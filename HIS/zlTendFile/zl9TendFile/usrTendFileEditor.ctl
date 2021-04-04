VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl usrTendFileEditor 
   AutoRedraw      =   -1  'True
   ClientHeight    =   5550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8565
   KeyPreview      =   -1  'True
   ScaleHeight     =   5550
   ScaleWidth      =   8565
   Begin VB.PictureBox picSignCheck 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2865
      Left            =   2580
      ScaleHeight     =   2865
      ScaleWidth      =   4815
      TabIndex        =   68
      Top             =   4380
      Visible         =   0   'False
      Width           =   4815
      Begin VB.CommandButton cmd取消 
         Caption         =   "取消"
         Height          =   350
         Left            =   3690
         TabIndex        =   71
         ToolTipText     =   "取消"
         Top             =   2370
         Width           =   840
      End
      Begin VB.CommandButton cmdSignCur 
         Caption         =   "验证"
         Height          =   350
         Left            =   2790
         TabIndex        =   70
         ToolTipText     =   "确认"
         Top             =   2370
         Width           =   840
      End
      Begin VB.CommandButton cmdSignAll 
         Caption         =   "全部"
         Height          =   350
         Left            =   270
         TabIndex        =   69
         ToolTipText     =   "确认"
         Top             =   2370
         Width           =   840
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfSignData 
         Height          =   1635
         Left            =   0
         TabIndex        =   72
         Top             =   630
         Width           =   4755
         _cx             =   8387
         _cy             =   2884
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
         BackColorSel    =   16764057
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   5000
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"usrTendFileEditor.ctx":0000
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
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   0   'False
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
      Begin VB.Image imgNote 
         Height          =   480
         Left            =   120
         Picture         =   "usrTendFileEditor.ctx":0062
         Top             =   90
         Width           =   480
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "以下是签名历史记录，可选择单行验证，也可进行全部验证。"
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   810
         TabIndex        =   73
         Top             =   150
         Width           =   3720
      End
   End
   Begin VB.PictureBox picPage 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1590
      ScaleHeight     =   285
      ScaleWidth      =   1095
      TabIndex        =   64
      Top             =   60
      Width           =   1095
      Begin VB.TextBox txtPage 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   420
         TabIndex        =   65
         Top             =   15
         Width           =   405
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "页"
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
         Left            =   870
         TabIndex        =   67
         Top             =   45
         Width           =   195
      End
      Begin VB.Label lblPage 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "定位"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   30
         TabIndex        =   66
         Top             =   45
         Width           =   360
      End
   End
   Begin MSComctlLib.ImageList imgRow 
      Left            =   6150
      Top             =   510
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
            Picture         =   "usrTendFileEditor.ctx":0CA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrTendFileEditor.ctx":103E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrTendFileEditor.ctx":13D8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtLength 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   1005
      Left            =   2340
      MultiLine       =   -1  'True
      TabIndex        =   17
      Top             =   3090
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4515
      Left            =   60
      ScaleHeight     =   4515
      ScaleWidth      =   8385
      TabIndex        =   12
      Top             =   510
      Width           =   8385
      Begin VB.PictureBox picInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   0
         ScaleHeight     =   330
         ScaleWidth      =   1200
         TabIndex        =   81
         Top             =   0
         Width           =   1200
         Begin VB.PictureBox picImg 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   840
            Picture         =   "usrTendFileEditor.ctx":172A
            ScaleHeight     =   225
            ScaleWidth      =   285
            TabIndex        =   82
            Top             =   45
            Width           =   285
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "温馨提示:"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   15
            TabIndex        =   83
            Top             =   45
            Width           =   810
         End
      End
      Begin VB.PictureBox picYear 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6825
         ScaleHeight     =   255
         ScaleWidth      =   765
         TabIndex        =   79
         Top             =   585
         Visible         =   0   'False
         Width           =   795
         Begin VB.ComboBox cboYear 
            BackColor       =   &H80000018&
            Height          =   300
            Left            =   -30
            Style           =   2  'Dropdown List
            TabIndex        =   80
            Top             =   -30
            Width           =   1080
         End
      End
      Begin VB.PictureBox PicLst 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H80000008&
         Height          =   2265
         Left            =   5430
         ScaleHeight     =   2235
         ScaleWidth      =   1185
         TabIndex        =   3
         Top             =   1590
         Visible         =   0   'False
         Width           =   1215
         Begin VB.TextBox txtLst 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   0
            MultiLine       =   -1  'True
            TabIndex        =   4
            Top             =   270
            Width           =   1215
         End
         Begin VB.ListBox lstSelect 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            Height          =   1470
            Index           =   0
            ItemData        =   "usrTendFileEditor.ctx":7F7C
            Left            =   -10
            List            =   "usrTendFileEditor.ctx":7F92
            TabIndex        =   5
            Top             =   825
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label lbllst 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000018&
            Caption         =   "选择："
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   1
            Left            =   15
            TabIndex        =   77
            Top             =   615
            Width           =   540
         End
         Begin VB.Label lbllst 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000018&
            Caption         =   "录入："
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   0
            Left            =   30
            TabIndex        =   76
            Top             =   30
            Width           =   540
         End
      End
      Begin VB.PictureBox picSign 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   2610
         ScaleHeight     =   195
         ScaleWidth      =   945
         TabIndex        =   74
         Tag             =   "225"
         Top             =   3615
         Visible         =   0   'False
         Width           =   975
         Begin VB.Image imgSign 
            Height          =   240
            Left            =   -30
            Picture         =   "usrTendFileEditor.ctx":7FCA
            Tag             =   "240"
            Top             =   -30
            Width           =   240
         End
         Begin VB.Label lbl验证签名 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "验证签名"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   210
            TabIndex        =   75
            Top             =   0
            Width           =   720
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfHead 
         Height          =   795
         Left            =   0
         TabIndex        =   63
         Top             =   930
         Width           =   4305
         _cx             =   7594
         _cy             =   1402
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
         BackColorSel    =   16764057
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   4
         FixedRows       =   3
         FixedCols       =   1
         RowHeightMin    =   255
         RowHeightMax    =   5000
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"usrTendFileEditor.ctx":E81C
         ScrollTrack     =   -1  'True
         ScrollBars      =   0
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
         AutoSizeMouse   =   0   'False
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
      Begin VB.PictureBox picDoubleChoose 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   6330
         ScaleHeight     =   240
         ScaleWidth      =   900
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   3300
         Visible         =   0   'False
         Width           =   930
         Begin VB.PictureBox picChooseRight 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   540
            ScaleHeight     =   255
            ScaleWidth      =   375
            TabIndex        =   59
            TabStop         =   0   'False
            Top             =   0
            Width           =   375
            Begin VB.ComboBox cboChoose 
               BackColor       =   &H80000018&
               Height          =   300
               Index           =   1
               ItemData        =   "usrTendFileEditor.ctx":E87E
               Left            =   -30
               List            =   "usrTendFileEditor.ctx":E88E
               Style           =   2  'Dropdown List
               TabIndex        =   60
               Top             =   -30
               Width           =   1605
            End
         End
         Begin VB.PictureBox picChooseLeft 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   435
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   0
            Width           =   435
            Begin VB.ComboBox cboChoose 
               BackColor       =   &H80000018&
               Height          =   300
               Index           =   0
               ItemData        =   "usrTendFileEditor.ctx":E8A0
               Left            =   -30
               List            =   "usrTendFileEditor.ctx":E8B0
               Style           =   2  'Dropdown List
               TabIndex        =   58
               Top             =   -30
               Width           =   1605
            End
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "/"
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
            Left            =   435
            TabIndex        =   61
            Top             =   30
            Width           =   105
         End
      End
      Begin VB.PictureBox picMutilInput 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   6135
         ScaleHeight     =   405
         ScaleWidth      =   1575
         TabIndex        =   10
         Top             =   3675
         Visible         =   0   'False
         Width           =   1600
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   225
            Index           =   0
            Left            =   810
            TabIndex        =   11
            Top             =   90
            Width           =   675
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "体温体录"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   0
            Left            =   15
            TabIndex        =   15
            Top             =   112
            Width           =   720
         End
      End
      Begin VB.CheckBox chkSwitch 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "全选"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   0
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   930
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.CommandButton cmdWord 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6120
         Picture         =   "usrTendFileEditor.ctx":E8C2
         Style           =   1  'Graphical
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1290
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.PictureBox picDouble 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   6330
         ScaleHeight     =   240
         ScaleWidth      =   900
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   2910
         Visible         =   0   'False
         Width           =   930
         Begin VB.PictureBox picDnInput 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   540
            ScaleHeight     =   255
            ScaleWidth      =   375
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   0
            Width           =   375
            Begin VB.Label lblDnInput 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   60
               TabIndex        =   22
               Top             =   30
               Width           =   315
            End
         End
         Begin VB.PictureBox picUpInput 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   435
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   0
            Width           =   435
            Begin VB.Label lblUpInput 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               ForeColor       =   &H80000008&
               Height          =   165
               Left            =   60
               TabIndex        =   21
               Top             =   30
               Width           =   315
            End
         End
         Begin VB.TextBox txtDnInput 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   525
            MaxLength       =   12
            TabIndex        =   9
            Top             =   30
            Width           =   345
         End
         Begin VB.TextBox txtUpInput 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   180
            Left            =   30
            MaxLength       =   12
            TabIndex        =   8
            Top             =   30
            Width           =   375
         End
         Begin VB.Label lblSplit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "/"
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
            Left            =   435
            TabIndex        =   18
            Top             =   30
            Width           =   105
         End
      End
      Begin VB.ListBox lstSelect 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   2340
         Index           =   1
         ItemData        =   "usrTendFileEditor.ctx":EC04
         Left            =   6675
         List            =   "usrTendFileEditor.ctx":EC1A
         Style           =   1  'Checkbox
         TabIndex        =   6
         Top             =   1590
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.PictureBox picInput 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000018&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5790
         ScaleHeight     =   225
         ScaleWidth      =   585
         TabIndex        =   1
         Top             =   1290
         Visible         =   0   'False
         Width           =   615
         Begin VB.TextBox txtInput 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   180
            Left            =   0
            MultiLine       =   -1  'True
            TabIndex        =   2
            Top             =   30
            Width           =   315
         End
         Begin VB.Label lblInput 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            Caption         =   "√"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   0
            TabIndex        =   16
            Top             =   30
            Width           =   315
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid VsfData 
         Height          =   2655
         Left            =   0
         TabIndex        =   0
         Top             =   930
         Width           =   4305
         _cx             =   7594
         _cy             =   4683
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
         BackColorSel    =   16764057
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   4
         FixedRows       =   3
         FixedCols       =   1
         RowHeightMin    =   255
         RowHeightMax    =   5000
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"usrTendFileEditor.ctx":EC52
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
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   0   'False
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
      Begin VSFlex8Ctl.VSFlexGrid vsTest 
         Height          =   495
         Left            =   1920
         TabIndex        =   62
         Top             =   930
         Visible         =   0   'False
         Width           =   1845
         _cx             =   3254
         _cy             =   873
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
         BackColorSel    =   16764057
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   4
         FixedRows       =   3
         FixedCols       =   1
         RowHeightMin    =   255
         RowHeightMax    =   5000
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"usrTendFileEditor.ctx":ECB4
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
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   0   'False
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
      Begin VB.Label lblCurPage 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "P333"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   7650
         TabIndex        =   55
         Top             =   240
         Width           =   555
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "一般护理记录单"
         Height          =   180
         Left            =   3450
         TabIndex        =   14
         Top             =   30
         Width           =   1275
      End
      Begin VB.Label lblSubhead 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名:##"
         Height          =   180
         Left            =   390
         TabIndex        =   13
         Top             =   540
         Width           =   720
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox picCloumn 
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   3075
      Left            =   525
      ScaleHeight     =   3075
      ScaleWidth      =   5955
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   1320
      Visible         =   0   'False
      Width           =   5955
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   2595
         MaxLength       =   20
         TabIndex        =   27
         Top             =   465
         Width           =   1200
      End
      Begin MSComctlLib.ListView lstColumnItems 
         Height          =   2490
         Left            =   45
         TabIndex        =   26
         Top             =   450
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   4392
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "项目序号"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "项目名称"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "部位"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.CommandButton cmdFilterOK 
         Height          =   315
         Left            =   2460
         Picture         =   "usrTendFileEditor.ctx":ED16
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "确认"
         Top             =   2310
         Width           =   450
      End
      Begin VB.CommandButton cmdFilterCancel 
         Height          =   315
         Left            =   3000
         Picture         =   "usrTendFileEditor.ctx":F2A0
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "取消"
         Top             =   2310
         Width           =   450
      End
      Begin VB.CommandButton cmdColumn 
         Caption         =   "选用(&S)"
         Height          =   300
         Index           =   0
         Left            =   2430
         TabIndex        =   28
         Top             =   1395
         Width           =   1100
      End
      Begin VB.CommandButton cmdColumn 
         Caption         =   "删除(&E)"
         Height          =   300
         Index           =   1
         Left            =   2430
         TabIndex        =   29
         Top             =   1725
         Width           =   1100
      End
      Begin VB.TextBox txtColumnNo 
         Height          =   300
         Left            =   4665
         MaxLength       =   20
         TabIndex        =   33
         Top             =   120
         Width           =   1185
      End
      Begin MSComctlLib.ListView lstColumnUsed 
         Height          =   2490
         Left            =   3840
         TabIndex        =   34
         Top             =   450
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   4392
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "项目序号"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "项目名称"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "部位"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label lblFind 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "查找"
         Height          =   180
         Left            =   2160
         TabIndex        =   78
         Top             =   525
         Width           =   360
      End
      Begin VB.Label lblNote 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "已发生数据，不允许调整设置。"
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   2385
         TabIndex        =   35
         Top             =   870
         Width           =   1260
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblColumnItems 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "可选护理记录项目"
         Height          =   180
         Left            =   105
         TabIndex        =   25
         Top             =   180
         Width           =   1440
      End
      Begin VB.Label lblColumnNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "表头名称"
         Height          =   180
         Left            =   3855
         TabIndex        =   32
         Top             =   180
         Width           =   720
      End
   End
   Begin VB.PictureBox picBiref 
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   1560
      ScaleHeight     =   3255
      ScaleWidth      =   5085
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   1755
      Visible         =   0   'False
      Width           =   5085
      Begin MSComCtl2.DTPicker DTPDate 
         Height          =   300
         Left            =   930
         TabIndex        =   39
         Top             =   180
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   114163715
         CurrentDate     =   40805
      End
      Begin VB.TextBox txt结束时点 
         Enabled         =   0   'False
         Height          =   300
         Left            =   3480
         MaxLength       =   5
         TabIndex        =   45
         Top             =   540
         Width           =   1365
      End
      Begin VB.TextBox txt开始时点 
         Enabled         =   0   'False
         Height          =   300
         Left            =   930
         MaxLength       =   5
         TabIndex        =   43
         Top             =   540
         Width           =   1365
      End
      Begin VB.ComboBox cbo标识 
         Height          =   300
         Left            =   930
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   900
         Width           =   3915
      End
      Begin VB.TextBox txt小结名称 
         Height          =   300
         Left            =   930
         TabIndex        =   49
         Top             =   1260
         Width           =   3885
      End
      Begin VB.ComboBox cbo小结 
         Height          =   300
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   180
         Width           =   1365
      End
      Begin VB.CommandButton cmdCancel 
         Height          =   315
         Left            =   4290
         Picture         =   "usrTendFileEditor.ctx":F82A
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "取消"
         Top             =   2850
         Width           =   450
      End
      Begin VB.CommandButton cmdOk 
         Height          =   315
         Left            =   3750
         Picture         =   "usrTendFileEditor.ctx":FDB4
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "确认"
         Top             =   2850
         Width           =   450
      End
      Begin VSFlex8Ctl.VSFlexGrid vfgItemList 
         Height          =   1100
         Left            =   930
         TabIndex        =   51
         Top             =   1635
         Width           =   3885
         _cx             =   6853
         _cy             =   1940
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
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483624
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   4
         FixedRows       =   3
         FixedCols       =   1
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
      Begin MSComctlLib.ImageList imgTrueFalse 
         Left            =   195
         Top             =   1965
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
               Picture         =   "usrTendFileEditor.ctx":1033E
               Key             =   "T"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "usrTendFileEditor.ctx":108D8
               Key             =   "F"
            EndProperty
         EndProperty
      End
      Begin VB.Label lblItemName 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "汇总项目"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   150
         TabIndex        =   50
         Top             =   1635
         Width           =   720
      End
      Begin VB.Label lbl结束时点 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "～ 结束时点"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2430
         TabIndex        =   44
         Top             =   600
         Width           =   990
      End
      Begin VB.Label lbl开始时点 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "开始时点"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   150
         TabIndex        =   42
         Top             =   600
         Width           =   720
      End
      Begin VB.Label lbl标识 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "标识"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   510
         TabIndex        =   46
         Top             =   960
         Width           =   360
      End
      Begin VB.Label lbl时间 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "时间"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   510
         TabIndex        =   38
         Top             =   240
         Width           =   360
      End
      Begin VB.Label lblCollectInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "保存数据后显示正确汇总"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   690
         TabIndex        =   54
         Top             =   2940
         Width           =   2010
         WordWrap        =   -1  'True
      End
      Begin VB.Label lbl小结名称 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "名称"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   510
         TabIndex        =   48
         Top             =   1320
         Width           =   360
      End
      Begin VB.Label lbl小结 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "小结"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3060
         TabIndex        =   40
         Top             =   240
         Width           =   360
      End
   End
   Begin VB.Line Line1 
      X1              =   3675
      X2              =   4875
      Y1              =   2535
      Y2              =   3015
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "usrTendFileEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'基础条件:
'1.护理记录同一时点只可能存在一条记录
'2.护理记录中不需要像体温单那样 , 记录病人是否外出, 拒测的数据, 测试了的数据才记录
'3.录入护理记录数据时,如果所录入的数据存在体温数据, 则提取过来
'4.护理记录单中不需要录入物理降温及脉搏短拙，如确需要可录入在护理摘要等文字型的列中
'#实现原理:
'1.对于用户修改过的数据,由于提供编辑状态页面切换的功能,对用户修改过的页数据进行整页复制,减少程序实现难度
'2.增加记录集记录哪些页哪些单元格被用户修改过
'3.任何编辑(粘贴,清空数据),都需要重新计算每行数据的占用行


'*******************************************************
'2012-01-06;zyb'活动项目选择中增加文本项目,一列只能绑定一个文本项目



Public mblnEditable As Boolean
'Public objFileSys As New FileSystemObject
'Public objStream As TextStream

Private mblnRestore As Boolean              '重新加载数据还是恢复页面数据
Private mfrmParent As Object
Private mblnInit As Boolean
Private mblnShow As Boolean                 '是否显示录入框
Private mblnVerify As Boolean               '是否审签模式(可修改,但不允许进行复制粘贴清除等操作,只能修改)
Private mstrVerify As String                '等待审签的ID串
Private mobjVerify As Collection        '等待审签的行信息(key存放记录ID,信息为:原发生时间)--81535:审签可修改时间调整
Private mintVerify As Integer               '当前操作员的最高级别
Private mintVerify_Last As Integer          '所选审签记录中最高级别
Private mblnBlowup As Boolean               '放大否？放大1/3，如字体9号放大为12号
Private mblnChange As Boolean               '是否修改数据
Private mstrData As String                  '进入编辑状态前保存之前的数据
Private mintNORule As Integer               '护理文件页码规则
Private mintPreDays As Long
Private mstrMaxDate As String
Private mlngSingerType As Long              '护士、签名人显示模式（是首行显示还是首尾显示等）

Private mint起始页码 As Integer
Private mint结束页 As Integer
Private mint页码 As Integer
Private mlng文件ID As Long
Private mlng格式ID As Long
Private mlng病人ID As Long
Private mlng主页ID As Long
Private mlng科室ID As Long
Private mlng病区ID As Long
Private mint婴儿 As Integer
Private mbln心率 As Boolean                 '是否需要录入心率
Private mstrPrivs As String

Private mintSymbol As Integer               '当前控件索引
Private mstrSymbol As String                 '特殊字符
Private mblnClear As Boolean                '如果为真,清除mrsDataMap记录集;当换页时应传假,保留用户修改的数据以备显示、保存使用
Private mstrCollectItems As String         '汇总项目集合
Private mstrColCollect As String             '汇总项目列集合:col;1|col;4,5
Private mstrColCorrelative As String       '汇总项目关联列集合:COl,3;COl,4|COl,5;COl,6(名称列号,项目序号;汇总列,项目序号),主要针对分类汇总
Private mstrColImCorrelative As String    '汇总项目关联列集合:COl,3;COl,4|COl,5;COl,6(名称列号,项目序号;汇总列,项目序号),主要针对入量导入
Private mblnCorrelative As Boolean        '是否启用了分类汇总
Private mstrCOLNothing As String          '未绑定的列集合+活动项目列(不管活动项目列是否绑定)
Private mstrCOLActive As String             '活动列集合
Private mstrCatercorner As String           '列对角线集合
Private mblnEditAssistant As Boolean        '当前选择的项目是否允许进行词句选择
Private mblnEditText As Boolean             '选择的项目是否是文本项目
Private mlngPageRows As Long                '此文件格式一页所显示的数据行
Private mArrPageInfo() As String            '每一页记录单的提示
Private mlngLitterRows() As Long            '记录跨页分组行的行数
Private mlngCurLitterRows() As Long         '记录跨页分组行在本页的实际行数
Private mlngOverrunRows As Long             '超出数据行
Private mlngReduceRow As Long               '减少数据行(合并文件起始页的起始行可能不是从1开始)
Private mlngRowCount As Long                '当前记录总行数
Private mlngRowCurrent As Long              '当前记录在本页的实际行数
Private mlngStartRowPage As Long            '当前记录的开始页号
Private mlngStartRowNo As Long              '当前记录单的开始行号
Private mlngDate As Long                    '日期
Private mlngTime As Long                    '时间
Private mlngChoose As Long                  '选择列
Private mlngYear As Long                    '年份:短日期格式时显示
Private mlngOperator As Long                '护士
Private mlngJoinSignName As Long            '交班签名人
Private mlngSignLevel As Long               '签名级别
Private mlngSigner As Long                  '签名信息
Private mlngSignName As Long                '签名人
Private mlngSignTime As Long                '签名时间
Private mlngRecord As Long                  '记录ID
Private mlngNoEditor As Long                '禁止编辑列,存在护士列则以护士列为准,不存在护士列则以签名列为准
Private mlngCollectType As Long             '汇总类别
Private mlngCollectText As Long             '汇总文本
Private mlngCollectStyle As Long            '汇总标记
Private mlngCollectDay As Long              '汇总日期:0-昨天;1-今天
Private mlngCollectStart As Long            '汇总开始时点
Private mlngCollectEnd As Long              '汇总结束时点
Private mlngDemo As Long                    '备用列
Private mlngActTime As Long                 '发生时间

Private mblnGroupNew As Boolean             '分组标志
Private mstrGroupRow As String              '分组起始行
Private mblnGroupApp As Boolean             '追加分组模式?
Private mblnSign As Boolean                 '是否签名
Private mblnArchive As Boolean              '是否归档
Private mintType As Integer                 '记录当前的编辑模式
Private mintCollectDef As Integer           '缺省小结格式
Private mlngCollectColor As Long            '小结标识颜色
Private mintPageSpan As Integer             '跨页显示；1-当前页；2-两页均显示
Private mintSignMode As Integer             '审签模式:0-聘任职务+审签权限;1-审签权限
Private mblnDateAd As Boolean               '日期缩写?
Private mstr开始时间 As String              '当前文件的开始时间
Private mstr结束时间 As String              '当前文件的结束时间
Private mstrYears As String                 '可选取的年份范围
Private CellRect As RECT
Private mbln护士 As Boolean                '是否绑定了护士列

Private rsTemp As New ADODB.Recordset
Private mrsItems As New ADODB.Recordset             '所有护理记录项目清单
Private mrsElement As New ADODB.Recordset           '使用于护理记录单的标签要素
Private mrsSelItems As New ADODB.Recordset          '当前录入的护理记录项目清单
Private mrsDataMap As New ADODB.Recordset           '当前操作员录入的数据镜像,与记录单格式一致,相关行数据全部保存以便迅速恢复
Private mrsCellMap As New ADODB.Recordset           '编辑过的数据镜像,字段有:页号,行号,列号,记录ID,数据,部位,删除
Private mrsCopyMap As New ADODB.Recordset           '复制行数据

Private mblnElement As Boolean                      '是否包含自定义标签要素
Private Enum ColIcon
    签名 = 1
    审签 = 2
    交班签名 = 3
End Enum
Private Enum SignLevel
    正高 = 1
    副高 = 2
    中级 = 3
    师级 = 4
    员士 = 5
    未定义 = 9
End Enum

Private Const conMenu_Save = 1

Private Const WS_MAXIMIZE = &H1000000
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_CAPTION = &HC00000
Private Const WS_SYSMENU = &H80000
Private Const WS_THICKFRAME = &H40000
Private Const WS_CHILD = &H40000000
Private Const WS_POPUP = &H80000000
Private Const SWP_NOZORDER = &H4
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOOWNERZORDER = &H200
Private Const SWP_NOREPOSITION = SWP_NOOWNERZORDER

Public Event AfterDataChanged(ByVal blnChange As Boolean)
Public Event AfterRefresh()
Public Event AfterRowColChange(ByVal strInfo As String, ByVal blnImportant As Boolean, ByVal blnSign As Boolean, ByVal blnArchive As Boolean)
Public Event ShowTipInfo(ByVal vsfObj As Object, ByVal strInfo As String, ByVal blnMultiRow As Boolean)
Dim strFields As String
Dim strValues As String
Dim blnScroll As Boolean

'记录上次选择行,顶行,以便刷新后重新定位
Dim lngLastRow As Long
Dim lngLastTopRow As Long
Dim lngLastPatientID As Long

Private mstrTag As String           '暂存

'病历文件格式定义相关
Private mintTabTiers As Integer     '表头层次
Private mintTagFormHour As Integer  '开始时间条件
Private mintTagToHour As Integer    '截止时间条件
Private mobjTagFont As New StdFont  '条件样式字体
Private mlngTagColor As Long        '条件样式颜色
Private mstrPaperSet As String      '格式
Private mstrPageHead As String      '页眉
Private mstrPageFoot As String      '页脚
Private mblnChildForm As Boolean
Private mstrSubhead As String       '表上标签
Private mstrTabHead As String       '表头单元
Private mstrColWidth As String      '列宽序列串
Private mstrColumns As String       '当前护理文件各列对应的项目
Private lngCurColor As Long, strCurFont As String, objFont As StdFont
'保存打开护理记录文件的SQL，在其它地方也有使用，不能修改
Private mstrSQL内 As String
Private mstrSQL中 As String
Private mstrSQL列 As String
Private mstrSQL条件 As String
Private mstrSQL As String

Private mcbrToolBar As CommandBar
Private mcbrPage As CommandBarControl
Const clngPage As Long = 3906
Const clngPageLocate As Long = 3907

'######################################################################################################################
'**********************************************************************************************************************
'以#分隔的区域内的代码都与绘图相关,没事别动
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long

Private Type POINTAPI
        X As Long
        Y As Long
End Type
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Const WHITE_BRUSH = 0    '白色画笔
Private Const cdblWidth As Double = 6          '一个英文字符的宽度
Private Const cHideCols = 4         '前缀隐藏列:备用,时间,选择,年度
Private Const cControlFields = 2    '记录集控制列:页号,行号

Private Function GetRBGFromOLEColor(ByVal dwOleColour As Long) As Long
    '将VB的颜色转换为RGB表示
    Dim clrref As Long
    Dim r As Long, g As Long, b As Long
    
    OleTranslateColor dwOleColour, 0, clrref
    
    b = (clrref \ 65536) And &HFF
    g = (clrref \ 256) And &HFF
    r = clrref And &HFF
    
    GetRBGFromOLEColor = RGB(r, g, b)
End Function

Private Function GetSymbolWidth(ByVal strPara As String) As Double
    '缺省是宋体9号,按字体大小同比放大
    Dim sinFontSize As Single
    Dim i As Integer, j As Integer
    
    j = Len(strPara)
    sinFontSize = VsfData.FontSize
    For i = 1 To j
        GetSymbolWidth = GetSymbolWidth + IIf(Asc(Mid(strPara, i, 1)) > 0, 1, 2) * cdblWidth * sinFontSize / 9
    Next
End Function

Private Sub DrawCell(ByVal hDC As Long, ByVal ROW As Long, ByVal COL As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    Dim strText As String
    Dim strLeft As String
    Dim strRight As String
    Dim lngLeft As Long
    Dim lngRight As Long
    Dim dblWidth As Double
    Dim lngBackColor As Long
    Dim lngForeColor As Long
    Dim blnDraw As Boolean
    '绘图相关
    Dim lngPen As Long
    Dim lngOldPen As Long
    Dim lngBrush As Long
    Dim lngOldBrush As Long
    Dim lpPoint As POINTAPI
    Dim t_ClientRect As RECT
    On Error GoTo ErrHand
    '******************************************
    '在此事件中不能对单元格的任何属性赋值,包括Celldata,否则会引起该事件的死循环,导致工具栏或计时器无法正常工作。
    '******************************************
    '使用匹配的背景色，前景色与字体进行文本输出。
    Done = True
    If Not mblnInit Then Exit Sub
    If VsfData.RowHidden(ROW) Then Exit Sub
    Done = False
    
    strText = FormatValue(VsfData.TextMatrix(ROW, COL))
    If IsDiagonal(COL) And InStr(1, strText, "/") <> 0 Then
        blnDraw = True
        '赋初值
        strLeft = Split(strText, "/")(0)
        strRight = Mid(strText, InStr(1, strText, "/") + 1)
        lngLeft = LenB(StrConv(strLeft, vbFromUnicode))
        lngRight = LenB(StrConv(strRight, vbFromUnicode))
        '取字符宽度
        dblWidth = GetSymbolWidth(strRight)
        '设定客户区域大小
        With t_ClientRect
            .Left = Left + 1
            .Top = Top + 1
            .Right = Right - 1
            .Bottom = Bottom - 1
        End With
        
        '1、清空内容
        '创建与背景色相同的刷子
        If ROW < VsfData.FixedRows Then
            lngBackColor = GetRBGFromOLEColor(VsfData.BackColorFixed)
            lngForeColor = GetRBGFromOLEColor(VsfData.ForeColorFixed)
        Else
            If ROW = VsfData.RowSel Then
                lngBackColor = GetRBGFromOLEColor(VsfData.BackColorSel)
                lngForeColor = RGB(0, 0, 0)
            Else
                lngBackColor = RGB(255, 255, 255)
                lngForeColor = GetRBGFromOLEColor(VsfData.Cell(flexcpForeColor, ROW, COL))
            End If

        End If
        lngBrush = CreateSolidBrush(lngBackColor)
        '使用该刷子填充背景色
        lngOldBrush = SelectObject(hDC, lngBrush)
        Call FillRect(hDC, t_ClientRect, lngBrush)
        '立即销毁临时使用的刷子并还原刷子
        Call SelectObject(hDC, lngOldBrush)
        Call DeleteObject(lngBrush)
        
        '2、准备画线
        '创建新画笔
        Call SetTextColor(hDC, lngForeColor)
        lngPen = CreatePen(0, 1, lngForeColor)
        lngOldPen = SelectObject(hDC, lngPen)
        '画线
        Call MoveToEx(hDC, Left, Bottom - 2, lpPoint)
        Call LineTo(hDC, Right, Top)
        '输出文本
        Call TextOut(hDC, Left, Top, strLeft, lngLeft)
        Call TextOut(hDC, IIf(Right - dblWidth >= Left, Right - dblWidth, Left), Bottom - 16, strRight, lngRight)
        
        '还原画笔并销毁
        Call SelectObject(hDC, lngOldPen)
        Call DeleteObject(lngPen)
        
        '已完成作图
        Done = True
    End If
    
    '3、如果是汇总行，则进行特殊处理
    
    If Val(VsfData.TextMatrix(ROW, mlngCollectType)) < 0 And Val(VsfData.TextMatrix(ROW, mlngCollectStyle)) > 0 _
        And (IIf(mblnDateAd = True, COL >= mlngYear, COL >= mlngDate) And COL < mlngNoEditor) Then
        Call DrawCollectCell(hDC, ROW, COL, Left, Top, Right, Bottom)
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub DrawCollectCell(ByVal hDC As Long, ByVal ROW As Long, ByVal COL As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
    Dim lngPen As Long, lngOldPen As Long
    Dim lpPoint As POINTAPI
    Dim lngRowCount As Long
    '创建新画笔
    lngPen = CreatePen(0, 1, mlngCollectColor)
    lngOldPen = SelectObject(hDC, lngPen)
    
    Select Case Val(VsfData.TextMatrix(ROW, mlngCollectStyle))
    Case 1  '上下划横线(起始行画上横线，结束行画下横线)
        '画线
        lngRowCount = Val(VsfData.TextMatrix(ROW, mlngRowCount))
        If lngRowCount > 1 Then
            If FormatValue(VsfData.TextMatrix(ROW, mlngRowCount)) = lngRowCount & "|1" Then
                Call MoveToEx(hDC, Left, Top + IIf(ROW = VsfData.FixedRows, 1, 0), lpPoint)
                Call LineTo(hDC, Right, Top + IIf(ROW = VsfData.FixedRows, 1, 0))
            ElseIf FormatValue(VsfData.TextMatrix(ROW, mlngRowCount)) = lngRowCount & "|" & lngRowCount Then
                Call MoveToEx(hDC, Left, Bottom - 2, lpPoint)
                Call LineTo(hDC, Right, Bottom - 2)
            End If
        Else
            Call MoveToEx(hDC, Left, Top + IIf(ROW = VsfData.FixedRows, 1, 0), lpPoint)
            Call LineTo(hDC, Right, Top + IIf(ROW = VsfData.FixedRows, 1, 0))
            Call MoveToEx(hDC, Left, Bottom - 2, lpPoint)
            Call LineTo(hDC, Right, Bottom - 2)
        End If
    Case 2  '汇总项下双横线
        If InStr(1, "|" & mstrColCollect & ";", "|" & COL - (cHideCols + VsfData.FixedCols - 1) & ";") <> 0 Then 'And Val(VsfData.TextMatrix(ROW, COL)) <> 0 Then
            '画线
            Call MoveToEx(hDC, Left, Bottom - 4, lpPoint)
            Call LineTo(hDC, Right, Bottom - 4)
            Call MoveToEx(hDC, Left, Bottom - 2, lpPoint)
            Call LineTo(hDC, Right, Bottom - 2)
        End If
    Case 3  '上横线
        '画线
        lngRowCount = Val(VsfData.TextMatrix(ROW, mlngRowCount))
        If FormatValue(VsfData.TextMatrix(ROW, mlngRowCount)) = lngRowCount & "|1" Then
            Call MoveToEx(hDC, Left, Top + IIf(ROW = VsfData.FixedRows, 1, 0), lpPoint)
            Call LineTo(hDC, Right, Top + IIf(ROW = VsfData.FixedRows, 1, 0))
        End If
    Case 4 '汇总项下单横线
        If InStr(1, "|" & mstrColCollect & ";", "|" & COL - (cHideCols + VsfData.FixedCols - 1) & ";") <> 0 Then 'And Val(VsfData.TextMatrix(ROW, COL)) <> 0 Then
            '画线
            Call MoveToEx(hDC, Left, Bottom - 2, lpPoint)
            Call LineTo(hDC, Right, Bottom - 2)
        End If
    End Select
    
    '还原画笔并销毁
    Call SelectObject(hDC, lngOldPen)
    Call DeleteObject(lngPen)
End Sub

'######################################################################################################################
'**********************************************************************************************************************
'以#分隔的区域内的代码都与分行相关,没事别动
Private Function GetData(ByVal strInput As String) As Variant
    Dim arrData
    Dim strData As String
    Dim strLine(256) As Byte
    Dim lngRow As Long, lngRows As Long, lngLen As Long
    
    GetData = ""
    lngRows = SendMessage(txtLength.hWnd, EM_GETLINECOUNT, 0&, 0&)
    For lngRow = 1 To lngRows
        Call ClearArray(strLine)
        lngLen = SendMessage(txtLength.hWnd, EM_GETLINE, lngRow - 1, strLine(0))
        Call ClearArray(strLine, lngLen)
        strData = StrConv(strLine, vbUnicode)
        strData = TruncZero(strData)
        GetData = GetData & IIf(GetData = "", "", "|ZYB.ZLSOFT|") & strData & IIf(lngRow < lngRows, vbCrLf, "")
    Next
    GetData = Split(GetData, "|ZYB.ZLSOFT|")
End Function

Private Sub ClearArray(strLine() As Byte, Optional ByVal lngPos As Long = 0)
    Dim intDo As Integer, intMax As Integer
    intMax = UBound(strLine)
    For intDo = lngPos To intMax
        strLine(intDo) = 0
        If lngPos > 0 Then Exit Sub     '不为零,表示仅设置字符串结束符
    Next
    strLine(1) = 1
End Sub

Private Function TrimStr(ByVal str As String) As String
'功能：去掉字符串中\0以后的字符，并且去掉两端的空格

    If InStr(str, Chr(0)) > 0 Then
        TrimStr = Trim(Left(str, InStr(str, Chr(0)) - 1))
    Else
        TrimStr = Trim(str)
    End If
End Function

Private Function TruncZero(ByVal strInput As String) As String
'功能：去掉字符串中\0以后的字符
    Dim lngPos As Long
    
    lngPos = InStr(strInput, Chr(0))
    If lngPos > 0 Then
        TruncZero = Mid(strInput, 1, lngPos - 1)
    Else
        TruncZero = strInput
    End If
End Function

'**********************************************************************************************************************
'######################################################################################################################


Private Sub BoundItems(ByVal intCol As Integer)
    Dim lstItem As ListItem
    Dim rsActive As New ADODB.Recordset
    On Error GoTo ErrHand
    '只提供数字型,选择项或汇总项及文本项的活动项目
    '绑定活动项目(绑定一个项目不控制,绑定两个项目时,项目类型必须=0且项目表示只能是数值,选择或汇总,且两个项目项目类型与项目表示方法必须一致)
    '51883,刘鹏飞,2012-08-02,提供单选和多选活动项目的绑定
    '100334,陈刘,2016-09-20,活动项目使用科室条件设置
    gstrSQL = "" & _
        " SELECT A.项目序号,A.部位,A.项目名称,B.列头名称,NVL(B.标志,0) AS 标志" & vbNewLine & _
        " FROM" & vbNewLine & _
        "     (SELECT A.项目序号,B.部位,B.部位||A.项目名称 AS 项目名称" & vbNewLine & _
        "     FROM 护理记录项目 A,体温部位 B" & vbNewLine & _
        "     WHERE A.项目序号 =B.项目序号(+) AND A.项目性质=2 And NVL(A.应用场合,0)<>1 And " & vbNewLine & _
        "     (A.适用科室 = 1 Or (A.适用科室 = 2 And Exists (Select 1 From 护理适用科室 C Where b.项目序号 = c.项目序号 And c.科室id = [4])))) A," & vbNewLine & _
        "     (SELECT A.列头名称,A.项目序号,A.部位||B.项目名称 AS 项目名称,1 AS 标志" & vbNewLine & _
        "     FROM 病人护理活动项目 A,护理记录项目 B" & vbNewLine & _
        "     WHERE A.项目序号=B.项目序号 AND A.文件ID=[1] AND A.页号=[2] AND A.列号=[3] ) B" & vbNewLine & _
        " WHERE A.项目序号=B.项目序号(+) AND A.项目名称=B.项目名称(+)" & vbNewLine & _
        " ORDER BY A.项目序号"
    Set rsActive = zlDatabase.OpenSQLRecord(gstrSQL, "提取未设置的活动项目", mlng文件ID, mint页码, intCol, mlng科室ID)
    If rsActive.RecordCount = 0 Then
        RaiseEvent AfterRowColChange("没有可供选择的活动项目，请在护理项目管理模块中进行设置！", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    
    '加入活动项目
    '65671:刘鹏飞,2013-09-22,绑定列之前清空活动项目标头名称
    txtFind.Text = ""
    txtColumnNo.Text = ""
    lstColumnItems.ListItems.Clear
    lstColumnUsed.ListItems.Clear
    With rsActive
        Do While Not .EOF
            If !标志 = 1 Then
                txtColumnNo.Text = NVL(!列头名称)
                Set lstItem = lstColumnUsed.ListItems.Add(, Now & "_" & !项目序号 & "_" & lstColumnUsed.ListItems.Count, !项目序号)
                lstItem.SubItems(1) = !项目名称
                lstItem.SubItems(2) = NVL(!部位)
            Else
                Set lstItem = lstColumnItems.ListItems.Add(, Now & "_" & !项目序号 & "_" & lstColumnItems.ListItems.Count + 100, !项目序号)
                lstItem.SubItems(1) = !项目名称
                lstItem.SubItems(2) = NVL(!部位)
            End If
            .MoveNext
        Loop
    End With
    
    '设置控件坐标（左边或右边超出屏幕大小则靠右或靠左显示，否则以列为中心显示）
    With picCloumn
        .Left = VsfData.Left + VsfData.CellLeft + VsfData.CellWidth / 2 - .Width / 2
        .Top = picMain.Top + VsfData.Top + VsfData.CellTop
        If .Height + .Top + picMain.Top > ScaleHeight Then
            .Top = ScaleHeight - picMain.Top - .Height
        End If
        If .Left + .Width > ScaleWidth Then
            .Left = ScaleWidth - .Width
        End If
        If .Left < VsfData.Left Then
            .Left = VsfData.Left
        End If
        .Visible = True
    End With
    
    lblNote.Visible = ISColHaveData
    cmdColumn(0).Enabled = Not lblNote.Visible
    cmdColumn(1).Enabled = Not lblNote.Visible
'    cmdFilterOK.Enabled = Not lblNote.Visible
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function GetPeriod() As String
    Dim rs As New ADODB.Recordset
    Dim strPeriod As String

    On Error GoTo ErrHand
    
    '53588:刘鹏飞,2013-4-25,修改数据的时间小于病人入院时间，床号，病区不能显示问题
    '如：病人入科时间为2013-03-13 11:23:34 文件开始时间和入科相同，此时录入数据时间为 2013-03-13 11:23
    '就会导致无法提取床号，应为保存的数据时间为2013-03-13 11:23:00 小于了病人入科时间导致无法提取到数据
    '获取病人的入院时间
    If mint婴儿 = 0 Then
        gstrSQL = "Select 开始时间, Sysdate As 结束时间" & vbNewLine & _
            " From 病人变动记录" & vbNewLine & _
            " Where 病人id = [1] And 主页id = [2] And 开始原因 = 2" & vbNewLine & _
            " Union All" & vbNewLine & _
            " Select 开始时间, Sysdate As 结束时间" & vbNewLine & _
            " From 病人变动记录 a" & vbNewLine & _
            " Where a.病人id = [1] And a.主页id = [2] And a.开始原因 = 1 And Not Exists" & vbNewLine & _
            " (Select 1 From 病人变动记录 Where 病人id = a.病人id And 主页id = a.主页id And 开始原因 = 2)"

    Else
        gstrSQL = " Select   出生时间 AS 开始时间,sysdate AS 结束时间 From 病人新生儿记录 Where 病人ID=[1] And 主页ID=[2] And 序号=[3]"
    End If
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "取入院日期或出生日期", mlng病人ID, mlng主页ID, mint婴儿)
    
    '获取指定页码的数据发生时间范围
    gstrSQL = " Select  MIN(发生时间) 开始时间,MAX(发生时间) AS 结束时间 From 病人护理打印 Where 文件ID=[1] And (开始页号=[2] OR 结束页号=[2])"
    Call SQLDIY(gstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取指定页码的数据发生时间范围", mlng文件ID, IIf(mint页码 < mint结束页 + 1, mint页码, mint结束页))
    If NVL(rsTemp!开始时间) = "" Then
        strPeriod = Format(rs!开始时间, "yyyy-MM-dd HH:mm:ss") & "～" & Format(rs!结束时间, "yyyy-MM-dd HH:mm:ss")
    Else
        If Format(rsTemp!开始时间, "yyyy-MM-dd HH:mm:ss") < Format(rs!开始时间, "yyyy-MM-dd HH:mm:ss") Then
            strPeriod = Format(rs!开始时间, "yyyy-MM-dd HH:mm:ss") & "～" & Format(rsTemp!结束时间, "yyyy-MM-dd HH:mm") & ":59"
        Else
            strPeriod = Format(rsTemp!开始时间, "yyyy-MM-dd HH:mm:ss") & "～" & Format(rsTemp!结束时间, "yyyy-MM-dd HH:mm") & ":59"
        End If
    End If
    GetPeriod = strPeriod
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ReadStruDef() As Boolean
    Dim lngCol As Long
    On Error GoTo ErrHand
    
    '读取文件属性
    mblnDateAd = False
    mbln护士 = False
    
    Call GetFileProperty
    
    '提取活动项目并加入列定义(格式：列号;表头名称|项目序号,部位;项目序号,部位||列号;表头名称...)
    mstrCOLActive = ""
    mstrCOLNothing = ""
    mstrCollectItems = ""
    mstrColCollect = ""
    mstrColCorrelative = ""
    mstrColImCorrelative = ""
    mblnCorrelative = True '初始化必须为真(兼容处理)
    gstrSQL = " Select   A.列号,A.列头名称,A.序号,A.项目序号,A.部位 From 病人护理活动项目 A " & _
              " Where A.文件ID=[1] And A.页号=[2] " & _
              " Order by A.列号,A.序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取出所有自定义的活动项目", mlng文件ID, mint页码)
    If rsTemp.RecordCount <> 0 Then
        Do While Not rsTemp.EOF
            If lngCol <> rsTemp!列号 Then
                lngCol = rsTemp!列号
                mstrCOLActive = mstrCOLActive & "||" & rsTemp!列号 & ";" & rsTemp!列头名称 & "|" & rsTemp!项目序号 & "," & NVL(rsTemp!部位)
            Else
                mstrCOLActive = mstrCOLActive & ";" & rsTemp!项目序号 & "," & NVL(rsTemp!部位)
            End If
            rsTemp.MoveNext
        Loop
    End If
    If mstrCOLActive <> "" Then mstrCOLActive = Mid(mstrCOLActive, 3)
    
    '读取病历文件格式定义
    gstrSQL = "Select   d.对象序号, d.内容文本, d.要素名称" & _
        " From 病历文件结构 d, 病历文件结构 p" & _
        " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表格样式'" & _
        " Order By d.对象序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取病历文件格式定义", mlng格式ID)
    With rsTemp
        Do While Not .EOF
            Select Case "" & !要素名称
            Case "表头层数": mintTabTiers = Val("" & !内容文本)
            Case "总列数"
                VsfData.Cols = Val("" & !内容文本)
                vsfHead.Cols = VsfData.Cols
            Case "最小行高"
                VsfData.RowHeightMin = BlowUp(Val("" & !内容文本))
                vsfHead.RowHeightMin = VsfData.RowHeightMin
            Case "文本字体"
                strCurFont = "" & !内容文本
                Set objFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = BlowUp(Val(Split(strCurFont, ",")(1)))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "粗") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "斜") > 0 Then .Italic = True
                End With
                Set VsfData.Font = objFont
                Set vsfHead.Font = objFont
                Set lblSubhead.Font = VsfData.Font
                Set Font = lblSubhead.Font
                Set picMain.Font = Font
                
            Case "文本颜色"
                VsfData.ForeColor = Val("" & !内容文本)
                vsfHead.ForeColor = VsfData.ForeColor
            Case "表格颜色"
                VsfData.GridColor = Val("" & !内容文本): VsfData.GridColorFixed = VsfData.GridColor
                vsfHead.GridColor = VsfData.GridColor: vsfHead.GridColorFixed = VsfData.GridColor
            Case "标题文本"
                lblTitle.Caption = "" & !内容文本
                lblTitle.AutoSize = True
            Case "标题字体"
                strCurFont = "" & !内容文本
                Set objFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = BlowUp(Val(Split(strCurFont, ",")(1)))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "粗") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "斜") > 0 Then .Italic = True
                End With
                Set lblTitle.Font = objFont
                lblTitle.AutoSize = False
            
            Case "开始时间": mintTagFormHour = Val("" & !内容文本)
            Case "终止时间": mintTagToHour = Val("" & !内容文本)
            Case "条件字体"
                strCurFont = "" & !内容文本
                Set objFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = BlowUp(Val(Split(strCurFont, ",")(1)))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "粗") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "斜") > 0 Then .Italic = True
                End With
                Set mobjTagFont = objFont
            Case "条件颜色": mlngTagColor = Val("" & !内容文本)
            Case "有效数据行"
                mlngOverrunRows = 0
                mlngReduceRow = 0
                mlngPageRows = Val("" & !内容文本)
            Case "分类汇总"
                mblnCorrelative = (Val("" & !内容文本) = 1)
            End Select
            .MoveNext
        Loop
    End With
    
    gstrSQL = "Select   格式, 页眉, 页脚,报表 From 病历页面格式 Where 种类 = 3 And 编号 In (Select 页面 From 病历文件列表 Where Id = [1])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取病历页面格式", mlng格式ID)
    If Not rsTemp.EOF Then
        mstrPaperSet = "" & rsTemp!格式: mstrPageHead = "" & rsTemp!页眉: mstrPageFoot = "" & rsTemp!页脚
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select   d.对象序号, d.内容文本, d.要素名称, Nvl(d.是否换行, 0) As 是否换行" & _
        " From 病历文件结构 d, 病历文件结构 p" & _
        " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表上标签'" & _
        " Order By d.对象序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取表上标签定义", mlng格式ID)
    With rsTemp
        mstrSubhead = ""
        Do While Not .EOF
            mstrSubhead = mstrSubhead & "|" & IIf(!是否换行 = 0, "", vbCrLf) & !内容文本 & "{" & !要素名称 & "}"
            .MoveNext
        Loop
        If mstrSubhead <> "" Then mstrSubhead = Mid(mstrSubhead, 2)
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select   d.对象序号, d.内容行次, d.内容文本" & _
        " From 病历文件结构 d, 病历文件结构 p" & _
        " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表头单元'" & _
        " Order By d.对象序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取表头单元定义", mlng格式ID)
    With rsTemp
        mstrTabHead = ""
        Do While Not .EOF
            mstrTabHead = mstrTabHead & "|" & !内容行次 - 1 & "," & !对象序号 & "," & !内容文本
            .MoveNext
        Loop
        If mstrTabHead <> "" Then mstrTabHead = Mid(mstrTabHead, 2)
    End With
    
    '查询语句组织
    '------------------------------------------------------------------------------------------------------------------
    Dim strSql外 As String, str格式 As String, strSqlNull As String
    Dim bln日期 As Boolean, bln时间 As Boolean, bln护士 As Boolean
    Dim bln签名人 As Boolean, bln签名时间 As Boolean, bln签名日期 As Boolean
    Dim bln对角线 As Boolean, bln选择项 As Boolean          '如果上一列是对角线且选择项,则直接提取各项数据,拼列头时在数值间加上/
    Dim lngColumn As Long, blnAddCollect As Boolean
    Dim strColCorrelative  As String
    
    gstrSQL = "Select   d.对象序号,d.对象标记, d.对象属性, d.内容行次, d.内容文本, upper(d.要素名称) AS 要素名称, d.要素单位,d.要素表示 " & _
        " From 病历文件结构 d, 病历文件结构 p" & _
        " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表列集合'" & _
        " Order By d.对象序号, d.内容行次"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取表列集合定义", mlng格式ID)
    With rsTemp
        lngColumn = 0: mstrColumns = "": mstrColWidth = "": mstrCatercorner = "": strColCorrelative = ""
        mstrSQL内 = "": mstrSQL中 = "": strSql外 = "": mstrSQL列 = "": mstrSQL条件 = "": strSqlNull = ""
        bln日期 = False: bln时间 = False: bln护士 = False
        bln签名人 = False: bln签名时间 = False: bln签名日期 = False
        Do While Not .EOF
            If lngColumn <> !对象序号 Then
                blnAddCollect = False
                If strColCorrelative <> "" Then
                    mstrColCorrelative = mstrColCorrelative & "|" & strColCorrelative
                End If
                strColCorrelative = ""
                mstrColumns = mstrColumns & IIf(mstrColumns = "", "", "'1'" & str格式) & "|" & !对象序号 & "'" & !要素名称
                mstrColWidth = mstrColWidth & "," & !对象属性 & "`" & !对象序号 & "`" & !要素表示
                If !要素表示 = 1 Then mstrCatercorner = mstrCatercorner & "," & !对象序号
                str格式 = ""
                If !要素名称 <> "" Then str格式 = "{" & NVL(!内容文本) & "[" & !要素名称 & "]" & NVL(!要素单位) & "}"
                
                If Mid(strSqlNull, 3) = "" Then
                    strSqlNull = "''"
                Else
                    strSqlNull = Mid(strSqlNull, 3)
                End If
                mstrSQL列 = mstrSQL列 & "," & IIf(Mid(strSql外, 3) = "", "''", "Decode(" & Mid(strSql外, 3) & "," & strSqlNull & ",''," & Mid(strSql外, 3) & ")") & " As C" & Format(lngColumn, "00")
                
                strSql外 = ""
                strSqlNull = ""
                lngColumn = !对象序号
                bln对角线 = (NVL(!要素表示, 0) = 1)
                bln选择项 = False
                mrsItems.Filter = "项目名称='" & NVL(!要素名称) & "'"
                If mrsItems.RecordCount <> 0 Then
                    bln选择项 = (mrsItems!项目表示 = 5)
                    If mrsItems!项目表示 = 4 Then   '汇总项目
                        blnAddCollect = True
                        mstrCollectItems = mstrCollectItems & "," & mrsItems!项目序号
                        mstrColCollect = mstrColCollect & "|" & !对象序号 & ";" & mrsItems!项目序号
                        If Val(NVL(!对象标记)) > 0 And Val(NVL(!对象序号)) <> Val(NVL(!对象标记)) Then
                            strColCorrelative = Val(NVL(!对象标记)) & ";" & !对象序号 & "," & mrsItems!项目序号
                        End If
                    End If
                End If
                mrsItems.Filter = 0
            Else
                mstrColumns = mstrColumns & "," & !要素名称
                str格式 = str格式 & "{" & NVL(!内容文本) & "[" & !要素名称 & "]" & NVL(!要素单位) & "}"
                mrsItems.Filter = "项目名称='" & NVL(!要素名称) & "'"
                If mrsItems.RecordCount <> 0 Then
                    If mrsItems!项目表示 = 4 Then   '汇总项目
                        mstrCollectItems = mstrCollectItems & "," & mrsItems!项目序号
                        If blnAddCollect Then
                            strColCorrelative = ""
                            mstrColCollect = mstrColCollect & "," & mrsItems!项目序号
                        Else    '有可能一列绑定两个项目,第一个项目不是汇总项目,第二个项目才是汇总项目,因此,下面的代码保证加上列序号
                            blnAddCollect = True
                            mstrColCollect = mstrColCollect & "|" & !对象序号 & ";" & mrsItems!项目序号
                            If Val(NVL(!对象标记)) > 0 And Val(NVL(!对象序号)) <> Val(NVL(!对象标记)) Then
                                strColCorrelative = Val(NVL(!对象标记)) & ";" & !对象序号 & "," & mrsItems!项目序号
                            End If
                        End If
                    End If
                End If
                mrsItems.Filter = 0
            End If
            
            Select Case !要素名称
            Case "日期"
                bln日期 = True
                mblnDateAd = (NVL(!要素表示, 0) = 1)
                mstrSQL中 = mstrSQL中 & ",日期"
                mstrSQL内 = mstrSQL内 & ",To_Char(l.发生时间, " & IIf(mblnDateAd, "'dd/MM'", "'yyyy-mm-dd'") & ") As 日期"
                strSql外 = strSql外 & "||" & !要素名称
            Case "时间"
                bln时间 = True
                mstrSQL中 = mstrSQL中 & ",时间"
                mstrSQL内 = mstrSQL内 & ",To_Char(l.发生时间, 'hh24:mi') As 时间"
                strSql外 = strSql外 & "||" & !要素名称
                
            Case "签名人"
                bln签名人 = True
                mstrSQL中 = mstrSQL中 & ",签名人"
                '51589:刘鹏飞,2013-03-01,添加交班签名
                'mstrSQL内 = mstrSQL内 & ",l.签名人"
                mstrSQL内 = mstrSQL内 & ",DECODE(TRIM(NVL(L.签名人,'')),'',TRIM(L.签名人),DECODE(TRIM(NVL(L.交班签名人,'')),'',TRIM(L.签名人), TRIM(L.签名人) || '/' || TRIM(L.交班签名人))) 签名人"
                strSql外 = strSql外 & "||" & !要素名称
                
            Case "签名时间"
                bln签名时间 = True
                mstrSQL中 = mstrSQL中 & ",签名时间"
                mstrSQL内 = mstrSQL内 & ",l.签名时间"
                strSql外 = strSql外 & "||" & !要素名称
                
            Case "护士"
                bln护士 = True
                mstrSQL中 = mstrSQL中 & ",护士"
                mstrSQL内 = mstrSQL内 & ",l.保存人 As 护士"
                strSql外 = strSql外 & "||" & !要素名称
            Case Else
                If !要素名称 <> "" Then
                    mstrSQL中 = mstrSQL中 & ",Max(""" & !要素名称 & """) As """ & !要素名称 & """"
                    mstrSQL条件 = mstrSQL条件 & " Or """ & !要素名称 & """ Is Not Null"
                    
                    strSql外 = strSql外 & "||'" & !内容文本 & "'||""" & !要素名称 & """||'" & !要素单位 & "'"
                    strSqlNull = strSqlNull & "||" & "'" & !内容文本 & "'||'" & !要素单位 & "'"
                    mstrSQL内 = mstrSQL内 & ", Decode(c.项目名称, '" & !要素名称 & "', c.记录内容, '') As """ & !要素名称 & """"
                    
'                    If bln对角线 And bln选择项 Then
'                        If strSql外 <> "" Then
'                            '第二项
'                            strSql外 = strSql外 & "||'/'||""" & !要素名称 & """"
'                        Else
'                            '第一项
'                            strSql外 = strSql外 & "||""" & !要素名称 & """"
'                        End If
'                    Else
'                        strSql外 = strSql外 & "||""" & !要素名称 & """"
'                        strSqlNull = strSqlNull & "||" & "'" & !内容文本 & "'||'" & !要素单位 & "'"
'                    End If
'
'                    If (Trim("" & !内容文本) = "" And Trim("" & !要素单位) = "") Or (bln对角线 And bln选择项) Then
'                        mstrSQL内 = mstrSQL内 & ", Decode(c.项目名称, '" & !要素名称 & "', Nvl(c.记录内容), '') As """ & !要素名称 & """"
'                    Else
'                        mstrSQL内 = mstrSQL内 & ", Decode(c.项目名称, '" & !要素名称 & "', Decode(c.记录内容,Null,'" & !内容文本 & "'||'" & !要素单位 & "','" & !内容文本 & "'||c.记录内容||'" & !要素单位 & "'),  '" & !内容文本 & "'||'" & !要素单位 & "') As """ & !要素名称 & """"
'                    End If
                Else
                    '为空表示未绑定列,强制加,后面进行替换
                    mstrCOLNothing = mstrCOLNothing & "," & Val(Format(!对象序号, "00"))
                    mstrSQL中 = mstrSQL中 & ",Max(""" & "C" & Format(!对象序号, "00") & """) As C" & Format(!对象序号, "00")
                    mstrSQL条件 = mstrSQL条件 & " Or """ & "C" & Format(!对象序号, "00") & """ Is Not Null"
                    mstrSQL内 = mstrSQL内 & ", C" & Format(!对象序号, "00") & " AS C" & Format(!对象序号, "00")
                End If
            End Select
            .MoveNext
        Loop
        
        mbln护士 = bln护士
        
        If mstrCollectItems <> "" Then
            mstrCollectItems = Mid(mstrCollectItems, 2)
            mstrColCollect = Mid(mstrColCollect, 2)
        End If
        '在InitRecords中需要给汇总项目关列的名称列明添加项目序号
        If Left(mstrColCorrelative, 1) = "|" Then mstrColCorrelative = Mid(mstrColCorrelative, 2)
        mstrCOLNothing = Mid(mstrCOLNothing, 2)
        mstrCatercorner = Mid(mstrCatercorner, 2)
        mstrColWidth = Mid(mstrColWidth, 2)
        '加入最后一列的格式
        mstrColumns = mstrColumns & IIf(mstrColumns = "", "", "'1'" & str格式) '& "|" & !对象序号 & "'" & !要素名称
        mstrColumns = Mid(mstrColumns, 2)     '格式如:列号;项目名称1,项目名称2|列号...,实例;1;体温|2;脉搏|3...
        
        If Mid(strSqlNull, 3) = "" Then
            strSqlNull = "''"
        Else
            strSqlNull = Mid(strSqlNull, 3)
        End If
        mstrSQL列 = mstrSQL列 & "," & IIf(Mid(strSql外, 3) = "", "''", "Decode(" & Mid(strSql外, 3) & "," & strSqlNull & ",''," & Mid(strSql外, 3) & ")") & " As C" & Format(lngColumn, "00")
        
        If mstrSQL条件 <> "" Then mstrSQL条件 = "(" & Mid(mstrSQL条件, 5) & ")"
        
        '如果没有出现日期，时间，护士，则内层需要补充，以保证中层分组的正常：
        If bln日期 = False Then mstrSQL内 = mstrSQL内 & ",To_Char(l.发生时间, 'yyyy-mm-dd') As 日期"
        If bln时间 = False Then mstrSQL内 = mstrSQL内 & ",To_Char(l.发生时间, 'hh24:mi') As 时间"
        
        '51589:刘鹏飞,2013-03-01,添加交班签名
        'If bln签名人 = False Then mstrSQL内 = mstrSQL内 & ",l.签名人 As 签名人"
        If bln签名人 = False Then mstrSQL内 = mstrSQL内 & ",DECODE(TRIM(NVL(L.签名人,'')),'',TRIM(L.签名人),DECODE(TRIM(NVL(L.交班签名人,'')),'',TRIM(L.签名人), TRIM(L.签名人) || '/' || TRIM(L.交班签名人))) 签名人"
        If bln签名时间 = False Then mstrSQL内 = mstrSQL内 & ",l.签名时间"
        
        If Mid(mstrSQL中, 2) = "" Then
            MsgBox "对不起，您没有定义当前护理单的显示列信息，请在病历文件管理中定义！", vbInformation, gstrSysName
            Exit Function
        End If
        '51589:刘鹏飞,2013-03-01,添加交班签名
        '程序内部控制增加固定列
        mstrSQL中 = UCase(mstrSQL中 & ",MAX(签名级别) AS 签名级别,MAX(签名信息) AS 签名信息,MAX(交班签名人) AS 交班签名人,MAX(记录ID) AS 记录ID,MAX(行数) AS 行数,MAX(实际行数) AS 实际行数,Max(开始页号) AS 开始页号,Max(开始行号) AS 开始行号,MAX(汇总类别) AS 汇总类别,MAX(汇总文本) AS 汇总文本,MAX(汇总标记) AS 汇总标记,MAX(汇总日期) AS 汇总日期,MAX(开始时点) AS 开始时点,MAX(结束时点) AS 结束时点")
        mstrSQL内 = UCase(mstrSQL内 & ",l.签名级别,l.签名人 AS 签名信息,l.交班签名人,C.记录ID,P.行数||'' AS 行数,DECODE(SIGN(P.结束页号-P.开始页号),1,DECODE(SIGN([5]-P.开始页号),1, P.结束行号,P.行数-P.结束行号 ),P.行数) AS 实际行数,P.开始页号,P.开始行号,NVL(L.汇总类别,0) AS 汇总类别,L.汇总文本,L.汇总标记,to_char(L.发生时间,'yyyy-MM-dd hh24:mi:ss')||'' AS 汇总日期,L.开始时点,L.结束时点")
        mstrSQL列 = UCase(mstrSQL列 & ",签名级别,签名信息,交班签名人,记录ID,行数,实际行数,开始页号,开始行号,汇总类别,汇总文本,汇总标记,汇总日期,开始时点,结束时点")
        
        '63706:刘鹏飞,2013-11-20,强制绑定护士列
'        If bln护士 = False Then
        '强制添加护士列,为了避免修改他人数据行(他人录入的数据,增加新数据也不允许)
        mstrSQL中 = mstrSQL中 & ",护士L"
        mstrSQL内 = mstrSQL内 & ",l.保存人 As 护士L"
        mstrSQL列 = mstrSQL列 & ",护士L"
'        End If
        
        '将活动项目加入到SQL中
        Call DelActiveNoUsed
        Call PreActiveCOL
        Call SQLCombination
    End With
    
    ReadStruDef = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub PreActiveHead(ByVal objVsf As VSFlexGrid)
    Dim arrData
    Dim intCol As Integer
    Dim strName As String
    Dim intDo As Integer, intCount As Integer
    '更新表头
    
    arrData = Split(mstrCOLActive, "||")
    intCount = UBound(arrData)
    For intDo = 0 To intCount
        intCol = Split(Split(arrData(intDo), "|")(0), ";")(0)
        strName = Split(Split(arrData(intDo), "|")(0), ";")(1)
        objVsf.TextMatrix(mintTabTiers - 1, intCol + cHideCols + objVsf.FixedCols - 1) = strName
        If mintTabTiers = 3 And objVsf.TextMatrix(1, intCol + cHideCols + objVsf.FixedCols - 1) = "" Then objVsf.TextMatrix(1, intCol + cHideCols + objVsf.FixedCols - 1) = strName
        If mintTabTiers = 2 And objVsf.TextMatrix(0, intCol + cHideCols + objVsf.FixedCols - 1) = "" Then objVsf.TextMatrix(0, intCol + cHideCols + objVsf.FixedCols - 1) = strName
    Next
End Sub

Private Sub PreActiveCOL()
    Dim arrData
    Dim arrCol
    Dim intCol As Integer, strName As String
    Dim strColFormat As String, strCOLNames As String, strCOLPart As String, strCOLCOND As String, strCOLDEF As String, strCOLMID As String, strCOLIN As String
    Dim intDo As Integer, intCount As Integer
    Dim intIn As Integer, intMax As Integer
    '将活动项目加入到查询SQL中，格式：列号;表头名称|项目序号,部位;项目序号,部位||列号;表头名称...
    '绑定多个项目，该列就自动转为对角线列
    
    arrData = Split(mstrCOLActive, "||")
    intCount = UBound(arrData)
    For intDo = 0 To intCount
        intCol = Split(Split(arrData(intDo), "|")(0), ";")(0)
        strName = Split(Split(arrData(intDo), "|")(0), ";")(1)
        arrCol = Split(Split(arrData(intDo), "|")(1), ";")
        
        '处理列表示(每列最多绑定两个项目)
        strCOLPart = ""
        strCOLNames = ""
        strColFormat = ""
        strCOLCOND = ""
        strCOLMID = ""
        strCOLIN = ""
        strCOLDEF = ""
        intMax = UBound(arrCol)
        For intIn = 0 To intMax
            strCOLPart = Split(arrCol(intIn), ",")(1)
            mrsItems.Filter = "项目序号=" & Val(Split(arrCol(intIn), ",")(0))
            strCOLNames = strCOLNames & "," & mrsItems!项目名称
            strCOLCOND = strCOLCOND & " OR """ & strCOLPart & mrsItems!项目名称 & """ IS NOT NULL"
            strCOLMID = strCOLMID & ",Max(""" & strCOLPart & mrsItems!项目名称 & """) As """ & strCOLPart & mrsItems!项目名称 & """"
            If intIn = 0 Then
                strCOLIN = strCOLIN & ", Decode(" & IIf(strCOLPart = "", "", "c.体温部位||") & "c.项目名称, '" & strCOLPart & mrsItems!项目名称 & "',c.记录内容, '') As """ & strCOLPart & mrsItems!项目名称 & """"
            Else
                strCOLIN = strCOLIN & ", Decode(" & IIf(strCOLPart = "", "", "c.体温部位||") & "c.项目名称, '" & strCOLPart & mrsItems!项目名称 & "', Decode(c.记录内容,Null,'/','/'||c.记录内容||''), '') As """ & strCOLPart & mrsItems!项目名称 & """"
            End If
            If intIn = 0 Then
                If intMax = 0 Then
                    strCOLDEF = strCOLDEF & " """ & strCOLPart & mrsItems!项目名称 & """ AS C" & Format(intCol, "00")
                Else
                    strCOLDEF = strCOLDEF & " """ & strCOLPart & mrsItems!项目名称 & """||"
                End If
            Else
                strCOLDEF = strCOLDEF & "NVL(""" & strCOLPart & mrsItems!项目名称 & """,'/')"
                If intIn = intMax Then
                    strCOLDEF = "Decode(" & strCOLDEF & ",'" & String(intMax, "/") & "',''," & strCOLDEF & ") As C" & Format(intCol, "00")
                End If
            End If
            
            strColFormat = strColFormat & "{[" & strCOLPart & mrsItems!项目名称 & "]" & IIf(intMax > 0 And intIn < intMax, "/", "") & "}"
        Next
        If strCOLPart <> "" Then
            strCOLPart = Mid(strCOLPart, 2)
        End If
        strCOLNames = Mid(strCOLNames, 2)
        
        '对角线
        If intMax > 0 Then
            mstrCatercorner = mstrCatercorner & IIf(mstrCatercorner = "", "", ",") & intCol
        End If
        '列格式:15'护士'1'{[护士]}
        '77476:LPF:活动列替换intcol前添加"|"字符,避免第3列和第13列都为活动项目时项目替换错误
        mstrColumns = Replace(mstrColumns, "|" & intCol & "''1'", "|" & intCol & "'" & strCOLNames & "'1'" & strColFormat)
        '列
        mstrSQL列 = Replace(mstrSQL列, "'' AS C" & Format(intCol, "00"), strCOLDEF)
        '条件
        '53893:刘鹏飞,2012-09-21,处理活动项目绑定在时间后面的情况
        'mstrSQL条件 = Replace(UCase(mstrSQL条件), " OR """ & "C" & Format(intCol, "00") & """ IS NOT NULL", strCOLCOND)
        mstrSQL条件 = Replace(UCase(Replace(UCase(mstrSQL条件), " OR """ & "C" & Format(intCol, "00") & """ IS NOT NULL", strCOLCOND)), """" & "C" & Format(intCol, "00") & """ IS NOT NULL", Mid(strCOLCOND, 5))
        '中
        mstrSQL中 = Replace(mstrSQL中, ",MAX(""" & "C" & Format(intCol, "00") & """) AS C" & Format(intCol, "00"), strCOLMID)
        '内
        mstrSQL内 = Replace(mstrSQL内, ", C" & Format(intCol, "00") & " AS C" & Format(intCol, "00"), strCOLIN)
    Next
    mrsItems.Filter = 0
    
    '将未绑定的列的SQL部分清除
    If mstrCOLNothing = "" Then Exit Sub
    arrData = Split(mstrCOLNothing, ",")
    intCount = UBound(arrData)
    For intDo = 0 To intCount
        '列(必须要保留)
'        mstrSQL列 = Replace(mstrSQL列, ",'' AS C" & arrData(intDo), "")
        '条件
        'mstrSQL条件 = Replace(UCase(mstrSQL条件), " OR """ & "C" & Format(arrData(intDo), "00") & """ IS NOT NULL", "")
        mstrSQL条件 = Replace(UCase(Replace(UCase(mstrSQL条件), " OR """ & "C" & Format(arrData(intDo), "00") & """ IS NOT NULL", "")), """" & "C" & Format(arrData(intDo), "00") & """ IS NOT NULL OR ", "")
        mstrSQL条件 = Replace(UCase(mstrSQL条件), "(""" & "C" & Format(arrData(intDo), "00") & """ IS NOT NULL)", "")

        '中
        mstrSQL中 = Replace(mstrSQL中, ",MAX(""" & "C" & Format(arrData(intDo), "00") & """) AS C" & Format(arrData(intDo), "00"), "")
        '内
        mstrSQL内 = Replace(mstrSQL内, ", C" & Format(arrData(intDo), "00") & " AS C" & Format(arrData(intDo), "00"), "")
    Next
End Sub

Private Sub SQLCombination(Optional ByVal lng记录ID As Long = 0)
    Dim str条件 As String
    str条件 = mstrSQL条件 & IIf(lng记录ID = 0, "", IIf(mstrSQL条件 = "", "", " And") & " 记录ID=[6]")
    
    mstrSQL = "Select '' AS 备用,to_char(发生时间,'yyyy-MM-dd hh24:mi:ss') AS 发生时间,'' AS 选择,to_char(发生时间,'YYYY') AS 年份," & Mid(mstrSQL列, 12) & vbCrLf & _
                " From (Select 记录组号,时间 as 备用,发生时间," & Mid(mstrSQL中, 2) & vbCrLf & _
                "        From (Select nvl(c.记录组号,0) 记录组号,l.发生时间," & Mid(mstrSQL内, 2) & vbCrLf & _
                "               From 病人护理数据 l, 病人护理明细 c,病人护理文件 f,病人护理打印 p " & vbCrLf & _
                "               Where l.ID=p.记录ID And l.Id = c.记录id And l.文件ID+0=f.ID+0 And f.ID=p.文件ID " & _
                "               And c.终止版本 Is Null And c.记录类型<>5  " & _
                "               And f.id=[1] And f.病人id = [2] And f.主页id = [3] And Nvl(f.婴儿,0)=[4] " & _
                IIf(mintPageSpan = 0, " And (P.开始页号=[5] Or P.结束页号=[5])", " And P.开始页号=[5]") & ")" & vbCrLf & _
                IIf(str条件 <> "", "Where " & str条件, "") & _
                "       Group By 日期, 时间, 发生时间,记录组号," & IIf(mbln护士 = True, "护士,", "护士L,") & "签名人,签名时间" & _
                                "       Order By 发生时间,记录组号," & IIf(mbln护士 = True, "护士,", "护士L,") & "签名人,签名时间)"
End Sub

Private Sub zlRefresh()
    Dim aryRow() As String, aryItem() As String
    Dim strPrefix As String, strItemName As String
    Dim lngRow As Long, lngCol As Long, lngCount As Long, strCell As String, strBed As String
    Dim strTmpSQL As String
    Dim aryPeriod() As String
    Dim strtmp As String
    Dim blnReplace As Boolean
    
    Err = 0: On Error GoTo ErrHand
    
    Call InitCons
    mblnElement = False
    '表上标签获取
    lblSubhead.Caption = ""
    lblSubhead.Tag = ""
    aryPeriod = Split(GetPeriod, "～")
    '87057,病人10:30:20转科入住建立护理文件时间为10:30:20,此时录入首条数据为10:30(记录单无法录入秒),导致无法显示新的科室
    aryPeriod(0) = Format(aryPeriod(0), "YYYY-MM-DD HH:mm") & ":59"
    '获取当前页之前的最后科室ID
    gstrSQL = "Select 科室ID From 病人变动记录 " & _
        "   Where  病人ID=[1] And 主页ID=[2] And [3]>=开始时间 " & _
        " And 开始时间 IS NOT NULL And 科室id IS NOT NULL And NVL(附加床位,0)=0 Order by 开始时间 DESC"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取当前页之前的最后科室ID", mlng病人ID, mlng主页ID, CDate(aryPeriod(0)))
    If rsTemp.RecordCount > 0 Then mlng科室ID = Val(rsTemp!科室ID)
    
    gstrSQL = "Select [1] || Zl_Replace_Element_Value([2],[3],[4],2,NULL,[5],[6]) as 信息 From Dual"
    aryItem = Split(mstrSubhead, "|")
    
        
    For lngCount = 0 To UBound(aryItem)
        strPrefix = Left(aryItem(lngCount), InStr(1, aryItem(lngCount), "{") - 1)
        strItemName = Mid(aryItem(lngCount), InStr(1, aryItem(lngCount), "{") + 1, InStr(1, aryItem(lngCount), "}") - InStr(1, aryItem(lngCount), "{") - 1)
        
        strtmp = strPrefix
        strCell = ""
        '68336
        blnReplace = True
        mrsElement.Filter = "中文名='" & strItemName & "'"
        If mrsElement.RecordCount > 0 Then
            blnReplace = Val(NVL(mrsElement!替换域, 0)) = 1
        End If
        Select Case strItemName
        Case "当前病区"
        
            strTmpSQL = "Select   b.名称" & vbNewLine & _
                        "From (Select 病区id, 开始时间, Nvl(终止时间, To_Date('3000-01-01', 'yyyy-mm-dd')) As 终止时间" & vbNewLine & _
                        "            From 病人变动记录" & vbNewLine & _
                        "            Where 病人id = [1] And 主页id = [2] And 科室id = [3] And NVL(附加床位,0)=0 And 开始时间 IS NOT NULL) a,部门表 b " & vbNewLine & _
                        "Where ([4] Between a.开始时间 And a.终止时间 Or [4] >= a.开始时间) And a.病区id Is Not Null And b.ID=a.病区id" & vbNewLine & _
                        "Order By a.开始时间"
                        
            Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "当前病区", mlng病人ID, mlng主页ID, mlng科室ID, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
            If rsTemp.BOF = False Then rsTemp.MoveLast
        Case "当前床号"

            strTmpSQL = "Select   a.床号" & vbNewLine & _
                        "From (Select 床号, 开始时间, Nvl(终止时间, To_Date('3000-01-01', 'yyyy-mm-dd')) As 终止时间" & vbNewLine & _
                        "            From 病人变动记录" & vbNewLine & _
                        "            Where 病人id = [1] And 主页id = [2] And 科室id = [3] And NVL(附加床位,0)=0 And 开始时间 IS NOT NULL) a" & vbNewLine & _
                        "Where ([4] Between a.开始时间 And a.终止时间 Or [4] >= a.开始时间) And a.床号 Is Not Null" & vbNewLine & _
                        "Order By a.开始时间"

            Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "当前床号", mlng病人ID, mlng主页ID, mlng科室ID, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
            If rsTemp.BOF = False Then rsTemp.MoveLast
        Case "床位变动"
            strTmpSQL = "Select   a.床号" & vbNewLine & _
                        "From (Select 床号, 开始时间, Nvl(终止时间, To_Date('3000-01-01', 'yyyy-mm-dd')) As 终止时间" & vbNewLine & _
                        "            From 病人变动记录" & vbNewLine & _
                        "            Where 病人id = [1] And 主页id = [2] And 科室id = [3] And NVL(附加床位,0)=0 And 开始时间 IS NOT NULL) a" & vbNewLine & _
                        "Where (a.终止时间>=[4] And a.开始时间<=[5]) And a.床号 Is Not Null" & vbNewLine & _
                        "Order By a.开始时间"

            Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "当前床号", mlng病人ID, mlng主页ID, mlng科室ID, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
            strCell = "": strBed = ""
            Do While Not rsTemp.EOF
                If strBed <> rsTemp.Fields(0).Value Then
                    strBed = rsTemp.Fields(0).Value
                    strCell = strCell & "->" & rsTemp.Fields(0).Value
                End If
            rsTemp.MoveNext
            Loop
            strCell = Mid(strCell, 3)
            If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
        Case "当前科室"
        
            strTmpSQL = "Select   名称 From 部门表 a Where a.ID=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "当前科室", mlng科室ID)
            
        Case "住院医师"
            strTmpSQL = "Select   a.经治医师" & vbNewLine & _
                        "From (Select 经治医师, 开始时间, Nvl(终止时间, To_Date('3000-01-01', 'yyyy-mm-dd')) As 终止时间" & vbNewLine & _
                        "            From 病人变动记录" & vbNewLine & _
                        "            Where 病人id = [1] And 主页id = [2] And 科室id = [3] And NVL(附加床位,0)=0 And 开始时间 IS NOT NULL) a" & vbNewLine & _
                        "Where ([4] Between a.开始时间 And a.终止时间 Or [4] >= a.开始时间) And a.经治医师 Is Not Null" & vbNewLine & _
                        "Order By a.开始时间"
            Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "住院医师", mlng病人ID, mlng主页ID, mlng科室ID, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
            If rsTemp.BOF = False Then rsTemp.MoveLast
        Case "责任护士"
        
            strTmpSQL = "Select   a.责任护士" & vbNewLine & _
                        "From (Select 责任护士, 开始时间, Nvl(终止时间, To_Date('3000-01-01', 'yyyy-mm-dd')) As 终止时间" & vbNewLine & _
                        "            From 病人变动记录" & vbNewLine & _
                        "            Where 病人id = [1] And 主页id = [2] And 科室id = [3] And NVL(附加床位,0)=0 And 开始时间 IS NOT NULL) a" & vbNewLine & _
                        "Where ([4] Between a.开始时间 And a.终止时间 Or [4] >= a.开始时间) And a.责任护士 Is Not Null" & vbNewLine & _
                        "Order By a.开始时间"
            Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "责任护士", mlng病人ID, mlng主页ID, mlng科室ID, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
            If rsTemp.BOF = False Then rsTemp.MoveLast
            
        Case "护理等级"
            strTmpSQL = "Select   b.名称" & vbNewLine & _
                        "From (Select 护理等级ID, 开始时间, Nvl(终止时间, To_Date('3000-01-01', 'yyyy-mm-dd')) As 终止时间" & vbNewLine & _
                        "            From 病人变动记录" & vbNewLine & _
                        "            Where 病人id = [1] And 主页id = [2] And 科室id = [3] And NVL(附加床位,0)=0 And 开始时间 IS NOT NULL) a,护理等级 b" & vbNewLine & _
                        "Where ([4] Between a.开始时间 And a.终止时间 Or [4] >= a.开始时间) And a.护理等级ID Is Not Null And b.序号=a.护理等级ID" & vbNewLine & _
                        "Order By a.开始时间"
            Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "护理等级", mlng病人ID, mlng主页ID, mlng科室ID, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
            If rsTemp.BOF = False Then rsTemp.MoveLast
        Case "最后诊断"
            strtmp = ""
            gstrSQL = "Select [1] || Zl_Replace_Element_Value([2],[3],[4],2,NULL,[5],[6]) as 信息 From Dual"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取要素", strPrefix, strItemName, mlng病人ID, mlng主页ID, mint婴儿, CDate(aryPeriod(0)))
        Case Else
            strtmp = ""
            If blnReplace = True Then
                gstrSQL = "Select [1] || Zl_Replace_Element_Value([2],[3],[4],2,NULL,[5],[6]) as 信息 From Dual"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取要素", strPrefix, strItemName, mlng病人ID, mlng主页ID, mint婴儿, CDate(aryPeriod(0)))
            Else
                mblnElement = True
                strtmp = strPrefix
                gstrSQL = "Select 内容 From 病人护理要素内容 Where 文件ID=[1] And 页号=[2] And 名称=[3]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取要素", mlng文件ID, mint页码, strItemName)
            End If
        End Select
        
        If rsTemp.BOF = False Then
            If strCell = "" Then
                If strtmp <> "" Then
                    lblSubhead.Tag = lblSubhead.Tag & " " & strtmp & rsTemp.Fields(0).Value
                Else
                    lblSubhead.Tag = lblSubhead.Tag & " " & rsTemp.Fields(0).Value
                End If
            Else
                lblSubhead.Tag = lblSubhead.Tag & " " & strtmp & strCell
            End If
        End If
    Next
    lblSubhead.Tag = Trim(lblSubhead.Tag)
    
    '表上标签分散处理
    Call zlLableBruit
    
    '产生列记录集
    Call InitRecords
    
    '装入数据
    Call SQLCombination
    gstrSQL = mstrSQL
    Call SQLDIY(gstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取护理数据", mlng文件ID, mlng病人ID, mlng主页ID, mint婴儿, mint页码)
    '清除并拷贝记录集结构
    Call DataMap_Init(rsTemp)
    '绑定数据并设置护理记录单的格式,同时实现一行数据分行显示的功能
    Call PreTendFormat(rsTemp)
    Call cbsThis_Resize
    
    lblCurPage.Caption = "P" & mint页码
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub DataMap_Init(ByVal rsSource As ADODB.Recordset)
    '初始化内存数据集
    
    If Not mblnClear Then Exit Sub
    
    '数据记录集,用于快速恢复
    Set mrsDataMap = CopyNewRec(rsSource)
    mrsDataMap.Sort = "页号,行号"
    '修改单元格记录,用于保存(标记主要用于保存导入入量的医嘱信息:医嘱ID:发送号 )
    Call Record_Init(mrsCellMap, "ID," & adLongVarChar & ",50|页号," & adDouble & ",18|行号," & adDouble & ",18|" & _
            "列号," & adDouble & ",18|起始行号," & adDouble & ",18|记录ID," & adDouble & ",18|数据," & adLongVarChar & ",4000|部位," & adLongVarChar & ",100|" & _
            "标记," & adLongVarChar & ",100|汇总," & adDouble & ",1|记录组号," & adDouble & ",1|删除," & adDouble & ",1")
    mrsCellMap.Sort = "页号,行号,列号"
    '复制记录集
    Set mrsCopyMap = New ADODB.Recordset
    Set mrsCopyMap = CopyNewRec(mrsDataMap, False)
    
    '为了不影响之后的换页,将此参数设置为假
    mblnClear = False
End Sub

Private Function DataMap_Save() As Boolean
    '将当前页面中用户编辑过的数据保存起来,页面切换或保存前触发
    Dim blnExit As Boolean
    Dim lngRow As Long, lngCol As Long, lngRows As Long, lngCols As Long
    On Error GoTo ErrHand
    
    '不管是否编辑过都保存
'    '如果当前页未编辑过,则不必保存
'    mrsCellMap.Filter = "页号=" & mint页码
'    blnExit = (mrsCellMap.RecordCount = 0)
'    If blnExit Then
'        mrsCellMap.Filter = 0
'        DataMap_Save = True
'        Exit Function
'    End If
'    mrsCellMap.Filter = 0
    
    If Not CheckFlip Then Exit Function
    
    '先删除指定页号的所有数据行
    mrsDataMap.Filter = "页号=" & mint页码
    Do While True
        If mrsDataMap.RecordCount = 0 Then Exit Do
        mrsDataMap.Delete
        mrsDataMap.MoveNext
    Loop
    mrsDataMap.Filter = 0
    
    '复制指定页号的所有数据行
    lngRows = VsfData.Rows - 1
    lngCols = VsfData.Cols - 1
    For lngRow = VsfData.FixedRows To lngRows
        mrsDataMap.AddNew
        mrsDataMap!页号 = mint页码
        mrsDataMap!行号 = lngRow
        mrsDataMap!删除 = IIf(VsfData.RowHidden(lngRow), 1, 0)
        For lngCol = 0 To lngCols - VsfData.FixedCols
            If lngCol + VsfData.FixedCols = mlngChoose Then
                mrsDataMap.Fields(cControlFields + lngCol).Value = VsfData.Cell(flexcpChecked, lngRow, mlngChoose)
            ElseIf InStr(1, "," & mlngCollectType & "," & mlngRecord & ",", "," & lngCol + VsfData.FixedCols & ",") <> 0 Then
                mrsDataMap.Fields(cControlFields + lngCol).Value = Val(FormatValue(VsfData.TextMatrix(lngRow, lngCol + VsfData.FixedCols)))
            Else
                mrsDataMap.Fields(cControlFields + lngCol).Value = IIf(FormatValue(VsfData.TextMatrix(lngRow, lngCol + VsfData.FixedCols)) = "", Null, FormatValue(VsfData.TextMatrix(lngRow, lngCol + VsfData.FixedCols)))
            End If
        Next
        mrsDataMap.Update
        
        mrsCellMap.Filter = "页号=" & mint页码 & " And 行号=" & lngRow
        If mrsCellMap.RecordCount > 0 And VsfData.RowHidden(lngRow) = True Then
            Do While Not mrsCellMap.EOF
                If mrsCellMap!列号 > mlngTime And mrsCellMap!列号 < mlngNoEditor Then
                    mrsCellMap!数据 = ""
                    mrsCellMap.Update
                End If
            mrsCellMap.MoveNext
            Loop
        End If
    Next
    mrsCellMap.Filter = ""
    DataMap_Save = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function DataMap_Restore(ByVal rsTemp As ADODB.Recordset) As Boolean
    '将指定页面的数据恢复到表格中
    Dim lngRow As Long, lngCol As Long, lngRows As Long, lngCols As Long
    On Error GoTo ErrHand
    
    mblnRestore = False
    If VsfData.Rows > VsfData.FixedRows Then
        VsfData.Cell(flexcpChecked, VsfData.FixedRows, mlngChoose, VsfData.Rows - 1, mlngChoose) = flexTSUnchecked
    End If
    '复制指定页号的所有数据行到表格中
    mrsDataMap.Filter = "页号=" & mint页码
    lngRows = mrsDataMap.RecordCount
    
    If lngRows = 0 Then
        '没有修改过的数据则绑定读取的记录集
        mrsDataMap.Filter = 0
        Set VsfData.DataSource = rsTemp
        DataMap_Restore = True
        Exit Function
    Else
        '此处只需要绑定一个空的记录集即可(恢复数据)
        Set VsfData.DataSource = rsTemp
        VsfData.Rows = VsfData.FixedRows
        mblnRestore = True
    End If
    
    mrsDataMap.MoveFirst
    lngCols = VsfData.Cols - 1
    For lngRow = 0 To lngRows - 1
        If lngRow > VsfData.Rows - VsfData.FixedRows - 1 Then VsfData.Rows = VsfData.Rows + 1
        For lngCol = 0 To lngCols - VsfData.FixedCols
            If lngCol + VsfData.FixedCols = mlngChoose Then
                If InStr(1, "3,4", NVL(mrsDataMap.Fields(cControlFields + lngCol).Value, 0)) <> 0 Then
                    VsfData.Cell(flexcpChecked, VsfData.FixedRows + lngRow, lngCol + VsfData.FixedCols) = NVL(mrsDataMap.Fields(cControlFields + lngCol).Value)
                End If
            Else
                VsfData.TextMatrix(VsfData.FixedRows + lngRow, lngCol + VsfData.FixedCols) = NVL(mrsDataMap.Fields(cControlFields + lngCol).Value)
            End If
        Next
        If mrsDataMap!删除 = 1 Then VsfData.RowHidden(VsfData.FixedRows + lngRow) = True
        
        mrsDataMap.MoveNext
    Next
    
    mrsDataMap.Filter = 0
    DataMap_Restore = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub CellMap_Update(ByVal lngStart As Long, ByVal lngDeff As Long, Optional ByVal blnBig As Boolean = True)
    Dim lngPos As Long
    Dim intCol As Integer
    
    '更新当前页面所有大于起始行的行号数据
    With mrsCellMap
        If lngDeff > 0 Then
            If .RecordCount = 0 Then Exit Sub
            If .RecordCount <> 0 Then .MoveLast
            If .BOF Then Exit Sub
            Do While Not mrsCellMap.BOF
                If !页号 = mint页码 And IIf(blnBig = True, !行号 > lngStart, !行号 = lngStart) Then
                    intCol = !列号
                    lngPos = .AbsolutePosition
                    !行号 = !行号 + lngDeff
                    !ID = mint页码 & "," & !行号 & "," & !列号
                    .Update
                    .MoveFirst
                    .Move lngPos - 2
                Else
                    .MovePrevious
                End If
            Loop
        ElseIf lngDeff < 0 Then
            If .RecordCount = 0 Then Exit Sub
            If .RecordCount <> 0 Then .MoveFirst
            If .EOF Then Exit Sub
            Do While Not mrsCellMap.EOF
                If !页号 = mint页码 And IIf(blnBig = True, !行号 > lngStart, !行号 = lngStart) Then
                    intCol = !列号
                    lngPos = .AbsolutePosition
                    !行号 = !行号 + lngDeff
                    !ID = mint页码 & "," & !行号 & "," & !列号
                    .Update
                    .MoveFirst
                    .Move lngPos
                Else
                    .MoveNext
                End If
            Loop
        End If
        If .RecordCount <> 0 Then .MoveFirst
    End With
End Sub

Public Function CopyNewRec(ByVal rsSource As ADODB.Recordset, Optional ByVal blnAddPage As Boolean = True) As ADODB.Recordset
    '只拷贝记录集的结构,同时增加页号,行号字段
    Dim rsTarget As New ADODB.Recordset
    Dim intFields As Integer
    
    Set rsTarget = New ADODB.Recordset
    With rsTarget
        If blnAddPage Then
            .Fields.Append "页号", adDouble, 18
            .Fields.Append "行号", adDouble, 18
        End If
        For intFields = 0 To rsSource.Fields.Count - 1
            If rsSource.Fields(intFields).Name = "汇总日期" Then
                .Fields.Append rsSource.Fields(intFields).Name, adLongVarChar, 50, adFldIsNullable      '0:表示新增
            ElseIf rsSource.Fields(intFields).Type = 200 Then       '日期型处理为字符型
                .Fields.Append rsSource.Fields(intFields).Name, adLongVarChar, rsSource.Fields(intFields).DefinedSize, adFldIsNullable     '0:表示新增
            Else
                .Fields.Append rsSource.Fields(intFields).Name, IIf(rsSource.Fields(intFields).Type = adNumeric, adDouble, rsSource.Fields(intFields).Type), rsSource.Fields(intFields).DefinedSize, adFldIsNullable    '0:表示新增
            End If
        Next
        If blnAddPage Then
            .Fields.Append "删除", adDouble, 1
        End If
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    Set CopyNewRec = rsTarget
End Function

Private Sub PreTendMutilRows()
    Dim arrData
    Dim intData As Integer, intDatas As Integer
    Dim lngRowCount As Long, lngRowCurrent As Long  '当前记录总行数,当前记录在本页的实际行数
    Dim lngCol As Long, lngMax As Long, lngRecordId As Long
    Dim lngRow As Long
    Dim str发生时间 As String, str发生时间_L As String, lngLastRow As Long
    Dim lngLiterrRow As Long
    Dim lngTestRow As Long, lngStartRow As Long
    Dim strDate As String
    Dim intCol As Integer, intCols As Integer
    Dim rsData As New ADODB.Recordset
    Dim strSignName As String
    Dim lngPrintedRow As Long, lngStart As Long
    Dim blnClear As Boolean
    Dim lngCount As Long
    Dim blnCollectType As Boolean  '记录正常数据行的上一行是否是汇总行
    Dim lngCurrRow As Long, lngCollectMutilRows As Long '汇总数据当前行、汇总列数据的行数
    Dim i As Integer, j As Integer, arrItem, arrCorrelative, arrLastRow, arrMutilRows '分类汇总项目数组
    On Error GoTo ErrHand
    
    arrItem = Split(mstrColCorrelative, "|")
    '如果一行显示不完则分行显示(根据当前数据占用行数先添加空白行并处理行坐标,然后再依次处理当前行的数据)
    '每页只显示实际的数据行,把'@处取消注释即可
    lngRow = VsfData.FixedRows
    Do While True
        If lngRow > VsfData.Rows - 1 Then Exit Do
        If InStr(1, FormatValue(VsfData.TextMatrix(lngRow, mlngRowCount)), "|") <> 0 Then Exit Do
        
        lngRowCount = Val(FormatValue(VsfData.TextMatrix(lngRow, mlngRowCount)))
        '@实际数据行
'        lngRowCurrent = Val(FormatValueVsfData.TextMatrix(lngRow, mlngRowCurrent)))
        
        str发生时间 = Format(VsfData.TextMatrix(lngRow, mlngActTime), "YYYY-MM-DD HH:mm:ss")
        If Val(VsfData.TextMatrix(lngRow, mlngCollectType)) < 0 Then
            '在处理分类汇总首条记录时(总量)，完成所有所有列的赋值 blnCollectType=false的情况
            If blnCollectType = False Then str发生时间_L = "": blnCollectType = True
            '分类汇总子类明细数据的处理(数据保存方式是一条护理数据对应多条明细,明细中的记录组号不同)
            If str发生时间_L <> "" And str发生时间_L = str发生时间 Then
                If UBound(arrItem) < 0 Then '如果当前没有设置汇总关系,但之前数据存在分类汇总的情况，按子分类条数循环处理
                    lngCurrRow = lngLastRow + lngCollectMutilRows '确定每一条子数据输出的起始位置
                    lngCollectMutilRows = 1
                    If lngCurrRow < lngRow Then
                        VsfData.TextMatrix(lngCurrRow, mlngYear) = ""
                        VsfData.TextMatrix(lngCurrRow, mlngDate) = ""
                        VsfData.TextMatrix(lngCurrRow, mlngTime) = ""
                        
                        For lngCol = mlngTime + 1 To mlngNoEditor - 1
                            If (lngCol <> mlngSignTime And VsfData.ColHidden(lngCol) = False) Then
                                '准备赋值
                                With txtLength
                                    .Width = VsfData.ColWidth(lngCol)
                                    '这里需要注意一点：提取子类数据的行数应该是lngRow而不是lngCurrRow，因为在处理汇总总量记录时会导致子类数据的行位置发生变化 (行数 = 主记录开始行号 + 数据行数)
                                    .Text = Replace(Replace(Replace(VsfData.TextMatrix(lngRow, lngCol), Chr(10), ""), Chr(13), ""), Chr(1), "")
                                    .FontName = VsfData.CellFontName
                                    .FontSize = VsfData.CellFontSize
                                    .FontBold = VsfData.CellFontBold
                                    .FontItalic = VsfData.CellFontItalic
                                End With
                                arrData = GetData(txtLength.Text)
                                intDatas = UBound(arrData)
                                
                                If intDatas >= 0 Then
                                    '循环赋值
                                    If intDatas + 1 > lngRow - lngCurrRow Then intDatas = lngRow - lngCurrRow - 1
                                    If lngCollectMutilRows < intDatas + 1 Then lngCollectMutilRows = intDatas + 1
                                    For intData = 0 To intDatas
                                        VsfData.TextMatrix(lngCurrRow + intData, lngCol) = Replace(Replace(Replace(arrData(intData), Chr(10), ""), Chr(13), ""), Chr(1), "")
                                    Next
                                End If
                            End If
                        Next lngCol
                    End If
                    lngLastRow = lngCurrRow
                Else
                    '设置了分类汇总关系，按照每个汇总项目依次展示数据
                    For i = 0 To UBound(arrItem)
                        lngCurrRow = Val(arrLastRow(i)) + Val(arrMutilRows(i)) '按项目分类,确定每条数据输出的起始位置
                        lngCollectMutilRows = 1
                        arrMutilRows(i) = lngCollectMutilRows
                        If lngCurrRow < lngRow Then
                            arrCorrelative = Split(arrItem(i), ";")
                            For j = 0 To 1
                                '准备赋值
                                    lngCol = Split(arrCorrelative(j), ",")(0) + cHideCols + VsfData.FixedCols - 1
                                    With txtLength
                                        .Width = VsfData.ColWidth(lngCol)
                                        '这里需要注意一点：提取子类数据的行数应该是lngRow而不是lngCurrRow，因为在处理汇总总量记录时会导致子类数据的行位置发生变化 (行数 = 主记录开始行号 + 数据行数)
                                        .Text = Replace(Replace(Replace(VsfData.TextMatrix(lngRow, lngCol), Chr(10), ""), Chr(13), ""), Chr(1), "")
                                        .FontName = VsfData.CellFontName
                                        .FontSize = VsfData.CellFontSize
                                        .FontBold = VsfData.CellFontBold
                                        .FontItalic = VsfData.CellFontItalic
                                    End With
                                    arrData = GetData(txtLength.Text)
                                    intDatas = UBound(arrData)
                                    
                                    If intDatas >= 0 Then
                                        If intDatas + 1 > lngRow - lngCurrRow Then intDatas = lngRow - lngCurrRow - 1
                                        If lngCollectMutilRows < intDatas + 1 Then lngCollectMutilRows = intDatas + 1
                                        arrMutilRows(i) = lngCollectMutilRows
                                        For intData = 0 To intDatas
                                            VsfData.TextMatrix(lngCurrRow + intData, lngCol) = Replace(Replace(Replace(arrData(intData), Chr(10), ""), Chr(13), ""), Chr(1), "")
                                        Next intData
                                    End If
                            Next j
                        End If
                        arrLastRow(i) = lngCurrRow
                    Next i
                End If
                '赋值完成后移除原有子类行
                VsfData.RowPosition(lngRow) = VsfData.Rows - 1
                VsfData.RemoveItem VsfData.Rows - 1
                GoTo NextData
            Else
                'If lngRow >= mlngPageRows + mlngOverrunRows - mlngReduceRow + VsfData.FixedRows Then Exit Do
                '总量行默认为一行(只是针对汇总列的数据)
                lngCollectMutilRows = 1
                lngLastRow = lngRow '记录分类汇总总量行的位置
                '确定分类汇总首条子分类数据每个汇总项目的起始位置
                arrLastRow = Array(): arrMutilRows = Array()
                For i = 0 To UBound(arrItem)
                    ReDim Preserve arrLastRow(UBound(arrLastRow) + 1)
                    arrLastRow(UBound(arrLastRow)) = lngLastRow
                    ReDim Preserve arrMutilRows(UBound(arrMutilRows) + 1)
                    arrMutilRows(UBound(arrMutilRows)) = lngCollectMutilRows
                Next i
            End If
        Else
            'If lngRow >= mlngPageRows + mlngOverrunRows - mlngReduceRow + VsfData.FixedRows Then Exit Do
            If blnCollectType = True Then str发生时间_L = "": blnCollectType = False
            If str发生时间_L <> "" And Mid(str发生时间_L, 1, 16) = Mid(str发生时间, 1, 16) And str发生时间_L <> str发生时间 Then
                '日期相同，秒数不同，且不是汇总数据行，则说明这些数据是一组，更新lngDemo列
                VsfData.TextMatrix(lngRow, mlngYear) = ""
                VsfData.TextMatrix(lngRow, mlngDate) = ""
                VsfData.TextMatrix(lngRow, mlngTime) = ""
                VsfData.TextMatrix(lngRow, mlngDemo) = lngRow - lngLastRow + 1
                If lngRow - lngLastRow = Val(FormatValue(VsfData.TextMatrix(lngLastRow, mlngRowCount))) Then
                    VsfData.TextMatrix(lngLastRow, mlngDemo) = 1
                End If
            Else
                lngLastRow = lngRow
            End If
        End If
        If lngRowCount > 1 Then
            '先增加空行
            VsfData.Rows = VsfData.Rows + lngRowCount - 1
            '从当前行的下一行开始，每行的位置+所增加的空白行数，保证新增的空白行从当前行的下一行开始
            For intData = VsfData.Rows - lngRowCount To lngRow + 1 Step -1
                VsfData.RowPosition(intData) = intData + lngRowCount - 1
            Next
            
            '循环处理当前行数据
            For lngCol = 0 To VsfData.Cols - 1
                If VsfData.ColHidden(lngCol) And lngCol <> mlngRowCount And lngCol <> mlngDemo Then
                    '循环赋值
                    For intData = 2 To lngRowCount
                        VsfData.TextMatrix(lngRow + intData - 1, lngCol) = FormatValue(VsfData.TextMatrix(lngRow, lngCol))
                    Next
                ElseIf (lngCol < mlngNoEditor And lngCol <> mlngDate And lngCol <> mlngTime And lngCol <> mlngYear) Then
                    '准备赋值
                    With txtLength
                        .Width = VsfData.ColWidth(lngCol)
                        .Text = Replace(Replace(Replace(VsfData.TextMatrix(lngRow, lngCol), Chr(10), ""), Chr(13), ""), Chr(1), "")
                        .FontName = VsfData.CellFontName
                        .FontSize = VsfData.CellFontSize
                        .FontBold = VsfData.CellFontBold
                        .FontItalic = VsfData.CellFontItalic
                    End With
                    arrData = GetData(txtLength.Text)
                    intDatas = UBound(arrData)
                    
                    If intDatas > 0 Then
                        '循环赋值
                        If intDatas + 1 > lngRowCount Then intDatas = lngRowCount - 1
                        For intData = 0 To intDatas
                            If VsfData.Rows <= lngRow + intData Then VsfData.Rows = VsfData.Rows + 1
                            VsfData.TextMatrix(lngRow + intData, lngCol) = Replace(Replace(Replace(arrData(intData), Chr(10), ""), Chr(13), ""), Chr(1), "")
                        Next
                    End If
                ElseIf lngCol = mlngNoEditor Then
                    '将行值改为从1开始,比如有4行数据,就是4|1
                    For intData = 1 To lngRowCount
                        VsfData.TextMatrix(lngRow + intData - 1, mlngRowCount) = lngRowCount & "|" & intData
                    Next
                    '最后一行需要填写封闭签名
                    If mlngSignName > 0 Then VsfData.TextMatrix(lngRow + lngRowCount - 1, mlngSignName) = FormatValue(VsfData.TextMatrix(lngRow, mlngSignName))
                    If mlngSignTime > 0 Then VsfData.TextMatrix(lngRow + lngRowCount - 1, mlngSignTime) = FormatValue(VsfData.TextMatrix(lngRow, mlngSignTime))
                    '--58414,刘鹏飞,2013-04-10,添加护士、签名列显示模式
                    Call SingerShowType(VsfData, lngRow, lngRow + lngRowCount - 1)
                Else
                    
                End If
            Next
            '@实际数据行
'            '如果本页第一行的数据不全,则先将该记录第一行的主数据(日期,时间,签名)信息复制到
'            If lngRow = VsfData.FixedRows And lngRowCount <> lngRowCurrent Then
'                '固定复制显示日期时间与签名列
'                lngMax = lngRowCount - lngRowCurrent
'                If mlngDate > -1 Then VsfData.TextMatrix(lngRow + lngMax, mlngDate) = VsfData.TextMatrix(lngRow, mlngDate)
'                If mlngTime > -1 Then VsfData.TextMatrix(lngRow + lngMax, mlngTime) = VsfData.TextMatrix(lngRow, mlngTime)
'                if mlngOperator <>-1 then VsfData.TextMatrix(lngRow + lngMax, mlngOperator) = VsfData.TextMatrix(lngRow, mlngOperator)
'                if mlngSignName <>-1 then VsfData.TextMatrix(lngRow + lngMax, mlngsignname) = VsfData.TextMatrix(lngRow, mlngsignname)
'                '删除多余的行
'                For lngCol = 1 To lngMax
'                    VsfData.RemoveItem lngRow
'                Next
'            End If
'            lngRow = lngRow + lngRowCurrent - 1 '加上该记录在本页实际的行数
            '@实际数据行要注释下面这行代码
            lngRow = lngRow + lngRowCount - 1
        Else
            VsfData.TextMatrix(lngRow, mlngRowCount) = "1|1"
            VsfData.TextMatrix(lngRow, mlngRowCurrent) = "1"
        End If
        lngRow = lngRow + 1
NextData:
        str发生时间_L = str发生时间
    Loop
    If mblnRestore Then Exit Sub
    
    'Modified by zyb 2011-09-15
    'Modify by LPF 2012-05-08
    '分组数据,只能在上一页完整显示并编辑,下一页不显示跨页的分组数据
    '检查当前页首是否为分组数据,是则删除这些分组数据(最后来处理)
    '检查最后的数据是否为分组数据,是则读取下一页,将跨页的部分分组数据组装在一起
    'If Val(VsfData.TextMatrix(VsfData.Rows - 1, mlngDemo)) > 0 And VsfData.Rows - VsfData.FixedRows >= mlngPageRows Then
    lngLiterrRow = 0
    mlngLitterRows(mint页码) = 0
    mlngCurLitterRows(mint页码) = 0
    mArrPageInfo(mint页码) = ""
    If VsfData.Rows > VsfData.FixedRows And Val(FormatValue(VsfData.TextMatrix(VsfData.Rows - 1, mlngRowCount))) > 0 And mint页码 >= mint起始页码 Then
        intCols = VsfData.Cols - 1
        lngTestRow = VsfData.FixedRows
        '获取数据起始行
        If Val(FormatValue(VsfData.TextMatrix(VsfData.Rows - 1, mlngRowCount))) = 1 Then
            lngStartRow = VsfData.Rows - 1
        Else
            lngStartRow = GetStartRow(VsfData.Rows - 1)
        End If
        lngRecordId = Val(FormatValue(VsfData.TextMatrix(lngStartRow, mlngRecord)))
        strDate = Format(FormatValue(VsfData.TextMatrix(lngStartRow, mlngActTime)), "YYYY-MM-DD HH:mm:ss")
        blnCollectType = Val(FormatValue(VsfData.TextMatrix(lngStartRow, mlngCollectType))) < 0
        Call SQLCombination
        
        gstrSQL = mstrSQL
        Call SQLDIY(gstrSQL)
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "提取护理数据", mlng文件ID, mlng病人ID, mlng主页ID, mint婴儿, mint页码 + 1)
        If rsData.RecordCount > 0 Then
            Set vsTest.DataSource = rsData
            Do While True
                If lngTestRow > vsTest.Rows - 1 Then Exit Do
                If Mid(strDate, 1, 16) <> Mid(Format(vsTest.TextMatrix(lngTestRow, mlngActTime), "YYYY-MM-DD HH:mm:ss"), 1, 16) Then Exit Do
                '82036:LPF,避免将汇总数据纳入分组数据中
                If Val(vsTest.TextMatrix(lngTestRow, mlngCollectType)) < 0 Or blnCollectType = True Then Exit Do
                
                If lngRecordId = Val(vsTest.TextMatrix(lngTestRow, mlngRecord)) Then GoTo ErrNext
                lngRowCount = Val(vsTest.TextMatrix(lngTestRow, mlngRowCount))
                VsfData.Rows = VsfData.Rows + lngRowCount
                lngLiterrRow = lngLiterrRow + lngRowCount
                
                For intCol = 0 To intCols
                    VsfData.TextMatrix(lngRow, intCol) = vsTest.TextMatrix(lngTestRow, intCol)
                Next
               '循环处理当前行数据
                For lngCol = 0 To VsfData.Cols - 1
                    If VsfData.ColHidden(lngCol) And lngCol <> mlngRowCount And lngCol <> mlngDemo Then
                        '循环赋值
                        For intData = 2 To lngRowCount
                            VsfData.TextMatrix(lngRow + intData - 1, lngCol) = vsTest.TextMatrix(lngTestRow, lngCol)
                        Next
                    ElseIf (lngCol < mlngNoEditor And lngCol <> mlngDate And lngCol <> mlngTime And lngCol <> mlngYear) Then
                        '准备赋值
                        With txtLength
                            .Width = VsfData.ColWidth(lngCol)
                            .Text = Replace(Replace(Replace(vsTest.TextMatrix(lngTestRow, lngCol), Chr(10), ""), Chr(13), ""), Chr(1), "")
                            .FontName = VsfData.CellFontName
                            .FontSize = VsfData.CellFontSize
                            .FontBold = VsfData.CellFontBold
                            .FontItalic = VsfData.CellFontItalic
                        End With
                        arrData = GetData(txtLength.Text)
                        intDatas = UBound(arrData)
                        
                        If intDatas > 0 Then
                            '循环赋值
                            If intDatas + 1 > lngRowCount Then intDatas = lngRowCount - 1
                            For intData = 0 To intDatas
                                If VsfData.Rows <= lngRow + intData Then VsfData.Rows = VsfData.Rows + 1
                                VsfData.TextMatrix(lngRow + intData, lngCol) = arrData(intData)
                            Next
                        End If
                    ElseIf lngCol = mlngNoEditor Then
                        '将行值改为从1开始,比如有4行数据,就是4|1
                        For intData = 1 To lngRowCount
                            VsfData.TextMatrix(lngRow + intData - 1, mlngRowCount) = lngRowCount & "|" & intData
                            VsfData.TextMatrix(lngRow, mlngYear) = ""
                            VsfData.TextMatrix(lngRow, mlngDate) = ""
                            VsfData.TextMatrix(lngRow, mlngTime) = ""
                        Next
                        '最后一行需要填写封闭签名
                        If mlngSignName > 0 Then VsfData.TextMatrix(lngRow + lngRowCount - 1, mlngSignName) = vsTest.TextMatrix(lngTestRow, mlngSignName)
                        If mlngSignTime > 0 Then VsfData.TextMatrix(lngRow + lngRowCount - 1, mlngSignTime) = vsTest.TextMatrix(lngTestRow, mlngSignTime)
                        '--58414,刘鹏飞,2013-04-10,添加护士、签名列显示模式
                        Call SingerShowType(VsfData, lngRow, lngRow + lngRowCount - 1)
                    End If
                Next

                VsfData.TextMatrix(lngRow, mlngDemo) = lngRow - lngLastRow + 1
                If lngRow - lngLastRow = Val(FormatValue(VsfData.TextMatrix(lngLastRow, mlngRowCount))) Then
                    VsfData.TextMatrix(lngLastRow, mlngDemo) = 1
                End If
                lngRow = lngRow + lngRowCount
ErrNext:
                lngTestRow = lngTestRow + 1
            Loop
        End If
        
        Set vsTest.DataSource = Nothing
        vsTest.Clear
        rsData.Close
        Set rsData = Nothing
    End If
    
    If lngLiterrRow <> 0 Then
        mArrPageInfo(mint页码) = mArrPageInfo(mint页码) & "[LPF]" & "当前页发生时间:" & Mid(strDate, 1, 16) & "的分组数据中有" & lngLiterrRow & "行数据为下一页的数据。"
    End If
    
    If mint页码 > mint起始页码 Then
        Call SQLCombination
        gstrSQL = mstrSQL
        Call SQLDIY(gstrSQL)
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "提取护理数据", mlng文件ID, mlng病人ID, mlng主页ID, mint婴儿, mint页码 - 1)
        If rsData.RecordCount > 0 Then
            Set vsTest.DataSource = rsData
            
            '跨页数据显示在当前行时，由于下一页无法提取到跨页数据，下一页的实际行数应该等于本页数据行+本页跨页行-上一页跨页的行数
            '例如：记录单行数20，第一页跨页数据跨了5行，当跨页数据显示在当前页时，第二页的行数应该为15行。
            If mintPageSpan = 1 Then
                lngRowCount = Val(FormatValue(vsTest.TextMatrix(vsTest.Rows - 1, mlngRowCount)))
                lngRowCurrent = Val(FormatValue(vsTest.TextMatrix(vsTest.Rows - 1, mlngRowCurrent)))
                If lngRowCount > lngRowCurrent Then
                    mlngLitterRows(mint页码) = lngRowCount - lngRowCurrent
                    mlngCurLitterRows(mint页码) = mlngLitterRows(mint页码)
                    mArrPageInfo(mint页码) = mArrPageInfo(mint页码) & "[LPF]" & "由于勾选了参数:跨页数据只显示在当前页,当前页有" & lngRowCount - lngRowCurrent & "行数据显示在上一页。" & _
                        IIf(Val(vsTest.TextMatrix(vsTest.Rows - 1, mlngCollectType)) < 0, "数据小结名称:" & vsTest.TextMatrix(vsTest.Rows - 1, mlngCollectText), "数据发生时间:" & Mid(vsTest.TextMatrix(vsTest.Rows - 1, mlngActTime), 1, 16))
                End If
            End If
            If VsfData.Rows > VsfData.FixedRows Then
                If Val(VsfData.TextMatrix(VsfData.FixedRows, mlngRowCount)) > 0 Then
                    '50503:刘鹏飞,2012-09-12,分组数据显示在分组起始页，此处单独处理避免过滤掉了调跨页的数据
                    lngTestRow = vsTest.Rows - 1
                    lngRecordId = Val(VsfData.TextMatrix(VsfData.FixedRows, mlngRecord))
                    Do While True
                        If lngTestRow < vsTest.FixedRows Then GoTo ErrEnd
                        If lngRecordId = Val(vsTest.TextMatrix(lngTestRow, mlngRecord)) Then
                            lngTestRow = lngTestRow - 1
                        Else
                            Exit Do
                        End If
                    Loop
                    lngCount = 0
                    If Mid(FormatValue(VsfData.TextMatrix(VsfData.FixedRows, mlngActTime)), 1, 16) = Mid(vsTest.TextMatrix(lngTestRow, mlngActTime), 1, 16) And _
                        FormatValue(VsfData.TextMatrix(VsfData.FixedRows, mlngActTime)) <> vsTest.TextMatrix(lngTestRow, mlngActTime) And _
                        Not (Val(FormatValue(VsfData.TextMatrix(VsfData.FixedRows, mlngCollectType))) < 0 Or Val(vsTest.TextMatrix(lngTestRow, mlngCollectType)) < 0) Then
                        strDate = Mid(FormatValue(VsfData.TextMatrix(VsfData.FixedRows, mlngActTime)), 1, 16)
                        VsfData.Rows = VsfData.Rows + 1
                        lngRowCount = Val(FormatValue(VsfData.TextMatrix(VsfData.FixedRows, mlngRowCount)))
                        lngRowCurrent = Val(FormatValue(VsfData.TextMatrix(VsfData.FixedRows, mlngRowCurrent)))
                        For intData = 1 To lngRowCount
                            If VsfData.Rows > VsfData.FixedRows Then VsfData.RemoveItem VsfData.FixedRows
                        Next intData
                        lngCount = lngRowCount
                        mlngLitterRows(mint页码) = Val(CStr(mlngLitterRows(mint页码))) + lngRowCount
                        mlngCurLitterRows(mint页码) = Val(CStr(mlngCurLitterRows(mint页码))) + lngRowCurrent
                        Do While True
                            If Val(FormatValue(VsfData.TextMatrix(VsfData.FixedRows, mlngDemo))) <= 1 Then Exit Do
                            lngRowCount = Val((VsfData.TextMatrix(VsfData.FixedRows, mlngRowCount)))
                            lngRowCurrent = Val((VsfData.TextMatrix(VsfData.FixedRows, mlngRowCurrent)))
                            For intData = 1 To lngRowCount
                                If VsfData.Rows > VsfData.FixedRows Then VsfData.RemoveItem VsfData.FixedRows
                            Next intData
                            lngCount = lngCount + lngRowCount
                            mlngLitterRows(mint页码) = Val(CStr(mlngLitterRows(mint页码))) + lngRowCount
                            mlngCurLitterRows(mint页码) = Val(CStr(mlngCurLitterRows(mint页码))) + lngRowCurrent
                        Loop
                        If VsfData.Rows - 1 > VsfData.FixedRows Then
                            VsfData.RemoveItem VsfData.Rows - 1
                        End If
                    End If
                    If lngCount > 0 Then
                        mArrPageInfo(mint页码) = mArrPageInfo(mint页码) & "[LPF]" & "当前页有" & lngCount & "行分组数据显示在上一页。分组数据发生时间:" & strDate
                    End If
                End If
            End If
        End If
ErrEnd:
        Set vsTest.DataSource = Nothing
        vsTest.Clear
        rsData.Close
        Set rsData = Nothing
    End If

    If Val(CStr(mlngLitterRows(mint页码))) - lngLiterrRow <= 0 Then
        mlngLitterRows(mint页码) = 0
    Else
        mlngLitterRows(mint页码) = Val(CStr(mlngLitterRows(mint页码))) - lngLiterrRow
    End If
    
    If Val(CStr(mlngCurLitterRows(mint页码))) - lngLiterrRow <= 0 Then
        mlngCurLitterRows(mint页码) = 0
    Else
        mlngCurLitterRows(mint页码) = Val(CStr(mlngCurLitterRows(mint页码))) - lngLiterrRow
    End If
    
    '63760:刘鹏飞,分组数据护士、签名人、签名时间的处理（同一个签名人始终显示一次）
    If mlngSingerType > 0 And VsfData.FixedRows <= VsfData.Rows - 1 Then
        lngPrintedRow = IIf(mlngOperator <> -1, mlngOperator, mlngSignName)
        lngRow = VsfData.FixedRows
        Do While True
            lngStart = GetStartRow(lngRow)
            lngRowCount = Val(VsfData.TextMatrix(lngStart, mlngRowCount))
            If lngRowCount <= 0 Then Exit Do
            
            If mlngSingerType = 3 Then '尾行签名
                strSignName = VsfData.TextMatrix(lngStart + lngRowCount - 1, lngPrintedRow)
            Else '首行签名或首尾签名
                strSignName = VsfData.TextMatrix(lngStart, lngPrintedRow)
            End If
            strSignName = FormatValue(strSignName)
            If Val(VsfData.TextMatrix(lngStart, mlngDemo)) = 1 And lngStart = lngRow And strSignName <> "" Then
                For lngRow = lngStart + lngRowCount To VsfData.Rows - 1
                    If lngRow = lngStart + lngRowCount Then
                    
                        If Val(VsfData.TextMatrix(lngRow, mlngDemo)) <= 1 Then Exit For
                        
                        lngRowCount = Val(VsfData.TextMatrix(lngRow, mlngRowCount))
                        If lngRowCount = 0 Then Exit For
                        
                        If mlngSingerType = 3 Then '尾行签名
                            If strSignName = FormatValue(VsfData.TextMatrix(lngRow + lngRowCount - 1, lngPrintedRow)) Then
                                If lngStart <= lngRow - 1 Then
                                    If mlngOperator <> -1 Then VsfData.TextMatrix(lngRow - 1, mlngOperator) = ""
                                    If mlngSignName <> -1 Then VsfData.TextMatrix(lngRow - 1, mlngSignName) = ""
                                    If mlngSignTime <> -1 Then VsfData.TextMatrix(lngRow - 1, mlngSignTime) = ""
                                End If
                            Else
                                If FormatValue(VsfData.TextMatrix(lngRow + lngRowCount - 1, lngPrintedRow)) <> "" Then
                                    strSignName = FormatValue(VsfData.TextMatrix(lngRow + lngRowCount - 1, lngPrintedRow))
                                End If
                            End If
                        Else '首行签名或首尾签名
                            If strSignName = FormatValue(VsfData.TextMatrix(lngRow, lngPrintedRow)) Then
                                '首行签名或首尾签名都需要去掉下一条数据的首行,但首尾签名需要注意分组中的最后一条数据行数=1的情况
                                blnClear = True
                                If mlngSingerType = 2 And lngRowCount = 1 Then
                                    If lngRow + lngRowCount < VsfData.Rows Then
                                        If Val(VsfData.TextMatrix(lngRow + lngRowCount, mlngDemo)) <= 1 Then
                                            blnClear = False
                                        End If
                                    Else
                                        blnClear = False
                                    End If
                                End If
                                
                                If blnClear Then
                                    If mlngOperator <> -1 Then VsfData.TextMatrix(lngRow, mlngOperator) = ""
                                    If mlngSignName <> -1 Then VsfData.TextMatrix(lngRow, mlngSignName) = ""
                                    If mlngSignTime <> -1 Then VsfData.TextMatrix(lngRow, mlngSignTime) = ""
                                End If
                                
                                If mlngSingerType = 2 And lngStart < lngRow - 1 Then '首尾签名还应该去掉上一条数据的尾行(上一行数据行数需要>1)
                                    If mlngOperator <> -1 Then VsfData.TextMatrix(lngRow - 1, mlngOperator) = ""
                                    If mlngSignName <> -1 Then VsfData.TextMatrix(lngRow - 1, mlngSignName) = ""
                                    If mlngSignTime <> -1 Then VsfData.TextMatrix(lngRow - 1, mlngSignTime) = ""
                                End If
                            Else
                                If FormatValue(VsfData.TextMatrix(lngRow, lngPrintedRow)) <> "" Then
                                    strSignName = FormatValue(VsfData.TextMatrix(lngRow, lngPrintedRow))
                                End If
                            End If
                        End If
                        
                        lngStart = lngRow
                    End If
                Next lngRow
            Else
                lngRow = lngStart + Val(VsfData.TextMatrix(lngStart, mlngRowCount))
            End If
            
            If lngRow > VsfData.Rows - 1 Then Exit Do
        Loop
    End If
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub PreTendFormat(ByVal rsTemp As ADODB.Recordset)
    Dim aryItem() As String, strInfo As String
    Dim lngRow As Long, lngCol As Long, lngCount As Long, strCell As String
    Dim lngStartRow As Long
    Dim blnAlign As Boolean
    
    On Error GoTo ErrHand
    
    '设置表头的格式
    With vsfHead
        .Redraw = flexRDNone
        .Clear
        Set .DataSource = rsTemp
        .Rows = 3
        
        '表头填写
        .MergeCells = flexMergeRestrictRows
        .MergeCellsFixed = flexMergeFree
        .MergeRow(0) = True
        .MergeRow(1) = True
        .MergeRow(2) = True
        
        '程序内部控制列隐藏
        .ColHidden(mlngDemo) = True
        .ColHidden(mlngActTime) = True
        .ColHidden(mlngChoose) = Not mblnVerify
        '69355:刘鹏飞,2014-01-07,日期存在对角线(短日期格式7/1),则显示年度列
        .ColHidden(mlngYear) = Not mblnDateAd
        .ColHidden(mlngRowCurrent) = True
        .ColHidden(mlngRowCount) = True
        .ColHidden(mlngStartRowPage) = True
        .ColHidden(mlngStartRowNo) = True
        '51589:刘鹏飞,2013-03-01,添加交班签名
        .ColHidden(mlngJoinSignName) = True
        .ColHidden(mlngRecord) = True
        .ColHidden(mlngSigner) = True
        .ColHidden(mlngSignLevel) = True
        .ColHidden(mlngCollectStyle) = True
        .ColHidden(mlngCollectText) = True
        .ColHidden(mlngCollectType) = True
        .ColHidden(mlngCollectDay) = True
        .ColHidden(mlngCollectStart) = True
        .ColHidden(mlngCollectEnd) = True
        '63706:刘鹏飞,2013-11-20
        .ColHidden(.Cols - 1) = True
        .ColWidth(0) = 250
        .ColWidth(mlngChoose) = 250      '选择列
        .ColWidth(mlngYear) = BlowUp(picMain.TextWidth("刘鹏飞"))
        
        .FrozenCols = mlngTime
        .SheetBorder = &H40C0&
        
        '设置列头
        aryItem = Split(mstrTabHead, "|")
        For lngCount = 0 To UBound(aryItem)
            strCell = aryItem(lngCount)
            lngRow = Left(strCell, InStr(1, strCell, ",") - 1): strCell = Mid(strCell, InStr(1, strCell, ",") + 1)
            lngCol = Left(strCell, InStr(1, strCell, ",") - 1): strCell = Mid(strCell, InStr(1, strCell, ",") + 1)
            .TextMatrix(lngRow, lngCol + cHideCols + .FixedCols - 1) = strCell
        Next
        
        
        '设置固定列及选择列
        .TextMatrix(0, 0) = " "
        .TextMatrix(1, 0) = " "
        .TextMatrix(2, 0) = " "
        .TextMatrix(0, mlngChoose) = " "
        .TextMatrix(1, mlngChoose) = " "
        .TextMatrix(2, mlngChoose) = " "
        .TextMatrix(0, mlngYear) = "年份"
        .TextMatrix(1, mlngYear) = "年份"
        .TextMatrix(2, mlngYear) = "年份"
        Call PreActiveHead(vsfHead)
        
        '列宽设置
        blnAlign = False
        aryItem = Split(mstrColWidth, ",")
        For lngCount = cHideCols + .FixedCols To .Cols - 1
            If Not .ColHidden(lngCount) Then
                .ColWidth(lngCount) = BlowUp(Val(Split(aryItem(lngCount - cHideCols - .FixedCols), "`")(0)))
                If InStr(1, aryItem(lngCount - cHideCols - .FixedCols), "`") <> 0 Then
                    blnAlign = True
                    .ColAlignment(lngCount) = Val(Split(aryItem(lngCount - cHideCols - .FixedCols), "`")(1))
                End If
            End If
        Next
        
        '固定行格式为居中
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
        '再按列合并
        For lngCount = 0 To .Cols - 1
            .MergeCol(lngCount) = True
        Next
        
        If blnAlign = False Then
            '改为根据用户的设置显示列对齐方式
            If .FixedRows < .Rows Then .Cell(flexcpAlignment, .FixedRows, 0, .Rows - 1, .Cols - 1) = flexAlignGeneralCenter
        End If
        For lngCount = 0 To .Rows - 1
            If .RowHeight(lngCount) <> .RowHeightMin Then .RowHeight(lngCount) = .RowHeightMin
        Next
        Select Case mintTabTiers
        Case 1
            .RowHidden(0) = False
            .RowHidden(1) = True
            .RowHidden(2) = True
        Case 2
            .RowHidden(0) = False
            .RowHidden(1) = False
            .RowHidden(2) = True
        Case 3
            .RowHidden(0) = False
            .RowHidden(1) = False
            .RowHidden(2) = False
        End Select
        For lngCount = 0 To .Cols - 1
            .MergeCol(lngCount) = True
        Next
        
        '可能固定行的行高不正确需要自动调整下
        .AutoResize = True
        .WordWrap = True
        .AutoSizeMode = flexAutoSizeRowHeight
        .AutoSize 0, .Cols - 1
        .AutoResize = False
        '将非固定行的行高设置为最小行高
        For lngCount = .FixedRows To .Rows - 1
            .RowHeight(lngCount) = .RowHeightMin
        Next
        
        .Height = 0
        For lngCount = 0 To .FixedRows - 1
            If Not .RowHidden(lngCount) Then
                If .RowHeight(lngCount) < .RowHeightMin Then
                    .Height = .Height + .RowHeightMin
                Else
                    .Height = .Height + .RowHeight(lngCount)
                End If
            End If
        Next
        .Height = .Height - 20
        .Redraw = flexRDDirect
    End With
    
    '设置护理记录单的格式
    With VsfData
        .Redraw = flexRDNone
        .Clear
        Call DataMap_Restore(rsTemp)
        
        '表头填写
        .MergeCells = flexMergeRestrictRows
        .MergeCellsFixed = flexMergeFree
        .MergeRow(0) = True
        .MergeRow(1) = True
        .MergeRow(2) = True
        
        '程序内部控制列隐藏
        .ColHidden(mlngDemo) = True
        .ColHidden(mlngActTime) = True
        .ColHidden(mlngChoose) = Not mblnVerify
        .ColHidden(mlngYear) = Not mblnDateAd
        .ColHidden(mlngRowCurrent) = True
        .ColHidden(mlngRowCount) = True
        '51589:刘鹏飞,2013-03-01,添加交班签名
        .ColHidden(mlngJoinSignName) = True
        .ColHidden(mlngRecord) = True
        .ColHidden(mlngStartRowPage) = True
        .ColHidden(mlngStartRowNo) = True
        .ColHidden(mlngSigner) = True
        .ColHidden(mlngSignLevel) = True
        .ColHidden(mlngCollectStyle) = True
        .ColHidden(mlngCollectText) = True
        .ColHidden(mlngCollectType) = True
        .ColHidden(mlngCollectDay) = True
        .ColHidden(mlngCollectStart) = True
        .ColHidden(mlngCollectEnd) = True
        '63706:刘鹏飞,2013-11-20
        .ColHidden(.Cols - 1) = True
        .ColWidth(0) = 250
        .ColWidth(mlngChoose) = 250      '选择列
        .ColWidth(mlngYear) = BlowUp(picMain.TextWidth("刘鹏飞"))
        
        .FrozenCols = mlngTime
        .SheetBorder = &H40C0&
        
        '设置列头
        aryItem = Split(mstrTabHead, "|")
        For lngCount = 0 To UBound(aryItem)
            strCell = aryItem(lngCount)
            lngRow = Left(strCell, InStr(1, strCell, ",") - 1): strCell = Mid(strCell, InStr(1, strCell, ",") + 1)
            lngCol = Left(strCell, InStr(1, strCell, ",") - 1): strCell = Mid(strCell, InStr(1, strCell, ",") + 1)
            .TextMatrix(lngRow, lngCol + cHideCols + .FixedCols - 1) = strCell
        Next
        
        '设置固定列及选择列
        .TextMatrix(0, 0) = " "
        .TextMatrix(1, 0) = " "
        .TextMatrix(2, 0) = " "
        .TextMatrix(0, mlngChoose) = " "
        .TextMatrix(1, mlngChoose) = " "
        .TextMatrix(2, mlngChoose) = " "
        .TextMatrix(0, mlngYear) = "年份"
        .TextMatrix(1, mlngYear) = "年份"
        .TextMatrix(2, mlngYear) = "年份"
        
        Call PreActiveHead(VsfData)
        
        '列宽设置
        blnAlign = False
        aryItem = Split(mstrColWidth, ",")
        For lngCount = cHideCols + .FixedCols To .Cols - 1
            If Not .ColHidden(lngCount) Then
                .ColWidth(lngCount) = BlowUp(Val(Split(aryItem(lngCount - cHideCols - .FixedCols), "`")(0)))
                If InStr(1, aryItem(lngCount - cHideCols - .FixedCols), "`") <> 0 Then
                    blnAlign = True
                    .ColAlignment(lngCount) = Val(Split(aryItem(lngCount - cHideCols - .FixedCols), "`")(1))
                End If
            End If
        Next
        
        '固定行格式为居中
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
        '再按列合并
        For lngCount = 0 To .Cols - 1
            .MergeCol(lngCount) = True
        Next
        
        If blnAlign = False Then
            '改为根据用户的设置显示列对齐方式
            If .FixedRows < .Rows Then .Cell(flexcpAlignment, .FixedRows, 0, .Rows - 1, .Cols - 1) = flexAlignGeneralCenter
        End If
        For lngCount = 0 To .Rows - 1
            If .RowHeight(lngCount) <> .RowHeightMin Then .RowHeight(lngCount) = .RowHeightMin
        Next
        Select Case mintTabTiers
        Case 1
            .RowHidden(0) = False
            .RowHidden(1) = True
            .RowHidden(2) = True
        Case 2
            .RowHidden(0) = False
            .RowHidden(1) = False
            .RowHidden(2) = True
        Case 3
            .RowHidden(0) = False
            .RowHidden(1) = False
            .RowHidden(2) = False
        End Select
        For lngCount = 0 To .Cols - 1
            .MergeCol(lngCount) = True
        Next
        
        strInfo = ""
        lblInfo.Tag = ""
        If .Rows = .FixedRows Then
            mlngOverrunRows = 0
            mlngReduceRow = 0
        Else
            '得到第一行减少的行(主要是针对合并打印的文件，可能上一份文件存在半页数据的情况)
            '例如：文件1结束行号为9，本文件和文件1合并，那么本文件的第一页的开始行号=10.此时记录单显示的行数需要减去9行
            If mint页码 = mint起始页码 And Val(VsfData.TextMatrix(3, mlngStartRowNo)) > 1 Then
                mlngReduceRow = Val(VsfData.TextMatrix(3, mlngStartRowNo)) - 1
                If picImg.Tag <> "" Then
                    strInfo = strInfo & "[LPF]" & "由于当前文件与文件'" & picImg.Tag & "'设置了合并打印,而被合并的文件最后一页数据未满页。因此当前页的数据有" & mlngReduceRow & "行显示在被合并的文件。"
                End If
            Else
                mlngReduceRow = 0
            End If
            '得到第一行的超出行
            mlngOverrunRows = Val(.TextMatrix(3, mlngRowCount)) - Val(.TextMatrix(3, mlngRowCurrent))
            If mlngOverrunRows > 0 Then
                If Val(.TextMatrix(3, mlngStartRowPage)) = mint页码 Then
                    strInfo = strInfo & "[LPF]" & IIf(Val(.TextMatrix(3, mlngCollectType)) < 0, "当前页小结:" & .TextMatrix(3, mlngCollectText), "当前页发生时间:" & Mid(.TextMatrix(3, mlngActTime), 1, 16)) & _
                        "的数据从第" & Val(.TextMatrix(3, mlngRowCurrent)) + 1 & "行开始跨页,跨页行数" & mlngOverrunRows & "行。"
                Else
                    strInfo = strInfo & "[LPF]" & IIf(Val(.TextMatrix(3, mlngCollectType)) < 0, "当前页小结:" & .TextMatrix(3, mlngCollectText), "当前页发生时间:" & Mid(.TextMatrix(3, mlngActTime), 1, 16)) & _
                        "的数据前" & Val(.TextMatrix(3, mlngRowCurrent)) & "行为上一页的数据。"
                End If
            End If
            '50503:刘鹏飞,2012-09-12,数据从某一页第一行就开始跨页，计算添加行不能重复计算，本次修改：
            '情况一:
            '如果:第一行和最后一行的记录起始行相同，那么说明是同一条数据，超出行累计不加上最后一行的超出行
            '情况二:81982,刘鹏飞,2015-01-30，分类汇总添加行的处理(分类汇总初始加载有多行记录)
            '否则:对于分类汇总数据第一次初始化时，数据可能存在多行(但实际多条数据是一条数据)，为了避免重复添加行，则需要根据记录ID判断是否是同一条数据。
            '原因:1.普通数据，不管数据是否展开始终可以从情况一判断的出来。2。任何数据只要已展开也可从情况以判断的出来。3.只有分类汇总数据展开前无法从情况一判断。
            lngStartRow = 3
            '加上最后一行的超出行
            If .Rows - 1 <> 3 Then
                If Val(.TextMatrix(.Rows - 1, mlngRowCount)) - Val(.TextMatrix(.Rows - 1, mlngRowCurrent)) > 0 Then
                    If lngStartRow <> GetStartRow(.Rows - 1) Then
                        If Val(.TextMatrix(lngStartRow, mlngRecord)) <> Val(.TextMatrix(.Rows - 1, mlngRecord)) And Val(.TextMatrix(lngStartRow, mlngRecord)) <> 0 And Val(.TextMatrix(.Rows - 1, mlngRecord)) <> 0 Then
                            mlngOverrunRows = mlngOverrunRows + Val(.TextMatrix(.Rows - 1, mlngRowCount)) - Val(.TextMatrix(.Rows - 1, mlngRowCurrent))
                            strInfo = strInfo & "[LPF]" & IIf(Val(.TextMatrix(.Rows - 1, mlngCollectType)) < 0, "当前页小结:" & .TextMatrix(.Rows - 1, mlngCollectText), "当前页发生时间:" & Mid(.TextMatrix(.Rows - 1, mlngActTime), 1, 16)) & _
                                "的数据从第" & Val(.TextMatrix(.Rows - 1, mlngRowCurrent)) + 1 & "行开始跨页,跨页行数" & Val(.TextMatrix(.Rows - 1, mlngRowCount)) - Val(.TextMatrix(.Rows - 1, mlngRowCurrent)) & "行。"
                        End If
                    End If
                End If
               ' mlngOverrunRows = mlngOverrunRows + Val(.TextMatrix(.Rows - 1, mlngRowCount)) - Val(.TextMatrix(.Rows - 1, mlngRowCurrent))
            End If
        End If
        Call PreTendMutilRows
        lblInfo.Tag = Trim(Mid(strInfo & mArrPageInfo(mint页码), 6))
        picInfo.Visible = lblInfo.Tag <> ""
        Call FillPage
        
        Call WriteColor
        
        '可能固定行的行高不正确需要自动调整下
        .AutoResize = True
        .WordWrap = True
        .AutoSizeMode = flexAutoSizeRowHeight
        .AutoSize 0, .Cols - 1
        .AutoResize = False
        
        .RowHidden(0) = True
        .RowHidden(1) = True
        .RowHidden(2) = True
        
        '将非固定行的行高设置为最小行高
        For lngCount = .FixedRows To .Rows - 1
            .RowHeight(lngCount) = .RowHeightMin
        Next
        
        .WordWrap = False           '不能自动换行,否则会出现最后一个字看不见的情况
        .Redraw = flexRDDirect
    End With
    
    With chkSwitch
        .Value = 0
        .Top = vsfHead.Top + vsfHead.Height - .Height - 50
        .Left = vsfHead.Left + vsfHead.Cell(flexcpLeft, mintTabTiers - 1, mlngChoose) + 50
        .Visible = mblnVerify
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub WriteColor()
    Dim blnTag As Boolean
    Dim lngCount As Long, lngCol As Long
    '晚班以红色显示，同时将非起始行设置为NoCheckBox，设置图标
    With VsfData
        For lngCount = .FixedRows To .Rows - 1
            If .TextMatrix(lngCount, 2) <> "" And Val(.TextMatrix(lngCount, mlngCollectType)) = 0 Then
                '晚班以红色显示
                blnTag = False
                If mintTagFormHour < mintTagToHour Then
                    blnTag = (Hour(.TextMatrix(lngCount, 2)) >= mintTagFormHour And Hour(.TextMatrix(lngCount, 2)) < mintTagToHour)
                Else
                    blnTag = (Hour(.TextMatrix(lngCount, 2)) >= mintTagFormHour Or Hour(.TextMatrix(lngCount, 2)) < mintTagToHour)
                End If
                If blnTag Then
                    Set .Cell(flexcpFont, lngCount, 0, lngCount, .Cols - 1) = mobjTagFont
                    .Cell(flexcpForeColor, lngCount, 0, lngCount, .Cols - 1) = mlngTagColor
                End If
            End If
            
            '处理汇总行,如果为零显示为空
            If Val(.TextMatrix(lngCount, mlngCollectType)) < 0 And FormatValue(VsfData.TextMatrix(lngCount, mlngRowCount)) Like "*|1" Then
                '88967:护士、签名列同时存在，且属于同一操作员，则应避免合并(打印签名人输出签名图片需注意签名人有回车的情况)
                For lngCol = mlngTime + 1 To IIf(mlngNoEditor < mlngSignName, mlngSignName, mlngNoEditor)
                    '52953,刘鹏飞,2012-08-24,汇总数据为0也要显示,关联问题60792
                    'If .TextMatrix(lngCount, lngCOL) = "0" Then .TextMatrix(lngCount, lngCOL) = ""
                    .TextMatrix(lngCount, lngCol) = FormatValue(.TextMatrix(lngCount, lngCol))
                    If Trim(.TextMatrix(lngCount, lngCol)) <> "" And .ColHidden(lngCol) = False Then
                        '66085:刘鹏飞,2012-09-26,避免相邻汇总列合并,将原来的列内容+空格同一改成在列后面在chr(13)
                        '避免因加空格后列宽不够导致内容显示不完全(主要针对右对其)
'                        Select Case .ColAlignment(lngCol)
'                            Case 6, 7, 8
'                                .TextMatrix(lngCount, lngCol) = IIf(lngCol Mod 2 = 1, " ", String(2, " ")) & .TextMatrix(lngCount, lngCol)
'                            Case 3, 4, 5
'                                .TextMatrix(lngCount, lngCol) = IIf(lngCol Mod 2 = 1, " ", String(2, " ")) & .TextMatrix(lngCount, lngCol) & IIf(lngCol Mod 2 = 1, " ", String(2, " "))
'                            Case 0, 1, 2
'                                .TextMatrix(lngCount, lngCol) = .TextMatrix(lngCount, lngCol) & IIf(lngCol Mod 2 = 1, " ", String(2, " "))
'                            Case Else
'                                .TextMatrix(lngCount, lngCol) = IIf(lngCol Mod 2 = 1, " ", String(2, " ")) & .TextMatrix(lngCount, lngCol)
'                        End Select
                        .TextMatrix(lngCount, lngCol) = .TextMatrix(lngCount, lngCol) & IIf(lngCol Mod 2 = 1, Chr(13), "")
                    End If
                    '.TextMatrix(lngCount, lngCOL) = IIf(lngCOL Mod 2 = 1, " ", "") & .TextMatrix(lngCount, lngCOL)
                Next
                .MergeRow(lngCount) = True
            Else
                .MergeRow(lngCount) = False
            End If
            
            '将非起始行设置为NoCheckBox
            If Not VsfData.RowHidden(lngCount) Then
                If Not FormatValue(VsfData.TextMatrix(lngCount, mlngRowCount)) Like "*|1" Then
                    VsfData.Cell(flexcpChecked, lngCount, mlngChoose) = flexNoCheckbox
                Else
                    If VsfData.Cell(flexcpChecked, lngCount, mlngChoose) <> flexTSChecked Then
                        VsfData.Cell(flexcpChecked, lngCount, mlngChoose) = flexTSUnchecked
                    End If
                    
                    '设置图标
                    If FormatValue(VsfData.TextMatrix(lngCount, mlngSigner)) = "" Then
                        VsfData.Cell(flexcpPicture, lngCount, 0) = Nothing
                    Else
                        If InStr(1, VsfData.TextMatrix(lngCount, mlngSigner), "/") <> 0 Then
                            VsfData.Cell(flexcpPicture, lngCount, 0) = imgRow.ListImages(审签).Picture
                        '51589:刘鹏飞,2013-03-01,添加交班签名
                        ElseIf Trim(VsfData.TextMatrix(lngCount, mlngJoinSignName)) <> "" Then '交班签名
                            VsfData.Cell(flexcpPicture, lngCount, 0) = imgRow.ListImages(交班签名).Picture
                        Else
                            VsfData.Cell(flexcpPicture, lngCount, 0) = imgRow.ListImages(签名).Picture
                        End If
                    End If
                
                    '处理小结的显示
                    If Val(FormatValue(VsfData.TextMatrix(lngCount, mlngCollectType))) <> 0 Then
                        If mblnDateAd Then VsfData.TextMatrix(lngCount, mlngYear) = FormatValue(VsfData.TextMatrix(lngCount, mlngCollectText))
                        VsfData.TextMatrix(lngCount, mlngDate) = FormatValue(VsfData.TextMatrix(lngCount, mlngCollectText))
                        VsfData.TextMatrix(lngCount, mlngTime) = FormatValue(VsfData.TextMatrix(lngCount, mlngCollectText))
                    End If
                End If
            End If
        Next
    End With
End Sub

Private Sub zlLableBruit()
    Dim lngScaleLeft As Long, lngScaleTop As Long, lngScaleRight As Long, lngScaleBottom As Long
    
    On Error Resume Next
    lblSubhead.Caption = lblSubhead.Tag
    lblSubhead.Top = lblTitle.Top + lblTitle.Height + 120
    Call cbsThis.GetClientRect(lngScaleLeft, lngScaleTop, lngScaleRight, lngScaleBottom)
'    VsfData.Move lngScaleLeft + 210, lblSubhead.Top + lblSubhead.Height + 45, lngScaleRight - lngScaleLeft - 210 * 2
'    VsfData.Height = picMain.Height - VsfData.Top
    vsfHead.Move lngScaleLeft + 210, lblSubhead.Top + lblSubhead.Height + 45, lngScaleRight - lngScaleLeft - 210 * 2
    VsfData.Move vsfHead.Left, vsfHead.Top + vsfHead.Height - 20, vsfHead.Width
    VsfData.Height = picMain.Height - vsfHead.Height - vsfHead.Top
End Sub

Private Sub GetFileProperty()
    '提取文件属性
    Dim lngYear As Long
    Dim strEndTime As String, strCurDate As String
    On Error GoTo ErrHand
    
    '"关于记录单数据行数显示的说明："
    '"1、分组数据只显示在分组起始页,这就意味着分组起始页的数据行数增加,下一页的数据行数减少。"
    '"2、跨页数据根据参数'跨页数据只显示在当前页'决定是显示在当前页还是两页都显示。如果显示在当前页就意味着显示在当前页数据行数增加，下一页的数据行数减少;如果两页都显示就意味着显示在两页的数据行数都增加。"
    '"3、当前文件如果与上一份文件设置了合并打印，如果上份文件最后一页未满页，那么当前文件显示的数据行数就会减少。"
    strCurDate = zlDatabase.Currentdate
    
    gstrSQL = " Select   开始时间,结束时间,格式ID,科室ID,归档人 From 病人护理文件 " & _
              " Where 病人ID=[1] And 主页ID=[2] And 婴儿=[3] And ID=[4] And Rownum<2"
    Call SQLDIY(gstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取护理文件数据", mlng病人ID, mlng主页ID, mint婴儿, mlng文件ID)
    If rsTemp.RecordCount <> 0 Then
        mlng格式ID = rsTemp!格式ID
        mlng科室ID = rsTemp!科室ID
        mblnArchive = (NVL(rsTemp!归档人) <> "")
        mstr开始时间 = Format(rsTemp!开始时间, "yyyy-MM-dd HH:mm:ss")
        mstr结束时间 = Format(rsTemp!结束时间, "yyyy-MM-dd HH:mm:ss")
    End If
    '69935:刘鹏飞,2014-1-7
    If IsDate(mstr结束时间) Then
        strEndTime = mstr结束时间
    Else
        strEndTime = mstrMaxDate
    End If
    mstrYears = ""
    For lngYear = Val(Format(strEndTime, "YYYY")) To Val(Format(mstr开始时间, "YYYY")) Step -1
        If Val(Format(strCurDate, "YYYY")) = lngYear Then
            mstrYears = mstrYears & "|" & lngYear
        Else
            mstrYears = mstrYears & "|" & lngYear
        End If
    Next lngYear
    mstrYears = Mid(mstrYears, 2)
    
    '如果页码=-1,说明缺省显示最后一页
    mint起始页码 = 1
    gstrSQL = " Select  MIN(开始页号) 起始页码, MAX(结束页号) AS 页码 From 病人护理打印 Where 文件ID=[1]"
    Call SQLDIY(gstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取指定页码的数据发生时间范围", mlng文件ID)
    mint起始页码 = NVL(rsTemp!起始页码, 1)
    mint结束页 = NVL(rsTemp!页码, 1)
    If mint页码 = -1 Then mint页码 = mint结束页
    If mint页码 < mint起始页码 Then mint页码 = mint起始页码
    If mint页码 > mint结束页 + 1 Then mint页码 = mint结束页
    
    '提取合并文件名称用于提示
    picImg.Tag = ""
    gstrSQL = "Select 文件名称 From 病人护理文件 Where 病人ID=[1] And 主页ID=[2] And NVL(婴儿,0)=[3] And 续打ID=[4]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取和当前文件合并的文件", mlng病人ID, mlng主页ID, mint婴儿, mlng文件ID)
    If rsTemp.RecordCount > 0 Then
        picImg.Tag = NVL(rsTemp!文件名称)
    End If
    
    Call InitPages
    
    If mblnClear = True Then
        ReDim mArrPageInfo(0 To mint结束页 + 1)
        ReDim mlngLitterRows(0 To mint结束页 + 1)
        ReDim mlngCurLitterRows(0 To mint结束页 + 1)
    Else
        ReDim Preserve mArrPageInfo(0 To mint结束页 + 1)
        ReDim Preserve mlngLitterRows(0 To mint结束页 + 1)
        ReDim Preserve mlngCurLitterRows(0 To mint结束页 + 1)
    End If
    RaiseEvent AfterRowColChange("", False, mblnSign, mblnArchive)
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub InitEnv()
    Dim rs As New ADODB.Recordset
    On Error GoTo ErrHand
    
    glngHours = Val(zlDatabase.GetPara("数据补录时限", glngSys))

    '打开现存在的所有护理记录项目
    gstrSQL = " Select   项目序号,upper(项目名称) AS 项目名称,项目类型,项目性质,项目长度,项目小数,项目表示,项目单位,项目值域,护理等级,应用方式,说明" & _
              " From 护理记录项目 B" & _
              " Order by 项目序号"
    Set mrsItems = zlDatabase.OpenSQLRecord(gstrSQL, "打开现存在的所有护理记录项目")
    
    '提取适用于记录单的诊治所见项目
    gstrSQL = _
        " Select i.分类id, i.编码, i.中文名, nvl(i.替换域,0) 替换域,i.类型,i.长度,i.小数,i.单位,i.表示法,i.数值域,i.必填" & vbNewLine & _
        " From 诊治所见项目 i, 诊治所见分类 k" & vbNewLine & _
        " Where k.Id = i.分类id And ((k.编码 In ('02', '05', '06') And i.替换域 = 1) Or (k.性质 = 2 And k.编码 = '06' And NVL(i.替换域,0) = 0))" & vbNewLine & _
        " Order By k.性质, k.编码, i.编码"
    Set mrsElement = zlDatabase.OpenSQLRecord(gstrSQL, "提取适用于记录单的诊治所见项目")
    
    '取当前操作员的级别
    mintVerify = 未定义
    mintVerify_Last = 未定义
    gstrSQL = "select  聘任技术职务 from 人员表 p where ID=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mRichEPR", glngUserId)
    If Not rs.EOF Then
        mintVerify = NVL(rs("聘任技术职务"), 未定义)
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub InitRecords()
    Dim i As Integer, j As Integer, k As Integer, l As Integer
    Dim lngCol As Long, lngOrder As Long, strName As String, intImmovable As Integer, strFormat As String
    Dim arrColumn, arrItem, arrCorrelative(), strColumns As String
    Dim blnSet As Boolean
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    
    strColumns = mstrColumns
    If Not mblnInit Then
        '初始化内存记录集(未对应项目的列为活动项目,其它列均为固定项)
        strFields = "列," & adDouble & ",18|序号," & adDouble & ",2|项目序号," & adDouble & ",18|项目名称," & adLongVarChar & ",20|固定," & adDouble & ",2|格式," & adLongVarChar & ",2000"
        Call Record_Init(mrsSelItems, strFields)
        strFields = "列|序号|项目序号|项目名称|固定|格式"
    End If
    
    '加入列定义
    If Not mblnInit Then
        arrColumn = Split(strColumns, "|")
        j = UBound(arrColumn)
        For i = 0 To j
            lngCol = Split(arrColumn(i), "'")(0)
            arrItem = Split(Split(arrColumn(i), "'")(1), ",")
            blnSet = False   '如果已设置以传入值为准'否则找不到项目就是活动项目
            If UBound(Split(arrColumn(i), "'")) > 1 Then
                blnSet = True
                intImmovable = Split(arrColumn(i), "'")(2)
            End If
            If UBound(Split(arrColumn(i), "'")) > 2 Then
                strFormat = Split(arrColumn(i), "'")(3)
            End If
            
            k = UBound(arrItem)
            For l = 0 To k
                strName = arrItem(l)
                mrsItems.Filter = "项目名称='" & strName & "'"
                If mrsItems.RecordCount <> 0 Then
                    lngOrder = mrsItems!项目序号
                    If Not blnSet Then intImmovable = 1   '固定不允许修改
                Else
                    lngOrder = 0
                    If Not blnSet Then intImmovable = 0
                    
                    '记录特殊列
                    Select Case strName
                    Case "日期"
                        mlngDate = i + cHideCols + VsfData.FixedCols
                    Case "时间"
                        mlngTime = i + cHideCols + VsfData.FixedCols
                    Case "护士"
                        mlngOperator = i + cHideCols + VsfData.FixedCols
                    Case "签名人"
                        mlngSignName = i + cHideCols + VsfData.FixedCols
                    Case "签名时间"
                        mlngSignTime = i + cHideCols + VsfData.FixedCols
                    End Select
                End If
                strValues = lngCol & "|" & l + 1 & "|" & lngOrder & "|" & strName & "|" & intImmovable & "|" & strFormat
                Call Record_Add(mrsSelItems, strFields, strValues)
            Next
        Next
        
        '整理分类汇总关联列信息
        arrCorrelative = Array()
        arrColumn = Split(mstrColCorrelative, "|")
        For i = 0 To UBound(arrColumn)
            arrItem = Split(arrColumn(i), ";")
            If UBound(arrItem) = 1 Then
                mrsSelItems.Filter = "列=" & Val(arrItem(0))
                If mrsSelItems.RecordCount = 1 Then
                    ReDim Preserve arrCorrelative(UBound(arrCorrelative) + 1)
                    arrCorrelative(UBound(arrCorrelative)) = Val(arrItem(0)) & "," & mrsSelItems!项目序号 & ";" & CStr(arrItem(1))
                End If
            End If
        Next i
        If UBound(arrCorrelative) = -1 Then
            mstrColCorrelative = ""
        Else
            mstrColCorrelative = Join(arrCorrelative, "|")
        End If
        mstrColImCorrelative = mstrColCorrelative
        If mblnCorrelative = False Then mstrColCorrelative = ""
        mrsSelItems.Filter = ""
        'Call OutputRsData(mrsSelItems)
        
        '加入程序内部控制列(列是在读取数据后绑定时增加的,此时只有预处理下)
        mlngDemo = VsfData.FixedCols
        mlngActTime = mlngDemo + 1
        mlngChoose = mlngActTime + 1
        mlngYear = mlngChoose + 1
        mlngSignLevel = VsfData.Cols + cHideCols + VsfData.FixedCols '加上隐藏列
        mlngSigner = mlngSignLevel + 1
        '51589:刘鹏飞,2013-03-01,添加交班签名
        mlngJoinSignName = mlngSigner + 1
        mlngRecord = mlngJoinSignName + 1
        mlngRowCount = mlngRecord + 1
        mlngRowCurrent = mlngRowCount + 1
        mlngStartRowPage = mlngRowCurrent + 1
        mlngStartRowNo = mlngStartRowPage + 1
        mlngCollectType = mlngStartRowNo + 1
        mlngCollectText = mlngCollectType + 1
        mlngCollectStyle = mlngCollectText + 1
        mlngCollectDay = mlngCollectStyle + 1
        mlngCollectStart = mlngCollectDay + 1
        mlngCollectEnd = mlngCollectStart + 1
        
        If mlngOperator <> -1 And mlngSignName <> -1 Then
            mlngNoEditor = IIf(mlngOperator < mlngSignName, mlngOperator, mlngSignName)
        Else
            mlngNoEditor = IIf(mlngOperator <> -1, mlngOperator, mlngSignName)
        End If
    End If
    
    mrsItems.Filter = 0
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub SetArchiveValue(ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal intBaby As Integer)
    mlng病人ID = lngPatiID
    mlng主页ID = lngPageId
    mint婴儿 = intBaby
End Sub

Public Sub ArchiveMe()
    On Error GoTo ErrHand
    
    If mlng病人ID = 0 Or gblnMoved Then Exit Sub
    If MsgBox("需要将该病人本次住院所有护理文件归档吗？", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbYes Then
        Dim strNow As String

        strNow = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
        gstrSQL = "ZL_病人护理文件_ARCHIVE(" & mlng病人ID & "," & mlng主页ID & "," & mint婴儿 & ",1)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "归档")

        mblnArchive = True
        RaiseEvent AfterRowColChange("", False, mblnSign, mblnArchive)
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub UnArchiveMe()
    On Error GoTo ErrHand
    
    If mlng病人ID = 0 Or gblnMoved Then Exit Sub
    If MsgBox("需要取消该病人的归档状态吗？", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbYes Then

        gstrSQL = "ZL_病人护理文件_ARCHIVE(" & mlng病人ID & "," & mlng主页ID & "," & mint婴儿 & ",0)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "撤销归档")
        
        mblnArchive = False
        RaiseEvent AfterRowColChange("", False, mblnSign, mblnArchive)
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function SignMe(Optional ByVal bln审签 As Boolean = False, Optional ByVal blnExchange As Boolean = False) As Boolean
    Dim blnSign As Boolean          '是否签名成功
    Dim blnRefresh As Boolean
    Dim strSignTime As String       '保证所有签名的签名时间一致,便于取消签名时按签名时间统一取消
    Dim str状态 As String           '保存签名选项,避免循环签名时不停的弹出签名窗口
    Dim str行错误 As String
    Dim str错误 As String
    Dim intRow As Integer, intRows As Integer
    Dim rsTemp As New ADODB.Recordset
    Dim strInfo As String
    
    Dim lngStart As Long, lngDemo As Long, lngRow As Long
    On Error GoTo ErrHand
    '按发生时间循环对所有未签名数据进行签名
    
    If mlng病人ID = 0 Then Exit Function
    
    '普签:对所有未签名的数据进行签名
    '审签:对所有已签名的数据进行审签
    If bln审签 Then
        blnExchange = False
        '43588:刘鹏飞,2012-09-13,添加记录单审签模式
        '0-缺省，聘任职务+审签权限，按审签时的职务高低进行控制，最多可达四级审签；1-审签权限，只有具有审签权限的人可以审签他人记录，审签后不能再次审签）
        If Not mblnVerify Then
            '汇总数据也要签名,因此去掉条件: And B.汇总类别=0
            If mintSignMode = 1 Then
                gstrSQL = " Select  distinct B.ID,B.发生时间 " & vbNewLine & _
                          " From 病人护理明细 A,病人护理数据 B,病人护理文件 C" & vbNewLine & _
                          " Where A.记录ID=B.ID And B.文件ID=C.ID And A.数据来源 in (0,3) And A.记录类型=5 AND A.终止版本 Is NULL And C.ID=[1] " & _
                          " MINUS" & vbNewLine & _
                          " Select  distinct B.ID,B.发生时间 " & vbNewLine & _
                          " From 病人护理明细 A,病人护理数据 B,病人护理文件 C" & vbNewLine & _
                          " Where A.记录ID=B.ID And B.文件ID=C.ID And A.数据来源 in (0,3)  And A.记录类型=15  AND A.终止版本 Is NULL And C.ID=[1] "
            Else
                gstrSQL = " Select  distinct B.ID,B.发生时间 " & vbNewLine & _
                          " From 病人护理明细 A,病人护理数据 B,病人护理文件 C" & vbNewLine & _
                          " Where A.记录ID=B.ID And B.文件ID=C.ID And A.数据来源 in (0,3)  And MOD(A.记录类型,10)=5  AND A.终止版本 Is NULL And C.ID=[1] "
            End If
            Call SQLDIY(gstrSQL)
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查是否存在未签名的数据", mlng文件ID)
            If rsTemp.RecordCount = 0 Then
                RaiseEvent AfterRowColChange("不存在已签名的数据！", True, mblnSign, mblnArchive)
                Exit Function
            End If
        
            '进入审签模式,可修改数据,可勾选数据
            mblnVerify = True
            chkSwitch.Visible = mblnVerify
            chkSwitch.ZOrder
            vsfHead.ColHidden(mlngChoose) = Not mblnVerify
            VsfData.ColHidden(mlngChoose) = Not mblnVerify
            VsfData.Cell(flexcpChecked, VsfData.FixedRows, mlngChoose, VsfData.Rows - 1, mlngChoose) = flexTSUnchecked
            Call WriteColor
            RaiseEvent AfterDataChanged(mblnChange Or mblnVerify)
            Exit Function
        Else
            '提取待审签的数据
            '汇总数据也要签名,因此去掉条件: And B.汇总类别=0
            If mintSignMode = 1 Then
                gstrSQL = " Select /*+ RULE */ distinct B.ID,B.发生时间 " & vbNewLine & _
                          " From 病人护理明细 A,病人护理数据 B,病人护理文件 C,(SELECT COLUMN_VALUE FROM TABLE(CAST(F_NUM2LIST([2]) AS ZLTOOLS.T_NUMLIST))) G" & vbNewLine & _
                          " Where A.记录ID=B.ID And B.ID=G.COLUMN_VALUE And B.文件ID=C.ID And A.记录类型=5  AND A.终止版本 Is NULL And C.ID=[1] " & _
                          " MINUS" & vbNewLine & _
                          " Select /*+ RULE */ distinct B.ID,B.发生时间 " & vbNewLine & _
                          " From 病人护理明细 A,病人护理数据 B,病人护理文件 C,(SELECT COLUMN_VALUE FROM TABLE(CAST(F_NUM2LIST([2]) AS ZLTOOLS.T_NUMLIST))) G" & vbNewLine & _
                          " Where A.记录ID=B.ID And B.ID=G.COLUMN_VALUE And B.文件ID=C.ID And A.记录类型=15  AND A.终止版本 Is NULL And C.ID=[1] "
            Else
                gstrSQL = " Select /*+ RULE */ distinct B.ID,B.发生时间 " & vbNewLine & _
                          " From 病人护理明细 A,病人护理数据 B,病人护理文件 C,(SELECT COLUMN_VALUE FROM TABLE(CAST(F_NUM2LIST([2]) AS ZLTOOLS.T_NUMLIST))) G" & vbNewLine & _
                          " Where A.记录ID=B.ID And B.ID=G.COLUMN_VALUE And B.文件ID=C.ID And MOD(A.记录类型,10)=5  AND A.终止版本 Is NULL And C.ID=[1] "
            End If
            Call SQLDIY(gstrSQL)
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查是否存在未签名的数据", mlng文件ID, mstrVerify)
        End If
    Else
        '仅对本人修改的数据进行签名(提取未签名数据-已签名数据)
        '汇总数据也要签名,因此去掉条件: And B.汇总类别=0
        mintVerify_Last = 未定义
        '51589:刘鹏飞,2013-03-01,添加交班签名
        If blnExchange = False Then
            gstrSQL = "" & _
                    "SELECT  DISTINCT B.ID,B.发生时间" & vbNewLine & _
                    "FROM 病人护理明细 A,病人护理数据 B" & vbNewLine & _
                    "WHERE A.记录ID=B.ID And A.数据来源 in (0,3) AND A.终止版本 IS NULL AND A.记录类型 =1 AND instr(NVL(B.签名人,'QMR'),'/',1)=0 AND A.记录人=[2] AND B.文件ID=[1]" & vbNewLine & _
                    "MINUS" & vbNewLine & _
                    "SELECT DISTINCT B.ID,B.发生时间" & vbNewLine & _
                    "FROM 病人护理明细 A,病人护理数据 B" & vbNewLine & _
                    "WHERE A.记录ID=B.ID And A.数据来源 in (0,3) AND A.终止版本 IS NULL AND A.记录类型 =5 AND A.记录人=[2] AND B.文件ID=[1]"
            Call SQLDIY(gstrSQL)
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查是否存在未签名的数据", mlng文件ID, gstrUserName)
            If rsTemp.RecordCount = 0 Then
                '101151,陈刘,2016-10-19,添加未签名数据记录人提示
                lngStart = GetStartRow(VsfData.ROW)
                If mbln护士 = True Then
                    strInfo = VsfData.TextMatrix(lngStart, mlngOperator)
                Else
                    strInfo = VsfData.TextMatrix(lngStart, VsfData.Cols - 1)
                End If
                If strInfo <> "" Then strInfo = "  当前数据记录人：" & strInfo
                RaiseEvent AfterRowColChange("没有找到需要签名的数据（只能对自己登记或修改的数据进行签名）！" & strInfo, True, mblnSign, mblnArchive)
                Exit Function
            End If
        Else '交班签名
            lngStart = GetStartRow(VsfData.ROW)
            '首先进行数据判断:是否选择已经签名的数据
            If Val(VsfData.TextMatrix(lngStart, mlngRecord)) = 0 Then
                RaiseEvent AfterRowColChange("请先选择要进行交班签名的数据行！", True, mblnSign, mblnArchive)
                Exit Function
            End If
            '对于分组数据交班签名时，只需要验证起始行
            lngDemo = Val(VsfData.TextMatrix(lngStart, mlngDemo))
            If lngDemo > 1 Then '寻找分组数据起始行
                lngRow = lngStart
                lngStart = lngRow - lngDemo + 1
                If Val(VsfData.TextMatrix(lngStart, mlngDemo)) <> 1 Then
                    For lngStart = lngRow To VsfData.FixedRows Step -1
                        If Val(VsfData.TextMatrix(lngStart, mlngDemo)) = 1 Then
                            lngRow = lngStart
                            Exit For
                        End If
                    Next lngStart
                    If lngStart < VsfData.FixedRows Then Exit Function
                    lngStart = lngRow
                End If
            End If
            
            gstrSQL = "" & _
                    "SELECT DISTINCT B.ID,B.发生时间" & vbNewLine & _
                    "FROM 病人护理明细 A,病人护理数据 B" & vbNewLine & _
                    "WHERE A.记录ID=B.ID And A.数据来源 in (0,3) AND A.终止版本 IS NULL AND A.记录类型 =5 AND Instr(NVL(B.签名人,'QMR'),'/',1)=0 And B.交班签名人 IS NULL AND A.记录ID=[2] AND B.文件ID=[1]"
            Call SQLDIY(gstrSQL)
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查是否存在未签名的数据", mlng文件ID, Val(VsfData.TextMatrix(lngStart, mlngRecord)))
            '记录集=0说明没有找到数据
            If rsTemp.RecordCount = 0 Then
                RaiseEvent AfterRowColChange("没有找到需要交班签名的数据（请确认当前选择的数据是否已经签名(不包含审签和交班签名)）！", True, mblnSign, mblnArchive)
                Exit Function
            End If
            '对于分组数据交班签名时，需要对本分组的所有行进行签名
            lngDemo = Val(VsfData.TextMatrix(lngStart, mlngDemo))
            If lngDemo > 0 Then
                '肯定找的到数据
                gstrSQL = "" & _
                    "SELECT DISTINCT B.ID,B.发生时间" & vbNewLine & _
                    "FROM 病人护理明细 A,病人护理数据 B" & vbNewLine & _
                    "WHERE A.记录ID=B.ID And A.数据来源 in (0,3) AND A.终止版本 IS NULL AND A.记录类型 =5 AND INSTR(NVL(B.签名人,'QMR'),'/',1)=0 And B.交班签名人 IS NULL AND B.发生时间 between [2] And [3] AND B.文件ID=[1]"
                Call SQLDIY(gstrSQL)
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查是否存在未签名的数据", mlng文件ID, CDate(Format(VsfData.TextMatrix(lngStart, mlngActTime), "YYYY-MM-DD HH:mm")), CDate(Format(VsfData.TextMatrix(lngStart, mlngActTime), "YYYY-MM-DD HH:mm") & ":59"))
            End If
            
            mintVerify_Last = Val(IIf(VsfData.TextMatrix(lngStart, mlngSignLevel) = "", 9, Val(VsfData.TextMatrix(lngStart, mlngSignLevel)) + 1))
            
        End If
    End If
    
    '准备签名
    strSignTime = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    With rsTemp
        Do While Not .EOF
            str行错误 = ""
            blnSign = SignName(Val(!ID), Format(!发生时间, "yyyy-MM-dd HH:mm:ss"), strSignTime, bln审签, str状态, str行错误, blnExchange)
            If str行错误 <> "" Then
                str错误 = str错误 & vbCrLf & "发生时间=[" & Format(!发生时间, "yyyy-MM-dd HH:mm:ss") & "]" & str行错误
            End If
            If Not blnSign Then Exit Do
            If Not blnRefresh Then blnRefresh = blnSign
            .MoveNext
        Loop
    End With
    
    
    If blnRefresh And Not mblnVerify Then Call ShowMe(mfrmParent, mlng文件ID, mlng病人ID, mlng主页ID, mlng病区ID, mint婴儿, mstrPrivs, mblnEditable, mint页码)
    If str错误 <> "" Then MsgBox "签名时发生以下错误：" & str错误, vbInformation, gstrSysName
    SignMe = blnRefresh
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub UnSignMe(Optional ByVal bln审签 As Boolean = False, Optional blnSingleCancel As Boolean = False)
    'blnSingleCancel 是否单条取消
    Dim intPos As Integer
    Dim lngStart As Long                '启始行
    Dim lngRecord As Long
    Dim blnOK As Boolean
    Dim strSignTime As String           '签名时间
    Dim strRecord As String, strSQLWhere As String
    Dim blnClear As Boolean             '取消签名时是否清除该版本的数据回退到上次签名后的状态
    Dim blnTrans As Boolean
    Dim strSQLTime() As String, strSQLSign() As String, strSQLCollect() As String
    Dim blnUnSign As Boolean, arrUnsign(), strUnsignID As String, strDays As String, strDate As String
    ReDim Preserve strSQLTime(1 To 1)
    ReDim Preserve strSQLSign(1 To 1)
    ReDim Preserve strSQLCollect(1 To 1)
    
    Dim clsSign As Object
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    '首先最后一次是本人的签名，根据当前选择数据的签名时间，批量取消签名
    
    If mlng病人ID = 0 Then Exit Sub
    
    '必要性检查
    '当前记录是新记录则退出
    If FormatValue(VsfData.TextMatrix(VsfData.ROW, mlngRowCount)) = "" Then Exit Sub
    lngStart = GetStartRow(VsfData.ROW)
    lngRecord = Val(VsfData.TextMatrix(lngStart, mlngRecord))
    If lngRecord = 0 Then
        RaiseEvent AfterRowColChange("新增记录不存在取消签名！", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    '当前记录未签名则退出
    If FormatValue(VsfData.TextMatrix(lngStart, mlngSigner)) = "" Then
        RaiseEvent AfterRowColChange("当前记录还未签名！", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    '审签：当前记录未审签则退出；平签：当前记录已审签则退出
    intPos = InStr(1, FormatValue(VsfData.TextMatrix(lngStart, mlngSigner)), "/")
    If bln审签 Then
        If intPos = 0 Then
            RaiseEvent AfterRowColChange("当前记录未审签，无法执行取消审签操作！", True, mblnSign, mblnArchive)
            Exit Sub
        End If
    Else
        If intPos <> 0 Then
            RaiseEvent AfterRowColChange("当前记录已审签，请取消审签后再操作！", True, mblnSign, mblnArchive)
            Exit Sub
        End If
    End If
    '当前记录的最后签名人不是本人则退出
    '汇总数据也要签名,因此去掉条件: And B.汇总类别=0
    gstrSQL = "" & _
              " SELECT  A.记录人,A.记录时间,A.项目名称,B.签名人" & vbNewLine & _
              " FROM 病人护理明细 A,病人护理数据 B" & vbNewLine & _
              " WHERE A.记录ID=B.ID AND B.文件ID=[1] AND A.记录ID=[2] AND A.记录类型=" & IIf(bln审签, 15, 5) & vbNewLine & _
              " ORDER BY A.项目名称 DESC"
    Call SQLDIY(gstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "当前记录的最后签名人不是本人则退出", mlng文件ID, lngRecord)
    If rsTemp.RecordCount = 0 Then
        RaiseEvent AfterRowColChange("无法找到已" & IIf(bln审签, "审签", "签名") & "的数据，可能是数据变化未刷新导致，请刷新数据后再试！", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    
    If bln审签 = False And InStr(1, NVL(rsTemp!签名人), "/") <> 0 Then
        RaiseEvent AfterRowColChange("当前记录已审签，可能是数据变化未刷新导致，请刷新数据、取消审签后再操作！", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    
    If NVL(rsTemp!记录人) <> gstrUserName Then
        RaiseEvent AfterRowColChange("您不是最后签名人，不能执行本操作！", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    
    
    '100658,陈刘,2016-10-20,增加单条记录回退签名,审签
     If blnSingleCancel = True Then
        strRecord = GetSelectRowRecordId(VsfData.ROW)
     Else
        strRecord = ""
     End If
     strSQLWhere = ""
     If strRecord <> "" Then
        strSQLWhere = " And B.id in (Select /*+ CARDINALITY(b 10) */  COLUMN_VALUE from  Table(f_Num2list([4])) b)"
     End If
    
    '提取所有数据准备取消签名或审签(记录时间不为空表示新版签名;)
    '汇总数据也要签名,因此去掉条件: And B.汇总类别=0
    If Not IsNull(rsTemp!记录时间) Then
        gstrSQL = "" & _
                  " SELECT  A.项目ID AS 证书ID,A.项目分组,B.发生时间,B.ID,B.签名人" & vbNewLine & _
                  " FROM 病人护理明细 A,病人护理数据 B" & vbNewLine & _
                  " WHERE A.记录ID=B.ID AND B.文件ID=[1] And A.记录人=[2] And A.记录时间=[3] " & strSQLWhere & _
                  " AND A.记录类型=" & IIf(bln审签, 15, 5) & _
                  " Order by B.发生时间"
                  
        Call SQLDIY(gstrSQL)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取所有数据准备取消签名或审签", mlng文件ID, gstrUserName, CDate(rsTemp!记录时间), strRecord)
    Else
        gstrSQL = "" & _
                  " SELECT  A.项目ID AS 证书ID,A.项目分组,B.发生时间,B.ID,B.签名人" & vbNewLine & _
                  " FROM 病人护理明细 A,病人护理数据 B" & vbNewLine & _
                  " WHERE A.记录ID=B.ID AND B.文件ID=[1] And A.记录人=[2] And A.项目名称=[3] " & strSQLWhere & _
                  " AND A.记录类型=" & IIf(bln审签, 15, 5) & _
                  " Order by B.发生时间"
        Call SQLDIY(gstrSQL)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取所有数据准备取消签名或审签", mlng文件ID, gstrUserName, CStr(rsTemp!项目名称), strRecord)
    End If
    
    '签名后不允许修改，如需修改必须回退签名，因此取消普签时不存在提示是否回退数据的问题，审签自动回退，所以取消提示
    '--------------------
    '询问是否需要清除数据
'    If Not bln审签 Then
'        blnClear = (MsgBox("取消签名时是否该版本的数据回退到上次签名后的状态？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes)
'    End If
    blnClear = True
    '--------------------
    arrUnsign = Array()
    strUnsignID = "": strDays = ""
    If bln审签 = True Then
        Do While Not rsTemp.EOF
            '81535:审签可能存在数据时间的修改, 回退审签则检查回退的时间是否存在已经签名的汇总数据
            If IsDate(NVL(rsTemp!项目分组)) Then
                If Format(rsTemp!发生时间, "yyyy-MM-dd HH:mm:ss") <> Format(rsTemp!项目分组, "yyyy-MM-dd HH:mm:ss") Then
                    If ISCollectSigned(mlng文件ID, Format(rsTemp!项目分组, "YYYY-MM-DD"), Format(rsTemp!项目分组, "HH:MM")) Then
                        ReDim Preserve arrUnsign(UBound(arrUnsign) + 1)
                        arrUnsign(UBound(arrUnsign)) = "当前数据时间：" & Format(rsTemp!发生时间, "yyyy-MM-dd HH:mm") & vbTab & "修改前数据时间：" & Format(rsTemp!项目分组, "yyyy-MM-dd HH:mm")
                        strUnsignID = strUnsignID & "," & Val(rsTemp!ID)
                    ElseIf ISCollectSigned(mlng文件ID, Format(rsTemp!发生时间, "YYYY-MM-DD"), Format(rsTemp!发生时间, "HH:MM")) Then
                        ReDim Preserve arrUnsign(UBound(arrUnsign) + 1)
                        arrUnsign(UBound(arrUnsign)) = "当前数据时间：" & Format(rsTemp!发生时间, "yyyy-MM-dd HH:mm") & vbTab & "修改前数据时间：" & Format(rsTemp!项目分组, "yyyy-MM-dd HH:mm")
                        strUnsignID = strUnsignID & "," & Val(rsTemp!ID)
                    Else
                        gstrSQL = "Zl_病人护理数据_发生时间(" & rsTemp!ID & ",to_date('" & rsTemp!项目分组 & "','yyyy-MM-dd hh24:mi:ss'))"
                        strSQLSign(ReDimArray(strSQLSign)) = gstrSQL
                        
                        '同时修正处理汇总数据
                        '必须要处理昨天，因为可能存在跨天汇总的数据，且当前时间刚好在第二天的情况
                        strDate = Format(rsTemp!发生时间, "yyyy-MM-dd")
                        If InStr(1, "," & strDays & ",", "," & Format(DateAdd("d", -1, CDate(strDate)), "yyyy-MM-dd") & ",") = 0 Then
                            gstrSQL = "ZL_汇总数据_UPDATE(" & mlng文件ID & ",'" & Format(DateAdd("d", -1, CDate(strDate)), "yyyy-MM-dd") & "')"
                            strSQLCollect(ReDimArray(strSQLCollect)) = gstrSQL
                            strDays = strDays & "," & Format(DateAdd("d", -1, CDate(strDate)), "yyyy-MM-dd")
                            If Left(strDays, 1) = "," Then strDays = Mid(strDays, 2)
                        End If
                        If InStr(1, "," & strDays & ",", "," & strDate & ",") = 0 Then
                            gstrSQL = "ZL_汇总数据_UPDATE(" & mlng文件ID & ",'" & strDate & "')"
                            strSQLCollect(ReDimArray(strSQLCollect)) = gstrSQL
                            strDays = strDays & "," & strDate
                            If Left(strDays, 1) = "," Then strDays = Mid(strDays, 2)
                        End If
                        strDate = Format(rsTemp!项目分组, "yyyy-MM-dd")
                        If InStr(1, "," & strDays & ",", "," & Format(DateAdd("d", -1, CDate(strDate)), "yyyy-MM-dd") & ",") = 0 Then
                            gstrSQL = "ZL_汇总数据_UPDATE(" & mlng文件ID & ",'" & Format(DateAdd("d", -1, CDate(strDate)), "yyyy-MM-dd") & "')"
                            strSQLCollect(ReDimArray(strSQLCollect)) = gstrSQL
                            strDays = strDays & "," & Format(DateAdd("d", -1, CDate(strDate)), "yyyy-MM-dd")
                            If Left(strDays, 1) = "," Then strDays = Mid(strDays, 2)
                        End If
                        If InStr(1, "," & strDays & ",", "," & strDate & ",") = 0 Then
                            gstrSQL = "ZL_汇总数据_UPDATE(" & mlng文件ID & ",'" & strDate & "')"
                            strSQLCollect(ReDimArray(strSQLCollect)) = gstrSQL
                            strDays = strDays & "," & strDate
                            If Left(strDays, 1) = "," Then strDays = Mid(strDays, 2)
                        End If
                    End If
                End If
            End If
        rsTemp.MoveNext
        Loop
        strUnsignID = Mid(strUnsignID, 2)
        If strUnsignID <> "" Then
            If MsgBox("您审签时修改了部分数据的时间，且这些数据中的部分数据对应的汇总数据已经签名，这些数据将不能进行回退，请问您是否继续？" & vbCrLf & _
                "是：继续，但部分数据将不能被回退" & vbCrLf & "否：终止本次审签回退" & vbCrLf & vbCrLf & _
                "不能回退的数据信息如下：" & vbCrLf & Join(arrUnsign, vbCrLf), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
    End If
    
    mstrVerify = ""
    If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
    Do While Not rsTemp.EOF
        If (bln审签 = False And InStr(1, NVL(rsTemp!签名人), "/") = 0) Or bln审签 = True Then
            blnUnSign = InStr(1, "," & strUnsignID & ",", "," & rsTemp!ID & ",") = 0
            If blnUnSign = True Then
                If NVL(rsTemp!证书ID, 0) > 0 Then
                    '数字签名验证，只验证一次
                    If clsSign Is Nothing Then
                        If gobjESign Is Nothing Then
                            On Error Resume Next
                            Set gobjESign = CreateObject("zl9ESign.clsESign")
                            If Err <> 0 Then Err.Clear
                            On Error GoTo 0
                            If Not gobjESign Is Nothing Then Call gobjESign.Initialize(gcnOracle, glngSys)
                        End If
                        Set clsSign = gobjESign
                        
                        If Not clsSign Is Nothing Then
                            If Not clsSign.CheckCertificate(gstrDBUser) Then
                                gcnOracle.RollbackTrans
                                Exit Sub
                            End If
                        Else
                            gcnOracle.RollbackTrans
                            RaiseEvent AfterRowColChange("电子签名部件未能正确安，回退操作不能继续！", True, mblnSign, mblnArchive)
                            Exit Sub
                        End If
                    End If
                End If
                
                '取消签名
                gstrSQL = "ZL_病人护理数据_UNSIGNNAME("
                gstrSQL = gstrSQL & mlng文件ID & ","
                gstrSQL = gstrSQL & "To_Date('" & Format(rsTemp!发生时间, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd hh24:mi:ss')," & IIf(blnClear, "1", "0") & ")"
                strSQLTime(ReDimArray(strSQLTime)) = gstrSQL
                'Call zlDatabase.ExecuteProcedure(gstrSQL, "执行取消签名")
                
                If InStr(1, mstrVerify & ",", "," & rsTemp!ID & ",") = 0 Then
                    mstrVerify = mstrVerify & "," & rsTemp!ID
                End If
            End If
        End If
        rsTemp.MoveNext
    Loop
    
    '撤销签名审前后重新结算数据行
    gcnOracle.BeginTrans
    blnTrans = True
    For intPos = 1 To UBound(strSQLTime)
        If strSQLTime(intPos) <> "" Then
            Call zlDatabase.ExecuteProcedure(strSQLTime(intPos), "执行取消签名")
        End If
    Next intPos
    '审签时修改了数据时间，回退时需要同步处理
    For intPos = 1 To UBound(strSQLSign)
        If strSQLSign(intPos) <> "" Then
            Call zlDatabase.ExecuteProcedure(strSQLSign(intPos), "发生时间修正")
        End If
    Next intPos
    '汇总数据修正
    For intPos = 1 To UBound(strSQLCollect)
        If strSQLCollect(intPos) <> "" Then
            Call zlDatabase.ExecuteProcedure(strSQLCollect(intPos), "汇总数据修正")
        End If
    Next intPos
    
    '取消审签才重新结算数据行，因为签名后不能修改数据，所以不牵扯数据行数变化的问题。
    '审签是可能原有数据5行审签时改为3行，回退审签是需要还原数据和行数
    If Not PreseData(bln审签) Then GoTo ErrHand:
    
    gcnOracle.CommitTrans
    blnTrans = False
    mstrVerify = ""
    '刷新数据
    Call ShowMe(mfrmParent, mlng文件ID, mlng病人ID, mlng主页ID, mlng病区ID, mint婴儿, mstrPrivs, mblnEditable, mint页码)
    Exit Sub
ErrHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Sub



Private Function SignName(ByVal lngRecordId As Long, ByVal strStart As String, ByVal strSignTime As String, ByVal bln审签 As Boolean, _
    str状态 As String, Optional str错误 As String, Optional ByVal blnExchange As Boolean = False) As Boolean
    '******************************************************************************************************************
    '功能:
    '
    '
    '******************************************************************************************************************
    Dim oSign As cTendSign
    Dim strSource As String             '审签源数据串
    Dim lngLoop As Long
    Dim rs As New ADODB.Recordset
    Dim strChangeTime As String
    On Error GoTo ErrHand
    
    '初始处理
    '------------------------------------------------------------------------------------------------------------------
    strSource = ""
    
    '获取要签名的内容(汇总数据也要签名,因此去掉条件: And B.汇总类别=0)
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = " Select a.记录类型,a.项目分组,a.项目序号,a.项目名称,a.项目类型,a.记录内容,a.项目单位,a.记录标记,a.体温部位,a.记录组号,a.复试合格,a.未记说明,a.记录人,a.记录时间  " & _
              " From 病人护理明细 a,病人护理数据 b,病人护理文件 c " & _
              " Where a.记录id=b.ID And b.文件ID=c.ID AND MOD(A.记录类型,10)<>5 And a.终止版本 Is Null And C.ID=[1] And b.发生时间=[2]" & _
              " Order by a.项目序号"
    Call SQLDIY(gstrSQL)
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "获取要签名的内容", mlng文件ID, CDate(strStart))
    If rs.BOF = False Then
        Do While Not rs.EOF
            For lngLoop = 0 To rs.Fields.Count - 1
                strSource = strSource & CStr(zlCommFun.NVL(rs.Fields(lngLoop).Value, ""))
            Next
            rs.MoveNext
        Loop
    End If
    'Debug.Print "开始签名：" & Now & vbCrLf & strSource
    If strSource = "" Then
        RaiseEvent AfterRowColChange("当前没有需要签名的信息！", True, mblnSign, mblnArchive)
        Exit Function
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    On Error Resume Next
    Err = 0
    '81535:提取原有数据的时间(审签可能修改数据发生时间)
    strChangeTime = ""
    If bln审签 = True And Not mobjVerify Is Nothing Then
        If mobjVerify.Count > 0 Then
            If Not Format(mobjVerify("_" & lngRecordId), "YYYY-MM-DD HH:mm") = Format(strStart, "YYYY-MM-DD HH:mm") Then
                strChangeTime = Format(mobjVerify("_" & lngRecordId), "YYYY-MM-DD HH:mm:ss")
            End If
        End If
    End If
    '76223:刘鹏飞,2012-09-13,电子签名添加时间戳信息
    '43588:刘鹏飞,2012-09-13,添加记录单审签模式
    Set oSign = frmTendFileSign.ShowMe(Me, mstrPrivs, mlng文件ID, mlng病区ID, mintVerify_Last, strSource, bln审签, str状态, str错误, mintSignMode, blnExchange)
    On Error GoTo ErrHand
    
    If Not oSign Is Nothing Then
        gstrSQL = "ZL_病人护理数据_SIGNNAME("
        gstrSQL = gstrSQL & mlng文件ID & ","
        gstrSQL = gstrSQL & "To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss')," & IIf(bln审签, 1, 0) & ","
        gstrSQL = gstrSQL & "'" & oSign.姓名 & "',"
        gstrSQL = gstrSQL & "'" & oSign.签名信息 & "'," & oSign.签名级别 & ","
        gstrSQL = gstrSQL & oSign.证书ID & ","
        gstrSQL = gstrSQL & oSign.签名方式 & ",'" & oSign.时间戳 & "'," & IIf(blnExchange, 1, 0) & ",'" & oSign.时间戳信息 & "',"
        gstrSQL = gstrSQL & "To_Date('" & strSignTime & "','yyyy-mm-dd hh24:mi:ss'),'" & strChangeTime & "')"
        
        Call zlDatabase.ExecuteProcedure(gstrSQL, "执行签名")
        SignName = True
    End If
    
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function PreseData(ByVal blnTrue As Boolean) As Boolean
'功能：取消审签数据后重新结算数据行
    Dim rsTemp As New ADODB.Recordset
    Dim str条件 As String, intPos As Long
    Dim lngRow As Long, lngCol As Long, lngRecord As Long, lngMutilRow As Long, strTime As String
    Dim arrData, arrMutilRow
    Dim strSQLData() As String
    Dim blnSave As Boolean, i As Long
    ReDim Preserve strSQLData(1 To 1)
    On Error GoTo ErrHand
    
    If Left(mstrVerify, 1) = "," Then mstrVerify = Mid(mstrVerify, 2)
    If mstrVerify = "" Or blnTrue = False Then GoTo ErrEnd
    
   
    str条件 = mstrSQL条件
    
    mstrSQL = "Select /*+ RULE */ '' AS 备用,to_char(发生时间,'yyyy-MM-dd hh24:mi:ss') AS 发生时间,'' AS 选择,to_char(发生时间,'YYYY') AS 年份," & Mid(mstrSQL列, 12) & vbCrLf & _
                " From (Select 记录组号,时间 as 备用,发生时间," & Mid(mstrSQL中, 2) & vbCrLf & _
                "        From (Select nvl(c.记录组号,0) 记录组号,l.发生时间," & Mid(mstrSQL内, 2) & vbCrLf & _
                "               From 病人护理数据 l, 病人护理明细 c,病人护理文件 f,病人护理打印 p," & vbCrLf & _
                "       (SELECT COLUMN_VALUE FROM TABLE(CAST(F_NUM2LIST([6]) AS ZLTOOLS.T_NUMLIST))) G" & vbNewLine & _
                "               Where l.ID=p.记录ID And l.Id = c.记录id And l.文件ID+0=f.ID+0 And f.ID=p.文件ID " & _
                "               And c.终止版本 Is Null And MOD(c.记录类型,10)<>5  " & _
                "               And f.id=[1] And f.病人id = [2] And f.主页id = [3] And Nvl(f.婴儿,0)=[4] And l.ID=G.COLUMN_VALUE)" & _
                IIf(str条件 <> "", "Where " & str条件, "") & _
                "       Group By 日期, 时间, 发生时间,记录组号," & IIf(mbln护士 = True, "护士,", "护士L,") & "签名人,签名时间" & _
                                "       Order By 发生时间,记录组号," & IIf(mbln护士 = True, "护士,", "护士L,") & "签名人,签名时间)"
     Call SQLDIY(mstrSQL)
     Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "检查是否存在未签名的数据", mlng文件ID, mlng病人ID, mlng主页ID, mint婴儿, mint页码, mstrVerify)
     If rsTemp.RecordCount > 0 Then
        blnSave = False: arrMutilRow = Array()
        Set vsTest.DataSource = rsTemp
        For lngRow = vsTest.FixedRows To vsTest.Rows - 1
            lngMutilRow = 0
            lngRecord = Val(vsTest.TextMatrix(lngRow, mlngRecord))
            If lngRecord <> 0 Then
                blnSave = True
                strTime = vsTest.TextMatrix(lngRow, mlngActTime)
                '分类汇总处理:计算数据行时需包含分类明细数据行
                If Val(vsTest.TextMatrix(lngRow, mlngCollectType)) < 0 Then
                    If lngRow + 1 < vsTest.Rows Then
                        If Val(vsTest.TextMatrix(lngRow + 1, mlngRecord)) > 0 And Val(vsTest.TextMatrix(lngRow + 1, mlngCollectType)) < 0 And _
                            strTime = vsTest.TextMatrix(lngRow + 1, mlngActTime) Then
                            blnSave = False
                        End If
                    End If
                End If
                
                For lngCol = mlngTime + 1 To mlngNoEditor - 1
                    If vsTest.TextMatrix(lngRow, lngCol) <> "" Then
                        '准备赋值
                        With txtLength
                            .Width = VsfData.ColWidth(lngCol)
                            .Text = Replace(Replace(Replace(vsTest.TextMatrix(lngRow, lngCol), Chr(10), ""), Chr(13), ""), Chr(1), "")
                            .FontName = VsfData.FontName
                            .FontSize = VsfData.FontSize
                            .FontBold = VsfData.FontBold
                            .FontItalic = VsfData.FontItalic
                        End With
                        arrData = GetData(txtLength.Text)
                        If UBound(arrData) > lngMutilRow Then
                            lngMutilRow = UBound(arrData)
                            If Trim(arrData(lngMutilRow)) = "" Then lngMutilRow = lngMutilRow - 1
                        End If
                    End If
                Next
                ReDim Preserve arrMutilRow(UBound(arrMutilRow) + 1)
                arrMutilRow(UBound(arrMutilRow)) = lngMutilRow + 1
                If blnSave = True Then
                    '----此处主要计算分类汇总的行数
                    lngMutilRow = 0
                    '计算分类明细的数据行数
                    For i = 1 To UBound(arrMutilRow)
                        lngMutilRow = lngMutilRow + arrMutilRow(i)
                    Next i
                    '汇总主数据行数如果大于分类明细数据行数+1(1为默认的总量行数),则以主数据行数为准,否则以明细数据行+1为准
                    If lngMutilRow + 1 > Val(arrMutilRow(0)) Then
                        lngMutilRow = lngMutilRow + 1
                    Else
                        lngMutilRow = Val(arrMutilRow(0))
                    End If
                    arrMutilRow = Array()
                    '一行结束时，产生打印解析数据
                    If Val(vsTest.TextMatrix(lngRow, mlngRowCount)) <> lngMutilRow Then
                        gstrSQL = "ZL_病人护理打印_UPDATE(" & mlng文件ID & ",to_date('" & strTime & "','yyyy-MM-dd hh24:mi:ss')" & "," & lngMutilRow & ")"
                        strSQLData(ReDimArray(strSQLData)) = gstrSQL
                    End If
                End If
            End If
        Next
     End If
     
    '执行过程
    For intPos = 1 To UBound(strSQLData)
        If strSQLData(intPos) <> "" Then
            Call zlDatabase.ExecuteProcedure(strSQLData(intPos), "产生打印解析数据")
        End If
    Next intPos
ErrEnd:
    PreseData = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function CancelMe() As Boolean
    CancelMe = True
    mblnVerify = False
    mstrVerify = ""
    Set mobjVerify = Nothing
    mblnChange = False
    Call ShowMe(mfrmParent, mlng文件ID, mlng病人ID, mlng主页ID, mlng病区ID, mint婴儿, mstrPrivs, mblnEditable, mint页码)
End Function

Public Function SaveME() As Boolean
    If Not CheckData Then Exit Function
    If Not SaveData Then Exit Function
    
    mblnShow = False
    Call InitCons
    SaveME = True
    
    Call ShowMe(mfrmParent, mlng文件ID, mlng病人ID, mlng主页ID, mlng病区ID, mint婴儿, mstrPrivs, mblnEditable, mint页码)
End Function

Public Function ShowMe(ByVal frmParent As Form, ByVal lngFileID As Long, ByVal lngPatiID As Long, ByVal lngPageId As Long, _
    ByVal lngDeptID As Long, ByVal intBaby As Integer, Optional ByVal strPrivs As String, Optional ByVal blnEditable As Boolean = True, _
    Optional ByVal int页码 As Integer = -1, Optional ByVal blnClear As Boolean = True) As Boolean
    '******************************************************************************************************************
    '功能： 显示护理记录文件内容
    '参数： frmParent           上级窗体对象
    '       lngPatiID           病人id
    '       lngPageID           主页id
    '       lngDeptID           要显示护理记录的科室
    '       intBaby             婴儿标志
    '       blnEditable         如果为假,说明是做为查询子窗体在使用,取消与编辑相关的功能
    '       blnClear            如果为真,清除mrsDataMap记录集;当换页时应传假,保留用户修改的数据以备显示、保存使用
    '返回： 无
    '******************************************************************************************************************
    Dim lngRow As Long
    Dim rsTemp As New ADODB.Recordset
    Dim strtmp As String
    On Error GoTo ErrHand
    Err = 0
    
    mblnInit = False
    If mblnChange Then
        If MsgBox("当前病人的数据还未保存，点“是”进行保存，点“否”将放弃本次修改！", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Call VsfData_EnterCell
            '60218:刘鹏飞,2013-04-22,添加CheckData，不然在没有切换页的情况下无法保存数据
            If CheckData Then Call SaveData
        End If
    End If
    
    mblnGroupNew = False
    mblnGroupApp = False
    mstrGroupRow = ""
    mblnClear = blnClear
    mint起始页码 = 1
    mint页码 = int页码
    mlng文件ID = lngFileID
    mlng病人ID = lngPatiID
    mlng主页ID = lngPageId
    mlng病区ID = lngDeptID
    mint婴儿 = intBaby
    mstrPrivs = strPrivs
    'mblnBlowup = (zlDatabase.GetPara("护理文件显示模式", glngSys, 1255, 0) = 1)
    UserControl.Font = IIf(mblnBlowup = True, 12, 9)
    Set mfrmParent = frmParent
    
    mintNORule = Val(zlDatabase.GetPara("护理文件页码规则", glngSys, 1255, 0))
    mintPreDays = Val(zlDatabase.GetPara("超期录入护理数据天数", glngSys, 1255, "1"))
    mstrMaxDate = Format(DateAdd("d", mintPreDays, zlDatabase.Currentdate), "yyyy-MM-dd HH:mm")
    mintCollectDef = Val(zlDatabase.GetPara("小结缺省格式", glngSys, 1255))
    mintPageSpan = Val(zlDatabase.GetPara("跨页数据只显示在第一页", glngSys, 1255))
    '68739:刘鹏飞,2014-1-2,添加"小结标识颜色"
    mlngCollectColor = Val(zlDatabase.GetPara("小结标识颜色", glngSys, 1255, "255"))
    
    '43588:刘鹏飞,2012-09-13,添加记录单审签模式
    strtmp = Val(zlDatabase.GetPara("记录单审签模式", glngSys, 1255))
    If Val(strtmp) >= 0 And Val(strtmp) <= 1 Then
        mintSignMode = CInt(Val(strtmp))
    Else
        mintSignMode = 0
    End If
    
    '--58414,刘鹏飞,2013-04-10,添加护士、签名列显示模式
    mlngSingerType = Val(zlDatabase.GetPara("护士、签名列显示模式", glngSys, 1255, "2"))
    If InStr(1, ",0,1,2,3,", "," & mlngSingerType & ",") = 0 Then mlngSingerType = 2
    
    If mrsItems.State = 0 Then
        Call InitMenuBar
        Call InitEnv            '初始化环境
        cbsThis.ActiveMenuBar.Visible = False
        cbsThis.RecalcLayout
    End If
    
    Call InitVariable
    Call InitCons
    
    If Not ReadStruDef Then Exit Function
    Call zlRefresh
    mblnInit = True
    mblnEditable = blnEditable And Not gblnMoved And Not mblnArchive
    
    '--48659:刘鹏飞,2012-09-14,添加字段'说明'
    RaiseEvent ShowTipInfo(VsfData, "", True)
    RaiseEvent AfterDataChanged(mblnChange Or mblnVerify)
    RaiseEvent AfterRefresh
    
'    Call OutputRsData(mrsSelItems)
    VsfData.Refresh
    ShowMe = True
    Exit Function
    
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub SetFontSize(ByVal bytSize As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置字体大小
    '入参:bytSize：0-小(缺省)，1-大
    '编制:刘鹏飞
    '日期:2012-06-19 15:16
    '问题:50807
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mblnBlowup = IIf(bytSize = 1, True, False)
    Call ReSetFontSize
    If mint页码 = -1 Or mblnInit = False Then Exit Sub
    If Not DataMap_Save Then Exit Sub
    '更新查询SQL
    '重新提取数据
    mblnInit = False
    Call InitVariable
    Call InitCons
    If Not ReadStruDef Then Exit Sub
    Call zlRefresh
    mblnInit = True
    VsfData.Refresh
    cbsThis.RecalcLayout
End Sub

Private Sub ReSetFontSize()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置字体大小
    '入参:bytSize：0-小(缺省)，1-大
    '编制:刘鹏飞
    '日期:2012-06-20 15:15:00
    '问题:50807
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCtrl As Control
    Dim CtlFont As StdFont
    Dim bytFontSize As Byte
    bytFontSize = BlowUp(9)
    
    UserControl.FontSize = bytFontSize
    UserControl.FontName = "宋体"
    
    Set CtlFont = cbsThis.Options.Font
    If CtlFont Is Nothing Then
        Set CtlFont = UserControl.Font
    End If
    CtlFont.Size = bytFontSize
    Set cbsThis.Options.Font = CtlFont
    
    lblPage.FontSize = bytFontSize
    lblPage.Left = 30
    txtPage.FontSize = bytFontSize
    txtPage.Height = TextHeight("刘") + TextHeight("刘") / 3
    txtPage.Width = TextWidth("9999")
    txtPage.Left = lblPage.Left + lblPage.Width
    Label3.FontSize = bytFontSize
    Label3.Left = txtPage.Left + txtPage.Width
    picPage.Width = Label3.Left + Label3.Width
    picPage.Height = txtPage.Height + 40
    txtPage.Top = (picPage.Height - txtPage.Height) \ 2
    lblPage.Top = txtPage.Top + (txtPage.Height - lblPage.Height) \ 2
    Label3.Top = lblPage.Top
End Sub

Private Function CheckFlip() As Boolean
    Dim blnExit As Boolean, blnNULL As Boolean
    Dim lngRow As Long, lngCol As Long, lngRows As Long, lngCols As Long
    Dim lngOldRow As Long, lngOldCol As Long, lngEditCol As Long
    Dim strInfo As String, blnShow As Boolean
    Dim strDate As String
    '页面切换前检查：日期时间正确才允许继续，这样在保存时就不必再检查其它页面的数据了（其它数据在录入时已经进行了检查，此处略过）
    
    '隐藏编辑控件
    lngOldRow = VsfData.ROW: lngOldCol = VsfData.COL
    
    blnShow = mblnShow
    Select Case mintType
    Case 0, 3
        picInput.Visible = False
    Case 1, 2
        lstSelect(mintType - 1).Visible = False
        If mintType = 1 Then
            txtLst.Visible = False
            PicLst.Visible = False
        End If
    Case 4, 5
        picDouble.Visible = False
    Case 6
        picMutilInput.Visible = False
    Case 7
        picDoubleChoose.Visible = False
    Case 8
        picYear.Visible = False
    End Select
    cmdWord.Visible = False
    mintType = -1
    mblnShow = False
    If mblnVerify = True Then
        '审签可能修改时间，此处检查修改了时间的数据对应的汇总数据是否已经签名
        For lngRow = VsfData.FixedRows To VsfData.Rows - 1
            If VsfData.Cell(flexcpChecked, lngRow, mlngChoose) = flexTSChecked And Val(FormatValue(VsfData.TextMatrix(lngRow, mlngDemo))) <= 1 And _
                Val(FormatValue(VsfData.TextMatrix(lngRow, mlngCollectType))) >= 0 And Val(VsfData.TextMatrix(lngRow, mlngRecord)) > 0 Then
                    If mblnDateAd Then
                        strDate = VsfData.TextMatrix(lngRow, mlngYear) & "-" & ToStandDate(VsfData.TextMatrix(lngRow, mlngDate)) & " " & VsfData.TextMatrix(lngRow, mlngTime)
                    Else
                        strDate = VsfData.TextMatrix(lngRow, mlngDate) & " " & VsfData.TextMatrix(lngRow, mlngTime)
                    End If
                    strDate = Format(strDate, "YYYY-MM-DD HH:mm")
                
                    If FormatValue(VsfData.TextMatrix(lngRow, mlngRowCount)) Like "*|1" And strDate <> Format(VsfData.TextMatrix(lngRow, mlngActTime), "YYYY-MM-DD HH:mm") Then
                        VsfData.ROW = lngRow: VsfData.COL = mlngTime
                        If Not CheckDateTime(VsfData.TextMatrix(VsfData.ROW, VsfData.COL), strInfo) Then
                            mblnShow = True: Call VsfData_EnterCell
                            If Not VsfData.RowIsVisible(lngRow) Then VsfData.TopRow = lngRow
                            RaiseEvent AfterRowColChange(strInfo, True, mblnSign, mblnArchive)
                            CheckFlip = False
                            Exit Function
                        End If
                    End If
            End If
        Next lngRow
        mblnShow = blnShow
        VsfData.Select lngOldRow, lngOldCol
        CheckFlip = True
        Exit Function
    End If
    
    mblnShow = False
    
    lngRows = VsfData.Rows - 1
    lngCols = VsfData.Cols - 1
    For lngRow = VsfData.FixedRows To lngRows
        If Val(FormatValue(VsfData.TextMatrix(lngRow, mlngDemo))) > 1 And Trim(FormatValue(VsfData.TextMatrix(lngRow, mlngSigner))) = "" And VsfData.RowHidden(lngRow) = False Then
            blnNULL = True
            For lngCol = mlngTime + 1 To lngCols - 1
                If Not VsfData.ColHidden(lngCol) And lngCol < mlngNoEditor And ISEditAssistant(lngCol) = False Then
                    If FormatValue(VsfData.TextMatrix(lngRow, lngCol)) <> "" And Not (IsDiagonal(lngCol) And InStr(1, FormatValue(VsfData.TextMatrix(lngRow, lngCol)), "/") <> 0) Then
                        blnNULL = False
                        Exit For
                    End If
                End If
            Next
            If blnNULL = True And cbsThis.FindControl(xtpControlButton, conMenu_Edit_Clear).Enabled = True Then
                VsfData.ROW = lngRow
                Call cbsThis_Execute(cbsThis.FindControl(xtpControlButton, conMenu_Edit_Clear))
            End If
        End If
    Next
    
    lngRows = VsfData.Rows - 1
    lngCols = VsfData.Cols - 1
    For lngRow = VsfData.FixedRows To lngRows
        mrsCellMap.Filter = "页号=" & mint页码 & " And 行号=" & lngRow & " And 列号>" & mlngTime
        If mrsCellMap.RecordCount = 0 And Val(VsfData.TextMatrix(lngRow, mlngRecord)) > 0 Then
            mrsCellMap.Filter = "页号=" & mint页码 & " And 行号=" & lngRow & " And 列号>=" & mlngDate
        End If
        'Call OutputRsData(mrsCellMap)
        If mrsCellMap.RecordCount <> 0 Then
            If Not VsfData.RowHidden(lngRow) And Val(FormatValue(VsfData.TextMatrix(lngRow, mlngDemo))) <= 1 And Val(FormatValue(VsfData.TextMatrix(lngRow, mlngCollectType))) >= 0 Then
                blnExit = (FormatValue(VsfData.TextMatrix(lngRow, mlngDate)) = "" Or FormatValue(VsfData.TextMatrix(lngRow, mlngTime)) = "")
                If blnExit = False And mblnDateAd Then
                    blnExit = FormatValue(VsfData.TextMatrix(lngRow, mlngYear)) = ""
                End If
                
                If blnExit Then
                    mrsCellMap.Filter = 0
                    If FormatValue(VsfData.TextMatrix(lngRow, mlngDate)) = "" Then
                        lngCol = mlngDate
                    ElseIf FormatValue(VsfData.TextMatrix(lngRow, mlngTime)) = "" Then
                        lngCol = mlngTime
                    Else
                        lngCol = mlngYear
                    End If
                    VsfData.ROW = lngRow: VsfData.COL = lngCol
                    mblnShow = True: Call VsfData_EnterCell
                    If Not VsfData.RowIsVisible(lngRow) Then VsfData.TopRow = lngRow
                    RaiseEvent AfterRowColChange("请补充日期时间！", True, mblnSign, mblnArchive)
                    CheckFlip = False
                    Exit Function
                Else
                    '日期不为空将检查日期的合法性
                    If Val(VsfData.TextMatrix(lngRow, mlngRecord)) > 0 Then
                        If mblnDateAd Then
                            strDate = VsfData.TextMatrix(lngRow, mlngYear) & "-" & ToStandDate(VsfData.TextMatrix(lngRow, mlngDate)) & " " & VsfData.TextMatrix(lngRow, mlngTime)
                        Else
                            strDate = VsfData.TextMatrix(lngRow, mlngDate) & " " & VsfData.TextMatrix(lngRow, mlngTime)
                        End If
                        strDate = Format(strDate, "YYYY-MM-DD HH:mm")
                        blnExit = (strDate = Format(VsfData.TextMatrix(lngRow, mlngActTime), "YYYY-MM-DD HH:mm"))
                    End If
                    If blnExit = False Then
                        VsfData.ROW = lngRow: VsfData.COL = mlngTime
                        If Not CheckDateTime(VsfData.TextMatrix(VsfData.ROW, VsfData.COL), strInfo) Then
                            mrsCellMap.Filter = ""
                            mblnShow = True: Call VsfData_EnterCell
                            If Not VsfData.RowIsVisible(lngRow) Then VsfData.TopRow = lngRow
                            RaiseEvent AfterRowColChange(strInfo, True, mblnSign, mblnArchive)
                            CheckFlip = False
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next
    '检查新增的子分组数据，如果汇总行已经签名则提示（只处理在原有数据新增的分组数据，因新增的分组数据上面已经检查）
    strDate = ""
    For lngRow = VsfData.FixedRows To lngRows
        If FormatValue(VsfData.TextMatrix(lngRow, mlngRowCount)) Like "*|1" Then
            If Not Val(FormatValue(VsfData.TextMatrix(lngRow, mlngDemo))) >= 1 Then strDate = ""
            If Val(FormatValue(VsfData.TextMatrix(lngRow, mlngDemo))) = 1 And Val(FormatValue(VsfData.TextMatrix(lngRow, mlngCollectType))) >= 0 _
                And Val(FormatValue(VsfData.TextMatrix(lngRow, mlngRecord))) > 0 Then
                If mblnDateAd Then
                    strDate = VsfData.TextMatrix(lngRow, mlngYear) & "-" & ToStandDate(VsfData.TextMatrix(lngRow, mlngDate)) & " " & VsfData.TextMatrix(lngRow, mlngTime)
                Else
                    strDate = VsfData.TextMatrix(lngRow, mlngDate) & " " & VsfData.TextMatrix(lngRow, mlngTime)
                End If
                strDate = Format(strDate, "YYYY-MM-DD HH:mm")
            End If
            
            If IsDate(strDate) And Not VsfData.RowHidden(lngRow) And Val(FormatValue(VsfData.TextMatrix(lngRow, mlngDemo))) > 1 And _
                Val(FormatValue(VsfData.TextMatrix(lngRow, mlngCollectType))) >= 0 And Val(FormatValue(VsfData.TextMatrix(lngRow, mlngRecord))) <= 0 Then
                mrsCellMap.Filter = "页号=" & mint页码 & " And 行号=" & lngRow & " And 列号>" & mlngTime
                If mrsCellMap.RecordCount > 0 Then
                    lngEditCol = 0
                    If CheckCollectIsData(lngRow, 1, lngEditCol) = True Then
                        If ISCollectSigned(mlng文件ID, Format(strDate, "YYYY-MM-DD"), Format(strDate, "HH:MM")) Then
                            VsfData.ROW = lngRow: VsfData.COL = lngEditCol
                            strInfo = "您新增的分组数据所对应的汇总行数据已签名，不允许再添加新的汇总列数据！"
                            mrsCellMap.Filter = ""
                            mblnShow = True: Call VsfData_EnterCell
                            If Not VsfData.RowIsVisible(lngRow) Then VsfData.TopRow = lngRow
                            RaiseEvent AfterRowColChange(strInfo, True, mblnSign, mblnArchive)
                            CheckFlip = False
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next
    mblnShow = blnShow
    VsfData.Select lngOldRow, lngOldCol
    mrsCellMap.Filter = ""
    CheckFlip = True
End Function

Private Function CheckData() As Boolean
    Dim intLevel As Integer
    Dim lngPage As Long
    On Error GoTo ErrHand
    '检查数据
    
    '如果修改了数据而日期时间不全则提示（数据合法性在录入时已经检查）
'    Call OutputRsData(mrsCellMap)
'    Call OutputRsData(mrsDataMap)
    If Not DataMap_Save Then Exit Function
    
    '如果是审签模式,则检查所选数据是否存在不能审签的情况
    If mblnVerify Then
        mstrVerify = ""
        Set mobjVerify = New Collection
        mintVerify_Last = 未定义
        '审签不允许新增数据
        For lngPage = mint起始页码 To mint结束页
            mrsDataMap.Filter = "页号=" & lngPage
            Do While Not mrsDataMap.EOF
                If NVL(mrsDataMap!选择, 0) = flexTSChecked Then
                    mstrVerify = mstrVerify & "," & mrsDataMap!记录ID
                    mobjVerify.Add Format(mrsDataMap!发生时间, "YYYY-MM-DD HH:mm:ss"), "_" & mrsDataMap!记录ID
                    If IsNull(mrsDataMap!签名级别) Then
                        intLevel = NVL(mrsDataMap!签名级别, 未定义)
                    Else
                        intLevel = Val(mrsDataMap!签名级别) + 1
                    End If
                    If mintVerify < intLevel Then mintVerify_Last = intLevel
                End If
                mrsDataMap.MoveNext
            Loop
        Next
        mrsDataMap.Filter = 0
        
        If mstrVerify = "" Then
            RaiseEvent AfterRowColChange("至少要选择一条数据才能完成审签操作！", True, mblnSign, mblnArchive)
            Exit Function
        End If
        mstrVerify = Mid(mstrVerify, 2)
    End If
    
    CheckData = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function SaveData() As Boolean
    Dim arrValue, arrOrder, arrPart, arrCollect
    Dim strSQL() As String, strSQLTime() As String, strCollectSQL() As String
    Dim intAllow As Integer
    Dim lngRecord As Long
    Dim blnTrans As Boolean, blnSaved As Boolean, blnDel As Boolean
    Dim intPos As Integer, intMax As Integer, intPage As Integer, intRow As Integer, intStarRow As Integer, intUsedRows As Integer
    Dim strReturn As String, strCellData As String, strPart As String
    Dim strMonth As String, strDay As String
    Dim strDate As String, strTime As String, strTemp As String, strYear As String
    Dim strDatetime As String, strCurrDate As String, strDays As String, strLastDate As String, strActTime As String
    Dim rsTime As New ADODB.Recordset, rsTimeCur As New ADODB.Recordset '数据发生时间变动
    Dim strFileds As String, strValues As String
    Dim strRelationNO As String
    ReDim Preserve strSQL(1 To 1)
    ReDim Preserve strSQLTime(1 To 1) '发生时间变动SQL数组
    ReDim Preserve strCollectSQL(1 To 1) '小结数据SQL，小结数据放在最后在进行处理
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    
    strFileds = "ID," & adDouble & ",18|时间," & adDate & ",20|发生时间," & adDate & ",20|标记," & adInteger & ",1"
    Call Record_Init(rsTime, strFileds)
    Call Record_Init(rsTimeCur, strFileds)
    '同行多列循环调用：ZL_病人护理数据_UPDATE
    '下一行前调用：
    '   1、ZL_病人护理数据_SYNCHRO，同步数据到体温单与护理记录单中，需要记录删除的明细ID串
    '   2、ZL_病人护理打印_UPDATE，完成打印数据解析
    '删除项目需记录，删除行也需要记录
    '修改数据的同步就将该行数据对应的日期与时间保存到mrsCellMap中
    
'    objStream.WriteLine (Now & "产生保存SQL")
    intAllow = IIf(InStr(mstrPrivs, "他人护理记录") > 0, 1, 0)
    strCurrDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
    With mrsCellMap
        '将有效数据过滤出来:记录ID>0的历史数据+新增的有效数据
        .Filter = "记录ID>0 or (记录ID=0 And 删除=0)"
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            '77436:LPF,审签时只保存要审签的数据(开始勾选记录修改数据，可能后面有取消了勾选)
            If (InStr(1, "," & mstrVerify & ",", "," & NVL(!记录ID, 0) & ",") <> 0 And mblnVerify = True) Or mblnVerify = False Then
                'Or (intPage = !页号 And NVL(!起始行号, !行号) = intStarRow) 这不内容主要针对分类汇总，因为数据记录只有一条，而明细有多条
                If Not ((intRow = !行号 And intPage = !页号) Or (intPage = !页号 And NVL(!起始行号, !行号) = intStarRow)) Then
endWork:
                    If intRow > 0 Then
                        mrsDataMap.Filter = "页号=" & intPage & " And 行号=" & intRow
                        If mrsDataMap.RecordCount <> 0 Then
                            blnDel = (mrsDataMap!删除 = 1)
                            intUsedRows = Val(Split(NVL(mrsDataMap!行数 & "|"), "|")(0))
                        Else
                            mrsDataMap.Filter = 0
                            intUsedRows = 1
                            RaiseEvent AfterRowColChange("第" & intRow & "行的数据内部错误，请记录本次操作步骤并反馈，然后重新录入数据，谢谢！", True, mblnSign, mblnArchive)
                            Exit Function
                        End If
                        mrsDataMap.Filter = 0
                    End If
    
                    If blnSaved Then
                        '完成打印数据解析
    '                    文件ID_IN IN 病人护理打印.文件ID%TYPE,
    '                    发生时间_IN IN 病人护理打印.发生时间%TYPE,
    '                    行数_IN IN 病人护理打印.行数%TYPE,
    '                    删除_IN Number:=0
                        gstrSQL = "ZL_病人护理打印_UPDATE(" & mlng文件ID & ",to_date('" & strDatetime & "','yyyy-MM-dd hh24:mi:ss')," & intUsedRows & "," & IIf(blnDel, "1", "0") & ")"
                        strSQL(ReDimArray(strSQL)) = gstrSQL
                        
                        '只要修改过数据,必然会执行打印解析,因此在这里进行汇总日期的处理
                        If strDate <> "" And .EOF Then
                            strLastDate = strDate
                            
                            '同步更新明天的汇总(夜班,全天汇总跨天的处理)
                            If InStr(1, "," & strDays & ",", "," & Format(DateAdd("d", -1, strDate), "yyyy-MM-dd") & ",") = 0 Then
                                strDays = strDays & "," & Format(DateAdd("d", -1, strDate), "yyyy-MM-dd")
                                gstrSQL = "ZL_汇总数据_UPDATE(" & mlng文件ID & ",'" & Format(DateAdd("d", -1, strDate), "yyyy-MM-dd") & "')"
                                strCollectSQL(ReDimArray(strCollectSQL)) = gstrSQL
                            End If
                            
                            If InStr(1, "," & strDays & ",", "," & strDate & ",") = 0 Then
                                strDays = strDays & "," & strDate
                                gstrSQL = "ZL_汇总数据_UPDATE(" & mlng文件ID & ",'" & strDate & "')"
                                strCollectSQL(ReDimArray(strCollectSQL)) = gstrSQL
                            End If
                            
                            strTemp = Format(DateAdd("d", 1, CDate(strDate)), "yyyy-MM-dd")
                            If InStr(1, "," & strDays & ",", "," & strTemp & ",") = 0 Then
                                strDays = strDays & "," & strTemp
                                gstrSQL = "ZL_汇总数据_UPDATE(" & mlng文件ID & ",'" & strTemp & "')"
                                strCollectSQL(ReDimArray(strCollectSQL)) = gstrSQL
                            End If
                        End If
                        
                        blnSaved = False
                        If .EOF Then Exit Do
                    End If
                    
                    '赋初值
                    intPage = !页号
                    intRow = !行号
                    intStarRow = NVL(!起始行号, !行号)
                    strDate = ""
                    strDatetime = ""
                    lngRecord = NVL(!记录ID, 0)
                End If
                
                If !列号 = mlngDate Then
                    If NVL(!汇总, 0) = 1 Then
                        arrCollect = Split(!数据, ";")
                        strDatetime = arrCollect(3)
                    '    文件ID_IN IN 病人护理数据.文件ID%TYPE,
                    '    发生时间_IN IN 病人护理数据.发生时间%TYPE,
                    '    汇总类别_IN IN 病人护理数据.汇总类别%TYPE,
                    '    汇总文本_IN IN 病人护理数据.汇总文本%TYPE,
                    '    汇总标记_IN IN 病人护理数据.汇总标记%TYPE,
                    '    删除_IN Number:=0
                        gstrSQL = "ZL_病人护理数据_COLLECT(" & mlng文件ID & ",to_date('" & arrCollect(3) & "','yyyy-MM-dd hh24:mi:ss')," & _
                                Val(arrCollect(1)) & ",'" & arrCollect(0) & "'," & Val(arrCollect(2)) & ",'" & arrCollect(4) & "','" & arrCollect(5) & "'," & !删除 & ")"
                        strSQL(ReDimArray(strSQL)) = gstrSQL
                        blnSaved = True
                    Else
                        strDate = NVL(!数据)
                        If strDate <> "" Then
                            If mblnDateAd Then
                                '69335:刘鹏飞,2014-1-7,短日期格式处理(日期部位存放年份)
                                strYear = NVL(!部位)
                                If InStr(1, "|" & mstrYears & "|", "|" & strYear & "|") <> 0 Then
                                    strDate = strYear & "-" & ToStandDate(strDate)
                                Else
                                    RaiseEvent AfterRowColChange("第" & !行号 & "行的[年份]数据错误，请记录本次操作步骤并反馈，然后重新录入数据，谢谢！", True, mblnSign, mblnArchive)
                                    Exit Function
                                End If
    '                            strDate = Mid(strCurrDate, 1, 5) & ToStandDate(strDate)
    '                            '检查是否翻年后编辑之前的时间(一个月的限制)
    '                            If CDate(strDate) > CDate(Mid(strCurrDate, 1, 10)) And Mid(strCurrDate, 6, 2) = "01" And Mid(strDate, 6, 2) = "12" Then
    '                                strDate = DateAdd("yyyy", -1, CDate(strDate))
    '                            End If
                            Else
                                strDate = Format(strDate, "yyyy-MM-dd")
                            End If
                        End If
                        
                        '已有数据修改时间的处理
                        If lngRecord <> 0 And strDate <> "" Then
                            mrsDataMap.Filter = "页号=" & !页号 & " And 行号=" & !行号
                            If mrsDataMap.RecordCount > 0 Then
                                strActTime = Format(mrsDataMap.Fields(cControlFields + mlngActTime - VsfData.FixedCols).Value, "YYYY-MM-DD")
                                If Format(strActTime, "YYYY-MM-DD") <> Format(strDate, "YYYY-MM-DD") Then
                                    '必须同时处理昨天:因为可能存在跨天汇总的数据，且当前时间刚好在第二天的情况
                                    If InStr(1, "," & strDays & ",", "," & Format(DateAdd("d", -1, strActTime), "yyyy-MM-dd") & ",") = 0 Then
                                        strDays = strDays & "," & Format(DateAdd("d", -1, strActTime), "yyyy-MM-dd")
                                        gstrSQL = "ZL_汇总数据_UPDATE(" & mlng文件ID & ",'" & Format(DateAdd("d", -1, strActTime), "yyyy-MM-dd") & "')"
                                        strCollectSQL(ReDimArray(strCollectSQL)) = gstrSQL
                                    End If
                                    '今天
                                    If InStr(1, "," & strDays & ",", "," & strActTime & ",") = 0 Then
                                        strDays = strDays & "," & strActTime
                                        gstrSQL = "ZL_汇总数据_UPDATE(" & mlng文件ID & ",'" & strActTime & "')"
                                        strCollectSQL(ReDimArray(strCollectSQL)) = gstrSQL
                                    End If
                                End If
                            End If
                        End If
                    End If
                    If strLastDate = "" Then strLastDate = strDate
                    
                    If strLastDate <> strDate Then
                        '只要修改过数据,必然会执行打印解析,因此在这里进行汇总日期的处理
                        '同步更新明天的汇总(夜班,全天汇总跨天的处理)
                        If InStr(1, "," & strDays & ",", "," & Format(DateAdd("d", -1, strLastDate), "yyyy-MM-dd") & ",") = 0 Then
                            strDays = strDays & "," & Format(DateAdd("d", -1, strLastDate), "yyyy-MM-dd")
                            gstrSQL = "ZL_汇总数据_UPDATE(" & mlng文件ID & ",'" & Format(DateAdd("d", -1, strLastDate), "yyyy-MM-dd") & "')"
                            strCollectSQL(ReDimArray(strCollectSQL)) = gstrSQL
                        End If
                        
                        If InStr(1, "," & strDays & ",", "," & strLastDate & ",") = 0 Then
                            strDays = strDays & "," & strLastDate
                            gstrSQL = "ZL_汇总数据_UPDATE(" & mlng文件ID & ",'" & strLastDate & "')"
                            strCollectSQL(ReDimArray(strCollectSQL)) = gstrSQL
                        End If
                        
                        strTemp = Format(DateAdd("d", 1, CDate(strLastDate)), "yyyy-MM-dd")
                        If InStr(1, "," & strDays & ",", "," & strTemp & ",") = 0 Then
                            strDays = strDays & "," & strTemp
                            gstrSQL = "ZL_汇总数据_UPDATE(" & mlng文件ID & ",'" & strTemp & "')"
                            strCollectSQL(ReDimArray(strCollectSQL)) = gstrSQL
                        End If
                        strLastDate = strDate
                    End If
                ElseIf !列号 = mlngTime Then
                    strTime = NVL(!数据)
                    If strDate = "" Then strDate = Mid(strCurrDate, 1, 10)
                    strDatetime = strDate & " " & strTime & ":00"
                    
                    '处理分组数据，保存时与普通数据无区别，只是秒数+
                    If Val(NVL(!部位)) >= 1 Then
                        'strDatetime = Mid(strDatetime, 1, 17) & String(2 - Len(!部位), "0") & Val(!部位) - 1
                        strDatetime = DateAdd("S", Val(!部位) - 1, CDate(strDatetime))
                    End If
                    
                    If lngRecord <> 0 Then
                        '更新发生时间
    '                    gstrSQL = "Zl_病人护理数据_发生时间(" & lngRecord & ",to_date('" & strDatetime & "','yyyy-MM-dd hh24:mi:ss'))"
    '                    strSQLTime(ReDimArray(strSQLTime)) = gstrSQL
                        mrsDataMap.Filter = "页号=" & !页号 & " And 行号=" & !行号
                        If mrsDataMap.RecordCount = 0 Then
                            RaiseEvent AfterRowColChange("第" & !行号 & "行的数据内部错误，请记录本次操作步骤并反馈，然后重新录入数据，谢谢！", True, mblnSign, mblnArchive)
                            Exit Function
                        End If
                        strActTime = Format(mrsDataMap.Fields(cControlFields + mlngActTime - VsfData.FixedCols).Value, "YYYY-MM-DD HH:mm:ss")
                        strValues = lngRecord & "|" & Format(strDatetime, "YYYY-MM-DD HH:mm:ss") & "|" & Format(strActTime, "YYYY-MM-DD HH:mm:ss") & "|0"
                        Call Record_Update(rsTime, "ID|时间|发生时间|标记", strValues, "ID|" & lngRecord)
                        Call Record_Update(rsTimeCur, "ID|时间|发生时间|标记", strValues, "ID|" & lngRecord)
                        blnSaved = True
                    End If
                Else
                    If !列号 > mlngTime Then
                        '取指定单元格的数据
                        strCellData = NVL(!数据)
                        strPart = NVL(!部位)
                        strReturn = ShowInput(!列号, strCellData, True)
                        'strOrders格式：项目序号,项目序号...
                        'strValues格式：值'值'值...
                        arrOrder = Split(Split(strReturn, "||")(0), ",")
                        arrValue = Split(Split(strReturn, "||")(1) & "'", "'")
                        arrPart = Split(strPart & "/////", "/")
                        
                        intMax = UBound(arrOrder)
                        For intPos = 0 To intMax
        '                    文件ID_IN IN 病人护理数据.文件ID%TYPE,
        '                    发生时间_IN IN 病人护理数据.发生时间%TYPE,
        '                    记录类型_IN IN 病人护理明细.记录类型%TYPE,          --护理项目=1，上标说明=2，手术日标记=4，签名记录=5，下标说明=6，入出量汇总=9
        '                    项目序号_IN IN 病人护理明细.项目序号%TYPE,          --护理项目的序号，非护理项目固定为0
        '                    记录内容_IN IN 病人护理明细.记录内容%TYPE := NULL,  --记录内容，如果内容为空，即清除以前的内容；37或38/37
        '                    体温部位_IN IN 病人护理明细.体温部位%TYPE := NULL,
        '                    他人记录_IN IN NUMBER := 1,
                            '65258:刘鹏飞,2013-11-1,小结为空也要显示(小结汇总项目强制插入Chr(13))
                            If NVL(!汇总, 0) = 1 And arrValue(intPos) = "" And InStr(1, "|" & mstrColCollect, "|" & Val(NVL(!列号, 0)) - cHideCols & ";") > 0 Then
                                arrValue(intPos) = Chr(13)
                            End If
                            '分类汇总根据汇总项目的行号和序号获取对应的关联项目序号
                            strRelationNO = ""
                            If NVL(!汇总, 0) = 1 And Val(NVL(mrsCellMap!记录组号)) > 0 Then
                                strRelationNO = GetRelatiionNo(Val(NVL(!列号, 0)) - cHideCols & "," & arrOrder(intPos))
                            End If
                            gstrSQL = "ZL_病人护理数据_UPDATE(" & mlng文件ID & ",to_date('" & strDatetime & "','yyyy-MM-dd hh24:mi:ss'),1," & _
                                    arrOrder(intPos) & ",'" & arrValue(intPos) & "','" & arrPart(intPos) & "'," & intAllow & ",0," & _
                                    IIf(mblnVerify, 1, 0) & ",NULL," & IIf(IsNull(mrsCellMap!记录组号), IIf(NVL(!汇总, 0) = 1, 0, "NULL"), Val(NVL(mrsCellMap!记录组号)))
                            If strRelationNO = "" Then
                                gstrSQL = gstrSQL & ",NULL,'" & NVL(!标记) & "')"
                            Else
                                gstrSQL = gstrSQL & "," & Val(strRelationNO) & ",'" & NVL(!标记) & "')"
                            End If
                            strSQL(ReDimArray(strSQL)) = gstrSQL
                            blnSaved = True
                        Next
                        mrsItems.Filter = 0
                    End If
                End If
            End If
            .MoveNext
        Loop
        
        If blnSaved Then GoTo endWork
        mrsDataMap.Filter = 0
    End With
    
    '更新数据发生时间，对于分组数据中间某行数据行数变化会引起后面本分组数据的其他分组数据时间发生变化。如(增加数据行)：
    ',ID:403,时间:2012/5/8 18:23:00,发生时间:2012/5/8 18:23:00
    ',ID:407,时间:2012/5/8 18:23:02,发生时间:2012/5/8 18:23:01
    ',ID:517,时间:2012/5/8 18:23:03,发生时间:2012/5/8 18:23:02
    '需要先更新最后一行发生时间：如(减少数据行):
    ',ID:403,时间:2012/5/8 18:23:00,发生时间:2012/5/8 18:23:00
    ',ID:407,时间:2012/5/8 18:23:01,发生时间:2012/5/8 18:23:02
    ',ID:517,时间:2012/5/8 18:23:02,发生时间:2012/5/8 18:23:03
    '需要从前往后更新
    strDays = ""
    rsTime.Filter = ""
    'Call OutputRsData(rsTime)
    rsTime.Sort = "时间 DESC"
    Do While Not rsTime.EOF
        If InStr(1, "," & strDays & ",", "," & rsTime!ID & ",") = 0 Then
            rsTimeCur.Filter = "发生时间='" & Format(rsTime!时间, "YYYY-MM-DD HH:mm:ss") & "'And 标记=0 And ID<>" & Val(rsTime!ID)
            If rsTimeCur.RecordCount > 0 Then
                lngRecord = rsTimeCur!ID
                gstrSQL = UpdateTime(rsTimeCur, Format(rsTimeCur!时间, "YYYY-MM-DD HH:mm:ss"), lngRecord)
                strSQLTime(ReDimArray(strSQLTime)) = gstrSQL
                rsTimeCur.Filter = ""
                Call Record_Update(rsTimeCur, "标记", "1", "ID|" & lngRecord)
                strDays = IIf(strDays = "", "", ",") & lngRecord
                GoTo ErrLoop
            Else
                lngRecord = rsTime!ID
                gstrSQL = "Zl_病人护理数据_发生时间(" & rsTime!ID & ",to_date('" & rsTime!时间 & "','yyyy-MM-dd hh24:mi:ss'))"
                strSQLTime(ReDimArray(strSQLTime)) = gstrSQL
                rsTimeCur.Filter = ""
                Call Record_Update(rsTimeCur, "标记", "1", "ID|" & lngRecord)
                strDays = IIf(strDays = "", "", ",") & lngRecord
            End If
        End If
    rsTime.MoveNext
ErrLoop:
    Loop
    
    '循环执行SQL保存数据
    'On Error Resume Next
    
    gcnOracle.BeginTrans
    blnTrans = True
    
    On Error GoTo ErrHand
    '先更新发生时间
    intMax = UBound(strSQLTime)
    If intMax > 0 Then
        For intPos = 1 To intMax
            If strSQLTime(intPos) <> "" Then
                'Debug.Print strSQLTime(intPos)
    '            objStream.WriteLine (Now & "；SQL：" & strSQLTime(intPos))
                Call zlDatabase.ExecuteProcedure(strSQLTime(intPos), "保存护理记录单数据")
            End If
        Next intPos
    End If
    '在更新数据
    intMax = UBound(strSQL)
    If intMax > 0 Then
'        objStream.WriteLine (Now & "准备保存数据")
        For intPos = 1 To intMax
            If strSQL(intPos) <> "" Then
                'Debug.Print strSQL(intPos)
    '            objStream.WriteLine (Now & "；SQL：" & strSQL(intPos))
                Call zlDatabase.ExecuteProcedure(strSQL(intPos), "保存护理记录单数据")
            End If
        Next
    '    objStream.WriteLine (Now & "保存数据完成")
    End If
    '最后更新小结内容
    intMax = UBound(strCollectSQL)
    If intMax > 0 Then
        For intPos = 1 To intMax
            If strCollectSQL(intPos) <> "" Then
                'Debug.Print strCollectSQL(intPos)
    '            objStream.WriteLine (Now & "；SQL：" & strCollectSQL(intPos))
                Call zlDatabase.ExecuteProcedure(strCollectSQL(intPos), "保存护理记录单数据")
            End If
        Next intPos
    End If
    
    If mblnVerify Then
        If Not SignMe(mblnVerify) Then
            gcnOracle.RollbackTrans
            Exit Function
        End If
    End If
    
    gcnOracle.CommitTrans
    SaveData = True
    blnTrans = False
    mblnChange = False
    mblnVerify = False
    mstrVerify = ""
    Set mobjVerify = Nothing
    
    RaiseEvent AfterDataChanged(mblnChange Or mblnVerify)
    RaiseEvent AfterRefresh
    Exit Function
ErrHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function UpdateTime(rsTimeCur As ADODB.Recordset, ByVal strTime As String, lngID As Long) As String
    Dim strSQL As String
    rsTimeCur.Filter = "发生时间='" & Format(strTime, "YYYY-MM-DD HH:mm:ss") & "' And 标记=0 And ID<>" & lngID
    If rsTimeCur.RecordCount > 0 Then
        lngID = Val(rsTimeCur!ID)
        strSQL = UpdateTime(rsTimeCur, Format(rsTimeCur!时间, "YYYY-MM-DD HH:mm:ss"), lngID)
    Else
        strSQL = "Zl_病人护理数据_发生时间(" & lngID & ",to_date('" & strTime & "','yyyy-MM-dd hh24:mi:ss'))"
    End If
    UpdateTime = strSQL
End Function

Private Sub cboChoose_GotFocus(Index As Integer)
    mblnEditAssistant = False
    mblnEditText = False
End Sub

Private Sub cboChoose_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Index = 0 Then
            cboChoose(1).SetFocus
        Else
            Call MoveNextCell
        End If
    End If
End Sub

Private Sub cboYear_GotFocus()
    mblnEditAssistant = False
    mblnEditText = False
End Sub

Private Sub cboYear_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call MoveNextCell
    End If
End Sub

Private Sub cbo小结_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txt开始时点.Enabled Then
            txt开始时点.SetFocus
        Else
            txt小结名称.SetFocus
        End If
    End If
End Sub


Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strDate As String, strTime As String, strYear As String
    Dim strLockItem As String                   '同步过来的数据,不允许修改或删除
    Dim lngTop As Long, lngHeight As Long
    Dim intMax As Integer                       '同步过来的数据占用的最大行数
    Dim intNULL As Integer, lngStartRow As Long, lngRowCount As Long, blnNULL As Boolean
    Dim lngRow As Long, lngCol As Long, lngRows As Long, lngCols As Long
    Dim strKey As String, strField As String, strValue As String
    Dim strPart As String, strCols As String
    Dim lngOrder As Long, intGroupFirstRows As Integer
    Dim lngCol1 As Long, lngRow1 As Long, lngCurRow As Long, strText As String, lngCount As Long
    Dim blnTure As Boolean
    Dim varAssistant() As Variant, strAssistantCols As String
    On Error GoTo err_exit
    
    Select Case Control.ID
    '粘贴,清除时需要同步mrsCellMap数据
    Case conMenu_Edit_Group_New
        '添加分组（当前记录的组号为1，但分组与普通数据没有区别，只是秒数上有变化；分组数据必须连续录入，不支持修改为分组或将分组数据修改为普通数据的功能）
        Control.Category = ""
        mblnGroupNew = mblnGroupNew Xor True
        If mblnGroupNew Then
            '记录起始行，非起始行不允许录入日期与时间
            Control.Category = VsfData.ROW
        End If
        Control.Checked = mblnGroupNew
        mstrGroupRow = Control.Category
    Case conMenu_Edit_Group_Append
        If VsfData.ROW < VsfData.FixedRows Then Exit Sub
        If InStr(1, FormatValue(VsfData.TextMatrix(VsfData.ROW, mlngRowCount)), "|") = 0 Then Exit Sub
        If Val(VsfData.TextMatrix(lngRow, mlngCollectType)) <> 0 Then Exit Sub
        
        '隐蔽已显示的录入控件
        Select Case mintType
        Case 0, 3
            picInput.Visible = False
        Case 1, 2
            lstSelect(mintType - 1).Visible = False
            If mintType = 1 Then
                txtLst.Visible = False
                PicLst.Visible = False
            End If
        Case 4, 5
            picDouble.Visible = False
        Case 6
            picMutilInput.Visible = False
        Case 7
            picDoubleChoose.Visible = False
        Case 8
            picYear.Visible = False
        End Select
        cmdWord.Visible = False
        mintType = -1
        
        If FormatValue(VsfData.TextMatrix(VsfData.ROW, mlngRowCount)) <> "1|1" Then
            lngRowCount = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0))
            lngStartRow = GetStartRow(VsfData.ROW)
        Else
            lngRowCount = 1
            lngStartRow = VsfData.ROW
        End If
        
        '追加数据时，需重新计算选中数据的行数，计算是不能包含大文本段信息。
        '如：一行数据5行，最大非大文本的内容只有3行，选中改行追加数据时，应该追加到第4行，demo=1的为3行，demo=4的为2行
        intNULL = lngStartRow + lngRowCount - 1
        For lngRow = lngRowCount To 1 Step -1
            blnNULL = True
            For lngCol = 0 To VsfData.Cols - 1
                If Not VsfData.ColHidden(lngCol) And lngCol < mlngNoEditor And ISEditAssistant(lngCol) = False Then
                    If FormatValue(VsfData.TextMatrix(lngRow + lngStartRow - 1, lngCol)) <> "" And Not (IsDiagonal(lngCol) And InStr(1, FormatValue(VsfData.TextMatrix(lngRow + lngStartRow - 1, lngCol)), "/") <> 0) Then
                        blnNULL = False
                        Exit For
                    End If
                End If
            Next
            
            If Not blnNULL Then Exit For
            intNULL = intNULL - 1
        Next
        '从新填写行序号
        If intNULL < lngStartRow Then intNULL = lngStartRow
        For lngRow = lngStartRow To intNULL
            VsfData.TextMatrix(lngRow, mlngRowCount) = (intNULL - lngStartRow + 1) & "|" & lngRow - lngStartRow + 1
            VsfData.TextMatrix(lngRow, mlngRowCurrent) = (intNULL - lngStartRow + 1)
        Next
        
        If mlngSignName <> -1 Then
            If Trim(FormatValue(VsfData.TextMatrix(lngStartRow + lngRowCount - 1, mlngSignName))) <> "" Then
                VsfData.TextMatrix(intNULL, mlngSignName) = FormatValue(VsfData.TextMatrix(lngStartRow + lngRowCount - 1, mlngSignName))
                If mlngSignTime <> -1 Then VsfData.TextMatrix(intNULL, mlngSignTime) = FormatValue(VsfData.TextMatrix(lngStartRow + lngRowCount - 1, mlngSignTime))
            End If
        End If
        
        If intNULL + 1 <= lngStartRow + lngRowCount - 1 Then
            For lngRow = intNULL + 1 To lngStartRow + lngRowCount - 1
                '清空隐藏列数据
                For lngCol = 0 To VsfData.Cols - 1
                    If VsfData.ColHidden(lngCol) = True Then VsfData.TextMatrix(lngRow, lngCol) = ""
                Next lngCol
                VsfData.TextMatrix(lngRow, mlngRowCount) = (lngStartRow + lngRowCount - intNULL - 1) & "|" & (lngRow - intNULL)
                VsfData.TextMatrix(lngRow, mlngRowCurrent) = (lngStartRow + lngRowCount - intNULL - 1)
            Next
            lngRowCount = Val(Split(FormatValue(VsfData.TextMatrix(lngStartRow, mlngRowCount)), "|")(0))
            '更新本列大文本列数据集信息
            Call CellMap_UpdateAssistant(lngStartRow)
            blnTure = False
        Else
            '检查下一行数据是否为空,如果不是空行直接添加到下一行
            lngCurRow = lngStartRow + lngRowCount
            blnTure = False
            If lngCurRow >= VsfData.Rows Then
                blnTure = True
            Else
                If Not VsfData.RowHidden(lngCurRow) Then
                    For lngCol = 0 To VsfData.Cols - 1
                        If Not VsfData.ColHidden(lngCol) And lngCol < mlngNoEditor Then
                            If FormatValue(VsfData.TextMatrix(lngCurRow, lngCol)) <> "" And Not (IsDiagonal(lngCol) And InStr(1, FormatValue(VsfData.TextMatrix(lngCurRow, lngCol)), "/") <> 0) Then
                                blnTure = True
                                Exit For
                            End If
                        End If
                    Next
                Else
                    blnTure = True
                End If
            End If
        End If
        If blnTure = True Then
            '追加分组，在当前行(只有一行的数据行)后增加空白行
            '1)先增加1个空行
            VsfData.Rows = VsfData.Rows + 1
            '   从当前行记录的空白行开始，每行的位置+所增加的空白行数
            For lngRow = VsfData.Rows - 2 To lngStartRow + lngRowCount Step -1
                VsfData.RowPosition(lngRow) = lngRow + 1
            Next
            lngRow = lngStartRow + lngRowCount - 1
            '2)当行号发生变化后，需同步更新mrsCellMap中大于该行号的行号数据
            Call CellMap_Update(lngRow, 1)
        End If
        '3)更新分组相关控制
        mintType = -1: mblnShow = False
        Call AppendGroup(lngStartRow)
        lngRow1 = VsfData.ROW
        lngCol1 = VsfData.COL
        If InStr(1, FormatValue(VsfData.TextMatrix(lngRow1, mlngRowCount)), "|") = 0 Then VsfData.TextMatrix(lngRow1, mlngRowCount) = "1|1"
        intGroupFirstRows = Val(Split(FormatValue(VsfData.TextMatrix(lngRow1, mlngRowCount)), "|")(0))
        '下一行如果存在数据则取消分组选择
        blnTure = False
        lngCurRow = lngRow1 + intGroupFirstRows
        If lngCurRow >= VsfData.Rows Then
            blnTure = True
        Else
            If Not VsfData.RowHidden(lngCurRow) Then
                For lngCol = 0 To VsfData.Cols - 1
                    If Not VsfData.ColHidden(lngCol) And lngCol < mlngNoEditor Then
                        If FormatValue(VsfData.TextMatrix(lngCurRow, lngCol)) <> "" And Not (IsDiagonal(lngCol) And InStr(1, FormatValue(VsfData.TextMatrix(lngCurRow, lngCol)), "/") <> 0) Then
                            blnTure = True
                            Exit For
                        End If
                    End If
                Next
            Else
                blnTure = True
            End If
        End If
        mblnGroupApp = blnTure
        '在原有分组数据上分组,需要处理分组序号
        If Val(FormatValue(VsfData.TextMatrix(lngStartRow, mlngDemo))) > 0 Then '分组数据行
            '确定分组起始行
            lngRow = lngStartRow
            lngStartRow = lngRow - Val(FormatValue(VsfData.TextMatrix(lngRow, mlngDemo))) + 1
            If Val(FormatValue(VsfData.TextMatrix(lngStartRow, mlngDemo))) <> 1 Then
                For lngStartRow = lngRow To VsfData.FixedRows Step -1
                    If Val(FormatValue(VsfData.TextMatrix(lngStartRow, mlngDemo))) = 1 Then
                        lngRow = lngStartRow
                        Exit For
                    End If
                Next lngStartRow
                If lngStartRow < VsfData.FixedRows Then GoTo ErrNext
                lngStartRow = lngRow
            End If
            '重新组织分组序号
            intGroupFirstRows = Val(Split(FormatValue(VsfData.TextMatrix(lngRow1, mlngRowCount)), "|")(0))
            lngCurRow = lngRow1
            For lngRow = lngRow1 + intGroupFirstRows To VsfData.Rows - 1
                If lngRow = lngCurRow + intGroupFirstRows Then
                    If Val(FormatValue(VsfData.TextMatrix(lngRow, mlngDemo))) <= 1 Then
                        Exit For
                    Else
                        VsfData.TextMatrix(lngRow, mlngDemo) = lngRow - Val(lngStartRow) + 1
                    End If
                    If InStr(1, FormatValue(VsfData.TextMatrix(lngRow, mlngRowCount)), "|") = 0 Then VsfData.TextMatrix(lngRow, mlngRowCount) = "1|1"
                    intGroupFirstRows = Val(Split(FormatValue(VsfData.TextMatrix(lngRow, mlngRowCount)), "|")(0))
                    lngCurRow = lngRow
                End If
            Next
            blnTure = False
            mblnEditAssistant = False
            For lngCol = mlngTime + 1 To mlngNoEditor - 1
                '寻找大文本列
                mrsSelItems.Filter = "列=" & lngCol - cHideCols
                If mrsSelItems.RecordCount > 0 Then
                    lngOrder = Val(mrsSelItems!项目序号)
                    mrsItems.Filter = "项目序号=" & lngOrder
                    If mrsItems.RecordCount = 0 Then
                        mrsItems.Filter = 0
                        GoTo ErrNext
                    End If
                    mblnEditAssistant = (mrsItems!项目类型 = 1 And mrsItems!项目长度 > 100) And Val(FormatValue(VsfData.TextMatrix(lngStartRow, mlngDemo))) <= 1
                    If Not mblnEditAssistant Then GoTo ErrNext
                        
                    If InStr(1, FormatValue(VsfData.TextMatrix(lngStartRow, mlngRowCount)), "|") = 0 Then VsfData.TextMatrix(lngStartRow, mlngRowCount) = "1|1"
                    intGroupFirstRows = Val(Split(FormatValue(VsfData.TextMatrix(lngStartRow, mlngRowCount)), "|")(0))
                    '为分组行时，选择数据起始行，编辑内容显示所有大文本行
                    strText = ""
                    If Val(FormatValue(VsfData.TextMatrix(lngStartRow, mlngDemo))) = 1 Then
                        For lngRow = 0 To intGroupFirstRows - 1
                            strText = strText & IIf(strText = "", "", vbCrLf) & Replace(Replace(Replace(VsfData.TextMatrix(lngRow + lngStartRow, lngCol), Chr(13), ""), Chr(10), ""), Chr(1), "")
                        Next lngRow
                        lngCount = lngStartRow + intGroupFirstRows - 1
                        For lngRow = lngStartRow + intGroupFirstRows To VsfData.Rows - 1
                            If VsfData.RowHidden(lngRow) = False Then
                                 '分组中每一个分组子行数据可能占用多行，只有在第一行保存了分组索引。分组总数据行数=每一个分组子行数+每一个分组子行的行数
                                If lngRow > lngCount Then
                                    If Val(FormatValue(VsfData.TextMatrix(lngRow, mlngDemo))) <= 1 Then Exit For
                                    If VsfData.TextMatrix(lngRow, mlngRowCount) = "" Then VsfData.TextMatrix(lngRow, mlngRowCount) = "1|1"
                                    lngCount = Val(Split(FormatValue(VsfData.TextMatrix(lngRow, mlngRowCount)), "|")(0)) + lngRow - 1
                                End If
                                    
                                strText = strText & IIf(strText = "", "", vbCrLf) & Replace(Replace(Replace(VsfData.TextMatrix(lngRow, lngCol), Chr(13), ""), Chr(10), ""), Chr(1), "")
                            Else
                                lngCount = lngCount + 1
                            End If
                        Next lngRow
                        mintType = -1: mblnShow = False
                        If strText = "" Then GoTo ErrNext ' strText = " "
                        VsfData.ROW = lngStartRow
                        VsfData.COL = lngCol
                        mintType = 0
                        'lngRow1 要追加的行
                        Call MoveNextCell(False, True, strText, lngRow1)
                        mintType = -1
                        blnTure = True
                    End If
ErrNext:
                End If
            Next lngCol
            'blnTrue=false 说明记录单没有大行文本(分组数据在起始行点击追加，在追加行录入数据，只能保存追加行和邻近下一行其中的一条数据)
            If blnTure = False Then
                '从起始行开始处理分组数据(防止已经保存的分组数据分组行和保存的时间不对应，导致存在两条相同时间的数据)
                '如：保存的数据起始行Demo=1，发生时间秒数为=01，此时追加一条新记录，Demo=2 保存时秒数也为01(如果修改了起始行数据就不存在这种情况)
                intGroupFirstRows = 0
                lngCurRow = lngStartRow
                For lngRow = lngStartRow To VsfData.Rows - 1
                    If lngRow = lngCurRow + intGroupFirstRows Then
                        If Val(FormatValue(VsfData.TextMatrix(lngRow, mlngDemo))) <= 1 And intGroupFirstRows > 0 Then Exit For
                        If InStr(1, FormatValue(VsfData.TextMatrix(lngRow, mlngRowCount)), "|") = 0 Then VsfData.TextMatrix(lngRow, mlngRowCount) = "1|1"
                        intGroupFirstRows = Val(Split(FormatValue(VsfData.TextMatrix(lngRow, mlngRowCount)), "|")(0))
                        lngCurRow = lngRow
                        strYear = ""
                        If CheckGroupDate(lngRow) = True Then
                            '保存后的修改才进入此流程，取该条记录的实际时间
                            If mblnDateAd Then
                                strYear = Format(FormatValue(VsfData.TextMatrix(lngRow, mlngActTime)), "YYYY")
                                strDate = Format(FormatValue(VsfData.TextMatrix(lngRow, mlngActTime)), "DD") & "/" & Format(FormatValue(VsfData.TextMatrix(lngRow, mlngActTime)), "MM")
                            Else
                                strDate = Mid(FormatValue(VsfData.TextMatrix(lngRow, mlngActTime)), 1, 10)
                            End If
                            strTime = Mid(FormatValue(VsfData.TextMatrix(lngRow, mlngActTime)), 12, 5)
                        Else
                            '新增时进入此流程
                            strDate = VsfData.TextMatrix(lngRow - Val(FormatValue(VsfData.TextMatrix(lngRow, mlngDemo))) + 1, mlngDate)
                            strTime = VsfData.TextMatrix(lngRow - Val(FormatValue(VsfData.TextMatrix(lngRow, mlngDemo))) + 1, mlngTime)
                            If mblnDateAd Then strYear = VsfData.TextMatrix(lngRow - Val(FormatValue(VsfData.TextMatrix(lngRow, mlngDemo))) + 1, mlngYear)
                        End If
                        
                        '1\日期
                        strField = "ID|页号|行号|列号|记录ID|数据|部位|删除"
                        If mlngDate <> -1 Then
                            strKey = mint页码 & "," & lngRow & "," & mlngDate
                            strValue = strKey & "|" & mint页码 & "|" & lngRow & "|" & mlngDate & "|" & _
                                Val(VsfData.TextMatrix(lngRow, mlngRecord)) & "|" & strDate & "|" & strYear & "|0"
                            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                        End If
                        '2\时间
                        strKey = mint页码 & "," & lngRow & "," & mlngTime
                        strValue = strKey & "|" & mint页码 & "|" & lngRow & "|" & mlngTime & "|" & _
                            Val(VsfData.TextMatrix(lngRow, mlngRecord)) & "|" & strTime & "|" & _
                            VsfData.TextMatrix(lngRow, mlngDemo) & "|0"
                        Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                    End If
                Next lngRow
            Else
                '还原选择列
                mblnShow = False
                mintType = -1
                VsfData.ROW = lngRow1
                VsfData.COL = lngCol1
            End If
        End If
        If InStr(1, VsfData.TextMatrix(lngRow1, mlngRowCount), "|") = 0 Then VsfData.TextMatrix(lngRow1, mlngRowCount) = "1|1"
    Case conMenu_Edit_Copy
        '复制指定数据行的数据
        If VsfData.ROW < VsfData.FixedRows Then Exit Sub
        If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then Exit Sub
        lngRow = GetStartRow(VsfData.ROW)
        If Val(VsfData.TextMatrix(lngRow, mlngCollectType)) <> 0 Then Exit Sub
        
        '复制记录集
        Set mrsCopyMap = New ADODB.Recordset
        Set mrsCopyMap = CopyNewRec(mrsDataMap, False)
        
        '得到指定数据行的起始行,结束行
        lngCols = VsfData.Cols - 1
        lngRows = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0))
        lngRows = lngRow + lngRows - 1
        For lngRow = lngRow To lngRows
            mrsCopyMap.AddNew
            mrsCopyMap!页号 = mint页码
            mrsCopyMap!行号 = lngRow
            For lngCol = 0 To lngCols - VsfData.FixedCols    '多了一个固定列
                mrsCopyMap.Fields(cControlFields + lngCol).Value = IIf(FormatValue(VsfData.TextMatrix(lngRow, lngCol + VsfData.FixedCols)) = "", Null, FormatValue(VsfData.TextMatrix(lngRow, lngCol + VsfData.FixedCols)))
            Next
            mrsCopyMap.Update
        Next
    Case conMenu_Edit_PASTE
        '粘贴时，将目标行整体覆盖，同步过来的数据列，活动列除外
        '活动项目可能不同页面项目不同，部位不同，所以不考虑活动项目
        '同步行所占用的行数不变，如不够再添加空白行，再行粘贴
        If VsfData.ROW < VsfData.FixedRows Then Exit Sub
        If mrsCopyMap.RecordCount = 0 Then Exit Sub
        
        '跨页数据行不允许对整行进行粘贴,删除,只能编辑除活动项目外的列
        If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then
            lngStartRow = VsfData.ROW
        Else
            lngStartRow = GetStartRow(VsfData.ROW)
        End If
        If Val(VsfData.TextMatrix(lngStartRow, mlngCollectType)) <> 0 Then Exit Sub
        If Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) >= 1 Then Exit Sub
        
        If InStr(1, VsfData.TextMatrix(lngStartRow, mlngRowCount), "|") <> 0 And lngStartRow = 3 And Val(VsfData.TextMatrix(lngStartRow, mlngStartRowPage)) <> mint页码 Then
            If Val(Split(VsfData.TextMatrix(lngStartRow, mlngRowCount), "|")(0)) <> Val(VsfData.TextMatrix(lngStartRow, mlngRowCurrent)) Then
                RaiseEvent AfterRowColChange("跨页数据行不允许粘贴，请切换到上一页进行操作！", True, mblnSign, mblnArchive)
                Exit Sub
            End If
        End If
        
        If VsfData.TextMatrix(lngStartRow, mlngSigner) <> "" Then
            RaiseEvent AfterRowColChange("已签名的数据不允许粘贴，请取消签名后再试！", True, mblnSign, mblnArchive)
            Exit Sub
        End If
        
        '对于已经存在数据，如果汇总数据已经签名不能粘贴
        blnTure = False
        If Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) > 0 And mstrColCollect <> "" Then
            '找出数据不为空的汇总列
            For lngRow = 0 To UBound(Split(mstrColCollect, "|"))
                strValue = GetRelatiionNo(CStr(Split(mstrColCollect, "|")(lngRow)), 2)
                strCols = strCols & "," & IIf(strValue = "", "", strValue & ",") & Split(Split(mstrColCollect, "|")(lngRow), ";")(0)
            Next
            strCols = Mid(strCols, 2)
            If strCols <> "" Then
                If ISCollectSigned(mlng文件ID, Format(VsfData.TextMatrix(lngStartRow, mlngActTime), "YYYY-MM-DD"), Format(VsfData.TextMatrix(lngStartRow, mlngActTime), "HH:MM")) Then
                    blnTure = True
                    If MsgBox("您要修改的数据所对应的汇总数据已签名，复制数据中所包含的汇总列数据将不能被粘贴，请问您是否继续。", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Sub
                End If
            End If
        End If
        
        '检查目标数据行是否存在同步过来的数据,如果有则跳过同步的记录
        strLockItem = GetSynItems(2, intMax)        '1.返回项目序号;2.返回列号
        
        '得到目标数据行的起始行,结束行
        strField = "ID|页号|行号|列号|记录ID|数据|部位|删除"
        lngCols = VsfData.Cols - 1
        strYear = ""
        If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then
            lngRow = VsfData.ROW
            If Val(VsfData.TextMatrix(lngRow, mlngCollectType)) <> 0 Then Exit Sub
            lngStartRow = lngRow
            If mlngDate > -1 Then strDate = VsfData.TextMatrix(lngRow, mlngDate)
            strTime = VsfData.TextMatrix(lngRow, mlngTime)
            strYear = VsfData.TextMatrix(lngRow, mlngYear)
        Else
            '删除多余的数据行,仅留一行
            lngRow = GetStartRow(VsfData.ROW)
            If Val(VsfData.TextMatrix(lngRow, mlngCollectType)) <> 0 Then Exit Sub
            lngStartRow = lngRow
            strDate = VsfData.TextMatrix(lngRow, mlngDate)
            strTime = VsfData.TextMatrix(lngRow, mlngTime)
            strYear = VsfData.TextMatrix(lngRow, mlngYear)
            lngRows = Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(0)) - 1
            For intNULL = 1 To lngRows
                VsfData.RemoveItem lngRow + 1
            Next
            '更新数据行
            Call CellMap_Update(lngStartRow, -1 * lngRows)
        End If
        
        '往下搜索空行,如果有其它数据行则计算需增加的行数
        intNULL = mrsCopyMap.RecordCount - 1
        For lngRow = 1 To mrsCopyMap.RecordCount - 1
            '保证当前输入的内容在一页中显示全
            If lngRow + lngStartRow > VsfData.Rows - 1 Then Exit For
            
            If Val(VsfData.TextMatrix(lngRow + lngStartRow, mlngRecord)) = 0 And VsfData.TextMatrix(lngRow + lngStartRow, mlngRowCount) = "" Then
                intNULL = intNULL - 1
            Else
                Exit For
            End If
        Next
        '先增加空行
        If intNULL > 0 Then
            VsfData.Rows = VsfData.Rows + intNULL
            '从当前行记录的空白行开始，每行的位置+所增加的空白行数
            For lngRow = 1 To intNULL
                VsfData.RowPosition(VsfData.Rows - 1) = lngStartRow + 1
            Next
        End If
        
        '还原日期，时间，强制不允许修改
        VsfData.TextMatrix(lngStartRow, mlngDate) = strDate
        VsfData.TextMatrix(lngStartRow, mlngTime) = strTime
        VsfData.TextMatrix(lngStartRow, mlngYear) = strYear
        '记录用户修改过的单元格
        If mlngDate <> -1 Then
            strKey = mint页码 & "," & lngStartRow & "," & mlngDate
            strValue = strKey & "|" & mint页码 & "|" & lngStartRow & "|" & mlngDate & "|" & _
                Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "|" & VsfData.TextMatrix(lngStartRow, mlngDate) & "|" & strYear & "|0"
            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
        End If
        '2\时间
        strKey = mint页码 & "," & lngStartRow & "," & mlngTime
        strValue = strKey & "|" & mint页码 & "|" & lngStartRow & "|" & mlngTime & "|" & _
            Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "|" & VsfData.TextMatrix(lngStartRow, mlngTime) & "||0"
        Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
        '向表格填充数据
        With mrsCopyMap
            If .RecordCount <> 0 Then .MoveFirst
            Do While Not .EOF
                For lngCol = 0 To lngCols - VsfData.FixedCols
                    Select Case lngCol + VsfData.FixedCols
                    Case mlngDemo, mlngActTime, mlngChoose, mlngYear, mlngDate, mlngTime, mlngOperator, mlngSigner, mlngSignTime, mlngRecord, mlngSignName, mlngSignLevel, mlngJoinSignName
                    Case Else
                        If Not (blnTure = True And InStr(1, "," & strCols & ",", "," & lngCol - (cHideCols - 1) & ",") > 0) Then
                            If InStr(1, "," & strLockItem & ",", "," & lngCol - (cHideCols - 1) & ",") = 0 And InStr(1, "," & mstrCOLNothing & ",", "," & lngCol - (cHideCols - 1) & ",") = 0 Then
                                VsfData.TextMatrix(lngStartRow + .AbsolutePosition - 1, lngCol + VsfData.FixedCols) = NVL(.Fields(cControlFields + lngCol).Value)
                                
                                '修改标志
                                If .AbsolutePosition = .RecordCount And lngCol < mlngNoEditor Then
                                    strKey = mint页码 & "," & lngStartRow & "," & lngCol + VsfData.FixedCols
                                    strValue = strKey & "|" & mint页码 & "|" & lngStartRow & "|" & lngCol + VsfData.FixedCols & "|" & _
                                        Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "|" & GetMutilData(lngStartRow, lngCol + VsfData.FixedCols, lngTop, lngHeight) & "||0"
                                    Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                                End If
                            End If
                        End If
                    End Select
                Next
                .MoveNext
            Loop
        End With
        '当行号发生变化后，需同步更新mrsCellMap中大于该行号的行号数据
        Call CellMap_Update(lngStartRow, mrsCopyMap.RecordCount - 1)

        '表格上色
        'Call WriteColor
        mblnChange = True
        RaiseEvent AfterDataChanged(mblnChange Or mblnVerify)
    
    Case conMenu_Edit_Clear
        If VsfData.ROW < VsfData.FixedRows Then Exit Sub
        If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then Exit Sub
        If VsfData.TextMatrix(VsfData.ROW, mlngSigner) <> "" Then
            RaiseEvent AfterRowColChange("已签名的数据不允许删除！", True, mblnSign, mblnArchive)
            Exit Sub
        End If
        
        '跨页数据行不允许对整行进行粘贴,删除,只能编辑除活动项目外的列
        If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then
            lngStartRow = VsfData.ROW
        Else
            lngStartRow = GetStartRow(VsfData.ROW)
        End If
        If InStr(1, VsfData.TextMatrix(lngStartRow, mlngRowCount), "|") <> 0 And lngStartRow = 3 And Val(VsfData.TextMatrix(lngStartRow, mlngStartRowPage)) <> mint页码 Then
            If Val(Split(VsfData.TextMatrix(lngStartRow, mlngRowCount), "|")(0)) <> Val(VsfData.TextMatrix(lngStartRow, mlngRowCurrent)) Then
                RaiseEvent AfterRowColChange("跨页数据行不允许删除，请切换到上一页进行操作！", True, mblnSign, mblnArchive)
                Exit Sub
            End If
        End If
        
        lngRowCount = Split(VsfData.TextMatrix(lngStartRow, mlngRowCount), "|")(0)
        '检查目标数据行是否存在同步过来的数据,如果有则跳过同步的记录
        strLockItem = GetSynItems(2, intMax)        '1.返回项目序号;2.返回列号
        
        '准备删除
        strField = "ID|页号|行号|列号|记录ID|数据|汇总|删除"
        If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then
            lngRow = VsfData.ROW
            lngRows = 1
        Else
            lngRow = GetStartRow(VsfData.ROW)
            lngStartRow = lngRow
            If VsfData.TextMatrix(lngStartRow, mlngSigner) <> "" Then
                RaiseEvent AfterRowColChange("已签名的数据不允许删除！", True, mblnSign, mblnArchive)
                Exit Sub
            End If
            lngRows = Val(Split(VsfData.TextMatrix(lngStartRow, mlngRowCount), "|")(0))
        End If
        
        blnTure = False
        '已经分组的数据不允许删除起始行
        If Val(VsfData.TextMatrix(lngRow, mlngDemo)) = 1 And lngRow + lngRows < VsfData.Rows Then
            lngCount = lngRow + lngRows - 1
            For lngCurRow = lngRow + lngRows To VsfData.Rows - 1
                 '分组中每一个分组子行数据可能占用多行，只有在第一行保存了分组索引。分组总数据行数=每一个分组子行数+每一个分组子行的行数
                If lngCurRow > lngCount Then
                    If Val(VsfData.TextMatrix(lngCurRow, mlngDemo)) <= 1 Then Exit For
                    If VsfData.RowHidden(lngCurRow) = False Then blnTure = True: Exit For '只要存在一个没有隐藏的分组就退出
                    lngCount = Val(Split(VsfData.TextMatrix(lngCurRow, mlngRowCount), "|")(0)) + lngCurRow - 1
                End If
            Next lngCurRow
        End If
        If blnTure = True Then
            RaiseEvent AfterRowColChange("存在分组数据行时，不允许删除分组起始行。", True, mblnSign, mblnArchive)
            Exit Sub
        End If
        
        '已有的数据存在汇总已签名的数据不允许删除
        If Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) > 0 And Not Val(VsfData.TextMatrix(lngStartRow, mlngCollectType)) < 0 Then
            If CheckCollectIsData(lngStartRow, 1) = True Then
                If ISCollectSigned(mlng文件ID, Format(VsfData.TextMatrix(lngStartRow, mlngActTime), "YYYY-MM-DD"), Format(VsfData.TextMatrix(lngStartRow, mlngActTime), "HH:MM")) Then
                    RaiseEvent AfterRowColChange("您要删除的数据存在汇总列数据，且本条数据所对应的汇总数据已签名，不允许删除。", True, mblnSign, mblnArchive)
                    Exit Sub
                End If
            End If
        End If
        '隐蔽已显示的录入控件
        Select Case mintType
        Case 0, 3
            picInput.Visible = False
        Case 1, 2
            lstSelect(mintType - 1).Visible = False
            If mintType = 1 Then
                txtLst.Visible = False
                PicLst.Visible = False
            End If
        Case 4, 5
            picDouble.Visible = False
        Case 6
            picMutilInput.Visible = False
        Case 7
            picDoubleChoose.Visible = False
        Case 8
            picYear.Visible = False
        End Select
        cmdWord.Visible = False
        mintType = -1
        blnNULL = mblnShow
        mblnShow = False
        
        strAssistantCols = ""
        '获取分组数据的大文本列数据内容
        If Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) > 1 Then
            Call GetGroupAssistant(strAssistantCols, varAssistant)
        End If
        '删除所有数据行
        For intNULL = 2 To lngRows
            VsfData.RowHidden(lngRow + intNULL - 1) = True
        Next
        
        '清除非起始行分组数据，不清除大文本信息并取消该分组
        '如：本分组包含三组，清除第二组时，将第二组大文段内容累加在第3祖上
        '记录用户修改过的单元格
        If Val(VsfData.TextMatrix(lngStartRow, mlngCollectType)) = 0 Then
            strYear = ""
            If Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) > 0 Then
                If CheckGroupDate(lngStartRow) = True Then
                    '保存后的修改才进入此流程，取该条记录的实际时间
                    strDate = Mid(VsfData.TextMatrix(lngStartRow, mlngActTime), 1, 10)
                    strTime = Mid(VsfData.TextMatrix(lngStartRow, mlngActTime), 12, 5)
                    If mblnDateAd Then strYear = Mid(VsfData.TextMatrix(lngStartRow, mlngActTime), 1, 4)
                Else
                    '新增时进入此流程
                    strDate = VsfData.TextMatrix(lngStartRow - Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) + 1, mlngDate)
                    strTime = VsfData.TextMatrix(lngStartRow - Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) + 1, mlngTime)
                    If mblnDateAd Then strYear = VsfData.TextMatrix(lngStartRow - Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) + 1, mlngYear)
                End If
            Else
                '普通数据
                strDate = VsfData.TextMatrix(lngStartRow, mlngDate)
                strTime = VsfData.TextMatrix(lngStartRow, mlngTime)
                If mblnDateAd Then strYear = VsfData.TextMatrix(lngStartRow, mlngYear)
            End If
            If mblnDateAd Then
                If InStr(1, strDate, "/") <> 0 Then
                    strDate = Mid(zlDatabase.Currentdate, 1, 5) & Split(strDate, "/")(1) & "-" & Split(strDate, "/")(0)
                End If
                strDate = Mid(strDate, 9, 2) & "/" & Mid(strDate, 6, 2)
            End If
            
            strField = "ID|页号|行号|列号|记录ID|数据|部位|汇总|删除"
            strKey = mint页码 & "," & lngStartRow & "," & mlngDate
            strValue = strKey & "|" & mint页码 & "|" & lngStartRow & "|" & mlngDate & "|" & _
                Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "|" & strDate & "|" & strYear & "|0|1"
            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
            '2\时间
            strKey = mint页码 & "," & lngStartRow & "," & mlngTime
            strValue = strKey & "|" & mint页码 & "|" & lngStartRow & "|" & mlngTime & "|" & _
                Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "|" & strTime & "|" & VsfData.TextMatrix(lngStartRow, mlngDemo) & "|0|1"
            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
        Else
            '1\日期
            strKey = mint页码 & "," & lngStartRow & "," & mlngDate
            strValue = strKey & "|" & mint页码 & "|" & lngStartRow & "|" & mlngDate & "|" & Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "|" & _
                    VsfData.TextMatrix(lngStartRow, mlngCollectText) & ";" & Val(VsfData.TextMatrix(lngStartRow, mlngCollectType)) & ";" & _
                    Val(VsfData.TextMatrix(lngStartRow, mlngCollectStyle)) & ";" & VsfData.TextMatrix(lngStartRow, mlngCollectDay) & ";" & _
                    VsfData.TextMatrix(lngStartRow, mlngCollectStart) & ";" & VsfData.TextMatrix(lngStartRow, mlngCollectEnd) & "|1|1"
            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
        End If
        strField = "ID|页号|行号|列号|记录ID|数据|部位|汇总|删除"
        
        '删除启始行中非同步的数据
        If strLockItem = "" Then
            VsfData.RowHidden(lngRow) = True
            If Val(VsfData.TextMatrix(lngStartRow, mlngCollectType)) = 0 Then
                '填写修改标志
                For lngCol = mlngTime + 1 To mlngNoEditor - 1
                    If InStr(1, "," & mstrCatercorner & ",", "," & lngCol - (cHideCols + VsfData.FixedCols - 1) & ",") = 0 Then
                        strPart = GetActivePart(lngCol, 0)
                    Else
                        strPart = GetActivePart(lngCol, 0) & "/" & GetActivePart(lngCol, 1)
                    End If
                    strKey = mint页码 & "," & lngStartRow & "," & lngCol
                    strValue = strKey & "|" & mint页码 & "|" & lngStartRow & "|" & lngCol & "|" & _
                        Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "||" & strPart & "|0|1"
                    Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                Next
                If Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) >= 1 Then VsfData.TextMatrix(lngStartRow, mlngRowCount) = "1|1"
            End If
        Else
            '填写修改标志(存在同步数据,日期与时间列不允许清除)``
            For lngCol = mlngTime + 1 To mlngNoEditor - 1
                If InStr(1, "," & strLockItem & ",", "," & lngCol - (cHideCols + VsfData.FixedCols - 1) & ",") = 0 And lngCol <> mlngDate And lngCol <> mlngTime Then
                    VsfData.TextMatrix(lngStartRow, lngCol) = ""
                    If InStr(1, "," & mstrCatercorner & ",", "," & lngCol - (cHideCols + VsfData.FixedCols - 1) & ",") = 0 Then
                        strPart = GetActivePart(lngCol, 0)
                    Else
                        strPart = GetActivePart(lngCol, 0) & "/" & GetActivePart(lngCol, 1)
                    End If
                    strKey = mint页码 & "," & lngStartRow & "," & lngCol
                    strValue = strKey & "|" & mint页码 & "|" & lngStartRow & "|" & lngCol & "|" & _
                        Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "||" & strPart & "|0|1"
                    Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                End If
            Next
            VsfData.TextMatrix(lngStartRow, mlngRowCount) = "1|1"
        End If
        
        Call FillPage(False)
        
        If Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) >= 1 Then
            For intNULL = 2 To lngRows
                For lngCol = 0 To VsfData.Cols - 1
                        VsfData.TextMatrix(lngRow + 1, lngCol) = ""
                Next lngCol
                VsfData.RowPosition(lngRow + 1) = VsfData.Rows - 1
            Next
            Call CellMap_Update(lngStartRow, -1 * (lngRows - 1))
            lngRowCount = 1
            
            '重新组织分组行号和大文本段内容
            If strAssistantCols <> "" Then
                Call ReSetGroupAssistant(True, False, strAssistantCols, varAssistant)
            Else
                Call ReSetGroupDemo(lngStartRow)
            End If
        End If

        mblnShow = False
        If lngStartRow + lngRowCount < VsfData.Rows - 1 Then
            lngRow1 = lngStartRow + lngRowCount
            If Val(VsfData.TextMatrix(lngRow1, mlngRowCount)) > 1 Then
                lngRow1 = GetStartRow(lngRow1)
                If lngRow1 + Val(Split(VsfData.TextMatrix(lngRow1, mlngRowCount), "|")(0)) < VsfData.Rows - 1 Then
                    lngRow1 = lngRow1 + Val(Split(VsfData.TextMatrix(lngRow1, mlngRowCount), "|")(0))
                End If
            End If
            
            If VsfData.RowHidden(lngRow1) = False Then
                VsfData.ROW = lngRow1
            Else
                For lngRow = lngRow1 + 1 To VsfData.Rows - 1
                    If VsfData.RowHidden(lngRow) = False Then VsfData.ROW = lngRow: Exit For
                Next lngRow
            End If
        End If
        
        mblnChange = True
        mblnShow = blnNULL
        RaiseEvent AfterDataChanged(mblnChange Or mblnVerify)
        
    Case conMenu_Edit_SPECIALCHAR
        
        '检查当前录入控件
        On Error Resume Next
        Dim objTXT As TextBox
        Dim intPos As Integer, intLen As Integer
        
        mstrSymbol = frmInsSymbol.ShowMe(False, 0)
        If mintSymbol = -1 Then
            Set objTXT = txtInput
        Else
            Set objTXT = txt(mintSymbol)
        End If
        objTXT.SetFocus
        strText = objTXT.Text
        intPos = objTXT.SelStart
        intLen = Len(objTXT)
        objTXT.Text = Mid(strText, 1, intPos) & mstrSymbol & Mid(strText, intPos + 1)
    
        If mintSymbol = -1 Then
            Call txtInput_KeyDown(vbKeyReturn, 0)
        Else
            Call txt_KeyDown(Val(txt(mintSymbol)), vbKeyReturn, 0)
        End If
    Case conMenu_Edit_Element
        If frmTendFileElement.ShowMe(mfrmParent, mlng文件ID, mlng格式ID, mint页码, mrsElement, IIf(mblnBlowup = True, 1, 0)) = True Then
            '重新提取数据
            mblnInit = False
            Call InitVariable
            Call InitCons
            Call ReadStruDef
            Call zlRefresh
            mblnInit = True
            VsfData.Refresh
            mcbrPage.Caption = "页码选择：第" & mint页码 & "页"
            cbsThis.RecalcLayout
        End If
    Case conMenu_Edit_Append
        Call BoundItems(VsfData.COL - (cHideCols + VsfData.FixedCols - 1))
    Case conMenu_Edit_PrevPage
        If mint页码 > mint起始页码 Then
            If Not DataMap_Save Then Exit Sub
            mint页码 = mint页码 - 1
            '更新查询SQL
            '重新提取数据
            mblnInit = False
            Call InitVariable
            Call InitCons
            Call ReadStruDef
            Call zlRefresh
            mblnInit = True
            VsfData.Refresh
        
            mcbrPage.Caption = "页码选择：第" & mint页码 & "页"
            cbsThis.RecalcLayout
        End If
    Case conMenu_Edit_NextPage
        If mint页码 < mint结束页 + 1 Then
            If Not DataMap_Save Then Exit Sub
            mint页码 = mint页码 + 1
            '更新查询SQL
            '重新提取数据
            mblnInit = False
            Call InitVariable
            Call InitCons
            Call ReadStruDef
            Call zlRefresh
            mblnInit = True
            VsfData.Refresh
            
            mcbrPage.Caption = "页码选择：第" & mint页码 & "页"
            cbsThis.RecalcLayout
        End If
    Case conMenu_View_Jump
        If Not DataMap_Save Then Exit Sub
        
        '更新查询SQL
        '重新提取数据
        mint页码 = Control.Parameter
        mblnInit = False
        Call InitVariable
        Call InitCons
        Call ReadStruDef
        Call zlRefresh
        
        mblnInit = True
        VsfData.Refresh
        
        mcbrPage.Caption = "页码选择：第" & mint页码 & "页"
        cbsThis.RecalcLayout
    Case conMenu_Edit_Word
        Call cmdWord_Click
    Case conMenu_Edit_Brief
        Call ShowBrief
    Case conMenu_Edit_Import
        '导入入量
        If VsfData.ROW < VsfData.FixedRows Then Exit Sub '其它条件已经在update中进行了判断
        Call ImportAmount
    Case conMenu_Tool_SignEarse
        Call UnSignMe(False, True)
    Case conMenu_Tool_SignAuditCancel
        Call UnSignMe(True, True)
    End Select
    Exit Sub
err_exit:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim arrData
    Dim blnFind As Boolean
    Dim strItem As String
    Dim intDo  As Integer, intCount As Integer
    On Error GoTo ErrHand
    
    If Not mblnInit Then
        Control.Enabled = False
        Exit Sub
    Else
        Control.Enabled = True
        Select Case Control.ID
            Case conMenu_Edit_Group_New, conMenu_Edit_Group_Append, conMenu_Edit_Copy, conMenu_Edit_PASTE, _
                conMenu_Edit_Clear, conMenu_Edit_SPECIALCHAR, conMenu_Edit_Element, conMenu_Edit_Append, conMenu_Edit_Word, conMenu_Edit_Brief, conMenu_Edit_Import
                Control.Visible = InStr(1, mstrPrivs, "护理记录登记") <> 0
        End Select
    End If
    
    Select Case Control.ID
    Case conMenu_Edit_Group_New  '分组，只对新添加的数据有效
        Control.Checked = mblnGroupNew
        Control.Enabled = mblnEditable And Not mblnArchive And Not mblnVerify And Not mblnGroupApp _
            And Val(VsfData.TextMatrix(VsfData.ROW, mlngRecord)) = 0 And Val(mstrGroupRow) <= VsfData.ROW
        '63934:刘鹏飞,2013-07-25,小结行不能使用追加功能
        If Control.Enabled Then
            If VsfData.TextMatrix(VsfData.ROW, mlngRowCount) = "" Then
                intDo = VsfData.ROW
            Else
                intDo = GetStartRow(VsfData.ROW)
            End If
            Control.Enabled = Control.Visible And Val(VsfData.TextMatrix(intDo, mlngCollectType)) >= 0 And intDo = VsfData.ROW
        End If
    Case conMenu_Edit_Group_Append
        Control.Checked = False 'mblnGroupApp
        Control.Enabled = Control.Visible And mblnEditable And Not mblnArchive And Not mblnVerify And Not mblnGroupNew
        If Control.Enabled Then
            If VsfData.TextMatrix(VsfData.ROW, mlngRowCount) = "" Then
                Control.Enabled = False
            Else
                '63934:刘鹏飞,2013-07-25,小结行不能使用追加功能
                '1、已经签名的数据不能追加.2、追加只能在有数据的行追加(不包含大文本项目).3、小结行不能追加
                intDo = GetStartRow(VsfData.ROW)
                Control.Enabled = IIf(VsfData.TextMatrix(intDo, mlngSigner) <> "", False, True) And (ISGroupAppend = True) _
                    And Val(VsfData.TextMatrix(intDo, mlngCollectType)) >= 0
            End If
        End If
    Case conMenu_Edit_Copy
        Control.Enabled = False
        If VsfData.TextMatrix(VsfData.ROW, mlngRowCount) <> "" Then
            intDo = GetStartRow(VsfData.ROW)
        Else
            intDo = VsfData.ROW
        End If
        If Val(VsfData.TextMatrix(intDo, mlngCollectType)) <> 0 Then Exit Sub
        Control.Enabled = Control.Visible And Not mblnShow And Not mblnArchive And Not mblnVerify And mblnEditable
        
    Case conMenu_Edit_PASTE
        Control.Enabled = False
        If mrsCopyMap.State = 0 Then Exit Sub
        '签名数据不允许粘贴
        If VsfData.TextMatrix(VsfData.ROW, mlngRowCount) <> "" Then
            intDo = GetStartRow(VsfData.ROW)
        Else
            intDo = VsfData.ROW
        End If
        If Val(VsfData.TextMatrix(intDo, mlngDemo)) >= 1 Then Exit Sub
        If VsfData.TextMatrix(intDo, mlngSigner) <> "" Then Exit Sub
        If Val(VsfData.TextMatrix(intDo, mlngCollectType)) <> 0 Then Exit Sub
        '粘贴不能在复制行粘贴
        mrsCopyMap.Filter = "页号=" & mint页码 & " And 行号=" & intDo
        If mrsCopyMap.RecordCount > 0 Then Exit Sub
        mrsCopyMap.Filter = ""
        Control.Enabled = Control.Visible And Not mblnShow And Not mblnArchive And mblnEditable And mrsCopyMap.RecordCount
    Case conMenu_Edit_Clear
        Control.Enabled = False
        If VsfData.TextMatrix(VsfData.ROW, mlngRowCount) <> "" Then
            intDo = GetStartRow(VsfData.ROW)
        Else
            intDo = VsfData.ROW
        End If
        If VsfData.TextMatrix(intDo, mlngSigner) <> "" Then Exit Sub
        
        Control.Enabled = Control.Visible And Not mblnArchive And Not mblnVerify And mblnEditable
    Case conMenu_Edit_SPECIALCHAR
        Control.Enabled = Control.Visible And mblnShow And Not mblnArchive And mblnEditable And (mintType = 0 Or mintType = 6)
    Case conMenu_Edit_Element
        Control.Enabled = Control.Visible And Not mblnArchive And mblnEditable And mblnElement
    Case conMenu_Edit_Append
        Control.Enabled = Control.Visible And (InStr(1, "," & mstrCOLNothing & ",", "," & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ",") <> 0) And Not mblnArchive And mblnEditable
    Case conMenu_Edit_PrevPage
        Control.Enabled = (mint页码 > mint起始页码)
    Case conMenu_Edit_NextPage
        Control.Enabled = (mint页码 < mint结束页 + 1)
    Case conMenu_Edit_Word
        '60291:刘鹏飞,2013-04-17,只要是文本项目都允许进行词句选择
        Control.Enabled = Control.Visible And (mblnEditAssistant Or mblnEditText) And mblnShow And Not mblnArchive And mblnEditable
    Case conMenu_Edit_Brief
        Control.Enabled = Control.Visible And Not mblnArchive And Not mblnVerify And mblnEditable
    Case conMenu_Edit_Import '入量导入
        Control.Enabled = Control.Visible And Not mblnArchive And Not mblnVerify And mblnEditable And mblnShow And mstrColCollect <> ""
         If Control.Enabled Then
            '判断选择的列是否是汇总项目列(一列绑定两个汇总列也不能使用此功能)
            blnFind = False
            For intDo = 0 To UBound(Split(mstrColCollect, "|"))
                If VsfData.COL - (cHideCols + VsfData.FixedCols - 1) = Split(Split(mstrColCollect, "|")(intDo), ";")(0) And InStr(1, Split(Split(mstrColCollect, "|")(intDo), ";")(1), ",") = 0 Then
                    blnFind = True
                    Exit For
                End If
            Next
            intDo = GetStartRow(VsfData.ROW)
            Control.Enabled = blnFind And IIf(VsfData.TextMatrix(intDo, mlngSigner) <> "", False, True) And Val(VsfData.TextMatrix(intDo, mlngCollectType)) >= 0
        End If
        
    Case conMenu_View_Jump
        Control.Checked = (Val(Control.Parameter) = mint页码)
    Case conMenu_Tool_SignEarse, conMenu_Tool_SignAuditCancel '单个取消签名和审签
        If VsfData.TextMatrix(VsfData.ROW, mlngRowCount) = "" Then
            Control.Visible = False
        Else
            intDo = GetStartRow(VsfData.ROW)
            Control.Visible = VsfData.TextMatrix(intDo, mlngSigner) <> "" And Val(VsfData.TextMatrix(intDo, mlngRecord)) > 0
        End If
        Control.Enabled = Control.Visible And mblnEditable And Not mblnArchive And Not mblnVerify And Not mblnChange
        If Control.Enabled Then
            If Control.ID = conMenu_Tool_SignEarse Then
                Control.Enabled = Control.Enabled And InStr(1, VsfData.TextMatrix(intDo, mlngSigner), "/") = 0
            Else
                Control.Enabled = Control.Enabled And InStr(1, VsfData.TextMatrix(intDo, mlngSigner), "/") <> 0
            End If
        End If
        
    End Select
ErrHand:
End Sub

Private Function ISGroupAppend() As Boolean
'追加分组数据，在选择的行有数据才能追加（不包含大文本项目）
    Dim lngCol As Long, lngRow As Long
    Dim blnNULL As Boolean
    
    lngRow = VsfData.ROW
    If lngRow > VsfData.Rows - 1 Then lngRow = VsfData.Rows - 1
    blnNULL = True
    For lngCol = mlngTime + 1 To VsfData.Cols - 1
        If Not VsfData.ColHidden(lngCol) And lngCol < mlngNoEditor And ISEditAssistant(lngCol) = False Then
            If FormatValue(VsfData.TextMatrix(lngRow, lngCol)) <> "" And Not (IsDiagonal(lngCol) And InStr(1, FormatValue(VsfData.TextMatrix(lngRow, lngCol)), "/") <> 0) Then
                blnNULL = False
                Exit For
            End If
        End If
    Next
    
    ISGroupAppend = Not blnNULL
End Function

Private Sub chkSwitch_Click()
    Dim blnSel As Boolean            '是否全部选中
    Dim blnUpdate As Boolean
    Dim intLevel As Integer
    Dim lngRow As Long, lngRows As Long
    Dim strKey As String, strField As String, strValue As String
    Dim lngStart As Long, lngDemo As Long, lngCurRow As Long, lngRowCount As Long, lngNextGroupRow As Long
    Dim arrRow, blnTrue As Boolean
    '将所有列全部选中或取消选中，并保存更新
    
    If Not mblnInit Then Exit Sub
    lngRows = VsfData.Rows - 1
    strField = "ID|页号|行号|列号|记录ID|数据|删除"
    
    blnSel = chkSwitch.Value
    For lngRow = VsfData.FixedRows To lngRows
        If Not VsfData.RowHidden(lngRow) Then
            If FormatValue(VsfData.TextMatrix(lngRow, mlngRowCount)) Like "*|1" Then
                '汇总数据也要签名,因此注释
                'If Val(VsfData.TextMatrix(lngRow, mlngCollectType)) = 0 Then    '汇总行不允许编辑
                    blnUpdate = False
                    blnTrue = blnSel
                    If blnSel Then
                        '检查,签过名的记录,且当前操作员级别比上次签名级别高
                        If VsfData.TextMatrix(lngRow, mlngSignLevel) = "" Then
                            intLevel = 未定义
                        Else
                            intLevel = Val(VsfData.TextMatrix(lngRow, mlngSignLevel)) + 1
                        End If
                        '43588:刘鹏飞,2012-09-13,添加记录单审签模式
                        If IIf(mintSignMode = 0, mintVerify < intLevel, InStr(1, VsfData.TextMatrix(lngRow, mlngSigner), "/") = 0 And mintVerify <> 未定义) And intLevel <> 未定义 Then
                        'If mintVerify < intLevel And intLevel <> 未定义 then
'                            blnUpdate = (VsfData.Cell(flexcpChecked, lngRow, mlngChoose) <> flexTSChecked)
'                            VsfData.Cell(flexcpChecked, lngRow, mlngChoose) = flexTSChecked
                            blnTrue = True
                        Else
                            blnTrue = False
                        End If
                    Else
'                        blnUpdate = (VsfData.Cell(flexcpChecked, lngRow, mlngChoose) <> flexTSUnchecked)
'                        VsfData.Cell(flexcpChecked, lngRow, mlngChoose) = flexTSUnchecked
                        blnTrue = False
                    End If
                    blnUpdate = True
                    lngStart = lngRow
                    If blnUpdate Then
                        lngDemo = Val(VsfData.TextMatrix(lngStart, mlngDemo))
                        If lngDemo > 1 Then '寻找起始行
                            lngCurRow = lngStart
                            lngStart = lngCurRow - lngDemo + 1
                            If Val(VsfData.TextMatrix(lngStart, mlngDemo)) <> 1 Then
                                For lngStart = lngCurRow To VsfData.FixedRows Step -1
                                    If Val(VsfData.TextMatrix(lngStart, mlngDemo)) = 1 Then
                                        lngCurRow = lngStart
                                        Exit For
                                    End If
                                Next lngStart
                                If lngStart < VsfData.FixedRows Then Exit Sub
                                lngStart = lngCurRow
                            End If
                        End If
                        arrRow = Array()
                        ReDim Preserve arrRow(UBound(arrRow) + 1)
                        arrRow(UBound(arrRow)) = lngStart
                        lngRowCount = Val(Split(VsfData.TextMatrix(lngStart, mlngRowCount), "|")(0))
                        lngNextGroupRow = lngStart + lngRowCount - 1
                        '1.检查分组数据是否已经签名，2.记录分组数据开始行
                        For lngCurRow = lngStart + lngRowCount To VsfData.Rows - 1
                            '分组中每一个分组子行数据可能占用多行，只有在第一行保存了分组索引。分组总数据行数=每一个分组子行数+每一个分组子行的行数
                            If lngCurRow > lngNextGroupRow Then
                                If Val(VsfData.TextMatrix(lngCurRow, mlngDemo)) <= 1 Then Exit For      '不分组或遇新分组就退出
                                If blnSel = True Then
                                    If VsfData.TextMatrix(lngCurRow, mlngSignLevel) = "" Then
                                        intLevel = 未定义
                                    Else
                                        intLevel = Val(VsfData.TextMatrix(lngCurRow, mlngSignLevel)) + 1
                                    End If
                                    '43588:刘鹏飞,2012-09-13,添加记录单审签模式
                                    If Not (IIf(mintSignMode = 0, mintVerify < intLevel, InStr(1, VsfData.TextMatrix(lngCurRow, mlngSigner), "/") = 0 And mintVerify <> 未定义) And intLevel <> 未定义) And blnTrue = True Then
                                    'If Not (mintVerify < intLevel And intLevel <> 未定义) And blnTrue = True Then
                                        blnTrue = False
                                    End If
                                Else
                                    blnTrue = False
                                End If
                                lngNextGroupRow = Val(Split(VsfData.TextMatrix(lngCurRow, mlngRowCount), "|")(0)) + lngCurRow - 1
                                ReDim Preserve arrRow(UBound(arrRow) + 1)
                                arrRow(UBound(arrRow)) = lngCurRow
                            End If
                        Next lngCurRow
                        '选中所有分组数据
                        For lngCurRow = 0 To UBound(arrRow)
                            lngStart = Val(arrRow(lngCurRow))
                            If (VsfData.Cell(flexcpChecked, lngStart, mlngChoose) <> IIf(blnTrue = False, flexTSUnchecked, flexTSChecked)) Then
                                VsfData.Cell(flexcpChecked, lngStart, mlngChoose) = IIf(blnTrue = False, flexTSUnchecked, flexTSChecked)
                                '保存修改记录以便同步
                                strKey = mint页码 & "," & lngStart & "," & mlngChoose
                                strValue = strKey & "|" & mint页码 & "|" & lngStart & "|" & mlngChoose & "|" & Val(VsfData.TextMatrix(lngStart, mlngRecord)) & "|" & VsfData.Cell(flexcpChecked, lngStart, mlngChoose) & "|1"
                                Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                            End If
                        Next lngCurRow
                    End If
                    
                    '移动行
                    lngRow = Val(arrRow(UBound(arrRow))) + Val(Split(VsfData.TextMatrix(Val(arrRow(UBound(arrRow))), mlngRowCount), "|")(0)) - 1
                'End If
            End If
        End If
    Next
End Sub

Private Sub cmdCanCel_Click()
    picBiref.Visible = False
End Sub

Private Sub cmdColumn_Click(Index As Integer)
    Dim lstItem As ListItem
    
    If cmdColumn(Index).Enabled = False Then Exit Sub
    If Index = 0 Then
        'add
        If Not lstColumnItems.SelectedItem Is Nothing Then
            Set lstItem = lstColumnUsed.ListItems.Add(, lstColumnItems.SelectedItem.Key, lstColumnItems.SelectedItem.Text)
            lstItem.SubItems(1) = lstColumnItems.SelectedItem.SubItems(1)
            lstItem.SubItems(2) = lstColumnItems.SelectedItem.SubItems(2)
            lstColumnItems.ListItems.Remove lstColumnItems.SelectedItem.Index
        End If
        If txtColumnNo.Text = "" Then
            txtColumnNo.Text = Replace(lstItem.SubItems(1), lstItem.SubItems(2), "")
        End If
    Else
        'del
        If Not lstColumnUsed.SelectedItem Is Nothing Then
            Set lstItem = lstColumnItems.ListItems.Add(, lstColumnUsed.SelectedItem.Key, lstColumnUsed.SelectedItem.Text)
            lstItem.SubItems(1) = lstColumnUsed.SelectedItem.SubItems(1)
            lstItem.SubItems(2) = lstColumnUsed.SelectedItem.SubItems(2)
            lstColumnUsed.ListItems.Remove lstColumnUsed.SelectedItem.Index
            If lstColumnUsed.ListItems.Count = 0 Then txtColumnNo.Text = ""
        End If
    End If
End Sub

Private Sub cmdFilterCancel_Click()
    picCloumn.Visible = False
End Sub

Private Sub cmdFilterOK_Click()
    Dim strPara As String
    Dim strTest As String
    Dim lngCol As Long, lngRow As Long
    Dim intDo As Integer, intCount As Integer, intFace As Integer, intType As Integer
    On Error GoTo ErrHand
    
    '63401:刘鹏飞,2013-07-16,检查选择的列是否是活动项目列
    If InStr(1, "," & mstrCOLNothing & ",", "," & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ",") = 0 Then
        RaiseEvent AfterRowColChange("当前选中的列不能绑定活动项目，请重新选择列进行绑定！", True, mblnSign, mblnArchive)
        picCloumn.Visible = False
        Exit Sub
    End If
    
    If lstColumnUsed.ListItems.Count > 0 Then
        If Trim(txtColumnNo.Text) = "" Then
            RaiseEvent AfterRowColChange("表头名称不能为空！", True, mblnSign, mblnArchive)
            txtColumnNo.SetFocus
            Exit Sub
        End If
        If LenB(StrConv(txtColumnNo.Text, vbFromUnicode)) > 100 Then
            RaiseEvent AfterRowColChange("表头名称不能超过50个汉字或100个字符！", True, mblnSign, mblnArchive)
            txtColumnNo.SetFocus
            Exit Sub
        End If
    End If
    
    '拼串，格式：表头名称|项目序号,部位;项目序号,部位
    strPara = Trim(txtColumnNo.Text) & "|"
    intCount = lstColumnUsed.ListItems.Count
    If intCount > 2 Then
        RaiseEvent AfterRowColChange("每列绑定的项目数不能超过2个！", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    '51883,刘鹏飞,2012-08-02,活动项目支持单选和复选
    '规则为：1、文本项目和多选项目只能单独版定，2.每列绑定的活动项目不能超过两个。3、绑定多个项目时项目表示和项目类型必须一致
    '项目表示必须一致
    For intDo = 1 To intCount
        mrsItems.Filter = "项目序号=" & Val(lstColumnUsed.ListItems(intDo).Text)
        If mrsItems.RecordCount = 0 Then
            RaiseEvent AfterRowColChange("护理记录项目可能已经发生变化，请重新刷新护理页面", True, mblnSign, mblnArchive)
            mrsItems.Filter = 0
            Exit Sub
        End If
        If intDo = 1 Then
            intFace = Val(NVL(mrsItems!项目表示))
            intType = Val(NVL(mrsItems!项目类型))
        Else
            If Not (intFace = Val(NVL(mrsItems!项目表示)) And intType = Val(NVL(mrsItems!项目类型))) Then
                RaiseEvent AfterRowColChange("绑定的两个项目的表示和项目类型必须一致！（要么都是选择项，要么都是数值录入项）", True, mblnSign, mblnArchive)
                mrsItems.Filter = 0
                Exit Sub
            End If
            '文本型不允许绑定多个项目
            If Val(NVL(mrsItems!项目类型)) = 1 And Val(NVL(mrsItems!项目表示)) = 0 Then
                RaiseEvent AfterRowColChange("一列只能绑定一个文本型活动项目！", True, mblnSign, mblnArchive)
                mrsItems.Filter = 0
                Exit Sub
            End If
            '多选项目只能单独绑定
            If Val(NVL(mrsItems!项目表示, 0)) = 3 Then
                RaiseEvent AfterRowColChange("一列只能绑定一个多选型活动项目！", True, mblnSign, mblnArchive)
                mrsItems.Filter = 0
                Exit Sub
            End If
        End If
        
        '拼串
        strTest = lstColumnUsed.ListItems(intDo).Text
        '47764,刘鹏飞,2012-08-13,活动项目没有部位，对于不同列没有控制到不能绑定相同项目
'        If lstColumnUsed.ListItems(intDo).SubItems(2) <> "" Then
'            strTest = strTest & "," & lstColumnUsed.ListItems(intDo).SubItems(2)
'        End If
        strTest = strTest & "," & lstColumnUsed.ListItems(intDo).SubItems(2)
        If ISActiveUsed(strTest) Then Exit Sub
        
        strPara = strPara & IIf(intDo > 1, ";", "") & strTest
        mrsItems.Filter = 0
    Next
    
    '61852:刘鹏飞,2013-11-05,添加活动项目保存本页之前改变的数据
    If Not DataMap_Save Then picCloumn.Visible = False: Exit Sub
    
    '保存数据
    gstrSQL = "ZL_病人护理页面_UPDATE(" & mlng文件ID & "," & mint页码 & "," & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ",'" & strPara & "','" & gstrUserName & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存活动项目绑定数据")
    picCloumn.Visible = False
    lngCol = VsfData.COL
    lngRow = VsfData.ROW
    
    '更新查询SQL
    '重新提取数据
    mblnInit = False
    Call InitVariable
    Call InitCons
    Call ReadStruDef
    Call zlRefresh
    mblnInit = True
    
    VsfData.ROW = lngRow
    VsfData.COL = lngCol
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function DelActiveNoUsed() As Boolean
'------------------------------------------------
'功能:删除绑定在非空列上的活动项目列信息
'编制:刘鹏飞,2013-07-16
'问题号:63401
'------------------------------------------------
    Dim arrData, arrActive, arrCol
    Dim strSQL As String
    Dim lngCol As Long, intDo As Integer, intCount As Integer
    Dim blnTran As Boolean
    
    If mstrCOLNothing = "" Then DelActiveNoUsed = True: Exit Function
    arrActive = Array()
    arrCol = Array()
    arrData = Split(mstrCOLActive, "||")
    intCount = UBound(arrData)
    For intDo = 0 To intCount
        lngCol = Val(Split(Split(arrData(intDo), "|")(0), ";")(0))
        If InStr(1, "," & mstrCOLNothing & ",", "," & lngCol & ",") <> 0 Then
            '记录现有正常空列上的活动项目设置信息
            ReDim Preserve arrActive(UBound(arrActive) + 1)
            arrActive(UBound(arrActive)) = CStr(arrData(intDo))
        Else
            '记录将要移除的活动项目列号
            ReDim Preserve arrCol(UBound(arrCol) + 1)
            arrCol(UBound(arrCol)) = lngCol
        End If
    Next
    
    On Error GoTo ErrHand
    
    '删除不需要的活动项目信息(主要是修正之前错误的数据,发生情况较少)
    If UBound(arrCol) > 1 Then
        gcnOracle.BeginTrans
        blnTran = True
    End If
    
    For intDo = 0 To UBound(arrCol)
        If CStr(arrCol(intDo)) <> "" Then
            strSQL = "ZL_病人护理页面_UPDATE(" & mlng文件ID & "," & mint页码 & "," & Val(arrCol(intDo)) & ",NULL,'" & gstrUserName & "')"
            Call zlDatabase.ExecuteProcedure(strSQL, "保存活动项目绑定数据")
        End If
    Next
    If blnTran = True Then gcnOracle.CommitTrans
    
    '重新更新提取的活动项目列信息
    If UBound(arrActive) = -1 Then
        mstrCOLActive = ""
    Else
        mstrCOLActive = Join(arrActive, "||")
    End If
    
    DelActiveNoUsed = True
    Exit Function
ErrHand:
    If blnTran = True Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ISActiveUsed(ByVal strTest As String) As Boolean
    Dim arrData, arrCol
    Dim lngCol As Long
    Dim intDo As Integer, intCount As Integer
    Dim intIn As Integer, intMax As Integer
    '检查某个活动项目是否已被其它列绑定
    ISActiveUsed = True
    
    arrData = Split(mstrCOLActive, "||")
    intCount = UBound(arrData)
    For intDo = 0 To intCount
        arrCol = Split(Split(arrData(intDo), "|")(1), ";")
        lngCol = Split(Split(arrData(intDo), "|")(0), ";")(0)
        intMax = UBound(arrCol)
        For intIn = 0 To intMax
            If strTest = arrCol(intIn) And VsfData.COL - (cHideCols + VsfData.FixedCols - 1) <> lngCol Then
                RaiseEvent AfterRowColChange(Split(strTest, ",")(1) & mrsItems!项目名称 & " 已经被绑定到" & lngCol & "列，不允许重复绑定！", True, mblnSign, mblnArchive)
                Exit Function
            End If
        Next
    Next
    ISActiveUsed = False
End Function

Private Function GetActivePart(ByVal intFindCol As Integer, ByVal intItem As Integer) As String
    '获取指定列的活动项目
    Dim arrData
    Dim arrCol
    Dim intCol As Integer, strPart As String
    Dim intDo As Integer, intCount As Integer
    Dim intIn As Integer, intMax As Integer
    '将活动项目加入到查询SQL中，格式：列号;表头名称|项目序号,部位;项目序号,部位||列号;表头名称...
    '绑定多个项目，该列就自动转为对角线列
    
    arrData = Split(mstrCOLActive, "||")
    intCount = UBound(arrData)
    For intDo = 0 To intCount
        intCol = Split(Split(arrData(intDo), "|")(0), ";")(0)
        If intCol = intFindCol - cHideCols Then
            arrCol = Split(Split(arrData(intDo), "|")(1), ";")
            strPart = Split(arrCol(intItem), ",")(1)
            Exit For
        End If
    Next
    GetActivePart = strPart
End Function

Private Function CalcCollect(ByVal lngItem As Long, ByVal strStart As String, ByVal strEnd As String) As String
    Dim strCollect As String
    On Error GoTo ErrHand
    
    gstrSQL = " SELECT  SUM(记录内容) AS 汇总" & _
              " From 病人护理明细 A,病人护理数据 B," & vbNewLine & _
              "      (Select 序号 From 护理汇总项目 Start With 序号=[2] Connect By Prior 序号=父序号) C" & vbNewLine & _
              " Where A.记录ID=B.ID And A.终止版本 Is NULL And A.记录类型=1 AND B.汇总类别=0 And A.项目序号=C.序号" & vbNewLine & _
              " And B.文件ID=[1] And B.发生时间 Between [3] And [4]"
    Call SQLDIY(gstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "汇总数据", mlng文件ID, lngItem, CDate(strStart), CDate(strEnd))
    strCollect = NVL(rsTemp!汇总)
    
    CalcCollect = strCollect
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function CaclCategorical(ByVal strItem As String, ByVal strStart As String, ByVal strEnd As String) As ADODB.Recordset
'--------------------------------------------------------
'功能:获取汇总项目分类汇总信息
'参数:
'    strItem:分类项目和汇总项目的项目序号,格式:6:7
'    strStart:汇总开始时间
'    strEnd:汇总结束时间
'--------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    
    gstrSQL = "Select /*+ Rule */ 发生时间, 项目名称, 项目内容" & vbNewLine & _
        " From (Select Min(发生时间) 发生时间, 项目名称, Sum(项目内容) 项目内容" & vbNewLine & _
        "       From (Select 发生时间, Max(项目名称) 项目名称, Nvl(Max(项目内容), 0) 项目内容" & vbNewLine & _
        "              From (Select a.Id, a.发生时间, Decode(b.项目序号, d.C1, b.记录内容, Null) 项目名称, Decode(b.项目序号, d.C2, b.记录内容, Null) 项目内容" & vbNewLine & _
        "                     From 病人护理数据 a, 病人护理明细 b, 病人护理打印 c," & vbNewLine & _
        "                          (Select C1, C2 From Table(Cast(f_Num2list2([2]) As Zltools.t_Numlist2))) d" & vbNewLine & _
        "                     Where a.Id = b.记录id And a.Id = c.记录id And Nvl(a.汇总类别, 0) = 0 And a.文件id = [1] And" & vbNewLine & _
        "                           (b.项目序号 = d.C1 Or b.项目序号 = d.C2) And b.终止版本 Is NULL And b.记录类型=1 And " & vbNewLine & _
        "                           a.发生时间 Between [3] And [4])" & vbNewLine & _
        "              Group By Id, 发生时间)" & vbNewLine & _
        "       Group By 项目名称)" & vbNewLine & _
        " Order By 发生时间"
    Call SQLDIY(gstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "分类汇总数据", mlng文件ID, strItem, CDate(strStart), CDate(strEnd))
    Set CaclCategorical = rsTemp
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub cmdOK_Click()
    Dim arrTime, arrColCollect '要汇总的咧
    Dim arrItem, arrData
    Dim intType As Integer      '小结类别
    Dim arrValue, arrCorrelative() As String
    Dim bln跨天 As Boolean, blnExit As Boolean
    Dim lngStart As Long
    Dim lngCol As Long, lngCount As Long, lngRow As Long, lngRows As Long, i As Long, j As Long
    Dim strToday As String, str发生时间 As String
    Dim strStartDate As String, strEndDate As String
    Dim strStartTime As String, strEndTime As String
    Dim strKey As String, strField As String, strValue As String, strtmp As String
    Dim lngMaxIndex As Long, intDatas As Integer
    Dim rsCategorical As New ADODB.Recordset, lngMaxMutilRows As Long, lngMutilRows As Long, arrMutilRows
    
    On Error GoTo ErrHand
    '产生一条新的汇总记录
    
    If cbo小结.Text = "临时" And Val(txt小结名称.Tag) = 0 Then
        RaiseEvent AfterRowColChange("开始时点或结束时点格式不正确！", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    
    If InStr(1, txt小结名称.Text, ";") <> 0 Then
        RaiseEvent AfterRowColChange("小结名称中不能含有分号！", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    If InStr(1, txt小结名称.Text, "'") <> 0 Then
        RaiseEvent AfterRowColChange("小结名称中不能含有单引号！", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    If LenB(StrConv(txt小结名称.Text, vbFromUnicode)) > 50 Then
        RaiseEvent AfterRowColChange("小结名称不能超过50个字符或25个汉字！", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    
    '至少需要选择一个汇总项目
    arrColCollect = Array()
    With vfgItemList
        For lngCol = .FixedCols To .Cols - 1
            If Val(.Cell(flexcpData, .FixedRows, lngCol)) = 1 Then
                ReDim Preserve arrColCollect(UBound(arrColCollect) + 1)
                arrColCollect(UBound(arrColCollect)) = .ColData(lngCol)
            End If
        Next lngCol
    End With
    
    If UBound(arrColCollect) < 0 Then
        RaiseEvent AfterRowColChange("至少要选择一个汇总列！", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    '检查时间范围是否跨天
    '以指定的日期为准
    '    白 今天
    '    夜 今天 - 明天
    '    全 今天 - 明天
    strToday = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    arrTime = Split(cbo小结.Tag, ";")   '格式:开始时间,结束时间;开始时间,结束时间
    strStartTime = txt开始时点.Text
    strEndTime = txt结束时点.Text
    If strEndTime < strStartTime Then bln跨天 = True
    If bln跨天 = True Then
        strStartDate = strToday & " " & strStartTime & ":00"
        strEndDate = DateAdd("d", 1, CDate(strToday)) & " " & strEndTime & IIf(cbo小结.Text <> "临时", ":59", ":00")
    Else
        strStartDate = strToday & " " & strStartTime & ":00"
        strEndDate = strToday & " " & strEndTime & IIf(cbo小结.Text <> "临时", ":59", ":00")
    End If

    strStartDate = Format(DateAdd("d", -1 * DateDiff("d", CDate(DTPDate.Value), CDate(strToday)), CDate(strStartDate)), "yyyy-MM-dd HH:mm:ss")
    strEndDate = Format(DateAdd("d", -1 * DateDiff("d", CDate(DTPDate.Value), CDate(strToday)), CDate(strEndDate)), "yyyy-MM-dd HH:mm:ss")
    
    '63765:刘鹏飞,2013-11-22,修正小结保存的时间
    lngMaxIndex = cbo小结.ListCount
    If cbo小结.Text <> "临时" Then
        '小结发生时间等于小结结束时间－小结索引
        intType = -1 * cbo小结.ItemData(cbo小结.ListIndex)
        If cbo小结.ItemData(cbo小结.ListIndex) = 999 Then '全天小结-1s
            str发生时间 = Format(DateAdd("s", -1, strEndDate), "YYYY-MM-DD HH:mm:ss")
        Else
            str发生时间 = Format(DateAdd("s", -1 * (lngMaxIndex - cbo小结.ListIndex), strEndDate), "YYYY-MM-DD HH:mm:ss")
        End If
    Else
        '临时小结为-998
        intType = -1 * 998
        '55892:刘鹏飞,2012-11-30,临时小结结束时间-1S，如：8:00-18:00 指的就是汇总8点到17:59:59的
        strEndDate = Format(DateAdd("s", -1, strEndDate), "YYYY-MM-DD HH:mm:ss")
        str发生时间 = strEndDate
        strEndTime = Format(strEndDate, "HH:mm")
    End If
    
    
    '检查是否已经存在该数据
    blnExit = False
    mrsDataMap.Filter = "删除=0 And 汇总类别=" & intType & " And 汇总日期='" & str发生时间 & "'"    '记录ID>0的数据,都是当天的数据
    blnExit = (mrsDataMap.RecordCount)
    mrsDataMap.Filter = 0
    
    If blnExit Then
        RaiseEvent AfterRowColChange("您要添加的小结数据已存在！", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    
    If CDate(Format(strEndDate, "YYYY-MM-DD HH:mm:ss")) < CDate(Format(mstr开始时间, "YYYY-MM-DD HH:mm:ss")) Then
       RaiseEvent AfterRowColChange("小结的结束时间不能小于文件开始时间:[" & CDate(Format(mstr开始时间, "YYYY-MM-DD HH:mm:ss")) & "]", True, mblnSign, mblnArchive)
       Exit Sub
    End If
    
    '查找空白行
    lngRows = VsfData.Rows - 1
    For lngRow = VsfData.FixedRows To lngRows
        If Val(VsfData.TextMatrix(lngRow, mlngRecord)) = 0 And VsfData.TextMatrix(lngRow, mlngRowCount) = "" Then
            lngStart = lngRow
            Exit For
        End If
    Next
    If lngStart = 0 Then
        '说明没有找到空白行
        VsfData.Rows = VsfData.Rows + 1
        lngStart = VsfData.Rows - 1
    End If
    
    '统计汇总数据(从数据库中汇总,当前数据只记录了是否修改,并不知道原值是多少,所以当前未保存的数据不汇总)
    '汇总项目集合
    '汇总项目列集合:col;1|col;4,5
    arrValue = Array()
    arrItem = Split(mstrColCollect, "|")
    For i = 0 To UBound(arrItem)
        If InStr(1, "," & Join(arrColCollect, ",") & ",", "," & Split(arrItem(i), ";")(0) & ",") > 0 Then
            arrData = Split(Split(arrItem(i), ";")(1), ",")
            For j = 0 To UBound(arrData)
                ReDim Preserve arrValue(UBound(arrValue) + 1)
                arrValue(UBound(arrValue)) = CalcCollect(arrData(j), strStartDate, strEndDate)
            Next j
        End If
    Next i
    
    '通用部分
    VsfData.TextMatrix(lngStart, mlngYear) = txt小结名称.Text
    VsfData.TextMatrix(lngStart, mlngDate) = txt小结名称.Text
    VsfData.TextMatrix(lngStart, mlngTime) = txt小结名称.Text
    VsfData.TextMatrix(lngStart, mlngRowCount) = "1|1"                          '为了保证时间不重复,采取结束时间+秒的方式
    VsfData.TextMatrix(lngStart, mlngRowCurrent) = "1"
    VsfData.TextMatrix(lngStart, mlngCollectText) = txt小结名称.Text
    VsfData.TextMatrix(lngStart, mlngCollectType) = intType                     '表示小结;-1白班;-2夜班;3-全天
    VsfData.TextMatrix(lngStart, mlngCollectStyle) = cbo标识.ListIndex         '不足24小时,上下划红线
    VsfData.TextMatrix(lngStart, mlngCollectDay) = str发生时间
    VsfData.TextMatrix(lngStart, mlngCollectStart) = strStartTime
    VsfData.TextMatrix(lngStart, mlngCollectEnd) = strEndTime
    VsfData.MergeRow(lngStart) = True
    '同步保存日期与时间列的数据
    strField = "ID|页号|行号|列号|起始行号|记录ID|数据|汇总|记录组号|删除"
    '1\日期
    strKey = mint页码 & "," & lngStart & "," & mlngDate
    strValue = strKey & "|" & mint页码 & "|" & lngStart & "|" & mlngDate & "|" & lngStart & "|" & Val(VsfData.TextMatrix(lngStart, mlngRecord)) & "|" & _
            txt小结名称.Text & ";" & intType & ";" & cbo标识.ListIndex & ";" & str发生时间 & ";" & strStartTime & ";" & strEndTime & "|1|0|0"
    Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
    
    '展现
    arrItem = Split(mstrColCollect, "|")
    lngCount = 0
    lngRows = UBound(arrItem)
    For lngRow = 0 To lngRows
        lngCol = Split(arrItem(lngRow), ";")(0)
        If InStr(1, "," & Join(arrColCollect, ",") & ",", "," & lngCol & ",") > 0 Then
            If UBound(Split(Split(arrItem(lngRow), ";")(1), ",")) = 1 Then
                strValue = arrValue(lngCount) & "/" & arrValue(lngCount + 1)
                lngCount = lngCount + 2
            Else
                strValue = arrValue(lngCount)
                lngCount = lngCount + 1
            End If
            
            VsfData.TextMatrix(lngStart, lngCol + cHideCols) = strValue
            
            '52953,刘鹏飞,2012-08-24,汇总数据为0也要显示,避免相邻数据列合并，关联问题:60792
            If lngCol + cHideCols > mlngTime And lngCol + cHideCols < mlngNoEditor Then
                VsfData.TextMatrix(lngStart, lngCol + cHideCols) = FormatValue(VsfData.TextMatrix(lngStart, lngCol + cHideCols))
                If Trim(VsfData.TextMatrix(lngStart, lngCol + cHideCols)) <> "" Then
                    '66085:刘鹏飞,2012-09-26,避免相邻汇总列合并,将原来的列内容+空格同一改成在列后面在chr(13)
                    '避免因加空格后列宽不够导致内容显示不完全(主要针对右对其)
    '                Select Case VsfData.ColAlignment(lngCol + cHideCols)
    '                    Case 6, 7, 8
    '                        VsfData.TextMatrix(lngStart, lngCol + cHideCols) = IIf((lngCol + cHideCols) Mod 2 = 1, " ", String(2, " ")) & VsfData.TextMatrix(lngStart, lngCol + cHideCols)
    '                    Case 3, 4, 5
    '                        VsfData.TextMatrix(lngStart, lngCol + cHideCols) = IIf((lngCol + cHideCols) Mod 2 = 1, " ", String(2, " ")) & VsfData.TextMatrix(lngStart, lngCol + cHideCols) & IIf((lngCol + cHideCols) Mod 2 = 1, " ", String(2, " "))
    '                    Case 0, 1, 2
    '                        VsfData.TextMatrix(lngStart, lngCol + cHideCols) = VsfData.TextMatrix(lngStart, lngCol + cHideCols) & IIf((lngCol + cHideCols) Mod 2 = 1, " ", String(2, " "))
    '                    Case Else
    '                        VsfData.TextMatrix(lngStart, lngCol + cHideCols) = IIf((lngCol + cHideCols) Mod 2 = 1, " ", String(2, " ")) & VsfData.TextMatrix(lngStart, lngCol + cHideCols)
    '                End Select
                    VsfData.TextMatrix(lngStart, lngCol + cHideCols) = VsfData.TextMatrix(lngStart, lngCol + cHideCols) & IIf((lngCol + cHideCols) Mod 2 = 1, Chr(13), "")
                End If
            End If
            
            strKey = mint页码 & "," & lngStart & "," & lngCol + cHideCols
            strValue = strKey & "|" & mint页码 & "|" & lngStart & "|" & lngCol + cHideCols & "|" & lngStart & "|" & _
                Val(VsfData.TextMatrix(lngStart, mlngRecord)) & "|" & VsfData.TextMatrix(lngStart, lngCol + cHideCols) & "|1|0|0"
            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
        End If
    Next
    '分类汇总数据处理
    arrMutilRows = Array()
    lngMaxMutilRows = 0: lngMutilRows = 0
    arrItem = Split(mstrColCorrelative, "|")
    For lngCount = 0 To UBound(arrItem)
        lngRow = lngStart + Val(VsfData.TextMatrix(lngStart, mlngRowCurrent))
        lngMaxMutilRows = 0
        ReDim Preserve arrMutilRows(UBound(arrMutilRows) + 1)
        arrMutilRows(UBound(arrMutilRows)) = lngMaxMutilRows
        arrCorrelative = Split(arrItem(lngCount), ";")
        If InStr(1, "," & Join(arrColCollect, ",") & ",", "," & Split(arrCorrelative(1), ",")(0) & ",") > 0 Then
            Set rsCategorical = CaclCategorical(Split(arrCorrelative(0), ",")(1) & ":" & Split(arrCorrelative(1), ",")(1), strStartDate, strEndDate)
            '存在分类汇总，总量汇总列关联的名称列填写名称"总量"
            If rsCategorical.RecordCount > 0 Then
                lngCol = Split(arrCorrelative(0), ",")(0) + cHideCols + VsfData.FixedCols - 1
                VsfData.TextMatrix(lngStart, lngCol) = "总量"
                strKey = mint页码 & "," & lngStart & "," & lngCol
                strValue = strKey & "|" & mint页码 & "|" & lngStart & "|" & lngCol & "|" & lngStart & "|" & _
                    Val(VsfData.TextMatrix(lngStart, mlngRecord)) & "|" & VsfData.TextMatrix(lngStart, lngCol) & "|1|0|0"
                Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
            End If
            With rsCategorical
                Do While Not .EOF
                    lngMutilRows = 0
                    For i = 0 To 1
                        lngRows = lngRow
                        lngCol = Split(arrCorrelative(i), ",")(0) + cHideCols + VsfData.FixedCols - 1
                        strtmp = IIf(i = 0, CStr(NVL(!项目名称)), CStr(NVL(!项目内容)))
                        With txtLength
                            .Width = VsfData.ColWidth(lngCol)
                            .Text = Replace(Replace(Replace(strtmp, Chr(10), ""), Chr(13), ""), Chr(1), "")
                            .FontName = VsfData.CellFontName
                            .FontSize = VsfData.CellFontSize
                            .FontBold = VsfData.CellFontBold
                            .FontItalic = VsfData.CellFontItalic
                        End With
                        arrData = GetData(txtLength.Text)
                        intDatas = UBound(arrData)
                        If intDatas < 0 Then ReDim arrData(0): intDatas = 0
                        If lngMutilRows < intDatas + 1 Then lngMutilRows = intDatas + 1
                        lngRows = lngRows + intDatas
                        If VsfData.Rows <= lngRows Then VsfData.Rows = VsfData.Rows + (lngRows - VsfData.Rows + 1)
                        For j = 0 To intDatas
                            VsfData.TextMatrix(lngRow + j, lngCol) = CStr(arrData(j))
                        Next j
                        
                        strKey = mint页码 & "," & lngRow & "," & lngCol
                        strValue = strKey & "|" & mint页码 & "|" & lngRow & "|" & lngCol & "|" & lngStart & "|" & _
                            Val(VsfData.TextMatrix(lngStart, mlngRecord)) & "|" & strtmp & "|1|" & .AbsolutePosition & "|0"
                        Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                    Next i
                    lngMaxMutilRows = lngMaxMutilRows + lngMutilRows
                    lngRow = lngRow + lngMutilRows
                    arrMutilRows(UBound(arrMutilRows)) = lngMaxMutilRows
                .MoveNext
                Loop
            End With
        End If
    Next
    '获取最大行高
    lngMaxMutilRows = 0
    For i = 0 To UBound(arrMutilRows)
        If lngMaxMutilRows < Val(arrMutilRows(i)) Then lngMaxMutilRows = Val(arrMutilRows(i))
    Next i
    If lngMaxMutilRows > 0 Then
        lngMaxMutilRows = lngMaxMutilRows + Val(VsfData.TextMatrix(lngStart, mlngRowCurrent))
        For lngRow = lngStart To lngStart + lngMaxMutilRows - 1
            VsfData.TextMatrix(lngRow, mlngRowCount) = lngMaxMutilRows & "|" & lngRow - lngStart + 1                       '为了保证时间不重复,采取结束时间+秒的方式
            VsfData.TextMatrix(lngRow, mlngRowCurrent) = lngMaxMutilRows
            VsfData.TextMatrix(lngRow, mlngCollectText) = VsfData.TextMatrix(lngStart, mlngCollectText)
            VsfData.TextMatrix(lngRow, mlngCollectType) = VsfData.TextMatrix(lngStart, mlngCollectType)
            VsfData.TextMatrix(lngRow, mlngCollectStyle) = VsfData.TextMatrix(lngStart, mlngCollectStyle)
            VsfData.TextMatrix(lngRow, mlngCollectDay) = VsfData.TextMatrix(lngStart, mlngCollectDay)
            VsfData.TextMatrix(lngRow, mlngCollectStart) = VsfData.TextMatrix(lngStart, mlngCollectStart)
            VsfData.TextMatrix(lngRow, mlngCollectEnd) = VsfData.TextMatrix(lngStart, mlngCollectEnd)
        Next lngRow
    End If
    
'    '合并单元格
'    lngRows = Split(Split(mstrColCollect, "|")(0), ";")(0) + cHideCols - 1
'    For lngRow = mlngTime + 1 To lngRows
'        VsfData.TextMatrix(lngStart, lngRow) = txt小结名称.Text
'    Next
'    VsfData.MergeCells = flexMergeRestrictRows          '冻结单元格竟然是单独合并,合并后会有两个合并单元格
'    VsfData.MergeRow(lngStart) = True
    mblnChange = True
    picBiref.Visible = False
    RaiseEvent AfterDataChanged(mblnChange)
 
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdWord_Click()
    Dim strInput As String
    '弹出词句选择器
    
    If Val(cmdWord.Tag) = -1 Then
        strInput = txtInput.Text
    Else
        strInput = txt(Val(cmdWord.Tag)).Text
    End If
    strInput = frmEditAssistant.ShowMe(Me, mlng病人ID, mlng主页ID, mint婴儿, strInput)
    
    If Val(cmdWord.Tag) = -1 Then
        txtInput.Text = strInput
        Call txtInput_KeyDown(vbKeyReturn, 0)
    Else
        txt(Val(cmdWord.Tag)).Text = strInput
        Call txt_KeyDown(Val(cmdWord.Tag), vbKeyReturn, 0)
    End If
End Sub

Private Sub ShowBrief()
    Dim strStart As String, strEnd As String, strMaxTime As String
    Dim strHave As String, strDate As String
    Dim strTag As String    'cbo小结的tag中保存时间段，格式：开始,结束;开始,结束
    Dim rsData As New ADODB.Recordset
    Dim i As Integer
    Dim strCurDate As String
    Dim intStart As Integer, intEnd As Integer, intCur As Integer, intIndex As Integer, intCount As Integer
    On Error GoTo ErrHand
    '显示小结窗体
    
    If Not DataMap_Save Then Exit Sub       '保存数据,以便选择小结的时候进行数据检查
    '本记录单是否存在汇总项目列，如果不存在则退出
    If mstrCollectItems = "" Then
        RaiseEvent AfterRowColChange("当前文件中未使用汇总项目！", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    
    '提取汇总时段(类别=3为全天小结)
    strDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    gstrSQL = "Select NVL(科室ID,0) AS 科室ID,单据,类别,名称,开始,结束 From 护理汇总时段 Order by 类别 "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取小结")
    rsTemp.Filter = "单据=2"
    If rsTemp.RecordCount = 0 Then
        rsTemp.Filter = 0
        RaiseEvent AfterRowColChange("还未设置记录单小结,请先在护理项目管理模块的汇总项目中设置！", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    
    rsTemp.Filter = "单据=1 And 类别=3"
    If rsTemp.RecordCount = 0 Then
        rsTemp.Filter = 0
        RaiseEvent AfterRowColChange("全天汇总时段未设置,请先在护理项目管理模块的汇总项目中设置！", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    strStart = NVL(rsTemp!开始)
    strEnd = NVL(rsTemp!结束)
    rsTemp.Filter = 0
    
    strCurDate = Format(zlDatabase.Currentdate, "YYYY-MM-DD")
    
    On Error Resume Next
    With DTPDate
        '问题号:50201 LPF 取消原来限制的一个月时间，小结时间可以在文件当前的有效时间范围即可
        '.MinDate = Format(DateAdd("D", -30, CDate(strCurDate)), "YYYY-MM-DD")
        .MinDate = Format(mstr开始时间, "YYYY-MM-DD")
        If CDate(.MinDate) < CDate(Format(mstr开始时间, "YYYY-MM-DD")) Then .MinDate = Format(mstr开始时间, "YYYY-MM-DD")
        '提取病人变动记录中的最大时间
        gstrSQL = "SELECT MAX(NVL(终止时间, SYSDATE+" & mintPreDays & ")) AS 出院时间" & vbNewLine & _
            " FROM 病人变动记录" & vbNewLine & _
            " WHERE 开始时间 IS NOT NULL AND 病人ID =[1] AND 主页ID =[2]"
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "提取病人最大时间", mlng病人ID, mlng主页ID)
        If rsData.RecordCount > 0 Then
            strMaxTime = Format(rsData!出院时间, "YYYY-MM-DD")
        Else
            strMaxTime = Format(DateAdd("D", mintPreDays, CDate(strCurDate)), "YYYY-MM-DD")
        End If
        '文件结束时间不为空，最大日期范围不能操作病人当前有效最大时间
        If IsDate(mstr结束时间) Then
            If CDate(Format(strMaxTime, "YYYY-MM-DD")) > CDate(Format(mstr结束时间, "YYYY-MM-DD")) Then
                strMaxTime = Format(mstr结束时间, "YYYY-MM-DD")
            End If
        End If
        If CDate(Format(strMaxTime, "YYYY-MM-DD")) < CDate(Format(mstr开始时间, "YYYY-MM-DD")) Then
            strMaxTime = Format(mstr开始时间, "YYYY-MM-DD")
        End If
        .MaxDate = Format(strMaxTime, "YYYY-MM-DD")
        strMaxTime = Format(DateAdd("D", -1, CDate(strCurDate)), "YYYY-MM-DD")
        
        If CDate(Format(strMaxTime, "YYYY-MM-DD")) < CDate(Format(.MinDate, "YYYY-MM-DD")) Then
            strMaxTime = Format(.MinDate, "YYYY-MM-DD")
        End If
        If CDate(Format(strMaxTime, "YYYY-MM-DD")) > CDate(Format(.MaxDate, "YYYY-MM-DD")) Then
            strMaxTime = Format(.MaxDate, "YYYY-MM-DD")
        End If
        .Value = Format(strMaxTime, "YYYY-MM-DD")
    End With
    
    '加载汇总类别(记录单小结最多3个)
    intIndex = 0
    intCount = 0
    intCur = Format(Now, "HH")
    cbo小结.Clear
    If strStart <> "" Or strEnd <> "" Then
        cbo小结.AddItem "全天小结"
        cbo小结.ItemData(cbo小结.NewIndex) = 999
        strTag = strTag & ";" & strStart & "," & strEnd
    End If
    intCount = intCount + 1
    
    With rsTemp
        rsTemp.Filter = "单据 = 2 And 科室ID=" & mlng病区ID
        If rsTemp.RecordCount = 0 Then rsTemp.Filter = "单据 = 2 And 科室ID=0"
        rsTemp.Sort = "开始 ASC"
        Do While Not .EOF
            If Not (NVL(!开始) = "" Or NVL(!结束) = "") Then
                cbo小结.AddItem !名称
                cbo小结.ItemData(cbo小结.NewIndex) = Val(!类别)
                strTag = strTag & ";" & !开始 & "," & !结束
                
                '定位当前时点对应的小结
                intStart = Val(!开始)
                intEnd = Val(!结束)
                If intStart <= intEnd Then
                    If intCur >= intStart And intCur <= intEnd Then intIndex = intCount
                Else
                    If intCur >= intStart Then intIndex = intCount
                End If
            End If
            
            intCount = intCount + 1
            .MoveNext
        Loop
        If strTag <> "" Then
            cbo小结.Tag = Mid(strTag, 2)
            cbo小结.ListIndex = intIndex
        Else
            rsTemp.Filter = 0
            RaiseEvent AfterRowColChange("当天的汇总已全部添加！", True, mblnSign, mblnArchive)
            Exit Sub
        End If
        cbo小结.AddItem "临时"
    End With
    
    
    rsTemp.Filter = 0
    
    With cbo标识
        .Clear
        .AddItem "不处理"
        .AddItem "上下画横线标识"
        .AddItem "汇总值下方画双横线标识"
        .AddItem "上方画横线标识"
        .AddItem "汇总值下方画单横线标识"
        If mintCollectDef > 0 And mintCollectDef < .ListCount Then
            .ListIndex = mintCollectDef
        Else
            .ListIndex = 0
        End If
    End With
    
    Call LoadCollectItem
    '设置坐标
    cmdOk.Top = vfgItemList.Top + vfgItemList.Height + 120
    cmdCancel.Top = cmdOk.Top
    lblCollectInfo.Top = cmdOk.Top + 80
    With picBiref
        .Top = VsfData.Top + picMain.Top + 20 ' 200
        .Left = (ScaleWidth - .Width) / 2
        .Height = cmdOk.Top + cmdOk.Height + 100
        If .Top + .Height > .Top + VsfData.Height Then
            .Top = .Top + VsfData.Height - .Height
            If .Top < 0 Then .Top = 0
        End If
        .Visible = True
    End With
    
    On Error Resume Next
    cbo小结.SetFocus
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub LoadCollectItem()
    Dim arrItem() As String, strCols As String, strCell As String
    Dim lngStartCol As Long, lngRow As Long, lngCol As Long
    Dim lngCount As Long
    Dim lngHeight As Long, lngWidth As Long
    '设置列头
    With vfgItemList
        .Clear
        arrItem = Split(mstrColCollect, "|")
        .Cols = UBound(arrItem) + 2
        .Rows = 4
        .FixedRows = 3
        .FixedCols = 1
        .MergeCells = flexMergeFixedOnly
        .MergeCellsFixed = flexMergeFree
        .MergeRow(0) = True
        .MergeRow(1) = True
        .MergeRow(2) = True
        
        For lngCol = 0 To .Cols - 1
            .MergeCol(lngCol) = True
        Next
        
        .TextMatrix(0, 0) = "汇总列"
        .TextMatrix(1, 0) = "汇总列"
        .TextMatrix(2, 0) = "汇总列"
        .TextMatrix(3, 0) = "是否汇总"
        
        strCols = ""
        For lngCount = 0 To UBound(arrItem)
            strCols = strCols & "," & Split(arrItem(lngCount), ";")(0)
        Next lngCount
        strCols = Mid(strCols, 2)
        arrItem = Split(mstrTabHead, "|")
    
        For lngCount = 0 To UBound(arrItem)
            strCell = arrItem(lngCount)
            lngRow = Left(strCell, InStr(1, strCell, ",") - 1): strCell = Mid(strCell, InStr(1, strCell, ",") + 1)
            lngCol = Left(strCell, InStr(1, strCell, ",") - 1): strCell = Mid(strCell, InStr(1, strCell, ",") + 1)
            If InStr(1, "," & strCols & ",", "," & lngCol & ",") <> 0 Then
                For lngStartCol = 0 To UBound(Split(strCols, ","))
                    If Split(strCols, ",")(lngStartCol) = lngCol Then
                        .TextMatrix(lngRow, lngStartCol + .FixedCols) = strCell
                        .ColData(lngStartCol + .FixedCols) = lngCol
                    End If
                Next lngStartCol
            End If
        Next
        Set .Cell(flexcpPicture, .FixedRows, .FixedCols, .FixedRows, .Cols - 1) = imgTrueFalse.ListImages("T").Picture
        .Cell(flexcpData, .FixedRows, .FixedCols, .FixedRows, .Cols - 1) = 1
        
        Select Case mintTabTiers
        Case 1
            .RowHidden(0) = False
            .RowHidden(1) = True
            .RowHidden(2) = True
        Case 2
            .RowHidden(0) = False
            .RowHidden(1) = False
            .RowHidden(2) = True
        Case 3
            .RowHidden(0) = False
            .RowHidden(1) = False
            .RowHidden(2) = False
        End Select
        
        .RowHeight(-1) = 255

        '可能固定行的行高不正确需要自动调整下
        .AutoResize = True
        .WordWrap = True
        .AutoSizeMode = flexAutoSizeRowHeight
        .AutoSize 0, .Cols - 1
        .AutoResize = False
        
        .Cell(flexcpAlignment, 0, .FixedCols, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
        .Cell(flexcpPictureAlignment, .FixedRows, .FixedCols, .FixedRows, .Cols - 1) = flexAlignCenterCenter
        
        '87428:根据行的高度重新调整表格高度，因为当固定行的高度+横向滚动条的高度大于表格高度，纵向滚动条无法显示
        '因表格是存在一行可编辑，目前处理为完全显示看到表格所有行
        .Height = 1100 '默认初始高度
        lngHeight = 0
        For lngRow = 0 To .Rows - 1
            If .RowHidden(lngRow) = False Then lngHeight = lngHeight + .RowHeight(lngRow)
        Next
        lngWidth = 0
        For lngCol = 0 To .Cols - 1
            lngWidth = lngWidth + .ColWidth(lngCol)
        Next
        
        If lngWidth > .Width Then
            If lngHeight + 350 > .Height Then .Height = lngHeight + 350
        Else
            If lngHeight + 100 > .Height Then .Height = lngHeight + 100
        End If
    End With
End Sub

Private Sub DTPDate_Change()
    cbo小结_Click
End Sub

Private Sub DTPDate_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyReturn Then cbo小结.SetFocus
End Sub

Private Sub lstSelect_Click(Index As Integer)
    If Index = 0 Then
        If PicLst.Visible = False Then Exit Sub
        If lstSelect(0).ListIndex > 0 Then
            txtLst.Text = lstSelect(0).Text
        End If
    End If
End Sub


Private Sub picImg_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strInfo As String
    Dim arrCode, i As Long
    strInfo = lblInfo.Tag
    arrCode = Split(strInfo, "[LPF]")
    If UBound(arrCode) = 0 Then
        strInfo = CStr(arrCode(i))
    Else
        For i = 0 To UBound(arrCode)
            If i = 0 Then
                strInfo = i + 1 & "、" & CStr(arrCode(i))
            Else
                strInfo = strInfo & vbCrLf & i + 1 & "、" & CStr(arrCode(i))
            End If
        Next
    End If
    If UBound(arrCode) >= 0 Then
        strInfo = strInfo & vbCrLf & vbCrLf & "说明：本提示信息的准确性依赖于数据已完全保存。"
    End If
    Call zlCommFun.ShowTipInfo(picImg.hWnd, strInfo, True)
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

Private Sub txtFind_Change()
    Call txtFind_KeyDown(10, 0)
End Sub

Private Sub txtFind_GotFocus()
    txtFind.SelStart = 0
    txtFind.SelLength = 100
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    Static lngPreIndex As Long
    Dim strText As String
    Dim lngIndex As Long
    
    '61855:刘鹏飞,2013-11-07,绑定活动项目怎么加搜索功能
    strText = Trim(txtFind.Text)
    If KeyCode = 10 Or strText = "" Then
        '主要是用于清除变量值
        lngPreIndex = 0
    ElseIf KeyCode = vbKeyReturn And strText <> "" Then
        If Not (lngPreIndex > 0 And lngPreIndex < lstColumnItems.ListItems.Count) Then lngPreIndex = 1
        For lngIndex = lngPreIndex To lstColumnItems.ListItems.Count
            If UCase(lstColumnItems.ListItems(lngIndex).SubItems(1)) Like UCase(strText) & "*" Then
                lstColumnItems.ListItems(lngIndex).Selected = True
                lstColumnItems.ListItems(lngIndex).EnsureVisible
                Exit For
            End If
        Next
        
        If lngIndex > lstColumnItems.ListItems.Count Then
            If lngPreIndex > 1 Then
                For lngIndex = 1 To lstColumnItems.ListItems.Count
                    If UCase(lstColumnItems.ListItems(lngIndex).SubItems(1)) Like UCase(strText) & "*" Then
                        lstColumnItems.ListItems(lngIndex).Selected = True
                        lstColumnItems.ListItems(lngIndex).EnsureVisible
                        Exit For
                    End If
                Next
            End If
            lngPreIndex = 1
        Else
            lngPreIndex = lngIndex + 1
        End If
    End If
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
    If KeyAscii = 58 And VsfData.COL = mlngTime Then KeyAscii = 0
End Sub

Private Sub txtLst_Change()
    Dim intRows As Integer
    Dim lngHeight As Long, lngCurHeight As Long, lngDiffHeight As Long
    
    If PicLst.Visible = False Then Exit Sub
    '获取输入的行数
    intRows = SendMessage(txtLst.hWnd, EM_GETLINECOUNT, 0&, 0&)
    lngCurHeight = PicLst.TextHeight("高") * intRows + PicLst.TextHeight("高") * 1 / 3
    lngHeight = txtLst.Height
    lngDiffHeight = lngCurHeight - lngHeight
    If lngCurHeight < Val(txtLst.Tag) Then lngCurHeight = Val(txtLst.Height)
    txtLst.Height = lngCurHeight
    lbllst(1).Top = txtLst.Top + txtLst.Height
    lstSelect(mintType - 1).Top = lbllst(1).Top + lbllst(1).Height + 20
    PicLst.Height = lstSelect(mintType - 1).Top + PicLst.TextHeight("高") * (lstSelect(mintType - 1).ListCount) + PicLst.TextHeight("高") \ 3
    If PicLst.Top + PicLst.Height + picMain.Top > ScaleHeight Then
        If ScaleHeight - PicLst.Top - picMain.Top < 0 Then
            PicLst.Top = 10
            PicLst.Height = ScaleHeight - picMain.Top - 10
        Else
            PicLst.Height = ScaleHeight - picMain.Top
        End If
    End If
    lstSelect(mintType - 1).Height = IIf(PicLst.Height - lstSelect(mintType - 1).Top < PicLst.TextHeight("刘") + PicLst.TextHeight("刘") \ 3, PicLst.TextHeight("刘") + PicLst.TextHeight("刘") \ 3, PicLst.Height - lstSelect(mintType - 1).Top)
    If lstSelect(mintType - 1).Top + lstSelect(mintType - 1).Height <> PicLst.Height Then
        PicLst.Height = lstSelect(mintType - 1).Top + lstSelect(mintType - 1).Height
    End If
End Sub

Private Sub txtLst_GotFocus()
    mblnEditAssistant = False
    mblnEditText = False
    PicLst.Tag = 1
    Call zlControl.TxtSelAll(txtLst)
    lstSelect(0).ListIndex = -1
End Sub

Private Sub txtLst_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And Shift = vbCtrlMask Then Exit Sub
    If KeyCode = vbKeyReturn Or _
        (KeyCode = vbKeyRight And txtLst.SelStart = Len(txtLst.Text)) Or _
        (KeyCode = vbKeyLeft And txtLst.SelStart = 0) Then
        '移动到下一个单元格
        Call MoveNextCell(Not (KeyCode = vbKeyLeft))
    End If
    If Shift = vbShiftMask And KeyCode = vbKeyDown Then
        KeyCode = 0
        lstSelect(0).SetFocus
    End If
End Sub

Private Sub txtPage_GotFocus()
    txtPage.SelStart = 0
    txtPage.SelLength = 3
End Sub

Private Sub txtPage_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Val(txtPage.Text) < mint起始页码 Or Val(txtPage.Text) > mint结束页 + 1 Then
        MsgBox "输入的页码无效，当前文件页码有效页码范围：第" & mint起始页码 & "页 ～ 第" & mint结束页 + 1 & "页！", vbInformation, gstrSysName
        txtPage.SetFocus
        Exit Sub
    End If
    
    If Not DataMap_Save Then Exit Sub
    
    '更新查询SQL
    '重新提取数据
    mint页码 = Val(txtPage.Text)
    mblnInit = False
    Call InitVariable
    Call InitCons
    Call ReadStruDef
    Call zlRefresh
    
    mblnInit = True
    VsfData.Refresh
    
    mcbrPage.Caption = "页码选择：第" & mint页码 & "页"
    cbsThis.RecalcLayout
    
End Sub

Private Sub txtPage_KeyPress(KeyAscii As Integer)
    If (Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And (KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack)) Then KeyAscii = 0
End Sub

Private Sub vfgItemList_DblClick()
    Dim intValue As Integer
    With vfgItemList
        If Not (.Rows > .FixedRows) Or Not (.Cols > .FixedCols) Then Exit Sub
        If .ROW >= .FixedRows And .COL >= .FixedCols Then
            intValue = Val(.Cell(flexcpData, .ROW, .COL))
            Set .Cell(flexcpPicture, .ROW, .COL) = imgTrueFalse.ListImages(IIf(intValue = 1, "F", "T")).Picture
            .Cell(flexcpData, .ROW, .COL) = IIf(intValue = 1, 0, 1)
        End If
    End With
End Sub

Private Sub VsfData_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Dim lngRow As Long, lngCol As Long
    Dim dblHeight As Double, dblWidth As Double
    Dim strItemInfo As String
    If Not mblnInit Then Exit Sub
    Call InitCons
    
    If OldLeftCol <> NewLeftCol Then
        vsfHead.LeftCol = NewLeftCol
        VsfData.LeftCol = vsfHead.LeftCol
    End If
    
    If mblnEditable = False Then Exit Sub
    '当列不可见时不显示说明信息
    If VsfData.RowIsVisible(VsfData.ROW) = True And VsfData.ColIsVisible(VsfData.COL) = True Then
         '显示当前项目的相关信息
        mrsSelItems.Filter = "列=" & VsfData.COL - (cHideCols + VsfData.FixedCols - 1)
        If mrsSelItems.RecordCount <> 0 Then
            mrsItems.Filter = "项目序号=" & mrsSelItems!项目序号
            If mrsItems.RecordCount <> 0 Then
                strItemInfo = Trim(NVL(mrsItems!说明, ""))
            End If
        End If
        mrsSelItems.Filter = 0
        mrsItems.Filter = 0
    End If
     '--48659:刘鹏飞,2012-09-14,添加字段'说明'
    RaiseEvent ShowTipInfo(VsfData, strItemInfo, True)
'    '计算固定行的高度
'    For lngRow = 0 To 2
'        If Not VsfData.RowHidden(lngRow) Then dblHeight = dblHeight + VsfData.ROWHEIGHT(lngRow)
'    Next
'    '从可见行开始向下查找最后一个可见行
'    For lngRow = NewTopRow To VsfData.Rows - 1
'        If Not VsfData.RowIsVisible(lngRow) Then
'            lngRow = lngRow - 1
'            Exit For
'        End If
'    Next
'    '从可见列开始查找最后一个可见列
'    For lngCol = NewLeftCol To VsfData.Cols - 1
'        If Not VsfData.ColIsVisible(lngCol) Then
'            lngCol = lngCol - 1
'            Exit For
'        Else
'            dblWidth = dblWidth + VsfData.ColWidth(lngCol)
'        End If
'    Next
'
'    If Not VsfData.RowIsVisible(VsfData.Row) Then
'        VsfData.Row = VsfData.Row + IIf(OldTopRow < NewTopRow, 1, -1)
'    Else
'        '当前数据行的高度+固定行的高度如果大于表格控件的高度,说明当前选择的数据行存在遮住部分的情况
'        If VsfData.Row >= lngRow - 1 And CellRect.Bottom * (lngRow - NewTopRow + 1) + dblHeight >= VsfData.ClientHeight Then
'            '遮住部分的情况下
'            VsfData.Row = VsfData.Row + IIf(OldTopRow < NewTopRow, 1, -1)
'        End If
'    End If
'
'    If Not VsfData.ColIsVisible(VsfData.Col) Then
'        VsfData.Col = VsfData.Col + IIf(OldLeftCol < NewLeftCol, 1, -1)
'    Else
'        '当前数据行的高度+固定行的高度如果大于表格控件的高度,说明当前选择的数据行存在遮住部分的情况
'        If VsfData.Col = lngCol And dblWidth >= VsfData.ClientWidth Then
'            '遮住部分的情况下
'            VsfData.Col = VsfData.Col + IIf(OldLeftCol < NewLeftCol, 1, -1)
'        End If
'    End If
'
'    Call VsfData_EnterCell
End Sub

Private Sub VsfData_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    Dim blnResult As Boolean
    If Not mblnInit Then Exit Sub
    If mintType = -1 Then Exit Sub
    blnResult = MoveNextCell(True, True)
    Cancel = Not blnResult
End Sub

Private Sub VsfData_DblClick()
    Call vsfdata_KeyDown(Asc("A"), 0)
End Sub

Private Sub VsfData_EnterCell()
    Dim strCols As String
    Dim strName As String
    Dim intMax As Integer
    Dim lngStart As Long
    Dim strDate As String, strYear As String
    Dim strCorrelative As String, arrCorrelative, i As Long
    Dim blnCheck As Boolean
    '隐蔽已显示的录入控件
    Select Case mintType
    Case 0, 3
        picInput.Visible = False
    Case 1, 2
        lstSelect(mintType - 1).Visible = False
        If mintType = 1 Then
            txtLst.Visible = False
            PicLst.Visible = False
        End If
    Case 4, 5
        picDouble.Visible = False
    Case 6
        picMutilInput.Visible = False
    Case 7
        picDoubleChoose.Visible = False
    Case 8
        picYear.Visible = False
    End Select
    
    cmdWord.Visible = False
    
    '未定义的列不允许录入数据
    mintType = -1
    
    If mblnInit = False Then Exit Sub
    
    Call ShowSignMarker
    
    If InStr(1, mstrPrivs, "护理记录登记") = 0 Then Exit Sub
    
    If VsfData.TextMatrix(VsfData.ROW, mlngRowCount) = "" Then
        lngStart = VsfData.ROW
    Else
        lngStart = GetStartRow(VsfData.ROW)
    End If
    
    '分类汇总关联的名称列确定
    arrCorrelative = Split(mstrColCorrelative, "|")
    For i = 0 To UBound(arrCorrelative)
        strCorrelative = strCorrelative & "," & Val(Split(arrCorrelative(i), ",")(0))
    Next i
    strCorrelative = Mid(strCorrelative, 2)
    If Val(VsfData.TextMatrix(lngStart, mlngCollectType)) < 0 And _
        (VsfData.COL <= mlngTime And IIf(mblnVerify = True, VsfData.COL <> mlngChoose, True) Or _
        InStr(1, "|" & mstrColCollect, "|" & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ";") <> 0 Or _
        InStr(1, "," & strCorrelative & ",", "," & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ",") <> 0) Then
        RaiseEvent AfterRowColChange("汇总行不允许修改日期时间,以及汇总列和汇总列关联列的数据！", True, mblnSign, mblnArchive)
        Exit Sub '汇总行不允许修改日期时间,以及汇总列和汇总列关联名称列的数据
    End If
    
    If mblnVerify Then  '必须放在mblnShow判断语句的上面
        If VsfData.COL = mlngChoose Then Call vsfdata_KeyDown(vbKeySpace, 0): Exit Sub
        '81535:刘鹏飞,审签时允许修改日期时间列
        'If VsfData.COL = mlngDate Or VsfData.COL = mlngTime Or VsfData.COL = mlngYear Then Exit Sub
        If Val(VsfData.TextMatrix(lngStart, mlngRecord)) = 0 Then Exit Sub
        If VsfData.TextMatrix(lngStart, mlngSigner) = "" Then Exit Sub
        If VsfData.Cell(flexcpChecked, lngStart, mlngChoose) = flexTSUnchecked Then Exit Sub '没有选中的记录不能编辑
    Else
        '审签过的数据只能在审签状态下修改
        If InStr(1, VsfData.TextMatrix(lngStart, mlngSigner), "/") <> 0 Then
            RaiseEvent AfterRowColChange("已审签的数据不允许编辑！", True, mblnSign, mblnArchive)
            Exit Sub
        End If
        
        '只要是签名数据就不允许修改
        '--------------------------
'        '如果当前操作员的级别比已签名操作员的级别低,不允许其编辑数据
'        If VsfData.TextMatrix(lngStart, mlngSigner) <> "" Then
'            If mintVerify > Val(VsfData.TextMatrix(lngStart, mlngSignLevel)) + 1 Then
'                RaiseEvent AfterRowColChange("当前操作员的级别比已签名操作员的级别低,不允许编辑数据！", True, mblnSign, mblnArchive)
'                Exit Sub
'            End If
'        End If
        If VsfData.TextMatrix(lngStart, mlngSigner) <> "" Then
            RaiseEvent AfterRowColChange("已签名的数据不允许再次编辑，请取消签名后再试！", True, mblnSign, mblnArchive)
            Exit Sub
        End If
        '--------------------------
        '63706:刘鹏飞,2013-11-20,最后一列始终绑定的是护士,不管是否在记录单绑定了护士列
        '默认签名人与保存人相同,不具有修改他人护理记录权限的操作员,不允许修改他人的数据
        strName = FormatValue(VsfData.TextMatrix(lngStart, VsfData.Cols - 1))
        If strName <> "" Then
            If strName <> gstrUserName And _
                InStr(1, mstrPrivs, "他人护理记录") = 0 Then
                RaiseEvent AfterRowColChange("您没有修改他人护理记录数据的权限！原操作员:" & strName, True, mblnSign, mblnArchive)
                Exit Sub
            End If
        End If
    End If
    If mblnArchive Then Exit Sub
    If Not mblnShow Or Not mblnEditable Then Exit Sub
    If VsfData.TextMatrix(lngStart, mlngDemo) <> "" Then
        '只有新增的未保存的数据，才允许修改日期与时间
        If (VsfData.COL = mlngDate Or VsfData.COL = mlngTime Or VsfData.COL = mlngYear) Then
            If Val(VsfData.TextMatrix(lngStart, mlngDemo)) > 1 Then
                Exit Sub
            Else
                'If Val(VsfData.TextMatrix(lngStart, mlngRecord)) > 0 Then Exit Sub
            End If
        End If
    End If
    '未绑定项目的空列不能编辑
    If (InStr(1, "," & mstrCOLNothing & ",", "," & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ",") <> 0) Then
        mrsSelItems.Filter = "列=" & VsfData.COL - (cHideCols + VsfData.FixedCols - 1)
        If mrsSelItems.RecordCount = 0 Then
            RaiseEvent AfterRowColChange("不允许编辑为绑定项目的列，请手工添加活动项目！", True, mblnSign, mblnArchive)
            Exit Sub
        End If
    End If
    '跨页数据行不允许对整行进行粘贴,删除,只能编辑除活动项目外的列
    If InStr(1, VsfData.TextMatrix(lngStart, mlngRowCount), "|") <> 0 And lngStart = 3 And Val(VsfData.TextMatrix(lngStart, mlngStartRowPage)) <> mint页码 Then
        If Val(Split(VsfData.TextMatrix(lngStart, mlngRowCount), "|")(0)) <> Val(VsfData.TextMatrix(lngStart, mlngRowCurrent)) Then
            If InStr(1, "," & mstrCOLNothing & ",", "," & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ",") <> 0 Then
                RaiseEvent AfterRowColChange("不允许修改跨页数据行的活动项目数据！", True, mblnSign, mblnArchive)
                Exit Sub
            End If
        End If
    End If
    '同步数据列不允许编辑
    strCols = GetSynItems(2, intMax)
    If strCols <> "" Then
        '存在同步数据的行,日期与时间是不允许修改的
        strCols = "," & strCols & ","
        If VsfData.COL = mlngDate Or VsfData.COL = mlngTime Or VsfData.COL = mlngYear Or _
            InStr(1, strCols, "," & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ",") <> 0 Then
            RaiseEvent AfterRowColChange("同步的数据行，不允许修改时间和同步的数据列！", True, mblnSign, mblnArchive)
            Exit Sub
        End If
    End If
    '汇总行涉及到的明细,如果汇总行已签名则其汇总列不允许修改
    If mstrColCollect <> "" Then
        If Val(VsfData.TextMatrix(lngStart, mlngRecord)) > 0 Then
            '不允许修改汇总列数据，也不允许修改日期与时间
            If InStr(1, "|" & mstrColCollect, "|" & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ";") <> 0 Then
                blnCheck = True
            ElseIf InStr(1, "|" & mstrColCorrelative, "|" & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ",") <> 0 And mstrColCorrelative <> "" Then
                blnCheck = True
            ElseIf VsfData.COL <= mlngTime Then
                blnCheck = True
            End If
            If blnCheck = True Then
                If ISCollectSigned(mlng文件ID, Mid(VsfData.TextMatrix(lngStart, mlngActTime), 1, 10), Format(VsfData.TextMatrix(lngStart, mlngActTime), "HH:MM")) Then
                    RaiseEvent AfterRowColChange("该条数据所对应的汇总行数据已签名，不允许修改当前汇总列或日期时间列！", True, mblnSign, mblnArchive)
                    Exit Sub
                End If
            End If
        End If
    End If
    
    If VsfData.COL <= mlngNoEditor - 1 Then Call ShowInput
    
    On Error Resume Next
    '让控件获得焦点
    Select Case mintType
    Case 0, 3
        picInput.SetFocus
    Case 1, 2
        If mintType = 2 Then
            lstSelect(mintType - 1).SetFocus
        Else
            PicLst.SetFocus
        End If
    Case 4, 5
        picDouble.SetFocus
    Case 6
        picMutilInput.SetFocus
    Case 7
        cboChoose(0).SetFocus
    Case 8
        cboYear.SetFocus
    End Select
    If Err <> 0 Then Err.Clear
End Sub

Private Sub vsfData_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strInfo As String
    Dim strCols As String
    Dim intMax As Integer
    Dim strItemInfo As String
    Dim blnExit As Boolean
    
    On Error GoTo ErrHand
    
    If mblnInit = False Then Exit Sub
    If mblnEditable = False Then Exit Sub
    If OldRow = NewRow And OldCol = NewCol Then Exit Sub
    
    '63401:刘鹏飞,2013-07-16,如果选择的列不是活动列将隐藏活动项目设置
    '                       如果选择的列是其他活动项目列将重新加载活动项目设置界面
    If OldCol <> NewCol And picCloumn.Visible = True Then
        If InStr(1, "," & mstrCOLNothing & ",", "," & NewCol - (cHideCols + VsfData.FixedCols - 1) & ",") <> 0 Then
            Call BoundItems(NewCol - (cHideCols + VsfData.FixedCols - 1))
        Else
            picCloumn.Visible = False
        End If
    End If
    
    '选择列,同步数据列直接退出,避免此处清除提示信息
    If NewCol = mlngChoose Then Exit Sub
    strCols = GetSynItems(2, intMax)
    If strCols <> "" Then
        strCols = "," & strCols & ","
        If InStr(1, strCols, "," & NewCol - (cHideCols + VsfData.FixedCols - 1) & ",") <> 0 Then blnExit = True
    End If
    
    '显示当前项目的相关信息
    mrsSelItems.Filter = "列=" & NewCol - (cHideCols + VsfData.FixedCols - 1)
    If mrsSelItems.RecordCount <> 0 Then
        mrsItems.Filter = "项目序号=" & mrsSelItems!项目序号
        If mrsItems.RecordCount <> 0 Then
            If NVL(mrsItems!项目值域) <> "" Then
                If mrsItems!项目类型 = 0 Then
                    strInfo = "有效范围:" & Split(mrsItems!项目值域, ";")(0) & "～" & Split(mrsItems!项目值域, ";")(1)
                Else
                    strInfo = "有效范围:" & mrsItems!项目值域
                End If
            Else
                strInfo = ""
            End If
            strItemInfo = Trim(NVL(mrsItems!说明, ""))
        End If
    End If
    mrsSelItems.Filter = 0
    mrsItems.Filter = 0
    
    '--48659:刘鹏飞,2012-09-14,添加字段'说明'
    RaiseEvent ShowTipInfo(VsfData, strItemInfo, True)
    
    If blnExit = True Then Exit Sub
    
    '检查是否已签名
    If VsfData.TextMatrix(VsfData.ROW, mlngRowCount) = "" Then
        intMax = VsfData.ROW
    Else
        intMax = GetStartRow(VsfData.ROW)
    End If
    mblnSign = (VsfData.TextMatrix(intMax, mlngSigner) <> "")
    
    RaiseEvent AfterRowColChange(strInfo, False, mblnSign, mblnArchive)
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub vsfData_DrawCell(ByVal hDC As Long, ByVal ROW As Long, ByVal COL As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    Call DrawCell(hDC, ROW, COL, Left, Top, Right, Bottom, Done)
End Sub

Private Sub vsfdata_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngStart As Long
    Dim intLevel As Integer
    Dim strField As String, strKey As String, strValue As String
    Dim lngDemo As Long, lngRow As Long, lngRowCount As Long, lngNextGroupRow As Long
    Dim arrRow
    On Error GoTo ErrHand
    
    If mblnInit = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Then
        If Not mblnShow And VsfData.COL = mlngDate Then
            mblnShow = True
            Call VsfData_EnterCell
        Else
            Call MoveNextCell
        End If
    ElseIf KeyCode = vbKeySpace And mblnVerify Then
        '只勾选起始行
        lngStart = GetStartRow(VsfData.ROW)
        If VsfData.TextMatrix(lngStart, mlngTime) = "" And Val(VsfData.TextMatrix(lngStart, mlngDemo)) <= 1 Then Exit Sub
        
        If mintVerify = 未定义 Then
            RaiseEvent AfterRowColChange("您当前还未设置聘任技术职务，请在人员管理中设置！", True, mblnSign, mblnArchive)
            Exit Sub
        End If
        '审签时,当前记录已签名,且操作员的签名级别比上次签名级别高才允许
        If VsfData.TextMatrix(lngStart, mlngSignLevel) = "" Then
            RaiseEvent AfterRowColChange("该数据还未签名，不能进行审签！", True, mblnSign, mblnArchive)
            Exit Sub
        Else
            intLevel = Val(VsfData.TextMatrix(lngStart, mlngSignLevel)) + 1
        End If
         '43588:刘鹏飞,2012-09-13,添加记录单审签模式
        If InStr(1, VsfData.TextMatrix(lngStart, mlngSigner), "/") <> 0 And mintSignMode = 1 Then
            RaiseEvent AfterRowColChange("当前审签模式为【1-审签权限】，您只能对未审签的数据进行操作！" & vbCrLf & "详细信息：该行的数据已经审签", True, mblnSign, mblnArchive)
            Exit Sub
        End If
        
        If mintVerify >= intLevel And mintSignMode = 0 Then
            RaiseEvent AfterRowColChange("您的级别[" & GetVerify(mintVerify) & "]要比上次审签人的级别[" & GetVerify(intLevel) & "]高才能勾选该记录！", True, mblnSign, mblnArchive)
            Exit Sub
        End If
        
        '对于分组数据审签时，需要选择本分组所有行
        lngDemo = Val(VsfData.TextMatrix(lngStart, mlngDemo))
        If lngDemo > 1 Then '寻找起始行
            lngRow = lngStart
            lngStart = lngRow - lngDemo + 1
            If Val(VsfData.TextMatrix(lngStart, mlngDemo)) <> 1 Then
                For lngStart = lngRow To VsfData.FixedRows Step -1
                    If Val(VsfData.TextMatrix(lngStart, mlngDemo)) = 1 Then
                        lngRow = lngStart
                        Exit For
                    End If
                Next lngStart
                If lngStart < VsfData.FixedRows Then Exit Sub
                lngStart = lngRow
            End If
            
            '选择的是非分组起始行，要先检查起始行
            If VsfData.TextMatrix(lngStart, mlngSignLevel) = "" Then
                RaiseEvent AfterRowColChange("该分组起始行中的数据还未签名，请先签名后在进行审签！", True, mblnSign, mblnArchive)
                Exit Sub
            Else
                intLevel = Val(VsfData.TextMatrix(lngStart, mlngSignLevel)) + 1
            End If
            '43588:刘鹏飞,2012-09-13,添加记录单审签模式
            If InStr(1, VsfData.TextMatrix(lngStart, mlngSigner), "/") <> 0 And mintSignMode = 1 Then
                RaiseEvent AfterRowColChange("当前审签模式为【1-审签权限】，您只能对未审签的数据进行操作！" & vbCrLf & "详细信息：该分组起始行的数据已经审签", True, mblnSign, mblnArchive)
                Exit Sub
            End If
            If mintVerify >= intLevel And mintSignMode = 0 Then
                RaiseEvent AfterRowColChange("您的级别[" & GetVerify(mintVerify) & "]要比该分组起始行的审签或签名人的级别[" & GetVerify(intLevel) & "]高才能勾选该记录！", True, mblnSign, mblnArchive)
                Exit Sub
            End If
        End If
        arrRow = Array()
        ReDim Preserve arrRow(UBound(arrRow) + 1)
        arrRow(UBound(arrRow)) = lngStart
        lngRowCount = Val(Split(VsfData.TextMatrix(lngStart, mlngRowCount), "|")(0))
        lngNextGroupRow = lngStart + lngRowCount - 1
        '1.检查分组数据是否已经签名，2.记录分组数据开始行
        For lngRow = lngStart + lngRowCount To VsfData.Rows - 1
            '分组中每一个分组子行数据可能占用多行，只有在第一行保存了分组索引。分组总数据行数=每一个分组子行数+每一个分组子行的行数
            If lngRow > lngNextGroupRow Then
                If Val(VsfData.TextMatrix(lngRow, mlngDemo)) <= 1 Then Exit For      '不分组或遇新分组就退出
                If VsfData.TextMatrix(lngRow, mlngSignLevel) = "" Then
                    RaiseEvent AfterRowColChange("该分组中存在未签名的数据，请先签名后在进行审签！", True, mblnSign, mblnArchive)
                    Exit Sub
                End If
                intLevel = Val(VsfData.TextMatrix(lngRow, mlngSignLevel)) + 1
                '43588:刘鹏飞,2012-09-13,添加记录单审签模式
                If InStr(1, VsfData.TextMatrix(lngRow, mlngSigner), "/") <> 0 And mintSignMode = 1 Then
                    RaiseEvent AfterRowColChange("当前审签模式为【1-审签权限】，您只能对未审签的数据进行操作！" & vbCrLf & "详细信息：该分组中第【" & UBound(arrRow) + 2 & "】组的数据已经审签", True, mblnSign, mblnArchive)
                    Exit Sub
                End If
                If mintVerify >= intLevel And mintSignMode = 0 Then
                    RaiseEvent AfterRowColChange("您的级别[" & GetVerify(mintVerify) & "]要比该分组中第【" & UBound(arrRow) + 2 & "】组的审签或签名人的级别[" & GetVerify(intLevel) & "]高才能勾选该记录！", True, mblnSign, mblnArchive)
                    Exit Sub
                End If
                lngNextGroupRow = Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(0)) + lngRow - 1
                ReDim Preserve arrRow(UBound(arrRow) + 1)
                arrRow(UBound(arrRow)) = lngRow
            End If
        Next lngRow
        '选中所有分组数据
        For lngRow = 0 To UBound(arrRow)
            lngStart = Val(arrRow(lngRow))
            VsfData.Cell(flexcpChecked, lngStart, mlngChoose) = IIf(VsfData.Cell(flexcpChecked, lngStart, mlngChoose) = flexTSChecked, flexTSUnchecked, flexTSChecked)
            '保存修改记录以便同步
            strField = "ID|页号|行号|列号|记录ID|数据|删除"
            strKey = mint页码 & "," & lngStart & "," & mlngChoose
            strValue = strKey & "|" & mint页码 & "|" & lngStart & "|" & mlngChoose & "|" & Val(VsfData.TextMatrix(lngStart, mlngRecord)) & "|" & VsfData.Cell(flexcpChecked, lngStart, mlngChoose) & "|1"
            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
        Next lngRow
    Else
        If Not (KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or Shift <> 0) Then
            mblnShow = True
            Call VsfData_EnterCell
        End If
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function GetVerify(ByVal intVerify As Integer) As String
'获取当前级别对应的聘任技术职务
    Dim strVerify As String
    Select Case intVerify
        Case 正高
            strVerify = "主任护师"
        Case 副高
            strVerify = "副主任护师"
        Case 中级
            strVerify = "主管护师"
        Case 师级
            strVerify = "护师"
        Case 员士
            strVerify = "护士"
        Case Else
            strVerify = "未定义"
    End Select
    GetVerify = strVerify
End Function

Private Sub InitVariable()
    '清除常量
    mlngDate = -1
    mlngTime = -1
    mlngOperator = -1
    mlngSigner = -1
    mlngSignTime = -1
    mlngSignName = -1
    mlngRecord = -1
    mlngNoEditor = -1
    
    mblnChange = False
    mblnShow = False
    mblnSign = False
    mblnArchive = False
    mblnEditAssistant = False
    mblnEditText = False
    mblnElement = False
End Sub

Private Sub InitCons()
    '隐藏输入控件
    picInput.Visible = False
    lstSelect(0).Visible = False
    lstSelect(1).Visible = False
    picDouble.Visible = False
    picDoubleChoose.Visible = False
    picMutilInput.Visible = False
    cmdWord.Visible = False
    txtLst.Visible = False
    PicLst.Visible = False
    picBiref.Visible = False
    picCloumn.Visible = False
End Sub

Private Function InitMenuBar() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom
    Dim cbrPop As CommandBarControl
    Dim cbrToolBar As CommandBar
    Dim rs As ADODB.Recordset
    Dim objExtendedBar As CommandBar
    
    On Error GoTo ErrHand
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    cbsThis.ActiveMenuBar.Title = "菜单栏"
    cbsThis.ActiveMenuBar.Visible = False
    
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    With cbsThis.Options
        .AlwaysShowFullMenus = False
        .ShowExpandButtonAlways = False
        .UseDisabledIcons = True
        .SetIconSize True, 16, 16
        .LargeIcons = True
    End With

    '------------------------------------------------------------------------------------------------------------------
    '菜单定义
    cbsThis.ActiveMenuBar.Title = "菜单"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    
        '------------------------------------------------------------------------------------------------------------------
        '工具栏定义
        Set cbrToolBar = cbsThis.Add("标准", xtpBarTop)
        cbrToolBar.ShowTextBelowIcons = False
        cbrToolBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
        cbrToolBar.ModifyStyle XTP_CBRS_GRIPPER, 0
        With cbrToolBar.Controls
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Group_New, "分组"): cbrControl.IconId = 3096: cbrControl.ToolTipText = "开始分组"
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Group_Append, "追加"): cbrControl.IconId = 3045: cbrControl.ToolTipText = "追加分组(Ctrl+A)"
            
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Copy, "复制"): cbrControl.ToolTipText = "复制(Ctrl+C)": cbrControl.BeginGroup = True
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_PASTE, "粘贴"):  cbrControl.ToolTipText = "粘贴(Ctrl+V)"
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Clear, "清除"):   cbrControl.ToolTipText = "清除"
            
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_SPECIALCHAR, "特殊符号"):  cbrControl.ToolTipText = "插入特殊符号(Ctrl+D)": cbrControl.BeginGroup = True
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Word, "词句选择"):  cbrControl.ToolTipText = "词句选择(Ctrl+W)"
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Import, "入量导入"):  cbrControl.ToolTipText = "入量导入(Ctrl+I)"
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Brief, "小结"): cbrControl.ToolTipText = "小结"
            
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Element, "标签要素"): cbrControl.IconId = conMenu_Edit_Append: cbrControl.BeginGroup = True: cbrControl.ToolTipText = "自定义标签要素内容录入"
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Append, "列绑定"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "列绑定"
        
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_PrevPage, "上页"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "上页"
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NextPage, "下页"):   cbrControl.ToolTipText = "下页"
        End With
    
        For Each cbrControl In cbrToolBar.Controls
            If cbrControl.Type <> xtpControlCustom And cbrControl.Type <> xtpControlLabel Then
                cbrControl.Style = xtpButtonIconAndCaption
            End If
        Next
        Set mcbrToolBar = cbrToolBar
    
         '快键绑定
        With cbsThis.KeyBindings
            .Add FCONTROL, Asc("C"), conMenu_Edit_Copy
            .Add FCONTROL, Asc("V"), conMenu_Edit_PASTE
            .Add FCONTROL, Asc("D"), conMenu_Edit_SPECIALCHAR
            .Add FCONTROL, Asc("A"), conMenu_Edit_Group_Append
            .Add FCONTROL, Asc("W"), conMenu_Edit_Word
            .Add FCONTROL, Asc("S"), conMenu_Save
            .Add FCONTROL, Asc("I"), conMenu_Edit_Import
        End With
    
    InitMenuBar = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function CheckTxtTime(objText As TextBox) As String
    Dim strInput As String
    Dim strHour As String, strMin As String
    '检查时点录入是否合法，并返回数据
    
    If Trim(objText.Text) <> "" Then
        strInput = Trim(objText.Text)
        If InStr(1, strInput, ":") > 0 Then
            strHour = Split(strInput, ":")(0)
            strMin = Split(strInput, ":")(1)
        ElseIf InStr(1, strInput, "：") > 0 Then
            strHour = Split(strInput, "：")(0)
            strMin = Split(strInput, "：")(1)
        Else
            strHour = strInput
            strMin = "00"
        End If
        strHour = Format(strHour, "00")
        strMin = Format(strMin, "00")
        If Not IsNumeric(strHour) Then
            RaiseEvent AfterRowColChange("开始时点中含有非法字符！", True, mblnSign, mblnArchive)
            Exit Function
        End If
        If Val(strHour) < 0 Or Val(strHour) > 23 Then
            RaiseEvent AfterRowColChange("开始时点不正确，小时值应该>0且小于24！", True, mblnSign, mblnArchive)
            Exit Function
        End If
        If Not IsNumeric(strMin) Then
            RaiseEvent AfterRowColChange("开始时点中含有非法字符！", True, mblnSign, mblnArchive)
            Exit Function
        End If
        If Val(strMin) < 0 Or Val(strMin) > 59 Then
            RaiseEvent AfterRowColChange("开始时点不正确，分钟值应该>0且小于60！", True, mblnSign, mblnArchive)
            Exit Function
        End If
        strInput = strHour & ":" & strMin
    End If
    CheckTxtTime = strInput
End Function

Private Function CheckTime(ByVal lngRow As Long, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal int婴儿 As Integer, _
    ByVal strTime As String, ByVal strCurTime As String, ByRef strMsg As String) As Boolean
    Dim blnMsg As Boolean
    Dim blnExist As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strBabyOutTime As String
    On Error GoTo ErrHand
    '数据发生时间必须在当前科室的有效时间范围内
    
    blnMsg = (strMsg <> "")
    
    '检查文件开始,结束时间
    If Format(strTime, "YYYY-MM-DD HH:mm") < Format(mstr开始时间, "yyyy-MM-dd HH:mm") Then
        strMsg = "发生时间不能小于文件开始时间[" & mstr开始时间 & "]"
        GoTo exitHand
    End If
    If mstr结束时间 <> "" Then
        If Format(strTime, "YYYY-MM-DD HH:mm") > Format(mstr结束时间, "yyyy-MM-dd HH:mm") Then
            strMsg = "发生时间不能大于文件结束时间[" & mstr结束时间 & "]"
            GoTo exitHand
        End If
    End If
    
    '75760:刘鹏飞,处理婴儿存在出院医嘱的情况
    If int婴儿 <> 0 Then
        strBabyOutTime = GetAdviceOutTime(lng病人ID, lng主页ID, int婴儿)
        If strBabyOutTime <> "" Then
            If Format(strTime, "YYYY-MM-DD HH:mm") > Format(strBabyOutTime, "YYYY-MM-DD HH:mm") Then
                strMsg = "第" & lngRow & "行的发生时间" & strTime & "有误！[发生时间不能大于出院时间:" & Format(strBabyOutTime, "YYYY-MM-DD HH:mm") & "]"
                GoTo exitHand
            End If
            '补录小时检查
            If Format(DateAdd("H", glngHours, strBabyOutTime), "yyyy-MM-dd HH:mm") < Format(strCurTime, "yyyy-MM-dd HH:mm") Then
                strMsg = "第" & lngRow & "行的发生时间" & strTime & "有误！[超过数据补录的有效时限:" & glngHours & "小时]"
                GoTo exitHand
            End If
            CheckTime = True
            Exit Function
        End If
    End If
    
    '根据病人变动记录进行检查
    gstrSQL = " Select   开始原因,病区ID,to_char(开始时间,'yyyy-MM-dd hh24:mi') AS 开始时间,to_char(NVL(终止时间,sysDate+" & mintPreDays & "),'yyyy-MM-dd hh24:mi') AS 终止时间 " & _
              " From 病人变动记录 " & _
              " Where 病人ID=[1] And 主页ID=[2]" & _
              " Order by 开始时间,开始原因"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取当前科室有效时间范围", lng病人ID, lng主页ID)
    With rsTemp
        .Filter = "病区ID=" & mlng病区ID
        Do While Not .EOF
            If Format(strTime, "YYYY-MM-DD HH:mm") >= Format(!开始时间, "YYYY-MM-DD HH:mm") And Format(strTime, "YYYY-MM-DD HH:mm") <= Format(!终止时间, "YYYY-MM-DD HH:mm") Then
                blnExist = True
                Exit Do
            End If
            .MoveNext
        Loop
        .Filter = 0
        '找到了就退出
        If blnExist Then
            If Not IsAllowInput(lng病人ID, lng主页ID, strTime, strCurTime) Then
                strMsg = "第" & lngRow & "行的发生时间" & strTime & "有误！[超过数据补录的有效时限:" & glngHours & "小时]"
                GoTo exitHand
            End If
            
            CheckTime = True
            Exit Function
        End If
        
        '没找到,就整理原因进行准确性提示
        .Filter = "开始原因=1"
        If .RecordCount <> 0 Then
            If !开始原因 = 1 And Format(strTime, "YYYY-MM-DD HH:mm") < Format(!开始时间, "YYYY-MM-DD HH:mm") Then
                strMsg = "第" & lngRow & "行的发生时间" & strTime & "有误！[发生时间不能小于病人入院时间:" & !开始时间 & "]"
                GoTo exitHand
            End If
        End If
        .Filter = "开始原因=2"
        If .RecordCount <> 0 Then
            If !开始原因 = 2 And Format(strTime, "YYYY-MM-DD HH:mm") < Format(!开始时间, "YYYY-MM-DD HH:mm") Then
                strMsg = "第" & lngRow & "行的发生时间" & strTime & "有误！[发生时间不能小于病人入科时间:" & !开始时间 & "]"
                GoTo exitHand
            End If
        End If
        .Filter = "开始原因=10"
        If .RecordCount <> 0 Then
            If !开始原因 = 10 And Format(strTime, "YYYY-MM-DD HH:mm") > Format(!终止时间, "YYYY-MM-DD HH:mm") Then
                strMsg = "第" & lngRow & "行的发生时间" & strTime & "有误！[发生时间不能大于出院时间:" & !终止时间 & "]"
                GoTo exitHand
            End If
        End If
        .Filter = 0
        '其他情况说明
        strMsg = "第" & lngRow & "行的发生时间" & strTime & "有误！[不在当前病区的有效时间范围内]"
        GoTo exitHand
    End With
    
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
exitHand:
    rsTemp.Filter = 0
    If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
End Function

Private Function CheckInput(strReturn As String, strInfo As String) As Boolean
    Dim i As Integer, j As Integer
    Dim strOrders As String, strText As String
    '检查录入数据的合法性(中文也认为是一个字符,考虑到体温项目等存在不升\外出等信息)
    '返回的数据,如果一列绑定多个项目,以单引号做为分隔符
    
    'mintType:0=文本框录入;1=单选;2=多选;3=选择;4-血压或一列绑定了两个项目,其格式类似血压的输入项目;5=一列绑定了两个项目且均是选择项目;
    '6=一列绑定N个项目,手工录入
    Select Case mintType
    Case 0
        strText = txtInput.Text
        strOrders = txtInput.Tag
    Case 1, 2   '免检
        If mintType = 1 Then
            If Val(PicLst.Tag) = 0 Then
                txtLst.Text = ""
                If InStr(1, lstSelect(mintType - 1).Text, "-") <> 0 Then
                    strText = Mid(lstSelect(mintType - 1).Text, InStr(1, lstSelect(mintType - 1).Text, "-") + 1)
                Else
                    strText = ""
                End If
            Else
                strText = Trim(txtLst.Text)
            End If
        Else
            j = lstSelect(mintType - 1).ListCount
            For i = 1 To j
                If lstSelect(mintType - 1).Selected(i - 1) Then
                    strText = strText & "," & Mid(lstSelect(mintType - 1).List(i - 1), InStr(1, lstSelect(mintType - 1).List(i - 1), "-") + 1)
                End If
            Next
            If strText <> "" Then strText = Mid(strText, 2)
        End If
        strOrders = lstSelect(mintType - 1).Tag
    Case 4
        strText = txtUpInput.Text & "'" & txtDnInput.Text
        strOrders = txtUpInput.Tag & "'" & txtDnInput.Tag
    Case 6
        j = txt.Count
        For i = 1 To j
            strText = strText & "'" & txt(i - 1).Text
            strOrders = strOrders & "'" & txt(i - 1).Tag
        Next
        If strText <> "" Then
            strText = Mid(strText, 2)
            strOrders = Mid(strOrders, 2)
        End If
    Case 3      '免检
        strText = lblInput.Caption
    Case 5      '免检
        strText = lblUpInput.Caption & "/" & lblDnInput.Caption
    Case 7
        strText = cboChoose(0).Text & "/" & cboChoose(1).Text
    Case 8
        strText = cboYear.Text
    End Select
    strText = Replace(Replace(Replace(strText, Chr(10), ""), Chr(13), ""), Chr(1), "")
    If Val(strOrders) <> 0 Then
        If Not CheckValid(strText, strOrders, strInfo) Then Exit Function
    ElseIf VsfData.COL = mlngDate Or VsfData.COL = mlngTime Then
        If Not CheckDateTime(strText, strInfo) Then Exit Function
    End If
    
    strReturn = strText
    CheckInput = True
End Function

Private Function CheckDateTime(strText As String, strInfo As String) As Boolean
    Dim blnCheck As Boolean, blnExist As Boolean
    Dim strCurrDate As String
    Dim strDate As String
    Dim rsCheck As New ADODB.Recordset
    Dim arrTime As Variant
    On Error GoTo ErrHand
    
    '69355:刘鹏飞,2013-01-07,短日期格式的处理
    strCurrDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
    If VsfData.COL = mlngDate Then
        If mblnDateAd Then
            If Trim(strText) = "" Then
                strInfo = "日期不能为空！"
                Exit Function
            End If
            If InStr(1, strText, "/") = 0 Then
                strInfo = "日期格式错误，如1月12日：12/01"
                Exit Function
            End If
            
            If VsfData.TextMatrix(VsfData.ROW, mlngYear) = "" Then
                strDate = Mid(strCurrDate, 1, 5) & ToStandDate(strText)
                '检查是否翻年后编辑之前的时间(一个月的限制)
                If CDate(strDate) > CDate(Mid(strCurrDate, 1, 10)) And Mid(strCurrDate, 6, 2) = "01" And Mid(strDate, 6, 2) = "12" Then
                    strDate = DateAdd("yyyy", -1, CDate(strDate))
                End If
            Else
                strDate = VsfData.TextMatrix(VsfData.ROW, mlngYear) & "-" & ToStandDate(strText)
            End If
            If Not IsDate(strDate) Then
                strInfo = "录入的数据不是合法的日期，如1月12日：12/01"
                Exit Function
            Else
                VsfData.TextMatrix(VsfData.ROW, mlngYear) = Format(strDate, "YYYY")
            End If
        Else
            If Trim(strText) = "" Then
                strInfo = "日期不能为空！"
                Exit Function
            End If
            If Not IsDate(strText) Then
                strInfo = "录入的数据不是合法的日期，如1月12日：2011-01-12"
                Exit Function
            End If
            strDate = Format(strText, "yyyy-MM-dd")
            VsfData.TextMatrix(VsfData.ROW, mlngYear) = Format(strDate, "YYYY")
        End If
        If Format(strDate, "YYYY-MM-DD") > Format(DateAdd("d", mintPreDays, CDate(strCurrDate)), "YYYY-MM-DD") Then
            strInfo = "录入的日期已超出参数[超期录入天数：" & mintPreDays & "天]所指定的范围！"
            Exit Function
        End If
        
'        If VsfData.TextMatrix(VsfData.ROW, mlngTime) <> "" Then
'            blnCheck = True
'            strDate = strDate & " " & VsfData.TextMatrix(VsfData.ROW, mlngTime)
'        End If
    Else
        If Trim(strText) = "" Then
            strInfo = "时间不能为空！"
            Exit Function
        End If
        If InStr(1, Trim(strText), ":") = 0 Then
            Select Case Len(strText)
            Case 3, 4
                strText = String(4 - Len(strText), "0") & strText
                strText = Mid(strText, 1, 2) & ":" & Mid(strText, 3)
            Case Is < 3
                strText = String(2 - Len(strText), "0") & strText
                strText = Format(Now, "HH") & ":" & strText
            End Select
        End If
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
            strInfo = "录入的时点格式非法！[小时:分钟]"
            Exit Function
        End If
        If Mid(strText, 3, 1) <> ":" Then
            strInfo = "录入的时点格式非法！[小时:分钟]"
            Exit Function
        End If
        If Val(arrTime(0)) < 0 Or Val(arrTime(0)) > 23 Then
            strInfo = "录入的时点格式非法！[小时应在0至23之间]"
            Exit Function
        End If
        If Val(arrTime(1)) < 0 Or Val(arrTime(1)) > 59 Then
            strInfo = "录入的时点格式非法！[分钟应在0至59之间]"
            Exit Function
        End If
        
        '进行合法性检查
        If VsfData.TextMatrix(VsfData.ROW, mlngDate) <> "" Then
            strDate = VsfData.TextMatrix(VsfData.ROW, mlngDate)
            If mblnDateAd Then
                If VsfData.TextMatrix(VsfData.ROW, mlngYear) = "" Then
                    strDate = Mid(strCurrDate, 1, 5) & ToStandDate(strDate)
                    '检查是否翻年后编辑之前的时间(一个月的限制)
                    If CDate(strDate) > CDate(Mid(strCurrDate, 1, 10)) And Mid(strCurrDate, 6, 2) = "01" And Mid(strDate, 6, 2) = "12" Then
                        strDate = DateAdd("yyyy", -1, CDate(strDate))
                    End If
                Else
                    strDate = VsfData.TextMatrix(VsfData.ROW, mlngYear) & "-" & ToStandDate(strDate)
                End If
                If IsDate(strDate) Then
                    VsfData.TextMatrix(VsfData.ROW, mlngYear) = Format(strDate, "YYYY")
                Else
                    strInfo = "录入的数据不是合法的日期，如1月12日：12/01"
                    Exit Function
                End If
            Else
                strDate = Format(VsfData.TextMatrix(VsfData.ROW, mlngDate), "yyyy-MM-dd")
            End If
            
            strDate = Format(strDate & " " & strText, "YYYY-MM-DD HH:mm:ss")
            '70990:刘鹏飞,2014-03-13,超期补录天数控制修改
            If Format(strDate, "YYYY-MM-DD HH:mm") > Format(DateAdd("d", mintPreDays, CDate(strCurrDate)), "YYYY-MM-DD HH:mm") Then
                strInfo = "录入的日期已超出参数[超期录入天数：" & mintPreDays & "天]所指定的范围！"
                Exit Function
            End If
        
            blnCheck = True
        End If
    End If
    
    If blnCheck Then
        '不管是新录入还是修改的数据 如果存在历史数据都不允许修改
        gstrSQL = " Select 1 From 病人护理数据 Where 文件ID=[1] And 发生时间=[2] And ([3]=0 OR ID<>[3])"
        Set rsCheck = zlDatabase.OpenSQLRecord(gstrSQL, "检查发生时间", mlng文件ID, CDate(strDate), Val(VsfData.TextMatrix(VsfData.ROW, mlngRecord)))
        If rsCheck.RecordCount > 0 Then
            strInfo = "您录入的时点已经存在历史数据！"
            Exit Function
        End If
        
        '从数据库中没有找到，开始从用户录入的数据寻找
        If Not CheckChangeDataTime(VsfData.ROW, strDate, strInfo) Then Exit Function
        
        '81535:修改时间的对应的汇总列如果存在数据，则检查是否已经存在相应的小结并进行了签名
        '规则:新增的数据强制检查;已有的数据则只需要检查时间变化的数据(因可能存在A操作员在开始无汇总行或有但未签名进行了时间调整，B操作员签名了，A操作员在保存的情况)
        '说明：新增的数据是可以修改任何列；已有的数据如果汇总行已经签名是不允许修改日期时间列和汇总列的（未处理对修改的时间在同一个汇总范围的判断）
        blnCheck = True
        If Val(VsfData.TextMatrix(VsfData.ROW, mlngRecord)) > 0 Then
            If Format(VsfData.TextMatrix(VsfData.ROW, mlngActTime), "YYYY-MM-DD HH:mm") = Format(strDate, "YYYY-MM-DD HH:mm") Then
                blnCheck = False
            End If
        End If
        If blnCheck = True Then
            If CheckCollectIsData(VsfData.ROW) = True Then
                If ISCollectSigned(mlng文件ID, Format(strDate, "YYYY-MM-DD"), Format(strDate, "HH:MM")) Then
                    strInfo = "您录入的时点所对应的汇总行数据已签名，不允许再添加新的汇总列数据！"
                    Exit Function
                End If
            End If
        End If
        '数据发生时间不能在当前操作员所属科室的有效时间以前
        If Not CheckTime(VsfData.ROW, mlng病人ID, mlng主页ID, mint婴儿, strDate, strCurrDate, strInfo) Then
            Exit Function
        End If
    End If
    
    CheckDateTime = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function CheckChangeDataTime(ByVal lngRow As Long, ByVal strCurDate As String, ByRef strMsg As String) As Boolean
'检查新录入的时间，是否与现有的时间相同，如果相同则提示不能录入
    Dim strDateHistory As String, strTimeHistory As String, strDatetime As String '用户已经录入的日期和时间
    Dim lngCurRow As Long, intPage As Integer, blnDel As Boolean, blnTrue As Boolean
    Dim strCurrDate As String, lngRecord As Long, strActiveTime As String
    Dim strRows As String, strPages As String, strTimes As String, lngCol As Long
    Dim arrRows
    On Error GoTo ErrHand

    strCurrDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    With mrsCellMap
        .Filter = "列号=" & mlngDate & " OR 列号=" & mlngTime
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            If Not (lngCurRow = !行号 And intPage = !页号) Then
                blnDel = False
endWork:
                If lngCurRow = lngRow And intPage = mint页码 Then GoTo ErrNext
                If lngCurRow > 0 Then
                    mrsDataMap.Filter = "页号=" & intPage & " And 行号=" & lngCurRow
                    If mrsDataMap.RecordCount <> 0 Then
                        blnDel = (mrsDataMap!删除 = 1)
                        If mint页码 = intPage Then
                            If blnDel = False Then
                                blnDel = VsfData.RowHidden(lngCurRow)
                            Else
                                mrsDataMap!删除 = IIf(VsfData.RowHidden(lngCurRow) = True, 1, 0)
                                mrsDataMap.Update
                            End If
                        End If
                        
                        strActiveTime = Format(mrsDataMap.Fields(cControlFields + mlngActTime - VsfData.FixedCols).Value, "YYYY-MM-DD HH:mm:ss")
                    Else
                        '非本编辑页的数据,MrsCellMap中有记录，mrsDataMap中必然存在对应的数据
                        If mint页码 = intPage Then
                            blnDel = VsfData.RowHidden(lngCurRow)
                            strActiveTime = Format(VsfData.TextMatrix(lngCurRow, mlngActTime), "YYYY-MM-DD HH:mm:ss")
                        Else
                            strMsg = "第" & intPage & "页，第" & lngCurRow & "行的数据内部错误,请检查、记录本次操作并反馈，谢谢！"
                            Exit Function
                        End If
                    End If
                    mrsDataMap.Filter = 0
                End If
                
                If blnTrue = True And strDatetime <> "" Then
                    If Format(strDatetime, "YYYY-MM-DD HH:mm:ss") = Format(strCurDate, "YYYY-MM-DD HH:mm:ss") Then
                        '存在相同时间的数据没有删除，直接进行提示
                        If blnDel = False Then
                            strMsg = "第" & intPage & "页，第" & lngCurRow & "行已经存在相同时点的数据，请检查！"
                            Exit Function
                        Else
                            If lngRecord > 0 Then '保存的数据删除，如果时间和原有时间相同直接提示，不相同恢复时间为原有时间
                                If Format(strDatetime, "YYYY-MM-DD HH:mm:ss") = Format(strActiveTime, "YYYY-MM-DD HH:mm:ss") Then
                                    strMsg = "您录入的时点已经存在历史数据！"
                                    Exit Function
                                Else '恢复时间为原有时间
                                    mrsDataMap.Filter = "页号=" & intPage & " And 行号=" & lngCurRow
                                    If mrsDataMap.RecordCount <> 0 Then
                                        strActiveTime = Format(mrsDataMap.Fields(cControlFields + mlngActTime - VsfData.FixedCols).Value, "YYYY-MM-DD HH:mm:ss")
                                        mrsDataMap.Fields(cControlFields + mlngDate - VsfData.FixedCols).Value = Format(strActiveTime, "YYYY-MM-DD")
                                        mrsDataMap.Fields(cControlFields + mlngTime - VsfData.FixedCols).Value = Mid(strActiveTime, 12, 5)
                                        mrsDataMap.Update
                                    Else
                                        '非本编辑页的数据,MrsCellMap中有记录，mrsDataMap中必然存在对应的数据
                                        '上面已经检查,肯定是本页
                                        If mint页码 = intPage Then
                                            VsfData.TextMatrix(lngCurRow, mlngDate) = Format(strActiveTime, "YYYY-MM-DD")
                                            VsfData.TextMatrix(lngCurRow, mlngTime) = Mid(strActiveTime, 12, 5)
                                        End If
                                    End If
                                    mrsDataMap.Filter = 0
                                    '记录行号和页号
                                    strRows = strRows & "," & lngCurRow
                                    strPages = strPages & "," & intPage
                                    strTimes = strTimes & "," & strActiveTime
                                End If
                            Else '未保存的数据删除，直接清空记录集内容信息(不包含页号、行号、删除)
                                mrsDataMap.Filter = "页号=" & intPage & " And 行号=" & lngCurRow
                                If mrsDataMap.RecordCount <> 0 Then
                                    For lngCol = cControlFields To mrsDataMap.Fields.Count - 2
                                        If InStr(1, "," & mlngCollectType & "," & mlngRecord & ",", "," & (lngCol - cControlFields + VsfData.FixedCols) & ",") <> 0 Then
                                            mrsDataMap.Fields(lngCol).Value = 0
                                        Else
                                            mrsDataMap.Fields(lngCol).Value = Null
                                        End If
                                    Next lngCol
                                    mrsDataMap.Update
                                End If
                                If mint页码 = intPage Then
                                    For lngCol = VsfData.FixedCols To VsfData.Cols - 1
                                        VsfData.TextMatrix(lngCurRow, lngCol) = ""
                                    Next lngCol
                                End If
                                mrsDataMap.Filter = 0
                                '记录行号和页号
                                strRows = strRows & "," & lngCurRow
                                strPages = strPages & "," & intPage
                                strTimes = strTimes & "," & "[LPF]"
                            End If
                        End If
                    End If
ErrNext:
                    blnTrue = False
                    If .EOF Then Exit Do
                End If
                '赋初值
                intPage = !页号
                lngCurRow = !行号
                strDateHistory = ""
                strTimeHistory = ""
                strDatetime = ""
                lngRecord = NVL(!记录ID, 0)
                blnTrue = False
            End If
            
            If !列号 = mlngDate Then
                If NVL(!汇总, 0) <> 1 Then
                    strDateHistory = NVL(!数据)
                    If strDateHistory <> "" Then
                        If mblnDateAd Then
                            If InStr(1, "|" & mstrYears & "|", "|" & Val(NVL(!部位)) & "|") <> 0 Then
                                strDateHistory = NVL(!部位) & "-" & ToStandDate(strDateHistory)
                            Else
                                strMsg = "第" & intPage & "页，第" & lngCurRow & "行的[年份]数据错误，请记录本次操作步骤并反馈，然后重新录入数据，谢谢！"
                                Exit Function
'                                strDateHistory = Mid(strCurrDate, 1, 5) & ToStandDate(strDateHistory)
'                                '检查是否翻年后编辑之前的时间(一个月的限制)
'                                If CDate(strDateHistory) > CDate(Mid(strCurrDate, 1, 10)) And Mid(strCurrDate, 6, 2) = "01" And Mid(strDateHistory, 6, 2) = "12" Then
'                                    strDateHistory = DateAdd("yyyy", -1, CDate(strDateHistory))
'                                End If
                            End If
                        Else
                            strDateHistory = Format(strDateHistory, "yyyy-MM-dd")
                        End If
                    End If
                End If
            Else '时间列
                strTimeHistory = NVL(!数据, "00:00")
                If strDateHistory = "" Then strDateHistory = Mid(strCurrDate, 1, 10)
                strDatetime = strDateHistory & " " & strTimeHistory & ":00"

                '处理分组数据，保存时与普通数据无区别，只是秒数+
                If Val(NVL(!部位)) >= 1 Then
                    strDatetime = Mid(strDatetime, 1, 17) & String(2 - Len(!部位), "0") & Val(!部位) - 1
                End If
                strDatetime = Format(strDatetime, "YYYY-MM-DD HH:mm:ss")
                blnTrue = True
            End If
        .MoveNext
        Loop
        
        If blnTrue Then GoTo endWork
        mrsDataMap.Filter = 0
    End With
    
    '更新mrsCellMap记录集
    If Left(strRows, 1) = "," Then strRows = Mid(strRows, 2)
    If Left(strPages, 1) = "," Then strPages = Mid(strPages, 2)
    If Left(strTimes, 1) = "," Then strTimes = Mid(strTimes, 2)
    arrRows = Split(strRows, ",")
    For lngCurRow = 0 To UBound(arrRows)
        mrsCellMap.Filter = "页号=" & Val(Split(strPages, ",")(lngCurRow)) & " And 行号=" & Val(arrRows(lngCurRow))
        If CStr(Split(strTimes, ",")(lngCurRow)) = "[LPF]" Then
            Do While Not mrsCellMap.EOF
                mrsCellMap.Delete
                mrsCellMap.Update
                mrsCellMap.MoveNext
            Loop
        Else
            Do While Not mrsCellMap.EOF
                If mrsCellMap!列号 = mlngDate Then
                    mrsCellMap!数据 = Format(CStr(Split(strTimes, ",")(lngCurRow)), "YYYY-MM-DD")
                    mrsCellMap.Update
                ElseIf mrsCellMap!列号 = mlngTime Then
                    mrsCellMap!数据 = Mid(CStr(Split(strTimes, ",")(lngCurRow)), 12, 5)
                    mrsCellMap.Update
                End If
            mrsCellMap.MoveNext
            Loop
        End If
        
    Next lngCurRow
    
    mrsCellMap.Filter = 0
    strMsg = ""
    CheckChangeDataTime = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function CheckValid(strReturn As String, ByVal strOrders As String, strInfo As String) As Boolean
    Dim arrData, arrOrder
    Dim blnCheck As Boolean, blnNumber As Boolean
    Dim i As Integer, j As Integer
    Dim dblMin As Double, dblMax As Double
    Dim strText As String, strName As String, strFormat As String, strFormat1 As String
    
    '按列格式组装数据
    mrsSelItems.Filter = "列=" & VsfData.COL - (cHideCols + VsfData.FixedCols - 1)
    If mrsSelItems.RecordCount <> 0 Then
        '有此列但未进行定义
        strFormat = NVL(mrsSelItems!格式)   '{P[体温]C}{...}
        strFormat1 = strFormat
    End If
    mrsSelItems.Filter = 0
    
    '检查数据
    arrData = Split(strReturn, "'")
    arrOrder = Split(strOrders, "'")
    j = UBound(arrData)
    For i = 0 To j
        strText = arrData(i)
        strOrders = arrOrder(i)
        If Val(strOrders) <> 0 Then
            mrsItems.Filter = "项目序号=" & strOrders
            strName = GetActivePart(VsfData.COL, i) & UCase(mrsItems!项目名称)
            If strText <> "" Then
                blnCheck = True
                blnNumber = False
                If blnCheck Then
                    If mrsItems!项目类型 = 0 And InStr(1, "0,4", mrsItems!项目表示) <> 0 Then
                        blnNumber = True
                        strText = Val(strText)
                        If NVL(mrsItems!项目小数, 0) <> 0 Then   '等于零是通过控件的MaxLength来控制的
                            If InStr(1, strText, ".") <> 0 Then strText = Mid(strText, 1, InStr(1, strText, ".") - 1)
                            If Len(strText) > mrsItems!项目长度 Then
                                mrsItems.Filter = 0
                                strInfo = "[" & strName & "]录入的数据超过了合法精度！"
                                Exit Function
                            End If
                            
                            strText = Val(arrData(i))
                            If InStr(1, strText, ".") <> 0 Then
                                strText = Mid(strText, InStr(1, strText, ".") + 1)
                                If Len(strText) > mrsItems!项目小数 Then
                                    mrsItems.Filter = 0
                                    strInfo = "[" & strName & "]录入的小数部分超过了合法精度！"
                                    Exit Function
                                End If
                            End If
                            strText = Val(arrData(i))
                        End If
                        If mrsItems!项目表示 = 0 Then
                            If Not IsNull(mrsItems!项目值域) Then
                                dblMin = Val(Split(mrsItems!项目值域, ";")(0))
                                dblMax = Val(Split(mrsItems!项目值域, ";")(1))
                                If Not (Val(strText) >= dblMin And Val(strText) <= dblMax) Then
                                    mrsItems.Filter = 0
                                    strInfo = "[" & strName & "]录入的数据不在" & Format(dblMin, "#0.00") & "～" & Format(dblMax, "#0.00") & "的有效范围！"
                                    Exit Function
                                End If
                            End If
                        End If
                    Else
                        If LenB(StrConv(strText, vbFromUnicode)) > mrsItems!项目长度 Then
                            strInfo = "[" & strName & "]录入的数据超过了最大长度：" & mrsItems!项目长度 & "！"
                            mrsItems.Filter = 0
                            Exit Function
                        End If
                    End If
                End If
                If IsNumeric(strText) And blnNumber = True Then
                    If Val(strText) < 1 And Val(strText) > 0 Then strText = "0" & Val(strText)
                End If
                strFormat = Replace(strFormat, "[" & strName & "]", strText)
            Else
                '删除该项目
                If InStr(1, "," & mstrCOLNothing & ",", "," & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ",") = 0 Then
                    Call SubstrPro(strFormat, strName)
                Else
                    '当项目无数据时,如果当前列具有对角线属性,则不清除
                    strFormat = Replace(strFormat, "[" & strName & "]", strText)
                End If
            End If
        Else
            strFormat = strReturn
        End If
    Next
    If j = -1 Then
        strOrders = arrOrder(i)
        If Val(strOrders) <> 0 Then
            mrsItems.Filter = "项目序号=" & strOrders
            strName = GetActivePart(VsfData.COL, i) & mrsItems!项目名称
            strFormat = Replace(strFormat, "[" & strName & "]", strText)
        End If
    End If
    mrsItems.Filter = 0
    
    strFormat = Replace(strFormat, "{", "")
    strFormat = Replace(strFormat, "}", "")
    If InStr(1, "," & mstrCOLNothing & ",", "," & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ",") = 0 Then
        If strFormat = SubstrFormat(strFormat1, arrOrder) Then strFormat = ""
    End If

    strReturn = strFormat
    
    CheckValid = True
End Function

Public Function SubstrFormat(ByVal strData As String, ByVal arrOrder As Variant) As String
    '获取绑定项目的前后缀符号
    Dim i As Integer
    Dim strOrders As String, strName As String
    For i = 0 To UBound(arrOrder)
        strOrders = CStr(arrOrder(i))
        If Val(strOrders) <> 0 Then
            mrsItems.Filter = "项目序号=" & strOrders
            strName = GetActivePart(VsfData.COL, i) & mrsItems!项目名称
        End If
        strData = Replace(strData, "[" & strName & "]", "")
    Next i
    strData = Replace(strData, "{", "")
    strData = Replace(strData, "}", "")
    
    SubstrFormat = strData
End Function

Public Function SubstrVal(ByVal strData As String, ByVal strFormat As String, ByVal strName As String, intPos As Integer) As String
    Dim i As Integer, j As Integer, l As Integer, r As Integer
    Dim strQZ As String, strHZ As String
    '返回前一个项目的后缀符号+当前项目的前缀符号的位置
    
    If strData = "" Then Exit Function
    strData = strData
    j = Len(strFormat)
    l = InStr(1, strFormat, "[" & strName & "]")
    If l = 0 Then Exit Function
    '得到前缀
    For i = l To 1 Step -1
        If Mid(strFormat, i, 1) = "{" Then Exit For
    Next
    strQZ = Mid(strFormat, i + 1, l - i - 1)
    '找到该项目格式串中的结束符号
    i = l + Len(strName) + 2
    For r = i To j
        If Mid(strFormat, r, 1) = "}" Then Exit For
    Next
    '得到后缀
    strHZ = Mid(strFormat, i, r - i)
    '如果后缀为空,继续向后寻找下一个项目的前缀符号
    If strHZ = "" And r < j Then
        For r = r + 1 To j
            If Mid(strFormat, r, 1) = "[" Then Exit For
        Next
        strHZ = Mid(strFormat, InStr(i, strFormat, "{") + 1, r - InStr(i, strFormat, "{") - 1)
    End If
    '取出指定项目完整的数据串
    If strHZ <> "" Then
        j = InStr(intPos, strData, strHZ) '因为是连续取数,考虑到分隔符可能相同的情况,记录上一次的最后位置,下次从这个位置往后取数据
        If j = 0 Then
            '有可能中间存在回车换行符
            j = InStr(intPos, Replace(strData, vbCrLf, ""), strHZ)
            If j = 0 Then Exit Function
        End If
    End If
    strData = Mid(strData, intPos)
    '前缀为空,继续向前寻找上一个项目的后缀符号
'    If strQZ = "" And i > 1 And intPos > 1 Then
'        For i = i - 1 To 1 Step -1
'            If Mid(strFormat, i, 1) = "]" Then Exit For
'        Next
'        strQZ = Mid(strFormat, i + 1, InStr(i, strFormat, "}") - i - 1)
'    End If
    
    SubstrVal = SubstrAnaly(strData, strHZ, strQZ)
    intPos = intPos + Len(strQZ & SubstrVal & strHZ)
    '如果是数字型则去掉回车换行符返回,如果是字符型则原样返回
'    If strHZ <> "" Then
'
'        strData = Mid(strData, 1, InStr(1, Replace(strData, vbCrLf, ""), strHZ) - 1) '丢弃该项目后的数据
'        intPOS = i + Len(strHZ)
'    End If
'    If strQZ <> "" Then strData = Mid(strData, InStr(1, strData, strQZ) + Len(strQZ)) '丢弃该项目后的数据
'    SubstrVal = strData ' Replace(strData, vbCrLf, "")
End Function

Private Function SubstrAnaly(ByVal strData As String, ByVal strHZ As String, ByVal strQZ As String) As String
    Dim strText As String
    Dim strCompare As String           '对比串
    Dim intLen As Integer, intActLen As Integer           '前缀/后缀的长度
    Dim intPos As Integer, intEnd As Integer
    Dim lngASC As Long
    Dim blnFind As Boolean
    '遇到回车换行符忽略,空格重新比对
    
    strText = strData
    If strHZ <> "" Then
        '把后缀去掉
        strHZ = Replace(strHZ, vbCrLf, "")
        intEnd = Len(strText)
        intLen = Len(strHZ)
        For intPos = 1 To intEnd
            lngASC = Asc(Mid(strText, intPos, 1))
            intActLen = intActLen + 1
            If Not (lngASC = 13 Or lngASC = 10) Then
                If lngASC = 32 Then
                    strCompare = ""
                    intActLen = 0
                Else
                    strCompare = strCompare & Mid(strText, intPos, 1)
                End If
                If Len(strCompare) = intLen Then
                    If strCompare = strHZ Then
                        blnFind = True
                        intPos = intPos - intActLen
                    Else
                        strCompare = ""
                        intPos = intPos - intActLen + 1
                        intActLen = 0
                    End If
                End If
            End If
            If blnFind Then Exit For
        Next
        '肯定有
        strText = Mid(strText, 1, intPos)
    End If
    
    '再去掉前缀
    If strQZ <> "" Then
        If InStr(1, strText, strQZ) = 0 Then strText = strQZ & strText
        strQZ = Replace(strQZ, vbCrLf, "")
        intEnd = Len(strText)
        intLen = Len(strQZ)
        strCompare = ""
        intActLen = 0
        blnFind = False
        For intPos = 1 To intEnd
            lngASC = Asc(Mid(strText, intPos, 1))
            intActLen = intActLen + 1
            If Not (lngASC = 13 Or lngASC = 10) Then
                If lngASC = 32 Then
                    strCompare = ""
                    intActLen = 0
                Else
                    strCompare = strCompare & Mid(strText, intPos, 1)
                End If
                If Len(strCompare) = intLen Then
                    If strCompare = strQZ Then
                        blnFind = True
                        intPos = intPos + 1
                    Else
                        strCompare = ""
                        intPos = intPos + 1
                        intActLen = 0
                    End If
                End If
            End If
            If blnFind Then Exit For
        Next
        strText = Mid(strText, intPos)
    End If
    
    If IsNumeric(Replace(strText, vbCrLf, "")) Then
        SubstrAnaly = Replace(strText, vbCrLf, "")
    Else
        SubstrAnaly = strText
    End If
End Function

Public Sub SubstrPro(strFormat As String, ByVal strName As String, Optional ByVal intType As Integer = 0)
    Dim i As Integer, j As Integer, l As Integer, r As Integer, strHZ As String, strQZ As String
    'intType=0-删除指定格式串;1-得到指定格式串
    j = Len(strFormat)
    i = InStr(1, strFormat, "[" & strName & "]")
    If i = 0 Then Exit Sub
    
    For l = i To 1 Step -1
        If Mid(strFormat, l, 1) = "{" Then Exit For
    Next
    strQZ = Mid(strFormat, l + 1, i - l - 1)
    For r = i To j
        If Mid(strFormat, r, 1) = "}" Then Exit For
    Next
    strHZ = Mid(strFormat, i + Len(strName) + 2, r - i - Len(strName) - 2)
    If intType = 0 Then
        'strFormat = Mid(strFormat, 1, l - 1) & strQZ & strHZ & Mid(strFormat, r + 1)
        If Mid(strFormat, 1, l - 1) = "" And Mid(strFormat, r + 1) = "" Then
            strFormat = ""
        Else
            strFormat = Mid(strFormat, 1, l - 1) & strQZ & strHZ & Mid(strFormat, r + 1)
        End If
    Else
        strFormat = Mid(strFormat, l, r - l + 1)
    End If
End Sub

Private Function MoveNextCell(Optional ByVal blnNext As Boolean = True, Optional ByVal blnNoMove As Boolean = False, Optional ByVal strText As String = "", Optional ByVal lngDemoRow As Long = 0) As Boolean
    '----------------------------------------------
    '修改人：LPF 2012-04-20
    '修改内容：允许非分组起始行，也可以录入多行数据
    '----------------------------------------------
    Dim arrData
    Dim blnNULL As Boolean                      '是否为空行
    Dim blnGroup As Boolean                     '分组行
    Dim strDate As String, strTime As String, strYear As String    '分组首记录的日期与时间
    Dim strReturn As String, strMsg As String, strPart As String
    Dim lngStart As Long, lngStartGroup As Long, lngMutilRows As Long, lngDeff As Long, intGroupFirstRows As Integer, intBound As Integer, intRowCount As Integer
    Dim intRow As Integer, intRowGroup As Integer, intCount As Integer, intNULL As Integer  '其后有多少空行
    Dim blnTrue As Boolean, blnDate As Boolean, strRows As String, strRowsDel As String
    Dim lngDemo As Long, lngLastNull As Long, lngLastNoNull As Long
    '赋值然后移动到下一个有效单元格
    Dim strKey As String, strField As String, strValue As String, strAppend As String
    Dim blnCallback As Boolean, blnReseGroupAssistant As Boolean, blnGroupAddNum As Boolean '分组数据增加行
    '大文本列和内容信息
    Dim varAssistant() As Variant, strAssistantCols As String
    On Error GoTo ErrHand
    blnReseGroupAssistant = False
    
    '检查数据,不合格就再次弹出要求录入
    If mintType >= 0 Then
        If strText = "" Then
            strReturn = Replace(Replace(Replace(strReturn, Chr(10), ""), Chr(13), ""), Chr(1), "")
            If Not CheckInput(strReturn, strMsg) Then
                RaiseEvent AfterRowColChange(strMsg, True, mblnSign, mblnArchive)
                Exit Function
            End If
            strText = strReturn
        Else
            strReturn = strText
            mstrData = strText
        End If
        '标记当前行为分组行
        blnDate = (InStr(1, "," & mlngYear & "," & mlngDate & "," & mlngTime & ",", "," & VsfData.COL & ",") > 0)
        If mstrGroupRow <> "" And Val(VsfData.TextMatrix(VsfData.ROW, mlngRecord)) = 0 And Val(mstrGroupRow) <= VsfData.ROW Then
            VsfData.TextMatrix(VsfData.ROW, mlngDemo) = VsfData.ROW - Val(mstrGroupRow) + 1
            'blnGroup = True
        Else
            'blnGroup = ((VsfData.TextMatrix(VsfData.ROW, mlngDemo) = "1") And mblnEditAssistant) Or (Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) > 1) '大段文本列才自动分解
        End If
        
        lngDemo = Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo))
        blnGroup = Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) >= 1
        If VsfData.TextMatrix(VsfData.ROW, mlngRowCount) = "" Then VsfData.TextMatrix(VsfData.ROW, mlngRowCount) = "1|1"
        'blnGroup = ((VsfData.TextMatrix(VsfData.ROW, mlngDemo) = "1") And (mblnEditAssistant Or blnDate)) Or (Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) > 1) '大段文本列才自动分解
        '如果修改的是非大文本列或时间列的分组数据，检查修改内容行数是否发生变化，如果变化就当分组数据处理，否则以普通数据处理
        If blnGroup = True And Not (mblnEditAssistant = True Or blnDate = True) Then
            intGroupFirstRows = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0))
            lngStart = GetStartRow(VsfData.ROW)
            '如果编辑的是分组数据最后一个分组行，则当普通数据处理
            If lngStart + intGroupFirstRows < VsfData.Rows Then
                If Val(VsfData.TextMatrix(lngStart + intGroupFirstRows, mlngDemo)) <= 1 Then
                    blnGroup = False: GoTo ErrBegin
                End If
            ElseIf lngStart + intGroupFirstRows >= VsfData.Rows Then
                blnGroup = False
                GoTo ErrBegin
            End If
            
            With txtLength
                '日期与时间列的宽度不管,为了避免返回多行,强制设置为5000
                .Width = VsfData.CellWidth
                .Text = Replace(Replace(Replace(IIf(strReturn = "", "a", strReturn), Chr(10), ""), Chr(13), ""), Chr(1), "")
                .FontName = VsfData.CellFontName
                .FontSize = VsfData.CellFontSize
                .FontBold = VsfData.CellFontBold
                .FontItalic = VsfData.CellFontItalic
            End With
            arrData = GetData(txtLength.Text)
            
            blnGroup = False
            intBound = -1
            If (UBound(arrData) + 1) > intGroupFirstRows Then
                blnGroup = True
            ElseIf (UBound(arrData) + 1) < intGroupFirstRows Then
                '得到本条数据占用最大行的列(不包含大文本项目)
                blnNULL = True
                For intRow = lngStart + intGroupFirstRows - 1 To lngStart Step -1
                    For intCount = 0 To mlngNoEditor - 1
                        If VsfData.ColHidden(intCount) = False And ISEditAssistant(intCount) = False Then
                            If FormatValue(VsfData.TextMatrix(intRow, intCount)) <> "" And Not (IsDiagonal(intCount) And InStr(1, FormatValue(VsfData.TextMatrix(intRow, intCount)), "/") <> 0) Then
                                blnNULL = False
                                If intCount = VsfData.COL Then
                                    intBound = intCount
                                Else
                                    intBound = intCount
                                    Exit For
                                End If
                            End If
                        End If
                    Next intCount
                    If blnNULL = False Then Exit For
                Next intRow
                
                If blnNULL = False Then
                    intNULL = intRow - lngStart + 1
                    If intBound = VsfData.COL Then
                        blnGroup = True
                    Else
                        blnGroup = (intNULL < intGroupFirstRows)
                    End If
                Else
                    blnGroup = True
                End If
            End If
        End If
ErrBegin:
        blnTrue = False
        lngMutilRows = 1
        intGroupFirstRows = 1
        
        If Not blnGroup Then
            If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") <> 0 Then
                lngMutilRows = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0))
            End If
            lngStart = GetStartRow(VsfData.ROW)
            '--58414,刘鹏飞,2013-04-10,添加护士、签名列显示模式
            Call ReSingDataToStart(VsfData, lngStart, lngStart + lngMutilRows - 1)
        Else
            lngMutilRows = 1
            If VsfData.TextMatrix(VsfData.ROW, mlngDemo) = 1 And (mblnEditAssistant Or blnDate) Then
                '记录分组起始行的数据行数
                intGroupFirstRows = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0))
                intBound = VsfData.ROW + intGroupFirstRows - 1
                '--58414,刘鹏飞,2013-04-10,添加护士、签名列显示模式
                Call ReSingDataToStart(VsfData, VsfData.ROW, intBound)
                
                For intCount = VsfData.ROW + intGroupFirstRows To VsfData.Rows - 1
                    '分组中每一个分组子行数据可能占用多行，只有在第一行保存了分组索引。分组总数据行数=每一个分组子行数+每一个分组子行的行数
                    If intCount > intBound Then
                        If Val(VsfData.TextMatrix(intCount, mlngDemo)) <= 1 Then Exit For      '不分组或遇新分组就退出
                        intBound = Val(Split(VsfData.TextMatrix(intCount, mlngRowCount), "|")(0)) + intCount - 1
                        '--58414,刘鹏飞,2013-04-10,添加护士、签名列显示模式
                        Call ReSingDataToStart(VsfData, intCount, intBound)
                    End If
                    lngMutilRows = lngMutilRows + 1
                Next
                lngMutilRows = lngMutilRows + intGroupFirstRows - 1 '保证数据行数的准确性
            Else
                '记录分组起始行的数据行数
                If VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1 >= VsfData.FixedRows Then
                    intGroupFirstRows = Val(Split(VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngRowCount), "|")(0))
                End If
                intGroupFirstRows = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0))
                lngMutilRows = intGroupFirstRows
                '--58414,刘鹏飞,2013-04-10,添加护士、签名列显示模式
                Call ReSingDataToStart(VsfData, VsfData.ROW, VsfData.ROW + lngMutilRows - 1)
            End If
            lngStart = VsfData.ROW
        End If
       
        '准备赋值
        With txtLength
            '日期与时间列的宽度不管,为了避免返回多行,强制设置为5000
            .Width = IIf(VsfData.COL = mlngDate Or VsfData.COL = mlngTime Or VsfData.COL = mlngYear, 5000, VsfData.CellWidth)
            .Text = Replace(Replace(Replace(strReturn, Chr(10), ""), Chr(13), ""), Chr(1), "")
            .FontName = VsfData.CellFontName
            .FontSize = VsfData.CellFontSize
            .FontBold = VsfData.CellFontBold
            .FontItalic = VsfData.CellFontItalic
        End With
        arrData = GetData(txtLength.Text)
        intCount = UBound(arrData)
        If intCount = -1 Then
            arrData = Array()
            ReDim Preserve arrData(UBound(arrData) + 1)
            arrData(UBound(arrData)) = ""
            intCount = 1
        End If
        lngLastNull = VsfData.ROW + intGroupFirstRows - 1: lngLastNoNull = VsfData.ROW + intGroupFirstRows - 1
        '分组数据中可能存在隐藏的行(点击清除功能时),只有是选择大文本才处理
        If blnGroup = True And mblnEditAssistant = True And blnDate = False Then
            If Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) = 1 Then
                intNULL = intGroupFirstRows
                lngDeff = VsfData.ROW + intGroupFirstRows - 1
                For intRow = VsfData.ROW + intGroupFirstRows To VsfData.Rows - 1
                    If intRow > lngDeff Then
                        If Val(VsfData.TextMatrix(intRow, mlngDemo)) <= 1 Or intNULL > intCount Then Exit For     '不分组或遇新分组就退出
                        lngDeff = Val(Split(VsfData.TextMatrix(intRow, mlngRowCount), "|")(0)) + intRow - 1
                    End If
                    If VsfData.RowHidden(intRow) = True Then  '删除的分组列
                        '重新组织文本数组，保证分组数据正常复制
                        ReDim Preserve arrData(UBound(arrData) + 1)
                        For intBound = UBound(arrData) To intRow - VsfData.ROW + 1 Step -1
                            arrData(intBound) = arrData(intBound - 1)
                        Next intBound
                        arrData(intRow - VsfData.ROW) = ""
                        '记录最后一次隐藏的行
                        lngLastNull = intRow
                    Else
                        intNULL = intNULL + 1
                        '记录最后一次没有隐藏的行
                        lngLastNoNull = intRow
                    End If
                Next
            End If
        End If
        intCount = UBound(arrData)
        
        lngDeff = 0
        strRowsDel = ""
        blnGroupAddNum = False
        blnTrue = blnGroup = True And mblnEditAssistant
        If intCount > lngMutilRows - 1 Then
            '对于新增分组数据时，必须要先录入完分组数据才能录入大文本段数据
            If mblnEditAssistant = True And Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) = 1 And lngMutilRows = intGroupFirstRows Then
                strMsg = "新增分组数据时，请先完成数据的分组，最后在录入大文段项目内容！"
                RaiseEvent AfterRowColChange(strMsg, True, mblnSign, mblnArchive)
                strMsg = ""
                Exit Function
            End If
            '往下搜索空行,如果有其它数据行则计算需增加的行数
            '20110830分组号算做同一数据行，将多行文本分解到各行，多余的文本放在统一放在最后一行上;在非首行按回车,只对现有数据进行修改,不对行发生变化
            intNULL = intCount - (lngMutilRows - 1)
            For intRow = lngMutilRows To intCount
                '保证当前输入的内容在一页中显示全
                If intRow + lngStart > VsfData.Rows - 1 Then Exit For
                
                If Val(VsfData.TextMatrix(intRow + lngStart, mlngRecord)) = 0 And VsfData.TextMatrix(intRow + lngStart, mlngRowCount) = "" Then
                    intNULL = intNULL - 1
                    If VsfData.RowHidden(intRow + lngStart) = True Then VsfData.RowHidden(intRow + lngStart) = False
                Else
                    Exit For
                End If
            Next
            '先增加空行
            If intNULL > 0 Then
                lngDeff = intNULL
                VsfData.Rows = VsfData.Rows + intNULL
                '从当前行记录的空白行开始，每行的位置+所增加的空白行数
                For intRow = VsfData.Rows - intNULL - 1 To lngStart + intCount - intNULL + 1 Step -1
                    VsfData.RowPosition(intRow) = intRow + intNULL
                Next
            End If
            
            '当行号发生变化后，需同步更新mrsCellMap中大于该行号的行号数据
            If lngDeff <> 0 Then
                If Not blnGroup Then
                    Call CellMap_Update(lngStart, lngDeff)
                Else
                    Call CellMap_Update(lngStart + lngMutilRows - 1, lngDeff)  '分组行数据从最大一条明细行之后开始处理
                End If
            End If
            '对分组数据最后一个分组行为隐藏行的处理
            '例如：该分组具有2个分组，第一组为一行大文本内容为A，第二组为一行大文本内容为C(改行隐藏).此时添加大文本内容为A、B占两行，此时组织得到本组的数据为A、C、B
            '计算方式为:第一行占用内容+隐藏行内容+多出的内容。此处就会把隐藏行放在最后，最后得到的内容为A、B、C，在下面循环赋值中，第一组就为占用2行内容为A、B
            '说明：如果中间存在隐藏行，最后一组没有隐藏，多出的数据就会追加在最后一组数据的后面
            If (lngLastNull - lngLastNoNull) > 0 Then
                For intRow = lngLastNoNull + 1 To lngLastNull
                    strValue = arrData(lngLastNoNull + 1 - VsfData.ROW)
                    For intBound = lngLastNoNull + 1 - VsfData.ROW To UBound(arrData) - 1
                        arrData(intBound) = arrData(intBound + 1)
                    Next intBound
                    arrData(UBound(arrData)) = strValue
                    VsfData.RowPosition(lngLastNoNull + 1) = lngLastNull + (intCount - (lngMutilRows - 1))
                Next intRow
                '更新记录
                For intRow = lngLastNull To lngLastNoNull + 1 Step -1
                    Call CellMap_Update(intRow, intCount - (lngMutilRows - 1), False)
                Next intRow
            End If
            '循环赋值
            intCount = UBound(arrData)
            intBound = 0
            blnReseGroupAssistant = (blnGroup = True And Not (mblnEditAssistant Or blnDate))
            blnGroupAddNum = blnReseGroupAssistant
            If blnGroup = True And blnDate = False Then strReturn = ""
            For intRow = 0 To intCount
                VsfData.TextMatrix(lngStart + intRow, VsfData.COL) = Replace(Replace(Replace(arrData(intRow), Chr(10), ""), Chr(13), ""), Chr(1), "")
                '修改非分组数据或非大文本或日期的分组数据
                '非分组数据：直接处理计算行数并更新数据
                '非大文本或日期的分组数据：1、直接处理计算行数并更新数据，2、需要重新处理大文段的内容显示位置
                If (Not blnGroup) Then
                    VsfData.TextMatrix(lngStart + intRow, mlngRowCount) = intCount + 1 & "|" & intRow + 1
                    VsfData.TextMatrix(lngStart + intRow, mlngRowCurrent) = intCount + 1
                    If intRow > 0 And intRow < intCount Then
                        If mlngSignName <> -1 Then VsfData.TextMatrix(lngStart + intRow, mlngSignName) = ""
                        If mlngSignTime <> -1 Then VsfData.TextMatrix(lngStart + intRow, mlngSignTime) = ""
                    ElseIf intRow = intCount Then
                        If mlngSignName <> -1 Then VsfData.TextMatrix(lngStart + intRow, mlngSignName) = VsfData.TextMatrix(lngStart, mlngSignName)
                        If mlngSignTime <> -1 Then VsfData.TextMatrix(lngStart + intRow, mlngSignTime) = VsfData.TextMatrix(lngStart, mlngSignTime)
                    End If
                Else
                    '修改分组大文段或日期，需从分组起始行到分组结束行从新整理文本内容显示或日期
                    '分组行的特殊处理,更新内部记录集的代码较多
                    '##########################################
                    If Val(VsfData.TextMatrix(lngStart + intRow, mlngDemo)) >= 1 Then
                        If (mblnEditAssistant = True Or blnDate) And Val(VsfData.TextMatrix(lngStart, mlngDemo)) = 1 Then
                            VsfData.TextMatrix(lngStart + intRow, mlngDemo) = intRow + 1
                        End If
                        intRowCount = 1
                        '获取该分组数据行的行数
                        For intBound = intRow + 1 To intCount
                             If Val(VsfData.TextMatrix(lngStart + intBound, mlngDemo)) >= 1 Then Exit For
                             intRowCount = intRowCount + 1
                        Next intBound
                        intBound = intRow
                        VsfData.TextMatrix(lngStart + intRow, mlngRowCount) = intRowCount & "|1"
                        VsfData.TextMatrix(lngStart + intRow, mlngRowCurrent) = intRowCount
                        If Not blnDate Then strReturn = ""
                    Else
                        VsfData.TextMatrix(lngStart + intRow, mlngRowCurrent) = intRowCount
                        VsfData.TextMatrix(lngStart + intRow, mlngRowCount) = intRowCount & "|" & intRow - intBound + 1
                        If mlngSignName <> -1 Then VsfData.TextMatrix(lngStart + intRow, mlngSignName) = ""
                        If mlngSignTime <> -1 Then VsfData.TextMatrix(lngStart + intRow, mlngSignTime) = ""
                    End If
                    If Not blnDate Then
                        strReturn = strReturn & Replace(Replace(Replace(VsfData.TextMatrix(lngStart + intRow, VsfData.COL), Chr(10), ""), Chr(13), ""), Chr(1), "")
                    End If
                    '到该分组数据行数最后一行才执行更新操作
                    If intRow = intBound + intRowCount - 1 Then
                        If mlngSignName <> -1 Then VsfData.TextMatrix(lngStart + intRow, mlngSignName) = VsfData.TextMatrix(lngStart + intBound, mlngSignName)
                        If mlngSignTime <> -1 Then VsfData.TextMatrix(lngStart + intRow, mlngSignTime) = VsfData.TextMatrix(lngStart + intBound, mlngSignTime)
                        '保存数据
                        strYear = ""
                        If Val(VsfData.TextMatrix(lngStart + intBound, mlngDemo)) > 0 Then
                            If CheckGroupDate(lngStart + intBound) = True Then
                                '保存后的修改才进入此流程，取该条记录的实际时间
                                If mblnDateAd Then
                                    strYear = Format(VsfData.TextMatrix(lngStart + intBound, mlngActTime), "YYYY")
                                    strDate = Format(VsfData.TextMatrix(lngStart + intBound, mlngActTime), "DD") & "/" & Format(VsfData.TextMatrix(lngStart + intBound, mlngActTime), "MM")
                                Else
                                    strDate = Mid(VsfData.TextMatrix(lngStart + intBound, mlngActTime), 1, 10)
                                End If
                                strTime = Mid(VsfData.TextMatrix(lngStart + intBound, mlngActTime), 12, 5)
                            Else
                                '新增时进入此流程
                                strDate = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngDate)
                                strTime = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngTime)
                                If mblnDateAd Then strYear = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngYear)
                            End If
                        Else
                            '普通数据
                            strDate = VsfData.TextMatrix(lngStart + intBound, mlngDate)
                            strTime = VsfData.TextMatrix(lngStart + intBound, mlngTime)
                            If mblnDateAd Then strYear = VsfData.TextMatrix(lngStart + intBound, mlngYear)
                        End If
                        
                        '1\日期
                        strField = "ID|页号|行号|列号|记录ID|数据|部位|删除"
                        If mlngDate <> -1 Then
                            strKey = mint页码 & "," & lngStart + intBound & "," & mlngDate
                            strValue = strKey & "|" & mint页码 & "|" & lngStart + intBound & "|" & mlngDate & "|" & _
                                Val(VsfData.TextMatrix(lngStart + intBound, mlngRecord)) & "|" & strDate & "|" & strYear & "|" & IIf(VsfData.RowHidden(lngStart + intBound) = True, 1, 0)
                            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                        End If
                        '2\时间
                        strKey = mint页码 & "," & lngStart + intBound & "," & mlngTime
                        strValue = strKey & "|" & mint页码 & "|" & lngStart + intBound & "|" & mlngTime & "|" & _
                            Val(VsfData.TextMatrix(lngStart + intBound, mlngRecord)) & "|" & strTime & "|" & _
                            VsfData.TextMatrix(lngStart + intBound, mlngDemo) & "|" & IIf(VsfData.RowHidden(lngStart + intBound) = True, 1, 0)
                        Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                        
                        If Not blnDate Then
                            strField = "ID|页号|行号|列号|记录ID|数据|部位|删除"
                            If InStr(1, "," & mstrCatercorner & ",", "," & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ",") = 0 Then
                                strPart = GetActivePart(VsfData.COL, 0)
                            Else
                                strPart = GetActivePart(VsfData.COL, 0) & "/" & GetActivePart(VsfData.COL, 1)
                            End If
                            strKey = mint页码 & "," & lngStart + intBound & "," & VsfData.COL
                            strValue = strKey & "|" & mint页码 & "|" & lngStart + intBound & "|" & VsfData.COL & "|" & _
                                Val(VsfData.TextMatrix(lngStart + intBound, mlngRecord)) & "|" & strReturn & "|" & strPart & "|" & IIf(VsfData.RowHidden(lngStart + intBound) = True, 1, 0)
                            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                        End If
                    End If
                    '##########################################
                End If
            Next
            
            '所有隐蔽列进行赋值
            intBound = lngStart + intCount
            For intRow = lngStart + 1 To intBound
                For intCount = 0 To VsfData.Cols - 1
                    VsfData.Cell(flexcpForeColor, intRow, intCount) = VsfData.Cell(flexcpForeColor, lngStart, intCount)
                    If VsfData.ColHidden(intCount) And InStr(1, "," & mlngRowCount & "," & mlngRowCurrent & ",", "," & intCount & ",") = 0 Then
                        If blnGroup And InStr(1, "," & mlngDemo & "," & mlngRecord & "," & mlngActTime & ",", "," & intCount & ",") = 0 Then
                            VsfData.TextMatrix(intRow, intCount) = VsfData.TextMatrix(lngStart, intCount)
                        End If
                    End If
                Next
            Next
            lngMutilRows = lngStart + lngMutilRows - 1
        Else
            blnReseGroupAssistant = False
            If blnGroup = True And blnDate = False Then strReturn = ""
            '对该列重新赋值（当只输入一个数字时，不知为何会产生字符ASCII码为1的符号）
            For intRow = 0 To intCount
                VsfData.TextMatrix(lngStart + intRow, VsfData.COL) = Replace(Replace(Replace(arrData(intRow), Chr(10), ""), Chr(13), ""), Chr(1), "")
                If blnGroup = True Then
                    '分组行的特殊处理,更新内部记录集的代码较多
                    '##########################################
                    If Val(VsfData.TextMatrix(lngStart + intRow, mlngDemo)) >= 1 Then
                        If (mblnEditAssistant = True Or blnDate = True) And Val(VsfData.TextMatrix(lngStart, mlngDemo)) = 1 Then
                            VsfData.TextMatrix(lngStart + intRow, mlngDemo) = intRow + 1
                        End If
                        intRowCount = 1
                        '获取该分组数据行的行数
                        For intBound = intRow + 1 To intCount
                             If Val(VsfData.TextMatrix(lngStart + intBound, mlngDemo)) >= 1 Then Exit For
                             intRowCount = intRowCount + 1
                        Next intBound
                        intBound = intRow
                        If Not blnDate Then strReturn = ""
                    End If
                    If Not blnDate Then
                        strReturn = strReturn & Replace(Replace(Replace(VsfData.TextMatrix(lngStart + intRow, VsfData.COL), Chr(10), ""), Chr(13), ""), Chr(1), "")
                    End If
                    '到该分组数据行数最后一行才执行更新操作
                    If intRow = intBound + intRowCount - 1 Then
                        '保存数据
                        strYear = ""
                        If Val(VsfData.TextMatrix(lngStart + intBound, mlngDemo)) > 0 Then
                            If CheckGroupDate(lngStart + intBound) = True Then
                                '保存后的修改才进入此流程，取该条记录的实际时间
                                If mblnDateAd Then
                                    strYear = Format(VsfData.TextMatrix(lngStart + intBound, mlngActTime), "YYYY")
                                    strDate = Format(VsfData.TextMatrix(lngStart + intBound, mlngActTime), "DD") & "/" & Format(VsfData.TextMatrix(lngStart + intBound, mlngActTime), "MM")
                                Else
                                    strDate = Mid(VsfData.TextMatrix(lngStart + intBound, mlngActTime), 1, 10)
                                End If
                                strTime = Mid(VsfData.TextMatrix(lngStart + intBound, mlngActTime), 12, 5)
                            Else
                                '新增时进入此流程
                                strDate = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngDate)
                                strTime = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngTime)
                                If mblnDateAd Then strYear = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngYear)
                            End If
                        Else
                            '普通数据
                            strDate = VsfData.TextMatrix(lngStart + intBound, mlngDate)
                            strTime = VsfData.TextMatrix(lngStart + intBound, mlngTime)
                            If mblnDateAd Then strYear = VsfData.TextMatrix(lngStart + intBound, mlngYear)
                        End If
                        
                        '1\日期
                        strField = "ID|页号|行号|列号|记录ID|数据|部位|删除"
                        If mlngDate <> -1 Then
                            strKey = mint页码 & "," & lngStart + intBound & "," & mlngDate
                            strValue = strKey & "|" & mint页码 & "|" & lngStart + intBound & "|" & mlngDate & "|" & _
                                Val(VsfData.TextMatrix(lngStart + intBound, mlngRecord)) & "|" & strDate & "|" & strYear & "|" & IIf(VsfData.RowHidden(lngStart + intBound) = True, 1, 0)
                            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                        End If
                        '2\时间
                        strKey = mint页码 & "," & lngStart + intBound & "," & mlngTime
                        strValue = strKey & "|" & mint页码 & "|" & lngStart + intBound & "|" & mlngTime & "|" & _
                            Val(VsfData.TextMatrix(lngStart + intBound, mlngRecord)) & "|" & strTime & "|" & _
                            VsfData.TextMatrix(lngStart + intBound, mlngDemo) & "|" & IIf(VsfData.RowHidden(lngStart + intBound) = True, 1, 0)
                        Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                        
                        If Not blnDate Then
                            strField = "ID|页号|行号|列号|记录ID|数据|部位|删除"
                            If InStr(1, "," & mstrCatercorner & ",", "," & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ",") = 0 Then
                                strPart = GetActivePart(VsfData.COL, 0)
                            Else
                                strPart = GetActivePart(VsfData.COL, 0) & "/" & GetActivePart(VsfData.COL, 1)
                            End If
                            strKey = mint页码 & "," & lngStart + intBound & "," & VsfData.COL
                            strValue = strKey & "|" & mint页码 & "|" & lngStart + intBound & "|" & VsfData.COL & "|" & _
                                Val(VsfData.TextMatrix(lngStart + intBound, mlngRecord)) & "|" & strReturn & "|" & strPart & "|" & IIf(VsfData.RowHidden(lngStart + intBound) = True, 1, 0)
                            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                        End If
                    End If
                    '##########################################
                End If
            Next
            strRows = ""
            strRowsDel = ""
            lngStartGroup = -1
            For intRow = intCount + 1 To lngMutilRows - 1
                VsfData.TextMatrix(lngStart + intRow, VsfData.COL) = ""
            Next intRow
            blnReseGroupAssistant = False
            If intCount < (lngMutilRows - 1) Then
                blnReseGroupAssistant = (blnGroup And Not (mblnEditAssistant Or blnDate))
            End If
            
            For intRow = intCount + 1 To lngMutilRows - 1
                '分组行的特殊处理,更新内部记录集的代码较多
                '##########################################
                '保存数据
                If (blnGroup And (mblnEditAssistant Or blnDate)) Then
                    '获取改行起始行
                    If lngStartGroup <> GetStartRow(lngStart + intRow) Then
                        intNULL = GetStartRow(lngStart + intRow)
                        '寻找的起始列mlngDemo肯定>0
                        If Val(VsfData.TextMatrix(intNULL, mlngDemo)) <= 0 Then
                            For intRowGroup = lngStart + intRow To lngStart Step -1
                                If Val(VsfData.TextMatrix(intRowGroup, mlngDemo)) > 0 Then
                                    intNULL = intRowGroup
                                    Exit For
                                End If
                            Next intRowGroup
                            If intNULL = lngStartGroup Then GoTo ErrDemo
                        End If
                        lngStartGroup = intNULL
                        '根据行数据重新填写行序列,intNULL记录最后一条不为空行的行号
                        If VsfData.TextMatrix(lngStartGroup, mlngRowCount) = "" Then VsfData.TextMatrix(lngStartGroup, mlngRowCount) = "1|1"
                        intGroupFirstRows = Val(Split(VsfData.TextMatrix(lngStartGroup, mlngRowCount), "|")(0))
                        intNULL = lngStartGroup + intGroupFirstRows - 1
                        For intRowGroup = intNULL To lngStartGroup Step -1
                            blnNULL = True
                            For intBound = 0 To VsfData.Cols - 1
                                If Not VsfData.ColHidden(intBound) And intBound < mlngNoEditor Then
                                    If VsfData.TextMatrix(intRowGroup, intBound) <> "" And Not (IsDiagonal(intBound) And InStr(1, VsfData.TextMatrix(intRowGroup, intBound), "/") <> 0) Then
                                        blnNULL = False
                                        Exit For
                                    End If
                                End If
                            Next
                            If Not blnNULL Then Exit For
                            intNULL = intNULL - 1
                            If intRowGroup = lngStartGroup Then
                                 intNULL = intNULL + 1
                            Else
                                If InStr(1, strRows & ",", "," & intRowGroup & ",") = 0 Then strRows = strRows & "," & intRowGroup
                            End If
                        Next intRowGroup
                        
                        '重新填写数据行数
                        For intRowGroup = lngStartGroup To intNULL
                            VsfData.TextMatrix(intRowGroup, mlngRowCount) = intNULL - lngStartGroup + 1 & "|" & intRowGroup - lngStartGroup + 1
                            VsfData.TextMatrix(intRowGroup, mlngRowCurrent) = intNULL - lngStartGroup + 1
                        Next intRowGroup
                        If mlngSignName <> -1 Then
                            If Trim(VsfData.TextMatrix(lngStartGroup + intGroupFirstRows - 1, mlngSignName)) <> "" Then
                                VsfData.TextMatrix(intNULL, mlngSignName) = VsfData.TextMatrix(lngStartGroup + intGroupFirstRows - 1, mlngSignName)
                                If mlngSignTime <> -1 Then VsfData.TextMatrix(intNULL, mlngSignTime) = VsfData.TextMatrix(lngStartGroup + intGroupFirstRows - 1, mlngSignTime)
                            End If
                        End If
                        For intRowGroup = intNULL + 1 To lngStartGroup + intGroupFirstRows - 1
                            VsfData.TextMatrix(intRowGroup, mlngRowCount) = ""
                            VsfData.TextMatrix(intRowGroup, mlngRowCurrent) = ""
                            VsfData.TextMatrix(intRowGroup, mlngRecord) = ""
                            If mlngSignName <> -1 Then VsfData.TextMatrix(intRowGroup, mlngSignName) = ""
                            If mlngOperator <> -1 Then VsfData.TextMatrix(intRowGroup, mlngOperator) = ""
                            If mlngSignTime <> -1 Then VsfData.TextMatrix(intRowGroup, mlngSignTime) = ""
                        Next
                    End If
ErrDemo:
                    If Val(VsfData.TextMatrix(lngStart + intRow, mlngDemo)) > 0 And intRow > intCount Then
                        VsfData.TextMatrix(lngStart + intRow, mlngDemo) = intRow + 1
                        strYear = ""
                        If CheckGroupDate(lngStart + intRow) = True Then
                            '保存后的修改才进入此流程，取该条记录的实际时间
                            If mblnDateAd Then
                                strYear = Format(VsfData.TextMatrix(lngStart + intRow, mlngActTime), "YYYY")
                                strDate = Format(VsfData.TextMatrix(lngStart + intRow, mlngActTime), "DD") & "/" & Format(VsfData.TextMatrix(lngStart + intRow, mlngActTime), "MM")
                            Else
                                strDate = Mid(VsfData.TextMatrix(lngStart + intRow, mlngActTime), 1, 10)
                            End If
                            strTime = Mid(VsfData.TextMatrix(lngStart + intRow, mlngActTime), 12, 5)
                        Else
                            '新增时进入此流程
                            strDate = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngDate)
                            strTime = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngTime)
                            If mblnDateAd Then strYear = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngYear)
                        End If
                    
                        '分组起始行的行数减少时，重新设置分组号
                        If Left(strRows, 1) = "," Then strRows = Mid(strRows, 2)
                        If strRows <> "" Then
                            intNULL = 0
                            For intBound = 0 To UBound(Split(strRows, ","))
                                If Val(Split(strRows, ",")(intBound)) < (lngStart + intRow) Then
                                    intNULL = intNULL + 1
                                End If
                            Next intBound
                            VsfData.TextMatrix(lngStart + intRow, mlngDemo) = Val(VsfData.TextMatrix(lngStart + intRow, mlngDemo)) - intNULL
                        End If
                        
                        '1\日期
                        strField = "ID|页号|行号|列号|记录ID|数据|部位|删除"
                        If mlngDate <> -1 Then
                            strKey = mint页码 & "," & lngStart + intRow & "," & mlngDate
                            strValue = strKey & "|" & mint页码 & "|" & lngStart + intRow & "|" & mlngDate & "|" & _
                                Val(VsfData.TextMatrix(lngStart + intRow, mlngRecord)) & "|" & strDate & "|" & strYear & "|" & IIf(VsfData.RowHidden(lngStart + intRow) = True, 1, 0)
                            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                        End If
                        '2\时间
                        strKey = mint页码 & "," & lngStart + intRow & "," & mlngTime
                        strValue = strKey & "|" & mint页码 & "|" & lngStart + intRow & "|" & mlngTime & "|" & _
                            Val(VsfData.TextMatrix(lngStart + intRow, mlngRecord)) & "|" & strTime & "|" & _
                            VsfData.TextMatrix(lngStart + intRow, mlngDemo) & "|" & IIf(VsfData.RowHidden(lngStart + intRow) = True, 1, 0)
                        Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                        
                        If Not blnDate Then
                            strField = "ID|页号|行号|列号|记录ID|数据|部位|删除"
                            If InStr(1, "," & mstrCatercorner & ",", "," & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ",") = 0 Then
                                strPart = GetActivePart(VsfData.COL, 0)
                            Else
                                strPart = GetActivePart(VsfData.COL, 0) & "/" & GetActivePart(VsfData.COL, 1)
                            End If
                            strKey = mint页码 & "," & lngStart + intRow & "," & VsfData.COL
                            strValue = strKey & "|" & mint页码 & "|" & lngStart + intRow & "|" & VsfData.COL & "|" & _
                                Val(VsfData.TextMatrix(lngStart + intRow, mlngRecord)) & "|" & "" & "|" & strPart & "|1"
                            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                        End If
                    End If
                End If
                '##########################################
            Next
            '修改分组非大文本或日期行时，需要获取分组数据大文本段内容信息，重新组织文本显示
            '如有3组数据，第二2行有3行，修改为1行，第3组数据应该紧接着显示在第2组下面(第二组此时只有1行)
            If blnReseGroupAssistant = True Then Call GetGroupAssistant(strAssistantCols, varAssistant)
            lngMutilRows = Val(Split(VsfData.TextMatrix(lngStart, mlngRowCount), "|")(0))
            '根据行数据重新填写行序列,intNULL记录最后一条不为空行的行号
            intNULL = lngStart + lngMutilRows - 1
            For intRow = lngMutilRows To 1 Step -1
                blnNULL = True
                For intCount = 0 To VsfData.Cols - 1
                    If Not VsfData.ColHidden(intCount) And intCount < mlngNoEditor And IIf(blnReseGroupAssistant = True, ISEditAssistant(intCount) = False, True) Then
                        If VsfData.TextMatrix(intRow + lngStart - 1, intCount) <> "" And Not (IsDiagonal(intCount) And InStr(1, VsfData.TextMatrix(intRow + lngStart - 1, intCount), "/") <> 0) Then
                            blnNULL = False
                            Exit For
                        End If
                    End If
                Next
                
                If Not blnNULL Then Exit For
                intNULL = intNULL - 1
            Next
            '从新填写行序号
            If Not blnGroup Then
                If intNULL < lngStart Then intNULL = lngStart
                For intRow = lngStart To intNULL
                    VsfData.TextMatrix(intRow, mlngRowCount) = (intNULL - lngStart + 1) & "|" & intRow - lngStart + 1
                    VsfData.TextMatrix(intRow, mlngRowCurrent) = (intNULL - lngStart + 1)
                Next
                If mlngSignName <> -1 Then
                    If Trim(VsfData.TextMatrix(lngMutilRows + lngStart - 1, mlngSignName)) <> "" Then
                        VsfData.TextMatrix(intNULL, mlngSignName) = VsfData.TextMatrix(lngMutilRows + lngStart - 1, mlngSignName)
                        If mlngSignTime <> -1 Then VsfData.TextMatrix(intNULL, mlngSignTime) = VsfData.TextMatrix(lngMutilRows + lngStart - 1, mlngSignTime)
                    End If
                End If
                strRows = ""
            Else '分组行以保存的数据删除时，不清空行号
                For intRow = intNULL + 1 To lngStart + lngMutilRows - 1
                    If intRow = lngStart Then intNULL = intNULL + 1
                Next intRow
                
                For intRow = lngStart To intNULL
                    VsfData.TextMatrix(intRow, mlngRowCount) = (intNULL - lngStart + 1) & "|" & intRow - lngStart + 1
                    VsfData.TextMatrix(intRow, mlngRowCurrent) = (intNULL - lngStart + 1)
                Next
                If mlngSignName <> -1 Then
                    If Trim(VsfData.TextMatrix(lngStart + lngMutilRows - 1, mlngSignName)) <> "" Then
                        VsfData.TextMatrix(intNULL, mlngSignName) = VsfData.TextMatrix(lngStart + lngMutilRows - 1, mlngSignName)
                        If mlngSignTime <> -1 Then VsfData.TextMatrix(intNULL, mlngSignTime) = VsfData.TextMatrix(lngStart + lngMutilRows - 1, mlngSignTime)
                    End If
                End If
            End If
            If Left(Trim(strRows), 1) = "," Then strRows = Mid(Trim(strRows), 2)
            If Right(strRows, 1) = "," Then strRows = Mid(strRows, 1, Len(strRows) - 1)
            For intRow = intNULL + 1 To lngStart + lngMutilRows - 1
                VsfData.TextMatrix(intRow, mlngRowCount) = ""
                VsfData.TextMatrix(intRow, mlngRowCurrent) = ""
                VsfData.TextMatrix(intRow, mlngRecord) = ""
                If mlngSignName <> -1 Then VsfData.TextMatrix(intRow, mlngSignName) = ""
                If mlngOperator <> -1 Then VsfData.TextMatrix(intRow, mlngOperator) = ""
                If mlngSignTime <> -1 Then VsfData.TextMatrix(intRow, mlngSignTime) = ""
                If blnReseGroupAssistant = True Then
                    If InStr(1, strRows & ",", "," & intRow & ",") = 0 Then strRows = strRows & "," & intRow
                ElseIf Not blnGroup Then
                    If InStr(1, strRows & ",", "," & intRow & ",") = 0 Then strRows = strRows & "," & intRow
                End If
            Next
            '更新记录集大文段信息
            If blnReseGroupAssistant = True Then Call CellMap_UpdateAssistant(lngStart)
        End If
        
        '获取分组起始行所有行信息
        If blnTrue = True Then 'blnTrue为真说明选择的是分组行的起始行，并且是大文本段
            strReturn = ""
            intCount = Val(Split(VsfData.TextMatrix(lngStart, mlngRowCount), "|")(0))
            For intRow = 0 To intCount - 1
                strReturn = strReturn & Replace(Replace(Replace(VsfData.TextMatrix(VsfData.ROW + intRow, VsfData.COL), Chr(10), ""), Chr(13), ""), Chr(1), "")
            Next intRow
        End If
        
        If mstrData <> strReturn Or blnTrue = True Then
            If strText <> mstrData Then mblnChange = True
            '同步保存日期与时间列的数据
            If Val(VsfData.TextMatrix(lngStart, mlngCollectType)) >= 0 Then
                strYear = ""
                If Val(VsfData.TextMatrix(lngStart, mlngDemo)) > 0 Then
                    If CheckGroupDate(lngStart) = True Then
                        '保存后的修改才进入此流程，取该条记录的实际时间
                        If mblnDateAd Then
                            strYear = Format(VsfData.TextMatrix(lngStart, mlngActTime), "YYYY")
                            strDate = Format(VsfData.TextMatrix(lngStart, mlngActTime), "DD") & "/" & Format(VsfData.TextMatrix(lngStart, mlngActTime), "MM")
                        Else
                            strDate = Mid(VsfData.TextMatrix(lngStart, mlngActTime), 1, 10)
                        End If
                        strTime = Mid(VsfData.TextMatrix(lngStart, mlngActTime), 12, 5)
                    Else
                        '新增时进入此流程
                        strDate = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngDate)
                        strTime = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngTime)
                        If mblnDateAd Then strYear = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngYear)
                    End If
                Else
                    '普通数据
                    strDate = VsfData.TextMatrix(lngStart, mlngDate)
                    strTime = VsfData.TextMatrix(lngStart, mlngTime)
                    If mblnDateAd Then strYear = VsfData.TextMatrix(lngStart, mlngYear)
                End If
                
                '1\日期
                strField = "ID|页号|行号|列号|记录ID|数据|部位|删除"
                If mlngDate <> -1 Then
                    strKey = mint页码 & "," & lngStart & "," & mlngDate
                    strValue = strKey & "|" & mint页码 & "|" & lngStart & "|" & mlngDate & "|" & _
                        Val(VsfData.TextMatrix(lngStart, mlngRecord)) & "|" & strDate & "|" & strYear & "|0"
                    Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                End If
                '2\时间
                strKey = mint页码 & "," & lngStart & "," & mlngTime
                strValue = strKey & "|" & mint页码 & "|" & lngStart & "|" & mlngTime & "|" & _
                    Val(VsfData.TextMatrix(lngStart, mlngRecord)) & "|" & strTime & "|" & _
                    VsfData.TextMatrix(lngStart, mlngDemo) & "|0"
                Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
            Else
                strField = "ID|页号|行号|列号|记录ID|数据|汇总|删除"
                strKey = mint页码 & "," & lngStart & "," & mlngDate
                strValue = strKey & "|" & mint页码 & "|" & lngStart & "|" & mlngDate & "|" & Val(VsfData.TextMatrix(lngStart, mlngRecord)) & "|" & _
                        VsfData.TextMatrix(lngStart, mlngCollectText) & ";" & VsfData.TextMatrix(lngStart, mlngCollectType) & ";" & _
                        VsfData.TextMatrix(lngStart, mlngCollectStyle) & ";" & VsfData.TextMatrix(lngStart, mlngCollectDay) & ";" & _
                    VsfData.TextMatrix(lngStart, mlngCollectStart) & ";" & VsfData.TextMatrix(lngStart, mlngCollectEnd) & "|1|0"
                Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                
            End If
            
            If (Not blnGroup Or blnTrue) And Not blnDate Then
                '记录用户修改过的单元格
                If InStr(1, "," & mstrCatercorner & ",", "," & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ",") = 0 Then
                    strPart = GetActivePart(VsfData.COL, 0)
                Else
                    strPart = GetActivePart(VsfData.COL, 0) & "/" & GetActivePart(VsfData.COL, 1)
                End If
                
                strField = "ID|页号|行号|列号|记录ID|数据|部位|删除"
                strKey = mint页码 & "," & lngStart & "," & VsfData.COL
                strValue = strKey & "|" & mint页码 & "|" & lngStart & "|" & VsfData.COL & "|" & _
                    Val(VsfData.TextMatrix(lngStart, mlngRecord)) & "|" & strReturn & "|" & strPart & "|" & IIf(strReturn = "", "1", "0")
                Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
            End If
        End If
        
        '52953,刘鹏飞,2012-08-24,汇总数据为0也要显示,避免相邻数据列合并，关联问题:60792
        '避免汇总数据行其他列数据合并
        If Val(VsfData.TextMatrix(lngStart, mlngCollectType)) < 0 Then
            VsfData.TextMatrix(lngStart, VsfData.COL) = FormatValue(VsfData.TextMatrix(lngStart, VsfData.COL))
            If Trim(VsfData.TextMatrix(lngStart, VsfData.COL)) <> "" Then
                '66085:刘鹏飞,2012-09-26,避免相邻汇总列合并,将原来的列内容+空格同一改成在列后面在chr(13)
                '避免因加空格后列宽不够导致内容显示不完全(主要针对右对其)
'                Select Case VsfData.ColAlignment(VsfData.COL)
'                    Case 6, 7, 8
'                        VsfData.TextMatrix(lngStart, VsfData.COL) = IIf(VsfData.COL Mod 2 = 1, " ", String(2, " ")) & VsfData.TextMatrix(lngStart, VsfData.COL)
'                    Case 3, 4, 5
'                        VsfData.TextMatrix(lngStart, VsfData.COL) = IIf(VsfData.COL Mod 2 = 1, " ", String(2, " ")) & VsfData.TextMatrix(lngStart, VsfData.COL) & IIf(VsfData.COL Mod 2 = 1, " ", String(2, " "))
'                    Case 0, 1, 2
'                        VsfData.TextMatrix(lngStart, VsfData.COL) = VsfData.TextMatrix(lngStart, VsfData.COL) & IIf(VsfData.COL Mod 2 = 1, " ", String(2, " "))
'                    Case Else
'                        VsfData.TextMatrix(lngStart, VsfData.COL) = IIf(VsfData.COL Mod 2 = 1, " ", String(2, " ")) & VsfData.TextMatrix(lngStart, VsfData.COL)
'                End Select
                VsfData.TextMatrix(lngStart, VsfData.COL) = VsfData.TextMatrix(lngStart, VsfData.COL) & IIf(VsfData.COL Mod 2 = 1, Chr(13), "")
            End If
        End If
            
        RaiseEvent AfterDataChanged(mblnChange Or mblnVerify)
    End If
    '数据行数减少时，将空白行移至到最后一行
    If Left(Trim(strRows), 1) = "," Then strRows = Mid(Trim(strRows), 2)
    If Right(strRows, 1) = "," Then strRows = Mid(strRows, 1, Len(strRows) - 1)
    strRows = Replace("," & strRows & ",", "," & lngDemoRow & ",", "") '不能删除要追加的行
    For intRow = VsfData.Rows - 1 To VsfData.FixedRows Step -1
        If InStr(1, strRows, "," & intRow & ",") <> 0 Then
            Call CellMap_Update(intRow, -1)
            VsfData.TextMatrix(intRow, mlngDemo) = ""
        End If
    Next intRow
    
    For intRow = VsfData.Rows - 1 To VsfData.FixedRows Step -1
        If InStr(1, strRows, "," & intRow & ",") <> 0 Then
            '清空改行所有信息
            For intBound = 0 To VsfData.Cols - 1
                VsfData.TextMatrix(intRow, intBound) = ""
            Next intBound
            VsfData.RowPosition(intRow) = VsfData.Rows - 1
        End If
    Next intRow
    
    '--58414,刘鹏飞,2013-04-10,添加护士、签名列显示模式
    If Val(VsfData.TextMatrix(lngStart, mlngDemo)) = 1 Then
        '记录分组起始行的数据行数
        intGroupFirstRows = Val(Split(VsfData.TextMatrix(lngStart, mlngRowCount), "|")(0))
        intBound = lngStart + intGroupFirstRows - 1
        Call SingerShowType(VsfData, lngStart, intBound)
        For intCount = lngStart + intGroupFirstRows To VsfData.Rows - 1
            '分组中每一个分组子行数据可能占用多行，只有在第一行保存了分组索引。分组总数据行数=每一个分组子行数+每一个分组子行的行数
            If intCount > intBound Then
                If Val(VsfData.TextMatrix(intCount, mlngDemo)) <= 1 Then Exit For      '不分组或遇新分组就退出
                intBound = Val(Split(VsfData.TextMatrix(intCount, mlngRowCount), "|")(0)) + intCount - 1
                '--58414,刘鹏飞,2013-04-10,添加护士、签名列显示模式
                Call SingerShowType(VsfData, intCount, intBound)
            End If
        Next
    Else
        intGroupFirstRows = Val(Split(VsfData.TextMatrix(lngStart, mlngRowCount), "|")(0))
        '--58414,刘鹏飞,2013-04-10,添加护士、签名列显示模式
        Call SingerShowType(VsfData, lngStart, lngStart + intGroupFirstRows - 1)
    End If
    
    'Call OutputRsData(mrsCellMap)

    '重新组织分组数据内容
    If blnReseGroupAssistant = True Then
        If blnGroupAddNum = True Then Call GetGroupAssistant(strAssistantCols, varAssistant)
        If strAssistantCols <> "" Then
            Call ReSetGroupAssistant(blnNoMove, blnNext, strAssistantCols, varAssistant)
        Else
            Call ReSetGroupDemo(lngStart)
        End If
    End If
    
    MoveNextCell = True
    
    If blnNoMove Then Exit Function
    If blnNext Then
toMoveNextCol:
        If VsfData.COL < mlngNoEditor - 1 Then       '护理记录单肯定有护士签名列
            VsfData.COL = VsfData.COL + 1
            If VsfData.ColWidth(VsfData.COL) = 0 Or VsfData.ColHidden(VsfData.COL) Or mintType = -1 Then GoTo toMoveNextCol
        Else
toMoveNextRow:
            '跳到下一行
            If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") <> 0 Then
                intRow = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0))
                intRow = intRow - Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(1)) + 1
            Else
                intRow = 1
            End If
            mblnShow = False
            If VsfData.ROW + intRow < VsfData.Rows Then
                '只有在追加的模式下，才清除mstrGroupRow
                If mblnGroupApp Then
                    mblnGroupApp = False
                    mstrGroupRow = ""
                End If
                VsfData.ROW = VsfData.ROW + intRow
            End If
            If VsfData.RowHidden(VsfData.ROW) Then
                If VsfData.ROW < VsfData.Rows - 1 Then
                    GoTo toMoveNextRow
                Else
                    For intRow = VsfData.Rows - 1 To VsfData.FixedRows Step -1
                        If VsfData.RowHidden(intRow) = False Then
                            VsfData.ROW = GetStartRow(intRow)
                            Exit For
                        End If
                    Next intRow
                End If
            End If
            mblnShow = True
            VsfData.COL = IIf(mlngDate > 0, mlngDate, mlngTime)
            If Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) > 0 Then VsfData.COL = VsfData.COL + 2
        End If
    Else
toMovePrevCol:
        If VsfData.COL > mlngDate Then      '护理记录单肯定有护士签名列
            VsfData.COL = VsfData.COL - 1
            If VsfData.ColWidth(VsfData.COL) = 0 Or VsfData.ColHidden(VsfData.COL) Or mintType = -1 Then GoTo toMovePrevCol
        Else
toMovePrevRow:
'            '跳到上一行
'            intRow = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0))
'            intRow = intRow - Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(1)) + 1
'            If VsfData.ROW + intRow < VsfData.Rows Then
'                VsfData.ROW = VsfData.ROW + intRow
'            End If
'            If VsfData.RowHidden(VsfData.ROW) Then GoTo toMovePrevRow
'            VsfData.COL = IIf(mlngDate > 0, mlngDate, mlngTime)
        End If
    End If
    If VsfData.ColIsVisible(VsfData.COL) = False Then
        VsfData.LeftCol = VsfData.COL
    End If
    If VsfData.RowIsVisible(VsfData.ROW) = False Then
        VsfData.TopRow = VsfData.ROW
    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function CheckGroupDate(ByVal lngRow As Long) As Boolean
'--功能：检查分组数据起始行时间和保存时间是否相等
    Dim strDate As String, strTime As String, strYear As String
    Dim strDate1 As String, strTime1 As String, strYear1 As String
    Dim lngStart As Long
    
    lngStart = lngRow - Val(VsfData.TextMatrix(lngRow, mlngDemo)) + 1
    
    If Val(VsfData.TextMatrix(lngRow, mlngRecord)) > 0 Then
        If Val(VsfData.TextMatrix(lngStart, mlngDemo)) <> 1 Then CheckGroupDate = True: Exit Function
        strDate = VsfData.TextMatrix(lngStart, mlngDate)
        strTime = VsfData.TextMatrix(lngStart, mlngTime)
        strYear = VsfData.TextMatrix(lngStart, mlngYear)
        If mblnDateAd Then
            strDate1 = Format(VsfData.TextMatrix(lngStart, mlngActTime), "DD") & "/" & Format(VsfData.TextMatrix(lngStart, mlngActTime), "MM")
        Else
            strDate1 = Mid(VsfData.TextMatrix(lngStart, mlngActTime), 1, 10)
        End If
        strTime1 = Mid(VsfData.TextMatrix(lngStart, mlngActTime), 12, 5)
        strYear1 = Format(VsfData.TextMatrix(lngStart, mlngActTime), "YYYY")
        If strDate <> strDate1 Or strTime <> strTime1 Or IIf(mblnDateAd, strYear <> strYear1, False) Then
            CheckGroupDate = False
        Else
            CheckGroupDate = True
        End If
    Else
        CheckGroupDate = False
    End If
End Function

Private Sub CellMap_UpdateAssistant(ByVal lngStartRow As Long)
'功能：更新记录集大文段信息
    Dim strDate As String, strTime As String, strYear As String
    Dim strKey As String, strField As String, strValue As String, strPart As String
    Dim lngCol As Long, lngRow As Long, lngRowCount As Long, strReturn As String
    
    On Error GoTo ErrHand
    
    If VsfData.TextMatrix(lngStartRow, mlngRowCount) = "" Then VsfData.TextMatrix(lngStartRow, mlngRowCount) = "1|1"
    lngRowCount = Val(Split(VsfData.TextMatrix(lngStartRow, mlngRowCount), "|")(0))
    
    strYear = ""
    If Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) > 0 Then
        If CheckGroupDate(lngStartRow) = True Then
            '保存后的修改才进入此流程，取该条记录的实际时间
            If mblnDateAd Then
                strYear = Format(VsfData.TextMatrix(lngStartRow, mlngActTime), "YYYY")
                strDate = Format(VsfData.TextMatrix(lngStartRow, mlngActTime), "DD") & "/" & Format(VsfData.TextMatrix(lngStartRow, mlngActTime), "MM")
            Else
                strDate = Mid(VsfData.TextMatrix(lngStartRow, mlngActTime), 1, 10)
            End If
            strTime = Mid(VsfData.TextMatrix(lngStartRow, mlngActTime), 12, 5)
        Else
            '新增时进入此流程
            strDate = VsfData.TextMatrix(lngStartRow - Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) + 1, mlngDate)
            strTime = VsfData.TextMatrix(lngStartRow - Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) + 1, mlngTime)
            If mblnDateAd Then strYear = VsfData.TextMatrix(lngStartRow - Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) + 1, mlngYear)
        End If
    Else
        '普通数据
        strDate = VsfData.TextMatrix(lngStartRow, mlngDate)
        strTime = VsfData.TextMatrix(lngStartRow, mlngTime)
        If mblnDateAd Then strYear = VsfData.TextMatrix(lngStartRow, mlngYear)
    End If
    
    '1\日期
    strField = "ID|页号|行号|列号|记录ID|数据|部位|删除"
    If mlngDate <> -1 Then
        strKey = mint页码 & "," & lngStartRow & "," & mlngDate
        strValue = strKey & "|" & mint页码 & "|" & lngStartRow & "|" & mlngDate & "|" & _
            Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "|" & strDate & "|" & strYear & "|0"
        Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
    End If
    '2\时间
    strKey = mint页码 & "," & lngStartRow & "," & mlngTime
    strValue = strKey & "|" & mint页码 & "|" & lngStartRow & "|" & mlngTime & "|" & _
        Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "|" & strTime & "|" & _
        VsfData.TextMatrix(lngStartRow, mlngDemo) & "|0"
    Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
    
    For lngCol = mlngTime + 1 To mlngNoEditor - 1
        If ISEditAssistant(lngCol) Then
            strReturn = ""
            For lngRow = 0 To lngRowCount - 1
                strReturn = strReturn & Replace(Replace(Replace(VsfData.TextMatrix(lngStartRow + lngRow, lngCol), Chr(10), ""), Chr(13), ""), Chr(1), "")
            Next lngRow
            '记录用户修改过的单元格
            If InStr(1, "," & mstrCatercorner & ",", "," & lngCol - (cHideCols + VsfData.FixedCols - 1) & ",") = 0 Then
                strPart = GetActivePart(lngCol, 0)
            Else
                strPart = GetActivePart(lngCol, 0) & "/" & GetActivePart(lngCol, 1)
            End If
            
            strField = "ID|页号|行号|列号|记录ID|数据|部位|删除"
            strKey = mint页码 & "," & lngStartRow & "," & lngCol
            strValue = strKey & "|" & mint页码 & "|" & lngStartRow & "|" & lngCol & "|" & _
                Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "|" & strReturn & "|" & strPart & "|" & IIf(strReturn = "", "1", "0")
            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
        End If
    Next lngCol
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub ReSetGroupAssistant(blnNoMove As Boolean, blnNext As Boolean, ByVal strAssistantCols As String, varAssistantText() As Variant)
'功能：重新排列大文本列在每一行的数据
'说明：对于修改非大文段或日期时间列的分组数据时才调用(先调用GetGroupAssistant方法在调用此方法)
    Dim lngCol As Long, lngRow As Long, lngStartRow As Long, varCol
    Dim lngOldRow As Long, lngOldCol As Long, intType As Integer, blnTrue As Boolean
    Dim strText As String, blnOldNoMove As Boolean, blnOldNext As Boolean
    
    lngOldRow = VsfData.ROW
    lngOldCol = VsfData.COL
    intType = mintType
    blnOldNoMove = blnNoMove
    blnOldNext = blnNext
    
    '获取编辑当前行的起始行行
    If VsfData.TextMatrix(VsfData.ROW, mlngRowCount) <> "1|1" Then
        lngRow = GetStartRow(VsfData.ROW)
    Else
        lngRow = VsfData.ROW
    End If
    '获取分组数据的第一行
    lngStartRow = lngRow - Val(VsfData.TextMatrix(lngRow, mlngDemo)) + 1
    If Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) <> 1 Then
        For lngStartRow = lngRow To VsfData.FixedRows Step -1
            If Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) = 1 Then
                lngRow = lngStartRow
                Exit For
            End If
        Next lngStartRow
        If lngRow < VsfData.FixedRows Then Exit Sub
        lngStartRow = lngRow
    End If
    
    If Left(strAssistantCols, 1) = "," Then strAssistantCols = Mid(strAssistantCols, 2)
    varCol = Split(strAssistantCols, ",")
    For lngCol = 0 To UBound(varAssistantText)
        strText = CStr(varAssistantText(lngCol))
        mintType = -1: mblnShow = False
        VsfData.ROW = lngStartRow
        VsfData.COL = Val(varCol(lngCol))
        mblnEditAssistant = True
        blnTrue = True
        mintType = 0
        Call MoveNextCell(False, True, strText)
        mintType = -1
    Next lngCol
    
    '恢复列
    If blnTrue = True Then
        VsfData.ROW = lngOldRow
        VsfData.COL = lngOldCol
        mintType = intType
    End If
    mblnEditAssistant = False
    mblnShow = True
    
    blnNoMove = blnOldNoMove
    blnNext = blnOldNext
End Sub

Private Sub GetGroupAssistant(strAssistantCols As String, varAssistantText() As Variant)
'功能：获取大文本段信息
'说明：对于修改非大文段或日期时间列的分组数据时才调用
    Dim lngRow As Long, lngCol As Long, lngOrder As Long, intGroupFirstRows As Integer, lngCount As Long
    Dim lngStartRow As Long
    Dim strText As String
    
    strAssistantCols = ""
    varAssistantText = Array()
    
    '获取编辑当前行的起始行行
    If VsfData.TextMatrix(VsfData.ROW, mlngRowCount) <> "1|1" Then
        lngRow = GetStartRow(VsfData.ROW)
    Else
        lngRow = VsfData.ROW
    End If
    '获取分组数据的第一行
    lngStartRow = lngRow - Val(VsfData.TextMatrix(lngRow, mlngDemo)) + 1
    If Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) <> 1 Then
        For lngStartRow = lngRow To VsfData.FixedRows Step -1
            If Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) = 1 Then
                lngRow = lngStartRow
                Exit For
            End If
        Next lngStartRow
        If lngStartRow < VsfData.FixedRows Then Exit Sub
        lngStartRow = lngRow
    End If
    
    For lngCol = mlngTime + 1 To mlngNoEditor - 1
        '寻找大文本列
        mrsSelItems.Filter = "列=" & lngCol - cHideCols
        If mrsSelItems.RecordCount > 0 Then
            lngOrder = Val(mrsSelItems!项目序号)
            mrsItems.Filter = "项目序号=" & lngOrder
            If mrsItems.RecordCount = 0 Then
                mrsItems.Filter = 0
                GoTo ErrNext
            End If
            mblnEditAssistant = (mrsItems!项目类型 = 1 And mrsItems!项目长度 > 100) And Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) <= 1
            If Not mblnEditAssistant Then GoTo ErrNext
                
            If InStr(1, VsfData.TextMatrix(lngStartRow, mlngRowCount), "|") = 0 Then VsfData.TextMatrix(lngStartRow, mlngRowCount) = "1|1"
            intGroupFirstRows = Val(Split(VsfData.TextMatrix(lngStartRow, mlngRowCount), "|")(0))
            '为分组行时，选择数据起始行，编辑内容显示所有大文本行
            strText = ""
            If Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) = 1 Then
                For lngRow = 0 To intGroupFirstRows - 1
                    strText = strText & IIf(strText = "", "", vbCrLf) & Replace(Replace(Replace(VsfData.TextMatrix(lngRow + lngStartRow, lngCol), Chr(13), ""), Chr(10), ""), Chr(1), "")
                Next lngRow
                lngCount = lngStartRow + intGroupFirstRows - 1
                For lngRow = lngStartRow + intGroupFirstRows To VsfData.Rows - 1
                    If VsfData.RowHidden(lngRow) = False Then
                        '分组中每一个分组子行数据可能占用多行，只有在第一行保存了分组索引。分组总数据行数=每一个分组子行数+每一个分组子行的行数
                        If lngRow > lngCount Then
                            If Val(VsfData.TextMatrix(lngRow, mlngDemo)) <= 1 Then Exit For  '不分组或遇新分组就退出
                            If VsfData.TextMatrix(lngRow, mlngRowCount) = "" Then VsfData.TextMatrix(lngRow, mlngRowCount) = "1|1"
                            lngCount = Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(0)) + lngRow - 1
                        End If
                        strText = strText & IIf(strText = "", "", vbCrLf) & Replace(Replace(Replace(VsfData.TextMatrix(lngRow, lngCol), Chr(13), ""), Chr(10), ""), Chr(1), "")
                    Else
                        lngCount = lngCount + 1
                    End If
                Next lngRow
                
                If strText = "" Then GoTo ErrNext
                strAssistantCols = strAssistantCols & "," & lngCol
                ReDim Preserve varAssistantText(UBound(varAssistantText) + 1)
                varAssistantText(UBound(varAssistantText)) = strText
            End If
ErrNext:
        End If
    Next lngCol
End Sub

Private Sub ReSetGroupDemo(ByVal lngRow As Long)
'功能：设置分组行的行号和记录集信息
'在修改分组行数据时，如果包含大文本切文本内容不为空通过GetGroupAssistant和ReSetGroupAssistant完成设置，如果没有则调用此函数完成设置
    Dim strDate As String, strTime As String, strYear As String
    Dim intNULL As Integer, lngStartRow As Long, lngRowCount As Long, blnNULL As Boolean
    Dim lngCurRow As Long, lngCol As Long, lngRows As Long, lngCols As Long
    Dim strKey As String, strField As String, strValue As String
    Dim intGroupFirstRows As Integer
    Dim varAssistant() As Variant, strAssistantCols As String
    
    If Val(VsfData.TextMatrix(lngRow, mlngRowCount)) > 1 Then
        lngStartRow = GetStartRow(lngRow)
    Else
        lngStartRow = lngRow
    End If
    '确定分组起始行
    lngRow = lngStartRow
    lngStartRow = lngRow - Val(VsfData.TextMatrix(lngRow, mlngDemo)) + 1
    If Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) <> 1 Then
        For lngStartRow = lngRow To VsfData.FixedRows Step -1
            If Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) = 1 Then
                lngRow = lngStartRow
                Exit For
            End If
        Next lngStartRow
        lngStartRow = lngRow
    End If
    '重新组织分组序号
    VsfData.TextMatrix(lngStartRow, mlngDemo) = 1
    intGroupFirstRows = Val(Split(VsfData.TextMatrix(lngStartRow, mlngRowCount), "|")(0))
    lngCurRow = lngStartRow
    For lngRow = lngStartRow + intGroupFirstRows To VsfData.Rows - 1
        If lngRow = lngCurRow + intGroupFirstRows Then
            If Val(VsfData.TextMatrix(lngRow, mlngDemo)) <= 1 Then
                Exit For
            Else
                VsfData.TextMatrix(lngRow, mlngDemo) = lngRow - Val(lngStartRow) + 1
            End If
            If InStr(1, VsfData.TextMatrix(lngRow, mlngRowCount), "|") = 0 Then VsfData.TextMatrix(lngRow, mlngRowCount) = "1|1"
            intGroupFirstRows = Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(0))
            lngCurRow = lngRow
        End If
    Next

    '从起始行开始处理分组数据记录集
    intGroupFirstRows = 0
    lngCurRow = lngStartRow
    For lngRow = lngStartRow To VsfData.Rows - 1
        If lngRow = lngCurRow + intGroupFirstRows Then
            If Val(VsfData.TextMatrix(lngRow, mlngDemo)) <= 1 And intGroupFirstRows > 0 Then Exit For
            If InStr(1, VsfData.TextMatrix(lngRow, mlngRowCount), "|") = 0 Then VsfData.TextMatrix(lngRow, mlngRowCount) = "1|1"
            intGroupFirstRows = Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(0))
            lngCurRow = lngRow
            strYear = ""
            If CheckGroupDate(lngRow) = True Then
                '保存后的修改才进入此流程，取该条记录的实际时间
                If mblnDateAd Then
                    strYear = Format(VsfData.TextMatrix(lngRow, mlngActTime), "YYYY")
                    strDate = Format(VsfData.TextMatrix(lngRow, mlngActTime), "DD") & "/" & Format(VsfData.TextMatrix(lngRow, mlngActTime), "MM")
                Else
                    strDate = Mid(VsfData.TextMatrix(lngRow, mlngActTime), 1, 10)
                End If
                strTime = Mid(VsfData.TextMatrix(lngRow, mlngActTime), 12, 5)
            Else
                '新增时进入此流程
                strDate = VsfData.TextMatrix(lngRow - Val(VsfData.TextMatrix(lngRow, mlngDemo)) + 1, mlngDate)
                strTime = VsfData.TextMatrix(lngRow - Val(VsfData.TextMatrix(lngRow, mlngDemo)) + 1, mlngTime)
                If mblnDateAd Then strYear = VsfData.TextMatrix(lngRow - Val(VsfData.TextMatrix(lngRow, mlngDemo)) + 1, mlngYear)
            End If
            
            '1\日期
            strField = "ID|页号|行号|列号|记录ID|数据|部位|删除"
            If mlngDate <> -1 Then
                strKey = mint页码 & "," & lngRow & "," & mlngDate
                strValue = strKey & "|" & mint页码 & "|" & lngRow & "|" & mlngDate & "|" & _
                    Val(VsfData.TextMatrix(lngRow, mlngRecord)) & "|" & strDate & "|" & strYear & "|" & IIf(VsfData.RowHidden(lngRow) = True, 1, 0)
                Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
            End If
            '2\时间
            strKey = mint页码 & "," & lngRow & "," & mlngTime
            strValue = strKey & "|" & mint页码 & "|" & lngRow & "|" & mlngTime & "|" & _
                Val(VsfData.TextMatrix(lngRow, mlngRecord)) & "|" & strTime & "|" & _
                VsfData.TextMatrix(lngRow, mlngDemo) & "|" & IIf(VsfData.RowHidden(lngRow) = True, 1, 0)
            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
        End If
    Next lngRow
End Sub

Private Function ISEditAssistant(ByVal lngCol As Long) As Boolean
'是否编辑的是大文本项目
    Dim blnTrue As Boolean, lngOrder As Long
    
    mrsSelItems.Filter = "列=" & lngCol - cHideCols
    If mrsSelItems.RecordCount > 0 Then
        lngOrder = Val(mrsSelItems!项目序号)
        mrsItems.Filter = "项目序号=" & lngOrder
        If mrsItems.RecordCount = 0 Then
            mrsItems.Filter = 0
            Exit Function
        End If
        blnTrue = (mrsItems!项目类型 = 1 And mrsItems!项目长度 > 100)
    End If
    ISEditAssistant = blnTrue
End Function

Private Sub AppendGroup(ByVal lngStartRow As Long)
    Dim lngDemo As Long, lngStart As Long, lngRows As Long
    Dim blnGroup As Boolean
    '追加分组行(只能在单数据行后追加分组行)
    If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then VsfData.TextMatrix(VsfData.ROW, mlngRowCount) = "1|1"
    lngRows = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0))
    '先检查当前行是否为分组行
    blnGroup = (VsfData.TextMatrix(lngStartRow, mlngDemo) <> "")
    If Not blnGroup Then
        lngDemo = 1
        VsfData.TextMatrix(lngStartRow, mlngDemo) = 1
    Else
        lngDemo = VsfData.TextMatrix(lngStartRow, mlngDemo)
    End If
    VsfData.TextMatrix(lngStartRow + lngRows, mlngDemo) = lngDemo + lngRows
    VsfData.ROW = lngStartRow + lngRows
    lngStart = VsfData.ROW - VsfData.TextMatrix(lngStartRow + lngRows, mlngDemo) + 1
    'mstrGroupRow = lngStart
End Sub

Private Function GetStartRow(ByVal lngRow As Long) As Long
    Dim lngStart As Long
    Dim lngCurRows As Long, lngRows As Long
    '提取多行数据起始行,超出本页则返回0
    '如果本页未显示全,则说明超出本页,也返回0
    '不允许在连续的数据行中插入新行
    If VsfData.TextMatrix(lngRow, mlngRowCount) = "" Then VsfData.TextMatrix(lngRow, mlngRowCount) = "1|1"
    lngRows = Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(0))    '总行数
    If InStr(1, VsfData.TextMatrix(lngRow, mlngRowCount), "|") = 0 Then
        lngCurRows = 1
    Else
        lngCurRows = Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(1)) '当前行
    End If
    If lngCurRows = 1 Then
        GetStartRow = lngRow
        Exit Function
    End If
    
    '寻找起始行
    For lngRow = lngRow To 3 Step -1
        If FormatValue(VsfData.TextMatrix(lngRow, mlngRowCount)) = lngRows & "|1" Then
            lngStart = lngRow
            Exit For
        End If
    Next
    
    GetStartRow = lngStart
End Function

Private Function GetMutilData(ByVal lngRow As Long, ByVal lngCol As Long, dblTop As Long, dblHeight As Long) As String
    Dim lngCurRow As Long
    Dim lngCount As Long
    Dim lngStart As Long    '起始行
    Dim lngRecordId As Long
    Dim strReturn As String
    Dim blnAdjust As Boolean
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    '返回第一行的坐标
    '不分行直接取，分行时检查如果当页显示全就拼接，否则从库中读取
    
    If VsfData.TextMatrix(lngRow, mlngRowCount) = "" Then
        GetMutilData = VsfData.TextMatrix(lngRow, lngCol)
        Exit Function
    End If
    lngRecordId = Val(VsfData.TextMatrix(lngRow, mlngRecord))
    lngCount = Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(0))
    lngCurRow = Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(1))
    
    If lngCount > 1 Then
        lngStart = GetStartRow(lngRow)
    Else
        lngStart = lngRow
    End If
    If lngRecordId <> 0 And (lngStart = 0 Or lngStart + lngCount > VsfData.Rows) Then   '页有效行=固定数据行+表头
        '从数据库中提取
        Call SQLCombination(lngRecordId)
        gstrSQL = mstrSQL
        Call SQLDIY(gstrSQL)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取护理数据", mlng文件ID, mlng病人ID, mlng主页ID, mint婴儿, mint页码, lngRecordId)
        strReturn = NVL(rsTemp.Fields(lngCol).Value)
        If lngStart = 0 Then lngStart = 3       '如果未找到启始行则设定为第1行
        blnAdjust = True
    Else
        For lngRow = lngStart To lngStart + lngCount - 1
            If Val(VsfData.TextMatrix(lngRow, mlngCollectType)) < 0 And lngRow = lngStart Then
                strReturn = strReturn & Replace(Replace(Replace(VsfData.TextMatrix(lngRow, lngCol), Chr(13), ""), Chr(10), ""), Chr(1), "") & IIf(lngRow = lngStart + lngCount - 1, "", vbCrLf)
            Else
                strReturn = strReturn & Replace(Replace(Replace(VsfData.TextMatrix(lngRow, lngCol), Chr(13), ""), Chr(10), ""), Chr(1), "") & IIf(lngRow = lngStart + lngCount - 1, "", vbCrLf)
            End If
        Next
    End If
    
'    '校正行高(有可能实际内容占5行而当前页面只显示了3行,若以3行显示数据怕显不全,所以还是以原来的行高显示数据,以下代码屏蔽)
'    If blnAdjust Then
'        If lngStart = 3 Then
'            lngCurRow = Val(Split(VsfData.TextMatrix(lngStart, mlngRowCount), "|")(1))
'            lngCount = lngCount - lngCurRow + 1
'        Else
'            lngCount = mlngPageRows +mlngOverrunRows + VsfData.FixedRows - lngStart
'        End If
'    End If
    '取行高
    VsfData.ROW = lngStart
    dblHeight = lngCount * VsfData.RowHeightMin + 20
    dblTop = VsfData.Top + VsfData.CellTop
    
    GetMutilData = strReturn
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ShowInput(Optional ByVal intCol As Integer = -1, Optional ByVal strCellData As String = "", Optional ByVal blnAnalyse As Boolean = False) As String
    Dim arrData, arrValue
    Dim lngOrder As Long
    Dim i As Integer, j As Integer, intPos As Integer, intIndex As Integer
    Dim strFormat As String, strText As String, strValue As String  '格式串,数据串,数值串
    Dim strOrders As String, strTypes As String, strBounds As String, strLen As String, strName As String, strState As String
    Const txtHeight = 300
    On Error GoTo ErrHand
    
    '病历文件构造管理模块需要处理:
    '1、一列绑定一个项目的不用管
    '2、一列绑定两个项目的，血压必须成对，要么都是录入，要么都是选择，不允许交叉出现，也不允许出现单选、复选
    '3、一列绑定多个项目的，只能是录入项目
    '由于以上条件限制，只取第一个项目的性质即可
    
    '如果是保存处调用则做如下处理
    If intCol = -1 Then intCol = VsfData.COL
    If blnAnalyse Then
        strText = Replace(Replace(Replace(strCellData, Chr(10), ""), Chr(13), ""), Chr(1), "")
    Else
        '取当前单元格的属性
        CellRect.Left = VsfData.CellLeft + VsfData.Left
        CellRect.Top = VsfData.CellTop + VsfData.Top
        CellRect.Bottom = VsfData.CellHeight + 20
        CellRect.Right = VsfData.CellWidth + 20
        strText = GetMutilData(VsfData.ROW, intCol, CellRect.Top, CellRect.Bottom)
    End If
    strText = Replace(Replace(Replace(strText, Chr(10), ""), Chr(13), ""), Chr(1), "")
    mstrData = strText
    If mblnDateAd And mlngYear = intCol Then
        mintType = 8
        strValue = strText
    Else
        mintType = 0
    End If
    intIndex = 0
    
    '取当前列的绑定项目
    intPos = 1
    mrsSelItems.Filter = "列=" & intCol - cHideCols
    Do While Not mrsSelItems.EOF
        lngOrder = mrsSelItems!项目序号
        If lngOrder = 0 Then
            strLen = 0
            strValue = strText
            Exit Do
        End If
        
        '项目表示:2单选;3-多选;4-汇总;5-选择
        '项目值域:项目表示为0-表示最小值;最大值;项目表示为2,3-表示项目A;项目B,前有勾的表示缺省项
        strFormat = NVL(mrsSelItems!格式)
        strOrders = strOrders & "," & lngOrder
        If lngOrder <> 0 Then
            mrsItems.Filter = "项目序号=" & lngOrder
            strName = strName & "," & GetActivePart(intCol, intIndex) & UCase(mrsItems!项目名称)
            strLen = strLen & "," & mrsItems!项目长度 & ";" & NVL(mrsItems!项目小数)
            strState = strState & "," & mrsItems!项目类型
            strTypes = strTypes & "," & mrsItems!项目表示
            strBounds = strBounds & "," & mrsItems!项目值域
            strValue = strValue & "'" & SubstrVal(strText, strFormat, GetActivePart(intCol, intIndex) & UCase(mrsItems!项目名称), intPos)
            
            Select Case mrsItems!项目表示
            Case 0  '文本录入项
                If mrsSelItems.RecordCount = 2 Then
                    If InStr(1, strState & ",", ",1,") = 0 Then
                        mintType = 4
                    Else
                        mintType = 6
                    End If
                ElseIf mrsSelItems.RecordCount > 2 Then
                    mintType = 6
                End If
            Case 2  '单选
                If mrsSelItems.RecordCount = 1 Then
                    mintType = 1
                ElseIf mrsSelItems.RecordCount = 2 Then
                    mintType = 7
                End If
            Case 3  '多选
                mintType = 2
            Case 4  '汇总
                If mrsSelItems.RecordCount = 2 Then
                    mintType = 4
                ElseIf mrsSelItems.RecordCount > 2 Then
                    mintType = 6
                End If
            Case 5  '选择
                If mrsSelItems.RecordCount = 1 Then
                    mintType = 3
                Else
                    mintType = 5
                End If
            End Select
        Else
            strState = strState & ","
            strTypes = strTypes & ","
            strBounds = strBounds & ","
            strLen = strLen & ","
            strName = strName & ","
        End If
        
        intIndex = intIndex + 1
        mrsSelItems.MoveNext
    Loop
    If strOrders <> "" Then
        strOrders = Mid(strOrders, 2)
        strName = Mid(strName, 2)
        strLen = Mid(strLen, 2)
        strState = Mid(strState, 2)
        strTypes = Mid(strTypes, 2)
        strBounds = Mid(strBounds, 2)
        strValue = Mid(strValue, 2)
    End If
    mrsSelItems.Filter = 0
    mrsItems.Filter = 0
    strValue = Replace(Replace(Replace(strValue, Chr(10), ""), Chr(13), ""), Chr(1), "")
    If blnAnalyse Then
        ShowInput = strOrders & "||" & strValue
        Exit Function
    End If
    
    '针对4进行校对,如果表头文本不含/则处理为6
    If mintType = 4 Then
        If Not IsDiagonal(intCol) Then
            mintType = 6
        End If
    End If
    
    '判断当前列的性质
    'mintType:0=文本框录入;1=单选;2=多选;3=选择;4-血压或一列绑定了两个项目,其格式类似血压的输入项目;5=一列绑定了两个项目且均是选择项目;
    '6=一列绑定2个及以上项目,手工录入,7=一列绑定了两个单选项目
    arrValue = Split(strValue & "'", "'")
    Select Case mintType
    Case 0, 3
        With picInput
            .Left = CellRect.Left
            .Top = CellRect.Top
            .Width = CellRect.Right
            .Height = CellRect.Bottom
            If .Height + .Top + picMain.Top > ScaleHeight Then
                If ScaleHeight - picMain.Top - .Height < 0 Then
                    .Height = ScaleHeight - picMain.Top
                Else
                    .Top = ScaleHeight - picMain.Top - .Height
                End If
            End If
            
            If .Top < 0 Then .Top = 0
            .Visible = True
            .ZOrder 0
        End With
        If mintType = 0 Then
            txtInput.Visible = True
            If Val(strLen) <> 0 And Val(strOrders) <> 10 Then
                txtInput.MaxLength = Val(Split(strLen, ";")(0)) + IIf(Val(Split(strLen, ";")(1)) = 0, 0, 1) '小数位数要加上小数点
            Else
                txtInput.MaxLength = 0
            End If
            txtInput.Tag = lngOrder
        Else
            txtInput.Visible = False
        End If
        With txtInput
            .Top = 0
            .Text = strValue
            .Width = CellRect.Right
            .Height = CellRect.Bottom
            .FontName = VsfData.FontName
            .FontSize = VsfData.FontSize
            .Width = .Width - (180 + IIf(mblnBlowup, 180 * 1 / 3, 0)) / 2 + 3 '宋体9号时减去90,字体越大扣除的边距越小,以保证文本框分行与实际一致
        End With
        With lblInput
            .FontName = VsfData.FontName
            .FontSize = VsfData.FontSize
            .Height = CellRect.Bottom
            .Width = CellRect.Right
            .Top = 50
            .Tag = lngOrder
            .Caption = strValue
            .Visible = (mintType = 3)
        End With
        
        '如果是日期或时间列，设定固定值
        If mintType = 0 And txtInput.Text = "" Then
            If intCol = mlngDate Then
                If mblnDateAd Then
                    txtInput.Text = Format(zlDatabase.Currentdate, "d-M")
                    txtInput.Text = Replace(txtInput.Text, "-", "/")
                Else
                    txtInput.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
                End If
            ElseIf intCol = mlngTime Then
                txtInput.Text = Format(zlDatabase.Currentdate, "HH:mm")
            End If
        End If
    Case 1, 2
        '56439:刘鹏飞,2012-11-30,单选项目如果未设置缺省想，默认定位到清除选择，以前的方式是定位到
        '实际数据项，对于有些项目不需要录入，就会要操作员手工选择到清除选择项，在操作上很麻烦。
        '加载数据
        lstSelect(mintType - 1).Clear
        If mintType = 1 Then lstSelect(mintType - 1).AddItem "清除选择"
        If strBounds = "" Then strBounds = ";"
        arrData = Split(strBounds, ";")
        j = UBound(arrData)
        For i = 0 To j
            If arrData(i) <> "" Then
                If Mid(arrData(i), 1, 1) = "√" Then
                    lstSelect(mintType - 1).AddItem lstSelect(mintType - 1).NewIndex + 1 & "-" & Mid(arrData(i), 2)
                    If strText = "" Then lstSelect(mintType - 1).ListIndex = lstSelect(mintType - 1).NewIndex
                Else
                    lstSelect(mintType - 1).AddItem lstSelect(mintType - 1).NewIndex + 1 & "-" & arrData(i)
                End If
            End If
        Next
        
        '多选且已录入数据的情况下
        If strValue <> "" Then
            strValue = Replace(strValue, vbCrLf, "")
            txtLst.Text = strValue
            PicLst.Tag = "1"
            j = lstSelect(mintType - 1).ListCount - 1
            For i = 0 To j
                '单选的第一个项目是清除选择，需要跳过此项，多选项目则直接进入
                If Not (mintType = 1 And i = 0) Then
                    If InStr(1, "," & strValue & ",", "," & Mid(lstSelect(mintType - 1).List(i), InStr(1, lstSelect(mintType - 1).List(i), "-") + 1) & ",") <> 0 Then
                        lstSelect(mintType - 1).Selected(i) = True
                        txtLst.Text = ""
                        PicLst.Tag = "0"
                    End If
                End If
            Next
        Else
            txtLst.Text = ""
            PicLst.Tag = "0"
        End If
        
        '控件显示
        '51134,刘鹏飞,2012-07-11,单选提供文本录入
        PicLst.FontName = VsfData.FontName
        PicLst.FontSize = VsfData.FontSize
        If mintType = 1 Then
        
            With PicLst
                .Left = CellRect.Left
                .Top = CellRect.Top
                .Width = LenB(StrConv(lstSelect(mintType - 1).List(lstSelect(mintType - 1).ListCount \ 2), vbFromUnicode)) * 120 + 500    '以中间项的长度为依据
                If .Width < CellRect.Right Then .Width = CellRect.Right
            End With
            
            With lbllst(0)
                .Left = 20
                .Top = 20
                If .Width > PicLst.Width Then
                    PicLst.Width = .Width + PicLst.TextWidth("刘")
                End If
                .FontName = VsfData.FontName
                .FontSize = VsfData.FontSize
                .Visible = True
            End With
            
            With txtLst
                .Top = lbllst(0).Top + lbllst(0).Height + 20
                .Left = -10
                .Width = PicLst.Width
                
                If .Text <> "" Then
                    txtLength.Width = .Width
                    txtLength.Text = Replace(Replace(Replace(.Text, Chr(10), ""), Chr(13), ""), Chr(1), "")
                    txtLength.FontName = VsfData.CellFontName
                    txtLength.FontSize = VsfData.CellFontSize
                    txtLength.FontBold = VsfData.CellFontBold
                    txtLength.FontItalic = VsfData.CellFontItalic
                    arrData = GetData(txtLength.Text)
                    .Text = Join(arrData, "")
                    If PicLst.TextHeight("刘") * (UBound(arrData) + 1) + PicLst.TextHeight("刘") \ 3 < VsfData.CellHeight + 20 Then
                        .Height = VsfData.CellHeight + 20
                    Else
                        .Height = PicLst.TextHeight("刘") * (UBound(arrData) + 1) + PicLst.TextHeight("刘") \ 3
                    End If
                Else
                    .Height = VsfData.CellHeight + 20
                End If
                .FontName = VsfData.FontName
                .FontSize = VsfData.FontSize
                .Tag = VsfData.CellHeight + 20 '最小高度
                If strLen <> "" Then
                    .MaxLength = Val(Split(strLen, ";")(0)) + IIf(Val(Split(strLen, ";")(1)) = 0, 0, 1) '小数位数要加上小数点
                End If
                .Visible = True
            End With
            
            With lbllst(1)
                .Left = 20
                .Top = txtLst.Top + txtLst.Height + 20
                .FontName = VsfData.FontName
                .FontSize = VsfData.FontSize
                .Visible = True
            End With
            '56047:刘鹏飞,2012-11-22,修改PicLst的坐标
            With PicLst
                'list控件的高度最小是240，后面是字体高度比例增长：故高度计算=.ListCount * PicLst.TextHeight("刘") + PicLst.TextHeight("刘") \ 3
                .Height = lbllst(1).Top + lbllst(1).Height + 20 + lstSelect(mintType - 1).ListCount * PicLst.TextHeight("刘") + PicLst.TextHeight("刘") \ 3
                If .Height < CellRect.Bottom Then .Height = CellRect.Bottom
                If .Height + .Top + picMain.Top > ScaleHeight Then
                    If ScaleHeight - picMain.Top - .Height < 0 Then
                        .Top = 10
                        .Height = ScaleHeight - picMain.Top - 10
                    Else
                        .Top = ScaleHeight - picMain.Top - .Height
                    End If
                End If
                .Visible = True
                .ZOrder 0
            End With
            
            With lstSelect(mintType - 1)
                .Top = lbllst(1).Top + lbllst(1).Height + 20
                .Left = -10
                .FontName = VsfData.FontName
                .FontSize = VsfData.FontSize
                .Width = PicLst.Width
                .Height = IIf(PicLst.Height - .Top < PicLst.TextHeight("刘") + PicLst.TextHeight("刘") \ 3, PicLst.TextHeight("刘") + PicLst.TextHeight("刘") \ 3, PicLst.Height - .Top)
                .Tag = lngOrder
                .Visible = True
                
                If .Top + .Height <> PicLst.Height Then
                    PicLst.Height = .Top + .Height
                End If
            End With
        Else
            '56047:刘鹏飞,2012-11-22,修改lstSelect的坐标
            '显示
            With lstSelect(mintType - 1)
                .Left = CellRect.Left
                .Top = CellRect.Top
                .FontName = VsfData.FontName
                .FontSize = VsfData.FontSize
                .Height = .ListCount * (PicLst.TextHeight("刘")) + PicLst.TextHeight("刘") \ 3
                If .Height < CellRect.Bottom Then .Height = CellRect.Bottom
                .Width = LenB(StrConv(.List(.ListCount \ 2), vbFromUnicode)) * 120 + 500    '以中间项的长度为依据
                If .Width < CellRect.Right Then .Width = CellRect.Right
                If .Height + .Top + picMain.Top > ScaleHeight Then
                    If ScaleHeight - picMain.Top - .Height < 0 Then
                        .Top = 10
                        .Height = IIf(ScaleHeight - picMain.Top - 10 < PicLst.TextHeight("刘") + PicLst.TextHeight("刘") \ 3, PicLst.TextHeight("刘") + PicLst.TextHeight("刘") \ 3, ScaleHeight - picMain.Top - 10)
                    Else
                        .Top = ScaleHeight - picMain.Top - .Height
                    End If
                End If
                .Tag = lngOrder
                .Visible = True
                .ZOrder 0
            End With
        End If
        
    Case 4, 5
        With picDouble
            .Left = CellRect.Left
            .Top = CellRect.Top
            .Height = CellRect.Bottom
            If .Height < 280 Then .Height = 280
            .Width = CellRect.Right
            If .Width < 820 Then .Width = 820
            If .Height + .Top + picMain.Top > ScaleHeight Then
                .Top = ScaleHeight - picMain.Top - .Height
            End If
            .Visible = True
            .ZOrder 0
        End With
        lblSplit.FontName = VsfData.FontName
        lblSplit.FontSize = VsfData.FontSize
        lblSplit.Left = (picDouble.Width - lblSplit.Width) / 2
        If mblnBlowup Then
            lblSplit.Width = 150
        Else
            lblSplit.Width = 105
        End If
        
        With txtUpInput
            .Text = arrValue(0)
            .FontName = VsfData.FontName
            .FontSize = VsfData.FontSize
            .Width = (picDouble.Width - lblSplit.Width) * 0.4
            .ZOrder IIf(mintType = 4, 0, 1)
            .Locked = Not (mintType = 4)
            .Tag = Split(strOrders, ",")(0)
        End With
        With picUpInput
            .Left = txtUpInput.Left
            .Width = txtUpInput.Width
            .Height = CellRect.Bottom
            .ZOrder IIf(mintType = 5, 0, 1)
            .Tag = Split(strOrders, ",")(0)
        End With
        With lblUpInput
            .Alignment = 2
            .Caption = arrValue(0)
            .Left = 0
            .Top = 50
            .FontName = VsfData.FontName
            .FontSize = VsfData.FontSize
            .Width = txtUpInput.Width
            .Height = CellRect.Bottom
            .Tag = Split(strOrders, ",")(0)
        End With
        With txtDnInput
            .Text = arrValue(1)
            .FontName = VsfData.FontName
            .FontSize = VsfData.FontSize
            .Left = lblSplit.Left + lblSplit.Width
            .Width = picDouble.Width - .Left
            .ZOrder IIf(mintType = 4, 0, 1)
            .Locked = Not (mintType = 4)
            .Tag = Split(strOrders, ",")(1)
        End With
        With picDnInput
            .Left = txtDnInput.Left
            .Height = CellRect.Bottom
            .Width = txtDnInput.Width
            .ZOrder IIf(mintType = 5, 0, 1)
            .Tag = Split(strOrders, ",")(1)
        End With
        With lblDnInput
            .Alignment = 2
            .Caption = arrValue(1)
            .Left = 0
            .Top = 50
            .Height = CellRect.Bottom
            .Width = txtDnInput.Width
            .FontName = VsfData.FontName
            .FontSize = VsfData.FontSize
            .Tag = Split(strOrders, ",")(1)
        End With
        
        If mintType = 4 Then
            If strLen <> "" Then txtUpInput.MaxLength = Val(Split(Split(strLen, ",")(0), ";")(0)) + IIf(Val(Split(Split(strLen, ",")(0), ";")(1)) = 0, 0, 1) '小数位数要加上小数点
            If strLen <> "" Then txtDnInput.MaxLength = Val(Split(Split(strLen, ",")(1), ";")(0)) + IIf(Val(Split(Split(strLen, ",")(1), ";")(1)) = 0, 0, 1) '小数位数要加上小数点
        End If
    Case 6
        '先删除以前的控件
        j = txt.Count - 1
        For i = 1 To j
            Unload lbl(i)
            Unload txt(i)
        Next
        '设定坐标
        With picMutilInput
            .Left = CellRect.Left
            .Top = CellRect.Top
            .Width = IIf(CellRect.Right < 1600, 1600, CellRect.Right)
        End With
        '对缺省控件赋值
        arrData = Split(strOrders, ",")
        j = UBound(arrData)
        lbl(0).Top = 130
        lbl(0).Caption = Split(strName, ",")(0)
        lbl(0).FontName = VsfData.FontName
        lbl(0).FontSize = VsfData.FontSize
        txt(0).Tag = arrData(0)
        txt(0).FontName = VsfData.FontName
        txt(0).FontSize = VsfData.FontSize
        txt(0).Width = picMutilInput.Width - txt(0).Left - 100
        txt(0).MaxLength = Val(Split(Split(strLen, ",")(0), ";")(0)) + IIf(Val(Split(Split(strLen, ",")(0), ";")(1)) = 0, 0, 1)  '小数位数要加上小数点
        txt(0).Text = arrValue(0)
        If Not mblnBlowup Then
            txt(0).Height = 225
        End If
        
        '加载控件
        For i = 1 To j
            Load lbl(i)
            With lbl(i)
                .Caption = Split(strName, ",")(i)
                .Left = lbl(0).Left + lbl(0).Width - .Width
                .Top = lbl(i - 1).Top + txtHeight + IIf(mblnBlowup, txtHeight * 1 / 3, 0)
                .Visible = True
            End With
            Load txt(i)
            With txt(i)
                .TabIndex = txt(i - 1).TabIndex + 1
                .Left = txt(0).Left
                .Top = txt(i - 1).Top + txtHeight + IIf(mblnBlowup, txtHeight * 1 / 3, 0)
                .Tag = arrData(i)
                If strLen <> "" Then
                    .MaxLength = Val(Split(Split(strLen, ",")(i), ";")(0)) + IIf(Val(Split(Split(strLen, ",")(i), ";")(1)) = 0, 0, 1) '小数位数要加上小数点
                End If
                .Text = arrValue(i)
                .Visible = True
            End With
        Next
        
        With picMutilInput
            .Height = txt(j).Top + txt(j).Height + 120
            If .Height < CellRect.Bottom Then .Height = CellRect.Bottom
            If .Height + .Top + picMain.Top > ScaleHeight Then
                .Top = ScaleHeight - picMain.Top - .Height
            End If
            If .Top < 0 Then .Top = 0
            .Visible = True
            .ZOrder 0
        End With
    Case 7
        cboChoose(0).Clear
        cboChoose(0).FontName = VsfData.FontName
        cboChoose(0).FontSize = VsfData.FontSize
        cboChoose(0).Tag = Split(strOrders, ",")(0)
        arrData = Split(Split(strBounds, ",")(0), ";")
        j = UBound(arrData)
        For i = 0 To j
            If Mid(arrData(i), 1, 1) = "√" Then
                cboChoose(0).AddItem Mid(arrData(i), 2)
                If strValue = "" Then
                    cboChoose(0).ListIndex = i
                Else
                    If Mid(arrData(i), 2) = Split(strValue, "'")(0) Then
                        cboChoose(0).ListIndex = i
                    End If
                End If
            Else
                cboChoose(0).AddItem arrData(i)
                If strValue <> "" Then
                    If arrData(i) = Split(strValue, "'")(0) Then
                        cboChoose(0).ListIndex = i
                    End If
                End If
            End If
        Next
        
        cboChoose(1).Clear
        cboChoose(1).FontName = VsfData.FontName
        cboChoose(1).FontSize = VsfData.FontSize
        cboChoose(1).Tag = Split(strOrders, ",")(1)
        arrData = Split(Split(strBounds, ",")(1), ";")
        j = UBound(arrData)
        For i = 0 To j
            If Mid(arrData(i), 1, 1) = "√" Then
                cboChoose(1).AddItem Mid(arrData(i), 2)
                If strValue = "" Then
                    cboChoose(1).ListIndex = i
                Else
                    If Mid(arrData(i), 2) = Split(strValue, "'")(1) Then
                        cboChoose(1).ListIndex = i
                    End If
                End If
            Else
                cboChoose(1).AddItem arrData(i)
                If strValue <> "" Then
                    If arrData(i) = Split(strValue, "'")(1) Then
                        cboChoose(1).ListIndex = i
                    End If
                End If
            End If
        Next
        
        With picDoubleChoose
            .Left = CellRect.Left
            .Top = CellRect.Top
            .Height = CellRect.Bottom
            If .Height < 280 Then .Height = 280
            .Width = CellRect.Right
            If .Width < 820 Then .Width = 820
            If .Height + .Top + picMain.Top > ScaleHeight Then
                .Top = ScaleHeight - picMain.Top - .Height
            End If
            .Visible = True
            .ZOrder 0
        End With
        lblSplit.FontName = VsfData.FontName
        lblSplit.FontSize = VsfData.FontSize
        lblSplit.Left = (picDoubleChoose.Width - lblSplit.Width) / 2
        If mblnBlowup Then
            lblSplit.Width = 150
        Else
            lblSplit.Width = 105
        End If
        picChooseRight.Left = lblSplit.Left + 150
        cboChoose(0).SetFocus
    Case 8
        cboYear.Clear
        arrData = Split(mstrYears, "|")
        For j = 0 To UBound(arrData)
            cboYear.AddItem Val(arrData(j))
            If Val(strValue) = Val(arrData(j)) Then
                cboYear.ListIndex = j
            End If
        Next j
        
         With picYear
            .Left = CellRect.Left
            .Top = CellRect.Top
            .Height = CellRect.Bottom
            .Width = CellRect.Right
            If .Height + .Top + picMain.Top > ScaleHeight Then
                .Top = ScaleHeight - picMain.Top - .Height
            End If
            .Visible = True
            .ZOrder 0
        End With
        
        With cboYear
            .Left = -10
            .Top = -10
            .Width = picYear.Width + 300
            .FontName = VsfData.FontName
            .FontSize = VsfData.FontSize
            .Visible = True
        End With
        
        cboYear.SetFocus
        
    End Select
    Exit Function
    
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub CheckFormat(ByVal strNames As String, ByVal strFormat As String)
    '如果格式与血压的方式不同,则将样式处理为6
    
    '去掉前缀后进行对比
    strFormat = Mid(strFormat, InStr(1, strFormat, "["))
    strFormat = Replace(strFormat, "[", "")
    strFormat = Replace(strFormat, "]", "")
    If Not (strFormat Like Split(strNames, ",")(0) & "/}*" Or strFormat Like "{/*" & Split(strNames, ",")(1)) Then
        mintType = 6
    End If
End Sub

Private Function IsDiagonal(ByVal intCol As Integer) As Boolean
    Dim arrCol, arrData
    Dim intDo As Integer, intCount As Integer
    '判断指定列是否设置了列对角线（mstrColWidth的格式：765`11`1`1,765`11`2`1,...，对象属性`对象序号`列对角线）
    
    IsDiagonal = (InStr(1, "," & mstrCatercorner & ",", "," & intCol - (cHideCols + VsfData.FixedCols - 1) & ",") <> 0)
End Function

Private Sub ISAssistant(ByVal lngOrder As Long, ByVal objTXT As TextBox)
    Dim intIndex As Integer, intType As Integer
    Dim objParent As Object
    Dim intRow As Integer, intCount As Integer, i As Integer, intGroupFirstRows As Integer, intHidden As Integer
    Dim strText As String, lngCount As Long
    Dim arrData, lngStartRow As Long
    '根据项目的长度决定是否允许进行词句选择
    mblnEditAssistant = False
    mblnEditText = False
    cmdWord.Visible = mblnEditAssistant
    
    mrsItems.Filter = "项目序号=" & lngOrder
    If mrsItems.RecordCount = 0 Then
        mrsItems.Filter = 0
        Exit Sub
    End If
    intType = mintType
    mblnEditAssistant = (mrsItems!项目类型 = 1 And mrsItems!项目长度 > 100) 'And Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) <= "1"
    mblnEditText = (mrsItems!项目类型 = 1 And NVL(mrsItems!项目表示, 0) = 0)
    If mblnEditText = True And mblnEditAssistant = False Then
        If UCase(objTXT.Name) = "TXTINPUT" Then
            cmdWord.Tag = -1  '表示txtInput
        Else
            cmdWord.Tag = objTXT.Index
        End If
    End If
    mrsItems.Filter = 0
    lngStartRow = VsfData.ROW
    '获取分组数据的第一行
    If Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) > 1 Then
        lngStartRow = VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1
        If Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) <> 1 Then
            For lngStartRow = VsfData.ROW To VsfData.FixedRows Step -1
                If Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) = 1 Then
                    Exit For
                End If
            Next lngStartRow
            If lngStartRow < VsfData.FixedRows Then Exit Sub
        End If
    End If
    
    '如果允许词句选择,显示并定位
    If mblnEditAssistant Then
        mintType = -1
        VsfData.ROW = lngStartRow
        mintType = intType
        
        If UCase(objTXT.Name) = "TXTINPUT" Then
            intIndex = -1 '表示txtInput
            Set objParent = picInput
        Else
            intIndex = objTXT.Index
            Set objParent = picMutilInput
        End If
        With cmdWord
            .Tag = intIndex
            .Top = objParent.Top + objTXT.Top + 25
            .Left = objParent.Left + objTXT.Left + objTXT.Width - .Width + 25
            .Visible = True
            .ZOrder 0
        End With
        strText = ""
        intCount = 0
        intHidden = 0
        If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then VsfData.TextMatrix(VsfData.ROW, mlngRowCount) = "1|1"
        intGroupFirstRows = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0))
        '为分组行时，选择数据起始行，编辑内容显示所有大文本行
        If Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) = 1 Then
            For intRow = 0 To intGroupFirstRows - 1
                strText = strText & IIf(strText = "", "", vbCrLf) & Replace(Replace(Replace(VsfData.TextMatrix(intRow + VsfData.ROW, VsfData.COL), Chr(13), ""), Chr(10), ""), Chr(1), "")
            Next intRow
            lngCount = VsfData.ROW + intGroupFirstRows - 1
            For intRow = VsfData.ROW + intGroupFirstRows To VsfData.Rows - 1
                If VsfData.RowHidden(intRow) = False Then
                    '分组中每一个分组子行数据可能占用多行，只有在第一行保存了分组索引。分组总数据行数=每一个分组子行数+每一个分组子行的行数
                    If intRow > lngCount Then
                        If Val(VsfData.TextMatrix(intRow, mlngDemo)) <= 1 Then
                            'If strText <> "" Then strText = Mid(strText, 1, Len(strText) - 2)
                            Exit For
                        End If
                        lngCount = Val(Split(VsfData.TextMatrix(intRow, mlngRowCount), "|")(0)) + intRow - 1
                    End If
                    strText = strText & IIf(strText = "", "", vbCrLf) & Replace(Replace(Replace(VsfData.TextMatrix(intRow, VsfData.COL), Chr(13), ""), Chr(10), ""), Chr(1), "")
                    'strText = strText & IIf(intRow > VsfData.ROW And strText <> "", vbCrLf, "") & Replace(VsfData.TextMatrix(intRow, VsfData.COL), vbCrLf, "")
                Else
                    lngCount = lngCount + 1
                    intHidden = intHidden + 1
                End If
            Next intRow
            '准备赋值
            With txtLength
                '日期与时间列的宽度不管,为了避免返回多行,强制设置为5000
                .Width = VsfData.CellWidth
                .Text = Replace(Replace(Replace(strText, Chr(10), ""), Chr(13), ""), Chr(1), "")
                .FontName = VsfData.CellFontName
                .FontSize = VsfData.CellFontSize
                .FontBold = VsfData.CellFontBold
                .FontItalic = VsfData.CellFontItalic
            End With
            arrData = GetData(txtLength.Text)
            intCount = UBound(arrData)
            strText = ""
            For i = 0 To intCount
                strText = strText & CStr(arrData(i))
            Next i
            intRow = intRow - VsfData.ROW - intHidden
            picInput.Height = intRow * VsfData.RowHeightMin + 20
            If picInput.Height + picInput.Top + picMain.Top > ScaleHeight Then
                picInput.Top = ScaleHeight - picMain.Top - picInput.Height
            End If
            txtInput.Height = picInput.Height
            txtInput.Text = strText
            mstrData = Replace(Replace(Replace(strText, Chr(10), ""), Chr(13), ""), Chr(1), "")
            Call zlControl.TxtSelAll(txtInput)
            lblInput.Height = picInput.Height
        End If
    End If
End Sub

Private Sub FillPage(Optional ByVal blnLocate As Boolean = True)
    Dim lngRow As Long, lngRows As Long, lngCount As Long, lngData As Long
    '保证每页有效数据行
    
    lngRows = VsfData.Rows - 1
    For lngRow = VsfData.FixedRows To lngRows
        If VsfData.TextMatrix(lngRow, mlngRowCount) <> "" Then
            lngData = lngData + 1
        End If
        If Not VsfData.RowHidden(lngRow) Then
            lngCount = lngCount + 1
        End If
    Next
    
    If lngCount < mlngPageRows + mlngOverrunRows - mlngReduceRow Then
        VsfData.Rows = VsfData.Rows + (mlngPageRows + mlngOverrunRows - mlngReduceRow - lngCount)
'    Else
'        VsfData.Rows = VsfData.FixedRows + (mlngPageRows + mlngOverrunRows)
    End If
    
    If mint页码 <= mint结束页 And mint页码 > 0 Then
        If mblnRestore = False Then
            VsfData.Rows = VsfData.Rows - Val(CStr(mlngLitterRows(mint页码)))
        Else
            '跨页数据修改后(不保存)mlngOverrunRows值会变为0，此处需要做处理保证数据可以正常展示
            '示例：前提是跨页数据显示在当前页，并且第一页数据跨页，如果第二页数据最后一行数据跨页(数据总行数5,跨2行,mlngOverrunRows=2),当修改这条数据的其他内容
            '会重新给行数和实际行数列赋值，就会导致mlngOverrunRows=0.通过切换页后回到本页就会导致VsfData.Rows不正确。
            '算法示例：mlngPageRows=20，第一页跨页行数5行（mlngCurLitterRows(mint页码)=5），第二页跨页行数2行（mlngOverrunRows=2）。
            '1、第一次加载数据第二页总行数应该为20-5+2=17，此时如果修改第二页跨页数据(例如：将体温37给为38)就会导致mlngOverrunRows=0
            '2、切换第一页在回到第二页，第二页的总行数就是20-5-0=15，这样就会导致第二页跨页数据的两行数据无法显示。
            If lngCount > VsfData.Rows - Val(CStr(mlngCurLitterRows(mint页码))) - VsfData.FixedRows Then
                VsfData.Rows = lngCount + VsfData.FixedRows
            Else
                VsfData.Rows = VsfData.Rows - Val(CStr(mlngCurLitterRows(mint页码)))
            End If
        End If
    End If
    
    On Error Resume Next
    If Not blnLocate Then Exit Sub
    If lngData + VsfData.FixedRows < VsfData.Rows - 1 Then
        VsfData.ROW = lngData + VsfData.FixedRows
    Else
        VsfData.ROW = VsfData.FixedRows
    End If
    If Not VsfData.RowIsVisible(VsfData.ROW) Then VsfData.TopRow = VsfData.ROW
    '如果最后一行是空行,则进入编辑状态
    If VsfData.TextMatrix(VsfData.ROW, mlngRowCount) = "" Then
        VsfData.SetFocus
        VsfData.COL = mlngDate
    End If
End Sub

Public Function GetSynItems(ByVal intType As Integer, ByRef intMax As Integer) As String
    Dim arrCols
    Dim strItems As String
    Dim strCols As String
    Dim strNames As String
    Dim lngRecord As Long, lngStartRow As Long, lngEndRow As Long
    Dim intIn As Integer, intOut As Integer, intInMAX As Integer, intOutMax As Integer, intCount As Integer
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    'intType，确定函数返回值，1)返回项目序号;2)返回列号
    'intMAX，返回同步数据列所占用的行高
    '返回同步数据列(一份文件中不可能出现重复的项目,所以,判断时不必检查列号)
    
    lngRecord = Val(VsfData.TextMatrix(VsfData.ROW, mlngRecord))
    If lngRecord = 0 Then Exit Function
    
    gstrSQL = "" & _
        " SELECT  B.项目序号,B.项目名称,A.对象序号 AS 列号" & vbNewLine & _
        " FROM 病历文件结构 A,病人护理明细 B" & vbNewLine & _
        " WHERE A.要素名称=B.项目名称 AND A.父ID=" & vbNewLine & _
        "      (SELECT A.ID FROM 病历文件结构 A,病人护理文件 B " & vbNewLine & _
        "       WHERE B.ID=[2] And A.文件ID=B.格式ID AND A.对象序号=4 AND A.父ID IS NULL)" & vbNewLine & _
        " AND B.数据来源>0 and B.数据来源 <> 3 AND B.记录ID=[1]"
    Call SQLDIY(gstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "返回同步数据列", lngRecord, mlng文件ID)
    If rsTemp.RecordCount = 0 Then Exit Function
    
    '获取同步相关信息
    Do While Not rsTemp.EOF
        If InStr(1, "," & strCols & ",", "," & rsTemp!列号 & ",") = 0 Then strCols = strCols & "," & rsTemp!列号
        strItems = strItems & "," & rsTemp!项目序号
        strNames = strNames & "," & rsTemp!项目名称
        rsTemp.MoveNext
    Loop
    strCols = Mid(strCols, 2)
    strItems = Mid(strItems, 2)
    strNames = Mid(strNames, 2)
    
    '根据列循环检查内容所占行高
    If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then
        lngStartRow = VsfData.ROW
        lngEndRow = VsfData.ROW
        intInMAX = 1
    Else
        lngStartRow = GetStartRow(VsfData.ROW)
        intInMAX = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0))
        lngEndRow = lngStartRow + intInMAX - 1
    End If
    
    intCount = 1    '同步的只能是数字型项目，所以占用行只可能是1行，以下内容不再需要检查
'    '数据占用超过1行才检查
'    If intInMAX > 1 Then
'        arrCols = Split(strCols, ",")
'        intOutMax = UBound(arrCols)
'        For intOut = 0 To intOutMax
'            For intIn = 2 To intInMAX
'                If VsfData.TextMatrix(intIn + lngStartRow - 1, arrCols(intOut) + 1) <> "" Then
'                    If intIn > intCount Then intCount = intIn
'                End If
'            Next
'        Next
'    End If
    
    intMax = intCount
    GetSynItems = IIf(intType = 1, strItems, strCols)
    If strNames <> "" Then
        RaiseEvent AfterRowColChange("日期列,时间列,以及 " & strNames & " 是同步过来的数据，不允许修改或删除！", True, mblnSign, mblnArchive)
    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ISColHaveData() As Boolean
    Dim arrData
    Dim arrCol
    Dim intCol As Integer
    Dim intDo As Integer, intCount As Integer
    Dim intIn As Integer, intMax As Integer
    Dim strCond As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    '从数据库中提取数据，如果当前活动项目列存在数据则不允许调整活动项目设置
    
    '将活动项目加入到查询SQL中，格式：列号;表头名称|项目序号,部位;项目序号,部位||列号;表头名称...
    '绑定多个项目，该列就自动转为对角线列
    arrData = Split(mstrCOLActive, "||")
    intCount = UBound(arrData)
    For intDo = 0 To intCount
        intCol = Split(Split(arrData(intDo), "|")(0), ";")(0)
        If intCol = VsfData.COL - cHideCols - VsfData.FixedCols + 1 Then
            arrCol = Split(Split(arrData(intDo), "|")(1), ";")
            intMax = UBound(arrCol)
            For intIn = 0 To intMax
                strCond = strCond & " OR (项目序号=" & Split(arrCol(intIn), ",")(0)
                If Split(arrCol(intIn), ",")(1) = "" Then
                    strCond = strCond & ")"
                Else
                    strCond = strCond & " AND NVL(体温部位,'TWBW')='" & Split(arrCol(intIn), ",")(1) & "')"
                End If
            Next
            
            Exit For
        End If
    Next
    
    If strCond <> "" Then
        strCond = " AND (" & Mid(strCond, 4) & ")"
        '查询数据库
        gstrSQL = " SELECT  1 FROM 病人护理明细 A,病人护理数据 B,病人护理打印 C" & vbNewLine & _
                  " Where A.记录ID=B.ID And B.汇总类别=0 And B.ID=C.记录ID And C.文件ID=B.文件ID " & vbNewLine & _
                  " And C.文件ID=[1] And (C.结束页号=[2] OR C.开始页号=[2])" & strCond & " AND ROWNUM<2"
        Call SQLDIY(gstrSQL)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "查询数据库当前页面指定活动列是否存在活动项目", mlng文件ID, mint页码)
        ISColHaveData = rsTemp.RecordCount
    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


'######################################################################################################################
'**********************************************************************************************************************
'以下是基础函数或过程

Private Sub txt结束时点_GotFocus()
    Call zlControl.TxtSelAll(txt结束时点)
End Sub

Private Sub txt结束时点_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim blnCancel As Boolean
    If KeyCode = vbKeyReturn Then
        Call txt结束时点_Validate(blnCancel)
        txt小结名称.SetFocus
    End If
End Sub

Private Sub txt结束时点_Validate(Cancel As Boolean)
    Dim strFormat As String
    Dim intDef As Integer   '时差+1
    '检查开始时点,结束时点合法性
    
    strFormat = CheckTxtTime(txt开始时点)
    If strFormat = "" Then Exit Sub
    txt开始时点.Text = strFormat
    strFormat = CheckTxtTime(txt结束时点)
    If strFormat = "" Then Exit Sub
    txt结束时点.Text = strFormat
    
    '更新小结名称
    If txt结束时点.Text > txt开始时点.Text Then
        intDef = Val(txt结束时点.Text) - Val(txt开始时点.Text)
    Else
        intDef = Val(txt结束时点.Text) + 24 - Val(txt开始时点.Text)
    End If
    '如果分钟数是59，则加1小时
    If Split(txt结束时点.Text, ":")(1) = "59" Then intDef = intDef + 1
    '71794:刘鹏飞,2014-05-06,临时小结不足一小时也可以小结
    '对于33(含33)以后版本，此标识只是一个数据合法的判断
    If intDef = 0 Then
        txt小结名称.Tag = 1
        txt小结名称.Text = "不足1小时小结"
    Else
        txt小结名称.Tag = intDef
        txt小结名称.Text = intDef & "小时小结"
    End If
End Sub

Private Sub txt小结名称_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdOk.SetFocus
End Sub

Private Sub txt开始时点_GotFocus()
    Call zlControl.TxtSelAll(txt开始时点)
End Sub

Private Sub txt开始时点_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txt结束时点.SetFocus
End Sub

Private Sub lblDnInput_Click()
    txtDnInput.SetFocus
End Sub

Private Sub lblUpInput_Click()
    txtUpInput.SetFocus
End Sub

Private Sub lstColumnItems_DblClick()
    Call cmdColumn_Click(0)
End Sub

Private Sub lstColumnItems_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call lstColumnItems_DblClick
End Sub

Private Sub lstColumnUsed_DblClick()
    Call cmdColumn_Click(1)
End Sub

Private Sub lstColumnUsed_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call lstColumnUsed_DblClick
End Sub

Private Sub lstSelect_DblClick(Index As Integer)
    Call lstSelect_KeyDown(Index, vbKeyReturn, 0)
End Sub

Private Sub lstSelect_GotFocus(Index As Integer)
    Dim i As Integer, j As Integer
    mblnEditAssistant = False
    mblnEditText = False
    PicLst.Tag = 0
    j = lstSelect(Index).ListCount - 1
    If Index = 0 And j >= 0 Then
        If lstSelect(Index).ListIndex < 0 Then lstSelect(Index).ListIndex = 0
    End If
End Sub

Private Sub txtColumnNo_GotFocus()
    txtColumnNo.SelStart = 0
    txtColumnNo.SelLength = 100
End Sub

Private Sub txtColumnNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" &[]{}+'""|", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtDnInput_GotFocus()
    txtDnInput.SelStart = 0
    txtDnInput.SelLength = 100
    Call ISAssistant(Val(txtDnInput.Tag), txtDnInput)
End Sub

Private Sub txtInput_GotFocus()
    txtInput.SelStart = 0
    txtInput.SelLength = 100
    Call zlControl.TxtSelAll(txtInput)
    mintSymbol = -1
    Call ISAssistant(Val(txtInput.Tag), txtInput)
End Sub

Private Sub txtUpInput_GotFocus()
    txtUpInput.SelStart = 0
    txtUpInput.SelLength = 100
    Call ISAssistant(Val(txtUpInput.Tag), txtUpInput)
End Sub

Private Sub txt_GotFocus(Index As Integer)
    txt(Index).SelStart = 0
    txt(Index).SelLength = 100
    mintSymbol = Index
    Call ISAssistant(Val(txt(Index).Tag), txt(Index))
End Sub

Private Sub lblUpInput_DblClick()
    lblUpInput.Caption = IIf(lblUpInput.Caption = "", "√", "")
    txtUpInput.SetFocus
End Sub

Private Sub lblDnInput_DblClick()
    lblDnInput.Caption = IIf(lblDnInput.Caption = "", "√", "")
    txtDnInput.SetFocus
End Sub

Private Sub lblInput_DblClick()
    lblInput.Caption = IIf(lblInput.Caption = "", "√", "")
End Sub

Private Sub txtUpInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtDnInput.SetFocus
    ElseIf KeyCode = vbKeyRight Then
        If txtUpInput.SelStart = Len(txtUpInput.Text) Then txtDnInput.SetFocus
    ElseIf KeyCode = vbKeyLeft And txtUpInput.SelStart = 0 Then
        Call MoveNextCell(False)
    ElseIf KeyCode = vbKeySpace And txtUpInput.Locked Then
        lblUpInput.Caption = IIf(lblUpInput.Caption = "", "√", "")
    End If
End Sub

Private Sub txtDnInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyRight And txtDnInput.SelStart = Len(txtDnInput.Text)) Then
        Call picDouble_KeyDown(KeyCode, Shift)
    ElseIf KeyCode = vbKeyLeft Then
        If txtDnInput.SelStart = 0 Then txtUpInput.SetFocus
    ElseIf KeyCode = vbKeySpace And txtDnInput.Locked Then
        lblDnInput.Caption = IIf(lblDnInput.Caption = "", "√", "")
    End If
End Sub

Private Sub picMutilInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call MoveNextCell
End Sub

Private Sub picDouble_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Then
        Call MoveNextCell(Not (KeyCode = vbKeyLeft))
    End If
End Sub

Private Sub picInput_GotFocus()
    If txtInput.Visible Then
        txtInput.SetFocus
    End If
End Sub

Private Sub picInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not txtInput.Visible Then
        If KeyCode = vbKeySpace Then
            Call lblInput_DblClick
        End If
    End If
    
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Then
        '移动到下一个单元格
        Call MoveNextCell(Not (KeyCode = vbKeyLeft))
    End If
End Sub

Private Sub lstSelect_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call MoveNextCell
    If Index = 0 And Shift = vbShiftMask And KeyCode = vbKeyUp Then
        KeyCode = 0
        txtLst.SetFocus
    End If
End Sub

Private Sub picMutilInput_GotFocus()
    On Error Resume Next
    txt(0).SetFocus
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Index < txt.Count - 1 Then
            txt(Index + 1).SetFocus
        Else
            Call picMutilInput_KeyDown(KeyCode, Shift)
        End If
    End If
End Sub

Private Sub picDouble_GotFocus()
    If txtUpInput.Visible Then
        txtUpInput.SetFocus
    End If
End Sub

Private Sub picMain_Resize()
    picMain.Left = 0
    VsfData.Width = picMain.Width
    VsfData.Height = IIf(picMain.Height - VsfData.Top < 0, 0, picMain.Height - VsfData.Top)
End Sub

Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And Shift = vbCtrlMask Then Exit Sub
    
    If KeyCode = vbKeyReturn Or _
        (KeyCode = vbKeyRight And txtInput.SelStart = Len(txtInput.Text)) Or _
        (KeyCode = vbKeyLeft And txtInput.SelStart = 0) Then
        Call picInput_KeyDown(KeyCode, Shift)
    End If
End Sub

Private Sub txtUpInput_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("/") Then
        KeyAscii = 0
        txtDnInput.SetFocus
    End If
End Sub

Private Sub txt小结名称_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = Asc(";") Then KeyAscii = 0
End Sub

Private Sub cbo小结_Click()
    If cbo小结.Tag = "" Then Exit Sub
    
    txt开始时点.Enabled = (cbo小结.Text = "临时")
    txt结束时点.Enabled = txt开始时点.Enabled
    If cbo小结.Text <> "临时" Then
        txt开始时点.Text = Split(Split(cbo小结.Tag, ";")(cbo小结.ListIndex), ",")(0)
        txt结束时点.Text = Split(Split(cbo小结.Tag, ";")(cbo小结.ListIndex), ",")(1)
        'txt小结名称.Text = Format(DateAdd("d", -1 * cbo小结范围.ListIndex, zldatabase.Currentdate), "MM-DD") & " " & cbo小结.Text
        txt小结名称.Text = Format(DTPDate.Value, "MM-DD") & " " & cbo小结.Text
        txt小结名称.Tag = 0
    Else
        txt小结名称.Text = ""
        txt开始时点.Text = ""
        txt结束时点.Text = ""
        txt小结名称.Tag = 0
    End If
End Sub


Private Sub UserControl_GotFocus()
    On Error Resume Next
    VsfData.SetFocus
End Sub

Private Sub UserControl_Initialize()
    mblnShow = False
    mblnChange = False
    mblnInit = False
    
'    Set objStream = objFileSys.OpenTextFile("C:\WORKLOG.txt", ForAppending, True)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    '以下字符做为数据分隔符或更新记录集的分隔符，因此不允许录入
    If KeyAscii = 39 Or KeyAscii = 13 Or KeyAscii = Asc("|") Then KeyAscii = 0: Exit Sub
    If KeyAscii = vbKeyEscape And mblnShow Then
        mblnShow = False
        mintType = -1
        Call InitCons
    End If
End Sub

Private Sub cbsThis_Resize()
    Dim lngScaleLeft As Long, lngScaleTop As Long, lngScaleRight As Long, lngScaleBottom As Long
    Call cbsThis.GetClientRect(lngScaleLeft, lngScaleTop, lngScaleRight, lngScaleBottom)
    
    Err = 0: On Error Resume Next
    lblTitle.Move lngScaleLeft, lngScaleTop + 120, lngScaleRight - lngScaleLeft
    With lblSubhead
        .Left = lngScaleLeft + 210: .Width = lngScaleRight - lngScaleLeft - 210 * 2
        .Top = lblTitle.Top + lblTitle.Height + 120
    End With
    picMain.Move lngScaleLeft, lngScaleTop, lngScaleRight, lngScaleBottom - lngScaleTop
    vsfHead.Move lngScaleLeft + 210, lblSubhead.Top + lblSubhead.Height + 45, lngScaleRight - lngScaleLeft - 210 * 2
    VsfData.Move vsfHead.Left, vsfHead.Top + vsfHead.Height - 20, vsfHead.Width
    VsfData.Height = picMain.Height - vsfHead.Height - vsfHead.Top
    
    lblCurPage.Top = picMain.Top
    lblCurPage.Left = picMain.Width - lblCurPage.Width
    
    picInfo.Top = 60
    picInfo.Left = 60
    '表上标签分散处理
    Call zlLableBruit
End Sub

Private Sub UserControl_Terminate()
'    objStream.Close
End Sub

Private Sub SetDockRight(BarToDock As CommandBar, BarOnLeft As CommandBar)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    Dim Left As Long
    Dim Top As Long
    Dim Right As Long
    Dim Bottom As Long
    
    cbsThis.RecalcLayout
    BarOnLeft.GetWindowRect Left, Top, Right, Bottom
    cbsThis.DockToolBar BarToDock, Right, (Bottom + Top) / 2, BarOnLeft.Position
End Sub

Private Function BlowUp(ByRef dblChange As Double) As Double
    '放大：字体，单元格宽度
    BlowUp = dblChange
    If Not mblnBlowup Then Exit Function
    BlowUp = dblChange + (dblChange * 1 / 3)
End Function

Private Sub InitPages()
    Dim intPage As Integer
    Dim cbrItem As CommandBarControl
    Dim cbrCus As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom
    
    '增加页选择
    If Not mcbrPage Is Nothing Then
        mcbrPage.CommandBar.Controls.DeleteAll
    Else
        Set mcbrPage = mcbrToolBar.Controls.Add(xtpControlPopup, clngPage, "页码选择")
        mcbrPage.BeginGroup = True
        mcbrPage.IconId = clngPage
        mcbrPage.Style = xtpButtonIconAndCaption
    End If
    
    Set cbrCus = mcbrToolBar.Controls.Find(, 999902)
    If Not cbrCus Is Nothing Then
        cbrCus.Delete
    End If
    Set cbrCustom = mcbrToolBar.Controls.Add(xtpControlCustom, 999902, "")
    cbrCustom.IconId = 0
    cbrCustom.Handle = picPage.hWnd
        
    For intPage = mint起始页码 To mint结束页 + 1
        Set cbrItem = mcbrPage.CommandBar.Controls.Add(xtpControlButton, conMenu_View_Jump, "第" & intPage & "页", -1, False)
        cbrItem.Parameter = intPage
    Next
    
    txtPage.Text = ""
    mcbrPage.Caption = "页码选择：第" & mint页码 & "页"
    cbsThis.RecalcLayout
End Sub

Private Sub imgSign_Click()
    Call picSign_Click
End Sub

Private Sub lbl验证签名_Click()
    Call picSign_Click
End Sub

Private Sub picSign_Click()
    '加载签名历史记录
    Dim str发生时间 As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    
    vsfSignData.Clear
    str发生时间 = VsfData.TextMatrix(VsfData.ROW, 2)
    gstrSQL = "" & _
        " SELECT A.记录人 AS 签名人,NVL(to_char(A.记录时间,'yyyy-MM-dd hh24:mi:ss'),A.项目名称) AS 签名时间,A.记录内容 AS 签名信息,A.记录标记 AS 签名规则,A.ID,DECODE(A.项目ID,NULL,'有效','未验证') AS 有效性,A.开始版本,NVL(A.项目序号,2) AS 签名规则版本" & vbNewLine & _
        " FROM 病人护理明细 A,病人护理数据 B,病人护理文件 C" & vbNewLine & _
        " WHERE A.记录ID=B.ID And B.文件ID=C.ID AND MOD(A.记录类型,10)=5" & vbNewLine & _
        " AND C.ID=[1] AND B.发生时间=[2] " & vbNewLine & _
        " Order by A.项目名称 Desc"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取签名历史记录", mlng文件ID, CDate(str发生时间))
    
    Set vsfSignData.DataSource = rsTemp
    With vsfSignData
        .ColWidth(0) = 1000
        .ColWidth(1) = 1800
        .ColWidth(2) = 0
        .ColWidth(3) = 0
        .ColWidth(4) = 0
        .ColWidth(6) = 0
        .ColWidth(7) = 0
        .ROW = 1: .COL = 5
    End With
    
    picSign.Visible = False
    With picSignCheck
        .Left = VsfData.Left + (VsfData.Width - .Width) / 2
        .Top = VsfData.Top + (VsfData.Height - .Height) / 2
        .Visible = True
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmd取消_Click()
    picSignCheck.Visible = False
End Sub

Private Sub cmdSignCur_Click()
    '单行验证
    Dim lngLoop As Long
    Dim int版本 As Integer
    Dim strSource As String, str发生时间 As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    
    If (Val(vsfSignData.TextMatrix(vsfSignData.ROW, 4)) = 0) Then Exit Sub
    If (Val(vsfSignData.TextMatrix(vsfSignData.ROW, 7)) < 2) Then
        MsgBox "由于签名规则变化，老版签名数据暂不支持签名校验功能！", vbInformation, gstrSysName
        Exit Sub
    End If
    '获取要签名的内容
    '------------------------------------------------------------------------------------------------------------------
    int版本 = vsfSignData.TextMatrix(vsfSignData.ROW, 6)
    str发生时间 = VsfData.TextMatrix(VsfData.ROW, 2)
    Set rsTemp = GetSignData(str发生时间, int版本)
    Do While Not rsTemp.EOF
        For lngLoop = 0 To rsTemp.Fields.Count - 1
            strSource = strSource & CStr(zlCommFun.NVL(rsTemp.Fields(lngLoop).Value, ""))
        Next
        rsTemp.MoveNext
    Loop
    Debug.Print "验证签名：" & Now & vbCrLf & strSource
    
    '数字签名
    If gobjESign Is Nothing Then
        On Error Resume Next
        Set gobjESign = CreateObject("zl9ESign.clsESign")
        If Err <> 0 Then Err.Clear
        On Error GoTo 0
        If Not gobjESign Is Nothing Then
            Call gobjESign.Initialize(gcnOracle, glngSys)
        End If
    End If
    If gobjESign Is Nothing Then
        MsgBox "电子签名部件未能正确安装，验证操作不能继续！", vbInformation, gstrSysName
        Exit Sub
    End If
    If gobjESign.VerifySignature(strSource, Val(vsfSignData.TextMatrix(vsfSignData.ROW, 4)), 6) Then
        vsfSignData.TextMatrix(vsfSignData.ROW, 5) = "有效"
        Call vsfSignData_EnterCell
    End If
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdSignAll_Click()
    Dim lngSel As Long
    Dim lngRow As Long, lngRows As Long
    '全部验证
    
    lngSel = vsfSignData.ROW
    vsfSignData.Redraw = flexRDNone
    lngRows = vsfSignData.Rows - 1
    For lngRow = 1 To lngRows
        If (vsfSignData.TextMatrix(lngRow, 5) <> "有效") Then
            vsfSignData.ROW = lngRow
            Call cmdSignCur_Click
        End If
    Next
    vsfSignData.ROW = lngSel
    vsfSignData.Redraw = flexRDDirect
End Sub

Private Function ShowSignMarker(Optional ByVal bln外部 As Boolean = False) As Boolean
    Dim str发生时间 As String
    Dim rsTemp As New ADODB.Recordset
    '显示历史签名标记
    
    picSign.Visible = False
    picSignCheck.Visible = False
    If Not bln外部 Then
        If VsfData.COL <> mlngSignName Then Exit Function
    End If
    If VsfData.TextMatrix(VsfData.ROW, mlngSigner) = "" Then Exit Function
    
    str发生时间 = VsfData.TextMatrix(VsfData.ROW, 2)
    gstrSQL = "" & _
        " SELECT A.记录人 AS 签名人,NVL(to_char(A.记录时间,'yyyy-MM-dd hh24:mi:ss'),A.项目名称) AS 签名时间,A.记录内容 AS 签名信息,A.记录标记 AS 签名规则,A.ID,DECODE(A.项目ID,NULL,'有效','未验证') AS 有效性,A.开始版本,NVL(A.项目序号,2) AS 签名规则版本" & vbNewLine & _
        " FROM 病人护理明细 A,病人护理数据 B,病人护理文件 C" & vbNewLine & _
        " WHERE A.记录ID=B.ID And B.文件ID=C.ID AND MOD(A.记录类型,10)=5" & vbNewLine & _
        " AND C.ID=[1] AND B.发生时间=[2] And Rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取签名历史记录", mlng文件ID, CDate(str发生时间))
    If rsTemp.RecordCount = 0 Then Exit Function
    
    With picSign
        .Top = VsfData.Top + VsfData.CellTop + VsfData.CellHeight - .Height
        .Left = VsfData.Left + VsfData.CellLeft + 500
        .Visible = True
    End With
    ShowSignMarker = True
End Function

Private Sub VsfData_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objPopup As CommandBar
    Dim objControl As CommandBarControl
    If VsfData.ROW < VsfData.FixedRows Then Exit Sub
    If Button = 2 And Y >= VsfData.CellTop And Y <= VsfData.CellTop + VsfData.CellHeight Then
        If Val(VsfData.TextMatrix(VsfData.ROW, mlngRecord)) > 0 And VsfData.TextMatrix(VsfData.ROW, mlngSigner) <> "" Then
            Set objPopup = cbsThis.Add("Popup", xtpBarPopup)
            With objPopup.Controls
                Set objControl = .Add(xtpControlButton, conMenu_Tool_SignEarse, "取消签名"): objControl.IconId = 229
                Set objControl = .Add(xtpControlButton, conMenu_Tool_SignAuditCancel, "取消审签"): objControl.IconId = 229
            End With
            objPopup.ShowPopup
        End If
    End If
End Sub

Private Sub vsfSignData_EnterCell()
    cmdSignCur.Enabled = (vsfSignData.TextMatrix(vsfSignData.ROW, 5) <> "有效")
End Sub

Private Function GetSignData(ByVal str发生时间 As String, ByVal int版本 As Integer) As ADODB.Recordset
    On Error GoTo ErrHand
    Dim rsTemp As New ADODB.Recordset
    
    If int版本 = 1 Then
        gstrSQL = "" & _
            "Select a.记录类型,a.项目分组,a.项目序号,a.项目名称,a.项目类型,a.记录内容,a.项目单位,a.记录标记,a.体温部位,a.记录组号,a.复试合格,a.未记说明,a.记录人,a.记录时间" & vbNewLine & _
            "  From 病人护理明细 a, 病人护理数据 b,病人护理文件 C" & vbNewLine & _
            " Where c.ID=[1] And b.发生时间 =[2]" & vbNewLine & _
            "   And a.记录id = b.ID and B.文件ID=C.ID and MOD(A.记录类型,10) <>5 and A.开始版本=1" & vbNewLine & _
            " ORDER BY 项目序号"
    Else
        gstrSQL = "" & _
            "Select a.记录类型,a.项目分组,a.项目序号,a.项目名称,a.项目类型,a.记录内容,a.项目单位,a.记录标记,a.体温部位,a.记录组号,a.复试合格,a.未记说明,a.记录人,a.记录时间" & vbNewLine & _
            "  From 病人护理明细 a, 病人护理数据 b,病人护理文件 C" & vbNewLine & _
            " Where c.ID=[1] And b.发生时间 =[2]" & vbNewLine & _
            "   And a.记录id = b.ID and B.文件ID=C.ID and MOD(A.记录类型,10) <>5" & vbNewLine & _
            "   and (A.开始版本=[3] or (A.开始版本 <[3] and A.终止版本 IS NULL) or (A.开始版本<[3] and A.终止版本>[3]))" & vbNewLine & _
            " ORDER BY 项目序号"
    End If
    Set GetSignData = zlDatabase.OpenSQLRecord(gstrSQL, "提取指定版本的数据", mlng文件ID, CDate(str发生时间), int版本)
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub SignMarker()
    '供外部主程序调用
    If Not ShowSignMarker(True) Then Exit Sub
    Call picSign_Click
End Sub

Private Sub SingerShowType(ByVal vsfObj As VSFlexGrid, ByVal lngStartRow As Long, ByVal lngEndRow As Long)
'-------------------------------------------------
'功能：护士签名人显示方式
''--58414,刘鹏飞,2013-04-10,添加护士、签名列显示模式
'-------------------------------------------------
    Dim lngRow As Integer
    
    Select Case mlngSingerType
        Case 0 '所有行显示
            For lngRow = lngStartRow To lngEndRow
                If mlngOperator > 0 Then vsfObj.TextMatrix(lngRow, mlngOperator) = vsfObj.TextMatrix(lngStartRow, mlngOperator)
                If mlngSignName > 0 Then vsfObj.TextMatrix(lngRow, mlngSignName) = vsfObj.TextMatrix(lngStartRow, mlngSignName)
                If mlngSignTime > 0 Then vsfObj.TextMatrix(lngRow, mlngSignTime) = vsfObj.TextMatrix(lngStartRow, mlngSignTime)
            Next
        Case 1 '首行显示
            For lngRow = lngStartRow To lngEndRow
                If lngRow = lngStartRow Then
                    If mlngOperator > 0 Then vsfObj.TextMatrix(lngRow, mlngOperator) = vsfObj.TextMatrix(lngStartRow, mlngOperator)
                    If mlngSignName > 0 Then vsfObj.TextMatrix(lngRow, mlngSignName) = vsfObj.TextMatrix(lngStartRow, mlngSignName)
                    If mlngSignTime > 0 Then vsfObj.TextMatrix(lngRow, mlngSignTime) = vsfObj.TextMatrix(lngStartRow, mlngSignTime)
                Else
                    If mlngOperator > 0 Then vsfObj.TextMatrix(lngRow, mlngOperator) = ""
                    If mlngSignName > 0 Then vsfObj.TextMatrix(lngRow, mlngSignName) = ""
                    If mlngSignTime > 0 Then vsfObj.TextMatrix(lngRow, mlngSignTime) = ""
                End If
            Next
        Case 3 '尾行显示
            If mlngOperator > 0 Then
                If vsfObj.TextMatrix(lngStartRow, mlngOperator) = "" Then vsfObj.TextMatrix(lngStartRow, mlngOperator) = vsfObj.TextMatrix(lngEndRow, mlngOperator)
            End If
            If mlngSignName > 0 Then
                If vsfObj.TextMatrix(lngStartRow, mlngSignName) = "" Then vsfObj.TextMatrix(lngStartRow, mlngSignName) = vsfObj.TextMatrix(lngEndRow, mlngSignName)
            End If
            If mlngSignTime > 0 Then
                If vsfObj.TextMatrix(lngStartRow, mlngSignTime) = "" Then vsfObj.TextMatrix(lngStartRow, mlngSignTime) = vsfObj.TextMatrix(lngEndRow, mlngSignTime)
            End If
            For lngRow = lngEndRow To lngStartRow Step -1
                If lngRow = lngEndRow Then
                    If mlngOperator > 0 Then vsfObj.TextMatrix(lngRow, mlngOperator) = vsfObj.TextMatrix(lngStartRow, mlngOperator)
                    If mlngSignName > 0 Then vsfObj.TextMatrix(lngRow, mlngSignName) = vsfObj.TextMatrix(lngStartRow, mlngSignName)
                    If mlngSignTime > 0 Then vsfObj.TextMatrix(lngRow, mlngSignTime) = vsfObj.TextMatrix(lngStartRow, mlngSignTime)
                Else
                    If mlngOperator > 0 Then vsfObj.TextMatrix(lngRow, mlngOperator) = ""
                    If mlngSignName > 0 Then vsfObj.TextMatrix(lngRow, mlngSignName) = ""
                    If mlngSignTime > 0 Then vsfObj.TextMatrix(lngRow, mlngSignTime) = ""
                End If
            Next
        Case Else '首尾显示
            '最后一行需要填写封闭签名
            For lngRow = lngStartRow To lngEndRow
                If lngRow = lngStartRow Or lngRow = lngEndRow Then
                    If mlngOperator > 0 Then vsfObj.TextMatrix(lngRow, mlngOperator) = vsfObj.TextMatrix(lngStartRow, mlngOperator)
                    If mlngSignName > 0 Then vsfObj.TextMatrix(lngRow, mlngSignName) = vsfObj.TextMatrix(lngStartRow, mlngSignName)
                    If mlngSignTime > 0 Then vsfObj.TextMatrix(lngRow, mlngSignTime) = vsfObj.TextMatrix(lngStartRow, mlngSignTime)
                Else
                    If mlngOperator > 0 Then vsfObj.TextMatrix(lngRow, mlngOperator) = ""
                    If mlngSignName > 0 Then vsfObj.TextMatrix(lngRow, mlngSignName) = ""
                    If mlngSignTime > 0 Then vsfObj.TextMatrix(lngRow, mlngSignTime) = ""
                End If
            Next
    End Select
End Sub

Private Sub ReSingDataToStart(ByVal vsfObj As VSFlexGrid, ByVal lngStartRow As Long, ByVal lngEndRow As Long)
'--58414,刘鹏飞,2013-04-10,添加护士、签名列显示模式
    If mlngSingerType = 3 Then '尾行签名,在开始将最后一行的信息赋予给开始行，以便后面SingerShowType重新组织显示方式
        If mlngOperator > 0 Then
            If vsfObj.TextMatrix(lngStartRow, mlngOperator) = "" Then vsfObj.TextMatrix(lngStartRow, mlngOperator) = vsfObj.TextMatrix(lngEndRow, mlngOperator)
        End If
        If mlngSignName > 0 Then
            If vsfObj.TextMatrix(lngStartRow, mlngSignName) = "" Then vsfObj.TextMatrix(lngStartRow, mlngSignName) = vsfObj.TextMatrix(lngEndRow, mlngSignName)
        End If
        If mlngSignTime > 0 Then
            If vsfObj.TextMatrix(lngStartRow, mlngSignTime) = "" Then vsfObj.TextMatrix(lngStartRow, mlngSignTime) = vsfObj.TextMatrix(lngEndRow, mlngSignTime)
        End If
    End If
End Sub

Private Function GetRelatiionNo(ByVal strKey As String, Optional ByVal bytType As Byte = 1, Optional ByVal blnCorrelative As Boolean = True) As String
'---------------------------------------------------
'功能:获取汇总项目关联的名称列的项目序号或列号(分类汇总)
'strKey 汇总项目的列号和序号,格式:列号,序号
'bytType 1:项目序号,2:列号
'blnCorrelative TRUE:分类汇总,FALSE:入量导入
'返回值:为空表示汇总项目没有设置关联列
'---------------------------------------------------
    Dim arrItem, arrCorrelative, i As Long
    Dim strValue As String
    If blnCorrelative = True Then
        arrItem = Split(mstrColCorrelative, "|")
    Else
        arrItem = Split(mstrColImCorrelative, "|")
    End If
    For i = 0 To UBound(arrItem)
        arrCorrelative = Split(arrItem(i), ";")
        If InStr(1, strKey, ";") <> 0 Then
            strKey = Split(strKey, ";")(0) & "," & Split(strKey, ";")(1)
        End If
        If strKey = arrCorrelative(1) Then
            If bytType = 1 Then
                strValue = Split(arrCorrelative(0), ",")(1)
            Else
                strValue = Split(arrCorrelative(0), ",")(0)
            End If
            Exit For
        End If
    Next i
    
    GetRelatiionNo = strValue
End Function

Private Function CheckCollectIsData(ByVal lngStartRow As Long, Optional ByVal bytMode As Byte = 0, Optional ByRef lngEditCol As Long = 0) As Boolean
'功能:检查汇总列及关联列是否存在数据，只要有一列存在就退出
'入参：bytMode：主要针对分组数据，是检查整个分组数据还是值检查子数据：0-整个都检查,1- 只检查子数据
'出参：汇总列不为空则返回行号
    Dim strCols As String, strValue As String
    Dim i As Integer, arrCol
    Dim lngRow As Long, lngCol As Long, lngRows As Long, lngRowCount As Long
    If mstrColCollect <> "" Then
        '1、获取汇总相关列号
        arrCol = Split(mstrColCollect, "|")
        For i = 0 To UBound(arrCol)
            strValue = GetRelatiionNo(CStr(arrCol(i)), 2)
            strCols = strCols & "," & IIf(strValue = "", "", strValue & ",") & Split(arrCol(i), ";")(0)
        Next
        strCols = Mid(strCols, 2)
        
        lngStartRow = GetStartRow(lngStartRow)
        '2、检查对应的列是否存在汇总数据
         '如果lngStartRow不是分组起始行，首先获取分组数据的第一行
        If Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) > 1 And bytMode = 0 Then
            lngRow = lngStartRow - Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) + 1
            If Val(VsfData.TextMatrix(lngRow, mlngDemo)) <> 1 Then
                For lngRow = lngStartRow To VsfData.FixedRows Step -1
                    If Val(VsfData.TextMatrix(lngRow, mlngDemo)) = 1 Then
                        Exit For
                    End If
                Next lngRow
                If lngRow >= VsfData.FixedRows Then lngStartRow = lngRow
            End If
        End If
        '获取数据的总行数
        lngRows = lngStartRow
        lngRowCount = Val(VsfData.TextMatrix(lngStartRow, mlngRowCount))
        If lngRowCount <= 0 Then lngRowCount = 1
        lngRows = lngRows + lngRowCount - 1
        
        If Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) = 1 And bytMode = 0 Then
            For lngRow = lngStartRow + lngRowCount To VsfData.Rows - 1
                If Not VsfData.RowHidden(lngRow) Then
                    If Val(VsfData.TextMatrix(lngRow, mlngDemo)) <= 1 Then
                        Exit For
                    End If
                    lngRows = lngRows + 1
                End If
            Next lngRow
        End If
        
        For lngRow = lngStartRow To lngRows
            For lngCol = mlngTime + 1 To mlngNoEditor - 1
                If InStr(1, "," & strCols & ",", "," & lngCol - (cHideCols + VsfData.FixedCols - 1) & ",") > 0 And Trim(FormatValue(VsfData.TextMatrix(lngRow, lngCol))) <> "" Then
                    lngEditCol = lngCol
                    CheckCollectIsData = True
                    Exit Function
                End If
            Next lngCol
        Next lngRow
    End If
    CheckCollectIsData = False
End Function

Private Sub ImportAmount()
'从医嘱导入入量，导入的入量当分组数据处理
    Dim cbrControl As CommandBarControl
    Dim rsImpAmount As ADODB.Recordset
    Dim strDate As String, blnImportName As Boolean, strValue As String
    Dim intCount As Integer, blnFind As Boolean, i As Integer, lngCurRow As Long
    Dim lngNameRow As Long, lngNumRow As Long '名称列，汇总列
    Dim lngNameOrder As Long, lngNumOrder As Long '名称列项目序号,汇总列项目序号
    
    On Error GoTo ErrHand
    For intCount = 0 To UBound(Split(mstrColCollect, "|"))
        If VsfData.COL - (cHideCols + VsfData.FixedCols - 1) = Split(Split(mstrColCollect, "|")(intCount), ";")(0) Then
            lngNumOrder = Split(Split(mstrColCollect, "|")(intCount), ";")(1)
            blnFind = True
            Exit For
        End If
    Next
    If blnFind = False Then Exit Sub
    lngNumRow = VsfData.COL
    strValue = GetRelatiionNo(CStr(Split(mstrColCollect, "|")(intCount)), 2, False)
    If strValue <> "" Then
        lngNameRow = Val(strValue) + (cHideCols + VsfData.FixedCols - 1)
        lngNameOrder = GetRelatiionNo(CStr(Split(mstrColCollect, "|")(intCount)), 1, False)
        blnImportName = True
    End If
    
    '返回记录集内容包含:key,名称,用量
    Set rsImpAmount = frmImportOrder.ShowMe(Me, mlng文件ID, mlng病人ID, mlng主页ID, mint婴儿, blnImportName, lngNumOrder, strDate)
    If rsImpAmount Is Nothing Then Call SetControlValue(lngNumOrder, "", False): Exit Sub
    If rsImpAmount.RecordCount = 0 Then Call SetControlValue(lngNumOrder, "", False): Exit Sub
    '导入入量
    If rsImpAmount.RecordCount > 0 Then rsImpAmount.MoveFirst
    For intCount = 1 To rsImpAmount.RecordCount
        VsfData.COL = lngNumRow
        If mblnShow = False Then Call VsfData_DblClick '追加后会取消编辑，此处需要重新设置
        lngCurRow = GetStartRow(VsfData.ROW)
        If SetControlValue(lngNumOrder, NVL(rsImpAmount("用量").Value)) = True Then
            '确定项目对应的编辑控件
            If blnImportName Then
               VsfData.COL = lngNameRow
               Call SetControlValue(lngNameOrder, NVL(rsImpAmount("名称").Value))
            End If
            If intCount = rsImpAmount.RecordCount Then
                Call MoveNextCell(True, True)
            Else
                Call MoveNextCell(VsfData.COL < mlngNoEditor - 1, True)
            End If
            '完成医嘱信息赋值(入量列)
            If Record_Locate(mrsCellMap, "ID|" & mint页码 & "," & lngCurRow & "," & lngNumRow) = True Then
                mrsCellMap.Fields("标记").Value = NVL(rsImpAmount("key").Value)
                mrsCellMap.Update
            End If
            
            If intCount < rsImpAmount.RecordCount Then
                '如果之前是启用了分组，此处需要禁用分组，直接使用追加
                If mblnGroupNew = True Then Call cbsThis_Execute(cbsThis.FindControl(, conMenu_Edit_Group_New))
                '使用追加功能，追加一行(直接在当前行下面追加)
                Set cbrControl = cbsThis.FindControl(, conMenu_Edit_Group_Append)
                If Not cbrControl Is Nothing Then
                    Call cbsThis_Execute(cbrControl)
                Else
                    Exit For
                End If
            End If
        End If
        rsImpAmount.MoveNext
    Next
    
    '隐蔽已显示的录入控件
    Select Case mintType
    Case 0, 3
        picInput.Visible = False
    Case 1, 2
        lstSelect(mintType - 1).Visible = False
        If mintType = 1 Then
            txtLst.Visible = False
            PicLst.Visible = False
        End If
    Case 4, 5
        picDouble.Visible = False
    Case 6
        picMutilInput.Visible = False
    Case 7
        picDoubleChoose.Visible = False
    Case 8
        picYear.Visible = False
    End Select
    cmdWord.Visible = False
    mintType = -1
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function SetControlValue(ByVal lngOrder As Long, ByVal strValue As String, Optional ByVal blnMode As Boolean = True) As Boolean
'功能：根据项目序号完成相应编辑控件的赋值(必须在编辑状态下)
'       blnMode:True 赋值,False 设置焦点
    Dim i As Integer, j As Integer
    Dim objControl As Object
    If mintType = -1 Then Exit Function
    On Error Resume Next
    If blnMode = True Then
        Select Case mintType
            Case 0
                txtInput.Text = strValue
            Case 1, 2
                If strValue <> "" Then
                    strValue = Replace(strValue, vbCrLf, "")
                    txtLst.Text = strValue
                    PicLst.Tag = "1"
                    j = lstSelect(mintType - 1).ListCount - 1
                    For i = 0 To j
                        '单选的第一个项目是清除选择，需要跳过此项，多选项目则直接进入
                        If Not (mintType = 1 And i = 0) Then
                            If InStr(1, "," & strValue & ",", "," & Mid(lstSelect(mintType - 1).List(i), InStr(1, lstSelect(mintType - 1).List(i), "-") + 1) & ",") <> 0 Then
                                lstSelect(mintType - 1).Selected(i) = True
                                txtLst.Text = ""
                                PicLst.Tag = "0"
                            End If
                        End If
                    Next
                Else
                    txtLst.Text = ""
                    PicLst.Tag = "0"
                End If
            Case 3
                If strValue = "√" Or strValue = "" Then lblInput.Caption = strValue
            Case 4
                If lngOrder = Val(txtUpInput.Tag) Then
                    txtUpInput.Text = strValue
                Else
                    txtDnInput.Text = strValue
                End If
            Case 5
                If lngOrder = Val(lblUpInput.Tag) Then
                     If strValue = "√" Or strValue = "" Then lblUpInput.Caption = strValue
                Else
                     If strValue = "√" Or strValue = "" Then lblDnInput.Caption = strValue
                End If
            Case 6
                For i = 0 To txt.Count - 1
                    If lngOrder = Val(txt(i).Tag) Then
                        txt(i).Text = strValue
                    End If
                Next
            Case 7
                If lngOrder = Val(cboChoose(0).Tag) Then
                    j = 0
                Else
                    j = 1
                End If
                For i = 0 To cboChoose(j).ListCount - 1
                    If strValue = cboChoose(j).List(i) Then
                        cboChoose(j).ListIndex = i
                    End If
                Next
            Case Else
                SetControlValue = False
                Exit Function
        End Select
    Else
        Select Case mintType
            Case 0
                Set objControl = txtInput
            Case 1, 2
                Set objControl = lstSelect(mintType - 1)
            Case 3
                Set objControl = lblInput
            Case 4
                If lngOrder = Val(txtUpInput.Tag) Then
                    Set objControl = txtUpInput
                Else
                    Set objControl = txtDnInput
                End If
            Case 5
                If lngOrder = Val(lblUpInput.Tag) Then
                     Set objControl = lblUpInput
                Else
                     Set objControl = lblDnInput
                End If
            Case 6
                For i = 0 To txt.Count - 1
                    If lngOrder = Val(txt(i).Tag) Then
                        Set objControl = txt(i)
                    End If
                Next
            Case 7
                If lngOrder = Val(cboChoose(0).Tag) Then
                    j = 0
                Else
                    j = 1
                End If
                Set objControl = cboChoose(j)
            Case Else
                Set objControl = Nothing
        End Select
    End If
    If Not objControl Is Nothing Then
        If objControl.Visible And objControl.Enabled Then objControl.SetFocus
    End If
    SetControlValue = True
    If Err <> 0 Then Err.Clear
End Function


Private Function GetSelectRowRecordId(ByVal lngRow As Long) As String
    '功能：返回指定行记录ID信息,分组数据则返回改组内数据的记录ID，以“，”号分割
    Dim lngDemo As Long, lngStart As Long
    Dim strRecordid As String, lngStartID As Long
    
    If lngRow < VsfData.FixedRows Or lngRow > VsfData.Rows Then Exit Function
    lngStart = GetStartRow(lngRow)
    lngDemo = VsfData.TextMatrix(lngStart, mlngDemo)
    If lngDemo > 1 Then '数据为分组数据,需找到起始行
        lngRow = lngStart
        lngStart = lngRow - lngDemo + 1
        If VsfData.TextMatrix(lngStart, mlngDemo) <> 1 Then
            For lngStart = lngRow To VsfData.FixedRows Step -1
                lngRow = lngStart
                Exit For
            Next lngStart
            If lngStart < VsfData.FixedRows Then Exit Function
            lngStart = lngRow
        End If
    End If
    strRecordid = ""
    lngDemo = VsfData.TextMatrix(lngRow, mlngDemo)
    lngStartID = Val(VsfData.TextMatrix(lngRow, mlngRecord))
    If lngDemo = 1 Then
        For lngRow = lngStart To VsfData.Rows
            If Val(VsfData.TextMatrix(lngRow, mlngDemo)) <= 1 And lngStartID <> Val(VsfData.TextMatrix(lngRow, mlngRecord)) Then Exit For
            If Val(VsfData.TextMatrix(lngRow, mlngRecord)) > 0 And InStr(1, "," & strRecordid & ",", "," & Val(VsfData.TextMatrix(lngRow, mlngRecord)) & ",") = 0 Then
                strRecordid = strRecordid & "," & Val(VsfData.TextMatrix(lngRow, mlngRecord))
            End If
        Next
        If Left(strRecordid, 1) = "," Then strRecordid = Mid(strRecordid, 2)
    Else
        strRecordid = IIf(Val(VsfData.TextMatrix(lngStart, mlngRecord)) > 0, Val(VsfData.TextMatrix(lngStart, mlngRecord)), "")
    End If
    GetSelectRowRecordId = strRecordid
End Function
