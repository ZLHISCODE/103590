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
      Begin VB.CommandButton cmdȡ�� 
         Caption         =   "ȡ��"
         Height          =   350
         Left            =   3690
         TabIndex        =   71
         ToolTipText     =   "ȡ��"
         Top             =   2370
         Width           =   840
      End
      Begin VB.CommandButton cmdSignCur 
         Caption         =   "��֤"
         Height          =   350
         Left            =   2790
         TabIndex        =   70
         ToolTipText     =   "ȷ��"
         Top             =   2370
         Width           =   840
      End
      Begin VB.CommandButton cmdSignAll 
         Caption         =   "ȫ��"
         Height          =   350
         Left            =   270
         TabIndex        =   69
         ToolTipText     =   "ȷ��"
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
            Name            =   "����"
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
         Caption         =   "������ǩ����ʷ��¼����ѡ������֤��Ҳ�ɽ���ȫ����֤��"
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
            Name            =   "����"
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
         Caption         =   "ҳ"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "��λ"
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
            Caption         =   "��ܰ��ʾ:"
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
            Caption         =   "ѡ��"
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
            Caption         =   "¼�룺"
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
         Begin VB.Label lbl��֤ǩ�� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��֤ǩ��"
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
            Name            =   "����"
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
               Name            =   "����"
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
            Caption         =   "������¼"
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
         Caption         =   "ȫѡ"
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
               Name            =   "����"
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
            Caption         =   "��"
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
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
         Caption         =   "һ�㻤���¼��"
         Height          =   180
         Left            =   3450
         TabIndex        =   14
         Top             =   30
         Width           =   1275
      End
      Begin VB.Label lblSubhead 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����:##"
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
            Text            =   "��Ŀ���"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "��Ŀ����"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "��λ"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.CommandButton cmdFilterOK 
         Height          =   315
         Left            =   2460
         Picture         =   "usrTendFileEditor.ctx":ED16
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "ȷ��"
         Top             =   2310
         Width           =   450
      End
      Begin VB.CommandButton cmdFilterCancel 
         Height          =   315
         Left            =   3000
         Picture         =   "usrTendFileEditor.ctx":F2A0
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "ȡ��"
         Top             =   2310
         Width           =   450
      End
      Begin VB.CommandButton cmdColumn 
         Caption         =   "ѡ��(&S)"
         Height          =   300
         Index           =   0
         Left            =   2430
         TabIndex        =   28
         Top             =   1395
         Width           =   1100
      End
      Begin VB.CommandButton cmdColumn 
         Caption         =   "ɾ��(&E)"
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
            Text            =   "��Ŀ���"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "��Ŀ����"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "��λ"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label lblFind 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
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
         Caption         =   "�ѷ������ݣ�������������á�"
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
         Caption         =   "��ѡ�����¼��Ŀ"
         Height          =   180
         Left            =   105
         TabIndex        =   25
         Top             =   180
         Width           =   1440
      End
      Begin VB.Label lblColumnNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ͷ����"
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
      Begin VB.TextBox txt����ʱ�� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   3480
         MaxLength       =   5
         TabIndex        =   45
         Top             =   540
         Width           =   1365
      End
      Begin VB.TextBox txt��ʼʱ�� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   930
         MaxLength       =   5
         TabIndex        =   43
         Top             =   540
         Width           =   1365
      End
      Begin VB.ComboBox cbo��ʶ 
         Height          =   300
         Left            =   930
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   900
         Width           =   3915
      End
      Begin VB.TextBox txtС������ 
         Height          =   300
         Left            =   930
         TabIndex        =   49
         Top             =   1260
         Width           =   3885
      End
      Begin VB.ComboBox cboС�� 
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
         ToolTipText     =   "ȡ��"
         Top             =   2850
         Width           =   450
      End
      Begin VB.CommandButton cmdOk 
         Height          =   315
         Left            =   3750
         Picture         =   "usrTendFileEditor.ctx":FDB4
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "ȷ��"
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
            Name            =   "����"
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
         Caption         =   "������Ŀ"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   150
         TabIndex        =   50
         Top             =   1635
         Width           =   720
      End
      Begin VB.Label lbl����ʱ�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�� ����ʱ��"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2430
         TabIndex        =   44
         Top             =   600
         Width           =   990
      End
      Begin VB.Label lbl��ʼʱ�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ʼʱ��"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   150
         TabIndex        =   42
         Top             =   600
         Width           =   720
      End
      Begin VB.Label lbl��ʶ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ʶ"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   510
         TabIndex        =   46
         Top             =   960
         Width           =   360
      End
      Begin VB.Label lblʱ�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ʱ��"
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
         Caption         =   "�������ݺ���ʾ��ȷ����"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   690
         TabIndex        =   54
         Top             =   2940
         Width           =   2010
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblС������ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   510
         TabIndex        =   48
         Top             =   1320
         Width           =   360
      End
      Begin VB.Label lblС�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "С��"
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
'��������:
'1.�����¼ͬһʱ��ֻ���ܴ���һ����¼
'2.�����¼�в���Ҫ�����µ����� , ��¼�����Ƿ����, �ܲ������, �����˵����ݲż�¼
'3.¼�뻤���¼����ʱ,�����¼������ݴ�����������, ����ȡ����
'4.�����¼���в���Ҫ¼�������¼�������׾����ȷ��Ҫ��¼���ڻ���ժҪ�������͵�����
'#ʵ��ԭ��:
'1.�����û��޸Ĺ�������,�����ṩ�༭״̬ҳ���л��Ĺ���,���û��޸Ĺ���ҳ���ݽ�����ҳ����,���ٳ���ʵ���Ѷ�
'2.���Ӽ�¼����¼��Щҳ��Щ��Ԫ���û��޸Ĺ�
'3.�κα༭(ճ��,�������),����Ҫ���¼���ÿ�����ݵ�ռ����


'*******************************************************
'2012-01-06;zyb'���Ŀѡ���������ı���Ŀ,һ��ֻ�ܰ�һ���ı���Ŀ



Public mblnEditable As Boolean
'Public objFileSys As New FileSystemObject
'Public objStream As TextStream

Private mblnRestore As Boolean              '���¼������ݻ��ǻָ�ҳ������
Private mfrmParent As Object
Private mblnInit As Boolean
Private mblnShow As Boolean                 '�Ƿ���ʾ¼���
Private mblnVerify As Boolean               '�Ƿ���ǩģʽ(���޸�,����������и���ճ������Ȳ���,ֻ���޸�)
Private mstrVerify As String                '�ȴ���ǩ��ID��
Private mobjVerify As Collection        '�ȴ���ǩ������Ϣ(key��ż�¼ID,��ϢΪ:ԭ����ʱ��)--81535:��ǩ���޸�ʱ�����
Private mintVerify As Integer               '��ǰ����Ա����߼���
Private mintVerify_Last As Integer          '��ѡ��ǩ��¼����߼���
Private mblnBlowup As Boolean               '�Ŵ�񣿷Ŵ�1/3��������9�ŷŴ�Ϊ12��
Private mblnChange As Boolean               '�Ƿ��޸�����
Private mstrData As String                  '����༭״̬ǰ����֮ǰ������
Private mintNORule As Integer               '�����ļ�ҳ�����
Private mintPreDays As Long
Private mstrMaxDate As String
Private mlngSingerType As Long              '��ʿ��ǩ������ʾģʽ����������ʾ������β��ʾ�ȣ�

Private mint��ʼҳ�� As Integer
Private mint����ҳ As Integer
Private mintҳ�� As Integer
Private mlng�ļ�ID As Long
Private mlng��ʽID As Long
Private mlng����ID As Long
Private mlng��ҳID As Long
Private mlng����ID As Long
Private mlng����ID As Long
Private mintӤ�� As Integer
Private mbln���� As Boolean                 '�Ƿ���Ҫ¼������
Private mstrPrivs As String

Private mintSymbol As Integer               '��ǰ�ؼ�����
Private mstrSymbol As String                 '�����ַ�
Private mblnClear As Boolean                '���Ϊ��,���mrsDataMap��¼��;����ҳʱӦ����,�����û��޸ĵ������Ա���ʾ������ʹ��
Private mstrCollectItems As String         '������Ŀ����
Private mstrColCollect As String             '������Ŀ�м���:col;1|col;4,5
Private mstrColCorrelative As String       '������Ŀ�����м���:COl,3;COl,4|COl,5;COl,6(�����к�,��Ŀ���;������,��Ŀ���),��Ҫ��Է������
Private mstrColImCorrelative As String    '������Ŀ�����м���:COl,3;COl,4|COl,5;COl,6(�����к�,��Ŀ���;������,��Ŀ���),��Ҫ�����������
Private mblnCorrelative As Boolean        '�Ƿ������˷������
Private mstrCOLNothing As String          'δ�󶨵��м���+���Ŀ��(���ܻ��Ŀ���Ƿ��)
Private mstrCOLActive As String             '��м���
Private mstrCatercorner As String           '�жԽ��߼���
Private mblnEditAssistant As Boolean        '��ǰѡ�����Ŀ�Ƿ�������дʾ�ѡ��
Private mblnEditText As Boolean             'ѡ�����Ŀ�Ƿ����ı���Ŀ
Private mlngPageRows As Long                '���ļ���ʽһҳ����ʾ��������
Private mArrPageInfo() As String            'ÿһҳ��¼������ʾ
Private mlngLitterRows() As Long            '��¼��ҳ�����е�����
Private mlngCurLitterRows() As Long         '��¼��ҳ�������ڱ�ҳ��ʵ������
Private mlngOverrunRows As Long             '����������
Private mlngReduceRow As Long               '����������(�ϲ��ļ���ʼҳ����ʼ�п��ܲ��Ǵ�1��ʼ)
Private mlngRowCount As Long                '��ǰ��¼������
Private mlngRowCurrent As Long              '��ǰ��¼�ڱ�ҳ��ʵ������
Private mlngStartRowPage As Long            '��ǰ��¼�Ŀ�ʼҳ��
Private mlngStartRowNo As Long              '��ǰ��¼���Ŀ�ʼ�к�
Private mlngDate As Long                    '����
Private mlngTime As Long                    'ʱ��
Private mlngChoose As Long                  'ѡ����
Private mlngYear As Long                    '���:�����ڸ�ʽʱ��ʾ
Private mlngOperator As Long                '��ʿ
Private mlngJoinSignName As Long            '����ǩ����
Private mlngSignLevel As Long               'ǩ������
Private mlngSigner As Long                  'ǩ����Ϣ
Private mlngSignName As Long                'ǩ����
Private mlngSignTime As Long                'ǩ��ʱ��
Private mlngRecord As Long                  '��¼ID
Private mlngNoEditor As Long                '��ֹ�༭��,���ڻ�ʿ�����Ի�ʿ��Ϊ׼,�����ڻ�ʿ������ǩ����Ϊ׼
Private mlngCollectType As Long             '�������
Private mlngCollectText As Long             '�����ı�
Private mlngCollectStyle As Long            '���ܱ��
Private mlngCollectDay As Long              '��������:0-����;1-����
Private mlngCollectStart As Long            '���ܿ�ʼʱ��
Private mlngCollectEnd As Long              '���ܽ���ʱ��
Private mlngDemo As Long                    '������
Private mlngActTime As Long                 '����ʱ��

Private mblnGroupNew As Boolean             '�����־
Private mstrGroupRow As String              '������ʼ��
Private mblnGroupApp As Boolean             '׷�ӷ���ģʽ?
Private mblnSign As Boolean                 '�Ƿ�ǩ��
Private mblnArchive As Boolean              '�Ƿ�鵵
Private mintType As Integer                 '��¼��ǰ�ı༭ģʽ
Private mintCollectDef As Integer           'ȱʡС���ʽ
Private mlngCollectColor As Long            'С���ʶ��ɫ
Private mintPageSpan As Integer             '��ҳ��ʾ��1-��ǰҳ��2-��ҳ����ʾ
Private mintSignMode As Integer             '��ǩģʽ:0-Ƹ��ְ��+��ǩȨ��;1-��ǩȨ��
Private mblnDateAd As Boolean               '������д?
Private mstr��ʼʱ�� As String              '��ǰ�ļ��Ŀ�ʼʱ��
Private mstr����ʱ�� As String              '��ǰ�ļ��Ľ���ʱ��
Private mstrYears As String                 '��ѡȡ����ݷ�Χ
Private CellRect As RECT
Private mbln��ʿ As Boolean                '�Ƿ���˻�ʿ��

Private rsTemp As New ADODB.Recordset
Private mrsItems As New ADODB.Recordset             '���л����¼��Ŀ�嵥
Private mrsElement As New ADODB.Recordset           'ʹ���ڻ����¼���ı�ǩҪ��
Private mrsSelItems As New ADODB.Recordset          '��ǰ¼��Ļ����¼��Ŀ�嵥
Private mrsDataMap As New ADODB.Recordset           '��ǰ����Ա¼������ݾ���,���¼����ʽһ��,���������ȫ�������Ա�Ѹ�ٻָ�
Private mrsCellMap As New ADODB.Recordset           '�༭�������ݾ���,�ֶ���:ҳ��,�к�,�к�,��¼ID,����,��λ,ɾ��
Private mrsCopyMap As New ADODB.Recordset           '����������

Private mblnElement As Boolean                      '�Ƿ�����Զ����ǩҪ��
Private Enum ColIcon
    ǩ�� = 1
    ��ǩ = 2
    ����ǩ�� = 3
End Enum
Private Enum SignLevel
    ���� = 1
    ���� = 2
    �м� = 3
    ʦ�� = 4
    Աʿ = 5
    δ���� = 9
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

'��¼�ϴ�ѡ����,����,�Ա�ˢ�º����¶�λ
Dim lngLastRow As Long
Dim lngLastTopRow As Long
Dim lngLastPatientID As Long

Private mstrTag As String           '�ݴ�

'�����ļ���ʽ�������
Private mintTabTiers As Integer     '��ͷ���
Private mintTagFormHour As Integer  '��ʼʱ������
Private mintTagToHour As Integer    '��ֹʱ������
Private mobjTagFont As New StdFont  '������ʽ����
Private mlngTagColor As Long        '������ʽ��ɫ
Private mstrPaperSet As String      '��ʽ
Private mstrPageHead As String      'ҳü
Private mstrPageFoot As String      'ҳ��
Private mblnChildForm As Boolean
Private mstrSubhead As String       '���ϱ�ǩ
Private mstrTabHead As String       '��ͷ��Ԫ
Private mstrColWidth As String      '�п����д�
Private mstrColumns As String       '��ǰ�����ļ����ж�Ӧ����Ŀ
Private lngCurColor As Long, strCurFont As String, objFont As StdFont
'����򿪻����¼�ļ���SQL���������ط�Ҳ��ʹ�ã������޸�
Private mstrSQL�� As String
Private mstrSQL�� As String
Private mstrSQL�� As String
Private mstrSQL���� As String
Private mstrSQL As String

Private mcbrToolBar As CommandBar
Private mcbrPage As CommandBarControl
Const clngPage As Long = 3906
Const clngPageLocate As Long = 3907

'######################################################################################################################
'**********************************************************************************************************************
'��#�ָ��������ڵĴ��붼���ͼ���,û�±�
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

Private Const WHITE_BRUSH = 0    '��ɫ����
Private Const cdblWidth As Double = 6          'һ��Ӣ���ַ��Ŀ��
Private Const cHideCols = 4         'ǰ׺������:����,ʱ��,ѡ��,���
Private Const cControlFields = 2    '��¼��������:ҳ��,�к�

Private Function GetRBGFromOLEColor(ByVal dwOleColour As Long) As Long
    '��VB����ɫת��ΪRGB��ʾ
    Dim clrref As Long
    Dim r As Long, g As Long, b As Long
    
    OleTranslateColor dwOleColour, 0, clrref
    
    b = (clrref \ 65536) And &HFF
    g = (clrref \ 256) And &HFF
    r = clrref And &HFF
    
    GetRBGFromOLEColor = RGB(r, g, b)
End Function

Private Function GetSymbolWidth(ByVal strPara As String) As Double
    'ȱʡ������9��,�������Сͬ�ȷŴ�
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
    '��ͼ���
    Dim lngPen As Long
    Dim lngOldPen As Long
    Dim lngBrush As Long
    Dim lngOldBrush As Long
    Dim lpPoint As POINTAPI
    Dim t_ClientRect As RECT
    On Error GoTo ErrHand
    '******************************************
    '�ڴ��¼��в��ܶԵ�Ԫ����κ����Ը�ֵ,����Celldata,�����������¼�����ѭ��,���¹��������ʱ���޷�����������
    '******************************************
    'ʹ��ƥ��ı���ɫ��ǰ��ɫ����������ı������
    Done = True
    If Not mblnInit Then Exit Sub
    If VsfData.RowHidden(ROW) Then Exit Sub
    Done = False
    
    strText = FormatValue(VsfData.TextMatrix(ROW, COL))
    If IsDiagonal(COL) And InStr(1, strText, "/") <> 0 Then
        blnDraw = True
        '����ֵ
        strLeft = Split(strText, "/")(0)
        strRight = Mid(strText, InStr(1, strText, "/") + 1)
        lngLeft = LenB(StrConv(strLeft, vbFromUnicode))
        lngRight = LenB(StrConv(strRight, vbFromUnicode))
        'ȡ�ַ����
        dblWidth = GetSymbolWidth(strRight)
        '�趨�ͻ������С
        With t_ClientRect
            .Left = Left + 1
            .Top = Top + 1
            .Right = Right - 1
            .Bottom = Bottom - 1
        End With
        
        '1���������
        '�����뱳��ɫ��ͬ��ˢ��
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
        'ʹ�ø�ˢ����䱳��ɫ
        lngOldBrush = SelectObject(hDC, lngBrush)
        Call FillRect(hDC, t_ClientRect, lngBrush)
        '����������ʱʹ�õ�ˢ�Ӳ���ԭˢ��
        Call SelectObject(hDC, lngOldBrush)
        Call DeleteObject(lngBrush)
        
        '2��׼������
        '�����»���
        Call SetTextColor(hDC, lngForeColor)
        lngPen = CreatePen(0, 1, lngForeColor)
        lngOldPen = SelectObject(hDC, lngPen)
        '����
        Call MoveToEx(hDC, Left, Bottom - 2, lpPoint)
        Call LineTo(hDC, Right, Top)
        '����ı�
        Call TextOut(hDC, Left, Top, strLeft, lngLeft)
        Call TextOut(hDC, IIf(Right - dblWidth >= Left, Right - dblWidth, Left), Bottom - 16, strRight, lngRight)
        
        '��ԭ���ʲ�����
        Call SelectObject(hDC, lngOldPen)
        Call DeleteObject(lngPen)
        
        '�������ͼ
        Done = True
    End If
    
    '3������ǻ����У���������⴦��
    
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
    '�����»���
    lngPen = CreatePen(0, 1, mlngCollectColor)
    lngOldPen = SelectObject(hDC, lngPen)
    
    Select Case Val(VsfData.TextMatrix(ROW, mlngCollectStyle))
    Case 1  '���»�����(��ʼ�л��Ϻ��ߣ������л��º���)
        '����
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
    Case 2  '��������˫����
        If InStr(1, "|" & mstrColCollect & ";", "|" & COL - (cHideCols + VsfData.FixedCols - 1) & ";") <> 0 Then 'And Val(VsfData.TextMatrix(ROW, COL)) <> 0 Then
            '����
            Call MoveToEx(hDC, Left, Bottom - 4, lpPoint)
            Call LineTo(hDC, Right, Bottom - 4)
            Call MoveToEx(hDC, Left, Bottom - 2, lpPoint)
            Call LineTo(hDC, Right, Bottom - 2)
        End If
    Case 3  '�Ϻ���
        '����
        lngRowCount = Val(VsfData.TextMatrix(ROW, mlngRowCount))
        If FormatValue(VsfData.TextMatrix(ROW, mlngRowCount)) = lngRowCount & "|1" Then
            Call MoveToEx(hDC, Left, Top + IIf(ROW = VsfData.FixedRows, 1, 0), lpPoint)
            Call LineTo(hDC, Right, Top + IIf(ROW = VsfData.FixedRows, 1, 0))
        End If
    Case 4 '�������µ�����
        If InStr(1, "|" & mstrColCollect & ";", "|" & COL - (cHideCols + VsfData.FixedCols - 1) & ";") <> 0 Then 'And Val(VsfData.TextMatrix(ROW, COL)) <> 0 Then
            '����
            Call MoveToEx(hDC, Left, Bottom - 2, lpPoint)
            Call LineTo(hDC, Right, Bottom - 2)
        End If
    End Select
    
    '��ԭ���ʲ�����
    Call SelectObject(hDC, lngOldPen)
    Call DeleteObject(lngPen)
End Sub

'######################################################################################################################
'**********************************************************************************************************************
'��#�ָ��������ڵĴ��붼��������,û�±�
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
        If lngPos > 0 Then Exit Sub     '��Ϊ��,��ʾ�������ַ���������
    Next
    strLine(1) = 1
End Sub

Private Function TrimStr(ByVal str As String) As String
'���ܣ�ȥ���ַ�����\0�Ժ���ַ�������ȥ�����˵Ŀո�

    If InStr(str, Chr(0)) > 0 Then
        TrimStr = Trim(Left(str, InStr(str, Chr(0)) - 1))
    Else
        TrimStr = Trim(str)
    End If
End Function

Private Function TruncZero(ByVal strInput As String) As String
'���ܣ�ȥ���ַ�����\0�Ժ���ַ�
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
    'ֻ�ṩ������,ѡ����������ı���Ļ��Ŀ
    '�󶨻��Ŀ(��һ����Ŀ������,��������Ŀʱ,��Ŀ���ͱ���=0����Ŀ��ʾֻ������ֵ,ѡ������,��������Ŀ��Ŀ��������Ŀ��ʾ��������һ��)
    '51883,������,2012-08-02,�ṩ��ѡ�Ͷ�ѡ���Ŀ�İ�
    '100334,����,2016-09-20,���Ŀʹ�ÿ�����������
    gstrSQL = "" & _
        " SELECT A.��Ŀ���,A.��λ,A.��Ŀ����,B.��ͷ����,NVL(B.��־,0) AS ��־" & vbNewLine & _
        " FROM" & vbNewLine & _
        "     (SELECT A.��Ŀ���,B.��λ,B.��λ||A.��Ŀ���� AS ��Ŀ����" & vbNewLine & _
        "     FROM �����¼��Ŀ A,���²�λ B" & vbNewLine & _
        "     WHERE A.��Ŀ��� =B.��Ŀ���(+) AND A.��Ŀ����=2 And NVL(A.Ӧ�ó���,0)<>1 And " & vbNewLine & _
        "     (A.���ÿ��� = 1 Or (A.���ÿ��� = 2 And Exists (Select 1 From �������ÿ��� C Where b.��Ŀ��� = c.��Ŀ��� And c.����id = [4])))) A," & vbNewLine & _
        "     (SELECT A.��ͷ����,A.��Ŀ���,A.��λ||B.��Ŀ���� AS ��Ŀ����,1 AS ��־" & vbNewLine & _
        "     FROM ���˻�����Ŀ A,�����¼��Ŀ B" & vbNewLine & _
        "     WHERE A.��Ŀ���=B.��Ŀ��� AND A.�ļ�ID=[1] AND A.ҳ��=[2] AND A.�к�=[3] ) B" & vbNewLine & _
        " WHERE A.��Ŀ���=B.��Ŀ���(+) AND A.��Ŀ����=B.��Ŀ����(+)" & vbNewLine & _
        " ORDER BY A.��Ŀ���"
    Set rsActive = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡδ���õĻ��Ŀ", mlng�ļ�ID, mintҳ��, intCol, mlng����ID)
    If rsActive.RecordCount = 0 Then
        RaiseEvent AfterRowColChange("û�пɹ�ѡ��Ļ��Ŀ�����ڻ�����Ŀ����ģ���н������ã�", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    
    '������Ŀ
    '65671:������,2013-09-22,����֮ǰ��ջ��Ŀ��ͷ����
    txtFind.Text = ""
    txtColumnNo.Text = ""
    lstColumnItems.ListItems.Clear
    lstColumnUsed.ListItems.Clear
    With rsActive
        Do While Not .EOF
            If !��־ = 1 Then
                txtColumnNo.Text = NVL(!��ͷ����)
                Set lstItem = lstColumnUsed.ListItems.Add(, Now & "_" & !��Ŀ��� & "_" & lstColumnUsed.ListItems.Count, !��Ŀ���)
                lstItem.SubItems(1) = !��Ŀ����
                lstItem.SubItems(2) = NVL(!��λ)
            Else
                Set lstItem = lstColumnItems.ListItems.Add(, Now & "_" & !��Ŀ��� & "_" & lstColumnItems.ListItems.Count + 100, !��Ŀ���)
                lstItem.SubItems(1) = !��Ŀ����
                lstItem.SubItems(2) = NVL(!��λ)
            End If
            .MoveNext
        Loop
    End With
    
    '���ÿؼ����꣨��߻��ұ߳�����Ļ��С���һ�����ʾ����������Ϊ������ʾ��
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
    
    '53588:������,2013-4-25,�޸����ݵ�ʱ��С�ڲ�����Ժʱ�䣬���ţ�����������ʾ����
    '�磺�������ʱ��Ϊ2013-03-13 11:23:34 �ļ���ʼʱ��������ͬ����ʱ¼������ʱ��Ϊ 2013-03-13 11:23
    '�ͻᵼ���޷���ȡ���ţ�ӦΪ���������ʱ��Ϊ2013-03-13 11:23:00 С���˲������ʱ�䵼���޷���ȡ������
    '��ȡ���˵���Ժʱ��
    If mintӤ�� = 0 Then
        gstrSQL = "Select ��ʼʱ��, Sysdate As ����ʱ��" & vbNewLine & _
            " From ���˱䶯��¼" & vbNewLine & _
            " Where ����id = [1] And ��ҳid = [2] And ��ʼԭ�� = 2" & vbNewLine & _
            " Union All" & vbNewLine & _
            " Select ��ʼʱ��, Sysdate As ����ʱ��" & vbNewLine & _
            " From ���˱䶯��¼ a" & vbNewLine & _
            " Where a.����id = [1] And a.��ҳid = [2] And a.��ʼԭ�� = 1 And Not Exists" & vbNewLine & _
            " (Select 1 From ���˱䶯��¼ Where ����id = a.����id And ��ҳid = a.��ҳid And ��ʼԭ�� = 2)"

    Else
        gstrSQL = " Select   ����ʱ�� AS ��ʼʱ��,sysdate AS ����ʱ�� From ������������¼ Where ����ID=[1] And ��ҳID=[2] And ���=[3]"
    End If
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��Ժ���ڻ��������", mlng����ID, mlng��ҳID, mintӤ��)
    
    '��ȡָ��ҳ������ݷ���ʱ�䷶Χ
    gstrSQL = " Select  MIN(����ʱ��) ��ʼʱ��,MAX(����ʱ��) AS ����ʱ�� From ���˻����ӡ Where �ļ�ID=[1] And (��ʼҳ��=[2] OR ����ҳ��=[2])"
    Call SQLDIY(gstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡָ��ҳ������ݷ���ʱ�䷶Χ", mlng�ļ�ID, IIf(mintҳ�� < mint����ҳ + 1, mintҳ��, mint����ҳ))
    If NVL(rsTemp!��ʼʱ��) = "" Then
        strPeriod = Format(rs!��ʼʱ��, "yyyy-MM-dd HH:mm:ss") & "��" & Format(rs!����ʱ��, "yyyy-MM-dd HH:mm:ss")
    Else
        If Format(rsTemp!��ʼʱ��, "yyyy-MM-dd HH:mm:ss") < Format(rs!��ʼʱ��, "yyyy-MM-dd HH:mm:ss") Then
            strPeriod = Format(rs!��ʼʱ��, "yyyy-MM-dd HH:mm:ss") & "��" & Format(rsTemp!����ʱ��, "yyyy-MM-dd HH:mm") & ":59"
        Else
            strPeriod = Format(rsTemp!��ʼʱ��, "yyyy-MM-dd HH:mm:ss") & "��" & Format(rsTemp!����ʱ��, "yyyy-MM-dd HH:mm") & ":59"
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
    
    '��ȡ�ļ�����
    mblnDateAd = False
    mbln��ʿ = False
    
    Call GetFileProperty
    
    '��ȡ���Ŀ�������ж���(��ʽ���к�;��ͷ����|��Ŀ���,��λ;��Ŀ���,��λ||�к�;��ͷ����...)
    mstrCOLActive = ""
    mstrCOLNothing = ""
    mstrCollectItems = ""
    mstrColCollect = ""
    mstrColCorrelative = ""
    mstrColImCorrelative = ""
    mblnCorrelative = True '��ʼ������Ϊ��(���ݴ���)
    gstrSQL = " Select   A.�к�,A.��ͷ����,A.���,A.��Ŀ���,A.��λ From ���˻�����Ŀ A " & _
              " Where A.�ļ�ID=[1] And A.ҳ��=[2] " & _
              " Order by A.�к�,A.���"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�������Զ���Ļ��Ŀ", mlng�ļ�ID, mintҳ��)
    If rsTemp.RecordCount <> 0 Then
        Do While Not rsTemp.EOF
            If lngCol <> rsTemp!�к� Then
                lngCol = rsTemp!�к�
                mstrCOLActive = mstrCOLActive & "||" & rsTemp!�к� & ";" & rsTemp!��ͷ���� & "|" & rsTemp!��Ŀ��� & "," & NVL(rsTemp!��λ)
            Else
                mstrCOLActive = mstrCOLActive & ";" & rsTemp!��Ŀ��� & "," & NVL(rsTemp!��λ)
            End If
            rsTemp.MoveNext
        Loop
    End If
    If mstrCOLActive <> "" Then mstrCOLActive = Mid(mstrCOLActive, 3)
    
    '��ȡ�����ļ���ʽ����
    gstrSQL = "Select   d.�������, d.�����ı�, d.Ҫ������" & _
        " From �����ļ��ṹ d, �����ļ��ṹ p" & _
        " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '�����ʽ'" & _
        " Order By d.�������"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����ļ���ʽ����", mlng��ʽID)
    With rsTemp
        Do While Not .EOF
            Select Case "" & !Ҫ������
            Case "��ͷ����": mintTabTiers = Val("" & !�����ı�)
            Case "������"
                VsfData.Cols = Val("" & !�����ı�)
                vsfHead.Cols = VsfData.Cols
            Case "��С�и�"
                VsfData.RowHeightMin = BlowUp(Val("" & !�����ı�))
                vsfHead.RowHeightMin = VsfData.RowHeightMin
            Case "�ı�����"
                strCurFont = "" & !�����ı�
                Set objFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = BlowUp(Val(Split(strCurFont, ",")(1)))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "��") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "б") > 0 Then .Italic = True
                End With
                Set VsfData.Font = objFont
                Set vsfHead.Font = objFont
                Set lblSubhead.Font = VsfData.Font
                Set Font = lblSubhead.Font
                Set picMain.Font = Font
                
            Case "�ı���ɫ"
                VsfData.ForeColor = Val("" & !�����ı�)
                vsfHead.ForeColor = VsfData.ForeColor
            Case "�����ɫ"
                VsfData.GridColor = Val("" & !�����ı�): VsfData.GridColorFixed = VsfData.GridColor
                vsfHead.GridColor = VsfData.GridColor: vsfHead.GridColorFixed = VsfData.GridColor
            Case "�����ı�"
                lblTitle.Caption = "" & !�����ı�
                lblTitle.AutoSize = True
            Case "��������"
                strCurFont = "" & !�����ı�
                Set objFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = BlowUp(Val(Split(strCurFont, ",")(1)))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "��") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "б") > 0 Then .Italic = True
                End With
                Set lblTitle.Font = objFont
                lblTitle.AutoSize = False
            
            Case "��ʼʱ��": mintTagFormHour = Val("" & !�����ı�)
            Case "��ֹʱ��": mintTagToHour = Val("" & !�����ı�)
            Case "��������"
                strCurFont = "" & !�����ı�
                Set objFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = BlowUp(Val(Split(strCurFont, ",")(1)))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "��") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "б") > 0 Then .Italic = True
                End With
                Set mobjTagFont = objFont
            Case "������ɫ": mlngTagColor = Val("" & !�����ı�)
            Case "��Ч������"
                mlngOverrunRows = 0
                mlngReduceRow = 0
                mlngPageRows = Val("" & !�����ı�)
            Case "�������"
                mblnCorrelative = (Val("" & !�����ı�) = 1)
            End Select
            .MoveNext
        Loop
    End With
    
    gstrSQL = "Select   ��ʽ, ҳü, ҳ��,���� From ����ҳ���ʽ Where ���� = 3 And ��� In (Select ҳ�� From �����ļ��б� Where Id = [1])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����ҳ���ʽ", mlng��ʽID)
    If Not rsTemp.EOF Then
        mstrPaperSet = "" & rsTemp!��ʽ: mstrPageHead = "" & rsTemp!ҳü: mstrPageFoot = "" & rsTemp!ҳ��
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select   d.�������, d.�����ı�, d.Ҫ������, Nvl(d.�Ƿ���, 0) As �Ƿ���" & _
        " From �����ļ��ṹ d, �����ļ��ṹ p" & _
        " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '���ϱ�ǩ'" & _
        " Order By d.�������"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���ϱ�ǩ����", mlng��ʽID)
    With rsTemp
        mstrSubhead = ""
        Do While Not .EOF
            mstrSubhead = mstrSubhead & "|" & IIf(!�Ƿ��� = 0, "", vbCrLf) & !�����ı� & "{" & !Ҫ������ & "}"
            .MoveNext
        Loop
        If mstrSubhead <> "" Then mstrSubhead = Mid(mstrSubhead, 2)
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select   d.�������, d.�����д�, d.�����ı�" & _
        " From �����ļ��ṹ d, �����ļ��ṹ p" & _
        " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '��ͷ��Ԫ'" & _
        " Order By d.�������"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��ͷ��Ԫ����", mlng��ʽID)
    With rsTemp
        mstrTabHead = ""
        Do While Not .EOF
            mstrTabHead = mstrTabHead & "|" & !�����д� - 1 & "," & !������� & "," & !�����ı�
            .MoveNext
        Loop
        If mstrTabHead <> "" Then mstrTabHead = Mid(mstrTabHead, 2)
    End With
    
    '��ѯ�����֯
    '------------------------------------------------------------------------------------------------------------------
    Dim strSql�� As String, str��ʽ As String, strSqlNull As String
    Dim bln���� As Boolean, blnʱ�� As Boolean, bln��ʿ As Boolean
    Dim blnǩ���� As Boolean, blnǩ��ʱ�� As Boolean, blnǩ������ As Boolean
    Dim bln�Խ��� As Boolean, blnѡ���� As Boolean          '�����һ���ǶԽ�����ѡ����,��ֱ����ȡ��������,ƴ��ͷʱ����ֵ�����/
    Dim lngColumn As Long, blnAddCollect As Boolean
    Dim strColCorrelative  As String
    
    gstrSQL = "Select   d.�������,d.������, d.��������, d.�����д�, d.�����ı�, upper(d.Ҫ������) AS Ҫ������, d.Ҫ�ص�λ,d.Ҫ�ر�ʾ " & _
        " From �����ļ��ṹ d, �����ļ��ṹ p" & _
        " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '���м���'" & _
        " Order By d.�������, d.�����д�"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���м��϶���", mlng��ʽID)
    With rsTemp
        lngColumn = 0: mstrColumns = "": mstrColWidth = "": mstrCatercorner = "": strColCorrelative = ""
        mstrSQL�� = "": mstrSQL�� = "": strSql�� = "": mstrSQL�� = "": mstrSQL���� = "": strSqlNull = ""
        bln���� = False: blnʱ�� = False: bln��ʿ = False
        blnǩ���� = False: blnǩ��ʱ�� = False: blnǩ������ = False
        Do While Not .EOF
            If lngColumn <> !������� Then
                blnAddCollect = False
                If strColCorrelative <> "" Then
                    mstrColCorrelative = mstrColCorrelative & "|" & strColCorrelative
                End If
                strColCorrelative = ""
                mstrColumns = mstrColumns & IIf(mstrColumns = "", "", "'1'" & str��ʽ) & "|" & !������� & "'" & !Ҫ������
                mstrColWidth = mstrColWidth & "," & !�������� & "`" & !������� & "`" & !Ҫ�ر�ʾ
                If !Ҫ�ر�ʾ = 1 Then mstrCatercorner = mstrCatercorner & "," & !�������
                str��ʽ = ""
                If !Ҫ������ <> "" Then str��ʽ = "{" & NVL(!�����ı�) & "[" & !Ҫ������ & "]" & NVL(!Ҫ�ص�λ) & "}"
                
                If Mid(strSqlNull, 3) = "" Then
                    strSqlNull = "''"
                Else
                    strSqlNull = Mid(strSqlNull, 3)
                End If
                mstrSQL�� = mstrSQL�� & "," & IIf(Mid(strSql��, 3) = "", "''", "Decode(" & Mid(strSql��, 3) & "," & strSqlNull & ",''," & Mid(strSql��, 3) & ")") & " As C" & Format(lngColumn, "00")
                
                strSql�� = ""
                strSqlNull = ""
                lngColumn = !�������
                bln�Խ��� = (NVL(!Ҫ�ر�ʾ, 0) = 1)
                blnѡ���� = False
                mrsItems.Filter = "��Ŀ����='" & NVL(!Ҫ������) & "'"
                If mrsItems.RecordCount <> 0 Then
                    blnѡ���� = (mrsItems!��Ŀ��ʾ = 5)
                    If mrsItems!��Ŀ��ʾ = 4 Then   '������Ŀ
                        blnAddCollect = True
                        mstrCollectItems = mstrCollectItems & "," & mrsItems!��Ŀ���
                        mstrColCollect = mstrColCollect & "|" & !������� & ";" & mrsItems!��Ŀ���
                        If Val(NVL(!������)) > 0 And Val(NVL(!�������)) <> Val(NVL(!������)) Then
                            strColCorrelative = Val(NVL(!������)) & ";" & !������� & "," & mrsItems!��Ŀ���
                        End If
                    End If
                End If
                mrsItems.Filter = 0
            Else
                mstrColumns = mstrColumns & "," & !Ҫ������
                str��ʽ = str��ʽ & "{" & NVL(!�����ı�) & "[" & !Ҫ������ & "]" & NVL(!Ҫ�ص�λ) & "}"
                mrsItems.Filter = "��Ŀ����='" & NVL(!Ҫ������) & "'"
                If mrsItems.RecordCount <> 0 Then
                    If mrsItems!��Ŀ��ʾ = 4 Then   '������Ŀ
                        mstrCollectItems = mstrCollectItems & "," & mrsItems!��Ŀ���
                        If blnAddCollect Then
                            strColCorrelative = ""
                            mstrColCollect = mstrColCollect & "," & mrsItems!��Ŀ���
                        Else    '�п���һ�а�������Ŀ,��һ����Ŀ���ǻ�����Ŀ,�ڶ�����Ŀ���ǻ�����Ŀ,���,����Ĵ��뱣֤���������
                            blnAddCollect = True
                            mstrColCollect = mstrColCollect & "|" & !������� & ";" & mrsItems!��Ŀ���
                            If Val(NVL(!������)) > 0 And Val(NVL(!�������)) <> Val(NVL(!������)) Then
                                strColCorrelative = Val(NVL(!������)) & ";" & !������� & "," & mrsItems!��Ŀ���
                            End If
                        End If
                    End If
                End If
                mrsItems.Filter = 0
            End If
            
            Select Case !Ҫ������
            Case "����"
                bln���� = True
                mblnDateAd = (NVL(!Ҫ�ر�ʾ, 0) = 1)
                mstrSQL�� = mstrSQL�� & ",����"
                mstrSQL�� = mstrSQL�� & ",To_Char(l.����ʱ��, " & IIf(mblnDateAd, "'dd/MM'", "'yyyy-mm-dd'") & ") As ����"
                strSql�� = strSql�� & "||" & !Ҫ������
            Case "ʱ��"
                blnʱ�� = True
                mstrSQL�� = mstrSQL�� & ",ʱ��"
                mstrSQL�� = mstrSQL�� & ",To_Char(l.����ʱ��, 'hh24:mi') As ʱ��"
                strSql�� = strSql�� & "||" & !Ҫ������
                
            Case "ǩ����"
                blnǩ���� = True
                mstrSQL�� = mstrSQL�� & ",ǩ����"
                '51589:������,2013-03-01,��ӽ���ǩ��
                'mstrSQL�� = mstrSQL�� & ",l.ǩ����"
                mstrSQL�� = mstrSQL�� & ",DECODE(TRIM(NVL(L.ǩ����,'')),'',TRIM(L.ǩ����),DECODE(TRIM(NVL(L.����ǩ����,'')),'',TRIM(L.ǩ����), TRIM(L.ǩ����) || '/' || TRIM(L.����ǩ����))) ǩ����"
                strSql�� = strSql�� & "||" & !Ҫ������
                
            Case "ǩ��ʱ��"
                blnǩ��ʱ�� = True
                mstrSQL�� = mstrSQL�� & ",ǩ��ʱ��"
                mstrSQL�� = mstrSQL�� & ",l.ǩ��ʱ��"
                strSql�� = strSql�� & "||" & !Ҫ������
                
            Case "��ʿ"
                bln��ʿ = True
                mstrSQL�� = mstrSQL�� & ",��ʿ"
                mstrSQL�� = mstrSQL�� & ",l.������ As ��ʿ"
                strSql�� = strSql�� & "||" & !Ҫ������
            Case Else
                If !Ҫ������ <> "" Then
                    mstrSQL�� = mstrSQL�� & ",Max(""" & !Ҫ������ & """) As """ & !Ҫ������ & """"
                    mstrSQL���� = mstrSQL���� & " Or """ & !Ҫ������ & """ Is Not Null"
                    
                    strSql�� = strSql�� & "||'" & !�����ı� & "'||""" & !Ҫ������ & """||'" & !Ҫ�ص�λ & "'"
                    strSqlNull = strSqlNull & "||" & "'" & !�����ı� & "'||'" & !Ҫ�ص�λ & "'"
                    mstrSQL�� = mstrSQL�� & ", Decode(c.��Ŀ����, '" & !Ҫ������ & "', c.��¼����, '') As """ & !Ҫ������ & """"
                    
'                    If bln�Խ��� And blnѡ���� Then
'                        If strSql�� <> "" Then
'                            '�ڶ���
'                            strSql�� = strSql�� & "||'/'||""" & !Ҫ������ & """"
'                        Else
'                            '��һ��
'                            strSql�� = strSql�� & "||""" & !Ҫ������ & """"
'                        End If
'                    Else
'                        strSql�� = strSql�� & "||""" & !Ҫ������ & """"
'                        strSqlNull = strSqlNull & "||" & "'" & !�����ı� & "'||'" & !Ҫ�ص�λ & "'"
'                    End If
'
'                    If (Trim("" & !�����ı�) = "" And Trim("" & !Ҫ�ص�λ) = "") Or (bln�Խ��� And blnѡ����) Then
'                        mstrSQL�� = mstrSQL�� & ", Decode(c.��Ŀ����, '" & !Ҫ������ & "', Nvl(c.��¼����), '') As """ & !Ҫ������ & """"
'                    Else
'                        mstrSQL�� = mstrSQL�� & ", Decode(c.��Ŀ����, '" & !Ҫ������ & "', Decode(c.��¼����,Null,'" & !�����ı� & "'||'" & !Ҫ�ص�λ & "','" & !�����ı� & "'||c.��¼����||'" & !Ҫ�ص�λ & "'),  '" & !�����ı� & "'||'" & !Ҫ�ص�λ & "') As """ & !Ҫ������ & """"
'                    End If
                Else
                    'Ϊ�ձ�ʾδ����,ǿ�Ƽ�,��������滻
                    mstrCOLNothing = mstrCOLNothing & "," & Val(Format(!�������, "00"))
                    mstrSQL�� = mstrSQL�� & ",Max(""" & "C" & Format(!�������, "00") & """) As C" & Format(!�������, "00")
                    mstrSQL���� = mstrSQL���� & " Or """ & "C" & Format(!�������, "00") & """ Is Not Null"
                    mstrSQL�� = mstrSQL�� & ", C" & Format(!�������, "00") & " AS C" & Format(!�������, "00")
                End If
            End Select
            .MoveNext
        Loop
        
        mbln��ʿ = bln��ʿ
        
        If mstrCollectItems <> "" Then
            mstrCollectItems = Mid(mstrCollectItems, 2)
            mstrColCollect = Mid(mstrColCollect, 2)
        End If
        '��InitRecords����Ҫ��������Ŀ���е��������������Ŀ���
        If Left(mstrColCorrelative, 1) = "|" Then mstrColCorrelative = Mid(mstrColCorrelative, 2)
        mstrCOLNothing = Mid(mstrCOLNothing, 2)
        mstrCatercorner = Mid(mstrCatercorner, 2)
        mstrColWidth = Mid(mstrColWidth, 2)
        '�������һ�еĸ�ʽ
        mstrColumns = mstrColumns & IIf(mstrColumns = "", "", "'1'" & str��ʽ) '& "|" & !������� & "'" & !Ҫ������
        mstrColumns = Mid(mstrColumns, 2)     '��ʽ��:�к�;��Ŀ����1,��Ŀ����2|�к�...,ʵ��;1;����|2;����|3...
        
        If Mid(strSqlNull, 3) = "" Then
            strSqlNull = "''"
        Else
            strSqlNull = Mid(strSqlNull, 3)
        End If
        mstrSQL�� = mstrSQL�� & "," & IIf(Mid(strSql��, 3) = "", "''", "Decode(" & Mid(strSql��, 3) & "," & strSqlNull & ",''," & Mid(strSql��, 3) & ")") & " As C" & Format(lngColumn, "00")
        
        If mstrSQL���� <> "" Then mstrSQL���� = "(" & Mid(mstrSQL����, 5) & ")"
        
        '���û�г������ڣ�ʱ�䣬��ʿ�����ڲ���Ҫ���䣬�Ա�֤�в�����������
        If bln���� = False Then mstrSQL�� = mstrSQL�� & ",To_Char(l.����ʱ��, 'yyyy-mm-dd') As ����"
        If blnʱ�� = False Then mstrSQL�� = mstrSQL�� & ",To_Char(l.����ʱ��, 'hh24:mi') As ʱ��"
        
        '51589:������,2013-03-01,��ӽ���ǩ��
        'If blnǩ���� = False Then mstrSQL�� = mstrSQL�� & ",l.ǩ���� As ǩ����"
        If blnǩ���� = False Then mstrSQL�� = mstrSQL�� & ",DECODE(TRIM(NVL(L.ǩ����,'')),'',TRIM(L.ǩ����),DECODE(TRIM(NVL(L.����ǩ����,'')),'',TRIM(L.ǩ����), TRIM(L.ǩ����) || '/' || TRIM(L.����ǩ����))) ǩ����"
        If blnǩ��ʱ�� = False Then mstrSQL�� = mstrSQL�� & ",l.ǩ��ʱ��"
        
        If Mid(mstrSQL��, 2) = "" Then
            MsgBox "�Բ�����û�ж��嵱ǰ��������ʾ����Ϣ�����ڲ����ļ������ж��壡", vbInformation, gstrSysName
            Exit Function
        End If
        '51589:������,2013-03-01,��ӽ���ǩ��
        '�����ڲ��������ӹ̶���
        mstrSQL�� = UCase(mstrSQL�� & ",MAX(ǩ������) AS ǩ������,MAX(ǩ����Ϣ) AS ǩ����Ϣ,MAX(����ǩ����) AS ����ǩ����,MAX(��¼ID) AS ��¼ID,MAX(����) AS ����,MAX(ʵ������) AS ʵ������,Max(��ʼҳ��) AS ��ʼҳ��,Max(��ʼ�к�) AS ��ʼ�к�,MAX(�������) AS �������,MAX(�����ı�) AS �����ı�,MAX(���ܱ��) AS ���ܱ��,MAX(��������) AS ��������,MAX(��ʼʱ��) AS ��ʼʱ��,MAX(����ʱ��) AS ����ʱ��")
        mstrSQL�� = UCase(mstrSQL�� & ",l.ǩ������,l.ǩ���� AS ǩ����Ϣ,l.����ǩ����,C.��¼ID,P.����||'' AS ����,DECODE(SIGN(P.����ҳ��-P.��ʼҳ��),1,DECODE(SIGN([5]-P.��ʼҳ��),1, P.�����к�,P.����-P.�����к� ),P.����) AS ʵ������,P.��ʼҳ��,P.��ʼ�к�,NVL(L.�������,0) AS �������,L.�����ı�,L.���ܱ��,to_char(L.����ʱ��,'yyyy-MM-dd hh24:mi:ss')||'' AS ��������,L.��ʼʱ��,L.����ʱ��")
        mstrSQL�� = UCase(mstrSQL�� & ",ǩ������,ǩ����Ϣ,����ǩ����,��¼ID,����,ʵ������,��ʼҳ��,��ʼ�к�,�������,�����ı�,���ܱ��,��������,��ʼʱ��,����ʱ��")
        
        '63706:������,2013-11-20,ǿ�ư󶨻�ʿ��
'        If bln��ʿ = False Then
        'ǿ����ӻ�ʿ��,Ϊ�˱����޸�����������(����¼�������,����������Ҳ������)
        mstrSQL�� = mstrSQL�� & ",��ʿL"
        mstrSQL�� = mstrSQL�� & ",l.������ As ��ʿL"
        mstrSQL�� = mstrSQL�� & ",��ʿL"
'        End If
        
        '�����Ŀ���뵽SQL��
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
    '���±�ͷ
    
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
    '�����Ŀ���뵽��ѯSQL�У���ʽ���к�;��ͷ����|��Ŀ���,��λ;��Ŀ���,��λ||�к�;��ͷ����...
    '�󶨶����Ŀ�����о��Զ�תΪ�Խ�����
    
    arrData = Split(mstrCOLActive, "||")
    intCount = UBound(arrData)
    For intDo = 0 To intCount
        intCol = Split(Split(arrData(intDo), "|")(0), ";")(0)
        strName = Split(Split(arrData(intDo), "|")(0), ";")(1)
        arrCol = Split(Split(arrData(intDo), "|")(1), ";")
        
        '�����б�ʾ(ÿ������������Ŀ)
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
            mrsItems.Filter = "��Ŀ���=" & Val(Split(arrCol(intIn), ",")(0))
            strCOLNames = strCOLNames & "," & mrsItems!��Ŀ����
            strCOLCOND = strCOLCOND & " OR """ & strCOLPart & mrsItems!��Ŀ���� & """ IS NOT NULL"
            strCOLMID = strCOLMID & ",Max(""" & strCOLPart & mrsItems!��Ŀ���� & """) As """ & strCOLPart & mrsItems!��Ŀ���� & """"
            If intIn = 0 Then
                strCOLIN = strCOLIN & ", Decode(" & IIf(strCOLPart = "", "", "c.���²�λ||") & "c.��Ŀ����, '" & strCOLPart & mrsItems!��Ŀ���� & "',c.��¼����, '') As """ & strCOLPart & mrsItems!��Ŀ���� & """"
            Else
                strCOLIN = strCOLIN & ", Decode(" & IIf(strCOLPart = "", "", "c.���²�λ||") & "c.��Ŀ����, '" & strCOLPart & mrsItems!��Ŀ���� & "', Decode(c.��¼����,Null,'/','/'||c.��¼����||''), '') As """ & strCOLPart & mrsItems!��Ŀ���� & """"
            End If
            If intIn = 0 Then
                If intMax = 0 Then
                    strCOLDEF = strCOLDEF & " """ & strCOLPart & mrsItems!��Ŀ���� & """ AS C" & Format(intCol, "00")
                Else
                    strCOLDEF = strCOLDEF & " """ & strCOLPart & mrsItems!��Ŀ���� & """||"
                End If
            Else
                strCOLDEF = strCOLDEF & "NVL(""" & strCOLPart & mrsItems!��Ŀ���� & """,'/')"
                If intIn = intMax Then
                    strCOLDEF = "Decode(" & strCOLDEF & ",'" & String(intMax, "/") & "',''," & strCOLDEF & ") As C" & Format(intCol, "00")
                End If
            End If
            
            strColFormat = strColFormat & "{[" & strCOLPart & mrsItems!��Ŀ���� & "]" & IIf(intMax > 0 And intIn < intMax, "/", "") & "}"
        Next
        If strCOLPart <> "" Then
            strCOLPart = Mid(strCOLPart, 2)
        End If
        strCOLNames = Mid(strCOLNames, 2)
        
        '�Խ���
        If intMax > 0 Then
            mstrCatercorner = mstrCatercorner & IIf(mstrCatercorner = "", "", ",") & intCol
        End If
        '�и�ʽ:15'��ʿ'1'{[��ʿ]}
        '77476:LPF:����滻intcolǰ���"|"�ַ�,�����3�к͵�13�ж�Ϊ���Ŀʱ��Ŀ�滻����
        mstrColumns = Replace(mstrColumns, "|" & intCol & "''1'", "|" & intCol & "'" & strCOLNames & "'1'" & strColFormat)
        '��
        mstrSQL�� = Replace(mstrSQL��, "'' AS C" & Format(intCol, "00"), strCOLDEF)
        '����
        '53893:������,2012-09-21,������Ŀ����ʱ���������
        'mstrSQL���� = Replace(UCase(mstrSQL����), " OR """ & "C" & Format(intCol, "00") & """ IS NOT NULL", strCOLCOND)
        mstrSQL���� = Replace(UCase(Replace(UCase(mstrSQL����), " OR """ & "C" & Format(intCol, "00") & """ IS NOT NULL", strCOLCOND)), """" & "C" & Format(intCol, "00") & """ IS NOT NULL", Mid(strCOLCOND, 5))
        '��
        mstrSQL�� = Replace(mstrSQL��, ",MAX(""" & "C" & Format(intCol, "00") & """) AS C" & Format(intCol, "00"), strCOLMID)
        '��
        mstrSQL�� = Replace(mstrSQL��, ", C" & Format(intCol, "00") & " AS C" & Format(intCol, "00"), strCOLIN)
    Next
    mrsItems.Filter = 0
    
    '��δ�󶨵��е�SQL�������
    If mstrCOLNothing = "" Then Exit Sub
    arrData = Split(mstrCOLNothing, ",")
    intCount = UBound(arrData)
    For intDo = 0 To intCount
        '��(����Ҫ����)
'        mstrSQL�� = Replace(mstrSQL��, ",'' AS C" & arrData(intDo), "")
        '����
        'mstrSQL���� = Replace(UCase(mstrSQL����), " OR """ & "C" & Format(arrData(intDo), "00") & """ IS NOT NULL", "")
        mstrSQL���� = Replace(UCase(Replace(UCase(mstrSQL����), " OR """ & "C" & Format(arrData(intDo), "00") & """ IS NOT NULL", "")), """" & "C" & Format(arrData(intDo), "00") & """ IS NOT NULL OR ", "")
        mstrSQL���� = Replace(UCase(mstrSQL����), "(""" & "C" & Format(arrData(intDo), "00") & """ IS NOT NULL)", "")

        '��
        mstrSQL�� = Replace(mstrSQL��, ",MAX(""" & "C" & Format(arrData(intDo), "00") & """) AS C" & Format(arrData(intDo), "00"), "")
        '��
        mstrSQL�� = Replace(mstrSQL��, ", C" & Format(arrData(intDo), "00") & " AS C" & Format(arrData(intDo), "00"), "")
    Next
End Sub

Private Sub SQLCombination(Optional ByVal lng��¼ID As Long = 0)
    Dim str���� As String
    str���� = mstrSQL���� & IIf(lng��¼ID = 0, "", IIf(mstrSQL���� = "", "", " And") & " ��¼ID=[6]")
    
    mstrSQL = "Select '' AS ����,to_char(����ʱ��,'yyyy-MM-dd hh24:mi:ss') AS ����ʱ��,'' AS ѡ��,to_char(����ʱ��,'YYYY') AS ���," & Mid(mstrSQL��, 12) & vbCrLf & _
                " From (Select ��¼���,ʱ�� as ����,����ʱ��," & Mid(mstrSQL��, 2) & vbCrLf & _
                "        From (Select nvl(c.��¼���,0) ��¼���,l.����ʱ��," & Mid(mstrSQL��, 2) & vbCrLf & _
                "               From ���˻������� l, ���˻�����ϸ c,���˻����ļ� f,���˻����ӡ p " & vbCrLf & _
                "               Where l.ID=p.��¼ID And l.Id = c.��¼id And l.�ļ�ID+0=f.ID+0 And f.ID=p.�ļ�ID " & _
                "               And c.��ֹ�汾 Is Null And c.��¼����<>5  " & _
                "               And f.id=[1] And f.����id = [2] And f.��ҳid = [3] And Nvl(f.Ӥ��,0)=[4] " & _
                IIf(mintPageSpan = 0, " And (P.��ʼҳ��=[5] Or P.����ҳ��=[5])", " And P.��ʼҳ��=[5]") & ")" & vbCrLf & _
                IIf(str���� <> "", "Where " & str����, "") & _
                "       Group By ����, ʱ��, ����ʱ��,��¼���," & IIf(mbln��ʿ = True, "��ʿ,", "��ʿL,") & "ǩ����,ǩ��ʱ��" & _
                                "       Order By ����ʱ��,��¼���," & IIf(mbln��ʿ = True, "��ʿ,", "��ʿL,") & "ǩ����,ǩ��ʱ��)"
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
    '���ϱ�ǩ��ȡ
    lblSubhead.Caption = ""
    lblSubhead.Tag = ""
    aryPeriod = Split(GetPeriod, "��")
    '87057,����10:30:20ת����ס���������ļ�ʱ��Ϊ10:30:20,��ʱ¼����������Ϊ10:30(��¼���޷�¼����),�����޷���ʾ�µĿ���
    aryPeriod(0) = Format(aryPeriod(0), "YYYY-MM-DD HH:mm") & ":59"
    '��ȡ��ǰҳ֮ǰ��������ID
    gstrSQL = "Select ����ID From ���˱䶯��¼ " & _
        "   Where  ����ID=[1] And ��ҳID=[2] And [3]>=��ʼʱ�� " & _
        " And ��ʼʱ�� IS NOT NULL And ����id IS NOT NULL And NVL(���Ӵ�λ,0)=0 Order by ��ʼʱ�� DESC"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��ǰҳ֮ǰ��������ID", mlng����ID, mlng��ҳID, CDate(aryPeriod(0)))
    If rsTemp.RecordCount > 0 Then mlng����ID = Val(rsTemp!����ID)
    
    gstrSQL = "Select [1] || Zl_Replace_Element_Value([2],[3],[4],2,NULL,[5],[6]) as ��Ϣ From Dual"
    aryItem = Split(mstrSubhead, "|")
    
        
    For lngCount = 0 To UBound(aryItem)
        strPrefix = Left(aryItem(lngCount), InStr(1, aryItem(lngCount), "{") - 1)
        strItemName = Mid(aryItem(lngCount), InStr(1, aryItem(lngCount), "{") + 1, InStr(1, aryItem(lngCount), "}") - InStr(1, aryItem(lngCount), "{") - 1)
        
        strtmp = strPrefix
        strCell = ""
        '68336
        blnReplace = True
        mrsElement.Filter = "������='" & strItemName & "'"
        If mrsElement.RecordCount > 0 Then
            blnReplace = Val(NVL(mrsElement!�滻��, 0)) = 1
        End If
        Select Case strItemName
        Case "��ǰ����"
        
            strTmpSQL = "Select   b.����" & vbNewLine & _
                        "From (Select ����id, ��ʼʱ��, Nvl(��ֹʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) As ��ֹʱ��" & vbNewLine & _
                        "            From ���˱䶯��¼" & vbNewLine & _
                        "            Where ����id = [1] And ��ҳid = [2] And ����id = [3] And NVL(���Ӵ�λ,0)=0 And ��ʼʱ�� IS NOT NULL) a,���ű� b " & vbNewLine & _
                        "Where ([4] Between a.��ʼʱ�� And a.��ֹʱ�� Or [4] >= a.��ʼʱ��) And a.����id Is Not Null And b.ID=a.����id" & vbNewLine & _
                        "Order By a.��ʼʱ��"
                        
            Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "��ǰ����", mlng����ID, mlng��ҳID, mlng����ID, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
            If rsTemp.BOF = False Then rsTemp.MoveLast
        Case "��ǰ����"

            strTmpSQL = "Select   a.����" & vbNewLine & _
                        "From (Select ����, ��ʼʱ��, Nvl(��ֹʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) As ��ֹʱ��" & vbNewLine & _
                        "            From ���˱䶯��¼" & vbNewLine & _
                        "            Where ����id = [1] And ��ҳid = [2] And ����id = [3] And NVL(���Ӵ�λ,0)=0 And ��ʼʱ�� IS NOT NULL) a" & vbNewLine & _
                        "Where ([4] Between a.��ʼʱ�� And a.��ֹʱ�� Or [4] >= a.��ʼʱ��) And a.���� Is Not Null" & vbNewLine & _
                        "Order By a.��ʼʱ��"

            Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "��ǰ����", mlng����ID, mlng��ҳID, mlng����ID, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
            If rsTemp.BOF = False Then rsTemp.MoveLast
        Case "��λ�䶯"
            strTmpSQL = "Select   a.����" & vbNewLine & _
                        "From (Select ����, ��ʼʱ��, Nvl(��ֹʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) As ��ֹʱ��" & vbNewLine & _
                        "            From ���˱䶯��¼" & vbNewLine & _
                        "            Where ����id = [1] And ��ҳid = [2] And ����id = [3] And NVL(���Ӵ�λ,0)=0 And ��ʼʱ�� IS NOT NULL) a" & vbNewLine & _
                        "Where (a.��ֹʱ��>=[4] And a.��ʼʱ��<=[5]) And a.���� Is Not Null" & vbNewLine & _
                        "Order By a.��ʼʱ��"

            Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "��ǰ����", mlng����ID, mlng��ҳID, mlng����ID, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
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
        Case "��ǰ����"
        
            strTmpSQL = "Select   ���� From ���ű� a Where a.ID=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "��ǰ����", mlng����ID)
            
        Case "סԺҽʦ"
            strTmpSQL = "Select   a.����ҽʦ" & vbNewLine & _
                        "From (Select ����ҽʦ, ��ʼʱ��, Nvl(��ֹʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) As ��ֹʱ��" & vbNewLine & _
                        "            From ���˱䶯��¼" & vbNewLine & _
                        "            Where ����id = [1] And ��ҳid = [2] And ����id = [3] And NVL(���Ӵ�λ,0)=0 And ��ʼʱ�� IS NOT NULL) a" & vbNewLine & _
                        "Where ([4] Between a.��ʼʱ�� And a.��ֹʱ�� Or [4] >= a.��ʼʱ��) And a.����ҽʦ Is Not Null" & vbNewLine & _
                        "Order By a.��ʼʱ��"
            Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "סԺҽʦ", mlng����ID, mlng��ҳID, mlng����ID, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
            If rsTemp.BOF = False Then rsTemp.MoveLast
        Case "���λ�ʿ"
        
            strTmpSQL = "Select   a.���λ�ʿ" & vbNewLine & _
                        "From (Select ���λ�ʿ, ��ʼʱ��, Nvl(��ֹʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) As ��ֹʱ��" & vbNewLine & _
                        "            From ���˱䶯��¼" & vbNewLine & _
                        "            Where ����id = [1] And ��ҳid = [2] And ����id = [3] And NVL(���Ӵ�λ,0)=0 And ��ʼʱ�� IS NOT NULL) a" & vbNewLine & _
                        "Where ([4] Between a.��ʼʱ�� And a.��ֹʱ�� Or [4] >= a.��ʼʱ��) And a.���λ�ʿ Is Not Null" & vbNewLine & _
                        "Order By a.��ʼʱ��"
            Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "���λ�ʿ", mlng����ID, mlng��ҳID, mlng����ID, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
            If rsTemp.BOF = False Then rsTemp.MoveLast
            
        Case "����ȼ�"
            strTmpSQL = "Select   b.����" & vbNewLine & _
                        "From (Select ����ȼ�ID, ��ʼʱ��, Nvl(��ֹʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) As ��ֹʱ��" & vbNewLine & _
                        "            From ���˱䶯��¼" & vbNewLine & _
                        "            Where ����id = [1] And ��ҳid = [2] And ����id = [3] And NVL(���Ӵ�λ,0)=0 And ��ʼʱ�� IS NOT NULL) a,����ȼ� b" & vbNewLine & _
                        "Where ([4] Between a.��ʼʱ�� And a.��ֹʱ�� Or [4] >= a.��ʼʱ��) And a.����ȼ�ID Is Not Null And b.���=a.����ȼ�ID" & vbNewLine & _
                        "Order By a.��ʼʱ��"
            Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "����ȼ�", mlng����ID, mlng��ҳID, mlng����ID, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
            If rsTemp.BOF = False Then rsTemp.MoveLast
        Case "������"
            strtmp = ""
            gstrSQL = "Select [1] || Zl_Replace_Element_Value([2],[3],[4],2,NULL,[5],[6]) as ��Ϣ From Dual"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡҪ��", strPrefix, strItemName, mlng����ID, mlng��ҳID, mintӤ��, CDate(aryPeriod(0)))
        Case Else
            strtmp = ""
            If blnReplace = True Then
                gstrSQL = "Select [1] || Zl_Replace_Element_Value([2],[3],[4],2,NULL,[5],[6]) as ��Ϣ From Dual"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡҪ��", strPrefix, strItemName, mlng����ID, mlng��ҳID, mintӤ��, CDate(aryPeriod(0)))
            Else
                mblnElement = True
                strtmp = strPrefix
                gstrSQL = "Select ���� From ���˻���Ҫ������ Where �ļ�ID=[1] And ҳ��=[2] And ����=[3]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡҪ��", mlng�ļ�ID, mintҳ��, strItemName)
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
    
    '���ϱ�ǩ��ɢ����
    Call zlLableBruit
    
    '�����м�¼��
    Call InitRecords
    
    'װ������
    Call SQLCombination
    gstrSQL = mstrSQL
    Call SQLDIY(gstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��������", mlng�ļ�ID, mlng����ID, mlng��ҳID, mintӤ��, mintҳ��)
    '�����������¼���ṹ
    Call DataMap_Init(rsTemp)
    '�����ݲ����û����¼���ĸ�ʽ,ͬʱʵ��һ�����ݷ�����ʾ�Ĺ���
    Call PreTendFormat(rsTemp)
    Call cbsThis_Resize
    
    lblCurPage.Caption = "P" & mintҳ��
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub DataMap_Init(ByVal rsSource As ADODB.Recordset)
    '��ʼ���ڴ����ݼ�
    
    If Not mblnClear Then Exit Sub
    
    '���ݼ�¼��,���ڿ��ٻָ�
    Set mrsDataMap = CopyNewRec(rsSource)
    mrsDataMap.Sort = "ҳ��,�к�"
    '�޸ĵ�Ԫ���¼,���ڱ���(�����Ҫ���ڱ��浼��������ҽ����Ϣ:ҽ��ID:���ͺ� )
    Call Record_Init(mrsCellMap, "ID," & adLongVarChar & ",50|ҳ��," & adDouble & ",18|�к�," & adDouble & ",18|" & _
            "�к�," & adDouble & ",18|��ʼ�к�," & adDouble & ",18|��¼ID," & adDouble & ",18|����," & adLongVarChar & ",4000|��λ," & adLongVarChar & ",100|" & _
            "���," & adLongVarChar & ",100|����," & adDouble & ",1|��¼���," & adDouble & ",1|ɾ��," & adDouble & ",1")
    mrsCellMap.Sort = "ҳ��,�к�,�к�"
    '���Ƽ�¼��
    Set mrsCopyMap = New ADODB.Recordset
    Set mrsCopyMap = CopyNewRec(mrsDataMap, False)
    
    'Ϊ�˲�Ӱ��֮��Ļ�ҳ,���˲�������Ϊ��
    mblnClear = False
End Sub

Private Function DataMap_Save() As Boolean
    '����ǰҳ�����û��༭�������ݱ�������,ҳ���л��򱣴�ǰ����
    Dim blnExit As Boolean
    Dim lngRow As Long, lngCol As Long, lngRows As Long, lngCols As Long
    On Error GoTo ErrHand
    
    '�����Ƿ�༭��������
'    '�����ǰҳδ�༭��,�򲻱ر���
'    mrsCellMap.Filter = "ҳ��=" & mintҳ��
'    blnExit = (mrsCellMap.RecordCount = 0)
'    If blnExit Then
'        mrsCellMap.Filter = 0
'        DataMap_Save = True
'        Exit Function
'    End If
'    mrsCellMap.Filter = 0
    
    If Not CheckFlip Then Exit Function
    
    '��ɾ��ָ��ҳ�ŵ�����������
    mrsDataMap.Filter = "ҳ��=" & mintҳ��
    Do While True
        If mrsDataMap.RecordCount = 0 Then Exit Do
        mrsDataMap.Delete
        mrsDataMap.MoveNext
    Loop
    mrsDataMap.Filter = 0
    
    '����ָ��ҳ�ŵ�����������
    lngRows = VsfData.Rows - 1
    lngCols = VsfData.Cols - 1
    For lngRow = VsfData.FixedRows To lngRows
        mrsDataMap.AddNew
        mrsDataMap!ҳ�� = mintҳ��
        mrsDataMap!�к� = lngRow
        mrsDataMap!ɾ�� = IIf(VsfData.RowHidden(lngRow), 1, 0)
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
        
        mrsCellMap.Filter = "ҳ��=" & mintҳ�� & " And �к�=" & lngRow
        If mrsCellMap.RecordCount > 0 And VsfData.RowHidden(lngRow) = True Then
            Do While Not mrsCellMap.EOF
                If mrsCellMap!�к� > mlngTime And mrsCellMap!�к� < mlngNoEditor Then
                    mrsCellMap!���� = ""
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
    '��ָ��ҳ������ݻָ��������
    Dim lngRow As Long, lngCol As Long, lngRows As Long, lngCols As Long
    On Error GoTo ErrHand
    
    mblnRestore = False
    If VsfData.Rows > VsfData.FixedRows Then
        VsfData.Cell(flexcpChecked, VsfData.FixedRows, mlngChoose, VsfData.Rows - 1, mlngChoose) = flexTSUnchecked
    End If
    '����ָ��ҳ�ŵ����������е������
    mrsDataMap.Filter = "ҳ��=" & mintҳ��
    lngRows = mrsDataMap.RecordCount
    
    If lngRows = 0 Then
        'û���޸Ĺ���������󶨶�ȡ�ļ�¼��
        mrsDataMap.Filter = 0
        Set VsfData.DataSource = rsTemp
        DataMap_Restore = True
        Exit Function
    Else
        '�˴�ֻ��Ҫ��һ���յļ�¼������(�ָ�����)
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
        If mrsDataMap!ɾ�� = 1 Then VsfData.RowHidden(VsfData.FixedRows + lngRow) = True
        
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
    
    '���µ�ǰҳ�����д�����ʼ�е��к�����
    With mrsCellMap
        If lngDeff > 0 Then
            If .RecordCount = 0 Then Exit Sub
            If .RecordCount <> 0 Then .MoveLast
            If .BOF Then Exit Sub
            Do While Not mrsCellMap.BOF
                If !ҳ�� = mintҳ�� And IIf(blnBig = True, !�к� > lngStart, !�к� = lngStart) Then
                    intCol = !�к�
                    lngPos = .AbsolutePosition
                    !�к� = !�к� + lngDeff
                    !ID = mintҳ�� & "," & !�к� & "," & !�к�
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
                If !ҳ�� = mintҳ�� And IIf(blnBig = True, !�к� > lngStart, !�к� = lngStart) Then
                    intCol = !�к�
                    lngPos = .AbsolutePosition
                    !�к� = !�к� + lngDeff
                    !ID = mintҳ�� & "," & !�к� & "," & !�к�
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
    'ֻ������¼���Ľṹ,ͬʱ����ҳ��,�к��ֶ�
    Dim rsTarget As New ADODB.Recordset
    Dim intFields As Integer
    
    Set rsTarget = New ADODB.Recordset
    With rsTarget
        If blnAddPage Then
            .Fields.Append "ҳ��", adDouble, 18
            .Fields.Append "�к�", adDouble, 18
        End If
        For intFields = 0 To rsSource.Fields.Count - 1
            If rsSource.Fields(intFields).Name = "��������" Then
                .Fields.Append rsSource.Fields(intFields).Name, adLongVarChar, 50, adFldIsNullable      '0:��ʾ����
            ElseIf rsSource.Fields(intFields).Type = 200 Then       '�����ʹ���Ϊ�ַ���
                .Fields.Append rsSource.Fields(intFields).Name, adLongVarChar, rsSource.Fields(intFields).DefinedSize, adFldIsNullable     '0:��ʾ����
            Else
                .Fields.Append rsSource.Fields(intFields).Name, IIf(rsSource.Fields(intFields).Type = adNumeric, adDouble, rsSource.Fields(intFields).Type), rsSource.Fields(intFields).DefinedSize, adFldIsNullable    '0:��ʾ����
            End If
        Next
        If blnAddPage Then
            .Fields.Append "ɾ��", adDouble, 1
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
    Dim lngRowCount As Long, lngRowCurrent As Long  '��ǰ��¼������,��ǰ��¼�ڱ�ҳ��ʵ������
    Dim lngCol As Long, lngMax As Long, lngRecordId As Long
    Dim lngRow As Long
    Dim str����ʱ�� As String, str����ʱ��_L As String, lngLastRow As Long
    Dim lngLiterrRow As Long
    Dim lngTestRow As Long, lngStartRow As Long
    Dim strDate As String
    Dim intCol As Integer, intCols As Integer
    Dim rsData As New ADODB.Recordset
    Dim strSignName As String
    Dim lngPrintedRow As Long, lngStart As Long
    Dim blnClear As Boolean
    Dim lngCount As Long
    Dim blnCollectType As Boolean  '��¼���������е���һ���Ƿ��ǻ�����
    Dim lngCurrRow As Long, lngCollectMutilRows As Long '�������ݵ�ǰ�С����������ݵ�����
    Dim i As Integer, j As Integer, arrItem, arrCorrelative, arrLastRow, arrMutilRows '���������Ŀ����
    On Error GoTo ErrHand
    
    arrItem = Split(mstrColCorrelative, "|")
    '���һ����ʾ�����������ʾ(���ݵ�ǰ����ռ����������ӿհ��в�����������,Ȼ�������δ���ǰ�е�����)
    'ÿҳֻ��ʾʵ�ʵ�������,��'@��ȡ��ע�ͼ���
    lngRow = VsfData.FixedRows
    Do While True
        If lngRow > VsfData.Rows - 1 Then Exit Do
        If InStr(1, FormatValue(VsfData.TextMatrix(lngRow, mlngRowCount)), "|") <> 0 Then Exit Do
        
        lngRowCount = Val(FormatValue(VsfData.TextMatrix(lngRow, mlngRowCount)))
        '@ʵ��������
'        lngRowCurrent = Val(FormatValueVsfData.TextMatrix(lngRow, mlngRowCurrent)))
        
        str����ʱ�� = Format(VsfData.TextMatrix(lngRow, mlngActTime), "YYYY-MM-DD HH:mm:ss")
        If Val(VsfData.TextMatrix(lngRow, mlngCollectType)) < 0 Then
            '�ڴ���������������¼ʱ(����)��������������еĸ�ֵ blnCollectType=false�����
            If blnCollectType = False Then str����ʱ��_L = "": blnCollectType = True
            '�������������ϸ���ݵĴ���(���ݱ��淽ʽ��һ���������ݶ�Ӧ������ϸ,��ϸ�еļ�¼��Ų�ͬ)
            If str����ʱ��_L <> "" And str����ʱ��_L = str����ʱ�� Then
                If UBound(arrItem) < 0 Then '�����ǰû�����û��ܹ�ϵ,��֮ǰ���ݴ��ڷ�����ܵ���������ӷ�������ѭ������
                    lngCurrRow = lngLastRow + lngCollectMutilRows 'ȷ��ÿһ���������������ʼλ��
                    lngCollectMutilRows = 1
                    If lngCurrRow < lngRow Then
                        VsfData.TextMatrix(lngCurrRow, mlngYear) = ""
                        VsfData.TextMatrix(lngCurrRow, mlngDate) = ""
                        VsfData.TextMatrix(lngCurrRow, mlngTime) = ""
                        
                        For lngCol = mlngTime + 1 To mlngNoEditor - 1
                            If (lngCol <> mlngSignTime And VsfData.ColHidden(lngCol) = False) Then
                                '׼����ֵ
                                With txtLength
                                    .Width = VsfData.ColWidth(lngCol)
                                    '������Ҫע��һ�㣺��ȡ�������ݵ�����Ӧ����lngRow������lngCurrRow����Ϊ�ڴ������������¼ʱ�ᵼ���������ݵ���λ�÷����仯 (���� = ����¼��ʼ�к� + ��������)
                                    .Text = Replace(Replace(Replace(VsfData.TextMatrix(lngRow, lngCol), Chr(10), ""), Chr(13), ""), Chr(1), "")
                                    .FontName = VsfData.CellFontName
                                    .FontSize = VsfData.CellFontSize
                                    .FontBold = VsfData.CellFontBold
                                    .FontItalic = VsfData.CellFontItalic
                                End With
                                arrData = GetData(txtLength.Text)
                                intDatas = UBound(arrData)
                                
                                If intDatas >= 0 Then
                                    'ѭ����ֵ
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
                    '�����˷�����ܹ�ϵ������ÿ��������Ŀ����չʾ����
                    For i = 0 To UBound(arrItem)
                        lngCurrRow = Val(arrLastRow(i)) + Val(arrMutilRows(i)) '����Ŀ����,ȷ��ÿ�������������ʼλ��
                        lngCollectMutilRows = 1
                        arrMutilRows(i) = lngCollectMutilRows
                        If lngCurrRow < lngRow Then
                            arrCorrelative = Split(arrItem(i), ";")
                            For j = 0 To 1
                                '׼����ֵ
                                    lngCol = Split(arrCorrelative(j), ",")(0) + cHideCols + VsfData.FixedCols - 1
                                    With txtLength
                                        .Width = VsfData.ColWidth(lngCol)
                                        '������Ҫע��һ�㣺��ȡ�������ݵ�����Ӧ����lngRow������lngCurrRow����Ϊ�ڴ������������¼ʱ�ᵼ���������ݵ���λ�÷����仯 (���� = ����¼��ʼ�к� + ��������)
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
                '��ֵ��ɺ��Ƴ�ԭ��������
                VsfData.RowPosition(lngRow) = VsfData.Rows - 1
                VsfData.RemoveItem VsfData.Rows - 1
                GoTo NextData
            Else
                'If lngRow >= mlngPageRows + mlngOverrunRows - mlngReduceRow + VsfData.FixedRows Then Exit Do
                '������Ĭ��Ϊһ��(ֻ����Ի����е�����)
                lngCollectMutilRows = 1
                lngLastRow = lngRow '��¼������������е�λ��
                'ȷ��������������ӷ�������ÿ��������Ŀ����ʼλ��
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
            If blnCollectType = True Then str����ʱ��_L = "": blnCollectType = False
            If str����ʱ��_L <> "" And Mid(str����ʱ��_L, 1, 16) = Mid(str����ʱ��, 1, 16) And str����ʱ��_L <> str����ʱ�� Then
                '������ͬ��������ͬ���Ҳ��ǻ��������У���˵����Щ������һ�飬����lngDemo��
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
            '�����ӿ���
            VsfData.Rows = VsfData.Rows + lngRowCount - 1
            '�ӵ�ǰ�е���һ�п�ʼ��ÿ�е�λ��+�����ӵĿհ���������֤�����Ŀհ��дӵ�ǰ�е���һ�п�ʼ
            For intData = VsfData.Rows - lngRowCount To lngRow + 1 Step -1
                VsfData.RowPosition(intData) = intData + lngRowCount - 1
            Next
            
            'ѭ������ǰ������
            For lngCol = 0 To VsfData.Cols - 1
                If VsfData.ColHidden(lngCol) And lngCol <> mlngRowCount And lngCol <> mlngDemo Then
                    'ѭ����ֵ
                    For intData = 2 To lngRowCount
                        VsfData.TextMatrix(lngRow + intData - 1, lngCol) = FormatValue(VsfData.TextMatrix(lngRow, lngCol))
                    Next
                ElseIf (lngCol < mlngNoEditor And lngCol <> mlngDate And lngCol <> mlngTime And lngCol <> mlngYear) Then
                    '׼����ֵ
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
                        'ѭ����ֵ
                        If intDatas + 1 > lngRowCount Then intDatas = lngRowCount - 1
                        For intData = 0 To intDatas
                            If VsfData.Rows <= lngRow + intData Then VsfData.Rows = VsfData.Rows + 1
                            VsfData.TextMatrix(lngRow + intData, lngCol) = Replace(Replace(Replace(arrData(intData), Chr(10), ""), Chr(13), ""), Chr(1), "")
                        Next
                    End If
                ElseIf lngCol = mlngNoEditor Then
                    '����ֵ��Ϊ��1��ʼ,������4������,����4|1
                    For intData = 1 To lngRowCount
                        VsfData.TextMatrix(lngRow + intData - 1, mlngRowCount) = lngRowCount & "|" & intData
                    Next
                    '���һ����Ҫ��д���ǩ��
                    If mlngSignName > 0 Then VsfData.TextMatrix(lngRow + lngRowCount - 1, mlngSignName) = FormatValue(VsfData.TextMatrix(lngRow, mlngSignName))
                    If mlngSignTime > 0 Then VsfData.TextMatrix(lngRow + lngRowCount - 1, mlngSignTime) = FormatValue(VsfData.TextMatrix(lngRow, mlngSignTime))
                    '--58414,������,2013-04-10,��ӻ�ʿ��ǩ������ʾģʽ
                    Call SingerShowType(VsfData, lngRow, lngRow + lngRowCount - 1)
                Else
                    
                End If
            Next
            '@ʵ��������
'            '�����ҳ��һ�е����ݲ�ȫ,���Ƚ��ü�¼��һ�е�������(����,ʱ��,ǩ��)��Ϣ���Ƶ�
'            If lngRow = VsfData.FixedRows And lngRowCount <> lngRowCurrent Then
'                '�̶�������ʾ����ʱ����ǩ����
'                lngMax = lngRowCount - lngRowCurrent
'                If mlngDate > -1 Then VsfData.TextMatrix(lngRow + lngMax, mlngDate) = VsfData.TextMatrix(lngRow, mlngDate)
'                If mlngTime > -1 Then VsfData.TextMatrix(lngRow + lngMax, mlngTime) = VsfData.TextMatrix(lngRow, mlngTime)
'                if mlngOperator <>-1 then VsfData.TextMatrix(lngRow + lngMax, mlngOperator) = VsfData.TextMatrix(lngRow, mlngOperator)
'                if mlngSignName <>-1 then VsfData.TextMatrix(lngRow + lngMax, mlngsignname) = VsfData.TextMatrix(lngRow, mlngsignname)
'                'ɾ���������
'                For lngCol = 1 To lngMax
'                    VsfData.RemoveItem lngRow
'                Next
'            End If
'            lngRow = lngRow + lngRowCurrent - 1 '���ϸü�¼�ڱ�ҳʵ�ʵ�����
            '@ʵ��������Ҫע���������д���
            lngRow = lngRow + lngRowCount - 1
        Else
            VsfData.TextMatrix(lngRow, mlngRowCount) = "1|1"
            VsfData.TextMatrix(lngRow, mlngRowCurrent) = "1"
        End If
        lngRow = lngRow + 1
NextData:
        str����ʱ��_L = str����ʱ��
    Loop
    If mblnRestore Then Exit Sub
    
    'Modified by zyb 2011-09-15
    'Modify by LPF 2012-05-08
    '��������,ֻ������һҳ������ʾ���༭,��һҳ����ʾ��ҳ�ķ�������
    '��鵱ǰҳ���Ƿ�Ϊ��������,����ɾ����Щ��������(���������)
    '������������Ƿ�Ϊ��������,�����ȡ��һҳ,����ҳ�Ĳ��ַ���������װ��һ��
    'If Val(VsfData.TextMatrix(VsfData.Rows - 1, mlngDemo)) > 0 And VsfData.Rows - VsfData.FixedRows >= mlngPageRows Then
    lngLiterrRow = 0
    mlngLitterRows(mintҳ��) = 0
    mlngCurLitterRows(mintҳ��) = 0
    mArrPageInfo(mintҳ��) = ""
    If VsfData.Rows > VsfData.FixedRows And Val(FormatValue(VsfData.TextMatrix(VsfData.Rows - 1, mlngRowCount))) > 0 And mintҳ�� >= mint��ʼҳ�� Then
        intCols = VsfData.Cols - 1
        lngTestRow = VsfData.FixedRows
        '��ȡ������ʼ��
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
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��������", mlng�ļ�ID, mlng����ID, mlng��ҳID, mintӤ��, mintҳ�� + 1)
        If rsData.RecordCount > 0 Then
            Set vsTest.DataSource = rsData
            Do While True
                If lngTestRow > vsTest.Rows - 1 Then Exit Do
                If Mid(strDate, 1, 16) <> Mid(Format(vsTest.TextMatrix(lngTestRow, mlngActTime), "YYYY-MM-DD HH:mm:ss"), 1, 16) Then Exit Do
                '82036:LPF,���⽫���������������������
                If Val(vsTest.TextMatrix(lngTestRow, mlngCollectType)) < 0 Or blnCollectType = True Then Exit Do
                
                If lngRecordId = Val(vsTest.TextMatrix(lngTestRow, mlngRecord)) Then GoTo ErrNext
                lngRowCount = Val(vsTest.TextMatrix(lngTestRow, mlngRowCount))
                VsfData.Rows = VsfData.Rows + lngRowCount
                lngLiterrRow = lngLiterrRow + lngRowCount
                
                For intCol = 0 To intCols
                    VsfData.TextMatrix(lngRow, intCol) = vsTest.TextMatrix(lngTestRow, intCol)
                Next
               'ѭ������ǰ������
                For lngCol = 0 To VsfData.Cols - 1
                    If VsfData.ColHidden(lngCol) And lngCol <> mlngRowCount And lngCol <> mlngDemo Then
                        'ѭ����ֵ
                        For intData = 2 To lngRowCount
                            VsfData.TextMatrix(lngRow + intData - 1, lngCol) = vsTest.TextMatrix(lngTestRow, lngCol)
                        Next
                    ElseIf (lngCol < mlngNoEditor And lngCol <> mlngDate And lngCol <> mlngTime And lngCol <> mlngYear) Then
                        '׼����ֵ
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
                            'ѭ����ֵ
                            If intDatas + 1 > lngRowCount Then intDatas = lngRowCount - 1
                            For intData = 0 To intDatas
                                If VsfData.Rows <= lngRow + intData Then VsfData.Rows = VsfData.Rows + 1
                                VsfData.TextMatrix(lngRow + intData, lngCol) = arrData(intData)
                            Next
                        End If
                    ElseIf lngCol = mlngNoEditor Then
                        '����ֵ��Ϊ��1��ʼ,������4������,����4|1
                        For intData = 1 To lngRowCount
                            VsfData.TextMatrix(lngRow + intData - 1, mlngRowCount) = lngRowCount & "|" & intData
                            VsfData.TextMatrix(lngRow, mlngYear) = ""
                            VsfData.TextMatrix(lngRow, mlngDate) = ""
                            VsfData.TextMatrix(lngRow, mlngTime) = ""
                        Next
                        '���һ����Ҫ��д���ǩ��
                        If mlngSignName > 0 Then VsfData.TextMatrix(lngRow + lngRowCount - 1, mlngSignName) = vsTest.TextMatrix(lngTestRow, mlngSignName)
                        If mlngSignTime > 0 Then VsfData.TextMatrix(lngRow + lngRowCount - 1, mlngSignTime) = vsTest.TextMatrix(lngTestRow, mlngSignTime)
                        '--58414,������,2013-04-10,��ӻ�ʿ��ǩ������ʾģʽ
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
        mArrPageInfo(mintҳ��) = mArrPageInfo(mintҳ��) & "[LPF]" & "��ǰҳ����ʱ��:" & Mid(strDate, 1, 16) & "�ķ�����������" & lngLiterrRow & "������Ϊ��һҳ�����ݡ�"
    End If
    
    If mintҳ�� > mint��ʼҳ�� Then
        Call SQLCombination
        gstrSQL = mstrSQL
        Call SQLDIY(gstrSQL)
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��������", mlng�ļ�ID, mlng����ID, mlng��ҳID, mintӤ��, mintҳ�� - 1)
        If rsData.RecordCount > 0 Then
            Set vsTest.DataSource = rsData
            
            '��ҳ������ʾ�ڵ�ǰ��ʱ��������һҳ�޷���ȡ����ҳ���ݣ���һҳ��ʵ������Ӧ�õ��ڱ�ҳ������+��ҳ��ҳ��-��һҳ��ҳ������
            '���磺��¼������20����һҳ��ҳ���ݿ���5�У�����ҳ������ʾ�ڵ�ǰҳʱ���ڶ�ҳ������Ӧ��Ϊ15�С�
            If mintPageSpan = 1 Then
                lngRowCount = Val(FormatValue(vsTest.TextMatrix(vsTest.Rows - 1, mlngRowCount)))
                lngRowCurrent = Val(FormatValue(vsTest.TextMatrix(vsTest.Rows - 1, mlngRowCurrent)))
                If lngRowCount > lngRowCurrent Then
                    mlngLitterRows(mintҳ��) = lngRowCount - lngRowCurrent
                    mlngCurLitterRows(mintҳ��) = mlngLitterRows(mintҳ��)
                    mArrPageInfo(mintҳ��) = mArrPageInfo(mintҳ��) & "[LPF]" & "���ڹ�ѡ�˲���:��ҳ����ֻ��ʾ�ڵ�ǰҳ,��ǰҳ��" & lngRowCount - lngRowCurrent & "��������ʾ����һҳ��" & _
                        IIf(Val(vsTest.TextMatrix(vsTest.Rows - 1, mlngCollectType)) < 0, "����С������:" & vsTest.TextMatrix(vsTest.Rows - 1, mlngCollectText), "���ݷ���ʱ��:" & Mid(vsTest.TextMatrix(vsTest.Rows - 1, mlngActTime), 1, 16))
                End If
            End If
            If VsfData.Rows > VsfData.FixedRows Then
                If Val(VsfData.TextMatrix(VsfData.FixedRows, mlngRowCount)) > 0 Then
                    '50503:������,2012-09-12,����������ʾ�ڷ�����ʼҳ���˴��������������˵��˵���ҳ������
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
                        mlngLitterRows(mintҳ��) = Val(CStr(mlngLitterRows(mintҳ��))) + lngRowCount
                        mlngCurLitterRows(mintҳ��) = Val(CStr(mlngCurLitterRows(mintҳ��))) + lngRowCurrent
                        Do While True
                            If Val(FormatValue(VsfData.TextMatrix(VsfData.FixedRows, mlngDemo))) <= 1 Then Exit Do
                            lngRowCount = Val((VsfData.TextMatrix(VsfData.FixedRows, mlngRowCount)))
                            lngRowCurrent = Val((VsfData.TextMatrix(VsfData.FixedRows, mlngRowCurrent)))
                            For intData = 1 To lngRowCount
                                If VsfData.Rows > VsfData.FixedRows Then VsfData.RemoveItem VsfData.FixedRows
                            Next intData
                            lngCount = lngCount + lngRowCount
                            mlngLitterRows(mintҳ��) = Val(CStr(mlngLitterRows(mintҳ��))) + lngRowCount
                            mlngCurLitterRows(mintҳ��) = Val(CStr(mlngCurLitterRows(mintҳ��))) + lngRowCurrent
                        Loop
                        If VsfData.Rows - 1 > VsfData.FixedRows Then
                            VsfData.RemoveItem VsfData.Rows - 1
                        End If
                    End If
                    If lngCount > 0 Then
                        mArrPageInfo(mintҳ��) = mArrPageInfo(mintҳ��) & "[LPF]" & "��ǰҳ��" & lngCount & "�з���������ʾ����һҳ���������ݷ���ʱ��:" & strDate
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

    If Val(CStr(mlngLitterRows(mintҳ��))) - lngLiterrRow <= 0 Then
        mlngLitterRows(mintҳ��) = 0
    Else
        mlngLitterRows(mintҳ��) = Val(CStr(mlngLitterRows(mintҳ��))) - lngLiterrRow
    End If
    
    If Val(CStr(mlngCurLitterRows(mintҳ��))) - lngLiterrRow <= 0 Then
        mlngCurLitterRows(mintҳ��) = 0
    Else
        mlngCurLitterRows(mintҳ��) = Val(CStr(mlngCurLitterRows(mintҳ��))) - lngLiterrRow
    End If
    
    '63760:������,�������ݻ�ʿ��ǩ���ˡ�ǩ��ʱ��Ĵ���ͬһ��ǩ����ʼ����ʾһ�Σ�
    If mlngSingerType > 0 And VsfData.FixedRows <= VsfData.Rows - 1 Then
        lngPrintedRow = IIf(mlngOperator <> -1, mlngOperator, mlngSignName)
        lngRow = VsfData.FixedRows
        Do While True
            lngStart = GetStartRow(lngRow)
            lngRowCount = Val(VsfData.TextMatrix(lngStart, mlngRowCount))
            If lngRowCount <= 0 Then Exit Do
            
            If mlngSingerType = 3 Then 'β��ǩ��
                strSignName = VsfData.TextMatrix(lngStart + lngRowCount - 1, lngPrintedRow)
            Else '����ǩ������βǩ��
                strSignName = VsfData.TextMatrix(lngStart, lngPrintedRow)
            End If
            strSignName = FormatValue(strSignName)
            If Val(VsfData.TextMatrix(lngStart, mlngDemo)) = 1 And lngStart = lngRow And strSignName <> "" Then
                For lngRow = lngStart + lngRowCount To VsfData.Rows - 1
                    If lngRow = lngStart + lngRowCount Then
                    
                        If Val(VsfData.TextMatrix(lngRow, mlngDemo)) <= 1 Then Exit For
                        
                        lngRowCount = Val(VsfData.TextMatrix(lngRow, mlngRowCount))
                        If lngRowCount = 0 Then Exit For
                        
                        If mlngSingerType = 3 Then 'β��ǩ��
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
                        Else '����ǩ������βǩ��
                            If strSignName = FormatValue(VsfData.TextMatrix(lngRow, lngPrintedRow)) Then
                                '����ǩ������βǩ������Ҫȥ����һ�����ݵ�����,����βǩ����Ҫע������е����һ����������=1�����
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
                                
                                If mlngSingerType = 2 And lngStart < lngRow - 1 Then '��βǩ����Ӧ��ȥ����һ�����ݵ�β��(��һ������������Ҫ>1)
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
    
    '���ñ�ͷ�ĸ�ʽ
    With vsfHead
        .Redraw = flexRDNone
        .Clear
        Set .DataSource = rsTemp
        .Rows = 3
        
        '��ͷ��д
        .MergeCells = flexMergeRestrictRows
        .MergeCellsFixed = flexMergeFree
        .MergeRow(0) = True
        .MergeRow(1) = True
        .MergeRow(2) = True
        
        '�����ڲ�����������
        .ColHidden(mlngDemo) = True
        .ColHidden(mlngActTime) = True
        .ColHidden(mlngChoose) = Not mblnVerify
        '69355:������,2014-01-07,���ڴ��ڶԽ���(�����ڸ�ʽ7/1),����ʾ�����
        .ColHidden(mlngYear) = Not mblnDateAd
        .ColHidden(mlngRowCurrent) = True
        .ColHidden(mlngRowCount) = True
        .ColHidden(mlngStartRowPage) = True
        .ColHidden(mlngStartRowNo) = True
        '51589:������,2013-03-01,��ӽ���ǩ��
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
        '63706:������,2013-11-20
        .ColHidden(.Cols - 1) = True
        .ColWidth(0) = 250
        .ColWidth(mlngChoose) = 250      'ѡ����
        .ColWidth(mlngYear) = BlowUp(picMain.TextWidth("������"))
        
        .FrozenCols = mlngTime
        .SheetBorder = &H40C0&
        
        '������ͷ
        aryItem = Split(mstrTabHead, "|")
        For lngCount = 0 To UBound(aryItem)
            strCell = aryItem(lngCount)
            lngRow = Left(strCell, InStr(1, strCell, ",") - 1): strCell = Mid(strCell, InStr(1, strCell, ",") + 1)
            lngCol = Left(strCell, InStr(1, strCell, ",") - 1): strCell = Mid(strCell, InStr(1, strCell, ",") + 1)
            .TextMatrix(lngRow, lngCol + cHideCols + .FixedCols - 1) = strCell
        Next
        
        
        '���ù̶��м�ѡ����
        .TextMatrix(0, 0) = " "
        .TextMatrix(1, 0) = " "
        .TextMatrix(2, 0) = " "
        .TextMatrix(0, mlngChoose) = " "
        .TextMatrix(1, mlngChoose) = " "
        .TextMatrix(2, mlngChoose) = " "
        .TextMatrix(0, mlngYear) = "���"
        .TextMatrix(1, mlngYear) = "���"
        .TextMatrix(2, mlngYear) = "���"
        Call PreActiveHead(vsfHead)
        
        '�п�����
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
        
        '�̶��и�ʽΪ����
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
        '�ٰ��кϲ�
        For lngCount = 0 To .Cols - 1
            .MergeCol(lngCount) = True
        Next
        
        If blnAlign = False Then
            '��Ϊ�����û���������ʾ�ж��뷽ʽ
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
        
        '���̶ܹ��е��и߲���ȷ��Ҫ�Զ�������
        .AutoResize = True
        .WordWrap = True
        .AutoSizeMode = flexAutoSizeRowHeight
        .AutoSize 0, .Cols - 1
        .AutoResize = False
        '���ǹ̶��е��и�����Ϊ��С�и�
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
    
    '���û����¼���ĸ�ʽ
    With VsfData
        .Redraw = flexRDNone
        .Clear
        Call DataMap_Restore(rsTemp)
        
        '��ͷ��д
        .MergeCells = flexMergeRestrictRows
        .MergeCellsFixed = flexMergeFree
        .MergeRow(0) = True
        .MergeRow(1) = True
        .MergeRow(2) = True
        
        '�����ڲ�����������
        .ColHidden(mlngDemo) = True
        .ColHidden(mlngActTime) = True
        .ColHidden(mlngChoose) = Not mblnVerify
        .ColHidden(mlngYear) = Not mblnDateAd
        .ColHidden(mlngRowCurrent) = True
        .ColHidden(mlngRowCount) = True
        '51589:������,2013-03-01,��ӽ���ǩ��
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
        '63706:������,2013-11-20
        .ColHidden(.Cols - 1) = True
        .ColWidth(0) = 250
        .ColWidth(mlngChoose) = 250      'ѡ����
        .ColWidth(mlngYear) = BlowUp(picMain.TextWidth("������"))
        
        .FrozenCols = mlngTime
        .SheetBorder = &H40C0&
        
        '������ͷ
        aryItem = Split(mstrTabHead, "|")
        For lngCount = 0 To UBound(aryItem)
            strCell = aryItem(lngCount)
            lngRow = Left(strCell, InStr(1, strCell, ",") - 1): strCell = Mid(strCell, InStr(1, strCell, ",") + 1)
            lngCol = Left(strCell, InStr(1, strCell, ",") - 1): strCell = Mid(strCell, InStr(1, strCell, ",") + 1)
            .TextMatrix(lngRow, lngCol + cHideCols + .FixedCols - 1) = strCell
        Next
        
        '���ù̶��м�ѡ����
        .TextMatrix(0, 0) = " "
        .TextMatrix(1, 0) = " "
        .TextMatrix(2, 0) = " "
        .TextMatrix(0, mlngChoose) = " "
        .TextMatrix(1, mlngChoose) = " "
        .TextMatrix(2, mlngChoose) = " "
        .TextMatrix(0, mlngYear) = "���"
        .TextMatrix(1, mlngYear) = "���"
        .TextMatrix(2, mlngYear) = "���"
        
        Call PreActiveHead(VsfData)
        
        '�п�����
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
        
        '�̶��и�ʽΪ����
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
        '�ٰ��кϲ�
        For lngCount = 0 To .Cols - 1
            .MergeCol(lngCount) = True
        Next
        
        If blnAlign = False Then
            '��Ϊ�����û���������ʾ�ж��뷽ʽ
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
            '�õ���һ�м��ٵ���(��Ҫ����Ժϲ���ӡ���ļ���������һ���ļ����ڰ�ҳ���ݵ����)
            '���磺�ļ�1�����к�Ϊ9�����ļ����ļ�1�ϲ�����ô���ļ��ĵ�һҳ�Ŀ�ʼ�к�=10.��ʱ��¼����ʾ��������Ҫ��ȥ9��
            If mintҳ�� = mint��ʼҳ�� And Val(VsfData.TextMatrix(3, mlngStartRowNo)) > 1 Then
                mlngReduceRow = Val(VsfData.TextMatrix(3, mlngStartRowNo)) - 1
                If picImg.Tag <> "" Then
                    strInfo = strInfo & "[LPF]" & "���ڵ�ǰ�ļ����ļ�'" & picImg.Tag & "'�����˺ϲ���ӡ,�����ϲ����ļ����һҳ����δ��ҳ����˵�ǰҳ��������" & mlngReduceRow & "����ʾ�ڱ��ϲ����ļ���"
                End If
            Else
                mlngReduceRow = 0
            End If
            '�õ���һ�еĳ�����
            mlngOverrunRows = Val(.TextMatrix(3, mlngRowCount)) - Val(.TextMatrix(3, mlngRowCurrent))
            If mlngOverrunRows > 0 Then
                If Val(.TextMatrix(3, mlngStartRowPage)) = mintҳ�� Then
                    strInfo = strInfo & "[LPF]" & IIf(Val(.TextMatrix(3, mlngCollectType)) < 0, "��ǰҳС��:" & .TextMatrix(3, mlngCollectText), "��ǰҳ����ʱ��:" & Mid(.TextMatrix(3, mlngActTime), 1, 16)) & _
                        "�����ݴӵ�" & Val(.TextMatrix(3, mlngRowCurrent)) + 1 & "�п�ʼ��ҳ,��ҳ����" & mlngOverrunRows & "�С�"
                Else
                    strInfo = strInfo & "[LPF]" & IIf(Val(.TextMatrix(3, mlngCollectType)) < 0, "��ǰҳС��:" & .TextMatrix(3, mlngCollectText), "��ǰҳ����ʱ��:" & Mid(.TextMatrix(3, mlngActTime), 1, 16)) & _
                        "������ǰ" & Val(.TextMatrix(3, mlngRowCurrent)) & "��Ϊ��һҳ�����ݡ�"
                End If
            End If
            '50503:������,2012-09-12,���ݴ�ĳһҳ��һ�оͿ�ʼ��ҳ����������в����ظ����㣬�����޸ģ�
            '���һ:
            '���:��һ�к����һ�еļ�¼��ʼ����ͬ����ô˵����ͬһ�����ݣ��������ۼƲ��������һ�еĳ�����
            '�����:81982,������,2015-01-30�������������еĴ���(������ܳ�ʼ�����ж��м�¼)
            '����:���ڷ���������ݵ�һ�γ�ʼ��ʱ�����ݿ��ܴ��ڶ���(��ʵ�ʶ���������һ������)��Ϊ�˱����ظ�����У�����Ҫ���ݼ�¼ID�ж��Ƿ���ͬһ�����ݡ�
            'ԭ��:1.��ͨ���ݣ����������Ƿ�չ��ʼ�տ��Դ����һ�жϵĳ�����2���κ�����ֻҪ��չ��Ҳ�ɴ�������жϵĳ�����3.ֻ�з����������չ��ǰ�޷������һ�жϡ�
            lngStartRow = 3
            '�������һ�еĳ�����
            If .Rows - 1 <> 3 Then
                If Val(.TextMatrix(.Rows - 1, mlngRowCount)) - Val(.TextMatrix(.Rows - 1, mlngRowCurrent)) > 0 Then
                    If lngStartRow <> GetStartRow(.Rows - 1) Then
                        If Val(.TextMatrix(lngStartRow, mlngRecord)) <> Val(.TextMatrix(.Rows - 1, mlngRecord)) And Val(.TextMatrix(lngStartRow, mlngRecord)) <> 0 And Val(.TextMatrix(.Rows - 1, mlngRecord)) <> 0 Then
                            mlngOverrunRows = mlngOverrunRows + Val(.TextMatrix(.Rows - 1, mlngRowCount)) - Val(.TextMatrix(.Rows - 1, mlngRowCurrent))
                            strInfo = strInfo & "[LPF]" & IIf(Val(.TextMatrix(.Rows - 1, mlngCollectType)) < 0, "��ǰҳС��:" & .TextMatrix(.Rows - 1, mlngCollectText), "��ǰҳ����ʱ��:" & Mid(.TextMatrix(.Rows - 1, mlngActTime), 1, 16)) & _
                                "�����ݴӵ�" & Val(.TextMatrix(.Rows - 1, mlngRowCurrent)) + 1 & "�п�ʼ��ҳ,��ҳ����" & Val(.TextMatrix(.Rows - 1, mlngRowCount)) - Val(.TextMatrix(.Rows - 1, mlngRowCurrent)) & "�С�"
                        End If
                    End If
                End If
               ' mlngOverrunRows = mlngOverrunRows + Val(.TextMatrix(.Rows - 1, mlngRowCount)) - Val(.TextMatrix(.Rows - 1, mlngRowCurrent))
            End If
        End If
        Call PreTendMutilRows
        lblInfo.Tag = Trim(Mid(strInfo & mArrPageInfo(mintҳ��), 6))
        picInfo.Visible = lblInfo.Tag <> ""
        Call FillPage
        
        Call WriteColor
        
        '���̶ܹ��е��и߲���ȷ��Ҫ�Զ�������
        .AutoResize = True
        .WordWrap = True
        .AutoSizeMode = flexAutoSizeRowHeight
        .AutoSize 0, .Cols - 1
        .AutoResize = False
        
        .RowHidden(0) = True
        .RowHidden(1) = True
        .RowHidden(2) = True
        
        '���ǹ̶��е��и�����Ϊ��С�и�
        For lngCount = .FixedRows To .Rows - 1
            .RowHeight(lngCount) = .RowHeightMin
        Next
        
        .WordWrap = False           '�����Զ�����,�����������һ���ֿ����������
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
    '����Ժ�ɫ��ʾ��ͬʱ������ʼ������ΪNoCheckBox������ͼ��
    With VsfData
        For lngCount = .FixedRows To .Rows - 1
            If .TextMatrix(lngCount, 2) <> "" And Val(.TextMatrix(lngCount, mlngCollectType)) = 0 Then
                '����Ժ�ɫ��ʾ
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
            
            '���������,���Ϊ����ʾΪ��
            If Val(.TextMatrix(lngCount, mlngCollectType)) < 0 And FormatValue(VsfData.TextMatrix(lngCount, mlngRowCount)) Like "*|1" Then
                '88967:��ʿ��ǩ����ͬʱ���ڣ�������ͬһ����Ա����Ӧ����ϲ�(��ӡǩ�������ǩ��ͼƬ��ע��ǩ�����лس������)
                For lngCol = mlngTime + 1 To IIf(mlngNoEditor < mlngSignName, mlngSignName, mlngNoEditor)
                    '52953,������,2012-08-24,��������Ϊ0ҲҪ��ʾ,��������60792
                    'If .TextMatrix(lngCount, lngCOL) = "0" Then .TextMatrix(lngCount, lngCOL) = ""
                    .TextMatrix(lngCount, lngCol) = FormatValue(.TextMatrix(lngCount, lngCol))
                    If Trim(.TextMatrix(lngCount, lngCol)) <> "" And .ColHidden(lngCol) = False Then
                        '66085:������,2012-09-26,�������ڻ����кϲ�,��ԭ����������+�ո�ͬһ�ĳ����к�����chr(13)
                        '������ӿո���п�������������ʾ����ȫ(��Ҫ����Ҷ���)
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
            
            '������ʼ������ΪNoCheckBox
            If Not VsfData.RowHidden(lngCount) Then
                If Not FormatValue(VsfData.TextMatrix(lngCount, mlngRowCount)) Like "*|1" Then
                    VsfData.Cell(flexcpChecked, lngCount, mlngChoose) = flexNoCheckbox
                Else
                    If VsfData.Cell(flexcpChecked, lngCount, mlngChoose) <> flexTSChecked Then
                        VsfData.Cell(flexcpChecked, lngCount, mlngChoose) = flexTSUnchecked
                    End If
                    
                    '����ͼ��
                    If FormatValue(VsfData.TextMatrix(lngCount, mlngSigner)) = "" Then
                        VsfData.Cell(flexcpPicture, lngCount, 0) = Nothing
                    Else
                        If InStr(1, VsfData.TextMatrix(lngCount, mlngSigner), "/") <> 0 Then
                            VsfData.Cell(flexcpPicture, lngCount, 0) = imgRow.ListImages(��ǩ).Picture
                        '51589:������,2013-03-01,��ӽ���ǩ��
                        ElseIf Trim(VsfData.TextMatrix(lngCount, mlngJoinSignName)) <> "" Then '����ǩ��
                            VsfData.Cell(flexcpPicture, lngCount, 0) = imgRow.ListImages(����ǩ��).Picture
                        Else
                            VsfData.Cell(flexcpPicture, lngCount, 0) = imgRow.ListImages(ǩ��).Picture
                        End If
                    End If
                
                    '����С�����ʾ
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
    '��ȡ�ļ�����
    Dim lngYear As Long
    Dim strEndTime As String, strCurDate As String
    On Error GoTo ErrHand
    
    '"���ڼ�¼������������ʾ��˵����"
    '"1����������ֻ��ʾ�ڷ�����ʼҳ,�����ζ�ŷ�����ʼҳ��������������,��һҳ�������������١�"
    '"2����ҳ���ݸ��ݲ���'��ҳ����ֻ��ʾ�ڵ�ǰҳ'��������ʾ�ڵ�ǰҳ������ҳ����ʾ�������ʾ�ڵ�ǰҳ����ζ����ʾ�ڵ�ǰҳ�����������ӣ���һҳ��������������;�����ҳ����ʾ����ζ����ʾ����ҳ���������������ӡ�"
    '"3����ǰ�ļ��������һ���ļ������˺ϲ���ӡ������Ϸ��ļ����һҳδ��ҳ����ô��ǰ�ļ���ʾ�����������ͻ���١�"
    strCurDate = zlDatabase.Currentdate
    
    gstrSQL = " Select   ��ʼʱ��,����ʱ��,��ʽID,����ID,�鵵�� From ���˻����ļ� " & _
              " Where ����ID=[1] And ��ҳID=[2] And Ӥ��=[3] And ID=[4] And Rownum<2"
    Call SQLDIY(gstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����ļ�����", mlng����ID, mlng��ҳID, mintӤ��, mlng�ļ�ID)
    If rsTemp.RecordCount <> 0 Then
        mlng��ʽID = rsTemp!��ʽID
        mlng����ID = rsTemp!����ID
        mblnArchive = (NVL(rsTemp!�鵵��) <> "")
        mstr��ʼʱ�� = Format(rsTemp!��ʼʱ��, "yyyy-MM-dd HH:mm:ss")
        mstr����ʱ�� = Format(rsTemp!����ʱ��, "yyyy-MM-dd HH:mm:ss")
    End If
    '69935:������,2014-1-7
    If IsDate(mstr����ʱ��) Then
        strEndTime = mstr����ʱ��
    Else
        strEndTime = mstrMaxDate
    End If
    mstrYears = ""
    For lngYear = Val(Format(strEndTime, "YYYY")) To Val(Format(mstr��ʼʱ��, "YYYY")) Step -1
        If Val(Format(strCurDate, "YYYY")) = lngYear Then
            mstrYears = mstrYears & "|" & lngYear
        Else
            mstrYears = mstrYears & "|" & lngYear
        End If
    Next lngYear
    mstrYears = Mid(mstrYears, 2)
    
    '���ҳ��=-1,˵��ȱʡ��ʾ���һҳ
    mint��ʼҳ�� = 1
    gstrSQL = " Select  MIN(��ʼҳ��) ��ʼҳ��, MAX(����ҳ��) AS ҳ�� From ���˻����ӡ Where �ļ�ID=[1]"
    Call SQLDIY(gstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡָ��ҳ������ݷ���ʱ�䷶Χ", mlng�ļ�ID)
    mint��ʼҳ�� = NVL(rsTemp!��ʼҳ��, 1)
    mint����ҳ = NVL(rsTemp!ҳ��, 1)
    If mintҳ�� = -1 Then mintҳ�� = mint����ҳ
    If mintҳ�� < mint��ʼҳ�� Then mintҳ�� = mint��ʼҳ��
    If mintҳ�� > mint����ҳ + 1 Then mintҳ�� = mint����ҳ
    
    '��ȡ�ϲ��ļ�����������ʾ
    picImg.Tag = ""
    gstrSQL = "Select �ļ����� From ���˻����ļ� Where ����ID=[1] And ��ҳID=[2] And NVL(Ӥ��,0)=[3] And ����ID=[4]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�͵�ǰ�ļ��ϲ����ļ�", mlng����ID, mlng��ҳID, mintӤ��, mlng�ļ�ID)
    If rsTemp.RecordCount > 0 Then
        picImg.Tag = NVL(rsTemp!�ļ�����)
    End If
    
    Call InitPages
    
    If mblnClear = True Then
        ReDim mArrPageInfo(0 To mint����ҳ + 1)
        ReDim mlngLitterRows(0 To mint����ҳ + 1)
        ReDim mlngCurLitterRows(0 To mint����ҳ + 1)
    Else
        ReDim Preserve mArrPageInfo(0 To mint����ҳ + 1)
        ReDim Preserve mlngLitterRows(0 To mint����ҳ + 1)
        ReDim Preserve mlngCurLitterRows(0 To mint����ҳ + 1)
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
    
    glngHours = Val(zlDatabase.GetPara("���ݲ�¼ʱ��", glngSys))

    '���ִ��ڵ����л����¼��Ŀ
    gstrSQL = " Select   ��Ŀ���,upper(��Ŀ����) AS ��Ŀ����,��Ŀ����,��Ŀ����,��Ŀ����,��ĿС��,��Ŀ��ʾ,��Ŀ��λ,��Ŀֵ��,����ȼ�,Ӧ�÷�ʽ,˵��" & _
              " From �����¼��Ŀ B" & _
              " Order by ��Ŀ���"
    Set mrsItems = zlDatabase.OpenSQLRecord(gstrSQL, "���ִ��ڵ����л����¼��Ŀ")
    
    '��ȡ�����ڼ�¼��������������Ŀ
    gstrSQL = _
        " Select i.����id, i.����, i.������, nvl(i.�滻��,0) �滻��,i.����,i.����,i.С��,i.��λ,i.��ʾ��,i.��ֵ��,i.����" & vbNewLine & _
        " From ����������Ŀ i, ������������ k" & vbNewLine & _
        " Where k.Id = i.����id And ((k.���� In ('02', '05', '06') And i.�滻�� = 1) Or (k.���� = 2 And k.���� = '06' And NVL(i.�滻��,0) = 0))" & vbNewLine & _
        " Order By k.����, k.����, i.����"
    Set mrsElement = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����ڼ�¼��������������Ŀ")
    
    'ȡ��ǰ����Ա�ļ���
    mintVerify = δ����
    mintVerify_Last = δ����
    gstrSQL = "select  Ƹ�μ���ְ�� from ��Ա�� p where ID=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mRichEPR", glngUserId)
    If Not rs.EOF Then
        mintVerify = NVL(rs("Ƹ�μ���ְ��"), δ����)
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
        '��ʼ���ڴ��¼��(δ��Ӧ��Ŀ����Ϊ���Ŀ,�����о�Ϊ�̶���)
        strFields = "��," & adDouble & ",18|���," & adDouble & ",2|��Ŀ���," & adDouble & ",18|��Ŀ����," & adLongVarChar & ",20|�̶�," & adDouble & ",2|��ʽ," & adLongVarChar & ",2000"
        Call Record_Init(mrsSelItems, strFields)
        strFields = "��|���|��Ŀ���|��Ŀ����|�̶�|��ʽ"
    End If
    
    '�����ж���
    If Not mblnInit Then
        arrColumn = Split(strColumns, "|")
        j = UBound(arrColumn)
        For i = 0 To j
            lngCol = Split(arrColumn(i), "'")(0)
            arrItem = Split(Split(arrColumn(i), "'")(1), ",")
            blnSet = False   '����������Դ���ֵΪ׼'�����Ҳ�����Ŀ���ǻ��Ŀ
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
                mrsItems.Filter = "��Ŀ����='" & strName & "'"
                If mrsItems.RecordCount <> 0 Then
                    lngOrder = mrsItems!��Ŀ���
                    If Not blnSet Then intImmovable = 1   '�̶��������޸�
                Else
                    lngOrder = 0
                    If Not blnSet Then intImmovable = 0
                    
                    '��¼������
                    Select Case strName
                    Case "����"
                        mlngDate = i + cHideCols + VsfData.FixedCols
                    Case "ʱ��"
                        mlngTime = i + cHideCols + VsfData.FixedCols
                    Case "��ʿ"
                        mlngOperator = i + cHideCols + VsfData.FixedCols
                    Case "ǩ����"
                        mlngSignName = i + cHideCols + VsfData.FixedCols
                    Case "ǩ��ʱ��"
                        mlngSignTime = i + cHideCols + VsfData.FixedCols
                    End Select
                End If
                strValues = lngCol & "|" & l + 1 & "|" & lngOrder & "|" & strName & "|" & intImmovable & "|" & strFormat
                Call Record_Add(mrsSelItems, strFields, strValues)
            Next
        Next
        
        '���������ܹ�������Ϣ
        arrCorrelative = Array()
        arrColumn = Split(mstrColCorrelative, "|")
        For i = 0 To UBound(arrColumn)
            arrItem = Split(arrColumn(i), ";")
            If UBound(arrItem) = 1 Then
                mrsSelItems.Filter = "��=" & Val(arrItem(0))
                If mrsSelItems.RecordCount = 1 Then
                    ReDim Preserve arrCorrelative(UBound(arrCorrelative) + 1)
                    arrCorrelative(UBound(arrCorrelative)) = Val(arrItem(0)) & "," & mrsSelItems!��Ŀ��� & ";" & CStr(arrItem(1))
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
        
        '��������ڲ�������(�����ڶ�ȡ���ݺ��ʱ���ӵ�,��ʱֻ��Ԥ������)
        mlngDemo = VsfData.FixedCols
        mlngActTime = mlngDemo + 1
        mlngChoose = mlngActTime + 1
        mlngYear = mlngChoose + 1
        mlngSignLevel = VsfData.Cols + cHideCols + VsfData.FixedCols '����������
        mlngSigner = mlngSignLevel + 1
        '51589:������,2013-03-01,��ӽ���ǩ��
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
    mlng����ID = lngPatiID
    mlng��ҳID = lngPageId
    mintӤ�� = intBaby
End Sub

Public Sub ArchiveMe()
    On Error GoTo ErrHand
    
    If mlng����ID = 0 Or gblnMoved Then Exit Sub
    If MsgBox("��Ҫ���ò��˱���סԺ���л����ļ��鵵��", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbYes Then
        Dim strNow As String

        strNow = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
        gstrSQL = "ZL_���˻����ļ�_ARCHIVE(" & mlng����ID & "," & mlng��ҳID & "," & mintӤ�� & ",1)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "�鵵")

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
    
    If mlng����ID = 0 Or gblnMoved Then Exit Sub
    If MsgBox("��Ҫȡ���ò��˵Ĺ鵵״̬��", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbYes Then

        gstrSQL = "ZL_���˻����ļ�_ARCHIVE(" & mlng����ID & "," & mlng��ҳID & "," & mintӤ�� & ",0)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "�����鵵")
        
        mblnArchive = False
        RaiseEvent AfterRowColChange("", False, mblnSign, mblnArchive)
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function SignMe(Optional ByVal bln��ǩ As Boolean = False, Optional ByVal blnExchange As Boolean = False) As Boolean
    Dim blnSign As Boolean          '�Ƿ�ǩ���ɹ�
    Dim blnRefresh As Boolean
    Dim strSignTime As String       '��֤����ǩ����ǩ��ʱ��һ��,����ȡ��ǩ��ʱ��ǩ��ʱ��ͳһȡ��
    Dim str״̬ As String           '����ǩ��ѡ��,����ѭ��ǩ��ʱ��ͣ�ĵ���ǩ������
    Dim str�д��� As String
    Dim str���� As String
    Dim intRow As Integer, intRows As Integer
    Dim rsTemp As New ADODB.Recordset
    Dim strInfo As String
    
    Dim lngStart As Long, lngDemo As Long, lngRow As Long
    On Error GoTo ErrHand
    '������ʱ��ѭ��������δǩ�����ݽ���ǩ��
    
    If mlng����ID = 0 Then Exit Function
    
    '��ǩ:������δǩ�������ݽ���ǩ��
    '��ǩ:��������ǩ�������ݽ�����ǩ
    If bln��ǩ Then
        blnExchange = False
        '43588:������,2012-09-13,��Ӽ�¼����ǩģʽ
        '0-ȱʡ��Ƹ��ְ��+��ǩȨ�ޣ�����ǩʱ��ְ��ߵͽ��п��ƣ����ɴ��ļ���ǩ��1-��ǩȨ�ޣ�ֻ�о�����ǩȨ�޵��˿�����ǩ���˼�¼����ǩ�����ٴ���ǩ��
        If Not mblnVerify Then
            '��������ҲҪǩ��,���ȥ������: And B.�������=0
            If mintSignMode = 1 Then
                gstrSQL = " Select  distinct B.ID,B.����ʱ�� " & vbNewLine & _
                          " From ���˻�����ϸ A,���˻������� B,���˻����ļ� C" & vbNewLine & _
                          " Where A.��¼ID=B.ID And B.�ļ�ID=C.ID And A.������Դ in (0,3) And A.��¼����=5 AND A.��ֹ�汾 Is NULL And C.ID=[1] " & _
                          " MINUS" & vbNewLine & _
                          " Select  distinct B.ID,B.����ʱ�� " & vbNewLine & _
                          " From ���˻�����ϸ A,���˻������� B,���˻����ļ� C" & vbNewLine & _
                          " Where A.��¼ID=B.ID And B.�ļ�ID=C.ID And A.������Դ in (0,3)  And A.��¼����=15  AND A.��ֹ�汾 Is NULL And C.ID=[1] "
            Else
                gstrSQL = " Select  distinct B.ID,B.����ʱ�� " & vbNewLine & _
                          " From ���˻�����ϸ A,���˻������� B,���˻����ļ� C" & vbNewLine & _
                          " Where A.��¼ID=B.ID And B.�ļ�ID=C.ID And A.������Դ in (0,3)  And MOD(A.��¼����,10)=5  AND A.��ֹ�汾 Is NULL And C.ID=[1] "
            End If
            Call SQLDIY(gstrSQL)
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����Ƿ����δǩ��������", mlng�ļ�ID)
            If rsTemp.RecordCount = 0 Then
                RaiseEvent AfterRowColChange("��������ǩ�������ݣ�", True, mblnSign, mblnArchive)
                Exit Function
            End If
        
            '������ǩģʽ,���޸�����,�ɹ�ѡ����
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
            '��ȡ����ǩ������
            '��������ҲҪǩ��,���ȥ������: And B.�������=0
            If mintSignMode = 1 Then
                gstrSQL = " Select /*+ RULE */ distinct B.ID,B.����ʱ�� " & vbNewLine & _
                          " From ���˻�����ϸ A,���˻������� B,���˻����ļ� C,(SELECT COLUMN_VALUE FROM TABLE(CAST(F_NUM2LIST([2]) AS ZLTOOLS.T_NUMLIST))) G" & vbNewLine & _
                          " Where A.��¼ID=B.ID And B.ID=G.COLUMN_VALUE And B.�ļ�ID=C.ID And A.��¼����=5  AND A.��ֹ�汾 Is NULL And C.ID=[1] " & _
                          " MINUS" & vbNewLine & _
                          " Select /*+ RULE */ distinct B.ID,B.����ʱ�� " & vbNewLine & _
                          " From ���˻�����ϸ A,���˻������� B,���˻����ļ� C,(SELECT COLUMN_VALUE FROM TABLE(CAST(F_NUM2LIST([2]) AS ZLTOOLS.T_NUMLIST))) G" & vbNewLine & _
                          " Where A.��¼ID=B.ID And B.ID=G.COLUMN_VALUE And B.�ļ�ID=C.ID And A.��¼����=15  AND A.��ֹ�汾 Is NULL And C.ID=[1] "
            Else
                gstrSQL = " Select /*+ RULE */ distinct B.ID,B.����ʱ�� " & vbNewLine & _
                          " From ���˻�����ϸ A,���˻������� B,���˻����ļ� C,(SELECT COLUMN_VALUE FROM TABLE(CAST(F_NUM2LIST([2]) AS ZLTOOLS.T_NUMLIST))) G" & vbNewLine & _
                          " Where A.��¼ID=B.ID And B.ID=G.COLUMN_VALUE And B.�ļ�ID=C.ID And MOD(A.��¼����,10)=5  AND A.��ֹ�汾 Is NULL And C.ID=[1] "
            End If
            Call SQLDIY(gstrSQL)
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����Ƿ����δǩ��������", mlng�ļ�ID, mstrVerify)
        End If
    Else
        '���Ա����޸ĵ����ݽ���ǩ��(��ȡδǩ������-��ǩ������)
        '��������ҲҪǩ��,���ȥ������: And B.�������=0
        mintVerify_Last = δ����
        '51589:������,2013-03-01,��ӽ���ǩ��
        If blnExchange = False Then
            gstrSQL = "" & _
                    "SELECT  DISTINCT B.ID,B.����ʱ��" & vbNewLine & _
                    "FROM ���˻�����ϸ A,���˻������� B" & vbNewLine & _
                    "WHERE A.��¼ID=B.ID And A.������Դ in (0,3) AND A.��ֹ�汾 IS NULL AND A.��¼���� =1 AND instr(NVL(B.ǩ����,'QMR'),'/',1)=0 AND A.��¼��=[2] AND B.�ļ�ID=[1]" & vbNewLine & _
                    "MINUS" & vbNewLine & _
                    "SELECT DISTINCT B.ID,B.����ʱ��" & vbNewLine & _
                    "FROM ���˻�����ϸ A,���˻������� B" & vbNewLine & _
                    "WHERE A.��¼ID=B.ID And A.������Դ in (0,3) AND A.��ֹ�汾 IS NULL AND A.��¼���� =5 AND A.��¼��=[2] AND B.�ļ�ID=[1]"
            Call SQLDIY(gstrSQL)
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����Ƿ����δǩ��������", mlng�ļ�ID, gstrUserName)
            If rsTemp.RecordCount = 0 Then
                '101151,����,2016-10-19,���δǩ�����ݼ�¼����ʾ
                lngStart = GetStartRow(VsfData.ROW)
                If mbln��ʿ = True Then
                    strInfo = VsfData.TextMatrix(lngStart, mlngOperator)
                Else
                    strInfo = VsfData.TextMatrix(lngStart, VsfData.Cols - 1)
                End If
                If strInfo <> "" Then strInfo = "  ��ǰ���ݼ�¼�ˣ�" & strInfo
                RaiseEvent AfterRowColChange("û���ҵ���Ҫǩ�������ݣ�ֻ�ܶ��Լ��Ǽǻ��޸ĵ����ݽ���ǩ������" & strInfo, True, mblnSign, mblnArchive)
                Exit Function
            End If
        Else '����ǩ��
            lngStart = GetStartRow(VsfData.ROW)
            '���Ƚ��������ж�:�Ƿ�ѡ���Ѿ�ǩ��������
            If Val(VsfData.TextMatrix(lngStart, mlngRecord)) = 0 Then
                RaiseEvent AfterRowColChange("����ѡ��Ҫ���н���ǩ���������У�", True, mblnSign, mblnArchive)
                Exit Function
            End If
            '���ڷ������ݽ���ǩ��ʱ��ֻ��Ҫ��֤��ʼ��
            lngDemo = Val(VsfData.TextMatrix(lngStart, mlngDemo))
            If lngDemo > 1 Then 'Ѱ�ҷ���������ʼ��
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
                    "SELECT DISTINCT B.ID,B.����ʱ��" & vbNewLine & _
                    "FROM ���˻�����ϸ A,���˻������� B" & vbNewLine & _
                    "WHERE A.��¼ID=B.ID And A.������Դ in (0,3) AND A.��ֹ�汾 IS NULL AND A.��¼���� =5 AND Instr(NVL(B.ǩ����,'QMR'),'/',1)=0 And B.����ǩ���� IS NULL AND A.��¼ID=[2] AND B.�ļ�ID=[1]"
            Call SQLDIY(gstrSQL)
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����Ƿ����δǩ��������", mlng�ļ�ID, Val(VsfData.TextMatrix(lngStart, mlngRecord)))
            '��¼��=0˵��û���ҵ�����
            If rsTemp.RecordCount = 0 Then
                RaiseEvent AfterRowColChange("û���ҵ���Ҫ����ǩ�������ݣ���ȷ�ϵ�ǰѡ��������Ƿ��Ѿ�ǩ��(��������ǩ�ͽ���ǩ��)����", True, mblnSign, mblnArchive)
                Exit Function
            End If
            '���ڷ������ݽ���ǩ��ʱ����Ҫ�Ա�����������н���ǩ��
            lngDemo = Val(VsfData.TextMatrix(lngStart, mlngDemo))
            If lngDemo > 0 Then
                '�϶��ҵĵ�����
                gstrSQL = "" & _
                    "SELECT DISTINCT B.ID,B.����ʱ��" & vbNewLine & _
                    "FROM ���˻�����ϸ A,���˻������� B" & vbNewLine & _
                    "WHERE A.��¼ID=B.ID And A.������Դ in (0,3) AND A.��ֹ�汾 IS NULL AND A.��¼���� =5 AND INSTR(NVL(B.ǩ����,'QMR'),'/',1)=0 And B.����ǩ���� IS NULL AND B.����ʱ�� between [2] And [3] AND B.�ļ�ID=[1]"
                Call SQLDIY(gstrSQL)
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����Ƿ����δǩ��������", mlng�ļ�ID, CDate(Format(VsfData.TextMatrix(lngStart, mlngActTime), "YYYY-MM-DD HH:mm")), CDate(Format(VsfData.TextMatrix(lngStart, mlngActTime), "YYYY-MM-DD HH:mm") & ":59"))
            End If
            
            mintVerify_Last = Val(IIf(VsfData.TextMatrix(lngStart, mlngSignLevel) = "", 9, Val(VsfData.TextMatrix(lngStart, mlngSignLevel)) + 1))
            
        End If
    End If
    
    '׼��ǩ��
    strSignTime = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    With rsTemp
        Do While Not .EOF
            str�д��� = ""
            blnSign = SignName(Val(!ID), Format(!����ʱ��, "yyyy-MM-dd HH:mm:ss"), strSignTime, bln��ǩ, str״̬, str�д���, blnExchange)
            If str�д��� <> "" Then
                str���� = str���� & vbCrLf & "����ʱ��=[" & Format(!����ʱ��, "yyyy-MM-dd HH:mm:ss") & "]" & str�д���
            End If
            If Not blnSign Then Exit Do
            If Not blnRefresh Then blnRefresh = blnSign
            .MoveNext
        Loop
    End With
    
    
    If blnRefresh And Not mblnVerify Then Call ShowMe(mfrmParent, mlng�ļ�ID, mlng����ID, mlng��ҳID, mlng����ID, mintӤ��, mstrPrivs, mblnEditable, mintҳ��)
    If str���� <> "" Then MsgBox "ǩ��ʱ�������´���" & str����, vbInformation, gstrSysName
    SignMe = blnRefresh
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub UnSignMe(Optional ByVal bln��ǩ As Boolean = False, Optional blnSingleCancel As Boolean = False)
    'blnSingleCancel �Ƿ���ȡ��
    Dim intPos As Integer
    Dim lngStart As Long                '��ʼ��
    Dim lngRecord As Long
    Dim blnOK As Boolean
    Dim strSignTime As String           'ǩ��ʱ��
    Dim strRecord As String, strSQLWhere As String
    Dim blnClear As Boolean             'ȡ��ǩ��ʱ�Ƿ�����ð汾�����ݻ��˵��ϴ�ǩ�����״̬
    Dim blnTrans As Boolean
    Dim strSQLTime() As String, strSQLSign() As String, strSQLCollect() As String
    Dim blnUnSign As Boolean, arrUnsign(), strUnsignID As String, strDays As String, strDate As String
    ReDim Preserve strSQLTime(1 To 1)
    ReDim Preserve strSQLSign(1 To 1)
    ReDim Preserve strSQLCollect(1 To 1)
    
    Dim clsSign As Object
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    '�������һ���Ǳ��˵�ǩ�������ݵ�ǰѡ�����ݵ�ǩ��ʱ�䣬����ȡ��ǩ��
    
    If mlng����ID = 0 Then Exit Sub
    
    '��Ҫ�Լ��
    '��ǰ��¼���¼�¼���˳�
    If FormatValue(VsfData.TextMatrix(VsfData.ROW, mlngRowCount)) = "" Then Exit Sub
    lngStart = GetStartRow(VsfData.ROW)
    lngRecord = Val(VsfData.TextMatrix(lngStart, mlngRecord))
    If lngRecord = 0 Then
        RaiseEvent AfterRowColChange("������¼������ȡ��ǩ����", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    '��ǰ��¼δǩ�����˳�
    If FormatValue(VsfData.TextMatrix(lngStart, mlngSigner)) = "" Then
        RaiseEvent AfterRowColChange("��ǰ��¼��δǩ����", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    '��ǩ����ǰ��¼δ��ǩ���˳���ƽǩ����ǰ��¼����ǩ���˳�
    intPos = InStr(1, FormatValue(VsfData.TextMatrix(lngStart, mlngSigner)), "/")
    If bln��ǩ Then
        If intPos = 0 Then
            RaiseEvent AfterRowColChange("��ǰ��¼δ��ǩ���޷�ִ��ȡ����ǩ������", True, mblnSign, mblnArchive)
            Exit Sub
        End If
    Else
        If intPos <> 0 Then
            RaiseEvent AfterRowColChange("��ǰ��¼����ǩ����ȡ����ǩ���ٲ�����", True, mblnSign, mblnArchive)
            Exit Sub
        End If
    End If
    '��ǰ��¼�����ǩ���˲��Ǳ������˳�
    '��������ҲҪǩ��,���ȥ������: And B.�������=0
    gstrSQL = "" & _
              " SELECT  A.��¼��,A.��¼ʱ��,A.��Ŀ����,B.ǩ����" & vbNewLine & _
              " FROM ���˻�����ϸ A,���˻������� B" & vbNewLine & _
              " WHERE A.��¼ID=B.ID AND B.�ļ�ID=[1] AND A.��¼ID=[2] AND A.��¼����=" & IIf(bln��ǩ, 15, 5) & vbNewLine & _
              " ORDER BY A.��Ŀ���� DESC"
    Call SQLDIY(gstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ǰ��¼�����ǩ���˲��Ǳ������˳�", mlng�ļ�ID, lngRecord)
    If rsTemp.RecordCount = 0 Then
        RaiseEvent AfterRowColChange("�޷��ҵ���" & IIf(bln��ǩ, "��ǩ", "ǩ��") & "�����ݣ����������ݱ仯δˢ�µ��£���ˢ�����ݺ����ԣ�", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    
    If bln��ǩ = False And InStr(1, NVL(rsTemp!ǩ����), "/") <> 0 Then
        RaiseEvent AfterRowColChange("��ǰ��¼����ǩ�����������ݱ仯δˢ�µ��£���ˢ�����ݡ�ȡ����ǩ���ٲ�����", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    
    If NVL(rsTemp!��¼��) <> gstrUserName Then
        RaiseEvent AfterRowColChange("���������ǩ���ˣ�����ִ�б�������", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    
    
    '100658,����,2016-10-20,���ӵ�����¼����ǩ��,��ǩ
     If blnSingleCancel = True Then
        strRecord = GetSelectRowRecordId(VsfData.ROW)
     Else
        strRecord = ""
     End If
     strSQLWhere = ""
     If strRecord <> "" Then
        strSQLWhere = " And B.id in (Select /*+ CARDINALITY(b 10) */  COLUMN_VALUE from  Table(f_Num2list([4])) b)"
     End If
    
    '��ȡ��������׼��ȡ��ǩ������ǩ(��¼ʱ�䲻Ϊ�ձ�ʾ�°�ǩ��;)
    '��������ҲҪǩ��,���ȥ������: And B.�������=0
    If Not IsNull(rsTemp!��¼ʱ��) Then
        gstrSQL = "" & _
                  " SELECT  A.��ĿID AS ֤��ID,A.��Ŀ����,B.����ʱ��,B.ID,B.ǩ����" & vbNewLine & _
                  " FROM ���˻�����ϸ A,���˻������� B" & vbNewLine & _
                  " WHERE A.��¼ID=B.ID AND B.�ļ�ID=[1] And A.��¼��=[2] And A.��¼ʱ��=[3] " & strSQLWhere & _
                  " AND A.��¼����=" & IIf(bln��ǩ, 15, 5) & _
                  " Order by B.����ʱ��"
                  
        Call SQLDIY(gstrSQL)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��������׼��ȡ��ǩ������ǩ", mlng�ļ�ID, gstrUserName, CDate(rsTemp!��¼ʱ��), strRecord)
    Else
        gstrSQL = "" & _
                  " SELECT  A.��ĿID AS ֤��ID,A.��Ŀ����,B.����ʱ��,B.ID,B.ǩ����" & vbNewLine & _
                  " FROM ���˻�����ϸ A,���˻������� B" & vbNewLine & _
                  " WHERE A.��¼ID=B.ID AND B.�ļ�ID=[1] And A.��¼��=[2] And A.��Ŀ����=[3] " & strSQLWhere & _
                  " AND A.��¼����=" & IIf(bln��ǩ, 15, 5) & _
                  " Order by B.����ʱ��"
        Call SQLDIY(gstrSQL)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��������׼��ȡ��ǩ������ǩ", mlng�ļ�ID, gstrUserName, CStr(rsTemp!��Ŀ����), strRecord)
    End If
    
    'ǩ���������޸ģ������޸ı������ǩ�������ȡ����ǩʱ��������ʾ�Ƿ�������ݵ����⣬��ǩ�Զ����ˣ�����ȡ����ʾ
    '--------------------
    'ѯ���Ƿ���Ҫ�������
'    If Not bln��ǩ Then
'        blnClear = (MsgBox("ȡ��ǩ��ʱ�Ƿ�ð汾�����ݻ��˵��ϴ�ǩ�����״̬��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes)
'    End If
    blnClear = True
    '--------------------
    arrUnsign = Array()
    strUnsignID = "": strDays = ""
    If bln��ǩ = True Then
        Do While Not rsTemp.EOF
            '81535:��ǩ���ܴ�������ʱ����޸�, ������ǩ������˵�ʱ���Ƿ�����Ѿ�ǩ���Ļ�������
            If IsDate(NVL(rsTemp!��Ŀ����)) Then
                If Format(rsTemp!����ʱ��, "yyyy-MM-dd HH:mm:ss") <> Format(rsTemp!��Ŀ����, "yyyy-MM-dd HH:mm:ss") Then
                    If ISCollectSigned(mlng�ļ�ID, Format(rsTemp!��Ŀ����, "YYYY-MM-DD"), Format(rsTemp!��Ŀ����, "HH:MM")) Then
                        ReDim Preserve arrUnsign(UBound(arrUnsign) + 1)
                        arrUnsign(UBound(arrUnsign)) = "��ǰ����ʱ�䣺" & Format(rsTemp!����ʱ��, "yyyy-MM-dd HH:mm") & vbTab & "�޸�ǰ����ʱ�䣺" & Format(rsTemp!��Ŀ����, "yyyy-MM-dd HH:mm")
                        strUnsignID = strUnsignID & "," & Val(rsTemp!ID)
                    ElseIf ISCollectSigned(mlng�ļ�ID, Format(rsTemp!����ʱ��, "YYYY-MM-DD"), Format(rsTemp!����ʱ��, "HH:MM")) Then
                        ReDim Preserve arrUnsign(UBound(arrUnsign) + 1)
                        arrUnsign(UBound(arrUnsign)) = "��ǰ����ʱ�䣺" & Format(rsTemp!����ʱ��, "yyyy-MM-dd HH:mm") & vbTab & "�޸�ǰ����ʱ�䣺" & Format(rsTemp!��Ŀ����, "yyyy-MM-dd HH:mm")
                        strUnsignID = strUnsignID & "," & Val(rsTemp!ID)
                    Else
                        gstrSQL = "Zl_���˻�������_����ʱ��(" & rsTemp!ID & ",to_date('" & rsTemp!��Ŀ���� & "','yyyy-MM-dd hh24:mi:ss'))"
                        strSQLSign(ReDimArray(strSQLSign)) = gstrSQL
                        
                        'ͬʱ���������������
                        '����Ҫ�������죬��Ϊ���ܴ��ڿ�����ܵ����ݣ��ҵ�ǰʱ��պ��ڵڶ�������
                        strDate = Format(rsTemp!����ʱ��, "yyyy-MM-dd")
                        If InStr(1, "," & strDays & ",", "," & Format(DateAdd("d", -1, CDate(strDate)), "yyyy-MM-dd") & ",") = 0 Then
                            gstrSQL = "ZL_��������_UPDATE(" & mlng�ļ�ID & ",'" & Format(DateAdd("d", -1, CDate(strDate)), "yyyy-MM-dd") & "')"
                            strSQLCollect(ReDimArray(strSQLCollect)) = gstrSQL
                            strDays = strDays & "," & Format(DateAdd("d", -1, CDate(strDate)), "yyyy-MM-dd")
                            If Left(strDays, 1) = "," Then strDays = Mid(strDays, 2)
                        End If
                        If InStr(1, "," & strDays & ",", "," & strDate & ",") = 0 Then
                            gstrSQL = "ZL_��������_UPDATE(" & mlng�ļ�ID & ",'" & strDate & "')"
                            strSQLCollect(ReDimArray(strSQLCollect)) = gstrSQL
                            strDays = strDays & "," & strDate
                            If Left(strDays, 1) = "," Then strDays = Mid(strDays, 2)
                        End If
                        strDate = Format(rsTemp!��Ŀ����, "yyyy-MM-dd")
                        If InStr(1, "," & strDays & ",", "," & Format(DateAdd("d", -1, CDate(strDate)), "yyyy-MM-dd") & ",") = 0 Then
                            gstrSQL = "ZL_��������_UPDATE(" & mlng�ļ�ID & ",'" & Format(DateAdd("d", -1, CDate(strDate)), "yyyy-MM-dd") & "')"
                            strSQLCollect(ReDimArray(strSQLCollect)) = gstrSQL
                            strDays = strDays & "," & Format(DateAdd("d", -1, CDate(strDate)), "yyyy-MM-dd")
                            If Left(strDays, 1) = "," Then strDays = Mid(strDays, 2)
                        End If
                        If InStr(1, "," & strDays & ",", "," & strDate & ",") = 0 Then
                            gstrSQL = "ZL_��������_UPDATE(" & mlng�ļ�ID & ",'" & strDate & "')"
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
            If MsgBox("����ǩʱ�޸��˲������ݵ�ʱ�䣬����Щ�����еĲ������ݶ�Ӧ�Ļ��������Ѿ�ǩ������Щ���ݽ����ܽ��л��ˣ��������Ƿ������" & vbCrLf & _
                "�ǣ����������������ݽ����ܱ�����" & vbCrLf & "����ֹ������ǩ����" & vbCrLf & vbCrLf & _
                "���ܻ��˵�������Ϣ���£�" & vbCrLf & Join(arrUnsign, vbCrLf), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
    End If
    
    mstrVerify = ""
    If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
    Do While Not rsTemp.EOF
        If (bln��ǩ = False And InStr(1, NVL(rsTemp!ǩ����), "/") = 0) Or bln��ǩ = True Then
            blnUnSign = InStr(1, "," & strUnsignID & ",", "," & rsTemp!ID & ",") = 0
            If blnUnSign = True Then
                If NVL(rsTemp!֤��ID, 0) > 0 Then
                    '����ǩ����֤��ֻ��֤һ��
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
                            RaiseEvent AfterRowColChange("����ǩ������δ����ȷ�������˲������ܼ�����", True, mblnSign, mblnArchive)
                            Exit Sub
                        End If
                    End If
                End If
                
                'ȡ��ǩ��
                gstrSQL = "ZL_���˻�������_UNSIGNNAME("
                gstrSQL = gstrSQL & mlng�ļ�ID & ","
                gstrSQL = gstrSQL & "To_Date('" & Format(rsTemp!����ʱ��, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd hh24:mi:ss')," & IIf(blnClear, "1", "0") & ")"
                strSQLTime(ReDimArray(strSQLTime)) = gstrSQL
                'Call zlDatabase.ExecuteProcedure(gstrSQL, "ִ��ȡ��ǩ��")
                
                If InStr(1, mstrVerify & ",", "," & rsTemp!ID & ",") = 0 Then
                    mstrVerify = mstrVerify & "," & rsTemp!ID
                End If
            End If
        End If
        rsTemp.MoveNext
    Loop
    
    '����ǩ����ǰ�����½���������
    gcnOracle.BeginTrans
    blnTrans = True
    For intPos = 1 To UBound(strSQLTime)
        If strSQLTime(intPos) <> "" Then
            Call zlDatabase.ExecuteProcedure(strSQLTime(intPos), "ִ��ȡ��ǩ��")
        End If
    Next intPos
    '��ǩʱ�޸�������ʱ�䣬����ʱ��Ҫͬ������
    For intPos = 1 To UBound(strSQLSign)
        If strSQLSign(intPos) <> "" Then
            Call zlDatabase.ExecuteProcedure(strSQLSign(intPos), "����ʱ������")
        End If
    Next intPos
    '������������
    For intPos = 1 To UBound(strSQLCollect)
        If strSQLCollect(intPos) <> "" Then
            Call zlDatabase.ExecuteProcedure(strSQLCollect(intPos), "������������")
        End If
    Next intPos
    
    'ȡ����ǩ�����½��������У���Ϊǩ�������޸����ݣ����Բ�ǣ�����������仯�����⡣
    '��ǩ�ǿ���ԭ������5����ǩʱ��Ϊ3�У�������ǩ����Ҫ��ԭ���ݺ�����
    If Not PreseData(bln��ǩ) Then GoTo ErrHand:
    
    gcnOracle.CommitTrans
    blnTrans = False
    mstrVerify = ""
    'ˢ������
    Call ShowMe(mfrmParent, mlng�ļ�ID, mlng����ID, mlng��ҳID, mlng����ID, mintӤ��, mstrPrivs, mblnEditable, mintҳ��)
    Exit Sub
ErrHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Sub



Private Function SignName(ByVal lngRecordId As Long, ByVal strStart As String, ByVal strSignTime As String, ByVal bln��ǩ As Boolean, _
    str״̬ As String, Optional str���� As String, Optional ByVal blnExchange As Boolean = False) As Boolean
    '******************************************************************************************************************
    '����:
    '
    '
    '******************************************************************************************************************
    Dim oSign As cTendSign
    Dim strSource As String             '��ǩԴ���ݴ�
    Dim lngLoop As Long
    Dim rs As New ADODB.Recordset
    Dim strChangeTime As String
    On Error GoTo ErrHand
    
    '��ʼ����
    '------------------------------------------------------------------------------------------------------------------
    strSource = ""
    
    '��ȡҪǩ��������(��������ҲҪǩ��,���ȥ������: And B.�������=0)
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = " Select a.��¼����,a.��Ŀ����,a.��Ŀ���,a.��Ŀ����,a.��Ŀ����,a.��¼����,a.��Ŀ��λ,a.��¼���,a.���²�λ,a.��¼���,a.���Ժϸ�,a.δ��˵��,a.��¼��,a.��¼ʱ��  " & _
              " From ���˻�����ϸ a,���˻������� b,���˻����ļ� c " & _
              " Where a.��¼id=b.ID And b.�ļ�ID=c.ID AND MOD(A.��¼����,10)<>5 And a.��ֹ�汾 Is Null And C.ID=[1] And b.����ʱ��=[2]" & _
              " Order by a.��Ŀ���"
    Call SQLDIY(gstrSQL)
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҪǩ��������", mlng�ļ�ID, CDate(strStart))
    If rs.BOF = False Then
        Do While Not rs.EOF
            For lngLoop = 0 To rs.Fields.Count - 1
                strSource = strSource & CStr(zlCommFun.NVL(rs.Fields(lngLoop).Value, ""))
            Next
            rs.MoveNext
        Loop
    End If
    'Debug.Print "��ʼǩ����" & Now & vbCrLf & strSource
    If strSource = "" Then
        RaiseEvent AfterRowColChange("��ǰû����Ҫǩ������Ϣ��", True, mblnSign, mblnArchive)
        Exit Function
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    On Error Resume Next
    Err = 0
    '81535:��ȡԭ�����ݵ�ʱ��(��ǩ�����޸����ݷ���ʱ��)
    strChangeTime = ""
    If bln��ǩ = True And Not mobjVerify Is Nothing Then
        If mobjVerify.Count > 0 Then
            If Not Format(mobjVerify("_" & lngRecordId), "YYYY-MM-DD HH:mm") = Format(strStart, "YYYY-MM-DD HH:mm") Then
                strChangeTime = Format(mobjVerify("_" & lngRecordId), "YYYY-MM-DD HH:mm:ss")
            End If
        End If
    End If
    '76223:������,2012-09-13,����ǩ�����ʱ�����Ϣ
    '43588:������,2012-09-13,��Ӽ�¼����ǩģʽ
    Set oSign = frmTendFileSign.ShowMe(Me, mstrPrivs, mlng�ļ�ID, mlng����ID, mintVerify_Last, strSource, bln��ǩ, str״̬, str����, mintSignMode, blnExchange)
    On Error GoTo ErrHand
    
    If Not oSign Is Nothing Then
        gstrSQL = "ZL_���˻�������_SIGNNAME("
        gstrSQL = gstrSQL & mlng�ļ�ID & ","
        gstrSQL = gstrSQL & "To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss')," & IIf(bln��ǩ, 1, 0) & ","
        gstrSQL = gstrSQL & "'" & oSign.���� & "',"
        gstrSQL = gstrSQL & "'" & oSign.ǩ����Ϣ & "'," & oSign.ǩ������ & ","
        gstrSQL = gstrSQL & oSign.֤��ID & ","
        gstrSQL = gstrSQL & oSign.ǩ����ʽ & ",'" & oSign.ʱ��� & "'," & IIf(blnExchange, 1, 0) & ",'" & oSign.ʱ�����Ϣ & "',"
        gstrSQL = gstrSQL & "To_Date('" & strSignTime & "','yyyy-mm-dd hh24:mi:ss'),'" & strChangeTime & "')"
        
        Call zlDatabase.ExecuteProcedure(gstrSQL, "ִ��ǩ��")
        SignName = True
    End If
    
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function PreseData(ByVal blnTrue As Boolean) As Boolean
'���ܣ�ȡ����ǩ���ݺ����½���������
    Dim rsTemp As New ADODB.Recordset
    Dim str���� As String, intPos As Long
    Dim lngRow As Long, lngCol As Long, lngRecord As Long, lngMutilRow As Long, strTime As String
    Dim arrData, arrMutilRow
    Dim strSQLData() As String
    Dim blnSave As Boolean, i As Long
    ReDim Preserve strSQLData(1 To 1)
    On Error GoTo ErrHand
    
    If Left(mstrVerify, 1) = "," Then mstrVerify = Mid(mstrVerify, 2)
    If mstrVerify = "" Or blnTrue = False Then GoTo ErrEnd
    
   
    str���� = mstrSQL����
    
    mstrSQL = "Select /*+ RULE */ '' AS ����,to_char(����ʱ��,'yyyy-MM-dd hh24:mi:ss') AS ����ʱ��,'' AS ѡ��,to_char(����ʱ��,'YYYY') AS ���," & Mid(mstrSQL��, 12) & vbCrLf & _
                " From (Select ��¼���,ʱ�� as ����,����ʱ��," & Mid(mstrSQL��, 2) & vbCrLf & _
                "        From (Select nvl(c.��¼���,0) ��¼���,l.����ʱ��," & Mid(mstrSQL��, 2) & vbCrLf & _
                "               From ���˻������� l, ���˻�����ϸ c,���˻����ļ� f,���˻����ӡ p," & vbCrLf & _
                "       (SELECT COLUMN_VALUE FROM TABLE(CAST(F_NUM2LIST([6]) AS ZLTOOLS.T_NUMLIST))) G" & vbNewLine & _
                "               Where l.ID=p.��¼ID And l.Id = c.��¼id And l.�ļ�ID+0=f.ID+0 And f.ID=p.�ļ�ID " & _
                "               And c.��ֹ�汾 Is Null And MOD(c.��¼����,10)<>5  " & _
                "               And f.id=[1] And f.����id = [2] And f.��ҳid = [3] And Nvl(f.Ӥ��,0)=[4] And l.ID=G.COLUMN_VALUE)" & _
                IIf(str���� <> "", "Where " & str����, "") & _
                "       Group By ����, ʱ��, ����ʱ��,��¼���," & IIf(mbln��ʿ = True, "��ʿ,", "��ʿL,") & "ǩ����,ǩ��ʱ��" & _
                                "       Order By ����ʱ��,��¼���," & IIf(mbln��ʿ = True, "��ʿ,", "��ʿL,") & "ǩ����,ǩ��ʱ��)"
     Call SQLDIY(mstrSQL)
     Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "����Ƿ����δǩ��������", mlng�ļ�ID, mlng����ID, mlng��ҳID, mintӤ��, mintҳ��, mstrVerify)
     If rsTemp.RecordCount > 0 Then
        blnSave = False: arrMutilRow = Array()
        Set vsTest.DataSource = rsTemp
        For lngRow = vsTest.FixedRows To vsTest.Rows - 1
            lngMutilRow = 0
            lngRecord = Val(vsTest.TextMatrix(lngRow, mlngRecord))
            If lngRecord <> 0 Then
                blnSave = True
                strTime = vsTest.TextMatrix(lngRow, mlngActTime)
                '������ܴ���:����������ʱ�����������ϸ������
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
                        '׼����ֵ
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
                    '----�˴���Ҫ���������ܵ�����
                    lngMutilRow = 0
                    '���������ϸ����������
                    For i = 1 To UBound(arrMutilRow)
                        lngMutilRow = lngMutilRow + arrMutilRow(i)
                    Next i
                    '��������������������ڷ�����ϸ��������+1(1ΪĬ�ϵ���������),��������������Ϊ׼,��������ϸ������+1Ϊ׼
                    If lngMutilRow + 1 > Val(arrMutilRow(0)) Then
                        lngMutilRow = lngMutilRow + 1
                    Else
                        lngMutilRow = Val(arrMutilRow(0))
                    End If
                    arrMutilRow = Array()
                    'һ�н���ʱ��������ӡ��������
                    If Val(vsTest.TextMatrix(lngRow, mlngRowCount)) <> lngMutilRow Then
                        gstrSQL = "ZL_���˻����ӡ_UPDATE(" & mlng�ļ�ID & ",to_date('" & strTime & "','yyyy-MM-dd hh24:mi:ss')" & "," & lngMutilRow & ")"
                        strSQLData(ReDimArray(strSQLData)) = gstrSQL
                    End If
                End If
            End If
        Next
     End If
     
    'ִ�й���
    For intPos = 1 To UBound(strSQLData)
        If strSQLData(intPos) <> "" Then
            Call zlDatabase.ExecuteProcedure(strSQLData(intPos), "������ӡ��������")
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
    Call ShowMe(mfrmParent, mlng�ļ�ID, mlng����ID, mlng��ҳID, mlng����ID, mintӤ��, mstrPrivs, mblnEditable, mintҳ��)
End Function

Public Function SaveME() As Boolean
    If Not CheckData Then Exit Function
    If Not SaveData Then Exit Function
    
    mblnShow = False
    Call InitCons
    SaveME = True
    
    Call ShowMe(mfrmParent, mlng�ļ�ID, mlng����ID, mlng��ҳID, mlng����ID, mintӤ��, mstrPrivs, mblnEditable, mintҳ��)
End Function

Public Function ShowMe(ByVal frmParent As Form, ByVal lngFileID As Long, ByVal lngPatiID As Long, ByVal lngPageId As Long, _
    ByVal lngDeptID As Long, ByVal intBaby As Integer, Optional ByVal strPrivs As String, Optional ByVal blnEditable As Boolean = True, _
    Optional ByVal intҳ�� As Integer = -1, Optional ByVal blnClear As Boolean = True) As Boolean
    '******************************************************************************************************************
    '���ܣ� ��ʾ�����¼�ļ�����
    '������ frmParent           �ϼ��������
    '       lngPatiID           ����id
    '       lngPageID           ��ҳid
    '       lngDeptID           Ҫ��ʾ�����¼�Ŀ���
    '       intBaby             Ӥ����־
    '       blnEditable         ���Ϊ��,˵������Ϊ��ѯ�Ӵ�����ʹ��,ȡ����༭��صĹ���
    '       blnClear            ���Ϊ��,���mrsDataMap��¼��;����ҳʱӦ����,�����û��޸ĵ������Ա���ʾ������ʹ��
    '���أ� ��
    '******************************************************************************************************************
    Dim lngRow As Long
    Dim rsTemp As New ADODB.Recordset
    Dim strtmp As String
    On Error GoTo ErrHand
    Err = 0
    
    mblnInit = False
    If mblnChange Then
        If MsgBox("��ǰ���˵����ݻ�δ���棬�㡰�ǡ����б��棬�㡰�񡱽����������޸ģ�", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Call VsfData_EnterCell
            '60218:������,2013-04-22,���CheckData����Ȼ��û���л�ҳ��������޷���������
            If CheckData Then Call SaveData
        End If
    End If
    
    mblnGroupNew = False
    mblnGroupApp = False
    mstrGroupRow = ""
    mblnClear = blnClear
    mint��ʼҳ�� = 1
    mintҳ�� = intҳ��
    mlng�ļ�ID = lngFileID
    mlng����ID = lngPatiID
    mlng��ҳID = lngPageId
    mlng����ID = lngDeptID
    mintӤ�� = intBaby
    mstrPrivs = strPrivs
    'mblnBlowup = (zlDatabase.GetPara("�����ļ���ʾģʽ", glngSys, 1255, 0) = 1)
    UserControl.Font = IIf(mblnBlowup = True, 12, 9)
    Set mfrmParent = frmParent
    
    mintNORule = Val(zlDatabase.GetPara("�����ļ�ҳ�����", glngSys, 1255, 0))
    mintPreDays = Val(zlDatabase.GetPara("����¼�뻤����������", glngSys, 1255, "1"))
    mstrMaxDate = Format(DateAdd("d", mintPreDays, zlDatabase.Currentdate), "yyyy-MM-dd HH:mm")
    mintCollectDef = Val(zlDatabase.GetPara("С��ȱʡ��ʽ", glngSys, 1255))
    mintPageSpan = Val(zlDatabase.GetPara("��ҳ����ֻ��ʾ�ڵ�һҳ", glngSys, 1255))
    '68739:������,2014-1-2,���"С���ʶ��ɫ"
    mlngCollectColor = Val(zlDatabase.GetPara("С���ʶ��ɫ", glngSys, 1255, "255"))
    
    '43588:������,2012-09-13,��Ӽ�¼����ǩģʽ
    strtmp = Val(zlDatabase.GetPara("��¼����ǩģʽ", glngSys, 1255))
    If Val(strtmp) >= 0 And Val(strtmp) <= 1 Then
        mintSignMode = CInt(Val(strtmp))
    Else
        mintSignMode = 0
    End If
    
    '--58414,������,2013-04-10,��ӻ�ʿ��ǩ������ʾģʽ
    mlngSingerType = Val(zlDatabase.GetPara("��ʿ��ǩ������ʾģʽ", glngSys, 1255, "2"))
    If InStr(1, ",0,1,2,3,", "," & mlngSingerType & ",") = 0 Then mlngSingerType = 2
    
    If mrsItems.State = 0 Then
        Call InitMenuBar
        Call InitEnv            '��ʼ������
        cbsThis.ActiveMenuBar.Visible = False
        cbsThis.RecalcLayout
    End If
    
    Call InitVariable
    Call InitCons
    
    If Not ReadStruDef Then Exit Function
    Call zlRefresh
    mblnInit = True
    mblnEditable = blnEditable And Not gblnMoved And Not mblnArchive
    
    '--48659:������,2012-09-14,����ֶ�'˵��'
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
    '����:���������С
    '���:bytSize��0-С(ȱʡ)��1-��
    '����:������
    '����:2012-06-19 15:16
    '����:50807
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mblnBlowup = IIf(bytSize = 1, True, False)
    Call ReSetFontSize
    If mintҳ�� = -1 Or mblnInit = False Then Exit Sub
    If Not DataMap_Save Then Exit Sub
    '���²�ѯSQL
    '������ȡ����
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
    '����:���������С
    '���:bytSize��0-С(ȱʡ)��1-��
    '����:������
    '����:2012-06-20 15:15:00
    '����:50807
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCtrl As Control
    Dim CtlFont As StdFont
    Dim bytFontSize As Byte
    bytFontSize = BlowUp(9)
    
    UserControl.FontSize = bytFontSize
    UserControl.FontName = "����"
    
    Set CtlFont = cbsThis.Options.Font
    If CtlFont Is Nothing Then
        Set CtlFont = UserControl.Font
    End If
    CtlFont.Size = bytFontSize
    Set cbsThis.Options.Font = CtlFont
    
    lblPage.FontSize = bytFontSize
    lblPage.Left = 30
    txtPage.FontSize = bytFontSize
    txtPage.Height = TextHeight("��") + TextHeight("��") / 3
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
    'ҳ���л�ǰ��飺����ʱ����ȷ����������������ڱ���ʱ�Ͳ����ټ������ҳ��������ˣ�����������¼��ʱ�Ѿ������˼�飬�˴��Թ���
    
    '���ر༭�ؼ�
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
        '��ǩ�����޸�ʱ�䣬�˴�����޸���ʱ������ݶ�Ӧ�Ļ��������Ƿ��Ѿ�ǩ��
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
        mrsCellMap.Filter = "ҳ��=" & mintҳ�� & " And �к�=" & lngRow & " And �к�>" & mlngTime
        If mrsCellMap.RecordCount = 0 And Val(VsfData.TextMatrix(lngRow, mlngRecord)) > 0 Then
            mrsCellMap.Filter = "ҳ��=" & mintҳ�� & " And �к�=" & lngRow & " And �к�>=" & mlngDate
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
                    RaiseEvent AfterRowColChange("�벹������ʱ�䣡", True, mblnSign, mblnArchive)
                    CheckFlip = False
                    Exit Function
                Else
                    '���ڲ�Ϊ�ս�������ڵĺϷ���
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
    '����������ӷ������ݣ�����������Ѿ�ǩ������ʾ��ֻ������ԭ�����������ķ������ݣ��������ķ������������Ѿ���飩
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
                mrsCellMap.Filter = "ҳ��=" & mintҳ�� & " And �к�=" & lngRow & " And �к�>" & mlngTime
                If mrsCellMap.RecordCount > 0 Then
                    lngEditCol = 0
                    If CheckCollectIsData(lngRow, 1, lngEditCol) = True Then
                        If ISCollectSigned(mlng�ļ�ID, Format(strDate, "YYYY-MM-DD"), Format(strDate, "HH:MM")) Then
                            VsfData.ROW = lngRow: VsfData.COL = lngEditCol
                            strInfo = "�������ķ�����������Ӧ�Ļ�����������ǩ����������������µĻ��������ݣ�"
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
    '�������
    
    '����޸������ݶ�����ʱ�䲻ȫ����ʾ�����ݺϷ�����¼��ʱ�Ѿ���飩
'    Call OutputRsData(mrsCellMap)
'    Call OutputRsData(mrsDataMap)
    If Not DataMap_Save Then Exit Function
    
    '�������ǩģʽ,������ѡ�����Ƿ���ڲ�����ǩ�����
    If mblnVerify Then
        mstrVerify = ""
        Set mobjVerify = New Collection
        mintVerify_Last = δ����
        '��ǩ��������������
        For lngPage = mint��ʼҳ�� To mint����ҳ
            mrsDataMap.Filter = "ҳ��=" & lngPage
            Do While Not mrsDataMap.EOF
                If NVL(mrsDataMap!ѡ��, 0) = flexTSChecked Then
                    mstrVerify = mstrVerify & "," & mrsDataMap!��¼ID
                    mobjVerify.Add Format(mrsDataMap!����ʱ��, "YYYY-MM-DD HH:mm:ss"), "_" & mrsDataMap!��¼ID
                    If IsNull(mrsDataMap!ǩ������) Then
                        intLevel = NVL(mrsDataMap!ǩ������, δ����)
                    Else
                        intLevel = Val(mrsDataMap!ǩ������) + 1
                    End If
                    If mintVerify < intLevel Then mintVerify_Last = intLevel
                End If
                mrsDataMap.MoveNext
            Loop
        Next
        mrsDataMap.Filter = 0
        
        If mstrVerify = "" Then
            RaiseEvent AfterRowColChange("����Ҫѡ��һ�����ݲ��������ǩ������", True, mblnSign, mblnArchive)
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
    Dim rsTime As New ADODB.Recordset, rsTimeCur As New ADODB.Recordset '���ݷ���ʱ��䶯
    Dim strFileds As String, strValues As String
    Dim strRelationNO As String
    ReDim Preserve strSQL(1 To 1)
    ReDim Preserve strSQLTime(1 To 1) '����ʱ��䶯SQL����
    ReDim Preserve strCollectSQL(1 To 1) 'С������SQL��С�����ݷ�������ڽ��д���
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    
    strFileds = "ID," & adDouble & ",18|ʱ��," & adDate & ",20|����ʱ��," & adDate & ",20|���," & adInteger & ",1"
    Call Record_Init(rsTime, strFileds)
    Call Record_Init(rsTimeCur, strFileds)
    'ͬ�ж���ѭ�����ã�ZL_���˻�������_UPDATE
    '��һ��ǰ���ã�
    '   1��ZL_���˻�������_SYNCHRO��ͬ�����ݵ����µ��뻤���¼���У���Ҫ��¼ɾ������ϸID��
    '   2��ZL_���˻����ӡ_UPDATE����ɴ�ӡ���ݽ���
    'ɾ����Ŀ���¼��ɾ����Ҳ��Ҫ��¼
    '�޸����ݵ�ͬ���ͽ��������ݶ�Ӧ��������ʱ�䱣�浽mrsCellMap��
    
'    objStream.WriteLine (Now & "��������SQL")
    intAllow = IIf(InStr(mstrPrivs, "���˻����¼") > 0, 1, 0)
    strCurrDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
    With mrsCellMap
        '����Ч���ݹ��˳���:��¼ID>0����ʷ����+��������Ч����
        .Filter = "��¼ID>0 or (��¼ID=0 And ɾ��=0)"
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            '77436:LPF,��ǩʱֻ����Ҫ��ǩ������(��ʼ��ѡ��¼�޸����ݣ����ܺ�����ȡ���˹�ѡ)
            If (InStr(1, "," & mstrVerify & ",", "," & NVL(!��¼ID, 0) & ",") <> 0 And mblnVerify = True) Or mblnVerify = False Then
                'Or (intPage = !ҳ�� And NVL(!��ʼ�к�, !�к�) = intStarRow) �ⲻ������Ҫ��Է�����ܣ���Ϊ���ݼ�¼ֻ��һ��������ϸ�ж���
                If Not ((intRow = !�к� And intPage = !ҳ��) Or (intPage = !ҳ�� And NVL(!��ʼ�к�, !�к�) = intStarRow)) Then
endWork:
                    If intRow > 0 Then
                        mrsDataMap.Filter = "ҳ��=" & intPage & " And �к�=" & intRow
                        If mrsDataMap.RecordCount <> 0 Then
                            blnDel = (mrsDataMap!ɾ�� = 1)
                            intUsedRows = Val(Split(NVL(mrsDataMap!���� & "|"), "|")(0))
                        Else
                            mrsDataMap.Filter = 0
                            intUsedRows = 1
                            RaiseEvent AfterRowColChange("��" & intRow & "�е������ڲ��������¼���β������貢������Ȼ������¼�����ݣ�лл��", True, mblnSign, mblnArchive)
                            Exit Function
                        End If
                        mrsDataMap.Filter = 0
                    End If
    
                    If blnSaved Then
                        '��ɴ�ӡ���ݽ���
    '                    �ļ�ID_IN IN ���˻����ӡ.�ļ�ID%TYPE,
    '                    ����ʱ��_IN IN ���˻����ӡ.����ʱ��%TYPE,
    '                    ����_IN IN ���˻����ӡ.����%TYPE,
    '                    ɾ��_IN Number:=0
                        gstrSQL = "ZL_���˻����ӡ_UPDATE(" & mlng�ļ�ID & ",to_date('" & strDatetime & "','yyyy-MM-dd hh24:mi:ss')," & intUsedRows & "," & IIf(blnDel, "1", "0") & ")"
                        strSQL(ReDimArray(strSQL)) = gstrSQL
                        
                        'ֻҪ�޸Ĺ�����,��Ȼ��ִ�д�ӡ����,�����������л������ڵĴ���
                        If strDate <> "" And .EOF Then
                            strLastDate = strDate
                            
                            'ͬ����������Ļ���(ҹ��,ȫ����ܿ���Ĵ���)
                            If InStr(1, "," & strDays & ",", "," & Format(DateAdd("d", -1, strDate), "yyyy-MM-dd") & ",") = 0 Then
                                strDays = strDays & "," & Format(DateAdd("d", -1, strDate), "yyyy-MM-dd")
                                gstrSQL = "ZL_��������_UPDATE(" & mlng�ļ�ID & ",'" & Format(DateAdd("d", -1, strDate), "yyyy-MM-dd") & "')"
                                strCollectSQL(ReDimArray(strCollectSQL)) = gstrSQL
                            End If
                            
                            If InStr(1, "," & strDays & ",", "," & strDate & ",") = 0 Then
                                strDays = strDays & "," & strDate
                                gstrSQL = "ZL_��������_UPDATE(" & mlng�ļ�ID & ",'" & strDate & "')"
                                strCollectSQL(ReDimArray(strCollectSQL)) = gstrSQL
                            End If
                            
                            strTemp = Format(DateAdd("d", 1, CDate(strDate)), "yyyy-MM-dd")
                            If InStr(1, "," & strDays & ",", "," & strTemp & ",") = 0 Then
                                strDays = strDays & "," & strTemp
                                gstrSQL = "ZL_��������_UPDATE(" & mlng�ļ�ID & ",'" & strTemp & "')"
                                strCollectSQL(ReDimArray(strCollectSQL)) = gstrSQL
                            End If
                        End If
                        
                        blnSaved = False
                        If .EOF Then Exit Do
                    End If
                    
                    '����ֵ
                    intPage = !ҳ��
                    intRow = !�к�
                    intStarRow = NVL(!��ʼ�к�, !�к�)
                    strDate = ""
                    strDatetime = ""
                    lngRecord = NVL(!��¼ID, 0)
                End If
                
                If !�к� = mlngDate Then
                    If NVL(!����, 0) = 1 Then
                        arrCollect = Split(!����, ";")
                        strDatetime = arrCollect(3)
                    '    �ļ�ID_IN IN ���˻�������.�ļ�ID%TYPE,
                    '    ����ʱ��_IN IN ���˻�������.����ʱ��%TYPE,
                    '    �������_IN IN ���˻�������.�������%TYPE,
                    '    �����ı�_IN IN ���˻�������.�����ı�%TYPE,
                    '    ���ܱ��_IN IN ���˻�������.���ܱ��%TYPE,
                    '    ɾ��_IN Number:=0
                        gstrSQL = "ZL_���˻�������_COLLECT(" & mlng�ļ�ID & ",to_date('" & arrCollect(3) & "','yyyy-MM-dd hh24:mi:ss')," & _
                                Val(arrCollect(1)) & ",'" & arrCollect(0) & "'," & Val(arrCollect(2)) & ",'" & arrCollect(4) & "','" & arrCollect(5) & "'," & !ɾ�� & ")"
                        strSQL(ReDimArray(strSQL)) = gstrSQL
                        blnSaved = True
                    Else
                        strDate = NVL(!����)
                        If strDate <> "" Then
                            If mblnDateAd Then
                                '69335:������,2014-1-7,�����ڸ�ʽ����(���ڲ�λ������)
                                strYear = NVL(!��λ)
                                If InStr(1, "|" & mstrYears & "|", "|" & strYear & "|") <> 0 Then
                                    strDate = strYear & "-" & ToStandDate(strDate)
                                Else
                                    RaiseEvent AfterRowColChange("��" & !�к� & "�е�[���]���ݴ������¼���β������貢������Ȼ������¼�����ݣ�лл��", True, mblnSign, mblnArchive)
                                    Exit Function
                                End If
    '                            strDate = Mid(strCurrDate, 1, 5) & ToStandDate(strDate)
    '                            '����Ƿ����༭֮ǰ��ʱ��(һ���µ�����)
    '                            If CDate(strDate) > CDate(Mid(strCurrDate, 1, 10)) And Mid(strCurrDate, 6, 2) = "01" And Mid(strDate, 6, 2) = "12" Then
    '                                strDate = DateAdd("yyyy", -1, CDate(strDate))
    '                            End If
                            Else
                                strDate = Format(strDate, "yyyy-MM-dd")
                            End If
                        End If
                        
                        '���������޸�ʱ��Ĵ���
                        If lngRecord <> 0 And strDate <> "" Then
                            mrsDataMap.Filter = "ҳ��=" & !ҳ�� & " And �к�=" & !�к�
                            If mrsDataMap.RecordCount > 0 Then
                                strActTime = Format(mrsDataMap.Fields(cControlFields + mlngActTime - VsfData.FixedCols).Value, "YYYY-MM-DD")
                                If Format(strActTime, "YYYY-MM-DD") <> Format(strDate, "YYYY-MM-DD") Then
                                    '����ͬʱ��������:��Ϊ���ܴ��ڿ�����ܵ����ݣ��ҵ�ǰʱ��պ��ڵڶ�������
                                    If InStr(1, "," & strDays & ",", "," & Format(DateAdd("d", -1, strActTime), "yyyy-MM-dd") & ",") = 0 Then
                                        strDays = strDays & "," & Format(DateAdd("d", -1, strActTime), "yyyy-MM-dd")
                                        gstrSQL = "ZL_��������_UPDATE(" & mlng�ļ�ID & ",'" & Format(DateAdd("d", -1, strActTime), "yyyy-MM-dd") & "')"
                                        strCollectSQL(ReDimArray(strCollectSQL)) = gstrSQL
                                    End If
                                    '����
                                    If InStr(1, "," & strDays & ",", "," & strActTime & ",") = 0 Then
                                        strDays = strDays & "," & strActTime
                                        gstrSQL = "ZL_��������_UPDATE(" & mlng�ļ�ID & ",'" & strActTime & "')"
                                        strCollectSQL(ReDimArray(strCollectSQL)) = gstrSQL
                                    End If
                                End If
                            End If
                        End If
                    End If
                    If strLastDate = "" Then strLastDate = strDate
                    
                    If strLastDate <> strDate Then
                        'ֻҪ�޸Ĺ�����,��Ȼ��ִ�д�ӡ����,�����������л������ڵĴ���
                        'ͬ����������Ļ���(ҹ��,ȫ����ܿ���Ĵ���)
                        If InStr(1, "," & strDays & ",", "," & Format(DateAdd("d", -1, strLastDate), "yyyy-MM-dd") & ",") = 0 Then
                            strDays = strDays & "," & Format(DateAdd("d", -1, strLastDate), "yyyy-MM-dd")
                            gstrSQL = "ZL_��������_UPDATE(" & mlng�ļ�ID & ",'" & Format(DateAdd("d", -1, strLastDate), "yyyy-MM-dd") & "')"
                            strCollectSQL(ReDimArray(strCollectSQL)) = gstrSQL
                        End If
                        
                        If InStr(1, "," & strDays & ",", "," & strLastDate & ",") = 0 Then
                            strDays = strDays & "," & strLastDate
                            gstrSQL = "ZL_��������_UPDATE(" & mlng�ļ�ID & ",'" & strLastDate & "')"
                            strCollectSQL(ReDimArray(strCollectSQL)) = gstrSQL
                        End If
                        
                        strTemp = Format(DateAdd("d", 1, CDate(strLastDate)), "yyyy-MM-dd")
                        If InStr(1, "," & strDays & ",", "," & strTemp & ",") = 0 Then
                            strDays = strDays & "," & strTemp
                            gstrSQL = "ZL_��������_UPDATE(" & mlng�ļ�ID & ",'" & strTemp & "')"
                            strCollectSQL(ReDimArray(strCollectSQL)) = gstrSQL
                        End If
                        strLastDate = strDate
                    End If
                ElseIf !�к� = mlngTime Then
                    strTime = NVL(!����)
                    If strDate = "" Then strDate = Mid(strCurrDate, 1, 10)
                    strDatetime = strDate & " " & strTime & ":00"
                    
                    '����������ݣ�����ʱ����ͨ����������ֻ������+
                    If Val(NVL(!��λ)) >= 1 Then
                        'strDatetime = Mid(strDatetime, 1, 17) & String(2 - Len(!��λ), "0") & Val(!��λ) - 1
                        strDatetime = DateAdd("S", Val(!��λ) - 1, CDate(strDatetime))
                    End If
                    
                    If lngRecord <> 0 Then
                        '���·���ʱ��
    '                    gstrSQL = "Zl_���˻�������_����ʱ��(" & lngRecord & ",to_date('" & strDatetime & "','yyyy-MM-dd hh24:mi:ss'))"
    '                    strSQLTime(ReDimArray(strSQLTime)) = gstrSQL
                        mrsDataMap.Filter = "ҳ��=" & !ҳ�� & " And �к�=" & !�к�
                        If mrsDataMap.RecordCount = 0 Then
                            RaiseEvent AfterRowColChange("��" & !�к� & "�е������ڲ��������¼���β������貢������Ȼ������¼�����ݣ�лл��", True, mblnSign, mblnArchive)
                            Exit Function
                        End If
                        strActTime = Format(mrsDataMap.Fields(cControlFields + mlngActTime - VsfData.FixedCols).Value, "YYYY-MM-DD HH:mm:ss")
                        strValues = lngRecord & "|" & Format(strDatetime, "YYYY-MM-DD HH:mm:ss") & "|" & Format(strActTime, "YYYY-MM-DD HH:mm:ss") & "|0"
                        Call Record_Update(rsTime, "ID|ʱ��|����ʱ��|���", strValues, "ID|" & lngRecord)
                        Call Record_Update(rsTimeCur, "ID|ʱ��|����ʱ��|���", strValues, "ID|" & lngRecord)
                        blnSaved = True
                    End If
                Else
                    If !�к� > mlngTime Then
                        'ȡָ����Ԫ�������
                        strCellData = NVL(!����)
                        strPart = NVL(!��λ)
                        strReturn = ShowInput(!�к�, strCellData, True)
                        'strOrders��ʽ����Ŀ���,��Ŀ���...
                        'strValues��ʽ��ֵ'ֵ'ֵ...
                        arrOrder = Split(Split(strReturn, "||")(0), ",")
                        arrValue = Split(Split(strReturn, "||")(1) & "'", "'")
                        arrPart = Split(strPart & "/////", "/")
                        
                        intMax = UBound(arrOrder)
                        For intPos = 0 To intMax
        '                    �ļ�ID_IN IN ���˻�������.�ļ�ID%TYPE,
        '                    ����ʱ��_IN IN ���˻�������.����ʱ��%TYPE,
        '                    ��¼����_IN IN ���˻�����ϸ.��¼����%TYPE,          --������Ŀ=1���ϱ�˵��=2�������ձ��=4��ǩ����¼=5���±�˵��=6�����������=9
        '                    ��Ŀ���_IN IN ���˻�����ϸ.��Ŀ���%TYPE,          --������Ŀ����ţ��ǻ�����Ŀ�̶�Ϊ0
        '                    ��¼����_IN IN ���˻�����ϸ.��¼����%TYPE := NULL,  --��¼���ݣ��������Ϊ�գ��������ǰ�����ݣ�37��38/37
        '                    ���²�λ_IN IN ���˻�����ϸ.���²�λ%TYPE := NULL,
        '                    ���˼�¼_IN IN NUMBER := 1,
                            '65258:������,2013-11-1,С��Ϊ��ҲҪ��ʾ(С�������Ŀǿ�Ʋ���Chr(13))
                            If NVL(!����, 0) = 1 And arrValue(intPos) = "" And InStr(1, "|" & mstrColCollect, "|" & Val(NVL(!�к�, 0)) - cHideCols & ";") > 0 Then
                                arrValue(intPos) = Chr(13)
                            End If
                            '������ܸ��ݻ�����Ŀ���кź���Ż�ȡ��Ӧ�Ĺ�����Ŀ���
                            strRelationNO = ""
                            If NVL(!����, 0) = 1 And Val(NVL(mrsCellMap!��¼���)) > 0 Then
                                strRelationNO = GetRelatiionNo(Val(NVL(!�к�, 0)) - cHideCols & "," & arrOrder(intPos))
                            End If
                            gstrSQL = "ZL_���˻�������_UPDATE(" & mlng�ļ�ID & ",to_date('" & strDatetime & "','yyyy-MM-dd hh24:mi:ss'),1," & _
                                    arrOrder(intPos) & ",'" & arrValue(intPos) & "','" & arrPart(intPos) & "'," & intAllow & ",0," & _
                                    IIf(mblnVerify, 1, 0) & ",NULL," & IIf(IsNull(mrsCellMap!��¼���), IIf(NVL(!����, 0) = 1, 0, "NULL"), Val(NVL(mrsCellMap!��¼���)))
                            If strRelationNO = "" Then
                                gstrSQL = gstrSQL & ",NULL,'" & NVL(!���) & "')"
                            Else
                                gstrSQL = gstrSQL & "," & Val(strRelationNO) & ",'" & NVL(!���) & "')"
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
    
    '�������ݷ���ʱ�䣬���ڷ��������м�ĳ�����������仯��������汾�������ݵ�������������ʱ�䷢���仯����(����������)��
    ',ID:403,ʱ��:2012/5/8 18:23:00,����ʱ��:2012/5/8 18:23:00
    ',ID:407,ʱ��:2012/5/8 18:23:02,����ʱ��:2012/5/8 18:23:01
    ',ID:517,ʱ��:2012/5/8 18:23:03,����ʱ��:2012/5/8 18:23:02
    '��Ҫ�ȸ������һ�з���ʱ�䣺��(����������):
    ',ID:403,ʱ��:2012/5/8 18:23:00,����ʱ��:2012/5/8 18:23:00
    ',ID:407,ʱ��:2012/5/8 18:23:01,����ʱ��:2012/5/8 18:23:02
    ',ID:517,ʱ��:2012/5/8 18:23:02,����ʱ��:2012/5/8 18:23:03
    '��Ҫ��ǰ�������
    strDays = ""
    rsTime.Filter = ""
    'Call OutputRsData(rsTime)
    rsTime.Sort = "ʱ�� DESC"
    Do While Not rsTime.EOF
        If InStr(1, "," & strDays & ",", "," & rsTime!ID & ",") = 0 Then
            rsTimeCur.Filter = "����ʱ��='" & Format(rsTime!ʱ��, "YYYY-MM-DD HH:mm:ss") & "'And ���=0 And ID<>" & Val(rsTime!ID)
            If rsTimeCur.RecordCount > 0 Then
                lngRecord = rsTimeCur!ID
                gstrSQL = UpdateTime(rsTimeCur, Format(rsTimeCur!ʱ��, "YYYY-MM-DD HH:mm:ss"), lngRecord)
                strSQLTime(ReDimArray(strSQLTime)) = gstrSQL
                rsTimeCur.Filter = ""
                Call Record_Update(rsTimeCur, "���", "1", "ID|" & lngRecord)
                strDays = IIf(strDays = "", "", ",") & lngRecord
                GoTo ErrLoop
            Else
                lngRecord = rsTime!ID
                gstrSQL = "Zl_���˻�������_����ʱ��(" & rsTime!ID & ",to_date('" & rsTime!ʱ�� & "','yyyy-MM-dd hh24:mi:ss'))"
                strSQLTime(ReDimArray(strSQLTime)) = gstrSQL
                rsTimeCur.Filter = ""
                Call Record_Update(rsTimeCur, "���", "1", "ID|" & lngRecord)
                strDays = IIf(strDays = "", "", ",") & lngRecord
            End If
        End If
    rsTime.MoveNext
ErrLoop:
    Loop
    
    'ѭ��ִ��SQL��������
    'On Error Resume Next
    
    gcnOracle.BeginTrans
    blnTrans = True
    
    On Error GoTo ErrHand
    '�ȸ��·���ʱ��
    intMax = UBound(strSQLTime)
    If intMax > 0 Then
        For intPos = 1 To intMax
            If strSQLTime(intPos) <> "" Then
                'Debug.Print strSQLTime(intPos)
    '            objStream.WriteLine (Now & "��SQL��" & strSQLTime(intPos))
                Call zlDatabase.ExecuteProcedure(strSQLTime(intPos), "���滤���¼������")
            End If
        Next intPos
    End If
    '�ڸ�������
    intMax = UBound(strSQL)
    If intMax > 0 Then
'        objStream.WriteLine (Now & "׼����������")
        For intPos = 1 To intMax
            If strSQL(intPos) <> "" Then
                'Debug.Print strSQL(intPos)
    '            objStream.WriteLine (Now & "��SQL��" & strSQL(intPos))
                Call zlDatabase.ExecuteProcedure(strSQL(intPos), "���滤���¼������")
            End If
        Next
    '    objStream.WriteLine (Now & "�����������")
    End If
    '������С������
    intMax = UBound(strCollectSQL)
    If intMax > 0 Then
        For intPos = 1 To intMax
            If strCollectSQL(intPos) <> "" Then
                'Debug.Print strCollectSQL(intPos)
    '            objStream.WriteLine (Now & "��SQL��" & strCollectSQL(intPos))
                Call zlDatabase.ExecuteProcedure(strCollectSQL(intPos), "���滤���¼������")
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
    rsTimeCur.Filter = "����ʱ��='" & Format(strTime, "YYYY-MM-DD HH:mm:ss") & "' And ���=0 And ID<>" & lngID
    If rsTimeCur.RecordCount > 0 Then
        lngID = Val(rsTimeCur!ID)
        strSQL = UpdateTime(rsTimeCur, Format(rsTimeCur!ʱ��, "YYYY-MM-DD HH:mm:ss"), lngID)
    Else
        strSQL = "Zl_���˻�������_����ʱ��(" & lngID & ",to_date('" & strTime & "','yyyy-MM-dd hh24:mi:ss'))"
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

Private Sub cboС��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txt��ʼʱ��.Enabled Then
            txt��ʼʱ��.SetFocus
        Else
            txtС������.SetFocus
        End If
    End If
End Sub


Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strDate As String, strTime As String, strYear As String
    Dim strLockItem As String                   'ͬ������������,�������޸Ļ�ɾ��
    Dim lngTop As Long, lngHeight As Long
    Dim intMax As Integer                       'ͬ������������ռ�õ��������
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
    'ճ��,���ʱ��Ҫͬ��mrsCellMap����
    Case conMenu_Edit_Group_New
        '��ӷ��飨��ǰ��¼�����Ϊ1������������ͨ����û������ֻ���������б仯���������ݱ�������¼�룬��֧���޸�Ϊ����򽫷��������޸�Ϊ��ͨ���ݵĹ��ܣ�
        Control.Category = ""
        mblnGroupNew = mblnGroupNew Xor True
        If mblnGroupNew Then
            '��¼��ʼ�У�����ʼ�в�����¼��������ʱ��
            Control.Category = VsfData.ROW
        End If
        Control.Checked = mblnGroupNew
        mstrGroupRow = Control.Category
    Case conMenu_Edit_Group_Append
        If VsfData.ROW < VsfData.FixedRows Then Exit Sub
        If InStr(1, FormatValue(VsfData.TextMatrix(VsfData.ROW, mlngRowCount)), "|") = 0 Then Exit Sub
        If Val(VsfData.TextMatrix(lngRow, mlngCollectType)) <> 0 Then Exit Sub
        
        '��������ʾ��¼��ؼ�
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
        
        '׷������ʱ�������¼���ѡ�����ݵ������������ǲ��ܰ������ı�����Ϣ��
        '�磺һ������5�У����Ǵ��ı�������ֻ��3�У�ѡ�и���׷������ʱ��Ӧ��׷�ӵ���4�У�demo=1��Ϊ3�У�demo=4��Ϊ2��
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
        '������д�����
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
                '�������������
                For lngCol = 0 To VsfData.Cols - 1
                    If VsfData.ColHidden(lngCol) = True Then VsfData.TextMatrix(lngRow, lngCol) = ""
                Next lngCol
                VsfData.TextMatrix(lngRow, mlngRowCount) = (lngStartRow + lngRowCount - intNULL - 1) & "|" & (lngRow - intNULL)
                VsfData.TextMatrix(lngRow, mlngRowCurrent) = (lngStartRow + lngRowCount - intNULL - 1)
            Next
            lngRowCount = Val(Split(FormatValue(VsfData.TextMatrix(lngStartRow, mlngRowCount)), "|")(0))
            '���±��д��ı������ݼ���Ϣ
            Call CellMap_UpdateAssistant(lngStartRow)
            blnTure = False
        Else
            '�����һ�������Ƿ�Ϊ��,������ǿ���ֱ����ӵ���һ��
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
            '׷�ӷ��飬�ڵ�ǰ��(ֻ��һ�е�������)�����ӿհ���
            '1)������1������
            VsfData.Rows = VsfData.Rows + 1
            '   �ӵ�ǰ�м�¼�Ŀհ��п�ʼ��ÿ�е�λ��+�����ӵĿհ�����
            For lngRow = VsfData.Rows - 2 To lngStartRow + lngRowCount Step -1
                VsfData.RowPosition(lngRow) = lngRow + 1
            Next
            lngRow = lngStartRow + lngRowCount - 1
            '2)���кŷ����仯����ͬ������mrsCellMap�д��ڸ��кŵ��к�����
            Call CellMap_Update(lngRow, 1)
        End If
        '3)���·�����ؿ���
        mintType = -1: mblnShow = False
        Call AppendGroup(lngStartRow)
        lngRow1 = VsfData.ROW
        lngCol1 = VsfData.COL
        If InStr(1, FormatValue(VsfData.TextMatrix(lngRow1, mlngRowCount)), "|") = 0 Then VsfData.TextMatrix(lngRow1, mlngRowCount) = "1|1"
        intGroupFirstRows = Val(Split(FormatValue(VsfData.TextMatrix(lngRow1, mlngRowCount)), "|")(0))
        '��һ���������������ȡ������ѡ��
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
        '��ԭ�з��������Ϸ���,��Ҫ����������
        If Val(FormatValue(VsfData.TextMatrix(lngStartRow, mlngDemo))) > 0 Then '����������
            'ȷ��������ʼ��
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
            '������֯�������
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
                'Ѱ�Ҵ��ı���
                mrsSelItems.Filter = "��=" & lngCol - cHideCols
                If mrsSelItems.RecordCount > 0 Then
                    lngOrder = Val(mrsSelItems!��Ŀ���)
                    mrsItems.Filter = "��Ŀ���=" & lngOrder
                    If mrsItems.RecordCount = 0 Then
                        mrsItems.Filter = 0
                        GoTo ErrNext
                    End If
                    mblnEditAssistant = (mrsItems!��Ŀ���� = 1 And mrsItems!��Ŀ���� > 100) And Val(FormatValue(VsfData.TextMatrix(lngStartRow, mlngDemo))) <= 1
                    If Not mblnEditAssistant Then GoTo ErrNext
                        
                    If InStr(1, FormatValue(VsfData.TextMatrix(lngStartRow, mlngRowCount)), "|") = 0 Then VsfData.TextMatrix(lngStartRow, mlngRowCount) = "1|1"
                    intGroupFirstRows = Val(Split(FormatValue(VsfData.TextMatrix(lngStartRow, mlngRowCount)), "|")(0))
                    'Ϊ������ʱ��ѡ��������ʼ�У��༭������ʾ���д��ı���
                    strText = ""
                    If Val(FormatValue(VsfData.TextMatrix(lngStartRow, mlngDemo))) = 1 Then
                        For lngRow = 0 To intGroupFirstRows - 1
                            strText = strText & IIf(strText = "", "", vbCrLf) & Replace(Replace(Replace(VsfData.TextMatrix(lngRow + lngStartRow, lngCol), Chr(13), ""), Chr(10), ""), Chr(1), "")
                        Next lngRow
                        lngCount = lngStartRow + intGroupFirstRows - 1
                        For lngRow = lngStartRow + intGroupFirstRows To VsfData.Rows - 1
                            If VsfData.RowHidden(lngRow) = False Then
                                 '������ÿһ�������������ݿ���ռ�ö��У�ֻ���ڵ�һ�б����˷�����������������������=ÿһ������������+ÿһ���������е�����
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
                        'lngRow1 Ҫ׷�ӵ���
                        Call MoveNextCell(False, True, strText, lngRow1)
                        mintType = -1
                        blnTure = True
                    End If
ErrNext:
                End If
            Next lngCol
            'blnTrue=false ˵����¼��û�д����ı�(������������ʼ�е��׷�ӣ���׷����¼�����ݣ�ֻ�ܱ���׷���к��ڽ���һ�����е�һ������)
            If blnTure = False Then
                '����ʼ�п�ʼ�����������(��ֹ�Ѿ�����ķ������ݷ����кͱ����ʱ�䲻��Ӧ�����´���������ͬʱ�������)
                '�磺�����������ʼ��Demo=1������ʱ������Ϊ=01����ʱ׷��һ���¼�¼��Demo=2 ����ʱ����ҲΪ01(����޸�����ʼ�����ݾͲ������������)
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
                            '�������޸ĲŽ�������̣�ȡ������¼��ʵ��ʱ��
                            If mblnDateAd Then
                                strYear = Format(FormatValue(VsfData.TextMatrix(lngRow, mlngActTime)), "YYYY")
                                strDate = Format(FormatValue(VsfData.TextMatrix(lngRow, mlngActTime)), "DD") & "/" & Format(FormatValue(VsfData.TextMatrix(lngRow, mlngActTime)), "MM")
                            Else
                                strDate = Mid(FormatValue(VsfData.TextMatrix(lngRow, mlngActTime)), 1, 10)
                            End If
                            strTime = Mid(FormatValue(VsfData.TextMatrix(lngRow, mlngActTime)), 12, 5)
                        Else
                            '����ʱ���������
                            strDate = VsfData.TextMatrix(lngRow - Val(FormatValue(VsfData.TextMatrix(lngRow, mlngDemo))) + 1, mlngDate)
                            strTime = VsfData.TextMatrix(lngRow - Val(FormatValue(VsfData.TextMatrix(lngRow, mlngDemo))) + 1, mlngTime)
                            If mblnDateAd Then strYear = VsfData.TextMatrix(lngRow - Val(FormatValue(VsfData.TextMatrix(lngRow, mlngDemo))) + 1, mlngYear)
                        End If
                        
                        '1\����
                        strField = "ID|ҳ��|�к�|�к�|��¼ID|����|��λ|ɾ��"
                        If mlngDate <> -1 Then
                            strKey = mintҳ�� & "," & lngRow & "," & mlngDate
                            strValue = strKey & "|" & mintҳ�� & "|" & lngRow & "|" & mlngDate & "|" & _
                                Val(VsfData.TextMatrix(lngRow, mlngRecord)) & "|" & strDate & "|" & strYear & "|0"
                            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                        End If
                        '2\ʱ��
                        strKey = mintҳ�� & "," & lngRow & "," & mlngTime
                        strValue = strKey & "|" & mintҳ�� & "|" & lngRow & "|" & mlngTime & "|" & _
                            Val(VsfData.TextMatrix(lngRow, mlngRecord)) & "|" & strTime & "|" & _
                            VsfData.TextMatrix(lngRow, mlngDemo) & "|0"
                        Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                    End If
                Next lngRow
            Else
                '��ԭѡ����
                mblnShow = False
                mintType = -1
                VsfData.ROW = lngRow1
                VsfData.COL = lngCol1
            End If
        End If
        If InStr(1, VsfData.TextMatrix(lngRow1, mlngRowCount), "|") = 0 Then VsfData.TextMatrix(lngRow1, mlngRowCount) = "1|1"
    Case conMenu_Edit_Copy
        '����ָ�������е�����
        If VsfData.ROW < VsfData.FixedRows Then Exit Sub
        If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then Exit Sub
        lngRow = GetStartRow(VsfData.ROW)
        If Val(VsfData.TextMatrix(lngRow, mlngCollectType)) <> 0 Then Exit Sub
        
        '���Ƽ�¼��
        Set mrsCopyMap = New ADODB.Recordset
        Set mrsCopyMap = CopyNewRec(mrsDataMap, False)
        
        '�õ�ָ�������е���ʼ��,������
        lngCols = VsfData.Cols - 1
        lngRows = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0))
        lngRows = lngRow + lngRows - 1
        For lngRow = lngRow To lngRows
            mrsCopyMap.AddNew
            mrsCopyMap!ҳ�� = mintҳ��
            mrsCopyMap!�к� = lngRow
            For lngCol = 0 To lngCols - VsfData.FixedCols    '����һ���̶���
                mrsCopyMap.Fields(cControlFields + lngCol).Value = IIf(FormatValue(VsfData.TextMatrix(lngRow, lngCol + VsfData.FixedCols)) = "", Null, FormatValue(VsfData.TextMatrix(lngRow, lngCol + VsfData.FixedCols)))
            Next
            mrsCopyMap.Update
        Next
    Case conMenu_Edit_PASTE
        'ճ��ʱ����Ŀ�������帲�ǣ�ͬ�������������У���г���
        '���Ŀ���ܲ�ͬҳ����Ŀ��ͬ����λ��ͬ�����Բ����ǻ��Ŀ
        'ͬ������ռ�õ��������䣬�粻������ӿհ��У�����ճ��
        If VsfData.ROW < VsfData.FixedRows Then Exit Sub
        If mrsCopyMap.RecordCount = 0 Then Exit Sub
        
        '��ҳ�����в���������н���ճ��,ɾ��,ֻ�ܱ༭�����Ŀ�����
        If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then
            lngStartRow = VsfData.ROW
        Else
            lngStartRow = GetStartRow(VsfData.ROW)
        End If
        If Val(VsfData.TextMatrix(lngStartRow, mlngCollectType)) <> 0 Then Exit Sub
        If Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) >= 1 Then Exit Sub
        
        If InStr(1, VsfData.TextMatrix(lngStartRow, mlngRowCount), "|") <> 0 And lngStartRow = 3 And Val(VsfData.TextMatrix(lngStartRow, mlngStartRowPage)) <> mintҳ�� Then
            If Val(Split(VsfData.TextMatrix(lngStartRow, mlngRowCount), "|")(0)) <> Val(VsfData.TextMatrix(lngStartRow, mlngRowCurrent)) Then
                RaiseEvent AfterRowColChange("��ҳ�����в�����ճ�������л�����һҳ���в�����", True, mblnSign, mblnArchive)
                Exit Sub
            End If
        End If
        
        If VsfData.TextMatrix(lngStartRow, mlngSigner) <> "" Then
            RaiseEvent AfterRowColChange("��ǩ�������ݲ�����ճ������ȡ��ǩ�������ԣ�", True, mblnSign, mblnArchive)
            Exit Sub
        End If
        
        '�����Ѿ��������ݣ�������������Ѿ�ǩ������ճ��
        blnTure = False
        If Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) > 0 And mstrColCollect <> "" Then
            '�ҳ����ݲ�Ϊ�յĻ�����
            For lngRow = 0 To UBound(Split(mstrColCollect, "|"))
                strValue = GetRelatiionNo(CStr(Split(mstrColCollect, "|")(lngRow)), 2)
                strCols = strCols & "," & IIf(strValue = "", "", strValue & ",") & Split(Split(mstrColCollect, "|")(lngRow), ";")(0)
            Next
            strCols = Mid(strCols, 2)
            If strCols <> "" Then
                If ISCollectSigned(mlng�ļ�ID, Format(VsfData.TextMatrix(lngStartRow, mlngActTime), "YYYY-MM-DD"), Format(VsfData.TextMatrix(lngStartRow, mlngActTime), "HH:MM")) Then
                    blnTure = True
                    If MsgBox("��Ҫ�޸ĵ���������Ӧ�Ļ���������ǩ���������������������Ļ��������ݽ����ܱ�ճ�����������Ƿ������", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Sub
                End If
            End If
        End If
        
        '���Ŀ���������Ƿ����ͬ������������,�����������ͬ���ļ�¼
        strLockItem = GetSynItems(2, intMax)        '1.������Ŀ���;2.�����к�
        
        '�õ�Ŀ�������е���ʼ��,������
        strField = "ID|ҳ��|�к�|�к�|��¼ID|����|��λ|ɾ��"
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
            'ɾ�������������,����һ��
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
            '����������
            Call CellMap_Update(lngStartRow, -1 * lngRows)
        End If
        
        '������������,�������������������������ӵ�����
        intNULL = mrsCopyMap.RecordCount - 1
        For lngRow = 1 To mrsCopyMap.RecordCount - 1
            '��֤��ǰ�����������һҳ����ʾȫ
            If lngRow + lngStartRow > VsfData.Rows - 1 Then Exit For
            
            If Val(VsfData.TextMatrix(lngRow + lngStartRow, mlngRecord)) = 0 And VsfData.TextMatrix(lngRow + lngStartRow, mlngRowCount) = "" Then
                intNULL = intNULL - 1
            Else
                Exit For
            End If
        Next
        '�����ӿ���
        If intNULL > 0 Then
            VsfData.Rows = VsfData.Rows + intNULL
            '�ӵ�ǰ�м�¼�Ŀհ��п�ʼ��ÿ�е�λ��+�����ӵĿհ�����
            For lngRow = 1 To intNULL
                VsfData.RowPosition(VsfData.Rows - 1) = lngStartRow + 1
            Next
        End If
        
        '��ԭ���ڣ�ʱ�䣬ǿ�Ʋ������޸�
        VsfData.TextMatrix(lngStartRow, mlngDate) = strDate
        VsfData.TextMatrix(lngStartRow, mlngTime) = strTime
        VsfData.TextMatrix(lngStartRow, mlngYear) = strYear
        '��¼�û��޸Ĺ��ĵ�Ԫ��
        If mlngDate <> -1 Then
            strKey = mintҳ�� & "," & lngStartRow & "," & mlngDate
            strValue = strKey & "|" & mintҳ�� & "|" & lngStartRow & "|" & mlngDate & "|" & _
                Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "|" & VsfData.TextMatrix(lngStartRow, mlngDate) & "|" & strYear & "|0"
            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
        End If
        '2\ʱ��
        strKey = mintҳ�� & "," & lngStartRow & "," & mlngTime
        strValue = strKey & "|" & mintҳ�� & "|" & lngStartRow & "|" & mlngTime & "|" & _
            Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "|" & VsfData.TextMatrix(lngStartRow, mlngTime) & "||0"
        Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
        '�����������
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
                                
                                '�޸ı�־
                                If .AbsolutePosition = .RecordCount And lngCol < mlngNoEditor Then
                                    strKey = mintҳ�� & "," & lngStartRow & "," & lngCol + VsfData.FixedCols
                                    strValue = strKey & "|" & mintҳ�� & "|" & lngStartRow & "|" & lngCol + VsfData.FixedCols & "|" & _
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
        '���кŷ����仯����ͬ������mrsCellMap�д��ڸ��кŵ��к�����
        Call CellMap_Update(lngStartRow, mrsCopyMap.RecordCount - 1)

        '�����ɫ
        'Call WriteColor
        mblnChange = True
        RaiseEvent AfterDataChanged(mblnChange Or mblnVerify)
    
    Case conMenu_Edit_Clear
        If VsfData.ROW < VsfData.FixedRows Then Exit Sub
        If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then Exit Sub
        If VsfData.TextMatrix(VsfData.ROW, mlngSigner) <> "" Then
            RaiseEvent AfterRowColChange("��ǩ�������ݲ�����ɾ����", True, mblnSign, mblnArchive)
            Exit Sub
        End If
        
        '��ҳ�����в���������н���ճ��,ɾ��,ֻ�ܱ༭�����Ŀ�����
        If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then
            lngStartRow = VsfData.ROW
        Else
            lngStartRow = GetStartRow(VsfData.ROW)
        End If
        If InStr(1, VsfData.TextMatrix(lngStartRow, mlngRowCount), "|") <> 0 And lngStartRow = 3 And Val(VsfData.TextMatrix(lngStartRow, mlngStartRowPage)) <> mintҳ�� Then
            If Val(Split(VsfData.TextMatrix(lngStartRow, mlngRowCount), "|")(0)) <> Val(VsfData.TextMatrix(lngStartRow, mlngRowCurrent)) Then
                RaiseEvent AfterRowColChange("��ҳ�����в�����ɾ�������л�����һҳ���в�����", True, mblnSign, mblnArchive)
                Exit Sub
            End If
        End If
        
        lngRowCount = Split(VsfData.TextMatrix(lngStartRow, mlngRowCount), "|")(0)
        '���Ŀ���������Ƿ����ͬ������������,�����������ͬ���ļ�¼
        strLockItem = GetSynItems(2, intMax)        '1.������Ŀ���;2.�����к�
        
        '׼��ɾ��
        strField = "ID|ҳ��|�к�|�к�|��¼ID|����|����|ɾ��"
        If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then
            lngRow = VsfData.ROW
            lngRows = 1
        Else
            lngRow = GetStartRow(VsfData.ROW)
            lngStartRow = lngRow
            If VsfData.TextMatrix(lngStartRow, mlngSigner) <> "" Then
                RaiseEvent AfterRowColChange("��ǩ�������ݲ�����ɾ����", True, mblnSign, mblnArchive)
                Exit Sub
            End If
            lngRows = Val(Split(VsfData.TextMatrix(lngStartRow, mlngRowCount), "|")(0))
        End If
        
        blnTure = False
        '�Ѿ���������ݲ�����ɾ����ʼ��
        If Val(VsfData.TextMatrix(lngRow, mlngDemo)) = 1 And lngRow + lngRows < VsfData.Rows Then
            lngCount = lngRow + lngRows - 1
            For lngCurRow = lngRow + lngRows To VsfData.Rows - 1
                 '������ÿһ�������������ݿ���ռ�ö��У�ֻ���ڵ�һ�б����˷�����������������������=ÿһ������������+ÿһ���������е�����
                If lngCurRow > lngCount Then
                    If Val(VsfData.TextMatrix(lngCurRow, mlngDemo)) <= 1 Then Exit For
                    If VsfData.RowHidden(lngCurRow) = False Then blnTure = True: Exit For 'ֻҪ����һ��û�����صķ�����˳�
                    lngCount = Val(Split(VsfData.TextMatrix(lngCurRow, mlngRowCount), "|")(0)) + lngCurRow - 1
                End If
            Next lngCurRow
        End If
        If blnTure = True Then
            RaiseEvent AfterRowColChange("���ڷ���������ʱ��������ɾ��������ʼ�С�", True, mblnSign, mblnArchive)
            Exit Sub
        End If
        
        '���е����ݴ��ڻ�����ǩ�������ݲ�����ɾ��
        If Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) > 0 And Not Val(VsfData.TextMatrix(lngStartRow, mlngCollectType)) < 0 Then
            If CheckCollectIsData(lngStartRow, 1) = True Then
                If ISCollectSigned(mlng�ļ�ID, Format(VsfData.TextMatrix(lngStartRow, mlngActTime), "YYYY-MM-DD"), Format(VsfData.TextMatrix(lngStartRow, mlngActTime), "HH:MM")) Then
                    RaiseEvent AfterRowColChange("��Ҫɾ�������ݴ��ڻ��������ݣ��ұ�����������Ӧ�Ļ���������ǩ����������ɾ����", True, mblnSign, mblnArchive)
                    Exit Sub
                End If
            End If
        End If
        '��������ʾ��¼��ؼ�
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
        '��ȡ�������ݵĴ��ı�����������
        If Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) > 1 Then
            Call GetGroupAssistant(strAssistantCols, varAssistant)
        End If
        'ɾ������������
        For intNULL = 2 To lngRows
            VsfData.RowHidden(lngRow + intNULL - 1) = True
        Next
        
        '�������ʼ�з������ݣ���������ı���Ϣ��ȡ���÷���
        '�磺������������飬����ڶ���ʱ�����ڶ�����Ķ������ۼ��ڵ�3����
        '��¼�û��޸Ĺ��ĵ�Ԫ��
        If Val(VsfData.TextMatrix(lngStartRow, mlngCollectType)) = 0 Then
            strYear = ""
            If Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) > 0 Then
                If CheckGroupDate(lngStartRow) = True Then
                    '�������޸ĲŽ�������̣�ȡ������¼��ʵ��ʱ��
                    strDate = Mid(VsfData.TextMatrix(lngStartRow, mlngActTime), 1, 10)
                    strTime = Mid(VsfData.TextMatrix(lngStartRow, mlngActTime), 12, 5)
                    If mblnDateAd Then strYear = Mid(VsfData.TextMatrix(lngStartRow, mlngActTime), 1, 4)
                Else
                    '����ʱ���������
                    strDate = VsfData.TextMatrix(lngStartRow - Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) + 1, mlngDate)
                    strTime = VsfData.TextMatrix(lngStartRow - Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) + 1, mlngTime)
                    If mblnDateAd Then strYear = VsfData.TextMatrix(lngStartRow - Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) + 1, mlngYear)
                End If
            Else
                '��ͨ����
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
            
            strField = "ID|ҳ��|�к�|�к�|��¼ID|����|��λ|����|ɾ��"
            strKey = mintҳ�� & "," & lngStartRow & "," & mlngDate
            strValue = strKey & "|" & mintҳ�� & "|" & lngStartRow & "|" & mlngDate & "|" & _
                Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "|" & strDate & "|" & strYear & "|0|1"
            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
            '2\ʱ��
            strKey = mintҳ�� & "," & lngStartRow & "," & mlngTime
            strValue = strKey & "|" & mintҳ�� & "|" & lngStartRow & "|" & mlngTime & "|" & _
                Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "|" & strTime & "|" & VsfData.TextMatrix(lngStartRow, mlngDemo) & "|0|1"
            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
        Else
            '1\����
            strKey = mintҳ�� & "," & lngStartRow & "," & mlngDate
            strValue = strKey & "|" & mintҳ�� & "|" & lngStartRow & "|" & mlngDate & "|" & Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "|" & _
                    VsfData.TextMatrix(lngStartRow, mlngCollectText) & ";" & Val(VsfData.TextMatrix(lngStartRow, mlngCollectType)) & ";" & _
                    Val(VsfData.TextMatrix(lngStartRow, mlngCollectStyle)) & ";" & VsfData.TextMatrix(lngStartRow, mlngCollectDay) & ";" & _
                    VsfData.TextMatrix(lngStartRow, mlngCollectStart) & ";" & VsfData.TextMatrix(lngStartRow, mlngCollectEnd) & "|1|1"
            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
        End If
        strField = "ID|ҳ��|�к�|�к�|��¼ID|����|��λ|����|ɾ��"
        
        'ɾ����ʼ���з�ͬ��������
        If strLockItem = "" Then
            VsfData.RowHidden(lngRow) = True
            If Val(VsfData.TextMatrix(lngStartRow, mlngCollectType)) = 0 Then
                '��д�޸ı�־
                For lngCol = mlngTime + 1 To mlngNoEditor - 1
                    If InStr(1, "," & mstrCatercorner & ",", "," & lngCol - (cHideCols + VsfData.FixedCols - 1) & ",") = 0 Then
                        strPart = GetActivePart(lngCol, 0)
                    Else
                        strPart = GetActivePart(lngCol, 0) & "/" & GetActivePart(lngCol, 1)
                    End If
                    strKey = mintҳ�� & "," & lngStartRow & "," & lngCol
                    strValue = strKey & "|" & mintҳ�� & "|" & lngStartRow & "|" & lngCol & "|" & _
                        Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "||" & strPart & "|0|1"
                    Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                Next
                If Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) >= 1 Then VsfData.TextMatrix(lngStartRow, mlngRowCount) = "1|1"
            End If
        Else
            '��д�޸ı�־(����ͬ������,������ʱ���в��������)``
            For lngCol = mlngTime + 1 To mlngNoEditor - 1
                If InStr(1, "," & strLockItem & ",", "," & lngCol - (cHideCols + VsfData.FixedCols - 1) & ",") = 0 And lngCol <> mlngDate And lngCol <> mlngTime Then
                    VsfData.TextMatrix(lngStartRow, lngCol) = ""
                    If InStr(1, "," & mstrCatercorner & ",", "," & lngCol - (cHideCols + VsfData.FixedCols - 1) & ",") = 0 Then
                        strPart = GetActivePart(lngCol, 0)
                    Else
                        strPart = GetActivePart(lngCol, 0) & "/" & GetActivePart(lngCol, 1)
                    End If
                    strKey = mintҳ�� & "," & lngStartRow & "," & lngCol
                    strValue = strKey & "|" & mintҳ�� & "|" & lngStartRow & "|" & lngCol & "|" & _
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
            
            '������֯�����кźʹ��ı�������
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
        
        '��鵱ǰ¼��ؼ�
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
        If frmTendFileElement.ShowMe(mfrmParent, mlng�ļ�ID, mlng��ʽID, mintҳ��, mrsElement, IIf(mblnBlowup = True, 1, 0)) = True Then
            '������ȡ����
            mblnInit = False
            Call InitVariable
            Call InitCons
            Call ReadStruDef
            Call zlRefresh
            mblnInit = True
            VsfData.Refresh
            mcbrPage.Caption = "ҳ��ѡ�񣺵�" & mintҳ�� & "ҳ"
            cbsThis.RecalcLayout
        End If
    Case conMenu_Edit_Append
        Call BoundItems(VsfData.COL - (cHideCols + VsfData.FixedCols - 1))
    Case conMenu_Edit_PrevPage
        If mintҳ�� > mint��ʼҳ�� Then
            If Not DataMap_Save Then Exit Sub
            mintҳ�� = mintҳ�� - 1
            '���²�ѯSQL
            '������ȡ����
            mblnInit = False
            Call InitVariable
            Call InitCons
            Call ReadStruDef
            Call zlRefresh
            mblnInit = True
            VsfData.Refresh
        
            mcbrPage.Caption = "ҳ��ѡ�񣺵�" & mintҳ�� & "ҳ"
            cbsThis.RecalcLayout
        End If
    Case conMenu_Edit_NextPage
        If mintҳ�� < mint����ҳ + 1 Then
            If Not DataMap_Save Then Exit Sub
            mintҳ�� = mintҳ�� + 1
            '���²�ѯSQL
            '������ȡ����
            mblnInit = False
            Call InitVariable
            Call InitCons
            Call ReadStruDef
            Call zlRefresh
            mblnInit = True
            VsfData.Refresh
            
            mcbrPage.Caption = "ҳ��ѡ�񣺵�" & mintҳ�� & "ҳ"
            cbsThis.RecalcLayout
        End If
    Case conMenu_View_Jump
        If Not DataMap_Save Then Exit Sub
        
        '���²�ѯSQL
        '������ȡ����
        mintҳ�� = Control.Parameter
        mblnInit = False
        Call InitVariable
        Call InitCons
        Call ReadStruDef
        Call zlRefresh
        
        mblnInit = True
        VsfData.Refresh
        
        mcbrPage.Caption = "ҳ��ѡ�񣺵�" & mintҳ�� & "ҳ"
        cbsThis.RecalcLayout
    Case conMenu_Edit_Word
        Call cmdWord_Click
    Case conMenu_Edit_Brief
        Call ShowBrief
    Case conMenu_Edit_Import
        '��������
        If VsfData.ROW < VsfData.FixedRows Then Exit Sub '���������Ѿ���update�н������ж�
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
                Control.Visible = InStr(1, mstrPrivs, "�����¼�Ǽ�") <> 0
        End Select
    End If
    
    Select Case Control.ID
    Case conMenu_Edit_Group_New  '���飬ֻ������ӵ�������Ч
        Control.Checked = mblnGroupNew
        Control.Enabled = mblnEditable And Not mblnArchive And Not mblnVerify And Not mblnGroupApp _
            And Val(VsfData.TextMatrix(VsfData.ROW, mlngRecord)) = 0 And Val(mstrGroupRow) <= VsfData.ROW
        '63934:������,2013-07-25,С���в���ʹ��׷�ӹ���
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
                '63934:������,2013-07-25,С���в���ʹ��׷�ӹ���
                '1���Ѿ�ǩ�������ݲ���׷��.2��׷��ֻ���������ݵ���׷��(���������ı���Ŀ).3��С���в���׷��
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
        'ǩ�����ݲ�����ճ��
        If VsfData.TextMatrix(VsfData.ROW, mlngRowCount) <> "" Then
            intDo = GetStartRow(VsfData.ROW)
        Else
            intDo = VsfData.ROW
        End If
        If Val(VsfData.TextMatrix(intDo, mlngDemo)) >= 1 Then Exit Sub
        If VsfData.TextMatrix(intDo, mlngSigner) <> "" Then Exit Sub
        If Val(VsfData.TextMatrix(intDo, mlngCollectType)) <> 0 Then Exit Sub
        'ճ�������ڸ�����ճ��
        mrsCopyMap.Filter = "ҳ��=" & mintҳ�� & " And �к�=" & intDo
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
        Control.Enabled = (mintҳ�� > mint��ʼҳ��)
    Case conMenu_Edit_NextPage
        Control.Enabled = (mintҳ�� < mint����ҳ + 1)
    Case conMenu_Edit_Word
        '60291:������,2013-04-17,ֻҪ���ı���Ŀ��������дʾ�ѡ��
        Control.Enabled = Control.Visible And (mblnEditAssistant Or mblnEditText) And mblnShow And Not mblnArchive And mblnEditable
    Case conMenu_Edit_Brief
        Control.Enabled = Control.Visible And Not mblnArchive And Not mblnVerify And mblnEditable
    Case conMenu_Edit_Import '��������
        Control.Enabled = Control.Visible And Not mblnArchive And Not mblnVerify And mblnEditable And mblnShow And mstrColCollect <> ""
         If Control.Enabled Then
            '�ж�ѡ������Ƿ��ǻ�����Ŀ��(һ�а�����������Ҳ����ʹ�ô˹���)
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
        Control.Checked = (Val(Control.Parameter) = mintҳ��)
    Case conMenu_Tool_SignEarse, conMenu_Tool_SignAuditCancel '����ȡ��ǩ������ǩ
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
'׷�ӷ������ݣ���ѡ����������ݲ���׷�ӣ����������ı���Ŀ��
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
    Dim blnSel As Boolean            '�Ƿ�ȫ��ѡ��
    Dim blnUpdate As Boolean
    Dim intLevel As Integer
    Dim lngRow As Long, lngRows As Long
    Dim strKey As String, strField As String, strValue As String
    Dim lngStart As Long, lngDemo As Long, lngCurRow As Long, lngRowCount As Long, lngNextGroupRow As Long
    Dim arrRow, blnTrue As Boolean
    '��������ȫ��ѡ�л�ȡ��ѡ�У����������
    
    If Not mblnInit Then Exit Sub
    lngRows = VsfData.Rows - 1
    strField = "ID|ҳ��|�к�|�к�|��¼ID|����|ɾ��"
    
    blnSel = chkSwitch.Value
    For lngRow = VsfData.FixedRows To lngRows
        If Not VsfData.RowHidden(lngRow) Then
            If FormatValue(VsfData.TextMatrix(lngRow, mlngRowCount)) Like "*|1" Then
                '��������ҲҪǩ��,���ע��
                'If Val(VsfData.TextMatrix(lngRow, mlngCollectType)) = 0 Then    '�����в�����༭
                    blnUpdate = False
                    blnTrue = blnSel
                    If blnSel Then
                        '���,ǩ�����ļ�¼,�ҵ�ǰ����Ա������ϴ�ǩ�������
                        If VsfData.TextMatrix(lngRow, mlngSignLevel) = "" Then
                            intLevel = δ����
                        Else
                            intLevel = Val(VsfData.TextMatrix(lngRow, mlngSignLevel)) + 1
                        End If
                        '43588:������,2012-09-13,��Ӽ�¼����ǩģʽ
                        If IIf(mintSignMode = 0, mintVerify < intLevel, InStr(1, VsfData.TextMatrix(lngRow, mlngSigner), "/") = 0 And mintVerify <> δ����) And intLevel <> δ���� Then
                        'If mintVerify < intLevel And intLevel <> δ���� then
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
                        If lngDemo > 1 Then 'Ѱ����ʼ��
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
                        '1.�����������Ƿ��Ѿ�ǩ����2.��¼�������ݿ�ʼ��
                        For lngCurRow = lngStart + lngRowCount To VsfData.Rows - 1
                            '������ÿһ�������������ݿ���ռ�ö��У�ֻ���ڵ�һ�б����˷�����������������������=ÿһ������������+ÿһ���������е�����
                            If lngCurRow > lngNextGroupRow Then
                                If Val(VsfData.TextMatrix(lngCurRow, mlngDemo)) <= 1 Then Exit For      '����������·�����˳�
                                If blnSel = True Then
                                    If VsfData.TextMatrix(lngCurRow, mlngSignLevel) = "" Then
                                        intLevel = δ����
                                    Else
                                        intLevel = Val(VsfData.TextMatrix(lngCurRow, mlngSignLevel)) + 1
                                    End If
                                    '43588:������,2012-09-13,��Ӽ�¼����ǩģʽ
                                    If Not (IIf(mintSignMode = 0, mintVerify < intLevel, InStr(1, VsfData.TextMatrix(lngCurRow, mlngSigner), "/") = 0 And mintVerify <> δ����) And intLevel <> δ����) And blnTrue = True Then
                                    'If Not (mintVerify < intLevel And intLevel <> δ����) And blnTrue = True Then
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
                        'ѡ�����з�������
                        For lngCurRow = 0 To UBound(arrRow)
                            lngStart = Val(arrRow(lngCurRow))
                            If (VsfData.Cell(flexcpChecked, lngStart, mlngChoose) <> IIf(blnTrue = False, flexTSUnchecked, flexTSChecked)) Then
                                VsfData.Cell(flexcpChecked, lngStart, mlngChoose) = IIf(blnTrue = False, flexTSUnchecked, flexTSChecked)
                                '�����޸ļ�¼�Ա�ͬ��
                                strKey = mintҳ�� & "," & lngStart & "," & mlngChoose
                                strValue = strKey & "|" & mintҳ�� & "|" & lngStart & "|" & mlngChoose & "|" & Val(VsfData.TextMatrix(lngStart, mlngRecord)) & "|" & VsfData.Cell(flexcpChecked, lngStart, mlngChoose) & "|1"
                                Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                            End If
                        Next lngCurRow
                    End If
                    
                    '�ƶ���
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
    
    '63401:������,2013-07-16,���ѡ������Ƿ��ǻ��Ŀ��
    If InStr(1, "," & mstrCOLNothing & ",", "," & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ",") = 0 Then
        RaiseEvent AfterRowColChange("��ǰѡ�е��в��ܰ󶨻��Ŀ��������ѡ���н��а󶨣�", True, mblnSign, mblnArchive)
        picCloumn.Visible = False
        Exit Sub
    End If
    
    If lstColumnUsed.ListItems.Count > 0 Then
        If Trim(txtColumnNo.Text) = "" Then
            RaiseEvent AfterRowColChange("��ͷ���Ʋ���Ϊ�գ�", True, mblnSign, mblnArchive)
            txtColumnNo.SetFocus
            Exit Sub
        End If
        If LenB(StrConv(txtColumnNo.Text, vbFromUnicode)) > 100 Then
            RaiseEvent AfterRowColChange("��ͷ���Ʋ��ܳ���50�����ֻ�100���ַ���", True, mblnSign, mblnArchive)
            txtColumnNo.SetFocus
            Exit Sub
        End If
    End If
    
    'ƴ������ʽ����ͷ����|��Ŀ���,��λ;��Ŀ���,��λ
    strPara = Trim(txtColumnNo.Text) & "|"
    intCount = lstColumnUsed.ListItems.Count
    If intCount > 2 Then
        RaiseEvent AfterRowColChange("ÿ�а󶨵���Ŀ�����ܳ���2����", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    '51883,������,2012-08-02,���Ŀ֧�ֵ�ѡ�͸�ѡ
    '����Ϊ��1���ı���Ŀ�Ͷ�ѡ��Ŀֻ�ܵ����涨��2.ÿ�а󶨵Ļ��Ŀ���ܳ���������3���󶨶����Ŀʱ��Ŀ��ʾ����Ŀ���ͱ���һ��
    '��Ŀ��ʾ����һ��
    For intDo = 1 To intCount
        mrsItems.Filter = "��Ŀ���=" & Val(lstColumnUsed.ListItems(intDo).Text)
        If mrsItems.RecordCount = 0 Then
            RaiseEvent AfterRowColChange("�����¼��Ŀ�����Ѿ������仯��������ˢ�»���ҳ��", True, mblnSign, mblnArchive)
            mrsItems.Filter = 0
            Exit Sub
        End If
        If intDo = 1 Then
            intFace = Val(NVL(mrsItems!��Ŀ��ʾ))
            intType = Val(NVL(mrsItems!��Ŀ����))
        Else
            If Not (intFace = Val(NVL(mrsItems!��Ŀ��ʾ)) And intType = Val(NVL(mrsItems!��Ŀ����))) Then
                RaiseEvent AfterRowColChange("�󶨵�������Ŀ�ı�ʾ����Ŀ���ͱ���һ�£���Ҫô����ѡ���Ҫô������ֵ¼���", True, mblnSign, mblnArchive)
                mrsItems.Filter = 0
                Exit Sub
            End If
            '�ı��Ͳ�����󶨶����Ŀ
            If Val(NVL(mrsItems!��Ŀ����)) = 1 And Val(NVL(mrsItems!��Ŀ��ʾ)) = 0 Then
                RaiseEvent AfterRowColChange("һ��ֻ�ܰ�һ���ı��ͻ��Ŀ��", True, mblnSign, mblnArchive)
                mrsItems.Filter = 0
                Exit Sub
            End If
            '��ѡ��Ŀֻ�ܵ�����
            If Val(NVL(mrsItems!��Ŀ��ʾ, 0)) = 3 Then
                RaiseEvent AfterRowColChange("һ��ֻ�ܰ�һ����ѡ�ͻ��Ŀ��", True, mblnSign, mblnArchive)
                mrsItems.Filter = 0
                Exit Sub
            End If
        End If
        
        'ƴ��
        strTest = lstColumnUsed.ListItems(intDo).Text
        '47764,������,2012-08-13,���Ŀû�в�λ�����ڲ�ͬ��û�п��Ƶ����ܰ���ͬ��Ŀ
'        If lstColumnUsed.ListItems(intDo).SubItems(2) <> "" Then
'            strTest = strTest & "," & lstColumnUsed.ListItems(intDo).SubItems(2)
'        End If
        strTest = strTest & "," & lstColumnUsed.ListItems(intDo).SubItems(2)
        If ISActiveUsed(strTest) Then Exit Sub
        
        strPara = strPara & IIf(intDo > 1, ";", "") & strTest
        mrsItems.Filter = 0
    Next
    
    '61852:������,2013-11-05,��ӻ��Ŀ���汾ҳ֮ǰ�ı������
    If Not DataMap_Save Then picCloumn.Visible = False: Exit Sub
    
    '��������
    gstrSQL = "ZL_���˻���ҳ��_UPDATE(" & mlng�ļ�ID & "," & mintҳ�� & "," & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ",'" & strPara & "','" & gstrUserName & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "������Ŀ������")
    picCloumn.Visible = False
    lngCol = VsfData.COL
    lngRow = VsfData.ROW
    
    '���²�ѯSQL
    '������ȡ����
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
'����:ɾ�����ڷǿ����ϵĻ��Ŀ����Ϣ
'����:������,2013-07-16
'�����:63401
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
            '��¼�������������ϵĻ��Ŀ������Ϣ
            ReDim Preserve arrActive(UBound(arrActive) + 1)
            arrActive(UBound(arrActive)) = CStr(arrData(intDo))
        Else
            '��¼��Ҫ�Ƴ��Ļ��Ŀ�к�
            ReDim Preserve arrCol(UBound(arrCol) + 1)
            arrCol(UBound(arrCol)) = lngCol
        End If
    Next
    
    On Error GoTo ErrHand
    
    'ɾ������Ҫ�Ļ��Ŀ��Ϣ(��Ҫ������֮ǰ���������,�����������)
    If UBound(arrCol) > 1 Then
        gcnOracle.BeginTrans
        blnTran = True
    End If
    
    For intDo = 0 To UBound(arrCol)
        If CStr(arrCol(intDo)) <> "" Then
            strSQL = "ZL_���˻���ҳ��_UPDATE(" & mlng�ļ�ID & "," & mintҳ�� & "," & Val(arrCol(intDo)) & ",NULL,'" & gstrUserName & "')"
            Call zlDatabase.ExecuteProcedure(strSQL, "������Ŀ������")
        End If
    Next
    If blnTran = True Then gcnOracle.CommitTrans
    
    '���¸�����ȡ�Ļ��Ŀ����Ϣ
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
    '���ĳ�����Ŀ�Ƿ��ѱ������а�
    ISActiveUsed = True
    
    arrData = Split(mstrCOLActive, "||")
    intCount = UBound(arrData)
    For intDo = 0 To intCount
        arrCol = Split(Split(arrData(intDo), "|")(1), ";")
        lngCol = Split(Split(arrData(intDo), "|")(0), ";")(0)
        intMax = UBound(arrCol)
        For intIn = 0 To intMax
            If strTest = arrCol(intIn) And VsfData.COL - (cHideCols + VsfData.FixedCols - 1) <> lngCol Then
                RaiseEvent AfterRowColChange(Split(strTest, ",")(1) & mrsItems!��Ŀ���� & " �Ѿ����󶨵�" & lngCol & "�У��������ظ��󶨣�", True, mblnSign, mblnArchive)
                Exit Function
            End If
        Next
    Next
    ISActiveUsed = False
End Function

Private Function GetActivePart(ByVal intFindCol As Integer, ByVal intItem As Integer) As String
    '��ȡָ���еĻ��Ŀ
    Dim arrData
    Dim arrCol
    Dim intCol As Integer, strPart As String
    Dim intDo As Integer, intCount As Integer
    Dim intIn As Integer, intMax As Integer
    '�����Ŀ���뵽��ѯSQL�У���ʽ���к�;��ͷ����|��Ŀ���,��λ;��Ŀ���,��λ||�к�;��ͷ����...
    '�󶨶����Ŀ�����о��Զ�תΪ�Խ�����
    
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
    
    gstrSQL = " SELECT  SUM(��¼����) AS ����" & _
              " From ���˻�����ϸ A,���˻������� B," & vbNewLine & _
              "      (Select ��� From ���������Ŀ Start With ���=[2] Connect By Prior ���=�����) C" & vbNewLine & _
              " Where A.��¼ID=B.ID And A.��ֹ�汾 Is NULL And A.��¼����=1 AND B.�������=0 And A.��Ŀ���=C.���" & vbNewLine & _
              " And B.�ļ�ID=[1] And B.����ʱ�� Between [3] And [4]"
    Call SQLDIY(gstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��������", mlng�ļ�ID, lngItem, CDate(strStart), CDate(strEnd))
    strCollect = NVL(rsTemp!����)
    
    CalcCollect = strCollect
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function CaclCategorical(ByVal strItem As String, ByVal strStart As String, ByVal strEnd As String) As ADODB.Recordset
'--------------------------------------------------------
'����:��ȡ������Ŀ���������Ϣ
'����:
'    strItem:������Ŀ�ͻ�����Ŀ����Ŀ���,��ʽ:6:7
'    strStart:���ܿ�ʼʱ��
'    strEnd:���ܽ���ʱ��
'--------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    
    gstrSQL = "Select /*+ Rule */ ����ʱ��, ��Ŀ����, ��Ŀ����" & vbNewLine & _
        " From (Select Min(����ʱ��) ����ʱ��, ��Ŀ����, Sum(��Ŀ����) ��Ŀ����" & vbNewLine & _
        "       From (Select ����ʱ��, Max(��Ŀ����) ��Ŀ����, Nvl(Max(��Ŀ����), 0) ��Ŀ����" & vbNewLine & _
        "              From (Select a.Id, a.����ʱ��, Decode(b.��Ŀ���, d.C1, b.��¼����, Null) ��Ŀ����, Decode(b.��Ŀ���, d.C2, b.��¼����, Null) ��Ŀ����" & vbNewLine & _
        "                     From ���˻������� a, ���˻�����ϸ b, ���˻����ӡ c," & vbNewLine & _
        "                          (Select C1, C2 From Table(Cast(f_Num2list2([2]) As Zltools.t_Numlist2))) d" & vbNewLine & _
        "                     Where a.Id = b.��¼id And a.Id = c.��¼id And Nvl(a.�������, 0) = 0 And a.�ļ�id = [1] And" & vbNewLine & _
        "                           (b.��Ŀ��� = d.C1 Or b.��Ŀ��� = d.C2) And b.��ֹ�汾 Is NULL And b.��¼����=1 And " & vbNewLine & _
        "                           a.����ʱ�� Between [3] And [4])" & vbNewLine & _
        "              Group By Id, ����ʱ��)" & vbNewLine & _
        "       Group By ��Ŀ����)" & vbNewLine & _
        " Order By ����ʱ��"
    Call SQLDIY(gstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�����������", mlng�ļ�ID, strItem, CDate(strStart), CDate(strEnd))
    Set CaclCategorical = rsTemp
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub cmdOK_Click()
    Dim arrTime, arrColCollect 'Ҫ���ܵ���
    Dim arrItem, arrData
    Dim intType As Integer      'С�����
    Dim arrValue, arrCorrelative() As String
    Dim bln���� As Boolean, blnExit As Boolean
    Dim lngStart As Long
    Dim lngCol As Long, lngCount As Long, lngRow As Long, lngRows As Long, i As Long, j As Long
    Dim strToday As String, str����ʱ�� As String
    Dim strStartDate As String, strEndDate As String
    Dim strStartTime As String, strEndTime As String
    Dim strKey As String, strField As String, strValue As String, strtmp As String
    Dim lngMaxIndex As Long, intDatas As Integer
    Dim rsCategorical As New ADODB.Recordset, lngMaxMutilRows As Long, lngMutilRows As Long, arrMutilRows
    
    On Error GoTo ErrHand
    '����һ���µĻ��ܼ�¼
    
    If cboС��.Text = "��ʱ" And Val(txtС������.Tag) = 0 Then
        RaiseEvent AfterRowColChange("��ʼʱ������ʱ���ʽ����ȷ��", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    
    If InStr(1, txtС������.Text, ";") <> 0 Then
        RaiseEvent AfterRowColChange("С�������в��ܺ��зֺţ�", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    If InStr(1, txtС������.Text, "'") <> 0 Then
        RaiseEvent AfterRowColChange("С�������в��ܺ��е����ţ�", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    If LenB(StrConv(txtС������.Text, vbFromUnicode)) > 50 Then
        RaiseEvent AfterRowColChange("С�����Ʋ��ܳ���50���ַ���25�����֣�", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    
    '������Ҫѡ��һ��������Ŀ
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
        RaiseEvent AfterRowColChange("����Ҫѡ��һ�������У�", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    '���ʱ�䷶Χ�Ƿ����
    '��ָ��������Ϊ׼
    '    �� ����
    '    ҹ ���� - ����
    '    ȫ ���� - ����
    strToday = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    arrTime = Split(cboС��.Tag, ";")   '��ʽ:��ʼʱ��,����ʱ��;��ʼʱ��,����ʱ��
    strStartTime = txt��ʼʱ��.Text
    strEndTime = txt����ʱ��.Text
    If strEndTime < strStartTime Then bln���� = True
    If bln���� = True Then
        strStartDate = strToday & " " & strStartTime & ":00"
        strEndDate = DateAdd("d", 1, CDate(strToday)) & " " & strEndTime & IIf(cboС��.Text <> "��ʱ", ":59", ":00")
    Else
        strStartDate = strToday & " " & strStartTime & ":00"
        strEndDate = strToday & " " & strEndTime & IIf(cboС��.Text <> "��ʱ", ":59", ":00")
    End If

    strStartDate = Format(DateAdd("d", -1 * DateDiff("d", CDate(DTPDate.Value), CDate(strToday)), CDate(strStartDate)), "yyyy-MM-dd HH:mm:ss")
    strEndDate = Format(DateAdd("d", -1 * DateDiff("d", CDate(DTPDate.Value), CDate(strToday)), CDate(strEndDate)), "yyyy-MM-dd HH:mm:ss")
    
    '63765:������,2013-11-22,����С�ᱣ���ʱ��
    lngMaxIndex = cboС��.ListCount
    If cboС��.Text <> "��ʱ" Then
        'С�ᷢ��ʱ�����С�����ʱ�䣭С������
        intType = -1 * cboС��.ItemData(cboС��.ListIndex)
        If cboС��.ItemData(cboС��.ListIndex) = 999 Then 'ȫ��С��-1s
            str����ʱ�� = Format(DateAdd("s", -1, strEndDate), "YYYY-MM-DD HH:mm:ss")
        Else
            str����ʱ�� = Format(DateAdd("s", -1 * (lngMaxIndex - cboС��.ListIndex), strEndDate), "YYYY-MM-DD HH:mm:ss")
        End If
    Else
        '��ʱС��Ϊ-998
        intType = -1 * 998
        '55892:������,2012-11-30,��ʱС�����ʱ��-1S���磺8:00-18:00 ָ�ľ��ǻ���8�㵽17:59:59��
        strEndDate = Format(DateAdd("s", -1, strEndDate), "YYYY-MM-DD HH:mm:ss")
        str����ʱ�� = strEndDate
        strEndTime = Format(strEndDate, "HH:mm")
    End If
    
    
    '����Ƿ��Ѿ����ڸ�����
    blnExit = False
    mrsDataMap.Filter = "ɾ��=0 And �������=" & intType & " And ��������='" & str����ʱ�� & "'"    '��¼ID>0������,���ǵ��������
    blnExit = (mrsDataMap.RecordCount)
    mrsDataMap.Filter = 0
    
    If blnExit Then
        RaiseEvent AfterRowColChange("��Ҫ��ӵ�С�������Ѵ��ڣ�", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    
    If CDate(Format(strEndDate, "YYYY-MM-DD HH:mm:ss")) < CDate(Format(mstr��ʼʱ��, "YYYY-MM-DD HH:mm:ss")) Then
       RaiseEvent AfterRowColChange("С��Ľ���ʱ�䲻��С���ļ���ʼʱ��:[" & CDate(Format(mstr��ʼʱ��, "YYYY-MM-DD HH:mm:ss")) & "]", True, mblnSign, mblnArchive)
       Exit Sub
    End If
    
    '���ҿհ���
    lngRows = VsfData.Rows - 1
    For lngRow = VsfData.FixedRows To lngRows
        If Val(VsfData.TextMatrix(lngRow, mlngRecord)) = 0 And VsfData.TextMatrix(lngRow, mlngRowCount) = "" Then
            lngStart = lngRow
            Exit For
        End If
    Next
    If lngStart = 0 Then
        '˵��û���ҵ��հ���
        VsfData.Rows = VsfData.Rows + 1
        lngStart = VsfData.Rows - 1
    End If
    
    'ͳ�ƻ�������(�����ݿ��л���,��ǰ����ֻ��¼���Ƿ��޸�,����֪��ԭֵ�Ƕ���,���Ե�ǰδ��������ݲ�����)
    '������Ŀ����
    '������Ŀ�м���:col;1|col;4,5
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
    
    'ͨ�ò���
    VsfData.TextMatrix(lngStart, mlngYear) = txtС������.Text
    VsfData.TextMatrix(lngStart, mlngDate) = txtС������.Text
    VsfData.TextMatrix(lngStart, mlngTime) = txtС������.Text
    VsfData.TextMatrix(lngStart, mlngRowCount) = "1|1"                          'Ϊ�˱�֤ʱ�䲻�ظ�,��ȡ����ʱ��+��ķ�ʽ
    VsfData.TextMatrix(lngStart, mlngRowCurrent) = "1"
    VsfData.TextMatrix(lngStart, mlngCollectText) = txtС������.Text
    VsfData.TextMatrix(lngStart, mlngCollectType) = intType                     '��ʾС��;-1�װ�;-2ҹ��;3-ȫ��
    VsfData.TextMatrix(lngStart, mlngCollectStyle) = cbo��ʶ.ListIndex         '����24Сʱ,���»�����
    VsfData.TextMatrix(lngStart, mlngCollectDay) = str����ʱ��
    VsfData.TextMatrix(lngStart, mlngCollectStart) = strStartTime
    VsfData.TextMatrix(lngStart, mlngCollectEnd) = strEndTime
    VsfData.MergeRow(lngStart) = True
    'ͬ������������ʱ���е�����
    strField = "ID|ҳ��|�к�|�к�|��ʼ�к�|��¼ID|����|����|��¼���|ɾ��"
    '1\����
    strKey = mintҳ�� & "," & lngStart & "," & mlngDate
    strValue = strKey & "|" & mintҳ�� & "|" & lngStart & "|" & mlngDate & "|" & lngStart & "|" & Val(VsfData.TextMatrix(lngStart, mlngRecord)) & "|" & _
            txtС������.Text & ";" & intType & ";" & cbo��ʶ.ListIndex & ";" & str����ʱ�� & ";" & strStartTime & ";" & strEndTime & "|1|0|0"
    Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
    
    'չ��
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
            
            '52953,������,2012-08-24,��������Ϊ0ҲҪ��ʾ,�������������кϲ�����������:60792
            If lngCol + cHideCols > mlngTime And lngCol + cHideCols < mlngNoEditor Then
                VsfData.TextMatrix(lngStart, lngCol + cHideCols) = FormatValue(VsfData.TextMatrix(lngStart, lngCol + cHideCols))
                If Trim(VsfData.TextMatrix(lngStart, lngCol + cHideCols)) <> "" Then
                    '66085:������,2012-09-26,�������ڻ����кϲ�,��ԭ����������+�ո�ͬһ�ĳ����к�����chr(13)
                    '������ӿո���п�������������ʾ����ȫ(��Ҫ����Ҷ���)
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
            
            strKey = mintҳ�� & "," & lngStart & "," & lngCol + cHideCols
            strValue = strKey & "|" & mintҳ�� & "|" & lngStart & "|" & lngCol + cHideCols & "|" & lngStart & "|" & _
                Val(VsfData.TextMatrix(lngStart, mlngRecord)) & "|" & VsfData.TextMatrix(lngStart, lngCol + cHideCols) & "|1|0|0"
            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
        End If
    Next
    '����������ݴ���
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
            '���ڷ�����ܣ����������й�������������д����"����"
            If rsCategorical.RecordCount > 0 Then
                lngCol = Split(arrCorrelative(0), ",")(0) + cHideCols + VsfData.FixedCols - 1
                VsfData.TextMatrix(lngStart, lngCol) = "����"
                strKey = mintҳ�� & "," & lngStart & "," & lngCol
                strValue = strKey & "|" & mintҳ�� & "|" & lngStart & "|" & lngCol & "|" & lngStart & "|" & _
                    Val(VsfData.TextMatrix(lngStart, mlngRecord)) & "|" & VsfData.TextMatrix(lngStart, lngCol) & "|1|0|0"
                Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
            End If
            With rsCategorical
                Do While Not .EOF
                    lngMutilRows = 0
                    For i = 0 To 1
                        lngRows = lngRow
                        lngCol = Split(arrCorrelative(i), ",")(0) + cHideCols + VsfData.FixedCols - 1
                        strtmp = IIf(i = 0, CStr(NVL(!��Ŀ����)), CStr(NVL(!��Ŀ����)))
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
                        
                        strKey = mintҳ�� & "," & lngRow & "," & lngCol
                        strValue = strKey & "|" & mintҳ�� & "|" & lngRow & "|" & lngCol & "|" & lngStart & "|" & _
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
    '��ȡ����и�
    lngMaxMutilRows = 0
    For i = 0 To UBound(arrMutilRows)
        If lngMaxMutilRows < Val(arrMutilRows(i)) Then lngMaxMutilRows = Val(arrMutilRows(i))
    Next i
    If lngMaxMutilRows > 0 Then
        lngMaxMutilRows = lngMaxMutilRows + Val(VsfData.TextMatrix(lngStart, mlngRowCurrent))
        For lngRow = lngStart To lngStart + lngMaxMutilRows - 1
            VsfData.TextMatrix(lngRow, mlngRowCount) = lngMaxMutilRows & "|" & lngRow - lngStart + 1                       'Ϊ�˱�֤ʱ�䲻�ظ�,��ȡ����ʱ��+��ķ�ʽ
            VsfData.TextMatrix(lngRow, mlngRowCurrent) = lngMaxMutilRows
            VsfData.TextMatrix(lngRow, mlngCollectText) = VsfData.TextMatrix(lngStart, mlngCollectText)
            VsfData.TextMatrix(lngRow, mlngCollectType) = VsfData.TextMatrix(lngStart, mlngCollectType)
            VsfData.TextMatrix(lngRow, mlngCollectStyle) = VsfData.TextMatrix(lngStart, mlngCollectStyle)
            VsfData.TextMatrix(lngRow, mlngCollectDay) = VsfData.TextMatrix(lngStart, mlngCollectDay)
            VsfData.TextMatrix(lngRow, mlngCollectStart) = VsfData.TextMatrix(lngStart, mlngCollectStart)
            VsfData.TextMatrix(lngRow, mlngCollectEnd) = VsfData.TextMatrix(lngStart, mlngCollectEnd)
        Next lngRow
    End If
    
'    '�ϲ���Ԫ��
'    lngRows = Split(Split(mstrColCollect, "|")(0), ";")(0) + cHideCols - 1
'    For lngRow = mlngTime + 1 To lngRows
'        VsfData.TextMatrix(lngStart, lngRow) = txtС������.Text
'    Next
'    VsfData.MergeCells = flexMergeRestrictRows          '���ᵥԪ��Ȼ�ǵ����ϲ�,�ϲ�����������ϲ���Ԫ��
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
    '�����ʾ�ѡ����
    
    If Val(cmdWord.Tag) = -1 Then
        strInput = txtInput.Text
    Else
        strInput = txt(Val(cmdWord.Tag)).Text
    End If
    strInput = frmEditAssistant.ShowMe(Me, mlng����ID, mlng��ҳID, mintӤ��, strInput)
    
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
    Dim strTag As String    'cboС���tag�б���ʱ��Σ���ʽ����ʼ,����;��ʼ,����
    Dim rsData As New ADODB.Recordset
    Dim i As Integer
    Dim strCurDate As String
    Dim intStart As Integer, intEnd As Integer, intCur As Integer, intIndex As Integer, intCount As Integer
    On Error GoTo ErrHand
    '��ʾС�ᴰ��
    
    If Not DataMap_Save Then Exit Sub       '��������,�Ա�ѡ��С���ʱ��������ݼ��
    '����¼���Ƿ���ڻ�����Ŀ�У�������������˳�
    If mstrCollectItems = "" Then
        RaiseEvent AfterRowColChange("��ǰ�ļ���δʹ�û�����Ŀ��", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    
    '��ȡ����ʱ��(���=3Ϊȫ��С��)
    strDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    gstrSQL = "Select NVL(����ID,0) AS ����ID,����,���,����,��ʼ,���� From �������ʱ�� Order by ��� "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡС��")
    rsTemp.Filter = "����=2"
    If rsTemp.RecordCount = 0 Then
        rsTemp.Filter = 0
        RaiseEvent AfterRowColChange("��δ���ü�¼��С��,�����ڻ�����Ŀ����ģ��Ļ�����Ŀ�����ã�", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    
    rsTemp.Filter = "����=1 And ���=3"
    If rsTemp.RecordCount = 0 Then
        rsTemp.Filter = 0
        RaiseEvent AfterRowColChange("ȫ�����ʱ��δ����,�����ڻ�����Ŀ����ģ��Ļ�����Ŀ�����ã�", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    strStart = NVL(rsTemp!��ʼ)
    strEnd = NVL(rsTemp!����)
    rsTemp.Filter = 0
    
    strCurDate = Format(zlDatabase.Currentdate, "YYYY-MM-DD")
    
    On Error Resume Next
    With DTPDate
        '�����:50201 LPF ȡ��ԭ�����Ƶ�һ����ʱ�䣬С��ʱ��������ļ���ǰ����Чʱ�䷶Χ����
        '.MinDate = Format(DateAdd("D", -30, CDate(strCurDate)), "YYYY-MM-DD")
        .MinDate = Format(mstr��ʼʱ��, "YYYY-MM-DD")
        If CDate(.MinDate) < CDate(Format(mstr��ʼʱ��, "YYYY-MM-DD")) Then .MinDate = Format(mstr��ʼʱ��, "YYYY-MM-DD")
        '��ȡ���˱䶯��¼�е����ʱ��
        gstrSQL = "SELECT MAX(NVL(��ֹʱ��, SYSDATE+" & mintPreDays & ")) AS ��Ժʱ��" & vbNewLine & _
            " FROM ���˱䶯��¼" & vbNewLine & _
            " WHERE ��ʼʱ�� IS NOT NULL AND ����ID =[1] AND ��ҳID =[2]"
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�������ʱ��", mlng����ID, mlng��ҳID)
        If rsData.RecordCount > 0 Then
            strMaxTime = Format(rsData!��Ժʱ��, "YYYY-MM-DD")
        Else
            strMaxTime = Format(DateAdd("D", mintPreDays, CDate(strCurDate)), "YYYY-MM-DD")
        End If
        '�ļ�����ʱ�䲻Ϊ�գ�������ڷ�Χ���ܲ������˵�ǰ��Ч���ʱ��
        If IsDate(mstr����ʱ��) Then
            If CDate(Format(strMaxTime, "YYYY-MM-DD")) > CDate(Format(mstr����ʱ��, "YYYY-MM-DD")) Then
                strMaxTime = Format(mstr����ʱ��, "YYYY-MM-DD")
            End If
        End If
        If CDate(Format(strMaxTime, "YYYY-MM-DD")) < CDate(Format(mstr��ʼʱ��, "YYYY-MM-DD")) Then
            strMaxTime = Format(mstr��ʼʱ��, "YYYY-MM-DD")
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
    
    '���ػ������(��¼��С�����3��)
    intIndex = 0
    intCount = 0
    intCur = Format(Now, "HH")
    cboС��.Clear
    If strStart <> "" Or strEnd <> "" Then
        cboС��.AddItem "ȫ��С��"
        cboС��.ItemData(cboС��.NewIndex) = 999
        strTag = strTag & ";" & strStart & "," & strEnd
    End If
    intCount = intCount + 1
    
    With rsTemp
        rsTemp.Filter = "���� = 2 And ����ID=" & mlng����ID
        If rsTemp.RecordCount = 0 Then rsTemp.Filter = "���� = 2 And ����ID=0"
        rsTemp.Sort = "��ʼ ASC"
        Do While Not .EOF
            If Not (NVL(!��ʼ) = "" Or NVL(!����) = "") Then
                cboС��.AddItem !����
                cboС��.ItemData(cboС��.NewIndex) = Val(!���)
                strTag = strTag & ";" & !��ʼ & "," & !����
                
                '��λ��ǰʱ���Ӧ��С��
                intStart = Val(!��ʼ)
                intEnd = Val(!����)
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
            cboС��.Tag = Mid(strTag, 2)
            cboС��.ListIndex = intIndex
        Else
            rsTemp.Filter = 0
            RaiseEvent AfterRowColChange("����Ļ�����ȫ����ӣ�", True, mblnSign, mblnArchive)
            Exit Sub
        End If
        cboС��.AddItem "��ʱ"
    End With
    
    
    rsTemp.Filter = 0
    
    With cbo��ʶ
        .Clear
        .AddItem "������"
        .AddItem "���»����߱�ʶ"
        .AddItem "����ֵ�·���˫���߱�ʶ"
        .AddItem "�Ϸ������߱�ʶ"
        .AddItem "����ֵ�·��������߱�ʶ"
        If mintCollectDef > 0 And mintCollectDef < .ListCount Then
            .ListIndex = mintCollectDef
        Else
            .ListIndex = 0
        End If
    End With
    
    Call LoadCollectItem
    '��������
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
    cboС��.SetFocus
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
    '������ͷ
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
        
        .TextMatrix(0, 0) = "������"
        .TextMatrix(1, 0) = "������"
        .TextMatrix(2, 0) = "������"
        .TextMatrix(3, 0) = "�Ƿ����"
        
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

        '���̶ܹ��е��и߲���ȷ��Ҫ�Զ�������
        .AutoResize = True
        .WordWrap = True
        .AutoSizeMode = flexAutoSizeRowHeight
        .AutoSize 0, .Cols - 1
        .AutoResize = False
        
        .Cell(flexcpAlignment, 0, .FixedCols, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
        .Cell(flexcpPictureAlignment, .FixedRows, .FixedCols, .FixedRows, .Cols - 1) = flexAlignCenterCenter
        
        '87428:�����еĸ߶����µ������߶ȣ���Ϊ���̶��еĸ߶�+����������ĸ߶ȴ��ڱ��߶ȣ�����������޷���ʾ
        '�����Ǵ���һ�пɱ༭��Ŀǰ����Ϊ��ȫ��ʾ�������������
        .Height = 1100 'Ĭ�ϳ�ʼ�߶�
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
    cboС��_Click
End Sub

Private Sub DTPDate_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyReturn Then cboС��.SetFocus
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
                strInfo = i + 1 & "��" & CStr(arrCode(i))
            Else
                strInfo = strInfo & vbCrLf & i + 1 & "��" & CStr(arrCode(i))
            End If
        Next
    End If
    If UBound(arrCode) >= 0 Then
        strInfo = strInfo & vbCrLf & vbCrLf & "˵��������ʾ��Ϣ��׼ȷ����������������ȫ���档"
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
    
    '61855:������,2013-11-07,�󶨻��Ŀ��ô����������
    strText = Trim(txtFind.Text)
    If KeyCode = 10 Or strText = "" Then
        '��Ҫ�������������ֵ
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
    '��ȡ���������
    intRows = SendMessage(txtLst.hWnd, EM_GETLINECOUNT, 0&, 0&)
    lngCurHeight = PicLst.TextHeight("��") * intRows + PicLst.TextHeight("��") * 1 / 3
    lngHeight = txtLst.Height
    lngDiffHeight = lngCurHeight - lngHeight
    If lngCurHeight < Val(txtLst.Tag) Then lngCurHeight = Val(txtLst.Height)
    txtLst.Height = lngCurHeight
    lbllst(1).Top = txtLst.Top + txtLst.Height
    lstSelect(mintType - 1).Top = lbllst(1).Top + lbllst(1).Height + 20
    PicLst.Height = lstSelect(mintType - 1).Top + PicLst.TextHeight("��") * (lstSelect(mintType - 1).ListCount) + PicLst.TextHeight("��") \ 3
    If PicLst.Top + PicLst.Height + picMain.Top > ScaleHeight Then
        If ScaleHeight - PicLst.Top - picMain.Top < 0 Then
            PicLst.Top = 10
            PicLst.Height = ScaleHeight - picMain.Top - 10
        Else
            PicLst.Height = ScaleHeight - picMain.Top
        End If
    End If
    lstSelect(mintType - 1).Height = IIf(PicLst.Height - lstSelect(mintType - 1).Top < PicLst.TextHeight("��") + PicLst.TextHeight("��") \ 3, PicLst.TextHeight("��") + PicLst.TextHeight("��") \ 3, PicLst.Height - lstSelect(mintType - 1).Top)
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
        '�ƶ�����һ����Ԫ��
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
    If Val(txtPage.Text) < mint��ʼҳ�� Or Val(txtPage.Text) > mint����ҳ + 1 Then
        MsgBox "�����ҳ����Ч����ǰ�ļ�ҳ����Чҳ�뷶Χ����" & mint��ʼҳ�� & "ҳ �� ��" & mint����ҳ + 1 & "ҳ��", vbInformation, gstrSysName
        txtPage.SetFocus
        Exit Sub
    End If
    
    If Not DataMap_Save Then Exit Sub
    
    '���²�ѯSQL
    '������ȡ����
    mintҳ�� = Val(txtPage.Text)
    mblnInit = False
    Call InitVariable
    Call InitCons
    Call ReadStruDef
    Call zlRefresh
    
    mblnInit = True
    VsfData.Refresh
    
    mcbrPage.Caption = "ҳ��ѡ�񣺵�" & mintҳ�� & "ҳ"
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
    '���в��ɼ�ʱ����ʾ˵����Ϣ
    If VsfData.RowIsVisible(VsfData.ROW) = True And VsfData.ColIsVisible(VsfData.COL) = True Then
         '��ʾ��ǰ��Ŀ�������Ϣ
        mrsSelItems.Filter = "��=" & VsfData.COL - (cHideCols + VsfData.FixedCols - 1)
        If mrsSelItems.RecordCount <> 0 Then
            mrsItems.Filter = "��Ŀ���=" & mrsSelItems!��Ŀ���
            If mrsItems.RecordCount <> 0 Then
                strItemInfo = Trim(NVL(mrsItems!˵��, ""))
            End If
        End If
        mrsSelItems.Filter = 0
        mrsItems.Filter = 0
    End If
     '--48659:������,2012-09-14,����ֶ�'˵��'
    RaiseEvent ShowTipInfo(VsfData, strItemInfo, True)
'    '����̶��еĸ߶�
'    For lngRow = 0 To 2
'        If Not VsfData.RowHidden(lngRow) Then dblHeight = dblHeight + VsfData.ROWHEIGHT(lngRow)
'    Next
'    '�ӿɼ��п�ʼ���²������һ���ɼ���
'    For lngRow = NewTopRow To VsfData.Rows - 1
'        If Not VsfData.RowIsVisible(lngRow) Then
'            lngRow = lngRow - 1
'            Exit For
'        End If
'    Next
'    '�ӿɼ��п�ʼ�������һ���ɼ���
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
'        '��ǰ�����еĸ߶�+�̶��еĸ߶�������ڱ��ؼ��ĸ߶�,˵����ǰѡ��������д�����ס���ֵ����
'        If VsfData.Row >= lngRow - 1 And CellRect.Bottom * (lngRow - NewTopRow + 1) + dblHeight >= VsfData.ClientHeight Then
'            '��ס���ֵ������
'            VsfData.Row = VsfData.Row + IIf(OldTopRow < NewTopRow, 1, -1)
'        End If
'    End If
'
'    If Not VsfData.ColIsVisible(VsfData.Col) Then
'        VsfData.Col = VsfData.Col + IIf(OldLeftCol < NewLeftCol, 1, -1)
'    Else
'        '��ǰ�����еĸ߶�+�̶��еĸ߶�������ڱ��ؼ��ĸ߶�,˵����ǰѡ��������д�����ס���ֵ����
'        If VsfData.Col = lngCol And dblWidth >= VsfData.ClientWidth Then
'            '��ס���ֵ������
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
    '��������ʾ��¼��ؼ�
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
    
    'δ������в�����¼������
    mintType = -1
    
    If mblnInit = False Then Exit Sub
    
    Call ShowSignMarker
    
    If InStr(1, mstrPrivs, "�����¼�Ǽ�") = 0 Then Exit Sub
    
    If VsfData.TextMatrix(VsfData.ROW, mlngRowCount) = "" Then
        lngStart = VsfData.ROW
    Else
        lngStart = GetStartRow(VsfData.ROW)
    End If
    
    '������ܹ�����������ȷ��
    arrCorrelative = Split(mstrColCorrelative, "|")
    For i = 0 To UBound(arrCorrelative)
        strCorrelative = strCorrelative & "," & Val(Split(arrCorrelative(i), ",")(0))
    Next i
    strCorrelative = Mid(strCorrelative, 2)
    If Val(VsfData.TextMatrix(lngStart, mlngCollectType)) < 0 And _
        (VsfData.COL <= mlngTime And IIf(mblnVerify = True, VsfData.COL <> mlngChoose, True) Or _
        InStr(1, "|" & mstrColCollect, "|" & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ";") <> 0 Or _
        InStr(1, "," & strCorrelative & ",", "," & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ",") <> 0) Then
        RaiseEvent AfterRowColChange("�����в������޸�����ʱ��,�Լ������кͻ����й����е����ݣ�", True, mblnSign, mblnArchive)
        Exit Sub '�����в������޸�����ʱ��,�Լ������кͻ����й��������е�����
    End If
    
    If mblnVerify Then  '�������mblnShow�ж���������
        If VsfData.COL = mlngChoose Then Call vsfdata_KeyDown(vbKeySpace, 0): Exit Sub
        '81535:������,��ǩʱ�����޸�����ʱ����
        'If VsfData.COL = mlngDate Or VsfData.COL = mlngTime Or VsfData.COL = mlngYear Then Exit Sub
        If Val(VsfData.TextMatrix(lngStart, mlngRecord)) = 0 Then Exit Sub
        If VsfData.TextMatrix(lngStart, mlngSigner) = "" Then Exit Sub
        If VsfData.Cell(flexcpChecked, lngStart, mlngChoose) = flexTSUnchecked Then Exit Sub 'û��ѡ�еļ�¼���ܱ༭
    Else
        '��ǩ��������ֻ������ǩ״̬���޸�
        If InStr(1, VsfData.TextMatrix(lngStart, mlngSigner), "/") <> 0 Then
            RaiseEvent AfterRowColChange("����ǩ�����ݲ�����༭��", True, mblnSign, mblnArchive)
            Exit Sub
        End If
        
        'ֻҪ��ǩ�����ݾͲ������޸�
        '--------------------------
'        '�����ǰ����Ա�ļ������ǩ������Ա�ļ����,��������༭����
'        If VsfData.TextMatrix(lngStart, mlngSigner) <> "" Then
'            If mintVerify > Val(VsfData.TextMatrix(lngStart, mlngSignLevel)) + 1 Then
'                RaiseEvent AfterRowColChange("��ǰ����Ա�ļ������ǩ������Ա�ļ����,������༭���ݣ�", True, mblnSign, mblnArchive)
'                Exit Sub
'            End If
'        End If
        If VsfData.TextMatrix(lngStart, mlngSigner) <> "" Then
            RaiseEvent AfterRowColChange("��ǩ�������ݲ������ٴα༭����ȡ��ǩ�������ԣ�", True, mblnSign, mblnArchive)
            Exit Sub
        End If
        '--------------------------
        '63706:������,2013-11-20,���һ��ʼ�հ󶨵��ǻ�ʿ,�����Ƿ��ڼ�¼�����˻�ʿ��
        'Ĭ��ǩ�����뱣������ͬ,�������޸����˻����¼Ȩ�޵Ĳ���Ա,�������޸����˵�����
        strName = FormatValue(VsfData.TextMatrix(lngStart, VsfData.Cols - 1))
        If strName <> "" Then
            If strName <> gstrUserName And _
                InStr(1, mstrPrivs, "���˻����¼") = 0 Then
                RaiseEvent AfterRowColChange("��û���޸����˻����¼���ݵ�Ȩ�ޣ�ԭ����Ա:" & strName, True, mblnSign, mblnArchive)
                Exit Sub
            End If
        End If
    End If
    If mblnArchive Then Exit Sub
    If Not mblnShow Or Not mblnEditable Then Exit Sub
    If VsfData.TextMatrix(lngStart, mlngDemo) <> "" Then
        'ֻ��������δ��������ݣ��������޸�������ʱ��
        If (VsfData.COL = mlngDate Or VsfData.COL = mlngTime Or VsfData.COL = mlngYear) Then
            If Val(VsfData.TextMatrix(lngStart, mlngDemo)) > 1 Then
                Exit Sub
            Else
                'If Val(VsfData.TextMatrix(lngStart, mlngRecord)) > 0 Then Exit Sub
            End If
        End If
    End If
    'δ����Ŀ�Ŀ��в��ܱ༭
    If (InStr(1, "," & mstrCOLNothing & ",", "," & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ",") <> 0) Then
        mrsSelItems.Filter = "��=" & VsfData.COL - (cHideCols + VsfData.FixedCols - 1)
        If mrsSelItems.RecordCount = 0 Then
            RaiseEvent AfterRowColChange("������༭Ϊ����Ŀ���У����ֹ���ӻ��Ŀ��", True, mblnSign, mblnArchive)
            Exit Sub
        End If
    End If
    '��ҳ�����в���������н���ճ��,ɾ��,ֻ�ܱ༭�����Ŀ�����
    If InStr(1, VsfData.TextMatrix(lngStart, mlngRowCount), "|") <> 0 And lngStart = 3 And Val(VsfData.TextMatrix(lngStart, mlngStartRowPage)) <> mintҳ�� Then
        If Val(Split(VsfData.TextMatrix(lngStart, mlngRowCount), "|")(0)) <> Val(VsfData.TextMatrix(lngStart, mlngRowCurrent)) Then
            If InStr(1, "," & mstrCOLNothing & ",", "," & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ",") <> 0 Then
                RaiseEvent AfterRowColChange("�������޸Ŀ�ҳ�����еĻ��Ŀ���ݣ�", True, mblnSign, mblnArchive)
                Exit Sub
            End If
        End If
    End If
    'ͬ�������в�����༭
    strCols = GetSynItems(2, intMax)
    If strCols <> "" Then
        '����ͬ�����ݵ���,������ʱ���ǲ������޸ĵ�
        strCols = "," & strCols & ","
        If VsfData.COL = mlngDate Or VsfData.COL = mlngTime Or VsfData.COL = mlngYear Or _
            InStr(1, strCols, "," & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ",") <> 0 Then
            RaiseEvent AfterRowColChange("ͬ���������У��������޸�ʱ���ͬ���������У�", True, mblnSign, mblnArchive)
            Exit Sub
        End If
    End If
    '�������漰������ϸ,�����������ǩ����������в������޸�
    If mstrColCollect <> "" Then
        If Val(VsfData.TextMatrix(lngStart, mlngRecord)) > 0 Then
            '�������޸Ļ��������ݣ�Ҳ�������޸�������ʱ��
            If InStr(1, "|" & mstrColCollect, "|" & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ";") <> 0 Then
                blnCheck = True
            ElseIf InStr(1, "|" & mstrColCorrelative, "|" & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ",") <> 0 And mstrColCorrelative <> "" Then
                blnCheck = True
            ElseIf VsfData.COL <= mlngTime Then
                blnCheck = True
            End If
            If blnCheck = True Then
                If ISCollectSigned(mlng�ļ�ID, Mid(VsfData.TextMatrix(lngStart, mlngActTime), 1, 10), Format(VsfData.TextMatrix(lngStart, mlngActTime), "HH:MM")) Then
                    RaiseEvent AfterRowColChange("������������Ӧ�Ļ�����������ǩ�����������޸ĵ�ǰ�����л�����ʱ���У�", True, mblnSign, mblnArchive)
                    Exit Sub
                End If
            End If
        End If
    End If
    
    If VsfData.COL <= mlngNoEditor - 1 Then Call ShowInput
    
    On Error Resume Next
    '�ÿؼ���ý���
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
    
    '63401:������,2013-07-16,���ѡ����в��ǻ�н����ػ��Ŀ����
    '                       ���ѡ��������������Ŀ�н����¼��ػ��Ŀ���ý���
    If OldCol <> NewCol And picCloumn.Visible = True Then
        If InStr(1, "," & mstrCOLNothing & ",", "," & NewCol - (cHideCols + VsfData.FixedCols - 1) & ",") <> 0 Then
            Call BoundItems(NewCol - (cHideCols + VsfData.FixedCols - 1))
        Else
            picCloumn.Visible = False
        End If
    End If
    
    'ѡ����,ͬ��������ֱ���˳�,����˴������ʾ��Ϣ
    If NewCol = mlngChoose Then Exit Sub
    strCols = GetSynItems(2, intMax)
    If strCols <> "" Then
        strCols = "," & strCols & ","
        If InStr(1, strCols, "," & NewCol - (cHideCols + VsfData.FixedCols - 1) & ",") <> 0 Then blnExit = True
    End If
    
    '��ʾ��ǰ��Ŀ�������Ϣ
    mrsSelItems.Filter = "��=" & NewCol - (cHideCols + VsfData.FixedCols - 1)
    If mrsSelItems.RecordCount <> 0 Then
        mrsItems.Filter = "��Ŀ���=" & mrsSelItems!��Ŀ���
        If mrsItems.RecordCount <> 0 Then
            If NVL(mrsItems!��Ŀֵ��) <> "" Then
                If mrsItems!��Ŀ���� = 0 Then
                    strInfo = "��Ч��Χ:" & Split(mrsItems!��Ŀֵ��, ";")(0) & "��" & Split(mrsItems!��Ŀֵ��, ";")(1)
                Else
                    strInfo = "��Ч��Χ:" & mrsItems!��Ŀֵ��
                End If
            Else
                strInfo = ""
            End If
            strItemInfo = Trim(NVL(mrsItems!˵��, ""))
        End If
    End If
    mrsSelItems.Filter = 0
    mrsItems.Filter = 0
    
    '--48659:������,2012-09-14,����ֶ�'˵��'
    RaiseEvent ShowTipInfo(VsfData, strItemInfo, True)
    
    If blnExit = True Then Exit Sub
    
    '����Ƿ���ǩ��
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
        'ֻ��ѡ��ʼ��
        lngStart = GetStartRow(VsfData.ROW)
        If VsfData.TextMatrix(lngStart, mlngTime) = "" And Val(VsfData.TextMatrix(lngStart, mlngDemo)) <= 1 Then Exit Sub
        
        If mintVerify = δ���� Then
            RaiseEvent AfterRowColChange("����ǰ��δ����Ƹ�μ���ְ��������Ա���������ã�", True, mblnSign, mblnArchive)
            Exit Sub
        End If
        '��ǩʱ,��ǰ��¼��ǩ��,�Ҳ���Ա��ǩ��������ϴ�ǩ������߲�����
        If VsfData.TextMatrix(lngStart, mlngSignLevel) = "" Then
            RaiseEvent AfterRowColChange("�����ݻ�δǩ�������ܽ�����ǩ��", True, mblnSign, mblnArchive)
            Exit Sub
        Else
            intLevel = Val(VsfData.TextMatrix(lngStart, mlngSignLevel)) + 1
        End If
         '43588:������,2012-09-13,��Ӽ�¼����ǩģʽ
        If InStr(1, VsfData.TextMatrix(lngStart, mlngSigner), "/") <> 0 And mintSignMode = 1 Then
            RaiseEvent AfterRowColChange("��ǰ��ǩģʽΪ��1-��ǩȨ�ޡ�����ֻ�ܶ�δ��ǩ�����ݽ��в�����" & vbCrLf & "��ϸ��Ϣ�����е������Ѿ���ǩ", True, mblnSign, mblnArchive)
            Exit Sub
        End If
        
        If mintVerify >= intLevel And mintSignMode = 0 Then
            RaiseEvent AfterRowColChange("���ļ���[" & GetVerify(mintVerify) & "]Ҫ���ϴ���ǩ�˵ļ���[" & GetVerify(intLevel) & "]�߲��ܹ�ѡ�ü�¼��", True, mblnSign, mblnArchive)
            Exit Sub
        End If
        
        '���ڷ���������ǩʱ����Ҫѡ�񱾷���������
        lngDemo = Val(VsfData.TextMatrix(lngStart, mlngDemo))
        If lngDemo > 1 Then 'Ѱ����ʼ��
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
            
            'ѡ����ǷǷ�����ʼ�У�Ҫ�ȼ����ʼ��
            If VsfData.TextMatrix(lngStart, mlngSignLevel) = "" Then
                RaiseEvent AfterRowColChange("�÷�����ʼ���е����ݻ�δǩ��������ǩ�����ڽ�����ǩ��", True, mblnSign, mblnArchive)
                Exit Sub
            Else
                intLevel = Val(VsfData.TextMatrix(lngStart, mlngSignLevel)) + 1
            End If
            '43588:������,2012-09-13,��Ӽ�¼����ǩģʽ
            If InStr(1, VsfData.TextMatrix(lngStart, mlngSigner), "/") <> 0 And mintSignMode = 1 Then
                RaiseEvent AfterRowColChange("��ǰ��ǩģʽΪ��1-��ǩȨ�ޡ�����ֻ�ܶ�δ��ǩ�����ݽ��в�����" & vbCrLf & "��ϸ��Ϣ���÷�����ʼ�е������Ѿ���ǩ", True, mblnSign, mblnArchive)
                Exit Sub
            End If
            If mintVerify >= intLevel And mintSignMode = 0 Then
                RaiseEvent AfterRowColChange("���ļ���[" & GetVerify(mintVerify) & "]Ҫ�ȸ÷�����ʼ�е���ǩ��ǩ���˵ļ���[" & GetVerify(intLevel) & "]�߲��ܹ�ѡ�ü�¼��", True, mblnSign, mblnArchive)
                Exit Sub
            End If
        End If
        arrRow = Array()
        ReDim Preserve arrRow(UBound(arrRow) + 1)
        arrRow(UBound(arrRow)) = lngStart
        lngRowCount = Val(Split(VsfData.TextMatrix(lngStart, mlngRowCount), "|")(0))
        lngNextGroupRow = lngStart + lngRowCount - 1
        '1.�����������Ƿ��Ѿ�ǩ����2.��¼�������ݿ�ʼ��
        For lngRow = lngStart + lngRowCount To VsfData.Rows - 1
            '������ÿһ�������������ݿ���ռ�ö��У�ֻ���ڵ�һ�б����˷�����������������������=ÿһ������������+ÿһ���������е�����
            If lngRow > lngNextGroupRow Then
                If Val(VsfData.TextMatrix(lngRow, mlngDemo)) <= 1 Then Exit For      '����������·�����˳�
                If VsfData.TextMatrix(lngRow, mlngSignLevel) = "" Then
                    RaiseEvent AfterRowColChange("�÷����д���δǩ�������ݣ�����ǩ�����ڽ�����ǩ��", True, mblnSign, mblnArchive)
                    Exit Sub
                End If
                intLevel = Val(VsfData.TextMatrix(lngRow, mlngSignLevel)) + 1
                '43588:������,2012-09-13,��Ӽ�¼����ǩģʽ
                If InStr(1, VsfData.TextMatrix(lngRow, mlngSigner), "/") <> 0 And mintSignMode = 1 Then
                    RaiseEvent AfterRowColChange("��ǰ��ǩģʽΪ��1-��ǩȨ�ޡ�����ֻ�ܶ�δ��ǩ�����ݽ��в�����" & vbCrLf & "��ϸ��Ϣ���÷����еڡ�" & UBound(arrRow) + 2 & "����������Ѿ���ǩ", True, mblnSign, mblnArchive)
                    Exit Sub
                End If
                If mintVerify >= intLevel And mintSignMode = 0 Then
                    RaiseEvent AfterRowColChange("���ļ���[" & GetVerify(mintVerify) & "]Ҫ�ȸ÷����еڡ�" & UBound(arrRow) + 2 & "�������ǩ��ǩ���˵ļ���[" & GetVerify(intLevel) & "]�߲��ܹ�ѡ�ü�¼��", True, mblnSign, mblnArchive)
                    Exit Sub
                End If
                lngNextGroupRow = Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(0)) + lngRow - 1
                ReDim Preserve arrRow(UBound(arrRow) + 1)
                arrRow(UBound(arrRow)) = lngRow
            End If
        Next lngRow
        'ѡ�����з�������
        For lngRow = 0 To UBound(arrRow)
            lngStart = Val(arrRow(lngRow))
            VsfData.Cell(flexcpChecked, lngStart, mlngChoose) = IIf(VsfData.Cell(flexcpChecked, lngStart, mlngChoose) = flexTSChecked, flexTSUnchecked, flexTSChecked)
            '�����޸ļ�¼�Ա�ͬ��
            strField = "ID|ҳ��|�к�|�к�|��¼ID|����|ɾ��"
            strKey = mintҳ�� & "," & lngStart & "," & mlngChoose
            strValue = strKey & "|" & mintҳ�� & "|" & lngStart & "|" & mlngChoose & "|" & Val(VsfData.TextMatrix(lngStart, mlngRecord)) & "|" & VsfData.Cell(flexcpChecked, lngStart, mlngChoose) & "|1"
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
'��ȡ��ǰ�����Ӧ��Ƹ�μ���ְ��
    Dim strVerify As String
    Select Case intVerify
        Case ����
            strVerify = "���λ�ʦ"
        Case ����
            strVerify = "�����λ�ʦ"
        Case �м�
            strVerify = "���ܻ�ʦ"
        Case ʦ��
            strVerify = "��ʦ"
        Case Աʿ
            strVerify = "��ʿ"
        Case Else
            strVerify = "δ����"
    End Select
    GetVerify = strVerify
End Function

Private Sub InitVariable()
    '�������
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
    '��������ؼ�
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
    '���ܣ�
    '������
    '���أ�
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
    cbsThis.ActiveMenuBar.Title = "�˵���"
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
    '�˵�����
    cbsThis.ActiveMenuBar.Title = "�˵�"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    
        '------------------------------------------------------------------------------------------------------------------
        '����������
        Set cbrToolBar = cbsThis.Add("��׼", xtpBarTop)
        cbrToolBar.ShowTextBelowIcons = False
        cbrToolBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
        cbrToolBar.ModifyStyle XTP_CBRS_GRIPPER, 0
        With cbrToolBar.Controls
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Group_New, "����"): cbrControl.IconId = 3096: cbrControl.ToolTipText = "��ʼ����"
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Group_Append, "׷��"): cbrControl.IconId = 3045: cbrControl.ToolTipText = "׷�ӷ���(Ctrl+A)"
            
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Copy, "����"): cbrControl.ToolTipText = "����(Ctrl+C)": cbrControl.BeginGroup = True
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_PASTE, "ճ��"):  cbrControl.ToolTipText = "ճ��(Ctrl+V)"
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Clear, "���"):   cbrControl.ToolTipText = "���"
            
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_SPECIALCHAR, "�������"):  cbrControl.ToolTipText = "�����������(Ctrl+D)": cbrControl.BeginGroup = True
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Word, "�ʾ�ѡ��"):  cbrControl.ToolTipText = "�ʾ�ѡ��(Ctrl+W)"
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Import, "��������"):  cbrControl.ToolTipText = "��������(Ctrl+I)"
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Brief, "С��"): cbrControl.ToolTipText = "С��"
            
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Element, "��ǩҪ��"): cbrControl.IconId = conMenu_Edit_Append: cbrControl.BeginGroup = True: cbrControl.ToolTipText = "�Զ����ǩҪ������¼��"
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Append, "�а�"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "�а�"
        
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_PrevPage, "��ҳ"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "��ҳ"
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NextPage, "��ҳ"):   cbrControl.ToolTipText = "��ҳ"
        End With
    
        For Each cbrControl In cbrToolBar.Controls
            If cbrControl.Type <> xtpControlCustom And cbrControl.Type <> xtpControlLabel Then
                cbrControl.Style = xtpButtonIconAndCaption
            End If
        Next
        Set mcbrToolBar = cbrToolBar
    
         '�����
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
    '���ʱ��¼���Ƿ�Ϸ�������������
    
    If Trim(objText.Text) <> "" Then
        strInput = Trim(objText.Text)
        If InStr(1, strInput, ":") > 0 Then
            strHour = Split(strInput, ":")(0)
            strMin = Split(strInput, ":")(1)
        ElseIf InStr(1, strInput, "��") > 0 Then
            strHour = Split(strInput, "��")(0)
            strMin = Split(strInput, "��")(1)
        Else
            strHour = strInput
            strMin = "00"
        End If
        strHour = Format(strHour, "00")
        strMin = Format(strMin, "00")
        If Not IsNumeric(strHour) Then
            RaiseEvent AfterRowColChange("��ʼʱ���к��зǷ��ַ���", True, mblnSign, mblnArchive)
            Exit Function
        End If
        If Val(strHour) < 0 Or Val(strHour) > 23 Then
            RaiseEvent AfterRowColChange("��ʼʱ�㲻��ȷ��СʱֵӦ��>0��С��24��", True, mblnSign, mblnArchive)
            Exit Function
        End If
        If Not IsNumeric(strMin) Then
            RaiseEvent AfterRowColChange("��ʼʱ���к��зǷ��ַ���", True, mblnSign, mblnArchive)
            Exit Function
        End If
        If Val(strMin) < 0 Or Val(strMin) > 59 Then
            RaiseEvent AfterRowColChange("��ʼʱ�㲻��ȷ������ֵӦ��>0��С��60��", True, mblnSign, mblnArchive)
            Exit Function
        End If
        strInput = strHour & ":" & strMin
    End If
    CheckTxtTime = strInput
End Function

Private Function CheckTime(ByVal lngRow As Long, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal intӤ�� As Integer, _
    ByVal strTime As String, ByVal strCurTime As String, ByRef strMsg As String) As Boolean
    Dim blnMsg As Boolean
    Dim blnExist As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strBabyOutTime As String
    On Error GoTo ErrHand
    '���ݷ���ʱ������ڵ�ǰ���ҵ���Чʱ�䷶Χ��
    
    blnMsg = (strMsg <> "")
    
    '����ļ���ʼ,����ʱ��
    If Format(strTime, "YYYY-MM-DD HH:mm") < Format(mstr��ʼʱ��, "yyyy-MM-dd HH:mm") Then
        strMsg = "����ʱ�䲻��С���ļ���ʼʱ��[" & mstr��ʼʱ�� & "]"
        GoTo exitHand
    End If
    If mstr����ʱ�� <> "" Then
        If Format(strTime, "YYYY-MM-DD HH:mm") > Format(mstr����ʱ��, "yyyy-MM-dd HH:mm") Then
            strMsg = "����ʱ�䲻�ܴ����ļ�����ʱ��[" & mstr����ʱ�� & "]"
            GoTo exitHand
        End If
    End If
    
    '75760:������,����Ӥ�����ڳ�Ժҽ�������
    If intӤ�� <> 0 Then
        strBabyOutTime = GetAdviceOutTime(lng����ID, lng��ҳID, intӤ��)
        If strBabyOutTime <> "" Then
            If Format(strTime, "YYYY-MM-DD HH:mm") > Format(strBabyOutTime, "YYYY-MM-DD HH:mm") Then
                strMsg = "��" & lngRow & "�еķ���ʱ��" & strTime & "����[����ʱ�䲻�ܴ��ڳ�Ժʱ��:" & Format(strBabyOutTime, "YYYY-MM-DD HH:mm") & "]"
                GoTo exitHand
            End If
            '��¼Сʱ���
            If Format(DateAdd("H", glngHours, strBabyOutTime), "yyyy-MM-dd HH:mm") < Format(strCurTime, "yyyy-MM-dd HH:mm") Then
                strMsg = "��" & lngRow & "�еķ���ʱ��" & strTime & "����[�������ݲ�¼����Чʱ��:" & glngHours & "Сʱ]"
                GoTo exitHand
            End If
            CheckTime = True
            Exit Function
        End If
    End If
    
    '���ݲ��˱䶯��¼���м��
    gstrSQL = " Select   ��ʼԭ��,����ID,to_char(��ʼʱ��,'yyyy-MM-dd hh24:mi') AS ��ʼʱ��,to_char(NVL(��ֹʱ��,sysDate+" & mintPreDays & "),'yyyy-MM-dd hh24:mi') AS ��ֹʱ�� " & _
              " From ���˱䶯��¼ " & _
              " Where ����ID=[1] And ��ҳID=[2]" & _
              " Order by ��ʼʱ��,��ʼԭ��"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��ǰ������Чʱ�䷶Χ", lng����ID, lng��ҳID)
    With rsTemp
        .Filter = "����ID=" & mlng����ID
        Do While Not .EOF
            If Format(strTime, "YYYY-MM-DD HH:mm") >= Format(!��ʼʱ��, "YYYY-MM-DD HH:mm") And Format(strTime, "YYYY-MM-DD HH:mm") <= Format(!��ֹʱ��, "YYYY-MM-DD HH:mm") Then
                blnExist = True
                Exit Do
            End If
            .MoveNext
        Loop
        .Filter = 0
        '�ҵ��˾��˳�
        If blnExist Then
            If Not IsAllowInput(lng����ID, lng��ҳID, strTime, strCurTime) Then
                strMsg = "��" & lngRow & "�еķ���ʱ��" & strTime & "����[�������ݲ�¼����Чʱ��:" & glngHours & "Сʱ]"
                GoTo exitHand
            End If
            
            CheckTime = True
            Exit Function
        End If
        
        'û�ҵ�,������ԭ�����׼ȷ����ʾ
        .Filter = "��ʼԭ��=1"
        If .RecordCount <> 0 Then
            If !��ʼԭ�� = 1 And Format(strTime, "YYYY-MM-DD HH:mm") < Format(!��ʼʱ��, "YYYY-MM-DD HH:mm") Then
                strMsg = "��" & lngRow & "�еķ���ʱ��" & strTime & "����[����ʱ�䲻��С�ڲ�����Ժʱ��:" & !��ʼʱ�� & "]"
                GoTo exitHand
            End If
        End If
        .Filter = "��ʼԭ��=2"
        If .RecordCount <> 0 Then
            If !��ʼԭ�� = 2 And Format(strTime, "YYYY-MM-DD HH:mm") < Format(!��ʼʱ��, "YYYY-MM-DD HH:mm") Then
                strMsg = "��" & lngRow & "�еķ���ʱ��" & strTime & "����[����ʱ�䲻��С�ڲ������ʱ��:" & !��ʼʱ�� & "]"
                GoTo exitHand
            End If
        End If
        .Filter = "��ʼԭ��=10"
        If .RecordCount <> 0 Then
            If !��ʼԭ�� = 10 And Format(strTime, "YYYY-MM-DD HH:mm") > Format(!��ֹʱ��, "YYYY-MM-DD HH:mm") Then
                strMsg = "��" & lngRow & "�еķ���ʱ��" & strTime & "����[����ʱ�䲻�ܴ��ڳ�Ժʱ��:" & !��ֹʱ�� & "]"
                GoTo exitHand
            End If
        End If
        .Filter = 0
        '�������˵��
        strMsg = "��" & lngRow & "�еķ���ʱ��" & strTime & "����[���ڵ�ǰ��������Чʱ�䷶Χ��]"
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
    '���¼�����ݵĺϷ���(����Ҳ��Ϊ��һ���ַ�,���ǵ�������Ŀ�ȴ��ڲ���\�������Ϣ)
    '���ص�����,���һ�а󶨶����Ŀ,�Ե�������Ϊ�ָ���
    
    'mintType:0=�ı���¼��;1=��ѡ;2=��ѡ;3=ѡ��;4-Ѫѹ��һ�а���������Ŀ,���ʽ����Ѫѹ��������Ŀ;5=һ�а���������Ŀ�Ҿ���ѡ����Ŀ;
    '6=һ�а�N����Ŀ,�ֹ�¼��
    Select Case mintType
    Case 0
        strText = txtInput.Text
        strOrders = txtInput.Tag
    Case 1, 2   '���
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
    Case 3      '���
        strText = lblInput.Caption
    Case 5      '���
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
    
    '69355:������,2013-01-07,�����ڸ�ʽ�Ĵ���
    strCurrDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
    If VsfData.COL = mlngDate Then
        If mblnDateAd Then
            If Trim(strText) = "" Then
                strInfo = "���ڲ���Ϊ�գ�"
                Exit Function
            End If
            If InStr(1, strText, "/") = 0 Then
                strInfo = "���ڸ�ʽ������1��12�գ�12/01"
                Exit Function
            End If
            
            If VsfData.TextMatrix(VsfData.ROW, mlngYear) = "" Then
                strDate = Mid(strCurrDate, 1, 5) & ToStandDate(strText)
                '����Ƿ����༭֮ǰ��ʱ��(һ���µ�����)
                If CDate(strDate) > CDate(Mid(strCurrDate, 1, 10)) And Mid(strCurrDate, 6, 2) = "01" And Mid(strDate, 6, 2) = "12" Then
                    strDate = DateAdd("yyyy", -1, CDate(strDate))
                End If
            Else
                strDate = VsfData.TextMatrix(VsfData.ROW, mlngYear) & "-" & ToStandDate(strText)
            End If
            If Not IsDate(strDate) Then
                strInfo = "¼������ݲ��ǺϷ������ڣ���1��12�գ�12/01"
                Exit Function
            Else
                VsfData.TextMatrix(VsfData.ROW, mlngYear) = Format(strDate, "YYYY")
            End If
        Else
            If Trim(strText) = "" Then
                strInfo = "���ڲ���Ϊ�գ�"
                Exit Function
            End If
            If Not IsDate(strText) Then
                strInfo = "¼������ݲ��ǺϷ������ڣ���1��12�գ�2011-01-12"
                Exit Function
            End If
            strDate = Format(strText, "yyyy-MM-dd")
            VsfData.TextMatrix(VsfData.ROW, mlngYear) = Format(strDate, "YYYY")
        End If
        If Format(strDate, "YYYY-MM-DD") > Format(DateAdd("d", mintPreDays, CDate(strCurrDate)), "YYYY-MM-DD") Then
            strInfo = "¼��������ѳ�������[����¼��������" & mintPreDays & "��]��ָ���ķ�Χ��"
            Exit Function
        End If
        
'        If VsfData.TextMatrix(VsfData.ROW, mlngTime) <> "" Then
'            blnCheck = True
'            strDate = strDate & " " & VsfData.TextMatrix(VsfData.ROW, mlngTime)
'        End If
    Else
        If Trim(strText) = "" Then
            strInfo = "ʱ�䲻��Ϊ�գ�"
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
            strInfo = "¼���ʱ���ʽ�Ƿ���[Сʱ:����]"
            Exit Function
        Else
            If Len(Trim(arrTime(0))) < 2 Then arrTime(0) = String(2 - Len(Trim(arrTime(0))), "0") & Trim(arrTime(0))
            If Len(Trim(arrTime(1))) < 2 Then arrTime(1) = String(2 - Len(Trim(arrTime(1))), "0") & Trim(arrTime(1))
            strText = arrTime(0) & ":" & arrTime(1)
        End If
        
        '�Ϸ��Լ��
        If IsNumeric(arrTime(0)) = False Or IsNumeric(arrTime(1)) = False Or Len(Trim(arrTime(0))) > 2 Or Len(Trim(arrTime(1))) > 2 Then
            strInfo = "¼���ʱ���ʽ�Ƿ���[Сʱ:����]"
            Exit Function
        End If
        If Mid(strText, 3, 1) <> ":" Then
            strInfo = "¼���ʱ���ʽ�Ƿ���[Сʱ:����]"
            Exit Function
        End If
        If Val(arrTime(0)) < 0 Or Val(arrTime(0)) > 23 Then
            strInfo = "¼���ʱ���ʽ�Ƿ���[СʱӦ��0��23֮��]"
            Exit Function
        End If
        If Val(arrTime(1)) < 0 Or Val(arrTime(1)) > 59 Then
            strInfo = "¼���ʱ���ʽ�Ƿ���[����Ӧ��0��59֮��]"
            Exit Function
        End If
        
        '���кϷ��Լ��
        If VsfData.TextMatrix(VsfData.ROW, mlngDate) <> "" Then
            strDate = VsfData.TextMatrix(VsfData.ROW, mlngDate)
            If mblnDateAd Then
                If VsfData.TextMatrix(VsfData.ROW, mlngYear) = "" Then
                    strDate = Mid(strCurrDate, 1, 5) & ToStandDate(strDate)
                    '����Ƿ����༭֮ǰ��ʱ��(һ���µ�����)
                    If CDate(strDate) > CDate(Mid(strCurrDate, 1, 10)) And Mid(strCurrDate, 6, 2) = "01" And Mid(strDate, 6, 2) = "12" Then
                        strDate = DateAdd("yyyy", -1, CDate(strDate))
                    End If
                Else
                    strDate = VsfData.TextMatrix(VsfData.ROW, mlngYear) & "-" & ToStandDate(strDate)
                End If
                If IsDate(strDate) Then
                    VsfData.TextMatrix(VsfData.ROW, mlngYear) = Format(strDate, "YYYY")
                Else
                    strInfo = "¼������ݲ��ǺϷ������ڣ���1��12�գ�12/01"
                    Exit Function
                End If
            Else
                strDate = Format(VsfData.TextMatrix(VsfData.ROW, mlngDate), "yyyy-MM-dd")
            End If
            
            strDate = Format(strDate & " " & strText, "YYYY-MM-DD HH:mm:ss")
            '70990:������,2014-03-13,���ڲ�¼���������޸�
            If Format(strDate, "YYYY-MM-DD HH:mm") > Format(DateAdd("d", mintPreDays, CDate(strCurrDate)), "YYYY-MM-DD HH:mm") Then
                strInfo = "¼��������ѳ�������[����¼��������" & mintPreDays & "��]��ָ���ķ�Χ��"
                Exit Function
            End If
        
            blnCheck = True
        End If
    End If
    
    If blnCheck Then
        '��������¼�뻹���޸ĵ����� ���������ʷ���ݶ��������޸�
        gstrSQL = " Select 1 From ���˻������� Where �ļ�ID=[1] And ����ʱ��=[2] And ([3]=0 OR ID<>[3])"
        Set rsCheck = zlDatabase.OpenSQLRecord(gstrSQL, "��鷢��ʱ��", mlng�ļ�ID, CDate(strDate), Val(VsfData.TextMatrix(VsfData.ROW, mlngRecord)))
        If rsCheck.RecordCount > 0 Then
            strInfo = "��¼���ʱ���Ѿ�������ʷ���ݣ�"
            Exit Function
        End If
        
        '�����ݿ���û���ҵ�����ʼ���û�¼�������Ѱ��
        If Not CheckChangeDataTime(VsfData.ROW, strDate, strInfo) Then Exit Function
        
        '81535:�޸�ʱ��Ķ�Ӧ�Ļ���������������ݣ������Ƿ��Ѿ�������Ӧ��С�Ტ������ǩ��
        '����:����������ǿ�Ƽ��;���е�������ֻ��Ҫ���ʱ��仯������(����ܴ���A����Ա�ڿ�ʼ�޻����л��е�δǩ��������ʱ�������B����Աǩ���ˣ�A����Ա�ڱ�������)
        '˵���������������ǿ����޸��κ��У����е���������������Ѿ�ǩ���ǲ������޸�����ʱ���кͻ����еģ�δ������޸ĵ�ʱ����ͬһ�����ܷ�Χ���жϣ�
        blnCheck = True
        If Val(VsfData.TextMatrix(VsfData.ROW, mlngRecord)) > 0 Then
            If Format(VsfData.TextMatrix(VsfData.ROW, mlngActTime), "YYYY-MM-DD HH:mm") = Format(strDate, "YYYY-MM-DD HH:mm") Then
                blnCheck = False
            End If
        End If
        If blnCheck = True Then
            If CheckCollectIsData(VsfData.ROW) = True Then
                If ISCollectSigned(mlng�ļ�ID, Format(strDate, "YYYY-MM-DD"), Format(strDate, "HH:MM")) Then
                    strInfo = "��¼���ʱ������Ӧ�Ļ�����������ǩ����������������µĻ��������ݣ�"
                    Exit Function
                End If
            End If
        End If
        '���ݷ���ʱ�䲻���ڵ�ǰ����Ա�������ҵ���Чʱ����ǰ
        If Not CheckTime(VsfData.ROW, mlng����ID, mlng��ҳID, mintӤ��, strDate, strCurrDate, strInfo) Then
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
'�����¼���ʱ�䣬�Ƿ������е�ʱ����ͬ�������ͬ����ʾ����¼��
    Dim strDateHistory As String, strTimeHistory As String, strDatetime As String '�û��Ѿ�¼������ں�ʱ��
    Dim lngCurRow As Long, intPage As Integer, blnDel As Boolean, blnTrue As Boolean
    Dim strCurrDate As String, lngRecord As Long, strActiveTime As String
    Dim strRows As String, strPages As String, strTimes As String, lngCol As Long
    Dim arrRows
    On Error GoTo ErrHand

    strCurrDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    With mrsCellMap
        .Filter = "�к�=" & mlngDate & " OR �к�=" & mlngTime
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            If Not (lngCurRow = !�к� And intPage = !ҳ��) Then
                blnDel = False
endWork:
                If lngCurRow = lngRow And intPage = mintҳ�� Then GoTo ErrNext
                If lngCurRow > 0 Then
                    mrsDataMap.Filter = "ҳ��=" & intPage & " And �к�=" & lngCurRow
                    If mrsDataMap.RecordCount <> 0 Then
                        blnDel = (mrsDataMap!ɾ�� = 1)
                        If mintҳ�� = intPage Then
                            If blnDel = False Then
                                blnDel = VsfData.RowHidden(lngCurRow)
                            Else
                                mrsDataMap!ɾ�� = IIf(VsfData.RowHidden(lngCurRow) = True, 1, 0)
                                mrsDataMap.Update
                            End If
                        End If
                        
                        strActiveTime = Format(mrsDataMap.Fields(cControlFields + mlngActTime - VsfData.FixedCols).Value, "YYYY-MM-DD HH:mm:ss")
                    Else
                        '�Ǳ��༭ҳ������,MrsCellMap���м�¼��mrsDataMap�б�Ȼ���ڶ�Ӧ������
                        If mintҳ�� = intPage Then
                            blnDel = VsfData.RowHidden(lngCurRow)
                            strActiveTime = Format(VsfData.TextMatrix(lngCurRow, mlngActTime), "YYYY-MM-DD HH:mm:ss")
                        Else
                            strMsg = "��" & intPage & "ҳ����" & lngCurRow & "�е������ڲ�����,���顢��¼���β�����������лл��"
                            Exit Function
                        End If
                    End If
                    mrsDataMap.Filter = 0
                End If
                
                If blnTrue = True And strDatetime <> "" Then
                    If Format(strDatetime, "YYYY-MM-DD HH:mm:ss") = Format(strCurDate, "YYYY-MM-DD HH:mm:ss") Then
                        '������ͬʱ�������û��ɾ����ֱ�ӽ�����ʾ
                        If blnDel = False Then
                            strMsg = "��" & intPage & "ҳ����" & lngCurRow & "���Ѿ�������ͬʱ������ݣ����飡"
                            Exit Function
                        Else
                            If lngRecord > 0 Then '���������ɾ�������ʱ���ԭ��ʱ����ֱͬ����ʾ������ͬ�ָ�ʱ��Ϊԭ��ʱ��
                                If Format(strDatetime, "YYYY-MM-DD HH:mm:ss") = Format(strActiveTime, "YYYY-MM-DD HH:mm:ss") Then
                                    strMsg = "��¼���ʱ���Ѿ�������ʷ���ݣ�"
                                    Exit Function
                                Else '�ָ�ʱ��Ϊԭ��ʱ��
                                    mrsDataMap.Filter = "ҳ��=" & intPage & " And �к�=" & lngCurRow
                                    If mrsDataMap.RecordCount <> 0 Then
                                        strActiveTime = Format(mrsDataMap.Fields(cControlFields + mlngActTime - VsfData.FixedCols).Value, "YYYY-MM-DD HH:mm:ss")
                                        mrsDataMap.Fields(cControlFields + mlngDate - VsfData.FixedCols).Value = Format(strActiveTime, "YYYY-MM-DD")
                                        mrsDataMap.Fields(cControlFields + mlngTime - VsfData.FixedCols).Value = Mid(strActiveTime, 12, 5)
                                        mrsDataMap.Update
                                    Else
                                        '�Ǳ��༭ҳ������,MrsCellMap���м�¼��mrsDataMap�б�Ȼ���ڶ�Ӧ������
                                        '�����Ѿ����,�϶��Ǳ�ҳ
                                        If mintҳ�� = intPage Then
                                            VsfData.TextMatrix(lngCurRow, mlngDate) = Format(strActiveTime, "YYYY-MM-DD")
                                            VsfData.TextMatrix(lngCurRow, mlngTime) = Mid(strActiveTime, 12, 5)
                                        End If
                                    End If
                                    mrsDataMap.Filter = 0
                                    '��¼�кź�ҳ��
                                    strRows = strRows & "," & lngCurRow
                                    strPages = strPages & "," & intPage
                                    strTimes = strTimes & "," & strActiveTime
                                End If
                            Else 'δ���������ɾ����ֱ����ռ�¼��������Ϣ(������ҳ�š��кš�ɾ��)
                                mrsDataMap.Filter = "ҳ��=" & intPage & " And �к�=" & lngCurRow
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
                                If mintҳ�� = intPage Then
                                    For lngCol = VsfData.FixedCols To VsfData.Cols - 1
                                        VsfData.TextMatrix(lngCurRow, lngCol) = ""
                                    Next lngCol
                                End If
                                mrsDataMap.Filter = 0
                                '��¼�кź�ҳ��
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
                '����ֵ
                intPage = !ҳ��
                lngCurRow = !�к�
                strDateHistory = ""
                strTimeHistory = ""
                strDatetime = ""
                lngRecord = NVL(!��¼ID, 0)
                blnTrue = False
            End If
            
            If !�к� = mlngDate Then
                If NVL(!����, 0) <> 1 Then
                    strDateHistory = NVL(!����)
                    If strDateHistory <> "" Then
                        If mblnDateAd Then
                            If InStr(1, "|" & mstrYears & "|", "|" & Val(NVL(!��λ)) & "|") <> 0 Then
                                strDateHistory = NVL(!��λ) & "-" & ToStandDate(strDateHistory)
                            Else
                                strMsg = "��" & intPage & "ҳ����" & lngCurRow & "�е�[���]���ݴ������¼���β������貢������Ȼ������¼�����ݣ�лл��"
                                Exit Function
'                                strDateHistory = Mid(strCurrDate, 1, 5) & ToStandDate(strDateHistory)
'                                '����Ƿ����༭֮ǰ��ʱ��(һ���µ�����)
'                                If CDate(strDateHistory) > CDate(Mid(strCurrDate, 1, 10)) And Mid(strCurrDate, 6, 2) = "01" And Mid(strDateHistory, 6, 2) = "12" Then
'                                    strDateHistory = DateAdd("yyyy", -1, CDate(strDateHistory))
'                                End If
                            End If
                        Else
                            strDateHistory = Format(strDateHistory, "yyyy-MM-dd")
                        End If
                    End If
                End If
            Else 'ʱ����
                strTimeHistory = NVL(!����, "00:00")
                If strDateHistory = "" Then strDateHistory = Mid(strCurrDate, 1, 10)
                strDatetime = strDateHistory & " " & strTimeHistory & ":00"

                '����������ݣ�����ʱ����ͨ����������ֻ������+
                If Val(NVL(!��λ)) >= 1 Then
                    strDatetime = Mid(strDatetime, 1, 17) & String(2 - Len(!��λ), "0") & Val(!��λ) - 1
                End If
                strDatetime = Format(strDatetime, "YYYY-MM-DD HH:mm:ss")
                blnTrue = True
            End If
        .MoveNext
        Loop
        
        If blnTrue Then GoTo endWork
        mrsDataMap.Filter = 0
    End With
    
    '����mrsCellMap��¼��
    If Left(strRows, 1) = "," Then strRows = Mid(strRows, 2)
    If Left(strPages, 1) = "," Then strPages = Mid(strPages, 2)
    If Left(strTimes, 1) = "," Then strTimes = Mid(strTimes, 2)
    arrRows = Split(strRows, ",")
    For lngCurRow = 0 To UBound(arrRows)
        mrsCellMap.Filter = "ҳ��=" & Val(Split(strPages, ",")(lngCurRow)) & " And �к�=" & Val(arrRows(lngCurRow))
        If CStr(Split(strTimes, ",")(lngCurRow)) = "[LPF]" Then
            Do While Not mrsCellMap.EOF
                mrsCellMap.Delete
                mrsCellMap.Update
                mrsCellMap.MoveNext
            Loop
        Else
            Do While Not mrsCellMap.EOF
                If mrsCellMap!�к� = mlngDate Then
                    mrsCellMap!���� = Format(CStr(Split(strTimes, ",")(lngCurRow)), "YYYY-MM-DD")
                    mrsCellMap.Update
                ElseIf mrsCellMap!�к� = mlngTime Then
                    mrsCellMap!���� = Mid(CStr(Split(strTimes, ",")(lngCurRow)), 12, 5)
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
    
    '���и�ʽ��װ����
    mrsSelItems.Filter = "��=" & VsfData.COL - (cHideCols + VsfData.FixedCols - 1)
    If mrsSelItems.RecordCount <> 0 Then
        '�д��е�δ���ж���
        strFormat = NVL(mrsSelItems!��ʽ)   '{P[����]C}{...}
        strFormat1 = strFormat
    End If
    mrsSelItems.Filter = 0
    
    '�������
    arrData = Split(strReturn, "'")
    arrOrder = Split(strOrders, "'")
    j = UBound(arrData)
    For i = 0 To j
        strText = arrData(i)
        strOrders = arrOrder(i)
        If Val(strOrders) <> 0 Then
            mrsItems.Filter = "��Ŀ���=" & strOrders
            strName = GetActivePart(VsfData.COL, i) & UCase(mrsItems!��Ŀ����)
            If strText <> "" Then
                blnCheck = True
                blnNumber = False
                If blnCheck Then
                    If mrsItems!��Ŀ���� = 0 And InStr(1, "0,4", mrsItems!��Ŀ��ʾ) <> 0 Then
                        blnNumber = True
                        strText = Val(strText)
                        If NVL(mrsItems!��ĿС��, 0) <> 0 Then   '��������ͨ���ؼ���MaxLength�����Ƶ�
                            If InStr(1, strText, ".") <> 0 Then strText = Mid(strText, 1, InStr(1, strText, ".") - 1)
                            If Len(strText) > mrsItems!��Ŀ���� Then
                                mrsItems.Filter = 0
                                strInfo = "[" & strName & "]¼������ݳ����˺Ϸ����ȣ�"
                                Exit Function
                            End If
                            
                            strText = Val(arrData(i))
                            If InStr(1, strText, ".") <> 0 Then
                                strText = Mid(strText, InStr(1, strText, ".") + 1)
                                If Len(strText) > mrsItems!��ĿС�� Then
                                    mrsItems.Filter = 0
                                    strInfo = "[" & strName & "]¼���С�����ֳ����˺Ϸ����ȣ�"
                                    Exit Function
                                End If
                            End If
                            strText = Val(arrData(i))
                        End If
                        If mrsItems!��Ŀ��ʾ = 0 Then
                            If Not IsNull(mrsItems!��Ŀֵ��) Then
                                dblMin = Val(Split(mrsItems!��Ŀֵ��, ";")(0))
                                dblMax = Val(Split(mrsItems!��Ŀֵ��, ";")(1))
                                If Not (Val(strText) >= dblMin And Val(strText) <= dblMax) Then
                                    mrsItems.Filter = 0
                                    strInfo = "[" & strName & "]¼������ݲ���" & Format(dblMin, "#0.00") & "��" & Format(dblMax, "#0.00") & "����Ч��Χ��"
                                    Exit Function
                                End If
                            End If
                        End If
                    Else
                        If LenB(StrConv(strText, vbFromUnicode)) > mrsItems!��Ŀ���� Then
                            strInfo = "[" & strName & "]¼������ݳ�������󳤶ȣ�" & mrsItems!��Ŀ���� & "��"
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
                'ɾ������Ŀ
                If InStr(1, "," & mstrCOLNothing & ",", "," & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ",") = 0 Then
                    Call SubstrPro(strFormat, strName)
                Else
                    '����Ŀ������ʱ,�����ǰ�о��жԽ�������,�����
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
            mrsItems.Filter = "��Ŀ���=" & strOrders
            strName = GetActivePart(VsfData.COL, i) & mrsItems!��Ŀ����
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
    '��ȡ����Ŀ��ǰ��׺����
    Dim i As Integer
    Dim strOrders As String, strName As String
    For i = 0 To UBound(arrOrder)
        strOrders = CStr(arrOrder(i))
        If Val(strOrders) <> 0 Then
            mrsItems.Filter = "��Ŀ���=" & strOrders
            strName = GetActivePart(VsfData.COL, i) & mrsItems!��Ŀ����
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
    '����ǰһ����Ŀ�ĺ�׺����+��ǰ��Ŀ��ǰ׺���ŵ�λ��
    
    If strData = "" Then Exit Function
    strData = strData
    j = Len(strFormat)
    l = InStr(1, strFormat, "[" & strName & "]")
    If l = 0 Then Exit Function
    '�õ�ǰ׺
    For i = l To 1 Step -1
        If Mid(strFormat, i, 1) = "{" Then Exit For
    Next
    strQZ = Mid(strFormat, i + 1, l - i - 1)
    '�ҵ�����Ŀ��ʽ���еĽ�������
    i = l + Len(strName) + 2
    For r = i To j
        If Mid(strFormat, r, 1) = "}" Then Exit For
    Next
    '�õ���׺
    strHZ = Mid(strFormat, i, r - i)
    '�����׺Ϊ��,�������Ѱ����һ����Ŀ��ǰ׺����
    If strHZ = "" And r < j Then
        For r = r + 1 To j
            If Mid(strFormat, r, 1) = "[" Then Exit For
        Next
        strHZ = Mid(strFormat, InStr(i, strFormat, "{") + 1, r - InStr(i, strFormat, "{") - 1)
    End If
    'ȡ��ָ����Ŀ���������ݴ�
    If strHZ <> "" Then
        j = InStr(intPos, strData, strHZ) '��Ϊ������ȡ��,���ǵ��ָ���������ͬ�����,��¼��һ�ε����λ��,�´δ����λ������ȡ����
        If j = 0 Then
            '�п����м���ڻس����з�
            j = InStr(intPos, Replace(strData, vbCrLf, ""), strHZ)
            If j = 0 Then Exit Function
        End If
    End If
    strData = Mid(strData, intPos)
    'ǰ׺Ϊ��,������ǰѰ����һ����Ŀ�ĺ�׺����
'    If strQZ = "" And i > 1 And intPos > 1 Then
'        For i = i - 1 To 1 Step -1
'            If Mid(strFormat, i, 1) = "]" Then Exit For
'        Next
'        strQZ = Mid(strFormat, i + 1, InStr(i, strFormat, "}") - i - 1)
'    End If
    
    SubstrVal = SubstrAnaly(strData, strHZ, strQZ)
    intPos = intPos + Len(strQZ & SubstrVal & strHZ)
    '�������������ȥ���س����з�����,������ַ�����ԭ������
'    If strHZ <> "" Then
'
'        strData = Mid(strData, 1, InStr(1, Replace(strData, vbCrLf, ""), strHZ) - 1) '��������Ŀ�������
'        intPOS = i + Len(strHZ)
'    End If
'    If strQZ <> "" Then strData = Mid(strData, InStr(1, strData, strQZ) + Len(strQZ)) '��������Ŀ�������
'    SubstrVal = strData ' Replace(strData, vbCrLf, "")
End Function

Private Function SubstrAnaly(ByVal strData As String, ByVal strHZ As String, ByVal strQZ As String) As String
    Dim strText As String
    Dim strCompare As String           '�Աȴ�
    Dim intLen As Integer, intActLen As Integer           'ǰ׺/��׺�ĳ���
    Dim intPos As Integer, intEnd As Integer
    Dim lngASC As Long
    Dim blnFind As Boolean
    '�����س����з�����,�ո����±ȶ�
    
    strText = strData
    If strHZ <> "" Then
        '�Ѻ�׺ȥ��
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
        '�϶���
        strText = Mid(strText, 1, intPos)
    End If
    
    '��ȥ��ǰ׺
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
    'intType=0-ɾ��ָ����ʽ��;1-�õ�ָ����ʽ��
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
    '�޸��ˣ�LPF 2012-04-20
    '�޸����ݣ�����Ƿ�����ʼ�У�Ҳ����¼���������
    '----------------------------------------------
    Dim arrData
    Dim blnNULL As Boolean                      '�Ƿ�Ϊ����
    Dim blnGroup As Boolean                     '������
    Dim strDate As String, strTime As String, strYear As String    '�����׼�¼��������ʱ��
    Dim strReturn As String, strMsg As String, strPart As String
    Dim lngStart As Long, lngStartGroup As Long, lngMutilRows As Long, lngDeff As Long, intGroupFirstRows As Integer, intBound As Integer, intRowCount As Integer
    Dim intRow As Integer, intRowGroup As Integer, intCount As Integer, intNULL As Integer  '����ж��ٿ���
    Dim blnTrue As Boolean, blnDate As Boolean, strRows As String, strRowsDel As String
    Dim lngDemo As Long, lngLastNull As Long, lngLastNoNull As Long
    '��ֵȻ���ƶ�����һ����Ч��Ԫ��
    Dim strKey As String, strField As String, strValue As String, strAppend As String
    Dim blnCallback As Boolean, blnReseGroupAssistant As Boolean, blnGroupAddNum As Boolean '��������������
    '���ı��к�������Ϣ
    Dim varAssistant() As Variant, strAssistantCols As String
    On Error GoTo ErrHand
    blnReseGroupAssistant = False
    
    '�������,���ϸ���ٴε���Ҫ��¼��
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
        '��ǵ�ǰ��Ϊ������
        blnDate = (InStr(1, "," & mlngYear & "," & mlngDate & "," & mlngTime & ",", "," & VsfData.COL & ",") > 0)
        If mstrGroupRow <> "" And Val(VsfData.TextMatrix(VsfData.ROW, mlngRecord)) = 0 And Val(mstrGroupRow) <= VsfData.ROW Then
            VsfData.TextMatrix(VsfData.ROW, mlngDemo) = VsfData.ROW - Val(mstrGroupRow) + 1
            'blnGroup = True
        Else
            'blnGroup = ((VsfData.TextMatrix(VsfData.ROW, mlngDemo) = "1") And mblnEditAssistant) Or (Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) > 1) '����ı��в��Զ��ֽ�
        End If
        
        lngDemo = Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo))
        blnGroup = Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) >= 1
        If VsfData.TextMatrix(VsfData.ROW, mlngRowCount) = "" Then VsfData.TextMatrix(VsfData.ROW, mlngRowCount) = "1|1"
        'blnGroup = ((VsfData.TextMatrix(VsfData.ROW, mlngDemo) = "1") And (mblnEditAssistant Or blnDate)) Or (Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) > 1) '����ı��в��Զ��ֽ�
        '����޸ĵ��ǷǴ��ı��л�ʱ���еķ������ݣ�����޸����������Ƿ����仯������仯�͵��������ݴ�����������ͨ���ݴ���
        If blnGroup = True And Not (mblnEditAssistant = True Or blnDate = True) Then
            intGroupFirstRows = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0))
            lngStart = GetStartRow(VsfData.ROW)
            '����༭���Ƿ����������һ�������У�����ͨ���ݴ���
            If lngStart + intGroupFirstRows < VsfData.Rows Then
                If Val(VsfData.TextMatrix(lngStart + intGroupFirstRows, mlngDemo)) <= 1 Then
                    blnGroup = False: GoTo ErrBegin
                End If
            ElseIf lngStart + intGroupFirstRows >= VsfData.Rows Then
                blnGroup = False
                GoTo ErrBegin
            End If
            
            With txtLength
                '������ʱ���еĿ�Ȳ���,Ϊ�˱��ⷵ�ض���,ǿ������Ϊ5000
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
                '�õ���������ռ������е���(���������ı���Ŀ)
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
            '--58414,������,2013-04-10,��ӻ�ʿ��ǩ������ʾģʽ
            Call ReSingDataToStart(VsfData, lngStart, lngStart + lngMutilRows - 1)
        Else
            lngMutilRows = 1
            If VsfData.TextMatrix(VsfData.ROW, mlngDemo) = 1 And (mblnEditAssistant Or blnDate) Then
                '��¼������ʼ�е���������
                intGroupFirstRows = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0))
                intBound = VsfData.ROW + intGroupFirstRows - 1
                '--58414,������,2013-04-10,��ӻ�ʿ��ǩ������ʾģʽ
                Call ReSingDataToStart(VsfData, VsfData.ROW, intBound)
                
                For intCount = VsfData.ROW + intGroupFirstRows To VsfData.Rows - 1
                    '������ÿһ�������������ݿ���ռ�ö��У�ֻ���ڵ�һ�б����˷�����������������������=ÿһ������������+ÿһ���������е�����
                    If intCount > intBound Then
                        If Val(VsfData.TextMatrix(intCount, mlngDemo)) <= 1 Then Exit For      '����������·�����˳�
                        intBound = Val(Split(VsfData.TextMatrix(intCount, mlngRowCount), "|")(0)) + intCount - 1
                        '--58414,������,2013-04-10,��ӻ�ʿ��ǩ������ʾģʽ
                        Call ReSingDataToStart(VsfData, intCount, intBound)
                    End If
                    lngMutilRows = lngMutilRows + 1
                Next
                lngMutilRows = lngMutilRows + intGroupFirstRows - 1 '��֤����������׼ȷ��
            Else
                '��¼������ʼ�е���������
                If VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1 >= VsfData.FixedRows Then
                    intGroupFirstRows = Val(Split(VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngRowCount), "|")(0))
                End If
                intGroupFirstRows = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0))
                lngMutilRows = intGroupFirstRows
                '--58414,������,2013-04-10,��ӻ�ʿ��ǩ������ʾģʽ
                Call ReSingDataToStart(VsfData, VsfData.ROW, VsfData.ROW + lngMutilRows - 1)
            End If
            lngStart = VsfData.ROW
        End If
       
        '׼����ֵ
        With txtLength
            '������ʱ���еĿ�Ȳ���,Ϊ�˱��ⷵ�ض���,ǿ������Ϊ5000
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
        '���������п��ܴ������ص���(����������ʱ),ֻ����ѡ����ı��Ŵ���
        If blnGroup = True And mblnEditAssistant = True And blnDate = False Then
            If Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) = 1 Then
                intNULL = intGroupFirstRows
                lngDeff = VsfData.ROW + intGroupFirstRows - 1
                For intRow = VsfData.ROW + intGroupFirstRows To VsfData.Rows - 1
                    If intRow > lngDeff Then
                        If Val(VsfData.TextMatrix(intRow, mlngDemo)) <= 1 Or intNULL > intCount Then Exit For     '����������·�����˳�
                        lngDeff = Val(Split(VsfData.TextMatrix(intRow, mlngRowCount), "|")(0)) + intRow - 1
                    End If
                    If VsfData.RowHidden(intRow) = True Then  'ɾ���ķ�����
                        '������֯�ı����飬��֤����������������
                        ReDim Preserve arrData(UBound(arrData) + 1)
                        For intBound = UBound(arrData) To intRow - VsfData.ROW + 1 Step -1
                            arrData(intBound) = arrData(intBound - 1)
                        Next intBound
                        arrData(intRow - VsfData.ROW) = ""
                        '��¼���һ�����ص���
                        lngLastNull = intRow
                    Else
                        intNULL = intNULL + 1
                        '��¼���һ��û�����ص���
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
            '����������������ʱ������Ҫ��¼����������ݲ���¼����ı�������
            If mblnEditAssistant = True And Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) = 1 And lngMutilRows = intGroupFirstRows Then
                strMsg = "������������ʱ������������ݵķ��飬�����¼����Ķ���Ŀ���ݣ�"
                RaiseEvent AfterRowColChange(strMsg, True, mblnSign, mblnArchive)
                strMsg = ""
                Exit Function
            End If
            '������������,�������������������������ӵ�����
            '20110830���������ͬһ�����У��������ı��ֽ⵽���У�������ı�����ͳһ�������һ����;�ڷ����а��س�,ֻ���������ݽ����޸�,�����з����仯
            intNULL = intCount - (lngMutilRows - 1)
            For intRow = lngMutilRows To intCount
                '��֤��ǰ�����������һҳ����ʾȫ
                If intRow + lngStart > VsfData.Rows - 1 Then Exit For
                
                If Val(VsfData.TextMatrix(intRow + lngStart, mlngRecord)) = 0 And VsfData.TextMatrix(intRow + lngStart, mlngRowCount) = "" Then
                    intNULL = intNULL - 1
                    If VsfData.RowHidden(intRow + lngStart) = True Then VsfData.RowHidden(intRow + lngStart) = False
                Else
                    Exit For
                End If
            Next
            '�����ӿ���
            If intNULL > 0 Then
                lngDeff = intNULL
                VsfData.Rows = VsfData.Rows + intNULL
                '�ӵ�ǰ�м�¼�Ŀհ��п�ʼ��ÿ�е�λ��+�����ӵĿհ�����
                For intRow = VsfData.Rows - intNULL - 1 To lngStart + intCount - intNULL + 1 Step -1
                    VsfData.RowPosition(intRow) = intRow + intNULL
                Next
            End If
            
            '���кŷ����仯����ͬ������mrsCellMap�д��ڸ��кŵ��к�����
            If lngDeff <> 0 Then
                If Not blnGroup Then
                    Call CellMap_Update(lngStart, lngDeff)
                Else
                    Call CellMap_Update(lngStart + lngMutilRows - 1, lngDeff)  '���������ݴ����һ����ϸ��֮��ʼ����
                End If
            End If
            '�Է����������һ��������Ϊ�����еĴ���
            '���磺�÷������2�����飬��һ��Ϊһ�д��ı�����ΪA���ڶ���Ϊһ�д��ı�����ΪC(��������).��ʱ��Ӵ��ı�����ΪA��Bռ���У���ʱ��֯�õ����������ΪA��C��B
            '���㷽ʽΪ:��һ��ռ������+����������+��������ݡ��˴��ͻ�������з���������õ�������ΪA��B��C��������ѭ����ֵ�У���һ���Ϊռ��2������ΪA��B
            '˵��������м���������У����һ��û�����أ���������ݾͻ�׷�������һ�����ݵĺ���
            If (lngLastNull - lngLastNoNull) > 0 Then
                For intRow = lngLastNoNull + 1 To lngLastNull
                    strValue = arrData(lngLastNoNull + 1 - VsfData.ROW)
                    For intBound = lngLastNoNull + 1 - VsfData.ROW To UBound(arrData) - 1
                        arrData(intBound) = arrData(intBound + 1)
                    Next intBound
                    arrData(UBound(arrData)) = strValue
                    VsfData.RowPosition(lngLastNoNull + 1) = lngLastNull + (intCount - (lngMutilRows - 1))
                Next intRow
                '���¼�¼
                For intRow = lngLastNull To lngLastNoNull + 1 Step -1
                    Call CellMap_Update(intRow, intCount - (lngMutilRows - 1), False)
                Next intRow
            End If
            'ѭ����ֵ
            intCount = UBound(arrData)
            intBound = 0
            blnReseGroupAssistant = (blnGroup = True And Not (mblnEditAssistant Or blnDate))
            blnGroupAddNum = blnReseGroupAssistant
            If blnGroup = True And blnDate = False Then strReturn = ""
            For intRow = 0 To intCount
                VsfData.TextMatrix(lngStart + intRow, VsfData.COL) = Replace(Replace(Replace(arrData(intRow), Chr(10), ""), Chr(13), ""), Chr(1), "")
                '�޸ķǷ������ݻ�Ǵ��ı������ڵķ�������
                '�Ƿ������ݣ�ֱ�Ӵ��������������������
                '�Ǵ��ı������ڵķ������ݣ�1��ֱ�Ӵ�������������������ݣ�2����Ҫ���´�����Ķε�������ʾλ��
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
                    '�޸ķ�����Ķλ����ڣ���ӷ�����ʼ�е���������д��������ı�������ʾ������
                    '�����е����⴦��,�����ڲ���¼���Ĵ���϶�
                    '##########################################
                    If Val(VsfData.TextMatrix(lngStart + intRow, mlngDemo)) >= 1 Then
                        If (mblnEditAssistant = True Or blnDate) And Val(VsfData.TextMatrix(lngStart, mlngDemo)) = 1 Then
                            VsfData.TextMatrix(lngStart + intRow, mlngDemo) = intRow + 1
                        End If
                        intRowCount = 1
                        '��ȡ�÷��������е�����
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
                    '���÷��������������һ�в�ִ�и��²���
                    If intRow = intBound + intRowCount - 1 Then
                        If mlngSignName <> -1 Then VsfData.TextMatrix(lngStart + intRow, mlngSignName) = VsfData.TextMatrix(lngStart + intBound, mlngSignName)
                        If mlngSignTime <> -1 Then VsfData.TextMatrix(lngStart + intRow, mlngSignTime) = VsfData.TextMatrix(lngStart + intBound, mlngSignTime)
                        '��������
                        strYear = ""
                        If Val(VsfData.TextMatrix(lngStart + intBound, mlngDemo)) > 0 Then
                            If CheckGroupDate(lngStart + intBound) = True Then
                                '�������޸ĲŽ�������̣�ȡ������¼��ʵ��ʱ��
                                If mblnDateAd Then
                                    strYear = Format(VsfData.TextMatrix(lngStart + intBound, mlngActTime), "YYYY")
                                    strDate = Format(VsfData.TextMatrix(lngStart + intBound, mlngActTime), "DD") & "/" & Format(VsfData.TextMatrix(lngStart + intBound, mlngActTime), "MM")
                                Else
                                    strDate = Mid(VsfData.TextMatrix(lngStart + intBound, mlngActTime), 1, 10)
                                End If
                                strTime = Mid(VsfData.TextMatrix(lngStart + intBound, mlngActTime), 12, 5)
                            Else
                                '����ʱ���������
                                strDate = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngDate)
                                strTime = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngTime)
                                If mblnDateAd Then strYear = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngYear)
                            End If
                        Else
                            '��ͨ����
                            strDate = VsfData.TextMatrix(lngStart + intBound, mlngDate)
                            strTime = VsfData.TextMatrix(lngStart + intBound, mlngTime)
                            If mblnDateAd Then strYear = VsfData.TextMatrix(lngStart + intBound, mlngYear)
                        End If
                        
                        '1\����
                        strField = "ID|ҳ��|�к�|�к�|��¼ID|����|��λ|ɾ��"
                        If mlngDate <> -1 Then
                            strKey = mintҳ�� & "," & lngStart + intBound & "," & mlngDate
                            strValue = strKey & "|" & mintҳ�� & "|" & lngStart + intBound & "|" & mlngDate & "|" & _
                                Val(VsfData.TextMatrix(lngStart + intBound, mlngRecord)) & "|" & strDate & "|" & strYear & "|" & IIf(VsfData.RowHidden(lngStart + intBound) = True, 1, 0)
                            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                        End If
                        '2\ʱ��
                        strKey = mintҳ�� & "," & lngStart + intBound & "," & mlngTime
                        strValue = strKey & "|" & mintҳ�� & "|" & lngStart + intBound & "|" & mlngTime & "|" & _
                            Val(VsfData.TextMatrix(lngStart + intBound, mlngRecord)) & "|" & strTime & "|" & _
                            VsfData.TextMatrix(lngStart + intBound, mlngDemo) & "|" & IIf(VsfData.RowHidden(lngStart + intBound) = True, 1, 0)
                        Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                        
                        If Not blnDate Then
                            strField = "ID|ҳ��|�к�|�к�|��¼ID|����|��λ|ɾ��"
                            If InStr(1, "," & mstrCatercorner & ",", "," & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ",") = 0 Then
                                strPart = GetActivePart(VsfData.COL, 0)
                            Else
                                strPart = GetActivePart(VsfData.COL, 0) & "/" & GetActivePart(VsfData.COL, 1)
                            End If
                            strKey = mintҳ�� & "," & lngStart + intBound & "," & VsfData.COL
                            strValue = strKey & "|" & mintҳ�� & "|" & lngStart + intBound & "|" & VsfData.COL & "|" & _
                                Val(VsfData.TextMatrix(lngStart + intBound, mlngRecord)) & "|" & strReturn & "|" & strPart & "|" & IIf(VsfData.RowHidden(lngStart + intBound) = True, 1, 0)
                            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                        End If
                    End If
                    '##########################################
                End If
            Next
            
            '���������н��и�ֵ
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
            '�Ը������¸�ֵ����ֻ����һ������ʱ����֪Ϊ�λ�����ַ�ASCII��Ϊ1�ķ��ţ�
            For intRow = 0 To intCount
                VsfData.TextMatrix(lngStart + intRow, VsfData.COL) = Replace(Replace(Replace(arrData(intRow), Chr(10), ""), Chr(13), ""), Chr(1), "")
                If blnGroup = True Then
                    '�����е����⴦��,�����ڲ���¼���Ĵ���϶�
                    '##########################################
                    If Val(VsfData.TextMatrix(lngStart + intRow, mlngDemo)) >= 1 Then
                        If (mblnEditAssistant = True Or blnDate = True) And Val(VsfData.TextMatrix(lngStart, mlngDemo)) = 1 Then
                            VsfData.TextMatrix(lngStart + intRow, mlngDemo) = intRow + 1
                        End If
                        intRowCount = 1
                        '��ȡ�÷��������е�����
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
                    '���÷��������������һ�в�ִ�и��²���
                    If intRow = intBound + intRowCount - 1 Then
                        '��������
                        strYear = ""
                        If Val(VsfData.TextMatrix(lngStart + intBound, mlngDemo)) > 0 Then
                            If CheckGroupDate(lngStart + intBound) = True Then
                                '�������޸ĲŽ�������̣�ȡ������¼��ʵ��ʱ��
                                If mblnDateAd Then
                                    strYear = Format(VsfData.TextMatrix(lngStart + intBound, mlngActTime), "YYYY")
                                    strDate = Format(VsfData.TextMatrix(lngStart + intBound, mlngActTime), "DD") & "/" & Format(VsfData.TextMatrix(lngStart + intBound, mlngActTime), "MM")
                                Else
                                    strDate = Mid(VsfData.TextMatrix(lngStart + intBound, mlngActTime), 1, 10)
                                End If
                                strTime = Mid(VsfData.TextMatrix(lngStart + intBound, mlngActTime), 12, 5)
                            Else
                                '����ʱ���������
                                strDate = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngDate)
                                strTime = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngTime)
                                If mblnDateAd Then strYear = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngYear)
                            End If
                        Else
                            '��ͨ����
                            strDate = VsfData.TextMatrix(lngStart + intBound, mlngDate)
                            strTime = VsfData.TextMatrix(lngStart + intBound, mlngTime)
                            If mblnDateAd Then strYear = VsfData.TextMatrix(lngStart + intBound, mlngYear)
                        End If
                        
                        '1\����
                        strField = "ID|ҳ��|�к�|�к�|��¼ID|����|��λ|ɾ��"
                        If mlngDate <> -1 Then
                            strKey = mintҳ�� & "," & lngStart + intBound & "," & mlngDate
                            strValue = strKey & "|" & mintҳ�� & "|" & lngStart + intBound & "|" & mlngDate & "|" & _
                                Val(VsfData.TextMatrix(lngStart + intBound, mlngRecord)) & "|" & strDate & "|" & strYear & "|" & IIf(VsfData.RowHidden(lngStart + intBound) = True, 1, 0)
                            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                        End If
                        '2\ʱ��
                        strKey = mintҳ�� & "," & lngStart + intBound & "," & mlngTime
                        strValue = strKey & "|" & mintҳ�� & "|" & lngStart + intBound & "|" & mlngTime & "|" & _
                            Val(VsfData.TextMatrix(lngStart + intBound, mlngRecord)) & "|" & strTime & "|" & _
                            VsfData.TextMatrix(lngStart + intBound, mlngDemo) & "|" & IIf(VsfData.RowHidden(lngStart + intBound) = True, 1, 0)
                        Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                        
                        If Not blnDate Then
                            strField = "ID|ҳ��|�к�|�к�|��¼ID|����|��λ|ɾ��"
                            If InStr(1, "," & mstrCatercorner & ",", "," & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ",") = 0 Then
                                strPart = GetActivePart(VsfData.COL, 0)
                            Else
                                strPart = GetActivePart(VsfData.COL, 0) & "/" & GetActivePart(VsfData.COL, 1)
                            End If
                            strKey = mintҳ�� & "," & lngStart + intBound & "," & VsfData.COL
                            strValue = strKey & "|" & mintҳ�� & "|" & lngStart + intBound & "|" & VsfData.COL & "|" & _
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
                '�����е����⴦��,�����ڲ���¼���Ĵ���϶�
                '##########################################
                '��������
                If (blnGroup And (mblnEditAssistant Or blnDate)) Then
                    '��ȡ������ʼ��
                    If lngStartGroup <> GetStartRow(lngStart + intRow) Then
                        intNULL = GetStartRow(lngStart + intRow)
                        'Ѱ�ҵ���ʼ��mlngDemo�϶�>0
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
                        '����������������д������,intNULL��¼���һ����Ϊ���е��к�
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
                        
                        '������д��������
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
                            '�������޸ĲŽ�������̣�ȡ������¼��ʵ��ʱ��
                            If mblnDateAd Then
                                strYear = Format(VsfData.TextMatrix(lngStart + intRow, mlngActTime), "YYYY")
                                strDate = Format(VsfData.TextMatrix(lngStart + intRow, mlngActTime), "DD") & "/" & Format(VsfData.TextMatrix(lngStart + intRow, mlngActTime), "MM")
                            Else
                                strDate = Mid(VsfData.TextMatrix(lngStart + intRow, mlngActTime), 1, 10)
                            End If
                            strTime = Mid(VsfData.TextMatrix(lngStart + intRow, mlngActTime), 12, 5)
                        Else
                            '����ʱ���������
                            strDate = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngDate)
                            strTime = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngTime)
                            If mblnDateAd Then strYear = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngYear)
                        End If
                    
                        '������ʼ�е���������ʱ���������÷����
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
                        
                        '1\����
                        strField = "ID|ҳ��|�к�|�к�|��¼ID|����|��λ|ɾ��"
                        If mlngDate <> -1 Then
                            strKey = mintҳ�� & "," & lngStart + intRow & "," & mlngDate
                            strValue = strKey & "|" & mintҳ�� & "|" & lngStart + intRow & "|" & mlngDate & "|" & _
                                Val(VsfData.TextMatrix(lngStart + intRow, mlngRecord)) & "|" & strDate & "|" & strYear & "|" & IIf(VsfData.RowHidden(lngStart + intRow) = True, 1, 0)
                            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                        End If
                        '2\ʱ��
                        strKey = mintҳ�� & "," & lngStart + intRow & "," & mlngTime
                        strValue = strKey & "|" & mintҳ�� & "|" & lngStart + intRow & "|" & mlngTime & "|" & _
                            Val(VsfData.TextMatrix(lngStart + intRow, mlngRecord)) & "|" & strTime & "|" & _
                            VsfData.TextMatrix(lngStart + intRow, mlngDemo) & "|" & IIf(VsfData.RowHidden(lngStart + intRow) = True, 1, 0)
                        Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                        
                        If Not blnDate Then
                            strField = "ID|ҳ��|�к�|�к�|��¼ID|����|��λ|ɾ��"
                            If InStr(1, "," & mstrCatercorner & ",", "," & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ",") = 0 Then
                                strPart = GetActivePart(VsfData.COL, 0)
                            Else
                                strPart = GetActivePart(VsfData.COL, 0) & "/" & GetActivePart(VsfData.COL, 1)
                            End If
                            strKey = mintҳ�� & "," & lngStart + intRow & "," & VsfData.COL
                            strValue = strKey & "|" & mintҳ�� & "|" & lngStart + intRow & "|" & VsfData.COL & "|" & _
                                Val(VsfData.TextMatrix(lngStart + intRow, mlngRecord)) & "|" & "" & "|" & strPart & "|1"
                            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                        End If
                    End If
                End If
                '##########################################
            Next
            '�޸ķ���Ǵ��ı���������ʱ����Ҫ��ȡ�������ݴ��ı���������Ϣ��������֯�ı���ʾ
            '����3�����ݣ��ڶ�2����3�У��޸�Ϊ1�У���3������Ӧ�ý�������ʾ�ڵ�2������(�ڶ����ʱֻ��1��)
            If blnReseGroupAssistant = True Then Call GetGroupAssistant(strAssistantCols, varAssistant)
            lngMutilRows = Val(Split(VsfData.TextMatrix(lngStart, mlngRowCount), "|")(0))
            '����������������д������,intNULL��¼���һ����Ϊ���е��к�
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
            '������д�����
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
            Else '�������Ա��������ɾ��ʱ��������к�
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
            '���¼�¼�����Ķ���Ϣ
            If blnReseGroupAssistant = True Then Call CellMap_UpdateAssistant(lngStart)
        End If
        
        '��ȡ������ʼ����������Ϣ
        If blnTrue = True Then 'blnTrueΪ��˵��ѡ����Ƿ����е���ʼ�У������Ǵ��ı���
            strReturn = ""
            intCount = Val(Split(VsfData.TextMatrix(lngStart, mlngRowCount), "|")(0))
            For intRow = 0 To intCount - 1
                strReturn = strReturn & Replace(Replace(Replace(VsfData.TextMatrix(VsfData.ROW + intRow, VsfData.COL), Chr(10), ""), Chr(13), ""), Chr(1), "")
            Next intRow
        End If
        
        If mstrData <> strReturn Or blnTrue = True Then
            If strText <> mstrData Then mblnChange = True
            'ͬ������������ʱ���е�����
            If Val(VsfData.TextMatrix(lngStart, mlngCollectType)) >= 0 Then
                strYear = ""
                If Val(VsfData.TextMatrix(lngStart, mlngDemo)) > 0 Then
                    If CheckGroupDate(lngStart) = True Then
                        '�������޸ĲŽ�������̣�ȡ������¼��ʵ��ʱ��
                        If mblnDateAd Then
                            strYear = Format(VsfData.TextMatrix(lngStart, mlngActTime), "YYYY")
                            strDate = Format(VsfData.TextMatrix(lngStart, mlngActTime), "DD") & "/" & Format(VsfData.TextMatrix(lngStart, mlngActTime), "MM")
                        Else
                            strDate = Mid(VsfData.TextMatrix(lngStart, mlngActTime), 1, 10)
                        End If
                        strTime = Mid(VsfData.TextMatrix(lngStart, mlngActTime), 12, 5)
                    Else
                        '����ʱ���������
                        strDate = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngDate)
                        strTime = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngTime)
                        If mblnDateAd Then strYear = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngYear)
                    End If
                Else
                    '��ͨ����
                    strDate = VsfData.TextMatrix(lngStart, mlngDate)
                    strTime = VsfData.TextMatrix(lngStart, mlngTime)
                    If mblnDateAd Then strYear = VsfData.TextMatrix(lngStart, mlngYear)
                End If
                
                '1\����
                strField = "ID|ҳ��|�к�|�к�|��¼ID|����|��λ|ɾ��"
                If mlngDate <> -1 Then
                    strKey = mintҳ�� & "," & lngStart & "," & mlngDate
                    strValue = strKey & "|" & mintҳ�� & "|" & lngStart & "|" & mlngDate & "|" & _
                        Val(VsfData.TextMatrix(lngStart, mlngRecord)) & "|" & strDate & "|" & strYear & "|0"
                    Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                End If
                '2\ʱ��
                strKey = mintҳ�� & "," & lngStart & "," & mlngTime
                strValue = strKey & "|" & mintҳ�� & "|" & lngStart & "|" & mlngTime & "|" & _
                    Val(VsfData.TextMatrix(lngStart, mlngRecord)) & "|" & strTime & "|" & _
                    VsfData.TextMatrix(lngStart, mlngDemo) & "|0"
                Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
            Else
                strField = "ID|ҳ��|�к�|�к�|��¼ID|����|����|ɾ��"
                strKey = mintҳ�� & "," & lngStart & "," & mlngDate
                strValue = strKey & "|" & mintҳ�� & "|" & lngStart & "|" & mlngDate & "|" & Val(VsfData.TextMatrix(lngStart, mlngRecord)) & "|" & _
                        VsfData.TextMatrix(lngStart, mlngCollectText) & ";" & VsfData.TextMatrix(lngStart, mlngCollectType) & ";" & _
                        VsfData.TextMatrix(lngStart, mlngCollectStyle) & ";" & VsfData.TextMatrix(lngStart, mlngCollectDay) & ";" & _
                    VsfData.TextMatrix(lngStart, mlngCollectStart) & ";" & VsfData.TextMatrix(lngStart, mlngCollectEnd) & "|1|0"
                Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                
            End If
            
            If (Not blnGroup Or blnTrue) And Not blnDate Then
                '��¼�û��޸Ĺ��ĵ�Ԫ��
                If InStr(1, "," & mstrCatercorner & ",", "," & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ",") = 0 Then
                    strPart = GetActivePart(VsfData.COL, 0)
                Else
                    strPart = GetActivePart(VsfData.COL, 0) & "/" & GetActivePart(VsfData.COL, 1)
                End If
                
                strField = "ID|ҳ��|�к�|�к�|��¼ID|����|��λ|ɾ��"
                strKey = mintҳ�� & "," & lngStart & "," & VsfData.COL
                strValue = strKey & "|" & mintҳ�� & "|" & lngStart & "|" & VsfData.COL & "|" & _
                    Val(VsfData.TextMatrix(lngStart, mlngRecord)) & "|" & strReturn & "|" & strPart & "|" & IIf(strReturn = "", "1", "0")
                Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
            End If
        End If
        
        '52953,������,2012-08-24,��������Ϊ0ҲҪ��ʾ,�������������кϲ�����������:60792
        '����������������������ݺϲ�
        If Val(VsfData.TextMatrix(lngStart, mlngCollectType)) < 0 Then
            VsfData.TextMatrix(lngStart, VsfData.COL) = FormatValue(VsfData.TextMatrix(lngStart, VsfData.COL))
            If Trim(VsfData.TextMatrix(lngStart, VsfData.COL)) <> "" Then
                '66085:������,2012-09-26,�������ڻ����кϲ�,��ԭ����������+�ո�ͬһ�ĳ����к�����chr(13)
                '������ӿո���п�������������ʾ����ȫ(��Ҫ����Ҷ���)
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
    '������������ʱ�����հ������������һ��
    If Left(Trim(strRows), 1) = "," Then strRows = Mid(Trim(strRows), 2)
    If Right(strRows, 1) = "," Then strRows = Mid(strRows, 1, Len(strRows) - 1)
    strRows = Replace("," & strRows & ",", "," & lngDemoRow & ",", "") '����ɾ��Ҫ׷�ӵ���
    For intRow = VsfData.Rows - 1 To VsfData.FixedRows Step -1
        If InStr(1, strRows, "," & intRow & ",") <> 0 Then
            Call CellMap_Update(intRow, -1)
            VsfData.TextMatrix(intRow, mlngDemo) = ""
        End If
    Next intRow
    
    For intRow = VsfData.Rows - 1 To VsfData.FixedRows Step -1
        If InStr(1, strRows, "," & intRow & ",") <> 0 Then
            '��ո���������Ϣ
            For intBound = 0 To VsfData.Cols - 1
                VsfData.TextMatrix(intRow, intBound) = ""
            Next intBound
            VsfData.RowPosition(intRow) = VsfData.Rows - 1
        End If
    Next intRow
    
    '--58414,������,2013-04-10,��ӻ�ʿ��ǩ������ʾģʽ
    If Val(VsfData.TextMatrix(lngStart, mlngDemo)) = 1 Then
        '��¼������ʼ�е���������
        intGroupFirstRows = Val(Split(VsfData.TextMatrix(lngStart, mlngRowCount), "|")(0))
        intBound = lngStart + intGroupFirstRows - 1
        Call SingerShowType(VsfData, lngStart, intBound)
        For intCount = lngStart + intGroupFirstRows To VsfData.Rows - 1
            '������ÿһ�������������ݿ���ռ�ö��У�ֻ���ڵ�һ�б����˷�����������������������=ÿһ������������+ÿһ���������е�����
            If intCount > intBound Then
                If Val(VsfData.TextMatrix(intCount, mlngDemo)) <= 1 Then Exit For      '����������·�����˳�
                intBound = Val(Split(VsfData.TextMatrix(intCount, mlngRowCount), "|")(0)) + intCount - 1
                '--58414,������,2013-04-10,��ӻ�ʿ��ǩ������ʾģʽ
                Call SingerShowType(VsfData, intCount, intBound)
            End If
        Next
    Else
        intGroupFirstRows = Val(Split(VsfData.TextMatrix(lngStart, mlngRowCount), "|")(0))
        '--58414,������,2013-04-10,��ӻ�ʿ��ǩ������ʾģʽ
        Call SingerShowType(VsfData, lngStart, lngStart + intGroupFirstRows - 1)
    End If
    
    'Call OutputRsData(mrsCellMap)

    '������֯������������
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
        If VsfData.COL < mlngNoEditor - 1 Then       '�����¼���϶��л�ʿǩ����
            VsfData.COL = VsfData.COL + 1
            If VsfData.ColWidth(VsfData.COL) = 0 Or VsfData.ColHidden(VsfData.COL) Or mintType = -1 Then GoTo toMoveNextCol
        Else
toMoveNextRow:
            '������һ��
            If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") <> 0 Then
                intRow = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0))
                intRow = intRow - Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(1)) + 1
            Else
                intRow = 1
            End If
            mblnShow = False
            If VsfData.ROW + intRow < VsfData.Rows Then
                'ֻ����׷�ӵ�ģʽ�£������mstrGroupRow
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
        If VsfData.COL > mlngDate Then      '�����¼���϶��л�ʿǩ����
            VsfData.COL = VsfData.COL - 1
            If VsfData.ColWidth(VsfData.COL) = 0 Or VsfData.ColHidden(VsfData.COL) Or mintType = -1 Then GoTo toMovePrevCol
        Else
toMovePrevRow:
'            '������һ��
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
'--���ܣ�������������ʼ��ʱ��ͱ���ʱ���Ƿ����
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
'���ܣ����¼�¼�����Ķ���Ϣ
    Dim strDate As String, strTime As String, strYear As String
    Dim strKey As String, strField As String, strValue As String, strPart As String
    Dim lngCol As Long, lngRow As Long, lngRowCount As Long, strReturn As String
    
    On Error GoTo ErrHand
    
    If VsfData.TextMatrix(lngStartRow, mlngRowCount) = "" Then VsfData.TextMatrix(lngStartRow, mlngRowCount) = "1|1"
    lngRowCount = Val(Split(VsfData.TextMatrix(lngStartRow, mlngRowCount), "|")(0))
    
    strYear = ""
    If Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) > 0 Then
        If CheckGroupDate(lngStartRow) = True Then
            '�������޸ĲŽ�������̣�ȡ������¼��ʵ��ʱ��
            If mblnDateAd Then
                strYear = Format(VsfData.TextMatrix(lngStartRow, mlngActTime), "YYYY")
                strDate = Format(VsfData.TextMatrix(lngStartRow, mlngActTime), "DD") & "/" & Format(VsfData.TextMatrix(lngStartRow, mlngActTime), "MM")
            Else
                strDate = Mid(VsfData.TextMatrix(lngStartRow, mlngActTime), 1, 10)
            End If
            strTime = Mid(VsfData.TextMatrix(lngStartRow, mlngActTime), 12, 5)
        Else
            '����ʱ���������
            strDate = VsfData.TextMatrix(lngStartRow - Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) + 1, mlngDate)
            strTime = VsfData.TextMatrix(lngStartRow - Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) + 1, mlngTime)
            If mblnDateAd Then strYear = VsfData.TextMatrix(lngStartRow - Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) + 1, mlngYear)
        End If
    Else
        '��ͨ����
        strDate = VsfData.TextMatrix(lngStartRow, mlngDate)
        strTime = VsfData.TextMatrix(lngStartRow, mlngTime)
        If mblnDateAd Then strYear = VsfData.TextMatrix(lngStartRow, mlngYear)
    End If
    
    '1\����
    strField = "ID|ҳ��|�к�|�к�|��¼ID|����|��λ|ɾ��"
    If mlngDate <> -1 Then
        strKey = mintҳ�� & "," & lngStartRow & "," & mlngDate
        strValue = strKey & "|" & mintҳ�� & "|" & lngStartRow & "|" & mlngDate & "|" & _
            Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "|" & strDate & "|" & strYear & "|0"
        Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
    End If
    '2\ʱ��
    strKey = mintҳ�� & "," & lngStartRow & "," & mlngTime
    strValue = strKey & "|" & mintҳ�� & "|" & lngStartRow & "|" & mlngTime & "|" & _
        Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "|" & strTime & "|" & _
        VsfData.TextMatrix(lngStartRow, mlngDemo) & "|0"
    Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
    
    For lngCol = mlngTime + 1 To mlngNoEditor - 1
        If ISEditAssistant(lngCol) Then
            strReturn = ""
            For lngRow = 0 To lngRowCount - 1
                strReturn = strReturn & Replace(Replace(Replace(VsfData.TextMatrix(lngStartRow + lngRow, lngCol), Chr(10), ""), Chr(13), ""), Chr(1), "")
            Next lngRow
            '��¼�û��޸Ĺ��ĵ�Ԫ��
            If InStr(1, "," & mstrCatercorner & ",", "," & lngCol - (cHideCols + VsfData.FixedCols - 1) & ",") = 0 Then
                strPart = GetActivePart(lngCol, 0)
            Else
                strPart = GetActivePart(lngCol, 0) & "/" & GetActivePart(lngCol, 1)
            End If
            
            strField = "ID|ҳ��|�к�|�к�|��¼ID|����|��λ|ɾ��"
            strKey = mintҳ�� & "," & lngStartRow & "," & lngCol
            strValue = strKey & "|" & mintҳ�� & "|" & lngStartRow & "|" & lngCol & "|" & _
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
'���ܣ��������д��ı�����ÿһ�е�����
'˵���������޸ķǴ��Ķλ�����ʱ���еķ�������ʱ�ŵ���(�ȵ���GetGroupAssistant�����ڵ��ô˷���)
    Dim lngCol As Long, lngRow As Long, lngStartRow As Long, varCol
    Dim lngOldRow As Long, lngOldCol As Long, intType As Integer, blnTrue As Boolean
    Dim strText As String, blnOldNoMove As Boolean, blnOldNext As Boolean
    
    lngOldRow = VsfData.ROW
    lngOldCol = VsfData.COL
    intType = mintType
    blnOldNoMove = blnNoMove
    blnOldNext = blnNext
    
    '��ȡ�༭��ǰ�е���ʼ����
    If VsfData.TextMatrix(VsfData.ROW, mlngRowCount) <> "1|1" Then
        lngRow = GetStartRow(VsfData.ROW)
    Else
        lngRow = VsfData.ROW
    End If
    '��ȡ�������ݵĵ�һ��
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
    
    '�ָ���
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
'���ܣ���ȡ���ı�����Ϣ
'˵���������޸ķǴ��Ķλ�����ʱ���еķ�������ʱ�ŵ���
    Dim lngRow As Long, lngCol As Long, lngOrder As Long, intGroupFirstRows As Integer, lngCount As Long
    Dim lngStartRow As Long
    Dim strText As String
    
    strAssistantCols = ""
    varAssistantText = Array()
    
    '��ȡ�༭��ǰ�е���ʼ����
    If VsfData.TextMatrix(VsfData.ROW, mlngRowCount) <> "1|1" Then
        lngRow = GetStartRow(VsfData.ROW)
    Else
        lngRow = VsfData.ROW
    End If
    '��ȡ�������ݵĵ�һ��
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
        'Ѱ�Ҵ��ı���
        mrsSelItems.Filter = "��=" & lngCol - cHideCols
        If mrsSelItems.RecordCount > 0 Then
            lngOrder = Val(mrsSelItems!��Ŀ���)
            mrsItems.Filter = "��Ŀ���=" & lngOrder
            If mrsItems.RecordCount = 0 Then
                mrsItems.Filter = 0
                GoTo ErrNext
            End If
            mblnEditAssistant = (mrsItems!��Ŀ���� = 1 And mrsItems!��Ŀ���� > 100) And Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) <= 1
            If Not mblnEditAssistant Then GoTo ErrNext
                
            If InStr(1, VsfData.TextMatrix(lngStartRow, mlngRowCount), "|") = 0 Then VsfData.TextMatrix(lngStartRow, mlngRowCount) = "1|1"
            intGroupFirstRows = Val(Split(VsfData.TextMatrix(lngStartRow, mlngRowCount), "|")(0))
            'Ϊ������ʱ��ѡ��������ʼ�У��༭������ʾ���д��ı���
            strText = ""
            If Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) = 1 Then
                For lngRow = 0 To intGroupFirstRows - 1
                    strText = strText & IIf(strText = "", "", vbCrLf) & Replace(Replace(Replace(VsfData.TextMatrix(lngRow + lngStartRow, lngCol), Chr(13), ""), Chr(10), ""), Chr(1), "")
                Next lngRow
                lngCount = lngStartRow + intGroupFirstRows - 1
                For lngRow = lngStartRow + intGroupFirstRows To VsfData.Rows - 1
                    If VsfData.RowHidden(lngRow) = False Then
                        '������ÿһ�������������ݿ���ռ�ö��У�ֻ���ڵ�һ�б����˷�����������������������=ÿһ������������+ÿһ���������е�����
                        If lngRow > lngCount Then
                            If Val(VsfData.TextMatrix(lngRow, mlngDemo)) <= 1 Then Exit For  '����������·�����˳�
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
'���ܣ����÷����е��кźͼ�¼����Ϣ
'���޸ķ���������ʱ������������ı����ı����ݲ�Ϊ��ͨ��GetGroupAssistant��ReSetGroupAssistant������ã����û������ô˺����������
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
    'ȷ��������ʼ��
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
    '������֯�������
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

    '����ʼ�п�ʼ����������ݼ�¼��
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
                '�������޸ĲŽ�������̣�ȡ������¼��ʵ��ʱ��
                If mblnDateAd Then
                    strYear = Format(VsfData.TextMatrix(lngRow, mlngActTime), "YYYY")
                    strDate = Format(VsfData.TextMatrix(lngRow, mlngActTime), "DD") & "/" & Format(VsfData.TextMatrix(lngRow, mlngActTime), "MM")
                Else
                    strDate = Mid(VsfData.TextMatrix(lngRow, mlngActTime), 1, 10)
                End If
                strTime = Mid(VsfData.TextMatrix(lngRow, mlngActTime), 12, 5)
            Else
                '����ʱ���������
                strDate = VsfData.TextMatrix(lngRow - Val(VsfData.TextMatrix(lngRow, mlngDemo)) + 1, mlngDate)
                strTime = VsfData.TextMatrix(lngRow - Val(VsfData.TextMatrix(lngRow, mlngDemo)) + 1, mlngTime)
                If mblnDateAd Then strYear = VsfData.TextMatrix(lngRow - Val(VsfData.TextMatrix(lngRow, mlngDemo)) + 1, mlngYear)
            End If
            
            '1\����
            strField = "ID|ҳ��|�к�|�к�|��¼ID|����|��λ|ɾ��"
            If mlngDate <> -1 Then
                strKey = mintҳ�� & "," & lngRow & "," & mlngDate
                strValue = strKey & "|" & mintҳ�� & "|" & lngRow & "|" & mlngDate & "|" & _
                    Val(VsfData.TextMatrix(lngRow, mlngRecord)) & "|" & strDate & "|" & strYear & "|" & IIf(VsfData.RowHidden(lngRow) = True, 1, 0)
                Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
            End If
            '2\ʱ��
            strKey = mintҳ�� & "," & lngRow & "," & mlngTime
            strValue = strKey & "|" & mintҳ�� & "|" & lngRow & "|" & mlngTime & "|" & _
                Val(VsfData.TextMatrix(lngRow, mlngRecord)) & "|" & strTime & "|" & _
                VsfData.TextMatrix(lngRow, mlngDemo) & "|" & IIf(VsfData.RowHidden(lngRow) = True, 1, 0)
            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
        End If
    Next lngRow
End Sub

Private Function ISEditAssistant(ByVal lngCol As Long) As Boolean
'�Ƿ�༭���Ǵ��ı���Ŀ
    Dim blnTrue As Boolean, lngOrder As Long
    
    mrsSelItems.Filter = "��=" & lngCol - cHideCols
    If mrsSelItems.RecordCount > 0 Then
        lngOrder = Val(mrsSelItems!��Ŀ���)
        mrsItems.Filter = "��Ŀ���=" & lngOrder
        If mrsItems.RecordCount = 0 Then
            mrsItems.Filter = 0
            Exit Function
        End If
        blnTrue = (mrsItems!��Ŀ���� = 1 And mrsItems!��Ŀ���� > 100)
    End If
    ISEditAssistant = blnTrue
End Function

Private Sub AppendGroup(ByVal lngStartRow As Long)
    Dim lngDemo As Long, lngStart As Long, lngRows As Long
    Dim blnGroup As Boolean
    '׷�ӷ�����(ֻ���ڵ������к�׷�ӷ�����)
    If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then VsfData.TextMatrix(VsfData.ROW, mlngRowCount) = "1|1"
    lngRows = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0))
    '�ȼ�鵱ǰ���Ƿ�Ϊ������
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
    '��ȡ����������ʼ��,������ҳ�򷵻�0
    '�����ҳδ��ʾȫ,��˵��������ҳ,Ҳ����0
    '���������������������в�������
    If VsfData.TextMatrix(lngRow, mlngRowCount) = "" Then VsfData.TextMatrix(lngRow, mlngRowCount) = "1|1"
    lngRows = Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(0))    '������
    If InStr(1, VsfData.TextMatrix(lngRow, mlngRowCount), "|") = 0 Then
        lngCurRows = 1
    Else
        lngCurRows = Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(1)) '��ǰ��
    End If
    If lngCurRows = 1 Then
        GetStartRow = lngRow
        Exit Function
    End If
    
    'Ѱ����ʼ��
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
    Dim lngStart As Long    '��ʼ��
    Dim lngRecordId As Long
    Dim strReturn As String
    Dim blnAdjust As Boolean
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    '���ص�һ�е�����
    '������ֱ��ȡ������ʱ��������ҳ��ʾȫ��ƴ�ӣ�����ӿ��ж�ȡ
    
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
    If lngRecordId <> 0 And (lngStart = 0 Or lngStart + lngCount > VsfData.Rows) Then   'ҳ��Ч��=�̶�������+��ͷ
        '�����ݿ�����ȡ
        Call SQLCombination(lngRecordId)
        gstrSQL = mstrSQL
        Call SQLDIY(gstrSQL)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��������", mlng�ļ�ID, mlng����ID, mlng��ҳID, mintӤ��, mintҳ��, lngRecordId)
        strReturn = NVL(rsTemp.Fields(lngCol).Value)
        If lngStart = 0 Then lngStart = 3       '���δ�ҵ���ʼ�����趨Ϊ��1��
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
    
'    'У���и�(�п���ʵ������ռ5�ж���ǰҳ��ֻ��ʾ��3��,����3����ʾ�������Բ�ȫ,���Ի�����ԭ�����и���ʾ����,���´�������)
'    If blnAdjust Then
'        If lngStart = 3 Then
'            lngCurRow = Val(Split(VsfData.TextMatrix(lngStart, mlngRowCount), "|")(1))
'            lngCount = lngCount - lngCurRow + 1
'        Else
'            lngCount = mlngPageRows +mlngOverrunRows + VsfData.FixedRows - lngStart
'        End If
'    End If
    'ȡ�и�
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
    Dim strFormat As String, strText As String, strValue As String  '��ʽ��,���ݴ�,��ֵ��
    Dim strOrders As String, strTypes As String, strBounds As String, strLen As String, strName As String, strState As String
    Const txtHeight = 300
    On Error GoTo ErrHand
    
    '�����ļ��������ģ����Ҫ����:
    '1��һ�а�һ����Ŀ�Ĳ��ù�
    '2��һ�а�������Ŀ�ģ�Ѫѹ����ɶԣ�Ҫô����¼�룬Ҫô����ѡ�񣬲���������֣�Ҳ��������ֵ�ѡ����ѡ
    '3��һ�а󶨶����Ŀ�ģ�ֻ����¼����Ŀ
    '���������������ƣ�ֻȡ��һ����Ŀ�����ʼ���
    
    '����Ǳ��洦�����������´���
    If intCol = -1 Then intCol = VsfData.COL
    If blnAnalyse Then
        strText = Replace(Replace(Replace(strCellData, Chr(10), ""), Chr(13), ""), Chr(1), "")
    Else
        'ȡ��ǰ��Ԫ�������
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
    
    'ȡ��ǰ�еİ���Ŀ
    intPos = 1
    mrsSelItems.Filter = "��=" & intCol - cHideCols
    Do While Not mrsSelItems.EOF
        lngOrder = mrsSelItems!��Ŀ���
        If lngOrder = 0 Then
            strLen = 0
            strValue = strText
            Exit Do
        End If
        
        '��Ŀ��ʾ:2��ѡ;3-��ѡ;4-����;5-ѡ��
        '��Ŀֵ��:��Ŀ��ʾΪ0-��ʾ��Сֵ;���ֵ;��Ŀ��ʾΪ2,3-��ʾ��ĿA;��ĿB,ǰ�й��ı�ʾȱʡ��
        strFormat = NVL(mrsSelItems!��ʽ)
        strOrders = strOrders & "," & lngOrder
        If lngOrder <> 0 Then
            mrsItems.Filter = "��Ŀ���=" & lngOrder
            strName = strName & "," & GetActivePart(intCol, intIndex) & UCase(mrsItems!��Ŀ����)
            strLen = strLen & "," & mrsItems!��Ŀ���� & ";" & NVL(mrsItems!��ĿС��)
            strState = strState & "," & mrsItems!��Ŀ����
            strTypes = strTypes & "," & mrsItems!��Ŀ��ʾ
            strBounds = strBounds & "," & mrsItems!��Ŀֵ��
            strValue = strValue & "'" & SubstrVal(strText, strFormat, GetActivePart(intCol, intIndex) & UCase(mrsItems!��Ŀ����), intPos)
            
            Select Case mrsItems!��Ŀ��ʾ
            Case 0  '�ı�¼����
                If mrsSelItems.RecordCount = 2 Then
                    If InStr(1, strState & ",", ",1,") = 0 Then
                        mintType = 4
                    Else
                        mintType = 6
                    End If
                ElseIf mrsSelItems.RecordCount > 2 Then
                    mintType = 6
                End If
            Case 2  '��ѡ
                If mrsSelItems.RecordCount = 1 Then
                    mintType = 1
                ElseIf mrsSelItems.RecordCount = 2 Then
                    mintType = 7
                End If
            Case 3  '��ѡ
                mintType = 2
            Case 4  '����
                If mrsSelItems.RecordCount = 2 Then
                    mintType = 4
                ElseIf mrsSelItems.RecordCount > 2 Then
                    mintType = 6
                End If
            Case 5  'ѡ��
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
    
    '���4����У��,�����ͷ�ı�����/����Ϊ6
    If mintType = 4 Then
        If Not IsDiagonal(intCol) Then
            mintType = 6
        End If
    End If
    
    '�жϵ�ǰ�е�����
    'mintType:0=�ı���¼��;1=��ѡ;2=��ѡ;3=ѡ��;4-Ѫѹ��һ�а���������Ŀ,���ʽ����Ѫѹ��������Ŀ;5=һ�а���������Ŀ�Ҿ���ѡ����Ŀ;
    '6=һ�а�2����������Ŀ,�ֹ�¼��,7=һ�а���������ѡ��Ŀ
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
                txtInput.MaxLength = Val(Split(strLen, ";")(0)) + IIf(Val(Split(strLen, ";")(1)) = 0, 0, 1) 'С��λ��Ҫ����С����
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
            .Width = .Width - (180 + IIf(mblnBlowup, 180 * 1 / 3, 0)) / 2 + 3 '����9��ʱ��ȥ90,����Խ��۳��ı߾�ԽС,�Ա�֤�ı��������ʵ��һ��
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
        
        '��������ڻ�ʱ���У��趨�̶�ֵ
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
        '56439:������,2012-11-30,��ѡ��Ŀ���δ����ȱʡ�룬Ĭ�϶�λ�����ѡ����ǰ�ķ�ʽ�Ƕ�λ��
        'ʵ�������������Щ��Ŀ����Ҫ¼�룬�ͻ�Ҫ����Ա�ֹ�ѡ�����ѡ����ڲ����Ϻ��鷳��
        '��������
        lstSelect(mintType - 1).Clear
        If mintType = 1 Then lstSelect(mintType - 1).AddItem "���ѡ��"
        If strBounds = "" Then strBounds = ";"
        arrData = Split(strBounds, ";")
        j = UBound(arrData)
        For i = 0 To j
            If arrData(i) <> "" Then
                If Mid(arrData(i), 1, 1) = "��" Then
                    lstSelect(mintType - 1).AddItem lstSelect(mintType - 1).NewIndex + 1 & "-" & Mid(arrData(i), 2)
                    If strText = "" Then lstSelect(mintType - 1).ListIndex = lstSelect(mintType - 1).NewIndex
                Else
                    lstSelect(mintType - 1).AddItem lstSelect(mintType - 1).NewIndex + 1 & "-" & arrData(i)
                End If
            End If
        Next
        
        '��ѡ����¼�����ݵ������
        If strValue <> "" Then
            strValue = Replace(strValue, vbCrLf, "")
            txtLst.Text = strValue
            PicLst.Tag = "1"
            j = lstSelect(mintType - 1).ListCount - 1
            For i = 0 To j
                '��ѡ�ĵ�һ����Ŀ�����ѡ����Ҫ���������ѡ��Ŀ��ֱ�ӽ���
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
        
        '�ؼ���ʾ
        '51134,������,2012-07-11,��ѡ�ṩ�ı�¼��
        PicLst.FontName = VsfData.FontName
        PicLst.FontSize = VsfData.FontSize
        If mintType = 1 Then
        
            With PicLst
                .Left = CellRect.Left
                .Top = CellRect.Top
                .Width = LenB(StrConv(lstSelect(mintType - 1).List(lstSelect(mintType - 1).ListCount \ 2), vbFromUnicode)) * 120 + 500    '���м���ĳ���Ϊ����
                If .Width < CellRect.Right Then .Width = CellRect.Right
            End With
            
            With lbllst(0)
                .Left = 20
                .Top = 20
                If .Width > PicLst.Width Then
                    PicLst.Width = .Width + PicLst.TextWidth("��")
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
                    If PicLst.TextHeight("��") * (UBound(arrData) + 1) + PicLst.TextHeight("��") \ 3 < VsfData.CellHeight + 20 Then
                        .Height = VsfData.CellHeight + 20
                    Else
                        .Height = PicLst.TextHeight("��") * (UBound(arrData) + 1) + PicLst.TextHeight("��") \ 3
                    End If
                Else
                    .Height = VsfData.CellHeight + 20
                End If
                .FontName = VsfData.FontName
                .FontSize = VsfData.FontSize
                .Tag = VsfData.CellHeight + 20 '��С�߶�
                If strLen <> "" Then
                    .MaxLength = Val(Split(strLen, ";")(0)) + IIf(Val(Split(strLen, ";")(1)) = 0, 0, 1) 'С��λ��Ҫ����С����
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
            '56047:������,2012-11-22,�޸�PicLst������
            With PicLst
                'list�ؼ��ĸ߶���С��240������������߶ȱ����������ʸ߶ȼ���=.ListCount * PicLst.TextHeight("��") + PicLst.TextHeight("��") \ 3
                .Height = lbllst(1).Top + lbllst(1).Height + 20 + lstSelect(mintType - 1).ListCount * PicLst.TextHeight("��") + PicLst.TextHeight("��") \ 3
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
                .Height = IIf(PicLst.Height - .Top < PicLst.TextHeight("��") + PicLst.TextHeight("��") \ 3, PicLst.TextHeight("��") + PicLst.TextHeight("��") \ 3, PicLst.Height - .Top)
                .Tag = lngOrder
                .Visible = True
                
                If .Top + .Height <> PicLst.Height Then
                    PicLst.Height = .Top + .Height
                End If
            End With
        Else
            '56047:������,2012-11-22,�޸�lstSelect������
            '��ʾ
            With lstSelect(mintType - 1)
                .Left = CellRect.Left
                .Top = CellRect.Top
                .FontName = VsfData.FontName
                .FontSize = VsfData.FontSize
                .Height = .ListCount * (PicLst.TextHeight("��")) + PicLst.TextHeight("��") \ 3
                If .Height < CellRect.Bottom Then .Height = CellRect.Bottom
                .Width = LenB(StrConv(.List(.ListCount \ 2), vbFromUnicode)) * 120 + 500    '���м���ĳ���Ϊ����
                If .Width < CellRect.Right Then .Width = CellRect.Right
                If .Height + .Top + picMain.Top > ScaleHeight Then
                    If ScaleHeight - picMain.Top - .Height < 0 Then
                        .Top = 10
                        .Height = IIf(ScaleHeight - picMain.Top - 10 < PicLst.TextHeight("��") + PicLst.TextHeight("��") \ 3, PicLst.TextHeight("��") + PicLst.TextHeight("��") \ 3, ScaleHeight - picMain.Top - 10)
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
            If strLen <> "" Then txtUpInput.MaxLength = Val(Split(Split(strLen, ",")(0), ";")(0)) + IIf(Val(Split(Split(strLen, ",")(0), ";")(1)) = 0, 0, 1) 'С��λ��Ҫ����С����
            If strLen <> "" Then txtDnInput.MaxLength = Val(Split(Split(strLen, ",")(1), ";")(0)) + IIf(Val(Split(Split(strLen, ",")(1), ";")(1)) = 0, 0, 1) 'С��λ��Ҫ����С����
        End If
    Case 6
        '��ɾ����ǰ�Ŀؼ�
        j = txt.Count - 1
        For i = 1 To j
            Unload lbl(i)
            Unload txt(i)
        Next
        '�趨����
        With picMutilInput
            .Left = CellRect.Left
            .Top = CellRect.Top
            .Width = IIf(CellRect.Right < 1600, 1600, CellRect.Right)
        End With
        '��ȱʡ�ؼ���ֵ
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
        txt(0).MaxLength = Val(Split(Split(strLen, ",")(0), ";")(0)) + IIf(Val(Split(Split(strLen, ",")(0), ";")(1)) = 0, 0, 1)  'С��λ��Ҫ����С����
        txt(0).Text = arrValue(0)
        If Not mblnBlowup Then
            txt(0).Height = 225
        End If
        
        '���ؿؼ�
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
                    .MaxLength = Val(Split(Split(strLen, ",")(i), ";")(0)) + IIf(Val(Split(Split(strLen, ",")(i), ";")(1)) = 0, 0, 1) 'С��λ��Ҫ����С����
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
            If Mid(arrData(i), 1, 1) = "��" Then
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
            If Mid(arrData(i), 1, 1) = "��" Then
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
    '�����ʽ��Ѫѹ�ķ�ʽ��ͬ,����ʽ����Ϊ6
    
    'ȥ��ǰ׺����жԱ�
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
    '�ж�ָ�����Ƿ��������жԽ��ߣ�mstrColWidth�ĸ�ʽ��765`11`1`1,765`11`2`1,...����������`�������`�жԽ��ߣ�
    
    IsDiagonal = (InStr(1, "," & mstrCatercorner & ",", "," & intCol - (cHideCols + VsfData.FixedCols - 1) & ",") <> 0)
End Function

Private Sub ISAssistant(ByVal lngOrder As Long, ByVal objTXT As TextBox)
    Dim intIndex As Integer, intType As Integer
    Dim objParent As Object
    Dim intRow As Integer, intCount As Integer, i As Integer, intGroupFirstRows As Integer, intHidden As Integer
    Dim strText As String, lngCount As Long
    Dim arrData, lngStartRow As Long
    '������Ŀ�ĳ��Ⱦ����Ƿ�������дʾ�ѡ��
    mblnEditAssistant = False
    mblnEditText = False
    cmdWord.Visible = mblnEditAssistant
    
    mrsItems.Filter = "��Ŀ���=" & lngOrder
    If mrsItems.RecordCount = 0 Then
        mrsItems.Filter = 0
        Exit Sub
    End If
    intType = mintType
    mblnEditAssistant = (mrsItems!��Ŀ���� = 1 And mrsItems!��Ŀ���� > 100) 'And Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) <= "1"
    mblnEditText = (mrsItems!��Ŀ���� = 1 And NVL(mrsItems!��Ŀ��ʾ, 0) = 0)
    If mblnEditText = True And mblnEditAssistant = False Then
        If UCase(objTXT.Name) = "TXTINPUT" Then
            cmdWord.Tag = -1  '��ʾtxtInput
        Else
            cmdWord.Tag = objTXT.Index
        End If
    End If
    mrsItems.Filter = 0
    lngStartRow = VsfData.ROW
    '��ȡ�������ݵĵ�һ��
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
    
    '�������ʾ�ѡ��,��ʾ����λ
    If mblnEditAssistant Then
        mintType = -1
        VsfData.ROW = lngStartRow
        mintType = intType
        
        If UCase(objTXT.Name) = "TXTINPUT" Then
            intIndex = -1 '��ʾtxtInput
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
        'Ϊ������ʱ��ѡ��������ʼ�У��༭������ʾ���д��ı���
        If Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) = 1 Then
            For intRow = 0 To intGroupFirstRows - 1
                strText = strText & IIf(strText = "", "", vbCrLf) & Replace(Replace(Replace(VsfData.TextMatrix(intRow + VsfData.ROW, VsfData.COL), Chr(13), ""), Chr(10), ""), Chr(1), "")
            Next intRow
            lngCount = VsfData.ROW + intGroupFirstRows - 1
            For intRow = VsfData.ROW + intGroupFirstRows To VsfData.Rows - 1
                If VsfData.RowHidden(intRow) = False Then
                    '������ÿһ�������������ݿ���ռ�ö��У�ֻ���ڵ�һ�б����˷�����������������������=ÿһ������������+ÿһ���������е�����
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
            '׼����ֵ
            With txtLength
                '������ʱ���еĿ�Ȳ���,Ϊ�˱��ⷵ�ض���,ǿ������Ϊ5000
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
    '��֤ÿҳ��Ч������
    
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
    
    If mintҳ�� <= mint����ҳ And mintҳ�� > 0 Then
        If mblnRestore = False Then
            VsfData.Rows = VsfData.Rows - Val(CStr(mlngLitterRows(mintҳ��)))
        Else
            '��ҳ�����޸ĺ�(������)mlngOverrunRowsֵ���Ϊ0���˴���Ҫ������֤���ݿ�������չʾ
            'ʾ����ǰ���ǿ�ҳ������ʾ�ڵ�ǰҳ�����ҵ�һҳ���ݿ�ҳ������ڶ�ҳ�������һ�����ݿ�ҳ(����������5,��2��,mlngOverrunRows=2),���޸��������ݵ���������
            '�����¸�������ʵ�������и�ֵ���ͻᵼ��mlngOverrunRows=0.ͨ���л�ҳ��ص���ҳ�ͻᵼ��VsfData.Rows����ȷ��
            '�㷨ʾ����mlngPageRows=20����һҳ��ҳ����5�У�mlngCurLitterRows(mintҳ��)=5�����ڶ�ҳ��ҳ����2�У�mlngOverrunRows=2����
            '1����һ�μ������ݵڶ�ҳ������Ӧ��Ϊ20-5+2=17����ʱ����޸ĵڶ�ҳ��ҳ����(���磺������37��Ϊ38)�ͻᵼ��mlngOverrunRows=0
            '2���л���һҳ�ڻص��ڶ�ҳ���ڶ�ҳ������������20-5-0=15�������ͻᵼ�µڶ�ҳ��ҳ���ݵ����������޷���ʾ��
            If lngCount > VsfData.Rows - Val(CStr(mlngCurLitterRows(mintҳ��))) - VsfData.FixedRows Then
                VsfData.Rows = lngCount + VsfData.FixedRows
            Else
                VsfData.Rows = VsfData.Rows - Val(CStr(mlngCurLitterRows(mintҳ��)))
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
    '������һ���ǿ���,�����༭״̬
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
    'intType��ȷ����������ֵ��1)������Ŀ���;2)�����к�
    'intMAX������ͬ����������ռ�õ��и�
    '����ͬ��������(һ���ļ��в����ܳ����ظ�����Ŀ,����,�ж�ʱ���ؼ���к�)
    
    lngRecord = Val(VsfData.TextMatrix(VsfData.ROW, mlngRecord))
    If lngRecord = 0 Then Exit Function
    
    gstrSQL = "" & _
        " SELECT  B.��Ŀ���,B.��Ŀ����,A.������� AS �к�" & vbNewLine & _
        " FROM �����ļ��ṹ A,���˻�����ϸ B" & vbNewLine & _
        " WHERE A.Ҫ������=B.��Ŀ���� AND A.��ID=" & vbNewLine & _
        "      (SELECT A.ID FROM �����ļ��ṹ A,���˻����ļ� B " & vbNewLine & _
        "       WHERE B.ID=[2] And A.�ļ�ID=B.��ʽID AND A.�������=4 AND A.��ID IS NULL)" & vbNewLine & _
        " AND B.������Դ>0 and B.������Դ <> 3 AND B.��¼ID=[1]"
    Call SQLDIY(gstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ͬ��������", lngRecord, mlng�ļ�ID)
    If rsTemp.RecordCount = 0 Then Exit Function
    
    '��ȡͬ�������Ϣ
    Do While Not rsTemp.EOF
        If InStr(1, "," & strCols & ",", "," & rsTemp!�к� & ",") = 0 Then strCols = strCols & "," & rsTemp!�к�
        strItems = strItems & "," & rsTemp!��Ŀ���
        strNames = strNames & "," & rsTemp!��Ŀ����
        rsTemp.MoveNext
    Loop
    strCols = Mid(strCols, 2)
    strItems = Mid(strItems, 2)
    strNames = Mid(strNames, 2)
    
    '������ѭ�����������ռ�и�
    If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then
        lngStartRow = VsfData.ROW
        lngEndRow = VsfData.ROW
        intInMAX = 1
    Else
        lngStartRow = GetStartRow(VsfData.ROW)
        intInMAX = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0))
        lngEndRow = lngStartRow + intInMAX - 1
    End If
    
    intCount = 1    'ͬ����ֻ������������Ŀ������ռ����ֻ������1�У��������ݲ�����Ҫ���
'    '����ռ�ó���1�вż��
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
        RaiseEvent AfterRowColChange("������,ʱ����,�Լ� " & strNames & " ��ͬ�����������ݣ��������޸Ļ�ɾ����", True, mblnSign, mblnArchive)
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
    '�����ݿ�����ȡ���ݣ������ǰ���Ŀ�д�������������������Ŀ����
    
    '�����Ŀ���뵽��ѯSQL�У���ʽ���к�;��ͷ����|��Ŀ���,��λ;��Ŀ���,��λ||�к�;��ͷ����...
    '�󶨶����Ŀ�����о��Զ�תΪ�Խ�����
    arrData = Split(mstrCOLActive, "||")
    intCount = UBound(arrData)
    For intDo = 0 To intCount
        intCol = Split(Split(arrData(intDo), "|")(0), ";")(0)
        If intCol = VsfData.COL - cHideCols - VsfData.FixedCols + 1 Then
            arrCol = Split(Split(arrData(intDo), "|")(1), ";")
            intMax = UBound(arrCol)
            For intIn = 0 To intMax
                strCond = strCond & " OR (��Ŀ���=" & Split(arrCol(intIn), ",")(0)
                If Split(arrCol(intIn), ",")(1) = "" Then
                    strCond = strCond & ")"
                Else
                    strCond = strCond & " AND NVL(���²�λ,'TWBW')='" & Split(arrCol(intIn), ",")(1) & "')"
                End If
            Next
            
            Exit For
        End If
    Next
    
    If strCond <> "" Then
        strCond = " AND (" & Mid(strCond, 4) & ")"
        '��ѯ���ݿ�
        gstrSQL = " SELECT  1 FROM ���˻�����ϸ A,���˻������� B,���˻����ӡ C" & vbNewLine & _
                  " Where A.��¼ID=B.ID And B.�������=0 And B.ID=C.��¼ID And C.�ļ�ID=B.�ļ�ID " & vbNewLine & _
                  " And C.�ļ�ID=[1] And (C.����ҳ��=[2] OR C.��ʼҳ��=[2])" & strCond & " AND ROWNUM<2"
        Call SQLDIY(gstrSQL)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯ���ݿ⵱ǰҳ��ָ������Ƿ���ڻ��Ŀ", mlng�ļ�ID, mintҳ��)
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
'�����ǻ������������

Private Sub txt����ʱ��_GotFocus()
    Call zlControl.TxtSelAll(txt����ʱ��)
End Sub

Private Sub txt����ʱ��_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim blnCancel As Boolean
    If KeyCode = vbKeyReturn Then
        Call txt����ʱ��_Validate(blnCancel)
        txtС������.SetFocus
    End If
End Sub

Private Sub txt����ʱ��_Validate(Cancel As Boolean)
    Dim strFormat As String
    Dim intDef As Integer   'ʱ��+1
    '��鿪ʼʱ��,����ʱ��Ϸ���
    
    strFormat = CheckTxtTime(txt��ʼʱ��)
    If strFormat = "" Then Exit Sub
    txt��ʼʱ��.Text = strFormat
    strFormat = CheckTxtTime(txt����ʱ��)
    If strFormat = "" Then Exit Sub
    txt����ʱ��.Text = strFormat
    
    '����С������
    If txt����ʱ��.Text > txt��ʼʱ��.Text Then
        intDef = Val(txt����ʱ��.Text) - Val(txt��ʼʱ��.Text)
    Else
        intDef = Val(txt����ʱ��.Text) + 24 - Val(txt��ʼʱ��.Text)
    End If
    '�����������59�����1Сʱ
    If Split(txt����ʱ��.Text, ":")(1) = "59" Then intDef = intDef + 1
    '71794:������,2014-05-06,��ʱС�᲻��һСʱҲ����С��
    '����33(��33)�Ժ�汾���˱�ʶֻ��һ�����ݺϷ����ж�
    If intDef = 0 Then
        txtС������.Tag = 1
        txtС������.Text = "����1СʱС��"
    Else
        txtС������.Tag = intDef
        txtС������.Text = intDef & "СʱС��"
    End If
End Sub

Private Sub txtС������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdOk.SetFocus
End Sub

Private Sub txt��ʼʱ��_GotFocus()
    Call zlControl.TxtSelAll(txt��ʼʱ��)
End Sub

Private Sub txt��ʼʱ��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txt����ʱ��.SetFocus
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
    lblUpInput.Caption = IIf(lblUpInput.Caption = "", "��", "")
    txtUpInput.SetFocus
End Sub

Private Sub lblDnInput_DblClick()
    lblDnInput.Caption = IIf(lblDnInput.Caption = "", "��", "")
    txtDnInput.SetFocus
End Sub

Private Sub lblInput_DblClick()
    lblInput.Caption = IIf(lblInput.Caption = "", "��", "")
End Sub

Private Sub txtUpInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtDnInput.SetFocus
    ElseIf KeyCode = vbKeyRight Then
        If txtUpInput.SelStart = Len(txtUpInput.Text) Then txtDnInput.SetFocus
    ElseIf KeyCode = vbKeyLeft And txtUpInput.SelStart = 0 Then
        Call MoveNextCell(False)
    ElseIf KeyCode = vbKeySpace And txtUpInput.Locked Then
        lblUpInput.Caption = IIf(lblUpInput.Caption = "", "��", "")
    End If
End Sub

Private Sub txtDnInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyRight And txtDnInput.SelStart = Len(txtDnInput.Text)) Then
        Call picDouble_KeyDown(KeyCode, Shift)
    ElseIf KeyCode = vbKeyLeft Then
        If txtDnInput.SelStart = 0 Then txtUpInput.SetFocus
    ElseIf KeyCode = vbKeySpace And txtDnInput.Locked Then
        lblDnInput.Caption = IIf(lblDnInput.Caption = "", "��", "")
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
        '�ƶ�����һ����Ԫ��
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

Private Sub txtС������_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = Asc(";") Then KeyAscii = 0
End Sub

Private Sub cboС��_Click()
    If cboС��.Tag = "" Then Exit Sub
    
    txt��ʼʱ��.Enabled = (cboС��.Text = "��ʱ")
    txt����ʱ��.Enabled = txt��ʼʱ��.Enabled
    If cboС��.Text <> "��ʱ" Then
        txt��ʼʱ��.Text = Split(Split(cboС��.Tag, ";")(cboС��.ListIndex), ",")(0)
        txt����ʱ��.Text = Split(Split(cboС��.Tag, ";")(cboС��.ListIndex), ",")(1)
        'txtС������.Text = Format(DateAdd("d", -1 * cboС�᷶Χ.ListIndex, zldatabase.Currentdate), "MM-DD") & " " & cboС��.Text
        txtС������.Text = Format(DTPDate.Value, "MM-DD") & " " & cboС��.Text
        txtС������.Tag = 0
    Else
        txtС������.Text = ""
        txt��ʼʱ��.Text = ""
        txt����ʱ��.Text = ""
        txtС������.Tag = 0
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
    '�����ַ���Ϊ���ݷָ�������¼�¼���ķָ�������˲�����¼��
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
    '���ϱ�ǩ��ɢ����
    Call zlLableBruit
End Sub

Private Sub UserControl_Terminate()
'    objStream.Close
End Sub

Private Sub SetDockRight(BarToDock As CommandBar, BarOnLeft As CommandBar)
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
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
    '�Ŵ����壬��Ԫ����
    BlowUp = dblChange
    If Not mblnBlowup Then Exit Function
    BlowUp = dblChange + (dblChange * 1 / 3)
End Function

Private Sub InitPages()
    Dim intPage As Integer
    Dim cbrItem As CommandBarControl
    Dim cbrCus As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom
    
    '����ҳѡ��
    If Not mcbrPage Is Nothing Then
        mcbrPage.CommandBar.Controls.DeleteAll
    Else
        Set mcbrPage = mcbrToolBar.Controls.Add(xtpControlPopup, clngPage, "ҳ��ѡ��")
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
        
    For intPage = mint��ʼҳ�� To mint����ҳ + 1
        Set cbrItem = mcbrPage.CommandBar.Controls.Add(xtpControlButton, conMenu_View_Jump, "��" & intPage & "ҳ", -1, False)
        cbrItem.Parameter = intPage
    Next
    
    txtPage.Text = ""
    mcbrPage.Caption = "ҳ��ѡ�񣺵�" & mintҳ�� & "ҳ"
    cbsThis.RecalcLayout
End Sub

Private Sub imgSign_Click()
    Call picSign_Click
End Sub

Private Sub lbl��֤ǩ��_Click()
    Call picSign_Click
End Sub

Private Sub picSign_Click()
    '����ǩ����ʷ��¼
    Dim str����ʱ�� As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    
    vsfSignData.Clear
    str����ʱ�� = VsfData.TextMatrix(VsfData.ROW, 2)
    gstrSQL = "" & _
        " SELECT A.��¼�� AS ǩ����,NVL(to_char(A.��¼ʱ��,'yyyy-MM-dd hh24:mi:ss'),A.��Ŀ����) AS ǩ��ʱ��,A.��¼���� AS ǩ����Ϣ,A.��¼��� AS ǩ������,A.ID,DECODE(A.��ĿID,NULL,'��Ч','δ��֤') AS ��Ч��,A.��ʼ�汾,NVL(A.��Ŀ���,2) AS ǩ������汾" & vbNewLine & _
        " FROM ���˻�����ϸ A,���˻������� B,���˻����ļ� C" & vbNewLine & _
        " WHERE A.��¼ID=B.ID And B.�ļ�ID=C.ID AND MOD(A.��¼����,10)=5" & vbNewLine & _
        " AND C.ID=[1] AND B.����ʱ��=[2] " & vbNewLine & _
        " Order by A.��Ŀ���� Desc"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡǩ����ʷ��¼", mlng�ļ�ID, CDate(str����ʱ��))
    
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

Private Sub cmdȡ��_Click()
    picSignCheck.Visible = False
End Sub

Private Sub cmdSignCur_Click()
    '������֤
    Dim lngLoop As Long
    Dim int�汾 As Integer
    Dim strSource As String, str����ʱ�� As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    
    If (Val(vsfSignData.TextMatrix(vsfSignData.ROW, 4)) = 0) Then Exit Sub
    If (Val(vsfSignData.TextMatrix(vsfSignData.ROW, 7)) < 2) Then
        MsgBox "����ǩ������仯���ϰ�ǩ�������ݲ�֧��ǩ��У�鹦�ܣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    '��ȡҪǩ��������
    '------------------------------------------------------------------------------------------------------------------
    int�汾 = vsfSignData.TextMatrix(vsfSignData.ROW, 6)
    str����ʱ�� = VsfData.TextMatrix(VsfData.ROW, 2)
    Set rsTemp = GetSignData(str����ʱ��, int�汾)
    Do While Not rsTemp.EOF
        For lngLoop = 0 To rsTemp.Fields.Count - 1
            strSource = strSource & CStr(zlCommFun.NVL(rsTemp.Fields(lngLoop).Value, ""))
        Next
        rsTemp.MoveNext
    Loop
    Debug.Print "��֤ǩ����" & Now & vbCrLf & strSource
    
    '����ǩ��
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
        MsgBox "����ǩ������δ����ȷ��װ����֤�������ܼ�����", vbInformation, gstrSysName
        Exit Sub
    End If
    If gobjESign.VerifySignature(strSource, Val(vsfSignData.TextMatrix(vsfSignData.ROW, 4)), 6) Then
        vsfSignData.TextMatrix(vsfSignData.ROW, 5) = "��Ч"
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
    'ȫ����֤
    
    lngSel = vsfSignData.ROW
    vsfSignData.Redraw = flexRDNone
    lngRows = vsfSignData.Rows - 1
    For lngRow = 1 To lngRows
        If (vsfSignData.TextMatrix(lngRow, 5) <> "��Ч") Then
            vsfSignData.ROW = lngRow
            Call cmdSignCur_Click
        End If
    Next
    vsfSignData.ROW = lngSel
    vsfSignData.Redraw = flexRDDirect
End Sub

Private Function ShowSignMarker(Optional ByVal bln�ⲿ As Boolean = False) As Boolean
    Dim str����ʱ�� As String
    Dim rsTemp As New ADODB.Recordset
    '��ʾ��ʷǩ�����
    
    picSign.Visible = False
    picSignCheck.Visible = False
    If Not bln�ⲿ Then
        If VsfData.COL <> mlngSignName Then Exit Function
    End If
    If VsfData.TextMatrix(VsfData.ROW, mlngSigner) = "" Then Exit Function
    
    str����ʱ�� = VsfData.TextMatrix(VsfData.ROW, 2)
    gstrSQL = "" & _
        " SELECT A.��¼�� AS ǩ����,NVL(to_char(A.��¼ʱ��,'yyyy-MM-dd hh24:mi:ss'),A.��Ŀ����) AS ǩ��ʱ��,A.��¼���� AS ǩ����Ϣ,A.��¼��� AS ǩ������,A.ID,DECODE(A.��ĿID,NULL,'��Ч','δ��֤') AS ��Ч��,A.��ʼ�汾,NVL(A.��Ŀ���,2) AS ǩ������汾" & vbNewLine & _
        " FROM ���˻�����ϸ A,���˻������� B,���˻����ļ� C" & vbNewLine & _
        " WHERE A.��¼ID=B.ID And B.�ļ�ID=C.ID AND MOD(A.��¼����,10)=5" & vbNewLine & _
        " AND C.ID=[1] AND B.����ʱ��=[2] And Rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡǩ����ʷ��¼", mlng�ļ�ID, CDate(str����ʱ��))
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
                Set objControl = .Add(xtpControlButton, conMenu_Tool_SignEarse, "ȡ��ǩ��"): objControl.IconId = 229
                Set objControl = .Add(xtpControlButton, conMenu_Tool_SignAuditCancel, "ȡ����ǩ"): objControl.IconId = 229
            End With
            objPopup.ShowPopup
        End If
    End If
End Sub

Private Sub vsfSignData_EnterCell()
    cmdSignCur.Enabled = (vsfSignData.TextMatrix(vsfSignData.ROW, 5) <> "��Ч")
End Sub

Private Function GetSignData(ByVal str����ʱ�� As String, ByVal int�汾 As Integer) As ADODB.Recordset
    On Error GoTo ErrHand
    Dim rsTemp As New ADODB.Recordset
    
    If int�汾 = 1 Then
        gstrSQL = "" & _
            "Select a.��¼����,a.��Ŀ����,a.��Ŀ���,a.��Ŀ����,a.��Ŀ����,a.��¼����,a.��Ŀ��λ,a.��¼���,a.���²�λ,a.��¼���,a.���Ժϸ�,a.δ��˵��,a.��¼��,a.��¼ʱ��" & vbNewLine & _
            "  From ���˻�����ϸ a, ���˻������� b,���˻����ļ� C" & vbNewLine & _
            " Where c.ID=[1] And b.����ʱ�� =[2]" & vbNewLine & _
            "   And a.��¼id = b.ID and B.�ļ�ID=C.ID and MOD(A.��¼����,10) <>5 and A.��ʼ�汾=1" & vbNewLine & _
            " ORDER BY ��Ŀ���"
    Else
        gstrSQL = "" & _
            "Select a.��¼����,a.��Ŀ����,a.��Ŀ���,a.��Ŀ����,a.��Ŀ����,a.��¼����,a.��Ŀ��λ,a.��¼���,a.���²�λ,a.��¼���,a.���Ժϸ�,a.δ��˵��,a.��¼��,a.��¼ʱ��" & vbNewLine & _
            "  From ���˻�����ϸ a, ���˻������� b,���˻����ļ� C" & vbNewLine & _
            " Where c.ID=[1] And b.����ʱ�� =[2]" & vbNewLine & _
            "   And a.��¼id = b.ID and B.�ļ�ID=C.ID and MOD(A.��¼����,10) <>5" & vbNewLine & _
            "   and (A.��ʼ�汾=[3] or (A.��ʼ�汾 <[3] and A.��ֹ�汾 IS NULL) or (A.��ʼ�汾<[3] and A.��ֹ�汾>[3]))" & vbNewLine & _
            " ORDER BY ��Ŀ���"
    End If
    Set GetSignData = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡָ���汾������", mlng�ļ�ID, CDate(str����ʱ��), int�汾)
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub SignMarker()
    '���ⲿ���������
    If Not ShowSignMarker(True) Then Exit Sub
    Call picSign_Click
End Sub

Private Sub SingerShowType(ByVal vsfObj As VSFlexGrid, ByVal lngStartRow As Long, ByVal lngEndRow As Long)
'-------------------------------------------------
'���ܣ���ʿǩ������ʾ��ʽ
''--58414,������,2013-04-10,��ӻ�ʿ��ǩ������ʾģʽ
'-------------------------------------------------
    Dim lngRow As Integer
    
    Select Case mlngSingerType
        Case 0 '��������ʾ
            For lngRow = lngStartRow To lngEndRow
                If mlngOperator > 0 Then vsfObj.TextMatrix(lngRow, mlngOperator) = vsfObj.TextMatrix(lngStartRow, mlngOperator)
                If mlngSignName > 0 Then vsfObj.TextMatrix(lngRow, mlngSignName) = vsfObj.TextMatrix(lngStartRow, mlngSignName)
                If mlngSignTime > 0 Then vsfObj.TextMatrix(lngRow, mlngSignTime) = vsfObj.TextMatrix(lngStartRow, mlngSignTime)
            Next
        Case 1 '������ʾ
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
        Case 3 'β����ʾ
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
        Case Else '��β��ʾ
            '���һ����Ҫ��д���ǩ��
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
'--58414,������,2013-04-10,��ӻ�ʿ��ǩ������ʾģʽ
    If mlngSingerType = 3 Then 'β��ǩ��,�ڿ�ʼ�����һ�е���Ϣ�������ʼ�У��Ա����SingerShowType������֯��ʾ��ʽ
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
'����:��ȡ������Ŀ�����������е���Ŀ��Ż��к�(�������)
'strKey ������Ŀ���кź����,��ʽ:�к�,���
'bytType 1:��Ŀ���,2:�к�
'blnCorrelative TRUE:�������,FALSE:��������
'����ֵ:Ϊ�ձ�ʾ������Ŀû�����ù�����
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
'����:�������м��������Ƿ�������ݣ�ֻҪ��һ�д��ھ��˳�
'��Σ�bytMode����Ҫ��Է������ݣ��Ǽ�������������ݻ���ֵ��������ݣ�0-���������,1- ֻ���������
'���Σ������в�Ϊ���򷵻��к�
    Dim strCols As String, strValue As String
    Dim i As Integer, arrCol
    Dim lngRow As Long, lngCol As Long, lngRows As Long, lngRowCount As Long
    If mstrColCollect <> "" Then
        '1����ȡ��������к�
        arrCol = Split(mstrColCollect, "|")
        For i = 0 To UBound(arrCol)
            strValue = GetRelatiionNo(CStr(arrCol(i)), 2)
            strCols = strCols & "," & IIf(strValue = "", "", strValue & ",") & Split(arrCol(i), ";")(0)
        Next
        strCols = Mid(strCols, 2)
        
        lngStartRow = GetStartRow(lngStartRow)
        '2������Ӧ�����Ƿ���ڻ�������
         '���lngStartRow���Ƿ�����ʼ�У����Ȼ�ȡ�������ݵĵ�һ��
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
        '��ȡ���ݵ�������
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
'��ҽ������������������������������ݴ���
    Dim cbrControl As CommandBarControl
    Dim rsImpAmount As ADODB.Recordset
    Dim strDate As String, blnImportName As Boolean, strValue As String
    Dim intCount As Integer, blnFind As Boolean, i As Integer, lngCurRow As Long
    Dim lngNameRow As Long, lngNumRow As Long '�����У�������
    Dim lngNameOrder As Long, lngNumOrder As Long '��������Ŀ���,��������Ŀ���
    
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
    
    '���ؼ�¼�����ݰ���:key,����,����
    Set rsImpAmount = frmImportOrder.ShowMe(Me, mlng�ļ�ID, mlng����ID, mlng��ҳID, mintӤ��, blnImportName, lngNumOrder, strDate)
    If rsImpAmount Is Nothing Then Call SetControlValue(lngNumOrder, "", False): Exit Sub
    If rsImpAmount.RecordCount = 0 Then Call SetControlValue(lngNumOrder, "", False): Exit Sub
    '��������
    If rsImpAmount.RecordCount > 0 Then rsImpAmount.MoveFirst
    For intCount = 1 To rsImpAmount.RecordCount
        VsfData.COL = lngNumRow
        If mblnShow = False Then Call VsfData_DblClick '׷�Ӻ��ȡ���༭���˴���Ҫ��������
        lngCurRow = GetStartRow(VsfData.ROW)
        If SetControlValue(lngNumOrder, NVL(rsImpAmount("����").Value)) = True Then
            'ȷ����Ŀ��Ӧ�ı༭�ؼ�
            If blnImportName Then
               VsfData.COL = lngNameRow
               Call SetControlValue(lngNameOrder, NVL(rsImpAmount("����").Value))
            End If
            If intCount = rsImpAmount.RecordCount Then
                Call MoveNextCell(True, True)
            Else
                Call MoveNextCell(VsfData.COL < mlngNoEditor - 1, True)
            End If
            '���ҽ����Ϣ��ֵ(������)
            If Record_Locate(mrsCellMap, "ID|" & mintҳ�� & "," & lngCurRow & "," & lngNumRow) = True Then
                mrsCellMap.Fields("���").Value = NVL(rsImpAmount("key").Value)
                mrsCellMap.Update
            End If
            
            If intCount < rsImpAmount.RecordCount Then
                '���֮ǰ�������˷��飬�˴���Ҫ���÷��飬ֱ��ʹ��׷��
                If mblnGroupNew = True Then Call cbsThis_Execute(cbsThis.FindControl(, conMenu_Edit_Group_New))
                'ʹ��׷�ӹ��ܣ�׷��һ��(ֱ���ڵ�ǰ������׷��)
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
    
    '��������ʾ��¼��ؼ�
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
'���ܣ�������Ŀ��������Ӧ�༭�ؼ��ĸ�ֵ(�����ڱ༭״̬��)
'       blnMode:True ��ֵ,False ���ý���
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
                        '��ѡ�ĵ�һ����Ŀ�����ѡ����Ҫ���������ѡ��Ŀ��ֱ�ӽ���
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
                If strValue = "��" Or strValue = "" Then lblInput.Caption = strValue
            Case 4
                If lngOrder = Val(txtUpInput.Tag) Then
                    txtUpInput.Text = strValue
                Else
                    txtDnInput.Text = strValue
                End If
            Case 5
                If lngOrder = Val(lblUpInput.Tag) Then
                     If strValue = "��" Or strValue = "" Then lblUpInput.Caption = strValue
                Else
                     If strValue = "��" Or strValue = "" Then lblDnInput.Caption = strValue
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
    '���ܣ�����ָ���м�¼ID��Ϣ,���������򷵻ظ��������ݵļ�¼ID���ԡ������ŷָ�
    Dim lngDemo As Long, lngStart As Long
    Dim strRecordid As String, lngStartID As Long
    
    If lngRow < VsfData.FixedRows Or lngRow > VsfData.Rows Then Exit Function
    lngStart = GetStartRow(lngRow)
    lngDemo = VsfData.TextMatrix(lngStart, mlngDemo)
    If lngDemo > 1 Then '����Ϊ��������,���ҵ���ʼ��
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
