VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.UserControl usrTendFileMutilEditor 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   5850
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11460
   KeyPreview      =   -1  'True
   ScaleHeight     =   5850
   ScaleWidth      =   11460
   Begin VSFlex8Ctl.VSFlexGrid vsfPrint 
      Height          =   1005
      Left            =   4845
      TabIndex        =   47
      Top             =   4140
      Visible         =   0   'False
      Width           =   3495
      _cx             =   6165
      _cy             =   1773
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
      GridColor       =   0
      GridColorFixed  =   0
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
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
      AutoResize      =   0   'False
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
   Begin VB.PictureBox picSplit 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   30
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   8415
      TabIndex        =   31
      Top             =   3840
      Width           =   8415
   End
   Begin VB.PictureBox pic�������� 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   450
      ScaleHeight     =   330
      ScaleWidth      =   10695
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   30
      Width           =   10695
      Begin VB.OptionButton optLevel 
         Caption         =   "����"
         Height          =   180
         Index           =   1
         Left            =   9720
         TabIndex        =   45
         Top             =   67
         Width           =   735
      End
      Begin VB.OptionButton optLevel 
         Caption         =   "����"
         Height          =   180
         Index           =   0
         Left            =   8640
         TabIndex        =   44
         Top             =   67
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.CommandButton cmdˢ�� 
         Caption         =   "ˢ��(&R)"
         Height          =   315
         Left            =   6660
         TabIndex        =   30
         Top             =   0
         Width           =   885
      End
      Begin VB.ComboBox cbo���� 
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   3600
         Style           =   2  'Dropdown List
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   0
         Width           =   1425
      End
      Begin VB.CheckBox chk��Ժ 
         Caption         =   "��Ժ"
         ForeColor       =   &H008080FF&
         Height          =   195
         Left            =   5940
         TabIndex        =   28
         ToolTipText     =   "��ѡ��ʾ��ȡ��Ժ����"
         Top             =   60
         Width           =   675
      End
      Begin VB.CheckBox chk���� 
         Caption         =   "����"
         ForeColor       =   &H008080FF&
         Height          =   195
         Left            =   5190
         TabIndex        =   27
         ToolTipText     =   "��ѡ��ʾ��ȡ���Ʋ���"
         Top             =   60
         Width           =   675
      End
      Begin VB.ComboBox cbo�����ļ���ʽ 
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   810
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   0
         Width           =   2205
      End
      Begin VB.Label lblEntry 
         AutoSize        =   -1  'True
         Caption         =   "¼�뷽ʽ"
         Height          =   180
         Left            =   7800
         TabIndex        =   46
         Top             =   67
         Width           =   720
      End
      Begin VB.Label lbl���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   3180
         TabIndex        =   25
         Top             =   60
         Width           =   360
      End
      Begin VB.Label lbl�ļ���ʽ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ļ���ʽ"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   60
         TabIndex        =   23
         Top             =   60
         Width           =   720
      End
   End
   Begin MSComctlLib.ImageList imgRow 
      Left            =   30
      Top             =   510
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
            Picture         =   "usrTendFileMutilEditor.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrTendFileMutilEditor.ctx":039A
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
      TabIndex        =   15
      Top             =   2550
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3315
      Left            =   60
      ScaleHeight     =   3315
      ScaleWidth      =   8385
      TabIndex        =   12
      Top             =   510
      Width           =   8385
      Begin VB.PictureBox PicLst 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H80000008&
         Height          =   2310
         Left            =   5310
         ScaleHeight     =   2280
         ScaleWidth      =   1185
         TabIndex        =   3
         Top             =   660
         Visible         =   0   'False
         Width           =   1215
         Begin VB.ListBox lstSelect 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            Height          =   1470
            Index           =   0
            ItemData        =   "usrTendFileMutilEditor.ctx":0734
            Left            =   -10
            List            =   "usrTendFileMutilEditor.ctx":074A
            TabIndex        =   5
            Top             =   825
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txtLst 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   -10
            MultiLine       =   -1  'True
            TabIndex        =   4
            Top             =   270
            Width           =   1215
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
            TabIndex        =   43
            Top             =   30
            Width           =   540
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
            TabIndex        =   42
            Top             =   615
            Width           =   540
         End
      End
      Begin VB.PictureBox picDoubleChoose 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   4680
         ScaleHeight     =   240
         ScaleWidth      =   900
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   360
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
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   0
            Width           =   375
            Begin VB.ComboBox cboChoose 
               BackColor       =   &H80000018&
               Height          =   300
               Index           =   1
               ItemData        =   "usrTendFileMutilEditor.ctx":0782
               Left            =   -30
               List            =   "usrTendFileMutilEditor.ctx":0792
               Style           =   2  'Dropdown List
               TabIndex        =   40
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
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   0
            Width           =   435
            Begin VB.ComboBox cboChoose 
               BackColor       =   &H80000018&
               Height          =   300
               Index           =   0
               ItemData        =   "usrTendFileMutilEditor.ctx":07A4
               Left            =   -30
               List            =   "usrTendFileMutilEditor.ctx":07B4
               Style           =   2  'Dropdown List
               TabIndex        =   38
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
            TabIndex        =   41
            Top             =   30
            Width           =   105
         End
      End
      Begin VB.CommandButton cmdWord 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6300
         Picture         =   "usrTendFileMutilEditor.ctx":07C6
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   90
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.PictureBox picDouble 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   5640
         ScaleHeight     =   240
         ScaleWidth      =   900
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   360
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
            TabIndex        =   18
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
               TabIndex        =   20
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
            TabIndex        =   17
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
               TabIndex        =   19
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
            TabIndex        =   16
            Top             =   30
            Width           =   105
         End
      End
      Begin VB.ListBox lstSelect 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   2550
         Index           =   1
         ItemData        =   "usrTendFileMutilEditor.ctx":0B08
         Left            =   6540
         List            =   "usrTendFileMutilEditor.ctx":0B1E
         Style           =   1  'Checkbox
         TabIndex        =   6
         Top             =   660
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.PictureBox picInput 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000018&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5970
         ScaleHeight     =   225
         ScaleWidth      =   585
         TabIndex        =   1
         Top             =   90
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
            TabIndex        =   14
            Top             =   30
            Width           =   315
         End
      End
      Begin VB.PictureBox picMutilInput 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   6810
         ScaleHeight     =   435
         ScaleWidth      =   1575
         TabIndex        =   10
         Top             =   150
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
            TabIndex        =   13
            Top             =   112
            Width           =   720
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid VsfData 
         Height          =   2655
         Left            =   0
         TabIndex        =   0
         Top             =   390
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
         FormatString    =   $"usrTendFileMutilEditor.ctx":0B56
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
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "һ�㻤���¼��"
         Height          =   180
         Left            =   3720
         TabIndex        =   29
         Top             =   90
         Width           =   1275
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfHistory 
      Height          =   1545
      Left            =   45
      TabIndex        =   32
      Top             =   3870
      Width           =   4305
      _cx             =   7594
      _cy             =   2725
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
      FormatString    =   $"usrTendFileMutilEditor.ctx":0BB8
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
   Begin VB.PictureBox picNull 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1245
      Left            =   1935
      ScaleHeight     =   1215
      ScaleWidth      =   7335
      TabIndex        =   33
      Top             =   1410
      Width           =   7365
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Ȼ����ˢ�°�ťװ������..."
         BeginProperty Font 
            Name            =   "����"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Index           =   1
         Left            =   30
         TabIndex        =   35
         Top             =   540
         Width           =   8115
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "��ѡ��һ�ֻ����ļ���ʽ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Index           =   0
         Left            =   30
         TabIndex        =   34
         Top             =   60
         Width           =   8145
      End
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
Attribute VB_Name = "usrTendFileMutilEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Public objFileSys As New FileSystemObject
'Public objStream As TextStream

Private mfrmParent As Object
Private mblnInit As Boolean
Private mblnHistory As Boolean              '��ʷ����Ƿ��ѳ�ʼ��
Private mblnShow As Boolean                 '�Ƿ���ʾ¼���
Private mblnBlowup As Boolean               '�Ŵ�񣿷Ŵ�1/3��������9�ŷŴ�Ϊ12��
Private mblnChange As Boolean               '�Ƿ��޸�����
Private mblnSaved As Boolean                '�Ƿ��ѱ���
Private mblnSigned As Boolean               '�Ƿ���ǩ��
Private mstrData As String                  '����༭״̬ǰ����֮ǰ������
Private mintPreDays As Long
Private mstrMaxDate As String
Private mlngSingerType As Long              '��ʿ��ǩ������ʾģʽ

Private mlng�ļ�ID As Long
Private mlng��ʽID As Long
Private mlng����ID As Long
Private mlng����ID As Long
Private mintҳ�� As Integer
Private mstrPrivs As String

Private mdtOutEnd As Date
Private mdtOutbegin As Date
Private mintChange As Integer
Private mstrBPItem As String                'ѪѹƵ�ζ�Ӧ��������ĿID,��ʽ1,2

Private mintSymbol As Integer               '��ǰ�ؼ�����
Private mstrSymbol As String                '�����ַ�
Private mstrCollectItems As String          '������Ŀ����
Private mstrColCollect As String              '������Ŀ�м���:col;1|col;4,5
Private mstrColCorrelative As String        '������Ŀ�����м���:COl,3;COl,4|COl,5;COl,6(�����к�,��Ŀ���;������,��Ŀ���),��Ҫ��Է������
Private mstrColImCorrelative As String    '������Ŀ�����м���:COl,3;COl,4|COl,5;COl,6(�����к�,��Ŀ���;������,��Ŀ���),��Ҫ�����������
Private mblnCorrelative As Boolean        '�Ƿ������˷������
Private mstrCOLNothing As String            'δ�󶨵��м���+���Ŀ��(���ܻ��Ŀ���Ƿ��)
Private mstrCOLActive As String             '��м���
Private mstrCatercorner As String           '�жԽ��߼���
Private mblnEditAssistant As Boolean        '��ǰѡ�����Ŀ�Ƿ�������дʾ�ѡ��
Private mblnEditText As Boolean             'ѡ�����Ŀ�Ƿ����ı���Ŀ
Private mblnEditHistoryAssistant As Boolean
Private mlngPageRows As Long                '���ļ���ʽһҳ����ʾ��������
Private mlngOverrunRows As Long             '����������
Private mlngRowCount As Long                '��ǰ��¼������
Private mlngRowCurrent As Long              '��ǰ��¼�ڱ�ҳ��ʵ������
Private mlngDate As Long                    '����
Private mlngTime As Long                    'ʱ��
Private mlngOperator As Long                '��ʿ
Private mlngSignLevel As Long               'ǩ������
Private mlngSigner As Long                  'ǩ����Ϣ
Private mlngSignName As Long                'ǩ����
Private mlngSignTime As Long                'ǩ��ʱ��
Private mlngRecord As Long                  '��¼ID
Private mlngNoEditor As Long                '��ֹ�༭��,���ڻ�ʿ�����Ի�ʿ��Ϊ׼,�����ڻ�ʿ������ǩ����Ϊ׼
Private mlngActiveTime As Long              '����ʱ��

Private mintType As Integer                 '��¼��ǰ�ı༭ģʽ
Private mblnDateAd As Boolean               '������д?
Private CellRect As RECT

Private rsTemp As New ADODB.Recordset
Private mrsItems As New ADODB.Recordset             '���л����¼��Ŀ�嵥
Private mrsSelItems As New ADODB.Recordset          '��ǰ¼��Ļ����¼��Ŀ�嵥
Private mrsDataMap As New ADODB.Recordset           '��ǰ����Ա¼������ݾ���,���¼����ʽһ��,���������ȫ�������Ա�Ѹ�ٻָ�
Private mrsCellMap As New ADODB.Recordset           '�༭�������ݾ���,�ֶ���:ҳ��,�к�,�к�,��¼ID,����,��λ,ɾ��
Private mrsCopyMap As New ADODB.Recordset           '����������

Private Enum ColIcon
    ǩ�� = 1
    ��ǩ = 2
End Enum
Private Enum SignLevel
    ���� = 1
    ���� = 2
    �м� = 3
    ʦ�� = 4
    Աʿ = 5
    δ���� = 9
End Enum

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

Public Event AfterRowColChange(ByVal strInfo As String, ByVal blnImportant As Boolean)
Public Event UsrHelp()
Public Event UsrExit()
'59118:������,2013-03-05
Public Event ShowTipInfo(ByVal vsfObj As Object, ByVal strInfo As String, ByVal blnMultiRow As Boolean)

Dim strFields As String
Dim strValues As String
Dim blnScroll As Boolean

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

Private Const WHITE_BRUSH = 0                   '��ɫ����
Private Const cdblWidth As Double = 6           'һ��Ӣ���ַ��Ŀ��
Private Const cHideCols = 8                    'ǰ׺��:����,����
Private Const cControlFields = 2                '��¼��������:ҳ��,�к�
Private Const mlngDemo As Long = 1                  '������
Private Const c���� As Integer = 1
Private Const c�ļ�ID As Integer = 2
Private Const c���� As Integer = 3
Private Const c���� As Integer = 4
Private Const c����ID As Integer = 5
Private Const c��ҳID As Integer = 6
Private Const cӤ�� As Integer = 7
Private Const cѪѹƵ�� As Integer = 8

Private Const pסԺ��ʿվ = 1262

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
    If Not mblnInit Then Exit Sub
    If VsfData.RowHidden(ROW) Then Exit Sub
    Done = False

    strText = VsfData.TextMatrix(ROW, COL)
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
'
'    '3������ǻ����У���������⴦��
'    If Val(VsfData.TextMatrix(Row, mlngCollectType)) < 0 And Val(VsfData.TextMatrix(Row, mlngCollectStyle)) = 1 _
'        And (Col >= mlngDate And Col < mlngNoEditor) Then
'        Call DrawCollectCell(hDC, Left, Top, Right, Bottom)
'    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Sub DrawCellHistory(ByVal hDC As Long, ByVal ROW As Long, ByVal COL As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
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
    If Not mblnInit Then Exit Sub
    If vsfHistory.RowHidden(ROW) Then Exit Sub
    Done = False

    strText = vsfHistory.TextMatrix(ROW, COL)
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
        If ROW < vsfHistory.FixedRows Then
            lngBackColor = GetRBGFromOLEColor(vsfHistory.BackColorFixed)
            lngForeColor = GetRBGFromOLEColor(vsfHistory.ForeColorFixed)
        Else
            If ROW = vsfHistory.RowSel Then
                lngBackColor = GetRBGFromOLEColor(vsfHistory.BackColorSel)
                lngForeColor = RGB(0, 0, 0)
            Else
                lngBackColor = RGB(255, 255, 255)
                lngForeColor = GetRBGFromOLEColor(vsfHistory.Cell(flexcpForeColor, ROW, COL))
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
'
'    '3������ǻ����У���������⴦��
'    If Val(vsfHistory.TextMatrix(Row, mlngCollectType)) < 0 And Val(vsfHistory.TextMatrix(Row, mlngCollectStyle)) = 1 _
'        And (Col >= mlngDate And Col < mlngNoEditor) Then
'        Call DrawCollectCell(hDC, Left, Top, Right, Bottom)
'    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
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

Private Function ReadStruDef() As Boolean
    Dim lngCol As Long
    On Error GoTo ErrHand

    '��ȡ�ļ�����
    mblnDateAd = False

    '��ȡ���Ŀ�������ж���(��ʽ���к�;��ͷ����|��Ŀ���,��λ;��Ŀ���,��λ||�к�;��ͷ����...)
    mstrCOLActive = ""
    mstrCOLNothing = ""
    mstrCollectItems = ""
    mstrColCollect = ""
    mstrColCorrelative = ""
    mstrColImCorrelative = ""
    mblnCorrelative = True
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
            Case "������":  VsfData.Cols = Val("" & !�����ı�): vsfHistory.Cols = Val("" & !�����ı�)
            Case "��С�и�": VsfData.RowHeightMin = BlowUp(Val("" & !�����ı�)): vsfHistory.RowHeightMin = BlowUp(Val("" & !�����ı�))
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
                Set vsfHistory.Font = objFont
                Set Font = objFont
            Case "�ı���ɫ"
                VsfData.ForeColor = Val("" & !�����ı�)
                vsfHistory.ForeColor = Val("" & !�����ı�)
            Case "�����ɫ"
                VsfData.GridColor = Val("" & !�����ı�): VsfData.GridColorFixed = VsfData.GridColor
                vsfHistory.GridColor = Val("" & !�����ı�): vsfHistory.GridColorFixed = vsfHistory.GridColor
            
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
                mlngPageRows = Val("" & !�����ı�)
            Case "�������"
                mblnCorrelative = (Val("" & !�����ı�) = 1)
            End Select
            .MoveNext
        Loop
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
    
    gstrSQL = "Select   d.�������,d.������,d.��������, d.�����д�, d.�����ı�, upper(d.Ҫ������) AS Ҫ������, d.Ҫ�ص�λ,d.Ҫ�ر�ʾ " & _
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
                mstrSQL�� = mstrSQL�� & ",l.ǩ����"
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
'                        'mstrSQL�� = mstrSQL�� & ", Decode(c.��Ŀ����, '" & !Ҫ������ & "', Decode(c.��¼����,Null,'" & !�����ı� & "'||'" & !Ҫ�ص�λ & "','" & !�����ı� & "'||c.��¼����||'" & !Ҫ�ص�λ & "'), '') As """ & !Ҫ������ & """"
'                        mstrSQL�� = mstrSQL�� & ", Decode(c.��Ŀ����, '" & !Ҫ������ & "', Decode(c.��¼����,Null,'" & !�����ı� & "'||'" & !Ҫ�ص�λ & "','" & !�����ı� & "'||c.��¼����||'" & !Ҫ�ص�λ & "'),  '" & !�����ı� & "'||'" & !Ҫ�ص�λ & "') As """ & !Ҫ������ & """"
'                    End If
                Else
'                    'Ϊ�ձ�ʾδ����,ǿ�Ƽ�,��������滻
                    mstrCOLNothing = mstrCOLNothing & "," & Val(Format(!�������, "00"))
'                    mstrSQL�� = mstrSQL�� & ",Max(""" & "C" & Format(!�������, "00") & """) As C" & Format(!�������, "00")
'                    mstrSQL���� = mstrSQL���� & " Or """ & "C" & Format(!�������, "00") & """ Is Not Null"
'                    mstrSQL�� = mstrSQL�� & ", C" & Format(!�������, "00") & " AS C" & Format(!�������, "00")
                End If
            End Select
            .MoveNext
        Loop

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
        If bln��ʿ = False Then mstrSQL�� = mstrSQL�� & ",l.������ As ��ʿ"

        If blnǩ���� = False Then mstrSQL�� = mstrSQL�� & ",l.ǩ���� As ǩ����"
        If blnǩ��ʱ�� = False Then mstrSQL�� = mstrSQL�� & ",l.ǩ��ʱ��"

        If Mid(mstrSQL��, 2) = "" Then
            MsgBox "�Բ�����û�ж��嵱ǰ��������ʾ����Ϣ�����ڲ����ļ������ж��壡", vbInformation, gstrSysName
            Exit Function
        End If

        '�����ڲ��������ӹ̶���
        mstrSQL�� = UCase(mstrSQL�� & ",MAX(ǩ������) AS ǩ������,MAX(ǩ����Ϣ) AS ǩ����Ϣ,MAX(��¼ID) AS ��¼ID,MAX(����) AS ����,MAX(ʵ������) AS ʵ������")
        mstrSQL�� = UCase(mstrSQL�� & ",l.ǩ������,l.ǩ���� AS ǩ����Ϣ,C.��¼ID,P.����||'' AS ����,1 AS ʵ������")
        mstrSQL�� = UCase(mstrSQL�� & ",ǩ������,ǩ����Ϣ,��¼ID,����,ʵ������")

        Call SQLCombination
    End With
    ReadStruDef = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub SQLCombination(Optional ByVal lng��¼ID As Long = 0)
    Dim str���� As String
    str���� = mstrSQL����
    
    mstrSQL = "Select   '' AS ����,0 AS �ļ�ID,'' AS ����,'' AS ����,0 AS ����ID,0 AS ��ҳID,0 AS Ӥ��,'' as ѪѹƵ��," & Mid(mstrSQL��, 12) & ",����ʱ��" & vbCrLf & _
                " From (Select ��¼���,ʱ�� as ����,to_char(����ʱ��,'yyyy-MM-dd hh24:mi:ss') ����ʱ��," & Mid(mstrSQL��, 2) & vbCrLf & _
                "        From (Select c.��¼���,l.����ʱ��," & Mid(mstrSQL��, 2) & vbCrLf & _
                "               From ���˻������� l, ���˻�����ϸ c,���˻����ļ� f,���˻����ӡ p " & vbCrLf & _
                "               Where l.ID=p.��¼ID And l.Id = c.��¼id And l.�ļ�ID+0=f.ID+0 And f.ID=p.�ļ�ID " & _
                "               And nvl(l.�������,0)=0 And c.��ֹ�汾 Is Null And c.��¼����<>5  " & _
                "               And f.id=[1] And 1=2)" & vbCrLf & _
                IIf(str���� <> "", "Where " & str����, "") & _
                "       Group By ����, ʱ��, ����ʱ��,��¼���,��ʿ,ǩ����,ǩ��ʱ��" & _
                                "       Order By ����, ʱ��, ����ʱ��,��¼���,��ʿ,ǩ����,ǩ��ʱ��)"
End Sub

Private Sub zlRefresh()
    Err = 0: On Error GoTo ErrHand
    Call InitCons

    '�����м�¼��
    Call InitRecords

    'װ������
    Call SQLCombination
    gstrSQL = mstrSQL
    Call SQLDIY(gstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��������", mlng�ļ�ID)
    '�����������¼���ṹ
    Call DataMap_Init(rsTemp)
    '�����ݲ����û����¼���ĸ�ʽ,ͬʱʵ��һ�����ݷ�����ʾ�Ĺ���
    Call PreTendFormat(rsTemp)
    
    '��ʼ����ʷ���
    Call PreTendFormatHistory(rsTemp)
    
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub DataMap_Init(ByVal rsSource As ADODB.Recordset)
    '��ʼ���ڴ����ݼ�

    '���ݼ�¼��,���ڿ��ٻָ�
    Set mrsDataMap = CopyNewRec(rsSource)
    mrsDataMap.Sort = "ҳ��,�к�"
    '�޸ĵ�Ԫ���¼,���ڱ���
    Call Record_Init(mrsCellMap, "ID," & adLongVarChar & ",50|ҳ��," & adDouble & ",18|�к�," & adDouble & ",18|" & _
            "�к�," & adDouble & ",18|��¼ID," & adDouble & ",18|����," & adLongVarChar & ",4000|��λ," & adLongVarChar & ",100|" & _
            "���," & adLongVarChar & ",100|����," & adDouble & ",1|ɾ��," & adDouble & ",1")
    mrsCellMap.Sort = "ҳ��,�к�,�к�"
    '���Ƽ�¼��
    Set mrsCopyMap = New ADODB.Recordset
    Set mrsCopyMap = CopyNewRec(mrsDataMap, False)
End Sub

Private Function DataMap_Save() As Boolean
    '����ǰҳ�����û��༭�������ݱ�������,ҳ���л��򱣴�ǰ����
    Dim blnExit As Boolean, blnNULL As Boolean
    Dim lngRow As Long, lngCol As Long, lngRows As Long, lngCols As Long
    Dim intCount As Integer
    Dim arrRows()
    On Error GoTo ErrHand
    
    '��ɾ��ָ��ҳ�ŵ�����������
    If mrsDataMap.RecordCount <> 0 Then mrsDataMap.MoveFirst
    Do While True
        If mrsDataMap.EOF Then Exit Do
        mrsDataMap.Delete
        mrsDataMap.MoveNext
    Loop
    
    '����ָ��ҳ�ŵ�����������
    lngRows = VsfData.Rows - 1
    lngCols = VsfData.Cols - 1
    
    arrRows = Array()
    '������ݼ�¼ID
    For lngRow = VsfData.FixedRows To lngRows
        blnNULL = True
        If VsfData.RowHidden(lngRow) = False Then
            For intCount = mlngTime + 1 To mlngNoEditor - 1
                If Not VsfData.ColHidden(intCount) Then
                    If VsfData.TextMatrix(lngRow, intCount) <> "" And Not (IsDiagonal(intCount) And InStr(1, VsfData.TextMatrix(lngRow, intCount), "/") <> 0) Then
                        blnNULL = False
                        Exit For
                    End If
                End If
            Next
            If blnNULL Then
               VsfData.TextMatrix(lngRow, mlngRecord) = ""
            End If
        Else
            VsfData.TextMatrix(lngRow, mlngRecord) = ""
            ReDim Preserve arrRows(UBound(arrRows) + 1)
            arrRows(UBound(arrRows)) = lngRow
        End If
    Next
    '������ص���
    For lngRow = UBound(arrRows) To 0 Step -1
        If VsfData.ROW >= Val(arrRows(lngRow)) Then
            VsfData.ROW = VsfData.ROW - 1
        End If
        VsfData.RemoveItem Val(arrRows(lngRow))
    Next lngRow
    
    lngRows = VsfData.Rows - 1
    '����������
    For lngRow = VsfData.FixedRows To lngRows
        mrsDataMap.AddNew
        mrsDataMap!ҳ�� = mintҳ��
        mrsDataMap!�к� = lngRow
        mrsDataMap!ɾ�� = IIf(VsfData.RowHidden(lngRow), 1, 0)
        For lngCol = 0 To lngCols - VsfData.FixedCols
            mrsDataMap.Fields(cControlFields + lngCol).Value = IIf(VsfData.TextMatrix(lngRow, lngCol + VsfData.FixedCols) = "", Null, VsfData.TextMatrix(lngRow, lngCol + VsfData.FixedCols))
        Next
        mrsDataMap.Update
    Next
    
    DataMap_Save = True
    
    'ˢ����ʷ����
    Call RefreshHistoryData(VsfData.ROW)
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function DataMap_Restore() As Boolean
    '��ָ��ҳ������ݻָ��������
    Dim lngRow As Long, lngCol As Long, lngRows As Long, lngCols As Long
    On Error GoTo ErrHand
    
    mrsDataMap.MoveFirst
    lngCols = VsfData.Cols - 1
    lngRows = mrsDataMap.RecordCount
    VsfData.Rows = VsfData.FixedRows
    For lngRow = 0 To lngRows - 1
        If lngRow > VsfData.Rows - VsfData.FixedRows - 1 Then VsfData.Rows = VsfData.Rows + 1
        For lngCol = 0 To lngCols - VsfData.FixedCols
            VsfData.TextMatrix(VsfData.FixedRows + lngRow, lngCol + VsfData.FixedCols) = NVL(mrsDataMap.Fields(cControlFields + lngCol).Value)
        Next
        If mrsDataMap!ɾ�� = 1 Then VsfData.RowHidden(VsfData.FixedRows + lngRow) = True
        mrsDataMap.MoveNext
    Next
    
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
    Dim lngRowCount As Long, lngRowCurrent As Long  '��ǰ��¼������,��ǰ��¼�ڱ�ҳ��ʵ������
    Dim lngCol As Long, lngMax As Long
    Dim lngRow As Long, lngLastRow As Long
    Dim str����ʱ�� As String, str����ʱ��_L As String
    Dim lngStart As Long, lngPrintedRow As Long
    Dim strSignName As String
    Dim blnClear As Boolean
    
    On Error GoTo ErrHand

    Dim arrData
    Dim intData As Integer, intDatas As Integer
    '���һ����ʾ�����������ʾ(���ݵ�ǰ����ռ����������ӿհ��в�����������,Ȼ�������δ���ǰ�е�����)
    'ÿҳֻ��ʾʵ�ʵ�������,��'@��ȡ��ע�ͼ���

    lngRow = vsfHistory.FixedRows
    Do While True
        If lngRow > vsfHistory.Rows - 1 Then Exit Do
        'If lngRow >= mlngPageRows + mlngOverrunRows + vsfHistory.FixedRows Then Exit Do
        If InStr(1, vsfHistory.TextMatrix(lngRow, mlngRowCount), "|") <> 0 Then Exit Do
        lngRowCount = Val(vsfHistory.TextMatrix(lngRow, mlngRowCount))
        '@ʵ��������
'        lngRowCurrent = Val(vsfhistory.TextMatrix(lngRow, mlngRowCurrent))
        str����ʱ�� = Format(vsfHistory.TextMatrix(lngRow, mlngActiveTime), "YYYY-MM-DD HH:mm:ss")
        If str����ʱ��_L <> "" And Mid(str����ʱ��_L, 1, 16) = Mid(str����ʱ��, 1, 16) Then
            '������ͬ��������ͬ���Ҳ��ǻ��������У���˵����Щ������һ�飬����lngDemo��
            vsfHistory.TextMatrix(lngRow, mlngDate) = ""
            vsfHistory.TextMatrix(lngRow, mlngTime) = ""
            vsfHistory.TextMatrix(lngRow, mlngDemo) = lngRow - lngLastRow + 1
            If lngRow - lngLastRow = Val(vsfHistory.TextMatrix(lngLastRow, mlngRowCount)) Then
                vsfHistory.TextMatrix(lngLastRow, mlngDemo) = 1
            End If
        Else
            lngLastRow = lngRow
        End If
        
        If lngRowCount > 1 Then
            '�����ӿ���
            vsfHistory.Rows = vsfHistory.Rows + lngRowCount - 1
            '�ӵ�ǰ�е���һ�п�ʼ��ÿ�е�λ��+�����ӵĿհ���������֤�����Ŀհ��дӵ�ǰ�е���һ�п�ʼ
            For intData = vsfHistory.Rows - lngRowCount To lngRow + 1 Step -1
                vsfHistory.RowPosition(intData) = intData + lngRowCount - 1
            Next

            'ѭ������ǰ������
            For lngCol = 0 To vsfHistory.Cols - 1
                If vsfHistory.ColHidden(lngCol) And lngCol <> mlngRowCount And lngCol <> mlngDemo Then
                    'ѭ����ֵ
                    For intData = 2 To lngRowCount
                        vsfHistory.TextMatrix(lngRow + intData - 1, lngCol) = vsfHistory.TextMatrix(lngRow, lngCol)
                    Next
                ElseIf (lngCol < mlngNoEditor And lngCol <> mlngDate And lngCol <> mlngTime) Then
                    '׼����ֵ
                    With txtLength
                        .Width = vsfHistory.ColWidth(lngCol)
                        .Text = Replace(Replace(Replace(vsfHistory.TextMatrix(lngRow, lngCol), Chr(10), ""), Chr(13), ""), Chr(1), "")
                        .FontName = vsfHistory.CellFontName
                        .FontSize = vsfHistory.CellFontSize
                        .FontBold = vsfHistory.CellFontBold
                        .FontItalic = vsfHistory.CellFontItalic
                    End With
                    arrData = GetData(txtLength.Text)
                    intDatas = UBound(arrData)

                    If intDatas > 0 Then
                        'ѭ����ֵ
                        For intData = 0 To intDatas
                            vsfHistory.TextMatrix(lngRow + intData, lngCol) = Replace(Replace(Replace(arrData(intData), Chr(10), ""), Chr(13), ""), Chr(1), "")
                        Next
                    End If
                ElseIf lngCol = mlngNoEditor Then
                        '����ֵ��Ϊ��1��ʼ,������4������,����4|1
                        For intData = 1 To lngRowCount
                            vsfHistory.TextMatrix(lngRow + intData - 1, mlngRowCount) = lngRowCount & "|" & intData
                        Next
                        '���һ����Ҫ��д���ǩ��
                        If mlngSignName > 0 Then vsfHistory.TextMatrix(lngRow + lngRowCount - 1, mlngSignName) = vsfHistory.TextMatrix(lngRow, mlngSignName)
                        If mlngSignTime > 0 Then vsfHistory.TextMatrix(lngRow + lngRowCount - 1, mlngSignTime) = vsfHistory.TextMatrix(lngRow, mlngSignTime)
                        '--58414,������,2013-04-10,��ӻ�ʿ��ǩ������ʾģʽ
                        Call SingerShowType(vsfHistory, lngRow, lngRow + lngRowCount - 1)
                    Else
                End If
            Next
            '@ʵ��������
'            '�����ҳ��һ�е����ݲ�ȫ,���Ƚ��ü�¼��һ�е�������(����,ʱ��,ǩ��)��Ϣ���Ƶ�
'            If lngRow = vsfhistory.FixedRows And lngRowCount <> lngRowCurrent Then
'                '�̶�������ʾ����ʱ����ǩ����
'                lngMax = lngRowCount - lngRowCurrent
'                If mlngDate > -1 Then vsfhistory.TextMatrix(lngRow + lngMax, mlngDate) = vsfhistory.TextMatrix(lngRow, mlngDate)
'                If mlngTime > -1 Then vsfhistory.TextMatrix(lngRow + lngMax, mlngTime) = vsfhistory.TextMatrix(lngRow, mlngTime)
'                if mlngOperator <>-1 then vsfhistory.TextMatrix(lngRow + lngMax, mlngOperator) = vsfhistory.TextMatrix(lngRow, mlngOperator)
'                if mlngOperator <>-1 then vsfhistory.TextMatrix(lngRow + lngMax, mlngsignname) = vsfhistory.TextMatrix(lngRow, mlngsignname)
'                'ɾ���������
'                For lngCol = 1 To lngMax
'                    vsfhistory.RemoveItem lngRow
'                Next
'            End If
'            lngRow = lngRow + lngRowCurrent - 1 '���ϸü�¼�ڱ�ҳʵ�ʵ�����
            '@ʵ��������Ҫע���������д���
            lngRow = lngRow + lngRowCount - 1
        Else
            vsfHistory.TextMatrix(lngRow, mlngRowCount) = "1|1"
        End If
        lngRow = lngRow + 1
        str����ʱ��_L = str����ʱ��
    Loop
    
    '63760:������,�������ݻ�ʿ��ǩ���ˡ�ǩ��ʱ��Ĵ���ͬһ��ǩ����ʼ����ʾһ�Σ�
    If mlngSingerType > 0 And vsfHistory.FixedRows <= vsfHistory.Rows - 1 Then
        lngPrintedRow = IIf(mlngOperator <> -1, mlngOperator, mlngSignName)
        lngRow = vsfHistory.FixedRows
        Do While True
            lngStart = GetStartRowHistory(lngRow)
            lngRowCount = Val(vsfHistory.TextMatrix(lngStart, mlngRowCount))
            If lngRowCount <= 0 Then Exit Do
            
            If mlngSingerType = 3 Then 'β��ǩ��
                strSignName = vsfHistory.TextMatrix(lngStart + lngRowCount - 1, lngPrintedRow)
            Else '����ǩ������βǩ��
                strSignName = vsfHistory.TextMatrix(lngStart, lngPrintedRow)
            End If
            strSignName = FormatValue(strSignName)
            
            If Val(vsfHistory.TextMatrix(lngStart, mlngDemo)) = 1 And lngStart = lngRow And strSignName <> "" Then
                For lngRow = lngStart + lngRowCount To vsfHistory.Rows - 1
                    If lngRow = lngStart + lngRowCount Then
                    
                        If Val(vsfHistory.TextMatrix(lngRow, mlngDemo)) <= 1 Then Exit For
                        
                        lngRowCount = Val(vsfHistory.TextMatrix(lngRow, mlngRowCount))
                        If lngRowCount = 0 Then Exit For
                        
                        If mlngSingerType = 3 Then 'β��ǩ��
                            If strSignName = FormatValue(vsfHistory.TextMatrix(lngRow + lngRowCount - 1, lngPrintedRow)) Then
                                If lngStart <= lngRow - 1 Then
                                    If mlngOperator <> -1 Then vsfHistory.TextMatrix(lngRow - 1, mlngOperator) = ""
                                    If mlngSignName <> -1 Then vsfHistory.TextMatrix(lngRow - 1, mlngSignName) = ""
                                    If mlngSignTime <> -1 Then vsfHistory.TextMatrix(lngRow - 1, mlngSignTime) = ""
                                End If
                            Else
                                If FormatValue(vsfHistory.TextMatrix(lngRow + lngRowCount - 1, lngPrintedRow)) <> "" Then
                                    strSignName = FormatValue(vsfHistory.TextMatrix(lngRow + lngRowCount - 1, lngPrintedRow))
                                End If
                            End If
                        Else '����ǩ������βǩ��
                            If strSignName = FormatValue(vsfHistory.TextMatrix(lngRow, lngPrintedRow)) Then
                                '����ǩ������βǩ������Ҫȥ����һ�����ݵ�����,����βǩ����Ҫע������е����һ����������=1�����
                                blnClear = True
                                If mlngSingerType = 2 And lngRowCount = 1 Then
                                    If lngRow + lngRowCount < vsfHistory.Rows Then
                                        If Val(vsfHistory.TextMatrix(lngRow + lngRowCount, mlngDemo)) <= 1 Then
                                            blnClear = False
                                        End If
                                    Else
                                        blnClear = False
                                    End If
                                End If
                                If blnClear Then
                                    If mlngOperator <> -1 Then vsfHistory.TextMatrix(lngRow, mlngOperator) = ""
                                    If mlngSignName <> -1 Then vsfHistory.TextMatrix(lngRow, mlngSignName) = ""
                                    If mlngSignTime <> -1 Then vsfHistory.TextMatrix(lngRow, mlngSignTime) = ""
                                End If
                                '��βǩ����Ӧ��ȥ����һ�����ݵ�β��(��һ������������Ҫ>1)
                                If mlngSingerType = 2 And lngStart < lngRow - 1 Then
                                    If mlngOperator <> -1 Then vsfHistory.TextMatrix(lngRow - 1, mlngOperator) = ""
                                    If mlngSignName <> -1 Then vsfHistory.TextMatrix(lngRow - 1, mlngSignName) = ""
                                    If mlngSignTime <> -1 Then vsfHistory.TextMatrix(lngRow - 1, mlngSignTime) = ""
                                End If
                            Else
                                If FormatValue(vsfHistory.TextMatrix(lngRow, lngPrintedRow)) <> "" Then
                                    strSignName = FormatValue(vsfHistory.TextMatrix(lngRow, lngPrintedRow))
                                End If
                            End If
                        End If
                        
                        lngStart = lngRow
                    End If
                Next lngRow
            Else
                lngRow = lngStart + Val(vsfHistory.TextMatrix(lngStart, mlngRowCount))
            End If
            
            If lngRow > vsfHistory.Rows - 1 Then Exit Do
        Loop
    End If
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub PreTendFormat(ByVal rsTemp As ADODB.Recordset)
    Dim aryItem() As String
    Dim lngRow As Long, lngCol As Long, lngCount As Long, strCell As String
    Dim stdPicFont As StdFont
    On Error GoTo ErrHand

    '���û����¼���ĸ�ʽ
    With VsfData
        .Redraw = flexRDNone
        .Clear
        Set .DataSource = rsTemp

        '��ͷ��д
        .MergeCells = flexMergeFixedOnly
        .MergeCellsFixed = flexMergeFree
        .MergeRow(0) = True
        .MergeRow(1) = True
        .MergeRow(2) = True

        '�����ڲ�����������\
        .ColHidden(c����) = True
        .ColHidden(c�ļ�ID) = True
        .ColHidden(c����ID) = True
        .ColHidden(c��ҳID) = True
        .ColHidden(cӤ��) = True
        .ColHidden(cѪѹƵ��) = (mstrBPItem = "")
        .ColHidden(mlngRowCurrent) = True
        .ColHidden(mlngRowCount) = True
        .ColHidden(mlngRecord) = True
        .ColHidden(mlngSigner) = True
        .ColHidden(mlngSignLevel) = True
        .ColHidden(mlngActiveTime) = True
        .ColWidth(0) = 250
        .ColWidth(c����) = 1500
        .ColAlignment(c����) = flexAlignRightCenter
        Set stdPicFont = picMain.Font
        Set picMain.Font = .Font
        .ColWidth(cѪѹƵ��) = (picMain.TextWidth("Ѫ") * 5)
        Set picMain.Font = stdPicFont
        
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
        .TextMatrix(0, c����) = "����"
        .TextMatrix(1, c����) = "����"
        .TextMatrix(2, c����) = "����"
        .TextMatrix(0, c�ļ�ID) = "�ļ�ID"
        .TextMatrix(1, c�ļ�ID) = "�ļ�ID"
        .TextMatrix(2, c�ļ�ID) = "�ļ�ID"
        .TextMatrix(0, c����) = "����"
        .TextMatrix(1, c����) = "����"
        .TextMatrix(2, c����) = "����"
        .TextMatrix(0, c����) = "����"
        .TextMatrix(1, c����) = "����"
        .TextMatrix(2, c����) = "����"
        .TextMatrix(0, c����ID) = "����ID"
        .TextMatrix(1, c����ID) = "����ID"
        .TextMatrix(2, c����ID) = "����ID"
        .TextMatrix(0, c��ҳID) = "��ҳID"
        .TextMatrix(1, c��ҳID) = "��ҳID"
        .TextMatrix(2, c��ҳID) = "��ҳID"
        .TextMatrix(0, cӤ��) = "Ӥ��"
        .TextMatrix(1, cӤ��) = "Ӥ��"
        .TextMatrix(2, cӤ��) = "Ӥ��"
        .TextMatrix(0, cѪѹƵ��) = "ѪѹƵ��"
        .TextMatrix(1, cѪѹƵ��) = "ѪѹƵ��"
        .TextMatrix(2, cѪѹƵ��) = "ѪѹƵ��"

        '�п�����
        Dim blnAlign As Boolean
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

        If .Rows = .FixedRows Then
            mlngOverrunRows = 0
        Else
            '�õ���һ�еĳ�����
            mlngOverrunRows = Val(.TextMatrix(3, mlngRowCount)) - Val(.TextMatrix(3, mlngRowCurrent))
            '�������һ�еĳ�����
            mlngOverrunRows = mlngOverrunRows + Val(.TextMatrix(.Rows - 1, mlngRowCount)) - Val(.TextMatrix(.Rows - 1, mlngRowCurrent))
        End If

        Call FillPage
        Call WriteColor
        
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
        
        .ROW = .FixedRows
        .Redraw = flexRDDirect
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub PreTendFormatHistory(ByVal rsTemp As ADODB.Recordset)
    Dim aryItem() As String
    Dim lngRow As Long, lngCol As Long, lngCount As Long, strCell As String
    On Error GoTo ErrHand

    '���û����¼���ĸ�ʽ
    With vsfHistory
        .Redraw = flexRDNone
        .Clear
        Set .DataSource = rsTemp

        '��ͷ��д
        .MergeCells = flexMergeFixedOnly
        .MergeCellsFixed = flexMergeFree
        .MergeRow(0) = True
        .MergeRow(1) = True
        .MergeRow(2) = True

        '�����ڲ�����������
        .ColHidden(c����) = True
        .ColHidden(c�ļ�ID) = True
        .ColHidden(c����ID) = True
        .ColHidden(c��ҳID) = True
        .ColHidden(cӤ��) = True
        .ColHidden(c����) = True
        .ColHidden(c����) = True
        .ColHidden(cѪѹƵ��) = True
        .ColHidden(mlngRowCurrent) = True
        .ColHidden(mlngRowCount) = True
        .ColHidden(mlngRecord) = True
        .ColHidden(mlngSigner) = True
        .ColHidden(mlngSignLevel) = True
        .ColHidden(mlngActiveTime) = True
        .ColWidth(0) = 250

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
        .TextMatrix(0, c����) = "����"
        .TextMatrix(1, c����) = "����"
        .TextMatrix(2, c����) = "����"
        .TextMatrix(0, c�ļ�ID) = "�ļ�ID"
        .TextMatrix(1, c�ļ�ID) = "�ļ�ID"
        .TextMatrix(2, c�ļ�ID) = "�ļ�ID"
        .TextMatrix(0, c����) = "����"
        .TextMatrix(1, c����) = "����"
        .TextMatrix(2, c����) = "����"
        .TextMatrix(0, c����) = "����"
        .TextMatrix(1, c����) = "����"
        .TextMatrix(2, c����) = "����"
        .TextMatrix(0, c����ID) = "����ID"
        .TextMatrix(1, c����ID) = "����ID"
        .TextMatrix(2, c����ID) = "����ID"
        .TextMatrix(0, c��ҳID) = "��ҳID"
        .TextMatrix(1, c��ҳID) = "��ҳID"
        .TextMatrix(2, c��ҳID) = "��ҳID"
        .TextMatrix(0, cӤ��) = "Ӥ��"
        .TextMatrix(1, cӤ��) = "Ӥ��"
        .TextMatrix(2, cӤ��) = "Ӥ��"
        .TextMatrix(0, cѪѹƵ��) = "Ӥ��"
        .TextMatrix(1, cѪѹƵ��) = "Ӥ��"
        .TextMatrix(2, cѪѹƵ��) = "Ӥ��"

        '�п�����
        Dim blnAlign As Boolean
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

        If .Rows = .FixedRows Then
            mlngOverrunRows = 0
        Else
            '�õ���һ�еĳ�����
            mlngOverrunRows = Val(.TextMatrix(3, mlngRowCount)) - Val(.TextMatrix(3, mlngRowCurrent))
            '�������һ�еĳ�����
            mlngOverrunRows = mlngOverrunRows + Val(.TextMatrix(.Rows - 1, mlngRowCount)) - Val(.TextMatrix(.Rows - 1, mlngRowCurrent))
        End If
        
        Call PreTendMutilRows
        Call WriteColorHistory
        
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
        
        .Redraw = flexRDDirect
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub WriteColor()
    Dim blnTag As Boolean
    Dim lngCount As Long
    '����Ժ�ɫ��ʾ��ͬʱ������ʼ������ΪNoCheckBox������ͼ��
    With VsfData
        For lngCount = .FixedRows To .Rows - 1
            If .TextMatrix(lngCount, 0) <> "" Then
                '����Ժ�ɫ��ʾ
                blnTag = False
                If mintTagFormHour < mintTagToHour Then
                    blnTag = (Hour(.TextMatrix(lngCount, 0)) >= mintTagFormHour And Hour(.TextMatrix(lngCount, 0)) < mintTagToHour)
                Else
                    blnTag = (Hour(.TextMatrix(lngCount, 0)) >= mintTagFormHour Or Hour(.TextMatrix(lngCount, 0)) < mintTagToHour)
                End If
                If blnTag Then
                    Set .Cell(flexcpFont, lngCount, 0, lngCount, .Cols - 1) = mobjTagFont
                    .Cell(flexcpForeColor, lngCount, 0, lngCount, .Cols - 1) = mlngTagColor
                End If
            End If

            '������ʼ������ΪNoCheckBox
            If Not VsfData.RowHidden(lngCount) Then
                If FormatValue(VsfData.TextMatrix(lngCount, mlngRowCount)) Like "*|1" Then
                    '����ͼ��
                    If VsfData.TextMatrix(lngCount, mlngSigner) = "" Then
                        VsfData.Cell(flexcpPicture, lngCount, 0) = Nothing
                    Else
                        If InStr(1, VsfData.TextMatrix(lngCount, mlngSigner), "/") <> 0 Then
                            VsfData.Cell(flexcpPicture, lngCount, 0) = imgRow.ListImages(��ǩ).Picture
                        Else
                            VsfData.Cell(flexcpPicture, lngCount, 0) = imgRow.ListImages(ǩ��).Picture
                        End If
                    End If
                End If
            End If
        Next
    End With
    
    Call SetActiveColColor
End Sub

Private Sub WriteColorHistory()
    Dim blnTag As Boolean
    Dim lngCount As Long
    '����Ժ�ɫ��ʾ��ͬʱ������ʼ������ΪNoCheckBox������ͼ��
    With VsfData
        For lngCount = .FixedRows To .Rows - 1
            If .TextMatrix(lngCount, 0) <> "" Then
                '����Ժ�ɫ��ʾ
                blnTag = False
                If mintTagFormHour < mintTagToHour Then
                    blnTag = (Hour(.TextMatrix(lngCount, 0)) >= mintTagFormHour And Hour(.TextMatrix(lngCount, 0)) < mintTagToHour)
                Else
                    blnTag = (Hour(.TextMatrix(lngCount, 0)) >= mintTagFormHour Or Hour(.TextMatrix(lngCount, 0)) < mintTagToHour)
                End If
                If blnTag Then
                    Set .Cell(flexcpFont, lngCount, 0, lngCount, .Cols - 1) = mobjTagFont
                    .Cell(flexcpForeColor, lngCount, 0, lngCount, .Cols - 1) = mlngTagColor
                End If
            End If

            '������ʼ������ΪNoCheckBox
            If Not VsfData.RowHidden(lngCount) Then
                If FormatValue(VsfData.TextMatrix(lngCount, mlngRowCount)) Like "*|1" Then
                    '����ͼ��
                    If VsfData.TextMatrix(lngCount, mlngSigner) = "" Then
                        VsfData.Cell(flexcpPicture, lngCount, 0) = Nothing
                    Else
                        If InStr(1, VsfData.TextMatrix(lngCount, mlngSigner), "/") <> 0 Then
                            VsfData.Cell(flexcpPicture, lngCount, 0) = imgRow.ListImages(��ǩ).Picture
                        Else
                            VsfData.Cell(flexcpPicture, lngCount, 0) = imgRow.ListImages(ǩ��).Picture
                        End If
                    End If
                End If
            End If
        Next
    End With
    
    Call SetActiveColColor
End Sub

Private Sub zlLableBruit()
    Dim lngScaleLeft As Long, lngScaleTop As Long, lngScaleRight As Long, lngScaleBottom As Long

    Call cbsThis.GetClientRect(lngScaleLeft, lngScaleTop, lngScaleRight, lngScaleBottom)
    VsfData.Move lngScaleLeft + 210, lblTitle.Top + lblTitle.Height + 300, lngScaleRight - lngScaleLeft - 210 * 2
    VsfData.Height = picMain.Height - VsfData.Top
End Sub

Private Sub InitEnv()
    Dim curDate As Date
    Dim intDay As Integer
    Dim rs As New ADODB.Recordset
    Dim blntype As Boolean
    On Error GoTo ErrHand
    
    glngHours = Val(zlDatabase.GetPara("���ݲ�¼ʱ��", glngSys))
    mintChange = Val(zlDatabase.GetPara("���ת������", glngSys, pסԺ��ʿվ, 7))
    '��Ժ����ʱ�䷶Χ
    curDate = zlDatabase.Currentdate
    intDay = Val(zlDatabase.GetPara("��Ժ���˽������", glngSys, pסԺ��ʿվ, 7))
    mdtOutEnd = Format(curDate + intDay, "yyyy-MM-dd 23:59:59")
    intDay = Val(zlDatabase.GetPara("��Ժ���˿�ʼ���", glngSys, pסԺ��ʿվ, 30))
    mdtOutbegin = Format(mdtOutEnd - intDay, "yyyy-MM-dd 00:00:00")
    
    blntype = Val(GetSetting("ZLSOFT", "˽��ģ��\usrTendFileMutilEditor\" & gstrUserName, "Value")) = 0
    If blntype Then
        optLevel(0).Value = True
    Else
        optLevel(1).Value = True
    End If
    
    '���ִ��ڵ����л����¼��Ŀ
    gstrSQL = " Select   ��Ŀ���,upper(��Ŀ����) AS ��Ŀ����,��Ŀ����,��Ŀ����,��Ŀ����,��ĿС��,��Ŀ��ʾ,��Ŀ��λ,��Ŀֵ��,����ȼ�,Ӧ�÷�ʽ,˵��" & _
              " From �����¼��Ŀ B" & _
              " Order by ��Ŀ���"
    Set mrsItems = zlDatabase.OpenSQLRecord(gstrSQL, "���ִ��ڵ����л����¼��Ŀ")
    
    '��ȡ�����µ��Ͳ���ͼ��Ļ����ļ��嵥
    gstrSQL = " Select  ID,���� FROM �����ļ��б� " & vbNewLine & _
              " WHERE ����=3 AND DECODE(����,-1,0,1,0,1)=1 AND (ͨ�� =1 OR (ͨ��=2 And ID IN (Select �ļ�ID FROM ����Ӧ�ÿ��� Where ����ID=[1])))" & vbNewLine & _
              " ORDER BY ��� "
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����µ���Ļ����ļ��嵥", mlng����ID)
    With rs
        cbo�����ļ���ʽ.Clear
        Do While Not .EOF
            cbo�����ļ���ʽ.AddItem !����
            cbo�����ļ���ʽ.ItemData(cbo�����ļ���ʽ.NewIndex) = !ID
            .MoveNext
        Loop
        If .RecordCount <> 0 Then cbo�����ļ���ʽ.ListIndex = 0
    End With
    
    '��ȡ��ǰ�����µ����п���
    gstrSQL = " Select distinct B.ID,B.����||'-'||B.���� AS ����" & _
              " From �������Ҷ�Ӧ A,���ű� B,������Ա C,��Ա�� D" & _
              " Where A.����ID = b.ID And A.����ID=C.����ID And C.��ԱID=D.ID And A.����ID = [1]" & _
              IIf(InStr(1, mstrPrivs, "��ǰ����") <> 0, "", " And D.ID=[2]") & _
              " Order by B.����||'-'||B.����"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��ǰ�����µ����п���", mlng����ID, glngUserId)
    With cbo����
        .Clear
        If InStr(1, mstrPrivs, "��ǰ����") <> 0 Then
            .AddItem "���п���"
            .ItemData(.NewIndex) = -1
        End If
        Do While Not rs.EOF
            .AddItem rs!����
            .ItemData(.NewIndex) = rs!ID
            rs.MoveNext
        Loop
        If rs.RecordCount <> 0 Then .ListIndex = 0
    End With
    
    '��ȡ�󶨵Ĳ�Ѫѹ��Ŀ
    mstrBPItem = ""
    gstrSQL = "Select a.Xh ��Ŀ" & vbNewLine & _
        " From ������������ʽ p, Xmltable('/ITEMLIST/ITEM/XH' Passing p.������Ŀ Columns Xh Varchar2(256) Path '/XH') a" & vbNewLine & _
        " Where p.����id = [1] And ���� = '��Ѫѹ�б�'"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "������������ʽ", mlng����ID)
    Do While Not rs.EOF
        If Not IsNull(rs!��Ŀ) Then
            mstrBPItem = mstrBPItem & "," & Val(rs!��Ŀ)
        End If
        rs.MoveNext
    Loop
    mstrBPItem = Mid(mstrBPItem, 2)
    
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
        
        '��������ڲ�������(�����ڶ�ȡ���ݺ��ʱ���ӵ�,��ʱֻ��Ԥ������)
        mlngSignLevel = VsfData.Cols + cHideCols + VsfData.FixedCols '����������
        mlngSigner = mlngSignLevel + 1
        mlngRecord = mlngSigner + 1
        mlngRowCount = mlngRecord + 1
        mlngRowCurrent = mlngRowCount + 1
        mlngActiveTime = mlngRowCurrent + 1
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

Public Function SignMe() As Boolean
    Dim blnSign As Boolean          '�Ƿ�ǩ���ɹ�
    Dim blnRefresh As Boolean
    Dim strTime As String
    Dim strSignTime As String       '��֤����ǩ����ǩ��ʱ��һ��,����ȡ��ǩ��ʱ��ǩ��ʱ��ͳһȡ��
    Dim str״̬ As String           '����ǩ��ѡ��,����ѭ��ǩ��ʱ��ͣ�ĵ���ǩ������
    Dim str�д��� As String
    Dim str���� As String
    Dim intRow As Integer, intRows As Integer
    On Error GoTo ErrHand
    
    '��ǩ:�Ե�ǰ������������ݽ���ǩ��
    '׼��ǩ��
    strSignTime = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    intRows = VsfData.Rows - 1
    For intRow = VsfData.FixedRows To intRows
        If Val(VsfData.TextMatrix(intRow, mlngRecord)) > 0 And VsfData.TextMatrix(intRow, mlngSigner) = "" Then
            str�д��� = ""
'            If InStr(1, VsfData.TextMatrix(intRow, mlngDate), "/") <> 0 Then
'                strTime = Format(Now, "yyyy") & "-" & ToStandDate(VsfData.TextMatrix(intRow, mlngDate)) & " " & VsfData.TextMatrix(intRow, mlngTime) & ":00"
'            Else
'                strTime = VsfData.TextMatrix(intRow, mlngDate) & " " & VsfData.TextMatrix(intRow, mlngTime) & ":00"
'            End If
            strTime = VsfData.TextMatrix(intRow, mlngActiveTime)
            
            blnSign = SignName(intRow, strTime, strSignTime, str״̬, str�д���)
            If Not blnSign Then Exit For
            If Not blnRefresh Then blnRefresh = blnSign
            If str�д��� <> "" Then
                str���� = str���� & vbCrLf & "����ʱ��=[" & strTime & "]" & str�д���
            End If
        End If
    Next
    
    SignMe = blnRefresh
    mblnSigned = blnRefresh
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub UnSignMe()
    Dim lngRecord As Long
    Dim blnOK As Boolean
    Dim strTime As String
    Dim blnTrans As Boolean
    Dim lngRow As Long, lngRows As Long
    Dim clsSign As Object
    On Error GoTo ErrHand
    '�������һ���Ǳ��˵�ǩ�������ݵ�ǰѡ�����ݵ�ǩ��ʱ�䣬����ȡ��ǩ��

    gcnOracle.BeginTrans
    blnTrans = True
    lngRows = VsfData.Rows - 1
    For lngRow = VsfData.FixedRows To lngRows
        If Val(VsfData.TextMatrix(lngRow, mlngRecord)) > 0 And VsfData.TextMatrix(lngRow, mlngSigner) <> "" Then
            If Val(VsfData.TextMatrix(lngRow, mlngSignLevel)) > 0 Then
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
                        RaiseEvent AfterRowColChange("����ǩ������δ����ȷ��װ�����˲������ܼ�����", True)
                        Exit Sub
                    End If
                End If
            End If
        End If
        
        If Val(VsfData.TextMatrix(lngRow, mlngRecord)) > 0 Then
            '��ȡ����ʱ��
'            If InStr(1, VsfData.TextMatrix(lngRow, mlngDate), "/") <> 0 Then
'                strTime = Format(Now, "yyyy") & "-" & ToStandDate(VsfData.TextMatrix(lngRow, mlngDate)) & " " & VsfData.TextMatrix(lngRow, mlngTime) & ":00"
'            Else
'                strTime = VsfData.TextMatrix(lngRow, mlngDate) & " " & VsfData.TextMatrix(lngRow, mlngTime) & ":00"
'            End If
            strTime = VsfData.TextMatrix(lngRow, mlngActiveTime)
            'ȡ��ǩ��
            gstrSQL = "ZL_���˻�������_UNSIGNNAME("
            gstrSQL = gstrSQL & VsfData.TextMatrix(lngRow, c�ļ�ID) & ","
            gstrSQL = gstrSQL & "To_Date('" & strTime & "','yyyy-MM-dd hh24:mi:ss'),1)"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "ִ��ȡ��ǩ��")
            '����ͼ��
            VsfData.Cell(flexcpPicture, lngRow, 0) = Nothing
            If mlngSignName <> -1 Then VsfData.TextMatrix(lngRow, mlngSignName) = ""
            VsfData.TextMatrix(lngRow, mlngSignLevel) = 0
            VsfData.TextMatrix(lngRow, mlngSigner) = ""
            If mlngSignTime <> -1 Then VsfData.TextMatrix(lngRow, mlngSignTime) = ""
            '--58414,������,2013-04-10,��ӻ�ʿ��ǩ������ʾģʽ
            Call SingerShowType(VsfData, lngRow, lngRow + Val(VsfData.TextMatrix(lngRow, mlngRowCount)) - 1, True)
        End If
    Next
    gcnOracle.CommitTrans
    blnTrans = False
    mblnSigned = False
    Exit Sub
ErrHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function SignName(ByVal intRow As Integer, ByVal strStart As String, ByVal strSignTime As String, _
    str״̬ As String, Optional str���� As String) As Boolean
    '******************************************************************************************************************
    '����:
    '
    '
    '******************************************************************************************************************
    Dim oSign As cTendSign
    Dim strSource As String             '��ǩԴ���ݴ�
    Dim lngLoop As Long
    Dim rs As New ADODB.Recordset

    On Error GoTo ErrHand

    '��ʼ����
    '------------------------------------------------------------------------------------------------------------------
    strSource = ""

    '��ȡҪǩ��������
    '------------------------------------------------------------------------------------------------------------------
   gstrSQL = " Select  a.��¼����,a.��Ŀ����,a.��Ŀ���,a.��Ŀ����,a.��Ŀ����,a.��¼����,a.��Ŀ��λ,a.��¼���,a.���²�λ,a.��¼���,a.���Ժϸ�,a.δ��˵��,a.��¼��,a.��¼ʱ��  " & _
              " From ���˻�����ϸ a,���˻������� b,���˻����ļ� c " & _
              " Where a.��¼id=b.ID And B.�������=0 AND MOD(A.��¼����,10)<>5 And b.�ļ�ID=c.ID And a.��ֹ�汾 Is Null And C.ID=[1] And b.����ʱ��=[2]" & _
              " Order by a.��Ŀ���"
    Call SQLDIY(gstrSQL)
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҪǩ��������", Val(VsfData.TextMatrix(intRow, c�ļ�ID)), CDate(strStart))
    If rs.BOF = False Then
        Do While Not rs.EOF
            For lngLoop = 0 To rs.Fields.Count - 1
                strSource = strSource & CStr(zlCommFun.NVL(rs.Fields(lngLoop).Value, ""))
            Next
            rs.MoveNext
        Loop
    End If
    If strSource = "" Then
        RaiseEvent AfterRowColChange("��ǰû����Ҫǩ������Ϣ��", True)
        Exit Function
    End If

    '------------------------------------------------------------------------------------------------------------------
    On Error Resume Next
    Err = 0
    '76223:������,2012-09-13,����ǩ�����ʱ�����Ϣ
    Set oSign = frmTendFileSign.ShowMe(Me, mstrPrivs, Val(VsfData.TextMatrix(intRow, c�ļ�ID)), mlng����ID, δ����, strSource, False, str״̬, str����)
    On Error GoTo ErrHand

    If Not oSign Is Nothing Then
        gstrSQL = "ZL_���˻�������_SIGNNAME("
        gstrSQL = gstrSQL & Val(VsfData.TextMatrix(intRow, c�ļ�ID)) & ","
        gstrSQL = gstrSQL & "To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss'),0,"
        gstrSQL = gstrSQL & "'" & oSign.���� & "',"
        gstrSQL = gstrSQL & "'" & oSign.ǩ����Ϣ & "'," & oSign.ǩ������ & ","
        gstrSQL = gstrSQL & oSign.֤��ID & ","
        gstrSQL = gstrSQL & oSign.ǩ����ʽ & ",'" & oSign.ʱ��� & "',0,'" & oSign.ʱ�����Ϣ & "',To_Date('" & strSignTime & "','yyyy-mm-dd hh24:mi:ss'))"
        
        Call zlDatabase.ExecuteProcedure(gstrSQL, "ִ��ǩ��")
        SignName = True
        
        VsfData.TextMatrix(intRow, mlngSignLevel) = oSign.֤��ID
        VsfData.TextMatrix(intRow, mlngSigner) = "SignName"
        '����ͼ��
        VsfData.Cell(flexcpPicture, intRow, 0) = imgRow.ListImages(ǩ��).Picture
        If mlngSignName <> -1 Then VsfData.TextMatrix(intRow, mlngSignName) = gstrUserName
        If mlngSignTime <> -1 Then VsfData.TextMatrix(intRow, mlngSignTime) = oSign.ʱ���
        '--58414,������,2013-04-10,��ӻ�ʿ��ǩ������ʾģʽ
        Call SingerShowType(VsfData, intRow, intRow + Val(VsfData.TextMatrix(intRow, mlngRowCount)) - 1)
    End If

    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function CancelMe() As Boolean
    CancelMe = True
    mblnChange = False
    mintType = -1
    
    '�ڴ��¼�����
    mrsCellMap.Filter = 0
    If mrsCellMap.RecordCount <> 0 Then mrsCellMap.MoveFirst
    Do While True
        If mrsCellMap.EOF Then Exit Do
        mrsCellMap.Delete
        mrsCellMap.Update
        mrsCellMap.MoveNext
    Loop
    
    Call DataMap_Restore
    
    Call InitCons
End Function

Public Function SaveME() As Boolean
    If Not CheckData Then Exit Function
    If Not SaveData Then Exit Function

    mblnShow = False
    Call InitCons
    SaveME = True
    RaiseEvent AfterRowColChange("����ɹ���", False)
    
    '--48659:������,2012-09-14,����ֶ�'˵��'
    RaiseEvent ShowTipInfo(VsfData, "", True)
    
    If VsfData.ROW < VsfData.Rows And mlngDate < VsfData.Cols Then
        VsfData.Select VsfData.ROW, mlngDate
    End If
End Function

Public Function ShowMe(ByVal frmParent As Form, ByVal lngDeptID As Long, Optional ByVal strPrivs As String, Optional ByVal bytSize As Byte = 0) As Boolean
    '******************************************************************************************************************
    '���ܣ� ��ʾ�����¼�ļ�����
    '������ frmParent           �ϼ��������
    '       lngDeptID           Ҫ��ʾ�����¼�Ŀ���
    '���أ� ��
    '******************************************************************************************************************
    Dim lngRow As Long
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    Err = 0

    mblnInit = False
    mblnHistory = False
    
    mintҳ�� = 1
    mlng����ID = lngDeptID
    mstrPrivs = strPrivs
    mblnBlowup = (bytSize = 1) '(zlDatabase.GetPara("�����ļ���ʾģʽ", glngSys, 1255, 0) = 1)
    Set mfrmParent = frmParent

    mintPreDays = Val(zlDatabase.GetPara("����¼�뻤����������", glngSys, 1255, "1"))
    mstrMaxDate = Format(DateAdd("d", mintPreDays, zlDatabase.Currentdate), "yyyy-MM-dd HH:mm")
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
    
     '--48659:������,2012-09-14,����ֶ�'˵��'
    RaiseEvent ShowTipInfo(VsfData, "", True)
    
    If cbo����.ListCount = 0 Then
        MsgBox "�������ڵ�ǰ�������κο��ң�����ʹ�øù��ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    Call ReSetFontSize
    ShowMe = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

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
    bytFontSize = IIf(mblnBlowup = True, 12, 9)
    
    UserControl.FontSize = bytFontSize
    UserControl.FontName = "����"
    For Each objCtrl In UserControl.Controls
        Select Case UCase(TypeName(objCtrl))
        Case UCase("Label")
            Select Case UCase(objCtrl.Name)
            Case UCase("lbl�ļ���ʽ"), UCase("lbl����")
                objCtrl.FontSize = bytFontSize
                objCtrl.Height = TextHeight("��") + 20
            End Select
        Case UCase("ComboBox")
            Select Case UCase(objCtrl.Name)
            Case UCase("cbo�����ļ���ʽ"), UCase("cbo����")
                objCtrl.FontSize = bytFontSize
            End Select
        Case UCase("CheckBox")
            Select Case UCase(objCtrl.Name)
            Case UCase("chk����"), UCase("chk��Ժ")
                objCtrl.FontSize = bytFontSize
                objCtrl.Width = TextWidth("����" & objCtrl.Caption) - TextWidth("��") / 3
            End Select
        Case UCase("CommandBars")
            Set CtlFont = objCtrl.Options.Font
            If CtlFont Is Nothing Then
                Set CtlFont = UserControl.Font
            End If
            CtlFont.Size = bytFontSize
            Set objCtrl.Options.Font = CtlFont
        Case UCase("CommandButton")
            If UCase(objCtrl.Name) = UCase("cmdˢ��") Then
                objCtrl.FontSize = bytFontSize
                objCtrl.Width = TextWidth(" " & IIf(objCtrl.Caption = "", "  ", objCtrl.Caption) & " ")
            End If
        End Select
    Next
    
    '�ƶ��ؼ�λ��
    cbo�����ļ���ʽ.Top = (pic��������.Height - cbo�����ļ���ʽ.Height) \ 2
    lbl�ļ���ʽ.Left = 60
    cbo�����ļ���ʽ.Left = lbl�ļ���ʽ.Left + lbl�ļ���ʽ.Width + TextWidth("��") / 2
    lbl�ļ���ʽ.Top = cbo�����ļ���ʽ.Top + (cbo�����ļ���ʽ.Height - lbl�ļ���ʽ.Height) \ 2
    lbl����.Left = cbo�����ļ���ʽ.Left + cbo�����ļ���ʽ.Width + TextWidth("��")
    lbl����.Top = lbl�ļ���ʽ.Top
    cbo����.Left = lbl����.Left + lbl����.Width + TextWidth("��") / 2
    cbo����.Top = cbo�����ļ���ʽ.Top
    chk����.Left = cbo����.Left + cbo����.Width + TextWidth("��")
    chk����.Top = lbl�ļ���ʽ.Top
    chk��Ժ.Left = chk����.Left + chk����.Width + TextWidth("��") / 2
    chk��Ժ.Top = chk����.Top
    cmdˢ��.Height = cbo����.Height + 15
    cmdˢ��.Left = chk��Ժ.Left + chk��Ժ.Width + TextWidth("��")
    lblEntry.Left = cmdˢ��.Left + cmdˢ��.Width + TextWidth("��")
    lblEntry.Top = cmdˢ��.Top + (cmdˢ��.Height - lblEntry.Height) \ 2
    lblEntry.Height = cmdˢ��.Height
    optLevel(0).Left = lblEntry.Left + lblEntry.Width + TextWidth("��") / 2
    optLevel(0).Top = cmdˢ��.Top + 10
    optLevel(0).Height = cmdˢ��.Height
    optLevel(1).Left = optLevel(0).Left + optLevel(0).Width + TextWidth("��")
    optLevel(1).Top = optLevel(0).Top
    optLevel(1).Height = optLevel(0).Height
    
    
    pic��������.Width = optLevel(1).Left + optLevel(1).Width + 50
End Sub

Private Function CheckFlip() As Boolean
    Dim blnExit As Boolean, blnNULL As Boolean
    Dim lngRow As Long, lngCol As Long, lngRows As Long, lngCols As Long
    Dim lngOldRow As Long, lngOldCol As Long, lngEditCol As Long, blnShow As Boolean
    Dim strDate As String, strInfo As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHand
    
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
    End Select
    
    cmdWord.Visible = False
    mintType = -1
    mblnShow = False
    
    lngRows = VsfData.Rows - 1
    lngCols = VsfData.Cols - 1
    For lngRow = VsfData.FixedRows To lngRows
        If Val(VsfData.TextMatrix(lngRow, mlngDemo)) > 1 And Trim(VsfData.TextMatrix(lngRow, mlngSigner)) = "" And VsfData.RowHidden(lngRow) = False Then
            blnNULL = True
            For lngCol = mlngTime + 1 To lngCols - 1
                If Not VsfData.ColHidden(lngCol) And lngCol < mlngNoEditor And ISEditAssistant(lngCol) = False Then
                    If VsfData.TextMatrix(lngRow, lngCol) <> "" And Not (IsDiagonal(lngCol) And InStr(1, VsfData.TextMatrix(lngRow, lngCol), "/") <> 0) Then
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
    
    'ҳ���л�ǰ��飺����ʱ����ȷ����������������ڱ���ʱ�Ͳ����ټ������ҳ��������ˣ�����������¼��ʱ�Ѿ������˼�飬�˴��Թ���
    lngRows = VsfData.Rows - 1
    lngCols = VsfData.Cols - 1
    For lngRow = VsfData.FixedRows To lngRows
        mrsCellMap.Filter = "ҳ��=" & mintҳ�� & " And �к�=" & lngRow & " And �к�>" & mlngTime
        If mrsCellMap.RecordCount = 0 And Val(VsfData.TextMatrix(lngRow, mlngRecord)) > 0 Then
            mrsCellMap.Filter = "ҳ��=" & mintҳ�� & " And �к�=" & lngRow & " And �к�>=" & mlngDate
        End If
        If mrsCellMap.RecordCount <> 0 Then
            If Not VsfData.RowHidden(lngRow) And Val(VsfData.TextMatrix(lngRow, mlngDemo)) <= 1 Then
                blnExit = (VsfData.TextMatrix(lngRow, mlngDate) = "" Or VsfData.TextMatrix(lngRow, mlngTime) = "")
                If blnExit Then
                    mrsCellMap.Filter = 0
                    If VsfData.TextMatrix(lngRow, mlngDate) = "" Then
                        lngCol = mlngDate
                    Else
                        lngCol = mlngTime
                    End If
                    VsfData.ROW = lngRow: VsfData.COL = lngCol
                    mblnShow = True: Call VsfData_EnterCell
                    If Not VsfData.RowIsVisible(lngRow) Then VsfData.TopRow = lngRow
                    CheckFlip = False
                    RaiseEvent AfterRowColChange("�벹������ʱ�䣡", True)
                    Exit Function
                Else
                    '���ڲ�Ϊ�ս�������ڵĺϷ���
                    If Val(VsfData.TextMatrix(lngRow, mlngRecord)) > 0 Then
                        If mblnDateAd Then
                            strDate = Mid(VsfData.TextMatrix(lngRow, mlngActiveTime), 1, 4) & "-" & ToStandDate(VsfData.TextMatrix(lngRow, mlngDate)) & " " & VsfData.TextMatrix(lngRow, mlngTime)
                        Else
                            strDate = VsfData.TextMatrix(lngRow, mlngDate) & " " & VsfData.TextMatrix(lngRow, mlngTime)
                        End If
                        strDate = Format(strDate, "YYYY-MM-DD HH:mm")
                        blnExit = (strDate = Format(VsfData.TextMatrix(lngRow, mlngActiveTime), "YYYY-MM-DD HH:mm"))
                    End If
                    If blnExit = False Then
                        VsfData.ROW = lngRow: VsfData.COL = mlngTime
                        If Not CheckDateTime(VsfData.TextMatrix(VsfData.ROW, VsfData.COL), strInfo) Then
                            mrsCellMap.Filter = ""
                            mblnShow = True: Call VsfData_EnterCell
                            If Not VsfData.RowIsVisible(lngRow) Then VsfData.TopRow = lngRow
                            RaiseEvent AfterRowColChange(strInfo, True)
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
            If Val(FormatValue(VsfData.TextMatrix(lngRow, mlngDemo))) = 1 And Val(FormatValue(VsfData.TextMatrix(lngRow, mlngRecord))) > 0 Then
                If mblnDateAd Then
                    strDate = Mid(VsfData.TextMatrix(lngRow, mlngActiveTime), 1, 4) & "-" & ToStandDate(VsfData.TextMatrix(lngRow, mlngDate)) & " " & VsfData.TextMatrix(lngRow, mlngTime)
                Else
                    strDate = VsfData.TextMatrix(lngRow, mlngDate) & " " & VsfData.TextMatrix(lngRow, mlngTime)
                End If
                strDate = Format(strDate, "YYYY-MM-DD HH:mm")
            End If
            
            If IsDate(strDate) And Not VsfData.RowHidden(lngRow) And Val(FormatValue(VsfData.TextMatrix(lngRow, mlngDemo))) > 1 And _
                 Val(FormatValue(VsfData.TextMatrix(lngRow, mlngRecord))) <= 0 Then
                mrsCellMap.Filter = "ҳ��=" & mintҳ�� & " And �к�=" & lngRow & " And �к�>" & mlngTime
                If mrsCellMap.RecordCount > 0 Then
                    lngEditCol = 0
                    If CheckCollectIsData(lngRow, 1, lngEditCol) = True Then
                        If ISCollectSigned(Val(VsfData.TextMatrix(lngRow, c�ļ�ID)), Format(strDate, "YYYY-MM-DD"), Format(strDate, "HH:MM")) Then
                            VsfData.ROW = lngRow: VsfData.COL = lngEditCol
                            strInfo = "�������ķ�����������Ӧ�Ļ�����������ǩ����������������µĻ��������ݣ�"
                            mrsCellMap.Filter = ""
                            mblnShow = True: Call VsfData_EnterCell
                            If Not VsfData.RowIsVisible(lngRow) Then VsfData.TopRow = lngRow
                            RaiseEvent AfterRowColChange(strInfo, True)
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
    mrsCellMap.Filter = 0
    CheckFlip = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    mrsCellMap.Filter = 0
End Function

Private Function CheckData() As Boolean
    Dim intLevel As Integer
    Dim lngPage As Long
    On Error GoTo ErrHand
    '�������

    '����޸������ݶ�����ʱ�䲻ȫ����ʾ�����ݺϷ�����¼��ʱ�Ѿ���飩
'    Call OutputRsData(mrsCellMap)
'    Call OutputRsData(mrsDataMap)
    If Not CheckFlip Then Exit Function

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
    Dim intPos As Integer, intMax As Integer, intPage As Integer, intRow As Integer, intUsedRows As Integer
    Dim strReturn As String, strCellData As String, strPart As String
    Dim strMonth As String, strDay As String
    Dim strDate As String, strTime As String, strTemp As String
    Dim strDatetime As String, strCurrDate As String, strDays As String
    Dim strSaveRows As String
    Dim str����ʱ��_L As String, str����ʱ�� As String, str�ļ�ID As String
    Dim lngLastRow As Long

    ReDim Preserve strSQL(1 To 1)
    ReDim Preserve strSQLTime(1 To 1) '����ʱ��䶯SQL����
    ReDim Preserve strCollectSQL(1 To 1) 'С��SQL����
    
    Dim rsTemp As New ADODB.Recordset, rsTime As New ADODB.Recordset, rsTimeCur As New ADODB.Recordset
    Dim strFileds As String, strValues As String
    On Error GoTo ErrHand
    
    strFileds = "ID," & adDouble & ",18|�ļ�ID," & adDouble & ",18|ʱ��," & adDate & ",20|����ʱ��," & adDate & ",20|���," & adInteger & ",1"
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
            If intRow <> !�к� Then
endWork:
                If intRow > 0 Then
                    blnDel = VsfData.RowHidden(intRow)
                    intUsedRows = Val(Split(VsfData.TextMatrix(intRow, mlngRowCount), "|")(0))
                End If

                If blnSaved Then
                    strSaveRows = strSaveRows & "," & intRow
                     
                    '��ɴ�ӡ���ݽ���
'                    �ļ�ID_IN IN ���˻����ӡ.�ļ�ID%TYPE,
'                    ����ʱ��_IN IN ���˻����ӡ.����ʱ��%TYPE,
'                    ����_IN IN ���˻����ӡ.����%TYPE,
'                    ɾ��_IN Number:=0
                    gstrSQL = "ZL_���˻����ӡ_UPDATE(" & Val(VsfData.TextMatrix(intRow, c�ļ�ID)) & ",to_date('" & strDatetime & "','yyyy-MM-dd hh24:mi:ss')," & intUsedRows & "," & IIf(blnDel, "1", "0") & ")"
                    strSQL(ReDimArray(strSQL)) = gstrSQL

                    'ֻҪ�޸Ĺ�����,��Ȼ��ִ�д�ӡ����,�����������л������ڵĴ���
                    strTemp = Format(DateAdd("d", -1, CDate(strDatetime)), "yyyy-MM-dd")
                    If InStr(1, "," & strDays & ",", "," & strTemp & ",") = 0 Then
                        strDays = strDays & "," & strTemp
                        gstrSQL = "ZL_��������_UPDATE(" & Val(VsfData.TextMatrix(intRow, c�ļ�ID)) & ",'" & strTemp & "')"
                        strCollectSQL(ReDimArray(strCollectSQL)) = gstrSQL
                    End If
                    
                    If InStr(1, "," & strDays & ",", "," & Mid(strDatetime, 1, 10) & ",") = 0 Then
                        'ͬ����������Ļ���(ҹ��,ȫ����ܿ���Ĵ���)
                        strDays = strDays & "," & Mid(strDatetime, 1, 10)
                        gstrSQL = "ZL_��������_UPDATE(" & Val(VsfData.TextMatrix(intRow, c�ļ�ID)) & ",'" & Mid(strDatetime, 1, 10) & "')"
                        strCollectSQL(ReDimArray(strCollectSQL)) = gstrSQL
                    End If
                        
                    strTemp = Format(DateAdd("d", 1, CDate(strDatetime)), "yyyy-MM-dd")
                    If InStr(1, "," & strDays & ",", "," & strTemp & ",") = 0 Then
                        strDays = strDays & "," & strTemp
                        gstrSQL = "ZL_��������_UPDATE(" & Val(VsfData.TextMatrix(intRow, c�ļ�ID)) & ",'" & strTemp & "')"
                        strCollectSQL(ReDimArray(strCollectSQL)) = gstrSQL
                    End If

                    blnSaved = False
                    If .EOF Then Exit Do
                End If

                '����ֵ
                intPage = !ҳ��
                intRow = !�к�
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
                    gstrSQL = "ZL_���˻�������_COLLECT(" & Val(VsfData.TextMatrix(intRow, c�ļ�ID)) & ",to_date('" & arrCollect(3) & "','yyyy-MM-dd hh24:mi:ss')," & _
                            Val(arrCollect(1)) & ",'" & arrCollect(0) & "'," & Val(arrCollect(2)) & "," & !ɾ�� & ")"
                    strSQL(ReDimArray(strSQL)) = gstrSQL
                    blnSaved = True
                Else
                    strDate = NVL(!����)
                    If strDate <> "" Then
                        If mblnDateAd Then
                            strDate = Mid(strCurrDate, 1, 5) & ToStandDate(strDate)
                            '����Ƿ����༭֮ǰ��ʱ��(һ���µ�����)
                            If CDate(strDate) > CDate(Mid(strCurrDate, 1, 10)) And Mid(strCurrDate, 6, 2) = "01" And Mid(strDate, 6, 2) = "12" Then
                                strDate = DateAdd("yyyy", -1, CDate(strDate))
                            End If
                        Else
                            strDate = Format(strDate, "yyyy-MM-dd")
                        End If
                    End If
                End If
            ElseIf !�к� = mlngTime Then
                strTime = NVL(!����)
                If strDatetime = "" Then
                    If strDate = "" Then strDate = Mid(strCurrDate, 1, 10)
                    strDatetime = strDate & " " & strTime & ":00"
                End If
                '����������ݣ�����ʱ����ͨ����������ֻ������+
                If Val(NVL(!��λ)) >= 1 Then
                    'strDatetime = Mid(strDatetime, 1, 17) & String(2 - Len(!��λ), "0") & Val(!��λ) - 1
                    strDatetime = DateAdd("S", Val(!��λ) - 1, CDate(strDatetime))
                End If
                If lngRecord <> 0 Then
                    '���·���ʱ��
'                    gstrSQL = "Zl_���˻�������_����ʱ��(" & lngRecord & ",to_date('" & strDatetime & "','yyyy-MM-dd hh24:mi:ss'))"
'                    strSQL(ReDimArray(strSQL)) = gstrSQL
                    strValues = lngRecord & "|" & Val(VsfData.TextMatrix(intRow, c�ļ�ID)) & "|" & Format(strDatetime, "YYYY-MM-DD HH:mm:ss") & "|" & Format(VsfData.TextMatrix(intRow, mlngActiveTime), "YYYY-MM-DD HH:mm:ss") & "|0"
                    Call Record_Update(rsTime, "ID|�ļ�ID|ʱ��|����ʱ��|���", strValues, "ID|" & lngRecord)
                    
                    Call Record_Update(rsTimeCur, "ID|�ļ�ID|ʱ��|����ʱ��|���", strValues, "ID|" & lngRecord)
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
                        If Not (Val(VsfData.TextMatrix(intRow, mlngRecord)) = 0 And arrValue(intPos) = "") Then
    '                    �ļ�ID_IN IN ���˻�������.�ļ�ID%TYPE,
    '                    ����ʱ��_IN IN ���˻�������.����ʱ��%TYPE,
    '                    ��¼����_IN IN ���˻�����ϸ.��¼����%TYPE,          --������Ŀ=1���ϱ�˵��=2�������ձ��=4��ǩ����¼=5���±�˵��=6�����������=9
    '                    ��Ŀ���_IN IN ���˻�����ϸ.��Ŀ���%TYPE,          --������Ŀ����ţ��ǻ�����Ŀ�̶�Ϊ0
    '                    ��¼����_IN IN ���˻�����ϸ.��¼����%TYPE := NULL,  --��¼���ݣ��������Ϊ�գ��������ǰ�����ݣ�37��38/37
    '                    ���²�λ_IN IN ���˻�����ϸ.���²�λ%TYPE := NULL,
    '                    ���˼�¼_IN IN NUMBER := 1,
                        gstrSQL = "ZL_���˻�������_UPDATE(" & Val(VsfData.TextMatrix(intRow, c�ļ�ID)) & ",to_date('" & strDatetime & "','yyyy-MM-dd hh24:mi:ss'),1," & _
                                arrOrder(intPos) & ",'" & arrValue(intPos) & "','" & arrPart(intPos) & "'," & intAllow & ",0,0,NULL,NULL,NULL,'" & NVL(!���) & "')"
                        strSQL(ReDimArray(strSQL)) = gstrSQL
                        blnSaved = True
                        End If
                    Next
                    mrsItems.Filter = 0
                End If
            End If

            .MoveNext
        Loop

        If blnSaved Then GoTo endWork
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
                gstrSQL = UpdateTime(rsTimeCur, Format(rsTimeCur!ʱ��, "YYYY-MM-DD HH:mm:ss"), lngRecord, Val(rsTime!�ļ�ID))
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
    On Error Resume Next

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
    
    intMax = UBound(strSQL)
    If intMax > 0 Then
'        objStream.WriteLine (Now & "׼����������")
        For intPos = 1 To intMax
            If strSQL(intPos) <> "" Then
    '            objStream.WriteLine (Now & "��SQL��" & strSQL(intPos))
                Call zlDatabase.ExecuteProcedure(strSQL(intPos), "���滤���¼������")
            End If
        Next
    '    objStream.WriteLine (Now & "�����������")
    End If
    
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
    
    gcnOracle.CommitTrans
    SaveData = True
    blnTrans = False
    mblnSaved = True
    mblnChange = False
    
    '���������еļ�¼ID��,��ʾ�������ѱ���
    strSaveRows = strSaveRows & ","
    For intRow = VsfData.FixedRows To VsfData.Rows - 1
        If InStr(1, strSaveRows, "," & intRow & ",") <> 0 Then
            strDatetime = ""
            If Val(VsfData.TextMatrix(intRow, mlngDemo)) > 0 Then
                If CheckGroupDate(intRow) = True Then
                    '�������޸ĲŽ�������̣�ȡ������¼��ʵ��ʱ��
                    If mblnDateAd Then
                        strDate = Format(VsfData.TextMatrix(intRow, mlngActiveTime), "DD") & "/" & Format(VsfData.TextMatrix(intRow, mlngActiveTime), "MM")
                    Else
                        strDate = Mid(VsfData.TextMatrix(intRow, mlngActiveTime), 1, 10)
                    End If
                    strTime = Mid(VsfData.TextMatrix(intRow, mlngActiveTime), 12, 5)
                Else
                    strDate = VsfData.TextMatrix(intRow - Val(VsfData.TextMatrix(intRow, mlngDemo)) + 1, mlngDate)
                    strTime = VsfData.TextMatrix(intRow - Val(VsfData.TextMatrix(intRow, mlngDemo)) + 1, mlngTime)
                End If
            Else
                '��ͨ����
                strDate = VsfData.TextMatrix(intRow, mlngDate)
                strTime = VsfData.TextMatrix(intRow, mlngTime)
            End If

            If strDate <> "" Then
                If mblnDateAd Then
                    strDate = Mid(strCurrDate, 1, 5) & ToStandDate(strDate)
                Else
                    strDate = Format(strDate, "yyyy-MM-dd")
                End If
                strDatetime = strDate & " " & strTime & ":00"
                If Val(VsfData.TextMatrix(intRow, mlngDemo)) >= 1 Then
                    strDatetime = Mid(strDatetime, 1, 17) & String(2 - Len(VsfData.TextMatrix(intRow, mlngDemo)), "0") & Val(VsfData.TextMatrix(intRow, mlngDemo)) - 1
                End If
            End If
            
            If strDatetime <> "" Then
                gstrSQL = " Select A.ID,A.����ʱ��,A.������ From ���˻������� A,���˻����ļ� B" & vbNewLine & _
                          " Where A.�ļ�ID=B.ID And B.ID=[1] And A.����ʱ��=[2]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��¼ID", Val(VsfData.TextMatrix(intRow, c�ļ�ID)), CDate(strDatetime))
                If rsTemp.RecordCount <> 0 Then
                    VsfData.TextMatrix(intRow, mlngRecord) = rsTemp!ID
                    VsfData.TextMatrix(intRow, mlngActiveTime) = Format(rsTemp!����ʱ��, "YYYY-MM-DD HH:mm:ss")
                    If mlngOperator <> -1 Then VsfData.TextMatrix(intRow, mlngOperator) = NVL(rsTemp!������)
                    '--58414,������,2013-04-10,��ӻ�ʿ��ǩ������ʾģʽ
                    Call SingerShowType(VsfData, intRow, intRow + Val(VsfData.TextMatrix(intRow, mlngRowCount)) - 1)
                End If
            End If
        End If
    Next
    
    '94211,2016-4-14,���� ����¼��
    ReDim Preserve strCollectSQL(1 To 1)
    For intRow = VsfData.FixedRows To VsfData.Rows - 1
        If VsfData.RowHidden(intRow) = True Then
            strCollectSQL(ReDimArray(strCollectSQL)) = intRow
        End If
    Next
    
    For intRow = 1 To UBound(strCollectSQL)
        If strCollectSQL(intRow) <> "" Then
            VsfData.RowPosition(Val(strCollectSQL(intRow))) = VsfData.Rows - 1
        End If
    Next
    
    str�ļ�ID = ""
    For intRow = VsfData.FixedRows To VsfData.Rows - 1
        
        If VsfData.TextMatrix(intRow, mlngRowCount) Like "*|1" And VsfData.RowHidden(intRow) = False Then
            If str�ļ�ID <> VsfData.TextMatrix(intRow, c�ļ�ID) Then str����ʱ��_L = ""
            If VsfData.TextMatrix(intRow, mlngActiveTime) <> "" Then str����ʱ�� = VsfData.TextMatrix(intRow, mlngActiveTime)
        
            If VsfData.TextMatrix(intRow, c�ļ�ID) = str�ļ�ID And str����ʱ��_L <> "" And Mid(str����ʱ��_L, 1, 16) = Mid(str����ʱ��, 1, 16) And str����ʱ��_L <> str����ʱ�� Then
                '������ͬ��������ͬ���Ҳ��ǻ��������У���˵����Щ������һ�飬����lngDemo��
                VsfData.TextMatrix(intRow, mlngDemo) = intRow - lngLastRow + 1
                If intRow - lngLastRow = Val(FormatValue(VsfData.TextMatrix(lngLastRow, mlngRowCount))) Then
                    VsfData.TextMatrix(lngLastRow, mlngDemo) = 1
                End If
            Else
                lngLastRow = intRow
                str����ʱ��_L = str����ʱ��
                str�ļ�ID = VsfData.TextMatrix(intRow, c�ļ�ID)
            End If
        End If
    Next
    
    '�ڴ��¼�����
    mrsCellMap.Filter = 0
    If mrsCellMap.RecordCount <> 0 Then mrsCellMap.MoveFirst
    Do While True
        If mrsCellMap.EOF Then Exit Do
        mrsCellMap.Delete
        mrsCellMap.Update
        mrsCellMap.MoveNext
    Loop
    '���浱ǰ����
    Call InitCons
    mblnShow = False
    Call DataMap_Save
    If VsfData.Rows > VsfData.FixedRows Then
        VsfData.ROW = VsfData.FixedRows
    Else
        vsfHistory.Rows = vsfHistory.FixedRows
    End If
    
    Exit Function
ErrHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function UpdateTime(rsTimeCur As ADODB.Recordset, ByVal strTime As String, lngID As Long, lng�ļ�ID As Long) As String
    Dim strSQL As String
    rsTimeCur.Filter = "����ʱ��='" & Format(strTime, "YYYY-MM-DD HH:mm:ss") & "' And ���=0 And ID<>" & lngID & " and  �ļ�ID=" & lng�ļ�ID
    If rsTimeCur.RecordCount > 0 Then
        lngID = Val(rsTimeCur!ID)
        lng�ļ�ID = Val(rsTimeCur!�ļ�ID)
        strSQL = UpdateTime(rsTimeCur, Format(rsTimeCur!ʱ��, "YYYY-MM-DD HH:mm:ss"), lngID, lng�ļ�ID)
    Else
        strSQL = "Zl_���˻�������_����ʱ��(" & lngID & ",to_date('" & strTime & "','yyyy-MM-dd hh24:mi:ss'))"
    End If
    UpdateTime = strSQL
End Function

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strDate As String, strTime As String
    Dim strLockItem As String                   'ͬ������������,�������޸Ļ�ɾ��
    Dim lngTop As Long, lngHeight As Long
    Dim intMax As Integer                       'ͬ������������ռ�õ��������
    Dim intNULL As Integer, lngStartRow As Long, lngRowCount As Long, blnNULL As Boolean
    Dim lngRow As Long, lngCol As Long, lngRows As Long, lngCols As Long
    Dim strKey As String, strField As String, strValue As String
    Dim strPart As String
    Dim lngOrder As Long, intGroupFirstRows As Integer
    Dim lngCol1 As Long, lngRow1 As Long, lngCurRow As Long, strText As String, lngCount As Long
    Dim blnTure As Boolean, blnShow As Boolean
    Dim varAssistant() As Variant, strAssistantCols As String
    Dim strCols As String
    
    On Error GoTo err_exit
    
    Select Case Control.ID
    
    Case conMenu_File_PrintSet

        Call zl9PrintMode.zlPrintSet

    Case conMenu_File_Preview

        Call zlRptPrint(2)

    Case conMenu_File_Print

        Call zlRptPrint(1)

    Case conMenu_File_Excel

        Call zlRptPrint(3)
        
    '���ݷ��飬��������ǰ�ķ���ͱ�����׷�ӷ���
    Case conMenu_Edit_Group_Append
        '��ӷ��飬�ڵ�ǰ��(ֻ��һ�е�������)�������������ӿհ���
        If VsfData.ROW < VsfData.FixedRows Then Exit Sub
        '0)���ݳ�ʼ��
        '��������ʾ��¼��ؼ�
        cmdWord.Visible = False
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
        End Select
        blnShow = mblnShow
        mintType = -1: mblnShow = False
        
        If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then
            lngStartRow = VsfData.ROW
            lngRowCount = 1
        Else
            If VsfData.TextMatrix(VsfData.ROW, mlngRowCount) <> "1|1" Then
                lngStartRow = GetStartRow(VsfData.ROW)
                lngRowCount = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0))
            Else
                lngStartRow = VsfData.ROW
                lngRowCount = 1
            End If
        End If
        
        'ȷ��������ʼ��(ȡ���˼�飬����ʱ�ڽ��м��)
        lngRow = lngStartRow
'        If Val(VsfData.TextMatrix(lngRow, mlngDemo)) > 1 Then '����������
'            lngStartRow = lngRow - Val(VsfData.TextMatrix(lngRow, mlngDemo)) + 1
'            If Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) <> 1 Then
'                For lngStartRow = lngRow To VsfData.FixedRows Step -1
'                    If Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) = 1 Then
'                        Exit For
'                    End If
'                Next lngStartRow
'                If lngStartRow < VsfData.FixedRows Then Exit Sub
'            End If
'        End If
'
'        If VsfData.TextMatrix(lngStartRow, mlngDate) = "" Or VsfData.TextMatrix(lngStartRow, mlngTime) = "" Then
'            VsfData.ROW = lngStartRow
'            VsfData.COL = mlngDate
'            RaiseEvent AfterRowColChange("�������ݷ���ʱ��������ʼ�����ڻ�ʱ�䲻��Ϊ�ա�", True)
'            Exit Sub
'        End If
        
        lngStartRow = lngRow
        '׷������ʱ�������¼���ѡ�����ݵ������������ǲ��ܰ������ı�����Ϣ��
        '�磺һ������5�У����Ǵ��ı�������ֻ��3�У�ѡ�и���׷������ʱ��Ӧ��׷�ӵ���4�У�demo=1��Ϊ3�У�demo=4��Ϊ2��
        intNULL = lngStartRow + lngRowCount - 1
        For lngRow = lngRowCount To 1 Step -1
            blnNULL = True
            For lngCol = cHideCols + 1 To VsfData.Cols - 1
                If Not VsfData.ColHidden(lngCol) And lngCol < mlngNoEditor And ISEditAssistant(lngCol) = False Then
                    If VsfData.TextMatrix(lngRow + lngStartRow - 1, lngCol) <> "" And Not (IsDiagonal(lngCol) And InStr(1, VsfData.TextMatrix(lngRow + lngStartRow - 1, lngCol), "/") <> 0) Then
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
            If Trim(VsfData.TextMatrix(lngStartRow + lngRowCount - 1, mlngSignName)) <> "" Then
                VsfData.TextMatrix(intNULL, mlngSignName) = VsfData.TextMatrix(lngStartRow + lngRowCount - 1, mlngSignName)
                If mlngSignTime <> -1 Then VsfData.TextMatrix(intNULL, mlngSignTime) = VsfData.TextMatrix(lngStartRow + lngRowCount - 1, mlngSignTime)
            End If
        End If
        
        If intNULL + 1 <= lngStartRow + lngRowCount - 1 Then
            For lngRow = intNULL + 1 To lngStartRow + lngRowCount - 1
                '�������������
                For lngCol = cHideCols + 1 To VsfData.Cols - 1
                    If VsfData.ColHidden(lngCol) = True Then VsfData.TextMatrix(lngRow, lngCol) = ""
                Next lngCol
                VsfData.TextMatrix(lngRow, mlngRowCount) = (lngStartRow + lngRowCount - intNULL - 1) & "|" & (lngRow - intNULL)
                VsfData.TextMatrix(lngRow, mlngRowCurrent) = (lngStartRow + lngRowCount - intNULL - 1)
            Next
            lngRowCount = Val(Split(VsfData.TextMatrix(lngStartRow, mlngRowCount), "|")(0))
            '���±��д��ı������ݼ���Ϣ
            Call CellMap_UpdateAssistant(lngStartRow)
            blnTure = False
        Else
            '�����һ�������Ƿ�Ϊ��,������ǿ���ֱ����ӵ���һ��
            lngCurRow = lngStartRow + lngRowCount
            blnTure = False
            If lngCurRow >= VsfData.Rows Then
                blnTure = True
            ElseIf VsfData.TextMatrix(lngStartRow, c�ļ�ID) <> VsfData.TextMatrix(lngCurRow, c�ļ�ID) Then
                blnTure = True
            Else
                If Not VsfData.RowHidden(lngCurRow) Then
                    For lngCol = cHideCols + 1 To VsfData.Cols - 1
                        If Not VsfData.ColHidden(lngCol) And lngCol < mlngNoEditor Then
                            If VsfData.TextMatrix(lngCurRow, lngCol) <> "" Then
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
            '1)�����һ������
            VsfData.Rows = VsfData.Rows + 1
            VsfData.TextMatrix(VsfData.Rows - 1, c�ļ�ID) = VsfData.TextMatrix(lngStartRow, c�ļ�ID)
            VsfData.TextMatrix(VsfData.Rows - 1, c����ID) = VsfData.TextMatrix(lngStartRow, c����ID)
            VsfData.TextMatrix(VsfData.Rows - 1, c��ҳID) = VsfData.TextMatrix(lngStartRow, c��ҳID)
            VsfData.TextMatrix(VsfData.Rows - 1, cӤ��) = VsfData.TextMatrix(lngStartRow, cӤ��)
            
            '�ӵ�ǰ�м�¼�Ŀհ��п�ʼ��ÿ�е�λ��+�����ӵĿհ�����
            For lngRow = VsfData.Rows - 2 To lngStartRow + lngRowCount Step -1
                VsfData.RowPosition(lngRow) = lngRow + 1
            Next
            lngRow = lngStartRow + lngRowCount - 1
            '2)���кŷ����仯����ͬ������mrsCellMap�д��ڸ��кŵ��к�����
            Call CellMap_Update(lngRow, 1)
        End If
        '3)���·�����ؿ���
        Call AppendGroup(lngStartRow)
        lngRow1 = VsfData.ROW
        lngCol1 = VsfData.COL
        If InStr(1, VsfData.TextMatrix(lngRow1, mlngRowCount), "|") = 0 Then VsfData.TextMatrix(lngRow1, mlngRowCount) = "1|1"
        '4)��ԭ�з��������Ϸ���,��Ҫ����������
        If Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) > 0 Then '����������
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
                If lngStartRow < VsfData.FixedRows Then GoTo ErrNext
                lngStartRow = lngRow
            End If
            '������֯�������
            intGroupFirstRows = Val(Split(VsfData.TextMatrix(lngRow1, mlngRowCount), "|")(0))
            lngCurRow = lngRow1
            For lngRow = lngRow1 + intGroupFirstRows To VsfData.Rows - 1
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
                                    If Val(VsfData.TextMatrix(lngRow, mlngDemo)) <= 1 Then Exit For
                                    If VsfData.TextMatrix(lngRow, mlngRowCount) = "" Then VsfData.TextMatrix(lngRow, mlngRowCount) = "1|1"
                                    lngCount = Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(0)) + lngRow - 1
                                End If
                                    
                                strText = strText & IIf(strText = "", "", vbCrLf) & Replace(Replace(Replace(VsfData.TextMatrix(lngRow, lngCol), Chr(13), ""), Chr(10), ""), Chr(1), "")
                            Else
                                lngCount = lngCount + 1
                            End If
                        Next lngRow
                        mintType = -1: mblnShow = False
                        If strText = "" Then GoTo ErrNext
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
                        If Val(VsfData.TextMatrix(lngRow, mlngDemo)) <= 1 And intGroupFirstRows > 0 Then Exit For
                        If InStr(1, VsfData.TextMatrix(lngRow, mlngRowCount), "|") = 0 Then VsfData.TextMatrix(lngRow, mlngRowCount) = "1|1"
                        intGroupFirstRows = Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(0))
                        lngCurRow = lngRow
                        If CheckGroupDate(lngRow) = True Then
                            '�������޸ĲŽ�������̣�ȡ������¼��ʵ��ʱ��
                            If mblnDateAd Then
                                strDate = Format(VsfData.TextMatrix(lngRow, mlngActiveTime), "DD") & "/" & Format(VsfData.TextMatrix(lngRow, mlngActiveTime), "MM")
                            Else
                                strDate = Mid(VsfData.TextMatrix(lngRow, mlngActiveTime), 1, 10)
                            End If
                            strTime = Mid(VsfData.TextMatrix(lngRow, mlngActiveTime), 12, 5)
                        Else
                            '����ʱ���������
                            strDate = VsfData.TextMatrix(lngRow - Val(VsfData.TextMatrix(lngRow, mlngDemo)) + 1, mlngDate)
                            strTime = VsfData.TextMatrix(lngRow - Val(VsfData.TextMatrix(lngRow, mlngDemo)) + 1, mlngTime)
                        End If
                        
                        '1\����
                        strField = "ID|ҳ��|�к�|�к�|��¼ID|����|��λ|ɾ��"
                        If mlngDate <> -1 Then
                            strKey = mintҳ�� & "," & lngRow & "," & mlngDate
                            strValue = strKey & "|" & mintҳ�� & "|" & lngRow & "|" & mlngDate & "|" & _
                                Val(VsfData.TextMatrix(lngRow, mlngRecord)) & "|" & strDate & "|" & VsfData.TextMatrix(lngRow, mlngDemo) & "|0"
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
        Call SetActiveColColor
    'ճ��,���ʱ��Ҫͬ��mrsCellMap����
    Case conMenu_Edit_Copy
        '����ָ�������е�����
        If VsfData.ROW < VsfData.FixedRows Then Exit Sub
        If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then Exit Sub
        lngRow = GetStartRow(VsfData.ROW)

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
                mrsCopyMap.Fields(cControlFields + lngCol).Value = IIf(VsfData.TextMatrix(lngRow, lngCol + VsfData.FixedCols) = "", Null, VsfData.TextMatrix(lngRow, lngCol + VsfData.FixedCols))
            Next
            mrsCopyMap.Update
        Next
    Case conMenu_Edit_PASTE
        'ճ��ʱ����Ŀ�������帲�ǣ�ͬ�������������У���г���
        '���Ŀ���ܲ�ͬҳ����Ŀ��ͬ����λ��ͬ�����Բ����ǻ��Ŀ
        'ͬ������ռ�õ��������䣬�粻������ӿհ��У�����ճ��
        If VsfData.ROW < VsfData.FixedRows Then Exit Sub
        If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then VsfData.TextMatrix(VsfData.ROW, mlngRowCount) = "1|1"
        lngStartRow = GetStartRow(VsfData.ROW)
        If VsfData.TextMatrix(lngStartRow, mlngDemo) <> "" Then
            RaiseEvent AfterRowColChange("Ҫճ���������У������Ƿ��������С�", True)
            Exit Sub
        End If
        If mrsCopyMap.RecordCount = 0 Then Exit Sub
        
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
                If ISCollectSigned(Val(VsfData.TextMatrix(lngStartRow, c�ļ�ID)), Format(VsfData.TextMatrix(lngStartRow, mlngActiveTime), "YYYY-MM-DD"), Format(VsfData.TextMatrix(lngStartRow, mlngActiveTime), "HH:MM")) Then
                    blnTure = True
                    If MsgBox("��Ҫ�޸ĵ���������Ӧ�Ļ���������ǩ���������������������Ļ��������ݽ����ܱ�ճ�����������Ƿ������", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Sub
                End If
            End If
        End If
        
        '�õ�Ŀ�������е���ʼ��,������
        strField = "ID|ҳ��|�к�|�к�|��¼ID|����|ɾ��"
        lngCols = VsfData.Cols - 1
        If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then
            lngRow = VsfData.ROW
            lngStartRow = lngRow
            If mlngDate > -1 Then strDate = VsfData.TextMatrix(lngRow, mlngDate)
            strTime = VsfData.TextMatrix(lngRow, mlngTime)
        Else
            'ɾ�������������,����һ��
            lngRow = GetStartRow(VsfData.ROW)
            lngStartRow = lngRow
            strDate = VsfData.TextMatrix(lngRow, mlngDate)
            strTime = VsfData.TextMatrix(lngRow, mlngTime)
            lngRows = Val(Split(VsfData.TextMatrix(lngStartRow, mlngRowCount), "|")(0)) - 1
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

            If Val(VsfData.TextMatrix(lngRow + lngStartRow, c����ID)) = 0 And VsfData.TextMatrix(lngRow + lngStartRow, mlngRowCount) = "" Then
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
        '��¼�û��޸Ĺ��ĵ�Ԫ��
        If mlngDate <> -1 Then
            strKey = mintҳ�� & "," & lngStartRow & "," & mlngDate
            strValue = strKey & "|" & mintҳ�� & "|" & lngStartRow & "|" & mlngDate & "|" & _
                Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "|" & VsfData.TextMatrix(lngStartRow, mlngDate) & "|0"
            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
        End If
        '2\ʱ��
        strKey = mintҳ�� & "," & lngStartRow & "," & mlngTime
        strValue = strKey & "|" & mintҳ�� & "|" & lngStartRow & "|" & mlngTime & "|" & _
            Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "|" & VsfData.TextMatrix(lngStartRow, mlngTime) & "|0"
        Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
        '�����������
        With mrsCopyMap
            If .RecordCount <> 0 Then .MoveFirst
            Do While Not .EOF
                For lngCol = 0 To lngCols - VsfData.FixedCols
                    Select Case lngCol + VsfData.FixedCols
                    Case 1, c�ļ�ID, c����, c����, c����ID, c��ҳID, cӤ��, cѪѹƵ��, _
                         mlngDate, mlngTime, mlngOperator, mlngSigner, mlngSignTime, mlngRecord, mlngSignName
                    Case Else
                        If Not (blnTure = True And InStr(1, "," & strCols & ",", "," & lngCol - (cHideCols - 1) & ",") > 0) Then
                            If InStr(1, "," & strLockItem & ",", "," & lngCol - (cHideCols - 1) & ",") = 0 And InStr(1, "," & mstrCOLNothing & ",", "," & lngCol - (cHideCols - 1) & ",") = 0 Then
                                VsfData.TextMatrix(lngStartRow + .AbsolutePosition - 1, lngCol + VsfData.FixedCols) = NVL(.Fields(cControlFields + lngCol).Value)
    
                                '�޸ı�־
                                If .AbsolutePosition = .RecordCount Then
                                    strKey = mintҳ�� & "," & lngStartRow & "," & lngCol + VsfData.FixedCols
                                    strValue = strKey & "|" & mintҳ�� & "|" & lngStartRow & "|" & lngCol + VsfData.FixedCols & "|" & _
                                        Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "|" & GetMutilData(lngStartRow, lngCol + VsfData.FixedCols, lngTop, lngHeight) & "|0"
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
        Call SetActiveColColor
        mblnChange = True

    Case conMenu_Edit_Clear
        If VsfData.ROW < VsfData.FixedRows Then Exit Sub
        If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then Exit Sub
        
        lngRow = GetStartRow(VsfData.ROW)
        lngStartRow = lngRow
        lngRows = Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(0))
        
        '׼��ɾ��
        strField = "ID|ҳ��|�к�|�к�|��¼ID|����|��λ|ɾ��"
        
        If VsfData.TextMatrix(lngStartRow, mlngSigner) <> "" Then
            RaiseEvent AfterRowColChange("��ǩ�������ݲ�����ɾ����", True)
            Exit Sub
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
            RaiseEvent AfterRowColChange("���ڷ���������ʱ��������ɾ��������ʼ�С�", True)
            Exit Sub
        End If
        
        '���е����ݴ��ڻ�����ǩ�������ݲ�����ɾ��
        If Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) > 0 Then
            If CheckCollectIsData(lngStartRow, 1) = True Then
                If ISCollectSigned(Val(VsfData.TextMatrix(lngStartRow, c�ļ�ID)), Format(VsfData.TextMatrix(lngStartRow, mlngActiveTime), "YYYY-MM-DD"), Format(VsfData.TextMatrix(lngStartRow, mlngActiveTime), "HH:MM")) Then
                    RaiseEvent AfterRowColChange("��Ҫɾ�������ݴ��ڻ��������ݣ��ұ�����������Ӧ�Ļ���������ǩ����������ɾ����", True)
                    Exit Sub
                End If
            End If
        End If
        
        cmdWord.Visible = False
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
        End Select
        mintType = -1
        blnNULL = mblnShow
        mblnShow = False
        'ɾ������������
        lngRows = Val(Split(VsfData.TextMatrix(lngStartRow, mlngRowCount), "|")(0))
        lngRowCount = lngRows
        strAssistantCols = ""
        '��ȡ�������ݵĴ��ı�����������
        If Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) > 1 Then
            Call GetGroupAssistant(strAssistantCols, varAssistant)
        End If
        
        For intNULL = 2 To lngRows
            VsfData.RowHidden(lngRow + intNULL - 1) = True
        Next
        '�������ʼ�з������ݣ���������ı���Ϣ��ȡ���÷���
        '�磺������������飬����ڶ���ʱ�����ڶ�����Ķ������ۼ��ڵ�3����
        '��¼�û��޸Ĺ��ĵ�Ԫ��
        If Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) > 0 Then
            If CheckGroupDate(lngStartRow) = True Then
                '�������޸ĲŽ�������̣�ȡ������¼��ʵ��ʱ��
                If mblnDateAd Then
                    strDate = Format(VsfData.TextMatrix(lngStartRow, mlngActiveTime), "DD") & "/" & Format(VsfData.TextMatrix(lngStartRow, mlngActiveTime), "MM")
                Else
                    strDate = Mid(VsfData.TextMatrix(lngStartRow, mlngActiveTime), 1, 10)
                End If
                strTime = Mid(VsfData.TextMatrix(lngStartRow, mlngActiveTime), 12, 5)
            Else
                strDate = VsfData.TextMatrix(lngStartRow - Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) + 1, mlngDate)
                strTime = VsfData.TextMatrix(lngStartRow - Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) + 1, mlngTime)
            End If
        Else
            '��ͨ����
            strDate = VsfData.TextMatrix(lngStartRow, mlngDate)
            strTime = VsfData.TextMatrix(lngStartRow, mlngTime)
        End If
      
        strField = "ID|ҳ��|�к�|�к�|��¼ID|����|��λ|����|ɾ��"
        '��¼�û��޸Ĺ��ĵ�Ԫ��
        strKey = mintҳ�� & "," & lngStartRow & "," & mlngDate
        strValue = strKey & "|" & mintҳ�� & "|" & lngStartRow & "|" & mlngDate & "|" & _
            Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "|" & strDate & "|" & VsfData.TextMatrix(lngStartRow, mlngDemo) & "|0|1"
        Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
        '2\ʱ��
        strKey = mintҳ�� & "," & lngStartRow & "," & mlngTime
        strValue = strKey & "|" & mintҳ�� & "|" & lngStartRow & "|" & mlngTime & "|" & _
            Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "|" & strTime & "|" & VsfData.TextMatrix(lngStartRow, mlngDemo) & "|0|1"
        Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
        strField = "ID|ҳ��|�к�|�к�|��¼ID|����|����|ɾ��"
        'ɾ����ʼ���з�ͬ��������
        If strLockItem = "" Then
            VsfData.RowHidden(lngRow) = True
            '��д�޸ı�־
            For lngCol = mlngTime + 1 To mlngNoEditor - 1
                strKey = mintҳ�� & "," & lngStartRow & "," & lngCol
                strValue = strKey & "|" & mintҳ�� & "|" & lngStartRow & "|" & lngCol & "|" & _
                    Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "||0|1"
                Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
            Next
            If Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) >= 1 Then VsfData.TextMatrix(lngStartRow, mlngRowCount) = "1|1"
        Else
            '��д�޸ı�־(����ͬ������,������ʱ���в��������)``
            For lngCol = mlngTime + 1 To mlngNoEditor - 1
                If InStr(1, "," & strLockItem & ",", "," & lngCol - (cHideCols + VsfData.FixedCols - 1) & ",") = 0 And lngCol <> mlngDate And lngCol <> mlngTime Then
                    VsfData.TextMatrix(lngStartRow, lngCol) = ""

                    strKey = mintҳ�� & "," & lngStartRow & "," & lngCol
                    strValue = strKey & "|" & mintҳ�� & "|" & lngStartRow & "|" & lngCol & "|" & _
                        Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "||0|1"
                    Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                End If
            Next
            VsfData.TextMatrix(lngStartRow, mlngRowCount) = "1|1"
        End If
        
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
        mblnShow = blnNULL
        mblnChange = True
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
        strText = objTXT.Text
        intPos = objTXT.SelStart
        intLen = Len(objTXT)
        objTXT.Text = Mid(strText, 1, intPos) & mstrSymbol & Mid(strText, intPos + 1)
    
        If mintSymbol = -1 Then
            Call txtInput_KeyDown(vbKeyReturn, 0)
        Else
            Call txt_KeyDown(Val(txt(mintSymbol)), vbKeyReturn, 0)
        End If
    
    Case conMenu_Edit_Word
        Call cmdWord_Click
    Case conMenu_Edit_Import
        '��������
        If VsfData.ROW < VsfData.FixedRows Then Exit Sub '���������Ѿ���update�н������ж�
        Call ImportAmount
    Case conMenu_Edit_NewItem
        '�ڵ�ǰ��Ч�����У����ܵ�ǰ��Ч�������Ƕ��У�֮������һ�հ���
        If VsfData.ROW < VsfData.FixedRows Then Exit Sub
        If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then
            lngStartRow = VsfData.ROW
        Else
            lngStartRow = GetStartRow(VsfData.ROW)
        End If
        
        'ȷ��������ʼ��
        lngRow = lngStartRow
        If Val(VsfData.TextMatrix(lngRow, mlngDemo)) > 1 Then '����������
            lngStartRow = lngRow - Val(VsfData.TextMatrix(lngRow, mlngDemo)) + 1
            If Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) <> 1 Then
                For lngStartRow = lngRow To VsfData.FixedRows Step -1
                    If Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) = 1 Then
                        Exit For
                    End If
                Next lngStartRow
                If lngStartRow < VsfData.FixedRows Then Exit Sub
            End If
        End If
        lngRow = lngStartRow
        
        VsfData.Rows = VsfData.Rows + 1
        VsfData.TextMatrix(VsfData.Rows - 1, c�ļ�ID) = VsfData.TextMatrix(lngStartRow, c�ļ�ID)
        VsfData.TextMatrix(VsfData.Rows - 1, c����) = VsfData.TextMatrix(lngStartRow, c����)
        VsfData.TextMatrix(VsfData.Rows - 1, c����) = VsfData.TextMatrix(lngStartRow, c����)
        VsfData.TextMatrix(VsfData.Rows - 1, c����ID) = VsfData.TextMatrix(lngStartRow, c����ID)
        VsfData.TextMatrix(VsfData.Rows - 1, c��ҳID) = VsfData.TextMatrix(lngStartRow, c��ҳID)
        VsfData.TextMatrix(VsfData.Rows - 1, cӤ��) = VsfData.TextMatrix(lngStartRow, cӤ��)
        VsfData.TextMatrix(VsfData.Rows - 1, cѪѹƵ��) = VsfData.TextMatrix(lngStartRow, cѪѹƵ��)
        
        lngRows = Val(Split(VsfData.TextMatrix(lngStartRow, mlngRowCount), "|")(0))
        intNULL = lngStartRow + lngRows - 1
        'ȷ����ǰ�е��������л����ݵ����һ��
        For lngRow = lngStartRow + lngRows To VsfData.Rows - 1
            If lngRow > intNULL Then
                '������ÿһ�������������ݿ���ռ�ö��У�ֻ���ڵ�һ�б����˷�����������������������=ÿһ������������+ÿһ���������е�����
                If Val(VsfData.TextMatrix(lngRow, mlngDemo)) <= 1 Then
                       lngStartRow = lngRow - 1
                       Exit For
                End If
                intNULL = Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(0)) + lngRow - 1
            End If
        Next lngRow
        
'        strKey = VsfData.TextMatrix(lngStartRow, mlngRowCount)
'        If InStr(1, strKey, "|") <> 0 And strKey <> "1|1" Then
'            strKey = Split(strKey, "|")(0)
'            strKey = strKey & "|" & strKey
'            For lngRow = VsfData.ROW + 1 To VsfData.Rows - 1
'                If VsfData.TextMatrix(lngRow, mlngRowCount) = strKey Then
'                    lngStartRow = lngRow + 1
'                    Exit For
'                End If
'            Next
'        Else
'            lngStartRow = VsfData.ROW + 1
'        End If
        
        For lngRow = VsfData.Rows - 2 To lngStartRow + 1 Step -1  '�ӵ����ڶ��п�ʼ
            VsfData.RowPosition(lngRow) = lngRow + 1
        Next
        VsfData.ROW = lngStartRow + 1
        Call CellMap_Update(VsfData.ROW, 1)
        Call SetActiveColColor
        mblnChange = True
    Case conMenu_Edit_Save
        Call SaveME
    Case conMenu_Edit_Transf_Cancle
        Call CancelMe
    Case conMenu_Tool_Sign
        Call SignMe
    Case conMenu_Tool_SignEarse
        Call UnSignMe
    Case conMenu_Help_Help '����
        RaiseEvent UsrHelp
    Case conMenu_File_Exit '�˳�
        RaiseEvent UsrExit
    End Select

err_exit:
End Sub

Public Sub zlRptPrint(ByVal bytMode As Byte)
    '����:�����ݸ��Ƶ��ɴ�ӡ�Ķ��󣬵��ô�ӡ
    '����:  bytMode��1-��ӡ;2-Ԥ��;3-�����EXCEL
'    If mrsSelItems.Records.Count = 0 Then Exit Sub
    
    '-------------------------------------------------
    '�������ݱ��
    If zlDataToPrint(vsfPrint, VsfData) = False Then Exit Sub
    
    '-------------------------------------------------
    '���ô�ӡ��������
    Dim objPrint As New zlPrint1Grd
    Dim objAppRow As zlTabAppRow
    
    Set objPrint.Body = vsfPrint
    
    objPrint.Title.Text = cbo�����ļ���ʽ.Text
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("��ӡʱ��:" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim arrData
    Dim blnFind As Boolean, blnAllow As Boolean
    Dim strItem As String
    Dim intDo  As Integer, intCount As Integer
    
    blnAllow = (cmdˢ��.Tag <> "")
    Select Case Control.ID
    Case conMenu_Edit_Group_Append
        Control.Enabled = False
        If Not mblnInit Then Exit Sub
        Control.Enabled = Not mblnSigned And blnAllow And VsfData.ROW >= VsfData.FixedRows And (ISGroupAppend = True)
    Case conMenu_Edit_Copy
        Control.Enabled = Not mblnShow And blnAllow
    Case conMenu_Edit_PASTE
        Control.Enabled = False
        If Not mblnInit Then Exit Sub
        If mblnSigned Then Exit Sub
        If mrsCopyMap.State = 0 Then Exit Sub
        Control.Enabled = Not mblnShow And mrsCopyMap.RecordCount And blnAllow
    Case conMenu_Edit_Clear
        Control.Enabled = Not mblnSigned And blnAllow
    Case conMenu_Edit_SPECIALCHAR
        Control.Enabled = mblnShow And (mintType = 0 Or mintType = 6) And blnAllow
    Case conMenu_Edit_Word
        '60291:������,2013-04-17,ֻҪ���ı���Ŀ��������дʾ�ѡ��
        Control.Enabled = (mblnEditAssistant Or mblnEditText) And Not mblnSigned And blnAllow
    Case conMenu_Edit_NewItem
        Control.Enabled = Not mblnSigned And blnAllow
    Case conMenu_Edit_Save
        Control.Enabled = mblnChange And Not mblnSigned And blnAllow
    Case conMenu_Edit_Transf_Cancle
        Control.Enabled = mblnChange And blnAllow
    Case conMenu_Tool_Sign
        Control.Enabled = mblnSaved And Not mblnSigned And Not mblnChange And blnAllow
    Case conMenu_Tool_SignEarse
        Control.Enabled = mblnSaved And mblnSigned And Not mblnChange And blnAllow
    Case conMenu_Edit_Import '��������
        Control.Enabled = False
        If Not mblnInit Then Exit Sub
        Control.Enabled = Not mblnSigned And blnAllow And VsfData.ROW >= VsfData.FixedRows And mblnShow And mstrColCollect <> ""
        If Control.Enabled Then
            '�ж�ѡ������Ƿ��ǻ�����Ŀ��(һ�а�����������Ҳ����ʹ�ô˹���)
            blnFind = False
            For intDo = 0 To UBound(Split(mstrColCollect, "|"))
                If VsfData.COL - (cHideCols + VsfData.FixedCols - 1) = Split(Split(mstrColCollect, "|")(intDo), ";")(0) And InStr(1, Split(Split(mstrColCollect, "|")(intDo), ";")(1), ",") = 0 Then
                    blnFind = True
                    Exit For
                End If
            Next
            Control.Enabled = blnFind
        End If
    End Select
End Sub

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
                RaiseEvent AfterRowColChange(Split(strTest, ",")(1) & mrsItems!��Ŀ���� & " �Ѿ����󶨵�" & lngCol & "�У��������ظ��󶨣�", True)
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



Private Sub cmdWord_Click()
    Dim strInput As String
    '�����ʾ�ѡ����

    If Val(cmdWord.Tag) = -1 Then
        strInput = txtInput.Text
    Else
        strInput = txt(Val(cmdWord.Tag)).Text
    End If
    strInput = frmEditAssistant.ShowMe(Me, Val(VsfData.TextMatrix(VsfData.ROW, c����ID)), Val(VsfData.TextMatrix(VsfData.ROW, c��ҳID)), Val(VsfData.TextMatrix(VsfData.ROW, cӤ��)), strInput)
    
    If Val(cmdWord.Tag) = -1 Then
        txtInput.Text = strInput
        Call txtInput_KeyDown(vbKeyReturn, 0)
    Else
        txt(Val(cmdWord.Tag)).Text = strInput
        Call txt_KeyDown(Val(cmdWord.Tag), vbKeyReturn, 0)
    End If
End Sub

Private Sub cmdˢ��_Click()
    '��ȡ�ļ���ʽ
    mblnInit = False
    mblnHistory = False
    picNull.Visible = False
    mlng��ʽID = cbo�����ļ���ʽ.ItemData(cbo�����ļ���ʽ.ListIndex)
    mlng����ID = cbo����.ItemData(cbo����.ListIndex)
    
    Call InitVariable
    Call InitCons
    If Not ReadStruDef Then Exit Sub
    Call zlRefresh
    cmdˢ��.Tag = 1
    mblnInit = True
    
    '���浱ǰ����
    Call DataMap_Save
End Sub



Private Sub lstSelect_Click(Index As Integer)
    If Index = 0 Then
        If PicLst.Visible = False Then Exit Sub
        If lstSelect(0).ListIndex > 0 Then
            txtLst.Text = lstSelect(0).Text
        End If
    End If
End Sub

Private Sub optLevel_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If optLevel(0).Value Then
         SaveSetting "ZLSOFT", "˽��ģ��\usrTendFileMutilEditor\" & gstrUserName, "Value", 0
    Else
        SaveSetting "ZLSOFT", "˽��ģ��\usrTendFileMutilEditor\" & gstrUserName, "Value", 1
    End If
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
    If Shift = vbShiftMask And KeyCode = vbKeyDown Then KeyCode = 0: lstSelect(0).SetFocus
End Sub

Private Sub picSplit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    picSplit.Tag = 1
End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    If Val(picSplit.Tag) = 0 Then Exit Sub
    
    If picSplit.Top + Y < 4000 Then
        picSplit.Top = 4000
    ElseIf ScaleHeight - (picSplit.Top + Y) < 3000 Then
        picSplit.Top = ScaleHeight - 3000
    Else
        picSplit.Move picSplit.Left, picSplit.Top + Y
    End If
End Sub

Private Sub picSplit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Val(picSplit.Tag) = 1 Then Call cbsThis_Resize

    picSplit.Tag = 0
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
    If KeyAscii = 58 And VsfData.COL = mlngTime Then KeyAscii = 0
End Sub

Private Sub VsfData_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Dim lngRow As Long, lngCol As Long
    Dim dblHeight As Double, dblWidth As Double
    Dim strItemInfo As String
    
    If Not mblnInit Then Exit Sub
    Call InitCons
    
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
    Dim intMax As Integer
    Dim lngStart As Long
    Dim blnCheck As Boolean
    
    On Error Resume Next
    
    If Not mblnInit Then Exit Sub
    If VsfData.ROW < VsfData.FixedRows Then Exit Sub
    '��������ʾ��¼��ؼ�
    cmdWord.Visible = False
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
    End Select

    'δ������в�����¼������
    mintType = -1
    If InStr(1, mstrPrivs, "�����¼�Ǽ�") = 0 Then Exit Sub
    If mblnSigned Then Exit Sub
    If Not mblnShow Then Exit Sub
    
    If VsfData.TextMatrix(VsfData.ROW, mlngRowCount) = "" Then
        lngStart = VsfData.ROW
    Else
        lngStart = GetStartRow(VsfData.ROW)
    End If
    
    '����ǻ��Ŀ������༭
    If InStr(1, "," & mstrCOLNothing & ",", "," & VsfData.COL - cHideCols & ",") <> 0 Then Exit Sub
    If VsfData.COL <= cѪѹƵ�� Then Exit Sub
    
    If VsfData.TextMatrix(lngStart, mlngSigner) <> "" Then
        RaiseEvent AfterRowColChange("��ǩ�������ݲ������ٴα༭����ȡ��ǩ�������ԣ�", True)
        Exit Sub
    End If
    
    If VsfData.TextMatrix(VsfData.ROW, mlngDemo) <> "" Then
        'ֻ��������δ��������ݣ��������޸�������ʱ��
        If (VsfData.COL = mlngDate Or VsfData.COL = mlngTime) Then
            If Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) > 1 Then
                Exit Sub
            Else
                'If Val(VsfData.TextMatrix(VsfData.ROW, mlngRecord)) > 0 Then Exit Sub
            End If
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
            ElseIf VsfData.COL = mlngTime Or VsfData.COL = mlngDate Then
                blnCheck = True
            End If
            If blnCheck = True Then
                If ISCollectSigned(Val(VsfData.TextMatrix(lngStart, c�ļ�ID)), Mid(VsfData.TextMatrix(lngStart, mlngActiveTime), 1, 10), Format(VsfData.TextMatrix(lngStart, mlngActiveTime), "HH:MM")) Then
                    RaiseEvent AfterRowColChange("������������Ӧ�Ļ�����������ǩ�����������޸ĵ�ǰ�����л�����ʱ���У�", True)
                    Exit Sub
                End If
            End If
        End If
    End If
    
    If VsfData.COL <= mlngNoEditor - 1 Then Call ShowInput
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
    End Select
End Sub

Private Sub vsfData_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strInfo As String
    Dim strCols As String
    Dim intMax As Integer
    Dim strStart As String, strEnd As String
    Dim strItemInfo As String
    
    If mblnInit = False Then Exit Sub
    If OldRow = NewRow And OldCol = NewCol Then Exit Sub
    On Error GoTo ErrHand

    'ѡ����,ͬ��������ֱ���˳�,����˴������ʾ��Ϣ
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
    
    '��ȡ�ò�����ʷ����
    If OldRow <> NewRow Then
        Call RefreshHistoryData(NewRow)
    End If
    
    RaiseEvent AfterRowColChange(strInfo, False)
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub RefreshHistoryData(ByVal lngRow As Long)
'ˢ����ʷ����
    Dim strStart As String
    Dim rsTemp As New ADODB.Recordset
    Dim strCurrDate As String
    
    On Error GoTo ErrHand
    
    strCurrDate = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:mm")
    
    If lngRow > VsfData.Rows Then
        lngRow = VsfData.Rows - 1
        VsfData.ROW = lngRow
    End If
    If Val(VsfData.TextMatrix(lngRow, c�ļ�ID)) > 0 Then
        mstrMaxDate = Format(DateAdd("d", mintPreDays, CDate(strCurrDate)), "yyyy-MM-dd HH:mm")
        strStart = Format(DateAdd("d", -1 * (mintPreDays + 1), CDate(strCurrDate)), "yyyy-MM-dd") & " 23:59:59"  '���������
        
        'װ������
        Call SQLCombination
        gstrSQL = Replace(mstrSQL, " And 1=2", " And l.����ʱ�� between [2] and [3]")
        Call SQLDIY(gstrSQL)
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��������", Val(VsfData.TextMatrix(lngRow, c�ļ�ID)), CDate(strStart), CDate(mstrMaxDate))
        '�����ݲ����û����¼���ĸ�ʽ,ͬʱʵ��һ�����ݷ�����ʾ�Ĺ���
        Call PreTendFormatHistory(rsTemp)
    End If
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
    On Error GoTo ErrHand
    
    If Not mblnInit Then Exit Sub
    If KeyCode = vbKeyReturn Then
        Call MoveNextCell
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

Private Sub InitVariable()
    '�������
    mlngDate = -1
    mlngTime = -1
    mlngOperator = -1
    mlngSigner = -1
    mlngSignName = -1
    mlngSignTime = -1
    mlngRecord = -1
    mlngNoEditor = -1
    mlngActiveTime = -1
    mintType = -1
    
    mblnShow = False
    mblnSigned = False
    mblnSaved = False
    mblnChange = False
    mblnEditAssistant = False
    mblnEditText = False
    mblnEditHistoryAssistant = False
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
    mintType = -1
End Sub

Private Function InitMenuBar() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrControlMenu As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom
    Dim cbrPop As CommandBarControl
    Dim cbrToolBar As CommandBar
    Dim rs As ADODB.Recordset
    Dim objExtendedBar As CommandBar
    Dim objMenu As CommandBarPopup

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
        .SetIconSize True, 24, 24
        .LargeIcons = True
    End With

    '------------------------------------------------------------------------------------------------------------------
    '�˵�����
    cbsThis.ActiveMenuBar.Title = "�˵�"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)

        '------------------------------------------------------------------------------------------------------------------
        '����������
        
        
        Set cbrToolBar = cbsThis.Add("��׼����", xtpBarTop)
        cbrToolBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
        cbrToolBar.ModifyStyle XTP_CBRS_GRIPPER, 0
        cbrToolBar.ShowTextBelowIcons = False
        With cbrToolBar.Controls
            Set objMenu = .Add(xtpControlPopup, conMenu_FilePopup, "��ӡ(&F)", -1, False)
            objMenu.ID = conMenu_FilePopup
            
            Set cbrControlMenu = objMenu.CommandBar.Controls.Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)")
            Set cbrControlMenu = objMenu.CommandBar.Controls.Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
            Set cbrControlMenu = objMenu.CommandBar.Controls.Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
            Set cbrControlMenu = objMenu.CommandBar.Controls.Add(xtpControlButton, conMenu_File_Excel, "�����&Excel")
            
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Group_Append, "׷��"): cbrControl.IconId = 3045: cbrControl.ToolTipText = "׷�ӷ���"
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Copy, "����"): cbrControl.ToolTipText = "����(Ctrl+C)": cbrControl.BeginGroup = True
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_PASTE, "ճ��"):  cbrControl.ToolTipText = "ճ��(Ctrl+V)"
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Clear, "���"):   cbrControl.ToolTipText = "���"

            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_SPECIALCHAR, "�������"):  cbrControl.ToolTipText = "�����������(Ctrl+D)": cbrControl.BeginGroup = True
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Word, "�ʾ�ѡ��"):  cbrControl.ToolTipText = "�ʾ�ѡ��(Ctrl+W)"
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Import, "��������"):  cbrControl.ToolTipText = "��������(Ctrl+I)"
            
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "���ӿ���"
            
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "����"): cbrControl.BeginGroup = True
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Transf_Cancle, "ȡ��")
            Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Sign, "ǩ��"): cbrControl.BeginGroup = True
            Set cbrControl = .Add(xtpControlButton, conMenu_Tool_SignEarse, "ȡ��"): cbrControl.IconId = 229
            
            Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): cbrControl.BeginGroup = True
            Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
        End With

        For Each cbrControl In cbrToolBar.Controls
            If cbrControl.Type <> xtpControlCustom And cbrControl.Type <> xtpControlLabel Then
                cbrControl.Style = xtpButtonIconAndCaption
            End If
        Next

        '------------------------------------------------------------------------------------------------------------------
        '����������
        pic��������.Height = IIf(mblnBlowup = True, 375, 300)
        Set cbrToolBar = cbsThis.Add("��������", xtpBarTop)
        cbrToolBar.ShowTextBelowIcons = False
        cbrToolBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
        cbrToolBar.ModifyStyle XTP_CBRS_GRIPPER, 0
        With cbrToolBar.Controls
            Set cbrCustom = .Add(xtpControlCustom, 0, "")
            cbrCustom.Flags = xtpFlagAlignLeft
            cbrCustom.Handle = pic��������.hWnd
            cbrCustom.ToolTipText = "����"
        End With

         '�����
        With cbsThis.KeyBindings
            .Add FCONTROL, Asc("C"), conMenu_Edit_Copy
            .Add FCONTROL, Asc("V"), conMenu_Edit_PASTE
            .Add FCONTROL, Asc("D"), conMenu_Edit_SPECIALCHAR
            .Add FCONTROL, Asc("W"), conMenu_Edit_Word
            .Add FCONTROL, Asc("I"), conMenu_Edit_Import
            .Add 0, VK_F1, conMenu_Help_Help
        End With

    InitMenuBar = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function CheckTime(ByVal lngRow As Long, ByVal strTime As String, ByVal strCurTime As String, ByRef strMsg As String) As Boolean
    Dim blnMsg As Boolean
    Dim blnExist As Boolean
    Dim lng�ļ�ID As Long, lng����ID As Long, lng��ҳID As Long, intӤ�� As Integer
    Dim rsTemp As New ADODB.Recordset
    Dim strBabyOutTime As String
    On Error GoTo ErrHand
    '���ݷ���ʱ������ڵ�ǰ���ҵ���Чʱ�䷶Χ��
    lng�ļ�ID = Val(VsfData.TextMatrix(lngRow, c�ļ�ID))
    lng����ID = Val(VsfData.TextMatrix(lngRow, c����ID))
    lng��ҳID = Val(VsfData.TextMatrix(lngRow, c��ҳID))
    intӤ�� = Val(VsfData.TextMatrix(lngRow, cӤ��))
    
    blnMsg = (strMsg <> "")
    
    gstrSQL = "Select ��ʼʱ��,����ʱ�� From ���˻����ļ� Where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�ļ�ID", lng�ļ�ID)
    If rsTemp.RecordCount > 0 Then
        '����ļ���ʼ,����ʱ��
        If Format(strTime, "yyyy-MM-dd HH:mm") < Format(NVL(rsTemp!��ʼʱ��), "yyyy-MM-dd HH:mm") Then
            strMsg = "����ʱ�䲻��С���ļ���ʼʱ��[" & NVL(rsTemp!��ʼʱ��) & "]"
            GoTo exitHand
        End If
        If NVL(rsTemp!����ʱ��) <> "" Then
            If Format(strTime, "yyyy-MM-dd HH:mm") <= Format(NVL(rsTemp!����ʱ��), "yyyy-MM-dd HH:mm") Then
                strMsg = "����ʱ�䲻�ܴ����ļ�����ʱ��[" & NVL(rsTemp!����ʱ��) & "]"
                GoTo exitHand
            End If
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
            If Format(strTime, "yyyy-MM-dd HH:mm") >= Format(!��ʼʱ��, "yyyy-MM-dd HH:mm") And Format(strTime, "yyyy-MM-dd HH:mm") <= Format(!��ֹʱ��, "yyyy-MM-dd HH:mm") Then
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
            If !��ʼԭ�� = 1 And Format(strTime, "yyyy-MM-dd HH:mm") < Format(!��ʼʱ��, "yyyy-MM-dd HH:mm") Then
                strMsg = "��" & lngRow & "�еķ���ʱ��" & strTime & "����[����ʱ�䲻��С�ڲ�����Ժʱ��:" & !��ʼʱ�� & "]"
                GoTo exitHand
            End If
        End If
        .Filter = "��ʼԭ��=2"
        If .RecordCount <> 0 Then
            If !��ʼԭ�� = 2 And Format(strTime, "yyyy-MM-dd HH:mm") < Format(!��ʼʱ��, "yyyy-MM-dd HH:mm") Then
                strMsg = "��" & lngRow & "�еķ���ʱ��" & strTime & "����[����ʱ�䲻��С�ڲ������ʱ��:" & !��ʼʱ�� & "]"
                GoTo exitHand
            End If
        End If
        .Filter = "��ʼԭ��=10"
        If .RecordCount <> 0 Then
            If !��ʼԭ�� = 10 And Format(strTime, "yyyy-MM-dd HH:mm") > Format(!��ֹʱ��, "yyyy-MM-dd HH:mm") Then
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
    End Select
    If Val(strOrders) <> 0 Then
        If Not CheckValid(strText, strOrders, strInfo) Then Exit Function
    ElseIf VsfData.COL = mlngDate Or VsfData.COL = mlngTime Then
        If Not CheckDateTime(strText, strInfo) Then Exit Function
    End If

    strReturn = strText
    CheckInput = True
End Function

Private Function CheckDateTime(strText As String, strInfo As String) As Boolean
    Dim arrData
    Dim blnCheck As Boolean
    Dim strCurrDate As String
    Dim strDate As String, strMonth As String, strDay As String
    Dim rsCheck As New ADODB.Recordset
    Dim arrTime As Variant
    
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

            strDate = Mid(strCurrDate, 1, 5) & ToStandDate(strText)
            '����Ƿ����༭֮ǰ��ʱ��(һ���µ�����)
            If CDate(strDate) > CDate(Mid(strCurrDate, 1, 10)) And Mid(strCurrDate, 6, 2) = "01" And Mid(strDate, 6, 2) = "12" Then
                strDate = DateAdd("yyyy", -1, CDate(strDate))
            End If
            If Not IsDate(strDate) Then
                strInfo = "¼������ݲ��ǺϷ������ڣ���1��12�գ�12/01"
                Exit Function
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
        End If
        
        If Format(strDate, "YYYY-MM-DD") > Format(DateAdd("d", mintPreDays, CDate(strCurrDate)), "YYYY-MM-DD") Then
            strInfo = "¼��������ѳ�������[����¼��������" & mintPreDays & "��]��ָ���ķ�Χ��"
            Exit Function
        End If
        
        'Ϊ�˱����û����������ݼ�����ʱ�����ûس����Լ��������ݽ��м��
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
                strDate = Mid(strCurrDate, 1, 5) & ToStandDate(strDate)
            Else
                strDate = Format(VsfData.TextMatrix(VsfData.ROW, mlngDate), "yyyy-MM-dd")
            End If
            '����Ƿ����༭֮ǰ��ʱ��(һ���µ�����)
            If CDate(strDate) > CDate(Mid(strCurrDate, 1, 10)) And Mid(strCurrDate, 6, 2) = "01" And Mid(strDate, 6, 2) = "12" Then
                strDate = DateAdd("yyyy", -1, CDate(strDate))
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
        Set rsCheck = zlDatabase.OpenSQLRecord(gstrSQL, "��鷢��ʱ��", Val(VsfData.TextMatrix(VsfData.ROW, c�ļ�ID)), CDate(strDate), Val(VsfData.TextMatrix(VsfData.ROW, mlngRecord)))
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
            If Format(VsfData.TextMatrix(VsfData.ROW, mlngActiveTime), "YYYY-MM-DD HH:mm") = Format(strDate, "YYYY-MM-DD HH:mm") Then
                blnCheck = False
            End If
        End If
        If blnCheck = True Then
            If CheckCollectIsData(VsfData.ROW) = True Then
                If ISCollectSigned(Val(VsfData.TextMatrix(VsfData.ROW, c�ļ�ID)), Format(strDate, "YYYY-MM-DD"), Format(strDate, "HH:MM")) Then
                    strInfo = "��¼���ʱ������Ӧ�Ļ�����������ǩ����������������µĻ��������ݣ�"
                    Exit Function
                End If
            End If
        End If
        
        '70990:������,2014-03-13
        '���ݷ���ʱ�䲻���ڵ�ǰ����Ա�������ҵ���Чʱ����ǰ
        If Not CheckTime(VsfData.ROW, strDate, strCurrDate, strInfo) Then
            Exit Function
        End If
    End If

    CheckDateTime = True
End Function

Private Function CheckChangeDataTime(ByVal lngRow As Long, ByVal strCurDate As String, ByRef strMsg As String) As Boolean
'�����¼���ʱ�䣬�Ƿ������е�ʱ����ͬ�������ͬ����ʾ����¼��
    Dim strDateHistory As String, strTimeHistory As String, strDatetime As String '�û��Ѿ�¼������ں�ʱ��
    Dim lngCurRow As Long, intPage As Integer, blnDel As Boolean, blnTrue As Boolean
    Dim strCurrDate As String, lngRecord As Long, strActiveTime As String
    Dim strRows As String, strPages As String, strTimes As String, lngCol As Long
    Dim lng�ļ�ID As Long
    Dim arrRows
    On Error GoTo ErrHand
    
    lng�ļ�ID = Val(VsfData.TextMatrix(lngRow, c�ļ�ID))
    strCurrDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    With mrsCellMap
        .Filter = "�к�=" & mlngDate & " OR �к�=" & mlngTime
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            If Not (lngCurRow = !�к� And intPage = !ҳ��) Then
                blnDel = False
endWork:
                If Not (lng�ļ�ID = Val(VsfData.TextMatrix(lngCurRow, c�ļ�ID))) Then GoTo ErrNext
                If lngCurRow = lngRow And intPage = mintҳ�� Then GoTo ErrNext
                If lngCurRow > 0 Then
                    blnDel = VsfData.RowHidden(lngCurRow)
                    strActiveTime = Format(VsfData.TextMatrix(lngCurRow, mlngActiveTime), "YYYY-MM-DD HH:mm:ss")
                End If
                
                If blnTrue = True And strDatetime <> "" Then
                    If Format(strDatetime, "YYYY-MM-DD HH:mm:ss") = Format(strCurDate, "YYYY-MM-DD HH:mm:ss") Then
                        '������ͬʱ�������û��ɾ����ֱ�ӽ�����ʾ
                        If blnDel = False Then
                            strMsg = "��" & lngCurRow & "���Ѿ�������ͬʱ������ݣ����飡"
                            Exit Function
                        Else
                            If lngRecord > 0 Then '���������ɾ�������ʱ���ԭ��ʱ����ֱͬ����ʾ������ͬ�ָ�ʱ��Ϊԭ��ʱ��
                                If Format(strDatetime, "YYYY-MM-DD HH:mm:ss") = Format(strActiveTime, "YYYY-MM-DD HH:mm:ss") Then
                                    strMsg = "��¼���ʱ���Ѿ�������ʷ���ݣ�"
                                    Exit Function
                                Else '�ָ�ʱ��Ϊԭ��ʱ��
                                    VsfData.TextMatrix(lngCurRow, mlngDate) = Format(strActiveTime, "YYYY-MM-DD")
                                    VsfData.TextMatrix(lngCurRow, mlngTime) = Mid(strActiveTime, 12, 5)
                                    '��¼�кź�ҳ��
                                    strRows = strRows & "," & lngCurRow
                                    strPages = strPages & "," & intPage
                                    strTimes = strTimes & "," & strActiveTime
                                End If
                            Else 'δ���������ɾ����ֱ����ռ�¼��������Ϣ
                                For lngCol = mlngDate To VsfData.Cols - 1
                                    VsfData.TextMatrix(lngCurRow, lngCol) = ""
                                Next lngCol
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
                            strDateHistory = Mid(strCurrDate, 1, 5) & ToStandDate(strDateHistory)
                            '����Ƿ����༭֮ǰ��ʱ��(һ���µ�����)
                            If CDate(strDateHistory) > CDate(Mid(strCurrDate, 1, 10)) And Mid(strCurrDate, 6, 2) = "01" And Mid(strDateHistory, 6, 2) = "12" Then
                                strDateHistory = DateAdd("yyyy", -1, CDate(strDateHistory))
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
    Dim i As Integer, j As Integer, blnNumber As Boolean
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
            strName = GetActivePart(VsfData.COL, i) & mrsItems!��Ŀ����
            blnNumber = False
            If strText <> "" Then
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
                            dblMin = Split(mrsItems!��Ŀֵ��, ";")(0)
                            dblMax = Split(mrsItems!��Ŀֵ��, ";")(1)
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
                If IsNumeric(strText) And blnNumber = True Then
                    If Val(strText) < 1 And Val(strText) > 0 Then strText = "0" & Val(strText)
                End If
                strFormat = Replace(strFormat, "[" & strName & "]", strText)
            Else
                'ɾ������Ŀ
                Call SubstrPro(strFormat, strName)
            End If
        Else
            strFormat = strReturn
        End If
    Next
    If j = -1 Then
        strOrders = arrOrder(i)
        If Val(strOrders) <> 0 Then
            mrsItems.Filter = "��Ŀ���=" & strOrders
            strName = mrsItems!��Ŀ����
            strFormat = Replace(strFormat, "[" & strName & "]", strText)
        End If
    End If
    mrsItems.Filter = 0

    strFormat = Replace(strFormat, "{", "")
    strFormat = Replace(strFormat, "}", "")
    If strFormat = SubstrFormat(strFormat1, arrOrder) Then strFormat = ""
    
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

'Public Sub SubstrPro(strFormat As String, ByVal strName As String, Optional ByVal intType As Integer = 0)
'    Dim i As Integer, j As Integer, l As Integer, r As Integer
'    'intType=0-ɾ��ָ����ʽ��;1-�õ�ָ����ʽ��
'    j = Len(strFormat)
'    i = InStr(1, strFormat, "[" & strName & "]")
'    If i = 0 Then Exit Sub
'
'    For l = i To 1 Step -1
'        If Mid(strFormat, l, 1) = "{" Then Exit For
'    Next
'    For r = i To j
'        If Mid(strFormat, r, 1) = "}" Then Exit For
'    Next
'    If intType = 0 Then
'        strFormat = Mid(strFormat, 1, l - 1) & Mid(strFormat, r + 1)
'    Else
'        strFormat = Mid(strFormat, l, r - l + 1)
'    End If
'End Sub

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
    Dim strDate As String, strTime As String    '�����׼�¼��������ʱ��
    Dim strReturn As String, strMsg As String, strPart As String
    Dim lngStart As Long, lngStartGroup As Long, lngMutilRows As Long, lngDeff As Long, intGroupFirstRows As Integer, intBound As Integer, intRowCount As Integer
    Dim intRow As Integer, intRowGroup As Integer, intCount As Integer, intNULL As Integer  '����ж��ٿ���
    Dim blnTrue As Boolean, blnDate As Boolean, strRows As String
    Dim lngDemo As Long, lngLastNull As Long, lngLastNoNull As Long
    '��ֵȻ���ƶ�����һ����Ч��Ԫ��
    Dim strKey As String, strField As String, strValue As String, strAppend As String
    Dim blnCallback As Boolean, blnReseGroupAssistant As Boolean, blnGroupAddNum As Boolean '��������������
    '���ı��к�������Ϣ
    Dim varAssistant() As Variant, strAssistantCols As String
    Dim blnNewRow As Boolean
    
    On Error GoTo ErrHand
    If VsfData.ROW >= VsfData.Rows Then Exit Function
    blnReseGroupAssistant = False
    blnNewRow = Val(GetSetting("ZLSOFT", "˽��ģ��\usrTendFileMutilEditor\" & gstrUserName, "Value")) = 0
    '�������,���ϸ���ٴε���Ҫ��¼��
    If mintType >= 0 Then
        If strText = "" Then
            strReturn = Replace(Replace(Replace(strReturn, Chr(10), ""), Chr(13), ""), Chr(1), "")
            If Not CheckInput(strReturn, strMsg) Then
                RaiseEvent AfterRowColChange(strMsg, True)
                Exit Function
            End If
            strText = strReturn
        Else
            strReturn = strText
            mstrData = strText
        End If
        '��ǵ�ǰ��Ϊ������
        blnDate = (InStr(1, "," & mlngDate & "," & mlngTime & ",", "," & VsfData.COL & ",") > 0)
        lngDemo = Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo))
        blnGroup = Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) >= 1
        If VsfData.TextMatrix(VsfData.ROW, mlngRowCount) = "" Then VsfData.TextMatrix(VsfData.ROW, mlngRowCount) = "1|1"
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
                            If VsfData.TextMatrix(intRow, intCount) <> "" And Not (IsDiagonal(intCount) And InStr(1, VsfData.TextMatrix(intRow, intCount), "/") <> 0) Then
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
        Else
            lngMutilRows = 1
            If VsfData.TextMatrix(VsfData.ROW, mlngDemo) = 1 And (mblnEditAssistant Or blnDate) Then
                '��¼������ʼ�е���������
                intGroupFirstRows = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0))
                intBound = VsfData.ROW + intGroupFirstRows - 1
                For intCount = VsfData.ROW + intGroupFirstRows To VsfData.Rows - 1
                    '������ÿһ�������������ݿ���ռ�ö��У�ֻ���ڵ�һ�б����˷�����������������������=ÿһ������������+ÿһ���������е�����
                    If intCount > intBound Then
                        If Val(VsfData.TextMatrix(intCount, mlngDemo)) <= 1 Then Exit For      '����������·�����˳�
                        intBound = Val(Split(VsfData.TextMatrix(intCount, mlngRowCount), "|")(0)) + intCount - 1
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
            End If
            lngStart = VsfData.ROW
        End If
        
        '׼����ֵ
        With txtLength
            '������ʱ���еĿ�Ȳ���,Ϊ�˱��ⷵ�ض���,ǿ������Ϊ5000
            .Width = IIf(VsfData.COL = mlngDate Or VsfData.COL = mlngTime, 5000, VsfData.CellWidth)
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
        blnGroupAddNum = False
        blnTrue = blnGroup = True And mblnEditAssistant
        If intCount > lngMutilRows - 1 Then
            '����������������ʱ������Ҫ��¼����������ݲ���¼����ı�������
'            If mblnEditAssistant = True And Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) = 1 And lngMutilRows = intGroupFirstRows Then
'                strMsg = "������������ʱ������������ݵķ��飬�����¼����Ķ���Ŀ���ݣ�"
'                RaiseEvent AfterRowColChange(strMsg, True)
'                strMsg = ""
'                Exit Function
'            End If
            '������������,�������������������������ӵ�����
            '20110830���������ͬһ�����У��������ı��ֽ⵽���У�������ı�����ͳһ�������һ����;�ڷ����а��س�,ֻ���������ݽ����޸�,�����з����仯
            intNULL = intCount - (lngMutilRows - 1)
            For intRow = lngMutilRows To intCount
                '��֤��ǰ�����������һҳ����ʾȫ
                If intRow + lngStart > VsfData.Rows - 1 Then Exit For
                
                If Val(VsfData.TextMatrix(intRow + lngStart, c����ID)) = 0 And VsfData.TextMatrix(intRow + lngStart, mlngRowCount) = "" Then
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
                For intRow = lngMutilRows To intCount
                    VsfData.TextMatrix(intRow + lngStart, c�ļ�ID) = VsfData.TextMatrix(lngStart, c�ļ�ID)
                    VsfData.TextMatrix(intRow + lngStart, c����ID) = VsfData.TextMatrix(lngStart, c����ID)
                    VsfData.TextMatrix(intRow + lngStart, c��ҳID) = VsfData.TextMatrix(lngStart, c��ҳID)
                    VsfData.TextMatrix(intRow + lngStart, cӤ��) = VsfData.TextMatrix(lngStart, cӤ��)
                Next intRow

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
                        If Val(VsfData.TextMatrix(lngStart + intBound, mlngDemo)) > 0 Then
                            If CheckGroupDate(lngStart + intBound) = True Then
                                '�������޸ĲŽ�������̣�ȡ������¼��ʵ��ʱ��
                                If mblnDateAd Then
                                    strDate = Format(VsfData.TextMatrix(lngStart + intBound, mlngActiveTime), "DD") & "/" & Format(VsfData.TextMatrix(lngStart + intBound, mlngActiveTime), "MM")
                                Else
                                    strDate = Mid(VsfData.TextMatrix(lngStart + intBound, mlngActiveTime), 1, 10)
                                End If
                                strTime = Mid(VsfData.TextMatrix(lngStart + intBound, mlngActiveTime), 12, 5)
                            Else
                                '����ʱ���������
                                strDate = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngDate)
                                strTime = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngTime)
                            End If
                        Else
                            '��ͨ����
                            strDate = VsfData.TextMatrix(lngStart + intBound, mlngDate)
                            strTime = VsfData.TextMatrix(lngStart + intBound, mlngTime)
                        End If
                        
                        '1\����
                        strField = "ID|ҳ��|�к�|�к�|��¼ID|����|��λ|ɾ��"
                        If mlngDate <> -1 Then
                            strKey = mintҳ�� & "," & lngStart + intBound & "," & mlngDate
                            strValue = strKey & "|" & mintҳ�� & "|" & lngStart + intBound & "|" & mlngDate & "|" & _
                                Val(VsfData.TextMatrix(lngStart + intBound, mlngRecord)) & "|" & strDate & "|" & VsfData.TextMatrix(lngStart + intBound, mlngDemo) & "|" & IIf(VsfData.RowHidden(lngStart + intBound) = True, 1, 0)
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
                        If blnGroup And InStr(1, "," & mlngDemo & "," & mlngRecord & "," & mlngActiveTime & ",", "," & intCount & ",") = 0 Then
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
                        If Val(VsfData.TextMatrix(lngStart + intBound, mlngDemo)) > 0 Then
                            If CheckGroupDate(lngStart + intBound) = True Then
                                '�������޸ĲŽ�������̣�ȡ������¼��ʵ��ʱ��
                                If mblnDateAd Then
                                    strDate = Format(VsfData.TextMatrix(lngStart + intBound, mlngActiveTime), "DD") & "/" & Format(VsfData.TextMatrix(lngStart + intBound, mlngActiveTime), "MM")
                                Else
                                    strDate = Mid(VsfData.TextMatrix(lngStart + intBound, mlngActiveTime), 1, 10)
                                End If
                                strTime = Mid(VsfData.TextMatrix(lngStart + intBound, mlngActiveTime), 12, 5)
                            Else
                                '����ʱ���������
                                strDate = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngDate)
                                strTime = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngTime)
                            End If
                        Else
                            '��ͨ����
                            strDate = VsfData.TextMatrix(lngStart + intBound, mlngDate)
                            strTime = VsfData.TextMatrix(lngStart + intBound, mlngTime)
                        End If
                        
                        '1\����
                        strField = "ID|ҳ��|�к�|�к�|��¼ID|����|��λ|ɾ��"
                        If mlngDate <> -1 Then
                            strKey = mintҳ�� & "," & lngStart + intBound & "," & mlngDate
                            strValue = strKey & "|" & mintҳ�� & "|" & lngStart + intBound & "|" & mlngDate & "|" & _
                                Val(VsfData.TextMatrix(lngStart + intBound, mlngRecord)) & "|" & strDate & "|" & VsfData.TextMatrix(lngStart + intBound, mlngDemo) & "|" & IIf(VsfData.RowHidden(lngStart + intBound) = True, 1, 0)
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
                        If CheckGroupDate(lngStart + intRow) = True Then
                            '�������޸ĲŽ�������̣�ȡ������¼��ʵ��ʱ��
                            If mblnDateAd Then
                                strDate = Format(VsfData.TextMatrix(lngStart + intRow, mlngActiveTime), "DD") & "/" & Format(VsfData.TextMatrix(lngStart + intRow, mlngActiveTime), "MM")
                            Else
                                strDate = Mid(VsfData.TextMatrix(lngStart + intRow, mlngActiveTime), 1, 10)
                            End If
                            strTime = Mid(VsfData.TextMatrix(lngStart + intRow, mlngActiveTime), 12, 5)
                        Else
                            '����ʱ���������
                            strDate = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngDate)
                            strTime = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngTime)
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
                                Val(VsfData.TextMatrix(lngStart + intRow, mlngRecord)) & "|" & strDate & "|" & VsfData.TextMatrix(lngStart + intRow, mlngDemo) & "|" & IIf(VsfData.RowHidden(lngStart + intRow) = True, 1, 0)
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
            If Left(strRows, 1) = "," Then strRows = Mid(strRows, 2)
            If Right(strRows, 1) = "," Then strRows = Mid(strRows, 1, Len(strRows) - 1)
            For intRow = intNULL + 1 To lngStart + lngMutilRows - 1
                VsfData.TextMatrix(intRow, mlngRowCount) = ""
                VsfData.TextMatrix(intRow, mlngRowCurrent) = ""
                VsfData.TextMatrix(intRow, mlngRecord) = ""
                If mlngSignName <> -1 Then VsfData.TextMatrix(intRow, mlngSignName) = ""
                If mlngOperator <> -1 Then VsfData.TextMatrix(intRow, mlngOperator) = ""
                If mlngSignTime <> -1 Then VsfData.TextMatrix(intRow, mlngSignTime) = ""
                If blnReseGroupAssistant = True Then
                    If InStr(1, "," & strRows & ",", "," & intRow & ",") = 0 Then strRows = strRows & "," & intRow
                ElseIf Not blnGroup Then
                    If InStr(1, "," & strRows & ",", "," & intRow & ",") = 0 Then strRows = strRows & "," & intRow
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
            If Val(VsfData.TextMatrix(lngStart, mlngDemo)) > 0 Then
                If CheckGroupDate(lngStart) = True Then
                    '�������޸ĲŽ�������̣�ȡ������¼��ʵ��ʱ��
                    If mblnDateAd Then
                        strDate = Format(VsfData.TextMatrix(lngStart, mlngActiveTime), "DD") & "/" & Format(VsfData.TextMatrix(lngStart, mlngActiveTime), "MM")
                    Else
                        strDate = Mid(VsfData.TextMatrix(lngStart, mlngActiveTime), 1, 10)
                    End If
                    strTime = Mid(VsfData.TextMatrix(lngStart, mlngActiveTime), 12, 5)
                Else
                    '����ʱ���������
                    strDate = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngDate)
                    strTime = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngTime)
                End If
            Else
                '��ͨ����
                strDate = VsfData.TextMatrix(lngStart, mlngDate)
                strTime = VsfData.TextMatrix(lngStart, mlngTime)
            End If
            
            '1\����
            strField = "ID|ҳ��|�к�|�к�|��¼ID|����|��λ|ɾ��"
            If mlngDate <> -1 Then
                strKey = mintҳ�� & "," & lngStart & "," & mlngDate
                strValue = strKey & "|" & mintҳ�� & "|" & lngStart & "|" & mlngDate & "|" & _
                    Val(VsfData.TextMatrix(lngStart, mlngRecord)) & "|" & strDate & "|" & VsfData.TextMatrix(lngStart, mlngDemo) & "|0"
                Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
            End If
            '2\ʱ��
            strKey = mintҳ�� & "," & lngStart & "," & mlngTime
            strValue = strKey & "|" & mintҳ�� & "|" & lngStart & "|" & mlngTime & "|" & _
                Val(VsfData.TextMatrix(lngStart, mlngRecord)) & "|" & strTime & "|" & _
                VsfData.TextMatrix(lngStart, mlngDemo) & "|0"
            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
         
            If Not blnGroup Or blnTrue Then
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
            Call SetActiveColColor
        End If
    End If
    
    '������������ʱ�����հ������������һ��
    If Left(Trim(strRows), 1) = "," Then strRows = Mid(Trim(strRows), 2)
    If Right(strRows, 1) = "," Then strRows = Mid(strRows, 1, Len(strRows) - 1)
    strRows = Replace("," & strRows & ",", ",,", "") '����ɾ��Ҫ׷�ӵ���
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
            VsfData.RowHidden(intRow) = True
            VsfData.RowPosition(intRow) = VsfData.Rows - 1
        End If
    Next intRow
   ' Call OutputRsData(mrsCellMap)

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
        If blnNewRow = False Then
        '����ţ�56592,����,����¼���������ת
              '������һ��
toMoveNextRow2:
            If VsfData.ROW < VsfData.Rows - 1 Then
                If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") <> 0 Then
                    intRow = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0))
                    intRow = intRow - Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(1)) + 1
                Else
                    intRow = 1
                End If
                If VsfData.ROW + intRow < VsfData.Rows Then
                    VsfData.ROW = VsfData.ROW + intRow
                Else
                    GoTo toMoveNextCol2
                End If
                If VsfData.RowHidden(VsfData.ROW) Then
                    If VsfData.ROW < VsfData.Rows - 1 Then
                        GoTo toMoveNextRow2
                    Else
                        For intRow = VsfData.Rows - 1 To VsfData.FixedRows Step -1
                            If VsfData.RowHidden(intRow) = False Then
                                VsfData.ROW = GetStartRow(intRow)
                                Exit For
                            End If
                        Next intRow
                    End If
                End If
            Else
toMoveNextCol2:
                If VsfData.COL < mlngNoEditor - 1 Then
                    VsfData.ROW = VsfData.FixedRows
                    VsfData.COL = VsfData.COL + 1
                    If VsfData.ColWidth(VsfData.COL) = 0 Or VsfData.ColHidden(VsfData.COL) Or _
                        InStr(1, "," & mstrCOLNothing & ",", "," & VsfData.COL - cHideCols & ",") <> 0 Then
                        GoTo toMoveNextCol2
                    End If
                Else
                    VsfData.ROW = VsfData.FixedRows
                    VsfData.COL = mlngNoEditor - 1
                End If
            End If
            
            
            If VsfData.ColIsVisible(VsfData.COL) = False Then
                VsfData.LeftCol = VsfData.COL
            End If
            If VsfData.RowIsVisible(VsfData.ROW) = False Then
                VsfData.TopRow = VsfData.ROW
            End If
        Else
toMoveNextCol:
            If VsfData.COL < mlngNoEditor - 1 Then       '�����¼���϶��л�ʿǩ����
                VsfData.COL = VsfData.COL + 1
                If VsfData.ColWidth(VsfData.COL) = 0 Or VsfData.ColHidden(VsfData.COL) Or _
                    InStr(1, "," & mstrCOLNothing & ",", "," & VsfData.COL - cHideCols & ",") <> 0 Then
                    GoTo toMoveNextCol
                End If
            Else
toMoveNextRow:
                '������һ��
                If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") <> 0 Then
                    intRow = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0))
                    intRow = intRow - Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(1)) + 1
                Else
                    intRow = 1
                End If
                If VsfData.ROW + intRow < VsfData.Rows Then
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
          
                VsfData.COL = IIf(mlngDate > 0, mlngDate, mlngTime)
            End If
        End If
    Else
toMovePrevCol:
        If VsfData.COL > mlngDate Then      '�����¼���϶��л�ʿǩ����
            VsfData.COL = VsfData.COL - 1
            If VsfData.ColWidth(VsfData.COL) = 0 Or VsfData.ColHidden(VsfData.COL) Then GoTo toMovePrevCol
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

Private Function GetStartRow(ByVal lngRow As Long) As Long
    Dim lngStart As Long
    Dim lngCurRows As Long, lngRows As Long
    '��ȡ������ʼ��,������ҳ�򷵻�0
    '�����ҳδ��ʾȫ,��˵��������ҳ,Ҳ����0
    '���������������������в�������

    If InStr(1, VsfData.TextMatrix(lngRow, mlngRowCount), "|") = 0 Then
        GetStartRow = lngRow
        Exit Function
    End If
    
    lngRows = Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(0))    '������
    lngCurRows = Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(1)) '��ǰ��
    If lngCurRows = 1 Then
        GetStartRow = lngRow
        Exit Function
    End If

    'Ѱ����ʼ��
    For lngRow = lngRow To 3 Step -1
        If Format(VsfData.TextMatrix(lngRow, mlngRowCount)) = lngRows & "|1" Then
            lngStart = lngRow
            Exit For
        End If
    Next

    GetStartRow = lngStart
End Function

Private Function GetStartRowHistory(ByVal lngRow As Long) As Long
    Dim lngStart As Long
    Dim lngCurRows As Long, lngRows As Long
    '��ȡ������ʼ��,������ҳ�򷵻�0
    '�����ҳδ��ʾȫ,��˵��������ҳ,Ҳ����0
    '���������������������в�������
    
    If InStr(1, vsfHistory.TextMatrix(lngRow, mlngRowCount), "|") = 0 Then
        GetStartRowHistory = lngRow
        Exit Function
    End If
    
    lngRows = Val(Split(vsfHistory.TextMatrix(lngRow, mlngRowCount), "|")(0))    '������
    lngCurRows = Val(Split(vsfHistory.TextMatrix(lngRow, mlngRowCount), "|")(1)) '��ǰ��
    If lngCurRows = 1 Then
        GetStartRowHistory = lngRow
        Exit Function
    End If

    'Ѱ����ʼ��
    For lngRow = lngRow To 3 Step -1
        If Format(vsfHistory.TextMatrix(lngRow, mlngRowCount)) = lngRows & "|1" Then
            lngStart = lngRow
            Exit For
        End If
    Next

    GetStartRowHistory = lngStart
End Function

Private Function GetMutilData(ByVal lngRow As Long, ByVal lngCol As Long, dblTop As Long, dblHeight As Long) As String
    Dim lngCurRow As Long
    Dim lngCount As Long
    Dim lngStart As Long    '��ʼ��
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
    lngCount = Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(0))
    lngCurRow = Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(1))

    If lngCount > 1 Then
        lngStart = GetStartRow(lngRow)
    Else
        lngStart = lngRow
    End If
    For lngRow = lngStart To lngStart + lngCount - 1
        strReturn = strReturn & VsfData.TextMatrix(lngRow, lngCol)
    Next
    
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

Private Function GetMutilDataHistory(ByVal lngRow As Long, ByVal lngCol As Long, dblTop As Long, dblHeight As Long) As String
    Dim lngCurRow As Long
    Dim lngCount As Long
    Dim lngStart As Long    '��ʼ��
    Dim strReturn As String
    Dim blnAdjust As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim lngOrder As Long
    Dim intCount As Integer, intBound As Integer
    On Error GoTo ErrHand
    '���ص�һ�е�����
    '������ֱ��ȡ������ʱ��������ҳ��ʾȫ��ƴ�ӣ�����ӿ��ж�ȡ
    mblnEditHistoryAssistant = False
    mblnEditHistoryAssistant = ISEditAssistant(lngCol)
   
    If vsfHistory.TextMatrix(lngRow, mlngRowCount) = "" Then
        GetMutilDataHistory = vsfHistory.TextMatrix(lngRow, lngCol)
        Exit Function
    End If
    
    '��ȡ������ʼ��
    lngRow = GetStartRowHistory(lngRow)
    lngStart = lngRow
    If Val(vsfHistory.TextMatrix(lngRow, mlngDemo)) > 0 And mblnEditHistoryAssistant Then 'ѡ����Ƿ�����ʼ��(���ı�)
        '��ȡ�������ݵĵ�һ��
        If Val(vsfHistory.TextMatrix(lngRow, mlngDemo)) > 1 Then
            lngStart = lngRow - Val(vsfHistory.TextMatrix(lngRow, mlngDemo)) + 1
            If Val(vsfHistory.TextMatrix(lngStart, mlngDemo)) <> 1 Then
                For lngStart = lngRow To vsfHistory.FixedRows Step -1
                    If Val(vsfHistory.TextMatrix(lngStart, mlngDemo)) = 1 Then
                        Exit For
                    End If
                Next lngStart
                If lngStart < vsfHistory.FixedRows Then lngStart = lngRow
            End If
        End If
        lngRow = lngStart
        lngCount = Val(Split(vsfHistory.TextMatrix(lngRow, mlngRowCount), "|")(0))
        intBound = lngRow + lngCount - 1
        
        For intCount = lngRow + lngCount To vsfHistory.Rows - 1
            If intCount > intBound Then
                If Val(vsfHistory.TextMatrix(intCount, mlngDemo)) <= 1 Then Exit For      '����������·�����˳�
                intBound = Val(Split(vsfHistory.TextMatrix(intCount, mlngRowCount), "|")(0)) + intCount - 1
            End If
            lngCount = lngCount + 1
        Next
    Else
        lngCount = Val(Split(vsfHistory.TextMatrix(lngRow, mlngRowCount), "|")(0))
    End If
    
    lngStart = lngRow
    strReturn = ""
    For lngRow = lngStart To lngStart + lngCount - 1
        strReturn = strReturn & vsfHistory.TextMatrix(lngRow, lngCol)
    Next
    strReturn = Replace(Replace(Replace(strReturn, Chr(10), ""), Chr(13), ""), Chr(1), "")
    'ȡ�и�
    vsfHistory.ROW = lngStart
    dblHeight = lngCount * vsfHistory.RowHeightMin + 20
    dblTop = vsfHistory.Top + vsfHistory.CellTop

    GetMutilDataHistory = strReturn
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ShowInput(Optional ByVal intCol As Integer = -1, Optional ByVal strCellData As String = "", Optional ByVal blnAnalyse As Boolean = False) As String
    Dim arrData, arrValue
    Dim lngOrder As Long
    Dim blnMake As Boolean
    Dim i As Integer, j As Integer, intPos As Integer, intIndex As Integer
    Dim strFormat As String, strText As String, strValue As String  '��ʽ��,���ݴ�,��ֵ��
    Dim strOrders As String, strTypes As String, strBounds As String
    Dim strLen As String, strName As String, strState As String
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
        strText = strCellData
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
    mintType = 0
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
            strName = strName & "," & mrsItems!��Ŀ����
            strLen = strLen & "," & mrsItems!��Ŀ���� & ";" & NVL(mrsItems!��ĿС��)
            strState = strState & "," & mrsItems!��Ŀ����
            strTypes = strTypes & "," & mrsItems!��Ŀ��ʾ
            strBounds = strBounds & "," & mrsItems!��Ŀֵ��
            strValue = strValue & "'" & SubstrVal(strText, strFormat, GetActivePart(intCol, intIndex) & mrsItems!��Ŀ����, intPos)

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
    '6=һ�а�2����������Ŀ,�ֹ�¼��
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
            .Width = .Width - (180 + IIf(mblnBlowup, 180 * 1 / 3, 0)) / 2 '����9��ʱ��ȥ90,����Խ��۳��ı߾�ԽС,�Ա�֤�ı��������ʵ��һ��
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
            Dim lngStart As Long
            blnMake = True
            If intCol = mlngDate Then
                If VsfData.ROW > VsfData.FixedRows Then
                    lngStart = GetStartRow(VsfData.ROW - 1)
                    blnMake = (VsfData.TextMatrix(lngStart, mlngDate) = "")
                End If
                If blnMake Then
                    If mblnDateAd Then
                        txtInput.Text = Format(zlDatabase.Currentdate, "d-M")
                        txtInput.Text = Replace(txtInput.Text, "-", "/")
                    Else
                        txtInput.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
                    End If
                Else
                    txtInput.Text = VsfData.TextMatrix(lngStart, mlngDate)
                End If
            ElseIf intCol = mlngTime Then
                If VsfData.ROW > VsfData.FixedRows Then
                    lngStart = GetStartRow(VsfData.ROW - 1)
                    blnMake = (VsfData.TextMatrix(lngStart, mlngTime) = "")
                End If
                If blnMake Then
                    txtInput.Text = Format(zlDatabase.Currentdate, "HH:mm")
                Else
                    txtInput.Text = VsfData.TextMatrix(lngStart, mlngTime)
                End If
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
                '��ѡ�ĵ�һ����Ŀ�����ѡ����Ҫ��������,��ѡ��Ŀ��ֱ�ӽ���
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
                If .Top + .Height <> PicLst.Height Then
                    PicLst.Height = .Top + .Height
                End If
                .Visible = True
            End With
        Else
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
        arrData = Split(Split(strBounds, ",")(1), ";")
        cboChoose(1).Tag = Split(strOrders, ",")(1)
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

Private Sub FillPage()
    Dim strPatient As String
    Dim rsTemp As New ADODB.Recordset
    Dim strCSQL As String
    On Error GoTo ErrHand
    '��ȡ���������Ĳ����嵥(��Ժ����+�������ת�Ʋ���+ָ��ʱ�䷶Χ�ڳ�Ժ����),�����嵥����������
    
    '58890:������,2013-02-26,��Ժ���˶�ȡ�����Ż�(������Ժ���˱���в�ѯ)
    '��Ժ�����嵥
    strPatient = "" & _
        " SELECT 1 AS ����,B.����ID, B.��ҳID, NVL(B.����,A.����) ����, NVL(B.�Ա�,A.�Ա�) �Ա�, B.סԺ��, lpad(B.��Ժ����,10,' ') AS ����,0 AS Ӥ��" & _
        " FROM ������Ϣ A,������ҳ B,��Ժ���� C" & _
        " Where A.����ID = B.����ID And A.��ҳID=B.��ҳID And NVL(b.��ҳID, 0) <> 0 " & _
        " And Nvl(B.״̬,0)<>1 AND Nvl(B.����״̬,0)<>5 AND B.���ʱ�� is NULL And A.����ID=C.����ID And C.����ID=[3] " & _
        IIf(mlng����ID = -1, "", " And C.����ID=[4]")
    If chk��Ժ.Value = 1 Then
        '��������Ժ�����嵥
        strPatient = strPatient & _
            " UNION " & _
            " SELECT 3 AS ����,B.����ID, B.��ҳID, NVL(B.����,A.����) ����, NVL(B.�Ա�,A.�Ա�) �Ա�, B.סԺ��, lpad(B.��Ժ����,10,' ') AS ����,0 AS Ӥ��" & _
            " FROM ������Ϣ A,������ҳ B" & _
            " Where A.����ID = b.����ID And NVL(b.��ҳID, 0) <> 0 And b.��ǰ����ID + 0 = [3]" & _
            " AND B.��Ժ���� BETWEEN [1] AND [2] AND Nvl(B.����״̬,0)<>5 AND B.���ʱ�� is NULL" & _
            IIf(mlng����ID = -1, "", " And B.��Ժ����ID+0=[4]")
    End If
    If chk����.Value = 1 Then
        '�������ת�Ʋ����嵥
        strPatient = strPatient & _
            " UNION " & _
            " Select 2 AS ����,B.����ID, B.��ҳID, NVL(B.����,A.����) ����, NVL(B.�Ա�,A.�Ա�) �Ա�, B.סԺ��, lpad(c.����,10,' ') AS ����,0 AS Ӥ��" & _
            " From ������Ϣ A,������ҳ B,���˱䶯��¼ C" & _
            " Where A.����ID=B.����ID And Nvl(B.��ҳID,0)<>0 " & _
            " And Nvl(B.״̬,0)<>2 And Nvl(C.���Ӵ�λ,0)=0 " & _
            " And B.����ID=C.����ID And B.��ҳID=C.��ҳID And C.����ID+0=[3]" & IIf(mlng����ID = -1, "", " And B.��Ժ����ID<>[4] And C.����ID+0=[4]") & _
            " And C.��ֹԭ��=3 And C.��ֹʱ�� Between Sysdate-" & mintChange & " And Sysdate" & _
            " And Nvl(B.����״̬,0)<>5 And B.���ʱ�� is NULL"
    End If
    '��ȡ�������б�
    strPatient = strPatient & _
              " UNION " & _
              " Select B.����,B.����ID,B.��ҳID,NVL(A.Ӥ������,B.����||'֮��'||A.���) AS ����,B.�Ա�,B.סԺ��,lpad(b.����,10,' ') AS ����,A.��� AS Ӥ��" & _
              " From ������������¼ A,(" & strPatient & ") B" & _
              " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID"
    
    If mstrBPItem = "" Then
        gstrSQL = " SELECT  A.����,A.����ID,A.��ҳID,A.Ӥ��,A.����,lpad(a.����,10,' ') AS ����,MAX(B.ID) AS �ļ�ID,'' ѪѹƵ��" & _
                  " FROM (" & strPatient & ") A,���˻����ļ� B" & _
                  " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID And A.Ӥ��=B.Ӥ�� " & _
                  " And B.�鵵�� is null And B.����ʱ�� is null And B.��ʽID=[5]" & _
                  " GROUP BY A.����,A.����ID,A.��ҳID,A.Ӥ��,A.���� ,A.����" & _
                  " Order by A.����,A.����"
    Else
        strCSQL = ",(Select ѪѹƵ��" & vbNewLine & _
                    "From (Select Distinct ����id, ��ҳid, NVL(Ӥ��,0) Ӥ��, First_Value(b.Ӣ������) Over(Partition By ����id, ��ҳid, NVL(Ӥ��,0) Order By ��ʼִ��ʱ�� Desc) ѪѹƵ��" & vbNewLine & _
                    "       From ����ҽ����¼ a, ����Ƶ����Ŀ b, (Select Column_Value From Table(f_Num2list('" & mstrBPItem & "'))) c" & vbNewLine & _
                    "       Where ((a.ҽ����Ч = 0 And a.ҽ��״̬ In (3, 5, 6, 7,8,9) And ((a.ִ����ֹʱ�� Is Null Or a.ִ����ֹʱ�� >= Sysdate))) Or (a.ҽ����Ч = 1 And a.ҽ��״̬ In (3, 5, 6, 7, 8)" & vbNewLine & _
                    "       And  a.��ʼִ��ʱ�� Between Sysdate-1 And Sysdate)) And a.ִ��Ƶ�� = b.���� And" & vbNewLine & _
                    "             a.������Ŀid = c.Column_Value) C Where A.����ID=C.����ID And A.��ҳID=C.��ҳID And A.Ӥ��=C.Ӥ��) ѪѹƵ��"

        gstrSQL = " SELECT /*+ Rule */  A.����,A.����ID,A.��ҳID,A.Ӥ��,A.����,lpad(a.����,10,' ') AS ����,MAX(B.ID) AS �ļ�ID" & strCSQL & _
                  " FROM (" & strPatient & ") A,���˻����ļ� B" & _
                  " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID And A.Ӥ��=B.Ӥ�� " & _
                  " And B.�鵵�� is null And B.����ʱ�� is null And B.��ʽID=[5]" & _
                  " GROUP BY A.����,A.����ID,A.��ҳID,A.Ӥ��,A.���� ,A.����" & _
                  " Order by A.����,A.����"
    End If
    
    
    Call SQLDIY(gstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����嵥", mdtOutbegin, mdtOutEnd, mlng����ID, mlng����ID, mlng��ʽID, mstrBPItem)
    
    '������ݵ����
    With rsTemp
        Do While Not .EOF
            If .AbsolutePosition > VsfData.Rows - VsfData.FixedRows Then VsfData.Rows = VsfData.Rows + 1
            
            VsfData.TextMatrix(.AbsolutePosition + VsfData.FixedRows - 1, c�ļ�ID) = !�ļ�ID
            VsfData.TextMatrix(.AbsolutePosition + VsfData.FixedRows - 1, c����) = Trim(NVL(!����))
            VsfData.TextMatrix(.AbsolutePosition + VsfData.FixedRows - 1, c����) = IIf(!Ӥ�� > 0, Space(4), "") & !����
            VsfData.TextMatrix(.AbsolutePosition + VsfData.FixedRows - 1, c����ID) = !����ID
            VsfData.TextMatrix(.AbsolutePosition + VsfData.FixedRows - 1, c��ҳID) = !��ҳID
            VsfData.TextMatrix(.AbsolutePosition + VsfData.FixedRows - 1, cӤ��) = !Ӥ��
            VsfData.TextMatrix(.AbsolutePosition + VsfData.FixedRows - 1, cѪѹƵ��) = NVL(!ѪѹƵ��)
            If mlngRowCount < VsfData.Cols Then VsfData.TextMatrix(.AbsolutePosition + VsfData.FixedRows - 1, mlngRowCount) = ""
            .MoveNext
        Loop
    End With
    
    If VsfData.Rows <= VsfData.FixedRows Then VsfData.Rows = VsfData.Rows + 1
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

'######################################################################################################################
'**********************************************************************************************************************
'�����ǻ������������
Private Sub lblDnInput_Click()
    txtDnInput.SetFocus
End Sub

Private Sub lblUpInput_Click()
    txtUpInput.SetFocus
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

Private Sub UserControl_GotFocus()
    On Error Resume Next
    VsfData.SetFocus
End Sub

Private Sub UserControl_Initialize()
    mblnShow = False
    mblnSigned = False
    mblnSaved = False
    mblnChange = False
    mblnInit = False

'    Set objStream = objFileSys.OpenTextFile("C:\WORKLOG.txt", ForAppending, True)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    '�����ַ���Ϊ���ݷָ�������¼�¼���ķָ�������˲�����¼��
    If KeyAscii = 39 Or KeyAscii = 13 Or KeyAscii = Asc("|") Then KeyAscii = 0: Exit Sub
    If KeyAscii = vbKeyEscape And mblnShow Then
        mblnShow = False
        Call InitCons
    End If
End Sub

Private Sub cbsThis_Resize()
    Dim lngScaleLeft As Long, lngScaleTop As Long, lngScaleRight As Long, lngScaleBottom As Long
    Call cbsThis.GetClientRect(lngScaleLeft, lngScaleTop, lngScaleRight, lngScaleBottom)

    Err = 0: On Error Resume Next
    If Not mblnInit Then picSplit.Top = ScaleHeight - 3000
    picSplit.Left = VsfData.Left
    picSplit.Width = VsfData.Width
    
    lblTitle.Move lngScaleLeft, 120, lngScaleRight - lngScaleLeft
    picMain.Move lngScaleLeft, lngScaleTop, lngScaleRight, picSplit.Top - lngScaleTop
    VsfData.Move lngScaleLeft + 210, lblTitle.Top + lblTitle.Height + 300, lngScaleRight - lngScaleLeft - 210 * 2
    VsfData.Height = picMain.Height - VsfData.Top
    
    vsfHistory.Left = VsfData.Left
    vsfHistory.Top = picSplit.Top + picSplit.Height
    vsfHistory.Height = lngScaleBottom - picSplit.Top - 50
    vsfHistory.Width = VsfData.Width
    
    picNull.Move lngScaleLeft, lngScaleTop, lngScaleRight, lngScaleBottom
    lblInfo(0).Move lngScaleLeft, (lngScaleBottom - lngScaleTop) / 2 - lblInfo(0).Height, lngScaleRight
    lblInfo(1).Move lngScaleLeft, (lngScaleBottom - lngScaleTop) / 2 + 100, lngScaleRight
    
    '���ϱ�ǩ��ɢ����
    Call zlLableBruit
End Sub

Private Sub UserControl_Terminate()
'    objStream.Close
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

Private Sub SetActiveColColor()
    '��еı���ɫ����Ϊ��ɫ,��ʾ������༭
    Dim aryItem, lngRow As Long
    aryItem = Split(mstrCOLNothing, ",")
    For lngRow = 0 To UBound(aryItem)
        VsfData.Cell(flexcpBackColor, VsfData.FixedRows, Val(aryItem(lngRow)) + cHideCols, VsfData.Rows - 1, Val(aryItem(lngRow)) + cHideCols) = &H8000000F
        '.ColHidden(Val(aryItem(lngCount)) + cHideCols) = True
    Next
End Sub

Private Sub vsfHistory_DblClick()
    Call vsfHistory_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub vsfHistory_DrawCell(ByVal hDC As Long, ByVal ROW As Long, ByVal COL As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    Call DrawCellHistory(hDC, ROW, COL, Left, Top, Right, Bottom, Done)
End Sub

Private Sub vsfHistory_KeyDown(KeyCode As Integer, Shift As Integer)
'----------------------------------------------
'�޸��ˣ�LPF 2012-04-20
'�޸����ݣ�����Ƿ�����ʼ�У�Ҳ����¼���������
'----------------------------------------------
    
    Dim arrData, i As Integer
    Dim blnNULL As Boolean                      '�Ƿ�Ϊ����
    Dim strReturn As String, strMsg As String, strPart As String
    Dim lngStart As Long, lngMutilRows As Long, lngDeff As Long, lngStartGroup As Long
    Dim intRow As Integer, intCount As Integer, intNULL As Integer, intBound As Integer, intRowCount As Integer  '����ж��ٿ���
    Dim dblTop As Long, dblHeight As Long, intRowGroup As Integer
    Dim blnGroup As Boolean, blnDate As Boolean, strRows As String, lngDemo As Long, lngLastNull As Long, lngLastNoNull As Long
    Dim strDate As String, strTime As String
    Dim strKey As String, strField As String, strValue As String, strCols As String
    Dim intGroupFirstRows As Integer, blnTrue As Boolean
    Dim blnReseGroupAssistant As Boolean, blnGroupAddNum As Boolean '��������������
    '���ı��к�������Ϣ
    Dim varAssistant() As Variant, strAssistantCols As String
    On Error GoTo ErrHand
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Not mblnInit Then Exit Sub
    '����ʷ���ǰ��Ԫ������ݸ��Ƶ���ǰ�����
    '��ֵȻ���ƶ�����һ����Ч��Ԫ��
    '��ȡ��ʷ����
    If vsfHistory.COL <= mlngTime Or vsfHistory.COL >= mlngNoEditor Then Exit Sub
    If Val(vsfHistory.TextMatrix(vsfHistory.ROW, mlngRecord)) = 0 Then Exit Sub
    
    cmdWord.Visible = False
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
    End Select
    mintType = -1: mblnShow = False: blnReseGroupAssistant = False
    
    strReturn = GetMutilDataHistory(vsfHistory.ROW, vsfHistory.COL, dblTop, dblHeight)
    If Val(VsfData.TextMatrix(VsfData.ROW, mlngRecord)) > 0 Then
        lngStart = GetStartRow(VsfData.ROW)
        For i = 0 To UBound(Split(mstrColCollect, "|"))
            strValue = GetRelatiionNo(CStr(Split(mstrColCollect, "|")(i)), 2)
            strCols = strCols & "," & IIf(strValue = "", "", strValue & ",") & Split(Split(mstrColCollect, "|")(i), ";")(0)
        Next
        strCols = Mid(strCols, 2)
        If InStr(1, "," & strCols & ",", "," & vsfHistory.COL - (cHideCols + vsfHistory.FixedCols - 1) & ",") > 0 Then
            If ISCollectSigned(Val(VsfData.TextMatrix(lngStart, c�ļ�ID)), Format(VsfData.TextMatrix(lngStart, mlngActiveTime), "YYYY-MM-DD"), Format(VsfData.TextMatrix(lngStart, mlngActiveTime), "HH:MM")) Then
                RaiseEvent AfterRowColChange("��Ҫճ��������Ϊ���������ݣ���Ŀ����������Ӧ�Ļ���������ǩ�������ݽ����ܱ�ճ����", True)
                Exit Sub
            End If
        End If
    End If
    mblnEditAssistant = mblnEditHistoryAssistant
    VsfData.COL = vsfHistory.COL
    If Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) > 0 And mblnEditHistoryAssistant Then 'ѡ����Ƿ�����ʼ��(���ı�)
        '��ȡ�������ݵĵ�һ��
        If Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) > 1 Then
            lngStart = VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1
            If Val(VsfData.TextMatrix(lngStart, mlngDemo)) <> 1 Then
                For lngStart = VsfData.ROW To VsfData.FixedRows Step -1
                    If Val(VsfData.TextMatrix(lngStart, mlngDemo)) = 1 Then
                        Exit For
                    End If
                Next lngStart
                If lngStart < VsfData.FixedRows Then lngStart = VsfData.ROW
            End If
            VsfData.ROW = lngStart
        End If
    End If
    blnDate = (InStr(1, "," & mlngDate & "," & mlngTime & ",", "," & VsfData.COL & ",") > 0)
    lngDemo = Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo))
    blnGroup = Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) >= 1
    If VsfData.TextMatrix(VsfData.ROW, mlngRowCount) = "" Then VsfData.TextMatrix(VsfData.ROW, mlngRowCount) = "1|1"
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
                        If VsfData.TextMatrix(intRow, intCount) <> "" And Not (IsDiagonal(intCount) And InStr(1, VsfData.TextMatrix(intRow, intCount), "/") <> 0) Then
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
    Else
        lngMutilRows = 1
        If VsfData.TextMatrix(VsfData.ROW, mlngDemo) = 1 And (mblnEditAssistant Or blnDate) Then
            '��¼������ʼ�е���������
            intGroupFirstRows = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0))
            intBound = VsfData.ROW + intGroupFirstRows - 1
            For intCount = VsfData.ROW + intGroupFirstRows To VsfData.Rows - 1
                '������ÿһ�������������ݿ���ռ�ö��У�ֻ���ڵ�һ�б����˷�����������������������=ÿһ������������+ÿһ���������е�����
                If intCount > intBound Then
                    If Val(VsfData.TextMatrix(intCount, mlngDemo)) <= 1 Then Exit For      '����������·�����˳�
                    intBound = Val(Split(VsfData.TextMatrix(intCount, mlngRowCount), "|")(0)) + intCount - 1
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
        End If
        lngStart = VsfData.ROW
    End If
    
    '׼����ֵ
    With txtLength
        '������ʱ���еĿ�Ȳ���,Ϊ�˱��ⷵ�ض���,ǿ������Ϊ5000
        .Width = IIf(VsfData.COL = mlngDate Or VsfData.COL = mlngTime, 5000, VsfData.CellWidth)
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
    blnGroupAddNum = False
    blnTrue = blnGroup = True And mblnEditAssistant
    If intCount > lngMutilRows - 1 Then
        '����������������ʱ������Ҫ��¼����������ݲ���¼����ı�������
'        If mblnEditAssistant = True And Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) = 1 And lngMutilRows = intGroupFirstRows Then
'            strMsg = "������������ʱ������������ݵķ��飬�����¼����Ķ���Ŀ���ݣ�"
'            RaiseEvent AfterRowColChange(strMsg, True)
'            strMsg = ""
'            Exit Sub
'        End If
        '������������,�������������������������ӵ�����
        '20110830���������ͬһ�����У��������ı��ֽ⵽���У�������ı�����ͳһ�������һ����;�ڷ����а��س�,ֻ���������ݽ����޸�,�����з����仯
        intNULL = intCount - (lngMutilRows - 1)
        For intRow = lngMutilRows To intCount
            '��֤��ǰ�����������һҳ����ʾȫ
            If intRow + lngStart > VsfData.Rows - 1 Then Exit For
            
            If Val(VsfData.TextMatrix(intRow + lngStart, c����ID)) = 0 And VsfData.TextMatrix(intRow + lngStart, mlngRowCount) = "" Then
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
            For intRow = lngMutilRows To intCount
                VsfData.TextMatrix(intRow + lngStart, c�ļ�ID) = VsfData.TextMatrix(lngStart, c�ļ�ID)
                VsfData.TextMatrix(intRow + lngStart, c����ID) = VsfData.TextMatrix(lngStart, c����ID)
                VsfData.TextMatrix(intRow + lngStart, c��ҳID) = VsfData.TextMatrix(lngStart, c��ҳID)
                VsfData.TextMatrix(intRow + lngStart, cӤ��) = VsfData.TextMatrix(lngStart, cӤ��)
            Next intRow
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
                    If Val(VsfData.TextMatrix(lngStart + intBound, mlngDemo)) > 0 Then
                        If CheckGroupDate(lngStart + intBound) = True Then
                            '�������޸ĲŽ�������̣�ȡ������¼��ʵ��ʱ��
                            If mblnDateAd Then
                                strDate = Format(VsfData.TextMatrix(lngStart + intBound, mlngActiveTime), "DD") & "/" & Format(VsfData.TextMatrix(lngStart + intBound, mlngActiveTime), "MM")
                            Else
                                strDate = Mid(VsfData.TextMatrix(lngStart + intBound, mlngActiveTime), 1, 10)
                            End If
                            strTime = Mid(VsfData.TextMatrix(lngStart + intBound, mlngActiveTime), 12, 5)
                        Else
                            '����ʱ���������
                            strDate = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngDate)
                            strTime = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngTime)
                        End If
                    Else
                        '��ͨ����
                        strDate = VsfData.TextMatrix(lngStart + intBound, mlngDate)
                        strTime = VsfData.TextMatrix(lngStart + intBound, mlngTime)
                    End If
                    
                    '1\����
                    strField = "ID|ҳ��|�к�|�к�|��¼ID|����|��λ|ɾ��"
                    If mlngDate <> -1 Then
                        strKey = mintҳ�� & "," & lngStart + intBound & "," & mlngDate
                        strValue = strKey & "|" & mintҳ�� & "|" & lngStart + intBound & "|" & mlngDate & "|" & _
                            Val(VsfData.TextMatrix(lngStart + intBound, mlngRecord)) & "|" & strDate & "|" & VsfData.TextMatrix(lngStart + intBound, mlngDemo) & "|" & IIf(VsfData.RowHidden(lngStart + intBound) = True, 1, 0)
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
                    If blnGroup And InStr(1, "," & mlngDemo & "," & mlngRecord & "," & mlngActiveTime & ",", "," & intCount & ",") = 0 Then
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
                    If Val(VsfData.TextMatrix(lngStart + intBound, mlngDemo)) > 0 Then
                        If CheckGroupDate(lngStart + intBound) = True Then
                            '�������޸ĲŽ�������̣�ȡ������¼��ʵ��ʱ��
                            If mblnDateAd Then
                                strDate = Format(VsfData.TextMatrix(lngStart + intBound, mlngActiveTime), "DD") & "/" & Format(VsfData.TextMatrix(lngStart + intBound, mlngActiveTime), "MM")
                            Else
                                strDate = Mid(VsfData.TextMatrix(lngStart + intBound, mlngActiveTime), 1, 10)
                            End If
                            strTime = Mid(VsfData.TextMatrix(lngStart + intBound, mlngActiveTime), 12, 5)
                        Else
                            '����ʱ���������
                            strDate = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngDate)
                            strTime = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngTime)
                        End If
                    Else
                        '��ͨ����
                        strDate = VsfData.TextMatrix(lngStart + intBound, mlngDate)
                        strTime = VsfData.TextMatrix(lngStart + intBound, mlngTime)
                    End If
                    
                    '1\����
                    strField = "ID|ҳ��|�к�|�к�|��¼ID|����|��λ|ɾ��"
                    If mlngDate <> -1 Then
                        strKey = mintҳ�� & "," & lngStart + intBound & "," & mlngDate
                        strValue = strKey & "|" & mintҳ�� & "|" & lngStart + intBound & "|" & mlngDate & "|" & _
                            Val(VsfData.TextMatrix(lngStart + intBound, mlngRecord)) & "|" & strDate & "|" & VsfData.TextMatrix(lngStart + intBound, mlngDemo) & "|" & IIf(VsfData.RowHidden(lngStart + intBound) = True, 1, 0)
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
                            If Val(VsfData.TextMatrix(intRowGroup, mlngDemo)) >= 0 Then
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
                    If CheckGroupDate(lngStart + intRow) = True Then
                        '�������޸ĲŽ�������̣�ȡ������¼��ʵ��ʱ��
                        If mblnDateAd Then
                            strDate = Format(VsfData.TextMatrix(lngStart + intRow, mlngActiveTime), "DD") & "/" & Format(VsfData.TextMatrix(lngStart + intRow, mlngActiveTime), "MM")
                        Else
                            strDate = Mid(VsfData.TextMatrix(lngStart + intRow, mlngActiveTime), 1, 10)
                        End If
                        strTime = Mid(VsfData.TextMatrix(lngStart + intRow, mlngActiveTime), 12, 5)
                    Else
                        '����ʱ���������
                        strDate = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngDate)
                        strTime = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngTime)
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
                            Val(VsfData.TextMatrix(lngStart + intRow, mlngRecord)) & "|" & strDate & "|" & VsfData.TextMatrix(lngStart + intRow, mlngDemo) & "|" & IIf(VsfData.RowHidden(lngStart + intRow) = True, 1, 0)
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
    mblnChange = True
           
    If Val(VsfData.TextMatrix(lngStart, mlngDemo)) > 0 Then
        If CheckGroupDate(lngStart) = True Then
            '�������޸ĲŽ�������̣�ȡ������¼��ʵ��ʱ��
            If mblnDateAd Then
                strDate = Format(VsfData.TextMatrix(lngStart, mlngActiveTime), "DD") & "/" & Format(VsfData.TextMatrix(lngStart, mlngActiveTime), "MM")
            Else
                strDate = Mid(VsfData.TextMatrix(lngStart, mlngActiveTime), 1, 10)
            End If
            strTime = Mid(VsfData.TextMatrix(lngStart, mlngActiveTime), 12, 5)
        Else
            '����ʱ���������
            strDate = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngDate)
            strTime = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngTime)
        End If
    Else
        '��ͨ����
        strDate = VsfData.TextMatrix(lngStart, mlngDate)
        strTime = VsfData.TextMatrix(lngStart, mlngTime)
    End If
    
    '1\����
    strField = "ID|ҳ��|�к�|�к�|��¼ID|����|��λ|ɾ��"
    If mlngDate <> -1 Then
        strKey = mintҳ�� & "," & lngStart & "," & mlngDate
        strValue = strKey & "|" & mintҳ�� & "|" & lngStart & "|" & mlngDate & "|" & _
            Val(VsfData.TextMatrix(lngStart, mlngRecord)) & "|" & strDate & "|" & VsfData.TextMatrix(lngStart, mlngDemo) & "|0"
        Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
    End If
    '2\ʱ��
    strKey = mintҳ�� & "," & lngStart & "," & mlngTime
    strValue = strKey & "|" & mintҳ�� & "|" & lngStart & "|" & mlngTime & "|" & _
        Val(VsfData.TextMatrix(lngStart, mlngRecord)) & "|" & strTime & "|" & _
        VsfData.TextMatrix(lngStart, mlngDemo) & "|0"
    Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
          
    If Not blnGroup Or blnTrue Then
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
    Call SetActiveColColor
    
    '������������ʱ�����հ������������һ��
    If Left(Trim(strRows), 1) = "," Then strRows = Mid(Trim(strRows), 2)
    If Right(strRows, 1) = "," Then strRows = Mid(strRows, 1, Len(strRows) - 1)
    strRows = Replace("," & strRows & ",", ",,", "")
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
            VsfData.RowHidden(intRow) = True
            VsfData.RowPosition(intRow) = VsfData.Rows - 1
        End If
    Next intRow

    '������֯������������
    If blnReseGroupAssistant = True Then
        If blnGroupAddNum = True Then Call GetGroupAssistant(strAssistantCols, varAssistant)
        If strAssistantCols <> "" Then
            Call ReSetGroupAssistant(True, False, strAssistantCols, varAssistant)
        Else
            Call ReSetGroupDemo(lngStart)
        End If
    End If
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

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
    Dim strDate As String, strTime As String
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
            If CheckGroupDate(lngRow) = True Then
                '�������޸ĲŽ�������̣�ȡ������¼��ʵ��ʱ��
                If mblnDateAd Then
                    strDate = Format(VsfData.TextMatrix(lngRow, mlngActiveTime), "DD") & "/" & Format(VsfData.TextMatrix(lngRow, mlngActiveTime), "MM")
                Else
                    strDate = Mid(VsfData.TextMatrix(lngRow, mlngActiveTime), 1, 10)
                End If
                strTime = Mid(VsfData.TextMatrix(lngRow, mlngActiveTime), 12, 5)
            Else
                '����ʱ���������
                strDate = VsfData.TextMatrix(lngRow - Val(VsfData.TextMatrix(lngRow, mlngDemo)) + 1, mlngDate)
                strTime = VsfData.TextMatrix(lngRow - Val(VsfData.TextMatrix(lngRow, mlngDemo)) + 1, mlngTime)
            End If
            
            '1\����
            strField = "ID|ҳ��|�к�|�к�|��¼ID|����|��λ|ɾ��"
            If mlngDate <> -1 Then
                strKey = mintҳ�� & "," & lngRow & "," & mlngDate
                strValue = strKey & "|" & mintҳ�� & "|" & lngRow & "|" & mlngDate & "|" & _
                    Val(VsfData.TextMatrix(lngRow, mlngRecord)) & "|" & strDate & "|" & VsfData.TextMatrix(lngRow, mlngDemo) & "|" & IIf(VsfData.RowHidden(lngRow) = True, 1, 0)
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

Private Sub CellMap_UpdateAssistant(ByVal lngStartRow As Long)
'���ܣ����¼�¼�����Ķ���Ϣ
    Dim strDate As String, strTime As String
    Dim strKey As String, strField As String, strValue As String, strPart As String
    Dim lngCol As Long, lngRow As Long, lngRowCount As Long, strReturn As String
    
    On Error GoTo ErrHand
    
    If VsfData.TextMatrix(lngStartRow, mlngRowCount) = "" Then VsfData.TextMatrix(lngStartRow, mlngRowCount) = "1|1"
    lngRowCount = Val(Split(VsfData.TextMatrix(lngStartRow, mlngRowCount), "|")(0))
    
    If Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) > 0 Then
        If CheckGroupDate(lngStartRow) = True Then
            '�������޸ĲŽ�������̣�ȡ������¼��ʵ��ʱ��
            If mblnDateAd Then
                strDate = Format(VsfData.TextMatrix(lngStartRow, mlngActiveTime), "DD") & "/" & Format(VsfData.TextMatrix(lngStartRow, mlngActiveTime), "MM")
            Else
                strDate = Mid(VsfData.TextMatrix(lngStartRow, mlngActiveTime), 1, 10)
            End If
            strTime = Mid(VsfData.TextMatrix(lngStartRow, mlngActiveTime), 12, 5)
        Else
            '����ʱ���������
            strDate = VsfData.TextMatrix(lngStartRow - Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) + 1, mlngDate)
            strTime = VsfData.TextMatrix(lngStartRow - Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) + 1, mlngTime)
        End If
    Else
        '��ͨ����
        strDate = VsfData.TextMatrix(lngStartRow, mlngDate)
        strTime = VsfData.TextMatrix(lngStartRow, mlngTime)
    End If
    
    '1\����
    strField = "ID|ҳ��|�к�|�к�|��¼ID|����|��λ|ɾ��"
    If mlngDate <> -1 Then
        strKey = mintҳ�� & "," & lngStartRow & "," & mlngDate
        strValue = strKey & "|" & mintҳ�� & "|" & lngStartRow & "|" & mlngDate & "|" & _
            Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "|" & strDate & "|" & VsfData.TextMatrix(lngStartRow, mlngDemo) & "|0"
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

Private Function CheckGroupDate(ByVal lngRow As Long) As Boolean
'--���ܣ�������������ʼ��ʱ��ͱ���ʱ���Ƿ����
    Dim strDate As String, strTime As String
    Dim strDate1 As String, strTime1 As String
    Dim lngStart As Long
    
    lngStart = lngRow - Val(VsfData.TextMatrix(lngRow, mlngDemo)) + 1
    
    If Val(VsfData.TextMatrix(lngRow, mlngRecord)) > 0 Then
        If Val(VsfData.TextMatrix(lngStart, mlngDemo)) <> 1 Then CheckGroupDate = True: Exit Function
        strDate = VsfData.TextMatrix(lngStart, mlngDate)
        strTime = VsfData.TextMatrix(lngStart, mlngTime)
        If mblnDateAd Then
            strDate1 = Format(VsfData.TextMatrix(lngStart, mlngActiveTime), "DD") & "/" & Format(VsfData.TextMatrix(lngStart, mlngActiveTime), "MM")
        Else
            strDate1 = Mid(VsfData.TextMatrix(lngStart, mlngActiveTime), 1, 10)
        End If
        strTime1 = Mid(VsfData.TextMatrix(lngStart, mlngActiveTime), 12, 5)
        If strDate <> strDate1 Or strTime <> strTime1 Then
            CheckGroupDate = False
        Else
            CheckGroupDate = True
        End If
    Else
        CheckGroupDate = False
    End If
End Function

Private Function ISGroupAppend() As Boolean
'׷�ӷ������ݣ���ѡ����������ݲ���׷�ӣ����������ı���Ŀ��
    Dim lngCol As Long, lngRow As Long
    Dim blnNULL As Boolean
    
    lngRow = VsfData.ROW
    If lngRow > VsfData.Rows - 1 Then lngRow = VsfData.Rows - 1
    blnNULL = True
    For lngCol = mlngTime + 1 To VsfData.Cols - 1
        If Not VsfData.ColHidden(lngCol) And lngCol < mlngNoEditor And ISEditAssistant(lngCol) = False Then
            If VsfData.TextMatrix(lngRow, lngCol) <> "" And Not (IsDiagonal(lngCol) And InStr(1, VsfData.TextMatrix(lngRow, lngCol), "/") <> 0) Then
                blnNULL = False
                Exit For
            End If
        End If
    Next
    
    ISGroupAppend = Not blnNULL
End Function

Private Sub SingerShowType(ByVal vsfObj As VSFlexGrid, ByVal lngStartRow As Long, ByVal lngEndRow As Long, Optional ByVal blnUnSingMe As Boolean = False)
'-------------------------------------------------
'���ܣ���ʿǩ������ʾ��ʽ
''--58414,������,2013-04-10,��ӻ�ʿ��ǩ������ʾģʽ
'-------------------------------------------------
    Dim lngRow As Integer
    'ȡ��ǩ��
    If blnUnSingMe = True Then
        For lngRow = lngStartRow To lngEndRow
            If mlngSignName > 0 Then vsfObj.TextMatrix(lngRow, mlngSignName) = ""
            If mlngSignTime > 0 Then vsfObj.TextMatrix(lngRow, mlngSignTime) = ""
        Next
        Exit Sub
    End If
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
    arrItem = Split(mstrColCorrelative, "|")
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
    Set rsImpAmount = frmImportOrder.ShowMe(Me, Val(VsfData.TextMatrix(VsfData.ROW, c�ļ�ID)), Val(VsfData.TextMatrix(VsfData.ROW, c����ID)), Val(VsfData.TextMatrix(VsfData.ROW, c��ҳID)), Val(VsfData.TextMatrix(VsfData.ROW, cӤ��)), blnImportName, lngNumOrder, strDate)
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


