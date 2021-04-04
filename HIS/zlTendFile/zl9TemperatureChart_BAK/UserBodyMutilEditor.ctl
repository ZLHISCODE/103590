VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.Unicode.9600.ocx"
Begin VB.UserControl UserBodyMutilEditor 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   8220
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12090
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   8220
   ScaleWidth      =   12090
   Begin VB.PictureBox picFilter 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1725
      Left            =   4800
      ScaleHeight     =   1695
      ScaleWidth      =   2115
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   465
      Visible         =   0   'False
      Width           =   2145
      Begin VB.ListBox lstFilter 
         Appearance      =   0  'Flat
         Height          =   1290
         Left            =   -15
         Style           =   1  'Checkbox
         TabIndex        =   49
         Top             =   -15
         Width           =   2145
      End
      Begin VB.CommandButton cmdFilterCancel 
         Height          =   315
         Left            =   1530
         Picture         =   "UserBodyMutilEditor.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "ȡ��"
         Top             =   1320
         Width           =   450
      End
      Begin VB.CommandButton cmdFilterOK 
         Height          =   315
         Left            =   990
         Picture         =   "UserBodyMutilEditor.ctx":058A
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "ȷ��"
         Top             =   1320
         Width           =   450
      End
   End
   Begin VB.PictureBox picPati 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6855
      Left            =   6900
      ScaleHeight     =   6825
      ScaleWidth      =   5145
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   435
      Visible         =   0   'False
      Width           =   5175
      Begin XtremeReportControl.ReportControl rptPati 
         Height          =   5370
         Left            =   0
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   915
         Width           =   5160
         _Version        =   589884
         _ExtentX        =   9102
         _ExtentY        =   9472
         _StockProps     =   0
         BorderStyle     =   1
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin VB.CommandButton cmdFilterUserOk 
         Height          =   315
         Left            =   3990
         Picture         =   "UserBodyMutilEditor.ctx":0B14
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "ȷ��"
         Top             =   6435
         Width           =   450
      End
      Begin VB.CommandButton cmdFilterUserCancle 
         Height          =   315
         Left            =   4530
         Picture         =   "UserBodyMutilEditor.ctx":109E
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "ȡ��"
         Top             =   6435
         Width           =   450
      End
      Begin VB.CheckBox chkSwitch 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   240
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   6510
         Width           =   195
      End
      Begin VB.CheckBox chkPati 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "���˱���"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   705
         TabIndex        =   37
         Top             =   6510
         Width           =   1095
      End
      Begin VB.CheckBox chkPati 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Ӥ��"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   1905
         TabIndex        =   36
         Top             =   6510
         Width           =   735
      End
      Begin VB.CheckBox chkScope 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "�����"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   35
         Top             =   135
         Width           =   915
      End
      Begin VB.CheckBox chkScope 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "��Ժ"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   1147
         TabIndex        =   34
         Top             =   135
         Width           =   750
      End
      Begin VB.CheckBox chkScope 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ת��"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   3
         Left            =   1980
         TabIndex        =   33
         Top             =   135
         Width           =   735
      End
      Begin VB.CheckBox chkScope 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "��Ժ"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   2
         Left            =   150
         TabIndex        =   32
         Top             =   510
         Width           =   795
      End
      Begin VB.TextBox txtChange 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         IMEMode         =   3  'DISABLE
         Left            =   3600
         MaxLength       =   3
         TabIndex        =   31
         Text            =   "7"
         Top             =   120
         Width           =   285
      End
      Begin VB.Frame fraChange 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   3570
         TabIndex        =   30
         Top             =   315
         Width           =   300
      End
      Begin VB.CommandButton CmdRef 
         Caption         =   "ˢ��"
         Height          =   315
         Left            =   4485
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "ȡ��"
         Top             =   480
         Width           =   555
      End
      Begin MSComCtl2.DTPicker dtpE 
         Height          =   300
         Index           =   0
         Left            =   3045
         TabIndex        =   41
         Top             =   480
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   127795203
         CurrentDate     =   37068
      End
      Begin MSComCtl2.DTPicker dtpB 
         Height          =   300
         Index           =   0
         Left            =   1380
         TabIndex        =   42
         Top             =   465
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   127795203
         CurrentDate     =   37068
      End
      Begin VB.Label lbl��Ժʱ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ʱ��"
         Height          =   180
         Left            =   945
         TabIndex        =   45
         Top             =   510
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Index           =   0
         Left            =   2760
         TabIndex        =   44
         Top             =   525
         Width           =   180
      End
      Begin VB.Label lblת�� 
         AutoSize        =   -1  'True
         Caption         =   "��ʾ���    ���ת������"
         Height          =   180
         Left            =   2820
         TabIndex        =   43
         Top             =   105
         Width           =   2160
      End
   End
   Begin VB.PictureBox picNull 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   870
      Left            =   2250
      ScaleHeight     =   840
      ScaleWidth      =   7530
      TabIndex        =   50
      Top             =   3105
      Visible         =   0   'False
      Width           =   7560
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "�������˻���Ӳ��˽����������"
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
         Left            =   0
         TabIndex        =   51
         Top             =   120
         Width           =   6960
      End
   End
   Begin VB.PictureBox picSplit 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   375
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   6255
      TabIndex        =   57
      Top             =   5460
      Width           =   6255
   End
   Begin MSComctlLib.ImageList imgRPT 
      Left            =   11400
      Top             =   480
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
            Picture         =   "UserBodyMutilEditor.ctx":1628
            Key             =   "woman"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserBodyMutilEditor.ctx":7E8A
            Key             =   "man"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   11400
      Top             =   120
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
            Picture         =   "UserBodyMutilEditor.ctx":E6EC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pic�������� 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   480
      ScaleHeight     =   345
      ScaleWidth      =   11130
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   11130
      Begin VB.ComboBox cboPati 
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   7380
         Style           =   2  'Dropdown List
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   0
         Width           =   1185
      End
      Begin VB.CommandButton cmdSift 
         Appearance      =   0  'Flat
         Height          =   260
         Left            =   6560
         Picture         =   "UserBodyMutilEditor.ctx":EA86
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "ѡ����Ŀ(F4)"
         Top             =   10
         Width           =   270
      End
      Begin VB.TextBox txtFilter 
         Height          =   300
         Left            =   4810
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   0
         Width           =   2040
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Left            =   840
         TabIndex        =   3
         Top             =   0
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   127795203
         CurrentDate     =   40624
      End
      Begin VB.CommandButton cmdAddUser 
         Caption         =   "��Ӳ���(&A)"
         Height          =   315
         Left            =   9720
         TabIndex        =   10
         Top             =   -15
         Width           =   1245
      End
      Begin VB.ComboBox cboUnit 
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   2820
         Style           =   2  'Dropdown List
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   0
         Width           =   1425
      End
      Begin VB.CommandButton cmdFilter 
         Caption         =   "����(&F)"
         Height          =   315
         Left            =   8640
         TabIndex        =   9
         Top             =   0
         Width           =   885
      End
      Begin VB.Label lblPati 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   6960
         TabIndex        =   25
         Top             =   60
         Width           =   360
      End
      Begin VB.Label lblFilter 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   4380
         TabIndex        =   6
         Top             =   60
         Width           =   360
      End
      Begin VB.Label lblDate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����ʱ��"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   60
         TabIndex        =   2
         Top             =   60
         Width           =   720
      End
      Begin VB.Label lblUnit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   2400
         TabIndex        =   4
         Top             =   60
         Width           =   360
      End
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4500
      Left            =   0
      ScaleHeight     =   4500
      ScaleWidth      =   4845
      TabIndex        =   1
      Top             =   360
      Width           =   4845
      Begin VB.PictureBox PicLst 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H80000008&
         Height          =   1410
         Left            =   1080
         ScaleHeight     =   1380
         ScaleWidth      =   1185
         TabIndex        =   19
         Top             =   3150
         Visible         =   0   'False
         Width           =   1215
         Begin VB.ListBox lstSelect 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            Height          =   570
            Index           =   0
            ItemData        =   "UserBodyMutilEditor.ctx":EB7C
            Left            =   -10
            List            =   "UserBodyMutilEditor.ctx":EB89
            TabIndex        =   21
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
            TabIndex        =   20
            Top             =   255
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
            TabIndex        =   60
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
            TabIndex        =   59
            Top             =   615
            Width           =   540
         End
      End
      Begin VB.ComboBox cbo���±�ʶ 
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   3600
         Style           =   2  'Dropdown List
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   3300
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.ListBox lstSelect 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1080
         Index           =   1
         ItemData        =   "UserBodyMutilEditor.ctx":EBA2
         Left            =   2280
         List            =   "UserBodyMutilEditor.ctx":EBAF
         Style           =   1  'Checkbox
         TabIndex        =   22
         Top             =   3120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.PictureBox picDouble 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1080
         ScaleHeight     =   240
         ScaleWidth      =   900
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   2880
         Visible         =   0   'False
         Width           =   930
         Begin VB.TextBox txtDnInput 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   525
            MaxLength       =   12
            TabIndex        =   17
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
            TabIndex        =   16
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
      Begin VB.ListBox lstNote 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   930
         ItemData        =   "UserBodyMutilEditor.ctx":EBC8
         Left            =   120
         List            =   "UserBodyMutilEditor.ctx":EBD2
         TabIndex        =   14
         Top             =   3120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.PictureBox picInput 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000018&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   150
         ScaleHeight     =   225
         ScaleWidth      =   945
         TabIndex        =   11
         Top             =   2880
         Visible         =   0   'False
         Width           =   975
         Begin VB.TextBox txtInput 
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   270
            Left            =   0
            TabIndex        =   12
            Top             =   0
            Width           =   180
         End
         Begin VB.Label lblCheck 
            Alignment       =   2  'Center
            BackColor       =   &H80000018&
            Height          =   135
            Left            =   240
            TabIndex        =   13
            Top             =   0
            Visible         =   0   'False
            Width           =   135
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid VsfData 
         Height          =   2655
         Left            =   0
         TabIndex        =   23
         Top             =   480
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
         AllowUserResizing=   1
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
         FormatString    =   $"UserBodyMutilEditor.ctx":EBE2
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
   End
   Begin VB.PictureBox picHistory 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2115
      Left            =   360
      ScaleHeight     =   2115
      ScaleWidth      =   5625
      TabIndex        =   52
      Top             =   5655
      Width           =   5625
      Begin VB.PictureBox picDay 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   105
         ScaleHeight     =   345
         ScaleWidth      =   4470
         TabIndex        =   58
         Top             =   165
         Width           =   4470
         Begin VB.TextBox txt��ʾ���� 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   855
            MaxLength       =   2
            TabIndex        =   54
            TabStop         =   0   'False
            Text            =   "1"
            Top             =   0
            Width           =   645
         End
         Begin VB.Label lblDayInfo 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Height          =   180
            Left            =   1695
            TabIndex        =   55
            Top             =   60
            Width           =   90
         End
         Begin VB.Label lblDay 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "��ʾ����"
            Height          =   180
            Left            =   60
            TabIndex        =   53
            Top             =   60
            Width           =   720
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfHistory 
         Height          =   1695
         Left            =   45
         TabIndex        =   56
         Top             =   600
         Width           =   4305
         _cx             =   7594
         _cy             =   2990
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
         AllowUserResizing=   1
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
         FormatString    =   $"UserBodyMutilEditor.ctx":EC44
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
Attribute VB_Name = "UserBodyMutilEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Enum PATI_COLUMN
    c_ѡ�� = 0
    c_ͼ�� = 1
    c_���� = 2
    c_״̬ = 3
    c_���� = 4
    C_����ID = 5
    c_��ҳID = 6
    c_���� = 7
    c_���� = 8
    c_סԺ�� = 9
    c_��Ժ���� = 10
    c_��Ժ���� = 11
End Enum

Private Enum Scope
    ����� = 0
    ��Ժ = 1
    ��Ժ = 2
    ת�� = 3
End Enum

Private Const c�ļ�ID As Integer = 1
Private Const c���� As Integer = 2
Private Const c���� As Integer = 3
Private Const c���� As Integer = 4
Private Const c����ID As Integer = 5
Private Const c��ҳID As Integer = 6
Private Const cӤ�� As Integer = 7
Private Const c��¼ID As Integer = 8
Private Const c����ȼ� As Integer = 9
Private Const c���±�ʶ As Integer = 10
Private Const c��Ժ As Integer = 11
Private Const c���� As Integer = 12
Private Const cʱ�� As Integer = 13
Private Const RootCol As Integer = 14  '�̶���ͷ����

Private mcbrMenuBar��λ As CommandBarControl
Private mcbrToolBar As CommandBar

'---���˻�����Ϣ
Private mlng�ļ�ID As Long
Private mlng����ID As Long
Private mlng��ҳID As Long
Private mlngBaby As Long

Private mrsItems As New ADODB.Recordset
Private mrsCell As New ADODB.Recordset
Private mrsPati As New ADODB.Recordset
Private mrsPart As New ADODB.Recordset
Private mrsCopy As New ADODB.Recordset
Private mrsData As New ADODB.Recordset
Private mrsHistory As New ADODB.Recordset

Private mstrSQL As String
Private mfrmParent As Object
Private mlng����ID As Long
Private mlng����id As Long '�û�ѡ��Ŀ���ID
Private mlng��ʽID As Long '���µ���ʽID
Private mstrDate As String '�û�ѡ���ʱ��
Private mblnInit As Boolean
Private mstrPrivs As String
Private mintBigSize As Integer '�����ļ���ʾģʽ
Private mintPreDays As Integer '����¼������
Private mlngHours As Long    '���ݲ�¼ʱ��
Private mstrScope As String  '����������ʾ��Χ
Private mintChange As Integer '�������ת������
Private mdtOutEnd As String '������Ժ��ʾ��ֹʱ��
Private mdtOutBegin As String '������Ժ��ʾ��ʼʱ��
Private mblnShow As Boolean
Private mblnChage As Boolean
Private mblnNullRow As Boolean
Private mblnClearRow As Boolean
Private mblnRefreshData As Boolean
Private mbln��Ժ As Boolean
Private mblnSaveData As Boolean
Private mblnDateFouces As Boolean
Private mblnChkClick As Boolean
Private mstrTabHead As String ' ��ͷ��Ϣ
Private mstrItemNo As String '��Ŀ�����Ϣ
Private mintPatiNo As Integer '�������� (���С����ˡ�Ӥ��)
Private mint����Ӧ�� As Integer
Private mstrNote As String '��������δ��˵����Ϣ
Private mintType As Integer
Private mstrModifyTime As String
Private mint������Դ As Integer, mintModify As Integer
Private mstr�������� As String
Private mbln����������ʾ As Boolean

Public Event AfterRowColChange(ByVal strInfo As String, ByVal blnImportant As Boolean)
Public Event UsrHelp()
Public Event UsrExit()

Public Function ShowMe(ByVal frmParent As Form, ByVal lngDeptID As Long, Optional ByVal strPrivs As String, Optional ByVal bytSize As Byte = 0) As Boolean
    '******************************************************************************************************************
    '���ܣ� ��ʾ���µ�����
    '������ frmParent           �ϼ��������
    '       lngDeptID           Ҫ��ʾ�����¼�Ŀ���
    '���أ� ��
    '******************************************************************************************************************
    Dim lngRow As Long
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo Errhand
    Err = 0

    mblnInit = False
    mlng����ID = lngDeptID
    mstrPrivs = strPrivs
    mintBigSize = bytSize ' zlDatabase.GetPara("�����ļ���ʾģʽ", glngSys, 1255, 0)
    Set mfrmParent = frmParent

    mintPreDays = Val(zlDatabase.GetPara("����¼�뻤����������", glngSys, 1255, "1"))
    mbln����������ʾ = (Val(zlDatabase.GetPara("���������(����/����)��ʽ¼��", glngSys, 1255, "1")) = 1)
    
    If mrsItems.State = 0 Then
        Call InitMenuBar
        Call InitEnv            '��ʼ������
        cbsThis.ActiveMenuBar.Visible = False
        cbsThis.RecalcLayout
    End If
    Call GetLocalSetting '��ע����ж�ȡ����������
    Call InitCons
    Call InitVariable
    
    If cboUnit.ListCount = 0 Then
        MsgBox "�������ڵ�ǰ�������κο��ң�����ʹ�øù��ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    Call ReSetFontSize
    ShowMe = True
    Exit Function
Errhand:
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
    bytFontSize = IIf(mintBigSize = 0, 9, IIf(mintBigSize = 1, 12, mintBigSize))
    
    UserControl.FontSize = bytFontSize
    UserControl.FontName = "����"
    For Each objCtrl In UserControl.Controls
        Select Case UCase(TypeName(objCtrl))
        Case UCase("Label")
            If UCase(objCtrl.Name) <> UCase("lblInfo") Then
            objCtrl.FontSize = bytFontSize
            objCtrl.Height = TextHeight("��") + 20
            End If
        Case UCase("ListBox")
            objCtrl.FontSize = bytFontSize
        Case UCase("VsFlexGrid")
            objCtrl.FontSize = bytFontSize
        Case UCase("ComboBox")
            objCtrl.FontSize = bytFontSize
        Case UCase("OptionButton")
            objCtrl.FontSize = bytFontSize
            objCtrl.Width = TextWidth("����" & objCtrl.Caption) - TextWidth("��") / 3
        Case UCase("CheckBox")
            objCtrl.FontSize = bytFontSize
            If UCase(objCtrl.Name) <> UCase("chkSwitch") Then
                objCtrl.Width = TextWidth("����" & objCtrl.Caption) - TextWidth("��") / 3
            End If
        Case UCase("DTPicker")
            objCtrl.Font.Size = bytFontSize
            objCtrl.Width = TextWidth("2012-01-01") + 400
            objCtrl.Height = TextHeight("��") * 1.5
        Case UCase("TextBox")
          objCtrl.FontSize = bytFontSize
        Case UCase("ReportControl")
            Set CtlFont = objCtrl.PaintManager.CaptionFont
            CtlFont.Size = bytFontSize
            Set objCtrl.PaintManager.CaptionFont = CtlFont
            
            Set CtlFont = objCtrl.PaintManager.TextFont
            CtlFont.Size = bytFontSize
            Set objCtrl.PaintManager.TextFont = CtlFont
            objCtrl.Redraw
        Case UCase("DockingPane")
            Set CtlFont = objCtrl.PaintManager.CaptionFont
            If CtlFont Is Nothing Then
                Set CtlFont = UserControl.Font
            End If
            CtlFont.Size = bytFontSize
            Set objCtrl.PaintManager.CaptionFont = CtlFont
        Case UCase("CommandBars")
            Set CtlFont = objCtrl.Options.Font
            If CtlFont Is Nothing Then
                Set CtlFont = UserControl.Font
            End If
            CtlFont.Size = bytFontSize
            Set objCtrl.Options.Font = CtlFont
        Case UCase("TabControl")
            Set CtlFont = objCtrl.PaintManager.Font
            If CtlFont Is Nothing Then
                Set CtlFont = UserControl.Font
            End If
            CtlFont.Size = bytFontSize
            Set objCtrl.PaintManager.Font = CtlFont
        Case UCase("CommandButton")
            objCtrl.FontSize = bytFontSize
            objCtrl.Width = TextWidth(" " & IIf(objCtrl.Caption = "", "  ", objCtrl.Caption) & " ")
        End Select
    Next
    
    '�ƶ��ؼ�λ��
    lblDate.Left = 60
    dtpDate.Left = lblDate.Left + lblDate.Width + TextWidth("��") / 2
    lblUnit.Left = dtpDate.Left + dtpDate.Width + TextWidth("��")
    cboUnit.Left = lblUnit.Left + lblUnit.Width + TextWidth("��") / 2
    lblFilter.Left = cboUnit.Left + cboUnit.Width + TextWidth("��")
    txtFilter.Left = lblFilter.Left + lblFilter.Width + TextWidth("��") / 2
    cmdSift.Height = txtFilter.Height - TextHeight("��") / 4
    cmdSift.Width = TextWidth("��") + TextWidth("��") / 2
    cmdSift.Left = txtFilter.Left + txtFilter.Width - cmdSift.Width
    lblPati.Left = txtFilter.Left + txtFilter.Width + TextWidth("��")
    cboPati.Left = lblPati.Left + lblPati.Width + TextWidth("��") / 2
    cmdFilter.Left = IIf(lblPati.Visible = True, cboPati.Left + cboPati.Width + (TextWidth("��") / 2) + 15, lblPati.Left)
    cmdFilter.Top = cboUnit.Top
    cmdFilter.Height = TextHeight("��") + 100
    cmdAddUser.Left = cmdFilter.Left + cmdFilter.Width + TextWidth("��") + 75
    cmdAddUser.Top = cmdFilter.Top
    cmdAddUser.Height = cmdFilter.Height
    pic��������.Width = cmdAddUser.Left + cmdAddUser.Width + TextWidth("��") / 2
    
    txt��ʾ����.Left = lblDay.Left + lblDay.Width + TextWidth("��")
    lblDay.Top = txt��ʾ����.Top + (txt��ʾ����.Height - lblDay.Height) \ 2
    lblDayInfo.Top = lblDate.Top
    lblDayInfo.Left = txt��ʾ����.Left + txt��ʾ����.Width + TextWidth("��")
    
    
    '��Ӳ����ƶ�
    chkScope(0).Left = 150
    chkScope(1).Left = chkScope(0).Left + chkScope(0).Width + TextWidth("��") / 2
    chkScope(1).Left = chkScope(0).Left + chkScope(0).Width + TextWidth("��") / 2
    chkScope(3).Left = chkScope(1).Left + chkScope(1).Width + TextWidth("��") / 2
    
    chkScope(2).Left = chkScope(0).Left
    lbl��Ժʱ��.Left = chkScope(2).Left + chkScope(2).Width + TextWidth("��") / 2
    lbl��Ժʱ��.Top = chkScope(2).Top
    dtpB(0).Left = lbl��Ժʱ��.Left + lbl��Ժʱ��.Width + TextWidth("��") / 2
    dtpB(0).Top = lbl��Ժʱ��.Top - (dtpB(0).Height - lbl��Ժʱ��.Height) \ 2
    Label2(0).Left = dtpB(0).Left + dtpB(0).Width + TextWidth("��") / 2
    Label2(0).Top = lbl��Ժʱ��.Top
    dtpE(0).Left = Label2(0).Left + Label2(0).Width + TextWidth("��") / 2
    dtpE(0).Top = dtpB(0).Top
    CmdRef.Left = dtpE(0).Left + dtpE(0).Width + TextWidth("��") / 2
    CmdRef.Height = TextHeight("��") + 100
    CmdRef.Top = dtpE(0).Top
    
    lblת��.Left = Label2(0).Left
    fraChange.Left = lblת��.Left + TextWidth("��ʾ���")
    fraChange.Top = lblת��.Height + lblת��.Top
    fraChange.Width = TextWidth("ת��")
    txtChange.Left = fraChange.Left + (fraChange.Width - txtChange.Width) / 2
    txtChange.Height = TextHeight("��")
    txtChange.Top = fraChange.Top - txtChange.Height
    
    chkPati(1).Left = chkPati(0).Left + chkPati(0).Width + TextWidth("��") / 2
    picPati.Width = lblת��.Left + lblת��.Width + TextWidth("��") / 2
    rptPati.Width = picPati.Width - 15
    
    '����
    picFilter.Width = TextWidth("���������´��ڳ���37.5�ȵĲ���") + TextWidth("������")
    If picFilter.Width < 2145 Then picFilter.Width = 2145
    lstFilter.Height = lstFilter.ListCount * TextHeight("��") + 30
    picFilter.Height = lstFilter.Height + cmdFilterOK.Height + 120
End Sub

Private Sub GetLocalSetting()
'���ܣ���ע����ȡ��Ժ���˵�ʱ�䷶Χ
    Dim i As Integer
    Dim curDate As Date, intDay As Integer

    '������ʾ��Χ
    mintChange = Val(zlDatabase.GetPara("���ת������", glngSys, pסԺ��ʿվ, 7))
    '�������30���ȡȱʡֵ
    If mintChange > 30 Then mintChange = 7
    
    '��Ժ����ʱ�䷶Χ
    curDate = zlDatabase.Currentdate
    mdtOutEnd = Format(curDate, "yyyy-MM-dd")
    mdtOutBegin = Format(CDate(mdtOutEnd) - 3, "yyyy-MM-dd")
    
    txtChange.Text = mintChange
    dtpE(0).Value = mdtOutEnd
    dtpE(0).Tag = mdtOutEnd
    dtpB(0).Value = mdtOutBegin
    dtpB(0).Tag = mdtOutBegin
    
    For i = 0 To 3
        If i = ��Ժ Then
            chkScope(i).Value = 1
        Else
            chkScope(i).Value = 0
        End If
    Next i
    
    dtpB(0).Enabled = False
    dtpE(0).Enabled = False
    txtChange.Enabled = False
End Sub

Private Function RefreshHistoryData(ByVal lngRow As Long) As Boolean
'��ȡ������ʷ������Ϣ
    Dim lng�ļ�ID As Long, int����ȼ� As Integer, intӤ�� As Integer
    Dim rsTemp As New ADODB.Recordset
    Dim i As Integer
    Dim lngItemNO As Long, lngHistoryRow As Long, lngHistoryCol As Long
    Dim strContent As String
    Dim strValues As String, strFileds As String
    Dim strKey As String, strKeys As String
    Dim intCOl As Integer
    Dim strPart As String, lng��ԴID As Long, int���� As Integer, int��ʾ As Integer, intModify As Integer, strNote As String
    Dim bln���� As Boolean, blnAllow As Boolean, bln���ʵ��� As Boolean
    Dim lngRecordID As Long
    Dim arrRecordID
    Dim strStart As String, strEnd As String, strTime As String
    
    On Error GoTo Errhand:
    
    If Not mblnInit Then Exit Function
    If lngRow < VsfData.FixedRows - 1 Or Val(VsfData.TextMatrix(lngRow, c�ļ�ID)) <= 0 Then Exit Function
    lng�ļ�ID = Val(VsfData.TextMatrix(lngRow, c�ļ�ID))
    int����ȼ� = Val(VsfData.TextMatrix(lngRow, c����ȼ�))
    intӤ�� = Val(VsfData.TextMatrix(lngRow, cӤ��))
    
    bln���ʵ��� = True
    bln���� = False
    mrsItems.Filter = 0
    mrsItems.Filter = "��Ŀ���=" & gint����
    If mrsItems.RecordCount > 0 Then bln���� = True
    
    If mrsHistory Is Nothing Then Exit Function
    strFileds = "ID|�к�|��Ŀ���|����|��λ|������Դ|��ԴID|����|��ʾ|�޸�|״̬"
    
    '�����ʷ���ݼ�
    mrsHistory.Filter = 0
    If mrsHistory.RecordCount <> 0 Then mrsHistory.MoveFirst
    Do While True
        If mrsHistory.EOF Then Exit Do
        mrsHistory.Delete
        mrsHistory.Update
        mrsHistory.MoveNext
    Loop
    vsfHistory.Rows = vsfHistory.FixedRows
    lngHistoryRow = vsfHistory.Rows - 1
    vsfHistory.RowHeight(-1) = 300 + mintBigSize * 300 / 3
    vsfHistory.FontSize = 9 + mintBigSize * 9 / 3
    
    lngRecordID = 0
    arrRecordID = Array()
    
    'ʱ�䷶Χ
    strStart = Format(DateAdd("d", -1 * Val(txt��ʾ����.Text), IIf(IsDate(mstrDate), mstrDate, zlDatabase.Currentdate)), "yyyy-MM-dd") & " 00:00:00"
    strEnd = Format(DateAdd("d", mintPreDays, IIf(IsDate(mstrDate), mstrDate, zlDatabase.Currentdate)), "yyyy-MM-dd") & " 23:59:59"
    
    strTime = VsfData.TextMatrix(lngRow, c����)
    If strTime <> "" Then
        If InStr(1, strTime, ";") <> 0 Then strTime = Split(strTime, ";")(0)
        If IsDate(Format(strTime, "YYYY-MM-DD HH:mm:ss")) Then
            If CDate(strStart) < CDate(strTime) Then strStart = Format(strTime, "YYYY-MM-DD HH:mm:ss")
        End If
        txt��ʾ����.Text = DateDiff("d", CDate(strStart), IIf(IsDate(mstrDate), mstrDate, zlDatabase.Currentdate))
    End If

    lblDayInfo.Caption = "ʱ�䷶Χ��" & strStart & " �� " & strEnd
    
    '��ȡѡ���˵���ʷ����
    mstrSQL = "SELECT B.ID,B.����ʱ��,C.��Ŀ���,C.��¼����,C.������Դ,C.���²�λ,C.δ��˵��,C.��ԴID,C.����,C.��ʾ,DECODE(C.��Ŀ���,-1,1,C.��¼���) ��¼���" & vbNewLine & _
        " FROM ���˻����ļ� A,���˻������� B,���˻�����ϸ C,�����¼��Ŀ D,���¼�¼��Ŀ E" & vbNewLine & _
        " WHERE A.ID=B.�ļ�ID AND B.ID=C.��¼ID AND A.ID=[1]" & vbNewLine & _
        " AND Mod(C.��¼����,5)<>5 And C.��ֹ�汾 IS NULL  AND B.����ʱ�� Between [2] And [3] And C.��Ŀ���=D.��Ŀ��� And D.��Ŀ���=E.��Ŀ��� AND nvl(D.����ȼ�,3) >=[4] And Nvl(D.���ò���,0) In (0,[5])" & vbNewLine & _
        " Order By B.����ʱ��,DECODE(C.��Ŀ���,-1,1,0),DECODE(C.��Ŀ���,-1,1,C.��¼���),��Ŀ���"
    Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "ͬ������", lng�ļ�ID, CDate(strStart), CDate(strEnd), int����ȼ�, IIf(intӤ�� = 0, 1, 2))
    
    If rsTemp.RecordCount = 0 Then GoTo NextPos
    rsTemp.MoveFirst
    '�����ڶ��ټ�¼
    With rsTemp
        Do While Not .EOF
            If lngRecordID <> Nvl(rsTemp!Id, 0) Then
                ReDim Preserve arrRecordID(UBound(arrRecordID) + 1)
                arrRecordID(UBound(arrRecordID)) = rsTemp!Id
                lngHistoryRow = lngHistoryRow + 1
                If lngHistoryRow > vsfHistory.Rows - 1 Then vsfHistory.Rows = vsfHistory.Rows + 1
                vsfHistory.TextMatrix(lngHistoryRow, c����) = Format(rsTemp!����ʱ��, "yyyy-MM-dd")
                vsfHistory.TextMatrix(lngHistoryRow, cʱ��) = Format(rsTemp!����ʱ��, "HH:mm:ss")
            End If
            lngRecordID = Val(Nvl(rsTemp!Id, 0))
        .MoveNext
        Loop
    End With
    
    'ѭ����ֵ
    For lngRecordID = 0 To UBound(arrRecordID)
        lngHistoryRow = lngRecordID + vsfHistory.FixedRows
        rsTemp.Filter = "ID=" & Val(arrRecordID(lngRecordID))
        For lngHistoryCol = RootCol To vsfHistory.Cols - 1
            lngItemNO = Val(vsfHistory.TextMatrix(0, lngHistoryCol))
            
            If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
    
            strContent = ""
            strPart = ""
            strNote = ""
            lng��ԴID = 0
            int���� = 0
            int��ʾ = 0
            intModify = 0
            With rsTemp
                Do While Not .EOF
                    If lngItemNO <> 4 Then
                        blnAllow = False
                        bln���ʵ��� = False
                        intModify = 0
                        
                        If InStr(1, "," & gint���� & "," & gint���� & "," & gint���� & ",", "," & Val(Nvl(!��Ŀ���)) & ",") > 0 Then
                            Select Case Val(Nvl(!��Ŀ���))
                                Case gint����
                                    If gint���� = lngItemNO Then blnAllow = True
                                Case gint����
                                    If gint���� = lngItemNO Then blnAllow = True
                                Case gint����
                                    If bln���� = True And mint����Ӧ�� = 2 Then
                                        If gint���� = lngItemNO Then blnAllow = True
                                    Else
                                        If gint���� = lngItemNO Then blnAllow = True: bln���ʵ��� = True
                                    End If
                            End Select
                            
                            If blnAllow = True Then
                                If Val(Nvl(!��¼���)) = 0 Then
                                    strContent = Nvl(!��¼����)
                                    strPart = Nvl(!���²�λ)
                                    lng��ԴID = Val(Nvl(!��ԴID))
                                    int���� = Val(Nvl(!����))
                                    int��ʾ = Val(Nvl(!��ʾ))
                                    strNote = Nvl(!δ��˵��)
                                Else '��װ�����º���������
                                    If bln���ʵ��� = False Then
                                        If strContent <> "" Then
                                            If InStr(1, strContent, "/") = 0 Then
                                                '���������ʾ��ʽ:����/����
                                                If mbln����������ʾ And lngItemNO = 2 Then
                                                    strContent = Nvl(!��¼����) & "/" & strContent
                                                Else
                                                    strContent = strContent & "/" & Nvl(!��¼����)
                                                End If
                                            Else
                                                '���������ʾ��ʽ:����/����
                                                If mbln����������ʾ And lngItemNO = 2 Then
                                                    strContent = Nvl(!��¼����) & "/" & Split(strContent, "/")(0)
                                                Else
                                                    strContent = Split(strContent, "/")(0) & "/" & Nvl(!��¼����)
                                                End If
                                                
                                            End If
                                        Else
                                            strContent = Nvl(!��¼����)
                                        End If
                                        
                                        Exit Do
                                    Else
                                        strPart = Nvl(!���²�λ)
                                        lng��ԴID = Val(Nvl(!��ԴID))
                                        int���� = Val(Nvl(!����))
                                        int��ʾ = Val(Nvl(!��ʾ))
                                        strContent = Nvl(!��¼����)
                                        strNote = Nvl(!δ��˵��)
                                        Exit Do
                                    End If
                                End If
                            End If
                        Else
                            If Val(Nvl(!��Ŀ���)) = lngItemNO Then
                                strPart = Nvl(!���²�λ)
                                lng��ԴID = Val(Nvl(!��ԴID))
                                int���� = Val(Nvl(!����))
                                int��ʾ = Val(Nvl(!��ʾ))
                                strContent = Nvl(!��¼����)
                                strNote = Nvl(!δ��˵��)
                                Exit Do
                            End If
                        End If
                    ElseIf InStr(1, ",4,5,", "," & Val(!��Ŀ���) & ",") <> 0 And lngItemNO = 4 Then
                        Select Case Val(!��Ŀ���)
                            Case 4
                                If strContent <> "" Or Nvl(!��¼����) <> "" Then
                                    If InStr(1, strContent, "/") > 0 Then
                                        strContent = Nvl(!��¼����) & "/" & Trim(Split(strContent, "/")(1))
                                    Else
                                        strContent = Nvl(!��¼����) & "/"
                                    End If
                                    strPart = Nvl(!���²�λ)
                                    lng��ԴID = Val(Nvl(!��ԴID))
                                    int���� = Val(Nvl(!����))
                                    int��ʾ = Val(Nvl(!��ʾ))
                                End If
                            Case 5
                                If strContent <> "" Or Nvl(!��¼����) <> "" Then
                                    If InStr(1, strContent, "/") > 0 Then
                                        strContent = Trim(Split(strContent, "/")(0)) & "/" & Nvl(!��¼����)
                                    Else
                                        strContent = "/" & Nvl(!��¼����)
                                    End If
                                End If
                        End Select
                    End If
                    .MoveNext
                Loop
                
                If strContent = "/" Then strContent = ""
                If lngItemNO = 4 Then
                    If InStr(1, strContent, "/") <> 0 Then
                        '�����:53505,�޸��ˣ����Σ�Ѫѹ��ʾ����
                        If Split(strContent, "/")(0) = "�ܲ�" Or Split(strContent, "/")(0) = "���" Or Split(strContent, "/")(0) = "���" Or Split(strContent, "/")(0) = "δ��" Then
                            strContent = Split(strContent, "/")(0)
                        Else
                            If Not IsNumeric(Split(strContent, "/")(0)) And Not IsNumeric(Split(strContent, "/")(1)) Then
                                strContent = ""
                            End If
                        End If
                    End If
                End If
                
                '���������������Ŀ�����Ҳ�λ��Ϊ��
                mrsItems.Filter = "��Ŀ���=" & lngItemNO
                If mrsItems.RecordCount > 0 Then
                    If Nvl(mrsItems!��¼��, 2) = 1 Or lngItemNO = gint���� Then
                        If strNote <> "" And strContent = "" Then
                            strContent = strNote
                            strPart = ""
                        Else
                            If strContent <> "" Then strContent = IIf(strPart = "", "", strPart & ":") & strContent
                        End If
                    Else
                        strPart = ""
                    End If
                Else
                    strPart = ""
                End If
                
                If strContent <> "" Then
                    '��ͬ��������װ�ص���¼����
                    strKey = lngHistoryRow & "," & lngHistoryCol
                    strValues = strKey & "|" & lngHistoryRow & "|" & lngItemNO & "|" & strContent & "|" & strPart & "|1|" & lng��ԴID & "|" & int���� & "|" & int��ʾ & "|" & intModify & "|0"
                    Call Record_Update(mrsHistory, strFileds, strValues, "ID|" & strKey)
                    vsfHistory.TextMatrix(lngHistoryRow, lngHistoryCol) = strContent
                    vsfHistory.RowData(lngHistoryRow) = Val(arrRecordID(lngRecordID))
                    If lng��ԴID <> 0 Then 'ͬ������������
                        vsfHistory.Cell(flexcpForeColor, lngHistoryRow, lngHistoryCol, lngHistoryRow, lngHistoryCol) = 255
                    Else
                        vsfHistory.Cell(flexcpForeColor, lngHistoryRow, lngHistoryCol, lngHistoryRow, lngHistoryCol) = &H80000008
                    End If
'                    If lngItemNo = gint���� Or (lngItemNo = gint���� And mint����Ӧ�� = 2) Then
'                        vsfHistory.Cell(flexcpForeColor, lngHistoryRow, lngHistoryCol, lngHistoryRow, lngHistoryCol) = RGB(0, 0, 255)
'                    Else
'                        vsfHistory.Cell(flexcpForeColor, lngHistoryRow, lngHistoryCol, lngHistoryRow, lngHistoryCol) = 255 '&H8080FF
'                    End If
                End If
            End With
        Next lngHistoryCol
    Next lngRecordID
    
    vsfHistory.Cell(flexcpAlignment, vsfHistory.FixedRows, cʱ��, vsfHistory.Rows - 1, vsfHistory.Cols - 1) = flexAlignCenterCenter
    
    '�����ô˲��˵�������
    For lngHistoryCol = RootCol To vsfHistory.Cols - 1
        vsfHistory.ColHidden(lngHistoryCol) = False
        lngItemNO = Val(vsfHistory.TextMatrix(0, lngHistoryCol))
        mrsItems.Filter = 0
        mrsItems.Filter = "��Ŀ���=" & lngItemNO & " And ����ȼ�>=" & int����ȼ�
        If mrsItems.RecordCount = 0 Then
            vsfHistory.ColHidden(lngHistoryCol) = True
        Else
            '����Ƿ������ڴ˲���
            If Val(vsfHistory.TextMatrix(2, lngHistoryCol)) = 1 Then
                vsfHistory.ColHidden(lngHistoryCol) = IIf(intӤ�� = 0, False, True)
            ElseIf vsfHistory.TextMatrix(2, lngHistoryCol) = 2 Then
                vsfHistory.ColHidden(lngHistoryCol) = IIf(intӤ�� <> 0, False, True)
            End If
        End If
    Next lngHistoryCol
    vsfHistory.RowHeight(-1) = 300 + mintBigSize * 300 / 3
    vsfHistory.FontSize = 9 + mintBigSize * 9 / 3
    
    mrsItems.Filter = 0
NextPos:
    RefreshHistoryData = True
    
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub RefreshPatiList()
    'ˢ�²����嵥
    Call LoadPatient
    If mrsPati.RecordCount > 0 Then mrsPati.MoveFirst
    rptPati.Records.DeleteAll
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
    Dim RS As ADODB.Recordset
    Dim objExtendedBar As CommandBar

    On Error GoTo Errhand

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
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Send, "ͬ��"):   cbrControl.ToolTipText = "����ͬ��"
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Clear, "���"):   cbrControl.ToolTipText = "�����ɾ��ĳ������"
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "���ӿ���"
        Set mcbrMenuBar��λ = .Add(xtpControlButtonPopup, conMenu_Edit_Compend, "��λ"): mcbrMenuBar��λ.ToolTipText = "���²�λ"
        
        Set cbrPop = .Add(xtpControlButtonPopup, conMenu_Edit_Append, "���⴦��"): cbrPop.BeginGroup = True
        Set cbrControl = cbrPop.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Append * 10 + 0, "����", -1, False): cbrControl.IconId = 1: cbrControl.Parameter = ""
        Set cbrControl = cbrPop.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Append * 10 + 1, "�೦[E]", -1, False):  cbrControl.IconId = 1: cbrControl.Parameter = "E"
        Set cbrControl = cbrPop.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Append * 10 + 2, "�೦����[/E]", -1, False):  cbrControl.IconId = 1: cbrControl.Parameter = "/E"
        Set cbrControl = cbrPop.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Append * 10 + 3, "���ʧ��[��]", -1, False): cbrControl.IconId = 1: cbrControl.Parameter = "��"
        Set cbrControl = cbrPop.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Append * 10 + 4, "�˹�����[��]", -1, False): cbrControl.IconId = 1: cbrControl.Parameter = "��"
        Set cbrControl = cbrPop.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Append * 10 + 5, "����[C]", -1, False):   cbrControl.IconId = 1: cbrControl.Parameter = "C"
        Set cbrControl = cbrPop.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Append * 10 + 6, "��������[/C]", -1, False):   cbrControl.IconId = 1: cbrControl.Parameter = "/C"

        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Blankoff, "��ձ��"): cbrControl.ToolTipText = "������������": cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Transf_Cancle, "ȡ��")
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����")
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With

    For Each cbrControl In cbrToolBar.Controls
        If cbrControl.Type <> xtpControlCustom And cbrControl.Type <> xtpControlLabel Then
            cbrControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    Set mcbrToolBar = cbrToolBar
    '------------------------------------------------------------------------------------------------------------------
    '����������
    Set cbrToolBar = cbsThis.Add("��������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    cbrToolBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    With cbrToolBar.Controls
        Set cbrCustom = .Add(xtpControlCustom, conMenu_View_LocationItem, "")
        cbrCustom.flags = xtpFlagAlignLeft
        cbrCustom.Handle = pic��������.hWnd
        cbrCustom.ToolTipText = "����"
    End With

    '�����
    With cbsThis.KeyBindings
        
        .Add FALT, Asc("0"), conMenu_Edit_Append * 10
        .Add FALT, Asc("1"), (conMenu_Edit_Append * 10 + 1)
        .Add FALT, Asc("2"), (conMenu_Edit_Append * 10 + 2)
        .Add FALT, Asc("3"), (conMenu_Edit_Append * 10 + 3)
        .Add FALT, Asc("4"), (conMenu_Edit_Append * 10 + 4)
        .Add FALT, Asc("5"), (conMenu_Edit_Append * 10 + 5)
        .Add FALT, Asc("6"), (conMenu_Edit_Append * 10 + 6)
        
        .Add FCONTROL, Asc("S"), conMenu_Edit_Save
        .Add FCONTROL, Asc("R"), conMenu_Edit_Transf_Cancle
        .Add 0, VK_F1, conMenu_Help_Help
    End With

    InitMenuBar = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub AddActiveMenu(ByVal lngItemNO As Long)
    '------------------------------------------------------------
    '������Ŀ��Ӳ˵�(��Ҫ�����������������Ŀ��λ��Ϣ)
    Dim varTmp As Variant
    Dim strPart As String
    Dim RS As New ADODB.Recordset
    Dim cbrControl As CommandBarControl
    Dim i As Integer
    
    On Error GoTo Errhand
    
    If Not mcbrMenuBar��λ Is Nothing Then
        If mcbrMenuBar��λ.CommandBar.Controls.Count <> 0 Then
            Call mcbrMenuBar��λ.CommandBar.Controls.DeleteAll
        End If
    End If
    
    If mrsPart Is Nothing Then Exit Sub
    If lngItemNO = 0 Then Exit Sub
    
    If lngItemNO = gint���� Then '����
        mstrSQL = "Select ��¼�� From ���¼�¼��Ŀ Where ��Ŀ���=[1]"
        Set RS = zlDatabase.OpenSQLRecord(mstrSQL, "���µ�����¼��", gint����)
        If RS.BOF = False Then
            varTmp = Split(Nvl(RS("��¼��").Value, "��,��,��"), ",")
        Else
            varTmp = Split("��,��,��", ",")
        End If
        
        Set cbrControl = mcbrMenuBar��λ.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Compend * 10, "����" & varTmp(0) & " (&1)", -1, False): cbrControl.Parameter = "����": cbrControl.IconId = 1
        Set cbrControl = mcbrMenuBar��λ.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Compend * 10, "Ҹ��" & varTmp(1) & " (&2)", -1, False): cbrControl.Parameter = "Ҹ��": cbrControl.IconId = 1
        Set cbrControl = mcbrMenuBar��λ.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Compend * 10, "����" & varTmp(2) & " (&3)", -1, False): cbrControl.Parameter = "����": cbrControl.IconId = 1
    ElseIf lngItemNO = gint���� Then '����
        mstrSQL = "Select ��¼�� From ���¼�¼��Ŀ Where ��Ŀ���=[1]"
        Set RS = zlDatabase.OpenSQLRecord(mstrSQL, "���µ�����¼��", gint����)
        If RS.BOF = False Then
            varTmp = Nvl(RS("��¼��").Value, "��")
        Else
            varTmp = "��"
        End If
        
        Set cbrControl = mcbrMenuBar��λ.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Compend * 10, "��������" & varTmp & " (&1)", -1, False): cbrControl.Parameter = "��������": cbrControl.IconId = 1
        Set cbrControl = mcbrMenuBar��λ.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Compend * 10, "������ (&2)", -1, False): cbrControl.Parameter = "������": cbrControl.IconId = 1
    ElseIf lngItemNO = gint���� Then '����
        mstrSQL = "Select ��¼�� From ���¼�¼��Ŀ Where ��Ŀ���=[1]"
        Set RS = zlDatabase.OpenSQLRecord(mstrSQL, "���µ�����¼��", gint����)
        
        If RS.BOF = False Then
            varTmp = Nvl(RS("��¼��").Value, "+")
        Else
            varTmp = "+"
        End If
        
        Set cbrControl = mcbrMenuBar��λ.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Compend * 10, "      " & varTmp & " (&1)", -1, False): cbrControl.Parameter = "": cbrControl.IconId = 1
        Set cbrControl = mcbrMenuBar��λ.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Compend * 10, "����" & " (&2)", -1, False): cbrControl.Parameter = "����": cbrControl.IconId = 1
    Else '����������Ŀ��λ��Ϣ
        varTmp = ""
        mstrSQL = "Select ��¼�� From ���¼�¼��Ŀ Where ��Ŀ���=[1]"
        Set RS = zlDatabase.OpenSQLRecord(mstrSQL, "���µ�����¼��", lngItemNO)
        If RS.BOF = False Then
            varTmp = Nvl(RS("��¼��").Value)
        End If
        mrsPart.Filter = 0
        mrsPart.Filter = "��Ŀ���=" & lngItemNO
        If mrsPart.RecordCount > 1 Then
            i = 1
            varTmp = varTmp & String(mrsPart.RecordCount - 1 - UBound(Split(varTmp, ",")), ",")
            Do While Not mrsPart.EOF
                strPart = Nvl(mrsPart!��λ)
                If strPart = "" Then strPart = "   "
                varTmp = Split(varTmp, ",")
                Set cbrControl = mcbrMenuBar��λ.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Compend * 10, strPart & varTmp(i - 1) & " (&1)", -1, False): cbrControl.Parameter = strPart: cbrControl.IconId = 1
                i = i + 1
            mrsPart.MoveNext
            Loop
        End If
    End If
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub InitCons()
    '��������ؼ�
    picFilter.Visible = False
    picPati.Visible = False
    picInput.Visible = False
    picDouble.Visible = False
    lstNote.Visible = False
    lstSelect(0).Visible = False
    lstSelect(1).Visible = False
    cbo���±�ʶ.Visible = False
    PicLst.Visible = False
    txtLst.Visible = False
    mintType = 0
End Sub

Private Sub InitVariable(Optional ByVal blnClearDate As Boolean = True)
    '�������
    mstrModifyTime = ""
    mblnChage = False
    mblnSaveData = False
    If blnClearDate = True Then
        mint����Ӧ�� = 0
        mbln��Ժ = False
    End If
    mstrTabHead = ""
    mstrItemNo = ""
    mint������Դ = 0
    mintModify = 0
    mintType = 0
    mblnShow = False
    mblnNullRow = False
    mblnClearRow = False
    mblnRefreshData = False
    mblnChkClick = False
    mblnDateFouces = False
End Sub

Private Function InitFilter() As Boolean
'���ܣ���ʼ�����µ�����¼���������
    Dim strFilter As String, strFilterID As String
    Dim arrFilter() As String, arrFilterID() As String
    Dim arrSel() As String
    Dim strSel As String
    Dim i As Integer
    Dim blnSelAll As Boolean
    
    strSel = zlDatabase.GetPara("���µ���������", glngSys, 1255)
    
    '51286,������,2012-07-11,��ӹ���һ�������ϻ���ǼǵĲ���
    If strSel = "" Then
        strSel = "1;1;1;1;1;1;1"
    Else
        arrSel = Split(strSel, ";")
        strSel = strSel & String(IIf(6 - UBound(arrSel) < 0, 0, 6 - UBound(arrSel)), ";")
    End If
    
    arrSel = Split(strSel, ";")
    txtFilter.Tag = ""
    txtFilter.Text = ""
    strFilter = "ȫ��;��Ժ�����ڵĲ���;���������ڵĲ���;���������´��ڳ���37.5�ȵĲ���;Σ/�ز���;ת�������ڵĲ���;һ�������ϻ���ȼ��Ĳ���" & IIf(mstr�������� = "����" Or mstr�������� = "����", ";����������ڵĲ���", "")
    strFilterID = "0;1;2;3;4;5;6" & IIf(mstr�������� = "����" Or mstr�������� = "����", ";7", "")
    arrFilter = Split(strFilter, ";")
    arrFilterID = Split(strFilterID, ";")
    
    blnSelAll = True
    lstFilter.Clear
    For i = 0 To UBound(arrFilter)
        lstFilter.AddItem CStr(arrFilter(i))
        lstFilter.ItemData(lstFilter.NewIndex) = Val(arrFilterID(i))
        
        If i <> 0 Then
            If Val(arrSel(i - 1)) = 1 Then
                txtFilter.Text = txtFilter.Text & ";" & arrFilter(i)
                txtFilter.Tag = txtFilter.Tag & ";" & arrFilterID(i)
            Else
                blnSelAll = False
            End If
        End If
    Next i
    
    If blnSelAll = True Then
        txtFilter.Text = "ȫ��"
        txtFilter.Tag = 0
    Else
        txtFilter.Text = Mid(txtFilter.Text, 2)
        txtFilter.Tag = Mid(txtFilter.Tag, 2)
    End If
    
    '����������С
    picFilter.Width = LenB(StrConv(lstFilter.List(lstFilter.ListCount \ 2), vbFromUnicode)) * 160 + 500
    If picFilter.Width < TextWidth("���������´��ڳ���37.5�ȵĲ���") + TextWidth("������") Then
        picFilter.Width = TextWidth("���������´��ڳ���37.5�ȵĲ���") + TextWidth("������")
    End If
    If picFilter.Width < 2145 Then picFilter.Width = 2145
    lstFilter.Height = lstFilter.ListCount * TextHeight("��") + 30
    picFilter.Height = lstFilter.Height + cmdFilterOK.Height + 120
    
    InitFilter = True
    Exit Function
End Function

Private Sub InitEnv()
    Dim curDate As Date
    Dim intDay As Integer
    Dim RS As New ADODB.Recordset
    Dim blnVisible As Boolean
    On Error GoTo Errhand
    
    mlngHours = Val(Mid(Val(zlDatabase.GetPara("���ݲ�¼ʱ��", glngSys)), 1, 6))
    txt��ʾ����.Tag = 1
    
    dtpDate.Value = Format(date, "YYYY-MM-DD")
    
    If mrsPart Is Nothing Then Set mrsPart = New ADODB.Recordset
    If mrsPart.State = 1 Then mrsPart.Close
    
    '��ȡ���в�λ��Ϣ
    mstrSQL = "SELECT ��Ŀ���,��λ,ȱʡ��,�̶��� FROM ���²�λ"
    Call zlDatabase.OpenRecordset(mrsPart, mstrSQL, "��ȡ��λ��ȡ")
    
    '���ִ��ڵ����л����¼��Ŀ
    mstrSQL = " Select B.������,B.��Ŀ���,B.��Ŀ����,B.��Ŀ����,B.��Ŀ����,B.��Ŀ����,B.��ĿС��,B.��Ŀ��ʾ,B.��Ŀ��λ,B.��Ŀֵ��,B.����ȼ�,B.���ò���,B.Ӧ�÷�ʽ,nvl(A.��¼��,2) ��¼��" & _
              " From �����¼��Ŀ B,���¼�¼��Ŀ A" & _
              " Where B.��Ŀ���=A.��Ŀ���(+) And B.Ӧ�÷�ʽ<>0 " & _
              " Order by B.��Ŀ���"
    Set mrsItems = zlDatabase.OpenSQLRecord(mstrSQL, "���ִ��ڵ����л����¼��Ŀ")
    
    '��ȡδ��˵����Ϣ
    mstrNote = ""
    mstrSQL = "Select ����,���� From ��������˵��"
    Call zlDatabase.OpenRecordset(RS, mstrSQL, "δ��˵����Ϣ")
    lstNote.Clear
    With RS
        Do While Not .EOF
            lstNote.AddItem Nvl(!����)
            lstNote.ItemData(lstNote.NewIndex) = Val(!����)
            mstrNote = mstrNote & "," & Nvl(!����)
        .MoveNext
        Loop
    End With
    If lstNote.ListCount > 0 Then lstNote.ListIndex = 0
    
    If Left(mstrNote, 1) = "," Then mstrNote = Mid(mstrNote, 2)
    
    '��ȡ���µ��嵥
    gstrSQL = " Select ID FROM �����ļ��б� WHERE ����=3 AND ����=-1 AND NVL(ͨ��,0)>0 "
    Set RS = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����µ���Ļ����ļ��嵥")
    
    If RS.RecordCount > 0 Then
        mlng��ʽID = Val(Nvl(RS!Id))
    Else
        mlng��ʽID = 0
        MsgBox "�ڲ����ļ��б���û���ҵ����µ���ص��ļ�,����!", vbInformation, gstrSysName
    End If
    
    blnVisible = False
    '��ȡ��ǰ�����µ����п���
    mstrSQL = " Select distinct B.ID,B.����||'-'||B.���� AS ����,decode(nvl(E.��������,''),'����',1,0) ����" & _
              " From �������Ҷ�Ӧ A,���ű� B,������Ա C,��Ա�� D,��������˵�� E" & _
              " Where A.����ID = b.ID And A.����ID=C.����ID And C.��ԱID=D.ID And A.����ID = [1]" & _
              IIf(InStr(1, mstrPrivs, "��ǰ����") <> 0, "", " And D.ID=[2]") & _
              " And B.ID=E.����ID(+) And E.��������(+)='����'" & _
              " Order by B.����||'-'||B.����"
    Set RS = zlDatabase.OpenSQLRecord(mstrSQL, "��ȡ��ǰ�����µ����п���", mlng����ID, glngUserId)
    With cboUnit
        .Clear
        .Tag = ""
        If InStr(1, mstrPrivs, "��ǰ����") <> 0 Then
            .AddItem "���п���"
            .ItemData(.NewIndex) = -1
        End If
        Do While Not RS.EOF
            .AddItem zlCommFun.Nvl(RS!����)
            .ItemData(.NewIndex) = RS!Id
            .Tag = .Tag & "[LPF]" & RS!����
            If blnVisible = False Then blnVisible = (Val(RS!����) = 1)
            RS.MoveNext
        Loop
        .Tag = IIf(blnVisible = True, 1, 0) & .Tag
        If Left(.Tag, 5) = "[LPF]" Then .Tag = Mid(.Tag, 6)
        If .ListCount <> 0 Then .ListIndex = 0
    End With
    
    '���ع���������Ϣ
    Call InitFilter
    
    '���ز���ѡ��
    With cboPati
        .AddItem "����": .ItemData(.NewIndex) = 0
        .AddItem "���˱���": .ItemData(.NewIndex) = 1
        .AddItem "Ӥ��": .ItemData(.NewIndex) = 2
        .ListIndex = 0
    End With
    
    cbo���±�ʶ.Clear
    cbo���±�ʶ.AddItem "2��/��"
    cbo���±�ʶ.AddItem "4��/��"
    cbo���±�ʶ.AddItem "6��/��"
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Sub LoadPatient()
    Dim strSQL As String
    On Error GoTo Errhand
    '58890:������,2013-02-26,��Ժ���˶�ȡ�����Ż�(������Ժ���˱���в�ѯ)
    '��Ժ����ƺ�ת�ƴ���Ʋ���(���˿��������Ĳ������ɽ���)
    'c.����id + 0,˵����ͨ��H����������ӹ��˺󣬼�¼�������٣�������B�������
    If chkScope(�����).Value = 1 Then
        strSQL = _
            "Select /*+ RULE */Distinct" & vbNewLine & _
            " Decode(B.״̬,1,0,Decode(c.��ʼԭ��,3,1,2)) As ����, Decode(Nvl(b.����״̬, 0), 0, 999, b.����״̬) As ����2," & _
            " Decode(B.״̬,1,'��Ժ����ס����',Decode(c.��ʼԭ��,3,'ת�ƴ���ס����','ת��������ס����')) As ����," & _
            " a.����id, b.��ҳid, A.�����,B.סԺ��, a.����, a.�Ա�, b.����," & vbNewLine & _
            " d.���� As ����, c.����id, c.����ҽʦ As סԺҽʦ,b.���λ�ʿ, b.����״̬, lpad(nvl(C.����,' '),10,' ') as ����," & _
            " e.���� As ����ȼ�, b.�ѱ�,b.��ǰ����, b.��Ժ����, b.��Ժ����,B.��Ժ��ʽ, b.��������, b.״̬, b.����, a.���￨��," & vbNewLine & _
            " -1 As ·��״̬,trunc(sysdate)-trunc(b.��Ժ����)+1 as סԺ����,Z.��ɫ" & vbNewLine & _
            "From ������Ϣ A, ������ҳ B, ���˱䶯��¼ C, ���ű� D, �շ���ĿĿ¼ E,�������� Z,��Ժ���� R" & vbNewLine & _
            "Where B.��������=Z.����(+) And A.����ID=R.����ID And a.����id = b.����id And Nvl(b.��ҳid, 0) <> 0 And b.����id = c.����id And b.��ҳid = c.��ҳid And c.����id = d.Id" & vbNewLine & _
            "      And (d.վ��='" & gstrNodeNo & "' Or d.վ�� is Null)" & vbNewLine & _
            "      And b.����ȼ�id = e.Id(+) And Nvl(c.���Ӵ�λ, 0) = 0 And c.��ֹʱ�� Is Null" & vbNewLine & _
            "      And ((c.��ʼԭ�� in(1,3) And Exists(Select 1 From �������Ҷ�Ӧ H Where c.����id = h.����id And h.����id = [1])) or (c.��ʼԭ��=15 And c.����id = [1]))" & vbNewLine & _
            "      And ((c.��ʼԭ�� = 1 And b.״̬ = 1) Or (c.��ʼԭ�� in (3,15) And c.��ʼʱ�� Is Null And b.״̬ = 2)) "
    End If
    '��Ժ����
    If chkScope(��Ժ).Value = 1 Then
        strSQL = strSQL & IIf(strSQL <> "", " Union All ", "") & _
            "Select /*+ RULE */ Decode(B.״̬,3,4,DECODE(B.��Ժ����, NULL, 3.1,DECODE(B.״̬,2,3.2,3))) as ����," & _
            " Decode(Nvl(B.����״̬,0),0,999,B.����״̬) as ����2," & _
            " Decode(B.״̬,3,'Ԥ��Ժ����',DECODE(B.��Ժ����, NULL, '��ͥ����',DECODE(B.״̬,2,'Ԥת�Ʋ���', '��Ժ����'))) as ����," & _
            " A.����ID,B.��ҳID,A.�����,B.סԺ��,A.����,A.�Ա�,B.����,C.���� as ����,B.��Ժ����ID ����ID,B.סԺҽʦ,B.���λ�ʿ,B.����״̬," & _
            " lpad(nvl(B.��Ժ����,' '),10,' ') as ����,E.���� as ����ȼ�,B.�ѱ�,B.��ǰ����,B.��Ժ����,B.��Ժ����,B.��Ժ��ʽ,B.��������," & _
            " B.״̬,B.����,A.���￨��,Nvl(B.·��״̬,-1) ·��״̬,trunc(sysdate)-trunc(b.��Ժ����)+1 as סԺ����,z.��ɫ" & _
            " From ������Ϣ A,������ҳ B,���ű� C,�շ���ĿĿ¼ E,�������� Z,��Ժ���� R" & _
            " Where B.��������=Z.����(+) And A.����ID=B.����ID And A.סԺ����=B.��ҳID And Nvl(B.��ҳID,0)<>0 And Nvl(B.״̬,0)<>1" & _
            " And B.��Ժ����ID=C.ID And B.����ȼ�ID=E.ID(+) And (c.վ��='" & gstrNodeNo & "' Or c.վ�� is Null)" & _
            " And Nvl(B.����״̬,0)<>5 And B.���ʱ�� is NULL And A.����ID=R.����ID And R.����ID=[1]"
    End If
    '��Ժ����:��Ժ���˿������ж��סԺ
    If chkScope(��Ժ).Value = 1 Then
        strSQL = strSQL & IIf(strSQL <> "", " Union All ", "") & _
            "Select /*+ RULE */ Decode(B.��Ժ��ʽ,'����',7,6) as ����," & _
            " Decode(Nvl(B.����״̬,0),0,999,B.����״̬) as ����2," & _
            " Decode(B.��Ժ��ʽ,'����','��������','��Ժ����') as ����," & _
            " A.����ID,B.��ҳID,A.�����,B.סԺ��,A.����,A.�Ա�,B.����,C.���� as ����,B.��Ժ����ID ����ID,B.סԺҽʦ,B.���λ�ʿ,B.����״̬," & _
            " lpad(nvl(B.��Ժ����,' '),10,' ') AS ����,E.���� as ����ȼ�,B.�ѱ�,B.��ǰ����,B.��Ժ����,B.��Ժ����,B.��Ժ��ʽ,B.��������," & _
            " B.״̬,B.����,A.���￨��,Nvl(B.·��״̬,-1) ·��״̬,trunc(b.��Ժ����)-trunc(b.��Ժ����)+1 as סԺ����,z.��ɫ" & _
            " From ������Ϣ A,������ҳ B,���ű� C,�շ���ĿĿ¼ E,�������� Z" & _
            " Where B.��������=Z.����(+) And A.����ID=B.����ID And Nvl(B.��ҳID,0)<>0 And B.״̬=0" & _
            " And B.��Ժ����ID=C.ID And B.����ȼ�ID=E.ID(+) And B.��ǰ����ID+0=[1] And (c.վ��='" & gstrNodeNo & "' Or c.վ�� is Null)" & _
            " And B.��Ժ���� Between [2] And [3] And Nvl(B.����״̬,0)<>5 And B.���ʱ�� is NULL"
    End If
    'ת������:��Ժ,ҽ���ʹ�����ʾ����ת��ǰ��
    If chkScope(ת��).Value = 1 Then
        strSQL = strSQL & IIf(strSQL <> "", " Union All ", "") & _
            "Select /*+ RULE */ Distinct 5 as ����,Decode(Nvl(B.����״̬,0),0,999,B.����״̬) as ����2,'ת������' as ����," & _
            " A.����ID,B.��ҳID,A.�����,B.סԺ��,A.����,A.�Ա�,B.����,D.���� as ����,C.����ID,C.����ҽʦ as סԺҽʦ,B.���λ�ʿ,B.����״̬," & _
            " lpad(nvl(C.����,' '),10,' ') as ����,E.���� as ����ȼ�,B.�ѱ�,B.��ǰ����,B.��Ժ����,B.��Ժ����,B.��Ժ��ʽ,B.��������," & _
            " B.״̬,B.����,A.���￨��,Nvl(B.·��״̬,-1) ·��״̬,trunc(sysdate)-trunc(b.��Ժ����)+1 as סԺ����,z.��ɫ" & _
            " From ������Ϣ A,������ҳ B,���˱䶯��¼ C,���ű� D,�շ���ĿĿ¼ E,�������� Z" & _
            " Where B.��������=Z.����(+) And A.����ID=B.����ID And Nvl(B.��ҳID,0)<>0 And B.����ȼ�ID=E.ID(+)" & _
            " And B.����ID=C.����ID And B.��ҳID=C.��ҳID" & _
            " And B.��ǰ����ID<>[1] And C.����ID+0=[1] And C.����ID=D.ID" & _
            " And Nvl(C.���Ӵ�λ,0)=0 And C.��ֹԭ�� In(3,15) And C.��ֹʱ�� Between Sysdate-[4] And Sysdate" & _
            " And Nvl(B.״̬,0)<>2 And Nvl(B.����״̬,0)<>5 And B.���ʱ�� is NULL"
    End If
    '�ٴι��˳������µ��ļ��Ĳ���
    
    strSQL = "SELECT A.����,A.����2,A.����,A.����ID,A.��ҳID,A.�����,A.סԺ��,A.����,A.�Ա�,A.����,A.����,A.����ID,A.סԺҽʦ,A.���λ�ʿ,A.����״̬," & _
            " lpad(nvl(A.����,' '),10,' ') as ����,A.����ȼ�,A.�ѱ�,A.��ǰ����,A.��Ժ����,A.��Ժ����,A.��Ժ��ʽ,A.��������," & _
            " A.״̬,A.����,A.���￨��,A.·��״̬,A.סԺ����,A.��ɫ" & _
            " From (" & strSQL & ") A,���˻����ļ� B" & _
            " Where A.����ID=B.����ID and A.��ҳID=B.��ҳID And nvl(B.Ӥ��,0)=0 And B.�鵵�� is null and B.����ʱ�� is null and B.��ʽID=[5]"
    strSQL = strSQL & " Order by A.����,A.����,A.��ҳID"
    
    Screen.MousePointer = 11
    On Error GoTo Errhand
    Set mrsPati = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�����б�", mlng����ID, _
        CDate(Format(mdtOutBegin, "yyyy-MM-dd 00:00:00")), CDate(Format(mdtOutEnd, "yyyy-MM-dd 23:59:59")), _
        mintChange, mlng��ʽID)
    Screen.MousePointer = 0
    Exit Sub
Errhand:
    Screen.MousePointer = 0
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cboUnit_Click()
    Dim ArrCode() As String
    Dim blnVisble As Boolean
    
    On Error GoTo Errhand
    
    If cboUnit.ListCount = 0 Then GoTo ErrNext
    If cboUnit.Tag = "" Then cboUnit.Tag = "0"

    ArrCode = Split(cboUnit.Tag, "[LPF]")
    mlng����id = Val(cboUnit.ItemData(cboUnit.ListIndex))
    '��ȡ��������
    Call Get��������(mlng����id)
    
    '����ˢ�¹�������
    Call InitFilter
    
    'ֻ�п���Ϊ�����ƲŽ���Ӥ������
    blnVisble = (Val(ArrCode(cboUnit.ListIndex)) = 1)
    lblPati.Visible = blnVisble
    cboPati.Visible = blnVisble
    cboPati.Enabled = blnVisble
    cmdFilter.Left = IIf(blnVisble = True, cboPati.Left + cboPati.Width + 75, lblPati.Left)
    cmdAddUser.Left = cmdFilter.Left + cmdFilter.Width + 195
ErrNext:
    Call dtpDate_GotFocus
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cboUnit_KeyPress(KeyAscii As Integer)
    Call zlControl.CboMatchIndex(cboUnit.hWnd, KeyAscii)
End Sub

Private Sub cbo���±�ʶ_Click()
    Call Save���±�ʶ(VsfData.Row)
End Sub

Private Sub Save���±�ʶ(ByVal lngCurRow As Long, Optional ByVal str���±�ʶ As String = "")
    Dim lngRow As Long
    Dim lng����ID As Long, lng��ҳID As Long, lngӤ�� As Long
    On Error GoTo Errhand
    '���没�����±�ʶ
    lng����ID = Val(VsfData.TextMatrix(lngCurRow, c����ID))
    lng��ҳID = Val(VsfData.TextMatrix(lngCurRow, c��ҳID))
    lngӤ�� = Val(VsfData.TextMatrix(lngCurRow, cӤ��))
    
    gstrSQL = "ZL_�������±�ʶ_Update(" & lng����ID & "," & lng��ҳID & "," & lngӤ�� & ",'" & IIf(str���±�ʶ = "", cbo���±�ʶ.Text, str���±�ʶ) & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "���没�����±�ʶ")
    
    For lngRow = VsfData.FixedRows To VsfData.Rows - 1
        If lng����ID = Val(VsfData.TextMatrix(lngRow, c����ID)) And lng��ҳID = Val(VsfData.TextMatrix(lngRow, c��ҳID)) And lngӤ�� = Val(VsfData.TextMatrix(lngRow, cӤ��)) Then
            VsfData.TextMatrix(lngRow, c���±�ʶ) = IIf(str���±�ʶ = "", cbo���±�ʶ.Text, str���±�ʶ)
        End If
    Next lngRow
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cbo���±�ʶ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call MoveNextCell
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngStartRow As Long, lngRow As Long, lngCol As Long, lngItemNO As Long, lngRow1 As Long
    Dim strKey As String, strFileds As String, strValues As String
    Dim strPart As String, strValue As String, strPart1 As String, strPart2 As String
    Dim strTime As String, strPatientTime As String, strInfo As String
    Dim arrValue() As Variant, arrCOL() As Variant, arrPart() As Variant, i As Long, intSate As Integer
    Dim arrID() As Variant
    Dim cbrCheck As CommandBarControl
    
    '���±�ʶ
    Dim str���±�ʶ As String
    
    Select Case Control.Id
        Case conMenu_Edit_Send '����ͬ��
            If VsfData.Row < VsfData.FixedRows Then Exit Sub
            arrID = Array()
            '��ѡ���е����� ͬ��������������
            lngStartRow = VsfData.Row
            strTime = VsfData.TextMatrix(lngStartRow, cʱ��)
            str���±�ʶ = Trim(VsfData.TextMatrix(lngStartRow, c���±�ʶ))
            strFileds = "ID|�к�|��Ŀ���|����|��λ|������Դ|״̬"
            '��ȡ���е�������Ϣ
            If mrsCell Is Nothing Then Exit Sub
            mrsCell.Filter = 0
            mrsCopy.Filter = 0
            '���ڱ���������mrscell��¼������Ϊ��,�˴����и���,�����ɾ����ֵ����Mrscell����
            mrsCopy.Filter = "�к�=" & lngStartRow
            Do While Not mrsCopy.EOF
                mrsCell.Filter = "ID='" & Nvl(mrsCopy!Id) & "' And ״̬=1"
                If mrsCell.RecordCount = 0 Then
                    strValues = Nvl(mrsCopy!Id) & "|" & Val(Nvl(mrsCopy!�к�)) & "|" & Val(Nvl(mrsCopy!��Ŀ���)) & "|" & Nvl(mrsCopy!����) & "|" & _
                        Nvl(mrsCopy!��λ) & "|" & Val(Nvl(mrsCopy!������Դ)) & "|" & 0
                    Call Record_Update(mrsCell, strFileds, strValues, "ID|" & Nvl(mrsCopy!Id))
                    ReDim Preserve arrID(UBound(arrID) + 1)
                    arrID(UBound(arrID)) = Nvl(mrsCopy!Id)
                End If
                mrsCopy.MoveNext
            Loop
            
            mrsCell.Filter = 0
            mrsCell.Filter = "�к�=" & lngStartRow & " And ״̬=1"
            If mrsCell.RecordCount = 0 Then Exit Sub
            arrValue = Array()
            arrCOL = Array()
            arrPart = Array()
            Do While Not mrsCell.EOF
                lngCol = Val(Split(mrsCell!Id, ",")(1))
                lngItemNO = Val(Nvl(mrsCell!��Ŀ���))
                strPart = Trim(Nvl(mrsCell!��λ))
                strPart1 = Trim(GetPart(lngItemNO))
                strPart2 = ""
                If strPart <> strPart1 Then
                    strPart2 = strPart
                End If
                strValue = Val(Nvl(mrsCell!��Ŀ���)) & "|" & Nvl(mrsCell!����) & "|" & Nvl(mrsCell!��λ) & "|0|1"
                ReDim Preserve arrValue(UBound(arrValue) + 1)
                arrValue(UBound(arrValue)) = strValue
                ReDim Preserve arrCOL(UBound(arrCOL) + 1)
                arrCOL(UBound(arrCOL)) = lngCol
                ReDim Preserve arrPart(UBound(arrPart) + 1)
                arrPart(UBound(arrPart)) = strPart2
            mrsCell.MoveNext
            Loop
            
            mrsCell.Filter = 0
            '��ʼ�������� �����ݵ��в����и�ֵ
            For lngRow = VsfData.FixedRows To VsfData.Rows - 1
                If lngRow <> lngStartRow And VsfData.RowHidden(lngRow) = False Then
                    '����û��Ѿ�����ʱ�� �Ͳ��ڽ���ʱ���ͬ��
                    If Trim(VsfData.TextMatrix(lngRow, cʱ��)) = "" And strTime <> "" Then
                        '�û�û��¼��ʱ�� ����Ҫ���ͬ����ʱ���Ƿ�Ϸ�(���Ϸ������и��� �û���Ҫ�ֹ�¼��)
                        strPatientTime = VsfData.TextMatrix(lngRow, c����)
                        If CheckDateTime(strTime, strPatientTime, strInfo) = True Then
                            VsfData.TextMatrix(lngRow, cʱ��) = strTime
                        End If
                    End If
                    
                    For i = 0 To UBound(arrValue)
                        strKey = lngRow & "," & Val(arrCOL(i))
                        mrsCell.Filter = "ID='" & strKey & "' And ״̬=1"
                        If mrsCell.RecordCount = 0 Then
                            strValues = strKey & "|" & lngRow & "|" & CStr(arrValue(i))
                            strValue = Split(CStr(arrValue(i)), "|")(1)
                            If Trim(CStr(arrPart(i))) <> "" Then
                                strValue = CStr(arrPart(i)) & ":" & strValue
                            End If
                            Call Record_Update(mrsCell, strFileds, strValues, strKey)
                            VsfData.TextMatrix(lngRow, Val(arrCOL(i))) = strValue
                            mblnChage = True
                        End If
                    Next i
                    'ͬ�����±�ʶ
                    If str���±�ʶ <> "" And Trim(VsfData.TextMatrix(lngRow, c���±�ʶ)) = "" Then
                        Call Save���±�ʶ(lngRow, str���±�ʶ)
                    End If
                End If
            Next lngRow
            
            'ͬ������˵�����ɾ���ղŸ��Ƶ���Ϣ
            mrsCell.Filter = 0
            For i = 0 To UBound(arrID)
                mrsCell.Filter = "ID='" & CStr(arrID(i)) & "'"
                mrsCell.Delete
                mrsCell.Update
            Next i
            
            VsfData.Cell(flexcpAlignment, VsfData.FixedRows, cʱ��, VsfData.Rows - 1, VsfData.Cols - 1) = flexAlignCenterCenter
            Call InitCons
        Case conMenu_Edit_Clear '���
           Call Edit_Clear
        Case conMenu_Edit_NewItem '��ӿ���
            If VsfData.Row < VsfData.FixedRows Then Exit Sub
            lngStartRow = VsfData.Row
            lngRow1 = VsfData.Row + 1
            VsfData.Rows = VsfData.Rows + 1
            VsfData.TextMatrix(VsfData.Rows - 1, c�ļ�ID) = VsfData.TextMatrix(lngStartRow, c�ļ�ID)
            VsfData.TextMatrix(VsfData.Rows - 1, c����) = VsfData.TextMatrix(lngStartRow, c����)
            VsfData.TextMatrix(VsfData.Rows - 1, c����) = VsfData.TextMatrix(lngStartRow, c����)
            VsfData.TextMatrix(VsfData.Rows - 1, c����) = VsfData.TextMatrix(lngStartRow, c����)
            VsfData.TextMatrix(VsfData.Rows - 1, c����ID) = VsfData.TextMatrix(lngStartRow, c����ID)
            VsfData.TextMatrix(VsfData.Rows - 1, c��ҳID) = VsfData.TextMatrix(lngStartRow, c��ҳID)
            VsfData.TextMatrix(VsfData.Rows - 1, cӤ��) = VsfData.TextMatrix(lngStartRow, cӤ��)
            VsfData.TextMatrix(VsfData.Rows - 1, c����ȼ�) = VsfData.TextMatrix(lngStartRow, c����ȼ�)
            VsfData.TextMatrix(VsfData.Rows - 1, c���±�ʶ) = VsfData.TextMatrix(lngStartRow, c���±�ʶ)
            VsfData.TextMatrix(VsfData.Rows - 1, c����) = VsfData.TextMatrix(lngStartRow, c����)
            VsfData.TextMatrix(VsfData.Rows - 1, c��Ժ) = VsfData.TextMatrix(lngStartRow, c��Ժ)
            lngStartRow = lngStartRow + 1
            
            For lngRow = VsfData.Rows - 2 To lngStartRow Step -1
                mrsCell.Filter = "�к�=" & lngRow
                If mrsCell.RecordCount > 0 Then
                    mrsCell.MoveFirst
                    Do While Not mrsCell.EOF
                        strFileds = "ID|�к�"
                        strKey = Nvl(mrsCell!Id)
                        lngCol = Val(Split(Nvl(mrsCell!Id, ","), ",")(1))
                        strValues = lngRow + 1 & "," & lngCol & "|" & lngRow + 1
                        Call Record_Update(mrsCell, strFileds, strValues, "ID|" & strKey)
                    mrsCell.MoveNext
                    Loop
                End If
                '����mrsCopy���ݼ�
                If Not mrsCopy Is Nothing Then
                    mrsCopy.Filter = "�к�=" & lngRow
                    If mrsCopy.RecordCount > 0 Then
                        mrsCopy.MoveFirst
                        Do While Not mrsCopy.EOF
                            strFileds = "ID|�к�"
                            strKey = Nvl(mrsCopy!Id)
                            lngCol = Val(Split(Nvl(mrsCopy!Id, ","), ",")(1))
                            strValues = lngRow + 1 & "," & lngCol & "|" & lngRow + 1
                            Call Record_Update(mrsCopy, strFileds, strValues, "ID|" & strKey)
                        mrsCopy.MoveNext
                        Loop
                    End If
                End If
                
                If Not mrsData Is Nothing Then
                    '���»ָ����ݵ��к�
                    mrsData.Filter = "�к�=" & lngRow
                    If mrsData.RecordCount > 0 Then
                        mrsData.MoveFirst
                        Do While Not mrsData.EOF
                            strFileds = "�к�"
                            strValues = lngRow + 1
                            Call Record_Update(mrsData, strFileds, strValues, "�к�|" & lngRow)
                        mrsData.MoveNext
                        Loop
                    End If
                End If
                
                VsfData.RowPosition(lngRow) = lngRow + 1
            Next lngRow
            VsfData.Cell(flexcpAlignment, VsfData.FixedRows, cʱ��, VsfData.Rows - 1, VsfData.Cols - 1) = flexAlignCenterCenter
            mblnChage = True
            VsfData.RowHeight(-1) = 300 + mintBigSize * 300 / 3
            VsfData.Select lngRow1, cʱ��
            VsfData.SetFocus
            '���ñ༭��ɫ
            Call SetTabEditColor
    
        Case conMenu_Edit_Save '����
            If Not SaveDate Then Exit Sub
        Case conMenu_Edit_Transf_Cancle 'ȡ��
            If Not EditCancle Then Exit Sub
        Case conMenu_Edit_Blankoff '��ձ��������(���������ݴ���)
            Call InitCons  '���ر༭�ؼ�
            Call InitVariable(False)
            '����ִ��¼��
            mrsCell.Filter = 0
            Do While Not mrsCell.EOF
               mrsCell.Delete
               mrsCell.Update
               mrsCell.MoveNext
            Loop
            mrsCopy.Filter = 0
            Do While Not mrsCopy.EOF
               mrsCopy.Delete
               mrsCopy.Update
               mrsCopy.MoveNext
            Loop
            mrsData.Filter = 0
            Do While Not mrsData.EOF
               mrsData.Delete
               mrsData.Update
               mrsData.MoveNext
            Loop
            mrsHistory.Filter = 0
            Do While Not mrsHistory.EOF
                mrsHistory.Delete
                mrsHistory.Update
                mrsHistory.MoveNext
            Loop
            
            Call ColligationTab(False)
            Call ColligationHistoryTab
            VsfData.Select VsfData.FixedRows, cʱ��
            Call AdjustRowFlag(VsfData, VsfData.FixedRows)
        Case conMenu_Edit_Compend * 10  '��λ
            If VsfData.Row < VsfData.FixedRows Then Exit Sub
            strPart = Trim(Control.Parameter)
            lngRow = VsfData.Row
            lngCol = VsfData.Col
            lngItemNO = Val(VsfData.TextMatrix(0, lngCol))
            strValue = Trim(VsfData.TextMatrix(lngRow, lngCol))
            If InStr(1, strValue, ":") <> 0 Then
                strValue = Split(strValue, ":")(1)
            End If
            '�������vsfdataû����ֵ,�����û��Ƿ��Ѿ�¼������
            If picInput.Visible = True And mintType = 1 Then
                strValue = txtInput.Text
                strPart2 = GetPart(lngItemNO)
                txtInput.Tag = Trim(strPart)
                '���²�λ�˵���ѡ����
                Call VsfData_AfterRowColChange(lngRow, cʱ��, lngRow, lngCol)
                For Each cbrCheck In mcbrToolBar.Controls(4).CommandBar.Controls
                    If cbrCheck.Parameter = Control.Parameter Then
                        cbrCheck.Checked = True
                    Else
                        cbrCheck.Checked = False
                    End If
                Next
                Exit Sub
            End If
            
            Call InitCons
            mintType = 0
            strFileds = "ID|�к�|��Ŀ���|����|��λ|������Դ|״̬"
            If strValue <> "" And IsNumeric(strValue) Then
                strKey = lngRow & "," & lngCol
                mrsCell.Filter = "ID='" & strKey & "'"
                If mrsCell.RecordCount = 0 Then '��mrscopy��ֵ���ƹ���
                    mrsCopy.Filter = "ID='" & strKey & "'"
                    strValues = Nvl(mrsCopy!Id) & "|" & Val(Nvl(mrsCopy!�к�)) & "|" & Val(Nvl(mrsCopy!��Ŀ���)) & "|" & Nvl(mrsCopy!����) & "|" & _
                        Nvl(mrsCopy!��λ) & "|" & Val(Nvl(mrsCopy!������Դ)) & "|" & 0
                    Call Record_Update(mrsCell, strFileds, strValues, "ID|" & Nvl(mrsCopy!Id))
                End If
                strFileds = "��λ|״̬"
                strValues = strPart
                If mrsCell.RecordCount > 0 Then
                    strPart1 = Trim(Nvl(mrsCell!��λ))
                    strPart2 = GetPart(lngItemNO)
                    If strPart1 = "" Then strPart1 = strPart2
                    If strPart1 <> strPart Then
                        Call Record_Update(mrsCell, strFileds, strValues & "|1", "ID|" & strKey)
                        If strPart <> strPart2 And strPart <> "" Then
                            VsfData.TextMatrix(lngRow, lngCol) = strPart & ":" & strValue
                        Else
                            VsfData.TextMatrix(lngRow, lngCol) = strValue
                        End If
                        If picInput.Visible = True Then txtInput.Tag = Trim(strPart)
                        mblnChage = True
                        '���²�λ�˵���ѡ����
                        Call VsfData_AfterRowColChange(lngRow, cʱ��, lngRow, lngCol)
                    End If
                End If
            End If
        Case conMenu_Edit_Append * 10, conMenu_Edit_Append * 10 + 1, conMenu_Edit_Append * 10 + 2, conMenu_Edit_Append * 10 + 3, conMenu_Edit_Append * 10 + 4, conMenu_Edit_Append * 10 + 5, conMenu_Edit_Append * 10 + 6
            If VsfData.Row < VsfData.FixedRows Then Exit Sub
            lngRow = VsfData.Row
            lngCol = VsfData.Col
            lngItemNO = Val(VsfData.TextMatrix(0, lngCol))
            strValue = Trim(VsfData.TextMatrix(lngRow, lngCol))
            
            If mintType = 1 And picInput.Visible = True Then strValue = txtInput.Text
            strPart = ""
            If InStr(1, "," & gint��� & "," & gint��Һ & ",", "," & lngItemNO & ",") = 0 Then Exit Sub
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
                Case conMenu_Edit_Append * 10 + 3
                    strPart = "��"
                    strValue = ""
                Case conMenu_Edit_Append * 10 + 4
                    strPart = "��"
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
                    If lngItemNO = gint��� Then
                        For i = 0 To 4
                            Select Case i
                                Case 0
                                    strPart1 = "E"
                                Case 1
                                    strPart1 = "/"
                                Case 2
                                    strPart1 = "*"
                                Case 3
                                    strPart1 = "��"
                                Case 4
                                    strPart1 = "��"
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
            
            If mintType = 1 And picInput.Visible = True Then
                If Len(txtInput.Text) > txtInput.MaxLength Then
                    RaiseEvent AfterRowColChange("ѡ������ݳ�����Ŀ����,���ڻ����¼�������������Ŀ����.", True)
                    Exit Sub
                End If
                txtInput.Text = strValue
                For Each cbrCheck In mcbrToolBar.Controls(5).CommandBar.Controls
                    If cbrCheck.Id = Control.Id Then
                        cbrCheck.Checked = True
                    Else
                        cbrCheck.Checked = False
                    End If
                Next

                Exit Sub
            End If
            
            Call InitCons
            mintType = 0
            '�������ݱ䶯��¼��
            strFileds = "ID|�к�|��Ŀ���|����|��λ|������Դ|״̬"
            strKey = lngRow & "," & lngCol
            mrsCell.Filter = "ID='" & strKey & "'"
            If mrsCell.RecordCount = 0 Then '��mrscopy��ֵ���ƹ���
                mrsCopy.Filter = "ID='" & strKey & "'"
                If mrsCopy.RecordCount > 0 Then
                    strValues = Nvl(mrsCopy!Id) & "|" & Val(Nvl(mrsCopy!�к�)) & "|" & Val(Nvl(mrsCopy!��Ŀ���)) & "|" & Nvl(mrsCopy!����) & "|" & _
                        Nvl(mrsCopy!��λ) & "|" & Val(Nvl(mrsCopy!������Դ)) & "|" & 0
                    Call Record_Update(mrsCell, strFileds, strValues, "ID|" & Nvl(mrsCopy!Id))
                    intSate = 1
                Else
                    '����µļ�¼
                    If Trim(strValue) = "" Then Exit Sub
                    strValues = strKey & "|" & lngRow & "|" & lngItemNO & "|" & strValue & "|" & _
                        "" & "|" & 0 & "|" & 1
                    Call Record_Update(mrsCell, strFileds, strValues, "ID|" & strKey)
                    GoTo ErrGO
                End If
            Else
                intSate = IIf(Trim(strValue) = "", 3, 1)
            End If
            mrsCell.Filter = "ID='" & strKey & "'"
            strFileds = "����|״̬"
            strValues = strValue & "|" & intSate
            If mrsCell.RecordCount > 0 Then
                If Trim(Nvl(mrsCell!����)) <> strValue Then
                    Call Record_Update(mrsCell, strFileds, strValues, "ID|" & strKey)
ErrGO:
                    VsfData.TextMatrix(lngRow, lngCol) = strValue
                    mblnChage = True
                    Call VsfData_AfterRowColChange(lngRow, cʱ��, lngRow, lngCol)
                End If
            End If
            
        Case conMenu_Help_Help '����
            RaiseEvent UsrHelp
        Case conMenu_File_Exit '�˳�
            RaiseEvent UsrExit
    End Select
End Sub

Private Sub cbsThis_Resize()
    Dim lngScaleLeft As Long, lngScaleTop As Long, lngScaleRight As Long, lngScaleBottom As Long
    Call cbsThis.GetClientRect(lngScaleLeft, lngScaleTop, lngScaleRight, lngScaleBottom)
    
    Err = 0: On Error Resume Next
    
    If Not mblnInit Then picSplit.Top = ScaleHeight - 3000
    
    picMain.Move lngScaleLeft, lngScaleTop, lngScaleRight, picSplit.Top - lngScaleTop
    VsfData.Move lngScaleLeft + 100, 100, lngScaleRight - lngScaleLeft - 100 * 2
    VsfData.Height = picMain.Height - VsfData.Top
    
    picHistory.Move lngScaleLeft, picSplit.Top + picSplit.Height, lngScaleRight, lngScaleBottom - picSplit.Top
    picDay.Left = VsfData.Left
    picDay.Top = 60
    picDay.Width = VsfData.Width
    vsfHistory.Left = VsfData.Left
    vsfHistory.Top = picDay.Top + picDay.Height + 60
    vsfHistory.Height = picHistory.Height - vsfHistory.Top - 50
    vsfHistory.Width = VsfData.Width
    picSplit.Left = lngScaleLeft
    picSplit.Width = picMain.Width
    picNull.Move lngScaleLeft, lngScaleTop, lngScaleRight, lngScaleBottom
    With lblInfo
        .Top = (picNull.Height - .Height) / 2
        .Left = (picNull.Width - .Width) / 2
    End With
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.Id
        Case conMenu_Edit_Save, conMenu_Edit_Transf_Cancle
            Control.Enabled = mblnChage
        Case conMenu_Edit_NewItem '��ӿ���
            Control.Enabled = (mblnNullRow And mblnInit)
        Case conMenu_Edit_Compend '��λ
            Control.Enabled = (mcbrMenuBar��λ.CommandBar.Controls.Count <> 0)
        Case conMenu_Edit_Clear, conMenu_Edit_Send '��� ����ͬ��
            Control.Enabled = (mblnClearRow And mblnInit)
        Case conMenu_Edit_Blankoff '��ձ��������(���������ݴ���)
            Control.Enabled = mblnNullRow
            picNull.Visible = Not mblnNullRow
            If mblnNullRow <> (VsfData.ScrollBars = flexScrollBarBoth) Then
                VsfData.ScrollBars = IIf(mblnNullRow, flexScrollBarBoth, flexScrollBarNone)
            End If
        Case conMenu_Edit_Append * 10 + 0, conMenu_Edit_Append
            Control.Enabled = is������Һ(1) Or is������Һ(2)
        Case conMenu_Edit_Append * 10 + 1, conMenu_Edit_Append * 10 + 2, conMenu_Edit_Append * 10 + 3, conMenu_Edit_Append * 10 + 4
            Control.Enabled = is������Һ(1)
        Case conMenu_Edit_Append * 10 + 5, conMenu_Edit_Append * 10 + 6
            Control.Enabled = is������Һ(2)
        Case conMenu_View_LocationItem
            'dtpDate.Enabled = Not mblnInit
    End Select
End Sub

Private Sub Edit_Clear()
'---------------------------------------
'����:���������Ϣ
'---------------------------------------
    Dim lngStartRow As Long, lngRow As Long, lngCol As Long, lngItemNO As Long
    Dim lngRow1 As Long
    Dim strKey As String, strFileds As String, strValues As String
    
    '����Ѿ�¼���������Ϣ
    If VsfData.Row < VsfData.FixedRows Then Exit Sub
    strFileds = "ID|�к�|��Ŀ���|����|��λ|������Դ|״̬"
    On Error GoTo Errhand
    
    lngRow = VsfData.Row
    lngRow1 = lngRow
    '����е�������Ϣ
    For lngCol = cʱ�� To VsfData.Cols - 1
        VsfData.TextMatrix(lngRow, lngCol) = ""
    Next lngCol
   
    '�����¼����Ϣ
    mrsCell.Filter = "�к�=" & lngRow
    mrsCell.Sort = "ID"
    Do While Not mrsCell.EOF
        mrsCell.Delete
        mrsCell.Update
        mblnChage = True
    mrsCell.MoveNext
    Loop
    
    mrsCell.Filter = "�к�=" & lngRow
    mrsCopy.Filter = "�к�=" & lngRow
    If mrsCopy.RecordCount > 0 Then
        Do While Not mrsCopy.EOF
            strValues = Nvl(mrsCopy!Id) & "|" & lngRow & "|" & Val(Nvl(mrsCopy!��Ŀ���)) & "|"
            If InStr(1, ",0,9", "," & mrsCopy!������Դ & ",") <> 0 Then
                strValues = strValues & "|" & Nvl(mrsCopy!��λ) & "|" & Nvl(mrsCopy!������Դ) & "|1"
                Call Record_Update(mrsCell, strFileds, strValues, "ID|" & Nvl(mrsCopy!Id))
            Else
                strValues = strValues & Nvl(mrsCopy!����) & "|" & Nvl(mrsCopy!��λ) & "|" & Nvl(mrsCopy!������Դ) & "|0"
                Call Record_Update(mrsCell, strFileds, strValues, "ID|" & Nvl(mrsCopy!Id))
            End If
            mrsCopy.MoveNext
        Loop
'        mrsCell.Filter = 0
'        Call OutputRsData(mrsCell, True)
        
        mrsData.Filter = "�к�=" & lngRow
        If mrsData.RecordCount > 0 Then
            VsfData.TextMatrix(lngRow, cʱ��) = mrsData.Fields(cʱ��).Value
            mrsData!ɾ�� = 1
            mrsData.Update
        End If
        'ɾ�������һ�п���,����ɾ����
        VsfData.RowHidden(lngRow) = True
        lngStartRow = lngRow
        VsfData.Rows = VsfData.Rows + 1
        VsfData.TextMatrix(VsfData.Rows - 1, c�ļ�ID) = VsfData.TextMatrix(lngStartRow, c�ļ�ID)
        VsfData.TextMatrix(VsfData.Rows - 1, c����) = VsfData.TextMatrix(lngStartRow, c����)
        VsfData.TextMatrix(VsfData.Rows - 1, c����) = VsfData.TextMatrix(lngStartRow, c����)
        VsfData.TextMatrix(VsfData.Rows - 1, c����) = VsfData.TextMatrix(lngStartRow, c����)
        VsfData.TextMatrix(VsfData.Rows - 1, c����ID) = VsfData.TextMatrix(lngStartRow, c����ID)
        VsfData.TextMatrix(VsfData.Rows - 1, c��ҳID) = VsfData.TextMatrix(lngStartRow, c��ҳID)
        VsfData.TextMatrix(VsfData.Rows - 1, cӤ��) = VsfData.TextMatrix(lngStartRow, cӤ��)
        VsfData.TextMatrix(VsfData.Rows - 1, c����ȼ�) = VsfData.TextMatrix(lngStartRow, c����ȼ�)
        VsfData.TextMatrix(VsfData.Rows - 1, c����) = VsfData.TextMatrix(lngStartRow, c����)
        VsfData.TextMatrix(VsfData.Rows - 1, c��Ժ) = VsfData.TextMatrix(lngStartRow, c��Ժ)
        
        lngStartRow = lngStartRow + 1
        For lngRow = VsfData.Rows - 2 To lngStartRow Step -1
            '����ԭʼ��¼��
            mrsCell.Filter = "�к�=" & lngRow
            If mrsCell.RecordCount > 0 Then
                mrsCell.MoveFirst
                Do While Not mrsCell.EOF
                    strFileds = "ID|�к�"
                    strKey = Nvl(mrsCell!Id)
                    lngCol = Val(Split(Nvl(mrsCell!Id, ","), ",")(1))
                    strValues = lngRow + 1 & "," & lngCol & "|" & lngRow + 1
                    Call Record_Update(mrsCell, strFileds, strValues, "ID|" & strKey)
                mrsCell.MoveNext
                Loop
            End If
            '����mrsCopy���ݼ�
            mrsCopy.Filter = "�к�=" & lngRow
            If mrsCopy.RecordCount > 0 Then
                mrsCopy.MoveFirst
                Do While Not mrsCopy.EOF
                    strFileds = "ID|�к�"
                    strKey = Nvl(mrsCopy!Id)
                    lngCol = Val(Split(Nvl(mrsCopy!Id, ","), ",")(1))
                    strValues = lngRow + 1 & "," & lngCol & "|" & lngRow + 1
                    Call Record_Update(mrsCopy, strFileds, strValues, "ID|" & strKey)
                mrsCopy.MoveNext
                Loop
            End If

            '���»ָ����ݵ��к�
            mrsData.Filter = "�к�=" & lngRow
            If mrsData.RecordCount > 0 Then
                mrsData.MoveFirst
                Do While Not mrsData.EOF
                    strFileds = "�к�"
                    strValues = lngRow + 1
                    Call Record_Update(mrsData, strFileds, strValues, "�к�|" & lngRow)
                mrsData.MoveNext
                Loop
            End If
            VsfData.RowPosition(lngRow) = lngRow + 1
        Next lngRow
        VsfData.Cell(flexcpAlignment, VsfData.FixedRows, cʱ��, VsfData.Rows - 1, VsfData.Cols - 1) = flexAlignCenterCenter
        lngRow1 = lngRow1 + 1
    End If
    VsfData.RowHeight(-1) = 300 + mintBigSize * 300 / 3
    VsfData.Select lngRow1, cʱ��
    VsfData.SetFocus
    
    '���ñ༭��ɫ
    Call SetTabEditColor
    mblnChage = True
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function EditCancle() As Boolean
'---------------------------------------------------
'����:�û�ȡ������
'---------------------------------------------------
    '�û�ȡ������ʱ���¼��������Ϣ,�����б���Ϣ����(ȡ���ظ���)
    Dim lngRow As Long, lngCol As Long, lngRows As Long, lngCOls As Long
    Dim lng�к� As Long
    Dim rsPati As New ADODB.Recordset
    Dim strFileds As String, strValues As String
    Dim lngID As Long
    
    On Error GoTo Errhand
    
    VsfData.Cell(flexcpText, lngRow, cʱ��, VsfData.Rows - 1, VsfData.Cols - 1) = ""
        
    Set mrsCell = New ADODB.Recordset
    gstrFields = "ID," & adLongVarChar & ",40|�к�," & adDouble & ",18|��Ŀ���," & adDouble & ",18|����," & adLongVarChar & ",40|" & _
        "��λ," & adLongVarChar & ",20|������Դ," & adDouble & ",1|��ԴID," & adDouble & ",18|����," & adDouble & ",1|��ʾ," & adDouble & ",1|" & _
        "�޸�," & adDouble & ",1|״̬," & adDouble & ",1"
    Call Record_Init(mrsCell, gstrFields)
    
    If mblnSaveData = False Then
        Call Record_Init(mrsCopy, gstrFields)
    End If
    
    gstrFields = "ID|�к�|��Ŀ���|����|��λ|������Դ|��ԴID|����|��ʾ|�޸�|״̬"
    '���¼��ر����Ϣ
    Call ColligationTab(False)
    
    '��ʼ�ָ�����
    mrsData.Filter = 0
    If mrsData.RecordCount > 0 Then mrsData.MoveFirst
    lngCOls = VsfData.Cols - 1
    lngRows = mrsData.RecordCount - 1
    
    For lngRow = 0 To lngRows
        If lngRow > VsfData.Rows - VsfData.FixedRows - 1 Then VsfData.Rows = VsfData.Rows + 1
        For lngCol = c�ļ�ID To lngCOls
            If lngCol = c���� Then
                VsfData.TextMatrix(VsfData.FixedRows + lngRow, lngCol) = IIf(Val(Nvl(mrsData.Fields(cӤ��).Value)) > 0, Space(4), "") & Nvl(mrsData.Fields(lngCol).Value)
            Else
                VsfData.TextMatrix(VsfData.FixedRows + lngRow, lngCol) = Nvl(mrsData.Fields(lngCol).Value)
            End If
        Next
        If mrsData!ɾ�� = 1 Then VsfData.RowHidden(VsfData.FixedRows + lngRow) = True
        '���������к�
        lng�к� = Val(Nvl(mrsData!�к�))
        mrsCopy.Filter = "�к�=" & lng�к�
        Do While Not mrsCopy.EOF
            mrsCopy!�к� = lngRow + VsfData.FixedRows
            mrsCopy.Update
        mrsCopy.MoveNext
        Loop
        mrsData!�к� = lngRow + VsfData.FixedRows
        mrsData.MoveNext
    Next
    
    If VsfData.Rows <= VsfData.FixedRows Then VsfData.Rows = VsfData.Rows + 1
    VsfData.RowHeight(-1) = 300 + mintBigSize * 300 / 3
    VsfData.Cell(flexcpAlignment, VsfData.FixedRows, cʱ��, VsfData.Rows - 1, VsfData.Cols - 1) = flexAlignCenterCenter
    '���ñ༭��ɫ
    Call SetTabEditColor
    
    VsfData.Select VsfData.FixedRows, cʱ��
    VsfData.SetFocus
    
    mblnChage = False
    mblnShow = False
    mbln��Ժ = False
    
    Call InitCons
    
    EditCancle = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function CheckEditData() As Boolean
'-------------------------------------------------
'����:����Ƿ��Ѿ��������ݱ༭
'-------------------------------------------------
    Dim lngCOls As Long, lngCol As Long
    On Error GoTo Errhand
    
    '���ڲ����б�ȫ���ֹ���ӵĲ���,�����ʱ�����û��¼���κ���������Խ������ڵ��л�
    If Format(mstrDate, "YYYY-MM-DD") = Format(dtpDate.Value, "YYYY-MM-DD") Then Exit Function
    
    If Not mblnRefreshData Then
        If mblnSaveData = True Then
            'ȫ���ֹ���ӵĲ��ˣ�������л����ھ�ֻ����������Ϣ,������Ϣȫ�����
            lngCOls = VsfData.Cols - 1
            mrsData.Filter = 0
            If mrsData.RecordCount > 0 Then mrsData.MoveFirst
            Do While Not mrsData.EOF
                For lngCol = cʱ�� To lngCOls
                    mrsData.Fields(lngCol) = ""
                Next lngCol
                mrsData("ɾ��") = 0
                mrsData.Update
            Loop
        End If
        mblnSaveData = False
        Call EditCancle
        Exit Function
'        If mrsCell Is Nothing Then Exit Function
'        mrsCell.Filter = 0
'        mrsCell.Filter = "״̬<>3"
'        If mrsCell.RecordCount = 0 Then
'            VsfData.Cell(flexcpText, VsfData.FixedRows, cʱ��, VsfData.Rows - 1, VsfData.Cols - 1) = ""
'            mblnChage = False
'            Call InitCons
'            mblnSaveData = False
'            Exit Function
'        Else
'            'MsgBox "�����Ѿ��������ݵ����ڽ����޸�ʱ,���ȵ��ȡ����ť,�ڽ��������л��˲�����", vbInformation, gstrSysName
'            Call EditCancle
'            mblnSaveData = False
'            Exit Function
'        End If
    Else '��������б��а������˳����Ĳ�������Ҫ�ڸı����ں�����ˢ�²�����Ϣ
'        If MsgBox("���ڱ��л��������Ѿ���������,�������������Ҫ�ֹ����¹���/��Ӳ���,�����Ƿ����?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
'            VsfData.Rows = VsfData.FixedRows + 1
'            VsfData.Cell(flexcpText, VsfData.FixedRows, 0, VsfData.Rows - 1, VsfData.Cols - 1) = ""
'            Call InitCons
'            Call InitVariable
'            If cmdFilter.Enabled = True Then cmdFilter.SetFocus
'            Exit Function
'        End If
        'ֱ�����¹�����Ϣ
        Call cmdFilter_Click
        Exit Function
    End If
    CheckEditData = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function GetPart(ByVal lng��Ŀ��� As Long) As String
'����:��ȡĬ�ϵ����²�λ
    Dim strPart As String
    mrsPart.Filter = "��Ŀ���=" & lng��Ŀ��� & " and ȱʡ��=1"
    If mrsPart.RecordCount > 0 Then strPart = Trim(zlCommFun.Nvl(mrsPart("��λ")))
    GetPart = strPart
End Function

Private Sub chkPati_Click(Index As Integer)
    Dim i As Integer
    Dim blnTrue As Boolean
    
    For i = 0 To chkPati.Count - 1
        If i <> Index Then
            blnTrue = (chkPati(i).Value <> 0)
        End If
    Next i
    
    If Not blnTrue And chkPati(Index).Value = 0 Then chkPati(IIf(Index = 0, 1, 0)).Value = 1
    blnTrue = (chkPati(IIf(Index = 0, 1, 0)).Value <> 0)
    
    If blnTrue And chkPati(Index).Value <> 0 Then
        mintPatiNo = 0
    ElseIf blnTrue Then
        mintPatiNo = IIf(Index = 0, 2, 1)
    Else
        mintPatiNo = IIf(Index = 0, 1, 2)
    End If
    
    For i = 0 To cboPati.ListCount - 1
        If mintPatiNo = cboPati.ItemData(i) Then Call zlControl.CboSetIndex(cboPati.hWnd, i)
    Next i
End Sub

Private Sub chkScope_Click(Index As Integer)
    Dim blnEnable As Boolean
    blnEnable = (chkScope(Index).Value = 1)
    
    Select Case Index
        Case ��Ժ
            dtpB(0).Enabled = blnEnable
            dtpE(0).Enabled = blnEnable
        Case ת��
            txtChange.Enabled = blnEnable
    End Select
End Sub

Private Sub chkScope_Validate(Index As Integer, Cancel As Boolean)
    Dim i As Integer
    Dim blnAll As Boolean
    
    For i = 0 To 3
        If blnAll = False Then
            blnAll = (chkScope(i).Value = 1)
        End If
    Next i
    
    If blnAll = False Then
        chkScope(Index).Value = 1
        RaiseEvent AfterRowColChange("��Ӳ���ʱ����������Ҫѡ��һ��", True)
        Cancel = True
    End If
End Sub

Private Sub chkSwitch_Click()
    '��ʼ���в�������ѡ��
    Dim intValue As Integer
    Dim lngLoop As Long
    Dim objRow As ReportRow
    Dim arrIndex()
    Dim i As Integer
    
    If mblnChkClick = True Then mblnChkClick = False: Exit Sub
    
    intValue = chkSwitch.Value
    
    arrIndex = Array()
    '��¼չ�����������
    For Each objRow In rptPati.Rows
       If objRow.GroupRow Then
           If objRow.Expanded = True Then
               ReDim Preserve arrIndex(UBound(arrIndex) + 1)
               arrIndex(UBound(arrIndex)) = objRow.Index
           End If
       End If
    Next
    
    '��������ѡ��
    For Each objRow In rptPati.Rows
        If objRow.GroupRow And objRow.Childs.Count > 0 Then
            For lngLoop = 0 To objRow.Childs.Count - 1
                If Not (objRow.Childs(lngLoop).Record Is Nothing) Then
                    'If Trim(objRow.Childs(lngLoop).Record.Item(c_��Ժ����).Value) <> "" Then Exit For
                    objRow.Childs(lngLoop).Record.Item(c_ѡ��).Checked = IIf(intValue = 0, False, True)
                End If
            Next lngLoop
        End If
    Next
    
    rptPati.Populate
    
    '��ԭչ�������
    For Each objRow In rptPati.Rows
       If objRow.GroupRow Then
           objRow.Expanded = False
           For i = 0 To UBound(arrIndex)
               If objRow.Index = Val(arrIndex(i)) Then
                   objRow.Expanded = True
                   Exit For
               End If
           Next i
       End If
    Next
End Sub

Private Sub chkSwitch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    zlCommFun.ShowTipInfo chkSwitch.hWnd, "�Բ��˽�������ȫѡ/��ѡ����(��������Ժ����)"
End Sub

Private Sub cmdAddUser_Click()
    Dim lngColor As Long
    Dim lngLoop As Long
    Dim objRow As ReportRow
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim strPatient As String '�����б���Ϣ
    Dim lngRow As Long, lngID As Long 'VSFѡ��Ĳ���ID
    Dim lngLeft As Long, lngTop  As Long, lngRight As Long, lngBottom As Long
    Dim ArrCode() As String
    Dim blnVisible As Boolean
    
    If mrsPati.State = 0 Then
        Call RefreshPatiList 'ˢ�²����б���Ϣ
    End If
    
    mrsPati.Filter = ""
    If rptPati.Records.Count = 0 And mrsPati.RecordCount > 0 Then
        '��ʾ�����б�ѡ��
        With mrsPati
            .MoveFirst
            
            Do While Not .EOF
                Set objRecord = rptPati.Records.Add()
                objRecord.Tag = CStr(!����ID & "," & !��ҳID)
                Set objItem = objRecord.AddItem("")
                objItem.HasCheckbox = True
                objItem.Checked = False
                
                Set objItem = objRecord.AddItem(""): objItem.Icon = IIf(!�Ա� = "��", 1, 0)
                Set objItem = objRecord.AddItem(CStr(!����))
                objItem.Caption = CStr(!���� & !����)
                Set objItem = objRecord.AddItem(CStr(!���� & !����))
                objItem.Caption = CStr(!���� & !����)
                
                Set objItem = objRecord.AddItem(LPAD(Nvl(!����), 10, " "))
                objItem.Caption = Trim(Nvl(!����, " "))
                objRecord.AddItem Val(!����ID)
                objRecord.AddItem Val(!��ҳID)
                objRecord.AddItem CStr(Nvl(!����))
                objRecord.AddItem CStr(Nvl(!����))
                Set objItem = objRecord.AddItem(CStr(Nvl(!סԺ��)))
                objItem.Caption = Nvl(!סԺ��, " ")
                
                '53881:������,2012-09-19,���Ժ����Ӧ�ü���ʱ���룬������¼��ʱ�䷢������
                Set objItem = objRecord.AddItem(Format(!��Ժ����, "yyyy-MM-dd HH:mm:ss"))
                objItem.Caption = Format(!��Ժ����, "yyyy-MM-dd  HH:mm:ss")
                Set objItem = objRecord.AddItem(Format(!��Ժ����, "yyyy-MM-dd HH:mm:ss"))
                objItem.Caption = Format(!��Ժ����, "yyyy-MM-dd  HH:mm:ss")
                
                '��ȡ�������͵���ɫ
                lngColor = Nvl(!��ɫ, 0)
                If lngColor <> 0 Then objRecord.Item(c_����).ForeColor = lngColor
                
                .MoveNext
            Loop
            
            .MoveFirst
        End With
    End If
    
    If cboUnit.Tag = "" Then cboUnit.Tag = "0"
    ArrCode = Split(cboUnit.Tag, "[LPF]")
    blnVisible = (Val(ArrCode(0)) = 1)
    chkPati(0).Visible = blnVisible
    chkPati(1).Visible = blnVisible
    
    Select Case mintPatiNo
        Case 1
            chkPati(0).Value = 1
            chkPati(1).Value = 0
        Case 2
            chkPati(0).Value = 0
            chkPati(1).Value = 1
        Case Else
            chkPati(0).Value = 1
            chkPati(1).Value = 1
    End Select
    chkPati(0).Enabled = (Not mblnNullRow And blnVisible)
    chkPati(1).Enabled = (Not mblnNullRow And blnVisible)
    '��������
    rptPati.Populate 'ȱʡ��ѡ���κ���
    picPati.Left = cmdAddUser.Left + 60
    picPati.Top = picMain.Top
    picPati.Visible = True
    
    With chkSwitch
        .Value = 0
        .Top = rptPati.Top + 100
        .Left = rptPati.Left + (rptPati.Columns(c_ѡ��).Width * Screen.TwipsPerPixelX - .Width) / 2
        .ZOrder 0
    End With
    
    cmdFilterUserCancle.Left = picPati.ScaleWidth - cmdFilterUserCancle.Width - 100
    cmdFilterUserOk.Left = cmdFilterUserCancle.Left - cmdFilterUserCancle.Width - 60
    
    strPatient = ""
    lngRow = 0
    If mblnInit = True Then
        If VsfData.Cols >= RootCol Then
            For lngRow = VsfData.FixedRows To VsfData.Rows - 1
                strPatient = strPatient & "," & VsfData.TextMatrix(lngRow, c����ID)
                If VsfData.Row = lngRow Then
                    lngID = Val(VsfData.TextMatrix(lngRow, c����ID))
                End If
            Next lngRow
        End If
    End If
    
    If Left(strPatient, 1) = "," Then strPatient = Mid(strPatient, 2)
    
    '�������ѡ����
    For lngLoop = 0 To rptPati.Rows.Count - 1
         If Not (rptPati.Rows(lngLoop).Record Is Nothing) Then
            rptPati.Rows(lngLoop).Record.Item(c_ѡ��).Checked = False
         End If
    Next
    
    '����Ѿ�������ˢ�� �͹�ѡ�Ѿ����˳����Ĳ���
'    For lngLoop = 0 To rptPati.Rows.Count - 1
'         If Not (rptPati.Rows(lngLoop).Record Is Nothing) Then
'             If InStr(1, "," & strPatient & ",", "," & Val(rptPati.Rows(lngLoop).Record.Item(c_����ID).Value) & ",") <> 0 Then
'                 rptPati.Rows(lngLoop).Record.Item(c_ѡ��).Checked = True
'             Else
'                rptPati.Rows(lngLoop).Record.Item(c_ѡ��).Checked = False
'             End If
'         End If
'     Next
    
    'ѡ�е�ǰ����(���۵���Ļ�,Rows.Countֻ����ĸ�����,�����ȶ�λ,���۵�)
    For lngLoop = 0 To rptPati.Rows.Count - 1
        If Not (rptPati.Rows(lngLoop).Record Is Nothing) Then
            If lngID <> 0 Then
                If Val(rptPati.Rows(lngLoop).Record.Item(C_����ID).Value) = lngID Then
                    Set rptPati.FocusedRow = rptPati.Rows(lngLoop)
                    Exit For
                End If
            Else
                 Set rptPati.FocusedRow = rptPati.Rows(lngLoop)
                 Exit For
            End If
        End If
    Next
    
    '�۵������� (ѡ�в�����һ�鲻�۵�)
    For Each objRow In rptPati.Rows
        If objRow.GroupRow And objRow.Index <> rptPati.FocusedRow.ParentRow.Index Then
            objRow.Expanded = False
        End If
    Next
    
    chkSwitch.Enabled = (rptPati.Records.Count > 0)
    
    If rptPati.Records.Count > 0 Then rptPati.FocusedRow.EnsureVisible
    rptPati.SetFocus
End Sub

Private Sub cmdFilter_Click()
'�����û����õĹ����������˲�����Ϣ
    mblnInit = False
    mlng����id = Val(cboUnit.ItemData(cboUnit.ListIndex))

    mstrDate = Format(dtpDate.Value, "YYYY-MM-DD")
    Call InitCons  '���ر༭�ؼ�
    Call InitVariable '���������Ϣ
    Call zlRefreshDate 'ˢ������
    mblnInit = True
    
    '�������ݼ�
    Call Data_Save
End Sub

Private Function zlRefreshDate(Optional blnFillPage As Boolean = True) As Boolean
'-----------------------------------------------------
'����:ˢ������
'blnFillPage �Ƿ�������ȡ������Ϣ
'-----------------------------------------------------
    Dim ArrCode() As String
    Dim blnVisible As Boolean
    
    'ֻ�п���Ϊ�����ƲŽ���Ӥ������
    If cboUnit.Tag = "" Then cboUnit.Tag = "0"
    ArrCode = Split(cboUnit.Tag, "[LPF]")
    blnVisible = (Val(ArrCode(cboUnit.ListIndex)) = 1)
    If blnVisible = True Then
        mintPatiNo = cboPati.ItemData(cboPati.ListIndex)
    Else
        mintPatiNo = 1
    End If
    '��ȡ������������
    Call InitCurveDate
    '�󶨱����
    Call ColligationTab(blnFillPage)
    '����ʱ��ʷ������
    Call ColligationHistoryTab
End Function

Private Sub InitCurveDate()
'----------------------------------------
'��ȡ�ճ�Ҫ�༭����������
'----------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strTmp As String
    Dim i As Integer
    Dim strFind As String
    On Error GoTo Errhand
        
        '��ʼ�����ݼ�¼��
        If Not (mrsCell Is Nothing) Then Set mrsCell = Nothing
        If Not (mrsCopy Is Nothing) Then Set mrsCopy = Nothing
        Set mrsCell = New ADODB.Recordset
        Set mrsCopy = New ADODB.Recordset
        
        gstrFields = "ID," & adLongVarChar & ",40|�к�," & adDouble & ",18|��Ŀ���," & adDouble & ",18|����," & adLongVarChar & ",40|" & _
            "��λ," & adLongVarChar & ",20|������Դ," & adDouble & ",1|��ԴID," & adDouble & ",18|����," & adDouble & ",1|��ʾ," & adDouble & ",1|" & _
            "�޸�," & adDouble & ",1|״̬," & adDouble & ",1"
        Call Record_Init(mrsCell, gstrFields)
        Call Record_Init(mrsCopy, gstrFields)
        Call Record_Init(mrsHistory, gstrFields)
        
        gstrFields = "ID|�к�|��Ŀ���|����|��λ|������Դ|��ԴID|����|��ʾ|�޸�|״̬"
        
        mstrTabHead = "|�ļ�ID|����|����|����|����ID|��ҳID|Ӥ��|��¼ID|����ȼ�|���±�ʶ|��Ժ|����|ʱ��"
        mstrItemNo = ""
        
        Select Case mintPatiNo
            Case 1
                strFind = " And instr('0,1',B.���ò���)<>0"
            Case 2
                strFind = " And instr('0,2',B.���ò���)<>0"
            Case Else
                strFind = ""
        End Select
        '��ȡҪ¼��ı��������Ϣ
        mstrSQL = "SELECT /*+ RULE */ A.��Ŀ���,DECODE(A.��Ŀ���,4,'Ѫѹ',A.��¼��) || DECODE(nvl(A.��λ,''),'','', '(' || A.��λ || ')') ��Ŀ����,A.�������,B.������  FROM ���¼�¼��Ŀ A,����������Ŀ C, �����¼��Ŀ B" & vbNewLine & _
                "WHERE  B.��ĿID=C.ID(+) AND A.��Ŀ���=B.��Ŀ��� AND NVL(B.Ӧ�÷�ʽ,0)=1 And A.��Ŀ���<>5 And B.��Ŀ����=1 " & strFind & vbNewLine & _
                "AND (B.���ÿ���=1 OR (B.���ÿ���=2 AND EXISTS (SELECT 1 FROM �������ÿ��� D," & vbNewLine & _
                "Table(Cast(f_num2list([1]) As zlTools.t_Numlist)) E WHERE D.��Ŀ���=B.��Ŀ��� AND D.����ID=E.Column_Value)))" & vbNewLine & _
                "ORDER BY Decode(A.��Ŀ���,1,0,1),A.�������"

        If mlng����id = -1 Then
            For i = 1 To cboUnit.ListCount - 1
                strTmp = strTmp & "," & cboUnit.ItemData(i)
            Next i
        Else
            strTmp = CStr(mlng����id)
        End If
        
        If Left(strTmp, 1) = "," Then strTmp = Mid(strTmp, 2)
        
        Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "���µ�����¼��", strTmp)
        
        '��ȡ����������Ŀ
        rsTemp.Filter = "������='1)����������Ŀ'"
        With rsTemp
            Do While Not .EOF
                mstrTabHead = mstrTabHead & "|" & Nvl(!��Ŀ����)
                mstrItemNo = mstrItemNo & "|" & Val(Nvl(!��Ŀ���))
            .MoveNext
            Loop
        End With
        
        If Left(mstrItemNo, 1) = "|" Then mstrItemNo = Mid(mstrItemNo, 2)
        '��ȡ����ѹ����ѹ
        rsTemp.Filter = "��Ŀ���=4"
        'mrsItems.Filter="��Ŀ���=4"
        If rsTemp.RecordCount > 0 Then '����ѹ������ѹ����ͬʱ����
            mstrTabHead = mstrTabHead & "|" & Nvl(rsTemp!��Ŀ����)    ' "|Ѫѹ(" & Nvl(mrsItems!��Ŀ��λ) & ")"
            mstrItemNo = mstrItemNo & "|4"
        End If
        
        '��ȡʣ�����±����Ŀ
        rsTemp.Filter = "������<>'1)����������Ŀ' and ��Ŀ���<>4"
        rsTemp.Sort = "�������"
        With rsTemp
            Do While Not .EOF
                mstrTabHead = mstrTabHead & "|" & Nvl(!��Ŀ����)
                mstrItemNo = mstrItemNo & "|" & Val(Nvl(!��Ŀ���))
            .MoveNext
            Loop
        End With
        
        'ȷ�������Ƿ����������
        mrsItems.Filter = "��Ŀ���=" & gint����
        If mrsItems.RecordCount > 0 Then mint����Ӧ�� = Val(Nvl(mrsItems!Ӧ�÷�ʽ, 0))
        mrsItems.Filter = 0
        
        Set mrsData = CopyNewRs
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function CopyNewRs() As ADODB.Recordset
'����:��ʼ����Ŀ�м�¼��
    Dim arrCOL() As String
    Dim i As Integer
    Dim strHead As String
    Dim rsNewRs As New ADODB.Recordset
    strHead = Mid(mstrTabHead, 2)
    arrCOL = Split(strHead, "|")
    
    '��¼����ʽ
    '"�к�|�ļ�ID|����|����|����ID|��ҳID|Ӥ��|����|��Ժ|ʱ��" + ����������Ŀ
    With rsNewRs
        .Fields.Append "�к�", adDouble, 18
        For i = 0 To UBound(arrCOL)
            Select Case CStr(arrCOL(i))
                Case "�ļ�ID,����ID,��ҳID,��¼ID"
                    .Fields.Append CStr(arrCOL(i)), adDouble, 18, adFldIsNullable
                Case "Ӥ��,��Ժ,����ȼ�"
                    .Fields.Append CStr(arrCOL(i)), adDouble, 1, adFldIsNullable
                Case "����"
                    .Fields.Append CStr(arrCOL(i)), adLongVarChar, 50, adFldIsNullable
                Case Else
                    .Fields.Append CStr(arrCOL(i)), adLongVarChar, 20, adFldIsNullable
            End Select
        Next i
        .Fields.Append "ɾ��", adDouble, 1 '-- 1��ʾ�����ɾ�� 2��ʾ������޸���ʱ�� ,0 δ��������
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    Set CopyNewRs = rsNewRs
End Function

Private Sub ColligationTab(Optional blnFillPage As Boolean = True)
'-------------------------------------------------
'�󶨱��������
'-------------------------------------------------
    Dim arrCOL() As String, arrNo() As String
    Dim lngCount As Long
    Dim lngRow As Long, lngCol As Long
    
    
    arrCOL = Split(mstrTabHead, "|")
    If mstrItemNo <> "" Then arrNo = Split(mstrItemNo, "|")
    With VsfData
        .Clear
        .Cols = IIf((UBound(arrCOL) + 1) = 0, RootCol, UBound(arrCOL) + 1)
        .FixedRows = 4
        .FixedCols = 1
        .Rows = 5
         
         '���ز�����
        .ColHidden(c�ļ�ID) = True
        .ColHidden(c����ID) = True
        .ColHidden(c��ҳID) = True
        .ColHidden(cӤ��) = True
        .ColHidden(c��¼ID) = True
        .ColHidden(c����) = True
        .ColHidden(c��Ժ) = True
        .ColHidden(c����ȼ�) = True
        .RowHidden(0) = True
        .RowHidden(1) = True
        .RowHidden(2) = True
        .ColWidth(0) = 250
        .ColWidth(c����) = 1500 + mintBigSize * 1500 / 3
        .ColAlignment(c����) = flexAlignLeftCenter
        .ColAlignment(c����) = flexAlignRightCenter
        .ColAlignment(c����) = flexAlignLeftCenter
        .ColWidth(c���±�ʶ) = 1000

        .FrozenCols = cʱ��
        .SheetBorder = &H40C0&
        
        .RowHeight(-1) = 300 + mintBigSize * 300 / 3
        .FontName = "����"
        .Font.Size = 9 + mintBigSize * 9 / 3
        '������ͷ
        For lngCount = 0 To UBound(arrCOL)
            .TextMatrix(.FixedRows - 1, lngCount) = arrCOL(lngCount)
            If lngCount >= cʱ�� Then
                .ColWidth(lngCount) = 1200 + mintBigSize * 1200 / 3
                .ColAlignment(lngCount) = flexAlignCenterCenter
            End If
        Next lngCount
        
        '����������
        For lngCol = 0 To .Cols - 1
            If lngCol < RootCol Then
                .TextMatrix(0, lngCol) = ""
                .TextMatrix(1, lngCol) = ""
                .TextMatrix(2, lngCol) = ""
            Else
                mrsItems.Filter = "��Ŀ���=" & Val(arrNo(lngCol - RootCol))
                .TextMatrix(0, lngCol) = mrsItems!��Ŀ���
                .TextMatrix(1, lngCol) = Nvl(mrsItems!��Ŀ����, 0) & "|" & Val(Nvl(mrsItems!��ĿС��, 0)) & "|" & Nvl(mrsItems!��Ŀֵ��)
                .TextMatrix(2, lngCol) = Val(Nvl(mrsItems!���ò���, 0))
            End If
        Next lngCol
        
         '�̶��и�ʽΪ����
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
        .Cell(flexcpAlignment, .FixedRows, cʱ��, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
        .Cell(flexcpText, .FixedRows, cʱ��, .Rows - 1, .Cols - 1) = ""
        .Cell(flexcpForeColor, 0, 0, .Rows - 1, .Cols - 1) = &H80000012
        
         If blnFillPage = True Then Call FillPage
    End With
End Sub

Private Sub FillPage()
'-----------------------------------------------------------------------------------------------------------------
'����:��ȡ���������Ĳ����б���Ϣ  ��Ժ�����ڵĲ��� + ���������ڵĲ��� + ���������´��ڳ���37.5�ȵĲ��� + Σ/�ز���
'-----------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim arrFilter() As String
    Dim strFilter As String, strPatient As String
    Dim strOutTime As String
    Dim i As Integer
    Dim strBegin As String, strEnd As String
    Dim strFind As String
    On Error GoTo Errhand
    
    strBegin = Format(Format(CDate(mstrDate) - 3, "YYYY-MM-DD") & " 00:00:00", "YYYY-MM-DD HH:mm:ss")
    strEnd = Format(Format(CDate(mstrDate), "YYYY-MM-DD") & " 23:59:59", "YYYY-MM-DD HH:mm:ss")
    
    'txtFilter.Tag ��ʾ������������
    strFilter = txtFilter.Tag
    If Val(txtFilter.Tag) = 0 Then
       strFilter = "1;1;1;1;1;1;1"
    Else
        strFilter = ";;;;;;"
        arrFilter = Split(strFilter, ";")
        For i = 0 To UBound(Split(txtFilter.Tag, ";"))
            arrFilter(Val(Split(txtFilter.Tag, ";")(i)) - 1) = 1
        Next i
        strFilter = Join(arrFilter, ";")
    End If
    
    arrFilter = Split(strFilter, ";")
    
    strPatient = ""
    '58890:������,2013-02-26,��Ժ���˶�ȡ�����Ż�(������Ժ���˱���в�ѯ)
    '�˴����ڳ�Ժ���˲�������ȡ
    If Val(arrFilter(0)) = 1 Then '��Ժ�����ڵĲ���
        strPatient = "" & _
            " SELECT 1 AS ����,B.����ID, B.��ҳID, A.����, A.�Ա�,B.����, B.סԺ��, lpad(nvl(B.��Ժ����,' '),10,' ') AS ����,0 AS Ӥ��" & _
            " FROM ������Ϣ A,������ҳ B,��Ժ���� C" & _
            " Where A.����ID = B.����ID And A.סԺ����=B.��ҳID And NVL(B.��ҳID, 0) <> 0 " & _
            " AND Nvl(B.����״̬,0)<>5 AND B.���ʱ�� is NULL" & _
            " AND B.��Ժ���� BETWEEN [1] AND [2] And A.����ID=C.����ID And C.����ID=[3]" & _
            IIf(mlng����id = -1, "", " And C.����ID=[4]")
    End If
    
    If Val(arrFilter(1)) = 1 Then '���������ڵĲ���
        If strPatient <> "" Then strPatient = strPatient & " UNION "
        '��ȡ���µ������ɱ�ע�����Ĳ���
        strPatient = strPatient & _
                " SELECT 1 AS ����,B.����ID,B.��ҳID, A.����, A.�Ա�,B.����,B.סԺ��, lpad(nvl(B.��Ժ����,' '),10,' ') AS ����,0 AS Ӥ��" & vbNewLine & _
                " FROM ������Ϣ A,������ҳ B,��Ժ���� F, ���˻����ļ� C ,���˻������� D,���˻�����ϸ E" & vbNewLine & _
                " WHERE A.����ID = B.����ID And A.סԺ����=B.��ҳID And A.����ID=F.����ID AND NVL(B.��ҳID, 0) <> 0 AND F.����ID = [3]" & vbNewLine & _
                " AND NVL(B.����״̬,0)<>5 AND B.���ʱ�� IS NULL" & vbNewLine & _
                " AND B.����ID=C.����ID AND B.��ҳID=C.��ҳID AND C.��ʽID=[5] AND C.ID=D.�ļ�ID AND D.ID=E.��¼ID" & vbNewLine & _
                " AND E.��¼����=4 AND E.��Ŀ����<>'����' AND E.��ֹ�汾 IS NULL" & vbNewLine & _
                " AND D.����ʱ�� BETWEEN [1] AND [2]" & vbNewLine & _
                IIf(mlng����id = -1, "", " And F.����ID=[4]")
                
        If strPatient <> "" Then strPatient = strPatient & " UNION "
        '��ҽ������ȡ����������Ϣ
        strPatient = strPatient & _
                    "SELECT 1 AS ����,B.����ID,B.��ҳID, A.����, A.�Ա�,B.����,B.סԺ��, lpad(nvl(B.��Ժ����,' '),10,' ') AS ����,0 AS Ӥ��" & vbNewLine & _
                    "FROM  ������Ϣ A,������ҳ B,��Ժ���� F,(SELECT D.����ID,D.��ҳID FROM (SELECT DISTINCT A.����ID,A.��ҳID" & vbNewLine & _
                    "           FROM ����ҽ����¼ A, ������ĿĿ¼ B" & vbNewLine & _
                    "           WHERE A.������ĿID = B.ID AND A.������� = 'F' And A.���ID is null And A.ҽ��״̬ in (3,8) AND A.��ʼִ��ʱ�� BETWEEN [1] AND [2]" & vbNewLine & _
                    "           UNION" & vbNewLine & _
                    "           SELECT DISTINCT A.����ID,A.��ҳID FROM ������������¼ A WHERE A.����ʱ�� BETWEEN [1] AND [2]) D GROUP BY D.����ID,D.��ҳID) C" & vbNewLine & _
                    "WHERE A.����ID = B.����ID And A.סԺ����=B.��ҳID And A.����ID=F.����ID AND NVL(B.��ҳID, 0) <> 0 AND F.����ID = [3]" & vbNewLine & _
                    "AND NVL(B.����״̬,0)<>5 AND B.���ʱ�� IS NULL" & vbNewLine & _
                    "AND B.����ID=C.����ID AND B.��ҳID=C.��ҳID" & vbNewLine & _
                    IIf(mlng����id = -1, "", " And F.����ID=[4]")
    End If
    
    If Val(arrFilter(2)) = 1 Then '���������´��ڳ���37.5�ȵĲ���
        If strPatient <> "" Then strPatient = strPatient & " UNION "
        strPatient = strPatient & _
                    " SELECT 1 AS ����,B.����ID,B.��ҳID, A.����, A.�Ա�,B.����,B.סԺ��, lpad(nvl(B.��Ժ����,' '),10,' ') AS ����,0 AS Ӥ��" & vbNewLine & _
                    " FROM ������Ϣ A,������ҳ B,��Ժ���� F, ���˻����ļ� C ,���˻������� D,���˻�����ϸ E" & vbNewLine & _
                    " WHERE A.����ID = B.����ID And A.סԺ����=B.��ҳID And A.����ID=F.����ID AND NVL(B.��ҳID, 0) <> 0 AND F.����ID = [3]" & vbNewLine & _
                    " AND NVL(B.����״̬,0)<>5 AND B.���ʱ�� IS NULL" & vbNewLine & _
                    " AND B.����ID=C.����ID AND B.��ҳID=C.��ҳID AND C.��ʽID=[5] AND C.ID=D.�ļ�ID AND D.ID=E.��¼ID" & vbNewLine & _
                    " AND E.��¼����=1 AND E.��Ŀ���=1 AND LENGTH( TRANSLATE(E.��¼����,'-.0123456789' || E.��¼����,'-.0123456789')) =LENGTH(E.��¼����)" & vbNewLine & _
                    " AND zl_to_number(E.��¼����)>=37.5 AND E.��ֹ�汾 IS NULL" & vbNewLine & _
                    " AND D.����ʱ�� BETWEEN [1] AND [2]" & vbNewLine & _
                    IIf(mlng����id = -1, "", " And F.����ID=[4]")
    End If
    
    If Val(arrFilter(3)) = 1 Then 'Σ/�ز���
        If strPatient <> "" Then strPatient = strPatient & " UNION "
        strPatient = strPatient & _
               " SELECT 1 AS ����,B.����ID,B.��ҳID, A.����, A.�Ա�,B.����,B.סԺ��, lpad(nvl(B.��Ժ����,' '),10,' ') AS ����,0 AS Ӥ��" & _
               " FROM ������Ϣ A,������ҳ B,��Ժ���� F " & _
               " Where A.����ID = b.����ID And A.סԺ����=B.��ҳID And A.����ID=F.����ID And NVL(b.��ҳID, 0) <> 0 And F.����ID = [3]" & _
               " AND Nvl(B.����״̬,0)<>5 AND B.���ʱ�� is NULL" & _
               " AND Instr(',' || 'Σ,��' || ',',','|| B.��ǰ���� || ',')>0 " & _
               IIf(mlng����id = -1, "", " And F.����ID=[4]")
    End If
    
    If Val(arrFilter(4)) = 1 Then 'ת�������ڵĲ���
        If strPatient <> "" Then strPatient = strPatient & " UNION "
        strPatient = strPatient & _
            " SELECT 1 AS ����,B.����ID, B.��ҳID, A.����, A.�Ա�,B.����, B.סԺ��, lpad(nvl(B.��Ժ����,' '),10,' ') AS ����,0 AS Ӥ��" & vbNewLine & _
            " FROM ������Ϣ A,������ҳ B,���˱䶯��¼ C,��Ժ���� F" & vbNewLine & _
            " Where A.����ID = b.����ID And A.סԺ����=B.��ҳID And NVL(b.��ҳID, 0) <> 0 And b.����id = c.����id And b.��ҳid = c.��ҳid " & vbNewLine & _
            " And A.����ID=F.����ID And F.����ID= [3] AND Nvl(B.����״̬,0)<>5 AND B.���ʱ�� is NULL" & vbNewLine & _
            " AND Nvl(c.���Ӵ�λ, 0) = 0 And C.��ֹʱ�� IS null  And C.��ʼԭ�� in (3,15) And C.��ʼʱ��  is not null And B.״̬=0" & vbNewLine & _
            " AND C.��ʼʱ�� BETWEEN [1] AND [2]" & vbNewLine & _
             IIf(mlng����id = -1, "", " And F.����ID=[4]")
    End If
    
    '51286,������,2012-07-11,���"һ�������ϻ���ȼ��Ĳ���"
    If Val(arrFilter(5)) = 1 Then
        If strPatient <> "" Then strPatient = strPatient & " UNION "
        strPatient = strPatient & _
               " SELECT 1 AS ����,B.����ID,B.��ҳID, A.����, A.�Ա�,B.����,B.סԺ��, lpad(nvl(B.��Ժ����,' '),10,' ') AS ����,0 AS Ӥ��" & _
               " FROM ������Ϣ A,������ҳ B,��Ժ���� F " & _
               " Where A.����ID = b.����ID And A.סԺ����=B.��ҳID And A.����ID=F.����ID And NVL(b.��ҳID, 0) <> 0 And F.����ID = [3]" & _
               " AND Nvl(B.����״̬,0)<>5 AND B.���ʱ�� is NULL" & _
               " AND zl_PatitTendGrade(B.����ID,B.��ҳID)<=1 " & _
               IIf(mlng����id = -1, "", " And F.����ID=[4]")
    End If
    
    If mstr�������� = "����" Or mstr�������� = "����" Then
        If Val(arrFilter(6)) = 1 Then '����������ڵĲ���
            If strPatient <> "" Then strPatient = strPatient & " UNION "
            strPatient = strPatient & _
                    " SELECT 1 AS ����,B.����ID,B.��ҳID, A.����, A.�Ա�,B.����,B.סԺ��, lpad(nvl(B.��Ժ����,' '),10,' ') AS ����,0 AS Ӥ��" & vbNewLine & _
                    " FROM ������Ϣ A,������ҳ B,��Ժ���� F, ���˻����ļ� C ,���˻������� D,���˻�����ϸ E" & vbNewLine & _
                    " WHERE A.����ID = B.����ID And A.סԺ����=B.��ҳID And A.����ID=F.����ID AND NVL(B.��ҳID, 0) <> 0 AND F.����ID=[3]" & vbNewLine & _
                    " AND NVL(B.����״̬,0)<>5 AND B.���ʱ�� IS NULL" & vbNewLine & _
                    " AND B.����ID=C.����ID AND B.��ҳID=C.��ҳID AND C.��ʽID=[5] AND C.ID=D.�ļ�ID AND D.ID=E.��¼ID" & vbNewLine & _
                    " AND E.��¼����=4 AND E.��Ŀ����='����' AND E.��ֹ�汾 IS NULL" & vbNewLine & _
                    " AND D.����ʱ�� BETWEEN [1] AND [2]" & vbNewLine & _
                    IIf(mlng����id = -1, "", " And F.����ID=[4]")
        End If
    End If
    
    If strPatient = "" Then Exit Sub
    
    Select Case mintPatiNo
        Case 1
            'ֻ��ȡ���˱���
            strPatient = strPatient
        Case 2
            'ֻ��ȡӤ����Ϣ
            strPatient = " Select B.����,B.����ID,B.��ҳID,NVL(A.Ӥ������,B.����||'֮��'||A.���) AS ����,B.�Ա�,Zl_Age_Calc(0,A.����ʱ��,sysdate) ����,B.סԺ��,lpad(B.����,10,' ') as ����,A.��� AS Ӥ��" & _
              " From ������������¼ A,(" & strPatient & ") B" & _
              " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID"
        Case Else
             '��ȡ���˼��������б�
            strPatient = strPatient & _
                  " UNION " & _
                  " Select B.����,B.����ID,B.��ҳID,NVL(A.Ӥ������,B.����||'֮��'||A.���) AS ����,B.�Ա�,Decode(nvl(A.���,0),0,B.����,Zl_Age_Calc(0,A.����ʱ��,sysdate)) ����,B.סԺ��,lpad(B.����,10,' ') as ����,A.��� AS Ӥ��" & _
                  " From ������������¼ A,(" & strPatient & ") B" & _
                  " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID"
    End Select
   
    mstrSQL = " SELECT  A.����,A.����ID,A.��ҳID,A.Ӥ��,A.����,A.����,lpad(A.����,10,' ') as ����,nvl(zl_PatitTendGrade(A.����ID,A.��ҳID),3) ����ȼ�,C.��Ϣֵ AS ���±�ʶ, MAX(B.ID) AS �ļ�ID,B.��ʼʱ��" & _
              " FROM (" & strPatient & ") A,���˻����ļ� B,������ҳ�ӱ� C" & _
              " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID And A.Ӥ��=B.Ӥ�� And A.����ID=C.����ID(+) And A.��ҳID=C.��ҳID(+) And C.��Ϣ��(+)='���±�ʶ'||DECODE(A.Ӥ��,0,'',A.Ӥ��) " & _
              " And B.�鵵�� is null And B.����ʱ�� is null And B.��ʽID=[5]" & _
              " GROUP BY A.����,A.����ID,A.��ҳID,A.Ӥ��,C.��Ϣֵ,A.���� ,A.����,A.����,B.��ʼʱ��" & _
              " Order by A.����,A.����,A.Ӥ��"
    Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "��ȡ�����嵥", CDate(strBegin), CDate(strEnd), mlng����ID, mlng����id, mlng��ʽID)
     
    strOutTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:mm:ss")
     
    '������ݵ����
    With rsTemp
        Do While Not .EOF
            mblnNullRow = True
            mblnRefreshData = True
            If .AbsolutePosition > VsfData.Rows - VsfData.FixedRows Then VsfData.Rows = VsfData.Rows + 1
            VsfData.TextMatrix(.AbsolutePosition + VsfData.FixedRows - 1, c�ļ�ID) = !�ļ�ID
            VsfData.TextMatrix(.AbsolutePosition + VsfData.FixedRows - 1, c����) = Nvl(!����)
            VsfData.TextMatrix(.AbsolutePosition + VsfData.FixedRows - 1, c����) = IIf(!Ӥ�� > 0, Space(4), "") & !����
            VsfData.TextMatrix(.AbsolutePosition + VsfData.FixedRows - 1, c����) = Nvl(!����)
            VsfData.TextMatrix(.AbsolutePosition + VsfData.FixedRows - 1, c����ID) = !����ID
            VsfData.TextMatrix(.AbsolutePosition + VsfData.FixedRows - 1, c��ҳID) = !��ҳID
            VsfData.TextMatrix(.AbsolutePosition + VsfData.FixedRows - 1, cӤ��) = Nvl(!Ӥ��, 0)
            VsfData.TextMatrix(.AbsolutePosition + VsfData.FixedRows - 1, c����ȼ�) = Val(!����ȼ�)
            VsfData.TextMatrix(.AbsolutePosition + VsfData.FixedRows - 1, c���±�ʶ) = Nvl(!���±�ʶ)
            VsfData.TextMatrix(.AbsolutePosition + VsfData.FixedRows - 1, c����) = Format(!��ʼʱ��, "YYYY-MM-DD HH:mm:ss") & ";" & strOutTime
            VsfData.TextMatrix(.AbsolutePosition + VsfData.FixedRows - 1, c��Ժ) = 0
            .MoveNext
        Loop
    End With
    
    If VsfData.Rows <= VsfData.FixedRows Then VsfData.Rows = VsfData.Rows + 1
    VsfData.RowHeight(-1) = 300 + mintBigSize * 300 / 3
    VsfData.Select VsfData.FixedRows, cʱ��
    '���ñ༭��ɫ
    Call SetTabEditColor
    
    VsfData.Cell(flexcpForeColor, VsfData.FixedRows, c���±�ʶ, VsfData.Rows - 1, c���±�ʶ) = RGB(0, 0, 255)
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub ColligationHistoryTab()
'-------------------------------------------------
'�󶨱��������
'-------------------------------------------------
    Dim arrCOL() As String, arrNo() As String
    Dim lngCount As Long
    Dim lngRow As Long, lngCol As Long
    
    
    arrCOL = Split(mstrTabHead, "|")
    If mstrItemNo <> "" Then arrNo = Split(mstrItemNo, "|")
    With vsfHistory
        .Clear
        .Cols = IIf((UBound(arrCOL) + 1) = 0, RootCol, UBound(arrCOL) + 1)
        .FixedRows = 4
        .FixedCols = 1
        .Rows = 4
         
         '���ز�����
        .ColHidden(c�ļ�ID) = True
        .ColHidden(c����) = True
        .ColHidden(c����) = True
        .ColHidden(c����) = True
        .ColHidden(c����ID) = True
        .ColHidden(c��ҳID) = True
        .ColHidden(cӤ��) = True
        .ColHidden(c��¼ID) = True
        .ColHidden(c����ȼ�) = True
        .ColHidden(c���±�ʶ) = True
        .ColHidden(c��Ժ) = True
        
        .RowHidden(0) = True
        .RowHidden(1) = True
        .RowHidden(2) = True
        .ColWidth(0) = 250
        
        .FrozenCols = cʱ��
        .SheetBorder = &H40C0&
        
        .RowHeight(-1) = 300 + mintBigSize * 300 / 3
        .FontName = "����"
        .Font.Size = 9 + mintBigSize * 9 / 3
        '������ͷ
        For lngCount = 0 To UBound(arrCOL)
            .TextMatrix(.FixedRows - 1, lngCount) = arrCOL(lngCount)
            If lngCount >= c���� Then
                .ColWidth(lngCount) = 1200 + mintBigSize * 1200 / 3
                .ColAlignment(lngCount) = flexAlignCenterCenter
            End If
        Next lngCount
        
        '����������
        For lngCol = 0 To .Cols - 1
            If lngCol < RootCol Then
                .TextMatrix(0, lngCol) = ""
                .TextMatrix(1, lngCol) = ""
                .TextMatrix(2, lngCol) = ""
            Else
                mrsItems.Filter = "��Ŀ���=" & Val(arrNo(lngCol - RootCol))
                .TextMatrix(0, lngCol) = mrsItems!��Ŀ���
                .TextMatrix(1, lngCol) = Nvl(mrsItems!��Ŀ����, 0) & "|" & Val(Nvl(mrsItems!��ĿС��, 0)) & "|" & Nvl(mrsItems!��Ŀֵ��)
                .TextMatrix(2, lngCol) = Val(Nvl(mrsItems!���ò���, 0))
            End If
        Next lngCol
        
         '�̶��и�ʽΪ����
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
        If .Rows > .FixedRows Then
            .Cell(flexcpAlignment, .FixedRows, c����, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
            .Cell(flexcpText, .FixedRows, c����, .Rows - 1, .Cols - 1) = ""
        End If
        .Cell(flexcpForeColor, 0, 0, .Rows - 1, .Cols - 1) = &H80000012
    End With
End Sub

Private Sub cmdFilterCancel_Click()
    picFilter.Visible = False
End Sub

Private Sub cmdFilterOK_Click()
    Dim i As Integer
    Dim strValue As String
    Dim arrValue() As String, ArrCode() As String
    
    If lstFilter.SelCount = 0 Then
        MsgBox "������ѡ��һ�ֹ���������", vbInformation, gstrSysName
        lstFilter.SetFocus
        Exit Sub
    End If
    
    If lstFilter.Selected(0) = True Then
        txtFilter.Text = "ȫ��"
        txtFilter.Tag = 0
    Else
        txtFilter.Text = ""
        txtFilter.Tag = ""
        For i = 1 To lstFilter.ListCount - 1
            If lstFilter.Selected(i) Then
                txtFilter.Text = txtFilter.Text & ";" & lstFilter.List(i)
                txtFilter.Tag = txtFilter.Tag & ";" & lstFilter.ItemData(i)
            End If
        Next
        txtFilter.Text = Mid(txtFilter.Text, 2)
        txtFilter.Tag = Mid(txtFilter.Tag, 2)
    End If
    
    txtFilter.SetFocus
    picFilter.Visible = False
    
    '�������������Ϣ
    If Val(txtFilter.Tag) = 0 Then
        strValue = "1;1;1;1;1;1;1"
    Else
        strValue = "0;0;0;0;0;0;0"
        arrValue = Split(strValue, ";")
        ArrCode = Split(txtFilter.Tag, ";")
        For i = 0 To UBound(ArrCode)
            arrValue(Val(ArrCode(i)) - 1) = 1
        Next i
        strValue = Join(arrValue, ";")
    End If
    
    Call zlDatabase.SetPara("���µ���������", strValue, glngSys, 1255)
    
    '��ʼ���¼���������Ϣ
    'Call cmdFilter_Click
End Sub

Private Sub cmdFilterUserCancle_Click()
    picPati.Visible = False
    VsfData.SetFocus
End Sub

Private Sub cmdFilterUserOk_Click()
    '��Ӳ���
    Dim rsTemp As New ADODB.Recordset
    Dim objRow As ReportRow
    Dim lngLoop As Long
    Dim strPatient As String, strSQL As String
    Dim lngRow As Long, lngTempRow As Long
    Dim strCurDate As String, strInTime As String, strOutTime As String
    Dim blnNullRow As Long, blnOut As Boolean
    
    '������Ϣ����
    Dim lng����ID As Long, lng��ҳID As Long, str���� As String, str�Ա� As String, str���� As String, strסԺ�� As String, str���� As String, intBaby As Integer
    
    strPatient = ""
    For Each objRow In rptPati.Rows
        If objRow.GroupRow And objRow.Childs.Count > 0 Then
            For lngLoop = 0 To objRow.Childs.Count - 1
                If Not (objRow.Childs(lngLoop).Record Is Nothing) Then
                    If objRow.Childs(lngLoop).Record.Item(c_ѡ��).Checked = True Then
                        lng����ID = Val(objRow.Childs(lngLoop).Record.Item(C_����ID).Value)
                        lng��ҳID = Val(objRow.Childs(lngLoop).Record.Item(c_��ҳID).Value)
                        str���� = objRow.Childs(lngLoop).Record.Item(c_����).Value
                        str�Ա� = IIf(Val(objRow.Childs(lngLoop).Record.Item(c_ͼ��).Icon) = 1, "��", "Ů")
                        str���� = objRow.Childs(lngLoop).Record.Item(c_����).Value
                        strסԺ�� = Val(objRow.Childs(lngLoop).Record.Item(c_סԺ��).Value)
                        str���� = objRow.Childs(lngLoop).Record.Item(c_����).Value
                        strOutTime = objRow.Childs(lngLoop).Record.Item(c_��Ժ����).Value
                        intBaby = 0
                        
                        strSQL = ""
                        strSQL = "SELECT 1 ����,"
                        strSQL = strSQL & lng����ID & " ����ID,"
                        strSQL = strSQL & lng��ҳID & " ��ҳID,"
                        strSQL = strSQL & "'" & str���� & "' ����,"
                        strSQL = strSQL & "'" & str�Ա� & "' �Ա�,"
                        strSQL = strSQL & "'" & str���� & "' ����,"
                        strSQL = strSQL & "" & strסԺ�� & " סԺ��,"
                        strSQL = strSQL & "'" & str���� & "' ����,"
                        strSQL = strSQL & "" & intBaby & " Ӥ��,"
                        strSQL = strSQL & "'" & strOutTime & "' ��Ժ����"
                        strSQL = strSQL & " FROM dual"
                        
                        strPatient = strPatient & vbCrLf & IIf(strPatient = "", strSQL, " UNION " & vbCrLf & strSQL)
                    End If
                End If
            Next lngLoop
        End If
    Next
    
    '��������PIC
    Call InitCons
    
    If Trim(strPatient) = "" Then Exit Sub
    On Error GoTo Errhand:
    
    '�����δ����д˴���Ҫ�������Ϣ
    If Not mblnNullRow Then
        mstrDate = Format(dtpDate.Value, "YYYY-MM-DD")
        Call InitVariable
        Call zlRefreshDate(False)
        mblnInit = True
    End If
    
    blnNullRow = mblnNullRow
    
    strPatient = "SELECT ����,����ID,��ҳID,����,�Ա�,����,סԺ��,lpad(����,10,' ') as  ����,Ӥ��,��Ժ���� FROM (" & strPatient & ")"
    
    Select Case mintPatiNo
        Case 1
            strPatient = strPatient
        Case 2
            strPatient = " Select B.����,B.����ID,B.��ҳID,NVL(A.Ӥ������,B.����||'֮��'||A.���) AS ����,B.�Ա�,Zl_Age_Calc(0,A.����ʱ��,sysdate) ����,B.סԺ��,lpad(B.����,10,' ') as ����,A.��� AS Ӥ��,B.��Ժ����" & _
                  " From ������������¼ A,(" & strPatient & ") B" & _
                  " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID"
        Case Else
            '��ȡ���˺��������б�
            strPatient = strPatient & _
                  " UNION " & _
                  " Select B.����,B.����ID,B.��ҳID,NVL(A.Ӥ������,B.����||'֮��'||A.���) AS ����,B.�Ա�,Decode(nvl(A.���,0),0,B.����,Zl_Age_Calc(0,A.����ʱ��,sysdate)) ����,B.סԺ��,lpad(B.����,10,' ') as ����,A.��� AS Ӥ��,B.��Ժ����" & _
                  " From ������������¼ A,(" & strPatient & ") B" & _
                  " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID"
    End Select

     mstrSQL = " SELECT  A.����, A.����ID,A.��ҳID,A.Ӥ��,nvl(zl_PatitTendGrade(A.����ID,A.��ҳID),3) ����ȼ�,C.��Ϣֵ AS ���±�ʶ,A.����,A.����,lpad(A.����,10,' ') as ����,A.��Ժ����,MAX(B.ID) AS �ļ�ID,B.��ʼʱ��" & _
              " FROM (" & strPatient & ") A,���˻����ļ� B,������ҳ�ӱ� C" & _
              " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID And A.Ӥ��=B.Ӥ�� And A.����ID=C.����ID(+) And A.��ҳID=C.��ҳID(+) And C.��Ϣ��(+)='���±�ʶ'||DECODE(A.Ӥ��,0,'',A.Ӥ��) " & _
              " And B.�鵵�� is null And B.����ʱ�� is null And B.��ʽID=[1]" & _
              " GROUP BY A.����,A.����ID,A.��ҳID,A.Ӥ��,C.��Ϣֵ,A.����,A.����,A.����,A.��Ժ����,B.��ʼʱ��" & _
              " Order by A.����,A.����,A.Ӥ��"
     Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "��ȡ�����嵥", mlng��ʽID)
     
     strCurDate = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:mm:ss")
    '������ݵ����
    lngRow = 0
    With rsTemp
        Do While Not .EOF
            blnOut = True
            mblnNullRow = True
            
            If blnNullRow = False Then
                If .AbsolutePosition > VsfData.Rows - VsfData.FixedRows Then VsfData.Rows = VsfData.Rows + 1
                lngTempRow = .AbsolutePosition + VsfData.FixedRows - 1
            Else
                If VsfData.Rows > VsfData.FixedRows Then
                    If VsfData.TextMatrix(VsfData.Rows - 1, c�ļ�ID) <> 0 Then
                        VsfData.Rows = VsfData.Rows + 1
                    End If
                Else
                    VsfData.Rows = VsfData.Rows + 1
                End If
                
                lngTempRow = VsfData.Rows - 1
            End If
            strOutTime = Trim(Nvl(!��Ժ����))
            If strOutTime = "" Then strOutTime = strCurDate: blnOut = False
            
            VsfData.TextMatrix(lngTempRow, c�ļ�ID) = !�ļ�ID
            VsfData.TextMatrix(lngTempRow, c����) = Nvl(!����)
            VsfData.TextMatrix(lngTempRow, c����) = IIf(!Ӥ�� > 0, Space(4), "") & !����
            VsfData.TextMatrix(lngTempRow, c����) = Nvl(!����)
            VsfData.TextMatrix(lngTempRow, c����ID) = !����ID
            VsfData.TextMatrix(lngTempRow, c��ҳID) = !��ҳID
            VsfData.TextMatrix(lngTempRow, cӤ��) = Nvl(!Ӥ��, 0)
            VsfData.TextMatrix(lngTempRow, c����ȼ�) = Val(!����ȼ�)
            VsfData.TextMatrix(lngTempRow, c���±�ʶ) = Nvl(!���±�ʶ)
            VsfData.TextMatrix(lngTempRow, c����) = Format(!��ʼʱ��, "YYYY-MM-DD HH:mm:ss") & ";" & strOutTime
            VsfData.TextMatrix(lngTempRow, c��Ժ) = IIf(blnOut = True, 1, 0)
            
            If lngRow = 0 Then lngRow = lngTempRow
            .MoveNext
        Loop
    End With
    
    If VsfData.Rows <= VsfData.FixedRows Then VsfData.Rows = VsfData.Rows + 1
    VsfData.RowHeight(-1) = 300 + mintBigSize * 300 / 3
    
    If lngRow = 0 Then lngRow = VsfData.Rows - 1
    VsfData.Select lngRow, cʱ��

    '���ñ༭��ɫ
    Call SetTabEditColor
    '�������ݼ�
    If Not mblnSaveData Then
        Call Data_Save
    End If

    VsfData.Cell(flexcpForeColor, VsfData.FixedRows, c���±�ʶ, VsfData.Rows - 1, c���±�ʶ) = RGB(0, 0, 255)
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub SetTabEditColor()
'-----------------------------------------------
'����:�жϸò��˵Ļ���ȼ��Ƿ���ʹ��ĳ����Ŀ
'-----------------------------------------------
    Dim intRow As Integer, intCOl As Integer
    Dim int����ȼ� As Integer, intӤ�� As Integer
    Dim lngItemNO As Long
    Dim blnTrue As Boolean
    
    For intRow = VsfData.FixedRows To VsfData.Rows - 1
        If VsfData.RowHidden(intRow) = False And Val(VsfData.TextMatrix(intRow, c�ļ�ID)) <> 0 Then
            int����ȼ� = Val(VsfData.TextMatrix(intRow, c����ȼ�))
            intӤ�� = Val(VsfData.TextMatrix(intRow, cӤ��))
            For intCOl = RootCol To VsfData.Cols - 1
                blnTrue = False
                lngItemNO = Val(VsfData.TextMatrix(0, intCOl))
                mrsItems.Filter = 0
                mrsItems.Filter = "��Ŀ���=" & lngItemNO & " And ����ȼ�>=" & int����ȼ�
                If mrsItems.RecordCount > 0 Then
                    VsfData.Cell(flexcpBackColor, intRow, intCOl, intRow, intCOl) = &H80000005
                Else
                    VsfData.Cell(flexcpBackColor, intRow, intCOl, intRow, intCOl) = &H80000016
                    blnTrue = True
                End If
                '����Ƿ������ڴ˲���
                If Not blnTrue Then
                    If Val(VsfData.TextMatrix(2, intCOl)) = 1 Then
                        VsfData.Cell(flexcpBackColor, intRow, intCOl, intRow, intCOl) = IIf(intӤ�� = 0, &H80000005, &H80000016)
                    ElseIf VsfData.TextMatrix(2, intCOl) = 2 Then
                        VsfData.Cell(flexcpBackColor, intRow, intCOl, intRow, intCOl) = IIf(intӤ�� <> 0, &H80000005, &H80000016)
                    End If
                End If
            Next intCOl
        End If
    Next intRow
    VsfData.Cell(flexcpForeColor, VsfData.FixedRows, c���±�ʶ, VsfData.Rows - 1, c���±�ʶ) = RGB(0, 0, 255)
End Sub

Private Sub CmdRef_Click()
    Set mrsPati = New ADODB.Recordset
    If mrsPati.State = adStateOpen Then mrsPart.Close
    Call cmdAddUser_Click
End Sub

Private Sub cmdSift_Click()
    Dim i As Integer
    
    For i = 0 To lstFilter.ListCount - 1
        If Val(txtFilter.Tag) = 0 Then
            lstFilter.Selected(i) = True
        ElseIf InStr(1, ";" & txtFilter.Tag & ";", ";" & lstFilter.ItemData(i) & ";") <> 0 Then
            lstFilter.Selected(i) = True
        Else
            lstFilter.Selected(i) = False
        End If
    Next i
    lstFilter.ListIndex = 0
    With picFilter
        .Top = picMain.Top
        .Left = txtFilter.Left + 60
        .Visible = True
        .ZOrder 0
    End With
End Sub

Private Sub dtpB_Change(Index As Integer)
'ʱ�䷶Χ�ı�ʱˢ��
    If dtpB(Index).Value >= dtpE(Index).Value Then
        RaiseEvent AfterRowColChange("ʱ�䷶Χ�Ŀ�ʼʱ��ӦС�ڽ���ʱ��", True)
        dtpB(Index).Value = dtpB(Index).Tag
        dtpB(Index).SetFocus: Exit Sub
    Else
        dtpB(Index).Tag = dtpB(Index).Value
        If Index = 0 Then mdtOutBegin = dtpB(Index).Value
    End If
End Sub

Private Sub dtpE_Change(Index As Integer)
    If dtpB(Index).Value >= dtpE(Index).Value Then
        RaiseEvent AfterRowColChange("ʱ�䷶Χ�Ŀ�ʼʱ��ӦС�ڽ���ʱ��", True)
        dtpE(Index).Value = dtpE(Index).Tag
        dtpE(Index).SetFocus: Exit Sub
    Else
        dtpE(Index).Tag = dtpE(Index).Value
        If Index = 0 Then mdtOutEnd = dtpE(Index).Value
    End If
End Sub

Private Sub dtpDate_Change()
    Dim blnCancle As Boolean
    Call dtpDate_Validate(blnCancle)
    If blnCancle = True Then
        dtpDate.SetFocus
    End If
End Sub

Private Sub dtpDate_GotFocus()
    If Not mblnDateFouces Then Call InitCons
End Sub

Private Sub dtpDate_Validate(Cancel As Boolean)
    If Not mblnInit Then Exit Sub
    If CheckEditData Then
        Cancel = True
        dtpDate.Value = Format(mstrDate, "YYYY-MM-DD")
        Exit Sub
    End If
    mstrDate = Format(dtpDate.Value, "YYYY-MM-DD")
End Sub

Private Sub lblCheck_DblClick()
    Call picInput_KeyPress(vbKeySpace)
End Sub

Private Sub lstFilter_ItemCheck(Item As Integer)
    Dim i As Integer
    
    If Item = 0 Then
        For i = 1 To lstFilter.ListCount - 1
            lstFilter.Selected(i) = lstFilter.Selected(0)
        Next
    ElseIf Not lstFilter.Selected(Item) Then
        lstFilter.Selected(0) = False
    ElseIf lstFilter.SelCount = lstFilter.ListCount - 1 Then
        lstFilter.Selected(0) = True
    End If
End Sub

Private Sub lstFilter_LostFocus()
    If Not UserControl.ActiveControl Is cmdFilterOK _
        And Not UserControl.ActiveControl Is cmdFilterCancel _
        And Not UserControl.ActiveControl Is lstFilter _
        And Not UserControl.ActiveControl Is picFilter Then picFilter.Visible = False: mblnDateFouces = False
End Sub

Private Sub lstNote_DblClick()
    Call lstNote_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub lstNote_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngRow As Long, lngCol As Long, lngItemNO As Long
    Dim strNote As String
    Dim intCount As Integer, intCOl As Integer, intCols As Integer
    Dim intStartCol As Integer, intEndCol As Integer
    Dim blnAll As Boolean
    
    If KeyCode = vbKeyReturn Then
        If Shift <> 0 Then Exit Sub
        If picInput.Visible = False Then Exit Sub
        
        lngRow = VsfData.Row 'Val(Split(txtInput.Tag, "|")(0))
        lngCol = VsfData.Col 'Val(Split(txtInput.Tag, "|")(1))
        strNote = lstNote.Text
        
        VsfData.TextMatrix(lngRow, lngCol) = strNote
        txtInput.Text = strNote
        mrsItems.Filter = 0
        
        '��������������Ƿ�����ֵ
        intStartCol = RootCol
        intCount = 0
        intCols = 0
        For intCOl = intStartCol To VsfData.Cols - 1
            lngItemNO = Val(VsfData.TextMatrix(0, intCOl))
            mrsItems.Filter = "��Ŀ���=" & lngItemNO
            If Trim(Nvl(mrsItems!������)) = "1)����������Ŀ" Then
                If Trim(VsfData.TextMatrix(lngRow, intCOl)) = "" Then
                    intCount = intCount + 1
                End If
                intCols = intCols + 1
                intEndCol = intCOl
            End If
        Next intCOl
        
        'ѭ����ֵ
        If intCount = intCols - 1 Then
            For intCOl = intStartCol To intEndCol
                VsfData.TextMatrix(lngRow, intCOl) = strNote
            Next intCOl
            blnAll = True
        Else
            intCount = 0
            intCols = 1
            blnAll = False
        End If
        
        If blnAll = True Then
            '��λ����һ��������Ŀ
            VsfData.Col = intStartCol
        Else
            VsfData.Col = lngCol
        End If
        
        For intCOl = 1 To intCols
            picInput.Tag = ""
            mblnDateFouces = True
            Call MoveNextCell(vbKeyReturn)
        Next intCOl
        
    ElseIf KeyCode = vbKeyEscape And Shift = 0 Then
        If picInput.Visible = True Then picInput.SetFocus
    End If
End Sub

Private Sub lstNote_LostFocus()
    Call lstNote_KeyDown(vbKeyEscape, 0)
End Sub



Private Sub picDouble_GotFocus()
    If picDouble.Visible = True Then txtUpInput.SetFocus
End Sub

Private Sub picFilter_GotFocus()
    lstFilter.SetFocus
End Sub

Private Sub picFilter_Resize()
    On Error Resume Next
    lstFilter.Left = -15
    lstFilter.Top = -15
    lstFilter.Width = picFilter.Width
    
    cmdFilterCancel.Left = picFilter.ScaleWidth - cmdFilterCancel.Width - 100
    cmdFilterOK.Left = cmdFilterCancel.Left - cmdFilterOK.Width - 60
    
    cmdFilterOK.Top = lstFilter.Height + (picFilter.ScaleHeight - lstFilter.Height - cmdFilterOK.Height) / 2
    cmdFilterCancel.Top = cmdFilterOK.Top
End Sub

Private Sub picHistory_Resize()
    txt��ʾ����.Left = lblDay.Left + lblDay.Width + TextWidth("��")
    lblDay.Top = txt��ʾ����.Top + (txt��ʾ����.Height - lblDay.Height) \ 2
    lblDayInfo.Top = lblDate.Top
    lblDayInfo.Left = txt��ʾ����.Left + txt��ʾ����.Width + TextWidth("��")
End Sub

Private Sub picInput_DblClick()
    Call picInput_KeyPress(vbKeySpace)
End Sub

Private Sub picInput_GotFocus()
    If picInput.Visible = True And txtInput.Visible = True Then txtInput.SetFocus
End Sub

Private Sub picInput_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        picInput.Visible = False
        lstNote.Visible = False
        picInput.Tag = ""
        txtInput.Tag = ""
        txtInput.Text = ""
        lstNote.Tag = ""
        mintType = 0
        mblnShow = False
        VsfData.SetFocus
    ElseIf KeyAscii = vbKeySpace Then
        If lblCheck.Caption = "��" Then
            lblCheck.Caption = ""
        Else
            lblCheck.Caption = "��"
        End If
    ElseIf KeyAscii = vbKeyReturn Then
        If txtInput.Visible = False Then
            mblnDateFouces = True
            Call VsfData_KeyDown(vbKeyReturn, 0)
        End If
    ElseIf KeyAscii = vbKeyLeft Then
        If txtInput.Visible = False Then
            mblnDateFouces = True
            Call MoveNextCell(KeyAscii)
        End If
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

Private Sub rptPati_ItemCheck(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Dim objRow As ReportRow
    Dim lngLoop As Long
    Dim blnAll As Boolean
    
    If Item.Index = c_ѡ�� Then
        For Each objRow In rptPati.Rows
            If objRow.GroupRow And objRow.Childs.Count > 0 Then
                For lngLoop = 0 To objRow.Childs.Count - 1
                    If Not (objRow.Childs(lngLoop).Record Is Nothing) Then
                        If Trim(objRow.Childs(lngLoop).Record.Item(c_��Ժ����).Value) <> "" Then Exit For
                        blnAll = True
                        If objRow.Childs(lngLoop).Record.Item(c_ѡ��).Checked = False Then
                            blnAll = False
                            GoTo NextCheck
                        End If
                    End If
                Next lngLoop
            End If
        Next
    End If
NextCheck:
    mblnChkClick = True
    chkSwitch.Value = IIf(blnAll = True, 1, 0)
End Sub

Private Sub rptPati_LostFocus()
    If Not UserControl.ActiveControl Is cmdFilterUserOk _
        And Not UserControl.ActiveControl Is cmdFilterUserCancle _
        And Not UserControl.ActiveControl Is rptPati _
        And Not UserControl.ActiveControl Is picPati Then picPati.Visible = False: mblnDateFouces = False
End Sub

Private Sub rptPati_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    '��Ӳ�����Ϣ
    If Not Row.Record Is Nothing Then
        Row.Record.Item(c_ѡ��).Checked = True
        Call cmdFilterUserOk_Click
    End If
End Sub

Private Sub txtChange_KeyPress(KeyAscii As Integer)
    If InStr("1234567890", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then KeyAscii = 0
    If KeyAscii <> vbKeyReturn Then Exit Sub
    mintChange = Val(txtChange.Text)
End Sub

Private Sub txtChange_GotFocus()
    Call zlControl.TxtSelAll(txtChange)
End Sub

Private Sub txtChange_Validate(Cancel As Boolean)
    If Val(txtChange.Text) > 30 Then
        RaiseEvent AfterRowColChange("ת�������������ܳ���30��", True)
        Cancel = True
    Else
        mintChange = Val(txtChange.Text)
    End If
End Sub

Private Sub txtDnInput_GotFocus()
    Call zlControl.TxtSelAll(txtDnInput)
End Sub

Private Sub txtDnInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Shift = vbCtrlMask Then
            Exit Sub
        Else
            Call VsfData_KeyDown(KeyCode, Shift)
        End If
    End If
    
    If KeyCode = vbKeyLeft Then
        If txtDnInput.SelStart = 0 Then
            txtUpInput.SetFocus
        End If
    End If
End Sub

Private Sub txtDnInput_KeyPress(KeyAscii As Integer)
    Call txtUpInput_KeyPress(KeyAscii)
End Sub

Private Sub txtDnInput_LostFocus()
    mblnDateFouces = False
End Sub





Private Sub txtInput_GotFocus()
    Call zlControl.TxtSelAll(txtInput)
    lstNote.Visible = False
End Sub

Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Shift = vbCtrlMask Then
            Exit Sub
        Else
            Call VsfData_KeyDown(KeyCode, Shift)
        End If
    End If
    
    If KeyCode = vbKeyLeft Then
        Call MoveNextCell(vbKeyLeft)
    End If
    
    If KeyCode = vbKeyDown Then '��ʾδ��˵����Ϣ
        If picInput.Visible = False Or txtInput.Visible = False Then Exit Sub
        If VsfData.Col < RootCol Or VsfData.Col > VsfData.Cols - 2 Then Exit Sub
        If InStr(1, ",0,9,", "," & mint������Դ & ",") = 0 Then Exit Sub
        
        With lstNote
            .Top = picInput.Top + picInput.Height
            .Left = picInput.Left
            .FontName = VsfData.FontName
            .Font.Size = VsfData.Font.Size
            .Width = LenB(StrConv(.List(.ListCount \ 2), vbFromUnicode)) * 160 + 500
            If .Width < picInput.Width Then .Width = picInput.Width
            .Height = .ListCount * 210 + 30
            If .Height + .Top + picMain.Top > ScaleHeight Then
                .Top = ScaleHeight - picMain.Top - .Height
            End If
            .Visible = True
            lstNote.SetFocus
        End With
    End If
    
    '���ر༭��
    If KeyCode = vbKeyEscape And Shift = 0 Then
        Call picInput_KeyPress(vbKeyEscape)
    End If
End Sub

Private Sub lstSelect_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Shift = vbShiftMask Then Exit Sub
        mblnDateFouces = True
        Call VsfData_KeyDown(KeyCode, Shift)
    ElseIf KeyCode = vbKeyLeft Then
        mblnDateFouces = True
        Call MoveNextCell(vbKeyLeft)
    ElseIf Index = 0 And Shift = vbShiftMask And KeyCode = vbKeyUp Then
        KeyCode = 0
        txtLst.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        lstSelect(Index).Visible = False
        If Index = 0 Then
            PicLst.Visible = False
            txtLst.Visible = False
        End If
        mblnShow = False
        VsfData.SetFocus
    End If
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

Private Sub txtInput_LostFocus()
    mblnDateFouces = False
End Sub

Private Sub txtLst_GotFocus()
    PicLst.Tag = 1
    Call zlControl.TxtSelAll(txtLst)
    lstSelect(0).ListIndex = -1
End Sub

Private Sub txtLst_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim blnAllow As Boolean
    
    If KeyCode = vbKeyReturn And Shift = vbShiftMask Then Exit Sub
    If KeyCode = vbKeyReturn Then
        mblnDateFouces = True
        Call VsfData_KeyDown(vbKeyReturn, 0)
    ElseIf KeyCode = vbKeyLeft And txtLst.SelStart = 0 Then
        mblnDateFouces = True
        Call MoveNextCell(KeyCode)
    ElseIf Shift = vbShiftMask And KeyCode = vbKeyDown Then
        KeyCode = 0
        lstSelect(0).SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        lstSelect(0).Visible = False
        txtLst.Visible = False
        PicLst.Visible = False
        mblnShow = False
        VsfData.SetFocus
    End If
End Sub

Private Sub txtUpInput_GotFocus()
    Call zlControl.TxtSelAll(txtUpInput)
End Sub

Private Sub txtUpInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Shift = vbCtrlMask Then
            Exit Sub
        Else
            txtDnInput.SetFocus
        End If
    End If
    
    If KeyCode = vbKeyLeft And txtUpInput.SelStart = 0 Then
        Call MoveNextCell(vbKeyLeft)
    End If
End Sub

Private Sub txtUpInput_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        picDouble.Visible = False
        picDouble.Tag = ""
        mblnShow = False
        VsfData.SetFocus
    End If
End Sub

Private Sub txtUpInput_LostFocus()
    mblnDateFouces = False
End Sub

Private Sub txt��ʾ����_GotFocus()
    Call zlControl.TxtSelAll(txt��ʾ����)
End Sub

Private Sub txt��ʾ����_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim blnCancel As Boolean
    If KeyCode = vbKeyReturn Then Call txt��ʾ����_Validate(blnCancel)
End Sub

Private Sub txt��ʾ����_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = vbKeyBack) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt��ʾ����_Validate(Cancel As Boolean)
    If Val(txt��ʾ����.Text) = Val(txt��ʾ����.Tag) Then Exit Sub
    txt��ʾ����.Tag = txt��ʾ����.Text
    Call RefreshHistoryData(VsfData.Row)
End Sub

Private Sub UserControl_Initialize()
    '��ʼ������ѡ����
    Dim objCol As ReportColumn
    With rptPati
        Set objCol = .Columns.Add(c_ѡ��, "", 18, False): objCol.AllowDrag = False
        Set objCol = .Columns.Add(c_ͼ��, "", 18, False): objCol.Alignment = xtpAlignmentCenter: objCol.AllowDrag = False
        Set objCol = .Columns.Add(c_����, "����", 0, True)
        Set objCol = .Columns.Add(c_״̬, "״̬", 0, True)
        Set objCol = .Columns.Add(c_����, "����", 40, True)
        Set objCol = .Columns.Add(C_����ID, "����ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(c_��ҳID, "��ҳID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(c_����, "����", 60, True)
        Set objCol = .Columns.Add(c_����, "����", 60, True)
        Set objCol = .Columns.Add(c_סԺ��, "סԺ��", 60, True)
        Set objCol = .Columns.Add(c_��Ժ����, "��Ժ����", 120, True)
        Set objCol = .Columns.Add(c_��Ժ����, "��Ժ����", 120, True)
        For Each objCol In .Columns
            If objCol.Index <> c_ѡ�� Then
                objCol.Editable = False
            Else
                objCol.Sortable = True
                objCol.Editable = True
            End If
            objCol.Groupable = (objCol.Index = c_״̬)
        Next
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�в���..."
        End With
        
        .PreviewMode = False
        .AllowColumnSort = True
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        .SetImageList UserControl.imgRPT
        
        .GroupsOrder.DeleteAll
        .GroupsOrder.Add .Columns.Find(c_����)
        .GroupsOrder(0).SortAscending = True
        .SortOrder.Add .Columns.Find(c_����)
    End With
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Call InitCons
End Sub

Private Sub UserControl_Terminate()
    Dim strValue As String
    Dim i As Integer
    Dim arrValue() As String, ArrCode() As String
    
    mstrNote = ""
    If Not (mrsItems Is Nothing) Then Set mrsItems = Nothing
    If Not (mrsPati Is Nothing) Then Set mrsPati = Nothing
    If Not (mrsCell Is Nothing) Then Set mrsCell = Nothing
    If Not (mrsPart Is Nothing) Then Set mrsPart = Nothing
    If Not (mrsCopy Is Nothing) Then Set mrsCopy = Nothing
    If Not (mrsData Is Nothing) Then Set mrsData = Nothing
    '�������������Ϣ
'    If Val(txtFilter.Tag) = 0 Then
'        strValue = "1;1;1;1"
'    Else
'        strValue = "0;0;0;0"
'        arrValue = Split(strValue, ";")
'        ArrCode = Split(txtFilter.Tag, ";")
'        For i = 0 To UBound(ArrCode)
'            arrValue(Val(ArrCode(i)) - 1) = 1
'        Next i
'        strValue = Join(arrValue, ";")
'    End If
    
    'Call zlDatabase.SetPara("���µ���������", strValue, glngSys, 1255)
End Sub

Private Sub VsfData_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strInfo As String
    Dim strCols As String
    Dim intMax As Integer
    Dim lngItemNO As Long
    Dim strText As String, strPart As String, strKey As String
    Dim lngCol As Long
    Dim cbrControl As CommandBarControl
    Dim blnCheck As Boolean
    
    If mblnInit = False Then Exit Sub
    If OldRow = NewRow And OldCol = NewCol Then Exit Sub
    
    Call AdjustRowFlag(VsfData, NewRow)
    mblnClearRow = False
    
    If NewRow >= VsfData.FixedRows Then
        For lngCol = cʱ�� + 1 To VsfData.Cols - 1
            If Trim(VsfData.TextMatrix(NewRow, lngCol)) <> "" Then
                mblnClearRow = True
                Exit For
            End If
        Next lngCol
    End If
        
    If NewCol >= RootCol And NewRow >= VsfData.FixedRows Then
        lngItemNO = Val(VsfData.TextMatrix(0, NewCol))
    Else
        If NewCol <> c���±�ʶ Then
            Call AddActiveMenu(0)
            GoTo ErrInfo
        End If
    End If
    '��ʾ��ǰ��Ŀ�������Ϣ
    mrsItems.Filter = 0
    mrsItems.Filter = "��Ŀ���=" & lngItemNO
    If mrsItems.RecordCount <> 0 Then
        If Nvl(mrsItems!��Ŀֵ��) <> "" Then
            If mrsItems!��Ŀ���� = 0 Then
                strInfo = "��Ч��Χ:" & Split(mrsItems!��Ŀֵ��, ";")(0) & "��" & Split(mrsItems!��Ŀֵ��, ";")(1)
            Else
                strInfo = "��Ч��Χ:" & mrsItems!��Ŀֵ��
            End If
        Else
            strInfo = ""
        End If
        
        If lngItemNO = gint���� Then
            strInfo = strInfo & Space(4) & "������:38/37"
        ElseIf lngItemNO = gint���� And mint����Ӧ�� = 2 Then
            strInfo = strInfo & Space(4) & "��������:100/120"
        ElseIf lngItemNO = 4 Then
            strInfo = strInfo & Space(4) & "����ѹ/����ѹ:110/90"
        End If
        
        If Trim(Nvl(mrsItems!������)) = "1)����������Ŀ" Then
             strInfo = strInfo & Space(4) & "��������δ��˵��ѡ��"
        End If
        
        '����������Ŀ���в�λ��Ϣ
        If Trim(Nvl(mrsItems!������)) <> "1)����������Ŀ" Then
             lngItemNO = 0
        Else
            If Val(VsfData.TextMatrix(VsfData.Row, c����ID)) = 0 Or Val(VsfData.TextMatrix(VsfData.Row, c�ļ�ID)) = 0 Then lngItemNO = 0
        End If
        
        Call AddActiveMenu(lngItemNO)
        
        If lngItemNO <> 0 Then
            strText = Trim(VsfData.TextMatrix(NewRow, NewCol))
            If strText = "" Then
                strPart = ""
            Else
                strKey = NewRow & "," & NewCol
                mrsCell.Filter = "ID='" & strKey & "'"
                strPart = ""
                If mrsCell.RecordCount > 0 Then
                    strPart = Trim(Nvl(mrsCell!��λ))
                End If
            End If
            
            If strPart = "" Then
                mrsPart.Filter = "��Ŀ���=" & lngItemNO & " and ȱʡ��=1"
                If mrsPart.RecordCount > 0 Then strPart = Trim(Nvl(mrsPart!��λ))
                If lngItemNO = gint���� And strPart = "" Then
                    strPart = "��������"
                End If
            End If
            
            '���ݲ�λ��Ϣѡ��λ�˵��Ĳ�λ
            For Each cbrControl In mcbrToolBar.Controls(4).CommandBar.Controls
                If Trim(cbrControl.Parameter) = Trim(strPart) Then
                    cbrControl.Checked = True
                Else
                    cbrControl.Checked = False
                End If
            Next
        Else
            'ȷ��������Һ����
            lngItemNO = Val(VsfData.TextMatrix(0, NewCol))
            strText = Trim(VsfData.TextMatrix(NewRow, NewCol))
            blnCheck = False
            For Each cbrControl In mcbrToolBar.Controls(5).CommandBar.Controls
                cbrControl.Checked = False
                If lngItemNO = gint��� Then
                    Select Case cbrControl.Id
                        Case conMenu_Edit_Append * 10 + 1
                            cbrControl.Checked = (InStr(1, UCase(strText), "/E") = 0 And InStr(1, UCase(strText), "E") > 0)
                        Case conMenu_Edit_Append * 10 + 2
                            cbrControl.Checked = (InStr(1, UCase(strText), "/E") > 0)
                        Case conMenu_Edit_Append * 10 + 3
                            cbrControl.Checked = (UCase(strText) = "*" Or UCase(strText) = "��")
                        Case conMenu_Edit_Append * 10 + 4
                            cbrControl.Checked = (UCase(strText) = "��")
                    End Select
                   
                ElseIf lngItemNO = gint��Һ Then
                    Select Case cbrControl.Id
                        Case conMenu_Edit_Append * 10 + 5
                            cbrControl.Checked = (InStr(1, UCase(strText), "/C") = 0 And InStr(1, UCase(strText), "C") > 0)
                        Case conMenu_Edit_Append * 10 + 6
                            cbrControl.Checked = InStr(1, UCase(strText), "/C") > 0
                    End Select
                End If
                If blnCheck = False Then blnCheck = cbrControl.Checked
            Next
            If blnCheck = False Then
                 mcbrToolBar.Controls(5).CommandBar.Controls(1).Checked = True
            End If
        End If
    End If

    mrsItems.Filter = 0
    
ErrInfo:
    RaiseEvent AfterRowColChange(strInfo, False)
    '��ȡ�ò�����ʷ����
    If OldRow <> NewRow Then
        Call RefreshHistoryData(NewRow)
    End If
End Sub

Private Sub VsfData_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    If Not mblnInit Then Exit Sub
    Call InitCons
End Sub

Private Sub VsfData_DblClick()
    Call VsfData_KeyDown(Asc("L"), 0)
End Sub

Private Sub VsfData_EnterCell()
    Dim lngItemNO As Long
    Dim strName As String
    Dim int����ȼ� As Integer
    Dim strKey As String, strInfo As String
    Dim rsObj As New ADODB.Recordset
    
    picInput.Visible = False
    lstNote.Visible = False
    picInput.Tag = ""
    lstNote.Tag = ""
    txtInput.Tag = ""
    picDouble.Visible = False
    picDouble.Tag = ""
    lstSelect(0).Visible = False
    lstSelect(1).Visible = False
    PicLst.Visible = False
    PicLst.Tag = ""
    txtLst.Visible = False
    cbo���±�ʶ.Visible = False
    
    mintType = 0
    
    VsfData.SetFocus
    
    If Not mblnInit Then Exit Sub
    If Not mblnShow Then Exit Sub
    If VsfData.Col < RootCol - 1 And VsfData.Col <> c���±�ʶ Then Exit Sub
    '����޲�����ϢҲ���ܱ༭
    If Val(VsfData.TextMatrix(VsfData.Row, c����ID)) = 0 Or Val(VsfData.TextMatrix(VsfData.Row, c�ļ�ID)) = 0 Then Exit Sub
    
    '��������Ѿ����棬���Ҹ��д���ͬ�����������ݡ��Ͳ������޸�ʱ��
    If VsfData.Col = cʱ�� Then
        mrsCopy.Filter = 0
        mrsCopy.Filter = "�к�=" & VsfData.Row
        Do While Not mrsCopy.EOF
            mint������Դ = Val(Nvl(mrsCopy!������Դ))
            If InStr(1, ",0,9,", "," & mint������Դ & ",") = 0 Then
                strInfo = "���������Ѿ����沢�Ұ���ͬ������������,�����޸�ʱ��."
                RaiseEvent AfterRowColChange(strInfo, True)
                Exit Sub
            End If
            mrsCopy.MoveNext
        Loop
    End If
    
    mint������Դ = 0
    mintModify = 0
    strName = VsfData.TextMatrix(VsfData.FixedRows - 1, VsfData.Col)
    lngItemNO = Val(VsfData.TextMatrix(0, VsfData.Col))
    int����ȼ� = Val(VsfData.TextMatrix(VsfData.Row, c����ȼ�))
    
    '��黤��ȼ������ò���
    If VsfData.Col >= RootCol Then
        mrsItems.Filter = "��Ŀ���=" & lngItemNO & " And ����ȼ�>=" & int����ȼ�
        If mrsItems.RecordCount = 0 Then
            strInfo = "��Ŀ[" & strName & "]�Ļ���ȼ������øò���."
            RaiseEvent AfterRowColChange(strInfo, True)
            Exit Sub
        End If
        
        '�Ƿ����ò���
        If Val(VsfData.TextMatrix(2, VsfData.Col)) = 1 Then
            If Val(VsfData.TextMatrix(VsfData.Row, cӤ��)) <> 0 Then
                strInfo = "��Ŀ[" & strName & "]ֻ�����ڲ���."
                RaiseEvent AfterRowColChange(strInfo, True)
                Exit Sub
            End If
        ElseIf VsfData.TextMatrix(2, VsfData.Col) = 2 Then
           If Val(VsfData.TextMatrix(VsfData.Row, cӤ��)) = 0 Then
                strInfo = "��Ŀ[" & strName & "]ֻ������Ӥ��."
                RaiseEvent AfterRowColChange(strInfo, True)
                Exit Sub
            End If
        End If
    End If
    
    '��������Ƿ���ͬ��������
    mrsCell.Filter = 0
    strKey = VsfData.Row & "," & VsfData.Col
    mrsCell.Filter = "ID='" & strKey & "'"
    If mrsCell.RecordCount > 0 Then '�����mrsCellΪ��
        Set rsObj = mrsCell.Clone
    Else
        Set rsObj = mrsCopy.Clone
    End If
    rsObj.Filter = "ID='" & strKey & "'"
    
    If rsObj.RecordCount > 0 Then
        lngItemNO = Val(Nvl(rsObj!��Ŀ���))
        mint������Դ = Val(Nvl(rsObj!������Դ))
        mintModify = Val(Nvl(rsObj!�޸�))
        If InStr(1, ",0,9,", "," & Val(rsObj!������Դ) & ",") = 0 Then
            If Not (lngItemNO = gint���� Or (lngItemNO = gint���� And mint����Ӧ�� = 2)) Then
                strInfo = "ͬ��������[" & strName & "]���ݲ��ܽ����޸�."
                RaiseEvent AfterRowColChange(strInfo, True)
                Exit Sub
            Else
                If mintModify = 1 Then
                    If lngItemNO = gint���� Then
                        strInfo = "ͬ��������[" & strName & "]����������������²��ܽ����޸�."
                    Else
                        strInfo = "ͬ��������[" & strName & "]������������������᲻�ܽ����޸�."
                    End If
                    RaiseEvent AfterRowColChange(strInfo, True)
                    Exit Sub
                End If
            End If
        End If
    End If
    
    Call ShowInput
End Sub

Private Sub VsfData_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       '������һ�л���һ��
       Call MoveNextCell
    Else
        If Not (KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or Shift <> 0) Then
            mblnShow = True
            Call VsfData_EnterCell
        End If
    End If
End Sub

Private Sub VsfData_GotFocus()
    picFilter.Visible = False
    picPati.Visible = False
End Sub

Private Sub ShowInput()
'----------------------------------
'������ʾ�������Ϣ
'----------------------------------
    Dim strText As String, strText1 As String, strPart As String
    Dim intCOl As Integer, intRow As Integer
    Dim CellRect As RECT
    Dim lngItemNO As Long
    Dim intType As Integer, intIndex As Integer
    Dim strLen As String
    Dim strTmp As String, strPoint As String
    Dim arrValue() As String, arrValue1() As String
    Dim blnSelect As Boolean
    Dim i As Integer, j As Integer
    
    Call InitCons
    intType = -1
    intCOl = VsfData.Col
    intRow = VsfData.Row
    
    CellRect.Left = VsfData.CellLeft + VsfData.Left
    CellRect.Top = VsfData.CellTop + VsfData.Top
    CellRect.Bottom = VsfData.CellHeight + 20
    CellRect.Right = VsfData.CellWidth + 20
    
    strPart = ""
    If intCOl = cʱ�� Then
        strText1 = Trim(VsfData.TextMatrix(intRow, intCOl))
        If strText1 = "" Then
            '����û��Ѿ�¼��ʱ����Ϣ���������ʱ���Դ�ʱ��Ϊ׼
            If Not IsDate(mstrModifyTime) Then
                strText = Format(zlDatabase.Currentdate, "HH:mm")
            Else
                strText = Format(mstrModifyTime, "HH:mm")
            End If
        Else
            strText = Format(strText1, "HH:mm")
        End If
        intType = -1
    ElseIf intCOl = c���±�ʶ Then
        Call zlControl.CboLocate(cbo���±�ʶ, VsfData.TextMatrix(intRow, intCOl))
        intType = -2
    Else
        strText = Trim(VsfData.TextMatrix(intRow, intCOl))
        If InStr(1, strText, ":") <> 0 Then
            strPart = Trim(Split(strText, ":")(0))
            strText = Trim(Split(strText, ":")(1))
        End If
        strText1 = strText
        lngItemNO = VsfData.TextMatrix(0, intCOl)
        intType = 0
    End If
    
    If intType = 0 Then
        If lngItemNO <> 4 Then
            mintType = 1
            mrsItems.Filter = "��Ŀ���=" & lngItemNO
            If InStr(1, ",2,3,5,", "," & Val(Nvl(mrsItems!��Ŀ��ʾ)) & ",") = 0 Then
                strLen = Nvl(mrsItems!��Ŀ����, 0) & ";" & Nvl(mrsItems!��ĿС��, 0)
                If lngItemNO = gint���� Or (lngItemNO = gint���� And mint����Ӧ�� = 2) Then
                    strLen = (Val(Split(strLen, ";")(0)) + Val(Split(strLen, ";")(0)) + 1) & ";" & IIf(Val(Split(strLen, ";")(1)) = 0, 0, 1) * 2
                End If
                
                If Val(strLen) <> 0 Then
                    txtInput.MaxLength = Val(Split(strLen, ";")(0)) + IIf(Val(Split(strLen, ";")(1)) = 0, 0, 1)
                Else
                    txtInput.MaxLength = 0
                End If
            Else
                mintType = Val(Nvl(mrsItems!��Ŀ��ʾ))
                strText1 = Nvl(mrsItems!��Ŀֵ��, ";")
            End If
        Else
            mintType = 4
            mrsItems.Filter = "��Ŀ���=4 or ��Ŀ���=5"
            mrsItems.Sort = "��Ŀ���"
            Do While Not mrsItems.EOF
                strTmp = Val(strTmp) + Val(Nvl(mrsItems!��Ŀ����))
                strPoint = Val(strPoint) + Val(Nvl(mrsItems!��ĿС��))
                strLen = strTmp & ";" & strPoint
                Select Case Val(mrsItems!��Ŀ���)
                    Case 4
                        If Val(strLen) <> 0 Then
                            txtUpInput.MaxLength = Val(Split(strLen, ";")(0)) + IIf(Val(Split(strLen, ";")(1)) = 0, 0, 1)
                        Else
                            txtUpInput.MaxLength = 0
                        End If
                    Case 5
                        If Val(strLen) <> 0 Then
                            txtDnInput.MaxLength = Val(Split(strLen, ";")(0)) + IIf(Val(Split(strLen, ";")(1)) = 0, 0, 1)
                        Else
                            txtDnInput.MaxLength = 0
                        End If
                End Select
            mrsItems.MoveNext
            Loop
        End If
    ElseIf intType = -1 Then
        mintType = 1
        txtInput.MaxLength = 5
    Else
        mintType = -2
    End If
    
    Select Case mintType
        Case -2 '���±�ʶ
            With cbo���±�ʶ
                .Left = CellRect.Left
                .Top = CellRect.Top
                .Width = CellRect.Right
                .FontName = VsfData.FontName
                .Font.Size = VsfData.Font.Size
                .Visible = True
                .ZOrder 0
            End With
        Case 1
            With picInput
                .Left = CellRect.Left
                .Top = CellRect.Top
                .Width = CellRect.Right
                .Height = CellRect.Bottom
                .FontName = VsfData.FontName
                .Font.Size = VsfData.Font.Size
                .Visible = True
                .ZOrder 0
            End With
            
            lblCheck.Visible = False
            
            With txtInput
                .Top = 0
                .Left = 0
                .Width = CellRect.Right
                .Height = CellRect.Bottom
                .FontName = VsfData.FontName
                .Font.Size = VsfData.Font.Size
                .Width = .Width - (180 + IIf(mintBigSize, 180 * 1 / 3, 0)) / 2 '����9��ʱ��ȥ90,����Խ��۳��ı߾�ԽС,�Ա�֤�ı��������ʵ��һ��
                .Visible = True
                .Text = strText
                .Tag = strPart  'intRow & "|" & intCOl
                .ZOrder 0
                picInput.Tag = strText1
            End With
            
            picInput.SetFocus
        Case 2, 3 '��ѡ���ѡ
            Select Case mintType
                Case 2
                    intIndex = 0
                    If Left(strText1, 1) <> ";" Then strText1 = ";" & strText1
                Case 3
                    intIndex = 1
            End Select
            
            strText = Trim(VsfData.TextMatrix(intRow, intCOl))
            arrValue = Split(strText1, ";") 'ֵ��
            lstSelect(intIndex).Clear
        
            PicLst.Tag = "1"
            For i = 0 To UBound(arrValue)
                If Left(arrValue(i), 1) = "��" Then arrValue(i) = Mid(arrValue(i), 2): strText1 = arrValue(i)
                lstSelect(intIndex).AddItem arrValue(i), i
                 
                If intIndex = 0 Then
                   ReDim arrValue1(0)
                   arrValue1(0) = strText
                   txtLst.Text = strText
                Else
                   arrValue1 = Split(strText, ",")
                End If
                For j = 0 To UBound(arrValue1)
                    If arrValue1(j) = arrValue(i) Then
                        lstSelect(intIndex).Selected(i) = True
                        blnSelect = True
                    End If
                Next j
            Next i
            If blnSelect = False And strText1 <> "" And IIf(intIndex = 0, Trim(txtLst.Text) = "", True) Then
                For i = 0 To lstSelect(intIndex).ListCount - 1
                    If lstSelect(intIndex).List(i) = strText1 Then
                        lstSelect(intIndex).Selected(i) = True
                    End If
                Next i
            End If
            If lstSelect(intIndex).ListIndex >= 0 Then txtLst.Text = "": PicLst.Tag = "0"
            
            '�ؼ���ʾ
            '51600,������,2012-07-16,��ѡ��Ŀ�ṩ����ѡ���¼�빦��
            If intIndex = 0 Then
                mrsItems.Filter = "��Ŀ���=" & lngItemNO
                If mrsItems.RecordCount > 0 Then
                    strLen = Nvl(mrsItems!��Ŀ����, 0) & ";" & Nvl(mrsItems!��ĿС��, 0)
                End If
                With PicLst
                    .FontName = VsfData.FontName
                    .FontSize = VsfData.FontSize
                    .Left = CellRect.Left
                    .Top = CellRect.Top
                    .Height = 80 + CellRect.Bottom + PicLst.TextHeight("��") * 2 + lstSelect(intIndex).ListCount * (PicLst.TextHeight("��") + PicLst.TextHeight("��") \ 4)
                    If .Height < CellRect.Bottom Then .Height = CellRect.Bottom
                    .Width = LenB(StrConv(lstSelect(intIndex).List(lstSelect(intIndex).ListCount \ 2), vbFromUnicode)) * 100 + 500    '���м���ĳ���Ϊ����
                    If .Width < CellRect.Right Then .Width = CellRect.Right
                    If .Height > VsfData.Height Then
                        .Height = VsfData.Height
                    End If
                    If .Top + .Height > VsfData.Height Then
                        .Top = CellRect.Top + CellRect.Bottom - .Height
                    End If
                    If .Top < 0 Then .Top = VsfData.Top
                
                    PicLst.Visible = True
                    PicLst.ZOrder 0
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
                    .Height = CellRect.Bottom
                    .FontName = VsfData.FontName
                    .FontSize = VsfData.FontSize
                    .MaxLength = Val(Split(strLen, ";")(0)) + IIf(Val(Split(strLen, ";")(1)) = 0, 0, 1)
                    .Visible = True
                End With
                
                With lbllst(1)
                    .Left = 20
                    .Top = txtLst.Top + txtLst.Height + 20
                    .FontName = VsfData.FontName
                    .FontSize = VsfData.FontSize
                    .Visible = True
                End With
                
                With lstSelect(intIndex)
                    .Top = lbllst(1).Top + lbllst(1).Height + 20
                    .Left = -10
                    .FontName = VsfData.FontName
                    .FontSize = VsfData.FontSize
                    .Width = PicLst.Width
                    .Height = PicLst.Height - .Top
                    .Visible = True
                    .Enabled = True
                    .ZOrder 0
                    .Tag = strText
                End With
                If lstSelect(intIndex).Top + lstSelect(intIndex).Height <> PicLst.Height Then
                    PicLst.Height = lstSelect(intIndex).Top + lstSelect(intIndex).Height
                End If
                PicLst.SetFocus
            Else
                lstSelect(intIndex).Top = CellRect.Top
                lstSelect(intIndex).Left = CellRect.Left
                lstSelect(intIndex).FontName = VsfData.FontName
                lstSelect(intIndex).FontSize = VsfData.FontSize
                lstSelect(intIndex).Height = lstSelect(intIndex).ListCount * (PicLst.TextHeight("��") + PicLst.TextHeight("��") \ 4)
                If lstSelect(intIndex).Height < CellRect.Bottom Then lstSelect(intIndex).Height = CellRect.Bottom
                lstSelect(intIndex).Width = LenB(StrConv(lstSelect(intIndex).List(lstSelect(intIndex).ListCount \ 2), vbFromUnicode)) * 100 + 500    '���м���ĳ���Ϊ����
                If lstSelect(intIndex).Width < CellRect.Right Then lstSelect(intIndex).Width = CellRect.Right
                If lstSelect(intIndex).Height > VsfData.Height Then
                    lstSelect(intIndex).Height = VsfData.Height
                End If
                If lstSelect(intIndex).Top + lstSelect(intIndex).Height > VsfData.Height Then
                    lstSelect(intIndex).Top = CellRect.Top + CellRect.Bottom - lstSelect(intIndex).Height
                End If
                If lstSelect(intIndex).Top < 0 Then lstSelect(intIndex).Top = VsfData.Top
                
                lstSelect(intIndex).Visible = True
                lstSelect(intIndex).Enabled = True
                lstSelect(intIndex).ZOrder 0
                
                lstSelect(intIndex).Tag = strText
                lstSelect(intIndex).SetFocus
            End If
        Case 4 'Ѫѹ
            With picDouble
                .Left = CellRect.Left
                .Top = CellRect.Top
                .Width = CellRect.Right
                .Height = CellRect.Bottom
                .FontName = VsfData.FontName
                .Font.Size = VsfData.Font.Size
                .Visible = True
                .Tag = strText1
                .ZOrder 0
            End With
            
            If strText = "" Then strText = "/"
            arrValue = Split(strText, "/")
            
            lblSplit.FontName = VsfData.FontName
            lblSplit.FontSize = VsfData.FontSize
            lblSplit.Left = (picDouble.Width - lblSplit.Width) / 2
            If mintBigSize = 1 Then
                lblSplit.Width = 150
            Else
                lblSplit.Width = 105
            End If
    
            With txtUpInput
                .Text = arrValue(0)
                .FontName = VsfData.FontName
                .Font.Size = VsfData.Font.Size
                .Width = (picDouble.Width - lblSplit.Width) * 0.4
                .ZOrder 0
            End With
            
            With txtDnInput
                .Text = arrValue(1)
                .FontName = VsfData.FontName
                .Font.Size = VsfData.Font.Size
                .Left = lblSplit.Left + lblSplit.Width
                .Width = picDouble.Width - .Left
                .ZOrder 0
            End With
            
            picDouble.SetFocus
        Case 5 'ѡ��
            With picInput
                .Left = CellRect.Left
                .Top = CellRect.Top
                .Width = CellRect.Right
                .Height = CellRect.Bottom
                .FontName = VsfData.FontName
                .Font.Size = VsfData.Font.Size
                .Visible = True
                .ZOrder 0
            End With
            
            txtInput.Visible = False
            
            With lblCheck
                .Top = 0
                .Left = 0
                .Caption = strText
                .Width = CellRect.Right
                .Height = CellRect.Bottom
                .FontName = VsfData.FontName
                .Font.Size = VsfData.Font.Size
                .Width = .Width - (180 + IIf(mintBigSize, 180 * 1 / 3, 0)) / 2 '����9��ʱ��ȥ90,����Խ��۳��ı߾�ԽС,�Ա�֤�ı��������ʵ��һ��
                .Visible = True
                .ZOrder 0
                picInput.Tag = strText1
            End With
            
            picInput.SetFocus
    End Select
End Sub

Private Sub MoveNextCell(Optional KeyCode As Integer = vbKeyReturn)
'--------------------------------------------
'����:������ݲ���ֵ ���ƶ�����һ�л���һ��
'--------------------------------------------
    Dim lngItemNO As Integer, i As Integer, intIndex As Integer
    Dim strText As String, strErrMsg As String, strPatiTime As String, strOldValue As String
    Dim intCOl As Integer, intRow As Integer
    Dim blnValidate As Boolean, blnSave As Boolean
    Dim strFileds As String, strValues As String, strKey As String, strPart As String
    Dim int��ԴID As Long, int���� As Integer, int��ʾ As Integer, int�޸� As Integer
    Dim intState As Integer
    Dim strData As String
   
    'If picInput.Visible = False Then Exit Sub
    
    If mblnInit = False Then Exit Sub
    
    strFileds = "ID|�к�|��Ŀ���|����|��λ|������Դ|��ԴID|����|��ʾ|�޸�|״̬"
    intCOl = VsfData.Col
    intRow = VsfData.Row
    blnValidate = False
    blnSave = False
    strOldValue = ""
    If KeyCode = vbKeyReturn And InStr(1, ",0,-2,", "," & mintType & ",") = 0 Then ' (picInput.Visible = True Or picDouble.Visible = True) Then
        mlng�ļ�ID = Val(VsfData.TextMatrix(intRow, c�ļ�ID))
        mlng����ID = Val(VsfData.TextMatrix(intRow, c����ID))
        mlng��ҳID = Val(VsfData.TextMatrix(intRow, c��ҳID))
        mlngBaby = Val(VsfData.TextMatrix(intRow, cӤ��))
        strPatiTime = VsfData.TextMatrix(intRow, c����)
        mbln��Ժ = (Val(VsfData.TextMatrix(intRow, c��Ժ)) = 1)
        
        Select Case mintType
            Case 1
                strText = Trim(txtInput.Text)
'                If InStr(txtInput.Text, "/") > 0 And mbln����������ʾ Then
'                    strData = Split(Trim(txtInput.Text), "/")(1) & "/" & Split(Trim(txtInput.Text), "/")(0)
'                    strText = strData
'                End If
                
                strPart = Trim(txtInput.Tag)
                strOldValue = picInput.Tag
            Case 2, 3
                If mintType = 2 Then
                    intIndex = 0
                Else
                    intIndex = 1
                End If
                strText = ""
                strPart = ""
                For i = 0 To lstSelect(intIndex).ListCount - 1
                  If lstSelect(intIndex).Selected(i) = True Then
                      strText = strText & "," & Replace(lstSelect(intIndex).List(i), ",", "")
                  End If
                Next i
                If Left(strText, 1) = "," Then strText = Mid(strText, 2)
                '51600�������ɣ�2012-07-16����ѡ����¼��Ҳ��������
                If intIndex = 0 And Val(PicLst.Tag) = 1 Then strText = Trim(txtLst.Text)
                strOldValue = lstSelect(intIndex).Tag
            Case 4
                strText = Trim(txtUpInput.Text) & "/" & Trim(txtDnInput.Text)
                strPart = ""
                If strText = "/" Then strText = ""
                strOldValue = picDouble.Tag
            Case 5
                strText = lblCheck.Caption
                strPart = ""
                strOldValue = picInput.Tag
        End Select
        
        '���ʱ��������Ƿ�Ϸ�
        If intCOl = cʱ�� Then
            If Not CheckDateTime(strText, strPatiTime, strErrMsg) Then picInput.SetFocus: GoTo ErrInfo
            '�˴����»�ȡ�к�,��Ϊ���ڱ���������޸�ʱ����ɾ��ԭ��ʱ�����ݣ�������һ����ʱ���µ�����(���ظ��У�����һ��������)
            intRow = VsfData.Row
            mstrModifyTime = Format(strText, "HH:mm")
            blnValidate = True
        ElseIf intCOl > cʱ�� Then
            lngItemNO = Val(VsfData.TextMatrix(0, intCOl))
            If Not CheckValid(strText, lngItemNO, strErrMsg) Then
                Select Case mintType
                    Case 1
                        picInput.SetFocus
                    Case 2, 3
                        If mintType = 2 Then
                            intIndex = 0
                        Else
                            intIndex = 1
                        End If
                        lstSelect(intIndex).SetFocus
                    Case 4
                        picDouble.SetFocus
                    Case Else
                        picInput.SetFocus
                End Select
                GoTo ErrInfo
            End If
            blnValidate = True
            If mlng����ID = 0 Or mlng�ļ�ID = 0 Or mlng��ҳID = 0 Then
                blnSave = False
            Else
                blnSave = True
            End If
        End If
        
        If blnValidate = True Then
            mrsCopy.Filter = 0
            VsfData.TextMatrix(intRow, intCOl) = IIf(strPart = "", "", strPart & ":") & strText
            VsfData.Cell(flexcpAlignment, intRow, intCOl, intRow, intCOl) = flexAlignCenterCenter
            '�������ݴ���
            If blnSave = True Then
                    strKey = intRow & "," & intCOl
                    '����޸ĵ������Ƿ��Ѿ�����
                    mrsCopy.Filter = "ID='" & strKey & "'"
                    If mrsCopy.RecordCount > 0 Then
                        int��ԴID = Val(Nvl(mrsCopy!��ԴID))
                        int���� = Val(Nvl(mrsCopy!����))
                        int��ʾ = Val(Nvl(mrsCopy!��ʾ))
                        int�޸� = Val(Nvl(mrsCopy!�޸�))
                        strPatiTime = Nvl(mrsCopy!��λ)
                        intState = 1
                    Else
                        int��ԴID = 0: int���� = 0: int��ʾ = 0: int�޸� = 0: strPatiTime = ""
                        intState = IIf(Trim(strText) = "", 3, 1)
                        mrsCell.Filter = "ID='" & strKey & "'"
                        If mrsCell.RecordCount > 0 Then
                            int��ԴID = Val(Nvl(mrsCell!��ԴID))
                            int���� = Val(Nvl(mrsCell!����))
                            int��ʾ = Val(Nvl(mrsCell!��ʾ))
                            int�޸� = Val(Nvl(mrsCell!�޸�))
                            strPatiTime = Nvl(mrsCell!��λ)
                        End If
                    End If
                    If Trim(strOldValue) <> Trim(strText) Or (Trim(strText) <> "" And strPart <> strPatiTime) Then
                        strValues = strKey & "|" & intRow & "|" & lngItemNO & "|" & strText & "|" & strPart & "|" & mint������Դ & "|" & _
                            int��ԴID & "|" & int���� & "|" & int��ʾ & "|" & int�޸� & "|" & intState
                        Call Record_Update(mrsCell, strFileds, strValues, "ID|" & strKey)
                    End If
                    mblnChage = True
            End If
        End If
    End If
    
    mintType = 0
    '��ʼ�ƶ��л���
    With VsfData
        If KeyCode = vbKeyReturn Then
          
NextCol2: '������һ��
            If .Col < .FixedCols Then
                .Col = .Col + 1: GoTo NextCol2
            End If
            If .Col < .Cols - 1 Then
                .Col = .Col + 1
                If .ColHidden(.Col) = True Then GoTo NextCol2
            Else
NextRow2: '������һ��
                If .Row < .Rows - 1 Then
                    intRow = .Row + 1
                    If .RowHidden(intRow) = True Then GoTo NextRow2
                    intCOl = cʱ��
                    .Select intRow, intCOl
                Else
                    intRow = .Row
                    intCOl = cʱ��
                    .Select intRow, intCOl
                End If
            End If
            '������л��в��ɼ����Զ���ʾ����
            If .ColIsVisible(.Col) = False Then
                .LeftCol = .Col
            End If
            If .RowIsVisible(.Row) = False Then
                .TopRow = .Row
            End If

            Exit Sub
        End If
        '���
        If KeyCode = vbKeyLeft Then
PreCol2:
            If .Col > cʱ�� Then
                .Col = .Col - 1
                If .ColHidden(.Col) = True Then GoTo PreCol2
            Else
PreRow2:
                If .Row > .FixedRows Then
                    intRow = .Row - 1
                    If .RowHidden(intRow) Then GoTo PreRow2
                    intCOl = .Cols - 1
                    .Select intRow, intCOl
                End If
            End If
            '������л��в��ɼ����Զ���ʾ����
            If .ColIsVisible(.Col) = False Then
                .LeftCol = .Col
            End If
            If .RowIsVisible(.Row) = False Then
                .TopRow = .Row
            End If
            Exit Sub
        End If
    End With
    
    Exit Sub
ErrInfo:
    RaiseEvent AfterRowColChange(strErrMsg, True)
End Sub

Private Function SaveDate() As Boolean
'------------------------------------------
'����:����������Ϣ
'------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim blnTrans As Boolean, blnTimeNull As Boolean
    Dim lngRow As Long, lngCol As Long, lngItemCode As Long, lngRecordID As Long
    Dim strKey As String, strPart As String, strValue As String
    Dim strTime As String, strEnd As String, strMarkTime As String, strSQL As String
    Dim arrSQL() As String, arrData() As String
    Dim i As Integer, intRow As Integer
    Dim strValues As String, strNote As String, strSaveRows As String
    Dim blnData As Boolean, blnSave As Boolean
    '���������Ϣ
    Dim lng�ļ�ID As Long, lng����ID As Long, lng��ҳID As Long, lngӤ�� As Long
    On Error GoTo Errhand
    
    mrsCell.Filter = 0
    '��������ݵ����Ƿ���дʱ��
    For lngRow = VsfData.FixedRows To VsfData.Rows - 1
        If Val(VsfData.TextMatrix(lngRow, c�ļ�ID)) <> 0 And VsfData.RowHidden(lngRow) = False Then
            blnTimeNull = IIf(Trim(VsfData.TextMatrix(lngRow, cʱ��)) = "", True, False)
            If blnTimeNull = True Then
                mrsCell.Filter = "�к�=" & lngRow & " And ״̬=1"
                If mrsCell.RecordCount > 0 Then
                    mblnShow = True
                    VsfData.Select lngRow, cʱ��
                    RaiseEvent AfterRowColChange("ʱ�䲻��Ϊ��,��¼��ʱ��.", True)
                    Exit Function
                End If
            End If
        End If
    Next lngRow
    
    Screen.MousePointer = 11
          
    ReDim Preserve arrSQL(1 To 1)
    
    strSaveRows = ""
    '���ȼ��ʱ���Ƿ����
    mrsData.Filter = 0
    For lngRow = VsfData.FixedRows To VsfData.Rows - 1
        If VsfData.RowHidden(lngRow) = False Then
            mrsData.Filter = "�к�=" & lngRow
            If mrsData.RecordCount > 0 Then
                lngRecordID = Val(Nvl(mrsData!��¼ID))
                If Val(Nvl(mrsData!ɾ��)) = 2 And lngRecordID > 0 Then '��ʾʱ���޸�
                    strTime = Format(Format(mstrDate, "YYYY-MM-DD") & " " & VsfData.TextMatrix(lngRow, cʱ��), "YYYY-MM-DD HH:mm:ss")
                    strMarkTime = strTime
                    strMarkTime = "To_Date('" & strMarkTime & "','yyyy-mm-dd hh24:mi:ss')"
                    strSQL = "ZL_���µ�����_����ʱ��("
                    'ID_IN       IN ���˻�������.ID%TYPE,
                    strSQL = strSQL & lngRecordID & ","
                    '����ʱ��_IN IN ���˻�������.����ʱ��%TYPE
                    strSQL = strSQL & strMarkTime & ")"
                    arrSQL(ReDimArray(arrSQL)) = strSQL
                    
                    strSaveRows = strSaveRows & "," & lngRow
                End If
            End If
        End If
    Next lngRow
    
    If Left(strSaveRows, 1) = "," Then strSaveRows = Mid(strSaveRows, 2)
    
    intRow = 0
    blnSave = False
    '���ݼ��ɹ���ʼ��ȡ��¼��
    mrsCell.Filter = 0
    mrsCell.Sort = "�к�"
    With mrsCell
        Do While Not .EOF
            If Val(Nvl(mrsCell!״̬)) = 1 Then
                If intRow <> Val(!�к�) Then
ErrRow:
                    If blnSave = True Then
                        If InStr(1, "," & strSaveRows & ",", "," & lngRow & ",") = 0 Then
                            strSaveRows = strSaveRows & "," & lngRow
                        End If
                        intRow = lngRow
                        blnSave = False
                        If .EOF Then Exit Do
                    End If
                End If
                
                strKey = !Id
                lngRow = Val(Split(strKey, ",")(0))
                lngCol = Val(Split(strKey, ",")(1))
                
                strTime = Format(Format(mstrDate, "YYYY-MM-DD") & " " & VsfData.TextMatrix(lngRow, cʱ��), "YYYY-MM-DD HH:mm:ss")
                strEnd = strTime
                strMarkTime = strTime
                
                lngItemCode = Val(!��Ŀ���)
                strPart = Nvl(!��λ)
                strValue = Nvl(!����)
                strNote = ""
                
            
                
                lng�ļ�ID = Val(VsfData.TextMatrix(lngRow, c�ļ�ID))
                lng����ID = Val(VsfData.TextMatrix(lngRow, c����ID))
                lng��ҳID = Val(VsfData.TextMatrix(lngRow, c��ҳID))
                lngӤ�� = Val(VsfData.TextMatrix(lngRow, cӤ��))
                
                strMarkTime = "To_Date('" & strMarkTime & "','yyyy-mm-dd hh24:mi:ss')"

                mrsItems.Filter = 0
                mrsItems.Filter = "��Ŀ���=" & lngItemCode
                If mrsItems!������ = "1)����������Ŀ" Then
                    '--��¼����
                    If strValue = "����" And lngItemCode = gint���� Then
                        strNote = ""
                    Else
                        If IsNumeric(strValue) Or InStr(1, strValue, "/") > 0 Then
                             strNote = ""
                        Else
                            strNote = strValue
                            strValue = ""
                        End If
                    End If
                Else
                     strNote = ""
                End If
                
                '�����:56853,�޸���:����,����������ʾ��ʽ������/����
                If lngItemCode = gint���� And mint����Ӧ�� = 2 Then
                    If mbln����������ʾ And strValue <> "" Then
                        strValue = Split(Nvl(!����), "/")(1) & "/" & Split(Nvl(!����), "/")(0)
                    End If
                End If
                
                If lngItemCode = 4 And (Nvl(!����) = "δ��" Or Nvl(!����) = "�ܲ�" Or Nvl(!����) = "���" Or Nvl(!����) = "���") Then
                    strValue = Nvl(!����) & "/" & Nvl(!����)
                End If
                    
                '����������Ϣ
                strSQL = "Zl_���µ�����_Update("
                '�ļ�id_In   In ���˻����ļ�.Id%Type,  --���˻����ļ�ID
                strSQL = strSQL & Val(lng�ļ�ID) & ","
                '����ʱ��_In In ���˻�������.����ʱ��%Type, --�������ݵķ���ʱ��
                strSQL = strSQL & strMarkTime & ","
                '��¼����_In In ���˻�����ϸ.��¼����%Type, --������Ŀ=1���ϱ�˵��=2�����ת���=3�������ձ��=4,�±�˵��=6
                strSQL = strSQL & "1,"
                '��Ŀ���_In In ���˻�����ϸ.��Ŀ���%Type, --������Ŀ����ţ��ǻ�����Ŀ�̶�Ϊ0
                strSQL = strSQL & lngItemCode & ","
                '��¼����_In In ���˻�����ϸ.��¼����%Type := Null, --��¼���ݣ��������Ϊ�գ��������ǰ������  36��36/37
                strSQL = strSQL & "'" & strValue & "',"
                '���²�λ_In In ���˻�����ϸ.���²�λ%Type := Null, --ɾ������ʱ������д��λ �����Ŀ��
                strSQL = strSQL & IIf(strValue <> "", "'" & strPart & "'", "NULL") & ","
                '���Ժϸ�_In In Number := 0,
                strSQL = strSQL & "NULL,"
                'δ��˵��_In In ���˻�����ϸ.δ��˵��%Type := Null, --δ��˵��
                strSQL = strSQL & "'" & strNote & "',"
                '���˼�¼_In In Number := 1,
                strSQL = strSQL & "1,"
                '������Դ_In In ���˻�����ϸ.������Դ%Type := 0,
                strSQL = strSQL & "0,"
                '��Դid_In   In ���˻�����ϸ.��Դid%Type := Null,
                strSQL = strSQL & IIf(Val(Nvl(!��ԴID)) = 0, "NULL", Val(Nvl(!��ԴID))) & ","
                '����_In     In ���˻�����ϸ.����%Type := 0,
                strSQL = strSQL & Val(Nvl(!����))
                strSQL = strSQL & ")"

                arrSQL(ReDimArray(arrSQL)) = strSQL
                
                If intRow <> Val(!�к�) Then blnSave = True
            
            End If
        .MoveNext
        Loop
        If blnSave = True Then GoTo ErrRow
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    'ѭ��ִ��SQL��������
    gcnOracle.BeginTrans
    blnTrans = True
    
    blnData = False
    
    'Debug.Print "---���濪ʼ:" & Now
    For i = 1 To UBound(arrSQL)
        If arrSQL(i) <> "" Then Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "������������"): blnData = True: 'Debug.Print CStr(arrSQL(i))
    Next
    gcnOracle.CommitTrans
    
    'Debug.Print "---�������:" & Now
    blnTrans = False
    
    Screen.MousePointer = 0
    mblnChage = False
    mblnShow = False
    mblnSaveData = True
    
    Call InitCons
    
    If Left(strSaveRows, 1) = "," Then strSaveRows = Mid(strSaveRows, 2)
    '���¼�¼ID
    For lngRow = VsfData.FixedRows To VsfData.Rows - 1
        blnTimeNull = IIf(Trim(VsfData.TextMatrix(lngRow, cʱ��)) = "", True, False)
        If Not blnTimeNull And VsfData.RowHidden(lngRow) = False Then
            If InStr(1, "," & strSaveRows & ",", "," & lngRow & ",") <> 0 Then
                strTime = Format(Format(mstrDate, "YYYY-MM-DD") & " " & VsfData.TextMatrix(lngRow, cʱ��), "YYYY-MM-DD HH:mm:ss")
                strSQL = " Select A.ID From ���˻������� A,���˻����ļ� B" & vbNewLine & _
                              " Where A.�ļ�ID=B.ID And B.ID=[1] And A.����ʱ��=[2]"
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��¼ID", Val(VsfData.TextMatrix(lngRow, c�ļ�ID)), CDate(strTime))
                If rsTemp.RecordCount <> 0 Then
                    VsfData.TextMatrix(lngRow, c��¼ID) = Val(Nvl(rsTemp!Id))
                End If
            End If
        End If
    Next lngRow
    
    SaveDate = True
    
    If blnData = True Then
        '�������ݼ�
        Call CopyCellData
        Call Data_Save
    End If
    Exit Function
Errhand:
    Screen.MousePointer = 0
    If blnTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CopyCellData() As Boolean
'------------------------------------------------
'����:Copy����������
'------------------------------------------------
    Dim i As Integer
    
    'ɾ��״̬=3�����ݻ���ֵΪ�յ�����
    mrsCell.Filter = 0
    mrsCell.Filter = "״̬=3 or ����=''"
    Do While Not mrsCell.EOF
        mrsCell.Delete
        mrsCell.Update
        mrsCell.MoveNext
    Loop
    '�޸�״̬Ϊ0
    mrsCell.Filter = 0
    Do While Not mrsCell.EOF
        mrsCell!״̬ = 0
        mrsCell.Update
        mrsCell.MoveNext
    Loop
    
    'mrsCopy���е� mrscellû�и�ֵ��mrscell
    mrsCopy.Filter = 0
    Do While Not mrsCopy.EOF
        mrsCell.Filter = "ID='" & Nvl(mrsCopy!Id) & "'"
        If mrsCell.RecordCount = 0 Then
            mrsCell.AddNew
            For i = 0 To mrsCopy.Fields.Count - 1
                'ĿǰMrsCell��¼��ֻ���� adLongVarChar �� adDouble ��������
                If mrsCopy.Fields(i).Type = adLongVarChar Then
                    mrsCell.Fields(mrsCopy.Fields(i).Name).Value = Nvl(mrsCopy.Fields(i).Value)
                Else
                    mrsCell.Fields(mrsCopy.Fields(i).Name).Value = Val(Nvl(mrsCopy.Fields(i).Value))
                End If
            Next i
            mrsCell.Update
        End If
    mrsCopy.MoveNext
    Loop
    
    'ɾ����¼����Ϣ
    mrsCopy.Filter = 0
    Do While Not mrsCopy.EOF
        mrsCopy.Delete
        mrsCopy.Update
        mrsCopy.MoveNext
    Loop
    
    '��ʼ��������
    mrsCell.Filter = 0
    mrsCell.Sort = "�к�,ID"
    Do While Not mrsCell.EOF
        mrsCopy.AddNew
        For i = 0 To mrsCell.Fields.Count - 1
            'ĿǰMrsCell��¼��ֻ���� adLongVarChar �� adDouble ��������
            If mrsCell.Fields(i).Type = adLongVarChar Then
                mrsCopy.Fields(mrsCell.Fields(i).Name).Value = Nvl(mrsCell.Fields(i).Value)
            Else
                mrsCopy.Fields(mrsCell.Fields(i).Name).Value = Val(Nvl(mrsCell.Fields(i).Value))
            End If
        Next i
        mrsCopy.Update
    mrsCell.MoveNext
    Loop

    'ɾ����¼����Ϣ
    mrsCell.Filter = 0
    Do While Not mrsCell.EOF
        mrsCell.Delete
        mrsCell.Update
        mrsCell.MoveNext
    Loop
End Function

Private Function Data_Save() As Boolean
'-------------------------------------------------------
'����:�������ݱ���ĺ������Ϣ,��ˢ�º������Ϣ,һ��˧��
'------------------------------------------------------
    Dim lngRows As Long, lngStartRow As Long, lngCol As Long, lngCOls As Long
    On Error GoTo Errhand
    
    If mrsData Is Nothing Then Exit Function
    '����ڴ漯
    mrsData.Filter = 0
    Do While Not mrsData.EOF
        mrsData.Delete
        mrsData.Update
        mrsData.MoveNext
    Loop
    
    lngRows = VsfData.Rows - 1
    lngCOls = VsfData.Cols - 1
    
    '��ʼ����������
    For lngStartRow = VsfData.FixedRows To lngRows
        mrsData.AddNew
        mrsData("�к�") = lngStartRow
        For lngCol = c�ļ�ID To lngCOls
            mrsData.Fields(lngCol).Value = Trim(VsfData.TextMatrix(lngStartRow, lngCol))
        Next lngCol
        mrsData("ɾ��") = IIf(VsfData.RowHidden(lngStartRow), 1, 0)
        mrsData.Update
    Next lngStartRow
    
    If mblnNullRow Then Call RefreshHistoryData(VsfData.Row)
    
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function CheckDateTime(strReturn As String, ByVal strPatientTime As String, strInfo As String) As Boolean
'-----------------------------------------------------------------------------
'����:���¼���ʱ���Ƿ�Ϸ�
'strPatientTime ���µ���ʼʱ��;���˳�Ժʱ��
'-----------------------------------------------------------------------------
    Dim strText As String, strTime As String
    
    strText = Trim(strReturn)
    
    If Trim(strText) = "" Then
        strInfo = "ʱ�䲻��Ϊ�գ�"
        Exit Function
    End If
    If Len(strText) <= 2 Then
        strText = String(2 - Len(strText), "0") & strText
        strText = strText & ":00"
    End If
    If Val(Mid(strText, 1, 2)) < 0 Or Val(Mid(strText, 1, 2)) > 23 Then
        strInfo = "¼���ʱ����Ч��СʱӦ����0-23֮�䣡"
        Exit Function
    End If
    If Mid(strText, 3, 1) <> ":" Then
        strInfo = "¼���ʱ���ʽ����[04:00]��"
        Exit Function
    End If
    If Len(strText) < 5 Then strText = strText & String(5 - Len(strText), "0")
    If Not (Val(Mid(strText, 4, 2)) >= 0 And Val(Mid(strText, 4, 2)) <= 59) Then
        strInfo = "¼���ʱ����Ч������Ӧ����0-59֮�䣡"
        Exit Function
    End If
    If Len(strText) > 5 Then
        strInfo = "¼���ʱ���ʽ����[04:00]��"
        Exit Function
    End If
    
    If Trim(strText) <> Trim(picInput.Tag) Then
        strTime = Format(Format(mstrDate, "YYYY-MM-DD") & " " & strText, "YYYY-MM-DD HH:mm:ss")
        '���¼�����ݵ�ʱ���Ƿ񳬹����µ���ʼʱ������ݲ�¼ʱ��
        If Not CheckTime(strTime, strPatientTime, strInfo) Then Exit Function
        '���ݼ��ɹ������ʱ���Ƿ����ͬ��������������Ϣ
        If Not CheckPaseDate(strTime) Then Exit Function
    End If
    
    strReturn = strText
    CheckDateTime = True
End Function

Private Function CheckPaseDate(ByVal strTime As String) As Boolean
'------------------------------------------------------------------
'���õ��Ƿ����������Ϣ
'------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim i As Integer
    Dim lngItemNO As Long, lngRow As Long, lngCol As Long
    Dim strContent As String
    Dim strValues As String, strFileds As String
    Dim arrValues() As Variant, arrKeys() As Variant, arrID() As String
    Dim strKey As String, strKeys As String
    Dim intCOl As Integer
    Dim blnUpdate As Boolean
    Dim strPart As String, lng��ԴID As Long, int���� As Integer, int��ʾ As Integer, intModify As Integer, strNote As String
    Dim bln���� As Boolean, blnAllow As Boolean, bln���ʵ��� As Boolean
    Dim intState As Integer
    
    On Error GoTo Errhand:
    
    arrValues = Array()
    arrKeys = Array()
    
    bln���ʵ��� = True
    bln���� = False
    mrsItems.Filter = 0
    mrsItems.Filter = "��Ŀ���=" & gint����
    If mrsItems.RecordCount > 0 Then bln���� = True
    
    If mrsCell Is Nothing Then Exit Function
    strFileds = "ID|�к�|��Ŀ���|����|��λ|������Դ|��ԴID|����|��ʾ|�޸�|״̬"
    lngRow = VsfData.Row
    
    VsfData.Cell(flexcpForeColor, lngRow, cʱ��, lngRow, VsfData.Cols - 1) = &H80000012
    
    blnUpdate = False
    '�޸�ʱ���� ����Ƿ��Ǳ��������
    mrsCopy.Filter = 0
    mrsData.Filter = 0
    mrsCopy.Filter = "�к�=" & lngRow
    If mrsCopy.RecordCount > 0 Then
        mrsData.Filter = "�к�=" & lngRow
        If Format(strTime, "HH:mm") <> Format(mrsData.Fields(cʱ��).Value, "HH:mm") Then
            '�޸�mrsdata��¼��ɾ��=2 ��ʾʱ�䷢���޸�
            intState = Val(Nvl(mrsData!ɾ��))
            If intState = 1 Then
                intState = 1
            Else
                intState = 2
            End If
            mrsData!ɾ�� = intState
            mrsData.Update
            mblnChage = True
'            '��ȡ��һ��������Ϣ
'            For lngCol = RootCol To VsfData.Cols - 1
'                blnUpdate = False
'                strKey = lngRow & "," & lngCol
'
'                mrsCell.Filter = 0
'                mrsCell.Filter = "ID='" & strKey & "'"
'                If Not mrsCell.EOF Then
'                    strValues = Nvl(mrsCell!��Ŀ���) & "|" & VsfData.TextMatrix(lngRow, lngCol) & "|" & Nvl(mrsCell!��λ) & "|" & _
'                        Nvl(mrsCell!������Դ) & "|" & Nvl(mrsCell!��ԴID) & "|" & Nvl(mrsCell!����) & "|" & Nvl(mrsCell!��ʾ) & "|" & _
'                        Nvl(mrsCell!�޸�) & "|" & 1
'                    blnUpdate = True
'                Else
'                    mrsCopy.Filter = "ID='" & strKey & "'"
'                    If Not mrsCopy.EOF Then
'                        strValues = Nvl(mrsCopy!��Ŀ���) & "|" & VsfData.TextMatrix(lngRow, lngCol) & "|" & Nvl(mrsCopy!��λ) & "|" & _
'                            Nvl(mrsCopy!������Դ) & "|" & Nvl(mrsCopy!��ԴID) & "|" & Nvl(mrsCopy!����) & "|" & Nvl(mrsCopy!��ʾ) & "|" & _
'                            Nvl(mrsCopy!�޸�) & "|" & 1
'                         blnUpdate = True
'                    End If
'                End If
'
'                If blnUpdate = True Then
'                    ReDim Preserve arrValues(UBound(arrValues) + 1)
'                    arrValues(UBound(arrValues)) = strValues
'                    ReDim Preserve arrKeys(UBound(arrKeys) + 1)
'                    arrKeys(UBound(arrKeys)) = lngRow + 1 & "," & lngCol
'                End If
'            Next lngCol
'            'ʱ�䷢���ı� Ϊ��¼��mrscell����ɾ����ǲ�����һ���µ�����
'            Call Edit_Clear
'            VsfData.Row = lngRow + 1
        End If
    End If
    'lngRow = VsfData.Row
    
    '��ʼ���лָ�����(�Ա���������޸�ʱ��)
'    For i = 0 To UBound(arrValues)
'        lngCol = Val(Split(CStr(arrKeys(i)), ",")(1))
'        strValues = CStr(arrKeys(i)) & "|" & lngRow & "|" & CStr(arrValues(i))
'        Call Record_Update(mrsCell, strFileds, strValues, "ID|" & CStr(arrKeys(i)))
'        VsfData.TextMatrix(lngRow, lngCol) = Split(CStr(arrValues(i)), "|")(1)
'    Next i

    mrsCell.Filter = 0
    strKeys = ""
    '��������ͬ��������������Դ
    mrsCell.Filter = "�к�=" & lngRow
    With mrsCell
        Do While Not .EOF
            If InStr(1, ",0,9,", "," & Val(Nvl(mrsCell!������Դ)) & ",") = 0 Then
                strKey = Nvl(mrsCell!Id, ",")
                intCOl = Val(Split(strKey, ",")(1))
                strKeys = strKeys & "|" & strKey
            Else
                If mblnChage = False Then mblnChage = True
            End If
        .MoveNext
        Loop
    End With
    
    mrsCell.Filter = 0
    '���������Դ��¼��
    If Left(strKeys, 1) = "|" Then strKeys = Mid(strKeys, 2)
    If strKeys <> "" Then
        arrID = Split(strKeys, "|")
        For i = 0 To UBound(arrID)
            mrsCell.Filter = "ID='" & CStr(arrID(i)) & "'"
            mrsCell!������Դ = 0
            mrsCell!״̬ = 1
            mrsCell.Update
            blnUpdate = True
        Next i
    End If
    mrsCell.Filter = 0
    strKey = ""
    
    '���õ��Ƿ����ͬ������������
    mstrSQL = "SELECT C.��Ŀ���,C.��¼����,C.������Դ,C.���²�λ,C.δ��˵��,C.��ԴID,C.����,C.��ʾ,DECODE(C.��Ŀ���,-1,1,C.��¼���) ��¼���" & vbNewLine & _
        " FROM ���˻����ļ� A,���˻������� B,���˻�����ϸ C,�����¼��Ŀ D" & vbNewLine & _
        " WHERE A.ID=B.�ļ�ID AND B.ID=C.��¼ID AND A.ID=[1] AND A.����ID=[2] AND A.��ҳID=[3]" & vbNewLine & _
        " AND nvl(C.��ԴID,0)>0 And Mod(C.��¼����,5)<>5 AND C.��ֹ�汾 IS NULL  AND B.����ʱ��=[4] And C.��Ŀ���=D.��Ŀ��� AND nvl(D.����ȼ�,3) >=[5] And Nvl(D.���ò���,0) In (0,[6])" & vbNewLine & _
        " Order By B.����ʱ��,DECODE(C.��Ŀ���,-1,1,0),DECODE(C.��Ŀ���,-1,1,C.��¼���)"
    Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "ͬ������", mlng�ļ�ID, mlng����ID, mlng��ҳID, CDate(strTime), Val(VsfData.TextMatrix(lngRow, c����ȼ�)), IIf(Val(VsfData.TextMatrix(lngRow, cӤ��)) = 0, 1, 2))
    
    If rsTemp.RecordCount = 0 Then GoTo NextPos
    
    For i = RootCol To VsfData.Cols - 1
        lngItemNO = Val(VsfData.TextMatrix(0, i))
        
        If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst

        strContent = ""
        strPart = ""
        strNote = ""
        lng��ԴID = 0
        int���� = 0
        int��ʾ = 0
        intModify = 0
        With rsTemp
            Do While Not .EOF
                If lngItemNO <> 4 Then
                    blnAllow = False
                    bln���ʵ��� = False
                    intModify = 0
                    
                    If InStr(1, "," & gint���� & "," & gint���� & "," & gint���� & ",", "," & Val(Nvl(!��Ŀ���)) & ",") > 0 Then
                        Select Case Val(Nvl(!��Ŀ���))
                            Case gint����
                                If gint���� = lngItemNO Then blnAllow = True
                            Case gint����
                                If gint���� = lngItemNO Then blnAllow = True
                            Case gint����
                                If bln���� = True And mint����Ӧ�� = 2 Then
                                    If gint���� = lngItemNO Then blnAllow = True
                                Else
                                    If gint���� = lngItemNO Then blnAllow = True: bln���ʵ��� = True
                                End If
                        End Select
                        
                        If blnAllow = True Then
                            If Val(Nvl(!��¼���)) = 0 And InStr(1, ",0,9,", "," & Val(Nvl(!������Դ)) & ",") = 0 Then
                                    
                                strContent = Nvl(!��¼����)
                                strPart = Nvl(!���²�λ)
                                lng��ԴID = Val(Nvl(!��ԴID))
                                int���� = Val(Nvl(!����))
                                int��ʾ = Val(Nvl(!��ʾ))
                                strNote = Nvl(!δ��˵��)
                            Else '��װ�����º���������
                                If bln���ʵ��� = False Then
                                    If strContent <> "" Then
                                        If InStr(1, strContent, "/") = 0 Then
                                            strContent = strContent & "/" & Nvl(!��¼����)
                                        Else
                                            strContent = Split(strContent, "/")(0) & "/" & Nvl(!��¼����)
                                        End If
                                    Else
                                        strContent = Nvl(!��¼����)
                                    End If
                                    
                                    If InStr(1, ",0,9,", "," & Val(Nvl(!������Դ)) & ",") = 0 Then
                                        intModify = 1
                                    End If
                                    
                                    Exit Do
                                Else
                                    If InStr(1, ",0,9,", "," & Val(Nvl(!������Դ)) & ",") = 0 Then
                                        strPart = Nvl(!���²�λ)
                                        lng��ԴID = Val(Nvl(!��ԴID))
                                        int���� = Val(Nvl(!����))
                                        int��ʾ = Val(Nvl(!��ʾ))
                                        intModify = 1
                                        strContent = Nvl(!��¼����)
                                        strNote = Nvl(!δ��˵��)
                                        Exit Do
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If Val(Nvl(!��Ŀ���)) = lngItemNO And InStr(1, ",0,9,", "," & Val(Nvl(!������Դ)) & ",") = 0 Then
                            strPart = Nvl(!���²�λ)
                            lng��ԴID = Val(Nvl(!��ԴID))
                            int���� = Val(Nvl(!����))
                            int��ʾ = Val(Nvl(!��ʾ))
                            strContent = Nvl(!��¼����)
                            strNote = Nvl(!δ��˵��)
                            intModify = 1
                            Exit Do
                        End If
                    End If
                ElseIf InStr(1, ",4,5,", "," & Val(!��Ŀ���) & ",") <> 0 And lngItemNO = 4 Then
                    Select Case Val(!��Ŀ���)
                        Case 4
                            If strContent <> "" Or Nvl(!��¼����) <> "" Then
                                If InStr(1, strContent, "/") > 0 Then
                                    strContent = Nvl(!��¼����) & "/" & Trim(Split(strContent, "/")(1))
                                Else
                                    strContent = Nvl(!��¼����) & "/"
                                End If
                                strNote = Nvl(!δ��˵��)
                                strPart = Nvl(!���²�λ)
                                lng��ԴID = Val(Nvl(!��ԴID))
                                int���� = Val(Nvl(!����))
                                int��ʾ = Val(Nvl(!��ʾ))
                                intModify = 1 '���ܽ����޸�
                            End If
                        Case 5
                            If strContent <> "" Or Nvl(!��¼����) <> "" Then
                                If InStr(1, strContent, "/") > 0 Then
                                    strContent = Trim(Split(strContent, "/")(0)) & "/" & Nvl(!��¼����)
                                Else
                                    strContent = "/" & Nvl(!��¼����)
                                End If
                            End If
                    End Select
                End If
                .MoveNext
            Loop
            
            If strContent = "/" Then strContent = ""
            If lngItemNO = 4 Then
                If InStr(1, strContent, "/") <> 0 Then
                    If Not IsNumeric(Split(strContent, "/")(0)) And Not IsNumeric(Split(strContent, "/")(1)) Then
                        strContent = ""
                    End If
                End If
            End If
            
            '���������������Ŀ�����Ҳ�λ��Ϊ��
            mrsItems.Filter = "��Ŀ���=" & lngItemNO
            If mrsItems.RecordCount > 0 Then
                If Nvl(mrsItems!��¼��, 2) = 1 Then
                    If strNote <> "" And strContent = "" Then
                        strContent = strNote
                    Else
                        If strContent <> "" Then strContent = IIf(strPart = "", "", strPart & ":") & strContent
                    End If
                End If
            End If

            If strContent <> "" Then
                '��ͬ��������װ�ص���¼����
                strKey = lngRow & "," & i
                strValues = strKey & "|" & lngRow & "|" & lngItemNO & "|" & strContent & "|" & strPart & "|1|" & lng��ԴID & "|" & int���� & "|" & int��ʾ & "|" & intModify & "|0"
                Call Record_Update(mrsCell, strFileds, strValues, "ID|" & strKey)
                VsfData.TextMatrix(lngRow, i) = strContent
                If lngItemNO = gint���� Or (lngItemNO = gint���� And mint����Ӧ�� = 2) Then
                    VsfData.Cell(flexcpForeColor, lngRow, i, lngRow, i) = RGB(0, 0, 255)
                Else
                    VsfData.Cell(flexcpForeColor, lngRow, i, lngRow, i) = 255 '&H8080FF
                End If
            End If
        End With
    Next i
    VsfData.Cell(flexcpAlignment, VsfData.FixedRows, cʱ��, VsfData.Rows - 1, VsfData.Cols - 1) = flexAlignCenterCenter
    mrsItems.Filter = 0
NextPos:
    CheckPaseDate = True
    
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function CheckTime(ByVal strTime As String, ByVal strPatientTime As String, strInfo As String) As Boolean
'-------------------------------------------------------------
'����:������ݲ�¼�ͳ���¼��
'strPatientTime ���µ���ʼ����;���˳�Ժ����
'-------------------------------------------------------------
    Dim strInTime As String, strOutTime As String, strCurrDate As String
    
    On Error GoTo Errhand
    
    strInTime = Split(strPatientTime, ";")(0)
    strOutTime = Split(strPatientTime, ";")(1)
    
    If mbln��Ժ = False Then
        strOutTime = DateAdd("d", mintPreDays, CDate(strOutTime))
    End If
    
    If Format(strTime, "YYYY-MM-DD HH:mm") > Format(strOutTime, "YYYY-MM-DD HH:mm") Then
        If mbln��Ժ = False Then
            strInfo = "��¼����ʱ���ѳ�������[����¼��������" & mintPreDays & "��]��ָ���ķ�Χ!"
        Else
            strInfo = "��¼����ʱ�䲻�ܴ���[���˳�Ժʱ�䣺" & Format(strOutTime, "YYYY-MM-DD HH:mm:ss") & "]!"
        End If
        Exit Function
    End If
    
    If Format(strTime, "YYYY-MM-DD HH:mm") < Format(strInTime, "YYYY-MM-DD HH:mm") Then
        strInfo = strInfo & "��¼����ʱ�䲻��С��[���µ���ʼʱ�䣺" & Format(strInTime, "YYYY-MM-DD HH:mm:ss") & "]!"
        Exit Function
    End If
    
    strCurrDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
    If Not IsAllowInput(mlng����ID, mlng��ҳID, strTime, strCurrDate) Then
        strInfo = "��¼����ʱ��[" & strTime & "]����![�������ݲ�¼����Чʱ��:" & mlngHours & "Сʱ]"
        Exit Function
    End If
    
    CheckTime = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function IsAllowInput(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal strTime As String, ByVal strCurTime As String) As Boolean
    'ȡ��ָ��������ָ��ʱ��֮��ؼ����ʱ��
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo Errhand
    
    IsAllowInput = True
    gstrSQL = "" & _
              " SELECT DECODE(��ֹԭ��,1,'��Ժ',3,'ת��',10,'Ԥ��Ժ',15,'ת����',DECODE(��ʼԭ��,10,'��Ժ','δ����')) AS ����,��ֹʱ�� AS ʱ��" & _
              " From ���˱䶯��¼" & _
              " WHERE (��ֹԭ�� IN (1,3,10,15) OR ��ʼԭ��=10) And ����ID=[1] And ��ҳID=[2] And [3] <= ��ֹʱ��" & _
              " ORDER BY ��ֹʱ��"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��ָ��������ָ��ʱ��֮��ؼ����ʱ��", lng����ID, lng��ҳID, CDate(strTime))
    If rsTemp.RecordCount = 0 Then Exit Function
    
    'ֻȡ��һ�����ϵļ�¼
    strTime = Format(DateAdd("H", mlngHours, rsTemp!ʱ��), "yyyy-MM-dd HH:mm")
    strCurTime = Format(strCurTime, "yyyy-MM-dd HH:mm")
    
    If strTime < strCurTime Then IsAllowInput = False
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function CheckValid(strReturn As String, ByVal lngItemNO As Long, strInfo As String) As Boolean
    Dim dblMin As Double, dblMax As Double
    Dim strText As String, strText1 As String, strName As String, strGroupName As String, strFormat As String, strFormat1 As String
    Dim arrValue() As String
    Dim strValue As String
    Dim i As Integer
    Dim blnCheck As Boolean
    Dim blnAllow As Boolean
    
    strText = Trim(strReturn)
    mrsItems.Filter = 0
    mrsItems.Filter = "��Ŀ���=" & lngItemNO
    If mrsItems.RecordCount = 0 Then Exit Function
    
    strName = mrsItems!��Ŀ����
    strGroupName = mrsItems!������
    
    blnAllow = True
    If strName = "���" Or strName = "����" Then
        blnAllow = IsNumeric(strInfo)
    Else
        blnAllow = IIf(InStr(1, "," & gint��� & "," & gint��Һ & ",", "," & lngItemNO & ",") > 0, False, True)
    End If
    
    If strText <> "" Then
        If mrsItems!��Ŀ���� = 0 And InStr(1, "0,4", Nvl(mrsItems!��Ŀ��ʾ, 0)) <> 0 Then
            If lngItemNO = 4 Then
                If InStr(1, strText, "/") = 0 Then
                    strInfo = "[Ѫѹ]���ݸ�ʽ��������ѹ/����ѹ��"
                    Exit Function
                End If
                
                '--�����53505,�޸��ˣ����Σ�Ѫѹ��ʾ����
                If Trim(Split(strText, "/")(0)) = "�ܲ�" Or Trim(Split(strText, "/")(0)) = "δ��" Or Trim(Split(strText, "/")(0)) = "���" Or Trim(Split(strText, "/")(0)) = "���" Then
                     strReturn = Trim(Split(strText, "/")(0))
                     CheckValid = True
                     Exit Function
                Else
                    If Trim(Split(strText, "/")(0)) = "" Or Trim(Split(strText, "/")(1)) = "" Then
                        strInfo = "[Ѫѹ]���ݸ�ʽ¼���������ѹ/����ѹ��"
                        Exit Function
                    End If
                End If
            ElseIf lngItemNO = gint���� And mint����Ӧ�� <> 2 Then
                If InStr(1, strText, "/") <> 0 Then
                    strInfo = "[" & strName & "]���ݸ�ʽ¼�����,���������Ƿ���������ã�"
                    Exit Function
                End If
            ElseIf lngItemNO <> gint���� And Not (lngItemNO = gint���� And mint����Ӧ�� = 2) And blnAllow = True Then
                If InStr(1, strText, "/") <> 0 Then
                    strInfo = "[" & strName & "]���ݸ�ʽ¼�����,���飡"
                    Exit Function
                End If
            End If
            
            arrValue = Split(strText, "/")
            
            For i = 0 To UBound(arrValue)
                strText = arrValue(i)
                blnCheck = False
                
                If strGroupName = "1)����������Ŀ" Then
                    If Not IsNumeric(strText) Then
                        If InStr(1, "," & mstrNote & "," & IIf(lngItemNO = gint����, ",����,", ""), "," & strText & ",") <> 0 Then
                            blnCheck = True
                        Else
                            strInfo = "[" & strName & "]���ݸ�ʽ¼�����,���飡"
                            Exit Function
                        End If
                    End If
                Else
                    If Not IsNumeric(strText) And blnAllow = True Then
                        strInfo = "[" & strName & "]���ݸ�ʽ¼�����,���飡"
                        Exit Function
                    End If
                End If
                
                If blnCheck = True Then
                    If UBound(arrValue) > 0 Then
                        strInfo = "[" & strName & "]���ݸ�ʽ¼�����,���飡"
                        Exit Function
                    End If
                End If
                
                If Nvl(mrsItems!��ĿС��, 0) <> 0 And blnAllow = True Then  '��������ͨ���ؼ���MaxLength�����Ƶ�
                    If InStr(1, strText, ".") <> 0 Then strText1 = Mid(strText, 1, InStr(1, strText, ".") - 1)
                    If Len(strText1) > mrsItems!��Ŀ���� Then
                        mrsItems.Filter = 0
                        strInfo = "[" & strName & "]¼������ݳ����˺Ϸ����ȣ�"
                        Exit Function
                    End If
        
                    If InStr(1, strText, ".") <> 0 Then
                        strText1 = Mid(strText, InStr(1, strText, ".") + 1)
                        If Len(strText1) > mrsItems!��ĿС�� Then
                            mrsItems.Filter = 0
                            strInfo = "[" & strName & "]¼���С�����ֳ����˺Ϸ����ȣ�"
                            Exit Function
                        End If
                    End If
                End If
                If Not IsNull(mrsItems!��Ŀֵ��) And Not blnCheck And blnAllow = True And Nvl(mrsItems!��Ŀ��ʾ, 0) = 0 Then
                    dblMin = Split(mrsItems!��Ŀֵ��, ";")(0)
                    dblMax = Split(mrsItems!��Ŀֵ��, ";")(1)
                    If Not (Val(strText) >= dblMin And Val(strText) <= dblMax) Then
                        mrsItems.Filter = 0
                        strInfo = "[" & strName & "]¼������ݲ���" & Format(dblMin, "#0.00") & "��" & Format(dblMax, "#0.00") & "����Ч��Χ��"
                        Exit Function
                    End If
                End If
                
                If blnCheck = True Then
                    strFormat = strText
                Else
                    strFormat = strFormat & "/" & IIf(blnAllow = True, Val(strText), strText)
                End If
                
                If i = UBound(arrValue) Then
                    If Left(strFormat, 1) = "/" Then strFormat = Mid(strFormat, 2)
                End If
            Next i
        Else '�ı�����
            If LenB(StrConv(strText, vbFromUnicode)) > mrsItems!��Ŀ���� Then
                strInfo = "[" & strName & "]¼������ݳ�������󳤶ȣ�" & mrsItems!��Ŀ���� & "��"
                mrsItems.Filter = 0
                Exit Function
            End If
            strFormat = strText
        End If
    Else
    
    End If
    
    strFormat1 = strFormat
    
    '����������Դ<>0,9�� ����,�������� ���б༭(�������º������������¼��������,��������)
    If InStr(1, ",0,9,", "," & mint������Դ & ",") = 0 Then
        If lngItemNO = gint���� Or (lngItemNO = gint���� And mint����Ӧ�� = 2) Then
            strValue = picInput.Tag
            If InStr(1, strFormat1, "/") <> 0 Then
                strFormat1 = Split(strFormat1, "/")(0)
            End If
            If InStr(1, strValue, "/") = 0 Then
                If Trim(strFormat1) <> Trim(strValue) Then
                    If lngItemNO = 1 Then
                        strInfo = "ͬ��������[" & strName & "]����ֻ�����޸������²���."
                    Else
                        strInfo = "ͬ��������[" & strName & "]����ֻ�����޸��������Ჿ��."
                    End If
                    
                    txtInput.Text = strValue
                    Exit Function
                End If
            Else
                If mintModify = 1 Then
                    If strFormat <> strValue Then
                        If lngItemNO = 1 Then
                            strInfo = "ͬ��������[" & strName & "]�����������������,�������޸�."
                        Else
                            strInfo = "ͬ��������[" & strName & "]�������������������,�������޸�."
                        End If
                        txtInput.Text = strValue
                        Exit Function
                    End If
                Else
                    If strFormat1 <> Split(strValue, "/")(0) Then
                        If lngItemNO = 1 Then
                            strInfo = "ͬ��������[" & strName & "]����ֻ�����޸������²���."
                        Else
                            strInfo = "ͬ��������[" & strName & "]����ֻ�����޸��������Ჿ��."
                        End If
                        txtInput.Text = strValue
                        Exit Function
                    End If
                End If
            End If
        End If
    End If
    
    mrsItems.Filter = 0
    strReturn = strFormat
    CheckValid = True
End Function

Private Sub AdjustRowFlag(ByRef objVsf As Object, ByVal intRow As Integer)
    '-----------------------------------------------------------------------------------------
    '����:
    '����:
    '-----------------------------------------------------------------------------------------
    If objVsf.FixedCols = 0 Then Exit Sub
    
    If Not (objVsf.Cell(flexcpPicture, intRow, 0) Is Nothing) Then Exit Sub
    Set objVsf.Cell(flexcpPicture, 1, 0, objVsf.Rows - 1, 0) = Nothing
    Set objVsf.Cell(flexcpPicture, intRow, 0) = ils16.ListImages(1).Picture
    
End Sub

Private Sub vsfHistory_DblClick()
    Call vsfHistory_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub vsfHistory_KeyDown(KeyCode As Integer, Shift As Integer)
    '���ܣ�������ʷ���ݵ�ֵ
    Dim lngItemNO As Integer
    Dim strText As String, strInfo As String, strOldValue As String, strName As String
    Dim intCOl As Integer, intRow As Integer
    Dim blnValidate As Boolean, blnSave As Boolean
    Dim strFileds As String, strValues As String, strKey As String, strPart As String
    Dim int��ԴID As Long, int���� As Integer, int��ʾ As Integer, int�޸� As Integer
    Dim intState As Integer
    Dim rsObj As New ADODB.Recordset
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If vsfHistory.Col < RootCol Or vsfHistory.Row < vsfHistory.FixedRows Then Exit Sub
    If Val(vsfHistory.RowData(vsfHistory.Row)) = 0 Then Exit Sub
    If vsfHistory.ColHidden(vsfHistory.Col) Then Exit Sub
    If mrsHistory Is Nothing Then Exit Sub
    
    strFileds = "ID|�к�|��Ŀ���|����|��λ|������Դ|��ԴID|����|��ʾ|�޸�|״̬"
    VsfData.Col = vsfHistory.Col
    intCOl = VsfData.Col
    intRow = VsfData.Row
    blnValidate = True
    blnSave = True
    strOldValue = ""
    
    mrsHistory.Filter = 0
    strKey = vsfHistory.Row & "," & vsfHistory.Col
    mrsHistory.Filter = "ID='" & strKey & "'"
    If mrsHistory.RecordCount > 0 Then '���˵�˵����������
        strText = Nvl(mrsHistory!����)
        strPart = Nvl(mrsHistory!��λ)
    Else
        strText = vsfHistory.TextMatrix(vsfHistory.Row, vsfHistory.Col)
        strPart = ""
    End If
    strOldValue = VsfData.TextMatrix(intRow, intCOl)
    
    '���Ҫ�滻�������Ƿ���ͬ��������
    strName = VsfData.TextMatrix(VsfData.FixedRows - 1, intCOl)
    mrsCell.Filter = 0
    strKey = intRow & "," & intCOl
    mrsCell.Filter = "ID='" & strKey & "'"
    If mrsCell.RecordCount > 0 Then '�����mrsCellΪ��
        Set rsObj = mrsCell.Clone
    Else
        Set rsObj = mrsCopy.Clone
    End If
    rsObj.Filter = "ID='" & strKey & "'"
    If rsObj.RecordCount > 0 Then
        lngItemNO = Val(Nvl(rsObj!��Ŀ���))
        mint������Դ = Val(Nvl(rsObj!������Դ))
        mintModify = Val(Nvl(rsObj!�޸�))
        If InStr(1, ",0,9,", "," & Val(rsObj!������Դ) & ",") = 0 Then
            If Not (lngItemNO = gint���� Or (lngItemNO = gint���� And mint����Ӧ�� = 2)) Then
                strInfo = "ͬ��������[" & strName & "]���ݲ��ܽ����޸�."
                RaiseEvent AfterRowColChange(strInfo, True)
                Exit Sub
            Else
                If mintModify = 1 Then
                    If lngItemNO = gint���� Then
                        strInfo = "ͬ��������[" & strName & "]����������������²��ܽ����޸�."
                    Else
                        strInfo = "ͬ��������[" & strName & "]������������������᲻�ܽ����޸�."
                    End If
                    RaiseEvent AfterRowColChange(strInfo, True)
                    Exit Sub
                Else
                    If InStr(1, strOldValue, "/") <> 0 Then
                        strInfo = Split(strOldValue, "/")(0)
                    Else
                        strInfo = strOldValue
                    End If
                    strInfo = Mid(strInfo, InStr(1, strInfo, ":") + 1)
                    If strInfo <> "" Then
                        If InStr(1, strText, "/") = 0 Then
                            strInfo = "ͬ��������[" & strName & "]���ݲ��ܽ����޸�."
                            RaiseEvent AfterRowColChange(strInfo, True)
                            Exit Sub
                        Else
                            strText = strInfo & "/" & Split(strText, "/")(1)
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    lngItemNO = Val(vsfHistory.TextMatrix(0, intCOl))
    If blnValidate = True Then
        mrsCopy.Filter = 0
        VsfData.TextMatrix(intRow, intCOl) = IIf(strPart = "", "", strPart & ":") & strText
        VsfData.Cell(flexcpAlignment, intRow, intCOl, intRow, intCOl) = flexAlignCenterCenter
        '�������ݴ���
        If blnSave = True Then
            If Trim(strOldValue) <> Trim(strText) Then
                strKey = intRow & "," & intCOl
                '����޸ĵ������Ƿ��Ѿ�����
                mrsCopy.Filter = "ID='" & strKey & "'"
                If mrsCopy.RecordCount > 0 Then
                    int��ԴID = Val(Nvl(mrsCopy!��ԴID))
                    int���� = Val(Nvl(mrsCopy!����))
                    int��ʾ = Val(Nvl(mrsCopy!��ʾ))
                    int�޸� = Val(Nvl(mrsCopy!�޸�))
                    intState = 1
                Else
                    int��ԴID = 0: int���� = 0: int��ʾ = 0: int�޸� = 0
                    intState = IIf(Trim(strText) = "", 3, 1)
                End If
                strValues = strKey & "|" & intRow & "|" & lngItemNO & "|" & strText & "|" & strPart & "|" & mint������Դ & "|" & _
                    int��ԴID & "|" & int���� & "|" & int��ʾ & "|" & int�޸� & "|" & intState
                Call Record_Update(mrsCell, strFileds, strValues, "ID|" & strKey)
                mblnChage = True
            End If
        End If
    End If
    
    If mblnShow And mintType <> 0 Then Call ShowInput
    mblnShow = False: mintType = 0
    
    mblnClearRow = False
    For intCOl = cʱ�� + 1 To VsfData.Cols - 1
        If Trim(VsfData.TextMatrix(intRow, intCOl)) <> "" Then
            mblnClearRow = True
            Exit For
        End If
    Next intCOl
End Sub

Private Function is������Һ(ByVal intType As Integer) As Boolean
'����Ƿ��Ǵ����Ŀ����ҹ��Ŀ  �����Ŀ���=10 ��ҹ=9
'intType=1 Ϊ�����Ŀ ����Ϊ��Һ��Ŀ
    Dim lngItemNO As Long
    Dim strKey As String
    Dim rsObj As New ADODB.Recordset
    
    On Error GoTo Errhand
    
    If VsfData.Col < RootCol Or VsfData.Row < VsfData.FixedRows Then Exit Function
    If mblnInit = False Or mblnNullRow = False Then Exit Function
    
    '��ȡ��Ŀ���
    lngItemNO = Val(VsfData.TextMatrix(0, VsfData.Col))
    If intType = 1 Then
        If lngItemNO <> 10 Then Exit Function
    Else
        If lngItemNO <> 9 Then Exit Function
    End If
    
    mrsItems.Filter = "��Ŀ���=" & lngItemNO
    If InStr(1, ",2,3,5,", "," & Val(Nvl(mrsItems!��Ŀ��ʾ)) & ",") > 0 Then Exit Function
    
    '����Ƿ���ͬ��������
    mrsCell.Filter = 0
    strKey = VsfData.Row & "," & VsfData.Col
    mrsCell.Filter = "ID='" & strKey & "'"
    If mrsCell.RecordCount > 0 Then '�����mrsCellΪ��
        Set rsObj = mrsCell.Clone
    Else
        Set rsObj = mrsCopy.Clone
    End If
    rsObj.Filter = "ID='" & strKey & "'"
    
    If rsObj.RecordCount > 0 Then
        If InStr(1, ",0,9,", "," & Val(rsObj!������Դ) & ",") = 0 Then Exit Function
    End If
    
    is������Һ = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function Get��������(ByVal lng����ID As Long)
    Dim rsTemp As ADODB.Recordset
    gstrSQL = "select �������� from ��������˵�� where ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��������", lng����ID)
    If Not rsTemp.BOF Then
        mstr�������� = rsTemp!��������
    Else
        mstr�������� = "����"
    End If
End Function


