VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FBAFE9A8-8B26-4559-9D12-D70E36A97BE3}#2.1#0"; "zlRichEditor.ocx"
Begin VB.Form frmPathAppend 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���·������Ŀ"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11925
   Icon            =   "frmPathAppend.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   11925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fraERPType 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   8040
      TabIndex        =   39
      Top             =   3930
      Width           =   2535
      Begin VB.OptionButton optEPRType 
         Caption         =   "�ϰ�"
         Height          =   180
         Index           =   0
         Left            =   960
         TabIndex        =   41
         Top             =   0
         Width           =   735
      End
      Begin VB.OptionButton optEPRType 
         Caption         =   "�°�"
         Height          =   180
         Index           =   1
         Left            =   1680
         TabIndex        =   40
         Top             =   0
         Width           =   735
      End
      Begin VB.Label lblEPR 
         Caption         =   "�����汾"
         Height          =   180
         Left            =   120
         TabIndex        =   42
         Top             =   0
         Width           =   855
      End
   End
   Begin VB.ComboBox cboItems 
      Height          =   300
      Left            =   3720
      Style           =   2  'Dropdown List
      TabIndex        =   37
      ToolTipText     =   "Ĭ�ϲ���ѡ��λ��֮ǰ"
      Top             =   960
      Width           =   2175
   End
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   11400
      MouseIcon       =   "frmPathAppend.frx":1708A
      MousePointer    =   99  'Custom
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   3885
      Width           =   330
   End
   Begin VB.Frame fraVariation 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   2760
      Left            =   6000
      TabIndex        =   32
      Top             =   960
      Width           =   5895
      Begin VB.TextBox txtVariation 
         Height          =   300
         Left            =   3885
         MaxLength       =   1000
         TabIndex        =   4
         Top             =   0
         Width           =   1890
      End
      Begin VSFlex8Ctl.VSFlexGrid vsVariation 
         Height          =   2380
         Left            =   0
         TabIndex        =   5
         Top             =   360
         Width           =   5775
         _cx             =   10186
         _cy             =   4198
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
         BackColorFixed  =   15597549
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   16777215
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   32768
         GridColorFixed  =   32768
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   320
         ColWidthMin     =   0
         ColWidthMax     =   8000
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmPathAppend.frx":171DC
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
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
         AllowUserFreezing=   0
         BackColorFrozen =   14811105
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lblSearch 
         Caption         =   "����(&F)"
         Height          =   255
         Left            =   3120
         TabIndex        =   3
         Top             =   60
         Width           =   735
      End
      Begin VB.Label lblVariation 
         Caption         =   "����ԭ��"
         Height          =   255
         Left            =   0
         TabIndex        =   33
         Top             =   60
         Width           =   855
      End
   End
   Begin VB.Frame fraItemKind 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   4440
      TabIndex        =   29
      Top             =   3930
      Width           =   3615
      Begin VB.OptionButton optType 
         Caption         =   "������"
         Height          =   180
         Index           =   1
         Left            =   1680
         TabIndex        =   7
         Top             =   0
         Width           =   840
      End
      Begin VB.OptionButton optType 
         Caption         =   "ҽ����"
         Height          =   180
         Index           =   0
         Left            =   810
         TabIndex        =   6
         Top             =   0
         Value           =   -1  'True
         Width           =   840
      End
      Begin VB.OptionButton optType 
         Caption         =   "������"
         Height          =   180
         Index           =   2
         Left            =   2550
         TabIndex        =   8
         Top             =   0
         Width           =   840
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ŀ����"
         Height          =   180
         Left            =   0
         TabIndex        =   30
         Top             =   0
         Width           =   720
      End
   End
   Begin VB.Frame fraExecute 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1420
      Left            =   0
      TabIndex        =   26
      Top             =   6650
      Width           =   11895
      Begin VB.Frame fraExecutor 
         Appearance      =   0  'Flat
         Caption         =   "ִ����"
         ForeColor       =   &H80000008&
         Height          =   1275
         Left            =   3480
         TabIndex        =   27
         Top             =   90
         Width           =   1305
         Begin VB.OptionButton optExecutor 
            Caption         =   "��ʿ"
            Height          =   180
            Index           =   1
            Left            =   240
            TabIndex        =   10
            Top             =   720
            Width           =   660
         End
         Begin VB.OptionButton optExecutor 
            Caption         =   "ҽ��"
            Height          =   180
            Index           =   0
            Left            =   240
            TabIndex        =   9
            Top             =   360
            Value           =   -1  'True
            Width           =   660
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsResult 
         Height          =   1290
         Left            =   5880
         TabIndex        =   11
         Top             =   90
         Width           =   5895
         _cx             =   10398
         _cy             =   2275
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
         BackColorFixed  =   15597549
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   32768
         GridColorFixed  =   32768
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
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmPathAppend.frx":17241
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
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   5
         X1              =   0
         X2              =   11880
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   4
         X1              =   0
         X2              =   11880
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   120
         Picture         =   "frmPathAppend.frx":172BF
         Top             =   120
         Width           =   720
      End
      Begin VB.Label lblExePrompt 
         Caption         =   $"frmPathAppend.frx":17ABD
         Height          =   1335
         Left            =   960
         TabIndex        =   31
         Top             =   90
         Width           =   2295
      End
      Begin VB.Label lblResult 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ִ�н��"
         Height          =   180
         Left            =   5040
         TabIndex        =   28
         Top             =   90
         Width           =   720
      End
   End
   Begin VB.TextBox txtReason 
      Height          =   2415
      Left            =   1080
      MaxLength       =   1000
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1320
      Width           =   4815
   End
   Begin VB.ComboBox cboItemType 
      Height          =   300
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox txtItem 
      Height          =   300
      Left            =   1050
      MaxLength       =   1000
      TabIndex        =   1
      Top             =   3900
      Width           =   3255
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00EFF0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   11925
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   8055
      Width           =   11925
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   9435
         TabIndex        =   14
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   10665
         TabIndex        =   15
         Top             =   120
         Width           =   1100
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   0
         X2              =   11880
         Y1              =   30
         Y2              =   30
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   0
         X2              =   11760
         Y1              =   45
         Y2              =   45
      End
   End
   Begin VB.PictureBox picInfo 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   11925
      TabIndex        =   16
      Top             =   0
      Width           =   11925
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   2
         X1              =   0
         X2              =   11880
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   3
         X1              =   0
         X2              =   11880
         Y1              =   825
         Y2              =   825
      End
      Begin VB.Image imgInfo 
         Height          =   720
         Left            =   195
         Picture         =   "frmPathAppend.frx":17B5F
         Top             =   45
         Width           =   720
      End
      Begin VB.Label lblNote 
         BackStyle       =   0  'Transparent
         Caption         =   "    ��׼·���ж������Ŀ�������㲡��ʵ�����󣬵��ֲ�����������Щ���ص�Ӱ����˳�·����������ʱ���·������Ŀ��"
         Height          =   360
         Left            =   1095
         TabIndex        =   18
         Top             =   360
         Width           =   10605
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "·������Ŀ"
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
         Left            =   1095
         TabIndex        =   17
         Top             =   120
         Width           =   975
      End
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   0
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin zlRichEditor.Editor edtEditor 
      Height          =   495
      Left            =   0
      TabIndex        =   24
      Top             =   2040
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
   End
   Begin MSComctlLib.ImageList imgNature 
      Left            =   0
      Top             =   2760
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
            Picture         =   "frmPathAppend.frx":196A1
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathAppend.frx":19C3B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathAppend.frx":1A1D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathAppend.frx":1A76F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathAppend.frx":1AD09
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathAppend.frx":1B2A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathAppend.frx":1B83D
            Key             =   "Selected"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathAppend.frx":1BBD7
            Key             =   "UnSelected"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraAdvice 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2380
      Left            =   120
      TabIndex        =   21
      Top             =   4250
      Width           =   11655
      Begin zlCISPath.UCAdviceList UCAdvice 
         Height          =   1935
         Left            =   0
         TabIndex        =   34
         Top             =   360
         Width           =   11655
         _extentx        =   20558
         _extenty        =   3413
      End
      Begin VB.CommandButton cmdAdvice 
         Caption         =   "ҽ���༭(&E)"
         Height          =   350
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   1380
      End
   End
   Begin VB.Frame fraEPR 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2380
      Left            =   120
      TabIndex        =   20
      Top             =   4250
      Width           =   11775
      Begin VSFlex8Ctl.VSFlexGrid vsEPR 
         Height          =   2375
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   11640
         _cx             =   20532
         _cy             =   4189
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
         BackColorFixed  =   15597549
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   32768
         GridColorFixed  =   32768
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
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   1
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   270
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmPathAppend.frx":1BF71
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
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   7
         X1              =   0
         X2              =   10000
         Y1              =   1335
         Y2              =   1335
      End
   End
   Begin VB.Label lblItemPosition 
      AutoSize        =   -1  'True
      Caption         =   "����λ��"
      Height          =   180
      Left            =   2880
      TabIndex        =   38
      Top             =   1020
      Width           =   720
   End
   Begin VB.Label lblIcon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��Ŀͼ��"
      Height          =   180
      Left            =   10680
      TabIndex        =   36
      Top             =   3960
      Width           =   720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   8
      X1              =   120
      X2              =   11780
      Y1              =   3810
      Y2              =   3810
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   6
      X1              =   120
      X2              =   11780
      Y1              =   3825
      Y2              =   3825
   End
   Begin VB.Label lblReason 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����˵��"
      Height          =   180
      Left            =   210
      TabIndex        =   25
      Top             =   1320
      Width           =   720
   End
   Begin XtremeCommandBars.CommandBars cbsIcon 
      Left            =   0
      Top             =   720
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Label lbl���� 
      AutoSize        =   -1  'True
      Caption         =   "��Ŀ����"
      Height          =   180
      Left            =   210
      TabIndex        =   23
      Top             =   1020
      Width           =   720
   End
   Begin VB.Label lblItem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��Ŀ����"
      Height          =   180
      Left            =   210
      TabIndex        =   22
      Top             =   3960
      Width           =   720
   End
End
Attribute VB_Name = "frmPathAppend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'��ڲ���
Private mlngFun As Long '0-ֱ�����,1-ҽ���¿�ʱ���,2-�޸�·������Ŀ(���ṩҽ���Ͳ����ļ��������޸�)
Private mfrmParent As Object
Private mPP As TYPE_PATH_Pati
Private mPati As TYPE_Pati
Private mstrItemType As String
Private mint���� As Integer
Private mstrҽ��IDs As String   'mlngFun=1ʱ����
Private mlngִ��ID As Long       'mlngFun=1��2ʱ����
Private mclsMipModule As zl9ComLib.clsMipModule ' ��Ϣƽ̨����

'�������
Private mrsResult As ADODB.Recordset 'ִ�н����
Private mrsNature As ADODB.Recordset '������ʼ�

Private mblnUseExecute As Boolean

Private mvItem As TYPE_PATH_ITEM
Private mblnReturn As Boolean
Private mblnOK As Boolean
Private marrSQL As Variant          'ҽ��У�ԵĲ������SQL
Private mdatPathOut As Date     'mlngFun=0 �� mlngFun=1 ʱ,סԺҽ���༭���淵��
Private mlng�׶�ID As Long      'mlngFun=0ʱ,·������Ŀ���ɵĽ׶�ID
Private mlng���� As Long        'mlngFun=0ʱ,·������Ŀ���ɵ�����

Private Enum CONST_COL_ִ�н��
    colִ��ͼ�� = 0
    colִ�н�� = 1
    col������� = 2
    colȱʡ��� = 3
End Enum

Private Enum CONST_COL_����ԭ��
    col������� = 0
    col����ԭ�� = 1
    col����ѡ�� = 2
End Enum


Public Function ShowMe(frmParent As Object, int���� As Integer, t_pati As TYPE_Pati, t_pp As TYPE_PATH_Pati, _
    ByVal strItemType As String, ByVal bytUseType As Byte, ByVal strҽ��IDs As String, ByVal lngִ��ID As Long, Optional ByRef objMip As Object, _
    Optional ByVal datDate As Date = CDate(0)) As Boolean
    
    Set mfrmParent = frmParent
    mint���� = int����
    mPati = t_pati
    mPP = t_pp
    mstrItemType = strItemType
    mlngFun = bytUseType
    mstrҽ��IDs = strҽ��IDs
    mlngִ��ID = lngִ��ID
    mblnOK = False
    mdatPathOut = datDate
    mlng�׶�ID = 0
    mlng���� = 0
    If Not (objMip Is Nothing) Then Set mclsMipModule = objMip
    Me.Show 1, frmParent
    ShowMe = mblnOK
End Function

Private Sub LoadItemType()
'���ܣ�����·������
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim i As Long

    strSql = "Select ���� From �ٴ�·������ Where ·��ID = [1] And �汾�� = [2] And NVL(��֧ID,0)=[3] Order by ���"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.·��ID, mPP.�汾��, mPP.��ǰ�׶η�֧ID)
    cboItemType.Clear
    For i = 1 To rsTmp.RecordCount
        cboItemType.AddItem rsTmp!����
        If rsTmp!���� = mstrItemType Then cboItemType.ListIndex = cboItemType.NewIndex
        rsTmp.MoveNext
    Next
    If cboItemType.ListIndex = -1 And cboItemType.ListCount > 0 Then cboItemType.ListIndex = 0
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cboItemType_Click()
  Call LoadCurrentItems
End Sub

Private Sub cbsIcon_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Control.ID = -1 Then
        picIcon.Cls
        mvItem.ͼ��ID = 0
    Else
        mvItem.ͼ��ID = Control.ID
        Call DrawPicture(GetPathIcon(mvItem.ͼ��ID))
    End If
End Sub

Private Sub DrawPicture(objPic As StdPicture)
    Dim X As Long, Y As Long, W As Long, H As Long
    
    W = picIcon.ScaleX(objPic.Width, vbHimetric, vbTwips)
    H = picIcon.ScaleY(objPic.Height, vbHimetric, vbTwips)
    
    X = (picIcon.ScaleWidth - W) / 2
    Y = (picIcon.ScaleHeight - H) / 2
    
    picIcon.PaintPicture objPic, X, Y, W, H
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub Form_Activate()
    If vsVariation.Enabled And vsVariation.Visible Then vsVariation.SetFocus
End Sub

Private Sub Form_Load()
    Dim vItem As TYPE_PATH_ITEM

    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsIcon.VisualTheme = xtpThemeOffice2003
    cbsIcon.ActiveMenuBar.Visible = False
    With Me.cbsIcon.Options
        .ToolBarAccelTips = True
        .LargeIcons = False
        .SetIconSize False, 16, 16
    End With
    fraVariation.BackColor = Me.BackColor
    fraAdvice.BackColor = Me.BackColor
    fraEPR.BackColor = Me.BackColor
    fraExecute.BackColor = Me.BackColor

    mblnUseExecute = Val(zlDatabase.GetPara("�Ƿ�����·��ִ�л���", glngSys, p�ٴ�·��Ӧ��, 1))
    If mblnUseExecute = False Then
        fraExecute.Visible = False
        fraAdvice.Height = fraAdvice.Height + fraExecute.Height + 60
        fraEPR.Height = fraAdvice.Height
        vsEPR.Height = fraEPR.Height
    Else
        Call Init����ִ�н��
    End If

    vsEPR.Row = 0: vsEPR.Row = 1: vsEPR.Col = 0
    vsResult.Row = 0: vsResult.Row = 1: vsResult.Col = 0
    
    Call LoadItemType
    
    mvItem = vItem  '���֮ǰ���ܱ�������Ϣ
    Call InitVariation
    
    If mlngFun = 1 Or mlngFun = 2 Then
        fraItemKind.Enabled = False
        optType(0).Enabled = False: optType(1).Enabled = False: optType(2).Enabled = False
        cboItemType.Enabled = True

        cmdAdvice.Visible = False
        UCAdvice.Top = 0
        UCAdvice.Height = fraAdvice.Height
    ElseIf mlngFun = 0 Then
        If mint���� = 1 Then '��ʿ����,ֻ������ӣ������ࣩ�ı�
            optType(0).Enabled = False: optType(1).Enabled = False: optType(2).Enabled = True
            optType(2).Value = True
            optExecutor(1).Value = True
        ElseIf mint���� = 0 Then
            optExecutor(0).Value = True
        End If
        cboItemType.Enabled = True
    End If
    
    If mlngFun = 1 Then
        optType(0).Value = True
        cmdOK.Left = cmdCancel.Left
        cmdCancel.Visible = False

        mvItem.ҽ��IDs = mstrҽ��IDs
        Call ShowAdvice
    ElseIf mlngFun = 2 Then
        Me.Caption = "�޸�·������Ŀ"
        Call LoadData
        Call SetFormFace
        fraEPR.Enabled = False: vsEPR.Enabled = False
    Else
        UCAdvice.Height = fraAdvice.Height - cmdAdvice.Height - 60
    End If

    Call Cbo.SetListHeight(cboItems, 500)
    Call Cbo.SetListWidth(cboItems.Hwnd, 3500)
    Call SetFormFace
End Sub

Private Sub LoadData()
'���ܣ���ȡ·������Ŀ�������Ϣ
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim i As Long, strҽ��IDs As String, lngRow As Long
    Dim j As Long, str��Ŀ��� As String, strȱʡ��� As String
    Dim arrtmp As Variant, arrtmp2 As Variant
 
    strSql = "Select a.����,a.��Ŀ����,a.ִ����,a.��Ŀ���,a.���ԭ��,a.ͼ��ID,a.����ԭ��,b.����ҽ��ID " & vbNewLine & _
            " From ����·��ִ�� A,����·��ҽ�� B Where a.ID = [1] And a.id = b.·��ִ��id(+)"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngִ��ID)
    For i = 1 To rsTmp.RecordCount
        If i = 1 Then
            Call Cbo.Locate(cboItemType, rsTmp!����)
            txtItem.Text = "" & rsTmp!��Ŀ����
            txtReason.Text = "" & rsTmp!���ԭ��
            
            If Not IsNull(rsTmp!����ԭ��) Then
                lngRow = vsVariation.FindRow(CStr(rsTmp!����ԭ��)) '���������rowdata
                If lngRow > 0 Then
                    vsVariation.Row = lngRow
                    vsVariation.TopRow = lngRow
                    Call vsVariation_DblClick
                End If
            End If
            
            mvItem.ͼ��ID = Val("" & rsTmp!ͼ��ID)
            If mvItem.ͼ��ID <> 0 Then Call DrawPicture(GetPathIcon(mvItem.ͼ��ID))
            
            If mblnUseExecute Then
                If Not IsNull(rsTmp!ִ����) Then
                    If rsTmp!ִ���� = 1 Then
                        optExecutor(0).Value = True
                    Else
                        optExecutor(1).Value = True
                    End If
                End If
                
                If Not IsNull(rsTmp!��Ŀ���) Then
                    arrtmp = Split(rsTmp!��Ŀ���, vbTab)
                    str��Ŀ��� = arrtmp(0)
                    If UBound(arrtmp) > 0 Then strȱʡ��� = arrtmp(1)
                    
                    arrtmp = Split(str��Ŀ���, ",")
                    With vsResult
                        .Rows = .FixedRows + UBound(arrtmp) + 2 '�������һ������
                        For j = 0 To UBound(arrtmp)
                            arrtmp2 = Split(arrtmp(j), "|")
                            lngRow = .FixedRows + j
                            .TextMatrix(lngRow, colִ�н��) = arrtmp2(0)
                            If arrtmp2(0) = strȱʡ��� Then
                                .TextMatrix(lngRow, colȱʡ���) = 1
                                Call vsResult_AfterEdit(lngRow, colȱʡ���)
                            End If
                            
                            If UBound(arrtmp2) > 0 Then
                                mrsNature.Filter = "����='" & arrtmp2(1) & "'"
                                Set .Cell(flexcpPicture, lngRow, colִ��ͼ��) = imgNature.ListImages(Val(arrtmp2(1))).Picture
                                .TextMatrix(lngRow, col�������) = mrsNature!����
                            End If
                        Next
                    End With
                End If
            End If
        End If
        
        If Not IsNull(rsTmp!����ҽ��id) Then
            strҽ��IDs = strҽ��IDs & "," & rsTmp!����ҽ��id
        End If
        rsTmp.MoveNext
    Next
    
    If strҽ��IDs <> "" Then
        optType(0).Value = True
        mvItem.ҽ��IDs = Mid(strҽ��IDs, 2)
        Call ShowAdvice
    Else
        strSql = "Select a.�������� From ���Ӳ�����¼ a Where a.·��ִ��id = [1]"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngִ��ID)
        If rsTmp.RecordCount = 0 Then
            optType(2).Value = True
        Else
            optType(1).Value = True
            vsEPR.Rows = vsEPR.FixedRows + rsTmp.RecordCount
            For i = 1 To rsTmp.RecordCount
                vsEPR.TextMatrix(vsEPR.FixedRows - 1 + i, 0) = rsTmp!��������
                rsTmp.MoveNext
            Next
        End If
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Init����ִ�н��()
'���ܣ���ʼ�����ػ���ִ�н��
    Dim strSql As String, rsTmp As ADODB.Recordset
          
    '��ȡ������ʼ�
    On Error GoTo errH
    strSql = "Select ����,���� From ·��������� Order by ����"
    Set mrsNature = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(mrsNature, strSql, Me.Caption)
    strSql = ""
    Do While Not mrsNature.EOF
        strSql = strSql & "|" & mrsNature!���� & "-" & mrsNature!����
        mrsNature.MoveNext
    Loop
    vsResult.ColData(col�������) = Mid(strSql, 2)
    
    '��ȡ���ý����
    strSql = "Select A.����,A.����,Nvl(����,0) as ����,B.���� as ����" & _
        " From ·��������� A,·��������� B" & _
        " Where A.ĩ��=1 And Nvl(A.����,0)=B.����(+)" & _
        " Order by A.����"
    Set mrsResult = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(mrsResult, strSql, Me.Caption)
    
    If mlngFun <> 2 Then
        mrsResult.Filter = "����=1"
        If Not mrsResult.EOF Then
            vsResult.Rows = vsResult.FixedRows + 1
            Call SetResultInput(vsResult.FixedRows, mrsResult)
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub optType_Click(Index As Integer)
    If Index = 1 Then
        If InStr(GetInsidePrivs(pסԺ��������), ";������д;") = 0 Then
            MsgBox "��û�в�����д��Ȩ�ޣ��������ɰ���������·����Ŀ��", vbInformation + vbOKOnly, gstrSysName
            optType(2).Value = True
            Exit Sub
        End If
    ElseIf Index = 0 Then
        If InStr(GetInsidePrivs(pסԺҽ���´�), ";ҽ���´�;") = 0 Then
            MsgBox "��û��ҽ���´��Ȩ�ޣ���������ҽ�����·����Ŀ��", vbInformation + vbOKOnly, gstrSysName
            optType(2).Value = True
            Exit Sub
        End If
    End If
    
    Call SetFormFace
    
    If Visible Then
        If Index = 0 Then
            If cmdAdvice.Enabled Then cmdAdvice.SetFocus
        ElseIf Index = 1 Then
            vsEPR.SetFocus
        End If
    End If
End Sub


Private Sub picBottom_GotFocus()
    If cmdOK.Enabled And cmdOK.Visible Then
        cmdOK.SetFocus
    ElseIf cmdCancel.Enabled And cmdCancel.Visible Then
        cmdCancel.SetFocus
    End If
End Sub

Private Sub SetFormFace()
'���ܣ����������������ý���Ŀɼ����ݺͳߴ�
    fraAdvice.Enabled = optType(0).Value: fraAdvice.Visible = fraAdvice.Enabled
    fraEPR.Enabled = optType(1).Value: fraEPR.Visible = fraEPR.Enabled
    If optType(1).Value Then
        If gobjEmr Is Nothing Then
            fraERPType.Visible = False
            optEPRType(0).Value = True
        Else
            fraERPType.Visible = True
            optEPRType(1).Value = True
        End If
    Else
        fraERPType.Visible = False
        fraItemKind.Move txtItem.Left + txtItem.Width + 120
    End If
    If fraERPType.Visible Then
        txtItem.Width = txtReason.Width - 1600
        fraItemKind.Left = txtItem.Left + txtItem.Width + 120
    Else
        txtItem.Width = txtReason.Width
        fraItemKind.Left = txtItem.Left + txtItem.Width + 120
    End If
End Sub

Private Sub picIcon_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, i As Long
    
    Dim objPopup As CommandBar
    Dim objControl As CommandBarControl
    Dim vPoint As POINTAPI
    
    On Error GoTo errH
    
    If img16.ListImages.count = 0 Then
        strSql = "Select ID,Nvl(����,0) as ���� From �ٴ�·��ͼ�� Order by ���� Desc,ID"
        Set rsTmp = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(rsTmp, strSql, Me.Caption)
        Do While Not rsTmp.EOF
            img16.ListImages.Add , "_" & IIf(rsTmp!���� = 1, 1, -1) * rsTmp!ID, GetPathIcon(rsTmp!ID)
            img16.ListImages(img16.ListImages.count).Tag = CStr(rsTmp!ID) 'ҪCStr
            rsTmp.MoveNext
        Loop
        cbsIcon.AddImageList img16
    End If
    
    Set objPopup = cbsIcon.Add("Popup", xtpBarPopup)
    objPopup.SetPopupToolBar True
    objPopup.Width = 260
    With objPopup.Controls
        Set objControl = .Add(xtpControlButton, -1, "�����Ŀͼ��")
        objControl.Flags = xtpFlagControlStretched
        For i = 1 To img16.ListImages.count
            Set objControl = .Add(xtpControlButton, img16.ListImages(i).Tag, "")
            If i = 1 Then
                objControl.BeginGroup = True
            ElseIf Val(Mid(img16.ListImages(i).Key, 2)) < 0 Then
                If Val(Mid(img16.ListImages(i - 1).Key, 2)) > 0 Then
                    objControl.BeginGroup = True
                End If
            End If
        Next
    End With
    
    vPoint.X = (picIcon.Left + picIcon.Width) / Screen.TwipsPerPixelX + 1
    vPoint.Y = picIcon.Top / Screen.TwipsPerPixelY
    ClientToScreen Me.Hwnd, vPoint
    objPopup.ShowPopup , vPoint.X * Screen.TwipsPerPixelX, vPoint.Y * Screen.TwipsPerPixelY
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtItem_GotFocus()
    Call zlControl.TxtSelAll(txtItem)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Or KeyAscii = Asc("|") Then
        KeyAscii = 0
    ElseIf KeyAscii = 13 Then
        If TypeName(ActiveControl) <> "VSFlexGrid" Then
            KeyAscii = 0
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub cmdOK_Click()
    Dim strFilter As String
    Dim i As Long
    Dim str����ԭ�� As String
    Dim strTmp As String
    
    '���ݼ��
    If Trim(txtItem.Text) = "" Then
        MsgBox "������·����Ŀ�����ݡ�", vbInformation, gstrSysName
        txtItem.SetFocus: Exit Sub
    End If
    If zlCommFun.ActualLen(txtItem.Text) > txtItem.MaxLength Then
        MsgBox "��Ŀ������������� " & txtItem.MaxLength \ 2 & " �����ֻ��� " & txtItem.MaxLength & "���ַ���", vbInformation, gstrSysName
        txtItem.SetFocus: Exit Sub
    End If

    '��������ݣ������ѡ��һ������ԭ�򣬱���˵�����Բ���
    If vsVariation.Rows > vsVariation.FixedRows Then
        With vsVariation
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpData, i, col����ѡ��) = 1 Then
                    mvItem.����ԭ�� = .RowData(i)
                    str����ԭ�� = Mid(.TextMatrix(i, col����ԭ��), InStr(.TextMatrix(i, col����ԭ��), "-") + 1)
                    Exit For
                End If
            Next
            If i > .Rows - 1 Then
                MsgBox "��ѡ��һ�ֱ���ԭ��", vbInformation, gstrSysName
                If vsVariation.Enabled Then vsVariation.SetFocus
                Exit Sub
            End If
        End With
    End If

    '�������ԭ����������Ҫ�������д����˵��
    If str����ԭ�� = "����" Or str����ԭ�� = "����" Then
        If Trim(txtReason.Text) = "" Then
            MsgBox "����ԭ��Ϊ�����ģ�������д����˵����", vbInformation, gstrSysName
            If txtReason.Enabled Then txtReason.SetFocus
            Exit Sub
        End If
    End If

    If zlCommFun.ActualLen(txtReason.Text) > txtReason.MaxLength Then
        MsgBox "���ԭ����������� " & txtReason.MaxLength \ 2 & " �����ֻ��� " & txtReason.MaxLength & "���ַ���", vbInformation, gstrSysName
        txtReason.SetFocus: Exit Sub
    End If


    '���ҽ��
    If mlngFun = 0 Then
        If optType(0).Value Then
            If mvItem.ҽ��IDs = "" Then
                MsgBox "û�ж��嵱ǰ��Ŀ����Ӧ��ҽ�����ݡ�", vbInformation, gstrSysName
                If cmdAdvice.Enabled Then cmdAdvice.SetFocus
                Exit Sub
            End If
        Else
            If mvItem.ҽ��IDs <> "" Then
                gcnOracle.RollbackTrans
                mvItem.ҽ��IDs = ""
                mdatPathOut = "0:00:00": mlng�׶�ID = 0: mlng���� = 0
                If cboItems.Tag <> mPP.��ǰ�׶�ID & "_" & mPP.��ǰ���� Then Call LoadCurrentItems
            End If
        End If

        '��鲡��
        If optType(1).Value Then
            With vsEPR
                strFilter = ""
                mvItem.����IDs = "": mvItem.�°没��IDs = ""
                For i = .FixedRows To .Rows - 1
                    If .RowData(i) <> 0 Then
                         strTmp = .RowData(i) '��ʽ��(NEW/OLD)|ID/ID:�༭��ʽ
                        If Split(strTmp, "|")(0) = "OLD" Then
                            mvItem.����IDs = mvItem.����IDs & "," & Split(strTmp, "|")(1) '���д��ID:�༭��ʽ
                        Else
                            mvItem.�°没��IDs = mvItem.�°没��IDs & "," & Split(strTmp, "|")(1)
                        End If
                        If InStr(strFilter & ",", "," & .TextMatrix(i, 0) & ",") = 0 Then
                            strFilter = strFilter & "," & .TextMatrix(i, 0)
                        Else
                            MsgBox "ָ�����ظ��Ĳ����ļ�""" & .TextMatrix(i, 0) & """��", vbInformation, gstrSysName
                            .Row = i: Call .ShowCell(.Row, .Col)
                            .SetFocus: Exit Sub
                        End If
                    End If
                Next
                mvItem.����IDs = Mid(mvItem.����IDs, 2)
                mvItem.�°没��IDs = Mid(mvItem.�°没��IDs, 2)
                If mvItem.����IDs = "" And mvItem.�°没��IDs = "" Then
                    MsgBox "��ָ����Ŀ����Ӧ�Ĳ����ļ���", vbInformation, gstrSysName
                    .SetFocus: Exit Sub
                End If
            End With
        Else
            mvItem.����IDs = "": mvItem.�°没��IDs = ""
        End If
        
    End If

    'ִ�н��
    If fraExecute.Visible Then
        With vsResult
            strFilter = ""
            mvItem.��Ŀ��� = ""
            For i = .FixedRows To .Rows - 1
                If Trim(.TextMatrix(i, colִ�н��)) <> "" Then
                    If InStr(strFilter & ",", "," & .TextMatrix(i, colִ�н��) & ",") = 0 _
                       And InStr(strFilter, "," & .TextMatrix(i, colִ�н��) & "|") = 0 Then
                        strFilter = strFilter & "," & .TextMatrix(i, colִ�н��)
                        If .TextMatrix(i, col�������) <> "" Then
                            mrsNature.Filter = "����='" & .TextMatrix(i, col�������) & "'"
                            strFilter = strFilter & "|" & mrsNature!����
                        End If
                    Else
                        MsgBox "ָ�����ظ���ִ�н��""" & .TextMatrix(i, colִ�н��) & """��", vbInformation, gstrSysName
                        .Row = i: Call .ShowCell(.Row, .Col)
                        .SetFocus: Exit Sub
                    End If

                    'ȱʡ���
                    If Val(.TextMatrix(i, colȱʡ���)) <> 0 Then
                        mvItem.��Ŀ��� = .TextMatrix(i, colִ�н��)
                    End If
                End If
            Next
            strFilter = Mid(strFilter, 2)
            If strFilter = "" Then
                MsgBox "��ָ����Ŀ����Ӧ��ִ�н����", vbInformation, gstrSysName
                .SetFocus: Exit Sub
            End If
            If mvItem.��Ŀ��� = "" Then
                MsgBox "��ָ����Ŀ����Ӧ��ȱʡ�����", vbInformation, gstrSysName
                .SetFocus: Exit Sub
            End If
            mvItem.��Ŀ��� = strFilter & vbTab & mvItem.��Ŀ���
        End With
    End If


    '���������ռ�
    mvItem.��Ŀ���� = txtItem.Text
    If fraExecute.Visible Then
        mvItem.ִ���� = IIf(optExecutor(0).Value, 1, 2)
    Else
        mvItem.ִ���� = 0
    End If

    Call SaveData
      
    mblnOK = True
    Unload Me
End Sub

Private Sub SaveData()
'����:bln������� -�Ƿ��޸��������
    Dim i As Long, blnTrans As Boolean
    Dim arrSQL As Variant, arrFile As Variant, arrItem As Variant
    Dim strAddDate As String, strDate As String
    Dim lng����ID As Long, str���˲���IDs As String, str����IDs As String
    Dim lngPosition As Long '����λ��id
    Dim str����IDs As String
    Dim strTmp As String, strSql As String
    
    Dim colNewDoc As New Collection  '����°没������ID
    Dim rsTmp As ADODB.Recordset
    Dim lng�׶�ID As Long
    Dim lng���� As Long
    Dim lng��� As Long, lng�Զ�ִ�� As Long
    
    lngPosition = cboItems.ItemData(cboItems.ListIndex)
    
    On Error GoTo errH
    If mlngFun = 1 Or mlngFun = 2 Then
        If lngPosition = mlngִ��ID Then lngPosition = 0    '����λ�ò������䶯
        strSql = "Zl_����·������_Update(" & mlngִ��ID & ",'" & cboItemType.Text & "','" & mvItem.��Ŀ���� & _
                 "'," & mvItem.ִ���� & ",'" & mvItem.��Ŀ��� & "'," & ZVal(mvItem.ͼ��ID) & ",'" & txtReason.Text & "','" & mvItem.����ԭ�� & _
                 "'," & lngPosition & "," & mPP.����·��ID & "," & mPP.��ǰ�׶�ID & "," & mPP.��ǰ���� & ")"
        Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
        
        If mlngFun = 1 And mdatPathOut <> "0:00:00" And mdatPathOut < mPP.��ǰ���� Then
            strSql = "Select �׶�id, ����,����" & vbNewLine & _
                    "From ����·��ִ�� A" & vbNewLine & _
                    "Where ID = [1] And Exists" & vbNewLine & _
                    " (Select 1 From ����·������ B Where b.·����¼id = a.·����¼id And b.�׶�id = a.�׶�id And b.���� = a.����)"

            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngִ��ID)
            If rsTmp.RecordCount = 1 Then
                strDate = "To_Date('" & Format(rsTmp!����, "yyyy-MM-dd") & "','YYYY-MM-DD')"
                strSql = "Zl_����·������_Insert(" & 2 & "," & mPP.����·��ID & "," & rsTmp!�׶�id & _
                "," & strDate & "," & rsTmp!���� & ",NULL," & 1 & ",NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,1,NULL,0,NULL,1)"
                Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
            End If
        End If
    Else
        strAddDate = "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
        arrFile = Array()
        If mvItem.����IDs <> "" Then
            arrFile = Split(mvItem.����IDs, ",")
            For i = 0 To UBound(arrFile)
                str����IDs = str����IDs & "," & Split(arrFile(i), ":")(0)

                lng����ID = zlDatabase.GetNextID("���Ӳ�����¼")
                str���˲���IDs = str���˲���IDs & "," & lng����ID
            Next
            str����IDs = Mid(str����IDs, 2)
            str���˲���IDs = Mid(str���˲���IDs, 2)
        End If
        
        If mvItem.�°没��IDs <> "" And Not gobjEmr Is Nothing Then '�°�
            If Not gobjEmr.IsInited Or gobjEmr.IsOffline Then Set gobjEmr = Nothing
            If Not gobjEmr Is Nothing Then
                strTmp = "": str����IDs = ""
                For i = 0 To UBound(Split(mvItem.�°没��IDs, ","))
                    strTmp = "<parameter><antetypeid>" & Split(mvItem.�°没��IDs, ",")(i) & "</antetypeid><patient>" & mPati.����ID & "</patient></parameter>"
                    '��¼�������ֶΣ�ԭ��ID,����ID,����ʱ��,��ʼʱ��,��ֹʱ�䣻
                    On Error Resume Next
                    Set rsTmp = gobjEmr.MakeBeforTask(strTmp)
                    Err.Clear: On Error GoTo 0
                    If rsTmp.State <> adStateClosed Then
                        If rsTmp.RecordCount = 1 Then
                            str����IDs = str����IDs & "," & rsTmp!����ID
                        End If
                    End If
                Next
                str����IDs = Mid(str����IDs, 2)
                colNewDoc.Add str����IDs, "C" & (colNewDoc.count + 1) '��¼���ص�����ID,���������ύʧ��,ɾ�����ɳɹ����°没��
            End If
        End If
        
        If mvItem.ҽ��IDs <> "" Then
            If mdatPathOut < mPP.��ǰ���� Then
                '��ȡ����ҽ����ĿӦ������Ľ׶�ID,����,����(��¼)
                strDate = "To_Date('" & Format(mdatPathOut, "yyyy-MM-dd") & "','YYYY-MM-DD')"
                lng�׶�ID = mlng�׶�ID
                lng���� = mlng����
                lng��� = 1 '��¼���
                lng�Զ�ִ�� = 1 '�Զ�ִ��
            ElseIf mdatPathOut > mPP.��ǰ���� Then
                lng��� = 2 '�ݴ���
            End If
        End If
        
        If lng��� = 0 Or lng��� = 2 Then 'ҽ����ʼ����=��ǰ����\>��ǰ����
            strDate = "To_Date('" & Format(mPP.��ǰ����, "yyyy-MM-dd") & "','YYYY-MM-DD')"
            lng�׶�ID = mPP.��ǰ�׶�ID
            lng���� = mPP.��ǰ����
        End If
        
        arrSQL = Array()
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_����·������_Insert(1," & mPati.����ID & "," & mPati.��ҳID & ",'0'," & mPati.����ID & "," & _
                 mPP.����·��ID & "," & lng�׶�ID & _
                 "," & strDate & "," & lng���� & _
                 ",'" & cboItemType.Text & "',Null,'" & mvItem.ҽ��IDs & "','" & str����IDs & "','" & str���˲���IDs & "'" & _
                 ",'" & UserInfo.���� & "'," & strAddDate & ",'" & mvItem.��Ŀ���� & _
                 "'," & mvItem.ִ���� & ",'" & mvItem.��Ŀ��� & "'," & ZVal(mvItem.ͼ��ID) & ",'" & txtReason.Text & "','" & mvItem.����ԭ�� & _
                 "'," & IIf(lng�Զ�ִ�� = 0, "NULL", "1") & ",Null,Null,Null,Null," & lngPosition & "," & IIf(mint���� = 0, 1, 2) & ",'" & str����IDs & IIf(lng��� = 0, "')", "'," & lng��� & ")")
       
        If mdatPathOut <> "0:00:00" And mdatPathOut < mPP.��ǰ���� Then
         '�޸���������ͱ���ԭ��
            If CheckPathIsEvaluated(mPP.����·��ID, lng�׶�ID, mdatPathOut) Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_����·������_Insert(" & 2 & "," & mPP.����·��ID & "," & lng�׶�ID & _
                "," & strDate & "," & lng���� & ",NULL," & 1 & ",NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,1,NULL,0,NULL,1)"
            End If
        End If
        
        If mvItem.ҽ��IDs = "" Then
            gcnOracle.BeginTrans: blnTrans = True
        Else
            'ҽ���Ѳ��벢��ʼ������
            '1.ҽ��У�ԵĲ������SQL
            blnTrans = True
            For i = 0 To UBound(marrSQL)
                zlDatabase.ExecuteProcedure CStr(marrSQL(i)), Me.Caption
            Next
        End If

        '2.��������·������
        For i = 0 To UBound(arrSQL)
            Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
        Next

        '3.���������ļ�RTF����
        If mvItem.����IDs <> "" Then
            For i = 0 To UBound(arrFile)
                arrItem = Split(arrFile(i), ":")
                If arrItem(1) = 0 Or arrItem(1) = 1 Then     'ȫ�ı༭��ʽ�Ĳ���
                    lng����ID = Split(str���˲���IDs, ",")(i)
                    Call ReadRTFData(lng����ID, edtEditor)
                    Call SaveRTFData(lng����ID, mPati.����ID, mPati.��ҳID, 0, edtEditor)   '�ݲ�֧��ѡ��ָ��Ӥ���Ĳ���
                End If
            Next
        End If
        gcnOracle.CommitTrans: blnTrans = False
        
        '����ҽ���¿�����Ϣ
        Call ZLHIS_CIS_001(mclsMipModule, mPati.����ID, mPati.��ҳID, mPati.����ID, mPati.����ID)
    End If


    '���ҽ��ID,�Ա��˳�ʱ�ж��Ƿ��������
    mvItem.ҽ��IDs = ""
    Exit Sub
errH:
    If blnTrans Then
        gcnOracle.RollbackTrans
        '�ع��°没��
        If Not gobjEmr Is Nothing Then
            If Not gobjEmr.IsInited Or gobjEmr.IsOffline Then Set gobjEmr = Nothing
            If Not gobjEmr Is Nothing Then
                For i = 1 To colNewDoc.count
                    strTmp = "<parameter><taskid>" & colNewDoc("C" & i) & "</taskid></parameter>"
                    On Error Resume Next
                    Call gobjEmr.DeleteTask(strTmp)
                    Err.Clear: On Error GoTo 0
                Next
            End If
        End If
    End If
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnOK = False And mlngFun = 1 Then
        MsgBox "��ѡ�����ԭ����ȷ����ť��", vbInformation, gstrSysName
        If vsVariation.Enabled And vsVariation.Visible Then vsVariation.SetFocus
        Cancel = 1: Exit Sub
    End If

    If Not mblnOK And txtItem.Text <> "" Then
        If mlngFun = 2 Then
            If MsgBox("ȷʵҪ����·������Ŀ���޸���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = 1: Exit Sub
            End If
        Else
            If MsgBox("ȷʵҪ����·������Ŀ�������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = 1: Exit Sub
            End If
        End If

    End If

    If Not mrsResult Is Nothing Then
        If mrsResult.State = 1 Then mrsResult.Close
        Set mrsResult = Nothing
    End If
    If Not mrsNature Is Nothing Then
        If mrsNature.State = 1 Then mrsNature.Close
        Set mrsNature = Nothing
    End If
    Set marrSQL = Nothing
    Set mclsMipModule = Nothing
    If mvItem.ҽ��IDs <> "" And mlngFun = 0 Then gcnOracle.RollbackTrans
End Sub

Private Sub cmdAdvice_Click()
'���ܣ��༭��Ŀ����Ӧ��ҽ��
    Dim strItem As String, i As Long
    Dim strAdviceOfItem As String
    Dim arrSQL As Variant
    Dim datPathOut As Date
    Dim rsTmp As ADODB.Recordset
    
    Call gobjKernel.ShowAdviceEdit(mfrmParent, mint����, 2, mPati.����ID, mPati.��ҳID, mvItem.ҽ��IDs, mPP.��ǰ����, arrSQL, strAdviceOfItem, , , , mclsMipModule, mdatPathOut)
        
    If mvItem.ҽ��IDs <> strAdviceOfItem Then
        mvItem.ҽ��IDs = strAdviceOfItem
        marrSQL = arrSQL
    End If
    
    If mdatPathOut <> CDate(0) And mdatPathOut < mPP.��ǰ���� Then
        Set rsTmp = GetPatiPathAppend(mPP.����·��ID, mdatPathOut)   '��ȡҪ����Ľ׶�
        If rsTmp.RecordCount > 0 Then
            mlng�׶�ID = Val(rsTmp!�׶�id & "")
            mlng���� = Val(rsTmp!���� & "")
        End If
        If cboItems.Tag <> mlng�׶�ID & "_" & mlng���� Then Call LoadCurrentItems
    Else
        mlng�׶�ID = 0: mlng���� = 0
        If cboItems.Tag <> mPP.��ǰ�׶�ID & "_" & mPP.��ǰ���� Then Call LoadCurrentItems
    End If
    'ˢ����ʾ
    Call ShowAdvice
End Sub

Private Function ShowAdvice() As Boolean
'���ܣ���ʾ·����Ŀ��Ӧ��ҽ������

    If mvItem.ҽ��IDs = "" Then
        Call UCAdvice.ShowAdvice(2, "", 0, "")
        ShowAdvice = True: Exit Function
    Else
        Call UCAdvice.ShowAdvice(2, "", 0, mvItem.ҽ��IDs)
    End If
    
    If mlngFun <> 2 Then
        'ȱʡ��Ŀ����
        If txtItem.Text = "" Then
            txtItem.Text = UCAdvice.GetAdviceTitle
        End If
    End If
End Function


Private Sub txtReason_GotFocus()
    Call zlControl.TxtSelAll(txtReason)
End Sub

Private Sub txtVariation_GotFocus()
    Call zlControl.TxtSelAll(txtVariation)
End Sub

Private Sub txtVariation_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Dim i As Long, strtxt As String
        strtxt = "*" & UCase(Trim(txtVariation.Text)) & "*"
        With vsVariation
            For i = .FixedRows To .Rows - 1
                If .TextMatrix(i, col����ԭ��) Like strtxt Or .RowData(i) Like strtxt Or .Cell(flexcpData, i, col����ԭ��) Like strtxt Then
                    .SetFocus
                    .Row = i
                    .TopRow = i
                    Exit Sub
                End If
            Next
        End With
    End If
End Sub

Private Sub vsEPR_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call vsEPR_AfterRowColChange(-1, -1, Row, Col)
End Sub

Private Sub vsEPR_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsEPR
        If NewCol = 0 Then
            .ComboList = "..."
            .FocusRect = flexFocusSolid
        Else
            .FocusRect = flexFocusLight
            .ComboList = ""
        End If
    End With
End Sub

Private Sub vsEPR_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnCancel As Boolean
    Dim vPoint As POINTAPI
    
    With vsEPR
        If optEPRType(0).Value Then  '��ʾ�°���Ӳ���δ��װ��װ�����������ϰ����̣�
            strSql = "Select A.ID,Decode(A.����,2,'סԺ����',4,'������',5,'����֤������',6,'֪���ļ�') as ����," & _
                " A.���,A.����,A.˵��,A.���� as �༭��ʽ From �����ļ��б� A" & _
                " Where A.���� IN(2,4,5,6) And Nvl(A.����,0) IN(0,1,2) And A.ͨ�� IN(1,2)" & _
                " Order by A.����,A.���"
        Else
            '�°�����
            If Not gobjEmr.IsInited Or gobjEmr.IsOffline Then
                Set gobjEmr = Nothing
                MsgBox "���Ӳ��������������߻򵼺�̨��¼ʱδ�ܳɹ����ӵ��Ӳ���������!", vbInformation + vbOKOnly, gstrSysName
                Exit Sub
            Else
                'gobjEmr.GetAntetypeList(byref strParameter as string) as Adodb.RecordSet
                '��¼�������ֶΣ������ţ��������ƣ��������ƣ�ID����ţ����ƣ�˵��
                On Error Resume Next
                Set rsTmp = gobjEmr.GetAntetypeList("")
                Err.Clear: On Error GoTo 0
                strSql = Rec.ToSQL(rsTmp)
            End If
        End If
        vPoint = zlControl.GetCoordPos(.Hwnd, .CellLeft + 15, .CellTop)
        Set rsTmp = zlDatabase.ShowSQLMultiSelect(Me, strSql, 0, "�����ļ�", False, "", "", False, False, True, _
            vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True)
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "û�в����ļ����ݿ���ѡ��", vbInformation, gstrSysName
            End If
        Else
            Call SetEPRInput(Row, rsTmp)
            Call EPREnterNextCell(True)
        End If
    End With
End Sub

Private Sub vsEPR_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    
    With vsEPR
        If KeyCode = vbKeyDelete Then
            If .TextMatrix(.Row, 0) <> "" Then
                If MsgBox("ȷʵҪ������в����ļ���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    .RemoveItem .Row
                End If
            End If
        ElseIf KeyCode > 127 Then
            '���ֱ�����뺺�ֵ�����
            Call vsEPR_KeyPress(KeyCode)
        End If
    End With
End Sub

Private Sub vsEPR_KeyPress(KeyAscii As Integer)
    With vsEPR
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call EPREnterNextCell
        ElseIf .Col = 0 Then
            If KeyAscii = Asc("*") Then
                KeyAscii = 0
                Call vsEPR_CellButtonClick(.Row, .Col)
            Else
                .ComboList = "" 'ʹ��ť״̬��������״̬
            End If
        End If
    End With
End Sub

Private Sub vsEPR_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then
        mblnReturn = True
    Else
        mblnReturn = False
    End If
End Sub

Private Sub vsEPR_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsEPR.EditSelStart = 0
    vsEPR.EditSelLength = zlCommFun.ActualLen(vsEPR.EditText)
End Sub

Private Sub vsEPR_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnCancel As Boolean
    Dim strInput As String, vPoint As POINTAPI
    
    With vsEPR
        If Col = 0 Then
            If .EditText = "" Then
                .EditText = .Cell(flexcpData, Row, Col)
                If mblnReturn Then Call EPREnterNextCell
            ElseIf .EditText = .Cell(flexcpData, Row, Col) Then
                If mblnReturn Then Call EPREnterNextCell
            Else
                strInput = UCase(.EditText)
                If optEPRType(0).Value Then
                    strSql = "Select A.ID,Decode(A.����,2,'סԺ����',4,'������',5,'����֤������',6,'֪���ļ�') as ����," & _
                        " A.���,A.����,A.˵��,A.���� as �༭��ʽ From �����ļ��б� A" & _
                        " Where A.���� IN(2,4,5,6) And Nvl(A.����,0) IN(0,1,2)" & _
                        " And A.ͨ�� IN(1,2) And (A.��� Like [1] Or A.���� Like [2] Or zlSpellCode(A.����) Like [2])" & _
                        " Order by A.����,A.���"
                    
                Else
                '�°�����
                    If Not gobjEmr.IsInited Or gobjEmr.IsOffline Then
                        Set gobjEmr = Nothing
                        MsgBox "���Ӳ��������������߻򵼺�̨��¼ʱδ�ܳɹ����ӵ��Ӳ���������!", vbInformation + vbOKOnly, gstrSysName
                        Exit Sub
                    Else
                        'gobjEmr.GetAntetypeList(byref strParameter as string) as Adodb.RecordSet
                        '��¼�������ֶΣ������ţ��������ƣ��������ƣ�ID����ţ����ƣ�˵��
                        On Error Resume Next
                        Set rsTmp = gobjEmr.GetAntetypeList("")
                        Err.Clear: On Error GoTo 0
                        strSql = Rec.ToSQL(rsTmp)
                        strSql = "select A.* from (" & strSql & ") A where A.��� Like [1] Or A.���� Like [2] Or zlSpellCode(A.����) Like [2]  order by ������,���"
                    End If
                End If
                vPoint = zlControl.GetCoordPos(.Hwnd, .CellLeft + 15, .CellTop)
                Set rsTmp = zlDatabase.ShowSQLMultiSelect(Me, strSql, 0, "�����ļ�", False, "", "", False, _
                    False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                    strInput & "%", gstrLike & strInput & "%")
                If rsTmp Is Nothing Then
                    If Not blnCancel Then
                        MsgBox "û���ҵ�ƥ��Ĳ����ļ���", vbInformation, gstrSysName
                    End If
                    Cancel = True
                Else
                    Call SetEPRInput(Row, rsTmp): .EditText = .Text
                    If mblnReturn Then Call EPREnterNextCell(True)
                End If
            End If
            mblnReturn = False
        End If
    End With
End Sub

Private Sub SetEPRInput(ByVal lngRow As Long, rsInput As ADODB.Recordset)
'���ܣ��������ļ�������
    Dim strItem As String, i As Long
    
    With vsEPR
        For i = 1 To rsInput.RecordCount
            If i > 1 Then
                .AddItem "", lngRow + 1
                lngRow = lngRow + 1
            End If
            If optEPRType(0).Value Then
               '�ɰ�
                .RowData(lngRow) = "OLD|" & Val(rsInput!ID) & ":" & rsInput!�༭��ʽ
            Else
               .RowData(lngRow) = "NEW|" & Val(rsInput!ID)
            End If
            .TextMatrix(lngRow, 0) = rsInput!����
            .Cell(flexcpData, lngRow, 0) = .TextMatrix(lngRow, 0)
            
            strItem = strItem & "��" & rsInput!����
            
            rsInput.MoveNext
        Next
        
        'ȱʡ��Ŀ����
        If txtItem.Text = "" Then txtItem.Text = "��д" & Mid(strItem, 2)
        
        'ʼ�ձ���һ����
        If lngRow = .Rows - 1 Then
            .AddItem "", lngRow + 1
        End If
        
    End With
End Sub

Private Sub EPREnterNextCell(Optional ByVal blnNewRow As Boolean)
    Dim i As Long, j As Long
    
    With vsEPR
        If blnNewRow Then
            .Row = .Rows - 1: .Col = 0
            .ShowCell .Row, .Col
        Else
            If .Row + 1 <= .Rows - 1 Then
                .Row = .Row + 1
                .ShowCell .Row, .Col
            Else
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        End If
    End With
End Sub


Private Sub vsResult_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    
    With vsResult
        If Col = col������� Then
            .TextMatrix(Row, Col) = zlCommFun.GetNeedName(.TextMatrix(Row, Col))
            If .TextMatrix(Row, Col) <> "" Then
                mrsNature.Filter = "����='" & .TextMatrix(Row, Col) & "'"
                Set .Cell(flexcpPicture, .Row, colִ��ͼ��) = imgNature.ListImages(Val(mrsNature!����)).Picture
            Else
                Set .Cell(flexcpPicture, .Row, colִ��ͼ��) = Nothing
            End If
        ElseIf Col = colȱʡ��� Then
            If Val(.TextMatrix(Row, Col)) <> 0 Then
                For i = .FixedRows To .Rows - 1
                    If i <> Row Then .TextMatrix(i, colȱʡ���) = 0
                Next
            End If
        End If
    End With
End Sub

Private Sub vsResult_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsResult
        If Not ResultCellEditable(NewRow, NewCol) Then
            .FocusRect = flexFocusLight
            .ComboList = ""
        ElseIf NewCol = colִ�н�� Then
            .ComboList = "..."
            .FocusRect = flexFocusSolid
        ElseIf NewCol = col������� Then
            .ComboList = .ColData(NewCol)
            .FocusRect = flexFocusSolid
        Else
            .FocusRect = flexFocusLight
            .ComboList = ""
        End If
    End With
End Sub

Private Sub vsResult_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnCancel As Boolean
    Dim vPoint As POINTAPI
    
    With vsResult
        '������Ӳ�ѯ������������˳�򲻶ԣ���Ҫ�ر�����
        strSql = _
            " Select A.���� As ID,A.�ϼ� As �ϼ�id,A.����,A.����,A.����,Nvl(A.ĩ��,0) As ĩ��,B.���� As ����" & _
            " From ·��������� A,·��������� B Where Nvl(A.����,0)=B.����(+)" & _
            " Start With A.�ϼ� Is Null Connect By Prior A.����=A.�ϼ�"
        vPoint = zlControl.GetCoordPos(.Hwnd, .CellLeft + 15, .CellTop)
        Set rsTmp = zlDatabase.ShowSQLMultiSelect(Me, strSql, 2, "�������", True, "", "", False, True, True, _
            vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True)
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "û�г���������ݿ���ѡ��", vbInformation, gstrSysName
            End If
        Else
            Call SetResultInput(Row, rsTmp)
            Call ResultEnterNextCell
        End If
    End With
End Sub

Private Sub vsResult_ComboDropDown(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    
    If Col = col������� Then
        With vsResult
            If .TextMatrix(Row, Col) <> "" Then
                For i = 0 To .ComboCount - 1
                    If zlCommFun.GetNeedName(.ComboItem(i)) = .TextMatrix(Row, Col) Then
                        .ComboIndex = i: Exit For
                    End If
                Next
            End If
        End With
    End If
End Sub

Private Sub vsResult_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
        
    With vsResult
        If KeyCode = vbKeyDelete Then
            If .Col = col������� Then
                .TextMatrix(.Row, .Col) = ""
                Set .Cell(flexcpPicture, .Row, colִ��ͼ��) = Nothing
            ElseIf .TextMatrix(.Row, colִ�н��) <> "" Then
                If MsgBox("ȷʵҪ������н����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    .RemoveItem .Row
                End If
            End If
        ElseIf KeyCode > 127 Then
            '���ֱ�����뺺�ֵ�����
            Call vsResult_KeyPress(KeyCode)
        End If
    End With
End Sub

Private Sub vsResult_KeyPress(KeyAscii As Integer)
    With vsResult
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call ResultEnterNextCell
        ElseIf KeyAscii = Asc(",") Then
            KeyAscii = 0
        ElseIf .Col = colִ�н�� Then
            If KeyAscii = Asc("*") Then
                KeyAscii = 0
                Call vsResult_CellButtonClick(.Row, .Col)
            Else
                .ComboList = "" 'ʹ��ť״̬��������״̬
            End If
        End If
    End With
End Sub

Private Sub vsResult_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then
        mblnReturn = True
    Else
        mblnReturn = False
    End If
    If KeyAscii = Asc(",") Then KeyAscii = 0
End Sub

Private Sub vsResult_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsResult.EditSelStart = 0
    vsResult.EditSelLength = zlCommFun.ActualLen(vsResult.EditText)
End Sub

Private Sub vsResult_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not ResultCellEditable(Row, Col) Then Cancel = True
End Sub

Private Sub vsResult_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnCancel As Boolean
    Dim strInput As String, vPoint As POINTAPI
    
    With vsResult
        If Col = colִ�н�� Then
            If .EditText = "" Then
                .EditText = .Cell(flexcpData, Row, Col)
                If mblnReturn Then Call ResultEnterNextCell
            ElseIf .EditText = .Cell(flexcpData, Row, Col) Then
                If mblnReturn Then Call ResultEnterNextCell
            Else
                strInput = UCase(.EditText)
                strSql = "Select A.���� as ID,A.����,A.����,A.����,B.���� as ����" & _
                    " From ·��������� A,·��������� B" & _
                    " Where Nvl(A.����,0)=B.����(+) And A.ĩ��=1" & _
                    " And (A.���� Like [1] Or A.���� Like [2] Or A.���� Like [2])" & _
                    " Order by A.����"
                vPoint = zlControl.GetCoordPos(.Hwnd, .CellLeft + 15, .CellTop)
                Set rsTmp = zlDatabase.ShowSQLMultiSelect(Me, strSql, 0, "�������", _
                    False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                    strInput & "%", gstrLike & strInput & "%")
                If blnCancel Then '��ƥ������ʱ,���������봦��,ȡ����ͬ
                    Cancel = True
                Else
                    Call SetResultInput(Row, rsTmp): .EditText = .Text
                    If mblnReturn Then Call ResultEnterNextCell
                End If
            End If
            mblnReturn = False
        ElseIf Col = col������� Then
            If mblnReturn Then Call ResultEnterNextCell
            mblnReturn = False
        End If
    End With
End Sub


Private Sub SetResultInput(ByVal lngRow As Long, rsInput As ADODB.Recordset)
'���ܣ�������Ŀ���������
    Dim i As Long
    
    With vsResult
        If Not rsInput Is Nothing Then
            For i = 1 To rsInput.RecordCount
                If i > 1 Then
                    .AddItem "", lngRow + 1
                    lngRow = lngRow + 1
                End If
                
                .TextMatrix(lngRow, colִ�н��) = rsInput!����
                
                '����������
                If Not IsNull(rsInput!����) Then
                    mrsNature.Filter = "����='" & rsInput!���� & "'"
                    Set .Cell(flexcpPicture, lngRow, colִ��ͼ��) = imgNature.ListImages(Val(mrsNature!����)).Picture
                    .TextMatrix(lngRow, col�������) = rsInput!����
                End If
                
                If i = 1 And lngRow = .FixedRows Then
                    .TextMatrix(lngRow, colȱʡ���) = 1
                    Call vsResult_AfterEdit(lngRow, colȱʡ���)
                End If
                
                rsInput.MoveNext
            Next
        Else
            .TextMatrix(lngRow, colִ�н��) = .EditText
        End If
        .Cell(flexcpData, lngRow, colִ�н��) = .TextMatrix(lngRow, colִ�н��)
                
        'ʼ�ձ���һ����
        If lngRow = .Rows - 1 Then
            .AddItem "", lngRow + 1
        End If
    End With
End Sub

Private Sub ResultEnterNextCell()
    With vsResult
        If .Col + 1 <= .Cols - 1 Then
            .Col = .Col + 1
        ElseIf .Row + 1 <= .Rows - 1 Then
            .Row = .Row + 1: .Col = colִ�н��
            .ShowCell .Row, .Col
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End With
End Sub

Private Function ResultCellEditable(ByVal lngRow As Long, ByVal lngCol As Long) As Boolean
    ResultCellEditable = True
    
    With vsResult
        If lngCol = colִ��ͼ�� Then
            ResultCellEditable = False
        ElseIf lngCol > colִ�н�� And .TextMatrix(lngRow, colִ�н��) = "" Then
            ResultCellEditable = False
        ElseIf lngCol = col������� And .TextMatrix(lngRow, colִ�н��) <> "" Then
            '�ֵ��еĽ�����ʲ��������,�ֹ�����Ĳ�����
            If .TextMatrix(lngRow, col�������) <> "" Then
                mrsResult.Filter = "����='" & .TextMatrix(lngRow, colִ�н��) & "'"
                If Not mrsResult.EOF Then
                    If NVL(mrsResult!����) = .TextMatrix(lngRow, col�������) Then
                        ResultCellEditable = False
                    End If
                End If
            End If
        End If
    End With
End Function

Private Sub InitVariation()
'���ܣ���ȡ�����ر���ԭ������
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim i As Long
    
    strSql = "Select b.���� As ����, a.����, a.����, a.����" & vbNewLine & _
            "From ���쳣��ԭ�� A, ���쳣��ԭ�� B" & vbNewLine & _
            "Where a.ĩ�� = 1 And a.�ϼ� = b.���� and a.����=1" & vbNewLine & _
            "order by b.����"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
        
    With vsVariation
        .Redraw = flexRDNone
        .Rows = .FixedRows
        If rsTmp.RecordCount > 0 Then
            .MergeCol(col�������) = True
            .Rows = .FixedRows + rsTmp.RecordCount
            'ȱʡ��ѡ��
            Set .Cell(flexcpPicture, .FixedRows, col����ѡ��, .Rows - 1, col����ѡ��) = imgNature.ListImages("UnSelected").Picture
            .Cell(flexcpPictureAlignment, .FixedRows, col����ѡ��, .Rows - 1, col����ѡ��) = flexPicAlignCenterCenter
            For i = .FixedRows To rsTmp.RecordCount
                .Cell(flexcpData, i, col����ѡ��) = 0
                
                .RowData(i) = CStr(rsTmp!����)    '����
                .TextMatrix(i, col�������) = rsTmp!����
                .TextMatrix(i, col����ԭ��) = rsTmp!���� & "-" & rsTmp!����
                .Cell(flexcpData, i, col����ԭ��) = "" & rsTmp!����
                rsTmp.MoveNext
            Next
        End If
        .Redraw = flexRDDirect
    End With
    Exit Sub
errH:
    vsVariation.Redraw = flexRDDirect
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsVariation_DblClick()
    With vsVariation
        If .Row >= .FixedRows Then
            Dim i As Long
            .Redraw = flexRDNone
            If .Cell(flexcpData, .Row, col����ѡ��) = 0 Then
                Set .Cell(flexcpPicture, .Row, col����ѡ��) = imgNature.ListImages("Selected").Picture
                .Cell(flexcpData, .Row, col����ѡ��) = 1
                For i = .FixedRows To .Rows - 1
                    If i <> .Row Then
                        If .Cell(flexcpData, i, col����ѡ��) = 1 Then
                            Set .Cell(flexcpPicture, i, col����ѡ��) = imgNature.ListImages("UnSelected").Picture
                            .Cell(flexcpData, i, col����ѡ��) = 0
                        End If
                    End If
                Next
            Else
                Set .Cell(flexcpPicture, .Row, col����ѡ��) = imgNature.ListImages("UnSelected").Picture
                .Cell(flexcpData, .Row, col����ѡ��) = 0
            End If
            .Redraw = flexRDDirect
        End If
    End With
End Sub

Private Sub vsVariation_GotFocus()
    If vsVariation.Row < vsVariation.FixedRows And vsVariation.Rows > vsVariation.FixedRows Then vsVariation.Row = vsVariation.FixedRows
End Sub

Private Sub vsVariation_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        Call vsVariation_DblClick
    End If
End Sub

Private Sub LoadCurrentItems()
'����:����·����Ŀ
    Dim strSql      As String
    Dim rsData      As Recordset
    Dim strTmp      As String
    Dim lng�׶�ID   As Long
    Dim lng����     As Long
    
    cboItems.Clear
    strSql = "Select Decode(a.��Ŀ����, Null, b.��Ŀ����, a.��Ŀ����) As ��Ŀ����, a.Id, a.��Ŀ���" & vbNewLine & _
             "From ����·��ִ�� A, �ٴ�·����Ŀ B" & vbNewLine & _
             "Where a.·����¼id = [1] And a.�׶�id = [2] And a.��Ŀid = b.Id(+) And a.���� = [3] And a.���� = [4]" & vbNewLine & _
             "Order By a.��Ŀ���"
    strTmp = cboItemType.List(cboItemType.ListIndex)
    On Error GoTo errH
    If mlng�׶�ID <> 0 And mlng���� <> 0 Then 'ָ��ĳ���׶�
        lng�׶�ID = mlng�׶�ID
        lng���� = mlng����
        cboItems.Tag = lng�׶�ID & "_" & lng����   '
    Else
        lng�׶�ID = mPP.��ǰ�׶�ID
        lng���� = mPP.��ǰ����
        cboItems.Tag = lng�׶�ID & "_" & lng����
    End If
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID, lng�׶�ID, lng����, strTmp)
    
    Call Cbo.AddData(cboItems, rsData)

    If mlngFun = 0 Or mlngFun = 1 Then  '0-ֱ�����,1-ҽ���¿�ʱ���,2-�޸�·������Ŀ
        rsData.Filter = "��Ŀ����='·������Ŀ'"
        If rsData.RecordCount < 1 Then
            cboItems.AddItem "·������Ŀ"
            cboItems.ItemData(cboItems.ListCount - 1) = 0
        End If
        If cboItems.ListCount > 0 Then Call Cbo.Locate(cboItems, "·������Ŀ", False)
    ElseIf mlngFun = 2 Then
        If Not Cbo.Locate(cboItems, mlngִ��ID, True) Then
            If cboItems.ListCount = 0 Then
                cboItems.AddItem ""
                cboItems.ItemData(cboItems.ListCount - 1) = 0
            End If
            If cboItems.ListCount > 0 Then cboItems.ListIndex = cboItems.ListCount - 1
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function GetPatiPathAppend(ByVal lng��¼ID As Long, ByVal dat���� As Date) As ADODB.Recordset
'����:�������ڻ�ȡ��Ч�׶κ�����
    Dim strSql      As String
    
    strSql = "Select �׶�id,����" & vbNewLine & _
            "From (Select a.�׶�id, a.����, a.�Ǽ�ʱ��" & vbNewLine & _
            "       From ����·��ִ�� A" & vbNewLine & _
            "       Where a.·����¼id = [1] And a.���� = [2]" & vbNewLine & _
            "       Order By a.�Ǽ�ʱ�� Desc)" & vbNewLine & _
            "Where Rownum < 2"
        
    On Error GoTo errH
    Set GetPatiPathAppend = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng��¼ID, dat����)
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckPathIsEvaluated(ByVal lng·����¼ID As Long, ByVal lng�׶�ID As Long, ByVal dat���� As Date) As Boolean
'-------------------------------------------------------------------------------------------------------------------------------------
'���ܣ���鵱ǰ���˵�ǰ·������,ĳһ��,ĳһ�׶��Ƿ�����
'������
'   lng·����¼Id-����·����¼ID
'   lng�׶�ID  -��ǰ�׶�ID
'   dat����-��ǰ����
'����:True-����;False-δ����
'-------------------------------------------------------------------------------------------------------------------------------------
     Dim strSql As String, rsTmp As Recordset
     
     strSql = "Select 1 From ����·������ Where ·����¼ID=[1] And �׶�ID =[2] And ���� =[3] "
     
     On Error GoTo errH
     Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng·����¼ID, lng�׶�ID, dat����)
     
     CheckPathIsEvaluated = rsTmp.RecordCount > 0
     Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
