VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelector 
   Caption         =   "ҩƷѡ����"
   ClientHeight    =   6465
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9975
   Icon            =   "frmSelector.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6465
   ScaleWidth      =   9975
   Begin VB.CheckBox chkView 
      Caption         =   "��ʾͣ��ҩƷ(&V)"
      Height          =   255
      Left            =   5640
      TabIndex        =   29
      Top             =   143
      Width           =   1650
   End
   Begin MSComctlLib.ImageList imgsDrug 
      Left            =   2160
      Top             =   480
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
            Picture         =   "frmSelector.frx":058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelector.frx":180C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelector.frx":1DA6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "ˢ��(F5)"
      Height          =   350
      Left            =   8280
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "F5���ܼ�ˢ�»���"
      Top             =   90
      Width           =   975
   End
   Begin VB.CheckBox chkContinue 
      Caption         =   "����ѡ��(&M)"
      Height          =   180
      Left            =   6840
      TabIndex        =   4
      Top             =   180
      Width           =   1335
   End
   Begin VB.TextBox txtFilterFind 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   2400
      TabIndex        =   3
      Top             =   120
      Width           =   3135
   End
   Begin VB.OptionButton optFilterFind 
      Caption         =   "����(&I)"
      Height          =   180
      Index           =   1
      Left            =   1200
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   160
      Width           =   950
   End
   Begin VB.OptionButton optFilterFind 
      Caption         =   "����(&F)"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   160
      Value           =   -1  'True
      Width           =   950
   End
   Begin MSComctlLib.ImageList imgsMain 
      Left            =   9000
      Top             =   120
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
            Picture         =   "frmSelector.frx":2340
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelector.frx":2692
            Key             =   "Up"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picSplit02_S 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   40
      Left            =   2880
      ScaleHeight     =   45
      ScaleWidth      =   2535
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   4080
      Width           =   2535
   End
   Begin VB.PictureBox picѡ���� 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   3120
      ScaleHeight     =   1455
      ScaleWidth      =   4815
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4560
      Width           =   4815
      Begin VB.PictureBox picOK 
         BackColor       =   &H00FFEDDD&
         BorderStyle     =   0  'None
         Height          =   220
         Left            =   3240
         Picture         =   "frmSelector.frx":29E4
         ScaleHeight     =   225
         ScaleWidth      =   270
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "ѡ��"
         Top             =   0
         Width           =   270
      End
      Begin VB.PictureBox picUpDown01 
         BackColor       =   &H00FFEDDD&
         BorderStyle     =   0  'None
         Height          =   220
         Left            =   3600
         Picture         =   "frmSelector.frx":2D26
         ScaleHeight     =   225
         ScaleWidth      =   270
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   0
         Width           =   270
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfѡ�� 
         Height          =   1125
         Left            =   0
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   360
         Width           =   4275
         _cx             =   7541
         _cy             =   1984
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   15724527
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
         Cols            =   1
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmSelector.frx":3068
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
      Begin VB.Label lblѡ�� 
         BackColor       =   &H00FFEDDD&
         Caption         =   "ѡ��ҩƷ"
         Height          =   255
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   3885
      End
   End
   Begin VB.PictureBox picҩƷ�� 
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   3120
      ScaleHeight     =   3375
      ScaleWidth      =   4695
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   600
      Width           =   4695
      Begin VB.PictureBox picSetCols 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   220
         Index           =   0
         Left            =   120
         Picture         =   "frmSelector.frx":30DD
         ScaleHeight     =   225
         ScaleWidth      =   255
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox picSplit04_S 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   40
         Left            =   2040
         MousePointer    =   7  'Size N S
         ScaleHeight     =   45
         ScaleWidth      =   2535
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1560
         Width           =   2535
      End
      Begin VSFlex8Ctl.VSFlexGrid vsf���� 
         Height          =   1485
         Left            =   0
         TabIndex        =   9
         Top             =   1800
         Width           =   4275
         _cx             =   7541
         _cy             =   2619
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   15724527
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
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmSelector.frx":360F
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
         Begin VB.PictureBox picSetCols 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   220
            Index           =   1
            Left            =   0
            Picture         =   "frmSelector.frx":3684
            ScaleHeight     =   225
            ScaleWidth      =   255
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   0
            Width           =   255
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsf��� 
         Height          =   1485
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   4275
         _cx             =   7541
         _cy             =   2619
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   15724527
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
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmSelector.frx":3BB6
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
   End
   Begin VB.PictureBox picSplit01_S 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4215
      Left            =   2760
      MousePointer    =   9  'Size W E
      ScaleHeight     =   4215
      ScaleWidth      =   45
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1200
      Width           =   40
   End
   Begin VB.PictureBox pic������ 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5775
      Left            =   0
      ScaleHeight     =   5775
      ScaleWidth      =   2655
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   600
      Width           =   2655
      Begin VB.Frame fra��̬ 
         Caption         =   "��ҩ��̬"
         Height          =   495
         Left            =   0
         TabIndex        =   25
         Top             =   3120
         Width           =   2655
         Begin VB.CheckBox chk��̬ 
            BackColor       =   &H00FFEDDD&
            Caption         =   "ɢװ"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   28
            Top             =   200
            Width           =   700
         End
         Begin VB.CheckBox chk��̬ 
            BackColor       =   &H00FFEDDD&
            Caption         =   "��Ƭ"
            Height          =   180
            Index           =   1
            Left            =   960
            TabIndex        =   27
            Top             =   200
            Width           =   700
         End
         Begin VB.CheckBox chk��̬ 
            BackColor       =   &H00FFEDDD&
            Caption         =   "����"
            Height          =   180
            Index           =   2
            Left            =   1710
            TabIndex        =   26
            Top             =   200
            Width           =   900
         End
      End
      Begin VB.CheckBox chkChoose 
         BackColor       =   &H00FFEDDD&
         Caption         =   "ȫѡ"
         Height          =   180
         Left            =   1560
         TabIndex        =   22
         Top             =   3600
         Width           =   700
      End
      Begin VB.PictureBox picSplit03_S 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   40
         Left            =   0
         MousePointer    =   7  'Size N S
         ScaleHeight     =   45
         ScaleWidth      =   2535
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   3000
         Width           =   2535
      End
      Begin MSComctlLib.ListView lvw���� 
         Height          =   1995
         Left            =   0
         TabIndex        =   7
         Top             =   3840
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   3519
         View            =   1
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         _Version        =   393217
         Icons           =   "imgsDrug"
         SmallIcons      =   "imgsDrug"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "����"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.TreeView tvw��� 
         Height          =   2925
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   5159
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   494
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "imgsDrug"
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.Label lbl���� 
         BackColor       =   &H00FFEDDD&
         Caption         =   "����"
         Height          =   255
         Left            =   0
         TabIndex        =   12
         Top             =   3600
         Width           =   2565
      End
   End
   Begin VB.Label lblFilter 
      AutoSize        =   -1  'True
      Caption         =   "����(&F)"
      Height          =   180
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   630
   End
End
Attribute VB_Name = "frmSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'��ʽ�� "����,,3,1000,r|..."
'   Ԫ��1��Keyֵ��
'   Ԫ��2��Captionֵ��Ĭ��ΪKeyֵ����
'   Ԫ��3�������ԣ�0���ڲ���ʾ�����ƶ���1���ڲ����أ������ƶ���������ʾ��2���û����أ�3���û���ʾ(Ĭ��ֵ)��
'   Ԫ��4���п�ȣ�Ĭ��0����
'   Ԫ��5����ʾ��ʽ��s���ַ����� n�����֣� d�����ڣ� t��ʱ�䣻 dt������ʱ��
Private Enum enmColProperty
    cpKey = 0
    cpCaption
    cpDisplay
    cpWidth
    cpFormat
End Enum

Private Const MCON_��� = _
    "����,,,1000|��ҩ��̬,,,1000,s|ҩ������,,1,0|��Դ,,,1000|����ҩ��,,,1000|ҩ��ID,,1,0|���ñ�־,,1,0|��;����ID,,1,0|������λ,,1,1000" & _
    "|ҩƷ����,,,1000|ͨ������,,,1000|ҩƷ����,,1,1000|��Ʒ��,,,1000|���,,0,1000|������,,,1000|ԭ����,,,1000|ҩ��ID,,1,0" & _
    "|ҩƷID,,1,0|�ϴβɹ���,,1,1000,n|�ۼ�,,,1000,n|�ۼ۵�λ,,0,1000|�ۼ۰�װ,,0,1000,n|���ﵥλ,,0,1000" & _
    "|�����װ,,0,1000,n|סԺ��λ,,0,1000|סԺ��װ,,0,1000,n|ҩ�ⵥλ,,0,1000|ҩ���װ,,0,1000,n|��������,,,1000,n" & _
    "|�������,,1,1000,n|�����,,1,1000,n|�����,,1,1000,n|��Ч��,,,1000,n|ҩ�����,,,1000|ҩ������,,,1000" & _
    "|ʱ��,,,1000|ָ��������,,1,1000,n|�ӳ���,,1,1000,n|�ⷿ��λ,,,1000|��׼�ĺ�,,,1000|ʵ������,,1,1000,n" & _
    "|��ͬ��λ,,0,1000|ҩ�ۼ���,,,1000|��������,,1,,n|����,,1,0|���ּ���,,1,0|�����,,1,0"

Private Const MCON_���� = _
    "RID,,1,0|�ⷿ,,,1000|����,,1,1000|�������,,0,1000,d|����,,,1000|��������,,,1000,d|��Ч��,,,1000,d|������,,,1000|ԭ����,,,1000" & _
    "|�ɱ���,,,1000,n|�ۼ�,,,1000,n|��������,,,1000,n|�������,,,1000,n|�����,,,1000,n|�����,,,1000,n" & _
    "|�ϴι�Ӧ��ID,,1,0|ʵ������,,1,0,n|��׼�ĺ�,,,1000|��Ӧ��,,,1000"

Private Const MCON_ѡ�� = _
    "����,,1,1000|ҩ������,,1,0|��Դ,,1,1000|����ҩ��,,1,1000|ҩƷ����,,0,1000|ͨ������,,,1000|ҩ��ID,,1,0|��;����ID,,1,0" & _
    "|������λ,,1,1000|ҩƷ����,,1,1000|��Ʒ��,,,1000|���,,0,1000|������,,0,1000|ԭ����,,0,1000|ҩ��ID,,1,0|ҩƷID,,1,0|����,,0,1000" & _
    "|�ϴβɹ���,,1,1000,n|�ۼ�,,0,1000,n|�ۼ۵�λ,,1,1000|�ۼ۰�װ,,1,1000,n|���ﵥλ,,1,1000" & _
    "|�����װ,,1,1000,n|סԺ��λ,,1,1000|סԺ��װ,,1,1000,n|ҩ�ⵥλ,,1,1000|ҩ���װ,,1,1000,n|��������,,1,1000,n" & _
    "|�������,,1,1000,n|�����,,1,1000,n|�����,,1,1000,n|���Ч��,,1,1000,n|ҩ�����,,1,1000|ҩ������,,1,1000" & _
    "|ʱ��,,1,1000|ָ��������,,1,1000,n|�ӳ���,,1,1000,n|��׼�ĺ�,,1,1000" & _
    "|��ͬ��λ,,1,1000|ҩ�ۼ���,,1,1000|��������,,1,,n" & _
    "|����,,1,1000|��������,,1,1000,d|��Ч��,��Ч����,1,1000,d|ʵ������,,1,1000,n" & _
    "|�ϴι�Ӧ��ID,,1,0|�ɱ���,,1,1000,n"

Private Const mconint�ۼ۵�λ As Integer = 1
Private Const mconint���ﵥλ As Integer = 2
Private Const mconintסԺ��λ As Integer = 3
Private Const mconintҩ�ⵥλ As Integer = 4

Private mintUnit As Integer             '��λϵ����1-�ۼ�;2-����;3-סԺ;4-ҩ��
Private mintCostDigit As Integer        '�ɱ���С��λ��
Private mintPriceDigit As Integer       '�ۼ�С��λ��
Private mintNumberDigit As Integer      '����С��λ��
Private mintMoneyDigit As Integer       '���С��λ��

Private mstrCostFormat As String
Private mstrPriceFormat As String
Private mstrNumberFormat As String
Private mstrMoneyFormat As String

Private mStr�ɱ��� As String
Private mStr���� As String
Private mStr���� As String
Private mStr��� As String

Private Type WinLocate
    Left As Double
    Top As Double
End Type
Private WindowPosition As WinLocate         '����λ��

Private MStrCaption As String
Private mbytStyle As Long                   'ѡ������ʾģʽ��   0������ѡ��ģʽ�� 1��ģ��¼��ģʽ
'Private mlngSys As Long                     'ϵͳ��
'Private mlngMode As Long                    'ģ���
Private mstrPrivs As String                 '��ǰ����ģ��Ȩ��
Private mfrmMain As Form                    '������
Private mbyt�༭ģʽ As Byte                '1����⣻ 2������
Private mstr���� As String                  '��ѯ¼��ļ���
Private mlng��Դ�ⷿ As Long                '��Դ�ⷿID
Private mlngĿ��ⷿ As Long                'Ŀ��ⷿID
Private mlngʹ�ò��� As Long                'ʹ�ò���ID
Private mlng��Ӧ�� As Long                  '��Ӧ��ID
Private mbyt����ͣ��ҩƷ As Byte         '�Ƿ���ʾͣ��ҩƷ��0-����ʾͣ��ҩƷ��1-��ʾͣ��ҩƷ��2-����ע��������ȷ����
Private mbln��ҩ�ⷿ As Boolean             '�Ƿ���ҩ�ⷿ   True����  False����
Private mintStockCheck As Integer           '�����       0-����飻1-��飬�������ѣ�2-��飬�����ֹ
Private mbyt�ⷿ���� As Byte                '�ⷿ����       1-ҩ�⣻2-ҩ����3-�Ƽ���
Private mrsReturn As ADODB.Recordset        '����ѡ��ҩƷ����
Private mstrFilterClass As String           'ҩƷ����Ĺ�������
Private mstrMatch As String                 '����ƥ�䷽ʽ   0-˫��ƥ�䣻 1-������ƥ��
'Private mbln��ȷ�������� As Boolean         '��ȷ����ҩƷ����
'Private mbln���쵥 As Boolean               'True���쵥��False�ƿⵥ
Private mblnCheck As Boolean                '�Ƿ�����(�̵���)
Private mblnPrice As Boolean                '�Ƿ�����ʱ�ۻ�����ҩƷ�����
Private mblnStore As Boolean                '��ʾ���
Private mbln������ As Boolean               '�����μ�¼��ʾ
Private mstr�����οⷿ As String            '�����μ�¼��ʾ�Ŀⷿ
Private mblnMultiSel As Boolean             '�ɶ�ѡ��¼
Private mblnOK As Boolean
Private mblnCostView As Boolean             '�鿴�ɱ��� true-����鿴 false-������鿴

Private mstr��� As String, mstr���� As String, mstrѡ�� As String      '�û��Զ������ͷ����˳��
Private mlngLast As Long    '���һ��ѡ��ҩƷ
Private mblnLoad As Boolean     '�����ͣ��жϴ����Ƿ����ڼ��� true-���ڣ�false-�Ѿ��������
Private mint�����γ��� As Integer           '0-�������γ���,1-�����γ���
Public Function ShowMe( _
    ByVal FrmMain As Form, _
    ByVal bytStyle As Byte, _
    ByVal byt�༭ģʽ As Byte, _
    Optional ByVal str���� As String, _
    Optional ByVal lngWinLeft As Long = 0, _
    Optional ByVal lngWinTop As Long = 0, _
    Optional ByVal lng��Դ�ⷿ As Long = 0, _
    Optional ByVal lngĿ��ⷿ As Long = 0, _
    Optional ByVal lngʹ�ò��� As Long = 0, _
    Optional ByVal lng��Ӧ�� As Long = 0, _
    Optional ByVal bln����� As Boolean = True, _
    Optional ByVal bln������λ�ʱ�� As Boolean = True, _
    Optional ByVal bln��ʾ��� As Boolean = True, _
    Optional ByVal byt����ͣ��ҩƷ As Byte = 0, _
    Optional ByVal bln�ɶ�ѡ As Boolean = True, _
    Optional ByVal strPrivs As String = "" _
) As ADODB.Recordset
    Dim strKeyName As String
    
    If grsMaster Is Nothing Or grsSlave Is Nothing Then
        MsgBox "ҩƷѡ����������δ���ɣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    WindowPosition.Left = lngWinLeft
    WindowPosition.Top = lngWinTop
    
    mbytStyle = bytStyle
    Set mfrmMain = FrmMain
'    mlngSys = lngSys
'    mlngMode = lngMode
    mbyt�༭ģʽ = byt�༭ģʽ
    mstr���� = str���� 'VerifyFilterStr(str����)
    mlng��Դ�ⷿ = lng��Դ�ⷿ
    mlngĿ��ⷿ = lngĿ��ⷿ
    mlngʹ�ò��� = lngʹ�ò���
    mlng��Ӧ�� = lng��Ӧ��
    mblnCheck = bln�����
    mblnPrice = bln������λ�ʱ��
    mblnStore = bln��ʾ���
    mbyt����ͣ��ҩƷ = byt����ͣ��ҩƷ
    mblnMultiSel = bln�ɶ�ѡ
    mstrPrivs = strPrivs
    
    '�ָ�����ѡ��״̬
    Select Case UCase(mfrmMain.Name)
        Case UCase("frmTransferCard")
            strKeyName = "ҩƷ�ƿ����"
        Case UCase("frmRequestDrugCard")
            strKeyName = "ҩƷ�������"
        Case UCase("frmDrawCard")
            strKeyName = "ҩƷ���ù���"
    End Select
    
    If UCase(mfrmMain.Name) = UCase("frmRequestDrugCard") Then
        mint�����γ��� = Val(zldatabase.GetPara("ҩƷ�����γ���", glngSys, 1343, 0))
    ElseIf UCase(mfrmMain.Name) = UCase("frmTransferCard") Then
        mint�����γ��� = Val(zldatabase.GetPara("ҩƷ�����γ���", glngSys, 1304, 1))
    ElseIf UCase(mfrmMain.Name) = UCase("frmDrawCard") Then
        mint�����γ��� = Val(zldatabase.GetPara("ҩƷ�����γ���", glngSys, 1305, 1))
    Else
        mint�����γ��� = 1
    End If
    
    mbln������ = False
    '�̵㵥Ҫ��¼
    If UCase(FrmMain.Name) = UCase("frmCheckCard") Or UCase(FrmMain.Name) = UCase("frmCheckCourseCard") Then
        mbln������ = True
        If grsSlave.State = adStateOpen And grsSlave.RecordCount > 0 Then
            grsSlave.MoveFirst
            mstr�����οⷿ = zlStr.Nvl(grsSlave!�ⷿ)
        Else
            mstr�����οⷿ = Get��������(IIf(lng��Դ�ⷿ = 0, lngĿ��ⷿ, lng��Դ�ⷿ))
        End If
    ElseIf UCase(mfrmMain.Name) = UCase("frmTransferCard") Or UCase(mfrmMain.Name) = UCase("frmRequestDrugCard") Or UCase(mfrmMain.Name) = UCase("frmDrawCard") Then
'        chkBatch.Visible = True
    End If
    
    If txtFilterFind.Tag = "1" Then txtFilterFind.SetFocus
    
    Show vbModal, FrmMain
    Set ShowMe = mrsReturn
    Unload Me
End Function

Private Sub chkBatch_Click()
    If vsfѡ��.rows > 1 And mint�����γ��� = 0 Then
        If MsgBox("�Ѿ���ѡ��ҩƷ���ڣ�ȡ���������γ��⡱�������ѡ����ҩƷ����ȷ����" _
            , vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            vsfѡ��.rows = 1
            chkContinue.Tag = ""
            Call Form_Resize
        End If
    End If

    ViewVSF���� vsf���
End Sub

Private Sub chkChoose_Click()
    Dim i As Integer
    If chkChoose.Value = 2 Then Exit Sub
    For i = 1 To lvw����.ListItems.count
        lvw����.ListItems(i).Checked = chkChoose.Value
    Next
    '����ͣ��ҩƷ�Ƿ���ʾ
    myFilter
    SetColor
End Sub

Private Sub chkContinue_Click()
    picѡ����.Visible = chkContinue.Value
    picѡ����.TabStop = chkContinue.Value
    picSplit02_S.Visible = chkContinue.Value
    If chkContinue.Value = 1 And chkContinue.Tag <> "msg" Then
        vsfѡ��.rows = 1
        chkContinue.Tag = ""
        lblѡ��.Caption = "ѡ��ҩƷ"
    Else
        If vsfѡ��.rows > 1 And chkContinue.Tag <> "msg" Then
            If MsgBox("�Ѿ���ѡ��ҩƷ���ڣ�ȡ��������ѡ�񡱽������ѡ����ҩƷ����ȷ����" _
                , vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                vsfѡ��.rows = 1
                chkContinue.Tag = ""
                Call Form_Resize
                Exit Sub
            End If
            chkContinue.Tag = "msg"
            chkContinue.Value = 1
            Exit Sub
        Else
            chkContinue.Tag = ""
        End If
    End If
    Call Form_Resize
End Sub

Private Sub chkView_Click()
    If chkView.Visible = False Then Exit Sub
    
    myFilter
    SetColor
End Sub

Private Sub chk��̬_Click(index As Integer)
    Call myFilter
    SetColor
End Sub

'
'
'
Private Sub cmdRefresh_Click()
    Dim strFind As String
    
    cmdRefresh.Enabled = False
    Me.MousePointer = vbHourglass
    On Error GoTo errHandle
    If mfrmMain.Caption Like "ҩƷ�ƿⵥ*" Then
        Call SetSelectorRS(mbyt�༭ģʽ, "ҩƷ�ƿ����", mlng��Դ�ⷿ, mlngĿ��ⷿ, , mlng��Ӧ��, , False)
    
    ElseIf mfrmMain.Caption Like "ҩƷ�̵��¼��*" Or mfrmMain.Caption Like "ҩƷ�̵��*" Then
        Call SetSelectorRS(mbyt�༭ģʽ, "ҩƷ�̵����", mlng��Դ�ⷿ, mlngĿ��ⷿ, , mlng��Ӧ��, , IIf(mbyt����ͣ��ҩƷ = 1, True, False))
    
    ElseIf mfrmMain.Caption Like "����۵�����*" Then
        Call SetSelectorRS(mbyt�༭ģʽ, "����۵�������", mlng��Դ�ⷿ, mlngĿ��ⷿ, , mlng��Ӧ��, , False)
    
    ElseIf mfrmMain.Caption Like "ҩƷ���õ�*" Then
        Call SetSelectorRS(mbyt�༭ģʽ, "ҩƷ���ù���", mlng��Դ�ⷿ, mlngĿ��ⷿ, , mlng��Ӧ��, , False)
    
    ElseIf mfrmMain.Caption Like "ҩƷ������ⵥ*" Then
        Call SetSelectorRS(mbyt�༭ģʽ, "ҩƷ����������", mlng��Դ�ⷿ, mlngĿ��ⷿ, , mlng��Ӧ��, , False)
    
    ElseIf mfrmMain.Caption Like "ҩƷ�������ⵥ*" Then
        Call SetSelectorRS(mbyt�༭ģʽ, "ҩƷ�����������", mlng��Դ�ⷿ, mlngĿ��ⷿ, , mlng��Ӧ��, , False)
    
    ElseIf mfrmMain.Caption Like "ҩƷ�⹺��ⵥ*" Then
        Call SetSelectorRS(mbyt�༭ģʽ, "ҩƷ�⹺������", mlng��Դ�ⷿ, mlngĿ��ⷿ, , mlng��Ӧ��, , False)
    
    ElseIf mfrmMain.Caption Like "ҩƷ�ƿⵥ*" Then
        Call SetSelectorRS(mbyt�༭ģʽ, "ҩƷ�ƿ����", mlng��Դ�ⷿ, mlngĿ��ⷿ, , mlng��Ӧ��, , False)
    
    ElseIf mfrmMain.Caption Like "ҩƷ���쵥*" Then
        Call SetSelectorRS(mbyt�༭ģʽ, "ҩƷ�������", mlng��Դ�ⷿ, mlngĿ��ⷿ, , mlng��Ӧ��, , False)
    
    ElseIf mfrmMain.Caption Like "ҩƷ�ƻ���*" Then
        Call SetSelectorRS(mbyt�༭ģʽ, "ҩƷ�ƻ�����", mlng��Դ�ⷿ, mlngĿ��ⷿ, , mlng��Ӧ��, , False)
    Else '����
        Call SetSelectorRS(mbyt�༭ģʽ, "����ҩƷ", mlng��Դ�ⷿ, mlngĿ��ⷿ, , mlng��Ӧ��, , True)
    End If
    Me.MousePointer = vbDefault
    '�����������
    If mbytStyle = 0 Then 'ѡ��ģʽ
        mstrFilterClass = GetFilterClass(tvw���.SelectedItem)
        If tvw���.SelectedItem.Children = 0 And tvw���.SelectedItem.Key Like "Root*" Then '����Ǹ��ڵ��Ҹýڵ����������ݵĻ���û������
            mstrFilterClass = "��;����id=99999999999999"
        Else
            mstrFilterClass = Left(mstrFilterClass, Len(mstrFilterClass) - 4)
        End If
        grsMaster.Filter = mstrFilterClass
        Call FillVSF(grsMaster, vsf���)
    Else '¼��ģʽ
        strFind = Trim(txtFilterFind.Text)
        txtFilterFind.Tag = ""
        mstrFilterClass = GetFilterSimpleCode(strFind)
        grsMasterInput.Filter = mstrFilterClass
        Call FillVSF(grsMasterInput, vsf���)
    End If
    ViewVSF���� vsf���
    cmdRefresh.Enabled = True
    
    myFilter
    
    SetColor
    
    Exit Sub
    
errHandle:
    Me.MousePointer = vbDefault
    cmdRefresh.Enabled = True
    Call ErrCenter
End Sub

Private Sub Form_Activate()
    '����¼��ƥ��ֻ��һ������
    If vsf���.rows = 2 And mbytStyle = 1 Then
        If vsf����.Visible = True Then
            If vsf����.rows = 2 Then
                Call vsf���_DblClick
                If mblnOK Then Unload Me
            End If
        Else
            Call vsf���_DblClick
            If mblnOK Then Unload Me
        End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    ElseIf KeyCode = vbKeyF5 Then
        If cmdRefresh.Enabled Then Call cmdRefresh_Click
    End If
End Sub

Private Sub Form_Load()
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo errHandle
    mblnLoad = True
    mblnCostView = zlStr.IsHavePrivs(mstrPrivs, "�鿴�ɱ���")
    picSplit03_S.Top = Me.ScaleHeight - lvw����.Height - lbl����.Height - picSplit03_S.Height
    picSplit04_S.Top = Me.ScaleHeight - vsf����.Height - picSplit04_S.Height - picҩƷ��.Top
    
    If mbytStyle = 1 Then Height = 4000
    
    MStrCaption = GetText(GetParentWindow(mfrmMain.hWnd))
    Call RestoreWinState(Me, App.ProductName, mfrmMain.Caption & mbytStyle)
    
    picSplit02_S.Visible = False
    picSplit04_S.Visible = False
    picѡ����.Visible = False
    vsf����.Visible = False
    vsf����.TabStop = False
    
    optFilterFind(0).Visible = False: optFilterFind(1).Visible = False  '��ȡ�����ؼ��ʹ�����ʱ����
    optFilterFind(0).TabStop = False: optFilterFind(1).TabStop = False
    
    pic������.Visible = mbytStyle = 0
    lblFilter.Visible = mbytStyle = 1
    txtFilterFind.Visible = mbytStyle = 1: txtFilterFind.TabStop = mbytStyle = 1
    
    '����ƥ�䷽ʽ
    If mbytStyle = 1 Then mstrMatch = IIf(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0) = "0", "%", "")
    
    '��ȡ�Ƿ���������
    gstrSQL = "Select Nvl(��鷽ʽ,0) ����� From ҩƷ������ Where �ⷿID=[1]"
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "��ȡ�Ƿ���������", mlng��Դ�ⷿ)
    If Not rsTmp.EOF Then
        mintStockCheck = rsTmp!�����
    End If
    rsTmp.Close
    
    '���Դ�ⷿ�Ƿ�Ϊҩ��
    mbyt�ⷿ���� = GetStockType(mlng��Դ�ⷿ)
        
    '������λ
    If MStrCaption Like "ҩƷ�������*" Then
        Call GetDrugDigit(mlngʹ�ò���, MStrCaption, mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
        mlngʹ�ò��� = 0
    ElseIf MStrCaption Like "ҩƷ�ƿ����*" Then
        Call GetDrugDigit(mlngʹ�ò���, MStrCaption, mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
        mlngʹ�ò��� = 0
    Else
        Call GetDrugDigit(IIf(mlng��Դ�ⷿ = 0, mlngĿ��ⷿ, mlng��Դ�ⷿ), MStrCaption, mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
    End If

    '���ۡ�����ʽ
'    mstrCostFormat = "'999999999990." & String(gtype_UserDrugDigits.Digit_�ɱ���, "0") & "'"
'    mstrPriceFormat = "'999999999990." & String(gtype_UserDrugDigits.Digit_���ۼ�, "0") & "'"
'    mstrNumberFormat = "'999999999990." & String(gtype_UserDrugDigits.Digit_����, "0") & "'"
'    mstrMoneyFormat = "'999999999990." & String(gtype_UserDrugDigits.Digit_���, "0") & "'"
'    mStr�ɱ��� = "####0." & String(gtype_UserDrugDigits.Digit_�ɱ���, "0") & ";-####0." & String(gtype_UserDrugDigits.Digit_�ɱ���, "0") & "; ;"
'    mStr���� = "####0." & String(gtype_UserDrugDigits.Digit_���ۼ�, "0") & ";-####0." & String(gtype_UserDrugDigits.Digit_���ۼ�, "0") & "; ;"
'    mStr���� = "####0." & String(gtype_UserDrugDigits.Digit_����, "0") & ";-####0." & String(gtype_UserDrugDigits.Digit_����, "0") & "; ;"
'    mStr��� = "####0." & String(gtype_UserDrugDigits.Digit_���, "0") & ";-####0." & String(gtype_UserDrugDigits.Digit_���, "0") & "; ;"
    mStr�ɱ��� = "####0." & String(mintCostDigit, "0") & ";-####0." & String(mintCostDigit, "0") & "; ;"
    mStr���� = "####0." & String(mintPriceDigit, "0") & ";-####0." & String(mintPriceDigit, "0") & "; ;"
    mStr���� = "####0." & String(mintNumberDigit, "0") & ";-####0." & String(mintNumberDigit, "0") & "; ;"
    mStr��� = "####0." & String(mintMoneyDigit, "0") & ";-####0." & String(mintMoneyDigit, "0") & "; ;"
    
    'VSF���ͷ�����ظ�����
    mstr��� = Trim(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\" & Me.Name & mfrmMain.Caption & mbytStyle & "\VSFlexGrid", _
        vsf���.Name & vsf���.Tag & "����", ""))
    mstr���� = Trim(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\" & Me.Name & mfrmMain.Caption & mbytStyle & "\VSFlexGrid", _
        vsf����.Name & vsf����.Tag & "����", ""))
    mstrѡ�� = Trim(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\" & Me.Name & mfrmMain.Caption & mbytStyle & "\VSFlexGrid", _
        vsfѡ��.Name & vsfѡ��.Tag & "����", ""))
    
    '����������ʾ
    Call ParamsColsHead
    
    InitVSF vsf���
    InitVSF vsf����
    InitVSF vsfѡ��
    SetVSFHead vsf���, mstr���
    SetVSFHead vsf����, mstr����
    SetVSFHead vsfѡ��, mstrѡ��
    vsfѡ��.rows = 1
    
    'װ������ tvw���lvw����
    If mbytStyle = 0 Then Call Fill_TVW���
    
    If mbytStyle = 1 Then
        'txtFilterFind.Tag = "1"
        txtFilterFind.Text = mstr����
        'txtFilterFind.Tag = ""
        'WindowPosition��ShowMe()�Ѿ���ֵ
    Else
        If Not tvw���.SelectedItem Is Nothing Then
            tvw���_NodeClick tvw���.SelectedItem
        End If
        '��Ļ����
        'Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
        '�����߾���
        WindowPosition.Left = mfrmMain.Left + (mfrmMain.Width - Me.Width) \ 2
        If WindowPosition.Left < 0 Then WindowPosition.Left = 0
        WindowPosition.Top = mfrmMain.Top + (mfrmMain.Height - Me.Height) \ 2
        If WindowPosition.Top < 0 Then WindowPosition.Top = 0
    End If
    Move WindowPosition.Left, WindowPosition.Top
    
    vsf���.TabIndex = 0: vsf����.TabIndex = 1
    
    '��ʼ���������ݼ�����
    Call InitReturnRecord
    
    chkContinue.Visible = mblnMultiSel
    chkView.Visible = mbyt����ͣ��ҩƷ = 2
    
    chkView.Value = GetSetting("ZLSOFT", "˽��ģ��\ZLHIS\zl9MediStore", "��ʾͣ��ҩƷ", 0)
    
    myFilter
    SetColor
    
    '��ʾ����
    If vsf���.rows > 1 Then
        For i = 1 To vsf���.rows - 1
            If vsf���.RowHidden(i) = False Then
                vsf���.Row = i
                Exit For
            End If
        Next
        
        
        ViewVSF���� vsf���
'        If mbln������ Then
'            If vsf����.Rows > 2 Then
'                vsf����.Row = 2
'            ElseIf vsf����.Rows > 1 Then
'                vsf����.Row = 1
'            End If
'        Else
'            If vsf����.Rows > 1 Then vsf����.Row = 1
'        End If
    End If
    
    mblnLoad = False
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub




Private Sub myFilter()
'���ܣ����ݡ���ʾͣ��ҩƷ���������͡��͡���ҩ��̬����������
    Dim i As Integer
    Dim str��̬ As String
    Dim str���� As String
    
     '1���ж�ѡ�����̬
    str��̬ = ";"
    For i = 0 To chk��̬.count - 1
        If chk��̬(i).Value = 1 Then str��̬ = str��̬ & chk��̬(i).Caption & ";"
    Next
    If str��̬ = ";" Then '������ѡ��Ĭ��Ϊȫѡ
        For i = 0 To chk��̬.count - 1
            str��̬ = str��̬ & chk��̬(i).Caption & ";"
        Next
    End If
    '2���ж�ѡ��ļ���
    If chkChoose.Value = 2 Then '����ѡ��
        With lvw����
            str���� = ";"
            For i = 1 To .ListItems.count
                If .ListItems(i).Checked Then str���� = str���� & .ListItems(i).Text & ";"
            Next
        End With
    Else 'ȫѡ
        With lvw����
            str���� = ";"
            For i = 1 To .ListItems.count
                str���� = str���� & .ListItems(i).Text & ";"
            Next
        End With
    End If
    
    '3������

    With vsf���
         .Redraw = flexRDNone
        For i = 1 To .rows - 1
        
            .RowHidden(i) = False 'ÿ������ǰ������ʾ
            
            'ͣ��ҩƷ����ʾ
            If mbyt����ͣ��ҩƷ = 2 Then
                If chkView.Value = 0 Then
                    .RowHidden(i) = .TextMatrix(i, .ColIndex("ͣ��")) = "��"
                End If
            Else
                If mbyt����ͣ��ҩƷ <> 1 Then .RowHidden(i) = .TextMatrix(i, .ColIndex("ͣ��")) = "��"
            End If
                
            'ģ����ѯģʽ�����˼���
            If mbytStyle <> 1 And .RowHidden(i) = False Then .RowHidden(i) = InStr(str����, ";" & .TextMatrix(i, .ColIndex("����")) & ";") = 0 '���͹���
            
            If .RowHidden(i) = False And fra��̬.Visible = True Then 'ֻ����в�ҩ����̬����
                .RowHidden(i) = InStr(str��̬, ";" & .TextMatrix(i, .ColIndex("��ҩ��̬")) & ";") = 0
            End If
 
        Next
        .Redraw = flexRDDirect
    End With
    
End Sub

Private Function GetStockType(ByVal lngStockid As Long) As Byte
'------------------------------------------------------------------------
'���ܣ���ȡҩƷ�ⷿ������
'������
'   lngStockID���ⷿID
'���أ�0��δ�ҵ�  1��ҩ��  2��ҩ��  3���Ƽ���
'------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strsql As String
    
    If lngStockid <= 0 Then Exit Function
    
    On Error GoTo errHandle
    GetStockType = 3
    strsql = "select count(����ID) rec from ��������˵�� where (�������� like '%�Ƽ���' or �������� like '%ҩ��') And ����id=[1] "
    Set rsTmp = zldatabase.OpenSQLRecord(strsql, "��ȡָ����������", lngStockid)
    If rsTmp!Rec > 0 Then
        GetStockType = 2
    Else
        strsql = "select count(����ID) rec from ��������˵�� where �������� like '%ҩ��' And ����id=[1] "
        Set rsTmp = zldatabase.OpenSQLRecord(strsql, "��ȡָ����������", lngStockid)
        If rsTmp!Rec > 0 Then
            GetStockType = 1
        End If
    End If
    rsTmp.Close
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Resize()
    Dim bln��ҩ As Boolean
    
    If WindowState = 1 Then Exit Sub
    
    On Error Resume Next
    If mbytStyle = 1 Then
        If Me.Height < 3000 Then Me.Height = 3000
    Else
        If Me.Height < 5835 Then Me.Height = 5835
    End If
    If Me.Width < 8415 Then Me.Width = 8415
    
    cmdRefresh.Left = ScaleWidth - cmdRefresh.Width - 150
    If mbytStyle = 1 Then
        chkContinue.Left = cmdRefresh.Left - chkContinue.Width - 150
        chkView.Left = cmdRefresh.Left - chkView.Width
    Else
        chkContinue.Left = picSplit01_S.Left + picSplit01_S.Width + 75
        chkView.Left = picSplit01_S.Left + picSplit01_S.Width + 75
    End If

'    If optFilterFind(1).Visible Then
'        txtFilterFind.Left = optFilterFind(1).Left + optFilterFind(1).Width + 50
'        txtFilterFind.Width = ScaleWidth - txtFilterFind.Left - chkContinue.Width - cmdRefresh.Width - 500
'    Else
'        txtFilterFind.Left = optFilterFind(1).Left
'        txtFilterFind.Width = ScaleWidth - optFilterFind(1).Left - chkContinue.Width - cmdRefresh.Width - 500
'    End If
    If lblFilter.Visible Then
        lblFilter.Top = txtFilterFind.Top + 30
        txtFilterFind.Left = lblFilter.Left + lblFilter.Width + 50
        txtFilterFind.Width = ScaleWidth - txtFilterFind.Left - chkContinue.Width - cmdRefresh.Width - 500
    End If
    
    'pic������.Visible = mbytStyle <> 1
    picSplit01_S.Visible = mbytStyle <> 1       '������ģ��¼��ģʽ
    picSplit03_S.Visible = mbytStyle <> 1
    
    If pic������.Visible Then
        If Not tvw���.SelectedItem Is Nothing Then
            If tvw���.SelectedItem.Tag = "3" Or tvw���.SelectedItem.Tag = "Root7" Then
                bln��ҩ = True
            End If
        End If
        With pic������
            .Top = 20
            .Left = 0
            .Height = ScaleHeight - .Top
            .Width = picSplit01_S.Left
        End With
        With tvw���
            .Top = 0
            .Left = 0
            .Width = pic������.Width
            .Height = picSplit03_S.Top
        End With
        With picSplit03_S
            If .Top > ScaleHeight - 2000 Then .Top = ScaleHeight - 2000
            .Left = 0
            .Width = pic������.Width
        End With
        With fra��̬
            If bln��ҩ = True Then
                .Visible = True
                .Top = picSplit03_S.Top + picSplit03_S.Height
                .Left = 0
                .Width = pic������.Width
            Else
                .Visible = False
            End If
        End With
        With lbl����
            If bln��ҩ = True Then
                .Top = fra��̬.Top + fra��̬.Height
            Else
                .Top = picSplit03_S.Top + picSplit03_S.Height
            End If
            .Left = 0
            .Width = pic������.Width
        End With
        With chkChoose
            .Top = lbl����.Top + 10
            .Left = lbl����.Width - chkChoose.Width
        End With
        With lvw����
            .Top = lbl����.Height + lbl����.Top
            .Left = 0
            .Height = pic������.Height - lbl����.Top - lbl����.Height
            .Width = pic������.Width
        End With
        
        With picSplit01_S
            .Visible = mbytStyle <> 1
            .Top = pic������.Top
            '.Left = pic������.Width
            .Height = pic������.Height
        End With
    End If
    
    With picҩƷ��
        .Top = 550
        .Left = IIf(pic������.Visible, picSplit01_S.Left + picSplit01_S.Width, 0)
        .Width = IIf(pic������.Visible, ScaleWidth - pic������.Width - picSplit01_S.Width, ScaleWidth)
        If chkContinue.Value Then
            .Height = ScaleHeight - .Top - IIf(pic������.Tag = "չ��", picѡ����.Height, lblѡ��.Height + picSplit02_S.Height)
        Else
            .Height = ScaleHeight - .Top
        End If
    End With
    With vsf���
        .Top = 0
        .Left = 0
        .Height = IIf(vsf����.Visible, picSplit04_S.Top, picҩƷ��.Height)
        .Width = picҩƷ��.Width
    End With
    With picSetCols(0)
        .Top = 30
        .Left = 40
        .Height = 220
        .Width = 220
    End With
    
    If picSplit04_S.Visible Then
        With picSplit04_S
            If .Top > picҩƷ��.ScaleHeight - 1000 Then .Top = picҩƷ��.ScaleHeight - 1000
            .Left = 0
            .Width = picҩƷ��.Width
        End With
        With vsf����
            If .Visible Then
                .Top = picSplit04_S.Top + picSplit04_S.Height
                .Left = 0
                .Width = picҩƷ��.Width
                .Height = picҩƷ��.Height - picSplit04_S.Top
            End If
        End With
        With picSetCols(1)
            If .Visible Then
                .Top = 30
                .Left = 40
                .Height = 220
                .Width = 220
            End If
        End With
    End If
    
    If picSplit02_S.Visible Then
        With picSplit02_S
            .Top = picҩƷ��.Top + picҩƷ��.Height
            .Left = picҩƷ��.Left
            .Width = picҩƷ��.Width
        End With
    End If
    
    If picѡ����.Visible Then
        With picѡ����
            .Tag = "����"
            picSplit02_S.MousePointer = 0
            Set picUpDown01.Picture = imgsMain.ListImages(2).Picture
            If .Tag = "չ��" Then
                .Height = ScaleHeight - picҩƷ��.Top - picҩƷ��.Height - picSplit02_S.Height
            Else
                .Height = lblѡ��.Height
            End If
            .Top = ScaleHeight - .Height
            .Left = picҩƷ��.Left 'IIf(pic������.Visible, pic������.Width, 0)
            .Width = ScaleWidth - IIf(pic������.Visible, pic������.Width + picSplit01_S.Width, 0)
        End With
        With lblѡ��
            .Top = 0
            .Left = 0
            .Width = picѡ����.Width
        End With
        With picUpDown01
            .Left = picѡ����.Width - .Width
            .Top = 0
        End With
        With picOK
            .Left = picUpDown01.Left - .Width
            .Top = 0
        End With
        With vsfѡ��
            .Visible = picѡ����.Tag = "չ��"
            .Top = lblѡ��.Height
            .Left = 0
            .Width = lblѡ��.Width
            If .Visible Then
                .Height = picѡ����.Height - lblѡ��.Height
            End If
        End With
    End If
    err.Clear: On Error GoTo 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strKeyName As String
    
    Call SaveWinState(Me, App.ProductName, mfrmMain.Caption & mbytStyle)
    '��������VSF���ͷ״̬
    mstr��� = GetVSFHead(vsf���)
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\" & Me.Name & mfrmMain.Caption & mbytStyle & "\VSFlexGrid", _
        vsf���.Name & vsf���.Tag & "����", _
        IIf(mstr��� = "", MCON_���, mstr���)

    mstr���� = GetVSFHead(vsf����)
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\" & Me.Name & mfrmMain.Caption & mbytStyle & "\VSFlexGrid", _
        vsf����.Name & vsf����.Tag & "����", _
        IIf(mstr���� = "", MCON_����, mstr����)
    
    mstrѡ�� = GetVSFHead(vsfѡ��)
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\" & Me.Name & mfrmMain.Caption & mbytStyle & "\VSFlexGrid", _
        vsfѡ��.Name & vsfѡ��.Tag & "����", _
        IIf(mstrѡ�� = "", MCON_ѡ��, mstrѡ��)
    
    '����ע�����Ϣ(�Ƿ���ʾͣ��ҩƷ)
    SaveSetting "ZLSOFT", "˽��ģ��\ZLHIS\zl9MediStore", "��ʾͣ��ҩƷ", chkView.Value
    
'    If chkBatch.Visible = True Then
'        '����ѡ��״̬���浽ע���
'        Select Case UCase(mfrmMain.Name)
'            Case UCase("frmTransferCard")
'                strKeyName = "ҩƷ�ƿ����"
'            Case UCase("frmRequestDrugCard")
'                strKeyName = "ҩƷ�������"
'            Case UCase("frmDrawCard")
'                strKeyName = "ҩƷ���ù���"
'        End Select
'
'        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & strKeyName, "�������", IIf(chkBatch.Value = 1, 1, 0)
'    End If
    
    mlngLast = 0
End Sub

Private Sub Lvw����_ItemCheck(ByVal Item As MSComctlLib.listItem)
    Dim i As Integer
    
    If Item.Checked Then
        chkChoose.Value = 2
        For i = 1 To lvw����.ListItems.count
            If lvw����.ListItems(i).Checked = False Then
               
                myFilter
                SetColor
                Exit Sub
            End If
        Next
        chkChoose.Value = 1
    Else
        For i = 1 To lvw����.ListItems.count
            If lvw����.ListItems(i).Checked Then
                chkChoose.Value = 2
                
                myFilter
                SetColor
                Exit Sub
            End If
        Next
        chkChoose.Value = 0
    End If
    
    myFilter
    SetColor
End Sub

Private Sub optFilterFind_Click(index As Integer)
    txtFilterFind.SetFocus
End Sub

Private Sub picSetCols_Click(index As Integer)
    Dim frm������ As New frmVsColSel
    Dim vRect As RECT
    
    If index = 0 Then
        vRect = zlControl.GetControlRect(vsf���.hWnd)
        frm������.ShowColSet Me, "������", vsf���, _
            vRect.Left, vRect.Top + picSetCols(0).Top + picSetCols(0).Height + 10, _
            Me.Top + Me.Height - (vRect.Top + picSetCols(0).Top + picSetCols(0).Height + 120)
    Else
        vRect = zlControl.GetControlRect(vsf����.hWnd)
        frm������.ShowColSet Me, "������", vsf����, _
            vRect.Left, vRect.Top - 4000, _
            4000
    End If
End Sub

Private Sub picSetCols_MouseDown(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    picSetCols(index).BorderStyle = 1
End Sub

Private Sub picSetCols_MouseUp(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    picSetCols(index).BorderStyle = 0
End Sub

Private Sub picSplit01_S_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    With picSplit01_S
        If .Left + x < 2000 Then Exit Sub
        If .Left + x > ScaleWidth - 2000 Then Exit Sub
        .Move .Left + x, .Top
    End With
    With pic������
        .Width = .Width + x
    End With
    With picҩƷ��
        .Left = .Left + x
        .Width = .Width + x
    End With
    With picSplit02_S
        .Left = picҩƷ��.Left
        .Width = picҩƷ��.Width
    End With
    With picѡ����
        .Left = picҩƷ��.Left
        .Width = picҩƷ��.Width
    End With
    Call Form_Resize
End Sub

Private Sub picSplit02_S_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    If Not (picSplit02_S.MousePointer = 7 Or picSplit02_S.MousePointer = 0) Then Exit Sub
    With picSplit02_S
        If .Top + y < 1500 Then Exit Sub
        If .Top + y > Me.ScaleHeight - picSplit02_S.Height - lblѡ��.Height Then Exit Sub
        .Move .Left, .Top + y
    End With
    With picѡ����
        .Top = picSplit02_S.Top + picSplit02_S.Height
        .Height = Me.ScaleHeight - .Top
    End With
    With vsfѡ��
        .Top = lblѡ��.Height
        .Height = picѡ����.Height - lblѡ��.Height
    End With
End Sub

Private Sub picSplit02_S_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If picSplit02_S.Top >= Me.ScaleHeight - picSplit02_S.Height - lblѡ��.Height Then
        Call Form_Resize
    End If
End Sub

Private Sub picSplit03_S_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    With picSplit03_S
        If .Top + y < tvw���.Top + 1000 Then Exit Sub
        If .Top + y > ScaleHeight - 2000 Then Exit Sub
        .Move .Left, .Top + y
    End With
    With tvw���
        .Height = picSplit03_S.Top - .Top
    End With
    With fra��̬
        .Top = picSplit03_S.Top + picSplit03_S.Height
    End With
    With lbl����
         If tvw���.SelectedItem.Tag = "3" Then
            .Top = fra��̬.Top + fra��̬.Height
        Else
            .Top = picSplit03_S.Top + picSplit03_S.Height
        End If
        .Left = 0
        .Width = pic������.Width
    End With
    With chkChoose
        .Top = lbl����.Top + 10
    End With
    With lvw����
        .Top = lbl����.Height + lbl����.Top
        .Left = 0
        .Height = pic������.Height - lbl����.Top - lbl����.Height
        .Width = pic������.Width
    End With
End Sub

Private Sub picSplit04_S_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    With picSplit04_S
        If .Top + y < vsf���.Top + 1000 Then Exit Sub
        If .Top + y > picҩƷ��.ScaleHeight - 1000 Then Exit Sub
        .Move .Left, .Top + y
    End With
    With vsf���
        .Height = picSplit04_S.Top - .Top
    End With
    If vsf����.Visible Then
        With vsf����
            .Top = picSplit04_S.Top + picSplit04_S.Height
            .Height = picҩƷ��.Height - .Top
        End With
    End If
End Sub

Private Sub picOK_Click()
    Call CombinateRec
    Unload Me
End Sub

Private Sub picUpDown01_Click()
    If picѡ����.Tag = "չ��" Then
        picѡ����.Tag = "����"
        Set picUpDown01.Picture = imgsMain.ListImages(2).Picture
    Else
        picѡ����.Tag = "չ��"
        Set picUpDown01.Picture = imgsMain.ListImages(1).Picture
    End If
    ViewVSFѡ�� picѡ����.Tag = "չ��"
End Sub

Private Sub ViewVSFѡ��(ByVal blnDisp As Boolean)
    Dim i As Integer
    Dim y As Single
    
    vsfѡ��.Visible = blnDisp: vsfѡ��.TabStop = blnDisp
    If blnDisp Then
        picSplit02_S.MousePointer = 7
        For i = picSplit02_S.Top To Me.ScaleHeight \ 2 Step -100
            picSplit02_S.Top = i
            picSplit02_S_MouseMove 1, 0, 0, y
        Next
    Else
        picSplit02_S.MousePointer = 0
        For i = picSplit02_S.Top To Me.ScaleHeight - picSplit02_S.Height - lblѡ��.Height Step 100
            picSplit02_S.Top = i
            picSplit02_S_MouseMove 1, 0, 0, y
        Next
    End If
End Sub

Private Sub Fill_TVW���()
    Dim nodTmp As Node
    Dim rsTmp As ADODB.Recordset
    Dim lng�ⷿid As Long
    Dim Intĩ�� As Integer
    
    On Error GoTo errHandle
    gstrSQL = "Select ����, ���� From ������Ŀ��� " & _
              "Where Instr([1], ����, 1) > 0 " & _
              "Order by ���� "
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, "567")
    
    With tvw���
        .Nodes.Clear
'        Set nodTmp = .Nodes.Add(, , "Root", "����", 2, 2)
        Do While Not rsTmp.EOF
            Set nodTmp = .Nodes.Add(, , "Root" & rsTmp!����, rsTmp!����, 2, 2)
            nodTmp.Tag = "Root" & rsTmp!����
            rsTmp.MoveNext
        Loop
        rsTmp.Close
    End With
    
    '�������⣬�����ⷿΪ׼�������Գ���ⷿΪ׼
    lng�ⷿid = IIf(mbyt�༭ģʽ = 1, mlngĿ��ⷿ, mlng��Դ�ⷿ)
    If lng�ⷿid <> 0 Then
        '��ȡ�ÿⷿ���м��ͣ����û�ѡ��
        gstrSQL = "Select 1 From ��������˵�� " & _
                 " Where �������� Like '��ҩ%' And ����ID = [1] "
        Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[��鲿������]", lng�ⷿid)
        
        mbln��ҩ�ⷿ = Not rsTmp.EOF
        
        gstrSQL = "Select Distinct J.����,J.���� " & _
                  "From ����ִ�п��� A, ҩƷ���� B, ҩƷ���� J " & _
                  "Where A.������ĿID=B.ҩ��ID And B.ҩƷ����=J.���� And A.ִ�п���ID=[1] " & _
                  "Order by J.���� "
        Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[��ȡ�ÿⷿ���ڼ���]", lng�ⷿid)
    Else
        gstrSQL = "Select ����,���� From ҩƷ���� order by ���� "
        Call zldatabase.OpenRecordset(rsTmp, gstrSQL, "��ȡ����ҩƷ����")
    End If
    
    With rsTmp
        lvw����.ListItems.Clear
        Do While Not .EOF
            lvw����.ListItems.Add , "K" & !����, !����, 1, 1
            .MoveNext
        Loop
        If .State = 1 Then .Close
        
        gstrSQL = "Select ID, �ϼ�ID, ����, 1 as ĩ��, decode(����,1,'����ҩ',2,'�г�ҩ','�в�ҩ') as ����, ���� " & _
                  "From ���Ʒ���Ŀ¼ " & _
                  "Where ���� in (1,2,3) " & _
                  "Start With �ϼ�ID IS NULL Connect By Prior ID=�ϼ�ID Order by level,ID "
    End With
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "��ȡҩƷ��;����")
    With rsTmp
        If .EOF Then
            MsgBox "���ʼ��ҩƷ��;���ࣨҩƷ��;���ࣩ��", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '��ҩƷ��;��������װ��
        Do While Not .EOF
            Intĩ�� = IIf(!ĩ�� = 1, 3, 2)
            If IsNull(!�ϼ�ID) Then
                Set nodTmp = tvw���.Nodes.Add("Root" & !����, 4, "K_" & !Id, !����, Intĩ��, Intĩ��)
            Else
                Set nodTmp = tvw���.Nodes.Add("K_" & !�ϼ�ID, 4, "K_" & !Id, !����, Intĩ��, Intĩ��)
            End If
            nodTmp.Tag = !����   '��ŷ�������:1-����ҩ,2-�г�ҩ,3-�в�ҩ
            .MoveNext
        Loop
    End With

    With tvw���
        .Nodes(1).Selected = True
        If .Nodes(1).Children <> 0 Then
            Intĩ�� = 1
            .Nodes(Intĩ��).Child.Selected = True
            .SelectedItem.Selected = True
        ElseIf .Nodes(2).Children <> 0 Then
            Intĩ�� = 2
            .Nodes(Intĩ��).Child.Selected = True
            .SelectedItem.Selected = True
        ElseIf .Nodes(3).Children <> 0 Then
            Intĩ�� = 3
            .Nodes(Intĩ��).Child.Selected = True
            .SelectedItem.Selected = True
        Else
            Intĩ�� = 0
            .Nodes(1).Selected = True
            .SelectedItem.Selected = True
        End If
        If Intĩ�� <> 0 Then .Nodes(Intĩ��).Expanded = True
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetVSFHead(ByVal vsfObject As VSFlexGrid, ByVal strHead As String)
'--------------------------------
'���ܣ���ʼ��VSFlexGrid�ؼ����ͷ
'������
'  vsfObject��Ŀ��ؼ���
'  strHead�����ͷ�ĳ�ʼ���ִ�
'--------------------------------
    Dim arrCols As Variant, arrRows As Variant
    Dim i As Integer
    
    On Error GoTo errHandle
    
    arrRows = Split(strHead, "|")
    With vsfObject
        If .rows = 0 Then .rows = 1
        .Cols = UBound(arrRows) + 1
        For i = LBound(arrRows) To UBound(arrRows)
            If arrRows(i) <> "" Then
                arrCols = Split(arrRows(i), ",")
                '��1Ԫ�أ�Keyֵ
                .ColKey(i) = arrCols(0)
                '��2Ԫ�أ�Captionֵ
                If arrCols(1) = "" Then
                    .TextMatrix(0, i) = arrCols(0)
                Else
                    .TextMatrix(0, i) = arrCols(1)
                End If
                '��3Ԫ�أ�������
                If arrCols(2) = "" Then
                    .ColData(i) = 3
                Else
                    .ColData(i) = Val(arrCols(2))
                End If
                '��4Ԫ�أ����
                .ColWidth(i) = Val(arrCols(3))
                '��5Ԫ�أ���ʾ��ʽ
                If UBound(arrCols) > 3 Then
                    If UCase(arrCols(4)) = "D" Then
                        .ColFormat(i) = "yyyy-mm-dd"
                        .ColAlignment(i) = flexAlignCenterCenter
                    ElseIf UCase(arrCols(4)) = "T" Then
                        .ColFormat(i) = "hh:mm:ss"
                        .ColAlignment(i) = flexAlignCenterCenter
                    ElseIf UCase(arrCols(4)) = "DT" Then
                        .ColFormat(i) = "yyyy-mm-dd hh:mm:ss"
                        .ColAlignment(i) = flexAlignCenterCenter
                    ElseIf UCase(arrCols(4)) = "N" Then
                        .ColAlignment(i) = flexAlignRightCenter
                    Else
                        .ColAlignment(i) = flexAlignLeftCenter
                    End If
                Else
                    .ColAlignment(i) = flexAlignLeftCenter
                End If
                '������
                If Val(arrCols(2)) = 1 Or Val(arrCols(2)) = 2 Then
                    .ColHidden(i) = True
                Else
                    .ColHidden(i) = False
                End If
                
            End If
        Next
        If .Cols > 0 Then .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
    End With
    Exit Sub
    
errHandle:
    Call ErrCenter
End Sub

Private Function GetVSFHead(ByVal vsfObject As VSFlexGrid) As String
'---------------------------------
'���ܣ���ȡVSFĿ��ؼ��ı��ͷ�ִ�
'������vsfObject��Ŀ��ؼ�
'���أ����ͷ�ִ�
'---------------------------------
    Dim i As Integer
    Dim strHead As String, strCol As String
    
    With vsfObject
        strHead = ""
        For i = 0 To .Cols - 1
            '��1Ԫ�أ�Key
            strCol = .ColKey(i) & ","
            '��2Ԫ�أ�Caption
            If strCol = .TextMatrix(0, i) & "," Then
                strCol = strCol & ","
            Else
                strCol = strCol & .TextMatrix(0, i) & ","
            End If
            '��3Ԫ�أ�������
            If Val(.ColData(i)) = 3 Then
                If .ColHidden(i) Then
                    strCol = strCol & "2,"
                Else
                    strCol = strCol & ","
                End If
            Else
                If .ColHidden(i) = False And Val(.ColData(i)) = 2 Then
                    strCol = strCol & "3,"
                Else
                    strCol = strCol & .ColData(i) & ","
                End If
            End If
            '��4Ԫ�أ��п�
            If Val(.ColWidth(i)) = 0 Then
                strCol = strCol & ","
            Else
                strCol = strCol & .ColWidth(i) & ","
            End If
            '��5Ԫ�أ���ʾ��ʽ
            If Trim(.ColFormat(i)) = "" Then
                If .ColAlignment(i) = flexAlignRightCenter Then
                    strCol = strCol & "n"
                Else
                    strCol = Left(strCol, Len(strCol) - 1)
                End If
            Else
                If .ColFormat(i) = "yyyy-mm-dd" Then
                    strCol = strCol & "d"
                ElseIf .ColFormat(i) = "hh:mm:ss" Then
                    strCol = strCol & "t"
                ElseIf .ColFormat(i) = "yyyy-mm-dd hh:mm:ss" Then
                    strCol = strCol & "dt"
                End If
            End If
            '�������
            strHead = strHead & strCol & IIf(i = .Cols - 1, "", "|")
        Next
    End With
    GetVSFHead = strHead
End Function

Private Function GetVSFRow(ByVal vsfVal As VSFlexGrid) As Long
    Dim i As Long
    With vsfVal
        For i = 1 To vsfVal.rows - 1
            If vsfVal.RowHidden(i) = False Then
                GetVSFRow = i
                Exit Function
            End If
        Next
    End With
End Function

Private Sub SetColor()
    '���ñ����ɫ
    '�����ͣ��ҩƷ������������Ϊ��ɫ
    Dim lngRow As Long
    
    With vsf���
        If .rows > 1 Then
            For lngRow = 1 To .rows - 1
                If .TextMatrix(lngRow, .ColIndex("ͣ��")) = "��" Then
                    .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = vbRed
                End If
            Next
        End If
    End With
End Sub

Private Sub tvw���_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim strCols As String
    On Error GoTo errHandle
    If tvw���.Tag <> Node.Key Then
        Me.MousePointer = vbHourglass
        strCols = GetVSFHead(vsf���)
        mstrFilterClass = GetFilterClass(Node)
        
        If Node.Children = 0 And Node.Key Like "Root*" Then '����Ǹ��ڵ��Ҹýڵ����������ݵĻ���û������
            mstrFilterClass = "��;����id=99999999999999"
        Else
            mstrFilterClass = Left(mstrFilterClass, Len(mstrFilterClass) - 4)
        End If
        
        grsMaster.Filter = mstrFilterClass
        
        vsf���.rows = 1
        Set vsf���.DataSource = grsMaster
        '����ColKeyֵ
        SetColKey vsf���
        '��ʽ��VSF����
        FormatCols strCols
        '����������������
        With fra��̬
            If Node.Tag = "3" Then
                .Visible = True
            Else
                .Visible = False
            End If
        End With
        Call myFilter
        '������ɫ
        Call SetColor
    
        If Node.Tag = "3" Or Node.Tag = "Root7" Then '���в�ҩ����ʾ��ҩ��̬��ԭ����
            vsf���.ColHidden(vsf���.ColIndex("��ҩ��̬")) = False
            vsf���.ColHidden(vsf���.ColIndex("ԭ����")) = False
            vsf���.ColData(vsf���.ColIndex("ԭ����")) = 3
            vsf����.ColData(vsf����.ColIndex("ԭ����")) = 3
            vsf���.ColData(vsf���.ColIndex("��ҩ��̬")) = 3
            vsf����.ColWidth(vsf����.ColIndex("ԭ����")) = 1000
        Else
            vsf���.ColHidden(vsf���.ColIndex("��ҩ��̬")) = True
            vsf���.ColHidden(vsf���.ColIndex("ԭ����")) = True
            vsf���.ColData(vsf���.ColIndex("ԭ����")) = 1
            vsf����.ColData(vsf����.ColIndex("ԭ����")) = 1
            vsf���.ColData(vsf���.ColIndex("��ҩ��̬")) = 1
            vsf����.ColWidth(vsf����.ColIndex("ԭ����")) = 0
        End If
        
        Call Form_Resize
        
        'ˢ��VSF����
        If vsf���.rows > 1 Then
            vsf���.Row = 1
            ViewVSF���� vsf���
        Else
            vsf����.rows = 1
            picSplit04_S.Visible = False
            vsf����.Visible = False: vsf����.TabStop = False
            Call Form_Resize
        End If
        
        If chkChoose.Value = 2 Then
            Call myFilter
            SetColor
            If GetVSFlexRows(vsf���) <= 1 Then
                picSplit04_S.Visible = False
                vsf����.Visible = False: vsf����.TabStop = False
                Call Form_Resize
            Else
                vsf���.Row = GetVSFRow(vsf���)
            End If
        End If
        tvw���.Tag = Node.Key
    End If
    Me.MousePointer = vbDefault
    Exit Sub

errHandle:
    Me.MousePointer = vbDefault
    Call ErrCenter
End Sub

Private Sub txtFilterFind_Change()
    Dim strCols As String, strFind As String
    Dim i As Integer
    Dim rstemp As ADODB.Recordset
    
    strCols = GetVSFHead(vsf���)
    strFind = Trim(txtFilterFind.Text)
    txtFilterFind.Tag = ""
    
    mstrFilterClass = GetFilterSimpleCode(strFind)
    err.Clear: On Error GoTo errHandle
    grsMasterInput.Filter = mstrFilterClass
    err.Clear: On Error GoTo 0
    vsf���.rows = 1
    Set vsf���.DataSource = grsMasterInput
    '����ColKeyֵ
    SetColKey vsf���
    '��ʽ��VSF����
    FormatCols strCols
    '������ɫ
    Call SetColor
    'ˢ������
    ViewVSF���� vsf���
    '��ҩ������"ԭ����"��"��ҩ��̬"��
    Call HiddenColumns
    
    If mblnLoad = True Then
        If mbytStyle = 1 Then
            '¼��ģʽ�ż��
            gstrSQL = "Select a.���� as ҩƷ����,a.���� as ҩƷ����, b.���� As ��������, b.����, b.���ּ���, b.�����" & vbNewLine & _
                            "From �շ���ĿĿ¼ A," & vbNewLine & _
                            "     (Select �շ�ϸĿid, Max(Decode(����, '3', ����, Null)) ���ּ���, Max(Decode(����, '1', ����, Null)) ����," & vbNewLine & _
                            "              Max(Decode(����, '2', ����, Null)) �����, ����" & vbNewLine & _
                            "       From �շ���Ŀ����" & vbNewLine & _
                            "       Where ���� In (1, 2, 3) And ���� In (1, 3, 9)" & vbNewLine & _
                            "       Group By �շ�ϸĿid, ����) B" & vbNewLine & _
                            "Where a.Id = b.�շ�ϸĿid And a.��� In ('5', '6', '7') And (a.վ�� = '0' Or a.վ�� Is Null)"
    
            Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "��ѯ����")
            
            If mbyt�༭ģʽ = 1 Then
                '���
                rstemp.Filter = mstrFilterClass
                If grsMasterInput.RecordCount = 0 And rstemp.RecordCount = 0 Then
                    MsgBox "�޴�ҩƷ��", vbInformation, gstrSysName
                End If
            Else
                '����
                rstemp.Filter = mstrFilterClass
                If grsMasterInput.RecordCount = 0 And rstemp.RecordCount > 0 And mint�����γ��� = 1 Then
                    MsgBox "��ҩƷ�޿���ˣ�", vbInformation, gstrSysName
                ElseIf grsMasterInput.RecordCount = 0 And rstemp.RecordCount = 0 Then
                    MsgBox "�޴�ҩƷ��", vbInformation, gstrSysName
                End If
            End If
        End If
    End If
    Exit Sub
    
errHandle:
    MsgBox "�ı������Ƿ��ַ���", vbInformation, gstrSysName
    txtFilterFind.Text = "": txtFilterFind.Tag = "1"
    If txtFilterFind.Enabled And txtFilterFind.Visible Then txtFilterFind.SetFocus
End Sub

Private Sub txtFilterFind_GotFocus()
    txtFilterFind.SelStart = 0
    txtFilterFind.SelLength = Len(txtFilterFind.Text)
End Sub

Private Sub txtFilterFind_KeyDown(KeyCode As Integer, Shift As Integer)
    'If KeyCode = vbKeyReturn Then zlcommfun.PressKey vbKeyTab
End Sub

Private Sub txtFilterFind_KeyPress(KeyAscii As Integer)
    If InStr("`~!@#$%^&*()_-+={[}]|\:;""'<,>.?/", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    ElseIf KeyAscii >= 97 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    Else
    End If
End Sub

Private Sub vsf���_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow <> NewRow Then
        ViewVSF���� vsf���
    End If
End Sub

Private Function SetColPropStr(ByVal strCol As String, ByVal intProp As Integer, ByVal strParam As String) As String
'----------------------------------
'���ܣ��ı�ָ����������ֵ
'������
'  strCol���е����������ַ���
'  intProp���е������к�
'  strParam��Ҫ�ı��е�Ŀ��ֵ
'���أ��µ��������ַ���
'----------------------------------
    Dim arrElement As Variant
    Dim i As Integer, n As Integer
    Dim strTmp As String, strReturn As String
    
    arrElement = Split(strCol, ",")
    n = UBound(arrElement)
    For i = 0 To 4
        If intProp > i Then
            strTmp = "," & arrElement(i)
        ElseIf intProp = i Then
            strTmp = "," & strParam
        Else
            If i > n Then
                strTmp = ""
            Else
                strTmp = "," & arrElement(i)
            End If
        End If
        strReturn = strReturn & strTmp
    Next
    SetColPropStr = Right(strReturn, Len(strReturn) - 1)
End Function

Private Sub ParamsColsHead()
'----------------------------------
'���ܣ����ݲ������ã���Ӧ�����⴦��
'----------------------------------
    Dim intBegin As Integer, intLen As Integer
    Dim strColHead As String
    Dim arrCols��� As Variant, arrCols���� As Variant, arrColsѡ�� As Variant
    Dim i As Integer
    Dim strTmp As String

    'VSF��ͷ����
    SyncColumns MCON_���, mstr���
    SyncColumns MCON_����, mstr����
    SyncColumns MCON_ѡ��, mstrѡ��

    On Error GoTo errHandle
    
    arrCols��� = Split(mstr���, "|")
    arrCols���� = Split(mstr����, "|")
    arrColsѡ�� = Split(mstrѡ��, "|")
    
    strTmp = ""
    For i = LBound(arrCols���) To UBound(arrCols���)
        If InStr(";" & arrCols���(i) & ";", ";ͨ������,") > 0 Then
            If mbytStyle = 1 Then
                '¼��ҩƷ����   0��ƥ����ʾ�� 1��ͬʱ��ʾͨ��������Ʒ��
                strTmp = strTmp & SetColPropStr(arrCols���(i), enmColProperty.cpDisplay, IIf(gint����ҩƷ��ʾ = 0, "1", "0")) & "|"
            Else
                '��ʾҩƷ����   0����ʾͨ������ 1����ʾ��Ʒ���� 2��ͬʱ��ʾͨ��������Ʒ��
                strTmp = strTmp & SetColPropStr(arrCols���(i), enmColProperty.cpDisplay, IIf(gintҩƷ������ʾ = 1, "1", "0")) & "|"
            End If
        ElseIf InStr(";" & arrCols���(i) & ";", ";��Ʒ��,") > 0 Then
            If mbytStyle = 1 Then
                '¼��ҩƷ����   0��ƥ����ʾ�� 1��ͬʱ��ʾͨ��������Ʒ��
                strTmp = strTmp & SetColPropStr(arrCols���(i), enmColProperty.cpDisplay, IIf(gint����ҩƷ��ʾ = 0, "1", "0")) & "|"
            Else
                '��ʾҩƷ����   0����ʾͨ������ 1����ʾ��Ʒ���� 2��ͬʱ��ʾͨ��������Ʒ��
                strTmp = strTmp & SetColPropStr(arrCols���(i), enmColProperty.cpDisplay, IIf(gintҩƷ������ʾ = 0, "1", "0")) & "|"
            End If
        ElseIf InStr(";" & arrCols���(i) & ";", ";ҩƷ����,") > 0 Then
            If mbytStyle = 1 Then
                '¼��ҩƷ����   0��ƥ����ʾ�� 1��ͬʱ��ʾͨ��������Ʒ��
                strTmp = strTmp & SetColPropStr(arrCols���(i), enmColProperty.cpDisplay, IIf(gint����ҩƷ��ʾ = 0, "0", "1")) & "|"
            Else
                strTmp = strTmp & arrCols���(i) & "|"
            End If
        ElseIf InStr(";" & arrCols���(i) & ";", ";�ۼ۵�λ,") > 0 Then
            strTmp = strTmp & SetColPropStr(arrCols���(i), enmColProperty.cpDisplay, IIf(mintUnit = 1, "0", "1")) & "|"
        ElseIf InStr(";" & arrCols���(i) & ";", ";���ﵥλ,") > 0 Then
            strTmp = strTmp & SetColPropStr(arrCols���(i), enmColProperty.cpDisplay, IIf(mintUnit = 2, "0", "1")) & "|"
        ElseIf InStr(";" & arrCols���(i) & ";", ";סԺ��λ,") > 0 Then
            strTmp = strTmp & SetColPropStr(arrCols���(i), enmColProperty.cpDisplay, IIf(mintUnit = 3, "0", "1")) & "|"
        ElseIf InStr(";" & arrCols���(i) & ";", ";ҩ�ⵥλ,") > 0 Then
            strTmp = strTmp & SetColPropStr(arrCols���(i), enmColProperty.cpDisplay, IIf(mintUnit = 4, "0", "1")) & "|"
        ElseIf InStr(";" & arrCols���(i) & ";", ";�ۼ۰�װ,") > 0 Then
            strTmp = strTmp & SetColPropStr(arrCols���(i), enmColProperty.cpDisplay, IIf(mintUnit = 1, "0", "1")) & "|"
        ElseIf InStr(";" & arrCols���(i) & ";", ";�����װ,") > 0 Then
            strTmp = strTmp & SetColPropStr(arrCols���(i), enmColProperty.cpDisplay, IIf(mintUnit = 2, "0", "1")) & "|"
        ElseIf InStr(";" & arrCols���(i) & ";", ";סԺ��װ,") > 0 Then
            strTmp = strTmp & SetColPropStr(arrCols���(i), enmColProperty.cpDisplay, IIf(mintUnit = 3, "0", "1")) & "|"
        ElseIf InStr(";" & arrCols���(i) & ";", ";ҩ���װ,") > 0 Then
            strTmp = strTmp & SetColPropStr(arrCols���(i), enmColProperty.cpDisplay, IIf(mintUnit = 4, "0", "1")) & "|"
        ElseIf InStr(";" & arrCols���(i) & ";", ";�ϴβɹ���,") > 0 Then
            strTmp = strTmp & SetColPropStr(arrCols���(i), enmColProperty.cpDisplay, IIf(mblnCostView = False, "1", "0")) & "|"
        ElseIf InStr(";" & arrCols���(i) & ";", ";�����,") > 0 Then
            strTmp = strTmp & SetColPropStr(arrCols���(i), enmColProperty.cpDisplay, IIf(mblnCostView = False, "1", "0")) & "|"
        ElseIf InStr(";" & arrCols���(i) & ";", ";ָ��������,") > 0 Then
            strTmp = strTmp & SetColPropStr(arrCols���(i), enmColProperty.cpDisplay, IIf(mblnCostView = False, "1", "0")) & "|"
        ElseIf InStr(";" & arrCols���(i) & ";", ";��ͬ��λ,") > 0 Then
            '1300��ҩƷ�⹺��⣻ 1330��ҩƷ�ƻ�����
            strTmp = strTmp & SetColPropStr(arrCols���(i), enmColProperty.cpDisplay, IIf(glngModul = 1300 Or glngModul = 1330, "0", "1")) & "|"
        Else
            strTmp = strTmp & arrCols���(i) & "|"
        End If
    Next
    mstr��� = Left(strTmp, Len(strTmp) - 1)
    
    '����
    strTmp = ""
    For i = LBound(arrCols����) To UBound(arrCols����)
        If InStr(";" & arrCols����(i) & ";", ";��Ч��,") > 0 Then
            'Ч��     0��ʧЧ�ڣ�  1����Ч��
            If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 0 Then
                strTmp = strTmp & SetColPropStr(arrCols����(i), enmColProperty.cpCaption, "ʧЧ��") & "|"
            Else
                strTmp = strTmp & SetColPropStr(arrCols����(i), enmColProperty.cpCaption, "��Ч����") & "|"
            End If
        ElseIf InStr(";" & arrCols����(i) & ";", ";�������,") > 0 Then
            '1304��ҩƷ�ƿ���� 1343��ҩƷ�������
            strTmp = strTmp & SetColPropStr(arrCols����(i), enmColProperty.cpDisplay, IIf(glngModul = 1304 Or glngModul = 1343, "0", "1")) & "|"
        ElseIf InStr(";" & arrCols����(i) & ";", ";�ɱ���,") > 0 Then
            strTmp = strTmp & SetColPropStr(arrCols����(i), enmColProperty.cpDisplay, IIf(mblnCostView = False, "1", "0")) & "|"
        ElseIf InStr(";" & arrCols����(i) & ";", ";�����,") > 0 Then
            strTmp = strTmp & SetColPropStr(arrCols����(i), enmColProperty.cpDisplay, IIf(mblnCostView = False, "1", "0")) & "|"
        Else
            strTmp = strTmp & arrCols����(i) & "|"
        End If
    Next
    mstr���� = Left(strTmp, Len(strTmp) - 1)
    
    'ѡ��
    strTmp = ""
    For i = LBound(arrColsѡ��) To UBound(arrColsѡ��)
        If InStr(";" & arrColsѡ��(i) & ";", ";ͨ������,") > 0 Then
            If mbytStyle = 1 Then
                '¼��ҩƷ����   0��ƥ����ʾ�� 1��ͬʱ��ʾͨ��������Ʒ��
                strTmp = strTmp & SetColPropStr(arrColsѡ��(i), enmColProperty.cpDisplay, IIf(gint����ҩƷ��ʾ = 0, "1", "0")) & "|"
            Else
                '��ʾҩƷ����   0����ʾͨ������ 1����ʾ��Ʒ���� 2��ͬʱ��ʾͨ��������Ʒ��
                strTmp = strTmp & SetColPropStr(arrColsѡ��(i), enmColProperty.cpDisplay, IIf(gintҩƷ������ʾ = 1, "1", "0")) & "|"
            End If
        ElseIf InStr(";" & arrColsѡ��(i) & ";", ";��Ʒ��,") > 0 Then
            If mbytStyle = 1 Then
                '¼��ҩƷ����   0��ƥ����ʾ�� 1��ͬʱ��ʾͨ��������Ʒ��
                strTmp = strTmp & SetColPropStr(arrColsѡ��(i), enmColProperty.cpDisplay, IIf(gint����ҩƷ��ʾ = 0, "1", "0")) & "|"
            Else
                '��ʾҩƷ����   0����ʾͨ������ 1����ʾ��Ʒ���� 2��ͬʱ��ʾͨ��������Ʒ��
                strTmp = strTmp & SetColPropStr(arrColsѡ��(i), enmColProperty.cpDisplay, IIf(gintҩƷ������ʾ = 0, "1", "0")) & "|"
            End If
        ElseIf InStr(";" & arrColsѡ��(i) & ";", ";ҩƷ����,") > 0 Then
            If mbytStyle = 1 Then
                '¼��ҩƷ����   0��ƥ����ʾ�� 1��ͬʱ��ʾͨ��������Ʒ��
                strTmp = strTmp & SetColPropStr(arrColsѡ��(i), enmColProperty.cpDisplay, IIf(gint����ҩƷ��ʾ = 0, "0", "1")) & "|"
            Else
                strTmp = strTmp & arrColsѡ��(i) & "|"
            End If
        ElseIf InStr(";" & arrColsѡ��(i) & ";", ";�ϴβɹ���,") > 0 Then
            strTmp = strTmp & SetColPropStr(arrColsѡ��(i), enmColProperty.cpDisplay, IIf(mblnCostView = False, "1", "0")) & "|"
        ElseIf InStr(";" & arrColsѡ��(i) & ";", ";�����,") > 0 Then
            strTmp = strTmp & SetColPropStr(arrColsѡ��(i), enmColProperty.cpDisplay, IIf(mblnCostView = False, "1", "0")) & "|"
        ElseIf InStr(";" & arrColsѡ��(i) & ";", ";ָ��������,") > 0 Then
            strTmp = strTmp & SetColPropStr(arrColsѡ��(i), enmColProperty.cpDisplay, IIf(mblnCostView = False, "1", "0")) & "|"
        ElseIf InStr(";" & arrColsѡ��(i) & ";", ";��ͬ��λ,") > 0 Then
            '1300��ҩƷ�⹺��⣻ 1330��ҩƷ�ƻ�����
            strTmp = strTmp & SetColPropStr(arrColsѡ��(i), enmColProperty.cpDisplay, IIf(glngModul = 1300 Or glngModul = 1330, "0", "1")) & "|"
        ElseIf InStr(";" & arrColsѡ��(i) & ";", ";��Ч��,") > 0 Then
            'Ч��     0��ʧЧ�ڣ�  1����Ч��
            If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 0 Then
                strTmp = strTmp & SetColPropStr(arrColsѡ��(i), enmColProperty.cpCaption, "ʧЧ��") & "|"
            Else
                strTmp = strTmp & SetColPropStr(arrColsѡ��(i), enmColProperty.cpCaption, "��Ч����") & "|"
            End If
        Else
            strTmp = strTmp & arrColsѡ��(i) & "|"
        End If
    Next
    mstrѡ�� = Left(strTmp, Len(strTmp) - 1)
    
    Exit Sub
    
errHandle:
    Call ErrCenter
End Sub

Private Function GetFilterSimpleCode(ByVal strSimpleCode As String) As String
'--------------------------------------------
'���ܣ��õ�����ƥ��Ĺ�������
'������strSimpleCode��¼��ļ���
'���أ����������ַ���
'--------------------------------------------
    Dim strReturn As String, strTemp As String
    
    If strSimpleCode = "" Then Exit Function
    
    If gint���뷽ʽ = 1 Then
        '�����
        strTemp = "�����"
    Else
        'ƴ����
        strTemp = "����"
    End If
    
    If IsNumeric(strSimpleCode) Then
        '������
        strReturn = "ҩƷ���� like '" & mstrMatch & strSimpleCode & "%'" & _
                    " Or ���ּ��� like '" & mstrMatch & strSimpleCode & "%'" & _
                    " or ҩƷ���� like '" & mstrMatch & strSimpleCode & "%'"
    ElseIf zlStr.IsCharAlpha(strSimpleCode) Then
        '����ĸ
        strReturn = strTemp & " like '" & mstrMatch & strSimpleCode & "%'"
    ElseIf zlStr.IsCharChinese(strSimpleCode) Then
        '������
        strReturn = "ҩƷ���� like '" & mstrMatch & strSimpleCode & "%'"
    Else
        strReturn = "ҩƷ���� like '" & mstrMatch & strSimpleCode & "%'" & _
                    " Or " & strTemp & " like '" & mstrMatch & strSimpleCode & "%'" & _
                    " Or ҩƷ���� like '" & mstrMatch & strSimpleCode & "%'"
    End If
    
    GetFilterSimpleCode = strReturn
End Function

Private Function GetFilterClass(ByVal objNode As Node) As String
'--------------------------------------------
'���ܣ��õ���ǰ�ڵ������е�����ڵ�Ĺ�������
'������objNode����ǰ�ڵ����
'���أ����������ַ���
'--------------------------------------------
    Dim i, n As Integer
    Dim strReturn As String
    Dim objTmp As Node
    Dim strsql As String
    Dim bln��� As Boolean

    n = objNode.Children
        
    If Left(objNode.Key, 2) = "K_" Then
        strReturn = strReturn & "��;����id=" & Mid(objNode.Key, 3) & " or "
    End If
    If n > 0 Then
        Set objTmp = objNode.Child
        strReturn = strReturn & GetFilterClass(objTmp)
        For i = 2 To n
            Set objTmp = objTmp.Next
            strReturn = strReturn & GetFilterClass(objTmp)
        Next
    End If
    
    GetFilterClass = strReturn
End Function

Private Sub FillVSF(ByVal rsVal As ADODB.Recordset, ByVal vsfVal As VSFlexGrid)
'------------------------------
'���ܣ�ΪVSF�ؼ��������
'������
'  rsVal���������ݼ�
'------------------------------
    Dim i, j As Long
    Dim strData As String
    
    With rsVal
        vsfVal.rows = 1
        If .RecordCount > 0 Then
            .MoveFirst
        End If
        vsfVal.Redraw = False
        vsfVal.rows = .RecordCount + 1
        For i = 1 To .RecordCount
            For j = 0 To .Fields.count - 1
                If vsfVal.ColIndex(.Fields(j).Name) > -1 Then
                    vsfVal.TextMatrix(i, vsfVal.ColIndex(.Fields(j).Name)) = FieldValueDisp(.Fields, j, vsfVal.Name)
                Else
                    'Debug.Print .Fields(j).Name & "(��)"
                End If
            Next
            .MoveNext
        Next
        vsfVal.Redraw = True
    End With
    If glngModul = 1305 Then
        With vsf����
            If InStr(1, gstrprivs, "��ʾ�Է����") = 0 Then
                .ColData(.ColIndex("�������")) = 1
                .ColData(.ColIndex("�����")) = 1
                .ColData(.ColIndex("�����")) = 1
            Else
                .ColData(.ColIndex("�������")) = IIf(.ColData(.ColIndex("�������")) = 2, 2, 3)
                .ColData(.ColIndex("�����")) = IIf(.ColData(.ColIndex("�����")) = 2, 2, 3)
                .ColData(.ColIndex("�����")) = IIf(.ColData(.ColIndex("�����")) = 2, 2, 3)
            End If

        
            If InStr(1, gstrprivs, "��ʾ�Է����") = 0 Then
                For i = 1 To .rows - 1
                    If Val(.TextMatrix(i, .ColIndex("�������"))) > 0 Then
                        .TextMatrix(i, .ColIndex("�������")) = "��"
                    Else
                        .TextMatrix(i, .ColIndex("�������")) = "��"
                    End If
                Next
                .ColData(.ColIndex("�����")) = 1
                .ColData(.ColIndex("�����")) = 1
                .ColHidden(.ColIndex("�����")) = True
                .ColHidden(.ColIndex("�����")) = True
            Else
'                .ColData(.ColIndex("�������")) = IIf(.ColData(.ColIndex("�������")) = 2, 2, 3)
                .ColData(.ColIndex("�����")) = IIf(.ColData(.ColIndex("�����")) = 2, 2, 3)
                .ColData(.ColIndex("�����")) = IIf(.ColData(.ColIndex("�����")) = 2, 2, 3)
                .ColHidden(.ColIndex("�����")) = IIf(.ColData(.ColIndex("�����")) = 2, True, False)
                .ColHidden(.ColIndex("�����")) = IIf(.ColData(.ColIndex("�����")) = 2, True, False)
            End If
        End With
    End If
    
    If glngModul = 1343 Then '����
        With vsf����
            If InStr(1, gstrprivs, "��ʾ�Է����") = 0 Then
                For i = 1 To .rows - 1
                    If Val(.TextMatrix(i, .ColIndex("��������"))) > 0 Then
                        .TextMatrix(i, .ColIndex("��������")) = "��"
                    Else
                        .TextMatrix(i, .ColIndex("��������")) = "��"
                    End If
                Next
            End If
        End With
    End If
    
End Sub

Private Function FieldValueDisp(ByVal objFields As ADODB.Fields, ByVal intCol As Integer, ByVal strVSFName As String) As String
'--------------------------------
'���ܣ����ݲ������ã�������ʾ����
'������
'  objFields���м���
'  intCol�������
'  strVSFName��VSF�ؼ�Nameֵ
'���أ������������(�ַ���)
'--------------------------------
    Dim strReturn As String
    Dim dblUnit As Double
    
    Select Case mintUnit
    Case mconint���ﵥλ
        dblUnit = zlStr.Nvl(objFields("�����װ").Value, 0)
    Case mconintסԺ��λ
        dblUnit = zlStr.Nvl(objFields("סԺ��װ").Value, 0)
    Case mconintҩ�ⵥλ
        dblUnit = zlStr.Nvl(objFields("ҩ���װ").Value, 0)
    Case Else
        dblUnit = 1
    End Select
    
    If objFields(intCol).Name = "�ϴβɹ���" Or objFields(intCol).Name = "�ɱ���" Then
        '�ϴβɹ��ۡ��ۼۣ�vsf���ؼ����У� �ɱ��ۣ�vsf���οؼ�����
        If strVSFName <> "vsf���" Then
            strReturn = Format(zlStr.Nvl(objFields(intCol).Value, 0) * dblUnit, mStr�ɱ���)
        Else
            strReturn = Format(zlStr.Nvl(objFields(intCol).Value, 0), mStr�ɱ���)
        End If
    ElseIf objFields(intCol).Name = "�ۼ�" Then
        If strVSFName <> "vsf���" Then
            strReturn = Format(zlStr.Nvl(objFields(intCol).Value, 0) * dblUnit, mStr����)
        Else
            strReturn = Format(zlStr.Nvl(objFields(intCol).Value, 0), mStr����)
        End If
    ElseIf objFields(intCol).Name = "��������" Or objFields(intCol).Name = "�������" Or objFields(intCol).Name = "ʵ������" Then
        If strVSFName <> "vsf���" Then
            strReturn = Format(zlStr.Nvl(objFields(intCol).Value, 0) / dblUnit, mStr����)
        Else
            strReturn = Format(zlStr.Nvl(objFields(intCol).Value, 0), mStr����)
        End If
    ElseIf objFields(intCol).Name = "�����" Or objFields(intCol).Name = "�����" Then
        strReturn = Format(zlStr.Nvl(objFields(intCol).Value, 0), mStr���)
    Else
        strReturn = zlStr.Nvl(objFields(intCol).Value)
    End If
    
    '�̵��¼�����鿴�̵㵥��桱����
    If glngModul <> 1343 Then '������
        If mblnStore = False And (objFields(intCol).Name = "�ۼ�" Or objFields(intCol).Name = "�ɱ���" _
            Or objFields(intCol).Name = "��������" Or objFields(intCol).Name = "�������" Or objFields(intCol).Name = "�����" _
            Or objFields(intCol).Name = "�����") Then
            strReturn = ""
        End If
    Else '����
        If mblnStore = False And (objFields(intCol).Name = "�ۼ�" Or objFields(intCol).Name = "�ɱ���" _
            Or objFields(intCol).Name = "�������" Or objFields(intCol).Name = "�����" _
            Or objFields(intCol).Name = "�����") Then
            strReturn = ""
        End If
    End If
    
    FieldValueDisp = Trim(strReturn)
End Function

Private Sub ViewVSF����(ByVal vsfVal As VSFlexGrid)
    Dim strFilter As String
    Dim blnVisible As Boolean
    Dim int���� As Integer
    Dim lngRow As Long
    
    If mbyt�༭ģʽ <> 2 Then Exit Sub
    blnVisible = vsf����.Visible
    If grsSlave.State = adStateClosed Then
        picSplit04_S.Visible = False
        vsf����.Visible = False: vsf����.TabStop = False
        If blnVisible <> vsf����.Visible Then Call Form_Resize
        Exit Sub
    End If
    
    If GetVSFlexRows(vsfVal) <= 1 Then Exit Sub
    
    strFilter = "ҩƷID=" & Val(vsfVal.TextMatrix(vsfVal.Row, vsfVal.ColIndex("ҩƷID")))
    grsSlave.Filter = strFilter
    FillVSF grsSlave, vsf����
    
    int���� = Get�ⷿ����()
    With vsfVal
        '�����ҩƷ������
        If Not ((int���� = 3 And mbyt�ⷿ���� <> 3) Or (int���� = 1 And mbyt�ⷿ���� = 1) Or (int���� = 2 And mbyt�ⷿ���� = 2)) Then
            picSplit04_S.Visible = False
            vsf����.Visible = False: vsf����.TabStop = False
        Else
            '�ƿ���Ը��ݰ����θ�ѡ�����
            If UCase(mfrmMain.Name) = UCase("frmTransferCard") Or UCase(mfrmMain.Name) = UCase("frmRequestDrugCard") Or UCase(mfrmMain.Name) = UCase("frmDrawCard") Then
                If mint�����γ��� = 1 Then
                    picSplit04_S.Visible = True
                    vsf����.Visible = True: vsf����.TabStop = True
                Else
                    picSplit04_S.Visible = False
                    vsf����.Visible = False: vsf����.TabStop = False
                End If
            ElseIf UCase(mfrmMain.Name) = UCase("frmcheckCoursecard") Or UCase(mfrmMain.Name) = UCase("frmCheckCard") Then
                picSplit04_S.Visible = True
                vsf����.Visible = True: vsf����.TabStop = True
                '���ӿ����μ�¼
                If mbln������ Then
                    With vsf����
                        .rows = .rows + 1
                        lngRow = .rows - 1
                        .TextMatrix(lngRow, .ColIndex("RID")) = "1"
                        .TextMatrix(lngRow, .ColIndex("�ⷿ")) = mstr�����οⷿ
                        .TextMatrix(lngRow, .ColIndex("����")) = "-1"
                        .TextMatrix(lngRow, .ColIndex("����")) = "��������ҩƷ"
                        .TextMatrix(lngRow, .ColIndex("��Ч��")) = zldatabase.Currentdate
                        .RowPosition(lngRow) = 1
                    End With
                End If
            Else
                If grsSlave.RecordCount > 0 Then
                    picSplit04_S.Visible = True
                    vsf����.Visible = True: vsf����.TabStop = True
                Else
                    '���ӿ����μ�¼��ֻ���̵���Ч
                    If mbln������ Then
                        picSplit04_S.Visible = True
                        vsf����.Visible = True: vsf����.TabStop = True
                        With vsf����
                            .rows = .rows + 1
                            lngRow = .rows - 1
                            .TextMatrix(lngRow, .ColIndex("RID")) = "1"
                            .TextMatrix(lngRow, .ColIndex("�ⷿ")) = mstr�����οⷿ
                            .TextMatrix(lngRow, .ColIndex("����")) = "-1"
                            .TextMatrix(lngRow, .ColIndex("����")) = "��������ҩƷ"
                            .TextMatrix(lngRow, .ColIndex("��Ч��")) = Sys.Currentdate
                            .RowPosition(lngRow) = 1
                        End With
                    Else
                        picSplit04_S.Visible = False
                        vsf����.Visible = False: vsf����.TabStop = False
                    End If
                End If
            End If
        End If
    End With
    
    'ˢ��
    If blnVisible <> vsf����.Visible Then Call Form_Resize
    If mbln������ Then
        If vsf����.rows > 2 Then
            vsf����.Row = 2
        ElseIf vsf����.rows > 1 Then
            vsf����.Row = 1
        End If
    Else
        If vsf����.rows > 1 Then vsf����.Row = 1
    End If
End Sub




Private Sub SyncColumns(ByVal strInit As String, ByRef strRegister As String)
'-----------------------------------------------
'���ܣ�VSF��ͷ��ע��������ͷ������������һ��
'������
'  strInit����ͷ�ĳ�ʼֵ
'  strRegister��������ע������ͷ����������
'-----------------------------------------------
    Dim i As Integer, j As Integer, intOrder As Integer
    Dim arrInit As Variant, arrRegister As Variant
    Dim blnFind As Boolean
    Dim strTmp As String
    Dim bytInit As Byte, bytReg As Byte
    Dim rstemp As ADODB.Recordset
    
    Set rstemp = New ADODB.Recordset
    With rstemp
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Fields.Append "SN", adInteger
        .Fields.Append "VAL", adVarChar, 100
        .Open
    End With
    
    arrInit = Split(strInit, "|")
    arrRegister = Split(strRegister, "|")
    For i = LBound(arrInit) To UBound(arrInit)
        blnFind = False
        For j = LBound(arrRegister) To UBound(arrRegister)
            If Split(arrInit(i), ",")(0) = Split(arrRegister(j), ",")(0) Then
                rstemp.AddNew
                rstemp!SN = j
                '�Ƚ�Keyֵ
                '��ͷ�����������Ըı䣬ע��"0"��Val()��0�ǲ�ͬ�ģ�Val()=0������3����Ĭ����ʾ����
                If Split(arrRegister(j), ",")(2) = "0" Or Split(arrRegister(j), ",")(2) = "1" Then bytReg = 1 Else bytReg = 0
                If Split(arrInit(i), ",")(2) = "0" Or Split(arrInit(i), ",")(2) = "1" Then bytInit = 1 Else bytInit = 0
                If bytInit = bytReg Then
                    rstemp!Val = arrRegister(j)
                Else
                    strTmp = arrRegister(j)
                    strTmp = SetColPropStr(strTmp, enmColProperty.cpDisplay, Split(arrInit(i), ",")(2))
                    rstemp!Val = strTmp
                End If
                '�Ƚ�Formatֵ
                If UBound(Split(arrInit(i), ",")) >= 4 Then
                    If UBound(Split(arrRegister(j), ",")) >= 3 Then
                        strTmp = rstemp!Val
                        If UBound(Split(arrRegister(j), ",")) >= 4 Then
                            strTmp = SetColPropStr(strTmp, enmColProperty.cpFormat, Split(arrInit(i), ",")(4))
                        Else
                            strTmp = strTmp & "," & Split(arrInit(i), ",")(4)
                        End If
                        rstemp!Val = strTmp
                    End If
                End If
                rstemp.Update
                blnFind = True
                Exit For
            End If
        Next
        If blnFind = False Then
            '��������
            rstemp.AddNew
            rstemp!SN = i
            rstemp!Val = arrInit(i)
            rstemp.Update
        End If
    Next
    
    '����
    rstemp.Sort = "SN"
    strTmp = ""
    rstemp.MoveFirst
    Do While Not rstemp.EOF
        strTmp = strTmp & rstemp!Val & "|"
        rstemp.MoveNext
    Loop
    rstemp.Close
    strRegister = Left(strTmp, Len(strTmp) - 1)
End Sub

Private Function VerifyFilterStr(ByVal strFilter As String) As String
'----------------------------------------------------
'���ܣ����¼���������ַ������������ַ�������������
'----------------------------------------------------
    Dim i As Integer
    Dim strTmp As String
    
    If Len(strFilter) < 1 Then Exit Function
                                                                                                                                                                                                                                                               
    For i = 1 To Len(strFilter)
        If InStr("`~!@#$%^&*()_-+={[}]|\:;""'<,>.?/", Mid(strFilter, i, 1)) = 0 Then
            strTmp = strTmp & Mid(strFilter, i, 1)
        End If
    Next
    VerifyFilterStr = strTmp
End Function

Private Function CheckData() As Boolean
    Dim DblCurStock As Double       '��ǰ�����
    '����Ƿ�����ѡ��
    CheckData = False
    
'    If BlnSelect = False Then Exit Function
    
    'lng��Ӧ��ID��Ϊ�㣬��ʾ�˻����޿��ʱ��׼����
    If mlng��Ӧ�� <> 0 Then
        If vsf����.Visible Then
            If Val(vsf����.TextMatrix(vsf����.Row, vsf����.ColIndex("�ϴι�Ӧ��ID"))) <> 0 _
                And mlng��Ӧ�� <> Val(vsf����.TextMatrix(vsf����.Row, vsf����.ColIndex("�ϴι�Ӧ��ID"))) Then
                MsgBox "��ѡ����˻��̲��Ǹ�ҩƷ�Ĺ�Ӧ�̣����ܼ���������", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If

    '���Դ�ⷿ��Ŀ�ⷿΪ�գ��������ҩƷĿ¼�Լ��ڽ��г������ã����ж�
    If (mlng��Դ�ⷿ = 0 And mlngĿ��ⷿ = 0) Then
        CheckData = True
        Exit Function
    End If
    
    '������̵㵥����ҩƷѡ�����������жϣ�ֱ���˳�
    'If bln�̵㵥 Then
    If glngModul = 1307 Or glngModul = 1303 Then   'ҩƷ�̵����ҩƷ����۵���
        CheckData = True
        Exit Function
    End If
    
    If vsf����.Visible Then
        If Trim(vsf����.TextMatrix(vsf����.Row, vsf����.ColIndex("��������"))) = "��" Then
            CheckData = True
            Exit Function
        End If
        
        DblCurStock = Val(vsf����.TextMatrix(vsf����.Row, vsf����.ColIndex("��������")))
    Else
        DblCurStock = Val(vsf���.TextMatrix(vsf���.Row, vsf���.ColIndex("��������")))
    End If
    
    If DblCurStock > 0 Or mint�����γ��� = 0 Then
        '�������Ĳ������(����/�ƿ�/����)
        CheckData = True
        Exit Function
    Else
        Select Case mintStockCheck
        Case 1
            If MsgBox("��ҩƷ�Ѿ�û�п�棬�Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        Case 2
            MsgBox "��ҩƷ�Ѿ�û�п�棬���ܼ���������", vbInformation, gstrSysName
            Exit Function
        End Select
    End If
        
    CheckData = True
End Function

Private Sub vsf���_Click()
    If picѡ����.Tag = "չ��" Then picUpDown01_Click
End Sub

Private Sub vsf���_DblClick()
    Dim int���� As Integer
    
    mblnOK = False
    
    If GetVSFlexRows(vsf���) <= 1 Then Exit Sub
    
    If glngModul = 1305 Then
        With vsf���
            If .TextMatrix(.Row, .ColIndex("���ñ�־")) = "0" Or .TextMatrix(.Row, .ColIndex("���ñ�־")) = "" Then
                MsgBox "�ò��Ų��������ø�ҩƷ���뵽ҩƷ�����޶������ã�", vbInformation, gstrSysName
                Exit Sub
            End If
            
            '���۹���ѡ��ҩƷ����������Ҫ�����ܽ��г���ҵ��
            If gtype_UserSysParms.P275_���۹���ģʽ = 2 And IsPriceAdjustMod(Val(vsf���.TextMatrix(vsf���.Row, vsf���.ColIndex("ҩƷID")))) = True Then
                If CheckPriceAdjust(Val(vsf���.TextMatrix(vsf���.Row, vsf���.ColIndex("ҩƷID"))), mlng��Դ�ⷿ, -1) = False Then
                    MsgBox "��ҩƷ�������۹���ģʽ�����ɱ��ۺ��ۼ۲�һ�£����ܿ�չҵ�����ȵ����۸�", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
        End With
    End If
    If vsf����.Visible Then
        '�ⷿ����
        If GetVSFlexRows(vsf����) > 2 Then Exit Sub
        Call vsf����_DblClick
        Exit Sub
    Else
        If glngModul = 1304 Or glngModul = 1305 Or glngModul = 1306 Or glngModul = 1307 Or glngModul = 1343 Then
            '���۹���ѡ��ҩƷ����������Ҫ�����ܽ��г���ҵ��
            If gtype_UserSysParms.P275_���۹���ģʽ = 2 And IsPriceAdjustMod(Val(vsf���.TextMatrix(vsf���.Row, vsf���.ColIndex("ҩƷID")))) = True Then
                If CheckPriceAdjust(Val(vsf���.TextMatrix(vsf���.Row, vsf���.ColIndex("ҩƷID"))), mlng��Դ�ⷿ, -1) = False Then
                    MsgBox "��ҩƷ�������۹���ģʽ�����ɱ��ۺ��ۼ۲�һ�£����ܿ�չҵ�����ȵ����۸�", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
        End If
                
        If UCase(mfrmMain.Name) <> UCase("frmOtherOutputCard") And UCase(mfrmMain.Name) <> UCase("frmTransferCard") And UCase(mfrmMain.Name) <> UCase("frmPurchaseCard") Then
            If FillVSFѡ��(Val(vsf���.TextMatrix(vsf���.Row, vsf���.ColIndex("ҩƷID")))) = False Then Exit Sub
        Else
            '���������������⡢�ƿ⡢�⹺�˻����������޿�治���ƿ�ͳ���
            If mbyt�༭ģʽ = 2 Then
                '������жϷ�������
                int���� = Get�ⷿ����()
                If (int���� = 3 And mbyt�ⷿ���� <> 3) Or (int���� = 1 And mbyt�ⷿ���� = 1) Or (int���� = 2 And mbyt�ⷿ���� = 2) Then
                    '�ⷿ����
                    If grsSlave.RecordCount > 0 Or mint�����γ��� = 0 Then
                        If FillVSFѡ��(Val(vsf���.TextMatrix(vsf���.Row, vsf���.ColIndex("ҩƷID")))) = False Then Exit Sub
                    Else
                        MsgBox "��ҩƷ�Ƿ���ҩƷ��û�п�棬���ܼ���������", vbInformation, gstrSysName
                        Exit Sub
                    End If
                Else
                    '������
                    If FillVSFѡ��(Val(vsf���.TextMatrix(vsf���.Row, vsf���.ColIndex("ҩƷID")))) = False Then Exit Sub
                End If
            Else
                If FillVSFѡ��(Val(vsf���.TextMatrix(vsf���.Row, vsf���.ColIndex("ҩƷID")))) = False Then Exit Sub
            End If
        End If
    End If
    
    If chkContinue.Value <> 1 Then
        Call CombinateRec
        mblnOK = True
        Unload Me
    End If
End Sub

Private Sub vsf���_EnterCell()
    '����ѵ�ִ�����ڶ��۸�δִ�У�ִ�м������
    Dim lngҩƷid As Long
    
    lngҩƷid = Val(vsf���.TextMatrix(vsf���.Row, vsf���.ColIndex("ҩƷID")))
    If lngҩƷid = 0 Then Exit Sub
    
    If mlngLast <> lngҩƷid Then
        Call AutoAdjustPrice_ByID(lngҩƷid)
    End If
    mlngLast = lngҩƷid
End Sub

Private Sub vsf���_GotFocus()
    SetGridFocus vsf���, True
End Sub

Private Sub vsf���_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    If KeyCode = vbKeyReturn Then
        If vsf���.RowHidden(vsf���.Row) Then Exit Sub '��ǰ���������˳�����
        
        Call vsf���_DblClick
    End If
    
End Sub

Private Sub vsf���_LostFocus()
    SetGridFocus vsf���, False
End Sub

Private Sub vsf����_DblClick()
    Dim intRow As Integer
    If GetVSFlexRows(vsf����) <= 1 Then
        If mint�����γ��� = 1 Then
            MsgBox "����ҩƷ�����γ��⣬�޿�治�ܼ���������"
        End If
        Exit Sub
    End If
    If glngModul = 1305 And mint�����γ��� = 0 Then '����
        With vsf���
            If .TextMatrix(.Row, .ColIndex("���ñ�־")) = "0" Or .TextMatrix(.Row, .ColIndex("���ñ�־")) = "" Then
                MsgBox "�ò��Ų��������ø�ҩƷ���뵽ҩƷ�����޶������ã�", vbInformation, gstrSysName
                Exit Sub
            End If
            
            For intRow = 1 To vsfѡ��.rows - 1
                If .TextMatrix(.Row, .ColIndex("ҩƷid")) = vsfѡ��.TextMatrix(intRow, vsfѡ��.ColIndex("ҩƷid")) Then
                    MsgBox "��ҩƷ�Ѿ���ҩƷѡ�����У���������ͬ�ŵ��������ö�Σ�", vbInformation, gstrSysName
                    Exit Sub
                End If
            Next
        End With
    End If
    
    If glngModul = 1304 And mint�����γ��� = 0 Then  '�ƿ�
        With vsf���
            For intRow = 1 To vsfѡ��.rows - 1
                If .TextMatrix(.Row, .ColIndex("ҩƷid")) = vsfѡ��.TextMatrix(intRow, vsfѡ��.ColIndex("ҩƷid")) Then
                    MsgBox "��ҩƷ�Ѿ���ҩƷѡ�����У���������ͬ�ŵ������ƿ��Σ�", vbInformation, gstrSysName
                    Exit Sub
                End If
            Next
        End With
    End If
    
    If glngModul = 1343 And mint�����γ��� = 0 Then    '����
        With vsf���
            For intRow = 1 To vsfѡ��.rows - 1
                If .TextMatrix(.Row, .ColIndex("ҩƷid")) = vsfѡ��.TextMatrix(intRow, vsfѡ��.ColIndex("ҩƷid")) Then
                    MsgBox "��ҩƷ�Ѿ���ҩƷѡ�����У���������ͬ�ŵ����������Σ�", vbInformation, gstrSysName
                    Exit Sub
                End If
            Next
        End With
    End If
    
    '���۹���ѡ��ҩƷ����������Ҫ�����ܽ��������ҵ��
'    If glngModul <> 1307 Or (glngModul = 1307 And vsf����.TextMatrix(vsf����.Row, vsf����.ColIndex("����")) <> "��������ҩƷ") Then
        If gtype_UserSysParms.P275_���۹���ģʽ = 2 And IsPriceAdjustMod(Val(vsf���.TextMatrix(vsf���.Row, vsf���.ColIndex("ҩƷID")))) = True Then
            If CheckPriceAdjust(Val(vsf���.TextMatrix(vsf���.Row, vsf���.ColIndex("ҩƷID"))), mlng��Դ�ⷿ, vsf����.TextMatrix(vsf����.Row, vsf����.ColIndex("����"))) = False Then
                MsgBox "��ҩƷ�������۹���ģʽ�����ɱ��ۺ��ۼ۲�һ�£����ܿ�չҵ�����ȵ����۸�", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
'    End If
    
    
    If FillVSFѡ��(Val(vsf���.TextMatrix(vsf���.Row, vsf���.ColIndex("ҩƷID"))), vsf����.TextMatrix(vsf����.Row, vsf����.ColIndex("����"))) = False Then Exit Sub
    
    If chkContinue.Value <> 1 Then
        If CombinateRec = False Then Exit Sub
        Unload Me
    End If
End Sub

Private Function Get�ⷿ����() As Integer
'-------------------------------------------------------------------------------------------
'���ܣ�����ⷿ��������
'���أ�0���ⷿ�������� 1��ҩ�������ҩ���������� 2��ҩ�ⲻ������ҩ�������� 3���ⷿ������
'-------------------------------------------------------------------------------------------
    Dim intReturn As Integer
    With vsf���
        If .TextMatrix(.Row, .ColIndex("ҩ�����")) = "��" Or .TextMatrix(.Row, .ColIndex("ҩ������")) = "��" Then
            If .TextMatrix(.Row, .ColIndex("ҩ�����")) = "��" And .TextMatrix(.Row, .ColIndex("ҩ������")) = "��" Then
                intReturn = 3
            ElseIf .TextMatrix(.Row, .ColIndex("ҩ�����")) = "��" Then
                intReturn = 1
            Else
                intReturn = 2
            End If
        End If
    End With
    Get�ⷿ���� = intReturn
End Function

Private Function FillVSFѡ��(ByVal lngDrugID As Long, Optional ByVal str���� As String) As Boolean
    Dim blnValid As Boolean
    Dim lngRow As Long, i As Long
    Dim int���� As Integer
    Dim dblPrice As Double
    Dim rstemp As ADODB.Recordset
    
    On Error GoTo errHandle
    '���ҩƷ�ظ�
    If chkContinue.Value = 1 Then
        For i = 1 To vsfѡ��.rows - 1
            If Val(vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("ҩƷID"))) = lngDrugID Then
                If vsf����.Visible Then
                    If vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("����")) = str���� Then
                        Exit Function
                    End If
                Else
                    Exit Function
                End If
            End If
        Next
    End If
    
    '�������͵����ݼ��
    If mbyt�༭ģʽ = 2 Then If CheckData = False Then Exit Function
'
'    '�������������������Ƿ�һ��
'    blnValid = ���������(IIf(mbyt�༭ģʽ = 2, mlng��Դ�ⷿ, mlngĿ��ⷿ), lngDrugID)
'    If blnValid = False Then
'        MsgBox "���ָ�ҩƷ�ڵ�ǰ�ⷿ�еĿ���¼���ڴ��󣨿����ǻ����������ô������鵱ǰ�ⷿ�Ĳ������ʼ���ҩƷ�ķ������ԣ���", vbInformation, gstrSysName
'        Exit Function
'    End If
    
    '����
'    int���� = Get�ⷿ����()
        
    '���vsfѡ��
    With vsfѡ��
        .rows = .rows + 1
        lngRow = .rows - 1
        .TextMatrix(lngRow, .ColIndex("����")) = vsf���.TextMatrix(vsf���.Row, vsf���.ColIndex("����"))
        .TextMatrix(lngRow, .ColIndex("ҩ������")) = vsf���.TextMatrix(vsf���.Row, vsf���.ColIndex("ҩ������"))
        .TextMatrix(lngRow, .ColIndex("��Դ")) = vsf���.TextMatrix(vsf���.Row, vsf���.ColIndex("��Դ"))
        .TextMatrix(lngRow, .ColIndex("����ҩ��")) = vsf���.TextMatrix(vsf���.Row, vsf���.ColIndex("����ҩ��"))
        .TextMatrix(lngRow, .ColIndex("ͨ������")) = vsf���.TextMatrix(vsf���.Row, vsf���.ColIndex("ͨ������"))
        .TextMatrix(lngRow, .ColIndex("ҩ��ID")) = vsf���.TextMatrix(vsf���.Row, vsf���.ColIndex("ҩ��ID"))
        .TextMatrix(lngRow, .ColIndex("��;����ID")) = vsf���.TextMatrix(vsf���.Row, vsf���.ColIndex("��;����ID"))
        .TextMatrix(lngRow, .ColIndex("������λ")) = vsf���.TextMatrix(vsf���.Row, vsf���.ColIndex("������λ"))
        .TextMatrix(lngRow, .ColIndex("ҩƷ����")) = vsf���.TextMatrix(vsf���.Row, vsf���.ColIndex("ҩƷ����"))
        .TextMatrix(lngRow, .ColIndex("ҩƷ����")) = vsf���.TextMatrix(vsf���.Row, vsf���.ColIndex("ҩƷ����"))
        .TextMatrix(lngRow, .ColIndex("��Ʒ��")) = vsf���.TextMatrix(vsf���.Row, vsf���.ColIndex("��Ʒ��"))
        .TextMatrix(lngRow, .ColIndex("���")) = vsf���.TextMatrix(vsf���.Row, vsf���.ColIndex("���"))
        .TextMatrix(lngRow, .ColIndex("ҩ��ID")) = vsf���.TextMatrix(vsf���.Row, vsf���.ColIndex("ҩ��ID"))
        .TextMatrix(lngRow, .ColIndex("ҩƷID")) = vsf���.TextMatrix(vsf���.Row, vsf���.ColIndex("ҩƷID"))
        
        If Get�ۼ�(Val(.TextMatrix(lngRow, .ColIndex("ҩƷID"))), dblPrice) = False Then
            .RemoveItem .rows - 1
            Exit Function
        End If
        .TextMatrix(lngRow, .ColIndex("�ۼ�")) = vsf���.TextMatrix(vsf���.Row, vsf���.ColIndex("�ۼ�"))
        
        .TextMatrix(lngRow, .ColIndex("�ۼ۵�λ")) = vsf���.TextMatrix(vsf���.Row, vsf���.ColIndex("�ۼ۵�λ"))
        .TextMatrix(lngRow, .ColIndex("�ۼ۰�װ")) = vsf���.TextMatrix(vsf���.Row, vsf���.ColIndex("�ۼ۰�װ"))
        .TextMatrix(lngRow, .ColIndex("���Ч��")) = vsf���.TextMatrix(vsf���.Row, vsf���.ColIndex("��Ч��"))
        .TextMatrix(lngRow, .ColIndex("���ﵥλ")) = vsf���.TextMatrix(vsf���.Row, vsf���.ColIndex("���ﵥλ"))
        .TextMatrix(lngRow, .ColIndex("�����װ")) = vsf���.TextMatrix(vsf���.Row, vsf���.ColIndex("�����װ"))
        .TextMatrix(lngRow, .ColIndex("סԺ��λ")) = vsf���.TextMatrix(vsf���.Row, vsf���.ColIndex("סԺ��λ"))
        .TextMatrix(lngRow, .ColIndex("סԺ��װ")) = vsf���.TextMatrix(vsf���.Row, vsf���.ColIndex("סԺ��װ"))
        .TextMatrix(lngRow, .ColIndex("ҩ�ⵥλ")) = vsf���.TextMatrix(vsf���.Row, vsf���.ColIndex("ҩ�ⵥλ"))
        .TextMatrix(lngRow, .ColIndex("ҩ���װ")) = vsf���.TextMatrix(vsf���.Row, vsf���.ColIndex("ҩ���װ"))
        .TextMatrix(lngRow, .ColIndex("ҩ�����")) = vsf���.TextMatrix(vsf���.Row, vsf���.ColIndex("ҩ�����"))
        .TextMatrix(lngRow, .ColIndex("ҩ������")) = vsf���.TextMatrix(vsf���.Row, vsf���.ColIndex("ҩ������"))
        .TextMatrix(lngRow, .ColIndex("ʱ��")) = vsf���.TextMatrix(vsf���.Row, vsf���.ColIndex("ʱ��"))
        .TextMatrix(lngRow, .ColIndex("��׼�ĺ�")) = vsf���.TextMatrix(vsf���.Row, vsf���.ColIndex("��׼�ĺ�"))
        .TextMatrix(lngRow, .ColIndex("ָ��������")) = vsf���.TextMatrix(vsf���.Row, vsf���.ColIndex("ָ��������"))
        .TextMatrix(lngRow, .ColIndex("�ӳ���")) = vsf���.TextMatrix(vsf���.Row, vsf���.ColIndex("�ӳ���"))
        .TextMatrix(lngRow, .ColIndex("������")) = vsf���.TextMatrix(vsf���.Row, vsf���.ColIndex("������"))
        .TextMatrix(lngRow, .ColIndex("ԭ����")) = vsf���.TextMatrix(vsf���.Row, vsf���.ColIndex("ԭ����"))
        
        '�ɱ���
        dblPrice = Val(vsf���.TextMatrix(vsf���.Row, vsf���.ColIndex("�ϴβɹ���")))
        If dblPrice = 0 Then
            dblPrice = Val(vsf���.TextMatrix(vsf���.Row, vsf���.ColIndex("ָ��������")))
        End If
'        Select Case mintUnit
'        Case mconint���ﵥλ
'            dblPrice = dblPrice / IIf(Val(.TextMatrix(lngRow, .ColIndex("�����װ"))) = 0, 1, Val(.TextMatrix(lngRow, .ColIndex("�����װ"))))
'        Case mconintסԺ��λ
'            dblPrice = dblPrice / IIf(Val(.TextMatrix(lngRow, .ColIndex("סԺ��װ"))) = 0, 1, Val(.TextMatrix(lngRow, .ColIndex("סԺ��װ"))))
'        Case mconintҩ�ⵥλ
'            dblPrice = dblPrice / IIf(Val(.TextMatrix(lngRow, .ColIndex("ҩ���װ"))) = 0, 1, Val(.TextMatrix(lngRow, .ColIndex("ҩ���װ"))))
'        End Select
        .TextMatrix(lngRow, .ColIndex("�ɱ���")) = dblPrice
        
        'If mbyt�༭ģʽ = 2 And (int���� = 3 And mbyt�ⷿ���� <> 3) Or (int���� = 1 And mbyt�ⷿ���� = 1) Or (int���� = 2 And mbyt�ⷿ���� = 2) Then
        If vsf����.Visible Then
            .TextMatrix(lngRow, .ColIndex("����")) = vsf����.TextMatrix(vsf����.Row, vsf����.ColIndex("����"))
            .TextMatrix(lngRow, .ColIndex("��������")) = vsf����.TextMatrix(vsf����.Row, vsf����.ColIndex("��������"))
            .TextMatrix(lngRow, .ColIndex("�ϴι�Ӧ��ID")) = vsf����.TextMatrix(vsf����.Row, vsf����.ColIndex("�ϴι�Ӧ��ID"))
            .TextMatrix(lngRow, .ColIndex("��Ч��")) = vsf����.TextMatrix(vsf����.Row, vsf����.ColIndex("��Ч��"))
            .TextMatrix(lngRow, .ColIndex("��׼�ĺ�")) = vsf����.TextMatrix(vsf����.Row, vsf����.ColIndex("��׼�ĺ�"))
            .TextMatrix(lngRow, .ColIndex("������")) = vsf����.TextMatrix(vsf����.Row, vsf����.ColIndex("������"))
            .TextMatrix(lngRow, .ColIndex("ԭ����")) = vsf����.TextMatrix(vsf����.Row, vsf����.ColIndex("ԭ����"))
            
            .TextMatrix(lngRow, .ColIndex("����")) = vsf����.TextMatrix(vsf����.Row, vsf����.ColIndex("����"))
            '����̵��¼�����ޡ��鿴�̵㵥��桱��������
            If Not mblnStore Then
                ProcQuantity vsfѡ��, lngRow, .TextMatrix(lngRow, .ColIndex("ҩƷID"))
            Else
                .TextMatrix(lngRow, .ColIndex("��������")) = vsf����.TextMatrix(vsf����.Row, vsf����.ColIndex("��������"))
                .TextMatrix(lngRow, .ColIndex("�������")) = vsf����.TextMatrix(vsf����.Row, vsf����.ColIndex("�������"))
                .TextMatrix(lngRow, .ColIndex("�����")) = vsf����.TextMatrix(vsf����.Row, vsf����.ColIndex("�����"))
                .TextMatrix(lngRow, .ColIndex("�����")) = vsf����.TextMatrix(vsf����.Row, vsf����.ColIndex("�����"))
                .TextMatrix(lngRow, .ColIndex("ʵ������")) = vsf����.TextMatrix(vsf����.Row, vsf����.ColIndex("ʵ������"))
            End If
        Else
            '��ȡ������ҩƷ��������Ч����Ϣ
            gstrSQL = "Select �ϴ�����,Ч��,�ϴι�Ӧ��id,�ϴ��������� AS ��������,��׼�ĺ�,�ϴβ��� From ҩƷ��� " & _
                     " Where �ⷿID=[1] And ҩƷID=[2] And ����=1 "
            Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[��ȡ������ҩƷ��������Ч����Ϣ]", mlng��Դ�ⷿ, CLng(lngDrugID))
            If rstemp.RecordCount > 0 Then
                .TextMatrix(lngRow, .ColIndex("����")) = zlStr.Nvl(rstemp!�ϴ�����)
                If Not IsNull(rstemp!��������) Then
                    .TextMatrix(lngRow, .ColIndex("��������")) = zlStr.Nvl(rstemp!��������)
                End If
                .TextMatrix(lngRow, .ColIndex("�ϴι�Ӧ��ID")) = zlStr.Nvl(rstemp!�ϴι�Ӧ��ID)
                .TextMatrix(lngRow, .ColIndex("������")) = zlStr.Nvl(rstemp!�ϴβ���)
                
                If Not IsNull(rstemp!Ч��) Then
                    'If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1 And Nvl(!Ч��) <> "" Then
                    If .TextMatrix(0, .ColIndex("��Ч��")) = "��Ч����" Then
                        '����Ϊ��Ч��
                        .TextMatrix(lngRow, .ColIndex("��Ч��")) = Format(DateAdd("D", -1, rstemp!Ч��), "yyyy-mm-dd")
                    Else
                        .TextMatrix(lngRow, .ColIndex("��Ч��")) = zlStr.Nvl(rstemp!Ч��)
                    End If
                End If
                .TextMatrix(lngRow, .ColIndex("��׼�ĺ�")) = zlStr.Nvl(rstemp!��׼�ĺ�)
            End If
            rstemp.Close
            
            .TextMatrix(lngRow, .ColIndex("����")) = "0"
            '����̵��¼�����ޡ��鿴�̵㵥��桱��������
            If Not mblnStore Then
                ProcQuantity vsfѡ��, lngRow, .TextMatrix(lngRow, .ColIndex("ҩƷID"))
            Else
                .TextMatrix(lngRow, .ColIndex("��������")) = vsf���.TextMatrix(vsf���.Row, vsf���.ColIndex("��������"))
                .TextMatrix(lngRow, .ColIndex("�������")) = vsf���.TextMatrix(vsf���.Row, vsf���.ColIndex("�������"))
                .TextMatrix(lngRow, .ColIndex("�����")) = vsf���.TextMatrix(vsf���.Row, vsf���.ColIndex("�����"))
                .TextMatrix(lngRow, .ColIndex("�����")) = vsf���.TextMatrix(vsf���.Row, vsf���.ColIndex("�����"))
                .TextMatrix(lngRow, .ColIndex("ʵ������")) = vsf���.TextMatrix(vsf���.Row, vsf���.ColIndex("ʵ������"))
            End If
        End If
    End With
    
    lblѡ��.Caption = "ѡ��ҩƷ��" & vsfѡ��.rows - 1 & "����"
    
    'ѡ�����
    FillVSFѡ�� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ProcQuantity(ByVal vsfVal As VSFlexGrid, ByVal lngRow As Long, ByVal lngDrugID As Long)
'----------------------------------------------------------
'���ܣ������������̵��¼�����ޡ��鿴�̵㵥��桱��������
'----------------------------------------------------------
    Dim dbl�������� As Double, dbl������� As Double, dblʵ������ As Double
    Dim dbl����� As Double, dbl����� As Double
    With grsMaster
        .Find "ҩƷid=" & lngDrugID
        If Not .EOF Then
            dbl�������� = zlStr.Nvl(!��������, 0)
            dbl������� = zlStr.Nvl(!�������, 0)
            dbl����� = zlStr.Nvl(!�����, 0)
            dbl����� = zlStr.Nvl(!�����, 0)
            dblʵ������ = zlStr.Nvl(!ʵ������, 0)
        End If
    End With
    With vsfVal
        .TextMatrix(lngRow, .ColIndex("��������")) = dbl��������
        .TextMatrix(lngRow, .ColIndex("�������")) = dbl�������
        .TextMatrix(lngRow, .ColIndex("�����")) = dbl�����
        .TextMatrix(lngRow, .ColIndex("�����")) = dbl�����
        .TextMatrix(lngRow, .ColIndex("ʵ������")) = dblʵ������
    End With
End Sub

Private Function Get�ۼ�(ByVal lngDrugID As Long, ByRef dblPrice As Double) As Boolean
'-----------------------------------
'���ܣ���ȡָ��ҩƷ�����۵�λ�۸�
'���أ�True�ɹ���Falseʧ��
'-----------------------------------
    Dim rstemp As ADODB.Recordset
    Dim strMsg As String
    
    On Error GoTo errHandle
    gstrSQL = "Select A.�ּ�, B.ָ��������, B.ָ�����ۼ�, C.���� ҩ������, C.���� ͨ������ " & _
              "From �շѼ�Ŀ A, ҩƷ��� B, �շ���ĿĿ¼ C " & _
              "Where A.�շ�ϸĿid = B.ҩƷid And B.ҩƷID = C.ID " & _
              "  And Sysdate Between A.ִ������ And Nvl(A.��ֹ����,Sysdate) And A.�շ�ϸĿID=[1] " & _
              GetPriceClassString("A")
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[��ȡ��ҩƷ�����۵�λ�۸�]", lngDrugID)
    
    If Not rstemp.EOF Then
        dblPrice = zlStr.Nvl(rstemp!�ּ�, 0)
    Else
        dblPrice = 0
    End If
    
    '���ָ�������ۣ�ָ�����ۼۣ�Ϊ0ʱ������Ը�ҩƷ����
    strMsg = ""
    If Not rstemp.EOF Then
        If rstemp!ָ�������� = 0 And rstemp!ָ�����ۼ� = 0 Then
            strMsg = "�ɹ��޼ۺ�ָ���ۼ�Ϊ0���������ü۸�"
        ElseIf rstemp!ָ�������� = 0 Then
            strMsg = "�ɹ��޼�Ϊ0���������ü۸�"
        ElseIf rstemp!ָ�����ۼ� = 0 Then
            strMsg = "ָ���ۼ�Ϊ0���������ü۸�"
        End If
        If strMsg <> "" Then strMsg = "[" & zlStr.Nvl(rstemp!ҩ������) & zlStr.Nvl(rstemp!ͨ������) & "]" & strMsg
    End If
    rstemp.Close
    
    If strMsg <> "" Then
        MsgBox strMsg, vbInformation, gstrSysName
        Get�ۼ� = False
        Exit Function
    End If
    
    Get�ۼ� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CombinateRec() As Boolean
    Dim i As Long
    Dim dblPrice As Double
    
    With vsfѡ��
        For i = 1 To .rows - 1
            If Val(.TextMatrix(i, .ColIndex("ҩƷID"))) > 0 Then
                mrsReturn.AddNew
                mrsReturn!���� = IIf(.TextMatrix(i, .ColIndex("����")) = "", Null, .TextMatrix(i, .ColIndex("����")))
                mrsReturn!ҩƷ���� = .TextMatrix(i, .ColIndex("ҩƷ����"))
                mrsReturn!ҩƷ��Դ = .TextMatrix(i, .ColIndex("��Դ"))
                mrsReturn!����ҩ�� = .TextMatrix(i, .ColIndex("����ҩ��"))
                mrsReturn!ͨ���� = .TextMatrix(i, .ColIndex("ͨ������"))
                mrsReturn!ҩ��ID = Val(.TextMatrix(i, .ColIndex("ҩ��ID")))
                mrsReturn!��;����id = Val(.TextMatrix(i, .ColIndex("��;����ID")))
                mrsReturn!������λ = .TextMatrix(i, .ColIndex("������λ"))
                mrsReturn!��Ʒ�� = .TextMatrix(i, .ColIndex("��Ʒ��"))
                mrsReturn!��� = .TextMatrix(i, .ColIndex("���"))
                mrsReturn!���� = .TextMatrix(i, .ColIndex("������"))
                mrsReturn!ԭ���� = .TextMatrix(i, .ColIndex("ԭ����"))
                mrsReturn!ҩ��ID = Val(.TextMatrix(i, .ColIndex("ҩ��ID")))
                mrsReturn!ҩƷID = Val(.TextMatrix(i, .ColIndex("ҩƷID")))
                
                mrsReturn!�ۼ۵�λ = .TextMatrix(i, .ColIndex("�ۼ۵�λ"))
                mrsReturn!����ϵ�� = Val(.TextMatrix(i, .ColIndex("�ۼ۰�װ")))
                mrsReturn!���Ч�� = .TextMatrix(i, .ColIndex("���Ч��"))
                mrsReturn!���ﵥλ = .TextMatrix(i, .ColIndex("���ﵥλ"))
                mrsReturn!�����װ = Val(.TextMatrix(i, .ColIndex("�����װ")))
                mrsReturn!סԺ��λ = .TextMatrix(i, .ColIndex("סԺ��λ"))
                mrsReturn!סԺ��װ = Val(.TextMatrix(i, .ColIndex("סԺ��װ")))
                mrsReturn!ҩ�ⵥλ = .TextMatrix(i, .ColIndex("ҩ�ⵥλ"))
                mrsReturn!ҩ���װ = Val(.TextMatrix(i, .ColIndex("ҩ���װ")))
                mrsReturn!ҩ����� = IIf(.TextMatrix(i, .ColIndex("ҩ�����")) = "��", 1, 0)
                mrsReturn!ҩ������ = IIf(.TextMatrix(i, .ColIndex("ҩ������")) = "��", 1, 0)
                mrsReturn!ʱ�� = IIf(.TextMatrix(i, .ColIndex("ʱ��")) = "��", 1, 0)
                mrsReturn!�ϴι�Ӧ��ID = Val(.TextMatrix(i, .ColIndex("�ϴι�Ӧ��ID")))
                mrsReturn!��׼�ĺ� = .TextMatrix(i, .ColIndex("��׼�ĺ�"))
                mrsReturn!���� = IIf(.TextMatrix(i, .ColIndex("����")) = "", "0", .TextMatrix(i, .ColIndex("����")))
                mrsReturn!���� = IIf(.TextMatrix(i, .ColIndex("����")) = "", Null, .TextMatrix(i, .ColIndex("����")))
                mrsReturn!�������� = IIf(.TextMatrix(i, .ColIndex("��������")) = "", Null, .TextMatrix(i, .ColIndex("��������")))
                mrsReturn!Ч�� = IIf(.TextMatrix(i, .ColIndex("��Ч��")) = "", Null, .TextMatrix(i, .ColIndex("��Ч��")))
                mrsReturn!�������� = Val(.TextMatrix(i, .ColIndex("��������")))
                mrsReturn!ʵ������ = Val(.TextMatrix(i, .ColIndex("ʵ������")))
                mrsReturn!ʵ�ʽ�� = Val(.TextMatrix(i, .ColIndex("�����")))
                mrsReturn!ʵ�ʲ�� = Val(.TextMatrix(i, .ColIndex("�����")))
                mrsReturn!������� = Val(.TextMatrix(i, .ColIndex("�������")))
                
                dblPrice = Val(.TextMatrix(i, .ColIndex("ָ��������")))
                Select Case mintUnit
                    Case mconint���ﵥλ
                        dblPrice = dblPrice / IIf(Val(.TextMatrix(i, .ColIndex("�����װ"))) = 0, 1, Val(.TextMatrix(i, .ColIndex("�����װ"))))
                    Case mconintסԺ��λ
                        dblPrice = dblPrice / IIf(Val(.TextMatrix(i, .ColIndex("סԺ��װ"))) = 0, 1, Val(.TextMatrix(i, .ColIndex("סԺ��װ"))))
                    Case mconintҩ�ⵥλ
                        dblPrice = dblPrice / IIf(Val(.TextMatrix(i, .ColIndex("ҩ���װ"))) = 0, 1, Val(.TextMatrix(i, .ColIndex("ҩ���װ"))))
                End Select
                mrsReturn!ָ�������� = dblPrice
                                
                mrsReturn!�ӳ��� = Val(.TextMatrix(i, .ColIndex("�ӳ���")))
                
                dblPrice = Val(.TextMatrix(i, .ColIndex("�ۼ�")))
                Select Case mintUnit
                    Case mconint���ﵥλ
                        dblPrice = dblPrice / IIf(Val(.TextMatrix(i, .ColIndex("�����װ"))) = 0, 1, Val(.TextMatrix(i, .ColIndex("�����װ"))))
                    Case mconintסԺ��λ
                        dblPrice = dblPrice / IIf(Val(.TextMatrix(i, .ColIndex("סԺ��װ"))) = 0, 1, Val(.TextMatrix(i, .ColIndex("סԺ��װ"))))
                    Case mconintҩ�ⵥλ
                        dblPrice = dblPrice / IIf(Val(.TextMatrix(i, .ColIndex("ҩ���װ"))) = 0, 1, Val(.TextMatrix(i, .ColIndex("ҩ���װ"))))
                End Select
                mrsReturn!�ۼ� = dblPrice
                
                dblPrice = Val(.TextMatrix(i, .ColIndex("�ɱ���")))
                Select Case mintUnit
                    Case mconint���ﵥλ
                        dblPrice = dblPrice / IIf(Val(.TextMatrix(i, .ColIndex("�����װ"))) = 0, 1, Val(.TextMatrix(i, .ColIndex("�����װ"))))
                    Case mconintסԺ��λ
                        dblPrice = dblPrice / IIf(Val(.TextMatrix(i, .ColIndex("סԺ��װ"))) = 0, 1, Val(.TextMatrix(i, .ColIndex("סԺ��װ"))))
                    Case mconintҩ�ⵥλ
                        dblPrice = dblPrice / IIf(Val(.TextMatrix(i, .ColIndex("ҩ���װ"))) = 0, 1, Val(.TextMatrix(i, .ColIndex("ҩ���װ"))))
                End Select
                mrsReturn!�ɱ��� = dblPrice
                
                mrsReturn.Update
            End If
        Next
    End With
    CombinateRec = True
End Function

Private Sub InitReturnRecord()
    Set mrsReturn = New ADODB.Recordset
    With mrsReturn
        .Fields.Append "����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "ҩ������", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "ҩƷ��Դ", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "����ҩ��", adVarChar, 30, adFldIsNullable
        .Fields.Append "ͨ����", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "ҩ��ID", adDouble, 18, adFldIsNullable
        .Fields.Append "��;����ID", adDouble, 18, adFldIsNullable
        .Fields.Append "������λ", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "ҩƷ����", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "��Ʒ��", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "���", adLongVarChar, 30, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "ԭ����", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "ҩ��ID", adDouble, 18, adFldIsNullable
        .Fields.Append "ҩƷID", adDouble, 18, adFldIsNullable
        .Fields.Append "�ۼ�", adDouble, 18, adFldIsNullable
        .Fields.Append "�ۼ۵�λ", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "����ϵ��", adDouble, 11, adFldIsNullable
        .Fields.Append "���Ч��", adDouble, 5, adFldIsNullable
        .Fields.Append "���ﵥλ", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "�����װ", adDouble, 11, adFldIsNullable
        .Fields.Append "סԺ��λ", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "סԺ��װ", adDouble, 11, adFldIsNullable
        .Fields.Append "ҩ�ⵥλ", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "ҩ���װ", adDouble, 11, adFldIsNullable
        .Fields.Append "ҩ�����", adDouble, 2, adFldIsNullable
        .Fields.Append "ҩ������", adDouble, 2, adFldIsNullable
        .Fields.Append "ʱ��", adDouble, 2, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "��������", adDate, , adFldIsNullable
        .Fields.Append "Ч��", adDate, , adFldIsNullable
        .Fields.Append "��������", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "ʵ������", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "ʵ�ʽ��", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "ʵ�ʲ��", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "ָ��������", adDouble, 11, adFldIsNullable
        .Fields.Append "�ӳ���", adDouble, 11, adFldIsNullable
        .Fields.Append "�ϴι�Ӧ��ID", adDouble, 18, adFldIsNullable
        .Fields.Append "�������", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "��׼�ĺ�", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "�ɱ���", adDouble, 11, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

'Private Function Get���ÿ��(ByVal lngҩƷID As Long, Optional ByVal lng���� As Long = 0) As Single
'    Dim rsStock As New ADODB.Recordset
'
'    gstrSQL = " Select Sum(A.��������" & StrUnitString & ") ��������,Sum(A.ʵ������" & StrUnitString & ") ʵ������,sum(A.ʵ�ʽ��) ʵ�ʽ��,sum(A.ʵ�ʲ��) ʵ�ʲ��,Sum(A.ʵ������) ������� " & _
'              " From ҩƷ��� A,ҩƷ��� B " & _
'              " Where A.ҩƷID=B.ҩƷID And A.����=1 And A.ҩƷID=[1] " & IIf(lng���� = 0, "", " And Nvl(A.����,0)=[2] ")
'    If mlng��Դ�ⷿ <> 0 Or mlngĿ��ⷿ <> 0 Then
'        gstrSQL = gstrSQL & " And A.�ⷿID=[3]"
'    End If
'    gstrSQL = gstrSQL & " Group By A.ҩƷid"
'
'    Set rsStock = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[��ȡ���ÿ��]", lngҩƷID, lng����, IIf(mlng��Դ�ⷿ = 0, mlngĿ��ⷿ, mlng��Դ�ⷿ))
'
'    mdbl�������� = 0
'    mdblʵ�ʲ�� = 0
'    mdblʵ�ʽ�� = 0
'    mdblʵ������ = 0
'    mdbl������� = 0
'    If Not rsStock.EOF Then
'        mdbl�������� = IIf(IsNull(rsStock!��������), 0, rsStock!��������)
'        mdblʵ�ʲ�� = IIf(IsNull(rsStock!ʵ�ʲ��), 0, rsStock!ʵ�ʲ��)
'        mdblʵ�ʽ�� = IIf(IsNull(rsStock!ʵ�ʽ��), 0, rsStock!ʵ�ʽ��)
'        mdblʵ������ = IIf(IsNull(rsStock!ʵ������), 0, rsStock!ʵ������)
'        mdbl������� = IIf(IsNull(rsStock!�������), 0, rsStock!�������)
'    End If
'    Get���ÿ�� = mdbl��������
'End Function

Private Sub vsf����_GotFocus()
    SetGridFocus vsf����, True
End Sub

Private Sub vsf����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call vsf����_DblClick
End Sub

Private Sub vsf����_LostFocus()
    SetGridFocus vsf����, False
End Sub

Private Sub vsfѡ��_GotFocus()
    SetGridFocus vsfѡ��, True
End Sub

Private Sub vsfѡ��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If vsfѡ��.rows > 1 Then
            vsfѡ��.RemoveItem vsfѡ��.Row
            If vsfѡ��.rows = 1 Then
                lblѡ��.Caption = "ѡ��ҩƷ"
            Else
                lblѡ��.Caption = "ѡ��ҩƷ��" & vsfѡ��.rows - 1 & "����"
            End If
        End If
    End If
End Sub

Private Function Get��������(ByVal lngDeptId As Long) As String
    Dim strsql As String
    Dim rstemp As ADODB.Recordset
    
    On Error GoTo errHandle
    strsql = "select ���� from ���ű� where ID=[1] "
    Set rstemp = zldatabase.OpenSQLRecord(strsql, "��ȡ��������", lngDeptId)
    If Not rstemp.EOF Then
        Get�������� = zlStr.Nvl(rstemp!����)
    End If
    rstemp.Close
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub FormatCols(ByVal strCols As String)
'���ܣ�VSF�����ݼ�����󣬶�VSF�еĵ���
    Dim i As Integer
    Dim arrCols As Variant, arrColumn As Variant

    arrCols = Split(strCols, "|")
    With vsf���
        .Redraw = False
        For i = LBound(arrCols) To UBound(arrCols)
            arrColumn = Split(arrCols(i), ",")
            If .ColIndex(arrColumn(0)) >= 0 Then
                '��˳��
                .ColPosition(.ColIndex(arrColumn(0))) = i
                '������
                If UBound(arrColumn) > 1 Then
                    .ColData(i) = IIf(arrColumn(2) = "", 3, Val(arrColumn(2)))
                Else
                    .ColData(i) = 3
                End If
                If .ColData(i) = 1 Or .ColData(i) = 2 Then
                    .ColHidden(i) = True
                Else
                    .ColHidden(i) = False
                End If
                '�п��
                If UBound(arrColumn) > 2 Then
                    .ColWidth(i) = Val(arrColumn(3))
                Else
                    .ColWidth(i) = 0
                End If
                '��ʾ��ʽ
                If UBound(arrColumn) > 3 Then
                    If UCase(arrColumn(4)) = "D" Then
                        .ColFormat(i) = "yyyy-mm-dd"
                        .ColAlignment(i) = flexAlignCenterCenter
                    ElseIf UCase(arrColumn(4)) = "T" Then
                        .ColFormat(i) = "hh:mm:ss"
                        .ColAlignment(i) = flexAlignCenterCenter
                    ElseIf UCase(arrColumn(4)) = "DT" Then
                        .ColFormat(i) = "yyyy-mm-dd hh:mm:ss"
                        .ColAlignment(i) = flexAlignCenterCenter
                    ElseIf UCase(arrColumn(4)) = "N" Then
                        .ColAlignment(i) = flexAlignRightCenter
                    Else
                        .ColAlignment(i) = flexAlignLeftCenter
                    End If
                Else
                    .ColAlignment(i) = flexAlignLeftCenter
                End If
            End If
        Next
        '��ͷ�ı�������ʾ
        If .Cols > 0 Then .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .Redraw = True
    End With
End Sub

Private Sub SetColKey(ByVal vsfVal As VSFlexGrid)
'���ܣ�VSF�ؼ������ݼ�����ʱ��ColKeyֵû�У���0�е���ֵ��ΪColKeyֵ
    Dim i As Integer
    For i = 0 To vsfVal.Cols - 1
        vsfVal.ColKey(i) = vsfVal.TextMatrix(0, i)
    Next
End Sub

Private Sub InitVSF(ByVal vsfVal As VSFlexGrid)
  With vsfVal
    .AllowUserResizing = flexResizeColumns
    .Appearance = flexFlat
    .BackColorAlternate = .BackColor
    .BackColorSel = glngRowByNotFocus
    .BackColorBkg = &H8000000C
    .ForeColorSel = vbBlack
    .ExplorerBar = flexExSortShowAndMove
    '.FixedCols = 0
    .GridColor = &H80000010
    .GridLinesFixed = flexGridFlat
    .SelectionMode = flexSelectionByRow
    .SheetBorder = &H80000005
  End With
End Sub

Private Sub vsfѡ��_LostFocus()
    SetGridFocus vsfѡ��, True
End Sub
Private Sub HiddenColumns()
    Dim rsDetail As New Recordset
    Dim bln��ҩ�ⷿ As Boolean
    
    On Error GoTo errHandle
        
    With vsf���
        If .rows > 1 Then
            If mlngĿ��ⷿ = 0 And mlng��Դ�ⷿ = 0 Then
                .ColWidth(.ColIndex("ԭ����")) = 0
                vsf����.ColWidth(vsf����.ColIndex("ԭ����")) = 0
                .ColData(.ColIndex("ԭ����")) = 1
                vsf����.ColData(vsf����.ColIndex("ԭ����")) = 1
                Exit Sub
            End If
            
            gstrSQL = "select ��� from �շ���ĿĿ¼  where id=[1]"
            Set rsDetail = zldatabase.OpenSQLRecord(gstrSQL, "�жϿⷿ����", .TextMatrix(1, .ColIndex("ҩƷid")))
    
            If rsDetail!��� = "7" Then bln��ҩ�ⷿ = True
            If bln��ҩ�ⷿ Then
                .ColWidth(.ColIndex("ԭ����")) = 1000
                vsf����.ColWidth(vsf����.ColIndex("ԭ����")) = 1000
                .ColData(.ColIndex("ԭ����")) = 3
                vsf����.ColData(vsf����.ColIndex("ԭ����")) = 3
            Else
                vsf���.ColWidth(vsf���.ColIndex("ԭ����")) = 0
                vsf����.ColWidth(vsf����.ColIndex("ԭ����")) = 0
                .ColData(.ColIndex("ԭ����")) = 1
                vsf����.ColData(vsf����.ColIndex("ԭ����")) = 1
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
