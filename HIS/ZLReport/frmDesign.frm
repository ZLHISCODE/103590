VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0BE3824E-5AFE-4B11-A6BC-4B3AD564982A}#8.0#0"; "olch2x8.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDesign 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "�������"
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   10635
   FillColor       =   &H80000012&
   Icon            =   "frmDesign.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MouseIcon       =   "frmDesign.frx":020A
   ScaleHeight     =   6870
   ScaleWidth      =   10635
   Begin VB.ComboBox CboTest 
      Height          =   300
      Left            =   -8888
      Style           =   2  'Dropdown List
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   1980
      Width           =   1695
   End
   Begin MSComctlLib.ImageList ImgTool 
      Left            =   9720
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":035C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":0A56
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":1150
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":184A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":1F44
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":263E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":2DB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":34B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":3BAC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picFormat 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   2340
      ScaleHeight     =   405
      ScaleWidth      =   6180
      TabIndex        =   34
      TabStop         =   0   'False
      ToolTipText     =   "X����"
      Top             =   1410
      Width           =   6180
      Begin VB.CommandButton cmdDel 
         Height          =   375
         Left            =   5760
         Picture         =   "frmDesign.frx":49FE
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "ɾ�������ʽ"
         Top             =   15
         Width           =   405
      End
      Begin VB.CommandButton cmdAdd 
         Height          =   375
         Left            =   5325
         Picture         =   "frmDesign.frx":4D40
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "���ӱ����ʽ"
         Top             =   15
         Width           =   405
      End
      Begin MSComctlLib.ImageCombo cboFormat 
         Height          =   315
         Left            =   900
         TabIndex        =   0
         ToolTipText     =   "��������޸ĸ�ʽ����"
         Top             =   45
         Width           =   4410
         _ExtentX        =   7779
         _ExtentY        =   556
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         ImageList       =   "img16"
      End
      Begin VB.Label lblFormat 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�����ʽ"
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   105
         TabIndex        =   35
         Top             =   105
         Width           =   720
      End
   End
   Begin VB.PictureBox picSQL 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4980
      Left            =   0
      ScaleHeight     =   4980
      ScaleWidth      =   2280
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1530
      Width           =   2280
      Begin MSComctlLib.TreeView tvwSQL 
         Height          =   2085
         Left            =   90
         TabIndex        =   4
         Top             =   315
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   3678
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   494
         LabelEdit       =   1
         PathSeparator   =   "."
         Style           =   7
         Checkboxes      =   -1  'True
         ImageList       =   "img16"
         Appearance      =   1
      End
      Begin MSComctlLib.ListView lvwPar 
         DragIcon        =   "frmDesign.frx":5082
         Height          =   2325
         Left            =   90
         TabIndex        =   5
         Top             =   2775
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   4101
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "img16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "���"
            Object.Width           =   961
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "����"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "����"
            Object.Width           =   961
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "ȱʡֵ"
            Object.Width           =   2646
         EndProperty
      End
      Begin MSComctlLib.ImageList img16 
         Left            =   1680
         Top             =   540
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   11
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDesign.frx":538C
               Key             =   "SQL_Custom"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDesign.frx":54E6
               Key             =   "SQL_Group"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDesign.frx":5640
               Key             =   "Root"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDesign.frx":579A
               Key             =   "Other"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDesign.frx":58F4
               Key             =   "String"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDesign.frx":5A4E
               Key             =   "Number"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDesign.frx":5BA8
               Key             =   "Date"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDesign.frx":5D02
               Key             =   "Bin"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDesign.frx":5E5C
               Key             =   "Format"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDesign.frx":5FB6
               Key             =   "Pars"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDesign.frx":C818
               Key             =   "ParsRoot"
            EndProperty
         EndProperty
      End
      Begin VB.Label lblPar 
         Alignment       =   2  'Center
         BackColor       =   &H009B6737&
         Caption         =   "���ݲ���"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         MousePointer    =   7  'Size N S
         TabIndex        =   26
         Top             =   2505
         Width           =   2040
      End
      Begin VB.Label lblSQL 
         Alignment       =   2  'Center
         BackColor       =   &H009B6737&
         Caption         =   "��������Դ"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   105
         TabIndex        =   25
         Top             =   60
         Width           =   2055
      End
   End
   Begin VB.PictureBox picL 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4980
      Left            =   2280
      MousePointer    =   9  'Size W E
      ScaleHeight     =   4980
      ScaleWidth      =   45
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   1530
      Width           =   45
      Begin VB.Line Line2 
         BorderColor     =   &H80000015&
         X1              =   30
         X2              =   30
         Y1              =   0
         Y2              =   15360
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   15360
      End
   End
   Begin VB.PictureBox picR 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4980
      Left            =   8190
      MousePointer    =   9  'Size W E
      ScaleHeight     =   4980
      ScaleWidth      =   45
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   1530
      Width           =   45
      Begin VB.Line Line4 
         BorderColor     =   &H80000015&
         X1              =   30
         X2              =   30
         Y1              =   -60
         Y2              =   15300
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   0
         Y1              =   -60
         Y2              =   15300
      End
   End
   Begin VB.PictureBox picAtt 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4980
      Left            =   8235
      ScaleHeight     =   4980
      ScaleWidth      =   2400
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1530
      Width           =   2400
      Begin VB.CommandButton cmdAtt 
         Caption         =   "��"
         Height          =   285
         Left            =   1725
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   2130
         Visible         =   0   'False
         Width           =   300
      End
      Begin MSComCtl2.DTPicker dtpAtt 
         Height          =   300
         Left            =   960
         TabIndex        =   48
         Top             =   2520
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "HH:mm:ss"
         Format          =   298254338
         CurrentDate     =   41766
      End
      Begin VB.ComboBox cboText 
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1080
         TabIndex        =   47
         TabStop         =   0   'False
         Text            =   "cboAttText"
         Top             =   2400
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.TextBox txtAtt 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   1155
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   1890
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.PictureBox picM 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   45
         Left            =   150
         MousePointer    =   7  'Size N S
         ScaleHeight     =   45
         ScaleWidth      =   1740
         TabIndex        =   32
         Top             =   4125
         Width           =   1740
      End
      Begin VB.ComboBox cboAtt 
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1110
         Style           =   2  'Dropdown List
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   2475
         Width           =   960
      End
      Begin VSFlex8Ctl.VSFlexGrid mshAtt 
         Height          =   1095
         Left            =   120
         TabIndex        =   44
         Top             =   2400
         Width           =   1695
         _cx             =   1964641198
         _cy             =   1964640139
         Appearance      =   0
         BorderStyle     =   0
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
         BackColorSel    =   11103813
         ForeColorSel    =   16777215
         BackColorBkg    =   -2147483634
         BackColorAlternate=   -2147483643
         GridColor       =   0
         GridColorFixed  =   0
         TreeColor       =   -2147483644
         FloodColor      =   192
         SheetBorder     =   -2147483631
         FocusRect       =   1
         HighLight       =   0
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmDesign.frx":1307A
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
      Begin MSComctlLib.Toolbar tbrTool 
         Height          =   720
         Left            =   90
         TabIndex        =   37
         Top             =   300
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         ToolTips        =   0   'False
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ImgTool"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   9
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ѡ��"
               Key             =   "Point"
               Object.ToolTipText     =   "ѡ�����"
               Object.Tag             =   "ѡ��"
               ImageIndex      =   1
               Style           =   2
               Value           =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Line"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   2
               Style           =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Frame"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   3
               Style           =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ǩ"
               Key             =   "Note"
               Object.ToolTipText     =   "��ǩ"
               Object.Tag             =   "��ǩ"
               ImageIndex      =   4
               Style           =   2
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ͼƬ"
               Key             =   "Picture"
               Object.ToolTipText     =   "ͼƬ"
               Object.Tag             =   "ͼƬ"
               ImageIndex      =   5
               Style           =   2
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "���"
               Key             =   "Table"
               Object.ToolTipText     =   "���"
               Object.Tag             =   "���"
               ImageIndex      =   6
               Style           =   2
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ͼ��"
               Key             =   "Chart"
               Object.ToolTipText     =   "ͼ��"
               Object.Tag             =   "ͼ��"
               ImageIndex      =   7
               Style           =   2
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "BarCode"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   8
               Style           =   2
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��Ƭ"
               Key             =   "Card"
               Object.ToolTipText     =   "��Ƭ"
               Object.Tag             =   "��Ƭ"
               ImageIndex      =   9
               Style           =   2
            EndProperty
         EndProperty
      End
      Begin VB.Label lblAtt 
         Alignment       =   2  'Center
         BackColor       =   &H009B6737&
         Caption         =   "Ԫ������"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   105
         TabIndex        =   28
         Top             =   1605
         Width           =   2040
      End
      Begin VB.Label lblNote 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H00FFFFFF&
         Height          =   915
         Left            =   60
         TabIndex        =   29
         Top             =   4200
         Width           =   2250
      End
      Begin VB.Label lblTool 
         Alignment       =   2  'Center
         BackColor       =   &H009B6737&
         Caption         =   "����Ԫ��"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   90
         TabIndex        =   27
         Top             =   105
         Width           =   2220
      End
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   2760
      MouseIcon       =   "frmDesign.frx":13157
      ScaleHeight     =   3855
      ScaleWidth      =   5355
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2265
      Width           =   5355
      Begin VB.PictureBox picPaperSize 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   60
         Index           =   2
         Left            =   5190
         MousePointer    =   8  'Size NW SE
         ScaleHeight     =   60
         ScaleWidth      =   60
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "�϶��ɸı�ֽ�Ÿ߶ȺͿ��"
         Top             =   3690
         Width           =   60
      End
      Begin VB.PictureBox picPaperSize 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   60
         Index           =   1
         Left            =   105
         MousePointer    =   7  'Size N S
         ScaleHeight     =   60
         ScaleWidth      =   4935
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "�϶��ɸı�ֽ�Ÿ߶�"
         Top             =   3705
         Width           =   4935
      End
      Begin VB.PictureBox picPaperSize 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3480
         Index           =   0
         Left            =   5190
         MousePointer    =   9  'Size W E
         ScaleHeight     =   3480
         ScaleWidth      =   60
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "�϶��ɸı�ֽ�ſ��"
         Top             =   105
         Width           =   60
      End
      Begin VB.PictureBox picPaper 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         DrawStyle       =   2  'Dot
         ForeColor       =   &H00FF0000&
         Height          =   3525
         Left            =   105
         MouseIcon       =   "frmDesign.frx":132A9
         ScaleHeight     =   3525
         ScaleWidth      =   5025
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   120
         Width           =   5025
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   615
            Index           =   0
            Left            =   -8888
            ScaleHeight     =   585
            ScaleWidth      =   585
            TabIndex        =   42
            Top             =   1080
            Width           =   615
         End
         Begin C1Chart2D8.Chart2D Chart 
            Height          =   1440
            Index           =   0
            Left            =   -8888
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   1200
            Visible         =   0   'False
            Width           =   2100
            _Version        =   524288
            _Revision       =   7
            _ExtentX        =   3704
            _ExtentY        =   2540
            _StockProps     =   0
            ControlProperties=   "frmDesign.frx":133FB
         End
         Begin VB.PictureBox PicSplit 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1155
            Left            =   -8888
            ScaleHeight     =   1155
            ScaleWidth      =   15
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   150
            Visible         =   0   'False
            Width           =   15
            Begin VB.Line LineSplit 
               BorderStyle     =   3  'Dot
               X1              =   0
               X2              =   0
               Y1              =   0
               Y2              =   8000
            End
         End
         Begin VB.PictureBox PicFontTest 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   60
            Left            =   -8888
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   0
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.PictureBox LblSize 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   0
            Left            =   -8888
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   780
            Visible         =   0   'False
            Width           =   60
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh1 
            Height          =   585
            Index           =   0
            Left            =   -8888
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   2130
            _ExtentX        =   3757
            _ExtentY        =   1032
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            BackColorFixed  =   15724527
            ForeColorFixed  =   0
            ForeColorSel    =   16777215
            BackColorBkg    =   16777215
            BackColorUnpopulated=   16777215
            GridColor       =   0
            GridColorFixed  =   0
            GridColorUnpopulated=   16777215
            WordWrap        =   -1  'True
            AllowBigSelection=   0   'False
            ScrollTrack     =   -1  'True
            FocusRect       =   0
            HighLight       =   0
            GridLinesFixed  =   1
            GridLinesUnpopulated=   1
            ScrollBars      =   0
            MergeCells      =   1
            AllowUserResizing=   1
            Appearance      =   0
            MousePointer    =   99
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "frmDesign.frx":13A5A
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VSFlex8Ctl.VSFlexGrid msh 
            Height          =   1575
            Index           =   0
            Left            =   360
            TabIndex        =   45
            Top             =   50000
            Visible         =   0   'False
            Width           =   3135
            _cx             =   1964643738
            _cy             =   1964640986
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
            MouseIcon       =   "frmDesign.frx":13D74
            BackColor       =   16777215
            ForeColor       =   0
            BackColorFixed  =   16777215
            ForeColorFixed  =   0
            BackColorSel    =   10251637
            ForeColorSel    =   16777215
            BackColorBkg    =   16777215
            BackColorAlternate=   16777215
            GridColor       =   0
            GridColorFixed  =   0
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   0
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   5
            Cols            =   3
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   ""
            ScrollTrack     =   -1  'True
            ScrollBars      =   0
            ScrollTips      =   0   'False
            MergeCells      =   1
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
         Begin VB.Label lblshp 
            BackColor       =   &H8000000E&
            Height          =   735
            Index           =   0
            Left            =   -50000
            TabIndex        =   46
            Top             =   1560
            Width           =   1335
         End
         Begin VB.Shape Shp 
            Height          =   1575
            Index           =   0
            Left            =   -50000
            Top             =   1200
            Width           =   2040
         End
         Begin VB.Image ImgCode 
            Appearance      =   0  'Flat
            Height          =   735
            Index           =   0
            Left            =   -8888
            MouseIcon       =   "frmDesign.frx":14C4E
            Stretch         =   -1  'True
            Top             =   1230
            Width           =   555
         End
         Begin VB.Image Img 
            Appearance      =   0  'Flat
            Height          =   735
            Index           =   0
            Left            =   -8888
            MouseIcon       =   "frmDesign.frx":14F58
            Stretch         =   -1  'True
            Top             =   390
            Width           =   555
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            BackColor       =   &H00EFEFEF&
            Caption         =   "��ǩ"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   -2205
            MouseIcon       =   "frmDesign.frx":15262
            MousePointer    =   99  'Custom
            TabIndex        =   17
            Top             =   255
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.Label lblLine 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            ForeColor       =   &H80000008&
            Height          =   15
            Index           =   0
            Left            =   -2235
            MouseIcon       =   "frmDesign.frx":1556C
            MousePointer    =   99  'Custom
            TabIndex        =   16
            Top             =   75
            Visible         =   0   'False
            Width           =   1410
         End
      End
   End
   Begin VB.VScrollBar scrVsc 
      DragIcon        =   "frmDesign.frx":15876
      Height          =   3870
      LargeChange     =   20
      Left            =   8205
      Max             =   100
      MouseIcon       =   "frmDesign.frx":15B80
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2280
      Width           =   250
   End
   Begin VB.HScrollBar scrHsc 
      DragIcon        =   "frmDesign.frx":15E8A
      Height          =   250
      LargeChange     =   20
      Left            =   2790
      Max             =   100
      MouseIcon       =   "frmDesign.frx":16194
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   6165
      Width           =   5400
   End
   Begin VB.PictureBox picRulerH 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   350
      Left            =   2370
      ScaleHeight     =   345
      ScaleWidth      =   6105
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "X����"
      Top             =   1875
      Width           =   6105
   End
   Begin VB.PictureBox picRulerV 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4140
      Left            =   2370
      ScaleHeight     =   4140
      ScaleWidth      =   345
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Y����"
      Top             =   2265
      Width           =   350
   End
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   1530
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   10635
      _ExtentX        =   18759
      _ExtentY        =   2699
      BandCount       =   2
      _CBWidth        =   10635
      _CBHeight       =   1530
      _Version        =   "6.7.9782"
      BandForeColor1  =   255
      Caption1        =   "ϵͳ"
      Child1          =   "tbr1"
      MinHeight1      =   720
      Width1          =   1305
      Key1            =   "System"
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      BandForeColor2  =   16711680
      Caption2        =   "��ʽ"
      Child2          =   "tbr2"
      MinHeight2      =   720
      Width2          =   915
      Key2            =   "Format"
      NewRow2         =   -1  'True
      Begin MSComctlLib.Toolbar tbr2 
         Height          =   720
         Left            =   585
         TabIndex        =   9
         Top             =   780
         Width           =   9960
         _ExtentX        =   17568
         _ExtentY        =   1270
         ButtonWidth     =   1138
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "img��ɫ"
         HotImageList    =   "img��ɫ"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   17
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�����"
               Key             =   "Left"
               Description     =   "�����"
               Object.ToolTipText     =   "ѡ����Ŀ�����"
               Object.Tag             =   "�����"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�Ҷ���"
               Key             =   "Right"
               Description     =   "�Ҷ���"
               Object.ToolTipText     =   "ѡ����Ŀ�Ҷ���"
               Object.Tag             =   "�Ҷ���"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�϶���"
               Key             =   "Up"
               Description     =   "�϶���"
               Object.ToolTipText     =   "ѡ����Ŀ�϶���"
               Object.Tag             =   "�϶���"
               ImageIndex      =   13
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�¶���"
               Key             =   "Down"
               Description     =   "�¶���"
               Object.ToolTipText     =   "ѡ����Ŀ�¶���"
               Object.Tag             =   "�¶���"
               ImageIndex      =   14
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�����"
               Key             =   "Hsc"
               Description     =   "�����"
               Object.ToolTipText     =   "ѡ����Ŀ�������"
               Object.Tag             =   "�����"
               ImageIndex      =   15
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "������"
               Key             =   "Vsc"
               Description     =   "������"
               Object.ToolTipText     =   "ѡ����Ŀ�������"
               Object.Tag             =   "������"
               ImageIndex      =   16
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ͬ���"
               Key             =   "Width"
               Description     =   "ͬ���"
               Object.ToolTipText     =   "ѡ����Ŀ�����ͬ"
               Object.Tag             =   "ͬ���"
               ImageIndex      =   17
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ͬ�߶�"
               Key             =   "Height"
               Description     =   "ͬ�߶�"
               Object.ToolTipText     =   "ѡ����Ŀ�߶���ͬ"
               Object.Tag             =   "ͬ�߶�"
               ImageIndex      =   18
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ͬ���"
               Key             =   "WH"
               Description     =   "ͬ���"
               Object.ToolTipText     =   "ѡ����Ŀ�����ͬ"
               Object.Tag             =   "ͬ���"
               ImageIndex      =   19
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�����"
               Key             =   "VscSpace"
               Description     =   "�����"
               Object.ToolTipText     =   "����ѡ����Ŀ������"
               Object.Tag             =   "�����"
               ImageIndex      =   20
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   3
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "VscSame"
                     Object.Tag             =   "��ͬ(&S)"
                     Text            =   "��ͬ(&S)"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "VscAdd"
                     Object.Tag             =   "����(&A)"
                     Text            =   "����(&A)"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "VscDec"
                     Object.Tag             =   "����(&D)"
                     Text            =   "����(&D)"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "HscSpace"
               Description     =   "����"
               Object.ToolTipText     =   "����ѡ����Ŀ������"
               Object.Tag             =   "����"
               ImageIndex      =   21
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   3
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "HscSame"
                     Object.Tag             =   "��ͬ(&S)"
                     Text            =   "��ͬ(&S)"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "HscAdd"
                     Object.Tag             =   "����(&A)"
                     Text            =   "����(&A)"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "HscDec"
                     Object.Tag             =   "����(&D)"
                     Text            =   "����(&D)"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Scale"
               Description     =   "����"
               Object.ToolTipText     =   "����ҳ�����ʾ����"
               Object.Tag             =   "����"
               ImageIndex      =   22
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   9
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Page"
                     Object.Tag             =   "��ҳ��ʾ(&P)"
                     Text            =   "��ҳ��ʾ(&P)"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Width"
                     Object.Tag             =   "��Ӧ���(&W)"
                     Text            =   "��Ӧ���(&W)"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Height"
                     Object.Tag             =   "��Ӧ�߶�(&H)"
                     Text            =   "��Ӧ�߶�(&H)"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "-Menu1"
                     Object.Tag             =   "-"
                     Text            =   "-"
                  EndProperty
                  BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Scale200"
                     Object.Tag             =   "200%"
                     Text            =   "200%"
                  EndProperty
                  BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Scale100"
                     Object.Tag             =   "100%"
                     Text            =   "100%"
                  EndProperty
                  BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Scale75"
                     Object.Tag             =   "75%"
                     Text            =   "75%"
                  EndProperty
                  BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Scale50"
                     Object.Tag             =   "50%"
                     Text            =   "50%"
                  EndProperty
                  BeginProperty ButtonMenu9 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Scale25"
                     Object.Tag             =   "25%"
                     Text            =   "25%"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Lock"
               Description     =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   23
               Style           =   1
               Value           =   1
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tbr1 
         Height          =   720
         Left            =   585
         TabIndex        =   8
         Top             =   30
         Width           =   9960
         _ExtentX        =   17568
         _ExtentY        =   1270
         ButtonWidth     =   1138
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "img��ɫ"
         HotImageList    =   "img��ɫ"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   21
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ִ��"
               Key             =   "Report"
               Description     =   "ִ��"
               Object.ToolTipText     =   "ִ�б���"
               Object.Tag             =   "ִ��"
               ImageKey        =   "Report"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Save"
               Description     =   "����"
               Object.ToolTipText     =   "���汨��"
               Object.Tag             =   "����"
               ImageKey        =   "Save"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��"
               Key             =   "Guide"
               Description     =   "��"
               Object.ToolTipText     =   "������"
               Object.Tag             =   "��"
               ImageKey        =   "Guide"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ҳ��"
               Key             =   "Page"
               Description     =   "ҳ��"
               Object.ToolTipText     =   "ҳ������"
               Object.Tag             =   "ҳ��"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�¸�ʽ"
               Key             =   "AddFormat"
               Object.ToolTipText     =   "������һ�ֱ����ʽ"
               Object.Tag             =   "�¸�ʽ"
               ImageKey        =   "AddFormat"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ɾ��ʽ"
               Key             =   "DelFormat"
               Object.ToolTipText     =   "ɾ����ǰ�����ʽ"
               Object.Tag             =   "ɾ��"
               ImageKey        =   "DelFormat"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "������"
               Key             =   "New"
               Description     =   "������"
               Object.ToolTipText     =   "����������Դ"
               Object.Tag             =   "������"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�޸�"
               Key             =   "Modi"
               Description     =   "�޸�"
               Object.ToolTipText     =   "�޸ĵ�ǰ����Դ"
               Object.Tag             =   "�޸�"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ɾ��Դ"
               Key             =   "Del"
               Description     =   "ɾ��"
               Object.ToolTipText     =   "ɾ����ǰ����Դ"
               Object.Tag             =   "ɾ��"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Data_"
               Style           =   3
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��Ԫ��"
               Key             =   "Item"
               Description     =   "��Ԫ��"
               Object.ToolTipText     =   "���ӱ���Ԫ��"
               Object.Tag             =   "��Ԫ��"
               ImageIndex      =   6
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   8
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Line"
                     Object.Tag             =   "����"
                     Text            =   "����"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Frame"
                     Object.Tag             =   "����"
                     Text            =   "����"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "_1"
                     Object.Tag             =   "-"
                     Text            =   "-"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Label"
                     Object.Tag             =   "��ǩ"
                     Text            =   "��ǩ"
                  EndProperty
                  BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Picture"
                     Object.Tag             =   "ͼƬ"
                     Text            =   "ͼƬ"
                  EndProperty
                  BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Table"
                     Object.Tag             =   "���"
                     Text            =   "���"
                  EndProperty
                  BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Chart"
                     Object.Tag             =   "ͼ��"
                     Text            =   "ͼ��"
                  EndProperty
                  BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "BarCode"
                     Object.Tag             =   "����"
                     Text            =   "����"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ɾԪ��"
               Key             =   "Remove"
               Description     =   "ɾ��"
               Object.ToolTipText     =   "ɾ��Ʊ����Ŀ"
               Object.Tag             =   "ɾ��"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Description     =   "����"
               Object.ToolTipText     =   "��������"
               Object.Tag             =   "����"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Quit"
               Description     =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageIndex      =   10
            EndProperty
         EndProperty
         Begin MSComctlLib.Toolbar tb2 
            Height          =   720
            Left            =   6200
            TabIndex        =   43
            Top             =   0
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   1270
            ButtonWidth     =   1455
            ButtonHeight    =   1270
            AllowCustomize  =   0   'False
            Wrappable       =   0   'False
            Style           =   1
            ImageList       =   "img��ɫ"
            HotImageList    =   "img��ɫ"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   1
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "��ʷ��¼"
                  Key             =   "History"
                  Description     =   "��ʷ��¼"
                  Object.ToolTipText     =   "��ʷ��¼"
                  Object.Tag             =   "��ʷ��¼"
                  ImageKey        =   "Guide"
               EndProperty
            EndProperty
         End
      End
   End
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   6
      ToolTipText     =   "����������Ϣ��ҵ��˾"
      Top             =   6510
      Width           =   10635
      _ExtentX        =   18759
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmDesign.frx":162E6
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12912
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   953
            MinWidth        =   88
            Text            =   "λ��"
            TextSave        =   "λ��"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   35
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
            Object.ToolTipText     =   "��ǰ���ּ�״̬"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   35
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
            Object.ToolTipText     =   "��ǰ��д��״̬"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imgLarge 
      Left            =   255
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":16B7A
            Key             =   "Fields"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":16E94
            Key             =   "Field"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":171AE
            Key             =   "Bill"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":174C8
            Key             =   "Not"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgSmall 
      Left            =   885
      Top             =   375
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":18312
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":188AC
            Key             =   "Field"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":18E46
            Key             =   "Bill"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":18FA0
            Key             =   "Fields"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdg 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontName        =   "����"
      FontSize        =   9
      Min             =   9
   End
   Begin MSComctlLib.ImageList img��ɫ 
      Left            =   165
      Top             =   855
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   26
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":1953A
            Key             =   "Page"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":19754
            Key             =   "Report"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":1996E
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":1A068
            Key             =   "Modi"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":1A762
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":1AE5C
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":1B556
            Key             =   "Remove"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":1B770
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":1B98A
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":1BBA4
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":1BDBE
            Key             =   "Left"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":1C4B8
            Key             =   "Right"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":1CBB2
            Key             =   "Up"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":1D2AC
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":1D9A6
            Key             =   "Hsc"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":1E0A0
            Key             =   "Vsc"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":1E79A
            Key             =   "Width"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":1EE94
            Key             =   "Height"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":1F58E
            Key             =   "WH"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":1FC88
            Key             =   "VscSpace"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":20382
            Key             =   "HscSpace"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":20A7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":20C96
            Key             =   "Lock"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":21390
            Key             =   "Guide"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":215AA
            Key             =   "AddFormat"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":217C4
            Key             =   "DelFormat"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img��ɫ 
      Left            =   810
      Top             =   885
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   26
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":219DE
            Key             =   "Page"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":21BF8
            Key             =   "Report"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":21E12
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":2250C
            Key             =   "Modi"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":22C06
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":23300
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":239FA
            Key             =   "Remove"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":23C14
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":23E2E
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":24048
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":24262
            Key             =   "Left"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":2495C
            Key             =   "Right"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":25056
            Key             =   "Up"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":25750
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":25E4A
            Key             =   "Hsc"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":26544
            Key             =   "Vsc"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":26C3E
            Key             =   "Width"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":27338
            Key             =   "Height"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":27A32
            Key             =   "WH"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":2812C
            Key             =   "VscSpace"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":28826
            Key             =   "HscSpace"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":28F20
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":2913A
            Key             =   "Lock"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":29834
            Key             =   "Guide"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":29A4E
            Key             =   "AddFormat"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":29C68
            Key             =   "DelFormat"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFile_Report 
         Caption         =   "ִ�б���(&E)"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFile_Save 
         Caption         =   "���汨��(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFile_Guide 
         Caption         =   "������(&G)"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuEdit_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_Page 
         Caption         =   "ҳ������(&S)"
      End
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_Quit 
         Caption         =   "�˳�(&X)"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuEdit_Copy 
         Caption         =   "����Ԫ��(&C)"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuEdit_Paste 
         Caption         =   "ճ��Ԫ��(&P)"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuEdit_7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_SelAll 
         Caption         =   "ȫ��ѡ��(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEdit_Inverse 
         Caption         =   "����ѡ��(&I)"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuEdit_4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_AddFormat 
         Caption         =   "���ӱ����ʽ(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuEdit_DelFormat 
         Caption         =   "ɾ�������ʽ(&E)"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuEdit_5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_New 
         Caption         =   "��������Դ(&N)"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuEdit_Modi 
         Caption         =   "�޸�����Դ(&M)"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuEdit_Del 
         Caption         =   "ɾ������Դ(&D)"
      End
      Begin VB.Menu mnuEdit_History 
         Caption         =   "��ʷ��¼(&H)"
      End
      Begin VB.Menu mnuEdit_Data_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Item 
         Caption         =   "����Ԫ��(&T)"
         Begin VB.Menu mnuEdit_ItemAdd 
            Caption         =   "����(&L)"
            Index           =   0
            Shortcut        =   ^L
         End
         Begin VB.Menu mnuEdit_ItemAdd 
            Caption         =   "����(&F)"
            Index           =   1
            Shortcut        =   ^T
         End
         Begin VB.Menu mnuEdit_ItemAdd 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu mnuEdit_ItemAdd 
            Caption         =   "��ǩ(&D)"
            Index           =   3
            Shortcut        =   ^D
         End
         Begin VB.Menu mnuEdit_ItemAdd 
            Caption         =   "ͼƬ(&P)"
            Index           =   4
            Shortcut        =   ^P
         End
         Begin VB.Menu mnuEdit_ItemAdd 
            Caption         =   "���(&B)"
            Index           =   5
            Shortcut        =   ^B
         End
         Begin VB.Menu mnuEdit_ItemAdd 
            Caption         =   "ͼ��(&C)"
            Index           =   6
            Shortcut        =   ^H
         End
         Begin VB.Menu mnuEdit_ItemAdd 
            Caption         =   "����(&R)"
            Index           =   7
            Shortcut        =   ^R
         End
      End
      Begin VB.Menu mnuEdit_Remove 
         Caption         =   "ɾ��Ԫ��(&R)"
         Shortcut        =   {DEL}
      End
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "��ʽ(&R)"
      Begin VB.Menu mnuFormat_Order 
         Caption         =   "���˳��(&O)"
         Begin VB.Menu mnuFormat_Front 
            Caption         =   "��ǰ(&F)"
         End
         Begin VB.Menu mnuFormat_Back 
            Caption         =   "�ú�(&B)"
         End
      End
      Begin VB.Menu mnuFormat_4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormat_Align 
         Caption         =   "����(&A)"
         Begin VB.Menu mnuFormat_DoAlign 
            Caption         =   "�����(&L)"
            Index           =   0
         End
         Begin VB.Menu mnuFormat_DoAlign 
            Caption         =   "�Ҷ���(&R)"
            Index           =   1
         End
         Begin VB.Menu mnuFormat_DoAlign 
            Caption         =   "�϶���(&U)"
            Index           =   2
         End
         Begin VB.Menu mnuFormat_DoAlign 
            Caption         =   "�¶���(&D)"
            Index           =   3
         End
         Begin VB.Menu mnuFormat_DoAlign 
            Caption         =   "�����(&C)"
            Index           =   4
         End
         Begin VB.Menu mnuFormat_DoAlign 
            Caption         =   "������(&M)"
            Index           =   5
         End
         Begin VB.Menu mnuFormat_DoAlign 
            Caption         =   "-"
            Index           =   6
         End
         Begin VB.Menu mnuFormat_DoAlign 
            Caption         =   "ˮƽ����(&H)"
            Index           =   7
         End
         Begin VB.Menu mnuFormat_DoAlign 
            Caption         =   "��ֱ����(&V)"
            Index           =   8
         End
      End
      Begin VB.Menu mnuFormat_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormat_Size 
         Caption         =   "�ߴ�(&S)"
         Begin VB.Menu mnuFormat_Width 
            Caption         =   "ͬ���(&W)"
         End
         Begin VB.Menu mnuFormat_Height 
            Caption         =   "ͬ�߶�(&H)"
         End
         Begin VB.Menu mnuFormat_WH 
            Caption         =   "ͬ���(&B)"
         End
      End
      Begin VB.Menu mnuFomrat_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormat_VscSapce 
         Caption         =   "�����(&V)"
         Begin VB.Menu mnuFormat_VscSpace_Same 
            Caption         =   "��ͬ(&S)"
         End
         Begin VB.Menu mnuFormat_VscSpace_Add 
            Caption         =   "����(&A)"
         End
         Begin VB.Menu mnuFormat_VscSpace_Dec 
            Caption         =   "����(&D)"
         End
      End
      Begin VB.Menu mnuFormat_HscSpace 
         Caption         =   "����(&H)"
         Begin VB.Menu mnuFormat_HscSpace_Same 
            Caption         =   "��ͬ(&S)"
         End
         Begin VB.Menu mnuFormat_HscSpace_Add 
            Caption         =   "����(&A)"
         End
         Begin VB.Menu mnuFormat_HscSpace_Dec 
            Caption         =   "����(&D)"
         End
      End
      Begin VB.Menu mnuFormat_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormat_Lock 
         Caption         =   "����Ԫ��(&L)"
         Checked         =   -1  'True
         Shortcut        =   ^K
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "��ͼ(&V)"
      Begin VB.Menu mnuViewTool 
         Caption         =   "������(&T)"
         Begin VB.Menu mnuViewToolSystem 
            Caption         =   "ϵͳ����(&B)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewToolFormat 
            Caption         =   "��ʽ����(&F)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuView_1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewToolText 
            Caption         =   "�ı���ǩ(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewScale 
         Caption         =   "��ʾ����(&C)"
         Begin VB.Menu mnuViewScaleMode 
            Caption         =   "��ҳ��ʾ(&P)"
            Index           =   0
         End
         Begin VB.Menu mnuViewScaleMode 
            Caption         =   "��Ӧ���(&W)"
            Index           =   1
         End
         Begin VB.Menu mnuViewScaleMode 
            Caption         =   "��Ӧ�߶�(&H)"
            Index           =   2
         End
         Begin VB.Menu mnuViewScaleMode 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnuViewScaleMode 
            Caption         =   "200%"
            Checked         =   -1  'True
            Index           =   4
         End
         Begin VB.Menu mnuViewScaleMode 
            Caption         =   "100%"
            Checked         =   -1  'True
            Index           =   5
         End
         Begin VB.Menu mnuViewScaleMode 
            Caption         =   "75%"
            Checked         =   -1  'True
            Index           =   6
         End
         Begin VB.Menu mnuViewScaleMode 
            Caption         =   "50%"
            Checked         =   -1  'True
            Index           =   7
         End
         Begin VB.Menu mnuViewScaleMode 
            Caption         =   "25%"
            Checked         =   -1  'True
            Index           =   8
         End
      End
      Begin VB.Menu mnuView_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewToolAttrib 
         Caption         =   "���Ա��(&A)"
         Checked         =   -1  'True
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuViewToolSQL 
         Caption         =   "����Դ��(&L)"
         Checked         =   -1  'True
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuViewToolRuler 
         Caption         =   "������(&U)"
         Checked         =   -1  'True
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "״̬��(&S)"
         Checked         =   -1  'True
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuView_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuView_reFlash 
         Caption         =   "ˢ��(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuHelpTitle 
         Caption         =   "��������(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "&WEB�ϵ�����"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "������ҳ(&H)"
         End
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "������̳(&F)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "���ͷ���(&K)..."
         End
      End
      Begin VB.Menu mnuHelp_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)..."
      End
   End
   Begin VB.Menu mnuCustom 
      Caption         =   "������"
      Visible         =   0   'False
      Begin VB.Menu mnuCustom_Head 
         Caption         =   "��ͷ����"
         Begin VB.Menu mnuCustom_Head_Insert 
            Caption         =   "�����ͷ��(&I)"
            Begin VB.Menu mnuCustom_Head_Insert_UP 
               Caption         =   "��ǰ������(&U)"
            End
            Begin VB.Menu mnuCustom_Head_Insert_Down 
               Caption         =   "��ǰ������(&D)"
            End
         End
         Begin VB.Menu mnuCustom_Head_Del 
            Caption         =   "ɾ����ͷ��(&D)"
         End
         Begin VB.Menu mnuCustom_Head_Auto 
            Caption         =   "�Զ����к�(&N)"
         End
         Begin VB.Menu mnuCustom_Head_Clear 
            Caption         =   "��ձ�ͷ(&C)"
         End
         Begin VB.Menu mnuCustom_Head_1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuCustom_Head_Text 
            Caption         =   "��Ԫ������(&T)"
         End
         Begin VB.Menu mnuCustom_Head_Merge 
            Caption         =   "��Ԫ��ϲ�(&M)"
         End
         Begin VB.Menu mnuCustom_Head_Split 
            Caption         =   "��Ԫ����(&S)"
         End
      End
      Begin VB.Menu mnuCustom_Col 
         Caption         =   "���в���"
         Begin VB.Menu mnuCustom_Col_Insert 
            Caption         =   "�������(&I)"
            Begin VB.Menu mnuCustom_Col_Insert_Left 
               Caption         =   "��ǰ������(&L)"
            End
            Begin VB.Menu mnuCustom_Col_Insert_Right 
               Caption         =   "��ǰ������(&R)"
            End
         End
         Begin VB.Menu mnuCustom_Col_Del 
            Caption         =   "ɾ������(&R)"
         End
         Begin VB.Menu mnuCustom_Col_Clear 
            Caption         =   "��ձ���(&C)"
         End
         Begin VB.Menu mnuCustom_Col_1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuCustom_Col_Data 
            Caption         =   "��������(&D)"
         End
         Begin VB.Menu mnuCustom_Col_State 
            Caption         =   "���л���(&S)"
            Begin VB.Menu mnuCustom_Col_State_Style 
               Caption         =   "��(&0)"
               Checked         =   -1  'True
               Index           =   0
            End
            Begin VB.Menu mnuCustom_Col_State_Style 
               Caption         =   "���(&1)"
               Index           =   1
            End
            Begin VB.Menu mnuCustom_Col_State_Style 
               Caption         =   "��ƽ��ֵ(&2)"
               Index           =   2
            End
            Begin VB.Menu mnuCustom_Col_State_Style 
               Caption         =   "�����ֵ(&3)"
               Index           =   3
            End
            Begin VB.Menu mnuCustom_Col_State_Style 
               Caption         =   "����Сֵ(&4)"
               Index           =   4
            End
            Begin VB.Menu mnuCustom_Col_State_Style 
               Caption         =   "���¼��(&5)"
               Index           =   5
            End
         End
         Begin VB.Menu mnuCustom_Col_Align 
            Caption         =   "���ж���(&A)"
            Begin VB.Menu mnuCustom_Col_Align_Style 
               Caption         =   "�����(&L)"
               Checked         =   -1  'True
               Index           =   0
            End
            Begin VB.Menu mnuCustom_Col_Align_Style 
               Caption         =   "���ж���(&M)"
               Index           =   1
            End
            Begin VB.Menu mnuCustom_Col_Align_Style 
               Caption         =   "�Ҷ���(&R)"
               Index           =   2
            End
         End
      End
   End
   Begin VB.Menu mnuClass 
      Caption         =   "������"
      Visible         =   0   'False
      Begin VB.Menu mnuClass_Insert 
         Caption         =   "��������Ŀ(&I)"
         Begin VB.Menu mnuClass_Insert_Before 
            Caption         =   "�ڵ�ǰ��֮ǰ(&B)"
         End
         Begin VB.Menu mnuClass_Insert_After 
            Caption         =   "�ڵ�ǰ��֮��(&A)"
         End
      End
      Begin VB.Menu mnuClass_Data 
         Caption         =   "��������(&D)"
      End
      Begin VB.Menu mnuClass_ExChange 
         Caption         =   "���жԻ�(&E)"
      End
      Begin VB.Menu mnuClass_Del 
         Caption         =   "ɾ������(&R)"
      End
      Begin VB.Menu mnuClass_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClass_State 
         Caption         =   "�������(&S)"
         Begin VB.Menu mnuClass_State_Style 
            Caption         =   "��(&0)"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuClass_State_Style 
            Caption         =   "���(&1)"
            Index           =   1
         End
         Begin VB.Menu mnuClass_State_Style 
            Caption         =   "��ƽ��ֵ(&2)"
            Index           =   2
         End
         Begin VB.Menu mnuClass_State_Style 
            Caption         =   "�����ֵ(&3)"
            Index           =   3
         End
         Begin VB.Menu mnuClass_State_Style 
            Caption         =   "����Сֵ(&4)"
            Index           =   4
         End
         Begin VB.Menu mnuClass_State_Style 
            Caption         =   "���¼��(&5)"
            Index           =   5
         End
      End
      Begin VB.Menu mnuClass_Align 
         Caption         =   "�������(&A)"
         Begin VB.Menu mnuClass_Align_Style 
            Caption         =   "�����(&L)"
            Index           =   0
         End
         Begin VB.Menu mnuClass_Align_Style 
            Caption         =   "�м����(&M)"
            Index           =   1
         End
         Begin VB.Menu mnuClass_Align_Style 
            Caption         =   "�Ҷ���(&R)"
            Checked         =   -1  'True
            Index           =   2
         End
      End
   End
End
Attribute VB_Name = "frmDesign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public lngRPTID As Long '�룺Ҫ��Ƶı���ID
Public mblnNotModiData As Boolean '�룺�Ƿ񲻿����޸�����Դ

Private objReport As Report 'Ҫ��Ƶı������
Private blnLock As Boolean
Private blnMax As Boolean
Private bytLine As Byte
Private blnDown As Boolean
Private lngPreX As Long, lngPreY As Long
Private preHsc As Long, preVsc As Long '���λ�ü�¼
Private bytCurTool As Byte '��ǰѡ��Ԫ��(0=ѡ��;1=����;2=����;3=��ǩ;4=ͼƬ;5=���;6=ͼ��;7=����,8=��Ƭ)
Private selArea As RECT '���ѡ��ľ��ο�
Private selCell As Cells '���ѡ��ı�ͷ��Χ
Private drgCell As Cells '�϶������еĵ�Ԫ��Χ,drgCell.Row����
Private intCurCol As Integer '��ǰѡ����������
Private objFont As New clsRotateFont '��ת�������
Private intMaxID As Integer '��ǰ���ؼ�����(��1��ʼ)
Private intCurID As Integer '��ǰѡ��ؼ�����(��1��ʼ)
Private BlnSave As Boolean
Private objLastSel As Object '���һ��ѡ�е�Ԫ�ؿؼ�
Private blnDrop As Boolean, blnHead As Boolean, blnSum As Boolean
Private strMenu As String
Private mblnFirst As Boolean
Private mobjMove As Object  'Ԫ������ĸ��ؼ�
Private mlngX As Long   '���븸�ؼ���λ��
Private mlngY As Long   '���븸�ؼ���λ��
Private mobjPicMERGE As IPictureDisp
Private mobjPicMove As IPictureDisp

'zyb#Add
Private blnDelReportFormat As Boolean   '�����Ƿ��ǹ̶�����(�̶���������ɾ��������ʽ)
Private mbytCurrFmt As Byte            'ѡ��ı�����ʽ(�����޸ı�����ʽ������)
Private blnAllowIn As Boolean           '�Ƿ����Change�¼�
Private blnRefresh As Boolean           '����ˢ��
Private blnModify As Boolean            '�Ƿ��������
Private blnAdjustRowHeight As Boolean   '����ı�̶��е��и�
Private blnAdjustColWidth As Boolean    '����ı������е��п�
Private sgnMode As Single
Private sgnLastMode As Single
Private Type WindowProperty
        l As Single
        H As Single
        T As Single
        W As Single
End Type
Private WinProperty As WindowProperty

Private Sub cboAtt_Click()
    Dim ItemThis As RPTItem, ItemSend As RPTItem, ItemFmt As RPTFmt
    Dim str���� As String, intID As Integer
    Dim strCurText As String, intType As Integer
    Dim objBarCode As StdPicture, lngSize As Long
    Dim strBarCode As String, sngWidth As Single
    Dim blnSeek As Boolean
    Dim k As Long, X As Long, Y As Long, tmpID As RelatID, ItemTmp As RPTItem
    Dim StrCompare As String, lngX As Long, lngY As Long
    Dim tmpItem As RPTItem, strSouse As String
    Dim j As Long, i As Long, blnYes As Boolean
    Dim tmpObj As PictureBox

    strCurText = mshAtt.TextMatrix(mshAtt.Row, 0)
    If intCurID = 0 And InStr(1, "���ͼ��,����Ԫ��", strCurText) = 0 Then Exit Sub '2002-03-26
    
    If intCurID <> 0 Then
        intType = objReport.Items("_" & intCurID).����
    End If
    Select Case strCurText
        Case "����"
            mshAtt.TextMatrix(mshAtt.Row, 1) = cboAtt.Text
            '�����ѡ�е�����Ԫ�ض�����ͬ��Ԫ��
            If lblSize.count > 9 Then
                For Each tmpObj In lblSize
                    If tmpObj.Index Mod 8 = 1 Then
                        lbl(tmpObj.Tag).Alignment = IIF(cboAtt.ListIndex <> 0, IIF(cboAtt.ListIndex = 1, 2, 1), 0)
                        objReport.Items("_" & tmpObj.Tag).���� = cboAtt.ListIndex
                    End If
                Next
            Else
                lbl(intCurID).Alignment = IIF(cboAtt.ListIndex <> 0, IIF(cboAtt.ListIndex = 1, 2, 1), 0)
                objReport.Items("_" & intCurID).���� = cboAtt.ListIndex
            End If
            BlnSave = False
        Case "����Ԫ��"
            If cboAtt.Text = "" Then Exit Sub
            'Call SelClear
            Call SelItem(cboAtt.ItemData(cboAtt.ListIndex), True)
            Call ShowAttrib(cboAtt.ItemData(cboAtt.ListIndex))
            BlnSave = False
        Case "���ͼ��"
            '���ļ���
            If blnModify = False Then Exit Sub
            For Each ItemFmt In objReport.Fmts
                If ItemFmt.��� = mbytCurrFmt Then
                    ItemFmt.ͼ�� = cboAtt.ItemData(cboAtt.ListIndex)
                    Exit For
                End If
            Next
            mshAtt.TextMatrix(mshAtt.Row, 1) = cboAtt.Text
            BlnSave = False
        Case "���ն���"
            If objReport.Items("_" & intCurID).���� = cboAtt.Text Then Exit Sub
            If (objReport.Items("_" & intCurID).���� = 4 Or objReport.Items("_" & intCurID).���� = 5) And cboAtt.Text <> "" Then
                If objReport.Items("_" & GetDependID(cboAtt.Text)).��ID <> 0 Then
                    MsgBox "��Ƭ�ڵı���������ø��ӱ��", vbInformation, App.Title
                    cboAtt.ListIndex = -1: cboAtt.SetFocus: Exit Sub
                End If
            End If
            Call CopyItem(ItemSend, objReport.Items("_" & intCurID))
            mshAtt.TextMatrix(mshAtt.Row, 1) = cboAtt.Text
            objReport.Items("_" & intCurID).���� = cboAtt.Text
            objReport.Items("_" & intCurID).���� = 0
            
            '������ն�����,���丽����Ŀ[����,����]Ϊ"����"
            If cboAtt.Text = "" Then
                With mshAtt
                    .TextMatrix(GetRow("����"), 1) = "����"
                    objReport.Items("_" & intCurID).���� = "0"
                End With
            Else
                With mshAtt
                    .TextMatrix(GetRow("���ն���"), 1) = objReport.Items("_" & intCurID).����
                    Select Case objReport.Items("_" & intCurID).����
                    Case 2
                        str���� = IIF(objReport.Items("_" & intCurID).���� = "0", "11", objReport.Items("_" & intCurID).����)
                        objReport.Items("_" & intCurID).���� = str����
                        .TextMatrix(GetRow("����"), 1) = IIF(Mid(str����, 1, 1) <> "2", "������", "������")
                        str���� = IIF(str���� = "0" Or str���� = "", "����", IIF(Mid(str����, 2) = "1", "����", IIF(Mid(str����, 2) = "2", "����", "����")))
                        .TextMatrix(GetRow("����"), 1) = str����
                    Case 4, 5
                        str���� = ""
                        For Each ItemThis In objReport.Items
                            If ItemThis.���� <> "" And ItemThis.��ʽ�� = mbytCurrFmt And ItemThis.���� = mshAtt.TextMatrix(GetRow("���ն���"), 1) And InStr(1, "4,5", ItemThis.����) <> 0 Then
                                If ItemThis.���� = 0 Then
                                    str���� = IIF(ItemThis.���� = 4, "����", "������")
                                Else
                                    str���� = IIF(ItemThis.���� = 1, "����", "������")
                                End If
                                Exit For
                            End If
                        Next
                        
                        If str���� = "" Then str���� = IIF(objReport.Items("_" & intCurID).���� = 4, "����", "������")
                        .TextMatrix(GetRow("����"), 1) = str����
                        objReport.Items("_" & intCurID).���� = IIF(str���� = "����", 1, 2)
                    End Select
                End With
                If objReport.Items("_" & GetDependID(cboAtt.Text)).��ID <> 0 Then
                    If objReport.Items("_" & intCurID).��ID = 0 Then
                        objReport.Items("_" & intCurID).��ID = objReport.Items("_" & GetDependID(cboAtt.Text)).��ID
                    End If
                Else
                    If objReport.Items("_" & intCurID).��ID <> 0 Then
                        objReport.Items("_" & intCurID).��ID = 0
                    End If
                End If
                Dim ParentItem As RPTItem
                For Each ParentItem In objReport.Items
                    If ParentItem.��ʽ�� = mbytCurrFmt And ParentItem.���� = objReport.Items("_" & intCurID).���� And ParentItem.���� = 5 And objReport.Items("_" & intCurID).���� = 5 Then
                        Call SetGridLike(msh(ParentItem.Key), msh(intCurID))
                        Exit For
                    End If
                Next
            End If
            
            Call ReferTo(ItemSend)
            If objReport.Items("_" & intCurID).ϵͳ Then Call AdjustCoordinate
            
            BlnSave = False
            
            If GetSelNum = 1 Then Call ShowAttrib(intCurID)
        Case "����Դ"
            If objReport.Items("_" & intCurID).����Դ = cboAtt.Text Then Exit Sub
            If Trim(cboAtt.Text) <> "" Then
                For Each ItemTmp In objReport.Items
                    If ItemTmp.��ID <> 0 And ItemTmp.��ID = intCurID Then
                        If ItemTmp.���� = 4 Then
                            For Each tmpID In ItemTmp.SubIDs
                                With objReport.Items("_" & tmpID.id)
                                    X = InStr(1, .����, "]")
                                    Y = InStr(1, .����, ".")
                                    k = InStr(1, .����, "[")
                                    If X > k And X > Y And X <> 0 And k <> 0 Then
                                        If Mid(.����, k + 1, Y - k - 1) <> cboAtt.Text Then
                                            MsgBox "��Ƭ�еı��󶨵������б�������ѡ������Դ�����飡", vbInformation, App.Title
                                            Call CboSetText(cboAtt, objReport.Items("_" & intCurID).����Դ)
                                            Exit Sub
                                        End If
                                    End If
                                End With
                            Next
                        End If
                    End If
                Next
            End If

            
            mshAtt.TextMatrix(mshAtt.Row, 1) = cboAtt.Text
            objReport.Items("_" & intCurID).����Դ = cboAtt.Text
            
            BlnSave = False
            
            If GetSelNum = 1 Then Call ShowAttrib(intCurID)
        Case "����"
            If objReport.Items("_" & intCurID).��ID = 0 Then
                StrCompare = "ҳ��"
            Else
                StrCompare = objReport.Items("_" & objReport.Items("_" & intCurID).��ID).����
            End If
            If StrCompare = cboAtt.Text Then Exit Sub
            '����Ƿ��Ƿ������
            If objReport.Items("_" & intCurID).���� = 4 And cboAtt.Text <> "ҳ��" Then
                If objReport.Items("_" & intCurID).���� > 1 Then
                    MsgBox "��Ƭ�в������������ı��", vbInformation, App.Title
                    Exit Sub
                End If
                '��Ƭ�ڲ������ӱ��
                For Each tmpItem In objReport.Items
                    If tmpItem.��ʽ�� = mbytCurrFmt Then
                        If tmpItem.���� = 5 Or tmpItem.���� = 4 Then
                            If tmpItem.���� = objReport.Items("_" & intCurID).���� Then
                                 MsgBox "������ڸ��ӱ�񣬲�������뿨Ƭ�У�", vbInformation, App.Title
                                 Exit Sub
                            End If
                        End If
                    End If
                Next
                '�����Ƭ������Դ�������������Դ�Ƿ�ƥ��
                If objReport.Items("_" & Val(cboAtt.ItemData(cboAtt.ListIndex) & "")).����Դ <> "" Then
                    For Each tmpID In objReport.Items("_" & intCurID).SubIDs
                        With objReport.Items("_" & tmpID.id)
                            i = InStr(1, .����, "]")
                            j = InStr(1, .����, ".")
                            k = InStr(1, .����, "[")
                            If i > k And i > j And i <> 0 And k <> 0 And j <> 0 Then
                                If Mid(.����, k + 1, j - k - 1) <> objReport.Items("_" & Val(cboAtt.ItemData(cboAtt.ListIndex) & "")).����Դ Then
                                    If blnYes = False Then
                                        If MsgBox("��ǰ��Ƭ��������Դ��������е������кͿ�Ƭ����Դ����ͬ�����뽫��ղ�ƥ����У��Ƿ����?", vbQuestion + vbYesNo + vbDefaultButton1, App.Title) = vbYes Then
                                            .���� = ""
                                            msh(intCurID).TextMatrix(1, .���) = ""
                                            blnYes = True
                                        Else
                                            Exit Sub
                                        End If
                                    Else
                                        .���� = ""
                                        msh(intCurID).TextMatrix(1, .���) = ""
                                    End If
                                End If
                            End If
                        End With
                    Next
                Else
                    '��Ƭû������Դ������ʾ�û��Ƿ��������Դ
                    For Each tmpID In objReport.Items("_" & intCurID).SubIDs
                        With objReport.Items("_" & tmpID.id)
                            i = InStr(1, .����, "]")
                            j = InStr(1, .����, ".")
                            k = InStr(1, .����, "[")
                            If i > k And i > j And i <> 0 And k <> 0 And j <> 0 Then
                                If InStr(strSouse, Mid(.����, k + 1, j - k - 1)) = 0 Then
                                    strSouse = strSouse & "," & Mid(.����, k + 1, j - k - 1)
                                End If
                            End If
                        End With
                    Next
                    strSouse = Mid(strSouse, 2)
                    'ֻ��һ������Դʱ����ʾ
                    If InStr(strSouse, ",") = 0 And strSouse <> "" Then
                        If MsgBox("��ǰ��Ƭδ������Դ���󶨺󽫷����ӡ���ſ�Ƭ������Դ�д���""�����ʶ""�ֶ���""�����ʶ""��ͬ��Ϊһ��,����һ������Ϊһ�飻" & vbCrLf & _
                             "������ֻ��ӡһ�ſ�Ƭ���Ƿ������Դ""" & strSouse & """?", vbQuestion + vbYesNo + vbDefaultButton1, App.Title) = vbYes Then
                            objReport.Items("_" & mobjMove.Index).����Դ = strSouse
                        End If
                    End If
                End If
            End If
            mshAtt.TextMatrix(mshAtt.Row, 1) = cboAtt.Text
            If cboAtt.Text = "ҳ��" Then
                objReport.Items("_" & intCurID).X = objReport.Items("_" & intCurID).X + objReport.Items("_" & objReport.Items("_" & intCurID).��ID).X
                objReport.Items("_" & intCurID).Y = objReport.Items("_" & intCurID).Y + objReport.Items("_" & objReport.Items("_" & intCurID).��ID).Y
                objReport.Items("_" & intCurID).��ID = 0
            Else
                If objReport.Items("_" & intCurID).��ID = 0 Then
                    lngX = objReport.Items("_" & Val(cboAtt.ItemData(cboAtt.ListIndex) & "")).X
                    lngY = objReport.Items("_" & Val(cboAtt.ItemData(cboAtt.ListIndex) & "")).Y
                Else
                    lngX = objReport.Items("_" & Val(cboAtt.ItemData(cboAtt.ListIndex) & "")).X - objReport.Items("_" & intCurID).X
                    lngY = objReport.Items("_" & Val(cboAtt.ItemData(cboAtt.ListIndex) & "")).Y - objReport.Items("_" & intCurID).Y
                End If
                objReport.Items("_" & intCurID).X = objReport.Items("_" & intCurID).X - lngX
                objReport.Items("_" & intCurID).Y = objReport.Items("_" & intCurID).Y - lngY
                objReport.Items("_" & intCurID).��ID = Val(cboAtt.ItemData(cboAtt.ListIndex) & "")
            End If
            If objReport.Items("_" & intCurID).���� = 4 Then
                '��������
                For Each tmpID In objReport.Items("_" & intCurID).SubIDs
                    objReport.Items("_" & tmpID.id).��ID = objReport.Items("_" & intCurID).��ID
                Next
            End If
            Call AdjustCoordinate(True)
            BlnSave = False
            If GetSelNum = 1 Then Call ShowAttrib(intCurID)
        Case "����"
            If mshAtt.TextMatrix(mshAtt.Row, 1) = cboAtt.Text Then Exit Sub
            objReport.Items("_" & intCurID).���� = IIF(cboAtt.Text = "������", "1", "2") & Mid(objReport.Items("_" & intCurID).����, 2)
            mshAtt.TextMatrix(mshAtt.Row, 1) = cboAtt.Text
            
            Call CopyItem(ItemSend, objReport.Items("_" & intCurID))
            Call ReferTo(ItemSend)
            If objReport.Items("_" & intCurID).ϵͳ Then Call AdjustCoordinate
            
            BlnSave = False
            
            If GetSelNum = 1 Then Call ShowAttrib(intCurID)
        Case "����"
            If mshAtt.TextMatrix(mshAtt.Row, 1) = cboAtt.Text Then Exit Sub
            Call CopyItem(ItemSend, objReport.Items("_" & intCurID))
            Select Case objReport.Items("_" & intCurID).����
            Case 2
                mshAtt.TextMatrix(mshAtt.Row, 1) = cboAtt.Text
                str���� = IIF(cboAtt.Text = "����", "1", IIF(cboAtt.Text = "����", "2", "3"))
                objReport.Items("_" & intCurID).���� = Mid(objReport.Items("_" & intCurID).����, 1, 1) & str����
            Case 5
                mshAtt.TextMatrix(mshAtt.Row, 1) = cboAtt.Text
                objReport.Items("_" & intCurID).���� = IIF(cboAtt.Text = "����", "1", "2")
            End Select
            
            Call ReferTo(ItemSend)
            If objReport.Items("_" & intCurID).ϵͳ Then Call AdjustCoordinate
            
            BlnSave = False
            
            If GetSelNum = 1 Then Call ShowAttrib(intCurID)
        Case "��������"
            If mshAtt.TextMatrix(mshAtt.Row, 1) = cboAtt.Text Then Exit Sub
            mshAtt.TextMatrix(mshAtt.Row, 1) = cboAtt.Text
            objReport.Items("_" & intCurID).��� = cboAtt.ItemData(cboAtt.ListIndex)
            BlnSave = False
            
            With objReport.Items("_" & intCurID)
                strBarCode = ReplaceBracket(.����)
                If strBarCode = "" Then strBarCode = "1234567890"
                
                Unload frmFlash 'ǿ�Ƴ�ʼPicture����Ȼ�л�����������
                If .��� = 1 Then
                    Set objBarCode = DrawBarCode128(frmFlash.picTemp, 3, strBarCode, Mid(.��ͷ, 1, 1) = "1")
                ElseIf .��� = 2 Then
                    Set objBarCode = DrawBarCode39(frmFlash.picTemp, 3, strBarCode, Mid(.��ͷ, 2, 1) = "1", Mid(.��ͷ, 1, 1) = "1")
                ElseIf .��� = 3 Then
                    If .�и� = 0 Then .�и� = 2
                    Set objBarCode = DrawBarCode128Auto(frmFlash.picTemp, strBarCode, sngWidth, .�и�, Mid(.��ͷ, 1, 1) = "1")
                ElseIf .��� = 10 Then
                    Set objBarCode = DrawBarCode2D(strBarCode, frmFlash.picTemp, lngSize)
                End If
                If Val(Mid(.��ͷ, 3, 1)) <> 0 Then
                    Set objBarCode = PictureSpin(objBarCode, Val(Mid(.��ͷ, 3, 1)), frmFlash.picTemp)
                End If
                Set ImgCode(intCurID).Picture = objBarCode
                
                If .��� = 3 Then
                    '128���Զ��������
                    If Val(Mid(.��ͷ, 3, 1)) = 0 Then
                        ImgCode(intCurID).Width = Format(Me.ScaleX(sngWidth, vbMillimeters, vbTwips) * sgnMode, "0.00")
                        .W = Me.ScaleX(sngWidth, vbMillimeters, vbTwips)
                    Else
                        ImgCode(intCurID).Height = Format(Me.ScaleY(sngWidth, vbMillimeters, vbTwips) * sgnMode, "0.00")
                        .H = Me.ScaleY(sngWidth, vbMillimeters, vbTwips)
                    End If
                    Call SeekItem(ImgCode(intCurID), ImgCode(intCurID).Left, ImgCode(intCurID).Top)
                ElseIf .��� = 10 Then
                    '��ά����ȱʡ�Զ�������С
                    .�Ե� = True
                    ImgCode(intCurID).Width = Format(lngSize * sgnMode, "0.00")
                    ImgCode(intCurID).Height = Format(lngSize * sgnMode, "0.00")
                    .W = lngSize: .H = lngSize
                    
                    Call SeekItem(ImgCode(intCurID), ImgCode(intCurID).Left, ImgCode(intCurID).Top)
                End If
            End With
            
            If GetSelNum = 1 Then Call ShowAttrib(intCurID)
        Case "�����߿�"
            If mshAtt.TextMatrix(mshAtt.Row, 1) = cboAtt.Text Then Exit Sub
            mshAtt.TextMatrix(mshAtt.Row, 1) = cboAtt.Text
            
            With objReport.Items("_" & intCurID)
                .�и� = Val(cboAtt.Text): BlnSave = False
                
                '�ػ�ͼ��
                strBarCode = ReplaceBracket(.����)
                If strBarCode = "" Then strBarCode = "1234567890"
                
                Unload frmFlash 'ǿ�Ƴ�ʼPicture����Ȼ�л�����������
                If .��� = 3 Then
                    Set objBarCode = DrawBarCode128Auto(frmFlash.picTemp, strBarCode, sngWidth, .�и�, Mid(.��ͷ, 1, 1) = "1")
                End If
                If Val(Mid(.��ͷ, 3, 1)) <> 0 Then
                    Set objBarCode = PictureSpin(objBarCode, Val(Mid(.��ͷ, 3, 1)), frmFlash.picTemp)
                End If
                Set ImgCode(intCurID).Picture = objBarCode
                
                If .��� = 3 Then
                    '128���Զ��������
                    If Val(Mid(.��ͷ, 3, 1)) = 0 Then
                        ImgCode(intCurID).Width = Format(Me.ScaleX(sngWidth, vbMillimeters, vbTwips) * sgnMode, "0.00")
                        .W = Me.ScaleX(sngWidth, vbMillimeters, vbTwips)
                    Else
                        ImgCode(intCurID).Height = Format(Me.ScaleY(sngWidth, vbMillimeters, vbTwips) * sgnMode, "0.00")
                        .H = Me.ScaleY(sngWidth, vbMillimeters, vbTwips)
                    End If
                    Call SeekItem(ImgCode(intCurID), ImgCode(intCurID).Left, ImgCode(intCurID).Top)
                End If
            End With
            
            If GetSelNum = 1 Then Call ShowAttrib(intCurID)
        Case "��ת����"
            If mshAtt.TextMatrix(mshAtt.Row, 1) = cboAtt.Text Then Exit Sub
            mshAtt.TextMatrix(mshAtt.Row, 1) = cboAtt.Text
            
            LockWindowUpdate Me.hwnd
            
            With objReport.Items("_" & intCurID)
                blnSeek = False
                
                '��ת֮��任���
                If Val(Mid(.��ͷ, 3, 1)) <> 0 And cboAtt.ListIndex = 0 _
                    Or Val(Mid(.��ͷ, 3, 1)) = 0 And cboAtt.ListIndex <> 0 Then
                    lngSize = .W: .W = .H: .H = lngSize
                    
                    lngSize = ImgCode(intCurID).Width
                    ImgCode(intCurID).Width = ImgCode(intCurID).Height
                    ImgCode(intCurID).Height = lngSize
                    
                    blnSeek = True
                End If
                .��ͷ = SetBit(.��ͷ, 3, cboAtt.ListIndex)
                BlnSave = False
                
                '�ػ�ͼ��
                strBarCode = ReplaceBracket(.����)
                If strBarCode = "" Then strBarCode = "1234567890"
                
                Unload frmFlash 'ǿ�Ƴ�ʼPicture����Ȼ�л�����������
                If .��� = 1 Then
                    Set objBarCode = DrawBarCode128(frmFlash.picTemp, 3, strBarCode, Mid(.��ͷ, 1, 1) = "1")
                ElseIf .��� = 2 Then
                    Set objBarCode = DrawBarCode39(frmFlash.picTemp, 3, strBarCode, Mid(.��ͷ, 2, 1) = "1", Mid(.��ͷ, 1, 1) = "1")
                ElseIf .��� = 3 Then
                    Set objBarCode = DrawBarCode128Auto(frmFlash.picTemp, strBarCode, sngWidth, .�и�, Mid(.��ͷ, 1, 1) = "1")
                End If
                If Val(Mid(.��ͷ, 3, 1)) <> 0 Then
                    Set objBarCode = PictureSpin(objBarCode, Val(Mid(.��ͷ, 3, 1)), frmFlash.picTemp)
                End If
                Set ImgCode(intCurID).Picture = objBarCode
                
                If .��� = 3 Then
                    '128���Զ��������
                    If Val(Mid(.��ͷ, 3, 1)) = 0 Then
                        ImgCode(intCurID).Width = Format(Me.ScaleX(sngWidth, vbMillimeters, vbTwips) * sgnMode, "0.00")
                        .W = Me.ScaleX(sngWidth, vbMillimeters, vbTwips)
                    Else
                        ImgCode(intCurID).Height = Format(Me.ScaleY(sngWidth, vbMillimeters, vbTwips) * sgnMode, "0.00")
                        .H = Me.ScaleY(sngWidth, vbMillimeters, vbTwips)
                    End If
                    blnSeek = True
                End If
            End With
            
            If blnSeek Then Call SeekItem(ImgCode(intCurID), ImgCode(intCurID).Left, ImgCode(intCurID).Top)
            If GetSelNum = 1 Then Call ShowAttrib(intCurID)
            
            LockWindowUpdate 0
        Case "��״"
            If mshAtt.TextMatrix(mshAtt.Row, 1) = cboAtt.Text Then Exit Sub
            objReport.Items("_" & intCurID).�߿� = IIF(cboAtt.Text = "����", False, True)
            mshAtt.TextMatrix(mshAtt.Row, 1) = cboAtt.Text
            Shp(intCurID).Shape = IIF(cboAtt.Text = "����", ShapeConstants.vbShapeRectangle, ShapeConstants.vbShapeOval)
            BlnSave = False

    End Select
    
    '���¶�λ�����У���Ϊѡ������ֵ֮����������з����˱仯
    If InStr("���ն���,����,����,��������,�����߿�,��ת����", strCurText) > 0 And GetSelNum = 1 Then
        For lngSize = 1 To mshAtt.Rows - 1
            If mshAtt.TextMatrix(lngSize, 0) = strCurText Then
                mshAtt.Row = lngSize: mshAtt.Col = 1: Exit For
            End If
        Next
    End If
End Sub

Private Sub cboFormat_Change()
    Dim tmpFmt As RPTFmt, bytOrder As Byte

    If blnAllowIn = False Then Exit Sub
    blnRefresh = False
    With cboFormat
        If Trim(.Text) = "" Then
            blnAllowIn = False
            Set .SelectedItem = .ComboItems("_" & mbytCurrFmt)
            blnAllowIn = True
            Exit Sub
        End If

        For Each tmpFmt In objReport.Fmts
            If tmpFmt.˵�� = Trim(.Text) Then Exit Sub
        Next

        '�޸ı�����ʽ����
        For Each tmpFmt In objReport.Fmts
            If tmpFmt.��� = mbytCurrFmt Then
                tmpFmt.˵�� = Trim(.Text)
                Exit For
            End If
        Next
        blnAllowIn = False
        .ComboItems("_" & mbytCurrFmt).Text = Trim(.Text)
        blnAllowIn = True
        BlnSave = False
    End With
End Sub

Private Sub cboFormat_Validate(Cancel As Boolean)
    Dim tmpFmt As RPTFmt
        
    'zyb#Add
    With cboFormat
        If Trim(.Text) = "" Then
            blnAllowIn = False
            Set .SelectedItem = .ComboItems("_" & mbytCurrFmt)
            blnAllowIn = True
            Exit Sub
        End If
        
        For Each tmpFmt In objReport.Fmts
            If tmpFmt.��� <> mbytCurrFmt And tmpFmt.˵�� = Trim(.Text) Then
                MsgBox "��ǰ����ı����ʽ�����Ѿ����ڣ����������룡"
                blnAllowIn = False
                Set .SelectedItem = .ComboItems("_" & mbytCurrFmt)
                blnAllowIn = True
                .SetFocus
                Cancel = True
                Exit Sub
            End If
        Next
    End With
End Sub

Private Sub cboText_Click()
    Dim strCurText As String, intType As Integer
    Dim ObjSel As Object
    
    strCurText = mshAtt.TextMatrix(mshAtt.Row, 0)
    Set ObjSel = GetInxObj(intCurID)
    
    Select Case strCurText
        Case "����"
            mshAtt.TextMatrix(mshAtt.Row, 1) = cboText.Text
            If UCase(TypeName(ObjSel)) = "LABEL" Then ObjSel.Caption = cboText.Text
            objReport.Items("_" & intCurID).���� = cboText.Text
        
            '�Ե��������LblSize�ؼ���λ��
            If UCase(TypeName(ObjSel)) = "LABEL" Then
                If ObjSel.AutoSize Then
                    'Call SelItem(ObjSel.Index, False)
                    'ж�ػᱨ���ƶ�λ�ü���
                    
                    Call SelMove(ObjSel.Index)
                End If
                objReport.Items("_" & intCurID).W = lbl(intCurID).Width / sgnMode
            End If
            
            cboText.Visible = False: mshAtt.SetFocus
            BlnSave = False
    End Select
    
End Sub

Private Sub cboText_KeyPress(KeyAscii As Integer)
    Dim ObjSel As Object
    Dim xx As Integer, yy As Integer, zz As Integer
    Dim strBarCode As String, objBarCode As StdPicture
    Dim strTemp As String
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        
        '�Ƿ���ָ��ַ�:
        If InString(cboText.Text, "'|~^") Then
            MsgBox "�����˷Ƿ��ַ���", vbInformation, App.Title
            cboText.SetFocus: Exit Sub
        End If
        Set ObjSel = GetInxObj(intCurID)

        Select Case mshAtt.TextMatrix(mshAtt.Row, 0)
            Case "����"
                If TLen(cboText.Text) > 255 Then
                    MsgBox "���ݲ��ܳ���255���ַ���", vbInformation, App.Title
                    cboText.SetFocus: Exit Sub
                End If
                
                Dim strNodeName As String, NodeThis As Node
                '�����adLongVarBinary���ֶ�,�������޸�
                xx = InStr(1, cboText, "]")
                yy = InStr(1, cboText, ".")
                zz = InStr(1, cboText, "[")
                If xx > zz And xx > yy And xx <> 0 And zz <> 0 Then
                    strNodeName = Mid(cboText, yy + 1, xx - yy - 1)
                    For Each NodeThis In tvwSQL.Nodes
                        If mdlPublic.GetStdNodeText(NodeThis.Text) = strNodeName And IsType(Val(NodeThis.Tag), adLongVarBinary) Then
                            MsgBox "����ѡ��ͼ���ֶ�Ϊ��ǩ�����ݣ�", vbInformation, App.Title
                            mshAtt.TextMatrix(mshAtt.Row, 1) = objReport.Items("_" & intCurID).����
                            Exit Sub
                        End If
                    Next
                End If
                
                mshAtt.TextMatrix(mshAtt.Row, 1) = cboText.Text
                If UCase(TypeName(ObjSel)) = "LABEL" Then ObjSel.Caption = cboText.Text
                objReport.Items("_" & intCurID).���� = cboText.Text
            
                '�Ե��������LblSize�ؼ���λ��
                If UCase(TypeName(ObjSel)) = "LABEL" Then
                    If ObjSel.AutoSize Then
                        Call SelItem(ObjSel.Index, False)
                        Call SelItem(ObjSel.Index, True)
                    End If
                    objReport.Items("_" & intCurID).W = lbl(intCurID).Width / sgnMode
                End If
                
                cboText.Visible = False: mshAtt.SetFocus
                BlnSave = False
        End Select
    Else
        Select Case mshAtt.TextMatrix(mshAtt.Row, 0)
            Case "����"
                If InStr("'|~^", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
        End Select
    End If
End Sub

Private Sub cbr_HeightChanged(ByVal NewHeight As Single)
    Form_Resize
End Sub

Private Sub CboFormat_Click()
    'zyb#Add
    If Trim(cboFormat.Text) = "" Then Exit Sub
    If mbytCurrFmt <> Mid(cboFormat.SelectedItem.Key, 2) Or blnRefresh Then
        blnRefresh = False
        mbytCurrFmt = Mid(cboFormat.SelectedItem.Key, 2)
        
        Call ShowSize: Call ShowScroll
        Call ReFlashReportBySelFormat
        Call picPaper_MouseDown(1, 0, 0, 0) '��ʾ��������
    End If
End Sub

Private Sub CboFormat_KeyPress(KeyAscii As Integer)
    'zyb#Add
    If blnDelReportFormat = False Then KeyAscii = 0: Exit Sub
    If KeyAscii = 39 Or KeyAscii = 22 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub Chart_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lngPreX = X: lngPreY = Y
    If Button = 1 Then
        If Shift = 2 Then
            If Mid(Chart(Index).Tag, 1, 2) = "" Then
                Call SelItem(Index, True) '��ѡ
                If GetSelNum() = 1 Then
                    Call ShowAttrib(Index) 'ֻѡ��һ������ʾ����
                Else
                    Call ShowAttrib '��ѡʱ����ʾ����
                End If
            Else
                Call SelItem(Index, False) '��ѡ
                If GetSelNum() = 1 Then
                    Call ShowAttrib(intCurID) 'ֻѡ��һ������ʾ����(ѡ�еĲ�һ���Ǹÿؼ�)
                Else
                    Call ShowAttrib '��ѡʱ����ʾ����
                End If
            End If
        Else
            If Mid(Chart(Index).Tag, 1, 2) = "" Then
                Call SelClear
                Call SelItem(Index, True)
                Call ShowAttrib(Index) 'ֻѡ��һ������ʾ����
            End If
        End If
    ElseIf Button = 2 Then
        PopupMenu mnuFormat, 2
    End If
End Sub

Private Sub Chart_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ObjSel As Object
    
    If objReport.Items("_" & Index).ϵͳ Then Exit Sub
    DrawXY X + Chart(Index).Left, Y + Chart(Index).Top
    Set ObjSel = Chart(Index)
    If Button = 1 And Mid(ObjSel.Tag, 1, 2) <> "" Then
        If blnLock Then Exit Sub
        Call MoveSelect(X - lngPreX, Y - lngPreY)
        If GetSelNum() = 1 Then ShowAttrib Index
    Else
        If objReport.Items("_" & Index).���� <> 12 Then
            ObjSel.MousePointer = 99
        End If
    End If
End Sub

Private Sub cmdAdd_Click()
    Dim bytOrder As Byte, strRptFmtName As String
    Dim objFmt As RPTFmt
    
    '���ӱ����ʽ��ֽ��ȱʡ���һ����ʽ��ͬ
    With cboFormat
        bytOrder = .ComboItems.count + 1
        If bytOrder < 100 Then
            strRptFmtName = GetRPtFmtName
            Set objFmt = objReport.Fmts(1)
            objReport.Fmts.Add bytOrder, strRptFmtName, objFmt.W, objFmt.H, objFmt.ֽ��, objFmt.ֽ��, objFmt.��ֽ̬��, 0, "_" & bytOrder
        
            blnAllowIn = False
            .ComboItems.Add , "_" & bytOrder, strRptFmtName, "Format"
            .ComboItems("_" & bytOrder).Selected = True
            Set .SelectedItem = .ComboItems("_" & bytOrder)
            .SetFocus
            
            blnAllowIn = True
            blnRefresh = True
            BlnSave = False
        Else
            MsgBox "�����ʽ̫�࣬���ܼ������ӣ������99�ָ�ʽ��", vbInformation, App.Title
            .SetFocus
            Exit Sub
        End If
    End With
    
    cmdDel.Enabled = (cboFormat.ComboItems.count > 1) And blnDelReportFormat
    tbr1.Buttons("DelFormat").Enabled = cmdDel.Enabled
    mnuEdit_DelFormat.Enabled = cmdDel.Enabled
    mbytCurrFmt = bytOrder
    
    Call CboFormat_Click
End Sub

Private Sub cmdDel_Click()
    Dim bytFmt As Byte, tmpFmt As RPTFmt, tmpItem As RPTItem
    Dim intModify As Integer, intDel As Integer
    
    'ɾ��������ʽ,�����¼���
    'zyb#Add
    With cboFormat
        If .ComboItems.count < 2 Then cmdDel.Enabled = False: Exit Sub
        If .SelectedItem Is Nothing Then Exit Sub
        If MsgBox("��ȷ��Ҫɾ���ø�ʽ�𣿣�ɾ���󽫲��ɻָ���", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then Exit Sub
        
        bytFmt = Mid(.SelectedItem.Key, 2)
        .ComboItems.Remove "_" & bytFmt
        For intModify = bytFmt To .ComboItems.count
            If Mid(.ComboItems("_" & intModify + 1).Key, 2) > bytFmt Then
                .ComboItems("_" & intModify + 1).Key = "_" & (Mid(.ComboItems("_" & intModify + 1).Key, 2) - 1)
            End If
        Next
        Set .SelectedItem = .ComboItems("_" & IIF(bytFmt = 1, 1, bytFmt - 1))
    End With
    
    'ɾ��������������ʽ�����
    objReport.Fmts.Remove "_" & bytFmt
    For intDel = bytFmt + 1 To objReport.Fmts.count + 1
        With objReport.Fmts("_" & intDel)
            If .��� > bytFmt Then
                objReport.Fmts.Add intDel - 1, .˵��, .W, .H, .ֽ��, .ֽ��, .��ֽ̬��, .ͼ��, "_" & intDel - 1
                objReport.Fmts.Remove "_" & intDel
            End If
        End With
    Next
    'ɾ���ñ�����ʽ��Ӧ�����б���Ԫ��
    For Each tmpItem In objReport.Items
        If tmpItem.��ʽ�� = bytFmt Then
            objReport.Items.Remove "_" & tmpItem.Key
        End If
    Next
    '�޸�����ı���Ԫ������Ӧ�ĸ�ʽ��
    For intDel = 1 To objReport.Items.count
        Set tmpItem = objReport.Items(intDel)
        If tmpItem.��ʽ�� > bytFmt Then
            tmpItem.��ʽ�� = tmpItem.��ʽ�� - 1
        End If
    Next
    
    blnRefresh = True
    BlnSave = False
    cmdDel.Enabled = (cboFormat.ComboItems.count > 1) And blnDelReportFormat
    Call CboFormat_Click
End Sub

Private Function GetInxObj(ByVal intIndex As Integer) As Object
'���ܣ�����ID����ö�Ӧ��Ԫ�ض���
    Dim ObjSel As Object
    
    Select Case objReport.Items("_" & intIndex).����
        Case 1
            Set ObjSel = lblLine(intIndex)
        Case 2, 3
            Set ObjSel = lbl(intIndex)
        Case 10
            Set ObjSel = Shp(intIndex)
        Case 4, 5
            Set ObjSel = msh(intIndex)
        Case 11
            Set ObjSel = img(intIndex)
        Case 12 '@@@
            Set ObjSel = Chart(intIndex)
        Case 13
            Set ObjSel = ImgCode(intIndex)
        Case 14
            Set ObjSel = pic(intIndex)
    End Select
    Set GetInxObj = ObjSel
End Function

Private Sub SetCmdAttBackColor(ByVal intIndex As Integer)
'���ܣ������ƶ�Ԫ�صı���ɫ
    Dim ObjSel As Object
    On Error Resume Next
    Set ObjSel = GetInxObj(intIndex)
    '�ؼ�����
    If objReport.Items("_" & intIndex).���� = 10 Then
        objReport.Items("_" & intIndex).���� = cdg.Color '�ȸ�ֵ
        Call DrawFrame(ObjSel)
    ElseIf objReport.Items("_" & intIndex).���� = 12 Then '@@@
        ObjSel.Interior.BackgroundColor = IIF(cdg.Color = &HFFFFFF, lbl(0).BackColor, cdg.Color) '��ɫ����
        objReport.Items("_" & intIndex).���� = cdg.Color
    Else
        If cdg.Color = &HFFFFFF Then
            If objReport.Items("_" & intIndex).���� = 4 Or objReport.Items("_" & intIndex).���� = 5 Then
                ObjSel.BackColor = cdg.Color
                ObjSel.BackColorFixed = lbl(0).BackColor '��ɫ���ֳ��ǹ̶�����
            Else
                ObjSel.BackColor = lbl(0).BackColor '��ɫ����
            End If
        Else
            ObjSel.BackColor = cdg.Color
            If objReport.Items("_" & intIndex).���� = 4 Or objReport.Items("_" & intIndex).���� = 5 Then
                ObjSel.BackColorFixed = cdg.Color
            End If
        End If
        If objReport.Items("_" & intIndex).���� = 4 Or objReport.Items("_" & intIndex).���� = 5 Then
            Call ResetColor(ObjSel.Index) '�ܹ�,���������Ԫˢ��
            If objReport.Items("_" & intIndex).���� = 4 Then
                Call SetCopyGrid(intIndex)
            End If
        End If
        objReport.Items("_" & intIndex).���� = cdg.Color
    End If
End Sub

Private Sub SetCmdAttForeColor(ByVal intIndex As Integer)
'���ܣ������ƶ�Ԫ�ص�ǰ��ɫ
    Dim ObjSel As Object
    On Error Resume Next
    
    Set ObjSel = GetInxObj(intIndex)
    If objReport.Items("_" & intIndex).���� = 1 Then
        '�������Ա���ɫ��ʾ
        If cdg.Color = &HFFFFFF Then
            ObjSel.BackColor = lbl(0).BackColor '��ɫʱ��ʾ̸ɫ
        Else
            ObjSel.BackColor = cdg.Color
        End If
    ElseIf objReport.Items("_" & intIndex).���� = 12 Then '@@@
        ObjSel.Interior.ForegroundColor = cdg.Color
        '��֪Ϊʲô�����ÿؼ�ǰ����Ч,��ͨ�����Կ����Ч
        ObjSel.ChartArea.Axes("X").AxisStyle.LineStyle.Color = cdg.Color
        ObjSel.ChartArea.Axes("Y").AxisStyle.LineStyle.Color = cdg.Color
    Else
        ObjSel.ForeColor = cdg.Color
    End If
    
    '����й̶�����ǰ��ɫ
    If objReport.Items("_" & intIndex).���� = 4 Or objReport.Items("_" & intIndex).���� = 5 Then
        ObjSel.ForeColorFixed = cdg.Color

        Call ResetColor(ObjSel.Index) '�ܹ�,���������Ԫˢ��
        If objReport.Items("_" & intIndex).���� = 4 Then
            Call SetCopyGrid(intIndex)
        End If
    End If
    
    '����ֵ
    objReport.Items("_" & intIndex).ǰ�� = cdg.Color
End Sub

Private Sub SetCmdAttFont(ByVal intIndex As Integer)
'���ܣ������ƶ�Ԫ�ص�����
    Dim ObjSel As Object, sgnH As Single
    Dim i As Long
    
    On Error Resume Next
    Set ObjSel = GetInxObj(intIndex)
    '��������
    objReport.Items("_" & intIndex).���� = cdg.FontName
    objReport.Items("_" & intIndex).�ֺ� = Format(cdg.FontSize, "0.0") '����С�ֺ�@@@
    objReport.Items("_" & intIndex).���� = cdg.FontBold
    objReport.Items("_" & intIndex).б�� = cdg.FontItalic
    If objReport.Items("_" & intIndex).���� <> 12 Then '@@@
        objReport.Items("_" & intIndex).���� = cdg.FontUnderline
        objReport.Items("_" & intIndex).ǰ�� = cdg.Color '�����ɫǰ��@@@
    End If

    '�ؼ�����
    If objReport.Items("_" & intIndex).���� = 12 Then '@@@
        Call SetChartStyleAndData(ObjSel, objReport.Items("_" & intIndex), , sgnMode, True)
    Else
        'Ϊ��������ؼ�װ����������
        PicFontTest.Font.name = cdg.FontName
        PicFontTest.Font.Size = cdg.FontSize * sgnMode '����С�ֺ�@@@
        PicFontTest.Font.Bold = cdg.FontBold
        PicFontTest.Font.Italic = cdg.FontItalic
        PicFontTest.Font.Underline = cdg.FontUnderline
        sgnH = (PicFontTest.TextHeight("��") + 15) * sgnMode
        
        ObjSel.Font.name = cdg.FontName
        ObjSel.Font.Size = cdg.FontSize * sgnMode '����С�ֺ�@@@
        ObjSel.Font.Bold = cdg.FontBold
        ObjSel.Font.Italic = cdg.FontItalic
        ObjSel.Font.Underline = cdg.FontUnderline
        ObjSel.ForeColor = cdg.Color '�����ɫǰ��@@@
        If TypeName(ObjSel) = "VSFlexGrid" Then
            ObjSel.ForeColorFixed = ObjSel.ForeColor
        End If
    End If
    
    '�������弰����ʾ�̶������Զ���������и�(����ʱ)
    Select Case objReport.Items("_" & intIndex).����
        Case 4, 5
            If ObjSel.RowHeight(0) < sgnH Then
                If Abs(Int(-ObjSel.Height / sgnH)) >= ObjSel.FixedRows + 2 Then
                    For i = 0 To ObjSel.Rows - 1
                        ObjSel.RowHeight(i) = sgnH
                    Next
                Else
                    For i = 0 To ObjSel.Rows - 1
                        ObjSel.RowHeight(i) = Abs(Int(-ObjSel.Height / (ObjSel.FixedRows + 2)))
                    Next
                End If
                objReport.Items("_" & intIndex).�и� = Format(ObjSel.RowHeight(0) / sgnMode, "0.00")
                mshAtt.TextMatrix(GetRow("�и�"), 1) = Format(ObjSel.RowHeight(0) / Twip_mm / sgnMode, "0.00")
                Call SetGridLine(intIndex)
            End If
            Call ResetColor(ObjSel.Index) '�ܹ�,���������Ԫˢ��
            If objReport.Items("_" & intIndex).���� = 4 Then
                Call SetCopyGrid(intIndex)
            End If
        Case 2, 3
            If ObjSel.Height < sgnH Then
                ObjSel.Height = sgnH: ObjSel.Width = PicFontTest.TextWidth("��") * TLen(ObjSel.Text) / 2
                objReport.Items("_" & intIndex).H = ObjSel.Height / sgnMode
                objReport.Items("_" & intIndex).W = ObjSel.Width / sgnMode
                mshAtt.TextMatrix(GetRow("�߶�"), 1) = Format(ObjSel.Height / Twip_mm / sgnMode, "0.00")
                mshAtt.TextMatrix(GetRow("���"), 1) = Format(ObjSel.Width / Twip_mm / sgnMode, "0.00")
            End If
    End Select
    
    i = intIndex
    intIndex = ObjSel.Index
    Call ReferTo
    intIndex = i
    
    Call SelItem(ObjSel.Index, False)
    Call SelItem(ObjSel.Index, True)
End Sub

Private Sub cmdAtt_Click()
    Dim ObjSel As Object, i As Integer
    Dim tmpItem As RPTItem
    Dim strInfo As String
    Dim lngReportID As Long
    Dim strReportID As String
    Dim X As Long, Y As Long, k As Long
    Dim tmpObj As PictureBox
    
    Set ObjSel = GetInxObj(intCurID)
    
    On Error Resume Next
    cdg.CancelError = True
    Select Case mshAtt.TextMatrix(mshAtt.Row, 0)
        Case "����"
            Set ObjSel = img(intCurID)
            cdg.Flags = &H80000 Or &H4 Or &H1000 Or &H200000 Or &H800
            cdg.Filter = "ͼƬ�ļ�(*.ico;*.cur;*.bmp;*.gif;*.jpg;*.rle;*.wmf;*.emf)|*.ico;*.cur;*.bmp;*.gif;*.jpg;*.rle;*.wmf;*.emf"
            cdg.InitDir = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.name, "ͼƬ·��", "C:\")
            cdg.ShowOpen
            If Err.Number = 0 Then
                ObjSel.Picture = LoadPicture(cdg.FileName)
                If Err.Number <> 0 Then
                    MsgBox "ѡ���ͼƬ�ļ���ʽ����", vbInformation, App.Title
                    Set ObjSel.Picture = Nothing
                    Exit Sub
                End If
                
                '�ȱ��汸��
                Set objReport.Items("_" & intCurID).ͼƬ = ObjSel.Picture
                
                '·��������ע���
                SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.name, "ͼƬ·��", Replace(cdg.FileName, cdg.FileTitle, "") 'Mid(cdg.FileName, 1, Len(cdg.FileName) - InStr(1, strReverse(cdg.FileName), "\"))
                
                '����ѡ��ؼ�
                If objReport.Items("_" & intCurID).�Ե� Then
                    ObjSel.Stretch = False
                    ObjSel.Width = ObjSel.Width * sgnMode
                    ObjSel.Height = ObjSel.Height * sgnMode
                    Call SelItem(ObjSel.Index, False)
                    Call SelItem(ObjSel.Index, True)
                    ObjSel.Stretch = True
                End If
                '���ֱ���
                If objReport.Items("_" & intCurID).���� Then
                    Set ObjSel.Picture = ScalePicture(PicFontTest, objReport.Items("_" & intCurID).ͼƬ, ObjSel.Width, ObjSel.Height)
                End If
                
                '��ֵ
                With objReport.Items("_" & intCurID)
                    .���� = cdg.FileName
                    .X = Format(ObjSel.Left / sgnMode, "0.00")
                    .Y = Format(ObjSel.Top / sgnMode, "0.00")
                    .H = Format(ObjSel.Height / sgnMode, "0.00")
                    .W = Format(ObjSel.Width / sgnMode, "0.00")
                    mshAtt.TextMatrix(mshAtt.Row, 1) = "[Picture]"
                    mshAtt.TextMatrix(GetRow("X����"), 1) = Format(.X / Twip_mm, "0.00")
                    mshAtt.TextMatrix(GetRow("Y����"), 1) = Format(.Y / Twip_mm, "0.00")
                    mshAtt.TextMatrix(GetRow("�߶�"), 1) = Format(.H / Twip_mm, "0.00")
                    mshAtt.TextMatrix(GetRow("���"), 1) = Format(.W / Twip_mm, "0.00")
                End With
                
                BlnSave = False
            Else
                Err.Clear
            End If
        Case "��������ɫ", "����ɫ"
            cdg.CancelError = True
            cdg.Flags = &H1 Or &H2
            cdg.Color = objReport.Items("_" & intCurID).����
            cdg.ShowColor
            If Err.Number = 0 Then
                '����ֵ����
                mshAtt.Col = 1
                mshAtt.CellForeColor = cdg.Color
                
                ObjSel.GridColor = cdg.Color
                If mshAtt.TextMatrix(mshAtt.Row, 0) = "����ɫ" Then
                    ObjSel.GridColorFixed = cdg.Color
                End If
                
                Call ResetColor(ObjSel.Index) '�ܹ�,���������Ԫˢ��
                If objReport.Items("_" & intCurID).���� = 4 Then
                    Call SetCopyGrid(intCurID)
                End If
                
                '����ֵ
                objReport.Items("_" & intCurID).���� = cdg.Color
                BlnSave = False
            Else
                Err.Clear
            End If
                
            '�������ӱ��������������һ��
            If objReport.Items("_" & intCurID).���� = "" And objReport.Items("_" & intCurID).���� = 5 Then
                For Each tmpItem In objReport.Items
                    If tmpItem.��ʽ�� = mbytCurrFmt And tmpItem.���� = objReport.Items("_" & intCurID).���� And tmpItem.���� = 5 Then
                        Call SetGridLike(msh(intCurID), msh(tmpItem.Key))
                    End If
                Next
            End If
        Case "��ͷ����ɫ"
            cdg.CancelError = True
            cdg.Flags = &H1 Or &H2
            cdg.Color = IIF(objReport.Items("_" & intCurID).��ʽ = "", objReport.Items("_" & intCurID).����, Val(objReport.Items("_" & intCurID).��ʽ))
            cdg.ShowColor
            If Err.Number = 0 Then
                '����ֵ����
                mshAtt.Col = 1
                mshAtt.CellForeColor = cdg.Color
                
                ObjSel.GridColorFixed = cdg.Color
                
                Call ResetColor(ObjSel.Index) '�ܹ�,���������Ԫˢ��
                If objReport.Items("_" & intCurID).���� = 4 Then
                    Call SetCopyGrid(intCurID)
                End If
                
                '����ֵ
                objReport.Items("_" & intCurID).��ʽ = cdg.Color
                BlnSave = False
            Else
                Err.Clear
            End If
                
        Case "ǰ��ɫ"
            cdg.Flags = &H1 Or &H2
            cdg.Color = objReport.Items("_" & intCurID).ǰ��
            cdg.ShowColor
            If Err.Number = 0 Then
                '����ֵ����
                mshAtt.Col = 1
                mshAtt.CellForeColor = cdg.Color
                
                '�����ѡ�е�����Ԫ�ض�����ͬ��Ԫ��
                If lblSize.count > 9 Then
                    For Each tmpObj In lblSize
                        If tmpObj.Index Mod 8 = 1 Then
                            Call SetCmdAttForeColor(tmpObj.Tag)
                        End If
                    Next
                Else
                    Call SetCmdAttForeColor(intCurID)
                End If
                
                BlnSave = False
            Else
                Err.Clear
            End If
                
            '�������ӱ��������������һ��
            If objReport.Items("_" & intCurID).���� = "" And objReport.Items("_" & intCurID).���� = 5 Then
                For Each tmpItem In objReport.Items
                    If tmpItem.��ʽ�� = mbytCurrFmt And tmpItem.���� = objReport.Items("_" & intCurID).���� And tmpItem.���� = 5 Then
                        Call SetGridLike(msh(intCurID), msh(tmpItem.Key))
                    End If
                Next
            End If
        Case "����ɫ"
            cdg.Flags = &H1 Or &H2
            cdg.Color = objReport.Items("_" & intCurID).����
            cdg.ShowColor
            If Err.Number = 0 Then
                '����ֵ����
                mshAtt.Col = 1
                mshAtt.CellForeColor = cdg.Color
                '�����ѡ�е�����Ԫ�ض�����ͬ��Ԫ��
                If lblSize.count > 9 Then
                    For Each tmpObj In lblSize
                        If tmpObj.Index Mod 8 = 1 Then
                            Call SetCmdAttBackColor(tmpObj.Tag)
                        End If
                    Next
                Else
                    Call SetCmdAttBackColor(intCurID)
                End If
                
                BlnSave = False
            Else
                Err.Clear
            End If
                
            '�������ӱ��������������һ��
            If objReport.Items("_" & intCurID).���� = "" And objReport.Items("_" & intCurID).���� = 5 Then
                For Each tmpItem In objReport.Items
                    If tmpItem.��ʽ�� = mbytCurrFmt And tmpItem.���� = objReport.Items("_" & intCurID).���� And tmpItem.���� = 5 Then
                        Call SetGridLike(msh(intCurID), msh(tmpItem.Key))
                    End If
                Next
            End If
        Case "����" '���Ըı�����,�ֺ�,����,б��
            cdg.Flags = &H3 Or &H400 Or &H200 Or &H10000
            If objReport.Items("_" & intCurID).���� <> 12 Then '@@@
                cdg.Flags = cdg.Flags Or &H100
            End If
            cdg.FontName = objReport.Items("_" & intCurID).����
            cdg.FontSize = objReport.Items("_" & intCurID).�ֺ�
            cdg.FontBold = objReport.Items("_" & intCurID).����
            cdg.FontItalic = objReport.Items("_" & intCurID).б��
            If objReport.Items("_" & intCurID).���� <> 12 Then '@@@
                cdg.FontUnderline = objReport.Items("_" & intCurID).����
                cdg.Color = objReport.Items("_" & intCurID).ǰ��
            End If
            
            cdg.ShowFont
            If Err.Number = 0 Then
                mshAtt.TextMatrix(mshAtt.Row, Val("1-������")) = cdg.FontName
                '�����ѡ�е�����Ԫ�ض�����ͬ��Ԫ��
                If lblSize.count > 9 Then
                    For Each tmpObj In lblSize
                        If tmpObj.Index Mod 8 = 1 Then
                            Call SetCmdAttFont(tmpObj.Tag)
                        End If
                    Next
                Else
                    Call SetCmdAttFont(intCurID)
                End If
                BlnSave = False
            Else
                Err.Clear
            End If
            
            '�������ӱ��������������һ��
            If objReport.Items("_" & intCurID).���� = "" And objReport.Items("_" & intCurID).���� = 5 Then
                For Each tmpItem In objReport.Items
                    If tmpItem.��ʽ�� = mbytCurrFmt And tmpItem.���� = objReport.Items("_" & intCurID).���� And tmpItem.���� = 5 Then
                        Call SetGridLike(msh(intCurID), msh(tmpItem.Key))
                    End If
                Next
            End If
        Case "����"
            Set tmpItem = objReport.Items("_" & intCurID)
            If frmChartSetup.ShowMe(Me, objReport.Datas, ObjSel, tmpItem) Then
                Call CopyItem(objReport.Items("_" & intCurID), tmpItem, False)
                Call ShowAttrib(intCurID)
                mshAtt.Row = 3
                mshAtt_AfterRowColChange 0, 0, mshAtt.Row, mshAtt.Col
                BlnSave = False
            End If
        Case "��������"
            X = InStr(1, objReport.Items("_" & intCurID).����, "]")
            Y = InStr(1, objReport.Items("_" & intCurID).����, ".")
            k = InStr(1, objReport.Items("_" & intCurID).����, "[")
            If X > k And X > Y And X <> 0 And k <> 0 Then
                strReportID = FindReport("", txtAtt.hwnd, strInfo, objReport.Items("_" & intCurID).Relations.Item(1).��������ID, objReport, objReport.Items("_" & intCurID).Relations, 2, Me, intCurID)
                If strReportID <> "" Then
                    mshAtt.TextMatrix(mshAtt.Row, 1) = strInfo
                    mshAtt.RowData(mshAtt.Row) = strReportID
                    txtAtt.Visible = False: mshAtt.SetFocus
                    BlnSave = False
                Else
                    '�ж���ȡ���������
                    If objReport.Items("_" & intCurID).Relations.count > 0 Then
                        txtAtt.SetFocus
                    Else
                        mshAtt.TextMatrix(mshAtt.Row, 1) = ""
                        mshAtt.RowData(mshAtt.Row) = 0
                        txtAtt.Text = ""
                        txtAtt.Visible = True
                        txtAtt.SetFocus
                        BlnSave = False
                    End If
                End If
            Else
                MsgBox "��ǰ��ǩ�����Ȱ�һ������Դ�����磺[����Դ.�ֶ�],�󶨺������ù�������", vbInformation, Me.Caption
            End If
    End Select
    
    mshAtt.SetFocus
End Sub

Private Sub dtpAtt_Change()
    Select Case mshAtt.TextMatrix(mshAtt.Row, 0)
        Case "��ֹ��ʼʱ��"
            objReport.��ֹ��ʼʱ�� = dtpAtt.Value
            mshAtt.TextMatrix(mshAtt.Row, 1) = Format(objReport.��ֹ��ʼʱ��, "HH:mm:ss")
            
            dtpAtt.Visible = False: If dtpAtt.Visible Then dtpAtt.SetFocus
            BlnSave = False
        Case "��ֹ����ʱ��"
            objReport.��ֹ����ʱ�� = dtpAtt.Value
            mshAtt.TextMatrix(mshAtt.Row, 1) = Format(objReport.��ֹ����ʱ��, "HH:mm:ss")
            
            dtpAtt.Visible = False: If dtpAtt.Visible Then dtpAtt.SetFocus
            BlnSave = False
    End Select
End Sub

Private Sub Form_Activate()
    If mblnFirst Then
        mblnFirst = False
        Me.Refresh
        On Error Resume Next
        picPaper.SetFocus
        Call picPaper_MouseDown(1, 0, 0, 0)
    End If
End Sub

Private Sub Form_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Call ClearXY
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim blnScroll As Boolean
    
    If Shift = 0 Then
        If ActiveControl Is Nothing Then
            blnScroll = True
        ElseIf TypeName(ActiveControl) <> "TextBox" _
            And TypeName(ActiveControl) <> "ComboBox" _
            And TypeName(ActiveControl) <> "ListView" _
            And UCase(ActiveControl.name) <> "MSHATT" Then
            blnScroll = True
        End If
    End If
    
    Select Case KeyCode
        Case vbKeyUp
            If blnScroll And scrVsc.Enabled And scrVsc.Value > scrVsc.Min Then scrVsc.Value = IIF(scrVsc.Value - scrVsc.LargeChange < scrVsc.Min, scrVsc.Min, scrVsc.Value - scrVsc.LargeChange)
            If Shift = 2 And GetSelNum > 0 Then
                MoveSelect 0, -15
                If GetSelNum = 1 Then ShowAttrib intCurID
            End If
        Case vbKeyDown
            If blnScroll And scrVsc.Enabled And scrVsc.Value < scrVsc.Max Then scrVsc.Value = IIF(scrVsc.Value + scrVsc.LargeChange > scrVsc.Max, scrVsc.Max, scrVsc.Value + scrVsc.LargeChange)
            If Shift = 2 And GetSelNum > 0 Then
                MoveSelect 0, 15
                If GetSelNum = 1 Then ShowAttrib intCurID
            End If
        Case vbKeyLeft
            If blnScroll And scrHsc.Enabled And scrHsc.Value > scrHsc.Min Then scrHsc.Value = IIF(scrHsc.Value - scrHsc.LargeChange < scrHsc.Min, scrHsc.Min, scrHsc.Value - scrHsc.LargeChange)
            If Shift = 2 And GetSelNum > 0 Then
                MoveSelect -15, 0
                If GetSelNum = 1 Then ShowAttrib intCurID
            End If
        Case vbKeyRight
            If blnScroll And scrHsc.Enabled And scrHsc.Value < scrHsc.Max Then scrHsc.Value = IIF(scrHsc.Value + scrHsc.LargeChange > scrHsc.Max, scrHsc.Max, scrHsc.Value + scrHsc.LargeChange)
            If Shift = 2 And GetSelNum > 0 Then
                MoveSelect 15, 0
                If GetSelNum = 1 Then ShowAttrib intCurID
            End If
    End Select
    
    If Shift = 4 Then
        If picPaper.MousePointer <> 99 Then
            Set picPaper.MouseIcon = picBack.MouseIcon
            picPaper.MousePointer = 99
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        'ȡ�����Ա༭״̬
        KeyAscii = 0
        If txtAtt.Visible Or cmdAtt.Visible Or cboAtt.Visible Then
            txtAtt.Visible = False
            cmdAtt.Visible = False
            cboAtt.Visible = False
            cboAtt.Clear: txtAtt.Text = ""
            mshAtt.SetFocus
        End If
    Else
        If InStr("'&", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    End If
End Sub

Private Sub Form_Keyup(KeyCode As Integer, Shift As Integer)
    If Shift = 0 Then picPaper.MousePointer = 0
End Sub

Private Sub Form_Load()
    Dim rsDelReportFormat As New ADODB.Recordset
    Dim strSQL As String
    
    Screen.MousePointer = vbHourglass
        
    RestoreWinState Me, App.ProductName
    tb2.ZOrder
    If mblnNotModiData Then
        mnuEdit_New.Visible = False
        mnuEdit_Modi.Visible = False
        mnuEdit_Del.Visible = False
        mnuEdit_Data_.Visible = False
        tbr1.Buttons("New").Visible = False
        tbr1.Buttons("Modi").Visible = False
        tbr1.Buttons("Del").Visible = False
        tbr1.Buttons("Data_").Visible = False
    End If
    
    '��ʾ����
    sgnMode = 1: sgnLastMode = 1
    mblnFirst = True
    
    picR.Visible = mnuViewToolAttrib.Checked
    picAtt.Visible = mnuViewToolAttrib.Checked

    picL.Visible = mnuViewToolSQL.Checked
    picSQL.Visible = mnuViewToolSQL.Checked

    picRulerH.Visible = mnuViewToolRuler.Checked
    picRulerV.Visible = mnuViewToolRuler.Checked

    '��ʼ����ת�������
    Set objFont = New clsRotateFont
    Set objFont.LogFont = New StdFont
    objFont.LogFont.name = "Times New Roman"
    objFont.LogFont.Size = 6.5
    objFont.Rotation = 90
    
    gblnModi = False
    intMaxID = 0: intCurID = 0
    blnLock = True
    bytCurTool = 0
    BlnSave = True
    Set objLastSel = Nothing
    
    selCell.Col1 = -1: selCell.Col2 = -1: selCell.Row1 = -1: selCell.Row2 = -1: selCell.Row = -1: sta.Panels(2).Text = ""
    intCurCol = -1
    
    'zyb#Add
    '��ʼ����ʾ�����˵�
    mnuViewScaleMode(0).Checked = False
    mnuViewScaleMode(1).Checked = False
    mnuViewScaleMode(2).Checked = False
    mnuViewScaleMode(4).Checked = False
    mnuViewScaleMode(5).Checked = True
    mnuViewScaleMode(6).Checked = False
    mnuViewScaleMode(7).Checked = False
    mnuViewScaleMode(8).Checked = False
    
    '��ȡͬʱ�ı�intMaxID
    blnAdjustRowHeight = False
    Set objReport = ReadReport(lngRPTID, intMaxID)
    Call GetInPaper
    
    If Not objReport Is Nothing Then
        '��ȡ�ñ����Ƿ�Ϊ�̶�����
        strSQL = "Select Nvl(ϵͳ,0) ϵͳ From zlReports Where ID=[1]"
        Set rsDelReportFormat = OpenSQLRecord(strSQL, Me.Caption, lngRPTID)
        blnDelReportFormat = (rsDelReportFormat!ϵͳ = 0)
        blnDelReportFormat = True '200312:�̶�����Ҳ����
        
        '��ʾ��������(����Դ��ֽ�š�Ԫ��)
        'zyb#Modify
        Call LoadReportFormat       '��ȡ�ñ�������и�ʽ
        Call ReFlashReport(False)
        
        Caption = Caption & " - [" & objReport.��� & "]" & objReport.���� & IIF(objReport.˵�� = "", "", "��" & objReport.˵��)
    Else
        Screen.MousePointer = vbDefault
        MsgBox "������ȷ��ȡ�������ݣ�", vbInformation, App.Title
        Unload Me: Exit Sub
    End If
    
    Screen.MousePointer = vbDefault
    If objReport.Items.count = 0 Then mnuFormat_Lock_Click
End Sub

Private Sub ShowPaperInfo()
    Dim objFmt As RPTFmt
    
    If Not objReport Is Nothing Then
        Set objFmt = objReport.Fmts("_" & mbytCurrFmt)
        sta.Panels(2).Text = "��ӡ��:" & objReport.��ӡ�� & "   ֽ��:" & GetPaperName(objFmt.ֽ��, objFmt.W, objFmt.H) & " " & _
            IIF(objFmt.ֽ�� = 256, CInt(objFmt.W / Twip_mm) & "mm �� " & CInt(objFmt.H / Twip_mm) & "mm", "") & _
            IIF(objFmt.ֽ�� = 1, "   ����", "   ����")
    Else
        sta.Panels(2).Text = ""
    End If
End Sub

Private Sub ReFlashReport(Optional blnReload As Boolean = False)
'���ܣ�����ˢ����ʾ��������
'������blnReLoad=�Ƿ����´����ݿ��м�������
    Dim objTmp As Object, tmpReport As Report, intPreMax As Long
    
    If blnReload Then
        intPreMax = intMaxID
        Set tmpReport = ReadReport(lngRPTID, intMaxID)
        If tmpReport Is Nothing Then
            MsgBox "��������ˢ��ʧ�ܣ�", vbInformation, App.Title
            intMaxID = intPreMax: Exit Sub
        End If
        Set objReport = tmpReport
        BlnSave = True
    End If
    
    For Each objTmp In lblSize
        If objTmp.Index <> 0 Then Unload objTmp
    Next
    For Each objTmp In lblLine
        If objTmp.Index <> 0 Then Unload objTmp
    Next
    For Each objTmp In lbl
        If objTmp.Index <> 0 Then Unload objTmp
    Next
    For Each objTmp In msh
        If objTmp.Index <> 0 Then Unload objTmp
    Next
    For Each objTmp In img
        If objTmp.Index <> 0 Then Unload objTmp
    Next
    For Each objTmp In ImgCode
        If objTmp.Index <> 0 Then Unload objTmp
    Next
    For Each objTmp In Chart
        If objTmp.Index <> 0 Then Unload objTmp
    Next
    For Each objTmp In pic
        If objTmp.Index <> 0 Then Unload objTmp
    Next
    For Each objTmp In Shp
        If objTmp.Index <> 0 Then Unload lblshp(objTmp.Index): Unload objTmp
    Next
    
    intCurID = 0
    Set objLastSel = Nothing
    
    '��ռ�����
    If Me.Visible Then Set objClip = New RPTItems
    
    Call ShowReportDetail
    Call ShowAttrib
    
    Me.Refresh
End Sub

Private Sub Form_Resize()
    Dim cbrH As Long    '������ռ�ø߶�
    Dim staH As Long    '״̬��ռ�ø߶�
    Dim attW As Long    '���Կ���ز���ռ�ÿ��
    Dim sqlW As Long    '�����б���ز���ռ�ÿ��
    Dim formatH As Long '��ʽ�߶�
    Dim rulW As Long    '��߿��
    Dim rulH As Long    '��߸߶�
    Dim i As Integer
    
    On Error Resume Next
    
    If WindowState = vbMinimized Then Exit Sub
    If blnMax Or WindowState = vbMaximized Then
        picSQL.Width = 3500
        picAtt.Width = 2400
        lvwPar.Height = 1700
        lblNote.Height = 900
        blnMax = False
    End If
    If WindowState = vbMaximized Then blnMax = True
    
    If Width < 8000 Then Width = 8000
    If Height < 5000 Then Height = 5000
    
    '����ؼ���Ⱥ͸߶�
    cbrH = IIF(cbr.Visible, cbr.Height, 0)
    staH = IIF(sta.Visible, sta.Height, 0)
    attW = IIF(picR.Visible, picR.Width + picAtt.Width, 0)
    sqlW = IIF(picL.Visible, picL.Width + picSQL.Width, 0)
    formatH = picFormat.Height
    rulW = IIF(picRulerV.Visible, picRulerV.Width, 0)
    rulH = IIF(picRulerH.Visible, picRulerH.Height, 0)
    
    'zyb#Add
    picFormat.Top = ScaleTop + cbrH
    picFormat.Left = ScaleLeft + sqlW
    picFormat.Width = Me.ScaleWidth - picFormat.Left - attW
    
    'zyb#Add
    cmdDel.Left = picFormat.Width - cmdDel.Width - 15
    cmdAdd.Left = cmdDel.Left - cmdAdd.Width - 15
    If cmdAdd.Left - cboFormat.Left - 50 > 3000 Then
        cboFormat.Width = cmdAdd.Left - cboFormat.Left - 30
    End If
    
    picRulerV.Top = picFormat.Top + formatH + rulH  'zyb#Modify
    picRulerV.Left = ScaleLeft + sqlW
    picRulerV.Height = ScaleHeight - cbrH - staH - rulH - scrHsc.Height - formatH   'zyb#Modify
    
    picRulerH.Left = ScaleLeft + sqlW
    picRulerH.Top = picFormat.Top + formatH 'zyb#Modify
    picRulerH.Width = ScaleWidth - sqlW - attW - scrVsc.Width
    
    scrHsc.Top = picRulerV.Top + picRulerV.Height
    scrHsc.Left = picRulerV.Left + rulW
    scrHsc.Width = picRulerH.Width - rulW
    
    scrVsc.Top = picRulerV.Top
    scrVsc.Left = picRulerH.Left + picRulerH.Width
    scrVsc.Height = picRulerV.Height
    
    picBack.Left = picRulerV.Left + rulW
    picBack.Top = picRulerH.Top + rulH
    picBack.Width = scrHsc.Width
    picBack.Height = scrVsc.Height
    
    lblSQL.Top = 15: lblSQL.Left = 30
    lblSQL.Width = picSQL.ScaleWidth - 60
    
    tvwSQL.Left = picSQL.ScaleLeft
    tvwSQL.Top = lblSQL.Height + 30
    tvwSQL.Width = picSQL.ScaleWidth
    tvwSQL.Height = (ScaleHeight - staH - cbrH) - lblSQL.Height - lblPar.Height - lvwPar.Height - 60
    
    lblPar.Top = tvwSQL.Top + tvwSQL.Height + 15
    lblPar.Left = picSQL.ScaleLeft + 30
    lblPar.Width = lblSQL.Width
    
    lvwPar.Top = lblPar.Top + lblPar.Height + 15
    lvwPar.Left = picSQL.ScaleLeft + 15
    lvwPar.Width = tvwSQL.Width
    
    lblTool.Top = 15: lblTool.Left = 30
    lblTool.Width = picAtt.ScaleWidth - 60
    
    tbrTool.Top = lblTool.Top + lblTool.Height
    tbrTool.Left = lblTool.Left
    tbrTool.Width = lblTool.Width
    
    lblAtt.Top = tbrTool.Top + tbrTool.Height + 45
    lblAtt.Left = 30
    lblAtt.Width = lblTool.Width
    
    mshAtt.Top = lblAtt.Top + lblAtt.Height + 15
    mshAtt.Left = picAtt.ScaleLeft
    mshAtt.Width = picAtt.ScaleWidth
    mshAtt.Height = (ScaleHeight - cbrH - staH) - (lblTool.Height + 30) - (lblAtt.Height + 45) - lblNote.Height - picM.Height - tbrTool.Height
    
    picM.Left = picAtt.ScaleLeft
    picM.Top = mshAtt.Top + mshAtt.Height
    picM.Width = picAtt.ScaleWidth
    
    lblNote.Top = picM.Top + picM.Height
    lblNote.Left = picAtt.ScaleLeft
    lblNote.Width = picAtt.ScaleWidth
    
    Call ShowSize
    Call ShowScroll
    If Not scrHsc.Enabled Then DrawRuler picRulerH
    If Not scrVsc.Enabled Then DrawRuler picRulerV
    
    Call NoneEdit
    
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim intR As VbMsgBoxResult
    Dim strInfo As String
    
    If Not BlnSave Then
        intR = MsgBox("�����е�ǰ�޸�������δ����,Ҫ������", vbQuestion + vbYesNoCancel, App.Title)
        If intR = vbCancel Then 'ȡ���˳�
            Cancel = 1: Exit Sub
        ElseIf intR = vbYes Then '�˳�ǰ�ȱ���
            strInfo = CheckData
            If strInfo <> "" Then MsgBox strInfo, vbInformation, App.Title: Cancel = 1: Exit Sub
            
            strInfo = CheckHead
            If strInfo <> "" Then MsgBox strInfo, vbInformation, App.Title: Cancel = 1: Exit Sub
            
            strInfo = CheckArea
            If strInfo <> "" Then MsgBox strInfo, vbInformation, App.Title: Cancel = 1: Exit Sub
            
            strInfo = CheckPars
            If strInfo <> "" Then MsgBox strInfo, vbInformation, App.Title: Exit Sub
    
            Call SelClear
            Refresh
            If Not SaveReport(lngRPTID, objReport, sta.Panels(2)) Then
                MsgBox "������ʧ��,�����Ա��������", vbInformation, App.Title
                Cancel = 1: Exit Sub
            End If
            Call UpdatePriv
            
            BlnSave = True
            gblnModi = True
            Refresh
            
            If Not CheckReportPriv(lngRPTID) Then
                If MsgBox("��û��Ȩ�޲�ѯ�ñ���ĳЩ����Դ�еĶ�����Ȼ��������" & vbCrLf & _
                          "�ر��棬������������Щ����֮ǰ�㲻������ʹ�øñ���" & vbCrLf & _
                          "ȷʵҪ�˳���ƻ�����", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then
                    Cancel = 1: Exit Sub
                End If
            End If
        End If
    Else
        If Not CheckReportPriv(lngRPTID) Then
            If MsgBox("��û��Ȩ�޲�ѯ�ñ���ĳЩ����Դ�еĶ���" & vbCrLf & _
                   "����������Щ����ǰ�㲻������ʹ�øñ���" & vbCrLf & _
                   "ȷʵҪ�˳���ƻ�����", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then
                Cancel = 1: Exit Sub
            End If
        End If
    End If
    
    lngRPTID = 0
    mblnNotModiData = False
    strMenu = ""
    Unload frmFlash
    
    SaveWinState Me, App.ProductName
End Sub

Private Sub lblshp_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lngPreX = X: lngPreY = Y
    Set mobjMove = Nothing: mlngX = 0: mlngY = 0
    If Button = 1 Then
        If Shift = 2 Then
            If Mid(Shp(Index).Tag, 1, 2) = "" Then
                Call SelItem(Index, True) '��ѡ
                If GetSelNum() = 1 Then
                    Call ShowAttrib(Index) 'ֻѡ��һ������ʾ����
                Else
                    Call ShowAttrib '��ѡʱ����ʾ����
                End If
            Else
                Call SelItem(Index, False) '��ѡ
                If GetSelNum() = 1 Then
                    Call ShowAttrib(intCurID) 'ֻѡ��һ������ʾ����(ѡ�еĲ�һ���Ǹÿؼ�)
                Else
                    Call ShowAttrib '��ѡʱ����ʾ����
                End If
            End If
        Else
            If Mid(Shp(Index).Tag, 1, 2) = "" Then
                Call SelClear
                Call SelItem(Index, True)
                Call ShowAttrib(Index) 'ֻѡ��һ������ʾ����
            End If
        End If
    ElseIf Button = 2 Then
        PopupMenu mnuFormat, 2
    End If
End Sub

Private Sub lblshp_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ObjSel As Shape

    If objReport.Items("_" & Index).ϵͳ Then Exit Sub
    lblshp(Index).ZOrder 1
    DrawXY X + Shp(Index).Left, Y + Shp(Index).Top

    If Button = 1 And Mid(Shp(Index).Tag, 1, 2) <> "" Then
        If blnLock Then Exit Sub
        Call MoveSelect(X - lngPreX, Y - lngPreY)
        If objReport.Items("_" & Index).���� = 10 Then Call DrawFrame(lblshp(Index))
        If GetSelNum() = 1 Then ShowAttrib Index
'    Else
'�������ƶ�ʱ��Shp��Χ������Ԫ����˸������
'        'zyb#Add
'        '���������ϻ�����Ԫ��
'        Set ObjSel = Shp(Index)
'
'        If X < 100 Or Y < 100 Or X > ObjSel.Width - 100 Or Y > ObjSel.Height - 100 Then
'            lblshp(Index).MousePointer = 99
'        Else
'            lblshp(Index).MousePointer = IIF(bytCurTool <> 0, 2, 0)
'            ObjSel.ZOrder 1
'            lblshp(Index).ZOrder 1
'            picPaper_MouseMove Button, Shift, X + Shp(Index).Left, Y + Shp(Index).Top
'        End If
    End If
End Sub

Private Sub lblshp_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    picPaper.Cls
    If Not mobjMove Is Nothing Then
        If Not mobjMove Is Shp(Index).Container Then
            If objReport.Items("_" & Index).���� = "" Then
                If GetDataSouse(objReport.Items("_" & Index).����) <> "" And UCase(mobjMove.name) = "PIC" Then
                    If objReport.Items("_" & mobjMove.Index).����Դ = "" Then
                        If MsgBox("��ǰ��Ƭδ������Դ���󶨺󽫷����ӡ���ſ�Ƭ������Դ�д���""�����ʶ""�ֶ���""�����ʶ""��ͬ��Ϊһ��,����һ������Ϊһ�飻" & vbCrLf & _
                             "������ֻ��ӡһ�ſ�Ƭ���Ƿ������Դ""" & GetDataSouse(objReport.Items("_" & Index).����) & """?", vbQuestion + vbYesNo + vbDefaultButton1, App.Title) = vbYes Then
                            objReport.Items("_" & mobjMove.Index).����Դ = GetDataSouse(objReport.Items("_" & Index).����)
                        End If
                    End If
                End If
                Set Shp(Index).Container = mobjMove
                Shp(Index).Top = mlngY: Shp(Index).Left = mlngX
                Set lblshp(Index).Container = mobjMove
                lblshp(Index).Top = mlngY: lblshp(Index).Left = mlngX
                If UCase(mobjMove.name) = "PIC" Then
                    objReport.Items("_" & Index).��ID = mobjMove.Index
                Else
                    objReport.Items("_" & Index).��ID = 0
                End If
                objReport.Items("_" & Index).X = mlngX: objReport.Items("_" & Index).Y = mlngY
                Set mobjMove = Nothing: mlngX = 0: mlngY = 0
                Call ShowAttrib(Index)
            End If
        End If
    End If
End Sub

Private Sub Img_DblClick(Index As Integer)
    Dim ObjSel As Image, ObjLeft As Single, ObjTop As Single
    
    If GetSelNum <> 1 Then Exit Sub
    Set ObjSel = img(Index)
    With ObjSel
        ObjLeft = .Left
        ObjTop = .Top
        .Stretch = False
        .Stretch = True
        .Width = .Width * sgnMode
        .Height = .Height * sgnMode
        .Left = ObjLeft
        .Top = ObjTop
    End With
    
    With objReport.Items("_" & Index)
        .X = Format(ObjSel.Left / sgnMode, "0.00")
        .Y = Format(ObjSel.Top / sgnMode, "0.00")
        .H = Format(ObjSel.Height / sgnMode, "0.00")
        .W = Format(ObjSel.Width / sgnMode, "0.00")
    End With
    Call SelItem(Index, False)
    Call SelItem(Index, True)
End Sub

Private Sub Img_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not mobjMove Is Nothing Then
        If Not mobjMove Is img(Index).Container Then
            Set img(Index).Container = mobjMove
            img(Index).Top = mlngY: img(Index).Left = mlngX
            If UCase(mobjMove.name) = "PIC" Then
                objReport.Items("_" & Index).��ID = mobjMove.Index
            Else
                objReport.Items("_" & Index).��ID = 0
            End If
            objReport.Items("_" & Index).X = mlngX: objReport.Items("_" & Index).Y = mlngY
        End If
        Set mobjMove = Nothing: mlngX = 0: mlngY = 0
        Call ShowAttrib(Index)
    End If
End Sub

Private Sub ImgCode_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not mobjMove Is Nothing Then
        If Not mobjMove Is ImgCode(Index).Container Then
            If GetDataSouse(objReport.Items("_" & Index).����) <> "" And UCase(mobjMove.name) = "PIC" Then
                If objReport.Items("_" & mobjMove.Index).����Դ = "" Then
                    If MsgBox("��ǰ��Ƭδ������Դ���󶨺󽫷����ӡ���ſ�Ƭ������Դ�д���""�����ʶ""�ֶ���""�����ʶ""��ͬ��Ϊһ��,����һ������Ϊһ�飻" & vbCrLf & _
                         "������ֻ��ӡһ�ſ�Ƭ���Ƿ������Դ""" & GetDataSouse(objReport.Items("_" & Index).����) & """?", vbQuestion + vbYesNo + vbDefaultButton1, App.Title) = vbYes Then
                        objReport.Items("_" & mobjMove.Index).����Դ = GetDataSouse(objReport.Items("_" & Index).����)
                    End If
                End If
            End If
            Set ImgCode(Index).Container = mobjMove
            ImgCode(Index).Top = mlngY: ImgCode(Index).Left = mlngX
            If UCase(mobjMove.name) = "PIC" Then
                objReport.Items("_" & Index).��ID = mobjMove.Index
            Else
                objReport.Items("_" & Index).��ID = 0
            End If
            objReport.Items("_" & Index).X = mlngX: objReport.Items("_" & Index).Y = mlngY
        End If
        Set mobjMove = Nothing: mlngX = 0: mlngY = 0
        Call ShowAttrib(Index)
    End If
End Sub

Private Function GetDataSouse(ByVal str���� As String) As String
'���ܣ��������ݻ�ü������Դ����
    Dim i As Long, j As Long, k As Long
    
    i = InStr(str����, "]")
    j = InStr(str����, ".")
    k = InStr(str����, "[")
    If i > k And i > j And i <> 0 And k <> 0 And j <> 0 Then
        GetDataSouse = Mid(str����, k + 1, j - k - 1)
    End If
End Function

Private Sub lblLine_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not mobjMove Is Nothing Then
        If Not mobjMove Is lblLine(Index).Container Then
            Set lblLine(Index).Container = mobjMove
            lblLine(Index).Top = mlngY: lblLine(Index).Left = mlngX
            If UCase(mobjMove.name) = "PIC" Then
                objReport.Items("_" & Index).��ID = mobjMove.Index
            Else
                objReport.Items("_" & Index).��ID = 0
            End If
            objReport.Items("_" & Index).X = mlngX: objReport.Items("_" & Index).Y = mlngY
        End If
        Set mobjMove = Nothing: mlngX = 0: mlngY = 0
        Call ShowAttrib(Index)
    End If
End Sub

Private Sub LblSize_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not mobjMove Is Nothing Then
        If UCase(mobjMove.name) = "PIC" Then
            If mobjMove.Index <> objReport.Items("_" & intCurID).��ID Then
                objReport.Items("_" & intCurID).��ID = mobjMove.Index
                objReport.Items("_" & intCurID).X = mlngX: objReport.Items("_" & intCurID).Y = mlngY
                Call AdjustCoordinate(True)
            End If
        Else
            If objReport.Items("_" & intCurID).��ID <> 0 Then
                objReport.Items("_" & intCurID).��ID = 0
                objReport.Items("_" & intCurID).X = mlngX: objReport.Items("_" & intCurID).Y = mlngY
                Call AdjustCoordinate(True)
            End If
        End If
  
        Set mobjMove = Nothing: mlngX = 0: mlngY = 0
        Call ShowAttrib(intCurID)
    End If
End Sub

Private Sub mnuEdit_History_Click()
'���ܣ��鿴��ʷ����Դ
    Dim rsTmp As Recordset, strSQL As String
    Dim strKey As String, strInfo As String
    Dim strPreName As String, strDBName As String
    
    If tvwSQL.Nodes.count = 1 Then
        MsgBox "��ǰû������Դ��", vbInformation, App.Title: Exit Sub
    End If
    If tvwSQL.SelectedItem.Key = "Root" Then
        MsgBox "��ѡ��Ҫ�鿴������Դ��", vbInformation, App.Title: Exit Sub
    End If
    
    If tvwSQL.SelectedItem.Parent.Key <> "Root" Then
        strKey = tvwSQL.SelectedItem.Parent.Key
    Else
        strKey = tvwSQL.SelectedItem.Key
    End If
    strPreName = objReport.Datas(strKey).����
    strDBName = objReport.Datas(strKey).ԭ����
    
    On Error GoTo errH
    strSQL = "select 1 from zlRPTSQLsHistory Where ����ID=[1] and ����Դ����=[2] And rownum<2"
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngRPTID, IIF(strDBName = "", strPreName, strDBName))
    If rsTmp.RecordCount = 0 Then
        MsgBox "��ǰ����Դû����ʷ��¼��", vbInformation, App.Title: Exit Sub
    End If
    
    Call frmSQLEdit.ShowMe(Me, IIF(glngSys <> 0, glngSys, objReport.ϵͳ), objReport.Datas(strKey), objReport.Datas, 1)
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mshAtt_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim i As Integer, j As Integer, k As Integer, intRow As Integer
    Dim ItemThis As RPTItem, StrCompare As String
    Dim DataThis As RPTData
    Dim lngLeft As Long, lngTop As Long
    Dim StrPar As String
    
    Call NoneEdit
    
    '����ѡ����ɫ
    mshAtt.Redraw = False
    mshAtt.Cell(flexcpBackColor, 1, 0, mshAtt.Rows - 1, 0) = mshAtt.BackColor
    mshAtt.Cell(flexcpForeColor, 1, 0, mshAtt.Rows - 1, 0) = mshAtt.ForeColor
    mshAtt.Cell(flexcpBackColor, 0, 0, 0, 1) = mshAtt.BackColorFixed
    mshAtt.Cell(flexcpForeColor, 0, 0, 0, 1) = mshAtt.ForeColorFixed
    

    mshAtt.Cell(flexcpBackColor, NewRow, 0) = mshAtt.BackColorSel
    mshAtt.Cell(flexcpForeColor, NewRow, 0) = mshAtt.ForeColorSel

    mshAtt.Redraw = True

    '����ע��
    Select Case mshAtt.TextMatrix(mshAtt.Row, 0)
        Case "����"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "Ԫ�ص�����,��Ϊ:����,����,��ǩ,ͼƬ,���,��Ƭ."
        Case "����"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "����Ԫ�ص����ƣ����ڲ��ն���."
        Case "����"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "��ͼ��Ԫ�صĸ�ʽ�����ݵ����ݽ�������"
        Case "����"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "��ǩ���ı����ݻ��Ӧ��������Ŀ."
        Case "X����"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "����Ԫ�����Ͻǵ����λ��,�Ժ���Ϊ��λ."
        Case "Y����"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "����Ԫ�����Ͻǵ��ϱ�λ��,�Ժ���Ϊ��λ."
        Case "���"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "����Ԫ�ص�������,�Ժ���Ϊ��λ."
        Case "�߶�"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "����Ԫ�ص�����߶�,�Ժ���Ϊ��λ."
        Case "�и�"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "�����ÿһ�еĸ߶�,�Ժ���Ϊ��λ."
        Case "����"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "��ǩ�����е�������ˮƽ�����ϵĶ��뷽ʽ."
        Case "��������ɫ"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "�����������������ɫ."
        Case "��ͷ����ɫ"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "����ͷ������������ɫ."
        Case "����ɫ"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "��������������ɫ."
        Case "ǰ��ɫ"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "����Ԫ�ص�������ɫ����������ɫ."
        Case "����ɫ"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "����Ԫ�صı�����ɫ."
        Case "�Զ�������С"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "���ñ�ǩ�����롢ͼ�εĳߴ��Ƿ��Զ�������С"
        Case "�Ӵ�"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "�������������ߵı߿��Ƿ�Ӵ�"
        Case "����߼Ӵ�"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "���ñ����������Ƿ�Ӵ�"
        Case "��״"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "���ÿ��ߵı߿���״"
        Case "����"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "����Ԫ�ص����������������."
        Case "�Զ�����"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "���ñ�ǩ�����ݹ���ʱ�Ƿ��Զ���С����ߴ���д�ӡ."
        Case "�߿�"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "ָ���Ƿ��ڱ�ǩ���ֱ�������Χ��һ���ο���."
        Case "����"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "��������������ݰ������Զ��������."
        Case "���ն���"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "�趨��ǩ��ָ�����ն����Ķ����ϵ."
        Case "����"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "�趨��ǩ�Ǳ�����Ǳ�����."
        Case "����"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "�趨��ǩ����ն����Ĳ��չ�ϵ."
        Case "��ʽ"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "���ñ�ǩ�������ֶε����������ʽ��,��VB��ʽ�ַ�����"
        Case "����Ԫ��"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "��ǰ�����ʽ�е����б���Ԫ��"
        Case "���ͼ��"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "��ǰ�����ʽͼ�������ģʽ"
        Case "Ʊ��"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "��ǰ�����Ƿ���Ʊ�ݵķ�ʽ�����������ӡ"
        Case "��ֹ��ʼʱ��"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "��ǰ���������ѯ��ʱ���(��ʼʱ��)��"
        Case "��ֹ����ʱ��"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "��ǰ���������ѯ��ʱ���(����ʱ��)��"
        Case "�ձ��ӡ"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "ģ�����ֱ�Ӵ�ӡʱ,���б������Ϊ���Ƿ���д�ӡ"
        Case "��ֽ̬��"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "�Ƿ���ݴ�ӡ�������Զ�����ֽ�Ÿ߶�"
        Case "��ӡ��"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "��ǰʹ�õĴ�ӡ��"
        Case "ֽ��"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "��ǰʹ�õ�ֽ������"
        Case "ֽ��"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "��ǰֽ�ŵķ���"
        Case "��ֽ��ʽ"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "��ǰֽ�ŵĽ�ֽ��ʽ"
        Case "����"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "����ӡ��Ԥ��ʱ�Ƿ��������"
        Case "����"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "����������ʱ�Ƿ��Զ�����,�������ͷ������ʼҪ����"
        Case "���ֱ���"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "ͼ������ʱ���Ƿ񱣳�ԭʼ�Ŀ�߱���"
        Case "����ͼ��"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "���ø�ͼ���Ƿ�����Ӱ���鱨��"
        Case "��������"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "�������������"
        Case "�����߿�"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "�������������Ŀ��(1-N)��������������Ŀ��"
        Case "��ʾ����"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "�Ƿ�������ͼ��������ʾ���������"
        Case "��У���"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "�Ƿ��29�����У���"
        Case "��ת����"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "����ͼ�����ʱ����ת����"
        Case "�о�"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "��������Ķ�������֮��ĵ��(0-100)"
        Case "����Դ"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "��ѡ��ǰ����һ�����������Դ��ȷ����Ƭ��̬��ӡ�ķ�ҳ"
        Case "���Ҽ��"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "��ҳ��̬��ӡ��Ƭʱ��ÿ�ſ�Ƭ�����Ҽ��,�Ժ���Ϊ��λ"
        Case "���¼��"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "��ҳ��̬��ӡ��Ƭʱ��ÿ�ſ�Ƭ�����¼��,�Ժ���Ϊ��λ"
        Case "����Դ�к�"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "�����ǩ�����а�������Դ�����Ʊ�ǩ��ʾ�ڼ��е�����"
        Case "�������"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "��ҳ��̬��ӡ��Ƭʱ�������ӡ�Ŀ�Ƭ����0Ϊ����Ӧ"
        Case "�������"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "��ҳ��̬��ӡ��Ƭʱ�������ӡ�Ŀ�Ƭ����0Ϊ����Ӧ"
        Case "����"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "��ѡ��ǰԪ�����������������Ԫ�صķ�Χ����������(��Ƭ)֮�ڲ���������"
        Case "��������"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "�������ѯ�˱���ʱ��˫����Ԫ�ؽ��������õĲ�����Դ��Ϊ��������Ĳ�����ִ�й�������"
        Case "ˮƽ��ת"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "ǩ��Ԫ����Ԥ������ӡʱˮƽ��ת��"
        Case Else
            lblNote.Caption = ""
    End Select
    
    If blnLock Then Exit Sub
    '�����ϵͳ������Ŀ����ֻ�����������弰���ն���
    If InStr(1, "����,���ն���,����,����", mshAtt.TextMatrix(mshAtt.Row, 0)) = 0 And intCurID <> 0 Then
        If objReport.Items("_" & intCurID).ϵͳ Then Exit Sub
    End If
    
    '����������ӱ����ӱ�,���������ø߶ȼ��и�
    If InStr(1, "�߶�,�и�,����,����ɫ,����ɫ", mshAtt.TextMatrix(mshAtt.Row, 0)) <> 0 And intCurID <> 0 Then
        If objReport.Items("_" & intCurID).���� <> "" And objReport.Items("_" & intCurID).���� = 5 Then Exit Sub
    End If
    
    '�����ͼ���ֶΣ���δ�������ݣ����������ã������˳�
    If mshAtt.TextMatrix(mshAtt.Row, 0) = "����" Then
        If objReport.Items("_" & intCurID).���� = 11 Then
            cmdAtt.Top = mshAtt.Top + mshAtt.CellTop
            cmdAtt.Left = mshAtt.Left + mshAtt.Cell(flexcpLeft, NewRow, 1) + mshAtt.CellWidth - cmdAtt.Width
            cmdAtt.Visible = True
            Exit Sub
        ElseIf objReport.Items("_" & intCurID).���� = 2 Then
            Dim strNodeName As String, NodeThis As Node
            '�����adLongVarBinary���ֶ�,�������޸�
            i = InStr(1, mshAtt.TextMatrix(mshAtt.Row, 1), "]")
            j = InStr(1, mshAtt.TextMatrix(mshAtt.Row, 1), ".")
            k = InStr(1, mshAtt.TextMatrix(mshAtt.Row, 1), "[")
            If i > k And i > j And i <> 0 And k <> 0 Then
                strNodeName = Mid(mshAtt.TextMatrix(mshAtt.Row, 1), j + 1, i - j - 1)
                
                For Each NodeThis In tvwSQL.Nodes
                    If mdlPublic.GetStdNodeText(NodeThis.Text) = strNodeName And IsType(Val(NodeThis.Tag), adLongVarBinary) Then Exit Sub
                Next
            End If
        End If
    End If
    
    mshAtt.ColWidth(1) = mshAtt.Cell(flexcpWidth, 0, 1)
    
    '�༭����
    Select Case mshAtt.TextMatrix(mshAtt.Row, 0)
        Case "����", "����", "X����", "Y����", "���", "�߶�", "�и�", "����", "�ֺ�", "��ʽ", "�о�" _
            , "���Ҽ��", "���¼��", "����Դ�к�", "�������", "�������"
            If intCurID = 0 Then Exit Sub
            If objReport.Items("_" & intCurID).���� <> 2 Or mshAtt.TextMatrix(mshAtt.Row, 0) <> "����" Then
                txtAtt.MaxLength = 0
                If InStr("X����,Y����,���,�߶�,�и�,���Ҽ��,���¼��,Դ�к�,�������,�������", mshAtt.TextMatrix(mshAtt.Row, 0)) > 0 Then
                    txtAtt.MaxLength = 7
                ElseIf mshAtt.TextMatrix(mshAtt.Row, 0) = "����" Then
                    '��������,���������÷���
                    If objReport.Items("_" & intCurID).���� <> "" Then Exit Sub
                    For Each ItemThis In objReport.Items
                        If ItemThis.��ʽ�� = mbytCurrFmt And InStr(1, "4,5", ItemThis.����) <> 0 Then
                            If ItemThis.��ʽ�� = mbytCurrFmt And ItemThis.Key <> intCurID And ItemThis.���� = objReport.Items("_" & intCurID).���� And InStr(1, "4,5", ItemThis.����) <> 0 Then Exit Sub
                        End If
                    Next
                    txtAtt.MaxLength = 2
                ElseIf mshAtt.TextMatrix(mshAtt.Row, 0) = "�ֺ�" Then
                    txtAtt.MaxLength = 7
                ElseIf mshAtt.TextMatrix(mshAtt.Row, 0) = "��ʽ" Then
                    txtAtt.MaxLength = 50
                ElseIf mshAtt.TextMatrix(mshAtt.Row, 0) = "�о�" Then
                    txtAtt.MaxLength = 3
                End If
                If InStr("X����,Y����", mshAtt.TextMatrix(mshAtt.Row, 0)) > 0 And objReport.Items("_" & intCurID).���� <> "" Then
                    If Not (mshAtt.TextMatrix(mshAtt.Row, 0) = "Y����" And objReport.Items("_" & intCurID).���� = 2) Then Exit Sub
                End If
                
                txtAtt.Left = mshAtt.Left + mshAtt.Cell(flexcpLeft, NewRow, 1) + 30
                txtAtt.Top = mshAtt.Top + mshAtt.CellTop + (mshAtt.CellHeight - txtAtt.Height) / 2
                txtAtt.Width = mshAtt.ColWidth(1) - 60
                txtAtt.Text = mshAtt.TextMatrix(mshAtt.Row, 1)
                txtAtt.Visible = True: txtAtt.SetFocus
            Else
                '��ǩ����ѡ���õ�����
                cboText.Clear
                For i = 1 To objReport.Datas.count
                    For j = 1 To objReport.Datas(i).Pars.count
                        If InStr(StrPar, objReport.Datas(i).Pars(j).����) = 0 Then
                            StrPar = StrPar & "|" & objReport.Datas(i).Pars(j).����
                        End If
                    Next
                Next
                StrPar = Mid(StrPar, 2)
                If StrPar <> "" Then
                    For i = 0 To UBound(Split(StrPar, "|"))
                        cboAtt.AddItem "[=" & Split(StrPar, "|")(i) & "]"
                    Next
                End If
                cboText.AddItem "[����Ա����]"
                cboText.AddItem "[����Ա���]"
                cboText.AddItem "[��λ����]"
                cboText.AddItem "[ҳ��]"
                cboText.AddItem "[ҳ��]"
                cboText.AddItem "[yyyy-mm-dd]"
                cboText.AddItem "[yyyy-mm-dd HH:MM]"
                cboText.AddItem "[yyyy-mm-dd HH:MM:SS]"
                
                cboText.Left = mshAtt.Left + mshAtt.Cell(flexcpLeft, NewRow, 1)
                cboText.Top = mshAtt.Top + mshAtt.CellTop
                cboText.Width = mshAtt.ColWidth(1) - 60
                cboText.Text = objReport.Items("_" & intCurID).����
                cboText.Visible = True:  cboText.SetFocus
            End If
        Case "����"
            cboAtt.Clear
            cboAtt.AddItem "�����"
            cboAtt.AddItem "�ж���"
            cboAtt.AddItem "�Ҷ���"
            cboAtt.Left = mshAtt.Left + mshAtt.Cell(flexcpLeft, NewRow, 1)
            cboAtt.Top = mshAtt.Top + mshAtt.CellTop
            cboAtt.Width = mshAtt.ColWidth(1)
            cboAtt.ListIndex = objReport.Items("_" & intCurID).����
            cboAtt.Visible = True:  cboAtt.SetFocus
        Case "��״"
            cboAtt.Clear
            cboAtt.AddItem "����"
            cboAtt.AddItem "Բ��"
            cboAtt.Left = mshAtt.Left + mshAtt.Cell(flexcpLeft, NewRow, 1)
            cboAtt.Top = mshAtt.Top + mshAtt.CellTop
            cboAtt.Width = mshAtt.ColWidth(1)
            cboAtt.ListIndex = IIF(objReport.Items("_" & intCurID).�߿�, 1, 0)
            cboAtt.Visible = True:  cboAtt.SetFocus
        Case "ǰ��ɫ", "����ɫ", "����", "��������ɫ", "����", "����ɫ", "��ͷ����ɫ"
            cmdAtt.Top = mshAtt.Top + mshAtt.CellTop
            cmdAtt.Left = mshAtt.Left + mshAtt.Cell(flexcpLeft, NewRow, 1) + mshAtt.ColWidth(1) - cmdAtt.Width
            cmdAtt.Visible = True
        Case "��������"
            cmdAtt.Top = mshAtt.Top + mshAtt.CellTop
            cmdAtt.Left = mshAtt.Left + mshAtt.Cell(flexcpLeft, NewRow, 1) + mshAtt.ColWidth(1) - cmdAtt.Width
            cmdAtt.Visible = True
            txtAtt.Left = mshAtt.Left + mshAtt.Cell(flexcpLeft, NewRow, 1) + 30
            txtAtt.Top = mshAtt.Top + mshAtt.CellTop + (mshAtt.CellHeight - txtAtt.Height) / 2
            txtAtt.Width = mshAtt.ColWidth(1) - 60
            txtAtt.Text = mshAtt.TextMatrix(mshAtt.Row, 1)
            If mshAtt.TextMatrix(mshAtt.Row, 1) <> "" Then
                txtAtt.Visible = False
            Else
                txtAtt.Visible = True: txtAtt.SetFocus
            End If
        Case "����Ԫ��"
            Call GetAllElement
            '����λ��
            cboAtt.Left = mshAtt.Left + mshAtt.Cell(flexcpLeft, NewRow, 1)
            cboAtt.Top = mshAtt.Top + mshAtt.CellTop
            cboAtt.Width = mshAtt.ColWidth(1)
            cboAtt.Visible = True:  cboAtt.SetFocus
        Case "��ֹ��ʼʱ��"
            dtpAtt.Left = mshAtt.Left + mshAtt.Cell(flexcpLeft, NewRow, 1)
            dtpAtt.Top = mshAtt.Top + mshAtt.CellTop
            dtpAtt.Width = mshAtt.ColWidth(1)
            dtpAtt.Value = Format(mshAtt.TextMatrix(mshAtt.Row, 1), "HH:mm:ss")
            dtpAtt.Visible = True:  dtpAtt.SetFocus
        Case "��ֹ����ʱ��"
            dtpAtt.Left = mshAtt.Left + mshAtt.Cell(flexcpLeft, NewRow, 1)
            dtpAtt.Top = mshAtt.Top + mshAtt.CellTop
            dtpAtt.Width = mshAtt.ColWidth(1)
            dtpAtt.Value = Format(mshAtt.TextMatrix(mshAtt.Row, 1), "HH:mm:ss")
            dtpAtt.Visible = True:  dtpAtt.SetFocus
        Case "���ͼ��"
            blnModify = False
            Call LoadOutChart
            
            '����λ��
            cboAtt.Left = mshAtt.Left + mshAtt.Cell(flexcpLeft, NewRow, 1)
            cboAtt.Top = mshAtt.Top + mshAtt.CellTop
            cboAtt.Width = mshAtt.ColWidth(1)
            cboAtt.Visible = True:  cboAtt.SetFocus
            
            Call LocateOutChart
            blnModify = True
        Case "���ն���"
            '��������ѷ���,���˳�
            If objReport.Items("_" & intCurID).���� > 1 Then Exit Sub
            '���б�񸽼��ڱ���,�򱾱������ò��ն��󼰷���
            For Each ItemThis In objReport.Items
                If ItemThis.��ʽ�� = mbytCurrFmt And InStr(1, "4,5", ItemThis.����) <> 0 Then
                    If ItemThis.Key <> intCurID And ItemThis.���� = objReport.Items("_" & intCurID).���� Then Exit Sub
                End If
            Next

            cboAtt.Clear
            cboAtt.AddItem ""
            
            '���������
            For Each ItemThis In objReport.Items
                If ItemThis.��ʽ�� = mbytCurrFmt Then
                    Select Case objReport.Items("_" & intCurID).����
                    Case "2"
                        If InStr(1, "|4,|5,", "|" & ItemThis.���� & ",") <> 0 And ItemThis.���� = "" Then
                            cboAtt.AddItem ItemThis.����
                        End If
                    Case "4", "5"
                        If ItemThis.���� < 2 And CheckTableProperty(ItemThis) Then cboAtt.AddItem ItemThis.����
                    End Select
                End If
            Next
            
            '����λ��
            cboAtt.Left = mshAtt.Left + mshAtt.Cell(flexcpLeft, NewRow, 1)
            cboAtt.Top = mshAtt.Top + mshAtt.CellTop
            cboAtt.Width = mshAtt.ColWidth(1)
            cboAtt.Visible = True:  cboAtt.SetFocus
            
            StrCompare = objReport.Items("_" & intCurID).����
            For i = 0 To cboAtt.ListCount - 1
                If cboAtt.List(i) = StrCompare Then
                    cboAtt.ListIndex = i
                    mshAtt.TextMatrix(mshAtt.Row, 1) = StrCompare
                    Exit For
                End If
            Next
            If cboAtt.Text <> StrCompare Then cboAtt.ListIndex = 0
        Case "����Դ"

            cboAtt.Clear
            cboAtt.AddItem ""
            
            '���������
            For Each DataThis In objReport.Datas
                If DataThis.���� = 0 Then
                    cboAtt.AddItem DataThis.����
                End If
            Next
            
            '����λ��
            cboAtt.Left = mshAtt.Left + mshAtt.Cell(flexcpLeft, NewRow, 1)
            cboAtt.Top = mshAtt.Top + mshAtt.CellTop
            cboAtt.Width = mshAtt.ColWidth(1)
            cboAtt.Visible = True:  cboAtt.SetFocus
            
            StrCompare = objReport.Items("_" & intCurID).����Դ
            For i = 0 To cboAtt.ListCount - 1
                If cboAtt.List(i) = StrCompare Then
                    cboAtt.ListIndex = i
                    mshAtt.TextMatrix(mshAtt.Row, 1) = StrCompare
                    Exit For
                End If
            Next
            If cboAtt.Text <> StrCompare Then cboAtt.ListIndex = 0
        Case "����"
            cboAtt.Clear
            cboAtt.AddItem "ҳ��"
            
            '���������
            If objReport.Items("_" & intCurID).��ID <> 0 Then
                lngLeft = objReport.Items("_" & objReport.Items("_" & intCurID).��ID).X
                lngTop = objReport.Items("_" & objReport.Items("_" & intCurID).��ID).Y
            End If
            For Each ItemThis In objReport.Items
                If ItemThis.���� = 14 And ItemThis.��ʽ�� = mbytCurrFmt Then
                    If objReport.Items("_" & intCurID).Y + lngTop >= ItemThis.Y And objReport.Items("_" & intCurID).X + lngLeft >= ItemThis.X And _
                            objReport.Items("_" & intCurID).H + objReport.Items("_" & intCurID).Y + lngTop <= ItemThis.Y + ItemThis.H And _
                            objReport.Items("_" & intCurID).W + objReport.Items("_" & intCurID).X + lngLeft <= ItemThis.X + ItemThis.W Then
                        cboAtt.AddItem ItemThis.����
                        cboAtt.ItemData(cboAtt.NewIndex) = ItemThis.id
                    ElseIf objReport.Items("_" & intCurID).��ID = ItemThis.id Then
                        cboAtt.AddItem ItemThis.����
                        cboAtt.ItemData(cboAtt.NewIndex) = ItemThis.id
                    End If
                End If
            Next
            
            '����λ��
            cboAtt.Left = mshAtt.Left + mshAtt.Cell(flexcpLeft, NewRow, 1)
            cboAtt.Top = mshAtt.Top + mshAtt.CellTop
            cboAtt.Width = mshAtt.ColWidth(1)
            cboAtt.Visible = True:  cboAtt.SetFocus
            
            If objReport.Items("_" & intCurID).��ID = 0 Then
                StrCompare = "ҳ��"
            Else
                StrCompare = objReport.Items("_" & objReport.Items("_" & intCurID).��ID).����
            End If
            For i = 0 To cboAtt.ListCount - 1
                If cboAtt.List(i) = StrCompare Then
                    cboAtt.ListIndex = i
                    mshAtt.TextMatrix(mshAtt.Row, 1) = StrCompare
                    Exit For
                End If
            Next
            If cboAtt.Text <> StrCompare Then cboAtt.ListIndex = 0
        Case "����"
            If mshAtt.TextMatrix(GetRow("���ն���"), 1) = "" Then
                mshAtt.TextMatrix(GetRow("����"), 1) = "����"
                Exit Sub
            End If
            
            cboAtt.Clear
            cboAtt.AddItem "������"
            cboAtt.AddItem "������"
            
            '����λ��
            cboAtt.Left = mshAtt.Left + mshAtt.Cell(flexcpLeft, NewRow, 1)
            cboAtt.Top = mshAtt.Top + mshAtt.CellTop
            cboAtt.Width = mshAtt.ColWidth(1)
            cboAtt.Visible = True:  cboAtt.SetFocus
            
            StrCompare = Mid(objReport.Items("_" & intCurID).����, 1, 1)
            StrCompare = IIF(StrCompare = "2", "������", "������")
            For i = 0 To cboAtt.ListCount - 1
                If cboAtt.List(i) = StrCompare Then
                    cboAtt.ListIndex = i
                    mshAtt.TextMatrix(mshAtt.Row, 1) = StrCompare
                    Exit For
                End If
            Next
            If cboAtt.Text <> StrCompare Then cboAtt.ListIndex = 0
        Case "����"
            If mshAtt.TextMatrix(GetRow("���ն���"), 1) = "" Then
                mshAtt.TextMatrix(GetRow("����"), 1) = "����"
                Exit Sub
            End If
            
            cboAtt.Clear
            Select Case objReport.Items("_" & intCurID).����
            Case 2
                cboAtt.AddItem "����"
                cboAtt.AddItem "����"
                cboAtt.AddItem "����"
            Case 4
                cboAtt.AddItem "����"
            Case 5
                cboAtt.AddItem "������"
            End Select
            
            '����λ��
            cboAtt.Left = mshAtt.Left + mshAtt.Cell(flexcpLeft, NewRow, 1)
            cboAtt.Top = mshAtt.Top + mshAtt.CellTop
            cboAtt.Width = mshAtt.ColWidth(1)
            cboAtt.Visible = True:  cboAtt.SetFocus
            
            Select Case objReport.Items("_" & intCurID).����
            Case 2
                StrCompare = Mid(objReport.Items("_" & intCurID).����, 2)
                StrCompare = IIF(StrCompare = "1", "����", IIF(StrCompare = "2", "����", "����"))
                For i = 0 To cboAtt.ListCount - 1
                    If cboAtt.List(i) = StrCompare Then
                        cboAtt.ListIndex = i
                        mshAtt.TextMatrix(mshAtt.Row, 1) = StrCompare
                        Exit For
                    End If
                Next
                If cboAtt.Text <> StrCompare Then cboAtt.ListIndex = 0
            Case 4, 5
                StrCompare = IIF(objReport.Items("_" & intCurID).���� = "1", "����", "������")
                For i = 0 To cboAtt.ListCount - 1
                    If cboAtt.List(i) = StrCompare Then
                        cboAtt.ListIndex = i
                        mshAtt.TextMatrix(mshAtt.Row, 1) = StrCompare
                        Exit For
                    End If
                Next
                If cboAtt.Text <> StrCompare Then cboAtt.ListIndex = 0
            End Select
        Case "��������"
            cboAtt.Clear
            cboAtt.AddItem "Code 128(����)": cboAtt.ItemData(cboAtt.NewIndex) = 1
            cboAtt.AddItem "Code 128 Auto": cboAtt.ItemData(cboAtt.NewIndex) = 3
            cboAtt.AddItem "Code 39": cboAtt.ItemData(cboAtt.NewIndex) = 2
            cboAtt.AddItem "QR Code": cboAtt.ItemData(cboAtt.NewIndex) = 10
            For i = 0 To cboAtt.ListCount - 1
                If cboAtt.ItemData(i) = objReport.Items("_" & intCurID).��� Then
                    CboSetIndex cboAtt.hwnd, i: Exit For
                End If
            Next
            If cboAtt.ListIndex = -1 Then CboSetIndex cboAtt.hwnd, 0
            
            cboAtt.Left = mshAtt.Left + mshAtt.Cell(flexcpLeft, NewRow, 1)
            cboAtt.Top = mshAtt.Top + mshAtt.CellTop
            cboAtt.Width = mshAtt.ColWidth(1)
            cboAtt.Visible = True:  cboAtt.SetFocus
        Case "�����߿�"
            cboAtt.Clear
            For i = 1 To 10
                cboAtt.AddItem i
            Next
            If objReport.Items("_" & intCurID).�и� <= 0 Then
                CboSetIndex cboAtt.hwnd, 1
            Else
                CboSetIndex cboAtt.hwnd, objReport.Items("_" & intCurID).�и� - 1
            End If
            
            cboAtt.Left = mshAtt.Left + mshAtt.Cell(flexcpLeft, NewRow, 1)
            cboAtt.Top = mshAtt.Top + mshAtt.CellTop
            cboAtt.Width = mshAtt.ColWidth(1)
            cboAtt.Visible = True:  cboAtt.SetFocus
        Case "��ת����"
            cboAtt.Clear
            cboAtt.AddItem "����ת"
            cboAtt.AddItem "˳ʱ��90��"
            cboAtt.AddItem "��ʱ��90��"
            CboSetIndex cboAtt.hwnd, Val(Mid(objReport.Items("_" & intCurID).��ͷ, 3, 1))
            
            cboAtt.Left = mshAtt.Left + mshAtt.Cell(flexcpLeft, NewRow, 1)
            cboAtt.Top = mshAtt.Top + mshAtt.CellTop
            cboAtt.Width = mshAtt.ColWidth(1)
            cboAtt.Visible = True:  cboAtt.SetFocus
    End Select
End Sub

Private Sub mshAtt_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    txtAtt.Visible = False
    cmdAtt.Visible = False
    cboAtt.Visible = False
    mshAtt.SetFocus
End Sub

Private Sub pic_DblClick(Index As Integer)
    Dim ObjSel As PictureBox, ObjLeft As Single, ObjTop As Single
    
    If GetSelNum <> 1 Then Exit Sub
    Set ObjSel = pic(Index)
    With ObjSel
        ObjLeft = .Left
        ObjTop = .Top
        .Width = .Width * sgnMode
        .Height = .Height * sgnMode
        .Left = ObjLeft
        .Top = ObjTop
    End With
    
    With objReport.Items("_" & Index)
        .X = Format(ObjSel.Left / sgnMode, "0.00")
        .Y = Format(ObjSel.Top / sgnMode, "0.00")
        .H = Format(ObjSel.Height / sgnMode, "0.00")
        .W = Format(ObjSel.Width / sgnMode, "0.00")
    End With
    Call SelItem(Index, False)
    Call SelItem(Index, True)
End Sub

Private Sub Img_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lngPreX = X: lngPreY = Y
    Set mobjMove = Nothing: mlngX = 0: mlngY = 0
    If Button = 1 Then
        If Shift = 2 Then
            If Mid(img(Index).Tag, 1, 2) = "" Then
                Call SelItem(Index, True) '��ѡ
                If GetSelNum() = 1 Then
                    Call ShowAttrib(Index) 'ֻѡ��һ������ʾ����
                Else
                    Call ShowAttrib '��ѡʱ����ʾ����
                End If
            Else
                Call SelItem(Index, False) '��ѡ
                If GetSelNum() = 1 Then
                    Call ShowAttrib(intCurID) 'ֻѡ��һ������ʾ����(ѡ�еĲ�һ���Ǹÿؼ�)
                Else
                    Call ShowAttrib '��ѡʱ����ʾ����
                End If
            End If
        Else
            If Mid(img(Index).Tag, 1, 2) = "" Then
                Call SelClear
                Call SelItem(Index, True)
                Call ShowAttrib(Index) 'ֻѡ��һ������ʾ����
            End If
        End If
    ElseIf Button = 2 Then
        PopupMenu mnuFormat, 2
    End If
End Sub

Private Sub pic_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    If UCase(Source.name) = "TVWSQL" Then
        selArea.Left = X: selArea.Top = Y
        Call AddReportItem(True, pic(Index))
        BlnSave = False
    End If

End Sub

Private Sub pic_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
    DrawXY CLng(X), CLng(Y)
    If UCase(Source.name) = "TVWSQL" Then
        If State = 1 Then
            Set tvwSQL.DragIcon = lvwPar.DragIcon
        ElseIf State = 0 Then
            If tvwSQL.SelectedItem.Children = 0 Then
                Set tvwSQL.DragIcon = scrHsc.DragIcon
            Else
                Set tvwSQL.DragIcon = scrVsc.DragIcon
            End If
        End If
    End If
End Sub

Private Sub pic_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lngPreX = X: lngPreY = Y
    If Button = 1 Then
        selArea.Left = X
        selArea.Top = Y
        blnDown = True
        If Shift = 2 Then
            If Mid(pic(Index).Tag, 1, 2) = "" Then
                Call SelItem(Index, True) '��ѡ
                If GetSelNum() = 1 Then
                    Call ShowAttrib(Index) 'ֻѡ��һ������ʾ����
                Else
                    Call ShowAttrib '��ѡʱ����ʾ����
                End If
            Else
                Call SelItem(Index, False) '��ѡ
                If GetSelNum() = 1 Then
                    Call ShowAttrib(intCurID) 'ֻѡ��һ������ʾ����(ѡ�еĲ�һ���Ǹÿؼ�)
                Else
                    Call ShowAttrib '��ѡʱ����ʾ����
                End If
            End If
        Else
            If Mid(pic(Index).Tag, 1, 2) = "" Then
                Call SelClear
                Call SelItem(Index, True)
                Call ShowAttrib(Index) 'ֻѡ��һ������ʾ����
            End If
        End If
    ElseIf Button = 2 Then
        PopupMenu mnuFormat, 2
    End If
End Sub

Private Sub Img_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ObjSel As Image
    
    If objReport.Items("_" & Index).ϵͳ Then Exit Sub
    DrawXY X + img(Index).Left, Y + img(Index).Top
    Set ObjSel = img(Index)
    
    If Button = 1 And Mid(ObjSel.Tag, 1, 2) <> "" Then
        If blnLock Then Exit Sub
        Call MoveSelect(X - lngPreX, Y - lngPreY)
        If GetSelNum() = 1 Then ShowAttrib Index
    Else
        ObjSel.MousePointer = 99
    End If
End Sub

Private Sub pic_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ObjSel As PictureBox
    Static PreX As Long, PreY As Long
    
    If bytCurTool <> 0 Then
        Call DrawXY(CLng(X), CLng(Y))
        
        '��ѡ�����
        If Button = 1 And blnDown And Shift <> 4 Then
            If PreX = Empty And PreY = Empty Then
                PreX = selArea.Left
                PreY = selArea.Top
            End If
            If bytCurTool <> 1 Then
                pic(Index).Line (selArea.Left, selArea.Top)-(PreX, PreY), picPaper.BackColor, B
                pic(Index).Line (selArea.Left, selArea.Top)-(X, Y), , B
            Else
                If Abs(X - selArea.Left) >= Abs(Y - selArea.Top) Then
                    '������
                    If bytLine = 2 Then pic(Index).Cls
                    pic(Index).Line (selArea.Left, selArea.Top)-(PreX, selArea.Top), picPaper.BackColor
                    pic(Index).Line (selArea.Left, selArea.Top)-(X, selArea.Top)
                    bytLine = 1
                Else
                    '������
                    If bytLine = 1 Then pic(Index).Cls
                    pic(Index).Line (selArea.Left, selArea.Top)-(selArea.Left, PreY), picPaper.BackColor
                    pic(Index).Line (selArea.Left, selArea.Top)-(selArea.Left, Y)
                    bytLine = 2
                End If
            End If
            PreX = X: PreY = Y
        End If
    Else
        If objReport.Items("_" & Index).ϵͳ Then Exit Sub
        DrawXY X + pic(Index).Left, Y + pic(Index).Top
        Set ObjSel = pic(Index)
        If Button = 1 And Mid(ObjSel.Tag, 1, 2) <> "" Then
            If blnLock Then Exit Sub
            Call MoveSelect(X - lngPreX, Y - lngPreY)
            If GetSelNum() = 1 Then ShowAttrib Index
        Else
            ObjSel.MousePointer = 99
        End If
    End If
End Sub

Private Sub ImgCode_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lngPreX = X: lngPreY = Y
    Set mobjMove = Nothing: mlngX = 0: mlngY = 0
    If Button = 1 Then
        If Shift = 2 Then
            If Mid(ImgCode(Index).Tag, 1, 2) = "" Then
                Call SelItem(Index, True) '��ѡ
                If GetSelNum() = 1 Then
                    Call ShowAttrib(Index) 'ֻѡ��һ������ʾ����
                Else
                    Call ShowAttrib '��ѡʱ����ʾ����
                End If
            Else
                Call SelItem(Index, False) '��ѡ
                If GetSelNum() = 1 Then
                    Call ShowAttrib(intCurID) 'ֻѡ��һ������ʾ����(ѡ�еĲ�һ���Ǹÿؼ�)
                Else
                    Call ShowAttrib '��ѡʱ����ʾ����
                End If
            End If
        Else
            If Mid(ImgCode(Index).Tag, 1, 2) = "" Then
                Call SelClear
                Call SelItem(Index, True)
                Call ShowAttrib(Index) 'ֻѡ��һ������ʾ����
            End If
        End If
    ElseIf Button = 2 Then
        PopupMenu mnuFormat, 2
    End If
End Sub

Private Sub ImgCode_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ObjSel As Image
    
    If objReport.Items("_" & Index).ϵͳ Then Exit Sub
    DrawXY X + ImgCode(Index).Left, Y + ImgCode(Index).Top
    Set ObjSel = ImgCode(Index)
    If Button = 1 And Mid(ObjSel.Tag, 1, 2) <> "" Then
        If blnLock Then Exit Sub
        Call MoveSelect(X - lngPreX, Y - lngPreY)
        If GetSelNum() = 1 Then ShowAttrib Index
    Else
        ObjSel.MousePointer = 99
    End If
End Sub

Private Sub lbl_DblClick(Index As Integer)
'���ܣ��Զ�������ǩ�Ĵ�С
    If Not blnLock And GetSelNum = 1 Then
        If objReport.Items("_" & Index).���� = 10 Then Exit Sub
        lbl(Index).AutoSize = True
        lbl(Index).AutoSize = False
        objReport.Items("_" & Index).W = Format(lbl(Index).Width / sgnMode, "0.00")
        objReport.Items("_" & Index).H = Format(lbl(Index).Height / sgnMode, "0.00")
        SeekItem lbl(Index), lbl(Index).Left, lbl(Index).Top
        Call ShowAttrib(Index)
        Call ReferTo
        BlnSave = False
    End If
End Sub

Private Sub lbl_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lngPreX = X: lngPreY = Y
    Set mobjMove = Nothing: mlngX = 0: mlngY = 0
    If Button = 1 Then
        If lbl(Index).MousePointer <> 99 Then 'zyb#Modify
            'zyb#Add
            picPaper_MouseDown Button, Shift, X + lbl(Index).Left, Y + lbl(Index).Top
        Else 'zyb#Modify
            If Shift = 2 Then
                If Mid(lbl(Index).Tag, 1, 2) = "" Then
                    Call SelItem(Index, True) '��ѡ
                    If GetSelNum() = 1 Then
                        Call ShowAttrib(Index) 'ֻѡ��һ������ʾ����
                    Else
                        Call ShowAttrib '��ѡʱ����ʾ����
                    End If
                Else
                    Call SelItem(Index, False) '��ѡ
                    If GetSelNum() = 1 Then
                        Call ShowAttrib(intCurID) 'ֻѡ��һ������ʾ����(ѡ�еĲ�һ���Ǹÿؼ�)
                    Else
                        Call ShowAttrib '��ѡʱ����ʾ����
                    End If
                End If
            Else
                If Mid(lbl(Index).Tag, 1, 2) = "" Then
                    Call SelClear
                    Call SelItem(Index, True)
                    Call ShowAttrib(Index) 'ֻѡ��һ������ʾ����
                End If
            End If
        End If 'zyb#Modify
    ElseIf Button = 2 Then
        PopupMenu mnuFormat, 2
    End If
End Sub

Private Sub lbl_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ObjSel As Label

    If objReport.Items("_" & Index).ϵͳ Then Exit Sub
    DrawXY X + lbl(Index).Left, Y + lbl(Index).Top
    
    If Button = 1 And Mid(lbl(Index).Tag, 1, 2) <> "" Then
        If blnLock Then Exit Sub
        Call MoveSelect(X - lngPreX, Y - lngPreY)
        If objReport.Items("_" & Index).���� = 10 Then Call DrawFrame(lbl(Index))
        If GetSelNum() = 1 Then ShowAttrib Index
    Else
        'zyb#Add
        '���������ϻ�����Ԫ��
        If objReport.Items("_" & Index).���� = 10 Then
            Set ObjSel = lbl(Index)
            
            If X < 100 Or Y < 100 Or X > ObjSel.Width - 100 Or Y > ObjSel.Height - 100 Then
                ObjSel.MousePointer = 99
            Else
                ObjSel.MousePointer = IIF(bytCurTool <> 0, 2, 0)
                ObjSel.ZOrder 1
                picPaper_MouseMove Button, Shift, X + lbl(Index).Left, Y + lbl(Index).Top
            End If
        End If
    End If
End Sub

Private Sub lbl_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lbl(Index).MousePointer <> 99 Then
        picPaper_MouseUp Button, Shift, X + lbl(Index).Left, Y + lbl(Index).Top
    End If
    picPaper.Cls
    If Not mobjMove Is Nothing Then
        If Not mobjMove Is lbl(Index).Container Then
            If objReport.Items("_" & Index).���� = "" Then
                If GetDataSouse(objReport.Items("_" & Index).����) <> "" And UCase(mobjMove.name) = "PIC" Then
                    If objReport.Items("_" & mobjMove.Index).����Դ = "" Then
                        If MsgBox("��ǰ��Ƭδ������Դ���󶨺󽫷����ӡ���ſ�Ƭ������Դ�д���""�����ʶ""�ֶ���""�����ʶ""��ͬ��Ϊһ��,����һ������Ϊһ�飻" & vbCrLf & _
                             "������ֻ��ӡһ�ſ�Ƭ���Ƿ������Դ""" & GetDataSouse(objReport.Items("_" & Index).����) & """?", vbQuestion + vbYesNo + vbDefaultButton1, App.Title) = vbYes Then
                            objReport.Items("_" & mobjMove.Index).����Դ = GetDataSouse(objReport.Items("_" & Index).����)
                        End If
                    End If
                End If
                Set lbl(Index).Container = mobjMove
                lbl(Index).Top = mlngY: lbl(Index).Left = mlngX
                If UCase(mobjMove.name) = "PIC" Then
                    objReport.Items("_" & Index).��ID = mobjMove.Index
                Else
                    objReport.Items("_" & Index).��ID = 0
                End If
                objReport.Items("_" & Index).X = mlngX: objReport.Items("_" & Index).Y = mlngY
                Set mobjMove = Nothing: mlngX = 0: mlngY = 0
                Call ShowAttrib(Index)
            End If
        End If
    End If
End Sub

Private Sub lblLine_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lngPreX = X: lngPreY = Y
    Set mobjMove = Nothing: mlngX = 0: mlngY = 0
    If Button = 1 Then
        If Shift = 2 Then
            If Mid(lblLine(Index).Tag, 1, 2) = "" Then
                Call SelItem(Index, True) '��ѡ
                If GetSelNum() = 1 Then
                    Call ShowAttrib(Index) 'ֻѡ��һ������ʾ����
                Else
                    Call ShowAttrib '��ѡʱ����ʾ����
                End If
            Else
                Call SelItem(Index, False) '��ѡ
                If GetSelNum() = 1 Then
                    Call ShowAttrib(intCurID) 'ֻѡ��һ������ʾ����(ѡ�еĲ�һ���Ǹÿؼ�)
                Else
                    Call ShowAttrib '��ѡʱ����ʾ����
                End If
            End If
        Else
            If Mid(lblLine(Index).Tag, 1, 2) = "" Then
                Call SelClear
                Call SelItem(Index, True)
                Call ShowAttrib(Index) 'ֻѡ��һ������ʾ����
            End If
        End If
    ElseIf Button = 2 Then
        PopupMenu mnuFormat, 2
    End If
End Sub

Private Sub lblLine_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If objReport.Items("_" & Index).ϵͳ Then Exit Sub
    DrawXY X + lblLine(Index).Left, Y + lblLine(Index).Top
    If Button = 1 And Mid(lblLine(Index).Tag, 1, 2) <> "" Then
        If blnLock Then Exit Sub
        Call MoveSelect(X - lngPreX, Y - lngPreY)
        If GetSelNum() = 1 Then ShowAttrib Index
    End If
End Sub

Private Sub lblPar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lngPreY = Y
End Sub

Private Sub lblPar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ClearXY
    If Button = 1 Then
        Call NoneEdit
        If tvwSQL.Height + Y - lngPreY < 2000 Or lvwPar.Height - (Y - lngPreY) < 600 Then Exit Sub
        lblPar.Top = lblPar.Top + Y - lngPreY
        tvwSQL.Height = tvwSQL.Height + Y - lngPreY
        lvwPar.Top = lvwPar.Top + Y - lngPreY
        lvwPar.Height = lvwPar.Height - (Y - lngPreY)
        Refresh
    End If
End Sub

Private Sub lblSize_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ObjSel As Control, tmpID As RelatID
    Dim lngMinW As Long, lngMinH As Long, i As Integer
    Dim xx As Integer, yy As Integer, zz As Integer
    Dim lngTop As Long, lngLeft As Long
   
    DrawXY X + lblSize(Index).Left, Y + lblSize(Index).Top
    
    If Button = 1 And GetSelNum = 1 And Not blnLock Then
        If objReport.Items("_" & intCurID).ϵͳ Then Exit Sub
        Select Case objReport.Items("_" & intCurID).����
            Case 1
                Set ObjSel = lblLine(intCurID)
                lngMinW = lblSize(0).Width: lngMinH = lblSize(0).Height
            Case 2, 3
                Set ObjSel = lbl(intCurID)
                lngMinW = lblSize(0).Width: lngMinH = lblSize(0).Height
            Case 10
                Set ObjSel = Shp(intCurID)
                lngMinW = lblSize(0).Width: lngMinH = lblSize(0).Height
            Case 4 '������
                Call ResetColor(intCurID)
                Set ObjSel = msh(intCurID)
                lngMinW = msh(intCurID).ColWidth(0) + 15
                
                lngMinH = 0
                For xx = 0 To msh(intCurID).FixedRows
                    If msh(intCurID).Rows - 1 >= msh(intCurID).FixedRows + 1 Then
                        lngMinH = lngMinH + msh(intCurID).RowHeight(xx)
                    Else
                        lngMinH = lngMinH + 255 * sgnMode
                    End If
                Next
                lngMinH = lngMinH + 15
                Call CustomColColor(intCurID, 0)
            Case 5 '���ܱ��
                Call ResetColor(intCurID)
                Set ObjSel = msh(intCurID)
                xx = msh(intCurID).FixedCols '���������Ŀ��
                yy = msh(intCurID).FixedRows - 1 '���������Ŀ��
                For Each tmpID In objReport.Items("_" & intCurID).SubIDs
                    If objReport.Items("_" & tmpID.id).���� = 9 Then zz = zz + 1 'ͳ����Ŀ��
                Next
                lngMinH = msh(intCurID).RowHeight(0) * (yy + 1) + 60 '��С��ͷ+1�и߶�
                For i = 0 To xx + zz - 1
                    lngMinW = lngMinW + msh(intCurID).ColWidth(i)
                Next
                lngMinW = lngMinW + 60
            Case 11
                Set ObjSel = img(intCurID)
                lngMinW = lblSize(0).Width: lngMinH = lblSize(0).Height
            Case 12 '@@@
                Set ObjSel = Chart(intCurID)
                lngMinW = Chart(0).Width: lngMinH = Chart(0).Height
            Case 13
                Set ObjSel = ImgCode(intCurID)
                lngMinW = lblSize(0).Width: lngMinH = lblSize(0).Height
            Case 14
                Set ObjSel = pic(intCurID)
                lngMinW = lblSize(0).Width: lngMinH = lblSize(0).Height
        End Select
        
        '@@@
        lngMinW = lngMinW * sgnMode
        lngMinH = lngMinH * sgnMode
        If UCase(ObjSel.Container.name) = "PIC" Then
            lngTop = ObjSel.Container.Top
            lngLeft = ObjSel.Container.Left
        End If
        '�ı�ؼ��ߴ��С
        Select Case Index Mod 8
            Case 1 '����
                If ObjSel.Height - Y < lngMinH Then Exit Sub
                If objReport.Items("_" & intCurID).���� = 12 Then '@@@
                    If ObjSel.Top + lngTop + Y < 0 Then Exit Sub
                End If
                
                lblSize(Index).Top = lblSize(Index).Top + Y
                lblSize(Index + 1).Top = lblSize(Index + 1).Top + Y
                lblSize(Index + 7).Top = lblSize(Index + 7).Top + Y
                
                ObjSel.Top = ObjSel.Top + Y
                ObjSel.Height = ObjSel.Height - Y
                If objReport.Items("_" & intCurID).���� = 10 Then
                    lblshp(intCurID).Top = lblshp(intCurID).Top + Y
                    lblshp(intCurID).Height = lblshp(intCurID).Height - Y
                End If
                
                lblSize(Index + 2).Top = ObjSel.Top + lngTop + (ObjSel.Height - lblSize(Index + 2).Height) / 2
                lblSize(Index + 6).Top = lblSize(Index + 2).Top
            Case 2 '����
                If ObjSel.Height - Y < lngMinH Or ObjSel.Width + X < lngMinW Then Exit Sub
                If objReport.Items("_" & intCurID).���� = 12 Then '@@@
                    If ObjSel.Top + lngTop + Y < 0 Then Exit Sub
                End If
                
                lblSize(Index - 1).Top = lblSize(Index - 1).Top + Y
                lblSize(Index).Top = lblSize(Index).Top + Y
                lblSize(Index + 6).Top = lblSize(Index + 6).Top + Y

                lblSize(Index).Left = lblSize(Index).Left + X
                lblSize(Index + 1).Left = lblSize(Index + 1).Left + X
                lblSize(Index + 2).Left = lblSize(Index + 2).Left + X
                
                ObjSel.Top = ObjSel.Top + Y
                ObjSel.Height = ObjSel.Height - Y
                ObjSel.Width = ObjSel.Width + X
                If objReport.Items("_" & intCurID).���� = 10 Then
                    lblshp(intCurID).Top = lblshp(intCurID).Top + Y
                    lblshp(intCurID).Height = lblshp(intCurID).Height - Y
                    lblshp(intCurID).Width = lblshp(intCurID).Width + X
                End If
                
                lblSize(Index + 1).Top = ObjSel.Top + lngTop + (ObjSel.Height - lblSize(Index + 1).Height) / 2
                lblSize(Index + 5).Top = lblSize(Index + 1).Top
                lblSize(Index - 1).Left = ObjSel.Left + lngLeft + (ObjSel.Width - lblSize(Index - 1).Width) / 2
                lblSize(Index + 3).Left = lblSize(Index - 1).Left
            Case 3 '����
                If ObjSel.Width + X < lngMinW Then Exit Sub
                
                lblSize(Index - 1).Left = lblSize(Index - 1).Left + X
                lblSize(Index).Left = lblSize(Index).Left + X
                lblSize(Index + 1).Left = lblSize(Index + 1).Left + X
                
                ObjSel.Width = ObjSel.Width + X
                If objReport.Items("_" & intCurID).���� = 10 Then
                    lblshp(intCurID).Width = lblshp(intCurID).Width + X
                End If
                
                lblSize(Index - 2).Left = ObjSel.Left + lngLeft + (ObjSel.Width - lblSize(Index - 2).Width) / 2
                lblSize(Index + 2).Left = lblSize(Index - 2).Left
            Case 4 '����
                If ObjSel.Height + Y < lngMinH Or ObjSel.Width + X < lngMinW Then Exit Sub
                
                lblSize(Index).Left = lblSize(Index).Left + X
                lblSize(Index - 1).Left = lblSize(Index - 1).Left + X
                lblSize(Index - 2).Left = lblSize(Index - 2).Left + X
                lblSize(Index).Top = lblSize(Index).Top + Y
                lblSize(Index + 1).Top = lblSize(Index + 1).Top + Y
                lblSize(Index + 2).Top = lblSize(Index + 2).Top + Y
                
                ObjSel.Height = ObjSel.Height + Y
                ObjSel.Width = ObjSel.Width + X
                If objReport.Items("_" & intCurID).���� = 10 Then
                    lblshp(intCurID).Height = lblshp(intCurID).Height + Y
                    lblshp(intCurID).Width = lblshp(intCurID).Width + X
                End If
                
                lblSize(Index - 1).Top = ObjSel.Top + lngTop + (ObjSel.Height - lblSize(Index - 1).Height) / 2
                lblSize(Index + 3).Top = lblSize(Index - 1).Top
                lblSize(Index - 3).Left = ObjSel.Left + lngLeft + (ObjSel.Width - lblSize(Index - 3).Width) / 2
                lblSize(Index + 1).Left = lblSize(Index - 3).Left
            Case 5 '����
                If ObjSel.Height + Y < lngMinH Then Exit Sub
                
                lblSize(Index).Top = lblSize(Index).Top + Y
                lblSize(Index - 1).Top = lblSize(Index - 1).Top + Y
                lblSize(Index + 1).Top = lblSize(Index + 1).Top + Y
                
                ObjSel.Height = ObjSel.Height + Y
                If objReport.Items("_" & intCurID).���� = 10 Then
                    lblshp(intCurID).Height = lblshp(intCurID).Height + Y
                End If
                
                lblSize(Index - 2).Top = ObjSel.Top + lngTop + (ObjSel.Height - lblSize(Index - 2).Height) / 2
                lblSize(Index + 2).Top = lblSize(Index - 2).Top
            Case 6 '����
                If ObjSel.Height + Y < lngMinH Or ObjSel.Width - X < lngMinW Then Exit Sub
                If objReport.Items("_" & intCurID).���� = 12 Then '@@@
                    If ObjSel.Left + lngLeft + X < 0 Then Exit Sub
                End If
                
                lblSize(Index).Top = lblSize(Index).Top + Y
                lblSize(Index - 1).Top = lblSize(Index - 1).Top + Y
                lblSize(Index - 2).Top = lblSize(Index - 2).Top + Y
                
                lblSize(Index).Left = lblSize(Index).Left + X
                lblSize(Index + 1).Left = lblSize(Index + 1).Left + X
                lblSize(Index + 2).Left = lblSize(Index + 2).Left + X
                
                ObjSel.Width = ObjSel.Width - X
                ObjSel.Height = ObjSel.Height + Y
                ObjSel.Left = ObjSel.Left + X
                If objReport.Items("_" & intCurID).���� = 10 Then
                    lblshp(intCurID).Left = lblshp(intCurID).Left + X
                    lblshp(intCurID).Height = lblshp(intCurID).Height + Y
                    lblshp(intCurID).Width = lblshp(intCurID).Width - X
                End If
                
                lblSize(Index - 1).Left = ObjSel.Left + lngLeft + (ObjSel.Width - lblSize(Index - 1).Width) / 2
                lblSize(Index - 5).Left = lblSize(Index - 1).Left
                lblSize(Index + 1).Top = ObjSel.Top + lngTop + (ObjSel.Height - lblSize(Index + 1).Height) / 2
                lblSize(Index - 3).Top = lblSize(Index + 1).Top
            Case 7 '����
                If ObjSel.Width - X < lngMinW Then Exit Sub
                If objReport.Items("_" & intCurID).���� = 12 Then '@@@
                    If ObjSel.Left + lngLeft + X < 0 Then Exit Sub
                End If
                
                lblSize(Index).Left = lblSize(Index).Left + X
                lblSize(Index + 1).Left = lblSize(Index + 1).Left + X
                lblSize(Index - 1).Left = lblSize(Index - 1).Left + X
                
                ObjSel.Width = ObjSel.Width - X
                ObjSel.Left = ObjSel.Left + X
                If objReport.Items("_" & intCurID).���� = 10 Then
                    lblshp(intCurID).Left = lblshp(intCurID).Left + X
                    lblshp(intCurID).Width = lblshp(intCurID).Width - X
                End If
                
                lblSize(Index - 6).Left = ObjSel.Left + lngLeft + (ObjSel.Width - lblSize(Index - 6).Width) / 2
                lblSize(Index - 2).Left = lblSize(Index - 6).Left
            Case 0 '����
                If ObjSel.Height - Y < lngMinH Or ObjSel.Width - X < lngMinW Then Exit Sub
                If objReport.Items("_" & intCurID).���� = 12 Then '@@@
                    If ObjSel.Top + lngTop + Y < 0 Then Exit Sub
                    If ObjSel.Left + lngLeft + X < 0 Then Exit Sub
                End If
                
                lblSize(Index).Left = lblSize(Index).Left + X
                lblSize(Index - 1).Left = lblSize(Index - 1).Left + X
                lblSize(Index - 2).Left = lblSize(Index - 2).Left + X
                
                lblSize(Index).Top = lblSize(Index).Top + Y
                lblSize(Index - 6).Top = lblSize(Index - 6).Top + Y
                lblSize(Index - 7).Top = lblSize(Index - 7).Top + Y
                
                ObjSel.Width = ObjSel.Width - X
                ObjSel.Height = ObjSel.Height - Y
                ObjSel.Left = ObjSel.Left + X
                ObjSel.Top = ObjSel.Top + Y
                If objReport.Items("_" & intCurID).���� = 10 Then
                    lblshp(intCurID).Left = lblshp(intCurID).Left + X
                    lblshp(intCurID).Height = lblshp(intCurID).Height - Y
                    lblshp(intCurID).Width = lblshp(intCurID).Width - X
                    lblshp(intCurID).Top = lblshp(intCurID).Top + Y
                End If
                
                lblSize(Index - 1).Top = ObjSel.Top + lngTop + (ObjSel.Height - lblSize(Index - 1).Height) / 2
                lblSize(Index - 5).Top = lblSize(Index - 1).Top
                lblSize(Index - 7).Left = ObjSel.Left + lngLeft + (ObjSel.Width - lblSize(Index - 7).Width) / 2
                lblSize(Index - 3).Left = lblSize(Index - 7).Left
        End Select
        If objReport.Items("_" & intCurID).���� <> 10 Then Me.Refresh
        
        'ͼƬ�ߴ����ʱ���ֱ���
        If objReport.Items("_" & intCurID).���� = 11 Then
            If Not objReport.Items("_" & intCurID).ͼƬ Is Nothing Then
                If objReport.Items("_" & intCurID).���� Then
                    Set ObjSel.Picture = ScalePicture(PicFontTest, objReport.Items("_" & intCurID).ͼƬ, ObjSel.Width, ObjSel.Height)
                End If
            End If
        End If
        
        Dim MainItem As RPTItem
        If InStr(1, "4,5", objReport.Items("_" & intCurID).����) > 0 And objReport.Items("_" & intCurID).���� <> "" Then
            For Each MainItem In objReport.Items
                If MainItem.��ʽ�� = mbytCurrFmt And MainItem.���� = objReport.Items("_" & intCurID).���� Then Exit For
            Next
        End If
        If InStr(1, "4,5", objReport.Items("_" & intCurID).����) > 0 Then
            Call SetGridLine(intCurID)  '�������
            If objReport.Items("_" & intCurID).���� = 4 Then
                Call SetCopyGrid(intCurID) '��������ؼ�
            End If
            If Not MainItem Is Nothing Then
                If objReport.Items("_" & intCurID).���� <> 1 Then
                    msh(MainItem.Key).Height = msh(intCurID).Height
                    objReport.Items("_" & MainItem.Key).H = msh(MainItem.Key).Height / sgnMode
                Else
                    msh(MainItem.Key).Width = msh(intCurID).Width
                    objReport.Items("_" & MainItem.Key).W = msh(MainItem.Key).Width / sgnMode
                End If
                Call SetMainWH(MainItem.Key)    '����ӱ�(�����ӻ򸽼�)
            Else
                Call SetChildWH(intCurID)    '����ӱ�(�����ӻ򸽼�)
            End If
        End If
        
        '�������ݶ���
        objReport.Items("_" & intCurID).X = Format(ObjSel.Left / sgnMode, "0.00")
        objReport.Items("_" & intCurID).Y = Format(ObjSel.Top / sgnMode, "0.00")
        If objReport.Items("_" & intCurID).���� = 1 Then
            If ObjSel.Width > ObjSel.Height Then
                objReport.Items("_" & intCurID).W = Format(ObjSel.Width / sgnMode, "0.00")
                objReport.Items("_" & intCurID).H = 0
            Else
                objReport.Items("_" & intCurID).W = 0
                objReport.Items("_" & intCurID).H = Format(ObjSel.Height / sgnMode, "0.00")
            End If
        Else
            objReport.Items("_" & intCurID).W = Format(ObjSel.Width / sgnMode, "0.00")
            objReport.Items("_" & intCurID).H = Format(ObjSel.Height / sgnMode, "0.00")
        End If
        
        If GetSelNum = 1 And InStr(1, "4,5", objReport.Items("_" & intCurID).����) <> 0 Then
            Call AdjustSelCons(msh(intCurID))
        End If
        Call MoveSelect(0, 0, True)
        Call ShowAttrib(intCurID)
        Call AdjustAll(True)
        BlnSave = False
    End If
End Sub

Private Sub lvwPar_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Call ClearXY
End Sub

Private Sub lvwPar_GotFocus()
    Call NoneEdit
End Sub

Private Sub mnuClass_Align_Style_Click(Index As Integer)
'���ܣ��Ի��ܱ��,���õ�ǰ������뷽ʽ
    Dim X As Integer, Y As Integer, Z As Integer, intDel As Integer
    Dim tmpID As RelatID, tmpItem As RPTItem
    
    If selCell.Col1 = -1 And selCell.Row1 = -1 Then sta.Panels(2).Text = "����ȷ����ǰ��Ŀ���ͼ�λ�ã�": PlayWarn: Exit Sub
    
    'ͳ���Χ
    For Each tmpID In objReport.Items("_" & intCurID).SubIDs
        Set tmpItem = objReport.Items("_" & tmpID.id)
        If tmpItem.���� = 9 And tmpItem.��� = selCell.Col1 - msh(intCurID).FixedCols Then
            tmpItem.���� = Index: Exit For
        End If
    Next
    BlnSave = False
End Sub

Private Sub mnuClass_Data_Click()
'���ܣ��Ի��ܱ��,���õ�ǰ��Ŀ��������Դ
    Dim intState As Integer
    Dim tmpID As RelatID, tmpItem As RPTItem
    Dim objRelations As New RPTRelations
    Dim objColProtertys As New RPTColProtertys
    Dim i As Long
    
    If selCell.Col1 = -1 Or selCell.Row1 = -1 Then
        sta.Panels(2).Text = "����ȷ����ǰ��Ŀ���ͼ�λ�ã�"
        Call PlayWarn
        Exit Sub
    End If
    
    '����ܱ��ͳ�������
    For Each tmpID In objReport.Items("_" & intCurID).SubIDs
        If objReport.Items("_" & tmpID.id).���� = 9 Then intState = intState + 1
    Next
    
    If selCell.Col1 <= msh(intCurID).FixedCols - 1 And selCell.Row1 >= msh(intCurID).FixedRows - 1 Then
        '������෶Χ
        frmData.I_strTitle = "�������������"
        frmData.I_bytType = 0
        frmData.I_strClass = objReport.Items("_" & intCurID).����
        frmData.mintEleID = intCurID
        For Each tmpID In objReport.Items("_" & intCurID).SubIDs
            Set tmpItem = objReport.Items("_" & tmpID.id)
            If tmpItem.��� = selCell.Col1 And tmpItem.���� = 7 Then
                frmData.IO_FontBold = IIF(tmpItem.����, 1, 0)
                frmData.IO_FontColor = tmpItem.ǰ��
                frmData.I_strOrder = tmpItem.����: Exit For
            End If
        Next
        Set frmData.objReport = objReport
        Call SetCopyRelations(tmpItem.Relations, objRelations)
        Set frmData.mobjRelations = objRelations
        Set frmData.frmParent = Me
        frmData.IO_strNode = GetDataName(msh(intCurID).TextMatrix(msh(intCurID).FixedRows - 1, selCell.Col1))
        frmData.Show 1, Me
        If gblnOK Then
            '������Ŀ�����ظ�
            For Each tmpID In objReport.Items("_" & intCurID).SubIDs
                Set tmpItem = objReport.Items("_" & tmpID.id)
                If tmpItem.���� = frmData.txtData.Text And (tmpItem.��� <> selCell.Col1 Or tmpItem.���� <> 7) Then
                    sta.Panels(2).Text = "����ѡ����������ڸû��ܱ�����Ѿ����ڣ�"
                    Unload frmData: PlayWarn: Exit Sub
                End If
            Next
            For Each tmpID In objReport.Items("_" & intCurID).SubIDs
                Set tmpItem = objReport.Items("_" & tmpID.id)
                If tmpItem.��� = selCell.Col1 And tmpItem.���� = 7 Then
                    tmpItem.���� = frmData.txtData.Text
                    If frmData.txtOrder.Text = "" Then
                        tmpItem.���� = ""
                    Else
                        tmpItem.���� = frmData.txtOrder.Text
                        If frmData.optDesc Then tmpItem.���� = "," & tmpItem.����
                    End If
                    tmpItem.���� = frmData.IO_FontBold
                    tmpItem.ǰ�� = frmData.IO_FontColor
                    Exit For
                End If
            Next
            '����������
            Set tmpItem.Relations = objRelations
            Unload frmData
            Call ReShowGrid(intCurID)
            Call ClassColor(intCurID, selCell)
            BlnSave = False
        End If
    ElseIf selCell.Row1 <= msh(intCurID).FixedRows - 2 Then
        '������෶Χ
        frmData.I_strTitle = "�������������"
        frmData.I_bytType = 1
        frmData.I_strClass = objReport.Items("_" & intCurID).����
        frmData.mintEleID = intCurID
        For Each tmpID In objReport.Items("_" & intCurID).SubIDs
            Set tmpItem = objReport.Items("_" & tmpID.id)
            If tmpItem.��� = selCell.Row1 And tmpItem.���� = 8 Then
                frmData.IO_FontBold = IIF(tmpItem.����, 1, 0)
                frmData.IO_FontColor = tmpItem.ǰ��
                frmData.I_strOrder = tmpItem.����: Exit For
            End If
        Next
        Set frmData.frmParent = Me
        Set frmData.objReport = objReport
        Call SetCopyRelations(tmpItem.Relations, objRelations)
        Set frmData.mobjRelations = objRelations
        frmData.IO_strNode = GetDataName(msh(intCurID).TextMatrix(selCell.Row1, msh(intCurID).FixedCols - 1))
        frmData.Show 1, Me
        If gblnOK Then
            '������Ŀ�����ظ�
            For Each tmpID In objReport.Items("_" & intCurID).SubIDs
                Set tmpItem = objReport.Items("_" & tmpID.id)
                If tmpItem.���� = frmData.txtData.Text And (tmpItem.��� <> selCell.Row1 Or tmpItem.���� <> 8) Then
                    sta.Panels(2).Text = "����ѡ����������ڸû��ܱ�����Ѿ����ڣ�"
                    Unload frmData: PlayWarn: Exit Sub
                End If
            Next
            For Each tmpID In objReport.Items("_" & intCurID).SubIDs
                Set tmpItem = objReport.Items("_" & tmpID.id)
                If tmpItem.��� = selCell.Row1 And tmpItem.���� = 8 Then
                    tmpItem.���� = frmData.txtData.Text
                    tmpItem.���� = frmData.IO_FontBold
                    tmpItem.ǰ�� = frmData.IO_FontColor
                    If frmData.txtOrder.Text = "" Then
                        tmpItem.���� = ""
                    Else
                        tmpItem.���� = frmData.txtOrder.Text
                        If frmData.optDesc Then tmpItem.���� = "," & tmpItem.����
                    End If
                    Exit For
                End If
            Next
            '����������
            Set tmpItem.Relations = objRelations
            Unload frmData
            Call ReShowGrid(intCurID)
            Call ClassColor(intCurID, selCell)
            BlnSave = False
        End If
    ElseIf selCell.Col1 >= msh(intCurID).FixedCols And selCell.Col1 <= msh(intCurID).FixedCols + intState - 1 And selCell.Row1 >= msh(intCurID).FixedRows - 1 Then
        'ͳ���Χ
        frmData.I_strTitle = "ͳ��������"
        frmData.I_strOrder = ""
        frmData.I_bytType = 2
        frmData.I_strClass = objReport.Items("_" & intCurID).����
        frmData.mintEleID = intCurID
        For Each tmpID In objReport.Items("_" & intCurID).SubIDs
            Set tmpItem = objReport.Items("_" & tmpID.id)
            If tmpItem.��� = selCell.Col1 - msh(intCurID).FixedCols And tmpItem.���� = 9 Then
                frmData.IO_FontBold = IIF(tmpItem.����, 1, 0)
                frmData.IO_FontColor = tmpItem.ǰ��
                frmData.I_strFormat = tmpItem.��ʽ: Exit For
            End If
        Next
        Set frmData.frmParent = Me
        Set frmData.objReport = objReport
        Call SetCopyRelations(tmpItem.Relations, objRelations)
        Set frmData.mobjRelations = objRelations
        Call SetCopyColProtertys(tmpItem.ColProtertys, objColProtertys)
        Set frmData.mobjColProtertys = objColProtertys
        frmData.IO_strNode = GetDataName(msh(intCurID).TextMatrix(msh(intCurID).FixedRows - 1, selCell.Col1))
        frmData.I_strSummaryFile = GetSummaryFile(intCurID)
        frmData.Show 1, Me
        If gblnOK Then
            '������Ŀ�����ظ�
            For Each tmpID In objReport.Items("_" & intCurID).SubIDs
                Set tmpItem = objReport.Items("_" & tmpID.id)
                If tmpItem.���� = frmData.txtData.Text And (tmpItem.��� <> selCell.Col1 - msh(intCurID).FixedCols Or tmpItem.���� <> 9) Then
                    sta.Panels(2).Text = "����ѡ����������ڸû��ܱ�����Ѿ����ڣ�"
                    Unload frmData: PlayWarn: Exit Sub
                End If
            Next
            For Each tmpID In objReport.Items("_" & intCurID).SubIDs
                Set tmpItem = objReport.Items("_" & tmpID.id)
                If tmpItem.��� = selCell.Col1 - msh(intCurID).FixedCols And tmpItem.���� = 9 Then
                    tmpItem.���� = frmData.txtData.Text
                    tmpItem.��ʽ = frmData.txtFormat.Text
                    tmpItem.���� = frmData.IO_FontBold
                    tmpItem.ǰ�� = frmData.IO_FontColor
                    Exit For
                End If
            Next
            '����������
            Set tmpItem.Relations = objRelations
            '����������
            Set tmpItem.ColProtertys = frmData.mobjColProtertys
            
            Unload frmData
            Call ReShowGrid(intCurID)
            Call ClassColor(intCurID, selCell, intState)
            BlnSave = False
        End If
    End If
End Sub

Private Function GetSummaryFile(ByVal intIndex As Integer)
'���ܣ���ȡ���ܱ��������ֶ�
    Dim i As Long, strReturn As String
    Dim strTmp As String
    
    For i = msh(intCurID).FixedCols To msh(intCurID).Cols - 1
        strTmp = GetDataName(msh(intIndex).TextMatrix(msh(intIndex).FixedRows - 1, i))
        If InStr(strReturn, strTmp) = 0 Then strReturn = strReturn & "," & strTmp
    Next
    GetSummaryFile = Mid(strReturn, 2)
End Function

Private Sub mnuClass_Del_Click()
'���ܣ��Ի��ܱ��,ɾ����ǰ��Ŀ
    Dim X As Integer, Y As Integer, Z As Integer, intDel As Integer
    Dim tmpID As RelatID, tmpItem As RPTItem
    
    If selCell.Col1 = -1 And selCell.Row1 = -1 Then sta.Panels(2).Text = "����ȷ����ǰ��Ŀ���ͼ�λ�ã�": PlayWarn: Exit Sub
    
    '����ܱ���������
    For Each tmpID In objReport.Items("_" & intCurID).SubIDs
        Select Case objReport.Items("_" & tmpID.id).����
            Case 7
                X = X + 1
            Case 8
                Y = Y + 1
            Case 9
                Z = Z + 1
        End Select
    Next
    If selCell.Col1 <= msh(intCurID).FixedCols - 1 And selCell.Row1 >= msh(intCurID).FixedRows - 1 Then
        '������෶Χ
        If X = 1 Then sta.Panels(2).Text = "����Ҫ��һ�����������Ŀ��": PlayWarn: Exit Sub
        For Each tmpID In objReport.Items("_" & intCurID).SubIDs
            Set tmpItem = objReport.Items("_" & tmpID.id)
            If tmpItem.���� = 7 Then
                If tmpItem.��� = selCell.Col1 Then
                    intDel = tmpItem.id
                ElseIf tmpItem.��� > selCell.Col1 Then
                    tmpItem.��� = tmpItem.��� - 1
                End If
            End If
        Next
    ElseIf selCell.Row1 <= msh(intCurID).FixedRows - 2 Then
        '������෶Χ
        If Y = 0 Then sta.Panels(2).Text = "�Ѿ�û�к��������Ŀ��": PlayWarn: Exit Sub
        For Each tmpID In objReport.Items("_" & intCurID).SubIDs
            Set tmpItem = objReport.Items("_" & tmpID.id)
            If tmpItem.���� = 8 Then
                If tmpItem.��� = selCell.Row1 Then
                    intDel = tmpItem.id
                ElseIf tmpItem.��� > selCell.Row1 Then
                    tmpItem.��� = tmpItem.��� - 1
                End If
            End If
        Next
    ElseIf selCell.Col1 >= msh(intCurID).FixedCols And selCell.Col1 <= msh(intCurID).FixedCols + Z - 1 And selCell.Row1 >= msh(intCurID).FixedRows - 1 Then
        'ͳ���Χ
        If Z = 1 Then sta.Panels(2).Text = "����Ҫ��һ��ͳ����Ŀ��": PlayWarn: Exit Sub
        For Each tmpID In objReport.Items("_" & intCurID).SubIDs
            Set tmpItem = objReport.Items("_" & tmpID.id)
            If tmpItem.���� = 9 Then
                If tmpItem.��� = selCell.Col1 - msh(intCurID).FixedCols Then
                    intDel = tmpItem.id
                ElseIf tmpItem.��� > selCell.Col1 - msh(intCurID).FixedCols Then
                    tmpItem.��� = tmpItem.��� - 1
                End If
            End If
        Next
    End If
    objReport.Items.Remove "_" & intDel
    objReport.Items("_" & intCurID).SubIDs.Remove "_" & intDel
    selCell.Col1 = -1: selCell.Row1 = -1
    Call ReShowGrid(intCurID)
    BlnSave = False
End Sub

Private Sub mnuClass_ExChange_Click()
'���ܣ���һ�����ܱ�����жԻ�
    Dim tmpID As RelatID, i As Integer
    
    For Each tmpID In objReport.Items("_" & intCurID).SubIDs
        If objReport.Items("_" & tmpID.id).���� = 8 Then i = i + 1
    Next
    If i = 0 Then sta.Panels(2).Text = "����û�к��������Ŀ,�����л���": PlayWarn: Exit Sub

    For Each tmpID In objReport.Items("_" & intCurID).SubIDs
        If objReport.Items("_" & tmpID.id).���� = 7 Then
            objReport.Items("_" & tmpID.id).���� = 77
        ElseIf objReport.Items("_" & tmpID.id).���� = 8 Then
            objReport.Items("_" & tmpID.id).���� = 88
        End If
    Next
    For Each tmpID In objReport.Items("_" & intCurID).SubIDs
        If objReport.Items("_" & tmpID.id).���� = 77 Then
            objReport.Items("_" & tmpID.id).���� = 8
        ElseIf objReport.Items("_" & tmpID.id).���� = 88 Then
            objReport.Items("_" & tmpID.id).���� = 7
            If objReport.Items("_" & tmpID.id).W = 0 Then objReport.Items("_" & tmpID.id).W = 1000
        End If
    Next
    selCell.Col1 = -1: selCell.Row1 = -1
    Call ReShowGrid(intCurID)
    BlnSave = False
End Sub

Private Sub mnuClass_Insert_After_Click()
'���ܣ��Ի��ܱ��,�Բ����������Ŀ�ں�
    Dim intState As Integer
    Dim tmpID As RelatID, tmpItem As RPTItem
    Dim objRelations As New RPTRelations
    Dim objColProtertys As New RPTColProtertys
    Dim i As Long
    
    If selCell.Col1 = -1 And selCell.Row1 = -1 Then sta.Panels(2).Text = "����ȷ����ǰ��Ŀ���ͼ�λ�ã�": PlayWarn: Exit Sub
    
    '����ܱ��ͳ�������
    For Each tmpID In objReport.Items("_" & intCurID).SubIDs
        If objReport.Items("_" & tmpID.id).���� = 9 Then intState = intState + 1
    Next
    
    If selCell.Col1 <= msh(intCurID).FixedCols - 1 And selCell.Row1 >= msh(intCurID).FixedRows - 1 Then
        '������෶Χ
        Set frmData.frmParent = Me
        frmData.I_strTitle = "�������������"
        frmData.I_bytType = 0
        frmData.I_strClass = objReport.Items("_" & intCurID).����
        frmData.mintEleID = intCurID
        frmData.IO_strNode = ""
        Set frmData.objReport = objReport
        Set frmData.mobjRelations = objRelations
        frmData.Show 1, Me
        If gblnOK Then
            '������Ŀ�����ظ�
            For Each tmpID In objReport.Items("_" & intCurID).SubIDs
                If objReport.Items("_" & tmpID.id).���� = frmData.txtData.Text Then
                    sta.Panels(2).Text = "����ѡ����������ڸû��ܱ�����Ѿ����ڣ�"
                    Unload frmData: PlayWarn: Exit Sub
                End If
            Next
            For Each tmpID In objReport.Items("_" & intCurID).SubIDs
                Set tmpItem = objReport.Items("_" & tmpID.id)
                If tmpItem.���� = 7 And tmpItem.��� > selCell.Col1 Then
                    tmpItem.��� = tmpItem.��� + 1
                End If
            Next
            intMaxID = intMaxID + 1
            Set tmpItem = objReport.Items.Add(intMaxID, mbytCurrFmt, "Ԫ��" & intMaxID, intCurID, 7, selCell.Col1 + 1 _
                , "", 0, frmData.txtData.Text, "", 0, 0, 1000, 0, 0, 0, False, "", 0, frmData.IO_FontBold = 1, False _
                , False, 0, frmData.IO_FontColor, 0, False, 0, IIF(frmData.optDesc, ",", "") & frmData.txtOrder.Text _
                , "", "", False, False, , False, , , , "_" & intMaxID)
                
            objReport.Items("_" & intCurID).SubIDs.Add intMaxID, "_" & intMaxID
            '����������
            Set tmpItem.Relations = objRelations
            Unload frmData
            selCell.Col1 = -1: selCell.Row1 = -1
            Call ReShowGrid(intCurID)
            BlnSave = False
        End If
    ElseIf selCell.Row1 <= msh(intCurID).FixedRows - 2 Then
        '������෶Χ
        Set frmData.frmParent = Me
        frmData.I_strTitle = "�������������"
        frmData.I_bytType = 1
        frmData.I_strClass = objReport.Items("_" & intCurID).����
        frmData.mintEleID = intCurID
        frmData.IO_strNode = ""
        Set frmData.objReport = objReport
        Set frmData.mobjRelations = objRelations
        frmData.Show 1, Me
        If gblnOK Then
            '������Ŀ�����ظ�
            For Each tmpID In objReport.Items("_" & intCurID).SubIDs
                If objReport.Items("_" & tmpID.id).���� = frmData.txtData.Text Then
                    sta.Panels(2).Text = "����ѡ����������ڸû��ܱ�����Ѿ����ڣ�"
                    Unload frmData: PlayWarn: Exit Sub
                End If
            Next
            For Each tmpID In objReport.Items("_" & intCurID).SubIDs
                Set tmpItem = objReport.Items("_" & tmpID.id)
                If tmpItem.���� = 8 And tmpItem.��� > selCell.Row1 Then
                    tmpItem.��� = tmpItem.��� + 1
                End If
            Next
            intMaxID = intMaxID + 1
            Set tmpItem = objReport.Items.Add(intMaxID, mbytCurrFmt, "Ԫ��" & intMaxID, intCurID, 8, selCell.Row1 + 1 _
                , "", 0, frmData.txtData.Text, "", 0, 0, 0, 0, 0, 0, False, "", 0, frmData.IO_FontBold = 1, False, False _
                , 0, frmData.IO_FontColor, 0, False, 0, IIF(frmData.optDesc, ",", "") & frmData.txtOrder.Text, "", "" _
                , False, False, , False, , , , "_" & intMaxID)
                
            objReport.Items("_" & intCurID).SubIDs.Add intMaxID, "_" & intMaxID
            '����������
            Set tmpItem.Relations = objRelations
            Unload frmData
            selCell.Col1 = -1: selCell.Row1 = -1
            Call ReShowGrid(intCurID)
            BlnSave = False
        End If
    ElseIf selCell.Col1 >= msh(intCurID).FixedCols And selCell.Col1 <= msh(intCurID).FixedCols + intState - 1 _
        And selCell.Row1 >= msh(intCurID).FixedRows - 1 Then
        'ͳ���Χ
        Set frmData.frmParent = Me
        frmData.I_strTitle = "ͳ��������"
        frmData.I_bytType = 2
        frmData.I_strClass = objReport.Items("_" & intCurID).����
        frmData.mintEleID = intCurID
        frmData.IO_strNode = ""
        Set frmData.objReport = objReport
        Set frmData.mobjRelations = objRelations
        Set frmData.mobjColProtertys = objColProtertys
        frmData.I_strSummaryFile = GetSummaryFile(intCurID)
        frmData.Show 1, Me
        If gblnOK Then
            '������Ŀ�����ظ�
            For Each tmpID In objReport.Items("_" & intCurID).SubIDs
                If objReport.Items("_" & tmpID.id).���� = frmData.txtData.Text Then
                    sta.Panels(2).Text = "����ѡ����������ڸû��ܱ�����Ѿ����ڣ�"
                    Unload frmData: PlayWarn: Exit Sub
                End If
            Next
            For Each tmpID In objReport.Items("_" & intCurID).SubIDs
                Set tmpItem = objReport.Items("_" & tmpID.id)
                If tmpItem.���� = 9 And tmpItem.��� > selCell.Col1 - msh(intCurID).FixedCols Then
                    tmpItem.��� = tmpItem.��� + 1
                End If
            Next
            intMaxID = intMaxID + 1
            Set tmpItem = objReport.Items.Add(intMaxID, mbytCurrFmt, "Ԫ��" & intMaxID, intCurID, 9 _
                , selCell.Col1 - msh(intCurID).FixedCols + 1, "", 0, frmData.txtData.Text, "", 0, 0, 1000, 0, 0, 0 _
                , False, "", 0, frmData.IO_FontBold = 1, False, False, 0, frmData.IO_FontColor, 0, False, 0, "" _
                , frmData.txtFormat.Text, "", False, False, , False, , , , "_" & intMaxID)
                
            objReport.Items("_" & intCurID).SubIDs.Add intMaxID, "_" & intMaxID
            '����������
            Set tmpItem.Relations = objRelations
            '����������
            Set tmpItem.ColProtertys = frmData.mobjColProtertys
            Unload frmData
            selCell.Col1 = -1: selCell.Row1 = -1
            Call ReShowGrid(intCurID)
            BlnSave = False
        End If
    End If
End Sub

Private Sub mnuClass_Insert_Before_Click()
'���ܣ��Ի��ܱ��,�Բ����������Ŀ��ǰ
    Dim intState As Integer
    Dim tmpID As RelatID, tmpItem As RPTItem
    Dim objRelations As New RPTRelations
    Dim objColProtertys As New RPTColProtertys
    Dim i As Long
    
    If selCell.Col1 = -1 And selCell.Row1 = -1 Then sta.Panels(2).Text = "����ȷ����ǰ��Ŀ���ͼ�λ�ã�": PlayWarn: Exit Sub
    
    '����ܱ��ͳ�������
    For Each tmpID In objReport.Items("_" & intCurID).SubIDs
        If objReport.Items("_" & tmpID.id).���� = 9 Then intState = intState + 1
    Next
    
    If selCell.Col1 <= msh(intCurID).FixedCols - 1 And selCell.Row1 >= msh(intCurID).FixedRows - 1 Then
        '������෶Χ
        Set frmData.frmParent = Me
        frmData.I_strTitle = "�������������"
        frmData.I_bytType = 0
        frmData.I_strClass = objReport.Items("_" & intCurID).����
        frmData.mintEleID = intCurID
        frmData.IO_strNode = ""
        Set frmData.objReport = objReport
        Set frmData.mobjRelations = objRelations
        frmData.Show 1, Me
        If gblnOK Then
            '������Ŀ�����ظ�
            For Each tmpID In objReport.Items("_" & intCurID).SubIDs
                If objReport.Items("_" & tmpID.id).���� = frmData.txtData.Text Then
                    sta.Panels(2).Text = "����ѡ����������ڸû��ܱ�����Ѿ����ڣ�"
                    Unload frmData: PlayWarn: Exit Sub
                End If
            Next
            For Each tmpID In objReport.Items("_" & intCurID).SubIDs
                Set tmpItem = objReport.Items("_" & tmpID.id)
                If tmpItem.���� = 7 And tmpItem.��� >= selCell.Col1 Then
                    tmpItem.��� = tmpItem.��� + 1
                End If
            Next
            intMaxID = intMaxID + 1
            Set tmpItem = objReport.Items.Add(intMaxID, mbytCurrFmt, "Ԫ��" & intMaxID, intCurID, 7, selCell.Col1, "", 0 _
                , frmData.txtData.Text, "", 0, 0, 1000, 0, 0, 0, False, "", 0, frmData.IO_FontBold = 1, False, False, 0 _
                , frmData.IO_FontColor, 0, False, 0, IIF(frmData.optDesc, ",", "") & frmData.txtOrder.Text, "", "" _
                , False, False, , False, , , , "_" & intMaxID)
            
            objReport.Items("_" & intCurID).SubIDs.Add intMaxID, "_" & intMaxID
            '����������
            Set tmpItem.Relations = objRelations
            Unload frmData
            selCell.Col1 = -1: selCell.Row1 = -1
            Call ReShowGrid(intCurID)
            BlnSave = False
        End If
    ElseIf selCell.Row1 <= msh(intCurID).FixedRows - 2 Then
        '������෶Χ
        Set frmData.frmParent = Me
        frmData.I_strTitle = "�������������"
        frmData.I_bytType = 1
        frmData.I_strClass = objReport.Items("_" & intCurID).����
        frmData.mintEleID = intCurID
        frmData.IO_strNode = ""
        Set frmData.objReport = objReport
        Set frmData.mobjRelations = objRelations
        frmData.Show 1, Me
        If gblnOK Then
            '������Ŀ�����ظ�
            For Each tmpID In objReport.Items("_" & intCurID).SubIDs
                If objReport.Items("_" & tmpID.id).���� = frmData.txtData Then
                    sta.Panels(2).Text = "����ѡ����������ڸû��ܱ�����Ѿ����ڣ�"
                    Unload frmData: PlayWarn: Exit Sub
                End If
            Next
            For Each tmpID In objReport.Items("_" & intCurID).SubIDs
                Set tmpItem = objReport.Items("_" & tmpID.id)
                If tmpItem.���� = 8 And tmpItem.��� >= selCell.Row1 Then
                    tmpItem.��� = tmpItem.��� + 1
                End If
            Next
            intMaxID = intMaxID + 1
            Set tmpItem = objReport.Items.Add(intMaxID, mbytCurrFmt, "Ԫ��" & intMaxID, intCurID, 8, selCell.Row1, "", 0 _
                , frmData.txtData.Text, "", 0, 0, 0, 0, 0, 0, False, "", 0, frmData.IO_FontBold = 1, False, False, 0 _
                , frmData.IO_FontColor, 0, False, 0, IIF(frmData.optDesc, ",", "") & frmData.txtOrder.Text, "", "" _
                , False, False, , False, , , , "_" & intMaxID)
                
            objReport.Items("_" & intCurID).SubIDs.Add intMaxID, "_" & intMaxID
            '����������
            Set tmpItem.Relations = objRelations
            Unload frmData
            selCell.Col1 = -1: selCell.Row1 = -1
            Call ReShowGrid(intCurID)
            BlnSave = False
        End If
    ElseIf selCell.Col1 >= msh(intCurID).FixedCols And selCell.Col1 <= msh(intCurID).FixedCols + intState - 1 _
        And selCell.Row1 >= msh(intCurID).FixedRows - 1 Then
        'ͳ���Χ
        Set frmData.frmParent = Me
        frmData.I_strTitle = "ͳ��������"
        frmData.I_bytType = 2
        frmData.I_strClass = objReport.Items("_" & intCurID).����
        frmData.mintEleID = intCurID
        frmData.IO_strNode = ""
        Set frmData.objReport = objReport
        Set frmData.mobjRelations = objRelations
        Set frmData.mobjColProtertys = objColProtertys
        frmData.I_strSummaryFile = GetSummaryFile(intCurID)
        frmData.Show 1, Me
        If gblnOK Then
            '������Ŀ�����ظ�
            For Each tmpID In objReport.Items("_" & intCurID).SubIDs
                If objReport.Items("_" & tmpID.id).���� = frmData.txtData.Text Then
                    sta.Panels(2).Text = "����ѡ����������ڸû��ܱ�����Ѿ����ڣ�"
                    Unload frmData: PlayWarn: Exit Sub
                End If
            Next
            For Each tmpID In objReport.Items("_" & intCurID).SubIDs
                Set tmpItem = objReport.Items("_" & tmpID.id)
                If tmpItem.���� = 9 And tmpItem.��� >= selCell.Col1 - msh(intCurID).FixedCols Then
                    tmpItem.��� = tmpItem.��� + 1
                End If
            Next
            intMaxID = intMaxID + 1
            Set tmpItem = objReport.Items.Add(intMaxID, mbytCurrFmt, "Ԫ��" & intMaxID, intCurID, 9 _
                , selCell.Col1 - msh(intCurID).FixedCols, "", 0, frmData.txtData.Text, "", 0, 0, 1000, 0, 0, 0, False, "", 0 _
                , frmData.IO_FontBold = 1, False, False, 0, frmData.IO_FontColor, 0, False, 0, "", frmData.txtFormat.Text, "" _
                , False, False, , False, , , , "_" & intMaxID)
                
            objReport.Items("_" & intCurID).SubIDs.Add intMaxID, "_" & intMaxID
            '����������
            Set tmpItem.Relations = objRelations
            '����������
            Set tmpItem.ColProtertys = frmData.mobjColProtertys
            Unload frmData
            selCell.Col1 = -1: selCell.Row1 = -1
            Call ReShowGrid(intCurID)
            BlnSave = False
        End If
    End If
End Sub

Private Sub mnuClass_State_Style_Click(Index As Integer)
'���ܣ��Ի��ܱ��,���õ�ǰ������ܷ�ʽ
    Dim tmpID As RelatID, tmpItem As RPTItem
    
    If selCell.Col1 = -1 And selCell.Row1 = -1 Then sta.Panels(2).Text = "����ȷ����ǰ��Ŀ���ͼ�λ�ã�": PlayWarn: Exit Sub
    
    If selCell.Col1 <= msh(intCurID).FixedCols - 1 And selCell.Row1 >= msh(intCurID).FixedRows - 1 Then
        '������෶Χ
        For Each tmpID In objReport.Items("_" & intCurID).SubIDs
            Set tmpItem = objReport.Items("_" & tmpID.id)
            If tmpItem.���� = 7 And tmpItem.��� = selCell.Col1 Then
                Select Case Index
                    Case 0
                        tmpItem.���� = ""
                    Case 1
                        tmpItem.���� = "SUM"
                    Case 2
                        tmpItem.���� = "AVG"
                    Case 3
                        tmpItem.���� = "MAX"
                    Case 4
                        tmpItem.���� = "MIN"
                    Case 5
                        tmpItem.���� = "COUNT"
                End Select
                If tmpItem.���� = "" Then
                    msh(intCurID).TextMatrix(msh(intCurID).FixedRows, selCell.Col1) = msh(intCurID).TextMatrix(msh(intCurID).FixedRows + 1, selCell.Col1)
                Else
                    msh(intCurID).TextMatrix(msh(intCurID).FixedRows, selCell.Col1) = tmpItem.����
                End If
                Exit For
            End If
        Next
    ElseIf selCell.Row1 <= msh(intCurID).FixedRows - 2 Then
        '������෶Χ
        For Each tmpID In objReport.Items("_" & intCurID).SubIDs
            Set tmpItem = objReport.Items("_" & tmpID.id)
            If tmpItem.���� = 8 And tmpItem.��� = selCell.Row1 Then
                Select Case Index
                    Case 0
                        tmpItem.���� = ""
                    Case 1
                        tmpItem.���� = "SUM"
                    Case 2
                        tmpItem.���� = "AVG"
                    Case 3
                        tmpItem.���� = "MAX"
                    Case 4
                        tmpItem.���� = "MIN"
                    Case 5
                        tmpItem.���� = "COUNT"
                End Select
                If tmpItem.���� = "" Then
                    msh(intCurID).TextMatrix(selCell.Row1, msh(intCurID).FixedCols) = msh(intCurID).TextMatrix(selCell.Row1, msh(intCurID).FixedCols + 1)
                Else
                    msh(intCurID).TextMatrix(selCell.Row1, msh(intCurID).FixedCols) = tmpItem.����
                End If
                msh(intCurID).MergeCol(msh(intCurID).FixedCols) = False
                Exit For
            End If
        Next
    End If
    BlnSave = False
End Sub

Private Sub mnuCustom_Col_Align_Style_Click(Index As Integer)
'���ܣ���������,���õ�ǰ�����ݶ��뷽ʽ
    Dim tmpItem As RPTItem, tmpID As RelatID
    
    If intCurCol = -1 Then sta.Panels(2).Text = "����ȷ����ǰ�����У�": PlayWarn: Exit Sub
    
    For Each tmpID In objReport.Items("_" & intCurID).SubIDs
        Set tmpItem = objReport.Items("_" & tmpID.id)
        If tmpItem.��� = intCurCol Then tmpItem.���� = Index
    Next
    msh(intCurID).ColAlignment(intCurCol) = Switch(Index = 0, 1, Index = 1, 4, Index = 2, 7)
    
    Call SetCopyGrid(intCurID)
    BlnSave = False
End Sub

Private Sub mnuCustom_Col_Clear_Click()
    Dim intCol As Integer, tmpRelatID As RelatID
    '��ձ���(�ѹ�ʽ���)
    If intCurID = 0 Then Exit Sub
    For Each tmpRelatID In objReport.Items("_" & intCurID).SubIDs
        objReport.Items("_" & tmpRelatID.Key).���� = ""
    Next
    Call ReShowGrid(intCurID)
    selCell.Col1 = -1: selCell.Col2 = -1: selCell.Row1 = -1: selCell.Row2 = -1: selCell.Row = -1
    Call ShowPaperInfo
    BlnSave = False
End Sub

Private Sub mnuCustom_Col_Data_Click()
'���ܣ���������,���õ�ǰ�е����ݼ��㷽ʽ
    Dim tmpID As RelatID, tmpItem As RPTItem
    Dim strFormula As String, strText As String, blnDo As Boolean
    Dim blnPreMerge As Boolean
    Dim k As Long, X As Long, Y As Long
    Dim objRelations As New RPTRelations
    Dim objColProtertys As New RPTColProtertys
    Dim i As Long
    
    If intCurCol = -1 Then sta.Panels(2).Text = "����ȷ����ǰ�����У�": PlayWarn: Exit Sub
    
    blnPreMerge = True
    For Each tmpID In objReport.Items("_" & intCurID).SubIDs
        If objReport.Items("_" & tmpID.id).��� = intCurCol Then
            Set tmpItem = objReport.Items("_" & tmpID.id)
        ElseIf objReport.Items("_" & tmpID.id).��� = intCurCol - 1 Then
            blnPreMerge = objReport.Items("_" & tmpID.id).�Ե�
        End If
    Next
    
    With frmFormula
        .strInit = tmpItem.����
        .strFormat = tmpItem.��ʽ
        .mbln��ҳ = tmpItem.�߿�
        .mblnCan��ҳ = True 'objReport.Items("_" & intCurID).���� <= 1
        .mblnMerge = tmpItem.�Ե�
        .mblnPreMerge = blnPreMerge
        .mblnVisible = tmpItem.���� = 1
        .intCol = intCurCol
        .intCur = intCurID
        
        '�иߣ���С���壩������Ӧ�и߻���
        If tmpItem.�и� = 1 Then
            .mblnAutoFont = True
            .mblnAutoRowHeight = False
        Else
            .mblnAutoFont = False
            .mblnAutoRowHeight = tmpItem.����Ӧ�и�
        End If
        
        Set .frmParent = Me
        Set .objReport = objReport
        Call SetCopyRelations(tmpItem.Relations, objRelations)
        Call SetCopyColProtertys(tmpItem.ColProtertys, objColProtertys)
        Set .mobjRelations = objRelations
        Set .mobjColProtertys = objColProtertys
        
        .Show vbModal, Me
    End With
    
    If gblnOK Then
        If tmpItem.�ϼ�ID <> 0 Then
            If objReport.Items("_" & tmpItem.�ϼ�ID).��ID <> 0 Then
                If objReport.Items("_" & objReport.Items("_" & tmpItem.�ϼ�ID).��ID).����Դ <> "" Then
                    X = InStr(1, frmFormula.txtFormula.Text, "]")
                    Y = InStr(1, frmFormula.txtFormula.Text, ".")
                    k = InStr(1, frmFormula.txtFormula.Text, "[")
                    If X > k And X > Y And X <> 0 And k <> 0 And Y <> 0 Then
                        If Mid(frmFormula.txtFormula.Text, k + 1, Y - k - 1) <> _
                            objReport.Items("_" & objReport.Items("_" & tmpItem.�ϼ�ID).��ID).����Դ Then
                            MsgBox "�󶨵������б������ڵ�ǰ��Ƭ����Դ�����飡", vbInformation, App.Title
                            Unload frmFormula
                            Exit Sub
                        End If
                    End If
                End If
            End If
        End If
        tmpItem.���� = frmFormula.txtFormula.Text
        tmpItem.��ʽ = frmFormula.txtFormat.Text
        If frmFormula.chkAutoFont.Value = 1 Then
            tmpItem.�и� = 1
            tmpItem.����Ӧ�и� = False
        Else
            tmpItem.�и� = 0
            tmpItem.����Ӧ�и� = frmFormula.chkAutoRowHeight.Value = 1
        End If
        tmpItem.���� = frmFormula.chkVisible.Value
        tmpItem.�Ե� = (frmFormula.chkMerge.Value = 1)
        tmpItem.�߿� = (frmFormula.chk��ҳ.Value = 1)
        If tmpItem.���� = "" Then tmpItem.���� = ""
        '����������
        Set tmpItem.Relations = objRelations
        
        '����������
        Set tmpItem.ColProtertys = frmFormula.mobjColProtertys
        
        Unload frmFormula
        msh(intCurID).TextMatrix(msh(intCurID).FixedRows, intCurCol) = tmpItem.����
        msh(intCurID).TextMatrix(msh(intCurID).FixedRows + 1, intCurCol) = tmpItem.����
        
        If Not tmpItem.�Ե� Then
            For Each tmpID In objReport.Items("_" & intCurID).SubIDs
                If objReport.Items("_" & tmpID.id).��� > intCurCol Then
                    objReport.Items("_" & tmpID.id).�Ե� = False
                End If
            Next
        End If
        If tmpItem.�߿� Then
            For Each tmpID In objReport.Items("_" & intCurID).SubIDs
                If objReport.Items("_" & tmpID.id).��� <> intCurCol Then
                    objReport.Items("_" & tmpID.id).�߿� = False
                End If
            Next
        End If
        
        '��ͷ����
        If Right(tmpItem.��ͷ, 1) = "#" And tmpItem.���� <> "" Then
            If Left(tmpItem.����, 1) = "[" And Right(tmpItem.����, 1) = "]" And InStr(tmpItem.����, ".") > 0 _
                And InStr(Mid(tmpItem.����, 2, Len(tmpItem.����) - 2), "[") = 0 Then
                
                strText = Mid(tmpItem.����, InStr(tmpItem.����, ".") + 1, Len(tmpItem.����) - 1 - InStr(tmpItem.����, "."))
                
                blnDo = True
                On Error Resume Next
                If msh(intCurID).TextMatrix(msh(intCurID).FixedRows - 1, intCurCol - 1) = strText And _
                    msh(intCurID).TextMatrix(msh(intCurID).FixedRows - 2, intCurCol) = strText Then
                    If Err.Number = 0 Then blnDo = False
                End If
                If msh(intCurID).TextMatrix(msh(intCurID).FixedRows - 1, intCurCol + 1) = strText And _
                    msh(intCurID).TextMatrix(msh(intCurID).FixedRows - 2, intCurCol) = strText Then
                    If Err.Number = 0 Then blnDo = False
                End If
                On Error GoTo 0
                
                If blnDo Then
                    tmpItem.��ͷ = Left(tmpItem.��ͷ, Len(tmpItem.��ͷ) - 1) & strText
                    msh(intCurID).TextMatrix(msh(intCurID).FixedRows - 1, intCurCol) = strText
                End If
            End If
        End If
        Call SetCopyGrid(intCurID)
        BlnSave = False
    End If
    Call ShowAttrib(intCurID)       'ˢ����������
End Sub

Private Sub mnuCustom_Col_Del_Click()
'���ܣ���������,ɾ����ǰ������
    Dim tmpItem As RPTItem, tmpID As RelatID
    Dim intDel As Integer
    
    If intCurCol = -1 Then sta.Panels(2).Text = "����ȷ����ǰ�����У�": PlayWarn: Exit Sub
    If objReport.Items("_" & intCurID).SubIDs.count < 2 Then sta.Panels(2).Text = "����ȫ��ɾ������У�": PlayWarn: Exit Sub
    
    For Each tmpID In objReport.Items("_" & intCurID).SubIDs
        Set tmpItem = objReport.Items("_" & tmpID.id)
        If tmpItem.��� = intCurCol Then
            intDel = tmpItem.id
        ElseIf tmpItem.��� > intCurCol Then
            tmpItem.��� = tmpItem.��� - 1
        End If
    Next
    
    objReport.Items.Remove "_" & intDel
    objReport.Items("_" & intCurID).SubIDs.Remove "_" & intDel
    
    Call ReShowGrid(intCurID)
    intCurCol = -1
    BlnSave = False
End Sub

Private Sub mnuCustom_Col_Insert_Left_Click()
'���ܣ���������,�ڵ�ǰ����߲���һ����
    Dim tmpItem As RPTItem, tmpID As RelatID
    Dim strHead As String, i As Integer
    
    If intCurCol = -1 Then sta.Panels(2).Text = "����ȷ����ǰ�����У�": PlayWarn: Exit Sub
    
    For Each tmpID In objReport.Items("_" & intCurID).SubIDs
        Set tmpItem = objReport.Items("_" & tmpID.id)
        If tmpItem.��� >= intCurCol Then tmpItem.��� = tmpItem.��� + 1
    Next
    
    For i = 0 To msh(intCurID).FixedRows - 1
        strHead = strHead & "|4^" & msh(intCurID).RowHeight(i) & "^#"
    Next
    strHead = Mid(strHead, 2)
    i = objReport.Items("_" & intCurID).id
    intMaxID = intMaxID + 1
    
    objReport.Items.Add intMaxID, mbytCurrFmt, "Ԫ��" & intMaxID, i, 6, intCurCol, "", 0, "", strHead, 0, 0, 1000, 0, 0, 0 _
        , False, "", 0, False, False, False, 0, 0, 0, False, 0, "", "", "", False, False, , False, , , , "_" & intMaxID
    objReport.Items("_" & intCurID).SubIDs.Add intMaxID, "_" & intMaxID
    
    Call ReShowGrid(intCurID)
    intCurCol = -1
    BlnSave = False
End Sub

Private Sub mnuCustom_Col_Insert_Right_Click()
'���ܣ���������,�ڵ�ǰ���ұ߲���һ����
    Dim tmpItem As RPTItem, tmpID As RelatID
    Dim strHead As String, i As Integer
    
    If intCurCol = -1 Then sta.Panels(2).Text = "����ȷ����ǰ�����У�": PlayWarn: Exit Sub
    
    For Each tmpID In objReport.Items("_" & intCurID).SubIDs
        Set tmpItem = objReport.Items("_" & tmpID.id)
        If tmpItem.��� > intCurCol Then tmpItem.��� = tmpItem.��� + 1
    Next
    
    For i = 0 To msh(intCurID).FixedRows - 1
        strHead = strHead & "|4^" & msh(intCurID).RowHeight(i) & "^#"
    Next
    strHead = Mid(strHead, 2)
    i = objReport.Items("_" & intCurID).id
    intMaxID = intMaxID + 1
    
    objReport.Items.Add intMaxID, mbytCurrFmt, "Ԫ��" & intMaxID, i, 6, intCurCol + 1, "", 0, "", strHead, 0, 0, 1000, 0, 0, 0 _
        , False, "", 0, False, False, False, 0, 0, 0, False, 0, "", "", "", False, False, , False, , , , "_" & intMaxID
    objReport.Items("_" & intCurID).SubIDs.Add intMaxID, "_" & intMaxID
    
    Call ReShowGrid(intCurID)
    intCurCol = -1
    BlnSave = False
End Sub

Private Sub mnuCustom_Col_State_Style_Click(Index As Integer)
'���ܣ���������,���õ�ǰ�еĻ��ܷ�ʽ
    Dim tmpItem As RPTItem, tmpID As RelatID
    
    If intCurCol = -1 Then sta.Panels(2).Text = "����ȷ����ǰ�����У�": PlayWarn: Exit Sub
    If msh(intCurID).TextMatrix(msh(intCurID).FixedRows, intCurCol) = "" Then sta.Panels(2).Text = "����û�ж���������Դ,���ܻ��ܣ�": PlayWarn: Exit Sub
    
    For Each tmpID In objReport.Items("_" & intCurID).SubIDs
        Set tmpItem = objReport.Items("_" & tmpID.id)
        If tmpItem.��� = intCurCol Then
            tmpItem.���� = Switch(Index = 0, "", Index = 1, "SUM", Index = 2, "AVG", Index = 3, "MAX", Index = 4, "MIN", Index = 5, "COUNT")
        End If
    Next
    msh(intCurID).TextMatrix(msh(intCurID).FixedRows + 1, intCurCol) = Switch(Index = 0, "", Index = 1, "SUM", Index = 2, "AVG", Index = 3, "MAX", Index = 4, "MIN", Index = 5, "COUNT")
    
    Call SetCopyGrid(intCurID)
    BlnSave = False
End Sub

Private Sub mnuCustom_Head_Auto_Click()
'���ܣ��Զ�����ǰ��ͷ����Ϊ�б�Ų�����
    Dim intBegin As Integer, i As Integer
    Dim arrHead() As String, IntAlig As Integer
    Dim tmpID As RelatID, tmpItem As RPTItem
    
    If selCell.Row = -1 Then sta.Panels(2).Text = "����ȷ����ǰ�У�": PlayWarn: Exit Sub
    
    frmInput.I_blnAllowNULL = False
    frmInput.I_bytType = 1
    frmInput.I_intMaxLen = 2
    frmInput.I_strInfo = "������������������д��ǰ��ͷ�п�ʼ���б�ţ�"
    frmInput.I_strTitle = "�Զ��б��"
    frmInput.I_strMask = "0123456789"
    frmInput.IO_strValue = "0"
    frmInput.Show 1, Me

    If gblnOK Then
        intBegin = CInt(frmInput.IO_strValue)
        IntAlig = frmInput.IO_IntAlig
        Unload frmInput
        
        For Each tmpID In objReport.Items("_" & intCurID).SubIDs
            Set tmpItem = objReport.Items("_" & tmpID.id)
            arrHead = Split(tmpItem.��ͷ, "|")
            tmpItem.��ͷ = ""
            For i = 0 To UBound(arrHead)
                If i = selCell.Row Then
                    msh(intCurID).Row = i: msh(intCurID).Col = tmpItem.���
                    msh(intCurID).CellAlignment = IntAlig
                    msh(intCurID).CellForeColor = frmInput.IO_FontColor
                    msh(intCurID).CellFontBold = frmInput.IO_FontBold
                    tmpItem.��ͷ = tmpItem.��ͷ & "|" & msh(intCurID).CellAlignment & "^" & msh(intCurID).RowHeight(i) & "^<" & intBegin + tmpItem.��� & ">" & "^" & IIF(msh(intCurID).CellFontBold, 1, 0) & "^" & msh(intCurID).CellForeColor
                Else
                    tmpItem.��ͷ = tmpItem.��ͷ & "|" & arrHead(i)
                End If
            Next
            tmpItem.��ͷ = Mid(tmpItem.��ͷ, 2)
        Next
        
        Call ReShowGrid(intCurID)
        selCell.Col1 = -1: selCell.Col2 = -1: selCell.Row1 = -1: selCell.Row2 = -1: selCell.Row = -1
        Call ShowPaperInfo
        BlnSave = False
    End If
End Sub

Private Sub mnuCustom_Head_Clear_Click()
    Dim intCol As Integer, tmpRelatID As RelatID
    '��ձ�ͷ,ֻ����һ�й̶���
    If intCurID = 0 Then Exit Sub
    msh(intCurID).FixedRows = 1
    
    For intCol = 0 To msh(intCurID).Cols - 1
        msh(intCurID).TextMatrix(0, intCol) = ""
    Next
    For Each tmpRelatID In objReport.Items("_" & intCurID).SubIDs
        msh(intCurID).Row = 0: msh(intCurID).Col = objReport.Items("_" & tmpRelatID.Key).���
        msh(intCurID).RowHeight(0) = 30
        objReport.Items("_" & tmpRelatID.Key).��ͷ = msh(intCurID).CellAlignment & "^" & msh(intCurID).RowHeight(0) & "^#"
    Next
    Call ReShowGrid(intCurID)
    selCell.Col1 = -1: selCell.Col2 = -1: selCell.Row1 = -1: selCell.Row2 = -1: selCell.Row = -1
    Call ShowPaperInfo
    BlnSave = False
End Sub

Private Sub mnuCustom_Head_Del_Click()
'���ܣ���������,ɾ����ǰѡ����
    Dim tmpID As RelatID, tmpItem As RPTItem, StrDelName As String
    Dim arrHead() As String, i As Integer
    
    If selCell.Row = -1 Then sta.Panels(2).Text = "����ȷ����ǰ�У�": PlayWarn: Exit Sub
    If msh(intCurID).FixedRows = 1 Then sta.Panels(2).Text = "��ͷ����Ҫ����һ�У�": PlayWarn: Exit Sub
    
    For Each tmpID In objReport.Items("_" & intCurID).SubIDs
        Set tmpItem = objReport.Items("_" & tmpID.id)
        arrHead = Split(tmpItem.��ͷ, "|")
        tmpItem.��ͷ = ""
        For i = 0 To UBound(arrHead)
            'ɾ����ǰ������
            If i <> selCell.Row Then
                If i = selCell.Row + 1 Then
                    StrDelName = msh(intCurID).TextMatrix(i, tmpItem.���)
                    msh(intCurID).Row = i: msh(intCurID).Col = tmpItem.���
                    tmpItem.��ͷ = tmpItem.��ͷ & "|" & msh(intCurID).CellAlignment & "^" & msh(intCurID).RowHeight(i) & "^" & StrDelName
                Else
                    tmpItem.��ͷ = tmpItem.��ͷ & "|" & arrHead(i)
                End If
            End If
        Next
        tmpItem.��ͷ = Mid(tmpItem.��ͷ, 2)
    Next
    Call ReShowGrid(intCurID)
    selCell.Col1 = -1: selCell.Col2 = -1: selCell.Row1 = -1: selCell.Row2 = -1: selCell.Row = -1
    Call ShowPaperInfo
    BlnSave = False
End Sub

Private Sub mnuCustom_Head_Insert_Down_Click()
'���ܣ����������ͷ,�ڵ�ǰѡ�����·�����һ������
'˵������SelCell.RowΪ������
    Dim tmpID As RelatID, tmpItem As RPTItem
    Dim arrHead() As String, i As Integer
    
    If selCell.Row = -1 Then sta.Panels(2).Text = "����ȷ����ǰ�У�": PlayWarn: Exit Sub
    For Each tmpID In objReport.Items("_" & intCurID).SubIDs
        Set tmpItem = objReport.Items("_" & tmpID.id)
        arrHead = Split(tmpItem.��ͷ, "|")
        tmpItem.��ͷ = ""
        For i = 0 To UBound(arrHead)
            If i = selCell.Row Then
                tmpItem.��ͷ = tmpItem.��ͷ & "|" & arrHead(i) & "|4^255^#" '�²���������
            Else
                tmpItem.��ͷ = tmpItem.��ͷ & "|" & arrHead(i)
            End If
        Next
        tmpItem.��ͷ = Mid(tmpItem.��ͷ, 2)
    Next
    Call ReShowGrid(intCurID)
    selCell.Col1 = -1: selCell.Col2 = -1: selCell.Row1 = -1: selCell.Row2 = -1: selCell.Row = -1
    Call ShowPaperInfo
    BlnSave = False
End Sub

Private Sub mnuCustom_Head_Insert_UP_Click()
'���ܣ����������ͷ,�ڵ�ǰѡ�����Ϸ�����һ������
'˵������SelCell.RowΪ������
    Dim tmpID As RelatID, tmpItem As RPTItem
    Dim arrHead() As String, i As Integer
    
    If selCell.Row = -1 Then sta.Panels(2).Text = "����ȷ����ǰ�У�": PlayWarn: Exit Sub
    For Each tmpID In objReport.Items("_" & intCurID).SubIDs
        Set tmpItem = objReport.Items("_" & tmpID.id)
        arrHead = Split(tmpItem.��ͷ, "|")
        tmpItem.��ͷ = ""
        For i = 0 To UBound(arrHead)
            If i = selCell.Row Then
                tmpItem.��ͷ = tmpItem.��ͷ & "|4^255^#|" & arrHead(i) '�²���������
            Else
                tmpItem.��ͷ = tmpItem.��ͷ & "|" & arrHead(i)
            End If
        Next
        tmpItem.��ͷ = Mid(tmpItem.��ͷ, 2)
    Next
    Call ReShowGrid(intCurID)
    selCell.Col1 = -1: selCell.Col2 = -1: selCell.Row1 = -1: selCell.Row2 = -1: selCell.Row = -1
    Call ShowPaperInfo
    BlnSave = False
End Sub

Private Sub mnuCustom_Head_Merge_Click()
'���ܣ����������ͷ,�ϲ���ǰѡ��Χ�ĵ�Ԫ��
'˵�������ú������ǰѡ��Ԫ��
    Dim i As Integer, j As Integer, strText As String
    Dim tmpID As RelatID, tmpItem As RPTItem
    Dim arrHead() As String, blnDo As Boolean
    
    If selCell.Row1 = -1 Then sta.Panels(2).Text = "����ȷ����ǰѡ��Ԫ��": PlayWarn: Exit Sub
    If selCell.Row1 = selCell.Row2 And selCell.Col1 = selCell.Col2 Then sta.Panels(2).Text = "һ����Ԫ���úϲ���": PlayWarn: Exit Sub
    If selCell.Row1 <> selCell.Row2 And selCell.Col1 <> selCell.Col2 Then sta.Panels(2).Text = "��Ԫ��ͬʱֻ����һ�������Ϻϲ���": PlayWarn: Exit Sub
    '�����ǰѡ��Χ��Ԫ������ȫ����ͬ���úϲ�
    For i = selCell.Row1 To selCell.Row2
        For j = selCell.Col1 To selCell.Col2
            If i = selCell.Row1 And j = selCell.Col1 Then
                strText = msh(intCurID).TextMatrix(i, j)
            Else
                If msh(intCurID).TextMatrix(i, j) <> strText Or strText = "" Then
                    blnDo = True: Exit For
                End If
            End If
        Next
    Next
    If Not blnDo Then sta.Panels(2).Text = "��ǰѡ��Ԫ���Ѿ��ϲ���": PlayWarn: Exit Sub
    
    frmInput.I_strTitle = "�ϲ���Ԫ��"
    frmInput.I_strInfo = "��������������������������ı�ͷ��Ԫ��ϲ�������֡�"
    frmInput.I_blnAllowNULL = False
    frmInput.I_intMaxLen = 50
    frmInput.IO_strValue = ""
    For i = selCell.Row1 To selCell.Row2
        For j = selCell.Col1 To selCell.Col2
            If msh(intCurID).TextMatrix(i, j) <> "" Then
                frmInput.IO_strValue = msh(intCurID).TextMatrix(i, j)
                msh(intCurID).Row = i: msh(intCurID).Col = j
                frmInput.IO_FontBold = IIF(msh(intCurID).CellFontBold, 1, 0)
                frmInput.IO_FontColor = msh(intCurID).CellForeColor
                Exit For
            End If
        Next
    Next
    frmInput.Show 1, Me
    If gblnOK Then
        '���¿ؼ�
        For i = selCell.Row1 To selCell.Row2
            For j = selCell.Col1 To selCell.Col2
                If CheckCell(intCurID, i, j, frmInput.IO_strValue) Then
                    msh(intCurID).Row = i: msh(intCurID).Col = j
                    msh(intCurID).CellAlignment = frmInput.IO_IntAlig
                    msh(intCurID).TextMatrix(i, j) = frmInput.IO_strValue
                    msh(intCurID).CellForeColor = frmInput.IO_FontColor
                    msh(intCurID).CellFontBold = frmInput.IO_FontBold
                Else
                    sta.Panels(2).Text = "��ͷ��Ԫ��ͬʱֻ����һ�������ϱ��ϲ�,��Ԫ�����ֲ���ȫ��д�룡": PlayWarn
                End If
            Next
        Next
        Unload frmInput
        '���¶���
        For Each tmpID In objReport.Items("_" & intCurID).SubIDs
            Set tmpItem = objReport.Items("_" & tmpID.id)
            If tmpItem.��� >= selCell.Col1 And tmpItem.��� <= selCell.Col2 Then
                arrHead = Split(tmpItem.��ͷ, "|")
                tmpItem.��ͷ = ""
                For i = 0 To UBound(arrHead)
                    If i >= selCell.Row1 And i <= selCell.Row2 Then
                        msh(intCurID).Row = i: msh(intCurID).Col = tmpItem.���
                        tmpItem.��ͷ = tmpItem.��ͷ & "|" & msh(intCurID).CellAlignment & "^" & msh(intCurID).RowHeight(i) & "^" & IIF(msh(intCurID).TextMatrix(i, tmpItem.���) = "", "#", msh(intCurID).TextMatrix(i, tmpItem.���)) & "^" & IIF(msh(intCurID).CellFontBold, 1, 0) & "^" & msh(intCurID).CellForeColor
                    Else
                        tmpItem.��ͷ = tmpItem.��ͷ & "|" & arrHead(i)
                    End If
                Next
                tmpItem.��ͷ = Mid(tmpItem.��ͷ, 2)
            End If
        Next
        Call SetCopyGrid(intCurID)
        BlnSave = False
    End If
End Sub

Private Sub mnuCustom_Head_Split_Click()
'���ܣ����������ͷ,��ֵ�ǰѡ��Ԫ��
    Dim i As Integer, j As Integer, strText As String
    Dim tmpID As RelatID, tmpItem As RPTItem, arrHead() As String
    
    If selCell.Row1 = -1 Then sta.Panels(2).Text = "����ȷ����ǰѡ��ĵ�Ԫ��": PlayWarn: Exit Sub
    If selCell.Row1 = selCell.Row2 And selCell.Col1 = selCell.Col2 Then sta.Panels(2).Text = "һ����Ԫ���ò�֣�": PlayWarn: Exit Sub
    
    '��ǰѡ��Χ��Ԫ�����ݱ���ȫ����ͬ
    For i = selCell.Row1 To selCell.Row2
        For j = selCell.Col1 To selCell.Col2
            If i = selCell.Row1 And j = selCell.Col1 Then
                strText = msh(intCurID).TextMatrix(i, j)
            Else
                If msh(intCurID).TextMatrix(i, j) <> strText Or strText = "" Then
                    sta.Panels(2).Text = "��ǰѡ��ķ�Χ��ֻһ����Ԫ��,���ܲ�֣�": PlayWarn: Exit Sub
                End If
            End If
        Next
    Next
    
    '��ֵ�Ԫ�������Ϊ��
    '���¿ؼ�����
    For i = selCell.Row1 To selCell.Row2
        For j = selCell.Col1 To selCell.Col2
            If Not (i = selCell.Row1 And j = selCell.Col1) Then
                msh(intCurID).TextMatrix(i, j) = ""
            End If
        Next
    Next
    
    '���¶�������
    For Each tmpID In objReport.Items("_" & intCurID).SubIDs
        Set tmpItem = objReport.Items("_" & tmpID.id)
        If tmpItem.��� >= selCell.Col1 And tmpItem.��� <= selCell.Col2 Then
            arrHead = Split(tmpItem.��ͷ, "|")
            tmpItem.��ͷ = ""
            For i = 0 To UBound(arrHead)
                If i >= selCell.Row1 And i <= selCell.Row2 Then
                    msh(intCurID).Row = i: msh(intCurID).Col = tmpItem.���
                    tmpItem.��ͷ = tmpItem.��ͷ & "|" & msh(intCurID).CellAlignment & "^" & msh(intCurID).RowHeight(i) & "^" & IIF(msh(intCurID).TextMatrix(i, tmpItem.���) = "", "#", msh(intCurID).TextMatrix(i, tmpItem.���))
                Else
                    tmpItem.��ͷ = tmpItem.��ͷ & "|" & arrHead(i)
                End If
            Next
            tmpItem.��ͷ = Mid(tmpItem.��ͷ, 2)
        End If
    Next
    Call SetCopyGrid(intCurID)
    BlnSave = False
End Sub

Private Sub mnuCustom_Head_Text_Click()
'���ܣ����������ͷ,�ڵ���Ԫ��(�����Ǻϲ���)������������
'˵�������ú������ǰѡ��Ԫ��
    Dim i As Integer, j As Integer, strText As String
    Dim tmpID As RelatID, tmpItem As RPTItem
    Dim arrHead() As String
    
    If selCell.Row1 = -1 Then sta.Panels(2).Text = "����ȷ����ǰѡ��ĵ�Ԫ��": PlayWarn: Exit Sub
    
    '��ǰѡ��Χ��Ԫ�����ݱ���ȫ����ͬ
    For i = selCell.Row1 To selCell.Row2
        For j = selCell.Col1 To selCell.Col2
            If i = selCell.Row1 And j = selCell.Col1 Then
                strText = msh(intCurID).TextMatrix(i, j)
            Else
                If msh(intCurID).TextMatrix(i, j) <> strText Or strText = "" Then
                    sta.Panels(2).Text = "��ǰѡ��ķ�Χ��ֻһ����Ԫ��,���Ⱥϲ���": PlayWarn: Exit Sub
                End If
            End If
        Next
    Next
    
    With frmInput
        .I_strTitle = "��Ԫ������"
        .I_strInfo = "��������������������������ı�ͷ��ǰ��Ԫ������֡�"
        .I_blnAllowNULL = True
        .I_intMaxLen = 200
        .IO_strValue = msh(intCurID).TextMatrix(selCell.Row1, selCell.Col1)
        msh(intCurID).Row = selCell.Row1: msh(intCurID).Col = selCell.Col1
        .IO_IntAlig = msh(intCurID).CellAlignment
        .IO_FontBold = IIF(msh(intCurID).CellFontBold, 1, 0)
        .IO_FontColor = msh(intCurID).CellForeColor
        .Show 1, Me
    End With
    If gblnOK Then
        '���¿ؼ�
        For i = selCell.Row1 To selCell.Row2
            For j = selCell.Col1 To selCell.Col2
                If CheckCell(intCurID, i, j, frmInput.IO_strValue) Then
                    msh(intCurID).TextMatrix(i, j) = frmInput.IO_strValue
                    msh(intCurID).Row = i: msh(intCurID).Col = j
                    msh(intCurID).CellAlignment = frmInput.IO_IntAlig
                    msh(intCurID).CellForeColor = frmInput.IO_FontColor
                    msh(intCurID).CellFontBold = frmInput.IO_FontBold
                Else
                    sta.Panels(2).Text = "��ͷ��Ԫ��ͬʱֻ����һ�������ϱ��ϲ�,��Ԫ�����ֲ���ȫ��д�룡": PlayWarn
                End If
            Next
        Next
        Unload frmInput
        '���¶���
        For Each tmpID In objReport.Items("_" & intCurID).SubIDs
            Set tmpItem = objReport.Items("_" & tmpID.id)
            arrHead = Split(tmpItem.��ͷ, "|")
            tmpItem.��ͷ = ""
            
            For i = 0 To UBound(arrHead)
                If i >= selCell.Row1 And i <= selCell.Row2 Then
                    msh(intCurID).Row = i: msh(intCurID).Col = tmpItem.���
                    tmpItem.��ͷ = tmpItem.��ͷ & IIF(tmpItem.��ͷ = "", "", "|") & msh(intCurID).CellAlignment & "^" & msh(intCurID).RowHeight(i) & "^" & IIF(msh(intCurID).TextMatrix(i, tmpItem.���) = "", "#", msh(intCurID).TextMatrix(i, tmpItem.���)) & "^" & IIF(msh(intCurID).CellFontBold, 1, 0) & "^" & msh(intCurID).CellForeColor
                Else
                    tmpItem.��ͷ = tmpItem.��ͷ & IIF(tmpItem.��ͷ = "", "", "|") & arrHead(i)
                End If
            Next
        Next
        Call SetCopyGrid(intCurID)
        BlnSave = False
    End If
End Sub

Private Sub mnuEdit_AddFormat_Click()
    cmdAdd_Click
End Sub

Private Sub mnuEdit_Copy_Click()
'���ܣ�����ǰѡ��ı���Ԫ�ظ��Ƶ����������(objClip)
'˵����
'     1.�Զ�����X,Y����
'     2.��������������,ճ��ʱ�ٴ���
'     3.����Ԫ������������(������������),ճ��ʱ�ٴ���
    Dim tmpObj As PictureBox, tmpItem As RPTItem, tmpID As RelatID
    Dim tmpItem1 As RPTItem
    
    If GetSelNum = 0 Then PlayWarn: Exit Sub
    
    Set objClip = New RPTItems
    For Each tmpObj In lblSize
        If tmpObj.Index Mod 8 = 1 Then
            With objReport.Items("_" & tmpObj.Tag)
                Set tmpItem = objClip.Add(.id, .��ʽ��, .����, .�ϼ�ID, .����, .���, .����, .����, .����, .��ͷ, .X, .Y, .W, .H, _
                .�и�, .����, .�Ե�, .����, .�ֺ�, .����, .����, .б��, .����, .ǰ��, .����, .�߿�, .����, .����, .��ʽ, .����, .����߼Ӵ�, _
                .����Ӧ�и�, .ͼƬ, .ϵͳ, .��ID, .SubIDs, .CopyIDs, "_" & .id, .����Դ, .���¼��, .���Ҽ��, .Դ�к�, .�������, .�������)
                '��������
                If .SubIDs.count > 0 Then
                    For Each tmpID In .SubIDs
                        With objReport.Items("_" & tmpID.id)
                            objClip.Add .id, .��ʽ��, .����, .�ϼ�ID, .����, .���, .����, .����, .����, .��ͷ, .X, .Y, .W, .H, _
                                .�и�, .����, .�Ե�, .����, .�ֺ�, .����, .����, .б��, .����, .ǰ��, .����, .�߿�, .����, .����, _
                                .��ʽ, .����, .����߼Ӵ�, .����Ӧ�и�, .ͼƬ, .ϵͳ, .��ID, .SubIDs, .CopyIDs, "_" & .id, .����Դ, _
                                .���¼��, .���Ҽ��, .Դ�к�, .�������, .�������
                        End With
                    Next
                End If
                If tmpItem.���� = "14" Then
                    For Each tmpItem1 In objReport.Items
                        If tmpItem1.��ID = tmpItem.id Then
                            With tmpItem1
                                objClip.Add .id, .��ʽ��, .����, .�ϼ�ID, .����, .���, .����, .����, .����, .��ͷ, .X, .Y, .W, .H, _
                                    .�и�, .����, .�Ե�, .����, .�ֺ�, .����, .����, .б��, .����, .ǰ��, .����, .�߿�, .����, .����, _
                                    .��ʽ, .����, .����߼Ӵ�, .����Ӧ�и�, .ͼƬ, .ϵͳ, .��ID, .SubIDs, .CopyIDs, "_" & .id, .����Դ, _
                                    .���¼��, .���Ҽ��, .Դ�к�, .�������, .�������
                            End With
                        End If
                    Next
                End If
                '�������
                Set tmpItem.CopyIDs = New RelatIDs
            End With
        End If
    Next
    
    Call SelClear
End Sub

Private Sub mnuEdit_Del_Click()
    Dim strKey As String, tmpItem As RPTItem
    Dim blnDo As Boolean, tmpMain As RPTItem, tmpID As RelatID
    
    If tvwSQL.Nodes.count = 1 Then
        MsgBox "��ǰû������Դ����ɾ����", vbInformation, App.Title: Exit Sub
    End If
    If tvwSQL.SelectedItem.Key = "Root" Then
        MsgBox "��ѡ��Ҫɾ��������Դ��", vbInformation, App.Title: Exit Sub
    End If
    
    If tvwSQL.SelectedItem.Parent.Key <> "Root" Then
        strKey = tvwSQL.SelectedItem.Parent.Key
    Else
        strKey = tvwSQL.SelectedItem.Key
    End If
    
    '��鱨��Ԫ�����Ƿ�ʹ���˸�����Դ,������ɾ��
    For Each tmpItem In objReport.Items
        If tmpItem.���� = 5 And tmpItem.���� = mdlPublic.GetStdNodeText(tvwSQL.Nodes(strKey).Text) Then  '���ܱ������
            MsgBox "�ڱ����з����л��ܱ��ʹ���˸�����Դ,����ɾ����", vbInformation, App.Title: Exit Sub
        ElseIf tmpItem.���� = 6 And InStr(tmpItem.����, mdlPublic.GetStdNodeText(tvwSQL.Nodes(strKey).Text) & ".") > 0 Then  '���ⱨ����
            MsgBox "�ڱ����з�����������ʹ���˸�����Դ��������,����ɾ����", vbInformation, App.Title: Exit Sub
        ElseIf tmpItem.���� = 3 And InStr(tmpItem.����, mdlPublic.GetStdNodeText(tvwSQL.Nodes(strKey).Text) & ".") > 0 Then  '���ݱ�ǩ
            MsgBox "�ڱ����з��������ݱ�ǩʹ���˸�����Դ��������,����ɾ����", vbInformation, App.Title: Exit Sub
        ElseIf tmpItem.���� = 12 And InStr(tmpItem.����, mdlPublic.GetStdNodeText(tvwSQL.Nodes(strKey).Text) & ".") > 0 Then  'ͼ��
            MsgBox "�ڱ����з���������ͼ��ʹ���˸�����Դ��������,����ɾ����", vbInformation, App.Title: Exit Sub
        End If
    Next
    
    If MsgBox("ȷʵҪɾ������Դ " & tvwSQL.Nodes(strKey).Text & " ��", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then Exit Sub
    
    On Error Resume Next
    If objClip.count > 0 Then
        For Each tmpItem In objClip
            If tmpItem.���� = 5 And tmpItem.���� = mdlPublic.GetStdNodeText(tvwSQL.Nodes(strKey).Text) Then  '���ܱ������
                For Each tmpID In tmpItem.SubIDs
                    objClip.Remove "_" & tmpID.id
                Next
                objClip.Remove "_" & tmpItem.id
                blnDo = True
            ElseIf tmpItem.���� = 6 And InStr(tmpItem.����, mdlPublic.GetStdNodeText(tvwSQL.Nodes(strKey).Text) & ".") > 0 Then  '���ⱨ����
                Set tmpMain = objClip("_" & tmpItem.�ϼ�ID)
                For Each tmpID In tmpMain.SubIDs
                    objClip.Remove "_" & tmpID.id
                Next
                objClip.Remove "_" & tmpMain.id
                blnDo = True
            ElseIf tmpItem.���� = 3 And InStr(tmpItem.����, mdlPublic.GetStdNodeText(tvwSQL.Nodes(strKey).Text) & ".") > 0 Then  '���ݱ�ǩ
                objClip.Remove "_" & tmpItem.id
                blnDo = True
            ElseIf tmpItem.���� = 12 And InStr(tmpItem.����, mdlPublic.GetStdNodeText(tvwSQL.Nodes(strKey).Text) & ".") > 0 Then  'ͼ��
                objClip.Remove "_" & tmpItem.id
                blnDo = True
            End If
        Next
    End If
    On Error GoTo 0
        
    If blnDo Then MsgBox "ϵͳ�������з�����ʹ���˸�����Դ��Ԫ�أ��Ѿ��Զ������ЩԪ�أ�", vbInformation, App.Title
    
    objReport.Datas.Remove tvwSQL.Nodes(strKey).Key
    tvwSQL.Nodes.Remove tvwSQL.Nodes(strKey).Key
    
    tvwSQL.Nodes(1).Selected = True
    tvwSQL_NodeClick tvwSQL.SelectedItem
    
    BlnSave = False
End Sub

Private Sub mnuEdit_DelFormat_Click()
    cmdDel_Click
End Sub

Private Sub mnuFile_Guide_Click()
    Dim tmpItem As RPTItem, tmpData As RPTData, tmpID As RelatID
    
    Set frmGuide.frmParent = Me
    Set frmGuide.objReport = objReport
    Set frmGuide.mobjFmt = objReport.Fmts("_" & mbytCurrFmt)
    frmGuide.Show 1, Me
    
    If gblnOK Then
        If frmGuide.objGuide.��� = "" Then
            '�����������
            Set objReport.Items = New RPTItems
            Set objReport.Datas = New RPTDatas
            Set objReport.Fmts = New RPTFmts
            Set objReport.Items = frmGuide.objGuide.Items
            Set objReport.Datas = frmGuide.objGuide.Datas
            Set objReport.Fmts = frmGuide.objGuide.Fmts
            mbytCurrFmt = 1
        Else
            '������������
            For Each tmpData In frmGuide.objGuide.Datas
                With tmpData
                    objReport.Datas.Add .����, .�������ӱ��, .SQL, .�ֶ�, .����, .����, .˵��, .Pars, "_" & .����
                End With
            Next
            For Each tmpItem In frmGuide.objGuide.Items
                With tmpItem
                    objReport.Items.Add .id, mbytCurrFmt, .����, .�ϼ�ID, .����, .���, .����, .����, .����, .��ͷ _
                        , .X, .Y, .W, .H, .�и�, .����, .�Ե�, .����, .�ֺ�, .����, .����, .б��, .����, .ǰ��, .���� _
                        , .�߿�, .����, .����, .��ʽ, .����, .����߼Ӵ�, .����Ӧ�и�, .ͼƬ, .ϵͳ, .��ID, .SubIDs _
                        , .CopyIDs, "_" & .id, .����Դ, .���¼��, .���Ҽ��, .Դ�к�, .�������, .�������
                End With
            Next
        End If
        
        Unload frmGuide
        
        '���¼������ؼ�����
        intMaxID = 0
        For Each tmpItem In objReport.Items
            If tmpItem.id > intMaxID Then intMaxID = tmpItem.id
            'ע�����
            For Each tmpID In tmpItem.CopyIDs
                If tmpID.id > intMaxID Then intMaxID = tmpID.id
            Next
        Next
        
        Set objLastSel = Nothing: intCurID = 0
        
        Call ReFlashReport
        Call LoadReportFormat
        BlnSave = False
    End If
End Sub

Private Sub mnuEdit_Inverse_Click()
'���ܣ�����ѡ�񱨱�Ԫ�ؿؼ�
    Dim tmpItem As RPTItem, ObjSel As Object
    Me.MousePointer = 11
    For Each tmpItem In objReport.Items
        If tmpItem.��ʽ�� = Mid(cboFormat.ComboItems("_" & mbytCurrFmt).Key, 2) Then
            If InStr(1, "|1,|2,|3,|4,|5,|10,|11,|12,|14,", "|" & tmpItem.����) <> 0 Then
                Set ObjSel = GetInxObj(tmpItem.id)
                
                If ObjSel.Tag = "" Then
                    SelItem tmpItem.id, True
                Else
                    SelItem tmpItem.id, False
                End If
            End If
        End If
    Next
    Call ShowAttrib
    If GetSelNum = 1 Then
        Call ShowAttrib(intCurID)
    End If
    Me.MousePointer = 0
End Sub

Private Sub mnuEdit_ItemAdd_Click(Index As Integer)
    Dim curNode As Object, objNode As Object
    Dim i As Integer, j As Integer
    
    '���⶯��
    selArea.Left = 500: selArea.Top = 300
    Select Case Index
        Case 0 '����
            bytCurTool = 1
            selArea.Bottom = selArea.Top
            selArea.Right = IIF(picPaper.Width < picBack.Width, picPaper.Width, picBack.Width) - selArea.Left
        Case 1, 3 '����,��ǩ
            bytCurTool = IIF(Index = 1, 2, Index)
            selArea.Bottom = selArea.Top + lbl(0).Height
            selArea.Right = selArea.Left + 2000
            If Index = 1 Then selArea.Bottom = selArea.Bottom + 1000
        Case 4
            bytCurTool = Index
            selArea.Right = selArea.Left + 1500
            selArea.Bottom = selArea.Top + 1300
        Case 5 '���
            bytCurTool = Index
            selArea.Right = selArea.Left + 5000
            selArea.Bottom = selArea.Top + 5000
        Case 6 'ͼ��@@@
            bytCurTool = Index
            selArea.Right = selArea.Left + 4500
            selArea.Bottom = selArea.Top + 3000
        Case 7 '����
            bytCurTool = Index
            selArea.Right = selArea.Left + 2500
            selArea.Bottom = selArea.Top + 1100
    End Select
    Call AddReportItem
    bytCurTool = 0
    
    BlnSave = False
End Sub

Private Sub mnuEdit_New_Click()
    Dim objData As RPTData
    If frmSQLEdit.ShowMe(Me, IIF(glngSys <> 0, glngSys, objReport.ϵͳ), objData, objReport.Datas, 0) Then
        With objData
            objReport.Datas.Add .����, .�������ӱ��, .SQL, .�ֶ�, .����, .����, .˵��, .Pars, "_" & .Key
        End With
        Call tvwSQL_NodeClick(tvwSQL.SelectedItem)
        Unload frmSQLEdit
        BlnSave = False
    End If
End Sub

Private Sub mnuEdit_Modi_Click()
    Dim strKey As String, strInfo As String
    Dim strPreName As String, strNewName As String
    Dim strDBName As String  '���ݿ��е�����
    Dim objData As RPTData
    
    If tvwSQL.Nodes.count = 1 Then
        MsgBox "��ǰû������Դ�����޸�,������������Դ��", vbInformation, App.Title: Exit Sub
    End If
    If tvwSQL.SelectedItem.Key = "Root" Then
        MsgBox "��ѡ��Ҫ�޸ĵ�����Դ��", vbInformation, App.Title: Exit Sub
    End If
    
    If tvwSQL.SelectedItem.Parent.Key <> "Root" Then
        strKey = tvwSQL.SelectedItem.Parent.Key
    Else
        strKey = tvwSQL.SelectedItem.Key
    End If
    strPreName = objReport.Datas(strKey).����
    
    Set objData = objReport.Datas(strKey)
    strDBName = objReport.Datas(strKey).ԭ����
    
    If frmSQLEdit.ShowMe(Me, IIF(glngSys <> 0, glngSys, objReport.ϵͳ), objData, objReport.Datas, 0) Then
        objReport.Datas.Remove strKey
        With objData
            objReport.Datas.Add .����, .�������ӱ��, .SQL, .�ֶ�, .����, .����, .˵��, .Pars, "_" & .Key
            strNewName = .����
        End With
        Call tvwSQL_NodeClick(tvwSQL.SelectedItem)
        
        '�������Դ���Ƹ���,������漰�ı���Ԫ�ص���Ӧ����(�ֶ���������޷�����)
        If strPreName <> strNewName Then
            Call ReplaceName(strPreName, strNewName)
            '����ǵ�һ���޸ģ����ԭ���Ƹ�ֵ
            If strDBName = "" Then
                objReport.Datas("_" & objData.Key).ԭ���� = strPreName
            Else
                objReport.Datas("_" & objData.Key).ԭ���� = strDBName
            End If
            If GetSelNum = 1 Then Call ShowAttrib(intCurID)
        End If
        
        '������ݶ�Ӧ��ϵ
        Me.Refresh
        strInfo = CheckData
        If strInfo <> "" Then
            MsgBox strInfo, vbInformation, App.Title
        End If
        Unload frmSQLEdit
        BlnSave = False
    End If
End Sub

Private Sub mnuEdit_Paste_Click()
    Dim tmpCopy As RPTItem, tmpItem As RPTItem, tmpID As RelatID
    Dim i As Integer, j As Integer, strName As String, tmpChange As RPTItem
    Dim RectTest As RECT
    Dim Col As New Collection
    Dim objClipTmp As RPTItems '���������
    Dim lng��ID As Integer, lng��IDTmp As Long
    Dim blnSouse As Boolean
    Dim k As Long, X As Long, Y As Long
    Dim strSouse As String
    Dim lngMinusX As Long, lngMinusY As Long, strCardIDs As String
    
    If objClip.count = 0 Then PlayWarn: Exit Sub
    On Error Resume Next
    If lblSize.count = 9 Then
        If objReport.Items("_" & Val(lblSize(1).Tag)).���� = "14" Then
            lng��IDTmp = objReport.Items("_" & Val(lblSize(1).Tag)).id
        End If
    End If
    Err.Clear: On Error GoTo 0
    Call SelClear
    
    'ֻճ��һ���ؼ�����������ն���
    If objClip.count = 1 Then objClip(1).���� = "": objClip(1).���� = 0
    
    'ճ������ؼ�,������,��������ռ�����
    Call CheckClip
    '���򣺷���Ԫ�ص��ȼ���
    Set objClipTmp = New RPTItems
    For j = 1 To objClip.count
        Set tmpCopy = objClip(j)
        If tmpCopy.���� = 5 Then
            
        End If
        If tmpCopy.��ID = 0 Then
            With tmpCopy
                objClipTmp.Add .id, .��ʽ��, .����, .�ϼ�ID, .����, .���, .����, .����, .����, .��ͷ, .X, .Y, .W, .H, _
                .�и�, .����, .�Ե�, .����, .�ֺ�, .����, .����, .б��, .����, .ǰ��, .����, .�߿�, .����, .����, .��ʽ, _
                .����, .����߼Ӵ�, .����Ӧ�и�, .ͼƬ, .ϵͳ, .��ID, .SubIDs, .CopyIDs, "_" & .id, .����Դ, .���¼��, _
                .���Ҽ��, .Դ�к�, .�������, .�������
                If .���� = 14 Then strCardIDs = strCardIDs & "," & .id
            End With
        End If
        If lng��IDTmp <> 0 And tmpCopy.�ϼ�ID = 0 Then
            If tmpCopy.���� = 14 Then
                MsgBox "��ƬԪ�ز������Ƶ���ƬԪ����ȥ��", vbInformation, App.Title
                Exit Sub
            ElseIf tmpCopy.���� = 4 And tmpCopy.���� > 1 Then
                MsgBox "��Ƭ�в������������ı��", vbInformation, App.Title
                Exit Sub
            End If
            If lngMinusX > tmpCopy.X Or lngMinusX = 0 Then lngMinusX = tmpCopy.X
            If lngMinusY > tmpCopy.Y Or lngMinusY = 0 Then lngMinusY = tmpCopy.Y
        End If
    Next
    strCardIDs = Mid(strCardIDs, 2)
    For j = 1 To objClip.count
        Set tmpCopy = objClip(j)
        If tmpCopy.��ID <> 0 Then
            With tmpCopy
                If lng��IDTmp = 0 Then
                    If InStr("," & strCardIDs & ",", "," & tmpCopy.��ID & ",") = 0 Then
                        .X = .X + pic(.��ID).Left
                        .Y = .Y + pic(.��ID).Top
                    Else
                        .X = .X - 200  '����Ҫ�����٣����Ա��ֲ���
                        .Y = .Y - 200
                    End If
                End If
                objClipTmp.Add .id, .��ʽ��, .����, .�ϼ�ID, .����, .���, .����, .����, .����, .��ͷ, .X, .Y, .W, .H, _
                    .�и�, .����, .�Ե�, .����, .�ֺ�, .����, .����, .б��, .����, .ǰ��, .����, .�߿�, .����, .����, _
                    .��ʽ, .����, .����߼Ӵ�, .����Ӧ�и�, .ͼƬ, .ϵͳ, .��ID, .SubIDs, .CopyIDs, "_" & .id, .����Դ, _
                    .���¼��, .���Ҽ��, .Դ�к�, .�������, .�������
            End With
        End If
    Next
    Set objClip = objClipTmp
    For j = 1 To objClip.count
        '���ڲ��ն����Ԫ��,��������ı�����,��ѭ���ı����ӱ�
        Set tmpCopy = objClip(j)
        If InStr(1, "|1,|2,|3,|4,|5,|10,|11,|12,|14,", "|" & tmpCopy.���� & ",") <> 0 Then
            strName = GetNextName(tmpCopy.����, True)
            If strName <> tmpCopy.���� Then
                For Each tmpChange In objClip
                    If tmpChange.���� = tmpCopy.���� And InStr(1, "4,5", tmpChange.����) > 0 Then tmpChange.���� = strName
                    If tmpChange.���� Like "��ǩ*" And tmpChange.���� = 2 Then tmpChange.���� = strName
                Next
                tmpCopy.���� = strName
            End If
        End If
    Next
    
    For j = 1 To objClip.count
        '���ڲ��ն����Ԫ��,��������ı�����,��ѭ���ı����ӱ�
        Set tmpCopy = objClip(j)
        blnSouse = False
        If tmpCopy.���� = "4" And lng��IDTmp <> 0 Then
            If objReport.Items("_" & lng��IDTmp).����Դ <> "" Then
                For Each tmpID In tmpCopy.SubIDs
                    With objReport.Items("_" & tmpID.id)
                        X = InStr(1, .����, "]")
                        Y = InStr(1, .����, ".")
                        k = InStr(1, .����, "[")
                        If X > k And X > Y And X <> 0 And k <> 0 And Y <> 0 Then
                            If Mid(.����, k + 1, Y - k - 1) <> objReport.Items("_" & lng��IDTmp).����Դ Then
                                strSouse = strSouse & "," & tmpCopy.����
                                blnSouse = True
                                Exit For
                            End If
                        End If
                    End With
                Next
            End If
        End If
        If blnSouse = False Then
            If InStr(1, "|1,|2,|3,|4,|5,|10,|11,|12,|13,|14,", "|" & tmpCopy.���� & ",") <> 0 Then
                intMaxID = intMaxID + 1
                With tmpCopy
                    RectTest.Left = .X
                    RectTest.Top = .Y
                    RectTest.Right = .W
                    RectTest.Bottom = .H
                    lng��ID = 0
                    If .��ID <> 0 Then
                        On Error Resume Next
                        lng��ID = Val(Col("_" & .��ID) & "")
                        Err.Clear: On Error GoTo 0
                    End If
                    If lng��IDTmp <> 0 And InStr(",5,12,14,", "," & tmpCopy.���� & ",") = 0 Then
                        lng��ID = lng��IDTmp
                    End If
                    If .��ʽ�� = mbytCurrFmt Then
                        If .���� = 2 Or .���� = 12 Then '@@@
                            '�ڸ߶��ϱ仯,Ϊ������յ�X���귢���ı�
                            If .���� <> "" Then
                                Select Case Mid(.����, 1, 1)
                                Case 1  '������
                                    If Not ((.Y - .H - 200) * sgnMode < 100) Then RectTest.Top = RectTest.Top - 200
                                    If .ϵͳ Then Call GetCoordinate(RectTest)
                                    Set tmpItem = objReport.Items.Add(intMaxID, mbytCurrFmt, .����, 0, .����, .���, _
                                        .����, .����, .����, .��ͷ, RectTest.Left, RectTest.Top, RectTest.Right, RectTest.Bottom, _
                                        .�и�, .����, .�Ե�, .����, .�ֺ�, .����, .����, .б��, .����, .ǰ��, .����, .�߿�, .����, _
                                        .����, .��ʽ, .����, .����߼Ӵ�, .����Ӧ�и�, .ͼƬ, .ϵͳ, lng��ID, , , "_" & intMaxID, .����Դ, _
                                        .���¼��, .���Ҽ��, .Դ�к�, .�������, .�������)
                                Case Else
                                    If (.Y + .H + 200) * sgnMode < picPaper.Height - 100 Then
                                        RectTest.Top = RectTest.Top + RectTest.Bottom + 200
                                    End If
                                    If .ϵͳ Then Call GetCoordinate(RectTest)
                                    Set tmpItem = objReport.Items.Add(intMaxID, mbytCurrFmt, .����, 0, .����, .���, _
                                        .����, .����, .����, .��ͷ, RectTest.Left, RectTest.Top, RectTest.Right, RectTest.Bottom, _
                                        .�и�, .����, .�Ե�, .����, .�ֺ�, .����, .����, .б��, .����, .ǰ��, .����, .�߿�, .����, _
                                        .����, .��ʽ, .����, .����߼Ӵ�, .����Ӧ�и�, .ͼƬ, .ϵͳ, lng��ID, , , "_" & intMaxID, .����Դ, _
                                        .���¼��, .���Ҽ��, .Դ�к�, .�������, .�������)
                                End Select
                            Else
                                If .ϵͳ Then
                                    Call GetCoordinate(RectTest)
                                Else
                                    RectTest.Top = RectTest.Top + 200
                                    RectTest.Left = RectTest.Left + 200
                                End If
                                Set tmpItem = objReport.Items.Add(intMaxID, mbytCurrFmt, .����, 0, .����, .���, _
                                    .����, .����, .����, .��ͷ, RectTest.Left, RectTest.Top, RectTest.Right, RectTest.Bottom, _
                                    .�и�, .����, .�Ե�, .����, .�ֺ�, .����, .����, .б��, .����, .ǰ��, .����, .�߿�, .����, _
                                    .����, .��ʽ, .����, .����߼Ӵ�, .����Ӧ�и�, .ͼƬ, .ϵͳ, lng��ID, , , "_" & intMaxID, .����Դ, _
                                    .���¼��, .���Ҽ��, .Դ�к�, .�������, .�������)
                            End If
                        Else
                            If .ϵͳ Then
                                Call GetCoordinate(RectTest)
                            Else
                                RectTest.Top = RectTest.Top + 200
                                RectTest.Left = RectTest.Left + 200
                            End If
                            Set tmpItem = objReport.Items.Add(intMaxID, mbytCurrFmt, .����, 0, .����, .���, _
                                .����, .����, .����, .��ͷ, RectTest.Left, RectTest.Top, RectTest.Right, RectTest.Bottom, _
                                .�и�, .����, .�Ե�, .����, .�ֺ�, .����, .����, .б��, .����, .ǰ��, .����, .�߿�, .����, _
                                .����, .��ʽ, .����, .����߼Ӵ�, .����Ӧ�и�, .ͼƬ, .ϵͳ, lng��ID, , , "_" & intMaxID, .����Դ, _
                                .���¼��, .���Ҽ��, .Դ�к�, .�������, .�������)
                        End If
                    Else
                        '������ԭ��ʽ�µ�����һ��
                        Set tmpItem = objReport.Items.Add(intMaxID, mbytCurrFmt, .����, 0, .����, .���, _
                            .����, .����, .����, .��ͷ, RectTest.Left, RectTest.Top, RectTest.Right, RectTest.Bottom, _
                            .�и�, .����, .�Ե�, .����, .�ֺ�, .����, .����, .б��, .����, .ǰ��, .����, .�߿�, .����, _
                            .����, .��ʽ, .����, .����߼Ӵ�, .����Ӧ�и�, .ͼƬ, .ϵͳ, lng��ID, , , "_" & intMaxID, .����Դ, _
                            .���¼��, .���Ҽ��, .Դ�к�, .�������, .�������)
                    End If
                    If .��ID = 0 Then
                        Col.Add intMaxID, "_" & .id
                    End If
                    If lng��ID <> 0 Then
                        tmpItem.X = tmpItem.X - lngMinusX
                        tmpItem.Y = tmpItem.Y - lngMinusY
                    End If
                    '��������
                    If (.���� = 4 Or .���� = 5) And .SubIDs.count > 0 Then
                        For Each tmpID In .SubIDs
                            intMaxID = intMaxID + 1
                            With objClip("_" & tmpID.id)
                                objReport.Items.Add intMaxID, mbytCurrFmt, strName, tmpItem.id, .����, .���, _
                                    .����, .����, .����, .��ͷ, .X + 300, .Y + 300, .W, .H, .�и�, .����, .�Ե�, _
                                    .����, .�ֺ�, .����, .����, .б��, .����, .ǰ��, .����, .�߿�, .����, .����, _
                                    .��ʽ, .����, .����߼Ӵ�, .����Ӧ�и�, .ͼƬ, .ϵͳ, lng��ID, , , "_" & intMaxID, _
                                    .����Դ, .���¼��, .���Ҽ��, .Դ�к�, .�������, .�������
                                tmpItem.SubIDs.Add intMaxID, "_" & intMaxID
                            End With
                        Next
                    End If
                
                    '�������
                    If .���� = 4 And .���� > 1 Then
                        For i = 1 To .���� - 1
                            intMaxID = intMaxID + 1
                            tmpItem.CopyIDs.Add intMaxID, "_" & intMaxID
                        Next
                    End If
                End With
                Call ShowItem(tmpItem.id)
                Call SelItem(tmpItem.id, True)
            End If
        End If
    Next
    
    If strSouse <> "" Then
        MsgBox "���" & Mid(strSouse, 2) & " �е��а󶨵����ݲ��Ǳ���Ƭ����Դ�е����ݣ����ܷ���˿�Ƭ�С�", vbInformation, App.Title
    End If
    
    If objClip.count = 1 Then
        Call ShowAttrib(tmpItem.id)
    End If
    Set objClip = New RPTItems
    
    BlnSave = False
End Sub

Private Sub mnuEdit_Remove_Click()
'���ܣ�ɾ����ǰѡ��ı���Ԫ��(һ������)
    Dim tmpObj As PictureBox, tmpID As RelatID, ItemThis As RPTItem
    Dim objControl As Object, i As Long
    Dim tmpObj1 As PictureBox, blntmp As Boolean
    
    Select Case GetSelNum
    Case 0
        MsgBox "û��ѡ���κα���Ԫ��,�޷�ɾ����", vbInformation, App.Title: Exit Sub
    Case 1
        If objReport.Items("_" & intCurID).ϵͳ Then
            MsgBox "��ǰѡ��ı���Ԫ����ϵͳ����Ԫ��,�޷�ɾ����", vbInformation, App.Title: Exit Sub
        End If
    End Select
    
    If MsgBox("��ѡ���� " & GetSelNum & " ��Ԫ��,ȷʵҪɾ����", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then Exit Sub
Res:
    For Each tmpObj In lblSize
        If lblSize.count > 1 Then
            On Error Resume Next
            If tmpObj.Index Mod 8 = 1 Then
                If blntmp = True Then blntmp = False: GoTo Res
                    If Not objReport.Items("_" & tmpObj.Tag).ϵͳ Then
                        Select Case objReport.Items("_" & tmpObj.Tag).����
                            Case 1
                                Unload lblLine(tmpObj.Tag)
                                objReport.Items.Remove "_" & tmpObj.Tag
                            Case 2, 3
                                Unload lbl(tmpObj.Tag)
                                objReport.Items.Remove "_" & tmpObj.Tag
                            Case 10
                                Unload Shp(tmpObj.Tag)
                                Unload lblshp(tmpObj.Tag)
                                objReport.Items.Remove "_" & tmpObj.Tag
                            Case 11
                                Unload img(tmpObj.Tag)
                                objReport.Items.Remove "_" & tmpObj.Tag
                            Case 14
                                For Each objControl In Me.Controls
                                     If InStr(";ImageList;CommonDialog;Menu;PictureBox;", ";" & TypeName(objControl) & ";") = 0 Then
                                        If objControl.Container Is pic(tmpObj.Tag) Then
                                            'ɾ���ӿؼ���λ�õ㣬��ɾ���ؼ���Ȼ������ѭ��
                                            For Each tmpObj1 In lblSize
                                                If tmpObj1.Index Mod 8 = 1 Then
                                                    If objReport.Items("_" & tmpObj1.Tag).id = objControl.Index Then
                                                        For i = tmpObj1.Index To tmpObj1.Index + 7
                                                            Unload lblSize(i)
                                                        Next
                                                        blntmp = True
                                                        Exit For
                                                    End If
                                                End If
                                            Next
                                            If objReport.Items("_" & objControl.Index).���� = 4 Then
                                                '�Ƴ��������
                                                For Each tmpID In objReport.Items("_" & objControl.Index).SubIDs
                                                    objReport.Items.Remove "_" & tmpID.id
                                                Next
                                            End If
                                            objReport.Items.Remove "_" & objControl.Index
                                            Unload objControl
                                        End If
                                    End If
                                Next
                                Unload pic(tmpObj.Tag)
                                objReport.Items.Remove "_" & tmpObj.Tag
                            Case 4, 5
                                '���ɾ������һ���ӱ����ӱ��������ӱ�񣩣�����������ӱ��λ��
                                If objReport.Items("_" & tmpObj.Tag).���� <> "" Then
                                    For Each ItemThis In objReport.Items
                                        If ItemThis.��ʽ�� = mbytCurrFmt And ItemThis.���� = objReport.Items("_" & tmpObj.Tag).���� Then
                                            If InStr(1, "4,5", ItemThis.����) <> 0 Then objReport.Items("_" & tmpObj.Tag).���� = "": SetChildWH (ItemThis.Key)
                                        End If
                                    Next
                                Else
                                    For Each ItemThis In objReport.Items
                                        If ItemThis.��ʽ�� = mbytCurrFmt And ItemThis.���� = objReport.Items("_" & tmpObj.Tag).���� Then ItemThis.���� = "": ItemThis.���� = 0
                                    Next
                                End If
                                
                                Unload msh(tmpObj.Tag)
                                '�Ƴ������ؼ�
                                For Each tmpID In objReport.Items("_" & tmpObj.Tag).CopyIDs
                                    Unload msh(tmpID.id)
                                Next
                                '�Ƴ��������
                                For Each tmpID In objReport.Items("_" & tmpObj.Tag).SubIDs
                                    objReport.Items.Remove "_" & tmpID.id
                                Next
                                objReport.Items.Remove "_" & tmpObj.Tag
                            Case 12 '@@@
                                Unload Chart(tmpObj.Tag)
                                objReport.Items.Remove "_" & tmpObj.Tag
                            Case 13
                                Unload ImgCode(tmpObj.Tag)
                                objReport.Items.Remove "_" & tmpObj.Tag
                        End Select
                    Else
                        Select Case objReport.Items("_" & tmpObj.Tag).����
                            Case 1
                                lblLine(tmpObj.Tag).Tag = ""
                            Case 2, 3
                                lbl(tmpObj.Tag).Tag = ""
                            Case 10
                                Shp(tmpObj.Tag).Tag = ""
                            Case 11
                                img(tmpObj.Tag).Tag = ""
                            Case 4, 5
                                msh(tmpObj.Tag).Tag = ""
                            Case 12 '@@@
                                Chart(tmpObj.Tag).Tag = ""
                            Case 13
                                ImgCode(tmpObj.Tag).Tag = ""
                            Case 14
                                pic(tmpObj.Tag).Tag = ""
                        End Select
                    End If
            End If
            If tmpObj.Index <> 0 Then
                Unload lblSize(tmpObj.Index)
            End If
        End If
    Next
    
    Set objLastSel = Nothing: intCurID = 0
    
    picPaper.SetFocus
    Call picPaper_MouseDown(1, 0, 0, 0)
    BlnSave = False
End Sub

Private Function CheckPars() As String
'���ܣ��������е�����Դ�󶨲���������Ƿ���ȷ
    Dim objData As RPTData, objPar As RPTPar
    Dim strSQL As String, strParName As String
    
    For Each objData In objReport.Datas
        For Each objPar In objData.Pars
            If objPar.��ϸSQL <> "" Then
                strSQL = objPar.��ϸSQL
                If CheckParsRela(strSQL, objReport.Datas, objPar.����, , , , strParName) = False Then
                    CheckPars = "����Դ[" & objData.���� & "]�еĲ���[" & objPar.���� & "]����ϸSQL�а󶨵Ĳ���[" & strParName & "]δ�����󶨵ľ��ǵ�ǰ���������顣"
                    Exit Function
                End If
            End If
            If objPar.����SQL <> "" Then
                strSQL = objPar.����SQL
                If CheckParsRela(strSQL, objReport.Datas, objPar.����, , , , strParName) = False Then
                    CheckPars = "����Դ[" & objData.���� & "]�еĲ���[" & objPar.���� & "]�ķ���SQL�а󶨵Ĳ���[" & strParName & "]δ�����󶨵ľ��ǵ�ǰ���������顣"
                    Exit Function
                End If
            End If
        Next
    Next
    
End Function

Private Sub mnuFile_Save_Click()
    Dim strInfo As String, LngItemKey As Long, i As Integer
    
    For i = 1 To cboFormat.ComboItems.count
        If InStr(cboFormat.ComboItems(i).Text, "'") > 0 Then
            MsgBox "�� " & i & " �������ʽ���������˷Ƿ��ַ������飡", vbInformation, App.Title
            Exit Sub
        End If
    Next
    
    '�������Դ
    strInfo = CheckData
    If strInfo <> "" Then MsgBox strInfo, vbInformation, App.Title: Exit Sub
    
    '����ͷ����
    strInfo = CheckHead
    If strInfo <> "" Then MsgBox strInfo, vbInformation, App.Title: Exit Sub
    
    strInfo = CheckArea
    If strInfo <> "" Then MsgBox strInfo, vbInformation, App.Title: Exit Sub
    
    strInfo = CheckPars
    If strInfo <> "" Then MsgBox strInfo, vbInformation, App.Title: Exit Sub
    
    LngItemKey = CheckCoordinate
    If LngItemKey <> 0 Then
        '���ϵͳ���Ƿ���ס������������
        MsgBox "ϵͳ����Ԫ��[" & objReport.Items("_" & LngItemKey).���� & "]������Ԫ����ס��������ֹ��", vbInformation, App.Title
        Exit Sub
    End If
    
    Call SelClear
    Refresh
    If Not SaveReport(lngRPTID, objReport, sta.Panels(2)) Then
        MsgBox "������ʧ��,�����Ա��������", vbInformation, App.Title
        Exit Sub
    End If
    Call UpdatePriv
    
    gblnModi = True
    BlnSave = True
    
    If Not CheckReportPriv(lngRPTID) Then
        MsgBox "��û��Ȩ�޲�ѯ�ñ���ĳЩ����Դ�еĶ�����Ȼ��������" & vbCrLf & _
               "�ر��棬������������Щ����֮ǰ�㲻������ʹ�øñ���", vbInformation, App.Title
    End If
End Sub

Private Sub mnuEdit_SelAll_Click()
'���ܣ�ѡ��ȫ������Ԫ�ؿؼ�
    Dim tmpItem As RPTItem
    
    Me.MousePointer = 11
    For Each tmpItem In objReport.Items
        If tmpItem.��ʽ�� = Mid(cboFormat.ComboItems("_" & mbytCurrFmt).Key, 2) Then
            If InStr(1, "|1,|2,|3,|4,|5,|10,|11,|12,|13,|14,", "|" & tmpItem.����) <> 0 Then '@@@
                SelItem tmpItem.id, True
            End If
        End If
    Next
    Call ShowAttrib
    If GetSelNum = 1 Then
        Call ShowAttrib(intCurID)
    End If
    Me.MousePointer = 0
End Sub

Private Sub mnuFile_Page_Click()
    If Printers.count = 0 Then
        MsgBox "��ϵͳ��û�м�⵽�κδ�ӡ�豸,���Ȱ�װ��ӡ���������Ըò�����" & vbCrLf & _
            "��������µĴ�ӡ��֮ǰ,ϵͳ����ȱʡֽ�Ž������á�", vbInformation, App.Title
        Exit Sub
    End If
    
    With frmPageSetup
        .strPrinter = objReport.��ӡ��
        .intBin = objReport.��ֽ
        .intPage = objReport.Fmts("_" & mbytCurrFmt).ֽ��
        .lngWidth = objReport.Fmts("_" & mbytCurrFmt).W
        .lngHeight = objReport.Fmts("_" & mbytCurrFmt).H
        .bytOrient = objReport.Fmts("_" & mbytCurrFmt).ֽ��
    End With
    frmPageSetup.Show 1, Me
    If gblnOK Then
        With frmPageSetup
            objReport.��ӡ�� = .strPrinter
            objReport.��ֽ = .intBin
            objReport.Fmts("_" & mbytCurrFmt).ֽ�� = .intPage
            objReport.Fmts("_" & mbytCurrFmt).W = .lngWidth '�����Զ���ֽ��ʱ,ҲҪ��ֵ
            objReport.Fmts("_" & mbytCurrFmt).H = .lngHeight
            objReport.Fmts("_" & mbytCurrFmt).ֽ�� = .bytOrient
            If objReport.Fmts("_" & mbytCurrFmt).ֽ�� = 2 Then
                objReport.Fmts("_" & mbytCurrFmt).��ֽ̬�� = False
            End If
        End With
        Unload frmPageSetup
        Call ShowSize: Call ShowScroll: Call GetInPaper
        BlnSave = False
    End If
End Sub

Private Sub mnuFile_Quit_Click()
    Unload Me
End Sub

Private Sub mnuFile_Report_Click()
    Dim strInfo As String, LngItemKey As Long

    strInfo = CheckData
    If strInfo <> "" Then MsgBox strInfo, vbInformation, App.Title: Exit Sub
    
    strInfo = CheckHead
    If strInfo <> "" Then MsgBox strInfo, vbInformation, App.Title: Exit Sub
    
    strInfo = CheckArea
    If strInfo <> "" Then MsgBox strInfo, vbInformation, App.Title: Exit Sub
    
    strInfo = CheckPars
    If strInfo <> "" Then MsgBox strInfo, vbInformation, App.Title: Exit Sub
    
    LngItemKey = CheckCoordinate
    If LngItemKey <> 0 Then
        '���ϵͳ���Ƿ���ס������������
        MsgBox "ϵͳ����Ԫ��[" & objReport.Items("_" & LngItemKey).���� & "]������Ԫ����ס��������ֹ��", vbInformation, App.Title
        Exit Sub
    End If
    
    If cboFormat.SelectedItem Is Nothing Then
        MsgBox "����ȷ����ǰ�����ʽ,��ѡ��һ�ֱ����ʽ��", vbInformation, App.Title
        cboFormat.SetFocus: Exit Sub
    End If
    
    'ִ�б���
    glngGroup = 0
    CopyReport objReport, gobjReport
    garrPars = Array("ReportFormat=" & cboFormat.SelectedItem.Index) 'ǿ��ʹ�õ�ǰ��ʽ
    If Not ShowReport(Me) Then MsgBox "�����ʧ�ܣ�", vbInformation, App.Title
End Sub

Private Sub mnuFormat_Back_Click()
    Call SetLevel(1)
End Sub

Private Sub mnuFormat_DoAlign_Click(Index As Integer)
    If Index <= 5 Then
        Call SetSelAlign(Index + 1)
    Else
        If Index = 7 Then
            Call SetSelCenter(0)
        ElseIf Index = 8 Then
            Call SetSelCenter(1)
        End If
    End If
End Sub

Private Sub mnuFormat_Front_Click()
    Call SetLevel
End Sub

Private Sub SetLevel(Optional bytOrder As Byte)
'���ܣ���ѡ��˳������ѡ��ؼ���ǰ��˳��
    Dim tmpObj As PictureBox, ObjSel As Object
    
    If GetSelNum = 0 Then Exit Sub
    
    For Each tmpObj In lblSize
        If tmpObj.Index Mod 8 = 1 Then
            Set ObjSel = GetInxObj(tmpObj.Tag)
        End If
    Next
    ObjSel.ZOrder bytOrder
End Sub

Private Sub mnuFormat_Height_Click()
    Call SetSelAlign(8)
End Sub

Private Sub mnuFormat_HscSpace_Add_Click()
    Call SetHscSpace(1)
End Sub

Private Sub mnuFormat_HscSpace_Dec_Click()
    Call SetHscSpace(-1)
End Sub

Private Sub mnuFormat_HscSpace_Same_Click()
    Call SetHscSpace(0)
End Sub

Private Sub mnuFormat_Lock_Click()
    mnuFormat_Lock.Checked = Not mnuFormat_Lock.Checked
    blnLock = Not blnLock
    If mnuFormat_Lock.Checked And tbr2.Buttons("Lock").Value = tbrUnpressed Then
        tbr2.Buttons("Lock").Value = tbrPressed
    ElseIf Not mnuFormat_Lock.Checked And tbr2.Buttons("Lock").Value = tbrPressed Then
        tbr2.Buttons("Lock").Value = tbrUnpressed
    End If
    Call SetLock(blnLock)
End Sub

Private Sub mnuFormat_VscSpace_Add_Click()
    Call SetVscSpace(1)
End Sub

Private Sub mnuFormat_VscSpace_Dec_Click()
    Call SetVscSpace(-1)
End Sub

Private Sub mnuFormat_VscSpace_Same_Click()
    Call SetVscSpace(0)
End Sub

Private Sub mnuFormat_WH_Click()
    Call SetSelAlign(9)
End Sub

Private Sub mnuFormat_Width_Click()
    Call SetSelAlign(7)
End Sub

Private Sub mnuHelpAbout_Click()
    Call ShowAbout(Me)
End Sub

Private Sub mnuHelpTitle_Click()
    Call ShowHelpRpt(Me.hwnd, "design", 0)
End Sub

Private Sub mnuView_reFlash_Click()
    If MsgBox("ȷʵҪˢ����ʾ����������", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then Exit Sub
    Call ReFlashReport(True)
    Call LoadReportFormat
End Sub

Private Sub mnuViewScaleMode_Click(Index As Integer)
    Dim MenuSelect As Menu   '��ʾ����
    Dim tmpItem As RPTItem, ObjSel As Control
    Dim lngRow As Long
    'zyb#Add
    '��������ʾ
    
    'ȡ�ϴε���ʾ����
    For Each MenuSelect In Me.mnuViewScaleMode
        If MenuSelect.Checked Then
            Select Case MenuSelect.Index
            Case 0, 1, 2
                sgnLastMode = GetAutoTest(MenuSelect.Index)
            Case 4
                sgnLastMode = 2
            Case 5
                sgnLastMode = 1
            Case 6
                sgnLastMode = 0.75
            Case 7
                sgnLastMode = 0.5
            Case 8
                sgnLastMode = 0.25
            End Select
        End If
    Next
    
    '���ѡ��
    mnuViewScaleMode(0).Checked = False
    mnuViewScaleMode(1).Checked = False
    mnuViewScaleMode(2).Checked = False
    mnuViewScaleMode(4).Checked = False
    mnuViewScaleMode(5).Checked = False
    mnuViewScaleMode(6).Checked = False
    mnuViewScaleMode(7).Checked = False
    mnuViewScaleMode(8).Checked = False
    
    '��ȡ��ʾ������������Ӧ�˵�
    Select Case Index
    Case 0, 1, 2
        mnuViewScaleMode(Index).Checked = True
        sgnMode = GetAutoTest(Index)
    Case 4  '200%
        mnuViewScaleMode(4).Checked = True
        sgnMode = 2
    Case 5  '100%
        mnuViewScaleMode(5).Checked = True
        sgnMode = 1
    Case 6  '75%
        mnuViewScaleMode(6).Checked = True
        sgnMode = 0.75
    Case 7  '50%
        mnuViewScaleMode(7).Checked = True
        sgnMode = 0.5
    Case 8  '25%
        mnuViewScaleMode(8).Checked = True
        sgnMode = 0.25
    End Select
    
    Call ShowSize
    Call ShowScroll
    Call ReFlashReportBySelFormat
    If Not scrHsc.Enabled Then DrawRuler picRulerH
    If Not scrVsc.Enabled Then DrawRuler picRulerV
    
    If GetSelNum = 1 Then ShowAttrib (intCurID)
End Sub

Private Sub msh_DblClick(Index As Integer)
    If Not (Left(msh(Index).Tag, 2) = "C_") Then
        If objReport.Items("_" & Index).ϵͳ Then Exit Sub
    End If
    strMenu = "DO"
    msh_MouseDown Index, 2, 0, CDbl(lngPreX), CDbl(lngPreY)
    Select Case strMenu
        Case "mnuClass_Data"
            strMenu = ""
            mnuClass_Data_Click
        Case "mnuCustom_Col_Data"
            strMenu = ""
            mnuCustom_Col_Data_Click
        Case "mnuCustom_Head_Text"
            strMenu = ""
            mnuCustom_Head_Text_Click
    End Select
End Sub

Private Sub msh_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim intIdx As Integer, intState As Integer
    Dim tmpID As RelatID, tmpItem As RPTItem
    
    lngPreX = X: lngPreY = Y
    Set mobjMove = Nothing: mlngX = 0: mlngY = 0
    blnHead = False
    
    '�����ؼ�Tag���="C_���ؼ�����"
    If Left(msh(Index).Tag, 2) = "C_" Then
        Call SetGridLine(CInt(Mid(msh(Index).Tag, 3)))
        Call SetCopyGrid(CInt(Mid(msh(Index).Tag, 3)))
        intIdx = CInt(Mid(msh(Index).Tag, 3))
    Else
        Call SetGridLine(Index)
        Call SetCopyGrid(Index)
        intIdx = Index
        If msh(Index).MouseRow < msh(Index).FixedRows Then blnHead = True
    End If
    
    Call ReFlashWidth
    
    If Shift = 2 Then
        If Mid(msh(intIdx).Tag, 1, 2) = "" Then
            Call SelItem(intIdx, True) '��ѡ
            If GetSelNum() = 1 Then
                Call ShowAttrib(intIdx) 'ֻѡ��һ������ʾ����
            Else
                Call ShowAttrib '��ѡʱ����ʾ����
            End If
        Else
            Call SelItem(intIdx, False) '��ѡ
            If GetSelNum() = 1 Then
                Call ShowAttrib(intCurID) 'ֻѡ��һ������ʾ����(ѡ�еĲ�һ���Ǹÿؼ�)
            Else
                Call ShowAttrib '��ѡʱ����ʾ����
            End If
        End If
    Else
        If Mid(msh(intIdx).Tag, 1, 2) = "" Then
            Call SelClear
            Call SelItem(intIdx, True)
        End If
    End If
    '���༭����.ֻѡ��һ����Ч;�����ؼ���Ч
    If Left(msh(Index).Tag, 2) <> "C_" And GetSelNum = 1 Then
        If objReport.Items("_" & intIdx).���� = 4 Then '������༭����
            If Button = 1 Then
                Call ResetColor(intIdx)
                selCell.Col1 = -1: selCell.Col2 = -1: selCell.Row1 = -1: selCell.Row2 = -1: selCell.Row = -1
                Call ShowPaperInfo
                drgCell.Col1 = -1: drgCell.Col2 = -1: drgCell.Row1 = -1: drgCell.Row2 = -1
                intCurCol = -1
            End If
            If msh(intIdx).MouseRow < msh(intIdx).FixedRows Then '��ͷ��Χ
                If Button = 1 Then
                    selCell = GetCellRange(msh(intIdx), msh(intIdx).MouseRow, msh(intIdx).MouseCol)
                    selCell.Row = msh(intIdx).MouseRow
                    '�϶�����ʼֵ
                    drgCell = selCell
                    Call CustomCellColor(intIdx, selCell, True)
                ElseIf Button = 2 Then
                    Call ShowPaperInfo
                    If strMenu = "" Then
                        If Not objReport.Items("_" & Index).ϵͳ Then PopupMenu mnuCustom_Head, 2
                    Else
                        strMenu = "mnuCustom_Head_Text"
                    End If
                End If
            Else '���з�Χ
                intCurCol = msh(intIdx).MouseCol
                If Button = 1 Then
                    Call CustomColColor(intIdx, intCurCol)
                Else
                    Call SetMenuDefault(intIdx, intCurCol)
                    If strMenu = "" Then
                        If Not objReport.Items("_" & Index).ϵͳ Then PopupMenu mnuCustom_Col, 2
                    Else
                        strMenu = "mnuCustom_Col_Data"
                    End If
                End If
            End If
    
            '����Ƿ�����ı�߶Ȼ���
            If blnAdjustRowHeight = False And msh(Index).MouseRow < msh(Index).FixedRows And Button = 1 And objReport.Items("_" & Index).ϵͳ = False Then
                msh(Index).Row = msh(Index).MouseRow
                msh(Index).Col = msh(Index).MouseCol
                If msh(Index).Row < msh(Index).FixedRows Then
                    blnAdjustRowHeight = (Button = 1 And (Y > msh(Index).CellTop + msh(Index).CellHeight - 100 And Y < msh(Index).CellTop + msh(Index).CellHeight + 100))
                End If
            ElseIf blnAdjustColWidth = False And Button = 1 And msh(Index).MouseRow = msh(Index).FixedRows And msh(Index).Col = 0 And objReport.Items("_" & Index).ϵͳ = False Then
                msh(Index).Row = msh(Index).MouseRow
                msh(Index).Col = msh(Index).MouseCol
                If msh(Index).Row = msh(Index).FixedRows And msh(Index).Col = 0 Then
                    blnAdjustColWidth = (Button = 1 And (X > msh(Index).CellLeft + msh(Index).CellWidth - 100 And X < msh(Index).CellLeft + msh(Index).CellWidth + 100))
                End If
                If blnAdjustColWidth Then
                    '��ʾ�ָ���,�������п�
                    msh(Index).MousePointer = 9
                    With PicSplit
                        .Left = X + msh(Index).Left
                        .Top = msh(Index).Top
                        .Height = msh(Index).CellTop + msh(Index).CellHeight
                        .ZOrder 0
                        .Visible = True
                    End With
                Else
                    PicSplit.Visible = False
                End If
            End If
        ElseIf objReport.Items("_" & intIdx).���� = 5 Then '���ܱ��༭����
            If Button = 1 Then
                Call ResetColor(intIdx)
                selCell.Col1 = -1: selCell.Row1 = -1
                Call ShowPaperInfo
            End If
            
            If msh(intIdx).MouseRow < msh(intIdx).FixedRows Then '��ͷ��Χ
                If Button = 1 Then
                    selCell = GetCellRange(msh(intIdx), msh(intIdx).MouseRow, msh(intIdx).MouseCol)
                    selCell.Row = msh(intIdx).MouseRow
                    '�϶�����ʼֵ
                    drgCell = selCell
                    Call CustomCellColor(intIdx, selCell, True)
                End If
            End If
            
            '��ͳ�������
            For Each tmpID In objReport.Items("_" & intIdx).SubIDs
                If objReport.Items("_" & tmpID.id).���� = 9 Then intState = intState + 1
            Next
            If msh(intIdx).MouseCol <= msh(intIdx).FixedCols - 1 And msh(intIdx).MouseRow >= msh(intIdx).FixedRows - 1 Then
                 '������෶Χ
                If Button = 1 Then
                    selCell.Row1 = msh(intIdx).MouseRow
                    selCell.Col1 = msh(intIdx).MouseCol
                    Call ClassColor(intIdx, selCell)
                ElseIf Button = 2 Then
                    If selCell.Col1 <> -1 And selCell.Row1 <> -1 Then
                        For Each tmpID In objReport.Items("_" & intIdx).SubIDs
                            Set tmpItem = objReport.Items("_" & tmpID.id)
                            If tmpItem.���� = 7 And tmpItem.��� = selCell.Col1 Then
                                Call SetDefaultState(tmpItem.����, tmpItem.����)
                            End If
                        Next
                    Else
                        Call SetDefaultState("", 0, True)
                    End If
                    mnuClass_Align.Visible = False '�������,�̶������
                    mnuClass_State.Visible = True
                    If strMenu = "" And Not ReferObj(Index) Then
                        If Not objReport.Items("_" & Index).ϵͳ Then PopupMenu mnuClass, 2
                    Else
                        strMenu = "mnuClass_Data"
                    End If
                End If
            ElseIf msh(intIdx).MouseRow <= msh(intIdx).FixedRows - 2 Then
                 '������෶Χ
                If Button = 1 Then
                    selCell.Row1 = msh(intIdx).MouseRow
                    selCell.Col1 = msh(intIdx).MouseCol
                    Call ClassColor(intIdx, selCell)
                ElseIf Button = 2 Then
                    If selCell.Col1 <> -1 And selCell.Row1 <> -1 Then
                        For Each tmpID In objReport.Items("_" & intIdx).SubIDs
                            Set tmpItem = objReport.Items("_" & tmpID.id)
                            If tmpItem.���� = 8 And tmpItem.��� = selCell.Row1 Then
                                Call SetDefaultState(tmpItem.����, tmpItem.����)
                            End If
                        Next
                    Else
                        Call SetDefaultState("", 0, True)
                    End If
                    mnuClass_Align.Visible = False '�������,�̶��ж���
                    mnuClass_State.Visible = True
                    If strMenu = "" Then
                        If Not objReport.Items("_" & Index).ϵͳ Then PopupMenu mnuClass, 2
                    Else
                        strMenu = "mnuClass_Data"
                    End If
                End If
            ElseIf msh(intIdx).MouseCol >= msh(intIdx).FixedCols And msh(intIdx).MouseRow >= msh(intIdx).FixedRows - 1 Then
                'ͳ���Χ
                If Button = 1 Then
                    selCell.Row1 = msh(intIdx).MouseRow
                    If msh(intIdx).MouseCol <= msh(intIdx).FixedCols + intState - 1 Then
                        selCell.Col1 = msh(intIdx).MouseCol
                    Else
                        selCell.Col1 = msh(intIdx).FixedCols + (msh(intIdx).MouseCol - msh(intIdx).FixedCols) Mod intState
                    End If
                    Call ClassColor(intIdx, selCell, intState)
                ElseIf Button = 2 Then
                    If selCell.Col1 <> -1 And selCell.Row1 <> -1 Then
                        For Each tmpID In objReport.Items("_" & intIdx).SubIDs
                            Set tmpItem = objReport.Items("_" & tmpID.id)
                            If tmpItem.���� = 9 And tmpItem.��� = selCell.Col1 - msh(intIdx).FixedCols Then
                                Call SetDefaultState(tmpItem.����, tmpItem.����)
                            End If
                        Next
                    Else
                        Call SetDefaultState("", 0, True)
                    End If
                    mnuClass_Align.Visible = True
                    mnuClass_State.Visible = False 'ͳ����,������ǻ���,����������ʽ�����ں��������
                    If strMenu = "" Then
                        If Not objReport.Items("_" & Index).ϵͳ Then PopupMenu mnuClass, 2
                    Else
                        strMenu = "mnuClass_Data"
                    End If
                End If
            End If
            
            '����Ƿ�����ı�߶Ȼ���
            If blnAdjustRowHeight = False And msh(Index).MouseRow < msh(Index).FixedRows And Button = 1 And objReport.Items("_" & Index).ϵͳ = False Then
                msh(Index).Row = msh(Index).MouseRow
                msh(Index).Col = msh(Index).MouseCol
                If msh(Index).Row < msh(Index).FixedRows Then
                    blnAdjustRowHeight = (Button = 1 And (Y > msh(Index).CellTop + msh(Index).CellHeight - 100 And Y < msh(Index).CellTop + msh(Index).CellHeight + 100))
                End If
            End If
        End If
    End If
    If GetSelNum = 1 Then Call ShowAttrib(intIdx) 'ֻѡ��һ������ʾ����
End Sub

Private Sub msh_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim tmpCell As Cells, RowH As Single, lngRow As Long, sgnH As Single
    Dim LngLastRow As Long, LngLastCol As Long, arrHead, i As Integer
    Dim sgnAlig As Single, strCaption As String, tmpID As RelatID, tmpItem As RPTItem
    
    If Not (Left(msh(Index).Tag, 2) = "C_") Then
        If objReport.Items("_" & Index).ϵͳ Then Exit Sub
    End If
    
    DrawXY X + msh(Index).Left, Y + msh(Index).Top
    If msh(Index).MouseRow >= 0 And msh(Index).MouseRow <= msh(Index).Rows - 1 And _
        msh(Index).MouseCol >= 0 And msh(Index).MouseCol <= msh(Index).Cols - 1 Then
        msh(Index).ToolTipText = msh(Index).TextMatrix(msh(Index).MouseRow, msh(Index).MouseCol)
    End If
    If Left(msh(Index).Tag, 2) = "C_" Then
        If Button = 1 And Mid(msh(CInt(Mid(msh(Index).Tag, 3))).Tag, 1, 2) <> "" Then
            If blnLock Then Exit Sub
            Call MoveSelect(X - lngPreX, Y - lngPreY)
            If GetSelNum() = 1 Then ShowAttrib CInt(Mid(msh(Index).Tag, 3))
        End If
    Else
        If objReport.Items("_" & Index).���� = 5 Then
            If msh(Index).MouseRow <= msh(Index).FixedRows Then
                If blnAdjustRowHeight = False And blnAdjustColWidth = False And Button = 0 And objReport.Items("_" & Index).ϵͳ = False Then
                    LngLastRow = msh(Index).Row: LngLastCol = msh(Index).Col
                    msh(Index).Row = msh(Index).MouseRow: msh(Index).Col = msh(Index).MouseCol
                    msh(Index).MousePointer = IIF(Y > msh(Index).CellTop + msh(Index).CellHeight - 100 And Y < msh(Index).CellTop + msh(Index).CellHeight + 100 And msh(Index).MouseRow <> msh(Index).FixedRows, 7, 99)
                    msh(Index).Row = LngLastRow: msh(Index).Col = LngLastCol
                Else
                    msh(Index).MousePointer = 99
                    If mobjPicMove Is Nothing Then Set mobjPicMove = LoadResPicture("MERGE", vbResCursor)
                    If Not msh(Index).MouseIcon = mobjPicMove Then
                        Set msh(Index).MouseIcon = LoadResPicture("MOVE", vbResCursor)
                        Set mobjPicMove = msh(Index).MouseIcon
                    End If
                End If
                
                If blnAdjustRowHeight Then
                    On Error Resume Next
                    If intCurID = 0 Then blnAdjustRowHeight = False: Exit Sub
                    If selCell.Row = -1 Then blnAdjustRowHeight = False: Exit Sub
                    msh(Index).MousePointer = 7
                    msh(Index).RowHeight(msh(Index).Row) = msh(Index).RowHeight(msh(Index).Row) + Y - lngPreY
                    lngPreY = Y

                    '������ʾ����������
                    LngLastCol = 0
                    For LngLastRow = 0 To msh(intCurID).FixedRows + 1
                        LngLastCol = LngLastCol + msh(intCurID).RowHeight(LngLastRow)
                    Next
                    If LngLastCol > msh(intCurID).Height - 100 * sgnMode Then
                        arrHead = Split(objReport.Items("_" & objReport.Items("_" & intCurID).SubIDs(1).Key).��ͷ, "|")
                        msh(intCurID).RowHeight(selCell.Row) = Split(arrHead(selCell.Row), "^")(1) * sgnMode
                    End If
'
                    '�����и�
                    For Each tmpID In objReport.Items("_" & intCurID).SubIDs
                        Set tmpItem = objReport.Items("_" & tmpID.id)
                        arrHead = Split(tmpItem.��ͷ, "|")
                        tmpItem.��ͷ = ""
                        For i = 0 To 2 'UBound(arrHead)
                            If i >= selCell.Row1 And i <= selCell.Row2 Then
                                msh(Index).Row = i: msh(Index).Col = tmpItem.���
                                sgnH = msh(Index).RowHeight(i)
                                sgnAlig = msh(Index).CellAlignment
                                strCaption = msh(Index).TextMatrix(msh(Index).Row, msh(Index).Col)
                                If strCaption = "" Then strCaption = "#"
                                tmpItem.��ͷ = tmpItem.��ͷ & "|" & sgnAlig & "^" & sgnH / sgnMode & "^" & strCaption
                            Else
                                tmpItem.��ͷ = tmpItem.��ͷ & "|" & arrHead(i)
                            End If
                        Next
                        tmpItem.��ͷ = Mid(tmpItem.��ͷ, 2)
                    Next
                    On Error GoTo 0
'                ElseIf msh(Index).MouseRow = -1 Then
'                    If Button = 1 And Mid(msh(Index).Tag, 1, 2) <> "" And blnAdjustColWidth = False Then
'                        If blnLock Then Exit Sub
'                        Call MoveSelect(X - lngPreX, Y - lngPreY)
'                        If GetSelNum() = 1 Then ShowAttrib Index
'                    End If
                End If
            ElseIf Button = 1 And Mid(msh(Index).Tag, 1, 2) <> "" Then
                If blnLock Then Exit Sub
                Call MoveSelect(X - lngPreX, Y - lngPreY)
                If GetSelNum() = 1 Then ShowAttrib Index
            Else
                msh(Index).MousePointer = 99
                If mobjPicMove Is Nothing Then Set mobjPicMove = LoadResPicture("MERGE", vbResCursor)
                If Not msh(Index).MouseIcon = mobjPicMove Then
                    Set msh(Index).MouseIcon = LoadResPicture("MOVE", vbResCursor)
                    Set mobjPicMove = msh(Index).MouseIcon
                End If
            End If
            
        ElseIf objReport.Items("_" & Index).���� = 4 Then
            If msh(Index).MouseRow < msh(Index).FixedRows Then
                If blnAdjustRowHeight = False And blnAdjustColWidth = False And Button = 0 And objReport.Items("_" & Index).ϵͳ = False Then
                    LngLastRow = msh(Index).Row: LngLastCol = msh(Index).Col
                    msh(Index).Row = msh(Index).MouseRow: msh(Index).Col = msh(Index).MouseCol
                    msh(Index).MousePointer = IIF(Y > msh(Index).CellTop + msh(Index).CellHeight - 100 And Y < msh(Index).CellTop + msh(Index).CellHeight + 100, 7, 99)
                    msh(Index).Row = LngLastRow: msh(Index).Col = LngLastCol
                Else
                    msh(Index).MousePointer = 99
                    If mobjPicMove Is Nothing Then Set mobjPicMove = LoadResPicture("MERGE", vbResCursor)
                    If Not msh(Index).MouseIcon = mobjPicMove Then
                        Set msh(Index).MouseIcon = LoadResPicture("MOVE", vbResCursor)
                        Set mobjPicMove = msh(Index).MouseIcon
                    End If
                End If
                If blnAdjustColWidth = False And msh(Index).MousePointer = 99 Then
                    If mobjPicMERGE Is Nothing Then Set mobjPicMERGE = LoadResPicture("MOVE", vbResCursor)
                    If Not msh(Index).MouseIcon = mobjPicMERGE Then
                        Set msh(Index).MouseIcon = LoadResPicture("MERGE", vbResCursor)
                        Set mobjPicMERGE = msh(Index).MouseIcon
                    End If
                End If
                If blnAdjustRowHeight Then
                    On Error Resume Next
                    If intCurID = 0 Then blnAdjustRowHeight = False: Exit Sub
                    If selCell.Row = -1 Then blnAdjustRowHeight = False: Exit Sub
                    
                    msh(Index).MousePointer = 7
                    PicFontTest.FontName = objReport.Items("_" & intCurID).����
                    PicFontTest.FontSize = objReport.Items("_" & intCurID).�ֺ�
                    sgnH = PicFontTest.TextHeight("��") + 15
                    
                    msh(Index).RowHeight(msh(Index).Row) = msh(Index).RowHeight(msh(Index).Row) + Y - lngPreY
                    For lngRow = selCell.Row1 To selCell.Row2
                        If Abs(msh(Index).RowHeight(lngRow)) < sgnH Then msh(Index).RowHeight(lngRow) = sgnH * sgnMode
                    Next
                    lngPreY = Y
                    
                    '������ʾ����������
                    LngLastCol = 0
                    For LngLastRow = 0 To msh(intCurID).FixedRows + 1
                        LngLastCol = LngLastCol + msh(intCurID).RowHeight(LngLastRow)
                    Next
                    If LngLastCol > msh(intCurID).Height - 100 * sgnMode Then
                        arrHead = Split(objReport.Items("_" & objReport.Items("_" & intCurID).SubIDs(1).Key).��ͷ, "|")
                        msh(intCurID).RowHeight(selCell.Row) = Split(arrHead(selCell.Row), "^")(1) * sgnMode
                    End If
                    
                    '�����и�
                    For Each tmpID In objReport.Items("_" & intCurID).SubIDs
                        Set tmpItem = objReport.Items("_" & tmpID.id)
                        arrHead = Split(tmpItem.��ͷ, "|")
                        tmpItem.��ͷ = ""
                        For i = 0 To UBound(arrHead)
                            If i >= selCell.Row1 And i <= selCell.Row2 Then
                                msh(Index).Row = i: msh(Index).Col = tmpItem.���
                                sgnH = msh(Index).RowHeight(i)
                                sgnAlig = msh(Index).CellAlignment
                                strCaption = msh(Index).TextMatrix(msh(Index).Row, msh(Index).Col)
                                If strCaption = "" Then strCaption = "#"
                                tmpItem.��ͷ = tmpItem.��ͷ & "|" & sgnAlig & "^" & sgnH / sgnMode & "^" & strCaption
                            Else
                                tmpItem.��ͷ = tmpItem.��ͷ & "|" & arrHead(i)
                            End If
                        Next
                        tmpItem.��ͷ = Mid(tmpItem.��ͷ, 2)
                    Next
                ElseIf msh(Index).MouseRow = -1 Then
                    If Button = 1 And Mid(msh(Index).Tag, 1, 2) <> "" And blnAdjustColWidth = False Then
                        If blnLock Then Exit Sub
                        Call MoveSelect(X - lngPreX, Y - lngPreY)
                        If GetSelNum() = 1 Then ShowAttrib Index
                    End If
                End If
            ElseIf msh(Index).MouseRow = msh(Index).FixedRows Then
                '�������϶�(ͬʱ�ı������еĿ��)
                If objReport.Items("_" & Index).ϵͳ = False And blnAdjustColWidth = False And msh(Index).MouseCol = 0 And Button = 0 Then
                    LngLastRow = msh(Index).Row: LngLastCol = msh(Index).Col
                    msh(Index).Row = msh(Index).MouseRow: msh(Index).Col = msh(Index).MouseCol
                    If msh(Index).Row = msh(Index).FixedRows And msh(Index).Col = 0 Then
                        msh(Index).MousePointer = IIF(X > msh(Index).CellLeft + msh(Index).CellWidth - 100 And X < msh(Index).CellLeft + msh(Index).CellWidth + 100, 9, 99)
                    End If
                    msh(Index).Row = LngLastRow: msh(Index).Col = LngLastCol
                End If
                If blnAdjustColWidth Then
                    '��ʾ�ָ���,�������п�
                    msh(Index).MousePointer = 9
                    With PicSplit
                        .Left = X + msh(Index).Left
                        .Top = msh(Index).Top
                        .Height = msh(Index).CellTop + msh(Index).CellHeight
                        .ZOrder 0
                        .Visible = True
                    End With
                Else
                    PicSplit.Visible = False
                    If msh(Index).MouseCol <> 0 Then
                        msh(Index).MousePointer = 99
                        If mobjPicMove Is Nothing Then Set mobjPicMove = LoadResPicture("MERGE", vbResCursor)
                        If Not msh(Index).MouseIcon = mobjPicMove Then
                            Set msh(Index).MouseIcon = LoadResPicture("MOVE", vbResCursor)
                            Set mobjPicMove = msh(Index).MouseIcon
                        End If
                    End If
                End If
            Else
                msh(Index).MousePointer = 99
                If mobjPicMove Is Nothing Then Set mobjPicMove = LoadResPicture("MERGE", vbResCursor)
                If Not msh(Index).MouseIcon = mobjPicMove Then
                    Set msh(Index).MouseIcon = LoadResPicture("MOVE", vbResCursor)
                    Set mobjPicMove = msh(Index).MouseIcon
                End If
            End If
            If Not blnHead Then
                If Button = 1 And Mid(msh(Index).Tag, 1, 2) <> "" And blnAdjustColWidth = False Then
                    If blnLock Then Exit Sub
                    Call MoveSelect(X - lngPreX, Y - lngPreY)
                    If GetSelNum() = 1 Then ShowAttrib Index
                End If
            ElseIf Button = 1 And blnAdjustRowHeight = False Then '�϶�ѡ��Ԫ��Χ
                If msh(Index).MouseRow >= 0 And msh(Index).MouseRow <= msh(Index).FixedRows - 1 And _
                    msh(Index).MouseCol >= 0 And msh(Index).MouseCol <= msh(Index).Cols - 1 Then
                    drgCell.Row2 = msh(Index).MouseRow
                    drgCell.Col2 = msh(Index).MouseCol
                    
                    tmpCell = MergeCell(Index, selCell, drgCell)
                    If tmpCell.Row1 <> -1 And tmpCell.Col1 <> -1 And tmpCell.Row2 <> -1 And tmpCell.Col2 <> -1 Then
                        Call CustomCellColor(Index, tmpCell)
                    Else
                        Call CustomCellColor(Index, selCell)
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub msh_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim tmpCell As Cells, i As Integer
    Dim tmpID As RelatID, tmpItem As RPTItem, arrHead, arrModify
    Dim sgnH As Long, sgnAlig As Long, strCaption As String
    Dim blnYes As Boolean, j As Long, k As Long
    Dim strSouse As String
    
    If blnAdjustColWidth Then
        PicSplit.Visible = False
        blnAdjustColWidth = False
        If PicSplit.Left < msh(Index).Left + 200 * sgnMode Or PicSplit.Left > msh(Index).Width - 200 * sgnMode Then Call VBA.Beep: Exit Sub
        sgnH = PicSplit.Left - msh(Index).Left
        For i = 0 To msh(Index).Cols - 1
            msh(Index).ColWidth(i) = sgnH
        Next
        
        '���ļ���
        For Each tmpID In objReport.Items("_" & intCurID).SubIDs
            Set tmpItem = objReport.Items("_" & tmpID.id)
            tmpItem.W = sgnH / sgnMode
        Next
        If GetSelNum = 1 Then ShowAttrib (Index)
    End If
    If blnAdjustRowHeight Then
        '����̶��е��и�
        If selCell.Row <> -1 Then
            For Each tmpID In objReport.Items("_" & intCurID).SubIDs
                Set tmpItem = objReport.Items("_" & tmpID.id)
                arrHead = Split(tmpItem.��ͷ, "|")
                tmpItem.��ͷ = ""
                For i = 0 To UBound(arrHead)
                    If i >= selCell.Row1 And i <= selCell.Row2 Then
                        msh(Index).Row = i: msh(Index).Col = tmpItem.���
                        sgnH = msh(Index).RowHeight(i)
                        sgnAlig = msh(Index).CellAlignment
                        strCaption = msh(Index).TextMatrix(msh(Index).Row, msh(Index).Col)
                        If strCaption = "" Then strCaption = "#"
                        tmpItem.��ͷ = tmpItem.��ͷ & "|" & sgnAlig & "^" & sgnH / sgnMode & "^" & strCaption
                    Else
                        tmpItem.��ͷ = tmpItem.��ͷ & "|" & arrHead(i)
                    End If
                Next
                tmpItem.��ͷ = Mid(tmpItem.��ͷ, 2)
            Next
        End If
        blnAdjustRowHeight = False
        If GetSelNum = 1 Then ShowAttrib (Index)
    End If
    
    If Left(msh(Index).Tag, 2) <> "C_" Then
        If objReport.Items("_" & Index).���� = 5 Then
        
        ElseIf objReport.Items("_" & Index).���� = 4 Then
            If Button = 1 And blnHead Then
                tmpCell = MergeCell(Index, selCell, drgCell)
                tmpCell.Row = selCell.Row
                If tmpCell.Row1 <> -1 And tmpCell.Col1 <> -1 And tmpCell.Row2 <> -1 And tmpCell.Col2 <> -1 Then selCell = tmpCell
                selCell = AdjustCell(selCell)
                Call CustomCellColor(Index, selCell)
            End If
        End If
   
        If Not mobjMove Is Nothing And objReport.Items("_" & Index).���� = 4 Then
            If Not mobjMove Is msh(Index).Container Or mobjMove Is picPaper Then
                If UCase(mobjMove.name) = "PIC" Then
                    If objReport.Items("_" & Index).���� > 1 Then
                        MsgBox "��Ƭ�в������������ı��", vbInformation, App.Title
                        Set mobjMove = Nothing: mlngX = 0: mlngY = 0
                        Exit Sub
                    End If
                    '��Ƭ�ڲ������ӱ��
                    For Each tmpItem In objReport.Items
                        If tmpItem.��ʽ�� = mbytCurrFmt Then
                            If tmpItem.���� = 5 Or tmpItem.���� = 4 Then
                                If tmpItem.���� = objReport.Items("_" & Index).���� Then
                                     MsgBox "������ڸ��ӱ�񣬲�������뿨Ƭ�У�", vbInformation, App.Title
                                     Set mobjMove = Nothing: mlngX = 0: mlngY = 0
                                     Exit Sub
                                End If
                            End If
                        End If
                    Next
                    '�����Ƭ������Դ�������������Դ�Ƿ�ƥ��
                    If objReport.Items("_" & mobjMove.Index).����Դ <> "" Then
                        For Each tmpID In objReport.Items("_" & Index).SubIDs
                            With objReport.Items("_" & tmpID.id)
                                i = InStr(1, .����, "]")
                                j = InStr(1, .����, ".")
                                k = InStr(1, .����, "[")
                                If i > k And i > j And i <> 0 And k <> 0 And j <> 0 Then
                                    If Mid(.����, k + 1, j - k - 1) <> objReport.Items("_" & mobjMove.Index).����Դ Then
                                        If blnYes = False Then
                                            If MsgBox("��ǰ��Ƭ��������Դ��������е������кͿ�Ƭ����Դ����ͬ�����뽫��ղ�ƥ����У��Ƿ����?", vbQuestion + vbYesNo + vbDefaultButton1, App.Title) = vbYes Then
                                                .���� = ""
                                                msh(Index).TextMatrix(1, .���) = ""
                                                blnYes = True
                                            Else
                                                Exit Sub
                                            End If
                                        Else
                                            .���� = ""
                                            msh(Index).TextMatrix(1, .���) = ""
                                        End If
                                    End If
                                End If
                            End With
                        Next
                    Else
                        '��Ƭû������Դ������ʾ�û��Ƿ��������Դ
                        For Each tmpID In objReport.Items("_" & Index).SubIDs
                            With objReport.Items("_" & tmpID.id)
                                i = InStr(1, .����, "]")
                                j = InStr(1, .����, ".")
                                k = InStr(1, .����, "[")
                                If i > k And i > j And i <> 0 And k <> 0 And j <> 0 Then
                                    If InStr(strSouse, Mid(.����, k + 1, j - k - 1)) = 0 Then
                                        strSouse = strSouse & "," & Mid(.����, k + 1, j - k - 1)
                                    End If
                                End If
                            End With
                        Next
                        strSouse = Mid(strSouse, 2)
                        'ֻ��һ������Դʱ����ʾ
                        If InStr(strSouse, ",") = 0 And strSouse <> "" Then
                            If MsgBox("��ǰ��Ƭδ������Դ���󶨺󽫷����ӡ���ſ�Ƭ������Դ�д���""�����ʶ""�ֶ���""�����ʶ""��ͬ��Ϊһ��,����һ������Ϊһ�飻" & vbCrLf & _
                                 "������ֻ��ӡһ�ſ�Ƭ���Ƿ������Դ""" & strSouse & """?", vbQuestion + vbYesNo + vbDefaultButton1, App.Title) = vbYes Then
                                objReport.Items("_" & mobjMove.Index).����Դ = strSouse
                            End If
                        End If
                    End If
                End If
                
                For Each tmpItem In objReport.Items
                    If tmpItem.��ʽ�� = mbytCurrFmt Then
                        If tmpItem.���� = 2 Then
                            If tmpItem.���� = objReport.Items("_" & Index).���� Then
                                If UCase(mobjMove.name) = "PIC" Then
                                    objReport.Items("_" & tmpItem.id).��ID = mobjMove.Index
                                    lbl(tmpItem.id).Left = lbl(tmpItem.id).Left + (mlngX - msh(Index).Left)
                                    lbl(tmpItem.id).Top = lbl(tmpItem.id).Top + (mlngY - msh(Index).Top)
                                    objReport.Items("_" & tmpItem.id).X = lbl(tmpItem.id).Left
                                    objReport.Items("_" & tmpItem.id).Y = lbl(tmpItem.id).Top
                                Else
                                    objReport.Items("_" & tmpItem.id).��ID = 0
                                    lbl(tmpItem.id).Left = lbl(tmpItem.id).Left + lbl(tmpItem.id).Container.Left
                                    lbl(tmpItem.id).Top = lbl(tmpItem.id).Top + lbl(tmpItem.id).Container.Top
                                    objReport.Items("_" & tmpItem.id).X = lbl(tmpItem.id).Left
                                    objReport.Items("_" & tmpItem.id).Y = lbl(tmpItem.id).Top
                                End If
                                Set lbl(tmpItem.id).Container = mobjMove
                            End If
                        End If
                    End If
                Next
                
                Set msh(Index).Container = mobjMove
                msh(Index).Top = mlngY: msh(Index).Left = mlngX
                If UCase(mobjMove.name) = "PIC" Then
                    objReport.Items("_" & Index).��ID = mobjMove.Index
                Else
                    objReport.Items("_" & Index).��ID = 0
                End If
                If objReport.Items("_" & Index).���� = 4 Then
                    '��������
                    For Each tmpID In objReport.Items("_" & Index).SubIDs
                        objReport.Items("_" & tmpID.id).��ID = objReport.Items("_" & Index).��ID
                    Next
                End If
                objReport.Items("_" & Index).X = mlngX: objReport.Items("_" & Index).Y = mlngY
                
            End If
            Set mobjMove = Nothing: mlngX = 0: mlngY = 0
            Call ShowAttrib(Index)
        End If
    Else
        If Not mobjMove Is Nothing Then
            If UCase(mobjMove.name) = "PIC" Then
                MsgBox "��Ƭ�в������������ı��", vbInformation, App.Title
                Set mobjMove = Nothing: mlngX = 0: mlngY = 0
            End If
        End If
    End If
End Sub

Private Function ItemIsGraph(ByVal intID As Integer) As Boolean
'���ܣ��ж�ָ���ı���Ԫ��(��ǩ)�Ƿ�ͼƬ�ֶ�
    Dim strNode As String, objNode As Node
    Dim i As Integer, j As Integer, k As Integer
    
    If objReport.Items("_" & intID).���� = 2 Then
        i = InStr(objReport.Items("_" & intID).����, "]")
        j = InStr(objReport.Items("_" & intID).����, ".")
        k = InStr(objReport.Items("_" & intID).����, "[")
        If i > k And i > j And i <> 0 And k <> 0 Then
            strNode = Mid(objReport.Items("_" & intID).����, j + 1, i - j - 1)
            For Each objNode In tvwSQL.Nodes
                If mdlPublic.GetStdNodeText(objNode.Text) = strNode And IsType(Val(objNode.Tag), adLongVarBinary) Then
                    ItemIsGraph = True
                End If
            Next
        End If
    End If
End Function

Private Sub SetAttAutoSize(ByVal intIndex As Integer, ByVal blnFlag As Boolean)
'���ܣ��������ñ�ǩ�Զ�������С
    Dim ObjSel As Object, intType As Integer, lngSize As Long, intID As Integer

    On Error Resume Next
    Set ObjSel = GetInxObj(intIndex)
    objReport.Items("_" & intIndex).�Ե� = blnFlag
            
    intType = objReport.Items("_" & intIndex).����
    If intType = 13 Then 'QR��ά����
        If blnFlag Then
            Set ObjSel.Picture = DrawBarCode2D(ReplaceBracket(objReport.Items("_" & intIndex).����), frmFlash.picTemp, lngSize)
            
            ObjSel.Height = Format(lngSize * sgnMode, "0.00")
            ObjSel.Width = Format(lngSize * sgnMode, "0.00")
            
            intID = intIndex
            Call SelItem(intID, False)
            Call SelItem(intID, True)
            
            objReport.Items("_" & intIndex).W = lngSize
            objReport.Items("_" & intIndex).H = lngSize
        End If
    Else
        '����Ǳ�ǩ�����ֶ�Ϊͼ�ͣ�����ͼ�ʹ���
        '��Ϊ�ֶ�ͼ�������޷��Ե�,ͼ�Ͷ�������Ե�
        If ItemIsGraph(intIndex) Then intType = 11
        If intType = 11 Then 'ͼƬ,������ΪͼƬ�ı�ǩ
            If blnFlag And Not objReport.Items("_" & intIndex).ͼƬ Is Nothing Then
                
                Set ObjSel.Picture = objReport.Items("_" & intIndex).ͼƬ '������ԭʼͼƬΪ׼
                
                ObjSel.Width = objReport.Items("_" & intIndex).ͼƬ.Width * (15 / 26.46) * sgnMode
                ObjSel.Height = objReport.Items("_" & intIndex).ͼƬ.Height * (15 / 26.46) * sgnMode
                
                intID = intIndex
                Call SelItem(intID, False)
                Call SelItem(intID, True)
                
                objReport.Items("_" & intIndex).X = ObjSel.Left / sgnMode
                objReport.Items("_" & intIndex).Y = ObjSel.Top / sgnMode
                objReport.Items("_" & intIndex).W = ObjSel.Width / sgnMode
                objReport.Items("_" & intIndex).H = ObjSel.Height / sgnMode
            End If
        ElseIf intType = 2 Then '��ǩ
            ObjSel.AutoSize = blnFlag
            If ObjSel.AutoSize Then '�Ե��������LblSize�ؼ���λ��
                intID = intIndex
                Call SelItem(intID, False)
                Call SelItem(intID, True)
                Call ReferTo
                objReport.Items("_" & intIndex).X = ObjSel.Left / sgnMode
                objReport.Items("_" & intIndex).Y = ObjSel.Top / sgnMode
                objReport.Items("_" & intIndex).W = ObjSel.Width / sgnMode
                objReport.Items("_" & intIndex).H = ObjSel.Height / sgnMode
            End If
        End If
    End If
End Sub

Private Sub mshAtt_DblClick()
    Dim intType As Integer
    Dim blnFlag As Boolean, blnFlagOld As Boolean
    Dim ObjSel As Object, objSub As RelatID
    Dim objBarCode As StdPicture
    Dim strBarCode As String
    Dim tmpObj As PictureBox
    
    If blnLock Then Exit Sub
    
    If Not (intCurID = 0 And (mshAtt.TextMatrix(mshAtt.Row, 0) = "Ʊ��" _
        Or mshAtt.TextMatrix(mshAtt.Row, 0) = "�ձ��ӡ" Or mshAtt.TextMatrix(mshAtt.Row, 0) = "��ֽ̬��")) Then
        
        If GetSelNum = 0 Then Exit Sub
    
        'ϵͳ��Ŀ(��ǩ)������༭
        If InStr(1, "�Զ�������С,�߿�", mshAtt.TextMatrix(mshAtt.Row, 0)) > 0 And objReport.Items("_" & intCurID).ϵͳ Then Exit Sub
            
        Set ObjSel = GetInxObj(intCurID)
    End If
    
    blnFlagOld = mshAtt.TextMatrix(mshAtt.Row, 1) = "��"
    If mshAtt.TextMatrix(mshAtt.Row, 1) = "��" Then
        mshAtt.TextMatrix(mshAtt.Row, 1) = "��"
        blnFlag = False: BlnSave = False
    ElseIf mshAtt.TextMatrix(mshAtt.Row, 1) = "��" Then
        mshAtt.TextMatrix(mshAtt.Row, 1) = "��"
        blnFlag = True: BlnSave = False
    End If
    
    Select Case mshAtt.TextMatrix(mshAtt.Row, 0)
        Case "Ʊ��"
            objReport.Ʊ�� = blnFlag
        Case "�ձ��ӡ"
            objReport.��ӡ��ʽ = IIF(blnFlag, 0, 1)
        Case "��ֽ̬��"
            objReport.Fmts("_" & mbytCurrFmt).��ֽ̬�� = blnFlag
        Case "����"
            objReport.Items("_" & intCurID).�߿� = blnFlag
        Case "����"
            ObjSel.FontBold = blnFlag
            If objReport.Items("_" & intCurID).���� = 4 Then Call SetCopyGrid(intCurID)
            objReport.Items("_" & intCurID).���� = blnFlag
        Case "б��"
            ObjSel.FontItalic = blnFlag
            If objReport.Items("_" & intCurID).���� = 4 Then Call SetCopyGrid(intCurID)
            objReport.Items("_" & intCurID).б�� = blnFlag
        Case "�»���"
            ObjSel.FontUnderline = blnFlag
            If objReport.Items("_" & intCurID).���� = 4 Then Call SetCopyGrid(intCurID)
            objReport.Items("_" & intCurID).���� = blnFlag
        Case "�߿�"
            '�����ѡ�е�����Ԫ�ض�����ͬ��Ԫ��
            If lblSize.count > 9 Then
                For Each tmpObj In lblSize
                    If tmpObj.Index Mod 8 = 1 Then
                        Set ObjSel = GetInxObj(tmpObj.Tag)
                        ObjSel.BorderStyle = IIF(blnFlag, 1, 0)
                        objReport.Items("_" & tmpObj.Tag).�߿� = blnFlag
                    End If
                Next
            Else
                ObjSel.BorderStyle = IIF(blnFlag, 1, 0)
                objReport.Items("_" & intCurID).�߿� = blnFlag
            End If
        Case "����"
            If blnFlag = False Then
                '���������еġ�����Ӧ�иߡ�����
                Set ObjSel = objReport.Items("_" & intCurID)
                If Not ObjSel Is Nothing Then
                    If ObjSel.���� = Val("4-�����") Then
                        For Each objSub In ObjSel.SubIDs
                            If objReport.Items("_" & objSub.id).����Ӧ�и� Then
                                ObjSel.�Ե� = blnFlagOld
                                mshAtt.TextMatrix(mshAtt.Row, 1) = IIF(blnFlagOld, "��", "��")
                                MsgBox "���Ƚ�����������еġ����е�Ԫ��ĸ߶��������Զ�����������ȡ����"
                                Exit Sub
                            End If
                        Next
                    End If
                End If
            End If
            objReport.Items("_" & intCurID).�Ե� = blnFlag
        Case "����"
            If cboAtt.Visible Then
                cboAtt.ListIndex = (cboAtt.ListIndex + 1) Mod 3
                Call cboAtt_Click
            End If
        Case "��״"
            If cboAtt.Visible Then
                cboAtt.ListIndex = (cboAtt.ListIndex + 1) Mod 2
                Call cboAtt_Click
            End If
        Case "���ֱ���"
            objReport.Items("_" & intCurID).���� = blnFlag
            '���ֱ���
            If Not objReport.Items("_" & intCurID).ͼƬ Is Nothing Then
                If objReport.Items("_" & intCurID).���� Then
                    Set ObjSel.Picture = ScalePicture(PicFontTest, objReport.Items("_" & intCurID).ͼƬ, ObjSel.Width, ObjSel.Height)
                Else
                    Set ObjSel.Picture = objReport.Items("_" & intCurID).ͼƬ
                End If
            End If
        Case "�Զ�����"
            '�����ѡ�е�����Ԫ�ض�����ͬ��Ԫ��
            If lblSize.count > 9 Then
                For Each tmpObj In lblSize
                    If tmpObj.Index Mod 8 = 1 Then
                        objReport.Items("_" & tmpObj.Tag).�и� = IIF(blnFlag, 1, 0)
                    End If
                Next
            Else
                objReport.Items("_" & intCurID).�и� = IIF(blnFlag, 1, 0)
            End If
        Case "�Զ�������С"
            '�����ѡ�е�����Ԫ�ض�����ͬ��Ԫ��
            If lblSize.count > 9 Then
                On Error Resume Next
                For Each tmpObj In lblSize
                    If tmpObj.Index Mod 8 = 1 Then
                        Call SetAttAutoSize(tmpObj.Tag, blnFlag)
                    End If
                Next
                On Error GoTo 0
            Else
                Call SetAttAutoSize(intCurID, blnFlag)
            End If
        Case "ˮƽ��ת"
            objReport.Items("_" & intCurID).ˮƽ��ת = blnFlag
        Case "�Ӵ�"
            objReport.Items("_" & intCurID).���� = blnFlag
            If objReport.Items("_" & intCurID).���� = 10 Then
                ObjSel.BorderWidth = IIF(blnFlag, 2, 1)
            ElseIf objReport.Items("_" & intCurID).���� = 1 Then
                ObjSel.Height = IIF(blnFlag, 30, 15)
            End If
        Case "����߼Ӵ�"
            objReport.Items("_" & intCurID).����߼Ӵ� = blnFlag
            If objReport.Items("_" & intCurID).���� = 4 Or objReport.Items("_" & intCurID).���� = 5 Then
                ObjSel.GridLineWidth = IIF(blnFlag, 2, 1)
            End If
        Case "����ͼ��"
            objReport.Items("_" & intCurID).���� = blnFlag
        Case "ǰ��ɫ", "����ɫ", "����", "��������ɫ", "����", "��������", "��ͷ����ɫ", "����ɫ"
            If objReport.Items("_" & intCurID).ϵͳ And InStr(1, "4,5", objReport.Items("_" & intCurID).����) <> 0 Then Exit Sub
            If cmdAtt.Visible Then cmdAtt_Click
        Case "��ʾ����" '��������
            objReport.Items("_" & intCurID).��ͷ = SetBit(objReport.Items("_" & intCurID).��ͷ, 1, IIF(blnFlag, 1, 0))
            
            With objReport.Items("_" & intCurID)
                strBarCode = ReplaceBracket(.����)
                If strBarCode = "" Then strBarCode = "1234567890"
                
                Unload frmFlash 'ǿ�Ƴ�ʼPicture����Ȼ�л�����������
                If .��� = 1 Then
                    Set objBarCode = DrawBarCode128(frmFlash.picTemp, 3, strBarCode, Mid(.��ͷ, 1, 1) = "1")
                ElseIf .��� = 2 Then
                    Set objBarCode = DrawBarCode39(frmFlash.picTemp, 3, strBarCode, Mid(.��ͷ, 2, 1) = "1", Mid(.��ͷ, 1, 1) = "1")
                ElseIf .��� = 3 Then
                    Set objBarCode = DrawBarCode128Auto(frmFlash.picTemp, strBarCode, 0, .�и�, Mid(.��ͷ, 1, 1) = "1")
                End If
                If Val(Mid(.��ͷ, 3, 1)) <> 0 Then
                    Set objBarCode = PictureSpin(objBarCode, Val(Mid(.��ͷ, 3, 1)), frmFlash.picTemp)
                End If
                Set ObjSel.Picture = objBarCode
            End With
        Case "��У���" '��������
            objReport.Items("_" & intCurID).��ͷ = SetBit(objReport.Items("_" & intCurID).��ͷ, 2, IIF(blnFlag, 1, 0))
            With objReport.Items("_" & intCurID)
                If .��� = 2 Then
                    Set objBarCode = DrawBarCode39(frmFlash.picTemp, 3, ReplaceBracket(.����), Mid(.��ͷ, 2, 1) = "1", Mid(.��ͷ, 1, 1) = "1")
                End If
                If Val(Mid(.��ͷ, 3, 1)) <> 0 Then
                    Set objBarCode = PictureSpin(objBarCode, Val(Mid(.��ͷ, 3, 1)), frmFlash.picTemp)
                End If
                Set ObjSel.Picture = objBarCode
            End With
    End Select
End Sub

Private Sub mshAtt_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Call ClearXY
End Sub

Private Sub mshAtt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
        KeyAscii = 0: mshAtt_DblClick
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0: mshAtt_AfterRowColChange 0, 0, mshAtt.Row, mshAtt.Col
    End If
End Sub

Private Sub mshAtt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mshAtt.MouseRow > 0 Then mshAtt_AfterRowColChange 0, 0, mshAtt.Row, mshAtt.Col
End Sub

Private Sub pic_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    If Button = 1 And bytCurTool <> 0 Then
        picPaper.Cls
        
        'ȷ��ѡ������
        selArea.Right = X
        selArea.Bottom = Y
        blnDown = False
        If bytCurTool = 1 Then
            '������������ͼ����(�ؼ�����Զ�����Ϊ15)
            If Abs(selArea.Right - selArea.Left) >= Abs(selArea.Bottom - selArea.Top) Then
                selArea.Bottom = selArea.Top
            Else
                selArea.Right = selArea.Left
            End If
        End If
        
        TrueArea selArea '����ѡ������
        
        If bytCurTool = 5 Then
            '��������С���(W=1000+15,H=255*3+15)
            If selArea.Right - selArea.Left < 1015 Then selArea.Right = selArea.Left + 1015
            If selArea.Bottom - selArea.Top < 780 Then selArea.Bottom = selArea.Top + 780
        ElseIf bytCurTool = 6 Then
            'ͼ����С�ߴ�
            If selArea.Right - selArea.Left < Chart(0).Width Then selArea.Right = selArea.Left + Chart(0).Width
            If selArea.Bottom - selArea.Top < Chart(0).Height Then selArea.Bottom = selArea.Top + Chart(0).Height
        End If
        
        If bytCurTool = 0 Then
            'ѡ������Ԫ��
            Call SelAreaItem(selArea)
            i = GetSelNum
            If i = 1 Then
                Call ShowAttrib(intCurID)
            ElseIf i = 0 Then
                Call ShowAttrib(, True)
            Else
                Call ShowAttrib
            End If
        Else
            '����Ԫ��
            If Not (Abs(selArea.Left - selArea.Right) = 0 And Abs(selArea.Top - selArea.Bottom) = 0) Then '��������
                Call AddReportItem(, pic(Index))
                BlnSave = False
            End If
        End If
        blnDown = False
    End If
End Sub

Private Sub picBack_GotFocus()
    Call NoneEdit
End Sub

Private Sub picL_GotFocus()
    Call NoneEdit
End Sub

Private Sub picM_GotFocus()
    Call NoneEdit
End Sub

Private Sub picPaper_DragDrop(Source As Control, X As Single, Y As Single)
    If UCase(Source.name) = "TVWSQL" Then
        selArea.Left = X: selArea.Top = Y
        Call AddReportItem(True)
        BlnSave = False
    End If
End Sub

Private Sub picPaper_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    DrawXY CLng(X), CLng(Y)
    If UCase(Source.name) = "TVWSQL" Then
        If State = 1 Then
            Set tvwSQL.DragIcon = lvwPar.DragIcon
        ElseIf State = 0 Then
            If tvwSQL.SelectedItem.Children = 0 Then
                Set tvwSQL.DragIcon = scrHsc.DragIcon
            Else
                Set tvwSQL.DragIcon = scrVsc.DragIcon
            End If
        End If
    End If
End Sub

Private Sub picPaper_GotFocus()
    Oldwinproc = GetWindowLong(picPaper.hwnd, GWL_WNDPROC)
    SetWindowLong picPaper.hwnd, GWL_WNDPROC, AddressOf FlexScroll
    Call NoneEdit
End Sub

Private Sub picPaper_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyPageDown Then
        '��
        If scrVsc.Value + scrVsc.Max / 10 > scrVsc.Max Then
            scrVsc.Value = scrVsc.Max
        Else
            scrVsc.Value = scrVsc.Value + scrVsc.Max / 10
        End If
    ElseIf KeyCode = vbKeyPageUp Then
        '��
        If scrVsc.Value - scrVsc.Max / 10 < 0 Then
            scrVsc.Value = 0
        Else
            scrVsc.Value = scrVsc.Value - scrVsc.Max / 10
        End If
    End If
End Sub

Private Sub picPaper_LostFocus()
    SetWindowLong picPaper.hwnd, GWL_WNDPROC, Oldwinproc
End Sub

Private Sub picPaperSize_GotFocus(Index As Integer)
    Call NoneEdit
End Sub

Private Sub picPaperSize_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    '�������Ԫ���Ƶ��ɼ�λ��
    Dim objItem As RPTItem
    Dim ObjSel As Object
    
    For Each objItem In objReport.Items
        If objItem.��ʽ�� = mbytCurrFmt Then
            If objItem.X + objItem.W > objReport.Fmts(mbytCurrFmt).W And objReport.Fmts(mbytCurrFmt).W - objItem.W >= 0 Then
                objItem.X = objReport.Fmts(mbytCurrFmt).W - objItem.W
                Set ObjSel = GetInxObj(objItem.id)
                ObjSel.Left = objItem.X
            End If
            If objItem.Y + objItem.H > objReport.Fmts(mbytCurrFmt).H And objReport.Fmts(mbytCurrFmt).H - objItem.H >= 0 Then
                objItem.Y = objReport.Fmts(mbytCurrFmt).H - objItem.H
                Set ObjSel = GetInxObj(objItem.id)
                ObjSel.Top = objItem.Y
            End If
        End If
    Next
    SelClear
End Sub

Private Sub picR_GotFocus()
    Call NoneEdit
End Sub

Private Sub picRulerH_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Call ClearXY
End Sub

Private Sub picRulerV_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Call ClearXY
End Sub

Private Sub scrHsc_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Call ClearXY
End Sub

Private Sub scrHsc_GotFocus()
    Call NoneEdit
End Sub

Private Sub scrVsc_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Call ClearXY
End Sub

Private Sub scrVsc_GotFocus()
    Call NoneEdit
End Sub

Private Sub sta_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Call ClearXY
End Sub

Private Sub sta_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ClearXY
End Sub

Private Sub tb2_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Key = "History" Then
        Call mnuEdit_History_Click
    End If
End Sub

Private Sub tbr1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Call NoneEdit
    Select Case ButtonMenu.Key
        Case "Line"
            Call mnuEdit_ItemAdd_Click(0)
        Case "Frame"
            Call mnuEdit_ItemAdd_Click(1)
        Case "Label"
            Call mnuEdit_ItemAdd_Click(3)
        Case "Picture"
            Call mnuEdit_ItemAdd_Click(4)
        Case "Table"
            Call mnuEdit_ItemAdd_Click(5)
        Case "Chart"
            Call mnuEdit_ItemAdd_Click(6)
        Case "BarCode"
            Call mnuEdit_ItemAdd_Click(7)
    End Select
End Sub

Private Sub tbr1_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Call ClearXY
End Sub

Private Sub tbr2_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Call ClearXY
End Sub

Private Sub tbr2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ClearXY
End Sub

Private Sub tbr1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ClearXY
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ClearXY
End Sub

Private Sub lblAtt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ClearXY
End Sub

Private Sub lblNote_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ClearXY
End Sub

Private Sub lblSQL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ClearXY
End Sub

Private Sub lblTool_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ClearXY
End Sub

Private Sub lvwPar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ClearXY
End Sub

Private Sub tbrTool_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim i As Integer
    
    bytCurTool = Button.Index - 1
    If bytCurTool = 0 Then
        picPaper.ForeColor = &HFF0000: picPaper.MousePointer = 0
    Else
        picPaper.ForeColor = &HFF&: picPaper.MousePointer = 2
    End If
End Sub

Private Sub tbrTool_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Call ClearXY
End Sub

Private Sub tbrTool_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ClearXY
End Sub

Private Sub tvwSQL_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Call ClearXY
    If State = 1 Then
        Set tvwSQL.DragIcon = lvwPar.DragIcon
    ElseIf State = 0 Then
        If tvwSQL.SelectedItem.Children = 0 Then
            Set tvwSQL.DragIcon = scrHsc.DragIcon
        Else
            Set tvwSQL.DragIcon = scrVsc.DragIcon
        End If
    End If
End Sub

Private Sub tvwSQL_GotFocus()
    Call NoneEdit
End Sub

Private Sub tvwSQL_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objNode As Object, tmpNode As Object
    Dim i As Integer, j As Integer
    
    lngPreY = Y: lngPreX = X
    
    Set objNode = tvwSQL.HitTest(X, Y)
    
    '�����Ƿ���϶�
    blnDrop = True
    blnSum = True
    If objNode Is Nothing Then
        blnDrop = False
    ElseIf objNode.Key = "Root" Then
        blnDrop = False
    ElseIf objNode.Children <> 0 Then
        If objReport.Datas(objNode.Key).���� = 0 Then
            '������
            If Not objNode.Checked Then blnDrop = False
        Else
            '���ܱ��
            Set tmpNode = objNode.Child
            Do While Not tmpNode Is Nothing
                If tmpNode.Checked Then
                    If IsType(Val(tmpNode.Tag), adLongVarBinary) Then
                        blnDrop = False: Exit Do '��ͼƬ�ֶβ����������ܱ��
                    ElseIf IsType(Val(tmpNode.Tag), adNumeric) Then
                        i = i + 1 'i��ʾ�������ֶθ���
                    Else
                        j = j + 1 'i��ʾ�����ֶθ���
                    End If
                End If
                Set tmpNode = tmpNode.Next
            Loop
            If i < 1 Then blnSum = False
            If i < 1 Or j < 1 Or i + j < 2 Then blnDrop = False
        End If
    End If
    If blnDrop Then
        Set tvwSQL.SelectedItem = objNode
        tvwSQL_NodeClick objNode
    End If
End Sub

Private Sub tvwsql_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    Dim objNode As Node
    
    
    Call ClearXY
    ClipCursor ByVal vbNullString
    If blnDrop And Button = 1 And (Abs(lngPreY - Y) > 300 Or Abs(lngPreX - X) > 300) Then
        If tvwSQL.SelectedItem.Children = 0 Then
            Set tvwSQL.DragIcon = scrHsc.DragIcon
        Else
            Set objNode = tvwSQL.SelectedItem.Child
            Do While Not objNode Is Nothing
                If objNode.Checked Then i = i + 1
                Set objNode = objNode.Next
            Loop
            If i = 0 Then
                ClipCursor GetObjRECT(tvwSQL.hwnd)
                Exit Sub
            End If
            Set tvwSQL.DragIcon = scrVsc.DragIcon
        End If
        tvwSQL.Drag 1
    ElseIf Button = 1 Then
        If blnSum = False And (X > (tvwSQL.Width - 200) Or Y > tvwSQL.Height - 200) Then
            MsgBox "����Դ����Ϊ������ܱ�ȱ�ٻ����ֶΣ�", vbInformation + vbOKOnly, Me.Caption
        End If
        ClipCursor GetObjRECT(tvwSQL.hwnd)
    End If
End Sub

Private Sub mshAtt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ClearXY
End Sub

Private Sub picback_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ClearXY
End Sub

Private Sub picRulerH_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ClearXY
End Sub

Private Sub picRulerV_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ClearXY
End Sub

Private Sub mnuViewToolAttrib_Click()
    mnuViewToolAttrib.Checked = Not mnuViewToolAttrib.Checked
    picR.Visible = Not picR.Visible
    picAtt.Visible = Not picAtt.Visible
    Call Form_Resize
End Sub

Private Sub mnuViewToolRuler_Click()
    mnuViewToolRuler.Checked = Not mnuViewToolRuler.Checked
    picRulerH.Visible = Not picRulerH.Visible
    picRulerV.Visible = Not picRulerV.Visible
    Form_Resize
End Sub

Private Sub mnuViewToolSQL_Click()
    mnuViewToolSQL.Checked = Not mnuViewToolSQL.Checked
    picL.Visible = Not picL.Visible
    picSQL.Visible = Not picSQL.Visible
    Call Form_Resize
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    sta.Visible = Not sta.Visible
    Call Form_Resize
End Sub

Private Sub picL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ClearXY
    If Button = 1 Then
        If picSQL.Width + X < 1000 Or picBack.Width - X < 2000 Then Exit Sub
        If cboFormat.Width - X < 3000 Then Exit Sub
        picSQL.Width = picSQL.Width + X
        lblSQL.Width = lblSQL.Width + X
        tvwSQL.Width = tvwSQL.Width + X
        lblPar.Width = lblPar.Width + X
        lvwPar.Width = lvwPar.Width + X
        picRulerV.Left = picRulerV.Left + X
        picRulerH.Left = picRulerH.Left + X
        picRulerH.Width = picRulerH.Width - X
        scrHsc.Left = scrHsc.Left + X
        scrHsc.Width = scrHsc.Width - X
        picBack.Left = picBack.Left + X
        picBack.Width = picBack.Width - X
    
        'zyb#Add
        picFormat.Left = picFormat.Left + X
        picFormat.Width = picFormat.Width - X
        cmdDel.Left = picFormat.Width - cmdDel.Width
        cmdAdd.Left = cmdDel.Left - cmdAdd.Width
        cboFormat.Width = cmdAdd.Left - cboFormat.Left - 50
        
        Call ShowSize
        Call ShowScroll
        If Not scrHsc.Enabled Then DrawRuler picRulerH
        
        Refresh
    End If
End Sub

Private Sub picPaper_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lngPreX = X: lngPreY = Y
    blnDown = True

    If Button = 1 Then
        selArea.Left = X
        selArea.Top = Y
        bytLine = 0
        If Shift <> 2 Then Call SelClear
        Call ShowAttrib(, True)
    End If
    If Button = 2 Then PopupMenu mnuFormat, 2
End Sub

Private Sub picPaper_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static PreX As Long, PreY As Long
    
    If Shift = 0 Then picPaper.MousePointer = 0
    Call DrawXY(CLng(X), CLng(Y))
    
    '��ѡ�����
    If Button = 1 And blnDown And Shift <> 4 Then
        If PreX = Empty And PreY = Empty Then
            PreX = selArea.Left
            PreY = selArea.Top
        End If
        If bytCurTool <> 1 Then
            picPaper.Line (selArea.Left, selArea.Top)-(PreX, PreY), picPaper.BackColor, B
            picPaper.Line (selArea.Left, selArea.Top)-(X, Y), , B
        Else
            If Abs(X - selArea.Left) >= Abs(Y - selArea.Top) Then
                '������
                If bytLine = 2 Then picPaper.Cls
                picPaper.Line (selArea.Left, selArea.Top)-(PreX, selArea.Top), picPaper.BackColor
                picPaper.Line (selArea.Left, selArea.Top)-(X, selArea.Top)
                bytLine = 1
            Else
                '������
                If bytLine = 1 Then picPaper.Cls
                picPaper.Line (selArea.Left, selArea.Top)-(selArea.Left, PreY), picPaper.BackColor
                picPaper.Line (selArea.Left, selArea.Top)-(selArea.Left, Y)
                bytLine = 2
            End If
        End If
        PreX = X: PreY = Y
    End If
    
    '�ƶ�ֽ��
    If Button = 1 And Shift = 4 And blnDown Then
        If scrVsc.Enabled Then
            If (Y - lngPreY) / Screen.TwipsPerPixelX > 0 Then
                scrVsc.Value = IIF(scrVsc.Value - (Y - lngPreY) / Screen.TwipsPerPixelX < scrVsc.Min, scrVsc.Min, scrVsc.Value - (Y - lngPreY) / Screen.TwipsPerPixelX)
            Else
                scrVsc.Value = IIF(scrVsc.Value - (Y - lngPreY) / Screen.TwipsPerPixelX > scrVsc.Max, scrVsc.Max, scrVsc.Value - (Y - lngPreY) / Screen.TwipsPerPixelX)
            End If
        End If
        If scrHsc.Enabled Then
            If (X - lngPreX) / Screen.TwipsPerPixelX > 0 Then
                scrHsc.Value = IIF(scrHsc.Value - (X - lngPreX) / Screen.TwipsPerPixelX < scrHsc.Min, scrHsc.Min, scrHsc.Value - (X - lngPreX) / Screen.TwipsPerPixelX)
            Else
                scrHsc.Value = IIF(scrHsc.Value - (X - lngPreX) / Screen.TwipsPerPixelX > scrHsc.Max, scrHsc.Max, scrHsc.Value - (X - lngPreX) / Screen.TwipsPerPixelX)
            End If
        End If
    End If
End Sub

Private Sub picPaper_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    If Button = 1 And blnDown Then
        picPaper.Cls
        
        'ȷ��ѡ������
        selArea.Right = X
        selArea.Bottom = Y
        
        If bytCurTool = 1 Then
            '������������ͼ����(�ؼ�����Զ�����Ϊ15)
            If Abs(selArea.Right - selArea.Left) >= Abs(selArea.Bottom - selArea.Top) Then
                selArea.Bottom = selArea.Top
            Else
                selArea.Right = selArea.Left
            End If
        End If
        
        TrueArea selArea '����ѡ������
        
        If bytCurTool = 5 Then
            '��������С���(W=1000+15,H=255*3+15)
            If selArea.Right - selArea.Left < 1015 Then selArea.Right = selArea.Left + 1015
            If selArea.Bottom - selArea.Top < 780 Then selArea.Bottom = selArea.Top + 780
        ElseIf bytCurTool = 6 Then
            'ͼ����С�ߴ�
            If selArea.Right - selArea.Left < Chart(0).Width Then selArea.Right = selArea.Left + Chart(0).Width
            If selArea.Bottom - selArea.Top < Chart(0).Height Then selArea.Bottom = selArea.Top + Chart(0).Height
        End If
        
        If bytCurTool = 0 Then
            'ѡ������Ԫ��
            Call SelAreaItem(selArea)
            i = GetSelNum
            If i = 1 Then
                Call ShowAttrib(intCurID)
            ElseIf i = 0 Then
                Call ShowAttrib(, True)
            Else
                Call ShowAttrib
            End If
        Else
            '����Ԫ��
            If Not (Abs(selArea.Left - selArea.Right) = 0 And Abs(selArea.Top - selArea.Bottom) = 0) Then '��������
                Call AddReportItem
                BlnSave = False
            End If
        End If
        blnDown = False
    End If
End Sub

Private Sub picR_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call ClearXY
    If Button = 1 Then
        If picAtt.Width - X < 2000 Or picBack.Width + X < 2000 Then Exit Sub
        If cboFormat.Width + X < 3000 Then Exit Sub
        lblTool.Width = lblTool.Width - X
        tbrTool.Width = lblTool.Width
        lblAtt.Top = tbrTool.Top + tbrTool.Height
        lblAtt.Width = lblAtt.Width - X
        mshAtt.Top = lblAtt.Top + lblAtt.Height
        mshAtt.Width = mshAtt.Width - X
        mshAtt.Height = picM.Top - mshAtt.Top
        lblNote.Width = lblNote.Width - X
        
        picAtt.Width = picAtt.Width - X
        picRulerH.Width = picRulerH.Width + X
        scrVsc.Left = scrVsc.Left + X
        scrHsc.Width = scrHsc.Width + X
        picBack.Width = picBack.Width + X
        picM.Width = picM.Width - X
        
        'zyb#Add
        picFormat.Width = picFormat.Width + X
        cmdDel.Left = picFormat.Width - cmdDel.Width
        cmdAdd.Left = cmdDel.Left - cmdAdd.Width
        cboFormat.Width = cmdAdd.Left - cboFormat.Left - 50
        
        Call ShowSize
        Call ShowScroll
        If Not scrHsc.Enabled Then DrawRuler picRulerH
        
        Refresh
    End If
End Sub

Private Sub picM_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If lblNote.Height - Y < 300 Or mshAtt.Height + Y < 2000 Then Exit Sub
        picM.Top = picM.Top + Y
        mshAtt.Height = mshAtt.Height + Y
        lblNote.Top = lblNote.Top + Y
        lblNote.Height = lblNote.Height - Y
        Refresh
    End If
End Sub

Private Sub tbr1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Call NoneEdit
    Select Case Button.Key
        Case "Quit"
            mnuFile_Quit_Click
        Case "New"
            mnuEdit_New_Click
        Case "Modi"
            mnuEdit_Modi_Click
        Case "Del"
            mnuEdit_Del_Click
        Case "Page"
            mnuFile_Page_Click
        Case "Save"
            mnuFile_Save_Click
        Case "Remove"
            mnuEdit_Remove_Click
        Case "Report"
            mnuFile_Report_Click
        Case "Guide"
            mnuFile_Guide_Click
        Case "Help"
            mnuHelpTitle_Click
        Case "AddFormat"
            mnuEdit_AddFormat_Click
        Case "DelFormat"
            mnuEdit_DelFormat_Click
    End Select
End Sub

Private Sub mnuViewToolText_Click()
    Dim But As Button
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    If mnuViewToolText.Checked Then
        For Each But In tbr1.Buttons
            But.Caption = But.Tag
        Next
        For Each But In tbr2.Buttons
            But.Caption = But.Tag
        Next
        For Each But In tbrTool.Buttons
            But.Caption = But.Tag
        Next
    Else
        For Each But In tbr1.Buttons
            But.Caption = ""
        Next
        For Each But In tbr2.Buttons
            But.Caption = ""
        Next
        For Each But In tbrTool.Buttons
            But.Caption = ""
        Next
    End If
    cbr.Bands("System").MinHeight = tbr1.ButtonHeight
    cbr.Bands("Format").MinHeight = tbr2.ButtonHeight
    tbrTool.Height = tbrTool.ButtonHeight
    Form_Resize
End Sub

Private Sub mnuViewToolFormat_Click()
    mnuViewToolFormat.Checked = Not mnuViewToolFormat.Checked
    cbr.Bands("Format").Visible = Not cbr.Bands("Format").Visible
    If cbr.Bands("Format").Visible And cbr.Bands("System").Visible Then
        If cbr.Bands("System").Position = 2 Then cbr.Bands("System").NewRow = True
        If cbr.Bands("Format").Position = 2 Then cbr.Bands("Format").NewRow = True
    End If
    If Not mnuViewToolFormat.Checked And Not mnuViewToolSystem.Checked Then
        cbr.Visible = False
        mnuViewToolText.Enabled = False
    Else
        cbr.Visible = True
        mnuViewToolText.Enabled = True
    End If
    Form_Resize
End Sub

Private Sub mnuViewToolSystem_Click()
    mnuViewToolSystem.Checked = Not mnuViewToolSystem.Checked
    cbr.Bands("System").Visible = Not cbr.Bands("System").Visible
    If cbr.Bands("Format").Visible And cbr.Bands("System").Visible Then
        If cbr.Bands("System").Position = 2 Then cbr.Bands("System").NewRow = True
        If cbr.Bands("Format").Position = 2 Then cbr.Bands("Format").NewRow = True
    End If
    If Not mnuViewToolFormat.Checked And Not mnuViewToolSystem.Checked Then
        cbr.Visible = False
        mnuViewToolText.Enabled = False
    Else
        cbr.Visible = True
        mnuViewToolText.Enabled = True
    End If
    Form_Resize
End Sub

Private Sub DrawRuler(picRuler As PictureBox, Optional lngBegin As Long = 0)
'����:��ʾ�������
'����:picRuler=��߿ؼ�;lngBegin=��ʼ����ֵ(��λ:Twip,X��Y),Ӧ��Ϊ����(<=0)
    Dim X As Long, Y As Long, IntStep As Integer
    Const FaceColor = &HA8A8A8
    
    IntStep = 283 * sgnMode
    With picRuler
        .Cls
        .DrawMode = vbCopyPen
        .FontName = "Times New Roman"
        .FontSize = 7.5
        .ForeColor = &H800000
        If .Width > .Height Then
            '����
            '����
            picRuler.Line (0, 0)-(Screen.Width, .ScaleHeight / 4), FaceColor, BF
            picRuler.Line (.ScaleHeight - .ScaleHeight / 4, .ScaleHeight - .ScaleHeight / 4)-(Screen.Width, .ScaleHeight), FaceColor, BF
            picRuler.Line (0, 0)-(.ScaleHeight / 4, .ScaleHeight), FaceColor, BF
            picRuler.Line (.ScaleWidth - .ScaleHeight / 4, 0)-(.ScaleWidth, .ScaleHeight), FaceColor, BF
            '��ע
            For X = .ScaleHeight + lngBegin To .ScaleWidth Step IntStep  '0.5cm
                If ((X - .ScaleHeight - lngBegin) / IntStep) Mod 2 = 0 Then
                    '����
                    .CurrentY = .ScaleHeight / 2 - .TextHeight("0") / 2
                    .CurrentX = X - .TextWidth(CStr(((X - .ScaleHeight - lngBegin) / IntStep) / 2) & "0") / 2
                    picRuler.Print ((X - .ScaleHeight - lngBegin) / IntStep) / 2
                    '�̶�
                    picRuler.Line (X, .ScaleHeight - .ScaleHeight / 4)-(X, .ScaleHeight), &HFFFFFF
                    picRuler.Line (X, 0)-(X, .ScaleHeight / 4), &HFFFFFF
                ElseIf ((X - .ScaleHeight - lngBegin) / IntStep) Mod 2 = 1 Then
                    picRuler.Line (X, .ScaleHeight - .ScaleHeight / 8 - 15)-(X, .ScaleHeight - .ScaleHeight / 8 + 15), &HFFFFFF
                    picRuler.Line (X, .ScaleHeight / 8 - 15)-(X, .ScaleHeight / 8 + 15), &HFFFFFF
                End If
            Next
        Else
            '����
            '����
            picRuler.Line (0, 0)-(.ScaleWidth / 4, Screen.Height), FaceColor, BF
            picRuler.Line (.ScaleWidth - .ScaleWidth / 4, 0)-(.ScaleWidth, Screen.Height), FaceColor, BF
            picRuler.Line (0, .ScaleHeight - .ScaleWidth / 4)-(.ScaleWidth, .ScaleHeight), FaceColor, BF
            '��ע
            For Y = lngBegin To .ScaleHeight Step IntStep  '0.5cm
                If ((Y - lngBegin) / IntStep) Mod 2 = 0 Then
                    '����
                    .CurrentX = .ScaleWidth / 4
                    .CurrentY = Y + .TextWidth(CStr(((Y - lngBegin) / IntStep) / 2)) / 2
                    objFont.OutPut picRuler, .CurrentX, .CurrentY, ((Y - lngBegin) / IntStep) / 2
                    '�̶�
                    picRuler.Line (.ScaleWidth - .ScaleWidth / 4, Y)-(.ScaleWidth, Y), &HFFFFFF
                    picRuler.Line (0, Y)-(.ScaleWidth / 4, Y), &HFFFFFF
                ElseIf ((Y - lngBegin) / IntStep) Mod 2 = 1 Then
                    picRuler.Line (.ScaleWidth - .ScaleWidth / 8 - 15, Y)-(.ScaleWidth - .ScaleWidth / 8 + 15, Y), &HFFFFFF
                    picRuler.Line (.ScaleWidth / 8 - 15, Y)-(.ScaleWidth / 8 + 15, Y), &HFFFFFF
                End If
            Next
        End If
        .ForeColor = &HFFFF00
        .DrawMode = vbXorPen
    End With
End Sub

Private Sub picPaperSize_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lngPreX = X: lngPreY = Y
End Sub

Private Sub picPaperSize_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'����:�ı䱨��ֽ�Ŵ�С
    Dim lngX As Long, lngY As Long
    Dim lngW As Long, lngH As Long
    
    Call ClearXY
    
    If Button = 1 Then
        If blnLock Then Exit Sub
        lngX = X - lngPreX
        lngY = Y - lngPreY
        
        'һ���ı�,����Ϊ�Զ���ֽ��
        With objReport.Fmts("_" & mbytCurrFmt)
            .ֽ�� = 256
            If .ֽ�� = 2 Then '�������
                .W = .W + .H
                .H = .W - .H
                .W = .W - .H
                .ֽ�� = 1
            End If
        End With
        
        Select Case Index
            Case 0 '�ı���
                If picPaper.Width + lngX < 283 * sgnMode Then Exit Sub
                
                picPaperSize(0).Left = picPaperSize(0).Left + lngX
                picPaper.Width = picPaper.Width + lngX
                objReport.Fmts("_" & mbytCurrFmt).W = picPaper.Width / sgnMode
                picPaperSize(2).Left = picPaperSize(2).Left + lngX
                picPaperSize(1).Width = picPaperSize(1).Width + lngX
                
                lngW = picBack.ScaleWidth - (picPaper.Width + picPaperSize(2).Width * 2 - Abs(picPaper.Left))
                If lngW < 0 Then lngW = 0
                
                If picBack.ScaleWidth >= picPaper.Width + picPaperSize(2).Width * 2 + lngW Then
                    scrHsc.Enabled = False
                Else
                    scrHsc.Enabled = True
                    scrHsc.Max = (picPaper.Width + picPaperSize(2).Width * 2 + lngW - picBack.ScaleWidth) / Screen.TwipsPerPixelX
                End If
            Case 1 '�ı�߶�
                If picPaper.Height + lngY < 283 * sgnMode Then Exit Sub
                
                picPaperSize(1).Top = picPaperSize(1).Top + lngY
                picPaper.Height = picPaper.Height + lngY
                objReport.Fmts("_" & mbytCurrFmt).H = picPaper.Height / sgnMode
                picPaperSize(0).Height = picPaperSize(0).Height + lngY
                picPaperSize(2).Top = picPaperSize(2).Top + lngY
                
                lngH = picBack.ScaleHeight - (picPaper.Height + picPaperSize(2).Width * 2 - Abs(picPaper.Top))
                If lngH < 0 Then lngH = 0
                
                If picBack.ScaleHeight >= picPaper.Height + picPaperSize(2).Width * 2 + lngH Then
                    scrVsc.Enabled = False
                Else
                    scrVsc.Enabled = True
                    scrVsc.Max = (picPaper.Height + picPaperSize(2).Width * 2 + lngH - picBack.ScaleHeight) / Screen.TwipsPerPixelX
                End If
            Case 2 '�ı���
                If picPaper.Height + lngY >= 283 * sgnMode Then
                    picPaper.Height = picPaper.Height + lngY
                    objReport.Fmts("_" & mbytCurrFmt).H = picPaper.Height / sgnMode
                    
                    picPaperSize(2).Top = picPaperSize(2).Top + lngY
                    picPaperSize(0).Height = picPaperSize(0).Height + lngY
                    picPaperSize(1).Top = picPaperSize(1).Top + lngY
                    
                    lngH = picBack.ScaleHeight - (picPaper.Height + picPaperSize(2).Width * 2 - Abs(picPaper.Top))
                    If lngH < 0 Then lngH = 0
                    
                    If picBack.ScaleHeight >= picPaper.Height + picPaperSize(2).Width * 2 + lngH Then
                        scrVsc.Enabled = False
                    Else
                        scrVsc.Enabled = True
                        scrVsc.Max = (picPaper.Height + picPaperSize(2).Width * 2 + lngH - picBack.ScaleHeight) / Screen.TwipsPerPixelX
                    End If
                End If
                If picPaper.Width + lngX >= 283 * sgnMode Then
                    picPaper.Width = picPaper.Width + lngX
                    objReport.Fmts("_" & mbytCurrFmt).W = picPaper.Width / sgnMode
                    
                    picPaperSize(2).Left = picPaperSize(2).Left + lngX
                    picPaperSize(0).Left = picPaperSize(0).Left + lngX
                    picPaperSize(1).Width = picPaperSize(1).Width + lngX
                    
                    lngW = picBack.ScaleWidth - (picPaper.Width + picPaperSize(2).Width * 2 - Abs(picPaper.Left))
                    If lngW < 0 Then lngW = 0
                    
                    If picBack.ScaleWidth >= picPaper.Width + picPaperSize(2).Width * 2 + lngW Then
                        scrHsc.Enabled = False
                    Else
                        scrHsc.Enabled = True
                        scrHsc.Max = (picPaper.Width + picPaperSize(2).Width * 2 + lngW - picBack.ScaleWidth) / Screen.TwipsPerPixelX
                    End If
                End If
        End Select
        BlnSave = False
    End If
End Sub

Private Sub scrhsc_Change()
    Call NoneEdit
    Call ClearXY
    Call ShowSize(-scrVsc.Value * 15#, -scrHsc.Value * 15#)
    If scrHsc.Value = 0 Then Call ShowScroll(1)
End Sub

Private Sub scrhsc_Scroll()
    Call NoneEdit
    Call ClearXY
    Call ShowSize(-scrVsc.Value * 15#, -scrHsc.Value * 15#)
    If scrHsc.Value = 0 Then Call ShowScroll(1)
End Sub

Private Sub scrVsc_Change()
    Call NoneEdit
    Call ClearXY
    Call ShowSize(-scrVsc.Value * 15#, -scrHsc.Value * 15#)
    If scrVsc.Value = 0 Then Call ShowScroll(2)
End Sub

Private Sub scrVsc_Scroll()
    Call NoneEdit
    Call ClearXY
    Call ShowSize(-scrVsc.Value * 15#, -scrHsc.Value * 15#)
    If scrVsc.Value = 0 Then Call ShowScroll(2)
End Sub

Private Sub ShowReportDetail()
'���ܣ�����objReport������ʾ��������
    Call ShowSQLs
    Call ShowSize
    Call ShowScroll
    Call ShowItems
End Sub

Private Sub ShowSize(Optional lngTop As Single = 0, Optional lngLeft As Single = 0)
'����:��ʾ����ֽ�Ŵ�С
    Dim lngW As Long, lngH As Long
    
    picPaper.Left = lngLeft
    picPaper.Top = lngTop
    
    '��ӡ��ֽ��ֻ�Ǽ򵥵ؽ�ֽ�ſ�Ⱥ͸߶ȶԵ�
    With objReport.Fmts("_" & mbytCurrFmt)
        If .ֽ�� = 1 Then
            lngW = .W: lngH = .H
        Else
            lngH = .W: lngW = .H
        End If
    End With
    
    picPaper.Width = Format(lngW * sgnMode, "0.00")
    picPaper.Height = Format(lngH * sgnMode, "0.00")
    
    '��Ӱ��������λ��
    picPaperSize(0).Top = picPaper.Top + picPaperSize(0).Width
    picPaperSize(0).Left = picPaper.Left + picPaper.Width
    picPaperSize(0).Height = picPaper.Height - picPaperSize(0).Width
    
    picPaperSize(1).Top = picPaper.Top + picPaper.Height
    picPaperSize(1).Left = picPaper.Left + picPaperSize(1).Height
    picPaperSize(1).Width = picPaper.Width - picPaperSize(1).Height
    
    picPaperSize(2).Top = picPaperSize(1).Top
    picPaperSize(2).Left = picPaperSize(0).Left
        
    '���
    DrawRuler picRulerH, picPaper.Left + 15
    DrawRuler picRulerV, picPaper.Top + 15
    
    With objReport.Fmts("_" & mbytCurrFmt)
        sta.Panels(2).Text = "��ӡ��:" & objReport.��ӡ�� & "   ֽ��:" & GetPaperName(.ֽ��, .W, .H) & " " & _
            IIF(.ֽ�� = 256, CInt(.W / Twip_mm) & "mm �� " & CInt(.H / Twip_mm) & "mm", "") & _
            IIF(.ֽ�� = 1, "   ����", "   ����")
    End With
    
    Me.Refresh
End Sub

Private Sub ShowScroll(Optional bytType As Byte = 3)
'����:���ù�����
'����:bytType=3-���߶�����(ȱʡֵ),1-������Hsc,2-������Vsc
    
    If bytType = 3 Or bytType = 2 Then
        If picBack.ScaleHeight >= picPaper.Height + picPaperSize(2).Width * 2 Then
            scrVsc.Enabled = False
        Else
            scrVsc.Max = (picPaper.Height + picPaperSize(2).Width * 2 - picBack.ScaleHeight) / Screen.TwipsPerPixelX 'ת��Ϊ����Ϊ��λ
            Call ShowSize(0, picPaper.Left)
            scrVsc.Value = 0
            scrVsc.Enabled = True
        End If
    End If
    If bytType = 3 Or bytType = 1 Then
        If picBack.ScaleWidth >= picPaper.Width + picPaperSize(2).Width * 2 Then
            scrHsc.Enabled = False
        Else
            scrHsc.Max = (picPaper.Width + picPaperSize(2).Width * 2 - picBack.ScaleWidth) / Screen.TwipsPerPixelX
            Call ShowSize(picPaper.Top, 0)
            scrHsc.Value = 0
            scrHsc.Enabled = True
        End If
    End If
End Sub

Private Sub ShowSQLs()
'���ܣ�����objReport������ʾ��������Դ������
'˵����Ϊ�ӿ��ٶ�,�Զ�����SQL�ֶ�,ʵ���õ�ʱ�ٴ����ݡ�
    Dim tmpData As New RPTData
    Dim objNode As Object
    Dim arrFields() As String
    Dim i As Integer
    Dim strSource As String
    
    '��ʾ�������Դ���ֶ�
    tvwSQL.Nodes.Clear
    Set objNode = tvwSQL.Nodes.Add(, , "Root", objReport.����, "Root")
    objNode.Selected = True
    objNode.Expanded = True
    
    For Each tmpData In objReport.Datas
        If tmpData.�������ӱ�� > 0 Then
            '��������������ʾ���ӵ�����
            strSource = GetDBConnectInfo(tmpData.�������ӱ��)
            If strSource = "" Then
                strSource = tmpData.����
            Else
                strSource = tmpData.���� & "��" & strSource & "��"
            End If
        Else
            strSource = tmpData.����
        End If
        
        If tmpData.���� = 0 Then
            Set objNode = tvwSQL.Nodes.Add("Root", 4, "_" & tmpData.����, strSource, "SQL_Custom")
        Else
            Set objNode = tvwSQL.Nodes.Add("Root", 4, "_" & tmpData.����, strSource, "SQL_Group")
        End If
        objNode.Expanded = True
        
        '�����ֶ�����
        If tmpData.�ֶ� <> "" Then
            arrFields = Split(tmpData.�ֶ�, "|")
            For i = 0 To UBound(arrFields)
                Select Case Split(arrFields(i), ",")(1)
                    Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR '�ı���(Varchar2,Long)
                        Set objNode = tvwSQL.Nodes.Add("_" & tmpData.Key, 4, , Split(arrFields(i), ",")(0), "String")
                    Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, _
                        adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, _
                        adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt  '������(Numeric(a,b),Sum)
                        Set objNode = tvwSQL.Nodes.Add("_" & tmpData.Key, 4, , Split(arrFields(i), ",")(0), "Number")
                    Case adDBTimeStamp, adDBTime, adDBDate, adDate '������(Date)
                        Set objNode = tvwSQL.Nodes.Add("_" & tmpData.Key, 4, , Split(arrFields(i), ",")(0), "Date")
                    Case adBinary, adVarBinary, adLongVarBinary '������(Long Raw)
                        Set objNode = tvwSQL.Nodes.Add("_" & tmpData.Key, 4, , Split(arrFields(i), ",")(0), "Bin")
                    Case Else '����
                        Set objNode = tvwSQL.Nodes.Add("_" & tmpData.Key, 4, , Split(arrFields(i), ",")(0), "Other")
                End Select
                objNode.Tag = Split(arrFields(i), ",")(1) '����ֶε����ͣ�����
            Next
        End If
    Next
    If Not tvwSQL.SelectedItem.Child Is Nothing Then tvwSQL.SelectedItem.Child.Selected = True
    Call tvwSQL_NodeClick(tvwSQL.SelectedItem)
End Sub

Private Sub tbr2_ButtonClick(ByVal Button As MSComctlLib.Button)
    Call NoneEdit
    Select Case Button.Key
        Case "Lock"
            mnuFormat_Lock_Click
        Case "Left", "Right", "Up", "Down", "Hsc", "Vsc"
            Call mnuFormat_DoAlign_Click(Button.Index - 1)
        Case "Width"
            mnuFormat_Width_Click
        Case "Height"
            mnuFormat_Height_Click
        Case "WH"
            mnuFormat_WH_Click
        Case "Scale"
            tbr2_ButtonMenuClick tbr2.Buttons("Scale").ButtonMenus(1)
    End Select
End Sub

Private Sub ClearXY()
    '������λ��״̬��
    If preHsc <> Empty Then
        picRulerH.Line (preHsc, 0)-(preHsc, picRulerH.ScaleHeight)
        preHsc = Empty
    End If
    If preVsc <> Empty Then
        picRulerV.Line (0, preVsc)-(picRulerH.ScaleWidth, preVsc)
        preVsc = Empty
    End If
    sta.Panels(3) = "λ��"
End Sub

Private Sub DrawXY(X As Long, Y As Long)
'����:�������ϵ�λ����
'����:�����PicPaper��λ��
    If preHsc <> Empty Then picRulerH.Line (preHsc, 0)-(preHsc, picRulerH.ScaleHeight)
    If preVsc <> Empty Then picRulerV.Line (0, preVsc)-(picRulerV.ScaleWidth, preVsc)
    
    picRulerH.Line (X + picRulerH.ScaleHeight + 15 - Abs(picPaper.Left), 0)-(X + picRulerH.ScaleHeight + 15 - Abs(picPaper.Left), picRulerH.ScaleHeight)
    picRulerV.Line (0, Y + 15 - Abs(picPaper.Top))-(picRulerV.ScaleWidth, Y + 15 - Abs(picPaper.Top))
    preHsc = X + picRulerH.ScaleHeight + 15 - Abs(picPaper.Left)
    preVsc = Y + 15 - Abs(picPaper.Top)
    
    sta.Panels(3) = "X=" & Format(X / sgnMode / Twip_mm, "0.00") & "mm Y=" & Format(Y / sgnMode / Twip_mm, "0.00") & "mm"
End Sub

Private Sub TrueArea(ByRef Area As RECT)
'����:��ѡ���������Ϊ���Һ�����,����Χ����
'����:ѡ��Χ
    Dim X1 As Long, Y1 As Long, X2 As Long, Y2 As Long
    X1 = Area.Left: Y1 = Area.Top
    X2 = Area.Right: Y2 = Area.Bottom
    If X2 < X1 Then Area.Left = X2: Area.Right = X1
    If Y2 < Y1 Then Area.Top = Y2: Area.Bottom = Y1
End Sub

Private Sub tvwSQL_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'��ѡ����
    Dim objNode As Object
    
    Set objNode = tvwSQL.HitTest(X, Y)
    
    If Not objNode Is Nothing Then
        If objNode.Key = "Root" Then
            objNode.Checked = False
        ElseIf IsType(Val(objNode.Tag), adLongVarBinary) Then '�������ֶβ��������ܱ��
            If objReport.Datas(objNode.Parent.Key).���� = 1 Then
                objNode.Checked = False
            End If
        ElseIf objNode.Image = "Other" Then '�����Ͳ����������
            objNode.Checked = False
        End If
    End If
End Sub

Private Sub tvwSQL_NodeCheck(ByVal Node As MSComctlLib.Node)
    Dim objNode As Object, blnCheck As Boolean
    
    Set objNode = Node
    
    If objNode.Key <> "Root" Then
        If objNode.Parent.Key = "Root" Then
            blnCheck = objNode.Checked
            Set objNode = objNode.Child
            Do While Not objNode Is Nothing
                'If Not IsType(Val(objNode.Tag), adLongVarBinary) And objNode.Image <> "Other" Then
                If objNode.Image <> "Other" And Not (objReport.Datas(objNode.Parent.Key).���� = 1 And IsType(Val(objNode.Tag), adLongVarBinary)) Then
                    '�������ֶβ��������ܱ��
                    objNode.Checked = blnCheck
                    If blnCheck = True And objNode.Text = "�����ʶ" Then
                        objNode.Checked = False
                    End If
                Else
                    objNode.Checked = False
                End If
                Set objNode = objNode.Next
            Loop
        Else
            If objNode.Checked Then
                'If Not IsType(Val(objNode.Tag), adLongVarBinary) And objNode.Image <> "Other" Then
                If objNode.Image <> "Other" Then
                    objNode.Parent.Checked = True
                End If
            Else
                blnCheck = False
                Set objNode = objNode.Parent.Child
                Do While Not objNode Is Nothing
                    blnCheck = blnCheck Or objNode.Checked
                    If objNode.Next Is Nothing Then objNode.Parent.Checked = blnCheck
                    Set objNode = objNode.Next
                Loop
            End If
        End If
    End If
End Sub

Private Sub tvwSQL_NodeClick(ByVal Node As MSComctlLib.Node)
'���ܣ���ʾ��ǰ����Դ�Ĳ����嵥
    Dim objItem As Object
    Dim tmpPar As RPTPar
    Dim strKey As String
    
    lvwPar.ListItems.Clear
    
    If Node.Key <> "Root" Then
        If Node.Children = 0 Then
            strKey = Node.Parent.Key
        Else
            strKey = Node.Key
        End If
        For Each tmpPar In objReport.Datas(strKey).Pars
            Set objItem = lvwPar.ListItems.Add(, "_" & tmpPar.���, tmpPar.���)
            objItem.SubItems(1) = tmpPar.����
            Select Case tmpPar.����
                Case 0
                    objItem.SubItems(2) = "�ַ�"
                    objItem.SmallIcon = "String"
                Case 1
                    objItem.SubItems(2) = "����"
                    objItem.SmallIcon = "Number"
                Case 2
                    objItem.SubItems(2) = "����"
                    objItem.SmallIcon = "Date"
                Case 3
                    objItem.SubItems(2) = "������"
                    objItem.SmallIcon = "Other"
            End Select
            objItem.SubItems(3) = tmpPar.ȱʡֵ
        Next
    End If
End Sub

Private Sub AddReportItem(Optional ByVal blnDrop As Boolean = False, Optional ByVal objParent As Object)
'����:���һ��������Ŀ
'ʹ�ò���:
'   SelArea=��ǰѡ��Χ(�Ӳ˵����϶�����Ҫ�ֶ�����)
'   bytCurTool=��ǰҪ��ӵ�Ԫ������/��tvwSQL.SelectedItem=��ǰҪ��ӵ����ݱ�����Ŀ
    Dim newObj As Control, objNode As Object, tmpItem As RPTItem
    Dim intCols As Integer, i As Integer, j As Integer, k As Integer, l As Integer
    Dim X As Integer, Y As Integer, Z As Integer, sngWidth As Single
    Dim bytAlign As Byte, arrAlign() As Byte, Str��ͷ As String, tmpID As RelatID
    Dim intMaxIDtmp As Integer, intCurIDtmp As Integer
    
    If cboFormat.ComboItems("_" & mbytCurrFmt) Is Nothing Then
        MsgBox "��ѡ��Ҫ��ӱ���Ԫ�صı����ʽ��", vbInformation, App.Title
        Exit Sub
    End If
    If (bytCurTool = 6 Or bytCurTool = 8) And Not objParent Is Nothing Then
        MsgBox "��Ƭ�в��������" & IIF(bytCurTool = 6, "ͼ��", "��Ƭ") & "��", vbInformation, App.Title
        objParent.Refresh
        Exit Sub
    End If
    intMaxIDtmp = intMaxID: intCurIDtmp = intCurID
    intMaxID = intMaxID + 1
    intCurID = intMaxID
    
    If Not blnDrop Then
        Select Case bytCurTool
            Case 1 '����(lblLine)
                Load lblLine(intCurID)
                Set newObj = lblLine(intCurID)
                
                '�������ݶ���
                objReport.Items.Add intCurID, mbytCurrFmt, GetNextName(1), 0, 1, 0, "", 0, "", "", _
                    Format(selArea.Left / sgnMode, "0.00"), Format(selArea.Top / sgnMode, "0.00"), _
                    Format((selArea.Right - selArea.Left) / sgnMode, "0.00"), _
                    Format((selArea.Bottom - selArea.Top) / sgnMode, "0.00"), _
                    0, 0, True, "", 0, False, False, False, 0, 0, 0, False, 0, "", "", "", False, _
                    False, , False, , , , "_" & intCurID
            Case 3  '��ǩ(lbl)
                Load lbl(intCurID)
                Set newObj = lbl(intCurID)
                If bytCurTool = 3 Then
                    newObj.Caption = GetNextName(2)
                Else
                    newObj.Caption = ""
                End If
                
                '�������ݶ���
                objReport.Items.Add intCurID, mbytCurrFmt, GetNextName(IIF(bytCurTool = 2, 10, 2)), _
                    0, IIF(bytCurTool = 2, 10, 2), 0, "", 0, GetNextName(IIF(bytCurTool = 2, 10, 2)), _
                    "", Format(selArea.Left / sgnMode, "0.00"), Format(selArea.Top / sgnMode, "0.00"), _
                    Format((selArea.Right - selArea.Left) / sgnMode, "0.00"), _
                    Format((selArea.Bottom - selArea.Top) / sgnMode, "0.00"), _
                    0, 0, IIF(bytCurTool = 3, True, False), "����", 9, False, False, False, 0, 0, &HFFFFFF, _
                    False, 0, "", "", "", False, False, , False, , , , "_" & intCurID
            Case 2 '����(shp)
                Load lblshp(intCurID)
                Load Shp(intCurID)
                Set newObj = Shp(intCurID)
                lblshp(intCurID).BackColor = picPaper.BackColor
                
                '�������ݶ���
                objReport.Items.Add intCurID, mbytCurrFmt, GetNextName(10), 0, 10, 0, "", 0, GetNextName(10), _
                    "", Format(selArea.Left / sgnMode, "0.00"), Format(selArea.Top / sgnMode, "0.00"), _
                    Format((selArea.Right - selArea.Left) / sgnMode, "0.00"), _
                    Format((selArea.Bottom - selArea.Top) / sgnMode, "0.00"), _
                    0, 0, IIF(bytCurTool = 3, True, False), "����", 9, False, False, False, 0, 0, &HFFFFFF, _
                    False, 0, "", "", "", False, False, , False, , , , "_" & intCurID
            Case 4 'ͼƬ
                Load img(intCurID)
                Set newObj = img(intCurID)
                newObj.BorderStyle = 1
                
                '��Ƭ�ֶ�
                '�������ݶ���
                objReport.Items.Add intCurID, mbytCurrFmt, GetNextName(11), 0, 11, 0, "", 0, "", "", _
                    Format(selArea.Left / sgnMode, "0.00"), Format(selArea.Top / sgnMode, "0.00"), _
                    Format((selArea.Right - selArea.Left) / sgnMode, "0.00"), _
                    Format((selArea.Bottom - selArea.Top) / sgnMode, "0.00"), _
                    0, 0, False, "����", 9, False, False, False, 0, 0, &HFFFFFF, True, 0, "", "", "", _
                    False, False, , False, , , , "_" & intCurID
            Case 8 '��Ƭ
                Load pic(intCurID)
                Set newObj = pic(intCurID)
                newObj.BorderStyle = 1
                
                '��Ƭ�ֶ�
                '�������ݶ���
                objReport.Items.Add intCurID, mbytCurrFmt, GetNextName(14), 0, 14, 0, "", 0, "", "", _
                    Format(selArea.Left / sgnMode, "0.00"), Format(selArea.Top / sgnMode, "0.00"), _
                    Format((selArea.Right - selArea.Left) / sgnMode, "0.00"), _
                    Format((selArea.Bottom - selArea.Top) / sgnMode, "0.00"), _
                    0, 0, False, "����", 9, False, False, False, 0, 0, &HFFFFFF, True, 0, "", "", "", _
                    False, False, , False, , , , "_" & intCurID
            Case 5 '������(����)
                Load msh(intCurID)
                            
                'ȱʡ1����ͷ��,�������ɸ߶ȼ���(����2��)��
                msh(intCurID).Rows = Abs(Int(-(selArea.Bottom - selArea.Top) / (255 * sgnMode))) - 1
                If Not objReport.Ʊ�� Then
                    msh(intCurID).FixedRows = msh(intCurID).Rows - 2
                Else
                    msh(intCurID).FixedRows = 1
                End If
                msh(intCurID).FixedCols = 0
                '�����ɿ�ȼ���,����1��
                msh(intCurID).Cols = Abs(Int(-(selArea.Right - selArea.Left) / 1000))
                If msh(intCurID).Cols > 1 Then msh(intCurID).Cols = msh(intCurID).Cols - 1
                
                msh(intCurID).TextMatrix(0, 0) = "���" & msh.count - 1
                                
                msh(intCurID).Row = 0: msh(intCurID).Col = 0
                SetHeadCenter msh(intCurID)
                                
                Set newObj = msh(intCurID)
                
                '�������ݶ���(1:����Ϊ��,������Դ,2:����=1,3:�и�=255)
                Set tmpItem = objReport.Items.Add(intCurID, mbytCurrFmt, GetNextName(4), 0, 4, 0, "", 0, "", "", _
                    Format(selArea.Left / sgnMode, "0.00"), Format(selArea.Top / sgnMode, "0.00"), _
                    Format((selArea.Right - selArea.Left) / sgnMode, "0.00"), _
                    Format((selArea.Bottom - selArea.Top) / sgnMode, "0.00"), _
                    255, 0, False, "����", 9, False, False, False, 0, 0, &HFFFFFF, True, 1, "", "", "", False, _
                    False, , False, , , , "_" & intCurID)
                
                For i = 0 To msh(intCurID).Cols - 1
                    msh(intCurID).ColWidth(i) = 1000 * sgnMode
                    msh(intCurID).ColAlignment(i) = 1
                    
                    intMaxID = intMaxID + 1
                    Str��ͷ = ""
                    For j = 0 To msh(intCurID).FixedRows - 1
                        msh(intCurID).Row = j: msh(intCurID).Col = i
                        Str��ͷ = Str��ͷ & "|" & msh(intCurID).CellAlignment & "^" & msh(intCurID).RowHeight(j) & "^#"
                    Next
                    Str��ͷ = Mid(Str��ͷ, 2)
                    
                    '�������ļ���(1:����Ϊ��,��������,2:������)
                    objReport.Items.Add intMaxID, mbytCurrFmt, "����" & intMaxID, intCurID, 6, i, "", 0, "", _
                        Str��ͷ, _
                        0, 0, msh(intCurID).ColWidth(i) / sgnMode, 0, _
                        0, 0, False, "", 0, False, False, False, 0, 0, 0, False, 0, "", "", "", False, False, _
                         , False, , , , "_" & intMaxID
                    
                    tmpItem.SubIDs.Add intMaxID, "_" & intMaxID
                Next
                '���Ժϲ�
                For i = 0 To msh(intCurID).FixedRows - 1
                    msh(intCurID).MergeRow(i) = True
                Next
                For i = 0 To msh(intCurID).Cols - 1
                    msh(intCurID).MergeCol(i) = True
                Next
                
            Case 6 'ͼ��@@@
                Load Chart(intCurID)
                Set newObj = Chart(intCurID)
                
                'ȱʡΪ����ͼ,ͼ���ڶ�
                objReport.Items.Add intCurID, mbytCurrFmt, GetNextName(12), 0, 12, 1, "", 0, "", "", _
                    Format(selArea.Left / sgnMode, "0.00"), Format(selArea.Top / sgnMode, "0.00"), _
                    Format((selArea.Right - selArea.Left) / sgnMode, "0.00"), _
                    Format((selArea.Bottom - selArea.Top) / sgnMode, "0.00"), _
                    0, 1, True, "����", 9, False, True, False, 0, 0, &HFFFFFF, False, 0, "", "", "", _
                    False, False, , False, , , , "_" & intCurID
            Case 7 '����
                Load ImgCode(intCurID)
                Set newObj = ImgCode(intCurID)
                newObj.BorderStyle = 0
                
                Unload frmFlash 'ǿ�Ƴ�ʼPicture����Ȼ�л�����������
                Set newObj.Picture = DrawBarCode128Auto(frmFlash.picTemp, "1234567890", sngWidth, 2, True)
                
                '�������ݶ���
                objReport.Items.Add intCurID, mbytCurrFmt, GetNextName(13), 0, 13, 3, "", 0, "", "100", _
                    Format(selArea.Left / sgnMode, "0.00"), Format(selArea.Top / sgnMode, "0.00"), _
                    Format(Me.ScaleX(sngWidth, vbMillimeters, vbTwips), "0.00"), _
                    Format((selArea.Bottom - selArea.Top) / sgnMode, "0.00"), _
                    2, 0, False, "����", 9, False, False, False, 0, 0, &HFFFFFF, False, 0, "", "", "", _
                    False, False, , False, , , , "_" & intCurID
        End Select
                
        tbrTool.Buttons(1).Value = tbrPressed
        tbrTool_ButtonClick tbrTool.Buttons(1)
    Else
        If tvwSQL.SelectedItem.Children = 0 Then
            If Not objParent Is Nothing Then
                If objReport.Items("_" & objParent.Index).����Դ = "" Then
                    If MsgBox("��ǰ��Ƭδ������Դ���󶨺󽫷����ӡ���ſ�Ƭ������Դ�д���""�����ʶ""�ֶ���""�����ʶ""��ͬ��Ϊһ��,����һ������Ϊһ�飻" & vbCrLf & _
                         "������ֻ��ӡһ�ſ�Ƭ���Ƿ������Դ""" & tvwSQL.SelectedItem.Parent.Text & """?" _
                        , vbQuestion + vbYesNo + vbDefaultButton1, App.Title) = vbYes Then
                        objReport.Items("_" & objParent.Index).����Դ = mdlPublic.GetStdNodeText(tvwSQL.SelectedItem.Parent.Text)
                    End If
                End If
            End If
            '��������
            Load lbl(intCurID)
            Set newObj = lbl(intCurID)
            newObj.Caption = "[" & LevelText(tvwSQL.SelectedItem) & "]" '[]�Ų��������ݿ�
            If IsType(Val(tvwSQL.SelectedItem.Tag), adLongVarBinary) Then
                '��Ƭ�ֶ�
                newObj.BorderStyle = 1
                selArea.Bottom = selArea.Top + 1500
                selArea.Right = selArea.Left + 1300
                
                '�������ݶ���
                objReport.Items.Add intCurID, mbytCurrFmt, GetNextName(2), 0, 2, 0, "", 0, newObj.Caption, "", _
                    Format(selArea.Left / sgnMode, "0.00"), Format(selArea.Top / sgnMode, "0.00"), _
                    Format((selArea.Right - selArea.Left) / sgnMode, "0.00"), _
                    Format((selArea.Bottom - selArea.Top) / sgnMode, "0.00"), _
                    0, 0, False, "����", 9, False, False, False, 0, 0, &HFFFFFF, True, 0, "", "", "", _
                    False, False, , False, , , , "_" & intCurID
            Else
                newObj.Caption = mdlPublic.GetStdNodeText(tvwSQL.SelectedItem.Text) & ":" & newObj.Caption
                selArea.Bottom = selArea.Top + lbl(0).Height * sgnMode
                selArea.Right = selArea.Left + TextWidth(lbl(intCurID).Caption & "��") * sgnMode
                
                Select Case tvwSQL.SelectedItem.Tag
                Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger _
                    , adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                    bytAlign = 2 '����ȱʡ�Ҷ���
                    lbl(intCurID).Alignment = 1
                Case adDBTimeStamp, adDBTime, adDBDate, adDate
                    bytAlign = 1 '����ȱʡ�ж���
                    lbl(intCurID).Alignment = 2
                End Select
                
                '�������ݶ���
                objReport.Items.Add intCurID, mbytCurrFmt, GetNextName(2), 0, 2, 0, "", 0, newObj.Caption, "", _
                    Format(selArea.Left / sgnMode, "0.00"), Format(selArea.Top / sgnMode, "0.00"), _
                    Format((selArea.Right - selArea.Left) / sgnMode, "0.00"), _
                    Format((selArea.Bottom - selArea.Top) / sgnMode, "0.00"), _
                    0, bytAlign, True, "����", 9, False, False, False, 0, 0, &HFFFFFF, False, 0, "", "", "", _
                    False, False, , False, , , , "_" & intCurID
            End If
        Else
            If objReport.Datas(tvwSQL.SelectedItem.Key).���� = 0 Then
                '�����ֻ�ܷ���Ϳ�Ƭ����Դ��ͬ������Դ
                If Not objParent Is Nothing Then
                    If mdlPublic.GetStdNodeText(tvwSQL.SelectedItem.Text) <> objReport.Items("_" & objParent.Index).����Դ _
                        And objReport.Items("_" & objParent.Index).����Դ <> "" Then
                        MsgBox "��Ƭ��������Դ������ֻ�ܼ���Ϳ�Ƭ��ͬ����Դ�ı��", vbInformation, App.Title
                        intMaxID = intMaxIDtmp: intCurID = intCurIDtmp
                        Exit Sub
                    End If
                    '�����Ƭδָ������Դ�����Զ�ָ��
                    If objReport.Items("_" & objParent.Index).����Դ = "" Then
                        If MsgBox("��ǰ��Ƭδ������Դ���󶨺󽫷����ӡ���ſ�Ƭ������Դ�д���""�����ʶ""�ֶ���""�����ʶ""��ͬ��Ϊһ��,����һ������Ϊһ�飻" & vbCrLf & _
                             "������ֻ��ӡһ�ſ�Ƭ���Ƿ������Դ""" & tvwSQL.SelectedItem.Text & """?" _
                            , vbQuestion + vbYesNo + vbDefaultButton1, App.Title) = vbYes Then
                            objReport.Items("_" & objParent.Index).����Դ = mdlPublic.GetStdNodeText(tvwSQL.SelectedItem.Text)
                        End If
                    End If
                End If
                '������
                Load msh(intCurID)
                            
                'ȱʡ1����ͷ��,5�������С�
                selArea.Bottom = selArea.Top + (1545 * sgnMode) '255 * 3 + 15
                msh(intCurID).Rows = 6
                msh(intCurID).FixedRows = 1
                msh(intCurID).FixedCols = 0
                msh(intCurID).SelectionMode = flexSelectionFree
                
                '�������������,����1��
                msh(intCurID).Cols = Abs(Int(-(selArea.Right - selArea.Left) / (1000 * sgnMode)))
                
                i = 0
                Set objNode = tvwSQL.SelectedItem.Child
                Do While Not objNode Is Nothing
                    If objNode.Checked Then
                        If i = 0 Then
                            ReDim arrAlign(i)
                        Else
                            ReDim Preserve arrAlign(i)
                        End If
                        Select Case objNode.Tag
                        Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger _
                            , adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                            arrAlign(i) = 2
                        Case adDBTimeStamp, adDBTime, adDBDate, adDate
                            arrAlign(i) = 1
                        Case adBinary, adVarBinary, adLongVarBinary
                            arrAlign(i) = 1
                        Case Else
                            arrAlign(i) = 0
                        End Select
                        
                        i = i + 1
                        
                        msh(intCurID).Cols = i
                        msh(intCurID).ColWidth(i - 1) = 1000 * sgnMode
                        msh(intCurID).ColAlignment(i - 1) = 1
                        msh(intCurID).TextMatrix(0, i - 1) = objNode.Text
                        msh(intCurID).TextMatrix(1, i - 1) = "[" & LevelText(objNode) & "]" '[]��Ҫ�������ݿ�
                    End If
                    Set objNode = objNode.Next
                Loop
                selArea.Right = selArea.Left + msh(intCurID).Cols * (1000 * sgnMode) + 30
                
                msh(intCurID).Row = 0: msh(intCurID).Col = 0
                SetHeadCenter msh(intCurID)
                
                Set newObj = msh(intCurID)
                
                '�������ݶ���(1:����Ϊ����Դ��)
                Set tmpItem = objReport.Items.Add(intCurID, mbytCurrFmt, GetNextName(4), 0, 4, 0, "", 0, _
                    mdlPublic.GetStdNodeText(tvwSQL.SelectedItem.Text), "", _
                    Format(selArea.Left / sgnMode, "0.00"), Format(selArea.Top / sgnMode, "0.00"), _
                    Format((selArea.Right - selArea.Left) / sgnMode, "0.00"), _
                    Format((selArea.Bottom - selArea.Top) / sgnMode, "0.00"), _
                    255, 0, False, "����", 9, False, False, False, 0, 0, &HFFFFFF, True, 1, "", "", "", _
                    False, False, , False, , , , "_" & intCurID)
                
                For i = 0 To msh(intCurID).Cols - 1
                    intMaxID = intMaxID + 1
                    '�������ļ���(1:�޻���)
                    Str��ͷ = ""
                    For j = 0 To msh(intCurID).FixedRows - 1
                        msh(intCurID).Row = j: msh(intCurID).Col = i
                        Str��ͷ = Str��ͷ & "|" & msh(intCurID).CellAlignment & _
                                    "^" & msh(intCurID).RowHeight(j) & "^" & msh(intCurID).TextMatrix(j, i)
                    Next
                    Str��ͷ = Mid(Str��ͷ, 2)
                    objReport.Items.Add intMaxID, mbytCurrFmt, "����" & intMaxID, intCurID, 6, i, _
                        "", 0, msh(intCurID).TextMatrix(1, i), Str��ͷ, 0, 0, msh(intCurID).ColWidth(i) / sgnMode, _
                        0, 0, arrAlign(i), False, "", 0, False, False, False, 0, 0, 0, False, 0, "", "", "", _
                        False, False, , False, , , , "_" & intMaxID
                    
                    tmpItem.SubIDs.Add intMaxID, "_" & intMaxID
                Next
                '���Ժϲ�
                For i = 0 To msh(intCurID).FixedRows - 1
                    msh(intCurID).MergeRow(i) = True
                Next
                For i = 0 To msh(intCurID).Cols - 1
                    msh(intCurID).MergeCol(i) = True
                Next
                
            Else
                '��Ƭ�����������ܱ�
                If Not objParent Is Nothing Then
                    MsgBox "��Ƭ�в����������ܱ�", vbInformation, App.Title
                    Exit Sub
                End If
                '���ܱ��
                Load msh(intCurID)
                
                'x,y,z�ֱ��ʾ����/����/ͳ�Ʒ�������
                Set objNode = tvwSQL.SelectedItem.Child
                Do While Not objNode Is Nothing
                    If objNode.Checked Then
                        '�������ж����Ƽ�������
                        Select Case objNode.Tag
                            Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR, _
                                adDBTimeStamp, adDBTime, adDBDate, adDate  '�ı���������
                                X = X + 1
                            Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, _
                                adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, _
                                adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt  '������
                                Z = Z + 1
                        End Select
                    End If
                    Set objNode = objNode.Next
                Loop
                '�������,����һ��Ϊ�������
                If X >= 2 Then X = X - 1: Y = Y + 1
                
                '�������
                With msh(intCurID)
                    '�����
                    .Rows = Y + 1 + X * 5
                    .FixedRows = Y + 1
                    .Cols = X + IIF(Y = 0, Z, Z * 2)
                    .FixedCols = X
                    selArea.Bottom = selArea.Top + .Rows * (255 * sgnMode) + 60
                    selArea.Right = selArea.Left + .Cols * (1000 * sgnMode) + 60
                    For i = 0 To .Cols - 1
                        .ColWidth(i) = 1000 * sgnMode
                        .ColAlignment(i) = 1
                    Next
                    For i = 0 To .FixedCols - 1
                        .MergeCol(i) = True
                    Next
                    For i = 0 To .FixedRows - 2
                        .MergeRow(i) = True
                    Next
                    
                    '������ݼ��������
                    
                    '�������ݶ���(1:����Ϊ����Դ��)
                    Set tmpItem = objReport.Items.Add(intCurID, mbytCurrFmt, GetNextName(5), 0, 5, 0, "", 0, _
                        mdlPublic.GetStdNodeText(tvwSQL.SelectedItem.Text), "", Format(selArea.Left / sgnMode, "0.00"), _
                        Format(selArea.Top / sgnMode, "0.00"), Format((selArea.Right - selArea.Left) / sgnMode, "0.00"), _
                        Format((selArea.Bottom - selArea.Top) / sgnMode, "0.00"), 255, 0, False, "����", 9, False, False, _
                        False, 0, 0, &HFFFFFF, True, 0, "", "", "", False, False, , False, , , , "_" & intCurID)
                    
                    i = 0: j = 0
                    Set objNode = tvwSQL.SelectedItem.Child
                    Do While Not objNode Is Nothing
                        If objNode.Checked Then
                            intMaxID = intMaxID + 1
                            Select Case objNode.Tag
                                Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR, _
                                    adDBTimeStamp, adDBTime, adDBDate, adDate  '�ַ�������
                                    i = i + 1
                                    If i <= X Then
                                        .TextMatrix(.FixedRows - 1, i - 1) = "[" & mdlPublic.GetStdNodeText(objNode.Text) & "]"  '[]���������ݿ�
                                        For k = .FixedRows To .Rows - 1
                                            .TextMatrix(k, i - 1) = mdlPublic.GetStdNodeText(objNode.Text)
                                        Next
                                        
                                        '�������ļ���(1:�޻���,2:����ֱ��Ϊ�ֶ���,��ΪA.B����ʽ,�������Ѵ��)
                                        objReport.Items.Add intMaxID, mbytCurrFmt, "����" & intMaxID, intCurID, 7, i - 1, _
                                            "", 0, mdlPublic.GetStdNodeText(objNode.Text), "", 0, 0, 1000, 0, 255, 0, False, _
                                            "", 0, False, False, False, 0, 0, 0, False, 0, "", "", "", False, False, , _
                                            False, , , , "_" & intMaxID
                                    Else
                                        For k = 0 To .FixedRows - 2
                                            For l = 0 To .FixedCols - 1
                                                .TextMatrix(k, l) = "[" & mdlPublic.GetStdNodeText(objNode.Text) & "]"
                                            Next
                                            For l = .FixedCols To .Cols - 1
                                                .TextMatrix(k, l) = mdlPublic.GetStdNodeText(objNode.Text)
                                            Next
                                        Next
                                        objReport.Items.Add intMaxID, mbytCurrFmt, "����" & intMaxID, intCurID, 8, i - X - 1, _
                                            "", 0, mdlPublic.GetStdNodeText(objNode.Text), "", 0, 0, 1000, 0, 255, 0, False, _
                                            "", 0, False, False, False, 0, 0, 0, False, 0, "", "", "", False, False, , _
                                            False, , , , "_" & intMaxID
                                    End If
                                Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt _
                                    , adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt _
                                    , adUnsignedTinyInt
                                    For k = .FixedCols To .Cols - 1 Step Z
                                        .TextMatrix(.FixedRows - 1, k + j) = "[" & mdlPublic.GetStdNodeText(objNode.Text) & "]"
                                    Next
                                    'ȱʡ�Ҷ���
                                    objReport.Items.Add intMaxID, mbytCurrFmt, "ͳ����" & intMaxID, intCurID, 9, j, _
                                        "", 0, mdlPublic.GetStdNodeText(objNode.Text), "", 0, 0, 1000, 0, 255, 2, _
                                        False, "", 0, False, False, False, 0, 0, 0, False, 0, "", "", "", False, _
                                        False, , False, , , , "_" & intMaxID
                                    j = j + 1 'jΪ���
                            End Select
                            tmpItem.SubIDs.Add intMaxID, "_" & intMaxID
                        End If
                        Set objNode = objNode.Next
                    Loop
                End With
                SetHeadCenter msh(intCurID)
                Set newObj = msh(intCurID)
            End If
        End If
    End If
    
    '��ʾ������Ŀ
    On Error Resume Next
    If Not objParent Is Nothing Then
        Set newObj.Container = objParent
        objReport.Items("_" & newObj.Index).��ID = objParent.Index
        For Each tmpID In objReport.Items("_" & newObj.Index).SubIDs
            With objReport.Items("_" & tmpID.id)
                .��ID = objParent.Index
            End With
        Next
    End If
    newObj.Left = selArea.Left
    newObj.Top = selArea.Top
    newObj.Width = objReport.Items("_" & newObj.Index).W * sgnMode
    newObj.Height = objReport.Items("_" & newObj.Index).H * sgnMode
    'newObj.Width = Abs(selArea.Right - selArea.Left)
    'newObj.Height = Abs(selArea.Bottom - selArea.Top)
    
    If UCase(TypeName(newObj)) = "LABEL" And objReport.Items("_" & newObj.Index).���� <> 1 Then
        newObj.AutoSize = objReport.Items("_" & newObj.Index).�Ե�
        objReport.Items("_" & newObj.Index).X = Format(newObj.Left / sgnMode, "0.00")
        objReport.Items("_" & newObj.Index).Y = Format(newObj.Top / sgnMode, "0.00")
        objReport.Items("_" & newObj.Index).W = Format(newObj.Width / sgnMode, "0.00")
        objReport.Items("_" & newObj.Index).H = Format(newObj.Height / sgnMode, "0.00")
    End If
    
    If objReport.Items("_" & newObj.Index).���� = 12 Then '@@@
        Call SetChartStyleAndData(newObj, objReport.Items("_" & newObj.Index), , sgnMode, True)
    Else
        newObj.FontSize = Format(newObj.FontSize * sgnMode, "0.0")
    End If
    
    '��ʼ����и�,�����Ժ��������弰�и�ʱ����(�м�,Loadʱһ������)��
    Select Case objReport.Items("_" & newObj.Index).����
        Case 4, 5
            Call InitRowHeight(newObj.Index)
            Call SetGridLine(newObj.Index)
            Call ReShowGrid(newObj.Index)
            objReport.Items("_" & newObj.Index).W = Format(newObj.Width / sgnMode, "0.00")
            objReport.Items("_" & newObj.Index).H = Format(newObj.Height / sgnMode, "0.00")
        Case 10
            newObj.BorderStyle = 1
            newObj.BackStyle = 0
            '����ǿ�����ͬ���޸�lblshp��λ��
            lblshp(newObj.Index).Left = newObj.Left
            lblshp(newObj.Index).Top = newObj.Top
            lblshp(newObj.Index).Width = newObj.Width
            lblshp(newObj.Index).Height = newObj.Height
            lblshp(newObj.Index).Visible = True
            lblshp(newObj.Index).ZOrder 1
            Call DrawFrame(newObj)
    End Select
    
    If objReport.Items("_" & newObj.Index).���� <> 10 Then newObj.ZOrder

    newObj.Visible = True
    
    If GetSelNum > 0 Then SelClear
    Call SelItem(newObj.Index, True)
    
    Call ShowAttrib(newObj.Index)
    
    '����Ԫ�غ�ȡ��Ԫ�ص�����
    If mnuFormat_Lock.Checked Then mnuFormat_Lock.Checked = False
    blnLock = False
    If mnuFormat_Lock.Checked And tbr2.Buttons("Lock").Value = tbrPressed Then
        tbr2.Buttons("Lock").Value = tbrUnpressed
    End If
    Call SetLock(blnLock)
    
    '����Ԫ�غ�Ĭ�ϲ�����Ԫ��
    If mnuFormat_Lock.Checked Then mnuFormat_Lock.Checked = False
    blnLock = False
    If Not mnuFormat_Lock.Checked And tbr2.Buttons("Lock").Value = tbrPressed Then
        tbr2.Buttons("Lock").Value = tbrUnpressed
    End If
    Call SetLock(blnLock)
    
    picPaper.SetFocus
End Sub

Private Function SelClear() As Integer
'����:���������ѡ�б���Ԫ�ص�ѡ��״̬
'����:���������Ŀ����
    Dim tmpObj As PictureBox
    
    For Each tmpObj In lblSize
        If tmpObj.Index <> 0 Then
            If tmpObj.Index Mod 8 = 1 Then '���б�־
                Select Case objReport.Items("_" & tmpObj.Tag).����
                    Case 1
                        lblLine(tmpObj.Tag).Tag = ""
                    Case 2, 3
                        lbl(tmpObj.Tag).Tag = ""
                    Case 10
                        Shp(tmpObj.Tag).Tag = ""
                    Case 4, 5
                        msh(tmpObj.Tag).Tag = ""
                        Call ResetColor(tmpObj.Tag)
                        Call SetGridLine(tmpObj.Tag)
                        If objReport.Items("_" & tmpObj.Tag).���� = 4 Then
                            Call CustomColColor(tmpObj.Tag, -9)
                            Call SetCopyGrid(tmpObj.Tag)
                        End If
                    Case 11
                        img(tmpObj.Tag).Tag = ""
                    Case 12 '@@@
                        Chart(tmpObj.Tag).Tag = ""
                    Case 13
                        ImgCode(tmpObj.Tag).Tag = ""
                    Case 14
                        pic(tmpObj.Tag).Tag = ""
                End Select
                SelClear = SelClear + 1
            End If
            Unload lblSize(tmpObj.Index)
        End If
    Next
    Set objLastSel = Nothing: intCurID = 0

    Call ShowAttrib
    
    selCell.Col1 = -1: selCell.Col2 = -1: selCell.Row1 = -1: selCell.Row2 = -1: selCell.Row = -1
    Call ShowPaperInfo
    intCurCol = -1
    
    picPaper.SetFocus
End Function

Private Sub SelMove(idx As Integer)
'���ܣ��ƶ�ѡ��Ԫ�صĵ㵽��ȷλ��
    Dim i As Integer
    Dim intBeginIdx As Integer
    Dim ObjSel As Control, tmpID As RelatID
    Dim blnUse(7) As Boolean '���Ƹ����ߴ��־�Ƿ���ʾ
    Dim lngTmp As Long
    Dim lngtmp1 As Long
    
    Select Case objReport.Items("_" & idx).����
        Case 2, 3 '�ı���ǩ�����ݱ�ǩ��ͬ���ؼ�
            Set ObjSel = lbl(idx)
        Case 10
            Set ObjSel = Shp(idx)
        Case 1 '����
            With lblLine(idx)
                If .Width > .Height Then
                    '����
                    For i = 0 To UBound(blnUse)
                        If Not (i = 2 Or i = 6) Then blnUse(i) = False
                    Next
                Else
                    '����
                    For i = 0 To UBound(blnUse)
                        If Not (i = 0 Or i = 4) Then blnUse(i) = False
                    Next
                End If
            End With
            Set ObjSel = lblLine(idx)
        Case 4, 5 '���ܱ��������ͬ���ؼ�
            Set ObjSel = msh(idx)
            '�����ؼ���ǰ(�����)
            For Each tmpID In objReport.Items("_" & idx).CopyIDs
                msh(tmpID.id).ZOrder
            Next
        Case 11
            Set ObjSel = img(idx)
        Case 12 '@@@
            Set ObjSel = Chart(idx)
        Case 13
            Set ObjSel = ImgCode(idx)
        Case 14
            Set ObjSel = pic(idx)
    End Select

    '����Ѿ�ѡ��,���ƶ�
    If Mid(ObjSel.Tag, 1, 2) = "S_" Then
        
        Set objLastSel = ObjSel: intCurID = idx
        
        
        '���¼���ʵ��λ�ü��������2002-03-26
        If ObjSel.Container.name <> "picPaper" Then
            lngTmp = ObjSel.Container.Top
            lngtmp1 = ObjSel.Container.Left
        End If
        
        intBeginIdx = CInt(Mid(ObjSel.Tag, 3))
        
         'Ԫ�ؿؼ���Tag��¼ѡ���־����ʼ����
        ObjSel.ZOrder IIF(objReport.Items("_" & ObjSel.Index).���� <> 10, 0, 1)
        If objReport.Items("_" & ObjSel.Index).���� = 10 Then lblshp(ObjSel.Index).ZOrder 1

        With WinProperty
            .H = lblSize(0).Height / Screen.TwipsPerPixelX
            .W = lblSize(0).Width / Screen.TwipsPerPixelX
        End With
        
        For i = intBeginIdx To intBeginIdx + 7 'ѡ���־��"����"��ʼ,"˳ʱ��"����
            Select Case IIF(i Mod 8 <> 0, i Mod 8, 8) '��λѡ��߿��λ��
                'zyb#Modify
                '����MoveWindow()���
                
                Case 1 '����
                    With WinProperty
                        .T = (ObjSel.Top + lngTmp - lblSize(i).Height) / Screen.TwipsPerPixelX
                        .l = (ObjSel.Left + lngtmp1 + (ObjSel.Width - lblSize(i).Width) / 2) / Screen.TwipsPerPixelX
                        Call MoveWindow(lblSize(i).hwnd, .l, .T, .W, .H, 1)
                    End With
                Case 2 '����
                    With WinProperty
                        .T = (ObjSel.Top + lngTmp - lblSize(i).Height) / Screen.TwipsPerPixelX
                        .l = (ObjSel.Left + lngtmp1 + ObjSel.Width) / Screen.TwipsPerPixelX
                        Call MoveWindow(lblSize(i).hwnd, .l, .T, .W, .H, 1)
                    End With
                Case 3 '����
                    With WinProperty
                        .T = (ObjSel.Top + lngTmp + (ObjSel.Height - lblSize(i).Height) / 2) / Screen.TwipsPerPixelX
                        .l = (ObjSel.Left + lngtmp1 + ObjSel.Width) / Screen.TwipsPerPixelX
                        Call MoveWindow(lblSize(i).hwnd, .l, .T, .W, .H, 1)
                    End With
                Case 4 '����
                    With WinProperty
                        .T = (ObjSel.Top + lngTmp + ObjSel.Height) / Screen.TwipsPerPixelX
                        .l = (ObjSel.Left + lngtmp1 + ObjSel.Width) / Screen.TwipsPerPixelX
                        Call MoveWindow(lblSize(i).hwnd, .l, .T, .W, .H, 1)
                    End With
                Case 5 '����
                    With WinProperty
                        .T = (ObjSel.Top + lngTmp + ObjSel.Height) / Screen.TwipsPerPixelX
                        .l = (ObjSel.Left + lngtmp1 + (ObjSel.Width - lblSize(i).Width) / 2) / Screen.TwipsPerPixelX
                        Call MoveWindow(lblSize(i).hwnd, .l, .T, .W, .H, 1)
                    End With
                Case 6 '����
                    With WinProperty
                        .T = (ObjSel.Top + lngTmp + ObjSel.Height) / Screen.TwipsPerPixelX
                        .l = (ObjSel.Left + lngtmp1 - lblSize(i).Width) / Screen.TwipsPerPixelX
                        Call MoveWindow(lblSize(i).hwnd, .l, .T, .W, .H, 1)
                    End With
                Case 7 '����
                    With WinProperty
                        .T = (ObjSel.Top + lngTmp + (ObjSel.Height - lblSize(i).Height) / 2) / Screen.TwipsPerPixelX
                        .l = (ObjSel.Left + lngtmp1 - lblSize(i).Width) / Screen.TwipsPerPixelX
                        Call MoveWindow(lblSize(i).hwnd, .l, .T, .W, .H, 1)
                    End With
                Case 8 '����
                    With WinProperty
                        .T = (ObjSel.Top + lngTmp - lblSize(i).Height) / Screen.TwipsPerPixelX
                        .l = (ObjSel.Left + lngtmp1 - lblSize(i).Width) / Screen.TwipsPerPixelX
                        Call MoveWindow(lblSize(i).hwnd, .l, .T, .W, .H, 1)
                    End With
            End Select
            lblSize(i).ZOrder
        Next
        
        Call SetSelFlag
        If GetSelNum > 1 Then Call SetSelFlag(intBeginIdx)
    End If
End Sub

Private Sub SelItem(idx As Integer, blnSel As Boolean)
'����:��ָ������Ԫ��ѡ��/��ѡ��
'����:idx=�ؼ�����,blnSel=�Ƿ�ѡ��
    Dim i As Integer
    Dim intBeginIdx As Integer
    Dim ObjSel As Control, tmpID As RelatID
    Dim blnUse(7) As Boolean '���Ƹ����ߴ��־�Ƿ���ʾ
    Dim lngTmp As Long
    Dim lngtmp1 As Long
    
    picPaper.SetFocus
    
    For i = 0 To UBound(blnUse)
        blnUse(i) = True
    Next
    
    Select Case objReport.Items("_" & idx).����
        Case 2, 3 '�ı���ǩ�����ݱ�ǩ��ͬ���ؼ�
            Set ObjSel = lbl(idx)
        Case 10
            Set ObjSel = Shp(idx)
        Case 1 '����
            With lblLine(idx)
                If .Width > .Height Then
                    '����
                    For i = 0 To UBound(blnUse)
                        If Not (i = 2 Or i = 6) Then blnUse(i) = False
                    Next
                Else
                    '����
                    For i = 0 To UBound(blnUse)
                        If Not (i = 0 Or i = 4) Then blnUse(i) = False
                    Next
                End If
            End With
            Set ObjSel = lblLine(idx)
        Case 4, 5 '���ܱ��������ͬ���ؼ�
            Set ObjSel = msh(idx)
            '�����ؼ���ǰ(�����)
            For Each tmpID In objReport.Items("_" & idx).CopyIDs
                msh(tmpID.id).ZOrder
            Next
        Case 11
            Set ObjSel = img(idx)
        Case 12 '@@@
            Set ObjSel = Chart(idx)
        Case 13
            Set ObjSel = ImgCode(idx)
        Case 14
            Set ObjSel = pic(idx)
    End Select

    If blnSel Then
        '����Ѿ�ѡ��,�����ظ�����
        If Mid(ObjSel.Tag, 1, 2) = "S_" Then Exit Sub
        
        Set objLastSel = ObjSel: intCurID = idx
        
        
        '���¼���ʵ��λ�ü��������2002-03-26
        If ObjSel.Container.name <> "picPaper" Then
            lngTmp = ObjSel.Container.Top
            lngtmp1 = ObjSel.Container.Left
        End If

        
        intBeginIdx = lblSize.UBound + 1 '����Ϊ(1n-8n),n>0
        
         'Ԫ�ؿؼ���Tag��¼ѡ���־����ʼ����
        ObjSel.Tag = "S_" & intBeginIdx
        ObjSel.ZOrder IIF(objReport.Items("_" & ObjSel.Index).���� <> 10, 0, 1)
        If objReport.Items("_" & ObjSel.Index).���� = 10 Then lblshp(ObjSel.Index).ZOrder 1

        With WinProperty
            .H = lblSize(0).Height / Screen.TwipsPerPixelX
            .W = lblSize(0).Width / Screen.TwipsPerPixelX
        End With
        
        For i = intBeginIdx To intBeginIdx + 7 'ѡ���־��"����"��ʼ,"˳ʱ��"����
            Load lblSize(i)
            Select Case IIF(i Mod 8 <> 0, i Mod 8, 8) '��λѡ��߿��λ��
                'zyb#Modify
                '����MoveWindow()���
                
                Case 1 '����
                    lblSize(i).Tag = idx '��һ��(����)��¼��Ӧ�ؼ�������
                    With WinProperty
                        .T = (ObjSel.Top + lngTmp - lblSize(i).Height) / Screen.TwipsPerPixelX
                        .l = (ObjSel.Left + lngtmp1 + (ObjSel.Width - lblSize(i).Width) / 2) / Screen.TwipsPerPixelX
                        Call MoveWindow(lblSize(i).hwnd, .l, .T, .W, .H, 1)
                    End With
                    lblSize(i).MousePointer = 7
                Case 2 '����
                    lblSize(i).Tag = ObjSel.name '�ڶ���(����)��¼��Ӧ�ؼ�������
                    With WinProperty
                        .T = (ObjSel.Top + lngTmp - lblSize(i).Height) / Screen.TwipsPerPixelX
                        .l = (ObjSel.Left + lngtmp1 + ObjSel.Width) / Screen.TwipsPerPixelX
                        Call MoveWindow(lblSize(i).hwnd, .l, .T, .W, .H, 1)
                    End With
                    lblSize(i).MousePointer = 6
                Case 3 '����
                    With WinProperty
                        .T = (ObjSel.Top + lngTmp + (ObjSel.Height - lblSize(i).Height) / 2) / Screen.TwipsPerPixelX
                        .l = (ObjSel.Left + lngtmp1 + ObjSel.Width) / Screen.TwipsPerPixelX
                        Call MoveWindow(lblSize(i).hwnd, .l, .T, .W, .H, 1)
                    End With
                    lblSize(i).MousePointer = 9
                Case 4 '����
                    With WinProperty
                        .T = (ObjSel.Top + lngTmp + ObjSel.Height) / Screen.TwipsPerPixelX
                        .l = (ObjSel.Left + lngtmp1 + ObjSel.Width) / Screen.TwipsPerPixelX
                        Call MoveWindow(lblSize(i).hwnd, .l, .T, .W, .H, 1)
                    End With
                    lblSize(i).MousePointer = 8
                Case 5 '����
                    With WinProperty
                        .T = (ObjSel.Top + lngTmp + ObjSel.Height) / Screen.TwipsPerPixelX
                        .l = (ObjSel.Left + lngtmp1 + (ObjSel.Width - lblSize(i).Width) / 2) / Screen.TwipsPerPixelX
                        Call MoveWindow(lblSize(i).hwnd, .l, .T, .W, .H, 1)
                    End With
                    lblSize(i).MousePointer = 7
                Case 6 '����
                    With WinProperty
                        .T = (ObjSel.Top + lngTmp + ObjSel.Height) / Screen.TwipsPerPixelX
                        .l = (ObjSel.Left + lngtmp1 - lblSize(i).Width) / Screen.TwipsPerPixelX
                        Call MoveWindow(lblSize(i).hwnd, .l, .T, .W, .H, 1)
                    End With
                    lblSize(i).MousePointer = 6
                Case 7 '����
                    With WinProperty
                        .T = (ObjSel.Top + lngTmp + (ObjSel.Height - lblSize(i).Height) / 2) / Screen.TwipsPerPixelX
                        .l = (ObjSel.Left + lngtmp1 - lblSize(i).Width) / Screen.TwipsPerPixelX
                        Call MoveWindow(lblSize(i).hwnd, .l, .T, .W, .H, 1)
                    End With
                    lblSize(i).MousePointer = 9
                Case 8 '����
                    With WinProperty
                        .T = (ObjSel.Top + lngTmp - lblSize(i).Height) / Screen.TwipsPerPixelX
                        .l = (ObjSel.Left + lngtmp1 - lblSize(i).Width) / Screen.TwipsPerPixelX
                        Call MoveWindow(lblSize(i).hwnd, .l, .T, .W, .H, 1)
                    End With
                    lblSize(i).MousePointer = 8
            End Select
            lblSize(i).ZOrder
            lblSize(i).Visible = blnUse(IIF(i Mod 8 <> 0, i Mod 8, 8) - 1)
        Next
        
        Call SetSelFlag
        If GetSelNum > 1 Then Call SetSelFlag(intBeginIdx)
    Else
        If Trim(ObjSel.Tag) = "" Then Exit Sub '�Ѵ��ڷ�ѡ��״̬
        
        If objReport.Items("_" & ObjSel.Index).���� = 4 Or objReport.Items("_" & ObjSel.Index).���� = 5 Then
            selCell.Col1 = -1: selCell.Col2 = -1: selCell.Row1 = -1: selCell.Row2 = -1: selCell.Row = -1
            Call ShowPaperInfo
            intCurCol = -1
        End If
        
        ObjSel.ZOrder IIF(objReport.Items("_" & ObjSel.Index).���� <> 10, 0, 1)
        intBeginIdx = CInt(Mid(ObjSel.Tag, 3))
        ObjSel.Tag = ""
        For i = intBeginIdx To intBeginIdx + 7
            Unload lblSize(i)
        Next
        
        Call SetSelFlag
        
        If GetSelNum > 0 Then
            If GetSelNum > 1 Then Call SetSelFlag(lblSize.UBound - 7)
            Select Case objReport.Items("_" & CInt(lblSize(lblSize.UBound - 7).Tag)).����
                Case 2, 3
                    Set objLastSel = lbl(CInt(lblSize(lblSize.UBound - 7).Tag))
                Case 10
                    Set objLastSel = Shp(CInt(lblSize(lblSize.UBound - 7).Tag))
                Case 1
                    Set objLastSel = lblLine(CInt(lblSize(lblSize.UBound - 7).Tag))
                Case 4, 5
                    Set objLastSel = msh(CInt(lblSize(lblSize.UBound - 7).Tag))
                Case 11
                    Set objLastSel = img(CInt(lblSize(lblSize.UBound - 7).Tag))
                Case 12 '@@@
                    Set objLastSel = Chart(CInt(lblSize(lblSize.UBound - 7).Tag))
                Case 13
                    Set objLastSel = ImgCode(CInt(lblSize(lblSize.UBound - 7).Tag))
                Case 14
                    Set objLastSel = pic(CInt(lblSize(lblSize.UBound - 7).Tag))
            End Select
            intCurID = CInt(lblSize(lblSize.UBound - 7).Tag)
        Else
            Set objLastSel = Nothing: intCurID = 0
        End If
    End If
End Sub

Private Sub SetSelFlag(Optional intBegin As Integer = 0)
'����:�趨���һ��ѡ�е���Ŀ���Ʊ�־��ǰ�治ͬ,��ָ����б�־Ϊ����ɫ
'������intBegin=(ָ���ؼ�Ԫ��)�ߴ��־��ʼ����,Ϊ0��ʾ�ָ����б�־Ϊ����ɫ
    Dim tmpObj As PictureBox
    Dim i As Integer
    If intBegin = 0 Then
        For Each tmpObj In lblSize
            If tmpObj.Index <> 0 Then
                If blnLock Then
                    tmpObj.BackColor = &HFF
                Else
                    tmpObj.BackColor = &HFF0000
                End If
            End If
        Next
    Else
        For i = intBegin To intBegin + 7
            lblSize(i).BackColor = &HC000&
        Next
    End If
End Sub

Private Function GetSelNum() As Integer
'���ܣ����ص�ǰѡ��Ԫ�ؿؼ�����
'˵��������ѡ��ؼ�һ�����ڳߴ��־
    Dim tmpObj As PictureBox, i As Integer
    For Each tmpObj In lblSize
        If tmpObj.Index <> 0 And tmpObj.Index Mod 8 = 0 Then i = i + 1
    Next
    GetSelNum = i
End Function

Private Function SelAreaItem(Area As RECT) As Integer
'���ܣ�ѡ����ָ�㶨�����ڵı���Ԫ�ؿؼ�(�����������ѡ���)
'���أ�ѡ�еĸ���
    Dim tmpItem As RPTItem, ObjSel As Object
    Dim lngLeft As Long, lngTop As Long
    
    For Each tmpItem In objReport.Items '@@@
        If InStr(1, "|1,|2,|3,|4,|5,|10,|11,|12,|13,|14,", "|" & tmpItem.����) <> 0 _
            And Mid(cboFormat.ComboItems("_" & mbytCurrFmt).Key, 2) = tmpItem.��ʽ�� Then
            Set ObjSel = GetInxObj(tmpItem.id)
            
            If tmpItem.���� <> 14 And tmpItem.��ID <> 0 Then
                lngLeft = pic(tmpItem.��ID).Left
                lngTop = pic(tmpItem.��ID).Top
            Else
                lngLeft = 0
                lngTop = 0
            End If
            If Not (ObjSel.Top + lngTop > Area.Bottom Or _
                ObjSel.Left + lngLeft > Area.Right Or _
                ObjSel.Top + lngTop + ObjSel.Height < Area.Top Or _
                ObjSel.Left + lngLeft + ObjSel.Width < Area.Left) Then
                Call SelItem(ObjSel.Index, True)
                SelAreaItem = SelAreaItem + 1
            End If
        End If
    Next
End Function

Private Sub SetLock(blnLock As Boolean)
'����:���ݿؼ��Ƿ������趨�ؼ�ѡ���߿���ɫ
    Dim tmpObj As PictureBox

    For Each tmpObj In lblSize
        If tmpObj.Index <> 0 Then
            If tmpObj.BackColor <> &HC000& Then
                If blnLock Then
                    tmpObj.BackColor = &HFF
                Else
                    tmpObj.BackColor = &HFF0000
                End If
            End If
        End If
    Next
End Sub

Private Sub ShowAttrib(Optional idx As Integer, Optional blnBase As Boolean = False)
'���ܣ���ʾָ�������ؼ���Ӧ�Ķ������Ի��������
'������idx=�ؼ�����,Ϊ0ʱ��ʾ������Կ�
    Dim intType As Integer, RowH As Single, i As Integer
    Dim arrHead As Variant, arrRowH As Variant
    Dim tmpID As RelatID, tmpItem As RPTItem
    Dim objFmt As RPTFmt
    Dim lngRow As Long, lngCol As Long
    Dim tmpObj As PictureBox
    Dim lngType As Long
    
    Set objFmt = objReport.Fmts("_" & mbytCurrFmt)
    
    With mshAtt
        lngRow = .Row: lngCol = .Col
        
        .Redraw = False
        .Clear
        .Rows = 2
        .TextMatrix(0, 0) = "��Ŀ": .TextMatrix(0, 1) = "����"
        .Row = 0: .Col = 0: .CellAlignment = 4: .Col = 1: .CellAlignment = 4
        .ColAlignment(0) = 1: .ColAlignment(1) = 1
        .ExtendLastCol = True
        .ColWidth(0) = 1200
        '.ColWidth(1) = 1200
        lblNote.Caption = ""
        If idx = 0 Then
            '�����ѡ�е�����Ԫ�ض�����ͬ��Ԫ��
            For Each tmpObj In lblSize
                If tmpObj.Index Mod 8 = 1 Then
                    If lngType <> 0 And lngType <> objReport.Items("_" & tmpObj.Tag).���� Then lngType = 0: idx = 0: Exit For
                    lngType = objReport.Items("_" & tmpObj.Tag).����
                    idx = objReport.Items("_" & tmpObj.Tag).id
                End If
            Next
        End If
        If idx <> 0 Then
            '����Ǳ�ǩ�����ֶ�Ϊͼ�ͣ�����ͼ�ʹ���
            intType = objReport.Items("_" & idx).����
            If ItemIsGraph(idx) Then intType = 11
            Select Case intType
                Case 1, 10 '����,����
                    .Rows = 8
                    .TextMatrix(1, 0) = "����": .TextMatrix(1, 1) = IIF(objReport.Items("_" & idx).���� = 1, "����", "����")
                    .TextMatrix(2, 0) = "����": .TextMatrix(2, 1) = objReport.Items("_" & idx).����
                    .TextMatrix(3, 0) = "X����": .TextMatrix(3, 1) = Format(objReport.Items("_" & idx).X / Twip_mm, "0.00")
                    .TextMatrix(4, 0) = "Y����": .TextMatrix(4, 1) = Format(objReport.Items("_" & idx).Y / Twip_mm, "0.00")
                    If objReport.Items("_" & idx).���� = 1 Then
                        If objReport.Items("_" & idx).W > objReport.Items("_" & idx).H Then .TextMatrix(5, 0) = "���": .TextMatrix(5, 1) = Format(objReport.Items("_" & idx).W / Twip_mm, "0.00")
                        If objReport.Items("_" & idx).W < objReport.Items("_" & idx).H Then .TextMatrix(5, 0) = "�߶�": .TextMatrix(5, 1) = Format(objReport.Items("_" & idx).H / Twip_mm, "0.00")
                        .TextMatrix(6, 0) = "ǰ��ɫ": .TextMatrix(6, 1) = "��": .Row = 6: .Col = 1: .CellForeColor = objReport.Items("_" & idx).ǰ��
                    Else
                        .TextMatrix(5, 0) = "�߶�": .TextMatrix(5, 1) = Format(objReport.Items("_" & idx).H / Twip_mm, "0.00")
                        .TextMatrix(6, 0) = "���": .TextMatrix(6, 1) = Format(objReport.Items("_" & idx).W / Twip_mm, "0.00")
                    End If
                    .TextMatrix(7, 0) = "�Ӵ�": .TextMatrix(7, 1) = IIF(objReport.Items("_" & idx).����, "��", "��")
                    If intType = 10 Then
                        .Rows = .Rows + 1
                        .TextMatrix(.Rows - 1, 0) = "��״": .TextMatrix(.Rows - 1, 1) = IIF(objReport.Items("_" & idx).�߿� = 0, "����", "Բ��")
                    End If
                Case 2 '��ǩ
                    Dim str���� As String
                    Dim str�������� As String
                    .Rows = 22
                    .TextMatrix(1, 0) = "����": .TextMatrix(1, 1) = "��ǩ"
                    .TextMatrix(2, 0) = "����": .TextMatrix(2, 1) = objReport.Items("_" & idx).����
                    .TextMatrix(3, 0) = "����": .TextMatrix(3, 1) = objReport.Items("_" & idx).����
                    .TextMatrix(4, 0) = "����": .TextMatrix(4, 1) = IIF(objReport.Items("_" & idx).���� = 0, "�����", IIF(objReport.Items("_" & idx).���� = 1, "�ж���", "�Ҷ���"))
                    .TextMatrix(5, 0) = "X����": .TextMatrix(5, 1) = Format(objReport.Items("_" & idx).X / Twip_mm, "0.00")
                    .TextMatrix(6, 0) = "Y����": .TextMatrix(6, 1) = Format(objReport.Items("_" & idx).Y / Twip_mm, "0.00")
                    .TextMatrix(7, 0) = "���": .TextMatrix(7, 1) = Format(objReport.Items("_" & idx).W / Twip_mm, "0.00")
                    .TextMatrix(8, 0) = "�߶�": .TextMatrix(8, 1) = Format(objReport.Items("_" & idx).H / Twip_mm, "0.00")
                    .TextMatrix(9, 0) = "�Զ�������С": .TextMatrix(9, 1) = IIF(objReport.Items("_" & idx).�Ե�, "��", "��")
                    .TextMatrix(10, 0) = "����": .TextMatrix(10, 1) = objReport.Items("_" & idx).����
                    .TextMatrix(11, 0) = "�Զ�����": .TextMatrix(11, 1) = IIF(objReport.Items("_" & idx).�и� = 1, "��", "��")
                    .TextMatrix(12, 0) = "�߿�": .TextMatrix(12, 1) = IIF(objReport.Items("_" & idx).�߿�, "��", "��")
                    .TextMatrix(13, 0) = "����ɫ": .TextMatrix(13, 1) = "��": .Row = 13: .Col = 1: .CellForeColor = objReport.Items("_" & idx).����
                    .TextMatrix(14, 0) = "���ն���": .TextMatrix(14, 1) = objReport.Items("_" & idx).����
                    str���� = objReport.Items("_" & idx).����
                    If str���� = "0" Or str���� = "" Then
                        str���� = "����"
                        .TextMatrix(15, 0) = "����": .TextMatrix(15, 1) = "����"
                    Else
                        str���� = Mid(str����, 2)
                        str���� = IIF(str���� = "1", "����", IIF(str���� = "2", "����", "����"))
                        .TextMatrix(15, 0) = "����": .TextMatrix(15, 1) = IIF(Mid(objReport.Items("_" & idx).����, 1, 1) = "1", "������", "������")
                    End If
                    .TextMatrix(16, 0) = "����": .TextMatrix(16, 1) = str����
                    .TextMatrix(17, 0) = "��ʽ": .TextMatrix(17, 1) = objReport.Items("_" & idx).��ʽ
                    .TextMatrix(18, 0) = "�о�": .TextMatrix(18, 1) = objReport.Items("_" & idx).����
                    .TextMatrix(19, 0) = "����Դ�к�": .TextMatrix(19, 1) = objReport.Items("_" & idx).Դ�к�
                    For i = 1 To objReport.Items("_" & idx).Relations.count
                        If InStr(str��������, "," & objReport.Items("_" & idx).Relations.Item(i).������������) = 0 Then
                            str�������� = str�������� & "," & objReport.Items("_" & idx).Relations.Item(i).������������
                        End If
                    Next
                    If str�������� <> "" Then
                        str�������� = Mid(str��������, 2)
                    End If
                    .TextMatrix(20, 0) = "��������": .TextMatrix(20, 1) = str��������
                    .TextMatrix(21, 0) = "ˮƽ��ת": .TextMatrix(21, 1) = IIF(objReport.Items("_" & idx).ˮƽ��ת, "��", "��")
                Case 4 '������
                    .Rows = 18
                    '��Ϊ������������Դ�ڶ������Դ,�������ݲ���ȷ��
                    .TextMatrix(1, 0) = "����": .TextMatrix(1, 1) = "������"
                    .TextMatrix(2, 0) = "����": .TextMatrix(2, 1) = objReport.Items("_" & idx).����
                    .TextMatrix(3, 0) = "X����": .TextMatrix(3, 1) = Format(objReport.Items("_" & idx).X / Twip_mm, "0.00")
                    .TextMatrix(4, 0) = "Y����": .TextMatrix(4, 1) = Format(objReport.Items("_" & idx).Y / Twip_mm, "0.00")
                    .TextMatrix(5, 0) = "���": .TextMatrix(5, 1) = Format(objReport.Items("_" & idx).W / Twip_mm, "0.00")
                    .TextMatrix(6, 0) = "�߶�": .TextMatrix(6, 1) = Format(objReport.Items("_" & idx).H / Twip_mm, "0.00")
                    If msh(idx).Row < msh(idx).FixedRows Then
                        RowH = 0
                        arrHead = Split(objReport.Items("_" & objReport.Items("_" & idx).SubIDs(1).Key).��ͷ, "|")
                        If selCell.Row <> -1 Then
                            For i = selCell.Row1 To selCell.Row2
                                arrRowH = Split(arrHead(i), "^")
                                On Error Resume Next
                                Err = 0
                                arrRowH(1) = arrRowH(1) + 0
                                If Err <> 0 Then
                                    RowH = RowH + 255
                                Else
                                    RowH = RowH + arrRowH(1)
                                End If
                            Next
                        Else
                            RowH = 255
                        End If
                        On Error GoTo 0
                        .TextMatrix(7, 0) = "�и�": .TextMatrix(7, 1) = Format(RowH / Twip_mm, "0.00")
                    Else
                        .TextMatrix(7, 0) = "�и�": .TextMatrix(7, 1) = Format(objReport.Items("_" & idx).�и� / Twip_mm, "0.00")
                    End If
                    .TextMatrix(8, 0) = "����": .TextMatrix(8, 1) = objReport.Items("_" & idx).����
                    .TextMatrix(9, 0) = "����ɫ": .TextMatrix(9, 1) = "��": .Row = 9: .Col = 1: .CellForeColor = objReport.Items("_" & idx).����
                    .TextMatrix(10, 0) = "��������ɫ": .TextMatrix(10, 1) = "��": .Row = 10: .Col = 1: .CellForeColor = objReport.Items("_" & idx).����
                    .TextMatrix(11, 0) = "��ͷ����ɫ": .TextMatrix(11, 1) = "��": .Row = 11: .Col = 1: .CellForeColor = IIF(objReport.Items("_" & idx).��ʽ = "", objReport.Items("_" & idx).����, Val(objReport.Items("_" & idx).��ʽ))
                    .TextMatrix(12, 0) = "���ն���": .TextMatrix(12, 1) = objReport.Items("_" & idx).����
                    .TextMatrix(13, 0) = "����": .TextMatrix(13, 1) = IIF(objReport.Items("_" & idx).���� = "", "����", IIF(objReport.Items("_" & idx).���� = "1", "����", "������"))
                    .TextMatrix(14, 0) = "����": .TextMatrix(14, 1) = IIF(objReport.Items("_" & idx).���� <= 1, 1, objReport.Items("_" & idx).����)
                    .TextMatrix(15, 0) = "����": .TextMatrix(15, 1) = IIF(objReport.Items("_" & idx).�߿�, "��", "��")
                    .TextMatrix(16, 0) = "����": .TextMatrix(16, 1) = IIF(objReport.Items("_" & idx).�Ե�, "��", "��")
                    .TextMatrix(17, 0) = "����߼Ӵ�": .TextMatrix(17, 1) = IIF(objReport.Items("_" & idx).����߼Ӵ�, "��", "��")
                Case 5 '���ܱ��
                    .Rows = 17
                    '���ܱ���ܷ���
                    .TextMatrix(1, 0) = "����": .TextMatrix(1, 1) = "���ܱ��"
                    .TextMatrix(2, 0) = "����": .TextMatrix(2, 1) = objReport.Items("_" & idx).����
                    .TextMatrix(3, 0) = "����": .TextMatrix(3, 1) = objReport.Items("_" & idx).����
                    .TextMatrix(4, 0) = "X����": .TextMatrix(4, 1) = Format(objReport.Items("_" & idx).X / Twip_mm, "0.00")
                    .TextMatrix(5, 0) = "Y����": .TextMatrix(5, 1) = Format(objReport.Items("_" & idx).Y / Twip_mm, "0.00")
                    .TextMatrix(6, 0) = "���": .TextMatrix(6, 1) = Format(objReport.Items("_" & idx).W / Twip_mm, "0.00")
                    .TextMatrix(7, 0) = "�߶�": .TextMatrix(7, 1) = Format(objReport.Items("_" & idx).H / Twip_mm, "0.00")
                    .TextMatrix(8, 0) = "�и�": .TextMatrix(8, 1) = Format(objReport.Items("_" & idx).�и� / Twip_mm, "0.00")
                    .TextMatrix(9, 0) = "����": .TextMatrix(9, 1) = objReport.Items("_" & idx).����
                    .TextMatrix(10, 0) = "����ɫ": .TextMatrix(10, 1) = "��": .Row = 10: .Col = 1: .CellForeColor = objReport.Items("_" & idx).����
                    .TextMatrix(11, 0) = "����ɫ": .TextMatrix(11, 1) = "��": .Row = 11: .Col = 1: .CellForeColor = objReport.Items("_" & idx).����
                    .TextMatrix(12, 0) = "���ն���": .TextMatrix(12, 1) = objReport.Items("_" & idx).����
                    .TextMatrix(13, 0) = "����": .TextMatrix(13, 1) = IIF(objReport.Items("_" & idx).���� = "", "����", IIF(objReport.Items("_" & idx).���� = "1", "����", "������"))
                    .TextMatrix(14, 0) = "����": .TextMatrix(14, 1) = IIF(objReport.Items("_" & idx).�߿�, "��", "��")
                    .TextMatrix(15, 0) = "����": .TextMatrix(15, 1) = IIF(objReport.Items("_" & idx).�Ե�, "��", "��")
                    .TextMatrix(16, 0) = "����߼Ӵ�": .TextMatrix(16, 1) = IIF(objReport.Items("_" & idx).����߼Ӵ�, "��", "��")
                Case 11 'ͼƬ
                    If objReport.Items("_" & idx).���� = 11 Then
                        .Rows = 12 '���������ͼƬԪ��
                    Else
                        .Rows = 11
                    End If
                    .TextMatrix(1, 0) = "����": .TextMatrix(1, 1) = "ͼƬ"
                    .TextMatrix(2, 0) = "����": .TextMatrix(2, 1) = objReport.Items("_" & idx).����
                    If objReport.Items("_" & idx).���� = 2 Then
                        .TextMatrix(3, 0) = "����": .TextMatrix(3, 1) = objReport.Items("_" & idx).����
                    Else
                        .TextMatrix(3, 0) = "����": .TextMatrix(3, 1) = IIF(Not objReport.Items("_" & idx).ͼƬ Is Nothing, "[Pictrue]", "")
                    End If
                    .TextMatrix(4, 0) = "X����": .TextMatrix(4, 1) = Format(objReport.Items("_" & idx).X / Twip_mm, "0.00")
                    .TextMatrix(5, 0) = "Y����": .TextMatrix(5, 1) = Format(objReport.Items("_" & idx).Y / Twip_mm, "0.00")
                    .TextMatrix(6, 0) = "���": .TextMatrix(6, 1) = Format(objReport.Items("_" & idx).W / Twip_mm, "0.00")
                    .TextMatrix(7, 0) = "�߶�": .TextMatrix(7, 1) = Format(objReport.Items("_" & idx).H / Twip_mm, "0.00")
                    .TextMatrix(8, 0) = "�߿�": .TextMatrix(8, 1) = IIF(objReport.Items("_" & idx).�߿�, "��", "��")
                    .TextMatrix(9, 0) = "���ֱ���": .TextMatrix(9, 1) = IIF(objReport.Items("_" & idx).����, "��", "��")
                    .TextMatrix(10, 0) = "�Զ�������С": .TextMatrix(10, 1) = IIF(objReport.Items("_" & idx).�Ե�, "��", "��")
                    If objReport.Items("_" & idx).���� = 11 Then
                        .TextMatrix(11, 0) = "����ͼ��": .TextMatrix(11, 1) = IIF(objReport.Items("_" & idx).����, "��", "��")
                    End If
                Case 14 '��Ƭ
                    .Rows = 13
                    .TextMatrix(1, 0) = "����": .TextMatrix(1, 1) = "��Ƭ"
                    .TextMatrix(2, 0) = "����": .TextMatrix(2, 1) = objReport.Items("_" & idx).����
                    .TextMatrix(3, 0) = "X����": .TextMatrix(3, 1) = Format(objReport.Items("_" & idx).X / Twip_mm, "0.00")
                    .TextMatrix(4, 0) = "Y����": .TextMatrix(4, 1) = Format(objReport.Items("_" & idx).Y / Twip_mm, "0.00")
                    .TextMatrix(5, 0) = "���": .TextMatrix(5, 1) = Format(objReport.Items("_" & idx).W / Twip_mm, "0.00")
                    .TextMatrix(6, 0) = "�߶�": .TextMatrix(6, 1) = Format(objReport.Items("_" & idx).H / Twip_mm, "0.00")
                    .TextMatrix(7, 0) = "�߿�": .TextMatrix(7, 1) = IIF(objReport.Items("_" & idx).�߿�, "��", "��")
                    .TextMatrix(8, 0) = "����Դ": .TextMatrix(8, 1) = objReport.Items("_" & idx).����Դ
                    .TextMatrix(9, 0) = "���¼��": .TextMatrix(9, 1) = Format(objReport.Items("_" & idx).���¼�� / Twip_mm, "0.00")
                    .TextMatrix(10, 0) = "���Ҽ��": .TextMatrix(10, 1) = Format(objReport.Items("_" & idx).���Ҽ�� / Twip_mm, "0.00")
                    .TextMatrix(11, 0) = "�������": .TextMatrix(11, 1) = objReport.Items("_" & idx).�������
                    .TextMatrix(12, 0) = "�������": .TextMatrix(12, 1) = objReport.Items("_" & idx).�������
                Case 12 'ͼ��@@@
                    .Rows = 11
                    .TextMatrix(1, 0) = "����": .TextMatrix(1, 1) = "ͼ��"
                    .TextMatrix(2, 0) = "����": .TextMatrix(2, 1) = objReport.Items("_" & idx).����
                    .TextMatrix(3, 0) = "����": .TextMatrix(3, 1) = ""
                    .TextMatrix(4, 0) = "X����": .TextMatrix(4, 1) = Format(objReport.Items("_" & idx).X / Twip_mm, "0.00")
                    .TextMatrix(5, 0) = "Y����": .TextMatrix(5, 1) = Format(objReport.Items("_" & idx).Y / Twip_mm, "0.00")
                    .TextMatrix(6, 0) = "���": .TextMatrix(6, 1) = Format(objReport.Items("_" & idx).W / Twip_mm, "0.00")
                    .TextMatrix(7, 0) = "�߶�": .TextMatrix(7, 1) = Format(objReport.Items("_" & idx).H / Twip_mm, "0.00")
                    .TextMatrix(8, 0) = "����": .TextMatrix(8, 1) = objReport.Items("_" & idx).����
                    .TextMatrix(9, 0) = "ǰ��ɫ": .TextMatrix(9, 1) = "��": .Row = 9: .Col = 1: .CellForeColor = objReport.Items("_" & idx).ǰ��
                    .TextMatrix(10, 0) = "����ɫ": .TextMatrix(10, 1) = "��": .Row = 10: .Col = 1: .CellForeColor = objReport.Items("_" & idx).����
                Case 13 '����
                    If objReport.Items("_" & idx).��� = 1 Then 'Code 128(����)
                        .Rows = 12
                    ElseIf objReport.Items("_" & idx).��� = 2 Then 'Code 39
                        .Rows = 13
                    ElseIf objReport.Items("_" & idx).��� = 3 Then 'Code 128 Auto
                        .Rows = 13
                    ElseIf objReport.Items("_" & idx).��� = 10 Then 'QR Code
                        .Rows = 11
                    End If
                    
                    .TextMatrix(1, 0) = "����": .TextMatrix(1, 1) = "����"
                    .TextMatrix(2, 0) = "����": .TextMatrix(2, 1) = objReport.Items("_" & idx).����
                    .TextMatrix(3, 0) = "����": .TextMatrix(3, 1) = objReport.Items("_" & idx).����
                    .TextMatrix(4, 0) = "X����": .TextMatrix(4, 1) = Format(objReport.Items("_" & idx).X / Twip_mm, "0.00")
                    .TextMatrix(5, 0) = "Y����": .TextMatrix(5, 1) = Format(objReport.Items("_" & idx).Y / Twip_mm, "0.00")
                    .TextMatrix(6, 0) = "���": .TextMatrix(6, 1) = Format(objReport.Items("_" & idx).W / Twip_mm, "0.00")
                    .TextMatrix(7, 0) = "�߶�": .TextMatrix(7, 1) = Format(objReport.Items("_" & idx).H / Twip_mm, "0.00")
                    .TextMatrix(8, 0) = "��������": .TextMatrix(8, 1) = Decode(objReport.Items("_" & idx).���, 1, "Code 128(����)", 3, "Code 128 Auto", 2, "Code 39", 10, "QR Code", "")
                    .TextMatrix(9, 0) = "����Դ�к�": .TextMatrix(9, 1) = objReport.Items("_" & idx).Դ�к�
                    If objReport.Items("_" & idx).��� = 1 Then
                        .TextMatrix(10, 0) = "��ʾ����": .TextMatrix(10, 1) = IIF(Mid(objReport.Items("_" & idx).��ͷ, 1, 1) = "1", "��", "��")
                        .TextMatrix(11, 0) = "��ת����": .TextMatrix(11, 1) = Decode(Val(Mid(objReport.Items("_" & idx).��ͷ, 3, 1)), 0, "����ת", 1, "˳ʱ��90��", 2, "��ʱ��90��", "")
                    ElseIf objReport.Items("_" & idx).��� = 2 Then
                        .TextMatrix(10, 0) = "��ʾ����": .TextMatrix(10, 1) = IIF(Mid(objReport.Items("_" & idx).��ͷ, 1, 1) = "1", "��", "��")
                        .TextMatrix(11, 0) = "��У���": .TextMatrix(11, 1) = IIF(Mid(objReport.Items("_" & idx).��ͷ, 2, 1) = "1", "��", "��")
                        .TextMatrix(12, 0) = "��ת����": .TextMatrix(12, 1) = Decode(Val(Mid(objReport.Items("_" & idx).��ͷ, 3, 1)), 0, "����ת", 1, "˳ʱ��90��", 2, "��ʱ��90��", "")
                    ElseIf objReport.Items("_" & idx).��� = 3 Then
                        .TextMatrix(10, 0) = "�����߿�": .TextMatrix(10, 1) = objReport.Items("_" & idx).�и�
                        .TextMatrix(11, 0) = "��ʾ����": .TextMatrix(11, 1) = IIF(Mid(objReport.Items("_" & idx).��ͷ, 1, 1) = "1", "��", "��")
                        .TextMatrix(12, 0) = "��ת����": .TextMatrix(12, 1) = Decode(Val(Mid(objReport.Items("_" & idx).��ͷ, 3, 1)), 0, "����ת", 1, "˳ʱ��90��", 2, "��ʱ��90��", "")
                    ElseIf objReport.Items("_" & idx).��� = 10 Then
                        .TextMatrix(10, 0) = "�Զ�������С": .TextMatrix(10, 1) = IIF(objReport.Items("_" & idx).�Ե�, "��", "��")
                    End If
            End Select
            If intType <> 12 And intType <> 14 And intType <> 5 Then
                '�������Ƿ�����ڿ�Ƭ��
                 .AddItem ""
                 .TextMatrix(.Rows - 1, 0) = "����"
                 If objReport.Items("_" & idx).��ID = 0 Then
                    .TextMatrix(.Rows - 1, 1) = "ҳ��"
                 Else
                    .TextMatrix(.Rows - 1, 1) = objReport.Items("_" & objReport.Items("_" & idx).��ID).����
                 End If
            End If
            '����Ƕ�ѡ����ֻ��ʾ������Ŀ
            If lngType <> 0 Then
                For i = 1 To .Rows - 1
                    If InStr("ǰ��ɫ;����ɫ;����;�Զ�������С;����;�Զ�����;�߿�", .TextMatrix(i, 0)) = 0 Then .RowHidden(i) = True
                Next
            End If
        End If
        
        '��ʾֽ�ŵȻ�������
        If blnBase Then
            '�����ӡ��֧�ֶ�ֽ̬��
            .Rows = IIF(objFmt.ֽ�� = 1, 14, 13)
            .TextMatrix(1, 0) = "����Ԫ��": .TextMatrix(1, 1) = ""
            .TextMatrix(2, 0) = "���ͼ��": .TextMatrix(2, 1) = GetCurOutChart
            .TextMatrix(3, 0) = "Ʊ��": .TextMatrix(3, 1) = IIF(objReport.Ʊ��, "��", "��")
            .TextMatrix(4, 0) = "�ձ��ӡ": .TextMatrix(4, 1) = IIF(objReport.��ӡ��ʽ = 0, "��", "��")
            .TextMatrix(5, 0) = "��ӡ��": .TextMatrix(5, 1) = objReport.��ӡ��
            .TextMatrix(6, 0) = "ֽ��": .TextMatrix(6, 1) = GetPaperName(objFmt.ֽ��, objFmt.W, objFmt.H)
            .TextMatrix(7, 0) = "ֽ��": .TextMatrix(7, 1) = IIF(objFmt.ֽ�� = 1, "����", "����")
            .TextMatrix(8, 0) = "�߶�": .TextMatrix(8, 1) = CLng(objFmt.H / Twip_mm) & "����"
            .TextMatrix(9, 0) = "���": .TextMatrix(9, 1) = CLng(objFmt.W / Twip_mm) & "����"
            .TextMatrix(10, 0) = "��ֽ��ʽ": .TextMatrix(10, 1) = CboTest.Text
            .TextMatrix(11, 0) = "��ֹ��ʼʱ��": .TextMatrix(11, 1) = Format(objReport.��ֹ��ʼʱ��, "HH:mm:ss")
            .TextMatrix(12, 0) = "��ֹ����ʱ��": .TextMatrix(12, 1) = Format(objReport.��ֹ����ʱ��, "HH:mm:ss")
            If objFmt.ֽ�� = 1 Then
                .TextMatrix(13, 0) = "��ֽ̬��": .TextMatrix(13, 1) = IIF(objFmt.��ֽ̬��, "��", "��")
            End If
        End If
        If lngRow <= .Rows - 1 Then
            .Row = IIF(lngRow <= .Rows - 1, lngRow, 1)
        Else
            .Row = 1
        End If
         .Col = 1
         mshAtt_AfterRowColChange 0, 0, mshAtt.Row, mshAtt.Col
        .Redraw = True
    End With
End Sub

Private Sub NoneEdit()
    txtAtt.Text = "": txtAtt.Visible = False
    cmdAtt.Visible = False
    cboAtt.Clear: cboAtt.Visible = False
    cboText.Clear: cboText.Visible = False
    dtpAtt.Visible = False
End Sub

Private Function SelIndex() As Integer
'���ܣ���ֻ��һ��Ԫ�ؿؼ���ѡ��ʱ,������ؼ�����
    Dim tmpObj As PictureBox
    
    If lblSize.count > 9 Or lblSize.count = 1 Then SelIndex = 0: Exit Function
    For Each tmpObj In lblSize
        If tmpObj.Index Mod 8 = 1 Then SelIndex = CInt(tmpObj.Tag): Exit Function
    Next
End Function

Private Sub MoveSelect(lngX As Long, lngY As Long, Optional ByVal blnReSize As Boolean)
'����:�ƶ�ѡ��Ԫ�ؿؼ�
'����:lngX=Xƫ����,lngY=Yƫ����,blnReSize=�ı�Ԫ�ش�Сʱ����
    Dim ObjSel As Control, tmpObj As PictureBox
    Dim tmpID As RelatID, objParent As VSFlexGrid
    Dim ItemThis As RPTItem, blnMove As Boolean
    Dim tmpObj1 As PictureBox
    Dim tmpItem As RPTItem, blntmp As Boolean
    Dim vPoint As PointAPI
    
    'Ϊ����ٶ�,����MoveWindow����
    If lngX = 0 And lngY = 0 And blnReSize = False Then Exit Sub
    If blnLock Then Exit Sub '����
    For Each tmpObj In lblSize
        If tmpObj.Index Mod 8 = 1 Then
            blnMove = Not objReport.Items("_" & tmpObj.Tag).ϵͳ
            If blnMove Then
                '������ƶ���Ƭ����Ƭ�ڲ���Ԫ�ز����ƶ�
                If objReport.Items("_" & tmpObj.Tag).��ID <> 0 Then
                    For Each tmpObj1 In lblSize
                        If tmpObj1.Index Mod 8 = 1 Then
                            If objReport.Items("_" & tmpObj.Tag).��ID = objReport.Items("_" & tmpObj1.Tag).id Then
                                GoTo NextObj
                            End If
                        End If
                    Next
                End If
                Select Case objReport.Items("_" & tmpObj.Tag).����
                    Case 1
                        Set ObjSel = lblLine(tmpObj.Tag)
                    Case 2, 3, 10, 12 '@@@
                        If objReport.Items("_" & tmpObj.Tag).���� = 12 Then
                            Set ObjSel = Chart(tmpObj.Tag)
                            If ObjSel.Top + lngY < 0 Then lngY = 0
                            If ObjSel.Left + lngX < 0 Then lngX = 0
                        ElseIf objReport.Items("_" & tmpObj.Tag).���� = 10 Then
                            Set ObjSel = Shp(tmpObj.Tag)
                        Else
                            Set ObjSel = lbl(tmpObj.Tag)
                        End If
                        
                        blnMove = (objReport.Items("_" & tmpObj.Tag).���� = "")
                        
                        If Not blnMove Then
                            lngX = 0
                            For Each ItemThis In objReport.Items
                                If ItemThis.��ʽ�� = mbytCurrFmt _
                                    And ItemThis.���� = objReport.Items("_" & tmpObj.Tag).���� _
                                    And InStr(1, "4,5", ItemThis.����) <> 0 Then
                                    On Error Resume Next
                                    Set objParent = msh(ItemThis.Key)
                                    Exit For
                                End If
                            Next
                            
                            If Mid(objReport.Items("_" & tmpObj.Tag).����, 1, 1) = 1 Then   '������
                                If ObjSel.Top + ObjSel.Height + lngY > objParent.Top - 100 Then lngY = 0
                            Else
                                If ObjSel.Top + lngY < objParent.Top + objParent.Height + 100 Then lngY = 0
                            End If
                            blnMove = True
                        End If
                    Case 11
                        Set ObjSel = img(tmpObj.Tag)
                    Case 14
                        Set ObjSel = pic(tmpObj.Tag)
                    Case 13
                        Set ObjSel = ImgCode(tmpObj.Tag)
                    Case 4, 5
                        '����Ǹ��ӱ��������ӱ��,���˳�
                        blnMove = (objReport.Items("_" & tmpObj.Tag).���� = "")
                        
                        Set ObjSel = msh(tmpObj.Tag)
                        
                        '��������
                        For Each tmpID In objReport.Items("_" & tmpObj.Tag).CopyIDs
                            msh(tmpID.id).Top = msh(tmpID.id).Top + lngY
                            msh(tmpID.id).Left = msh(tmpID.id).Left + lngX
                        Next
                        Call LinkMove(ObjSel.Index, lngX, lngY)
                End Select
            End If
            
            If blnMove Then
                '����Ԫ�ز��ж��Ƿ�����
                ObjSel.Top = ObjSel.Top + lngY
                ObjSel.Left = ObjSel.Left + lngX
                If objReport.Items("_" & tmpObj.Tag).���� = 10 Then
                    '����ͬ���ƶ�lblshp
                    lblshp(tmpObj.Tag).Top = ObjSel.Top
                    lblshp(tmpObj.Tag).Left = ObjSel.Left
                End If
                blntmp = False
                If objReport.Items("_" & tmpObj.Tag).���� <> 14 And lblSize.count = 9 And objReport.Items("_" & tmpObj.Tag).���� <> 12 Then
                    If UCase(ObjSel.Container.name) = "PIC" Then
                        '�������Ԫ�أ����ж��Ƿ��ƶ����˿�Ƭ����,��������Ƭ
                        If ObjSel.Top > ObjSel.Container.Height Or ObjSel.Left > ObjSel.Container.Width Or ObjSel.Top < -1 * ObjSel.Height Or ObjSel.Left < -1 * ObjSel.Width Then
                            '�Ƴ��˿�Ƭ
                            For Each tmpItem In objReport.Items
                                If tmpItem.���� = 14 And tmpItem.id <> ObjSel.Container.Index And tmpItem.��ʽ�� = mbytCurrFmt Then
                                    If ObjSel.Top + ObjSel.Container.Top >= tmpItem.Y And ObjSel.Left + ObjSel.Container.Left >= tmpItem.X And _
                                        ObjSel.Height + ObjSel.Top + ObjSel.Container.Top <= tmpItem.Y + tmpItem.H And ObjSel.Width + ObjSel.Left + ObjSel.Container.Left <= tmpItem.X + tmpItem.W Then
                                        blntmp = True
                                        mlngY = ObjSel.Top + ObjSel.Container.Top - tmpItem.Y
                                        mlngX = ObjSel.Left + ObjSel.Container.Left - tmpItem.X
                                        Set mobjMove = pic(tmpItem.id)
                                        Exit For
                                    Else
                                        mlngY = ObjSel.Top
                                        mlngX = ObjSel.Left
                                        Set mobjMove = picPaper
                                    End If
                                End If
                            Next
                            If blntmp = False Then
                                'û������������Ƭ�ͷ���ֽ����
                                If objReport.Items("_" & tmpObj.Tag).���� = 4 Then
                                    ObjSel.Top = ObjSel.Top + ObjSel.Container.Top
                                    ObjSel.Left = ObjSel.Left + ObjSel.Container.Left
                                    Set ObjSel.Container = picPaper
                                    objReport.Items("_" & ObjSel.Index).��ID = 0
                                    objReport.Items("_" & ObjSel.Index).X = ObjSel.Left + ObjSel.Container.Left
                                    objReport.Items("_" & ObjSel.Index).Y = ObjSel.Top + ObjSel.Container.Top
                                    mlngY = ObjSel.Top + ObjSel.Container.Top
                                    mlngX = ObjSel.Left + ObjSel.Container.Left
                                    Set mobjMove = picPaper
                                    '��������
                                    For Each tmpID In objReport.Items("_" & ObjSel.Index).SubIDs
                                        objReport.Items("_" & tmpID.id).��ID = 0
                                    Next
                                Else
                                    mlngY = ObjSel.Top + ObjSel.Container.Top
                                    mlngX = ObjSel.Left + ObjSel.Container.Left
                                    Set mobjMove = picPaper
                                End If
                            End If
                        Else
                            Set mobjMove = Nothing: mlngX = 0: mlngY = 0
                        End If
                    Else
                        '���������Ԫ�أ����ж��Ƿ��ƶ����˿�Ƭ��
                        For Each tmpItem In objReport.Items
                            If tmpItem.���� = 14 And tmpItem.��ʽ�� = mbytCurrFmt Then
                                If ObjSel.Top >= tmpItem.Y And ObjSel.Left >= tmpItem.X And _
                                    ObjSel.Height + ObjSel.Top <= tmpItem.Y + tmpItem.H And ObjSel.Width + ObjSel.Left <= tmpItem.X + tmpItem.W Then
                                    mlngY = ObjSel.Top - tmpItem.Y
                                    mlngX = ObjSel.Left - tmpItem.X
                                    Set mobjMove = pic(tmpItem.id)
                                    Exit For
                                Else
                                    mlngY = ObjSel.Top
                                    mlngX = ObjSel.Left
                                    Set mobjMove = picPaper
                                End If
                            End If
                        Next
                    End If
                End If
            
                '�������ݶ���
                objReport.Items("_" & tmpObj.Tag).X = Format(ObjSel.Left / sgnMode, "0.00")
                objReport.Items("_" & tmpObj.Tag).Y = Format(ObjSel.Top / sgnMode, "0.00")
            End If
        End If
NextObj:
        If tmpObj.Index <> 0 Then
            If blnMove Then
                With WinProperty
                    .T = (tmpObj.Top + lngY) / Screen.TwipsPerPixelX
                    .l = (tmpObj.Left + lngX) / Screen.TwipsPerPixelX
                    .H = tmpObj.Height / Screen.TwipsPerPixelX
                    .W = tmpObj.Width / Screen.TwipsPerPixelX
                    Call MoveWindow(tmpObj.hwnd, .l, .T, .W, .H, 1)
                End With
            End If
        End If

    Next
    
    Me.Refresh
    BlnSave = False
End Sub

Private Sub ResetColor(idx As Integer)
'���ܣ������༭��ɫ�ָ�Ϊ���ɫ
    Dim i As Integer, j As Integer
    Dim lngRow As Integer, lngCol As Integer
    msh(idx).Redraw = False
    lngRow = msh(idx).Row: lngCol = msh(idx).Col
    For i = 0 To msh(idx).Rows - 1
        msh(idx).Row = i
        For j = 0 To msh(idx).Cols - 1
            msh(idx).Col = j
            If i < msh(idx).FixedRows Or j < msh(idx).FixedCols Then
                msh(idx).CellBackColor = msh(idx).BackColorFixed
                'msh(idx).CellForeColor = msh(idx).ForeColorFixed
            Else
                msh(idx).CellBackColor = msh(idx).BackColor
                msh(idx).CellForeColor = msh(idx).ForeColor
            End If
        Next
    Next
    msh(idx).Row = lngRow: msh(idx).Col = lngCol
    msh(idx).Redraw = True
End Sub

Private Sub SetGridSame(mshS As Control, mshO As Control)
'����:������������ؼ�������ͬ
'˵��������ʱ����������������
    Dim i As Integer, j As Integer
    
    mshO.Redraw = False
    mshS.Redraw = False
    
    mshO.Width = mshS.Width
    mshO.Height = mshS.Height
    mshO.Rows = mshS.Rows
    mshO.Cols = mshS.Cols
    mshO.FixedCols = mshS.FixedCols
    mshO.FixedRows = mshS.FixedRows
    
    mshO.ForeColor = mshS.ForeColor
    mshO.BackColor = mshS.BackColor
    mshO.BackColorFixed = mshS.BackColorFixed
    mshO.ForeColorFixed = mshS.ForeColorFixed
    mshO.BackColorSel = mshS.BackColorSel
    mshO.ForeColorSel = mshS.ForeColorSel
    mshO.GridColor = mshS.GridColor
    mshO.GridColorFixed = mshS.GridColorFixed
    
    mshO.Font.Size = mshS.Font.Size
    mshO.Font.name = mshS.Font.name
    mshO.Font.Bold = mshS.Font.Bold
    mshO.Font.Underline = mshS.Font.Underline
    mshO.Font.Italic = mshS.Font.Italic
    
    For i = 0 To mshS.Rows - 1
        mshS.Row = i: mshO.Row = i
        mshO.RowHeight(i) = mshS.RowHeight(i)
        mshO.MergeRow(i) = mshS.MergeRow(i)
        For j = 0 To mshS.Cols - 1
            mshS.Col = j: mshO.Col = j
            mshO.CellAlignment = mshS.CellAlignment
            mshO.CellFontBold = mshS.CellFontBold
            mshO.CellFontName = mshS.CellFontName
            mshO.CellFontSize = mshS.CellFontSize
            mshO.CellFontItalic = mshS.CellFontItalic
            mshO.CellFontUnderline = mshS.CellFontUnderline
            mshO.TextMatrix(i, j) = mshS.TextMatrix(i, j)
            If i <= mshS.FixedRows - 1 Or j <= mshS.FixedCols - 1 Then
                mshO.CellBackColor = mshS.BackColorFixed
                mshO.CellForeColor = mshS.ForeColorFixed
            Else
                mshO.CellBackColor = mshS.BackColor
                mshO.CellForeColor = mshS.ForeColor
            End If
        Next
    Next
    For i = 0 To mshS.Cols - 1
        mshO.ColWidth(i) = mshS.ColWidth(i)
        mshO.ColAlignment(i) = mshS.ColAlignment(i)
        mshO.MergeCol(i) = mshS.MergeCol(i)
    Next
    
    mshO.Redraw = True
    mshS.Redraw = True
End Sub

Private Sub SetGridLike(mshS As Control, mshO As Control)
'����:������������ؼ�������ͬ(�����壬�и߼���ɫ��ͬ)
'˵��������ʱ����������������
    Dim i As Integer, j As Integer
    
    mshO.Redraw = False
    mshS.Redraw = False
    
    mshO.ForeColor = mshS.ForeColor
    mshO.BackColor = mshS.BackColor
    mshO.BackColorFixed = mshS.BackColorFixed
    mshO.ForeColorFixed = mshS.ForeColorFixed
    mshO.BackColorSel = mshS.BackColorSel
    mshO.ForeColorSel = mshS.ForeColorSel
    mshO.GridColor = mshS.GridColor
    mshO.GridColorFixed = mshS.GridColorFixed
    
    mshO.Font.Size = mshS.Font.Size
    mshO.Font.name = mshS.Font.name
    mshO.Font.Bold = mshS.Font.Bold
    mshO.Font.Underline = mshS.Font.Underline
    mshO.Font.Italic = mshS.Font.Italic
    mshO.RowHeightMin = mshS.RowHeightMin
    objReport.Items("_" & mshO.Index).�ֺ� = objReport.Items("_" & mshS.Index).�ֺ�
    objReport.Items("_" & mshO.Index).���� = objReport.Items("_" & mshS.Index).����
    objReport.Items("_" & mshO.Index).б�� = objReport.Items("_" & mshS.Index).б��
    objReport.Items("_" & mshO.Index).���� = objReport.Items("_" & mshS.Index).����
    objReport.Items("_" & mshO.Index).���� = objReport.Items("_" & mshS.Index).����
    objReport.Items("_" & mshO.Index).���� = objReport.Items("_" & mshS.Index).����
    objReport.Items("_" & mshO.Index).ǰ�� = objReport.Items("_" & mshS.Index).ǰ��
    objReport.Items("_" & mshO.Index).���� = objReport.Items("_" & mshS.Index).����
    objReport.Items("_" & mshO.Index).�и� = objReport.Items("_" & mshS.Index).�и�
    
    For i = 0 To mshS.Rows - 1
        If i <= mshO.Rows - 1 Then
            mshS.Row = i: mshO.Row = i
            mshO.RowHeight(i) = mshS.RowHeight(i)
            For j = 0 To mshS.Cols - 1
                If j <= mshO.Cols - 1 Then
                    mshS.Col = j
                    mshO.Col = j
                    mshO.CellFontBold = mshS.CellFontBold
                    mshO.CellFontName = mshS.CellFontName
                    mshO.CellFontSize = mshS.CellFontSize
                    mshO.CellFontItalic = mshS.CellFontItalic
                    mshO.CellFontUnderline = mshS.CellFontUnderline
                    If i <= mshS.FixedRows - 1 Or j <= mshS.FixedCols - 1 Then
                        mshO.CellBackColor = mshS.BackColorFixed
                        mshO.CellForeColor = mshS.ForeColorFixed
                    Else
                        mshO.CellBackColor = mshS.BackColor
                        mshO.CellForeColor = mshS.ForeColor
                    End If
                End If
            Next
        End If
    Next
    
    mshO.Redraw = True
    mshS.Redraw = True
End Sub

Private Sub SetCopyGrid(intIdx As Integer)
'����:��ָ�����ݱ�ķ������е���
    Dim i As Integer
    Dim tmpID As RelatID
    i = 0
    For Each tmpID In objReport.Items("_" & intIdx).CopyIDs
        i = i + 1
        Call SetGridSame(msh(intIdx), msh(tmpID.id))
        msh(tmpID.id).Top = msh(intIdx).Top
        msh(tmpID.id).Left = msh(intIdx).Left + (msh(intIdx).Width - 15) * i
    Next
End Sub

Private Sub SetSelAlign(bytAlign As Byte)
'����:����ѡ�пؼ�����
'����:
'     bytAlign=1:�����,2:�Ҷ���,3:�϶���,4:�¶���,5:ˮƽ���ж���,6:��ֱ���ж���,7:��ͬ���,8:��ͬ�߶�,9:�����ͬ
'˵��������7,8,9ʱ,ע�������С��ȼ��߶�

    Dim tmpObj As PictureBox, ObjSel As Control, tmpID As RelatID
    Dim ItemSend As RPTItem
    Dim lngPreX As Long, lngPreY As Long
    Dim lngOffX As Long, lngOffY As Long 'ǰ��ƫ����
    Dim lngMinW As Long, lngMinH As Long, i As Integer
    Dim xx As Integer, yy As Integer, zz As Integer
    
    If GetSelNum < 2 Then Exit Sub
    If objLastSel Is Nothing Then Exit Sub
    
    For Each tmpObj In lblSize
        If tmpObj.Index Mod 8 = 1 And tmpObj.Tag <> objLastSel.Index Then
            lngMinW = 0: lngMinH = 0
            Select Case objReport.Items("_" & tmpObj.Tag).����
                Case 1
                    Set ObjSel = lblLine(tmpObj.Tag)
                Case 2, 3
                    Set ObjSel = lbl(tmpObj.Tag)
                Case 10
                    Set ObjSel = Shp(tmpObj.Tag)
                Case 11
                    Set ObjSel = img(tmpObj.Tag)
                Case 14
                    Set ObjSel = pic(tmpObj.Tag)
                Case 4
                    Set ObjSel = msh(tmpObj.Tag)
                    lngMinW = msh(tmpObj.Tag).ColWidth(0) + 15
                    lngMinH = msh(tmpObj.Tag).RowHeight(0) * (msh(tmpObj.Tag).FixedRows + 2) + 15
                Case 5
                    Set ObjSel = msh(tmpObj.Tag)
                    xx = msh(tmpObj.Tag).FixedCols '���������Ŀ��
                    yy = msh(tmpObj.Tag).FixedRows - 1 '���������Ŀ��
                    For Each tmpID In objReport.Items("_" & tmpObj.Tag).SubIDs
                        If objReport.Items("_" & tmpID.id).���� = 9 Then zz = zz + 1 'ͳ����Ŀ��
                    Next
                    lngMinH = msh(tmpObj.Tag).RowHeight(0) * (yy + 1) + 15
                    For i = 0 To xx + zz - 1
                        lngMinW = lngMinW + msh(tmpObj.Tag).ColWidth(i)
                    Next
                    lngMinW = lngMinW + 60
                Case 12 '@@@
                    Set ObjSel = Chart(tmpObj.Tag)
                    lngMinW = Chart(0).Width: lngMinH = Chart(0).Height
                Case 13
                    Set ObjSel = ImgCode(tmpObj.Tag)
            End Select
                        
            '@@@
            lngMinW = lngMinW * sgnMode
            lngMinH = lngMinH * sgnMode
            
            If bytAlign < 7 Then '��������
                lngPreX = ObjSel.Left: lngPreY = ObjSel.Top
                Select Case bytAlign
                    Case 1
                        ObjSel.Left = objLastSel.Left
                    Case 2
                        ObjSel.Left = objLastSel.Left + objLastSel.Width - ObjSel.Width
                    Case 3
                        ObjSel.Top = objLastSel.Top
                    Case 4
                        ObjSel.Top = objLastSel.Top + objLastSel.Height - ObjSel.Height
                    Case 5
                        ObjSel.Top = objLastSel.Top + (objLastSel.Height - ObjSel.Height) / 2
                    Case 6
                        ObjSel.Left = objLastSel.Left + (objLastSel.Width - ObjSel.Width) / 2
                End Select
                
                If objReport.Items("_" & tmpObj.Tag).���� = 12 Then '@@@
                    If ObjSel.Left < 0 Then ObjSel.Left = 0
                    If ObjSel.Top < 0 Then ObjSel.Top = 0
                End If
                
                lngOffX = ObjSel.Left - lngPreX
                lngOffY = ObjSel.Top - lngPreY
                
                For i = CInt(Mid(ObjSel.Tag, 3)) To CInt(Mid(ObjSel.Tag, 3)) + 7
                    lblSize(i).Left = lblSize(i).Left + lngOffX
                    lblSize(i).Top = lblSize(i).Top + lngOffY
                Next
                
                '�������ݶ���
                objReport.Items("_" & tmpObj.Tag).X = Format(ObjSel.Left / sgnMode, "0.00")
                objReport.Items("_" & tmpObj.Tag).Y = Format(ObjSel.Top / sgnMode, "0.00")
            Else '�ߴ�����
                Select Case bytAlign
                    Case 7
                        If Not (objReport.Items("_" & tmpObj.Tag).���� = 1 And objReport.Items("_" & tmpObj.Tag).W < objReport.Items("_" & tmpObj.Tag).H) Then
                            If objReport.Items("_" & tmpObj.Tag).���� = 4 Or objReport.Items("_" & tmpObj.Tag).���� = 5 Then
                                If objLastSel.Width < lngMinW Then
                                    ObjSel.Width = lngMinW
                                Else
                                    ObjSel.Width = objLastSel.Width
                                End If
                            Else
                                ObjSel.Width = objLastSel.Width
                            End If
                        End If
                    Case 8
                        If Not (objReport.Items("_" & tmpObj.Tag).���� = 1 And objReport.Items("_" & tmpObj.Tag).H < objReport.Items("_" & tmpObj.Tag).W) Then
                            If objReport.Items("_" & tmpObj.Tag).���� = 4 Or objReport.Items("_" & tmpObj.Tag).���� = 5 Then
                                If objLastSel.Height < lngMinH Then
                                    ObjSel.Height = lngMinH
                                Else
                                    ObjSel.Height = objLastSel.Height
                                End If
                            Else
                                ObjSel.Height = objLastSel.Height
                            End If
                        End If
                    Case 9
                        If Not (objReport.Items("_" & tmpObj.Tag).���� = 1 And objReport.Items("_" & tmpObj.Tag).W < objReport.Items("_" & tmpObj.Tag).H) Then
                            If objReport.Items("_" & tmpObj.Tag).���� = 4 Or objReport.Items("_" & tmpObj.Tag).���� = 5 Then
                                If objLastSel.Width < lngMinW Then
                                    ObjSel.Width = lngMinW
                                Else
                                    ObjSel.Width = objLastSel.Width
                                End If
                            Else
                                ObjSel.Width = objLastSel.Width
                            End If
                        End If
                        If Not (objReport.Items("_" & tmpObj.Tag).���� = 1 And objReport.Items("_" & tmpObj.Tag).H < objReport.Items("_" & tmpObj.Tag).W) Then
                            If objReport.Items("_" & tmpObj.Tag).���� = 4 Or objReport.Items("_" & tmpObj.Tag).���� = 5 Then
                                If objLastSel.Height < lngMinH Then
                                    ObjSel.Height = lngMinH
                                Else
                                    ObjSel.Height = objLastSel.Height
                                End If
                            Else
                                ObjSel.Height = objLastSel.Height
                            End If
                        End If
                End Select
                
                If objReport.Items("_" & tmpObj.Tag).���� = 12 Then '@@@
                    If ObjSel.Left < 0 Then ObjSel.Left = 0
                    If ObjSel.Top < 0 Then ObjSel.Top = 0
                End If
                
                For i = CInt(Mid(ObjSel.Tag, 3)) To CInt(Mid(ObjSel.Tag, 3)) + 7
                    Select Case IIF(i Mod 8 <> 0, i Mod 8, 8) '��λѡ��߿��λ��
                        Case 1 '����
                            lblSize(i).Top = ObjSel.Top - lblSize(i).Height
                            lblSize(i).Left = ObjSel.Left + (ObjSel.Width - lblSize(i).Width) / 2
                        Case 2 '����
                            lblSize(i).Top = ObjSel.Top - lblSize(i).Height
                            lblSize(i).Left = ObjSel.Left + ObjSel.Width
                        Case 3 '����
                            lblSize(i).Top = ObjSel.Top + (ObjSel.Height - lblSize(i).Height) / 2
                            lblSize(i).Left = ObjSel.Left + ObjSel.Width
                        Case 4 '����
                            lblSize(i).Top = ObjSel.Top + ObjSel.Height
                            lblSize(i).Left = ObjSel.Left + ObjSel.Width
                        Case 5 '����
                            lblSize(i).Top = ObjSel.Top + ObjSel.Height
                            lblSize(i).Left = ObjSel.Left + (ObjSel.Width - lblSize(i).Width) / 2
                        Case 6 '����
                            lblSize(i).Top = ObjSel.Top + ObjSel.Height
                            lblSize(i).Left = ObjSel.Left - lblSize(i).Width
                        Case 7 '����
                            lblSize(i).Top = ObjSel.Top + (ObjSel.Height - lblSize(i).Height) / 2
                            lblSize(i).Left = ObjSel.Left - lblSize(i).Width
                        Case 8 '����
                            lblSize(i).Top = ObjSel.Top - lblSize(i).Height
                            lblSize(i).Left = ObjSel.Left - lblSize(i).Width
                    End Select
                Next
                
                '�������ݶ���
                If Not (objReport.Items("_" & tmpObj.Tag).���� = 1 And objReport.Items("_" & tmpObj.Tag).W < objReport.Items("_" & tmpObj.Tag).H) Then objReport.Items("_" & tmpObj.Tag).W = Format(ObjSel.Width / sgnMode, "0.00")
                If Not (objReport.Items("_" & tmpObj.Tag).���� = 1 And objReport.Items("_" & tmpObj.Tag).H < objReport.Items("_" & tmpObj.Tag).W) Then objReport.Items("_" & tmpObj.Tag).H = Format(ObjSel.Height / sgnMode, "0.00")
            End If
            
            If objReport.Items("_" & tmpObj.Tag).���� = 4 Or objReport.Items("_" & tmpObj.Tag).���� = 5 Then
                Call SetGridLine(tmpObj.Tag)
            End If
            
            '��������
            If objReport.Items("_" & tmpObj.Tag).���� = 4 Then
                Call SetCopyGrid(tmpObj.Tag)
            End If
        End If
        
        If Not ObjSel Is Nothing Then
            If InStr(1, "4,5", objReport.Items("_" & ObjSel.Index).����) <> 0 And objReport.Items("_" & ObjSel.Index).���� = "" Then
                Call CopyItem(ItemSend, objReport.Items("_" & ObjSel.Index))
                Call SetChildWH(ObjSel.Index)
                
                Dim ResizeItem As RPTItem, IntLastCurID As Integer
                IntLastCurID = intCurID
                For Each ResizeItem In objReport.Items
                    If ResizeItem.��ʽ�� = mbytCurrFmt And ResizeItem.���� = ItemSend.���� And ResizeItem.���� = 2 Then
                        intCurID = ResizeItem.Key
                        Call ReferTo
                    End If
                Next
                intCurID = IntLastCurID
            ElseIf objReport.Items("_" & ObjSel.Index).���� = 2 And objReport.Items("_" & ObjSel.Index).���� <> "" Then
                IntLastCurID = intCurID
                intCurID = ObjSel.Index
                Call ReferTo
                intCurID = IntLastCurID
            End If
        End If
    Next
    BlnSave = False
End Sub

Private Sub tbr2_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Call NoneEdit
    Select Case ButtonMenu.Key
        Case "HscSame"
            mnuFormat_HscSpace_Same_Click
        Case "HscAdd"
            mnuFormat_HscSpace_Add_Click
        Case "HscDec"
            mnuFormat_HscSpace_Dec_Click
        Case "VscSame"
            mnuFormat_VscSpace_Same_Click
        Case "VscAdd"
            mnuFormat_VscSpace_Add_Click
        Case "VscDec"
            mnuFormat_VscSpace_Dec_Click
        Case "Page", "Width", "Height", "Scale200", "Scale100", "Scale75", "Scale50", "Scale25"
            mnuViewScaleMode_Click ButtonMenu.Index - 1
    End Select
End Sub

Private Sub tbr2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuView, 2
End Sub

Private Sub tbr1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuView, 2
End Sub

Private Sub SetVscSpace(bytType As Integer)
'����:����ѡ�пؼ��Ĵ�ֱ���
'����:bytType=0:��ͬ,-1:����,1����
    Dim i As Integer, j As Integer, lngH As Long
    Dim tmpObj As PictureBox, ObjSel As Control, arrObj() As Control '���հ����ϵ��µ�˳����ѡ�пؼ�
    Dim ItemSend As RPTItem
    
    Const SPACE_STEP As Long = 75 'һ�����ӻ���ٵļ��
    
    On Error Resume Next
    
    i = GetSelNum()
    If i < 2 Or (i = 2 And bytType = 0) Then Exit Sub '��Ҫ���������ͬʱ,ѡ�пؼ����������3
    
    '�γ�ѡ�пؼ�����
    ReDim arrObj(i - 1)
    i = 0
    For Each tmpObj In lblSize
        If tmpObj.Index Mod 8 = 1 Then
            Select Case objReport.Items("_" & tmpObj.Tag).����
                Case 1
                    Set arrObj(i) = lblLine(tmpObj.Tag)
                Case 2, 3
                    Set arrObj(i) = lbl(tmpObj.Tag)
                Case 10
                    Set arrObj(i) = Shp(tmpObj.Tag)
                Case 11
                    Set arrObj(i) = img(tmpObj.Tag)
                Case 4, 5
                    Set arrObj(i) = msh(tmpObj.Tag)
                Case 12 '@@@
                    Set arrObj(i) = Chart(tmpObj.Tag)
                Case 13
                    Set arrObj(i) = ImgCode(tmpObj.Tag)
                Case 14
                    Set arrObj(i) = pic(tmpObj.Tag)
            End Select
            i = i + 1
        End If
    Next
    '�Կؼ����鰴Top��С�����˳��˳��
    For i = 0 To UBound(arrObj) - 1
        For j = i + 1 To UBound(arrObj)
            If arrObj(j).Top < arrObj(i).Top Then
                Set ObjSel = arrObj(j)
                Set arrObj(j) = arrObj(i)
                Set arrObj(i) = ObjSel
            End If
        Next
    Next
    Select Case bytType
        Case 0
            '��ƽ�����
            lngH = 0
            For i = 0 To UBound(arrObj) - 1
                lngH = lngH + (arrObj(i + 1).Top - (arrObj(i).Top + arrObj(i).Height))
            Next
            lngH = lngH \ UBound(arrObj)
            '��λ��Ŀ
            For i = 1 To UBound(arrObj)
                '����λʱ�ƶ���Ŀ��,�����ж�Ӧ����Ŀֵ����Ӧ�����仯(SET)
                Call SeekItem(arrObj(i), arrObj(i).Left, arrObj(i - 1).Top + arrObj(i - 1).Height + lngH)
            Next
        Case 1
            lngH = SPACE_STEP
            For i = 1 To UBound(arrObj)
                Call SeekItem(arrObj(i), arrObj(i).Left, arrObj(i).Top + lngH)
                lngH = lngH + SPACE_STEP
            Next
        Case -1
            lngH = SPACE_STEP
            For i = 1 To UBound(arrObj)
                Call SeekItem(arrObj(i), arrObj(i).Left, arrObj(i).Top - lngH)
                lngH = lngH + SPACE_STEP
            Next
    End Select
    
    For i = 0 To UBound(arrObj) - 1
        If Not arrObj(i) Is Nothing Then
            If InStr(1, "4,5", objReport.Items("_" & arrObj(i).Index).����) <> 0 And objReport.Items("_" & arrObj(i).Index).���� = "" Then
                Call CopyItem(ItemSend, objReport.Items("_" & arrObj(i).Index))
                Call SetChildWH(arrObj(i).Index)
                Dim ResizeItem As RPTItem, IntLastCurID As Integer
                IntLastCurID = intCurID
                For Each ResizeItem In objReport.Items
                    If ResizeItem.��ʽ�� = mbytCurrFmt And ResizeItem.���� = ItemSend.���� And ResizeItem.���� = 2 Then
                        intCurID = ResizeItem.Key
                        Call ReferTo
                    End If
                Next
                intCurID = IntLastCurID
            ElseIf objReport.Items("_" & arrObj(i).Index).���� = 2 And objReport.Items("_" & arrObj(i).Index).���� <> "" Then
                IntLastCurID = intCurID
                intCurID = arrObj(i).Index
                Call ReferTo
                intCurID = IntLastCurID
            End If
        End If
    Next
    BlnSave = False
End Sub

Private Sub SetHscSpace(bytType As Integer)
'����:����ѡ�пؼ���ˮƽ���
'����:bytType=0:��ͬ,-1:����,1����
    Dim i As Integer, j As Integer, lngW As Long
    Dim tmpObj As PictureBox, ObjSel As Control, arrObj() As Control  '���հ������ҵ�˳����ѡ�пؼ�
    Dim ItemSend As RPTItem
    
    Const SPACE_STEP As Long = 75 'һ�����ӻ���ٵļ��
    
    i = GetSelNum()
    If i < 2 Or (i = 2 And bytType = 0) Then Exit Sub '��Ҫ���������ͬʱ,ѡ�пؼ����������3
    
    '�γ�ѡ�пؼ�����
    ReDim arrObj(i - 1)
    i = 0
    For Each tmpObj In lblSize
        If tmpObj.Index Mod 8 = 1 Then
            Select Case objReport.Items("_" & tmpObj.Tag).����
                Case 1
                    Set arrObj(i) = lblLine(tmpObj.Tag)
                Case 2, 3
                    Set arrObj(i) = lbl(tmpObj.Tag)
                Case 10
                    Set arrObj(i) = Shp(tmpObj.Tag)
                Case 11
                    Set arrObj(i) = img(tmpObj.Tag)
                Case 4, 5
                    Set arrObj(i) = msh(tmpObj.Tag)
                Case 12 '@@@
                    Set arrObj(i) = Chart(tmpObj.Tag)
                Case 13
                    Set arrObj(i) = ImgCode(tmpObj.Tag)
                Case 14
                    Set arrObj(i) = pic(tmpObj.Tag)
            End Select
            i = i + 1
        End If
    Next
    '�Կؼ����鰴Left��С�����˳��˳��
    For i = 0 To UBound(arrObj) - 1
        For j = i + 1 To UBound(arrObj)
            If arrObj(j).Left < arrObj(i).Left Then
                Set ObjSel = arrObj(j)
                Set arrObj(j) = arrObj(i)
                Set arrObj(i) = ObjSel
            End If
        Next
    Next
    Select Case bytType
        Case 0
            '��ƽ�����
            lngW = 0
            For i = 0 To UBound(arrObj) - 1
                lngW = lngW + (arrObj(i + 1).Left - (arrObj(i).Left + arrObj(i).Width))
            Next
            lngW = lngW \ UBound(arrObj)
            '��λ��Ŀ
            For i = 1 To UBound(arrObj)
                '����λʱ�ƶ���Ŀ��,�����ж�Ӧ����Ŀֵ����Ӧ�����仯(SET)
                Call SeekItem(arrObj(i), arrObj(i - 1).Left + arrObj(i - 1).Width + lngW, arrObj(i).Top)
            Next
        Case 1
            lngW = SPACE_STEP
            For i = 1 To UBound(arrObj)
                Call SeekItem(arrObj(i), arrObj(i).Left + lngW, arrObj(i).Top)
                lngW = lngW + SPACE_STEP
            Next
        Case -1
            lngW = SPACE_STEP
            For i = 1 To UBound(arrObj)
                Call SeekItem(arrObj(i), arrObj(i).Left - lngW, arrObj(i).Top)
                lngW = lngW + SPACE_STEP
            Next
    End Select
    For i = 0 To UBound(arrObj) - 1
        If Not arrObj(i) Is Nothing Then
            If InStr(1, "4,5", objReport.Items("_" & arrObj(i).Index).����) <> 0 And objReport.Items("_" & arrObj(i).Index).���� = "" Then
                Call CopyItem(ItemSend, objReport.Items("_" & arrObj(i).Index))
                Call SetChildWH(arrObj(i).Index)
                Dim ResizeItem As RPTItem, IntLastCurID As Integer
                IntLastCurID = intCurID
                For Each ResizeItem In objReport.Items
                    If ResizeItem.��ʽ�� = mbytCurrFmt And ResizeItem.���� = ItemSend.���� And ResizeItem.���� = 2 Then
                        intCurID = ResizeItem.Key
                        Call ReferTo
                    End If
                Next
                intCurID = IntLastCurID
            ElseIf objReport.Items("_" & arrObj(i).Index).���� = 2 And objReport.Items("_" & arrObj(i).Index).���� <> "" Then
                IntLastCurID = intCurID
                intCurID = arrObj(i).Index
                Call ReferTo
                intCurID = IntLastCurID
            End If
        End If
    Next
    BlnSave = False
End Sub

Private Sub SeekItem(objSeek As Control, X As Long, Y As Long)
'����:��λ��Ŀ
'����:objSeek=������Ŀ
'˵��:�ú�����Ҫ��ˮƽ�ʹ�ֱ�����������������
    Dim i As Byte
    Dim lngTop As Long, lngLeft As Long
    
    objSeek.Top = Y: objSeek.Left = X
    If objReport.Items("_" & objSeek.Index).���� = 12 Then '@@@
        If objSeek.Top < 0 Then objSeek.Top = 0
        If objSeek.Left < 0 Then objSeek.Left = 0
    End If
    If UCase(objSeek.Container.name) = "PIC" Then
        lngTop = objSeek.Container.Top
        lngLeft = objSeek.Container.Left
    End If
    
    If Mid(objSeek.Tag, 1, 2) = "S_" Then
        For i = CInt(Mid(objSeek.Tag, 3)) To CInt(Mid(objSeek.Tag, 3)) + 7 '�ƶ�Size���
            Select Case IIF(i Mod 8 <> 0, i Mod 8, 8) '��λѡ��߿��λ��
                Case 1 '����
                    lblSize(i).Top = objSeek.Top + lngTop - lblSize(i).Height
                    lblSize(i).Left = objSeek.Left + lngLeft + (objSeek.Width - lblSize(i).Width) / 2
                Case 2 '����
                    lblSize(i).Top = objSeek.Top + lngTop - lblSize(i).Height
                    lblSize(i).Left = objSeek.Left + lngLeft + objSeek.Width
                Case 3 '����
                    lblSize(i).Top = objSeek.Top + lngTop + (objSeek.Height - lblSize(i).Height) / 2
                    lblSize(i).Left = objSeek.Left + lngLeft + objSeek.Width
                Case 4 '����
                    lblSize(i).Top = objSeek.Top + lngTop + objSeek.Height
                    lblSize(i).Left = objSeek.Left + lngLeft + objSeek.Width
                Case 5 '����
                    lblSize(i).Top = objSeek.Top + lngTop + objSeek.Height
                    lblSize(i).Left = objSeek.Left + lngLeft + (objSeek.Width - lblSize(i).Width) / 2
                Case 6 '����
                    lblSize(i).Top = objSeek.Top + lngTop + objSeek.Height
                    lblSize(i).Left = objSeek.Left + lngLeft - lblSize(i).Width
                Case 7 '����
                    lblSize(i).Top = objSeek.Top + lngTop + (objSeek.Height - lblSize(i).Height) / 2
                    lblSize(i).Left = objSeek.Left + lngLeft - lblSize(i).Width
                Case 8 '����
                    lblSize(i).Top = objSeek.Top + lngTop - lblSize(i).Height
                    lblSize(i).Left = objSeek.Left + lngLeft - lblSize(i).Width
            End Select
        Next
    End If
    '��������
    If objReport.Items("_" & objSeek.Index).���� = 4 Then
        Call SetCopyGrid(objSeek.Index)
    End If
    
    '�������ݶ���
    objReport.Items("_" & objSeek.Index).X = Format(objSeek.Left / sgnMode, "0.00")
    objReport.Items("_" & objSeek.Index).Y = Format(objSeek.Top / sgnMode, "0.00")
End Sub

Private Sub SetSelCenter(bytStyle As Byte)
'���ܣ�����ѡ��ؼ�ˮƽ���л�ֱ����
'������bytStyle=0:ˮƽ����,1:��ֱ����
    Dim tmpObj As PictureBox, ObjSel As Object
    Dim ItemSend As RPTItem, objFmt As RPTFmt
    Dim lngW As Long, lngH As Long
    
    If GetSelNum = 0 Then Exit Sub
    
    Set objFmt = objReport.Fmts("_" & mbytCurrFmt)
    If objFmt.ֽ�� = 1 Then
        lngW = objFmt.W
        lngH = objFmt.H
    Else
        lngW = objFmt.H
        lngH = objFmt.W
    End If
    For Each tmpObj In lblSize
        If tmpObj.Index Mod 8 = 1 Then
            Set ObjSel = GetInxObj(tmpObj.Tag)
            If bytStyle = 0 Then
                SeekItem ObjSel, (lngW - ObjSel.Width) / 2, ObjSel.Top
            Else
                SeekItem ObjSel, ObjSel.Left, (lngH - ObjSel.Height) / 2
            End If
            If Not ObjSel Is Nothing Then
                If InStr(1, "4,5", objReport.Items("_" & ObjSel.Index).����) <> 0 And objReport.Items("_" & ObjSel.Index).���� = "" Then
                    Call CopyItem(ItemSend, objReport.Items("_" & ObjSel.Index))
                    Call SetChildWH(ObjSel.Index)
                    Dim ResizeItem As RPTItem, IntLastCurID As Integer
                    IntLastCurID = intCurID
                    For Each ResizeItem In objReport.Items
                        If ResizeItem.��ʽ�� = mbytCurrFmt And ResizeItem.���� = ItemSend.���� And ResizeItem.���� = 2 Then
                            intCurID = ResizeItem.Key
                            Call ReferTo
                        End If
                    Next
                    intCurID = IntLastCurID
                ElseIf objReport.Items("_" & ObjSel.Index).���� = 2 And objReport.Items("_" & ObjSel.Index).���� <> "" Then
                    IntLastCurID = intCurID
                    intCurID = ObjSel.Index
                    Call ReferTo
                    intCurID = IntLastCurID
                End If
            End If
        End If
    Next
    If GetSelNum = 1 Then Call ShowAttrib(intCurID)
    BlnSave = False
End Sub

Private Sub txtAtt_GotFocus()
    SelAll txtAtt
End Sub

Private Sub txtAtt_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            mshAtt.SetFocus: mshAtt.Refresh: SendKeys "{DOWN}"
        Case vbKeyUp
            mshAtt.SetFocus: mshAtt.Refresh: SendKeys "{UP}"
    End Select
End Sub

Private Sub txtAtt_KeyPress(KeyAscii As Integer)
    Dim ObjSel As Object, tmpID As RelatID, tmpItem As RPTItem
    Dim i As Integer, xx As Long, yy As Long, zz As Long
    Dim lngMinW As Long, lngMinH As Long, sgnH As Single
    Dim arrHead, arrModify
    Dim objBarCode As StdPicture, lngSize As Long
    Dim strBarCode As String, sngWidth As Single
    Dim lngL As Long, lngW As Long
    Dim strInfo As String
    Dim lngReportID As Long
    Dim strReportID As String
    Dim X As Long, Y As Long, k As Long
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        
        '�Ƿ���ָ��ַ�:
        If InString(txtAtt.Text, "'|~^") Then
            MsgBox "�����˷Ƿ��ַ���", vbInformation, App.Title
            txtAtt.SetFocus: Exit Sub
        End If
        Set ObjSel = GetInxObj(intCurID)

        Select Case mshAtt.TextMatrix(mshAtt.Row, 0)
            Case "����"
                txtAtt = Trim(txtAtt.Text)
                If TLen(txtAtt.Text) > 50 Then
                    MsgBox "���Ʋ��ܳ���50���ַ���", vbInformation, App.Title
                    txtAtt.SetFocus: Exit Sub
                End If
                If txtAtt.Text = "" Then
                    MsgBox "���Ʋ���Ϊ�գ�", vbInformation, App.Title
                    txtAtt.SetFocus: Exit Sub
                End If
                
                '������ƵĺϷ���
                If CheckNameValid(txtAtt.Text) = False Then
                    MsgBox "�����еı����ʽ�з��������ظ���", vbInformation, App.Title
                    txtAtt.SetFocus
                    Exit Sub
                End If
                Call ChangeReferTo(objReport.Items("_" & intCurID).����, txtAtt.Text)      '�޸������ӱ�Ĳ��ն���
                mshAtt.TextMatrix(mshAtt.Row, 1) = txtAtt.Text
                objReport.Items("_" & intCurID).���� = txtAtt.Text
                
                txtAtt.Visible = False: mshAtt.SetFocus
                BlnSave = False
            Case "����"
                If TLen(txtAtt.Text) > 255 Then
                    MsgBox "���ݲ��ܳ���255���ַ���", vbInformation, App.Title
                    txtAtt.SetFocus: Exit Sub
                End If
                
                '����Ƿ��ַ����
                If objReport.Items("_" & intCurID).���� = 13 Then
                    If InString(txtAtt.Text, "[]") And Not BracketMatch(txtAtt.Text, "[]") Then
                        MsgBox "�������ݵ����Ų���ԣ�", vbInformation, App.Title
                        txtAtt.SetFocus: Exit Sub
                    End If
                    If ReplaceBracket(txtAtt.Text) <> "" Then
                        If objReport.Items("_" & intCurID).��� = 1 Or objReport.Items("_" & intCurID).��� = 3 Then
                            If Not MatchString(ReplaceBracket(txtAtt.Text), STR_CODE_128) Then
                                MsgBox "���������а����Ƿ��ַ���", vbInformation, App.Title
                                txtAtt.SetFocus: Exit Sub
                            End If
                        ElseIf objReport.Items("_" & intCurID).��� = 2 Then
                            If Not MatchString(ReplaceBracket(txtAtt.Text), STR_CODE_39) Then
                                MsgBox "���������а����Ƿ��ַ���", vbInformation, App.Title
                                txtAtt.SetFocus: Exit Sub
                            End If
                        End If
                    End If
                End If
                
                Dim strNodeName As String, NodeThis As Node
                '�����adLongVarBinary���ֶ�,�������޸�
                xx = InStr(1, txtAtt, "]")
                yy = InStr(1, txtAtt, ".")
                zz = InStr(1, txtAtt, "[")
                If xx > zz And xx > yy And xx <> 0 And zz <> 0 Then
                    strNodeName = Mid(txtAtt, yy + 1, xx - yy - 1)
                    For Each NodeThis In tvwSQL.Nodes
                        If mdlPublic.GetStdNodeText(NodeThis.Text) = strNodeName And IsType(Val(NodeThis.Tag), adLongVarBinary) Then
                            MsgBox "����ѡ��ͼ���ֶ�Ϊ��ǩ�����ݣ�", vbInformation, App.Title
                            mshAtt.TextMatrix(mshAtt.Row, 1) = objReport.Items("_" & intCurID).����
                            Exit Sub
                        End If
                    Next
                End If
                
                mshAtt.TextMatrix(mshAtt.Row, 1) = txtAtt.Text
                If UCase(TypeName(ObjSel)) = "LABEL" Then ObjSel.Caption = txtAtt.Text
                objReport.Items("_" & intCurID).���� = txtAtt.Text
            
                '�Ե��������LblSize�ؼ���λ��
                If UCase(TypeName(ObjSel)) = "LABEL" Then
                    If ObjSel.AutoSize Then
                        Call SelItem(ObjSel.Index, False)
                        Call SelItem(ObjSel.Index, True)
                    End If
                    objReport.Items("_" & intCurID).W = lbl(intCurID).Width / sgnMode
                ElseIf objReport.Items("_" & intCurID).���� = 13 Then
                    With objReport.Items("_" & intCurID)
                        strBarCode = ReplaceBracket(.����)
                        If strBarCode = "" Then strBarCode = "1234567890"
                        
                        Unload frmFlash 'ǿ�Ƴ�ʼPicture����Ȼ�л�����������
                        If .��� = 1 Then
                            Set objBarCode = DrawBarCode128(frmFlash.picTemp, 3, strBarCode, Mid(.��ͷ, 1, 1) = "1")
                        ElseIf .��� = 2 Then
                            Set objBarCode = DrawBarCode39(frmFlash.picTemp, 3, strBarCode, Mid(.��ͷ, 2, 1) = "1", Mid(.��ͷ, 1, 1) = "1")
                        ElseIf .��� = 3 Then
                            Set objBarCode = DrawBarCode128Auto(frmFlash.picTemp, strBarCode, sngWidth, .�и�, Mid(.��ͷ, 1, 1) = "1")
                        ElseIf .��� = 10 Then
                            Set objBarCode = DrawBarCode2D(strBarCode, frmFlash.picTemp, lngSize)
                        End If
                        If Val(Mid(.��ͷ, 3, 1)) <> 0 Then
                            Set objBarCode = PictureSpin(objBarCode, Val(Mid(.��ͷ, 3, 1)), frmFlash.picTemp)
                        End If
                        Set ObjSel.Picture = objBarCode
                        
                        If .��� = 3 Then
                            '128���Զ��������
                            If Val(Mid(.��ͷ, 3, 1)) = 0 Then
                                ObjSel.Width = Format(Me.ScaleX(sngWidth, vbMillimeters, vbTwips) * sgnMode, "0.00")
                                .W = Me.ScaleX(sngWidth, vbMillimeters, vbTwips)
                            Else
                                ObjSel.Height = Format(Me.ScaleY(sngWidth, vbMillimeters, vbTwips) * sgnMode, "0.00")
                                .H = Me.ScaleY(sngWidth, vbMillimeters, vbTwips)
                            End If
                            Call SeekItem(ObjSel, ObjSel.Left, ObjSel.Top)
                        ElseIf .��� = 10 And .�Ե� Then
                            '��ά����ȱʡ�Զ�������С
                            .W = lngSize: .H = lngSize
                            
                            ObjSel.Width = Format(lngSize * sgnMode, "0.00")
                            ObjSel.Height = Format(lngSize * sgnMode, "0.00")
                            
                            Call SeekItem(ObjSel, ObjSel.Left, ObjSel.Top)
                        End If
                    End With
                End If
                
                txtAtt.Visible = False: mshAtt.SetFocus
                BlnSave = False
            Case "X����"
                If Not IsNumeric(txtAtt.Text) Then
                    MsgBox "���������������ݣ�", vbInformation, App.Title
                    txtAtt.SetFocus: Exit Sub
                End If
                If Abs(CDbl(txtAtt.Text)) > 5000 Then
                    If CDbl(txtAtt.Text) > 0 Then
                        txtAtt.Text = 5000
                    Else
                        txtAtt.Text = -5000
                    End If
                End If
                If objReport.Items("_" & ObjSel.Index).���� = 12 Then '@@@
                    If Val(txtAtt.Text) < 0 Then txtAtt.Text = "0.00"
                End If
                
                ObjSel.Left = CDbl(txtAtt.Text) * Twip_mm * sgnMode
                Call SeekItem(ObjSel, ObjSel.Left, ObjSel.Top)
                
                If objReport.Items("_" & intCurID).���� = 4 Then
                    Call SetCopyGrid(intCurID)
                End If
                
                mshAtt.TextMatrix(mshAtt.Row, 1) = Format(txtAtt.Text, "0.00")
                objReport.Items("_" & intCurID).X = Format(ObjSel.Left / sgnMode, "0.00")
                
                txtAtt.Visible = False: mshAtt.SetFocus
                
                Select Case objReport.Items("_" & ObjSel.Index).����
                Case 2
                    AdjustAll (True)
                Case 4, 5
                    SetMainWH (ObjSel.Index)
                End Select
                BlnSave = False
            Case "Y����"
                If Not IsNumeric(txtAtt.Text) Then
                    MsgBox "���������������ݣ�", vbInformation, App.Title
                    txtAtt.SetFocus: Exit Sub
                End If
                If Abs(CDbl(txtAtt.Text)) > 5000 Then
                    If CDbl(txtAtt.Text) > 0 Then
                        txtAtt.Text = 5000
                    Else
                        txtAtt.Text = -5000
                    End If
                End If
                If objReport.Items("_" & ObjSel.Index).���� = 12 Then '@@@
                    If Val(txtAtt.Text) < 0 Then txtAtt.Text = "0.00"
                End If
                
                Call MoveSelect(0, (CDbl(txtAtt.Text) * Twip_mm - ObjSel.Top / sgnMode) * sgnMode)
                
                mshAtt.TextMatrix(mshAtt.Row, 1) = Format(txtAtt.Text, "0.00")
                objReport.Items("_" & intCurID).Y = Format(ObjSel.Top / sgnMode, "0.00")
                
                txtAtt.Visible = False: mshAtt.SetFocus
                
                Select Case objReport.Items("_" & ObjSel.Index).����
                Case 2
                    AdjustAll (True)
                Case 4, 5
                    SetMainWH (ObjSel.Index)
                End Select
                BlnSave = False
            Case "���"
                If Not IsNumeric(txtAtt.Text) Then
                    MsgBox "���������������ݣ�", vbInformation, App.Title
                    txtAtt.SetFocus: Exit Sub
                End If
                txtAtt.Text = Abs(CDbl(txtAtt.Text))
                If CDbl(txtAtt.Text) > 5000 Then txtAtt.Text = 5000
                
                '�����С���
                If objReport.Items("_" & intCurID).���� = 4 Then
                    lngMinW = ObjSel.ColWidth(0) + 15
                ElseIf objReport.Items("_" & intCurID).���� = 5 Then
                    xx = ObjSel.FixedCols '���������Ŀ��
                    For Each tmpID In objReport.Items("_" & intCurID).SubIDs
                        If objReport.Items("_" & tmpID.id).���� = 9 Then zz = zz + 1 'ͳ����Ŀ��
                    Next
                    For i = 0 To xx + zz - 1
                        lngMinW = lngMinW + ObjSel.ColWidth(i)
                    Next
                    lngMinW = lngMinW + 60
                ElseIf objReport.Items("_" & intCurID).���� = 12 Then '@@@
                    lngMinW = Chart(0).Width
                End If
                If CDbl(txtAtt.Text) * Twip_mm < lngMinW Then txtAtt.Text = lngMinW / Twip_mm
                
                ObjSel.Width = CDbl(txtAtt.Text) * Twip_mm * sgnMode
                Call SeekItem(ObjSel, ObjSel.Left, ObjSel.Top)
                
                mshAtt.TextMatrix(mshAtt.Row, 1) = Format(txtAtt.Text, "0.00")
                objReport.Items("_" & intCurID).W = Format(ObjSel.Width / sgnMode, "0.00")
                
                '����Ǳ��,Ҫ����������
                If InStr(1, "4,5", objReport.Items("_" & intCurID).����) <> 0 Then
                    Call SetGridLine(intCurID)
                End If
                If objReport.Items("_" & intCurID).���� = 4 Then
                    Call SetCopyGrid(intCurID)
                End If
                
                txtAtt.Visible = False: mshAtt.SetFocus
                
                Select Case objReport.Items("_" & ObjSel.Index).����
                Case 2
                    Call AdjustAll(True)
                Case 4, 5
                    Call SetMainWH(ObjSel.Index)
                End Select
                BlnSave = False
            Case "�߶�"
                If Not IsNumeric(txtAtt.Text) Then
                    MsgBox "���������������ݣ�", vbInformation, App.Title
                    txtAtt.SetFocus: Exit Sub
                End If
                txtAtt.Text = Abs(CDbl(txtAtt.Text))
                If CDbl(txtAtt.Text) > 5000 Then txtAtt.Text = 5000
                
                '�����С�߶�
                If objReport.Items("_" & intCurID).���� = 4 Then
                    lngMinH = ObjSel.RowHeight(0) * (ObjSel.FixedRows + 2) + 15
                ElseIf objReport.Items("_" & intCurID).���� = 5 Then
                    yy = ObjSel.FixedRows - 1 '���������Ŀ��
                    lngMinH = ObjSel.RowHeight(0) * (yy + 3) + 60
                ElseIf objReport.Items("_" & intCurID).���� = 12 Then '@@@
                    lngMinH = Chart(0).Height
                End If
                If CDbl(txtAtt.Text) * Twip_mm < lngMinH Then txtAtt.Text = lngMinH / Twip_mm
                
                ObjSel.Height = CDbl(txtAtt.Text) * Twip_mm * sgnMode
                Call SeekItem(ObjSel, ObjSel.Left, ObjSel.Top)
                
                mshAtt.TextMatrix(mshAtt.Row, 1) = Format(txtAtt.Text, "0.00")
                objReport.Items("_" & intCurID).H = Format(ObjSel.Height / sgnMode, "0.00")
                
                '����Ǳ��,Ҫ����������
                If InStr(1, "4,5", objReport.Items("_" & intCurID).����) <> 0 Then
                    Call SetGridLine(intCurID)
                End If
                If objReport.Items("_" & intCurID).���� = 4 Then
                    Call SetCopyGrid(intCurID)
                End If
                
                txtAtt.Visible = False: mshAtt.SetFocus
                
                Select Case objReport.Items("_" & ObjSel.Index).����
                Case 2
                    Call AdjustAll(True)
                Case 4, 5
                    Call SetMainWH(ObjSel.Index)
                End Select
                
                '�������ӱ��������������һ��
                If objReport.Items("_" & intCurID).���� = "" And objReport.Items("_" & intCurID).���� = 5 Then
                    For Each tmpItem In objReport.Items
                        If tmpItem.��ʽ�� = mbytCurrFmt And tmpItem.���� = objReport.Items("_" & intCurID).���� And tmpItem.���� = 5 Then
                            Call SetGridLike(msh(intCurID), msh(tmpItem.Key))
                        End If
                    Next
                End If
                BlnSave = False
            Case "�и�"
                On Error Resume Next
                If Not IsNumeric(txtAtt.Text) Then
                    MsgBox "���������������ݣ�", vbInformation, App.Title
                    txtAtt.SetFocus: Exit Sub
                End If
                txtAtt.Text = Abs(CDbl(txtAtt.Text))
                If CDbl(txtAtt.Text) > 5000 Then txtAtt.Text = 5000
                
                '��Ҫ���õ��и߱��뱣֤��������ʾ���й̶�����+2
                '�����ǰѡ�е��ǹ̶���,�����õ�ǰ�̶��еĸ߶�;�����������������и߶�
                PicFontTest.FontName = objReport.Items("_" & intCurID).����
                PicFontTest.FontSize = objReport.Items("_" & intCurID).�ֺ�
                sgnH = (PicFontTest.TextHeight("��") + 15) * sgnMode
                Dim SgnFixedRows As Single
                
                If ObjSel.Row >= ObjSel.FixedRows Then
                    If objReport.Items("_" & intCurID).���� = 4 Then
                        SgnFixedRows = 0
                        For i = 0 To ObjSel.FixedRows - 1
                            SgnFixedRows = SgnFixedRows + ObjSel.RowHeight(i)
                        Next
                        If Abs(Int(-CDbl(txtAtt.Text) * Twip_mm * sgnMode)) < sgnH Then
                            If objReport.Items("_" & intCurID).���� = 5 Then ObjSel.RowHeightMin = sgnH
                            For i = ObjSel.FixedRows To ObjSel.Rows - 1
                                ObjSel.RowHeight(i) = sgnH
                            Next
                        ElseIf Abs(Int((-ObjSel.Height + SgnFixedRows) / Abs(Int(-CDbl(txtAtt.Text) * Twip_mm * sgnMode)))) > 2 Then
                            If objReport.Items("_" & intCurID).���� = 5 Then ObjSel.RowHeightMin = Abs(Int(-ObjSel.Height / (ObjSel.FixedRows + 2)))
                            For i = ObjSel.FixedRows To ObjSel.Rows - 1
                                ObjSel.RowHeight(i) = Abs(Int(-CDbl(txtAtt.Text) * Twip_mm * sgnMode))
                            Next
                        End If
                        mshAtt.TextMatrix(mshAtt.Row, 1) = Format(txtAtt.Text, "0.00")
                        objReport.Items("_" & intCurID).�и� = Format(ObjSel.RowHeight(ObjSel.Row) / sgnMode, "0.00")
                    Else
                        If Abs(Int(-CDbl(txtAtt.Text) * Twip_mm * sgnMode)) < sgnH Then
                            For i = 0 To ObjSel.Rows - 1
                                ObjSel.RowHeight(i) = sgnH
                            Next
                        ElseIf Abs(Int(-ObjSel.Height / Abs(Int(-CDbl(txtAtt.Text) * Twip_mm * sgnMode)))) < ObjSel.FixedRows + 2 Then
                            For i = 0 To ObjSel.Rows - 1
                                ObjSel.RowHeight(i) = Abs(Int(-ObjSel.Height / (ObjSel.FixedRows + 2)))
                            Next
                        Else
                            For i = 0 To ObjSel.Rows - 1
                                ObjSel.RowHeight(i) = Abs(Int(-CDbl(txtAtt.Text) * Twip_mm * sgnMode))
                            Next
                        End If
                        mshAtt.TextMatrix(mshAtt.Row, 1) = Format(txtAtt.Text, "0.00")
                        objReport.Items("_" & intCurID).�и� = Format(ObjSel.RowHeight(0) / sgnMode, "0.00")
                    End If
                Else
                    '�̶���
                    If Abs(Int(-CDbl(txtAtt.Text) * Twip_mm * sgnMode)) / (selCell.Row2 - selCell.Row1 + 1) < sgnH Then
                        If objReport.Items("_" & intCurID).���� = 5 Then ObjSel.RowHeightMin = sgnH
                        For i = selCell.Row1 To selCell.Row2
                            ObjSel.RowHeight(i) = sgnH
                        Next
                    Else
                        If objReport.Items("_" & intCurID).���� = 5 Then ObjSel.RowHeightMin = sgnH
                        For i = selCell.Row1 To selCell.Row2
                            ObjSel.RowHeight(i) = Abs(Int(-CDbl(txtAtt.Text) * Twip_mm * sgnMode)) / (selCell.Row2 - selCell.Row1 + 1)
                        Next
                    End If
                    '�޸Ĺ̶��е��и�
                    For Each tmpID In objReport.Items("_" & ObjSel.Index).SubIDs
                        Set tmpItem = objReport.Items("_" & tmpID.id)
                        arrHead = Split(tmpItem.��ͷ, "|")
                        tmpItem.��ͷ = ""
                        For i = 0 To UBound(arrHead)
                            If i >= selCell.Row1 And i <= selCell.Row2 Then
                                arrModify = Split(arrHead(i), "^")
                                tmpItem.��ͷ = tmpItem.��ͷ & "|" & arrModify(0) & "^" & ObjSel.RowHeight(i) / sgnMode & "^" & arrModify(2)
                            Else
                                tmpItem.��ͷ = tmpItem.��ͷ & "|" & arrHead(i)
                            End If
                        Next
                        tmpItem.��ͷ = Mid(tmpItem.��ͷ, 2)
                    Next
                    mshAtt.TextMatrix(mshAtt.Row, 1) = txtAtt.Text
                End If
                
                Call SetGridLine(intCurID) '���Ҫ����������
                If objReport.Items("_" & intCurID).���� = 4 Then Call SetCopyGrid(intCurID)
                txtAtt.Visible = False: mshAtt.SetFocus
                
                If GetSelNum = 1 Then ShowAttrib (intCurID)
                SetMainWH (ObjSel.Index)
                
                '�������ӱ��������������һ��
                If objReport.Items("_" & intCurID).���� = "" And objReport.Items("_" & intCurID).���� = 5 Then
                    For Each tmpItem In objReport.Items
                        If tmpItem.��ʽ�� = mbytCurrFmt And tmpItem.���� = objReport.Items("_" & intCurID).���� And tmpItem.���� = 5 Then
                            Call SetGridLike(msh(intCurID), msh(tmpItem.Key))
                        End If
                    Next
                End If
                BlnSave = False
            Case "����"
                If Not IsNumeric(txtAtt.Text) Then
                    MsgBox "���������������ݣ�", vbInformation, App.Title
                    txtAtt.SetFocus: Exit Sub
                End If
                If CDbl(txtAtt.Text) > 20 Then
                    MsgBox "�������ֵ����", vbInformation, App.Title
                    txtAtt.SetFocus: Exit Sub
                End If
                If objReport.Items("_" & intCurID).��ID <> 0 And txtAtt.Text <> "1" Then
                    MsgBox "��Ƭ�ڵı�����������", vbInformation, App.Title
                    txtAtt.SetFocus: txtAtt.Text = "1": Exit Sub
                End If
                txtAtt.Text = Abs(CDbl(txtAtt.Text))
                
                objReport.Items("_" & intCurID).���� = IIF(CByte(txtAtt.Text) < 2, 1, CByte(txtAtt.Text))
                
                mshAtt.TextMatrix(mshAtt.Row, 1) = objReport.Items("_" & intCurID).����
                
                'ɾ��ԭ�з���
                For Each tmpID In objReport.Items("_" & intCurID).CopyIDs
                    Unload msh(tmpID.id)
                Next
                Set objReport.Items("_" & intCurID).CopyIDs = New RelatIDs
                
                '��������
                For i = 1 To objReport.Items("_" & intCurID).���� - 1
                    intMaxID = intMaxID + 1
                    Load msh(intMaxID)
                    msh(intMaxID).Tag = "C_" & intCurID
                    msh(intMaxID).ToolTipText = "�� " & i & " ��"
                    
                    msh(intMaxID).Top = msh(intCurID).Top
                    msh(intMaxID).Left = msh(intCurID).Left + (msh(intCurID).Width - 15) * i
                    msh(intMaxID).Width = msh(intCurID).Width
                    msh(intMaxID).Height = msh(intCurID).Height
                    
                    Call SetGridSame(msh(intCurID), msh(intMaxID))
                    
                    msh(intMaxID).ZOrder
                    msh(intMaxID).Visible = True
                    objReport.Items("_" & intCurID).CopyIDs.Add intMaxID, "_" & intMaxID
                Next
                msh(intCurID).ZOrder
                lblSize(Mid(msh(intCurID).Tag, 3) + 2).ZOrder
                
                '������ǩ(ֻ���Ǳ�ǩ�������)�����
                Dim ResizeItem As RPTItem, IntSaveCurID As Integer
                IntSaveCurID = intCurID
                For Each ResizeItem In objReport.Items
                    If ResizeItem.���� = 2 And ResizeItem.���� = objReport.Items("_" & IntSaveCurID).���� And ResizeItem.��ʽ�� = mbytCurrFmt Then
                        intCurID = ResizeItem.Key
                        Call ReferTo
                    End If
                Next
                intCurID = IntSaveCurID
                
                txtAtt.Visible = False: mshAtt.SetFocus
                BlnSave = False
            Case "�о�"
                If Not IsNumeric(txtAtt.Text) Then
                    MsgBox "���������������ݣ�", vbInformation, App.Title
                    txtAtt.SetFocus: Exit Sub
                End If
                If CDbl(txtAtt.Text) > 100 Then
                    MsgBox "�������ֵ����", vbInformation, App.Title
                    txtAtt.SetFocus: Exit Sub
                End If
                txtAtt.Text = Abs(CDbl(txtAtt.Text))
                
                objReport.Items("_" & intCurID).���� = CByte(txtAtt.Text)
                
                mshAtt.TextMatrix(mshAtt.Row, 1) = objReport.Items("_" & intCurID).����
                
                txtAtt.Visible = False: mshAtt.SetFocus
                BlnSave = False
            Case "��ʽ"
                If TLen(txtAtt.Text) > 50 Then
                    MsgBox "��ʽ�г��Ȳ��ܳ���50���ַ���", vbInformation, App.Title
                    txtAtt.SetFocus: Exit Sub
                End If
                mshAtt.TextMatrix(mshAtt.Row, 1) = txtAtt.Text
                objReport.Items("_" & intCurID).��ʽ = txtAtt.Text
                txtAtt.Visible = False: mshAtt.SetFocus
                BlnSave = False
            Case "���¼��"
                If Not IsNumeric(txtAtt.Text) Then
                    MsgBox "���������������ݣ�", vbInformation, App.Title
                    txtAtt.SetFocus: Exit Sub
                End If
                txtAtt.Text = Abs(CDbl(txtAtt.Text))
                
                objReport.Items("_" & intCurID).���¼�� = Val(txtAtt.Text) * Twip_mm
                mshAtt.TextMatrix(mshAtt.Row, 1) = Format(objReport.Items("_" & intCurID).���¼�� / Twip_mm, "0.00")
                
                txtAtt.Visible = False: mshAtt.SetFocus
                BlnSave = False
            Case "���Ҽ��"
                If Not IsNumeric(txtAtt.Text) Then
                    MsgBox "���������������ݣ�", vbInformation, App.Title
                    txtAtt.SetFocus: Exit Sub
                End If
                txtAtt.Text = Abs(CDbl(txtAtt.Text))
                
                objReport.Items("_" & intCurID).���Ҽ�� = Val(txtAtt.Text) * Twip_mm
                mshAtt.TextMatrix(mshAtt.Row, 1) = Format(objReport.Items("_" & intCurID).���Ҽ�� / Twip_mm, "0.00")
                
                txtAtt.Visible = False: mshAtt.SetFocus
                BlnSave = False
            Case "�������"
                If Not IsNumeric(txtAtt.Text) Then
                    MsgBox "���������������ݣ�", vbInformation, App.Title
                    txtAtt.SetFocus: Exit Sub
                End If
                
                txtAtt.Text = Abs(CDbl(txtAtt.Text))
                If Val(txtAtt.Text) > 0 Then
                    lngL = (objReport.Fmts("_" & mbytCurrFmt).W - objReport.Items("_" & intCurID).X + objReport.Items("_" & intCurID).���Ҽ��) \ (objReport.Items("_" & intCurID).W + objReport.Items("_" & intCurID).���Ҽ��)
                    If Val(txtAtt.Text) > lngL Then
                        MsgBox "�������Ҽ�ֽ࣬�ź�������" & lngL & "����", vbInformation, App.Title
                        txtAtt.SetFocus: Exit Sub
                    End If
                End If
                
                objReport.Items("_" & intCurID).������� = Val(txtAtt.Text)
                mshAtt.TextMatrix(mshAtt.Row, 1) = objReport.Items("_" & intCurID).�������
                
                txtAtt.Visible = False: mshAtt.SetFocus
                BlnSave = False
            Case "�������"
                If Not IsNumeric(txtAtt.Text) Then
                    MsgBox "���������������ݣ�", vbInformation, App.Title
                    txtAtt.SetFocus: Exit Sub
                End If
                
                txtAtt.Text = Abs(CDbl(txtAtt.Text))
                If Val(txtAtt.Text) > 0 Then
                    lngW = (objReport.Fmts("_" & mbytCurrFmt).H - objReport.Items("_" & intCurID).Y + objReport.Items("_" & intCurID).���¼��) \ (objReport.Items("_" & intCurID).H + objReport.Items("_" & intCurID).���¼��)
                    If Val(txtAtt.Text) > lngW Then
                        MsgBox "�������¼�ֽ࣬����������" & lngW & "����", vbInformation, App.Title
                        txtAtt.SetFocus: Exit Sub
                    End If
                End If
                
                objReport.Items("_" & intCurID).������� = Val(txtAtt.Text)
                mshAtt.TextMatrix(mshAtt.Row, 1) = objReport.Items("_" & intCurID).�������
                
                txtAtt.Visible = False: mshAtt.SetFocus
                BlnSave = False
            
            Case "����Դ�к�"
                If Not IsNumeric(txtAtt.Text) Then
                    MsgBox "���������������ݣ�", vbInformation, App.Title
                    txtAtt.SetFocus: Exit Sub
                End If
                txtAtt.Text = Abs(CDbl(txtAtt.Text))
                
                objReport.Items("_" & intCurID).Դ�к� = Val(txtAtt.Text)
                mshAtt.TextMatrix(mshAtt.Row, 1) = objReport.Items("_" & intCurID).Դ�к�
                
                txtAtt.Visible = False: mshAtt.SetFocus
                BlnSave = False
            Case "��������"
                X = InStr(1, objReport.Items("_" & intCurID).����, "]")
                Y = InStr(1, objReport.Items("_" & intCurID).����, ".")
                k = InStr(1, objReport.Items("_" & intCurID).����, "[")
                If X > k And X > Y And X <> 0 And k <> 0 Then
                    strReportID = FindReport(txtAtt.Text, txtAtt.hwnd, strInfo, objReport.Items("_" & intCurID).Relations.Item(1).��������ID, objReport, objReport.Items("_" & intCurID).Relations, 2, Me, intCurID)
                    If strReportID <> "" Then
                        mshAtt.TextMatrix(mshAtt.Row, 1) = strInfo
                        mshAtt.RowData(mshAtt.Row) = strReportID
                        txtAtt.Visible = False: mshAtt.SetFocus
                        BlnSave = False
                    Else
                        txtAtt.SetFocus
                    End If
                Else
                    MsgBox "��ǰ��ǩ�����Ȱ�һ������Դ�����磺[����Դ.�ֶ�],�󶨺������ù�������", vbInformation, Me.Caption
                End If
        End Select
        Call AdjustAll
    Else
        Select Case mshAtt.TextMatrix(mshAtt.Row, 0)
            Case "����"
                If InStr("'|~^", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
            Case "X����", "Y����", "���", "�߶�", "�и�", "���¼��", "���Ҽ��"
                If InStr("-0.123456789" & Chr(8) & Chr(3) & Chr(22) & Chr(24) & Chr(26), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
            Case "����", "�о�", "����Դ�к�", "�������", "�������"
                If InStr("0123456789" & Chr(8) & Chr(3) & Chr(22) & Chr(24) & Chr(26), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
            Case "��ʽ"
                If InStr("'|~^", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
        End Select
    End If
End Sub

Private Function GetRow(str As String) As Integer
    Dim i As Integer
    For i = 1 To mshAtt.Rows - 1
        If UCase(Trim(mshAtt.TextMatrix(i, 0))) = UCase(Trim(str)) Then
            GetRow = i: Exit Function
        End If
    Next
    GetRow = 0
End Function

Private Sub InitRowHeight(intIdx As Integer)
'���ܣ���ʼ����и�Ϊ255,�Ա��ڳ����ֶ������и�(��Ȼ����������Զ�����)
'˵������Loadһ���µı���,һ��Ҫ���øù���
    Dim i As Integer
    For i = 0 To msh(intIdx).Rows - 1
        msh(intIdx).RowHeight(i) = 255 * sgnMode
    Next
End Sub

Private Sub SetGridLine(idx As Integer)
'���ܣ�����ָ���������������,�����и����,�������������
'˵��������ʱ�ؼ���Ӧ�����ݶ���(Item)�����Ѿ����ڣ��Ҷ�Ӧ�ؼ��Ѿ�����������ͷ���
    Dim blnPre As Boolean, SinH As Single
    Dim X As Integer, Y As Integer, Z As Integer
    Dim tmpID As RelatID, i As Integer, j As Integer
    Dim intFixHeight As Integer

    blnPre = msh(idx).Redraw
    msh(idx).Redraw = False
    
    If objReport.Items("_" & idx).���� = 4 Then '���ܱ��
        '�������������������
        If objReport.Ʊ�� Then
            SinH = 0: X = msh(idx).FixedRows
            For i = 0 To msh(idx).FixedRows - 1
                SinH = SinH + msh(idx).RowHeight(i)
            Next
            msh(idx).Rows = Abs(Int((-(msh(idx).Height - SinH)) / (objReport.Items("_" & idx).�и� * sgnMode))) + X
            If msh(idx).Rows = X Then msh(idx).Rows = msh(idx).Rows + 2
            msh(idx).FixedRows = X
            For i = msh(idx).FixedRows To msh(idx).Rows - 1
                msh(idx).RowHeight(i) = objReport.Items("_" & idx).�и� * sgnMode
            Next
        End If
    ElseIf objReport.Items("_" & idx).���� = 5 Then '���ܱ��
        X = msh(idx).FixedCols '���������Ŀ��
        Y = msh(idx).FixedRows - 1 '���������Ŀ��
        For Each tmpID In objReport.Items("_" & idx).SubIDs
            If objReport.Items("_" & tmpID.id).���� = 9 Then Z = Z + 1 'ͳ����Ŀ��
        Next
        
        For i = 0 To msh(idx).FixedRows - 1
            intFixHeight = intFixHeight + msh(idx).RowHeight(i)
        Next
        '���ܱ���������������
        msh(idx).Rows = Abs(Int(-(msh(idx).Height - intFixHeight) / msh(idx).RowHeight(msh(idx).FixedRows))) + 2
        If msh(idx).Rows < msh(idx).FixedRows + 3 Then msh(idx).Rows = msh(idx).FixedRows + 3
        
        For i = msh(idx).FixedRows + 1 To msh(idx).Rows - 1
'            msh(idx).RowHeight(i) = msh(idx).RowHeight(0)
            For j = 0 To msh(idx).FixedCols - 1
                msh(idx).TextMatrix(i, j) = msh(idx).TextMatrix(msh(idx).FixedRows + 2, j)
            Next
        Next
        
        '����к�����ࣺ���ܱ���������ͳ����
        If msh(idx).FixedRows > 1 Then
            X = 0
            For i = 0 To msh(idx).FixedCols - 1
                X = X + msh(idx).ColWidth(i) '��������ܿ��
            Next
            Y = 0
            For i = msh(idx).FixedCols To msh(idx).FixedCols + Z - 1
                Y = Y + msh(idx).ColWidth(i) 'һ��ͳ�����ܿ��
            Next
            '���� = ͳ������ * ÿ������ + �����������
            msh(idx).Cols = Abs(Int(-(msh(idx).Width - X) / Y)) * Z + msh(idx).FixedCols
            'ÿ���ȼ�������ͬ
            For i = msh(idx).FixedCols + Z To msh(idx).Cols - 1
                For j = 0 To msh(idx).FixedRows - 2
                    msh(idx).TextMatrix(j, i) = msh(idx).TextMatrix(j, msh(idx).FixedCols + 1)
                Next
            Next
            For i = msh(idx).FixedCols + Z To msh(idx).Cols - 1 Step Z
                For j = 1 To Z
                    msh(idx).TextMatrix(msh(idx).FixedRows - 1, i + j - 1) = _
                    msh(idx).TextMatrix(msh(idx).FixedRows - 1, msh(idx).FixedCols + j - 1)
                    msh(idx).ColWidth(i + j - 1) = msh(idx).ColWidth(msh(idx).FixedCols + j - 1)
                    msh(idx).ColAlignment(i + j - 1) = msh(idx).ColAlignment(msh(idx).FixedCols + j - 1)
                Next
            Next
        End If
    End If
    
    msh(idx).Redraw = blnPre
End Sub

Private Sub ReplaceName(strPre As String, strNew As String)
'���ܣ��������漰strPre����Դ���Ƶı���Ԫ�ص���Ӧ�����滻��������Դ����strNew
    Dim tmpItem As RPTItem
    Dim i As Integer, j As Integer
    
    If strPre = strNew Then Exit Sub
    
    For Each tmpItem In objReport.Items
        Select Case tmpItem.����
            Case 2, 3, 11, 14 '��ǩ
                If InStr(tmpItem.����, strPre & ".") > 0 Then
                    tmpItem.���� = Replace(tmpItem.����, strPre & ".", strNew & ".")
                    If tmpItem.���� <> 11 And tmpItem.��ʽ�� = mbytCurrFmt Then lbl(tmpItem.id).Caption = Replace(lbl(tmpItem.id).Caption, strPre & ".", strNew & ".")
                End If
            Case 4 '������
                If tmpItem.��ʽ�� = mbytCurrFmt Then
                    For j = 0 To msh(tmpItem.id).Rows - 1
                        For i = 0 To msh(tmpItem.id).Cols - 1
                            If InStr(msh(tmpItem.id).TextMatrix(j, i), strPre & ".") > 0 Then
                                msh(tmpItem.id).TextMatrix(j, i) = _
                                    Replace(msh(tmpItem.id).TextMatrix(j, i), strPre & ".", strNew & ".")
                            End If
                        Next
                    Next
                End If
                
                For i = 1 To tmpItem.SubIDs.count
                    objReport.Items("_" & tmpItem.SubIDs(i).Key).��ͷ = Replace(objReport.Items("_" & tmpItem.SubIDs(i).Key).��ͷ, strPre & ".", strNew & ".")
                    objReport.Items("_" & tmpItem.SubIDs(i).Key).���� = Replace(objReport.Items("_" & tmpItem.SubIDs(i).Key).����, strPre & ".", strNew & ".")
                Next
                If tmpItem.��ʽ�� = mbytCurrFmt Then
                    Call SetCopyGrid(tmpItem.id)
                End If
            Case 5 '���ܱ��
                If tmpItem.���� = strPre Then
                    tmpItem.���� = strNew
                End If
            Case 6 '��������
                If InStr(tmpItem.����, strPre & ".") > 0 Then
                    tmpItem.���� = Replace(tmpItem.����, strPre & ".", strNew & ".")
                End If
            Case 12
                If tmpItem.���� <> "" Then
                    tmpItem.���� = Mid(Replace("|" & tmpItem.����, "|" & strPre & ".", "|" & strNew & "."), 2)
                End If
        End Select
    Next
End Sub

Private Function CheckData() As String
'���ܣ����Ԫ���е�����Ԫ���Ƿ����ҵ���Ӧ������Դ
'���أ���ʾ��Ϣ
    Dim tmpItem As RPTItem, objNode As Object
    Dim strFX As String, strFS As String, strFY As String, strDBConn As String
    Dim strData As String, blnExist As Boolean
    Dim lngL As Long, lngW As Long
    
    For Each tmpItem In objReport.Items
        blnExist = False
        Select Case tmpItem.����
            Case 2, 3 '��ǩ
                'ֻ��ĩ�����(����Դ="[����Դ��.�ֶ�]"
                blnExist = CheckText(tmpItem)
                If Not blnExist Then CheckData = "������Դ���Ҳ������ݱ�ǩ������""" & tmpItem.���� & """,�봦������Դ���ǩ��": Exit Function
            Case 5  '���ܱ��(����Դ��)
                'ֻ���м���
                blnExist = CheckText(tmpItem)
                If Not blnExist Then CheckData = "������Դ���Ҳ������ܱ�������""" & tmpItem.���� & """,�봦������Դ����": Exit Function
            Case 6 '����������
                'ֻ��ĩ�����(���������п����й�ʽ)
                If GetItemCount(tmpItem.����) > 0 Then
                    blnExist = CheckText(tmpItem)
                    If Not blnExist Then CheckData = "������Դ���Ҳ����������е�����""" & tmpItem.���� & """,�봦������Դ����": Exit Function
                End If
            Case 7, 8, 9 '���ܱ������(ֻ�����ֶ���)
                'ֻ��ĩ�����
                For Each objNode In tvwSQL.Nodes
                    If objNode.Key <> "Root" And objNode.Children = 0 Then
                        strDBConn = objReport.Items("_" & tmpItem.�ϼ�ID).����
                        If strDBConn Like "*��*��" Then
                            strDBConn = Left(strDBConn, InStrRev(strDBConn, "��") - 1)
                        End If
                        If LevelText(objNode) = strDBConn & "." & tmpItem.���� Then blnExist = True: Exit For
                    End If
                Next
                If Not blnExist Then CheckData = "������Դ���Ҳ������ܱ�������������""" & tmpItem.���� & """,�봦������Դ����": Exit Function
                
                If tmpItem.���� <> "" Then
                    blnExist = False
                    For Each objNode In tvwSQL.Nodes
                        If objNode.Key <> "Root" And objNode.Children = 0 Then
                            If LevelText(objNode) = mdlPublic.GetStdNodeText(objReport.Items("_" & tmpItem.�ϼ�ID).����) & "." & _
                                IIF(Left(tmpItem.����, 1) = ",", Mid(tmpItem.����, 2), tmpItem.����) Then blnExist = True: Exit For
                        End If
                    Next
                    If Not blnExist Then CheckData = "������Դ���Ҳ������ܱ�������������""" & IIF(Left(tmpItem.����, 1) = ",", Mid(tmpItem.����, 2), tmpItem.����) & """,�봦������Դ����": Exit Function
                End If
            Case 12 '@@@
                If tmpItem.���� <> "" Then
                    Call GetChartDataName(tmpItem.����, strFX, strFS, strFY, strData)
                    If strFX <> "" Then
                        blnExist = False
                        For Each objNode In tvwSQL.Nodes
                            If objNode.Key <> "Root" And objNode.Children = 0 Then
                                If LevelText(objNode) = strData & "." & strFX Then blnExist = True: Exit For
                            End If
                        Next
                        If Not blnExist Then CheckData = "������Դ���Ҳ���ͼ��""" & tmpItem.���� & """�ģ�ֵ�ֶΣ��봦������Դ��ͼ��": Exit Function
                    End If
                    
                    If strFS <> "" Then
                        blnExist = False
                        For Each objNode In tvwSQL.Nodes
                            If objNode.Key <> "Root" And objNode.Children = 0 Then
                                If LevelText(objNode) = strData & "." & strFS Then blnExist = True: Exit For
                            End If
                        Next
                        If Not blnExist Then CheckData = "������Դ���Ҳ���ͼ��""" & tmpItem.���� & """�������ֶΣ��봦������Դ��ͼ��": Exit Function
                    End If
                    
                    If strFY <> "" Then
                        blnExist = False
                        For Each objNode In tvwSQL.Nodes
                            If objNode.Key <> "Root" And objNode.Children = 0 Then
                                If LevelText(objNode) = strData & "." & strFY Then blnExist = True: Exit For
                            End If
                        Next
                        If Not blnExist Then CheckData = "������Դ���Ҳ���ͼ��""" & tmpItem.���� & """�ģ�ֵ�ֶΣ��봦������Դ��ͼ��": Exit Function
                    End If
                End If
            Case 13 '����
                'ֻ��ĩ�����(����Դ="[����Դ��.�ֶ�]"
                blnExist = CheckText(tmpItem)
                If Not blnExist Then CheckData = "������Դ���Ҳ������������""" & tmpItem.���� & """,�봦������Դ���������ݣ�": Exit Function
            Case 14 '��Ƭ
                If tmpItem.������� > 0 Then
                    lngW = (objReport.Fmts("_" & mbytCurrFmt).H - tmpItem.Y + tmpItem.���¼��) \ (tmpItem.H + tmpItem.���¼��)
                    If tmpItem.������� > lngW Then
                        CheckData = "�������¼�ֽ࣬����������" & lngW & "�������鿨Ƭ���ԡ�"
                        Exit Function
                    End If
                End If
                If tmpItem.������� > 0 Then
                    lngL = (objReport.Fmts("_" & mbytCurrFmt).W - tmpItem.X + tmpItem.���Ҽ��) \ (tmpItem.W + tmpItem.���Ҽ��)
                    If tmpItem.������� > lngL Then
                        CheckData = "�������Ҽ�ֽ࣬�ź�������" & lngL & "�������鿨Ƭ���ԡ�"
                        Exit Function
                    End If
                End If
        End Select
    Next
End Function

Private Function CheckArea() As String
'���ܣ�����Ƿ�����Ԫ�ض��ڱ����߷�Χ֮��,�Լ����������Ƿ������ʾ�ꡣ
    Dim tmpItem As RPTItem, bytFmt As Byte
    Dim StrFmt As String, objFmt As RPTFmt
    Dim lngW As Long, lngH As Long
    Dim strTmp As String
    
    Call ReFlashWidth
    
    For Each tmpItem In objReport.Items
        With tmpItem
            If InStr(1, "|1,|2,|3,|4,|5,|10,|11,|12,|14,", "|" & .���� & ",") <> 0 Then
                Set objFmt = objReport.Fmts("_" & .��ʽ��)
                If tmpItem.��ID = 0 Then
                    If objFmt.ֽ�� = 1 Then
                        lngW = objFmt.W
                        lngH = objFmt.H
                    Else
                        lngW = objFmt.H
                        lngH = objFmt.W
                    End If
                    strTmp = "ֽ��"
                Else
                    lngW = objReport.Items("_" & tmpItem.��ID).W
                    lngH = objReport.Items("_" & tmpItem.��ID).H
                    strTmp = "��Ƭ"
                End If
                StrFmt = objFmt.˵��
                If .X < 0 Or .Y < 0 Or (.X + .W) > lngW Or (.Y + .H) > lngH Then
                    
                    Select Case .����
                        Case 1
                            CheckArea = "��ʽ[" & StrFmt & "]��ĳ��������λ�ó�����" & strTmp & "�ĳߴ緶Χ,����������"
                        Case 2, 3
                            CheckArea = "��ʽ[" & StrFmt & "]��ĳ����ǩ��λ�ó�����" & strTmp & "�ĳߴ緶Χ,����������"
                        Case 4
                            CheckArea = "��ʽ[" & StrFmt & "]��ĳ���������λ�ó�����" & strTmp & "�ĳߴ緶Χ,����������"
                        Case 5
                            CheckArea = "��ʽ[" & StrFmt & "]��ĳ�����ܱ���λ�ó�����" & strTmp & "�ĳߴ緶Χ,����������"
                        Case 10
                            CheckArea = "��ʽ[" & StrFmt & "]��ĳ�����ߵ�λ�ó�����" & strTmp & "�ĳߴ緶Χ,����������"
                        Case 11
                            CheckArea = "��ʽ[" & StrFmt & "]��ĳ��ͼƬ��λ�ó�����" & strTmp & "�ĳߴ緶Χ,����������"
                        Case 12
                            CheckArea = "��ʽ[" & StrFmt & "]��ĳ��ͼ���λ�ó�����" & strTmp & "�ĳߴ緶Χ,����������"
                        Case 14
                            CheckArea = "��ʽ[" & StrFmt & "]��ĳ����Ƭ��λ�ó�����" & strTmp & "�ĳߴ緶Χ,����������"
                    End Select
                    Exit Function
                End If
                If .���� > 1 And .���� = 4 Then
                    If .X + .W * .���� > lngW Then
                        CheckArea = "������ĳ��������ķ�������,������" & strTmp & "�ĳߴ緶Χ,����������"
                        Exit Function
                    End If
                End If
            End If
        End With
    Next
End Function

Private Sub ShowItems()
'���ܣ�����objReport������ʾ����Ԫ��
    Dim tmpItem As RPTItem, bytFormat As Byte
    
    '����ʾͼ���Լ�������
    For Each tmpItem In objReport.Items
        If tmpItem.��ʽ�� = mbytCurrFmt And tmpItem.���� = 12 Then
            Call ShowItem(tmpItem.id)
        End If
    Next

    For Each tmpItem In objReport.Items
        If tmpItem.��ʽ�� = mbytCurrFmt And tmpItem.���� <> 12 Then
            Call ShowItem(tmpItem.id)
        End If
    Next
End Sub

Private Sub LoadReportFormat()
    Dim tmpFmt As RPTFmt
    Dim objItem As ComboItem
    
    With cboFormat
        blnAllowIn = False
        .ComboItems.Clear
        
        For Each tmpFmt In objReport.Fmts
            Set objItem = .ComboItems.Add(, "_" & tmpFmt.���, tmpFmt.˵��, "Root")
            If tmpFmt.��� = mbytCurrFmt Then
                objItem.Selected = True
                Set .SelectedItem = objItem
                .SelectedItem.Selected = True
            End If
        Next
        If .ComboItems.count > 0 And .SelectedItem Is Nothing Then
            .ComboItems(1).Selected = True
            Set .SelectedItem = .ComboItems(1)
            .SelectedItem.Selected = True
        End If
        mbytCurrFmt = Mid(cboFormat.SelectedItem.Key, 2)
        
        cmdAdd.Enabled = blnDelReportFormat
        cmdDel.Enabled = (.ComboItems.count > 1) And blnDelReportFormat
        tbr1.Buttons("AddFormat").Enabled = cmdAdd.Enabled
        tbr1.Buttons("DelFormat").Enabled = cmdDel.Enabled
        mnuEdit_AddFormat.Enabled = cmdAdd.Enabled
        mnuEdit_DelFormat.Enabled = cmdDel.Enabled
        
        blnAllowIn = True
    End With
End Sub

Private Sub ShowItem(idx As Integer)
'���ܣ���ʾָ���ı���Ԫ��(ShowItems���Ӻ���,Ҳ�ɵ�������)
'������idx=objReport�е�Ԫ������
    Dim i As Integer, j As Integer, tmpID As RelatID
    Dim ObjSel As Control
    Dim objBarCode As StdPicture, strBarCode As String
    Dim lngSize As Long, sngWidth As Single
    
    With objReport.Items("_" & idx)
        Select Case .����
            Case 1 '����
                Load lblLine(.id)
                Set ObjSel = lblLine(.id)
                ObjSel.Top = Format(.Y * sgnMode, "0.00")
                ObjSel.Left = Format(.X * sgnMode, "0.00")
                ObjSel.Height = Format(.H * sgnMode, "0.00")
                ObjSel.Width = Format(.W * sgnMode, "0.00")
                ObjSel.BackColor = .ǰ��
                If .���� Then ObjSel.Height = 30
                ObjSel.ZOrder
                ObjSel.Visible = True
            Case 2, 3 '��ǩ
                Load lbl(.id)
                Set ObjSel = lbl(.id)
                ObjSel.Top = Format(.Y * sgnMode, "0.00")
                ObjSel.Left = Format(.X * sgnMode, "0.00")
                ObjSel.Height = Format(.H * sgnMode, "0.00")
                ObjSel.Width = Format(.W * sgnMode, "0.00")
                ObjSel.ForeColor = .ǰ��
                ObjSel.BackColor = IIF(.���� = &HFFFFFF, lbl(0).BackColor, .����)
                ObjSel.Font.name = .����
                ObjSel.Font.Size = Format(.�ֺ� * sgnMode, "0.0")
                ObjSel.Font.Bold = .����
                ObjSel.Font.Italic = .б��
                ObjSel.Font.Underline = .����
                ObjSel.BorderStyle = IIF(.�߿�, 1, 0)
                ObjSel.Alignment = IIF(.���� <> 0, IIF(.���� = 1, 2, 1), 0)
                ObjSel.Caption = .����
                
                If Not ItemIsGraph(.id) Then
                    ObjSel.AutoSize = .�Ե�
                End If
                
                If InStr(1, "|11,", "|" & .���� & ",") <> 0 Then
                    ObjSel.BorderStyle = 1
                    ObjSel.BackStyle = 0
                    If .���� = 10 Then ObjSel.Caption = ""
                    
                    Call DrawFrame(ObjSel)
                End If
                ObjSel.ZOrder 0
                ObjSel.Visible = True
            Case 10 '����
                Load Shp(.id)
                Set ObjSel = Shp(.id)
                Load lblshp(.id)
                lblshp(.id).BackColor = picPaper.BackColor
                ObjSel.Top = Format(.Y * sgnMode, "0.00")
                ObjSel.Left = Format(.X * sgnMode, "0.00")
                ObjSel.Height = Format(.H * sgnMode, "0.00")
                ObjSel.Width = Format(.W * sgnMode, "0.00")
                lblshp(.id).Top = ObjSel.Top
                lblshp(.id).Left = ObjSel.Left
                lblshp(.id).Width = ObjSel.Width
                lblshp(.id).Height = ObjSel.Height
                ObjSel.BorderColor = .ǰ��
                ObjSel.BackColor = IIF(.���� = &HFFFFFF, Shp(0).BackColor, .����)
                ObjSel.BorderStyle = 1
                ObjSel.BackStyle = 0
                ObjSel.BorderWidth = IIF(.����, 2, 1)
                ObjSel.Shape = IIF(.�߿�, ShapeConstants.vbShapeOval, ShapeConstants.vbShapeRectangle)
                
                ObjSel.ZOrder 1
                ObjSel.Visible = True
                lblshp(.id).ZOrder 1
                lblshp(.id).Visible = True
            Case 4, 5 '������,���ܱ��
                Load msh(.id)
                Set ObjSel = msh(.id)
                '��ʽ����
                ObjSel.Top = Format(.Y * sgnMode, "0.00")
                ObjSel.Left = Format(.X * sgnMode, "0.00")
                ObjSel.Height = Format(.H * sgnMode, "0.00")
                ObjSel.Width = Format(.W * sgnMode, "0.00")
                ObjSel.Font.Size = Format(.�ֺ� * sgnMode, "0.0")
                '��������(����CopyIDs�Ѿ�����)
                i = 0
                For Each tmpID In .CopyIDs
                    i = i + 1
                    Load msh(tmpID.id)
                    msh(tmpID.id).Width = ObjSel.Width
                    msh(tmpID.id).Height = ObjSel.Height
                    msh(tmpID.id).Top = ObjSel.Top
                    msh(tmpID.id).Left = ObjSel.Left + (ObjSel.Width - 15) * i
                    msh(tmpID.id).Font.Size = ObjSel.Font.Size
                    msh(tmpID.id).Tag = "C_" & .id
                    msh(tmpID.id).ZOrder
                    msh(tmpID.id).Visible = True
                Next
                
                Call ReShowGrid(.id)
                If .���� = 4 Then Call CustomColColor(.id, -9)
                ObjSel.ZOrder
                ObjSel.Visible = True
            Case 11
                Load img(.id)
                Set ObjSel = img(.id)
                
                ObjSel.Top = Format(.Y * sgnMode, "0.00")
                ObjSel.Left = Format(.X * sgnMode, "0.00")
                ObjSel.Height = Format(.H * sgnMode, "0.00")
                ObjSel.Width = Format(.W * sgnMode, "0.00")
                ObjSel.BorderStyle = IIF(.�߿�, 1, 0)
                
                '���ֱ���
                If Not .ͼƬ Is Nothing Then
                    If .���� Then
                        Set ObjSel.Picture = ScalePicture(PicFontTest, .ͼƬ, ObjSel.Width, ObjSel.Height)
                    Else
                        Set ObjSel.Picture = .ͼƬ
                    End If
                End If
                
                ObjSel.ZOrder
                ObjSel.Visible = True
            Case 14
                Load pic(.id)
                Set ObjSel = pic(.id)
                
                ObjSel.Top = Format(.Y * sgnMode, "0.00")
                ObjSel.Left = Format(.X * sgnMode, "0.00")
                ObjSel.Height = Format(.H * sgnMode, "0.00")
                ObjSel.Width = Format(.W * sgnMode, "0.00")
                ObjSel.BorderStyle = IIF(.�߿�, 1, 0)
                
                ObjSel.ZOrder
                ObjSel.Visible = True
            Case 12 '@@@
                Load Chart(.id)
                Set ObjSel = Chart(.id)
                
                ObjSel.Top = Format(.Y * sgnMode, "0.00")
                ObjSel.Left = Format(.X * sgnMode, "0.00")
                ObjSel.Height = Format(.H * sgnMode, "0.00")
                ObjSel.Width = Format(.W * sgnMode, "0.00")
                
                Call SetChartStyleAndData(ObjSel, objReport.Items("_" & idx), , sgnMode, True)
                
                ObjSel.ZOrder
                ObjSel.Visible = True
            Case 13
                Load ImgCode(.id)
                Set ObjSel = ImgCode(.id)
                
                ObjSel.Top = Format(.Y * sgnMode, "0.00")
                ObjSel.Left = Format(.X * sgnMode, "0.00")
                ObjSel.Height = Format(.H * sgnMode, "0.00")
                ObjSel.Width = Format(.W * sgnMode, "0.00")
                ObjSel.BorderStyle = 0
                
                '��ʾ����ͼ��
                strBarCode = ReplaceBracket(.����)
                If strBarCode = "" Then strBarCode = "1234567890"
                
                Unload frmFlash 'ǿ�Ƴ�ʼPicture����Ȼ�л�����������
                If .��� = 1 Then
                    Set objBarCode = DrawBarCode128(frmFlash.picTemp, 3, strBarCode, Mid(.��ͷ, 1, 1) = "1")
                ElseIf .��� = 2 Then
                    Set objBarCode = DrawBarCode39(frmFlash.picTemp, 3, strBarCode, Mid(.��ͷ, 2, 1) = "1", Mid(.��ͷ, 1, 1) = "1")
                ElseIf .��� = 3 Then
                    Set objBarCode = DrawBarCode128Auto(frmFlash.picTemp, strBarCode, sngWidth, .�и�, Mid(.��ͷ, 1, 1) = "1")
                ElseIf .��� = 10 Then
                    Set objBarCode = DrawBarCode2D(strBarCode, frmFlash.picTemp, lngSize)
                End If
                If Val(Mid(.��ͷ, 3, 1)) <> 0 Then
                    Set objBarCode = PictureSpin(objBarCode, Val(Mid(.��ͷ, 3, 1)), frmFlash.picTemp)
                End If
                Set ObjSel.Picture = objBarCode
                
                If .��� = 3 Then
                    '128���Զ��������
                    If Val(Mid(.��ͷ, 3, 1)) = 0 Then
                        ObjSel.Width = Format(Me.ScaleX(sngWidth, vbMillimeters, vbTwips) * sgnMode, "0.00")
                        .W = Me.ScaleX(sngWidth, vbMillimeters, vbTwips)
                    Else
                        ObjSel.Height = Format(Me.ScaleY(sngWidth, vbMillimeters, vbTwips) * sgnMode, "0.00")
                        .H = Me.ScaleY(sngWidth, vbMillimeters, vbTwips)
                    End If
                ElseIf .��� = 10 And .�Ե� Then
                    '��ά����ȱʡ�Զ�������С
                    ObjSel.Width = Format(lngSize * sgnMode, "0.00")
                    ObjSel.Height = Format(lngSize * sgnMode, "0.00")
                    .W = lngSize: .H = lngSize
                End If

                ObjSel.ZOrder
                ObjSel.Visible = True
        End Select
        If .��ID <> 0 And InStr(",14,5,6,12,", "," & .���� & ",") = 0 Then
            Set ObjSel.Container = pic(.��ID)
        End If
    End With
End Sub

Private Sub ReShowGrid(idx As Integer)
'���ܣ�����objReport���������»��Ʊ������,��ʱˢ�·����ؼ�
'˵����1.objReport���������Ѵ���,2.��Ӧ�ؼ��Ѵ���

    Dim i As Integer, j As Integer, X As Integer, Y As Integer, Z As Integer
    Dim tmpID As RelatID, tmpItem As RPTItem, strCaption As String, sgnH As Long
    
    msh(idx).Redraw = False
    msh(idx).Clear
    With objReport.Items("_" & idx)
        If .���� = 4 Then '������
            '��ʽ����(λ�ü��ߴ粻��)
            msh(idx).ForeColor = .ǰ��
            msh(idx).ForeColorFixed = .ǰ��
            msh(idx).GridColor = .����
            msh(idx).GridColorFixed = IIF(.��ʽ = "", .����, Val(.��ʽ))
            
            msh(idx).BackColor = .����
            msh(idx).BackColorFixed = IIF(.���� = &HFFFFFF, lbl(0).BackColor, .����)
            
            msh(idx).Font.name = .����
            msh(idx).Font.Size = Format(.�ֺ� * sgnMode, "0.0")
            msh(idx).Font.Bold = .����
            msh(idx).Font.Italic = .б��
            msh(idx).Font.Underline = .����
            msh(idx).GridLineWidth = IIF(.����߼Ӵ�, 2, 1)
            '��������
            '����
            msh(idx).Cols = .SubIDs.count
            msh(idx).FixedCols = 0
            i = 0
            For Each tmpID In .SubIDs
                Set tmpItem = objReport.Items("_" & tmpID.id)
                
                If i = 0 Then '��С����
                    If objReport.Ʊ�� = False Then
                        msh(idx).Rows = UBound(Split(tmpItem.��ͷ, "|")) + 3
                        msh(idx).FixedRows = UBound(Split(tmpItem.��ͷ, "|")) + 1
                    Else
                        msh(idx).Rows = UBound(Split(tmpItem.��ͷ, "|")) + 3
                        msh(idx).FixedRows = UBound(Split(tmpItem.��ͷ, "|")) + 1
                    End If
                End If

                '����������
                msh(idx).ColWidth(tmpItem.���) = tmpItem.W * sgnMode
                msh(idx).ColAlignment(tmpItem.���) = Switch(tmpItem.���� = 0, 1, tmpItem.���� = 1, 4, tmpItem.���� = 2, 7)
                msh(idx).TextMatrix(msh(idx).FixedRows, tmpItem.���) = tmpItem.����
                msh(idx).TextMatrix(msh(idx).FixedRows + 1, tmpItem.���) = tmpItem.����
                                    
                '�Զ����ͷ����
                For i = 0 To msh(idx).FixedRows - 1
                    On Error Resume Next
                    
                    Err = 0
                    strCaption = Split(Split(tmpItem.��ͷ, "|")(i), "^")(2)
                    If Err <> 0 Then strCaption = ""
                    If strCaption = "#" Then
                        msh(idx).TextMatrix(i, tmpItem.���) = ""
                    ElseIf strCaption = "��" Then
                        msh(idx).TextMatrix(i, tmpItem.���) = msh(idx).TextMatrix(i, tmpItem.��� - 1)
                    ElseIf strCaption = "��" Then
                        msh(idx).TextMatrix(i, tmpItem.���) = msh(idx).TextMatrix(i - 1, tmpItem.���)
                    Else
                        msh(idx).TextMatrix(i, tmpItem.���) = strCaption
                    End If
                    
                    Err = 0
                    sgnH = Split(Split(tmpItem.��ͷ, "|")(i), "^")(1)
                    If Err <> 0 Then sgnH = 250
                    msh(idx).RowHeight(i) = sgnH * sgnMode
                    msh(idx).Row = i
                    msh(idx).Col = tmpItem.���
                    Err = 0
                    sgnH = Split(Split(tmpItem.��ͷ, "|")(i), "^")(0)
                    If Err <> 0 Then sgnH = 4
                    msh(idx).CellAlignment = sgnH
                    If UBound(Split(Split(tmpItem.��ͷ, "|")(i), "^")) > 2 Then
                        msh(idx).CellFontBold = Split(Split(tmpItem.��ͷ, "|")(i), "^")(3)
                        msh(idx).CellForeColor = Split(Split(tmpItem.��ͷ, "|")(i), "^")(4)
                    End If
                Next
            Next
            
            For i = msh(idx).FixedRows To msh(idx).Rows - 1
                msh(idx).RowHeight(i) = .�и� * sgnMode
            Next
            
            '�ϲ�����
            For i = 0 To msh(idx).FixedRows - 1
                msh(idx).MergeRow(i) = True
            Next
            For i = 0 To msh(idx).Cols - 1
                msh(idx).MergeCol(i) = True
            Next
            
'            Call SetHeadCenter(msh(idx)) '��ͷ���ݾ���
            Call SetGridLine(.id) '�������
            
            '��������(����CopyIDs�Ѿ�����)
            For Each tmpID In .CopyIDs
                Call SetGridSame(msh(idx), msh(tmpID.id))
            Next
        ElseIf .���� = 5 Then '���ܱ��
            msh(idx).ForeColor = .ǰ��
            msh(idx).ForeColorFixed = .ǰ��
            msh(idx).GridColor = .����
            msh(idx).GridColorFixed = .����
            
            msh(idx).BackColor = .����
            msh(idx).BackColorFixed = IIF(.���� = &HFFFFFF, lbl(0).BackColor, .����)
            
            msh(idx).Font.name = .����
            msh(idx).Font.Size = Format(.�ֺ� * sgnMode, "0.0")
            msh(idx).Font.Bold = .����
            msh(idx).Font.Italic = .б��
            msh(idx).Font.Underline = .����
            msh(idx).GridLineWidth = IIF(.����߼Ӵ�, 2, 1)
            
            X = 0: Y = 0: Z = 0
            For Each tmpID In .SubIDs
                Select Case objReport.Items("_" & tmpID.id).����
                    Case 7
                        X = X + 1 '���������
                    Case 8
                        Y = Y + 1 '���������
                    Case 9
                        Z = Z + 1 'ͳ������
                End Select
            Next
            '��С������
            msh(idx).Rows = Y + 4
            msh(idx).FixedRows = Y + 1
            If Y = 0 Then
                msh(idx).Cols = X + Z
            Else
                msh(idx).Cols = X + IIF(Z = 1, Z + 1, Z)
            End If
            msh(idx).FixedCols = X
            msh(idx).RowHeight(0) = .�и� * sgnMode '�и�0�Ǳ�׼
            msh(idx).RowHeightMin = msh(idx).RowHeight(0)
            
            '������������
            For Each tmpID In .SubIDs
                Set tmpItem = objReport.Items("_" & tmpID.id)
                Select Case tmpItem.����
                    Case 7 '�������
                        msh(idx).TextMatrix(msh(idx).FixedRows - 1, tmpItem.���) = "[" & tmpItem.���� & "]"
                        msh(idx).Cell(flexcpFontBold, msh(idx).FixedRows - 1, tmpItem.���) = tmpItem.����
                        msh(idx).Cell(flexcpForeColor, msh(idx).FixedRows - 1, tmpItem.���) = tmpItem.ǰ��
                        
                        For i = msh(idx).FixedRows To msh(idx).Rows - 1
                            msh(idx).TextMatrix(i, tmpItem.���) = tmpItem.����
                        Next
                        If tmpItem.���� <> "" Then
                            msh(idx).TextMatrix(msh(idx).FixedRows, tmpItem.���) = tmpItem.����
                        End If
                        
                        msh(idx).ColWidth(tmpItem.���) = tmpItem.W * sgnMode
                        msh(idx).ColAlignment(tmpItem.���) = Switch(tmpItem.���� = 0, 1, tmpItem.���� = 1, 4, tmpItem.���� = 2, 7)
                    Case 8 '�������
                        For i = 0 To msh(idx).FixedCols - 1
                            msh(idx).TextMatrix(tmpItem.���, i) = "[" & tmpItem.���� & "]"
                            msh(idx).Cell(flexcpFontBold, tmpItem.���, i) = tmpItem.����
                            msh(idx).Cell(flexcpForeColor, tmpItem.���, i) = tmpItem.ǰ��
                        Next
                        
                        For i = msh(idx).FixedCols To msh(idx).Cols - 1
                            msh(idx).TextMatrix(tmpItem.���, i) = tmpItem.����
                        Next
                        If tmpItem.���� <> "" Then
                            msh(idx).TextMatrix(tmpItem.���, msh(idx).FixedCols) = tmpItem.����
                        End If
                    Case 9 'ͳ����
                        msh(idx).TextMatrix(msh(idx).FixedRows - 1, msh(idx).FixedCols + tmpItem.���) = "[" & tmpItem.���� & "]"
                        msh(idx).ColWidth(msh(idx).FixedCols + tmpItem.���) = tmpItem.W * sgnMode
                        msh(idx).ColAlignment(msh(idx).FixedCols + tmpItem.���) = Switch(tmpItem.���� = 0, 1, tmpItem.���� = 1, 4, tmpItem.���� = 2, 7)
                        msh(idx).Cell(flexcpFontBold, msh(idx).FixedRows - 1, msh(idx).FixedCols + tmpItem.���, msh(idx).Rows - 1, msh(idx).FixedCols + tmpItem.���) = tmpItem.����
                        msh(idx).Cell(flexcpForeColor, msh(idx).FixedRows - 1, msh(idx).FixedCols + tmpItem.���, msh(idx).Rows - 1, msh(idx).FixedCols + tmpItem.���) = tmpItem.ǰ��
                End Select
            Next
            
            '�ϲ�����
            For i = 0 To msh(idx).FixedRows - 2
                msh(idx).MergeRow(i) = True
            Next
            For i = 0 To msh(idx).FixedCols - 1
                msh(idx).MergeCol(i) = True
            Next
            
            '��ͷ���и�
             '�Զ����ͷ����
            On Error Resume Next
            For i = 0 To msh(idx).FixedRows - 1
                Err = 0
                sgnH = Split(Split(tmpItem.��ͷ, "|")(i), "^")(1)
                If Err <> 0 Then sgnH = 250
                msh(idx).RowHeight(i) = sgnH * sgnMode
                msh(idx).Row = i
            Next
            On Error GoTo 0
            Call SetGridLine(.id)
'            Call SetHeadCenter(msh(idx))
        End If
    End With
    msh(idx).Redraw = True
End Sub

Private Sub CustomCellColor(idx As Integer, sCell As Cells, Optional blnClear As Boolean = True)
'���ܣ�����Ԫ���������������ɫ
'������blnClear=�Ƿ����ԭ�����
'      sCell=��䵥Ԫ��Χ,Row=-1�Ǳ�ʾ��������
    Dim i As Integer, j As Integer
    Dim sRow As Integer, sCol As Integer
    
    If sCell.Col1 = -1 Or sCell.Col2 = -1 Or sCell.Row1 = -1 Or sCell.Row2 = -1 Then Exit Sub
    
    msh(idx).Redraw = False
    
    If blnClear Then '������ϴε���ɫ,���������ɫ
        For i = 0 To msh(idx).FixedRows - 1
            msh(idx).Row = i
            For j = 0 To msh(idx).Cols - 1
                msh(idx).Col = j: msh(idx).CellBackColor = msh(idx).BackColorFixed
            Next
        Next
    End If
    
'    '��������ɫ
'    If sCell.Row <> -1 Then
'        msh(idx).Row = sCell.Row
'        For i = 0 To msh(idx).Cols - 1
'            msh(idx).Col = i: msh(idx).CellBackColor = &HC0C0C0
'        Next
'    End If
        
    '��Ԫ��(��Χ)��ɫ
    sRow = 1
    If sCell.Row2 < sCell.Row1 Then sRow = -1
    sCol = 1
    If sCell.Col2 < sCell.Col1 Then sCol = -1
    For i = sCell.Row1 To sCell.Row2 Step sRow
        msh(idx).Row = i
        For j = sCell.Col1 To sCell.Col2 Step sCol
            msh(idx).Col = j: msh(idx).CellBackColor = &HE7CFBA
        Next
    Next
    
    msh(idx).Redraw = True
End Sub

Private Function CheckCell(idx As Integer, Row As Integer, Col As Integer, Text As String) As Boolean
'���ܣ���鵱ǰ��Ԫ��������Ƿ��������,���������������ͷֻ���ڵ������Ϻϲ�
'������Text=��Ҫ�����Row,Col��Ԫ�������
    Dim tmpCell As Cells
    
    '�Ϸ���Ԫ��
    If Row - 1 >= 0 Then
        If Text = msh(idx).TextMatrix(Row - 1, Col) Then
            tmpCell = GetCellRange(msh(idx), Row - 1, Col)
            If Abs(tmpCell.Col1 - tmpCell.Col2) <> 0 Then Exit Function
        End If
    End If
    '�·���Ԫ��
    If Row + 1 <= msh(idx).FixedRows - 1 Then
        If Text = msh(idx).TextMatrix(Row + 1, Col) Then
            tmpCell = GetCellRange(msh(idx), Row + 1, Col)
            If Abs(tmpCell.Col1 - tmpCell.Col2) <> 0 Then Exit Function
        End If
    End If
    '��ߵ�Ԫ��
    If Col - 1 >= 0 Then
        If Text = msh(idx).TextMatrix(Row, Col - 1) Then
            tmpCell = GetCellRange(msh(idx), Row, Col - 1)
            If Abs(tmpCell.Row1 - tmpCell.Row2) <> 0 Then Exit Function
        End If
    End If
    '�ұߵ�Ԫ��
    If Col + 1 <= msh(idx).Cols - 1 Then
        If Text = msh(idx).TextMatrix(Row, Col + 1) Then
            tmpCell = GetCellRange(msh(idx), Row, Col + 1)
            If Abs(tmpCell.Row1 - tmpCell.Row2) <> 0 Then Exit Function
        End If
    End If
    CheckCell = True
End Function

Private Function MergeCell(idx As Integer, sCell As Cells, dCell As Cells) As Cells
'���ܣ��������ͷ,��������Ԫ��ѡ��Χ���кϲ�,������һ���ϲ���ĵ�Ԫ��Χ
'˵����������ܺϲ�,�򷵻�һ����Ч��Χ
    Dim i As Integer
    Dim tmpCell As Cells
    
    MergeCell.Row1 = -1: MergeCell.Col1 = -1: MergeCell.Row2 = -1: MergeCell.Col2 = -1
    
    If sCell.Row1 = sCell.Row2 And sCell.Col1 = sCell.Col2 Then '�ӵ���һ����Ԫ��ʼ�϶�
        MergeCell.Row1 = sCell.Row1
        MergeCell.Col1 = sCell.Col1
        If dCell.Row1 <> dCell.Row2 And dCell.Col1 <> dCell.Col2 Then
            If Abs(dCell.Row1 - dCell.Row2) >= Abs(dCell.Col1 - dCell.Col2) Then
                MergeCell.Row2 = sCell.Row1
                MergeCell.Col2 = dCell.Col2
            Else
                MergeCell.Row2 = dCell.Row2
                MergeCell.Col2 = sCell.Col1
            End If
        Else
            MergeCell.Row2 = dCell.Row2
            MergeCell.Col2 = dCell.Col2
        End If
    Else '���Ѻϲ���Ԫ��ʼ�϶�
        If sCell.Row1 = sCell.Row2 Then 'ͬһ�кϲ�
            If dCell.Row1 = dCell.Row2 Then
                MergeCell.Row1 = sCell.Row1
                MergeCell.Row2 = sCell.Row2
                If dCell.Col2 > dCell.Col1 Then '������
                    MergeCell.Col1 = sCell.Col1
                    MergeCell.Col2 = dCell.Col2
                Else '������
                    MergeCell.Col1 = dCell.Col2
                    MergeCell.Col2 = sCell.Col2
                End If
            End If
        ElseIf sCell.Col1 = sCell.Col2 Then 'ͬһ�кϲ�
            If dCell.Col1 = dCell.Col2 Then
                MergeCell.Col1 = sCell.Col1
                MergeCell.Col2 = sCell.Col2
                If dCell.Row2 > dCell.Row1 Then '������
                    MergeCell.Row1 = sCell.Row1
                    MergeCell.Row2 = dCell.Row2
                Else '������
                    MergeCell.Row1 = dCell.Row2
                    MergeCell.Row2 = sCell.Row2
                End If
            End If
        End If
    End If
    
    MergeCell = AdjustCell(MergeCell)

    '���������ϵ
    If MergeCell.Row1 >= sCell.Row1 And MergeCell.Row2 <= sCell.Row2 And MergeCell.Col1 >= sCell.Col1 And MergeCell.Col2 <= sCell.Col2 Then MergeCell = sCell
    
    '��;���෵����ϲ���Ԫ
    If MergeCell.Col1 <> -1 And MergeCell.Col2 <> -1 And MergeCell.Row1 <> -1 And MergeCell.Row2 <> -1 Then
        If MergeCell.Row1 = MergeCell.Row2 Then
            For i = MergeCell.Col1 To MergeCell.Col2
                If i < sCell.Col1 Or i > sCell.Col2 Then
                    tmpCell = GetCellRange(msh(idx), MergeCell.Row1, i)
                    If tmpCell.Row1 <> tmpCell.Row2 Then MergeCell = sCell: Exit For
                End If
            Next
        End If
        If MergeCell.Col1 = MergeCell.Col2 Then
            For i = MergeCell.Row1 To MergeCell.Row2
                If i < sCell.Row1 Or i > sCell.Row2 Then
                    tmpCell = GetCellRange(msh(idx), i, MergeCell.Col1)
                    If tmpCell.Col1 <> tmpCell.Col2 Then MergeCell = sCell: Exit For
                End If
            Next
        End If
    End If

    '�����ص�Ԫ
    If MergeCell.Row1 <> -1 And MergeCell.Row2 <> -1 And MergeCell.Col1 <> -1 And MergeCell.Col2 <> -1 Then
        tmpCell = GetCellRange(msh(idx), MergeCell.Row1, MergeCell.Col1)
        If tmpCell.Col1 <> tmpCell.Col2 And MergeCell.Row1 = MergeCell.Row2 Then MergeCell.Col1 = tmpCell.Col1
        If tmpCell.Row1 <> tmpCell.Row2 And MergeCell.Col1 = MergeCell.Col2 Then MergeCell.Row1 = tmpCell.Row1
        tmpCell = GetCellRange(msh(idx), MergeCell.Row2, MergeCell.Col2)
        If tmpCell.Col1 <> tmpCell.Col2 And MergeCell.Row1 = MergeCell.Row2 Then MergeCell.Col2 = tmpCell.Col2
        If tmpCell.Row1 <> tmpCell.Row2 And MergeCell.Col1 = MergeCell.Col2 Then MergeCell.Row2 = tmpCell.Row2
    End If
End Function

Private Function AdjustCell(sCell As Cells) As Cells
    Dim i As Integer
    If sCell.Row1 > sCell.Row2 Then
        i = sCell.Row1
        sCell.Row1 = sCell.Row2
        sCell.Row2 = i
    End If
    If sCell.Col1 > sCell.Col2 Then
        i = sCell.Col1
        sCell.Col1 = sCell.Col2
        sCell.Col2 = i
    End If
    AdjustCell = sCell
End Function

Private Sub CustomColColor(idx As Integer, Col As Integer, Optional Clear As Boolean = True)
'���ܣ����������������ɫ
    Dim i As Integer, j As Integer, LngCurRow As Long
    
    If Col = -1 Then Exit Sub
    
    LngCurRow = msh(idx).Row
    msh(idx).Redraw = False
    For i = msh(idx).FixedRows To msh(idx).Rows - 1
        msh(idx).Row = i
        For j = 0 To msh(idx).Cols - 1
            msh(idx).Col = j
            If j <> Col And Clear Then
                msh(idx).CellBackColor = msh(idx).BackColor
            ElseIf j = Col Then
                msh(idx).CellBackColor = &HE7CFBA
            End If
        Next
    Next
    msh(idx).Row = LngCurRow
    If msh(idx).Cols > 0 Then
        msh(idx).Col = IIF(Col = -9, 0, Col)
    End If
    msh(idx).Redraw = True
End Sub

Private Function CheckHead() As String
'���ܣ�����������ͷ�����Ƿ񳬳�
'���أ���ʾ��Ϣ
    Dim tmpItem As RPTItem, tmpID As RelatID
    
    For Each tmpItem In objReport.Items
        If tmpItem.���� = 4 Then
            For Each tmpID In tmpItem.SubIDs
                If LenB(StrConv(objReport.Items("_" & tmpID.id).��ͷ, vbFromUnicode)) > 4000 Then
                    CheckHead = "������ĳ��������ı�ͷ���ֹ�����ձ�ͷ�й���,���飡"
                    Exit Function
                End If
                If CheckText(objReport.Items("_" & tmpID.id), True) = False Then
                    CheckHead = "������ĳ��������ı�ͷ����Դ����ȷ,���飡"
                    Exit Function
                End If
            Next
        End If
    Next
End Function

Private Sub SetMenuDefault(idx As Integer, Col As Integer)
'���ܣ���������,���ݵ�ǰ�ж��뼰���ܷ�ʽ,���ò˵���ѡ��
    Dim tmpID As RelatID, tmpItem As RPTItem
    Dim lngType As Long, i As Integer
    
    mnuCustom_Col_State.Enabled = True
    mnuCustom_Col_Align.Enabled = True
    
    For i = 0 To mnuCustom_Col_Align_Style.UBound
        mnuCustom_Col_Align_Style(i).Checked = False
    Next
    For i = 0 To mnuCustom_Col_State_Style.UBound
        mnuCustom_Col_State_Style(i).Checked = False
    Next
    If Col = -1 Then Exit Sub
    
    For Each tmpID In objReport.Items("_" & intCurID).SubIDs
        Set tmpItem = objReport.Items("_" & tmpID.id)
        If tmpItem.��� = intCurCol Then
            Select Case tmpItem.����
                Case ""
                    mnuCustom_Col_State_Style(0).Checked = True
                Case "SUM"
                    mnuCustom_Col_State_Style(1).Checked = True
                Case "AVG"
                    mnuCustom_Col_State_Style(2).Checked = True
                Case "MAX"
                    mnuCustom_Col_State_Style(3).Checked = True
                Case "MIN"
                    mnuCustom_Col_State_Style(4).Checked = True
                Case "COUNT"
                    mnuCustom_Col_State_Style(5).Checked = True
            End Select
            mnuCustom_Col_Align_Style(tmpItem.����).Checked = True
            
            '���ݵ�һ�ֶ��������ÿ��ò˵�
            If tmpItem.���� Like "*.*" And Left(tmpItem.����, 1) = "[" And Right(tmpItem.����, 1) = "]" Then
                lngType = GetNodeType(Mid(tmpItem.����, 2, Len(tmpItem.����) - 2), tvwSQL)
                If IsType(lngType, adVarChar) Then
                    '�ַ�����ʹ�û���
                    mnuCustom_Col_State.Enabled = False
                End If
                If IsType(lngType, adLongVarBinary) Then
                    'ͼƬ��ʹ�û��ܺͶ���(�̶��ж���)
                    mnuCustom_Col_State.Enabled = False
                    mnuCustom_Col_Align.Enabled = False
                End If
            End If
            
            Exit For
        End If
    Next
End Sub

Private Sub ClassColor(idx As Integer, sCell As Cells, Optional intState As Integer)
'���ܣ����ݵ�ǰ����ֵ�����ܱ�����ɫ
'������sCell=��Row1,Col1��Ч;intState=ͳ������
    Dim i As Integer
    On Error Resume Next
    msh(idx).Redraw = False
    If sCell.Col1 <= msh(idx).FixedCols - 1 And sCell.Row1 >= msh(idx).FixedRows - 1 Then
         '������෶Χ
         msh(idx).Col = sCell.Col1
         For i = msh(idx).FixedRows - 1 To msh(idx).Rows - 1
            msh(idx).Row = i: msh(idx).CellBackColor = &HE7CFBA
         Next
    ElseIf sCell.Row1 <= msh(idx).FixedRows - 2 Then
         '������෶Χ
         msh(idx).Row = sCell.Row1
         For i = 0 To msh(idx).Cols - 1
            msh(idx).Col = i: msh(idx).CellBackColor = &HE7CFBA
         Next
    ElseIf sCell.Col1 >= msh(idx).FixedCols And sCell.Col1 <= msh(idx).FixedCols + intState - 1 And sCell.Row1 >= msh(idx).FixedRows - 1 Then
        'ͳ���Χ
        msh(idx).Col = sCell.Col1
        For i = msh(idx).FixedRows - 1 To msh(idx).Rows - 1
            msh(idx).Row = i: msh(idx).CellBackColor = &HE7CFBA
        Next
    End If
    msh(idx).Redraw = True
End Sub

Private Sub ReFlashWidth()
'���ܣ��������п�(���ݿؼ�����ˢ�¶�������)
    Dim tmpID As RelatID, tmpItem As RPTItem
    For Each tmpItem In objReport.Items
        If tmpItem.��ʽ�� = Mid(cboFormat.ComboItems("_" & mbytCurrFmt).Key, 2) Then  'zyb#Add
            If tmpItem.���� = 4 Then
                For Each tmpID In tmpItem.SubIDs
                    objReport.Items("_" & tmpID.id).W = msh(tmpItem.id).ColWidth(objReport.Items("_" & tmpID.id).���) / sgnMode
                Next
            ElseIf tmpItem.���� = 5 Then
                For Each tmpID In tmpItem.SubIDs
                    If objReport.Items("_" & tmpID.id).���� = 7 Then '��
                        objReport.Items("_" & tmpID.id).W = msh(tmpItem.id).ColWidth(objReport.Items("_" & tmpID.id).���) / sgnMode
                    ElseIf objReport.Items("_" & tmpID.id).���� = 9 Then 'ͳ
                        objReport.Items("_" & tmpID.id).W = msh(tmpItem.id).ColWidth(msh(tmpItem.id).FixedCols + objReport.Items("_" & tmpID.id).���) / sgnMode
                    End If
                Next
            End If
        End If
    Next
End Sub

Private Sub SetDefaultState(strState As String, intAlign As Integer, Optional blnClear As Boolean)
'���ܣ��Ի��ܱ��,���ݵ�ǰ������ܷ�ʽ���ò˵���ʾ
    Dim i As Integer
    For i = 0 To mnuClass_State_Style.UBound
        mnuClass_State_Style(i).Checked = False
    Next
    For i = 0 To mnuClass_Align_Style.UBound
        mnuClass_Align_Style(i).Checked = False
    Next
    If Not blnClear Then
        mnuClass_Align_Style(intAlign).Checked = True
        Select Case strState
            Case ""
                mnuClass_State_Style(0).Checked = True
            Case "SUM"
                mnuClass_State_Style(1).Checked = True
            Case "AVG"
                mnuClass_State_Style(2).Checked = True
            Case "MAX"
                mnuClass_State_Style(3).Checked = True
            Case "MIN"
                mnuClass_State_Style(4).Checked = True
            Case "COUNT"
                mnuClass_State_Style(5).Checked = True
        End Select
    End If
End Sub

Private Sub mnuHelpWebHome_Click()
    zlHomePage hwnd
End Sub

Private Sub mnuHelpWebMail_Click()
    zlMailTo hwnd
End Sub

Private Function UpdatePriv() As Boolean
'���ܣ����±������Ȩ��(���������Ȩ�޸��¼��������Ȩ�޸���)
'���أ��ɹ�=True,ʧ��=False
    Dim rsTmp As ADODB.Recordset, tmpData As RPTData, tmpPar As RPTPar
    Dim strObject As String, strOwner As String, strName As String
    Dim lngProgID As Long, i As Integer, j As Integer, blnTran As Boolean
    Dim strSQL As String, strTmp As String
    
    On Error GoTo errH
    
    '�ǵ�ǰ��¼�����ݿ����Ӻ���
    strObject = ""
    For Each tmpData In objReport.Datas
        If tmpData.���� <> "" And tmpData.�������ӱ�� <= 0 Then
            For j = 0 To UBound(Split(tmpData.����, ","))
                If InStr(strObject & ",", "," & Split(tmpData.����, ",")(j) & ",") = 0 Then
                    strObject = strObject & "," & Split(tmpData.����, ",")(j)
                End If
            Next
        End If
        If tmpData.�������ӱ�� <= 0 Then
            For Each tmpPar In tmpData.Pars
                If tmpPar.���� <> "" Then
                    For i = 0 To UBound(Split(tmpPar.����, "|"))
                        strTmp = Split(tmpPar.����, "|")(i)
                        If strTmp <> "" Then
                            For j = 0 To UBound(Split(strTmp, ","))
                                If InStr(strObject & ",", "," & Split(strTmp, ",")(j) & ",") = 0 Then
                                    strObject = strObject & "," & Split(strTmp, ",")(j)
                                End If
                            Next
                        End If
                    Next
                End If
            Next
        End If
    Next
    strObject = Mid(strObject, 2)
    
    '�Ƿ��ѷ���(���ַ�����ʽ)
    strSQL = _
        " Select " & IIF(objReport.ϵͳ = 0, "-Null", objReport.ϵͳ) & " as ϵͳ,����ID,����" & _
        " From zlReports Where ����ID is Not Null And ID=[1]" & _
        " Union" & _
        " Select " & IIF(objReport.ϵͳ = 0, "-Null", objReport.ϵͳ) & " as ϵͳ,A.����ID,B.����" & _
        " From zlRPTGroups A,zlRPTSubs B" & _
        " Where A.����ID is Not Null And A.ID=B.��ID And B.����ID=[1]" & _
        " Union " & _
        " Select ϵͳ,����ID,���� From zlRPTPuts Where ����ID=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngRPTID)
    If Not rsTmp.EOF Then
        On Error GoTo errH
        gcnOracle.BeginTrans: blnTran = True
        Do While Not rsTmp.EOF
            '������дȨ��
            If IsNull(rsTmp!����) Then
                gcnOracle.Execute "Delete From zlProgPrivs Where ���=" & rsTmp!����id & _
                                  " And ���� is Null And Nvl(ϵͳ,0)=" & Nvl(rsTmp!ϵͳ, 0)
            Else
                gcnOracle.Execute "Delete From zlProgPrivs Where ���=" & rsTmp!����id & _
                                  " And ����='" & rsTmp!���� & "' And Nvl(ϵͳ,0)=" & Nvl(rsTmp!ϵͳ, 0)
            End If
            If strObject <> "" Then '�ñ���п��ܲ��������ݿ�
                For i = 0 To UBound(Split(strObject, ","))
                    strOwner = Left(Split(strObject, ",")(i), InStr(Split(strObject, ",")(i), ".") - 1)
                    If strOwner <> "SYS" And strOwner <> "ZLTOOLS" And strOwner <> "SYSTEM" Then
                        strName = Mid(Split(strObject, ",")(i), InStr(Split(strObject, ",")(i), ".") + 1)
                        gcnOracle.Execute GetInsertProgPrivs(Nvl(rsTmp!ϵͳ, 0), Nvl(rsTmp!����id, 0) _
                                                , rsTmp!����, strName, strOwner, "SELECT")
                    End If
                Next
            End If
            rsTmp.MoveNext
        Loop
        gcnOracle.CommitTrans: blnTran = False
    End If
    UpdatePriv = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    If blnTran Then gcnOracle.RollbackTrans
    Call SaveErrLog
End Function

Private Sub ReFlashReportBySelFormat()
'zyb#Add
'���ܣ�����ָ�������ʽ,����ˢ����ʾ��������
    Dim objTmp As Object
    
    For Each objTmp In lblSize
        If objTmp.Index <> 0 Then Unload objTmp
    Next
    For Each objTmp In lblLine
        If objTmp.Index <> 0 Then Unload objTmp
    Next
    For Each objTmp In lbl
        If objTmp.Index <> 0 Then Unload objTmp
    Next
    For Each objTmp In msh
        If objTmp.Index <> 0 Then Unload objTmp
    Next
    For Each objTmp In img
        If objTmp.Index <> 0 Then Unload objTmp
    Next
    For Each objTmp In ImgCode
        If objTmp.Index <> 0 Then Unload objTmp
    Next
    For Each objTmp In Chart
        If objTmp.Index <> 0 Then Unload objTmp
    Next
    For Each objTmp In pic
        If objTmp.Index <> 0 Then Unload objTmp
    Next
    For Each objTmp In Shp
        If objTmp.Index <> 0 Then Unload lblshp(objTmp.Index): Unload objTmp
    Next
    
    intCurID = 0
    Set objLastSel = Nothing
    
    Call ShowItems
    Call ShowAttrib
    Call AdjustAll
    
    Me.Refresh
End Sub

Private Function GetRPtFmtName() As String
    Dim IntFmt As Integer, StrFmtName As String, CboItem As ComboItem, StrCompare As String
    'zyb#Add
    '����ȱʡ������ʽ������(���ظ�)
    
    StrFmtName = ""
    For IntFmt = 1 To cboFormat.ComboItems.count
        Set CboItem = cboFormat.ComboItems(IntFmt)
        StrCompare = CStr(Val(StrReverse(Val(StrReverse(CboItem.Text & "9"))))) '����һ������,�������һλ��Ķ�ʧ
        StrCompare = Mid(StrCompare, 1, Len(StrCompare) - 1)
        
        If Len(StrCompare) > Len(StrFmtName) Then
            StrFmtName = StrCompare
        ElseIf Len(StrCompare) = Len(StrFmtName) Then
            If StrFmtName < StrCompare Then StrFmtName = StrCompare
        End If
    Next
    
    '��ȡ�õ����ּ�1
    If StrFmtName = "" Then
        GetRPtFmtName = objReport.���� & "1"
    Else
        GetRPtFmtName = objReport.���� & (CLng(StrFmtName) + 1)
    End If
End Function

Private Sub DrawFrame(ByVal ObjOper As Label)
    'zyb#Add
    '���߿���
    
'    With WinProperty
'        .L = ObjOper.Left - 10
'        .T = ObjOper.Top - 10
'        .H = ObjOper.Top + ObjOper.Height - 10
'        .W = ObjOper.Left + ObjOper.Width - 10
'
'        picPaper.Line (.L, .T)-(.W, .H), objReport.Items("_" & ObjOper.Index).ǰ��, B
'    End With
End Sub

Private Function GetAutoTest(ByVal bytMode As Byte) As Single
    Dim sgnCompare As Single, sgnActure As Single
    Dim objFmt As RPTFmt
    
    '��ȡ��Ӧ�ı���
    Set objFmt = objReport.Fmts("_" & mbytCurrFmt)
    Select Case bytMode
    Case 0  '����Ӧ
        sgnCompare = IIF(objFmt.H > objFmt.W, objFmt.H, objFmt.W)
        sgnActure = IIF(picPaper.Height > picPaper.Width, scrHsc.Top - picRulerH.Top - picRulerH.Height, scrVsc.Left - picRulerV.Left - picRulerV.Width)
    Case 1  '��Ӧ���
        sgnCompare = IIF(objFmt.ֽ�� = 1, objFmt.W, objFmt.H)
        sgnActure = scrVsc.Left - picRulerV.Left - picRulerV.Width
    Case 2  '��Ӧ�߶�
        sgnCompare = IIF(objFmt.ֽ�� = 1, objFmt.H, objFmt.W)
        sgnActure = scrHsc.Top - picRulerH.Top - picRulerH.Height
    End Select
    GetAutoTest = (sgnActure - 200) / sgnCompare
End Function

Private Function CheckNameValid(ByVal strName As String) As Boolean
    Dim ItemThis As RPTItem
    '������ƵĺϷ���
    
    CheckNameValid = False
    For Each ItemThis In objReport.Items
        If ItemThis.��ʽ�� = mbytCurrFmt Then
            If ItemThis.Key <> intCurID And strName = ItemThis.���� Then Exit Function
        End If
    Next
    CheckNameValid = True
End Function

Private Sub ReferTo(Optional ByVal ItemTest As RPTItem)
    Dim ItemThis As RPTItem
    Dim ObjSel As Control, BytWay As Byte, bytKind As Byte
    Dim TargetObj As VSFlexGrid, ReferToObjname As String
    Dim DblAdd As Double

    '���ñ�ǩ��λ��,����ն���,�������ʵĲ�ͬ�������仯
    If intCurID = 0 Then Exit Sub
    If Val(objReport.Items("_" & intCurID).����) = 0 Then Exit Sub
    
    Select Case objReport.Items("_" & intCurID).����
        Case 2
            Set ObjSel = lbl(intCurID)
            If objReport.Items("_" & intCurID).��ID = 0 Then
                Set ObjSel.Container = picPaper
            Else
                Set ObjSel.Container = pic(objReport.Items("_" & intCurID).��ID)
            End If
            ReferToObjname = objReport.Items("_" & intCurID).����   '����
            
            For Each ItemThis In objReport.Items
                If ItemThis.���� = ReferToObjname And ItemThis.��ʽ�� = mbytCurrFmt Then Set TargetObj = msh(ItemThis.Key): Exit For
            Next
        
            '�ƶ��ؼ�
            Select Case Mid(objReport.Items("_" & intCurID).����, 1, 1)
            Case 1
                If Not (ObjSel.Top + ObjSel.Height + 100 * sgnMode < TargetObj.Top) Then
                    ObjSel.Top = TargetObj.Top - 100 * sgnMode - ObjSel.Height
                End If
            Case 2
                If Not (ObjSel.Top >= TargetObj.Top + GetTableHeight(TargetObj) + 50 * sgnMode And ObjSel.Top <= picPaper.Height - 200 * sgnMode) Then
                    ObjSel.Top = TargetObj.Top + GetTableHeight(TargetObj) + 100 * sgnMode
                End If
            End Select
            Select Case Mid(objReport.Items("_" & intCurID).����, 2)
            Case 1
                ObjSel.Left = TargetObj.Left
            Case 2
                ObjSel.Left = GetTableWidth(TargetObj) / 2 + TargetObj.Left - (ObjSel.Width / 2)
            Case 3
                ObjSel.Left = GetTableWidth(TargetObj) + TargetObj.Left - ObjSel.Width
            End Select
            objReport.Items("_" & ObjSel.Index).X = ObjSel.Left / sgnMode
            objReport.Items("_" & ObjSel.Index).Y = ObjSel.Top / sgnMode
            Call AdjustSelCons(ObjSel)
        Case 4, 5
            '��ȡ����
            For Each ItemThis In objReport.Items
                If ItemThis.���� = IIF(ReferToObjname = "", objReport.Items("_" & intCurID).����, ReferToObjname) And ItemThis.��ʽ�� = mbytCurrFmt Then Set TargetObj = msh(ItemThis.Key): Exit For
            Next
            
            If objReport.Items("_" & intCurID).���� <> "" And ItemTest.���� = "" Then
                BytWay = objReport.Items("_" & intCurID).����
                ReferToObjname = objReport.Items("_" & intCurID).����
                
                '�������������
                DblAdd = IIF(BytWay = 1, TargetObj.Top, TargetObj.Left) + IIF(BytWay = 1, TargetObj.Height, TargetObj.Width) - 15
                For Each ItemThis In objReport.Items
                    If ItemThis.���� = ReferToObjname And ItemThis.��ʽ�� = mbytCurrFmt And ItemThis.Key <> intCurID And InStr(1, "4,5", ItemThis.����) <> 0 Then
                        Set ObjSel = msh(ItemThis.Key)
                        DblAdd = DblAdd + IIF(BytWay = 1, ObjSel.Height, ObjSel.Width) - 15 * sgnMode
                    End If
                Next
                
                Set ObjSel = msh(intCurID)
                If objReport.Items("_" & intCurID).��ID = 0 Then
                    Set ObjSel.Container = picPaper
                Else
                    Set ObjSel.Container = pic(objReport.Items("_" & intCurID).��ID)
                End If
                ObjSel.Left = IIF(BytWay = 1, TargetObj.Left, DblAdd)
                ObjSel.Top = IIF(BytWay = 1, DblAdd, TargetObj.Top)
                If BytWay = 2 Then
                    ObjSel.Height = TargetObj.Height
                Else
                    ObjSel.Width = TargetObj.Width
                End If
                objReport.Items("_" & ObjSel.Index).Y = ObjSel.Top / sgnMode
                objReport.Items("_" & ObjSel.Index).X = ObjSel.Left / sgnMode
                objReport.Items("_" & ObjSel.Index).W = ObjSel.Width / sgnMode
                objReport.Items("_" & ObjSel.Index).H = ObjSel.Height / sgnMode
            Else
                BytWay = ItemTest.����
                ReferToObjname = ItemTest.����
                
                For Each ItemThis In objReport.Items
                    If ItemThis.���� = ReferToObjname And ItemThis.��ʽ�� = mbytCurrFmt And InStr(1, "4,5", ItemThis.����) <> 0 Then
                        Set ObjSel = msh(ItemThis.Key)
                        Select Case BytWay
                        Case 1
                            If ObjSel.Top > msh(intCurID).Top Then
                                ObjSel.Top = ObjSel.Top - msh(intCurID).Height
                                ItemThis.Y = ObjSel.Top / sgnMode
                            End If
                        Case 2
                            If ObjSel.Left > msh(intCurID).Left Then
                                ObjSel.Left = ObjSel.Left - msh(intCurID).Width
                                ItemThis.X = ObjSel.Left / sgnMode
                            End If
                        End Select
                    End If
                Next
                Set ObjSel = msh(intCurID)
                If objReport.Items("_" & intCurID).��ID = 0 Then
                    Set ObjSel.Container = picPaper
                Else
                    Set ObjSel.Container = pic(objReport.Items("_" & intCurID).��ID)
                End If
            
                BytWay = objReport.Items("_" & intCurID).����
                ReferToObjname = objReport.Items("_" & intCurID).����
                
                If ReferToObjname <> "" Then
                    '�������������
                    DblAdd = IIF(BytWay = 1, TargetObj.Top, TargetObj.Left) + IIF(BytWay = 1, TargetObj.Height, TargetObj.Width) - 15
                    For Each ItemThis In objReport.Items
                        If ItemThis.���� = ReferToObjname And ItemThis.��ʽ�� = mbytCurrFmt And ItemThis.Key <> intCurID And InStr(1, "4,5", ItemThis.����) <> 0 Then
                            Set ObjSel = msh(ItemThis.Key)
                            DblAdd = DblAdd + IIF(BytWay = 1, ObjSel.Height, ObjSel.Width) - 15 * sgnMode
                        End If
                    Next
                    
                    Set ObjSel = msh(intCurID)
                    ObjSel.Left = IIF(BytWay = 1, TargetObj.Left, DblAdd)
                    ObjSel.Top = IIF(BytWay = 1, DblAdd, TargetObj.Top)
                    ObjSel.Height = TargetObj.Height
                    ObjSel.Width = TargetObj.Width
                    objReport.Items("_" & ObjSel.Index).Y = ObjSel.Top / sgnMode
                    objReport.Items("_" & ObjSel.Index).X = ObjSel.Left / sgnMode
                    objReport.Items("_" & ObjSel.Index).W = ObjSel.Width / sgnMode
                    objReport.Items("_" & ObjSel.Index).H = ObjSel.Height / sgnMode
                End If
            End If
            SetGridLine (ObjSel.Index)
            Call AdjustSelCons(ObjSel)
            ObjSel.ZOrder 0
    End Select
End Sub

Private Sub AdjustSelCons(ByVal ObjSel As Object)
    '����ѡ��ؼ�LblSize��λ��
    Dim i As Integer, lngTop As Long, lngLeft As Long
    With WinProperty
        .H = lblSize(0).Height / Screen.TwipsPerPixelX
        .W = lblSize(0).Width / Screen.TwipsPerPixelX
    End With
    If UCase(ObjSel.Container.name) = "PIC" Then
        lngTop = ObjSel.Container.Top
        lngLeft = ObjSel.Container.Left
    End If

    If Mid(ObjSel.Tag, 3) = "" Then Exit Sub
    For i = Mid(ObjSel.Tag, 3) To Mid(ObjSel.Tag, 3) + 7 'ѡ���־��"����"��ʼ,"˳ʱ��"����
        Select Case IIF(i Mod 8 <> 0, i Mod 8, 8) '��λѡ��߿��λ��
            Case 1 '����
                With WinProperty
                    .T = (ObjSel.Top + lngTop - lblSize(i).Height) / Screen.TwipsPerPixelX
                    .l = (ObjSel.Left + lngLeft + (ObjSel.Width - lblSize(i).Width) / 2) / Screen.TwipsPerPixelX
                    Call MoveWindow(lblSize(i).hwnd, .l, .T, .W, .H, 1)
                End With
                lblSize(i).MousePointer = 7
            Case 2 '����
                With WinProperty
                    .T = (ObjSel.Top + lngTop - lblSize(i).Height) / Screen.TwipsPerPixelX
                    .l = (ObjSel.Left + lngLeft + ObjSel.Width) / Screen.TwipsPerPixelX
                    Call MoveWindow(lblSize(i).hwnd, .l, .T, .W, .H, 1)
                End With
                lblSize(i).MousePointer = 6
            Case 3 '����
                With WinProperty
                    .T = (ObjSel.Top + lngTop + (ObjSel.Height - lblSize(i).Height) / 2) / Screen.TwipsPerPixelX
                    .l = (ObjSel.Left + lngLeft + ObjSel.Width) / Screen.TwipsPerPixelX
                    Call MoveWindow(lblSize(i).hwnd, .l, .T, .W, .H, 1)
                End With
                lblSize(i).MousePointer = 9
            Case 4 '����
                With WinProperty
                    .T = (ObjSel.Top + lngTop + ObjSel.Height) / Screen.TwipsPerPixelX
                    .l = (ObjSel.Left + lngLeft + ObjSel.Width) / Screen.TwipsPerPixelX
                    Call MoveWindow(lblSize(i).hwnd, .l, .T, .W, .H, 1)
                End With
                lblSize(i).MousePointer = 8
            Case 5 '����
                With WinProperty
                    .T = (ObjSel.Top + lngTop + ObjSel.Height) / Screen.TwipsPerPixelX
                    .l = (ObjSel.Left + lngLeft + (ObjSel.Width - lblSize(i).Width) / 2) / Screen.TwipsPerPixelX
                    Call MoveWindow(lblSize(i).hwnd, .l, .T, .W, .H, 1)
                End With
                lblSize(i).MousePointer = 7
            Case 6 '����
                With WinProperty
                    .T = (ObjSel.Top + lngTop + ObjSel.Height) / Screen.TwipsPerPixelX
                    .l = (ObjSel.Left + lngLeft - lblSize(i).Width) / Screen.TwipsPerPixelX
                    Call MoveWindow(lblSize(i).hwnd, .l, .T, .W, .H, 1)
                End With
                lblSize(i).MousePointer = 6
            Case 7 '����
                With WinProperty
                    .T = (ObjSel.Top + lngTop + (ObjSel.Height - lblSize(i).Height) / 2) / Screen.TwipsPerPixelX
                    .l = (ObjSel.Left + lngLeft - lblSize(i).Width) / Screen.TwipsPerPixelX
                    Call MoveWindow(lblSize(i).hwnd, .l, .T, .W, .H, 1)
                End With
                lblSize(i).MousePointer = 9
            Case 8 '����
                With WinProperty
                    .T = (ObjSel.Top + lngTop - lblSize(i).Height) / Screen.TwipsPerPixelX
                    .l = (ObjSel.Left + lngLeft - lblSize(i).Width) / Screen.TwipsPerPixelX
                    Call MoveWindow(lblSize(i).hwnd, .l, .T, .W, .H, 1)
                End With
                lblSize(i).MousePointer = 8
        End Select
        lblSize(i).ZOrder
        lblSize(i).Visible = True
    Next
End Sub

Private Sub AdjustAll(Optional ByVal BlnSelected As Boolean = False)
    Dim ItemThis As RPTItem, intTmpIdx As Integer
    
    '����������װ���,�������пؼ�
    intTmpIdx = intCurID
    For Each ItemThis In objReport.Items
        If ItemThis.��ʽ�� = mbytCurrFmt Then
            Select Case ItemThis.����
                Case 2
                    intCurID = ItemThis.Key
                    Call ReferTo(ItemThis)
            End Select
        End If
    Next
    intCurID = intTmpIdx
    If GetSelNum = 1 And intCurID <> 0 Then
        Call ShowAttrib(intCurID)
    End If
End Sub

Private Function TestReferTo(ByVal LngIdx As Long, ByVal lngX As Long, ByVal lngY As Long) As Boolean
    Dim ItemThis As RPTItem
    Dim ObjSel As Label, BytWay As Byte, bytKind As Byte
    Dim TargetObj As VSFlexGrid, ReferToObjname As String
    '���ñ�ǩ��λ��,����ն���,�������ʵĲ�ͬ�������仯
    TestReferTo = False
    If LngIdx = 0 Then Exit Function
    If objReport.Items("_" & LngIdx).���� = "" Then Exit Function
    If Val(objReport.Items("_" & LngIdx).����) = 0 Then Exit Function
    
    Set ObjSel = lbl(LngIdx)
    ReferToObjname = objReport.Items("_" & LngIdx).����
    bytKind = objReport.Items("_" & LngIdx).����   '����
    
    For Each ItemThis In objReport.Items
        If ItemThis.���� = ReferToObjname And ItemThis.��ʽ�� = mbytCurrFmt Then Set TargetObj = msh(ItemThis.Key): Exit For
    Next
    
    '�ƶ��ؼ�
    Select Case Mid(bytKind, 1, 1)
    Case 1, 2
        If Mid(bytKind, 1, 1) = "1" Then
            If ObjSel.Top + ObjSel.Height + lngY + 100 * sgnMode >= TargetObj.Top Then TestReferTo = True
        Else
            If Not (ObjSel.Top + lngY >= TargetObj.Top + GetTableHeight(TargetObj) + 50 And ObjSel.Top + lngY <= picPaper.Height - 200) Then TestReferTo = True
        End If
        If TestReferTo = False Then
            Select Case Mid(bytKind, 2)
            Case 1
                If Not (ObjSel.Left + lngX >= TargetObj.Left And ObjSel.Left + lngX <= TargetObj.Left + 200 * sgnMode) Then TestReferTo = True
            Case 2
                If Not (ObjSel.Left + lngX >= GetTableWidth(TargetObj) / 2 + TargetObj.Left - (ObjSel.Width / 2) - 100 * sgnMode And ObjSel.Left + lngX <= GetTableWidth(TargetObj) / 2 + TargetObj.Left - (ObjSel.Width / 2) + 100 * sgnMode) Then TestReferTo = True
            Case 3
                If Not (ObjSel.Left + lngX >= GetTableWidth(TargetObj) + TargetObj.Left - ObjSel.Width - 100 * sgnMode And ObjSel.Left + lngX <= GetTableWidth(TargetObj) + TargetObj.Left - ObjSel.Width + 100 * sgnMode) Then TestReferTo = True
            End Select
        End If
    End Select
End Function

Private Function CheckTableProperty(ByVal ItemThis As RPTItem) As Boolean
    Dim ItemTest As RPTItem, IntTest As Integer, StrTest As String, StrThis As String
    '���ñ���Ƿ����������ӻ򸽼�
    CheckTableProperty = False
    
    'Keyֵ��ͬ���˳�
    Set ItemTest = objReport.Items("_" & intCurID)
    If ItemTest.Key = ItemThis.Key Then Exit Function
    If ItemThis.���� <> "" Then Exit Function
    
    Select Case ItemTest.����
    Case 4  '���ӱ��(ֻ���������)
        '��鸽�ӱ���Ƿ����(����Ϊ��)
        If InStr(1, "4,5", ItemThis.����) = 0 Then Exit Function
        If Not (ItemThis.���� < 2) Then Exit Function
        
    Case 5  '�����ӱ��(ֻ���Ƿ�����)
        If ItemTest.SubIDs.count = 0 Or ItemThis.SubIDs.count = 0 Then Exit Function
        If ItemThis.���� = 4 Then Exit Function
        
        '��鱻�����ӵı���������ӱ�����������Ƿ�һ��
        StrTest = "": StrThis = ""
        For IntTest = 1 To ItemTest.SubIDs.count
            If objReport.Items("_" & ItemTest.SubIDs(IntTest).Key).���� = 7 Then
                StrTest = StrTest & "|" & objReport.Items("_" & ItemTest.SubIDs(IntTest).Key).����
            End If
        Next
        For IntTest = 1 To ItemThis.SubIDs.count
            If objReport.Items("_" & ItemThis.SubIDs(IntTest).Key).���� = 7 Then
                StrThis = StrThis & "|" & objReport.Items("_" & ItemThis.SubIDs(IntTest).Key).����
            End If
        Next
        If StrTest <> StrThis Then Exit Function
    End Select
    
    CheckTableProperty = True
End Function

Private Sub LinkMove(ByVal LngIdx As Long, ByVal lngX As Long, ByVal lngY As Long)
    Dim ItemThis As RPTItem, ObjSel As Control
    '�����ƶ����ӱ�������ӱ��
    
    For Each ItemThis In objReport.Items
        If ItemThis.���� = objReport.Items("_" & LngIdx).���� _
            And ItemThis.��ʽ�� = mbytCurrFmt And ItemThis.���� <> "" Then
            Select Case ItemThis.����
                Case 2
                    Set ObjSel = lbl(ItemThis.Key)
                Case 4, 5
                    Set ObjSel = msh(ItemThis.Key)
                Case 12 '@@@
                    Set ObjSel = Chart(ItemThis.Key)
            End Select
            
            ObjSel.Left = ObjSel.Left + lngX
            ObjSel.Top = ObjSel.Top + lngY
            ItemThis.X = ObjSel.Left / sgnMode
            ItemThis.Y = ObjSel.Top / sgnMode
            ObjSel.ZOrder
            
            Call AdjustSelCons(ObjSel)
    
            '����Ϊѡ��״̬
            On Error Resume Next '���ܴ����������,��ʵ���ƶ������ӱ�
            lblSize(Mid(msh(LngIdx).Tag, 3) + 2).ZOrder
            lblSize(Mid(msh(LngIdx).Tag, 3) + 4).ZOrder
        End If
    Next
End Sub

Private Sub SetMainWH(ByVal LngIdx As Long)
    Dim ItemThis As RPTItem, SelObj As Control
    '�����ӱ�,�����������ӱ�һ��(��������߶�һ��,��������һ��)
    If InStr(1, "4,5", objReport.Items("_" & LngIdx).����) = 0 Then Exit Sub
    If objReport.Items("_" & LngIdx).���� = "" Then Call SetChildWH(LngIdx): Exit Sub
    
    '�����ӱ�,����Ӧ��������(��Ҫָ�����߶�)
    For Each ItemThis In objReport.Items
        If ItemThis.��ʽ�� = mbytCurrFmt And InStr(1, "4,5", ItemThis.����) <> 0 And ItemThis.���� = objReport.Items("_" & LngIdx).���� Then
            Set SelObj = msh(ItemThis.Key)
            If objReport.Items("_" & LngIdx).���� = 1 Then
                SelObj.Width = msh(LngIdx).Width
                ItemThis.W = SelObj.Width / sgnMode
            Else
                SelObj.Height = msh(LngIdx).Height
                ItemThis.H = SelObj.Height / sgnMode
            End If
            Call SetGridLine(SelObj.Index)
            Call SetChildWH(ItemThis.Key)
            Exit Sub
        End If
    Next
End Sub

Private Sub SetChildWH(ByVal LngIdx As Long)
    Dim TargetObj As Control, SelObj As Control, ItemThis As RPTItem, Int���� As Integer
    Dim OrderTable() As Long, StrParentName As String, ArrayCount As Integer, ArrayIn As Integer, ArrayOut As Integer
    '������������ӱ�Ŀ����߶�
    
    If objReport.Items("_" & LngIdx).���� <> "" Then Exit Sub
    Set TargetObj = msh(LngIdx)
    StrParentName = objReport.Items("_" & LngIdx).����
    ArrayCount = 0
    '��ȡ�����ӱ�
    For Each ItemThis In objReport.Items
        If ItemThis.��ʽ�� = mbytCurrFmt And ItemThis.���� = StrParentName And InStr(1, "4,5", ItemThis.����) <> 0 Then
            ReDim Preserve OrderTable(ArrayCount)
            OrderTable(ArrayCount) = ItemThis.Key
            Int���� = ItemThis.���� '�����ӱ������һ��
            ArrayCount = ArrayCount + 1
        End If
    Next
    
    'ʹ��������
    Dim lngTmp As Long
    For ArrayOut = 0 To ArrayCount - 1
        For ArrayIn = 0 To ArrayCount - 2
            Select Case Int����
            Case 1
                If objReport.Items("_" & OrderTable(ArrayIn + 1)).Y < objReport.Items("_" & OrderTable(ArrayIn)).Y Then
                    lngTmp = OrderTable(ArrayIn)
                    OrderTable(ArrayIn) = OrderTable(ArrayIn + 1)
                    OrderTable(ArrayIn + 1) = lngTmp
                End If
            Case 2
                If objReport.Items("_" & OrderTable(ArrayIn + 1)).X < objReport.Items("_" & OrderTable(ArrayIn)).X Then
                    lngTmp = OrderTable(ArrayIn)
                    OrderTable(ArrayIn) = OrderTable(ArrayIn + 1)
                    OrderTable(ArrayIn + 1) = lngTmp
                End If
            End Select
        Next
    Next
    
    '���������ӱ�
    lngTmp = IIF(Int���� = 1, TargetObj.Top + TargetObj.Height, TargetObj.Left + TargetObj.Width)
    For ArrayOut = 0 To ArrayCount - 1
        Set SelObj = msh(OrderTable(ArrayOut))
        SelObj.Top = IIF(Int���� = 1, lngTmp, TargetObj.Top)
        SelObj.Left = IIF(Int���� = 1, TargetObj.Left, lngTmp)
        If Int���� = 1 Then '���һ��
            SelObj.Width = TargetObj.Width
        Else
            SelObj.Height = TargetObj.Height
        End If
        Set ItemThis = objReport.Items("_" & OrderTable(ArrayOut))
        ItemThis.X = SelObj.Left / sgnMode
        ItemThis.Y = SelObj.Top / sgnMode
        ItemThis.W = SelObj.Width / sgnMode
        ItemThis.H = SelObj.Height / sgnMode
        Call SetGridLine(SelObj.Index)
        lngTmp = lngTmp + IIF(Int���� = 1, SelObj.Height, SelObj.Width)
    Next
    
    '����Ϊѡ��״̬
    Set SelObj = msh(LngIdx)
    Call AdjustSelCons(SelObj)
End Sub

Private Sub ChangeReferTo(ByVal strOldName As String, ByVal strNewName As String)
    Dim ItemThis  As RPTItem
    '�޸����е��ӱ�Ĳ��ն���
    
    For Each ItemThis In objReport.Items
        If ItemThis.���� = strOldName And ItemThis.��ʽ�� = mbytCurrFmt Then ItemThis.���� = strNewName
    Next
End Sub

Private Function GetTableWidth(ByVal TargetObj As Control) As Single
    Dim ItemThis As RPTItem, SgnWidth As Single, strName As String
    '����ָ������Ŀ��(���ӱ�Ŀ��)
    
    SgnWidth = TargetObj.Width
    If objReport.Items("_" & TargetObj.Index).���� > 1 Then SgnWidth = SgnWidth * objReport.Items("_" & TargetObj.Index).����
    strName = objReport.Items("_" & TargetObj.Index).����
    For Each ItemThis In objReport.Items
        If ItemThis.���� = strName And ItemThis.��ʽ�� = mbytCurrFmt And InStr(1, "4,5", ItemThis.����) <> 0 Then
            If ItemThis.���� = 2 Then SgnWidth = SgnWidth + msh(ItemThis.Key).Width
        End If
    Next
    GetTableWidth = SgnWidth
End Function

Private Function GetTableHeight(ByVal TargetObj As Control) As Single
    Dim ItemThis As RPTItem, SgnHeight As Single, strName As String
    '����ָ������ĸ߶�(���ӱ�ĸ߶�)
    
    SgnHeight = TargetObj.Height
    strName = objReport.Items("_" & TargetObj.Index).����
    For Each ItemThis In objReport.Items
        If ItemThis.���� = strName And ItemThis.��ʽ�� = mbytCurrFmt And InStr(1, "4,5", ItemThis.����) <> 0 Then
            If ItemThis.���� = 1 Then SgnHeight = SgnHeight + msh(ItemThis.Key).Height
        End If
    Next
    GetTableHeight = SgnHeight
End Function

Private Sub SetRowHeight(ByVal ObjSel As VSFlexGrid)
    Dim lngRow As Long, LngRows As Long
    Dim ItemThis As RPTItem, ArrayHeight
    '���ù̶��е��и�
    
    ArrayHeight = Split(objReport.Items("_" & objReport.Items("_" & ObjSel.Index).SubIDs(1).Key).��ͷ, "|")
    LngRows = UBound(ArrayHeight)
    For lngRow = 0 To LngRows
        ObjSel.RowHeight(lngRow) = Split(ArrayHeight, "^")(1) * sgnMode
    Next
End Sub

Private Function GetNextName(ByVal intType As Integer, Optional ByVal BlnTestClip As Boolean = False) As String
    Dim intMax As Integer, ItemThis As RPTItem, strName As String
    '����һ�����ظ�������
    
    intMax = 0
    For Each ItemThis In objReport.Items
        If ItemThis.��ʽ�� = mbytCurrFmt And ItemThis.���� = intType Then
            strName = Val(StrReverse(Val(StrReverse(ItemThis.���� & "9"))))
            strName = Mid(strName, 1, Len(strName) - 1)
            If intMax < CInt(Val(strName)) Then intMax = CInt(Val(strName))
        End If
    Next
    If BlnTestClip Then
        For Each ItemThis In objClip
            If ItemThis.���� = intType Then
                strName = Val(StrReverse(Val(StrReverse(ItemThis.���� & "9"))))
                strName = Mid(strName, 1, Len(strName) - 1)
                If intMax < CInt(Val(strName)) Then intMax = CInt(Val(strName))
            End If
        Next
    End If
    intMax = intMax + 1
    
    Select Case intType
    Case 1
        GetNextName = "����"
    Case 2, 3
        GetNextName = "��ǩ"
    Case 4
        GetNextName = "�����"
    Case 5
        GetNextName = "���ܱ�"
    Case 10
        GetNextName = "����"
    Case 11
        GetNextName = "ͼƬ"
    Case 12 '@@@
        GetNextName = "ͼ��"
    Case 13
        GetNextName = "����"
    Case 14
        GetNextName = "��Ƭ"
    End Select
    GetNextName = GetNextName & intMax
End Function

Private Function UserRefer(ByVal LngIdx As Long) As Boolean
    Dim ItemThis As RPTItem
    '������������Ĳ��ն���������ɾ��
    UserRefer = True
    
    For Each ItemThis In objReport.Items
        If ItemThis.��ʽ�� = mbytCurrFmt And ItemThis.Key <> LngIdx And ItemThis.���� = objReport.Items("_" & LngIdx).���� Then Exit Function
    Next
    
    UserRefer = False
End Function

Private Sub GetAllElement()
    Dim ItemThis As RPTItem
    
    '��ȡ��ǰ�����ʽ�е����б���Ԫ��
    cboAtt.Clear
    cboAtt.AddItem ""
    
    For Each ItemThis In objReport.Items
        If ItemThis.��ʽ�� = mbytCurrFmt And InStr(1, "|1,|2,|3,|4,|5,|10,|11,|12,|14,", "|" & ItemThis.���� & ",") <> 0 Then
            cboAtt.AddItem ItemThis.����
            cboAtt.ItemData(cboAtt.NewIndex) = ItemThis.Key
        End If
    Next
    cboAtt.ListIndex = 0
End Sub

Private Sub LoadOutChart()
    With cboAtt
        .Clear
        .AddItem "��ֹ���"
        .ItemData(.NewIndex) = 0
        .AddItem "���ͼ"
        .ItemData(.NewIndex) = 1        'xlArea
        .AddItem "����ͼ"
        .ItemData(.NewIndex) = 4        'xlLine
        .AddItem "��ͼ"
        .ItemData(.NewIndex) = 5        'xlPie
        .AddItem "����ͼ"
        .ItemData(.NewIndex) = 15       'xlBubble
        .AddItem "����ͼ"
        .ItemData(.NewIndex) = 51       'xlColumnClustered
        .AddItem "����ͼ"
        .ItemData(.NewIndex) = 57       'xlBarClustered
        .AddItem "����ͼ"
        .ItemData(.NewIndex) = 83       'xlSurface
        .AddItem "�ɼ�ͼ"
        .ItemData(.NewIndex) = 88       'xlStockHLC
        .AddItem "Բ��ͼ"
        .ItemData(.NewIndex) = 92       'xlCylinderColClustered
        .AddItem "Բ׶ͼ"
        .ItemData(.NewIndex) = 99       'xlConeColClustered
        .AddItem "��׶ͼ"
        .ItemData(.NewIndex) = 106      'xlPyramidColClustered
        .AddItem "ɢ��ͼ"
        .ItemData(.NewIndex) = -4169    'xlXYScatter
        .AddItem "Բ��ͼ"
        .ItemData(.NewIndex) = -4120    'xlDoughnut
        .AddItem "�״�ͼ"
        .ItemData(.NewIndex) = -4151    'xlRadar
        .AddItem "��ά����ͼ"
        .ItemData(.NewIndex) = -4100    'xl3DColumn
        .ListIndex = 0
    End With
End Sub

Private Function GetCurOutChart() As String
    Select Case objReport.Fmts("_" & mbytCurrFmt).ͼ��
        Case 0
            GetCurOutChart = "��ֹ���"
        Case 1        'xlArea
            GetCurOutChart = "���ͼ"
        Case 4        'xlLine
            GetCurOutChart = "����ͼ"
        Case 5        'xlPie
            GetCurOutChart = "��ͼ"
        Case 15       'xlBubble
            GetCurOutChart = "����ͼ"
        Case 51       'xlColumnClustered
            GetCurOutChart = "����ͼ"
        Case 57       'xlBarClustered
            GetCurOutChart = "����ͼ"
        Case 83       'xlSurface
            GetCurOutChart = "����ͼ"
        Case 88       'xlStockHLC
            GetCurOutChart = "�ɼ�ͼ"
        Case 92       'xlCylinderColClustered
            GetCurOutChart = "Բ��ͼ"
        Case 99       'xlConeColClustered
            GetCurOutChart = "Բ׶ͼ"
        Case 106      'xlPyramidColClustered
            GetCurOutChart = "��׶ͼ"
        Case -4169    'xlXYScatter
            GetCurOutChart = "ɢ��ͼ"
        Case -4120    'xlDoughnut
            GetCurOutChart = "Բ��ͼ"
        Case -4151    'xlRadar
            GetCurOutChart = "�״�ͼ"
        Case -4100    'xl3DColumn
            GetCurOutChart = "��ά����ͼ"
    End Select
End Function

Private Sub LocateOutChart()
    Dim i As Integer, IntOutChart As Integer
    '��ʾ����ʽ����Ӧ�����ģʽ
    
    blnModify = False
    IntOutChart = objReport.Fmts("_" & mbytCurrFmt).ͼ��
    For i = 0 To cboAtt.ListCount - 1
        If cboAtt.ItemData(i) = IntOutChart Then
            cboAtt.ListIndex = i
            Exit For
        End If
    Next
    If cboAtt.ListIndex < 0 Then cboAtt.ListIndex = 0
    If cboAtt.ItemData(cboAtt.ListIndex) <> IntOutChart Then cboAtt.ListIndex = 0
    blnModify = True
End Sub

Private Function CheckText(ByVal tmpItem As RPTItem, Optional ByVal blnHead As Boolean = False) As Boolean
    Dim objNode As Node, arrData As Variant
    Dim strFind As String, intFind As Integer, blnFind As Boolean
    
    CheckText = False
    strFind = GetLabelDataName(IIF(blnHead, tmpItem.��ͷ, tmpItem.����))
    arrData = Split(strFind, "|")
    
    For intFind = 0 To UBound(arrData)
        'ÿ������Դ����ƥ�䣬�����˳�
        blnFind = False
        For Each objNode In tvwSQL.Nodes
            If objNode.Key <> "Root" And objNode.Children = 0 Then
                If InStr(1, arrData(intFind), LevelText(objNode)) <> 0 Then blnFind = True: Exit For
            End If
        Next
        If blnFind = False Then Exit Function
    Next
    CheckText = True
End Function

Private Sub GetInPaper()
    Dim i As Integer, j As Integer, k As Integer
    Dim IntPaper As Integer, strTmp As String
    Dim strPaperBinName As String * 1000, strPaperBin As String * 100
    '--------------------------------------------------------------------------------------------
    
    If Printers.count = 0 Then
        MsgBox "��ϵͳ��û�м�⵽�κδ�ӡ�豸,�뾡�����,���򲿷ֲ�����������ִ�С�", vbInformation, App.Title
        Exit Sub
    End If
    
    IntPaper = DeviceCapabilities(Printer.DeviceName, Printer.Port, DC_BINNAMES, strPaperBinName, 0)
    IntPaper = DeviceCapabilities(Printer.DeviceName, Printer.Port, DC_BINS, strPaperBin, 0)
    
    CboTest.Clear
    j = 1
    For i = 1 To IntPaper
        k = 0
        '��ֽ����
        Do
            If Mid(strPaperBinName, j, 1) = Chr(0) Then
                If Trim(strTmp) <> "" Then
                    CboTest.AddItem Trim(strTmp)
                    
                    '��ֽ���
                    CboTest.ItemData(CboTest.ListCount - 1) = Asc(Mid(strPaperBin, i * 2, 1)) * 256# + Asc(Mid(strPaperBin, i * 2 - 1, 1))
                    If CboTest.ItemData(CboTest.ListCount - 1) = objReport.��ֽ Then
                        CboTest.ListIndex = CboTest.ListCount - 1 '��λ��ԭ������
                    End If
                    If CboTest.ListIndex = -1 And CboTest.ItemData(CboTest.ListCount - 1) = Printer.PaperBin Then
                        CboTest.ListIndex = CboTest.ListCount - 1 '��λ�ڴ�ӡ��ȱʡ������
                    End If
                End If
                
                j = 24 + j - LenB(StrConv(strTmp, vbFromUnicode))
                strTmp = ""
                Exit Do
            Else
                strTmp = strTmp & Mid(strPaperBinName, j, 1)
                j = j + 1
                k = k + 1
                If k > 24 Then Exit Do
            End If
        Loop
    Next
    '--------------------------------------------------------------------------------------------
    If CboTest.ListIndex = -1 And CboTest.ListCount > 0 Then CboTest.ListIndex = 0
End Sub

Private Function ReferObj(ByVal LngIdx As Long) As Boolean
    Dim ItemThis As RPTItem, StrMainObj As String
    '��⵱ǰԪ���Ƿ���������ӱ�
    
    ReferObj = False
    StrMainObj = objReport.Items("_" & LngIdx).����
    If objReport.Items("_" & LngIdx).���� <> "" Then ReferObj = True: Exit Function
    For Each ItemThis In objReport.Items
        If ItemThis.��ʽ�� = mbytCurrFmt And ItemThis.���� = StrMainObj And ItemThis.���� = 5 Then
            ReferObj = True
            Exit Function
        End If
    Next
End Function

Private Function CheckCoordinate() As Long
    Dim ItemCheck As RPTItem, blnCheck As Boolean
    Dim RectCheck As RECT
    '�������ϵͳ������Ŀ���������ס����������
    
    CheckCoordinate = 0
    For Each ItemCheck In objReport.Items
        If ItemCheck.ϵͳ Then
            RectCheck.Left = ItemCheck.X
            RectCheck.Top = ItemCheck.Y
            RectCheck.Right = ItemCheck.W
            RectCheck.Bottom = ItemCheck.H
            blnCheck = GetCoordinate(RectCheck, ItemCheck.Key, True)
            If blnCheck Then CheckCoordinate = ItemCheck.Key: Exit Function
        End If
    Next
End Function

Private Function GetCoordinate(ByRef Area As RECT, Optional ByVal IntStyle As Integer = 1, _
Optional ByVal lngKey As Long, Optional ByVal blnCheck As Boolean = False) As Boolean
'���ܣ��Զ���λ����ճ��ϵͳ������Ŀʱ��
'���أ�ѡ�еĸ���
    Dim tmpItem As RPTItem, ObjSel As Object, LngLoop As Long
    Dim ObjLeft As Single, ObjTop As Single, ObjHeight As Single, ObjWidth As Single
    
    For LngLoop = 1 To objReport.Items.count
        Set tmpItem = objReport.Items(LngLoop)
        If InStr(1, "|1,|2,|3,|4,|5,|10,|11,|12,|14,", "|" & tmpItem.����) <> 0 _
            And Mid(cboFormat.ComboItems("_" & mbytCurrFmt).Key, 2) = tmpItem.��ʽ�� And tmpItem.Key <> lngKey Then
            Set ObjSel = GetInxObj(tmpItem.id)
            
            ObjTop = objReport.Items("_" & tmpItem.id).Y
            ObjLeft = objReport.Items("_" & tmpItem.id).X
            ObjWidth = objReport.Items("_" & tmpItem.id).W
            ObjHeight = objReport.Items("_" & tmpItem.id).H
            
            If Not (ObjTop > Area.Bottom + Area.Top Or _
                ObjLeft > Area.Right + Area.Left Or _
                ObjTop + ObjHeight < Area.Top Or _
                ObjLeft + ObjWidth < Area.Left) Then
                If blnCheck Then GetCoordinate = True: Exit Function
                If IntStyle = 1 Then
                    Area.Top = ObjTop + ObjHeight + 100
                Else
                    Area.Top = ObjTop - Area.Bottom - 100
                End If
                LngLoop = 0
            End If
        End If
    Next
End Function

Private Function AdjustCoordinate(Optional ByVal BlnIn As Boolean)
'blnIn=��������ʱ����
    Dim ObjSel As Object, RectTest As RECT
    
    If objReport.Items("_" & intCurID).���� <> 0 Or BlnIn Then
        RectTest.Left = objReport.Items("_" & intCurID).X
        RectTest.Top = objReport.Items("_" & intCurID).Y
        RectTest.Bottom = objReport.Items("_" & intCurID).H
        RectTest.Right = objReport.Items("_" & intCurID).W
        If BlnIn = False Then
            Call GetCoordinate(RectTest, IIF(Mid(objReport.Items("_" & intCurID).����, 1, 1) = 1, 2, 1), intCurID)
        
            objReport.Items("_" & intCurID).X = RectTest.Left
            objReport.Items("_" & intCurID).Y = RectTest.Top
        End If
        Set ObjSel = GetInxObj(intCurID)
        
        With ObjSel
            .Left = RectTest.Left * sgnMode
            .Top = RectTest.Top * sgnMode
            If objReport.Items("_" & intCurID).��ID = 0 Then
                Set ObjSel.Container = picPaper
            Else
                Set ObjSel.Container = pic(objReport.Items("_" & intCurID).��ID)
            End If
        End With
        
        Call AdjustSelCons(ObjSel)
    End If
End Function

Private Function CheckClip()
    Dim ItemCheck As RPTItem, StrMainObj As String, ArrayRefer
    Dim i As Integer, blnFind As Boolean
    '���ճ���ؼ���������,����������ӿؼ��Ĳ��ն�������
    
    StrMainObj = ","
    For Each ItemCheck In objClip
        If ItemCheck.���� <> "" And InStr(1, StrMainObj, ItemCheck.����) = 0 Then
            StrMainObj = StrMainObj & IIF(StrMainObj = ",", "", ",") & ItemCheck.����
        End If
    Next
    For Each ItemCheck In objClip
        If ItemCheck.���� = 2 Then ItemCheck.���� = "": ItemCheck.���� = 0
    Next
    
    StrMainObj = Mid(StrMainObj, 2)
    ArrayRefer = Split(StrMainObj, ",")
    
    For i = 0 To UBound(ArrayRefer)
        blnFind = False
        For Each ItemCheck In objClip
            If ItemCheck.���� = "" And ArrayRefer(i) = ItemCheck.���� Then
                blnFind = True
                Exit For
            End If
        Next
        If blnFind = False Then
            '��������ӱ�Ĳ��ն�������
            For Each ItemCheck In objClip
                If ItemCheck.���� = ArrayRefer(i) Then ItemCheck.���� = "": ItemCheck.���� = 0
            Next
        End If
    Next
End Function

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

Private Function GetDependID(strName As String) As Integer
'���ܣ����ݲ�������,��ȡ������.
    Dim objItem As RPTItem
    
    For Each objItem In objReport.Items
        If objItem.��ʽ�� = mbytCurrFmt And objItem.���� = strName _
            And (objItem.���� = 4 Or objItem.���� = 5) And objItem.���� = 0 Then
            GetDependID = objItem.id: Exit Function
        End If
    Next
End Function

