VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0BE3824E-5AFE-4B11-A6BC-4B3AD564982A}#8.0#0"; "olch2x8.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReport 
   AutoRedraw      =   -1  'True
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9480
   Icon            =   "frmReport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6510
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picLR_S 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5010
      Left            =   4140
      MousePointer    =   9  'Size W E
      ScaleHeight     =   5010
      ScaleWidth      =   45
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1140
      Width           =   45
   End
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   6150
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmReport.frx":014A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11298
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   18
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
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
   Begin VB.VScrollBar scrVsc 
      DragIcon        =   "frmReport.frx":09DE
      Height          =   5175
      LargeChange     =   20
      Left            =   9225
      Max             =   100
      SmallChange     =   10
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   720
      Width           =   250
   End
   Begin VB.HScrollBar scrHsc 
      DragIcon        =   "frmReport.frx":0CE8
      Height          =   250
      LargeChange     =   20
      Left            =   4185
      Max             =   100
      SmallChange     =   10
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5895
      Width           =   4995
   End
   Begin VB.PictureBox picBack 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      Height          =   5145
      Left            =   4230
      ScaleHeight     =   5085
      ScaleWidth      =   4950
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   735
      Width           =   5010
      Begin VB.PictureBox picRotate 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   4440
         ScaleHeight     =   285
         ScaleWidth      =   330
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   1800
         Visible         =   0   'False
         Width           =   330
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfRelations 
         Height          =   1215
         Left            =   720
         TabIndex        =   35
         Top             =   3120
         Visible         =   0   'False
         Width           =   2055
         _cx             =   1964641833
         _cy             =   1964640351
         Appearance      =   1
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
         BackColor       =   14737632
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   14737632
         GridColor       =   16761024
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   30
         Cols            =   0
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
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
         FillStyle       =   1
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
      Begin VSFlex8Ctl.VSFlexGrid msh 
         Height          =   1575
         Index           =   0
         Left            =   240
         TabIndex        =   31
         Top             =   960
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
         MouseIcon       =   "frmReport.frx":0FF2
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
         HighLight       =   2
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
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
         ScrollBars      =   3
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
      Begin VB.PictureBox picTemp 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   765
         Left            =   3960
         ScaleHeight     =   765
         ScaleWidth      =   330
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1815
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.PictureBox picPaper 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3390
         Index           =   0
         Left            =   255
         ScaleHeight     =   3390
         ScaleWidth      =   3315
         TabIndex        =   6
         Top             =   165
         Width           =   3315
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   -8888
            ScaleHeight     =   225
            ScaleWidth      =   345
            TabIndex        =   30
            Top             =   0
            Width           =   375
         End
      End
      Begin VB.PictureBox picShadow 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3390
         Left            =   330
         ScaleHeight     =   3390
         ScaleWidth      =   3315
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   255
         Width           =   3315
      End
   End
   Begin VB.PictureBox picGroup 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5010
      Left            =   0
      ScaleHeight     =   5010
      ScaleWidth      =   4140
      TabIndex        =   9
      TabStop         =   0   'False
      Tag             =   "Save"
      Top             =   1140
      Width           =   4140
      Begin VB.PictureBox picPar 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00F4F4F4&
         Height          =   3090
         Left            =   45
         ScaleHeight     =   3030
         ScaleWidth      =   4050
         TabIndex        =   15
         TabStop         =   0   'False
         Tag             =   "Save"
         Top             =   2325
         Width           =   4110
         Begin VB.CommandButton cmdSelAll 
            Caption         =   "ȫѡ"
            Height          =   350
            Left            =   120
            TabIndex        =   34
            Top             =   930
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.CommandButton cmdSelNone 
            Cancel          =   -1  'True
            Caption         =   "ȫ��"
            Height          =   350
            Left            =   765
            TabIndex        =   33
            Top             =   930
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.CommandButton cmdLoad 
            BackColor       =   &H00F4F4F4&
            Caption         =   "ȷ��(&O)"
            Height          =   350
            Left            =   1680
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   930
            Visible         =   0   'False
            Width           =   1100
         End
         Begin VB.CommandButton cmdDefault 
            BackColor       =   &H00F4F4F4&
            Caption         =   "����(&D)"
            Height          =   350
            Left            =   2850
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   930
            Visible         =   0   'False
            Width           =   1100
         End
         Begin VB.Frame fraGroup 
            BackColor       =   &H00F4F4F4&
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   0
            Left            =   210
            TabIndex        =   16
            Top             =   -60
            Visible         =   0   'False
            Width           =   3825
         End
         Begin VB.Frame fra 
            BackColor       =   &H00F4F4F4&
            ForeColor       =   &H00800000&
            Height          =   645
            Index           =   0
            Left            =   210
            TabIndex        =   17
            Top             =   60
            Visible         =   0   'False
            Width           =   3825
            Begin VB.OptionButton opt 
               BackColor       =   &H00F4F4F4&
               Caption         =   "#"
               Height          =   180
               Index           =   0
               Left            =   105
               MaskColor       =   &H8000000F&
               TabIndex        =   18
               Top             =   270
               Visible         =   0   'False
               Width           =   1150
            End
         End
         Begin VB.CheckBox chk 
            BackColor       =   &H00F4F4F4&
            Caption         =   "#"
            Height          =   195
            Index           =   0
            Left            =   1455
            TabIndex        =   23
            Top             =   255
            Visible         =   0   'False
            Width           =   2775
         End
         Begin VB.ComboBox cbo 
            BackColor       =   &H00F4F4F4&
            Height          =   300
            Index           =   0
            Left            =   1455
            TabIndex        =   21
            Top             =   195
            Visible         =   0   'False
            Width           =   2460
         End
         Begin VB.TextBox txt 
            BackColor       =   &H00F4F4F4&
            Height          =   300
            Index           =   0
            Left            =   1455
            TabIndex        =   20
            Top             =   195
            Visible         =   0   'False
            Width           =   2475
         End
         Begin VB.CommandButton cmd 
            BackColor       =   &H00F4F4F4&
            Caption         =   "��"
            Height          =   240
            Index           =   0
            Left            =   4425
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "�� F2 ��ѡ����"
            Top             =   225
            Visible         =   0   'False
            Width           =   270
         End
         Begin MSComCtl2.DTPicker dtp 
            Height          =   300
            Index           =   0
            Left            =   1455
            TabIndex        =   22
            Top             =   195
            Visible         =   0   'False
            Width           =   2460
            _ExtentX        =   4339
            _ExtentY        =   529
            _Version        =   393216
            CalendarBackColor=   16053492
            CalendarTitleBackColor=   12946264
            CalendarTitleForeColor=   16053492
            CustomFormat    =   "yyyy��MM��dd�� HH:mm:ss"
            Format          =   43778051
            CurrentDate     =   36731
         End
         Begin VB.Frame fraSplit 
            BackColor       =   &H00F4F4F4&
            Height          =   75
            Left            =   -180
            TabIndex        =   25
            Top             =   750
            Visible         =   0   'False
            Width           =   10000
         End
         Begin VB.Label lblName 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��������"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   0
            Left            =   675
            TabIndex        =   24
            Top             =   255
            Visible         =   0   'False
            Width           =   720
         End
      End
      Begin MSComctlLib.ListView lvw 
         Height          =   1845
         Left            =   45
         TabIndex        =   11
         Tag             =   "Save"
         Top             =   225
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   3254
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "img32"
         SmallIcons      =   "img16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "����"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "���"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "˵��"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label lblPar_S 
         BackColor       =   &H009B6737&
         Caption         =   " ��������"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   45
         MousePointer    =   7  'Size N S
         TabIndex        =   13
         Top             =   2100
         Width           =   4080
      End
      Begin VB.Label lblGroup_S 
         BackColor       =   &H009B6737&
         Caption         =   " ������"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   45
         TabIndex        =   12
         Top             =   15
         Width           =   4095
      End
   End
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   1140
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   2011
      _CBWidth        =   9480
      _CBHeight       =   1140
      _Version        =   "6.7.9782"
      Child1          =   "tbr"
      MinWidth1       =   4500
      MinHeight1      =   720
      Width1          =   810
      NewRow1         =   0   'False
      Caption2        =   "��ʽ"
      Child2          =   "cboFormat"
      MinWidth2       =   2505
      MinHeight2      =   330
      Width2          =   4005
      NewRow2         =   0   'False
      Caption3        =   "����"
      Child3          =   "txtFind"
      MinWidth3       =   1005
      MinHeight3      =   330
      Width3          =   1935
      NewRow3         =   0   'False
      Begin VB.TextBox txtFind 
         Height          =   330
         Left            =   585
         TabIndex        =   32
         Top             =   780
         Width           =   8805
      End
      Begin MSComctlLib.ImageCombo cboFormat 
         Height          =   315
         Left            =   6000
         TabIndex        =   8
         Top             =   225
         Width           =   3390
         _ExtentX        =   5980
         _ExtentY        =   556
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   16053492
         Locked          =   -1  'True
         ImageList       =   "img16"
      End
      Begin MSComctlLib.Toolbar tbr 
         Height          =   720
         Left            =   165
         TabIndex        =   2
         Top             =   30
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imgGray"
         HotImageList    =   "imgColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   18
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ԥ��"
               Key             =   "Preview"
               Description     =   "Ԥ��"
               Object.ToolTipText     =   "��ӡԤ��"
               Object.Tag             =   "Ԥ��"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "Print"
               Description     =   "��ӡ"
               Object.ToolTipText     =   "��ӡ"
               Object.Tag             =   "��ӡ"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ͼ��"
               Key             =   "Graph"
               Description     =   "ͼ��"
               Object.ToolTipText     =   "�Ե�ǰ������ͼ�η���"
               Object.Tag             =   "ͼ��"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Par"
               Description     =   "����"
               Object.ToolTipText     =   "��������"
               Object.Tag             =   "����"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Par_"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�п�"
               Key             =   "ColWidth"
               Description     =   "�п�"
               Object.ToolTipText     =   "�п�"
               Object.Tag             =   "�п�"
               ImageIndex      =   4
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   3
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Auto"
                     Object.Tag             =   "����ƥ��"
                     Text            =   "����ƥ��"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Fill"
                     Object.Tag             =   "������"
                     Text            =   "������"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Def"
                     Object.Tag             =   "ȱʡ����"
                     Text            =   "ȱʡ����"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ѡ��"
               Key             =   "SelMode"
               Description     =   "ѡ��"
               Object.ToolTipText     =   "�������ѡ��ģʽ"
               Object.Tag             =   "ѡ��"
               ImageKey        =   "SelMode"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "RowMode"
                     Object.Tag             =   "����ѡ��"
                     Text            =   "����ѡ��"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "ColMode"
                     Object.Tag             =   "����ѡ��"
                     Text            =   "����ѡ��"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�б�"
               Key             =   "Style"
               Object.ToolTipText     =   "�������б���ʾ��ʽ"
               Object.Tag             =   "�б�"
               ImageKey        =   "Style"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Large"
                     Object.Tag             =   "��ͼ��"
                     Text            =   "��ͼ��"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Small"
                     Object.Tag             =   "Сͼ��"
                     Text            =   "Сͼ��"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "List"
                     Object.Tag             =   "�б�"
                     Text            =   "�б�"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Detail"
                     Object.Tag             =   "��ϸ����"
                     Text            =   "��ϸ����"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Style_"
               Style           =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ǰ��"
               Key             =   "Pre"
               Description     =   "ǰ��"
               Object.ToolTipText     =   "�л���ǰһ�ű���(Page Up)"
               Object.Tag             =   "ǰ��"
               ImageKey        =   "Pre"
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Next"
               Description     =   "����"
               Object.ToolTipText     =   "�л�����һ�ű���(Page Down)"
               Object.Tag             =   "����"
               ImageKey        =   "Next"
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Page_"
               Style           =   3
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Description     =   "����"
               Object.ToolTipText     =   "��ǰ��������"
               Object.Tag             =   "����"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Quit"
               Description     =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageIndex      =   6
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   705
      Top             =   465
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":18CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":1AE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":1D00
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":1F1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":2134
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":234E
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":2568
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":2782
            Key             =   "Style"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":299C
            Key             =   "Pre"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":2BB6
            Key             =   "Next"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":2DD0
            Key             =   "SelMode"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   75
      Top             =   465
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":2FEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":3204
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":341E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":3638
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":3852
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":3A6C
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":3C86
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":3EA0
            Key             =   "Style"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":40BA
            Key             =   "Pre"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":42D4
            Key             =   "Next"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":44EE
            Key             =   "SelMode"
         EndProperty
      EndProperty
   End
   Begin VB.Timer timHead 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin MSScriptControlCtl.ScriptControl Srt 
      Left            =   6855
      Top             =   75
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin MSComctlLib.ImageList img32 
      Left            =   2745
      Top             =   1125
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":4708
            Key             =   "Report"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   2100
      Top             =   1125
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
            Picture         =   "frmReport.frx":4A22
            Key             =   "Format"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":4B7C
            Key             =   "Report"
         EndProperty
      EndProperty
   End
   Begin C1Chart2D8.Chart2D Chart 
      Height          =   1230
      Index           =   0
      Left            =   4275
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   4470
      Visible         =   0   'False
      Width           =   1650
      _Version        =   524288
      _Revision       =   7
      _ExtentX        =   2910
      _ExtentY        =   2170
      _StockProps     =   0
      ControlProperties=   "frmReport.frx":4CD6
   End
   Begin VB.Image imgCode 
      Appearance      =   0  'Flat
      Height          =   330
      Index           =   0
      Left            =   4230
      Stretch         =   -1  'True
      Top             =   2415
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image img 
      Appearance      =   0  'Flat
      Height          =   330
      Index           =   0
      Left            =   4230
      Stretch         =   -1  'True
      Top             =   2415
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Line lin 
      Index           =   0
      Visible         =   0   'False
      X1              =   4380
      X2              =   5655
      Y1              =   1905
      Y2              =   1905
   End
   Begin VB.Shape Shp 
      FillColor       =   &H80000005&
      Height          =   315
      Index           =   0
      Left            =   4365
      Top             =   1995
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "#"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   4380
      MouseIcon       =   "frmReport.frx":5335
      TabIndex        =   14
      Top             =   1560
      Visible         =   0   'False
      Width           =   930
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFile_Setup 
         Caption         =   "��ӡ����(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFile_Preview 
         Caption         =   "��ӡԤ��(&V)"
      End
      Begin VB.Menu mnuFile_Print 
         Caption         =   "��ӡ����(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFile_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_Excel 
         Caption         =   "Excel�������(&E)"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuFile_Graph 
         Caption         =   "Excelͼ�η���(&G)"
         Shortcut        =   ^G
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
      Begin VB.Menu mnuEdit_Par 
         Caption         =   "��������(&R)"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuEdit_Par_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_SetCol 
         Caption         =   "�����п�(&C)"
         Begin VB.Menu mnuEdit_SetCol_Auto 
            Caption         =   "����ƥ��(&A)"
         End
         Begin VB.Menu mnuEdit_SetCol_Fill 
            Caption         =   "������(&I)"
         End
         Begin VB.Menu mnuEdit_SetCol_Def 
            Caption         =   "ȱʡ����(&D)"
         End
      End
      Begin VB.Menu mnuEdit_SelMode 
         Caption         =   "ѡ��ģʽ(&S)"
         Begin VB.Menu mnuEdit_SelMode_Row 
            Caption         =   "����ѡ��(&R)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuEdit_SelMode_Col 
            Caption         =   "����ѡ��(&C)"
         End
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "��ͼ(&V)"
      Begin VB.Menu mnuViewTool 
         Caption         =   "������(&T)"
         Begin VB.Menu mnuViewToolButton 
            Caption         =   "��׼��ť(&B)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewToolFormat 
            Caption         =   "�����ʽ(&F)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewToolGroup 
            Caption         =   "������(&G)"
            Checked         =   -1  'True
            Shortcut        =   {F11}
         End
         Begin VB.Menu mnuView_Tlb_1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewToolText 
            Caption         =   "�ı���ǩ(&L)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "״̬��(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuEdit_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewStyle 
         Caption         =   "��ͼ��(&G)"
         Index           =   0
      End
      Begin VB.Menu mnuViewStyle 
         Caption         =   "Сͼ��(&M)"
         Index           =   1
      End
      Begin VB.Menu mnuViewStyle 
         Caption         =   "�б�(&L)"
         Checked         =   -1  'True
         Index           =   2
      End
      Begin VB.Menu mnuViewStyle 
         Caption         =   "��ϸ����(&D)"
         Index           =   3
      End
      Begin VB.Menu mnuViewStyle_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuView_Pre 
         Caption         =   "ǰһ��(&P)"
      End
      Begin VB.Menu mnuView_Next 
         Caption         =   "��һ��(&N)"
      End
      Begin VB.Menu mnuView_Page_ 
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
         Caption         =   "WEB�ϵ�����"
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
   Begin VB.Menu mnuPop 
      Caption         =   "�����˵�"
      Visible         =   0   'False
      Begin VB.Menu mnuPop_Cond 
         Caption         =   "����1"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPop_Split1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPop_Save 
         Caption         =   "����(&S)"
      End
      Begin VB.Menu mnuPop_SaveAs 
         Caption         =   "���Ϊ(&A)"
      End
      Begin VB.Menu mnuPop_Del 
         Caption         =   "ɾ��(&C)"
      End
      Begin VB.Menu mnuPop_Split2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPop_Default 
         Caption         =   "ȱʡ(&D)"
      End
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mobjCurDLL As clsReport '�򿪵�ǰ����ʱ����ʼ����һ������������(DLL:clsReport,��Ҫ���������¼�)
Public mbytStyle As Byte    '�����������ʾ����ʽ

Public mblnDisabledPrint As Boolean     'ֻԤ���������ӡ
Public mblnPrintEmpty As Boolean '�Ƿ��ӡ�ޱ�����ݵĸ�ʽ
Public bytFormat As Byte '��ǰ���������򿪵ĸ�ʽ��
Public marrPars As Variant '������������,���ڱ��ű������

Public frmParent As Object '����ˢ��
Public mobjReport As Report '��ǰ�������(�����������еĵ�ǰ)

Private arrReport() As Report '��������,��ű������еĶ��ű���
Private arrLibDatas() As LibDatas '�������и������������Դ����(ע�⣺ռ��Դ,��Ϊ�л���ʽʱ�����¶�ȡ,����Ҫÿ������)
Private arrDefPars() As RPTPars '�������и��������ȱʡ��������

Public intReport As Integer '>=0,������ʱ��ǰ�������,��ӦpicPaper������

'���ڲ�����(���ڴ�ӡ��Ԥ����Excel��ͼ�η���)-------------------------
Public mLibDatas As LibDatas '���������еĶ������Դ����,�򿪱���ʱ����
Public marrPage As Variant   '����PageCells���ϵ�����,��ӡ��Ԫ��������,����Ԥ�����ӡ
Public marrPageCard As Variant   '��;��Ƭ��ҳ��
Public mcolRowIDs As New Collection '�������ڼ�¼���������ID(ID��Դ������Դ�ֶ�,��һ�����ڱ����)

'ģ�����------------------------------------------------------------
Private mstrExcelFile As String
Private mblnAllFormat As Boolean
Private lngPreX As Long, lngPreY As Long
Private intGridCount As Integer '��ǰ���������еĶ��������(һ������������)
Private intGridID As Integer '���ֻ��һ���������,��Ϊ��ؼ�ID
Private objCurGrid As Object
Private mobjPars As RPTPars '�������д����������ʱʹ��
Private mobjDefPars As RPTPars '��ŵ�ǰ����ԭʼ�Ĳ�������,���ڻָ�ȱʡֵ
Private objScript As clsScript
Private blnMatch As Boolean, blnExcel As Boolean
Private blnRefresh As Boolean
Private lngCurInx As Long
Private lngTmpColor As Long
Private mstrPDFFile As String
Private mlngReportID As Long
Private mlngRPTID As Long               '����ӱ���ID���������ID
Private mblnLeftClick As Boolean
Private mlngRelationReport As Long
Private mintGridIndex As Integer
Private mintLblIndex As Integer
Private mlngRelationMouseRow As Long
Private mlngRelationMouseCol As Long
Private mlngBackX As Long
Private mlngBackY As Long
Private mbytType As Byte
Private mintCurMenuIndex As Integer
Private mintCurCondID As Integer
Private mobjfrmShow As frmPreview
Private mobjfrmShowDock As frmPreviewDock
Private mlngSys As Long

Private Const CON_SETFOCES As Long = &H9C6D75

Public Sub ShowMe(objParent As Object, objCurDLL As clsReport, arrPars As Variant, ByVal bytStyle As Byte)
    Set frmParent = objParent
    Set mobjCurDLL = objCurDLL
    marrPars = arrPars
    mbytStyle = bytStyle
    mlngSys = glngSys
    
    On Error Resume Next
    
    If mbytStyle <> 0 Then
        Load Me
        If Err.Number = 0 Then
            If mbytStyle = 1 Then       '�Զ�Ԥ��
                mnuFile_Preview_Click
            ElseIf mbytStyle = 2 Then   '�Զ���ӡ
                mnuFile_Print_Click
            ElseIf mbytStyle = 3 Then   '�����Excel
                mnuFile_Excel_Click
            ElseIf mbytStyle = 4 Then   '�̶������PDF
                mnuFile_Print_Click
            End If
        ElseIf Err.Number <> 0 Then
            '364:������ж��(��Form_Load�ڲ�Unload,��ȡ����������)
            Err.Clear
        End If
        Unload Me
    Else
        '�ȳ����Է�ģ̬��ʾ����
        If frmParent Is Nothing Then
            Me.Show
        ElseIf frmParent.name = "frmDesign" Then
            Me.Show 1, frmParent
        Else
            Me.Show , frmParent
        End If
        
        '������������Է�ģ̬��ʾ
        If Err.Number = 373 Or Err.Number = 401 Then
            '373:��֧�ֱ������ƻ����������ڲ�����(Դ�������zlReport.dll,��֧�ּӸ�����)
            '401:������ģʽ����ʱ������ʾ��ģʽ����
            '���Զ�Load������ʾʱ�����ټ���Form_Load�¼�
            Err.Clear: Me.Show 1
        ElseIf Err.Number = 364 Then
            '364:������ж��(��Form_Load�ڲ�Unload,��ȡ����������)
            Err.Clear
        ElseIf Err.Number <> 0 Then
            Err.Clear: Unload Me '���Զ�Load��δ֪����ʱж�ش���
        End If
    End If
End Sub

Private Sub CopyLibDatas(objS As LibDatas, objO As LibDatas)
'���ܣ�������ͬ����֮��Ķ������Դ
    Dim tmpData As LibData
    
    Set objO = New LibDatas
    
    For Each tmpData In objS
        objO.Add tmpData.DataName, tmpData.DataSet.Clone, "_" & tmpData.DataName
    Next
End Sub

Private Sub CboFormat_Click()
    Dim strErr As String
    Dim strStartTime As String
    
    If CByte(Mid(cboFormat.SelectedItem.Key, 2)) = bytFormat Then Exit Sub
    bytFormat = CByte(Mid(cboFormat.SelectedItem.Key, 2))
    mobjReport.bytFormat = bytFormat
    
    mnuFile_Graph.Enabled = (mobjReport.Fmts("_" & bytFormat).ͼ�� <> 0)
    tbr.Buttons("Graph").Enabled = (mobjReport.Fmts("_" & bytFormat).ͼ�� <> 0)
    
    If mobjReport.blnLoad Then
        If gblnReportRunLog Then
            strStartTime = Format(Currentdate, "YYYY-MM-DD HH:mm:SS")
        End If
        '��ȡ��ǰ��ʽ��Ҫ������(���Ѿ���ȡ��������ʽ����Դ�������)
        strErr = OpenReportData(False)
        If strErr <> "" Then
            MsgBox "�ڶ�ȡ��������""" & strErr & """ʱ�����������,�����ܲ�����", vbInformation, App.Title
            Exit Sub
        End If

        '���������ύ�¼�
        If Not mobjCurDLL Is Nothing Then
            mobjCurDLL.Act_CommitCondition mobjReport.���, GetParsStr(MakeNamePars(mobjReport, True)), Me
        End If
        
        Call ShowItems
        If Val(cboFormat.Tag) = 0 Then
            If mlngReportID > 0 Then
                Call RecordsExecute(mlngReportID, strStartTime, 2)
            ElseIf lvw.Visible Then
                Call RecordsExecute(Val(Mid(lvw.SelectedItem.Key, 2)), strStartTime, 2)
            End If
        End If
    End If
End Sub

Private Sub CboFormat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Set cboFormat.SelectedItem = cboFormat.ComboItems("_" & bytFormat)
        KeyAscii = 0
    End If
End Sub

Private Sub cbr_Resize()
    Form_Resize
End Sub

Private Sub Chart_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sngX As Single, sngY As Single
    Dim lngSeries As Long, lngPoint As Long, lngDS As Long
    Dim strSeries As String, vArea As RegionConstants
    Dim dblX As Double, strX As String, strY As String
    Dim strLabelX As String, strLabelY As String
    
    With Chart(Index).ChartGroups(1)
        sngX = X / Screen.TwipsPerPixelX
        sngY = Y / Screen.TwipsPerPixelY
        vArea = .CoordToDataIndex(sngX, sngY, oc2dFocusXY, lngSeries, lngPoint, lngDS)
        If vArea = oc2dRegionInChartArea Then
            If lngDS <= 3 Then
                strSeries = ""
                If lngSeries <= .SeriesLabels.count Then
                    strSeries = .SeriesLabels(lngSeries).Text & ":"
                End If
                                
                If .Data.Layout = oc2dDataGeneral Then
                    dblX = .Data.X(lngSeries, lngPoint)
                Else
                    dblX = .Data.X(1, lngPoint)
                End If
                
                If Chart(Index).ChartArea.Axes("X").AnnotationMethod = oc2dAnnotateTimeLabels Then '��1970-01-01 08:00:00��ʼ������
                    strX = Format(DateAdd("s", dblX, CDate("1970-01-01 08:00:00")), "yyyy-MM-dd HH:mm:ss")
                    strX = Replace(strX, " 00:00:00", "")
                    strX = Replace(strX, ":00:00", "")
                    strX = Replace(strX, ":00", "")
                Else
                    strX = dblX
                End If
                strY = .Data.Y(lngSeries, lngPoint)
                
                If Chart(Index).ChartArea.Axes("X").Title.Text <> "" Then
                    strLabelX = Chart(Index).ChartArea.Axes("X").Title.Text & "="
                End If
                If Chart(Index).ChartArea.Axes("Y").Title.Text <> "" Then
                    strLabelY = Chart(Index).ChartArea.Axes("Y").Title.Text & "="
                End If
                
                sta.Panels(3).Text = strSeries & strLabelX & strX & "," & strLabelY & strY
            Else
                sta.Panels(3).Text = ""
            End If
        Else
            sta.Panels(3).Text = ""
        End If
    End With
End Sub

Private Sub cmdDefault_Click()
    Dim sngTop As Single
    
    sngTop = cmdDefault.Top + cmdDefault.Height + picPar.Top + IIF(cbr.Visible, cbr.Height, 0) + 15
    Call Me.PopupMenu(mnuPop, , cmdDefault.Left + 30, sngTop)
End Sub

Private Sub cmdLoad_Click()
    mnuView_reFlash_Click
End Sub

Private Sub cmdSelAll_Click()
    Dim chkTmp As CheckBox
    
    For Each chkTmp In chk
        chkTmp.Value = 1
    Next
End Sub

Private Sub cmdSelNone_Click()
    Dim chkTmp As CheckBox
    
    For Each chkTmp In chk
        chkTmp.Value = 0
    Next
End Sub

Private Sub Form_Activate()
    Dim tmpMsh As Object
    Static blnAct As Boolean
    
    If blnExcel Then blnExcel = False: Exit Sub
    
    cbr.Bands(2).Width = cbr.Bands(2).Width + 15
    cbr.Bands(2).Width = cbr.Bands(2).Width - 15
    
    '�����¼�
    If Not mobjCurDLL Is Nothing Then
        Call mobjCurDLL.Act_ReportActive(mobjReport.���, Me)
    End If
    
    If cbr.Bands(1).Visible Then cbr.Bands(1).MinHeight = tbr.ButtonHeight

    '��λ�ڵ�һ�������
    If Not blnAct Then
        blnAct = True
        For Each tmpMsh In msh
            If tmpMsh.Index <> 0 And tmpMsh.Container Is picPaper(intReport) And Not tmpMsh.Tag Like "H_*" Then
                Call msh_EnterCell(tmpMsh.Index)
                On Error Resume Next
                tmpMsh.SetFocus: Exit For
            End If
        Next
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not (scrVsc.Visible And scrHsc.Visible) And KeyCode <> vbKeyF3 Then Exit Sub
    Select Case KeyCode
        Case vbKeyUp
            If scrVsc.Enabled And scrVsc.Value > scrVsc.Min Then
                If Shift = 2 Then
                    scrVsc.Value = IIF(scrVsc.Value - scrVsc.LargeChange < scrVsc.Min, scrVsc.Min, scrVsc.Value - scrVsc.LargeChange)
                Else
                    scrVsc.Value = IIF(scrVsc.Value - scrVsc.SmallChange < scrVsc.Min, scrVsc.Min, scrVsc.Value - scrVsc.SmallChange)
                End If
            End If
        Case vbKeyDown
            If scrVsc.Enabled And scrVsc.Value < scrVsc.Max Then
                If Shift = 2 Then
                    scrVsc.Value = IIF(scrVsc.Value + scrVsc.LargeChange > scrVsc.Max, scrVsc.Max, scrVsc.Value + scrVsc.LargeChange)
                Else
                    scrVsc.Value = IIF(scrVsc.Value + scrVsc.SmallChange > scrVsc.Max, scrVsc.Max, scrVsc.Value + scrVsc.SmallChange)
                End If
            End If
        Case vbKeyLeft
            If scrHsc.Enabled And scrHsc.Value > scrHsc.Min Then
                If Shift = 2 Then
                    scrHsc.Value = IIF(scrHsc.Value - scrHsc.LargeChange < scrHsc.Min, scrHsc.Min, scrHsc.Value - scrHsc.LargeChange)
                Else
                    scrHsc.Value = IIF(scrHsc.Value - scrHsc.SmallChange < scrHsc.Min, scrHsc.Min, scrHsc.Value - scrHsc.SmallChange)
                End If
            End If
        Case vbKeyRight
            If scrHsc.Enabled And scrHsc.Value < scrHsc.Max Then
                If Shift = 2 Then
                    scrHsc.Value = IIF(scrHsc.Value + scrHsc.LargeChange > scrHsc.Max, scrHsc.Max, scrHsc.Value + scrHsc.LargeChange)
                Else
                    scrHsc.Value = IIF(scrHsc.Value + scrHsc.SmallChange > scrHsc.Max, scrHsc.Max, scrHsc.Value + scrHsc.SmallChange)
                End If
            End If
        Case vbKeyF3
            Call FindItem(txtFind.Text, True)
    End Select
End Sub

Private Sub Form_Load()
    Dim strErr As String, i As Integer, j As Integer
    Dim objItem As Object, rsTmp As ADODB.Recordset
    Dim strPrivs As String, lng����ID As Long, lngϵͳID As Long
    Dim blnPriv As Boolean, bytMode As Byte
    Dim strSQL As String, lngReport As Long
    Dim rsReport As New ADODB.Recordset
    Dim frmNewParInput As New frmParInput
    Dim strBasePrivs As String
    Dim strTmp As String
    Dim strStartTime As String
    Dim rsData As ADODB.Recordset
    Dim lngPersonID As Long
    
    Set objScript = New clsScript
    Srt.AddObject "clsScript", objScript, True
    
    garrBill = Empty
    mblnPrintEmpty = False
    bytFormat = 0
    blnExcel = False

    '��ȡ��������
    If gobjReport Is Nothing Then
        '�򿪱�����
        '��ʾ��������Ϣ
        Set rsTmp = GetGroupInfo(glngGroup)
        If rsTmp Is Nothing Then Unload Me: Exit Sub '�����˳�
        Caption = rsTmp!����
        lblGroup_S.Caption = lblGroup_S.Caption & ":" & rsTmp!����
        Me.Tag = rsTmp!��� '���뱨������
        
        lngϵͳID = IIF(IsNull(rsTmp!ϵͳ), 0, rsTmp!ϵͳ)
        lng����ID = IIF(IsNull(rsTmp!����id), 0, rsTmp!����id)
        
        '����ѡ��ģʽ
        bytMode = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.name & Me.Tag, "ѡ��ģʽ", 0)
        If bytMode = 1 Then
            Call mnuEdit_SelMode_Col_Click
        Else
            Call mnuEdit_SelMode_Row_Click
        End If
        
        'װ�뼰��ʾ����
        Set rsTmp = GetSubReport(glngGroup)
        If rsTmp Is Nothing Then
            MsgBox "�ñ�������û���κα������ִ�У�", vbInformation, App.Title
            Unload Me: Exit Sub '�����˳�
        End If
        Screen.MousePointer = 11
        strPrivs = GetPrivFunc(lngϵͳID, lng����ID)
        i = 0
        Do While Not rsTmp.EOF
            'δ��Ȩ���ӱ������ҷǹ����ߵ��ã��Ͳ��г����ӱ���
            If InStr(";" & strPrivs & ";", ";" & Nvl(rsTmp!����, "NONE") & ";") <= 0 _
                And Not mobjCurDLL Is Nothing Then
                GoTo makContinue
            End If
            
            blnPriv = True
            '�Ϸ����ж�
            blnPriv = CheckPass(rsTmp!����ID)
            'Ȩ���ж�
            If lng����ID > 0 And Not IsNull(rsTmp!����) And blnPriv Then
                blnPriv = (InStr(";" & strPrivs & ";", ";" & rsTmp!���� & ";") > 0)
            End If
            If blnPriv Then
                If i = 0 Then
                    ReDim arrReport(0)
                    ReDim arrLibDatas(0) '������ʱδװ��
                    ReDim arrDefPars(0)
                Else
                    Load picPaper(i): picPaper(i).Visible = False
                    ReDim Preserve arrReport(i)
                    ReDim Preserve arrLibDatas(i) '������ʱδװ��
                    ReDim Preserve arrDefPars(i)
                End If
                
                '��������
                Set arrReport(i) = New Report
                Set arrReport(i) = ReadReport(rsTmp!����ID)
                Call ReplaceSysNo(arrReport(i)) '������������е�ϵͳ����
                Call GetUserName(arrReport(i).ϵͳ, gstrUserName, gstrUserNO)
                Call SetReportIndex(i, arrReport(i))
                
                'ȱʡ��������
                Set arrDefPars(i) = New RPTPars
                Set arrDefPars(i) = MakeNamePars(arrReport(i))
                
                Set objItem = lvw.ListItems.Add(, "_" & rsTmp!����ID, arrReport(i).����, "Report", "Report")
                objItem.SubItems(1) = arrReport(i).���
                objItem.SubItems(2) = arrReport(i).˵��
                
                 '���еı����ʽ�в��õ�����Դɾ��
                If arrReport(i).Datas.count > 0 Then Call DelUnUseData(arrReport(i))
                '�滻�û�ͨ����������Ĳ���
                If ParCount(arrReport(i)) > 0 Then Call ReplaceUserPars(arrReport(i))
                
                '��ע�����ϴθ�ʽ,ȱʡ1
                arrReport(i).bytFormat = CByte(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.name & arrReport(i).���, "��ʽ", 1))
                
                i = i + 1
            End If
            
makContinue:
            rsTmp.MoveNext
        Loop
        
        Screen.MousePointer = 0
        If rsTmp.RecordCount > 0 And lvw.ListItems.count = 0 Then
            MsgBox "��û��Ȩ��ִ�иñ������е����б������鱨��ĺϷ����Լ��Ƿ���ȷ��Ȩ��", vbInformation, App.Title
            Unload Me: Exit Sub '�����˳�
        ElseIf lvw.ListItems.count = 0 Then
            Unload Me: Exit Sub '�����˳�
        End If
        
        mnuEdit_Par.Visible = False
        mnuEdit_Par_.Visible = False
        tbr.Buttons("Par").Caption = "��װ"
        tbr.Buttons("Par").Tag = "��װ"
        
        lvw.ColumnHeaders(2).Position = 1
        RestoreWinState Me, App.ProductName, Me.Tag

        SetView lvw.View
        
        '���ñ����б�߶ȣ��Ծ���������������߶�
        lvw.Height = lvw.ListItems.count * 350
        If lvw.Height < 1000 Then lvw.Height = 1000
        If lvw.Height > picGroup.Height / 2 Then
            lvw.Height = picGroup.Height / 2
        End If
        
        picLR_S.Visible = mnuViewToolGroup.Checked
        picGroup.Visible = mnuViewToolGroup.Checked
        
        If Not lvw.SelectedItem Is Nothing Then Call lvw_ItemClick(lvw.SelectedItem)
    Else
        '�򿪵�������
        picBack.BorderStyle = 0
        picLR_S.Visible = False
        picGroup.Visible = False
        For i = 0 To mnuViewStyle.UBound
            mnuViewStyle(i).Visible = False
        Next
        mnuViewStyle_.Visible = False
        mnuView_Pre.Visible = False
        mnuView_Next.Visible = False
        mnuView_Page_.Visible = False
        mnuViewToolGroup.Visible = False

        tbr.Buttons("Style").Visible = False
        tbr.Buttons("Style_").Visible = False
        tbr.Buttons("Pre").Visible = False
        tbr.Buttons("Next").Visible = False
        tbr.Buttons("Page_").Visible = False
        
        intReport = 0
        Call CopyReport(gobjReport, mobjReport)
        Call ReplaceSysNo(mobjReport) '������������е�ϵͳ����
        Call GetUserName(mobjReport.ϵͳ, gstrUserName, gstrUserNO)
        Call SetReportIndex(intReport, mobjReport)
        Caption = mobjReport.����
        
        If mbytStyle = 0 Then '����ʾ����ʱ�Ͳ������Լӿ��ٶ�
            RestoreWinState Me, App.ProductName, mobjReport.���
        End If
        
        '��ѯ�Ƿ������ڵ�ǰʱ��ִ�д˱���
        If Format(mobjReport.��ֹ����ʱ��, "HH:mm:ss") <> "00:00:00" Or Format(mobjReport.��ֹ��ʼʱ��, "HH:mm:ss") <> "00:00:00" Then
            If CDate(Format(mobjReport.��ֹ����ʱ��, "HH:mm:ss")) > CDate(Format(mobjReport.��ֹ��ʼʱ��, "HH:mm:ss")) Then
                If Between(CDate(Format(Currentdate, "HH:mm:ss")), CDate(Format(mobjReport.��ֹ��ʼʱ��, "HH:mm:ss")), CDate(Format(mobjReport.��ֹ����ʱ��, "HH:mm:ss"))) Then
                    MsgBox "��ǰ������" & CDate(Format(mobjReport.��ֹ��ʼʱ��, "HH:mm:ss")) & "-" & CDate(Format(mobjReport.��ֹ����ʱ��, "HH:mm:ss")) & "��ִֹ�У�������������ϵ��Ϣ�ơ�", vbInformation, App.Title
                    Unload Me: Exit Sub
                End If
            Else
                If CDate(Format(Currentdate, "HH:mm:ss")) < CDate(Format(mobjReport.��ֹ��ʼʱ��, "HH:mm:ss")) Or CDate(Format(Currentdate, "HH:mm:ss")) > CDate(Format(mobjReport.��ֹ����ʱ��, "HH:mm:ss")) Then
                    MsgBox "��ǰ������" & CDate(Format(mobjReport.��ֹ��ʼʱ��, "HH:mm:ss")) & "-�ڶ���" & CDate(Format(mobjReport.��ֹ����ʱ��, "HH:mm:ss")) & "��ִֹ�У�������������ϵ��Ϣ�ơ�", vbInformation, App.Title
                    Unload Me: Exit Sub
                End If
            End If
        End If
        
        '����ѡ��ģʽ
        bytMode = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.name & mobjReport.���, "ѡ��ģʽ", 0)
        If bytMode = 1 Then
            Call mnuEdit_SelMode_Col_Click
        Else
            Call mnuEdit_SelMode_Row_Click
        End If
        
        '��¼����ִ�п�ʼʱ��
        If gblnReportRunLog Then
            strStartTime = Format(Currentdate, "YYYY-MM-DD HH:mm:SS")
        End If
        If Not mobjCurDLL Is Nothing Then
            Call mobjCurDLL.Act_BeforeReportLoad(mobjReport.���, Me)
        End If
    
         '���еı����ʽ�в��õ�����Դɾ��
        If mobjReport.Datas.count > 0 Then Call DelUnUseData(mobjReport)
    
        'ȱʡ��ʾ��һ�ָ�ʽ
        bytFormat = 1
        
        '���ݱ��ش�ӡ���ö�ȡҪ��ӡ�ĸ�ʽ
        '����ǳ����ڲ����ô�ӡ��Ҫ��ӡ���и�ʽ����ǰĬ��Ϊ��һ�ָ�ʽ
        strTmp = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\LocalSet\" & mobjReport.���, "AllFormat", "")
        If strTmp = "" Then strTmp = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\LocalSet\" & mobjReport.���, "AllFormat", 0)
        mblnAllFormat = Val(strTmp) = 1
        If Not (mbytStyle = 2 And mblnAllFormat) Then
            '����ע���ȡ�û���ʽ
            bytFormat = CByte(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.name & mobjReport.���, "��ʽ", 1))
            '��ȡ��ӡ����ָ���ĸ�ʽ 'If mobjReport.Ʊ�� Then
            strTmp = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\LocalSet\" & mobjReport.���, "Format", "")
            If strTmp = "" Then
                i = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\LocalSet\" & mobjReport.���, "Format", -1))
            Else
                i = Val(strTmp)
            End If
            If i <> -1 Then bytFormat = i
        End If
        
        'ȡ�ñ����ID
        lngReport = 0
        If ReportReaded(, mobjReport.���, mobjReport.ϵͳ) Then
            lngReport = grsReport!id '���û���
        Else
            strSQL = "Select ID,���,����,˵��,����,��ӡ��,��ֽ,Ʊ��,��ӡ��ʽ,ϵͳ,����ID,����,�޸�ʱ��,����ʱ��,��ֹ��ʼʱ��,��ֹ����ʱ�� From zlReports Where ���=[1] And Nvl(ϵͳ,0)=[2]"
            Set rsReport = OpenSQLRecord(strSQL, Me.Caption, mobjReport.���, mobjReport.ϵͳ)
            If Not rsReport.EOF Then '���洦��
                Set grsReport = New ADODB.Recordset
                Set grsReport = rsReport
                gdatModiTime = grsReport!�޸�ʱ��
                
                lngReport = rsReport!id
            End If
        End If
        mlngReportID = lngReport
        mlngRPTID = lngReport
        
        '�����û��������ȡһЩ���Ʋ���
        If IsArray(marrPars) Then
            If UBound(marrPars) <> -1 Then
                For i = 0 To UBound(marrPars)
                    j = InStr(CStr(marrPars(i)), "=")
                    If j > 0 Then
                        'ReportFormat
                        If UCase(Trim(Left(CStr(marrPars(i)), j - 1))) = UCase("ReportFormat") Then
                            If IsNumeric(Trim(Mid(CStr(marrPars(i)), j + 1))) Then
                                bytFormat = CByte(Trim(Mid(CStr(marrPars(i)), j + 1)))
                                mblnAllFormat = False '����ָ���˸�ʽ���ӡ������Ч
                            End If
                        'DisabledPrint
                        ElseIf UCase(Trim(Left(CStr(marrPars(i)), j - 1))) = UCase("DisabledPrint") Then
                            If IsNumeric(Trim(Mid(CStr(marrPars(i)), j + 1))) Then
                                mblnDisabledPrint = CByte(Trim(Mid(CStr(marrPars(i)), j + 1))) = 1
                            End If
                        'PrintEmpty
                        ElseIf UCase(Trim(Left(CStr(marrPars(i)), j - 1))) = UCase("PrintEmpty") Then
                            If IsNumeric(Trim(Mid(CStr(marrPars(i)), j + 1))) Then
                                mblnPrintEmpty = CByte(Trim(Mid(CStr(marrPars(i)), j + 1))) = 1
                            End If
                        'ExcelFile
                        ElseIf UCase(Trim(Left(CStr(marrPars(i)), j - 1))) = UCase("ExcelFile") Then
                            mstrExcelFile = Trim(Mid(CStr(marrPars(i)), j + 1))
                        'PDF
                        ElseIf UCase(Trim(Left(CStr(marrPars(i)), j - 1))) = UCase("PDF") Then
                            mstrPDFFile = Trim(Mid(CStr(marrPars(i)), j + 1))
                        End If
                    End If
                Next
            End If
        End If
        
        '�������и�ʽ
        For i = 1 To mobjReport.Fmts.count
            Set objItem = cboFormat.ComboItems.Add(, "_" & mobjReport.Fmts(i).���, mobjReport.Fmts(i).˵��, "Format")
            If mobjReport.Fmts(i).��� = bytFormat Then objItem.Selected = True
        Next
        If cboFormat.SelectedItem Is Nothing And cboFormat.ComboItems.count > 0 Then
            cboFormat.ComboItems(1).Selected = True
            bytFormat = CByte(Mid(cboFormat.SelectedItem.Key, 2))
        End If
        mobjReport.bytFormat = bytFormat
        mnuFile_Graph.Enabled = (mobjReport.Fmts("_" & bytFormat).ͼ�� <> 0)
        tbr.Buttons("Graph").Enabled = (mobjReport.Fmts("_" & bytFormat).ͼ�� <> 0)
                
'        If cboFormat.ComboItems.Count = 1 Then
'            mnuViewToolFormat.Checked = False
'            cbr.Bands(2).Visible = False
'        End If
        cboFormat.Locked = cboFormat.ComboItems.count > 1
                
        '��������
        If ParCount(mobjReport) > 0 Then
            If Not ReplaceUserPars(mobjReport) Then
                'δȫ����ȷ�ش������,��Ҫ���������
                
                Set mobjPars = MakeNamePars(mobjReport)
                Call CopyPars(mobjPars, mobjDefPars)
                frmNewParInput.mlngReport = lngReport
                Set frmNewParInput.mobjPars = mobjPars
                Set frmNewParInput.mobjDefPars = mobjDefPars
                Set frmNewParInput.mobjRPTDatas = mobjReport.Datas
                
                frmNewParInput.mstrTitle = mobjReport.����
                frmNewParInput.mblnReset = False
                frmNewParInput.Show 1, Me
                
                If frmNewParInput.mblnOK Then
                    '���������ύ�¼�
                    If Not mobjCurDLL Is Nothing Then
                        mobjCurDLL.Act_CommitCondition mobjReport.���, GetParsStr(frmNewParInput.mobjPars), Me
                    End If
                    
                    ReplaceInputPars frmNewParInput.mobjPars
                    Unload frmNewParInput
                Else
                    Unload Me: Exit Sub '��һ��ȡ�����˳�
                End If
            Else
                'ȫ����ȷ�������,��Ҳ�������ñ�������
                tbr.Buttons("Par").Visible = False
                mnuEdit_Par.Visible = False
                tbr.Buttons("Par_").Visible = False
                mnuEdit_Par_.Visible = False
                
                Set mobjDefPars = MakeNamePars(mobjReport)
                
                '���������ύ�¼�
                If Not mobjCurDLL Is Nothing Then
                    mobjCurDLL.Act_CommitCondition mobjReport.���, GetParsStr(MakeNamePars(mobjReport, True)), Me
                End If
            End If
        Else
            '���������ύ�¼�
            If Not mobjCurDLL Is Nothing Then
                mobjCurDLL.Act_CommitCondition mobjReport.���, "ReportFormat=" & bytFormat, Me
            End If
            
            '���û�ж������,��Ҳ�������ñ�������
            tbr.Buttons("Par").Visible = False
            mnuEdit_Par.Visible = False
            tbr.Buttons("Par_").Visible = False
            mnuEdit_Par_.Visible = False
        End If
        
        'ʹ���ߴ����������������������ڱ��������ȱʡֵ��(mobjReport)
        '��������
        If Not frmParent Is Nothing Then frmParent.Refresh
        Me.Refresh
        strErr = OpenReportData(False)
        If strErr <> "" Then
            If gblnSilentMode = False Then
                MsgBox "�ڶ�ȡ��������""" & strErr & """ʱ�����������,�����ܲ�����", vbInformation, App.Title
            End If
            Unload Me: Exit Sub
        End If
        
        '��ʾ����
        Call ShowItems
       
        If Not mobjCurDLL Is Nothing Then
            Call mobjCurDLL.Act_AfterReportLoad(mobjReport.���, Me)
        End If
        Call RecordsExecute(lngReport, strStartTime, 3)
        
    End If
    
    'Excel�������ӡȨ���ж�
    strBasePrivs = GetPrivFunc(0, 16)
    If InStr(";" & strBasePrivs & ";", ";Excel���;") = 0 Then
        mnuFile_Excel.Visible = False
        mnuFile_Graph.Visible = False
        mnuFile_1.Visible = False
        tbr.Buttons("Graph").Visible = False
        tbr.Buttons(5).Visible = False
    End If
    If InStr(";" & strBasePrivs & ";", ";��ӡ;") = 0 Or mblnDisabledPrint Then
        mnuFile_Print.Visible = False
        tbr.Buttons("Print").Visible = False
    End If
    
End Sub

Private Sub RecordsExecute(ByVal lngReportID As Long, ByVal strStartTime As String, _
    Optional ByVal intType As Integer = 0)
'���ܣ���¼����ִ��
'������
'  lngReportID������ID
'  strStartTime��
'  intType��Ҫ��¼�����ͣ�2-����������־��3-����ʹ��״̬�ͱ���������־

    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim blnRunLog As Boolean
    Dim strEndTime As Date
    
    If intType <= 0 Then Exit Sub
    If Not (gblnReportUse Or gblnReportRunLog) Then Exit Sub
    
    On Error GoTo ErrHand
    strEndTime = Format(Currentdate, "YYYY-MM-DD HH:mm:SS")
    Select Case intType
    Case 2
        GoSub makTwo
    Case 3
        GoSub makOne
        GoSub makTwo
    End Select
    Exit Sub
    
makOne:
    If gblnReportUse Then
        strSQL = "Zl_Rptrun_Update(" & lngReportID & "," & _
                    "'" & gstrUserName & "')"
        Call ExecuteProcedure(strSQL, "�������м�¼")
    End If
    Return
    
makTwo:
    If gblnReportRunLog Then
        strSQL = "Zl_Rptrunhistory_Update(" & _
                    lngReportID & "," & _
                    "'" & gstrUserName & "'," & _
                    "to_date('" & strStartTime & "','YYYY-MM-DD HH24:MI:SS')," & _
                    "to_date('" & strEndTime & "','YYYY-MM-DD HH24:MI:SS'))"
        Call ExecuteProcedure(strSQL, "����������־")
    End If
    Return
    
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub lbl_Click(Index As Integer)
    Dim objRelations As RPTRelations
    Dim i As Long
    Dim lngRec As Long, strDataName As String
    Dim strLisName As String
    Dim lngRelationReport As Long
    Dim tmpData As RPTData, tmpPar As RPTPar
    
    If mblnLeftClick = False And mlngRelationReport = 0 Then Exit Sub
    
    If lbl(Index).Tag <> "" Then
        Set objRelations = mobjReport.Items("_" & Index).Relations
        lngRec = Val(lbl(Index).Tag)
        For i = 1 To objRelations.count
            If objRelations.Item(i).Ĭ�� = 1 Then
                lngRelationReport = objRelations.Item(i).��������ID
                Exit For
            End If
        Next
        If lngRelationReport = 0 Then lngRelationReport = objRelations.Item(1).��������ID
        If mlngRelationReport <> 0 Then lngRelationReport = mlngRelationReport
        If Not CheckReportPriv(lngRelationReport) Then
            MsgBox "��û��Ȩ�޲�ѯ�ñ���ĳЩ����Դ�еĶ���", vbInformation, App.Title: Exit Sub
        End If
        'ִ�б���
        If CheckPass(lngRelationReport) = False Then
            MsgBox "�������ݴ��󣬲���ִ�иñ���", vbInformation, App.Title: Exit Sub
        End If
        
        Set gobjReport = ReadReport(lngRelationReport)
        '��ʼ������
        garrPars = Array()
        '��λ��¼��
        On Error Resume Next
        For i = 1 To objRelations.count
            If objRelations.Item(i).��������ID = lngRelationReport Then
                If InStr(objRelations.Item(i).����ֵ��Դ, ".") > 0 Then
                    strDataName = Mid(objRelations.Item(i).����ֵ��Դ, 1, InStr(objRelations.Item(i).����ֵ��Դ, ".") - 1)
                End If
            End If
            If strDataName <> "" Then Exit For
        Next

        '����ľ���������
        If strDataName <> "" Then mLibDatas("_" & strDataName).DataSet.AbsolutePosition = lngRec
        
        For i = 1 To objRelations.count
            With objRelations.Item(i)
                strLisName = ""
                If objRelations.Item(i).��������ID = lngRelationReport Then
                    If InStr(.����ֵ��Դ, ".") > 0 Then
                        If mLibDatas("_" & strDataName).DataSet.RecordCount > 0 Then
                            strLisName = mLibDatas("_" & strDataName).DataSet.Fields(Mid(.����ֵ��Դ, InStr(.����ֵ��Դ, ".") + 1)).Value
                        End If
                    ElseIf InStr(.����ֵ��Դ, "=") = 1 Then
                        For Each tmpData In mobjReport.Datas
                            For Each tmpPar In tmpData.Pars
                                If tmpPar.���� = Mid(.����ֵ��Դ, 2) Then
                                    strLisName = tmpPar.ȱʡֵ
                                    Exit For
                                End If
                            Next
                            If strLisName <> "" Then Exit For
                        Next
                    End If
                    ReDim Preserve garrPars(UBound(garrPars) + 1)
                    garrPars(UBound(garrPars)) = .������ & "=" & strLisName
                End If
            End With
        Next
        
        
        If Not ShowReport(Me) Then MsgBox "�����ʧ�ܣ�", vbInformation, App.Title
    End If
End Sub

Private Sub lbl_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lbl(Index).Tag <> "" Then lbl(Index).MousePointer = 99
End Sub

Private Sub lbl_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    Dim dbRowHeight As Double

    If Button = 2 Then
        mintLblIndex = Index
        mblnLeftClick = False
        If lbl(Index).FontUnderline = True Then
            Call LoadRelation(1, Index)
            vsfRelations.Visible = True
            vsfRelations.SetFocus
            For i = 0 To vsfRelations.Rows - 1
                dbRowHeight = dbRowHeight + vsfRelations.RowHeight(i)
            Next
            vsfRelations.Height = dbRowHeight
            vsfRelations.Left = lbl(Index).Left + X + 150
            vsfRelations.Top = lbl(Index).Top + 90
        Else
            vsfRelations.Visible = False
        End If
    Else
        mblnLeftClick = True
    End If
End Sub

Private Sub lblPar_S_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then lngPreY = Y
End Sub

Private Sub lblPar_S_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If lblPar_S.Top + Y - lngPreY < 1000 Or picPar.Height - (Y - lngPreY) < 1000 Then Exit Sub
        lblPar_S.Top = lblPar_S.Top + Y - lngPreY
        lvw.Height = lvw.Height + Y - lngPreY
        picPar.Top = picPar.Top + Y - lngPreY
        picPar.Height = picPar.Height - (Y - lngPreY)
        Me.Refresh
    End If
End Sub

Private Sub lvw_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static blnDesc As Boolean
    Static intIdx As Integer
    
    If intIdx = ColumnHeader.Index Then
        blnDesc = Not blnDesc
    Else
        blnDesc = False
    End If
    lvw.SortKey = ColumnHeader.Index - 1
    If blnDesc Then
        lvw.SortOrder = lvwDescending
    Else
        lvw.SortOrder = lvwAscending
    End If
    lvw.Sorted = True
    intIdx = ColumnHeader.Index
    
    If Not lvw.SelectedItem Is Nothing Then lvw.SelectedItem.EnsureVisible
End Sub

Private Sub lvw_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim i As Integer, objItem As Object
    
    Set objCurGrid = Nothing
    
    LockWindowUpdate Me.hwnd
    
    '����ǰһ�ű��������״̬
    If Not mobjReport Is Nothing Then
        Call CopyReport(mobjReport, arrReport(intReport)) '����
        Call CopyPars(mobjDefPars, arrDefPars(intReport)) 'ȱʡ��������
        If mLibDatas Is Nothing Then '����Դ����
            Set arrLibDatas(intReport) = Nothing
        Else
            Call CopyLibDatas(mLibDatas, arrLibDatas(intReport))
        End If
    End If
    
    '��ȡ��ǰ��������״̬
    intReport = Item.Index - 1
    Call CopyReport(arrReport(intReport), mobjReport) '����
    Call CopyPars(arrDefPars(intReport), mobjDefPars) 'ȱʡ��������
    If arrLibDatas(intReport) Is Nothing Then '����Դ����
        Set mLibDatas = Nothing
    Else
        Call CopyLibDatas(arrLibDatas(intReport), mLibDatas)
    End If
    
    bytFormat = mobjReport.bytFormat
    intGridCount = mobjReport.intGridCount
    intGridID = mobjReport.intGridID
        
    '�����ʽ
    cboFormat.ComboItems.Clear
    For i = 1 To mobjReport.Fmts.count
        Set objItem = cboFormat.ComboItems.Add(, "_" & mobjReport.Fmts(i).���, mobjReport.Fmts(i).˵��, "Format")
        If mobjReport.Fmts(i).��� = bytFormat Then objItem.Selected = True
    Next
    If cboFormat.SelectedItem Is Nothing And cboFormat.ComboItems.count > 0 Then
        cboFormat.ComboItems(1).Selected = True
        bytFormat = CByte(Mid(cboFormat.SelectedItem.Key, 2))
        mobjReport.bytFormat = bytFormat
    End If
    cboFormat.Refresh
    cboFormat.Locked = cboFormat.ComboItems.count > 1
   
    mnuFile_Graph.Enabled = (mobjReport.Fmts("_" & bytFormat).ͼ�� <> 0)
    tbr.Buttons("Graph").Enabled = (mobjReport.Fmts("_" & bytFormat).ͼ�� <> 0)
    
    '��ʾ��ǰֽ��
    picBack.Visible = False
    For i = 0 To picPaper.UBound
        picPaper(i).Visible = (i = intReport)
    Next
    picPaper(intReport).ZOrder
    
    scrVsc.Visible = Not (intGridCount = 1 And Not mobjReport.Ʊ��)
    scrHsc.Visible = Not (intGridCount = 1 And Not mobjReport.Ʊ��)
    picShadow.Visible = Not (intGridCount = 1 And Not mobjReport.Ʊ��)
    If Not (intGridCount = 1 And Not mobjReport.Ʊ��) Then
        scrVsc.Value = scrVsc.Min
        scrHsc.Value = scrHsc.Min
        Call scrhsc_Change
        Call scrVsc_Change
    End If
    
    '����ҳ��
    Call Form_Resize
    picBack.Visible = True

    '��ʾ����
    picPar.Visible = False

    Call CopyPars(mobjDefPars, mobjPars)
    mlngRPTID = Val(Mid(lvw.SelectedItem.Key, 2))
    Call InitReportPars
    picPar.Visible = True
    
    LockWindowUpdate 0

    '��λ�ڵ�һ�������
    For Each objItem In msh
        If objItem.Index <> 0 And objItem.Container Is picPaper(intReport) And Not objItem.Tag Like "H_*" Then
            Call msh_EnterCell(objItem.Index)
            Exit For
        End If
    Next
End Sub

Private Sub mnuEdit_SelMode_Col_Click()
    Dim tmpMsh As Object
    
    '(��������)���б���
    For Each tmpMsh In msh
        If tmpMsh.Index <> 0 And Not tmpMsh.Tag Like "H_*" Then
            Call msh_LeaveCell(tmpMsh.Index)
        End If
    Next
    
    mnuEdit_SelMode_Col.Checked = True
    mnuEdit_SelMode_Row.Checked = False
    
    '(��������)���б���
    For Each tmpMsh In msh
        If tmpMsh.Index <> 0 And Not tmpMsh.Tag Like "H_*" Then
            Call msh_EnterCell(tmpMsh.Index)
        End If
        '��ѡ��
        msh(tmpMsh.Index).SelectionMode = flexSelectionByColumn
    Next
End Sub

Private Sub mnuEdit_SelMode_Row_Click()
    Dim tmpMsh As Object
    
    '(��������)���б���
    For Each tmpMsh In msh
        If tmpMsh.Index <> 0 And Not tmpMsh.Tag Like "H_*" Then
            Call msh_LeaveCell(tmpMsh.Index)
        End If
    Next
    
    mnuEdit_SelMode_Row.Checked = True
    mnuEdit_SelMode_Col.Checked = False
    
    '(��������)���б���
    For Each tmpMsh In msh
        If tmpMsh.Index <> 0 And Not tmpMsh.Tag Like "H_*" Then
            Call msh_EnterCell(tmpMsh.Index)
        End If
        '��ѡ��
        msh(tmpMsh.Index).SelectionMode = flexSelectionByRow
    Next
End Sub

Private Sub mnuEdit_SetCol_Auto_Click()
'���ܣ���������п�Ϊ��С��Ӧ���ֿ��,�����һ�й̶���Ϊ׼(�����)
    Dim tmpMsh As Object
    Dim i As Integer

    If Not mobjReport.blnLoad Then Exit Sub
    
    On Error Resume Next
    
    Screen.MousePointer = 11
    
    LockWindowUpdate Me.hwnd
    
    For Each tmpMsh In msh
        If tmpMsh.Index <> 0 Then
            If tmpMsh.Container Is picPaper(intReport) Then
                Call SetColWidth(tmpMsh)
            ElseIf UCase(tmpMsh.Container.name) = "PIC" Then
                If tmpMsh.Container.Container Is picPaper(intReport) Then
                    Call SetColWidth(tmpMsh)
                End If
            End If
        End If
    Next
    '��ͷ�������ͬ�еĿ��ȡ�ϴ���
    For Each tmpMsh In msh
        If tmpMsh.Index <> 0 And (tmpMsh.Container Is picPaper(intReport) Or UCase(tmpMsh.Container.name) = "PIC") _
            And tmpMsh.FixedRows = 0 Then
            For i = 0 To tmpMsh.Cols - 1
                If tmpMsh.ColWidth(i) > msh(tmpMsh.Tag).ColWidth(i) Then
                    msh(tmpMsh.Tag).ColWidth(i) = tmpMsh.ColWidth(i)
                Else
                    tmpMsh.ColWidth(i) = msh(tmpMsh.Tag).ColWidth(i)
                End If
            Next
            tmpMsh.LeftCol = 0: msh(tmpMsh.Tag).LeftCol = 0
            
            '�����и�
            Call AdjustRowHight(tmpMsh.Index)
        End If
    Next
    
    Screen.MousePointer = 0
    LockWindowUpdate 0
End Sub

Private Sub mnuEdit_SetCol_Def_Click()
'���ܣ���ԭ����п�Ϊȱʡ����п�
    Dim objItem As RPTItem, objCurItem As RPTItem
    Dim tmpItem As RPTItem, tmpID As RelatID
    Dim i As Integer, j As Integer, strWidth As String
    Dim lngColB As Long, lngColE As Long
    
    If Not mobjReport.blnLoad Then Exit Sub
        
    On Error Resume Next
    
    LockWindowUpdate Me.hwnd
    
    For Each objItem In mobjReport.Items
        If objItem.��ʽ�� = bytFormat Then
            If objItem.���� = 4 Then
                With objItem
                    For Each tmpID In .SubIDs
                        Set tmpItem = mobjReport.Items("_" & tmpID.id)
                        msh(.id).ColWidth(tmpItem.���) = tmpItem.W
                        msh(.SubIDs(1).id).ColWidth(tmpItem.���) = tmpItem.W
                        msh(.id).LeftCol = 0: msh(.SubIDs(1).id).LeftCol = 0
                    Next
                    '�����и�
                    Call AdjustRowHight(objItem.id)
                End With
            ElseIf objItem.���� = 5 And objItem.���� = 0 Then
                For i = 0 To UBound(Split(objItem.��ͷ, "|"))
                    Set objCurItem = mobjReport.Items("_" & Split(Split(objItem.��ͷ, "|")(i), ",")(0))
                    With objCurItem
                        strWidth = ""
                        For Each tmpID In .SubIDs
                            Set tmpItem = mobjReport.Items("_" & tmpID.id)
                            Select Case tmpItem.����
                                Case 7
                                    If i = 0 Then msh(objItem.id).ColWidth(tmpItem.���) = tmpItem.W
                                Case 9
                                    strWidth = strWidth & "," & tmpItem.W
                            End Select
                        Next
                        strWidth = Mid(strWidth, 2)
                        
                        If i = 0 Then
                            lngColB = msh(objItem.id).FixedCols
                        Else
                            lngColB = lngColE + 1
                        End If
                        lngColE = CLng(Split(Split(objItem.��ͷ, "|")(i), ",")(1)) - 1
                        
                        For j = lngColB To lngColE
                            msh(objItem.id).ColWidth(j) = _
                                CLng(Split(strWidth, ",")((j - lngColB) Mod (UBound(Split(strWidth, ",")) + 1)))
                        Next
                    End With
                Next
            End If
        End If
    Next
    
    '��Ը��ӱ������⴦��
    Call SetGridAlign
    
    LockWindowUpdate 0
End Sub

Private Sub mnuEdit_SetCol_Fill_Click()
'���ܣ��Զ������п�(����ǰ���б�����������,�Ҹ������и����ұ��ж���)
    Dim tmpMsh As VSFlexGrid
    Dim i As Integer
    Dim lngCurW As Long
    Dim sngScale As Single
    
    If Not mobjReport.blnLoad Then Exit Sub
    
    On Error Resume Next
    
    LockWindowUpdate Me.hwnd
    
    Call SetGridAlign(Val("1-����"))
    
    LockWindowUpdate 0
    
'    '����
'    For Each tmpMsh In msh
'        If tmpMsh.Index <> 0 And (tmpMsh.Container Is picPaper(intReport) Or UCase(tmpMsh.Container.name) = "PIC") Then
'            tmpMsh.Redraw = False
'
'            lngCurW = GetGridColWidth(tmpMsh)
'            If lngCurW < tmpMsh.Width - 300 Then
'                '����ҳ����������ƽ�����ı���
'                sngScale = (tmpMsh.Width - 300) / lngCurW
'                For i = 0 To tmpMsh.Cols - 1
'                    tmpMsh.ColWidth(i) = tmpMsh.ColWidth(i) * sngScale
'                Next
'            End If
'
'            '�����и�
'            Call AdjustRowHight(tmpMsh.Index)
'
'            tmpMsh.Redraw = True
'        End If
'    Next
'
'    LockWindowUpdate 0
End Sub

Private Sub mnuFile_Graph_Click()
    Dim objHead As Object
    Dim objItem As RPTItem
    Dim bytKind As Byte
    Dim tmpMsh As Object
    
    If Not mobjReport.blnLoad Then Exit Sub
    
    If zlRegInfo("��Ȩ����") <> "1" Then
        MsgBox "���û���԰汾����ʹ�øù��ܡ�", vbInformation, App.Title
        Exit Sub
    End If
    
    If intGridCount = 0 Then
        MsgBox "��ǰ������û�����ݱ�ɹ�ͼ�η�����", vbInformation, App.Title
        Exit Sub
    End If
    If objCurGrid Is Nothing Then
        If msh.count > 1 Then
            For Each tmpMsh In msh
                If tmpMsh.Index <> 0 And (tmpMsh.Container Is picPaper(intReport) Or UCase(tmpMsh.Container.name) = "PIC") And Not tmpMsh.Tag Like "H_*" Then
                    Set objCurGrid = tmpMsh
                    Exit For
                End If
            Next
        End If
        If objCurGrid Is Nothing Then
            MsgBox "����ѡ��һ��Ҫ����ͼ�η��������ݱ�", vbInformation, App.Title
            Exit Sub
        End If
    End If
    If objCurGrid.Tag Like "H_*" Then
        MsgBox "���ݱ�ͷ��������ͼ�η�����", vbInformation, App.Title
        Exit Sub
    End If
    
    Set objItem = mobjReport.Items("_" & objCurGrid.Index)
    If objItem.���� = 4 Then
        bytKind = GetGridStyle(mobjReport, objItem.id)
        If bytKind = 0 Then Set objHead = msh(CInt(objCurGrid.Tag))
    End If
    blnExcel = True
    Call ExcelChart(Me, objCurGrid, objHead, IIF(mobjReport.Items("_" & objCurGrid.Index).���� = 5, 1, 2), mobjReport.����, mobjReport.Fmts("_" & bytFormat).ͼ��)
End Sub

Private Sub mnuFile_Preview_Click()
    Dim frmShow As New frmPreview

    If Not mobjReport.blnLoad Then Exit Sub
    
    If mobjReport.Items.count = 0 Then Exit Sub
    
    If Not InitPrinter(Me) Then
        gblnError = True
        MsgBox "�豸��ʼ��ʧ��.������ϵͳû�а�װ��ӡ�����뵱ǰ���ò����ݣ�", vbInformation, App.Title: Exit Sub
    End If
    
    If Not CalcCellPage Then
        gblnError = True
        MsgBox "�޷�����ı���ʽ,�������ܼ�����", vbInformation, App.Title: Exit Sub
    End If
    If lbl(lngCurInx).BackColor = CON_SETFOCES And lngCurInx <> 0 Then
        lbl(lngCurInx).BackColor = lngTmpColor
        lngCurInx = 0: lngTmpColor = 0
    End If
    
    SetRedraw False
    
    Set frmShow.frmParent = Me
    
    If mbytStyle = Val("1-�Զ�Ԥ��") Then
        If Not frmParent Is Nothing Then
            On Error Resume Next
            frmShow.Show 1, frmParent
            If Err.Number <> 0 Then
                On Error GoTo 0
                frmShow.Show 1
            End If
            On Error GoTo 0
        Else
            frmShow.Show 1
        End If
    Else
        frmShow.Show 1, Me
    End If
    
    SetRedraw True
End Sub

Private Sub mnuFile_Print_Click()
    Dim objItem As RPTItem, strSource As String
    Dim lngPrintH As Long, blnReset As Boolean
    Dim blnExit As Boolean, intCopy As Integer
    Dim blnDo As Boolean, blnCancel As Boolean
    Dim k As Integer, i As Integer, j As Integer
    Dim arrBill As Variant, strItem As String
    Dim objFmt As RPTFmt, blnGoOn As Boolean
    Dim blnPrint As Boolean, blnALLEmpty As Boolean
    Dim strTmp As String
    Dim strDefault As String
    Dim lngEndPage As Long
    
    If Not mobjReport.blnLoad Then Exit Sub
    If mobjReport.Items.count = 0 Then Exit Sub
    
    If Not mobjCurDLL Is Nothing Then
        mobjCurDLL.DataIsEmpty = False
    End If
    blnALLEmpty = True
    
    strDefault = mobjReport.Fmts(mobjReport.bytFormat).˵��
    strTmp = GetRegPrinterInfo("PaperCopy", mobjReport.���, strDefault)
    intCopy = Val(strTmp)
    If intCopy < 1 Then intCopy = 1
    If gblnSingleTask Then intCopy = 1 '�౨�������ӡʱ��֧�ִ�ӡ����
    If mobjReport.Ʊ�� Then intCopy = 1 '�����Ʊ�ݣ���ֻ�ܴ�ӡ1��
    
    cboFormat.Tag = "1"
    blnGoOn = True
    Do While blnGoOn
        Set objFmt = mobjReport.Fmts("_" & mobjReport.bytFormat)

        'ֱ�Ӵ�ӡʱ,��ǰ��ʽ�ı��Ϊ��ʱ,�򲻴�ӡ
        blnExit = False
        'mblnPrintEmpty=Fale��ʾ����ǿ�ƴ�ӡ�ձ� ��ӡ��ʽ=1��ʾ�����ձ��ӡ��
        If mblnPrintEmpty = False And mobjReport.��ӡ��ʽ = Val("1-�ձ���ӡ") And InStr(";0;2;4;", ";" & mbytStyle & ";") > 0 Then
            strSource = ""
            For Each objItem In mobjReport.Items
                If objItem.��ʽ�� = bytFormat Then
                    If objItem.���� = 4 Then        '������
                        strItem = GetGridSource(objItem, True) '"������Ϣ,ҩƷ��Ϣ,..."
                        If strItem <> "" Then strSource = strSource & "," & strItem
                    ElseIf objItem.���� = 5 Then    '���ܱ��
                        strSource = strSource & "," & objItem.����
                    End If
                End If
            Next
            'ʹ��������Դ(����)���ж�,�������ֻ��ӡһЩ��ǩ֮���
            If strSource <> "" Then
                blnExit = True
                strSource = Mid(strSource, 2)
                For i = 0 To UBound(Split(strSource, ","))
                    On Error Resume Next
                    blnExit = blnExit And mLibDatas("_" & Split(strSource, ",")(i)).DataSet.RecordCount = 0
                    Err.Clear: On Error GoTo 0
                Next
            End If
            If blnExit Then GoTo NextFormat
        End If
        blnALLEmpty = False
        
        On Error GoTo errH
        
        If mbytStyle = Val("4-PDF") Then
            If PDFInitialize(objFmt) Then
                If PDFFile(mstrPDFFile, , , True) = False Then
                    MsgBox "δָ��PDF���·�����ļ��������飡", vbInformation, App.Title
                    Exit Sub
                End If
            Else
                Exit Sub
            End If
        Else
            '��ʼ����ӡ����Ϣ
            If Not InitPrinter(Me, intCopy) Then
                MsgBox "�豸��ʼ��ʧ��.������ϵͳû�а�װ��ӡ�����뵱ǰ���ò����ݣ�", vbInformation, App.Title
                gblnError = True: GoTo ExitHandle
            End If
        End If
        
        k = intCopy 'ȱʡΪǿ��ѭ����ӡk��
        If Printer.Copies = intCopy Then k = 1 '֧��ʱʹ�ô�ӡ������
        
        '�����ӡ����
        If Not CalcCellPage Then
            gblnError = True
            MsgBox "�޷�����ı���ʽ,�������ܼ�����", vbInformation, App.Title: GoTo ExitHandle
        End If
        If mbytStyle <> 2 And mbytStyle <> 4 Then
            If MsgBox("�������ݼ������,��ӡ�豸׼��������", vbQuestion + vbYesNo, App.Title) = vbNo Then GoTo ExitHandle
        End If
        
        If lbl(lngCurInx).BackColor = CON_SETFOCES And lngCurInx <> 0 Then
            lbl(lngCurInx).BackColor = lngTmpColor
            lngCurInx = 0: lngTmpColor = 0
        End If
        
        '�������ӡ֮ǰ�������ӡ�¼�
        If Not mobjCurDLL Is Nothing Then
            arrBill = Empty: blnCancel = False: i = 1
            If IsArray(marrPage) Then i = UBound(marrPage) + 1
            Call mobjCurDLL.Act_BeforePrint(mobjReport.���, i * intCopy, blnCancel, arrBill)
            If blnCancel Then GoTo ExitHandle
            
            'ʵ��Ҫ��ӡ��Ʊ������
            If IsArray(arrBill) Then garrBill = arrBill
        End If
        
        SetRedraw False
        
        'ֱ�Ӵ�ӡ����
        If mbytStyle <> 2 Then Screen.MousePointer = 11
        
        j = 0
        blnReset = False
        Do
            k = k - 1
            j = j + 1
            If Not IsArray(marrPage) Then
                If IsArray(marrPageCard) Then
                    '��Ƭ
                    GoTo makPage
                End If
                
                If mbytStyle <> Val("2-��ӡ") Then
                    If Printer.Copies <> intCopy And intCopy <> 1 Then
                        ShowFlash "���" & mobjReport.���� & ",�� 1 ҳ " & intCopy & " ��,��ǰ�� " & j & " ��", j / intCopy, Me
                    Else
                        ShowFlash "���" & mobjReport.���� & "��", 1, Me
                    End If
                End If
                
                '��̬���㼰����ֽ�Ÿ߶�
                If objFmt.��ֽ̬�� And objFmt.ֽ�� = 1 Then
                    Call PrintPage(0, Me, Me, 1, False, True, lngPrintH)
                    blnDo = lngPrintH > 0 And lngPrintH < objFmt.H
                    If blnDo Then '�հײ��ݸ���30mm�Ҹ���ԭֽ�ŵ�1/8
                        blnDo = objFmt.H - lngPrintH > 30 * Twip_mm And objFmt.H - lngPrintH > objFmt.H / 8
                    End If
                    If blnDo Then
                        lngPrintH = lngPrintH + 567 '��ʵ�ʴ�ӡ����10mm�߶�
                        If Not SetPrinterPaper(Me.hwnd, mobjReport, lngPrintH, intCopy) Then
                            '����ʧ��ʱ�ָ���ԭʼֽ��
                            Call ResetPrinterPaper(Me.hwnd, mobjReport, intCopy)
                        End If
                    End If
                End If
                
                blnPrint = True
                Call PrintPage(0, Printer, Me)
            Else
makPage:
                If IsArray(marrPage) Then
                    lngEndPage = UBound(marrPage)
                ElseIf IsArray(marrPageCard) Then
                    lngEndPage = UBound(marrPageCard)
                Else
                    lngEndPage = -1
                End If
                
                For i = 0 To lngEndPage
                    If mbytStyle <> 2 Then
                        If Printer.Copies <> intCopy And intCopy <> 1 Then
                            ShowFlash "���" & mobjReport.���� & ",�� " & lngEndPage + 1 & " ҳ " & intCopy & " ��,��ǰ�� " & j & " ��", ((i + 1) + ((j - 1) * (lngEndPage + 1))) / ((lngEndPage + 1) * intCopy), Me
                        Else
                            ShowFlash "���" & mobjReport.���� & ",�� " & lngEndPage + 1 & " ҳ,��ǰ�� " & i + 1 & " ҳ��", (i + 1) / (lngEndPage + 1), Me
                        End If
                    End If
                    
                    '��̬���㼰����ֽ�Ÿ߶�
                    If objFmt.��ֽ̬�� And objFmt.ֽ�� = 1 Then
                        Call PrintPage(i, Me, Me, 1, False, True, lngPrintH)
                        blnDo = lngPrintH > 0 And lngPrintH < objFmt.H
                        If blnDo Then '�հײ��ݸ���30mm�Ҹ���ԭֽ�ŵ�1/8
                            blnDo = objFmt.H - lngPrintH > 30 * Twip_mm And objFmt.H - lngPrintH > objFmt.H / 8
                        End If
                        If blnDo Then
                            lngPrintH = lngPrintH + 567 '��ʵ�ʴ�ӡ����10mm�߶�
                            If Not SetPrinterPaper(Me.hwnd, mobjReport, lngPrintH, intCopy) Then
                                '����ʧ��ʱ�ָ���ԭʼֽ��
                                Call ResetPrinterPaper(Me.hwnd, mobjReport, intCopy)
                                blnReset = False
                            Else
                                blnReset = True '��ҳ�����ù���ֽ̬��,��ҳ�����������ʱҪ�ָ���ԭʼ��
                            End If
                        ElseIf blnReset Then
                            Call ResetPrinterPaper(Me.hwnd, mobjReport, intCopy)
                            blnReset = False
                        End If
                    End If
                    
                    blnPrint = True
                    If Not PrintPage(i, Printer, Me) Then Exit For
                    If i <> lngEndPage Then Printer.NewPage: blnPrint = True '��ҳ
                Next
            End If
            If k > 0 Then Printer.NewPage: blnPrint = True '���
        Loop Until k = 0
        
NextFormat:
        '����Ƿ������ӡ��һ��ʽ
        blnGoOn = False
        If InStr(";0;2;4;", ";" & mbytStyle & ";") > 0 And mblnAllFormat And cboFormat.ComboItems.count > 1 And cboFormat.SelectedItem.Index < cboFormat.ComboItems.count Then
            cboFormat.ComboItems(cboFormat.SelectedItem.Index + 1).Selected = True
            Call CboFormat_Click: blnGoOn = True
            If Not (mblnPrintEmpty = False And mobjReport.��ӡ��ʽ = 1) Or blnExit = False Then
                Printer.NewPage
            End If
            blnPrint = True '���ʽ�������ҳ��ʵ�����,�򲻻�����´�ӡҳ
        End If
    Loop
    cboFormat.Tag = ""

    If Not mobjCurDLL Is Nothing Then
        mobjCurDLL.DataIsEmpty = blnALLEmpty
    End If

ExitHandle:
    cboFormat.Tag = ""
    If blnPrint Then
        If gblnSingleTask Then
            Printer.NewPage '�����ҳ��ʵ�����,�򲻻�����´�ӡҳ
        Else
            Printer.EndDoc
        End If
        
        '���PDF����
        If mbytStyle = Val("4-PDF") Then
            Call PDFFileSuccess
        End If
        
        '�������ӡ�����������ӡ�¼�
        If Not mobjCurDLL Is Nothing Then
            Call mobjCurDLL.Act_AfterPrint(mobjReport.���)
        End If
    End If

    If mbytStyle <> 2 Then ShowFlash
    SetRedraw True
    Screen.MousePointer = 0
    Exit Sub
    
errH:
    cboFormat.Tag = ""
    Screen.MousePointer = 0
    If mbytStyle <> 2 Then Call ShowFlash
    Printer.KillDoc
    SetRedraw True
    MsgBox Err.Number & ":" & Err.Description & vbCrLf & "��ӡ���̱�ǿ���жϣ�", vbExclamation, App.Title
    Err.Clear
    gblnError = True
End Sub

Private Sub mnuHelpTitle_Click()
    If Me.Tag = "" Then
        Call ShowHelpRpt(Me.hwnd, mobjReport.���, Int((mobjReport.ϵͳ) / 100))
    Else
        Call ShowHelpRpt(Me.hwnd, Me.Tag, Int((mobjReport.ϵͳ) / 100))
    End If
End Sub

Private Sub mnuPop_Cond_Click(Index As Integer)
    Set mobjPars = mdlPublic.RPTParsCondExec(mlngRPTID, Val(mnuPop_Cond(Index).Tag), mobjDefPars)
    If Not mobjPars Is Nothing Then
        mintCurMenuIndex = Index
        mintCurCondID = Val(mnuPop_Cond(Index).Tag)
        Call InitReportPars
        If cmdLoad.Enabled And cmdLoad.Visible Then cmdLoad.SetFocus
    End If
End Sub

Private Sub mnuPop_Default_Click()
    '����ȱʡ�������󣬸��µ�ǰ��������
    Call CopyPars(mobjDefPars, mobjPars)
    If Not mobjPars Is Nothing Then
        mintCurMenuIndex = 0
        mintCurCondID = 0
        Call InitReportPars
        If cmdLoad.Enabled And cmdLoad.Visible Then cmdLoad.SetFocus
    End If
End Sub

Private Sub mnuPop_Del_Click()
    If mdlPublic.RPTParsCondDel(mlngRPTID, mintCurCondID) Then
        Call mnuPop_Default_Click
    End If
End Sub

Private Sub mnuPop_Save_Click()
    '��������
    If mdlPublic.RPTParsCondSave(mlngRPTID, mintCurCondID, mobjPars, mobjDefPars, Me) Then
        '���²����ؼ�
        If mintCurCondID = 0 Then
            '��ȱʡ״̬�±��棬����Ϊ����������
            Call mnuPop_Cond_Click(mnuPop_Cond.count - 1)
        Else
            '������״̬�±���
            Call mnuPop_Cond_Click(mintCurCondID)
        End If
    End If
End Sub

Private Sub mnuPop_SaveAs_Click()
    If mdlPublic.RPTParsCondSave(mlngRPTID, mintCurCondID, mobjPars, mobjDefPars, Me, True) Then
        '���²����ؼ�
        If mintCurCondID = 0 Then
            '��ȱʡ״̬�±��棬����Ϊ����������
            Call mnuPop_Cond_Click(mnuPop_Cond.count - 1)
        Else
            '������״̬�±���
            Call mnuPop_Cond_Click(mintCurCondID)
        End If
    End If
End Sub

Private Sub mnuView_Next_Click()
    Dim intIdx As Integer
    
    If lvw.SelectedItem Is Nothing Then Exit Sub
    intIdx = lvw.SelectedItem.Index
    If intIdx + 1 <= lvw.ListItems.count Then
        lvw.ListItems(intIdx + 1).Selected = True
        lvw.SelectedItem.EnsureVisible
        Call lvw_ItemClick(lvw.SelectedItem)
    End If
End Sub

Private Sub mnuView_Pre_Click()
    Dim intIdx As Integer
    
    If lvw.SelectedItem Is Nothing Then Exit Sub
    intIdx = lvw.SelectedItem.Index
    If intIdx - 1 >= 1 Then
        lvw.ListItems(intIdx - 1).Selected = True
        lvw.SelectedItem.EnsureVisible
        Call lvw_ItemClick(lvw.SelectedItem)
    End If
End Sub

Private Sub mnuView_reFlash_Click()
    Dim strErr As String, strCond As String
    Dim tmpMsh As Object
    Dim strStartTime As String, strInfo As String, strName As String
    Dim intState As Integer
    
    If gblnReportRunLog Then
        strStartTime = Format(Currentdate, "YYYY-MM-DD HH:mm:SS")
    End If
    'ȡ��ǰ����
    If lblName.UBound > 0 Then
        If Not ReSetReportPars Then Exit Sub
    End If
    
    '���������ύ�¼�
    If Not mobjCurDLL Is Nothing Then
        '��鱨��򱨱����״̬
        strName = mdlPublic.FormatString("��[1]��[2]", mobjReport.���, mobjReport.����)
        intState = mdlPublic.ReportStateSwitch(mlngSys, mobjReport.���, False, strInfo)
        Select Case intState
        Case Val("0-�������ڻ�δ����")
            If strInfo = "" Then
                MsgBox mdlPublic.FormatString("����[1]�����ڣ�����ϵ����Ա��", strName), vbInformation, App.Title
                Exit Sub
'            Else
'                MsgBox mdlPublic.FormatString("��[1]������δ����������ϵ����Ա��", strName), vbInformation, App.Title
            End If
        Case Val("1-����������")
            '����
        Case Val("2-����ͣ����")
            MsgBox mdlPublic.FormatString("��[1]������ͣ���У�����ϵ����Ա��", strName), vbInformation, App.Title
            Exit Sub
        Case Else
            Exit Sub
        End Select
        
        '�׳��¼�
        strCond = GetParsStr(MakeNamePars(mobjReport, True))
        mobjCurDLL.Act_CommitCondition mobjReport.���, strCond, Me
    End If
    
     '��ѯ�Ƿ������ڵ�ǰʱ��ִ�д˱���
    If Format(mobjReport.��ֹ����ʱ��, "HH:mm:ss") <> "00:00:00" Or Format(mobjReport.��ֹ��ʼʱ��, "HH:mm:ss") <> "00:00:00" Then
        If CDate(Format(mobjReport.��ֹ����ʱ��, "HH:mm:ss")) > CDate(Format(mobjReport.��ֹ��ʼʱ��, "HH:mm:ss")) Then
            If Between(CDate(Format(Currentdate, "HH:mm:ss")), CDate(Format(mobjReport.��ֹ��ʼʱ��, "HH:mm:ss")), CDate(Format(mobjReport.��ֹ����ʱ��, "HH:mm:ss"))) Then
                MsgBox "��ǰ������" & CDate(Format(mobjReport.��ֹ��ʼʱ��, "HH:mm:ss")) & "-" & CDate(Format(mobjReport.��ֹ����ʱ��, "HH:mm:ss")) & "��ִֹ�У�������������ϵ��Ϣ�ơ�", vbInformation, App.Title
                Exit Sub
            End If
        Else
            If CDate(Format(Currentdate, "HH:mm:ss")) < CDate(Format(mobjReport.��ֹ��ʼʱ��, "HH:mm:ss")) Or CDate(Format(Currentdate, "HH:mm:ss")) > CDate(Format(mobjReport.��ֹ����ʱ��, "HH:mm:ss")) Then
                MsgBox "��ǰ������" & CDate(Format(mobjReport.��ֹ��ʼʱ��, "HH:mm:ss")) & "-�ڶ���" & CDate(Format(mobjReport.��ֹ����ʱ��, "HH:mm:ss")) & "��ִֹ�У�������������ϵ��Ϣ�ơ�", vbInformation, App.Title
                Exit Sub
            End If
        End If
    End If
    
    '��������
    strErr = OpenReportData(True)
    If strErr <> "" Then
        MsgBox "�ڶ�ȡ��������""" & strErr & """ʱ�����������,�����ܲ�����", vbInformation, App.Title
        Exit Sub
    End If
    '��������
    Call ShowItems
    
    If lblName.UBound > 0 Then
        '������ʾ����(ע�⣺��Ҫ������ʾ)
        picPar.Visible = False
        Set mobjPars = New RPTPars
        Set mobjPars = MakeNamePars(mobjReport)
        Call InitReportPars
        picPar.Visible = True
        
        '���ݵ�ǰ����,�滻������������ͬ������ֵ
        Call KeepParsSame
    End If
    
    '��λ�ڵ�һ�������
    For Each tmpMsh In msh
        If tmpMsh.Index <> 0 And tmpMsh.Container Is picPaper(intReport) And Not tmpMsh.Tag Like "H_*" Then
            Call msh_EnterCell(tmpMsh.Index)
            On Error Resume Next
            tmpMsh.SetFocus
            On Error GoTo 0
            Exit For
        End If
    Next
    If lvw.ListItems.count > 0 Then
        Call RecordsExecute(Val(Mid(lvw.SelectedItem.Key, 2)), strStartTime, 2)
    Else
        Call RecordsExecute(mlngReportID, strStartTime, 2)
    End If
End Sub

Private Sub KeepParsSame()
'���ܣ��������У����ݵ�ǰ������Ч�Ĳ�������,��������������ͬ������ֵ��ͬ
'���룺mobjPars=��ǰ������ʹ�õ��Ĳ���
'˵����1.�����͵Ĳ���������,��Ϊ���ܲ���ֵ��д��ʽ��һ��
    Dim objPar As RPTPar, tmpPar As RPTPar
    Dim objData As RPTData, i As Integer
    For i = 0 To UBound(arrReport)
        If i <> intReport Then
            For Each objData In arrReport(i).Datas
                For Each objPar In objData.Pars
                    For Each tmpPar In mobjPars
                        If tmpPar.���� = objPar.���� _
                            And tmpPar.���� = objPar.���� _
                            And objPar.���� <> 3 Then
                            objPar.ȱʡֵ = tmpPar.ȱʡֵ
                            objPar.Reserve = tmpPar.Reserve
                        End If
                    Next
                Next
            Next
        End If
    Next
End Sub

Private Sub mnuViewToolFormat_Click()
    mnuViewToolFormat.Checked = Not mnuViewToolFormat.Checked
    If mnuViewToolButton.Checked Then cbr.Bands(1).Visible = False
    cbr.Bands(2).Visible = Not cbr.Bands(2).Visible
    If mnuViewToolButton.Checked Then cbr.Bands(1).Visible = True
    cbr.Visible = cbr.Bands(2).Visible Or cbr.Bands(1).Visible
    Form_Resize
End Sub

Private Sub mnuViewToolGroup_Click()
    mnuViewToolGroup.Checked = Not mnuViewToolGroup.Checked
    picLR_S.Visible = Not picLR_S.Visible
    picGroup.Visible = Not picGroup.Visible
    Call Form_Resize
End Sub

Private Sub msh_AfterSort(Index As Integer, ByVal Col As Long, Order As Integer)
    Dim objRelations As RPTRelations
    Dim i As Long
    Dim intIdx As Integer
    Dim colRelations As Collection
    
    If mobjReport.Items("_" & Index).���� = 6 Then
        intIdx = Val(Mid(msh(Index).Tag, 3))
    
        '����Sort����ֵ�����е����ӱ�����Ϣ�ᷢ���仯����ˣ���Ҫ�Ȼ�������������
        For i = 0 To msh(intIdx).Cols - 1
            If TypeName(msh(intIdx).Cell(flexcpData, 0, i)) = "RPTRelations" Then
                '���󻺴�
                If colRelations Is Nothing Then
                    Set colRelations = New Collection
                End If
                colRelations.Add msh(intIdx).Cell(flexcpData, 0, i), "_" & i
                '���ԭ��¼Data��Ϣ
                msh(intIdx).Cell(flexcpData, 0, i) = Empty
            End If
        Next
    
        msh(intIdx).Col = Col
        msh(intIdx).Sort = Order
        
        '����ָ�
        If Not colRelations Is Nothing Then
            For i = 0 To msh(intIdx).Cols - 1
                Set objRelations = Nothing
                On Error Resume Next
                Set objRelations = colRelations("_" & i)
                On Error GoTo 0
                If Not objRelations Is Nothing Then
                    Set msh(intIdx).Cell(flexcpData, 0, i) = objRelations
                End If
            Next
        End If
    End If
End Sub

Private Sub msh_AfterUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim i As Integer, intBegin As Integer, intEnd As Integer
    Dim lngID As Long
    Dim objRelaID As RelatID, objItem As RPTItem, objBody As RPTItem
    Dim sngNew As Single, sngOld As Single
    
    If mobjReport.Items("_" & Index).���� = Val("6-�������ͷ��") Then
        If msh(Index).Tag Like "H_*" Then
            '�����ж���(RPTItem)���
            lngID = Val(Mid(msh(Index).Tag, 3))
            Set objBody = mobjReport.Items("_" & lngID)
            If Not objBody Is Nothing Then
                intBegin = -1
                intEnd = -1
                For Each objRelaID In objBody.SubIDs
                    Set objItem = mobjReport.Items("_" & objRelaID.id)
                    If objItem.��� = Col Then
                        sngOld = objItem.W
                        objItem.W = msh(Index).ColWidth(Col)
                        sngNew = msh(Index).ColWidth(Col)
                    End If
                    If objItem.����Ӧ�и� Then
                        If intBegin < 0 Then intBegin = objItem.���
                        intEnd = objItem.���
                    End If
                Next
            End If
            '����������
            msh(lngID).ColWidth(Col) = msh(Index).ColWidth(Col)
            '�����и�
            If intBegin >= 0 Then
                msh(lngID).AutoSize intBegin, intEnd
            End If
        End If
        If Not mobjCurDLL Is Nothing Then
            Call mobjCurDLL.Act_ColResize(mobjReport.���, CInt(Col), sngNew, sngOld)
        End If
    End If
End Sub

Private Sub msh_Click(Index As Integer)
    Dim objRelations As RPTRelations
    Dim i As Long
    Dim lngRec As Long, strDataName As String
    Dim strLisName As String
    Dim tmpData As RPTData, tmpPar As RPTPar
    Dim tmpID As RelatID, tmpItem As RPTItem
    Dim strFilter As String, lngItemID As Long
    Dim lngMouseRow As Long, lngMouseCol As Long
    Dim lngRelationReport As Long
    
    If mblnLeftClick = False And mlngRelationReport = 0 Then Exit Sub
    If mblnLeftClick = False And mlngRelationReport <> 0 Then
        lngMouseRow = mlngRelationMouseRow
        lngMouseCol = mlngRelationMouseCol
    Else
        lngMouseRow = msh(Index).MouseRow: lngMouseCol = msh(Index).MouseCol
    End If

    If grsObject Is Nothing Then Set grsObject = UserObject
    If grsObject Is Nothing Then Exit Sub
    If grsObject.State = adStateClosed Then
        Set grsObject = Nothing
        Set grsObject = UserObject
        If grsObject Is Nothing Then Exit Sub
    End If
    
    If lngMouseRow > -1 And lngMouseCol > -1 Then
        If msh(Index).Cell(flexcpFontUnderline, lngMouseRow, lngMouseCol) = True Then
            If mobjReport.Items("_" & Index).���� = 4 Then
'                Set objRelations = msh(Index).Cell(flexcpData, lngMouseRow, lngMouseCol)(2)
'                '�����ɶ�λ�����������
'                lngRec = msh(Index).Cell(flexcpData, lngMouseRow, lngMouseCol)(1)
                
                '�Ż�
                Set objRelations = msh(Index).Cell(flexcpData, 0, lngMouseCol)  '�ӵ�һ�еĵ�Ԫ���л�ȡRelations����
                lngRec = msh(Index).RowData(lngMouseRow)                        '���е�RowData�л�ȡ��¼���к�
            Else
                '�Ż���ֻ���ض��а����������
                lngRec = msh(Index).FixedRows
                If TypeName(msh(Index).Cell(flexcpData, lngRec, lngMouseCol)) = "Empty" Then Exit Sub
                If msh(Index).Cell(flexcpData, lngRec, lngMouseCol).Relations.count <= 0 Then Exit Sub
                
                Set objRelations = msh(Index).Cell(flexcpData, lngRec, lngMouseCol).Relations
                lngItemID = msh(Index).Cell(flexcpData, lngRec, lngMouseCol).id
            End If
            
            For i = 1 To objRelations.count
                If objRelations.Item(i).Ĭ�� = 1 Then
                    lngRelationReport = objRelations.Item(i).��������ID
                    Exit For
                End If
            Next
            If lngRelationReport = 0 Then lngRelationReport = objRelations.Item(1).��������ID
            If mlngRelationReport <> 0 Then lngRelationReport = mlngRelationReport
            If Not CheckReportPriv(lngRelationReport) Then
                MsgBox "��û��Ȩ�޲�ѯ�ñ���ĳЩ����Դ�еĶ���", vbInformation, App.Title: Exit Sub
            End If
            'ִ�б���
            If CheckPass(lngRelationReport) = False Then
                MsgBox "�������ݴ��󣬲���ִ�иñ���", vbInformation, App.Title: Exit Sub
            End If
            
            Set gobjReport = ReadReport(lngRelationReport)
            '��ʼ������
            garrPars = Array()
            '��λ��¼��
            On Error Resume Next
            For i = 1 To objRelations.count
                If objRelations.Item(i).��������ID = lngRelationReport Then
                    If InStr(objRelations.Item(i).����ֵ��Դ, ".") > 0 Then
                        strDataName = Mid(objRelations.Item(i).����ֵ��Դ, 1, InStr(objRelations.Item(i).����ֵ��Դ, ".") - 1)
                    End If
                End If
                If strDataName <> "" Then Exit For
            Next
            If mobjReport.Items("_" & Index).���� = 4 Then
                '������ܹ�ȷ������ľ���������
                If strDataName <> "" Then mLibDatas("_" & strDataName).DataSet.AbsolutePosition = lngRec
            Else
                '���ܱ�ֻ�ܸ�������ͺ�����ඨλ������
                If strDataName <> "" Then
                    For Each tmpID In mobjReport.Items("_" & Index).SubIDs
                        Set tmpItem = mobjReport.Items("_" & tmpID.id)
                        Select Case mobjReport.Items("_" & lngItemID).����
                            Case 7 '�������
                                If tmpItem.���� = 7 Then
                                    If Decode(Trim(msh(Index).TextMatrix(lngMouseRow, tmpItem.���)), "�ϼ�", 1, "ƽ��ֵ", 2, "���ֵ", 3, "��Сֵ", 4, "��¼��", 5, 0) > 0 Then
                                        '����Ǻϼ�����ȡ����һ��
                                        lngMouseRow = lngMouseRow - 1
                                    End If
                                    strFilter = strFilter & " And " & tmpItem.���� & "='" & msh(Index).TextMatrix(lngMouseRow, tmpItem.���) & "'"
                                End If
                            Case 8 '�������
                                If tmpItem.���� = 8 Then
                                    strFilter = strFilter & " And " & tmpItem.���� & "='" & msh(Index).TextMatrix(tmpItem.���, lngMouseCol) & "'"
                                End If
                            Case 9 'ͳ����
                                'ͳ������ݺ��������������ȷ��
                                If tmpItem.���� = 7 Then
                                    If Decode(Trim(msh(Index).TextMatrix(lngMouseRow, tmpItem.���)), "�ϼ�", 1, "ƽ��ֵ", 2, "���ֵ", 3, "��Сֵ", 4, "��¼��", 5, 0) > 0 Then
                                        '����Ǻϼ�����ȡ����һ��
                                        lngMouseRow = lngMouseRow - 1
                                    End If
                                    strFilter = strFilter & " And " & tmpItem.���� & "='" & msh(Index).TextMatrix(lngMouseRow, tmpItem.���) & "'"
                                ElseIf tmpItem.���� = 8 Then
                                    strFilter = strFilter & " And " & tmpItem.���� & "='" & msh(Index).TextMatrix(tmpItem.���, lngMouseCol) & "'"
                                End If
                        End Select
                    Next
                    mLibDatas("_" & strDataName).DataSet.Filter = Mid(strFilter, 6)
                End If
            End If
            For i = 1 To objRelations.count
                With objRelations.Item(i)
                    strLisName = ""
                    If objRelations.Item(i).��������ID = lngRelationReport Then
                        If InStr(.����ֵ��Դ, ".") > 0 Then
                            If mLibDatas("_" & strDataName).DataSet.RecordCount > 0 Then
                                strLisName = mLibDatas("_" & strDataName).DataSet.Fields(Mid(.����ֵ��Դ, InStr(.����ֵ��Դ, ".") + 1)).Value
                            End If
                        ElseIf InStr(.����ֵ��Դ, "=") = 1 Then
                            For Each tmpData In mobjReport.Datas
                                For Each tmpPar In tmpData.Pars
                                    If tmpPar.���� = Mid(.����ֵ��Դ, 2) Then
                                        strLisName = tmpPar.ȱʡֵ
                                        Exit For
                                    End If
                                Next
                                If strLisName <> "" Then Exit For
                            Next
                        End If
                    
                        ReDim Preserve garrPars(UBound(garrPars) + 1)
                        garrPars(UBound(garrPars)) = .������ & "=" & strLisName
                    End If
                End With
            Next
            
            If Not ShowReport(Me) Then MsgBox "�����ʧ�ܣ�", vbInformation, App.Title
        End If
    End If
End Sub

Private Sub msh_DblClick(Index As Integer)
    If Not mobjCurDLL Is Nothing And IsNumeric(msh(Index).Tag) Then
        Call mobjCurDLL.Act_SheetDblClick(mobjReport.���, msh(Index), Me)
        msh(Index).SetFocus
    End If
End Sub

Private Sub msh_EnterCell(Index As Integer)
    Dim i As Long, lngRow As Long, lngCol As Long
    Dim strRowText As String, strText As String
    Dim intA As Integer, intB As Integer
    Static strRow As String
    Static strCol As String
    
    If blnRefresh = False Then Exit Sub
    Set objCurGrid = msh(Index)
    
    If Not mobjCurDLL Is Nothing And IsNumeric(msh(Index).Tag) Then
        lngRow = msh(Index).Row
        lngCol = msh(Index).Col
        intA = IIF(lngRow > 30000, 30000, lngRow)
        intB = IIF(lngCol > 30000, 30000, lngCol)
        strText = msh(Index).Text
        Call mobjCurDLL.Act_EnterCell(mobjReport.���, intA, intB, strText)
        '�ı���ֵ
        If lngRow >= 0 And lngRow <= msh(Index).Rows - 1 And lngCol >= 0 And lngCol <= msh(Index).Cols - 1 Then
            msh(Index).Row = lngRow
            msh(Index).Col = lngCol
            msh(Index).Text = strText
        End If
        
        If strRow <> Index & "," & msh(Index).Row Then
            For i = 0 To msh(Index).Cols - 1
                strRowText = strRowText & "|" & msh(Index).TextMatrix(msh(Index).Row, i)
            Next
            intA = IIF(msh(Index).Row > 30000, 30000, msh(Index).Row)
            Call mobjCurDLL.Act_EnterRow(mobjReport.���, intA, Mid(strRowText, 2), msh(Index))
            strRow = Index & "," & msh(Index).Row
        End If
        
        If strCol <> Index & "," & msh(Index).Col Then
            Call mobjCurDLL.Act_EnterCol(mobjReport.���, msh(Index).Col, msh(Index))
            strCol = Index & "," & msh(Index).Col
        End If
    End If
End Sub

Private Sub msh_GotFocus(Index As Integer)
    On Error Resume Next
    If msh(Index).Tag Like "H_*" Then
        msh(CInt(Mid(msh(Index).Tag, 3))).SetFocus
    Else
        Call msh_EnterCell(Index)
    End If
End Sub

Private Sub msh_LeaveCell(Index As Integer)
    Dim intRow As Integer
    
    If blnRefresh = False Then Exit Sub
    If Not mobjCurDLL Is Nothing And IsNumeric(msh(Index).Tag) Then
        intRow = IIF(msh(Index).Row > 30000, 30000, msh(Index).Row)
        Call mobjCurDLL.Act_LevelCell(mobjReport.���, intRow, msh(Index).Col, msh(Index).Text)
    End If
End Sub

Private Sub msh_LostFocus(Index As Integer)
    If Not msh(Index).Tag Like "H_*" Then Call msh_LeaveCell(Index)
End Sub

Private Sub msh_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not mobjCurDLL Is Nothing And IsNumeric(msh(Index).Tag) Then
        Call mobjCurDLL.Act_SheetMouseDown(mobjReport.���, Button, Shift, X, Y, msh(Index), Me)
    End If
End Sub

Private Sub msh_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    msh(Index).ToolTipText = msh(Index).TextMatrix(msh(Index).MouseRow, msh(Index).MouseCol)
    
    If Not mobjCurDLL Is Nothing And IsNumeric(msh(Index).Tag) Then
        Call mobjCurDLL.Act_SheetMouseMove(mobjReport.���, Button, Shift, X, Y, msh(Index), Me)
    End If
    If msh(Index).MouseRow > -1 And msh(Index).MouseCol > -1 Then
        If msh(Index).Cell(flexcpFontUnderline, msh(Index).MouseRow, msh(Index).MouseCol) = True Then
            msh(Index).MousePointer = 99
        Else
            msh(Index).MousePointer = 0
        End If
    Else
        msh(Index).MousePointer = 0
    End If
End Sub

Private Sub msh_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    Dim dbRowHeight As Double
    
    If Not mobjCurDLL Is Nothing And IsNumeric(msh(Index).Tag) Then
        Call mobjCurDLL.Act_SheetMouseUp(mobjReport.���, Button, Shift, X, Y, msh(Index), Me)
    End If
    
    If Button = vbRightButton Then
        If msh(Index).MouseRow < 0 Or msh(Index).MouseCol < 0 Then
            vsfRelations.Visible = False
            Exit Sub
        End If
        mintGridIndex = Index
        mblnLeftClick = False
        mlngRelationMouseRow = msh(Index).MouseRow
        mlngRelationMouseCol = msh(Index).MouseCol
        If msh(Index).Cell(flexcpFontUnderline, msh(Index).MouseRow, msh(Index).MouseCol) Then
            Call LoadRelation(0, Index, msh(Index).MouseRow, msh(Index).MouseCol)
            If TypeName(msh(Index).Cell(flexcpData, msh(Index).MouseRow, msh(Index).MouseCol)) = "RPTRelations" Then
                vsfRelations.Visible = True
                vsfRelations.SetFocus
            End If
            
            For i = 0 To vsfRelations.Rows - 1
                dbRowHeight = dbRowHeight + vsfRelations.RowHeight(i)
            Next
            vsfRelations.Height = dbRowHeight
            vsfRelations.Left = msh(Index).Left + X + 150
            vsfRelations.Top = msh(Index).Top + Y + 90
        Else
            vsfRelations.Visible = False
        End If
    Else
        mblnLeftClick = True
    End If
End Sub

Private Sub LoadRelation(ByVal bytType As Byte, ByVal cIndex As Integer, Optional lngMouseRow As Long, Optional lngMouseCol As Long)
    Dim i As Long
    Dim objRelations As RPTRelations
    Dim strFlag As String
    
    mbytType = bytType
    If mbytType = 0 Then
        If TypeName(msh(cIndex).Cell(flexcpData, lngMouseRow, lngMouseCol)) <> "RPTRelations" Then
            Exit Sub
        End If
        If mobjReport.Items("_" & cIndex).���� = 4 Then
            Set objRelations = msh(cIndex).Cell(flexcpData, lngMouseRow, lngMouseCol)(2)
        Else
            Set objRelations = msh(cIndex).Cell(flexcpData, lngMouseRow, lngMouseCol).Relations
        End If
    ElseIf mbytType = 1 Then
        Set objRelations = mobjReport.Items("_" & cIndex).Relations
    End If
    If objRelations.count = 0 Then Exit Sub

    With vsfRelations
        .Rows = 0
        If .Cols = 0 Then
            .Cols = 2
            .ColKey(0) = "ID"
            .ColDataType(0) = flexDTString
            .ColWidth(0) = 0
            .ColKey(1) = "����"
            .ColDataType(1) = flexDTString
            .ColWidth(1) = vsfRelations.Width
        End If
        For i = 1 To objRelations.count
            If InStr(strFlag, "," & objRelations.Item(i).��������ID) = 0 Then
                strFlag = strFlag & "," & objRelations.Item(i).��������ID
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, .ColIndex("ID")) = objRelations.Item(i).��������ID
                .TextMatrix(.Rows - 1, .ColIndex("����")) = " " & Split(objRelations.Item(i).������������, "(")(0)
            End If
        Next
    End With
End Sub

Private Sub msh_AfterScroll(Index As Integer, ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Dim intPre As Integer
    
    If Not mobjCurDLL Is Nothing And IsNumeric(msh(Index).Tag) Then
        Call mobjCurDLL.Act_SheetScroll(mobjReport.���, msh(Index))
    End If
    
    If IsNumeric(msh(Index).Tag) Then
        intPre = msh(msh(Index).Tag).LeftCol
        msh(msh(Index).Tag).LeftCol = msh(Index).LeftCol
        If msh(msh(Index).Tag).LeftCol = intPre Then msh(Index).LeftCol = intPre
    ElseIf Left(msh(Index).Tag, 2) = "H_" Then
        intPre = msh(Mid(msh(Index).Tag, 3)).LeftCol
        msh(Mid(msh(Index).Tag, 3)).LeftCol = msh(Index).LeftCol
        If msh(Mid(msh(Index).Tag, 3)).LeftCol = intPre Then msh(Index).LeftCol = intPre
    End If
End Sub

Private Sub opt_GotFocus(Index As Integer)
    If opt(Index).Value Then
        '��������Ŀ���Ǳ��ⰴTAB��ʱ�Զ��л�����һ��ѡ��
        opt(Index).Value = False
        opt(Index).Value = True
    End If
End Sub

Private Sub picLR_S_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objTmp As Object
    
    On Error Resume Next
    
    If Button = 1 Then
        If picGroup.Width + X < 1000 Or picBack.Width - X < 3000 Then Exit Sub
        picLR_S.Left = picLR_S.Left + X

        picGroup.Width = picGroup.Width + X
        picBack.Left = picBack.Left + X
        picBack.Width = picBack.Width - X
        scrHsc.Left = scrHsc.Left + X
        scrHsc.Width = scrHsc.Width - X
        
        lblGroup_S.Width = lblGroup_S.Width + X
        lvw.Width = lvw.Width + X
        lblPar_S.Width = lblPar_S.Width + X
        picPar.Width = picPar.Width + X
        
        lvw.ColumnHeaders(1).Width = lvw.Width - 500    '��̬�п�
        
        For Each objTmp In fraGroup
            objTmp.Width = picGroup.ScaleWidth - objTmp.Left * 2
        Next
        For Each objTmp In fra
            objTmp.Width = picGroup.ScaleWidth - objTmp.Left * 2
        Next
        
        picPaper(intReport).Cls
        Call SetPaper
        Call SetPlace
        Me.Refresh
    End If
End Sub

Private Sub picPane_Click()

End Sub

Private Sub picPaper_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim blnPop As Boolean
    
    lngPreX = X: lngPreY = Y
    
    If Not mobjCurDLL Is Nothing Then
        blnPop = True
        Call mobjCurDLL.Act_PaperMouseDown(mobjReport.���, Button, Shift, X, Y, blnPop)
        If blnPop Then
            If Button = 2 Then PopupMenu mnuEdit, 2
        End If
    Else
        If Button = 2 Then PopupMenu mnuEdit, 2
    End If
End Sub

Private Sub picPaper_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not mobjCurDLL Is Nothing Then
        Call mobjCurDLL.Act_PaperMouseMove(mobjReport.���, Button, Shift, X, Y)
    End If
    If Button = 1 Then
        If scrVsc.Enabled And scrVsc.Visible Then
            If (Y - lngPreY) / 15 > 0 Then
                scrVsc.Value = IIF(scrVsc.Value - (Y - lngPreY) / 15 < scrVsc.Min, scrVsc.Min, scrVsc.Value - (Y - lngPreY) / 15)
            Else
                scrVsc.Value = IIF(scrVsc.Value - (Y - lngPreY) / 15 > scrVsc.Max, scrVsc.Max, scrVsc.Value - (Y - lngPreY) / 15)
            End If
        End If
        If scrHsc.Enabled And scrHsc.Visible Then
            If (X - lngPreX) / 15 > 0 Then
                scrHsc.Value = IIF(scrHsc.Value - (X - lngPreX) / 15 < scrHsc.Min, scrHsc.Min, scrHsc.Value - (X - lngPreX) / 15)
            Else
                scrHsc.Value = IIF(scrHsc.Value - (X - lngPreX) / 15 > scrHsc.Max, scrHsc.Max, scrHsc.Value - (X - lngPreX) / 15)
            End If
        End If
    End If
End Sub

Private Sub picPaper_GotFocus(Index As Integer)
    Oldwinproc = GetWindowLong(picPaper(Index).hwnd, GWL_WNDPROC)
    SetWindowLong picPaper(Index).hwnd, GWL_WNDPROC, AddressOf FlexScroll
End Sub

Private Sub picPaper_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
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

Private Sub picPaper_LostFocus(Index As Integer)
    SetWindowLong picPaper(Index).hwnd, GWL_WNDPROC, Oldwinproc
End Sub

Private Sub mnuEdit_Par_Click()
'���ܣ����ñ�������
    Dim strErr As String, objPars As RPTPars
    Dim strCond As String, blnInhere As Boolean
    Dim lngReport As Long, strSQL As String
    Dim rsReport As New ADODB.Recordset
    Dim frmNewParInput As New frmParInput
    Dim strStartTime As String
    
    'ȡ�ñ����ID
    lngReport = 0
    strSQL = "Select ID from zlReports Where ���=[1]"
    Set rsReport = OpenSQLRecord(strSQL, Me.Caption, mobjReport.���)
    If Not rsReport.EOF Then lngReport = rsReport!id
    
    If gblnReportRunLog Then
        strStartTime = Format(Currentdate, "YYYY-MM-DD HH:mm:SS")
    End If
    
    If Not mobjCurDLL Is Nothing Then
        blnInhere = True
        Set objPars = MakeNamePars(mobjReport, True)
        strCond = GetParsStr(objPars)
        
        '�������������¼�
        mobjCurDLL.Act_ResetCondition mobjReport.���, strCond, blnInhere, Me
        
        If Not blnInhere Then
             '���ó���ȡ����������,��������ʽ���ô���
            If strCond = "" Or Not strCond Like "*=*" Then Exit Sub
            
            Set objPars = SetStrPars(strCond, objPars)
            
            '���������ύ�¼�
            strCond = GetParsStr(objPars)
            mobjCurDLL.Act_CommitCondition mobjReport.���, strCond, Me
            
            ReplaceInputPars objPars
            
            Me.Refresh
            strErr = OpenReportData(True)
            If strErr <> "" Then MsgBox "�ڶ�ȡ��������""" & strErr & """ʱ�����������,�����ܲ�����", vbInformation, App.Title: Exit Sub
            Call ShowItems
        Else
            Set objPars = MakeNamePars(mobjReport) '��Ҫ���������ַ�ʽȡ
            
            frmNewParInput.mlngReport = lngReport
            Set frmNewParInput.mobjPars = objPars
            Set frmNewParInput.mobjDefPars = mobjDefPars
            Set frmNewParInput.mobjRPTDatas = mobjReport.Datas
            frmNewParInput.mstrTitle = mobjReport.����
            frmNewParInput.mblnReset = True
            frmNewParInput.Show 1, Me
            If frmNewParInput.mblnOK Then
                '���������ύ�¼�
                strCond = GetParsStr(frmNewParInput.mobjPars)
                mobjCurDLL.Act_CommitCondition mobjReport.���, strCond, Me
                
                ReplaceInputPars frmNewParInput.mobjPars
                Unload frmNewParInput
                
                '��������
                Me.Refresh
                strErr = OpenReportData(True)
                If strErr <> "" Then MsgBox "�ڶ�ȡ��������""" & strErr & """ʱ�����������,�����ܲ�����", vbInformation, App.Title: Exit Sub
                Call ShowItems
            End If
        End If
    Else
        frmNewParInput.mlngReport = lngReport
        Set frmNewParInput.mobjPars = MakeNamePars(mobjReport)
        Set frmNewParInput.mobjDefPars = mobjDefPars
        Set frmNewParInput.mobjRPTDatas = mobjReport.Datas
        frmNewParInput.mstrTitle = mobjReport.����
        frmNewParInput.mblnReset = True
        frmNewParInput.Show 1, Me
        If frmNewParInput.mblnOK Then
            ReplaceInputPars frmNewParInput.mobjPars
            Unload frmNewParInput
           
            '��������
            Me.Refresh
            strErr = OpenReportData(True)
            If strErr <> "" Then MsgBox "�ڶ�ȡ��������""" & strErr & """ʱ�����������,�����ܲ�����", vbInformation, App.Title: Exit Sub
            Call ShowItems
        End If
    End If
    Call RecordsExecute(lngReport, strStartTime, 2)
End Sub

Private Sub Form_Resize()
    Dim cbrH As Long '������ռ�ø߶�
    Dim staH As Long '״̬��ռ�ø߶�
    Dim lngTmp As Long
    
    On Error Resume Next
    
    If WindowState = 1 Then Exit Sub
    
    'lblPar_S��lblGropu_S��picLR_S ��������ʽ��Ϊ�˴�����书�ܵĴ���
    
    '����ؼ���Ⱥ͸߶�
    cbrH = IIF(cbr.Visible, cbr.Height, 0)
    staH = IIF(sta.Visible, sta.Height, 0)
    
    lblGroup_S.Width = picGroup.ScaleWidth - lblGroup_S.Left * 2
    
    lblPar_S.Width = lblGroup_S.Width
    
    lvw.Top = lblGroup_S.Top + lblGroup_S.Height + 15
    lvw.Width = picGroup.ScaleWidth
    lvw.Height = lblPar_S.Top - lblGroup_S.Top - lblGroup_S.Height - 15 * 2
    
    picPar.Top = lblPar_S.Top + lblPar_S.Height + 15
    picPar.Left = 0
    picPar.Width = lvw.Width
    picPar.Height = ScaleHeight - staH - cbrH - (lblGroup_S.Height + 30) - (lblPar_S.Height + 30) - lvw.Height
    
    picBack.Top = ScaleTop + cbrH
    picBack.Left = ScaleLeft + IIF(picGroup.Visible, picGroup.Width + picLR_S.Width, 0)
    picBack.Width = ScaleWidth - IIF(scrVsc.Visible, scrVsc.Width, 0) - IIF(picGroup.Visible, picGroup.Width + picLR_S.Width, 0)
    picBack.Height = ScaleHeight - staH - cbrH - IIF(scrHsc.Visible, scrHsc.Height, 0)
    
    If scrVsc.Visible Then
        scrVsc.Top = picBack.Top
        scrVsc.Left = ScaleWidth - scrVsc.Width
        scrVsc.Height = picBack.Height
        
        scrHsc.Left = picBack.Left
        scrHsc.Top = picBack.Top + picBack.Height
        scrHsc.Width = picBack.Width
    End If
    
    On Error GoTo 0
    
    If Not mobjReport Is Nothing And Visible Then
        picPaper(intReport).Cls
        Call SetPaper
        Call SetPlace
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer, bytMode As Byte
    
    bytMode = IIF(mnuEdit_SelMode_Row.Checked, 0, 1)
    
    If lvw.ListItems.count > 0 Then
        SaveWinState Me, App.ProductName, Me.Tag
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.name & Me.Tag, "ѡ��ģʽ", bytMode
        For i = 0 To UBound(arrReport)
            SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.name & arrReport(i).���, "��ʽ", arrReport(i).bytFormat
        Next
    ElseIf Not mobjReport Is Nothing Then
        If mbytStyle = 0 Then '����ʾ����ʱ�Ͳ������Լӿ��ٶ�
            SaveWinState Me, App.ProductName, mobjReport.���
        End If
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.name & mobjReport.���, "ѡ��ģʽ", bytMode
    End If
    
    If Not mobjReport Is Nothing Then
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.name & mobjReport.���, "��ʽ", bytFormat
    End If
    
    '�����ж���¼�
    If Not mobjCurDLL Is Nothing And Not mobjReport Is Nothing Then
        Call mobjCurDLL.Act_ReportUnload(mobjReport.���, Me)
    End If
    
    '�ͷ�ģ�����
    '---------------------------------------------------
    mbytStyle = 0
    mstrExcelFile = ""
    mstrPDFFile = ""
    
    Unload frmFlash
    
    Set frmParent = Nothing
    Set mobjCurDLL = Nothing
    Set mobjReport = Nothing
    Set mLibDatas = Nothing
    Set objCurGrid = Nothing
    Set mobjPars = Nothing
    Set mobjDefPars = Nothing
    Set objScript = Nothing
    
    Erase arrReport, arrLibDatas, arrDefPars

    If IsArray(marrPars) Then Erase marrPars
    If IsArray(marrPage) Then Erase marrPage
    marrPars = Empty
    marrPage = Empty

    Err.Clear
End Sub

Private Sub mnuFile_Quit_Click()
    Unload Me
End Sub

Private Sub mnuFile_Setup_Click()
    Dim objFmt As RPTFmt
    Dim strTmp As String
    Dim strDefault As String
    
    Set objFmt = mobjReport.Fmts("_" & mobjReport.bytFormat)
    strTmp = GetRegPrinterInfo("Printer", mobjReport.���, objFmt.˵��, mobjReport)
    If Not ReportLocalSet(mobjReport.ϵͳ, mobjReport.���, False, mobjReport.bytFormat, Me) Then Exit Sub
    sta.Panels(2) = "��ӡ��:" & strTmp & _
        "   ֽ��:" & GetPaperName(objFmt.ֽ��, objFmt.W, objFmt.H) & " " & _
        IIF(objFmt.ֽ�� = 256, CInt(objFmt.W / Twip_mm) & "mm �� " & CInt(objFmt.H / Twip_mm) & "mm", "") & _
        IIF(objFmt.ֽ�� = 1, "   ����", "   ����")
End Sub

Private Sub mnuHelpAbout_Click()
    Call ShowAbout(Me)
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    sta.Visible = Not sta.Visible
    Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    cbr.Bands(1).Visible = Not cbr.Bands(1).Visible
    mnuViewToolText.Enabled = Not mnuViewToolText.Enabled
    cbr.Visible = cbr.Bands(2).Visible Or cbr.Bands(1).Visible
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim i As Integer
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For i = 1 To tbr.Buttons.count
        tbr.Buttons(i).Caption = IIF(mnuViewToolText.Checked, tbr.Buttons(i).Tag, "")
    Next
    cbr.Bands(1).MinHeight = tbr.ButtonHeight
    Form_Resize
End Sub

Private Sub picPaper_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not mobjCurDLL Is Nothing Then
        Call mobjCurDLL.Act_PaperMouseUp(mobjReport.���, Button, Shift, X, Y)
    End If
End Sub

Private Sub tbr_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Quit"
            mnuFile_Quit_Click
        Case "Par"
            If lvw.ListItems.count = 0 Then
                mnuEdit_Par_Click '��������ʱ��������
            Else
                mnuView_reFlash_Click
            End If
        Case "Preview"
            mnuFile_Preview_Click
        Case "Print"
            mnuFile_Print_Click
        Case "Graph"
            mnuFile_Graph_Click
        Case "Help"
            mnuHelpTitle_Click
        Case "Style"
            Call SetView((lvw.View + 1) Mod 4)
        Case "Pre"
            mnuView_Pre_Click
        Case "Next"
            mnuView_Next_Click
        Case "ColWidth"
            mnuEdit_SetCol_Auto_Click
        Case "SelMode"
            If mnuEdit_SelMode_Row.Checked Then
                mnuEdit_SelMode_Col_Click
            Else
                mnuEdit_SelMode_Row_Click
            End If
    End Select
End Sub

Private Sub SetPaper()
'���ܣ����ñ���ֽ�ųߴ�,λ��,����
'˵�����������ڴ�ӡ�豸
    Dim strPrinter As String
    Dim strDefault As String
    
    strDefault = mobjReport.Fmts(mobjReport.bytFormat).˵��
    strPrinter = GetRegPrinterInfo("Printer", mobjReport.���, strDefault, mobjReport)
    With mobjReport.Fmts("_" & mobjReport.bytFormat)
        sta.Panels(2).Text = "��ӡ��:" & strPrinter & "   ֽ��:" & GetPaperName(.ֽ��, .W, .H) & " " & _
            IIF(.ֽ�� = 256, CInt(.W / Twip_mm) & "mm �� " & CInt(.H / Twip_mm) & "mm", "") & _
            IIF(.ֽ�� = 1, "   ����", "   ����")
    End With
    On Error GoTo errH
    
    If intGridCount = 1 And Not mobjReport.Ʊ�� Then
        picPaper(intReport).Top = 45
        picPaper(intReport).Left = 45
        picPaper(intReport).Width = picBack.ScaleWidth - picPaper(intReport).Left * 2
        picPaper(intReport).Height = picBack.ScaleHeight - picPaper(intReport).Top * 2
    Else
        With mobjReport.Fmts("_" & mobjReport.bytFormat)
            If .ֽ�� = 1 Then
                picPaper(intReport).Width = .W
                picPaper(intReport).Height = .H
            Else
                picPaper(intReport).Width = .H
                picPaper(intReport).Height = .W
            End If
        End With
        picShadow.Width = picPaper(intReport).Width
        picShadow.Height = picPaper(intReport).Height
        
        If picBack.ScaleWidth >= picPaper(intReport).Width + 180 Then
            picPaper(intReport).Left = (picBack.ScaleWidth - (picPaper(intReport).Width + 180)) / 2 + 60
            scrHsc.Enabled = False
        Else
            picPaper(intReport).Left = 60
            scrHsc.Max = (picPaper(intReport).Width + 180 - picBack.ScaleWidth) / 15
            If scrHsc.Max / 3 < scrHsc.SmallChange Then
                scrHsc.LargeChange = scrHsc.SmallChange
            Else
                scrHsc.LargeChange = scrHsc.Max / 3
            End If
            scrHsc.Enabled = True
        End If
        
        If picBack.ScaleHeight >= picPaper(intReport).Height + 180 Then
            picPaper(intReport).Top = (picBack.ScaleHeight - (picPaper(intReport).Height + 180)) / 2 + 60
            scrVsc.Enabled = False
        Else
            picPaper(intReport).Top = 60
            scrVsc.Max = (picPaper(intReport).Height + 180 - picBack.ScaleHeight) / 15
            If scrVsc.Max / 3 < scrVsc.SmallChange Then
                scrVsc.LargeChange = scrVsc.SmallChange
            Else
                scrVsc.LargeChange = scrVsc.Max / 3
            End If
            scrVsc.Enabled = True
        End If
        
        picShadow.Top = picPaper(intReport).Top + 60
        picShadow.Left = picPaper(intReport).Left + 60
    End If
    Exit Sub
errH:
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub scrhsc_Change()
    Dim sngPer As Single, sngPre As Single
    
    If Not mobjCurDLL Is Nothing Then
        sngPer = scrHsc.Value / (scrHsc.Max - scrHsc.Min) * 100
        sngPre = sngPer
        Call mobjCurDLL.Act_PaperScroll(mobjReport.���, 0, sngPer)
        If sngPre <> sngPer And sngPer >= 0 And sngPer <= 100 Then
            scrHsc.Value = (scrHsc.Max - scrHsc.Min) * (sngPer / 100)
        End If
    End If
    picPaper(intReport).Left = -scrHsc.Value * 15# + 60
    picShadow.Left = picPaper(intReport).Left + 60
    Me.Refresh
End Sub

Private Sub scrhsc_Scroll()
    Dim sngPer As Single, sngPre As Single
    
    If Not mobjCurDLL Is Nothing Then
        sngPer = scrHsc.Value / (scrHsc.Max - scrHsc.Min) * 100
        sngPre = sngPer
        Call mobjCurDLL.Act_PaperScroll(mobjReport.���, 0, sngPer)
        If sngPre <> sngPer And sngPer >= 0 And sngPer <= 100 Then
            scrHsc.Value = (scrHsc.Max - scrHsc.Min) * (sngPer / 100)
        End If
    End If
    picPaper(intReport).Left = -scrHsc.Value * 15# + 60
    picShadow.Left = picPaper(intReport).Left + 60
    Me.Refresh
End Sub

Private Sub scrVsc_Change()
    Dim sngPer As Single, sngPre As Single
    
    If Not mobjCurDLL Is Nothing Then
        sngPer = scrVsc.Value / (scrVsc.Max - scrVsc.Min) * 100
        sngPre = sngPer
        Call mobjCurDLL.Act_PaperScroll(mobjReport.���, 1, sngPer)
        If sngPre <> sngPer And sngPer >= 0 And sngPer <= 100 Then
            scrVsc.Value = (scrVsc.Max - scrVsc.Min) * (sngPer / 100)
        End If
    End If
    picPaper(intReport).Top = -scrVsc.Value * 15# + 60
    picShadow.Top = picPaper(intReport).Top + 60
    Me.Refresh
End Sub

Private Sub scrVsc_Scroll()
    Dim sngPer As Single, sngPre As Single
    
    If Not mobjCurDLL Is Nothing Then
        sngPer = scrVsc.Value / (scrVsc.Max - scrVsc.Min) * 100
        sngPre = sngPer
        Call mobjCurDLL.Act_PaperScroll(mobjReport.���, 1, sngPer)
        If sngPre <> sngPer And sngPer >= 0 And sngPer <= 100 Then
            scrVsc.Value = (scrVsc.Max - scrVsc.Min) * (sngPer / 100)
        End If
    End If
    picPaper(intReport).Top = -scrVsc.Value * 15# + 60
    picShadow.Top = picPaper(intReport).Top + 60
    Me.Refresh
End Sub

Private Sub tbr_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
        Case "Auto"
            mnuEdit_SetCol_Auto_Click
        Case "Def"
            mnuEdit_SetCol_Def_Click
        Case "Fill"
            mnuEdit_SetCol_Fill_Click
        Case "Large"
            Call SetView(0)
        Case "Small"
            Call SetView(1)
        Case "List"
            Call SetView(2)
        Case "Detail"
            Call SetView(3)
        Case "RowMode"
            mnuEdit_SelMode_Row_Click
        Case "ColMode"
            mnuEdit_SelMode_Col_Click
    End Select
End Sub

Private Sub tbr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuView, 2
End Sub

Private Function GetGridSource(objItem As RPTItem, Optional ByVal blnHead As Boolean) As String
'���ܣ��������������������õ�������Դ��
'������objItem=����������
'      blnHead=�Ƿ�ӱ�ͷ��ǩ�м��
'���أ�"������Ϣ,ҩƷ��Ϣ,...",""
    Dim tmpID As RelatID
    Dim strSource As String, strFormula As String
    
    For Each tmpID In objItem.SubIDs
        strFormula = mobjReport.Items("_" & tmpID.id).����
        Do While InStr(strFormula, "[") > 0
            strSource = Trim(Mid(strFormula, InStr(strFormula, "[") + 1, InStr(strFormula, "]") - InStr(strFormula, "[") - 1))
            strFormula = Mid(strFormula, InStr(strFormula, "]") + 1)
            If InStr(strSource, ".") > 0 Then
                If InStr(GetGridSource & ",", "," & Left(strSource, InStr(strSource, ".") - 1) & ",") = 0 Then
                    GetGridSource = GetGridSource & "," & Left(strSource, InStr(strSource, ".") - 1)
                End If
            End If
        Loop
        
        If blnHead Then
            strFormula = mobjReport.Items("_" & tmpID.id).��ͷ
            Do While InStr(strFormula, "[") > 0
                strSource = Trim(Mid(strFormula, InStr(strFormula, "[") + 1, InStr(strFormula, "]") - InStr(strFormula, "[") - 1))
                strFormula = Mid(strFormula, InStr(strFormula, "]") + 1)
                If InStr(strSource, ".") > 0 Then
                    If InStr(GetGridSource & ",", "," & Left(strSource, InStr(strSource, ".") - 1) & ",") = 0 Then
                        GetGridSource = GetGridSource & "," & Left(strSource, InStr(strSource, ".") - 1)
                    End If
                End If
            Loop
        End If
    Next
    If GetGridSource <> "" Then GetGridSource = Mid(GetGridSource, 2)
End Function

Private Function ReplaceUserPars(objReport As Report) As Boolean
'���ܣ�����ʹ���ߴ���������ñ������ֵ
'���أ�ʹ�����Ƿ���(ȫ��)(��ȷ)����
'˵����Ϊ�˱���"="���������Ͳ�����ͻ,����ʹ���߲�����ʽʱ����Split����,����Instr����
    Dim tmpData As RPTData, tmpPar As RPTPar
    Dim i As Integer, j As Integer, k As Integer
    Dim blnCur As Boolean, blnALL As Boolean
    Dim strTmp As String
    
    If Not IsArray(marrPars) Then Exit Function
    If UBound(marrPars) <> -1 Then
        '���ж��Ƿ�ȫ������
        blnALL = True
        For Each tmpData In objReport.Datas
            For Each tmpPar In tmpData.Pars
                blnCur = False: k = k + 1
                For i = 0 To UBound(marrPars)
                    '����������ͬ�Ҹ�ʽ�Ϸ����滻
                    j = InStr(CStr(marrPars(i)), "=")
                    If j > 0 Then
                        If UCase(Trim(Left(CStr(marrPars(i)), j - 1))) = UCase(tmpPar.����) Then
                            strTmp = Trim(Mid(CStr(marrPars(i)), j + 1))
                            If strTmp <> "" Then
                                If InStr(strTmp, "|") > 0 And (tmpPar.ȱʡֵ = "ѡ�������塭" Or tmpPar.ȱʡֵ = "�̶�ֵ�б�") Then
                                    blnCur = True: Exit For
                                Else
                                    Select Case tmpPar.����
                                        Case 0, 3
                                            blnCur = True: Exit For
                                        Case 1
                                            If IsNumeric(strTmp) Then blnCur = True: Exit For
                                        Case 2
                                            If IsDate(strTmp) Then blnCur = True: Exit For
                                    End Select
                                End If
                            End If
                        End If
                    End If
                Next
                blnALL = blnALL And blnCur
            Next
        Next
        
        '�ٴ���
        For Each tmpData In objReport.Datas
            For Each tmpPar In tmpData.Pars
                k = k + 1
                For i = 0 To UBound(marrPars)
                    '����������ͬ�Ҹ�ʽ�Ϸ����滻
                    j = InStr(CStr(marrPars(i)), "=")
                    If j > 0 Then
                        If UCase(Trim(Left(CStr(marrPars(i)), j - 1))) = UCase(tmpPar.����) Then
                            strTmp = Trim(Mid(CStr(marrPars(i)), j + 1))
                            If strTmp <> "" Then
                                '���ɳ���ֻ�����˰�ֵʱ,ģ��Ϊ��������ʾֵ
                                If InStr(strTmp, "|") = 0 And (tmpPar.ȱʡֵ = "ѡ�������塭" Or tmpPar.ȱʡֵ = "�̶�ֵ�б�") Then
                                    If tmpPar.ȱʡֵ = "�̶�ֵ�б�" Then
                                        For j = 0 To UBound(Split(tmpPar.ֵ�б�, "|"))
                                            If Split(Split(tmpPar.ֵ�б�, "|")(j), ",")(1) = strTmp Then
                                                strTmp = Split(Split(tmpPar.ֵ�б�, "|")(j), ",")(0) & "|" & strTmp
                                                If Left(strTmp, 1) = "��" Then strTmp = Mid(strTmp, 2)
                                                Exit For
                                            End If
                                        Next
                                    Else
                                        '�����������������������,�����ڲ�����������������
                                        strTmp = "������|" & strTmp
                                    End If
                                End If
                                If InStr(strTmp, "|") > 0 And (tmpPar.ȱʡֵ = "ѡ�������塭" Or tmpPar.ȱʡֵ = "�̶�ֵ�б�") Then
                                    '����ʾֵ,��ֵ����Ĳ�����
                                    If Not blnALL Then
                                        '��δ��������������,ԭ���ǽ���������ģ��Ϊ����ʱ��ֵ
                                        tmpPar.Reserve = strTmp
                                    Else
                                        '�ڴ���������,ԭ���ǽ���������ģ��Ϊ��Ҫִ��ʱ��ֵ
                                        tmpPar.Reserve = tmpPar.ȱʡֵ & "|" & Split(strTmp, "|")(0)
                                        tmpPar.ȱʡֵ = Split(strTmp, "|")(1)
                                    End If
                                    Exit For '��ǰ�������滻,��������
                                Else
                                    'һ�㴫�����,һ��Ҫ����,�����������ѡ��
                                    '�����Ƿ���,��ֱ�Ӵ���ȱʡֵ
                                    If tmpPar.Reserve = "" And Left(tmpPar.ȱʡֵ, 1) = "&" Then
                                        tmpPar.Reserve = tmpPar.ȱʡֵ
                                    End If
                                    Select Case tmpPar.����
                                        Case 0, 3
                                            tmpPar.ȱʡֵ = strTmp: Exit For
                                        Case 1
                                            If IsNumeric(strTmp) Then tmpPar.ȱʡֵ = strTmp: Exit For
                                        Case 2
                                            If IsDate(strTmp) Then tmpPar.ȱʡֵ = strTmp: Exit For
                                    End Select
                                End If
                            End If
                        End If
                    End If
                Next
            Next
        Next
    End If
    ReplaceUserPars = blnALL
End Function

Private Function ParCount(objReport As Report) As Integer
'���ܣ��ӱ�������з��ز��ظ����Ʋ�������
    Dim tmpPar As RPTPar, tmpData As RPTData, StrPar As String
    
    If objReport.Datas.count = 0 Then ParCount = 0: Exit Function
    For Each tmpData In objReport.Datas
        For Each tmpPar In tmpData.Pars
            If InStr(StrPar & ",", "," & tmpPar.���� & ",") = 0 Then
                StrPar = StrPar & "," & tmpPar.����
                ParCount = ParCount + 1
            End If
        Next
    Next
End Function

Private Sub ReplaceInputPars(objPars As RPTPars)
'���ܣ����ݲ������봰������Ĳ���ֵ(����Ψһ)�滻��������Դ�Ĳ�����
    Dim tmpData As RPTData, tmpPar As RPTPar, objPar As RPTPar
    
    For Each tmpData In mobjReport.Datas
        For Each tmpPar In tmpData.Pars
            '�Ե�ǰ���������滻
            For Each objPar In objPars
                If objPar.���� = tmpPar.���� Then
                    tmpPar.ȱʡֵ = objPar.ȱʡֵ
                    tmpPar.Reserve = objPar.Reserve
                    Exit For '��һ���������
                End If
            Next
        Next
    Next
End Sub

Private Function OpenReportData(Optional ByVal blnAllReLoad As Boolean = True) As String
'���ܣ����ݱ������(mobjReport)��ǰ��ʽ������Դ����,��������ı������ݼ�
'���ܣ�blnAllReLoad=�Ƿ���ȫ�����¶�ȡ����Դ(����ˢ��,��������ʱ,���л���ʽʱֻ�ض���Ҫ��)
'���أ��ɹ�="",ʧ��="����Դ��"
    Dim tmpData As RPTData, strName As String
    Dim rsTmp As ADODB.Recordset
    Dim blnDo As Boolean, i As Integer
    
    'û�ж�������Դ
    mobjReport.blnLoad = True '��ʾ�����ζ�ȡ�Ƿ���ȷ
    If mobjReport.Datas.count = 0 Then Exit Function
    
    If blnAllReLoad Then
        Set mLibDatas = Nothing
        Set mLibDatas = New LibDatas
    ElseIf mLibDatas Is Nothing Then
        Set mLibDatas = New LibDatas
    End If
    
    On Error GoTo hErr
            
    For Each tmpData In mobjReport.Datas
        '�жϸ�����Դ�Ƿ��Ѷ�ȡ
        blnDo = True
        For i = 1 To mLibDatas.count
            If mLibDatas(i).Key = tmpData.���� Then
                blnDo = False: Exit For
            End If
        Next
        '��ȡ��ǰ��ʽ�õ�������Դ
        If blnDo And DataUsed(mobjReport, tmpData.����, True) Then
            strName = tmpData.����
            Set rsTmp = Nothing
            Set rsTmp = OpenReportSQL(tmpData)
            If rsTmp Is Nothing Then
                OpenReportData = tmpData.����
                mobjReport.blnLoad = False
                Call ShowFlash: Exit Function
            End If
            mLibDatas.Add strName, rsTmp, "_" & strName
        End If
    Next
    
    Call ShowFlash
    Exit Function
    
hErr:
    If ErrCenter = 1 Then Resume
End Function

Private Function OpenReportSQL(objData As RPTData) As ADODB.Recordset
'���ܣ���������Դ�������ݴ򿪼�¼��
'˵��������̬ADO.Command�������صļ�¼����Clone��ȥ�����Ҹ�Clone���ڴ�״̬ʱ���ظ�ִ��Command����ֶ����Ѵ򿪴���
'1.ִ�б�����ʱ,�����������,Clone�ļ�¼���ֲ��ܹر�,��˲�ʹ�þ�̬������
'2.�������������������,����Static����.���ú���Ӧ�ŵ�����ģ����,��Ȼ����ر���StaticҲû��Ч����
'  ��Ϊ���������ظ�ִ��Ƶ�ʲ���ܸ�,���Ӹ�ֵ��Ч��Ӱ����Ժ��ԡ�
'3.��������д�������󶨡��磺select '[0]' ���� from ...

    Dim rsTmp As New ADODB.Recordset
    Dim cmdData As New ADODB.Command
    Dim strLeft As String, strRight As String
    Dim StrPar As String, strParOld As String, bytPar As Byte
    Dim strSQL As String, strLog As String
    Dim intMax As Integer
    Dim strSQLtmp As String, i As Long, arrStr As Variant
    Dim strTmp As String, strSQLtmp1 As String
    Dim intDateType  As Integer  '0=�޼Ӽ����㣬1=�Ӽ�������2=�Ӽ����������ֶ�,3=�������㣬������ǰ�Ĺ��򣬴����ַ��󶨱���
    Dim j As Long, k As Long, datValue As Date
    Dim l As Long

    If mbytStyle = 0 Or mbytStyle = 1 Then
        ShowFlash "���ڶ�ȡ����""" & objData.���� & """�����Ժ򣮣���", , Me
    End If

    On Error GoTo errHandle

    '����ԭʼSQL
    'strSql = SQLOwner(TrimChar(objData.SQL), objData.����)
    strSQL = SQLOwner(RemoveNote(objData.SQL), objData.����)
    
    '������ʹ���˶�̬�ڴ������û��ʹ��/*+ XXX*/����ʾ��ʱ�Զ�����
    strSQLtmp = Trim(UCase(strSQL))
    If Mid(Trim(Mid(strSQLtmp, 7)), 1, 2) <> "/*" And Mid(strSQLtmp, 1, 6) = "SELECT" Then
        If Not Replace(strSQLtmp, " ", "") Like "*/[*]+CARDINALITY*[*]/*" Then      '/**/����ܳ��ֶ��CARDINALITY
            arrStr = Split("F_STR2LIST,F_NUM2LIST,F_NUM2LIST2,F_STR2LIST2", ",")
            For i = 0 To UBound(arrStr)
                strSQLtmp1 = strSQLtmp
                Do While InStr(strSQLtmp1, arrStr(i)) > 0
                    '�ж�ǰ���Ƿ�����IN �����򲻼�Rule
                    '���ҵ����һ��SELECT
                    strTmp = Mid(strSQLtmp1, 1, InStr(strSQLtmp1, arrStr(i)) - 1)
                    strTmp = Replace(TrimEx(Mid(strTmp, 1, InStrRev(strTmp, "SELECT") - 1)), " ", "")
                    If Len(strTmp) > 1 Then strTmp = Mid(strTmp, Len(strTmp) - 2)               'ȡ����3���ַ�
                    
                    If strTmp = "IN(" Then '����in(select��������������ѭ�������Ƿ����û��ʹ������д����������̬�ڴ溯��
                       strSQLtmp1 = Mid(strSQLtmp1, InStr(strSQLtmp1, arrStr(i)) + Len(arrStr(i)))
                    Else
                        Exit For
                    End If
                Loop
            Next
            If i <= UBound(arrStr) Then
                strSQL = "Select /*+ RULE*/" & Mid(Trim(strSQL), 7)
            End If
        End If
    End If
    
    strLog = strSQL
        
    i = 1
    Do While i <= Len(strLog)
        If InStr(i, strLog, "[") <= 0 Then
            i = i + 1
            GoTo makContinue1
        End If
        strLeft = Left(strLog, InStr(i, strLog, "[") - 1)
        strTmp = Mid(strLog, InStr(i, strLog, "["))
        If mdlPublic.AtString(strLeft) _
            And Not (strTmp Like "[[][0-9][]]*" Or strTmp Like "[[][0-9][0-9][]]*") Then
            '�������ڵ��ַ��������Ҹ�ʽ��[0-99]
            i = i + 1
            GoTo makContinue1
        End If
        
        If InStr(i, strLog, "]") <= 0 Then
            i = i + 1
            GoTo makContinue1
        End If
        strRight = Mid(strLog, InStr(i, strLog, "]") + 1)
        If strRight <> "" And mdlPublic.AtString(strRight) _
            And Not (strTmp Like "[[][0-9][]]*" Or strTmp Like "[[][0-9][0-9][]]*") Then
            '�������ڵ��ַ��������Ҹ�ʽ��[0-99]
            i = i + 1
            GoTo makContinue1
        End If
        
        '��������Ĳ�����
        i = InStr(i, strLog, "[")
        strRight = Mid(strLog, InStr(i, strLog, "]") + 1)
        
        StrPar = Mid(strLog, InStr(i, strLog, "[") + 1, InStr(i, strLog, "]") - InStr(i, strLog, "[") - 1)
        strParOld = StrPar
        bytPar = Val(StrPar)
        Select Case objData.Pars("_" & CInt(bytPar)).����
            Case 0 '�ַ�
                StrPar = "'" & Replace(objData.Pars("_" & CInt(bytPar)).ȱʡֵ, "'", "''") & "'"
            Case 1 '����
                StrPar = objData.Pars("_" & CInt(bytPar)).ȱʡֵ
            Case 2 '����
                If Left(objData.Pars("_" & CInt(bytPar)).ȱʡֵ, 1) = "&" Then
                    StrPar = GetParSQLMacro(objData.Pars("_" & CInt(bytPar)).ȱʡֵ)
                Else
                    If Format(objData.Pars("_" & CInt(bytPar)).ȱʡֵ, "HH:mm:ss") = "00:00:00" Then
                        '��ʱ���ʽ
                        StrPar = "To_Date('" & Format(objData.Pars("_" & CInt(bytPar)).ȱʡֵ, "yyyy-MM-dd") & "','YYYY-MM-DD')"
                    Else
                        '��ʱ���ʽ
                        StrPar = "To_Date('" & Format(objData.Pars("_" & CInt(bytPar)).ȱʡֵ, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    End If
                End If
            Case 3 '������
                StrPar = objData.Pars("_" & CInt(bytPar)).ȱʡֵ
        End Select
        strLog = strLeft & StrPar & strRight
        
        i = Len(strLeft & strParOld)
        
makContinue1:
    Loop
        
    If InStr(UCase(objData.SQL), "--UNBOUND") > 0 Then GoTo LineOld

    '�����󶨲���SQL
    cmdData.CommandText = ""                        '��Ϊ����ʱ�����������
    cmdData.CommandType = adCmdText                 '����ΪadCmdText���ܸ���
    
    '���ԭ�в���:��Ȼ�����ظ�ִ��
    Do While cmdData.Parameters.count > 0
        cmdData.Parameters.Delete 0
    Loop
    
    l = 1
    Do While l <= Len(strSQL)
        If InStr(l, strSQL, "[") <= 0 Then
            l = l + 1
            GoTo makContinue2
        End If
        strLeft = Left(strSQL, InStr(l, strSQL, "[") - 1)
        strTmp = Mid(strSQL, InStr(l, strSQL, "["))
        If mdlPublic.AtString(strLeft) _
            And Not (strTmp Like "[[][0-9][]]*" Or strTmp Like "[[][0-9][0-9][]]*") Then
            '�������ڵ��ַ��������Ҹ�ʽ��[0-99]
            l = l + 1
            GoTo makContinue2
        End If
        
        If InStr(l, strSQL, "]") <= 0 Then
            l = l + 1
            GoTo makContinue2
        End If
        strRight = Mid(strSQL, InStr(l, strSQL, "]") + 1)
        If strRight <> "" And mdlPublic.AtString(strRight) _
            And Not (strTmp Like "[[][0-9][]]*" Or strTmp Like "[[][0-9][0-9][]]*") Then
            '�������ڵ��ַ��������Ҹ�ʽ��[0-99]
            l = l + 1
            GoTo makContinue2
        End If
        
        '��������Ĳ�����
        l = InStr(l, strSQL, "[")
        strRight = Mid(strSQL, InStr(l, strSQL, "]") + 1)
        
        StrPar = Mid(strSQL, InStr(l, strSQL, "[") + 1, InStr(l, strSQL, "]") - InStr(l, strSQL, "[") - 1)
        strParOld = StrPar
        bytPar = Val(StrPar)
        intDateType = 0
        datValue = CDate(0)
        strTmp = ""
        
        Select Case objData.Pars("_" & CInt(bytPar)).����
            Case 0 '�ַ�
                StrPar = objData.Pars("_" & CInt(bytPar)).ȱʡֵ
                intMax = LenB(StrConv(StrPar, vbFromUnicode))
                
                If intMax <= 2000 Then
                    intMax = IIF(intMax <= 200, 200, 2000)
                    cmdData.Parameters.Append cmdData.CreateParameter("PAR" & cmdData.Parameters.count + 1, adVarChar, adParamInput, intMax, StrPar)
                Else
                    If intMax < 4000 Then intMax = 4000
                    cmdData.Parameters.Append cmdData.CreateParameter("PAR" & cmdData.Parameters.count + 1, adLongVarChar, adParamInput, intMax, StrPar)
                End If
                
                strSQL = strLeft & "?" & strRight
            Case 1 '����
                StrPar = objData.Pars("_" & CInt(bytPar)).ȱʡֵ
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & cmdData.Parameters.count + 1, adVarNumeric, adParamInput, 30, Val(StrPar))

                strSQL = strLeft & "?" & strRight
            Case 2 '����
                If Left(objData.Pars("_" & CInt(bytPar)).ȱʡֵ, 1) = "&" Then
                    StrPar = GetParVBMacro(objData.Pars("_" & CInt(bytPar)).ȱʡֵ)
                Else
                    If Format(objData.Pars("_" & CInt(bytPar)).ȱʡֵ, "HH:mm:ss") = "00:00:00" Then
                        '��ʱ���ʽ
                        StrPar = Format(objData.Pars("_" & CInt(bytPar)).ȱʡֵ, "yyyy-MM-dd")
                    Else
                        '��ʱ���ʽ
                        StrPar = Format(objData.Pars("_" & CInt(bytPar)).ȱʡֵ, "yyyy-MM-dd HH:mm:ss")
                    End If
                End If
'                1��������������ڣ��Ӽ������㣬��ֱ��ʹ�������Ͱ󶨱�����
'                2������������ڣ��Ӽ������㣬��������ǳ���������ִ��һ��SQL�����ڳ����м��㣬�õ�������ֵ�����磺XX+1/24�����ٸ��ݵõ���ֵʹ�������Ͱ󶨱�����
'                      ������治�ǳ��������磺����ĳ���ֶλ�sysdate�����㣩����ʹ�ð󶨱�����ֱ�Ӵ������ֵ(sqlƴ��)��
'                      ���ֲ�ʹ�ð󶨱�������Ȼÿ��ִ����ҪӲ�������������ִ�мƻ����ܳ������Ƚϣ����۸�СһЩ��
 '                     ֻʶ���õļӼ��㷨��1��+1-1/24/60/60  2�� -1/24/60/60+1 ��3��-1/24/60/60  4���Ӽ�һ�����֣�û�����ӵ����
 '                     ��������㳣���㷨�������ӣ�+1/24 ���ֱ�����ǰ�Ĺ��򣬴����ַ��󶨱���
 '����SQL��
'                select * from ���ű� where ����ʱ�� >1+ [0]- 1  and  ID>0
'                Union All
'                select * from ���ű� where  [0]- 1=����ʱ��  and  ID>0
'                Union All
'                select * from ���ű� where  ����ʱ��>[0]+1 - 1/24 /60 /60  and  ID>0
'                Union All
'                select * from ���ű� where  ����ʱ��>[0] - 1/24 /60 /60+1  and  ID>0
'                Union All
'                select * from ���ű� where  ����ʱ��>[0] - 1 /24 /60 /60  and  ID>0
'                Union All
'                select * from ���ű� where  ����ʱ��>[0] - 1 /24 /60   and  ID>0
'                Union All
'                select * from ���ű� where  ����ʱ��>1+[0] - 1 /24 /60/60   and  ID>0

                '�Ȳ鿴����Ƿ��мӼ�����
                datValue = CDate(StrPar)
                
                For i = 1 To Len(strRight)
                    If Mid(strRight, i, 1) <> " " Then
                        If InStr("+-", Mid(strRight, i, 1)) > 0 Then
                            For j = i + 1 To Len(strRight)
                                If Mid(strRight, j, 1) <> " " Then
                                    '�ҵ������ֵ
                                    For k = j + 1 To Len(strRight)
                                        If Mid(strRight, k, 1) = " " Or (IsNumeric(Mid(strRight, j, 1)) And Not IsNumeric(Mid(strRight, j, k - j + 1))) Then
                                            If Not Mid(strRight, k, 1) = " " And Not IsNumeric(Mid(strRight, k - 1, 1)) Then
                                                k = k - 1
                                            End If
                                            Exit For
                                        End If
                                    Next
                                    If IsNumeric(Mid(strRight, j, k - j)) Then
                                        intDateType = 1
                                        
                                        '�������ֵ
                                        '�����ļ��㷽ʽ�����жϳ���
                                        If InStr(Replace(strRight, " ", ""), "+1-1/24/60/60") = 1 Then
                                            datValue = datValue + 1 - 1 / 24 / 60 / 60
                                            strTmp = Mid(strRight, InStr(Mid(strRight, InStr(strRight, "60") + 2), "60") + 2 + InStr(strRight, "60") + 1)
                                        ElseIf InStr(Replace(strRight, " ", ""), "-1/24/60/60+1") = 1 Then
                                            datValue = datValue + 1 - 1 / 24 / 60 / 60
                                            strTmp = Mid(strRight, InStr(Mid(strRight, InStr(strRight, "+") + 1), "1") + InStr(strRight, "+") + 1)
                                        ElseIf InStr(Replace(strRight, " ", ""), "-1/24/60/60") = 1 Then
                                            datValue = datValue - 1 / 24 / 60 / 60
                                            strTmp = Mid(strRight, InStr(Mid(strRight, InStr(strRight, "60") + 2), "60") + 2 + InStr(strRight, "60") + 1)
                                        Else
                                            If Mid(strRight, i, 1) = "+" Then
                                                datValue = datValue + Val(Mid(strRight, j, k - j))
                                            Else
                                                datValue = datValue - Val(Mid(strRight, j, k - j))
                                            End If
                                            strTmp = Mid(strRight, k)
                                        End If
                                        If InStr("+-*/", Mid(Replace(strTmp, " ", ""), 1, 1)) > 0 And Replace(strRight, " ", "") <> "" Then
                                            '�������û��+-*/���ʾ�ǵ����ļӼ���,���򱣳���ǰ�Ĺ���
                                            intDateType = 3
                                        End If
                                    Else
                                        intDateType = 2
                                    End If
                                    Exit For
                                End If
                            Next
                        Else
                            Exit For
                        End If
                        Exit For
                    End If
                Next
                'ǰ��Ӽ������������ִ����ַ��󶨱����Ĺ���
                If intDateType <> 2 Then
                    For i = Len(strLeft) To 1 Step -1
                        If Mid(strLeft, i, 1) <> " " Then
                            If InStr("+-", Mid(strLeft, i, 1)) > 0 Then
                               intDateType = 3
                            End If
                            Exit For
                        End If
                    Next
                End If
                If intDateType = 2 Then
                    '��ʹ�ð󶨱���
                    strSQL = strLeft & "To_Date('" & Format(datValue, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')" & strRight
                ElseIf intDateType = 3 Then
                    '������ת��Ϊ�ַ����͵�SQL�а�,��Ϊ�������͵İ󶨱�������������Ҫ����
                    cmdData.Parameters.Append cmdData.CreateParameter("PAR" & cmdData.Parameters.count + 1, adVarChar, adParamInput, Len(StrPar), StrPar)
                    If StrPar Like "*:*:*" Then
                        strSQL = strLeft & "To_Date(?,'YYYY-MM-DD HH24:MI:SS')" & strRight
                    Else
                        strSQL = strLeft & "To_Date(?,'YYYY-MM-DD')" & strRight
                    End If
                Else
                    '�����︳ֵ��Ҫ�Ǵ��� �������������ģ���ǰ�����������
                    If intDateType = 1 Then strRight = strTmp
                    '����ǳ����Ӽ�����û�мӼ���ֱ��ʹ�����ڰ󶨱���
                    cmdData.Parameters.Append cmdData.CreateParameter("PAR" & cmdData.Parameters.count + 1, adDBTimeStamp, adParamInput, , datValue)
                    strSQL = strLeft & "?" & strRight
                End If
            Case 3 '������
                StrPar = objData.Pars("_" & CInt(bytPar)).ȱʡֵ

                strSQL = strLeft & StrPar & strRight
        End Select
        
        l = Len(strLeft & strParOld)
        
makContinue2:
    Loop
    
    '����ROLLUP�Ļ���WITH��ͷ�Ķ�����SELECT* Ƕ��
    If InStr(strSQL, "ROLLUP") > 0 Or Mid(strSQL, 1, 4) = "WITH" Then
        strSQL = "SELECT * FROM(" & strSQL & ")"
    End If
    'ִ�з��ؼ�¼��
'    If cmdData.ActiveConnection Is Nothing Then
'        Set cmdData.ActiveConnection = gcnOracle '���Ƚ���
'    End If
    Set cmdData.ActiveConnection = mdlPublic.GetDBConnection(objData.�������ӱ��)
    cmdData.CommandText = strSQL

LineBand:
    Call SQLTest(App.ProductName, "OpenReportSQL", strLog)
    
    If SQLExistLOB(objData) Then
        If rsTmp Is Nothing Then Set rsTmp = New ADODB.Recordset
        rsTmp.Open cmdData, , adOpenStatic, adLockOptimistic
    Else
        Set rsTmp = cmdData.Execute
    End If
    
    Set rsTmp.ActiveConnection = Nothing        '��zl9ComLib��Recordset������һ��
    Call SQLTest
    Set OpenReportSQL = rsTmp
    Exit Function
    
LineOld:
    Call OpenRecord(rsTmp, strLog, "OpenReportSQL", objData.�������ӱ��)
    Set rsTmp.ActiveConnection = Nothing        '��zl9ComLib��Recordset������һ��
    Set OpenReportSQL = rsTmp
    Exit Function
    
LineBlob:
    '����ODBC��֧�ֺ���LOB�е�SQL�����Ը�Ϊʹ��OLEDB��ѯ
    If objData.�������ӱ�� <= 0 Then
        If gcolOLEDBConnect Is Nothing Then
            Set gcolOLEDBConnect = New Collection
        End If
        '��ȡ�������Ӷ���
        Set gcnOLEDB = mdlPublic.GetOLEDBConnect(gcnOracle, gcolOLEDBConnect, gobjRegister)
        If gcnOLEDB Is Nothing Then
            Set gcnOLEDB = gobjRegister.ReGetConnection(Val("1-OracleOLEDB"), "", gcnOracle)
            
            If gcnOLEDB.State = adStateClosed Then
                strTmp = "������������ʧ�ܣ����飺" & Chr(10) & _
                         "1.ҵ�����ʹ��zlRegister���������Ĳ����Ƿ���ȷ��" & Chr(10) & _
                         "2.�ǵ���̨������ñ�������������Ӷ���΢��������������ͨ��" & Chr(10) & _
                         "  zlRegister�����������Ӷ���"
                If gblnSilentMode Then
                    gstrErrorContent = strTmp
                Else
                    MsgBox strTmp, vbInformation, App.Title
                End If
                Exit Function
            End If
            
            '����
            Call gcolOLEDBConnect.Add(gcnOLEDB)
        End If
        Set cmdData.ActiveConnection = gcnOLEDB
    Else
        Set cmdData.ActiveConnection = mdlPublic.GetDBConnectionEx(Val("1-OracleOLEDB"), objData.�������ӱ��)
    End If
    If Not cmdData.ActiveConnection Is Nothing Then
        'Set rsTmp = cmdData.Execute
        'CLOB��BLOB�ֶ��������ʹ��Command���󣬼�¼������Ĭ�ϵ���adOpenUnspecifiedִ�л����
        '��ˣ����ü�¼�������Open����
        If rsTmp Is Nothing Then Set rsTmp = New ADODB.Recordset
        rsTmp.Open cmdData, , adOpenStatic, adLockOptimistic
        Set rsTmp.ActiveConnection = Nothing        '��zl9ComLib��Recordset������һ��
        Set OpenReportSQL = rsTmp
    End If
    Exit Function
    
errHandle:
    'ORA-00979:���� GROUP BY ���ʽ
    'SQL�е�"?"���ύ��Oracleʱ��ADO˳��Ϊ":P1,:P2"��ʽ,Group by��Ϊ���ǲ�ͬ�ķ����ֶ�,��ʹ��ֵ��ͬ
    '�����ӷ�ʽ��,ADO��SQL�в���ʹ��":P"���ֲ���(ʼ��˵����δ����)
    '�����ӷ�ʽ��,ADO��SQL�п���ʹ��":P"���ֲ���,����Parameters�����д�����ֻ��˳���Ӧ,���Ʋ���Ӧ.
    '    ����ͨ������ʹ��Group���漰��������ͬ�ֶ�һ��,��������Ϊ������ͬ���ٴ���һЩ����
    If Err.Description Like "*ORA-00979*" Then Err.Clear: GoTo LineOld
    
    'ORA-00932: �������Ͳ�һ��: ӦΪ NUMBER, ��ȴ��� -
    '��Group By Rollup��Decode����ʱ�����ܻ���ָô�����δ֪��ȷԭ��,����ȷ����취
    'ʵ������������ָô���ʱ��SQL��δ����ִ�У��ٶ���Ӧû��Ӱ��
    If Err.Description Like "*ORA-00932*" Then Err.Clear: GoTo LineOld
    
     'MS��ODBC���ӣ���ѯBLOB�ֶ�ʱ�ᱨ��(ִ��һ���ṩ��������ʱ�����ṩ����ʧ�ܡ�)��������ʱ��ΪOraOLEDB���Ӷ���������
    If Err.Number = -2147467259 Then Err.Clear: GoTo LineBlob
    
    Call ShowFlash
    If Err.Description Like "*ORA-00920*" Then
        MsgBox "����������󣬵��²�����ȷ��ȡ����""" & objData.���� & """��", vbExclamation, App.Title
    ElseIf ErrCenter() = 1 Then
        If mbytStyle = 0 Or mbytStyle = 1 Then ShowFlash "���ڶ�ȡ����""" & objData.���� & """�����Ժ򣮣���", , Me
        Resume
    End If
    Call SaveErrLog
End Function

Private Function EvalFormula(ByVal strFormula As String, idx As Integer, Row As Long) As String
'���ܣ�������ʽ��ֵ
'������strFormula=��﹫ʽ,idx:�Ѵ���һ�������ݵı������,Row=��ǰ��
'���أ�������ֵ,������󷵻ؿ�
'�ο���mLibDatas
    Dim strLeft As String, strRight As String, strVar As String
    
    On Error Resume Next
    
    strFormula = Trim(strFormula)
    
    If strFormula = "" Then '����
        Exit Function
    ElseIf InStr(strFormula, "[") = 0 Then '��������
        EvalFormula = Srt.Eval(strFormula)
    ElseIf Left(strFormula, 1) = "[" And Right(strFormula, 1) = "]" And InStr(strFormula, ".") > 0 _
        And InStr(Mid(strFormula, 2, Len(strFormula) - 2), "[") = 0 Then
         'ֻ���ֶ����õ���
         EvalFormula = GetFieldValue(Me, Mid(strFormula, 2, Len(strFormula) - 2))
    ElseIf Left(strFormula, 1) = "[" And Right(strFormula, 1) = "]" And InStr(strFormula, ".") > 0 _
        And InStr(Mid(strFormula, 2, Len(strFormula) - 2), "[") = 0 Then
         'ֻ�������õ���
         EvalFormula = msh(idx).TextMatrix(Row, CInt(Mid(strFormula, 2, Len(strFormula) - 2)))
    Else '���ϼ���
        Do While InStr(strFormula, "[") > 0
            strLeft = Left(strFormula, InStr(strFormula, "[") - 1)
            strRight = Mid(strFormula, InStr(strFormula, "]") + 1)
            strVar = Mid(strFormula, InStr(strFormula, "[") + 1, InStr(strFormula, "]") - InStr(strFormula, "[") - 1)
            
            If IsNumeric(Mid(strVar, 2)) And Left(strVar, 1) = "@" Then
                If Row = msh(idx).FixedRows Then
                    strVar = "" '��һ�������޷�ȡ
                Else
                    If InStr(strFormula, """[" & strVar & "]""") > 0 And InStr(msh(idx).TextMatrix(Row - 1, CInt(Mid(strVar, 2))), """") > 0 Then
                        '�ַ������㼰��Ԫֵ�а����ַ���
                        strVar = Replace(msh(idx).TextMatrix(Row - 1, CInt(Mid(strVar, 2))), """", """""")
                    Else
                        strVar = msh(idx).TextMatrix(Row - 1, CInt(Mid(strVar, 2))) 'ֱ��ȡ��Ӧ��ֵ
                    End If
                End If
                If strVar = "" Then strVar = 0
            ElseIf IsNumeric(strVar) Then
                If InStr(strFormula, """[" & strVar & "]""") > 0 And InStr(msh(idx).TextMatrix(Row, CInt(strVar)), """") > 0 Then
                    '�ַ������㼰��Ԫֵ�а����ַ���
                    strVar = Replace(msh(idx).TextMatrix(Row, CInt(strVar)), """", """""")
                Else
                    strVar = msh(idx).TextMatrix(Row, CInt(strVar)) 'ֱ��ȡ��Ӧ��ֵ
                End If
                If strVar = "" Then strVar = 0
            ElseIf InStr(strVar, ".") > 0 Then
                '���Ϊ��,����"Null",���ʽҪ�������ж�
                strVar = GetFieldValue(Me, strVar, True) '���������ڻ��ַ���������ʱ,�Զ�ת����ʽ
            End If
            
            '�滻������ѭ��
            If InStr(strVar, "[") > 0 Or InStr(strVar, "]") > 0 Then
                strVar = Replace(strVar, "[", Chr(1) & "SKIPCYCLEFT" & Chr(1))
                strVar = Replace(strVar, "]", Chr(1) & "SKIPCYCRIGHT" & Chr(1))
            End If
            strFormula = strLeft & strVar & strRight
        Loop
        strFormula = Replace(strFormula, Chr(1) & "SKIPCYCLEFT" & Chr(1), "[")
        strFormula = Replace(strFormula, Chr(1) & "SKIPCYCRIGHT" & Chr(1), "]")
        EvalFormula = Srt.Eval(strFormula)
    End If
End Function

Private Function SortFormula(objItem As RPTItem) As Variant
    Dim tmpID As RelatID, tmpItem As RPTItem
    Dim arrFormula() As String, strTmp As String
    Dim strReferCols As String, intReferCols As Integer
    Dim intCol As Integer, intCur As Integer
    Dim i As Integer, j As Integer
    Dim strDie As String, strOrder As String
    
    
    ReDim arrFormula(objItem.SubIDs.count - 1) As String
    
    '��˳�����������
    For Each tmpID In objItem.SubIDs
        Set tmpItem = mobjReport.Items("_" & tmpID.id)
        arrFormula(tmpItem.���) = tmpItem.���� & "|" & tmpItem.��ʽ & "|" & tmpItem.��� & "|" & tmpItem.����
    Next
    
    '����"����"�İ�����ϵ����
    i = 0
    strOrder = GetOrder(arrFormula)
    Do While i <= UBound(arrFormula)
        '��������Щ��
        strReferCols = GetReferCols(CStr(Split(arrFormula(i), "|")(0)))
        intReferCols = UBound(Split(strReferCols, ","))
        
        intCur = i '���ǰλ��
        For j = 0 To intReferCols
            '�����ǰλ��
            intCol = GetReferLoc(arrFormula, CInt(Split(strReferCols, ",")(j)))
            If intCol > intCur Then
                strTmp = arrFormula(intCur)
                arrFormula(intCur) = arrFormula(intCol)
                arrFormula(intCol) = strTmp
                intCur = intCol
            End If
        Next
        '���һ��Ҳû�л�,��ʾû������,�������һ������
        '������Ȼ����λ�ÿ�ʼ��������
        strDie = GetOrder(arrFormula)
        If intCur = i Or (intCur <> i And strOrder = strDie) Then
            i = i + 1
            strOrder = strDie
        End If
    Loop
    
    SortFormula = arrFormula
End Function

Private Function GetOrder(arrFormula() As String) As String
'���ܣ����ص�ǰ��ʽ�����и��к����е�˳��,���ڼ����ѭ��
    Dim i As Integer
    For i = 0 To UBound(arrFormula)
        GetOrder = GetOrder & "," & CInt(Split(arrFormula(i), "|")(2))
    Next
    GetOrder = Mid(GetOrder, 2)
End Function

Private Function GetReferLoc(arrFormula() As String, intCol As Integer) As Integer
'���ܣ��������ΪintCol����Ŀ�������е�λ��
    Dim i As Integer
    For i = 0 To UBound(arrFormula)
        If CInt(Split(arrFormula(i), "|")(2)) = intCol Then
            GetReferLoc = i: Exit Function
        End If
    Next
End Function

Private Function GetReferCols(ByVal strFormula As String) As String
'���ܣ����ع�ʽ�����õ��к�,��"3,5,6"
    Dim strRight As String, strCol As String, strCols As String
    
    strFormula = Trim(strFormula)
    
    Do While InStr(strFormula, "[") > 0
        strRight = Mid(strFormula, InStr(strFormula, "]") + 1)
        strCol = Mid(strFormula, InStr(strFormula, "[") + 1, InStr(strFormula, "]") - InStr(strFormula, "[") - 1)
        If IsNumeric(strCol) Then strCols = strCols & "," & strCol
        strFormula = strRight
    Loop
    GetReferCols = Mid(strCols, 2)
End Function

Private Sub SetRedraw(blnDraw As Boolean)
    Dim obj As Object
    For Each obj In msh
        If obj.Index <> 0 And (obj.Container Is picPaper(intReport) Or UCase(obj.Container.name) = "PIC") Then obj.Redraw = blnDraw
    Next
End Sub

Private Sub mnuHelpWebHome_Click()
    zlHomePage hwnd
End Sub

Private Sub mnuHelpWebMail_Click()
    zlMailTo hwnd
End Sub

Private Function GetLR(msh As Object, Col As Integer) As Byte
    Select Case msh.ColAlignment(Col)
        Case 0, 1, 2 '�����
            GetLR = 2 '�Ҽӿո�
        Case 3, 4, 5 '�ж���
            GetLR = 1 '˫�ӿո�
        Case 6, 7, 8 '�Ҷ���
            GetLR = 0 '��ӿո�
    End Select
End Function

Private Function GetRowText(msh As Object, Row As Long, Col As Long) As String
    Dim i As Integer
    Dim strTmp As String
    
    For i = 0 To Col
        strTmp = strTmp & Trim(msh.TextMatrix(Row, i))
    Next
    GetRowText = strTmp
End Function

Private Function GetColText(msh As Object, Row As Long, Col As Long) As String
    Dim i As Integer
    Dim strTmp As String
    
    For i = 0 To Row
        strTmp = strTmp & Trim(msh.TextMatrix(i, Col))
    Next
    GetColText = strTmp
End Function

Private Function GetColType(ByVal strFormula As String) As Byte
'���ܣ��ж�������ĳһ�е���������
'������strFormula=�м��㹫ʽ
'���أ�0=��ȷ��,1-�ַ�(����),2=����,3=����
'�ο���mLibDatas
    Dim varR As Variant, strData As String, strField As String
    
    On Error Resume Next
    
    strFormula = Trim(strFormula)
    
    If strFormula = "" Then '����
        GetColType = 1
    ElseIf InStr(strFormula, "[") = 0 Then '��������
        varR = Srt.Eval(strFormula)
        If IsNumeric(varR) Then
            GetColType = 2
        ElseIf IsDate(varR) Then
            GetColType = 3
        Else
            GetColType = 1
        End If
    ElseIf Left(strFormula, 1) = "[" And Right(strFormula, 1) = "]" And InStr(strFormula, ".") > 0 _
        And InStr(Mid(strFormula, 2, Len(strFormula) - 2), "[") = 0 Then
         'ֻ���ֶ����õ���
        strFormula = Mid(strFormula, 2, Len(strFormula) - 2)
        strData = Left(strFormula, InStr(strFormula, ".") - 1)
        strField = Mid(strFormula, InStr(strFormula, ".") + 1)
        
        Select Case mLibDatas("_" & strData).DataSet.Fields(strField).type
            Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
                GetColType = 1
            Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                GetColType = 2
            Case adDBTimeStamp, adDBTime, adDBDate, adDate
                GetColType = 3
        End Select
    End If
End Function

Private Function GetParsStr(ByVal objPars As RPTPars) As String
'���ܣ������ظ����ƵĲ������е�����ת��Ϊ�ַ���
'���أ�"������=����ֵ|������=����ֵ..."
'˵�������������ж����ʽ,ͬʱ����"ReportFormat=x"
    Dim tmpPar As RPTPar
    Dim strPars As String
    
    If mobjReport.Fmts.count > 1 Then
        strPars = strPars & "|ReportFormat=" & bytFormat
    End If
    
    For Each tmpPar In objPars
        If tmpPar.ȱʡֵ Like "&*" And tmpPar.���� = 2 Then
            strPars = strPars & "|" & tmpPar.���� & "=" & GetParVBMacro(tmpPar.ȱʡֵ)
        Else
            strPars = strPars & "|" & tmpPar.���� & "=" & tmpPar.ȱʡֵ
        End If
    Next
    GetParsStr = Mid(strPars, 2)
End Function

Private Function SetStrPars(ByVal strPars As String, ByVal objPars As RPTPars) As RPTPars
'���ܣ����ַ����еĲ�������������д������������
'������strPars="������=����ֵ|������=����ֵ..."
'���أ������õĲ�������
'˵���������ǰ�����ж��ָ�ʽ�Ҳ����������и�ʽָ��,���滻
    Dim tmpPar As RPTPar, tmpPars As RPTPars
    Dim i As Integer, j As Integer
    Dim bytTmp As Byte, strTmp As String
    
    If strPars = "" Or Not strPars Like "*=*" Then Set SetStrPars = objPars: Exit Function
    
    Set tmpPars = objPars
    
    For i = 0 To UBound(Split(strPars, "|"))
        For Each tmpPar In tmpPars
            strTmp = Split(strPars, "|")(i)
            If UCase(Split(strTmp, "=")(0)) = UCase("ReportFormat") And mobjReport.Fmts.count > 1 Then
                If IsNumeric(Split(strTmp, "=")(1)) Then
                    bytTmp = CByte(Split(strTmp, "=")(1))
                    For j = 1 To cboFormat.ComboItems.count
                        If CByte(Mid(cboFormat.ComboItems(j).Key, 2)) = bytTmp Then
                            cboFormat.ComboItems(j).Selected = True
                            bytFormat = bytTmp: mobjReport.bytFormat = bytFormat: Exit For
                        End If
                    Next
                End If
            ElseIf UCase(tmpPar.����) = UCase(Split(strTmp, "=")(0)) Then
                Select Case tmpPar.����
                    Case 1 '������
                        If IsNumeric(Split(strTmp, "=")(1)) Then tmpPar.ȱʡֵ = Split(strTmp, "=")(1)
                    Case 2 '������
                        If IsDate(Split(strTmp, "=")(1)) Then tmpPar.ȱʡֵ = Split(strTmp, "=")(1)
                    Case Else
                        tmpPar.ȱʡֵ = Split(strTmp, "=")(1)
                End Select
            End If
        Next
    Next
    Set SetStrPars = tmpPars
End Function

Private Sub mnuFile_Excel_Click()
    Dim lngRow As Long, lngCol As Long
    Dim bytKind As Byte, tmpMsh As Object
    Dim i As Long, j As Long
    
    '��Ҫ�������
    If Not mobjReport.blnLoad Then Exit Sub
    
    If zlRegInfo("��Ȩ����") <> "1" Then
        MsgBox "���û���԰汾����ʹ�øù��ܡ�", vbInformation, App.Title
        Exit Sub
    End If
    
    If isExporting Then
        gblnError = True
        MsgBox "����һ�ű������������ Excel,���Ժ���ִ�иò�����", vbInformation, App.Title
        Exit Sub
    End If
    
    If intGridCount = 0 Then
        MsgBox "������û�����ݱ����������� Excel��", vbInformation, App.Title
        Exit Sub
    End If
    If objCurGrid Is Nothing Then
        If msh.count > 1 Then
            For Each tmpMsh In msh
                If tmpMsh.Index <> 0 And (tmpMsh.Container Is picPaper(intReport) Or UCase(tmpMsh.Container.name) = "PIC") And Not tmpMsh.Tag Like "H_*" Then
                    Set objCurGrid = tmpMsh
                    Exit For
                End If
            Next
        End If
        If objCurGrid Is Nothing Then
            MsgBox "����ѡ��һ��Ҫ����� Excel�����ݱ�", vbInformation, App.Title
            Exit Sub
        End If
    End If
    
    If Not HaveExcel Then
        gblnError = True
        MsgBox "ϵͳ��⵽����û�а�װ Microsoft Excel ����,�������ܼ�����", vbInformation, App.Title
        Exit Sub
    End If
    
    'ȷ��[��ͷ]������
    Set gobjHead = Nothing
    Set gobjBody = Nothing
    
    Set gobjBody = objCurGrid
    If Val(objCurGrid.Tag) > 0 Then
        bytKind = GetGridStyle(mobjReport, objCurGrid.Index)
        If bytKind <> 2 Then Set gobjHead = msh(CInt(objCurGrid.Tag))
    End If
    
    '�������ӱ�ǩ��Ŀ
    Call MakeAppend(Me, picPaper(intReport))
    
    '�����Excel
    lngRow = gobjBody.Row
    lngCol = gobjBody.Col
    If Not gobjHead Is Nothing Then gobjHead.Redraw = False
    gobjBody.Redraw = False
    
    '�滻�س����з�
    If Not gobjHead Is Nothing Then
        For i = 0 To gobjHead.Rows - 1
            For j = 0 To gobjHead.Cols - 1
                gobjHead.TextMatrix(i, j) = Replace(Replace(Replace(gobjHead.TextMatrix(i, j), vbCrLf, "<���зָ���>"), vbLf, "<���зָ���>"), vbCr, "<���зָ���>")
            Next
        Next
    End If
    For i = 0 To gobjBody.Rows - 1
        For j = 0 To gobjBody.Cols - 1
            gobjBody.TextMatrix(i, j) = Replace(Replace(Replace(gobjBody.TextMatrix(i, j), vbCrLf, "<���зָ���>"), vbLf, "<���зָ���>"), vbCr, "<���зָ���>")
        Next
    Next
    
    blnExcel = True
    Call ExportExcel(Me, IIF(mbytStyle = 3, mstrExcelFile, ""))
    
    gobjBody.Row = lngRow
    gobjBody.Col = lngCol
    Call msh_EnterCell(gobjBody.Index)
    If Not gobjHead Is Nothing Then
        gobjHead.Redraw = True
    
        '�ָ��س����з�
        For i = 0 To gobjHead.Rows - 1
            For j = 0 To gobjHead.Cols - 1
                gobjHead.TextMatrix(i, j) = Replace(gobjHead.TextMatrix(i, j), "<���зָ���>", vbCrLf)
            Next
        Next
    End If
    For i = 0 To gobjBody.Rows - 1
        For j = 0 To gobjBody.Cols - 1
            gobjBody.TextMatrix(i, j) = Replace(gobjBody.TextMatrix(i, j), "<���зָ���>", vbCrLf)
        Next
    Next
    gobjBody.Redraw = True
    
End Sub

Public Function DelUnUseData(objReport As Report) As Boolean
'���ܣ��Ӷ���mobjReport��ɾ��δʹ�õ�����Դ����
'���أ��Ƿ����δʹ�õ�����Դ
'˵����1.�ú���ֻ�ڴ򿪱���ǰ����һ��
'      2.��ɾ�����б����ʽ��δʹ�õġ�
    Dim tmpData As RPTData
    
    If objReport Is Nothing Then Exit Function
    
    For Each tmpData In objReport.Datas
        If Not DataUsed(objReport, tmpData.����) Then objReport.Datas.Remove "_" & tmpData.Key
    Next
End Function

Private Function GetStatText(strStat As String) As String
    Select Case strStat
        Case "SUM"
            GetStatText = "�ϼ�"
        Case "AVG"
            GetStatText = "ƽ��ֵ"
        Case "MAX"
            GetStatText = "���ֵ"
        Case "MIN"
            GetStatText = "��Сֵ"
        Case "COUNT"
            GetStatText = "��¼��"
    End Select
End Function

Public Sub AddCol(msh As Object, Optional ByVal intCol As Integer = -1, Optional ByVal intCols As Integer = 1)
'���ܣ���ָ�����msh�в���intCols����,����ĵ�һ�е��к�ΪintCol,���û��intCol����,��׷����
'˵���������к�,ֻ�����ݴ���,�Ը�ʽ��������(�����п�)
    Dim i As Integer, j As Integer, k As Integer
    
    If intCol >= msh.Cols Then intCol = -1
    msh.Cols = msh.Cols + intCols
    If intCol = -1 Then Exit Sub
    '�ƶ�����
    For j = msh.Cols - 1 To intCol + intCols Step -intCols
        For i = 0 To msh.FixedRows - 1
            For k = 0 To intCols - 1
                msh.TextMatrix(i, j - k) = msh.TextMatrix(i, j - k - intCols)
                msh.ColData(j - k) = msh.ColData(j - k - intCols)
            Next
        Next
    Next
    '�����������
    For j = intCol To intCol + intCols - 1
        For i = 0 To msh.FixedRows - 1
            msh.TextMatrix(i, j) = ""
        Next
    Next
End Sub

Private Sub ShowFreeGrid(objItem As RPTItem)
'���ܣ��ڲ�ѯ��������֯��ʾһ��������
    Dim strData As String, strTmp As String, bytKind As Byte
    Dim lngCol As Long, strState As String, arrState() As Variant
    Dim mshBody As Object, mshHead As Object
    Dim tmpItem As RPTItem, tmpID As RelatID, objBody As RPTItem
    Dim strValue As String, lngHead As Long, arrHead() As String
    Dim strSource As String, arrSource() As String, arrFormula() As String, arrField() As String
    Dim strFormula As String, strFormat As String, arrType() As Long
    Dim i As Long, j As Long, k As Long, l As Long, blnDo As Boolean
    Dim arrRowIDs() As Variant, strIDSource As String, strFirstSource As String
    Dim objPic As StdPicture
    Dim objColProtertys As RPTColProtertys
    Dim varIFValue As Variant
    Dim blntmp As Boolean, blnRPTLink As Boolean
    
    arrRowIDs = Array()
    On Error GoTo hErr
    
    With objItem
        bytKind = GetGridStyle(mobjReport, .id)
        Load msh(.id) '���岿��
        Load msh(.SubIDs(1).id) '��ͷ����
        Set msh(.id).Container = picPaper(intReport)
        Set msh(.SubIDs(1).id).Container = picPaper(intReport)
        If .��ID <> 0 Then
            Set msh(.id).Container = pic(.��ID)
            Set msh(.SubIDs(1).id).Container = pic(.��ID)
        End If
        Set mshBody = msh(.id)
        Set mshHead = msh(.SubIDs(1).id) '���õ�һ�е�ID��Ϊ�ؼ�����
        
        mshBody.Redraw = False
        mshHead.Redraw = False
                            
        mshHead.Tag = "H_" & mshBody.Index '��־�ñ��Ϊ�̶���ͷ
        mshBody.Tag = mshHead.Index
        
        '�������
        '��ͷ
        mshHead.Left = .X: mshHead.Top = .Y
        mshHead.Width = .W: mshHead.Height = .H 'Ϊ��ʹ��ͷ�ɹ���
        
        mshHead.Cols = .SubIDs.count
        mshHead.FixedCols = 0
        mshHead.Rows = UBound(Split(mobjReport.Items("_" & .SubIDs(1).id).��ͷ, "|")) + 2
        mshHead.RowHeight(mshHead.Rows - 1) = 0
        mshHead.FixedRows = mshHead.Rows - 1
        
        mshHead.ForeColor = .ǰ��
        mshHead.ForeColorFixed = .ǰ��
        mshHead.BackColor = .����
        mshHead.BackColorFixed = .����
        mshHead.GridColor = .����
        mshHead.GridColorFixed = IIF(.��ʽ = "", .����, Val(.��ʽ))
        mshHead.Font.name = .����
        mshHead.Font.Size = .�ֺ�
        mshHead.Font.Bold = .����
        mshHead.Font.Italic = .б��
        mshHead.Font.Underline = .����
        mshHead.GridLineWidth = IIF(.����߼Ӵ�, 2, 1)
        'Set mshHead.FontFixed = mshHead.Font
        '֧������
        mshHead.ExplorerBar = flexExSortShow

        '�����ͷ����(��Ԫ���롢�иߡ�����,�п�)
        For Each tmpID In .SubIDs
            Set tmpItem = mobjReport.Items("_" & tmpID.id)
            If tmpItem.Relations.count > 0 Then
                If blnRPTLink = False Then blnRPTLink = True
            End If
            arrHead = Split(tmpItem.��ͷ, "|")
            lngHead = 0 '����ͷ���ݸ߶�
            For i = 0 To UBound(arrHead) '����^�߶�^����
                mshHead.Col = tmpItem.���: mshHead.Row = i
                mshHead.CellAlignment = CInt(Split(arrHead(i), "^")(0))
                
                mshHead.RowHeight(i) = CLng(Split(arrHead(i), "^")(1))
                lngHead = lngHead + mshHead.RowHeight(i)
                
                If CStr(Split(arrHead(i), "^")(2)) = "#" Then 'Ϊ��
                    mshHead.TextMatrix(i, tmpItem.���) = ""
                ElseIf CStr(Split(arrHead(i), "^")(2)) = "��" Then '����ߵ�Ԫ����ͬ
                    mshHead.TextMatrix(i, tmpItem.���) = mshHead.TextMatrix(i, tmpItem.��� - 1)
                ElseIf CStr(Split(arrHead(i), "^")(2)) = "��" Then '���ϱߵ�Ԫ����ͬ
                    mshHead.TextMatrix(i, tmpItem.���) = mshHead.TextMatrix(i - 1, tmpItem.���)
                Else
                    strValue = CStr(Split(arrHead(i), "^")(2))
                    
                    '����ָ�븴λ(�����õ��������Դ������ֶ�)
                    '�ȴ��������ֶ�(��ѯʱֻȡ��һ��ֵ)
                    strData = GetLabelDataName(strValue)
                    If strData <> "" Then
                        For j = 0 To UBound(Split(strData, "|"))
                            strTmp = Split(Split(strData, "|")(j), ".")(0)
                            If mLibDatas("_" & strTmp).DataSet.RecordCount > 0 Then
                                mLibDatas("_" & strTmp).DataSet.MoveFirst
                            End If
                            strTmp = GetFieldValue(Me, CStr(Split(strData, "|")(j)))
                            strValue = Replace(strValue, "[" & Split(strData, "|")(j) & "]", strTmp)
                        Next
                    End If
                    
                    '�ٴ��������:[=������]��[n>=0]��[���ڸ�ʽ��][��λ����]
                    strValue = GetLabelMacro(Me, strValue)
                    
                    mshHead.TextMatrix(i, tmpItem.���) = strValue
                    
                End If
                If UBound(Split(arrHead(i), "^")) > 3 Then
                    '�����ͷ��ɫ���Ӵ�
                    If Split(arrHead(i), "^")(3) = 1 Then
                        mshHead.Cell(flexcpFontBold, i, tmpItem.���) = True
                    End If
                    mshHead.Cell(flexcpForeColor, i, tmpItem.���) = Val(Split(arrHead(i), "^")(4))
                End If
            Next
            mshHead.ColWidth(tmpItem.���) = tmpItem.W
        Next
        
        '��ͷ����ϲ�
        For i = 0 To mshHead.FixedRows - 1
            mshHead.MergeRow(i) = True
        Next
        For i = 0 To mshHead.Cols - 1
            mshHead.MergeCol(i) = True
        Next
        
        '�����ʽ
        If bytKind = 2 Then '���б���
            mshBody.Top = .Y: mshBody.Left = .X
            mshBody.Height = .H: mshBody.Width = .W
        Else
            mshBody.Top = .Y + lngHead: mshBody.Left = .X
            If .H - lngHead + 15 < 0 Then
                mshBody.Height = 0
            Else
                mshBody.Height = .H - lngHead + 15
            End If
            mshBody.Width = .W
        End If
        mshBody.Cols = .SubIDs.count: mshBody.FixedCols = 0
        mshBody.Rows = 1: mshBody.FixedRows = 0 '������ʱֻ��һ����
        mshBody.RowHeight(0) = .�и�
        mshBody.RowHeightMin = .�и�
        
        mshBody.ForeColor = .ǰ��
        mshBody.ForeColorFixed = .ǰ��
        mshBody.BackColor = .����
        mshBody.BackColorFixed = .����
        mshBody.GridColor = .����
        mshBody.GridColorFixed = .����
        mshBody.Font.name = .����
        mshBody.Font.Size = .�ֺ�
        mshBody.Font.Bold = .����
        mshBody.Font.Italic = .б��
        mshBody.Font.Underline = .����
        mshBody.GridLineWidth = IIF(.����߼Ӵ�, 2, 1)

        'Set mshBody.FontFixed = mshBody.Font
        
        '��������(�п��ж���)
        For Each tmpID In .SubIDs
            Set tmpItem = mobjReport.Items("_" & tmpID.id)
            With mshBody
                .ColData(tmpItem.���) = tmpItem
                .ColWidth(tmpItem.���) = tmpItem.W
                .ColAlignment(tmpItem.���) = Switch(tmpItem.���� = Val("0-��"), flexAlignLeftCenter _
                                                   , tmpItem.���� = Val("1-��"), flexAlignCenterCenter _
                                                   , tmpItem.���� = Val("2-��"), flexAlignRightCenter)
                If .FixedRows - 1 >= 0 And .Rows - 1 >= 0 Then
                    .Cell(flexcpAlignment, .FixedRows - 1, tmpItem.���, .Rows - 1, tmpItem.���) = .ColAlignment(tmpItem.���)
                End If
                .MergeCol(tmpItem.���) = tmpItem.�Ե�
            End With
        Next
        
        '--------------------------------------------------------------------------------------
        '�����������
        '--------------------------------------------------------------------------------------
        '1.�����ñ���õ�������Դ
        strSource = GetGridSource(objItem) '"������Ϣ,ҩƷ��Ϣ,..."
        
        '2.�Ա��й�ʽ������ϵ����
        arrFormula = SortFormula(objItem) '(����Ԫ��="��ʽ|��ʽ|�к�|����")
        
        '3.��ʼͳ������
        ReDim arrState(.SubIDs.count - 1)
        ReDim arrType(.SubIDs.count - 1) '������������(0=��ȷ��,1-�ַ�(����),2-����,3-����)
        If strSource <> "" Then
            arrSource = Split(strSource, ",")
            strFirstSource = arrSource(0)
            ''��һ��ʱ�ж��Ƿ�������:ֻҪһ��������,�򲻽���
            blnDo = False
            For i = 0 To UBound(arrSource)
                If mLibDatas("_" & arrSource(i)).DataSet.RecordCount > 0 Then
                    mLibDatas("_" & arrSource(i)).DataSet.MoveFirst 'ָ�븴λ
                End If
                blnDo = blnDo Or Not mLibDatas("_" & arrSource(i)).DataSet.EOF
                
                'ȷ����ID�ֶε�����Դ:�Ե�һ��Ϊ׼
                If strIDSource = "" Then
                    For j = 0 To mLibDatas("_" & arrSource(i)).DataSet.Fields.count - 1
                        If UCase(mLibDatas("_" & arrSource(i)).DataSet.Fields(j).name) = "ID" Then
                            If IsType(mLibDatas("_" & arrSource(i)).DataSet.Fields(j).type, adNumeric) Then
                                strIDSource = arrSource(i): Exit For
                            End If
                        End If
                    Next
                End If
            Next

        Else
            blnDo = True
        End If
        
        mshHead.WordWrap = .�Ե�
        mshBody.WordWrap = .�Ե�
        
        '����Ӧ�и�
        For Each tmpID In .SubIDs
            Set tmpItem = mobjReport.Items("_" & tmpID.id)
            If Not tmpItem Is Nothing Then
                If tmpItem.����Ӧ�и� Then
                    '��һ����������Ӧ�иߣ������������Ҫ���ø�����
                    mshBody.AutoSizeMode = flexAutoSizeRowHeight
                End If
            End If
        Next
        
        '4.��֯����
        j = 0
        Do While blnDo
            If j > 0 Then
                mshBody.Rows = mshBody.Rows + 1 'ȱʡ��һ��
                mshBody.RowHeight(mshBody.Rows - 1) = .�и�
            End If
            
            '�Ա���Ӧ��ID��������и�ֵ
            ReDim Preserve arrRowIDs(UBound(arrRowIDs) + 1)
            arrRowIDs(UBound(arrRowIDs)) = 0
            If strIDSource <> "" Then
                If Not mLibDatas("_" & strIDSource).DataSet.EOF Then
                    arrRowIDs(UBound(arrRowIDs)) = Val(Nvl(mLibDatas("_" & strIDSource).DataSet.Fields("ID").Value, 0))
                End If
            End If
            
            For i = 0 To UBound(arrFormula)
                arrField = Split(arrFormula(i), "|")
                strFormula = arrField(0)    '��ʽ
                strFormat = arrField(1)     '��ʽ
                lngCol = Val(arrField(2))   '�к�
                strState = arrField(3)      '����
                
                '���������
                strValue = EvalFormula(strFormula, mshBody.Index, j)
                'If gobjFile.FileExists(strValue) Then   '�÷����ٶ������ر����ڼ�¼�����������ر�����
                If LCase$(Right$(strValue, 4)) = ".pic" Then
                    Set objPic = Nothing
                    On Error Resume Next
                    Set objPic = LoadPicture(strValue)
                    gobjFile.DeleteFile strValue, True
                    On Error GoTo 0
                    
                    If Not objPic Is Nothing Then
                        mshBody.Row = j: mshBody.Col = lngCol
                        
                        Me.picTemp.Cls '����������
                        If objPic.Height / objPic.Width < mshBody.CellHeight / mshBody.CellWidth Then
                            Me.picTemp.Width = mshBody.CellWidth
                            Me.picTemp.Height = (objPic.Height / objPic.Width) * mshBody.CellWidth
                        Else
                            Me.picTemp.Height = mshBody.CellHeight
                            Me.picTemp.Width = (objPic.Width / objPic.Height) * mshBody.CellHeight
                        End If
                        Me.picTemp.PaintPicture objPic, 0, 0, Me.picTemp.Width, Me.picTemp.Height
                                            
                        Set mshBody.CellPicture = Me.picTemp.Image
                        mshBody.CellPictureAlignment = 4 '�̶��ж���
                    End If
                Else
                    mshBody.TextMatrix(j, lngCol) = strValue
                    mshBody.Cell(flexcpFont, j, lngCol) = mshBody.Font          'ǿ�Ƹ���������󣬵�Ԫ������Ӧ�и߲�����Ч
                    If (strIDSource <> "" Or strFirstSource <> "") And blnRPTLink = True Then
'                        Set colRelation = New Collection
'                        colRelation.Add mLibDatas("_" & IIF(strIDSource = "", strFirstSource, strIDSource)).DataSet.AbsolutePosition
'                        mshBody.Cell(flexcpData, j, lngCol) = colRelation
                        
                        '�Ż�
                        '�̶���RowData��ż�¼�����кţ���һ�еĵ�Ԫ���ű������ӹ�ϵ
                        mshBody.RowData(j) = mLibDatas("_" & IIF(strIDSource = "", strFirstSource, strIDSource)).DataSet.AbsolutePosition
                    End If
                    '����������
                    Set objColProtertys = mshBody.ColData(lngCol).ColProtertys
                    If objColProtertys.count > 0 Then
                        For l = 1 To objColProtertys.count
                            If InStr(objColProtertys.Item(l).����ֵ, strIDSource & ".") > 0 Then
                                varIFValue = EvalFormula("[" & objColProtertys.Item(l).����ֵ & "]", mshBody.Index, j)
                            Else
                                varIFValue = objColProtertys.Item(l).����ֵ
                            End If
                            If CheckColProtertys(EvalFormula("[" & objColProtertys.Item(l).�����ֶ� & "]", mshBody.Index, j), objColProtertys.Item(l).������ϵ, varIFValue) Then
                                If objColProtertys.Item(l).�Ƿ�����Ӧ�� Then
                                    mshBody.Cell(flexcpBackColor, j, mshBody.FixedCols, j, mshBody.Cols - 1) = objColProtertys.Item(l).������ɫ
                                    mshBody.Cell(flexcpForeColor, j, mshBody.FixedCols, j, mshBody.Cols - 1) = objColProtertys.Item(l).������ɫ
                                    mshBody.Cell(flexcpFontBold, j, mshBody.FixedCols, j, mshBody.Cols - 1) = objColProtertys.Item(l).�Ƿ�Ӵ�
                                Else
                                    mshBody.Cell(flexcpBackColor, j, lngCol) = objColProtertys.Item(l).������ɫ
                                    mshBody.Cell(flexcpForeColor, j, lngCol) = objColProtertys.Item(l).������ɫ
                                    mshBody.Cell(flexcpFontBold, j, lngCol) = objColProtertys.Item(l).�Ƿ�Ӵ�
                                End If
                                
                                '���뷽ʽ
                                Select Case objColProtertys.Item(l).����
                                Case Val("1-����")
                                    mshBody.Cell(flexcpAlignment, j, lngCol) = flexAlignLeftCenter
                                Case Val("2-����")
                                    mshBody.Cell(flexcpAlignment, j, lngCol) = flexAlignCenterCenter
                                Case Val("3-����")
                                    mshBody.Cell(flexcpAlignment, j, lngCol) = flexAlignRightCenter
                                Case Else
                                    'ȱʡ��������
                                End Select
                            End If
                        Next
                    End If
                End If
                
                '������������
                If j = 0 And (strState = "MAX" Or strState = "MIN") Then
                    arrType(lngCol) = GetColType(strFormula)
                    arrState(lngCol) = "��ʼֵ"
                End If
                If strState = "MAX" Or strState = "MIN" Then
                    If arrType(lngCol) = 0 Then
                        If IsNumeric(mshBody.TextMatrix(j, lngCol)) Then
                            arrType(lngCol) = 2
                        ElseIf IsDate(mshBody.TextMatrix(j, lngCol)) Then
                            arrType(lngCol) = 3
                        Else
                            arrType(lngCol) = 1
                        End If
                    End If
                End If
                
                '��������
                On Error Resume Next
                If mshBody.TextMatrix(j, lngCol) <> "" Then
                    Select Case strState
                        Case "SUM", "AVG" 'ƽ��ֵ�ȼ�(�ٳ�)
                            If IsNumeric(mshBody.TextMatrix(j, lngCol)) Then
                                arrState(lngCol) = arrState(lngCol) + CDbl(mshBody.TextMatrix(j, lngCol))
                            ElseIf IsDate(mshBody.TextMatrix(j, lngCol)) Then
                                arrState(lngCol) = arrState(lngCol) + CDate(mshBody.TextMatrix(j, lngCol))
                            Else
                                arrState(lngCol) = arrState(lngCol) + mshBody.TextMatrix(j, lngCol)
                            End If
                        Case "MAX"
                            If arrState(lngCol) = "��ʼֵ" Then
                                If arrType(lngCol) = 2 Then
                                    arrState(lngCol) = CDbl(mshBody.TextMatrix(j, lngCol))
                                ElseIf arrType(lngCol) = 3 Then
                                    arrState(lngCol) = CDate(mshBody.TextMatrix(j, lngCol))
                                Else
                                    arrState(lngCol) = mshBody.TextMatrix(j, lngCol)
                                End If
                            Else
                                If arrType(lngCol) = 2 Then
                                    If CDbl(mshBody.TextMatrix(j, lngCol)) > arrState(lngCol) Then
                                        arrState(lngCol) = CDbl(mshBody.TextMatrix(j, lngCol))
                                    End If
                                ElseIf arrType(lngCol) = 3 Then
                                    If CDate(mshBody.TextMatrix(j, lngCol)) > arrState(lngCol) Then
                                        arrState(lngCol) = CDate(mshBody.TextMatrix(j, lngCol))
                                    End If
                                Else
                                    If mshBody.TextMatrix(j, lngCol) > arrState(lngCol) Then
                                        arrState(lngCol) = mshBody.TextMatrix(j, lngCol)
                                    End If
                                End If
                            End If
                        Case "MIN"
                            If arrState(lngCol) = "��ʼֵ" Then
                                If arrType(lngCol) = 2 Then
                                    arrState(lngCol) = CDbl(mshBody.TextMatrix(j, lngCol))
                                ElseIf arrType(lngCol) = 3 Then
                                    arrState(lngCol) = CDate(mshBody.TextMatrix(j, lngCol))
                                Else
                                    arrState(lngCol) = mshBody.TextMatrix(j, lngCol)
                                End If
                            Else
                                If arrType(lngCol) = 2 Then
                                    If CDbl(mshBody.TextMatrix(j, lngCol)) < arrState(lngCol) Then
                                        arrState(lngCol) = CDbl(mshBody.TextMatrix(j, lngCol))
                                    End If
                                ElseIf arrType(lngCol) = 3 Then
                                    If CDate(mshBody.TextMatrix(j, lngCol)) < arrState(lngCol) Then
                                        arrState(lngCol) = CDate(mshBody.TextMatrix(j, lngCol))
                                    End If
                                Else
                                    If mshBody.TextMatrix(j, lngCol) < arrState(lngCol) Then
                                        arrState(lngCol) = mshBody.TextMatrix(j, lngCol)
                                    End If
                                End If
                            End If
                        Case "COUNT"
                            arrState(lngCol) = arrState(lngCol) + 1
                    End Select
                End If
                
                If Err.Number <> 0 Then Err.Clear
                On Error GoTo 0
                
                '�Ȼ����ٸ�ʽ��,�������
                If strFormat <> "" Then
                    On Error Resume Next
                    mshBody.TextMatrix(j, lngCol) = Format(mshBody.TextMatrix(j, lngCol), strFormat)
                    If Err.Number <> 0 Then Err.Clear
                    On Error GoTo 0
                End If
            Next
            
            If strSource <> "" Then
                'ֻҪһ��������,�򲻽���
                blnDo = False
                For i = 0 To UBound(arrSource)
                    If Not mLibDatas("_" & arrSource(i)).DataSet.EOF Then
                        mLibDatas("_" & arrSource(i)).DataSet.MoveNext
                    End If
                    blnDo = blnDo Or Not mLibDatas("_" & arrSource(i)).DataSet.EOF
                Next
            Else
                '���û���õ�����Դ,��ֻ��һ������
                blnDo = False
            End If
            
            j = j + 1
        Loop
        
        '5.���������,ֻҪ��һ���л���,����
        blnDo = False
        For i = 0 To UBound(arrFormula)
            blnDo = blnDo Or (Split(arrFormula(i), "|")(3) <> "")
            '�Ż�
            If blnDo Then Exit For
        Next
        If blnDo Then
            mshBody.Rows = mshBody.Rows + 1
            mshBody.RowHeight(mshBody.Rows - 1) = .�и�
            For i = 0 To UBound(arrFormula)
                arrField = Split(arrFormula(i), "|")
                strState = arrField(3)      '����
                lngCol = Val(arrField(2))   '�к�
                strFormat = arrField(1)     '��ʽ
                strFormula = arrField(0)    '��ʽ
                If strState = "AVG" Then
                    On Error Resume Next
                    mshBody.TextMatrix(j, lngCol) = arrState(lngCol) / j
                    If Err.Number <> 0 Then Err.Clear
                    On Error GoTo 0
                ElseIf strState <> "" Then
                    If TypeName(arrState(lngCol)) = "String" Then
                        If arrState(lngCol) = "��ʼֵ" Then arrState(lngCol) = ""
                    End If
                    mshBody.TextMatrix(j, lngCol) = arrState(lngCol)
                ElseIf strFormula <> "" Then
                    '�������У�û�л��ܵ�������й�ʽ������㹫ʽ
                    strValue = EvalFormula(strFormula, mshBody.Index, j)
                    'If gobjFile.FileExists(strValue) Then   '�÷����ٶ������ر����ڼ�¼�����������ر�����
                    If LCase$(Right$(strValue, 4)) = ".pic" Then
                        'ͼƬ�ֶ�
                        On Error Resume Next
                        gobjFile.DeleteFile strValue, True
                        On Error GoTo 0
                    Else
                        mshBody.TextMatrix(j, lngCol) = strValue
                    End If
                End If
                '��ʽ����Ԫֵ
                If strFormat <> "" And mshBody.TextMatrix(j, lngCol) <> "" Then
                    On Error Resume Next
                    mshBody.TextMatrix(j, lngCol) = Format(mshBody.TextMatrix(j, lngCol), strFormat)
                    If Err.Number <> 0 Then Err.Clear
                    On Error GoTo 0
                End If
            Next
            '��ʾ���ܱ�־
            For k = 0 To mshBody.Cols - 1
                If mshBody.ColWidth(k) > 0 Then Exit For
            Next
            If mshBody.TextMatrix(j, k) = "" Then
                blnDo = True: l = 0
                For i = 0 To UBound(arrFormula)
                    arrField = Split(arrFormula(i), "|")
                    If arrField(3) <> "" Then
                        If l = 0 Then
                            strState = arrField(3)
                        Else
                            blnDo = blnDo And (Split(arrFormula(i), "|")(3) = strState)
                        End If
                        l = l + 1
                    End If
                Next
                If blnDo Then 'һ�ֻ��ܷ�ʽ
                    mshBody.TextMatrix(j, k) = Switch(strState = "SUM", "�ϼ�", strState = "AVG", "ƽ��ֵ", strState = "MAX", "���ֵ", strState = "MIN", "��Сֵ", strState = "COUNT", "��¼��")
                Else '���ֻ��ܷ�ʽ
                    mshBody.TextMatrix(j, k) = "����"
                End If
                mshBody.Row = j: mshBody.Col = k: mshBody.CellAlignment = 4
            End If
        End If
        
        For i = 0 To mshBody.Rows - 1
            mshBody.RowHeight(i) = .�и�
        Next

        '��������
        mshHead.ScrollBars = flexScrollBarHorizontal
        mshBody.MergeCells = flexMergeRestrictRows
        mshBody.ScrollBars = flexScrollBarBoth
        mshHead.Row = mshHead.FixedRows
        mshBody.Row = 0: mshBody.Col = 0
        
        '��������(�п��ж���)
        For Each tmpID In .SubIDs
            Set tmpItem = mobjReport.Items("_" & tmpID.id)
            '���ù�����ѯ��������ʽ
            If tmpItem.Relations.count > 0 Then
                For i = 0 To mshBody.Rows - 1
                    '�ϼ��в�����
                    If TypeName(mshBody.RowData(i)) <> "Empty" Then
                        If mshBody.Cell(flexcpForeColor, i, tmpItem.���) = 0 Then
                            mshBody.Cell(flexcpForeColor, i, tmpItem.���) = &HFF0001
                        End If
                        mshBody.Cell(flexcpFontUnderline, i, tmpItem.���) = True
                    End If
                Next
                
                '�Ż�����һ�еĵ�Ԫ���ű������ӹ�ϵ����
                For i = 0 To mshBody.Cols - 1
                    If tmpItem.��� = i Then
                        '����ͬʱȡRelations����
                        mshBody.Cell(flexcpData, 0, i) = tmpItem.Relations
                        Exit For
                    End If
                Next
            End If
            '����û���κ�����ʱ����
            blntmp = False
            If tmpItem.���� = 1 Then
                For i = mshBody.FixedRows To mshBody.Rows - 1
                    If mshBody.TextMatrix(i, tmpItem.���) <> "" Then
                        blntmp = True: Exit For
                    End If
                Next
                If blntmp = False Then
                    mshBody.ColHidden(tmpItem.���) = True
                    mshHead.ColHidden(tmpItem.���) = True
                    mshBody.ColWidth(tmpItem.���) = 0
                    mshHead.ColWidth(tmpItem.���) = 0
                End If
            End If
        Next
        
        '�Զ����������и�
        For Each objBody In mobjReport.Items
            '��ǰ���
            If objBody.���� = Val("4-�����") And mshBody.Index = objBody.id Then
                Call AdjustRowHight(mshBody.Index)
                Exit For
            End If
        Next
        
        If bytKind <> 2 Then '���б���
            mshHead.ZOrder
            mshHead.Visible = True
        End If
        If bytKind <> 1 Then '���б�ͷ
            mshBody.ZOrder
            mshBody.Visible = True
        End If
    End With
    
    mshBody.Redraw = True
    mshHead.Redraw = True
    
    '���������е�����
    If UBound(arrRowIDs) + 1 < mshBody.Rows Then
        ReDim Preserve arrRowIDs(UBound(arrRowIDs) + (mshBody.Rows - (UBound(arrRowIDs) + 1)))
    End If
    mcolRowIDs.Add arrRowIDs, "_" & mshBody.Index
    
    Exit Sub
    
hErr:
    Call ErrCenter
End Sub

Private Sub AdjustRowHight(ByVal Index As Integer)
    Dim intBegin As Integer, intEnd As Integer, i As Integer
    Dim objBody As RPTItem, tmpItem As RPTItem
    
    intBegin = -1
    intEnd = -1
    Set objBody = mobjReport.Items("_" & Index)
    If objBody Is Nothing Then Exit Sub
    
    For i = 1 To objBody.SubIDs.count
        Set tmpItem = mobjReport.Items("_" & objBody.SubIDs(i).id)
        If Not tmpItem Is Nothing Then
            If tmpItem.����Ӧ�и� Then
                If intBegin < 0 Then intBegin = tmpItem.���
                intEnd = tmpItem.���
            End If
        End If
    Next
    If intBegin >= 0 Then
        '�����и�
        msh(objBody.id).AutoSize intBegin, intEnd
    End If
End Sub

Private Function CheckColProtertys(ByVal var�����ֶ� As Variant, ByVal str������ϵ As String, ByVal var����ֵ As Variant) As Boolean
'���ܣ����ݴ���������������ж��Ƿ�����,��Ϊ�գ�������ִ��
    
    Select Case str������ϵ
        Case ""
            CheckColProtertys = True
        Case "����"
            If IsNumeric(var�����ֶ�) Then var����ֵ = ValEx(var����ֵ): var�����ֶ� = ValEx(var�����ֶ�)
            CheckColProtertys = (var�����ֶ� = var����ֵ)
        Case "����"
            var����ֵ = ValEx(var����ֵ)
            var�����ֶ� = ValEx(var�����ֶ�)
            CheckColProtertys = (var�����ֶ� > var����ֵ)
        Case "С��"
            var����ֵ = ValEx(var����ֵ)
            var�����ֶ� = ValEx(var�����ֶ�)
            CheckColProtertys = (var�����ֶ� < var����ֵ)
        Case "������"
            var����ֵ = ValEx(var����ֵ)
            var�����ֶ� = ValEx(var�����ֶ�)
            CheckColProtertys = (var�����ֶ� <> var����ֵ)
        Case "���ڵ���"
            var����ֵ = ValEx(var����ֵ)
            var�����ֶ� = ValEx(var�����ֶ�)
            CheckColProtertys = (var�����ֶ� >= var����ֵ)
        Case "С�ڵ���"
            var����ֵ = ValEx(var����ֵ)
            var�����ֶ� = ValEx(var�����ֶ�)
            CheckColProtertys = (var�����ֶ� <= var����ֵ)
        Case "��ƥ��"
            If var�����ֶ� <> "" And var����ֵ <> "" Then
                CheckColProtertys = (var�����ֶ� Like var����ֵ & "*")
            End If
        Case "˫��ƥ��"
            If var�����ֶ� <> "" And var����ֵ <> "" Then
                CheckColProtertys = (var�����ֶ� Like "*" & var����ֵ & "*")
            End If
    End Select
End Function

Private Sub ShowItems()
    Dim i As Integer, lngW As Long, lngH As Long
    Dim objItem As RPTItem, objLoad As Object
    Dim strData As String, strFormat As String, strTmp As String
    Dim strValue As String, objPic As StdPicture
    Dim objFmt As RPTFmt, objFont As StdFont
    Dim lngSize As Long, sngWidth As Single
    Dim lngRec As Long
    Dim objRotate As clsRotateFont
    
    On Error GoTo errH
    blnRefresh = False
    If mbytStyle = 0 Or mbytStyle = 1 Then ShowFlash "������֯��������,���Ժ򣮣���", , Me

    LockWindowUpdate Me.hwnd
    
    Set mcolRowIDs = New Collection
    For Each objLoad In msh
        If objLoad.Index <> 0 Then
            If objLoad.Container Is picPaper(intReport) Then
                Unload objLoad
            ElseIf UCase(objLoad.Container.name) = "PIC" Then
                If objLoad.Container.Container Is picPaper(intReport) Then
                    Unload objLoad
                End If
            End If
        End If
    Next
    For Each objLoad In lbl
        If objLoad.Index <> 0 Then
            If objLoad.Container Is picPaper(intReport) Then
                Unload objLoad
            ElseIf UCase(objLoad.Container.name) = "PIC" Then
                If objLoad.Container.Container Is picPaper(intReport) Then
                    Unload objLoad
                End If
            End If
        End If
    Next
    For Each objLoad In img
        If objLoad.Index <> 0 Then
            If objLoad.Container Is picPaper(intReport) Then
                Unload objLoad
            ElseIf UCase(objLoad.Container.name) = "PIC" Then
                If objLoad.Container.Container Is picPaper(intReport) Then
                    Unload objLoad
                End If
            End If
        End If
    Next
    For Each objLoad In imgCode
        If objLoad.Index <> 0 Then
            If objLoad.Container Is picPaper(intReport) Then
                Unload objLoad
            ElseIf UCase(objLoad.Container.name) = "PIC" Then
                If objLoad.Container.Container Is picPaper(intReport) Then
                    Unload objLoad
                End If
            End If
        End If
    Next
    For Each objLoad In lin
        If objLoad.Index <> 0 Then
            If objLoad.Container Is picPaper(intReport) Then
                Unload objLoad
            ElseIf UCase(objLoad.Container.name) = "PIC" Then
                If objLoad.Container.Container Is picPaper(intReport) Then
                    Unload objLoad
                End If
            End If
        End If
    Next
    For Each objLoad In Shp
        If objLoad.Index <> 0 Then
            If objLoad.Container Is picPaper(intReport) Then
                Unload objLoad
            ElseIf UCase(objLoad.Container.name) = "PIC" Then
                If objLoad.Container.Container Is picPaper(intReport) Then
                    Unload objLoad
                End If
            End If
        End If
    Next
    For Each objLoad In Chart
        If objLoad.Index <> 0 Then
            If objLoad.Container Is picPaper(intReport) Then
                Unload objLoad
            ElseIf UCase(objLoad.Container.name) = "PIC" Then
                If objLoad.Container.Container Is picPaper(intReport) Then
                    Unload objLoad
                End If
            End If
        End If
    Next
    For Each objLoad In pic
        If objLoad.Index <> 0 And objLoad.Container Is picPaper(intReport) Then Unload objLoad
    Next
    
    picPaper(intReport).Cls
    intGridCount = 0
    intGridID = 0
    Set objCurGrid = Nothing
    
    Set objFmt = mobjReport.Fmts("_" & mobjReport.bytFormat)
    If objFmt.ֽ�� = 1 Then
        lngW = objFmt.W
        lngH = objFmt.H
    Else
        lngW = objFmt.H
        lngH = objFmt.W
    End If
    
    '�ȴ�����
    For Each objItem In mobjReport.Items
        '����Ϊ2��Ϊ�����ӱ��,�ڴ�������ձ���ͬʱ����
        If objItem.��ʽ�� = bytFormat Then
            With objItem
                If .���� = Val("14-��ƬԪ��") Then
                    Load pic(.id)
                    Set pic(.id).Container = picPaper(intReport)
                    Set objLoad = pic(.id)
                    .���� = "111"
                    objLoad.Left = .X
                    objLoad.Top = .Y
                    
                    objLoad.Height = IIF(.H > lngH, lngH, .H)
                    objLoad.Width = IIF(.W > lngW, lngW, .W)
                    objLoad.BorderStyle = IIF(.�߿�, 1, 0)
                    
                    objLoad.ZOrder
                    objLoad.Visible = True
                End If
            End With
        End If
    Next
    
    For Each objItem In mobjReport.Items
        '����Ϊ2��Ϊ�����ӱ��,�ڴ�������ձ���ͬʱ����
        If objItem.��ʽ�� = bytFormat Then
            With objItem
                Select Case .����
                    Case 1 '����
                        Load lin(.id)
                        Set lin(.id).Container = picPaper(intReport)
                        If .��ID <> 0 Then
                            Set lin(.id).Container = pic(.��ID)
                        End If
                        Set objLoad = lin(.id)
                        objLoad.X1 = .X
                        objLoad.X2 = IIF(.X + .W - IIF(.W > 0, Screen.TwipsPerPixelX, 0) > lngW, lngW, .X + .W - IIF(.W > 0, Screen.TwipsPerPixelX, 0))
                        objLoad.Y1 = .Y
                        objLoad.Y2 = IIF(.Y + .H - IIF(.H > 0, Screen.TwipsPerPixelY, 0) > lngH, lngH, .Y + .H - IIF(.H > 0, Screen.TwipsPerPixelY, 0))
                        objLoad.BorderColor = .ǰ��
                        If .���� Then objLoad.BorderWidth = 2
                        objLoad.ZOrder
                        objLoad.Visible = True
                    Case 10 '����
                        Load Shp(.id)
                        Set Shp(.id).Container = picPaper(intReport)
                        If .��ID <> 0 Then
                            Set Shp(.id).Container = pic(.��ID)
                        End If
                        Set objLoad = Shp(.id)
                        objLoad.Left = .X
                        objLoad.Top = .Y
                        objLoad.Height = IIF(.H > lngH, lngH, .H)
                        objLoad.Width = IIF(.W > lngW, lngW, .W)
                        objLoad.BorderColor = 0
                        If .���� Then objLoad.BorderWidth = 2
                        objLoad.Shape = IIF(.�߿�, ShapeConstants.vbShapeOval, ShapeConstants.vbShapeRectangle)
                        objLoad.ZOrder
                        objLoad.Visible = True
                    Case 11 'ͼƬ
                        Load img(.id)
                        Set img(.id).Container = picPaper(intReport)
                        If .��ID <> 0 Then
                            Set img(.id).Container = pic(.��ID)
                        End If
                        Set objLoad = img(.id)
                        objLoad.Left = .X
                        objLoad.Top = .Y
                        
                        Set objPic = LoadPictureFromPar(Me, .����)
                        If objPic Is Nothing Then Set objPic = .ͼƬ
                        If .�Ե� And Not objPic Is Nothing Then
                            .W = objPic.Width * (15 / 26.46)
                            .H = objPic.Height * (15 / 26.46)
                        End If
                        objLoad.Height = IIF(.H > lngH, lngH, .H)
                        objLoad.Width = IIF(.W > lngW, lngW, .W)
                        objLoad.BorderStyle = IIF(.�߿�, 1, 0)
                        
                        '���ֱ���
                        If Not objPic Is Nothing Then
                            If .���� Then
                                Set objLoad.Picture = ScalePicture(picTemp, objPic, objLoad.Width, objLoad.Height)
                            Else
                                Set objLoad.Picture = objPic
                            End If
                        End If
                        
                        objLoad.ZOrder
                        objLoad.Visible = True
                    Case 2, 3 '��ǩ,��ǩ��ͼƬ
                        strValue = .����
                        
                        '����ָ�븴λ(�����õ��������Դ������ֶ�)
                        strData = GetLabelDataName(strValue)
                        If strData <> "" Then
                            For i = 0 To UBound(Split(strData, "|"))
                                strTmp = Split(Split(strData, "|")(i), ".")(0)
                                If mLibDatas("_" & strTmp).DataSet.RecordCount > 0 Then
                                    If Val(.Դ�к� & "") <> 0 Then
                                        If mLibDatas("_" & strTmp).DataSet.RecordCount >= Val(.Դ�к� & "") Then
                                            mLibDatas("_" & strTmp).DataSet.AbsolutePosition = Val(.Դ�к� & "")
                                        Else
                                            mLibDatas("_" & strTmp).DataSet.MoveFirst
                                        End If
                                    Else
                                        mLibDatas("_" & strTmp).DataSet.MoveFirst
                                    End If
                                End If
                                
                                '�ȴ��������ֶ�(��ѯʱֻȡ��һ��ֵ)
                                strFormat = GetFieldValue(Me, CStr(Split(strData, "|")(i)))
                                If .��ʽ <> "" Then
                                    On Error Resume Next
                                    strFormat = Format(strFormat, .��ʽ)
                                    If Err.Number <> 0 Then Err.Clear
                                    On Error GoTo errH
                                End If
                                strValue = Replace(strValue, "[" & Split(strData, "|")(i) & "]", strFormat)
                            Next
                            lngRec = mLibDatas("_" & strTmp).DataSet.AbsolutePosition
                        End If
                        
                        '�ٴ��������:[=������]��[n>=0]��[���ڸ�ʽ��]��[��λ����]
                        strValue = GetLabelMacro(Me, strValue)
                        
                        If gobjFile.FileExists(strValue) Then
                            '�������ֶε���ͼ��
                            On Error Resume Next
                            Set .ͼƬ = LoadPicture(strValue)
                            If .ͼƬ Is Nothing Then Set .ͼƬ = New StdPicture '�Դ�������ͼƬ��������
                            Kill strValue
                            Err.Clear
                            On Error GoTo errH
                            
                            If .�Ե� Then
                                .W = .ͼƬ.Width * (15 / 26.46)
                                .H = .ͼƬ.Height * (15 / 26.46)
                            End If
                            
                            Load img(.id)
                            Set img(.id).Container = picPaper(intReport)
                            Set objLoad = img(.id)
                            objLoad.BorderStyle = IIF(.�߿�, 1, 0)
                            
                            '���ֱ���
                            If .���� Then
                                Set objLoad.Picture = ScalePicture(picTemp, .ͼƬ, objLoad.Width, objLoad.Height)
                            Else
                                Set objLoad.Picture = .ͼƬ
                            End If
                        Else
                            Set .ͼƬ = Nothing '�Դ�������ͼƬ��������
                            
                            If .�Ե� Then Call ItemAutoSize(objItem, strValue, picBack)
                            
                            Load lbl(.id)
                            Set lbl(.id).Container = picPaper(intReport)
                            If .��ID <> 0 Then
                                Set lbl(.id).Container = pic(.��ID)
                            End If
                            Set objLoad = lbl(.id)
                            
                            objLoad.FontName = .����
                            objLoad.FontSize = .�ֺ�
                            objLoad.FontBold = .����
                            objLoad.FontItalic = .б��
                            objLoad.FontUnderline = .����
                            
                            objLoad.Alignment = IIF(.���� = 2, 1, IIF(.���� = 1, 2, 0))
                            objLoad.BorderStyle = IIF(.�߿�, 1, 0)
                            objLoad.ForeColor = .ǰ��
                            objLoad.BackColor = .����
                            objLoad.Caption = strValue
                            '���ó���������
                            If objItem.Relations.count > 0 Then
                                objLoad.ForeColor = &HFF0001
                                objLoad.FontUnderline = True
                                objLoad.Tag = lngRec
                            End If
                        End If
                        objLoad.Left = .X
                        objLoad.Top = .Y
                        objLoad.Width = IIF(.W > lngW, lngW, .W)
                        objLoad.Height = IIF(.H > lngH, lngH, .H)
                        
                        If .ˮƽ��ת Then
                            Set objRotate = New clsRotateFont
                            
                            Load picRotate(.id)
                            With picRotate(.id)
                                Set .Font = objLoad.Font
                                If objItem.��ID = 0 Then
                                    Set .Container = picPaper(intReport)
                                Else
                                    Set .Container = objLoad.Container      '��Ƭ�ڵı�ǩ
                                End If
                                .AutoRedraw = True
                                .Left = objLoad.Left
                                .Top = objLoad.Top
                                .Width = objLoad.Width
                                .Height = objLoad.Height
                                .BackColor = objLoad.BackColor
                                If objItem.�߿� Then
                                    picRotate(objItem.id).Line (0, 0)-(objLoad.Width - 15, objLoad.Height - 15), , B
                                End If
                                .ForeColor = objLoad.ForeColor
                                .ZOrder
                                .Visible = True
                            End With
                            
                            Set objRotate.LogFont = picRotate(.id).Font
                            objRotate.OutputReverse picRotate(.id), lbl(.id).Caption, .����
                        Else
                            objLoad.ZOrder
                            objLoad.Visible = True
                        End If
                    Case 4 '������(������Ϊ6������)
                        If objItem.���� = 0 Then
                            If .��ID = 0 Then
                                '��Ƭ�еı���������
                                intGridCount = intGridCount + 1
                                intGridID = objItem.id
                            End If
                        End If
                        Call ShowFreeGrid(objItem)
                    Case 5 '������(������Ϊ7,8,9������)
                        If objItem.���� = 0 Then
                            intGridCount = intGridCount + 1
                            intGridID = objItem.id
                            Call ShowStatGrid(objItem)
                        End If
                    Case 12 'ͼ��@@@
                        Load Chart(.id)
                        Set Chart(.id).Container = picPaper(intReport)
                        If .��ID <> 0 Then
                            Set Chart(.id).Container = pic(.��ID)
                        End If
                        Set objLoad = Chart(.id)
                        
                        objLoad.Left = .X
                        objLoad.Top = .Y
                        objLoad.Height = IIF(.H > lngH, lngH, .H)
                        objLoad.Width = IIF(.W > lngW, lngW, .W)
                                                                
                        strTmp = GetChartFileFromPar(Me, .����)
                        If strTmp <> "" Then
                            Call objLoad.Load(strTmp)
                            objLoad.Height = IIF(.H > lngH, lngH, .H)
                            objLoad.Width = IIF(.W > lngW, lngW, .W)
                        Else
                            If objItem.���� <> "" Then
                                Call GetChartDataName(objItem.����, , , , strTmp)
                            End If
                            If strTmp <> "" Then
                                Call SetChartStyleAndData(objLoad, objItem, mLibDatas("_" & strTmp).DataSet)
                            Else
                                Call SetChartStyleAndData(objLoad, objItem, , , , True)
                            End If
                        End If
                        
                        objLoad.ZOrder
                        objLoad.Visible = True
                    Case 13 '����
                        Load imgCode(.id)
                        Set imgCode(.id).Container = picPaper(intReport)
                        If .��ID <> 0 Then
                            Set imgCode(.id).Container = pic(.��ID)
                        End If
                        Set objLoad = imgCode(.id)
                        
                        objLoad.Left = .X
                        objLoad.Top = .Y
                        objLoad.Height = IIF(.H > lngH, lngH, .H)
                        objLoad.Width = IIF(.W > lngW, lngW, .W)
                        objLoad.BorderStyle = 0
                        
                        '��ȡ��������
                        strValue = .����
                        
                        '����ָ�븴λ(�����õ��������Դ������ֶ�)
                        strData = GetLabelDataName(strValue)
                        If strData <> "" Then
                            For i = 0 To UBound(Split(strData, "|"))
                                strTmp = Split(Split(strData, "|")(i), ".")(0)
                                If mLibDatas("_" & strTmp).DataSet.RecordCount > 0 Then
                                    mLibDatas("_" & strTmp).DataSet.MoveFirst
                                End If
                            Next
                        End If
                        
                        '�ȴ��������ֶ�(��ѯʱֻȡ��һ��ֵ)
                        If strData <> "" Then
                            For i = 0 To UBound(Split(strData, "|"))
                                strTmp = GetFieldValue(Me, CStr(Split(strData, "|")(i)))
                                If .��ʽ <> "" Then
                                    On Error Resume Next
                                    strTmp = Format(strTmp, .��ʽ)
                                    If Err.Number <> 0 Then Err.Clear
                                    On Error GoTo errH
                                End If
                                strValue = Replace(strValue, "[" & Split(strData, "|")(i) & "]", strTmp)
                            Next
                        End If
                        
                        '�ٴ��������:[=������]��[n>=0]��[���ڸ�ʽ��]��[��λ����]
                        strValue = GetLabelMacro(Me, strValue)
                        '[ҳ��]��[ҳ��]Ԥ��ʱ����ֵ
                        strValue = Replace(strValue, "[ҳ��]", "")
                        strValue = Replace(strValue, "[ҳ��]", "")
                        
                        Set objPic = Nothing
                        If strValue <> "" Then
                            Unload frmFlash 'ǿ�Ƴ�ʼPicture����Ȼ�л�����������
                            If .��� = 1 Then
                                Set objPic = DrawBarCode128(frmFlash.picTemp, 3, strValue, Mid(.��ͷ, 1, 1) = "1")
                            ElseIf .��� = 2 Then
                                Set objPic = DrawBarCode39(frmFlash.picTemp, 3, strValue, Mid(.��ͷ, 2, 1) = "1", Mid(.��ͷ, 1, 1) = "1")
                            ElseIf .��� = 3 Then
                                Set objPic = DrawBarCode128Auto(frmFlash.picTemp, strValue, sngWidth, .�и�, Mid(.��ͷ, 1, 1) = "1")
                            ElseIf .��� = 10 Then
                                Set objPic = DrawBarCode2D(strValue, frmFlash.picTemp, lngSize)
                            End If
                            If Val(Mid(.��ͷ, 3, 1)) <> 0 Then
                                Set objPic = PictureSpin(objPic, Val(Mid(.��ͷ, 3, 1)), frmFlash.picTemp)
                            End If
                        End If
                        Set objLoad.Picture = objPic
                        
                        If .��� = 3 Then
                            '128���Զ��������
                            If Val(Mid(.��ͷ, 3, 1)) = 0 Then
                                .W = Me.ScaleX(sngWidth, vbMillimeters, vbTwips)
                                objLoad.Width = .W
                            Else
                                .H = Me.ScaleY(sngWidth, vbMillimeters, vbTwips)
                                objLoad.Height = .H
                            End If
                        ElseIf .��� = 10 And .�Ե� Then
                            '��ά����ȱʡ�Զ�������С
                            objLoad.Width = lngSize
                            objLoad.Height = lngSize
                            .W = lngSize: .H = lngSize
                        End If
                        
                        objLoad.ZOrder
                        objLoad.Visible = True
                End Select
            End With
        End If
    Next
    
    '�����ǩԪ��Ϊָ�����Ԫ��ֵ������
    For Each objItem In mobjReport.Items
        If objItem.���� = 4 Or objItem.���� = 5 Then
            Call mdlPublic.SetCellValue(Val("0-��Ԥ��"), Me, objItem)
        End If
    Next
    
    scrVsc.Visible = Not (intGridCount = 1 And Not mobjReport.Ʊ��)
    scrHsc.Visible = Not (intGridCount = 1 And Not mobjReport.Ʊ��)
    picShadow.Visible = Not (intGridCount = 1 And Not mobjReport.Ʊ��)
    
    '���ø�����ȱʡ����
    Call SetGridAlign
        
    mobjReport.intGridCount = intGridCount
    mobjReport.intGridID = intGridID
        
    Call Form_Resize
    
    ShowFlash
    blnRefresh = True
    LockWindowUpdate 0
    Exit Sub
errH:
    ShowFlash
    LockWindowUpdate 0
    If ErrCenter() = 1 Then
        If mbytStyle = 0 Or mbytStyle = 1 Then ShowFlash "������֯��������,���Ժ򣮣���", , Me
        LockWindowUpdate Me.hwnd
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetFixHeight(objGrid As Object) As Long
'���ܣ���ȡָ�����Ĺ̶����ݸ߶�
    Dim i As Integer, lngH As Long
    
    For i = 0 To objGrid.FixedRows - 1
        lngH = lngH + objGrid.RowHeight(i)
    Next
    GetFixHeight = lngH
End Function

Private Sub GetGridCurSize(ByVal intID As Integer, ByRef X As Long, ByRef Y As Long, _
    ByRef W As Long, ByRef H As Long, Optional ByRef Bottom As Long)
'���ܣ���ȡָ�����ĵ�ǰ��ʾ����ߴ�(ָ���������ӻ򸽼ӱ������)
'���أ�X,Y,W,H��Bottom(���ʱ)
    Dim objItem As RPTItem, tmpItem As RPTItem, lngCurH As Long, lngBottom As Long
    
    X = msh(intID).Left
    W = msh(intID).Width
    
    If Val(msh(intID).Tag) = 0 Then
        Y = msh(intID).Top
        lngCurH = msh(intID).Height
    Else
        Y = msh(CInt(msh(intID).Tag)).Top
        lngCurH = msh(CInt(msh(intID).Tag)).Height
    End If
        
    lngBottom = mobjReport.Items("_" & intID).Y + mobjReport.Items("_" & intID).H
    
    Set objItem = mobjReport.Items("_" & intID)
    W = W * objItem.����
    
    '���ܱ�Ҳ�����и��ӱ�
    For Each tmpItem In mobjReport.Items '���ϸ��ӱ�߶�
        If tmpItem.��ʽ�� = bytFormat And tmpItem.���� = 4 _
            And tmpItem.���� = 1 And tmpItem.���� = objItem.���� Then
            lngCurH = lngCurH + msh(CInt(msh(tmpItem.id).Tag)).Height
            lngBottom = lngBottom + tmpItem.H
        End If
    Next
    H = lngCurH
    Bottom = lngBottom
End Sub

Private Function GetDependID(strName As String) As Integer
'���ܣ����ݲ�������,��ȡ������.
    Dim objItem As RPTItem
    
    For Each objItem In mobjReport.Items
        If objItem.��ʽ�� = bytFormat And objItem.���� = strName _
            And (objItem.���� = 4 Or objItem.���� = 5) And objItem.���� = 0 Then
            GetDependID = objItem.id: Exit Function
        End If
    Next
End Function

Private Sub SetPlace()
'���ܣ����ݱ������ݣ����ñ�񡢱�ǩ�����λ��
    Dim objItem As RPTItem, tmpItem As RPTItem
    Dim lngDesignH As Long, lngShowH As Long, lngAppH As Long
    Dim lngCurH As Long, lngCurTop As Long, bytKind As Byte
    Dim strAppGrid As String, lngFixH As Long
    Dim strGridScale As String, i As Integer
    Dim intCurID As Integer, sngScale As Single
    Dim lngTX As Long, lngTY As Long, lngTW As Long, lngTH As Long
    Dim lngBottom As Long
    
    On Error GoTo errH
    
    If mobjReport Is Nothing Then Exit Sub
    If Not mobjReport.blnLoad Then Exit Sub
    
    '��������ʺϴ����С
    If intGridCount = 1 And Not mobjReport.Ʊ�� Then
        '��ֻ��һ�������ʱ(���������ӡ����ӱ��)��
        '1:��Top��Leftʹ�þ���λ��
        '2.��Widthʹ����Ա�Left�����λ��
        '3.���±�ǩ�ܸ߶Ȳ��ܳ�����ߵĵ�һ��,���ȡ�������±�ǩλ�õ���Ը߶�
        Set objItem = mobjReport.Items("_" & intGridID)
        
        '��������ƺ͵�ǰ�߶�(ע����������û�б�ͷ�����)
        lngDesignH = objItem.H '��Ƹ߶�
        If Val(msh(intGridID).Tag) > 0 Then
            '��ͷʵ�ʸ߶����������߶�
            lngShowH = msh(CInt(msh(intGridID).Tag)).Height
        Else
            lngShowH = msh(intGridID).Height
        End If
        
        For Each tmpItem In mobjReport.Items '���ϸ��ӱ�߶�
            If tmpItem.��ʽ�� = bytFormat And tmpItem.���� = 4 _
                And tmpItem.���� = 1 And tmpItem.���� = objItem.���� Then
                lngDesignH = lngDesignH + tmpItem.H
                lngShowH = lngShowH + msh(CInt(msh(tmpItem.id).Tag)).Height
                strAppGrid = strAppGrid & "," & tmpItem.id '���ӱ����
            End If
        Next
        strAppGrid = Mid(strAppGrid, 2)
        
        '������±�ǩռ���ܸ߶�
        For Each tmpItem In mobjReport.Items
            If tmpItem.��ʽ�� = bytFormat And tmpItem.���� = 2 _
                And tmpItem.ͼƬ Is Nothing And tmpItem.Y >= objItem.Y + lngDesignH Then
                If tmpItem.Y + lbl(tmpItem.id).Height > lngAppH Then
                    lngAppH = tmpItem.Y + lbl(tmpItem.id).Height '����ʵ�߶ȱȽ�
                End If
            End If
        Next
        
        If lngAppH > 0 Then lngAppH = lngAppH - (objItem.Y + lngDesignH)
        
        '�±߸߶Ȳ��ܳ�����ߵ�һ��
        If lngAppH > lngShowH / 2 Then lngAppH = lngShowH / 2
        
        msh(intGridID).Width = (picPaper(intReport).ScaleWidth - msh(intGridID).Left * 2) / objItem.����

        '��������Ӧ��ռ�ĸ߶�
        If Val(msh(intGridID).Tag) = 0 Then
            lngCurH = picPaper(intReport).ScaleHeight - msh(intGridID).Top - (lngAppH + 200)
        Else
            msh(CInt(msh(intGridID).Tag)).Width = msh(intGridID).Width
            lngCurH = picPaper(intReport).ScaleHeight - msh(CInt(msh(intGridID).Tag)).Top - (lngAppH + 200)
        End If
        
        If strAppGrid = "" Then 'û�и��ӱ��ʱ
            If objItem.���� = 5 Then
                msh(intGridID).Height = lngCurH
            Else
                bytKind = GetGridStyle(mobjReport, intGridID)
                If bytKind = 2 Then
                    msh(intGridID).Height = lngCurH
                Else
                    lngFixH = GetFixHeight(msh(CInt(msh(intGridID).Tag))) '����ͷ�߶�
                    If lngCurH < lngFixH + 300 Then lngCurH = lngFixH + 300 '����Ҫ��֤������ʾ��ͷ(�ӹ�����)
                    msh(CInt(msh(intGridID).Tag)).Height = lngCurH
                    msh(intGridID).Height = lngCurH - lngFixH
                End If
            End If
        Else
            '�и��ӱ��ʱ
            '����������ӱ��ռ�ܸ߶ȵı���
            strGridScale = "|" & objItem.id & "," & objItem.H / lngDesignH
            For i = 0 To UBound(Split(strAppGrid, ","))
                Set tmpItem = mobjReport.Items("_" & Split(strAppGrid, ",")(i))
                strGridScale = strGridScale & "|" & tmpItem.id & "," & tmpItem.H / lngDesignH
            Next
            strGridScale = Mid(strGridScale, 2) '"��ID,����|��ID,����..."
            lngCurTop = objItem.Y
            For i = 0 To UBound(Split(strGridScale, "|"))
                intCurID = CInt(Split(Split(strGridScale, "|")(i), ",")(0))
                sngScale = CSng(Split(Split(strGridScale, "|")(i), ",")(1))
                
                If i > 0 Then
                    msh(intCurID).Width = msh(intGridID).Width
                    msh(CInt(msh(intCurID).Tag)).Width = msh(intCurID).Width
                End If
                
                bytKind = GetGridStyle(mobjReport, intCurID)
                
                If Val(msh(intCurID).Tag) = 0 Then 'Ϊ�����,Ҳ���ܴ�����,��������Ϊ����
                    msh(intCurID).Height = lngCurH * sngScale
                    lngCurTop = lngCurTop + msh(intCurID).Height
                Else
                    lngFixH = GetFixHeight(msh(CInt(msh(intCurID).Tag)))
                    If lngCurH * sngScale < lngFixH + 300 Then '����Ҫ��֤������ʾ��ͷ(�ӹ�����)
                        msh(CInt(msh(intCurID).Tag)).Height = lngFixH + 300
                    Else
                        msh(CInt(msh(intCurID).Tag)).Height = lngCurH * sngScale
                    End If
                    msh(CInt(msh(intCurID).Tag)).Top = lngCurTop
                    lngCurTop = lngCurTop + msh(CInt(msh(intCurID).Tag)).Height
                    
                    bytKind = GetGridStyle(mobjReport, intCurID)
                    If bytKind = 2 Then
                        msh(intCurID).Top = msh(CInt(msh(intCurID).Tag)).Top
                        msh(intCurID).Height = msh(CInt(msh(intCurID).Tag)).Height
                    Else
                        msh(intCurID).Top = msh(CInt(msh(intCurID).Tag)).Top + lngFixH
                        msh(intCurID).Height = msh(CInt(msh(intCurID).Tag)).Height - lngFixH
                    End If
                End If
            Next
        End If
    End If
    
    For Each tmpItem In mobjReport.Items
        If tmpItem.��ʽ�� = bytFormat And tmpItem.���� = Val("4-�����") And tmpItem.���� > 1 _
            And tmpItem.���� = 0 And tmpItem.���� = "" Then
            '������־
            For i = 2 To tmpItem.����
                With msh(tmpItem.id)
                    DrawCell picPaper(intReport), "���ݷ���λ��", tmpItem.X + ((i - 1) * .Width), tmpItem.Y, .Width, _
                        msh(CInt(.Tag)).Height - 15, , , .GridColor, .ForeColor, .BackColor, .Font, , 1, 1
                End With
            Next
        End If
    Next
    
    '������ǩ�ʺ����ղα��λ��(����ʱ)
    For Each tmpItem In mobjReport.Items
        If tmpItem.��ʽ�� = bytFormat And tmpItem.���� = 2 And tmpItem.ͼƬ Is Nothing Then
            '��������ҿ��룺���ܶ��ٱ�񶼴���
            If tmpItem.���� <> 0 And tmpItem.���� <> "" Then
                GetGridCurSize GetDependID(tmpItem.����), lngTX, lngTY, lngTW, lngTH
                Select Case tmpItem.����
                    Case 11, 21 '����
                        lbl(tmpItem.id).Left = lngTX
                    Case 12, 22 '����
                        lbl(tmpItem.id).Left = lngTX + (lngTW - lbl(tmpItem.id).Width) / 2
                    Case 13, 23 '����
                        lbl(tmpItem.id).Left = lngTX + lngTW - lbl(tmpItem.id).Width
                End Select
            End If
            '���µ��Զ����룺ֻ��һ��������ʱ�Ŵ���(���б�ǩ,��������)
            If intGridCount = 1 Then
                GetGridCurSize intGridID, lngTX, lngTY, lngTW, lngTH, lngBottom
                If tmpItem.Y >= lngBottom Then
                    lbl(tmpItem.id).Top = lngTY + lngTH + (tmpItem.Y - lngBottom)
                End If
            End If
        End If
    Next
    Exit Sub
errH:
    Err.Clear
    On Error GoTo 0
End Sub

Private Function GridHaveApp(intID As Integer) As Boolean
'���ܣ��ж�һ������Ƿ���и��ӱ��
    Dim tmpItem As RPTItem, strName As String
    
    strName = mobjReport.Items("_" & intID).����
    For Each tmpItem In mobjReport.Items
        If tmpItem.��ʽ�� = bytFormat And tmpItem.���� = 4 And tmpItem.���� = 1 And tmpItem.���� = strName Then
            GridHaveApp = True: Exit Function
        End If
    Next
End Function

Private Function GetGridDesignWidth(objItem As RPTItem) As Long
'���ܣ���ȡ�������������ӱ�������ʱ���ܿ��
    Dim lngW As Long, tmpItem As RPTItem
    
    lngW = objItem.W
    For Each tmpItem In mobjReport.Items
        If tmpItem.��ʽ�� = bytFormat And tmpItem.���� = 5 _
            And tmpItem.���� = 2 And tmpItem.���� = objItem.���� Then
            lngW = lngW + tmpItem.W
        End If
    Next
    GetGridDesignWidth = lngW
End Function

Private Function GetPreAppGrid(intID As Integer, arrGrids As Variant) As Long
'���ܣ���ȡ��ǰ���ӱ���ǰһ�����(����Ϊ���ӱ���������)
'������arrGrids=��XY�Ⱥ�˳���ŵı����������
'˵����1.�����е����ݱ��뱣֤���ӱ��Y�Ⱥ�˳����
'      2.����ǰ������Ϊ���ӱ��ʱ,����ձ��һ���Ѿ����
    Dim objItem As RPTItem, tmpItem As RPTItem, i As Integer
    
    Set objItem = mobjReport.Items("_" & intID)
    For i = 0 To UBound(arrGrids)
        Set tmpItem = mobjReport.Items("_" & arrGrids(i))
        If tmpItem.��ʽ�� = bytFormat And _
            ((tmpItem.���� = 4 And tmpItem.���� = 1 And tmpItem.���� = objItem.����) Or _
            (tmpItem.���� = 0 And objItem.���� = tmpItem.����)) Then
            If tmpItem.id <> intID Then
                GetPreAppGrid = tmpItem.id
            Else '��������ǰ���ӱ��ʱ,��һ���Ϊ����
                Exit Function
            End If
        End If
    Next
End Function

Private Function GetGridDesignHeight(intID As Integer) As Long
'���ܣ���ȡ�������ʱ�߶�(�������и��ӱ��)
'�����������������е��κ�һ���������
    Dim objItem As RPTItem, tmpItem As RPTItem
    Dim lngH As Long
    
    Set objItem = mobjReport.Items("_" & intID)
    If objItem.���� = 1 And objItem.���� <> "" Then
        Set objItem = mobjReport.Items("_" & GetDependID(objItem.����))
    End If
    
    lngH = objItem.H
    For Each tmpItem In mobjReport.Items
        If tmpItem.��ʽ�� = bytFormat And tmpItem.���� = 4 _
            And tmpItem.���� = 1 And tmpItem.���� = objItem.���� Then
            lngH = lngH + tmpItem.H
        End If
    Next
    GetGridDesignHeight = lngH
End Function

Private Function GetGridPageCol(objItem As RPTItem) As Integer
'���ܣ����������е������Զ���ҳ���к�
'������objItem=������
'���أ�-1=û��
    Dim tmpItem As RPTItem, tmpID As RelatID
    
    GetGridPageCol = -1
    If objItem.���� <> 4 Then Exit Function
    
    For Each tmpID In objItem.SubIDs
        Set tmpItem = mobjReport.Items("_" & tmpID.id)
        If tmpItem.�߿� Then
            GetGridPageCol = tmpItem.���
            Exit For
        End If
    Next
End Function

Private Sub AddPrintPage(ByVal intPage As Integer, ByVal objBody As Object, ByVal colCard As Collection _
    , ByVal lngPageBeginRow As Long, ByVal lngPageEndRow As Long _
    , ByVal lngW As Long, ByVal lngL As Long)

    '��̬�����������
    If intPage > 0 Then
        ReDim Preserve marrPageCard(intPage) As PageCards
    Else
        ReDim marrPageCard(intPage) As PageCards
    End If
    Set marrPageCard(intPage) = New PageCards
    
    '�����µĴ�ӡҳ����
    marrPageCard(intPage).Add objBody.Index, objBody.Left, objBody.Top, objBody.Width _
        , objBody.Height, lngPageBeginRow, lngPageEndRow, lngW, lngL, colCard, "_" & objBody.Index
End Sub

Private Function CalcCellPage() As Boolean
'���ܣ����㵥Ԫ����ҳ�Ķ�Ӧ��ϵ
'������mobjreport=�������
'      marrPage=��ӡҳ��
'���أ��Ƿ���Խ��д�ӡ��Ԥ��(��̶����гߴ�Ƚϱ��ߴ绹��)
'˵����������иú���֮��isArray(marrPage)=False,�����û�б�����
    Dim objBody As Control, objPageCell As PageCell, arrPage As Variant '��ǰ���������
    Dim lngFixW As Long, lngFixH As Long '��ǰ���̶����гߴ�
    Dim lngRowB As Long, lngRowE As Long
    Dim lngColB As Long, lngColE As Long '��ֹ����
    Dim lngBodyW As Long, lngBodyH As Long '�����̶����к���ÿ��
    Dim lngCurW As Long, lngCurH As Long '��ǰҳ�б������ۼƵ��Ŀ��
    Dim lngOutX As Long, lngOutY As Long '��ǰ���б�������ʵ��λ��(��Ҫ���ڸ��ӱ��)
    Dim bytKind As Byte, intPage As Integer  '��ǰ������ҳ(0-N)
    Dim i As Long, j As Long, k As Long, strTmp As String
    Dim objItem As RPTItem, blnHaveApp As Boolean, blnHorPage As Boolean
    Dim blnApp As Boolean, lngMinH As Long, arrGrids As Variant
    Dim lngPreID As Long, intDepend As Integer, lngDesignH As Long
    Dim tmpPageCell As PageCell
    Dim lngL As Long, lngW As Long, lngC As Long, lngZ As Long
    Dim lngTop As Long, lngLeft As Long
    Dim lngCount As Long, lngRowsHeight As Long
    Dim blnData As Boolean, tmpSubID As RelatID
    Dim Y As Long, X As Long, Z As Long
    
    '�����������Զ���ҳ��ر���
    Dim strCurText As String, blnNewPage As Boolean, lngPageCol As Long
    Dim lngBaseRows As Long, lngVRowE As Long
    Dim colCardRow As New Collection  '��¼��Ƭ�ڱ����С��ʾ����
    Dim lngLastID As Long, lngRow As Long
    Dim lngRowCount As Long, colCard As New Collection
    Dim blnRePage As Boolean, blnPage As Boolean
    Dim arrPageTmp As Variant, arrTmp As Variant
    Dim intGridID As Integer
    Dim vsfTmp As VSFlexGrid
    
    '�����X,Y�Ⱥ��������
    arrGrids = Array()
    For Each objBody In msh
        If objBody.Index <> 0 And (objBody.Container Is picPaper(intReport) Or objBody.Container.name = "pic") _
            And Left(objBody.Tag, 2) <> "H_" Then
            ReDim Preserve arrGrids(UBound(arrGrids) + 1)
            arrGrids(UBound(arrGrids)) = objBody.Left & "," & objBody.Top & "," & objBody.Index
        End If
    Next
    For i = 0 To UBound(arrGrids) - 1
        For j = i To UBound(arrGrids)
            If CLng(Split(arrGrids(j), ",")(0)) < CLng(Split(arrGrids(i), ",")(0)) Then
                strTmp = arrGrids(i): arrGrids(i) = arrGrids(j): arrGrids(j) = strTmp
            End If
        Next
    Next
    For i = 0 To UBound(arrGrids) - 1
        For j = i To UBound(arrGrids)
            If CLng(Split(arrGrids(j), ",")(1)) < CLng(Split(arrGrids(i), ",")(1)) Then
                strTmp = arrGrids(i): arrGrids(i) = arrGrids(j): arrGrids(j) = strTmp
            End If
        Next
    Next
    For i = 0 To UBound(arrGrids)
        arrGrids(i) = CInt(Split(arrGrids(i), ",")(2))
    Next
    
    arrPage = Empty
    marrPage = Empty
    marrPageCard = Empty
    
    For k = 0 To UBound(arrGrids)
        '���������
        Set objBody = msh(arrGrids(k))
        'ǿ�ƴ�����С�и�
        For i = 0 To objBody.Rows - 1
            If objBody.RowHeight(i) < objBody.RowHeightMin Then
                objBody.RowHeight(i) = objBody.RowHeightMin
            End If
        Next
        
        strTmp = ""
        lngLeft = 0: lngTop = 0
        If objBody.Container.name = "pic" Then
            If objBody.Container.Container Is picPaper(intReport) Then
                lngLeft = mobjReport.Items("_" & objBody.Container.Index).X
                lngTop = mobjReport.Items("_" & objBody.Container.Index).Y
            End If
        End If
        Set objItem = mobjReport.Items("_" & objBody.Index)
        blnApp = (objItem.���� = Val("4-�����") And objItem.���� = Val("1-���ӱ��") And objItem.���� <> "") '�Ƿ񸽼ӱ��
        
        '��ñ��̶��п��̶��и�
        lngFixW = 0: lngFixH = 0
        lngDesignH = GetGridDesignHeight(objItem.id) '�������ӱ��ĸ߶�
        
        If objItem.���� = Val("5-���ܱ�") Then
            For i = 0 To objBody.FixedCols - 1
                lngFixW = lngFixW + objBody.ColWidth(i)
            Next
            For i = 0 To objBody.FixedRows - 1
                lngFixH = lngFixH + objBody.RowHeight(i)
            Next
            '��ȥ�̶�����֮��һҳ���õĿ�Ⱥ͸߶�(�������)
            lngBodyW = GetGridDesignWidth(objItem) - lngFixW
            lngBodyH = lngDesignH - lngFixH
        Else
            bytKind = GetGridStyle(mobjReport, objBody.Index)
            For i = 0 To msh(CInt(objBody.Tag)).FixedRows - 1
                lngFixH = lngFixH + msh(objBody.Tag).RowHeight(i)
            Next
            Select Case bytKind
            Case Val("0-��ͷ����")
                lngBodyH = lngDesignH - lngFixH
            Case Val("1-��ͷ")
                lngBodyH = 0
            Case Val("2-����")
                lngBodyH = lngDesignH
                lngFixH = 0
            End Select
            lngBodyW = objItem.W
        End If
        
        '����Ϊ�������ص��ֵ
        lngBodyW = Round(lngBodyW / Screen.TwipsPerPixelX) * Screen.TwipsPerPixelX
        
        If objItem.���� = Val("4-�����") Then blnHaveApp = GridHaveApp(objItem.id)
        
        lngPageCol = GetGridPageCol(objItem) '��������,û��Ϊ-1
        lngRowB = objBody.FixedRows
        lngColB = objBody.FixedCols
        lngRowE = lngRowB - 1
        lngColE = lngColB - 1
        
        '��ǰ��������ʼҳ��
        If blnApp Then
            '���յı��ID
            intDepend = GetDependID(objItem.����)
            '��һ������ĸ��ӱ��ID
            '��Ϊ�ñ��Ϊ���ӱ��,����һ�����һ���Ѿ����
            lngPreID = GetPreAppGrid(objItem.id, arrGrids)
            intPage = -1
            If objItem.��ID <> 0 Then
                arrTmp = arrPageTmp
            Else
                arrTmp = arrPage
            End If
            For i = 0 To UBound(arrTmp)
                For Each objPageCell In arrTmp(i)
                    If objPageCell.id = lngPreID Then
                        '���һ��(����һ��)���ҳΪ�ϸ�����������ҳ
                        '(��Ϊ���ܱ���ܺ����ҳ,����Щ��ҳ��������ӱ��)
                        If objPageCell.RowE >= msh(objPageCell.id).Rows - 1 _
                            And objPageCell.ColB = msh(objPageCell.id).FixedCols Then
                            '�ж�ʣ��߶��Ƿ����(��С�߶�Ϊ��ͷ��һ��)
                            Select Case bytKind
                                Case 0 '�������
                                    lngMinH = lngFixH + objItem.�и�
                                Case 1 '�������ͷ
                                    lngMinH = lngFixH
                                Case 2 '���������
                                    lngMinH = objItem.�и�
                            End Select
                            If lngDesignH - ((objPageCell.Y + objPageCell.H) - mobjReport.Items("_" & intDepend).Y) >= lngMinH Then
                                lngOutX = objPageCell.X + lngLeft
                                lngOutY = objPageCell.Y + objPageCell.H + lngTop
                                Select Case bytKind
                                    Case 0 '�������
                                        lngBodyH = lngDesignH - ((objPageCell.Y + objPageCell.H) - mobjReport.Items("_" & intDepend).Y) - lngFixH
                                    Case 1 '�������ͷ
                                        lngBodyH = 0
                                    Case 2 '���������
                                        lngBodyH = lngDesignH - ((objPageCell.Y + objPageCell.H) - mobjReport.Items("_" & intDepend).Y)
                                End Select
                                intPage = i
                            Else
                                '����ҳ�п�ʼ���,��������������
                                lngOutX = mobjReport.Items("_" & intDepend).X + lngLeft
                                lngOutY = mobjReport.Items("_" & intDepend).Y + lngTop
                                Select Case bytKind
                                    Case 0 '�������
                                        lngBodyH = lngDesignH - lngFixH
                                    Case 1 '�������ͷ
                                        lngBodyH = 0
                                    Case 2 '���������
                                        lngBodyH = lngDesignH
                                End Select
                                intPage = i + 1
                                
                                '���ӱ����������ձ��ĺ����ҳ
                                For j = intPage To UBound(arrTmp)
                                    For Each tmpPageCell In arrTmp(j)
                                        If tmpPageCell.id = intDepend Then
                                            If tmpPageCell.ColB <> msh(intDepend).FixedCols Then
                                                intPage = intPage + 1
                                            End If
                                        End If
                                    Next
                                Next
                            End If
                            Exit For
                        End If
                    End If
                Next
                If intPage <> -1 Then Exit For
            Next
            If intPage = -1 Then intPage = 0
        Else
            lngOutX = objItem.X + lngLeft
            lngOutY = objItem.Y + lngTop
            intPage = 0
        End If
        
        'ҳ��ѭ��(ÿ������ڶ�ҳ�м���)
        Do
            'ҳ��ѭ��(����DO)
            
            '���㵱ǰҳ�з�Χ
            lngCurH = 0
            blnNewPage = False
            Do
                If lngPageCol <> -1 Then
                    If lngRowE + 1 = lngRowB Then
                        'ÿҳ��һ��ΪlngRowE=lngRowB-1,���ñȽ�,�Ҹ����ڸ�ҳ�ض�Ҫ��ӡ
                        strCurText = objBody.TextMatrix(lngRowE + 1, lngPageCol)
                    ElseIf lngRowE + 1 > lngRowB Then
                        If strCurText <> objBody.TextMatrix(lngRowE + 1, lngPageCol) Then
                            blnNewPage = True
                        End If
                    End If
                End If
                If Not blnNewPage Then
                    lngCurH = lngCurH + objBody.RowHeight(lngRowE + 1)
                    If lngCurH <= lngBodyH Then
                        lngRowE = lngRowE + 1   '����ʵ���и߼����ÿҳ��������������Լ����߶�
                        If lngPageCol <> -1 Then
                            strCurText = objBody.TextMatrix(lngRowE, lngPageCol)
                        End If
                    End If
                End If
            Loop Until (lngCurH > lngBodyH) Or (lngRowE = objBody.Rows - 1) Or blnNewPage
            
            'ȡʵ�ʸ߶�
            If lngCurH > lngBodyH Then lngCurH = lngCurH - objBody.RowHeight(lngRowE + 1)
            
            'ȡ��ǰҳ�����ɵ�ʵ������
            lngRowsHeight = 0
            lngBaseRows = 0
            For i = lngRowB To objBody.Rows - 1
                lngRowsHeight = lngRowsHeight + objBody.RowHeight(i)
                If lngBodyH < lngRowsHeight Then
                    Exit For
                Else
                    lngBaseRows = lngBaseRows + 1
                End If
            Next
            If lngBodyH > lngRowsHeight Then
                '��Ʊ�����ʵ�������ߵĲ�����������
                lngBaseRows = lngBaseRows + (lngBodyH - lngRowsHeight) \ objItem.�и�
            End If
            
            '�����ӡһ��,��ǿ�д�ӡһ��
            If lngRowE < lngRowB Then lngRowE = lngRowB
            
            '���������Ʊ��ʱ�������������������и���ͬ
            lngVRowE = 0 '��������Ʊ�����ʱ�����������
            If objItem.���� > 1 Then
                '�����ʵ��ҳβ��(ǰ��ֻ�ǿ��ܳ����߶���)
                If lngPageCol <> -1 Then
                    strCurText = objBody.TextMatrix(lngRowE, lngPageCol)
                    For i = lngRowE + 1 To objBody.Rows - 1
                        If i - lngRowB + 1 > lngBaseRows * objItem.���� Then
                            lngRowE = i - 1: Exit For
                        ElseIf strCurText <> objBody.TextMatrix(i, lngPageCol) Then
                            lngRowE = i - 1: Exit For
                        Else
                            lngRowE = i
                        End If
                        strCurText = objBody.TextMatrix(i, lngPageCol)
                    Next
                Else
                    For i = lngRowE + 1 To objBody.Rows - 1
                        If i - lngRowB + 1 > lngBaseRows * objItem.���� Then
                            lngRowE = i - 1: Exit For
                        Else
                            lngRowE = i
                        End If
                    Next
                End If
                '������⻻ҳβ��(���������Ŀհ���)
                If mobjReport.Ʊ�� Then
                    '��������Ʊ��ʱ������Ʋ�����
                    lngVRowE = lngRowE + (lngBaseRows * objItem.���� - (lngRowE - lngRowB + 1))
                Else
                    '��������Ʊ��ʱ��������ʵ�������������������
                    If lngRowE - lngRowB + 1 <= lngBaseRows Then
                        lngVRowE = lngRowE + (lngRowE - lngRowB + 1) * (objItem.���� - 1)
                    Else
                        lngVRowE = lngRowE + (lngBaseRows * objItem.���� - (lngRowE - lngRowB + 1))
                    End If
                End If
            Else
                'û�з���ʱ��Ʊ����Ҫ������
                '�����ǲ������ݱ仯ǿ�л�ҳ
                If mobjReport.Ʊ�� Then
                    lngVRowE = lngRowE + (lngBaseRows - (lngRowE - lngRowB + 1))
                End If
            End If
            If lngVRowE = lngRowE Then lngVRowE = 0
            
            '�����з�Χ(�����ҳ�Ƕ�ҳ)
            Do
                '���㵱ǰҳ�з�Χ
                lngCurW = 0
                Do
                    lngCurW = lngCurW + objBody.ColWidth(lngColE + 1)
                    If lngCurW <= lngBodyW Then lngColE = lngColE + 1
                Loop Until lngCurW > lngBodyW Or lngColE = objBody.Cols - 1
                
                'ȡ��ʵ���
                If lngCurW > lngBodyW Then lngCurW = lngCurW - objBody.ColWidth(lngColE + 1)
                
                '�����ӡһ��,��ǿ�д�ӡһ��
                If lngColE < lngColB Then lngColE = lngColB
                
                If objItem.��ID = 0 Then
                    '��Ƭ��ı��Ҫ��ҳ
                    blnPage = True
                Else
                    '��Ƭ�ڲ��ı��Ƭ������Դʱ����ҳ
                    If mobjReport.Items("_" & objItem.��ID).����Դ = "" Then
                        blnPage = True
                    Else
                        blnPage = False
                    End If
                End If
                
                '�µ�һҳ��ʼ
                If blnPage Then
                    If Not IsArray(arrPage) Then
                        ReDim arrPage(intPage) As PageCells  '��һ�γ�ʼҳ
                        Set arrPage(intPage) = New PageCells
                    ElseIf intPage > UBound(arrPage) Then
                        '�����ҳ�ѱ��������ռ��,�����ٳ�ʼ
                        ReDim Preserve arrPage(intPage) As PageCells
                        Set arrPage(intPage) = New PageCells
                    End If
                Else
                    If intPage = 0 Then
                        If Not IsArray(arrPageTmp) Then
                            ReDim arrPageTmp(intPage) As PageCells  '��һ�γ�ʼҳ
                            Set arrPageTmp(intPage) = New PageCells
                        ElseIf intPage > UBound(arrPageTmp) Then
                            '�����ҳ�ѱ��������ռ��,�����ٳ�ʼ
                            ReDim Preserve arrPageTmp(intPage) As PageCells
                            Set arrPageTmp(intPage) = New PageCells
                        End If
                    End If
                End If
                blnData = False
                If objBody.Container.name = "pic" Then
                    '��Ƭ
                    If objBody.Container.Container Is picPaper(intReport) Then
                        If mobjReport.Items("_" & objBody.Index).SubIDs.count > 0 And mobjReport.Items("_" & objBody.Container.Index).����Դ <> "" And lngLastID <> objBody.Index Then
                            For Each tmpSubID In mobjReport.Items("_" & objBody.Index).SubIDs
                                If mobjReport.Items("_" & tmpSubID.id).���� <> "" Then
                                    With mobjReport.Items("_" & tmpSubID.id)
                                        X = InStr(1, .����, "]")
                                        Y = InStr(1, .����, ".")
                                        Z = InStr(1, .����, "[")
                                        If X > Z And X > Y And X <> 0 And Z <> 0 Then
                                            If Mid(.����, Z + 1, Y - Z - 1) = mobjReport.Items("_" & objBody.Container.Index).����Դ Then
                                                blnData = True
                                                Exit For
                                            End If
                                        End If
                                    End With
                                End If
                            Next
                            If blnData Then
                                On Error Resume Next
                                If lngCurH \ mobjReport.Items("_" & objBody.Index).�и� < colCardRow("_" & objBody.Container.Index) Then
                                    If Err.Number = 0 Then colCardRow.Remove "_" & objBody.Container.Index
                                    colCardRow.Add lngCurH \ mobjReport.Items("_" & objBody.Index).�и�, "_" & objBody.Container.Index
                                End If
                                On Error GoTo 0
                            End If
                        End If
                    End If
                End If
                lngLastID = objBody.Index
                
                '�����µĴ�ӡҳ����
                'ֻ�п�Ƭ��ı��ŷ�ҳ
                If blnPage Then
                    arrPage(intPage).Add objBody.Index, lngOutX, lngOutY, lngCurW + lngFixW, lngCurH + lngFixH, _
                        lngDesignH, lngRowB, lngRowE, lngVRowE, lngColB, lngColE, _
                        lngFixW, lngFixH, objItem.����, "_" & objBody.Index
                Else
                    If intPage = 0 Then
                        arrPageTmp(intPage).Add objBody.Index, lngOutX, lngOutY, lngCurW + lngFixW, lngCurH + lngFixH, _
                            lngDesignH, lngRowB, lngRowE, lngVRowE, lngColB, lngColE, _
                            lngFixW, lngFixH, objItem.����, "_" & objBody.Index
                    End If
                End If
                lngColB = lngColE + 1
                lngColE = lngColB - 1
                
                intPage = intPage + 1
            
                If blnApp Then
                    '���ӱ����������ձ��ĺ����ҳ
                    If objItem.��ID <> 0 Then
                        arrTmp = arrPageTmp
                    Else
                        arrTmp = arrPage
                    End If
                    For i = intPage To UBound(arrTmp)
                        For Each objPageCell In arrTmp(i)
                            If objPageCell.id = intDepend Then
                                If objPageCell.ColB <> msh(intDepend).FixedCols Then
                                    intPage = intPage + 1
                                End If
                            End If
                        Next
                    Next
                    '������ҳ���õ�λ�óߴ�
                    '����ҳ�п�ʼ���,��������������
                    lngOutX = mobjReport.Items("_" & intDepend).X
                    lngOutY = mobjReport.Items("_" & intDepend).Y
                    Select Case bytKind
                        Case 0 '�������
                            lngBodyH = lngDesignH - lngFixH
                        Case 1 '�������ͷ
                            lngBodyH = 0
                        Case 2 '���������
                            lngBodyH = lngDesignH
                    End Select
                End If
            
            '������з���ʱ�����ڸ�������ʱ�������ҳ
            Loop Until lngColB > objBody.Cols - 1 Or _
                (objItem.���� = 4 And (objItem.���� > 1 Or objItem.���� = Val("1-���ӱ�") Or blnHaveApp))
            
            lngColB = objBody.FixedCols
            lngColE = lngColB - 1
                            
            lngRowB = lngRowE + 1
            lngRowE = lngRowB - 1
        '�����д�������,�ɸñ��Ҳ���ˣ�ֻ��ʾ��ͷʱֻ��һҳ
        Loop Until lngRowB > objBody.Rows - 1 Or (objItem.���� = 4 And bytKind = 1)
    Next
    
    '��Ƭ��̬��ӡ
    For Each objBody In pic
        '����Ƕ�̬��ӡ
        If objBody.Index <> 0 And Not mobjReport.Items("_" & objBody.Index) Is Nothing Then
            If mobjReport.Items("_" & objBody.Index).����Դ <> "" Then
                lngRowB = 0
                lngRowE = 0
                intPage = 0
                lngCount = 0
                If mobjReport.Items("_" & objBody.Index).������� = 0 Then
                    If mobjReport.Fmts.Item("_" & bytFormat).ֽ�� = 1 Then
                        lngL = (mobjReport.Fmts.Item("_" & bytFormat).W - objBody.Left + mobjReport.Items("_" & objBody.Index).���Ҽ��) \ (objBody.Width + mobjReport.Items("_" & objBody.Index).���Ҽ��)
                    Else
                        lngL = (mobjReport.Fmts.Item("_" & bytFormat).H - objBody.Left + mobjReport.Items("_" & objBody.Index).���Ҽ��) \ (objBody.Width + mobjReport.Items("_" & objBody.Index).���Ҽ��)
                    End If
                Else
                    lngL = mobjReport.Items("_" & objBody.Index).�������
                End If
                If mobjReport.Items("_" & objBody.Index).������� = 0 Then
                    If mobjReport.Fmts.Item("_" & bytFormat).ֽ�� = 1 Then
                        lngW = (mobjReport.Fmts.Item("_" & bytFormat).H - objBody.Top + mobjReport.Items("_" & objBody.Index).���¼��) \ (objBody.Height + mobjReport.Items("_" & objBody.Index).���¼��)
                    Else
                        lngW = (mobjReport.Fmts.Item("_" & bytFormat).W - objBody.Top + mobjReport.Items("_" & objBody.Index).���¼��) \ (objBody.Height + mobjReport.Items("_" & objBody.Index).���¼��)
                    End If
                Else
                    lngW = mobjReport.Items("_" & objBody.Index).�������
                End If
                
                'һҳ�������ɶ��ٿ�Ƭ
                lngC = lngW * lngL
                With mLibDatas("_" & mobjReport.Items("_" & objBody.Index).����Դ).DataSet
                    If .RecordCount > 0 Then .MoveFirst
                    
                    '�������ʶ��
                    For i = 0 To .Fields.count - 1
                        If .Fields(i).name = "�����ʶ" Then
                            Exit For
                        End If
                    Next
                    
                    If i >= 0 And i <= .Fields.count - 1 Then
                        '�С������ʶ����
                        
                        '��ȡ����ؼ�
                        intGridID = -1
                        For Each vsfTmp In msh
                            If vsfTmp.Index > 0 And Not vsfTmp.Container Is Nothing Then
                                If objBody.Index = vsfTmp.Container.Index Then
                                    intGridID = vsfTmp.Index
                                    Exit For
                                End If
                            End If
                        Next
                        
                        If intGridID >= 0 Then
                            '��ȡ����
                            lngDesignH = GetGridDesignHeight(intGridID)
                            lngTop = msh(CInt(msh(intGridID).Tag)).Top
                            lngFixH = 0
                            For i = 0 To msh(CInt(msh(intGridID).Tag)).FixedRows - 1
                                lngFixH = lngFixH + msh(CInt(msh(intGridID).Tag)).RowHeight(i)
                            Next
'                            Select Case bytKind
'                            Case Val("0-��ͷ����")
'                                lngBodyH = lngDesignH - lngFixH
'                            Case Val("1-��ͷ")
'                                lngBodyH = 0
'                            Case Val("2-����")
'                                lngBodyH = lngDesignH
'                                lngFixH = 0
'                            End Select
                            lngBodyH = lngDesignH - lngFixH
                            
                            lngRowB = 0                 '��Ƭ�Ŀ�ʼ��
                            lngRowE = 0                 '��Ƭ�Ľ�����
                            lngCurH = 0                 'ʵ���и�
                            i = 0                       'ҳ��Ƭ����
                            If .RecordCount > 0 Then
                                strTmp = "" & Nvl(!�����ʶ)
                            Else
                                strTmp = ""
                            End If
                            
                            '����������
                            Do While .EOF = False
                                lngRow = .AbsolutePosition - 1

                                'ҳ�����п�Ƭ
                                '�ֿ�Ƭ�߼���1.�ۼ��и� > ����ߣ� 2.�����ʶ
                                If lngCurH + msh(intGridID).RowHeight(lngRow) > lngBodyH _
                                    Or strTmp <> "" & !�����ʶ Then
                                    '�и߳�����Ƭ��
                                    colCard.Add lngRowB + 1 & "-" & lngRowE + 1
                                    strTmp = "" & !�����ʶ
                                    If lngCurH = 0 Then
                                        'ֻһ�У���������
                                    Else
                                        '����һ��
                                        .MovePrevious
                                    End If
                                    lngCurH = 0
                                    i = i + 1
                                    lngRowB = lngRowE + 1
                                    lngRowE = lngRowB
                                Else
                                    lngCurH = lngCurH + msh(intGridID).RowHeight(lngRow)
                                    lngRowE = lngRow
                                    strTmp = "" & !�����ʶ
                                End If
                                
                                .MoveNext

                                '����ҳ��Ƭ�����Ͳ���һҳ��marrPageCard����
                                If i > lngC - 1 Or .EOF Then
                                    If .EOF Then
                                        lngRowE = .RecordCount - 1
                                        colCard.Add lngRowB + 1 & "-" & lngRowE + 1
                                    End If
                                    Call AddPrintPage(intPage, objBody, colCard, lngRowB, lngRowE, lngW, lngL)
                                    Set colCard = New Collection
                                    intPage = intPage + 1
                                    i = 0
                                End If
                            Loop
                        Else
                            '������ؼ�
                            lngRowB = 0                 '��Ƭ�Ŀ�ʼ��
                            lngRowE = 0                 '��Ƭ�Ľ�����
                            i = 0                       'ҳ��Ƭ����
                            If .RecordCount > 0 Then
                                strTmp = "" & Nvl(!�����ʶ)
                            Else
                                strTmp = ""
                            End If
                            
                            '���㿨Ƭ����
                            Do While .EOF = False
                                lngRow = .AbsolutePosition - 1
                                
                                'ҳ�����п�Ƭ
                                '�ֿ�Ƭ�߼���1.�����ʶ
                                If strTmp <> "" & !�����ʶ Then
                                    colCard.Add lngRowB + 1 & "-" & lngRowE + 1
                                    strTmp = "" & !�����ʶ

                                    '����һ��
                                    .MovePrevious
                                    
                                    i = i + 1
                                    lngRowB = lngRowE + 1
                                    lngRowE = lngRowB
                                Else
                                    lngRowE = lngRow
                                    strTmp = "" & !�����ʶ
                                End If
                                
                                .MoveNext

                                '����ҳ��Ƭ�����Ͳ���һҳ��marrPageCard����
                                If i > lngC - 1 Or .EOF Then
                                    If .EOF Then
                                        lngRowE = .RecordCount - 1
                                        colCard.Add lngRowB + 1 & "-" & lngRowE + 1
                                    End If
                                    Call AddPrintPage(intPage, objBody, colCard, lngRowB, lngRowE, lngW, lngL)
                                    Set colCard = New Collection
                                    intPage = intPage + 1
                                    i = 0
                                End If
                            Loop
                        End If

                    Else
                        '�ޡ������ʶ���У�һ����Ƭֻ��һ����¼
                        If .RecordCount <= lngC Then
                            'ֻһҳ�࿨Ƭ
                            lngRowB = 0
                            lngRowE = .RecordCount - 1
                            For i = lngRowB To lngRowE
                                colCard.Add i + 1 & "-" & i + 1
                            Next
                            Call AddPrintPage(intPage, objBody, colCard, lngRowB, lngRowE, lngW, lngL)
                        Else
                            '��ҳ�࿨Ƭ
                            lngRowB = 0
                            Do While .EOF = False
                                '���п�Ƭ
                                lngRowE = .AbsolutePosition - 1
                                colCard.Add lngRowE + 1 & "-" & lngRowE + 1
                                
                                .MoveNext
                                
                                If colCard.count >= lngC Or .EOF Then
                                    Call AddPrintPage(intPage, objBody, colCard, lngRowB, lngRowE, lngW, lngL)
                                    Set colCard = New Collection
                                    intPage = intPage + 1
                                    lngRowB = lngRowE + 1
                                End If
                            Loop
                        End If
                    End If
                End With
            End If
        End If
    Next
    
    '��̬��ӡ�ı�񵥶�����
    If IsArray(arrPageTmp) Then
        If arrPageTmp(0).count > 0 Then
            For Each objPageCell In arrPageTmp(0)
                If IsArray(marrPageCard) Then
                    For i = 0 To UBound(marrPageCard)
                        On Error Resume Next
                        j = marrPageCard(i).Item("_" & mobjReport.Items("_" & objPageCell.id).��ID).id
                        If Err.Number = 0 Then
                            On Error GoTo 0
                            With marrPageCard(i).Item("_" & mobjReport.Items("_" & objPageCell.id).��ID)
                                    For j = 1 To .Item.count
                                        If Not IsArray(arrPage) Then
                                            ReDim arrPage(i) As PageCells  '��һ�γ�ʼҳ
                                            Set arrPage(i) = New PageCells
                                        ElseIf i > UBound(arrPage) Then
                                            '�����ҳ�ѱ��������ռ��,�����ٳ�ʼ
                                            ReDim Preserve arrPage(i) As PageCells
                                            Set arrPage(i) = New PageCells
                                        End If
                                        If mobjReport.Ʊ�� Then
                                            lngVRowE = (mobjReport.Items("_" & objPageCell.id).H - objPageCell.FixH) \ mobjReport.Items("_" & objPageCell.id).�и� _
                                                    - (Val(Mid(.Item(j), InStr(.Item(j), "-") + 1, Len(.Item(j)))) - Val(Mid(.Item(j), 1, InStr(.Item(j), "-") - 1)) + 1) _
                                                    + Val(Mid(.Item(j), InStr(.Item(j), "-") + 1, Len(.Item(j)))) - 1
                                        Else
                                            lngVRowE = objPageCell.VRowE
                                        End If
                                        arrPage(i).Add objPageCell.id _
                                            , objPageCell.X + ((j - 1) Mod .Col) * (mobjReport.Items("_" & mobjReport.Items("_" & objPageCell.id).��ID).W + mobjReport.Items("_" & mobjReport.Items("_" & objPageCell.id).��ID).���Ҽ��) _
                                            , objPageCell.Y + ((j - 1) \ .Col) * (mobjReport.Items("_" & mobjReport.Items("_" & objPageCell.id).��ID).H + mobjReport.Items("_" & mobjReport.Items("_" & objPageCell.id).��ID).���¼��) _
                                            , objPageCell.W, objPageCell.FixH + (Val(Mid(.Item(j), InStr(.Item(j), "-") + 1, Len(.Item(j)))) - Val(Mid(.Item(j), 1, InStr(.Item(j), "-") - 1)) + 1) * mobjReport.Items("_" & objPageCell.id).�и� _
                                            , objPageCell.MaxH, Val(Mid(.Item(j), 1, InStr(.Item(j), "-") - 1)) - 1, Val(Mid(.Item(j), InStr(.Item(j), "-") + 1, Len(.Item(j)))) - 1 _
                                            , lngVRowE, objPageCell.ColB, objPageCell.ColE, objPageCell.FixW, objPageCell.FixH _
                                            , objPageCell.Copys, "_" & objPageCell.id + (j - 1)
                                    Next
                                
                            End With
                        End If
                        On Error GoTo 0
                    Next
                End If
            Next
        End If
    End If
    
    marrPage = arrPage
    CalcCellPage = True
End Function

Private Sub SetReportIndex(intIndex As Integer, objReport As Report)
'���ܣ����ݵ�ǰҪ��ʾ�ı���,����Ԫ�ص��������Ϻ�׺
'˵������Ҫ�������𱨱����еĶ��ű���Ĳ�ͬԪ��
'ע�⣺��Ԫ�ص���ʵ�ؼ���Ҳһ��������
    Dim tmpItem As RPTItem, objItems As RPTItems
    Dim tmpSubID As RelatID, objSubIDs As RelatIDs
    Dim tmpCopyID As RelatID, objCopyIDs As RelatIDs
    
    Set objItems = New RPTItems
    For Each tmpItem In objReport.Items
        With tmpItem
            Set objSubIDs = New RelatIDs
            For Each tmpSubID In .SubIDs
                objSubIDs.Add tmpSubID.id & intIndex, "_" & tmpSubID.id & intIndex
            Next
            Set objCopyIDs = New RelatIDs
            For Each tmpCopyID In .CopyIDs
                objCopyIDs.Add tmpCopyID.id & intIndex, "_" & tmpCopyID.id & intIndex
            Next
            objItems.Add .id & intIndex, .��ʽ��, .����, .�ϼ�ID, .����, .���, .����, .����, .����, .��ͷ, _
                .X, .Y, .W, .H, .�и�, .����, .�Ե�, .����, .�ֺ�, .����, .����, .б��, .����, .ǰ��, .����, _
                .�߿�, .����, .����, .��ʽ, .����, .����߼Ӵ�, .����Ӧ�и�, .ͼƬ, .ϵͳ, _
                IIF(.��ID = 0, 0, .��ID & intIndex), objSubIDs, objCopyIDs, "_" & .id & intIndex, _
                .����Դ, .���¼��, .���Ҽ��, .Դ�к�, .�������, .�������, .Relations, _
                .ColProtertys, .ˮƽ��ת
        End With
    Next
    
    Set objReport.Items = New RPTItems
    For Each tmpItem In objItems
        With tmpItem
            objReport.Items.Add .id, .��ʽ��, .����, .�ϼ�ID, .����, .���, .����, .����, .����, .��ͷ, _
                .X, .Y, .W, .H, .�и�, .����, .�Ե�, .����, .�ֺ�, .����, .����, .б��, .����, .ǰ��, .����, _
                .�߿�, .����, .����, .��ʽ, .����, .����߼Ӵ�, .����Ӧ�и�, .ͼƬ, .ϵͳ, .��ID, .SubIDs, .CopyIDs, _
                "_" & .id, .����Դ, .���¼��, .���Ҽ��, .Դ�к�, .�������, .�������, .Relations, _
                .ColProtertys, .ˮƽ��ת
        End With
    Next
End Sub

Private Sub mnuViewStyle_Click(Index As Integer)
    Call SetView(CByte(Index))
End Sub

Private Sub SetView(bytStyle As Byte)
'���ܣ�������λ�б���ʾ��ʽ
'������bytstyle=0-��ͼ��,1-Сͼ��,2-�б�,3-��ϸ����
    mnuViewStyle(0).Checked = False
    mnuViewStyle(1).Checked = False
    mnuViewStyle(2).Checked = False
    mnuViewStyle(3).Checked = False
    mnuViewStyle(bytStyle).Checked = True
    lvw.View = bytStyle
End Sub

Private Function GetSubReport(lngGroup As Long) As ADODB.Recordset
'���ܣ����ݱ�����ID,��ȡ���ӱ������Ϣ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH

    strSQL = "Select ��ID,����ID,���,���� From zlRPTSubs Where ��ID=[1] Order by ���"
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngGroup)
    If Not rsTmp.EOF Then Set GetSubReport = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetGroupInfo(lngGroup As Long) As ADODB.Recordset
'���ܣ����ݱ�����ID��ȡ����Ϣ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select ID,���,����,˵��,ϵͳ,����ID,����ʱ�� From zlRPTGroups Where ID=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngGroup)
    If Not rsTmp.EOF Then Set GetGroupInfo = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub InitReportPars()
'���ܣ��ڱ�������,���ݵ�ǰ���������������ʾӦ�������
    Dim i As Integer, j As Integer
    Dim tmpPar As RPTPar, strTmp As String
    Dim lngCurH As Long, objTmp As Object
    Dim intCurTab As Integer, objLoad As Object
    Dim strGroup As String, objGroup As Object
    Dim blnCmd As Boolean, blnExist As Boolean
    Dim strPre As String, strCur As String
    Dim blntmp As Boolean, lngTmp As Long
    
    For Each objLoad In lblName
        If objLoad.Index <> 0 Then Unload objLoad
    Next
    For Each objLoad In txt
        If objLoad.Index <> 0 Then Unload objLoad
    Next
    For Each objLoad In cmd
        If objLoad.Index <> 0 Then Unload objLoad
    Next
    For Each objLoad In cbo
        If objLoad.Index <> 0 Then Unload objLoad
    Next
    For Each objLoad In dtp
        If objLoad.Index <> 0 Then Unload objLoad
    Next
    For Each objLoad In opt
        If objLoad.Index <> 0 Then Unload objLoad
    Next
    For Each objLoad In chk
        If objLoad.Index <> 0 Then Unload objLoad
    Next
    For Each objLoad In fra
        If objLoad.Index <> 0 Then Unload objLoad
    Next
    For Each objLoad In fraGroup
        If objLoad.Index <> 0 Then Unload objLoad
    Next
    
    blnMatch = False
    
    '���������������
    i = 0: lngCurH = lblName(0).Top
    For Each tmpPar In mobjPars
        i = i + 1
        
        Load lblName(i)
        lblName(i).Caption = tmpPar.���� & "(&" & i & ")"
        lblName(i).ToolTipText = tmpPar.����
        lblName(i).Left = txt(0).Left - lblName(i).Width - 30
        lblName(i).Top = lngCurH
        lblName(i).TabIndex = intCurTab: intCurTab = intCurTab + 1
        lblName(i).Visible = True
        
        If tmpPar.ȱʡֵ = "�̶�ֵ�б�" Then
            If tmpPar.��ʽ = 0 Then '������
                Load cbo(i): Set objTmp = cbo(i)
                If tmpPar.�Ƿ����� Then objTmp.Enabled = False
                cbo(i).TabIndex = intCurTab: intCurTab = intCurTab + 1
                cbo(i).Left = cbo(0).Left: cbo(i).Top = lblName(i).Top - (cbo(i).Height - lblName(i).Height) / 2
                '���õķָ���
                For j = 0 To UBound(Split(tmpPar.ֵ�б�, "|"))
                    strTmp = Split(Split(tmpPar.ֵ�б�, "|")(j), ",")(0)
                    
                    If Left(strTmp, 1) = "��" Then
                        cbo(i).AddItem Mid(strTmp, 2)
                        If cbo(i).ListIndex = -1 Then cbo(i).ListIndex = cbo(i).NewIndex
                    Else
                        cbo(i).AddItem strTmp
                    End If
                    '��������ʱReserve�����"��ʾֵ|��ֵ"
                    '�����ϴ���ʾֵ����λȱʡ��
                    If tmpPar.Reserve Like "*|*" Then
                        If Left(strTmp, 1) = "��" Then
                            If Split(tmpPar.Reserve, "|")(0) = Mid(strTmp, 2) Then cbo(i).ListIndex = cbo(i).NewIndex
                        Else
                            If Split(tmpPar.Reserve, "|")(0) = strTmp Then cbo(i).ListIndex = cbo(i).NewIndex
                        End If
                        
                        '�ϴ���Ϊ�����ֵ��ĳ����ֵ��ͬ,��λ
                        '��Ϊ���ѡ��ֵ�а�ֵ�����ظ�,���Դ˶οɲ�Ҫ
                        If Split(tmpPar.Reserve, "|")(0) = Split(Split(tmpPar.ֵ�б�, "|")(j), ",")(1) Then
                            cbo(i).ListIndex = cbo(i).NewIndex
                        End If
                    End If
                Next
                cbo(i).Visible = True
            ElseIf tmpPar.��ʽ = 1 Then '��ѡ��
                Load fra(i): Set objTmp = fra(i)
                fra(i).TabIndex = intCurTab: intCurTab = intCurTab + 1
                fra(i).Left = fra(0).Left: fra(i).Top = lblName(i).Top - 50
                
                lblName(i).Visible = False
                fra(i).Caption = lblName(i).Caption
                                
                j = UBound(Split(tmpPar.ֵ�б�, "|")) + 1 '��ѡ��
                j = CInt((j / 3) + 0.4) '����
                
                fra(i).Height = fra(0).Height + (j - 1) * (opt(0).Height * 1.6) - opt(0).Height * 0.3
                
                blnExist = False '�Ƿ��Ѿ����ϴ�����ֵ�����˵�ǰֵ
                '���õķָ���
                For j = 0 To UBound(Split(tmpPar.ֵ�б�, "|"))
                    strTmp = Split(Split(tmpPar.ֵ�б�, "|")(j), ",")(0)
                    
                    Load opt(opt.UBound + 1)
                    If tmpPar.�Ƿ����� Then opt(opt.UBound).Enabled = False
                    Set opt(opt.UBound).Container = fra(i)
                    opt(opt.UBound).TabIndex = intCurTab: intCurTab = intCurTab + 1
                    opt(opt.UBound).Tag = Split(Split(tmpPar.ֵ�б�, "|")(j), ",")(1) '��Ű�ֵ
                    
                    If InStr(",0,1,3,", "," & UBound(Split(tmpPar.ֵ�б�, "|")) & ",") > 0 Then
                        'ֻ��1,2,4����������⴦��
                        If j = 0 Or j = 1 Then 'Top
                            opt(opt.UBound).Top = opt(0).Top
                        Else
                            opt(opt.UBound).Top = opt(0).Top + opt(0).Height * 1.6
                        End If
                        If j = 0 Or j = 2 Then 'Left
                            opt(opt.UBound).Left = opt(0).Left + 150
                        Else
                            opt(opt.UBound).Left = opt(0).Left + (opt(0).Width * 1.4 + 60) + 150
                        End If
                        
                        If Left(strTmp, 1) = "��" Then
                            opt(opt.UBound).Caption = GetLenStr(Mid(strTmp, 2), opt(0).Width * 1.4 - 200, Me)
                            opt(opt.UBound).ToolTipText = Mid(strTmp, 2)
                            If Not blnExist Then opt(opt.UBound).Value = True
                        Else
                            opt(opt.UBound).Caption = GetLenStr(strTmp, opt(0).Width * 1.4 - 200, Me)
                            opt(opt.UBound).ToolTipText = strTmp
                        End If
                    Else
                        opt(opt.UBound).Top = opt(0).Top + (CInt(((j + 1) / 3) + 0.4) - 1) * (opt(0).Height * 1.6)
                        opt(opt.UBound).Left = opt(0).Left + (IIF(((j + 1) Mod 3) = 0, 3, ((j + 1) Mod 3)) - 1) * (opt(0).Width + 60)
                        
                        If Left(strTmp, 1) = "��" Then
                            opt(opt.UBound).Caption = GetLenStr(Mid(strTmp, 2), opt(0).Width - 200, Me)
                            opt(opt.UBound).ToolTipText = Mid(strTmp, 2)
                            If Not blnExist Then opt(opt.UBound).Value = True
                        Else
                            opt(opt.UBound).Caption = GetLenStr(strTmp, opt(0).Width - 200, Me)
                            opt(opt.UBound).ToolTipText = strTmp
                        End If
                    End If

                    opt(opt.UBound).Width = TextWidth(opt(opt.UBound).Caption) + 300
                    
                    '��������ʱReserve�����"��ʾֵ|��ֵ"
                    '�����ϴ�ѡ��ֵ����λȱʡ��
                    If tmpPar.Reserve Like "*|*" Then
                        If Left(strTmp, 1) = "��" Then
                            If Split(tmpPar.Reserve, "|")(0) = Mid(strTmp, 2) Then
                                opt(opt.UBound).Value = True
                                blnExist = True
                            End If
                        Else
                            If Split(tmpPar.Reserve, "|")(0) = strTmp Then
                                opt(opt.UBound).Value = True
                                blnExist = True
                            End If
                        End If
                    End If
                    
                    opt(opt.UBound).Visible = True
                Next
                
                fra(i).ZOrder 1 '����������
                fra(i).Visible = True
            ElseIf tmpPar.��ʽ = 2 Then '������ѡ��
                lblName(i).Visible = False
                
                blntmp = True
                Load chk(i): Set objTmp = chk(i)
                If tmpPar.�Ƿ����� Then objTmp.Enabled = False
                chk(i).Caption = lblName(i).Caption
                chk(i).TabIndex = intCurTab: intCurTab = intCurTab + 1
                chk(i).Left = chk(0).Left: chk(i).Top = lblName(i).Top - (chk(i).Height - lblName(i).Height) / 2
                chk(i).Width = TextWidth(chk(i).Caption) + 230
                
                '���õķָ���
                If Left(Split(Split(tmpPar.ֵ�б�, "|")(0), ",")(0), 1) = "��" Then chk(i).Value = 1
                For j = 0 To 1
                    strTmp = Split(Split(tmpPar.ֵ�б�, "|")(j), ",")(0)
                    '��������ʱReserve����ϴ���"��ʾֵ|��ֵ"
                    '�����ϴ�ѡ��ֵ����λ����ȱʡ��
                    If tmpPar.Reserve Like "*|*" Then
                        If Left(strTmp, 1) = "��" Then
                            If Split(tmpPar.Reserve, "|")(0) = Mid(strTmp, 2) Then
                                If Left(strTmp, 1) = "��" Then
                                    chk(i).Value = 1
                                Else
                                    chk(i).Value = 0
                                End If
                            End If
                        Else
                            If Split(tmpPar.Reserve, "|")(0) = strTmp Then
                                If Left(strTmp, 1) = "��" Then
                                    chk(i).Value = 1
                                Else
                                    chk(i).Value = 0
                                End If
                            End If
                        End If
                    End If
                Next
                chk(i).Visible = True
            End If
        ElseIf tmpPar.ȱʡֵ = "ѡ�������塭" Then
            Load txt(i): Set objTmp = txt(i)
            If tmpPar.�Ƿ����� Then objTmp.Enabled = False
            txt(i).TabIndex = intCurTab: intCurTab = intCurTab + 1
            txt(i).Left = txt(0).Left: txt(i).Top = lblName(i).Top - (txt(i).Height - lblName(i).Height) / 2
            txt(i).ToolTipText = "�� F2 ��ѡ����"
            txt(i).Locked = True
                                                
            blnCmd = True
            If tmpPar.Reserve Like "*|*" Then
                If Split(tmpPar.Reserve, "|")(0) <> "" Then
                    '��������ʱReserve�����"��ʾֵ|��ֵ"
                    txt(i).Text = Split(tmpPar.Reserve, "|")(0)
                    txt(i).Tag = Split(tmpPar.Reserve, "|")(1)
                    
                    '��Ȼ��ȱʡ,�����û��������ѡ�򲻿ɼ�
                    strTmp = ""
                    If InStr(tmpPar.����, "|") > 0 Then strTmp = Split(tmpPar.����, "|")(0)
                    strTmp = SQLOwner(Replace(RemoveNote(tmpPar.��ϸSQL), "[*]", ""), strTmp)
                    Call CheckParsRela(strTmp, Nothing, tmpPar.����, True, , mobjPars)
                    strTmp = GetDefaultValue(strTmp, tmpPar.��ϸ�ֶ�, , GetDBConnectNo(tmpPar, mobjReport.Datas))
                    If strTmp <> "" Then
                        blnCmd = (CLng((Split(strTmp, "|")(2))) > 0)
                    Else
                        blnCmd = False
                    End If
                Else
                    'ʹ��ȱʡ�����ȱʡֵ
                    If tmpPar.ֵ�б� Like "*|*" Then
                        txt(i).Text = Split(tmpPar.ֵ�б�, "|")(0)
                        txt(i).Tag = Split(tmpPar.ֵ�б�, "|")(1)
                    ElseIf tmpPar.��ϸSQL <> "" Then
                        'ȡ��ϸSQL����е�һ��ֵ,���ֻ��һ��,����ѡ
                        strTmp = ""
                        If InStr(tmpPar.����, "|") > 0 Then strTmp = Split(tmpPar.����, "|")(0)
                        strTmp = SQLOwner(Replace(RemoveNote(tmpPar.��ϸSQL), "[*]", ""), strTmp)
                        Call CheckParsRela(strTmp, Nothing, tmpPar.����, True, , mobjPars)
                        strTmp = GetDefaultValue(strTmp, tmpPar.��ϸ�ֶ�, , GetDBConnectNo(tmpPar, mobjReport.Datas))
                        If strTmp <> "" Then
                            txt(i).Text = Split(strTmp, "|")(0)
                            txt(i).Tag = Split(strTmp, "|")(1)
                            If tmpPar.��ʽ = 1 Then txt(i).Tag = " IN (" & txt(i).Tag & ") "
                            blnCmd = (CLng((Split(strTmp, "|")(2))) > 1)
                        Else
                            blnCmd = False
                        End If
                    End If
                End If
            Else
                If tmpPar.ֵ�б� Like "*|*" Then
                    'ʹ��ȱʡ�����ȱʡֵ
                    txt(i).Text = Split(tmpPar.ֵ�б�, "|")(0)
                    txt(i).Tag = Split(tmpPar.ֵ�б�, "|")(1)
                    
                    '��Ȼ��ȱʡ,�����û��������ѡ�򲻿ɼ�
                    strTmp = ""
                    If InStr(tmpPar.����, "|") > 0 Then strTmp = Split(tmpPar.����, "|")(0)
                    strTmp = SQLOwner(Replace(RemoveNote(tmpPar.��ϸSQL), "[*]", ""), strTmp)
                    Call CheckParsRela(strTmp, Nothing, tmpPar.����, True, , mobjPars)
                    strTmp = GetDefaultValue(strTmp, tmpPar.��ϸ�ֶ�, , GetDBConnectNo(tmpPar, mobjReport.Datas))
                    If strTmp <> "" Then
                        blnCmd = (CLng((Split(strTmp, "|")(2))) > 0)
                    Else
                        blnCmd = False
                    End If
                ElseIf tmpPar.��ϸSQL <> "" Then
                    'ȡ��ϸSQL����е�һ��ֵ,���ֻ��һ��,����ѡ
                    strTmp = ""
                    If InStr(tmpPar.����, "|") > 0 Then strTmp = Split(tmpPar.����, "|")(0)
                    strTmp = SQLOwner(Replace(RemoveNote(tmpPar.��ϸSQL), "[*]", ""), strTmp)
                    Call CheckParsRela(strTmp, Nothing, tmpPar.����, True, , mobjPars)
                    strTmp = GetDefaultValue(strTmp, tmpPar.��ϸ�ֶ�, , GetDBConnectNo(tmpPar, mobjReport.Datas))
                    If strTmp <> "" Then
                        txt(i).Text = Split(strTmp, "|")(0)
                        txt(i).Tag = Split(strTmp, "|")(1)
                        If tmpPar.��ʽ = 1 Then txt(i).Tag = " IN (" & txt(i).Tag & ") "
                        blnCmd = (CLng((Split(strTmp, "|")(2))) > 1)
                    Else
                        blnCmd = False
                    End If
                End If
            End If
                        
            Load cmd(i)
            If tmpPar.�Ƿ����� Then cmd(i).Enabled = False
            cmd(i).Top = txt(i).Top + 30
            cmd(i).Left = txt(i).Left + txt(i).Width - cmd(i).Width - 30
            cmd(i).Height = txt(i).Height - 45
            cmd(i).TabStop = False
            cmd(i).ZOrder
            
            txt(i).Visible = True
            cmd(i).Visible = blnCmd
            
            '�ɷ�����ƥ��
            txt(i).Locked = Not ((InStr(tmpPar.����SQL, "[*]") > 0 Or InStr(tmpPar.��ϸSQL, "[*]") > 0) And blnCmd)
        Else
            If tmpPar.���� = 2 Then
                Load dtp(i): Set objTmp = dtp(i)
                If tmpPar.�Ƿ����� Then objTmp.Enabled = False
                dtp(i).TabIndex = intCurTab: intCurTab = intCurTab + 1
                dtp(i).Left = dtp(0).Left: dtp(i).Top = lblName(i).Top - (dtp(i).Height - lblName(i).Height) / 2
                If InStr(tmpPar.ȱʡֵ, ":") > 0 Or InStr(tmpPar.ȱʡֵ, "ʱ��") > 0 Then
                    dtp(i).CustomFormat = "yyyy��MM��dd�� HH:mm:ss"
                    dtp(i).Width = 2460
                Else
                    dtp(i).CustomFormat = "yyyy��MM��dd��"
                    dtp(i).Width = 1635
                End If
                If tmpPar.ȱʡֵ <> "" Then
                    If Left(tmpPar.ȱʡֵ, 1) = "&" Then
                        dtp(i).Value = GetParVBMacro(tmpPar.ȱʡֵ)
                    Else
                        dtp(i).Value = Format(tmpPar.ȱʡֵ, dtp(i).CustomFormat)
                    End If
                Else
                    dtp(i).Value = Currentdate
                End If
                
'                'ע�����ֵ
'                If dtp(i).CustomFormat Like "*HH:mm:ss" And Left(tmpPar.ȱʡֵ, 1) <> "&" Then
'                    strTmp = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.name & mobjReport.����, lblName(i).ToolTipText & "ʱ��", Format(dtp(i).Value, "HH:mm:ss"))
'                    dtp(i).Value = CDate(Format(dtp(i).Value, Left(dtp(i).CustomFormat, InStr(dtp(i).CustomFormat, "HH:mm:ss") - 1)) & strTmp)
'                End If
                
                dtp(i).Visible = True
            Else
                Load txt(i): Set objTmp = txt(i)
                If tmpPar.�Ƿ����� Then objTmp.Enabled = False
                txt(i).Left = txt(0).Left: txt(i).Top = lblName(i).Top - (txt(i).Height - lblName(i).Height) / 2
                txt(i).TabIndex = intCurTab: intCurTab = intCurTab + 1
                txt(i).Text = tmpPar.ȱʡֵ
                txt(i).Visible = True
            End If
        End If
        If objTmp.name = "fra" Then
            lngCurH = lngCurH + objTmp.Height + 180
        Else
            lngCurH = lngCurH + txt(0).Height + 150
        End If
        
        lblName(i).Tag = tmpPar.���� & "," & objTmp.name
        If tmpPar.ȱʡֵ = "ѡ�������塭" Then lblName(i).Tag = lblName(i).Tag & ",cmd"
    Next
    
    fraSplit.Top = lngCurH
    
    '���������
    For i = 1 To lblName.UBound
        strCur = ""
        If strGroup <> CStr(Split(lblName(i).Tag, ",")(0)) And CStr(Split(lblName(i).Tag, ",")(0)) <> "" Then
            Load fraGroup(fraGroup.UBound + 1)
            Set objGroup = fraGroup(fraGroup.UBound)
            objGroup.Caption = CStr(Split(lblName(i).Tag, ",")(0))
            objGroup.Top = lblName(i).Top - 150
            objGroup.ZOrder 1
            objGroup.Visible = True
            
            Select Case CStr(Split(lblName(i).Tag, ",")(1))
                Case "txt"
                    Set objTmp = txt(i)
                Case "cbo"
                    Set objTmp = cbo(i)
                Case "dtp"
                    Set objTmp = dtp(i)
                Case "chk"
                    Set objTmp = chk(i)
            End Select
            
            lngCurH = 195 '��ǰTopλ��
            
            Set objTmp.Container = objGroup
            objTmp.Top = lngCurH + (300 - objTmp.Height) / 2
            objTmp.Left = 1250
            
            Set lblName(i).Container = objGroup
            lblName(i).Top = objTmp.Top + (objTmp.Height - lblName(i).Height) / 2
            lblName(i).Left = objTmp.Left - lblName(i).Width - 30
            lblName(i).Caption = GetLenStr(lblName(i).ToolTipText, 900, Me) & Mid(lblName(i).Caption, InStr(lblName(i).Caption, "("))
            
            If UBound(Split(lblName(i).Tag, ",")) = 2 Then
                Set cmd(i).Container = objGroup
                cmd(i).Top = objTmp.Top + 30
                cmd(i).Left = objTmp.Left + objTmp.Width - cmd(i).Width - 30
            End If

            lngCurH = lngCurH + txt(0).Height + 50 '��ǰTopλ��
        ElseIf strGroup = CStr(Split(lblName(i).Tag, ",")(0)) And CStr(Split(lblName(i).Tag, ",")(0)) <> "" Then
            strCur = "Add"
            Select Case CStr(Split(lblName(i).Tag, ",")(1))
                Case "txt"
                    Set objTmp = txt(i)
                Case "cbo"
                    Set objTmp = cbo(i)
                Case "dtp"
                    Set objTmp = dtp(i)
                Case "chk"
                    Set objTmp = chk(i)
            End Select
            
            Set objTmp.Container = objGroup
            objTmp.Top = lngCurH + (300 - objTmp.Height) / 2
            objTmp.Left = 1250
            
            Set lblName(i).Container = objGroup
            lblName(i).Top = objTmp.Top + (objTmp.Height - lblName(i).Height) / 2
            lblName(i).Left = objTmp.Left - lblName(i).Width - 30
            lblName(i).Caption = GetLenStr(lblName(i).ToolTipText, 900, Me) & Mid(lblName(i).Caption, InStr(lblName(i).Caption, "("))
            
            If UBound(Split(lblName(i).Tag, ",")) = 2 Then
                Set cmd(i).Container = objGroup
                cmd(i).Top = objTmp.Top + 30
                cmd(i).Left = objTmp.Left + objTmp.Width - cmd(i).Width - 30
            End If
                        
            lngCurH = lngCurH + txt(0).Height + 50 '��ǰTopλ��
            
            objGroup.Height = objTmp.Top + objTmp.Height + 90  '��߶�
            
            '�ÿ����µ���������ȫ������
            For j = i + 1 To lblName.UBound
                If Split(lblName(j).Tag, ",")(0) <> "fra" Then
                    lblName(j).Top = lblName(j).Top + 60
                    Select Case CStr(Split(lblName(j).Tag, ",")(1))
                        Case "txt"
                            txt(j).Top = txt(j).Top + 60
                        Case "cbo"
                            cbo(j).Top = cbo(j).Top + 60
                        Case "dtp"
                            dtp(j).Top = dtp(j).Top + 60
                        Case "chk"
                            chk(j).Top = chk(j).Top + 60
                    End Select
                    If UBound(Split(lblName(j).Tag, ",")) = 2 Then
                        cmd(j).Top = cmd(j).Top + 60
                    End If
                End If
            Next
        End If
        If strPre = "Add" And strCur = "" Then
            fraSplit.Top = fraSplit.Top + 60
        End If
        strPre = strCur
        strGroup = CStr(Split(lblName(i).Tag, ",")(0))
    Next
    
    'û�в����鵫�ж��ѡ��ʱ,��ÿ����
    If fraGroup.UBound = 0 And fra.UBound > 0 Then
        For Each objTmp In fra
            objTmp.Left = txt(0).Left - 1000
        Next
    End If
    
    cmdLoad.Top = fraSplit.Top + 180
    cmdDefault.Top = fraSplit.Top + 180
    
    fraSplit.Visible = (lblName.UBound > 0)
    cmdLoad.Visible = (lblName.UBound > 0)
    cmdDefault.Visible = (lblName.UBound > 0)
    
    cmdLoad.TabIndex = intCurTab: intCurTab = intCurTab + 1
    cmdDefault.TabIndex = intCurTab
    
    cmdSelAll.Top = cmdLoad.Top: cmdSelNone.Top = cmdSelAll.Top
    cmdSelAll.Visible = blntmp
    cmdSelNone.Visible = blntmp
    If Me.Visible Then
        On Error Resume Next
        If picPar.Height < cmdLoad.Top + cmdLoad.Height + 100 Then
            lngTmp = cmdLoad.Top + cmdLoad.Height + 100 - picPar.Height
            picPar.Height = picPar.Height + lngTmp
            picPar.Top = picPar.Top - lngTmp: lblPar_S.Top = lblPar_S.Top - lngTmp
            lvw.Height = lvw.Height - lngTmp
        End If
    End If
    
    '���µ����˵�
    Call LoadCondsMenu
End Sub

Private Function ReSetReportPars() As Boolean
'���ܣ��������õ�ǰ�������û�����Ĳ���
    Dim i As Integer, j As Integer
    Dim strTmp As String, strDisp As String
    Dim strParName As String, curDate As Date
    
    '�ȼ��Ϸ���
    For i = 1 To lblName.UBound
        strParName = lblName(i).ToolTipText
        
        If mobjPars("_" & strParName).ȱʡֵ = "�̶�ֵ�б�" Then
            Select Case mobjPars("_" & strParName).��ʽ
                Case 0
                    If Trim(cbo(i).Text) = "" Then
                        MsgBox "��ѡ��""" & strParName & """������ֵ��", vbInformation, App.Title
                        If cbo(i).Enabled And cbo(i).Visible Then cbo(i).SetFocus
                        Exit Function
                    End If
                    If GetCboIndex(cbo(i), cbo(i).Text) = -1 Then '�Ƿ���Ϊ����
                        '���ͼ��
                        Select Case mobjPars("_" & strParName).����
                            Case 1
                                If Not IsNumeric(cbo(i).Text) Then
                                    MsgBox "�������""" & strParName & """������ֵ����Ӧ��Ϊ�����ͣ�", vbInformation, App.Title
                                    If cbo(i).Enabled And cbo(i).Visible Then cbo(i).SetFocus
                                    Exit Function
                                End If
                            Case 2
                                If Not IsDate(cbo(i).Text) Then
                                    MsgBox "�������""" & strParName & """������ֵ����Ӧ��Ϊ�����ͣ�", vbInformation, App.Title
                                    If cbo(i).Enabled And cbo(i).Visible Then cbo(i).SetFocus
                                    Exit Function
                                End If
                        End Select
                    End If
            End Select
        ElseIf mobjPars("_" & strParName).ȱʡֵ = "ѡ�������塭" Then
            If Trim(txt(i).Text) = "" Then
                MsgBox "��ѡ��""" & strParName & """������ֵ��", vbInformation, App.Title
                If txt(i).Enabled And txt(i).Visible Then txt(i).SetFocus
                Exit Function
            End If
            If txt(i).Tag = "" Then '�Ƿ���Ϊ����
                If mobjPars("_" & strParName).ֵ�б� Like "*|*" Then
                    If Split(mobjPars("_" & strParName).ֵ�б�, "|")(0) <> txt(i).Text Then
                        '���ͼ��
                        Select Case mobjPars("_" & strParName).����
                            Case 1
                                If Not IsNumeric(txt(i).Text) Then
                                    MsgBox "�������""" & strParName & """������ֵ����Ӧ��Ϊ�����ͣ�", vbInformation, App.Title
                                    If txt(i).Enabled And txt(i).Visible Then txt(i).SetFocus
                                    Exit Function
                                End If
                            Case 2
                                If Not IsDate(txt(i).Text) Then
                                    MsgBox "�������""" & strParName & """������ֵ����Ӧ��Ϊ�����ͣ�", vbInformation, App.Title
                                    If txt(i).Enabled And txt(i).Visible Then txt(i).SetFocus
                                    Exit Function
                                End If
                        End Select
                    Else
                        '����ֵ�붨���ȱʡֵ��ͬ,��ԭΪȱʡֵ
                        txt(i).Tag = Split(mobjPars("_" & strParName).ֵ�б�, "|")(1)
                    End If
                Else
                    '���ͼ��
                    Select Case mobjPars("_" & strParName).����
                        Case 1
                            If Not IsNumeric(txt(i).Text) Then
                                MsgBox "�������""" & strParName & """������ֵ����Ӧ��Ϊ�����ͣ�", vbInformation, App.Title
                                If txt(i).Enabled And txt(i).Visible Then txt(i).SetFocus
                                Exit Function
                            End If
                        Case 2
                            If Not IsDate(txt(i).Text) Then
                                MsgBox "�������""" & strParName & """������ֵ����Ӧ��Ϊ�����ͣ�", vbInformation, App.Title
                                If txt(i).Enabled And txt(i).Visible Then txt(i).SetFocus
                                Exit Function
                            End If
                    End Select
                End If
            End If
        Else
            Select Case mobjPars("_" & strParName).����
                Case 0, 3
                    If Trim(txt(i).Text) = "" Then
                        MsgBox "������""" & strParName & """������ֵ��", vbInformation, App.Title
                        If txt(i).Enabled And txt(i).Visible Then txt(i).SetFocus
                        Exit Function
                    End If
                    If TLen(txt(i).Text) > 255 Then
                        MsgBox """" & strParName & """������ֵ���Ȳ��ܳ���255���ַ���", vbInformation, App.Title
                        If txt(i).Enabled And txt(i).Visible Then txt(i).SetFocus
                        Exit Function
                    End If
                Case 1
                    If Trim(txt(i).Text) = "" Then
                        MsgBox "������""" & strParName & """������ֵ��", vbInformation, App.Title
                        If txt(i).Enabled And txt(i).Visible Then txt(i).SetFocus
                        Exit Function
                    End If
                    If TLen(txt(i).Text) > 255 Then
                        MsgBox """" & strParName & """������ֵ���Ȳ��ܳ���255���ַ���", vbInformation, App.Title
                        If txt(i).Enabled And txt(i).Visible Then txt(i).SetFocus
                        Exit Function
                    End If
                    If Not IsNumeric(txt(i).Text) Then
                        MsgBox """" & strParName & """������ֵ����Ӧ��Ϊ�����ͣ�", vbInformation, App.Title
                        If txt(i).Enabled And txt(i).Visible Then txt(i).SetFocus
                        Exit Function
                    End If
                Case 2 '����ʱ�����ֵ���
                    curDate = Currentdate
                    If Not (mobjPars("_" & strParName).ȱʡֵ Like "&��һ*" Or mobjPars("_" & strParName).Reserve Like "&��һ*" Or _
                        mobjPars("_" & strParName).ȱʡֵ Like "&��һ*" Or mobjPars("_" & strParName).Reserve Like "&��һ*" Or _
                        mobjPars("_" & strParName).ȱʡֵ Like "&*����*" Or mobjPars("_" & strParName).Reserve Like "&*����*" Or _
                        mobjPars("_" & strParName).ȱʡֵ Like "&*��ĩ*" Or mobjPars("_" & strParName).ȱʡֵ Like "&*��ĩ*" Or _
                        mobjPars("_" & strParName).Reserve Like "&*��ĩ*" Or mobjPars("_" & strParName).Reserve Like "&*��ĩ*") Then
                        
                        If mobjPars("_" & strParName).ȱʡֵ Like "*ʱ��*" Or mobjPars("_" & strParName).Reserve Like "*ʱ��*" Then
                            If Format(dtp(i).Value, "yyyy-MM-dd HH:mm:ss") > Format(curDate, "yyyy-MM-dd HH:mm:ss") Then
                                MsgBox """" & strParName & """ ������ֵ���ܳ�����ǰʱ�䣡", vbInformation, App.Title
                                If dtp(i).Enabled And dtp(i).Visible Then dtp(i).SetFocus
                                Exit Function
                            End If
                        Else
                            If Format(dtp(i).Value, "yyyy-MM-dd") > Format(curDate, "yyyy-MM-dd") Then
                                MsgBox """" & strParName & """ ������ֵ���ܳ�����ǰ���ڣ�", vbInformation, App.Title
                                If dtp(i).Enabled And dtp(i).Visible Then dtp(i).SetFocus
                                Exit Function
                            End If
                        End If
                    End If
            End Select
        End If
    Next
        
    '��ȡֵ
    For i = 1 To lblName.UBound
        strParName = lblName(i).ToolTipText
        
        If mobjPars("_" & strParName).ȱʡֵ = "�̶�ֵ�б�" Then '���õķָ���
            Select Case mobjPars("_" & strParName).��ʽ
                Case 0
                    If GetCboIndex(cbo(i), cbo(i).Text) = -1 Then '�Ƿ���Ϊ����
                        'Reserve�ֶα��汾��������"������ֵ|��ʾֵ"
                        mobjPars("_" & strParName).Reserve = "�̶�ֵ�б�|" & cbo(i).Text
                        mobjPars("_" & strParName).ȱʡֵ = cbo(i).Text
                    Else
                        '�б�ѡ��
                        'Reserve�ֶα��汾��������"������ֵ|��ʾֵ"
                        mobjPars("_" & strParName).Reserve = "�̶�ֵ�б�|" & cbo(i).Text
                        strTmp = mobjPars("_" & strParName).ֵ�б�
                        For j = 0 To UBound(Split(strTmp, "|"))
                            strDisp = Split(Split(strTmp, "|")(j), ",")(0)
                            If Left(strDisp, 1) = "��" Then strDisp = Mid(strDisp, 2)
                            If strDisp = cbo(i).Text Then
                                mobjPars("_" & strParName).ȱʡֵ = Split(Split(strTmp, "|")(j), ",")(1)
                                Exit For
                            End If
                        Next
                    End If
                Case 1
                    For j = 1 To opt.UBound
                        If opt(j).Container.Index = i Then
                            If opt(j).Value Then
                                'Reserve�ֶα��汾��������"������ֵ|��ʾֵ"
                                mobjPars("_" & strParName).Reserve = "�̶�ֵ�б�|" & opt(j).ToolTipText
                                mobjPars("_" & strParName).ȱʡֵ = opt(j).Tag
                            End If
                        End If
                    Next
                Case 2
                    'Reserve�ֶα��汾��������"������ֵ|��ʾֵ"
                    strTmp = mobjPars("_" & strParName).ֵ�б�
                    For j = 0 To 1
                        strDisp = Split(Split(strTmp, "|")(j), ",")(0)
                        If chk(i).Value = 0 Then
                            If Left(strDisp, 1) <> "��" Then
                                mobjPars("_" & strParName).Reserve = "�̶�ֵ�б�|" & strDisp
                                mobjPars("_" & strParName).ȱʡֵ = Split(Split(strTmp, "|")(j), ",")(1)
                            End If
                        Else
                            If Left(strDisp, 1) = "��" Then
                                mobjPars("_" & strParName).Reserve = "�̶�ֵ�б�|" & Mid(strDisp, 2)
                                mobjPars("_" & strParName).ȱʡֵ = Split(Split(strTmp, "|")(j), ",")(1)
                            End If
                        End If
                    Next
            End Select
        ElseIf mobjPars("_" & strParName).ȱʡֵ = "ѡ�������塭" Then
            If txt(i).Tag = "" Then '�Ƿ���Ϊ����
                'Reserve�ֶα��汾��������"������ֵ|��ʾֵ"
                mobjPars("_" & strParName).Reserve = "ѡ�������塭|"
                mobjPars("_" & strParName).ȱʡֵ = txt(i).Text
            Else
                '�б�ѡ��
                'Reserve�ֶα��汾��������"������ֵ|��ʾֵ"
                mobjPars("_" & strParName).Reserve = "ѡ�������塭|" & txt(i).Text
                mobjPars("_" & strParName).ȱʡֵ = txt(i).Tag
            End If
        Else
            Select Case mobjPars("_" & strParName).����
                Case 0, 1, 3
                    mobjPars("_" & strParName).ȱʡֵ = txt(i).Text
                Case 2
                    If mobjPars("_" & strParName).ȱʡֵ Like "&*" Then
                        mobjPars("_" & strParName).Reserve = mobjPars("_" & strParName).ȱʡֵ
                    End If
                    mobjPars("_" & strParName).ȱʡֵ = Format(dtp(i).Value, dtp(i).CustomFormat)
                    '���浽ע���
                    If dtp(i).CustomFormat Like "*HH:mm:ss" Then
                        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.name & mobjReport.����, lblName(i).ToolTipText & "ʱ��", Format(dtp(i).Value, "HH:mm:ss")
                    End If
            End Select
        End If
    Next
    
    Call ReplaceInputPars(mobjPars)
    
    ReSetReportPars = True
End Function

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim LngIdx As Long
    
    If KeyAscii = 13 Then KeyAscii = 0: SendKeys "{Tab}": Exit Sub
    If InStr("~`!@#$^&"";|", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If mobjPars("_" & lblName(Index).ToolTipText).���� = 1 Then
        If InStr("-0.123456789" & Chr(8) & Chr(3) & Chr(22) & Chr(24) & Chr(26), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
    End If
    
    If KeyAscii <> 8 Then
        If SendMessage(cbo(Index).hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then SendKeys "{F4}"
        LngIdx = MatchIndex(cbo(Index), KeyAscii)
        If LngIdx <> -2 Then cbo(Index).ListIndex = LngIdx
    End If
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: SendKeys "{Tab}": Exit Sub
End Sub

Private Function GetValues() As Collection
'���ܣ���ȡ���еĽ����ϵĲ���ֵ
    Dim i As Integer, j As Integer
    Dim strParName As String, strTmp As String
    Dim strDisp As String, colValue As New Collection
     
    For i = 1 To lblName.UBound
        strParName = lblName(i).ToolTipText
        
        If mobjPars("_" & strParName).ȱʡֵ = "�̶�ֵ�б�" Then
            Select Case mobjPars("_" & strParName).��ʽ
                Case 0
                    If GetCboIndex(cbo(i), cbo(i).Text) = -1 Then '�Ƿ���Ϊ����
                        'Reserve�ֶα��汾��������"������ֵ|��ʾֵ"
                        colValue.Add cbo(i).Text, "_" & strParName
                    Else
                        '�б�ѡ��
                        'Reserve�ֶα��汾��������"������ֵ|��ʾֵ"
                        '���õķָ���
                        strTmp = mobjPars("_" & strParName).ֵ�б�
                        For j = 0 To UBound(Split(strTmp, "|"))
                            strDisp = Split(Split(strTmp, "|")(j), ",")(0)
                            If Left(strDisp, 1) = "��" Then strDisp = Mid(strDisp, 2)
                            If strDisp = cbo(i).Text Then
                                colValue.Add Split(Split(strTmp, "|")(j), ",")(1), "_" & strParName
                                Exit For
                            End If
                        Next
                    End If
                Case 1
                    For j = 1 To opt.UBound
                        If opt(j).Container.Index = i Then
                            If opt(j).Value Then
                                colValue.Add opt(j).Tag, "_" & strParName
                            End If
                        End If
                    Next
                Case 2
                    'Reserve�ֶα��汾��������"������ֵ|��ʾֵ"
                    '���õķָ���
                    strTmp = mobjPars("_" & strParName).ֵ�б�
                    For j = 0 To 1
                        strDisp = Split(Split(strTmp, "|")(j), ",")(0)
                        If chk(i).Value = 0 Then
                            If Left(strDisp, 1) <> "��" Then
                                colValue.Add Split(Split(strTmp, "|")(j), ",")(1), "_" & strParName
                            End If
                        Else
                            If Left(strDisp, 1) = "��" Then
                                colValue.Add Split(Split(strTmp, "|")(j), ",")(1), "_" & strParName
                            End If
                        End If
                    Next
            End Select
        ElseIf mobjPars("_" & strParName).ȱʡֵ = "ѡ�������塭" Then
            If txt(i).Tag = "" Then '�Ƿ���Ϊ����
                'Reserve�ֶα��汾��������"������ֵ|��ʾֵ"
                colValue.Add txt(i).Text, "_" & strParName
            Else
                '�б�ѡ��
                'Reserve�ֶα��汾��������"������ֵ|��ʾֵ"
                colValue.Add txt(i).Tag, "_" & strParName
            End If
        Else
            Select Case mobjPars("_" & strParName).����
                Case 0, 1, 3
                    colValue.Add txt(i).Text, "_" & strParName
                Case 2
                    colValue.Add Format(dtp(i).Value, dtp(i).CustomFormat), "_" & strParName
            End Select
        End If
    Next
    Set GetValues = colValue
End Function

Private Sub cmd_Click(Index As Integer)
    Dim tmpPar As RPTPar, str��ϸ���� As String, str������� As String
    Dim frmNewSelect As New frmSelect
    Dim strSQL��ϸ As String, strSQL���� As String
    Dim colValue As New Collection    '�������е�ֵ
    
    For Each tmpPar In mobjPars
        If tmpPar.���� = lblName(Index).ToolTipText Then
            If blnMatch And txt(Index).Tag = "" Then frmNewSelect.strMatch = txt(Index).Text
            
            If InStr(tmpPar.����, "|") > 0 Then
                str��ϸ���� = Split(tmpPar.����, "|")(0)
                str������� = Split(tmpPar.����, "|")(1)
            End If
            strSQL��ϸ = tmpPar.��ϸSQL
            strSQL���� = tmpPar.����SQL
            Set colValue = GetValues
            Call CheckParsRela(strSQL��ϸ, Nothing, tmpPar.����, True, colValue, mobjPars)
            Call CheckParsRela(strSQL����, Nothing, tmpPar.����, True, colValue, mobjPars)
            frmNewSelect.strSQLList = SQLOwner(RemoveNote(strSQL��ϸ), str��ϸ����)
            frmNewSelect.strSQLTree = SQLOwner(RemoveNote(strSQL����), str�������)
            frmNewSelect.strFLDList = tmpPar.��ϸ�ֶ�
            frmNewSelect.strFLDTree = tmpPar.�����ֶ�
            frmNewSelect.strParName = tmpPar.����
            frmNewSelect.bytType = tmpPar.����
            frmNewSelect.mblnMulti = tmpPar.��ʽ = 1
            frmNewSelect.mintConnect = GetDBConnectNo(tmpPar, mobjReport.Datas)
            frmNewSelect.lngSeekHwnd = cmd(Index).hwnd
            
            On Error Resume Next
            Err.Clear
            
            frmNewSelect.Show 1, Me
            If frmNewSelect.mblnOK Then
                txt(Index).Text = frmNewSelect.strOutDisp
                txt(Index).Tag = frmNewSelect.strOutBand
                Unload frmNewSelect
                SendKeys "{Tab}"
            ElseIf blnMatch Then
                txt(Index).Text = ""
                txt(Index).Tag = ""
            End If
            
            blnMatch = False
            Exit For
        End If
    Next
    txt(Index).SetFocus
End Sub

Private Sub dtp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Private Sub dtp_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub opt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: SendKeys "{Tab}": Exit Sub
End Sub

Private Sub txt_GotFocus(Index As Integer)
    SelAll txt(Index)
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 And txt(Index).ToolTipText <> "" Then
        If cmd(Index).Enabled And cmd(Index).Visible Then Call cmd_Click(Index)
    End If
    If txt(Index).Locked Then Exit Sub
    
    '��Ϊ����ʱ(��ѡ��)�������ֵ��Ϊ��Ϊ����ı�־
    '144=Num;112-123=F1-F12;229=��ʼ���뺺��
    If KeyCode >= 48 And KeyCode <> 144 And KeyCode <> 229 _
        And Not (KeyCode >= 112 And KeyCode <= 123) Then
        txt(Index).Tag = ""
    End If
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txt(Index).Tag = "" And txt(Index).ToolTipText <> "" Then
            '������ƥ��
            KeyAscii = 0
            If txt(Index).Text <> "" Then
                If cmd(Index).Enabled And cmd(Index).Visible Then
                    blnMatch = True
                    Call cmd_Click(Index)
                End If
            End If
            Exit Sub
        Else
            '���ƶ�����
            KeyAscii = 0: SendKeys "{Tab}": Exit Sub
        End If
    End If
    
    If txt(Index).Locked Then Exit Sub
    
    If InStr("~`!@#$^&"";|'" & Chr(3) & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0
    
    If txt(Index).ToolTipText = "" And mobjPars("_" & lblName(Index).ToolTipText).���� = 1 Then
        If InStr("-0.123456789" & Chr(8) & Chr(24) & Chr(26), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    '��Ϊ����ʱ(��ѡ��)�������ֵ��Ϊ��Ϊ����ı�־
    '����ֻ������,������KeyDown�д���
    If KeyAscii < 0 Then txt(Index).Tag = ""
End Sub

Private Sub ShowStatGrid(objItem As RPTItem)
'���ܣ��ڲ�ѯ��������֯��ʾһ��������(�����������ӱ��)
    Dim mshBody As Object, tmpItem As RPTItem, tmpID As RelatID
    Dim rsGroup As ADODB.Recordset, rsVsc As ADODB.Recordset, rsHsc As ADODB.Recordset
    Dim arrStat() As Variant, strVscStat As String, strHscStat As String, strStat As String
    Dim strVsc As String, strHsc As String, strVscOrder As String, strHscOrder As String
    Dim strFilter As String, strAlign As String, strTmp As String
    Dim i As Long, j As Long, k As Long, l As Long, M As Long
    Dim X As Long, Y As Long, Z As Long '����������
    Dim strFormat As String, strSort As String, blnHide As Boolean, blnDo As Boolean
    Dim arrLevel() As String, arrMerge() As String, arrCount() As Long
    Dim arrHead() As String
        
    '���������ӱ�����ӱ���
    Dim lngCurCols As Long, objCurItem As RPTItem
    Dim strLink As String, lngMaxY As Long
    Dim lngGrid As Long, strTopRow As String
    Dim lngDiff As Long
    Dim lngStatistics As Long
    
    '�Ľ����������������
    Dim colVsc As Collection, colHsc As Collection
    Dim strKey As String, lngRow As Long, lngCol As Long, StrFmt As String
    
    Dim varIFValue As Variant
    Dim objColProp As RPTColProterty
    Dim objStatusGridItem As RPTItem
    Dim colTmp As Collection
    
    On Error GoTo hErr
    
    With objItem
        Load msh(.id)
        Set msh(.id).Container = picPaper(intReport)
        Set mshBody = msh(.id)
        
        mshBody.Redraw = False
        
        mshBody.ForeColor = .ǰ��
        mshBody.ForeColorFixed = .ǰ��
        mshBody.BackColor = .����
        mshBody.BackColorFixed = .����
        mshBody.GridColor = .����
        mshBody.GridColorFixed = .����
        mshBody.Font.name = .����
        mshBody.Font.Size = .�ֺ�
        mshBody.Font.Bold = .����
        mshBody.Font.Italic = .б��
        mshBody.Font.Underline = .����
        mshBody.GridLineWidth = IIF(.����߼Ӵ�, 2, 1)
        'Set mshBody.FontFixed = mshBody.Font
        
        mshBody.Left = .X: mshBody.Top = .Y
        mshBody.Height = .H: mshBody.Width = 0
        mshBody.FixedRows = 0
    
        '��ȡ��ӽӱ�������Ϣ
        strLink = strLink & "|" & .id
        For Each tmpItem In mobjReport.Items
            If tmpItem.��ʽ�� = bytFormat And tmpItem.���� = 5 And tmpItem.���� = 2 And tmpItem.���� = .���� Then
                strLink = strLink & "|" & tmpItem.id
            End If
        Next
        strLink = Mid(strLink, 2)
    End With
        
    objItem.��ͷ = ""
    strTopRow = ""
    
    blnHide = True
    lngCurCols = 0
    lngMaxY = 0
    For lngGrid = 0 To UBound(Split(strLink, "|"))
        Set objCurItem = mobjReport.Items("_" & Split(strLink, "|")(lngGrid))
        With objCurItem
            mshBody.Width = mshBody.Width + .W
            
            'ͳ��������
            '�����ӵı���ٴ����������
            If lngGrid = 0 Then
                strVsc = "": strVscOrder = "": X = 0
            End If
            strHsc = "": strHscOrder = "" '��ź�������Ŀ���Ƽ������ֶ�����
            Y = 0: Z = 0
            For Each tmpID In .SubIDs
                Set tmpItem = mobjReport.Items("_" & tmpID.id)
                Select Case tmpItem.����
                    Case 7 '�������
                        If lngGrid = 0 Then
                            X = X + 1
                            If tmpItem.���� <> "" Then strVscOrder = strVscOrder & "|" & tmpItem.����
                            strVsc = strVsc & "|" & tmpItem.����
                        End If
                    Case 8 '�������
                        Y = Y + 1
                        If tmpItem.���� <> "" Then strHscOrder = strHscOrder & "|" & tmpItem.����
                        strHsc = strHsc & "|" & tmpItem.����
                    Case 9 'ͳ����
                        Z = Z + 1
                End Select
            Next
            If Y > lngMaxY Then lngMaxY = Y
            If lngGrid = 0 Then
                strVsc = Mid(strVsc, 2)
                strVscOrder = Mid(strVscOrder, 2)
            End If
            strHsc = Mid(strHsc, 2)
            strHscOrder = Mid(strHscOrder, 2)
            
            '����ձ���
            If lngGrid = 0 Then
                mshBody.FixedRows = Y + 1
            Else
                If Y + 1 > mshBody.FixedRows Then
                    lngDiff = Y + 1 - mshBody.FixedRows '������λƫ��(������ʱ�������ӹ̶�����)
                    For i = 1 To Y + 1 - mshBody.FixedRows
                        mshBody.AddItem "", mshBody.FixedRows
                        mshBody.FixedRows = mshBody.FixedRows + 1
                        For j = 0 To mshBody.Cols - 1 '�����º��������ʱ����ǰһ������������
                            mshBody.TextMatrix(mshBody.FixedRows - 1, j) = mshBody.TextMatrix(mshBody.FixedRows - 2, j)
                        Next
                    Next
                End If
            End If
            mshBody.Cols = lngCurCols + IIF(lngGrid = 0, X, 0) + Z
            If lngGrid = 0 Then
                mshBody.Rows = Y + 2
                mshBody.FixedCols = X
            End If
            lngStatistics = 0
            For Each tmpID In .SubIDs
                Set tmpItem = mobjReport.Items("_" & tmpID.id)
                Select Case tmpItem.����
                    Case 7 '�������
                        If lngGrid = 0 Then
                            For i = 0 To Y
                                mshBody.TextMatrix(i, tmpItem.���) = tmpItem.����
                            Next
                        End If
                    Case 8 '�������
                    Case 9 'ͳ����
                        lngStatistics = lngStatistics + 1
                        For i = mshBody.FixedRows - 1 To Y
                            mshBody.TextMatrix(i, lngCurCols + IIF(lngGrid = 0, X, 0) + tmpItem.���) = tmpItem.����
                        Next
                End Select
            Next
            
            '-------------------------------------------------------------------------------------
            '����������
            '-------------------------------------------------------------------------------------
            If mLibDatas("_" & .����).DataSet.RecordCount > 0 Then
                Set rsGroup = Nothing
                Set rsGroup = mLibDatas("_" & .����).DataSet.Clone
                
                '1.���ɱ����
                
                '1.1:������������ͷ����(���������ֶ�)
                If lngGrid = 0 Then
                    Set rsVsc = Nothing
                    Set rsVsc = New ADODB.Recordset
                    '1.1.1:��������������ֶ�
                    For i = 0 To UBound(Split(strVscOrder, "|"))
                        If Left(Split(strVscOrder, "|")(i), 1) = "," Then
                            With rsGroup.Fields(Mid(Split(strVscOrder, "|")(i), 2))
                                '������adNumeric������ʱ����,���滻ΪadBigInt��adSingle/adDouble
                                rsVsc.Fields.Append .name, IIF(IsType(.type, adNumeric), IIF(.NumericScale = 0, adBigInt, adDouble), IIF(.type = adWChar, adVarWChar, .type)), .DefinedSize
                            End With
                        Else
                            With rsGroup.Fields(Split(strVscOrder, "|")(i))
                                rsVsc.Fields.Append .name, IIF(IsType(.type, adNumeric), IIF(.NumericScale = 0, adBigInt, adDouble), IIF(.type = adWChar, adVarWChar, .type)), .DefinedSize
                            End With
                        End If
                    Next
                    '1.1.2:�������������ֶ�(���������ֶ��ظ�)
                    For i = 0 To UBound(Split(strVsc, "|"))
                        If InStr("|" & Replace(strVscOrder, ",", "") & "|", "|" & Split(strVsc, "|")(i) & "|") = 0 Then
                            With rsGroup.Fields(Split(strVsc, "|")(i))
                                rsVsc.Fields.Append .name, IIF(IsType(.type, adNumeric), IIF(.NumericScale = 0, adBigInt, adDouble), IIF(.type = adWChar, adVarWChar, .type)), .DefinedSize
                            End With
                        End If
                    Next
                    rsVsc.CursorLocation = adUseClient
                    rsVsc.LockType = adLockBatchOptimistic
                    rsVsc.CursorType = adOpenStatic
                    rsVsc.Open
                End If
                
                '1.2:�����������ͷ����(���������ֶ�)
                Set rsHsc = Nothing
                If strHsc <> "" Then
                    Set rsHsc = New ADODB.Recordset
                    '1.2.1:����Ӻ��������ֶ�
                    For i = 0 To UBound(Split(strHscOrder, "|"))
                        If Left(Split(strHscOrder, "|")(i), 1) = "," Then
                            With rsGroup.Fields(Mid(Split(strHscOrder, "|")(i), 2))
                                rsHsc.Fields.Append .name, IIF(IsType(.type, adNumeric), IIF(.NumericScale = 0, adBigInt, adDouble), IIF(.type = adWChar, adVarWChar, .type)), .DefinedSize
                            End With
                        Else
                            With rsGroup.Fields(Split(strHscOrder, "|")(i))
                                rsHsc.Fields.Append .name, IIF(IsType(.type, adNumeric), IIF(.NumericScale = 0, adBigInt, adDouble), IIF(.type = adWChar, adVarWChar, .type)), .DefinedSize
                            End With
                        End If
                    Next
                    '1.2.2:����Ӻ�������ֶ�(���������ֶ��ظ�)
                    For i = 0 To UBound(Split(strHsc, "|"))
                        If InStr("|" & Replace(strHscOrder, ",", "") & "|", "|" & Split(strHsc, "|")(i) & "|") = 0 Then
                            With rsGroup.Fields(Split(strHsc, "|")(i))
                                rsHsc.Fields.Append .name, IIF(IsType(.type, adNumeric), IIF(.NumericScale = 0, adBigInt, adDouble), IIF(.type = adWChar, adVarWChar, .type)), .DefinedSize
                            End With
                        End If
                    Next
                    rsHsc.CursorLocation = adUseClient
                    rsHsc.LockType = adLockBatchOptimistic
                    rsHsc.CursorType = adOpenStatic
                    rsHsc.Open
                End If
                
                '1.3:��ӱ�ͷ���ݼ�
                rsGroup.MoveFirst
                For i = 1 To rsGroup.RecordCount
                    '�����ͷ
                    If Not rsVsc Is Nothing And lngGrid = 0 Then
                        strFilter = "" '�Ƿ��Ѿ�����÷�����ֵ
                        For j = 0 To UBound(Split(strVsc, "|")) '��ʱ���������ֶ�
                            strFilter = strFilter & " And " & Split(strVsc, "|")(j) & "="
                            Select Case rsGroup.Fields(Split(strVsc, "|")(j)).type
                                Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
                                    If Not IsNull(rsGroup.Fields(Split(strVsc, "|")(j)).Value) Then
                                        strFilter = strFilter & "'" & Replace(rsGroup.Fields(Split(strVsc, "|")(j)).Value, " ", "���") & "'"
                                    Else
                                        strFilter = strFilter & "'#'"
                                    End If
                                Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                                    If Not IsNull(rsGroup.Fields(Split(strVsc, "|")(j)).Value) Then
                                        strFilter = strFilter & rsGroup.Fields(Split(strVsc, "|")(j)).Value
                                    Else
                                        strFilter = strFilter & "123456707654321"
                                    End If
                                Case adDBTimeStamp, adDBTime, adDBDate, adDate
                                    If Not IsNull(rsGroup.Fields(Split(strVsc, "|")(j)).Value) Then
                                        '�����ʽ������ȷʶ���ʽ,��#02-4-9#���ϳ�"2009-02-04"
                                        strFilter = strFilter & "#" & Format(rsGroup.Fields(Split(strVsc, "|")(j)).Value, "yyyy-MM-dd HH:mm:ss") & "#"
                                    Else
                                        strFilter = strFilter & "#3000-05-05#"
                                    End If
                            End Select
                        Next
                        rsVsc.Filter = Replace(Mid(strFilter, 6), "���", " ")
                        If rsVsc.EOF Then
                            rsVsc.AddNew
                            For j = 0 To rsVsc.Fields.count - 1 '�����µķ�����ֵ
                                If Not IsNull(rsGroup.Fields(rsVsc.Fields(j).name).Value) Then
                                    rsVsc.Fields(j).Value = rsGroup.Fields(rsVsc.Fields(j).name).Value
                                Else
                                    Select Case rsGroup.Fields(rsVsc.Fields(j).name).type
                                        Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
                                            rsVsc.Fields(j).Value = "#" '�ձ�־
                                        Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                                            rsVsc.Fields(j).Value = 123456707654321# '�ձ�־
                                        Case adDBTimeStamp, adDBTime, adDBDate, adDate
                                            rsVsc.Fields(j).Value = #5/5/3000#   '�ձ�־
                                    End Select
                                End If
                            Next
                        End If
                    End If
                    '�����ͷ
                    If Not rsHsc Is Nothing Then
                        strFilter = "" '�Ƿ��Ѿ�����÷�����ֵ
                        For j = 0 To UBound(Split(strHsc, "|")) '��ʱ���������ֶ�
                            strFilter = strFilter & " And " & Split(strHsc, "|")(j) & "="
                            Select Case rsGroup.Fields(Split(strHsc, "|")(j)).type
                                Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
                                    If Not IsNull(rsGroup.Fields(Split(strHsc, "|")(j)).Value) Then
                                        strFilter = strFilter & "'" & rsGroup.Fields(Split(strHsc, "|")(j)).Value & "'"
                                    Else
                                        strFilter = strFilter & "'#'"
                                    End If
                                Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                                    If Not IsNull(rsGroup.Fields(Split(strHsc, "|")(j)).Value) Then
                                        strFilter = strFilter & rsGroup.Fields(Split(strHsc, "|")(j)).Value
                                    Else
                                        strFilter = strFilter & "123456707654321"
                                    End If
                                Case adDBTimeStamp, adDBTime, adDBDate, adDate
                                    If Not IsNull(rsGroup.Fields(Split(strHsc, "|")(j)).Value) Then
                                        strFilter = strFilter & "#" & Format(rsGroup.Fields(Split(strHsc, "|")(j)).Value, "yyyy-MM-dd HH:mm:ss") & "#"
                                    Else
                                        strFilter = strFilter & "#3000-05-05#"
                                    End If
                            End Select
                        Next
                        rsHsc.Filter = Mid(strFilter, 6)
                        If rsHsc.EOF Then
                            rsHsc.AddNew
                            For j = 0 To rsHsc.Fields.count - 1 '�����µķ�����ֵ
                                If Not IsNull(rsGroup.Fields(rsHsc.Fields(j).name).Value) Then
                                    rsHsc.Fields(j).Value = rsGroup.Fields(rsHsc.Fields(j).name).Value
                                Else
                                    Select Case rsGroup.Fields(rsHsc.Fields(j).name).type
                                        Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
                                            rsHsc.Fields(j).Value = "#" '�ձ�־
                                        Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                                            rsHsc.Fields(j).Value = 123456707654321# '�ձ�־
                                        Case adDBTimeStamp, adDBTime, adDBDate, adDate
                                            rsHsc.Fields(j).Value = #5/5/3000#   '�ձ�־
                                    End Select
                                End If
                            Next
                        End If
                    End If
                    rsGroup.MoveNext
                Next
                If Not rsVsc Is Nothing And lngGrid = 0 Then
                    rsVsc.UpdateBatch adAffectAllChapters
                    rsVsc.Filter = 0
                End If
                If Not rsHsc Is Nothing Then
                    rsHsc.UpdateBatch adAffectAllChapters
                    rsHsc.Filter = 0
                End If
                
                '1.4:��ͷ���ݰ���������
                If Not rsVsc Is Nothing And lngGrid = 0 Then
                    strSort = ""
                    For i = 0 To UBound(Split(strVscOrder, "|"))
                        If Left(Split(strVscOrder, "|")(i), 1) = "," Then
                            strSort = strSort & "," & Mid(Split(strVscOrder, "|")(i), 2) & " Desc"
                        Else
                            strSort = strSort & "," & Split(strVscOrder, "|")(i)
                        End If
                    Next
                    If strSort <> "" Then rsVsc.Sort = Mid(strSort, 2)
                    rsVsc.MoveFirst
                End If
                If Not rsHsc Is Nothing Then
                    strSort = ""
                    For i = 0 To UBound(Split(strHscOrder, "|"))
                        If Left(Split(strHscOrder, "|")(i), 1) = "," Then
                            strSort = strSort & "," & Mid(Split(strHscOrder, "|")(i), 2) & " Desc"
                        Else
                            strSort = strSort & "," & Split(strHscOrder, "|")(i)
                        End If
                    Next
                    If strSort <> "" Then rsHsc.Sort = Mid(strSort, 2)
                    rsHsc.MoveFirst
                End If
                
                '1.5:��д�������ͷ��Ԫ����
                '�����ͷ
                If Not rsVsc Is Nothing And lngGrid = 0 Then
                    Set colVsc = New Collection
                    '���ܺ���
                    strVscStat = ""
                    For Each tmpID In .SubIDs
                        Set tmpItem = mobjReport.Items("_" & tmpID.id)
                        If tmpItem.���� = 7 Then strVscStat = strVscStat & "," & tmpItem.����
                    Next
                    strVscStat = Mid(strVscStat, 2)
                    
                    '������ͷ
                    k = Y  '��ǰӦ�ô������
                    ReDim arrLevel(X - 1) '���ж�ĳ�������Ƿ�Ӧ�ü����������
                    ReDim arrMerge(X - 1) '���ڴ���ͬ�����ܲ�ͬ�ϼ���ֹ�ϲ�
                    For i = 1 To X - 1
                        arrMerge(i) = Space(i Mod 2)
                    Next
                    For i = 1 To rsVsc.RecordCount
                        k = k + 1
                        If mshBody.Rows - 1 < k Then mshBody.Rows = mshBody.Rows + 1
                        strKey = ""
                        For j = 0 To X - 1
                            strTmp = Trim(mshBody.TextMatrix(Y, j))
                            Select Case rsVsc.Fields(strTmp).type
                                Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
                                    If rsVsc.Fields(strTmp).Value = "#" Then
                                        strKey = strKey & "^"
                                        mshBody.TextMatrix(k, j) = " " '�ÿո���Ϊ��ǿ�ƺϲ�
                                    Else
                                        strKey = strKey & "^" & Replace(rsVsc.Fields(CStr(Split(strVsc, "|")(j))).Value, " ", "��")
                                        mshBody.TextMatrix(k, j) = rsVsc.Fields(CStr(Split(strVsc, "|")(j))).Value
                                    End If
                                Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                                    If rsVsc.Fields(strTmp).Value = 123456707654321# Then
                                        strKey = strKey & "^"
                                        mshBody.TextMatrix(k, j) = " "
                                    Else
                                        strKey = strKey & "^" & Replace(rsVsc.Fields(CStr(Split(strVsc, "|")(j))).Value, " ", "��")
                                        mshBody.TextMatrix(k, j) = rsVsc.Fields(CStr(Split(strVsc, "|")(j))).Value
                                    End If
                                Case adDBTimeStamp, adDBTime, adDBDate, adDate
                                    If rsVsc.Fields(strTmp).Value = #5/5/3000# Then
                                        strKey = strKey & "^"
                                        mshBody.TextMatrix(k, j) = " "
                                    Else
                                        strKey = strKey & "^" & Replace(rsVsc.Fields(CStr(Split(strVsc, "|")(j))).Value, " ", "��")
                                        mshBody.TextMatrix(k, j) = rsVsc.Fields(CStr(Split(strVsc, "|")(j))).Value
                                    End If
                            End Select
                        Next
                        
                        '���������(�ڵ����ڶ���)
                        For j = X - 1 To 1 Step -1 'һ��Ҫ����
                            strTmp = GetRowText(mshBody, k, j - 1)
                            If strTmp <> arrLevel(j) And k > Y + 1 Then
                                If strVscStat <> "" Then
                                    If Split(strVscStat, ",")(j) <> "" Then
                                        mshBody.AddItem "", k
                                        mshBody.Row = k
                                        For l = 0 To j - 1
                                            mshBody.TextMatrix(k, l) = mshBody.TextMatrix(k - 1, l)
                                        Next
                                        For l = j To X - 1
                                            mshBody.Col = l
                                            mshBody.CellAlignment = 4
                                            'mshBody.TextMatrix(k, L) = Space(j Mod 2) & GetStatText(CStr(Split(strVscStat, ",")(j))) & Space(j Mod 2)
                                            mshBody.TextMatrix(k, l) = Space(j) & GetStatText(CStr(Split(strVscStat, ",")(j))) & Space(j)
                                        Next
                                        mshBody.RowData(k) = j + 1
                                        mshBody.MergeRow(k) = True
                                        
                                        k = k + 1
                                    End If
                                End If
                                arrMerge(j) = IIF(arrMerge(j) = "", " ", "")
                            End If
                        Next
                        
                        'ע�⣺kҪΪ���������֮�����(��Ϊ�����в��ǲ��뵽�����)
                        colVsc.Add k, "_" & Mid(strKey, 2) '���������ж�λ����
                        
                        '��ʱKΪ�ǻ�����(���һ��)
                        For j = 1 To X - 1
                            mshBody.TextMatrix(k, j) = mshBody.TextMatrix(k, j) & arrMerge(j)
                            arrLevel(j) = GetRowText(mshBody, k, j - 1)
                        Next
                        
                        rsVsc.MoveNext
                    Next
                    
                    '�������Ļ�����
                    k = mshBody.Rows
                    If strVscStat <> "" And k > Y + 1 Then
                        For j = X - 1 To 0 Step -1
                            If Split(strVscStat, ",")(j) <> "" Then
                                mshBody.AddItem "", k
                                mshBody.Row = k
                                For l = 0 To j - 1
                                    mshBody.TextMatrix(k, l) = mshBody.TextMatrix(k - 1, l)
                                Next
                                For l = j To X - 1
                                    mshBody.Col = l
                                    mshBody.CellAlignment = 4
                                    '���ڻ����У�0,2�кϼ�,1�в��ϼ�,��ʱ0,2�ĺϼư���һ����,��ɺ�����ͬʱ�ϲ���
                                    '��������Ϊ��ʾ��ʽ�е㲻ͬ�����Բ�����������⡣
                                    'mshBody.TextMatrix(k, L) = Space(j Mod 2) & GetStatText(CStr(Split(strVscStat, ",")(j))) & Space(j Mod 2)
                                    mshBody.TextMatrix(k, l) = Space(j) & GetStatText(CStr(Split(strVscStat, ",")(j))) & Space(j)
                                Next
                                mshBody.RowData(k) = j + 1
                                mshBody.MergeRow(k) = True
                                
                                k = k + 1
                            End If
                        Next
                    End If
                End If
                
                '�����ͷ
                If Y > 0 And Not rsHsc Is Nothing Then
                    Set colHsc = New Collection
                    '���ܺ���
                    strHscStat = ""
                    For Each tmpID In .SubIDs
                        Set tmpItem = mobjReport.Items("_" & tmpID.id)
                        If tmpItem.���� = 8 Then strHscStat = strHscStat & "," & tmpItem.����
                    Next
                    strHscStat = Mid(strHscStat, 2)
                    
                    '������ͷ
                    ReDim arrLevel(Y - 1) '���ж�ĳ�������Ƿ�Ӧ�ü����������
                    ReDim arrMerge(Y - 1) '���ڴ���ͬ�����ܲ�ͬ�ϼ���ֹ�ϲ�
                    For i = 1 To Y - 1
                        arrMerge(i) = Space(i Mod 2)
                    Next
                    l = lngCurCols + IIF(lngGrid = 0, X, 0) - Z    '��ǰӦ�ô������
                    For i = 1 To rsHsc.RecordCount
                        l = l + Z
                        If mshBody.Cols - 1 < l Then mshBody.Cols = mshBody.Cols + Z
                        strKey = "" '֮���������ﴦ��(�������ͷ��ͬ),��Ϊ��ȡ��¼ԭʼֵ
                        For j = 0 To Y - 1
                            For k = 0 To Z - 1
                                Select Case rsHsc.Fields(CStr(Split(strHsc, "|")(j))).type
                                    Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
                                        If rsHsc.Fields(CStr(Split(strHsc, "|")(j))).Value = "#" Then
                                            If k = 0 Then strKey = strKey & "^"
                                            mshBody.TextMatrix(j, l + k) = " " '�ÿո���Ϊ��ǿ�ƺϲ�
                                        Else
                                            If k = 0 Then strKey = strKey & "^" & Replace(rsHsc.Fields(CStr(Split(strHsc, "|")(j))).Value, " ", "��")
                                            mshBody.TextMatrix(j, l + k) = Space(j Mod 2) & rsHsc.Fields(CStr(Split(strHsc, "|")(j))).Value & Space(j Mod 2)
                                        End If
                                    Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                                        If rsHsc.Fields(CStr(Split(strHsc, "|")(j))).Value = 123456707654321# Then
                                            If k = 0 Then strKey = strKey & "^"
                                            mshBody.TextMatrix(j, l + k) = " "
                                        Else
                                            If k = 0 Then strKey = strKey & "^" & Replace(rsHsc.Fields(CStr(Split(strHsc, "|")(j))).Value, " ", "��")
                                            mshBody.TextMatrix(j, l + k) = Space(j Mod 2) & rsHsc.Fields(CStr(Split(strHsc, "|")(j))).Value & Space(j Mod 2)
                                        End If
                                    Case adDBTimeStamp, adDBTime, adDBDate, adDate
                                        If rsHsc.Fields(CStr(Split(strHsc, "|")(j))).Value = #5/5/3000# Then
                                            If k = 0 Then strKey = strKey & "^"
                                            mshBody.TextMatrix(j, l + k) = " "
                                        Else
                                            If k = 0 Then strKey = strKey & "^" & Replace(rsHsc.Fields(CStr(Split(strHsc, "|")(j))).Value, " ", "��")
                                            mshBody.TextMatrix(j, l + k) = Space(j Mod 2) & rsHsc.Fields(CStr(Split(strHsc, "|")(j))).Value & Space(j Mod 2)
                                        End If
                                End Select
                            Next
                        Next
                        
                        '���������
                        For j = Y - 1 To 1 Step -1 'һ��Ҫ����
                            strTmp = GetColText(mshBody, j - 1, l)
                            If strTmp <> arrLevel(j) And l > lngCurCols + IIF(lngGrid = 0, X, 0) Then
                                If strHscStat <> "" Then
                                    If Split(strHscStat, ",")(j) <> "" Then
                                        AddCol mshBody, l, Z
                                        For k = 0 To Z - 1
                                            For M = 0 To j - 1
                                                mshBody.TextMatrix(M, l + k) = mshBody.TextMatrix(M, l + k - Z)
                                            Next
                                            mshBody.Col = l + k
                                            mshBody.Row = j
                                            mshBody.CellAlignment = 4
                                            mshBody.TextMatrix(j, l + k) = Space((j + 1) Mod 2) & GetStatText(CStr(Split(strHscStat, ",")(j))) & Space((j + 1) Mod 2)
                                            mshBody.ColData(l + k) = j + 1
                                            mshBody.MergeCol(l + k) = True
                                        Next
                                        l = l + Z
                                    End If
                                End If
                                arrMerge(j) = IIF(arrMerge(j) = "", " ", "")
                            End If
                        Next
                        
                        'ע�⣺LҪΪ���������֮�����(��Ϊ�����в��ǲ��뵽�����)
                        colHsc.Add l, "_" & Mid(strKey, 2) '���������ж�λ����
                        
                        '��ʱLΪ�ǻ�����(���һ����)
                        For j = 1 To Y - 1
                            For k = 0 To Z - 1
                                mshBody.TextMatrix(j, l + k) = mshBody.TextMatrix(j, l + k) & arrMerge(j)
                            Next
                            arrLevel(j) = GetColText(mshBody, j - 1, l)
                        Next
                        rsHsc.MoveNext
                    Next
                    '�������Ļ�����
                    l = mshBody.Cols
                    If strHscStat <> "" And l > lngCurCols + IIF(lngGrid = 0, X, 0) Then
                        For j = Y - 1 To 0 Step -1
                            If Split(strHscStat, ",")(j) <> "" Then
                                AddCol mshBody, l, Z
                                For k = 0 To Z - 1
                                    For M = 0 To j - 1
                                        mshBody.TextMatrix(M, l + k) = mshBody.TextMatrix(M, l + k - Z)
                                    Next
                                    mshBody.Col = l + k
                                    mshBody.Row = j
                                    mshBody.CellAlignment = 4
                                    mshBody.TextMatrix(j, l + k) = Space((j + 1) Mod 2) & GetStatText(CStr(Split(strHscStat, ",")(j))) & Space((j + 1) Mod 2)
                                    mshBody.ColData(l + k) = j + 1
                                    mshBody.MergeCol(l + k) = True
                                Next
                                l = l + Z
                            End If
                        Next
                    End If
                End If
                
                '��дͳ�����ͷ
                strFormat = ""
                For Each tmpID In .SubIDs
                    Set tmpItem = mobjReport.Items("_" & tmpID.id)
                    If tmpItem.���� = 9 Then
                        strFormat = strFormat & "|~" & tmpItem.��� & "~"
                        For i = lngCurCols + IIF(lngGrid = 0, X, 0) To mshBody.Cols - 1 Step Z
                            '�����ǰ�������α�ǰһ�������,��Բ��ӵĹ̶��в���������
                            For j = Y To mshBody.FixedRows - 1
                                mshBody.TextMatrix(j, tmpItem.��� + i) = Space((tmpItem.��� + i) Mod 2) & tmpItem.���� & Space((tmpItem.��� + i) Mod 2)
                            Next
                            '�ϼ�������
                            If mshBody.ColData(i) > 0 Then
                                For M = Y - 1 To mshBody.ColData(i) Step -1
                                    For k = 0 To Z - 1
                                        mshBody.TextMatrix(M, i + k) = mshBody.TextMatrix(M + 1, i + k)
                                    Next
                                Next
                            End If
                        Next
                    End If
                Next
                strFormat = Mid(strFormat, 2)
                
                'ͳ�����ֶθ�ʽ
                For Each tmpID In .SubIDs
                    Set tmpItem = mobjReport.Items("_" & tmpID.id)
                    '����ַ�����Ϊ��ʽ�ַ�
                    If tmpItem.���� = 9 Then strFormat = Replace(strFormat, "~" & tmpItem.��� & "~", tmpItem.��ʽ)
                Next
                
                '�п�(��������)������
                strAlign = ""
                Set colTmp = New Collection
                For Each tmpID In .SubIDs
                    Set tmpItem = mobjReport.Items("_" & tmpID.id)
                    Select Case tmpItem.����
                        Case 7 '�������
                           If lngGrid = 0 Then mshBody.ColWidth(tmpItem.���) = tmpItem.W
                        Case 9 'ͳ����
                            strAlign = strAlign & "," & tmpItem.����
                            For i = lngCurCols + IIF(lngGrid = 0, X, 0) To mshBody.Cols - 1 Step Z
                                mshBody.ColAlignment(i + tmpItem.���) = Switch(tmpItem.���� = 0, 1, tmpItem.���� = 1, 4, tmpItem.���� = 2, 7)
                                If mshBody.FixedRows - 1 >= 0 And mshBody.Rows - 1 >= 0 Then mshBody.Cell(flexcpAlignment, mshBody.FixedRows - 1, i + tmpItem.���, mshBody.Rows - 1, i + tmpItem.���) = mshBody.ColAlignment(i + tmpItem.���)
                                mshBody.ColWidth(i + tmpItem.���) = tmpItem.W
                            Next
                            
                            '���潻��ͳ�Ƶ��ж���
                            colTmp.Add tmpItem, "_" & tmpItem.���
                    End Select
                Next
                strAlign = Mid(strAlign, 2)
                
                '�����������
                rsGroup.MoveFirst
                For i = 1 To rsGroup.RecordCount
                    'ȡ��
                    strKey = ""
                    For j = 0 To UBound(Split(strVsc, "|"))
                        strKey = strKey & "^" & IIF(IsNull(rsGroup.Fields(CStr(Split(strVsc, "|")(j))).Value), "", Replace(Nvl(rsGroup.Fields(CStr(Split(strVsc, "|")(j))).Value, ""), " ", "��"))
                    Next
                    
                    '���������ʱ��,����ұ߱�����ݱ���߶�,��Ϊ����Ϊ׼,������ܶ�λ��,�򲻴������ݡ�
                    lngRow = 0
                    If lngGrid > 0 Then On Local Error Resume Next
                    lngRow = CLng(colVsc("_" & Mid(strKey, 2))) + lngDiff
                    On Error GoTo 0
                    If lngRow > 0 Then
                        'ȡ��
                        lngCol = lngCurCols + IIF(lngGrid = 0, X, 0)
                        If strHsc <> "" Then
                            strKey = ""
                            For j = 0 To UBound(Split(strHsc, "|"))
                                strKey = strKey & "^" & IIF(IsNull(rsGroup.Fields(CStr(Split(strHsc, "|")(j))).Value), "", Replace(rsGroup.Fields(CStr(Split(strHsc, "|")(j))).Value & "", " ", "��"))
                            Next
                            lngCol = CLng(colHsc("_" & Mid(strKey, 2)))
                        End If
                        
                        '����(�ݲ������������)
                        For j = 0 To Z - 1
                            strTmp = Trim(mshBody.TextMatrix(Y, lngCurCols + IIF(lngGrid = 0, X, 0) + j))
                            If Not IsNull(rsGroup.Fields(strTmp).Value) Then
                                '����ʱ��ʽ��
                                StrFmt = ""
                                If strFormat <> "" Then StrFmt = CStr(Split(strFormat, "|")(j))
                                If StrFmt <> "" Then
                                    On Local Error Resume Next
                                    Select Case rsGroup.Fields(strTmp).type
                                        Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                                            mshBody.TextMatrix(lngRow, lngCol + j) = Format(Val(Replace(mshBody.TextMatrix(lngRow, lngCol + j), ",", "")) + rsGroup.Fields(strTmp).Value, StrFmt)
                                        Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
                                            mshBody.TextMatrix(lngRow, lngCol + j) = Format(Val(mshBody.TextMatrix(lngRow, lngCol + j)) + Val(rsGroup.Fields(strTmp).Value), StrFmt)
                                        Case adDBTimeStamp, adDBTime, adDBDate, adDate
                                            If mshBody.TextMatrix(lngRow, lngCol + j) = "" Then
                                                mshBody.TextMatrix(lngRow, lngCol + j) = Format(CDate(rsGroup.Fields(strTmp).Value), StrFmt)
                                            Else
                                                mshBody.TextMatrix(lngRow, lngCol + j) = Format(CDate(mshBody.TextMatrix(lngRow, lngCol + j)) + rsGroup.Fields(strTmp).Value, StrFmt)
                                            End If
                                    End Select
                                    On Local Error GoTo 0
                                Else
                                    Select Case rsGroup.Fields(strTmp).type
                                        Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                                            mshBody.TextMatrix(lngRow, lngCol + j) = Val(Replace(mshBody.TextMatrix(lngRow, lngCol + j), ",", "")) + rsGroup.Fields(strTmp).Value
                                        Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
                                            mshBody.TextMatrix(lngRow, lngCol + j) = Val(mshBody.TextMatrix(lngRow, lngCol + j)) + Val(rsGroup.Fields(strTmp).Value)
                                        Case adDBTimeStamp, adDBTime, adDBDate, adDate
                                            If mshBody.TextMatrix(lngRow, lngCol + j) = "" Then
                                                mshBody.TextMatrix(lngRow, lngCol + j) = CDate(rsGroup.Fields(strTmp).Value)
                                            Else
                                                mshBody.TextMatrix(lngRow, lngCol + j) = CDate(mshBody.TextMatrix(lngRow, lngCol + j)) + rsGroup.Fields(strTmp).Value
                                            End If
                                    End Select
                                End If
                                
                                '���ݶ������÷�ֹ�ϲ�
                                Select Case CByte(Split(strAlign, ",")(j))
                                    Case 0 '�����
                                        mshBody.TextMatrix(lngRow, lngCol + j) = mshBody.TextMatrix(lngRow, lngCol + j) & Space((lngRow + lngCol + j) Mod 2)
                                    Case 1 '�ж���
                                        mshBody.TextMatrix(lngRow, lngCol + j) = Space((lngRow + lngCol + j) Mod 2) & mshBody.TextMatrix(lngRow, lngCol + j) & Space((lngRow + lngCol + j) Mod 2)
                                    Case 2 '�Ҷ���
                                        mshBody.TextMatrix(lngRow, lngCol + j) = Space((lngRow + lngCol + j) Mod 2) & mshBody.TextMatrix(lngRow, lngCol + j)
                                End Select
                                
                                '�����嵱ǰ�ж���
                                Set objStatusGridItem = Nothing
                                For Each objStatusGridItem In colTmp
                                    If objStatusGridItem.��� = j And objStatusGridItem.���� = Val("9-����������") Then
                                        Exit For
                                    End If
                                Next
                                
                                If Not objStatusGridItem Is Nothing Then
                                    For k = 1 To objStatusGridItem.ColProtertys.count
                                        Set objColProp = objStatusGridItem.ColProtertys.Item(k)
                                        If InStr(objColProp.����ֵ, objCurItem.���� & ".") > 0 Then
                                            varIFValue = GetStatGridData(mshBody.Index, objColProp.����ֵ, lngRow, lngCol + j)
                                        Else
                                            varIFValue = objColProp.����ֵ
                                        End If
                                        If lngCol + j = mshBody.FixedCols And objColProp.�Ƿ�����Ӧ�� Then
                                            If CheckColProtertys(Trim(mshBody.TextMatrix(lngRow, lngCol + j)), objColProp.������ϵ, varIFValue) Then
                                                If objColProp.������ɫ <> vbWhite Then
                                                    mshBody.Cell(flexcpBackColor, lngRow, mshBody.FixedCols, lngRow, mshBody.Cols - 1) = objColProp.������ɫ
                                                End If
                                                If objColProp.������ɫ <> vbBlack Then
                                                    mshBody.Cell(flexcpForeColor, lngRow, mshBody.FixedCols, lngRow, mshBody.Cols - 1) = objColProp.������ɫ
                                                End If
                                                If objColProp.�Ƿ�Ӵ� Then
                                                    mshBody.Cell(flexcpFontBold, lngRow, mshBody.FixedCols, lngRow, mshBody.Cols - 1) = objColProp.�Ƿ�Ӵ�
                                                End If
                                            End If
                                        Else
                                            If CheckColProtertys(Trim(mshBody.TextMatrix(lngRow, lngCol + j)), objColProp.������ϵ, varIFValue) Then
                                                If objColProp.������ɫ <> vbWhite Then
                                                    mshBody.Cell(flexcpBackColor, lngRow, lngCol + j) = objColProp.������ɫ
                                                End If
                                                If objColProp.������ɫ <> vbBlack Then
                                                    mshBody.Cell(flexcpForeColor, lngRow, lngCol + j) = objColProp.������ɫ
                                                End If
                                                If objColProp.�Ƿ�Ӵ� Then
                                                    mshBody.Cell(flexcpFontBold, lngRow, lngCol + j) = objColProp.�Ƿ�Ӵ�
                                                End If
                                                
                                                '���뷽ʽ
                                                Select Case objColProp.����
                                                Case Val("1-����")
                                                    mshBody.Cell(flexcpAlignment, lngRow, lngCol + j) = flexAlignLeftCenter
                                                Case Val("2-����")
                                                    mshBody.Cell(flexcpAlignment, lngRow, lngCol + j) = flexAlignCenterCenter
                                                Case Val("3-����")
                                                    mshBody.Cell(flexcpAlignment, lngRow, lngCol + j) = flexAlignRightCenter
                                                Case Else
                                                    'ȱʡ��������
                                                End Select
                                            End If
                                        End If
                                    Next
                                End If
                                
                            End If
                        Next
                    End If
                    rsGroup.MoveNext
                Next
                Set colTmp = Nothing
                
                '���������������(��������)
                '���������
                If strHsc <> "" And strHscStat <> "" Then
                    For l = UBound(Split(strHsc, "|")) To 0 Step -1
                        strStat = CStr(Split(strHscStat, ",")(l))
                        If strStat <> "" Then
                            ReDim arrStat(mshBody.FixedRows To mshBody.Rows - 1, Z - 1)  '�����������
                            ReDim arrCount(mshBody.FixedRows To mshBody.Rows - 1, Z - 1) '����ǿռ�¼����
                            blnDo = False
                            For j = lngCurCols + IIF(lngGrid = 0, X, 0) To mshBody.Cols - 1 Step Z
                                For i = mshBody.FixedRows To mshBody.Rows - 1 '��Ϊ���ܶ�����������,Y��׼,��FixedRows
                                    '��ʾ�����н��
                                    If mshBody.ColData(j) = l + 1 Then
                                        For k = 0 To Z - 1
                                            If strStat = "AVG" Then
                                                strTmp = Trim(mshBody.TextMatrix(Y, lngCurCols + IIF(lngGrid = 0, X, 0) + k))
                                                Select Case rsGroup.Fields(strTmp).type
                                                    Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                                                        arrStat(i, k) = Val(arrStat(i, k) / arrCount(i, k))
                                                    Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
                                                        arrStat(i, k) = Val(arrStat(i, k) / arrCount(i, k))
                                                    Case adDBTimeStamp, adDBTime, adDBDate, adDate
                                                        arrStat(i, k) = CDate(arrStat(i, k) / arrCount(i, k))
                                                End Select
                                            End If
                                            StrFmt = ""
                                            If strFormat <> "" Then StrFmt = CStr(Split(strFormat, "|")(k))
                                            If StrFmt <> "" Then
                                                On Local Error Resume Next
                                                Select Case Split(strAlign, ",")(k)
                                                    Case 0 '��
                                                        mshBody.TextMatrix(i, j + k) = Format(arrStat(i, k), StrFmt) & Space((i + j + k) Mod 2)
                                                    Case 1 '��
                                                        mshBody.TextMatrix(i, j + k) = Space((i + j + k) Mod 2) & Format(arrStat(i, k), StrFmt) & Space((i + j + k) Mod 2)
                                                    Case 2 '��
                                                        mshBody.TextMatrix(i, j + k) = Space((i + j + k) Mod 2) & Format(arrStat(i, k), StrFmt)
                                                End Select
                                                On Local Error GoTo 0
                                            Else
                                                Select Case Split(strAlign, ",")(k)
                                                    Case 0 '��
                                                        mshBody.TextMatrix(i, j + k) = arrStat(i, k) & Space((i + j + k) Mod 2)
                                                    Case 1 '��
                                                        mshBody.TextMatrix(i, j + k) = Space((i + j + k) Mod 2) & arrStat(i, k) & Space((i + j + k) Mod 2)
                                                    Case 2 '��
                                                        mshBody.TextMatrix(i, j + k) = Space((i + j + k) Mod 2) & arrStat(i, k)
                                                End Select
                                            End If
                                        Next
                                    '�����������
                                    ElseIf mshBody.ColData(j) = 0 Then
                                        For k = 0 To Z - 1
                                            If Trim(mshBody.TextMatrix(i, j + k)) <> "" Then
                                                strTmp = Trim(mshBody.TextMatrix(Y, lngCurCols + IIF(lngGrid = 0, X, 0) + k))
                                                arrCount(i, k) = arrCount(i, k) + 1
                                                Select Case strStat
                                                    Case "SUM", "AVG"
                                                        Select Case rsGroup.Fields(strTmp).type
                                                            Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                                                                arrStat(i, k) = arrStat(i, k) + Val(Replace(Trim(mshBody.TextMatrix(i, j + k)), ",", ""))
                                                            Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
                                                                arrStat(i, k) = arrStat(i, k) + Val(Trim(mshBody.TextMatrix(i, j + k)))
                                                            Case adDBTimeStamp, adDBTime, adDBDate, adDate
                                                                arrStat(i, k) = arrStat(i, k) + CDate(Trim(mshBody.TextMatrix(i, j + k)))
                                                        End Select
                                                    Case "MIN"
                                                        If Not blnDo Then arrStat(i, k) = Trim(mshBody.TextMatrix(i, j + k)): blnDo = True
                                                        If Trim(mshBody.TextMatrix(i, j + k)) < arrStat(i, k) Then arrStat(i, k) = Trim(mshBody.TextMatrix(i, j + k))
                                                    Case "MAX"
                                                        If Not blnDo Then arrStat(i, k) = Trim(mshBody.TextMatrix(i, j + k)): blnDo = True
                                                        If Trim(mshBody.TextMatrix(i, j + k)) > arrStat(i, k) Then arrStat(i, k) = Trim(mshBody.TextMatrix(i, j + k))
                                                    Case "COUNT"
                                                        arrStat(i, k) = arrStat(i, k) + 1
                                                End Select
                                            End If
                                        Next
                                    End If
                                Next
                                If mshBody.ColData(j) = l + 1 Then
                                    ReDim arrStat(mshBody.FixedRows To mshBody.Rows - 1, Z - 1)  '�����������
                                    ReDim arrCount(mshBody.FixedRows To mshBody.Rows - 1, Z - 1) '����ǿռ�¼����
                                    blnDo = False
                                End If
                            Next
                        End If
                    Next
                End If

                '���������
                If strVscStat <> "" Then
                    For l = UBound(Split(strVsc, "|")) To 0 Step -1
                        strStat = CStr(Split(strVscStat, ",")(l))
                        If strStat <> "" Then
                            ReDim arrStat(lngCurCols + IIF(lngGrid = 0, X, 0) To mshBody.Cols - 1) '�����������
                            ReDim arrCount(lngCurCols + IIF(lngGrid = 0, X, 0) To mshBody.Cols - 1) '����ǿռ�¼����
                            blnDo = False
                            For i = mshBody.FixedRows To mshBody.Rows - 1 '��Ϊ���ܶ�����������,Y��׼,��FixedRows
                                For j = lngCurCols + IIF(lngGrid = 0, X, 0) To mshBody.Cols - 1
                                    '��ʾ�����н��
                                    If mshBody.RowData(i) = l + 1 Then
                                        If strStat = "AVG" Then
                                            strTmp = Trim(mshBody.TextMatrix(Y, j))
                                            Select Case rsGroup.Fields(strTmp).type
                                                Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                                                    arrStat(j) = Val(arrStat(j) / arrCount(j))
                                                Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
                                                    arrStat(j) = Val(arrStat(j) / arrCount(j))
                                                Case adDBTimeStamp, adDBTime, adDBDate, adDate
                                                    arrStat(j) = CDate(arrStat(j) / arrCount(j))
                                            End Select
                                        End If
                                        k = 0
                                        If Z > 1 Then k = ((j - (lngCurCols + IIF(lngGrid = 0, X, 0)) + 1) Mod Z) - 1
                                        If k = -1 Then k = Z - 1
                                        StrFmt = ""
                                        If strFormat <> "" Then StrFmt = CStr(Split(strFormat, "|")(k))
                                        If StrFmt <> "" Then
                                            On Local Error Resume Next
                                            Select Case Split(strAlign, ",")(k)
                                                Case 0 '��
                                                    mshBody.TextMatrix(i, j) = Format(arrStat(j), StrFmt) & Space((i + j) Mod 2)
                                                Case 1 '��
                                                    mshBody.TextMatrix(i, j) = Space((i + j) Mod 2) & Format(arrStat(j), StrFmt) & Space((i + j) Mod 2)
                                                Case 2 '��
                                                    mshBody.TextMatrix(i, j) = Space((i + j) Mod 2) & Format(arrStat(j), StrFmt)
                                            End Select
                                            On Local Error GoTo 0
                                        Else
                                            Select Case Split(strAlign, ",")(k)
                                                Case 0 '��
                                                    mshBody.TextMatrix(i, j) = arrStat(j) & Space((i + j) Mod 2)
                                                Case 1 '��
                                                    mshBody.TextMatrix(i, j) = Space((i + j) Mod 2) & arrStat(j) & Space((i + j) Mod 2)
                                                Case 2 '��
                                                    mshBody.TextMatrix(i, j) = Space((i + j) Mod 2) & arrStat(j)
                                            End Select
                                        End If
                                    '�����������
                                    ElseIf mshBody.RowData(i) = 0 And Trim(mshBody.TextMatrix(i, j)) <> "" Then
                                        strTmp = Trim(mshBody.TextMatrix(Y, j))
                                        arrCount(j) = arrCount(j) + 1
                                        Select Case strStat
                                            Case "SUM", "AVG"
                                                Select Case rsGroup.Fields(strTmp).type
                                                    Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                                                        arrStat(j) = arrStat(j) + Val(Replace(Trim(mshBody.TextMatrix(i, j)), ",", ""))
                                                    Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
                                                        arrStat(j) = arrStat(j) + Val(Trim(mshBody.TextMatrix(i, j)))
                                                    Case adDBTimeStamp, adDBTime, adDBDate, adDate
                                                        arrStat(j) = arrStat(j) + CDate(Trim(mshBody.TextMatrix(i, j)))
                                                End Select
                                            Case "MIN"
                                                If Not blnDo Then arrStat(j) = Trim(mshBody.TextMatrix(i, j)): blnDo = True
                                                If Trim(mshBody.TextMatrix(i, j)) < arrStat(j) Then arrStat(j) = Trim(mshBody.TextMatrix(i, j))
                                            Case "MAX"
                                                If Not blnDo Then arrStat(j) = Trim(mshBody.TextMatrix(i, j)): blnDo = True
                                                If Trim(mshBody.TextMatrix(i, j)) > arrStat(j) Then arrStat(j) = Trim(mshBody.TextMatrix(i, j))
                                            Case "COUNT"
                                                arrStat(j) = arrStat(j) + 1
                                        End Select
                                    End If
                                Next
                                If mshBody.RowData(i) = l + 1 Then
                                    ReDim arrStat(lngCurCols + IIF(lngGrid = 0, X, 0) To mshBody.Cols - 1)
                                    ReDim arrCount(lngCurCols + IIF(lngGrid = 0, X, 0) To mshBody.Cols - 1)
                                    blnDo = False
                                End If
                            Next
                        End If
                    Next
                End If
                For Each tmpID In .SubIDs
                    Set tmpItem = mobjReport.Items("_" & tmpID.id)
                    Select Case tmpItem.����
                        Case 7 '�������
                            '���������ѯ����������
                            If tmpItem.Relations.count > 0 Then
                                mshBody.Cell(flexcpForeColor, mshBody.FixedRows, tmpItem.���, mshBody.Rows - 1, tmpItem.���) = &HFF0001
                                mshBody.Cell(flexcpFontUnderline, mshBody.FixedRows, tmpItem.���, mshBody.Rows - 1, tmpItem.���) = True
                                mshBody.Cell(flexcpData, mshBody.FixedRows, tmpItem.���, mshBody.Rows - 1, tmpItem.���) = tmpItem
                            End If
                            '���ñ�񲿷�������ɫ�ͼӴ֣����ȣ�
                            mshBody.Cell(flexcpFontBold, mshBody.FixedRows, tmpItem.���, mshBody.Rows - 1, tmpItem.���) = tmpItem.����
                            If tmpItem.ǰ�� <> 0 Then mshBody.Cell(flexcpForeColor, mshBody.FixedRows, tmpItem.���, mshBody.Rows - 1, tmpItem.���) = tmpItem.ǰ��
                        Case 8 '�������
                            If tmpItem.Relations.count > 0 Then
                                mshBody.Cell(flexcpForeColor, tmpItem.���, mshBody.FixedCols, tmpItem.���, mshBody.Cols - 1) = &HFF0001
                                mshBody.Cell(flexcpFontUnderline, tmpItem.���, mshBody.FixedCols, tmpItem.���, mshBody.Cols - 1) = True
                                mshBody.Cell(flexcpData, tmpItem.���, mshBody.FixedCols, tmpItem.���, mshBody.Cols - 1) = tmpItem
                            End If
                            '���ñ�񲿷�������ɫ�ͼӴ֣����ȣ�
                            mshBody.Cell(flexcpFontBold, tmpItem.���, mshBody.FixedCols, tmpItem.���, mshBody.Cols - 1) = tmpItem.����
                            If tmpItem.ǰ�� <> 0 Then mshBody.Cell(flexcpForeColor, tmpItem.���, mshBody.FixedCols, tmpItem.���, mshBody.Cols - 1) = tmpItem.ǰ��
                        Case 9 'ͳ����
                            For j = mshBody.FixedCols To mshBody.Cols - 1 Step lngStatistics
                                On Error Resume Next
                                If tmpItem.Relations.count > 0 Then
                                    mshBody.Cell(flexcpForeColor, mshBody.FixedRows, tmpItem.��� + j, mshBody.Rows - 1, tmpItem.��� + j) = &HFF0001
                                    mshBody.Cell(flexcpFontUnderline, mshBody.FixedRows, tmpItem.��� + j, mshBody.Rows - 1, tmpItem.��� + j) = True
                                End If
                                '�Ż���ֻ���ض��а����������
                                mshBody.Cell(flexcpData, mshBody.FixedRows, tmpItem.��� + j, mshBody.FixedRows, tmpItem.��� + j) = tmpItem

'                                 '���ñ�񲿷�������ɫ�ͼӴ֣����ȣ�
'                                mshBody.Cell(flexcpFontBold, mshBody.FixedRows, tmpItem.��� + j, mshBody.Rows - 1, tmpItem.��� + j) = tmpItem.����
'                                If tmpItem.ǰ�� <> 0 Then mshBody.Cell(flexcpForeColor, mshBody.FixedRows, tmpItem.��� + j, mshBody.Rows - 1, tmpItem.��� + j) = tmpItem.ǰ��
                                On Error GoTo 0
                            Next
                    End Select
                Next
                'ȥ�������еĳ�����
                For i = 0 To mshBody.FixedCols - 1
                    For j = 0 To mshBody.Rows - 1
                        If Decode(Trim(mshBody.TextMatrix(j, i)), "�ϼ�", 1, "ƽ��ֵ", 2, "���ֵ", 3, "��Сֵ", 4, "��¼��", 5, 0) > 0 Then
                            mshBody.Cell(flexcpForeColor, j, i, j, mshBody.Cols - 1) = mshBody.ForeColor
                            mshBody.Cell(flexcpFontUnderline, j, i, j, mshBody.Cols - 1) = False
                            mshBody.Cell(flexcpData, j, i, j, mshBody.Cols - 1) = Empty
                            mshBody.Cell(flexcpFontBold, j, i, j, mshBody.Cols - 1) = False
                        End If
                    Next
                Next
                For j = 0 To mshBody.FixedRows - 1
                    For i = 0 To mshBody.Cols - 1
                        If Decode(Trim(mshBody.TextMatrix(j, i)), "�ϼ�", 1, "ƽ��ֵ", 2, "���ֵ", 3, "��Сֵ", 4, "��¼��", 5, 0) > 0 Then
                            mshBody.Cell(flexcpForeColor, j, i, mshBody.Rows - 1, i) = mshBody.ForeColor
                            mshBody.Cell(flexcpFontUnderline, j, i, mshBody.Rows - 1, i) = False
                            mshBody.Cell(flexcpData, j, i, mshBody.Rows - 1, i) = Empty
                            mshBody.Cell(flexcpFontBold, j, i, mshBody.Rows - 1, i) = False
                        End If
                    Next
                Next
                
                '����ͳ������ʽ
                If Z = 1 And Y > 0 Then
                    For i = lngCurCols + IIF(lngGrid = 0, X, 0) To mshBody.Cols - 1
                        For j = mshBody.FixedRows - 1 To Y Step -1
                            mshBody.TextMatrix(j, i) = mshBody.TextMatrix(Y - 1, i)
                        Next
                    Next
                Else
                    blnHide = False
                End If
                
                '��������"��ͷ"���Դ�Ÿ�������������,����"ȱʡ�п�"����
                objItem.��ͷ = objItem.��ͷ & "|" & .id & "," & mshBody.Cols
                strTopRow = strTopRow & "|" & objCurItem.���� & "," & mshBody.Cols
                                
                '��ǰ���������������
                lngCurCols = mshBody.Cols
            Else
                blnHide = False
                Call SetHeadCenter(mshBody)
                Exit For '��������û������,�������ӵı��Ҳ���ô�����
            End If
        End With
    Next
    
    objItem.��ͷ = Mid(objItem.��ͷ, 2)
    
    '���������ʱ,��һ����(�öο�ȡ��,��SQLʵ��)
'    strTopRow = Mid(strTopRow, 2)
'    strTmp = ""
'    For i = 0 To UBound(Split(strTopRow, "|"))
'        If InStr(strTmp & "|", "|" & Split(Split(strTopRow, "|")(i), ",")(0) & "|") = 0 Then
'            strTmp = strTmp & "|" & Split(Split(strTopRow, "|")(i), ",")(0)
'        End If
'    Next
'    If UBound(Split(Mid(strTmp, 2), "|")) > 0 Then
'        '������
'        mshBody.AddItem "", mshBody.FixedRows
'        mshBody.FixedRows = mshBody.FixedRows + 1
'        For i = mshBody.FixedRows - 1 To 1 Step -1
'            For j = 0 To mshBody.Cols - 1
'                mshBody.TextMatrix(i, j) = mshBody.TextMatrix(i - 1, j)
'                mshBody.RowHeight(i) = mshBody.RowHeight(i - 1)
'                mshBody.RowData(i) = mshBody.RowData(i - 1)
'            Next
'        Next
'        mshBody.RowData(0) = 0
'        mshBody.RowHeight(0) = objItem.�и�
'        mshBody.MergeRow(0) = True
'        For j = mshBody.FixedCols To mshBody.Cols - 1
'            mshBody.TextMatrix(0, j) = ""
'        Next
'
'        '��д����
'        For i = 0 To UBound(Split(strTopRow, "|"))
'            If i = 0 Then
'                lngColB = mshBody.FixedCols
'            Else
'                lngColB = lngColE + 1
'            End If
'            lngColE = CLng(Split(Split(strTopRow, "|")(i), ",")(1)) - 1
'            For j = lngColB To lngColE
'                mshBody.TextMatrix(0, j) = CStr(Split(Split(strTopRow, "|")(i), ",")(0))
'            Next
'        Next
'    End If
    
    '�̶����кϲ�
    For j = 0 To mshBody.Cols - 1
        mshBody.MergeCol(j) = True
    Next
    For i = 0 To mshBody.FixedRows - 2
        mshBody.MergeRow(i) = True
    Next
    
    
    '�����ͷ����(��Ԫ���롢�иߡ�����,�п�)
    For Each tmpID In objItem.SubIDs
        Set tmpItem = mobjReport.Items("_" & tmpID.id)
        arrHead = Split(tmpItem.��ͷ, "|")
        For i = 0 To UBound(arrHead) '����^�߶�^����
            mshBody.RowHeight(i) = CLng(Split(arrHead(i), "^")(1))
        Next
    Next
    
    '�и�(��������)
    For i = mshBody.FixedRows To mshBody.Rows - 1
        mshBody.RowHeight(i) = objItem.�и�
    Next
    
    '���ص���ͳ�����ͷ��
    '------�˶οɻ�Ϊ�¶�-----------------
    blnHide = True
    For i = mshBody.FixedRows - 1 To 1 Step -1
        For j = 0 To mshBody.Cols - 1
            If mshBody.TextMatrix(i, j) <> mshBody.TextMatrix(i - 1, j) Then
                blnHide = False: Exit For
            End If
        Next
        If blnHide Then
            mshBody.RowHeight(i) = 0
        Else
            Exit For
        End If
    Next
    '------�˶ο��滻�϶�-----------------
    'If blnHide Then mshBody.RowHeight(mshBody.FixedRows - 1) = 0
    
    '�̶����ж���
    mshBody.Cell(flexcpAlignment, 0, 0, mshBody.FixedRows - 1, mshBody.Cols - 1) = flexAlignCenterCenter
    mshBody.Cell(flexcpAlignment, 0, 0, mshBody.Rows - 1, mshBody.FixedCols - 1) = flexAlignCenterCenter
    
    '�̶��������(�Ǻϼ���)
    For i = mshBody.FixedRows To mshBody.Rows - 1
        If mshBody.RowData(i) = 0 Then
            mshBody.Row = i
            For j = 0 To mshBody.FixedCols - 1
                mshBody.Col = j
                mshBody.CellAlignment = 1
            Next
        End If
    Next
    
    mshBody.WordWrap = True
    
    mshBody.MergeCells = flexMergeFree
    mshBody.ScrollBars = flexScrollBarBoth
    mshBody.Row = mshBody.FixedRows
    mshBody.Col = mshBody.FixedCols
    mshBody.Redraw = flexRDBuffered
    mshBody.ZOrder
    mshBody.Visible = True
    Exit Sub
    
hErr:
    Call ErrCenter
End Sub

Private Function GetGridColWidth(ByVal objGrid As Object _
    , Optional ByRef intPageLastCol As Integer _
    , Optional ByVal lngMaxWidth As Long = 0) As Long
'���ܣ���ȡһ�������п��֮��

    Dim i As Integer
    Dim lngW As Long, lngGridWidth As Long
    
    intPageLastCol = -1
    lngW = 0
    For i = 0 To objGrid.Cols - 1
        If lngMaxWidth > 0 Then
            If lngMaxWidth >= lngW + objGrid.ColWidth(i) Then
                lngW = lngW + objGrid.ColWidth(i)
                intPageLastCol = i
            Else
                Exit For
            End If
        Else
            lngW = lngW + objGrid.ColWidth(i)
        End If
    Next
    GetGridColWidth = lngW
End Function

Private Sub SetGridAlign(Optional ByVal bytMode As Byte = 0)
'���ܣ��Զ��������ӱ�����п�(�������Ϊ׼����)
'˵����ֻ�����ӱ���,�Ұ���Ƴߴ紦��,ֻ�ڱ�����ʾǰ����һ��
    
    Dim tmpMsh As VSFlexGrid, tmpBody As VSFlexGrid
    Dim tmpItem As RPTItem
    Dim lngMaxW As Long, lngCurW As Long, lngWidth As Long, lngSum As Long
    Dim strIDs As String
    Dim i As Integer, j As Integer, intCurID As Integer, intCol As Integer
    Dim arrIDs As Variant
    Dim sngRate As Single

    If Not mobjReport.blnLoad Then Exit Sub

    On Error GoTo hErr

    For Each tmpItem In mobjReport.Items
        If tmpItem.��ʽ�� = bytFormat _
            And (tmpItem.���� = Val("4-���ɱ��") Or tmpItem.���� = Val("5-���ܱ��")) _
            And tmpItem.���� = "" And tmpItem.���� = 0 Then

            '�ж��Ƿ���ڸ��ӱ�
            If GridHaveApp(tmpItem.id) Then
                '���ӱ�1..n��
                strIDs = GetGridAppIDs(tmpItem.����)
                strIDs = tmpItem.id & "," & strIDs
                arrIDs = Split(strIDs, ",")

                '��ȡ�ο������е��ܿ�ȣ���ҳ�������㳬ҳ��������У�
                On Error Resume Next
                Set tmpMsh = Nothing
                Set tmpMsh = msh(Val(arrIDs(0)))
                On Error GoTo hErr

                lngMaxW = -1
                If Not tmpMsh Is Nothing Then
                    If bytMode = Val("1-����") Then
                        GoSub makPro
                    End If
                    lngMaxW = GetGridColWidth(tmpMsh, , tmpItem.W)              '�������ҳ�Ŀ��
                    lngWidth = GetGridColWidth(tmpMsh)                          '�����е��ܿ��
                    
                    '�и��ӱ�����������п�С�ڱ����ƿ�ʱ��������ƿ������������ܱ����п�
                    If tmpItem.W > lngWidth Then
                        'lngMaxW����Ϊ�������ص��ֵ
                        lngMaxW = Round(tmpItem.W / Screen.TwipsPerPixelX) * Screen.TwipsPerPixelX
                        sngRate = lngMaxW / IIF(lngWidth = 0, 1, lngWidth)
                        lngSum = 0
                        For j = 0 To tmpMsh.Cols - 1
                            'ȷ���������ص��ֵ
                            tmpMsh.ColWidth(j) = Round(tmpMsh.ColWidth(j) * sngRate / Screen.TwipsPerPixelX) * Screen.TwipsPerPixelX
                            If lngSum + tmpMsh.ColWidth(j) >= lngMaxW Then
                                '�������ƿ��
                                tmpMsh.ColWidth(j) = lngMaxW - lngSum
                            End If
                            lngSum = lngSum + tmpMsh.ColWidth(j)
                        Next
                        
                        '���ɱ��ı�ͷ
                        If tmpItem.���� = Val("4-���ɱ��") And Not tmpMsh Is Nothing Then
                            Set tmpMsh = msh(Val(tmpMsh.Tag))
                            If Not tmpMsh Is Nothing Then
                                For j = 0 To tmpMsh.Cols - 1
                                    'ȷ���������ص��ֵ
                                    tmpMsh.ColWidth(j) = Round(tmpMsh.ColWidth(j) * sngRate / Screen.TwipsPerPixelX) * Screen.TwipsPerPixelX
                                Next
                            End If
                        End If
                        
                        '�����ܿ��
                        lngMaxW = lngSum
                    End If
                End If

                If lngMaxW > -1 Then
                    '�������ӱ�����һ�еĿ����ο�����ȶ��롣���ӱ�ֻ�������ɱ�񣬲�������ܱ��
                    For i = 1 To UBound(arrIDs)
                        intCurID = Val(arrIDs(i))
                        Set tmpBody = Nothing
                        If mobjReport.Items("_" & intCurID).���� = Val("4-�����") Then
                            Set tmpMsh = msh(CInt(msh(intCurID).Tag))           '��ͷ
                            Set tmpBody = msh(intCurID)                         '����
                            
                            tmpMsh.Redraw = False
                            lngCurW = GetGridColWidth(tmpMsh, intCol, lngMaxW)  '���ӱ���ͷ��ҳ�Ŀ��
                            
                            '���ӱ��ο������
                            If intCol > -1 Then
                                If intCol < tmpMsh.Cols - 1 Then
                                    '��ҳ����ƽ���п���㣩
                                    For j = intCol + 1 To tmpMsh.Cols - 1
                                        tmpMsh.ColWidth(j) = (lngMaxW - lngCurW) \ (tmpMsh.Cols - 1 - intCol)
                                        tmpBody.ColWidth(j) = tmpMsh.ColWidth(j)
                                    Next
                                Else
                                    If lngMaxW >= lngCurW Then
                                        tmpMsh.ColWidth(tmpMsh.Cols - 1) = tmpMsh.ColWidth(tmpMsh.Cols - 1) + lngMaxW - lngCurW
                                        tmpBody.ColWidth(tmpBody.Cols - 1) = tmpMsh.ColWidth(tmpMsh.Cols - 1)
                                    Else
                                        Debug.Print ""
                                    End If
                                End If
                            End If
                            
                            tmpMsh.Redraw = True
                        End If
                    Next
                    Erase arrIDs
                End If
                
                '�������ɱ����и�
                If tmpItem.���� = Val("4-�����") Then
                    Call AdjustRowHight(tmpItem.id)
                End If
                
            ElseIf bytMode = Val("1-����") Then
                Set tmpMsh = msh(CInt(msh(tmpItem.id).Tag))     '��ͷ
                GoSub makPro
                Set tmpMsh = msh(tmpItem.id)                    '����
                GoSub makPro
            End If
        End If
    Next
    
    Exit Sub

hErr:
    Call ErrCenter
    Exit Sub
    
makPro:
    lngWidth = GetGridColWidth(tmpMsh)
    sngRate = (tmpMsh.Width - 300) / IIF(lngWidth = 0, 1, lngWidth)
    For j = 0 To tmpMsh.Cols - 1
        tmpMsh.ColWidth(j) = tmpMsh.ColWidth(j) * sngRate
    Next
    Return
End Sub

Private Function GetGridAppIDs(strName As String) As String
'���ܣ���ȡ���ձ��ĸ��ӱ���������
'������strName=������
    Dim tmpItem As RPTItem
    Dim strIDs As String
    
    For Each tmpItem In mobjReport.Items
        If tmpItem.��ʽ�� = bytFormat And tmpItem.���� = 4 _
            And tmpItem.���� = 1 And tmpItem.���� = strName Then
            strIDs = strIDs & "," & tmpItem.id
        End If
    Next
    GetGridAppIDs = Mid(strIDs, 2)
End Function

Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        lngTXTProc = GetWindowLong(txt(Index).hwnd, GWL_WNDPROC)
        Call SetWindowLong(txt(Index).hwnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txt(Index).hwnd, GWL_WNDPROC, lngTXTProc)
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    If txt(Index).Tag = "" And txt(Index).ToolTipText <> "" Then
        'ǿ������ƥ��
        If txt(Index).Text <> "" Then
            If cmd(Index).Enabled And cmd(Index).Visible Then
                blnMatch = True
                Call cmd_Click(Index)
            End If
            Cancel = True
        End If
    End If
End Sub

Private Sub ReplaceSysNo(objReport As Report)
    Dim i As Integer, j As Integer
    For i = 1 To objReport.Datas.count
        objReport.Datas(i).SQL = Replace(objReport.Datas(i).SQL, "[ϵͳ]", IIF(mlngSys <> 0, mlngSys, objReport.ϵͳ))
        For j = 1 To objReport.Datas(i).Pars.count
            objReport.Datas(i).Pars(j).��ϸSQL = Replace(objReport.Datas(i).Pars(j).��ϸSQL, "[ϵͳ]", IIF(mlngSys <> 0, mlngSys, objReport.ϵͳ))
            objReport.Datas(i).Pars(j).����SQL = Replace(objReport.Datas(i).Pars(j).����SQL, "[ϵͳ]", IIF(mlngSys <> 0, mlngSys, objReport.ϵͳ))
        Next
    Next
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

Private Function GetStatGridData(ByVal intIndex As Integer, ByVal strFiled As String, ByVal lngRow As Long, ByVal lngCol As Long)
'���ܣ�ȡ���ܱ��Ӧ���������ָ���е������ֶε�ֵ
    Dim str������� As String
    Dim i As Long, strFiledTmp As String
    
    With msh(intIndex)
        strFiledTmp = Mid(strFiled, InStr(strFiled, ".") + 1)
        str������� = .TextMatrix(0, lngCol)
        
        For i = lngCol To .Cols - 1
            If str������� = .TextMatrix(0, i) And strFiledTmp = Trim(.TextMatrix(.FixedRows - 1, i)) Then
                GetStatGridData = .TextMatrix(lngRow, i)
                Exit Function
            End If
        Next
        
        For i = lngCol To .FixedCols Step -1
            If str������� = .TextMatrix(0, i) And strFiledTmp = Trim(.TextMatrix(.FixedRows - 1, i)) Then
                GetStatGridData = .TextMatrix(lngRow, i)
                Exit Function
            End If
        Next
        GetStatGridData = strFiled
    End With
End Function

Private Sub FindItem(ByVal strFind As String, Optional ByVal blnNext As Boolean)
'���ܣ����ҽ����ϵĹؼ���
'������blnNext=������һ��
    Static lngindex As Long
    Static lngMshRow As Long
    Static lngMshcol As Long
    Static strFindLast As String
    Dim objControl As Object
    Dim blntmp As Boolean
    Dim i As Long, j As Long, k As Long
    
    If Trim(strFind) = "" Then Exit Sub
    If strFindLast <> strFind Then lngindex = 0
    strFindLast = strFind
    If lngCurInx <> 0 And lbl(lngCurInx).BackColor = CON_SETFOCES Then lbl(lngCurInx).BackColor = lngTmpColor
    For Each objControl In Me.Controls
        i = i + 1
        'ֻ���ұ�ǩ�ͱ��
        If i >= lngindex Then
            If objControl.name = "lbl" Then
                If i > lngindex Then
                    If objControl.Caption Like "*" & strFind & "*" Then
                        lngCurInx = objControl.Index
                        lngTmpColor = objControl.BackColor
                        objControl.BackColor = CON_SETFOCES
                        lngindex = i
                        blntmp = True
                        Exit Sub
                    End If
                End If
            ElseIf objControl.name = "msh" Then
                If lngindex <> i Then lngMshRow = 0: lngMshcol = 0
                If lngMshRow < objControl.Rows - 1 Or lngMshcol < objControl.Cols - 1 Then
                    For j = objControl.FixedRows To objControl.Rows - 1
                        For k = objControl.FixedCols To objControl.Cols - 1
                            If j = lngMshRow And k > lngMshcol Or j > lngMshRow Then
                                If objControl.TextMatrix(j, k) Like "*" & strFind & "*" Then
                                    objControl.Row = j: objControl.Col = k
                                    objControl.ShowCell j, k
                                    lngindex = i
                                    blntmp = True
                                    lngMshRow = j: lngMshcol = k
                                    objControl.SetFocus
                                    Exit Sub
                                End If
                            End If
                            
                        Next
                    Next
                End If
            End If
        End If
    Next
    If blntmp = False Then
        If lngindex <> 0 Then
            MsgBox "�Ѿ�ȫ��������ɡ�", vbInformation, App.Title
        Else
            MsgBox "û�в��ҵ���ص����֡�", vbInformation, App.Title
        End If
        lngindex = 0
    End If
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call FindItem(txtFind.Text)
    End If
End Sub

Private Sub vsfRelations_LostFocus()
    vsfRelations.Visible = False
End Sub

Private Sub vsfRelations_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strRelationReport As String
    If Button = 2 Then
        Exit Sub
    End If
    
    strRelationReport = vsfRelations.TextMatrix(vsfRelations.Row, vsfRelations.ColIndex("ID"))
    If strRelationReport <> "" Then
        mlngRelationReport = Val(strRelationReport)
        If mbytType = 0 Then
            Call msh_Click(mintGridIndex)
        ElseIf mbytType = 1 Then
            Call lbl_Click(mintLblIndex)
        End If
        mlngRelationReport = 0
        vsfRelations.Visible = False
    End If
End Sub

Private Sub LoadCondsMenu()
    Dim strSQL As String
    Dim i As Integer
    Dim rsPara As ADODB.Recordset
    Dim blnRetry As Boolean
    
    If mlngRPTID = 0 Then Exit Sub
    
    On Error GoTo hErr
    
    'ɾ�������˵�
    For i = mnuPop_Cond.count - 1 To 1 Step -1
        Unload mnuPop_Cond(i)
    Next
    
    blnRetry = True
    strSQL = "Select Distinct ������, �������� From zlRptConds Where ����ID=[1] Order by ������"
    Set rsPara = OpenSQLRecord(strSQL, "��ȡ��������ı�������", mlngRPTID)
    blnRetry = False
    
    With rsPara
        If .RecordCount = 0 Then
            mnuPop_Split1.Visible = False
            mnuPop_Del.Enabled = False
            mintCurCondID = 0
            mintCurMenuIndex = 0
        Else
            mnuPop_Split1.Visible = True
            mnuPop_Del.Enabled = mintCurCondID > 0
            Do While .EOF = False
                i = .AbsolutePosition
                Load mnuPop_Cond(i)
                mnuPop_Cond(i).Caption = Nvl(!��������) & "(&" & i & ")"
                mnuPop_Cond(i).Visible = True
                mnuPop_Cond(i).Tag = Nvl(!������, 0)
                
                If mintCurCondID = Nvl(!������, 0) Then
                    mnuPop_Cond(i).Checked = True
                Else
                    mnuPop_Cond(i).Checked = False
                End If
                
                .MoveNext
            Loop
        End If
        .Close
    End With
            
    mnuPop_Default.Checked = mintCurCondID = 0
    
    Exit Sub
    
hErr:
    If blnRetry Then
        If ErrCenter = 1 Then Resume
    Else
        Call ErrCenter
    End If
End Sub

Public Function GetReportForm(objParent As Object, objCurDLL As clsReport, LibDatas As Object, arrPars As Variant, ByVal bytStyle As Byte) As Object
    Set frmParent = objParent
    Set mobjCurDLL = objCurDLL
    marrPars = arrPars
    mbytStyle = bytStyle
    
    On Error Resume Next
        If Not LibDatas Is Nothing Then Set mLibDatas = LibDatas
        Load Me
    If Err.Number = 0 Then
        Set mobjfrmShowDock = New frmPreviewDock
        If Not mobjReport.blnLoad Then Exit Function
    
        If mobjReport.Items.count = 0 Then Exit Function
        
        If Not InitPrinter(Me) Then
            gblnError = True
            MsgBox "�豸��ʼ��ʧ��.������ϵͳû�а�װ��ӡ�����뵱ǰ���ò����ݣ�", vbInformation, App.Title: Exit Function
        End If
        
        If Not CalcCellPage Then
            gblnError = True
            MsgBox "�޷�����ı���ʽ,�������ܼ�����", vbInformation, App.Title: Exit Function
        End If
        If lbl(lngCurInx).BackColor = CON_SETFOCES And lngCurInx <> 0 Then
            lbl(lngCurInx).BackColor = lngTmpColor
            lngCurInx = 0: lngTmpColor = 0
        End If
        mobjfrmShowDock.BorderStyle = FormBorderStyleConstants.vbBSNone '����Ϊ�ޱ߿�
        mobjfrmShowDock.Caption = mobjfrmShowDock.Caption       '�ص�����һ��
        Set mobjfrmShowDock.frmParent = Me
        Load mobjfrmShowDock
        mobjfrmShowDock.LoadForm 1
        Set LibDatas = mLibDatas
        Set GetReportForm = mobjfrmShowDock
    ElseIf Err.Number <> 0 Then
        '364:������ж��(��Form_Load�ڲ�Unload,��ȡ����������)
        Err.Clear
    End If
End Function

Public Sub PrintReportForRec(objParent As Object, objCurDLL As clsReport, LibDatas As Object, arrPars As Variant, ByVal bytStyle As Byte)
    Set frmParent = objParent
    Set mobjCurDLL = objCurDLL
    marrPars = arrPars
    mbytStyle = bytStyle
    
    On Error Resume Next
    
    If mbytStyle <> 0 Then
        Set mLibDatas = LibDatas
        Load Me
        If Err.Number = 0 Then
            If mbytStyle = 1 Then       '�Զ�Ԥ��
                mnuFile_Preview_Click
            ElseIf mbytStyle = 2 Then   '�Զ���ӡ
                mnuFile_Print_Click
            ElseIf mbytStyle = 3 Then   '�����Excel
                mnuFile_Excel_Click
            ElseIf mbytStyle = 4 Then   '�̶������PDF
                mnuFile_Print_Click
            End If
        ElseIf Err.Number <> 0 Then
            '364:������ж��(��Form_Load�ڲ�Unload,��ȡ����������)
            Err.Clear
        End If
        Unload Me
    Else
        '�ȳ����Է�ģ̬��ʾ����
        If frmParent Is Nothing Then
            Me.Show
        ElseIf frmParent.name = "frmDesign" Then
            Me.Show 1, frmParent
        Else
            Me.Show , frmParent
        End If
        
        '������������Է�ģ̬��ʾ
        If Err.Number = 373 Or Err.Number = 401 Then
            '373:��֧�ֱ������ƻ����������ڲ�����(Դ�������zlReport.dll,��֧�ּӸ�����)
            '401:������ģʽ����ʱ������ʾ��ģʽ����
            '���Զ�Load������ʾʱ�����ټ���Form_Load�¼�
            Err.Clear: Me.Show 1
        ElseIf Err.Number = 364 Then
            '364:������ж��(��Form_Load�ڲ�Unload,��ȡ����������)
            Err.Clear
        ElseIf Err.Number <> 0 Then
            Err.Clear: Unload Me '���Զ�Load��δ֪����ʱж�ش���
        End If
    End If
End Sub

Private Function SQLExistLOB(ByVal clsData As RPTData) As Boolean
'���ܣ��ж�����Դ��SQL�Ƿ����LOB�ֶ�����
    
    Dim arrField As Variant
    Dim i As Integer, intType As Integer
    
    SQLExistLOB = False
    arrField = Split(clsData.�ֶ�, "|")
    For i = 0 To UBound(arrField)
        intType = Val(Split(arrField(i), ",")(1))
        Select Case intType
        Case adBinary, adVarBinary, adLongVarBinary
            SQLExistLOB = True
            Exit For
        End Select
    Next
End Function

