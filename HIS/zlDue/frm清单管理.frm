VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frm�嵥���� 
   Caption         =   "Ӧ�����ѯ"
   ClientHeight    =   6255
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9630
   Icon            =   "frm�嵥����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6255
   ScaleWidth      =   9630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CheckBox chk�������� 
      Caption         =   "��������ʱ���֮�󸶿��δ���嵥"
      Height          =   255
      Left            =   6720
      TabIndex        =   36
      Top             =   2400
      Width           =   3495
   End
   Begin VB.PictureBox picFind 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   6165
      ScaleHeight     =   315
      ScaleWidth      =   3945
      TabIndex        =   33
      Top             =   2835
      Width           =   3945
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   1365
         TabIndex        =   35
         Top             =   0
         Width           =   1785
      End
      Begin VB.Label lblFind 
         AutoSize        =   -1  'True
         Caption         =   "�����Ʋ���"
         Height          =   180
         Left            =   420
         TabIndex        =   34
         Top             =   60
         Width           =   900
      End
      Begin VB.Image imgSearch 
         Height          =   360
         Left            =   15
         Picture         =   "frm�嵥����.frx":08CA
         Top             =   -15
         Width           =   360
      End
   End
   Begin MSComctlLib.ImageList ilt24 
      Left            =   2505
      Top             =   825
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�嵥����.frx":0FB4
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�嵥����.frx":16AE
            Key             =   "FindH"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsList 
      Height          =   2100
      Left            =   2790
      TabIndex        =   31
      Top             =   3195
      Width           =   6780
      _cx             =   11959
      _cy             =   3704
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483644
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483644
      BackColorAlternate=   16777215
      GridColor       =   -2147483633
      GridColorFixed  =   12632256
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   12632256
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frm�嵥����.frx":1DA8
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
   Begin VB.Frame fraSearch 
      Height          =   5010
      Left            =   30
      TabIndex        =   17
      Top             =   765
      Visible         =   0   'False
      Width           =   2370
      Begin VB.Frame fraSplit 
         Height          =   30
         Left            =   0
         TabIndex        =   30
         Top             =   465
         Width           =   2355
      End
      Begin VB.PictureBox picClear 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   45
         ScaleHeight     =   315
         ScaleWidth      =   750
         TabIndex        =   29
         ToolTipText     =   "�����ǰ���������¿�ʼ!"
         Top             =   525
         Width           =   750
         Begin VB.Image img��� 
            Height          =   285
            Left            =   0
            Picture         =   "frm�嵥����.frx":1DE4
            Top             =   15
            Width           =   300
         End
      End
      Begin VB.PictureBox picHelp 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   960
         ScaleHeight     =   330
         ScaleWidth      =   435
         TabIndex        =   28
         ToolTipText     =   "�����ǰ���������¿�ʼ!"
         Top             =   510
         Width           =   435
         Begin VB.Image imgHelp 
            Height          =   270
            Left            =   45
            Picture         =   "frm�嵥����.frx":229A
            Top             =   30
            Width           =   300
         End
      End
      Begin VB.PictureBox PicClose 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2010
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   27
         ToolTipText     =   "�����ǰ���������¿�ʼ!"
         Top             =   135
         Width           =   300
      End
      Begin VB.PictureBox PicSearchBack 
         BackColor       =   &H8000000E&
         Height          =   4125
         Left            =   0
         ScaleHeight     =   4065
         ScaleWidth      =   2280
         TabIndex        =   18
         Top             =   870
         Width           =   2340
         Begin VB.PictureBox picSearch 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   4530
            Left            =   0
            ScaleHeight     =   4530
            ScaleWidth      =   2205
            TabIndex        =   20
            Top             =   -135
            Width           =   2205
            Begin VB.TextBox txt���� 
               Height          =   300
               Left            =   30
               MaxLength       =   13
               TabIndex        =   3
               Top             =   420
               Width           =   1965
            End
            Begin VB.CommandButton cmd���� 
               Caption         =   "��������(&S)"
               Height          =   350
               Left            =   45
               TabIndex        =   6
               Top             =   1380
               Width           =   1245
            End
            Begin VB.TextBox Txt���� 
               Height          =   300
               Left            =   45
               MaxLength       =   50
               TabIndex        =   5
               Top             =   960
               Width           =   1935
            End
            Begin VB.TextBox TxtOther 
               Height          =   300
               Index           =   0
               Left            =   240
               MaxLength       =   50
               TabIndex        =   23
               Top             =   2400
               Visible         =   0   'False
               Width           =   1710
            End
            Begin MSComCtl2.DTPicker DtpOther 
               Height          =   300
               Index           =   0
               Left            =   120
               TabIndex        =   21
               Top             =   3525
               Visible         =   0   'False
               Width           =   1740
               _ExtentX        =   3069
               _ExtentY        =   529
               _Version        =   393216
               CheckBox        =   -1  'True
               CustomFormat    =   "yyyy��MM��dd��"
               DateIsNull      =   -1  'True
               Format          =   292356099
               CurrentDate     =   37131
            End
            Begin VB.CheckBox chkOther 
               BackColor       =   &H8000000E&
               Caption         =   "ĩ��"
               Height          =   225
               Index           =   0
               Left            =   165
               TabIndex        =   22
               Top             =   2820
               Visible         =   0   'False
               Width           =   1515
            End
            Begin VB.Label lbl 
               BackStyle       =   0  'Transparent
               Caption         =   "��Ӧ�̱���"
               Height          =   180
               Index           =   0
               Left            =   45
               TabIndex        =   2
               Top             =   195
               Width           =   1980
            End
            Begin VB.Label lblHit 
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "����ѡ��>>"
               ForeColor       =   &H8000000D&
               Height          =   240
               Left            =   45
               TabIndex        =   26
               Top             =   1800
               Width           =   1905
            End
            Begin VB.Shape shpHit 
               Height          =   2505
               Left            =   45
               Top             =   1785
               Visible         =   0   'False
               Width           =   1920
            End
            Begin VB.Label lbl 
               BackStyle       =   0  'Transparent
               Caption         =   "��Ӧ�����ƻ����"
               Height          =   180
               Index           =   2
               Left            =   60
               TabIndex        =   4
               Top             =   735
               Width           =   1980
            End
            Begin VB.Label lblOther 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����"
               Height          =   180
               Index           =   0
               Left            =   150
               TabIndex        =   25
               Top             =   2145
               Visible         =   0   'False
               Width           =   360
            End
            Begin VB.Label lblDate 
               BackStyle       =   0  'Transparent
               Caption         =   "����ʱ��"
               Height          =   240
               Index           =   0
               Left            =   150
               TabIndex        =   24
               Top             =   3180
               Visible         =   0   'False
               Width           =   1695
            End
         End
         Begin VB.VScrollBar Scr 
            Height          =   4125
            Left            =   2055
            TabIndex        =   19
            Top             =   0
            Width           =   225
         End
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   165
         Index           =   1
         Left            =   60
         TabIndex        =   1
         Top             =   195
         Width           =   1860
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         DrawMode        =   9  'Not Mask Pen
         X1              =   900
         X2              =   900
         Y1              =   555
         Y2              =   795
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000006&
         DrawMode        =   16  'Merge Pen
         X1              =   915
         X2              =   915
         Y1              =   540
         Y2              =   810
      End
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   2595
      Top             =   3330
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�嵥����.frx":2714
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�嵥����.frx":2B6C
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�嵥����.frx":2FC4
            Key             =   "ItemNo"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�嵥����.frx":3418
            Key             =   "No"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�嵥����.frx":3870
            Key             =   "Write"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TabStrip tabSelect 
      Height          =   315
      Left            =   2850
      TabIndex        =   16
      Top             =   2820
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   556
      TabWidthStyle   =   1
      MultiRow        =   -1  'True
      Style           =   2
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "������ϸ��"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "�Ѹ��嵥"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "δ���嵥"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwList 
      Height          =   5355
      Left            =   0
      TabIndex        =   0
      Top             =   750
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   9446
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ils16"
      Appearance      =   1
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   7
      Top             =   5895
      Width           =   9630
      _ExtentX        =   16986
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frm�嵥����.frx":3CC8
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11906
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin MSComctlLib.ImageList ilsCold 
      Left            =   6150
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�嵥����.frx":455C
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�嵥����.frx":477C
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�嵥����.frx":499C
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�嵥����.frx":4BB8
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�嵥����.frx":4DD8
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�嵥����.frx":4FF8
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�嵥����.frx":5214
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�嵥����.frx":5430
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�嵥����.frx":564A
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�嵥����.frx":57A4
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�嵥����.frx":59C0
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�嵥����.frx":5BE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�嵥����.frx":5DFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�嵥����.frx":6014
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�嵥����.frx":622E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsHot 
      Left            =   6750
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�嵥����.frx":6448
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�嵥����.frx":6668
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�嵥����.frx":6888
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�嵥����.frx":6AA4
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�嵥����.frx":6CC4
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�嵥����.frx":6EE4
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�嵥����.frx":7100
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�嵥����.frx":731C
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�嵥����.frx":7536
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�嵥����.frx":7690
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�嵥����.frx":78B0
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�嵥����.frx":7AD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�嵥����.frx":7CEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�嵥����.frx":7F04
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�嵥����.frx":811E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrTool 
      Height          =   780
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   9570
      _ExtentX        =   16880
      _ExtentY        =   1376
      BandCount       =   2
      BandBorders     =   0   'False
      _CBWidth        =   9570
      _CBHeight       =   780
      _Version        =   "6.7.8988"
      Child1          =   "tlbThis"
      MinHeight1      =   720
      Width1          =   11040
      NewRow1         =   0   'False
      MinHeight2      =   0
      NewRow2         =   0   'False
      BandStyle2      =   1
      Begin MSComctlLib.Toolbar tlbThis 
         Height          =   720
         Left            =   165
         TabIndex        =   9
         Top             =   30
         Width           =   9315
         _ExtentX        =   16431
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ilsCold"
         HotImageList    =   "ilsHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   10
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ԥ��"
               Key             =   "PrintView"
               Description     =   "Ԥ��"
               Object.ToolTipText     =   "Ԥ��"
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
               Key             =   "PrintSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Search"
               Description     =   "����"
               Object.ToolTipText     =   "��������"
               Object.Tag             =   "����"
               ImageIndex      =   13
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "����"
               Key             =   "Find"
               Description     =   "��λ"
               Object.ToolTipText     =   "���ݶ�λ"
               Object.Tag             =   "��λ"
               ImageIndex      =   14
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Filter"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageKey        =   "Search"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ˢ��"
               Key             =   "Refresh"
               Description     =   "ˢ��"
               Object.ToolTipText     =   "ˢ��"
               Object.Tag             =   "ˢ��"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "FindSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Description     =   "��������"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Exit"
               Description     =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageIndex      =   11
            EndProperty
         EndProperty
         MouseIcon       =   "frm�嵥����.frx":8338
         Begin MSComctlLib.ImageList iltHelp 
            Left            =   3825
            Top             =   210
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   20
            ImageHeight     =   18
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   4
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm�嵥����.frx":8652
                  Key             =   "HELPB"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm�嵥����.frx":8ADC
                  Key             =   "HELPC"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm�嵥����.frx":8F66
                  Key             =   "SEARCHB"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm�嵥����.frx":942C
                  Key             =   "SEARCHC"
               EndProperty
            EndProperty
         End
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsHead 
      Height          =   1245
      Left            =   2865
      TabIndex        =   32
      Top             =   1140
      Width           =   6765
      _cx             =   11933
      _cy             =   2196
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483644
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483644
      BackColorAlternate=   16777215
      GridColor       =   -2147483633
      GridColorFixed  =   12632256
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   12632256
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frm�嵥����.frx":98F2
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
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "��λ����"
      Height          =   180
      Left            =   3945
      TabIndex        =   13
      Top             =   2475
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label lblTemp 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   1
      Left            =   2820
      TabIndex        =   12
      Top             =   2400
      Visible         =   0   'False
      Width           =   3690
   End
   Begin VB.Label lblVsc_s 
      Height          =   75
      Left            =   2820
      MousePointer    =   7  'Size N S
      TabIndex        =   15
      Top             =   2730
      Width           =   6750
   End
   Begin VB.Label lblHsc_s 
      Height          =   5355
      Left            =   2745
      MousePointer    =   9  'Size W E
      TabIndex        =   14
      Top             =   750
      Width           =   60
   End
   Begin VB.Label lblTemp 
      AutoSize        =   -1  'True
      Caption         =   "������Ϣ"
      Height          =   180
      Index           =   2
      Left            =   5595
      TabIndex        =   11
      Top             =   825
      Width           =   720
   End
   Begin VB.Label lblTemp 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Index           =   0
      Left            =   2820
      TabIndex        =   10
      Top             =   750
      Width           =   6750
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFilePrintSet 
         Caption         =   "��ӡ����(&S)"
      End
      Begin VB.Menu mnuFilePreView 
         Caption         =   "��ӡԤ��(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "��ӡ(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "�����&Excel"
      End
      Begin VB.Menu mnuFileLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileLocalSet 
         Caption         =   "���ز�������(&R)"
      End
      Begin VB.Menu mnuFileLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Visible         =   0   'False
      Begin VB.Menu mnuEditAdd 
         Caption         =   "����(&A)"
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "�޸�(&M)"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "ɾ��(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSplit 
         Caption         =   "�ƻ�(&S)"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "����(&R)"
      Visible         =   0   'False
      Begin VB.Menu mnuReportItem 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "�鿴(&V)"
      Begin VB.Menu mnuViewTool 
         Caption         =   "������(&T)"
         Begin VB.Menu mnuViewToolButton 
            Caption         =   "��׼��ť(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu menuViewLine1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewToolText 
            Caption         =   "�ı���ǩ(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "״̬��(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewUnit 
         Caption         =   "ҩƷ��Ӧ��(&L)"
         Index           =   0
      End
      Begin VB.Menu mnuViewUnit 
         Caption         =   "���ʹ�Ӧ��(&M)"
         Index           =   1
      End
      Begin VB.Menu mnuViewUnit 
         Caption         =   "�豸��Ӧ��(&E)"
         Index           =   2
      End
      Begin VB.Menu mnuViewUnit 
         Caption         =   "������Ӧ��(&O)"
         Index           =   3
      End
      Begin VB.Menu mnuViewUnit 
         Caption         =   "���Ĺ�Ӧ��(&W)"
         Index           =   4
      End
      Begin VB.Menu mnuViewLine4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewOpen 
         Caption         =   "��������(&J)"
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "���ݶ�λ(&F)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuViewFilter 
         Caption         =   "����(&S)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewLine5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
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
         Caption         =   "&Web�ϵ�����"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "������ҳ(&H)"
         End
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "������̳(&F)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "���ͷ���(&K)"
         End
      End
      Begin VB.Menu mnuHelpLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)"
      End
   End
End
Attribute VB_Name = "frm�嵥����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mdtBegin As Date, mdtEnd As Date
Private mstrKey As String, mstrData As String
Private mstrType As String    '��Ӧ������
Private msngDownX As Single, msngDownY As Single, mintOldSel As Integer, mstrDeptWhere As String
Private mlngModule As Long

Dim mlng��λID As Long      '�ϴ�ѡ��Ĺ�Ӧ��
Dim mstrPrivs As String
Dim mblnFirst As Boolean
Private mstrFiler As String     '����
Private mstrOthers() As String
Private mint��λ As Integer
Private mintFlag As Integer

Private Sub chkOther_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
        ScrCtl chkOther(Index)
    End If
End Sub
'����26224 by lesfeng 2010-02-08
Private Sub chk��������_Click()
    mintFlag = 1
End Sub

Private Sub DtpOther_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
        ScrCtl DtpOther(Index)
    End If
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    Call LoadOtherCon
    Me.imgHelp.Visible = False
    Me.picHelp.Visible = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)        '
        If KeyAscii = Asc("'") Then
            KeyAscii = 0
        End If
End Sub

Private Sub Form_Load()
    Dim strOthers(0 To 16) As String
    
    mintOldSel = 1
    mstrOthers = strOthers
    mstrPrivs = gstrPrivs: mlngModule = glngModul: mstrKey = "": mblnFirst = True
    
    mlng��λID = Val(zlDatabase.GetPara("�ϴ�ѡ��λID", glngSys, mlngModule))
    '����26224 by lesfeng 2010-02-08
    chk��������.Value = IIf(Val(zlDatabase.GetPara("����ʱ���֮�󸶿�", glngSys, mlngModule)) = 1, 1, 0)
    '����27878 by lesfeng 2010-02-25
    mint��λ = Val(zlDatabase.GetPara("��λ", glngSys, mlngModule))
    mintFlag = 0
    Call initvsHeadHead(True)
    
    mdtEnd = CDate(Format(zlDatabase.Currentdate, "yyyy-MM-dd"))
    mdtBegin = DateAdd("m", -1, mdtEnd) + 1
    mstrData = "00000"
    
    If Check���Ȩ��(mstrPrivs, "ҩƷ") = False Then
        mnuViewUnit(0).Checked = False
        mnuViewUnit(0).Enabled = False
    End If
    If Check���Ȩ��(mstrPrivs, "����") = False Then
        mnuViewUnit(1).Checked = False
        mnuViewUnit(1).Enabled = False
    End If
    If Check���Ȩ��(mstrPrivs, "�豸") = False Then
        mnuViewUnit(2).Checked = False
        mnuViewUnit(2).Enabled = False
    End If
    If Check���Ȩ��(mstrPrivs, "����") = False Then
        mnuViewUnit(3).Checked = False
        mnuViewUnit(3).Enabled = False
    End If
    If Check���Ȩ��(mstrPrivs, "����") = False Then
        mnuViewUnit(4).Checked = False
        mnuViewUnit(4).Enabled = False
    End If
    '�ָ���ز���
    RestoreWinState Me, App.ProductName
    InitColHead tabSelect.SelectedItem.Index
    FullDept
    '2006-04-25:���˺�,ͳһ���ӱ�������ģ��Ĺ���
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs)
    
    Call vsList_AfterRowColChange(1, 0, 1, 0)
End Sub
'����27878 by lesfeng 2010-02-25
Private Sub mnuFileLocalSet_Click()
    '���ز�������
    frm�嵥����Set.�������� Me, mlngModule, mstrPrivs
    mint��λ = Val(zlDatabase.GetPara("��λ", glngSys, mlngModule))
    '���³�ʹ����
    Select Case tabSelect.SelectedItem.Index
        Case 1
            zl_vsGrid_Para_Save mlngModule, vsList, Me.Caption, "������ϸ�б�"
            InitColHead 1
            Full������ϸ
        Case 2
            zl_vsGrid_Para_Save mlngModule, vsList, Me.Caption, "�Ѹ���ϸ�б�"
            InitColHead 2
            Full�Ѹ��嵥
        Case 3
             zl_vsGrid_Para_Save mlngModule, vsList, Me.Caption, "δ����ϸ�б�"
             InitColHead 3
            Fullδ���嵥
    End Select
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    Dim lng����id As Long
    Dim lng��Ӧ��ID As Long
    
    If Not tvwList.SelectedItem Is Nothing Then
        lng����id = Val(Mid(Me.tvwList.SelectedItem.Key, 2))
    End If
    
    lng��Ӧ��ID = vsHead.RowData(vsHead.Row)
    '2006-04-25:���˺�:�����Զ��屨������ģ��Ĺ���
    Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, "����=" & lng����id, "��Ӧ��=" & lng��Ӧ��ID)
End Sub

Private Sub FullCount()
    Dim rstList As New ADODB.Recordset, strSQL As String
'    '����26224 by lesfeng 2010-02-08
    If Me.mnuViewFilter.Checked Then
    Else
        If tvwList.SelectedItem.Key = mstrKey And mintFlag = 0 Then Exit Sub
    End If
        
    FillSum
    vsHead_EnterCell
End Sub

Private Sub initvsHeadHead(Optional bln��ʼ As Boolean = False)
    '��ʼ��ͷ
    Dim i As Long
    With vsHead
        .Redraw = False
        .Clear 1
        .Rows = 2
        .Cols = 5
        .TextMatrix(0, 0) = "��Ӧ������": .ColKey(0) = "��Ӧ������"
        .TextMatrix(0, 1) = "�ڳ�Ӧ��": .ColKey(1) = "�ڳ�Ӧ��"
        .TextMatrix(0, 2) = "�����޹�": .ColKey(2) = "�����޹�"
        .TextMatrix(0, 3) = "����֧��": .ColKey(3) = "����֧��"
        .TextMatrix(0, 4) = "��ĩӦ��": .ColKey(4) = "��ĩӦ��"
        For i = 0 To .Cols - 1
            .FixedAlignment(i) = flexAlignCenterCenter
            If .ColIndex("��Ӧ������") = i Then
                .ColAlignment(i) = 1
            Else
                .ColAlignment(i) = 7
            End If
        Next
        
        If bln��ʼ Then
            .ColWidth(.ColIndex("��Ӧ������")) = 3000
            .ColWidth(.ColIndex("�ڳ�Ӧ��")) = 1200
            .ColWidth(.ColIndex("�����޹�")) = 1200
            .ColWidth(.ColIndex("����֧��")) = 1200
            .ColWidth(.ColIndex("��ĩӦ��")) = 1200
        End If
        zl_vsGrid_Para_Restore mlngModule, vsHead, Me.Caption, "�����Ϣ�б�", True
        .Redraw = True
    End With
End Sub

Private Sub InitColHead(ByVal intType As Integer, Optional bln��ʼ As Boolean = True)
    '��ʼ��ͷ
    Dim i As Long
    If intType = 1 Then
        With vsList
                .Redraw = False
                .Clear
                .ExplorerBar = flexExMove
                .Rows = 2
                .FormatString = "^����|^���ݺ�|^ժҪ|^��λ|^����|^�ɹ�����|^�ɹ���|^Ӧ�����|^�Ѹ����|^���"
                .MergeCells = flexMergeNever
                .SelectionMode = flexSelectionByRow
                For i = 0 To .Cols - 1
                     .ColKey(i) = .TextMatrix(0, i)
                     .FixedAlignment(i) = flexAlignCenterCenter
                     Select Case i
                     Case .ColIndex("�ɹ�����"), .ColIndex("�ɹ���"), .ColIndex("Ӧ�����"), .ColIndex("�Ѹ����"), .ColIndex("���")
                         .ColAlignment(i) = flexAlignRightCenter
                         If bln��ʼ Then .ColWidth(i) = 1000
                     Case .ColIndex("����"), .ColIndex("���ݺ�"), .ColIndex("��λ")
                         .ColAlignment(i) = flexAlignCenterCenter
                         If bln��ʼ Then .ColWidth(i) = 1400
                     Case Else
                         .ColAlignment(i) = flexAlignLeftCenter
                         If bln��ʼ Then .ColWidth(i) = 1400
                     End Select
                Next
                zl_vsGrid_Para_Restore mlngModule, vsList, Me.Caption, "������ϸ�б�", True
               .Redraw = True
           End With
           Exit Sub
    End If
    If intType = 2 Then
        With vsList
            .Redraw = False
            .Clear
            .Rows = 2
            .ExplorerBar = flexExMove
            .FormatString = "^��ⵥ�ݺ�|^��Ʊ��|^��Ʊ����|^��Ʊ���|^����ݺ�|^����|^Ʒ��|^���|^��λ|^����|^����|^���"
            .MergeCells = flexMergeNever
            .SelectionMode = flexSelectionByRow
            For i = 0 To .Cols - 1
                 .ColKey(i) = .TextMatrix(0, i)
                 .FixedAlignment(i) = flexAlignCenterCenter
                 Select Case i
                 Case .ColIndex("����"), .ColIndex("���"), .ColIndex("��Ʊ���")
                     .ColAlignment(i) = flexAlignRightCenter
                     If bln��ʼ Then .ColWidth(i) = 1000
                 Case .ColIndex("��ⵥ�ݺ�"), .ColIndex("��Ʊ����"), .ColIndex("��λ"), .ColIndex("����")
                     .ColAlignment(i) = flexAlignCenterCenter
                     If bln��ʼ Then .ColWidth(i) = 1400
                 Case Else
                     .ColAlignment(i) = flexAlignLeftCenter
                     If bln��ʼ Then .ColWidth(i) = 1400
                 End Select
            Next
            zl_vsGrid_Para_Restore mlngModule, vsList, Me.Caption, "�Ѹ���ϸ�б�", True
            .Redraw = True
        End With
        Exit Sub
    End If
    With vsList
        .Redraw = False
        .Clear
        .Rows = 2
        .FormatString = "^����|^���ݺ�|^Ʒ��|^���|^��λ|^����|^��ⵥ�ݺ�|^����|^���ݽ��|^���|^��Ʊ��|^��Ʊ����|^�ƻ����|^�ƻ���������"
        .ExplorerBar = flexExMove
        .MergeCells = flexMergeNever
        .SelectionMode = flexSelectionByRow
        For i = 0 To .Cols - 1
             .ColKey(i) = .TextMatrix(0, i)
             .FixedAlignment(i) = flexAlignCenterCenter
             Select Case i
             Case .ColIndex("����"), .ColIndex("���ݽ��"), .ColIndex("���")
                 .ColAlignment(i) = flexAlignRightCenter
                 If bln��ʼ Then .ColWidth(i) = 1000
             Case .ColIndex("����"), .ColIndex("���ݺ�"), .ColIndex("��λ"), .ColIndex("��Ʊ����")
                 .ColAlignment(i) = flexAlignCenterCenter
                 If bln��ʼ Then .ColWidth(i) = 1400
             Case Else
                 .ColAlignment(i) = flexAlignLeftCenter
                 If bln��ʼ Then .ColWidth(i) = 1400
             End Select
        Next
        .MergeCol(0) = True
        .MergeCol(1) = True
        .MergeCol(2) = True
        If bln��ʼ Then
            .ColWidth(.ColIndex("Ʒ��")) = 2000
        End If
        zl_vsGrid_Para_Restore mlngModule, vsList, Me.Caption, "δ����ϸ�б�", True
        .Redraw = True
    End With
End Sub

Private Sub GetTypeCon()
    '��ȡ��������
    Dim intIndex As Integer
    Dim strTmp As String
    strTmp = ""
    For intIndex = 0 To 4
        If mnuViewUnit(intIndex).Checked Then
            strTmp = strTmp & "1"
        Else
            strTmp = strTmp & "0"
        End If
    Next
    mstrType = strTmp ' Bin2Dec(strTmp)
End Sub

Private Sub FullDept()
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��乩Ӧ��
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim intFilt As Integer
    Dim i As Long
    Dim str���� As String
    Dim strFilt As String
    
    Call GetTypeCon
    
    str���� = ""
    For i = 1 To Len(mstrType)
        If Mid(mstrType, i, 1) = 1 Then
            str���� = str���� & " or substr(����," & i & ",1)=1"
        End If
    Next
    If str���� <> "" Then
        str���� = " And (" & Mid(str����, 4) & ") "
    End If

    Dim strȨ�� As String
    strȨ�� = " and  " & Get����Ȩ��(gstrPrivs)
                
    strSQL = "" & _
        "   Select ID,�ϼ�ID,����,����,ĩ�� " & _
        "   From ��Ӧ�� " & _
        "   Where (����ʱ��=TO_DATE('3000-1-1','yyyy-MM-dd') or ����ʱ�� is null ) " & _
        "           and (ĩ��=0 Or (ĩ��=1 " & zl_��ȡվ������() & " " & str���� & strȨ�� & "))" & _
        "   Start with �ϼ�ID is null Connect by prior ID =�ϼ�ID"
    On Error GoTo errHandle
    zlDatabase.OpenRecordset rsTemp, strSQL, Me.Caption
    Dim curNode As Node
    tvwList.Nodes.Clear
    tvwList.Nodes.Add , , "Root", "���й�Ӧ��", 1
    tvwList.Nodes("Root").Selected = True
    tvwList.Nodes("Root").Expanded = True
    tvwList.Nodes("Root").Sorted = True
    While Not rsTemp.EOF
        If IsNull(rsTemp!�ϼ�ID) Then
            Set curNode = tvwList.Nodes.Add("Root", tvwChild, "K" & rsTemp!ID, "��" & rsTemp!���� & "��" & rsTemp!����, IIf(rsTemp!ĩ�� <> 1, 5, 2))
        Else
            Set curNode = tvwList.Nodes.Add("K" & rsTemp!�ϼ�ID, tvwChild, "K" & rsTemp!ID, "��" & rsTemp!���� & "��" & rsTemp!����, IIf(rsTemp!ĩ�� <> 1, 5, 2))
        End If
        If Nvl(rsTemp!ID) = mlng��λID Then
            curNode.Selected = True
            curNode.Expanded = True
        End If
        curNode.Sorted = True
        rsTemp.MoveNext
    Wend
    FullCount
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Full������ϸ()
    Dim rsTemp As New ADODB.Recordset
    Dim strBegin As String, strEnd As String
    Dim dblSum(1 To 2) As Double, dblBalance As Double
    Dim lngRow As Long, lngID As Long
    Dim strSelect As String
  
    '�õ���ѯ����
    lngID = Val(vsHead.Cell(flexcpData, vsHead.Row, vsHead.ColIndex("��Ӧ������")))
    If lngID = 0 Then Exit Sub
    
    '��ʼ��ѯ
    strBegin = Format(mdtBegin, "yyyy-MM-dd")
    strEnd = Format(mdtEnd + 1, "yyyy-MM-dd")
    On Error GoTo errHandle
    '���ȵõ���ĩ���
    gstrSQL = "" & _
        "   Select sum(nvl(���,0)) as ��� " & _
        "   From(   Select ���  From �����¼  " & _
        "           Where �������>=[2] and ��λID=[1]" & _
        "           Union All " & _
        "           Select -1 * nvl(��Ʊ���,0) as ��� from Ӧ����¼ " & _
        "           Where �������>=[2] and ��λID=[1] " & _
        "           Union All " & _
        "           Select ���  From Ӧ����� " & _
        "           Where ����=1 and ��λID=[1]) "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngID, CDate(strEnd))
    If Not rsTemp.EOF Then
        dblBalance = IIf(IsNull(rsTemp("���")), 0, rsTemp("���"))
    Else
        dblBalance = 0
    End If
    
    '�ٵõ���ϸ��
    '����27878 by lesfeng 2010-02-25
    Select Case mint��λ
    Case 0
        strSelect = "A.������λ,nvl(A.����,0) as �ɹ�����,nvl(A.�ɹ���,0) as �ɹ���,"
    Case 1
        strSelect = "B.���ﵥλ as ������λ,nvl(A.����,0)/decode(B.�����װ,null,1,0,1,B.�����װ) as �ɹ����� ,nvl(A.�ɹ���,0)*nvl(B.�����װ,1) as �ɹ���,"
    Case 2
        strSelect = "B.סԺ��λ as ������λ,nvl(A.����,0)/decode(B.סԺ��װ,null,1,0,1,B.סԺ��װ) as �ɹ����� ,nvl(A.�ɹ���,0)*nvl(B.סԺ��װ,1) as �ɹ���,"
    Case 3
        strSelect = "B.ҩ�ⵥλ as ������λ,nvl(A.����,0)/decode(B.ҩ���װ,null,1,0,1,B.ҩ���װ) as �ɹ����� ,nvl(A.�ɹ���,0)*nvl(B.ҩ���װ,1) as �ɹ���,"
    End Select
    '����27930 by lesfeng 2010-03-23
    gstrSQL = " Select * From ( " & _
              "  Select to_char(�������,'yyyy-MM-dd') as ����,decode(�ܸ���־,0,'��','��')||NO as NO, " & _
              "       decode(Ԥ����,1,'Ԥ����',decode(mod(��¼״̬,3),2,'������¼',ժҪ))||decode(���㷽ʽ,'','','('||���㷽ʽ||')') as ժҪ, " & _
              "       '' as ����,'' as ��λ,0 as �ɹ�����,0 as �ɹ���,0 as Ӧ�����,nvl(���,0) as �Ѹ���� " & _
              "       From �����¼ " & _
              "       where �������>=[1] and �������<[2] and ��λID=[3]" & _
              " "
    'ҩƷ����
    gstrSQL = gstrSQL & "  Union All  " & _
              "  select to_char(A.�������,'yyyy-MM-dd') as ����,A.��ⵥ�ݺ�,A.Ʒ��||decode(A.���,null,'','('||A.���||')') as ժҪ, " & _
              "       A.����," & strSelect & "nvl(A.��Ʊ���,0) as Ӧ�����,0 as �Ѹ���� " & _
              "       from Ӧ����¼ A,ҩƷ��� B " & _
              "       where not A.��¼���� in (-1,2) And A.�������>=[1] and A.�������<[2] and A.��λID=[3]" & _
              "         And A.��ĿID = B.ҩƷid And A.ϵͳ��ʶ = 1 "
    '��ҩƷ����
    gstrSQL = gstrSQL & "  Union All  " & _
              "  select to_char(�������,'yyyy-MM-dd') as ����,��ⵥ�ݺ�,Ʒ��||decode(���,null,'','('||���||')') as ժҪ, " & _
              "       ����,������λ,nvl(����,0),nvl(�ɹ���,0),nvl(��Ʊ���,0) as Ӧ�����,0 as �Ѹ���� " & _
              "       from Ӧ����¼  " & _
              "       where not ��¼���� in (-1,2) And �������>=[1] and �������<[2] and ��λID=[3]  And ϵͳ��ʶ <> 1 " & _
              ")  order by ����,no"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CDate(strBegin), CDate(strEnd), lngID)
    
    If rsTemp.RecordCount = 0 Then
        Exit Sub
    End If
    vsList.Rows = rsTemp.RecordCount + 3
    lngRow = 2
    With vsList
        .Redraw = False
        '"^����|^���ݺ�|^ժҪ|^��λ|^����|^�ɹ�����|^�ɹ���|^Ӧ�����|^�Ѹ����|^���"
        
        .TextMatrix(1, .ColIndex("����")) = Format(mdtBegin, "yyyy-MM-dd")
        .TextMatrix(1, .ColIndex("ժҪ")) = "�ڳ����"
        
        Do Until rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("����")) = Nvl(rsTemp("����"))
            .TextMatrix(lngRow, .ColIndex("���ݺ�")) = Nvl(rsTemp("NO"))
            .TextMatrix(lngRow, .ColIndex("ժҪ")) = Nvl(rsTemp("ժҪ"))
            .TextMatrix(lngRow, .ColIndex("��λ")) = Nvl(rsTemp("��λ"))
            .TextMatrix(lngRow, .ColIndex("����")) = Nvl(rsTemp("����"))
            .TextMatrix(lngRow, .ColIndex("�ɹ�����")) = Format(Val(Nvl(rsTemp("�ɹ�����"))), gVbFmtString.FM_����)
            .TextMatrix(lngRow, .ColIndex("�ɹ���")) = Format(Val(Nvl(rsTemp("�ɹ���"))), gVbFmtString.FM_�ɱ���)
            .TextMatrix(lngRow, .ColIndex("Ӧ�����")) = Format(Val(Nvl(rsTemp("Ӧ�����"))), gVbFmtString.FM_���)
            .TextMatrix(lngRow, .ColIndex("�Ѹ����")) = Format(Val(Nvl(rsTemp("�Ѹ����"))), gVbFmtString.FM_���)
            dblSum(1) = dblSum(1) + Nvl(rsTemp("Ӧ�����"), 0)
            dblSum(2) = dblSum(2) + Nvl(rsTemp("�Ѹ����"), 0)
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        
        .TextMatrix(lngRow, .ColIndex("����")) = Format(mdtEnd, "yyyy-MM-dd")
        .TextMatrix(lngRow, .ColIndex("ժҪ")) = "�ϼ�"
        .TextMatrix(lngRow, .ColIndex("Ӧ�����")) = Format(dblSum(1), gVbFmtString.FM_���)
        .TextMatrix(lngRow, .ColIndex("�Ѹ����")) = Format(dblSum(2), gVbFmtString.FM_���)
        .TextMatrix(lngRow, .ColIndex("���")) = Format(dblBalance, gVbFmtString.FM_���)
        
        Do Until lngRow = 1
            lngRow = lngRow - 1
            .TextMatrix(lngRow, .ColIndex("���")) = Format(dblBalance, gVbFmtString.FM_���)
            dblBalance = dblBalance + Val(.TextMatrix(lngRow, .ColIndex("�Ѹ����"))) - Val(.TextMatrix(lngRow, .ColIndex("Ӧ�����")))
        Loop
        If .Rows - 1 >= 2 Then .Row = 1
        .Col = 0: .LeftCol = 0
        .Redraw = True
    End With
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Full�Ѹ��嵥()
    Dim rsTemp As New ADODB.Recordset, intFilt As Integer, strFilt As String
    Dim strBegin As String, strEnd As String
    Dim dblSum As Double
    Dim lngRow As Long, lngCount As Long, lngTemp As Long
    Dim lngID As Long
    Dim strSelect As String
    
    
    '�õ���ѯ����
    lngID = Val(vsHead.Cell(flexcpData, vsHead.Row, vsHead.ColIndex("��Ӧ������")))
    If lngID = 0 Then Exit Sub
    '��ʼ��ѯ
    strBegin = Format(mdtBegin, "yyyy-MM-dd") & " 00:00:00"
    strEnd = Format(mdtEnd, "yyyy-MM-dd") & " 23:59:59"
    '�õ��Ѹ��嵥
    'by lesfeng 2009-12-2 �����Ż�
    '����27878 by lesfeng 2010-02-25
    Select Case mint��λ
    Case 0
        strSelect = "Max(A.������λ) As ��λ,Sum(Nvl(A.����, 0)) as ����,"
    Case 1
        strSelect = "Max(C.���ﵥλ) as ��λ,Sum(nvl(A.����,0)/decode(C.�����װ,null,1,0,1,C.�����װ)) as ����,"
    Case 2
        strSelect = "Max(C.סԺ��λ) as ��λ,Sum(nvl(A.����,0)/decode(C.סԺ��װ,null,1,0,1,C.סԺ��װ)) as ����,"
    Case 3
        strSelect = "Max(C.ҩ�ⵥλ) as ��λ,Sum(nvl(A.����,0)/decode(C.ҩ���װ,null,1,0,1,C.ҩ���װ)) as ����,"
    End Select
    '����27930 by lesfeng 2010-03-23
    gstrSQL = "Select * From ( " & _
             " Select A.���,Max(A.��ⵥ�ݺ�) || Decode(A.�ƻ����, Null, ' ', 0, ' ', '(�ƻ�:' || a.�ƻ���� || ')') As ��ⵥ�ݺ�, " & _
             "       sum(Nvl(A.��Ʊ���, 0)) As ���,B.NO   As ����ݺ�, " & _
             "       To_Char(Max(B.�������), 'yyyy-MM-dd') As ����, Max(A.Ʒ��) As ����, Max(A.���) ���, " & _
             "       Max(A.��Ʊ��) ��Ʊ��, To_Char(Max(A.��Ʊ����), 'yyyy-mm-dd') ��Ʊ����, Max(A.����) ����," & strSelect & _
             "       Sum(Decode(A.��¼����, -1, Nvl(A.�ƻ����, 0), Nvl(A.��Ʊ���, 0))) As ������ " & _
             " From��Ӧ����¼ A,ҩƷ��� C," & _
             "    (Select Distinct �������,NO||'('||decode(Ԥ����,1,'��Ԥ��',decode(�ܸ���־,1,'���','����'))||')' As No,������� " & _
             "     From �����¼ " & _
             "     Where nvl(Ԥ����,0)<>1 And ������� Between [1] And [2] " & _
             "           And ��λid=[3] )  B " & _
             " Where a.�������=b.������� And A.��λid=[3] And A.��ĿID = C.ҩƷid And A.ϵͳ��ʶ = 1 and a.��¼����<>2 " & _
             " Group By A.ϵͳ��ʶ, A.��¼����, A.NO, A.��Ŀid, A.���, A.�ƻ����, B.NO "
             
    gstrSQL = gstrSQL & "  Union All  " & _
             " Select A.���,Max(A.��ⵥ�ݺ�) || Decode(A.�ƻ����, Null, ' ', 0, ' ', '(�ƻ�:' || a.�ƻ���� || ')') As ��ⵥ�ݺ�, " & _
             "       sum(Nvl(A.��Ʊ���, 0)) As ���,B.NO   As ����ݺ�, " & _
             "       To_Char(Max(B.�������), 'yyyy-MM-dd') As ����, Max(A.Ʒ��) As ����, Max(A.���) ���, " & _
             "       Max(A.��Ʊ��) ��Ʊ��, To_Char(Max(A.��Ʊ����), 'yyyy-mm-dd') ��Ʊ����, Max(A.����) ����," & _
             "       Max(A.������λ) As ��λ,Sum(Nvl(A.����, 0)) ����," & _
             "       Sum(Decode(A.��¼����, -1, Nvl(A.�ƻ����, 0), Nvl(A.��Ʊ���, 0))) As ������ " & _
             " From��Ӧ����¼ A," & _
             "    (Select Distinct �������,NO||'('||decode(Ԥ����,1,'��Ԥ��', decode(�ܸ���־,1,'���','����'))||')' As No,������� " & _
             "     From �����¼ " & _
             "     Where nvl(Ԥ����,0)<>1 And ������� Between [1] And [2] " & _
             "           And ��λid=[3] )  B " & _
             " Where a.�������=b.������� And A.��λid=[3] And A.ϵͳ��ʶ <> 1 and a.��¼����<>2 " & _
             " Group By A.ϵͳ��ʶ, A.��¼����, A.NO, A.��Ŀid, A.���, A.�ƻ����, B.NO)" & _
             " Order by ��ⵥ�ݺ�,���"
                 
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CDate(strBegin), CDate(strEnd), lngID)
    If rsTemp.RecordCount = 0 Then
        Exit Sub
    Else
        vsList.Rows = rsTemp.RecordCount + 2
    End If
    lngRow = 1
    With vsList
        .Redraw = False
        Do Until rsTemp.EOF
            ' "^��ⵥ�ݺ�|^��Ʊ��|^��Ʊ����|^��Ʊ���|^����ݺ�|^����|^Ʒ��|^���|^��λ|^����|^����|^���"
            .TextMatrix(lngRow, .ColIndex("��ⵥ�ݺ�")) = Nvl(rsTemp!��ⵥ�ݺ�)
            .TextMatrix(lngRow, .ColIndex("��Ʊ��")) = Nvl(rsTemp!��Ʊ��)
            .TextMatrix(lngRow, .ColIndex("��Ʊ����")) = Nvl(rsTemp!��Ʊ����)
            .TextMatrix(lngRow, .ColIndex("��Ʊ���")) = Format(Val(Nvl(rsTemp!���)), gVbFmtString.FM_���)
            .TextMatrix(lngRow, .ColIndex("����ݺ�")) = Nvl(rsTemp!����ݺ�)
            .TextMatrix(lngRow, .ColIndex("����")) = Nvl(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("Ʒ��")) = Nvl(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("���")) = Nvl(rsTemp!���)
            .TextMatrix(lngRow, .ColIndex("��λ")) = Nvl(rsTemp!��λ)
            .TextMatrix(lngRow, .ColIndex("����")) = Nvl(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("����")) = Format(Val(Nvl(rsTemp!����)), gVbFmtString.FM_����)
            .TextMatrix(lngRow, .ColIndex("���")) = Format(Val(Nvl(rsTemp!������)), gVbFmtString.FM_���)
            dblSum = dblSum + Val(Nvl(rsTemp!������))
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        If .Rows > 2 Then
            .TextMatrix(lngRow, .ColIndex("��ⵥ�ݺ�")) = "�ϼ�"
            .TextMatrix(lngRow, .ColIndex("���")) = Format(dblSum, gVbFmtString.FM_���)
        End If
        If .Rows - 1 >= 2 Then .Row = 1
        .Col = 0: .LeftCol = 0
        .Redraw = True
    End With
End Sub

Private Sub Fullδ���嵥()
    Dim rsTemp As New ADODB.Recordset
    Dim dtStartdate As Date, dtEndDate As Date
    Dim dblSum As Double
    Dim lngRow As Long, lngCount As Long, lngTemp As Long
    Dim lngID As Long
    Dim strTemp As String
    Dim strSelect As String
    Dim strDSelect As String
    '�õ���ѯ����
    lngID = Val(vsHead.Cell(flexcpData, vsHead.Row, vsHead.ColIndex("��Ӧ������")))
    If lngID = 0 Then Exit Sub
    '��ʼ��ѯ
    dtStartdate = CDate(Format(mdtBegin, "yyyy-mm-dd") & " 00:00:00")
    dtEndDate = CDate(Format(mdtEnd + 1, "yyyy-mm-dd") & " 00:00:00")
    '����26224 by lesfeng 2010-02-08
    If chk��������.Value = 1 Then
        strTemp = " or ������� >= [3]"
    Else
        strTemp = ""
    End If
    '�õ�δ���嵥
    '����27878 by lesfeng 2010-02-25
    Select Case mint��λ
    Case 0
        strSelect = "Max(A.������λ) As ��λ,Sum(Nvl(A.����, 0)) as ����,"
        strDSelect = "A.������λ As ��λ, Nvl(A.����, 0) as ����,"
    Case 1
        strSelect = "Max(C.���ﵥλ) as ��λ,Sum(nvl(A.����,0)/decode(C.�����װ,null,1,0,1,C.�����װ)) as ����,"
        strDSelect = "C.���ﵥλ as ��λ,nvl(A.����,0)/decode(C.�����װ,null,1,0,1,C.�����װ) as ����,"
    Case 2
        strSelect = "Max(C.סԺ��λ) as ��λ,Sum(nvl(A.����,0)/decode(C.סԺ��װ,null,1,0,1,C.סԺ��װ)) as ����,"
        strDSelect = "C.סԺ��λ as ��λ,nvl(A.����,0)/decode(C.סԺ��װ,null,1,0,1,C.סԺ��װ) as ����,"
    Case 3
        strSelect = "Max(C.ҩ�ⵥλ) as ��λ,Sum(nvl(A.����,0)/decode(C.ҩ���װ,null,1,0,1,C.ҩ���װ)) as ����,"
        strDSelect = "C.ҩ�ⵥλ as ��λ,nvl(A.����,0)/decode(C.ҩ���װ,null,1,0,1,C.ҩ���װ) as ����,"
    End Select
    
    gstrSQL = " " & _
             "  Select  no as ���,0 as ��־,a.NO,max(a.��ⵥ�ݺ�) ��ⵥ�ݺ�,            " & _
             "      to_char(max(A.�������),'yyyy-MM-dd') as ����,max(A.��Ʊ��) ��Ʊ��,to_char(max(A.��Ʊ����),'yyyy-MM-dd') as ��Ʊ����,            " & _
             "      null as �ƻ�����,max(a.�ƻ����) �ƻ����,null,max(decode(��¼״̬,3,A.ID,1,A.ID,0)) ID,             " & _
             "      max(A.Ʒ��) as ����,max(���) ���,max(����) ����," & strSelect & "sum(nvl(A.��Ʊ���,0)) as ��� ,sum(nvl(A.���ݽ��,0)) ���ݽ��    " & _
             "  From    Ӧ����¼ a,ҩƷ��� C" & _
             "  Where  a.������� between [2] and [3] and �ƻ����� is null and (A.������� is  null or A.������� is not null and   " & _
             "      A.�������  in (Select ������� From �����¼ where (������� is null" & strTemp & ") and ��λID=[1]))" & _
             "      and A.��λID=[1] And A.��ĿID = C.ҩƷid And A.ϵͳ��ʶ = 1 and a.��¼���� <> 2 " & _
             "  group by a.ϵͳ��ʶ,a.��¼����,A.NO,A.��Ŀid,A.��� " & _
             "  having sum(nvl(A.��Ʊ���,0))<>0  "
             
    gstrSQL = gstrSQL & " UNION ALL " & _
             "  Select  no as ���,0 as ��־,a.NO,max(a.��ⵥ�ݺ�) ��ⵥ�ݺ�,            " & _
             "      to_char(max(A.�������),'yyyy-MM-dd') as ����,max(A.��Ʊ��) ��Ʊ��,to_char(max(A.��Ʊ����),'yyyy-MM-dd') as ��Ʊ����,            " & _
             "      null as �ƻ�����,max(a.�ƻ����) �ƻ����,null,max(decode(��¼״̬,3,A.ID,1,A.ID,0)) ID,             " & _
             "      max(A.Ʒ��) as ����,max(���) ���,max(����) ����,max(������λ) As ��λ,sum(nvl(����,0)) ����, sum(nvl(A.��Ʊ���,0)) as ��� ,sum(nvl(A.���ݽ��,0)) ���ݽ��    " & _
             "  From    Ӧ����¼ a " & _
             "  Where  a.������� between [2] and [3] and �ƻ����� is null and (A.������� is  null or A.������� is not null and   " & _
             "      A.�������  in (Select ������� From �����¼ where (������� is null" & strTemp & ") and ��λID=[1]))" & _
             "      and A.��λID=[1] And A.ϵͳ��ʶ <> 1 and a.��¼���� <> 2 " & _
             "  group by a.ϵͳ��ʶ,a.��¼����,A.NO,A.��Ŀid,A.��� " & _
             "  having sum(nvl(A.��Ʊ���,0))<>0 "

    gstrSQL = gstrSQL & " UNION ALL " & _
            " Select A.NO As ���, 0 As ��־, A.NO, Max(A.��ⵥ�ݺ�) ��ⵥ�ݺ�, To_Char(Max(A.�������), 'yyyy-MM-dd') As ����, " & _
            "       Max(A.��Ʊ��) ��Ʊ��, To_Char(Max(A.��Ʊ����), 'yyyy-MM-dd') As ��Ʊ����, Null As �ƻ�����, " & _
            "       -1 As �ƻ����, Null, Max(Decode(a.��¼״̬, 3, A.ID, 1, A.ID, 0)) ID, Max(A.Ʒ��) As ����, " & _
            "       Max(a.���) ���, Max(a.����) ����,Max(a.��λ) As ��λ, Max(Nvl(a.����, 0)) ����, " & _
            "       Max( Nvl(A.��Ʊ���, 0))-Sum( Nvl(B.�ƻ����, 0)) As ���, Max(Nvl(A.���ݽ��, 0)) ���ݽ�� " & _
            " From ( Select A.ID,A.No,A.��ⵥ�ݺ�,A.�������,A.��Ʊ��,A.��Ʊ����,A.Ʒ��,A.���,A.����," & strDSelect & "A.��Ʊ���," & _
            "               A.���ݽ��,A.��¼״̬,A.ϵͳ��ʶ,A.��¼����,A.��Ŀid,A.��� " & _
            "       From Ӧ����¼ A,ҩƷ��� C " & _
            "       Where A.������� Between [2] And [3] And �ƻ����� Is Not Null And A.��λid =[1]  And A.��ĿID = C.ҩƷid And A.ϵͳ��ʶ = 1" & _
            "       ) A,Ӧ����¼ B " & _
            " Where A.ID = B.ID " & _
            " Group By A.ϵͳ��ʶ, A.��¼����, A.NO, A.��Ŀid, A.��� " & _
            " Having Max(Nvl(A.��Ʊ���, 0)) - Sum(Nvl(B.�ƻ����, 0)) <> 0 "
            
    gstrSQL = gstrSQL & " UNION ALL " & _
            " Select A.NO As ���, 0 As ��־, A.NO, Max(A.��ⵥ�ݺ�) ��ⵥ�ݺ�, To_Char(Max(A.�������), 'yyyy-MM-dd') As ����, " & _
            "       Max(A.��Ʊ��) ��Ʊ��, To_Char(Max(A.��Ʊ����), 'yyyy-MM-dd') As ��Ʊ����, Null As �ƻ�����, " & _
            "       -1 As �ƻ����, Null, Max(Decode(a.��¼״̬, 3, A.ID, 1, A.ID, 0)) ID, Max(A.Ʒ��) As ����, " & _
            "       Max(a.���) ���, Max(a.����) ����, Max(a.������λ) As ��λ, Max(Nvl(a.����, 0)) ����, " & _
            "       Max( Nvl(A.��Ʊ���, 0))-Sum( Nvl(B.�ƻ����, 0)) As ���, Max(Nvl(A.���ݽ��, 0)) ���ݽ�� " & _
            " From ( Select ID,No,��ⵥ�ݺ�,�������,��Ʊ��,��Ʊ����,Ʒ��,���,����,������λ,����,��Ʊ���,���ݽ��,��¼״̬,ϵͳ��ʶ,��¼����,��Ŀid,��� " & _
            "       From Ӧ����¼ A " & _
            "       Where A.������� Between [2] And [3] And �ƻ����� Is Not Null And A.��λid =[1] And A.ϵͳ��ʶ <> 1" & _
            "       ) A,Ӧ����¼ B " & _
            " Where A.ID = B.ID " & _
            " Group By A.ϵͳ��ʶ, A.��¼����, A.NO, A.��Ŀid, A.��� " & _
            " Having Max(Nvl(A.��Ʊ���, 0)) - Sum(Nvl(B.�ƻ����, 0)) <> 0 "
      
    gstrSQL = gstrSQL & _
           "  UNION ALL  " & _
           "  Select  B.no as ���,1 as ��־,decode(b.no,null,a.no,b.no) as NO,a.��ⵥ�ݺ�,            " & _
           "      to_char(b.�������,'yyyy-MM-dd') as ����,A.��Ʊ��,to_char(A.��Ʊ����,'yyyy-MM-dd') as ��Ʊ����,            " & _
           "      to_char(a.�ƻ�����,'yyyy-mm-dd') as �ƻ�����,a.�ƻ����,null,A.ID,             " & _
           "      A.Ʒ�� as ����,a.���,a.����," & strDSelect & "nvl(A.�ƻ����,0) as ��� ,nvl(A.���ݽ��,0) ���ݽ��    " & _
           "  From   Ӧ����¼ a,ҩƷ��� C, " & _
           "          (Select * From Ӧ����¼   " & _
           "           WHERE ��λID=[1]  AND ������� between [2] and [3] and (��¼״̬=1 or mod(��¼״̬,3)=0) AND ��¼����<>-1 AND (������� IS NULL OR ������� IS NOT NULL AND  " & _
           "              ������� IN (Select ������� From �����¼  where (������� is null" & strTemp & ") and ��λID=[1]))  " & _
           "              AND �ƻ����� IS NOT  NULL  " & _
           "          ) b " & _
           "  Where  a.��¼����=-1  and (a.������� is  null or a.������� in (Select ������� from �����¼ where (������� is null" & strTemp & ") and ��λid=[1]))  and A.��λID=[1] AND a.ID=b.id " & _
           "     And A.��ĿID = C.ҩƷid And A.ϵͳ��ʶ = 1 "
    
    gstrSQL = gstrSQL & _
           "  UNION ALL  " & _
           "  Select  B.no as ���,1 as ��־,decode(b.no,null,a.no,b.no) as NO,a.��ⵥ�ݺ�,            " & _
           "      to_char(b.�������,'yyyy-MM-dd') as ����,A.��Ʊ��,to_char(A.��Ʊ����,'yyyy-MM-dd') as ��Ʊ����,            " & _
           "      to_char(a.�ƻ�����,'yyyy-mm-dd') as �ƻ�����,a.�ƻ����,null,A.ID,             " & _
           "      A.Ʒ�� as ����,a.���,a.����,a.������λ As ��λ,nvl(a.����,0), nvl(A.�ƻ����,0) as ��� ,nvl(A.���ݽ��,0) ���ݽ��    " & _
           "  From   Ӧ����¼ a, " & _
           "          (Select * From Ӧ����¼   " & _
           "           WHERE ��λID=[1]  AND ������� between [2] and [3] and (��¼״̬=1 or mod(��¼״̬,3)=0) AND ��¼����<>-1 AND (������� IS NULL OR ������� IS NOT NULL AND  " & _
           "              ������� IN (Select ������� From �����¼  where (������� is null" & strTemp & ") and ��λID=[1]))  " & _
           "              AND �ƻ����� IS NOT  NULL  " & _
           "          ) b " & _
           "  Where  a.��¼����=-1  and (a.������� is  null or a.������� in (Select ������� from �����¼ where (������� is null" & strTemp & ") and ��λid=[1]))  and A.��λID=[1] AND a.ID=b.id  And A.ϵͳ��ʶ <> 1 " & _
           "  ORDER BY ���,����,��־,�ƻ����"

    Err = 0: On Error GoTo ErrHand:
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngID, dtStartdate, dtEndDate)
    
    If rsTemp.RecordCount = 0 Then
        Exit Sub
    Else
        vsList.Rows = rsTemp.RecordCount + 2
    End If
    lngRow = 1
    With vsList
        .Redraw = False
        Do Until rsTemp.EOF
            '"^����|^���ݺ�|^Ʒ��|^���|^��λ|^����|^��ⵥ�ݺ�|^����|^���ݽ��|^���|^��Ʊ����|^�ƻ����|^�ƻ���������"
            .TextMatrix(lngRow, .ColIndex("����")) = Nvl(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("���ݺ�")) = Nvl(rsTemp!NO)
            .TextMatrix(lngRow, .ColIndex("Ʒ��")) = Nvl(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("���")) = Nvl(rsTemp!���)
            .TextMatrix(lngRow, .ColIndex("��λ")) = Nvl(rsTemp!��λ)
            .TextMatrix(lngRow, .ColIndex("����")) = Nvl(rsTemp!����)
            
            .TextMatrix(lngRow, .ColIndex("��ⵥ�ݺ�")) = Nvl(rsTemp!��ⵥ�ݺ�)
            
            .TextMatrix(lngRow, .ColIndex("����")) = Format(rsTemp("����"), gVbFmtString.FM_����)
            .TextMatrix(lngRow, .ColIndex("���ݽ��")) = Format(rsTemp("���ݽ��"), gVbFmtString.FM_���)
            .TextMatrix(lngRow, .ColIndex("���")) = Format(rsTemp("���"), gVbFmtString.FM_���)
                
            .TextMatrix(lngRow, .ColIndex("��Ʊ��")) = Nvl(rsTemp!��Ʊ��)
            .TextMatrix(lngRow, .ColIndex("��Ʊ����")) = Nvl(rsTemp!��Ʊ����)
            If Val(Nvl(rsTemp!�ƻ����)) = 0 Then
                .TextMatrix(lngRow, .ColIndex("�ƻ����")) = ""
            Else
                .TextMatrix(lngRow, .ColIndex("�ƻ����")) = IIf(Val(Nvl(rsTemp!�ƻ����)) = -1, "δ���Ƽƻ�", Nvl(rsTemp!�ƻ����))
            End If
            .TextMatrix(lngRow, .ColIndex("�ƻ���������")) = Nvl(rsTemp!�ƻ�����)
            dblSum = dblSum + Val(Nvl(rsTemp("���")))
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        If .Rows > 2 Then
            .TextMatrix(lngRow, .ColIndex("����")) = "�ϼ�"
            .TextMatrix(lngRow, .ColIndex("���")) = Format(dblSum, gVbFmtString.FM_���)
        End If
        If .Rows - 1 >= 2 Then .Row = 1
        .Col = 0: .LeftCol = 0
        
        .Redraw = True
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
    
    If Me.WindowState <> vbMaximized Then
        If Me.Height < 5000 Then
            Me.Height = 5000
        End If
        If Me.Width < 4500 Then
            Me.Width = 4500
        End If
    End If
    
    cbrTool.Move 0, 0, Me.ScaleWidth
    
    If lblHsc_s.Left > Me.ScaleWidth - 2000 Then lblHsc_s.Left = Me.ScaleWidth - 2000
    
    lblHsc_s.Top = IIf(cbrTool.Visible, cbrTool.Height + 30, 0)
    lblHsc_s.Height = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0) - lblHsc_s.Top - 15
    tvwList.Move 0, lblHsc_s.Top, lblHsc_s.Left, lblHsc_s.Height
    With fraSearch
        .Left = tvwList.Left
        .Top = tvwList.Top
        .Height = tvwList.Height
        .Width = tvwList.Width
    End With
    With fraSplit
        .Left = 0
        .Width = fraSearch.Width
    End With
    With PicSearchBack
        .Left = 0
        .Width = fraSearch.Width - 10
        .Height = IIf(fraSearch.Height - .Top < 0, 0, fraSearch.Height - .Top)
    End With
    Dim sngTmp As Single
    With PicClose
        sngTmp = PicSearchBack.Left + PicSearchBack.Width - .Width - 50
        .Left = sngTmp
    End With
    
    lblVsc_s.Left = lblHsc_s.Left + lblHsc_s.Width
    lblVsc_s.Width = Me.ScaleWidth - lblVsc_s.Left
    
    If lblVsc_s.Top > Me.ScaleHeight - 2000 Then lblVsc_s.Top = Me.ScaleHeight - 2000
    
    lblTemp(0).Move lblVsc_s.Left, lblHsc_s.Top, lblVsc_s.Width
    lblTemp(2).Move lblVsc_s.Left + (lblVsc_s.Width - lblTemp(2).Width) / 2, lblTemp(0).Top + (lblTemp(0).Height - lblTemp(2).Height) / 2
    vsHead.Move lblVsc_s.Left, lblTemp(0).Top + lblTemp(0).Height + 15, lblTemp(0).Width, lblVsc_s.Top - (lblTemp(0).Top + lblTemp(0).Height + 15)
    
    tabSelect.Move lblVsc_s.Left, lblVsc_s.Top + lblVsc_s.Height
    '����26224 by lesfeng 2010-02-08
    chk��������.Top = tabSelect.Top
    chk��������.Left = tabSelect.Left + tabSelect.Width + 200
    
    picFind.Move Me.ScaleWidth - picFind.Width - 5, tabSelect.Top
    If picFind.Left < lblVsc_s.Left Then picFind.Left = lblVsc_s.Left
       
    vsList.Move lblVsc_s.Left, tabSelect.Top + tabSelect.Height + 15, lblVsc_s.Width
    vsList.Height = Me.ScaleHeight - vsList.Top - IIf(stbThis.Visible, stbThis.Height, 0) - 30
     
    Call picFind_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
    zlDatabase.SetPara "�ϴ�ѡ��λID", mlng��λID, glngSys, mlngModule
    '����26224 by lesfeng 2010-02-08
    zlDatabase.SetPara "����ʱ���֮�󸶿�", IIf(chk��������.Value, 1, 0), glngSys, mlngModule
    zl_vsGrid_Para_Save mlngModule, vsHead, Me.Caption, "�����Ϣ�б�", True
    Select Case tabSelect.SelectedItem.Index
        Case 1
            zl_vsGrid_Para_Save mlngModule, vsList, Me.Caption, "������ϸ�б�", True
        Case 2
            zl_vsGrid_Para_Save mlngModule, vsList, Me.Caption, "�Ѹ���ϸ�б�", True
        Case 3
            zl_vsGrid_Para_Save mlngModule, vsList, Me.Caption, "δ����ϸ�б�", True
    End Select
End Sub

Private Sub mnuFileExcel_Click()
    subPrint 3
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFilePreView_Click()
    subPrint 2
End Sub

Private Sub mnuFilePrint_Click()
    subPrint 1
End Sub

Private Sub mnuFilePrintSet_Click()
    zlPrintSet
End Sub

Private Sub lblHsc_s_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    msngDownX = X
End Sub

Private Sub lblHsc_s_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        With lblHsc_s
            If .Left + X - msngDownX < 2000 Then Exit Sub
            If .Left + X - msngDownX > ScaleWidth - 2000 Then Exit Sub
            .Left = .Left + X - msngDownX
        End With
        Call Form_Resize
        
    End If
End Sub

Private Sub lblVsc_s_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    msngDownY = Y
End Sub

Private Sub lblVsc_s_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        With lblVsc_s
            If .Top + Y - msngDownY < 2000 Then Exit Sub
            If .Top + Y - msngDownY > ScaleHeight - 2000 Then Exit Sub
            .Top = .Top + Y - msngDownY
        End With
        Call Form_Resize
    End If
End Sub

Private Sub mnuViewFilter_Click()
    mnuViewFilter.Checked = Not mnuViewFilter.Checked
    
    If Not mnuViewFilter.Checked Then
        tlbThis.Buttons("Filter").Value = tbrUnpressed
        Me.fraSearch.Visible = False
    Else
        tlbThis.Buttons("Filter").Value = tbrPressed
        Me.fraSearch.Visible = True
        fraSearch.ZOrder
        Me.txt����.SetFocus
    End If
End Sub

Private Sub mnuViewFind_Click()
    '���ݶ�λ
    '����Ӧ���뵥�ݺŶ�λ
    Dim str���ݺ� As String, str��Ӧ��ID As String
    Dim rsTemp As New ADODB.Recordset
    Dim nod As MSComctlLib.Node, lngRow As Long, lngCol As Long
    
    If frmӦ���λ.Get��λ����(mstrPrivs, str���ݺ�, str��Ӧ��ID) = False Then
        Exit Sub
    End If
    
    If str���ݺ� <> "" Then
        '���ݵ��ݺ��ҵ���Ӧ��
        On Error GoTo errHandle
        gstrSQL = "select ��λID from Ӧ����¼ where ��ⵥ�ݺ�=[1] and ��λID is not null"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, str���ݺ�)
        
        If rsTemp.EOF = True Then
            MsgBox "���ݺ�Ϊ " & str���ݺ� & " �ļ�¼û���ҵ���", vbInformation, gstrSysName
            Exit Sub
        End If
        
        str��Ӧ��ID = rsTemp("��λID")
        rsTemp.Close
    End If
    
    On Error Resume Next
    Set nod = tvwList.Nodes("K" & str��Ӧ��ID)
    If Err <> 0 Then
        MsgBox "û�з���ָ����Ӧ�̣������Ѿ���ͣ�á�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    nod.Selected = True
    nod.EnsureVisible
    Call FullCount
    
    If str���ݺ� <> "" Then
        '�ҵ�����������
        If tabSelect.SelectedItem.Index = 1 Then
            lngCol = 1
        Else
            lngCol = 0
        End If
        
        With vsList
            For lngRow = .FixedRows To .Rows - 1
                If .TextMatrix(lngRow, lngCol) = str���ݺ� Then
                    .TopRow = lngRow
                    Exit For
                End If
            Next
        End With
    End If
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub mnuViewOpen_Click()
    If frmTimeSet.GetTimeScope(mdtBegin, mdtEnd, mstrData, Me) = True Then
        mstrKey = ""
        Call FullCount
    End If
End Sub

Private Sub mnuViewRefresh_Click()
    FullDept
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    cbrTool.Visible = mnuViewToolButton.Checked
    cbrTool.Bands(1).MinHeight = tlbThis.Height
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim buttTemp As Button
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For Each buttTemp In tlbThis.Buttons
        If mnuViewToolText.Checked Then
            buttTemp.Caption = buttTemp.Tag
        Else
            buttTemp.Caption = ""
        End If
    Next
    cbrTool.Bands(1).MinHeight = tlbThis.Height
    Form_Resize
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = mnuViewStatus.Checked
    Form_Resize
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpTitle_Click()
       ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(hwnd)
End Sub

Private Sub mnuViewUnit_Click(Index As Integer)
    mnuViewUnit(Index).Checked = Not mnuViewUnit(Index).Checked
    FullDept
End Sub

Private Sub picFind_Resize()
    Err = 0: On Error Resume Next
    With picFind
        txtFind.Left = lblFind.Width + lblFind.Left + 10
        txtFind.Width = .ScaleWidth - txtFind.Left
    End With
End Sub

Private Sub vsHead_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    zl_VsGridRowChange vsHead, OldRow, NewRow, OldCol, NewCol
End Sub

Private Sub vsHead_EnterCell()
    lblName = vsHead.TextMatrix(vsHead.Row, 0)
    lblName.Left = lblTemp(1).Left + (lblTemp(1).Width - lblName.Width) / 2

    Select Case tabSelect.SelectedItem.Index
        Case 1
            zl_vsGrid_Para_Save mlngModule, vsList, Me.Caption, "������ϸ�б�"
            InitColHead 1
            Full������ϸ
        Case 2
            zl_vsGrid_Para_Save mlngModule, vsList, Me.Caption, "�Ѹ���ϸ�б�"
            InitColHead 2
            Full�Ѹ��嵥
        Case 3
             zl_vsGrid_Para_Save mlngModule, vsList, Me.Caption, "δ����ϸ�б�"
             InitColHead 3
            Fullδ���嵥
    End Select
    '����26224 by lesfeng 2010-02-08
    mintFlag = 0
End Sub

Private Sub picClear_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ClearSearchData
    RaisEffect picClear, 0, "���", mRightAgnmt
    picClear.Tag = ""
    img���.Tag = ""
    ReleaseCapture
End Sub

Private Sub tabSelect_Click()
    '����26224 by lesfeng 2010-02-08
    If tabSelect.SelectedItem.Index = mintOldSel And mintFlag = 0 Then Exit Sub
    mintFlag = 0
    '������ʷ�ṹ
    Select Case mintOldSel
        Case 1
            zl_vsGrid_Para_Save mlngModule, vsList, Me.Caption, "������ϸ�б�", True
        Case 2
            zl_vsGrid_Para_Save mlngModule, vsList, Me.Caption, "�Ѹ���ϸ�б�", True
        Case 3
            zl_vsGrid_Para_Save mlngModule, vsList, Me.Caption, "δ����ϸ�б�", True
    End Select
    
    '�ָ���ʷ�ṹ
    vsList.Cols = 1
    Select Case tabSelect.SelectedItem.Index
        Case 1
            InitColHead 1
            Full������ϸ
        Case 2
            InitColHead 2
            Full�Ѹ��嵥
        Case 3
            InitColHead 3
            Fullδ���嵥
    End Select
    mintOldSel = tabSelect.SelectedItem.Index
    lblFind.Caption = "��" & vsList.ColKey(vsList.Col) & "����"
    lblFind.Tag = vsList.ColKey(vsList.Col)
    txtFind.Text = ""
    Call picFind_Resize
End Sub

Private Sub tlbthis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Find"
            mnuViewFind_Click
        Case "Search"
            mnuViewOpen_Click
        Case "PrintView"
            mnuFilePreView_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Filter"
            mnuViewFilter_Click
        Case "Refresh"
            Call mnuViewRefresh_Click
        Case "Help"
            Call mnuHelpTitle_Click
        Case "Exit"
            Call mnuFileExit_Click
    End Select
End Sub

Private Sub tlbthis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then Me.PopupMenu mnuViewTool
End Sub

Private Sub tvwList_NodeClick(ByVal Node As MSComctlLib.Node)
    If mlng��λID = Val(Mid(Node.Key, 2)) Then Exit Sub
    mlng��λID = Val(Mid(Node.Key, 2))
    FullCount
End Sub

Private Sub FillSum()
'����:װ�����ͳ������
    Dim rsTemp As New ADODB.Recordset
    Dim strBegin As String, strEnd As String
    Dim dblSum(1 To 4) As Double
    Dim lngRow As Long
    Dim blnSum As Boolean        '�ϼƵ���ʾ
    Dim i As Long
    Dim str���� As String
    Dim lng�ϼ�id As Long
    
    str���� = ""
    For i = 1 To Len(mstrType)
        If Mid(mstrType, i, 1) = 1 Then
            str���� = str���� & " or substr(b.����," & i & ",1)=1"
        End If
    Next
    If str���� <> "" Then
        str���� = " And (" & Mid(str����, 4) & ") "
    End If

    Dim strȨ�� As String
    strȨ�� = " and  " & Get����Ȩ��(gstrPrivs)
    
    stbThis.Panels(2).Text = "ʱ�䷶Χ��" & Format(mdtBegin, "yyyy-MM-dd") & " �� " & Format(mdtEnd, "yyyy-MM-dd")

    If tvwList.SelectedItem Is Nothing Then Exit Sub
    If mnuViewFilter.Checked Then
        mstrKey = ""
    Else
        If mstrKey = tvwList.SelectedItem.Key Then Exit Sub
        mstrKey = tvwList.SelectedItem.Key
    End If
    '��ʼ��ѯ
    'by lesfeng 2009-12-2 �����Ż�
    strBegin = Format(mdtBegin, "yyyyMMdd")
    strEnd = Format(mdtEnd, "yyyyMMdd")
    If UCase(mstrKey) = "ROOT" Then
        lng�ϼ�id = 0
    Else
        lng�ϼ�id = Val(Mid(mstrKey, 2))
    End If
    MousePointer = 11
    '���ȵõ��Ӳ�ѯ��SQL���
    If mnuViewFilter.Checked Then
        '����:
        gstrSQL = mstrFiler & " and  " & Get����Ȩ��(gstrPrivs, "B.")
    Else
        If tvwList.SelectedItem.Image = "2" Then
            gstrSQL = " and A.��λID=" & Mid(mstrKey, 2)
        ElseIf tvwList.SelectedItem.Image = "1" Then
            gstrSQL = " and A.��λID in (select ID from ��Ӧ�� where 1=1 " & zl_��ȡվ������() & "  " & Replace(str����, "b.����", "����") & strȨ�� & " start with �ϼ�ID is null connect by prior id=�ϼ�ID )"
        Else
            gstrSQL = " and A.��λID in (select ID from ��Ӧ�� where 1=1 " & zl_��ȡվ������() & "  " & Replace(str����, "b.����", "����") & strȨ�� & " start with �ϼ�ID =[2] connect by prior id=�ϼ�ID  )"
        End If
    End If
    
    If Mid(mstrData, 1, 1) = "1" Then
        gstrSQL = gstrSQL & " And A.�ڳ�Ӧ��<>0 "
    End If
    If Mid(mstrData, 2, 1) = "1" Then
        gstrSQL = gstrSQL & " And A.�����޹�<>0 "
    End If
    If Mid(mstrData, 3, 1) = "1" Then
        gstrSQL = gstrSQL & " And A.����֧��<>0 "
    End If
    If Mid(mstrData, 4, 1) = "1" Then
        gstrSQL = gstrSQL & " And A.��ĩӦ��<>0 "
    End If
    
    '�ٵõ�������SQL���
    gstrSQL = "select '��'||B.����||'��'|| B.���� as ����,B.ID,A.�ڳ�Ӧ��,A.�����޹�,A.����֧��,A.��ĩӦ�� from " & _
            "(select ��λID,sum(���-�ڳ�Ӧ��+�ڳ�����) as �ڳ�Ӧ��,sum(�ڳ�Ӧ��-��ĩӦ��) as �����޹� " & _
            "            ,sum(�ڳ�����-��ĩ����) as ����֧��,sum(���-��ĩӦ��+��ĩ����) as ��ĩӦ�� " & _
            "from( " & _
            "select ��λID,nvl(���,0)  as �ڳ�����, " & _
            "    decode(sign(to_char(�������,'yyyymmdd')-'" & strEnd & "'),1,nvl(���,0),0) as ��ĩ����, " & _
            "    0 as �ڳ�Ӧ��,0 as ��ĩӦ��,0 as ��� from �����¼ " & _
            "    where �������>=[1] " & _
            "Union All " & _
            "select ��λID ��λID,0 as �ڳ�����,0 as ��ĩ����, " & _
            "    ��Ʊ��� as �ڳ�Ӧ��,decode(sign(to_char(�������,'yyyymmdd')-'" & strEnd & "'),1,nvl(��Ʊ���,0),0) as ��ĩӦ��,0 as ��� from Ӧ����¼ " & _
            "    where ��¼����<>-1 And �������>=[1] " & _
            "Union All " & _
            "select ��λID ��λID,0 as �ڳ�����,0 as ��ĩ����,0 as �ڳ�Ӧ��,0 as ��ĩӦ��,nvl(���,0) as ��� from Ӧ����� " & _
            "    where ����=1) " & _
            "group by ��λID)A,��Ӧ�� B " & _
            "where A.��λID=B.ID  " & str���� & gstrSQL
    On Error GoTo errHandle
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CDate(Format(mdtBegin, "yyyy-MM-dd")), lng�ϼ�id, mstrOthers(0), mstrOthers(1), _
                            mstrOthers(2), mstrOthers(3), mstrOthers(4), mstrOthers(5), mstrOthers(6), mstrOthers(7), mstrOthers(8), mstrOthers(9), _
                            mstrOthers(10), mstrOthers(11), mstrOthers(12), mstrOthers(13), mstrOthers(14), mstrOthers(15))
    
    initvsHeadHead
    vsHead.Redraw = False
    If rsTemp.RecordCount = 0 Then
        vsHead.Rows = 2
        vsHead.RowData(1) = 0
    Else
        If rsTemp.RecordCount = 1 Then
            'ֻ��һ�У��Ͳ���ʾ�ϼ���
            vsHead.Rows = 2
            blnSum = False
        Else
            vsHead.Rows = rsTemp.RecordCount + 2
            blnSum = True
        End If
    End If
    lngRow = 1
    With vsHead
        Do Until rsTemp.EOF
            .RowData(lngRow) = rsTemp("ID")
            .TextMatrix(lngRow, .ColIndex("��Ӧ������")) = Nvl(rsTemp!����)
            .Cell(flexcpData, lngRow, .ColIndex("��Ӧ������")) = Nvl(rsTemp!ID)
            .TextMatrix(lngRow, .ColIndex("�ڳ�Ӧ��")) = Format(Val(Nvl(rsTemp!�ڳ�Ӧ��)), gVbFmtString.FM_���)
            .TextMatrix(lngRow, .ColIndex("�����޹�")) = Format(Val(Nvl(rsTemp!�����޹�)), gVbFmtString.FM_���)
            .TextMatrix(lngRow, .ColIndex("����֧��")) = Format(Val(Nvl(rsTemp!����֧��)), gVbFmtString.FM_���)
            .TextMatrix(lngRow, .ColIndex("��ĩӦ��")) = Format(Val(Nvl(rsTemp!��ĩӦ��)), gVbFmtString.FM_���)
            If blnSum = True Then
                dblSum(1) = dblSum(1) + Nvl(rsTemp("�ڳ�Ӧ��"), 0)
                dblSum(2) = dblSum(2) + Nvl(rsTemp("�����޹�"), 0)
                dblSum(3) = dblSum(3) + Nvl(rsTemp("����֧��"), 0)
                dblSum(4) = dblSum(4) + Nvl(rsTemp("��ĩӦ��"), 0)
            End If
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        If blnSum = True Then
            .TextMatrix(lngRow, 0) = "  �ϼ�"
            .Cell(flexcpData, lngRow, .ColIndex("��Ӧ������")) = 0
            .TextMatrix(lngRow, .ColIndex("�ڳ�Ӧ��")) = Format(dblSum(1), gVbFmtString.FM_���)
            .TextMatrix(lngRow, .ColIndex("�����޹�")) = Format(dblSum(2), gVbFmtString.FM_���)
            .TextMatrix(lngRow, .ColIndex("����֧��")) = Format(dblSum(3), gVbFmtString.FM_���)
            .TextMatrix(lngRow, .ColIndex("��ĩӦ��")) = Format(dblSum(4), gVbFmtString.FM_���)
        End If
                
    End With
    vsHead.Redraw = True
    
    MousePointer = 0
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then
        Resume
    Else
        MousePointer = 0
    End If
End Sub

Private Sub subPrint(bytMode As Byte)
'����:���д�ӡ,Ԥ���������EXCEL
'����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    
    If vsHead Is ActiveControl Then
        Set objPrint.Body = vsHead
        objPrint.Title.Text = "Ӧ���������Ϣ"
        objRow.Add " "
        objRow.Add "��ѯʱ�䣺" & Format(mdtBegin, "yyyy-MM-dd") & " �� " & Format(mdtEnd, "yyyy-MM-dd")
        objPrint.UnderAppRows.Add objRow
        
        Set objRow = New zlTabAppRow
        objRow.Add "��ӡ�ˣ�" & UserInfo.����
        objRow.Add "��ӡʱ�䣺" & Format(zlDatabase.Currentdate, "yyyy-MM-dd")
        objPrint.BelowAppRows.Add objRow
    Else
        Set objPrint.Body = vsList
        objPrint.Title.Text = tabSelect.SelectedItem.Caption
        objRow.Add "��Ӧ�̣�" & Mid(lblName.Caption, InStr(lblName.Caption, "��") + 1)
        objRow.Add "��ѯʱ�䣺" & Format(mdtBegin, "yyyy-MM-dd") & " �� " & Format(mdtEnd, "yyyy-MM-dd")
        objPrint.UnderAppRows.Add objRow
        
        Set objRow = New zlTabAppRow
        objRow.Add "��ӡ�ˣ�" & UserInfo.����
        objRow.Add "��ӡʱ�䣺" & Format(zlDatabase.Currentdate, "yyyy-MM-dd")
        objPrint.BelowAppRows.Add objRow
    End If
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Private Sub PicSearchBack_Resize()
    Dim i As Long
    Dim CtlWidth As Single
    Dim sngBottom As Single
    Dim blnOther As Boolean
    
    If InStr(1, Me.lblHit.Caption, ">>") <> 0 Then
        sngBottom = Me.lblHit.Top + Me.lblHit.Height
    Else
        sngBottom = shpHit.Top + shpHit.Height
    End If
    
    With picSearch
        .Left = PicSearchBack.ScaleLeft
        .Top = PicSearchBack.ScaleTop
    End With
    
    If PicSearchBack.ScaleHeight < sngBottom Then
        Scr.Visible = True
        picSearch.Width = IIf(PicSearchBack.ScaleWidth - Me.Scr.Width < 0, 0, PicSearchBack.ScaleWidth - Me.Scr.Width)
    Else
        Scr.Visible = False
        picSearch.Width = PicSearchBack.ScaleWidth
    End If
    With Scr
        .Left = picSearch.Left + picSearch.Width
        .Top = picSearch.Top
        .Height = PicSearchBack.ScaleHeight
    End With
    shpHit.Width = IIf(picSearch.Width - 100 < 0, 0, picSearch.Width - 100)
    CtlWidth = IIf(shpHit.Width - 100 < 0, 0, shpHit.Width - 100)
    lblHit.Width = shpHit.Width
    For i = 0 To lblOther.UBound
        lblOther(i).Width = CtlWidth
        TxtOther(i).Width = CtlWidth
    Next
    For i = 0 To chkOther.UBound
        chkOther(i).Width = CtlWidth
    Next
    For i = 0 To lblDate.UBound
        lblDate(i).Width = CtlWidth
        DtpOther(i).Width = CtlWidth
    Next
    Scr.Max = Int(Me.picSearch.Height / Me.PicSearchBack.Height + 0.5) * 12
End Sub

Private Sub Scr_Change()
    Scr_Scroll
End Sub

Private Sub Scr_Scroll()
    picSearch.Top = -Scr.Value * (Me.PicSearchBack.Height / 12) + 400
End Sub


Private Sub lblHit_Click()
    Dim i As Long
    Dim blnTrue As Boolean
    
    If InStr(1, lblHit.Caption, "<<") <> 0 Then
        blnTrue = False
        lblHit.Caption = Replace(Me.lblHit.Caption, "<<", ">>")
        lblHit.BackStyle = 0
        lblHit.ForeColor = &H8000000D
        shpHit.Visible = False
    Else
        blnTrue = True
        lblHit.Caption = Replace(Me.lblHit.Caption, ">>", "<<")
        lblHit.BackStyle = 1
        lblHit.ForeColor = &H8000000E
        shpHit.Visible = True
    End If
    For i = 0 To lblOther.UBound
        lblOther(i).Visible = shpHit.Visible
        TxtOther(i).Visible = shpHit.Visible
    Next
    For i = 0 To chkOther.UBound
        chkOther(i).Visible = chkOther(i).Visible And shpHit.Visible
    Next
    For i = 0 To lblDate.UBound
        lblDate(i).Visible = shpHit.Visible
        DtpOther(i).Visible = shpHit.Visible
    Next
    PicSearchBack_Resize
End Sub

Private Function ClearSearchData()
    '------------------------------------------------------------------
    '����:����������������
    '------------------------------------------------------------------
    Dim i As Long
    initvsHeadHead True
    Call InitColHead(1, True)
    Call InitColHead(2, True)
    
    Me.txt����.Text = ""
    Me.Txt����.Text = ""
    For i = 0 To TxtOther.UBound
        TxtOther(i).Text = ""
    Next
    For i = 0 To chkOther.UBound
        chkOther(i).Value = 0
    Next
    For i = 0 To DtpOther.UBound
        DtpOther(i).Value = ""
    Next
End Function

Private Sub PicClose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaisEffect PicClose, 2, "��", mCenterAgnmt, True
End Sub

Private Sub PicClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If PicClose.Tag = "In" Then
        If X < 0 Or Y < 0 Or X > PicClose.Width Or Y > PicClose.Height Then
            PicClose.Tag = ""
            ReleaseCapture
            RaisEffect PicClose, 0, "��", mCenterAgnmt, True
        End If
    Else
        PicClose.Tag = "In"
        SetCapture PicClose.hwnd
        MousePointer = 99
        RaisEffect PicClose, 1, "��", mCenterAgnmt, True
    End If
End Sub

Private Sub PicClose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PicClose.Tag = ""
    RaisEffect PicClose, 0, "��", mCenterAgnmt, True
    ReleaseCapture
    mnuViewFilter_Click
    Call FullCount
End Sub

Private Sub picHelp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaisEffect picHelp, 2
End Sub

Private Sub picHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If picHelp.Tag = "In" Then
        If X < 0 Or Y < 0 Or X > picHelp.Width Or Y > picHelp.Height Then
            picHelp.Tag = ""
            ReleaseCapture
            RaisEffect picHelp, 0
            Set Me.imgHelp.Picture = iltHelp.ListImages("HELPB").Picture
        End If
    Else
        picHelp.Tag = "In"
        SetCapture picHelp.hwnd
        MousePointer = 99
        RaisEffect picHelp, 1
        Set Me.imgHelp.Picture = iltHelp.ListImages("HELPC").Picture      'LoadResPicture("HELPC", 0)
    End If
End Sub

Private Sub picHelp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
    RaisEffect picHelp, 0
    picHelp.Tag = ""
    imgHelp.Tag = ""
    ReleaseCapture
End Sub

Private Sub picClear_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaisEffect picClear, 2, "���", mRightAgnmt
End Sub

Private Sub picClear_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If picClear.Tag = "In" Then
        If X < 0 Or Y < 0 Or X > picClear.Width Or Y > picClear.Height Then
            picClear.Tag = ""
            ReleaseCapture
            RaisEffect picClear, 0, "���", mRightAgnmt
            Set Me.img���.Picture = iltHelp.ListImages("SEARCHB").Picture   'LoadResPicture("SEARCHB", 0)
        End If
    Else
        picClear.Tag = "In"
        SetCapture picClear.hwnd
        MousePointer = 99
        RaisEffect picClear, 1, "���", mRightAgnmt
        Set Me.img���.Picture = iltHelp.ListImages("SEARCHC").Picture ' LoadResPicture("SEARCHC", 0)
    End If
End Sub

Private Sub imgHelp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picHelp_MouseDown Button, Shift, X, Y
End Sub

Private Sub imgHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If imgHelp.Tag = "In" Then
        If X < 0 Or Y < 0 Or X > imgHelp.Width Or Y > imgHelp.Height Then
            imgHelp.Tag = ""
            ReleaseCapture
            RaisEffect picHelp, 0
            Set Me.imgHelp.Picture = iltHelp.ListImages("HELPB").Picture
        End If
    Else
        imgHelp.Tag = "In"
        SetCapture picHelp.hwnd
        MousePointer = 99
        RaisEffect picHelp, 1
        Set Me.imgHelp.Picture = iltHelp.ListImages("HELPC").Picture      'LoadResPicture("HELPC", 0)
    End If
End Sub

Private Sub imgHelp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picHelp_MouseUp Button, Shift, X, Y
End Sub

Private Function IsValitSearchCon() As Boolean
    '------------------------------------------------------
    '����:������������Ƿ���Ч
    '------------------------------------------------------
    Dim i As Long
    IsValitSearchCon = False
    If InStr(1, Me.txt����.Text, "'") <> 0 Then
        MsgBox "���������к��÷Ƿ��ַ���", vbInformation, gstrSysName
        Exit Function
    End If
    If InStr(1, Me.Txt����.Text, "'") <> 0 Then
        MsgBox "�����к��÷Ƿ��ַ���", vbInformation, gstrSysName
        Exit Function
    End If
    For i = 0 To TxtOther.UBound
        If InStr(1, Me.TxtOther(i).Text, "'") <> 0 Then
            MsgBox Me.TxtOther(i).Tag & "�к��÷Ƿ��ַ���", vbInformation, gstrSysName
            Exit Function
        End If
    Next
    IsValitSearchCon = True
End Function

Private Sub TxtOther_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
        ScrCtl TxtOther(Index)
    End If
End Sub

Private Sub TxtOther_LostFocus(Index As Integer)
    Dim strIme As String
    zlCommFun.OpenIme (False)
End Sub

Private Sub txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub Txt����_GotFocus()
    Dim strIme As String
    Txt����.SelStart = 0
    Txt����.SelLength = Len(Txt����)
    zlCommFun.OpenIme (True)
End Sub

Private Sub cmd����_Click()
    Dim strWhere As String
    If Not IsValitSearchCon Then Exit Sub
    strWhere = Trim(GetSearchCon)
    If strWhere = "" Then
        ShowMsgbox "δ�����������,�����룡"
        Exit Sub
    End If
    mstrFiler = " AND (" & strWhere & " )"
   '����
   Call FullCount
End Sub

Private Sub LoadOtherCon()
    '---------------------------------------------------------------------------------
    '����:������������ѡ����
    '----------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim strFind As String
    Dim CurTop As Single
    Dim CtlWidth As Single
    Dim CtlLeft As Single
    CtlWidth = IIf(lblHit.Width - 50 < 0, 0, lblHit.Width - 50)
    CtlLeft = lblHit.Left + 50
    Dim sngTabIndex As Long
    sngTabIndex = 4
    RaisEffect PicClose, 0, "��", mCenterAgnmt, True
    RaisEffect picClear, 0, "���", mRightAgnmt

    '�����ı�����
    For i = 0 To 9
        strFind = Switch(i = 0, "��ַ", i = 1, "���֤��", i = 2, "ִ�պ�", i = 3, "˰��ǼǺ�", i = 4, "�ʺ�", _
             i = 5, "��ϵ��", i = 6, "��������", i = 7, "����ί����", i = 8, "������֤��", i = 9, "ҩ��ֱ�����")
             
        If i <> 0 Then
            Load lblOther(i)
            Load TxtOther(i)
            lblOther(i).Top = CurTop
            CurTop = CurTop + lblOther(i).Height + 50
            TxtOther(i).Top = CurTop
            CurTop = CurTop + TxtOther(i).Height + 100
        Else
            CurTop = TxtOther(i).Top + TxtOther(i).Height + 100
        End If
        lblOther(i).TabIndex = sngTabIndex
        sngTabIndex = sngTabIndex + 1
        TxtOther(i).TabIndex = sngTabIndex
        sngTabIndex = sngTabIndex + 1
        lblOther(i) = strFind
        TxtOther(i).Tag = strFind
        lblOther(i).Left = CtlLeft
        TxtOther(i).Left = CtlLeft
        TxtOther(i).Width = CtlWidth
        lblOther(i).Width = CtlWidth
    Next
    
    '����ѡ������
'    For i = 0 To 2
'        strFind = Switch(i = 0, "�޾��Բ���", i = 1, "һ���Բ���", i = 2, "�������")
'        If i <> 0 Then
'            Load chkOther(i)
'        End If
'        chkOther(i).Top = CurTop
'        chkOther(i).TabIndex = sngTabIndex
'        sngTabIndex = sngTabIndex + 1
'        CurTop = CurTop + chkOther(i).Height + 100
'        chkOther(i).Caption = strFind
'        chkOther(i).Tag = strFind
'        chkOther(i).Left = CtlLeft
'        chkOther(i).Width = CtlWidth
'        If marblnSelectWare And (marblnOnlyRation Or marBln�޾��Բ���) Then
'            chkOther(i).Enabled = False
'        Else
'            chkOther(i).Enabled = True
'        End If
'    Next
    '����ʱ��ѡ��ؼ�
    For i = 0 To 3
        strFind = Switch(i = 0, "����ʱ��", i = 1, "����ʱ��", i = 2, "���֤Ч��", i = 3, "ִ��Ч��")
        If i <> 0 Then
            Load lblDate(i)
            Load DtpOther(i)
        End If
        lblDate(i).TabIndex = sngTabIndex
        sngTabIndex = sngTabIndex + 1
        DtpOther(i).TabIndex = sngTabIndex
        sngTabIndex = sngTabIndex + 1
        lblDate(i) = strFind
        DtpOther(i).Tag = strFind
        lblDate(i).Left = CtlLeft
        DtpOther(i).Left = CtlLeft
        DtpOther(i).Width = CtlWidth
        lblDate(i).Width = CtlWidth
        lblDate(i).Top = CurTop
        CurTop = CurTop + lblDate(i).Height + 50
        DtpOther(i).Top = CurTop
        If i < 2 Then
            DtpOther(i).MaxDate = zlDatabase.Currentdate()
        Else
            DtpOther(i).MaxDate = CDate("3000-01-01")
        End If
        DtpOther(i).Value = zlDatabase.Currentdate()
        CurTop = CurTop + DtpOther(i).Height + 100
        DtpOther(i).Value = Null
        DtpOther(i).Enabled = True
    Next
    picSearch.Height = CurTop
    shpHit.Height = IIf(CurTop - shpHit.Top < 0, 0, CurTop - shpHit.Top)
    chkOther(0).Visible = False
End Sub

Private Function GetSearchCon() As String
    '---------------------------------------------------------------------------------------------------------
    '����:��ȡ��ѯ����
    '---------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim strWhere As String
    Dim strTemp As String
    Dim strField As String
    Dim LfPBF As String
    Dim RgPbf As String
    Dim strOthers(0 To 16) As String
    
    If gstrMatchMethod = "0" Then
        LfPBF = "%"
        RgPbf = "%"
    Else
        LfPBF = ""
        RgPbf = "%"
    End If
    
    strWhere = ""
    strTemp = Trim(txt����.Text)
    If strTemp <> "" Then
        If InStr(1, strTemp, "%") <> 0 Then
            strWhere = strWhere & "   or  (B.���� Like [3]) "
            strOthers(0) = strTemp
        Else
            strWhere = strWhere & "   or  (B.���� Like [3]) "
            strOthers(0) = LfPBF & strTemp & RgPbf
        End If
    End If
    
    strTemp = UCase(Trim(Txt����.Text))
    If strTemp <> "" Then
        If InStr(1, strTemp, "%") <> 0 Then
            strWhere = strWhere & "   or  (B.���� Like [4])  "
            strWhere = strWhere & "   or  (B.���� Like [4])  "
            strOthers(1) = strTemp
        Else
            strWhere = strWhere & "   or  (B.���� Like [4]) "
            strWhere = strWhere & "   or  (B.���� Like [4]) "
            strOthers(1) = LfPBF & strTemp & RgPbf
        End If
    End If
    
    If shpHit.Visible Then
        For i = 0 To TxtOther.UBound
            strField = " upper(B." & TxtOther(i).Tag & ")"
            strTemp = UCase(Trim(TxtOther(i).Text))
            If strTemp <> "" Then
                If InStr(1, strTemp, "%") <> 0 Then
                    strWhere = strWhere & "   or  (" & strField & "  Like [" & i + 5 & "]) "
                    strOthers(i + 2) = strTemp
                Else
                    strWhere = strWhere & "   or  ( " & strField & "  Like [" & i + 5 & "]) "
                    strOthers(i + 2) = LfPBF & strTemp & RgPbf
                End If
            End If
        Next
        For i = 0 To DtpOther.UBound
            strField = DtpOther(i).Tag
            If Not IsNull(DtpOther(i).Value) Then
'                strWhere = strWhere & "   or  ( to_char(B." & strField & ",'yyyy-MM-DD')  = '" & Format(DtpOther(i).Value, "yyyy-MM-DD") & "' ) "
                strWhere = strWhere & "   or  ( to_char(B." & strField & ",'yyyy-MM-DD')  = [" & i + 15 & "] ) "
                strOthers(i + 12) = Format(DtpOther(i).Value, "yyyy-MM-DD")
            End If
        Next
    End If
    mstrOthers = strOthers
    strWhere = Mid(strWhere, 6)
    GetSearchCon = strWhere
End Function

Private Sub Txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub ScrCtl(ByVal ctlObject As Object)
'    Err = 0: On Error Resume Next
'    If (ctlObject.Top + ctlObject.Height) + picSearch.Top + 1800 > PicSearchBack.ScaleHeight Then
'            If Scr.Value + 1 < Scr.Max Then
'                 Scr.Value = Scr.Value + 3
'            End If
'    End If
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

Private Sub vsHead_GotFocus()
    zl_VsGridGotFocus vsHead
End Sub

Private Sub vsHead_LostFocus()
    zl_VsGridLOSTFOCUS vsHead
End Sub

Private Sub vsList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    lblFind.Caption = "��" & vsList.ColKey(NewCol) & "����"
    If lblFind.Tag <> vsList.ColKey(NewCol) Then
        txtFind.Text = ""
    End If
    lblFind.Tag = vsList.ColKey(NewCol)
    
    Call picFind_Resize
    zl_VsGridRowChange vsList, OldRow, NewRow, OldCol, NewCol
End Sub

Private Sub vsList_GotFocus()
    zl_VsGridGotFocus vsList
End Sub

Private Sub vsList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyF3 Then Exit Sub
    If Trim(txtFind.Text) = "" Then Exit Sub
    FindRow Trim(txtFind.Text), IIf(vsList.Row + 1 >= vsList.Rows - 1, 1, vsList.Row + 1)
End Sub

Private Sub vsList_LostFocus()
    zl_VsGridLOSTFOCUS vsList
End Sub

Private Sub FindRow(ByVal strFind As String, Optional lngRow As Long = 1)
    '����:����ָ�е������Ƿ�������ص�����
    '����:intMachType:0-��ƥ��,1-��ȫƥ��
    Dim i As Long, lngCol As Long
    Dim blnAll As Boolean
Redo:
    With vsList
        lngCol = .ColIndex(lblFind.Tag)
        'δ�ҵ����˳�
        If lngCol < 0 Then Exit Sub
        
        If InStr(1, lblFind.Tag, "����") > 0 Then
            blnAll = True
            strFind = Format(Val(strFind), gVbFmtString.FM_����)
        ElseIf InStr(1, lblFind.Tag, "���") > 0 Then
            blnAll = True
            strFind = Format(Val(strFind), gVbFmtString.FM_���)
        ElseIf InStr(1, lblFind.Tag, "�ɹ���") > 0 Then
            blnAll = True
            strFind = Format(Val(strFind), gVbFmtString.FM_�ɱ���)
        ElseIf InStr(1, lblFind.Tag, "���ۼ�") > 0 Then
            blnAll = True
            strFind = Format(Val(strFind), gVbFmtString.FM_���ۼ�)
        ElseIf InStr(1, lblFind.Tag, "��") > 0 Then
            blnAll = True
            strFind = Format(Val(strFind), gVbFmtString.FM_���)
        ElseIf InStr(1, lblFind.Tag, "����") > 0 Then
            blnAll = False
            strFind = CheckIsDate(strFind, lblFind.Tag)
            If strFind = "" Then Exit Sub
        Else
            blnAll = False
        End If
       i = .FindRow(strFind, lngRow, lngCol, False, blnAll)
       If i > 0 Then
            .Row = i: .TopRow = i
       Else
            If lngRow = 1 Then
                ShowMsgbox "�Ѿ��鵽ĩβ,û�з�����������������,����"
            Else
                If MsgBox("�Ѿ��鵽ĩβ,û�з�����������������,�Ƿ����½��в���!", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                    lngRow = 1
                    GoTo Redo:
                End If
            End If
       End If
    End With
End Sub

Private Sub txtFind_GotFocus()
    zlControl.TxtSelAll txtFind
    If InStr(1, lblFind.Tag, "��") > 0 Or _
        InStr(1, lblFind.Tag, "����") > 0 Or _
        InStr(1, lblFind.Tag, "��") > 0 Then
        zlCommFun.OpenIme False
    Else
        zlCommFun.OpenIme True
    End If
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strNO As String
    If KeyCode <> vbKeyReturn Then
        If KeyCode = vbKeyF3 Then
            Call vsList_KeyDown(vbKeyF3, 0)
            Exit Sub
        End If
        Exit Sub
    End If
    If Trim(txtFind) = "" Then Exit Sub
    FindRow Trim(txtFind.Text)
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
        
    If InStr(1, lblFind.Tag, "����") > 0 Then
        zlControl.TxtCheckKeyPress txtFind, KeyAscii, m�����ʽ
    ElseIf InStr(1, lblFind.Tag, "���") > 0 Then
        zlControl.TxtCheckKeyPress txtFind, KeyAscii, m�����ʽ
    ElseIf InStr(1, lblFind.Tag, "�ɹ���") > 0 Then
        zlControl.TxtCheckKeyPress txtFind, KeyAscii, m�����ʽ
    ElseIf InStr(1, lblFind.Tag, "���ۼ�") > 0 Then
        zlControl.TxtCheckKeyPress txtFind, KeyAscii, m�����ʽ
    ElseIf InStr(1, lblFind.Tag, "��") > 0 Then
        zlControl.TxtCheckKeyPress txtFind, KeyAscii, m�����ʽ
    ElseIf InStr(1, lblFind.Tag, "����") > 0 Then
        zlControl.TxtCheckKeyPress txtFind, KeyAscii, m�ı�ʽ
    Else
        zlControl.TxtCheckKeyPress txtFind, KeyAscii, m�ı�ʽ
    End If
        
End Sub

