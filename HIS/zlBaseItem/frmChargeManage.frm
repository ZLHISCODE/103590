VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "CODEJO~1.OCX"
Begin VB.Form frmChargeManage 
   Caption         =   "�շ���Ŀ����"
   ClientHeight    =   7890
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   11250
   Icon            =   "frmChargeManage.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7890
   ScaleWidth      =   11250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picPage 
      BorderStyle     =   0  'None
      Height          =   1620
      Index           =   1
      Left            =   5490
      ScaleHeight     =   1620
      ScaleWidth      =   960
      TabIndex        =   28
      Top             =   2520
      Width           =   960
      Begin VSFlex8Ctl.VSFlexGrid msh��Ŀ 
         Height          =   1695
         Left            =   60
         TabIndex        =   30
         Top             =   570
         Width           =   3465
         _cx             =   6112
         _cy             =   2990
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
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
      Begin VB.CheckBox chk�۸� 
         Caption         =   "��ʾ��ʷ�۸�"
         Height          =   315
         Left            =   4770
         TabIndex        =   29
         Top             =   60
         Width           =   1425
      End
      Begin VB.Image img��Ŀ 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   540
         Left            =   0
         Picture         =   "frmChargeManage.frx":0442
         Top             =   0
         Width           =   540
      End
      Begin VB.Label lbl��Ŀ 
         Caption         =   "    �˴���ʾ�շ���Ŀ�ļ۸񣬱�����ɫΪ��ɫ���Ǽ����ǵ�ǰ�۸�"
         Height          =   435
         Left            =   780
         TabIndex        =   31
         Top             =   60
         Width           =   3795
      End
   End
   Begin VB.PictureBox picͣ��ԭ�� 
      BorderStyle     =   0  'None
      Height          =   220
      Left            =   150
      ScaleHeight     =   225
      ScaleWidth      =   3945
      TabIndex        =   42
      Top             =   5580
      Width           =   3945
      Begin VB.Label lblͣ��ԭ�� 
         Caption         =   "ͣ��ԭ��"
         Height          =   225
         Left            =   0
         TabIndex        =   43
         Top             =   0
         Width           =   2955
      End
   End
   Begin VB.PictureBox picPage 
      Height          =   1815
      Index           =   6
      Left            =   8700
      ScaleHeight     =   1755
      ScaleWidth      =   810
      TabIndex        =   11
      Top             =   4860
      Width           =   870
      Begin VSFlex8Ctl.VSFlexGrid vsWholeSet 
         Height          =   4680
         Left            =   0
         TabIndex        =   38
         Top             =   225
         Width           =   11355
         _cx             =   20029
         _cy             =   8255
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483634
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   16
         FixedRows       =   1
         FixedCols       =   2
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmChargeManage.frx":067C
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
         ExplorerBar     =   2
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
   Begin MSComctlLib.ImageList iltdept 
      Left            =   3180
      Top             =   3345
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
            Picture         =   "frmChargeManage.frx":08AA
            Key             =   "Dept"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picPage 
      Height          =   1815
      Index           =   7
      Left            =   9870
      ScaleHeight     =   1755
      ScaleWidth      =   810
      TabIndex        =   32
      Top             =   4860
      Width           =   870
      Begin MSComctlLib.ListView lvwUseDept 
         Height          =   1230
         Left            =   45
         TabIndex        =   36
         Top             =   165
         Width           =   4605
         _ExtentX        =   8123
         _ExtentY        =   2170
         View            =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "iltdept"
         SmallIcons      =   "iltdept"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "����"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.PictureBox picPage 
      BorderStyle     =   0  'None
      Height          =   1815
      Index           =   5
      Left            =   7500
      ScaleHeight     =   1815
      ScaleWidth      =   870
      TabIndex        =   33
      Top             =   4860
      Width           =   870
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh�ѱ� 
         Height          =   2475
         Left            =   0
         TabIndex        =   34
         Top             =   240
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   4366
         _Version        =   393216
         FixedCols       =   0
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483630
         FocusRect       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.PictureBox picPage 
      BorderStyle     =   0  'None
      Height          =   1650
      Index           =   3
      Left            =   7650
      ScaleHeight     =   1650
      ScaleWidth      =   900
      TabIndex        =   14
      Top             =   2520
      Width           =   900
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh���� 
         Height          =   1455
         Left            =   0
         TabIndex        =   15
         Top             =   810
         Width           =   3105
         _ExtentX        =   5477
         _ExtentY        =   2566
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         FocusRect       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
      End
      Begin VB.Image img���� 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   480
         Left            =   0
         Picture         =   "frmChargeManage.frx":0BC6
         Top             =   0
         Width           =   480
      End
      Begin VB.Label lbl���� 
         Caption         =   "    ������Ŀ��ָ�û��ڽ��е���¼���У����������շ���Ŀ�����Ӷ��Զ����ӵ��շ���Ŀ��"
         Height          =   435
         Left            =   870
         TabIndex        =   16
         Top             =   90
         Width           =   3795
      End
   End
   Begin VB.PictureBox picPage 
      BorderStyle     =   0  'None
      Height          =   1650
      Index           =   4
      Left            =   8730
      ScaleHeight     =   1650
      ScaleWidth      =   900
      TabIndex        =   12
      Top             =   2520
      Width           =   900
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshAlias 
         Height          =   2475
         Left            =   120
         TabIndex        =   13
         Top             =   0
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   4366
         _Version        =   393216
         FixedCols       =   0
         BackColorSel    =   -2147483639
         ForeColorSel    =   -2147483630
         FocusRect       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.PictureBox picTreeClass_S 
      BorderStyle     =   0  'None
      Height          =   1125
      Left            =   120
      ScaleHeight     =   1125
      ScaleWidth      =   2700
      TabIndex        =   9
      Top             =   4380
      Width           =   2700
      Begin XtremeSuiteControls.TabControl tbClassPage 
         Height          =   2700
         Left            =   -270
         TabIndex        =   10
         Top             =   -495
         Width           =   2175
         _Version        =   589884
         _ExtentX        =   3836
         _ExtentY        =   4762
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picTreeWholeSet 
      BorderStyle     =   0  'None
      Height          =   1605
      Left            =   60
      ScaleHeight     =   1605
      ScaleWidth      =   2700
      TabIndex        =   7
      Top             =   2520
      Width           =   2700
      Begin MSComctlLib.TreeView tvwWholeSet 
         Height          =   1485
         Left            =   0
         TabIndex        =   8
         Top             =   240
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   2619
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   494
         LabelEdit       =   1
         LineStyle       =   1
         Sorted          =   -1  'True
         Style           =   7
         ImageList       =   "ils16"
         Appearance      =   1
      End
   End
   Begin VB.PictureBox picTreeItem 
      BorderStyle     =   0  'None
      Height          =   1560
      Left            =   90
      ScaleHeight     =   1560
      ScaleWidth      =   2745
      TabIndex        =   5
      Top             =   825
      Width           =   2745
      Begin MSComctlLib.TreeView tvwMainItem 
         Height          =   1485
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   2619
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   494
         LabelEdit       =   1
         LineStyle       =   1
         Sorted          =   -1  'True
         Style           =   7
         ImageList       =   "ils16"
         Appearance      =   1
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   7530
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   635
      SimpleText      =   $"frmChargeManage.frx":1008
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmChargeManage.frx":104F
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14764
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
   Begin VB.PictureBox picNS 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   4050
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   3000
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2040
      Width           =   3000
   End
   Begin MSComctlLib.ListView lvwMain_S 
      Height          =   1065
      Left            =   8040
      TabIndex        =   2
      Top             =   810
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   1879
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ils32"
      SmallIcons      =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList Ilsrw 
      Left            =   3645
      Top             =   1710
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":18E3
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":1B03
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":1D23
            Key             =   "Parent"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":1F3F
            Key             =   "Child"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":215B
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":237B
            Key             =   "Raise"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":2597
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":27B7
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":29D7
            Key             =   "Start"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":2BF7
            Key             =   "View"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":2E17
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":3037
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":3257
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":3571
            Key             =   "verify"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList Ilscolor 
      Left            =   2925
      Top             =   1665
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":378B
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":39AB
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":3BCB
            Key             =   "Parent"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":3DE7
            Key             =   "Child"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":4003
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":4223
            Key             =   "Raise"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":443F
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":465F
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":487F
            Key             =   "Start"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":4A9F
            Key             =   "View"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":4CBF
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":4EDF
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":50FF
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":5419
            Key             =   "verify"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picSplit 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3225
      Left            =   4500
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3225
      ScaleWidth      =   45
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1920
      Width           =   45
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   1376
      BandCount       =   2
      _CBWidth        =   11250
      _CBHeight       =   780
      _Version        =   "6.7.8988"
      Child1          =   "Toolbar1"
      MinHeight1      =   720
      Width1          =   9795
      Key1            =   "only"
      NewRow1         =   0   'False
      Caption2        =   "����"
      Child2          =   "txtFind"
      MinHeight2      =   300
      Width2          =   1080
      NewRow2         =   0   'False
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   10410
         TabIndex        =   41
         Top             =   240
         Width           =   750
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   720
         Left            =   165
         TabIndex        =   40
         Top             =   30
         Width           =   9600
         _ExtentX        =   16933
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "Ilsrw"
         HotImageList    =   "Ilscolor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   19
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ԥ��"
               Key             =   "Preview"
               Object.ToolTipText     =   "Ԥ��"
               Object.Tag             =   "Ԥ��"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "Print"
               Object.ToolTipText     =   "��ӡ"
               Object.Tag             =   "��ӡ"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split0"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Parent"
               Object.ToolTipText     =   "���ӷ���"
               Object.Tag             =   "����"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��Ŀ"
               Key             =   "Child"
               Object.ToolTipText     =   "������Ŀ"
               Object.Tag             =   "��Ŀ"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�޸�"
               Key             =   "Modify"
               Object.ToolTipText     =   "�޸�"
               Object.Tag             =   "�޸�"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ɾ��"
               Key             =   "Delete"
               Object.ToolTipText     =   "ɾ��"
               Object.Tag             =   "ɾ��"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split1"
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Raise"
               Description     =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "���"
               Key             =   "RaiseVerify"
               Object.ToolTipText     =   "���"
               Object.Tag             =   "���"
               ImageIndex      =   14
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split2"
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Start"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ͣ��"
               Key             =   "Stop"
               Object.ToolTipText     =   "ͣ��"
               Object.Tag             =   "ͣ��"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split3"
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Find"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   13
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�鿴"
               Key             =   "View"
               Object.ToolTipText     =   "�鿴��ʽ"
               Object.Tag             =   "�鿴"
               ImageIndex      =   10
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "  ��ͼ��"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "  Сͼ��"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "  �б�"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "  ��ϸ����"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split4"
               Style           =   3
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Quit"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageIndex      =   12
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ils32 
      Left            =   3615
      Top             =   1080
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
            Picture         =   "frmChargeManage.frx":5633
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":5A8B
            Key             =   "ItemNo"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":5EDF
            Key             =   "ItemR"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":6D31
            Key             =   "ItemRNo"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   2940
      Top             =   1020
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":7B83
            Key             =   "RootS"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":7CDD
            Key             =   "Exp"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":7E37
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":8289
            Key             =   "RootR"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":86DB
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":8B33
            Key             =   "ItemNo"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":8F87
            Key             =   "Write"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":93DB
            Key             =   "Read"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":982F
            Key             =   "ItemR"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":A681
            Key             =   "ItemRNo"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picPage 
      BorderStyle     =   0  'None
      Height          =   1650
      Index           =   2
      Left            =   6660
      ScaleHeight     =   1650
      ScaleWidth      =   900
      TabIndex        =   17
      Top             =   2520
      Width           =   900
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   3000
         Left            =   120
         TabIndex        =   18
         Top             =   120
         Width           =   5025
         Begin VB.Frame Frame1 
            BorderStyle     =   0  'None
            Height          =   2130
            Left            =   120
            TabIndex        =   26
            Top             =   720
            Width           =   4875
            Begin MSComctlLib.ListView lvwOutIn 
               Height          =   1230
               Left            =   255
               TabIndex        =   27
               Top             =   360
               Width           =   4605
               _ExtentX        =   8123
               _ExtentY        =   2170
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   0
            End
         End
         Begin VB.OptionButton opt���� 
            Caption         =   "����ȷִ�п���"
            Height          =   255
            Index           =   0
            Left            =   45
            TabIndex        =   25
            Top             =   105
            Width           =   1590
         End
         Begin VB.OptionButton opt���� 
            Caption         =   "�������ڲ���"
            Height          =   255
            Index           =   2
            Left            =   45
            TabIndex        =   24
            Top             =   405
            Width           =   1470
         End
         Begin VB.OptionButton opt���� 
            Caption         =   "����Ա���ڿ���"
            Height          =   255
            Index           =   3
            Left            =   1725
            TabIndex        =   23
            Top             =   435
            Width           =   1665
         End
         Begin VB.OptionButton opt���� 
            Caption         =   "ָ������"
            Height          =   255
            Index           =   4
            Left            =   210
            TabIndex        =   22
            Top             =   750
            Width           =   1170
         End
         Begin VB.OptionButton opt���� 
            Caption         =   "�������ڿ���"
            Height          =   255
            Index           =   1
            Left            =   1725
            TabIndex        =   21
            Top             =   120
            Width           =   1530
         End
         Begin VB.OptionButton opt���� 
            Caption         =   "Ժ��ִ��"
            Height          =   195
            Index           =   5
            Left            =   3480
            TabIndex        =   20
            Top             =   135
            Visible         =   0   'False
            Width           =   1860
         End
         Begin VB.OptionButton opt���� 
            Caption         =   "���������ڿ���"
            Height          =   195
            Index           =   6
            Left            =   3480
            TabIndex        =   19
            Top             =   465
            Width           =   1860
         End
         Begin VB.Label lblMsg 
            Caption         =   "��ҳ������鿴�������޸���˫����"
            Height          =   375
            Left            =   1560
            TabIndex        =   39
            Top             =   0
            Width           =   2895
         End
      End
   End
   Begin XtremeSuiteControls.TabControl tbPage 
      Height          =   990
      Left            =   4860
      TabIndex        =   35
      Top             =   5370
      Width           =   2055
      _Version        =   589884
      _ExtentX        =   3625
      _ExtentY        =   1746
      _StockProps     =   64
   End
   Begin MSComctlLib.ListView lvwWholeSetItem_S 
      Height          =   1095
      Left            =   6960
      TabIndex        =   37
      Top             =   810
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1931
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ils32"
      SmallIcons      =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFileset 
         Caption         =   "��ӡ����(&S)"
      End
      Begin VB.Menu mnuFilepre 
         Caption         =   "��ӡԤ��(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "��ӡ(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "�����&Excel"
      End
      Begin VB.Menu mnuFileSplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileStdImp 
         Caption         =   "��׼����(&I)"
      End
      Begin VB.Menu mnuFileStdCheck 
         Caption         =   "��׼�˲�(&C)"
      End
      Begin VB.Menu mnuFileSplit3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileParameter 
         Caption         =   "��������(&R)"
      End
      Begin VB.Menu mnuFileSplit2 
         Caption         =   "-"
      End
      Begin VB.Menu mnufileexit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuEditWholeSet 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuEditWholeSetClassAdd 
         Caption         =   "���ӳ��׷���(&N)"
      End
      Begin VB.Menu mnuEditWholeSetClassModify 
         Caption         =   "�޸ĳ��׷���(&M)"
      End
      Begin VB.Menu mnuEditWholeSetClassDelete 
         Caption         =   "ɾ�����׷���(&L)"
      End
      Begin VB.Menu mnuEditWholeSplit 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuEditWholeSetItemAdd 
         Caption         =   "���ӳ�����Ŀ(&C)"
      End
      Begin VB.Menu mnuEditWholeSetItemModify 
         Caption         =   "�޸ĳ�����Ŀ(&I)"
      End
      Begin VB.Menu mnuEditWholeSetItemDelete 
         Caption         =   "ɾ��������Ŀ(&D)"
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuEditParent 
         Caption         =   "���ӷ���(&N)"
      End
      Begin VB.Menu mnuEditModifyAssort 
         Caption         =   "�޸ķ���(&M)"
      End
      Begin VB.Menu mnuEditDeleteAssort 
         Caption         =   "ɾ������(&L)"
      End
      Begin VB.Menu mnuEditSplit0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditChild 
         Caption         =   "������Ŀ(&C)"
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "��������(&O)"
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "�޸���Ŀ(&I)"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "ɾ����Ŀ(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditSplit2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditDept 
         Caption         =   "ִ�п���(&P)"
      End
      Begin VB.Menu mnuEditSlave 
         Caption         =   "������Ŀ(&V)"
      End
      Begin VB.Menu mnuEditItemGroup 
         Caption         =   "��Ŀ���(&G)"
      End
      Begin VB.Menu mnuEditSplit3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClassEdit 
         Caption         =   "���༭(&A)"
      End
      Begin VB.Menu mnuEditExcel 
         Caption         =   "������Ŀ"
      End
      Begin VB.Menu mnuEditSplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditStart 
         Caption         =   "����(&S)"
      End
      Begin VB.Menu mnuEditStop 
         Caption         =   "ͣ��(&T)"
      End
   End
   Begin VB.Menu mnuPrice 
      Caption         =   "��Ŀ����(&T)"
      Begin VB.Menu mnuPriceRaise 
         Caption         =   "����(&R)"
      End
      Begin VB.Menu mnuPriceRaiseMass 
         Caption         =   "��������(&P)"
      End
      Begin VB.Menu mnuPriceRaiseVerify 
         Caption         =   "�������(&V)"
      End
      Begin VB.Menu mnuPriceHistory 
         Caption         =   "ɾ��δִ�м۸�(&E)"
      End
      Begin VB.Menu mnuPriceChargeSet 
         Caption         =   "�ѱ�����(&C)"
      End
      Begin VB.Menu mnuPriceReport 
         Caption         =   "��Ŀ��(&J)"
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
         Begin VB.Menu mnuViewToolspilt1 
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
      Begin VB.Menu mnuviewsplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "��ͼ��(&G)"
         Index           =   0
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "Сͼ��(&M)"
         Index           =   1
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "�б�(&L)"
         Index           =   2
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "��ϸ����(&D)"
         Checked         =   -1  'True
         Index           =   3
      End
      Begin VB.Menu mnuViewSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewShowAll 
         Caption         =   "��ʾ�����¼�(&H)"
      End
      Begin VB.Menu mnuViewShowStop 
         Caption         =   "��ʾͣ����Ŀ(&P)"
      End
      Begin VB.Menu mnuViewSplit3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewSelect 
         Caption         =   "ѡ����(&C)"
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "����(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "ˢ��(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuHelpTopic 
         Caption         =   "��������(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "Web�ϵ�����"
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
      Begin VB.Menu mnuHelpSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)..."
      End
   End
   Begin VB.Menu mnuShort1 
      Caption         =   "��ݲ˵�"
      Visible         =   0   'False
      Begin VB.Menu mnuShortMenu1 
         Caption         =   "���ӷ���(&P)"
         Index           =   1
      End
      Begin VB.Menu mnuShortMenu1 
         Caption         =   "�޸�(&M)"
         Index           =   2
      End
      Begin VB.Menu mnuShortMenu1 
         Caption         =   "ɾ��(&D)"
         Index           =   3
      End
   End
   Begin VB.Menu mnuShort2 
      Caption         =   "��ݲ˵�"
      Visible         =   0   'False
      Begin VB.Menu mnuShortMenu2 
         Caption         =   "������Ŀ(&C)"
         Index           =   0
      End
      Begin VB.Menu mnuShortMenu2 
         Caption         =   "��������(&O)"
         Index           =   1
      End
      Begin VB.Menu mnuShortMenu2 
         Caption         =   "�޸�(&M)"
         Index           =   2
      End
      Begin VB.Menu mnuShortMenu2 
         Caption         =   "ִ�п���(&P)"
         Index           =   3
      End
      Begin VB.Menu mnuShortMenu2 
         Caption         =   "������Ŀ(&V)"
         Index           =   4
      End
      Begin VB.Menu mnuShortMenu2 
         Caption         =   "ɾ��(&D)"
         Index           =   5
      End
      Begin VB.Menu mnuShortMenu2 
         Caption         =   "��Ŀ���"
         Index           =   6
      End
      Begin VB.Menu mnuShortsplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShortIcon 
         Caption         =   "��ͼ��(&G)"
         Index           =   0
      End
      Begin VB.Menu mnuShortIcon 
         Caption         =   "Сͼ��(&M)"
         Index           =   1
      End
      Begin VB.Menu mnuShortIcon 
         Caption         =   "�б�(&L)"
         Index           =   2
      End
      Begin VB.Menu mnuShortIcon 
         Caption         =   "��ϸ����(&D)"
         Checked         =   -1  'True
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmChargeManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng���볤�� As Long
Dim mintColumn1 As Integer
Dim msngStartX As Single, msngStartY As Single    '�ƶ�ǰ����λ��
Dim mblnItem As Boolean  'Ϊ���ʾ������ListViewĳһ����
Dim mblnLoad As Boolean
Dim mintColumn As Integer
Dim mstrKey As String       'ǰһ�����ڵ�Ĺؼ�ֵ
Dim mstr�ϼ�Key As String   '��ǰ��Ŀ���ϼ����ڵ��Keyֵ����Ҫ��������ʾ�����¼���Ŀ
Dim mstrClass As String     '������¼������
Dim mstrClassName As String '������¼�������
Dim mbln����ҽ��ϵͳ As Boolean '�Ƿ�������ҽ��ϵͳ
Private Const mstrLvw As String = "����,1500,0,1;����,1000,0,2;��ʶ����,1400,0,0;��ʶ����,900,0,0;��ѡ��,1400,0,0;" & _
                                "���,550,0,0;���㵥λ,900,0,0;�������,1000,0,2;��������,900,0,0;�������,900,0,0;" & _
                                "˵��,1440,0,0;���ηѱ�,900,0,0;�Ƿ���,900,0,0;�Ӱ�Ӽ�,900,0,0;����ժҪ,900,0,0;" & _
                                "��Ŀ����,1100,0,2;����޼�,1000,1,0;����޼�,1000,1,0;����ʱ��,1100,0,0;����ʱ��,1100,0,0;" & _
                                "��������,1300,0,2;������Ŀ,1300,0,2;Ժ��,0,0,2"
Private Const mstrLvwWholeSet As String = "����,1500,0,1;����,800,0,2;ƴ��,1400,0,0;���,1400,0,0;ʹ�÷�Χ,1000,0,0;��������,2400,0,0"
Private mlngMode As Long
Public mstrPrivs As String                              'Ȩ�޴�
Private mint�ϴ�ϸĿҳ As Integer '
Private mint�ϴγ���ҳ As Integer '
Private mfrmEarnRS As New frmEarnRS
Private mblnVerifyFlow As Boolean   '�����Ƿ�������������̣�true-���ã�false-δ����
Private mblnVerifyPris As Boolean   '��˵��۵�Ȩ�� true-��Ȩ�ޣ�false-��Ȩ��

Private Enum mCalssPage
    pg_ϸĿ = 1
    pg_���� = 2
End Enum
Private Enum mItemPage
    pg_��Ŀ = 1
    pg_ִ�п��� = 2
    pg_������Ŀ = 3
    pg_���� = 4
    pg_�ѱ�ȼ� = 5
    pg_������� = 6
    pg_����ʹ�ÿ��� = 7
End Enum

Private Enum mIndex��Ŀ
    Col_�۸�ȼ� = 0
    Col_���ݺ�
    Col_ִ������
    Col_��ֹ����
    Col_������Ŀ
    Col_ԭ��
    Col_�ּ�
    Col_���������շ���
    Col_�Ӱ�Ӽ���
    Col_����˵��
    Col_ȱʡ�۸�
    Col_������
End Enum

Private mblnNotClick As Boolean
Private mblnCanUpdateAll As Boolean '�Ƿ��������������Ŀ��δ���ü۸�ȼ��������˼۸�ȼ��С�����Ժ����Ȩ��

Private Sub zlInitClassPage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������ҳ��
    '����:���˺�
    '����:2010-08-24 10:15:11
    '˵��:27327
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, ObjItem As TabControlItem
    Err = 0: On Error GoTo ErrHand:
    mblnNotClick = True
    Set ObjItem = tbClassPage.InsertItem(mCalssPage.pg_ϸĿ, "�շ�ϸĿ", picTreeItem.hwnd, 0)
    ObjItem.Tag = mCalssPage.pg_ϸĿ
    ObjItem.Selected = True
    Set ObjItem = tbClassPage.InsertItem(mCalssPage.pg_����, "������Ŀ", picTreeWholeSet.hwnd, 0)
    ObjItem.Tag = mCalssPage.pg_����
     With tbClassPage
        
        .PaintManager.Appearance = xtpTabAppearanceVisio
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = True
        .PaintManager.ClientFrame = xtpTabFrameBorder
    End With
    
    mblnNotClick = False
    Set ObjItem = tbPage.InsertItem(mItemPage.pg_��Ŀ, "�շѼ�Ŀ", picPage(1).hwnd, 0)
    ObjItem.Tag = mItemPage.pg_��Ŀ
    Set ObjItem = tbPage.InsertItem(mItemPage.pg_ִ�п���, "ִ�п���", picPage(2).hwnd, 0)
    ObjItem.Tag = mItemPage.pg_ִ�п���
    Set ObjItem = tbPage.InsertItem(mItemPage.pg_������Ŀ, "������Ŀ", picPage(3).hwnd, 0)
    ObjItem.Tag = mItemPage.pg_������Ŀ
    Set ObjItem = tbPage.InsertItem(mItemPage.pg_����, "����", picPage(4).hwnd, 0)
    ObjItem.Tag = mItemPage.pg_����
    Set ObjItem = tbPage.InsertItem(mItemPage.pg_�ѱ�ȼ�, "�ѱ�ȼ�", picPage(5).hwnd, 0)
    ObjItem.Tag = mItemPage.pg_�ѱ�ȼ�
    Set ObjItem = tbPage.InsertItem(mItemPage.pg_�������, "������Ŀ���", picPage(6).hwnd, 0)
    ObjItem.Tag = mItemPage.pg_�������
    Set ObjItem = tbPage.InsertItem(mItemPage.pg_����ʹ�ÿ���, "������Ŀʹ�ÿ���", picPage(7).hwnd, 0)
    ObjItem.Tag = mItemPage.pg_����ʹ�ÿ���
    tbPage.Item(0).Selected = True
    mint�ϴγ���ҳ = mItemPage.pg_�������
    mint�ϴ�ϸĿҳ = mItemPage.pg_��Ŀ
    
     With tbPage
        .PaintManager.Appearance = xtpTabAppearancePropertyPage
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = True
        .PaintManager.ClientFrame = xtpTabFrameBorder
    End With
    Call SetPageVisible
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function Check�շ���Ŀ(ByVal ID As Long, ByRef strMsg As String) As Boolean
    '����շ���Ŀ���ڵĸ���������ϵ
    Dim rs As New ADODB.Recordset
        
    On Error GoTo ErrHandle

    '1.���շ���Ŀ�Ƿ������������շѶ����С�
    gstrSQL = "Select 1 From �����շѹ�ϵ where RowNum=1 and �շ���Ŀid=[1] "
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "��������շѶ��ա�", ID)
    
    strMsg = IIF(rs.RecordCount = 0, "", "[����Ŀ���������շѶ��չ�ϵ��]" & vbCrLf)
        
    '2.���շ���Ŀ�Ƿ�����Ϊ������Ŀ�Ĵ�����Ŀ��
    gstrSQL = "Select 1 From �շѴ�����Ŀ where RowNum=1 and (����id=[1] or ����id=[1] )"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "����շѴ�����Ŀ��", ID)
    
    strMsg = strMsg & IIF(rs.RecordCount = 0, "", "[����Ŀ�����շѴ�����ϵ��]" & vbCrLf)
    
    '3.���շ���Ŀ�Ƿ��ض��շ���Ŀ��
    gstrSQL = "Select 1 From �շ��ض���Ŀ where RowNum=1 and �շ�ϸĿid=[1] "
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "����շ��ض���Ŀ��", ID)
    
    strMsg = strMsg & IIF(rs.RecordCount = 0, "", "[����Ŀ���շ��ض���Ŀ��]" & vbCrLf)
    
    '4.���շ���Ŀ�Ƿ�����Ϊ�Զ��Ƽ���Ŀ��
    gstrSQL = "Select 1 From �Զ��Ƽ���Ŀ where RowNum=1 and �շ�ϸĿid=[1] "
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "����Զ��Ƽ���Ŀ��", ID)
    
    strMsg = strMsg & IIF(rs.RecordCount = 0, "", "[����Ŀ���Զ��Ƽ���Ŀ��]" & vbCrLf)
    
    Check�շ���Ŀ = True
    Exit Function
    
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    
End Function


Private Sub Form_Activate()
    On Error GoTo ErrHandle
    
    mblnVerifyFlow = IIF(Val(zlDatabase.GetPara("������Ҫ���", glngSys, 1009, 0)) = 0, False, True)
    mblnVerifyPris = IIF(InStr(1, ";" & gstrPrivs & ";", ";�շѼ�Ŀ�������;") > 0, True, False)
    If mblnLoad = False Then Exit Sub
    Call Form_Resize 'Ϊ��ʹCoolBar����Ӧ�߶�
    Call FillTree:
    Call FillWholeSetTree
    mblnLoad = False
    
    If checkNotPrice(0) = False Then
        MsgBox "�շ�ϸĿ�л�����δ��˵ļ۸���ע����ˣ�", vbInformation, gstrSysName
    End If
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub zlInitLvwHeadCol()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��ListView��ͷ
    '����:���˺�
    '����:2010-08-24 14:21:48
    '˵��:27327
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '���ListView�Ļ�δ�����ã������һ��ʹ�ã��Ǿ͵���ȱʡ�ĳ�ʼ��
    lvwWholeSetItem_S.ListItems.Clear
    zlControl.LvwSelectColumns lvwWholeSetItem_S, mstrLvwWholeSet, True
    lvwMain_S.ListItems.Clear
    zlControl.LvwSelectColumns lvwMain_S, mstrLvw, True
End Sub

Private Sub Form_Load()
    Dim intType  As Integer
    On Error GoTo ErrHandle
    mblnLoad = True
    
    mlngMode = glngModul
    mstrPrivs = gstrPrivs
    mblnCanUpdateAll = IsPriceGradeEnabled() = False Or zlStr.IsHavePrivs(mstrPrivs, "����Ժ��")   '110070
    
    Call GetPriceGrade(gstrҩƷ�۸�ȼ�, gstr���ļ۸�ȼ�, gstr��ͨ�۸�ȼ�)
    Call zlInitClassPage
    
    'Ȩ��
    Call Ȩ�޿���

    '���������ɾ����ListView�������
    lvwMain_S.Tag = "�ɱ仯��"
    lvwWholeSetItem_S.Tag = "�ɱ仯��"
    mnuViewShowAll.Checked = (Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "��ʾ����", 0)) = 1)
    mnuViewShowStop.Checked = (Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "��ʾͣ��", 0)) = 1)
    chk�۸�.value = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "��ʾ��ʷ�۸�", "0"))
    
    Call zlInitLvwHeadCol
    
    '����lvwMain_S��ʾ���ö�Ӧ�˵�
     mnuViewIcon_Click lvwMain_S.View
    
    '��ʼ�����½ǵ�ִ�п�����
    zlControl.LvwSelectColumns lvwOutIn, "ִ�п���,3000,0,0;���˿���,8000,0,0", True
    zlControl.LvwFlatColumnHeader lvwOutIn
    
    RestoreWinState Me, App.ProductName
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngMode, mstrPrivs, "ZL1_INSIDE_1009")
    
    Call InitTable
    Call GetDefineSize
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub InitTable()
On Error GoTo ErrHandle
    '��ʼ���շѼ�Ŀ��ʹ�����Ŀ��
    msh��Ŀ.Cols = 12
    
    msh��Ŀ.ColWidth(Col_�۸�ȼ�) = 250
    msh��Ŀ.ColWidth(Col_���ݺ�) = 1000
    msh��Ŀ.ColWidth(Col_ִ������) = 2000
    msh��Ŀ.ColWidth(Col_��ֹ����) = 2000
    msh��Ŀ.ColWidth(Col_������Ŀ) = 1000
    msh��Ŀ.ColWidth(Col_ԭ��) = 1000
    msh��Ŀ.ColWidth(Col_�ּ�) = 1000
    msh��Ŀ.ColWidth(Col_���������շ���) = 1000
    msh��Ŀ.ColWidth(Col_�Ӱ�Ӽ���) = 1000
    msh��Ŀ.ColWidth(Col_����˵��) = 3000
    msh��Ŀ.ColWidth(Col_ȱʡ�۸�) = 1000
    msh��Ŀ.ColWidth(Col_������) = 800
    
    msh����.ColWidth(0) = 1000
    msh����.ColWidth(1) = 3000
    msh����.ColWidth(2) = 1000
    msh����.ColWidth(3) = 1500
    
    msh��Ŀ.TextMatrix(0, Col_�۸�ȼ�) = ""
    msh��Ŀ.TextMatrix(0, Col_���ݺ�) = "���ݺ�"
    msh��Ŀ.TextMatrix(0, Col_ִ������) = "ִ������"
    msh��Ŀ.TextMatrix(0, Col_��ֹ����) = "��ֹ����"
    msh��Ŀ.TextMatrix(0, Col_������Ŀ) = "������Ŀ"
    msh��Ŀ.TextMatrix(0, Col_ԭ��) = "ԭ��"
    msh��Ŀ.TextMatrix(0, Col_�ּ�) = "�ּ�"
    msh��Ŀ.TextMatrix(0, Col_���������շ���) = "���������շ���"
    msh��Ŀ.TextMatrix(0, Col_�Ӱ�Ӽ���) = "�Ӱ�Ӽ���"
    msh��Ŀ.TextMatrix(0, Col_����˵��) = "����˵��"
    msh��Ŀ.TextMatrix(0, Col_ȱʡ�۸�) = "ȱʡ�۸�"
    msh��Ŀ.TextMatrix(0, Col_������) = "������"
   
    mshAlias.Cols = 4
    mshAlias.ColWidth(0) = 1000
    mshAlias.ColWidth(1) = 4000
    mshAlias.ColWidth(2) = 800
    mshAlias.ColWidth(3) = 3000
    
    mshAlias.TextMatrix(0, 0) = "��������"
    mshAlias.TextMatrix(0, 1) = "����"
    mshAlias.TextMatrix(0, 2) = "����"
    mshAlias.TextMatrix(0, 3) = "����"
    
    msh����.TextMatrix(0, 0) = "�շ����"
    msh����.TextMatrix(0, 1) = "�շ���Ŀ"
    msh����.TextMatrix(0, 2) = "����"
    msh����.TextMatrix(0, 3) = "�̶�"
    msh����.TextMatrix(0, 4) = "״̬"
    msh����.Col = 0
    msh����.Row = 0
    msh����.ColSel = 3
    msh����.RowSel = 0
    msh����.FillStyle = flexFillRepeat
    msh����.CellAlignment = 4
    msh����.FillStyle = flexFillSingle
    msh����.AllowBigSelection = False
    msh����.Row = 1
    msh����.ColAlignment(3) = 1
    msh����.ColAlignment(0) = 1
    msh����.ColAlignment(1) = 1
    
    msh��Ŀ.ColAlignment(1) = 1
    msh��Ŀ.FillStyle = flexFillRepeat
    msh��Ŀ.CellAlignment = 4
    msh��Ŀ.FillStyle = flexFillSingle
    msh��Ŀ.AllowBigSelection = False
    msh��Ŀ.Row = 1
    msh��Ŀ.FixedAlignment(-1) = flexAlignCenterCenter
    msh��Ŀ.HighLight = flexHighlightNever
    
    mshAlias.Col = 0
    mshAlias.Row = 1
    mshAlias.ColSel = 1
    mshAlias.RowSel = 0
    mshAlias.FillStyle = flexFillRepeat
    mshAlias.CellAlignment = 4
    mshAlias.FillStyle = flexFillSingle
    mshAlias.Row = 1
    
    With msh�ѱ�
        .Cols = 4
        .ColWidth(0) = 1500
        .ColWidth(1) = 3000
        .ColWidth(2) = 1050
        .ColWidth(3) = 2000
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .TextMatrix(0, 0) = "�ѱ�"
        .TextMatrix(0, 1) = "Ӧ�ս��(Ԫ)"
        .TextMatrix(0, 2) = "ʵ�ձ���(%)"
        .TextMatrix(0, 3) = "���㷽��"
        
        .MergeCol(0) = True
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
 

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrHandle
    mstrKey = ""
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "��ʾ����", IIF(mnuViewShowAll.Checked, 1, 0)
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "��ʾͣ��", IIF(mnuViewShowStop.Checked, 1, 0)
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "��ʾ��ʷ�۸�", chk�۸�.value
    zl_vsGrid_Para_Save mlngMode, vsWholeSet, Me.Caption, "������Ŀ��ɱ���-������", True, True
    SaveWinState Me, App.ProductName
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
 
Private Sub lvwWholeSetItem_S_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo ErrHandle
    If mintColumn1 = ColumnHeader.Index - 1 Then '���Ǹղ�����
        lvwWholeSetItem_S.SortOrder = IIF(lvwWholeSetItem_S.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintColumn1 = ColumnHeader.Index - 1
        lvwWholeSetItem_S.SortKey = mintColumn1
        lvwWholeSetItem_S.SortOrder = lvwAscending
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub lvwWholeSetItem_S_DblClick()
    Dim strID As String, lng�ϼ�ID As Long
    If Not mblnItem Then Exit Sub
    If mnuEditWholeSetClassModify.Enabled And mnuEditWholeSetClassModify.Visible Then
        Call mnuEditWholeSetItemModify_Click
    Else
        If Me.lvwWholeSetItem_S.SelectedItem Is Nothing Then Exit Sub
        With lvwWholeSetItem_S
            strID = Val(Mid(.SelectedItem.Key, 2))
        End With
        If frmChargeWholeSetItemEdit.ShowCard(Me, EdI_�鿴, mstrPrivs, mlngMode, lng�ϼ�ID, strID) = False Then Exit Sub
    End If
End Sub

Private Sub lvwWholeSetItem_S_GotFocus()
    Call MenuSet
End Sub

Private Sub lvwWholeSetItem_S_ItemClick(ByVal Item As MSComctlLib.ListItem)
    '���˺�:27327
    'Ϊ������Ŀά��ʱ,��Ҫ��������
    mblnItem = True
    If lvwWholeSetItem_S.Tag <> Item.Key Then
        Call FillWholeSetItemChildData(Val(Mid(Item.Key, 2)))
    End If
    lvwWholeSetItem_S.Tag = Item.Key
End Sub

Private Sub lvwWholeSetItem_S_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If mnuEditWholeSetItemModify.Enabled And mnuEditWholeSetItemModify.Visible Then mnuEditWholeSetItemModify_Click
    End If
End Sub

Private Sub lvwWholeSetItem_S_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnItem = False
End Sub

Private Sub lvwWholeSetItem_S_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mnuEditWholeSet.Visible = False Then Exit Sub
    If Button <> 2 Then Exit Sub
    PopupMenu mnuEditWholeSet, vbPopupMenuRightButton
End Sub

Private Sub mnuEditExcel_Click()
    frmItemImport.ShowMe 1, Me
    Call FillTree
    
End Sub

Private Sub mnuEditWholeSetClassAdd_Click()
    Dim lng�ϼ�ID As Long
    With tvwWholeSet
        If .SelectedItem Is Nothing Then Exit Sub
        lng�ϼ�ID = Val(Mid(.SelectedItem.Key, 2))
    End With
    If InStr(1, mstrPrivs, ";���ӳ�����Ŀ;") = 0 Then Exit Sub
    If frmChargeWholeSetClassEdit.EditCard(Me, Ed_����, mstrPrivs, mlngMode, lng�ϼ�ID, "") = False Then Exit Sub
    Call FillWholeSetTree

End Sub

Private Sub mnuEditWholeSetClassDelete_Click()
    Dim strKey As String, intIndex As Long
    On Error GoTo ErrHandle
    With tvwWholeSet
        If .SelectedItem Is Nothing Then Exit Sub
        If .SelectedItem.Key = "Root" Then Exit Sub
        If MsgBox("��ȷ��Ҫɾ������Ϊ��" & .SelectedItem.Text & "���ķ�����", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
            Me.MousePointer = 11
            gstrSQL = "Zl_������Ŀ����_Delete(" & Val(Mid(.SelectedItem.Key, 2)) & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            Me.MousePointer = 0
             
            strKey = .SelectedItem.Key
            If Not .SelectedItem.Next Is Nothing Then
                 .SelectedItem.Next.Selected = True
                tvwWholeSet_NodeClick tvwWholeSet.SelectedItem
            Else
                If Not .SelectedItem.Parent Is Nothing Then
                     .SelectedItem.Parent.Selected = True
                End If
                If Not .SelectedItem Is Nothing Then
                    tvwWholeSet_NodeClick tvwWholeSet.SelectedItem
                End If
            End If
             .Nodes.Remove strKey
        End If
    End With
    MenuSet
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Me.MousePointer = 0
End Sub

Private Sub mnuEditWholeSetClassModify_Click()
    Dim lng�ϼ�ID As Long, lngId As Long
    If InStr(1, mstrPrivs, ";�޸ĳ�����Ŀ;") = 0 Then Exit Sub
    With tvwWholeSet
        If .SelectedItem Is Nothing Then Exit Sub
        If .SelectedItem.Key = "Root" Then
            lng�ϼ�ID = 0
            lngId = 0
        Else
            lng�ϼ�ID = Val(Mid(.SelectedItem.Parent.Key, 2))
            lngId = Mid(.SelectedItem.Key, 2)
        End If
    End With
    If frmChargeWholeSetClassEdit.EditCard(Me, Ed_�޸�, mstrPrivs, mlngMode, lng�ϼ�ID, lngId) = False Then Exit Sub
    Call FillWholeSetTree
End Sub

Private Sub mnuEditWholeSetItemAdd_Click()
    '������Ŀ����
    Dim lng�ϼ�ID As Long
    With tvwWholeSet
        If .SelectedItem Is Nothing Then Exit Sub
        lng�ϼ�ID = Val(Mid(.SelectedItem.Key, 2))
    End With
    If InStr(1, mstrPrivs, ";���ӳ�����Ŀ;") = 0 Then Exit Sub
    If frmChargeWholeSetItemEdit.ShowCard(Me, EdI_����, mstrPrivs, mlngMode, lng�ϼ�ID, "") = False Then Exit Sub
    Call FillWholeItem(lng�ϼ�ID)
End Sub

Private Sub mnuEditWholeSetItemDelete_Click()
    'ɾ����Ŀ
    Dim lngId As Long, strKey As String
    Dim intIndex As Long
     
    '�޸���Ŀ
    If Not (mnuEditWholeSetItemModify.Enabled And mnuEditWholeSetItemModify.Visible) Then Exit Sub
    If Me.lvwWholeSetItem_S.SelectedItem Is Nothing Then Exit Sub
    
    If InStr(1, mstrPrivs, ";�޸ĳ�����Ŀ;") = 0 Then Exit Sub
    With lvwWholeSetItem_S
        lngId = Val(Mid(.SelectedItem.Key, 2))
        strKey = .SelectedItem.Key
    End With
    If MsgBox("��ȷ��Ҫɾ������Ϊ��" & lvwWholeSetItem_S.SelectedItem.Text & "���ĳ����շ���Ŀ��", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    On Error GoTo ErrHandle
    Me.MousePointer = 11
    'Zl_�����շ���Ŀ_Delete(Id_In In �����շ���Ŀ.ID%Type)
    gstrSQL = "Zl_�����շ���Ŀ_Delete(" & lngId & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Me.MousePointer = 0
    With lvwWholeSetItem_S
        intIndex = .SelectedItem.Index
        .ListItems.Remove .SelectedItem.Key
        If .ListItems.Count > 0 Then
            intIndex = IIF(.ListItems.Count > intIndex, intIndex, .ListItems.Count)
            .ListItems(intIndex).Selected = True
            .ListItems(intIndex).EnsureVisible
            FillWholeSetItemChildData Val(Mid(.SelectedItem.Key, 2))
        Else
            FillWholeSetItemChildData 0
        End If
    End With
    MenuSet
    Me.MousePointer = 0
    Exit Sub
ErrHandle:
    Me.MousePointer = 0
    If ErrCenter() = 1 Then
        Me.MousePointer = 11
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuEditWholeSetItemModify_Click()
    Dim strID As String, lng�ϼ�ID As Long
    '�޸���Ŀ
    If Not (mnuEditWholeSetItemModify.Enabled And mnuEditWholeSetItemModify.Visible) Then Exit Sub
    If Me.lvwWholeSetItem_S.SelectedItem Is Nothing Then Exit Sub
    
    If InStr(1, mstrPrivs, ";�޸ĳ�����Ŀ;") = 0 Then Exit Sub
    With lvwWholeSetItem_S
        strID = Val(Mid(.SelectedItem.Key, 2))
    End With
    With tvwWholeSet
        If .SelectedItem Is Nothing Then
            lng�ϼ�ID = 0
        Else
            lng�ϼ�ID = Val(Mid(.SelectedItem.Key, 2))
        End If
    End With
    If frmChargeWholeSetItemEdit.ShowCard(Me, EdI_�޸�, mstrPrivs, mlngMode, lng�ϼ�ID, strID) = False Then Exit Sub
    Call FillWholeItem(lng�ϼ�ID)
End Sub

Private Sub mnuFileParameter_Click()
    frmParSetFeeItem.ShowMe Me
    
    mblnVerifyFlow = IIF(Val(zlDatabase.GetPara("������Ҫ���", glngSys, 1009, 0)) = 0, False, True)
    mnuPriceRaiseVerify.Visible = mblnVerifyFlow
    Toolbar1.Buttons("RaiseVerify").Visible = mblnVerifyFlow   '�������
End Sub

Private Sub mnuPriceRaiseVerify_Click()
    frmChargePriceVerify.ShowMe Me, mblnCanUpdateAll
End Sub

Private Sub msh��Ŀ_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = Col_�۸�ȼ� Then Cancel = True
End Sub

Private Sub picPage_DblClick(Index As Integer)
    
    On Error GoTo ErrHandle
    Select Case Index
        Case 1
            If mnuPriceRaise.Enabled = True And mnuPriceRaise.Visible = True Then Call mnuPriceRaise_Click
        Case 2
            If mnuEditDept.Enabled = True And mnuEditDept.Visible = True Then Call mnuEditDept_Click
        Case 3
            If mnuEditSlave.Enabled = True And mnuEditSlave.Visible = True Then Call mnuEditSlave_Click
        Case 4
            If mnuEditModify.Enabled = True And mnuEditModify.Visible = True Then Call mnuEditModify_Click
        Case 5
            If mnuPriceChargeSet.Enabled = True And mnuPriceChargeSet.Visible = True Then Call mnuPriceChargeSet_Click
    End Select
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub lvwMain_S_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo ErrHandle
    If mintColumn = ColumnHeader.Index - 1 Then '���Ǹղ�����
        lvwMain_S.SortOrder = IIF(lvwMain_S.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintColumn = ColumnHeader.Index - 1
        lvwMain_S.SortKey = mintColumn
        lvwMain_S.SortOrder = lvwAscending
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub lvwMain_S_DblClick()
On Error GoTo ErrHandle
    If mblnItem = True And mnuEditModify.Enabled And mnuEditModify.Visible Then mnuEditModify_Click
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub lvwMain_S_KeyPress(KeyAscii As Integer)
On Error GoTo ErrHandle
    If KeyAscii = vbKeyReturn Then
        If mnuEditModify.Enabled And mnuEditModify.Visible Then mnuEditModify_Click
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub lvwMain_S_GotFocus()
On Error GoTo ErrHandle
    Call MenuSet
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub lvwMain_S_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long
    Dim strTips As String
    Dim intName As Integer
    Dim intArrit As Integer
    Dim intAssort As Integer
    Dim intCode As Integer

On Error GoTo ErrHandle
    mstrClass = ""
    mstrClassName = ""
    If Item Is Nothing Then Exit Sub
    
    '�����Ϊ�Һš�������λ��
    strTips = ""
    For i = 1 To lvwMain_S.ColumnHeaders.Count - 1
        If lvwMain_S.ColumnHeaders(i).Text = "��Ŀ����" Then
            intArrit = i - 1
        ElseIf lvwMain_S.ColumnHeaders(i).Text = "����" Then
            intName = i - 1
        ElseIf lvwMain_S.ColumnHeaders(i).Text = "������Ŀ" Then
            intAssort = i - 1
        ElseIf lvwMain_S.ColumnHeaders(i).Text = "����" Then
            intCode = i - 1
        End If
    Next
    
    mblnItem = True
    FillItem Item.Key
    
    If lvwMain_S.ListItems.Count > 0 Then
        strTips = "��ǰ�����²��ҵ� " & lvwMain_S.ListItems.Count & " ����Ŀ"
    Else
        strTips = "��ǰ��������Ŀ"
    End If
    Me.stbThis.Panels(2).Text = strTips
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub lvwMain_S_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnItem = False
End Sub

Private Sub lvwMain_S_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrHandle
    Dim i As Long
    Dim intArrit As Integer
    For i = 1 To lvwMain_S.ColumnHeaders.Count - 1
        If lvwMain_S.ColumnHeaders(i).Text = "��Ŀ����" Then
            intArrit = i - 1
            Exit For
        End If
    Next
    
    If Button = 2 Then
        If lvwMain_S.ListItems.Count < 1 Then
            mnuEditCopy.Enabled = False
            mnuEditModify.Enabled = False
        End If
        mnuShortMenu2(0).Enabled = mnuEditChild.Enabled
        mnuShortMenu2(1).Enabled = mnuEditCopy.Enabled
        mnuShortMenu2(2).Enabled = mnuEditModify.Enabled
        mnuShortMenu2(3).Enabled = mnuEditDept.Enabled
        mnuShortMenu2(4).Enabled = mnuEditSlave.Enabled
        mnuShortMenu2(5).Enabled = mnuEditDelete.Enabled
        mnuShortMenu2(6).Enabled = mnuEditItemGroup.Enabled
        
        For i = 0 To 3
            mnuShortIcon(i).Checked = mnuViewIcon(i).Checked
        Next
        PopupMenu mnuShort2, vbPopupMenuRightButton
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuClassEdit_Click()
    frmChargeClassEdit.Show 1, Me
End Sub

Private Sub mnuPriceChargeSet_Click()
    If Me.lvwMain_S.ListItems.Count > 0 Then
        If frmChargeSortItemEdit.ShowMe(Me, 2, "", Val(Me.lvwMain_S.SelectedItem.Tag), Me.lvwMain_S.SelectedItem.Text) Then
            Call frmChargeManage.FillTree
        End If
    End If
End Sub

Private Sub mnuEditItemGroup_Click()
    If Me.lvwMain_S.ListItems.Count > 0 Then
        frmChargeGroupItem.ShowMe Me, Me.lvwMain_S.SelectedItem.Tag
    End If
End Sub

Private Sub mnuFileStdCheck_Click()
    frmStdCheck.Show 1, Me
End Sub

Private Sub mnuFileStdImp_Click()
    frmPriceImp.Show 1, Me
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuEditDept_Click()
On Error GoTo ErrHandle
    If mnuEdit.Visible = False Then Exit Sub
    
    ModifyMode 3  'editDept
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuEditSlave_Click()
On Error GoTo ErrHandle
    If mnuEdit.Visible = False Then Exit Sub
    ModifyMode 4    'editSlave
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuPriceHistory_Click()
'ɾ��δִ�м۸�
    Dim strNodeNo As String
    
    If lvwMain_S.SelectedItem Is Nothing Then Exit Sub
    '110133
    '���һ����Ժ������
    strNodeNo = lvwMain_S.SelectedItem.ListSubItems(lvwMain_S.SelectedItem.ListSubItems.Count).Tag
    If mblnCanUpdateAll Or strNodeNo = gstrNodeNo Then
        If MsgBox("��ȷ��Ҫɾ�����һ�ε�δִ�м۸���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    Else
        If MsgBox("��ȷ��Ҫɾ�����һ�ε�δִ�м۸���" & vbCrLf & vbCrLf & _
            "ע�⣺������û�С�����Ժ����Ȩ�ޣ�����ֻ��ɾ���۸�ȼ���" & gstr��ͨ�۸�ȼ� & "�����һ�ε�δִ�м۸�", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    On Error GoTo ErrHandle
    MousePointer = 11
    'Zl_�շѼ�Ŀ_Delete(
    gstrSQL = "zl_�շѼ�Ŀ_Delete("
    '  ϸĿid_In In �շѼ�Ŀ.�շ�ϸĿid%Type,
    gstrSQL = gstrSQL & "" & Val(lvwMain_S.SelectedItem.Tag) & ","
    '  վ��_In   In �շ���ĿĿ¼.վ��%Type := Null
    gstrSQL = gstrSQL & "" & IIF(mblnCanUpdateAll, "NULL", "'" & gstrNodeNo & "'") & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    FillItem lvwMain_S.SelectedItem.Key
    MousePointer = 0
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    MousePointer = 0
End Sub

Private Sub mnuPriceRaiseMass_Click()
'��������
On Error GoTo ErrHandle
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    Dim datToday As Date
    
    datToday = sys.Currentdate
    If tvwMainItem.SelectedItem Is Nothing Then Exit Sub
    With frmChargeBatchPrice
        .mstrPrivs = mstrPrivs
        .mblnCanUpdateAll = mblnCanUpdateAll
        
        .txtType.Text = mstrClassName
        .lbl���.Tag = mstrClass
        
        .txtChargeType.Text = tvwMainItem.SelectedItem.Text
        .lbl����.Tag = tvwMainItem.SelectedItem.Tag
        '�����ִ������
        strSQL = "select max(B.ִ������) as ������� from " & _
                "(select id from �շ���ĿĿ¼ where (����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null) And �Ƿ���=0 and " & IIF(.lbl����.Tag = "", "���=[1] and ����ID is null", "����ID=[2] ") & ") A" & _
                ",�շѼ�Ŀ B  Where A.ID = B.�շ�ϸĿID "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, .lbl���.Tag, .lbl����.Tag)
                
        If rsTemp("�������") > datToday Then
            .datSingle = rsTemp("�������") + 1
        Else
            .datSingle = datToday + 1
        End If
        
        '����С���������������Ŀ��
        strSQL = "select min(B.�ּ�) as ��С��� from " & _
                "(select id from �շ���ĿĿ¼ where (����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null) And �Ƿ���=0 and " & IIF(.lbl����.Tag = "", "���=[1] and ����ID is null", "����ID=[2] ") & ") A" & _
                ",�շѼ�Ŀ B  Where A.ID = B.�շ�ϸĿID And Sysdate Between b.ִ������ And Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD')) "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, .lbl���.Tag, .lbl����.Tag)
        
        .dblSingle = IIF(IsNull(rsTemp("��С���")), 0, rsTemp("��С���"))
        
        strSQL = "select max(B.ִ������) as ������� from " & _
                "(select c.id from �շ���ĿĿ¼ c,(Select id From �շѷ���Ŀ¼  start with  " & IIF(.lbl����.Tag = "", "���=[1] and �ϼ�ID is null", "�ϼ�ID=[2] ") & " connect by prior id=�ϼ�ID) d  " & _
                " where c.�Ƿ���=0 And c.����id=d.Id And (c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.����ʱ�� Is Null)) A" & _
                ",�շѼ�Ŀ B  Where A.ID = B.�շ�ϸĿID "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, .lbl���.Tag, .lbl����.Tag)
        
        If rsTemp("�������") > datToday Then
            .datAll = rsTemp("�������") + 1
        Else
            .datAll = datToday + 1
        End If
        
        strSQL = "select min(B.�ּ�) as ��С��� from " & _
                "(select c.id from �շ���ĿĿ¼ c,(Select id From �շѷ���Ŀ¼  start with  " & IIF(.lbl����.Tag = "", "���=[1] and �ϼ�ID is null", "�ϼ�ID=[2] ") & " connect by prior id=�ϼ�ID) d  " & _
                " where c.�Ƿ���=0 And c.����id=d.Id And (c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.����ʱ�� Is Null)) A" & _
                ",�շѼ�Ŀ B  Where A.ID = B.�շ�ϸĿID And Sysdate Between b.ִ������ And Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD')) "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, .lbl���.Tag, .lbl����.Tag)
        
        .dblAll = IIF(IsNull(rsTemp("��С���")), 0, rsTemp("��С���"))
                
        rsTemp.Close
        
        .dtpBegin.value = .datSingle
        .dtpBegin.MinDate = .datSingle
    End With
    frmChargeBatchPrice.Show vbModal, Me
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuPriceReport_Click()
'��Ŀ��
On Error GoTo ErrHandle
    Dim str��� As String
    Dim lng����id As Long
    Dim strCaption As String
    
    If tvwMainItem.Nodes.Count > 0 Then
        lng����id = Val(Mid(tvwMainItem.SelectedItem.Key, 2))
    End If
    
    On Error Resume Next
    ReportOpen gcnOracle, glngSys, "ZL1_INSIDE_1009", Me, strCaption, gstrUserName, "����=" & lng����id, _
        "վ��='" & IIF(mblnCanUpdateAll, "ȫԺ", gstrNodeNo) & "'"
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub mnuReportItem_Click(Index As Integer)
    'Ĭ�ϲ���������=����id����Ŀ=��Ŀid�����=�շ��������
    Dim lng����id As Long
    Dim lng��Ŀid As Long
    Dim str�շ���� As String
    
    If Not tvwMainItem.SelectedItem Is Nothing Then
        lng����id = Mid(tvwMainItem.SelectedItem.Key, 2)
    End If
    
    If Not lvwMain_S.SelectedItem Is Nothing Then
        lng��Ŀid = Mid(lvwMain_S.SelectedItem.Key, 3)
        str�շ���� = Replace(Replace(lvwMain_S.SelectedItem.SubItems(lvwMain_S.ColumnHeaders("_�������").Index - 1), "[", ""), "]", "")
    End If
    
    Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
        "����=" & IIF(lng����id = 0, "", lng����id), _
        "��Ŀ=" & IIF(lng��Ŀid = 0, "", lng��Ŀid), _
        "���=" & str�շ����)
End Sub

Private Sub mnuViewFind_Click()
    frmChargeItemFind.Show , Me
End Sub

Private Sub mnuViewRefresh_Click()
On Error GoTo ErrHandle
    FillTree
    Call FillWholeSetTree
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuViewShowAll_Click()
On Error GoTo ErrHandle
    mnuViewShowAll.Checked = Not mnuViewShowAll.Checked
    If tvwMainItem.SelectedItem Is Nothing Then
        If tvwMainItem.Nodes.Count > 0 Then
            MsgBox "��ѡ��һ�·��࣡", vbInformation, gstrSysName
        Else
            MsgBox "���κη������ʾ��", vbInformation, gstrSysName
        End If
        Exit Sub
    End If
    FillList tvwMainItem.SelectedItem.Tag
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuViewShowStop_Click()
On Error GoTo ErrHandle
    mnuViewShowStop.Checked = Not mnuViewShowStop.Checked
    If tvwMainItem.Nodes.Count > 0 Then
        FillList tvwMainItem.SelectedItem.Tag
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuShortMenu1_Click(Index As Integer)
On Error GoTo ErrHandle
    Select Case Index
        Case 1
            mnuEditParent_Click
        Case 2
            mnuEditModifyAssort_Click
        Case 3
            mnuEditDeleteAssort_Click
    End Select
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuShortMenu2_Click(Index As Integer)
On Error GoTo ErrHandle
    Select Case Index
        Case 0
            mnuEditChild_Click
        Case 1
            mnuEditCopy_Click
        Case 2
            mnuEditModify_Click
        Case 3
            mnuEditDept_Click
        Case 4
            mnuEditSlave_Click
        Case 5
            mnuEditDelete_Click
        Case 6
            mnuEditItemGroup_Click
        Case 7
            mnuPriceChargeSet_Click
    End Select
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuShortIcon_Click(Index As Integer)
On Error GoTo ErrHandle
    mnuViewIcon_Click Index
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuViewIcon_Click(Index As Integer)
On Error GoTo ErrHandle
    Dim i As Integer
    For i = 0 To 3
        mnuViewIcon(i).Checked = False
        Toolbar1.Buttons("View").ButtonMenus(i + 1).Text = Replace(Toolbar1.Buttons("View").ButtonMenus(i + 1).Text, "��", "  ")
    Next
    mnuViewIcon(Index).Checked = True
    Toolbar1.Buttons("View").ButtonMenus(Index + 1).Text = Replace(Toolbar1.Buttons("View").ButtonMenus(Index + 1).Text, "  ", "��")
    lvwMain_S.View = Index
    lvwWholeSetItem_S.View = Index
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuViewSelect_Click()
On Error GoTo ErrHandle
    If zlControl.LvwSelectColumns(lvwMain_S, mstrLvw) = True Then
        '���б仯��Ҫ����ˢ��
        If tvwMainItem.Nodes.Count > 0 Then
            FillList tvwMainItem.SelectedItem.Tag
        End If
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mshAlias_DblClick()
On Error GoTo ErrHandle
    If mnuEditModify.Enabled And mnuEditModify.Visible = True Then mnuEditModify_Click
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub msh����_DblClick()
On Error GoTo ErrHandle
    If mnuEditSlave.Enabled = True And mnuEditSlave.Visible Then Call mnuEditSlave_Click
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub msh�ѱ�_DblClick()
    Dim strCharge As String
    On Error GoTo ErrHandle
    
    If InStr(mstrPrivs, "�ѱ�����") = 0 Then Exit Sub
    
    If msh�ѱ�.Rows > 1 Then
        If msh�ѱ�.TextMatrix(msh�ѱ�.Rows - 1, 0) <> "" Then
            strCharge = msh�ѱ�.TextMatrix(msh�ѱ�.Row, 0)
        End If
    End If
    
    If mnuPriceChargeSet.Enabled = True Then
        If lvwMain_S.SelectedItem Is Nothing Then Exit Sub
        If frmChargeSortItemEdit.ShowMe(Me, 2, strCharge, Val(Me.lvwMain_S.SelectedItem.Tag), Me.lvwMain_S.SelectedItem.Text) Then
            Call frmChargeManage.FillTree
        End If
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub msh��Ŀ_DblClick()
On Error GoTo ErrHandle
    If mnuPriceRaise.Enabled = True Then Call mnuPriceRaise_Click
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub msh��Ŀ_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrHandle
    If Button = 2 Then
        If InStr(mstrPrivs, "��Ŀ����") > 0 Then PopupMenu mnuPrice, vbPopupMenuRightButton
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

 
Private Sub picTreeClass_S_Resize()
    Err = 0: On Error Resume Next
    With tbClassPage
        .Top = picTreeClass_S.ScaleTop
        .Height = picTreeClass_S.ScaleHeight
        .Left = picTreeClass_S.ScaleLeft
        .Width = picTreeClass_S.ScaleWidth
    End With
End Sub

Private Sub picTreeItem_Resize()
    Err = 0: On Error Resume Next
    With tvwMainItem
        .Top = picTreeItem.ScaleTop
        .Height = picTreeItem.ScaleHeight
        .Left = picTreeItem.ScaleLeft
        .Width = picTreeItem.ScaleWidth
    End With
End Sub

Private Sub picTreeWholeSet_Resize()
    Err = 0: On Error Resume Next
    With tvwWholeSet
        .Top = picTreeWholeSet.ScaleTop
        .Height = picTreeWholeSet.ScaleHeight
        .Left = picTreeWholeSet.ScaleLeft
        .Width = picTreeWholeSet.ScaleWidth
    End With
End Sub

Private Sub tbClassPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If mblnNotClick = True Then Exit Sub
    Select Case Val(Item.Tag)
    Case mCalssPage.pg_ϸĿ
        If Val(tbClassPage.Tag) <> mCalssPage.pg_ϸĿ Then
            '�ı���,��Ҫ���´�������
            '�ȱ����ϴ�ѡ�����
           ' SaveListViewState Me.lvwMain_S, Me.Name & "_" & Val(tbClassPage.Tag), lvwMain_S.View
            tbClassPage.Tag = mCalssPage.pg_ϸĿ: mstrKey = ""
            If Not tvwMainItem.SelectedItem Is Nothing Then
                 Call tvwMainItem_NodeClick(tvwMainItem.SelectedItem)
            End If
        End If
         If tvwMainItem.Enabled And tvwMainItem.Visible Then tvwMainItem.SetFocus
    Case mCalssPage.pg_����
        If Val(tbClassPage.Tag) <> mCalssPage.pg_���� Then
            '�ı���,��Ҫ���´�������
            '�ȱ����ϴ�ѡ�����
            'SaveListViewState Me.lvwMain_S, Me.Name & "_" & Val(tbClassPage.Tag), lvwMain_S.View
            tbClassPage.Tag = mCalssPage.pg_����: tvwWholeSet.Tag = ""
            If Not tvwWholeSet.SelectedItem Is Nothing Then
                 Call tvwWholeSet_NodeClick(tvwWholeSet.SelectedItem)
            End If
        End If
         If tvwWholeSet.Enabled And tvwWholeSet.Visible Then tvwWholeSet.SetFocus
    End Select
    Call SetPageVisible
    
End Sub

Private Sub tbPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If tbClassPage.Selected Is Nothing Then Exit Sub
    If Val(tbClassPage.Selected.Tag) = mCalssPage.pg_���� Then
        mint�ϴγ���ҳ = Val(Item.Tag)
    Else
        mint�ϴ�ϸĿҳ = Val(Item.Tag)
    End If
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
On Error GoTo ErrHandle
    Dim i As Integer
    For i = 0 To 3
        mnuViewIcon(i).Checked = False
    Next
    mnuViewIcon(ButtonMenu.Index - 1).Checked = True
    lvwMain_S.View = ButtonMenu.Index - 1
    lvwWholeSetItem_S.View = lvwMain_S.View
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Toolbar1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrHandle
    If Button = 2 Then
        PopupMenu mnuViewTool, 2
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub tvwMainItem_NodeClick(ByVal Node As MSComctlLib.Node)
On Error GoTo ErrHandle
    If Node Is Nothing Then Exit Sub
    If mstrKey = Node.Key Then Exit Sub
    mstrKey = Node.Key
    
    FillList Format(Node.Tag)
    If lvwMain_S.ListItems.Count > 0 Then
        If Not lvwMain_S.SelectedItem Is Nothing Then
            lvwMain_S_ItemClick lvwMain_S.SelectedItem
            Exit Sub
        End If
    End If
    
    msh��Ŀ.Clear 1
    msh��Ŀ.Rows = 2
    
    chk�۸�.value = 0
    opt����(0).value = 1
    opt����(1).value = 0
    opt����(2).value = 0
    opt����(3).value = 0
    opt����(4).value = 0
    opt����(5).value = 0
    opt����(6).value = 0
    lvwOutIn.ListItems.Clear
    msh����.Rows = 2
    msh����.TextMatrix(1, 0) = ""
    msh����.TextMatrix(1, 1) = ""
    msh����.TextMatrix(1, 2) = ""
    msh����.TextMatrix(1, 3) = ""
    
    mshAlias.Rows = 2
    mshAlias.TextMatrix(1, 0) = ""
    mshAlias.TextMatrix(1, 1) = ""
    mshAlias.TextMatrix(1, 2) = ""
    mshAlias.TextMatrix(1, 3) = ""
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub tvwMainItem_GotFocus()
    Call MenuSet
End Sub

Private Sub tvwMainItem_LostFocus()
    Call MenuSet
End Sub

Private Sub tvwMainItem_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrHandle
    If Button = 2 Then
        If mnuShortMenu1(1).Visible = False Then Exit Sub
        mnuShortMenu1(1).Enabled = mnuEditParent.Enabled
        mnuShortMenu1(2).Enabled = mnuEditModifyAssort.Enabled
        mnuShortMenu1(3).Enabled = mnuEditDeleteAssort.Enabled
        PopupMenu mnuShort1, vbPopupMenuRightButton
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    Dim sngTop As Single, sngBottom As Single
    On Error Resume Next
    sngTop = IIF(CoolBar1.Visible, CoolBar1.Top + CoolBar1.Height, 0) + Screen.TwipsPerPixelY
    sngBottom = Me.ScaleHeight - IIF(stbThis.Visible, stbThis.Height, 0)
    'picTreeClass_S��λ��
    picTreeClass_S.Top = sngTop
    picTreeClass_S.Height = IIF(sngBottom - picTreeClass_S.Top > 0, sngBottom - picTreeClass_S.Top, 0)
    picTreeClass_S.Left = 0
    'picSplit��λ��
    picSplit.Top = sngTop
    picSplit.Height = picTreeClass_S.Height
    picSplit.Left = picTreeClass_S.Left + picTreeClass_S.Width
    'lvwMain_S��λ��
    lvwMain_S.Top = sngTop - Screen.TwipsPerPixelY
    lvwMain_S.Left = picSplit.Left + picSplit.Width

    
    If Me.ScaleWidth - lvwMain_S.Left > 0 Then lvwMain_S.Width = Me.ScaleWidth - lvwMain_S.Left
    'picNS��λ��
    picNS.Left = lvwMain_S.Left
    picNS.Top = lvwMain_S.Top + lvwMain_S.Height
    picNS.Width = lvwMain_S.Width
    'picTreeClass_S��λ��
    tbPage.Left = lvwMain_S.Left
    tbPage.Top = picNS.Top + picNS.Height
    tbPage.Width = lvwMain_S.Width
    tbPage.Height = IIF(sngBottom - tbPage.Top > 0, sngBottom - tbPage.Top, 0)
    Me.picͣ��ԭ��.Left = Me.tbPage.Left + 4300
    Me.picͣ��ԭ��.Top = Me.tbPage.Top + 100
    lblͣ��ԭ��.ZOrder
    With lvwWholeSetItem_S
        .Top = lvwMain_S.Top
        .Left = lvwMain_S.Left
        .Width = lvwMain_S.Width
        .Height = lvwMain_S.Height
    End With
    lblMsg.Move opt����(1).Left + opt����(1).Width + 2500, opt����(1).Top
    CoolBar1.Bands(1).Width = Me.Width - 2000
    Me.Refresh
End Sub


Private Sub mnuEditChild_Click()
'������Ŀ
    On Error GoTo ErrHandle
    Dim strSQL As String

    If mnuEdit.Visible = False Then Exit Sub
    If tvwMainItem.SelectedItem Is Nothing Then Exit Sub
    If lvwMain_S.ListItems.Count > 0 And Not lvwMain_S.SelectedItem Is Nothing Then
        If IsNumeric(lvwMain_S.SelectedItem.Tag) Then
            If CLng(lvwMain_S.SelectedItem.Tag) > 0 Then
                Call frmChargeItem.�༭��Ŀ(mstrPrivs, mblnCanUpdateAll, "C" & IIF(mstrClass = "", " ", mstrClass) & tvwMainItem.SelectedItem.Tag, , , , mbln����ҽ��ϵͳ)
                Exit Sub
            End If
        End If
                
        Call frmChargeItem.�༭��Ŀ(mstrPrivs, mblnCanUpdateAll, "C" & IIF(mstrClass = "", " ", mstrClass) & tvwMainItem.SelectedItem.Tag, , , , mbln����ҽ��ϵͳ)
    Else
        '��ʱ�����Ϊ��
        Call frmChargeItem.�༭��Ŀ(mstrPrivs, mblnCanUpdateAll, "C " & tvwMainItem.SelectedItem.Tag, , , , mbln����ҽ��ϵͳ)
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuEditCopy_Click()
'��������
On Error GoTo ErrHandle
    If lvwMain_S.SelectedItem Is Nothing Then Exit Sub
    Call frmChargeItem.�༭��Ŀ(mstrPrivs, mblnCanUpdateAll, "C" & IIF(mstrClass = "", " ", mstrClass) & IIF(tvwMainItem.SelectedItem.Key = "Root", "", tvwMainItem.SelectedItem.Tag), lvwMain_S.SelectedItem.Tag, , 5, mbln����ҽ��ϵͳ)  'editCopy
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuEditParent_Click()
'��������
On Error GoTo ErrHandle
    With frmChargeSort
        If Me.tvwMainItem.SelectedItem Is Nothing Then
            .txtParent.Tag = 0
        Else
            .txtParent.Tag = Mid(Me.tvwMainItem.SelectedItem.Key, 2)
        End If
        .Tag = "����"
        If .ShowMe(1, Me) Then Call FillTree
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuEditModifyAssort_Click()
'�޸ķ���
Dim i As Long
Dim strSQL As String
Dim rsTmp As New ADODB.Recordset

On Error GoTo ErrHandle
    If tvwMainItem.SelectedItem Is Nothing Then Exit Sub
    With frmChargeSort
        .mblnCancel = True
        If Me.tvwMainItem.SelectedItem.Parent Is Nothing Then
            .txtParent.Tag = 0
            .txtParent.Text = "(��)"
            .txtUpCode.Text = ""
            .txtCode.Text = Mid(Split(Me.tvwMainItem.SelectedItem.Text, "]")(0), 2)
            .txtCode.MaxLength = Len(.txtCode.Text)
            .txtCode.Tag = .txtCode.MaxLength
        Else
            .txtParent.Tag = Mid(Me.tvwMainItem.SelectedItem.Parent.Key, 2)
            .txtParent.Text = Me.tvwMainItem.SelectedItem.Parent.Text
            .txtUpCode.Text = Mid(Split(Me.tvwMainItem.SelectedItem.Parent.Text, "]")(0), 2)
            .txtCode.Text = Mid(Split(Me.tvwMainItem.SelectedItem.Text, "]")(0), Len(.txtUpCode.Text) + 2)
            .txtCode.MaxLength = Len(.txtCode.Text)
            .txtCode.Tag = .txtCode.MaxLength
        End If
        .txtName = Split(Me.tvwMainItem.SelectedItem.Text, "]")(1)
        strSQL = "Select ���� from �շѷ���Ŀ¼ where id=[1] "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Me.tvwMainItem.SelectedItem.Tag))
        
        If rsTmp.RecordCount > 0 Then
            .txtSymbol = Nvl(rsTmp!����)
        Else
            .txtSymbol = ""
        End If
        .Tag = Mid(Me.tvwMainItem.SelectedItem.Key, 2)
        .mblnCancel = False
        If .ShowMe(1, Me) Then Call FillTree
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuEditModify_Click()
'�޸���Ŀ
On Error GoTo ErrHandle
    If mnuEdit.Visible = False Then Exit Sub
    If lvwMain_S.SelectedItem Is Nothing Then Exit Sub
    Call frmChargeItem.�༭��Ŀ(mstrPrivs, mblnCanUpdateAll, "C" & IIF(mstrClass = "", " ", mstrClass) & tvwMainItem.SelectedItem.Tag, lvwMain_S.SelectedItem.Tag, , 1, mbln����ҽ��ϵͳ) 'EditMode.editModify
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuPriceRaise_Click()
'����
On Error GoTo ErrHandle
    If mnuPrice.Visible = False Then Exit Sub
    ModifyMode 2    'editRaise
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ModifyMode(ByVal edit��ʽ As EditMode)
On Error GoTo ErrHandle
    If ActiveControl Is tvwMainItem And edit��ʽ < 2 Then  'editRaise
        With tvwMainItem.SelectedItem
            Call frmChargeItem.�༭��Ŀ(mstrPrivs, mblnCanUpdateAll, "C" & IIF(mstrClass = "", " ", mstrClass) & .Tag, .Tag, 0, edit��ʽ, mbln����ҽ��ϵͳ)
        End With
    Else
        If lvwMain_S.SelectedItem Is Nothing Then Exit Sub
        If checkNotPrice(lvwMain_S.SelectedItem.Tag) = False Then
            MsgBox "���շ�ϸĿ������δ��˵ļ۸�����˺��������ۣ�", vbInformation, gstrSysName
            Exit Sub
        End If
        Call frmChargeItem.�༭��Ŀ(mstrPrivs, mblnCanUpdateAll, "C" & IIF(mstrClass = "", " ", mstrClass) & IIF(tvwMainItem.SelectedItem.Key = "Root", "", tvwMainItem.SelectedItem.Tag), lvwMain_S.SelectedItem.Tag, , edit��ʽ, mbln����ҽ��ϵͳ)
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function checkNotPrice(ByVal lng�շ�ϸĿID As Long) As Boolean
    '����Ƿ񻹴���δ��Ч�ļ۸�
    Dim rsData As ADODB.Recordset
    Dim strWhere As String
    
    On Error GoTo ErrHandle
    If mblnCanUpdateAll = False Then
        strWhere = " And (b.վ��=[2]" & vbNewLine & _
                "       Or b.վ�� Is Null And a.�۸�ȼ� In(" & vbNewLine & _
                "           Select m.����" & vbNewLine & _
                "           From �շѼ۸�ȼ� M, �շѼ۸�ȼ�Ӧ�� N" & vbNewLine & _
                "           Where m.���� = n.�۸�ȼ� And Nvl(m.�Ƿ�������ͨ��Ŀ, 0) = 1 And n.վ�� = [2]" & vbNewLine & _
                "                 And (m.����ʱ�� Is Null Or m.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd'))))"
    End If
    If lng�շ�ϸĿID <> 0 Then
        gstrSQL = "Select 1 From �շѵ��ۼ�¼ A,�շ���ĿĿ¼ B Where a.�շ�ϸĿID = b.ID And a.��˱�־ = 0 And a.�շ�ϸĿid=[1]" & strWhere & " And Rownum < 2"
    Else
        gstrSQL = "Select 1 From �շѵ��ۼ�¼ A,�շ���ĿĿ¼ B" & _
                " Where a.�շ�ϸĿID = b.ID And a.��˱�־ = 0" & strWhere & " And Rownum < 2"
    End If
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "δ��Ч���ݲ�ѯ", lng�շ�ϸĿID, gstrNodeNo)
    If rsData.RecordCount > 0 Then
        checkNotPrice = False
    Else
        checkNotPrice = True
    End If
    
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub mnuEditDeleteAssort_Click()
'ɾ��
    Dim strKey As String
    Dim intIndex As Long
    
    On Error GoTo ErrHandle
    If tvwMainItem.SelectedItem Is Nothing Then Exit Sub
    
    If MsgBox("��ȷ��Ҫɾ������Ϊ��" & tvwMainItem.SelectedItem.Text & "���ķ�����", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
        Me.MousePointer = 11
        gstrSQL = "ZL_�շѷ���Ŀ¼_DELETE(" & tvwMainItem.SelectedItem.Tag & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        Me.MousePointer = 0
        
        strKey = tvwMainItem.SelectedItem.Key
        If Not tvwMainItem.SelectedItem.Next Is Nothing Then
            tvwMainItem.SelectedItem.Next.Selected = True
            tvwMainItem_NodeClick tvwMainItem.SelectedItem
        Else
            If Not tvwMainItem.SelectedItem.Parent Is Nothing Then
                tvwMainItem.SelectedItem.Parent.Selected = True
            End If
            If Not tvwMainItem.SelectedItem Is Nothing Then
                tvwMainItem_NodeClick tvwMainItem.SelectedItem
            End If
        End If
        tvwMainItem.Nodes.Remove strKey
    End If
    MenuSet
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Me.MousePointer = 0
End Sub

Private Sub mnuEditDelete_Click()
'ɾ��
    Dim strKey As String
    Dim intIndex As Long
    
    On Error GoTo ErrHandle
    If lvwMain_S.SelectedItem Is Nothing Then Exit Sub
    If MsgBox("��ȷ��Ҫɾ������Ϊ��" & lvwMain_S.SelectedItem.Text & "������Ŀ��", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
        Me.MousePointer = 11
        gstrSQL = "zl_�շ�ϸĿ_delete(" & lvwMain_S.SelectedItem.Tag & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        Me.MousePointer = 0
        
        With lvwMain_S
            intIndex = .SelectedItem.Index
            .ListItems.Remove .SelectedItem.Key
            If .ListItems.Count > 0 Then
                intIndex = IIF(.ListItems.Count > intIndex, intIndex, .ListItems.Count)
                .ListItems(intIndex).Selected = True
                .ListItems(intIndex).EnsureVisible
                FillItem .ListItems(intIndex).Key
            Else
                FillItem ""
            End If
        End With
    End If
    MenuSet
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Me.MousePointer = 0
End Sub

Private Sub mnuEditStart_Click()
    Dim strԭ�� As String
    On Error GoTo ErrHandle
    
    If lvwMain_S.SelectedItem Is Nothing Then Exit Sub
    
    mfrmEarnRS.ShowMe 1, strԭ��
    
    If strԭ�� = "" Then Exit Sub
    
    gstrSQL = "zl_�շ�ϸĿ_reuse(" & lvwMain_S.SelectedItem.Tag & ",'" & strԭ�� & "')"
    'ִ�����ù���
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    '�ı�ͼ�����ɫ
    With lvwMain_S.SelectedItem
        .Icon = "Item"
        .SmallIcon = "Item"
        .ForeColor = RGB(0, 0, 0)
        
        Dim i As Integer
        For i = 1 To lvwMain_S.ColumnHeaders.Count
            If i < lvwMain_S.ColumnHeaders.Count Then
                .ListSubItems(i).ForeColor = RGB(0, 0, 0)
            End If
            '���³���ʱ��
            If lvwMain_S.ColumnHeaders(i).Text = "����ʱ��" Then
                .SubItems(i - 1) = "3000-01-01"
            End If
        Next
    End With
    '�ı�״̬���Ͳ˵�
    MenuSet
    lvwMain_S_ItemClick lvwMain_S.SelectedItem
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuEditStop_Click()
    On Error GoTo ErrHandle
    Dim strKey As String
    Dim intIndex As Integer
    Dim strTmp As String
    Dim strԭ�� As String
    
    If lvwMain_S.SelectedItem Is Nothing Then Exit Sub
    
    strKey = lvwMain_S.SelectedItem.Tag
    
    If Not Check�շ���Ŀ(Val(strKey), strTmp) Then
        Exit Sub
    End If
    If strTmp <> "" Then
        If MsgBox("����Ŀ����������������ϵ��" & vbCrLf & strTmp & "�Ƿ�ͣ�ã�", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
    
    mfrmEarnRS.ShowMe 2, strԭ��
    
    If strԭ�� = "" Then Exit Sub
    
    gstrSQL = "zl_�շ�ϸĿ_stop(" & strKey & ",'" & strԭ�� & "')"
    'ִ�����ù���
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    '�ı�ͼ�����ɫ
    If mnuViewShowStop.Checked = True Then 'Ҫ��ʾͣ�ò���
        With lvwMain_S.SelectedItem
            .Icon = "ItemNo"
            .SmallIcon = "ItemNo"
            .ForeColor = RGB(255, 0, 0)
            
            Dim i As Integer
            For i = 1 To lvwMain_S.ColumnHeaders.Count
                If i < lvwMain_S.ColumnHeaders.Count Then
                    .ListSubItems(i).ForeColor = RGB(255, 0, 0)
                End If
                '���³���ʱ��
                If lvwMain_S.ColumnHeaders(i).Text = "����ʱ��" Then
                    .SubItems(i - 1) = Format(Date, "yyyy-MM-dd")
                End If
            Next
        End With
    Else '����ʾͣ�ò���
        With lvwMain_S
            intIndex = .SelectedItem.Index
            .ListItems.Remove .SelectedItem.Key
            If .ListItems.Count > 0 Then
                intIndex = IIF(.ListItems.Count > intIndex, intIndex, .ListItems.Count)
                .ListItems(intIndex).Selected = True
                .ListItems(intIndex).EnsureVisible
                FillItem .ListItems(intIndex).Key
            Else
                FillItem ""
            End If
        End With
    End If
    MenuSet
    lvwMain_S_ItemClick lvwMain_S.SelectedItem
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetCodeLength(ByVal strTable As String) As Long
'����:�ӱ��еõ��ֶεĳ���
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo ErrHandle
    
    GetCodeLength = 0
    gstrSQL = "select nvl(max(Vsize(����)),0) as LenCode from " & strTable
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If rsTmp.RecordCount > 0 Then
        GetCodeLength = rsTmp!lencode
    End If
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub picsplit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        msngStartX = X
    End If
End Sub

Private Sub picsplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    Dim sngTemp As Single
    If Button = 1 Then
        sngTemp = picSplit.Left + X - msngStartX
        If sngTemp > 500 And Me.ScaleWidth - (sngTemp + picSplit.Width) > 600 Then
            picSplit.Left = sngTemp
            picTreeClass_S.Width = picSplit.Left - picTreeClass_S.Left
            Form_Resize
        End If
    End If
End Sub

Private Sub picNS_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        msngStartY = Y
    End If
End Sub

Private Sub picNS_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    Dim sngTemp As Single
    If Button = 1 Then
        sngTemp = picNS.Top + Y - msngStartY
        If sngTemp > lvwMain_S.Top + 600 And Me.ScaleHeight - (sngTemp + picNS.Height) > 600 Then
            picNS.Top = sngTemp
            lvwMain_S.Height = picNS.Top - lvwMain_S.Top
            lvwWholeSetItem_S.Height = lvwMain_S.Height
            Form_Resize
        End If
    End If
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileExcel_Click()
    subPrint 3
End Sub

Private Sub mnufilepre_Click()
    subPrint 2
End Sub

Private Sub mnuFilePrint_Click()
    subPrint 1
End Sub

Private Sub mnufileset_Click()
    zlPrintSet
End Sub

'Private Sub tbPage_Click()
'    Dim i As Integer
'    tbPage.ZOrder 1
'    For i = 1 To 5
'        fra(i).ZOrder 1
'    Next
'    fra(tbPage.SelectedItem.Index).ZOrder
'End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Child"
            If tbClassPage.Selected Is Nothing Then Exit Sub
            If Val(tbClassPage.Selected.Tag) = mCalssPage.pg_���� Then
                '������Ŀ����
                Call mnuEditWholeSetItemAdd_Click
            Else
                 mnuEditChild_Click
            End If
        Case "Parent"
            If tbClassPage.Selected Is Nothing Then Exit Sub
            If Val(tbClassPage.Selected.Tag) = mCalssPage.pg_���� Then
                '���׷���
                Call mnuEditWholeSetClassAdd_Click
            Else
                mnuEditParent_Click
            End If
        Case "Modify"
            If tbClassPage.Selected Is Nothing Then Exit Sub
            If Val(tbClassPage.Selected.Tag) = mCalssPage.pg_���� Then
                '����
                If ActiveControl Is tvwWholeSet Then
                    mnuEditWholeSetClassModify_Click
                Else
                    mnuEditWholeSetItemModify_Click
                End If
            Else
                If ActiveControl Is tvwMainItem Then
                    If InStr(mstrPrivs, "������") > 0 Then
                        mnuEditModifyAssort_Click
                    End If
                Else
                    mnuEditModify_Click
                End If
            End If
        Case "Delete"
            If tbClassPage.Selected Is Nothing Then Exit Sub
            If Val(tbClassPage.Selected.Tag) = mCalssPage.pg_���� Then
                '����
                If ActiveControl Is tvwWholeSet Then
                    mnuEditWholeSetClassDelete_Click
                Else
                    mnuEditWholeSetItemDelete_Click
                End If
            Else
                If ActiveControl Is tvwMainItem Then
                    If InStr(mstrPrivs, "������") > 0 Then
                        mnuEditDeleteAssort_Click
                    End If
                Else
                    mnuEditDelete_Click
                End If
            End If
        Case "Raise"
            mnuPriceRaise_Click
        Case "RaiseVerify"
            frmChargePriceVerify.ShowMe Me, mblnCanUpdateAll
        Case "Start"
            mnuEditStart_Click
        Case "Stop"
            mnuEditStop_Click
        Case "Quit"
            mnuFileExit_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Find"
            mnuViewFind_Click
        Case "Preview"
            mnufilepre_Click
        Case "Help"
            mnuhelptopic_Click
        Case "View"
            mnuViewIcon(lvwMain_S.View).Checked = False
            If lvwMain_S.View = 3 Then
                mnuViewIcon(0).Checked = True
                lvwMain_S.View = 0
            Else
                mnuViewIcon(lvwMain_S.View + 1).Checked = True
                lvwMain_S.View = lvwMain_S.View + 1
            End If
    End Select
    
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    CoolBar1.Visible = mnuViewToolButton.Checked
    CoolBar1.Bands("only").MinHeight = Toolbar1.Height
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim buttTemp As Button
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For Each buttTemp In Toolbar1.Buttons
        If mnuViewToolText.Checked Then
            buttTemp.Caption = buttTemp.Tag
        Else
            buttTemp.Caption = ""
        End If
    Next
    CoolBar1.Bands("only").MinHeight = Toolbar1.Height
    Form_Resize
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = mnuViewStatus.Checked
    Form_Resize
End Sub

Private Sub mnuhelptopic_Click()
      ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hwnd)
End Sub
Private Sub subPrint(bytMode As Byte)
    '����:���д�ӡ,Ԥ���������EXCEL
    '����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Dim objPrint As New zlPrintLvw
    If gstrUserName = "" Then Call GetUserInfo
    If tbClassPage.Selected Is Nothing Then Exit Sub
    If Val(tbClassPage.Selected.Tag) = mCalssPage.pg_���� Then
        If tvwWholeSet.SelectedItem Is Nothing Then Exit Sub
        If lvwWholeSetItem_S.ListItems.Count = 0 Then Exit Sub
        objPrint.Title.Text = "�����շ���Ŀ"
        Set objPrint.Body.objData = lvwWholeSetItem_S
        objPrint.UnderAppItems.Add "���ࣺ" & tvwWholeSet.SelectedItem.Text
    Else
        If tvwMainItem.SelectedItem Is Nothing Then Exit Sub
        If lvwMain_S.ListItems.Count = 0 Then Exit Sub
        objPrint.Title.Text = "�շ���Ŀ"
        Set objPrint.Body.objData = lvwMain_S
        objPrint.UnderAppItems.Add "���ࣺ" & tvwMainItem.SelectedItem.Text
    End If
    objPrint.BelowAppItems.Add "��ӡ�ˣ�" & gstrUserName
    objPrint.BelowAppItems.Add "��ӡʱ�䣺" & Format(sys.Currentdate, "yyyy��MM��dd��")
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrViewLvw objPrint, 1
          Case 2
              zlPrintOrViewLvw objPrint, 2
          Case 3
              zlPrintOrViewLvw objPrint, 3
      End Select
    Else
        zlPrintOrViewLvw objPrint, bytMode
    End If
End Sub
Public Sub FillWholeSetTree()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����׷�������
    '����:���˺�
    '����:2010-08-24 14:55:07
    '˵��:27327
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String, objNode As Node
    Dim strPreKey As String
    Err = 0: On Error GoTo ErrHand:
    strSQL = "" & _
    "   Select id,�ϼ�ID,����,���� " & _
    "   From ������Ŀ����  " & _
    "   Start with �ϼ�id is null Connect by Prior   Id=�ϼ�ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With tvwWholeSet
        If Not .SelectedItem Is Nothing Then strPreKey = .SelectedItem.Key
        .Nodes.Clear
       Set objNode = .Nodes.Add(, , "Root", "���г���", "RootS", "Exp")
       objNode.Expanded = True
       objNode.Sorted = True
       Do While Not rsTemp.EOF
            If IsNull(rsTemp!�ϼ�id) Then
                Set objNode = .Nodes.Add("Root", tvwChild, "K" & Nvl(rsTemp!ID), Nvl(rsTemp!����) & "-" & Nvl(rsTemp!����), "RootS", "Exp")
            Else
                Set objNode = .Nodes.Add("K" & rsTemp!�ϼ�id, tvwChild, "K" & Nvl(rsTemp!ID), Nvl(rsTemp!����) & "-" & Nvl(rsTemp!����), "RootS", "Exp")
            End If
            objNode.Sorted = True
            If objNode.Key = strPreKey Then
                objNode.EnsureVisible
                objNode.Selected = True
                objNode.Expanded = True
            End If
            objNode.Sorted = True
            rsTemp.MoveNext
       Loop
       tvwWholeSet.Tag = ""
       If .SelectedItem Is Nothing Then .Nodes("Root").Selected = True
       If Not .SelectedItem Is Nothing Then
            
            Call tvwWholeSet_NodeClick(.SelectedItem)
       End If
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Sub
Private Function FillWholeItem(ByVal lng����id As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������Ŀ
    '���:lng����id-����ID,0-���з���
    '����:
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-08-25 15:41:48
    '����:27327
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String, strWhere As String, strOwner As String
    Dim strPreKey As String, objListItem As ListItem, lngCol As Long
    On Error GoTo ErrHandle
    
    Screen.MousePointer = vbHourglass
    Me.stbThis.Panels(2).Text = "���ڶ�ȡ�շѳ�����Ŀ�б�����,���Ժ� ������"
    Me.stbThis.Refresh
    
    If Not tbClassPage.Selected Is Nothing Then
        If Not lvwWholeSetItem_S.SelectedItem Is Nothing And Val(tbClassPage.Selected.Tag) = mCalssPage.pg_���� Then
            strPreKey = lvwWholeSetItem_S.SelectedItem.Key
        End If
    End If
    
    strSQL = "select ������ from zlsystems where ���=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��������", glngSys)
    If rsTemp.RecordCount = 1 Then
        strOwner = IIF(IsNull(rsTemp!������), "", rsTemp!������)
    End If
    rsTemp.Close
    
    If strOwner <> gstrDbUser Then
        strWhere = " And ( A.��ԱID=[2] "
        If InStr(1, mstrPrivs, ";���Ƴ��׷���;") > 0 Then
            strWhere = strWhere & " OR Exists(Select 1 From ������Ŀʹ�ÿ��� A1 ,������Ա B1 Where A1.����ID=A.ID And A1.����ID=B1.����Id and B1.��Աid=[2]) "
        End If
        If InStr(1, mstrPrivs, ";ȫԺ���׷���;") > 0 Then
            strWhere = strWhere & " OR nvl(A.��Χ,0)=0 "
        End If
        strWhere = strWhere & ")"
    End If
    
    strSQL = "" & _
    "   Select  A.Id,A.����ID,A.����,A.����,A.ƴ��,A.���,decode(nvl(��Χ,0),0,'ȫԺ',1,'ָ������',decode(A.��Աid,Null,'ָ������Ա',B.����)) As ʹ�÷�Χ," & _
    "              C.���� as �������� " & _
    "   From �����շ���Ŀ A,��Ա�� B " & _
            IIF(lng����id = 0, ",������Ŀ���� C", " ,(Select ID,�ϼ�ID,����,���� From  ������Ŀ����  Start With Id =[1] Connect By Prior Id=�ϼ�id ) C") & _
    "   Where a.��Աid=b.Id(+) And A.����id=C.ID " & strWhere
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����id, glngUserId)
    zlControl.FormLock lvwWholeSetItem_S.hwnd
    mblnNotClick = True
    With lvwWholeSetItem_S
        .ListItems.Clear
        Do While Not rsTemp.EOF
            '��ӽڵ�
            Set objListItem = .ListItems.Add(, "K" & rsTemp!ID, Nvl(rsTemp!����) & "-" & Nvl(rsTemp!����), "Item", "Item")
            objListItem.Tag = Nvl(rsTemp!����id)
            ' "����,1500,0,1;����,800,0,2;����,1400,0,0;ʹ�÷�Χ,400,0,0;��������,2400,0,0"
            '����ListView�����������ݿ�ȡ��
            For lngCol = 2 To lvwWholeSetItem_S.ColumnHeaders.Count
                objListItem.SubItems(lngCol - 1) = Nvl(rsTemp.Fields(lvwWholeSetItem_S.ColumnHeaders(lngCol).Text))
            Next
            If rsTemp.AbsolutePosition = 1 Then 'ȱʡΪ��һ��ѡ��
                objListItem.Selected = True
            End If
            If objListItem.Key = strPreKey Then
                objListItem.Selected = True
                objListItem.EnsureVisible
            End If
            rsTemp.MoveNext
        Loop
        If .ListItems.Count > 0 Then
                Me.stbThis.Panels(2).Text = "�շѳ�����Ŀ���ݶ�ȡ��ɣ�"
        Else
                Me.stbThis.Panels(2).Text = ""
        End If
    End With
    mblnNotClick = False
    lvwWholeSetItem_S.Tag = ""
    
    If Not lvwWholeSetItem_S.SelectedItem Is Nothing Then
        Call lvwWholeSetItem_S_ItemClick(lvwWholeSetItem_S.SelectedItem)
    Else
        '���������Ŀ��һЩ����
        Call zlClearDownWholeSetItem
    End If
    zlControl.FormLock 0
    Screen.MousePointer = vbDefault
    FillWholeItem = True
    Exit Function
ErrHandle:
    Screen.MousePointer = vbDefault
    If ErrCenter() = 1 Then
        Screen.MousePointer = vbHourglass
        Resume
    End If
    Me.stbThis.Panels(2).Text = ""
    Me.stbThis.Refresh
    mblnNotClick = False
    zlControl.FormLock 0
End Function
Private Function FillWholeSetItemChildData(ByVal lng����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���س�����Ŀ������
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-08-25 17:42:49
    '˵��:27327
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim objListItem As ListItem, lng���� As Long, j As Long, i As Long
    Dim strWherePriceGrade As String
    
    On Error GoTo ErrHandle

    'strSQL = "" & _
    "   Select '' as ��־,A.���, A.����id, A.�շ�ϸĿid, B.����, B.����, B.���㵥λ, B.���, A.��������, A.����, A.����, A.ִ�п���id, " & _
    "          decode(C.����,NULL,'',C.����||'-') ||C.���� As ִ�п��� " & _
    "   From �����շ���Ŀ��� A, �շ���ĿĿ¼ B, ���ű� C " & _
    "   Where A.�շ�ϸĿid = B.ID And A.ִ�п���id = C.ID(+)  And A.����id = [1]" & _
    "   Order By A.���"

    If gstr��ͨ�۸�ȼ� = "" And gstrҩƷ�۸�ȼ� = "" And gstr���ļ۸�ȼ� = "" Then
        strWherePriceGrade = " And j.�۸�ȼ� Is Null"
    Else
        strWherePriceGrade = "" & _
            " And ((Instr(';5;6;7;', ';' || k.��� || ';') > 0 And j.�۸�ȼ� = [2])" & vbNewLine & _
            "      Or (Instr(';4;', ';' || k.��� || ';') > 0 And j.�۸�ȼ� = [3])" & vbNewLine & _
            "      Or (Instr(';4;5;6;7;', ';' || k.��� || ';') = 0 And j.�۸�ȼ� = [4])" & vbNewLine & _
            "      Or (j.�۸�ȼ� Is Null" & vbNewLine & _
            "          And Not Exists (Select 1" & vbNewLine & _
            "                          From �շѼ�Ŀ" & vbNewLine & _
            "                          Where j.�շ�ϸĿid = �շ�ϸĿid And Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
            "                                And ((Instr(';5;6;7;', ';' || k.��� || ';') > 0 And �۸�ȼ� = [2])" & vbNewLine & _
            "                                      Or (Instr(';4;', ';' || k.��� || ';') > 0 And �۸�ȼ� = [3])" & vbNewLine & _
            "                                      Or (Instr(';4;5;6;7;', ';' || k.��� || ';') = 0 And �۸�ȼ� = [4])))))"
    End If
    
    gstrSQL = "" & _
    "   Select  /*+Rule */ A.����ID,A.�շ�ϸĿID,A.���,A.��������,A.����,A.����,A.����,A.ִ�п���ID, " & _
    "              B.���,B.����,B.����,B.���㵥λ,B.���,C.��ҩ��̬,D.���� as ���Ʊ���, " & _
    "              D.���� as ��������,D.���㵥λ as ������λ,C.����ϵ��, " & _
    "              E.���� As ִ�п��ұ���,E.���� As ִ�п�������, " & _
    "              M.���� As ���ױ���,M.���� As ��������,M.ƴ��,M.���,M.��ע,M.��Χ, " & _
    "              M.����ID,M.��ԱID,G.����,J.���� As �������,J.���� As �������� ,B.�Ƿ���,B.ִ�п���,C.ҩ��ID," & _
    "              Decode(B.�Ƿ���,1,'ʱ��',LTrim(To_Char(J1.�ּ�,'999999999.9999999'))) as �ּ�  " & _
    "   From �����շ���Ŀ M,������Ŀ���� J,�����շ���Ŀ��� A,�շ���ĿĿ¼ B,ҩƷ��� C,������ĿĿ¼ D, " & _
    "             ���ű� E,��Ա�� G," & _
    "             (Select j.�շ�ϸĿid, Sum(j.�ּ�) as �ּ�" & vbNewLine & _
    "              From �շѼ�Ŀ J,�շ���ĿĿ¼ K" & vbNewLine & _
    "              Where j.�շ�ϸĿID = k.ID And Sysdate Between J.ִ������ And Nvl(J.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD'))" & vbNewLine & _
                         strWherePriceGrade & vbNewLine & _
    "              Group By j.�շ�ϸĿid ) J1 " & _
    "   Where   M.����ID=J.Id And  M.��ԱID=G.Id(+) And M.Id=A.����ID(+)  " & _
    "               And A.�շ�ϸĿid=b.Id(+)  And a.�շ�ϸĿID=C.ҩƷID(+) And C.ҩ��ID=D.Id(+) " & _
    "               And A.�շ�ϸĿid=J1.�շ�ϸĿID(+)  And A.ִ�п���ID=E.Id(+)  " & _
    "               And M.ID=[1] Order by A.���"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng����ID, gstrҩƷ�۸�ȼ�, gstr���ļ۸�ȼ�, gstr��ͨ�۸�ȼ�)
    With vsWholeSet
        .redraw = flexRDNone
        .Clear 1
        .OutlineBar = flexOutlineBarSimple
        .OutlineCol = .ColIndex("��־"): .SubtotalPosition = flexSTAbove
        .Rows = IIF(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        i = 1
        lng���� = 0
        Do While Not rsTemp.EOF
            If Val(Nvl(rsTemp!��������)) = 0 Then
                    lng���� = Nvl(rsTemp!���)
            End If
            .TextMatrix(i, .ColIndex("���")) = Nvl(rsTemp!���)
            .TextMatrix(i, .ColIndex("��������")) = Nvl(rsTemp!��������)
            .TextMatrix(i, .ColIndex("�շ���Ŀ")) = Nvl(rsTemp!����) & "-" & Nvl(rsTemp!����)
            .TextMatrix(i, .ColIndex("���")) = Nvl(rsTemp!���)
            .TextMatrix(i, .ColIndex("ȱʡ����")) = IIF(Val(Nvl(rsTemp!����)) = 0, 1, Val(Nvl(rsTemp!����)))
            .TextMatrix(i, .ColIndex("ȱʡ����")) = FormatEx(Val(Nvl(rsTemp!����)), 5)
            .TextMatrix(i, .ColIndex("ȱʡ�۸�")) = FormatEx(Val(Nvl(rsTemp!����)), 8)
            .TextMatrix(i, .ColIndex("ȱʡִ�п���")) = Nvl(rsTemp!ִ�п�������)
            .TextMatrix(i, .ColIndex("��λ")) = Nvl(rsTemp!���㵥λ)
            .TextMatrix(i, .ColIndex("�ּ�")) = IIF(Nvl(rsTemp!�ּ�) = "ʵ��", "ʵ��", FormatEx(Val(Nvl(rsTemp!�ּ�)), 5))
            If Nvl(rsTemp!���) = "7" Then
                '��ҩ,��ʾ��������
                .TextMatrix(.Row, .ColIndex("ҩ��")) = Nvl(rsTemp!���Ʊ���) & "-" & Nvl(rsTemp!��������)
                .TextMatrix(.Row, .ColIndex("��λ")) = Nvl(rsTemp!������λ)
                .TextMatrix(.Row, .ColIndex("ȱʡ����")) = FormatEx(Val(.TextMatrix(.Row, .ColIndex("ȱʡ����"))) * Val(Nvl(rsTemp!����ϵ��)), 5)
                .TextMatrix(.Row, .ColIndex("ȱʡ�۸�")) = FormatEx(Val(.TextMatrix(.Row, .ColIndex("ȱʡ�۸�"))) / Val(Nvl(rsTemp!����ϵ��)), 8)
              '  .TextMatrix(.Row, .ColIndex("��ҩ��̬")) = Val(Nvl(rsTemp!��ҩ��̬))
            End If
        
            If Val(Nvl(rsTemp!��������)) = 0 Then
                    .IsSubtotal(i) = True
                    .RowOutlineLevel(i) = 1
            ElseIf lng���� = Val(.TextMatrix(i, .ColIndex("��������"))) Then
                    If i > 2 Then
                        If Val(.TextMatrix(i - 1, .ColIndex("��������"))) <> 0 Then
                            .IsSubtotal(i - 1) = False
                            .RowOutlineLevel(i - 1) = 1
                        End If
                    End If
                    .IsSubtotal(i) = True
                    .RowOutlineLevel(i) = 2
            End If
            i = i + 1
            rsTemp.MoveNext
        Loop
        
        zl_vsGrid_Para_Restore mlngMode, vsWholeSet, Me.Caption, "������Ŀ��ɱ���-������", True, True
        .redraw = flexRDBuffered
    End With
    strSQL = "" & _
    "   Select A.����ID,B.����,b.���� " & _
    "   From ������Ŀʹ�ÿ��� A,���ű�  B  " & _
    "   Where a.����id=b.Id And a.����ID=[1]" & _
    "   Order By ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
   With lvwUseDept
        .ListItems.Clear
        Do While Not rsTemp.EOF
            .ListItems.Add , "K" & rsTemp!����ID, Nvl(rsTemp!����) & "-" & Nvl(rsTemp!����), "Dept", "Dept"
            rsTemp.MoveNext
        Loop
   End With
    FillWholeSetItemChildData = True
    Exit Function
ErrHandle:
    vsWholeSet.redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
        vsWholeSet.redraw = flexRDNone
    End If
End Function
Private Sub zlClearDownWholeSetItem()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ָ��������Ŀ����ɺ�ʹ�ÿ�������
    '����:���˺�
    '����:2010-08-25 16:35:03
    '����:27327
    '---------------------------------------------------------------------------------------------------------------------------------------------

    With vsWholeSet
        .Rows = 2
        .Clear 1
    End With
    With lvwUseDept
        .ListItems.Clear
    End With
End Sub

Public Sub FillTree()
'����:װ���շ������շ�ϸĿ�����з��ൽtvwMainItem
    '�����������ڵ�����������KEYֵ��һ���ַ������ڶ�λ��������

    Dim rs���� As New ADODB.Recordset
    Dim strTemp As String
    Dim strKey As String
    Dim i As Long
    Dim objNode As Node
    
    mstrKey = ""     'ȫ��ˢ��ʱ���൱���û�û����κνڵ�
    
    If Not tvwMainItem.SelectedItem Is Nothing Then
    '��¼��ǰ�Ľڵ�
        strKey = tvwMainItem.SelectedItem.Key
    End If
    
    Screen.MousePointer = vbHourglass
    
    Me.stbThis.Panels(2).Text = "���ڶ�ȡ��������,���Ժ� ������"
    Me.stbThis.Refresh
    On Error GoTo ErrHandle
    zlControl.FormLock tvwMainItem.hwnd
    
    tvwMainItem.Nodes.Clear
    tvwMainItem.Sorted = False
    
    '��ʾ����
    gstrSQL = _
        "Select ID,�ϼ�id,����,����,���� " & vbCrLf & _
        "From �շѷ���Ŀ¼" & vbCrLf & _
        "Start With �ϼ�id Is Null" & vbCrLf & _
        "Connect By Prior id=�ϼ�id "
        
    rs����.CursorLocation = adUseClient
    Call zlDatabase.OpenRecordset(rs����, gstrSQL, Me.Caption)
    If rs����.RecordCount > 0 Then
        With rs����
            .MoveFirst
            For i = 0 To .RecordCount - 1
                If IsNull(rs����!�ϼ�id) Then
                    Set objNode = tvwMainItem.Nodes.Add(, , "R" & rs����!ID, "[" & rs����("����") & "]" & rs����("����"), "RootS", "Exp")
                Else
                    Set objNode = tvwMainItem.Nodes.Add("R" & rs����!�ϼ�id, tvwChild, "R" & rs����!ID, "[" & rs����("����") & "]" & rs����("����"), "RootS", "Exp")
                End If
                'objNode.ExpandedImage = "Exp"
                objNode.Tag = rs����!ID
                objNode.Sorted = True
                .MoveNext
            Next
        End With
        Me.stbThis.Panels(2).Text = ""
    Else
        Me.stbThis.Panels(2).Text = "���κη���!"
    End If
    tvwMainItem.Sorted = True
    
    zlControl.FormLock 0
    
    Dim nod As Node
    On Error Resume Next
    Set nod = tvwMainItem.Nodes(strKey)
    If Err <> 0 Then
        Set nod = tvwMainItem.Nodes(1)
        nod.Selected = True
        nod.Expanded = True
        tvwMainItem_NodeClick nod
    Else
        Err.Clear
        nod.Selected = True
        nod.Expanded = True
        nod.EnsureVisible
        tvwMainItem_NodeClick nod
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    zlControl.FormLock 0
End Sub

Public Sub FillList(ByVal str���� As String)
'����:װ���Ӧ�������Ŀ��lvwMain_S
'����:str���� ����ı�ʶ
    Dim rs��Ŀ As New ADODB.Recordset
    Dim strTemp As String
    Dim i As Integer
    Dim lst As ListItem
    Dim strKey As String
    Dim str���� As String
    Dim j As Long
    
    On Error GoTo errHandleList
    
    If Not lvwMain_S.SelectedItem Is Nothing Then
        '����ԭ�м�ֵ
        strKey = lvwMain_S.SelectedItem.Key
    End If
    rs��Ŀ.CursorLocation = adUseClient
    rs��Ŀ.CursorType = adOpenKeyset
    rs��Ŀ.LockType = adLockReadOnly
    
    Screen.MousePointer = vbHourglass
    Me.stbThis.Panels(2).Text = "���ڶ�ȡ�շ���Ŀ�б�����,���Ժ� ������"
    Me.stbThis.Refresh
    If mnuViewShowStop.Checked = False Then
        strTemp = " and (A.����ʱ�� = to_date('3000-01-01','YYYY-MM-DD') or A.����ʱ�� is null) "
    End If
    If mblnCanUpdateAll = False Then
        strTemp = strTemp & " And (a.վ�� Is Null Or a.վ�� = [2])"
    End If
    If mnuViewShowAll.Checked = True Then
            gstrSQL = _
                "Select A.ID,A.����ID,A.���,C.���� �������,C.�̶� ���̶�,A.����,A.��ʶ����,A.��ʶ����,A.����޼�,A.����޼�,A.��ѡ��,A.����,A.���,A.���㵥λ,A.��������, " & vbCrLf & _
                "       Decode(A.�������,1,'����',2,'סԺ',3,'������סԺ','��') as �������," & vbCrLf & _
                "       decode(A.����ժҪ,1,'��','') as ����ժҪ, A.˵��,decode(A.���ηѱ�,1,'��','') as ���ηѱ�," & vbCrLf & _
                "       decode(A.�Ƿ���,1,'��','') as �Ƿ���,decode(A.�Ӱ�Ӽ�,1,'��','') as �Ӱ�Ӽ�,decode(A.ִ�п���,1,4,2,1,3,3,0) AS ִ�п���, " & vbCrLf & _
                "       to_char(A.����ʱ��,'yyyy-mm-dd') as ����ʱ��,to_char(A.����ʱ��,'yyyy-mm-dd') as ����ʱ��," & vbCrLf & _
                "       decode(A.���,'1',decode(A.��Ŀ����,1,'������Ŀ','�Һ���Ŀ'),'H',decode(A.��Ŀ����,1,'����ȼ�',2,'��������','')) As ��Ŀ����," & vbCrLf & _
                "        Nvl(B.����,'') As ��������,a.������Ŀ,d.��� As վ����, d.���� As վ������" & vbCrLf & _
                " From " & vbCrLf & _
                "   (Select Id,���� From �շѷ���Ŀ¼" & vbCrLf & _
                "   Start With �ϼ�id  = [1]" & vbCrLf & _
                "    Connect By Prior id=�ϼ�id) B," & vbCrLf & _
                "    �շ���ĿĿ¼ A,�շ���Ŀ��� C, zlnodelist D " & vbCrLf & _
                "Where  A.���=C.���� And a.վ�� = d.���(+) and A.����id=B.Id And A.���<>'5'  And  A.���<>'6'  And  A.���<>'7'" & strTemp & vbCrLf & _
                "Union" & vbCrLf & _
                "Select A.ID,A.����ID,A.���,C.���� �������,C.�̶� ���̶�,A.����,A.��ʶ����,A.��ʶ����,A.����޼�,A.����޼�,A.��ѡ��,A.����,A.���,A.���㵥λ,A.��������, " & vbCrLf & _
                "       Decode(A.�������,1,'����',2,'סԺ',3,'������סԺ','��') as �������," & vbCrLf & _
                "       decode(A.����ժҪ,1,'��','') as ����ժҪ, A.˵��,decode(A.���ηѱ�,1,'��','') as ���ηѱ�," & vbCrLf & _
                "       decode(A.�Ƿ���,1,'��','') as �Ƿ���,decode(A.�Ӱ�Ӽ�,1,'��','') as �Ӱ�Ӽ�,decode(A.ִ�п���,1,4,2,1,3,3,0) AS ִ�п���, " & vbCrLf & _
                "       to_char(A.����ʱ��,'yyyy-mm-dd') as ����ʱ��,to_char(A.����ʱ��,'yyyy-mm-dd') as ����ʱ��," & vbCrLf & _
                "       decode(A.���,'1',decode(A.��Ŀ����,1,'������Ŀ','�Һ���Ŀ'),'H',decode(A.��Ŀ����,1,'����ȼ�',2,'��������','')) As ��Ŀ����," & vbCrLf & _
                "        Nvl(B.����,'') As ��������,a.������Ŀ,d.��� As վ����, d.���� As վ������" & vbCrLf & _
                " From �շ���ĿĿ¼ A,�շѷ���Ŀ¼ B,�շ���Ŀ��� C, zlnodelist D " & vbCrLf & _
                "Where A.���=C.���� And a.վ�� = d.���(+) and A.����id  = [1] And A.����id=B.ID And A.���<>'5'  And  A.���<>'6'  And  A.���<>'7'" & strTemp
    Else
            gstrSQL = _
                "Select A.ID,A.����ID,A.���,C.���� �������,C.�̶� ���̶�,A.����,A.��ʶ����,A.��ʶ����,A.����޼�,A.����޼�,A.��ѡ��,A.����,A.���,A.���㵥λ,A.��������, " & vbCrLf & _
                "       decode(A.�������,1,'����',2,'סԺ',3,'������סԺ','��') as �������," & vbCrLf & _
                "       decode(A.����ժҪ,1,'��','') as ����ժҪ, A.˵��,decode(A.���ηѱ�,1,'��','') as ���ηѱ�," & vbCrLf & _
                "       decode(A.�Ƿ���,1,'��','') as �Ƿ���,decode(A.�Ӱ�Ӽ�,1,'��','') as �Ӱ�Ӽ�,decode(A.ִ�п���,1,4,2,1,3,3,0) AS ִ�п���, " & vbCrLf & _
                "       to_char(A.����ʱ��,'yyyy-mm-dd') as ����ʱ��,to_char(A.����ʱ��,'yyyy-mm-dd') as ����ʱ��," & vbCrLf & _
                "       decode(A.���,'1',decode(A.��Ŀ����,1,'������Ŀ','�Һ���Ŀ'),'H',decode(A.��Ŀ����,1,'����ȼ�',2,'��������','')) As ��Ŀ����," & vbCrLf & _
                "        Nvl(B.����,'') As ��������,a.������Ŀ,d.��� As վ����, d.���� As վ������" & vbCrLf & _
                "From �շ���ĿĿ¼ A  ,�շѷ���Ŀ¼ B,�շ���Ŀ��� C, zlnodelist D " & vbCrLf & _
                "Where A.���=C.���� And a.վ�� = d.���(+) and A.����id=B.Id And  A.���<>'5'  And A.���<>'6' And A.���<>'7'" & strTemp & " And  A.����id  = [1] "
    End If
    
    Set rs��Ŀ = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(str����), gstrNodeNo)
    'Call zlInitLvwHeadCol(1 )
    zlControl.FormLock lvwMain_S.hwnd
    With lvwMain_S.ListItems
        .Clear
        '����;����;���;���㵥λ;��������;�������;˵��;���ηѱ�;�Ƿ���;�Ӱ�Ӽ�;����ժҪ;��Ŀ����;����ʱ��;����ʱ��;��������
        If rs��Ŀ.RecordCount > 0 Then
            rs��Ŀ.MoveFirst
            Dim lngCol  As Long
            Dim varValue As Variant
            For i = 0 To rs��Ŀ.RecordCount - 1
                '�ó���ȷ��ͼ��
                strTemp = "Item"
                If Not CDate(IIF(IsNull(rs��Ŀ("����ʱ��")), CDate("3000/1/1"), rs��Ŀ("����ʱ��"))) = CDate("3000/1/1") Then
                    strTemp = strTemp & "No"
                End If
                '��ӽڵ�
                Set lst = .Add(, "C" & rs��Ŀ("���") & rs��Ŀ("id"), rs��Ŀ("����"), strTemp, strTemp)
                If InStr(strTemp, "No") > 0 Then lst.ForeColor = RGB(255, 0, 0)
                lst.Tag = rs��Ŀ!ID
                
                '����ListView�����������ݿ�ȡ��
                For lngCol = 2 To lvwMain_S.ColumnHeaders.Count
                    If Trim(lvwMain_S.ColumnHeaders(lngCol).Text) = "�������" And rs��Ŀ!���̶� = 1 Then
                            varValue = "[" & rs��Ŀ(lvwMain_S.ColumnHeaders(lngCol).Text).value & "]"
                    ElseIf Trim(lvwMain_S.ColumnHeaders(lngCol).Text) = "����޼�" Or Trim(lvwMain_S.ColumnHeaders(lngCol).Text) = "����޼�" Then
                        If IsNull(rs��Ŀ(lvwMain_S.ColumnHeaders(lngCol).Text).value) Then
                            varValue = " "
                        Else
                            varValue = CStr(Format(rs��Ŀ(lvwMain_S.ColumnHeaders(lngCol).Text).value, "0.00"))
                        End If
                    ElseIf lvwMain_S.ColumnHeaders(lngCol).Text = "Ժ��" Then
                        varValue = rs��Ŀ("վ������").value
                    Else
                        varValue = rs��Ŀ(lvwMain_S.ColumnHeaders(lngCol).Text).value
                    End If
                    lst.SubItems(lngCol - 1) = IIF(IsNull(varValue), "", varValue)
                    If InStr(strTemp, "No") > 0 Then lst.ListSubItems(lngCol - 1).ForeColor = RGB(255, 0, 0)
                Next
                '��¼����Ŀ�����
                lst.ListSubItems(2).Tag = Nvl(rs��Ŀ("���"))
                '��¼����Ŀ��վ��
                lst.ListSubItems(lst.ListSubItems.Count).Tag = Nvl(rs��Ŀ("վ����"))
                rs��Ŀ.MoveNext
            Next
        End If
    End With
    If rs��Ŀ.RecordCount > 0 Then
        Me.stbThis.Panels(2).Text = "�շ���Ŀ���ݶ�ȡ��ɣ�"
        Me.mnuFileExcel.Enabled = True
        Me.mnuFilePrint.Enabled = True
        Me.mnuFilepre.Enabled = True
        Me.mnuFilePrint.Enabled = True
    Else
        Me.stbThis.Panels(2).Text = "��ǰ��������Ŀ"
        Me.mnuFileExcel.Enabled = False
        Me.mnuFilePrint.Enabled = False
        Me.mnuFilepre.Enabled = False
        Me.mnuFilePrint.Enabled = False
    End If
    Toolbar1.Buttons("Print").Enabled = mnuFileExcel.Enabled
    Toolbar1.Buttons("Preview").Enabled = mnuFileExcel.Enabled
    If Me.ActiveControl Is tvwMainItem Then
        If tvwMainItem.SelectedItem Is Nothing Then
            mnuEditModifyAssort.Enabled = False
            mnuEditDeleteAssort.Enabled = False
        End If
        mnuEditModify.Enabled = False
        mnuEditDelete.Enabled = False
        mnuEditStart.Enabled = False
        mnuEditStop.Enabled = mnuEditStart.Enabled
        mnuEditDept.Enabled = mnuEditStart.Enabled
        mnuEditSlave.Enabled = mnuEditStart.Enabled
        mnuEditItemGroup.Enabled = mnuEditStart.Enabled
        mnuPriceChargeSet.Enabled = mnuEditStart.Enabled
        mnuPriceHistory.Enabled = mnuEditStart.Enabled
        mnuPriceRaise.Enabled = mnuEditStart.Enabled
        mnuEditCopy.Enabled = mnuEditStart.Enabled
        
        Toolbar1.Buttons("Modify").Enabled = True
        Toolbar1.Buttons("Delete").Enabled = True
        Toolbar1.Buttons("Start").Enabled = mnuEditStart.Enabled
        Toolbar1.Buttons("Stop").Enabled = mnuEditStop.Enabled
    Else
        If lvwMain_S.ListItems.Count > 0 Then
            tbPage.Enabled = True
            
            Dim Item As ListItem
            On Error Resume Next
            Set Item = lvwMain_S.ListItems(strKey)
            If Err <> 0 Then
                Set Item = lvwMain_S.ListItems(1)
                Item.Selected = True
                Item.EnsureVisible
                lvwMain_S_ItemClick Item
            Else
                Err.Clear
                Item.Selected = True
                Item.EnsureVisible
                lvwMain_S_ItemClick Item
            End If
        Else
            mnuEditStart.Enabled = False
            mnuEditStop.Enabled = mnuEditStart.Enabled
            mnuEditDelete.Enabled = mnuEditStart.Enabled
            mnuEditDept.Enabled = mnuEditStart.Enabled
            mnuEditModify.Enabled = mnuEditStart.Enabled
            mnuEditSlave.Enabled = mnuEditStart.Enabled
            mnuEditItemGroup.Enabled = mnuEditStart.Enabled
            mnuPriceChargeSet.Enabled = mnuEditStart.Enabled
            mnuPriceHistory.Enabled = mnuEditStart.Enabled
            mnuPriceRaise.Enabled = mnuEditStart.Enabled
            mnuEditCopy.Enabled = mnuEditStart.Enabled
            
            Toolbar1.Buttons("Modify").Enabled = mnuEditStart.Enabled
            Toolbar1.Buttons("Delete").Enabled = mnuEditStart.Enabled
            Toolbar1.Buttons("Start").Enabled = mnuEditStart.Enabled
            Toolbar1.Buttons("Stop").Enabled = mnuEditStop.Enabled
            
            FillItem ""
        End If
    End If
    zlControl.FormLock 0
    Screen.MousePointer = vbDefault
    Exit Sub
errHandleList:
    Screen.MousePointer = vbDefault
    If ErrCenter() = 1 Then
        Screen.MousePointer = vbHourglass
        Resume
    End If
    Call SaveErrLog
    zlControl.FormLock 0
End Sub

Public Sub FillItem(ByVal str��Ŀ As String)
'����:��ʾĩ���շ�ϸĿ�ļ�Ŀ,������Ŀ��ִ�п���
'����:str��Ŀ ��Ŀ�ı�ʶ
    Dim rsTemp As New ADODB.Recordset
    Dim strID As String
    Dim lst As ListItem
    Dim i As Integer, j As Integer
    Dim datCurr As Date
    Dim strSQL As String
    Dim iRow As Integer, icol As Integer
    Dim strTmp As String, str�۸�ȼ� As String
    Dim ObjItem As ListItem
    
    On Error GoTo ErrHandle
    
    MenuSet
    If str��Ŀ = "" Then
        mstr�ϼ�Key = ""
        tbPage.Enabled = False
        
        msh��Ŀ.Clear 1
        msh��Ŀ.Rows = 2
        
        msh����.Rows = 2
        For i = 0 To msh����.Cols - 1
            msh����.TextMatrix(1, i) = ""
        Next
        mshAlias.Rows = 2
        For i = 0 To mshAlias.Cols - 1
            mshAlias.TextMatrix(1, i) = ""
        Next
        opt����(4).value = True
        
        If Mid(tvwMainItem.SelectedItem.Key, 2, 1) = "F" Then
            msh��Ŀ.ColWidth(Col_���������շ���) = 1500
            msh��Ŀ.TextMatrix(0, Col_���������շ���) = "���������շ���"
        Else
            msh��Ŀ.ColWidth(Col_���������շ���) = 0
        End If
    Else
        tbPage.Enabled = True
    End If
    rsTemp.CursorLocation = adUseClient
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockReadOnly
    
    If lvwMain_S.ListItems.Count = 0 Then Exit Sub
    Set lst = lvwMain_S.ListItems(str��Ŀ)
    strID = Mid(str��Ŀ, 3)
    
    '�������
    If Mid(str��Ŀ, 2, 1) = "F" Then
        msh��Ŀ.ColWidth(Col_���������շ���) = 1500
        msh��Ŀ.TextMatrix(0, Col_���������շ���) = "���������շ���"
    Else
        msh��Ŀ.ColWidth(Col_���������շ���) = 0
    End If
    
    gstrSQL = "select a.�Ƿ���,a.�Ӱ�Ӽ�,a.ִ�п���,b.����,b.���� ���,a.����ID,a.����ԭ��,a.ͣ��ԭ��,a.����ʱ�� from �շ���ĿĿ¼   A,�շ���Ŀ��� B  where   a.���=b.����  AND a.ID=[1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(strID))
    
    If Not CDate(IIF(IsNull(rsTemp("����ʱ��")), CDate("3000/1/1"), rsTemp("����ʱ��"))) = CDate("3000/1/1") Then
        If Nvl(rsTemp!ͣ��ԭ��) = "" Then
            Me.picͣ��ԭ��.Visible = False
        Else
            Me.picͣ��ԭ��.Visible = True
            Me.lblͣ��ԭ��.Caption = "ͣ��ԭ��" & rsTemp!ͣ��ԭ��
        End If
    Else
        If Nvl(rsTemp!����ԭ��) = "" Then
            Me.picͣ��ԭ��.Visible = False
        Else
            Me.picͣ��ԭ��.Visible = True
            Me.lblͣ��ԭ��.Caption = "����ԭ��" & rsTemp!����ԭ��
        End If
    End If
    
    If rsTemp.RecordCount > 0 Then
        If IsNull(rsTemp("����ID")) Then
            mstr�ϼ�Key = "R" & rsTemp("���")
        Else
            mstr�ϼ�Key = "C" & rsTemp("���") & rsTemp("����ID")
        End If
        mstrClass = Nvl(rsTemp!����)
        mstrClassName = Nvl(rsTemp!���)
    Else
        mstrClass = ""
        mstrClassName = ""
        mstr�ϼ�Key = ""
        MsgBox "����Ŀ�����ڣ�", vbExclamation, gstrSysName
        Call mnuViewRefresh_Click
        Exit Sub
    End If
    
    If rsTemp("�Ƿ���") = 1 Then
        msh��Ŀ.Rows = 2
        msh��Ŀ.TextMatrix(0, Col_ԭ��) = "����޼�"
        msh��Ŀ.TextMatrix(0, Col_�ּ�) = "����޼�"
        msh��Ŀ.ColWidth(Col_ȱʡ�۸�) = 1000
    Else
        msh��Ŀ.TextMatrix(0, Col_ԭ��) = "ԭ��"
        msh��Ŀ.TextMatrix(0, Col_�ּ�) = "�ּ�"
        msh��Ŀ.ColWidth(Col_ȱʡ�۸�) = 0
    End If
    
    If rsTemp("�Ӱ�Ӽ�") = 1 Then
        msh��Ŀ.ColWidth(Col_�Ӱ�Ӽ���) = 1500
        msh��Ŀ.TextMatrix(0, Col_�Ӱ�Ӽ���) = "�Ӱ�Ӽ���"
    Else
        msh��Ŀ.ColWidth(Col_�Ӱ�Ӽ���) = 0
    End If
    '��ʾ����
    opt����(IIF(rsTemp("ִ�п���") < 7, rsTemp("ִ�п���"), 0)).value = True
    lvwOutIn.ListItems.Clear
    
    rsTemp.Close
    If opt����(4).value = True Then
        gstrSQL = " Select  " & _
            "   decode(b.����,null,'','['||b.����||']'|| b.����) As ��������,  " & vbCrLf & _
            "    '['||c.����||']'|| c.���� As ִ�п���" & vbCrLf & _
            " from �շ�ִ�п��� A,���ű� B,���ű� C" & vbCrLf & _
            " Where a.��������id=b.Id(+) And a.ִ�п���id=C.id and ������Դ is null and A.�շ�ϸĿid=[1] " & _
            " order by  c.���� "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(strID))
                
        Do Until rsTemp.EOF
            If strTmp <> rsTemp!ִ�п��� Then
                i = i + 1
                Set ObjItem = Me.lvwOutIn.ListItems.Add(, "A" & i, rsTemp!ִ�п���)
                ObjItem.SubItems(1) = IIF(IsNull(rsTemp!��������), "�����в��ţ�", rsTemp!��������)
            Else
                Me.lvwOutIn.ListItems("A" & i).SubItems(1) = Me.lvwOutIn.ListItems("A" & i).SubItems(1) & "," & rsTemp!��������
            End If
            strTmp = rsTemp!ִ�п���
            rsTemp.MoveNext
        Loop
        If lvwOutIn.ListItems.Count > 0 Then
            lvwOutIn.ListItems(1).Selected = True
            lvwOutIn.ListItems(1).EnsureVisible
        End If
    ElseIf opt����(0).value = True Then
        '����ȷִ�п�����ʾ�����õ��ֹ�����ȱʡִ�п���
        gstrSQL = "" & _
            " Select '[' || b.���� || ']' || b.���� As ִ�п���" & vbNewLine & _
            " From �շ�ִ�п��� A, ���ű� B" & vbNewLine & _
            " Where a.ִ�п���id = b.Id And a.������Դ = 2 And a.�շ�ϸĿid = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(strID))
        If Not rsTemp.EOF Then
            lvwOutIn.ListItems.Add , , Nvl(rsTemp!ִ�п���)
        End If
    End If
    
    '��ʾ�շѼ�Ŀ
    Call Fill��Ŀ(Val(strID))
    
    '��ʾ������Ŀ
    gstrSQL = "select a.����ID,a.����ID,a.���д���,a.��������,b.����,b.���� ��Ŀ����,c.����,c.���� ���," & _
        "Nvl(B.����ʱ��,to_Date('3000-01-01','YYYY-MM-DD')) As ����ʱ�� from �շѴ�����Ŀ a,�շ���ĿĿ¼ b ,�շ���Ŀ��� c where c.����=b.��� and a.����ID=b.id and ����ID=[1] " & _
        " ORDER BY a.ROWID "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(strID))
        
    msh����.Rows = IIF(rsTemp.RecordCount = 0, 2, rsTemp.RecordCount + 1)
    If rsTemp.RecordCount = 0 Then
        For i = 0 To 3
            msh����.TextMatrix(1, i) = ""
        Next
    Else
        i = 1
        Do Until rsTemp.EOF
            msh����.TextMatrix(i, 0) = "(" & rsTemp("����") & ")" & rsTemp("���")
            msh����.TextMatrix(i, 1) = "[" & rsTemp("��Ŀ����") & "]" & rsTemp("����")
            msh����.TextMatrix(i, 2) = rsTemp("��������")
            If rsTemp("���д���") = 0 Then
                msh����.TextMatrix(i, 3) = "0-���̶�"
            ElseIf rsTemp("���д���") = 2 Then
                msh����.TextMatrix(i, 3) = "2-����������"
            Else
                msh����.TextMatrix(i, 3) = "1-�̶�"
            End If
            msh����.TextMatrix(i, 4) = IIF(Format(rsTemp!����ʱ��, "YYYY-MM-DD") <> "3000-01-01", "ͣ��", "")

            iRow = msh����.Row: icol = msh����.Col
            msh����.Row = i
            For j = 0 To msh����.Cols - 1
                msh����.Col = j
                msh����.CellForeColor = IIF(Format(rsTemp!����ʱ��, "YYYY-MM-DD") <> "3000-01-01", &HFF&, vbBlack)
            Next
            msh����.Row = iRow: msh����.Col = icol
            
            i = i + 1
            rsTemp.MoveNext
        Loop
    End If
    '��ʾ����
    gstrSQL = "select decode( ����,1,'����',2,'Ӣ����',3,'������',4,'��ѧ��',5,'��Ʒ��',9,'��������','') ��������,����," & _
        "   decode(����,1,'ƴ����',2,'�����',3,'������','')  ����,nvl(����,'') ����" & _
        " from �շ���Ŀ���� where �շ�ϸĿID=[1]" & _
        " order by �������� ,���� "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(strID))
        
    mshAlias.Clear
    mshAlias.Rows = IIF(rsTemp.RecordCount = 0, 2, rsTemp.RecordCount + 1)
    
    If rsTemp.RecordCount = 0 Then
        For i = 0 To 3
            mshAlias.TextMatrix(1, i) = ""
        Next
    Else
        mshAlias.TextMatrix(0, 0) = "��������"
        mshAlias.TextMatrix(0, 1) = "����"
        mshAlias.TextMatrix(0, 2) = "����"
        mshAlias.TextMatrix(0, 3) = "����"
        
        i = 1
        Do Until rsTemp.EOF
            mshAlias.TextMatrix(i, 0) = rsTemp("��������")
            mshAlias.TextMatrix(i, 1) = rsTemp("����")
            mshAlias.TextMatrix(i, 2) = rsTemp("����")
            mshAlias.TextMatrix(i, 3) = Nvl(rsTemp("����"))
            i = i + 1
            rsTemp.MoveNext
        Loop
        mshAlias.ColAlignment(0) = flexAlignLeftCenter
        mshAlias.ColAlignment(1) = 4
        mshAlias.ColAlignment(2) = flexAlignLeftCenter
        mshAlias.ColAlignment(3) = flexAlignLeftCenter
        mshAlias.MergeCells = flexMergeRestrictColumns
        mshAlias.MergeCol(0) = True
        mshAlias.MergeCol(1) = True
    End If
    
    '��ʾ�ѱ�ȼ�
    gstrSQL = "Select A.�ѱ�, A.�κ�, Ӧ�ն���ֵ, Ӧ�ն�βֵ, ʵ�ձ���, Decode(���㷽��, 1, '1-�ɱ��ۼ��ձ�������', '0-�ֶα�������') As ���㷽�� " & _
            " From �ѱ���ϸ A, �շ���ĿĿ¼ B " & _
            " Where A.�շ�ϸĿid = B.ID And A.�շ�ϸĿid = [1] " & _
            " Order By A.�ѱ�, A.�κ�, A.Ӧ�ն���ֵ"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(strID))
    
    msh�ѱ�.Clear
    msh�ѱ�.Rows = IIF(rsTemp.RecordCount = 0, 2, rsTemp.RecordCount + 1)
    
    With msh�ѱ�
        If rsTemp.RecordCount = 0 Then
            .TextMatrix(0, 0) = "�ѱ�"
            .TextMatrix(0, 1) = "Ӧ�ս��(Ԫ)"
            .TextMatrix(0, 2) = "ʵ�ձ���(%)"
            .TextMatrix(0, 3) = "���㷽��"
            For i = 0 To 3
                msh�ѱ�.TextMatrix(1, i) = ""
            Next
        Else
            .TextMatrix(0, 0) = "�ѱ�"
            .TextMatrix(0, 1) = "Ӧ�ս��(Ԫ)"
            .TextMatrix(0, 2) = "ʵ�ձ���(%)"
            .TextMatrix(0, 3) = "���㷽��"
            
            i = 1
            Do Until rsTemp.EOF
                .TextMatrix(i, 0) = rsTemp.Fields("�ѱ�").value
                .TextMatrix(i, 1) = Format(rsTemp.Fields("Ӧ�ն���ֵ").value, "##########0.00;-#########0.00;0.00;0.00") & _
                    " �� " & Format(rsTemp.Fields("Ӧ�ն�βֵ").value, "##########0.00;-#########0.00;0.00;0.00")
                .TextMatrix(i, 2) = Format(rsTemp.Fields("ʵ�ձ���").value, "###0.000;-##0.000;0.000;0.000")
                .TextMatrix(i, 3) = rsTemp.Fields("���㷽��").value
                i = i + 1
                rsTemp.MoveNext
            Loop
            .ColAlignment(1) = 1
            .ColAlignment(2) = 1
            .MergeCells = flexMergeRestrictColumns
            .MergeCol(0) = True
        End If
    End With
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    msh����.redraw = True
    msh��Ŀ.redraw = flexRDBuffered
    mshAlias.redraw = True
End Sub

Private Function Fill��Ŀ(ByVal lngϸĿID As Long) As Boolean
    '��ʾ�շѼ�Ŀ
    Dim rsTemp As New ADODB.Recordset
    Dim i As Integer
    Dim datCurr As Date
    Dim str�۸�ȼ� As String, strWhere As String
    
    On Error GoTo ErrHandle
    With msh��Ŀ
        .redraw = flexRDNone
        .Clear 1
        .Rows = 2
        .Cell(flexcpBackColor, 1, 0, 1, .Cols - 1) = &H80000005
        .Subtotal flexSTClear
        
        datCurr = sys.Currentdate
        If mblnCanUpdateAll Then
            strWhere = "      And (a.�۸�ȼ� Is Null Or Exists (Select 1 From �շѼ۸�ȼ�Ӧ�� Where �۸�ȼ� = a.�۸�ȼ�))"
        Else
            strWhere = _
                "      And (a.�۸�ȼ� Is Null" & vbNewLine & _
                "           Or Exists(Select 1" & vbNewLine & _
                "               From �շ���ĿĿ¼ M, �շѼ۸�ȼ�Ӧ�� N" & vbNewLine & _
                "               Where m.Id = a.�շ�ϸĿid And c.���� = n.�۸�ȼ� And n.վ�� = [2]))"
        End If
        gstrSQL = "" & _
                "Select a.�۸�ȼ�, a.No, a.Id, a.ԭ��id, a.������Ŀid, a.ԭ��, a.�ּ�, Nvl(a.ȱʡ�۸�, 0) As ȱʡ�۸�," & vbNewLine & _
                "       a.�շ�ϸĿid, b.����, a.�Ӱ�Ӽ���, a.�����շ���, a.�䶯ԭ��, a.����˵��, a.ִ������, a.��ֹ����, a.������" & vbNewLine & _
                "From �շѼ�Ŀ A, ������Ŀ B, �շѼ۸�ȼ� C" & vbNewLine & _
                "Where a.������Ŀid = b.Id And a.�۸�ȼ� = c.����(+) And a.�շ�ϸĿid = [1]" & vbNewLine & _
                    IIF(chk�۸�.value = 1, "", "And (a.��ֹ���� Is Null Or a.��ֹ���� > Sysdate)") & vbNewLine & _
                "      And (c.����ʱ�� Is Null Or c.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd'))" & vbNewLine & _
                    strWhere & vbNewLine & _
                "Order By Nvl(c.����, ' '), a.ִ������ Desc"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngϸĿID, gstrNodeNo)
        
        .Rows = IIF(rsTemp.RecordCount = 0, 2, rsTemp.RecordCount + 1)
        i = 1
        Do Until rsTemp.EOF
            If InStr("," & str�۸�ȼ� & ",", "," & Nvl(rsTemp!�۸�ȼ�, "ȱʡ") & ",") = 0 Then
                str�۸�ȼ� = str�۸�ȼ� & "," & Nvl(rsTemp!�۸�ȼ�, "ȱʡ")
            End If
            .TextMatrix(i, Col_�۸�ȼ�) = Nvl(rsTemp!�۸�ȼ�, "ȱʡ")
            .TextMatrix(i, Col_���ݺ�) = Nvl(rsTemp!NO)
            .TextMatrix(i, Col_ִ������) = Format(Nvl(rsTemp!ִ������), "yyyy-MM-dd hh:mm:ss")
            .TextMatrix(i, Col_��ֹ����) = Format(Nvl(rsTemp!��ֹ����), "yyyy-MM-dd hh:mm:ss")
            .TextMatrix(i, Col_������Ŀ) = Nvl(rsTemp!����)
            .TextMatrix(i, Col_ԭ��) = Format(Nvl(rsTemp!ԭ��), "###########0.000;-##########0.000; ; ")
            .TextMatrix(i, Col_�ּ�) = Format(Nvl(rsTemp!�ּ�), "###########0.000;-##########0.000; ; ")
            .TextMatrix(i, Col_���������շ���) = Val(Nvl(rsTemp!�����շ���))
            .TextMatrix(i, Col_�Ӱ�Ӽ���) = Val(Nvl(rsTemp!�Ӱ�Ӽ���))
            .TextMatrix(i, Col_����˵��) = Nvl(rsTemp!����˵��)
            .TextMatrix(i, Col_ȱʡ�۸�) = Format(Nvl(rsTemp!ȱʡ�۸�), "###########0.000;-##########0.000; ; ")
            .TextMatrix(i, Col_������) = Nvl(rsTemp!������)
            .RowData(i) = rsTemp("ID")
            '�����ǲ������м۸�
            If rsTemp("ִ������") <= datCurr Then
                If CDate(Nvl(rsTemp!��ֹ����, "3000-01-01")) >= datCurr Then
                    .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &HDFFFFF
                End If
            End If
            i = i + 1
            rsTemp.MoveNext
        Loop
        '������ʾ
        If str�۸�ȼ� <> "" Then str�۸�ȼ� = Mid(str�۸�ȼ�, 2)
        If UBound(Split(str�۸�ȼ�, ",")) <= 0 Then
            .ColHidden(Col_�۸�ȼ�) = True
        Else
            .ColHidden(Col_�۸�ȼ�) = False
            .OutlineBar = flexOutlineBarComplete
            .MultiTotals = True
    
            .Subtotal flexSTNone, Col_�۸�ȼ�, , , , , True, "%s", , True
            .SubtotalPosition = flexSTAbove
    
            .Outline Col_�۸�ȼ�
            .OutlineCol = Col_�۸�ȼ�
    
            .MergeCells = flexMergeRestrictRows
            .MergeRow(-1) = False
            
            For i = 1 To .Rows - 1
                If .IsSubtotal(i) Then
                    .Cell(flexcpText, i, 0, i, .Cols - 1) = .TextMatrix(i + 1, Col_�۸�ȼ�)
                    .MergeRow(i) = True '���кϲ�
                    .IsCollapsed(i) = flexOutlineExpanded  '�Ƿ�չ��״̬
                End If
            Next
        End If
        .redraw = flexRDBuffered
    End With
    Fill��Ŀ = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    msh��Ŀ.redraw = flexRDBuffered
End Function

Private Sub chk�۸�_Click()
    Dim rsTemp As New ADODB.Recordset
    Dim strID As String
    Dim i As Integer
    Dim datCurr As Date
    
    On Error GoTo ErrHandle
    msh��Ŀ.Clear 1
    msh��Ŀ.Rows = 2
    msh��Ŀ.Cell(flexcpBackColor, 1, 0, 1, msh��Ŀ.Cols - 1) = &H80000005
    msh��Ŀ.Subtotal flexSTClear
    
    If lvwMain_S.SelectedItem Is Nothing Then Exit Sub
    strID = Mid(lvwMain_S.SelectedItem.Key, 3)
    Call Fill��Ŀ(Val(strID))
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub SetPageVisible()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ó�����Ŀҳ�����ݻ�����ϸĿ����ҳ����ʾ
    '����:���˺�
    '����:2010-08-27 16:39:39
    '����:27327
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnShowWholeSet As Boolean  '�Ƿ������ʾ
    Dim i As Long
    If tbClassPage.Selected Is Nothing Then
        blnShowWholeSet = False
    Else
        blnShowWholeSet = Val(tbClassPage.Selected.Tag) <> mCalssPage.pg_ϸĿ
    End If
    With tbPage
        For i = 0 To .ItemCount - 1
            If Val(.Item(i).Tag) = mItemPage.pg_����ʹ�ÿ��� Or Val(.Item(i).Tag) = mItemPage.pg_������� Then
                .Item(i).Visible = blnShowWholeSet
                If Val(.Item(i).Tag) = mint�ϴγ���ҳ And .Item(i).Visible Then
                        .Item(i).Selected = True
                End If
            Else
                .Item(i).Visible = Not blnShowWholeSet
                If Val(.Item(i).Tag) = mint�ϴ�ϸĿҳ And .Item(i).Visible Then
                        .Item(i).Selected = True
                End If
            End If
        Next
    End With

End Sub
Private Sub zlSetWholeSetMenu()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ó�����Ŀ����ز˵�
    '����:���˺�
    '����:2010-08-25 17:01:26
    '����:27327
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '�����ò˵�,�������еķǳ�����Ŀ�༭
    Dim blnAdd As Boolean, blnModify As Boolean, blnDelete As Boolean
    Dim blnEdit As Boolean
    If Val(tbClassPage.Selected.Tag) = mCalssPage.pg_ϸĿ Then
        '��ϸĿ,��������Ȩ��
        mnuEdit.Visible = True: mnuPrice.Visible = True
        mnuViewSelect.Visible = True: mnuViewShowStop.Visible = True
        mnuViewFind.Visible = True
        Call Ȩ�޿���
        mnuEditWholeSet.Visible = False
        lvwWholeSetItem_S.Visible = False
        lvwMain_S.Visible = True
        mnuPriceRaiseVerify.Visible = mblnVerifyFlow
        Toolbar1.Buttons("RaiseVerify").Visible = mblnVerifyFlow   '�������
        If mblnVerifyPris = False And mnuPriceRaiseVerify.Visible = True Then
            mnuPriceRaiseVerify.Enabled = False
            Toolbar1.Buttons("RaiseVerify").Enabled = False
        End If
        Exit Sub
    End If
    lvwWholeSetItem_S.Visible = True
    lvwMain_S.Visible = False
    
    mnuEdit.Visible = False: mnuPrice.Visible = False
    '�в���,��������
    mnuViewSelect.Visible = False: mnuViewShowStop.Visible = False  'û����ʾͣ����
    mnuPriceRaiseVerify.Visible = mblnVerifyFlow
    mnuViewFind.Visible = False 'û�в���
    blnAdd = InStr(1, mstrPrivs, ";���ӳ�����Ŀ;") > 0
    blnModify = InStr(1, mstrPrivs, ";�޸ĳ�����Ŀ;") > 0
    blnDelete = InStr(1, mstrPrivs, ";ɾ��������Ŀ;") > 0
    If blnAdd Or blnModify Or blnDelete Then
        mnuEditWholeSet.Visible = True
        mnuEditWholeSetClassAdd.Visible = blnAdd
        mnuEditWholeSetClassModify.Visible = blnModify
        mnuEditWholeSetClassDelete.Visible = blnDelete
        mnuEditWholeSetItemAdd.Visible = blnAdd
        mnuEditWholeSetItemModify.Visible = blnModify
        mnuEditWholeSetItemDelete.Visible = blnDelete
        Toolbar1.Buttons("Split4").Visible = True
    Else
        mnuEditWholeSet.Visible = False
        Toolbar1.Buttons("Split4").Visible = False
    End If
    Toolbar1.Buttons("Parent").Visible = blnAdd
    Toolbar1.Buttons("Child").Visible = blnAdd
    Toolbar1.Buttons("Child").Enabled = blnAdd
    Toolbar1.Buttons("Modify").Visible = blnModify
    Toolbar1.Buttons("Delete").Visible = blnDelete
    
    Toolbar1.Buttons("Split1").Visible = False
    Toolbar1.Buttons("Raise").Visible = False   '����
    Toolbar1.Buttons("RaiseVerify").Visible = mblnVerifyFlow   '�������
    
    If mblnVerifyPris = False And mnuPriceRaiseVerify.Visible = True Then
        mnuPriceRaiseVerify.Enabled = False
        Toolbar1.Buttons("RaiseVerify").Enabled = False
    End If
        
    Toolbar1.Buttons("Split2").Visible = False
    Toolbar1.Buttons("Start").Visible = False   '����
    Toolbar1.Buttons("Stop").Visible = False   'ͣ��
    Toolbar1.Buttons("Split3").Visible = False
    Toolbar1.Buttons("Find").Visible = False  '����
    Toolbar1.Buttons("Split4").Visible = False
    '�����Ƿ�ɱ༭
    blnEdit = Not lvwWholeSetItem_S.SelectedItem Is Nothing
    
    mnuEditWholeSetItemModify.Enabled = blnEdit
    mnuEditWholeSetItemDelete.Enabled = blnEdit
    If Me.ActiveControl Is tvwWholeSet Then
       '��ǰѡ�е��Ƿ���
       If tvwWholeSet.SelectedItem Is Nothing Then
            blnEdit = False
       Else
            blnEdit = Val(Mid(tvwWholeSet.SelectedItem.Key, 2)) <> 0
       End If
       mnuEditWholeSetClassModify.Enabled = blnEdit
       mnuEditWholeSetClassDelete.Enabled = blnEdit
       Toolbar1.Buttons("Modify").Enabled = blnEdit
       Toolbar1.Buttons("Delete").Enabled = blnEdit
       With tvwWholeSet
            If .SelectedItem Is Nothing Then
                stbThis.Panels(2).Text = ""
            Else
                stbThis.Panels(2).Text = "�÷��๲��" & .SelectedItem.Children & "���¼����ࡣ"
            End If
        End With
       Exit Sub
    End If
    With lvwWholeSetItem_S
        stbThis.Panels(2).Text = "������Ŀ�б��й���ʾ��" & .ListItems.Count & "��������Ŀ��"
    End With
    Toolbar1.Buttons("Modify").Enabled = mnuEditWholeSetItemModify.Enabled
    Toolbar1.Buttons("Delete").Enabled = mnuEditWholeSetItemDelete.Enabled
End Sub
Public Sub MenuSet()
'����:��ʾ�˵��͹�������״̬
'�������
'   ���Ƕ����༭ʱ,������ɾ�����Ҳ������������Ŀ����ɾ��
'   ����ϵͳ��־ʱ,������ɾ����𡢵�����������Ŀ����ɾ��
'���ڷ��ࣺ
'   ��Զ������ɾ����
'������Ŀ��
'   �����Ƕ����༭ʱ�������޸ĺ�ɾ��
'   �����Ǵ���ͣ��ʱ���޸�Ҳ������
'   ��������Ϳ�Ȩ�޵Ŀ�����
On Error GoTo ErrHandle
    Dim blnClassModify As Boolean, blnItemModify As Boolean '��������޸ġ�ɾ��
    Dim blnPrice As Boolean '������ֵ��ۡ��������۵�
    Dim blnStart As Boolean '���õ�״̬
    Dim blnStop As Boolean 'ͣ�õ�״̬
    Dim blnPrint As Boolean '��ӡ
    Dim blnCanModify As Boolean
    
    '���˺�:27327
    Call zlSetWholeSetMenu
    If Val(tbClassPage.Selected.Tag) = mCalssPage.pg_���� Then Exit Sub
    
    If ActiveControl Is tvwMainItem Then
        With tvwMainItem
            If Not tvwMainItem.SelectedItem Is Nothing Then
                blnClassModify = True
                stbThis.Panels(2).Text = "�÷��๲��" & .SelectedItem.Children & "���¼����ࡣ"
            End If
        End With
    Else
        With lvwMain_S
            stbThis.Panels(2).Text = "��Ŀ�б��й���ʾ��" & .ListItems.Count & "����Ŀ��"
            blnPrint = .ListItems.Count > 0
            
            If Not .SelectedItem Is Nothing Then
                blnCanModify = mblnCanUpdateAll _
                    Or .SelectedItem.ListSubItems(.SelectedItem.ListSubItems.Count).Tag = gstrNodeNo
            
                If InStr(.SelectedItem.Icon, "No") > 0 Then
                    'ͣ��
                    blnStop = (.SelectedItem.Icon = "ItemNo") And blnCanModify
                Else
                    blnStart = (.SelectedItem.Icon = "Item") And blnCanModify
                    blnPrice = True
                End If
                blnItemModify = blnStart
            End If
        End With
    End If
    
    '�༭
    mnuEditParent.Enabled = True '��������
    mnuEditModifyAssort.Enabled = blnClassModify '�޸ķ���
    mnuEditDeleteAssort.Enabled = blnClassModify 'ɾ������
    
    mnuEditChild.Enabled = True '������Ŀ
    mnuEditCopy.Enabled = Not lvwMain_S.SelectedItem Is Nothing '��������
    mnuEditModify.Enabled = blnItemModify '�޸���Ŀ
    mnuEditDelete.Enabled = blnItemModify 'ɾ����Ŀ
    
    Toolbar1.Buttons("Modify").Enabled = blnClassModify Or blnItemModify
    Toolbar1.Buttons("Delete").Enabled = blnClassModify Or blnItemModify
    
    mnuEditDept.Enabled = blnItemModify   'ִ�п���
    mnuEditSlave.Enabled = blnItemModify '������Ŀ
    mnuEditItemGroup.Enabled = blnItemModify '��Ŀ���
    
    mnuEditStart.Enabled = blnStop   '����
    mnuEditStop.Enabled = blnStart  'ͣ��
    Toolbar1.Buttons("Start").Enabled = blnStop
    Toolbar1.Buttons("Stop").Enabled = blnStart
    
    '��Ŀ����
    mnuPriceRaise.Enabled = blnPrice '����
    Toolbar1.Buttons("Raise").Enabled = blnPrice
    If gstrҽ�۽ӿڱ�� <> "" And gbln����ҽ���շ���Ŀ Then
        mnuPriceRaiseMass.Enabled = False '��������
    Else
        mnuPriceRaiseMass.Enabled = blnPrice
    End If
    mnuPriceHistory.Enabled = blnPrice 'ɾ��δִ�м۸�
    
    mnuPriceChargeSet.Enabled = (InStr(mstrPrivs, "�ѱ�����") > 0) And Not lvwMain_S.SelectedItem Is Nothing '�ѱ�����
    mnuEditItemGroup.Enabled = (InStr(mstrPrivs, "��Ŀ����") > 0)
    
    '��ӡ
    mnuFilepre.Enabled = blnPrint
    mnuFilePrint.Enabled = blnPrint
    mnuFileExcel.Enabled = blnPrint
    Toolbar1.Buttons("Preview").Enabled = blnPrint
    Toolbar1.Buttons("Print").Enabled = blnPrint
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Ȩ�޿���()
    Dim rsTmp As New ADODB.Recordset
    
    '��ʼ����������ť
    On Error GoTo ErrHandle
    Toolbar1.Buttons("Split1").Visible = True
    Toolbar1.Buttons("Start").Visible = True
    Toolbar1.Buttons("Stop").Visible = True
    Toolbar1.Buttons("Raise").Visible = True
    Toolbar1.Buttons("Split2").Visible = True
    Toolbar1.Buttons("Find").Visible = True
    Toolbar1.Buttons("Split3").Visible = True
    Toolbar1.Buttons("Split4").Visible = True
    
    '����:�����е��û�Ȩ�޲���,��ʹһЩ�˵����ť���ɼ�
    If InStr(mstrPrivs, "��Ŀ����") = 0 Then
        mnuEditChild.Visible = False  '������Ŀ
        mnuEditCopy.Visible = False   '���ƿ���
        mnuEditModify.Visible = False '�޸�
        mnuEditDelete.Visible = False 'ɾ��
        mnuEditDept.Visible = False   'ִ�п���
        mnuEditStart.Visible = False  '����
        mnuEditStop.Visible = False   'ͣ��
        mnuEditSplit0.Visible = False '��һ���ָ�
        mnuEditSplit1.Visible = False '�ڶ����ָ�
        
        mnuShortMenu2(0).Visible = False  '��Ŀ�Ŀ�ݲ˵��༭���ܲ��ɼ�
        mnuShortMenu2(1).Visible = False
        mnuShortMenu2(2).Visible = False
        mnuShortMenu2(3).Visible = False
        mnuShortMenu2(5).Visible = False
        
        mnuShortsplit1.Visible = False
        Toolbar1.Buttons("Child").Visible = False
        Toolbar1.Buttons("Modify").Visible = False
        Toolbar1.Buttons("Delete").Visible = False
        Toolbar1.Buttons("Split1").Visible = False
        Toolbar1.Buttons("Start").Visible = False
        Toolbar1.Buttons("Stop").Visible = False
        Toolbar1.Buttons("Split3").Visible = False
    End If
    
    If InStr(mstrPrivs, "������") = 0 Then
        mnuClassEdit.Visible = False
        mnuEditSplit3.Visible = False
        Toolbar1.Buttons("Parent").Enabled = False
        Me.mnuEditParent.Enabled = False
        Me.mnuEditModifyAssort.Enabled = False
        Me.mnuEditDeleteAssort.Enabled = False
        Me.mnuShort1.Visible = False
        mnuEditParent.Visible = False '���ӷ���
        Toolbar1.Buttons("Parent").Visible = False
    End If
    
    If InStr(mstrPrivs, "��Ŀ�������") = 0 Then
        mnuEditSlave.Visible = False  '������Ŀ
        mnuEditItemGroup.Visible = False    '��Ŀ���
        mnuShortMenu2(4).Visible = False
        mnuShortMenu2(6).Visible = False
    End If
    
    If InStr(mstrPrivs, "��Ŀ����") = 0 Then
        mnuPrice.Visible = False
        mnuPriceRaise.Visible = False
        mnuPriceRaiseMass.Visible = False
        mnuPriceHistory.Visible = False
        Toolbar1.Buttons("Raise").Visible = False
        Toolbar1.Buttons("Split2").Visible = False
    End If
    
    If InStr(mstrPrivs, "�ѱ�����") = 0 Then
        mnuPriceChargeSet.Visible = False
    End If
    
    If InStr(mstrPrivs, "ҽ�۽ӿ�") = 0 Then
        Me.mnuFileStdImp.Visible = False
        Me.mnuFileStdCheck.Visible = False
        Me.mnuFileSplit1.Visible = False
    Else
        Me.mnuFileStdImp.Visible = True
        Me.mnuFileStdCheck.Visible = True
        Me.mnuFileSplit1.Visible = True
        
        '�õ�ҽ�۽ӿڱ���
        gstrSQL = "select ���,ҽ�� from ҽ�۽ӿ� where nvl(ѡ��,0)=1"
        Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, Me.Caption)
        If rsTmp.RecordCount > 0 Then
            gstrҽ�۽ӿڱ�� = Nvl(rsTmp!���)
            gbln����ҽ���շ���Ŀ = CStr(Nvl(rsTmp!ҽ��)) = "1"
            mnuFileStdImp.Enabled = True
            mnuFileStdCheck.Enabled = True
            mbln����ҽ��ϵͳ = True
        Else
            gstrҽ�۽ӿڱ�� = ""
            gbln����ҽ���շ���Ŀ = False
            mnuFileStdImp.Enabled = False
            mnuFileStdCheck.Enabled = False
        End If
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub GetDefineSize()
'���ܣ��õ����ݿ�ı��ֶεĳ���
On Error GoTo ErrHandle
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "Select ���� From �շ���ĿĿ¼ Where Rownum<0"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�շ���ĿĿ¼")
    
    mlng���볤�� = rsTemp.Fields("����").DefinedSize
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub


Private Sub picPage_Resize(Index As Integer)
        Err = 0: On Error Resume Next
        With picPage(Index)
                Select Case Index
                Case 1
                    msh��Ŀ.Left = 50
                    msh��Ŀ.Height = .ScaleHeight - 100 - msh��Ŀ.Top
                    msh��Ŀ.Width = .ScaleWidth - msh��Ŀ.Left - 100
                Case 2
                    Frame2.Top = 0
                    Frame2.Left = 0
                    Frame2.Width = .ScaleWidth
                    Frame2.Height = .ScaleHeight
                    
                    Frame1.Width = .ScaleWidth - 300
                    Frame1.Height = .ScaleHeight - Frame1.Top - 200
                    lvwOutIn.Top = 0
                    lvwOutIn.Left = 0
                    lvwOutIn.Width = Frame1.Width
                    lvwOutIn.Height = Frame1.Height - lvwOutIn.Top
                Case 3
                    msh����.Top = 800
                    msh����.Height = .ScaleHeight - 100 - msh����.Top
                    msh����.Width = .ScaleWidth - 300
                Case 4
                    mshAlias.Left = 0
                    mshAlias.Width = .ScaleWidth
                    mshAlias.Top = 0
                    mshAlias.Height = .ScaleHeight
                Case 5
                    msh�ѱ�.Left = 0
                    msh�ѱ�.Width = .ScaleWidth
                    msh�ѱ�.Top = 0
                    msh�ѱ�.Height = .ScaleHeight
                Case 6
                    vsWholeSet.Left = 0
                    vsWholeSet.Top = 0
                    vsWholeSet.Width = .ScaleWidth
                    vsWholeSet.Height = .ScaleHeight
                Case 7
                    lvwUseDept.Left = 0
                    lvwUseDept.Top = 0
                    lvwUseDept.Width = .ScaleWidth
                    lvwUseDept.Height = .ScaleHeight
                End Select
        End With
End Sub

Private Sub tvwWholeSet_GotFocus()
    Call MenuSet
End Sub

Private Sub tvwWholeSet_LostFocus()
    Call MenuSet
End Sub

Private Sub tvwWholeSet_NodeClick(ByVal Node As MSComctlLib.Node)
        '���س�����Ŀ����
        If tvwWholeSet.Tag <> Node.Key Then
            tvwWholeSet.Tag = Node.Key
            Call FillWholeItem(Val(Mid(Node.Key, 2)))
        End If
        Call MenuSet
End Sub

Private Sub txtFind_GotFocus()
    Call zlControl.TxtSelAll(txtFind)
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strSQL As String
    Dim rsTmp As Recordset
    Dim vRect As RECT, blnCancel As Boolean
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    On Error GoTo ErrHandle
    strSQL = "Select Distinct a.Id, a.���, b.����, a.����, a.��ʶ����, a.��ʶ����, b.����, c.���� As ����, a.����id, a.����ʱ��" & vbNewLine & _
            "From (Select ID, ���, ����id, ����, ����, ��ʶ����, ��ʶ����, ����ʱ��" & vbNewLine & _
            "       From �շ���ĿĿ¼" & vbNewLine & _
            "       Where ��� <> '5' And ��� <> '6' And ��� <> '7'" & _
            IIF(mnuViewShowStop.Checked, "", " And (����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null)") & vbNewLine & _
            ") A, (Select a.�շ�ϸĿid, a.����, a.���� || '/' || b.���� As ����" & vbNewLine & _
            "       From �շ���Ŀ���� A, �շ���Ŀ���� B" & vbNewLine & _
            "       Where a.�շ�ϸĿid = b.�շ�ϸĿid And a.���� = 1 And b.���� = 2) B, �շѷ���Ŀ¼ C" & vbNewLine & _
            "Where a.����id = c.Id(+) And a.Id = b.�շ�ϸĿid And c.���� Is Not Null"
    If txtFind.Text = "" Then Exit Sub
    If zlStr.IsCharChinese(txtFind.Text) Then
        strSQL = strSQL & " And b.���� Like [1]"
    ElseIf IsNumeric(txtFind.Text) Then
        strSQL = strSQL & " And a.���� Like [2]"
    Else
        strSQL = strSQL & " And (b.���� Like [1] Or b.���� Like [3])"
    End If
    
    vRect = zlControl.GetControlRect(txtFind.hwnd)
    If vRect.Left + 7350 > Screen.Width Then vRect.Left = Screen.Width - 7350
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "�շ�ϸĿѡ��", False, "", "", False, False, True, _
                        vRect.Left, vRect.Top, txtFind.Height, blnCancel, False, True, gstrLike & txtFind.Text & "%", txtFind.Text & "%", gstrLike & UCase(txtFind.Text) & "%")
    If blnCancel = True Then Exit Sub
    If Not rsTmp Is Nothing Then
        Call FindLocate(rsTmp)
    Else
        MsgBox "û���ҵ��������ҵ��շ���Ŀ��", vbInformation, Me.Caption
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FindLocate(ByVal rsTmp As Recordset)
'���Ҷ�λ
    Dim strKey As String
    
    On Error Resume Next
        strKey = "R" & rsTmp!����id
        If rsTmp!����id & "" <> "" Then
            Me.tvwMainItem.Nodes(strKey).Selected = True
            Me.tvwMainItem.Nodes(strKey).EnsureVisible
            Me.tvwMainItem_NodeClick Me.tvwMainItem.SelectedItem
            Err.Clear
            Me.lvwMain_S.ListItems("C" & rsTmp!��� & rsTmp!ID).Selected = True
            Me.lvwMain_S.ListItems("C" & rsTmp!��� & rsTmp!ID).EnsureVisible
            If Err.Number = 35601 Then
                MsgBox "���ҵ���������¼�����ѱ�ɾ����ͣ�ã���ˢ���б�", vbInformation, gstrSysName
                Err.Clear
                Exit Sub
            End If
            Me.lvwMain_S_ItemClick Me.lvwMain_S.SelectedItem
        Else
            Me.tvwMainItem.Nodes("Root").Selected = True
            Me.tvwMainItem.Nodes(strKey).EnsureVisible
            Me.tvwMainItem_NodeClick Me.tvwMainItem.SelectedItem
            Err.Clear
            Me.lvwMain_S.ListItems("C" & rsTmp!��� & rsTmp!ID).Selected = True
            Me.lvwMain_S.ListItems("C" & rsTmp!��� & rsTmp!ID).EnsureVisible
            If Err.Number = 35601 Then
                MsgBox "���ҵ���������¼�����ѱ�ɾ����ͣ�ã���ˢ���б�", vbInformation, gstrSysName
                Err.Clear
                Exit Sub
            End If
            Me.lvwMain_S_ItemClick Me.lvwMain_S.SelectedItem
        End If
    Err.Clear
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "'" Then KeyAscii = 0: Exit Sub
End Sub

Private Sub vsWholeSet_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngMode, vsWholeSet, Me.Caption, "������Ŀ��ɱ���-������", True, True
End Sub

Private Sub vsWholeSet_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngMode, vsWholeSet, Me.Caption, "������Ŀ��ɱ���-������", True, True
End Sub

Private Sub vsWholeSet_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsWholeSet
        Select Case Col
        Case .ColIndex("��־")
            Cancel = True
        Case Else
        End Select
    End With
End Sub
