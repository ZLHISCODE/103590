VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmStuffMgr 
   BackColor       =   &H8000000A&
   Caption         =   "����Ŀ¼����"
   ClientHeight    =   8280
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   11415
   Icon            =   "frmStuffMgr.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8280
   ScaleWidth      =   11415
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picClass 
      Height          =   4455
      Left            =   0
      ScaleHeight     =   4395
      ScaleWidth      =   2355
      TabIndex        =   16
      Top             =   840
      Width           =   2415
      Begin VB.CommandButton cmdKind 
         Caption         =   "���˽��(&1)"
         Height          =   300
         Index           =   1
         Left            =   0
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   300
         UseMaskColor    =   -1  'True
         Width           =   2295
      End
      Begin VB.CommandButton cmdKind 
         Caption         =   "��    ��(&0)"
         Height          =   300
         Index           =   0
         Left            =   0
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   15
         Width           =   2295
      End
      Begin MSComctlLib.TreeView tvwClass 
         Height          =   3030
         Left            =   0
         TabIndex        =   19
         Top             =   840
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   5345
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   494
         LabelEdit       =   1
         LineStyle       =   1
         Sorted          =   -1  'True
         Style           =   7
         ImageList       =   "imgList"
         BorderStyle     =   1
         Appearance      =   1
      End
   End
   Begin VB.PictureBox picCost 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2640
      Left            =   -810
      ScaleHeight     =   2640
      ScaleWidth      =   5730
      TabIndex        =   14
      Top             =   5310
      Width           =   5730
      Begin VSFlex8Ctl.VSFlexGrid vsCost 
         Height          =   2070
         Left            =   225
         TabIndex        =   15
         Top             =   30
         Width           =   7080
         _cx             =   12488
         _cy             =   3651
         Appearance      =   3
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
         ForeColorSel    =   0
         BackColorBkg    =   -2147483634
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
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
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmStuffMgr.frx":030A
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
         ExplorerBar     =   3
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
         BackColorFrozen =   16777215
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.PictureBox picVBar_S 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6660
      Left            =   2520
      MousePointer    =   9  'Size W E
      ScaleHeight     =   6660
      ScaleWidth      =   45
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   840
      Width           =   45
   End
   Begin VB.PictureBox picHBar_S 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   3360
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   6075
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5520
      Width           =   6075
   End
   Begin VB.PictureBox picTabPrice 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2550
      Left            =   2640
      ScaleHeight     =   2550
      ScaleWidth      =   3735
      TabIndex        =   9
      Top             =   4590
      Width           =   3735
      Begin VB.Frame fraComment 
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   60
         TabIndex        =   10
         Top             =   1905
         Width           =   7410
         Begin VB.Label lblComment 
            AutoSize        =   -1  'True
            Caption         =   "1��ʱ�۲��ϣ�ָ��������6Ԫ/֧������"
            Height          =   180
            Index           =   3
            Left            =   0
            TabIndex        =   12
            Top             =   30
            Width           =   3150
         End
         Begin VB.Label lblComment 
            AutoSize        =   -1  'True
            Caption         =   "2������ۼ�198.25Ԫ/֧�����ݲ�����ݷѱ�����Żݻ�Ӽۡ�"
            Height          =   180
            Index           =   4
            Left            =   0
            TabIndex        =   11
            Top             =   270
            Width           =   5040
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsPrice 
         Height          =   1725
         Left            =   90
         TabIndex        =   13
         Top             =   105
         Width           =   7365
         _cx             =   12991
         _cy             =   3043
         Appearance      =   3
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
         ForeColorSel    =   0
         BackColorBkg    =   -2147483634
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
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
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   15
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmStuffMgr.frx":0416
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
         ExplorerBar     =   3
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
         BackColorFrozen =   16777215
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.PictureBox picStuffLst 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2640
      Left            =   6885
      ScaleHeight     =   2640
      ScaleWidth      =   2670
      TabIndex        =   7
      Top             =   4680
      Width           =   2670
      Begin VSFlex8Ctl.VSFlexGrid vsStuff 
         Height          =   2070
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   7080
         _cx             =   12488
         _cy             =   3651
         Appearance      =   3
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
         ForeColorSel    =   0
         BackColorBkg    =   -2147483634
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
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
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   37
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmStuffMgr.frx":061F
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
         ExplorerBar     =   3
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
         BackColorFrozen =   16777215
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin XtremeSuiteControls.TabControl TbList 
      Height          =   3630
      Left            =   2445
      TabIndex        =   6
      Top             =   4170
      Width           =   8475
      _Version        =   589884
      _ExtentX        =   14949
      _ExtentY        =   6403
      _StockProps     =   64
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   10080
      Top             =   1200
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
            Picture         =   "frmStuffMgr.frx":0B4B
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":10E5
            Key             =   "expend"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":167F
            Key             =   "start"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":1B91
            Key             =   "stop"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwItems 
      Height          =   2895
      Left            =   2520
      TabIndex        =   1
      Top             =   960
      Width           =   7200
      _ExtentX        =   12700
      _ExtentY        =   5106
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgList"
      SmallIcons      =   "imgList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   7905
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmStuffMgr.frx":1DAB
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15055
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
   Begin ComCtl3.CoolBar clbThis 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   11415
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tlbThis"
      MinWidth1       =   24000
      MinHeight1      =   720
      Width1          =   8730
      FixedBackground1=   0   'False
      Key1            =   "Comm"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tlbThis 
         Height          =   720
         Left            =   30
         TabIndex        =   20
         Top             =   30
         Width           =   24000
         _ExtentX        =   42333
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
            NumButtons      =   20
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ԥ��"
               Key             =   "Preview"
               Description     =   "Ԥ��"
               Object.ToolTipText     =   "Ԥ����ǰ��"
               Object.Tag             =   "Ԥ��"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "Print"
               Description     =   "��ӡ"
               Object.ToolTipText     =   "��ӡ��ǰ��"
               Object.Tag             =   "��ӡ"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Class"
               Description     =   "����"
               Object.ToolTipText     =   "�������Ϸ���"
               Object.Tag             =   "����"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split4"
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Description     =   "����"
               Object.ToolTipText     =   "���ӹ���Ʒ��"
               Object.Tag             =   "����"
               ImageKey        =   "Add"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "����Ʒ��"
                     Object.Tag             =   "����Ʒ��"
                     Text            =   "����Ʒ��"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "���ӹ��"
                     Object.Tag             =   "���ӹ��"
                     Text            =   "���ӹ��"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�޸�"
               Key             =   "�޸�"
               Description     =   "�޸�"
               Object.ToolTipText     =   "�޸�"
               Object.Tag             =   "�޸�"
               ImageKey        =   "Modify"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ɾ��"
               Key             =   "ɾ��"
               Object.ToolTipText     =   "ɾ��"
               Object.Tag             =   "ɾ��"
               ImageKey        =   "Delete"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split1"
               Description     =   "ɾ��"
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Start"
               Description     =   "����"
               Object.ToolTipText     =   "����ָ����ͣ�ò���"
               Object.Tag             =   "����"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ͣ��"
               Key             =   "Stop"
               Description     =   "ͣ��"
               Object.ToolTipText     =   "ͣ��ָ�������ò���"
               Object.Tag             =   "ͣ��"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split3"
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split12"
               Style           =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�鿴"
               Key             =   "�鿴"
               Object.ToolTipText     =   "�鿴"
               Object.Tag             =   "�鿴"
               ImageKey        =   "View"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "��ͼ��"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Сͼ��"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "�б�"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "��ϸ����"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Find"
               Description     =   "����"
               Object.ToolTipText     =   "��������Ŀ¼"
               Object.Tag             =   "����"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Filter"
               Description     =   "����"
               Object.ToolTipText     =   "��������Ŀ¼"
               Object.Tag             =   "����"
               ImageIndex      =   16
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Description     =   "����"
               Object.ToolTipText     =   "��ǰ��������"
               Object.Tag             =   "����"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Exit"
               Description     =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageIndex      =   11
            EndProperty
         EndProperty
         Begin VB.PictureBox picFind 
            BorderStyle     =   0  'None
            Height          =   300
            Left            =   9000
            ScaleHeight     =   285.714
            ScaleMode       =   0  'User
            ScaleWidth      =   495
            TabIndex        =   22
            Top             =   240
            Width           =   495
            Begin VB.Label lbl���� 
               Caption         =   "����"
               Height          =   255
               Left            =   120
               TabIndex        =   23
               Top             =   75
               Width           =   495
            End
         End
         Begin VB.TextBox txtFind 
            Height          =   300
            Left            =   9600
            MaxLength       =   10
            TabIndex        =   21
            Tag             =   "����"
            Top             =   240
            Width           =   1425
         End
      End
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   7680
      Top             =   525
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":263D
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":2857
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":2A71
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":2C8B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":2EA5
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":359F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":37B9
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":39D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":40CD
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":42E7
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":4507
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":4727
            Key             =   "Add"
            Object.Tag             =   "Add"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":4941
            Key             =   "Modify"
            Object.Tag             =   "Modify"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":4B5B
            Key             =   "Delete"
            Object.Tag             =   "Delete"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":4D75
            Key             =   "View"
            Object.Tag             =   "View"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":4F8F
            Key             =   "Filter"
            Object.Tag             =   "Filter"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   6915
      Top             =   525
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":52A9
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":54C9
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":56E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":5903
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":5B1D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":6217
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":6431
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":664B
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":6D45
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":6F5F
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":717F
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":739F
            Key             =   "Add"
            Object.Tag             =   "Add"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":75B9
            Key             =   "Modify"
            Object.Tag             =   "Modify"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":77D3
            Key             =   "Delete"
            Object.Tag             =   "Delete"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":79ED
            Key             =   "View"
            Object.Tag             =   "View"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":7C07
            Key             =   "Filter"
            Object.Tag             =   "Filter"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picUp_S 
      Appearance      =   0  'Flat
      BackColor       =   &H80000011&
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2790
      ScaleHeight     =   255
      ScaleWidth      =   7170
      TabIndex        =   5
      Top             =   3900
      Width           =   7170
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFilePrintSet 
         Caption         =   "��ӡ����(&S)"
      End
      Begin VB.Menu mnuFilePreview 
         Caption         =   "��ӡԤ��(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "��ӡ(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "�����&Excel"
      End
      Begin VB.Menu mnuFileMlPrint 
         Caption         =   "Ŀ¼��ӡ(&M)"
      End
      Begin VB.Menu mnuFileMlPreview 
         Caption         =   "Ŀ¼Ԥ��(&Y)"
      End
      Begin VB.Menu mnuFileSpt1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePara 
         Caption         =   "��������(&A)"
         Shortcut        =   {F12}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileSpt2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuClass 
      Caption         =   "����(&K)"
      Begin VB.Menu mnuClassAdd 
         Caption         =   "����(&I)"
         Shortcut        =   +{INSERT}
      End
      Begin VB.Menu mnuClassMod 
         Caption         =   "�޸�(&U)"
      End
      Begin VB.Menu mnuClassDel 
         Caption         =   "ɾ��(&E)"
         Shortcut        =   +{DEL}
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "����(&E)"
      Begin VB.Menu mnuEditAddName 
         Caption         =   "����Ʒ��(&P)"
      End
      Begin VB.Menu mnuEditModifyName 
         Caption         =   "�޸�Ʒ��(&E)"
      End
      Begin VB.Menu mnuEditDeleName 
         Caption         =   "ɾ��Ʒ��(&D)"
      End
      Begin VB.Menu mnuEditSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditItemAdd 
         Caption         =   "�������(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditItemMod 
         Caption         =   "�޸Ĺ��(&M)"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuEditItemDel 
         Caption         =   "ɾ�����(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditSpt1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditItemPart 
         Caption         =   "�洢�ⷿ(&F)..."
      End
      Begin VB.Menu mnuEditSpecLimit 
         Caption         =   "��������(&L)..."
      End
      Begin VB.Menu mnuEditSpecSelf 
         Caption         =   "���Ʋ���(&H)..."
      End
      Begin VB.Menu mnuEditUnit 
         Caption         =   "�б굥λ(&Z)"
      End
      Begin VB.Menu mnuEditSpt2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSubsection 
         Caption         =   "�ֶμӳ���(&J)"
      End
      Begin VB.Menu mnuEditExcel 
         Caption         =   "������Ŀ"
      End
      Begin VB.Menu mnuEditSpt3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditStart 
         Caption         =   "����(&R)"
      End
      Begin VB.Menu mnuEditStop 
         Caption         =   "ͣ��(&S)"
      End
      Begin VB.Menu mnuEditSpt4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditPrice 
         Caption         =   "����(&T)"
      End
   End
   Begin VB.Menu mnuPrice 
      Caption         =   "�۸�(&P)"
      Begin VB.Menu mnuPriceTable 
         Caption         =   "���ۼ�¼��(&S)"
      End
      Begin VB.Menu mnuPriceLists 
         Caption         =   "���ϼ�Ŀ��(&L)..."
         Shortcut        =   ^L
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
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "������(&T)"
         Begin VB.Menu mnuViewToolbarStand 
            Caption         =   "��׼��ť(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewToolbarText 
            Caption         =   "�ı���ǩ(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStates 
         Caption         =   "״̬��(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewSpt1 
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
         Index           =   3
      End
      Begin VB.Menu mnuViewSp 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewStoped 
         Caption         =   "��ʾͣ�ò���(&C)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewPrices 
         Caption         =   "��ʾ��ʷ�۸�(&H)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewDownLevel 
         Caption         =   "��ʾ�����¼�(&X)"
      End
      Begin VB.Menu mnuViewSpt2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "����(&F)..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "����(&T)"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuViewSpt3 
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
         Caption         =   "Web�ϵ�����(&W)"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "������ҳ(&H)"
         End
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "������̳(&F)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "���ͷ���(&E)..."
         End
      End
      Begin VB.Menu mnuHelp1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)..."
      End
   End
End
Attribute VB_Name = "frmStuffMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mstrPrivs As String       '�û����б�����ľ���Ȩ��

Dim rsTemp As New ADODB.Recordset
Dim objNode As Node
Dim objItem As ListItem
Dim intCount As Integer, intRow As Integer, intCol As Integer
Dim strTemp As String
Dim mintUnit  As Integer
Private mlngModule As Long
Dim mstrMaterialID As String
Private Declare Function SetParent Lib "user32 " (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private mstrFindValue As String
Private mrsFind As ADODB.Recordset
Private mrsRefClasses As ADODB.Recordset
Private mrsRefRecords As ADODB.Recordset

Private Enum colPrice
    NO = 0
    ���
    ��������
    ָ������
    ����
    ָ���ۼ�
    ָ������
    ִ�����
    ���ηѱ�
    ����
    ��λ
    �ۼ�
    ������Ŀ
    ˵��
    ִ������
End Enum
Private mbln��ʾ�¼� As Boolean         '�Ƿ���ʾ�¼����в�
'----------------------------------------------------------------------------------------------------------
'���˺�:����С��λ���ĸ�ʽ��
'�޸�:2007/03/06
Private mFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------


Private Sub SetMenuEnable()
    Dim blnStop As Boolean  'Ʒ��ͣ��
    Dim blnListStop As Boolean  '���ͣ��
    Dim blnSel As Boolean   'ѡ��Ʒ�ֵ�
    Dim blnListSel As Boolean 'ѡ�еĲ���
    
    Dim blnData As Boolean '�������ݷ�
    Dim blnSel���� As Boolean 'ѡ�з���
        
    blnSel���� = Not tvwClass.SelectedItem Is Nothing
    
    blnData = Me.lvwItems.ListItems.count <> 0
    blnSel = Not Me.lvwItems.SelectedItem Is Nothing
    If blnSel Then
        blnStop = Me.lvwItems.SelectedItem.Icon = "stop"
    End If
    
    With vsStuff
        blnListSel = .RowData(.Row) <> 0
        If blnSel Then
            blnListStop = .TextMatrix(.Row, .ColIndex("����ʱ��")) <> ""
        End If
    End With
    
    
    '��Ʒ�ֵ���ز�����������
    mnuEditModifyName.Enabled = blnSel And Not blnStop
    mnuEditDeleName.Enabled = blnSel And Not blnStop
    
    'ȷ���������ϵ���ɾ��
    mnuEditItemAdd.Enabled = blnSel And Not blnStop
    tlbThis.Buttons("����").ButtonMenus("���ӹ��").Enabled = mnuEditItemAdd.Enabled
    
    mnuEditItemMod.Enabled = blnListSel And Not blnListStop
    mnuEditItemDel.Enabled = blnListSel And Not blnListStop
    mnuEditPrice.Enabled = blnListSel And Not blnListStop
    
    If Me.ActiveControl Is lvwItems Then
        tlbThis.Buttons("�޸�").Enabled = blnSel And Not blnStop
        tlbThis.Buttons("ɾ��").Enabled = blnSel And Not blnStop
        mnuEditStop.Enabled = Not blnStop And blnSel
        mnuEditStart.Enabled = blnStop And blnSel
    Else
        tlbThis.Buttons("�޸�").Enabled = blnListSel And Not blnListStop
        tlbThis.Buttons("ɾ��").Enabled = blnListSel And Not blnListStop
        mnuEditStop.Enabled = Not blnListStop And blnListSel
        mnuEditStart.Enabled = blnListStop And blnListSel
    End If
    
    tlbThis.Buttons("Start").Enabled = mnuEditStart.Enabled
    tlbThis.Buttons("Stop").Enabled = mnuEditStop.Enabled
    
    'ȷ����������
    mnuClassDel.Enabled = blnSel����
    mnuClassMod.Enabled = blnSel����
    
    '��Ҫ���Ȩ��
    Dim blnVisible As Boolean
    
    If Me.ActiveControl Is Me.vsStuff Then
        blnVisible = True
    ElseIf Me.ActiveControl Is Me.vsPrice Then
        blnVisible = zlStr.IsHavePrivs(mstrPrivs, "�ۼ۹���")
    ElseIf Me.ActiveControl Is Me.vsCost Then
        blnVisible = zlStr.IsHavePrivs(mstrPrivs, "�ɱ��۹���")
    Else
        blnVisible = True
    End If
    
    'ȷ����ӡ��Ԥ��
    mnuFileExcel.Enabled = blnData
    mnuFilePreview.Enabled = blnData
    mnuFilePrint.Enabled = blnData
    tlbThis.Buttons("Preview").Enabled = blnData
    tlbThis.Buttons("Print").Enabled = blnData
End Sub
Private Sub clbThis_Resize()
    Me.clbThis.Bands(1).MinHeight = Me.tlbThis.Height
    Me.clbThis.Refresh
    Call Form_Resize
End Sub

Private Sub cmdKind_Click(Index As Integer)
    Dim intCount As Integer
    Call SaveListViewState(Me.lvwItems, Me.Name & Val(Me.tvwClass.Tag), Me.lvwItems.View)
    For intCount = Me.cmdKind.LBound To Me.cmdKind.UBound
        If intCount <= Index Then
            Me.cmdKind(intCount).Tag = 0
        Else
            Me.cmdKind(intCount).Tag = 1
        End If
    Next
    
    '�����ݣ����ܴ�ӡ��Ԥ�� 2010-5-10
    mnuFileMlPrint.Enabled = Not tvwClass.SelectedItem Is Nothing
    mnuFileMlPreview.Enabled = Not tvwClass.SelectedItem Is Nothing
    
    'װ���ݲ���������
    If Me.lvwItems.Visible Then
        Call picClass_Resize
        Me.tvwClass.SetFocus
    End If
    If Index = 0 Then
        If Val(tvwClass.Tag) <> Index Then
            Me.tvwClass.Tag = Index
            Call zlRefClasses
        End If
        Me.mnuViewFind.Enabled = True
        Me.tlbThis.Buttons("Find").Enabled = True
    Else
        Me.tvwClass.Tag = Index
        Call zlRefClasses
        Me.mnuViewFind.Enabled = False
        Me.tlbThis.Buttons("Find").Enabled = False
        frmStuffFind.Hide
    End If
End Sub

Private Sub Form_Activate()
    Me.lvwItems.Visible = True
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    Dim rs������Ŀ As ADODB.Recordset
    Dim bln������Ŀ As Boolean
    
    '����ָ�
    Call RestoreWinState(Me, App.ProductName)
    mlngModule = glngModul
    mbln��ʾ�¼� = Val(zlDatabase.GetPara("�����¼�����", glngSys, mlngModule, "0")) = 1
    
    Me.mnuViewDownLevel.Checked = mbln��ʾ�¼�
    Me.mnuViewStoped.Checked = False
    Me.mnuViewPrices.Checked = False
    
    mnuViewIcon(0).Checked = False
    mnuViewIcon(1).Checked = False
    mnuViewIcon(2).Checked = False
    mnuViewIcon(3).Checked = False
    mnuViewIcon(lvwItems.View).Checked = True
    
    '����Ƿ��Կⷿ��λ��ʾ�۸�
    mintUnit = Get���۵�λ
  
    '���˺�:����С����ʽ����
    With mFMT
        .FM_�ɱ��� = GetFmtString(mintUnit, g_�ɱ���)
        .FM_��� = GetFmtString(mintUnit, g_���)
        .FM_���ۼ� = GetFmtString(mintUnit, g_�ۼ�)
        .FM_���� = GetFmtString(mintUnit, g_����)
    End With
    
    
    'mstrFormat = IIf(mintUnit, "####0.0000;-####0.0000; ;", "####0.0000000;-####0.0000000; ;")
    '��ֱ��ͨ���˵����е�Ȩ�޿���
    mstrPrivs = ";" & gstrPrivs & ";"

    '����Ȩ��
    Call SetPopedom
    '��ʼ�۸���ͷ
     
    Call InitHgdPrivceHeadcol
       
'    Me.picHBar_S.Top = Me.ScaleHeight - IIf(Me.stbThis.Visible, Me.stbThis.Height, 0) - 2500
    '��ʼ��ͷ
    Call InitLvwStuffColHead
    
    '��ȡ�������
    zlRefClasses
    '2006-04-25:���˺�,ͳһ���ӱ�������ģ��Ĺ���
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, gstrPrivs)
     
    Call SetMenuEnable
    Call InitTabCtl
    Call vsCost_LostFocus
    Call vsPrice_LostFocus
    Call vsStuff_LostFocus
    
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����", 0, 0)) = 1 Then
        strTemp = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\�ָ�", "����", "0")
        If strTemp <> "0" Then
            Me.picVBar_S.Left = CLng(strTemp)
        End If
        strTemp = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\�ָ�", "����", "0")
        If strTemp <> "0" Then
            Me.picHBar_S.Top = CLng(strTemp)
        End If
    End If
    
    zl_vsGrid_Para_Restore mlngModule, vsCost, Me.Caption, "�ɱ��۵��ۼ�¼"
    zl_vsGrid_Para_Restore mlngModule, vsPrice, Me.Caption, "�ۼ۵��ۼ�¼"
    zl_vsGrid_Para_Restore mlngModule, vsStuff, Me.Caption, "���Ϲ���¼"
    
    If vsStuff.ColWidth(vsStuff.ColIndex("����")) = 0 Then vsStuff.ColWidth(vsStuff.ColIndex("����")) = 1600
    vsStuff.ColHidden(vsStuff.ColIndex("����")) = False
    
    Call cmdKind_Click(0)
    
    gstrSQL = "select ����ֵ from zlParameters where ģ��=1711 and ������='������Ŀ��Ӧ'"
    Set rs������Ŀ = zlDatabase.OpenSQLRecord(gstrSQL, "������Ŀ", "������Ŀ��Ӧ")
    If rs������Ŀ.RecordCount > 0 Then
        If IsNull(rs������Ŀ!����ֵ) Then
            bln������Ŀ = True
        End If
    End If
    
    If bln������Ŀ = True Then
        MsgBox "�����ø����ʶ�Ӧ��������Ŀ��", vbInformation, gstrSysName
        frmStuffPara.ShowMe mstrPrivs, Me
        If gblnIncomeItem = False Then
            Unload Me
        End If
    End If
End Sub
Private Sub InitTabCtl()
    '-----------------------------------------------------------------------------------------------------------
    '����:��ʼ��Tab�ؼ�
    '����:���˺�
    '����:2007/05/24
    '-----------------------------------------------------------------------------------------------------------
    TbList.SetImageList Me.imgList
    TbList.InsertItem 0, "���Ϲ��(&1)", picStuffLst.hwnd, 0
    TbList.InsertItem 1, "�ۼ۵��ۼ�¼(&2)", picTabPrice.hwnd, 0
    TbList.InsertItem 2, "�ɱ��۵��ۼ�¼(&3)", picCost.hwnd, 0
    
    TbList.Item(2).Visible = InStr(1, mstrPrivs, ";�ɱ��۹���;") > 0
    TbList.PaintManager.Appearance = xtpTabAppearancePropertyPage2003
End Sub

Private Sub SetPopedom()
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:Ȩ�޿���
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    Dim blnModify As Boolean    '�����޸�
    Dim bln���� As Boolean
    
    blnModify = InStr(1, mstrPrivs, ";�ۼ۹���;") <> 0      '���ۼ�,���Ը���
    blnModify = blnModify Or InStr(1, mstrPrivs, ";�ɱ��۹���;") <> 0      '�гɱ���,Ҳ���Ը���
    blnModify = blnModify Or InStr(1, mstrPrivs, ";�������;") <> 0      '�п���,Ҳ���Ը���
    blnModify = blnModify Or InStr(1, mstrPrivs, ";ָ���۸����;") <> 0      '��ָ���۸�,Ҳ���Ը���
    blnModify = blnModify Or InStr(1, mstrPrivs, ";ҽ������Ŀ¼;") <> 0      '��;ҽ������Ŀ¼,Ҳ���Ը���
    
'    mnuFilePara.Visible = InStr(1, mstrPrivs, ";��������;") <> 0
'    mnuFileSpt2.Visible = mnuFilePara.Visible
    
    bln���� = zlStr.IsHavePrivs(mstrPrivs, "�ۼ۹���") Or zlStr.IsHavePrivs(mstrPrivs, "�ɱ��۹���")

    mnuFileMlPreview.Visible = InStr(1, mstrPrivs, ";Ŀ¼��ӡ;") <> 0
    mnuFileMlPrint.Visible = InStr(1, mstrPrivs, ";Ŀ¼��ӡ;") <> 0
    
    '�۸��Ӳ˵�
    If InStr(1, mstrPrivs, ";��ѯ���ļ�Ŀ��;") = 0 And InStr(1, mstrPrivs, ";���ۼ�¼��ѯ;") = 0 Then '�Ӳ˵�������ʾ��ֱ�ӿ��Ƹ��˵�����ʾ
        mnuPrice.Visible = False
    Else
        mnuPriceLists.Visible = InStr(1, mstrPrivs, ";��ѯ���ļ�Ŀ��;") <> 0
        mnuPriceTable.Visible = InStr(1, mstrPrivs, ";���ۼ�¼��ѯ;") <> 0
    End If
    
    mnuEditSubsection.Visible = InStr(1, mstrPrivs, ";�ֶμӳ���;") <> 0
    mnuEditSpt3.Visible = mnuEditSubsection.Visible
    
    mnuEditSpecLimit.Visible = InStr(1, mstrPrivs, ";�����޿���;") <> 0 Or InStr(1, mstrPrivs, ";�̵���������;") <> 0
    
    mnuEditItemAdd.Visible = InStr(1, mstrPrivs, ";����Ʒ������;") <> 0
    mnuEditItemMod.Visible = InStr(1, mstrPrivs, ";����Ʒ������;") <> 0 Or _
                InStr(1, mstrPrivs, ";�������;") <> 0 Or _
                InStr(1, mstrPrivs, ";ָ���۸����;") <> 0 Or _
                InStr(1, mstrPrivs, ";�ۼ۹���;") <> 0 Or _
                InStr(1, mstrPrivs, ";�ɱ��۹���;") <> 0 Or _
                InStr(1, mstrPrivs, ";ҽ������Ŀ¼;") <> 0 Or _
                InStr(1, mstrPrivs, ";�������;") <> 0
    mnuEditPrice.Visible = InStr(1, mstrPrivs, ";�ۼ۹���;") <> 0 And InStr(1, mstrPrivs, ";�ɱ��۹���;") <> 0
    mnuEditItemDel.Visible = mnuEditItemAdd.Visible
    
    mnuEditAddName.Visible = mnuEditItemAdd.Visible
    mnuEditDeleName.Visible = mnuEditItemAdd.Visible
    mnuEditModifyName.Visible = mnuEditItemAdd.Visible
    
    mnuEditSpt1.Visible = mnuEditItemAdd.Visible Or mnuEditItemMod.Visible
    mnuEditStop.Visible = mnuEditItemAdd.Visible
    mnuEditStart.Visible = mnuEditItemAdd.Visible
    mnuEditSpt2.Visible = mnuEditItemAdd.Visible
    
    tlbThis.Buttons("����").Visible = mnuEditItemAdd.Visible
    tlbThis.Buttons("�޸�").Visible = blnModify    ' mnuEditItemMod.Visible
    tlbThis.Buttons("ɾ��").Visible = mnuEditItemAdd.Visible
    
    tlbThis.Buttons("Split1").Visible = mnuEditItemAdd.Visible
    tlbThis.Buttons("Start").Visible = mnuEditItemAdd.Visible
    tlbThis.Buttons("Stop").Visible = mnuEditItemAdd.Visible
    
    
    mnuClass.Visible = InStr(1, mstrPrivs, ";���ķ������;") <> 0
    tlbThis.Buttons("Class").Visible = mnuClass.Visible
    tlbThis.Buttons("Split4").Visible = mnuClass.Visible
    
    tlbThis.Buttons("Split").Visible = mnuEditItemAdd.Visible Or mnuClass.Visible Or mnuEditItemMod.Visible
    tlbThis.Buttons("Split4").Visible = mnuEditItemAdd.Visible And mnuClass.Visible
    
    
End Sub
Private Sub InitHgdPrivceHeadcol()
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ʼ�۸���ͷ
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    '�۸�������
    With Me.vsPrice
        .Redraw = flexRDNone
        .Rows = 2
        For intCol = 0 To .Cols - 1
            .TextMatrix(1, intCol) = ""
            .Cell(flexcpData, 1, intCol) = ""
        Next
        .RowData(1) = 0
        .Redraw = flexRDBuffered
        .ExtendLastCol = True
    End With
    With Me.vsStuff
        .Redraw = flexRDNone
        .Rows = 2
        For intCol = 0 To .Cols - 1
            .TextMatrix(1, intCol) = ""
            .Cell(flexcpData, 1, intCol) = ""
        Next
                 
        .RowData(1) = 0
        .Redraw = flexRDBuffered
        .ExtendLastCol = True
    End With
    vsCost.ExtendLastCol = True
End Sub
Private Sub Form_Resize()
    Dim lngTools As Single, lngStatus As Single
    
    SetParent txtFind.hwnd, tlbThis.hwnd
    SetParent picFind.hwnd, tlbThis.hwnd
    txtFind.Left = Me.ScaleWidth - txtFind.Width - 200
    picFind.Left = txtFind.Left - 100 - picFind.Width
    
    If WindowState = 1 Then Exit Sub
    lngTools = IIf(Me.clbThis.Visible, Me.clbThis.Height, 0)
    lngStatus = IIf(Me.stbThis.Visible, Me.stbThis.Height, 0)
    
    err = 0: On Error Resume Next
    
    With Me.picVBar_S
        .Top = lngTools
        .Height = Me.ScaleHeight - lngTools - lngStatus
        If .Left < 2000 Then .Left = 2000
        If .Left > Me.ScaleWidth - 4000 Then .Left = Me.ScaleWidth - 4000
    End With
    With Me.picHBar_S
        .Left = Me.picVBar_S.Left + Me.picVBar_S.Width
        .Width = Me.ScaleWidth - .Left
        If .Top < 2000 Then .Top = 2000
        If .Top > Me.ScaleHeight - lngStatus - 2500 Then .Top = Me.ScaleHeight - lngStatus - 2500
    End With
    
    With Me.picClass
        .Left = Me.ScaleLeft
        .Top = lngTools
        .Height = Me.ScaleHeight - picClass.Top - lngStatus
        .Width = Me.picVBar_S.Left - Me.picClass.Left
    End With
    
    With Me.lvwItems
        .Left = Me.picVBar_S.Left + Me.picVBar_S.Width
        .Top = lngTools
        .Height = Me.picHBar_S.Top - .Top
        .Width = Me.ScaleWidth - .Left
    End With
    
    With Me.picUp_S
        .Left = Me.picVBar_S.Left + Me.picVBar_S.Width
        .Top = Me.picHBar_S.Top + Me.picHBar_S.Height
        '.Height = Me.ScaleHeight - lngStatus - .Top + 15
        .Width = Me.ScaleWidth - .Left + 15
    End With
    With Me.TbList
        .Left = picUp_S.Left
        .Top = picUp_S.Height + picUp_S.Top + 10
        .Height = ScaleHeight - .Top - lngStatus
        .Width = ScaleWidth - .Left
    End With
    
    
'    With Me.fraComment
'        .Left = picUp_S.Left + 20
'        .Width = ScaleWidth - .Left
'        .Top = ScaleHeight - .Height - lngStatus
'    End With
'
'    With Me.vsPrice
'        .Left = picUp_S.Left
'        .Top = picUp_S.Height + picUp_S.Top + 10
'        .Width = picUp_S.Width - 30
'        .Height = fraComment.Top - .Top - 30
'    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
    Call SaveListViewState(Me.lvwItems, Me.Name & Val(Me.tvwClass.Tag), Me.lvwItems.View)
    
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\�ָ�", "����", Me.picVBar_S.Left)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\�ָ�", "����", Me.picHBar_S.Top)
    
    Call zlDatabase.SetPara("�����¼�����", IIf(mnuViewDownLevel.Checked = True, 1, 0), glngSys, mlngModule)
    zl_vsGrid_Para_Save mlngModule, vsCost, Me.Caption, "�ɱ��۵��ۼ�¼"
    zl_vsGrid_Para_Save mlngModule, vsPrice, Me.Caption, "�ۼ۵��ۼ�¼"
    zl_vsGrid_Para_Save mlngModule, vsStuff, Me.Caption, "���Ϲ���¼"
    mstrFindValue = ""
    Set mrsFind = Nothing
End Sub



Private Sub mnuEditAddName_Click()
    '����Ʒ��
    Dim lng����id As Long
    If Me.tvwClass.SelectedItem Is Nothing Then
        ShowMsgBox "��δ���÷���,������ɾ�������ϣ�"
        Exit Sub
    End If
    If tvwClass.SelectedItem Is Nothing Then
        lng����id = 0
    Else
        lng����id = Val(Mid(tvwClass.SelectedItem.Key, 2))
    End If
    If frmStuffBreed.ShowEditCard(Me, g����, "", lng����id, mstrPrivs) = False Then
        Exit Sub
    End If
    If lvwItems.SelectedItem Is Nothing Then
        Call zlRefRecords
    Else
        Call zlRefRecords(Val(Mid(lvwItems.SelectedItem.Key, 2)))
    End If
    Call SetMenuEnable
End Sub

Private Sub mnuEditDeleName_Click()
    Dim lng����ID As Long
    Dim rsSpec As New ADODB.Recordset
    On Error GoTo ErrHand
    
    With Me.lvwItems
        If .SelectedItem Is Nothing Then Exit Sub
        If MsgBox("���ɾ��Ʒ��Ϊ��" & .SelectedItem.Text & "����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        
        lng����ID = Mid(.SelectedItem.Key, 2)
        
        '����Ϊ:����id
        gstrSQL = "Zl_����Ʒ��_Delete(" & lng����ID & ")"
        zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
        
        Call .ListItems.Remove(.SelectedItem.Key)
        If .SelectedItem Is Nothing Then
            Call InitHgdPrivceHeadcol
        Else
            Call lvwItems_ItemClick(.SelectedItem)
        End If
    End With
    Call SetMenuEnable
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditExcel_Click()
    frmItemImport.ShowMe 3, Me
    
    Call zlRefClasses
    Call zlRefRecords
End Sub

Private Sub mnuEditModifyName_Click()

    Dim lng����id As Long, lng����ID As Long
    If lvwItems.SelectedItem Is Nothing Then Exit Sub
    lng����ID = Val(Mid(lvwItems.SelectedItem.Key, 2))
    If lng����ID = 0 Then Exit Sub
    lng����id = Val(lvwItems.SelectedItem.Tag)
    
    If Me.lvwItems.SelectedItem.Icon = "stop" Then
        ShowMsgBox "���ܶ�ͣ����������Ʒ�ֽ����޸ģ�"
        Exit Sub
    End If
    
    If frmStuffBreed.ShowEditCard(Me, g�޸�, lng����ID, lng����id, mstrPrivs) = False Then
        Exit Sub
    End If
    Call zlRefRecords(lng����ID)
    Call SetMenuEnable

End Sub

Private Sub mnuFilter_Click()
    With FrmStuffFilter
        Call .ShowMe(Me, mnuViewStoped.Checked)
    End With
End Sub

Public Sub zlGetFilter(ByVal strMaterialId As String)
    mstrMaterialID = strMaterialId
    Call cmdKind_Click(1)
End Sub
Private Sub picClass_Resize()
    Dim intCount As Integer
    err = 0: On Error Resume Next
    For intCount = Me.cmdKind.LBound To Me.cmdKind.UBound
        Me.cmdKind(intCount).Left = Me.picClass.ScaleLeft + 15
        Me.cmdKind(intCount).Width = Me.picClass.ScaleWidth
        Me.cmdKind(intCount).Height = 300
        If Val(Me.cmdKind(intCount).Tag) = 0 Then
            Me.cmdKind(intCount).Top = Me.picClass.ScaleTop + 285 * intCount
            Me.tvwClass.Top = Me.picClass.ScaleTop + 285 * (intCount + 1)
        Else
            Me.cmdKind(intCount).Top = Me.picClass.ScaleHeight - 285 * (Me.cmdKind.UBound - intCount + 1)
        End If
    Next
    Me.tvwClass.Left = Me.picClass.ScaleLeft + 15
    Me.tvwClass.Width = Me.picClass.ScaleWidth
    Me.tvwClass.Height = Me.picClass.ScaleHeight - 285 * (Me.cmdKind.UBound + 1) - 15
End Sub


Private Sub picCost_Resize()
    err = 0: On Error Resume Next
    With picCost
        vsCost.Move .ScaleLeft, .ScaleTop, .ScaleWidth, .ScaleHeight
    End With
End Sub

Private Sub TbList_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Select Case Item.Index
    Case 0
        zlControl.ControlSetFocus vsStuff, True
    Case 1
        zlControl.ControlSetFocus vsPrice, True
    Case Else
        zlControl.ControlSetFocus vsCost, True
    End Select
End Sub

Private Sub txtFind_GotFocus()
    zlControl.TxtSelAll txtFind
    OS.OpenIme True
End Sub

Private Sub txtfind_KeyPress(KeyAscii As Integer)
    Dim strTag As String
    Dim strSearch As String
        
    On Error GoTo ErrHandle
    If KeyAscii = vbKeyReturn Then
        zlControl.TxtSelAll txtFind
        If txtFind.Text = "" Then Exit Sub
        If mstrFindValue <> txtFind.Text And txtFind.Text <> "" Then
            mstrFindValue = txtFind.Text
            Set mrsFind = Nothing
            
            strSearch = " And (C.���� Like [1] OR N.���� Like [1] OR N.���� LIKE upper([1]))"
            gstrSQL = "" & _
                    "   Select distinct I.����ID,B.����ID,B.����ID" & _
                    "   From ������ĿĿ¼ I,�շ���Ŀ���� N,�������� B,�շ���ĿĿ¼ C" & _
                    "   Where   I.���='4' And I.id=b.����id and b.����ID=N.�շ�ϸĿid and b.����id=C.id " & strSearch
            If mnuViewStoped.Checked = False Then gstrSQL = gstrSQL & " And (C.����ʱ�� Is NULL Or C.����ʱ�� >=to_date('3000-01-01 00:00:00','yyyy-mm-dd hh24:mi:ss'))"
            Set mrsFind = zlDatabase.OpenSQLRecord(gstrSQL, "���ϲ�ѯ", UCase(Trim(txtFind.Text)) & "%")
                    
            If mrsFind.RecordCount > 0 Then
                Call zlLocateItem(mrsFind!����id, mrsFind!����id, mrsFind!����ID)
            End If
        Else
            If Not mrsFind.EOF Then
                mrsFind.MoveNext
                If Not mrsFind.EOF Then
                    Call zlLocateItem(mrsFind!����id, mrsFind!����id, mrsFind!����ID)
                Else
                    MsgBox "�Ѳ�ѯ�����һ����¼��", vbInformation, gstrSysName
                    mrsFind.MoveFirst
                    Call zlLocateItem(mrsFind!����id, mrsFind!����id, mrsFind!����ID)
                End If
            ElseIf mrsFind.RecordCount <> 0 And mrsFind.EOF Then
                mrsFind.MoveFirst
                Call zlLocateItem(mrsFind!����id, mrsFind!����id, mrsFind!����ID)
            End If
        End If
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsCost_DblClick()
    Call ShowCostBill
    
End Sub

Private Sub vsCost_GotFocus()
    With vsCost
         .SelectionMode = flexSelectionByRow
         .BackColorSel = &H8000000D
    End With
End Sub

Private Sub vsCost_LostFocus()
    With vsCost
         .BackColorSel = GRD_LOSTFOCUS_COLORSEL
    End With
End Sub

Private Sub ShowCostBill()
    '-----------------------------------------------------------------------------------------------------------
    '����:��ʾ�ɱ��۵��ɽ��
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-11-11 16:57:43
    '-----------------------------------------------------------------------------------------------------------
    Dim strNo As String
    
    With vsCost
        strNo = Trim(.TextMatrix(.Row, .ColIndex("NO")))
        If strNo = "" Then Exit Sub
        frmDiffPriceAdjustCard.ShowCard Me, strNo, 4, 1
    End With
    
    
End Sub


Private Sub vsPrice_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 2 Then Exit Sub
    If mnuPrice.Visible = False Then Exit Sub
    Call PopupMenu(Me.mnuPrice, 2)
End Sub

Private Sub vsPrice_RowColChange()
    
    Dim rsCheck As New ADODB.Recordset
    With Me.vsPrice
        Me.lblComment(3).Caption = "1��" & _
            .TextMatrix(.Row, .ColIndex("��������")) & "�������ϣ�" & _
             "ָ������" & .TextMatrix(.Row, .ColIndex("ָ������")) & "Ԫ/" & .TextMatrix(.Row, colPrice.��λ) & "��" & _
            "�ɹ�����" & .TextMatrix(.Row, .ColIndex("����")) & "%��"

            Me.lblComment(4).Caption = "2��" & _
            "ָ���ۼ�" & .TextMatrix(.Row, .ColIndex("ָ���ۼ�")) & "Ԫ/" & .TextMatrix(.Row, colPrice.��λ) & "��" & _
            "ָ������" & .TextMatrix(.Row, .ColIndex("ָ�������")) & "%��" & _
            IIf(Val(.TextMatrix(.Row, .ColIndex("���ηѱ�"))) = 0, "���ݲ�����ݷѱ�����Żݻ�Ӽۡ�", "���ܲ�����ݷѱ�Ӱ�졣")
    End With
End Sub
Private Sub lvwItems_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Me.lvwItems.SortKey = ColumnHeader.Index - 1 Then
        Me.lvwItems.SortOrder = IIf(Me.lvwItems.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Me.lvwItems.SortKey = ColumnHeader.Index - 1
        Me.lvwItems.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwItems_DblClick()
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    Call frmStuffBreed.ShowEditCard(Me, g�鿴, Mid(Me.lvwItems.SelectedItem.Key, 2), , mstrPrivs)
End Sub

Private Function Load�۸���Ϣ(ByVal lng����ID As Long) As Boolean
    '--------------------------------------------------------------------------------
    '����: �����������ϵļ۸���Ϣ
    '����:lng����ID-����ID
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2007/05/24
    '--------------------------------------------------------------------------------
    err = 0: On Error GoTo ErrHand
    
    For intCount = Me.lblComment.LBound To Me.lblComment.UBound
        Me.lblComment(intCount).Caption = ""
    Next
    '----------��д�۸�-----------------
    gstrSQL = " Select P.ID,p.NO ,p.���,decode(I.�Ƿ���,1,'ʱ��','����') as ����,  " & _
             "      nvl(S.ָ��������,0) as ָ������,nvl(S.����,0) as ����, " & _
             "      nvl(S.ָ�����ۼ�,0) as ָ���ۼ�,nvl(S.ָ�������,0) as ָ������,nvl(I.���ηѱ�,0)  as ���ηѱ�, " & _
             "      decode(sign(P.ִ������-sysdate),1,1,decode(sign(P.��ֹ����-sysdate),-1,-1,0)) as ִ�����, " & _
             "      '['||I.����||']'||I.����||' '||I.���||' '||I.���� as ����,I.���㵥λ as ��λ,S.��װ��λ,S.����ϵ��,S.����ID, " & _
             "      P.�ּ� as �ۼ�,U.���� as ������Ŀ,P.����˵��, " & _
             "      to_char(P.ִ������,'YYYY-MM-DD HH24:MI:SS') as ִ������ " & _
             "   From �շѼ�Ŀ P,������Ŀ U,�շ���ĿĿ¼ I,�������� S" & _
             "   Where P.�շ�ϸĿID=I.ID and P.������ĿID=U.ID and I.ID=S.����ID" & _
             "       and S.����ID=[1]" & GetPriceClassString("P")
    
    If Me.mnuViewPrices.Checked = False Then
        gstrSQL = gstrSQL & "       and (P.��ֹ���� is null or P.��ֹ����>=sysdate)"
    End If
    gstrSQL = gstrSQL & " order by I.����,P.ִ������ desc"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng����ID)

    With rsTemp
        Me.vsPrice.Redraw = flexRDNone
        If .BOF Or .EOF Then
            With Me.vsPrice
                .Rows = .FixedRows + 1: .RowData(.FixedRows) = 0
                For intCol = 0 To .Cols - 1
                    .TextMatrix(.FixedRows, intCol) = ""
                Next
            End With
        Else
            Me.vsPrice.Rows = Me.vsPrice.FixedRows + .RecordCount
        End If

        Do While Not .EOF
            Me.vsPrice.RowData(.AbsolutePosition) = Val(zlStr.Nvl(!Id))
            Me.vsPrice.TextMatrix(.AbsolutePosition, vsPrice.ColIndex("NO")) = zlStr.Nvl(!NO)
            Me.vsPrice.TextMatrix(.AbsolutePosition, vsPrice.ColIndex("���")) = zlStr.Nvl(!���)
            Me.vsPrice.TextMatrix(.AbsolutePosition, vsPrice.ColIndex("��������")) = zlStr.Nvl(!����)
            Me.vsPrice.TextMatrix(.AbsolutePosition, vsPrice.ColIndex("ָ������")) = Format(!ָ������ * IIf(mintUnit = 0, 1, zlStr.Nvl(!����ϵ��, 1)), mFMT.FM_�ɱ���)
            Me.vsPrice.TextMatrix(.AbsolutePosition, vsPrice.ColIndex("����")) = Format(!����, GFM_VBKL)
            Me.vsPrice.TextMatrix(.AbsolutePosition, vsPrice.ColIndex("ָ���ۼ�")) = Format(!ָ���ۼ� * IIf(mintUnit = 0, 1, zlStr.Nvl(!����ϵ��, 1)), mFMT.FM_���ۼ�)
            Me.vsPrice.TextMatrix(.AbsolutePosition, vsPrice.ColIndex("ָ�������")) = Format(!ָ������, GFM_VBCJL)

            Me.vsPrice.TextMatrix(.AbsolutePosition, vsPrice.ColIndex("ִ�����")) = zlStr.Nvl(!ִ�����)
            Me.vsPrice.TextMatrix(.AbsolutePosition, vsPrice.ColIndex("����")) = zlStr.Nvl(!����)
            Me.vsPrice.Cell(flexcpData, .AbsolutePosition, vsPrice.ColIndex("����")) = zlStr.Nvl(!����ID)
            Me.vsPrice.TextMatrix(.AbsolutePosition, vsPrice.ColIndex("��λ")) = IIf(mintUnit = 0, zlStr.Nvl(!��λ), zlStr.Nvl(!��װ��λ))
            Me.vsPrice.TextMatrix(.AbsolutePosition, vsPrice.ColIndex("�ۼ�")) = Format(!�ۼ� * IIf(mintUnit = 0, 1, zlStr.Nvl(!����ϵ��, 1)), mFMT.FM_���ۼ�)

            Me.vsPrice.TextMatrix(.AbsolutePosition, vsPrice.ColIndex("������Ŀ")) = zlStr.Nvl(!������Ŀ)
            Me.vsPrice.TextMatrix(.AbsolutePosition, vsPrice.ColIndex("˵��")) = zlStr.Nvl(!����˵��)
            Me.vsPrice.TextMatrix(.AbsolutePosition, vsPrice.ColIndex("ִ������")) = zlStr.Nvl(!ִ������)
            Me.vsPrice.TextMatrix(.AbsolutePosition, vsPrice.ColIndex("���ηѱ�")) = zlStr.Nvl(!���ηѱ�)

            Me.vsPrice.Row = .AbsolutePosition
            For intCol = 0 To Me.vsPrice.Cols - 1
                Me.vsPrice.Col = intCol
                Select Case !ִ�����
                Case -1     '��ִ��
                    Me.vsPrice.CellBackColor = RGB(240, 240, 240)
                Case 0      '����ִ��
                    Me.vsPrice.CellBackColor = RGB(255, 255, 255)
                Case 1      'δִ��
                    Me.vsPrice.CellBackColor = RGB(225, 255, 255)
                End Select
            Next
            .MoveNext
        Loop

        Me.vsPrice.Row = Me.vsPrice.FixedRows
        If Me.vsPrice.ColWidth(vsPrice.ColIndex("����")) = 0 Or Me.vsPrice.ColWidth(vsPrice.ColIndex("��λ")) = 0 Then
            Me.vsPrice.ColWidth(vsPrice.ColIndex("����")) = 3500
            Me.vsPrice.ColWidth(vsPrice.ColIndex("��λ")) = 550
        End If
        Me.vsPrice.Redraw = flexRDBuffered
    End With
    Call vsPrice_RowColChange
    Load�۸���Ϣ = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub SetVsGridColor(ByVal objVsGrid As Object, ByVal lngRow As Long, ByVal OleColor As OLE_COLOR, ByVal blnBackColor As Boolean)
    '--------------------------------------------------------------------------------
    '����: ����ָ���еı�����ǰ��ɫ
    '����:objVsGrid-ָ��������ؼ�
    '     lngRow -ָ������
    '     olecolor-ָ������ɫֵ
    '     blnBackColor-���ñ���ɫ
    '����:���˺�
    '����:2007/05/24
    '--------------------------------------------------------------------------------
    Dim lngCol As Long
    Dim lngCurRow As Long, lngCurCol As Long
    With objVsGrid
        lngCurRow = .Row: lngCurCol = .Col
        For intCol = 0 To .Cols - 1
            .Col = intCol
            If blnBackColor = False Then
                .CellForeColor = OleColor
            Else
                .CellBackColor = OleColor
            End If
        Next
        .Row = lngCurRow: .Col = lngCurCol
    End With
End Sub
Private Function Load���Ϲ����Ϣ(ByVal lng����ID As Long) As Boolean
    '--------------------------------------------------------------------------------
    '����: �����������ϵĹ����Ϣ
    '����:lng����ID-����ID
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2007/05/24
    '--------------------------------------------------------------------------------
    Dim lngRow As Long
    Dim lngPre����ID As Long, lngLocaleRow As Long
    err = 0: On Error GoTo ErrHand:
    gstrSQL = " " & _
         " SELECT  Distinct b.id,a.����id, b.����,b.����,b.���,b.����,b.���㵥λ AS ɢװ��λ,a.��װ��λ, a.����ϵ��,b.�������,b.��������,a.�ɱ���," & _
         "      a.���ʷ���,a.�洢����,a.���֤��,a.���֤��Ч��,b.��ʶ����,b.��ʶ����,decode(b.�Ƿ���,1,'ʵ��','����') as �Ƿ���," & _
         "      a.���Ч��,a.���Ч��,decode(a.�б����,1,'��','��') as �б����,decode(a.�޾��Բ���,1,'��','��')  �޾��Բ���, " & _
         "      decode(a.һ���Բ���,1,'��','��') һ���Բ���  ,  decode(a.���Ʋ���,1,'��','��') ���Ʋ��� ,a.��Դ���,a.������Դ, " & _
         "      a.ָ��������,a.ָ�����ۼ�,a.ָ�������,a.����, decode(a.�ⷿ����,1,'��','��') �ⷿ����,decode(a.���÷���,1,'��','��') ���÷���," & _
         "      decode(A.ԭ����,1,'��','��') ԭ����,  decode(a.��������,1,'��','��') ��������,decode(a.�������,1,'��','��') �������," & _
         "      to_Char(b.����ʱ��,'yyyy-mm-dd') as  ����ʱ��," & _
         "      nvl(b.����ʱ��,to_date('3000-01-01','YYYY-MM-DD')) as ����ʱ��,C.���� As ��Ʒ��, A.��ֵ���� " & _
         " FROM �������� a,�շ���ĿĿ¼ b, �շ���Ŀ���� C " & _
         " WHERE a.����id=b.id  And B.ID = C.�շ�ϸĿid(+) And C.����(+) = 3 and a.����id=[1] "
 
    If Me.mnuViewStoped.Checked = False Then
        gstrSQL = gstrSQL & " and (B.����ʱ�� is null or B.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))"
    End If
    gstrSQL = gstrSQL & " order by B.����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng����ID)
    
    lngLocaleRow = -1
    With vsStuff
        .Redraw = flexRDNone
        lngPre����ID = .RowData(.Row)
        If rsTemp.EOF Then
            .Rows = .FixedRows + 1: .RowData(.FixedRows) = 0
            For intCol = 0 To .Cols - 1
                .TextMatrix(.FixedRows, intCol) = ""
            Next
        Else
            .Rows = .FixedRows + rsTemp.RecordCount
        End If
        
        lngRow = 1
        Do While Not rsTemp.EOF
            .RowData(lngRow) = Val(zlStr.Nvl(rsTemp!Id))
            
            .Cell(flexcpData, lngRow, .ColIndex("����")) = zlStr.Nvl(rsTemp!����id)
            .TextMatrix(lngRow, .ColIndex("ID")) = zlStr.Nvl(rsTemp!Id)
            .TextMatrix(lngRow, .ColIndex("����")) = zlStr.Nvl(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("���")) = zlStr.Nvl(rsTemp!���)
            .TextMatrix(lngRow, .ColIndex("����")) = zlStr.Nvl(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("��Ʒ��")) = zlStr.Nvl(rsTemp!��Ʒ��)
            .TextMatrix(lngRow, .ColIndex("ɢװ��λ")) = zlStr.Nvl(rsTemp!ɢװ��λ)
            .TextMatrix(lngRow, .ColIndex("��װ��λ")) = zlStr.Nvl(rsTemp!��װ��λ)
            .TextMatrix(lngRow, .ColIndex("����ϵ��")) = zlStr.Nvl(rsTemp!����ϵ��)
            If zlStr.Nvl(rsTemp!�������) = 1 Then
                .TextMatrix(lngRow, .ColIndex("�������")) = "����"
            ElseIf zlStr.Nvl(rsTemp!�������) = 2 Then
                .TextMatrix(lngRow, .ColIndex("�������")) = "סԺ"
            ElseIf zlStr.Nvl(rsTemp!�������) = 3 Then
                .TextMatrix(lngRow, .ColIndex("�������")) = "�����סԺ"
            Else
                .TextMatrix(lngRow, .ColIndex("�������")) = "��Ӧ���ڲ���"
            End If
            .TextMatrix(lngRow, .ColIndex("ҽ������")) = zlStr.Nvl(rsTemp!��������)
            .TextMatrix(lngRow, .ColIndex("���ʷ���")) = zlStr.Nvl(rsTemp!���ʷ���)
            .TextMatrix(lngRow, .ColIndex("�洢����")) = zlStr.Nvl(rsTemp!�洢����)
            .TextMatrix(lngRow, .ColIndex("���֤��")) = zlStr.Nvl(rsTemp!���֤��)
            .TextMatrix(lngRow, .ColIndex("���֤Ч��")) = zlStr.Nvl(rsTemp!���֤��Ч��)
            .TextMatrix(lngRow, .ColIndex("��ʶ����")) = zlStr.Nvl(rsTemp!��ʶ����)
            .TextMatrix(lngRow, .ColIndex("��ʶ����")) = zlStr.Nvl(rsTemp!��ʶ����)
            .TextMatrix(lngRow, .ColIndex("���Ч��")) = zlStr.Nvl(rsTemp!���Ч��)
            .TextMatrix(lngRow, .ColIndex("���Ч��")) = zlStr.Nvl(rsTemp!���Ч��)
            .TextMatrix(lngRow, .ColIndex("�б����")) = zlStr.Nvl(rsTemp!�б����)

            .TextMatrix(lngRow, .ColIndex("�޾�����")) = zlStr.Nvl(rsTemp!�޾��Բ���)
            .TextMatrix(lngRow, .ColIndex("һ���Բ���")) = zlStr.Nvl(rsTemp!һ���Բ���)
            .TextMatrix(lngRow, .ColIndex("���Ʋ���")) = zlStr.Nvl(rsTemp!���Ʋ���)
            .TextMatrix(lngRow, .ColIndex("��Դ���")) = zlStr.Nvl(rsTemp!��Դ���)
            .TextMatrix(lngRow, .ColIndex("������Դ")) = zlStr.Nvl(rsTemp!������Դ)
            .TextMatrix(lngRow, .ColIndex("ָ������")) = Format(rsTemp!ָ�������� * IIf(mintUnit = 0, 1, zlStr.Nvl(rsTemp!����ϵ��, 1)), mFMT.FM_�ɱ���)
            .TextMatrix(lngRow, .ColIndex("�ɱ���")) = Format(rsTemp!�ɱ��� * IIf(mintUnit = 0, 1, zlStr.Nvl(rsTemp!����ϵ��, 1)), mFMT.FM_�ɱ���)
            
            .TextMatrix(lngRow, .ColIndex("ָ���ۼ�")) = Format(rsTemp!ָ�����ۼ� * IIf(mintUnit = 0, 1, zlStr.Nvl(rsTemp!����ϵ��, 1)), mFMT.FM_���ۼ�)
            .TextMatrix(lngRow, .ColIndex("ָ�������")) = Format(rsTemp!ָ�������, GFM_VBCJL)
            .TextMatrix(lngRow, .ColIndex("����")) = Format(rsTemp!����, GFM_VBKL)
            .TextMatrix(lngRow, .ColIndex("�ⷿ����")) = zlStr.Nvl(rsTemp!�ⷿ����)
            .TextMatrix(lngRow, .ColIndex("���÷���")) = zlStr.Nvl(rsTemp!���÷���)
            .TextMatrix(lngRow, .ColIndex("ԭ�ϲ���")) = zlStr.Nvl(rsTemp!ԭ����)
            .TextMatrix(lngRow, .ColIndex("��������")) = zlStr.Nvl(rsTemp!��������)
            .TextMatrix(lngRow, .ColIndex("�������")) = zlStr.Nvl(rsTemp!�������)
            .TextMatrix(lngRow, .ColIndex("��ֵ����")) = IIf(IsNull(rsTemp!��ֵ����) Or rsTemp!��ֵ���� = "0", "��", "��")
            .TextMatrix(lngRow, .ColIndex("����ʱ��")) = zlStr.Nvl(rsTemp!����ʱ��)
            If .RowData(lngRow) = lngPre����ID Then
                lngLocaleRow = lngRow
            End If
            .Row = lngRow
            If Format(rsTemp!����ʱ��, "YYYY-MM-DD") <> "3000-01-01" Then
                .TextMatrix(lngRow, .ColIndex("����ʱ��")) = zlStr.Nvl(rsTemp!����ʱ��)
                Call SetVsGridColor(vsStuff, lngRow, &HFF&, False)
            Else
                .TextMatrix(lngRow, .ColIndex("����ʱ��")) = ""
                '������б���ϣ�����ɫ�����Ƿ���ʱ�ۻ��Ƕ��۲���
                If Trim(zlStr.Nvl(rsTemp!�б����)) = "��" Then
                    Call SetVsGridColor(vsStuff, lngRow, IIf(rsTemp!�Ƿ��� = "����", &H800000, &H800080), False)
                Else
                    Call SetVsGridColor(vsStuff, lngRow, IIf(rsTemp!�Ƿ��� = "����", &H0, &H40&), False)
                End If
            End If
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        If lngLocaleRow > 0 Then
            .Row = lngLocaleRow
        Else
            .Row = 1
        End If
        Call .ShowCell(.Row, .ColIndex("����"))
        .Redraw = flexRDBuffered
    End With
   Load���Ϲ����Ϣ = True
   Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Sub lvwItems_ItemClick(ByVal Item As MSComctlLib.ListItem)
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    Dim lng����ID As Long
    
    err = 0: On Error GoTo ErrHand
    lng����ID = Val(Mid(Item.Key, 2))
    Call Load�۸���Ϣ(lng����ID)
    Call LoadCostData(lng����ID)
    '���ع����Ϣ:
    Call Load���Ϲ����Ϣ(lng����ID)
    Call SetMenuEnable
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub lvwItems_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    Call lvwItems_DblClick
End Sub

Private Sub lvwItems_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 2 Then Exit Sub
   
   PopupMenu mnuEdit
End Sub
Private Sub mnuClassAdd_Click()
    With frmClinicClass
        .lblKind.Tag = 7
        If Me.tvwClass.SelectedItem Is Nothing Then
            .txtParent.Tag = 0
        Else
            .txtParent.Tag = Mid(Me.tvwClass.SelectedItem.Key, 2)
        End If
        .Tag = "����"
        .Show 1, Me
    End With
    
    Call zlRefClasses
End Sub

Private Sub mnuClassDel_Click()
    If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
    If MsgBox("���ɾ���÷��ࡰ" & Me.tvwClass.SelectedItem.Text & "����", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    err = 0: On Error GoTo ErrHand
    gstrSQL = "zl_���Ʒ���Ŀ¼_delete(" & Mid(Me.tvwClass.SelectedItem.Key, 2) & ")"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
                    
    
    Dim strParentKey As String
    If Me.tvwClass.SelectedItem.Next Is Nothing Then
        If Me.tvwClass.SelectedItem.Parent Is Nothing Then
            Call zlRefClasses
        Else
            strParentKey = Me.tvwClass.SelectedItem.Parent.Key
            Call Me.tvwClass.Nodes.Remove(Me.tvwClass.SelectedItem.Key)
            If Me.tvwClass.SelectedItem Is Nothing Then
                    Me.lvwItems.ListItems.Clear
                    Call InitHgdPrivceHeadcol
            Else
                Call zlRefRecords
            End If
        End If
    Else
        Call zlRefClasses
    End If
    SetMenuEnable
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuClassMod_Click()

    If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
    With frmClinicClass
        .lblKind.Tag = 7
        If Me.tvwClass.SelectedItem.Parent Is Nothing Then
            .txtParent.Tag = 0
            .txtParent.Text = "(��)"
            .txtUpCode.Text = ""
            .txtCode.Text = Mid(Split(Me.tvwClass.SelectedItem.Text, "]")(0), 2)
            .txtCode.MaxLength = Len(.txtCode.Text)
            .txtCode.Tag = .txtCode.MaxLength
        Else
            .txtParent.Tag = Mid(Me.tvwClass.SelectedItem.Parent.Key, 2)
            .txtParent.Text = Me.tvwClass.SelectedItem.Parent.Text
            .txtUpCode.Text = Mid(Split(Me.tvwClass.SelectedItem.Parent.Text, "]")(0), 2)
            .txtCode.Text = Mid(Split(Me.tvwClass.SelectedItem.Text, "]")(0), Len(.txtUpCode.Text) + 2)
            .txtCode.MaxLength = Len(.txtCode.Text)
            .txtCode.Tag = .txtCode.MaxLength
        End If
        .txtName = Split(Me.tvwClass.SelectedItem.Text, "]")(1)
        .txtSymbol = Me.tvwClass.SelectedItem.Tag
        .Tag = Mid(Me.tvwClass.SelectedItem.Key, 2)
        .Show 1, Me
    End With
    Call zlRefClasses
End Sub

Private Sub mnuEditItemAdd_Click()
    Dim lng����ID As Long
    Dim blnStop As Boolean
    Dim lng����id As Long
    
    If Me.lvwItems.ListItems.count = 0 Then
        ShowMsgBox "��δ����Ʒ��,������ɾ�������ϵĹ��"
        Exit Sub
    End If
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    
    blnStop = Me.lvwItems.SelectedItem.Icon = "stop"
    If blnStop Then
        ShowMsgBox "��Ʒ���Ѿ���ͣ��,�������ӹ��!"
        Exit Sub
    End If
    
    lng����ID = Val(Mid(Me.lvwItems.SelectedItem.Key, 2))
    
    If tvwClass.SelectedItem Is Nothing Then
        lng����id = 0
    Else
        lng����id = Val(Mid(tvwClass.SelectedItem.Key, 2))
    End If
    
    If frmStuffSpec.ShowEditCard(Me, g����, lng����ID, lng����id, "", mstrPrivs) = False Then
        Exit Sub
    End If
    Call Load���Ϲ����Ϣ(Val(Mid(lvwItems.SelectedItem.Key, 2)))
    Call SetMenuEnable
    
End Sub

Private Sub mnuEditItemDel_Click()
    Dim lng����ID As Long
    Dim lng����ID As Long
    Dim blnStop As Boolean
    Dim str��� As String
    With vsStuff
          lng����ID = .RowData(.Row)
          lng����ID = Val(.Cell(flexcpData, .Row, .ColIndex("����")))
          blnStop = Trim(.TextMatrix(.Row, .ColIndex("����ʱ��"))) <> ""
          str��� = Trim(.TextMatrix(.Row, .ColIndex("���")))
    End With
    If lng����ID = 0 Then Exit Sub
    If lng����ID = 0 Then Exit Sub
    
    On Error GoTo ErrHand
    If MsgBox("���ɾ�����Ϊ��" & str��� & "����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        
    '����Ϊ:����id
    gstrSQL = "zl_��������_DELETE(" & lng����ID & ")"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    With vsStuff
        If .Rows = 2 Then
            Call InitHgdPrivceHeadcol
        Else
            .RemoveItem .Row
        End If
    End With
    Call SetMenuEnable
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditItemMod_Click()
    Dim lng����ID As Long
    Dim lng����ID As Long
    Dim blnStop As Boolean
    Dim lng����id As Long
    
    With vsStuff
          lng����ID = .RowData(.Row)
          lng����ID = Val(.Cell(flexcpData, .Row, .ColIndex("����")))
          blnStop = Trim(.TextMatrix(.Row, .ColIndex("����ʱ��"))) <> ""
    End With
    If lng����ID = 0 Then Exit Sub
    If lng����ID = 0 Then Exit Sub
    If blnStop Then
        ShowMsgBox "���ܶ�ͣ���������Ͻ����޸ģ�"
        Exit Sub
    End If
    If tvwClass.SelectedItem Is Nothing Then
        lng����id = 0
    Else
        lng����id = Val(Mid(tvwClass.SelectedItem.Key, 2))
    End If
    If frmStuffSpec.ShowEditCard(Me, g�޸�, lng����ID, lng����id, CStr(lng����ID), mstrPrivs) = False Then
        Exit Sub
    End If
    If Me.lvwItems.SelectedItem Is Nothing Then
        Call InitHgdPrivceHeadcol
    Else
        Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
    End If
    Call SetMenuEnable
    
End Sub
Private Function Get����ID() As Long

    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "Select ����id From ������ĿĿ¼ where id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Me.lvwItems.SelectedItem.Tag))
    If rsTemp.EOF Then
        Get����ID = 0
    Else
        Get����ID = Val(zlStr.Nvl(rsTemp!����id))
    End If
    rsTemp.Close
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub mnuEditItemPart_Click()
    Dim bln�༭ As Boolean, lng����ID As Long
    With vsStuff
        lng����ID = .RowData(.Row)
    End With
    With frm�洢�ⷿ
        bln�༭ = (InStr(1, mstrPrivs, ";�洢�ⷿ;") <> 0)
        Call .ShowMe(Me, lng����ID, bln�༭)
    End With
End Sub



Private Sub mnuEditSpecLimit_Click()
    With frmStuffLimit
        If InStr(1, mstrPrivs, ";�����޿���;") = 0 And InStr(1, mstrPrivs, ";�̵���������;") = 0 Then
            .cmdClose.Tag = "����"
        Else
            .cmdClose.Tag = "�޸�"
        End If
        .strPrivs = mstrPrivs
        .Show 1, Me
    End With
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    
    Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
End Sub

Private Sub mnuEditSpecSelf_Click()
    Dim lng����ID As Long
    
    With vsStuff
        lng����ID = .RowData(.Row)
    End With
'
'    If Me.lvwItems.SelectedItem Is Nothing Then
'        lng����id = 0
'    Else
'        lng����id = Val(Mid(Me.lvwItems.SelectedItem.Key, 2))
'    End If

    With frmStuffMember
        If InStr(1, mstrPrivs, ";���Ʋ��Ϲ���;") = 0 Then
            .cmdClose.Tag = "����"
        Else
            .cmdClose.Tag = "�޸�"
        End If
          
        .lblMedi.Tag = lng����ID
        .msfMember.Tag = "����"
        .Show 1, Me
    End With

    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
End Sub
Private Sub mnuEditStart_Click()
    Dim lng����ID As Long
    Dim lng����ID As Long
    Dim blnStop As Boolean
    
    If Me.ActiveControl Is lvwItems Then
            With Me.lvwItems
                If .SelectedItem Is Nothing Then Exit Sub
                If .SelectedItem.Icon = "start" Then Exit Sub
                
                If MsgBox("�����������Ʒ��Ϊ��" & .SelectedItem.Text & "����", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
                
                '�շ���ĿĿ¼ID(����ID)
                gstrSQL = "Zl_����Ʒ��_Reuse(" & Val(Mid(.SelectedItem.Key, 2)) & ")"
                
                err = 0: On Error GoTo ErrHand
                zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
                .SelectedItem.SubItems(Me.lvwItems.ColumnHeaders("_����ʱ��").Index - 1) = ""
                                
                .SelectedItem.Icon = "start": .SelectedItem.SmallIcon = "start"
                '�ָ�������Ŀ��ʾ��ɫ
                .SelectedItem.ForeColor = .ForeColor
                For intCount = 1 To .ColumnHeaders.count - 1
                    .SelectedItem.ListSubItems(intCount).ForeColor = .ForeColor
                Next
                
            End With
    Else
            With Me.vsStuff
                lng����ID = .RowData(.Row)
                lng����ID = Val(.Cell(flexcpData, .Row, .ColIndex("����")))
                blnStop = Trim(.TextMatrix(.Row, .ColIndex("����ʱ��"))) <> ""
                
                If lng����ID = 0 Then Exit Sub
                If blnStop = False Then Exit Sub
                
                If MsgBox("����������ù��Ϊ:��" & .TextMatrix(.Row, .ColIndex("���")) & "����", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
                
                '�շ���ĿĿ¼ID(����ID)
                gstrSQL = "zl_��������_REUSE(" & lng����ID & ")"
                
                err = 0: On Error GoTo ErrHand
                zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
                .Cell(flexcpText, .Row, .ColIndex("����ʱ��")) = ""
                Call SetVsGridColor(vsStuff, .Row, .ForeColor, False)
                
                With lvwItems
                    lvwItems.SelectedItem.SubItems(Me.lvwItems.ColumnHeaders("_����ʱ��").Index - 1) = ""
                    lvwItems.SelectedItem.Icon = "start": lvwItems.SelectedItem.SmallIcon = "start"
                    '�ָ�������Ŀ��ʾ��ɫ
                    .SelectedItem.ForeColor = .ForeColor
                    For intCount = 1 To .ColumnHeaders.count - 1
                        .SelectedItem.ListSubItems(intCount).ForeColor = .ForeColor
                    Next
                End With
            End With
    End If
    Call SetMenuEnable
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditStop_Click()
    Dim lng����ID As Long
    Dim lng����ID As Long
    Dim blnStop As Boolean
    
    If Me.ActiveControl Is lvwItems Then
        With Me.lvwItems
            If .SelectedItem Is Nothing Then Exit Sub
            If .SelectedItem.Icon = "stop" Then Exit Sub
            If MsgBox("���Ҫͣ��Ʒ��Ϊ��" & .SelectedItem.Text & "����", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
            
            gstrSQL = "Zl_����Ʒ��_Stop(" & Val(Mid(.SelectedItem.Key, 2)) & ")"
        
            err = 0: On Error GoTo ErrHand
            zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
            
            If Me.mnuViewStoped.Checked = True Then
                .SelectedItem.Icon = "stop": .SelectedItem.SmallIcon = "stop"
                '��ͣ����Ŀ��ʾΪ��ɫ
                .SelectedItem.SubItems(Me.lvwItems.ColumnHeaders("_����ʱ��").Index - 1) = Format(sys.Currentdate, "yyyy-mm-dd")
                .SelectedItem.ForeColor = &HFF&
                For intCount = 1 To .ColumnHeaders.count - 1
                    .SelectedItem.ListSubItems(intCount).ForeColor = &HFF&
                Next
                
            Else
                Call .ListItems.Remove(.SelectedItem.Key)
            End If
            
        End With
    Else
            With Me.vsStuff
                lng����ID = .RowData(.Row)
                lng����ID = Val(.Cell(flexcpData, .Row, .ColIndex("����")))
                blnStop = Trim(.TextMatrix(.Row, .ColIndex("����ʱ��"))) <> ""
                
                If lng����ID = 0 Then Exit Sub
                If blnStop Then Exit Sub
                
                If MsgBox("�������ͣ�ù��Ϊ��" & .TextMatrix(.Row, .ColIndex("���")) & "����", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
                
                '�շ���ĿĿ¼ID(����ID)
                gstrSQL = "zl_��������_STOP(" & lng����ID & ")"
                
                err = 0: On Error GoTo ErrHand
                zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
                
                If Me.mnuViewStoped.Checked = True Then
                    .Cell(flexcpText, .Row, .ColIndex("����ʱ��")) = Format(sys.Currentdate, "yyyy-mm-dd")
                    Call SetVsGridColor(vsStuff, .Row, &HFF&, False)
                Else
                    If .Rows = 2 Then
                        Call InitHgdPrivceHeadcol
                    Else
                        .RemoveItem .Row
                    End If
                End If
            End With
    End If
    Call SetMenuEnable
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditSubsection_Click()
    Dim blnreturn As Boolean
    frm�ӳ�������.EditCard Me, mstrPrivs, blnreturn
End Sub

Private Sub mnuEditUnit_Click()
    Dim lng����ID  As Long, lng����ID As Long
    '�б굥λ
    '�б�����б굥λ����
    With vsStuff
          lng����ID = .RowData(.Row)
          lng����ID = Val(.Cell(flexcpData, .Row, .ColIndex("����")))
    End With
    
    With frmStuffUnitMgr
        Dim strType As String

        If lng����ID = 0 Then   '���û�м�¼���˳�����Ϊ�޷��ж�ҩƷ����
            .lblTag = 0
        Else
            .lblTag = lng����ID
        End If
        
        .frmTag = 4
        .strPrivs = mstrPrivs
        If InStr(1, mstrPrivs, ";�б굥λ;") <> 0 Then
            .cmdClose.Tag = ""
        Else
            .cmdClose.Tag = "����"
        End If
        .Show 1, Me
    End With
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
End Sub

Private Sub mnuFileExcel_Click()
    Call zlRptPrint(3)
End Sub

Private Sub mnuFileMlPreview_Click()
    Dim strPara As String
    If Val(Mid(Me.tvwClass.SelectedItem.Key, 2)) > 0 Then
        strPara = " ID=" & Val(Mid(Me.tvwClass.SelectedItem.Key, 2))
    Else
        strPara = " �ϼ�ID is null"
    End If
    Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1711", Me, "����=" & strPara & " ", 1)
End Sub

Private Sub mnuFileMlPrint_Click()
    Dim strPara As String
    If Val(Mid(Me.tvwClass.SelectedItem.Key, 2)) > 0 Then
        strPara = " ID=" & Val(Mid(Me.tvwClass.SelectedItem.Key, 2))
    Else
        strPara = " �ϼ�ID is null"
    End If
    Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1711", Me, "����=" & strPara & " ", 2)

End Sub

Private Sub mnuFilePara_Click()
    'ģ�鹫�������Ѿ�������ҩƷ��������ģ�飬Ŀǰû��˽�л򱾻���������ʱ���β������ý���
'   frmStuffPara.ShowMe mstrPrivs, Me
   'mbln��ʾ�¼� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\������ʾģʽ\", "��ʾ�¼�", 0)) = 1
   
End Sub

Private Sub mnuFilePreView_Click()
    Call zlRptPrint(0)
End Sub

Private Sub mnuFilePrint_Click()
    Call zlRptPrint(1)
End Sub

Private Sub mnuFilePrintSet_Click()
    Call zlPrintSet
End Sub

Private Sub mnuHelpAbout_Click()
    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
End Sub

Private Sub mnuHelpTitle_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
End Sub

Private Sub mnufileexit_Click()
    Unload Me
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hwnd)
End Sub


Private Sub mnuPriceLists_Click()
    Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1711_2", Me)
End Sub

Private Sub mnuPriceTable_Click()
    Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1711_1", Me)
End Sub

Private Sub mnuViewDownLevel_Click()
    mnuViewDownLevel.Checked = Not mnuViewDownLevel.Checked
    mbln��ʾ�¼� = mnuViewDownLevel.Checked
    Call mnuViewRefresh_Click
End Sub

Private Sub mnuViewFind_Click()
    With frmStuffFind
        Call .ShowMe(Me, mnuViewStoped.Checked)
    End With
End Sub

Private Sub mnuViewIcon_Click(Index As Integer)
   Dim intTemp As Integer
    For intTemp = 0 To 3
        mnuViewIcon(intTemp).Checked = False
    Next
    mnuViewIcon(Index).Checked = True
    lvwItems.View = Index
    lvwItems.Refresh
End Sub

Private Sub mnuViewPrices_Click()
    Me.mnuViewPrices.Checked = Not Me.mnuViewPrices.Checked
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
End Sub

Private Sub mnuViewRefresh_Click()
    If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
    Call zlRefRecords
End Sub

Private Sub mnuViewStates_Click()
    Me.mnuViewStates.Checked = Not Me.mnuViewStates.Checked
    Me.stbThis.Visible = Me.mnuViewStates.Checked
    Form_Resize
End Sub

Private Sub mnuViewStoped_Click()
    Me.mnuViewStoped.Checked = Not Me.mnuViewStoped.Checked
    Call zlRefRecords
End Sub
Private Sub mnuReportItem_Click(Index As Integer)
    Dim lng����id As Long
    Dim lng����ID As Long
    
    If tvwClass.SelectedItem Is Nothing Then
        lng����id = 0
    Else
        lng����id = Val(Mid(tvwClass.SelectedItem.Key, 2))
    End If
    
    lng����ID = 0
    If Not Me.lvwItems.SelectedItem Is Nothing Then
        lng����ID = Val(Mid(Me.lvwItems.SelectedItem.Key, 2))
    End If
    
    '2006-04-25:���˺�:�����Զ��屨������ģ��Ĺ���
    Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, "����=" & lng����id, "����=" & lng����ID)
End Sub
Private Sub mnuViewToolbarStAnd_Click()
    Me.mnuViewToolbarStand.Checked = Not Me.mnuViewToolbarStand.Checked
    Me.clbThis.Visible = Me.mnuViewToolbarStand.Checked
    Form_Resize
End Sub

Private Sub mnuViewToolbarText_Click()
    Dim i As Integer
    Me.mnuViewToolbarText.Checked = Not Me.mnuViewToolbarText.Checked
    If Me.mnuViewToolbarText.Checked Then
        For i = 1 To Me.tlbThis.Buttons.count
            Me.tlbThis.Buttons(i).Caption = Me.tlbThis.Buttons(i).Tag
        Next
    Else
        For i = 1 To Me.tlbThis.Buttons.count
            Me.tlbThis.Buttons(i).Caption = ""
        Next
    End If
    Me.clbThis.Bands(1).MinHeight = Me.tlbThis.Height
    Me.clbThis.Refresh
    Form_Resize
End Sub

Private Sub picHBar_S_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        picHBar_S.BackColor = &H8000000C
    End If
End Sub

Private Sub picHBar_S_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Me.picHBar_S.Top = Me.picHBar_S.Top + Y
    End If
End Sub

Private Sub picHBar_S_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Call Form_Resize
        picHBar_S.BackColor = &H8000000A
    End If
End Sub


Private Sub picUp_S_Paint()
    '����:��ӡʱ�䷶Χ
    picUp_S.Cls
    picUp_S.CurrentX = 90
    picUp_S.CurrentY = 60
    picUp_S.Print "���Ϲ�񼰼۸���Ϣ"
End Sub

Private Sub picVBar_S_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        picVBar_S.BackColor = &H8000000C
    End If
End Sub

Private Sub picVBar_S_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Me.picVBar_S.Left = Me.picVBar_S.Left + X
    End If
End Sub

Private Sub picVBar_S_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Call Form_Resize
        picVBar_S.BackColor = &H8000000A
    End If
End Sub

Private Sub tlbThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim lvwTemp As ListView
    
    Select Case Button.Key
    Case "Preview"
        Call mnuFilePreView_Click
    Case "Print"
        Call mnuFilePrint_Click
    Case "Class"
        Call PopupMenu(Me.mnuClass, 2)
    Case "����"
        If Me.ActiveControl Is lvwItems Then
            mnuEditAddName_Click
        Else
            mnuEditItemAdd_Click
        End If
    Case "�޸�"
        If Me.ActiveControl Is lvwItems Then
            If mnuEditModifyName.Visible Then
                mnuEditModifyName_Click
            End If
        Else
            mnuEditItemMod_Click
        End If
    Case "ɾ��"
        If Me.ActiveControl Is lvwItems Then
            mnuEditDeleName_Click
        Else
            mnuEditItemDel_Click
        End If
    Case "Start"
        Call mnuEditStart_Click
    Case "Stop"
        Call mnuEditStop_Click
    Case "Find"
        Call mnuViewFind_Click
    Case "Filter"
        Call mnuFilter_Click
    Case "View"
        Set lvwTemp = lvwItems
        mnuViewIcon(lvwTemp.View).Checked = False
        If lvwTemp.View = 3 Then
            mnuViewIcon(0).Checked = True
            lvwTemp.View = 0
        Else
            mnuViewIcon(lvwTemp.View + 1).Checked = True
            lvwTemp.View = lvwTemp.View + 1
        End If
        
    Case "Help"
        Call mnuHelpTitle_Click
    Case "Exit"
        Call mnufileexit_Click
    End Select
End Sub

Private Sub tlbThis_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Dim i As Integer
    Select Case ButtonMenu.Key
    Case "����Ʒ��"
        Call mnuEditAddName_Click
    Case "���ӹ��"
        Call mnuEditItemAdd_Click
    Case Else
        For i = 0 To 3
            mnuViewIcon(i).Checked = False
        Next
        mnuViewIcon(ButtonMenu.Index - 1).Checked = True
        lvwItems.View = ButtonMenu.Index - 1
    End Select
End Sub

Private Sub tlbThis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 2 Then Exit Sub
    PopupMenu Me.mnuViewToolbar, 2
End Sub

Private Sub tvwClass_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 2 Then Exit Sub
    If InStr(1, mstrPrivs, ";���ķ������;") = 0 Then Exit Sub
    Call PopupMenu(Me.mnuClass, 2)
End Sub

Private Sub tvwClass_NodeClick(ByVal Node As MSComctlLib.Node)
    If Me.lvwItems.Tag = Node.Key Then Exit Sub
    Me.lvwItems.Tag = Node.Key
    Call zlRefRecords
    
    Call SetMenuEnable

End Sub
Private Sub InitLvwStuffColHead()
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ʼ��ͷ
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    With Me.lvwItems.ColumnHeaders
        .Clear
        .Add , "_����", "����", 2500
        .Add , "_����", "����", 1000
        .Add , "_Ӣ������", "Ӣ������", 1000
        .Add , "_ɢװ��λ", "ɢװ��λ", 1000
        .Add , "_�����Ա�", "�����Ա�", 1200
        .Add , "_����ʱ��", "����ʱ��", 1200
        .Add , "_����ʱ��", "����ʱ��", 1200
    End With
    
    With Me.lvwItems
        .ListItems.Clear
        .ColumnHeaders("_����").Position = 1
        .SortKey = .ColumnHeaders("_����").Index - 1
        .SortOrder = lvwAscending
    End With
    Call RestoreListViewState(Me.lvwItems, Me.Name, Me.lvwItems.View)
    

End Sub

Private Sub RefClassesAppend()
'������̬��¼��
    On Error GoTo ErrHand
    Set mrsRefClasses = New ADODB.Recordset
    If mrsRefClasses.State <> 1 Then
        With mrsRefClasses
            Call .Fields.Append("ID", adDouble, 20, adFldIsNullable)
            Call .Fields.Append("�ϼ�id", adDouble, 20, adFldIsNullable)
            Call .Fields.Append("����", adLongVarChar, 20, adFldIsNullable)
            Call .Fields.Append("����", adLongVarChar, 100, adFldIsNullable)
            Call .Fields.Append("����", adLongVarChar, 100, adFldIsNullable)
            
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .Open '�򿪼�¼��
        End With
    End If
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub RefRecordsAppend()
'������̬��¼��
    On Error GoTo ErrHand
    Set mrsRefRecords = New ADODB.Recordset
    If mrsRefRecords.State <> 1 Then
        With mrsRefRecords
            Call .Fields.Append("����id", adDouble, 20, adFldIsNullable)
            Call .Fields.Append("����id", adDouble, 20, adFldIsNullable)
            Call .Fields.Append("����", adLongVarChar, 50, adFldIsNullable)
            Call .Fields.Append("����", adLongVarChar, 100, adFldIsNullable)
            Call .Fields.Append("Ӣ������", adLongVarChar, 100, adFldIsNullable)
            Call .Fields.Append("ɢװ��λ", adLongVarChar, 50, adFldIsNullable)
            Call .Fields.Append("����ʱ��", adDate, , adFldIsNullable)
            Call .Fields.Append("����ʱ��", adDate, , adFldIsNullable)
            Call .Fields.Append("�����Ա�", adDouble, 10, adFldIsNullable)

            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .Open '�򿪼�¼��
        End With
    End If
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub zlRefClasses()
    '---------------------------------------------
    '��д���Ʒ�����Ŀ
    '---------------------------------------------
    Dim lngNode As Long
    Dim arrExecute As Variant
    Dim strID As String
    Dim i As Integer
    Dim str��¼ID�� As String
    
    err = 0
    On Error GoTo ErrHand:
    
    If Val(tvwClass.Tag) = 0 Then
        gstrSQL = "" & _
            "   select ID,�ϼ�ID,����,����,����" & _
            "   From ���Ʒ���Ŀ¼" & _
            "   Where ���� = 7 " & _
            "   start with �ϼ�ID is null Connect by prior ID=�ϼ�ID"
            
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    Else
        If mstrMaterialID = "" Then
            tvwClass.Nodes.Clear
            Me.lvwItems.ListItems.Clear
            Exit Sub
        End If
        
        gstrSQL = "" & _
            "   select /*+ Rule*/ Distinct A.ID,A.�ϼ�ID,A.����,A.����,A.����" & _
            "   From ���Ʒ���Ŀ¼ A,������ĿĿ¼ B, Table(Cast(f_Num2List([1]) As zlTools.t_NumList)) C " & _
            "   Where A.���� = 7 " & _
            "  And A.id=B.����id And B.id=C.Column_Value "
            
        'mstrMaterialID�����ܳ���4K���ֽ��ֱ�ִ��SQL�ٻ������ݼ�
        Call RefClassesAppend
        arrExecute = GetArrayByStr(mstrMaterialID, 4000, ",")
        For i = 0 To UBound(arrExecute)
            strID = arrExecute(i)
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strID)
            
            Do While Not rsTemp.EOF
                If InStr(str��¼ID�� & ",", "," & rsTemp!Id & ",") = 0 Then
                    With mrsRefClasses
                        .AddNew
                        .Fields!Id = rsTemp!Id
                        .Fields!�ϼ�id = rsTemp!�ϼ�id
                        .Fields!���� = rsTemp!����
                        .Fields!���� = rsTemp!����
                        .Fields!���� = rsTemp!����
                        .Update
                    End With
                    str��¼ID�� = str��¼ID�� & "," & rsTemp!Id
                End If
                
                rsTemp.MoveNext
            Loop
        Next
        
        Set rsTemp = Nothing
        Set rsTemp = mrsRefClasses.Clone
    End If
    
    If Not tvwClass.SelectedItem Is Nothing Then
        lngNode = Val(Mid(tvwClass.SelectedItem.Key, 2))
    End If
    
    With rsTemp
        tvwClass.Nodes.Clear
        If Val(tvwClass.Tag) = 1 Then Set objNode = Me.tvwClass.Nodes.Add(, , "_ALL", "���й��˽��", "close")
        Do While Not .EOF
            If Val(tvwClass.Tag) = 0 Then
                If IsNull(!�ϼ�id) Then
                    Set objNode = Me.tvwClass.Nodes.Add(, , "_" & !Id, "[" & !���� & "]" & !����, "close", "expend")
                Else
                    Set objNode = Me.tvwClass.Nodes.Add("_" & !�ϼ�id, tvwChild, "_" & !Id, "[" & !���� & "]" & !����, "close", "expend")
                End If
            Else
                Set objNode = Me.tvwClass.Nodes.Add("_ALL", tvwChild, "_" & !Id, "[" & !���� & "]" & !����, "close", "expend")
            End If
            
            objNode.Sorted = True
            objNode.Tag = IIf(IsNull(!����), "", !����)
         '   objNode.ExpandedImage = "expend"
            .MoveNext
        Loop
    End With
    
    err = 0
    On Error Resume Next
    If Me.tvwClass.Nodes.count > 0 Then
        If lngNode <> 0 Then
            Me.tvwClass.Nodes("_" & lngNode).Selected = True
            Me.tvwClass.Nodes("_" & lngNode).Expanded = True
            If err <> 0 Then
                Me.tvwClass.Nodes(1).Selected = True
                Me.tvwClass.Nodes(1).Expanded = True
            End If
        Else
            Me.tvwClass.Nodes(1).Selected = True
             Me.tvwClass.Nodes(1).Expanded = True
        End If
        Call zlRefRecords
    Else
        Me.lvwItems.ListItems.Clear
    End If
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub zlRefRecords(Optional lngLocale����ID As Long)
    '---------------------------------------------
    '��д���������б�
    '---------------------------------------------
    Dim lngId As Long
    Dim lngPreItemID As Long
    Dim arrExecute As Variant
    Dim strID As String
    Dim i As Integer
    
    err = 0
    On Error GoTo ErrHand
    If Me.lvwItems.SelectedItem Is Nothing Then
        lngPreItemID = 0
    Else
        lngPreItemID = Val(Mid(Me.lvwItems.SelectedItem.Key, 2))
    End If
    If tvwClass.SelectedItem Is Nothing Then
        lngId = 0
    Else
        lngId = Val(Mid(tvwClass.SelectedItem.Key, 2))
    End If
    
    If Val(tvwClass.Tag) = 0 Then
        gstrSQL = " " & _
         " SELECT distinct a.id as ����ID,A.����ID, a.����,a.����,c.���� as Ӣ������,a.���㵥λ as ɢװ��λ," & _
         "      to_char(a.����ʱ��,'yyyy-mm-dd') as ����ʱ��,nvl(a.����ʱ��,to_date('3000-01-01','YYYY-MM-DD')) as ����ʱ��,Nvl(A.�����Ա�,0) As �����Ա� " & _
         " FROM ������ĿĿ¼ a ,������Ŀ���� c" & _
         " WHERE a.id=c.������Ŀid(+) and c.����(+)=2 " & _
                IIf(mbln��ʾ�¼�, " And a.����id in (Select ID From ���Ʒ���Ŀ¼ where ����=7 start with id=[1] connect by prior id=�ϼ�id  )", "  and  A.����id=[1]")
    Else
        gstrSQL = " " & _
         " SELECT /*+ Rule*/ distinct a.id as ����ID,A.����ID, a.����,a.����,c.���� as Ӣ������,a.���㵥λ as ɢװ��λ," & _
         "      to_char(a.����ʱ��,'yyyy-mm-dd') as ����ʱ��,nvl(a.����ʱ��,to_date('3000-01-01','YYYY-MM-DD')) as ����ʱ��,Nvl(A.�����Ա�,0) As �����Ա� " & _
         " FROM ������ĿĿ¼ a ,������Ŀ���� c, Table(cast(f_Num2List([2]) as zlTools.t_NumList)) B " & _
         " WHERE a.id=c.������Ŀid(+) and c.����(+)=2 And a.ID=b.Column_Value "
         
         If lngId = 0 Then
            gstrSQL = gstrSQL & IIf(mbln��ʾ�¼�, "", " And a.����id =[1]")
         Else
            gstrSQL = gstrSQL & " And a.����id =[1]"
         End If
    End If
    
    If Me.mnuViewStoped.Checked = False Then
        gstrSQL = gstrSQL & " and (a.����ʱ�� is null or a.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))"
    End If
    gstrSQL = gstrSQL & " order by a.����"

    If Val(tvwClass.Tag) = 0 Then
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngId)
    Else
        'mstrMaterialID�����ܳ���4K���ֽ��ֱ�ִ��SQL�ٻ������ݼ�
        Call RefRecordsAppend
        arrExecute = GetArrayByStr(mstrMaterialID, 4000, ",")
        For i = 0 To UBound(arrExecute)
            strID = arrExecute(i)
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngId, strID)
               
            Do While Not rsTemp.EOF
                    With mrsRefRecords
                        .AddNew
                        .Fields!����id = rsTemp!����id
                        .Fields!����id = rsTemp!����id
                        .Fields!���� = rsTemp!����
                        .Fields!���� = rsTemp!����
                        .Fields!Ӣ������ = rsTemp!Ӣ������
                        .Fields!ɢװ��λ = rsTemp!ɢװ��λ
                        .Fields!����ʱ�� = rsTemp!����ʱ��
                        .Fields!����ʱ�� = rsTemp!����ʱ��
                        .Fields!�����Ա� = rsTemp!�����Ա�
                        .Update
                    End With
                rsTemp.MoveNext
            Loop
        Next
    
        Set rsTemp = Nothing
        Set rsTemp = mrsRefRecords.Clone
    End If
    
    With rsTemp
            Me.lvwItems.ListItems.Clear
            Do While Not .EOF
                Set objItem = Me.lvwItems.ListItems.Add(, "_" & !����id, zlStr.Nvl(!����))
                If Format(!����ʱ��, "YYYY-MM-DD") = "3000-01-01" Then
                    objItem.Icon = "start": objItem.SmallIcon = "start"
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_����ʱ��").Index - 1) = ""
                Else
                    objItem.Icon = "stop": objItem.SmallIcon = "stop"
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_����ʱ��").Index - 1) = Format(!����ʱ��, "yyyy-mm-dd")
                End If
                
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_����").Index - 1) = !����
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_Ӣ������").Index - 1) = zlStr.Nvl(!Ӣ������)
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_ɢװ��λ").Index - 1) = zlStr.Nvl(!ɢװ��λ)
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_�����Ա�").Index - 1) = IIf(!�����Ա� = 1, "����", IIf(!�����Ա� = 2, "Ů��", "���Ա�����"))
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_����ʱ��").Index - 1) = zlStr.Nvl(!����ʱ��)
                objItem.Tag = zlStr.Nvl(!����id)
                If !����id = lngLocale����ID Then
                    objItem.Selected = True
                End If
                
                If Format(!����ʱ��, "YYYY-MM-DD") <> "3000-01-01" Then
                    objItem.ForeColor = &HFF&
                    For intCount = 1 To Me.lvwItems.ColumnHeaders.count - 1
                        objItem.ListSubItems(intCount).ForeColor = &HFF&
                    Next
                End If
                If Val(zlStr.Nvl(!����id)) = lngPreItemID Then
                    objItem.Selected = True
                    objItem.EnsureVisible
                    
                End If
                 .MoveNext
            Loop
    End With
    If Me.lvwItems.ListItems.count > 0 Then
    
        If Me.lvwItems.SelectedItem Is Nothing Then
            Me.lvwItems.ListItems(1).Selected = True
        End If
        Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
        
        err = 0: On Error Resume Next
        DoEvents: Me.lvwItems.SelectedItem.EnsureVisible
        Me.stbThis.Panels(2).Text = "�÷��๲��" & Me.lvwItems.ListItems.count & "������Ʒ��"
    Else
        
        For intCount = Me.lblComment.LBound To Me.lblComment.UBound
            Me.lblComment(intCount).Caption = ""
        Next
        Me.stbThis.Panels(2).Text = ""
        Call InitHgdPrivceHeadcol
    End If
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub zlRptPrint(ByVal bytMode As Byte)
    '-------------------------------------------------
    '����:��¼���ӡ
    '����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    '-------------------------------------------------
    Dim objPrint As New zlPrintLvw
    err = 0: On Error Resume Next
    Set objPrint.Body.objData = Me.lvwItems
    objPrint.Title.Text = "��������Ʒ���嵥"
    
    objPrint.UnderAppItems.Add "���ࣺ" & Me.tvwClass.SelectedItem.Text
    objPrint.BelowAppItems.Add "��ӡʱ�䣺" & Now
    
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrViewLvw objPrint, bytMode
    Else
        zlPrintOrViewLvw objPrint, bytMode
    End If
End Sub

Public Sub zlLocateItem(lng����id As Long, ByVal ln����ID As Long, lng����ID As Long)
    '---------------------------------------------
    '����:��λ��ָ���Ĳ�����ȥ���ڲ���ʱʹ��
    '����:���˺�
    '����:2007-05-28
    '---------------------------------------------
    Dim lngRow As Long
    Dim lstItem As ListItem, tvwNode As Node
    On Error GoTo ErrHand
    Set tvwNode = tvwClass.SelectedItem
    Set lstItem = lvwItems.SelectedItem
    
    Set Me.tvwClass.SelectedItem = Me.tvwClass.Nodes("_" & lng����id)
    Me.tvwClass.Nodes("_" & lng����id).Selected = True
    Me.tvwClass.SelectedItem.EnsureVisible
    Call zlRefRecords
    Set Me.lvwItems.SelectedItem = Me.lvwItems.ListItems("_" & ln����ID)
    Me.lvwItems.SelectedItem.EnsureVisible
    Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
    '��λ������ID��ȥ
    With vsStuff
        For lngRow = 1 To .Rows - 1
            If .RowData(lngRow) = lng����ID Then
                .Row = lngRow
                Call .ShowCell(.Row, 1)
            End If
        Next
    End With
    Exit Sub
ErrHand:
    Set tvwClass.SelectedItem = tvwNode
    Call zlRefRecords
    Set Me.lvwItems.SelectedItem = Me.lvwItems.ListItems(lstItem.Key)
    Me.lvwItems.SelectedItem.EnsureVisible
    Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

Private Sub picStuffLst_Resize()
    err = 0: On Error Resume Next
    With picStuffLst
        vsStuff.Top = 20
        vsStuff.Left = 20
        vsStuff.Width = .ScaleWidth - 40
        vsStuff.Height = .ScaleHeight - 40
    End With
End Sub
Private Sub picTabPrice_Resize()
    err = 0: On Error Resume Next
    With picTabPrice
        vsPrice.Top = 20
        vsPrice.Left = 20
        vsPrice.Width = .ScaleWidth - 40
        fraComment.Top = .ScaleHeight - fraComment.Height - 40
        fraComment.Left = vsPrice.Left
        vsPrice.Height = fraComment.Top - vsPrice.Top - 20
    End With
End Sub

Private Sub vsStuff_DblClick()
    Dim lng����ID As Long
    Dim lng����id As Long
    
    With vsStuff
        If .RowData(.Row) = 0 Then Exit Sub
        lng����ID = .RowData(.Row)
    End With
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    
    If tvwClass.SelectedItem Is Nothing Then
        lng����id = 0
    Else
        lng����id = Val(Mid(tvwClass.SelectedItem.Key, 2))
    End If
    Call frmStuffSpec.ShowEditCard(Me, g�鿴, Mid(Me.lvwItems.SelectedItem.Key, 2), lng����id, CStr(lng����ID), mstrPrivs)
    
End Sub

 

Private Sub vsStuff_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button <> 2 Then Exit Sub
   PopupMenu mnuEdit
End Sub

Private Sub vsStuff_RowColChange()
    Call SetMenuEnable
End Sub
Private Function LoadCostData(ByVal lng����ID As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:���ɱ��۵�����Ϣ
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-11-10 15:17:09
    '-----------------------------------------------------------------------------------------------------------
    Dim lngRow As Long
    
    
    'ԭ�ɱ��� NUMBER(16,7),Ч�� DATE,���� VARCHAR2(20),���� VARCHAR2(40)
    
    gstrSQL = "" & _
    " Select '' as NO, I.ID As ����ID, '[' || I.���� || ']' || I.���� || ' ' || I.��� || ' ' || A.���� As ����, " & _
    "        P.����,P.���� As �ⷿ, A.���� As ����, " & _
    "        I.���㵥λ As ��λ, S.��װ��λ,S.����ϵ��, " & _
    "        A.�³ɱ��� As �ɱ���,A.ԭ�ɱ���,'' as ִ������, 'δִ��' ժҪ " & _
    " From �ɱ��۵�����Ϣ A, �������� S, �շ���ĿĿ¼ I,���ű� P" & _
    " Where A.ҩƷID=S.����ID And A.ҩƷID=I.ID and A.�ⷿid=P.id" & _
    "       And S.����id = [1] And A.ִ������ is NULL " & _
            IIf(Me.mnuViewStoped.Checked = False, "and (I.����ʱ�� is null or I.����ʱ��>=to_date('3000-01-01','YYYY-MM-DD'))", "")
    
  gstrSQL = gstrSQL & " Union all " & vbCrLf & _
    " Select  B.NO as NO, I.ID As ����ID, '[' || I.���� || ']' || I.���� || ' ' || I.��� || ' ' || A.���� As ����, " & _
    "        P.����,P.���� As �ⷿ, A.���� , " & _
    "        I.���㵥λ As ��λ, S.��װ��λ,S.����ϵ��, " & _
    "        A.�³ɱ��� As �ɱ���,A.ԭ�ɱ���,to_char(A.ִ������,'yyyy-mm-dd') ִ������, B.ժҪ " & _
    " From �ɱ��۵�����Ϣ A,ҩƷ�շ���¼ B, �������� S, �շ���ĿĿ¼ I,���ű� P" & _
    " Where A.�շ�ID=B.ID and A.ҩƷID=S.����ID And A.ҩƷID=I.ID and A.�ⷿid=P.id" & _
    "       And S.����id = [1] And A.ִ������ is NOt NULL " & _
            IIf(Me.mnuViewStoped.Checked = False, "and (I.����ʱ�� is null or I.����ʱ��>=to_date('3000-01-01','YYYY-MM-DD'))", "")
    
    gstrSQL = gstrSQL & " Order By ����,����,ִ������ Desc, NO Desc "
    
    err = 0: On Error GoTo ErrHand:
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng����ID)
    
    With vsCost
        .Rows = 2
        .Clear 1
        .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
        If rsTemp.RecordCount = 0 Then
            LoadCostData = True: Exit Function
        End If
        
        .Rows = .FixedRows + rsTemp.RecordCount
        .Redraw = flexRDNone
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("NO")) = zlStr.Nvl(rsTemp!NO)
            .TextMatrix(lngRow, .ColIndex("������Ϣ")) = zlStr.Nvl(rsTemp!����)
            .Cell(flexcpData, lngRow, .ColIndex("������Ϣ")) = zlStr.Nvl(rsTemp!����ID)
            .TextMatrix(lngRow, .ColIndex("�ⷿ")) = zlStr.Nvl(rsTemp!�ⷿ)
            .TextMatrix(lngRow, .ColIndex("����")) = zlStr.Nvl(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("��λ")) = IIf(mintUnit = 0, zlStr.Nvl(rsTemp!��λ), zlStr.Nvl(rsTemp!��װ��λ))
            .Cell(flexcpData, lngRow, .ColIndex("��λ")) = IIf(mintUnit = 0, 1, Val(zlStr.Nvl(rsTemp!����ϵ��)))
            .TextMatrix(lngRow, .ColIndex("�ɱ���")) = Format(Val(zlStr.Nvl(rsTemp!�ɱ���)) * Val(.Cell(flexcpData, lngRow, .ColIndex("��λ"))), mFMT.FM_�ɱ���)
            .TextMatrix(lngRow, .ColIndex("ִ������")) = zlStr.Nvl(rsTemp!ִ������)
            .TextMatrix(lngRow, .ColIndex("˵��")) = zlStr.Nvl(rsTemp!ժҪ)
            
            .Cell(flexcpBackColor, lngRow, 0, lngRow, .Cols - 1) = IIf(zlStr.Nvl(rsTemp!ִ������) = "", RGB(225, 255, 255), RGB(240, 240, 240))
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        .Row = 1
        .Redraw = flexRDBuffered
    End With
    LoadCostData = True
    Exit Function
ErrHand:
    vsCost.Redraw = flexRDBuffered
    If ErrCenter = 1 Then
        Resume
    End If
End Function

 

Private Sub vsStuff_GotFocus()
    With vsStuff
        .SelectionMode = flexSelectionByRow
        .BackColorSel = &H8000000D
    End With
End Sub

Private Sub vsStuff_LostFocus()
    With vsStuff
        .BackColorSel = GRD_LOSTFOCUS_COLORSEL
    End With
End Sub

Private Sub vsPrice_GotFocus()
    With vsPrice
        .SelectionMode = flexSelectionByRow
        .BackColorSel = &H8000000D
    End With
End Sub

Private Sub vsPrice_LostFocus()
    With vsPrice
        .BackColorSel = GRD_LOSTFOCUS_COLORSEL
    End With
End Sub

Private Sub mnuEditPrice_Click()
    If checkNotExecutePrice(vsStuff.TextMatrix(vsStuff.Row, vsStuff.ColIndex("ID"))) Then Exit Sub
    Call frmStuffPriceCard.ShowMe(Me, 0, "", 0, 1, vsStuff.TextMatrix(vsStuff.Row, vsStuff.ColIndex("ID")))
    frmStuffMgr.SetFocus
End Sub

Private Function checkNotExecutePrice(Optional ByVal lngDrugID As Long = 0) As Boolean
    '���� ������Ƿ����δִ�еļ۸�
    Dim RecCheck As New ADODB.Recordset
    On Error GoTo ErrHandle
    
    checkNotExecutePrice = False
    '�ж��Ƿ���δִ�е���ʷ�۸�
    gstrSQL = " Select Count(*) Records From �շѼ�Ŀ Where �䶯ԭ��=0 " & GetPriceClassString("") & _
        " And ִ������ > Sysdate And �շ�ϸĿID=[1]"
    Set RecCheck = zlDatabase.OpenSQLRecord(gstrSQL, "", lngDrugID)

    With RecCheck
        If Not .EOF Then
            If Not IsNull(!Records) Then
                If !Records <> 0 Then
                    MsgBox "�ù�񻹴���δִ�е��ۼ۵��ۼ�¼��δִ�����Ĳ��ܵ��ۣ�", vbInformation, gstrSysName
                    checkNotExecutePrice = True
                    Exit Function
                End If
            End If
        End If
    End With

    '����Ƿ���δִ�еĳɱ��۵��ۼƻ�
    gstrSQL = "Select 1 From �ɱ��۵�����Ϣ Where ҩƷid = [1] And ִ������ Is Null And Rownum = 1 "
    Set RecCheck = zlDatabase.OpenSQLRecord(gstrSQL, "", lngDrugID)

    If RecCheck.RecordCount > 0 Then
        MsgBox "�ù�񻹴���δִ�еĳɱ��۵��ۼ�¼��δִ�����Ĳ��ܵ��ۣ�", vbInformation, gstrSysName
        checkNotExecutePrice = True
        Exit Function
    End If
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
