VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmChargeSortGrade 
   Caption         =   "�ѱ�ȼ�����"
   ClientHeight    =   6000
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10005
   Icon            =   "frmChargeSortGrade.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   10005
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab sstabItem 
      Height          =   3255
      Left            =   5040
      TabIndex        =   5
      Top             =   1080
      Width           =   4725
      _ExtentX        =   8334
      _ExtentY        =   5741
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "  ������Ŀ(&0)  "
      TabPicture(0)   =   "frmChargeSortGrade.frx":0582
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fgdDetail"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "  �շ���Ŀ(&1)  "
      TabPicture(1)   =   "frmChargeSortGrade.frx":059E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fgdItem"
      Tab(1).ControlCount=   1
      Begin VSFlex8Ctl.VSFlexGrid fgdDetail 
         Height          =   1935
         Left            =   360
         TabIndex        =   6
         Top             =   720
         Width           =   3735
         _cx             =   6588
         _cy             =   3413
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
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   16777215
         GridColor       =   -2147483631
         GridColorFixed  =   -2147483630
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
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
      Begin VSFlex8Ctl.VSFlexGrid fgdItem 
         Height          =   1935
         Left            =   -74640
         TabIndex        =   7
         Top             =   600
         Width           =   3735
         _cx             =   6588
         _cy             =   3413
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
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   16777215
         GridColor       =   -2147483631
         GridColorFixed  =   -2147483630
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
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
   End
   Begin VB.PictureBox picSplitV 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3225
      Left            =   4770
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3225
      ScaleWidth      =   45
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   960
      Width           =   45
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   10005
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "Toolbar1"
      MinHeight1      =   720
      Width1          =   615
      FixedBackground1=   0   'False
      Key1            =   "Comm"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   720
         Left            =   30
         TabIndex        =   2
         Top             =   30
         Width           =   9885
         _ExtentX        =   17436
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imgToolsStard"
         HotImageList    =   "imgToolsHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   11
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ԥ��"
               Key             =   "Preview"
               Object.ToolTipText     =   "��ӡԤ��"
               Object.Tag             =   "Ԥ��"
               ImageKey        =   "Preview"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "Print"
               Object.ToolTipText     =   "��ӡ"
               Object.Tag             =   "��ӡ"
               ImageKey        =   "Print"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "New"
               Object.ToolTipText     =   "���Ӵ�λ�ȼ�"
               Object.Tag             =   "����"
               ImageKey        =   "New"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�޸�"
               Key             =   "Modify"
               Object.ToolTipText     =   "�޸Ĵ�λ�ȼ�"
               Object.Tag             =   "�޸�"
               ImageKey        =   "Modify"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ɾ��"
               Key             =   "Delete"
               Object.ToolTipText     =   "ɾ����λ�ȼ�"
               Object.Tag             =   "ɾ��"
               ImageKey        =   "Delete"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split1"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�鿴"
               Key             =   "View"
               Object.ToolTipText     =   "�鿴��ʽ"
               Object.Tag             =   "�鿴"
               ImageKey        =   "View"
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
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Object.ToolTipText     =   "��������"
               Object.Tag             =   "����"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Exit"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imgToolsHot 
      Left            =   8280
      Top             =   690
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeSortGrade.frx":05BA
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeSortGrade.frx":07D4
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeSortGrade.frx":09EE
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeSortGrade.frx":0C08
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeSortGrade.frx":0E22
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeSortGrade.frx":103C
            Key             =   "View"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeSortGrade.frx":125C
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeSortGrade.frx":1476
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgToolsStard 
      Left            =   9000
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeSortGrade.frx":1690
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeSortGrade.frx":18AA
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeSortGrade.frx":1AC4
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeSortGrade.frx":1CDE
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeSortGrade.frx":1EF8
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeSortGrade.frx":2112
            Key             =   "View"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeSortGrade.frx":2332
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeSortGrade.frx":254C
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils32 
      Left            =   3090
      Top             =   5250
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeSortGrade.frx":2766
            Key             =   "Key"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeSortGrade.frx":2A80
            Key             =   "KeyD"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   2310
      Top             =   5310
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
            Picture         =   "frmChargeSortGrade.frx":2D9A
            Key             =   "Key"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeSortGrade.frx":2EF4
            Key             =   "KeyD"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwMain_S 
      Height          =   4755
      Left            =   30
      TabIndex        =   0
      Top             =   810
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   8387
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
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
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   5640
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   635
      SimpleText      =   $"frmChargeSortGrade.frx":304E
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmChargeSortGrade.frx":3095
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12568
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
      Begin VB.Menu mnusplit3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuEditAdd 
         Caption         =   "����(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "�޸�(&M)"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "ɾ��(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditItem 
         Caption         =   "�����շѱ༭(&S)"
      End
      Begin VB.Menu mnuEditUnion 
         Caption         =   "ͳһʵ�ձ���(&U)"
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
      Begin VB.Menu mnuViewSplit3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewColumn 
         Caption         =   "ѡ����(&C)"
      End
      Begin VB.Menu mnuViewSplit4 
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
      Begin VB.Menu mnuHelpSpt1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)..."
      End
   End
   Begin VB.Menu mnuShort2 
      Caption         =   "��ݲ˵�"
      Visible         =   0   'False
      Begin VB.Menu mnuShortMenu2 
         Caption         =   "����(&A)"
         Index           =   1
      End
      Begin VB.Menu mnuShortMenu2 
         Caption         =   "�޸�(&M)"
         Index           =   2
      End
      Begin VB.Menu mnuShortMenu2 
         Caption         =   "ɾ��(&D)"
         Index           =   3
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
Attribute VB_Name = "frmChargeSortGrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mrsDetail As New ADODB.Recordset
Dim mrsItem As New ADODB.Recordset
Dim msngStartX As Single
Dim mblnLoad As Boolean
Dim mintColumn As Integer
Dim mblnItem As Boolean
Private Const mstrLvw As String = "����,1200,0,1;����,600,0,2;����,600,0,0;��Ч�ڿ�ʼʱ��,1500,0,0;��Ч�ڽ���ʱ��,1500,0,0;���ÿ���,900,0,0;����,1000,0,0;���޳���,900,0,0;˵��,2000,0,0"
Private mlngMode As Long
Private mstrPrivs As String                              'Ȩ�޴�
Private mstrCharge As String
Private Sub SetItemMenu()
    Dim blnEnabled As Boolean
    If Me.ActiveControl Is fgdItem Then
        blnEnabled = (fgdItem.Rows > 1 And fgdItem.TextMatrix(1, 0) <> "")
        Toolbar1.Buttons("New").Enabled = True
        Toolbar1.Buttons("Modify").Enabled = blnEnabled
        Toolbar1.Buttons("Delete").Enabled = blnEnabled
        mnuEditAdd.Enabled = True
        mnuEditDelete.Enabled = blnEnabled
        mnuEditModify.Enabled = blnEnabled
        
        mnuEditItem.Visible = False
        mnuEditUnion.Visible = False
        mnuEditSplit.Visible = False
    End If
    
    If Me.ActiveControl Is fgdDetail Then
        Toolbar1.Buttons("New").Enabled = False
        Toolbar1.Buttons("Modify").Enabled = False
        Toolbar1.Buttons("Delete").Enabled = False
        mnuEditAdd.Enabled = False
        mnuEditDelete.Enabled = False
        mnuEditModify.Enabled = False
        
        mnuEditItem.Visible = True
        mnuEditUnion.Visible = True
        mnuEditSplit.Visible = True
    End If
End Sub

Private Sub fgdDetail_DblClick()
    If mnuEditModify.Visible = False Then Exit Sub
    If mnuEditItem.Enabled = False Then Exit Sub
    mnuEditItem_Click
End Sub


Private Sub fgdDetail_EnterCell()
    Call SetItemMenu
End Sub


Private Sub fgdDetail_LostFocus()
    Call SetMenu
End Sub

Private Sub fgdItem_DblClick()
    If mnuEditModify.Visible = False Then Exit Sub
    If mnuEditItem.Enabled = False Then Exit Sub
    mnuEditItem_Click
End Sub


Private Sub fgdItem_EnterCell()
    Call SetItemMenu
End Sub


Private Sub fgdItem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If fgdItem.Rows > 1 Then
            PopupMenu mnuEdit, 2
        End If
    End If
End Sub


Private Sub Form_Activate()
    Dim rsItem As New ADODB.Recordset
    
    Call Form_Resize 'Ϊ����ȷ����coolbar�ĸ߶�
    
    On Error GoTo ErrHandle
    If mblnLoad = True Then
        gstrSQL = "select ID,����,���� from ������Ŀ where ĩ��=1 and rownum<2"
        Set rsItem = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)

        If rsItem.RecordCount = 0 Then
            MsgBox "û�ҵ�������Ŀ���������зѱ�ȼ�����" & vbCrLf & "���ڡ�������Ŀ����������������Ŀ��", vbExclamation, "��ʾ"
            Unload Me
            Exit Sub
        End If
    
        Call FillList
    End If
    mblnLoad = False
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Dim Item As ListItem
    
    'ByZT20030722
    If glngSys Like "8??" Then
        Caption = "��Ա�ȼ�����"
    End If
    
    mlngMode = glngModul
    mstrPrivs = gstrPrivs
    
    Call Ȩ�޿���
    
    '���������ɾ����ListView�������
    lvwMain_S.Tag = "�ɱ仯��"
    '-----------
    RestoreWinState Me, App.ProductName
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngMode, mstrPrivs)
    
    '���ListView�Ļ�δ�����ã������һ��ʹ�ã��Ǿ͵���ȱʡ�ĳ�ʼ��
    If lvwMain_S.ColumnHeaders.Count = 0 Then
        zlControl.LvwSelectColumns lvwMain_S, mstrLvw, True
    End If
    '����LvwMain��ʾ���ö�Ӧ�˵�
     mnuViewIcon_Click lvwMain_S.View
    mblnLoad = True
    
    '���в��ֳ�ʼ��
    With fgdDetail
        .Cols = 5
        .ColWidth(0) = 0
        .ColWidth(1) = 1700
        .ColWidth(2) = 2200
        .ColWidth(3) = 1050
        .ColWidth(4) = 2000
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .TextMatrix(0, 1) = "������Ŀ"
        .TextMatrix(0, 2) = "Ӧ�ս��(Ԫ)"
        .TextMatrix(0, 3) = "ʵ�ձ���(%)"
        .TextMatrix(0, 4) = "���㷽��"
    End With
    
    With fgdItem
        .Cols = 5
        .ColWidth(0) = 0
        .ColWidth(1) = 2500
        .ColWidth(2) = 3000
        .ColWidth(3) = 1050
        .ColWidth(4) = 2000
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .TextMatrix(0, 1) = "�շ���Ŀ"
        .TextMatrix(0, 2) = "Ӧ�ս��(Ԫ)"
        .TextMatrix(0, 3) = "ʵ�ձ���(%)"
        .TextMatrix(0, 4) = "���㷽��"
    End With
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
    SizeControls
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Sub lvwMain_S_DblClick()
    If mblnItem Then
        If mnuEditModify.Enabled And mnuEditModify.Visible Then mnuEditModify_Click
    End If
End Sub

Private Sub lvwMain_S_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If Item.Tag <> "" Then stbThis.Panels(2).Text = "˵����" & Item.Tag
    mblnItem = True
    mstrCharge = lvwMain_S.SelectedItem.Text
    If sstabItem.Tab = 0 Then
        Call FillDetail
    Else
        Call FillItem
    End If
End Sub
Private Sub lvwMain_S_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If mnuEditModify.Enabled And mnuEditModify.Visible Then mnuEditModify_Click
    End If
End Sub

Private Sub lvwMain_S_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnItem = False
End Sub

Private Sub lvwMain_S_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    If Button = 2 Then
        mnuShortMenu2(2).Enabled = mnuEditModify.Enabled
        mnuShortMenu2(3).Enabled = mnuEditDelete.Enabled
        For i = 0 To 3
            mnuShortIcon(i).Checked = mnuViewIcon(i).Checked
        Next
        PopupMenu mnuShort2, vbPopupMenuRightButton
    End If
End Sub

Private Sub lvwMain_S_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If mintColumn = ColumnHeader.Index - 1 Then '���Ǹղ�����
        lvwMain_S.SortOrder = IIF(lvwMain_S.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintColumn = ColumnHeader.Index - 1
        lvwMain_S.SortKey = mintColumn
        lvwMain_S.SortOrder = lvwAscending
    End If
End Sub

Private Sub mnuEditAdd_Click()
    If sstabItem.Tab = 0 Then
        Call frmChargeSortEdit.�༭�ѱ�("")
    Else
        If frmChargeSortItemEdit.ShowMe(Me, 1, lvwMain_S.SelectedItem.Text, 0, "") = True Then
            Call FillItem
        End If
    End If
End Sub

Private Sub mnuEditDelete_Click()
    On Error GoTo ErrHandle
    If sstabItem.Tab = 0 Then
        If lvwMain_S.ListItems.Count = 0 Then Exit Sub
        If Not lvwMain_S.SelectedItem.Selected Then Exit Sub
        If MsgBox("�Ƿ�ɾ���ѱ�" & lvwMain_S.SelectedItem.Text, vbQuestion Or vbDefaultButton2 Or vbYesNo, gstrSysName) = vbNo Then Exit Sub
        Err = 0
        On Error Resume Next
        With lvwMain_S.SelectedItem
            gstrSQL = "zl_�ѱ�_delete('" & Mid(.Key, 2) & "')"
        End With
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        
        If Err <> 0 Then
            MsgBox "ɾ��ʧ�ܣ����ܸ÷ѱ��Ѿ�ʹ��", vbExclamation, gstrSysName
            Err.Clear
            Exit Sub
        End If
        
        Dim intIndex As Integer
        With lvwMain_S
            intIndex = .SelectedItem.Index
            .ListItems.Remove .SelectedItem.Key
            If .ListItems.Count > 0 Then
                intIndex = IIF(.ListItems.Count > intIndex, intIndex, .ListItems.Count)
                .ListItems(intIndex).Selected = True
                .ListItems(intIndex).EnsureVisible
            End If
        End With
        
        If sstabItem.Tab = 0 Then
            Call FillDetail
        Else
            Call FillItem
        End If
    Else
        If MsgBox("�Ƿ�ɾ���շ���Ŀ��[" & fgdItem.TextMatrix(fgdItem.Row, 1) & "]�ķѱ����ã�", vbQuestion Or vbDefaultButton2 Or vbYesNo, gstrSysName) = vbNo Then Exit Sub
        gstrSQL = "zl_�ѱ���ϸ_update('" & lvwMain_S.SelectedItem.Text & "'," & Val(fgdItem.TextMatrix(fgdItem.Row, 0)) & ",Null,0,1,Null)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        
        Call FillItem
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub mnuEditItem_Click()
'    With mrsDetail
'        .Filter = "������Ŀid=" & fgdDetail.TextMatrix(fgdDetail.Row, 0)
'        If .RecordCount <> 0 Then
'            frmChargeSortItemEdit.mstrGrade = .Fields("�ѱ�").Value
'            frmChargeSortItemEdit.mlngItemId = .Fields("������Ŀid").Value
'            frmChargeSortItemEdit.txtStage.Text = .RecordCount
'            frmChargeSortItemEdit.UdStage.Value = .RecordCount
'            frmChargeSortItemEdit.cboMeasure.ListIndex = Val(.Fields("���㷽��").Value)     '����Click�¼�������ؿؼ�
'            frmChargeSortItemEdit.lblItem.Caption = fgdDetail.TextMatrix(fgdDetail.Row, 1) & "�ֶ�����"
'            Do While Not .EOF
'                frmChargeSortItemEdit.lblNo(.AbsolutePosition - 1).Visible = True
'                frmChargeSortItemEdit.lblNo(.AbsolutePosition - 1).Caption = .AbsolutePosition
'                frmChargeSortItemEdit.txtMoney(.AbsolutePosition - 1).Visible = True
'                frmChargeSortItemEdit.txtMoney(.AbsolutePosition - 1).Text = Format(.Fields("Ӧ�ն���ֵ").Value, "###########0.00;-##########0.00;0.00;0.00")
'                frmChargeSortItemEdit.txtTax(.AbsolutePosition - 1).Visible = True
'                frmChargeSortItemEdit.txtTax(.AbsolutePosition - 1).Text = Format(.Fields("ʵ�ձ���").Value, "###0.000;-##0.000;0.000;0.000")
'                .MoveNext
'            Loop
'            frmChargeSortItemEdit.mblnChange = False
'            frmChargeSortItemEdit.Show 1, Me
'            .Filter = adFilterNone
'            If frmChargeSortItemEdit.mblnOK = True Then Call FillDetail
'        Else
'            .Filter = adFilterNone
'        End If
'
'    End With
    If sstabItem.Tab = 0 Then
        If frmChargeSortItemEdit.ShowMe(Me, 0, lvwMain_S.SelectedItem.Text, Val(fgdDetail.TextMatrix(fgdDetail.Row, 0)), fgdDetail.TextMatrix(fgdDetail.Row, 1)) = True Then
            Call FillDetail
        End If
    Else
        If frmChargeSortItemEdit.ShowMe(Me, 1, lvwMain_S.SelectedItem.Text, Val(fgdItem.TextMatrix(fgdItem.Row, 0)), fgdItem.TextMatrix(fgdItem.Row, 1)) = True Then
            Call FillItem
        End If
    End If
    Call SetItemMenu
End Sub

Private Sub mnuEditModify_Click()
    If sstabItem.Tab = 0 Then
        If lvwMain_S.ListItems.Count = 0 Then Exit Sub
        If Not lvwMain_S.SelectedItem.Selected Then Exit Sub
        
        Call frmChargeSortEdit.�༭�ѱ�(lvwMain_S.SelectedItem.Text)
    Else
        If frmChargeSortItemEdit.ShowMe(Me, 1, lvwMain_S.SelectedItem.Text, Val(fgdItem.TextMatrix(fgdItem.Row, 0)), fgdItem.TextMatrix(fgdItem.Row, 1)) = True Then
            Call FillItem
        End If
    End If
End Sub

Private Sub mnuEditUnion_Click()
    If lvwMain_S.SelectedItem Is Nothing Then Exit Sub
    
    If frmChargeSortRate.UnifyPercentage(lvwMain_S.SelectedItem.Text, Val(fgdDetail.TextMatrix(fgdDetail.Row, 3))) = True Then
        Call FillDetail
    End If
End Sub

Private Sub mnuFileExcel_Click()
    subPrint 3
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnufilePreview_Click()
    subPrint 2
End Sub

Private Sub mnuFilePrint_Click()
    subPrint 1
End Sub

Private Sub mnuFilePrintset_Click()
    zlPrintSet
End Sub

Private Sub subPrint(ByVal intMode As Byte)
'����:���д�ӡ,Ԥ���������EXCEL
'����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    
    If lvwMain_S.ListItems.Count = 0 Then Exit Sub
    objPrint.Title = IIF(sstabItem.Tab = 0, "�ѱ��(������Ŀ)", "�ѱ��(�շ���Ŀ)")
    
    If sstabItem.Tab = 0 Then
        Set objPrint.Body = fgdDetail
    Else
        Set objPrint.Body = fgdItem
    End If
    
    objRow.Add ""
    objRow.Add "�ѱ�ȼ���" & lvwMain_S.SelectedItem.Text & "    "
    objPrint.UnderAppRows.Add objRow
    
    If intMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, intMode
    End If
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuhelpTitle_Click()
   ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hwnd)
End Sub


Private Sub mnuReportItem_Click(Index As Integer)
    'Ĭ�ϲ���������=�ѱ����
    Dim str���� As String
        
    If Not lvwMain_S.SelectedItem Is Nothing Then
        str���� = Mid(lvwMain_S.SelectedItem.Key, 2)
    End If
    
    Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
        "����=" & str����)
End Sub

Private Sub mnuShortMenu2_Click(Index As Integer)
    Select Case Index
        Case 1
            mnuEditAdd_Click
        Case 2
            mnuEditModify_Click
        Case 3
            mnuEditDelete_Click
    End Select
        
End Sub

Private Sub mnuShortIcon_Click(Index As Integer)
    mnuViewIcon_Click Index
End Sub

Private Sub mnuViewColumn_Click()
    If zlControl.LvwSelectColumns(lvwMain_S, mstrLvw) = True Then
        '���б仯��Ҫ����ˢ��
        mnuViewRefresh_Click
    End If
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    Me.cbrThis.Visible = mnuViewToolButton.Checked
    SizeControls
End Sub

Private Sub mnuViewIcon_Click(Index As Integer)
    Dim i As Integer
    For i = 0 To 3
        mnuViewIcon(i).Checked = False
        Toolbar1.Buttons("View").ButtonMenus(i + 1).Text = Replace(Toolbar1.Buttons("View").ButtonMenus(i + 1).Text, "��", "  ")
    Next
    mnuViewIcon(Index).Checked = True
    Toolbar1.Buttons("View").ButtonMenus(Index + 1).Text = Replace(Toolbar1.Buttons("View").ButtonMenus(Index + 1).Text, "  ", "��")
    lvwMain_S.View = Index
End Sub

Private Sub mnuViewRefresh_Click()
    Call FillList
End Sub

Private Sub mnuViewStatus_Click()
    Me.mnuViewStatus.Checked = Not Me.mnuViewStatus.Checked
    Me.stbThis.Visible = Me.mnuViewStatus.Checked
    SizeControls
End Sub

Private Sub mnuViewToolText_Click()
    Dim intCount As Integer
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    If mnuViewToolText.Checked Then
        For intCount = 1 To Me.Toolbar1.Buttons.Count
            Me.Toolbar1.Buttons(intCount).Caption = Me.Toolbar1.Buttons(intCount).Tag
        Next
    Else
        For intCount = 1 To Me.Toolbar1.Buttons.Count
            Me.Toolbar1.Buttons(intCount).Caption = ""
        Next
    End If
    Me.cbrThis.Bands(1).MinHeight = Me.Toolbar1.Height
    Me.cbrThis.Refresh
    SizeControls
End Sub

Private Sub picSplitV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        msngStartX = X
    End If
End Sub

Private Sub picSplitV_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sngTemp As Single
    If Button = 1 Then
        sngTemp = picSplitV.Left + X - msngStartX
        If sngTemp > 1000 And Me.ScaleWidth - (sngTemp + picSplitV.Width) > 1000 Then
            picSplitV.Left = sngTemp
            lvwMain_S.Width = picSplitV.Left - lvwMain_S.Left
            sstabItem.Left = picSplitV.Left + picSplitV.Width
            sstabItem.Width = Me.ScaleWidth - sstabItem.Left
            
            fgdDetail.Top = 480
            fgdDetail.Width = sstabItem.Width - 240
            fgdItem.Top = 480
            fgdItem.Width = sstabItem.Width - 240
        End If
        lvwMain_S.SetFocus
    End If
End Sub


Private Sub SizeControls()
'����:���ı䴰�ڴ�Сʱ,�Ը����ؼ���λ�ý�����������
    
    Dim sngTop As Single, sngBottom As Single
    
    On Error Resume Next
    
    sngTop = IIF(Me.cbrThis.Visible, Me.cbrThis.Height, 0)
    sngBottom = IIF(Me.stbThis.Visible, Me.stbThis.Height, 0)
    
    lvwMain_S.Top = sngTop
    picSplitV.Top = sngTop
    sstabItem.Top = sngTop
    
    fgdDetail.Top = 480
    fgdItem.Top = 480
    
    lvwMain_S.Height = Me.ScaleHeight - sngBottom - lvwMain_S.Top
    picSplitV.Height = Me.ScaleHeight - sngTop - sngBottom
    sstabItem.Height = Me.ScaleHeight - sngTop - sngBottom
    sstabItem.Width = Me.ScaleWidth - sstabItem.Left
    
    fgdDetail.Height = sstabItem.Height - fgdDetail.Top - 100
    fgdItem.Height = sstabItem.Height - fgdItem.Top - 100
    
    lvwMain_S.Left = Me.ScaleLeft
    picSplitV.Left = lvwMain_S.Left + lvwMain_S.Width
    sstabItem.Left = picSplitV.Left + picSplitV.Width
    
    fgdDetail.Left = 80
    fgdItem.Left = 80
    
    fgdDetail.Width = sstabItem.Width - 240
    fgdItem.Width = sstabItem.Width - 240

End Sub

Private Sub sstabItem_Click(PreviousTab As Integer)
    fgdDetail.Visible = True
    fgdItem.Visible = True
    
    If sstabItem.Tab = 0 Then
        fgdItem.Visible = False
        Call FillDetail
    Else
        fgdDetail.Visible = False
        Call FillItem
    End If
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Dim i As Integer
    For i = 0 To 3
        mnuViewIcon(i).Checked = False
    Next
    mnuViewIcon(ButtonMenu.Index - 1).Checked = True
    lvwMain_S.View = ButtonMenu.Index - 1
End Sub

Private Sub Toolbar1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuViewTool
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Preview"
            mnufilePreview_Click
        Case "Print"
            mnuFilePrint_Click
        Case "New"
            mnuEditAdd_Click
        Case "Delete"
            mnuEditDelete_Click
        Case "Modify"
            mnuEditModify_Click
        Case "Help"
            mnuhelpTitle_Click
        Case "Exit"
            mnuFileExit_Click
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

Public Sub FillList()
    Dim rsItem As New ADODB.Recordset
    Dim lst As ListItem
    Dim strKey As String, strIcon As String
    
    'ˢ��ListView
    On Error GoTo ErrHandle
    If Not lvwMain_S.SelectedItem Is Nothing Then
        '����ԭ�м�ֵ
        strKey = lvwMain_S.SelectedItem.Key
    End If
    With rsItem
        gstrSQL = "select ����,����,����,ȱʡ��־ as ȱʡ��,˵��" & _
                   ",to_char(��Ч��ʼ,'yyyy-MM-dd') as ��Ч�ڿ�ʼʱ��,to_char(��Ч����,'yyyy-MM-dd') as ��Ч�ڽ���ʱ��" & _
                   ",decode(���ÿ���,2,'ָ��','ȫ��') as ���ÿ���,decode(����,2,'��̬����Ŀ','���Ψһ��Ŀ') as ����,decode(���޳���,1,'��','��') ���޳��� from �ѱ�"
        Set rsItem = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)

        Dim lngCol  As Long
        Dim varValue As Variant
        lvwMain_S.ListItems.Clear
        Do Until rsItem.EOF
            strIcon = IIF(rsItem("ȱʡ��") = 1, "KeyD", "Key")
            
            Set lst = lvwMain_S.ListItems.Add(, "C" & rsItem("����"), rsItem("����"), strIcon, strIcon)
        
            '����ListView�����������ݿ�ȡ��
            For lngCol = 2 To lvwMain_S.ColumnHeaders.Count
                varValue = rsItem(lvwMain_S.ColumnHeaders(lngCol).Text).Value
                lst.SubItems(lngCol - 1) = IIF(IsNull(varValue), "", varValue)
            Next
            lst.Tag = IIF(IsNull(rsItem("˵��")), "", rsItem("˵��"))
            rsItem.MoveNext
        Loop
        If rsItem.RecordCount > 0 Then
            On Error Resume Next
            Set lst = lvwMain_S.ListItems(strKey)
            If Err <> 0 Then
                Err.Clear
                Set lst = lvwMain_S.ListItems(1)
                lst.Selected = True
                lst.EnsureVisible
            Else
                lst.Selected = True
                lst.EnsureVisible
            End If
        End If
    End With
    
    mstrCharge = lvwMain_S.SelectedItem.Text
    If sstabItem.Tab = 0 Then
        Call FillDetail
    Else
        Call FillItem
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FillDetail()
    Dim i As Integer
    If lvwMain_S.SelectedItem Is Nothing Then
        fgdDetail.Rows = 2
        fgdDetail.RowData(1) = 0
        For i = 0 To fgdDetail.Cols - 1
            fgdDetail.TextMatrix(1, i) = ""
        Next
        Call SetMenu
        Exit Sub
    End If
     
    On Error GoTo ErrHandle
'    gstrSQL = "zl_�ѱ�_NEW('" & mstrCharge & "')"
'    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    If lvwMain_S.ListItems.Count > 0 Then
        mstrCharge = lvwMain_S.SelectedItem.Text
    End If
    
    gstrSQL = "select a.�ѱ�,a.������ĿID,b.���� as ������Ŀ,a.�κ�,Ӧ�ն���ֵ,Ӧ�ն�βֵ,ʵ�ձ���,Decode(���㷽��,1,'1-�ɱ��ۼ��ձ�������','0-�ֶα�������') as ���㷽��" & _
            " from �ѱ���ϸ A,������Ŀ B" & _
            " Where a.������ĿID = B.id" & _
            "       and �ѱ�=[1] " & _
            " Order by b.����,Ӧ�ն���ֵ"
    Set mrsDetail = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mstrCharge)
        
    With fgdDetail
        .Clear
        .redraw = False
        .Rows = IIF(mrsDetail.RecordCount > 0, mrsDetail.RecordCount + 1, 2)
        .TextMatrix(0, 1) = "������Ŀ"
        .TextMatrix(0, 2) = "Ӧ�ս��(Ԫ)"
        .TextMatrix(0, 3) = "ʵ�ձ���(%)"
        .TextMatrix(0, 4) = "���㷽��"
        .MergeCol(1) = True
        .MergeCol(2) = False
        .MergeCol(3) = False
        .MergeCol(4) = False
        
        Do While Not mrsDetail.EOF()
            .RowData(mrsDetail.AbsolutePosition) = mrsDetail.Fields("�κ�").Value
            .TextMatrix(mrsDetail.AbsolutePosition, 0) = mrsDetail.Fields("������ĿID").Value
            .TextMatrix(mrsDetail.AbsolutePosition, 1) = mrsDetail.Fields("������Ŀ").Value
            .TextMatrix(mrsDetail.AbsolutePosition, 2) = Format(mrsDetail.Fields("Ӧ�ն���ֵ").Value, "##########0.00;-#########0.00;0.00;0.00") & _
                    " �� " & Format(mrsDetail.Fields("Ӧ�ն�βֵ").Value, "##########0.00;-#########0.00;0.00;0.00")
            .TextMatrix(mrsDetail.AbsolutePosition, 3) = Format(mrsDetail.Fields("ʵ�ձ���").Value, "###0.00;-##0.00;0.00;0.00")
            .TextMatrix(mrsDetail.AbsolutePosition, 4) = mrsDetail.Fields("���㷽��").Value
            .Row = mrsDetail.AbsolutePosition
            .Col = 3
            .CellBackColor = &H80000005
            .Col = 1
            mrsDetail.MoveNext
        Loop
        .Row = 1
        .redraw = True
    End With
    Call SetMenu
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FillItem()
    Dim i As Integer
    If lvwMain_S.SelectedItem Is Nothing Then
        fgdItem.Rows = 2
        fgdItem.RowData(1) = 0
        For i = 0 To fgdItem.Cols - 1
            fgdItem.TextMatrix(1, i) = ""
        Next
        Call SetMenu
        Exit Sub
    End If
     
    On Error GoTo ErrHandle
    gstrSQL = "select a.�ѱ�,a.�շ�ϸĿid,B.���� || Decode(B.���, '5', '(' || B.���� || ')', '6', '(' || B.���� || ')', '7', '(' || B.���� || ')',  '('||C.����||')') As �շ���Ŀ,a.�κ�,Ӧ�ն���ֵ,Ӧ�ն�βֵ,ʵ�ձ���,Decode(���㷽��,1,'1-�ɱ��ۼ��ձ�������','0-�ֶα�������') as ���㷽��" & _
            " from �ѱ���ϸ A,�շ���ĿĿ¼ B, �շ���Ŀ��� C " & _
            " Where a.�շ�ϸĿid = B.id And B.��� = C.���� " & _
            "       and �ѱ�=[1] " & _
            " Order by C.����, B.����, A.Ӧ�ն���ֵ "
    Set mrsItem = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mstrCharge)
        
    With fgdItem
        .Clear
        .redraw = False
        .Rows = IIF(mrsItem.RecordCount > 0, mrsItem.RecordCount + 1, 2)
        .TextMatrix(0, 1) = "�շ���Ŀ"
        .TextMatrix(0, 2) = "Ӧ�ս��(Ԫ)"
        .TextMatrix(0, 3) = "ʵ�ձ���(%)"
        .TextMatrix(0, 4) = "���㷽��"
        .MergeCol(1) = True
        .MergeCol(2) = False
        .MergeCol(3) = False
        .MergeCol(4) = False
        
        Do While Not mrsItem.EOF()
            .RowData(mrsItem.AbsolutePosition) = mrsItem.Fields("�κ�").Value
            .TextMatrix(mrsItem.AbsolutePosition, 0) = mrsItem.Fields("�շ�ϸĿID").Value
            .TextMatrix(mrsItem.AbsolutePosition, 1) = mrsItem.Fields("�շ���Ŀ").Value
            .TextMatrix(mrsItem.AbsolutePosition, 2) = Format(mrsItem.Fields("Ӧ�ն���ֵ").Value, "##########0.00;-#########0.00;0.00;0.00") & _
                    " �� " & Format(mrsItem.Fields("Ӧ�ն�βֵ").Value, "##########0.00;-#########0.00;0.00;0.00")
            .TextMatrix(mrsItem.AbsolutePosition, 3) = Format(mrsItem.Fields("ʵ�ձ���").Value, "###0.00;-##0.00;0.00;0.00")
            .TextMatrix(mrsItem.AbsolutePosition, 4) = mrsItem.Fields("���㷽��").Value
            .Row = mrsItem.AbsolutePosition
            .Col = 3
            .CellBackColor = &H80000005
            .Col = 1
            mrsItem.MoveNext
        Loop
        .Row = 1
        .redraw = True
    End With
    Call SetMenu
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub Ȩ�޿���()
'����:�����е��û�Ȩ�޲���,��ʹһЩ�˵����ť���ɼ�
    If InStr(mstrPrivs, "��ɾ��") = 0 Then
        mnuEdit.Visible = False
        mnuEditModify.Visible = False
        mnuShortMenu2(1).Visible = False
        mnuShortMenu2(2).Visible = False
        mnuShortMenu2(3).Visible = False
        mnuShortsplit1.Visible = -False
        Toolbar1.Buttons("Split").Visible = False
        Toolbar1.Buttons("New").Visible = False
        Toolbar1.Buttons("Modify").Visible = False
        Toolbar1.Buttons("Delete").Visible = False
    End If
End Sub

Private Sub SetMenu()
'����:���ô�ӡ��Ԥ����ť����Чֵ
    Dim blnEnabled As Boolean
    
    mnuEditItem.Visible = True
    mnuEditUnion.Visible = True
    mnuEditSplit.Visible = True
    
    If sstabItem.Tab = 0 Then
        blnEnabled = Not (lvwMain_S.SelectedItem Is Nothing)
        Toolbar1.Buttons("New").Enabled = True
        Toolbar1.Buttons("Modify").Enabled = blnEnabled
        Toolbar1.Buttons("Delete").Enabled = blnEnabled
        mnuEditAdd.Enabled = True
        mnuEditDelete.Enabled = blnEnabled
        mnuEditModify.Enabled = blnEnabled
        
        blnEnabled = lvwMain_S.ListItems.Count > 0
        Toolbar1.Buttons("Print").Enabled = blnEnabled
        Toolbar1.Buttons("Preview").Enabled = blnEnabled
        mnuFilePreview.Enabled = blnEnabled
        mnuFilePrint.Enabled = blnEnabled
        mnuFileExcel.Enabled = blnEnabled
    End If
    
    If sstabItem.Tab = 1 Then
        blnEnabled = (fgdItem.Rows > 1 And fgdItem.TextMatrix(1, 0) <> "")
        Toolbar1.Buttons("New").Enabled = True
        Toolbar1.Buttons("Modify").Enabled = blnEnabled
        Toolbar1.Buttons("Delete").Enabled = blnEnabled
        mnuEditAdd.Enabled = True
        mnuEditDelete.Enabled = blnEnabled
        mnuEditModify.Enabled = blnEnabled
        
        mnuEditItem.Visible = False
        mnuEditUnion.Visible = False
        mnuEditSplit.Visible = False
    End If
    
End Sub


Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

