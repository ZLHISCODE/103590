VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmKind 
   Caption         =   "�����������"
   ClientHeight    =   6720
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10455
   Icon            =   "frmKind.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6720
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   6360
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmKind.frx":1CFA
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13361
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
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
   Begin MSComctlLib.TreeView tvw 
      Height          =   1770
      Left            =   330
      TabIndex        =   4
      Top             =   945
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   3122
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ils16"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   1035
      Top             =   4695
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
            Picture         =   "frmKind.frx":258E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKind.frx":29E0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   1485
      Left            =   3135
      TabIndex        =   3
      Top             =   870
      Width           =   2430
      _ExtentX        =   4286
      _ExtentY        =   2619
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
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
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList ils32 
      Left            =   405
      Top             =   4695
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
            Picture         =   "frmKind.frx":2CFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKind.frx":314C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   705
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   1244
      BandCount       =   1
      _CBWidth        =   10455
      _CBHeight       =   705
      _Version        =   "6.7.9782"
      Child1          =   "tbrThis"
      MinWidth1       =   4995
      MinHeight1      =   645
      Width1          =   8370
      Key1            =   "only"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   645
         Left            =   30
         TabIndex        =   2
         Top             =   30
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   1138
         ButtonWidth     =   820
         ButtonHeight    =   1138
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ilsMenu"
         HotImageList    =   "ilsMenuHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   15
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ԥ��"
               Key             =   "Ԥ��"
               Object.ToolTipText     =   "Ԥ��"
               Object.Tag             =   "Ԥ��"
               ImageKey        =   "Preview"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "��ӡ"
               Object.ToolTipText     =   "��ӡ"
               Object.Tag             =   "��ӡ"
               ImageKey        =   "Print"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_1"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageKey        =   "Class"
               Style           =   5
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_2"
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageKey        =   "Add"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�޸�"
               Key             =   "�޸�"
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
               Key             =   "Split_3"
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��Ŀ"
               Key             =   "��Ŀ"
               Object.ToolTipText     =   "��Ŀ"
               Object.Tag             =   "��Ŀ"
               ImageKey        =   "Item"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_4"
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�б�"
               Key             =   "�б�"
               Object.ToolTipText     =   "�б�"
               Object.Tag             =   "�б�"
               ImageKey        =   "View"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Large"
                     Text            =   "��ͼ��(&G)"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Small"
                     Text            =   "Сͼ��(&M)"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "List"
                     Text            =   "�б�(&L)"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Detail"
                     Text            =   "��ϸ����(&D)"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_5"
               Style           =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ilsMenu 
      Left            =   8325
      Top             =   3210
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKind.frx":3466
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKind.frx":3686
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKind.frx":38A6
            Key             =   "Class"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKind.frx":3AC2
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKind.frx":3CDE
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKind.frx":3EF8
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKind.frx":4118
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKind.frx":4338
            Key             =   "View"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKind.frx":4558
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKind.frx":4778
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsMenuHot 
      Left            =   7515
      Top             =   3210
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKind.frx":4998
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKind.frx":4BB8
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKind.frx":4DD8
            Key             =   "Class"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKind.frx":4FF4
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKind.frx":5210
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKind.frx":5562
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKind.frx":5782
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKind.frx":59A2
            Key             =   "View"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKind.frx":5BC2
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKind.frx":5DE2
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid vsfPrint 
      Height          =   780
      Left            =   5940
      TabIndex        =   5
      Top             =   1305
      Visible         =   0   'False
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   1376
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   270
      MergeCells      =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VSFlex8Ctl.VSFlexGrid vsf 
      Height          =   1740
      Left            =   3135
      TabIndex        =   6
      Top             =   2535
      Width           =   3030
      _cx             =   5345
      _cy             =   3069
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
      BackColorSel    =   16772055
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   12698049
      GridColorFixed  =   -2147483632
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
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
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
   Begin VB.Image imgX_S 
      Height          =   45
      Left            =   4860
      MousePointer    =   7  'Size N S
      Top             =   2385
      Width           =   5115
   End
   Begin VB.Image imgY_S 
      Height          =   4395
      Left            =   2670
      MousePointer    =   9  'Size W E
      Top             =   840
      Width           =   45
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFilePrintSet 
         Caption         =   "��ӡ����(&S)"
      End
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrintView 
         Caption         =   "��ӡԤ��(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "��ӡ(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileOutExcel 
         Caption         =   "�����&Excel"
      End
      Begin VB.Menu mnuFile_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuEditClass 
         Caption         =   "���ͷ���(&C)"
         Begin VB.Menu mnuEditClassAdd 
            Caption         =   "���ӷ���(&A)"
         End
         Begin VB.Menu mnuEditClassModify 
            Caption         =   "�޸ķ���(&M)"
         End
         Begin VB.Menu mnuEditClassDelete 
            Caption         =   "ɾ������(&D)"
         End
      End
      Begin VB.Menu mnuEdit_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditAdd 
         Caption         =   "��������(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "�޸�����(&M)"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "ɾ������(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEdit_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelect 
         Caption         =   "�����Ŀ(&S)"
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
         Begin VB.Menu mnuViewToolText 
            Caption         =   "�ı���ǩ(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "״̬��(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuView_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "��ͼ��(&G)"
         Checked         =   -1  'True
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
      Begin VB.Menu mnuView_2 
         Caption         =   "-"
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
         Caption         =   "&Web�ϵ�����"
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
      Begin VB.Menu h1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)..."
      End
   End
End
Attribute VB_Name = "frmKind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'���������弶��������**************************************************************************************************
Private mblnStartUp As Boolean                          '����������־
Private mstrVsf As String                               '����б���
Private mstrKey As String                               '������ǰ��ѡ��
Private Const mstrLvw As String = "����,2400,0,1;����,900,0,0;����,900,0,0;�����۸�,1200,1,0;���۸�,1200,1,0;�ۿ�,900,1,0;����,900,0,0,;˵��,1500,0,0;��������,1500,0,0"
Private mlngLoop As Long
Private WithEvents mobjPopMenu As clsPopMenu                '�Զ��嵯���˵�����
Attribute mobjPopMenu.VB_VarHelpID = -1
Private mbytPopMenu As Byte

Private Enum mCol
    ��Ŀ���� = 0
    ��Ŀ���
    ��鲿λ
    �ɼ���ʽ
    ����걾
    �����۸�
    ���۸�
    �ۿ�
End Enum


'�������Զ�����̻���************************************************************************************************
Private Function InitLoad() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:��ʼ�����ݣ������ڴ����Load�¼�
    '------------------------------------------------------------------------------------------------------------------
    Dim strVsf As String
    
    On Error GoTo errHand
    
    mstrKey = ""
    lvw.Tag = "�ɱ仯��"
        
    If Val(GetSetting("ZLSOFT", "˽��ȫ��\" & gstrDBUser, "ʹ�ø��Ի����", "0")) = 1 Then
        'ʹ�ø��Ի�����
        mstrVsf = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "������", mstrVsf)
                        
    End If
    
    If lvw.ListItems.Count = 0 Then zlControl.LvwSelectColumns lvw, mstrLvw, True
                
    mstrVsf = "��Ŀ����,3000,1,1,1,;��Ŀ���,900,1,1,1,;��鲿λ,1200,1,1,1,;�ɼ���ʽ,1200,1,1,1,;����걾,900,1,1,1,;�����۸�,1080,7,1,1,;���۸�,1080,7,1,1,;�ۿ�,1080,7,1,1,"
    Call CreateVsf(vsf, mstrVsf)
    
    vsf.ColFormat(mCol.�����۸�) = "0.00"
    vsf.ColFormat(mCol.���۸�) = "0.00"
    vsf.ColFormat(mCol.�ۿ�) = "0.000"
    InitLoad = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub ApplyPrivilege(ByVal strPrivilege As String)
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ� Ӧ��Ȩ�޴���
    '������ strPrivilege                    Ȩ��
    '------------------------------------------------------------------------------------------------------------------
    
    '�������
    'strPrivilege = "����;��ɾ��"
    
    '�����С���ɾ�ġ�Ȩ��ʱ
    If InStr(strPrivilege, "��ɾ��") = 0 Then
        mnuEdit.Visible = False
    End If
    
    tbrThis.Buttons("����").Visible = mnuEdit.Visible
    
    tbrThis.Buttons("����").Visible = mnuEdit.Visible
    tbrThis.Buttons("�޸�").Visible = mnuEdit.Visible
    tbrThis.Buttons("ɾ��").Visible = mnuEdit.Visible
    
    tbrThis.Buttons("Split_2").Visible = mnuEdit.Visible
    tbrThis.Buttons("Split_3").Visible = mnuEdit.Visible
    
    tbrThis.Buttons("��Ŀ").Visible = mnuEdit.Visible
    tbrThis.Buttons("Split_4").Visible = mnuEdit.Visible
End Sub

Private Sub AdjustEnableState()
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ� ���������ܲ˵��Ŀ���״̬
    '------------------------------------------------------------------------------------------------------------------
    
    mnuFilePrint.Enabled = True
    mnuFilePrintView.Enabled = True
    mnuFileOutExcel.Enabled = True
    
    mnuEditClassAdd.Enabled = True
    mnuEditClassModify.Enabled = True
    mnuEditClassDelete.Enabled = True
    
    mnuEditClass.Enabled = True
    
    mnuEditAdd.Enabled = True
    mnuEditModify.Enabled = True
    mnuEditDelete.Enabled = True
    
    mnuEditSelect.Enabled = True
    
    If lvw.ListItems.Count = 0 Then
                
        mnuEditModify.Enabled = False
        mnuEditDelete.Enabled = False
        
        mnuEditSelect.Enabled = False
    End If
    
    If Val(vsf.RowData(1)) = 0 Then
        mnuFilePrint.Enabled = False
        mnuFilePrintView.Enabled = False
        mnuFileOutExcel.Enabled = False
    End If
    
    If tvw.SelectedItem.Key = "K0" Then
        mnuEditClassModify.Enabled = False
        mnuEditClassDelete.Enabled = False
    End If
    
    tbrThis.Buttons("Ԥ��").Enabled = mnuFilePrintView.Enabled
    tbrThis.Buttons("��ӡ").Enabled = mnuFilePrint.Enabled
    
    tbrThis.Buttons("����").Enabled = mnuEditClassModify.Enabled Or mnuEditClassAdd.Enabled
    tbrThis.Buttons("����").Enabled = mnuEditAdd.Enabled
    tbrThis.Buttons("�޸�").Enabled = mnuEditModify.Enabled
    tbrThis.Buttons("ɾ��").Enabled = mnuEditDelete.Enabled
    tbrThis.Buttons("��Ŀ").Enabled = mnuEditSelect.Enabled
    
End Sub

Private Sub RefreshStateInfo()
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ�ˢ��״̬����ʾ��Ϣ
    '------------------------------------------------------------------------------------------------------------------
    If lvw.SelectedItem Is Nothing Then
        stbThis.Panels(2).Text = "���� " & lvw.ListItems.Count & " ��������ͣ�"
    Else
        If vsf.Rows = 2 And vsf.RowData(1) = 0 Then
            stbThis.Panels(2).Text = "���� " & lvw.ListItems.Count & " ��������ͣ�"
        Else
            stbThis.Panels(2).Text = "���� " & lvw.ListItems.Count & " ��������ͣ���" & lvw.SelectedItem.Text & "������ " & vsf.Rows - 1 & " �������Ŀ��"
        End If
    End If
    
End Sub

Public Function GetItem(ByRef lngKey As Long, ByVal intFoot As Integer) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ����༭���ݴ�����ã��ӿں���
    '------------------------------------------------------------------------------------------------------------------
    Dim lngIndex As Long
    Dim objItem As ListItem
    
    On Error GoTo errHand
    
    Set objItem = lvw.ListItems("K" & lngKey)
    If Not (objItem Is Nothing) Then
        
        lngIndex = objItem.Index
        lngIndex = lngIndex + intFoot
        
        Set objItem = Nothing
        Set objItem = lvw.ListItems(lngIndex)
        
        If Not (objItem Is Nothing) Then lngKey = Val(Mid(objItem.Key, 2))
            
        GetItem = True
    Else
        GetItem = False
    End If
    
    Exit Function
    
errHand:
    
End Function

Public Function EditRefresh(ByVal strMenuItem As String, ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ����༭���ݴ�����ã��ӿں���
    '------------------------------------------------------------------------------------------------------------------
    Dim lngSvrKey As Long
    
    On Error GoTo errHand

    Select Case strMenuItem
    Case "������ͷ���"
        
        Call ClearData("������ͷ���;�������;�����Ŀ")
        
        Call RefreshData("������ͷ���")
        
        On Error Resume Next
        tvw.Nodes("K" & lngKey).Selected = True
        tvw.Nodes("K" & lngKey).EnsureVisible
        On Error GoTo 0
        
        Call RefreshData("�������")
        Call RefreshData("�����Ŀ")
        
        
    Case "�������"
    
        Call ClearData("�������;�����Ŀ")
        Call RefreshData("�������")
        
        '�ָ��������
        Call zlControl.LvwRestoreItem(lvw, "K" & lngKey)
            
        Call RefreshData("�����Ŀ")
                
    Case "�����Ŀ"
        If lvw.SelectedItem.Key = "K" & lngKey Then
            mstrKey = ""
            Call lvw_ItemClick(lvw.SelectedItem)
        End If
    End Select
    
    Call AdjustEnableState
    Call RefreshStateInfo
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function ClearData(Optional ByVal strMenuItem As String = "") As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:
    '------------------------------------------------------------------------------------------------------------------
    
    strMenuItem = ";" & strMenuItem & ";"
    
    If InStr(strMenuItem, ";������ͷ���;") > 0 Then
        tvw.Nodes.Clear
    End If
    
    If InStr(strMenuItem, ";�������;") > 0 Then
        lvw.ListItems.Clear
    End If
    
    If InStr(strMenuItem, ";�������;") > 0 Then
        lvw.ListItems.Clear
    End If
    If InStr(strMenuItem, ";�����Ŀ;") > 0 Then
        Call ResetVsf(vsf)
    End If
        
End Function

Private Function RefreshData(Optional ByVal strMenuItem As String = "") As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ�ˢ��/װ������
    '------------------------------------------------------------------------------------------------------------------
    Dim lngKey As Long
    Dim rs As New ADODB.Recordset
    Dim objNode As Node
    Dim rsPrice As New ADODB.Recordset
    
    On Error GoTo errHand
    
    Select Case strMenuItem
    Case "������ͷ���"
        
        gstrSQL = GetPublicSQL(SQL.������ͷ���)
        
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        
        If rs.BOF = False Then Call FillTreeData(tvw, rs)
        
    Case "�������"
        
        If Val(Mid(tvw.SelectedItem.Key, 2)) = 0 Then
            
            gstrSQL = "select a.��� AS ID," & _
                            "a.����," & _
                            "a.����," & _
                            "a.����," & _
                            "DECODE(a.���÷�Χ,0,'����',1,'����',2,'����') AS ����," & _
                            "Trim(To_Char(c.�����۸�,'99999999.00')) As �����۸�," & _
                            "Trim(To_Char(c.���۸�,'99999999.00')) As ���۸�," & _
                            "Trim(To_Char(Decode(c.�����۸�,Null,0,0,0,10*c.���۸�/c.�����۸�),'99999999.000')) As �ۿ�," & _
                            "a.˵��," & _
                            "b.���� AS ��������," & _
                            "1 as ͼ�� " & _
                    "from ������� a," & _
                         "������� b," & _
                         "(Select a.���,Sum(b.�ּ�*a.����) As �����۸�,Sum(b.�ּ�*a.����*Nvl(a.�ۿ�,1)) As ���۸� " & _
                         "From ������ͼƼ� a," & _
                              "�շѼ�Ŀ b " & _
                         "Where b.�շ�ϸĿid=a.�շ�ϸĿid " & _
                               "and b.ִ������<=SYSDATE " & _
                               "and (b.��ֹ���� IS NULL OR b.��ֹ����>SYSDATE) " & _
                         "group by a.��� " & _
                         ") c " & _
                    "where a.ĩ�� = 1 AND a.�ϼ���� = b.���(+) " & _
                          "and a.���=c.���(+)"

        Else
            
            gstrSQL = "select a.��� AS ID," & _
                            "a.����," & _
                            "a.����," & _
                            "a.����," & _
                            "DECODE(a.���÷�Χ,0,'����',1,'����',2,'����') AS ����," & _
                            "Trim(To_Char(c.�����۸�,'99999999.00')) As �����۸�," & _
                            "Trim(To_Char(c.���۸�,'99999999.00')) As ���۸�," & _
                            "Trim(To_Char(Decode(c.�����۸�,Null,0,0,0,10*c.���۸�/c.�����۸�),'99999999.00')) As �ۿ�," & _
                            "a.˵��," & _
                            "b.���� AS ��������," & _
                            "1 as ͼ�� " & _
                    "from ������� a," & _
                         "������� b," & _
                         "(Select a.���,Sum(b.�ּ�*a.����) As �����۸�,Sum(b.�ּ�*a.����*Nvl(a.�ۿ�,1)) As ���۸� " & _
                         "From ������ͼƼ� a," & _
                              "�շѼ�Ŀ b " & _
                         "Where b.�շ�ϸĿid=a.�շ�ϸĿid " & _
                               "and b.ִ������<=SYSDATE " & _
                               "and (b.��ֹ���� IS NULL OR b.��ֹ����>SYSDATE) " & _
                         "group by a.��� " & _
                         ") c " & _
                    "where a.ĩ�� = 1 AND a.�ϼ���� = b.���(+) " & _
                          "and a.���=c.���(+) and a.�ϼ����=[1]"
            
        End If
        
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Mid(tvw.SelectedItem.Key, 2)))
        If rs.BOF = False Then Call FillLvw(lvw, rs)
                
    Case "�����Ŀ"
        
        If lvw.SelectedItem Is Nothing Then Exit Function
        lngKey = Val(Mid(lvw.SelectedItem.Key, 2))
        If lngKey = 0 Then Exit Function
        
        gstrSQL = GetPublicSQL(SQL.���������Ŀ, CStr(lngKey))
    
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
        
        If rs.BOF = False Then
            Do While Not rs.EOF
                
                If Val(vsf.RowData(vsf.Rows - 1)) > 0 Then
                    vsf.Rows = vsf.Rows + 1
                End If
                
                vsf.RowData(vsf.Rows - 1) = zlCommFun.NVL(rs("ID"), 0)
                vsf.TextMatrix(vsf.Rows - 1, mCol.��Ŀ����) = zlCommFun.NVL(rs("��Ŀ����"))
                vsf.TextMatrix(vsf.Rows - 1, mCol.��Ŀ���) = zlCommFun.NVL(rs("��Ŀ���"))
                vsf.TextMatrix(vsf.Rows - 1, mCol.��鲿λ) = zlCommFun.NVL(rs("��鲿λ"))
                vsf.TextMatrix(vsf.Rows - 1, mCol.�ɼ���ʽ) = zlCommFun.NVL(rs("�ɼ���ʽ"))
                vsf.TextMatrix(vsf.Rows - 1, mCol.����걾) = zlCommFun.NVL(rs("����걾"))
                vsf.TextMatrix(vsf.Rows - 1, mCol.�����۸�) = zlCommFun.NVL(rs("�����۸�"))
                vsf.TextMatrix(vsf.Rows - 1, mCol.���۸�) = zlCommFun.NVL(rs("���۸�"))
                vsf.TextMatrix(vsf.Rows - 1, mCol.�ۿ�) = zlCommFun.NVL(rs("�ۿ�"))
                                            
                rs.MoveNext
            Loop
        End If
    End Select
    
    RefreshData = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function MenuClick(ByVal strMenuItem As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ����ݱ༭/����
    '------------------------------------------------------------------------------------------------------------------
    Dim lngKey As Long
    Dim blnTran As Boolean
    Dim lngLoop As Long
    Dim strSQL() As String
                
    On Error GoTo errHand
    
    ReDim Preserve strSQL(1 To 1)
    
    Select Case strMenuItem
    Case "���ӷ���"
        
        If tvw.SelectedItem Is Nothing Then Exit Function
        
        If Not frmKindClassEdit.ShowEdit(Me, 0, Val(Mid(tvw.SelectedItem.Key, 2))) Then Exit Function
        
        
    Case "�޸ķ���"
        
        If tvw.SelectedItem Is Nothing Then Exit Function
        If tvw.SelectedItem.Key = "K0" Then Exit Function
                        
        If Not frmKindClassEdit.ShowEdit(Me, Val(Mid(tvw.SelectedItem.Key, 2)), Val(Mid(tvw.SelectedItem.Parent.Key, 2))) Then Exit Function
        
        
    Case "ɾ������"
        
        If tvw.SelectedItem Is Nothing Then Exit Function
        If tvw.SelectedItem.Key = "K0" Then Exit Function
        
        If MsgBox("�����Ҫɾ����" & tvw.SelectedItem.Text & "�����ࣿ" & vbCrLf & "ɾ������ͬʱҲɾ����Ӧ��������ͺ������Ŀ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        
        lngKey = Val(Mid(tvw.SelectedItem.Key, 2))
        If lngKey = 0 Then Exit Function
        
        strSQL(ReDimArray(strSQL)) = "ZL_�������_DELETE(" & lngKey & ")"
        
    Case "��������"
        If tvw.SelectedItem Is Nothing Then Exit Function
        
        If Not frmKindEdit.ShowEdit(Me, 0, Val(Mid(tvw.SelectedItem.Key, 2))) Then Exit Function
        
    Case "�޸�����"
        If tvw.SelectedItem Is Nothing Then Exit Function
        If lvw.SelectedItem Is Nothing Then Exit Function
        lngKey = Val(Mid(lvw.SelectedItem.Key, 2))
        If lngKey = 0 Then Exit Function
        
        If Not frmKindEdit.ShowEdit(Me, lngKey, Val(Mid(tvw.SelectedItem.Key, 2))) Then Exit Function
        
    Case "ɾ������"
        If tvw.SelectedItem Is Nothing Then Exit Function
        If lvw.SelectedItem Is Nothing Then Exit Function
        
        If MsgBox("�����Ҫɾ����" & lvw.SelectedItem.Text & "������Ӧ�������Ŀ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        
        lngKey = Val(Mid(lvw.SelectedItem.Key, 2))
        If lngKey = 0 Then Exit Function
        
        strSQL(ReDimArray(strSQL)) = "ZL_�������_DELETE(" & lngKey & ")"
        
    Case "�����Ŀ"
        If lvw.SelectedItem Is Nothing Then Exit Function
        lngKey = Val(Mid(lvw.SelectedItem.Key, 2))
        If lngKey = 0 Then Exit Function
        
        frmKindCustom.ShowEdit Me, lngKey
        
    End Select
    
    blnTran = True
    gcnOracle.BeginTrans
    
    For lngLoop = 1 To UBound(strSQL)
        If strSQL(lngLoop) <> "" Then Call zlDatabase.ExecuteProcedure(strSQL(lngLoop), Me.Caption)
    Next
    
    gcnOracle.CommitTrans
    blnTran = False
        
    Select Case strMenuItem
    Case "ɾ������"
        
        If Not (tvw.SelectedItem Is Nothing) Then tvw.Nodes.Remove tvw.SelectedItem.Index
        
        Call ClearData("�������;�����Ŀ")
        Call RefreshData("�������")
        If Not (lvw.SelectedItem Is Nothing) Then Call RefreshData("�����Ŀ")
                
        
    Case "ɾ������"
    
        'ɾ����
        lngLoop = lvw.SelectedItem.Index
        lvw.ListItems.Remove lngLoop
        Call NextLvwPos(lvw, lngLoop)
        
    End Select
    
    Call AdjustEnableState
    Call RefreshStateInfo
    
    MenuClick = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
    If blnTran Then gcnOracle.RollbackTrans
End Function

Private Sub PrintData(ByVal bytMode As Byte)
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ� ��ӡ����
    '������ bytMode                         ��ӡ��ʽ��1-��ӡ��2-Ԥ����3-�����Excel��
    '------------------------------------------------------------------------------------------------------------------
    Dim objPrint As New zlPrint1Grd
    Dim objRow As zlTabAppRow
    
    If lvw.SelectedItem Is Nothing Then Exit Sub
    
    If UserInfo.���� = "" Then Call GetUserInfo

    objPrint.Title.Text = "�����Ŀ�嵥"
    Call CopyGrid(vsf, vsfPrint)
    
    Set objRow = New zlTabAppRow
    objRow.Add "���ͣ�" & lvw.SelectedItem.Text
    objRow.Add ""
    
    objPrint.UnderAppRows.Add objRow
    
    Set objPrint.Body = vsfPrint

    If bytMode = 1 Then bytMode = zlPrintAsk(objPrint)

    If bytMode >= 1 And bytMode <= 3 Then Call zlPrintOrView1Grd(objPrint, bytMode)
        
End Sub





'���������弰��ؼ����¼�����******************************************************************************************
Private Sub Form_Activate()
    
    If mblnStartUp = False Then Exit Sub
    DoEvents
    
    Call mnuViewIcon_Click(lvw.View)
    Call mnuViewRefresh_Click
    
    mblnStartUp = False
End Sub

Private Sub Form_Load()
    
    mblnStartUp = True
    
    Call RestoreWinState(Me, App.ProductName)
    Call InitLoad
    Call ApplyPrivilege(gstrPrivs)
    
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    
    If Me.WindowState = 1 Then Exit Sub
    
    '�����������
    
    If imgX_S.Top > Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0) - 1000 Then
        imgX_S.Top = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0) - 1000
    End If
    
    If imgY_S.Left > Me.ScaleWidth - 1000 Then
        imgY_S.Left = Me.ScaleWidth - 1000
    End If
    
    With tvw
        .Left = 0
        .Top = IIf(cbrThis.Visible, cbrThis.Height, 0)
        .Width = imgY_S.Left
        .Height = Me.ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
    End With
    
    With imgY_S
        .Top = tvw.Top
        .Height = tvw.Height
    End With
    
    With lvw
        .Left = imgY_S.Left + imgY_S.Width
        .Top = tvw.Top
        .Width = Me.ScaleWidth - .Left
        .Height = imgX_S.Top - .Top
    End With
    
    With imgX_S
        .Left = lvw.Left
        .Width = lvw.Width
    End With
    
    With vsf
        .Left = lvw.Left
        .Top = imgX_S.Top + imgX_S.Height
        .Width = Me.ScaleWidth - .Left
        .Height = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0) - .Top
    End With

    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Cancel = mblnStartUp
    If Cancel Then Exit Sub
    
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "������", mstrVsf)
                
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub imgX_S_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    
    On Error Resume Next
    
    imgX_S.Top = imgX_S.Top + Y
    
    If imgX_S.Top < 1500 Then imgX_S.Top = 1500
    If Me.Height - imgX_S.Top - imgX_S.Height < 1000 Then imgX_S.Top = Me.Height - imgX_S.Height - 1000
    
            
    Form_Resize
End Sub

Private Sub imgY_S_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    
    imgY_S.Left = imgY_S.Left + X
    
    If imgY_S.Left < 1500 Then imgY_S.Left = 1500
    If Me.Width - imgY_S.Left - imgY_S.Width < 1000 Then imgY_S.Left = Me.Width - imgY_S.Width - 1000

    Form_Resize
End Sub

Private Sub lvw_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call zlControl.LvwSortColumn(lvw, ColumnHeader.Index)
End Sub

Private Sub lvw_DblClick()
    If mnuEdit.Visible And mnuEditModify.Visible And mnuEditModify.Enabled Then Call mnuEditModify_Click
End Sub

Private Sub lvw_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim lngKey As Long
    
    If mstrKey = Item.Key Then Exit Sub
    mstrKey = Item.Key
    
    '����
    lngKey = Val(vsf.RowData(vsf.Row))
    
    Call ClearData("�����Ŀ")
    Call RefreshData("�����Ŀ")
    
    '�ָ�
    Call AdjustEnableState
    Call RefreshStateInfo
    
End Sub

Private Sub lvw_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call lvw_DblClick
End Sub

Private Sub lvw_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 2 Then Exit Sub
    
    mbytPopMenu = 2
    Set mobjPopMenu = New clsPopMenu
    mobjPopMenu.ShowPopupMenuByCursor

End Sub

Private Sub mnuEditClassAdd_Click()
    Call MenuClick("���ӷ���")
End Sub

Private Sub mnuEditClassDelete_Click()
    Call MenuClick("ɾ������")
End Sub

Private Sub mnuEditClassModify_Click()
    Call MenuClick("�޸ķ���")
End Sub

Private Sub mnuEditDelete_Click()
    Call MenuClick("ɾ������")
End Sub

Private Sub mnuEditModify_Click()
    Call MenuClick("�޸�����")
End Sub

Private Sub mnuEditAdd_Click()
    Call MenuClick("��������")
End Sub

Private Sub mnuEditSelect_Click()
    Call MenuClick("�����Ŀ")
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileOutExcel_Click()
    Call PrintData(3)
End Sub

Private Sub mnuFilePrint_Click()
    Call PrintData(1)
End Sub

Private Sub mnuFilePrintSet_Click()
    Call zlPrintSet
End Sub

Private Sub mnuFilePrintView_Click()
    Call PrintData(2)
End Sub

Private Sub mnuHelpAbout_Click()
    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
End Sub

Private Sub mnuHelpTopic_Click()
   Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hWnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hWnd)
End Sub

Private Sub mnuViewIcon_Click(Index As Integer)
    mnuViewIcon(0).Checked = False
    mnuViewIcon(1).Checked = False
    mnuViewIcon(2).Checked = False
    mnuViewIcon(3).Checked = False
    
    mnuViewIcon(Index).Checked = True
    
    lvw.View = Index
End Sub

Private Sub mnuViewRefresh_Click()
    Dim strKey As String
    Dim strKeyClass As String
        
    '����������ͷ��ࡢ�������
    If Not (tvw.SelectedItem Is Nothing) Then strKeyClass = tvw.SelectedItem.Key
    strKey = zlControl.LvwSaveItem(lvw)
            
    Call ClearData("�����Ŀ;�������;������ͷ���")
    
    Call RefreshData("������ͷ���")
    
    '�ָ�ˢ��ǰѡ���������ͷ���
    
    tvw.Nodes(1).Selected = True
    tvw.Nodes(1).Expanded = True
    
    On Error Resume Next
    tvw.Nodes(strKeyClass).Selected = True
    tvw.Nodes(strKeyClass).EnsureVisible
    On Error GoTo 0
    
    If Not (tvw.SelectedItem Is Nothing) Then
        Call RefreshData("�������")
        
        '�ָ�ˢ��ǰѡ����������
        Call zlControl.LvwRestoreItem(lvw, strKey)
        
        If Not (lvw.SelectedItem Is Nothing) Then Call RefreshData("�����Ŀ")
    End If
    
    Call AdjustEnableState
    Call RefreshStateInfo
    
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = mnuViewStatus.Checked
    Call Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    cbrThis.Visible = mnuViewToolButton.Checked
    Call Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim intLoop As Integer
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For intLoop = 1 To tbrThis.Buttons.Count
        tbrThis.Buttons(intLoop).Caption = IIf(mnuViewToolText.Checked, tbrThis.Buttons(intLoop).Tag, "")
    Next
    cbrThis.Bands(1).MinHeight = tbrThis.Height
    Call Form_Resize
    
End Sub

Private Sub mobjPopMenu_MenuBeforeShow(Cancel As Boolean)
    
    Select Case mbytPopMenu
    Case 1
        
        If mnuEdit.Visible = False Then Exit Sub
        
        If mnuEditClassAdd.Visible Then mobjPopMenu.Add 1, mnuEditClassAdd.Caption, , , mnuEditClassAdd.Enabled
        If mnuEditClassModify.Visible Then mobjPopMenu.Add 2, mnuEditClassModify.Caption, , , mnuEditClassModify.Enabled
        If mnuEditClassDelete.Visible Then mobjPopMenu.Add 3, mnuEditClassDelete.Caption, , , mnuEditClassDelete.Enabled
        
    Case 2
        
        If mnuEdit.Visible = False Then Exit Sub
        
        If mnuEditAdd.Visible Then mobjPopMenu.Add 1, mnuEditAdd.Caption, , , mnuEditAdd.Enabled
        If mnuEditModify.Visible Then mobjPopMenu.Add 2, mnuEditModify.Caption, , , mnuEditModify.Enabled
        If mnuEditDelete.Visible Then mobjPopMenu.Add 3, mnuEditDelete.Caption, , , mnuEditDelete.Enabled
        
        mobjPopMenu.Add 4, "-", , 2, True
        
        If mnuEditSelect.Visible Then mobjPopMenu.Add 5, mnuEditSelect.Caption, , , mnuEditSelect.Enabled
    
    End Select
    
End Sub

Private Sub mobjPopMenu_MenuClick(ByVal Key As Long, ByVal Caption As String)
    Select Case mbytPopMenu
    Case 1
        Select Case Key
        Case 1
            Call mnuEditClassAdd_Click
        Case 2
            Call mnuEditClassModify_Click
        Case 3
            Call mnuEditClassDelete_Click
        End Select
    Case 2
        Select Case Key
        Case 1
            Call mnuEditAdd_Click
        Case 2
            Call mnuEditModify_Click
        Case 3
            Call mnuEditDelete_Click
        Case 5
            Call mnuEditSelect_Click
        End Select
    End Select
End Sub


Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(tbrThis.hWnd, objPoint)
    
    Select Case Button.Key
    Case "Ԥ��"
        Call mnuFilePrintView_Click
    Case "��ӡ"
        
        Call mnuFilePrint_Click
        
    Case "����"
                
        mbytPopMenu = 1
        Set mobjPopMenu = New clsPopMenu
        Call mobjPopMenu.ShowPopupMenu(objPoint.X * 15 + Button.Left - 15, objPoint.Y * 15 + Button.Top + Button.Height + 15)
        
    Case "����"
        Call mnuEditAdd_Click
    Case "�޸�"
        Call mnuEditModify_Click
    Case "ɾ��"
        Call mnuEditDelete_Click
    Case "��Ŀ"
        Call mnuEditSelect_Click
    Case "�б�"
        Call mnuViewIcon_Click(IIf(lvw.View = 3, 0, lvw.View + 1))
    Case "����"
        Call mnuHelpTopic_Click
    Case "�˳�"
        Call mnuFileExit_Click
    End Select
End Sub

Private Sub tbrThis_ButtonDropDown(ByVal Button As MSComctlLib.Button)
    Call tbrThis_ButtonClick(Button)
End Sub

Private Sub tbrThis_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
    Case "Large"
        Call mnuViewIcon_Click(0)
    Case "Small"
        Call mnuViewIcon_Click(1)
    Case "List"
        Call mnuViewIcon_Click(2)
    Case "Detail"
        Call mnuViewIcon_Click(3)
    End Select
End Sub

Private Sub tbrThis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuViewTool
End Sub

Private Sub tvw_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 2 Then
    
        mbytPopMenu = 1
        Set mobjPopMenu = New clsPopMenu
        mobjPopMenu.ShowPopupMenuByCursor
        
    End If
End Sub

Private Sub tvw_NodeClick(ByVal Node As MSComctlLib.Node)
    
    Call ClearData("�������;�����Ŀ")
    
    Call RefreshData("�������")
    
    If Not (lvw.SelectedItem Is Nothing) Then Call RefreshData("�����Ŀ")
    
    Call AdjustEnableState
    
End Sub

Private Sub vsf_GotFocus()
    vsf.BackColorSel = COLOR.����
End Sub

Private Sub vsf_LostFocus()
    vsf.BackColorSel = COLOR.�ǽ���
End Sub


Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

