VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmRequestStuffList 
   Caption         =   "�����������"
   ClientHeight    =   4980
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9615
   Icon            =   "frmRequestStuffList.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4980
   ScaleWidth      =   9615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picSeparate_s 
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   180
      MousePointer    =   7  'Size N S
      ScaleHeight     =   360
      ScaleWidth      =   4815
      TabIndex        =   0
      Top             =   2700
      Width           =   4815
      Begin VB.Label lblRange 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ѯ��Χ:1999��8��12����1999��9��12��"
         Height          =   180
         Left            =   0
         TabIndex        =   4
         Top             =   200
         Width           =   3690
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         Caption         =   "�ɱ���"
         Height          =   180
         Left            =   0
         TabIndex        =   3
         Top             =   20
         Width           =   900
      End
      Begin VB.Label lbl2 
         AutoSize        =   -1  'True
         Caption         =   "�ۼ۽�"
         Height          =   180
         Left            =   1890
         TabIndex        =   2
         Top             =   20
         Width           =   900
      End
      Begin VB.Label lbl3 
         AutoSize        =   -1  'True
         Caption         =   "��۽�"
         Height          =   180
         Left            =   3690
         TabIndex        =   1
         Top             =   20
         Width           =   900
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDetail 
      Height          =   975
      Left            =   360
      TabIndex        =   5
      Top             =   3120
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   1720
      _Version        =   393216
      FixedCols       =   0
      AllowBigSelection=   0   'False
      FocusRect       =   0
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin ComCtl3.CoolBar cbrTool 
      Height          =   780
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   1376
      BandCount       =   2
      _CBWidth        =   11775
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tlbTool"
      MinWidth1       =   6000
      MinHeight1      =   720
      Width1          =   6210
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Caption2        =   "���첿��"
      Child2          =   "cboStock"
      MinWidth2       =   3000
      MinHeight2      =   300
      Width2          =   3345
      NewRow2         =   0   'False
      Begin MSComctlLib.Toolbar tlbTool 
         Height          =   720
         Left            =   165
         TabIndex        =   8
         Top             =   30
         Width           =   7515
         _ExtentX        =   13256
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
            NumButtons      =   16
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
               Key             =   "Add"
               Description     =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   3
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Hank"
                     Text            =   "�ֹ���д"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Auto"
                     Text            =   "�Զ�����"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�޸�"
               Key             =   "Modify"
               Description     =   "�޸�"
               Object.ToolTipText     =   "�޸�"
               Object.Tag             =   "�޸�"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ɾ��"
               Key             =   "Delete"
               Description     =   "ɾ��"
               Object.ToolTipText     =   "ɾ��"
               Object.Tag             =   "ɾ��"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Edit1"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˲�"
               Key             =   "Check"
               Object.ToolTipText     =   "�˲�"
               Object.Tag             =   "�˲�"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Receive"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "DisReceive"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "EditSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Search"
               Description     =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ˢ��"
               Key             =   "Refresh"
               Description     =   "ˢ��"
               Object.ToolTipText     =   "ˢ��"
               Object.Tag             =   "ˢ��"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "FindSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Description     =   "��������"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Exit"
               Description     =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageIndex      =   11
            EndProperty
         EndProperty
         MouseIcon       =   "frmRequestStuffList.frx":014A
         Begin VB.Timer LimitTime 
            Enabled         =   0   'False
            Interval        =   8000
            Left            =   6660
            Top             =   180
         End
      End
      Begin VB.ComboBox cboStock 
         Height          =   300
         Left            =   8685
         TabIndex        =   7
         Text            =   "cboStock"
         Top             =   240
         Width           =   3000
      End
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   9
      Top             =   4620
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmRequestStuffList.frx":0464
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11880
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
      Left            =   0
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffList.frx":0CF8
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffList.frx":0F18
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffList.frx":1138
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffList.frx":1354
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffList.frx":1574
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffList.frx":1794
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffList.frx":19B0
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffList.frx":1BCC
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffList.frx":1DE6
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffList.frx":1F40
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffList.frx":215C
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffList.frx":237C
            Key             =   "check"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsHot 
      Left            =   600
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffList.frx":2596
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffList.frx":27B6
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffList.frx":29D6
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffList.frx":2BF2
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffList.frx":2E12
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffList.frx":3032
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffList.frx":324E
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffList.frx":346A
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffList.frx":3684
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffList.frx":37DE
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffList.frx":39FE
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffList.frx":3C1E
            Key             =   "check"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid mshList 
      Height          =   885
      Left            =   240
      TabIndex        =   10
      Top             =   1320
      Width           =   4935
      _cx             =   8705
      _cy             =   1561
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
      BackColorSel    =   16769992
      ForeColorSel    =   0
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
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
      ExplorerBar     =   1
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
      VirtualData     =   0   'False
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
      Begin VB.Menu mnuFileBillPrint 
         Caption         =   "���ݴ�ӡ(&B)"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuFileBillPreview 
         Caption         =   "����Ԥ��(&L)"
      End
      Begin VB.Menu mnuFileLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "�����&Excel"
      End
      Begin VB.Menu mnuFileLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileParameter 
         Caption         =   "��������(&R)"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFileLine3 
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
         Begin VB.Menu mnuEditAddHank 
            Caption         =   "�ֹ���д(&H)"
         End
         Begin VB.Menu mnuEditAddAuto 
            Caption         =   "�Զ�����(&A)"
         End
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "�޸�(&M)"
      End
      Begin VB.Menu mnuEditDel 
         Caption         =   "ɾ��(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCheck 
         Caption         =   "�˲�(&V)"
      End
      Begin VB.Menu mnuEditReceive 
         Caption         =   "����(&R)"
      End
      Begin VB.Menu mnuEditDisReceive 
         Caption         =   "����(&D)"
      End
      Begin VB.Menu mnuEditLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditImport 
         Caption         =   "�����깺��(&I)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditDisplay 
         Caption         =   "�鿴����(&W)"
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
      Begin VB.Menu mnuViewSearch 
         Caption         =   "����(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewLine4 
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
         Caption         =   "&WEB�ϵ�����"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "������ҳ(&H)"
         End
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "������̳(&F)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "���ͷ���(&M)..."
         End
      End
      Begin VB.Menu mnuHelpLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)..."
      End
   End
End
Attribute VB_Name = "frmRequestStuffList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrFind As String
Private mblnBootUp As Boolean
Private mlastRow As Long                '�ϴε������
Private mintPreCol As Integer           'ǰһ�ε���ͷ��������
Private mintsort As Integer             'ǰһ�ε���ͷ������
Private mintPreDetailCol As Integer     'ǰһ�ε������������
Private mintDetailsort As Integer       'ǰһ�ε����������

Private mdtStartDate As Date
Private mdtEndDate As Date
Private mdtVerifyStart As Date
Private mdtVerifyEnd As Date
Private mstrPrivs As String
Private mintUnit As String
Private mstrOthers() As String ' 0-��¼״̬(�ƻ�����),1-��ʼ����,2-��������,3-����id,4-�Է�����id(��������id����Ʒ���(�ƻ���)),5-������,6-�����,7-��Ӧ��ID,8-������,9-��ʼ��������,10-������������,11-��ʼ��Ʊ��,12-������Ʊ��,13-������Ϣ
Private mlngModule As Long
Private mblnCostView As Boolean             '�鿴�ɱ��������Ϣ true-����鿴 false-������鿴
Private Const mstrCaption As String = "�����������"
Private mbln����˲� As Boolean     '�����Ƿ���Ҫ�˲� true-��Ҫ false-����Ҫ
Private mintFindDay As Integer      '��ѯ������Χ
Private mint��ȷ���� As Integer             '��ʾ����д���쵥ʱ���Ƿ���ȷ���ĵ�����
Private mint�������� As Integer                          '0-����Ҫ����;1-��Ҫ����


'----------------------------------------------------------------------------------------------------------
'���˺�:����С��λ���ĸ�ʽ��
'�޸�:2007/03/06
Private mOraFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------


Private Sub cboStock_Click()
    If mblnBootUp Then mnuViewRefresh_Click
End Sub

Private Sub cboStock_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If cboStock.ListCount = 0 Then Call zlControl.ControlSetFocus(mshList): Exit Sub
    
    If cboStock.ListIndex >= 0 Then
        If Val(cboStock.Tag) = cboStock.ItemData(cboStock.ListIndex) Then
            Call zlControl.ControlSetFocus(mshList, True)
            Exit Sub
        End If
    End If
    
    If Select����ѡ����(Me, cboStock, Trim(cboStock.Text), "W,V,K", True) = False Then
        Exit Sub
    End If
    If cboStock.ListIndex >= 0 Then
        cboStock.Tag = cboStock.ItemData(cboStock.ListIndex)
    End If
End Sub

Private Sub cbrTool_Resize()
    If mblnBootUp = False Then Exit Sub
    Form_Resize
End Sub

Public Sub ShowList(ByVal frmMain As Variant)
    Dim strFind As String
    
    mblnBootUp = False
    If Not CheckDepend Then Exit Sub            '���������Բ���
                
    SetVisable  '����Ȩ�����ò�ͬ����ʾ��Ŀ
    mintFindDay = Val(zlDatabase.GetPara("��ѯ����", glngSys, mlngModule, 1))
    mdtStartDate = Format(DateAdd("d", -mintFindDay, sys.Currentdate), "yyyy-MM-dd")
    mdtEndDate = Format(sys.Currentdate, "yyyy-MM-dd")
    mdtVerifyStart = "1901-01-01"
    mdtVerifyEnd = "1901-01-01"
    
    strFind = " AND A.��¼״̬ = 1 And A.������� is Null And A.�������� Between To_Date('" & Format(mdtStartDate, "yyyy-mm-dd") & " 00:00:00','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(mdtEndDate, "yyyy-mm-dd") & " 23:59:59','YYYY-MM-DD HH24:MI:SS')"
    mstrFind = strFind
    
    GetList (mstrFind)  '�г�����ͷ
    RestoreWinState Me, App.ProductName, mstrCaption
    '�ָ����Ի��������ú󣬻���Ҫ��Ȩ�޿��Ƶ��н�һ������
    With mshList
        .ColWidth(.ColIndex("�ɱ����")) = IIf(mblnCostView = True, 900, 0) '���ϲ�����Ա�Ƿ���Կ��ɱ���
        .ColWidth(.ColIndex("��۽��")) = IIf(mblnCostView = True, 900, 0)
    End With
    
    With mshDetail
        .ColWidth(12) = IIf(mblnCostView = True, 900, 0)
        .ColWidth(13) = IIf(mblnCostView = True, 900, 0)
        .ColWidth(16) = IIf(mblnCostView = True, 900, 0)
    End With
    
    mblnBootUp = True
    
    If IsObject(frmMain) Then
        Me.Show , frmMain
    Else
        zlCommFun.ShowChildWindow Me.hwnd, frmMain
    End If
    
    Me.ZOrder 0

End Sub

'�������������
Private Function CheckDepend() As Boolean
    
    Dim rsDepend As New Recordset
    Dim strStock As String
    
    CheckDepend = False

    On Error GoTo ErrHandle
    strStock = " And B.���� In('�Ƽ���','���Ŀ�','���ϲ���')"
    
    gstrSQL = "" & _
        "   SELECT DISTINCT a.id, a.���� " & _
        "   FROM ��������˵�� c, �������ʷ��� b, ���ű� a " & _
        "   Where c.�������� = b.���� And (A.վ��=[1] or A.վ�� is null) " & strStock & _
        "           AND a.id = c.����id " & _
        "           AND TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01'"
    Set rsDepend = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, gstrNodeNo)
    If rsDepend.EOF Then
        MsgBox "����������Ϣ��ȫ,��鿴���Ź���", vbInformation, gstrSysName
        rsDepend.Close
        Exit Function
    End If
    rsDepend.Close
    
    gstrSQL = "" & _
        "   SELECT DISTINCT a.id, a.���� " & _
        "   FROM ��������˵�� c, �������ʷ��� b, ���ű� a " & _
        "   Where c.�������� = b.���� And (A.վ��=[2] or A.վ�� is null) " & strStock & _
        "           AND a.id = c.����id " & _
        "           and a.id in (select ����id from ������Ա where ��Աid= [1]) " & _
        "           AND TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01'"
    Set rsDepend = zlDatabase.OpenSQLRecord(gstrSQL, "�����������", UserInfo.Id, gstrNodeNo)
    If rsDepend.EOF Then
        MsgBox "�㲻�����Ŀ⡢���ϲ��š����Ƽ��ҵĹ�����Ա�����ܽ��룡", vbInformation, gstrSysName
        rsDepend.Close
        Exit Function
    End If
    
    With cboStock
        .Clear
        Do While Not rsDepend.EOF
            .AddItem rsDepend!����
            .ItemData(.NewIndex) = rsDepend!Id
            If rsDepend!Id = glngDeptId Then
                .ListIndex = .NewIndex
            End If
            rsDepend.MoveNext
        Loop
        rsDepend.Close
        
        If .ListIndex = -1 Then
            .ListIndex = 0
        End If
    End With
    
    CheckDepend = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub GetList(ByVal strFind As String)
    Dim rsList As New Recordset
    Dim strUserPart As String
    Dim intCol As Integer
    
    '����ͳ�ƺϼƽ��
    Dim dbl1 As Double
    Dim dbl2 As Double
    Dim dbl3 As Double
    Dim n As Long
    Dim strFormat As String
    
    On Error GoTo ErrHandle
    strFormat = "0.00##"
    
    mlastRow = 0
    
    mshList.Redraw = False
    strUserPart = " And A.�ⷿID+0=[1]"
    
    gstrSQL = "SELECT A.NO, C.���� AS ���Ͽⷿ,LTRIM(TO_CHAR (SUM (A.�ɱ����), " & mOraFMT.FM_��� & ")) AS �ɱ����, " & _
        " LTRIM(TO_CHAR ( (SUM (A.���۽��)), " & mOraFMT.FM_��� & ")) AS �ۼ۽��,LTRIM(TO_CHAR (SUM (A.���۽�� - A.�ɱ����)," & mOraFMT.FM_��� & ")) AS ��۽��, A.������, " & _
        " TO_CHAR (MIN(A.��������), 'YYYY-MM-DD HH24:MI:SS') AS ��������, " & _
        " a.�˲���, To_Char(Min(a.�˲�����), 'YYYY-MM-DD HH24:MI:SS') As �˲�����," & _
        " A.�����, TO_CHAR (MIN(A.�������), 'YYYY-MM-DD HH24:MI:SS') AS �������, A.��¼״̬, A.��ҩ�� ������,A.ժҪ,nvl(a.��ҩ��ʽ,0) as ��ҩ��ʽ " & _
        " FROM ҩƷ�շ���¼ A, ���ű� B,���ű� C " & _
        " WHERE A.�ⷿID = B.ID AND A.�Է�����ID=C.ID AND A.���� = 19 AND  A.���ϵ��=1 " & _
        " And (A.��ҩ�� Is NULL Or A.��ҩ���� Is Not NULL)" & _
        strUserPart & strFind & _
        " GROUP BY A.NO,C.����,A.������,A.�˲���,A.�����,A.��¼״̬,A.��ҩ��,A.ժҪ,nvl(a.��ҩ��ʽ,0) " & _
        " ORDER BY NO DESC, �������� ASC "
        
     'mstrOthers(0 To 13) As String ' 0-��¼״̬(�ƻ�����),1-��ʼ����,2-��������,3-����id,4-�Է�����id(��������id����Ʒ���(�ƻ���)),5-������,6-�����,7-��Ӧ��ID,8-������,9-��ʼ��������,10-������������,11-��ʼ��Ʊ��,12-������Ʊ��,13-������Ϣ
    '������Χ:[1]-�ⷿid,[2]:��ʼ��������,[3]������������,[4]��ʼ�������,[5] �����������,[6]-��¼״̬,[7]��ʼ���ݺ�,[8]�������ݺ�,[9]����id,[10]�Է�����id,[11]������,[12]�����[13]-��Ӧ��ID,[14]-������,[15]-��ʼ��������,[16]-������������,[17]-��ʼ��Ʊ��,[18]-������Ʊ��,[19]-������Ϣ
    
    '��ʼ��������
    mstrOthers(9) = IIf(Trim(mstrOthers(9)) = "", "1901-01-01", mstrOthers(9))
    mstrOthers(10) = IIf(Trim(mstrOthers(10)) = "", "1901-01-01", mstrOthers(10))
    
    Set rsList = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, cboStock.ItemData(cboStock.ListIndex), _
        CDate(Format(mdtStartDate, "yyyy-mm-dd") & " 00:00:00"), CDate(Format(mdtEndDate, "yyyy-mm-dd") & " 23:59:59"), _
        CDate(Format(mdtVerifyStart, "yyyy-mm-dd") & " 00:00:00"), CDate(Format(mdtVerifyEnd, "yyyy-mm-dd") & " 23:59:59"), _
        Val(mstrOthers(0)), mstrOthers(1), mstrOthers(2), Val(mstrOthers(3)), _
        Val(mstrOthers(4)), mstrOthers(5), mstrOthers(6), _
        Val(mstrOthers(7)), mstrOthers(8), CDate(mstrOthers(9) & " 00:00:00"), CDate(mstrOthers(10) & " 23:59:59"), _
         mstrOthers(11), mstrOthers(12), mstrOthers(13) & "%")
      
    Set mshList.DataSource = rsList
    With mshList
        If .Rows = 1 Then
            .Rows = .Rows + 100
            .Row = 1
'            .Redraw = True
            
            .TopRow = 1
            .Rows = .Rows - 99
        End If
        .Row = 1
        .Col = 0
'        .ColSel = .Cols - 1
        
        For intCol = 0 To .Cols - 1
            .ColKey(intCol) = .TextMatrix(0, intCol)
        Next
    End With
    
    SetListColWidth
    
    'ͳ�ƺϼƽ��
    If (Not rsList.EOF) And (Not rsList.BOF) Then
        rsList.MoveFirst
        Do While Not rsList.EOF
            dbl1 = dbl1 + IIf(IsNull(rsList!�ɱ����), 0, rsList!�ɱ����)
            dbl2 = dbl2 + IIf(IsNull(rsList!�ۼ۽��), 0, rsList!�ۼ۽��)
            dbl3 = dbl3 + IIf(IsNull(rsList!��۽��), 0, rsList!��۽��)
            rsList.MoveNext
        Loop
        rsList.MoveFirst
        
        lbl1.Caption = "�ɱ����ϼƣ�" & Format(dbl1, strFormat)
        lbl2.Caption = "�ۼ۽��ϼƣ�" & Format(dbl2, strFormat)
        lbl3.Caption = "��۽��ϼƣ�" & Format(dbl3, strFormat)
    Else
        lbl1.Caption = "�ɱ����ϼƣ�" & Format(0, strFormat)
        lbl2.Caption = "�ۼ۽��ϼƣ�" & Format(0, strFormat)
        lbl3.Caption = "��۽��ϼƣ�" & Format(0, strFormat)
    End If
    
    mshlist_EnterCell    '�г�������
    
    SetStrikeColor
    mshList.Redraw = True
    Call SetEnable
    
    staThis.Panels(2).Text = "��ǰ����" & rsList.RecordCount & "�ŵ���"
    rsList.Close
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub SetStrikeColor()
    Dim intStatus As Integer
    Dim intRow As Integer
    Dim intCol As Integer
    
    With mshList
        If .Rows <= 2 Then Exit Sub
        For intRow = 1 To .Rows - 1
            intStatus = Val(.TextMatrix(intRow, .ColIndex("��¼״̬")))
            If intStatus Mod 3 = 0 Then
                .Row = intRow
                For intCol = 0 To .Cols - 1
                    .Col = intCol
                    .CellForeColor = &H80000001
                Next
            End If
            If intStatus Mod 3 = 2 Then
                .Row = intRow
                For intCol = 0 To .Cols - 1
                    .Col = intCol
                    If .TextMatrix(intRow, .ColIndex("�����")) = "" Then
                        .CellForeColor = &HC0C0FF
                    Else
                        .CellForeColor = &HFF
                    End If
                Next
            End If
        Next
    End With
End Sub

'��ͷ�п��ʼ
Private Sub SetListColWidth()
    Dim intCol As Integer
    
    With mshList
        .ColAlignment(.ColIndex("�ɱ����")) = flexAlignRightCenter
        If mblnBootUp = False Then
            For intCol = 1 To .Cols - 1
                If intCol = 1 Then
                   .ColWidth(intCol) = 2000
                ElseIf intCol = .ColIndex("��¼״̬") Or intCol = .ColIndex("��ҩ��ʽ") Then
                    .ColWidth(intCol) = 0
                Else
                    .ColWidth(intCol) = 1000
                End If
            Next
        End If
        
        .ColWidth(.ColIndex("�ɱ����")) = IIf(mblnCostView = False, 0, 1000) '���ϲ�����Ա�Ƿ���Կ��ɱ���
        If mblnCostView = False Then
            .ColWidth(.ColIndex("��۽��")) = 0 '���ϲ�����Ա�Ƿ���Կ��ɱ���
        End If
    End With
End Sub


Private Sub SetDetailColWidth()
    Dim intCol As Integer
    Dim i As Integer
    On Error Resume Next
    
    With mshDetail
        .ColAlignment(8) = flexAlignRightCenter     '���Ч��
        .ColAlignment(9) = flexAlignRightCenter     '��д����
        .ColAlignment(10) = flexAlignCenterCenter    'ʵ������
        .ColAlignment(11) = flexAlignRightCenter     '��λ
        .ColAlignment(12) = flexAlignRightCenter     '�ɱ���
        .ColAlignment(13) = flexAlignRightCenter    '�ɱ����
        .ColAlignment(14) = flexAlignRightCenter    '�ۼ�
        .ColAlignment(15) = flexAlignRightCenter    '�ۼ۽��
        .ColAlignment(16) = flexAlignRightCenter    '���
                
        If mblnBootUp = False Then
            .ColWidth(0) = 0
            .ColWidth(1) = 2500
            For intCol = 2 To .Cols - 1
                .ColWidth(intCol) = 1000
            Next
            .ColWidth(16) = 0
        End If
        
        .ColWidth(12) = IIf(mblnCostView = False, 0, 1000)
        .ColWidth(13) = IIf(mblnCostView = False, 0, 1000)
        .ColWidth(16) = IIf(mblnCostView = False, 0, 1000)
    End With
End Sub


'����Ȩ�����ò�ͬ����ʾ��Ŀ
Private Sub SetVisable()
    '����������
    If mbln����˲� = False Then
        mnuEditCheck.Visible = False
        tlbTool.Buttons("Check").Visible = False
    Else
        mnuEditCheck.Visible = True
        tlbTool.Buttons("Check").Visible = True
    End If
    
    If Not zlStr.IsHavePrivs(gstrPrivs, "����") Then
        mnuEditAdd.Visible = False
        mnuEditModify.Visible = False
        mnuEditDel.Visible = False
        
        tlbTool.Buttons("Add").Visible = False
        tlbTool.Buttons("Modify").Visible = False
        tlbTool.Buttons("Delete").Visible = False
        tlbTool.Buttons("Edit1").Visible = False
        mnuEditLine1.Visible = False
    End If
    If Not zlStr.IsHavePrivs(gstrPrivs, "���") Then
        mnuEditReceive.Visible = False
        tlbTool.Buttons("Receive").Visible = False
    End If
    If Not zlStr.IsHavePrivs(gstrPrivs, "����") Then
        mnuEditDisReceive.Visible = False
        If mnuEditReceive.Visible = False Then mnuEditLine2.Visible = False
        tlbTool.Buttons("DisReceive").Visible = False
        tlbTool.Buttons("EditSeparate").Visible = mnuEditLine2.Visible
    End If
    If Not zlStr.IsHavePrivs(gstrPrivs, "���ݴ�ӡ") Then
        mnuFileBillPrint.Visible = False
        mnuFileBillPreview.Visible = False
    End If
    
End Sub

Private Sub Form_Activate()
    If mint�������� = 1 Then
        mnuEditDisReceive.Caption = "�������(&R)"
        tlbTool.Buttons("DisReceive").Caption = "�������"
    Else
        mnuEditDisReceive.Caption = "����(&D)"
        tlbTool.Buttons("DisReceive").Caption = "����"
    End If
End Sub

Private Sub Form_Load()
    Dim strReg As String
    Dim strOthers(0 To 13) As String
    Dim i As Integer
    mlngModule = glngModul
    mbln����˲� = IIf((zlDatabase.GetPara("������Ҫ�˲������ƿ�", glngSys, mlngModule, "0")) = 0, False, True)
    
    'ȡ�ƿ�ĳ����������
    mint�������� = Val(zlDatabase.GetPara("��������", glngSys, 1716))

    '������찴�������죬����ʹ��"����ɹ���"���ܡ�
    mint��ȷ���� = IIf(IS��������, 1, 0)
'    If mint��ȷ���� = 1 Then
'        mnuEditImport.Visible = False
'    End If
    
    For i = 0 To 13
        strOthers(i) = ""
    Next
    '������������
    strOthers(9) = "1901-01-01"
    strOthers(10) = "1901-01-01"
    
    '0-��¼״̬(�ƻ�����),1-��ʼ����,2-��������,3-����id,4-�Է�����id(��������id����Ʒ���(�ƻ���)),5-������,6-�����,7-��Ӧ��ID,8-������,9-��ʼ��������,10-������������,11-��ʼ��Ʊ��,12-������Ʊ��,13-������Ϣ
    mstrOthers = strOthers

    lblRange.Caption = "��ѯ��Χ:" & Format(sys.Currentdate, "yyyy��MM��dd��") & "��" & Format(sys.Currentdate, "yyyy��MM��dd��")
    
    mstrPrivs = gstrPrivs
    mblnCostView = zlStr.IsHavePrivs(mstrPrivs, "�鿴�ɱ���")
    If Not zlStr.IsHavePrivs(mstrPrivs, "��������") Then
        mnuFileParameter.Visible = False
    Else
        mnuFileParameter.Visible = True
    End If
    
    strReg = Val(zlDatabase.GetPara("���ĵ�λ", glngSys, mlngModule, "0"))
    mintUnit = Val(strReg)
  
    '���˺�:����С����ʽ����
    With mOraFMT
        .FM_�ɱ��� = GetFmtString(mintUnit, g_�ɱ���, True)
        .FM_��� = GetFmtString(mintUnit, g_���, True)
        .FM_���ۼ� = GetFmtString(mintUnit, g_�ۼ�, True)
        .FM_���� = GetFmtString(mintUnit, g_����, True)
    End With
    
    
    lbl1.Caption = ""
    lbl2.Caption = ""
    lbl3.Caption = ""
    lbl2.Left = lbl1.Left + lbl1.Width + 3000
    lbl3.Left = lbl2.Left + lbl2.Width + 3000
    If mblnCostView = False Then
        lbl1.Visible = False
        lbl3.Visible = False
    End If
   
   '2006-04-25:���˺�,ͳһ���ӱ�������ģ��Ĺ���
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, gstrPrivs)
     
End Sub

Private Sub Form_Resize()
    '����λ������
    
    On Error Resume Next
    If Me.WindowState <> vbMaximized Then
        If Me.Height < 8145 Then
            Me.Height = 8145
        End If
    End If
    
    cbrTool.Bands(1).MinHeight = tlbTool.Height
    With cbrTool
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth - .Left
    End With
    
    With picSeparate_s
        .Height = 360
        .Left = 0
        .Width = cbrTool.Width
        
    End With
    
    With mshList
        .Top = IIf(cbrTool.Visible, cbrTool.Height, 0)
        .Left = 0
        .Width = cbrTool.Width
        .Height = picSeparate_s.Top - .Top
    End With
    
    With mshDetail
        .Top = picSeparate_s.Top + picSeparate_s.Height + 100
        .Left = 0
        .Height = ScaleHeight - .Top - IIf(staThis.Visible, staThis.Height, 0)
        .Width = cbrTool.Width
    End With
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName, mstrCaption
End Sub

Private Sub mnuEditAddAuto_Click()
    '�Զ�����
    Dim strNo As String
    Dim blnSuccess As Boolean
    
    strNo = ""
    frmRequestStuffCard.ShowCard Me, strNo, 5, , mstrPrivs, blnSuccess, cboStock.ItemData(cboStock.ListIndex)
    If blnSuccess = True Then
        mnuViewRefresh_Click
    End If
    
End Sub

Private Sub mnuEditAddHank_Click()
    '�ֹ���д
    Dim strNo As String
    Dim blnSuccess As Boolean
    
    strNo = ""
    '����
    frmRequestStuffCard.ShowCard Me, strNo, 1, , mstrPrivs, blnSuccess, cboStock.ItemData(cboStock.ListIndex)
    If blnSuccess = True Then
        mnuViewRefresh_Click
    End If
    
End Sub

Private Sub mnuEditCheck_Click()
    '�˲飬����
    Dim strNo As String
    Dim blnSuccess As Boolean
    
    blnSuccess = False
    With mshList
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        strNo = .TextMatrix(.Row, 0)
        
        frmRequestStuffCard.ShowCard Me, strNo, 3, mshList.TextMatrix(mshList.Row, mshList.ColIndex("��¼״̬")), mstrPrivs, blnSuccess, cboStock.ItemData(cboStock.ListIndex)
        If blnSuccess = True Then
            mnuViewRefresh_Click
        End If
    End With
End Sub

Private Sub mnuEditDel_Click()
    'ɾ��
    Dim StrBillNo As String
    Dim intRow As Integer
    Dim strTitle As String
    Dim intReturn As Integer
    Dim intRecord As Integer
    
    With mshList
        If .TextMatrix(intRow, .Cols - 4) = "" And Val(.TextMatrix(.Row, .ColIndex("��¼״̬"))) = 1 Then
            If Not Check����(StrBillNo) Then
                MsgBox "��û��Ȩ��ɾ���ƿⵥ��", vbInformation, gstrSysName
                Exit Sub
            End If
        
            strTitle = "ҩƷ���쵥"
        ElseIf Val(.TextMatrix(.Row, .Cols - 3)) Mod 3 = 2 And mint�������� = 1 Then
            strTitle = "�������뵥"
        End If
        
        On Error GoTo ErrHandle
        intRow = .Row
        StrBillNo = .TextMatrix(intRow, 0)
        
        intReturn = MsgBox("��ȷʵҪɾ�����ݺ�Ϊ��" & StrBillNo & "����" & strTitle & "��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)
        intRecord = .Rows - 1
        If intReturn = vbYes Then
            gstrSQL = "zl_�����ƿ�_Delete('" & StrBillNo & "'," & Val(.TextMatrix(.Row, .ColIndex("��¼״̬"))) & " )"
            If gstrSQL = "" Then Exit Sub
            Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption & "-ɾ�����쵥")
            intRecord = intRecord - 1
            mlastRow = 0
            If .Rows > 2 Then
                .RemoveItem intRow
            ElseIf .Rows = 2 Then
                .Rows = 3
                .RemoveItem intRow
                With mshDetail
                    .Rows = 1
                    .Rows = 2
                    .FixedRows = 1
                    .Row = 1
                    .Col = 0
                    .ColSel = .Cols - 1
                End With
                SetEnable
                
            End If
                
            '.RowHeight(intRow) = 0
            If intRow < .Rows - 1 Then
                .Row = intRow
            Else
                If .Rows = 2 Then
                    .Row = 1
                Else
                    .Row = intRow - 1
                End If
            End If
            .Col = 0
'            .ColSel = .Cols - 1
            mshlist_EnterCell
        End If
    End With
    staThis.Panels(2).Text = "��ǰ����" & intRecord & "�ŵ���"
    Exit Sub

ErrHandle:
    If ErrCenter() = 1 Then Resume 'Resume����������õ���
    Call SaveErrLog
    
End Sub

Private Sub mnuEditDisplay_Click()
    '�鿴����
    
    Dim strNo As String
    With mshList
        strNo = .TextMatrix(.Row, 0)
        frmRequestStuffCard.ShowCard Me, strNo, 4, .TextMatrix(.Row, .ColIndex("��¼״̬")), mstrPrivs, , cboStock.ItemData(cboStock.ListIndex)
    End With
End Sub


'Modified By ���� 2003-12-10 ����������
Private Sub mnuEditDisReceive_Click()
    Dim strNo As String, blnSuccess As Boolean
    Dim int����ʽ As Integer
    
    If mnuEditDisReceive.Caption = "�������(&R)" Then
        int����ʽ = 1
    Else
        int����ʽ = 0
    End If
    
    With mshList
        strNo = .TextMatrix(.Row, 0)
        frmRequestStuffCard.ShowCard Me, strNo, 7, .TextMatrix(.Row, .ColIndex("��¼״̬")), mstrPrivs, blnSuccess, cboStock.ItemData(cboStock.ListIndex), int����ʽ
        If Not blnSuccess Then Exit Sub
    End With
    Call mnuViewRefresh_Click
End Sub

Private Sub mnuEditImport_Click()
    Dim blnSuccess As Boolean
    
    frmPurchaseImportFromPlane.ShowCard Me, cboStock.Text, cboStock.ItemData(cboStock.ListIndex), mintUnit, InStr(mstrPrivs, "���пⷿ") <> 0, blnSuccess, 1, 1722, mint��ȷ����
    
    If blnSuccess = True Then
        mnuViewRefresh_Click
    End If
End Sub

Private Sub mnuEditModify_Click()
    '�޸�
    Dim strNo As String
    Dim blnSuccess As Boolean
    
    blnSuccess = False
    With mshList
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        strNo = .TextMatrix(.Row, 0)
        If Not Check����(strNo) Then
            MsgBox "��û��Ȩ���޸��ƿⵥ��", vbInformation, gstrSysName
            Exit Sub
        End If
        
        frmRequestStuffCard.ShowCard Me, strNo, 2, mshList.TextMatrix(.Row, mshList.ColIndex("��¼״̬")), mstrPrivs, blnSuccess, cboStock.ItemData(cboStock.ListIndex)
        If blnSuccess = True Then
            mnuViewRefresh_Click
        End If
    End With
End Sub

Private Sub mnuEditReceive_Click()
    Dim strNo As String, blnSuccess As Boolean
    With mshList
        strNo = .TextMatrix(.Row, 0)
        frmRequestStuffCard.ShowCard Me, strNo, 6, .TextMatrix(.Row, .ColIndex("��¼״̬")), mstrPrivs, blnSuccess, cboStock.ItemData(cboStock.ListIndex)
    End With
    If blnSuccess = True Then
        mnuViewRefresh_Click
    End If
End Sub

Private Sub mnuFileBillPreview_Click()
    With mshList
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        ReportOpen gcnOracle, glngSys, "zl1_bill_1722", Me, "���ݱ��=" & .TextMatrix(.Row, 0), "��¼״̬=" & .TextMatrix(.Row, .ColIndex("��¼״̬")), "��λϵ��=" & mintUnit, 1
    End With
End Sub

Private Sub mnuFileBillPrint_Click()
    
    With mshList
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        ReportOpen gcnOracle, glngSys, "zl1_bill_1722", Me, "���ݱ��=" & .TextMatrix(.Row, 0), "��¼״̬=" & .TextMatrix(.Row, .ColIndex("��¼״̬")), "��λϵ��=" & mintUnit, 2
    End With
End Sub

Private Sub mnuFileExcel_Click()
    '�����Excel
    If Me.ActiveControl Is mshList Then
        mshList.Redraw = False
        subPrint 3
        mshList.Redraw = True
        mshList.Col = 0
'        mshList.ColSel = mshList.Cols - 1
    ElseIf Me.ActiveControl Is mshDetail Then
        mshDetail.Redraw = False
        subExcel 3
        mshDetail.Redraw = True
        mshDetail.Col = 0
        mshDetail.ColSel = mshDetail.Cols - 1
    End If
End Sub

Private Sub mnufileexit_Click()
    '�˳�
    Unload Me
    
End Sub

Private Sub mnuFileParameter_Click()
    Dim strReg As String
    '��������
'    frmRequestPara.���ò��� mlngModule, Me, mstrCaption, mstrPrivs
    frmParaset.���ò��� mlngModule, mstrPrivs, Me, mstrCaption
    
    strReg = Val(zlDatabase.GetPara("���ĵ�λ", glngSys, mlngModule, "0"))
    mintUnit = Val(strReg)
    mbln����˲� = IIf((zlDatabase.GetPara("������Ҫ�˲������ƿ�", glngSys, mlngModule, "0")) = 0, False, True)
    If mbln����˲� = False Then
        mnuEditCheck.Visible = False
        tlbTool.Buttons("Check").Visible = False
    Else
        mnuEditCheck.Visible = True
        tlbTool.Buttons("Check").Visible = True
    End If
    mintFindDay = Val(zlDatabase.GetPara("��ѯ����", glngSys, mlngModule, 1))
    mdtStartDate = Format(DateAdd("d", -mintFindDay, sys.Currentdate), "yyyy-MM-dd")
    mdtEndDate = Format(sys.Currentdate, "yyyy-MM-dd")
    
    Call GetList(mstrFind)
End Sub

Private Sub mnuFilePreView_Click()
    '��ӡԤ��
    mshList.Redraw = False
    subPrint 2
    mshList.Redraw = True
    mshList.Col = 0
'    mshList.ColSel = mshList.Cols - 1
    
End Sub

Private Sub mnuFilePrint_Click()
    '��ӡ
    mshList.Redraw = False
    subPrint 1
    mshList.Redraw = True
    mshList.Col = 0
'    mshList.ColSel = mshList.Cols - 1
End Sub

Private Sub mnuFilePrintSet_Click()
    '��ӡ����
    zlPrintSet
End Sub

Private Sub mnuHelpAbout_Click()
    '����
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpTitle_Click()
    '��������
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name)
End Sub

Private Sub mnuHelpWebHome_Click()
    '������ҳ
    Call zlHomePage(Me.hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    '���ͷ���
    Call zlMailTo(Me.hwnd)
End Sub

Private Sub mnuViewRefresh_Click()
    'ˢ��
    GetList mstrFind
End Sub

Private Sub mnuViewSearch_Click()
    
   '����
    
    Dim strCon As String
    Dim strFind As String
    Dim strOthers() As String
    
    strFind = FrmTransferSearch.GetSearch(Me, 1716, mdtStartDate, mdtEndDate, mdtVerifyStart, mdtVerifyEnd, mstrPrivs, strOthers)
    
    If strFind <> "" Then
        mstrFind = strFind
        mstrOthers = strOthers
        GetList mstrFind
        If Format(mdtStartDate, "yyyy-mm-dd") = "1901-01-01" And Format(mdtVerifyStart, "yyyy-mm-dd") = "1901-01-01" Then
            lblRange.Visible = False
        ElseIf Format(mdtStartDate, "yyyy-mm-dd") <> "1901-01-01" And Format(mdtVerifyStart, "yyyy-mm-dd") <> "1901-01-01" Then
            lblRange.Visible = True
            lblRange = "��ѯ��Χ:�������� " & Format(mdtStartDate, "yyyy��MM��dd��") & "��" & Format(mdtEndDate, "yyyy��MM��dd��") & "  ������� " & Format(mdtVerifyStart, "yyyy��MM��dd��") & "��" & Format(mdtVerifyEnd, "yyyy��MM��dd��")
        ElseIf Format(mdtStartDate, "yyyy-mm-dd") <> "1901-01-01" Then
            lblRange.Visible = True
            lblRange = "��ѯ��Χ:�������� " & Format(mdtStartDate, "yyyy��MM��dd��") & "��" & Format(mdtEndDate, "yyyy��MM��dd��")
        ElseIf Format(mdtVerifyStart, "yyyy-mm-dd") <> "1901-01-01" Then
            lblRange.Visible = True
            lblRange = "��ѯ��Χ:������� " & Format(mdtVerifyStart, "yyyy��MM��dd��") & "��" & Format(mdtVerifyEnd, "yyyy��MM��dd��")
        End If
    End If
    
End Sub

Private Sub mnuViewStatus_Click()
    With mnuViewStatus
        .Checked = Not .Checked
        staThis.Visible = .Checked
    End With
    
    Form_Resize
End Sub
Private Sub mnuReportItem_Click(Index As Integer)
    Dim strNo As String
    Dim intRecodeSta As Integer
    Dim lng�ⷿID As Long
    Dim lngCol As Long
    
    With mshList
        strNo = Trim(.TextMatrix(.Row, 0))
        lngCol = GetCol(mshList, "��¼״̬")
        If lngCol < 0 Then
            intRecodeSta = 1
        Else
            intRecodeSta = Val(.TextMatrix(.Row, lngCol))
        End If
    End With
    
    If cboStock.ListIndex < 0 Then
        lng�ⷿID = 0
    Else
        lng�ⷿID = cboStock.ItemData(cboStock.ListIndex)
    End If
    
    '2006-04-25:���˺�:�����Զ��屨������ģ��Ĺ���
    If Format(mdtStartDate, "yyyy-mm-dd") = "1990-01-01" Then
        Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, "NO=" & strNo, "��¼״̬=" & intRecodeSta, "���첿��=" & lng�ⷿID)
    Else
        Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, "NO=" & strNo, "��¼״̬=" & intRecodeSta, "���첿��=" & lng�ⷿID, "��ʼʱ��=" & Format(mdtStartDate, "yyyy-mm-dd"), "����ʱ��=" & Format(mdtEndDate, "yyyy-mm-dd"))
    End If
End Sub
Private Sub mnuViewToolButton_Click()
    With mnuViewToolButton
        .Checked = Not .Checked
        cbrTool.Visible = .Checked
        mnuViewToolText.Enabled = .Checked
    End With
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim intCount As Integer      '����������
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    With tlbTool.Buttons
        If mnuViewToolText.Checked = False Then
            'ȡ�����е��ı���ǩ��ʾ
            For intCount = 1 To .Count
                .Item(intCount).Caption = ""
            Next
        Else
            '�����е��ı���ǩ��ʾ��˵����Tag�зŵ��ı���ǩ
            For intCount = 1 To .Count
                .Item(intCount).Caption = .Item(intCount).Tag
            Next
        End If
    End With
    
    cbrTool.Bands(1).MinHeight = tlbTool.Height
    
    Form_Resize
End Sub

Private Sub mshDetail_Click()
    With mshDetail
        If .Row < 1 Or .TextMatrix(.Row, 0) = "" Then Exit Sub
        If .MouseRow = 0 Then
            DetailSort          '������
            Exit Sub
        End If
    End With
End Sub

Private Sub mshList_Click()
    With mshList
         If .Row < 1 Then Exit Sub
         If .MouseRow = 0 Then
            ListSort
            Exit Sub
         End If
    End With
End Sub

Private Sub mshlist_DblClick()
    If mnuEditModify.Visible = False Then Exit Sub
    If mnuEditModify.Enabled = False Then Exit Sub
    If mshList.MouseRow = 0 Then Exit Sub
    mnuEditModify_Click
End Sub

Private Sub mshlist_EnterCell()
    Dim rsDetail As New Recordset
    Dim strUnitQuantity As String               '��λ��������ʽ����
    Dim IntBill As Integer                      '��������  �磺1���⹺��⣻2��
    Dim strUnit As String                       '��λ����:�����ﵥλ��סԺ��λ��
    Dim str��װϵ�� As String
    Dim strOrder As String
    Dim strCompare As String
    
    On Error GoTo ErrHandle

    If mlastRow = mshList.Row Or LTrim(mshList.TextMatrix(mshList.Row, 0)) = "" Then
        If LTrim(mshList.TextMatrix(mshList.Row, 0)) = "" Then
            With mshDetail
                .Cols = IIf(gblnCode = True, 21, 19)
                .Rows = 2
                
                .Clear
                .TextMatrix(0, 0) = "���"
                .TextMatrix(0, 1) = "������Ϣ"
                .TextMatrix(0, 2) = "������Դ"
                .TextMatrix(0, 3) = "���"
                .TextMatrix(0, 4) = "����"
                .TextMatrix(0, 5) = "��׼�ĺ�"
                .TextMatrix(0, 6) = "����"
                .TextMatrix(0, 7) = "Ч��"
                .TextMatrix(0, 8) = "���Ч��"
                .TextMatrix(0, 9) = "��д����"
                .TextMatrix(0, 10) = "ʵ������"
                .TextMatrix(0, 11) = "��λ"
                .TextMatrix(0, 12) = "�ɱ���"
                .TextMatrix(0, 13) = "�ɱ����"
                .TextMatrix(0, 14) = "�ۼ�"
                .TextMatrix(0, 15) = "�ۼ۽��"
                .TextMatrix(0, 16) = "���"
                .TextMatrix(0, 17) = "�ⷿ��λ"
                .TextMatrix(0, 18) = "����"
                
                If gblnCode = True Then
                    .TextMatrix(0, 19) = "��Ʒ����"
                    .TextMatrix(0, 20) = "�ڲ�����"
                End If
                
                .Row = 1
                .Col = 0
                .ColSel = .Cols - 1
            End With
        End If
        Exit Sub
    End If
    mlastRow = mshList.Row
    SetEnable
    
    strOrder = zlDatabase.GetPara("��������", glngSys, mlngModule, "00")
    strCompare = Mid(strOrder, 1, 1)
    
    If mshList.Row >= 1 And LTrim(mshList.TextMatrix(mshList.Row, 0)) <> "" Then
        mshList.Col = 0
'        mshList.ColSel = mshList.Cols - 1
        
        mshDetail.Redraw = False
        
        Select Case mintUnit
            Case 0
                str��װϵ�� = "1"
            Case Else
                str��װϵ�� = "B.����ϵ��"
        End Select
            
        gstrSQL = "" & _
            "   SELECT * " & _
            "   FROM (  " & _
            "           SELECT DISTINCT ���,('['||D.����||']'||D.����) AS ������Ϣ,B.������Դ," & _
            "                       D.���,A.����,A.��׼�ĺ�, A.����, A.Ч��,to_char(a.���Ч��,'yyyy-mm-dd') as ���Ч��," & _
            "                       (TO_CHAR(A.��д���� /" & str��װϵ�� & "," & mOraFMT.FM_���� & ")) AS ��д����," & _
            "                       (TO_CHAR(A.ʵ������ /" & str��װϵ�� & "," & mOraFMT.FM_���� & ")) AS ʵ������," & _
                                    IIf(mintUnit = 0, "D.���㵥λ", "b.��װ��λ") & "  AS ��λ," & _
            "                       TO_CHAR (A.�ɱ���*" & str��װϵ�� & "," & mOraFMT.FM_�ɱ��� & ") AS �ɱ���," & _
            "                       TO_CHAR (A.�ɱ����, " & mOraFMT.FM_��� & ") AS �ɱ����," & _
            "                       TO_CHAR (A.���ۼ�*" & str��װϵ�� & "," & mOraFMT.FM_���ۼ� & ") AS �ۼ�," & _
            "                       TO_CHAR (A.���۽��, " & mOraFMT.FM_��� & ") AS �ۼ۽��," & _
            "                       TO_CHAR (A.���," & mOraFMT.FM_��� & ") AS ��� ,C.�ⷿ��λ ,NVL(E.����,D.����) as ���� "
            
        If gblnCode = True Then
            gstrSQL = gstrSQL & " ,A.��Ʒ����,A.�ڲ����� "
        End If
        
        gstrSQL = gstrSQL & _
            "           FROM ҩƷ�շ���¼ A, �������� B, �շ���Ŀ���� E, �շ���ĿĿ¼ D, ���ϴ����޶� C " & _
            "           WHERE A.ҩƷID = B.����ID AND B.����ID=D.ID " & _
            "                   AND B.����ID = E.�շ�ϸĿID(+) AND E.����(+)=3 " & _
            "                   AND A.��¼״̬ = [2]" & _
            "                   AND A.���� = 19 AND ���ϵ��=1 " & _
            "                   AND A.NO =[1]  AND A.ҩƷID=C.����ID(+) AND A.�ⷿID=C.�ⷿID(+)" & _
            "   )" & _
            " ORDER BY " & IIf(strCompare = "0", "���", IIf(strCompare = "1", "������Ϣ", IIf(strCompare = "2", "����", "�ⷿ��λ"))) & IIf(Right(strOrder, 1) = "0", " ASC", " DESC")
        
        Set rsDetail = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, mshList.TextMatrix(mshList.Row, 0), Val(mshList.TextMatrix(mshList.Row, mshList.ColIndex("��¼״̬"))))
                
        Set mshDetail.Recordset = rsDetail
    
        
        With rsDetail
            .Close
        End With
        With mshDetail
            If .Rows = 1 Then
                .Rows = .Rows + 100
                .Row = 1
                .Redraw = True
                .TopRow = 1
                .Rows = .Rows - 99
            End If
            .Row = 1
            .Col = 0
            .ColSel = .Cols - 1
        End With
        mshDetail.Redraw = True
    End If
    SetDetailColWidth
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mshlist_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If mnuEditModify.Visible = False Then Exit Sub
        If mnuEditModify.Enabled = False Then Exit Sub
        mnuEditModify_Click
    End If
        
End Sub

Private Sub mshlist_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button <> 2 Then Exit Sub
    If mnuEdit.Visible = False Then Exit Sub
    
    PopupMenu mnuEdit, 2
    
End Sub

Private Sub picSeparate_s_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    '�ָ�������
    
    If Button <> 1 Then Exit Sub
    
    With picSeparate_s
        If .Top + Y < 2000 Then Exit Sub
        If .Top + Y > ScaleHeight - 2000 Then Exit Sub
        .Move .Left, .Top + Y
    End With
    
    With mshList
        .Top = IIf(cbrTool.Visible, cbrTool.Height, 0)
        .Height = picSeparate_s.Top - .Top
    End With
    
    With mshDetail
        .Top = picSeparate_s.Top + picSeparate_s.Height + 100
        .Height = ScaleHeight - .Top - IIf(staThis.Visible, staThis.Height, 0)
    End With
    
End Sub

Private Sub tlbTool_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "PrintView"
            mnuFilePreView_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Add"
            mnuEditAddHank_Click
        Case "Modify"
            mnuEditModify_Click
        Case "Check"
            mnuEditCheck_Click
        Case "Delete"
            mnuEditDel_Click
        Case "Receive"
            mnuEditReceive_Click
        Case "DisReceive"
            mnuEditDisReceive_Click
        Case "Search"
            mnuViewSearch_Click
        Case "Refresh"
            mnuViewRefresh_Click
        Case "Help"
            mnuHelpTitle_Click
        Case "Exit"
            mnufileexit_Click
    End Select
End Sub

'���ò˵��͹��߰�ť�Ŀ�������
Private Sub SetEnable()
    Dim bln�ѷ��� As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    With mshList
        .ToolTipText = ""
        If .TextMatrix(.Row, 0) = "" Or .Row = 0 Then         'û�е�
            mnuFilePreView.Enabled = False
            mnuFilePrint.Enabled = False
            mnuFileBillPreview.Enabled = False
            mnuFileBillPrint.Enabled = False
            mnuFileExcel.Enabled = False
            tlbTool.Buttons("Print").Enabled = False
            tlbTool.Buttons("PrintView").Enabled = False
            
            If mnuEditCheck.Visible = True Then
                mnuEditCheck.Enabled = False
                tlbTool.Buttons("Check").Enabled = False
            End If
            If mnuEditModify.Visible = True Then
                mnuEditModify.Enabled = False
                tlbTool.Buttons("Modify").Enabled = False
            End If
            If mnuEditDel.Visible = True Then
                mnuEditDel.Enabled = False
                tlbTool.Buttons("Delete").Enabled = False
            End If
            If mnuEditDisplay.Visible = True Then
                mnuEditDisplay.Enabled = False
            End If
            If mnuEditReceive.Visible Then
                mnuEditReceive.Enabled = False
                tlbTool.Buttons("Receive").Enabled = False
            End If
            If mnuEditDisReceive.Visible Then
                mnuEditDisReceive.Enabled = False
                tlbTool.Buttons("DisReceive").Enabled = False
            End If
         Else
            mnuFilePreView.Enabled = True
            mnuFilePrint.Enabled = True
            mnuFileBillPreview.Enabled = True
            mnuFileBillPrint.Enabled = True
            mnuFileExcel.Enabled = True
            tlbTool.Buttons("Print").Enabled = True
            tlbTool.Buttons("PrintView").Enabled = True
                        
            If .TextMatrix(.Row, .ColIndex("�������")) = "" Then    'δ��˵�
                bln�ѷ��� = (mshList.TextMatrix(mshList.Row, .ColIndex("������")) <> "")
                If mnuEditModify.Visible = True Then
                    mnuEditModify.Enabled = True
                    tlbTool.Buttons("Modify").Enabled = True
                End If
                If mnuEditDel.Visible = True Then
                    mnuEditDel.Enabled = True
                    tlbTool.Buttons("Delete").Enabled = True
                End If
                 
                If mnuEditDisplay.Visible = True Then
                    mnuEditDisplay.Enabled = True
                End If
                
                If mnuEditModify.Visible = True Then
                    mnuEditModify.Enabled = True
                    tlbTool.Buttons("Modify").Enabled = True
                End If
                
                '���Ҫ���к˲�
                If mnuEditCheck.Visible = True Then
                    mnuEditCheck.Enabled = Not bln�ѷ���
                    tlbTool.Buttons("Check").Enabled = Not bln�ѷ���
                    If .TextMatrix(.Row, .ColIndex("�˲�����")) = "" Then    '�˲�����
                        'δ�˲�
                        If mnuEditReceive.Visible Then
                            mnuEditReceive.Enabled = bln�ѷ���  'δ�˲��ѷ��͵Ŀ��Խ��ܣ��ƿ�ģ�����
                            tlbTool.Buttons("Receive").Enabled = bln�ѷ���
                        End If
                    Else
                        '�Ѻ˲�
                        If mnuEditReceive.Visible Then
                            mnuEditReceive.Enabled = bln�ѷ���
                            tlbTool.Buttons("Receive").Enabled = bln�ѷ���
                        End If
                    End If
                Else
                '�����к˲�
                    If mnuEditReceive.Visible Then
                        mnuEditReceive.Enabled = bln�ѷ���
                        tlbTool.Buttons("Receive").Enabled = bln�ѷ���
                    End If
                End If

                If mnuEditDisReceive.Visible Then
                    If bln�ѷ��� Then
                        mnuEditDisReceive.Enabled = Not bln�ѷ���
                        tlbTool.Buttons("DisReceive").Enabled = Not bln�ѷ���
                    Else
                        mnuEditDisReceive.Enabled = False
                        tlbTool.Buttons("DisReceive").Enabled = False
                    End If
                End If
                
                '�����������δ��ˣ�������ɾ��
                If mint�������� = 1 Then
                    If Val(.TextMatrix(.Row, .ColIndex("��¼״̬"))) Mod 3 = 2 Then
                        mnuEditModify.Enabled = False
                        tlbTool.Buttons("Modify").Enabled = False
                        mnuEditReceive.Enabled = False
                        tlbTool.Buttons("Receive").Enabled = False
                        mnuEditDisReceive.Enabled = False
                        tlbTool.Buttons("DisReceive").Enabled = False
                        
                        mnuEditDel.Enabled = True
                        tlbTool.Buttons("Delete").Enabled = True
                    End If
                Else
                    If mnuEditDisReceive.Visible Then
                        If bln�ѷ��� Then
                            mnuEditDisReceive.Enabled = Not bln�ѷ���
                            tlbTool.Buttons("DisReceive").Enabled = Not bln�ѷ���
                        Else
                            mnuEditDisReceive.Enabled = False
                            tlbTool.Buttons("DisReceive").Enabled = False
                        End If
                    End If
                End If
            ElseIf .TextMatrix(.Row, .ColIndex("��¼״̬")) = 1 Then    '��˵�
                '�ж��Ƿ���ܣ���֧���ѳ������ݵĽ��ܹ��ܣ�����ȫ�˻��为���ķ�ʽ�������ΪҪʵ��������ܣ���Ҫ����ͳ��ʣ��������
                If mnuEditCheck.Visible = True Then
                    mnuEditCheck.Enabled = False
                    tlbTool.Buttons("Check").Enabled = False
                End If
                    
                If mnuEditModify.Visible = True Then
                    mnuEditModify.Enabled = False
                    tlbTool.Buttons("Modify").Enabled = False
                End If
                
                If mnuEditDel.Visible = True Then
                    mnuEditDel.Enabled = False
                    tlbTool.Buttons("Delete").Enabled = False
                End If
                 
                If mnuEditDisplay.Visible = True Then
                    mnuEditDisplay.Enabled = True
                End If
                If mnuEditReceive.Visible Then
                    mnuEditReceive.Enabled = False
                    tlbTool.Buttons("Receive").Enabled = False
                End If
                If mnuEditDisReceive.Visible Then
                    mnuEditDisReceive.Enabled = True
                    tlbTool.Buttons("DisReceive").Enabled = True
                End If
            Else   '2,3 ������
                If .TextMatrix(.Row, .ColIndex("��¼״̬")) Mod 3 = 0 Then
                    .ToolTipText = "�������ݵ�ԭ����"
                    If mnuEditDisReceive.Visible = True Then
                        mnuEditDisReceive.Enabled = True
                        tlbTool.Buttons("DisReceive").Enabled = True
                    End If
                ElseIf .TextMatrix(.Row, .ColIndex("��¼״̬")) Mod 3 = 2 Then
                    .ToolTipText = "��������"
                    If mnuEditDisReceive.Visible = True Then
                        mnuEditDisReceive.Enabled = False
                        tlbTool.Buttons("DisReceive").Enabled = False
                    End If
                End If
                If mnuEditCheck.Visible = True Then
                    mnuEditCheck.Enabled = False
                    tlbTool.Buttons("Check").Enabled = False
                End If
                If mnuEditModify.Visible = True Then
                    mnuEditModify.Enabled = False
                    tlbTool.Buttons("Modify").Enabled = False
                End If
                If mnuEditDel.Visible = True Then
                    mnuEditDel.Enabled = False
                    tlbTool.Buttons("Delete").Enabled = False
                End If
                 
                If mnuEditDisplay.Visible = True Then
                    mnuEditDisplay.Enabled = True
                End If
                If mnuEditReceive.Visible Then
                    mnuEditReceive.Enabled = False
                    tlbTool.Buttons("Receive").Enabled = False
                End If
            End If
        End If
        
    End With
End Sub

Private Sub subPrint(bytMode As Byte)
'����:���д�ӡ,Ԥ���������EXCEL
'����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
'    If gstrUserName = "" Then Call GetUserInfo
    Dim objPrint As Object
    Dim objRow As New zlTabAppRow
    Dim strRange As String
    
    If Format(mdtStartDate, "yyyy-mm-dd") = "1901-01-01" And Format(mdtVerifyStart, "yyyy-mm-dd") = "1901-01-01" Then
        strRange = "������� " & Format(mdtVerifyStart, "yyyy��MM��dd��") & "��" & Format(mdtVerifyEnd, "yyyy��MM��dd��")
    ElseIf Format(mdtVerifyStart, "yyyy-mm-dd") <> "1901-01-01" Then
        strRange = "�������� " & Format(mdtStartDate, "yyyy��MM��dd��") & "��" & Format(mdtEndDate, "yyyy��MM��dd��") & "  ������� " & Format(mdtVerifyStart, "yyyy��MM��dd��") & "��" & Format(mdtVerifyEnd, "yyyy��MM��dd��")
    Else
        strRange = "�������� " & Format(mdtStartDate, "yyyy��MM��dd��") & "��" & Format(mdtEndDate, "yyyy��MM��dd��")
    End If
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "����_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    
    objPrint.Title.Text = mstrCaption
        
    objRow.Add "ʱ�䣺" & strRange
    
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
        
    objRow.Add "��ӡ��:" & gstrUserName
    objRow.Add "��ӡ����:" & Format(sys.Currentdate, "yyyy��MM��dd��")
    objPrint.BelowAppRows.Add objRow
    
    Set objPrint.Body = mshList
    
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

Private Sub tlbTool_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    If ButtonMenu.Index = 1 Then
        mnuEditAddHank_Click
    Else
        mnuEditAddAuto_Click
    End If
End Sub

Private Sub tlbTool_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuViewTool
    End If
End Sub


'�Ե���ͷ������
Private Sub ListSort()
    Dim intCol As Integer
    Dim intRow As Integer
    Dim intTemp As String
    
    With mshList
        If .Rows > 1 Then
            .Redraw = False
            intCol = .MouseCol
            .Col = intCol
            .ColSel = intCol
            intTemp = .TextMatrix(.Row, 0)
            
            Select Case intCol
                Case 2
                    If intCol = mintPreCol And mintsort = flexSortNumericDescending Then
                       .Sort = flexSortNumericAscending
                       mintsort = flexSortNumericAscending
                    Else
                       .Sort = flexSortNumericDescending
                       mintsort = flexSortNumericDescending
                    End If
                Case Else
                    If intCol = mintPreCol And mintsort = flexSortStringNoCaseDescending Then
                       .Sort = flexSortStringNoCaseAscending
                       mintsort = flexSortStringNoCaseAscending
                    Else
                       .Sort = flexSortStringNoCaseDescending
                       mintsort = flexSortStringNoCaseDescending
                    End If
            End Select
            mintPreCol = intCol
            .Row = grid.MshGrdFindRow(mshList, intTemp, 0)
            If .RowPos(.Row) + .RowHeight(.Row) > .Height Then
                .TopRow = .Row
            Else
                .TopRow = 1
            End If
            .Col = 0
            .ColSel = .Cols - 1
            .Redraw = True
            .SetFocus
        Else
            .ColSel = 0
        End If
    End With
End Sub

'�Ե���ͷ������
Private Sub DetailSort()
    Dim intCol As Integer
    Dim intRow As Integer
    Dim intTemp As Integer
    
    With mshDetail
        If .Rows > 1 Then
            .Redraw = False
            intCol = .MouseCol
            .Col = intCol
            .ColSel = intCol
            intTemp = .TextMatrix(.Row, 0)
                
            Select Case intCol
                Case 6, 7, 9, 10, 11, 12, 13
                    If intCol = mintPreDetailCol And mintDetailsort = flexSortNumericDescending Then
                       .Sort = flexSortNumericAscending
                       mintDetailsort = flexSortNumericAscending
                    Else
                       .Sort = flexSortNumericDescending
                       mintDetailsort = flexSortNumericDescending
                    End If
                    
                Case Else
                    If intCol = mintPreDetailCol And mintDetailsort = flexSortStringNoCaseDescending Then
                       .Sort = flexSortStringNoCaseAscending
                       mintDetailsort = flexSortStringNoCaseAscending
                    Else
                       .Sort = flexSortStringNoCaseDescending
                       mintDetailsort = flexSortStringNoCaseDescending
                    End If
            End Select
                
            mintPreDetailCol = intCol
            .Row = grid.MshGrdFindRow(mshDetail, intTemp, 0)
            If .RowPos(.Row) + .RowHeight(.Row) > .Height Then
                .TopRow = .Row
            Else
                .TopRow = 1
            End If
            .Col = 0
            .ColSel = .Cols - 1
            .Redraw = True
            .SetFocus
        Else
            .ColSel = 0
        End If
    End With
End Sub

Private Sub subExcel(bytMode As Byte)
'����:���д�ӡ,Ԥ���������EXCEL
'����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
'    If gstrUserName = "" Then Call GetUserInfo
    Dim objPrint As Object
    Dim objRow As zlTabAppRow
    Dim strRange As String
    
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "����_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    
    objPrint.Title.Text = mstrCaption
        
    Set objRow = New zlTabAppRow
    objRow.Add ""
    objRow.Add "NO." & Trim(mshList.TextMatrix(mshList.Row, mshList.ColIndex("no")))
    objPrint.UnderAppRows.Add objRow

    Set objRow = New zlTabAppRow
    objRow.Add "�Ƴ��ⷿ��" & mshList.TextMatrix(mshList.Row, mshList.ColIndex("���Ͽⷿ"))
    objRow.Add "����ⷿ��" & gstrDeptName
    objPrint.UnderAppRows.Add objRow
        
    Set objRow = New zlTabAppRow
    objRow.Add "ժҪ:" & mshList.TextMatrix(mshList.Row, mshList.ColIndex("ժҪ"))
    objPrint.BelowAppRows.Add objRow
        
    Set objRow = New zlTabAppRow
    objRow.Add "������:" & mshList.TextMatrix(mshList.Row, mshList.ColIndex("������")) & "  ��������:" & mshList.TextMatrix(mshList.Row, mshList.ColIndex("��������"))
    
    objRow.Add "�����:" & mshList.TextMatrix(mshList.Row, mshList.ColIndex("�����")) & "  �������:" & mshList.TextMatrix(mshList.Row, mshList.ColIndex("�������"))
    objPrint.BelowAppRows.Add objRow
    
    Set objPrint.Body = mshDetail
    zlPrintOrView1Grd objPrint, bytMode
End Sub

Private Function Check����(ByVal StrBillNo As String) As Boolean
    Dim rsCheck As New ADODB.Recordset
    On Error GoTo ErrHandle
    
    '�ȼ���ǲ������쵥
    gstrSQL = " Select Nvl(��ҩ��ʽ,0) ���� From ҩƷ�շ���¼ " & _
              " Where ����=19 And NO=[1] And ���=1"
    
    Set rsCheck = zlDatabase.OpenSQLRecord(gstrSQL, "����ǲ������쵥", StrBillNo)
              
    Check���� = Not (rsCheck!���� = 0)
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

