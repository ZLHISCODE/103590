VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCheckMain 
   ClientHeight    =   4980
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9480
   Icon            =   "frmCheckMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin TabDlg.SSTab TabShow 
      Height          =   360
      Left            =   30
      TabIndex        =   6
      Top             =   720
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   635
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "�̵��¼���嵥(&1)"
      TabPicture(0)   =   "frmCheckMain.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "�̵���嵥(&2)"
      TabPicture(1)   =   "frmCheckMain.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
   End
   Begin MSComctlLib.ImageList ilsHot 
      Left            =   6270
      Top             =   120
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
            Picture         =   "frmCheckMain.frx":0342
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":0562
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":0782
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":099E
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":0BBE
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":0DDE
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":0FFA
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":1216
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":1430
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":158A
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":17AA
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsCold 
      Left            =   5670
      Top             =   120
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
            Picture         =   "frmCheckMain.frx":19CA
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":1BEA
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":1E0A
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":2026
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":2246
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":2466
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":2682
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":289E
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":2AB8
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":2C12
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":2E2E
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Cmd���� 
      Caption         =   "����(&V)"
      Height          =   350
      Left            =   5160
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2550
      Width           =   1100
   End
   Begin VB.PictureBox picSeparate_s 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   30
      MousePointer    =   7  'Size N S
      ScaleHeight     =   300
      ScaleWidth      =   4815
      TabIndex        =   4
      Top             =   2580
      Width           =   4815
      Begin VB.Label lbl�ɱ����� 
         AutoSize        =   -1  'True
         Caption         =   "�ɱ�����ϼƣ�"
         Height          =   170
         Left            =   3240
         TabIndex        =   8
         Top             =   40
         Width           =   1440
      End
      Begin VB.Label lblSum�ɱ���� 
         AutoSize        =   -1  'True
         Caption         =   "�̵�ɱ����ϼƣ�"
         Height          =   170
         Left            =   480
         TabIndex        =   7
         Top             =   40
         Width           =   1620
      End
   End
   Begin ComCtl3.CoolBar cbrTool 
      Height          =   780
      Left            =   0
      TabIndex        =   0
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
      Caption2        =   "�ⷿ"
      Child2          =   "cboStock"
      MinWidth2       =   3000
      MinHeight2      =   300
      Width2          =   3345
      NewRow2         =   0   'False
      Begin VB.ComboBox cboStock 
         Height          =   300
         Left            =   8685
         TabIndex        =   2
         Text            =   "cboStock"
         Top             =   240
         Width           =   3000
      End
      Begin MSComctlLib.Toolbar tlbTool 
         Height          =   720
         Left            =   165
         TabIndex        =   1
         Top             =   30
         Width           =   7875
         _ExtentX        =   13891
         _ExtentY        =   1270
         ButtonWidth     =   1138
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
               Caption         =   "��¼��"
               Key             =   "Bill"
               Description     =   "����"
               Object.ToolTipText     =   "��¼��"
               Object.Tag             =   "��¼��"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�̵��"
               Key             =   "Table"
               Object.ToolTipText     =   "�̵��"
               Object.Tag             =   "�̵��"
               ImageIndex      =   3
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   3
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Auto"
                     Text            =   "�Զ������̵��"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Total"
                     Text            =   "���ܼ�¼�������̵��"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Zero"
                     Text            =   "ȫ����Ϊ��"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�޸�"
               Key             =   "Modify"
               Description     =   "�޸�"
               Object.ToolTipText     =   "�޸�"
               Object.Tag             =   "�޸�"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ɾ��"
               Key             =   "Delete"
               Description     =   "ɾ��"
               Object.ToolTipText     =   "ɾ��"
               Object.Tag             =   "ɾ��"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "EditSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "���"
               Key             =   "Verify"
               Description     =   "���"
               Object.ToolTipText     =   "���"
               Object.Tag             =   "���"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Strike"
               Description     =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "VerifySeparate"
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
         Begin VB.Timer LimitTime 
            Enabled         =   0   'False
            Interval        =   8000
            Left            =   6660
            Top             =   180
         End
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   4620
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11642
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
   Begin VSFlex8Ctl.VSFlexGrid mshlist 
      Height          =   1005
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   6255
      _cx             =   11033
      _cy             =   1773
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
      ForeColorSel    =   -2147483640
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
      FormatString    =   $"frmCheckMain.frx":304E
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
   Begin VSFlex8Ctl.VSFlexGrid mshDetail 
      Height          =   975
      Left            =   360
      TabIndex        =   10
      Top             =   3120
      Width           =   5655
      _cx             =   9975
      _cy             =   1720
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
      BackColorSel    =   16053482
      ForeColorSel    =   -2147483640
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
      FormatString    =   $"frmCheckMain.frx":30C3
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
      Begin VB.Menu mnuEditAddBill 
         Caption         =   "���Ӽ�¼��(&B)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditAddTable 
         Caption         =   "�����̵��(&T)"
         Begin VB.Menu mnuEditAddTableAuto 
            Caption         =   "�Զ������̵��(&A)"
         End
         Begin VB.Menu mnuEditAddTableTotal 
            Caption         =   "���ܼ�¼�������̵��(&T)"
         End
         Begin VB.Menu mnuEditAddTableZero 
            Caption         =   "ȫ����Ϊ��(&Z)"
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
      Begin VB.Menu mnuEditVerify 
         Caption         =   "���(&C)"
      End
      Begin VB.Menu mnuEditStrike 
         Caption         =   "����(&K)"
      End
      Begin VB.Menu mnuEditLine2 
         Caption         =   "-"
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
Attribute VB_Name = "frmCheckMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngMode As Long
Private mstrFind As String
Private mblnBootUp As Boolean
Private mlastRow As Long                '�ϴε������
Private mstrTitle As String             '����ı���
Private mintPreCol As Integer           'ǰһ�ε���ͷ��������
Private mintsort As Integer             'ǰһ�ε���ͷ������
Private mintPreDetailCol As Integer     'ǰһ�ε������������
Private mintDetailsort As Integer       'ǰһ�ε����������
Public mstrPrivs As String                     'Ȩ��
Private mintUnit  As Integer                '��ʾ��λ:0-ɢװ��λ,1-��װ��λ
Private mintUnit1  As Integer                '��ʾ��λ:0-ɢװ��λ,1-��װ��λ
Private mstrOrder As String             '��¼����ʽ

Private mblnLoadGrid As Boolean

Private mintOldY  As Integer
Private mstrOthers() As String ' 0-��¼״̬(�ƻ�����),1-��ʼ����,2-��������,3-����id,4-�Է�����id(��������id����Ʒ���(�ƻ���)),5-������,6-�����,7-��Ӧ��ID,8-������,9-��ʼ��������,10-������������,11-��ʼ��Ʊ��,12-������Ʊ��
Private mblnCostView As Boolean             '�鿴�ɱ��������Ϣ true-����鿴 false-������鿴
Private Const mstrCaption As String = "�����̵����"
'�̵㵥
Private Const M_INT_COL�̵㵥NO As Integer = 0
Private Const M_INT_COL�̵㵥�̵�ʱ�� As Integer = 1
Private Const M_INT_COL�̵㵥������ As Integer = 2
Private Const M_INT_COL�̵㵥�������� As Integer = 3
Private Const M_INT_COL�̵㵥ժҪ As Integer = 4
Private Const M_INT_COL�̵㵥��� As Integer = 5
Private Const M_INT_COL�̵㵥���� As Integer = 6
Private Const M_INT_�̵㵥ALLCOLUMN As Integer = 7 '������
'�̵��
Private Const M_INT_COLNO As Integer = 0 ' "NO"
Private Const M_INT_COL�̵�ʱ�� As Integer = 1 ' "�̵�ʱ��"
Private Const M_INT_COL������ As Integer = 2 ' "������"
Private Const M_INT_COL�������� As Integer = 3 ' "��������"
Private Const M_INT_COL����� As Integer = 4 '"�����"
Private Const M_INT_COL������� As Integer = 5 '"�������"
Private Const M_INT_COL�̵��� As Integer = 6 '"�̵���"
Private Const M_INT_COL���� As Integer = 7 '"����"
Private Const M_INT_COL�̵�ɱ���� As Integer = 8 ' "�̵�ɱ����"
Private Const M_INT_COL�̵�ɱ����� As Integer = 9 ' "�̵�ɱ�����"
Private Const M_INT_COL��¼״̬ As Integer = 10 '"��¼״̬"
Private Const M_INT_COLժҪ As Integer = 11 '"ժҪ"
Private Const M_INT_ALLCOLUMN As Integer = 12 '������
 
'---------------------------------------------------------------------------------------------------------
'������صĹ�������:2008-08-22 16:35:52
'���˺�:
Private mblnNoClick As Boolean
Private mstr�������� As String
Private mbln����Ա���� As Boolean

'----------------------------------------------------------------------------------------------------------
'���˺�:����С��λ���ĸ�ʽ��
'�޸�:2007/03/06
Private mOraFMT As g_FmtString
Private mORaFMT��¼�� As g_FmtString
'----------------------------------------------------------------------------------------------------------


'��������
Private mdtStartDate As Date
Private mdtEndDate As Date
Private mdtVerifyStart As Date
Private mdtVerifyEnd As Date
Private mintFindDay As Integer  '��ѯ����

Private Sub cboStock_Click()
    If mblnNoClick Then Exit Sub
    If cboStock.ListIndex >= 0 Then cboStock.Tag = cboStock.ItemData(cboStock.ListIndex)
    If mblnBootUp Then mnuViewRefresh_Click
End Sub

Private Sub cboStock_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If cboStock.ListCount = 0 Then Call zlControl.ControlSetFocus(mshlist): Exit Sub
    
    If cboStock.ListIndex >= 0 Then
        If Val(cboStock.Tag) = cboStock.ItemData(cboStock.ListIndex) Then
            Call zlControl.ControlSetFocus(mshlist, True)
            Exit Sub
        End If
    End If
    
    If Select����ѡ����(Me, cboStock, Trim(cboStock.Text), mstr��������, mbln����Ա����) = False Then
        Exit Sub
    End If
    If cboStock.ListIndex >= 0 Then
        cboStock.Tag = cboStock.ItemData(cboStock.ListIndex)
    End If
End Sub

Private Sub cboStock_LostFocus()
    Dim i As Long
    If cboStock.ListCount = 0 Then Exit Sub
    If cboStock.ListIndex < 0 Then
        For i = 0 To cboStock.ListCount - 1
            If Val(cboStock.Tag) = cboStock.ItemData(i) Then
                mblnNoClick = True
                cboStock.ListIndex = i: Exit For
            End If
        Next
    End If
    mblnNoClick = False
End Sub

Private Sub cbrTool_Resize()
    If mblnBootUp = False Then Exit Sub
    Form_Resize
End Sub

Public Sub ShowList(ByVal lngMode As Long, ByVal strTitle As String, ByVal FrmMain As Variant)
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ʾָ���ĵ��ݹ���,
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------

    Dim strFind As String
    mblnBootUp = False
    mlngMode = lngMode
    mstrTitle = strTitle
    mstrPrivs = gstrPrivs
    
    If Not CheckDepend Then Exit Sub            '���������Բ���
    
    
    Me.Caption = strTitle
    
    SetVisable  '����Ȩ�����ò�ͬ����ʾ��Ŀ
    
    Call initGrid
    mintFindDay = Val(zlDataBase.GetPara("��ѯ����", glngSys, mlngMode, 1))
    mdtStartDate = Format(DateAdd("d", -mintFindDay, sys.Currentdate), "yyyy-MM-dd")
    mdtEndDate = Format(sys.Currentdate, "yyyy-MM-dd")
    mdtVerifyStart = "1901-01-01"
    mdtVerifyEnd = "1901-01-01"
    
    strFind = " AND A.��¼״̬ = 1 And A.������� is Null And A.�������� Between To_Date('" & Format(mdtStartDate, "yyyy-mm-dd") & " 00:00:00','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(mdtEndDate, "yyyy-mm-dd") & " 23:59:59','YYYY-MM-DD HH24:MI:SS')"
    mstrFind = strFind
    
    GetList (mstrFind)  '�г�����ͷ
    
    TabShow.Tab = 1
    Call SetListColWidth
    TabShow.Tab = 0
    RestoreWinState Me, App.ProductName, mstrTitle
    
    If TabShow.Tab = 1 Then
        With mshDetail
            .ColWidth(12) = IIf(mblnCostView = True, 1000, 0)
            .ColWidth(15) = IIf(mblnCostView = True, 1000, 0)
            .ColWidth(.Cols - 2) = IIf(mblnCostView = True, 1500, 0)
            .ColWidth(.Cols - 1) = IIf(mblnCostView = True, 1500, 0)
        End With
        With mshlist
            .ColWidth(M_INT_COL�̵�ɱ����) = IIf(mblnCostView = True, 1000, 0)
            .ColWidth(M_INT_COL�̵�ɱ�����) = IIf(mblnCostView = True, 1000, 0)
        End With
    End If
    mblnBootUp = True
    
    If IsObject(FrmMain) Then
        Me.Show , FrmMain
    Else
        OS.ShowChildWindow Me.hwnd, FrmMain
    End If
    
    Me.ZOrder 0
End Sub

'�������������
Private Function CheckDepend() As Boolean
    Dim rsTemp As New Recordset
    
    On Error GoTo errHandle
    CheckDepend = False
    mstr�������� = "V,W,K,12"
    gstrSQL = "" & _
            "   SELECT DISTINCT a.id, a.���� " & _
            "   FROM ��������˵�� c, �������ʷ��� b, ���ű� a " & _
            "   Where c.�������� = b.���� and (a.վ��=[2] or a.վ�� is null) " & _
            "       And b.���� In('V','K','W','12') " & _
            "       AND a.id = c.����id " & _
            "       AND TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01'" & _
            IIf(InStr(gstrPrivs, "���пⷿ") <> 0, "", " And a.ID IN (Select ����ID From ������Ա Where ��ԱID=[1])")
                        
    mbln����Ա���� = Not zlStr.IsHavePrivs(gstrPrivs, "���пⷿ")
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, mstrTitle, UserInfo.Id, gstrNodeNo)
    
    If rsTemp.EOF Then
        ShowMsgBox "����Ӧ������һ�����пⷿ���ʡ����ϲ���" & vbCrLf & "�����Ƽ������ʵĲ���,��鿴���Ź���"
        rsTemp.Close
        Exit Function
    End If
    
    With cboStock
        .Clear
        Do While Not rsTemp.EOF
            .AddItem rsTemp!����
            .ItemData(.NewIndex) = rsTemp!Id
            If rsTemp!Id = UserInfo.����ID Then
                .ListIndex = .NewIndex
            End If
            rsTemp.MoveNext
        Loop
        rsTemp.Close
        
        If .ListIndex = -1 Then
            If InStr(gstrPrivs, "���пⷿ") = 0 Then
                ShowMsgBox "�㲻�Ƿ��ϲ��Ż�ⷿ������Ա�Ҳ��������пⷿ��Ȩ�ޣ����ܽ��룡"
                Unload Me
                Exit Function
            End If
            .ListIndex = 0
        End If
    End With
    CheckDepend = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub GetList(ByVal strFind As String)
    Dim rsTemp As New Recordset
    Dim strUserPart As String
    Dim dbl�̵�ɱ���� As Double
    Dim dbl�̵�ɱ����� As Double
    Dim intCol As Integer
    Dim intRow As Integer
    
    mlastRow = 0
    On Error GoTo errHandle
    Call FS.ShowFlash("���������������ϼ�¼,���Ժ� ...", Me)
    DoEvents
    Screen.MousePointer = vbHourglass
    strUserPart = " And A.�ⷿID+0=[1]"
    
    mshlist.Redraw = False
    
    
    'Ƶ���ֶα���� �̵�ʱ��
    
    If TabShow.Tab = 1 Then
        gstrSQL = "" & _
            "   SELECT distinct a.no, Ƶ�� AS �̵�ʱ��," & _
            "           a.������,TO_CHAR (min(a.��������), 'yyyy-mm-dd HH24:Mi:SS') AS ��������, a.�����," & _
            "           TO_CHAR (min(a.�������), 'yyyy-mm-dd HH24:Mi:SS') AS �������, " & _
            "           ltrim(to_char((Sum(Nvl(����,0)*���ۼ�))," & mOraFMT.FM_��� & ")) �̵���," & _
            "           ltrim(to_char((Sum(���۽��*decode(sign(Nvl(����,0)-��д����),-1,-1,1)))," & mOraFMT.FM_��� & ")) ����," & _
            "           LTrim(to_Char(sum(a.�ɱ���+to_char(a.���۽��*a.���ϵ��*Decode(a.��¼״̬, 1, 1, Decode(Mod(a.��¼״̬, 3), 0, 1, -1))," & mORaFMT��¼��.FM_��� & ")" & "-(a.�ɱ����+to_char(a.���*a.���ϵ��*Decode(a.��¼״̬, 1, 1, Decode(Mod(a.��¼״̬, 3), 0, 1, -1))," & mORaFMT��¼��.FM_��� & ")))," & mORaFMT��¼��.FM_��� & ")) as  �̵�ɱ����, " & _
            "           LTrim(To_Char(sum(a.���۽��*a.���ϵ��-a.���*a.���ϵ��), " & mOraFMT.FM_��� & ")) as �̵�ɱ����� , " & _
            "           a.��¼״̬, a.ժҪ " & _
            "   FROM ҩƷ�շ���¼ a, ���ű� b " & _
            "   Where a.�ⷿid = b.ID AND a.���� =22  " & strUserPart & strFind & _
            "   Group by a.no,Ƶ��,a.������,a.�����,a.��¼״̬, a.ժҪ " & _
            "   ORDER BY no DESC,�������� ASC "
    Else
        gstrSQL = "" & _
            "   SELECT distinct a.no, Ƶ�� AS �̵�ʱ��," & _
            "           a.������,TO_CHAR (min(a.��������), 'yyyy-mm-dd HH24:Mi:SS') AS ��������,a.ժҪ,��� " & _
            "   FROM ҩƷ�շ���¼ a, ���ű� b " & _
            "   Where  a.�ⷿid = b.ID  and a.���� = 23  " & strUserPart & strFind & _
            "   Group by a.no,Ƶ��,a.������,a.ժҪ,��� " & _
            "   ORDER BY no DESC,�������� ASC "
    End If
    
    'mstrOthers(0 To 6) As String ' 0-��¼״̬(�ƻ�����),1-��ʼ����,2-��������,3-����id,4-�Է�����id(��������id����Ʒ���(�ƻ���)),5-������,6-�����,7-��Ӧ��ID,8-������,9-��ʼ��������,10-������������,11-��ʼ��Ʊ��,12-������Ʊ��
    '������Χ:[1]-�ⷿid,[2]:��ʼ��������,[3]������������,[4]��ʼ�������,[5] �����������,[6]-��¼״̬,[7]��ʼ���ݺ�,[8]�������ݺ�,[9]����id,[10]�Է�����id,[11]������,[12]�����
    ' δ�Ͳ���: [13]-��Ӧ��ID,[14]-������,[15]-��ʼ��������,[16]-������������,[17]-��ʼ��Ʊ��,[18]-������Ʊ��
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption, cboStock.ItemData(cboStock.ListIndex), _
        CDate(Format(mdtStartDate, "yyyy-mm-dd") & " 00:00:00"), CDate(Format(mdtEndDate, "yyyy-mm-dd") & " 23:59:59"), _
        CDate(Format(mdtVerifyStart, "yyyy-mm-dd") & " 00:00:00"), CDate(Format(mdtVerifyEnd, "yyyy-mm-dd") & " 23:59:59"), _
        Val(mstrOthers(0)), mstrOthers(1), mstrOthers(2), Val(mstrOthers(3)), _
        Val(mstrOthers(4)), mstrOthers(5), mstrOthers(6))
    
    With mshlist
        .Rows = 1
        If TabShow.Tab = 1 Then
            Do While Not rsTemp.EOF
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, M_INT_COLNO) = IIf(IsNull(rsTemp!NO), "", rsTemp!NO)
                .TextMatrix(.Rows - 1, M_INT_COL�̵�ʱ��) = IIf(IsNull(rsTemp!�̵�ʱ��), "", rsTemp!�̵�ʱ��)
                .TextMatrix(.Rows - 1, M_INT_COL������) = IIf(IsNull(rsTemp!������), "", rsTemp!������)
                .TextMatrix(.Rows - 1, M_INT_COL��������) = IIf(IsNull(rsTemp!��������), "", rsTemp!��������)
                .TextMatrix(.Rows - 1, M_INT_COL�����) = IIf(IsNull(rsTemp!�����), "", rsTemp!�����)
                .TextMatrix(.Rows - 1, M_INT_COL�������) = IIf(IsNull(rsTemp!�������), "", rsTemp!�������)
                .TextMatrix(.Rows - 1, M_INT_COL�̵���) = IIf(IsNull(rsTemp!�̵���), "", rsTemp!�̵���)
                .TextMatrix(.Rows - 1, M_INT_COL����) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
                .TextMatrix(.Rows - 1, M_INT_COL�̵�ɱ����) = IIf(IsNull(rsTemp!�̵�ɱ����), "", rsTemp!�̵�ɱ����)
                .TextMatrix(.Rows - 1, M_INT_COL�̵�ɱ�����) = IIf(IsNull(rsTemp!�̵�ɱ�����), "", rsTemp!�̵�ɱ�����)
                .TextMatrix(.Rows - 1, M_INT_COL��¼״̬) = IIf(IsNull(rsTemp!��¼״̬), "", rsTemp!��¼״̬)
                .TextMatrix(.Rows - 1, M_INT_COLժҪ) = IIf(IsNull(rsTemp!ժҪ), "", rsTemp!ժҪ)
                
                dbl�̵�ɱ���� = dbl�̵�ɱ���� + IIf(IsNull(rsTemp!�̵�ɱ����), "", rsTemp!�̵�ɱ����)
                dbl�̵�ɱ����� = dbl�̵�ɱ����� + IIf(IsNull(rsTemp!�̵�ɱ�����), "", rsTemp!�̵�ɱ�����)
                rsTemp.MoveNext
            Loop
            lblSum�ɱ����.Caption = "�̵�ɱ����ϼƣ�" & GetFormat(dbl�̵�ɱ����, g_С��λ��.obj_��װС��.���С��) & "Ԫ"
            lbl�ɱ�����.Caption = "�ɱ�����ϼƣ�" & GetFormat(dbl�̵�ɱ�����, g_С��λ��.obj_��װС��.���С��) & "Ԫ"
        Else
            Do While Not rsTemp.EOF
                .Rows = .Rows + 1
                                
                .TextMatrix(.Rows - 1, M_INT_COL�̵㵥NO) = IIf(IsNull(rsTemp!NO), "", rsTemp!NO)
                .TextMatrix(.Rows - 1, M_INT_COL�̵㵥�̵�ʱ��) = IIf(IsNull(rsTemp!�̵�ʱ��), "", rsTemp!�̵�ʱ��)
                .TextMatrix(.Rows - 1, M_INT_COL�̵㵥������) = IIf(IsNull(rsTemp!������), "", rsTemp!������)
                .TextMatrix(.Rows - 1, M_INT_COL�̵㵥��������) = IIf(IsNull(rsTemp!��������), "", rsTemp!��������)
                .TextMatrix(.Rows - 1, M_INT_COL�̵㵥ժҪ) = IIf(IsNull(rsTemp!ժҪ), "", rsTemp!ժҪ)
                .TextMatrix(.Rows - 1, M_INT_COL�̵㵥���) = IIf(IsNull(rsTemp!���), "0", rsTemp!���)
                If .TextMatrix(.Rows - 1, M_INT_COL�̵㵥���) <> 0 Then
                    .TextMatrix(.Rows - 1, M_INT_COL�̵㵥����) = "��"
                End If
                
                rsTemp.MoveNext
            Loop
        End If
    End With
'    Set mshList.Recordset = rsTemp
    With mshlist
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
    End With
    SetListColWidth
    
    mshlist_EnterCell    '�г�������
    
    If TabShow.Tab = 1 Then
        SetStrikeColor
    End If
    
    With mshlist
        .Row = 1
        .Col = 0
'        .ColSel = .Cols - 1
    End With
    
    mshlist.Redraw = True
    Call FS.StopFlash
    
    Screen.MousePointer = vbDefault
    stbThis.Panels(2).Text = "��ǰ����" & rsTemp.RecordCount & "�ŵ���"
    rsTemp.Close
    If mshlist.Visible = True Then
        mshlist.SetFocus
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub initGrid()
    mblnLoadGrid = False
    '��ʼ���б�
    With mshlist
        .Rows = 2
        If TabShow.Tab = 1 Then
            '�̵��
            .Cols = M_INT_ALLCOLUMN
            .TextMatrix(0, M_INT_COLNO) = "NO"
            .TextMatrix(0, M_INT_COL�̵�ʱ��) = "�̵�ʱ��"
            .TextMatrix(0, M_INT_COL������) = "������"
            .TextMatrix(0, M_INT_COL��������) = "��������"
            .TextMatrix(0, M_INT_COL�����) = "�����"
            .TextMatrix(0, M_INT_COL�������) = "�������"
            .TextMatrix(0, M_INT_COL�̵���) = "�̵���"
            .TextMatrix(0, M_INT_COL����) = "����"
            .TextMatrix(0, M_INT_COL�̵�ɱ����) = "�̵�ɱ����"
            .TextMatrix(0, M_INT_COL�̵�ɱ�����) = "�̵�ɱ�����"
            .TextMatrix(0, M_INT_COL��¼״̬) = "��¼״̬"
            .TextMatrix(0, M_INT_COLժҪ) = "ժҪ"
            
            .ColAlignment(M_INT_COLNO) = flexAlignLeftCenter  'no
            .ColAlignment(M_INT_COL�̵�ʱ��) = flexAlignLeftCenter '�̵�ʱ��
            .ColAlignment(M_INT_COL������) = flexAlignLeftCenter '������
            .ColAlignment(M_INT_COL�������) = flexAlignLeftCenter '��������
            .ColAlignment(M_INT_COL�����) = flexAlignLeftCenter '�����
            .ColAlignment(M_INT_COL�������) = flexAlignLeftCenter '�������
            .ColAlignment(M_INT_COL�̵���) = flexAlignLeftCenter '�̵���
            .ColAlignment(M_INT_COL����) = flexAlignRightCenter '����
            .ColAlignment(M_INT_COL�̵�ɱ����) = flexAlignRightCenter '�̵�ɱ����
            .ColAlignment(M_INT_COL�̵�ɱ�����) = flexAlignRightCenter '�̵�ɱ�����
            .ColAlignment(M_INT_COL��¼״̬) = flexAlignRightCenter '��¼״̬
            .ColAlignment(M_INT_COLժҪ) = flexAlignRightCenter 'ժҪ
        Else
            '�̵㵥
            .Cols = M_INT_�̵㵥ALLCOLUMN
            
            .TextMatrix(0, M_INT_COL�̵㵥����) = "����"
            .TextMatrix(0, M_INT_COL�̵㵥NO) = "NO"
            .TextMatrix(0, M_INT_COL�̵㵥�̵�ʱ��) = "�̵�ʱ��"
            .TextMatrix(0, M_INT_COL�̵㵥������) = "������"
            .TextMatrix(0, M_INT_COL�̵㵥��������) = "��������"
            .TextMatrix(0, M_INT_COL�̵㵥ժҪ) = "ժҪ"
            .TextMatrix(0, M_INT_COL�̵㵥���) = "���"
            
            .ColAlignment(M_INT_COL�̵㵥����) = flexAlignCenterCenter   '���
            .ColAlignment(M_INT_COL�̵㵥NO) = flexAlignLeftCenter  'no
            .ColAlignment(M_INT_COL�̵㵥�̵�ʱ��) = flexAlignLeftCenter '�̵�ʱ��
            .ColAlignment(M_INT_COL�̵㵥������) = flexAlignLeftCenter  '������
            .ColAlignment(M_INT_COL�̵㵥��������) = flexAlignLeftCenter '��������
            .ColAlignment(M_INT_COL�̵㵥ժҪ) = flexAlignLeftCenter 'ժҪ
            .ColAlignment(M_INT_COL�̵㵥���) = flexAlignLeftCenter  '���
        End If
    End With
    mblnLoadGrid = True
End Sub

Private Sub SetStrikeColor()
    Dim intStatus As Integer
    Dim intRow As Integer
    Dim intCol As Integer
    
    With mshlist
        If .Rows <= 2 Then Exit Sub
        For intRow = 1 To .Rows - 1
            intStatus = IIf(TabShow.Tab = 0, 1, Val(.TextMatrix(intRow, M_INT_COL��¼״̬)))
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
                    .CellForeColor = &HFF
                Next
            End If
        Next
    End With
End Sub

'��ͷ�п��ʼ
Private Sub SetListColWidth()
    Dim intCol As Integer
    
    With mshlist
        If TabShow.Tab = 1 Then
            If mblnBootUp = False Then
                For intCol = 1 To .Cols - 1
                    If intCol = 1 Then
                        .ColWidth(intCol) = 2000
                    ElseIf intCol = M_INT_COL��¼״̬ Then
                        .ColWidth(intCol) = 0
                    Else
                        If intCol = M_INT_COL�̵�ɱ���� Or intCol = M_INT_COL�̵�ɱ����� Then
                            .ColWidth(intCol) = 1500
                        Else
                            .ColWidth(intCol) = 1000
                        End If
                    End If
                Next
            End If
            
            .ColHidden(M_INT_COL��¼״̬) = True
            
            .ColWidth(M_INT_COL�������) = 1000
            .ColWidth(M_INT_COL�̵�ɱ����) = IIf(mblnCostView = False, 0, 1500)
            .ColWidth(M_INT_COL�̵�ɱ�����) = IIf(mblnCostView = False, 0, 1500)
        Else
            If mblnBootUp = False Then
                .ColWidth(M_INT_COL�̵㵥�̵�ʱ��) = 2000
                .ColWidth(M_INT_COL�̵㵥ժҪ) = 3000
            End If
            .ColWidth(M_INT_COL�̵㵥���) = 0
        End If
        Call RestoreFlexState(mshlist, TabShow.TabCaption(TabShow.Tab))
        If TabShow.Tab = 1 Then
            .ColWidth(M_INT_COL�̵�ɱ����) = IIf(mblnCostView = True, 1500, 0)
            .ColWidth(M_INT_COL�̵�ɱ�����) = IIf(mblnCostView = True, 1500, 0)
        End If
    End With
End Sub

Private Sub SetDetailColWidth()
    Dim intCol As Integer
    
    With mshDetail
        .ColAlignment(4) = flexAlignCenterCenter   '��λ
        .ColAlignment(IIf(TabShow.Tab = 1, 9, 7)) = flexAlignRightCenter 'ʵ����
        If TabShow.Tab = 1 Then
            .ColAlignment(8) = flexAlignRightCenter     '������
            .ColAlignment(10) = flexAlignCenterCenter    '��־
            .ColAlignment(11) = flexAlignRightCenter     '������
            .ColAlignment(12) = flexAlignRightCenter    '�ɱ���
            .ColAlignment(13) = flexAlignRightCenter    '�ۼ�
            .ColAlignment(14) = flexAlignRightCenter    '����
            .ColAlignment(15) = flexAlignRightCenter    '��۲�
            .ColAlignment(16) = flexAlignRightCenter    '�̵���
            .ColAlignment(.Cols - 2) = flexAlignRightCenter '�̵�ɱ����
            .ColAlignment(.Cols - 1) = flexAlignRightCenter '�̵�ɱ�����
        Else
            .ColAlignment(8) = flexAlignRightCenter '�ɱ���
            .ColAlignment(9) = flexAlignRightCenter '�ɱ����
            .ColAlignment(10) = flexAlignRightCenter '�ۼ�
            .ColAlignment(11) = flexAlignRightCenter '�ۼ۽��
            
        End If
        
        If TabShow.Tab = 1 Then
            .ColWidth(.Cols - 1) = 1500
            .ColWidth(.Cols - 2) = 1500
            
            If mblnBootUp = False Then
                .ColWidth(0) = 0
                .ColWidth(1) = 2500
                For intCol = 2 To .Cols - 1
                    .ColWidth(intCol) = 1000
                    If intCol = .Cols - 1 Or intCol = .Cols - 2 Then
                        .ColWidth(intCol) = 1500
                    End If
                Next
                If mlngMode = 1300 Then
                    .ColWidth(16) = 0
                End If
                .ColWidth(.Cols - 2) = 0
            End If
            
            .ColWidth(12) = IIf(mblnCostView = False, 0, 1000)
            .ColWidth(15) = IIf(mblnCostView = False, 0, 1000)
            .ColWidth(.Cols - 2) = IIf(mblnCostView = False, 0, 1500)
            .ColWidth(.Cols - 1) = IIf(mblnCostView = False, 0, 1500)
        Else
            .ColWidth(0) = 0
            .ColWidth(1) = 2500
            For intCol = 2 To .Cols - 1
                .ColWidth(intCol) = 1000
            Next
        End If
        Call RestoreFlexState(mshDetail, TabShow.TabCaption(TabShow.Tab))
        
        If TabShow.Tab = 1 Then
            .ColWidth(12) = IIf(mblnCostView = False, 0, 1000)
            .ColWidth(15) = IIf(mblnCostView = False, 0, 1000)
            .ColWidth(.Cols - 2) = IIf(mblnCostView = False, 0, 1500)
            .ColWidth(.Cols - 1) = IIf(mblnCostView = False, 0, 1500)
        End If
    End With
End Sub


'����Ȩ�����ò�ͬ����ʾ��Ŀ
Private Sub SetVisable()
    '�⹺�������Ȩ�ޣ��������á����������пⷿ���Ǽǡ��޸ġ�ɾ�������ա����������ݴ�ӡ
'    If InStr(1, gstrPrivs, "��������") = 0 Then
'         mnuFileParameter.Visible = False
'         mnuFileLine3.Visible = False                '��Ӧ�ķָ���
'    End If
'
    If InStr(1, gstrPrivs, "�Ǽ�") = 0 Then
        mnuEditAddBill.Visible = False
        mnuEditAddTable.Visible = False
        tlbTool.Buttons("Bill").Visible = False
        tlbTool.Buttons("Table").Visible = False
    End If
    
    If InStr(1, gstrPrivs, "�޸�") = 0 Then
        mnuEditModify.Visible = False
        tlbTool.Buttons("Modify").Visible = False
    End If
    
    If InStr(1, gstrPrivs, "ɾ��") = 0 Then
        mnuEditDel.Visible = False
        tlbTool.Buttons("Delete").Visible = False
         '��û�����б༭Ȩ��ʱ���Ѳ˵��͹������ϵ���Ӧ�ķָ������Ρ�
        If mnuEditAddBill.Visible = False And mnuEditModify.Visible = False Then
            mnuEditLine1.Visible = False
            tlbTool.Buttons("EditSeparate").Visible = False
        End If
    End If
    
    If InStr(1, gstrPrivs, "���") = 0 Then
        mnuEditVerify.Visible = False
        tlbTool.Buttons("Verify").Visible = False
    End If
    
    If InStr(1, gstrPrivs, "����") = 0 Then
        mnuEditStrike.Visible = False
        tlbTool.Buttons("Strike").Visible = False
        
        If mnuEditVerify.Visible = False Then
            mnuEditLine2.Visible = False
            tlbTool.Buttons("VerifySeparate").Visible = False
        End If
    End If
    If InStr(1, gstrPrivs, "���ݴ�ӡ") = 0 Then
        mnuFileBillPrint.Visible = False
        mnuFileBillPreview.Visible = False
    End If
End Sub

Private Sub Cmd����_Click()
    Call mnuEditDisplay_Click
End Sub

Private Sub Form_Load()
    Dim strReg As String
    Dim strOthers(0 To 6) As String
    Dim i As Integer
    For i = 0 To 6
        strOthers(i) = ""
    Next
    mstrOthers = strOthers
    strReg = Val(zlDataBase.GetPara("�̵��λ", glngSys, mlngMode, "0"))
    mintUnit = Val(strReg)
    mintUnit1 = IIf(Val(zlDataBase.GetPara("��¼����λ", glngSys, mlngMode, "0")) = 1, 1, 0)
    mstrOrder = zlDataBase.GetPara("��������", glngSys, mlngMode, "00")
    mblnCostView = zlStr.IsHavePrivs(mstrPrivs, "�鿴�ɱ���")
  
    '���˺�:����С����ʽ����
    With mOraFMT
        .FM_�ɱ��� = GetFmtString(mintUnit, g_�ɱ���, True)
        .FM_��� = GetFmtString(mintUnit, g_���, True)
        .FM_���ۼ� = GetFmtString(mintUnit, g_�ۼ�, True)
        .FM_���� = GetFmtString(mintUnit, g_����, True)
    End With
    With mORaFMT��¼��
        .FM_�ɱ��� = GetFmtString(mintUnit1, g_�ɱ���, True)
        .FM_��� = GetFmtString(mintUnit1, g_���, True)
        .FM_���ۼ� = GetFmtString(mintUnit1, g_�ۼ�, True)
        .FM_���� = GetFmtString(mintUnit1, g_����, True)
    End With

    '�ָ�����
    Me.Caption = mstrTitle
    PrintRange "��ѯ��Χ:" & Format(sys.Currentdate, "yyyy��MM��dd��") & "��" & Format(sys.Currentdate, "yyyy��MM��dd��")
    Call RestoreWinState(Me, App.ProductName)
    
    '2006-04-25:���˺�,ͳһ���ӱ�������ģ��Ĺ���
    Call zlDataBase.ShowReportMenu(Me, glngSys, glngModul, gstrPrivs)
     
End Sub

Private Sub Form_Resize()
    '����λ������
    
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
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
        .Height = 300
        .Left = 0
        .Width = cbrTool.Width
        
    End With
    
    With lbl�ɱ�����
        .Left = Me.Width - .Width - 1700
        .Top = picSeparate_s.Height - 200
    End With
    If mblnCostView = False Then
        lbl�ɱ�����.Visible = False
    End If
    
    With lblSum�ɱ����
        .Left = lbl�ɱ�����.Left - .Width - 600
        .Top = picSeparate_s.Height - 200
    End With
    If mblnCostView = False Then
        lblSum�ɱ����.Visible = False
    End If
    
    With TabShow
        .Top = IIf(cbrTool.Visible, cbrTool.Height, 0)
        .Left = 0
        .Width = cbrTool.Width
    End With
    
    With mshlist
        .Top = TabShow.Top + TabShow.Height
        .Left = 0
        .Width = cbrTool.Width
        .Height = picSeparate_s.Top - .Top
    End With
    
    With Cmd����
        .Left = Me.ScaleWidth - .Width - 100
        .Top = mshlist.Top + mshlist.Height + 30
        
    End With
    
    With mshDetail
        .Top = picSeparate_s.Top + picSeparate_s.Height + 100
        .Left = 0
        .Height = ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
        .Width = cbrTool.Width
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName, mstrTitle
    Call SaveFlexState(mshlist, TabShow.TabCaption(TabShow.Tab))
End Sub

Private Sub mnuEditaddBill_Click()
    Dim strNo As String
    Dim blnSuccess As Boolean
    
    frmCheckCourseCard.ShowCard Me, strNo, 1, , mstrPrivs, blnSuccess
    
    If blnSuccess Then Call mnuViewRefresh_Click
End Sub

Private Sub mnuEditAddTableAuto_Click()
    Dim strNo As String
    Dim blnSuccess As Boolean
    
    frmCheckCard.ShowCard Me, strNo, 1, , mstrPrivs, blnSuccess
    
    If blnSuccess Then Call mnuViewRefresh_Click
End Sub

Private Sub mnuEditAddTableTotal_Click()
    Dim strNo As String
    Dim blnSuccess As Boolean
    
    frmCheckCard.ShowCard Me, strNo, 5, , mstrPrivs, blnSuccess
    
    If blnSuccess Then Call mnuViewRefresh_Click
End Sub

Private Sub mnuEditAddTableZero_Click()
    Dim strNo As String
    Dim blnSuccess As Boolean
    
    frmCheckCard.ShowCard Me, strNo, 6, , mstrPrivs, blnSuccess
    
    If blnSuccess Then Call mnuViewRefresh_Click
End Sub

Private Sub mnuEditVerify_Click()
    '����
    
    Dim strNo As String
    Dim blnSuccess As Boolean
    
    With mshlist
        strNo = .TextMatrix(.Row, M_INT_COLNO)
        frmCheckCard.ShowCard Me, strNo, 3, .TextMatrix(.Row, M_INT_COL��¼״̬), mstrPrivs, blnSuccess
    End With
    
    If blnSuccess Then Call mnuViewRefresh_Click
End Sub

Private Sub mnuEditDel_Click()
    'ɾ��
    Dim StrBillNo As String
    Dim intRow As Integer
    Dim strTitle As String
    Dim intReturn As Integer
    Dim intRecord As Integer
     
    With mshlist
        strTitle = IIf(TabShow.Tab = 0, "�̵��¼��", "�̵��")
        
        On Error GoTo errHandle
        intRow = .Row
        StrBillNo = .TextMatrix(intRow, M_INT_COLNO)
        intReturn = MsgBox("��ȷʵҪɾ�����ݺ�Ϊ��" & StrBillNo & "����" & strTitle & "��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)
        intRecord = .Rows - 1
        If intReturn = vbYes Then
            If TabShow.Tab = 1 Then
                gstrSQL = "zl_�����̵�_Delete('" & StrBillNo & "')"
            Else
                gstrSQL = "zl_�����̵��¼��_Delete('" & StrBillNo & "')"
            End If
            zlDataBase.ExecuteProcedure gstrSQL, Me.Caption
            
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
'                    .ColSel = .Cols - 1
                End With
                SetEnable
                
            End If
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
    stbThis.Panels(2).Text = "��ǰ����" & intRecord & "�ŵ���"
    Exit Sub

errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditDisplay_Click()
    '�鿴����
    
    Dim strNo As String
    With mshlist
        strNo = .TextMatrix(.Row, M_INT_COLNO)
        If TabShow.Tab = 0 Then
            frmCheckCourseCard.ShowCard Me, strNo, 4
        Else
            frmCheckCard.ShowCard Me, strNo, 4, .TextMatrix(.Row, M_INT_COL��¼״̬), mstrPrivs
        End If
    End With
End Sub

Private Sub mnuEditStrike_Click()
    Dim blnPurchase As Boolean, blnRefresh As Boolean
    
    '������⹺(blnPurchaseΪ��)����ֱ�ӽ������
    'ѯ���Ƿ����(blnPurchaseΪ��ʾ�򷵻�ֵ)������������
    blnPurchase = (InStr(1, "1300,1302,1304,1305,1306", mlngMode) <> 0)
    With mshlist
        If Not blnPurchase Then
            blnPurchase = (MsgBox("��ȷʵҪ�������ݺ�Ϊ��" & .TextMatrix(.Row, M_INT_COLNO) & "���ĵ�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes)
        End If
        If blnPurchase Then
            blnRefresh = StrikeSave
            If blnRefresh Then mnuViewRefresh_Click
        End If
    End With
End Sub

Private Function StrikeSave() As Boolean
    Dim blnSuccess As Boolean
    Dim rsTemp As ADODB.Recordset
    Dim int����� As Integer
    Dim strMsg As String
    Dim n As Integer
    
    StrikeSave = False
    
    int����� = StuffWork_GetCheckStockRule(Val(cboStock.ItemData(cboStock.ListIndex)))
    
    On Error GoTo errHandle
    If int����� <> 0 Then
        gstrSQL = "Select Distinct A.ҩƷ��Ϣ " & _
            " From (Select  '(' || I.���� || ')' || Nvl(N.����, I.����) As ҩƷ��Ϣ, A.ʵ������, Nvl(K.ʵ������, 0) As ������� " & _
            " From ҩƷ�շ���¼ A, (Select ҩƷid, �ⷿid, ʵ������, Nvl(����, 0) ���� From ҩƷ��� Where ���� = 1) K, �������� B, �շ���ĿĿ¼ I, �շ���Ŀ���� N " & _
            " Where A.ҩƷid = K.ҩƷid(+) And A.�ⷿid = K.�ⷿid(+) And Nvl(A.����, 0) = K.����(+) And A.ҩƷid = B.����id And " & _
            " A.ҩƷid = I.ID And A.ҩƷid = N.�շ�ϸĿid(+) And N.����(+) = 3 And A.���� = 22 And A.���ϵ�� = 1 And A.NO = [1]) A " & _
            " Where A.ʵ������ > A.������� "
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "�����", mshlist.TextMatrix(mshlist.Row, 0))
        
        With rsTemp
            If .RecordCount > 0 Then
                For n = 1 To .RecordCount
                    If n > 5 Then
                        strMsg = strMsg & vbCrLf & "��������" & .RecordCount - 5 & "��ҩƷ......"
                        Exit For
                    End If
                    strMsg = IIf(strMsg = "", "", strMsg & "," & vbCrLf) & !ҩƷ��Ϣ
                    .MoveNext
                Next
                
                If int����� = 1 Then
                    If MsgBox("ע�⣬����ҩƷ��治�㣺" & vbCrLf & strMsg & vbCrLf & Space(4) & "�Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                ElseIf int����� = 2 Then
                    MsgBox "�Բ�������ҩƷ��治�㣬���ܳ�����" & vbCrLf & strMsg, vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End With
    End If
    
    With mshlist
        gstrSQL = "zl_�����̵�_Strike('" & .TextMatrix(.Row, M_INT_COLNO) & "','" & UserInfo.�û��� & "')"
        
        zlDataBase.ExecuteProcedure gstrSQL, Me.Caption
    End With
    StrikeSave = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub mnuEditModify_Click()
    '�޸�
    Dim strNo As String
    Dim blnSuccess As Boolean
    
    blnSuccess = False
    With mshlist
        If .TextMatrix(.Row, M_INT_COLNO) = "" Then Exit Sub
        strNo = .TextMatrix(.Row, M_INT_COLNO)
        If TabShow.Tab = 0 Then
            frmCheckCourseCard.ShowCard Me, strNo, 2, 1, mstrPrivs, blnSuccess
        Else
            frmCheckCard.ShowCard Me, strNo, 2, mshlist.TextMatrix(mshlist.Row, M_INT_COL��¼״̬), mstrPrivs, blnSuccess
        End If
        
        If blnSuccess Then Call mnuViewRefresh_Click
    End With
End Sub

Private Sub mnuFileBillPreview_Click()
    With mshlist
        If .TextMatrix(.Row, M_INT_COLNO) = "" Then Exit Sub
        ReportOpen gcnOracle, glngSys, "zl1_bill_1719", Me, "���ݱ��=" & .TextMatrix(.Row, M_INT_COLNO), "��¼״̬=" & .TextMatrix(.Row, M_INT_COL��¼״̬), "��λϵ��=" & mintUnit, 1
    End With
End Sub

Private Sub mnuFileBillPrint_Click()
    With mshlist
        If .TextMatrix(.Row, M_INT_COLNO) = "" Then Exit Sub
        ReportOpen gcnOracle, glngSys, "zl1_bill_1719", Me, "���ݱ��=" & .TextMatrix(.Row, M_INT_COLNO), "��¼״̬=" & .TextMatrix(.Row, M_INT_COL��¼״̬), "��λϵ��=" & mintUnit, 2
    End With
End Sub

Private Sub mnuFileExcel_Click()
    '�����Excel
    
    If Me.ActiveControl Is mshlist Then
        mshlist.Redraw = False
        subPrint 3
        mshlist.Redraw = True
        mshlist.Col = 0
'        mshlist.ColSel = mshlist.Cols - 1
    ElseIf Me.ActiveControl Is mshDetail Then
        mshDetail.Redraw = False
        subExcel 3
        mshDetail.Redraw = True
        mshDetail.Col = 0
'        mshDetail.ColSel = mshDetail.Cols - 1
    End If
End Sub

Private Sub mnufileexit_Click()
    '�˳�
    Unload Me
End Sub

Private Sub mnuFileParameter_Click()
    Dim strReg As String
    '��������
    Call frmParaset.���ò���(mlngMode, mstrPrivs, Me, mstrCaption)
     
    '�̵��¼���ĵ�λ
    mintUnit = Val(zlDataBase.GetPara("�̵��λ", glngSys, mlngMode, "0"))
    mintUnit1 = IIf(Val(zlDataBase.GetPara("��¼����λ", glngSys, mlngMode, "0")) = 1, 1, 0)
    mstrOrder = zlDataBase.GetPara("��������", glngSys, mlngMode, "00")
    mintFindDay = Val(zlDataBase.GetPara("��ѯ����", glngSys, mlngMode, 1))
    mdtStartDate = Format(DateAdd("d", -mintFindDay, sys.Currentdate), "yyyy-MM-dd")
    mdtEndDate = Format(sys.Currentdate, "yyyy-MM-dd")
  
    '���˺�:����С����ʽ����
    With mOraFMT
        .FM_�ɱ��� = GetFmtString(mintUnit, g_�ɱ���, True)
        .FM_��� = GetFmtString(mintUnit, g_���, True)
        .FM_���ۼ� = GetFmtString(mintUnit, g_�ۼ�, True)
        .FM_���� = GetFmtString(mintUnit, g_����, True)
    End With
    With mORaFMT��¼��
        .FM_�ɱ��� = GetFmtString(mintUnit1, g_�ɱ���, True)
        .FM_��� = GetFmtString(mintUnit1, g_���, True)
        .FM_���ۼ� = GetFmtString(mintUnit1, g_�ۼ�, True)
        .FM_���� = GetFmtString(mintUnit1, g_����, True)
    End With
    
    Call GetList(mstrFind)
End Sub

Private Sub mnuFilePreView_Click()
    '��ӡԤ��
    mshlist.Redraw = False
    subPrint 2
    mshlist.Redraw = True
    mshlist.Col = 0
'    mshlist.ColSel = mshlist.Cols - 1
End Sub

Private Sub mnuFilePrint_Click()
    '��ӡ
    mshlist.Redraw = False
    subPrint 1
    mshlist.Redraw = True
    mshlist.Col = 0
'    mshlist.ColSel = mshlist.Cols - 1
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
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
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
    
    Dim strFind As String
    Dim strOthers() As String
    strFind = FrmTransferSearch.GetSearch(Me, mlngMode, mdtStartDate, mdtEndDate, mdtVerifyStart, mdtVerifyEnd, mstrPrivs, strOthers)
    If strFind <> "" Then
        mstrFind = strFind
        mstrOthers = strOthers
       'mstrOthers(0 To 6) As String ' 0-��¼״̬(�ƻ�����),1-��ʼ����,2-��������,3-����id,4-�Է�����id(��������id����Ʒ���(�ƻ���)),5-������,6-�����)
        
        GetList mstrFind
        
        If Format(mdtStartDate, "yyyy-mm-dd") = "1901-01-01" And Format(mdtVerifyStart, "yyyy-mm-dd") = "1901-01-01" Then
        ElseIf Format(mdtStartDate, "yyyy-mm-dd") <> "1901-01-01" And Format(mdtVerifyStart, "yyyy-mm-dd") <> "1901-01-01" Then
            PrintRange "��ѯ��Χ:�������� " & Format(mdtStartDate, "yyyy��MM��dd��") & "��" & Format(mdtEndDate, "yyyy��MM��dd��") & "  ������� " & Format(mdtVerifyStart, "yyyy��MM��dd��") & "��" & Format(mdtVerifyEnd, "yyyy��MM��dd��")
        ElseIf Format(mdtStartDate, "yyyy-mm-dd") <> "1901-01-01" Then
            PrintRange "��ѯ��Χ:�������� " & Format(mdtStartDate, "yyyy��MM��dd��") & "��" & Format(mdtEndDate, "yyyy��MM��dd��")
        ElseIf Format(mdtVerifyStart, "yyyy-mm-dd") <> "1901-01-01" Then
            PrintRange "��ѯ��Χ:������� " & Format(mdtVerifyStart, "yyyy��MM��dd��") & "��" & Format(mdtVerifyEnd, "yyyy��MM��dd��")
        End If
    End If
    
End Sub

Private Sub mnuViewStatus_Click()
    With mnuViewStatus
        .Checked = Not .Checked
        stbThis.Visible = .Checked
    End With
    
    Form_Resize
End Sub
Private Sub mnuReportItem_Click(Index As Integer)
    Dim strNo As String
    Dim intRecodeSta As Integer
    Dim lng�ⷿID As Long
    Dim lngCol As Long
    
    With mshlist
        strNo = Trim(.TextMatrix(.Row, M_INT_COLNO))
        lngCol = GetCol(mshlist, "��¼״̬")
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
    If Format(mdtStartDate, "yyyy-mm-dd") = "1901-01-01" Then
        Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, "NO=" & strNo, "��¼״̬=" & intRecodeSta, "�ⷿ=" & lng�ⷿID)
    Else
        Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, "NO=" & strNo, "��¼״̬=" & intRecodeSta, "�ⷿ=" & lng�ⷿID, "��ʼʱ��=" & Format(mdtStartDate, "yyyy-mm-dd"), "����ʱ��=" & Format(mdtEndDate, "yyyy-mm-dd"))
    End If
End Sub
Private Sub mnuViewToolButton_Click()
    With mnuViewToolButton
        .Checked = Not .Checked
        cbrTool.Bands(1).Visible = .Checked
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
    With mshlist
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
    If mshlist.MouseRow = 0 Then Exit Sub
    mnuEditModify_Click
End Sub

Private Sub mshlist_EnterCell()
    Dim rsTemp As New Recordset
    Dim strUnitQuantity As String               '��λ��������ʽ����
    Dim IntBill As Integer                      '��������  �磺1���⹺��⣻2��
    Dim str��װϵ�� As String
    Dim intTmp As Integer
    Dim str���� As String
    Dim str���� As String
    
    If mblnLoadGrid = False Then Exit Sub
    If mlastRow = mshlist.Row Then Exit Sub
    mlastRow = mshlist.Row
        
    
    On Error GoTo errHandle
    
    If Mid(mstrOrder, 1, 1) = "0" Then
        str���� = " ���"
    ElseIf Mid(mstrOrder, 1, 1) = "1" Then
        str���� = " ������Ϣ"
    ElseIf Mid(mstrOrder, 1, 1) = "2" Then
        str���� = " ����"
    ElseIf Mid(mstrOrder, 1, 1) = "3" Then
        str���� = " �ⷿ��λ"
    End If
    
    If Mid(mstrOrder, 2, 1) = "0" Then
        str���� = str���� & " asc"
    ElseIf Mid(mstrOrder, 2, 1) = "1" Then
        str���� = str���� & " desc"
    End If
    
    If mshlist.Row >= 1 And LTrim(mshlist.TextMatrix(mshlist.Row, M_INT_COLNO)) <> "" Then
        mshlist.Col = 0
        mshlist.ColSel = mshlist.Cols - 1
        If mshlist.RowIsVisible(mshlist.Row) = False Then
           mshlist.TopRow = mshlist.Row
        End If
        mshDetail.Redraw = False
        intTmp = IIf(TabShow.Tab = 1, mintUnit, mintUnit1)
        Select Case intTmp
            Case 0
                strUnitQuantity = "to_char(A.ʵ������," & mOraFMT.FM_���� & ") AS ����," & _
                "c.���㵥λ AS ��λ,"
                str��װϵ�� = "1"
            Case Else
                strUnitQuantity = "(to_char(A.ʵ������ / B.����ϵ��," & mOraFMT.FM_���� & ")) AS ����," & _
                "B.��װ��λ AS ��λ,"
                str��װϵ�� = "B.����ϵ��"
        End Select
            
        IntBill = IIf(TabShow.Tab = 1, 22, 23)
        Dim int��¼״̬ As Integer
        
        If TabShow.Tab = 1 Then
            str���� = "���,������Ϣ,���,����,��׼�ĺ�,��λ,����,ʧЧ��,������,ʵ����,��־,������,�ɱ���,�ۼ�,����,��۲�,�̵���,�̵�ɱ����,�̵�ɱ�����"
            gstrSQL = "" & _
                "   SELECT " & str���� & _
                "   FROM (  SELECT DISTINCT ���,('[' || c.���� || ']' || c.����) AS ������Ϣ," & _
                "                   c.���,c.����,zlSpellCode(c.����) ����,a.����,a.��׼�ĺ�,a.�ⷿ��λ," & IIf(mintUnit = 0, "c.���㵥λ", "b.��װ��λ") & " as ��λ,a.����, to_char(A.Ч��,'yyyy-mm-dd') as ʧЧ��," & _
                "                   (to_char(A.��д���� /" & str��װϵ�� & "," & mOraFMT.FM_���� & ")) AS ������," & _
                "                   (to_char(A.���� /" & str��װϵ�� & "," & mOraFMT.FM_���� & ")) AS ʵ����," & _
                "                   Decode(Sign(A.����-A.��д����),-1,'��',1,'ӯ','ƽ') as ��־," & _
                "                   (to_char(A.ʵ������ /" & str��װϵ�� & "," & mOraFMT.FM_���� & ")) AS ������," & _
                "                   TO_CHAR (a.����*" & str��װϵ�� & "," & mOraFMT.FM_�ɱ��� & ") AS �ɱ���," & _
                "                   TO_CHAR (a.���ۼ�*" & str��װϵ�� & "," & mOraFMT.FM_���ۼ� & ") AS �ۼ�," & _
                "                   TO_CHAR (a.���۽��*a.���ϵ��, " & mOraFMT.FM_��� & ") AS ����," & _
                "                   TO_CHAR (a.���*a.���ϵ��, " & mOraFMT.FM_��� & ") AS ��۲�, " & _
                "                   TO_CHAR ((A.���� / " & str��װϵ�� & ")*(a.���ۼ�*" & str��װϵ�� & "), " & mOraFMT.FM_��� & ") as �̵���, " & _
                "                   To_Char(a.�ɱ���+to_char(a.���۽��*a.���ϵ��*Decode(a.��¼״̬, 1, 1, Decode(Mod(a.��¼״̬, 3), 0, 1, -1))," & mOraFMT.FM_��� & ")" & "-(a.�ɱ����+to_char(a.���*a.���ϵ��*Decode(a.��¼״̬, 1, 1, Decode(Mod(a.��¼״̬, 3), 0, 1, -1))," & mOraFMT.FM_��� & "))," & mOraFMT.FM_��� & ") as  �̵�ɱ����, " & _
                "                   To_Char(a.���۽��*a.���ϵ��  - a.���*a.���ϵ�� ," & mOraFMT.FM_��� & ") AS �̵�ɱ����� " & _
                "           FROM ҩƷ�շ���¼ a, ��������  b,�շ���ĿĿ¼ c" & _
                "           Where a.ҩƷid = b.����id and a.ҩƷid=c.id " & _
                "                   AND  a.��¼״̬ =[3] " & _
                "                   AND a.���� =[1] " & _
                "                   AND a.no =[2] " & _
                "   )" & _
                "  ORDER BY " & str����
                
            int��¼״̬ = Val(mshlist.TextMatrix(mshlist.Row, M_INT_COL��¼״̬))
        Else
            
            str���� = "���,������Ϣ,���,����,��λ,����,ʧЧ��,ʵ����,�ɱ���,�ɱ����,�ۼ�,�ۼ۽��"
            gstrSQL = "" & _
                "   SELECT " & str���� & _
                "   FROM (  SELECT DISTINCT ���,('[' || c.���� || ']' || c.����) AS ������Ϣ," & _
                "                   c.���,c.����,zlSpellCode(c.����) ����,a.����,a.�ⷿ��λ," & IIf(mintUnit1 = 0, "c.���㵥λ", "b.��װ��λ") & " as ��λ,a.����, to_char(A.Ч��,'yyyy-mm-dd') as ʧЧ��," & _
                "                   (to_char(A.���� /" & str��װϵ�� & "," & mORaFMT��¼��.FM_���� & ")) AS ʵ����," & _
                                    IIf(mintUnit1 = 0, "to_char(a.����," & mORaFMT��¼��.FM_�ɱ��� & ") �ɱ���, ", "to_char(a.���� * " & str��װϵ�� & "," & mORaFMT��¼��.FM_�ɱ��� & ") �ɱ���,") & _
                "                   to_char(A.���� * A.����," & mORaFMT��¼��.FM_��� & ") �ɱ����," & _
                                    IIf(mintUnit1 = 0, "to_char(a.���ۼ�," & mORaFMT��¼��.FM_���ۼ� & ") �ۼ�,", "to_char(a.���ۼ� * " & str��װϵ�� & "," & mORaFMT��¼��.FM_���ۼ� & ") �ۼ�,") & _
                "                   to_char(A.���ۼ� * A.����," & mORaFMT��¼��.FM_��� & ") �ۼ۽��" & _
                "           FROM ҩƷ�շ���¼ a, ��������  b,�շ���ĿĿ¼ c " & _
                "           Where a.ҩƷid = b.����id and a.ҩƷid=c.id " & _
                "                   AND a.��¼״̬ =[3] AND a.���� = [1] " & _
                "                   AND a.no = [2] " & _
                "       )" & _
                " ORDER BY  " & str����
                int��¼״̬ = 1
        End If
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption, IntBill, mshlist.TextMatrix(mshlist.Row, M_INT_COLNO), int��¼״̬)
        
        Set mshDetail.DataSource = rsTemp
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
'            .ColSel = .Cols - 1
        End With
        
        mshDetail.Redraw = True
    ElseIf LTrim(mshlist.TextMatrix(mshlist.Row, M_INT_COLNO)) = "" Then
        With mshDetail
            .Cols = IIf(TabShow.Tab = 1, 19, 12)
            .Rows = 2
            .Clear
            .TextMatrix(0, 0) = "���"
            .TextMatrix(0, 1) = "������Ϣ"
            .TextMatrix(0, 2) = "���"
            .TextMatrix(0, 3) = "����"
            .TextMatrix(0, 4) = "��׼�ĺ�"
            .TextMatrix(0, 5) = "��λ"
            .TextMatrix(0, 6) = "����"
            .TextMatrix(0, 7) = "ʧЧ��"
            .TextMatrix(0, IIf(TabShow.Tab = 1, 9, 7)) = "ʵ����"
            
            If TabShow.Tab = 1 Then
                .TextMatrix(0, 8) = "������"
                .TextMatrix(0, 10) = "��־"
                .TextMatrix(0, 11) = "������"
                .TextMatrix(0, 12) = "�ɱ���"
                .TextMatrix(0, 13) = "�ۼ�"
                .TextMatrix(0, 14) = "����"
                .TextMatrix(0, 15) = "��۲�"
                .TextMatrix(0, 16) = "�̵���"
                .TextMatrix(0, 17) = "�̵�ɱ����"
                .TextMatrix(0, 18) = "�̵�ɱ�����"
            Else
                .TextMatrix(0, 8) = "�ɱ���"
                .TextMatrix(0, 9) = "�ɱ����"
                .TextMatrix(0, 10) = "�ۼ�"
                .TextMatrix(0, 11) = "�ۼ۽��"
            End If
            
            .Row = 1
            .Col = 0
'            .ColSel = .Cols - 1
        End With
    End If
    SetDetailColWidth
    SetEnable
    Exit Sub
errHandle:
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

Private Sub mshlist_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    If mnuEdit.Visible = False Then Exit Sub
    
    PopupMenu mnuEdit, 2
End Sub

Private Sub picSeparate_s_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
        If Button <> 1 Then Exit Sub
        mintOldY = y
End Sub

Private Sub picSeparate_s_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    '�ָ�������
    
    If Button <> 1 Then Exit Sub
    
    With picSeparate_s
        If .Top + y < 2000 Then Exit Sub
        If .Top + y > ScaleHeight - 2000 Then Exit Sub
        .Move .Left, .Top + y - mintOldY
    End With
    
    With mshlist
        .Height = picSeparate_s.Top - .Top
    End With
    
    With Cmd����
        .Top = mshlist.Top + mshlist.Height + 30
    End With
    
    With mshDetail
        .Top = picSeparate_s.Top + picSeparate_s.Height + 100
        .Height = ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
    End With
End Sub

Private Sub picSeparate_s_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        If Button <> 1 Then Exit Sub
        mintOldY = 0
End Sub

Private Sub tabShow_Click(PreviousTab As Integer)
    Call SaveFlexState(mshlist, TabShow.TabCaption(PreviousTab))
    Call initGrid
    GetList (mstrFind)  '�г�����ͷ
    
    If TabShow.Tab = 1 Then
        lblSum�ɱ����.Visible = True
        lbl�ɱ�����.Visible = True
    Else
        lblSum�ɱ����.Visible = False
        lbl�ɱ�����.Visible = False
    End If
End Sub

Private Sub tlbTool_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "PrintView"
            mnuFilePreView_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Bill"
            mnuEditaddBill_Click
        Case "Table"
            mnuEditAddTableAuto_Click
        Case "Modify"
            mnuEditModify_Click
        Case "Delete"
            mnuEditDel_Click
        Case "Verify"
            mnuEditVerify_Click
        Case "Strike"
            mnuEditStrike_Click
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
    Dim strVerify As String, blnVisible As Boolean
    
    blnVisible = (TabShow.Tab = 1)
    mnuEditVerify.Visible = blnVisible And (InStr(1, gstrPrivs, "���") <> 0)
    mnuEditStrike.Visible = blnVisible And (InStr(1, gstrPrivs, "����") <> 0)
    mnuEditLine2.Visible = blnVisible And (mnuEditVerify.Visible Or mnuEditStrike.Visible)
    tlbTool.Buttons("Verify").Visible = blnVisible And (InStr(1, gstrPrivs, "���") <> 0)
    tlbTool.Buttons("Strike").Visible = blnVisible And (InStr(1, gstrPrivs, "����") <> 0)
    
    tlbTool.Buttons("VerifySeparate").Visible = mnuEditLine2.Visible
    
    mnuFileBillPreview.Visible = blnVisible And (InStr(1, gstrPrivs, "���ݴ�ӡ") <> 0)
    mnuFileBillPrint.Visible = blnVisible And (InStr(1, gstrPrivs, "���ݴ�ӡ") <> 0)
    
    With mshlist
        .ToolTipText = ""
        If .TextMatrix(.Row, M_INT_COLNO) = "" Or .Row = 0 Then          'û�е�
            mnuFilePreView.Enabled = False
            mnuFilePrint.Enabled = False
            mnuFileBillPreview.Enabled = False
            mnuFileBillPrint.Enabled = False
            mnuFileExcel.Enabled = False
            tlbTool.Buttons("Print").Enabled = False
            tlbTool.Buttons("PrintView").Enabled = False
        
            If mnuEditModify.Visible = True Then
                mnuEditModify.Enabled = False
                tlbTool.Buttons("Modify").Enabled = False
            End If
            If mnuEditDel.Visible = True Then
                mnuEditDel.Enabled = False
                tlbTool.Buttons("Delete").Enabled = False
            End If
            If mnuEditVerify.Visible = True Then
                mnuEditVerify.Enabled = False
                tlbTool.Buttons("Verify").Enabled = False
            End If
            
            If mnuEditStrike.Visible = True Then
                mnuEditStrike.Enabled = False
                tlbTool.Buttons("Strike").Enabled = False
            End If
             
            If mnuEditDisplay.Visible = True Then
                mnuEditDisplay.Enabled = False
            End If
         Else
            mnuFilePreView.Enabled = True
            mnuFilePrint.Enabled = True
            mnuFileBillPreview.Enabled = True
            mnuFileBillPrint.Enabled = True
            mnuFileExcel.Enabled = True
            tlbTool.Buttons("Print").Enabled = True
            tlbTool.Buttons("PrintView").Enabled = True
            
            If TabShow.Tab = 1 Then
                strVerify = .TextMatrix(.Row, M_INT_COL�������)
            Else
                strVerify = ""
            End If
            If strVerify = "" Then    'δ��˵�
                If mnuEditModify.Visible = True Then
                    mnuEditModify.Enabled = True
                    tlbTool.Buttons("Modify").Enabled = True
                End If
                If mnuEditDel.Visible = True Then
                    mnuEditDel.Enabled = True
                    tlbTool.Buttons("Delete").Enabled = True
                End If
                If mnuEditVerify.Visible = True Then
                    mnuEditVerify.Enabled = True
                    tlbTool.Buttons("Verify").Enabled = True
                End If
                
                If mnuEditStrike.Visible = True Then
                    mnuEditStrike.Enabled = False
                    tlbTool.Buttons("Strike").Enabled = False
                End If
                 
                If mnuEditDisplay.Visible = True Then
                    mnuEditDisplay.Enabled = True
                End If
            ElseIf .TextMatrix(.Row, M_INT_COL��¼״̬) = 1 Then     '��˵�
                If mnuEditModify.Visible = True Then
                    mnuEditModify.Enabled = False
                    tlbTool.Buttons("Modify").Enabled = False
                End If
                If mnuEditDel.Visible = True Then
                    mnuEditDel.Enabled = False
                    tlbTool.Buttons("Delete").Enabled = False
                End If
                If mnuEditVerify.Visible = True Then
                    mnuEditVerify.Enabled = False
                    tlbTool.Buttons("Verify").Enabled = False
                End If
                
                If mnuEditStrike.Visible = True Then
                    mnuEditStrike.Enabled = True
                    tlbTool.Buttons("Strike").Enabled = True
                End If
                 
                If mnuEditDisplay.Visible = True Then
                    mnuEditDisplay.Enabled = True
                End If
            Else   '2,3 ������
                If .TextMatrix(.Row, M_INT_COL��¼״̬) Mod 3 = 0 Then
                    .ToolTipText = "�������ݵ�ԭ����"
                    If mnuEditStrike.Visible = True Then
                        mnuEditStrike.Enabled = True
                        tlbTool.Buttons("Strike").Enabled = True
                    End If
                ElseIf .TextMatrix(.Row, M_INT_COL��¼״̬) Mod 3 = 2 Then
                    .ToolTipText = "��������"
                    If mnuEditStrike.Visible = True Then
                        mnuEditStrike.Enabled = False
                        tlbTool.Buttons("Strike").Enabled = False
                    End If
                End If
                If mnuEditModify.Visible = True Then
                    mnuEditModify.Enabled = False
                    tlbTool.Buttons("Modify").Enabled = False
                End If
                If mnuEditDel.Visible = True Then
                    mnuEditDel.Enabled = False
                    tlbTool.Buttons("Delete").Enabled = False
                End If
                If mnuEditVerify.Visible = True Then
                    mnuEditVerify.Enabled = False
                    tlbTool.Buttons("Verify").Enabled = False
                End If
                 
                If mnuEditDisplay.Visible = True Then
                    mnuEditDisplay.Enabled = True
                End If
            End If
        End If
    End With
    Cmd����.Enabled = mnuEditDisplay.Enabled
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
    objPrint.Title.Font.Name = "���� & _GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    
    objPrint.Title.Text = mstrTitle
        
    objRow.Add "ʱ�䣺" & strRange
    objRow.Add "���ţ�" & cboStock.Text
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
        
    objRow.Add "��ӡ��:" & UserInfo.�û���
    objRow.Add "��ӡ����:" & Format(sys.Currentdate, "yyyy��MM��dd��")
    objPrint.BelowAppRows.Add objRow
    
    Set objPrint.Body = mshlist
    
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

Private Sub subExcel(bytMode As Byte)
'����:���������EXCEL
'����:bytMode3 �����EXCEL

    Dim objPrint As Object
    Dim objRow As zlTabAppRow
    Dim strRange As String
    
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "���� & _GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    
    objPrint.Title.Text = mstrTitle
        
    Set objRow = New zlTabAppRow
    objRow.Add ""
    objRow.Add "NO." & Trim(mshlist.TextMatrix(mshlist.Row, GetCol(mshlist, "NO")))
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "�̵�ⷿ��" & Trim(cboStock.Text)
    objRow.Add "�̵�ʱ�䣺" & Trim(mshlist.TextMatrix(mshlist.Row, GetCol(mshlist, "�̵�ʱ��")))
    objPrint.UnderAppRows.Add objRow
        
    Set objRow = New zlTabAppRow
    objRow.Add "ժҪ:" & mshlist.TextMatrix(mshlist.Row, GetCol(mshlist, "ժҪ"))
    objPrint.BelowAppRows.Add objRow
        
    Set objRow = New zlTabAppRow
    objRow.Add "������:" & mshlist.TextMatrix(mshlist.Row, GetCol(mshlist, "������")) & "  ��������:" & mshlist.TextMatrix(mshlist.Row, GetCol(mshlist, "��������"))
    
    objRow.Add "�����:  " & "  �������:  "
    objPrint.BelowAppRows.Add objRow
    
    Set objPrint.Body = mshDetail
    zlPrintOrView1Grd objPrint, bytMode
End Sub

Private Sub tlbTool_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
    Case "Auto"
        Call mnuEditAddTableAuto_Click
    Case "Total"
        Call mnuEditAddTableTotal_Click
    Case "Zero"
        Call mnuEditAddTableZero_Click
    End Select
End Sub

Private Sub tlbTool_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu mnuViewTool
    End If
End Sub

'�Ե���ͷ������
Private Sub ListSort()
    Dim intCol As Integer
    Dim intRow As Integer
    Dim intTemp As String
    
    With mshlist
        If .Rows > 1 Then
            .Redraw = False
            intCol = .MouseCol
            .Col = intCol
            .ColSel = intCol
            intTemp = .TextMatrix(.Row, M_INT_COLNO)

            If intCol = mintPreCol And mintsort = flexSortStringNoCaseDescending Then
                .Sort = flexSortStringNoCaseAscending
                mintsort = flexSortStringNoCaseAscending
            Else
               .Sort = flexSortStringNoCaseDescending
               mintsort = flexSortStringNoCaseDescending
            End If

            mintPreCol = intCol
            .Row = Grid.MshGrdFindRow(mshlist, intTemp, 0)
            If .RowPos(.Row) + .RowHeight(.Row) > .Height Then
                .TopRow = .Row
            Else
                .TopRow = 1
            End If
            .Col = 0
            .ColSel = .Cols - 1
            If .Row = 0 Then
                .Row = 1
            End If
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
                Case 7, 8, 10, 11, 12, 13
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
            .Row = Grid.MshGrdFindRow(mshDetail, intTemp, 0)
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

Private Sub PrintRange(ByVal strRange As String)
    '����:��ӡʱ�䷶Χ
    picSeparate_s.Cls
    picSeparate_s.CurrentX = 50
    picSeparate_s.CurrentY = 100
    picSeparate_s.Print strRange
End Sub





Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

