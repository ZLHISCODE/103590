VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Begin VB.Form frmMainList 
   Caption         =   "ҩƷЭ��������"
   ClientHeight    =   4980
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9480
   Icon            =   "frmMainList.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picColor 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   4920
      ScaleHeight     =   255
      ScaleWidth      =   3615
      TabIndex        =   8
      Top             =   4320
      Width           =   3615
      Begin VB.PictureBox picColor3 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2280
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   11
         Top             =   0
         Width           =   260
      End
      Begin VB.PictureBox picColor2 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1320
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   10
         Top             =   0
         Width           =   260
      End
      Begin VB.PictureBox picColor1 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   9
         Top             =   0
         Width           =   260
      End
      Begin VB.Label lblColor3 
         AutoSize        =   -1  'True
         Caption         =   "�������"
         Height          =   180
         Left            =   2640
         TabIndex        =   14
         Top             =   30
         Width           =   720
      End
      Begin VB.Label lblColor1 
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Height          =   180
         Left            =   360
         TabIndex        =   13
         Top             =   37
         Width           =   720
      End
      Begin VB.Label lblColor2 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   1680
         TabIndex        =   12
         Top             =   37
         Width           =   360
      End
   End
   Begin VB.PictureBox picSeparate_s 
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   120
      MousePointer    =   7  'Size N S
      ScaleHeight     =   300
      ScaleWidth      =   4815
      TabIndex        =   4
      Top             =   2760
      Width           =   4815
      Begin VB.Label lblRange 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ѯ��Χ:1999��8��12����1999��9��12��"
         Height          =   180
         Left            =   0
         TabIndex        =   5
         Top             =   120
         Width           =   3690
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
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   3000
      End
      Begin MSComctlLib.Toolbar tlbTool 
         Height          =   720
         Left            =   165
         TabIndex        =   1
         Top             =   30
         Width           =   7785
         _ExtentX        =   13732
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
            NumButtons      =   15
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
               Key             =   "EditSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "���"
               Key             =   "Verify"
               Description     =   "���"
               Object.ToolTipText     =   "���"
               Object.Tag             =   "���"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Strike"
               Description     =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "VerifySeparate"
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Search"
               Description     =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ˢ��"
               Key             =   "Refresh"
               Description     =   "ˢ��"
               Object.ToolTipText     =   "ˢ��"
               Object.Tag             =   "ˢ��"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "FindSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Description     =   "��������"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Exit"
               Description     =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageIndex      =   11
            EndProperty
         EndProperty
         MouseIcon       =   "frmMainList.frx":014A
      End
   End
   Begin MSComctlLib.StatusBar staThis 
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
            Picture         =   "frmMainList.frx":0464
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
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":0CF8
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":0F18
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":1138
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":1354
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":1574
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":1794
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":19B0
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":1BCC
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":1DE6
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":1F40
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":215C
            Key             =   "Quit"
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
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":237C
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":259C
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":27BC
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":29D8
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":2BF8
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":2E18
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":3034
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":3250
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":346A
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":35C4
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":37E4
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfList 
      Height          =   1005
      Left            =   120
      TabIndex        =   6
      Top             =   1200
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
      FormatString    =   $"frmMainList.frx":3A04
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
   Begin VSFlex8Ctl.VSFlexGrid vsfDetail 
      Height          =   975
      Left            =   120
      TabIndex        =   7
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
      FormatString    =   $"frmMainList.frx":3A79
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
         Shortcut        =   ^A
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
Attribute VB_Name = "frmMainList"
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
Private mstrPrivs As String             '��ǰ�û����еĵ�ǰģ��Ĺ���
Private mblnViewCost As Boolean         '�鿴�ɱ���Ȩ�� true-�ܲ鿴�ɱ��� flase-���ܲ鿴�ɱ���

Private strStart As Date
Private strEnd As Date
Private strVerifyStart As Date
Private strVerifyEnd As Date

Private Type Type_SQLCondition
    strNO��ʼ As String
    strNO���� As String
    date����ʱ�俪ʼ As Date
    date����ʱ����� As Date
    date���ʱ�俪ʼ As Date
    date���ʱ����� As Date
    lngҩƷ As Long
    lng�Ƴ��ⷿ As Long
    str������ As String
    str����� As String
End Type

Private SQLCondition As Type_SQLCondition

Private mlng�ⷿid As Long
Private mintUnit As Integer                 '��λϵ����1-�ۼ�;2-����;3-סԺ;4-ҩ��
Private mint����� As Integer                       '��ʾҩƷ����ʱ�Ƿ���п���飺0-�����;1-��飬�������ѣ�2-��飬�����ֹ

'�Ӳ�������ȡҩƷ�۸����������С��λ������ʾ���ȣ�
Private mintShowCostDigit As Integer            '�ɱ���С��λ��
Private mintShowPriceDigit As Integer           '�ۼ�С��λ��
Private mintShowNumberDigit As Integer          '����С��λ��
Private mintShowMoneyDigit As Integer           '���С��λ��

Private mstrNumberFormat As String
Private mstrCostFormat As String
Private mstrPriceFormat As String
Private mstrMoneyFormat As String

Private Const mconint�ۼ۵�λ As Integer = 1
Private Const mconint���ﵥλ As Integer = 2
Private Const mconintסԺ��λ As Integer = 3
Private Const mconintҩ�ⵥλ As Integer = 4
Private Sub cboStock_Click()
    If mlng�ⷿid <> cboStock.ItemData(cboStock.ListIndex) Then
        mlng�ⷿid = cboStock.ItemData(cboStock.ListIndex)
        Call GetDrugDigit(mlng�ⷿid, Me.Tag, mintUnit, mintShowCostDigit, mintShowPriceDigit, mintShowNumberDigit, mintShowMoneyDigit)
        
        '��֯��ʽ����
        mstrCostFormat = "'999999999990." & String(mintShowCostDigit, "0") & "'"
        mstrPriceFormat = "'999999999990." & String(mintShowPriceDigit, "0") & "'"
        mstrNumberFormat = "'999999999990." & String(mintShowNumberDigit, "0") & "'"
        mstrMoneyFormat = "'999999999990." & String(mintShowMoneyDigit, "0") & "'"
        
        If mblnBootUp Then mnuViewRefresh_Click
    End If
End Sub

Private Sub cbrTool_Resize()
    If mblnBootUp = False Then Exit Sub
    Form_Resize
End Sub

Public Sub ShowList(ByVal lngMode As Long, ByVal strTitle As String, ByVal FrmMain As Variant)
    Dim strFind As String
    Dim dateCurDate As Date
    Dim strTemp As String
    Dim dateCurrentDate As Date
    
    mblnBootUp = False
    mlngMode = lngMode
    mstrTitle = strTitle
    mstrPrivs = gstrprivs
    Me.Tag = strTitle
    
    If Not CheckDepend Then     '���������Բ���
        Unload Me
        Exit Sub
    End If
    
    mlng�ⷿid = Me.cboStock.ItemData(Me.cboStock.ListIndex)
    Call GetDrugDigit(mlng�ⷿid, Me.Tag, mintUnit, mintShowCostDigit, mintShowPriceDigit, mintShowNumberDigit, mintShowMoneyDigit)
    
    '��֯��ʽ����
    mstrCostFormat = "'999999999990." & String(mintShowCostDigit, "0") & "'"
    mstrPriceFormat = "'999999999990." & String(mintShowPriceDigit, "0") & "'"
    mstrNumberFormat = "'999999999990." & String(mintShowNumberDigit, "0") & "'"
    mstrMoneyFormat = "'999999999990." & String(mintShowMoneyDigit, "0") & "'"
    
    SetVisable  '����Ȩ�����ò�ͬ����ʾ��Ŀ
    
    dateCurDate = Sys.Currentdate()
    strStart = Format(DateAdd("d", -7, dateCurDate), "yyyy-MM-dd")
    strEnd = Format(dateCurDate, "yyyy-MM-dd")
    
    strVerifyStart = "1901-01-01"
    strVerifyEnd = "1901-01-01"
    
    strFind = " AND A.��¼״̬ = 1 And A.������� is Null And A.�������� Between [3] And [4] "
    SQLCondition.date����ʱ�俪ʼ = CDate(Format(strStart, "yyyy-mm-dd") & " 00:00:00")
    SQLCondition.date����ʱ����� = CDate(Format(strEnd, "yyyy-mm-dd") & " 23:59:59")
        
    mstrFind = strFind
    
    GetList (mstrFind)  '�г�����ͷ
    RestoreWinState Me, App.ProductName, mstrTitle
    '�ָ����Ի����ú�Ȩ�޿�����ʾ������Ҫ��һ������
    With vsfDetail
        If mblnViewCost = False Then
            .ColHidden(.ColIndex("����")) = True
            .ColHidden(.ColIndex("���۽��")) = True
            .ColHidden(.ColIndex("���")) = True
        End If
    End With
    
    Call zldatabase.ShowReportMenu(Me, glngSys, mlngMode, mstrPrivs)
    
    mblnBootUp = True
    
    If IsObject(FrmMain) Then
        Me.Show , FrmMain
    Else
        'ZLBH�ں���ʾ
        OS.ShowChildWindow Me.hWnd, FrmMain
    End If
    Me.ZOrder 0
End Sub

'�������������
Private Function CheckDepend() As Boolean
    
    Dim rsDepend As New Recordset
    Dim strStock As String
    
    CheckDepend = False
    On Error GoTo errHandle
    Select Case mlngMode
        Case 1344
            strStock = "LMN"
        
        Case 1343
            strStock = "HIJKLMN"
        Case Else
            
            
    End Select

    gstrSQL = "SELECT DISTINCT a.id, a.���� " _
         & "FROM ��������˵�� c, �������ʷ��� b, ���ű� a " _
        & "Where (a.վ�� = '" & gstrNodeNo & "' Or a.վ�� is Null) And c.�������� = b.���� " _
          & "AND Instr([1],b.����,1) > 0 " _
         & " AND a.id = c.����id " _
          & "AND TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01'"
    Set rsDepend = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, strStock)
    
    If rsDepend.EOF Then
        MsgBox "����������Ϣ��ȫ,��鿴���Ź���", vbInformation, gstrSysName
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
            If Not IsHavePrivs(gstrprivs, "����ҩ��") Then
                MsgBox "�㲻��ҩ��������Ա�Ҳ��������пⷿ��Ȩ�ޣ����ܽ��룡", vbInformation, gstrSysName
                Unload Me
                Exit Function
            End If
            .ListIndex = 0
        ElseIf Not IsHavePrivs(gstrprivs, "����ҩ��") Then
            .Enabled = False
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
    Dim rsList As New Recordset
    Dim strUserPart As String
    Dim n As Integer
    
    On Error GoTo errHandle
    mlastRow = 0
'    strUserPart = " And A.�ⷿID+0=" & cboStock.ItemData(cboStock.ListIndex)
    strUserPart = " And A.�ⷿID+0=[11] "
    vsfList.Redraw = flexRDNone
    Select Case mlngMode
        Case 1344           'Э��ҩƷ������
            gstrSQL = "SELECT /*+ Rule*/ a.no, a.������, " _
                  & "TO_CHAR (a.��������, 'yyyy-mm-dd HH24:Mi:SS') AS ��������, a.�����, " _
                  & "TO_CHAR (a.�������, 'yyyy-mm-dd HH24:Mi:SS') AS �������, a.��¼״̬, a.ժҪ " _
             & "FROM ҩƷ�շ���¼ a, ���ű� b " _
            & "Where a.�ⷿid = b.ID AND a.���� = 3 and ���ϵ��=1 and  a.���=1  " _
                        & strUserPart & strFind _
           & "ORDER BY no DESC, �������� ASC "
       
    End Select
    
'    Call SQLTest(App.Title, Me.Caption, gstrSQL)
'    rsList.Open gstrSQL, gcnOracle
'    Call SQLTest
    
    Set rsList = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
        SQLCondition.strNO��ʼ, _
        SQLCondition.strNO����, _
        SQLCondition.date����ʱ�俪ʼ, _
        SQLCondition.date����ʱ�����, _
        SQLCondition.date���ʱ�俪ʼ, _
        SQLCondition.date���ʱ�����, _
        SQLCondition.lngҩƷ, _
        SQLCondition.lng�Ƴ��ⷿ, _
        SQLCondition.str������, _
        SQLCondition.str�����, _
        cboStock.ItemData(cboStock.ListIndex))
   
    Set vsfList.DataSource = rsList
    With vsfList
        If .rows = 1 Then
            .rows = .rows + 100
            .Row = 1
            .Redraw = flexRDDirect
            
            .TopRow = 1
            .rows = .rows - 99
            
        End If
        .Row = 1
        .Col = 0
        .ColSel = .Cols - 1
        
        For n = 0 To .Cols - 1
            .FixedAlignment(n) = flexAlignCenterCenter
        Next
    End With
    SetListColWidth
    vsfList.ColWidth(vsfList.Cols - 2) = 0
    vsfList_EnterCell    '�г�������
    
    SetStrikeColor
    vsfList.Redraw = flexRDDirect
    
    staThis.Panels(2).Text = "��ǰ����" & rsList.RecordCount & "�ŵ���"
    rsList.Close
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetStrikeColor()
    Dim intStatus As Integer
    Dim intRow As Integer
    Dim intCol As Integer
    
    With vsfList
        If .rows <= 2 Then Exit Sub
        For intRow = 1 To .rows - 1
            intStatus = .TextMatrix(intRow, .Cols - 2)
            If intStatus = 3 Then
                .Cell(flexcpForeColor, intRow, 0, intRow, .Cols - 1) = &H80000001 ' &H80000018
            End If
            If intStatus = 2 Then
                .Cell(flexcpForeColor, intRow, 0, intRow, .Cols - 1) = &HFF '&H80000001
            End If
        Next
    End With
                
End Sub

'��ͷ�п��ʼ
Private Sub SetListColWidth()
    Dim intCol As Integer
    
    With vsfList
        If mblnBootUp = False Then
            For intCol = 1 To .Cols - 1
                If intCol = 1 Then
                    .ColWidth(intCol) = 2000
                ElseIf intCol = .Cols - 2 Then
                    .ColWidth(intCol) = 0
                Else
                    .ColWidth(intCol) = 1000
                End If
            Next
        End If
            
    End With
End Sub


Private Sub SetDetailColWidth()
    Dim intCol As Integer
    Dim rsDetail As New Recordset
    Dim bln��ҩ�ⷿ As Boolean
    Dim str�ⷿ���� As String
    
    On Error GoTo errHandle
    str�ⷿ���� = ""
    gstrSQL = "Select a.�������� From ��������˵�� A Where a.����id =[1]"
    Set rsDetail = zldatabase.OpenSQLRecord(gstrSQL, "�ж��ǿⷿ����", cboStock.ItemData(cboStock.ListIndex))
    Do While Not rsDetail.EOF
        str�ⷿ���� = str�ⷿ���� & "," & rsDetail!��������
        rsDetail.MoveNext
    Loop
    If str�ⷿ���� Like "*��ҩ*" Or str�ⷿ���� Like "*�Ƽ���*" Then bln��ҩ�ⷿ = True
        
    With vsfDetail
        Select Case mlngMode
            Case 1344
                .ColAlignment(.ColIndex("����")) = flexAlignRightCenter     '����
                .ColAlignment(.ColIndex("��λ")) = flexAlignCenterCenter    '��λ
                .ColAlignment(.ColIndex("����")) = flexAlignRightCenter     '����
                .ColAlignment(.ColIndex("���۽��")) = flexAlignRightCenter '���۽��
                .ColAlignment(.ColIndex("�ۼ�")) = flexAlignRightCenter     '�ۼ�
                .ColAlignment(.ColIndex("�ۼ۽��")) = flexAlignRightCenter '�ۼ۽��
                .ColAlignment(.ColIndex("���")) = flexAlignRightCenter     '���
        End Select
        
        If mblnBootUp = False Then
            .ColWidth(0) = 0
            .ColWidth(.ColIndex("ҩƷ��Ϣ")) = 2500
            For intCol = 2 To .Cols - 1
                .ColWidth(intCol) = 1000
            Next
        End If
        If mblnViewCost = False Then
            .ColHidden(.ColIndex("����")) = True
            .ColHidden(.ColIndex("���۽��")) = True
            .ColHidden(.ColIndex("���")) = True
        End If
        
        If bln��ҩ�ⷿ Then
            .ColHidden(.ColIndex("ԭ����")) = False
        Else
            .ColHidden(.ColIndex("ԭ����")) = True
        End If
    End With
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'����Ȩ�����ò�ͬ����ʾ��Ŀ
Private Sub SetVisable()
    '�⹺�������Ȩ�ޣ��������á����������пⷿ���Ǽǡ��޸ġ�ɾ�������ա�����

    Select Case mlngMode
        Case 1344
            If Not IsHavePrivs(gstrprivs, "�Ǽ�") Then
                mnuEditAdd.Visible = False
                tlbTool.Buttons("Add").Visible = False
            End If
            
            If Not IsHavePrivs(gstrprivs, "�޸�") Then
                mnuEditModify.Visible = False
                tlbTool.Buttons("Modify").Visible = False
            End If
            
            If Not IsHavePrivs(gstrprivs, "ɾ��") Then
                mnuEditDel.Visible = False
                tlbTool.Buttons("Delete").Visible = False
                 '��û�����б༭Ȩ��ʱ���Ѳ˵��͹������ϵ���Ӧ�ķָ������Ρ�
                If mnuEditAdd.Visible = False And mnuEditModify.Visible = False Then
                    mnuEditLine1.Visible = False
                    tlbTool.Buttons("EditSeparate").Visible = False
                End If
            End If
            
            If Not IsHavePrivs(gstrprivs, "���") Then
                mnuEditVerify.Visible = False
                tlbTool.Buttons("Verify").Visible = False
            End If
            
            If Not IsHavePrivs(gstrprivs, "����") Then
                mnuEditStrike.Visible = False
                tlbTool.Buttons("Strike").Visible = False
                
                If mnuEditVerify.Visible = False Then
                    mnuEditLine2.Visible = False
                    tlbTool.Buttons("VerifySeparate").Visible = False
                End If
            End If
            If Not IsHavePrivs(gstrprivs, "���ݴ�ӡ") Then
               mnuFileBillPrint.Visible = False
               mnuFileBillPreview.Visible = False
            End If
        Case 1301
        
        Case Else
        
    End Select
    
End Sub

Private Sub Form_Load()
    Dim dateCurDate As Date
    
    mblnViewCost = IsHavePrivs(mstrPrivs, "�鿴�ɱ���")
    '�ָ�����
    dateCurDate = Sys.Currentdate()
    lblRange.Caption = "��ѯ��Χ:" & Format(dateCurDate, "yyyy��MM��dd��") & "��" & Format(dateCurDate, "yyyy��MM��dd��")
    
    staThis.Panels(2).Picture = picColor
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
        .Height = 300
        .Left = 0
        .Width = cbrTool.Width
        
    End With
    
    With vsfList
        .Top = IIf(cbrTool.Visible, cbrTool.Height, 0)
        .Left = 0
        .Width = cbrTool.Width
        .Height = picSeparate_s.Top - .Top
    End With
    
    With vsfDetail
        .Top = picSeparate_s.Top + picSeparate_s.Height + 100
        .Left = 0
        .Height = ScaleHeight - .Top - IIf(staThis.Visible, staThis.Height, 0)
        .Width = cbrTool.Width
    End With
        
    With picColor
        .Top = Me.ScaleHeight - .Height - 30
        .Left = Me.ScaleWidth - staThis.Panels(3).Width - staThis.Panels(4).Width - .Width - 300
    End With
    
    If mlngMode <> 1343 Then
        picColor3.Visible = False
        lblColor3.Visible = False
        picColor.Width = lblColor2.Left + lblColor2.Width + 20
    Else
        lblColor3.Caption = "δ��˳���"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName, mstrTitle
End Sub

Private Sub mnuEditAdd_Click()
    Dim strNo As String
    Dim BlnSuccess As Boolean
    
    '��鱾���Ƿ��Ѿ���˽�棬���δ��˽�����ܽ�������ҵ�����
'    If CheckIsAccount(Val(cboStock.ItemData(cboStock.ListIndex))) = False Then
'        Exit Sub
'    End If
    
    strNo = ""
    '����
    Select Case mlngMode
        Case 1344
            frmAccordDrugCard.ShowCard Me, strNo, 1, , BlnSuccess
        Case 1306
            '��frmtputCard.ShowCard Me, strNO, 1, , blnSuccess
    End Select
    
    If BlnSuccess = True Then
        mnuViewRefresh_Click
    End If
End Sub

Private Sub MnuEditVerify_Click()
    '����
    Dim strNo As String
    Dim BlnSuccess As Boolean
    
    With vsfList
        strNo = .TextMatrix(.Row, 0)
        Select Case mlngMode
            Case 1344
                frmAccordDrugCard.ShowCard Me, strNo, 3, .TextMatrix(.Row, .Cols - 2), BlnSuccess
            Case 1343
                'frmSelfMakeCard.ShowCard Me, StrNo, 3, .TextMatrix(.Row, .Cols - 2), BlnSuccess
            
        End Select
        
    End With
    If BlnSuccess = True Then
        mnuViewRefresh_Click
    End If
End Sub

Private Sub mnuEditDel_Click()
    'ɾ��
    Dim StrBillNo As String
    Dim intRow As Integer
    Dim strTitle As String
    Dim intReturn As Integer
    Dim intRecord As Integer
    
    With vsfList
        Select Case mlngMode
            Case 1344
                strTitle = "Э��ҩƷ��ⵥ"
            Case Else
                
        End Select
        
        On Error GoTo errHandle
        intRow = .Row
        StrBillNo = .TextMatrix(intRow, 0)
        intReturn = MsgBox("��ȷʵҪɾ�����ݺ�Ϊ��" & StrBillNo & "����" & strTitle & "��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)
        intRecord = .rows - 1
        If intReturn = vbYes Then
            Select Case mlngMode
                Case 1344
                    gstrSQL = "zl_Э�����_Delete('" & StrBillNo & "')"
            End Select
            If gstrSQL = "" Then Exit Sub
            Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            
            intRecord = intRecord - 1
            
            mlastRow = 0
            If .rows > 2 Then
                .RemoveItem intRow
            ElseIf .rows = 2 Then
                .rows = 3
                .RemoveItem intRow
                With vsfDetail
                    .rows = 1
                    .rows = 2
                    .FixedRows = 1
                    .Row = 1
                    .Col = 0
                    .ColSel = .Cols - 1
                End With
                SetEnable
                
            End If
                
            '.RowHeight(intRow) = 0
            If intRow < .rows - 1 Then
                .Row = intRow
            Else
                If .rows = 2 Then
                    .Row = 1
                Else
                    .Row = intRow - 1
                End If
            End If
            .Col = 0
            .ColSel = .Cols - 1
            vsfList_EnterCell
        End If
    End With
    staThis.Panels(2).Text = "��ǰ����" & intRecord & "�ŵ���"
    Exit Sub
    
errHandle:
    If ErrCenter() = 1 Then Resume 'Resume����������õ���
    Call SaveErrLog
    
End Sub

Private Sub mnuEditDisplay_Click()
    '�鿴����
    
    Dim strNo As String
    With vsfList
        strNo = .TextMatrix(.Row, 0)
        Select Case mlngMode
            Case 1344
                frmAccordDrugCard.ShowCard Me, strNo, 4, .TextMatrix(.Row, .Cols - 2)
            Case 1343
                'frmOtherInputCard.ShowCard Me, StrNo, 4, .TextMatrix(.Row, .Cols - 2)
        
        End Select
        
    End With
    
End Sub

Private Sub mnuEditStrike_Click()
    '����
    
    With vsfList
        If MsgBox("��ȷʵҪȫ���������ݺ�Ϊ��" & .TextMatrix(.Row, 0) & "���ĵ�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            If StrikeSave = True Then
                mnuViewRefresh_Click
            End If
        End If
    End With
End Sub

Private Function StrikeSave() As Boolean
    Dim n As Integer
    
    StrikeSave = False
    With vsfList
        Select Case mlngMode
            Case 1344
                mint����� = MediWork_GetCheckStockRule(cboStock.ItemData(cboStock.ListIndex))
                
                '�����������Ƿ��㹻����������Ϊ�������ʱ�����У����뵥�ݣ�ҩ��������飬���ݺţ���ţ�
                If mint����� <> 0 And .TextMatrix(.Row, 0) <> "" Then
                    For n = 1 To vsfDetail.rows - 1
                        If vsfDetail.TextMatrix(n, 0) <> "" Then
                            If CheckStrickUsable(3, 0, 0, vsfDetail.TextMatrix(n, 1), _
                                0, 0, mint�����, Trim(.TextMatrix(.Row, 0)), Val(vsfDetail.TextMatrix(n, 0))) = False Then
                                Exit Function
                            End If
                        End If
                    Next
                End If
                
                gstrSQL = "zl_Э�����_Strike('" & .TextMatrix(.Row, 0) & "','" & .TextMatrix(.Row, .Cols - 4) & "')"
            
        End Select
        
        On Error GoTo errHandle
        
        Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        
        '��ʾͣ��ҩƷ
        Call CheckStopMedi(���ݺ�.Эҩ��� & "|" & .TextMatrix(.Row, 0))
    End With
    StrikeSave = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

    'MsgBox "����ʧ�ܣ�", vbInformation, gstrSysName
End Function

Private Sub mnuEditModify_Click()
    '�޸�
    Dim strNo As String
    Dim BlnSuccess As Boolean
    
    BlnSuccess = False
    With vsfList
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        strNo = .TextMatrix(.Row, 0)
        
        Select Case mlngMode
            Case 1344
                frmAccordDrugCard.ShowCard Me, strNo, 2, vsfList.TextMatrix(vsfList.Row, vsfList.Cols - 2), BlnSuccess
            Case 1302
                
        End Select
        If BlnSuccess = True Then
            mnuViewRefresh_Click
        End If
    End With
End Sub

Private Sub mnuFileBillPreview_Click()
    Dim int��λϵ�� As Integer
    
    With vsfList
        If .TextMatrix(.Row, 0) = "" Then Exit Sub

        Select Case mintUnit
            Case mconint�ۼ۵�λ
                int��λϵ�� = 4
            Case mconint���ﵥλ
                int��λϵ�� = 2
            Case mconintסԺ��λ
                int��λϵ�� = 1
            Case mconintҩ�ⵥλ
                int��λϵ�� = 3
        End Select
    
        ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1344", "zl8_bill_1344"), Me, "���ݱ��=" & .TextMatrix(.Row, 0), "��¼״̬=" & .TextMatrix(.Row, .Cols - 2), "��λϵ��=" & int��λϵ��, 1
    End With
End Sub

Private Sub MnuFileBillprint_Click()
    Dim int��λϵ�� As Integer
    
    With vsfList
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        
        Select Case mintUnit
            Case mconint�ۼ۵�λ
                int��λϵ�� = 4
            Case mconint���ﵥλ
                int��λϵ�� = 2
            Case mconintסԺ��λ
                int��λϵ�� = 1
            Case mconintҩ�ⵥλ
                int��λϵ�� = 3
        End Select
        ReportOpen gcnOracle, glngSys, "zl1_bill_1344", Me, "���ݱ��=" & .TextMatrix(.Row, 0), "��¼״̬=" & .TextMatrix(.Row, .Cols - 2), "��λϵ��=" & int��λϵ��, 2
    End With
End Sub

Private Sub mnuFileExcel_Click()
    '�����Excel
    If Me.ActiveControl Is vsfList Then
        vsfList.Redraw = flexRDNone
        subPrint 3
        vsfList.Redraw = flexRDDirect
        vsfList.Col = 0
        vsfList.ColSel = vsfList.Cols - 1
    ElseIf Me.ActiveControl Is vsfDetail Then
        vsfDetail.Redraw = flexRDNone
        subExcel 3
        vsfDetail.Redraw = flexRDDirect
        vsfDetail.Col = 0
        vsfDetail.ColSel = vsfDetail.Cols - 1
    End If
End Sub

Private Sub mnufileexit_Click()
    '�˳�
    Unload Me
    
End Sub

Private Sub mnuFileParameter_Click()
    '��������
    frm��������.���ò��� Me, mstrPrivs, 1344, Me.Tag
    Call GetList(mstrFind)
End Sub

Private Sub mnuFilePreView_Click()
    '��ӡԤ��
    vsfList.Redraw = flexRDNone
    subPrint 2
    vsfList.Redraw = flexRDDirect
    vsfList.Col = 0
    vsfList.ColSel = vsfList.Cols - 1
End Sub

Private Sub mnuFilePrint_Click()
    '��ӡ

    vsfList.Redraw = flexRDNone
    subPrint 1
    vsfList.Redraw = flexRDDirect
    vsfList.Col = 0
    vsfList.ColSel = vsfList.Cols - 1
    
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
'    ReportMan gcnOracle, Me
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub mnuHelpWebHome_Click()
    '������ҳ
    Call zlHomePage(Me.hWnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    '���ͷ���
    Call zlMailTo(Me.hWnd)
End Sub

Private Sub mnuReportItem_Click(index As Integer)
    'Ĭ�ϲ�����ҩƷ=ҩƷid���ⷿ=�ⷿid����ʼʱ��=���ƿ�ʼʱ�䣬����ʱ��=���ƽ���ʱ�䣬NO=��ⵥNO
    Dim str��ʼʱ�� As String
    Dim str����ʱ�� As String
    Dim strNo As String
    
    If vsfList.Row >= 1 And LTrim(vsfList.TextMatrix(vsfList.Row, 0)) <> "" Then
        strNo = vsfList.TextMatrix(vsfList.Row, 0)
    End If
    
    str��ʼʱ�� = IIf(Format(SQLCondition.date����ʱ�俪ʼ, "yyyy-mm-dd") = "1899-12-30", "", Format(SQLCondition.date����ʱ�俪ʼ, "yyyy-mm-dd"))
    str����ʱ�� = IIf(Format(SQLCondition.date����ʱ�����, "yyyy-mm-dd") = "1899-12-30", "", Format(SQLCondition.date����ʱ�����, "yyyy-mm-dd"))
    
    Call ReportOpen(gcnOracle, Split(mnuReportItem(index).Tag, ",")(0), Split(mnuReportItem(index).Tag, ",")(1), Me, _
        "ҩƷ=" & IIf(SQLCondition.lngҩƷ = 0, "", SQLCondition.lngҩƷ), _
        "�ⷿ=" & Val(cboStock.ItemData(cboStock.ListIndex)), _
        "��ʼʱ��=" & str��ʼʱ��, _
        "����ʱ��=" & str����ʱ��, _
        "NO=" & strNo)
End Sub
Private Sub mnuViewRefresh_Click()
    'ˢ��
    GetList mstrFind
End Sub

Private Sub mnuViewSearch_Click()
    '����
    Dim strFind As String
    
    Select Case mlngMode
        Case 1344
            strFind = FrmListSearch.GetSearch(Me, mlngMode, cboStock.ItemData(cboStock.ListIndex), strStart, strEnd, strVerifyStart, strVerifyEnd, _
                SQLCondition.strNO��ʼ, _
                SQLCondition.strNO����, _
                SQLCondition.date����ʱ�俪ʼ, _
                SQLCondition.date����ʱ�����, _
                SQLCondition.date���ʱ�俪ʼ, _
                SQLCondition.date���ʱ�����, _
                SQLCondition.lngҩƷ, _
                SQLCondition.lng�Ƴ��ⷿ, _
                SQLCondition.str������, _
                SQLCondition.str�����)
        Case Else
    End Select
    '������ģ�������ʱ�䱣�浽ע�����
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & mlngMode, "Ĭ�ϲ�ѯʱ��", SQLCondition.date����ʱ�俪ʼ & "|" & SQLCondition.date����ʱ�����)
    
    If strFind <> "" Then
        mstrFind = strFind
        GetList mstrFind
        If Format(strStart, "yyyy-mm-dd") = "1901-01-01" And Format(strVerifyStart, "yyyy-mm-dd") = "1901-01-01" Then
            lblRange.Visible = False
        ElseIf Format(strStart, "yyyy-mm-dd") <> "1901-01-01" And Format(strVerifyStart, "yyyy-mm-dd") <> "1901-01-01" Then
            lblRange.Visible = True
            lblRange = "��ѯ��Χ:�������� " & Format(strStart, "yyyy��MM��dd��") & "��" & Format(strEnd, "yyyy��MM��dd��") & "  ������� " & Format(strVerifyStart, "yyyy��MM��dd��") & "��" & Format(strVerifyEnd, "yyyy��MM��dd��")
        ElseIf Format(strStart, "yyyy-mm-dd") <> "1901-01-01" Then
            lblRange.Visible = True
            lblRange = "��ѯ��Χ:�������� " & Format(strStart, "yyyy��MM��dd��") & "��" & Format(strEnd, "yyyy��MM��dd��")
        ElseIf Format(strVerifyStart, "yyyy-mm-dd") <> "1901-01-01" Then
            lblRange.Visible = True
            lblRange = "��ѯ��Χ:������� " & Format(strVerifyStart, "yyyy��MM��dd��") & "��" & Format(strVerifyEnd, "yyyy��MM��dd��")
        End If
             
    End If
    
End Sub

Private Sub mnuViewStatus_Click()
    With mnuViewStatus
        .Checked = Not .Checked  ' Xor True
        staThis.Visible = .Checked
    End With
    
    Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    With mnuViewToolButton
        .Checked = Not .Checked   ' Xor True
        cbrTool.Bands(1).Visible = .Checked
        mnuViewToolText.Enabled = .Checked
    End With
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim intCount As Integer      '����������
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked   ' Xor True
    With tlbTool.Buttons
        If mnuViewToolText.Checked = False Then
            'ȡ�����е��ı���ǩ��ʾ
            For intCount = 1 To .count
                .Item(intCount).Caption = ""
            Next
        Else
            '�����е��ı���ǩ��ʾ��˵����Tag�зŵ��ı���ǩ
            For intCount = 1 To .count
                .Item(intCount).Caption = .Item(intCount).Tag
            Next
        End If
    End With
    
    cbrTool.Bands(1).MinHeight = tlbTool.Height
    
    Form_Resize
End Sub

Private Sub vsfDetail_GotFocus()
    Call SetGridFocus(vsfDetail, True)
End Sub

Private Sub vsfDetail_LostFocus()
    Call SetGridFocus(vsfDetail, False)
End Sub

Private Sub vsfList_DblClick()
    If mnuEditModify.Visible = False Then Exit Sub
    If mnuEditModify.Enabled = False Then Exit Sub
    If vsfList.MouseRow = 0 Then Exit Sub
    mnuEditModify_Click
End Sub

Private Sub vsfList_EnterCell()
    Dim rsDetail As New Recordset
    Dim strUnitQuantity As String               '��λ��������ʽ����
    Dim IntBill As Integer                      '��������  �磺1���⹺��⣻2��
    Dim str��װϵ�� As String
    Dim strSqlҩ�� As String
    Dim intCol As Integer
    Dim n As Integer
    
    If mlastRow = vsfList.Row Then Exit Sub
    mlastRow = vsfList.Row
    
    On Error GoTo errHandle
    With vsfList
        .Redraw = flexRDNone
        .ForeColorSel = .Cell(flexcpForeColor, mlastRow, 1)
        .Redraw = flexRDDirect
    End With
    
    SetEnable
    
    If vsfList.Row >= 1 And LTrim(vsfList.TextMatrix(vsfList.Row, 0)) <> "" Then
        vsfList.Col = 0
        vsfList.ColSel = vsfList.Cols - 1
        
        vsfDetail.Redraw = flexRDNone
        
        Select Case mintUnit
            Case mconint�ۼ۵�λ
                strUnitQuantity = "LTRIM(to_char(A.ʵ������," & mstrNumberFormat & ")) AS ����," _
                    & "F.���㵥λ AS ��λ,"
                str��װϵ�� = "1"
            Case mconint���ﵥλ
                strUnitQuantity = "LTRIM(to_char(A.ʵ������ / B.�����װ," & mstrNumberFormat & ")) AS ����," _
                    & "B.���ﵥλ AS ��λ,"
                str��װϵ�� = "B.�����װ"
            Case mconintסԺ��λ
                strUnitQuantity = "LTRIM(to_char(A.ʵ������ / B.סԺ��װ," & mstrNumberFormat & ")) AS ����," _
                    & "B.סԺ��λ AS ��λ,"
                str��װϵ�� = "B.סԺ��װ"
            Case mconintҩ�ⵥλ
                strUnitQuantity = "LTRIM(to_char(A.ʵ������ / B.ҩ���װ," & mstrNumberFormat & ")) AS ����," _
                    & "B.ҩ�ⵥλ AS ��λ,"
                str��װϵ�� = "B.ҩ���װ"
        End Select
        
        If gintҩƷ������ʾ = 0 Then
            strSqlҩ�� = "('['||F.����||']'||F.����) AS ҩƷ��Ϣ"
        ElseIf gintҩƷ������ʾ = 1 Then
            strSqlҩ�� = "('['||F.����||']'||NVL(E.����,F.����)) AS ҩƷ��Ϣ"
        Else
            strSqlҩ�� = "('['||F.����||']'||F.����) AS ҩƷ��Ϣ,E.���� As ��Ʒ��"
        End If
        
        Select Case mlngMode
            Case 1344
                IntBill = 3
                gstrSQL = " SELECT * FROM " & _
                    "    (SELECT DISTINCT ���," & strSqlҩ�� & ",B.ҩƷ��Դ,B.����ҩ��,F.���,A.����,A.ԭ����," & _
                    strUnitQuantity & _
                    "    LTRIM(TO_CHAR (A.�ɱ���*" & str��װϵ�� & ", " & mstrCostFormat & ")) AS ����," & _
                    "    LTRIM(TO_CHAR (A.�ɱ����, " & mstrMoneyFormat & ")) AS ���۽��," & _
                    "    LTRIM(TO_CHAR (A.���ۼ�*" & str��װϵ�� & "," & mstrPriceFormat & ")) AS �ۼ�," & _
                    "    LTRIM(TO_CHAR (A.���۽��, " & mstrMoneyFormat & ")) AS �ۼ۽��," & _
                    "    LTRIM(TO_CHAR (A.���, " & mstrMoneyFormat & ")) AS ��� " & _
                    "    FROM ҩƷ�շ���¼ A, ҩƷ��� B,�շ���Ŀ���� E,�շ���ĿĿ¼ F " & _
                    "    WHERE A.ҩƷID = B.ҩƷID AND B.ҩƷID=F.ID" & _
                    "    AND B.ҩƷID = E.�շ�ϸĿID(+) AND E.����(+)=3 " & _
                    "    AND ��¼״̬ = [2] " & _
                    "    AND A.���� = 3 AND ���ϵ��=1 " & _
                    "    AND A.NO = [1])" & _
                    " ORDER BY ��� "
        End Select
        
        Set rsDetail = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, vsfList.TextMatrix(vsfList.Row, 0), Val(vsfList.TextMatrix(vsfList.Row, vsfList.Cols - 2)))
        
        Set vsfDetail.DataSource = rsDetail
        rsDetail.Close
        
        With vsfDetail
            .Row = 1
            .Col = 0
            .ColSel = .Cols - 1
            
            For n = 0 To .Cols - 1
                .ColKey(n) = .TextMatrix(0, n)
                .FixedAlignment(n) = flexAlignCenterCenter
            Next
        End With
        
        vsfDetail.Redraw = flexRDDirect
    ElseIf LTrim(vsfList.TextMatrix(vsfList.Row, 0)) = "" Then
        With vsfDetail
            .Redraw = flexRDNone
            Select Case mlngMode
                Case 1344
                    .Cols = IIf(gintҩƷ������ʾ = 2, 15, 14)
                    .rows = 2
                    .Clear
                    
                    intCol = 0
                    
                    .TextMatrix(0, intCol) = "���": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "ҩƷ��Ϣ": intCol = intCol + 1
                    
                    If gintҩƷ������ʾ = 2 Then
                        .TextMatrix(0, intCol) = "��Ʒ��": intCol = intCol + 1
                    End If
                    
                    .TextMatrix(0, intCol) = "ҩƷ��Դ": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "����ҩ��": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "���": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "������": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "ԭ����": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "����": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "��λ": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "����": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "���۽��": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "�ۼ�": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "�ۼ۽��": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "���": intCol = intCol + 1
                
            End Select
            .Row = 1
            .Col = 0
            .ColSel = .Cols - 1
            
            For n = 0 To .Cols - 1
                .ColKey(n) = .TextMatrix(0, n)
                .FixedAlignment(n) = flexAlignCenterCenter
            Next
            .Redraw = flexRDDirect
        End With
    End If
    SetDetailColWidth
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsfList_GotFocus()
    Call SetGridFocus(vsfList, True)
End Sub

Private Sub vsfList_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If mnuEditModify.Visible = False Then Exit Sub
        If mnuEditModify.Enabled = False Then Exit Sub
        mnuEditModify_Click
    End If
        
End Sub

Private Sub vsfList_LostFocus()
    Call SetGridFocus(vsfList, False)
End Sub

Private Sub vsfList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    If mnuEdit.Visible = False Then Exit Sub
    
    PopupMenu mnuEdit, 2
    
End Sub

Private Sub picSeparate_s_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    '�ָ�������
    
    If Button <> 1 Then Exit Sub
    
    With picSeparate_s
        If .Top + y < 2000 Then Exit Sub
        If .Top + y > ScaleHeight - 2000 Then Exit Sub
        .Move .Left, .Top + y
    End With
    
    With vsfList
        .Top = IIf(cbrTool.Visible, cbrTool.Height, 0)
        .Height = picSeparate_s.Top - .Top
    End With
    
    With vsfDetail
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
            mnuEditAdd_Click
        Case "Modify"
            mnuEditModify_Click
        Case "Delete"
            mnuEditDel_Click
        Case "Verify"
            MnuEditVerify_Click
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
    With vsfList
        .ToolTipText = ""
        If .TextMatrix(.Row, 0) = "" Or .Row = 0 Then         'û�е�
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
            
            
            If .TextMatrix(.Row, .Cols - 4) = "" Then    'δ��˵�
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
            ElseIf .TextMatrix(.Row, .Cols - 2) = 1 Then    '��˵�
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
                If .TextMatrix(.Row, .Cols - 2) = 3 Then
                    .ToolTipText = "�������ݵ�ԭ����"
                ElseIf .TextMatrix(.Row, .Cols - 2) = 2 Then
                    .ToolTipText = "��������"
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
                
                If mnuEditStrike.Visible = True Then
                    mnuEditStrike.Enabled = False
                    tlbTool.Buttons("Strike").Enabled = False
                End If
                 
                If mnuEditDisplay.Visible = True Then
                    mnuEditDisplay.Enabled = True
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
    
    If Format(strStart, "yyyy-mm-dd") = "1901-01-01" And Format(strVerifyStart, "yyyy-mm-dd") = "1901-01-01" Then
        strRange = "������� " & Format(strVerifyStart, "yyyy��MM��dd��") & "��" & Format(strVerifyEnd, "yyyy��MM��dd��")
    ElseIf Format(strVerifyStart, "yyyy-mm-dd") <> "1901-01-01" Then
        strRange = "�������� " & Format(strStart, "yyyy��MM��dd��") & "��" & Format(strEnd, "yyyy��MM��dd��") & "  ������� " & Format(strVerifyStart, "yyyy��MM��dd��") & "��" & Format(strVerifyEnd, "yyyy��MM��dd��")
    Else
        strRange = "�������� " & Format(strStart, "yyyy��MM��dd��") & "��" & Format(strEnd, "yyyy��MM��dd��")
    End If
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "����_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    
    objPrint.Title.Text = mstrTitle
        
    objRow.Add "ʱ�䣺" & strRange
    objRow.Add "���ţ�" & cboStock.Text
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
        
    objRow.Add "��ӡ��:" & gstrUserName
    objRow.Add "��ӡ����:" & Format(Sys.Currentdate, "yyyy��MM��dd��")
    objPrint.BelowAppRows.Add objRow
    
    Set objPrint.Body = vsfList
    
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

Private Sub tlbTool_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu mnuViewTool
    End If
End Sub

'Ѱ����ĳһ����ȵ���
Public Function FindRow(ByVal FlexTemp As MSHFlexGrid, ByVal intTemp As Variant, ByVal intCol As Integer) As Integer
    Dim i As Integer
    
    With FlexTemp
        For i = 1 To .rows - 1
            If IsDate(intTemp) Then
               If Format(.TextMatrix(i, intCol), "yyyy-mm-dd") = Format(intTemp, "yyyy-mm-dd") Then
                  FindRow = i
                  Exit Function
               End If
            Else
                If .TextMatrix(i, intCol) = intTemp Then
                  FindRow = i
                  Exit Function
                End If
            End If
        Next
    End With
    FindRow = 1
End Function

Private Sub subExcel(bytMode As Byte)
'����:������ݵ�EXCEL
'����:bytMode=3 �����EXCEL
'
    Dim objPrint As Object
    Dim objRow As zlTabAppRow
    Dim strRange As String
    
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "����_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    
    objPrint.Title.Text = mstrTitle
        
    Set objRow = New zlTabAppRow
    objRow.Add ""
    objRow.Add "NO." & Trim(vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "NO")))
    objPrint.UnderAppRows.Add objRow

    Set objRow = New zlTabAppRow
    objRow.Add "�ⷿ��" & Trim(cboStock.Text)
    objPrint.UnderAppRows.Add objRow
        
    Set objRow = New zlTabAppRow
    objRow.Add "ժҪ:" & vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "ժҪ"))
    objPrint.BelowAppRows.Add objRow
        
    Set objRow = New zlTabAppRow
    objRow.Add "������:" & vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "������")) & "  ��������:" & vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "��������"))
    
    objRow.Add "�����:" & vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "�����")) & "  �������:" & vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "�������"))
    objPrint.BelowAppRows.Add objRow
    
    Set objPrint.Body = vsfDetail
    zlPrintOrView1Grd objPrint, bytMode
End Sub




Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

