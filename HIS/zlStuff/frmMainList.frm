VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMainList 
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
   Begin VSFlex8Ctl.VSFlexGrid vsfCostlyInfo 
      Height          =   615
      Left            =   6360
      TabIndex        =   8
      Top             =   3480
      Width           =   2295
      _cx             =   4048
      _cy             =   1085
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
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
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
   Begin TabDlg.SSTab TabShow 
      Height          =   345
      Left            =   255
      TabIndex        =   6
      Top             =   1080
      Visible         =   0   'False
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   609
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "�Ƴ��ⷿ(&0)"
      TabPicture(0)   =   "frmMainList.frx":014A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "����ⷿ(&1)"
      TabPicture(1)   =   "frmMainList.frx":0166
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
   End
   Begin MSComctlLib.ImageList ilsHot 
      Left            =   7290
      Top             =   1830
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
            Picture         =   "frmMainList.frx":0182
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":03A2
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":05C2
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":07DE
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":09FE
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":0C1E
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":0E3A
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":1056
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":1270
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":13CA
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":15EA
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":180A
            Key             =   "Prepare"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":1F04
            Key             =   "Send"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":211E
            Key             =   "Back"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":2338
            Key             =   "Cancel"
            Object.Tag             =   "Cancel"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Cmd���� 
      Caption         =   "����(&V)"
      Height          =   350
      Left            =   5250
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2700
      Width           =   1100
   End
   Begin VB.PictureBox picSeparate_s 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   600
      MousePointer    =   7  'Size N S
      ScaleHeight     =   630
      ScaleWidth      =   4815
      TabIndex        =   4
      Top             =   2685
      Width           =   4815
   End
   Begin ComCtl3.CoolBar cbrTool 
      Height          =   1125
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   1984
      BandCount       =   2
      _CBWidth        =   9480
      _CBHeight       =   1125
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
         Left            =   585
         TabIndex        =   2
         Text            =   "cboStock"
         Top             =   780
         Width           =   8805
      End
      Begin MSComctlLib.Toolbar tlbTool 
         Height          =   720
         Left            =   165
         TabIndex        =   1
         Top             =   30
         Width           =   9225
         _ExtentX        =   16272
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
            NumButtons      =   21
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
               Object.Visible         =   0   'False
               Key             =   "PrepareSplit"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "�˲�"
               Key             =   "Check"
               Object.ToolTipText     =   "�˲�"
               Object.Tag             =   "�˲�"
               ImageKey        =   "Prepare"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "ȡ��"
               Key             =   "CancelCheck"
               Object.ToolTipText     =   "ȡ���˲�"
               Object.Tag             =   "ȡ��"
               ImageKey        =   "Cancel"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "����"
               Key             =   "Prepare"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageKey        =   "Prepare"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "����"
               Key             =   "Send"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageKey        =   "Send"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "����"
               Key             =   "Back"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageKey        =   "Back"
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "EditSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "���"
               Key             =   "Verify"
               Description     =   "���"
               Object.ToolTipText     =   "���"
               Object.Tag             =   "���"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Strike"
               Description     =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "VerifySeparate"
               Style           =   3
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Search"
               Description     =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ˢ��"
               Key             =   "Refresh"
               Description     =   "ˢ��"
               Object.ToolTipText     =   "ˢ��"
               Object.Tag             =   "ˢ��"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "FindSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Description     =   "��������"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Exit"
               Description     =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageIndex      =   11
            EndProperty
         EndProperty
         MouseIcon       =   "frmMainList.frx":2552
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
            Picture         =   "frmMainList.frx":286C
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
      Left            =   6675
      Top             =   1845
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
            Picture         =   "frmMainList.frx":3100
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":3320
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":3540
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":375C
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":397C
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":3B9C
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":3DB8
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":3FD4
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":41EE
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":4348
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":4564
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":4784
            Key             =   "Prepare"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":4E7E
            Key             =   "Send"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":5098
            Key             =   "Back"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":52B2
            Key             =   "Cancel"
            Object.Tag             =   "Cancel"
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshList 
      Height          =   1455
      Left            =   480
      TabIndex        =   9
      Top             =   1320
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   2566
      _Version        =   393216
      FixedCols       =   0
      BackColorBkg    =   -2147483628
      AllowBigSelection=   0   'False
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      RowSizingMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VSFlex8Ctl.VSFlexGrid mshDetail 
      Height          =   945
      Left            =   360
      TabIndex        =   10
      Top             =   3480
      Width           =   5985
      _cx             =   10557
      _cy             =   1667
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
      BackColor       =   -2147483628
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483628
      GridColor       =   12632256
      GridColorFixed  =   0
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483628
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
      FormatString    =   $"frmMainList.frx":54CC
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
      ExplorerBar     =   7
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
      Begin VB.Image imgLeft 
         Height          =   240
         Left            =   30
         Picture         =   "frmMainList.frx":55A1
         Top             =   45
         Width           =   240
      End
   End
   Begin VB.Label lblCostly 
      BackColor       =   &H8000000A&
      Caption         =   "��ֵ������Ϣ"
      Height          =   195
      Left            =   6360
      TabIndex        =   7
      Top             =   3300
      Width           =   1455
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
      Begin VB.Menu mnuPlugIn 
         Caption         =   "��չ(&E)"
         Visible         =   0   'False
         Begin VB.Menu mnuPlugItem 
            Caption         =   "����"
            Index           =   0
         End
      End
      Begin VB.Menu mnuEditLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCheckBatch 
         Caption         =   "�������������˲�(&H)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditCheck 
         Caption         =   "�˲�(&H)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditCancelCheck 
         Caption         =   "ȡ���˲�(&Q)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditCheckLine 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditVerifyBatch 
         Caption         =   "���������������(&T)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditVerify 
         Caption         =   "���(&C)"
      End
      Begin VB.Menu mnuEditStrike 
         Caption         =   "����(&K)"
      End
      Begin VB.Menu mnuEditLine0 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditRestore 
         Caption         =   "�����˻�(&R)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditBill 
         Caption         =   "�޸ķ�Ʊ��Ϣ(&B)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditReg 
         Caption         =   "�޸�ע��֤��(&G)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditAcc 
         Caption         =   "�������(&V)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditImport 
         Caption         =   "����ƻ���(&I)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditImportFile 
         Caption         =   "�����ⲿ�ļ�(&F)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditPrepare 
         Caption         =   "����(&P)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditSend 
         Caption         =   "����(&S)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditBack 
         Caption         =   "����(&B)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditPrePareSp 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditVerifySelect 
         Caption         =   "������˵���ѯ(&Y)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditDisplay 
         Caption         =   "�鿴����(&W)"
      End
      Begin VB.Menu mnuEditLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditTMPrint 
         Caption         =   "���������ӡ����"
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
      Begin VB.Menu mnuViewLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewColDefine 
         Caption         =   "��ѡ��(&C)"
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
Private mstrPrivs As String                     'Ȩ��
Private mblnFirst As Boolean
Private mblnPopupmenuCall As Boolean
Private mstrOrder As String             '��¼����ʽ
Private mStr�ⷿ As String              '��¼��ǰ����Ա���ܲ��������пⷿ
Private mbln����˲� As Boolean     '�����Ƿ���Ҫ�˲� true-��Ҫ false-����Ҫ
Private mintFindDay As Integer      '��ѯ������Χ

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
Private mFMT As g_FmtString

'----------------------------------------------------------------------------------------------------------
'��������
Private mdtStartDate As Date
Private mdtEndDate As Date
Private mdtVerifyStart As Date
Private mdtVerifyEnd As Date
Private mintOldY As Integer
Private mintUnit As Integer                 '0��ɢװ��λ��1����װ��λ
Private mstrPrintRange As String      '��ӡ��Χ�ı�
Private mstrMoneySum As String        '���ϼ�
Private mint�з�Ʊ As Integer
Private mint�޷�Ʊ As Integer
Private mstrOthers() As String ' 0-��¼״̬(�ƻ�����),1-��ʼ����,2-��������,3-����id,4-�Է�����id(��������id����Ʒ���(�ƻ���)),5-������,6-�����,7-��Ӧ��ID,8-������,9-��ʼ��������,10-������������,11-��ʼ��Ʊ��,12-������Ʊ��,13-������Ϣ
'------------------------------------------------------------------------------------------------------------------------
'--���˺�:20060803,����:8740
Private mbln���ϲ������� As Boolean       '������Ч
Private mblnֻ�߱���ͨ����       As Boolean       '��ǰ��Ա���ڿ���ֻ�߱���ͨ����,������Ч
'------------------------------------------------------------------------------------------------------------------------

Private mint�ƿ⴦������ As Integer                    '1-��Ҫ���ϡ����͡�������һ����  0-����Ҫ��һ����
Private mbln��Ҫ�˲�    As Boolean              'ֻ����⹺���
Private mint������˷�ʽ As Integer             '������ˣ�0����ͨ��ˣ�1����Ҫ�Ȳ������
Private mstr��ֵ�Ĳ� As String              '��¼�����������Ƿ�ѡ���˸�ֵ�Ĳ�
Private mint�������� As Integer             '0-����Ҫ������� 1-��Ҫ�������
Private mint������ʽ As Integer             '0������������ʽ��1�������������뵥�ݣ�2������Ѳ����ĳ������뵥��
Private mblnCostView As Boolean             '�鿴�ɱ��������Ϣ true-����鿴 false-������鿴
Private mbln�ƿ���ȷ���� As Boolean         '�Ƿ���ȷ���Σ������ƿⵥ��Ч
Public Sub SetMenu()
    '���ر��ϡ����͡���������
    If mlngMode <> 1716 Then Exit Sub
    
    mnuEditPrepare.Visible = False
    mnuEditSend.Visible = False
    mnuEditBack.Visible = False
    
    tlbTool.Buttons("Prepare").Visible = False
    tlbTool.Buttons("Send").Visible = False
    tlbTool.Buttons("Back").Visible = False
    
    mnuEditVerify.Visible = False
    mnuEditStrike.Visible = False
    tlbTool.Buttons("Verify").Visible = False
    tlbTool.Buttons("Strike").Visible = False
 
    mnuEditLine1.Visible = False
    mnuEditLine0.Visible = False
    mnuEditLine2.Visible = False
    mnuEditPrePareSp.Visible = False
    tlbTool.Buttons("EditSeparate").Visible = False
    tlbTool.Buttons("VerifySeparate").Visible = False
    
    '���ݵ�ǰҳ�濪��
    If TabShow.Tab = 0 Then
        If mlngMode = 1716 Then
            mint�ƿ⴦������ = IIf(Val(zlDatabase.GetPara("�ƿ�����", glngSys, mlngMode, "0", , , , cboStock.ItemData(cboStock.ListIndex))) = 1, 1, 0)
            
            If mint�ƿ⴦������ = 0 Then
                mnuEditPrepare.Visible = False
                mnuEditVerify.Visible = zlStr.IsHavePrivs(mstrPrivs, "���")
                mnuEditStrike.Visible = zlStr.IsHavePrivs(mstrPrivs, "����")
                mnuEditLine0.Visible = mnuEditVerify.Visible Or mnuEditAdd.Visible Or mnuEditModify.Visible Or mnuEditDel.Visible
                mnuEditLine1.Visible = mnuEditVerify.Visible And (mnuEditAdd.Visible Or mnuEditModify.Visible Or mnuEditDel.Visible)
                tlbTool.Buttons("Verify").Visible = mnuEditVerify.Visible
                tlbTool.Buttons("Strike").Visible = mnuEditStrike.Visible
                
                tlbTool.Buttons("VerifySeparate").Visible = mnuEditLine0.Visible
                 tlbTool.Buttons("PrintSeparate").Visible = mnuEditLine0.Visible
                mnuEditVerify.Caption = "���(&C)"
                tlbTool.Buttons("Verify").Caption = "���"
                tlbTool.Buttons("Verify").Tag = "���"
                tlbTool.Buttons("Verify").ToolTipText = "���"
            Else
                mnuEditVerify.Caption = "����(&C)"
                tlbTool.Buttons("Verify").Caption = "����"
                tlbTool.Buttons("Verify").Tag = "����"
                tlbTool.Buttons("Verify").ToolTipText = "����"
                mnuEditPrepare.Visible = zlStr.IsHavePrivs(mstrPrivs, "����")
                mnuEditLine1.Visible = mnuEditPrepare.Visible
                mnuEditPrePareSp.Visible = mnuEditPrepare.Visible
            End If
            
            mint�������� = IIf(Val(zlDatabase.GetPara("��������", glngSys, mlngMode, "0")) = 1, 1, 0)
            If mint�������� = 0 Then
                tlbTool.Buttons("Strike").Visible = False
                tlbTool.Buttons("VerifySeparate").Visible = False
                mnuEditStrike.Caption = "����(&K)"
            Else
                tlbTool.Buttons("Strike").ToolTipText = "��˳���"
                tlbTool.Buttons("Strike").Caption = "��˳���"
                tlbTool.Buttons("Strike").Tag = "��˳���"
                tlbTool.Buttons("Strike").Visible = zlStr.IsHavePrivs(mstrPrivs, "����")
                tlbTool.Buttons("VerifySeparate").Visible = tlbTool.Buttons("Strike").Visible
                mnuEditStrike.Caption = "��˳���(&K)"
            End If
            
            mnuEditSend.Visible = mnuEditPrepare.Visible
            mnuEditBack.Visible = mnuEditPrepare.Visible
            tlbTool.Buttons("Prepare").Visible = mnuEditPrepare.Visible
            tlbTool.Buttons("Send").Visible = mnuEditPrepare.Visible
            tlbTool.Buttons("Back").Visible = mnuEditPrepare.Visible
            tlbTool.Buttons("EditSeparate").Visible = mnuEditPrepare.Visible
        End If
    Else
        If mlngMode = 1716 Then '�ƿ�
            mint�������� = IIf(Val(zlDatabase.GetPara("��������", glngSys, mlngMode, "0")) = 1, 1, 0)
            If mint�������� = 0 Then
                tlbTool.Buttons("Strike").ToolTipText = "����"
                tlbTool.Buttons("Strike").Caption = "����"
                tlbTool.Buttons("Strike").Tag = "����"
                mnuEditStrike.Caption = "����(&K)"
            Else
                tlbTool.Buttons("Strike").ToolTipText = "�������"
                tlbTool.Buttons("Strike").Caption = "�������"
                tlbTool.Buttons("Strike").Tag = "�������"
                mnuEditStrike.Caption = "�������(&K)"
            End If
        End If
        mnuEditVerify.Visible = zlStr.IsHavePrivs(mstrPrivs, "���")
        mnuEditStrike.Visible = zlStr.IsHavePrivs(mstrPrivs, "����")
        
        mnuEditLine1.Visible = mnuEditAdd.Visible Or mnuEditDel.Visible Or mnuEditModify.Visible
        mnuEditPrePareSp.Visible = mnuEditVerify.Visible Or mnuEditStrike.Visible
        
        tlbTool.Buttons("Verify").Visible = mnuEditVerify.Visible
        tlbTool.Buttons("Strike").Visible = mnuEditStrike.Visible
        tlbTool.Buttons("VerifySeparate").Visible = True
    End If

End Sub


Private Sub cboStock_Click()
    On Error Resume Next
    Dim lng�ⷿID As Long
    Dim rsCheck As New ADODB.Recordset
    Dim str���� As String
    
    If mblnNoClick Then Exit Sub
    If cboStock.ListIndex >= 0 Then cboStock.Tag = cboStock.ItemData(cboStock.ListIndex)
    
    On Error GoTo ErrHandle
    '���ÿⷿ�Ƿ�Ϊ���Ŀ⣬ֻ�����Ŀ�������˻�
    lng�ⷿID = cboStock.ItemData(cboStock.ListIndex)
    If mlngMode = 1716 Or mlngMode = 1717 Or mlngMode = 1722 Then
        str���� = " in ('���Ŀ�')"
    Else
        str���� = " in ('���Ŀ�','����ⷿ')"
    End If
    
    gstrSQL = " SELECT DISTINCT 0 " & _
              " FROM ��������˵�� " & _
              " WHERE �������� " & str���� & _
              "         AND ����ID =[1]"
              
    Set rsCheck = zlDatabase.OpenSQLRecord(gstrSQL, "��鵱ǰ�ⷿ�Ƿ�Ϊ���Ŀ�", lng�ⷿID)
    
    mnuEditRestore.Enabled = (rsCheck.RecordCount <> 0)
'    mnuEditLine0.Enabled = (rsCheck.RecordCount <> 0)
    
    '�л��ⷿˢ�²˵���
    SetMenu
    
    If mblnBootUp Then mnuViewRefresh_Click
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
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

Public Sub ShowList(ByVal lngMode As Long, ByVal strTitle As String, ByVal frmMain As Variant)
    Dim strFind As String
    
    mblnBootUp = False
    mlngMode = lngMode
    mstrTitle = strTitle
    mstrPrivs = ";" & gstrPrivs & ";"
    
    
    mbln���ϲ������� = False
    If mlngMode = 1717 Then
        '���˺�:���ӿ������ϲ�������
        '����:8468
        mbln���ϲ������� = Val(zlDatabase.GetPara(132, glngSys, 0)) = 1
    End If
        
        
    If Not CheckDepend Then Exit Sub            '���������Բ���
    Me.Caption = strTitle
    Me.Tag = strTitle
                
    SetPopedom  '����Ȩ�����ò�ͬ����ʾ��Ŀ
    mintFindDay = Val(zlDatabase.GetPara("��ѯ����", glngSys, mlngMode, 1))
    mdtStartDate = Format(DateAdd("d", -mintFindDay, sys.Currentdate), "yyyy-MM-dd")
    mdtEndDate = Format(sys.Currentdate, "yyyy-MM-dd")
    mdtVerifyStart = "1901-01-01"
    mdtVerifyEnd = "1901-01-01"
    
    
    strFind = " AND A.��¼״̬ = 1 And A.������� is Null And A.�������� Between To_Date('" & Format(mdtStartDate, "yyyy-mm-dd") & " 00:00:00','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(mdtEndDate, "yyyy-mm-dd") & " 23:59:59','YYYY-MM-DD HH24:MI:SS')"
    
    mstrFind = strFind
    
   
    Call tabShow_Click(0)
    If mlngMode <> 1716 Then GetList (mstrFind) '�г�����ͷ
    
    
    
    RestoreWinState Me, App.ProductName, mstrTitle
    Call SetColCostPriceWidth
    mblnBootUp = True
    
    Call SetTlbAndMenuCaption
    Call SetMenu
    
    '2006-04-25:���˺�,ͳһ���ӱ�������ģ��Ĺ���
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, gstrPrivs)
     
    If IsObject(frmMain) Then
        Me.Show , frmMain
    Else
        OS.ShowChildWindow Me.hwnd, frmMain
    End If
    
    Me.ZOrder 0
End Sub

Private Sub SetColCostPriceWidth()
    Dim intCol As Integer
    Dim blnCol�ɱ� As Boolean
    
    With mshList
        For intCol = 1 To .Cols - 1
            If .TextMatrix(0, intCol) = "�����" Or .TextMatrix(0, intCol) = "��۽��" Or .TextMatrix(0, intCol) = "�ɱ����" Or .TextMatrix(0, intCol) = "������" Then
                .ColWidth(intCol) = IIf(mblnCostView = True, 1000, 0)
            End If
        Next
    End With
    With mshDetail
        Select Case mlngMode
            Case 1712 '�����⹺
                .ColWidth(.ColIndex("�����")) = IIf(mblnCostView = True, 1000, 0)
                .ColWidth(.ColIndex("������")) = IIf(mblnCostView = True, 1500, 0)
                .ColWidth(.ColIndex("���")) = IIf(mblnCostView = True, 1000, 0)
            Case 1713 '�������
                .ColWidth(.ColIndex("�ɹ���")) = IIf(mblnCostView = True, 1000, 0)
                .ColWidth(.ColIndex("�ɹ����")) = IIf(mblnCostView = True, 1500, 0)
                .ColWidth(.ColIndex("���")) = IIf(mblnCostView = True, 1000, 0)
            Case 1714 '�������
                .ColWidth(.ColIndex("�ɱ���")) = IIf(mblnCostView = True, 1000, 0)
                .ColWidth(.ColIndex("�ɱ����")) = IIf(mblnCostView = True, 1500, 0)
                .ColWidth(.ColIndex("���")) = IIf(mblnCostView = True, 1000, 0)
            Case 1716 '�����ƿ�
                .ColWidth(.ColIndex("�ɱ���")) = IIf(mblnCostView = True, 1000, 0)
                .ColWidth(.ColIndex("�ɱ����")) = IIf(mblnCostView = True, 1500, 0)
                .ColWidth(.ColIndex("���")) = IIf(mblnCostView = True, 1000, 0)
            Case 1717 '��������
                .ColWidth(.ColIndex("�ɱ���")) = IIf(mblnCostView = True, 1000, 0)
                .ColWidth(.ColIndex("�ɱ����")) = IIf(mblnCostView = True, 1500, 0)
                .ColWidth(.ColIndex("���")) = IIf(mblnCostView = True, 1000, 0)
            Case 1718 '��������
                .ColWidth(.ColIndex("�ɱ���")) = IIf(mblnCostView = True, 1000, 0)
                .ColWidth(.ColIndex("�ɱ����")) = IIf(mblnCostView = True, 1500, 0)
                .ColWidth(.ColIndex("���")) = IIf(mblnCostView = True, 1000, 0)
        End Select
    End With
End Sub
Private Sub SetTlbAndMenuCaption()
    '���ò˵��͹��������������
    
    If mlngMode = 1716 Then
        mnuEditVerify.Caption = "����(&C)"
        tlbTool.Buttons("Verify").Caption = "����"
        tlbTool.Buttons("Verify").Tag = "����"
        tlbTool.Buttons("Verify").ToolTipText = "����"
        TabShow.Visible = True
    Else
        TabShow.Visible = False
    End If
    

End Sub


'�������������
Private Function CheckDepend() As Boolean
    Dim rsTemp As New Recordset
    
    On Error GoTo ErrHandle
    CheckDepend = False
    mblnֻ�߱���ͨ���� = False
 
    '��ȡ�ɲ����Ŀⷿ
    Select Case mlngMode
        Case 1712                       '�����⹺������
            mstr�������� = "V,K,12"
        Case 1713                       '��������������
            mstr�������� = "V,K,12"
        Case 1714                       '��������������
            mstr�������� = "V,K,W,12"
        Case 1715                       '���Ŀ���۵���
            mstr�������� = "V,K,12"
        Case 1716                       '�����ƿ����
            '���ƿⵥ,���Ը����ϲ����ƿ�.ֻ���Ƶ����ϲ��ŵĲ��ϲ��ܱ�����.
            mstr�������� = "V,K,W"
        Case 1717                       '�������ù���
            '����:8468:20060803,��Ҫ���޸Ŀ������ϲ�������
            mstr�������� = "V,K" & IIf(mbln���ϲ������� = False, "", ",W")
            
            gstrSQL = "" & _
                "   SELECT /*+ Rule*/ DISTINCT a.id, a.���� " & _
                "   FROM ��������˵�� c, �������ʷ��� b, ���ű� a " & _
                "     , Table(Cast(f_Str2List([3]) as zlTools.t_StrList)) D " & _
                "   Where c.�������� = b.���� and (a.վ��=[2] or a.վ�� is null) and b.���� = D.Column_Value " & _
                "           AND a.id = c.����id " & _
                "           AND TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01'" & _
                "           And a.ID IN (Select ����ID From ������Ա Where ��ԱID=[1])"
            '"         AND instr(',V,K,W',b.����)>0"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ա�ⷿ����", UserInfo.Id, gstrNodeNo, mstr��������)
            If rsTemp.EOF Then
                mblnֻ�߱���ͨ���� = True
                '��������ͨ���ң���˾߱������пⷿ����,���ǲ��߱����;������Ȩ��
                If InStr(mstrPrivs, ";���пⷿ;") = 0 Then
                    mstrPrivs = mstrPrivs & ";���пⷿ;"
                End If
                mstrPrivs = Replace(mstrPrivs, ";���;", ";")
                mstrPrivs = Replace(mstrPrivs, ";����;", ";")
            Else
                mblnֻ�߱���ͨ���� = False
            End If
        Case 1718                       '���������������
            mstr�������� = "W,V,K,12"
        Case 1719                       '�����̵����
            mstr�������� = "V,K,12"
        Case Else
    End Select
    
    gstrSQL = "" & _
        "   SELECT /*+ Rule*/ DISTINCT a.id,a.����, a.����,a.����" & _
        "   FROM ��������˵�� c, �������ʷ��� b, ���ű� a " & _
        "     , Table(Cast(f_Str2List([3]) as zlTools.t_StrList)) D " & _
        "   Where c.�������� = b.���� and (a.վ��=[2] or a.վ�� is null) " & _
        "           AND b.���� = D.Column_Value " & _
        "           AND a.id = c.����id " & _
        "           AND TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01'" & _
        IIf(InStr(mstrPrivs, "���пⷿ") <> 0, "", " And a.ID IN (Select ����ID From ������Ա Where ��ԱID=[1])")
    
    mbln����Ա���� = Not zlStr.IsHavePrivs(mstrPrivs, "���пⷿ")
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ӧ�Ŀⷿ", UserInfo.Id, gstrNodeNo, mstr��������)
    
    If rsTemp.EOF Then
        If mlngMode = 1717 And mbln���ϲ������� = False Then
            ShowMsgBox "����Ӧ������һ���������Ŀ�����" & vbCrLf & "�����Ƽ������ʵĲ���,��鿴���Ź���"
        Else
            ShowMsgBox "����Ӧ������һ���������Ŀ����ʣ����ϲ�������" & vbCrLf & "�����Ƽ������ʵĲ���,��鿴���Ź���"
        End If
        rsTemp.Close
        Exit Function
    End If
    
    
    'װ��ⷿ����
    With cboStock
        .Clear
        Do While Not rsTemp.EOF
            .AddItem rsTemp!����
            .ItemData(.NewIndex) = rsTemp!Id
            mStr�ⷿ = mStr�ⷿ & rsTemp!Id & "," & rsTemp!���� & "|"
            If rsTemp!Id = UserInfo.����ID Then
                .ListIndex = .NewIndex
            End If
            rsTemp.MoveNext
        Loop
        rsTemp.Close
        If .ListIndex = -1 Then
            .ListIndex = 0
        End If
    End With
    mbln��Ҫ�˲� = False
    If mlngMode = 1712 Then
        '�⹺��⣬��Ҫȷ���Ƿ���Ҫ�˲鹦��
        mbln��Ҫ�˲� = Val(zlDatabase.GetPara("�����⹺��Ҫ�˲�", glngSys, "0")) = 1
    End If
    
    mint������˷�ʽ = 0
    If mlngMode = 1717 Then
        mint������˷�ʽ = Val(zlDatabase.GetPara("�������", glngSys, mlngMode, "0"))
    End If
    
    CheckDepend = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub GetList(ByVal strFind As String)
    Dim rsTemp As New Recordset
    Dim strUserPart As String
    Dim dbl1 As Double, dbl2 As Double, dbl3 As Double, dbl4 As Double
    Dim strTemp As String
    mlastRow = 0
    
    On Error GoTo ErrHandle
    Call FS.ShowFlash("���������������ϼ�¼,���Ժ� ...", Me)
    DoEvents
    Screen.MousePointer = vbHourglass
    strUserPart = " And A.�ⷿID+0=[1]"
    mshList.Redraw = False
    
    Select Case mlngMode
        Case 1712           '�����⹺������
            gstrSQL = "" & _
                "   SELECT  A.No, Decode(Nvl(A.��ҩ��ʽ, 0), 0, '��ⵥ', '�˿ⵥ') as ����˵��,C.���� AS ��Ӧ��,ltrim(to_char(SUM(A.�ɱ����)," & mOraFMT.FM_��� & ")) AS ������," & _
                "           ltrim(to_char((SUM(A.���۽��))," & mOraFMT.FM_��� & ")) AS �ۼ۽��," & _
                "           LTRIM(TO_CHAR((SUM(A.���۽�� - A.�ɱ����))," & mOraFMT.FM_��� & " )) AS ��۽��," & _
                "           Decode(Sign(Nvl(Max(a.����id), 0) - 1), 1, 0, Nvl(Max(a.����id), 0)) as �����־, A.������,TO_CHAR(min(A.��������), 'yyyy-mm-dd HH24:Mi:SS') AS ��������, " & _
                            IIf(mbln��Ҫ�˲� = False, "", "           A.��ҩ�� as �˲���,TO_CHAR(min(A.��ҩ����), 'yyyy-mm-dd HH24:Mi:SS') AS �˲�����, ") & _
                "           A.�����,TO_CHAR(min(A.�������), 'yyyy-mm-dd HH24:Mi:SS') AS �������," & _
                "           A.��¼״̬, " & _
                "           A.ժҪ, Max(����id) As ����id " & _
                "   FROM ҩƷ�շ���¼ A, ���ű� B, ��Ӧ�� C,�������� D, Ӧ����¼ E " & _
                "   Where A.�ⷿid = B.ID AND A.��ҩ��λid = C.Id and (c.վ��=[21] or c.վ�� is null) AND A.���� = 15 and a.ҩƷid=d.����id(+) and e.ϵͳ��ʶ(+) = 5 And e.��¼����(+) = 0 And a.Id = e.�շ�id(+) " & mstr��ֵ�Ĳ� & _
                        strUserPart & strFind & _
                "   GROUP BY A.No,C.����,A.������,A.�����,A.��ҩ�� ,A.��¼״̬,A.ժҪ,A.��ҩ��ʽ " & _
                "   ORDER BY   No Desc,�������� asc"
                
        Case 1713           '��������������
             
            gstrSQL = "" & _
                "   SELECT  a.no, c.���� AS �Ƽ���,ltrim(TO_CHAR (SUM (nvl(a.�ɱ����,0))," & mOraFMT.FM_��� & ")) AS �ɱ����," & _
                "           ltrim(TO_CHAR ( (SUM (nvl(a.���۽��,0))), " & mOraFMT.FM_��� & ")) AS �ۼ۽��," & _
                "           LTRIM(TO_CHAR((SUM(A.���۽�� - nvl(A.�ɱ����,0)))," & mOraFMT.FM_��� & " )) AS ��۽��," & _
                "           a.������, " & _
                "           TO_CHAR (min(a.��������), 'yyyy-mm-dd HH24:Mi:SS') AS ��������, a.�����, " & _
                "           TO_CHAR (min(a.�������), 'yyyy-mm-dd HH24:Mi:SS') AS �������, a.��¼״̬, a.ժҪ " & _
                "   FROM ҩƷ�շ���¼ a, ���ű� b ,���ű� c " & _
                "   Where   a.�ⷿid = b.ID AND a.�Է�����id=c.id AND a.���� = 16 and a.���ϵ��=1 " & _
                            strUserPart & strFind & _
                "   GROUP BY a.no,c.����,a.������,a.�����,a.��¼״̬,a.ժҪ " & _
                "   ORDER BY no DESC, �������� ASC "
    
        Case 1714           '��������������
            gstrSQL = "" & _
                "   SELECT  a.no, c.���� AS ������,ltrim(TO_CHAR (SUM (a.�ɱ����)," & mOraFMT.FM_��� & " )) AS �ɱ����," & _
                "           ltrim(TO_CHAR ((SUM (a.���۽��)), " & mOraFMT.FM_��� & " )) AS �ۼ۽��," & _
                "           LTRIM(TO_CHAR((SUM(A.���۽�� - A.�ɱ����))," & mOraFMT.FM_��� & " )) AS ��۽��," & _
                "           a.������,TO_CHAR (min(a.��������), 'yyyy-mm-dd HH24:Mi:SS') AS ��������, a.�����," & _
                "           TO_CHAR (min(a.�������), 'yyyy-mm-dd HH24:Mi:SS') AS �������, a.��¼״̬, a.ժҪ " & _
                "   FROM ҩƷ�շ���¼ a, ���ű� b,ҩƷ������ c " & _
                "   Where a.�ⷿid = b.ID AND a.������id = c.id AND a.���� = 17 " & _
                            strUserPart & strFind & _
                "   GROUP BY a.no,c.����,a.������,a.�����,a.��¼״̬,a.ժҪ " & _
                "   ORDER BY no DESC,�������� ASC "
                
        Case 1715           '����۵�������
            gstrSQL = "" & _
                "   SELECT  a.no, ltrim(TO_CHAR (SUM (a.���ۼ�), " & mOraFMT.FM_��� & " )) AS �����," & _
                "           ltrim(TO_CHAR (SUM (a.�ɱ���), " & mOraFMT.FM_��� & " )) AS �����," & _
                "           ltrim(TO_CHAR ( (SUM (a.���))," & mOraFMT.FM_��� & " )) AS ������, " & _
                "           a.������,TO_CHAR (min(a.��������), 'yyyy-mm-dd HH24:Mi:SS') AS ��������, a.�����," & _
                "           TO_CHAR(min(a.�������), 'yyyy-mm-dd HH24:Mi:SS') AS �������, a.��¼״̬, a.ժҪ,max(nvl(a.��ҩ��ʽ,0)) as ��ҩ��ʽ  " & _
                "   FROM ҩƷ�շ���¼ a, ���ű� b " & _
                "   Where   a.�ⷿid = b.ID  AND a.���� = 18 " & _
                            strUserPart & strFind & _
                "   GROUP BY a.no,a.������,a.�����,a.��¼״̬,a.ժҪ " & _
                "   ORDER BY no DESC,�������� ASC "
            
        Case 1716           '�����ƿ����
            If mbln����˲� = True Then
                strTemp = " and (Nvl(a.��ҩ��ʽ, 0) = 1 And a.�˲��� Is Not Null Or Nvl(a.��ҩ��ʽ, 0) = 0)"
            Else
                strTemp = ""
            End If
                If TabShow.Tab = 0 Then
                    strUserPart = " And A.�ⷿID+0=[1]"
                    
                    gstrSQL = "" & _
                        "   SELECT  a.no, c.���� AS ����ⷿ," & _
                        "           LTRIM(TO_CHAR (SUM (A.�ɱ����), " & mOraFMT.FM_��� & ")) AS �ɱ����, " & _
                        "           ltrim(TO_CHAR ((SUM (a.���۽��))," & mOraFMT.FM_��� & ")) AS �ۼ۽��," & _
                        "           LTRIM(TO_CHAR((SUM(A.���۽�� - A.�ɱ����))," & mOraFMT.FM_��� & " )) AS ��۽��," & _
                        "           a.������, " & _
                        "           TO_CHAR (min(a.��������), 'yyyy-mm-dd HH24:Mi:SS') AS ��������,a.����� as ������,  " & _
                        "           To_Char(Min(a.�������), 'yyyy-mm-dd HH24:Mi:SS') As ��������," & _
                        "           A.��ҩ�� AS ������,TO_CHAR(MIN(A.��ҩ����),'YYYY-MM-DD HH24:MI:SS') AS ��������," & _
                        "            a.��¼״̬, a.ժҪ " & _
                        "   FROM ҩƷ�շ���¼ a, ���ű� b ,���ű� c " & _
                        "   Where   a.�ⷿid = b.ID AND a.�Է�����id=c.id AND a.���� = 19 AND  a.���ϵ��=-1 " & _
                                    strUserPart & strFind & strTemp & _
                        "   GROUP BY a.no,c.����,a.������,a.�����,a.��ҩ��,a.��¼״̬,a.ժҪ " & _
                        "   ORDER BY a.no DESC,�������� ASC,a.��ҩ�� asc "
                Else
                    strUserPart = " And A.�Է�����ID+0=[1]"
                    gstrSQL = "" & _
                        "   SELECT  a.no, B.���� AS �Ƴ��ⷿ," & _
                        "           LTRIM(TO_CHAR (SUM (A.�ɱ����), " & mOraFMT.FM_��� & ")) AS �ɱ����, " & _
                        "           ltrim(TO_CHAR ((SUM (a.���۽��))," & mOraFMT.FM_��� & ")) AS �ۼ۽��," & _
                        "           LTRIM(TO_CHAR((SUM(A.���۽�� - A.�ɱ����))," & mOraFMT.FM_��� & " )) AS ��۽��," & _
                        "           a.������, " & _
                        "           TO_CHAR (min(a.��������), 'yyyy-mm-dd HH24:Mi:SS') AS ��������," & _
                        "           A.����� AS ������,TO_CHAR(MIN(A.�������),'YYYY-MM-DD HH24:MI:SS') AS �������," & _
                        "           a.��¼״̬, a.ժҪ " & _
                        "   FROM ҩƷ�շ���¼ a, ���ű� b ,���ű� c " & _
                        "   Where   a.�ⷿid = b.ID AND a.�Է�����id=c.id AND a.���� = 19 AND  a.���ϵ��=-1 " & _
                                    strUserPart & strFind & strTemp & _
                        "   GROUP BY a.no,b.����,a.������,a.�����,a.��¼״̬,a.ժҪ " & _
                        "   ORDER BY a.no DESC,�������� ASC,a.����� asc "
                End If
        Case 1717           '�������ù���
            gstrSQL = "" & _
                "   SELECT  a.no, c.���� AS ���ò���," & _
                "           LTRIM(TO_CHAR (SUM (A.�ɱ����), " & mOraFMT.FM_��� & ")) AS �ɱ����, " & _
                "           ltrim(TO_CHAR ((SUM (a.���۽��))," & mOraFMT.FM_��� & ")) AS �ۼ۽��," & _
                "           LTRIM(TO_CHAR((SUM(A.���۽�� - A.�ɱ����))," & mOraFMT.FM_��� & " )) AS ��۽��," & _
                "           a.������,a.������, " & _
                "           TO_CHAR (min(a.��������), 'yyyy-mm-dd HH24:Mi:SS') AS �������� "
                
                If mint������˷�ʽ = 1 Then
                    gstrSQL = gstrSQL & ", a.��ҩ�� As �˲���, TO_CHAR (min(a.��ҩ����), 'yyyy-mm-dd HH24:Mi:SS') AS �˲����� "
                End If
                
            gstrSQL = gstrSQL & ", a.�����, TO_CHAR (min(a.�������), 'yyyy-mm-dd HH24:Mi:SS') AS �������, a.��¼״̬, a.ժҪ " & _
                "   FROM ҩƷ�շ���¼ a, ���ű� b ,���ű� c " & _
                "   Where   a.�ⷿid = b.ID AND a.�Է�����id=c.id AND a.���� = 20 " & IIf(mblnֻ�߱���ͨ����, " and a.�Է�����id in (Select ����ID From ������Ա Where ��ԱID=[20])", "") & _
                            strUserPart & strFind & _
                "   GROUP BY a.no,c.����,a.������,a.������,a.��ҩ��,a.�����,a.��¼״̬,a.ժҪ " & _
                "   ORDER BY no DESC, �������� ASC "
                
        Case 1718          '���������������
            gstrSQL = "" & _
                "   SELECT /*+rule*/ a.no, c.���� AS ������,d.���� AS �Է���λ," & _
                "           LTRIM(TO_CHAR (SUM (A.�ɱ����), " & mOraFMT.FM_��� & ")) AS �ɱ����, " & _
                "           ltrim(TO_CHAR ((SUM (a.���۽��))," & mOraFMT.FM_��� & ")) AS �ۼ۽��," & _
                "           LTRIM(TO_CHAR((SUM(A.���۽�� - A.�ɱ����))," & mOraFMT.FM_��� & " )) AS ��۽��," & _
                "           LTrim(To_Char((Sum(A.���� * A.ʵ������)), '9999999999990.99')) As �������," & _
                "           a.������,TO_CHAR (min(a.��������), 'yyyy-mm-dd HH24:Mi:SS') AS ��������, a.�����," & _
                "           TO_CHAR (min(a.�������), 'yyyy-mm-dd HH24:Mi:SS') AS �������, a.��¼״̬, a.ժҪ, Max(����id) As ����id " & _
                "   FROM ҩƷ�շ���¼ a, ���ű� b,ҩƷ������ c,����������λ d " & _
                "   Where a.�ⷿid = b.ID AND a.������id = c.id AND A.��ҩ����=D.���� And a.���� = 21 " & strUserPart & strFind & _
                "   GROUP BY a.no,c.����,d.����,a.������,a.�����,a.��¼״̬,a.ժҪ "

            gstrSQL = gstrSQL & _
                " Union All " & _
                "   SELECT  a.no, c.���� AS ������,'' AS �Է���λ," & _
                "           LTRIM(TO_CHAR (SUM (A.�ɱ����), " & mOraFMT.FM_��� & ")) AS �ɱ����, " & _
                "           ltrim(TO_CHAR ((SUM (a.���۽��))," & mOraFMT.FM_��� & ")) AS �ۼ۽��," & _
                "           LTRIM(TO_CHAR((SUM(A.���۽�� - A.�ɱ����))," & mOraFMT.FM_��� & " )) AS ��۽��," & _
                "           LTrim(To_Char((Sum(A.���� * A.ʵ������)), '9999999999990.99')) As �������," & _
                "           a.������,TO_CHAR (min(a.��������), 'yyyy-mm-dd HH24:Mi:SS') AS ��������, a.�����," & _
                "           TO_CHAR (min(a.�������), 'yyyy-mm-dd HH24:Mi:SS') AS �������, a.��¼״̬, a.ժҪ, Max(����id) As ����id " & _
                "   FROM ҩƷ�շ���¼ a, ���ű� b,ҩƷ������ c " & _
                "   Where a.�ⷿid = b.ID AND a.������id = c.id AND A.��ҩ���� Is Not Null And A.��ҩ���� Not In (Select ���� From ����������λ) And a.���� = 21 " & strUserPart & strFind & _
                "   GROUP BY a.no,c.����,a.������,a.�����,a.��¼״̬,a.ժҪ "
            
            gstrSQL = gstrSQL & _
                " Union All " & _
                "   SELECT  a.no, c.���� AS ������,'' AS �Է���λ," & _
                "           LTRIM(TO_CHAR (SUM (A.�ɱ����), " & mOraFMT.FM_��� & ")) AS �ɱ����, " & _
                "           ltrim(TO_CHAR ((SUM (a.���۽��))," & mOraFMT.FM_��� & ")) AS �ۼ۽��," & _
                "           LTRIM(TO_CHAR((SUM(A.���۽�� - A.�ɱ����))," & mOraFMT.FM_��� & " )) AS ��۽��," & _
                "           LTrim(To_Char((Sum(A.���� * A.ʵ������)), '9999999999990.99')) As �������," & _
                "           a.������,TO_CHAR (min(a.��������), 'yyyy-mm-dd HH24:Mi:SS') AS ��������, a.�����," & _
                "           TO_CHAR (min(a.�������), 'yyyy-mm-dd HH24:Mi:SS') AS �������, a.��¼״̬, a.ժҪ, Max(����id) As ����id " & _
                "   FROM ҩƷ�շ���¼ a, ���ű� b,ҩƷ������ c " & _
                "   Where a.�ⷿid = b.ID AND a.������id = c.id And A.��ҩ���� Is Null And a.���� = 21 " & strUserPart & strFind & _
                "   GROUP BY a.no,c.����,a.������,a.�����,a.��¼״̬,a.ժҪ " & _
                "   ORDER BY no DESC,�������� ASC "
        Case 1719         '�����̵�
            'Ƶ���ֶα���� �̵�ʱ��
            gstrSQL = "" & _
                "   SELECT distinct a.no, Ƶ�� AS �̵�ʱ��," & _
                "           a.������,TO_CHAR (min(a.��������), 'yyyy-mm-dd HH24:Mi:SS') AS ��������, a.�����," & _
                "           TO_CHAR (min(a.�������), 'yyyy-mm-dd HH24:Mi:SS') AS �������, " & _
                "           ltrim(to_char((Sum(Nvl(����,0)*���ۼ�))," & mOraFMT.FM_��� & ")) �̵���," & _
                "           ltrim(to_char((Sum(���۽��))," & mOraFMT.FM_��� & ")) ����,a.��¼״̬, a.ժҪ " & _
                "   FROM ҩƷ�շ���¼ a, ���ű� b " & _
                "   Where a.�ⷿid = b.ID AND a.���� =22  " & strUserPart & strFind & _
                "   Group by a.no,Ƶ��,a.������,a.�����,a.��¼״̬, a.ժҪ " & _
                "   ORDER BY no DESC,�������� ASC "
    End Select
    
     'mstrOthers(0 To 13) As String ' 0-��¼״̬(�ƻ�����),1-��ʼ����,2-��������,3-����id,4-�Է�����id(��������id����Ʒ���(�ƻ���)),5-������,6-�����,7-��Ӧ��ID,8-������,9-��ʼ��������,10-������������,11-��ʼ��Ʊ��,12-������Ʊ��,13-������Ϣ
    '������Χ:[1]-�ⷿid,[2]:��ʼ��������,[3]������������,[4]��ʼ�������,[5] �����������,[6]-��¼״̬,[7]��ʼ���ݺ�,[8]�������ݺ�,[9]����id,[10]�Է�����id,[11]������,[12]�����[13]-��Ӧ��ID,[14]-������,[15]-��ʼ��������,[16]-������������,[17]-��ʼ��Ʊ��,[18]-������Ʊ��,[19]-������Ϣ
    
    '��ʼ��������
    mstrOthers(9) = IIf(Trim(mstrOthers(9)) = "", "1901-01-01", mstrOthers(9))
    mstrOthers(10) = IIf(Trim(mstrOthers(10)) = "", "1901-01-01", mstrOthers(10))
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
        cboStock.ItemData(cboStock.ListIndex), _
        CDate(Format(mdtStartDate, "yyyy-mm-dd") & " 00:00:00"), _
        CDate(Format(mdtEndDate, "yyyy-mm-dd") & " 23:59:59"), _
        CDate(Format(mdtVerifyStart, "yyyy-mm-dd") & " 00:00:00"), _
        CDate(Format(mdtVerifyEnd, "yyyy-mm-dd") & " 23:59:59"), _
        Val(mstrOthers(0)), _
        mstrOthers(1), _
        mstrOthers(2), _
        Val(mstrOthers(3)), _
        Val(mstrOthers(4)), _
        mstrOthers(5), _
        mstrOthers(6), _
        Val(mstrOthers(7)), _
        mstrOthers(8), _
        CDate(mstrOthers(9) & " 00:00:00"), _
        CDate(mstrOthers(10) & " 23:59:59"), _
        mstrOthers(11), _
        mstrOthers(12), _
        mstrOthers(13) & "%", _
        UserInfo.Id, _
        gstrNodeNo)
        
    Set mshList.DataSource = rsTemp
    
    With mshList
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
    
    Call SetListColWidth
    
    'ͳ�ƺϼƽ��
    If (Not rsTemp.EOF) And (Not rsTemp.BOF) Then
        rsTemp.MoveFirst
        Do While Not rsTemp.EOF
            Select Case mlngMode
                Case 1712
                    dbl1 = dbl1 + IIf(IsNull(rsTemp!������), 0, rsTemp!������)
                    dbl2 = dbl2 + IIf(IsNull(rsTemp!�ۼ۽��), 0, rsTemp!�ۼ۽��)
                    dbl3 = dbl3 + IIf(IsNull(rsTemp!��۽��), 0, rsTemp!��۽��)
                Case 1713, 1714, 1716, 1717
                    dbl1 = dbl1 + IIf(IsNull(rsTemp!�ɱ����), 0, rsTemp!�ɱ����)
                    dbl2 = dbl2 + IIf(IsNull(rsTemp!�ۼ۽��), 0, rsTemp!�ۼ۽��)
                    dbl3 = dbl3 + IIf(IsNull(rsTemp!��۽��), 0, rsTemp!��۽��)
                Case 1715
                    dbl1 = dbl1 + IIf(IsNull(rsTemp!�����), 0, rsTemp!�����)
                    dbl2 = dbl2 + IIf(IsNull(rsTemp!�����), 0, rsTemp!�����)
                    dbl3 = dbl3 + IIf(IsNull(rsTemp!������), 0, rsTemp!������)
                Case 1718
                    dbl1 = dbl1 + IIf(IsNull(rsTemp!�ɱ����), 0, rsTemp!�ɱ����)
                    dbl2 = dbl2 + IIf(IsNull(rsTemp!�ۼ۽��), 0, rsTemp!�ۼ۽��)
                    dbl3 = dbl3 + IIf(IsNull(rsTemp!��۽��), 0, rsTemp!��۽��)
                    dbl4 = dbl4 + IIf(IsNull(rsTemp!�������), 0, rsTemp!�������)
                Case 1719
                    dbl2 = dbl2 + IIf(IsNull(rsTemp!�̵���), 0, rsTemp!�̵���)
                    dbl3 = dbl3 + IIf(IsNull(rsTemp!����), 0, rsTemp!����)
            End Select
            rsTemp.MoveNext
        Loop
        rsTemp.MoveFirst
    End If
    Dim strText As String
    Select Case mlngMode
        Case 1712
            If mblnCostView = False Then
                strText = "�ۼ۽��ϼƣ�" & Format(dbl2, mFMT.FM_���)
            Else
                strText = "������ϼƣ�" & Format(dbl1, mFMT.FM_���)
                strText = strText & Space(10) & " �ۼ۽��ϼƣ�" & Format(dbl2, mFMT.FM_���)
                strText = strText & Space(10) & "��۽��ϼƣ�" & Format(dbl3, mFMT.FM_���)
            End If
        Case 1713, 1714, 1716, 1717
            If mblnCostView = False Then
                strText = "�ۼ۽��ϼƣ�" & Format(dbl2, mFMT.FM_���)
            Else
                strText = "�ɱ����ϼƣ�" & Format(dbl1, mFMT.FM_���)
                strText = strText & Space(10) & "�ۼ۽��ϼƣ�" & Format(dbl2, mFMT.FM_���)
                strText = strText & Space(10) & "��۽��ϼƣ�" & Format(dbl3, mFMT.FM_���)
            End If
        Case 1715
            strText = "�����ϼƣ�" & Format(dbl1, mFMT.FM_���)
            strText = strText & Space(10) & "����ۺϼƣ�" & Format(dbl2, mFMT.FM_���)
            strText = strText & Space(10) & "������ϼƣ�" & Format(dbl3, mFMT.FM_���)
        Case 1718
            If mblnCostView = False Then
                strText = "�ۼ۽��ϼƣ�" & Format(dbl2, mFMT.FM_���)
                strText = strText & Space(10) & "�������ϼƣ�" & Format(dbl4, mFMT.FM_���)
            Else
                strText = "�ɱ����ϼƣ�" & Format(dbl1, mFMT.FM_���)
                strText = strText & Space(10) & "�ۼ۽��ϼƣ�" & Format(dbl2, mFMT.FM_���)
                strText = strText & Space(10) & "��۽��ϼƣ�" & Format(dbl3, mFMT.FM_���)
                strText = strText & Space(10) & "�������ϼƣ�" & Format(dbl4, mFMT.FM_���)
            End If
        Case 1719
            strText = "�̵���ϼƣ�" & Format(dbl2, mFMT.FM_���)
            strText = strText & Space(10) & "����ϼƣ�" & Format(dbl3, mFMT.FM_���)
    End Select
    mstrMoneySum = strText
    PrintRange strText & Space(10) & vbCrLf & mstrPrintRange
    
    
    Call mshlist_EnterCell    '�г�������
    
    Call SetStrikeColor
    
    With mshList
        .Row = 1
        .Col = 0
        .ColSel = .Cols - 1
    End With
    mshList.Redraw = True
    
    Call FS.StopFlash
    
    Screen.MousePointer = vbDefault
    stbThis.Panels(2).Text = "��ǰ����" & rsTemp.RecordCount & "�ŵ���"
    
    rsTemp.Close
    If mshList.Visible = True Then
        mshList.SetFocus
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetStrikeColor()
    Dim int��¼״̬ As Integer, int�����־ As Integer
    Dim intRow As Integer, intCol As Integer
    Dim intCol��¼״̬ As Integer, intCol�����־ As Integer
    Dim int�Զ���� As Integer
    Dim intCol����� As Integer
        
    With mshList
        If .Rows <= 2 Then Exit Sub
        intCol��¼״̬ = GetCol(mshList, "��¼״̬")
        If intCol��¼״̬ < 0 Then Exit Sub
        intCol�����־ = GetCol(mshList, "�����־")
        int�Զ���� = GetCol(mshList, "����ID")
        If mlngMode = 1716 Then '�����ƿ�
            intCol����� = GetCol(mshList, "������")
        Else
            intCol����� = GetCol(mshList, "�����")
        End If
        
        For intRow = 1 To .Rows - 1
            int��¼״̬ = Val(.TextMatrix(intRow, intCol��¼״̬))
            If intCol�����־ >= 0 Then int�����־ = Val(.TextMatrix(intRow, intCol�����־))
            
            If int��¼״̬ Mod 3 = 0 Then
                .Row = intRow
                For intCol = 0 To .Cols - 1
                    .Col = intCol
                    .CellForeColor = &H80000001
                Next
            ElseIf int��¼״̬ Mod 3 = 2 Then
                .Row = intRow
                For intCol = 0 To .Cols - 1
                    .Col = intCol
                    If .TextMatrix(intRow, intCol�����) = "" Then
                        .CellForeColor = &HC0C0FF
                    Else
                        .CellForeColor = IIf(int�����־ = 1, &HC0C0FF, &HFF)   '�����־��ʾ��һ��
                    End If
                Next
            End If
            
            If int�Զ���� > 1 Then
                If Val(.TextMatrix(intRow, int�Զ����)) > 1 Then
                    .Row = intRow
                    For intCol = 0 To .Cols - 1
                        .Col = intCol
                        .CellForeColor = IIf(Val(.TextMatrix(intRow, int�Զ����)) > 1, &H808080, &H80000008)
                    Next
                End If
            End If
        Next
    End With
End Sub

'��ͷ�п��ʼ
Private Sub SetListColWidth()
    Dim intCol As Integer
    
    With mshList
        Select Case mlngMode
            Case 1712
                .ColAlignment(3) = flexAlignRightCenter
                .ColAlignment(4) = flexAlignRightCenter
                .ColAlignment(5) = flexAlignRightCenter
            Case 1713
                .ColAlignment(2) = flexAlignRightCenter
                .ColAlignment(3) = flexAlignRightCenter
                    
            Case 1714
                .ColAlignment(2) = flexAlignRightCenter
                .ColAlignment(3) = flexAlignRightCenter
            Case 1715
                .ColAlignment(1) = flexAlignRightCenter
                .ColAlignment(2) = flexAlignRightCenter
                .ColAlignment(3) = flexAlignRightCenter
                
            Case 1716
                .ColAlignment(2) = flexAlignRightCenter
            Case 1718
                .ColAlignment(2) = flexAlignLeftCenter
                .ColAlignment(3) = flexAlignRightCenter
                .ColAlignment(4) = flexAlignRightCenter
                .ColAlignment(5) = flexAlignRightCenter
                .ColAlignment(6) = flexAlignRightCenter
            Case 1719
                .ColAlignment(2) = flexAlignRightCenter         '�ۼ۽��
            Case 1720
                
            Case Else
            
        End Select
        intCol = GetCol(mshList, "��¼״̬")
        If intCol >= 0 Then mshList.ColWidth(intCol) = 0
        intCol = GetCol(mshList, "�����־")
        If intCol >= 0 Then .ColWidth(intCol) = 0
        intCol = GetCol(mshList, "����ID")
        If intCol >= 0 Then .ColWidth(intCol) = 0
        
        If mblnBootUp = False Then
            For intCol = 1 To .Cols - 1
                If intCol = 1 Then
                    If mlngMode = 1715 Then
                        .ColWidth(intCol) = 1000
                    ElseIf intCol = GetCol(mshList, "��������") Then
                        .ColWidth(intCol) = 900
                    Else
                        .ColWidth(intCol) = 2000
                    End If
                    
                ElseIf intCol = GetCol(mshList, "��¼״̬") Then
                    .ColWidth(intCol) = 0
                ElseIf intCol = GetCol(mshList, "�����־") Then
                    .ColWidth(intCol) = 0
                ElseIf intCol = GetCol(mshList, "��Ӧ��") Then
                     .ColWidth(intCol) = 2000
                Else
                    .ColWidth(intCol) = 1000
                End If
                If mlngMode = 1715 Then
                    If intCol = GetCol(mshList, "��ҩ��ʽ") Then
                        .ColWidth(intCol) = 0
                    End If
                End If
                
                If .TextMatrix(0, intCol) = "�����" Or .TextMatrix(0, intCol) = "��۽��" Or .TextMatrix(0, intCol) = "������" Or .TextMatrix(0, intCol) = "�ɱ����" Then
                    .ColWidth(intCol) = IIf(mblnCostView = True, 1000, 0)
                End If
            Next
        End If
    End With
End Sub


Private Sub SetDetailColWidth()
    Dim intCol As Integer
    Dim blnCol�ɱ� As Boolean
    
    With mshDetail
        For intCol = 0 To .Cols - 1
            .FixedAlignment(intCol) = flexAlignCenterCenter
            '��������
            .ColKey(intCol) = .TextMatrix(0, intCol)
        Next
        
        zl_vsGrid_Para_Restore mlngMode, mshDetail, mstrTitle, "mshDetail", False, True
        
        For intCol = 0 To .Cols - 1
'            .FixedAlignment(intCol) = flexAlignCenterCenter
'
'            '��������
'            .ColKey(intCol) = .TextMatrix(0, intCol)
            If mblnFirst Then
                If .ColWidth(intCol) = 0 Then .ColWidth(intCol) = 1000
            End If
            Select Case .ColKey(intCol)
            Case "��λ", "��Ʊ��", "�������", "��������", "���", "���۵�λ", "��Ʊ����"
                .ColAlignment(intCol) = flexAlignCenterCenter
                If .ColKey(intCol) = "���" Then
                    .ColWidth(intCol) = 0: .ColHidden(intCol) = True
                End If
            Case Else
                If .ColKey(intCol) = "������Ϣ" And mblnFirst = True Then
                    If .ColWidth(intCol) = 0 Then .ColWidth(intCol) = 2500
                End If
                .ColAlignment(intCol) = flexAlignLeftCenter
            End Select
            '.coldata(i):1-�̶�,-1-����ѡ,0-��ѡ
            If .ColKey(intCol) = "������Ϣ" Then .ColData(intCol) = 1
            If .ColKey(intCol) = "���" Then .ColData(intCol) = -1
            If .ColKey(intCol) Like "*����*" Or _
                .ColKey(intCol) Like "*��*" Or _
                .ColKey(intCol) Like "*��*" Or _
                .ColKey(intCol) Like "*��*" Or _
                .ColKey(intCol) Like "*��*" Then
                .ColAlignment(intCol) = flexAlignRightCenter
            End If
            '�������ۼ�\���۽��\���۲����ҪĬ��Ϊ��
            Select Case .ColKey(intCol)
            Case "���ۼ�", "���۵�λ", "���۽��", "���۲��"
                .ColHidden(intCol) = True
            Case Else
            End Select
            
            '����Ҫ����Ȩ�����ж��Ƿ���ʾ
            Select Case .ColKey(intCol)
            Case "��Ʒ����", "�ڲ�����"
                If gblnCode = False Then
                    .ColWidth(intCol) = 0
                    .ColHidden(intCol) = True
                End If
            End Select
        Next
        
        For intCol = 1 To mshList.Cols - 1
            If mshList.ColHeaderCaption(0, intCol) = "�ɱ����" Then
                blnCol�ɱ� = True
                Exit For
            End If
        Next
        
        Select Case mlngMode
            Case 1712 '�����⹺
                .ColWidth(.ColIndex("�����")) = IIf(mblnCostView = False, 0, 1000)
                If mblnCostView = False Then .ColData(.ColIndex("�����")) = -1
                .ColWidth(.ColIndex("������")) = IIf(mblnCostView = False, 0, 1500)
                If mblnCostView = False Then .ColData(.ColIndex("������")) = -1
                .ColWidth(.ColIndex("���")) = IIf(mblnCostView = False, 0, 1500)
                If mblnCostView = False Then .ColData(.ColIndex("���")) = -1
            Case 1713 '�������
                .ColWidth(.ColIndex("�ɹ���")) = IIf(mblnCostView = False, 0, 1000)
                If mblnCostView = False Then .ColData(.ColIndex("�ɹ���")) = -1
                .ColWidth(.ColIndex("�ɹ����")) = IIf(mblnCostView = False, 0, 1500)
                If mblnCostView = False Then .ColData(.ColIndex("�ɹ����")) = -1
                .ColWidth(.ColIndex("���")) = IIf(mblnCostView = False, 0, 1500)
                If mblnCostView = False Then .ColData(.ColIndex("���")) = -1
            Case 1714 '�������
                .ColWidth(.ColIndex("�ɱ���")) = IIf(mblnCostView = False, 0, 1000)
                If mblnCostView = False Then .ColData(.ColIndex("�ɱ���")) = -1
                .ColWidth(.ColIndex("�ɱ����")) = IIf(mblnCostView = False, 0, 1500)
                If mblnCostView = False Then .ColData(.ColIndex("�ɱ����")) = -1
                .ColWidth(.ColIndex("���")) = IIf(mblnCostView = False, 0, 1500)
                If mblnCostView = False Then .ColData(.ColIndex("���")) = -1
            Case 1716 '�����ƿ�
                .ColWidth(.ColIndex("�ɱ���")) = IIf(mblnCostView = False, 0, 1000)
                If mblnCostView = False Then .ColData(.ColIndex("�ɱ���")) = -1
                .ColWidth(.ColIndex("�ɱ����")) = IIf(mblnCostView = False, 0, 1500)
                If mblnCostView = False Then .ColData(.ColIndex("�ɱ����")) = -1
                .ColWidth(.ColIndex("���")) = IIf(mblnCostView = False, 0, 1500)
                If mblnCostView = False Then .ColData(.ColIndex("���")) = -1
                If blnCol�ɱ� = True Then
                    mshList.ColWidth(intCol) = IIf(mblnCostView = False, 0, 1500)
                End If
            Case 1717 '��������
                .ColWidth(.ColIndex("�ɱ���")) = IIf(mblnCostView = False, 0, 1000)
                If mblnCostView = False Then .ColData(.ColIndex("�ɱ���")) = -1
                .ColWidth(.ColIndex("�ɱ����")) = IIf(mblnCostView = False, 0, 1500)
                If mblnCostView = False Then .ColData(.ColIndex("�ɱ����")) = -1
                .ColWidth(.ColIndex("���")) = IIf(mblnCostView = False, 0, 1500)
                If mblnCostView = False Then .ColData(.ColIndex("���")) = -1
                If blnCol�ɱ� = True Then
                    mshList.ColWidth(intCol) = IIf(mblnCostView = False, 0, 1500)
                End If
            Case 1718 '��������
                .ColWidth(.ColIndex("�ɱ���")) = IIf(mblnCostView = False, 0, 1000)
                If mblnCostView = False Then .ColData(.ColIndex("�ɱ���")) = -1
                .ColWidth(.ColIndex("�ɱ����")) = IIf(mblnCostView = False, 0, 1500)
                If mblnCostView = False Then .ColData(.ColIndex("�ɱ����")) = -1
                .ColWidth(.ColIndex("���")) = IIf(mblnCostView = False, 0, 1500)
                If mblnCostView = False Then .ColData(.ColIndex("���")) = -1
                If blnCol�ɱ� = True Then
                    mshList.ColWidth(intCol) = IIf(mblnCostView = False, 0, 1500)
                End If
        End Select
    End With
End Sub


'����Ȩ�����ò�ͬ����ʾ��Ŀ
Private Sub SetPopedom()

    '�⹺�������Ȩ�ޣ��������á����������пⷿ���Ǽǡ��޸ġ�ɾ�������ա����������ݴ�ӡ
    Select Case mlngMode
        Case 1712, 1713, 1714, 1715, 1716, 1717, 1718, 1719
            If mlngMode = 1712 Then
                '���˺�:���Ӻ˲鹦��,2007/05/14
                mnuEditCheck.Visible = mbln��Ҫ�˲� And InStr(1, mstrPrivs, ";�˲�;") <> 0
                mnuEditCancelCheck.Visible = mnuEditCheck.Visible
                tlbTool.Buttons("Check").Visible = mnuEditCheck.Visible
                tlbTool.Buttons("CancelCheck").Visible = mnuEditCheck.Visible
                tlbTool.Buttons("PrepareSplit").Visible = mnuEditCheck.Visible
            End If
            
            If mlngMode = 1717 Then
                mnuEditCheck.Visible = (mint������˷�ʽ = 1) And InStr(1, mstrPrivs, ";�������;") <> 0
                mnuEditCancelCheck.Visible = mnuEditCheck.Visible
                mnuEditCheckLine.Visible = mnuEditCheck.Visible
                tlbTool.Buttons("Check").Visible = mnuEditCheck.Visible
                tlbTool.Buttons("CancelCheck").Visible = mnuEditCheck.Visible
                tlbTool.Buttons("PrepareSplit").Visible = mnuEditCheck.Visible
            End If
             
            If InStr(1, mstrPrivs, ";�Ǽ�;") = 0 Then
                mnuEditAdd.Visible = False
                mnuEditRestore.Visible = False
                tlbTool.Buttons("Add").Visible = False
            Else
                mnuEditRestore.Visible = True
            End If
            
            If InStr(1, mstrPrivs, ";�޸�;") = 0 Then
                mnuEditModify.Visible = False
                tlbTool.Buttons("Modify").Visible = False
            End If
            
            If InStr(1, mstrPrivs, ";ɾ��;") = 0 Then
                mnuEditDel.Visible = False
                tlbTool.Buttons("Delete").Visible = False
                 '��û�����б༭Ȩ��ʱ���Ѳ˵��͹������ϵ���Ӧ�ķָ������Ρ�
                If mnuEditAdd.Visible = False And mnuEditModify.Visible = False Then
                    mnuEditLine1.Visible = False
                    tlbTool.Buttons("EditSeparate").Visible = False
                End If
            End If
            
            If InStr(1, ";" & mstrPrivs & ";", ";���;") = 0 Then
                mnuEditVerify.Visible = False
                mnuEditBill.Visible = False
                mnuEditReg.Visible = False
                tlbTool.Buttons("Verify").Visible = False
            End If
            
            If InStr(1, mstrPrivs, ";����;") = 0 Then
                mnuEditStrike.Visible = False
                tlbTool.Buttons("Strike").Visible = False
                
                If mnuEditVerify.Visible = False Then
                    mnuEditLine2.Visible = False
                    tlbTool.Buttons("VerifySeparate").Visible = False
                End If
            End If
            If InStr(1, mstrPrivs, ";���ݴ�ӡ;") = 0 Then
                mnuFileBillPrint.Visible = False
                mnuFileBillPreview.Visible = False
            End If
        Case Else
        
    End Select
                        
    If mlngMode = 1712 Then
        mnuEditLine0.Visible = True
        mnuEditCheckBatch.Visible = mbln��Ҫ�˲� And InStr(1, mstrPrivs, ";�˲�;") <> 0
        mnuEditVerifyBatch.Visible = InStr(1, mstrPrivs, ";���;") <> 0
        mnuEditCheckLine.Visible = mnuEditCheck.Visible And (mnuEditVerify.Visible Or mnuEditStrike.Visible)
        If InStr(1, ";" & mstrPrivs & ";", ";���;") <> 0 Then
            mnuEditBill.Visible = True
            mnuEditReg.Visible = True
        Else
            mnuEditLine0.Visible = False
        End If
        If InStr(1, mstrPrivs, ";�������;") <> 0 Then
            mnuEditLine0.Visible = True
            mnuEditAcc.Visible = True
        Else
            If (mnuEditBill.Visible = False And mnuEditReg.Visible = False) Then mnuEditLine0.Visible = False
            mnuEditAcc.Visible = False
        End If
        If InStr(1, mstrPrivs, ";����ƻ���;") <> 0 Then
            mnuEditLine0.Visible = True
            mnuEditImport.Visible = True
        Else
            mnuEditImport.Visible = False
        End If
    ElseIf mlngMode = 1716 Then
        '�ƿⵥ
        mnuEditPrepare.Visible = InStr(1, ";" & mstrPrivs & ";", ";����;") <> 0
        mnuEditSend.Visible = mnuEditPrepare.Visible
        mnuEditBack.Visible = mnuEditPrepare.Visible
        mnuEditPrePareSp.Visible = mnuEditPrepare.Visible
            
        tlbTool.Buttons("PrepareSplit").Visible = mnuEditPrepare.Visible
        tlbTool.Buttons("Send").Visible = mnuEditPrepare.Visible
        tlbTool.Buttons("Back").Visible = mnuEditPrepare.Visible
        tlbTool.Buttons("Prepare").Visible = mnuEditPrepare.Visible
        
        If InStr(1, ";" & mstrPrivs & ";", ";���;") = 0 And _
           InStr(1, ";" & mstrPrivs & ";", ";����;") = 0 Then
            TabShow.TabVisible(1) = False
        End If
            
    Else
        mnuEditBill.Visible = False
        mnuEditBill.Visible = False
        mnuEditAcc.Visible = False
        mnuEditImport.Visible = False
        mnuEditLine0.Visible = False
    End If
    mnuEditRestore.Visible = mnuEditRestore.Visible And mlngMode = 1712
End Sub




Private Sub Cmd����_Click()
    Call mnuEditDisplay_Click
End Sub

Private Sub Form_Activate()
    PrintRange mstrMoneySum & Space(10) & vbCrLf & mstrPrintRange
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
End Sub

Private Sub Form_Load()
    Dim strOthers(0 To 13) As String
    Dim i As Integer
    mblnFirst = True
    
    mblnCostView = zlStr.IsHavePrivs(mstrPrivs, "�鿴�ɱ���")
    mbln����˲� = IIf((zlDatabase.GetPara("������Ҫ�˲������ƿ�", glngSys, 1722, "0")) = 0, False, True)
    
    mbln�ƿ���ȷ���� = IS�����ƿ�
'    If mlngMode = 1716 Then
'        mnuEditImport.Caption = "�����깺��(&I)"
'        mnuEditImport.Visible = True
'    End If
    
    For i = 0 To 13
        strOthers(i) = ""
    Next
    '������������
    strOthers(9) = "1901-01-01"
    strOthers(10) = "1901-01-01"
    
    '0-��¼״̬(�ƻ�����),1-��ʼ����,2-��������,3-����id,4-�Է�����id(��������id����Ʒ���(�ƻ���)),5-������,6-�����,7-��Ӧ��ID,8-������,9-��ʼ��������,10-������������,11-��ʼ��Ʊ��,12-������Ʊ��,13-������Ϣ
    mstrOthers = strOthers
    
    '�ָ�����
    Me.Caption = mstrTitle
    mstrPrintRange = "��ѯ��Χ:" & Format(sys.Currentdate, "yyyy��MM��dd��") & "��" & Format(sys.Currentdate, "yyyy��MM��dd��")
    
    PrintRange mstrMoneySum & Space(10) & vbCrLf & mstrPrintRange
    
    mintUnit = Val(IIf(Val(zlDatabase.GetPara("���ĵ�λ", glngSys, mlngMode, "0")) = 1, 1, 0))
    mstrOrder = zlDatabase.GetPara("��������", glngSys, mlngMode, "00")
  
    '���˺�:����С����ʽ����
    With mOraFMT
        .FM_�ɱ��� = GetFmtString(mintUnit, g_�ɱ���, True)
        .FM_��� = GetFmtString(mintUnit, g_���, True)
        .FM_���ۼ� = GetFmtString(mintUnit, g_�ۼ�, True)
        .FM_���� = GetFmtString(mintUnit, g_����, True)
        .FM_ɢװ���ۼ� = GetFmtString(0, g_�ۼ�, True)
    End With
    With mFMT
        .FM_�ɱ��� = GetFmtString(mintUnit, g_�ɱ���)
        .FM_��� = GetFmtString(mintUnit, g_���)
        .FM_���ۼ� = GetFmtString(mintUnit, g_�ۼ�)
        .FM_���� = GetFmtString(mintUnit, g_����)
        .FM_ɢװ���ۼ� = GetFmtString(0, g_�ۼ�)
    End With
        
    mnuViewLine3.Visible = mlngMode = 1712
    mnuViewColDefine.Visible = mlngMode = 1712
    mnuEditLine0.Visible = mlngMode = 1712
    mnuEditVerifySelect.Visible = mlngMode = 1712
    TabShow.Visible = (mlngMode = 1716)
    
    mnuEditTMPrint.Visible = mlngMode = 1712
    mnuEditLine3.Visible = mnuEditTMPrint.Visible
    
    '�����ⲿ�ļ�
    If InStr(mstrPrivs, ";�Ǽ�;") > 0 Then
        mnuEditImportFile.Visible = (mlngMode = 1712 Or mlngMode = 1714)
    Else
        mnuEditImportFile.Visible = False
    End If
    
    '��ֵ����
    With vsfCostlyInfo
        '.Cols = 4
        '.Rows = 2
        .RowHeight(0) = 300
        .AutoSizeMode = flexAutoSizeColWidth
        .Visible = False
        lblCostly.Visible = .Visible
    End With
    
    If mlngMode = 1712 Then
        If gobjPlugIn Is Nothing Then
            On Error Resume Next
            Set gobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
            If Not gobjPlugIn Is Nothing Then
                Call gobjPlugIn.Initialize(gcnOracle, glngSys, glngModul)
                If InStr(",438,0,", "," & err.Number & ",") = 0 Then
                    MsgBox "zlPlugIn ��Ҳ���ִ�� Initialize ʱ����" & vbCrLf & err.Number & vbCrLf & err.Description, vbInformation, gstrSysName
                End If
            End If
            err.Clear: On Error GoTo 0
        End If
         
        Call LoadPlugInMnu(Not gobjPlugIn Is Nothing)
    End If
    
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
        .Height = 400
        .Left = 0
        .Width = cbrTool.Width
        
    End With
   
    With TabShow
        .Left = 0
        .Top = cbrTool.Height
    End With
    
    With mshList
        .Top = IIf(cbrTool.Visible, cbrTool.Height, 0) + IIf(TabShow.Visible, TabShow.Height, 0)
        .Left = 0
        .Width = cbrTool.Width
        .Height = picSeparate_s.Top - .Top
    End With
    
    With Cmd����
        .Left = Me.ScaleWidth - .Width - 100
        .Top = mshList.Top + mshList.Height + 30
    End With
    
    If mlngMode = 1712 And vsfCostlyInfo.Visible Then
        '�����⹺�����Ҫ��ʾ��ֵ������Ϣ
        With mshDetail
            .Top = picSeparate_s.Top + picSeparate_s.Height + 100
            .Left = 0
            .Height = ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0) - lblCostly.Height - vsfCostlyInfo.Height
            .Width = cbrTool.Width
        End With
        With lblCostly
            .Top = mshDetail.Top + mshDetail.Height + 40
            .Left = 0
            .Width = cbrTool.Width
        End With
        With vsfCostlyInfo
            .Top = lblCostly.Top + lblCostly.Height
            .Left = 0
            .Width = cbrTool.Width
        End With
    Else
        With mshDetail
            .Top = picSeparate_s.Top + picSeparate_s.Height + 100
            .Left = 0
            .Height = ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
            .Width = cbrTool.Width
        End With
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName, mstrTitle
    
    Set gobjPlugIn = Nothing
End Sub


Private Sub imgLeft_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlControl.GetControlRect(mshDetail.hwnd)
    lngLeft = vRect.Left + imgLeft.Left
    lngTop = vRect.Top + imgLeft.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, mshDetail, lngLeft, lngTop, imgLeft.Height)

    zl_vsGrid_Para_Save mlngMode, mshDetail, mstrTitle, "mshDetail", False, True

End Sub

Private Sub mnuEditAcc_Click()
    Dim strNo As String
    Dim blnSuccess As Boolean
    
    With mshList
        strNo = .TextMatrix(.Row, 0)
        frmPurchaseCard.ShowCard Me, strNo, 7, .TextMatrix(.Row, GetCol(mshList, "��¼״̬")), mstrPrivs, blnSuccess
        If blnSuccess = True Then
            mnuViewRefresh_Click
        End If
    End With
End Sub

Private Sub mnuEditCheckBatch_Click()
    Dim frmPVB As New frmPurchaseVerifyBatch
    
    If Val(cboStock.Tag) > 0 Then
        frmPVB.ShowMe Me, mstrPrivs, 1, Val(cboStock.Tag)
        Call mnuViewRefresh_Click
    End If
End Sub

Private Sub mnuEditImport_Click()
    Dim blnSuccess As Boolean
    
    If mlngMode = 1712 Then
        frmPurchaseImportFromPlane.ShowCard Me, cboStock.Text, cboStock.ItemData(cboStock.ListIndex), mintUnit, InStr(mstrPrivs, "���пⷿ") <> 0, blnSuccess
    ElseIf mlngMode = 1716 Then
        frmPurchaseImportFromPlane.ShowCard Me, cboStock.Text, cboStock.ItemData(cboStock.ListIndex), mintUnit, InStr(mstrPrivs, "���пⷿ") <> 0, blnSuccess, 1, 1716, IIf(mbln�ƿ���ȷ����, 1, 0)
    End If
    
    If blnSuccess = True Then
        mnuViewRefresh_Click
    End If
End Sub
Private Sub mnuEditAdd_Click()
    Dim strNo As String
    Dim blnSuccess As Boolean

    strNo = ""
    '����
    Select Case mlngMode
        '�����⹺���
        Case 1712
            '���Popupmenuģ̬���壬���ܼ���Popupmenu
            If mblnPopupmenuCall Then
                mnuEditAdd.Tag = "1"
            Else
                frmPurchaseCard.ShowCard Me, strNo, 1, , mstrPrivs, blnSuccess
                mnuEditAdd.Tag = ""
            End If
        '�����������
        Case 1713
            frmSelfMakeCard.ShowCard Me, strNo, 1, , mstrPrivs, blnSuccess
        '�����������
        Case 1714
            frmOtherInputCard.ShowCard Me, strNo, 1, , mstrPrivs, blnSuccess
        '����۵���
        Case 1715
            frmDiffPriceAdjustCard.ShowCard Me, strNo, 1, , mstrPrivs, blnSuccess
        '�����ƿ�
        Case 1716
            frmTransferCard.ShowCard Me, strNo, 1, , mstrPrivs, blnSuccess
        '��������
        Case 1717
            frmDrawCard.ShowCard Me, strNo, 1, , mstrPrivs, blnSuccess
        '������������
        Case 1718
            frmOtherOutputCard.ShowCard Me, strNo, 1, , mstrPrivs, blnSuccess
    End Select
    
    If blnSuccess = True Then
        mnuViewRefresh_Click
    End If
End Sub

Private Sub mnuEditBack_Click()
    Dim strNo As String
    err = 0: On Error GoTo ErrHand
    '������һ��״̬
    '���δ����ֱ���˳���ֻ�ܴӷ��ͻ��˵����ϣ��ɱ��ϻ��˵��Ǳ��ϣ�
    strNo = mshList.TextMatrix(mshList.Row, 0)
    If strNo = "" Then Exit Sub
    
    gstrSQL = "ZL_�����ƿ�_BACK('" & strNo & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����")
    Call mnuViewRefresh_Click
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub mnuEditBill_Click()
    Dim strNo As String
    Dim blnSuccess As Boolean

    With mshList
        strNo = .TextMatrix(.Row, 0)
        frmPurchaseCard.ShowCard Me, strNo, 5, .TextMatrix(.Row, GetCol(mshList, "��¼״̬")), mstrPrivs, blnSuccess

        If blnSuccess = True Then
            mnuViewRefresh_Click
        End If
    End With
    
End Sub
Private Function CancelCheck() As Boolean
    '-----------------------------------------------------------------------------------------------------------------------
    '����:ȡ���˲鹦��
    '����:
    '����:ȡ���ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2007/05/13
    '-----------------------------------------------------------------------------------------------------------------------
    Dim blnSuccess As Boolean
    
    
    CancelCheck = False
    With mshList
        If mlngMode = 1712 Then
            ' Zl_�����⹺_Cancelcheck
            '  No_In In ҩƷ�շ���¼.NO%Type
            gstrSQL = "ZL_�����⹺_CANCELCHECK('" & .TextMatrix(.Row, 0) & "')"
        ElseIf mlngMode = 1717 Then
            '����
            gstrSQL = "Zl_��������_CancelVerify('" & .TextMatrix(.Row, 0) & "')"
        Else
            Exit Function
        End If
    End With
    
    err = 0: On Error GoTo ErrHandle
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption & "--ȡ���˲� ")
        
    CancelCheck = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub mnuEditCancelCheck_Click()
    '----------------------------------------------------------------------------------------------------------------------------
    '����:ȡ���˲�(ֻ���⹺���ž߱�ȡ���˲鹦��)
    '����:���˺�
    '����:2007/05/15
    '----------------------------------------------------------------------------------------------------------------------------
    Dim blnRefresh As Boolean
 
    If mlngMode = 1712 Or mlngMode = 1717 Then
        With mshList
            blnRefresh = (MsgBox("��ȷʵҪȡ���˲鵥�ݺ�Ϊ��" & .TextMatrix(.Row, 0) & "���ĵ�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes)
            If blnRefresh Then
                blnRefresh = CancelCheck()
                If blnRefresh Then mnuViewRefresh_Click
            End If
        End With
    End If
End Sub



Private Sub mnuEditCheck_Click()
    '----------------------------------------------------------------------------------------------------------------------------
    '����:�˲�ָ���ĵ���(ֻ���⹺���ž߱��˲鹦��)
    '����:���˺�
    '����:2007/05/15
    '----------------------------------------------------------------------------------------------------------------------------
    Dim strNo  As String
    Dim blnSuccess As Boolean
    With mshList
        strNo = mshList.TextMatrix(mshList.Row, 0)
        Select Case mlngMode
            Case 1712
                frmPurchaseCard.ShowCard Me, strNo, 9, .TextMatrix(.Row, GetCol(mshList, "��¼״̬")), mstrPrivs, blnSuccess
            Case 1717
                frmDrawCard.ShowCard Me, strNo, 5, .TextMatrix(.Row, GetCol(mshList, "��¼״̬")), mstrPrivs, blnSuccess
        End Select
        If blnSuccess = True Then
            mnuViewRefresh_Click
        End If
    End With
End Sub

Private Sub mnuEditImportFile_Click()
    Dim rsTmp As ADODB.Recordset
    Dim blnVirtualStock As Boolean
    
    If cboStock.ListCount < 1 Then Exit Sub
    
    On Error GoTo ErrHandle

    With frmPurchaseImportFile
        .EntryPort mlngMode, cboStock.ItemData(cboStock.ListIndex) & ";" & cboStock.Text
        .Show vbModal, Me
        If .Result Then
            Call mnuViewRefresh_Click
        End If
    End With
    Exit Sub
    
ErrHandle:
    Call ErrCenter
    Call SaveErrLog
End Sub

Private Sub mnuEditPrepare_Click()
    Dim strNo As String
    On Error GoTo ErrHand
    strNo = mshList.TextMatrix(mshList.Row, 0)
    
    If Trim(strNo) = "" Then Exit Sub
    
    gstrSQL = "zl_�����ƿ�_PREPARE('" & strNo & "','" & UserInfo.�û��� & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����")
    Call mnuViewRefresh_Click
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub mnuEditReg_Click()
    Dim strNo As String
    Dim blnSuccess As Boolean

    With mshList
        strNo = .TextMatrix(.Row, 0)
        frmPurchaseCard.ShowCard Me, strNo, 10, .TextMatrix(.Row, GetCol(mshList, "��¼״̬")), mstrPrivs, blnSuccess

        If blnSuccess = True Then
            mnuViewRefresh_Click
        End If
    End With
End Sub
Private Sub mnuEditRestore_Click()
    If mblnPopupmenuCall Then
        mnuEditRestore.Tag = "1"
    Else
        Dim strNo As String
        Dim blnSuccess As Boolean
        Call frmPurchaseCard.ShowCard(Me, strNo, 8, , mstrPrivs, blnSuccess)
        mnuEditRestore.Tag = ""
        If blnSuccess Then mnuViewRefresh_Click
    End If
End Sub

Private Sub mnuEditSend_Click()
    Dim strNo As String
    Dim blnSuccess As Boolean
    
    err = 0: On Error GoTo ErrHand
    strNo = mshList.TextMatrix(mshList.Row, 0)
    If Trim(strNo) = "" Then Exit Sub
    
    Call frmTransferCard.ShowCard(Me, strNo, 10, 1, mstrPrivs, blnSuccess)
    If blnSuccess Then mnuViewRefresh_Click
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub mnuEditTMPrint_Click()
    frmBarCodePrint.ShowMe Me, mOraFMT.FM_����, cboStock
End Sub


Private Sub mnuEditVerify_Click()
    '����
    
    Dim strNo As String
    Dim blnSuccess As Boolean
    
    With mshList
        strNo = .TextMatrix(.Row, 0)
        Select Case mlngMode
            Case 1712
                frmPurchaseCard.ShowCard Me, strNo, 3, .TextMatrix(.Row, GetCol(mshList, "��¼״̬")), mstrPrivs, blnSuccess
            Case 1713
                frmSelfMakeCard.ShowCard Me, strNo, 3, .TextMatrix(.Row, GetCol(mshList, "��¼״̬")), mstrPrivs, blnSuccess
            Case 1714
                frmOtherInputCard.ShowCard Me, strNo, 3, .TextMatrix(.Row, GetCol(mshList, "��¼״̬")), mstrPrivs, blnSuccess
            Case 1715
                frmDiffPriceAdjustCard.ShowCard Me, strNo, 3, .TextMatrix(.Row, GetCol(mshList, "��¼״̬")), mstrPrivs, blnSuccess
            Case 1716
                frmTransferCard.ShowCard Me, strNo, 3, .TextMatrix(.Row, GetCol(mshList, "��¼״̬")), mstrPrivs, blnSuccess
            Case 1717
                frmDrawCard.ShowCard Me, strNo, 3, .TextMatrix(.Row, GetCol(mshList, "��¼״̬")), mstrPrivs, blnSuccess
            Case 1718
                If Val(.TextMatrix(.Row, GetCol(mshList, "����ID"))) > 1 Then
                    MsgBox "�������ķ����Զ�����ĵ��ݲ������ֹ���ˣ�", vbInformation, gstrSysName
                    Exit Sub
                End If
                frmOtherOutputCard.ShowCard Me, strNo, 3, .TextMatrix(.Row, GetCol(mshList, "��¼״̬")), mstrPrivs, blnSuccess
        End Select
    End With
    If blnSuccess = True Then
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
    Dim rsCheck As New ADODB.Recordset
    Dim intCol��¼״̬ As Integer
     
    With mshList
        Select Case mlngMode
            Case 1712
                If Val(.TextMatrix(.Row, GetCol(mshList, "����ID"))) > 1 Then
                    MsgBox "�������ķ����Զ���������ⵥ�ݲ�����ɾ����", vbInformation, gstrSysName
                    Exit Sub
                End If
                
                strTitle = "�⹺��ⵥ"
            Case 1713
                strTitle = "������ⵥ"
            Case 1714
                strTitle = "������ⵥ"
            Case 1715
                strTitle = "����۵�����"
            Case 1716
                strTitle = "�����ƿⵥ"
            Case 1717
                strTitle = "�������õ�"
            Case 1718
                If Val(.TextMatrix(.Row, GetCol(mshList, "����ID"))) > 1 Then
                    MsgBox "�������ķ����Զ�����ĵ��ݲ�����ɾ����", vbInformation, gstrSysName
                    Exit Sub
                End If
                
                strTitle = "�����������ⵥ"
            Case 1719
                strTitle = "�����̵㵥"
        End Select
        
        On Error GoTo ErrHandle
        intRow = .Row
        StrBillNo = .TextMatrix(intRow, 0)
        intReturn = MsgBox("��ȷʵҪɾ�����ݺ�Ϊ��" & StrBillNo & "����" & strTitle & "��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)
        
        intRecord = .Rows - 1
        
        If intReturn = vbYes Then
            Select Case mlngMode
                Case 1712
                    gstrSQL = "zl_�����⹺_Delete('" & StrBillNo & "')"
                Case 1713
                    gstrSQL = "zl_���Ʋ������_Delete('" & StrBillNo & "')"
                Case 1714
                    gstrSQL = "zl_�����������_Delete('" & StrBillNo & "')"
                Case 1715
                    gstrSQL = "zl_���Ͽ���۵���_Delete('" & StrBillNo & "')"
                Case 1716
                    intCol��¼״̬ = GetCol(mshList, "��¼״̬")
                    If .TextMatrix(.Row, intCol��¼״̬) = 1 Then
                    '�ѱ��ϣ���д�������ˣ����ѷ��͵ĵ��ݣ���������ⷽ�޸Ĵ��൥��
                        If TabShow.Tab = 1 Then
                            If TestPrepare(StrBillNo) Then
                                MsgBox "�ѷ��͵ĵ��ݲ�����ɾ����", vbInformation, gstrSysName
                                Exit Sub
                            End If
                        End If
                    End If
                    gstrSQL = "zl_�����ƿ�_Delete('" & StrBillNo & "'," & .TextMatrix(.Row, intCol��¼״̬) & ")"
                    
'
'
'                    '�ȼ���ǲ������쵥
'                    gstrSQL = " Select Nvl(��ҩ��ʽ,0) ���� From ҩƷ�շ���¼ " & _
'                              " Where ����=19 And NO='" & strBillNo & "' And ���=1"
'                    Call OpenRecordset(rsCheck, "����ǲ������쵥")
'                    If rsCheck!���� = 0 Then
'                        gstrSQL = "zl_�����ƿ�_Delete('" & strBillNo & "')"
'                    Else
'                        gstrSQL = "zl_��������_Delete('" & strBillNo & "')"
'                    End If
'
                Case 1717
                    gstrSQL = "zl_��������_Delete('" & StrBillNo & "')"
                Case 1718
                    gstrSQL = "zl_������������_Delete('" & StrBillNo & "')"
                Case 1719
                    gstrSQL = "zl_�����̵�_Delete('" & StrBillNo & "')"
                Case Else
                
            End Select
            If gstrSQL = "" Then Exit Sub
            zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
            
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
            .ColSel = .Cols - 1
            mshlist_EnterCell
        End If
    End With
    stbThis.Panels(2).Text = "��ǰ����" & intRecord & "�ŵ���"
    Exit Sub

ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    
End Sub

Private Sub mnuEditDisplay_Click()
    '�鿴����
    
    Dim strNo As String
    With mshList
        strNo = .TextMatrix(.Row, 0)
        If strNo = "" Then Exit Sub
        Select Case mlngMode
            Case 1712
                frmPurchaseCard.ShowCard Me, strNo, 4, Val(.TextMatrix(.Row, GetCol(mshList, "��¼״̬"))), mstrPrivs
            Case 1713
                frmSelfMakeCard.ShowCard Me, strNo, 4, Val(.TextMatrix(.Row, GetCol(mshList, "��¼״̬"))), mstrPrivs
            Case 1714
                frmOtherInputCard.ShowCard Me, strNo, 4, Val(.TextMatrix(.Row, GetCol(mshList, "��¼״̬"))), mstrPrivs
            Case 1715
                frmDiffPriceAdjustCard.ShowCard Me, strNo, 4, Val(.TextMatrix(.Row, GetCol(mshList, "��¼״̬"))), mstrPrivs
            Case 1716
                frmTransferCard.ShowCard Me, strNo, 4, Val(.TextMatrix(.Row, GetCol(mshList, "��¼״̬"))), mstrPrivs
            Case 1717
                frmDrawCard.ShowCard Me, strNo, 4, Val(.TextMatrix(.Row, GetCol(mshList, "��¼״̬"))), mstrPrivs
            Case 1718
                frmOtherOutputCard.ShowCard Me, strNo, 4, Val(.TextMatrix(.Row, GetCol(mshList, "��¼״̬"))), mstrPrivs
        End Select
        
    End With
    
End Sub

Private Sub mnuEditStrike_Click()
    Dim blnPurchase As Boolean, blnRefresh As Boolean
    
    
    '������⹺(blnPurchaseΪ��)����ֱ�ӽ������
    'ѯ���Ƿ����(blnPurchaseΪ��ʾ�򷵻�ֵ)������������
    blnPurchase = (InStr(1, "1712,1714,1716,1717,1718", mlngMode) <> 0)
    With mshList
        If Not blnPurchase Then
            blnPurchase = (MsgBox("��ȷʵҪ�������ݺ�Ϊ��" & .TextMatrix(.Row, 0) & "���ĵ�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes)
        End If
        If blnPurchase Then
            blnRefresh = StrikeSave
            If blnRefresh Then mnuViewRefresh_Click
        End If
    End With
End Sub
Private Function CheckSelfMakeStock(ByVal str���ݺ� As String) As Boolean
    '------------------------------------------------------------------------------
    '����:�ڳ���ʱ������������Ŀ��������Ƿ����
    '����:����,����true,���򷵻�False
    '����:���˺�
    '����:2008/02/15
    '------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, int����� As Integer
    err = 0: On Error GoTo ErrHand:
    gstrSQL = "" & _
        "   Select ҩƷid,nvl(����,0) as ����,�ⷿid,sum(ʵ������) as ʵ������ " & _
        "   From ҩƷ�շ���¼ A " & _
        "   where ���� = 16 And A.NO = [1] And A.��¼״̬ = 1 And A.���ϵ��=1" & _
        "   Group by ҩƷID,nvl(����,0),�ⷿID"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, str���ݺ�)
    With rsTemp
        If .EOF Then Exit Function
        int����� = Get������(cboStock.ItemData(cboStock.ListIndex))
        Do While Not .EOF
            If Check��������(Val(zlStr.Nvl(rsTemp!�ⷿID)), Val(zlStr.Nvl(rsTemp!ҩƷID)), _
                Val(zlStr.Nvl(rsTemp!����, 0)), Val(zlStr.Nvl(rsTemp!ʵ������)), int�����) = False Then Exit Function
            .MoveNext
        Loop
    End With
    CheckSelfMakeStock = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function
Private Function StrikeSave() As Boolean
    Dim blnSuccess As Boolean
    
    StrikeSave = False
    With mshList
        Select Case mlngMode
            Case 1712
                frmPurchaseCard.ShowCard Me, .TextMatrix(.Row, 0), 6, mshList.TextMatrix(mshList.Row, GetCol(mshList, "��¼״̬")), mstrPrivs, blnSuccess
                StrikeSave = blnSuccess
                Exit Function
            Case 1713
                If CheckSelfMakeStock(.TextMatrix(.Row, 0)) = False Then Exit Function
                gstrSQL = "zl_���Ʋ������_Strike('" & .TextMatrix(.Row, 0) & "','" & UserInfo.�û��� & "')"
            Case 1714
                frmOtherInputCard.ShowCard Me, .TextMatrix(.Row, 0), 6, mshList.TextMatrix(mshList.Row, GetCol(mshList, "��¼״̬")), mstrPrivs, blnSuccess
                StrikeSave = blnSuccess
                Exit Function
            Case 1715
                gstrSQL = "zl_���Ͽ���۵���_Strike('" & .TextMatrix(.Row, 0) & "','" & UserInfo.�û��� & "')"
            Case 1716
                If mnuEditStrike.Caption = "����(&K)" Then
                    mint������ʽ = 0
                ElseIf mnuEditStrike.Caption = "�������(&K)" Then
                    mint������ʽ = 1
                ElseIf mnuEditStrike.Caption = "��˳���(&K)" Then
                    mint������ʽ = 2
                End If
                
                frmTransferCard.ShowCard Me, .TextMatrix(.Row, 0), 6, mshList.TextMatrix(mshList.Row, GetCol(mshList, "��¼״̬")), mstrPrivs, blnSuccess, mint������ʽ
                StrikeSave = blnSuccess
                Exit Function
            Case 1717
                frmDrawCard.ShowCard Me, .TextMatrix(.Row, 0), 6, mshList.TextMatrix(mshList.Row, GetCol(mshList, "��¼״̬")), mstrPrivs, blnSuccess
                StrikeSave = blnSuccess
                Exit Function
            Case 1718
                frmOtherOutputCard.ShowCard Me, .TextMatrix(.Row, 0), 6, mshList.TextMatrix(mshList.Row, GetCol(mshList, "��¼״̬")), mstrPrivs, blnSuccess
                StrikeSave = blnSuccess
                Exit Function
            Case 1719
                gstrSQL = "zl_�����̵�_Strike('" & .TextMatrix(.Row, 0) & "','" & UserInfo.�û��� & "')"
            Case Else
            
        End Select
        
        On Error GoTo ErrHandle
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption & "--���� ")
    End With
    StrikeSave = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog

End Function

Private Sub mnuEditModify_Click()
    '�޸�
    Dim strNo As String
    Dim blnSuccess As Boolean
    
    blnSuccess = False
    With mshList
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        strNo = .TextMatrix(.Row, 0)
        
        Select Case mlngMode
            Case 1712
                If Val(.TextMatrix(.Row, GetCol(mshList, "����ID"))) > 1 Then
                    MsgBox "�������ķ����Զ���������ⵥ�ݲ������޸ģ�", vbInformation, gstrSysName
                    Exit Sub
                End If
                frmPurchaseCard.ShowCard Me, strNo, 2, Val(mshList.TextMatrix(mshList.Row, GetCol(mshList, "��¼״̬"))), mstrPrivs, blnSuccess
            Case 1713
                frmSelfMakeCard.ShowCard Me, strNo, 2, Val(mshList.TextMatrix(mshList.Row, GetCol(mshList, "��¼״̬"))), mstrPrivs, blnSuccess
            Case 1714
                frmOtherInputCard.ShowCard Me, strNo, 2, Val(mshList.TextMatrix(mshList.Row, GetCol(mshList, "��¼״̬"))), mstrPrivs, blnSuccess
            Case 1715
                frmDiffPriceAdjustCard.ShowCard Me, strNo, 2, Val(mshList.TextMatrix(mshList.Row, GetCol(mshList, "��¼״̬"))), mstrPrivs, blnSuccess
            Case 1716
            
                '�ѱ��ϣ���д�������ˣ����ѷ��͵ĵ��ݣ���������ⷽ�޸Ĵ��൥��
                If TabShow.Tab = 1 Then
                    If TestPrepare(strNo) Then
                        MsgBox "�ѷ��͵ĵ��ݲ������޸ģ�", vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
                
                frmTransferCard.ShowCard Me, strNo, 2, Val(mshList.TextMatrix(mshList.Row, GetCol(mshList, "��¼״̬"))), mstrPrivs, blnSuccess, mint������ʽ
            Case 1717
                frmDrawCard.ShowCard Me, strNo, 2, Val(mshList.TextMatrix(mshList.Row, GetCol(mshList, "��¼״̬"))), mstrPrivs, blnSuccess
            Case 1718
                If Val(.TextMatrix(.Row, GetCol(mshList, "����ID"))) > 1 Then
                    MsgBox "�������ķ����Զ�����ĵ��ݲ������޸ģ�", vbInformation, gstrSysName
                    Exit Sub
                End If
                frmOtherOutputCard.ShowCard Me, strNo, 2, Val(mshList.TextMatrix(mshList.Row, GetCol(mshList, "��¼״̬"))), mstrPrivs, blnSuccess
        End Select
        If blnSuccess = True Then
            mnuViewRefresh_Click
        End If
    End With
End Sub

Private Sub mnuEditVerifyBatch_Click()
    Dim frmPVB As New frmPurchaseVerifyBatch
    
''    If mbln��Ҫ�˲� Then
''        If InStr(1, mstrPrivs, ";�˲�;") <= 0 And InStr(1, mstrPrivs, ";���;") <= 0 Then
''            MsgBox "�����������Ĳ������õġ������⹺��Ҫ�˲顱��������û�С��˲顱Ȩ�ޣ�Ҳû�С���ˡ�Ȩ�ޣ�", vbInformation, gstrSysName
''            Exit Sub
''        End If
''    Else
''        If InStr(1, mstrPrivs, ";���;") <= 0 Then
''            MsgBox "��û�С���ˡ�Ȩ�ޣ�����ϵ����Ա��", vbInformation, gstrSysName
''            Exit Sub
''        End If
''    End If
    If Val(cboStock.Tag) > 0 Then
        frmPVB.ShowMe Me, mstrPrivs, IIf(mbln��Ҫ�˲�, 2, 0), Val(cboStock.Tag)
        Call mnuViewRefresh_Click
    End If

End Sub

Private Sub mnuEditVerifySelect_Click()
    frmPurchaseVerifySelect.ShowMe Me, mStr�ⷿ, cboStock.ListIndex
End Sub

Private Sub mnuFileBillPreview_Click()
    Dim int��λϵ�� As Integer
    
    On Error GoTo ErrHandle
    With mshList
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        
        Select Case mintUnit
            Case 0
                int��λϵ�� = 0
            Case 1
                int��λϵ�� = 1
        End Select
        
        Select Case mlngMode
            Case 1712
                Dim rsTemp As New ADODB.Recordset
                Dim bln�˿ⵥ As Boolean
                
                gstrSQL = "Select Nvl(��ҩ��ʽ,0) ��־ From ҩƷ�շ���¼ Where NO=[1] And ��¼״̬=[2] and ����=15 And Rownum<2"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[�ж��Ƿ����˿ⵥ]", .TextMatrix(.Row, 0), Val(.TextMatrix(.Row, GetCol(mshList, "��¼״̬"))))
                
                bln�˿ⵥ = (rsTemp!��־ = 1)
            
                ReportOpen gcnOracle, glngSys, "zl1_bill_1712", Me, "���ݱ��=" & .TextMatrix(.Row, 0), "��¼״̬=" & Val(.TextMatrix(.Row, GetCol(mshList, "��¼״̬"))), "��λϵ��=" & int��λϵ��, IIf(bln�˿ⵥ, "�����˻���", "�����⹺��ⵥ"), 1
            Case 1713
                ReportOpen gcnOracle, glngSys, "zl1_bill_1713", Me, "���ݱ��=" & .TextMatrix(.Row, 0), "��¼״̬=" & Val(.TextMatrix(.Row, GetCol(mshList, "��¼״̬"))), "��λϵ��=" & int��λϵ��, 1
            Case 1714
                ReportOpen gcnOracle, glngSys, "zl1_bill_1714", Me, "���ݱ��=" & .TextMatrix(.Row, 0), "��¼״̬=" & Val(.TextMatrix(.Row, GetCol(mshList, "��¼״̬"))), "��λϵ��=" & int��λϵ��, 1
            Case 1715
                ReportOpen gcnOracle, glngSys, "zl1_bill_1715", Me, "���ݱ��=" & .TextMatrix(.Row, 0), "��¼״̬=" & Val(.TextMatrix(.Row, GetCol(mshList, "��¼״̬"))), "��λϵ��=" & int��λϵ��, 1
            Case 1716
                ReportOpen gcnOracle, glngSys, "zl1_bill_1716", Me, "���ݱ��=" & .TextMatrix(.Row, 0), "��¼״̬=" & Val(.TextMatrix(.Row, GetCol(mshList, "��¼״̬"))), "��λϵ��=" & int��λϵ��, 1
            Case 1717
                ReportOpen gcnOracle, glngSys, "zl1_bill_1717", Me, "���ݱ��=" & .TextMatrix(.Row, 0), "��¼״̬=" & Val(.TextMatrix(.Row, GetCol(mshList, "��¼״̬"))), "��λϵ��=" & int��λϵ��, 1
            Case 1718
                ReportOpen gcnOracle, glngSys, "zl1_bill_1718", Me, "���ݱ��=" & .TextMatrix(.Row, 0), "��¼״̬=" & Val(.TextMatrix(.Row, GetCol(mshList, "��¼״̬"))), "��λϵ��=" & int��λϵ��, 1
            Case 1719
                ReportOpen gcnOracle, glngSys, "zl1_bill_1719", Me, "���ݱ��=" & .TextMatrix(.Row, 0), "��¼״̬=" & Val(.TextMatrix(.Row, GetCol(mshList, "��¼״̬"))), "��λϵ��=" & int��λϵ��, 1
            Case Else
            
        End Select
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuFileBillPrint_Click()
    Dim strUnit As String
    Dim int��λϵ�� As Integer
    
    On Error GoTo ErrHandle
    With mshList
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        Select Case mintUnit
            Case 0
                int��λϵ�� = 0
            Case 1
                int��λϵ�� = 1
        End Select
        
        Select Case mlngMode
            Case 1712
                Dim rsTemp As New ADODB.Recordset
                Dim bln�˿ⵥ As Boolean
                gstrSQL = "Select Nvl(��ҩ��ʽ,0) ��־ From ҩƷ�շ���¼ Where NO=[1] And ��¼״̬=[2] and ����=15 And Rownum<2"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[�ж��Ƿ����˿ⵥ]", .TextMatrix(.Row, 0), Val(.TextMatrix(.Row, GetCol(mshList, "��¼״̬"))))
                bln�˿ⵥ = (rsTemp!��־ = 1)
                ReportOpen gcnOracle, glngSys, "zl1_bill_1712", Me, "���ݱ��=" & .TextMatrix(.Row, 0), "��¼״̬=" & .TextMatrix(.Row, GetCol(mshList, "��¼״̬")), "��λϵ��=" & int��λϵ��, IIf(bln�˿ⵥ, "�����˻���", "�����⹺��ⵥ"), 2
                
            Case 1713
                ReportOpen gcnOracle, glngSys, "zl1_bill_1713", Me, "���ݱ��=" & .TextMatrix(.Row, 0), "��¼״̬=" & .TextMatrix(.Row, GetCol(mshList, "��¼״̬")), "��λϵ��=" & int��λϵ��, 2
            Case 1714
                ReportOpen gcnOracle, glngSys, "zl1_bill_1714", Me, "���ݱ��=" & .TextMatrix(.Row, 0), "��¼״̬=" & .TextMatrix(.Row, GetCol(mshList, "��¼״̬")), "��λϵ��=" & int��λϵ��, 2
            Case 1715
                ReportOpen gcnOracle, glngSys, "zl1_bill_1715", Me, "���ݱ��=" & .TextMatrix(.Row, 0), "��¼״̬=" & .TextMatrix(.Row, GetCol(mshList, "��¼״̬")), "��λϵ��=" & int��λϵ��, 2
            Case 1716
                ReportOpen gcnOracle, glngSys, "zl1_bill_1716", Me, "���ݱ��=" & .TextMatrix(.Row, 0), "��¼״̬=" & .TextMatrix(.Row, GetCol(mshList, "��¼״̬")), "��λϵ��=" & int��λϵ��, 2
            Case 1717
                ReportOpen gcnOracle, glngSys, "zl1_bill_1717", Me, "���ݱ��=" & .TextMatrix(.Row, 0), "��¼״̬=" & .TextMatrix(.Row, GetCol(mshList, "��¼״̬")), "��λϵ��=" & int��λϵ��, 2
            Case 1718
                ReportOpen gcnOracle, glngSys, "zl1_bill_1718", Me, "���ݱ��=" & .TextMatrix(.Row, 0), "��¼״̬=" & .TextMatrix(.Row, GetCol(mshList, "��¼״̬")), "��λϵ��=" & int��λϵ��, 2
            Case 1719
                ReportOpen gcnOracle, glngSys, "zl1_bill_1719", Me, "���ݱ��=" & .TextMatrix(.Row, 0), "��¼״̬=" & .TextMatrix(.Row, GetCol(mshList, "��¼״̬")), "��λϵ��=" & int��λϵ��, 2
            Case Else
            
        End Select
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuFileExcel_Click()
    '�����Excel
    
    If Me.ActiveControl Is mshList Then
        mshList.Redraw = False
        subPrint 3
        mshList.Redraw = True
        mshList.Col = 0
        mshList.ColSel = mshList.Cols - 1
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
    '��������
    frmParaset.���ò��� mlngMode, mstrPrivs, Me, Me.Tag
    mintUnit = IIf(Val(zlDatabase.GetPara("���ĵ�λ", glngSys, mlngMode, "0")) = 1, 1, 0)
    
    mint������˷�ʽ = 0
    If mlngMode = 1717 Then
        mint������˷�ʽ = Val(zlDatabase.GetPara("�������", glngSys, mlngMode, "0"))
    End If
    
    SetPopedom
    Call SetMenu
    mstrOrder = zlDatabase.GetPara("��������", glngSys, mlngMode, "00")
    mintUnit = Val(IIf(Val(zlDatabase.GetPara("���ĵ�λ", glngSys, mlngMode, "0")) = 1, 1, 0))
    
    mintFindDay = Val(zlDatabase.GetPara("��ѯ����", glngSys, mlngMode, 1))
    mdtStartDate = Format(DateAdd("d", -mintFindDay, sys.Currentdate), "yyyy-MM-dd")
    mdtEndDate = Format(sys.Currentdate, "yyyy-MM-dd")
  
    '���˺�:����С����ʽ����
    With mOraFMT
        .FM_�ɱ��� = GetFmtString(mintUnit, g_�ɱ���, True)
        .FM_��� = GetFmtString(mintUnit, g_���, True)
        .FM_���ۼ� = GetFmtString(mintUnit, g_�ۼ�, True)
        .FM_���� = GetFmtString(mintUnit, g_����, True)
        .FM_ɢװ���ۼ� = GetFmtString(0, g_�ۼ�, True)
    End With
    With mFMT
        .FM_�ɱ��� = GetFmtString(mintUnit, g_�ɱ���)
        .FM_��� = GetFmtString(mintUnit, g_���)
        .FM_���ۼ� = GetFmtString(mintUnit, g_�ۼ�)
        .FM_���� = GetFmtString(mintUnit, g_����)
        .FM_ɢװ���ۼ� = GetFmtString(0, g_�ۼ�)
    End With
    
    Call GetList(mstrFind)
End Sub

Private Sub mnuFilePreView_Click()
    '��ӡԤ��
    mshList.Redraw = False
    subPrint 2
    mshList.Redraw = True
    mshList.Col = 0
    mshList.ColSel = mshList.Cols - 1
    
End Sub

Private Sub mnuFilePrint_Click()
    '��ӡ
    mshList.Redraw = False
    subPrint 1
    mshList.Redraw = True
    mshList.Col = 0
    mshList.ColSel = mshList.Cols - 1
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
    Dim StrWinName As String
    With mshList
        Select Case mlngMode
            Case 1712
                StrWinName = "frmMainList1"
            Case 1713
                StrWinName = "frmMainList2"
            Case 1714
                StrWinName = "frmMainList3"
            Case 1715
                StrWinName = "frmMainList4"
            Case 1716
                StrWinName = "frmMainList5"
            Case 1717
                StrWinName = "frmMainList6"
            Case 1718
                StrWinName = "frmMainList7"
            Case 1719
                StrWinName = "frmMainList8"
        End Select
    End With
    Call ShowHelp(App.ProductName, Me.hwnd, StrWinName, Int(glngSys / 100))
End Sub

Private Sub mnuHelpWebHome_Click()
    '������ҳ
    Call zlHomePage(Me.hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    '���ͷ���
    Call zlMailTo(Me.hwnd)
End Sub

Private Sub mnuPlugItem_Click(Index As Integer)
    Call ExcPlugInFun(mnuPlugItem(Index).Tag)
End Sub

Private Sub mnuViewColDefine_Click()
    Dim strColumn_All As String, strColumn_Select As String
    
    Select Case mlngMode
    Case 1712           '�����⹺������
        strColumn_All = "����,0|��Ʒ��,1|���,1|����,1|��׼�ĺ�,1|����,0|��������,1|��Ʒ����,1|�������,1|���ʧЧ��,1|Ч��,0|ע��֤��,1|��λ,1|����,0|ָ��������,1|�ɹ���,0|����,1|" & _
                        "�ӳ���,1|�����,0|������,0|�ۼ�,0|�ۼ۽��,0|���,0|���۵�λ,1|���ۼ�,1|���۽��,1|���۲��,1|�������,1|���ս���,1|��Ʊ��,0|��Ʊ����,0|��Ʊ����,0|��Ʊ���,0"
    Case 1713
        strColumn_All = "����,0|���,1|����,1|����,0|��������,1|�������,1|���ʧЧ��,1|Ч��,0|ע��֤��,1|��λ,1|����,0|ָ��������,1|�ɹ���,1|����,1|" & _
                        "�ӳ���,1|�����,0|������,0|�ۼ�,0|�ۼ۽��,0|���,0|���۵�λ,1|���ۼ�,1|���۽��,1|���۲��,1|�������,1|��Ʊ��,0|��Ʊ����,0|��Ʊ����,0|��Ʊ���,0"
    
    Case Else
        Exit Sub
    End Select

    'ȡ��ѡ���е���Ϣ'Me.Caption
    strColumn_Select = zlDatabase.GetPara("ѡ����", glngSys, mlngMode)
    If Not frmColSet.ShowMe(Me, strColumn_All, strColumn_Select) Then Exit Sub
    Call zlDatabase.SetPara("ѡ����", Split(strColumn_Select, "||")(0), glngSys, mlngMode)
    Call zlDatabase.SetPara("������", Split(strColumn_Select, "||")(1), glngSys, mlngMode)
End Sub

Private Sub mnuViewRefresh_Click()
    'ˢ��
    GetList mstrFind
End Sub

Private Sub mnuViewSearch_Click()
    '����
    Dim strFind As String
    Dim strOthers() As String

    Select Case mlngMode
        Case 1715, 1716, 1717, 1718, 1719
            strFind = FrmTransferSearch.GetSearch(Me, mlngMode, mdtStartDate, mdtEndDate, mdtVerifyStart, mdtVerifyEnd, mstrPrivs, strOthers)
        Case 1712
            strFind = FrmPurchaseSearch.GetSearch(Me, mdtStartDate, mdtEndDate, mdtVerifyStart, mdtVerifyEnd, strOthers, mstr��ֵ�Ĳ�, mint�޷�Ʊ, mint�з�Ʊ)
        Case 1713
            strFind = FrmSelfMakeSearch.GetSearch(Me, mdtStartDate, mdtEndDate, mdtVerifyStart, mdtVerifyEnd, strOthers)
        Case 1714
            strFind = FrmOtherInputSearch.GetSearch(Me, mdtStartDate, mdtEndDate, mdtVerifyStart, mdtVerifyEnd, strOthers)
    End Select
    
    If strFind <> "" Then
        mstrFind = strFind
        mstrOthers = strOthers
        
        GetList mstrFind
        If Format(mdtStartDate, "yyyy-mm-dd") = "1901-01-01" And Format(mdtVerifyStart, "yyyy-mm-dd") = "1901-01-01" Then
            mstrPrintRange = ""
        ElseIf Format(mdtStartDate, "yyyy-mm-dd") <> "1901-01-01" And Format(mdtVerifyStart, "yyyy-mm-dd") <> "1901-01-01" Then
            mstrPrintRange = "��ѯ��Χ:�������� " & Format(mdtStartDate, "yyyy��MM��dd��") & "��" & Format(mdtEndDate, "yyyy��MM��dd��") & "  ������� " & Format(mdtVerifyStart, "yyyy��MM��dd��") & "��" & Format(mdtVerifyEnd, "yyyy��MM��dd��")
        ElseIf Format(mdtStartDate, "yyyy-mm-dd") <> "1901-01-01" Then
            mstrPrintRange = "��ѯ��Χ:�������� " & Format(mdtStartDate, "yyyy��MM��dd��") & "��" & Format(mdtEndDate, "yyyy��MM��dd��")
        ElseIf Format(mdtVerifyStart, "yyyy-mm-dd") <> "1901-01-01" Then
            mstrPrintRange = "��ѯ��Χ:������� " & Format(mdtVerifyStart, "yyyy��MM��dd��") & "��" & Format(mdtVerifyEnd, "yyyy��MM��dd��")
        End If
        PrintRange mstrMoneySum & Space(10) & vbCrLf & mstrPrintRange
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

Private Sub mshDetail_AfterMoveColumn(ByVal Col As Long, Position As Long)
      zl_vsGrid_Para_Save mlngMode, mshDetail, mstrTitle, "mshDetail", False, True
End Sub

Private Sub mshDetail_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
      zl_vsGrid_Para_Save mlngMode, mshDetail, mstrTitle, "mshDetail", False, True
End Sub

Private Sub mshDetail_Click()
'    With mshDetail
'         If .Row < 1 Or .TextMatrix(.Row, 0) = "" Then Exit Sub
'         If .MouseRow = 0 Then
'            DetailSort          '������
'            Exit Sub
'         End If
'    End With
End Sub

Private Sub mshDetail_EnterCell()
    '�����⹺���
    On Error GoTo ErrHandle
    If mlngMode = 1712 Then
        Dim rsTmp As ADODB.Recordset
        Dim strTmp As String
        
        vsfCostlyInfo.Visible = False
        lblCostly.Visible = False
        
        If mshDetail.Rows <= 1 Or mshDetail.Row <= 0 Then
            Call Form_Resize
            Exit Sub
        End If
        
        strTmp = "Select A.����, A.��������, A.סԺ��, A.����, nvl(C.��ֵ����,0) ��ֵ���� " _
               & "From �շ���¼������Ϣ A, ҩƷ�շ���¼ B, �������� C " _
               & "Where A.�շ�id = B.ID And B.ҩƷid = C.����id and A.�շ�id=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strTmp, Me.Caption, mshDetail.TextMatrix(mshDetail.Row, mshDetail.ColIndex("�շ�ID")))
        Set vsfCostlyInfo.DataSource = rsTmp
        If rsTmp.RecordCount > 0 Then
            If rsTmp!��ֵ���� = 1 Then
                vsfCostlyInfo.Visible = True
                lblCostly.Visible = True
                vsfCostlyInfo.ColHidden(vsfCostlyInfo.ColIndex("��ֵ����")) = True
                vsfCostlyInfo.ColHidden(0) = True
            End If
        End If
        rsTmp.Close
        With vsfCostlyInfo
            .ColWidth(.ColIndex("����")) = 2000
            .ColWidth(.ColIndex("��������")) = 2000
            .ColWidth(.ColIndex("סԺ��")) = 2000
            .ColWidth(.ColIndex("����")) = 1000
            .ColAlignment(.ColIndex("סԺ��")) = flexAlignLeftCenter
        End With
        Call Form_Resize
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
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
    Dim rsTemp As New Recordset
    Dim strUnitQuantity As String               '��λ��������ʽ����
    Dim IntBill As Integer                      '��������  �磺1���⹺��⣻2��
    Dim strUnit As String                       '��λ����:�����ﵥλ��סԺ��λ��
    Dim str��װϵ�� As String
    Dim str���ۼ� As String
    Dim intCol As Integer
    Dim str���� As String
    Dim str���� As String
    Dim strTemp As String
    
'    If mlastRow = mshList.Row Then Exit Sub
    mlastRow = mshList.Row
        
    On Error GoTo ErrHandle
'    If mshList.Row >= 1 And LTrim(mshList.TextMatrix(mshList.Row, 0)) <> "" Then
    
        mshList.Col = 0
        mshList.ColSel = mshList.Cols - 1
        If mshList.RowIsVisible(mshList.Row) = False Then
           mshList.TopRow = mshList.Row
        End If
        
        If Mid(mstrOrder, 1, 1) = "0" Then
            str���� = " ���"
        ElseIf Mid(mstrOrder, 1, 1) = "1" Then
            str���� = " ������Ϣ"
        ElseIf Mid(mstrOrder, 1, 1) = "2" Then
            str���� = " ����"
        End If
        
        If Mid(mstrOrder, 2, 1) = "0" Then
            str���� = str���� & " asc"
        ElseIf Mid(mstrOrder, 2, 1) = "1" Then
            str���� = str���� & " desc"
        End If
        Select Case mintUnit
            Case 0
                strUnitQuantity = "ltrim(rtrim((to_char(A.ʵ������ ," & mOraFMT.FM_���� & ")))) AS ����," _
                    & "D.���㵥λ AS ��λ,"
                str��װϵ�� = "1"
            Case 1
                strUnitQuantity = "ltrim(rtrim((to_char(A.ʵ������ / b.����ϵ��," & mOraFMT.FM_���� & ")))) AS ����," _
                    & "B.��װ��λ AS ��λ,"
                str��װϵ�� = "B.����ϵ��"
        End Select
        
        Dim int���� As Integer
        Select Case mlngMode
            Case 1712       '�����⹺���
                IntBill = 1
                strTemp = ""
                
                If mint�޷�Ʊ = 1 And mint�з�Ʊ = 0 Then
                    strTemp = " and c.��Ʊ�� is null "
                End If
                If mint�з�Ʊ = 1 And mint�޷�Ʊ = 0 Then
                    strTemp = " and c.��Ʊ�� is not null "
                End If
                
                If mintUnit <> 0 Then
                    str���� = "���,������Ϣ,��Ʒ��,���,����,��׼�ĺ�,����,��������,ʧЧ��,ע��֤��,����,��λ,�����,������,����,�ۼ�,�ۼ۽��,���,���ۼ�,���۵�λ,���۽��,���۲��,�������,��Ʊ��,��Ʊ����,��Ʊ����,�������,��Ʊ���,�շ�id,��Ʒ����,�ڲ�����"
                    str���ۼ� = "" & _
                        "                   ltrim(rtrim(to_char(A.���ۼ�* " & str��װϵ�� & "," & mOraFMT.FM_���ۼ� & "))) as �ۼ� , " & _
                        "                   ltrim(rtrim(to_char(A.���۽��," & mOraFMT.FM_��� & ")))  as �ۼ۽��, " & _
                        "                   ltrim(rtrim(to_char(A.���, " & mOraFMT.FM_��� & "))) as ���," & _
                        "                   ltrim(rtrim(to_char(A.���ۼ�," & mOraFMT.FM_ɢװ���ۼ� & "))) as ���ۼ� , " & _
                        "                   D.���㵥λ as ���۵�λ , " & _
                        "                   ltrim(rtrim(to_char(A.���۽��," & mOraFMT.FM_��� & ")))  as ���۽��," & _
                        "                   ltrim(rtrim(to_char(A.���," & mOraFMT.FM_��� & "))) as ���۲��,"
                Else
                    str���� = "���,������Ϣ,��Ʒ��,���,����,��׼�ĺ�,����,��������,ʧЧ��,ע��֤��,����,��λ,�����,������,����,�ۼ�,�ۼ۽��,���,�������,��Ʊ��,��Ʊ����,�������,��Ʊ���,�շ�id,��Ʒ����,�ڲ�����"
                    str���ۼ� = "" & _
                    "                   ltrim(rtrim(to_char(A.���ۼ�*" & str��װϵ�� & "," & mOraFMT.FM_���ۼ� & "))) as �ۼ� , " & _
                    "                   ltrim(rtrim(to_char(A.���۽��," & mOraFMT.FM_��� & ")))  as �ۼ۽��, " & _
                    "                   ltrim(rtrim(to_char(A.���," & mOraFMT.FM_��� & "))) as ���,"

                End If
                gstrSQL = "" & _
                    "   SELECT " & str���� & _
                    "   From (  SELECT distinct a.���, '[' || D.���� || ']' || D.���� AS ������Ϣ,E.���� As ��Ʒ��,D.���,d.����,zlSpellCode(d.����) ����, A.����,A.��׼�ĺ�, A.����, to_char(A.��������,'yyyy-mm-dd') as ��������, to_char(A.Ч��,'yyyy-mm-dd') as ʧЧ�� ,a.ע��֤��," & _
                                        strUnitQuantity & _
                    "                   ltrim(rtrim(to_char((A.�ɱ���*" & str��װϵ�� & ")," & mOraFMT.FM_�ɱ��� & "))) AS �����, ltrim(rtrim(to_char(A.�ɱ����," & mOraFMT.FM_��� & "))) AS ������," & _
                    "                   DECODE(A.����, NULL, 0, A.����) AS ����, " & str���ۼ� & _
                    "                   C.�������,C.��Ʊ��,c.��Ʊ���� ,to_char(C.��Ʊ����,'yyyy-mm-dd') as ��Ʊ����,rtrim(ltrim(to_char(nvl(c.�������,0),'9999999999999999'))) as �������, ltrim(rtrim(to_char(C.��Ʊ���," & mOraFMT.FM_��� & "))) as ��Ʊ���, A.ID �շ�ID, a.��Ʒ����, a.�ڲ����� " & _
                    "           FROM  ҩƷ�շ���¼ A, �������� b,�շ���ĿĿ¼ D,�շ���Ŀ���� E, " & _
                    "                 (Select �շ�id,�������,��Ʊ��,��Ʊ����,��Ʊ����,�������,��Ʊ��� From Ӧ����¼ Where ϵͳ��ʶ=5 And ��¼����=0) C " & _
                    "           Where  A.ҩƷid = B.����id and A.ҩƷid=D.id AND A.Id = C.�շ�id (+) And D.ID = E.�շ�ϸĿid(+) And E.����(+) = 3 " & strTemp & _
                    "                   AND A.��¼״̬ =[3] " & _
                    "                   AND A.���� = [1] " & _
                    "                   AND A.No =[2]" & _
                    "       ) " & _
                    "   ORDER BY " & str����
                    int���� = 15
            Case 1713 '��������������
                IntBill = 2
                 gstrSQL = "" & _
                    "   select ���,������Ϣ,���,����,ʧЧ��,����,��λ,�ɹ���,�ɹ����,�ۼ�,�ۼ۽��,��� " & _
                    "   FROM (  SELECT DISTINCT ���,('[' || d.���� || ']' || d.����) AS ������Ϣ,d.���,d.����,zlSpellCode(d.����) ����,a.����, to_char(A.Ч��,'yyyy-mm-dd') as ʧЧ��," & _
                                        strUnitQuantity & _
                    "                   To_Char((a.�ɱ���*" & str��װϵ�� & ")," & mOraFMT.FM_�ɱ��� & ") AS �ɹ���," & _
                    "                   TO_CHAR (a.�ɱ����, " & mOraFMT.FM_��� & ") AS �ɹ����," & _
                    "                   TO_CHAR (a.���ۼ�*" & str��װϵ�� & ", " & mOraFMT.FM_���ۼ� & ") AS �ۼ�," & _
                    "                   TO_CHAR (a.���۽��," & mOraFMT.FM_��� & ") AS �ۼ۽��," & _
                    "                   TO_CHAR (a.���, " & mOraFMT.FM_��� & ") AS ��� " & _
                    "           FROM ҩƷ�շ���¼ a , �������� b,�շ���ĿĿ¼ D " & _
                    "           Where a.ҩƷid = b.����id and a.ҩƷid=d.id " & _
                    "                   AND a.��¼״̬ = [3] " & _
                    "                   AND a.���� = [1] AND ���ϵ��=1 " & _
                    "                   AND a.no = [2] " & _
                    "         )" & _
                    "   ORDER BY " & str����
                    
                    int���� = 16
            Case 1714       '�������
                IntBill = 4
                If mintUnit <> 0 Then
                    str���� = "���,������Ϣ,���,����,��׼�ĺ�,����,��������,ʧЧ��,����,��λ,�ɱ���,�ɱ����,�ۼ�,�ۼ۽��,���,���ۼ�,���۽��,���۲��,��Ʒ����,�ڲ�����"
                    str���ۼ� = "" & _
                        "                   ltrim(rtrim(to_char(((A.���۽��-to_number(nvl(to_char(A.�÷�," & gOraFmt_Max.FM_��� & "),'0')," & gOraFmt_Max.FM_��� & " ))/a.ʵ������)* " & str��װϵ�� & "," & mOraFMT.FM_���ۼ� & "))) as �ۼ� , " & _
                        "                   ltrim(rtrim(to_char(A.���۽��-to_number(nvl(to_char(A.�÷�," & gOraFmt_Max.FM_��� & "),'0')," & gOraFmt_Max.FM_��� & " )," & mOraFMT.FM_��� & ")))  as �ۼ۽��, " & _
                        "                   ltrim(rtrim(to_char(A.���-to_number(nvl(to_char(A.�÷�," & gOraFmt_Max.FM_��� & "),'0')," & gOraFmt_Max.FM_��� & " ), " & mOraFMT.FM_��� & "))) as ���," & _
                        "                   ltrim(rtrim(to_char(A.���ۼ�," & mOraFMT.FM_ɢװ���ۼ� & "))) as ���ۼ� , " & _
                        "                   D.���㵥λ as ���۵�λ , " & _
                        "                   ltrim(rtrim(to_char(A.���۽��," & mOraFMT.FM_��� & ")))  as ���۽��," & _
                        "                   ltrim(rtrim(to_char(A.���," & mOraFMT.FM_��� & "))) as ���۲�� "
                Else
                    str���� = "���,������Ϣ,���,����,��׼�ĺ�,����,��������,ʧЧ��,����,��λ,�ɱ���,�ɱ����,�ۼ�,�ۼ۽��,���,��Ʒ����,�ڲ�����"
                    str���ۼ� = "" & _
                    "                   ltrim(rtrim(to_char(A.���ۼ�*" & str��װϵ�� & "," & mOraFMT.FM_���ۼ� & "))) as �ۼ� , " & _
                    "                   ltrim(rtrim(to_char(A.���۽��," & mOraFMT.FM_��� & ")))  as �ۼ۽��, " & _
                    "                   ltrim(rtrim(to_char(A.���," & mOraFMT.FM_��� & "))) as ���"

                End If
                                
                                
                gstrSQL = "" & _
                    "   Select * " & _
                    "   From (  SELECT distinct ���, ('[' || D.���� || ']' || D.����) AS ������Ϣ," & _
                    "                   D.���,d.����,zlSpellCode(d.����) ����, A.����,A.��׼�ĺ�, A.����, to_char(A.��������,'yyyy-mm-dd') as ��������, to_char(A.Ч��,'yyyy-mm-dd') as ʧЧ��," & strUnitQuantity & _
                    "                   to_char(A.�ɱ���*" & str��װϵ�� & "," & mOraFMT.FM_�ɱ��� & ") AS �ɱ���, to_char(A.�ɱ����," & mOraFMT.FM_��� & ") AS �ɱ����," & str���ۼ� & _
                    "           , a.��Ʒ����, a.�ڲ����� " & _
                    "           FROM ҩƷ�շ���¼ A, �������� b,�շ���ĿĿ¼ D  " & _
                    "           Where  A.ҩƷid = B.����id and a.ҩƷid=d.id  " & _
                    "                   AND A.��¼״̬ =  [3] " & _
                    "                   AND A.���� = [1] " & _
                    "                   AND A.No =[2] " & _
                    "       ) " & _
                    "   ORDER BY " & str����
                int���� = 17
                
            Case 1715 '���Ŀ���۵���
                IntBill = 5
                
                gstrSQL = "" & _
                    "   Select ���,������Ϣ,���,����,����,ʧЧ��,��λ,�����,�����,������ " & _
                    "   From (  SELECT distinct ���, ('[' || D.���� || ']' || D.����) AS ������Ϣ," & _
                    "                   D.���,d.����,zlSpellCode(d.����) ����, A.����, A.����, to_char(A.Ч��,'yyyy-mm-dd') as ʧЧ��," & IIf(mintUnit = 0, "D.���㵥λ", "B.��װ��λ") & " as ��λ," & _
                    "                   to_char(A.���ۼ�," & mOraFMT.FM_��� & ") AS �����,to_char(A.�ɱ���," & mOraFMT.FM_��� & ") AS �����," & _
                    "                   to_char(A.���," & mOraFMT.FM_��� & ")  as ������ " & _
                    "           FROM ҩƷ�շ���¼ A, �������� b,�շ���ĿĿ¼ D" & _
                    "           Where  A.ҩƷid = B.����id and A.ҩƷid=d.id " & _
                    "                   AND A.��¼״̬ =  [3] " & _
                    "                   AND A.���� = [1] " & _
                    "                   AND A.No =[2] " & _
                    "       ) " & _
                    "   ORDER BY " & str����
                int���� = 18
                    
            Case 1716       '�����ƿ����
                IntBill = 6
                
                gstrSQL = "" & _
                    "   SELECT ���,������Ϣ,���,����,��׼�ĺ�,����,ʧЧ��,��д����,ʵ������,��λ,�ɱ���,�ɱ����,�ۼ�,�ۼ۽��,���,��Ʒ����,�ڲ����� " & _
                    "   FROM (  SELECT DISTINCT ���,('[' || D.���� || ']' || d.����) AS ������Ϣ,d.���,d.����,zlSpellCode(d.����) ����,a.����,a.��׼�ĺ�, a.����, to_char(A.Ч��,'yyyy-mm-dd') as ʧЧ��," & _
                    "                   (to_char(A.��д���� /" & str��װϵ�� & "," & mOraFMT.FM_���� & ")) AS ��д����,(to_char(A.ʵ������ /" & str��װϵ�� & "," & mOraFMT.FM_���� & ")) AS ʵ������," & IIf(mintUnit = 0, "D.���㵥λ", "B.��װ��λ") & " as ��λ," & _
                    "                   TO_CHAR (a.�ɱ���*" & str��װϵ�� & "," & mOraFMT.FM_�ɱ��� & ") AS �ɱ���," & _
                    "                   TO_CHAR (a.�ɱ����, " & mOraFMT.FM_��� & ") AS �ɱ����," & _
                    "                   TO_CHAR (a.���ۼ�*" & str��װϵ�� & ", " & mOraFMT.FM_���ۼ� & ") AS �ۼ�," & _
                    "                   TO_CHAR (a.���۽��, " & mOraFMT.FM_��� & ") AS �ۼ۽��," & _
                    "                   TO_CHAR (a.���, " & mOraFMT.FM_��� & ") AS ���, a.��Ʒ����, a.�ڲ����� " & _
                    "           FROM ҩƷ�շ���¼ a, �������� b,�շ���ĿĿ¼ D " & _
                    "           Where a.ҩƷid = b.����id and a.ҩƷid=d.id " & _
                    "                   AND a.��¼״̬ = [3] " & _
                    "                   AND a.���� = [1] AND ���ϵ��=-1 " & _
                    "                   AND a.no = [2] " & _
                    "           )" & _
                    "   ORDER BY " & str����
                    int���� = 19
                
            Case 1717       '����
                IntBill = 7
                
                gstrSQL = "" & _
                    "   SELECT ���,������Ϣ,���,����,��׼�ĺ�,����,ʧЧ��,��д����,ʵ������,��λ,�ɱ���,�ɱ����,�ۼ�,�ۼ۽��,���,��Ʒ����,�ڲ����� " & _
                    "   FROM (  SELECT DISTINCT ���,('[' || D.���� || ']' || D.����) AS ������Ϣ,D.���,d.����,zlSpellCode(d.����) ����,a.����,a.��׼�ĺ�, a.����, to_char(A.Ч��,'yyyy-mm-dd') as ʧЧ��," & _
                    "                   (to_char(A.��д���� /" & str��װϵ�� & "," & mOraFMT.FM_���� & ")) AS ��д����,(to_char(A.ʵ������ /" & str��װϵ�� & "," & mOraFMT.FM_���� & ")) AS ʵ������," & IIf(mintUnit = 0, "D.���㵥λ", "b.��װ��λ") & " as ��λ," & _
                    "                   TO_CHAR (A.�ɱ���*" & str��װϵ�� & ", " & mOraFMT.FM_�ɱ��� & ") AS �ɱ���," & _
                    "                   TO_CHAR (a.�ɱ����, " & mOraFMT.FM_��� & ") AS �ɱ����," & _
                    "                   TO_CHAR (a.���ۼ�*" & str��װϵ�� & "," & mOraFMT.FM_���ۼ� & ") AS �ۼ�," & _
                    "                   TO_CHAR (a.���۽��, " & mOraFMT.FM_��� & ") AS �ۼ۽��," & _
                    "                   TO_CHAR (a.���, " & mOraFMT.FM_��� & ") AS ���, a.��Ʒ����, a.�ڲ����� " & _
                    "           FROM ҩƷ�շ���¼ a , �������� b,�շ���ĿĿ¼ D" & _
                    "           Where a.ҩƷid = b.����id and a.ҩƷid=d.id " & _
                    "                   AND A.��¼״̬ = [3] " & _
                    "                   AND a.���� =[1] " & _
                    "                   AND a.no = [2] )" & _
                    "   ORDER BY " & str����
                int���� = 20
                    
            Case 1718   '��������
                IntBill = 11
                If mshList.TextMatrix(mshList.Row, 1) = "��������" Then
                    str���� = "���,������Ϣ,���,����,��׼�ĺ�,����,ʧЧ��,����,��λ,�ɱ���,�ɱ����,�ۼ�,�ۼ۽��,���,������,�������,��ֵ˰��,˰��,��Ʒ����,�ڲ�����"
                Else
                    str���� = "���,������Ϣ,���,����,��׼�ĺ�,����,ʧЧ��,����,��λ,�ɱ���,�ɱ����,�ۼ�,�ۼ۽��,���,��Ʒ����,�ڲ�����"
                End If
                
                gstrSQL = "" & _
                    "   Select " & str���� & _
                    "   From (  SELECT distinct ���, ('[' || d.���� || ']' ||d.����) AS ������Ϣ," & _
                    "                   d.���,d.����,zlSpellCode(d.����) ����, A.����,A.��׼�ĺ�, A.����, to_char(A.Ч��,'yyyy-mm-dd') as ʧЧ��," & strUnitQuantity & _
                    "                   to_char(a.�ɱ���*" & str��װϵ�� & "," & mOraFMT.FM_�ɱ��� & ") AS �ɱ���, to_char(A.�ɱ����," & mOraFMT.FM_��� & ") AS �ɱ����," & _
                    "                   to_char(A.���ۼ�*" & str��װϵ�� & "," & mOraFMT.FM_���ۼ� & ") as �ۼ� , to_char(A.���۽��," & mOraFMT.FM_��� & ")  as �ۼ۽��, to_char(A.���," & mOraFMT.FM_��� & ") as ��� "
                    
                If mshList.TextMatrix(mshList.Row, 1) = "��������" Then
                    gstrSQL = gstrSQL & " ,LTRIM(TO_CHAR(A.����*" & str��װϵ�� & "," & mOraFMT.FM_���ۼ� & ")) AS ������,LTRIM(TO_CHAR(A.����*A.ʵ������," & mOraFMT.FM_��� & ")) AS �������,LTRIM(TO_CHAR(Nvl(A.Ƶ��,0)/100," & mOraFMT.FM_��� & ")) As ��ֵ˰��,LTRIM(TO_CHAR(A.����*A.ʵ������*(Nvl(A.Ƶ��,0)/100/(1+Nvl(A.Ƶ��,0)/100))," & mOraFMT.FM_��� & ")) As ˰�� "
                End If
                    
                gstrSQL = gstrSQL & ", a.��Ʒ����, a.�ڲ�����  FROM ҩƷ�շ���¼ A , �������� b,�շ���ĿĿ¼ D" & _
                    "           Where  A.ҩƷid = B.����id and a.ҩƷid=d.id " & _
                    "                   AND A.��¼״̬ =  [3] " & _
                    "                   AND A.���� = [1] " & _
                    "                   AND A.No =[2] " & _
                    "       ) " & _
                    "   ORDER BY " & str����
                int���� = 21
            Case 1719 '�����̵����
                IntBill = 12
                
                gstrSQL = "" & _
                    "   SELECT * " & _
                    "   FROM (  SELECT DISTINCT ���,('[' || d.���� || ']' || d.����) AS ������Ϣ," & _
                    "                   d.���,d.����,zlSpellCode(d.����) ����,a.����," & IIf(strUnit = "��װ��λ", "d.��װ��λ", "b." & strUnit) & " as ��λ,a.����, to_char(A.Ч��,'yyyy-mm-dd') as ʧЧ��," & _
                    "                   (to_char(A.��д���� /" & str��װϵ�� & "," & mOraFMT.FM_���� & ")) AS ������," & _
                    "                   (to_char(A.���� /" & str��װϵ�� & "," & mOraFMT.FM_���� & ")) AS ʵ����," & _
                    "                   Decode(Sign(A.����-A.��д����),-1,'��',1,'ӯ','ƽ') as ��־," & _
                    "                   (to_char(A.ʵ������ /" & str��װϵ�� & "," & mOraFMT.FM_���� & ")) AS ������," & _
                    "                   TO_CHAR (a.���ۼ�*" & str��װϵ�� & ", " & mOraFMT.FM_���ۼ� & ") AS �ۼ�," & _
                    "                   TO_CHAR (a.���۽��, " & mOraFMT.FM_��� & ") AS ����," & _
                    "                   TO_CHAR (a.���, " & mOraFMT.FM_��� & ") AS ��۲�, " & _
                    "                   TO_CHAR ((A.���� / " & str��װϵ�� & ")*(a.���ۼ�*" & str��װϵ�� & "), " & mOraFMT.FM_��� & ") as �̵��� " & _
                    "           FROM ҩƷ�շ���¼ a, �������� b,�շ���ĿĿ¼ D" & _
                    "           Where a.ҩƷid = b.����id and a.ҩƷid=d.id  " & _
                    "                   AND ��¼״̬ = [3] " & _
                    "                   AND a.���� = [1] " & _
                    "                   AND a.no = [2] " & _
                    "       )" & _
                    "   ORDER BY " & str����
                int���� = 22
            End Select
            
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, int����, mshList.TextMatrix(mshList.Row, 0), Val(mshList.TextMatrix(mshList.Row, GetCol(mshList, "��¼״̬"))))
        
        Set mshDetail.DataSource = rsTemp
        With mshDetail
            If rsTemp.RecordCount = 0 Then
                .Rows = 2
                .Clear 1
            End If
            rsTemp.Close
        End With
        If mlngMode = 1712 Then
            mshDetail.ColHidden(mshDetail.ColIndex("�շ�ID")) = True
            Call mshDetail_EnterCell
        End If

    SetDetailColWidth
    SetEnable
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

Private Sub mshlist_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    If mnuEdit.Visible = False Then Exit Sub
    '���Popupmenuģ̬���壬���ܼ���Popupmenu
    mblnPopupmenuCall = True
    PopupMenu mnuEdit, 2
    mblnPopupmenuCall = False
    If mnuEditAdd.Tag = "1" Then
        Call mnuEditAdd_Click
    ElseIf mnuEditRestore.Tag = "1" Then
        Call mnuEditRestore_Click
    End If
    
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
    
    With mshList
        .Top = IIf(cbrTool.Visible, cbrTool.Height, 0) + IIf(TabShow.Visible, TabShow.Height, 0)
        .Height = picSeparate_s.Top - .Top
    End With
    
    With Cmd����
        .Top = mshList.Top + mshList.Height + 30
    End With
    
    With mshDetail
        .Top = picSeparate_s.Top + picSeparate_s.Height + 100
        If mlngMode = 1712 Then
            .Height = ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0) _
                    - IIf(vsfCostlyInfo.Visible, vsfCostlyInfo.Height, 0) _
                    - IIf(lblCostly.Visible, lblCostly.Height, 0)
        Else
            .Height = ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
        End If
    End With
    
End Sub

Private Sub picSeparate_s_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    mintOldY = 0
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
        Case "Prepare"
            mnuEditPrepare_Click
        Case "Send"
            mnuEditSend_Click
        Case "Back"
            mnuEditBack_Click
        Case "Verify"
            mnuEditVerify_Click
        Case "Check"
            mnuEditCheck_Click
        Case "CancelCheck"
            mnuEditCancelCheck_Click
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
    Dim strVerify As String
    Dim bln�˲� As Boolean
    Dim intCol As Integer
    Dim intTemp As Integer
    
    If mlngMode = 1712 Then
        If mbln��Ҫ�˲� Then
            mnuEditCheckBatch.Enabled = InStr(mstrPrivs, ";�˲�;") > 0
        End If
        mnuEditVerifyBatch.Enabled = InStr(mstrPrivs, ";���;") > 0
    End If
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
            tlbTool.Buttons("Strike").Enabled = False
        
            mnuEditCheck.Enabled = False
            mnuEditCancelCheck.Enabled = False
            tlbTool.Buttons("Check").Enabled = False
            tlbTool.Buttons("CancelCheck").Enabled = False
        
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
            
            If mnuEditBill.Visible = True Then
                mnuEditBill.Enabled = False
            End If
            
            If mnuEditReg.Visible = True Then
                mnuEditReg.Enabled = False
            End If
            
            If mnuEditAcc.Visible Then
                mnuEditAcc.Enabled = False
            End If
            
            If mnuEditImport.Visible Then
                mnuEditImport.Enabled = True
            End If
            
            If mnuEditPrepare.Visible Then
                mnuEditPrepare.Enabled = False
                mnuEditSend.Enabled = False
                mnuEditBack.Enabled = False
                tlbTool.Buttons("Prepare").Enabled = False
                tlbTool.Buttons("Send").Enabled = False
                tlbTool.Buttons("Back").Enabled = False
            End If
        Else
            mnuFilePreView.Enabled = True
            mnuFilePrint.Enabled = True
            mnuFileBillPreview.Enabled = True
            mnuFileBillPrint.Enabled = True
            mnuFileExcel.Enabled = True
            tlbTool.Buttons("Print").Enabled = True
            tlbTool.Buttons("PrintView").Enabled = True
            
            'ֻ���⹺��ⵥ����
            If mnuEditBill.Visible = True Then
                mnuEditBill.Enabled = False
            End If
            
            If mnuEditReg.Visible = True Then
                mnuEditReg.Enabled = False
            End If
            
            If mnuEditAcc.Visible Then
                mnuEditAcc.Enabled = False
            End If
            
            If mnuEditImport.Visible Then
                mnuEditImport.Enabled = True
            End If
            
            If mlngMode = 1719 Then
                strVerify = .TextMatrix(.Row, .Cols - 6)
            Else
                If mlngMode = 1716 Then '�����ƿ�
                    If mint�ƿ⴦������ = 1 Then
                        If TabShow.Tab = 0 Then
                            If Val(.TextMatrix(.Row, GetCol(mshList, "��¼״̬"))) Mod 3 = 2 Then
                                strVerify = .TextMatrix(.Row, GetCol(mshList, "������")) '���������Ƿ����
                            Else
                                strVerify = .TextMatrix(.Row, GetCol(mshList, "������")) '�Ƿ���
                            End If
                        Else
                            strVerify = .TextMatrix(.Row, GetCol(mshList, "������")) '�Ƿ����
                        End If
                    Else
                        strVerify = .TextMatrix(.Row, GetCol(mshList, "������")) '�Ƿ����
                    End If
                Else
                    strVerify = .TextMatrix(.Row, GetCol(mshList, "�������"))    '�������
                End If
            End If
            
            If strVerify = "" Then    'δ��˵�
                If mlngMode = 1712 Then
                
                    '���˺�:����˲�����2007/05/13
                    intCol = GetCol(mshList, "�˲�����")
                    If intCol >= 0 Then
                        bln�˲� = Trim(.TextMatrix(.Row, intCol)) <> ""
                    Else
                        bln�˲� = False
                    End If
                    
                    If mnuEditModify.Visible = True Then
                        mnuEditModify.Enabled = Not bln�˲�
                        tlbTool.Buttons("Modify").Enabled = Not bln�˲�
                    End If
                    If mnuEditDel.Visible = True Then
                        mnuEditDel.Enabled = Not bln�˲�
                        tlbTool.Buttons("Delete").Enabled = Not bln�˲�
                    End If
                    
                    mnuEditCheck.Enabled = Not bln�˲�
                    mnuEditCancelCheck.Enabled = bln�˲�
                    tlbTool.Buttons("Check").Enabled = Not bln�˲�
                    tlbTool.Buttons("CancelCheck").Enabled = bln�˲�
                    
                    If mnuEditVerify.Visible = True Then
                        mnuEditVerify.Enabled = IIf(mbln��Ҫ�˲�, bln�˲�, True)
                        tlbTool.Buttons("Verify").Enabled = IIf(mbln��Ҫ�˲�, bln�˲�, True)
                    End If
                    
                ElseIf mlngMode = 1717 Then
                    intCol = GetCol(mshList, "�˲�����")
                    If intCol >= 0 Then
                        bln�˲� = Trim(.TextMatrix(.Row, intCol)) <> ""
                    Else
                        bln�˲� = False
                    End If
                    
                    If mnuEditModify.Visible = True Then
                        mnuEditModify.Enabled = Not bln�˲�
                        tlbTool.Buttons("Modify").Enabled = Not bln�˲�
                    End If
                    If mnuEditDel.Visible = True Then
                        mnuEditDel.Enabled = Not bln�˲�
                        tlbTool.Buttons("Delete").Enabled = Not bln�˲�
                    End If
                    
                    mnuEditCheck.Enabled = Not bln�˲�
                    mnuEditCancelCheck.Enabled = bln�˲�
                    tlbTool.Buttons("Check").Enabled = Not bln�˲�
                    tlbTool.Buttons("CancelCheck").Enabled = bln�˲�
                    
                    If mnuEditVerify.Visible = True Then
                        mnuEditVerify.Enabled = IIf(mint������˷�ʽ = 1, bln�˲�, True)
                        tlbTool.Buttons("Verify").Enabled = mnuEditVerify.Enabled
                    End If
                Else
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
                End If
                If mnuEditStrike.Visible = True Then
                    mnuEditStrike.Enabled = False
                    tlbTool.Buttons("Strike").Enabled = False
                End If
                 
                If mnuEditDisplay.Visible = True Then
                    mnuEditDisplay.Enabled = True
                End If
                
                '�ƿⵥ�����ݵ�ǰѡ���ҳ�棬��ǰ�������ð�ť״̬
                If mlngMode = 1716 Then
                    If TabShow.Tab = 0 Then
                        If Val(.TextMatrix(.Row, GetCol(mshList, "��¼״̬"))) Mod 3 = 2 Then
                            tlbTool.Buttons("Modify").Enabled = False
                            tlbTool.Buttons("Delete").Enabled = False
                            tlbTool.Buttons("Prepare").Enabled = False
                            tlbTool.Buttons("Send").Enabled = False
                            tlbTool.Buttons("Back").Enabled = False
                            tlbTool.Buttons("Verify").Enabled = False
                            tlbTool.Buttons("Strike").Enabled = True
                            mnuEditModify.Enabled = False
                            mnuEditDel.Enabled = False
                            mnuEditPrepare.Enabled = False
                            mnuEditVerify.Enabled = False
                            mnuEditStrike.Enabled = True
                        Else
                            mnuEditPrepare.Enabled = (.TextMatrix(.Row, 0) <> "")
                            mnuEditSend.Enabled = False
                            mnuEditBack.Enabled = False
                            tlbTool.Buttons("Prepare").Enabled = mnuEditPrepare.Enabled
                            tlbTool.Buttons("Send").Enabled = False
                            tlbTool.Buttons("Back").Enabled = False
                            
                            '����õ�������ˣ�������ҩ�뷢��
                            If TestVerify(mshList.TextMatrix(mshList.Row, 0)) Then
                                mnuEditPrepare.Enabled = False
                                mnuEditSend.Enabled = False
                                mnuEditBack.Enabled = False
                                tlbTool.Buttons("Prepare").Enabled = False
                                tlbTool.Buttons("Send").Enabled = False
                                tlbTool.Buttons("Back").Enabled = False
                                tlbTool.Buttons("Strike").Enabled = False
                            Else
                                tlbTool.Buttons("Strike").Enabled = False
                            End If
                        End If
                    Else
                        If Val(.TextMatrix(.Row, GetCol(mshList, "��¼״̬"))) Mod 3 = 2 Then
                            tlbTool.Buttons("Modify").Enabled = False
                            tlbTool.Buttons("Strike").Enabled = False
                            tlbTool.Buttons("Verify").Enabled = False
                            tlbTool.Buttons("Delete").Enabled = True
                            mnuEditModify.Enabled = False
                            mnuEditStrike.Enabled = False
                            mnuEditVerify.Enabled = False
                            mnuEditDel.Enabled = True
                        Else
                            mnuEditVerify.Enabled = TestPrepare(.TextMatrix(.Row, 0))
                            tlbTool.Buttons("Verify").Enabled = mnuEditVerify.Enabled
                        End If
                    End If
                End If
                
            ElseIf .TextMatrix(.Row, GetCol(mshList, "��¼״̬")) = 1 Then  '��˵�
                If mlngMode = 1712 Or mlngMode = 1717 Then
                    '���˺�:����˲鹦��2007/05/13
                    mnuEditCheck.Enabled = False
                    tlbTool.Buttons("Check").Enabled = False
                    mnuEditCheck.Enabled = False
                    mnuEditCancelCheck.Enabled = False
                    tlbTool.Buttons("Check").Enabled = False
                    tlbTool.Buttons("CancelCheck").Enabled = False
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
                    If mlngMode = 1715 And .TextMatrix(.Row, .Cols - 1) = "1" Then
                        mnuEditStrike.Enabled = False
                        tlbTool.Buttons("Strike").Enabled = False
                    Else
                        mnuEditStrike.Enabled = True
                        tlbTool.Buttons("Strike").Enabled = True
                    End If
                End If
                 
                If mnuEditDisplay.Visible = True Then
                    mnuEditDisplay.Enabled = True
                End If
                
                'ֻ���⹺��ⵥ����
                If mnuEditBill.Visible = True Then
'                    If Val(mshDetail.TextMatrix(mshDetail.Row, mshDetail.Cols - 2)) = 0 Then
                        mnuEditBill.Enabled = True
'                    End If
                End If
                
                If mnuEditReg.Visible = True Then
                    If Val(mshDetail.TextMatrix(mshDetail.Row, mshDetail.Cols - 2)) = 0 Then
                        mnuEditReg.Enabled = True
                    End If
                End If
                
                If mlngMode = 1716 And TabShow.Tab = 0 Then
                    mnuEditAcc.Enabled = (Trim(mshList.TextMatrix(mshList.Row, GetCol(mshList, "��������"))) <> "")
                Else
                    mnuEditAcc.Enabled = (Trim(mshList.TextMatrix(mshList.Row, GetCol(mshList, "�������"))) <> "")
                End If
                
                If mlngMode = 1716 Then
                    If TabShow.Tab = 0 Then
                        mnuEditPrepare.Enabled = False
                        mnuEditSend.Enabled = (mshList.TextMatrix(mshList.Row, mshList.Cols - 3) = "")
                        mnuEditBack.Enabled = True
                        tlbTool.Buttons("Prepare").Enabled = False
                        tlbTool.Buttons("Send").Enabled = mnuEditSend.Enabled
                        tlbTool.Buttons("Back").Enabled = True
                        '����õ�������ˣ����������뷢��
                        If TestVerify(mshList.TextMatrix(mshList.Row, 0)) Then
                            mnuEditPrepare.Enabled = False
                            mnuEditSend.Enabled = False
                            mnuEditBack.Enabled = False
                            tlbTool.Buttons("Prepare").Enabled = False
                            tlbTool.Buttons("Send").Enabled = False
                            tlbTool.Buttons("Back").Enabled = False
                        End If
                        
                        If mnuEditStrike.Visible = True Then
                            If mint�������� = 1 Then
                                mnuEditStrike.Enabled = False
                                tlbTool.Buttons("Strike").Enabled = False
                            Else
                                mnuEditStrike.Enabled = True
                                tlbTool.Buttons("Strike").Enabled = True
                            End If
                        End If
                    Else
                        If mnuEditStrike.Visible = True Then
                            mnuEditStrike.Enabled = True
                            tlbTool.Buttons("Strike").Enabled = True
                        End If
                    End If
                End If
                
            Else   '2,3 ���������Ѹ���ĵ��ݲ����������ˣ�ͬ����������˺�ĵ��ݲ����������
                If mnuEditBill.Visible = True Then
                    mnuEditBill.Enabled = True
                End If
                
                If mnuEditReg.Visible = True Then
                    mnuEditReg.Enabled = True
                End If
                
                If mlngMode = 1712 Or mlngMode = 1717 Then
                    '���˺�:����˲鹦��2007/05/13
                    mnuEditCheck.Enabled = False
                    mnuEditCancelCheck.Enabled = False
                    tlbTool.Buttons("Check").Enabled = False
                    tlbTool.Buttons("CancelCheck").Enabled = False
                    
                End If
                
                If Val(.TextMatrix(.Row, GetCol(mshList, "��¼״̬"))) Mod 3 = 0 Then
                    '�������
                    intTemp = GetCol(mshList, "�����־")
                    If intTemp >= 0 Then intTemp = Val(.TextMatrix(.Row, intTemp))
                    .ToolTipText = IIf(intTemp = 1, "������˵�ԭ����", "�������ݵ�ԭ����")
                    If mnuEditStrike.Visible = True Then
                        mnuEditStrike.Enabled = True
                        tlbTool.Buttons("Strike").Enabled = True
                    Else
                        tlbTool.Buttons("Strike").Enabled = False
                    End If
                    If mlngMode = 1716 And TabShow.Tab = 0 Then
                        mnuEditAcc.Enabled = (Trim(mshList.TextMatrix(mshList.Row, GetCol(mshList, "��������"))) = "")
                    Else
                        mnuEditAcc.Enabled = (Trim(mshList.TextMatrix(mshList.Row, GetCol(mshList, "�������"))) = "")
                    End If
                    
                    If mlngMode = 1716 Then
                        If TabShow.Tab = 0 Then
                            mnuEditPrepare.Enabled = False
                            mnuEditSend.Enabled = (mshList.TextMatrix(mshList.Row, mshList.Cols - 3) = "")
                            mnuEditBack.Enabled = True
                            tlbTool.Buttons("Prepare").Enabled = False
                            tlbTool.Buttons("Send").Enabled = mnuEditSend.Enabled
                            tlbTool.Buttons("Back").Enabled = True
                            '����õ�������ˣ����������뷢��
                            If TestVerify(mshList.TextMatrix(mshList.Row, 0)) Then
                                mnuEditPrepare.Enabled = False
                                mnuEditSend.Enabled = False
                                mnuEditBack.Enabled = False
                                tlbTool.Buttons("Prepare").Enabled = False
                                tlbTool.Buttons("Send").Enabled = False
                                tlbTool.Buttons("Back").Enabled = False
                            End If
                            
                            If mnuEditStrike.Visible = True Then
                                If mint�������� = 1 Then
                                    mnuEditStrike.Enabled = False
                                    tlbTool.Buttons("Strike").Enabled = False
                                Else
                                    mnuEditStrike.Enabled = True
                                    tlbTool.Buttons("Strike").Enabled = True
                                End If
                            End If
                        Else
                            If mnuEditStrike.Visible = True Then
                                mnuEditStrike.Enabled = True
                                tlbTool.Buttons("Strike").Enabled = True
                            End If
                        End If
                    End If
                ElseIf .TextMatrix(.Row, GetCol(mshList, "��¼״̬")) Mod 3 = 2 Then
                    .ToolTipText = "��������"
                    '�������
                    intTemp = GetCol(mshList, "�����־")
                    If intTemp >= 0 Then intTemp = Val(.TextMatrix(.Row, intTemp))
                    .ToolTipText = IIf(intTemp = 1, "������˵ĳ�������", "��������")
                    
                    If mnuEditStrike.Visible = True Then
                        mnuEditStrike.Enabled = False
                        tlbTool.Buttons("Strike").Enabled = False
                    End If
                    
                    If mlngMode = 1716 Then '�ƿ�
                        If TabShow.Tab = 0 Then
                            If mnuEditVerify.Visible = True Then
                                mnuEditVerify.Enabled = False
                                tlbTool.Buttons("Verify").Enabled = False
                            End If
                            
                            If mint�������� = 1 Then mnuEditStrike.Visible = True
                            If strVerify = "" Then
                                mnuEditStrike.Enabled = True
                                tlbTool.Buttons("Strike").Enabled = True
                            Else
                                mnuEditStrike.Enabled = False
                                tlbTool.Buttons("Strike").Enabled = False
                            End If
                        Else

                        End If
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
                
                If mlngMode = 1304 Or mlngMode = 1716 Then
                    If TabShow.Tab = 0 Then
                        mnuEditPrepare.Enabled = False
                        mnuEditSend.Enabled = False
                        mnuEditBack.Enabled = False
                        tlbTool.Buttons("Prepare").Enabled = False
                        tlbTool.Buttons("Send").Enabled = mnuEditSend.Enabled
                        tlbTool.Buttons("Back").Enabled = False
                        '����õ�������ˣ�������ҩ�뷢��
                        If TestVerify(mshList.TextMatrix(mshList.Row, 0)) Then
                            mnuEditPrepare.Enabled = False
                            mnuEditSend.Enabled = False
                            mnuEditBack.Enabled = False
                            tlbTool.Buttons("Prepare").Enabled = False
                            tlbTool.Buttons("Send").Enabled = False
                            tlbTool.Buttons("Back").Enabled = False
                        End If
                    End If
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
    objPrint.Title.Font.Name = "����_GB2312"
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

Private Sub subExcel(bytMode As Byte)
'����:���������EXCEL
'����:bytMode3 �����EXCEL

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
    objRow.Add "NO." & Trim(mshList.TextMatrix(mshList.Row, GetCol(mshList, "NO")))
    objPrint.UnderAppRows.Add objRow
    
    Select Case mlngMode
        Case 1712       '�����⹺������
            Set objRow = New zlTabAppRow
            objRow.Add "�ⷿ��" & Trim(cboStock.Text)
            objRow.Add "��Ӧ�̣�" & Trim(mshList.TextMatrix(mshList.Row, GetCol(mshList, "��Ӧ��")))
            objPrint.UnderAppRows.Add objRow
                
        Case 1713       '��������������
            Set objRow = New zlTabAppRow
            objRow.Add "�ⷿ��" & Trim(cboStock.Text)
            objRow.Add "�Ƽ��ң�" & Trim(mshList.TextMatrix(mshList.Row, GetCol(mshList, "�Ƽ���")))
            objPrint.UnderAppRows.Add objRow
            
        Case 1714      '��������������
            Set objRow = New zlTabAppRow
            objRow.Add "�ⷿ��" & Trim(cboStock.Text)
            objRow.Add "������" & Trim(mshList.TextMatrix(mshList.Row, GetCol(mshList, "������")))
            objPrint.UnderAppRows.Add objRow
        Case 1715       '����۵�������
            Set objRow = New zlTabAppRow
            objRow.Add "�ⷿ��" & Trim(cboStock.Text)
            objPrint.UnderAppRows.Add objRow
            
        Case 1716       '�����ƿ����
            Set objRow = New zlTabAppRow
            If TabShow.Tab = 0 Then
                objRow.Add "�Ƴ��ⷿ��" & Trim(cboStock.Text)
                objRow.Add "����ⷿ��" & Trim(mshList.TextMatrix(mshList.Row, GetCol(mshList, "����ⷿ")))
            Else
                objRow.Add "����ⷿ��" & Trim(cboStock.Text)
                objRow.Add "�Ƴ��ⷿ��" & Trim(mshList.TextMatrix(mshList.Row, GetCol(mshList, "�Ƴ��ⷿ")))
            End If
            objPrint.UnderAppRows.Add objRow
        Case 1717       '�������ù���
            Set objRow = New zlTabAppRow
            objRow.Add "�����Ŀⷿ��" & Trim(cboStock.Text)
            objRow.Add "���ò��ţ�" & Trim(mshList.TextMatrix(mshList.Row, GetCol(mshList, "���ò���")))
            objPrint.UnderAppRows.Add objRow
            
        Case 1718       '���������������
            Set objRow = New zlTabAppRow
            objRow.Add "�ⷿ��" & Trim(cboStock.Text)
            objRow.Add "������" & Trim(mshList.TextMatrix(mshList.Row, GetCol(mshList, "������")))
            objPrint.UnderAppRows.Add objRow
        Case 1719       '�����̵����
            Set objRow = New zlTabAppRow
            objRow.Add "�̵�ⷿ��" & Trim(cboStock.Text)
            objRow.Add "�̵�ʱ�䣺" & Trim(mshList.TextMatrix(mshList.Row, GetCol(mshList, "�̵�ʱ��")))
            objPrint.UnderAppRows.Add objRow
    End Select
        
    Set objRow = New zlTabAppRow
    objRow.Add "ժҪ:" & mshList.TextMatrix(mshList.Row, GetCol(mshList, "ժҪ"))
    objPrint.BelowAppRows.Add objRow
        
    Set objRow = New zlTabAppRow
    objRow.Add "������:" & mshList.TextMatrix(mshList.Row, GetCol(mshList, "������")) & "  ��������:" & mshList.TextMatrix(mshList.Row, GetCol(mshList, "��������"))
    
    '���������ƿ�ģ��
    If mlngMode = 1716 Then
        If TabShow.Tab = 0 Then
            objRow.Add "�����:  �������:"
        Else
            objRow.Add "�����:" & mshList.TextMatrix(mshList.Row, GetCol(mshList, "������")) & "  �������:" & mshList.TextMatrix(mshList.Row, GetCol(mshList, "�������"))
        End If
    Else
        objRow.Add "�����:" & mshList.TextMatrix(mshList.Row, GetCol(mshList, "�����")) & "  �������:" & mshList.TextMatrix(mshList.Row, GetCol(mshList, "�������"))
    End If
    objPrint.BelowAppRows.Add objRow
    
    Set objPrint.Body = mshDetail
    zlPrintOrView1Grd objPrint, bytMode
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
    
    err = 0: On Error Resume Next
    With mshList
        If .Rows > 1 Then
            .Redraw = False
            intCol = .MouseCol
            .Col = intCol
            .ColSel = intCol
            intTemp = .TextMatrix(.Row, intCol)

            Select Case mlngMode
                Case 1712, 1718
                    If InStr(1, "345", intCol) <> 0 Then '345Ϊ����,����Ϊ�ַ�
                        If intCol = mintPreCol And mintsort = flexSortNumericDescending Then
                           .Sort = flexSortNumericAscending
                           mintsort = flexSortNumericAscending
                        Else
                           .Sort = flexSortNumericDescending
                           mintsort = flexSortNumericDescending
                        End If
                    Else
                        If intCol = mintPreCol And mintsort = flexSortStringNoCaseDescending Then
                           .Sort = flexSortStringNoCaseAscending
                           mintsort = flexSortStringNoCaseAscending
                        Else
                           .Sort = flexSortStringNoCaseDescending
                           mintsort = flexSortStringNoCaseDescending
                        End If
                    End If
                Case 1713, 1714   '2,34��Ϊ���֣�����Ϊ�ַ�
                    If InStr(1, "234", intCol) <> 0 Then
                        If intCol = mintPreCol And mintsort = flexSortNumericDescending Then
                           .Sort = flexSortNumericAscending
                           mintsort = flexSortNumericAscending
                        Else
                           .Sort = flexSortNumericDescending
                           mintsort = flexSortNumericDescending
                        End If
                    Else
                        If intCol = mintPreCol And mintsort = flexSortStringNoCaseDescending Then
                           .Sort = flexSortStringNoCaseAscending
                           mintsort = flexSortStringNoCaseAscending
                        Else
                           .Sort = flexSortStringNoCaseDescending
                           mintsort = flexSortStringNoCaseDescending
                        End If
                    End If
                Case 1715               '1,2,3��Ϊ���֣�����Ϊ�ַ�
                    If InStr(1, "123", intCol) <> 0 Then
                        If intCol = mintPreCol And mintsort = flexSortNumericDescending Then
                           .Sort = flexSortNumericAscending
                           mintsort = flexSortNumericAscending
                        Else
                           .Sort = flexSortNumericDescending
                           mintsort = flexSortNumericDescending
                        End If
                    Else
                        If intCol = mintPreCol And mintsort = flexSortStringNoCaseDescending Then
                           .Sort = flexSortStringNoCaseAscending
                           mintsort = flexSortStringNoCaseAscending
                        Else
                           .Sort = flexSortStringNoCaseDescending
                           mintsort = flexSortStringNoCaseDescending
                        End If
                    End If
                Case 1716, 1717 '2,3,4��Ϊ���֣�����Ϊ�ַ�
                    If InStr(1, "234", intCol) <> 0 Then
                        If intCol = mintPreCol And mintsort = flexSortNumericDescending Then
                           .Sort = flexSortNumericAscending
                           mintsort = flexSortNumericAscending
                        Else
                           .Sort = flexSortNumericDescending
                           mintsort = flexSortNumericDescending
                        End If
                    Else
                        If intCol = mintPreCol And mintsort = flexSortStringNoCaseDescending Then
                           .Sort = flexSortStringNoCaseAscending
                           mintsort = flexSortStringNoCaseAscending
                        Else
                           .Sort = flexSortStringNoCaseDescending
                           mintsort = flexSortStringNoCaseDescending
                        End If
                    End If
                Case 1719               'ȫΪ�ַ�
                    If intCol = mintPreCol And mintsort = flexSortStringNoCaseDescending Then
                        .Sort = flexSortStringNoCaseAscending
                        mintsort = flexSortStringNoCaseAscending
                    Else
                       .Sort = flexSortStringNoCaseDescending
                       mintsort = flexSortStringNoCaseDescending
                    End If
                Case Else

            End Select
            mintPreCol = intCol
            .Row = grid.MshGrdFindRow(mshList, intTemp, intCol)
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
            
            Select Case mlngMode
                Case 1712                   '6,8,9,10,11,12,13,16Ϊ���֣�����Ϊ�ַ�
                    Select Case intCol
                        Case 6, 8, 9, 10, 11, 12, 13, 16
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
                        
                
                Case 1713, 1714, 1718       '6,8,9,10,11,12Ϊ���֣�����Ϊ�ַ�
                    Select Case intCol
                        Case 6, 8, 9, 10, 11, 12
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
                Case 1715                   '7,8Ϊ���֣�����Ϊ�ַ�
                    Select Case intCol
                        Case 7, 8
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
                Case 1716, 1717             '6,7,9,10,11,12,13Ϊ���֣�����Ϊ�ַ�
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
                Case 1719                   '7,8,10,11,12,13Ϊ���֣�����Ϊ�ַ�
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

Private Sub PrintRange(ByVal strRange As String)
    '����:��ӡʱ�䷶Χ
    picSeparate_s.Cls
    picSeparate_s.CurrentX = 50
    picSeparate_s.CurrentY = 50
    picSeparate_s.Print strRange
End Sub
Private Function TestVerify(ByVal strNo As String) As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    '���õ����Ƿ�ͨ����ˣ�������ƿⵥ
    gstrSQL = "" & _
        "   Select ����� From ҩƷ�շ���¼ " & _
        "   Where ����=19 And NO=[1] And Rownum<2"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����Ƿ�ͨ�����", strNo)
        
    If Not IsNull(rsTemp!�����) Then
        TestVerify = True
    End If
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function TestPrepare(ByVal strNo As String) As Boolean
    Dim IntBill As Integer
    Dim rsTemp As New ADODB.Recordset
    '�����ҩ���Ƿ��Ѿ���д
    On Error GoTo ErrHandle
    Select Case mlngMode
    Case 1712
        IntBill = 15
    Case 1716
        IntBill = 19
    Case Else
        Exit Function
    End Select
    
    gstrSQL = "Select ��ҩ�� From ҩƷ�շ���¼ Where ����=[1] And NO=[2]  And Rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����Ƿ�ͨ���˲�", IntBill, strNo)
    If Not IsNull(rsTemp!��ҩ��) Then
        TestPrepare = True
    End If
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub tabShow_Click(PreviousTab As Integer)
    If mlngMode <> 1716 Then Exit Sub
    Call SetMenu

    Call GetList(mstrFind)
End Sub


Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

Private Sub LoadPlugInMnu(ByVal blnHave As Boolean)
'������blnHave true ��ʾ����������
    Dim strTmp As String
    Dim arrTmp As Variant
    Dim i As Integer
    
    mnuPlugIn.Visible = blnHave
 
    If blnHave Then
        'blnHave Ϊtrue ʱ����ȷ�� gobjPlugIn ����Ϊ Nothing
        On Error Resume Next
        strTmp = gobjPlugIn.GetFuncNames(glngSys, glngModul)
        If InStr(",438,0,", "," & err.Number & ",") = 0 Then
            MsgBox "zlPlugIn ��Ҳ���ִ�� GetFuncNames ʱ����" & vbCrLf & err.Number & vbCrLf & err.Description, vbInformation, gstrSysName
        End If
        err.Clear: On Error GoTo 0
        
        If strTmp = "" Then Exit Sub
        
        strTmp = Replace(strTmp, "Auto:", "")
        arrTmp = Split(strTmp, ",")
        For i = 0 To UBound(arrTmp)
            If i <> 0 Then
                Load mnuPlugItem(i)
            End If
            
            mnuPlugItem(i).Caption = CStr(arrTmp(i))
            mnuPlugItem(i).Tag = CStr(arrTmp(i))
            
            If i <= 9 Then
                mnuPlugItem(i).Caption = CStr(arrTmp(i)) & "(&" & IIf(i = 9, 0, i + 1) & ")"
            End If
        Next
    End If
End Sub

Private Sub ExcPlugInFun(ByVal strFunName As String)
    Dim lng�ⷿID As Long
    Dim int���� As Integer
    Dim strNo As String
    
    With mshList
        lng�ⷿID = Val(cboStock.ItemData(cboStock.ListIndex))
        If mlngMode = 1712 Then int���� = 15
        strNo = .TextMatrix(.Row, 0)
    End With
    
    On Error Resume Next
    
    If gobjPlugIn Is Nothing Then
        On Error Resume Next
        Set gobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
        If Not gobjPlugIn Is Nothing Then
            Call gobjPlugIn.Initialize(gcnOracle, glngSys, glngModul)
            If InStr(",438,0,", "," & err.Number & ",") = 0 Then
                MsgBox "zlPlugIn ��Ҳ���ִ�� Initialize ʱ����" & vbCrLf & err.Number & vbCrLf & err.Description, vbInformation, gstrSysName
            End If
        End If
        err.Clear: On Error GoTo 0
    End If
    
    If Not gobjPlugIn Is Nothing Then
        On Error Resume Next
        Call gobjPlugIn.DrugStuffWorkNoramal(mlngMode, strFunName, lng�ⷿID, strNo, int����)
        If InStr(",438,0,", "," & err.Number & ",") = 0 Then
            MsgBox "zlPlugIn ��Ҳ���ִ�� ExecuteFunc ʱ����" & vbCrLf & err.Number & vbCrLf & err.Description, vbInformation, gstrSysName
        End If
        err.Clear: On Error GoTo 0
    End If
End Sub
