VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
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
   Begin VB.PictureBox picColor 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   5280
      ScaleHeight     =   255
      ScaleWidth      =   3615
      TabIndex        =   15
      Top             =   4320
      Width           =   3615
      Begin VB.PictureBox picColor3 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2280
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
         Top             =   0
         Width           =   260
      End
      Begin VB.Label lblColor3 
         AutoSize        =   -1  'True
         Caption         =   "�������"
         Height          =   180
         Left            =   2640
         TabIndex        =   21
         Top             =   30
         Width           =   720
      End
      Begin VB.Label lblColor1 
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Height          =   180
         Left            =   360
         TabIndex        =   20
         Top             =   37
         Width           =   720
      End
      Begin VB.Label lblColor2 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   1680
         TabIndex        =   19
         Top             =   37
         Width           =   360
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfDetail 
      Height          =   1155
      Left            =   0
      TabIndex        =   12
      Top             =   3000
      Width           =   6255
      _cx             =   11033
      _cy             =   2037
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
      BackColorAlternate=   15724527
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
      FormatString    =   $"frmCheckMain.frx":030A
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
   Begin VSFlex8Ctl.VSFlexGrid vsfList 
      Height          =   1455
      Left            =   0
      TabIndex        =   11
      Top             =   1040
      Width           =   6255
      _cx             =   11033
      _cy             =   2566
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
      FormatString    =   $"frmCheckMain.frx":037F
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
   Begin TabDlg.SSTab TabShow 
      Height          =   360
      Left            =   30
      TabIndex        =   7
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
      TabPicture(0)   =   "frmCheckMain.frx":03F4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "�̵���嵥(&2)"
      TabPicture(1)   =   "frmCheckMain.frx":0410
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
   End
   Begin MSComctlLib.ImageList ilsHot 
      Left            =   7110
      Top             =   720
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
            Picture         =   "frmCheckMain.frx":042C
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":064C
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":086C
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":0A88
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":0CA8
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":0EC8
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":10E4
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":1300
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":151A
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":1734
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":188E
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":1AAE
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsCold 
      Left            =   6510
      Top             =   720
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
            Picture         =   "frmCheckMain.frx":1CCE
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":1EEE
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":210E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":232A
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":254A
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":276A
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":2986
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":2BA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":2DBC
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":2FD6
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":3130
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":334C
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Cmd���� 
      Caption         =   "����(&V)"
      Height          =   350
      Left            =   8040
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2555
      Width           =   1100
   End
   Begin VB.PictureBox picSeparate_s 
      BorderStyle     =   0  'None
      Height          =   370
      Left            =   0
      MousePointer    =   7  'Size N S
      ScaleHeight     =   375
      ScaleWidth      =   7935
      TabIndex        =   4
      Top             =   2520
      Width           =   7935
      Begin VB.Label lbl�ɱ����� 
         AutoSize        =   -1  'True
         Caption         =   "�ɱ�����ϼƣ�"
         Height          =   180
         Left            =   6480
         TabIndex        =   14
         Top             =   0
         Width           =   1440
      End
      Begin VB.Label lblSum�ɱ���� 
         AutoSize        =   -1  'True
         Caption         =   "�̵�ɱ����ϼƣ�"
         Height          =   180
         Left            =   4680
         TabIndex        =   13
         Top             =   0
         Width           =   1620
      End
      Begin VB.Label lbl3 
         AutoSize        =   -1  'True
         Caption         =   "�������ϼƣ�"
         Height          =   180
         Left            =   3000
         TabIndex        =   10
         Top             =   0
         Width           =   1440
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         Caption         =   "�̵���ϼƣ�"
         Height          =   180
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   1260
      End
      Begin VB.Label lbl2 
         AutoSize        =   -1  'True
         Caption         =   "����ϼƣ�"
         Height          =   180
         Left            =   1680
         TabIndex        =   8
         Top             =   0
         Width           =   1080
      End
      Begin VB.Label lblRange 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ѯ��Χ��1999��8��12����1999��9��12��"
         Height          =   180
         Left            =   0
         TabIndex        =   5
         Top             =   200
         Width           =   3420
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
            NumButtons      =   18
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
               Caption         =   "ȷ��"
               Key             =   "Affirmant"
               Object.ToolTipText     =   "�¶�ȷ��"
               Object.Tag             =   "ȷ��"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "AffirmantSplit"
               Style           =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Search"
               Description     =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ˢ��"
               Key             =   "Refresh"
               Description     =   "ˢ��"
               Object.ToolTipText     =   "ˢ��"
               Object.Tag             =   "ˢ��"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "FindSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Description     =   "��������"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Exit"
               Description     =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageIndex      =   12
            EndProperty
         EndProperty
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
      Begin VB.Menu mnuEditAffirmant 
         Caption         =   "�¶�ȷ��(&O)"
      End
      Begin VB.Menu mnuEditAffirmantSplit 
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
Private mblnViewCost As Boolean         '�鿴�ɱ���
'Private Const mstrTitle As String = "ҩƷ�̵����"

Public mstrPrivs As String              'Ȩ��

'��������
Private strStart As Date
Private strEnd As Date
Private strVerifyStart As Date
Private strVerifyEnd As Date

Private mlng�ⷿID As Long
Private mintUnit As Integer                 '��λϵ����1-�ۼ�;2-����;3-סԺ;4-ҩ��

Private Const mcstComment As String = "��-��ƽ;��-��ӯ;��-�̿�;����-ͣ��ҩƷ"

'�Ӳ�������ȡҩƷ�۸����������С��λ������ʾ���ȣ�
Private mintShowCostDigit As Integer            '�ɱ���С��λ��
Private mintShowPriceDigit As Integer           '�ۼ�С��λ��
Private mintShowNumberDigit As Integer          '����С��λ��
Private mintShowMoneyDigit As Integer           '���С��λ��

Private mstrCostFormat As String
Private mstrPriceFormat As String
Private mstrNumberFormat As String
Private mstrMoneyFormat As String

Private Const mconint�ۼ۵�λ As Integer = 1
Private Const mconint���ﵥλ As Integer = 2
Private Const mconintסԺ��λ As Integer = 3
Private Const mconintҩ�ⵥλ As Integer = 4

Private Type Type_SQLCondition
    strNO��ʼ As String
    strNO���� As String
    date����ʱ�俪ʼ As Date
    date����ʱ����� As Date
    date���ʱ�俪ʼ As Date
    date���ʱ����� As Date
    lngҩƷ As Long
    lng����ⷿ As Long
    str������ As String
    str����� As String
    lngҩƷ���� As Long
    str���� As String
End Type

Private SQLCondition As Type_SQLCondition
Private Sub cboStock_Click()
    If mlng�ⷿID <> Me.cboStock.ItemData(Me.cboStock.ListIndex) Then
        mlng�ⷿID = Me.cboStock.ItemData(Me.cboStock.ListIndex)
        Call GetDrugDigit(mlng�ⷿID, mstrTitle, mintUnit, mintShowCostDigit, mintShowPriceDigit, mintShowNumberDigit, mintShowMoneyDigit)
        
        '������֯��ʽ����
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
    Dim dateCurrentDate As Date
    Dim strTemp As String
    Dim int��ѯ���� As Integer
    
    mblnBootUp = False
    mlngMode = lngMode
    mstrTitle = strTitle
    mstrPrivs = gstrprivs
    Me.Caption = strTitle
    
    If Not CheckDepend Then Exit Sub            '���������Բ���
    
    mlng�ⷿID = Me.cboStock.ItemData(Me.cboStock.ListIndex)
    Call GetDrugDigit(mlng�ⷿID, mstrTitle, mintUnit, mintShowCostDigit, mintShowPriceDigit, mintShowNumberDigit, mintShowMoneyDigit)
    
    '��֯��ʽ����
    mstrCostFormat = "'999999999990." & String(mintShowCostDigit, "0") & "'"
    mstrPriceFormat = "'999999999990." & String(mintShowPriceDigit, "0") & "'"
    mstrNumberFormat = "'999999999990." & String(mintShowNumberDigit, "0") & "'"
    mstrMoneyFormat = "'999999999990." & String(mintShowMoneyDigit, "0") & "'"
        
    strVerifyStart = "1901-01-01"
    strVerifyEnd = "1901-01-01"
    
    dateCurrentDate = Sys.Currentdate
    int��ѯ���� = Val(zlDatabase.GetPara("��ѯ����", glngSys, mlngMode, 7)) - 1
    strStart = Format(DateAdd("d", -int��ѯ����, dateCurrentDate), "yyyy-MM-dd")
    strEnd = Format(dateCurrentDate, "yyyy-MM-dd")
    
    strFind = " AND A.��¼״̬ = 1 And A.������� is Null And A.�������� Between [3] And [4] "
    SQLCondition.date����ʱ�俪ʼ = CDate(Format(strStart, "yyyy-mm-dd") & " 00:00:00")
    SQLCondition.date����ʱ����� = CDate(Format(strEnd, "yyyy-mm-dd") & " 23:59:59")
    
    mstrFind = strFind
    
    SetVisable  '����Ȩ�����ò�ͬ����ʾ��Ŀ
    TabShow.Tab = 0
    GetList (mstrFind)  '�г�����ͷ
    
    RestoreWinState Me, App.ProductName, mstrTitle
        
    mblnBootUp = True
    
    If IsObject(FrmMain) Then
        Me.Show , FrmMain
    Else
        OS.ShowChildWindow Me.hWnd, FrmMain
    End If
    Me.ZOrder 0
End Sub

'�������������
Private Function CheckDepend() As Boolean
    
    Dim rsDepend As New Recordset
    Dim strStock As String
    
    On Error GoTo errHandle
    CheckDepend = False
    
    strStock = "HIJKLMN"
    gstrSQL = "SELECT DISTINCT a.id, a.���� " _
             & "FROM ��������˵�� c, �������ʷ��� b, ���ű� a " _
            & "Where (a.վ�� = '" & gstrNodeNo & "' Or a.վ�� is Null) And c.�������� = b.���� " _
              & "AND Instr([1],b.����,1) > 0 " _
             & " AND a.id = c.����id " _
              & "AND TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01'" _
              & IIf(zlStr.IsHavePrivs(mstrPrivs, "���пⷿ"), "", " And a.ID IN (Select ����ID From ������Ա Where ��ԱID=[2])")

    Set rsDepend = zlDatabase.OpenSQLRecord(gstrSQL, mstrTitle, strStock, UserInfo.�û�ID)
    
    If rsDepend.EOF Then
        MsgBox "����Ӧ������һ������ҩ�����ʣ�ҩ�����ʣ������Ƽ������ʵĲ���,��鿴���Ź���", vbInformation, gstrSysName
        rsDepend.Close
        Exit Function
    End If
            
    With cboStock
        .Clear
        Do While Not rsDepend.EOF
            .AddItem rsDepend!����
            .ItemData(.NewIndex) = rsDepend!id
            If rsDepend!id = UserInfo.����ID Then
                .ListIndex = .NewIndex
            End If
            rsDepend.MoveNext
        Loop
        rsDepend.Close
        
        If .ListIndex = -1 Then
            If Not zlStr.IsHavePrivs(mstrPrivs, "���пⷿ") Then
                MsgBox "�㲻��ҩ��������Ա�Ҳ��������пⷿ��Ȩ�ޣ����ܽ��룡", vbInformation, gstrSysName
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
    Dim rsList As New Recordset
    Dim strUserPart As String
    Dim str��װϵ�� As String
    Dim strSqlForm As String
    Dim n As Integer
    
    '����ͳ�ƺϼƽ��
    Dim dbl1 As Double
    Dim dbl2 As Double
    Dim dbl3 As Double
    Dim dbl�̵�ɱ���� As Double
    Dim dbl�̵���� As Double

    mlastRow = 0
    On Error GoTo errHandle

    Call FS.ShowFlash("��������ҩƷ��¼,���Ժ� ...", Me)
    DoEvents
    Screen.MousePointer = vbHourglass
    strUserPart = " And A.�ⷿID+0=[11] "
    
    Select Case mintUnit
        Case mconint�ۼ۵�λ
            str��װϵ�� = "1"
        Case mconint���ﵥλ
            str��װϵ�� = "B.�����װ"
        Case mconintסԺ��λ
            str��װϵ�� = "B.סԺ��װ"
        Case mconintҩ�ⵥλ
            str��װϵ�� = "B.ҩ���װ"
    End Select
    
    vsfList.Redraw = flexRDNone
    'Ƶ���ֶα���� �̵�ʱ��
    If TabShow.Tab = 1 Then
        If SQLCondition.str���� <> "" And SQLCondition.lngҩƷ���� = 0 Then
            strSqlForm = " , ������ĿĿ¼ G, ҩƷ���� H"
            strFind = strFind & " And b.ҩ��id = g.Id And g.Id = h.ҩ��id(+) and h.ҩƷ���� in (select * from Table(Cast(f_Str2list([13]) As zlTools.t_Strlist))) and (g.���='5' or g.���='6' or g.���='7')"
        ElseIf SQLCondition.str���� = "" And SQLCondition.lngҩƷ���� <> 0 Then
            strSqlForm = " , ������ĿĿ¼ G"
            strFind = strFind & " And b.ҩ��id = g.Id And g.����id + 0=[12] and (g.���='5' or g.���='6' or g.���='7')"
        ElseIf SQLCondition.str���� <> "" And SQLCondition.lngҩƷ���� <> 0 Then
            strSqlForm = " , ������ĿĿ¼ G, ҩƷ���� H"
            strFind = strFind & " And b.ҩ��id = g.Id And g.Id = h.ҩ��id(+) and h.ҩƷ���� in (select * from Table(Cast(f_Str2list([13]) As zlTools.t_Strlist))) and (g.���='5' or g.���='6' or g.���='7') and g.����id + 0=[12]"
        End If
        
        gstrSQL = "Select NO, �̵�ʱ��, ������, ��������, �����, �������, " & _
                "   to_char(Sum(�̵���), " & mstrMoneyFormat & ") �̵���, to_char(Sum(����), " & mstrMoneyFormat & ") ����,to_char(Sum(�������), " & mstrMoneyFormat & ") �������,to_char(Sum(�̵�ɱ����)," & mstrMoneyFormat & ") �̵�ɱ����, to_char(Sum(�ɱ�����)," & mstrMoneyFormat & ") �ɱ�����, ��¼״̬, ժҪ" & _
                " from ( SELECT a.no,a.���, Ƶ�� AS �̵�ʱ��," _
                & "a.������,TO_CHAR (min(a.��������), 'yyyy-mm-dd HH24:Mi:SS') AS ��������, a.�����," _
                & "TO_CHAR (min(a.�������), 'yyyy-mm-dd HH24:Mi:SS') AS �������, " _
                & "ltrim(to_char(A.�ɱ���+A.���ϵ��*A.���۽��*Decode(��¼״̬, 1, 1, Decode(Mod(��¼״̬, 3), 0, 1, -1))," & mstrMoneyFormat & ")) �̵���," _
                & "ltrim(to_char(���۽��*a.���ϵ��," & mstrMoneyFormat & ")) ����," _
                & "ltrim(to_char((A.����-A.��д����) * a.���ۼ�* Decode(��¼״̬, 1, 1, Decode(Mod(��¼״̬, 3), 0, 1, -1))," & mstrMoneyFormat & ")) AS �������," _
                & "ltrim(to_char((a.�ɱ���+to_char(a.���۽��*a.���ϵ��*Decode(a.��¼״̬, 1, 1, Decode(Mod(a.��¼״̬, 3), 0, 1, -1))," & mstrMoneyFormat & "))-(a.�ɱ����+to_char(a.���*a.���ϵ��*Decode(a.��¼״̬, 1, 1, Decode(Mod(a.��¼״̬, 3), 0, 1, -1))," & mstrMoneyFormat & "))," & mstrMoneyFormat & ")) as �̵�ɱ����," _
                & "ltrim(to_char(a.���۽��*a.���ϵ��-a.���*a.���ϵ��," & mstrMoneyFormat & ")) as �ɱ�����," _
                & " a.��¼״̬, a.ժҪ " _
                & " FROM ҩƷ�շ���¼ a,ҩƷ��� B " & strSqlForm _
                & " Where A.ҩƷID=B.ҩƷID And A.���� = 12  " & strUserPart & strFind _
                & " Group By a.No,a.���, Ƶ��, a.������, a.�����, a.�ɱ���, a.���ϵ��, a.�ɱ���,a.�ɱ����," & str��װϵ�� & ", a.���۽��, a.��¼״̬, a.����, a.��д����, a.���ۼ�, a.����, a.���, a.ժҪ) " _
                & " Group By NO, �̵�ʱ��, ������, ��������, �����, �������, ��¼״̬, ժҪ ORDER BY no DESC,�������� ASC"
    Else
        If SQLCondition.str���� <> "" And SQLCondition.lngҩƷ���� = 0 Then
            strSqlForm = " , ҩƷ��� F, ������ĿĿ¼ G, ҩƷ���� H"
            strFind = strFind & " and a.ҩƷid = f.ҩƷid And f.ҩ��id = g.Id And g.Id = h.ҩ��id(+) and h.ҩƷ���� in (select * from Table(Cast(f_Str2list([13]) As zlTools.t_Strlist))) and (g.���='5' or g.���='6' or g.���='7')"
        ElseIf SQLCondition.str���� = "" And SQLCondition.lngҩƷ���� <> 0 Then
            strSqlForm = " , ҩƷ��� F, ������ĿĿ¼ G"
            strFind = strFind & " and a.ҩƷid = f.ҩƷid And f.ҩ��id = g.Id And g.����id + 0=[12] and (g.���='5' or g.���='6' or g.���='7')"
        ElseIf SQLCondition.str���� <> "" And SQLCondition.lngҩƷ���� <> 0 Then
            strSqlForm = " , ҩƷ��� F, ������ĿĿ¼ G, ҩƷ���� H"
            strFind = strFind & " and a.ҩƷid = f.ҩƷid And f.ҩ��id = g.Id And g.Id = h.ҩ��id(+) and h.ҩƷ���� in(select * from Table(Cast(f_Str2list([13]) As zlTools.t_Strlist))) and (g.���='5' or g.���='6' or g.���='7') and g.����id + 0=[12]"
        End If
        gstrSQL = " SELECT a.no, Ƶ�� AS �̵�ʱ��," _
                    & "a.������,TO_CHAR (min(a.��������), 'yyyy-mm-dd HH24:Mi:SS') AS ��������,a.ժҪ " _
                    & " FROM ҩƷ�շ���¼ a " & strSqlForm _
                    & " Where a.���� = 14  " & strUserPart & strFind _
                    & " Group by a.no,Ƶ��,a.������,a.ժҪ " _
                    & " ORDER BY no DESC,�������� ASC "
    End If
    
    Set rsList = zlDatabase.OpenSQLRecord(gstrSQL, mstrTitle, _
        SQLCondition.strNO��ʼ, _
        SQLCondition.strNO����, _
        SQLCondition.date����ʱ�俪ʼ, _
        SQLCondition.date����ʱ�����, _
        SQLCondition.date���ʱ�俪ʼ, _
        SQLCondition.date���ʱ�����, _
        SQLCondition.lngҩƷ, _
        SQLCondition.lng����ⷿ, _
        SQLCondition.str������, _
        SQLCondition.str�����, _
        cboStock.ItemData(cboStock.ListIndex), _
        SQLCondition.lngҩƷ����, _
        SQLCondition.str����)
        
    Set vsfList.DataSource = rsList
    With vsfList
        If .rows = 1 Then
            .rows = .rows + 100
            .Row = 1
            .Redraw = flexRDDirect
            
            .TopRow = 1
            .rows = .rows - 99
        End If
        .ColAlignment(.ColIndex("�̵�ɱ����")) = flexAlignRightCenter
        .ColAlignment(.ColIndex("�̵���")) = flexAlignRightCenter
        .ColAlignment(.ColIndex("����")) = flexAlignRightCenter
        .ColAlignment(.ColIndex("�������")) = flexAlignRightCenter
        .ColAlignment(.ColIndex("�ɱ�����")) = flexAlignRightCenter
        .Row = 1
        .Col = 0
        .ColSel = .Cols - 1
        If TabShow.Tab = 1 Then
            .ColWidth(.Cols - 2) = 0         'ʼ������"��¼״̬"��һ��
        End If
        
        For n = 0 To .Cols - 1
            .FixedAlignment(n) = flexAlignCenterCenter
        Next
    End With
    SetListColWidth
    
    'ͳ�ƺϼƽ��
    lbl1.Caption = "�̵���ϼƣ�"
    lbl2.Caption = "����ϼƣ�"
    lbl3.Caption = "�������ϼƣ�"
    
    If TabShow.Tab = 1 Then
        If mblnViewCost = False Then
            lblSum�ɱ����.Visible = False
            lbl�ɱ�����.Visible = False
        Else
            lblSum�ɱ����.Visible = True
            lbl�ɱ�����.Visible = True
        End If
        If (Not rsList.EOF) And (Not rsList.BOF) Then
            rsList.MoveFirst
            Do While Not rsList.EOF
                dbl1 = dbl1 + IIf(IsNull(rsList!�̵���), 0, rsList!�̵���)
                dbl2 = dbl2 + IIf(IsNull(rsList!����), 0, rsList!����)
                dbl3 = dbl3 + IIf(IsNull(rsList!�������), 0, rsList!�������)
                dbl�̵�ɱ���� = dbl�̵�ɱ���� + IIf(IsNull(rsList!�̵�ɱ����), 0, rsList!�̵�ɱ����)
                dbl�̵���� = dbl�̵���� + IIf(IsNull(rsList!�ɱ�����), 0, rsList!�ɱ�����)
                rsList.MoveNext
            Loop
            rsList.MoveFirst
            
            lbl1.Caption = "�̵���ϼƣ�" & Format(dbl1, "0." & String(mintShowMoneyDigit, "0"))
            lbl2.Caption = "����ϼƣ�" & Format(dbl2, "0." & String(mintShowMoneyDigit, "0"))
            lbl3.Caption = "�������ϼƣ�" & Format(dbl3, "0." & String(mintShowMoneyDigit, "0"))
            lblSum�ɱ����.Caption = "�̵�ɱ����ϼƣ�" & Format(dbl�̵�ɱ����, "0." & String(mintShowMoneyDigit, "0"))
            lbl�ɱ�����.Caption = "�ɱ����" & Format(dbl�̵����, "0." & String(mintShowMoneyDigit, "0"))
        End If
    Else
        lblSum�ɱ����.Visible = False
        lbl�ɱ�����.Visible = False
    End If
    
    lbl2.Left = lbl1.Width + lbl1.Left + 200
    lbl3.Left = lbl2.Width + lbl2.Left + 200
    lblSum�ɱ����.Left = lbl3.Width + lbl3.Left + 200
    lbl�ɱ�����.Left = lblSum�ɱ����.Width + lblSum�ɱ����.Left + 200
    
    vsfList_EnterCell    '�г�������
    
    SetStrikeColor
    With vsfList
        .Row = 1
        .Col = 0
        .ColSel = .Cols - 1
    End With
    vsfList.Redraw = flexRDDirect
    Call FS.StopFlash
    Screen.MousePointer = vbDefault
    staThis.Panels(2).Text = "��ǰ����" & rsList.RecordCount & "�ŵ���"
    rsList.Close
    If vsfList.Visible = True Then
        vsfList.SetFocus
        vsfList.Row = 1
    End If
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
            intStatus = IIf(TabShow.Tab = 0, 1, Val(.TextMatrix(intRow, .Cols - 2)))
            If intStatus Mod 3 = 0 Then
                .Cell(flexcpForeColor, intRow, 0, intRow, .Cols - 1) = &H80000001
            End If
            If intStatus Mod 3 = 2 Then
                .Cell(flexcpForeColor, intRow, 0, intRow, .Cols - 1) = &HFF
            End If
        Next
    End With
End Sub

'��ͷ�п��ʼ
Private Sub SetListColWidth()
    Dim intCol As Integer
    
    With vsfList
        If TabShow.Tab = 1 Then
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
        Else
            If mblnBootUp = False Then
                .ColWidth(1) = 2000
                .ColWidth(4) = 3000
            End If
        End If
        .ColWidth(.ColIndex("�̵�ɱ����")) = 1500
    End With
    
    Call RestoreFlexState(vsfList, TabShow.TabCaption(TabShow.Tab))
    If TabShow.Tab = 1 And mblnViewCost = False Then
        vsfList.ColHidden(vsfList.ColIndex("�̵�ɱ����")) = True
        vsfList.ColHidden(vsfList.ColIndex("�ɱ�����")) = True
    End If
End Sub

Private Sub SetDetailColWidth()
    Dim intCol As Integer
    
    With vsfDetail
        .ColAlignment(.ColIndex("��λ")) = flexAlignCenterCenter    '��λ
        .ColAlignment(.ColIndex("ʵ����")) = flexAlignRightCenter 'ʵ����
        If TabShow.Tab = 1 Then
            .ColAlignment(.ColIndex("������")) = flexAlignRightCenter     '������
            .ColAlignment(.ColIndex("��־")) = flexAlignCenterCenter    '��־
            .ColAlignment(.ColIndex("������")) = flexAlignRightCenter     '������
            .ColAlignment(.ColIndex("�ɱ���")) = flexAlignRightCenter    '�ɱ���
            .ColAlignment(.ColIndex("�ۼ�")) = flexAlignRightCenter    '�ۼ�
            .ColAlignment(.ColIndex("����")) = flexAlignRightCenter    '����
            .ColAlignment(.ColIndex("��۲�")) = flexAlignRightCenter    '��۲�
            .ColAlignment(.ColIndex("�̵���")) = flexAlignRightCenter    '�̵���
            .ColAlignment(.ColIndex("�������")) = flexAlignRightCenter    '�������
            .ColAlignment(.ColIndex("�̵�ɱ����")) = flexAlignRightCenter    '�̵�ɱ����
            .ColAlignment(.ColIndex("�ɱ�����")) = flexAlignRightCenter    '�ɱ�����
            
        End If
        
        If TabShow.Tab = 1 Then
            If mblnBootUp = False Then
                .ColWidth(0) = 500
                .ColWidth(.ColIndex("ҩƷ��Ϣ")) = 2500
                For intCol = 2 To .Cols - 1
                    .ColWidth(intCol) = 1000
                Next
                .ColWidth(.ColIndex("����ʱ��")) = 0
                .ColWidth(.ColIndex("�̵�ɱ����")) = 1500
            End If
        Else
            .ColWidth(0) = 500
            .ColWidth(.ColIndex("ҩƷ��Ϣ")) = 2500
            For intCol = 2 To .Cols - 1
                .ColWidth(intCol) = 1000
            Next
        End If
        
        Call RestoreFlexState(vsfDetail, TabShow.TabCaption(TabShow.Tab))
        If TabShow.Tab = 1 And mblnViewCost = False Then
            .ColHidden(.ColIndex("�ɱ���")) = True
            .ColHidden(.ColIndex("��۲�")) = True
            .ColHidden(.ColIndex("�̵�ɱ����")) = True
            .ColHidden(.ColIndex("�ɱ�����")) = True
        End If
    End With
End Sub


'����Ȩ�����ò�ͬ����ʾ��Ŀ
Private Sub SetVisable()
    '�⹺�������Ȩ�ޣ��������á����������пⷿ���Ǽǡ��޸ġ�ɾ�������ա����������ݴ�ӡ
'    If Not zlStr.IsHavePrivs(mstrPrivs, "��������") Then
'         mnuFileParameter.Visible = False
'         mnuFileLine3.Visible = False                '��Ӧ�ķָ���
'    End If
     
    If Not zlStr.IsHavePrivs(mstrPrivs, "�Ǽ�") Then
        mnuEditAddBill.Visible = False
        mnuEditAddTable.Visible = False
        tlbTool.Buttons("Bill").Visible = False
        tlbTool.Buttons("Table").Visible = False
    End If
    
    If Not zlStr.IsHavePrivs(mstrPrivs, "�޸�") Then
        mnuEditModify.Visible = False
        tlbTool.Buttons("Modify").Visible = False
    End If
    
    If Not zlStr.IsHavePrivs(mstrPrivs, "ɾ��") Then
        mnuEditDel.Visible = False
        tlbTool.Buttons("Delete").Visible = False
         '��û�����б༭Ȩ��ʱ���Ѳ˵��͹������ϵ���Ӧ�ķָ������Ρ�
        If mnuEditAddBill.Visible = False And mnuEditModify.Visible = False Then
            mnuEditLine1.Visible = False
            tlbTool.Buttons("EditSeparate").Visible = False
        End If
    End If
    
    If Not zlStr.IsHavePrivs(mstrPrivs, "���") Then
        mnuEditVerify.Visible = False
        tlbTool.Buttons("Verify").Visible = False
    End If
    
    If Not zlStr.IsHavePrivs(mstrPrivs, "�¶�ȷ��") Then
        mnuEditAffirmant.Visible = False
        mnuEditAffirmantSplit.Visible = False
        tlbTool.Buttons("Affirmant").Visible = False
        tlbTool.Buttons("AffirmantSplit").Visible = False
    End If
    
    If Not zlStr.IsHavePrivs(mstrPrivs, "����") Then
        mnuEditStrike.Visible = False
        tlbTool.Buttons("Strike").Visible = False
        
        If mnuEditVerify.Visible = False Then
            mnuEditLine2.Visible = False
            tlbTool.Buttons("VerifySeparate").Visible = False
        End If
    End If
    If Not zlStr.IsHavePrivs(mstrPrivs, "ȫ����Ϊ��") Then
        mnuEditAddTableZero.Visible = False
        tlbTool.Buttons("Table").ButtonMenus("Zero").Visible = False
    End If
    If Not zlStr.IsHavePrivs(mstrPrivs, "���ݴ�ӡ") Then
        mnuFileBillPrint.Visible = False
        mnuFileBillPreview.Visible = False
    End If
    If Not zlStr.IsHavePrivs(mstrPrivs, "ҩƷ�̵��") Then
        mnuEditAddTable.Visible = False
        tlbTool.Buttons("Table").Visible = False
        TabShow.TabVisible(1) = False
    End If
End Sub

Private Sub Cmd����_Click()
    Call mnuEditDisplay_Click
End Sub

Private Sub Form_Activate()
    If vsfList.Visible = True Then
        vsfList.SetFocus
        vsfList.Row = 1
        vsfDetail.Row = 1
    End If
End Sub

Private Sub Form_Load()
    '�ָ�����
    Dim dateCurrentDate As Date
    
    Me.Caption = mstrTitle
    mblnViewCost = zlStr.IsHavePrivs(mstrPrivs, "�鿴�ɱ���")
    dateCurrentDate = Sys.Currentdate
    lblRange.Caption = "��ѯ��Χ:" & Format(dateCurrentDate, "yyyy��MM��dd��") & "��" & Format(dateCurrentDate, "yyyy��MM��dd��")
    
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngMode, mstrPrivs)
    staThis.Panels(2).Picture = picColor
End Sub
Private Sub Form_Resize()
    '����λ������
    
    On Error Resume Next

    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.WindowState <> vbMaximized Then
        If Me.Height < 8145 Then
            Me.Height = 8145
        End If
        Me.Top = (Screen.Height - Me.Height) / 2
    End If
   
    cbrTool.Bands(1).MinHeight = tlbTool.Height
    
    With cbrTool
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth - .Left
    End With
    
    With picSeparate_s
        .Height = 370
        .Left = 0
        .Width = cbrTool.Width
        
    End With
    
    With TabShow
        .Top = IIf(cbrTool.Visible, cbrTool.Height, 0)
        .Left = 0
        .Width = cbrTool.Width
    End With
    
    With vsfList
        .Top = TabShow.Top + TabShow.Height
        .Left = 0
        .Width = cbrTool.Width
        .Height = picSeparate_s.Top - .Top
    End With
    
    With Cmd����
        .Left = Me.ScaleWidth - .Width - 100
        .Top = vsfList.Top + vsfList.Height + 30
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
    picColor3.Visible = False
    lblColor3.Visible = False
    picColor.Width = lblColor2.Left + lblColor2.Width + 20
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName, mstrTitle
    Call SaveFlexState(vsfList, TabShow.TabCaption(TabShow.Tab))
    Call SaveFlexState(vsfDetail, TabShow.TabCaption(TabShow.Tab))
End Sub
Private Sub mnuEditaddBill_Click()
    Dim strNo As String
    Dim BlnSuccess As Boolean
    
    frmCheckCourseCard.ShowCard Me, strNo, 1, , BlnSuccess
    
    If BlnSuccess Then Call mnuViewRefresh_Click
End Sub

Private Sub mnuEditAddTableAuto_Click()
    Dim strNo As String
    Dim BlnSuccess As Boolean
    
    '��鱾���Ƿ��Ѿ���˽�棬���δ��˽�����ܽ�������ҵ�����
'    If CheckIsAccount(Val(cboStock.ItemData(cboStock.ListIndex))) = False Then
'        Exit Sub
'    End If
    
    frmCheckCard.ShowCard Me, strNo, 1, , BlnSuccess
    
    If BlnSuccess Then
        Call mnuViewRefresh_Click
    End If
End Sub

Private Sub mnuEditAddTableTotal_Click()
    Dim strNo As String
    Dim BlnSuccess As Boolean
    
    '��鱾���Ƿ��Ѿ���˽�棬���δ��˽�����ܽ�������ҵ�����
'    If CheckIsAccount(Val(cboStock.ItemData(cboStock.ListIndex))) = False Then
'        Exit Sub
'    End If
    
    frmCheckCard.ShowCard Me, strNo, 5, , BlnSuccess
    
    If BlnSuccess Then
        Call mnuViewRefresh_Click
    End If
End Sub

Private Sub mnuEditAddTableZero_Click()
    Dim strNo As String
    Dim BlnSuccess As Boolean
    
    '��鱾���Ƿ��Ѿ���˽�棬���δ��˽�����ܽ�������ҵ�����
'    If CheckIsAccount(Val(cboStock.ItemData(cboStock.ListIndex))) = False Then
'        Exit Sub
'    End If
   
    frmCheckCard.ShowCard Me, strNo, 6, , BlnSuccess
    
    If BlnSuccess Then
        Call mnuViewRefresh_Click
    End If
End Sub

Private Sub mnuEditAffirmant_Click()
    Dim str������� As String       'ȱʡ��Ϊȷ�ϼ�¼�Ľ�������
    '��д�¶�ȷ�ϼ�¼
    If TabShow.Tab = 1 Then
        str������� = vsfList.TextMatrix(vsfList.Row, 5)
    End If
    With frm�¶�ȷ��
        Call .ShowEditor(Me.cboStock.ItemData(Me.cboStock.ListIndex), str�������)
    End With
End Sub

Private Sub mnuEditVerify_Click()
    '����
    
    Dim strNo As String
    Dim BlnSuccess As Boolean
    
    With vsfList
        strNo = .TextMatrix(.Row, 0)
        frmCheckCard.ShowCard Me, strNo, 3, .TextMatrix(.Row, .Cols - 2), BlnSuccess
    End With
    
    If BlnSuccess Then Call mnuViewRefresh_Click
End Sub

Private Sub mnuEditDel_Click()
    'ɾ��
    Dim strBillNo As String
    Dim intRow As Integer
    Dim strTitle As String
    Dim intReturn As Integer
    Dim intRecord As Integer
     
    With vsfList
        strTitle = IIf(TabShow.Tab = 0, "�̵��¼��", "�̵��")
        
        On Error GoTo errHandle
        intRow = .Row
        strBillNo = .TextMatrix(intRow, 0)
        intReturn = MsgBox("��ȷʵҪɾ�����ݺ�Ϊ��" & strBillNo & "����" & strTitle & "��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)
        intRecord = .rows - 1
        If intReturn = vbYes Then
            If TabShow.Tab = 1 Then
                gstrSQL = "zl_ҩƷ�̵�_Delete('" & strBillNo & "')"
            Else
                gstrSQL = "zl_ҩƷ�̵��¼��_Delete('" & strBillNo & "')"
            End If
            Call zlDatabase.ExecuteProcedure(gstrSQL, mstrTitle)
            
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
    mnuViewRefresh_Click
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
        If TabShow.Tab = 0 Then
            frmCheckCourseCard.ShowCard Me, strNo, 4
        Else
            frmCheckCard.ShowCard Me, strNo, 4, .TextMatrix(.Row, .Cols - 2)
        End If
    End With
End Sub

Private Sub mnuEditStrike_Click()
    Dim blnPurchase As Boolean, blnRefresh As Boolean
    
    '������⹺(blnPurchaseΪ��)����ֱ�ӽ������
    'ѯ���Ƿ����(blnPurchaseΪ��ʾ�򷵻�ֵ)������������
    blnPurchase = (InStr(1, "1300,1302,1304,1305,1306", mlngMode) <> 0)
    With vsfList
        If Not blnPurchase Then
            blnPurchase = (MsgBox("��ȷʵҪȫ���������ݺ�Ϊ��" & .TextMatrix(.Row, 0) & "���ĵ�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes)
        End If
        If blnPurchase Then
            blnRefresh = StrikeSave
            If blnRefresh Then mnuViewRefresh_Click
        End If
    End With
End Sub

Private Function StrikeSave() As Boolean
    Dim BlnSuccess As Boolean
    Dim rsTemp As ADODB.Recordset
    Dim int����� As Integer
    Dim strMsg As String
    Dim n As Integer
    
    StrikeSave = False
    
    int����� = MediWork_GetCheckStockRule(mlng�ⷿID)
    
    On Error GoTo errHandle
    If int����� <> 0 Then
        gstrSQL = "Select A.ҩƷ��Ϣ " & _
            " From (Select Distinct '(' || I.���� || ')' || Nvl(N.����, I.����) As ҩƷ��Ϣ, A.ʵ������, Nvl(K.ʵ������, 0) As ������� " & _
            " From ҩƷ�շ���¼ A, (Select ҩƷid, �ⷿid, ʵ������, Nvl(����, 0) ���� From ҩƷ��� Where ���� = 1) K, ҩƷ��� B, �շ���ĿĿ¼ I, �շ���Ŀ���� N " & _
            " Where A.ҩƷid = K.ҩƷid(+) And A.�ⷿid = K.�ⷿid(+) And Nvl(A.����, 0) = K.����(+) And A.ҩƷid = B.ҩƷid And " & _
            " A.ҩƷid = I.ID And A.ҩƷid = N.�շ�ϸĿid(+) And N.����(+) = 3 And A.���� = 12 And A.���ϵ�� = 1 And A.NO = [1]) A " & _
            " Where A.ʵ������ > A.������� "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�����", vsfList.TextMatrix(vsfList.Row, 0))
        
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
    
    With vsfList
        gstrSQL = "zl_ҩƷ�̵�_Strike('" & .TextMatrix(.Row, 0) & "','" & UserInfo.�û����� & "')"
        
        Call zlDatabase.ExecuteProcedure(gstrSQL, mstrTitle)
        
        '��ʾͣ��ҩƷ
        Call CheckStopMedi(���ݺ�.�̵�� & "|" & .TextMatrix(.Row, 0))
    End With
    StrikeSave = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    
    'MsgBox "����ʧ�ܣ�", vbInformation, gstrSysName
    Call SaveErrLog

End Function


Private Sub mnuEditModify_Click()
    '�޸�
    Dim strNo As String
    Dim BlnSuccess As Boolean
    
    BlnSuccess = False
    With vsfList
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        strNo = .TextMatrix(.Row, 0)
        If TabShow.Tab = 0 Then
            frmCheckCourseCard.ShowCard Me, strNo, 2, 1, BlnSuccess
        Else
            frmCheckCard.ShowCard Me, strNo, 2, vsfList.TextMatrix(vsfList.Row, vsfList.Cols - 2), BlnSuccess
        End If
        
        If BlnSuccess Then Call mnuViewRefresh_Click
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
        ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1307", "zl8_bill_1307"), Me, "���ݱ��=" & .TextMatrix(.Row, 0), "��¼״̬=" & .TextMatrix(.Row, .Cols - 2), "��λϵ��=" & int��λϵ��, 1
    End With
End Sub
Private Sub mnuFileBillPrint_Click()
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
        ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1307", "zl8_bill_1307"), Me, "���ݱ��=" & .TextMatrix(.Row, 0), "��¼״̬=" & .TextMatrix(.Row, .Cols - 2), "��λϵ��=" & int��λϵ��, 2
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
    Dim dateCurrentDate As Date
    Dim int��ѯ���� As Integer
    
    frm��������.���ò��� Me, mstrPrivs, mstrTitle
    
    dateCurrentDate = Sys.Currentdate
    int��ѯ���� = Val(zlDatabase.GetPara("��ѯ����", glngSys, mlngMode, 7)) - 1
    strStart = Format(DateAdd("d", -int��ѯ����, dateCurrentDate), "yyyy-MM-dd")
    strEnd = Format(dateCurrentDate, "yyyy-MM-dd")
    
    SQLCondition.date����ʱ�俪ʼ = CDate(Format(strStart, "yyyy-mm-dd") & " 00:00:00")
    SQLCondition.date����ʱ����� = CDate(Format(strEnd, "yyyy-mm-dd") & " 23:59:59")
    
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
    Dim StrWinName As String
    With vsfList
        StrWinName = "frmMainList8"
    End With
    Call ShowHelp(App.ProductName, Me.hWnd, StrWinName)
End Sub

Private Sub mnuHelpWebHome_Click()
    '������ҳ
    Call zlHomePage(Me.hWnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    '���ͷ���
    Call zlMailTo(Me.hWnd)
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    'Ĭ�ϲ�����ҩƷ=ҩƷid���ⷿ=�ⷿid����ʼʱ��=���ƿ�ʼʱ�䣬����ʱ��=���ƽ���ʱ�䣬�̵㵥=�̵㵥NO���̵��=�̵��NO
    Dim str��ʼʱ�� As String
    Dim str����ʱ�� As String
    Dim strNo As String
    Dim strReportName As String
    
    strReportName = Split(mnuReportItem(Index).Tag, ",")(1)
    
    Select Case strReportName
        Case "ZL1_INSIDE_1307"
            Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1307", Me, "�ⷿ=" & Me.cboStock.Text & "|" & IIf(Me.cboStock.ItemData(Me.cboStock.ListIndex) = 0, " is not null ", "=" & Me.cboStock.ItemData(Me.cboStock.ListIndex)))
        Case "ZL1_INSIDE_1307_1"
            Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1307_1", Me, "�ⷿ=" & Me.cboStock.Text & "|" & IIf(Me.cboStock.ItemData(Me.cboStock.ListIndex) = 0, " is not null ", "=" & Me.cboStock.ItemData(Me.cboStock.ListIndex)), "��λ=" & Choose(mintUnit, "�ۼ۵�λ", "���ﵥλ", "סԺ��λ", "ҩ�ⵥλ") & "|" & Choose(mintUnit, 1, 3, 4, 2))
        Case Else
            If vsfList.Row >= 1 And LTrim(vsfList.TextMatrix(vsfList.Row, 0)) <> "" Then
                strNo = vsfList.TextMatrix(vsfList.Row, 0)
            End If
            
            str��ʼʱ�� = IIf(Format(SQLCondition.date����ʱ�俪ʼ, "yyyy-mm-dd") = "1899-12-30", "", Format(SQLCondition.date����ʱ�俪ʼ, "yyyy-mm-dd"))
            str����ʱ�� = IIf(Format(SQLCondition.date����ʱ�����, "yyyy-mm-dd") = "1899-12-30", "", Format(SQLCondition.date����ʱ�����, "yyyy-mm-dd"))
            
            Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
                "ҩƷ=" & IIf(SQLCondition.lngҩƷ = 0, "", SQLCondition.lngҩƷ), _
                "�ⷿ=" & IIf(Val(cboStock.ItemData(cboStock.ListIndex)) = 0, "", Val(cboStock.ItemData(cboStock.ListIndex))), _
                "��ʼʱ��=" & str��ʼʱ��, _
                "����ʱ��=" & str����ʱ��, _
                "�̵㵥=" & strNo, _
                "�̵��=" & strNo)
    End Select
End Sub

Private Sub mnuViewRefresh_Click()
    'ˢ��
    GetList mstrFind
End Sub

Private Sub mnuViewSearch_Click()
    '����
    
    Dim strFind As String
    
    strFind = FrmTransferSearch.GetSearch(Me, mlngMode, mlng�ⷿID, strStart, strEnd, strVerifyStart, strVerifyEnd, _
                SQLCondition.strNO��ʼ, _
                SQLCondition.strNO����, _
                SQLCondition.date����ʱ�俪ʼ, _
                SQLCondition.date����ʱ�����, _
                SQLCondition.date���ʱ�俪ʼ, _
                SQLCondition.date���ʱ�����, _
                SQLCondition.lngҩƷ, _
                SQLCondition.lng����ⷿ, _
                SQLCondition.str������, _
                SQLCondition.str�����, _
                SQLCondition.lngҩƷ����, _
                SQLCondition.str����)
    
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




Private Sub vsfDetail_EnterCell()
    With vsfDetail
        If .Row = 0 Then Exit Sub
        
        .Redraw = flexRDNone
        .ForeColorSel = .Cell(flexcpForeColor, .Row, 1)
        .Redraw = flexRDDirect
    End With
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
    Dim intBill As Integer                      '��������  �磺1���⹺��⣻2��
    Dim str��װϵ�� As String
    Dim str��λ�ֶ� As String
    Dim strOrder As String
    Dim strCompare As String
    Dim strSqlЧ�� As String
    Dim lngColor As Long
    Dim n As Long
    Dim i As Integer
    Dim intCol As Integer
    Dim strSqlҩ�� As String
    Dim strSqlOrder As String
    
    If mlastRow = vsfList.Row Then Exit Sub
    mlastRow = vsfList.Row
    
    On Error GoTo errHandle
    With vsfList
        .Redraw = flexRDNone
        .ForeColorSel = .Cell(flexcpForeColor, mlastRow, 1)
        .Redraw = flexRDDirect
    End With
    
    strOrder = zlDatabase.GetPara("����", glngSys, ģ���.ҩƷ�̵�)
    strCompare = Mid(strOrder, 1, 1)
    
    strSqlOrder = "���"
    
    If strCompare = "0" Then
        '���������
        strSqlOrder = "���"
    ElseIf strCompare = "1" Then
        '����������
        strSqlOrder = "ҩƷ��Ϣ"
    ElseIf strCompare = "2" Then
        '����������
        strSqlOrder = "Substr(ҩƷ��Ϣ, Instr(ҩƷ��Ϣ, ']') + 1)"
    End If
    
    strSqlOrder = strSqlOrder & IIf(Right(strOrder, 1) = "0", " ASC", " DESC") & ",ҩƷ��Ϣ,���"
    
    If vsfList.Row >= 1 And LTrim(vsfList.TextMatrix(vsfList.Row, 0)) <> "" Then
        vsfList.Col = 0
        vsfList.ColSel = vsfList.Cols - 1
        
        vsfDetail.Redraw = flexRDNone
        Select Case mintUnit
            Case mconint�ۼ۵�λ
                str��װϵ�� = "1"
                str��λ�ֶ� = "I.���㵥λ"
            Case mconint���ﵥλ
                str��װϵ�� = "B.�����װ"
                str��λ�ֶ� = "B.���ﵥλ"
            Case mconintסԺ��λ
                str��װϵ�� = "B.סԺ��װ"
                str��λ�ֶ� = "B.סԺ��λ"
            Case mconintҩ�ⵥλ
                str��װϵ�� = "B.ҩ���װ"
                str��λ�ֶ� = "B.ҩ�ⵥλ"
        End Select
        
        strSqlЧ�� = IIf(gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1, "TO_CHAR(A.Ч��-1,'YYYY-MM-DD') AS ��Ч����", "TO_CHAR(A.Ч��,'YYYY-MM-DD') AS ʧЧ��")
        
        If gintҩƷ������ʾ = 0 Then
            strSqlҩ�� = ",('['||I.����||']'||I.����) AS ҩƷ��Ϣ"
        ElseIf gintҩƷ������ʾ = 1 Then
            strSqlҩ�� = ",('['||I.����||']'||NVL(N.����,I.����)) AS ҩƷ��Ϣ"
        Else
            strSqlҩ�� = ",('['||I.����||']'||I.����) AS ҩƷ��Ϣ,N.���� As ��Ʒ��"
        End If
        
        intBill = IIf(TabShow.Tab = 1, 12, 14)
        If TabShow.Tab = 1 Then
            gstrSQL = "Select DISTINCT a.���" & strSqlҩ�� & "," _
                    & "     B.ҩƷ��Դ,B.����ҩ��,I.���,a.����," & str��λ�ֶ� & " as ��λ,a.����," & strSqlЧ�� & ",a.��׼�ĺ�," _
                    & "     LTRIM(to_char(A.��д���� /" & str��װϵ�� & ",decode(a.����,0,'999999999990.00000'," & mstrNumberFormat & "))) AS ������," _
                    & "     LTRIM(to_char(A.���� /" & str��װϵ�� & "," & mstrNumberFormat & ")) AS ʵ����," _
                    & "     Decode(Sign(A.����-A.��д����),-1,'��',1,'ӯ','ƽ') as ��־," _
                    & "     LTRIM(to_char(A.ʵ������ /" & str��װϵ�� & ",decode(a.����,0,'999999999990.00000'," & mstrNumberFormat & "))) AS ������," _
                    & "     LTRIM(TO_CHAR (a.����*" & str��װϵ�� & ", " & mstrCostFormat & ")) AS �ɱ���," _
                    & "     LTRIM(TO_CHAR (a.���ۼ�*" & str��װϵ�� & ", " & mstrPriceFormat & ")) AS �ۼ�," _
                    & "     LTRIM(TO_CHAR (a.���۽��*a.���ϵ��,decode(a.����,0,'999999999990.00000', " & mstrMoneyFormat & "))) AS ����," _
                    & "     LTRIM(TO_CHAR ((A.����-A.��д����) * a.���ۼ�* Decode(��¼״̬, 1, 1, Decode(Mod(��¼״̬, 3), 0, 1, -1))," & mstrMoneyFormat & ")) AS �������," _
                    & "     LTRIM(TO_CHAR (a.���*a.���ϵ��, decode(a.����,0,'999999999990.00000'," & mstrMoneyFormat & "))) AS ��۲�, " _
                    & "     LTrim(To_Char((a.���� / b.�����װ)*(a.���ۼ� * b.�����װ), " & mstrMoneyFormat & ")) As �̵���," _
                    & "     LTrim(To_Char(((a.�ɱ���+to_char(a.���۽��*a.���ϵ��*Decode(a.��¼״̬, 1, 1, Decode(Mod(a.��¼״̬, 3), 0, 1, -1))," & mstrMoneyFormat & "))-(a.�ɱ����+to_char(a.���*a.���ϵ��*Decode(a.��¼״̬, 1, 1, Decode(Mod(a.��¼״̬, 3), 0, 1, -1))," & mstrMoneyFormat & ")))," & mstrMoneyFormat & ")) as �̵�ɱ����, " _
                    & "     ltrim(To_Char((a.���۽��*a.���ϵ�� - a.���*a.���ϵ�� ), " & mstrMoneyFormat & ")) As �ɱ�����," _
                    & " Nvl(I.����ʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')) As ����ʱ��,e.�ⷿ��λ " _
                    & " From (Select a.���ϵ��,a.��¼״̬,a.���,a.ҩƷid,a.����,a.����,a.Ч��,A.��д����,A.����,A.ʵ������,a.�ɱ���,a.�ɱ����,a.���ۼ�,a.���۽��,a.���,a.����,a.��׼�ĺ�,a.�ⷿid" _
                    & "         From ҩƷ�շ���¼ a" _
                    & "        Where a.��¼״̬= [2] " _
                    & "             And a.����= 12 And a.No=[1]) a," _
                    & "        ҩƷ��� b,�շ���ĿĿ¼ I ,�շ���Ŀ���� n,ҩƷ�����޶� e" _
                    & " Where a.ҩƷid = b.ҩƷid And a.ҩƷid = i.Id" _
                    & "        And a.ҩƷid=n.�շ�ϸĿid(+) And n.����(+)=3 " _
                    & "        And a.ҩƷid = e.ҩƷid and a.�ⷿid = e.�ⷿid " _
                    & " ORDER BY " & strSqlOrder
        Else
            gstrSQL = "Select DISTINCT a.���" & strSqlҩ�� & "," _
                    & "     B.ҩƷ��Դ,B.����ҩ��,I.���,a.����," & str��λ�ֶ� & " as ��λ,a.����," & strSqlЧ�� & ",a.��׼�ĺ�," _
                    & "     to_char(A.���� /" & str��װϵ�� & "," & mstrNumberFormat & ") AS ʵ����" _
                    & " From (Select a.���,a.ҩƷid,a.����,a.����,a.Ч��,A.��д����,A.����,A.ʵ������,a.���ۼ�,a.���۽��,a.���,a.��׼�ĺ�" _
                    & "         From ҩƷ�շ���¼ a" _
                    & "        Where a.��¼״̬= 1 And a.����= 14 And a.No=[1]) a," _
                    & "        ҩƷ��� b,�շ���ĿĿ¼ I ,�շ���Ŀ���� n" _
                    & " Where a.ҩƷid = b.ҩƷid And a.ҩƷid = i.Id" _
                    & "        And a.ҩƷid=n.�շ�ϸĿid(+) And n.����(+)=3 " _
                    & " ORDER BY " & strSqlOrder
        End If
        Set rsDetail = zlDatabase.OpenSQLRecord(gstrSQL, mstrTitle, vsfList.TextMatrix(vsfList.Row, 0), vsfList.TextMatrix(vsfList.Row, vsfList.Cols - 2))
        
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
                
        '��ʽ������
'        With vsfDetail
'            Select Case TabShow.Tab
'            Case 0
'                For n = 1 To .rows - 1
'                    .TextMatrix(n, .ColIndex("ʵ����")) = zlStr.FormatEx(.TextMatrix(n, .ColIndex("ʵ����")), mintShowNumberDigit)
'                Next
'            Case 1
'                For n = 1 To .rows - 1
'                    .TextMatrix(n, .ColIndex("������")) = zlStr.FormatEx(.TextMatrix(n, .ColIndex("������")), mintShowNumberDigit)
'                    .TextMatrix(n, .ColIndex("ʵ����")) = zlStr.FormatEx(.TextMatrix(n, .ColIndex("ʵ����")), mintShowNumberDigit)
'                    .TextMatrix(n, .ColIndex("������")) = zlStr.FormatEx(.TextMatrix(n, .ColIndex("������")), mintShowNumberDigit)
'                Next
'            End Select
'        End With
        
        vsfDetail.Redraw = flexRDDirect
    ElseIf LTrim(vsfList.TextMatrix(vsfList.Row, 0)) = "" Then
        With vsfDetail
            .Cols = IIf(TabShow.Tab = 1, 24, 11)
            If gintҩƷ������ʾ = 2 Then .Cols = .Cols + 1
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
            .TextMatrix(0, intCol) = "����": intCol = intCol + 1
            .TextMatrix(0, intCol) = "��λ": intCol = intCol + 1
            .TextMatrix(0, intCol) = "����": intCol = intCol + 1
            .TextMatrix(0, intCol) = IIf(gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1, "��Ч����", "ʧЧ��"): intCol = intCol + 1
            .TextMatrix(0, intCol) = "��׼�ĺ�": intCol = intCol + 1
            If TabShow.Tab = 0 Then
                .TextMatrix(0, intCol) = "ʵ����": intCol = intCol + 1
            Else
                .TextMatrix(0, intCol) = "������": intCol = intCol + 1
                .TextMatrix(0, intCol) = "ʵ����": intCol = intCol + 1
                .TextMatrix(0, intCol) = "��־": intCol = intCol + 1
                .TextMatrix(0, intCol) = "������": intCol = intCol + 1
                .TextMatrix(0, intCol) = "�ɱ���": intCol = intCol + 1
                .TextMatrix(0, intCol) = "�ۼ�": intCol = intCol + 1
                .TextMatrix(0, intCol) = "����": intCol = intCol + 1
                .TextMatrix(0, intCol) = "��۲�": intCol = intCol + 1
                .TextMatrix(0, intCol) = "�̵�ɱ����": intCol = intCol + 1
                .TextMatrix(0, intCol) = "�̵���": intCol = intCol + 1
                .TextMatrix(0, intCol) = "�ɱ�����": intCol = intCol + 1
                .TextMatrix(0, intCol) = "�������": intCol = intCol + 1
                .TextMatrix(0, intCol) = "����ʱ��": intCol = intCol + 1
                .TextMatrix(0, intCol) = "�ⷿ��λ": intCol = intCol + 1
            End If
            
            .Row = 1
            .Col = 0
            .ColSel = .Cols - 1
            
            For n = 0 To .Cols - 1
                .ColKey(n) = .TextMatrix(0, n)
                .FixedAlignment(n) = flexAlignCenterCenter
            Next
        End With
    End If
    SetDetailColWidth
    SetEnable
    
    '��ɫ
    If TabShow.Tab = 1 Then
        With vsfDetail
            .Redraw = flexRDNone
            For n = 1 To .rows - 1
                If .TextMatrix(n, 0) <> "" Then
                    If .TextMatrix(n, .ColIndex("��־")) = "ӯ" Then
                        lngColor = vbRed
                    ElseIf .TextMatrix(n, .ColIndex("��־")) = "��" Then
                        lngColor = vbBlue
                    Else
                        lngColor = vbBlack
                    End If
                    
                    '�̿���ӯ������ɫ���֣�
                    If lngColor <> vbBlack Then
                        .Cell(flexcpForeColor, n, 0, n, .Cols - 1) = lngColor
                    End If
                    
                    '�����ͣ��ҩƷ�����д�����ʾ
                    If Format(.TextMatrix(n, .ColIndex("����ʱ��")), "YYYY-MM-DD") <> "3000-01-01" Then
                        .Cell(flexcpFontBold, n, 0, n, .Cols - 1) = True
                    End If
                End If
            Next
            .Redraw = flexRDDirect
        End With
    End If
    
    vsfDetail.Row = 1
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
        .Height = picSeparate_s.Top - .Top
    End With
    
    With Cmd����
        .Top = vsfList.Top + vsfList.Height + 30
    End With
    
    With vsfDetail
        .Top = picSeparate_s.Top + picSeparate_s.Height + 100
        .Height = ScaleHeight - .Top - IIf(staThis.Visible, staThis.Height, 0)
    End With
End Sub

Private Sub TabShow_Click(PreviousTab As Integer)
    Call SaveFlexState(vsfList, TabShow.TabCaption(PreviousTab))
    Call SaveFlexState(vsfDetail, TabShow.TabCaption(PreviousTab))
    mblnBootUp = False
    If TabShow.Tab = 1 Then
        vsfDetail.ToolTipText = mcstComment
    Else
        vsfDetail.ToolTipText = ""
    End If
    GetList (mstrFind)  '�г�����ͷ
    mblnBootUp = True
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
        Case "Affirmant"
            mnuEditAffirmant_Click
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
    mnuEditVerify.Visible = blnVisible And zlStr.IsHavePrivs(mstrPrivs, "���")
    mnuEditStrike.Visible = blnVisible And zlStr.IsHavePrivs(mstrPrivs, "����")
    mnuEditLine2.Visible = blnVisible
    tlbTool.Buttons("Verify").Visible = blnVisible And zlStr.IsHavePrivs(mstrPrivs, "���")
    tlbTool.Buttons("Strike").Visible = blnVisible And zlStr.IsHavePrivs(mstrPrivs, "����")
    tlbTool.Buttons("VerifySeparate").Visible = blnVisible
    
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
            mnuFileExcel.Enabled = True
            tlbTool.Buttons("Print").Enabled = True
            tlbTool.Buttons("PrintView").Enabled = True
            
            mnuFileBillPreview.Enabled = TabShow.Tab = 1
            mnuFileBillPrint.Enabled = TabShow.Tab = 1
            
            If TabShow.Tab = 1 Then
                strVerify = .TextMatrix(.Row, .Cols - 8)
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
                If .TextMatrix(.Row, .Cols - 2) Mod 3 = 0 Then
                    .ToolTipText = "�������ݵ�ԭ����"
                    If mnuEditStrike.Visible = True Then
                        mnuEditStrike.Enabled = True
                        tlbTool.Buttons("Strike").Enabled = True
                    End If
                ElseIf .TextMatrix(.Row, .Cols - 2) Mod 3 = 2 Then
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
        
    objRow.Add "��ӡ��:" & UserInfo.�û�����
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
    objRow.Add "NO." & Trim(vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "NO")))
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "�̵�ⷿ��" & Trim(cboStock.Text)
    objRow.Add "�̵�ʱ�䣺" & Trim(vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "�̵�ʱ��")))
    objPrint.UnderAppRows.Add objRow
        
    Set objRow = New zlTabAppRow
    objRow.Add "ժҪ:" & vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "ժҪ"))
    objPrint.BelowAppRows.Add objRow
        
    Set objRow = New zlTabAppRow
    objRow.Add "������:" & vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "������")) & "  ��������:" & vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "��������"))
    
    If TabShow.Tab = 1 Then
        objRow.Add "�����:" & vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "�����")) & "  �������:" & vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "�������"))
        objPrint.BelowAppRows.Add objRow
    End If
    
    Set objPrint.Body = vsfDetail
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

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub


