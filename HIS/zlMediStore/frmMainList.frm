VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmMainList 
   ClientHeight    =   4980
   ClientLeft      =   2400
   ClientTop       =   4365
   ClientWidth     =   9480
   Icon            =   "frmMainList.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4980
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picColor 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   4680
      ScaleHeight     =   255
      ScaleWidth      =   3615
      TabIndex        =   14
      Top             =   4320
      Width           =   3615
      Begin VB.PictureBox picColor1 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   17
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
         TabIndex        =   16
         Top             =   0
         Width           =   260
      End
      Begin VB.PictureBox picColor3 
         BackColor       =   &H00FF00FF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2280
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   15
         Top             =   0
         Width           =   260
      End
      Begin VB.Label lblColor2 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   1680
         TabIndex        =   20
         Top             =   37
         Width           =   360
      End
      Begin VB.Label lblColor1 
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Height          =   180
         Left            =   360
         TabIndex        =   19
         Top             =   37
         Width           =   720
      End
      Begin VB.Label lblColor3 
         AutoSize        =   -1  'True
         Caption         =   "�������"
         Height          =   180
         Left            =   2640
         TabIndex        =   18
         Top             =   30
         Width           =   720
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfList 
      Height          =   1005
      Left            =   60
      TabIndex        =   11
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
      FormatString    =   $"frmMainList.frx":014A
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
      Height          =   330
      Left            =   60
      TabIndex        =   7
      Top             =   870
      Visible         =   0   'False
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   582
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "�Ƴ��ⷿ(&0)"
      TabPicture(0)   =   "frmMainList.frx":01BF
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "����ⷿ(&1)"
      TabPicture(1)   =   "frmMainList.frx":01DB
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
   End
   Begin VB.CommandButton Cmd���� 
      Caption         =   "����(&V)"
      Height          =   350
      Left            =   5250
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2700
      Width           =   1100
   End
   Begin VB.PictureBox picSeparate_s 
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   120
      MousePointer    =   7  'Size N S
      ScaleHeight     =   360
      ScaleWidth      =   4815
      TabIndex        =   4
      Top             =   2550
      Width           =   4815
      Begin VB.Label lbl4 
         AutoSize        =   -1  'True
         Caption         =   "���������"
         Height          =   180
         Left            =   4560
         TabIndex        =   12
         Top             =   0
         Width           =   1260
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         Caption         =   "�ɱ���"
         Height          =   180
         Left            =   0
         TabIndex        =   10
         Top             =   20
         Width           =   900
      End
      Begin VB.Label lbl2 
         AutoSize        =   -1  'True
         Caption         =   "�ۼ۽�"
         Height          =   180
         Left            =   1890
         TabIndex        =   9
         Top             =   20
         Width           =   900
      End
      Begin VB.Label lbl3 
         AutoSize        =   -1  'True
         Caption         =   "��۽�"
         Height          =   180
         Left            =   3690
         TabIndex        =   8
         Top             =   20
         Width           =   900
      End
      Begin VB.Label lblRange 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ѯ��Χ:1999��8��12����1999��9��12��"
         Height          =   180
         Left            =   0
         TabIndex        =   5
         Top             =   200
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
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ilsCold"
         HotImageList    =   "ilsHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   24
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
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Object.Visible         =   0   'False
                     Key             =   "FromStore"
                     Text            =   "��ⷿ��ҩ"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Object.Visible         =   0   'False
                     Key             =   "FromLeave"
                     Text            =   "��������ҩ"
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
               Key             =   "EditSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "�˲�"
               Key             =   "Prepare"
               Object.ToolTipText     =   "�˲�"
               Object.Tag             =   "�˲�"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "��ҩ"
               Key             =   "PreparePhysic"
               Object.ToolTipText     =   "��ҩ"
               Object.Tag             =   "��ҩ"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "����"
               Key             =   "SendPhysic"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "����"
               Key             =   "Back"
               Object.ToolTipText     =   "���˵��ϴ�״̬"
               Object.Tag             =   "����"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "PrepareSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "���"
               Key             =   "Verify"
               Description     =   "���"
               Object.ToolTipText     =   "���"
               Object.Tag             =   "���"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Strike"
               Description     =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "�������"
               Key             =   "ApplyStrike"
               Description     =   "�������"
               Object.ToolTipText     =   "�������"
               Object.Tag             =   "�������"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "��˳���"
               Key             =   "VerifyStrike"
               Description     =   "��˳���"
               Object.ToolTipText     =   "��˳���"
               Object.Tag             =   "��˳���"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "VerifySeparate"
               Style           =   3
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Search"
               Description     =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ˢ��"
               Key             =   "Refresh"
               Description     =   "ˢ��"
               Object.ToolTipText     =   "ˢ��"
               Object.Tag             =   "ˢ��"
               ImageIndex      =   13
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "PlugInSeparator"
               Style           =   3
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "����1"
               Key             =   "PlugItem"
               ImageIndex      =   17
            EndProperty
            BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "FindSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Description     =   "��������"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   14
            EndProperty
            BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Exit"
               Description     =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageIndex      =   15
            EndProperty
         EndProperty
         MouseIcon       =   "frmMainList.frx":01F7
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
            Picture         =   "frmMainList.frx":0511
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
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":0DA5
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":0FC5
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":11E5
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":1401
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":1621
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":1841
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":1F3B
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":292D
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":331F
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":3539
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":3755
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":3971
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":3B8B
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":3CE5
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":3F01
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":4121
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":A983
            Key             =   "PlugIn"
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
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":B85D
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":BA7D
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":BC9D
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":BEB9
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":C0D9
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":C2F9
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":C9F3
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":D3E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":DDD7
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":DFF1
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":E20D
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":E429
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":E643
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":E79D
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":E9BD
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":EBDD
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":1543F
            Key             =   "PlugIn"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfDetail 
      Height          =   975
      Left            =   0
      TabIndex        =   13
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
      FormatString    =   $"frmMainList.frx":17BF1
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
      Begin VB.Menu mnuFileCodePrint 
         Caption         =   "�����ӡ(&C)"
         Begin VB.Menu mnuFileAllCodePrint 
            Caption         =   "������ҩƷ�����ӡ(&A)"
         End
         Begin VB.Menu mnuFileSelCodePrint 
            Caption         =   "ѡ����ҩƷ�����ӡ(&S)"
         End
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
      Begin VB.Menu mnuEditPrepare 
         Caption         =   "�˲�(&H)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditPreparePhysic 
         Caption         =   "��ҩ(&P)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditSendPhysic 
         Caption         =   "����(&S)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditBack 
         Caption         =   "����(&O)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditLine3 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditVerify 
         Caption         =   "���(&C)"
      End
      Begin VB.Menu mnuEditMark 
         Caption         =   "��Ʊ�˶�(&M)"
      End
      Begin VB.Menu mnuEditStrike 
         Caption         =   "����(&K)"
      End
      Begin VB.Menu mnuEditApplyStrike 
         Caption         =   "�������"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditVerifyStrike 
         Caption         =   "��˳���"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditWriteOff 
         Caption         =   "��������"
      End
      Begin VB.Menu mnuEditLine0 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditRestore 
         Caption         =   "ҩ���˻�(&R)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditBill 
         Caption         =   "�޸ķ�Ʊ��Ϣ(&B)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditAcc 
         Caption         =   "�������(&V)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditPric 
         Caption         =   "�ɱ��۵���(&Z)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditHandBack 
         Caption         =   "ҩƷ��ҩ�ƻ�(&P)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditMediPlanImport 
         Caption         =   "ҩƷ�ƻ�����������(&I)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditDeliveryInvoice 
         Caption         =   "�ͻ���Ʊ����(&E)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditVerifySelect 
         Caption         =   "������˵��ݲ�ѯ(&Y)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditDisplay 
         Caption         =   "�鿴����(&W)"
      End
      Begin VB.Menu mnuEditCodePrintLine 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditAllCodePrint 
         Caption         =   "������ҩƷ�����ӡ(&A)"
         Visible         =   0   'False
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
Private mblnStock As Boolean            '��ǰ����Ա�Ƿ���ҩ����Ա���������õ�����Ч
Private mblnBootUp As Boolean
Private mlastRow As Long                '�ϴε������
Private mstrTitle As String             '����ı���
Private mbln�˲� As Boolean
Private mstrPrivs As String                     'Ȩ��
Private mintListRow As Integer
Private mStr�ⷿ As String
Private mblnDo As Boolean
Private mbln����Ա����  As Boolean
Private mblnViewCost As Boolean      '�鿴�ɱ��� true-���Բ鿴 false-�����Բ鿴

Private mstrNumberFormat As String
Private mstrCostFormat As String
Private mstrPriceFormat As String
Private mstrMoneyFormat As String

Private mlng�ⷿID As Long
Private mintUnit As Integer                 '��λϵ����1-�ۼ�;2-����;3-סԺ;4-ҩ��
Private mblnBandEvent As Boolean            '��¼�Ƿ��vsf�ؼ�������datasource������true-�ǣ�false-��

Private mobjPlugIn As Object             '��ҽӿ�

'�Ӳ�������ȡҩƷ�۸����������С��λ������ʾ���ȣ�
Private mintShowCostDigit As Integer            '�ɱ���С��λ��
Private mintShowPriceDigit As Integer           '�ۼ�С��λ��
Private mintShowNumberDigit As Integer          '����С��λ��
Private mintShowMoneyDigit As Integer           '���С��λ��

Private Const mconint�ۼ۵�λ As Integer = 1
Private Const mconint���ﵥλ As Integer = 2
Private Const mconintסԺ��λ As Integer = 3
Private Const mconintҩ�ⵥλ As Integer = 4

'��������
Private strStart As Date
Private strEnd As Date
Private strVerifyStart As Date
Private strVerifyEnd As Date

Private int�ƿ⴦������ As Integer                    '1-��Ҫ��ҩ�����͡�������һ����  0-����Ҫ��һ����
Private mint�������� As Integer                       '0-����Ҫ����;1-��Ҫ����
Private mint���ó������� As Integer                   'ҩƷ����ģ�������ʽ��0-����Ҫ����;1-��Ҫ����
Private mint����� As Integer                       '��ʾҩƷ����ʱ�Ƿ���п���飺0-�����;1-��飬�������ѣ�2-��飬�����ֹ

Private Type Type_SQLCondition
    strNO��ʼ As String
    strNO���� As String
    date����ʱ�俪ʼ As Date
    date����ʱ����� As Date
    date���ʱ�俪ʼ As Date
    date���ʱ����� As Date
    lngҩƷ As Long
    lng�ⷿ As Long
    str������ As String
    str����� As String
    lng������ As Long
    str���� As String
    str��Ʊ�ſ�ʼ As String
    str��Ʊ�Ž��� As String
    lng������ As Long
    int�������һ����ѯ As Integer
    int�ޱ�� As Integer
    int�б�� As Integer
    int�޷�Ʊ As Integer
    int�з�Ʊ As Integer
    lngҩƷ���� As Long
    str���� As String
    date��Ʊ������ڿ�ʼ As Date
    date��Ʊ������ڽ��� As Date
End Type

Private SQLCondition As Type_SQLCondition

Private mstr������ As String

Private Enum �⹺����
    NO = 0
    ��Ӧ�� = 1
    �ɱ���� = 2
    �ۼ۽�� = 3
    ��۽�� = 4
    ���۽�� = 5
    ���۲�� = 6
    �������� = 7
    ������ = 8
    �������� = 9
    �޸��� = 10
    �޸����� = 11
    �˲��� = 12
    �˲����� = 13
    ����� = 14
    ������� = 15
    ��¼״̬ = 16
    ����˵�� = 17
    ժҪ = 18
    
    ���� = 19
End Enum

Private Function Is����(ByVal strBillNo As String) As Boolean
    Dim rsCheck As New ADODB.Recordset
    '�ȼ���ǲ������쵥
    On Error GoTo errHandle
    gstrSQL = " Select Nvl(��ҩ��ʽ,0) ���� From ҩƷ�շ���¼ " & _
              " Where ����=6 And NO=[1] And ���ϵ�� = -1 and rownum = 1"
    Set rsCheck = zlDataBase.OpenSQLRecord(gstrSQL, "����ǲ������쵥", strBillNo)
    
    Is���� = Not (rsCheck!���� = 0)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub PlugInFun(ByVal strFunName As String)
    'ִ����ҹ���
    Dim strParam As String
    Dim lng�ⷿID As Long
    Dim int���� As Integer
    Dim strNo As String
    
    If mlngMode <> ģ���.�⹺��� Then Exit Sub
    
    With vsfList
        If .TextMatrix(.Row, 0) <> "" Then
            lng�ⷿID = Val(cboStock.ItemData(cboStock.ListIndex))
            If mlngMode = ģ���.�⹺��� Then int���� = 1
            strNo = .TextMatrix(.Row, 0)
            
            strParam = lng�ⷿID & "," & int���� & "," & strNo
        End If
    End With
    
    Call zlPlugIn_Fun(glngSys, mlngMode, mobjPlugIn, Me, strFunName, strParam)
End Sub

Private Sub SetDetailFocus()
    vsfDetail.ForeColorFixed = glngFixedForeColorByFocus
    vsfDetail.BackColorSel = glngRowByFocus
'    If vsfDetail.Row > 0 Then
'        vsfDetail.ForeColorSel = vsfDetail.Cell(flexcpForeColor, vsfDetail.Row)
'    End If
    
    vsfList.ForeColorFixed = glngFixedForeColorNotFocus
    vsfList.BackColorSel = glngRowByNotFocus
End Sub

Public Sub SetMenu()
    '���ر�ҩ�����͡���������
    mnuEditPreparePhysic.Visible = False
    mnuEditSendPhysic.Visible = False
    mnuEditBack.Visible = False
    tlbTool.Buttons("PreparePhysic").Visible = False
    tlbTool.Buttons("SendPhysic").Visible = False
    tlbTool.Buttons("Back").Visible = False
    mnuEditVerify.Visible = False
    mnuEditStrike.Visible = False
    mnuEditWriteOff.Visible = False
    tlbTool.Buttons("Verify").Visible = False
    tlbTool.Buttons("Strike").Visible = False
    mnuEditLine3.Visible = False
    mnuEditLine0.Visible = False
    tlbTool.Buttons("PrepareSeparate").Visible = False
    tlbTool.Buttons("VerifySeparate").Visible = False
    
    '���ݵ�ǰҳ�濪��
    If TabShow.Tab = 0 Then
        If mlngMode = ģ���.ҩƷ�ƿ� Then
            If int�ƿ⴦������ = 0 Then
                mnuEditPreparePhysic.Visible = False
                mnuEditVerify.Visible = zlStr.IsHavePrivs(mstrPrivs, "���")
'                mnuEditStrike.Visible = zlStr.IsHavePrivs(mstrPrivs, "����")
                mnuEditWriteOff.Visible = False
                mnuEditStrike.Visible = False
                mnuEditLine0.Visible = True
                tlbTool.Buttons("Verify").Visible = mnuEditVerify.Visible
                tlbTool.Buttons("Strike").Visible = mnuEditStrike.Visible
                tlbTool.Buttons("VerifySeparate").Visible = True
                mnuEditVerify.Caption = "���(&C)"
                tlbTool.Buttons("Verify").Caption = "���"
                tlbTool.Buttons("Verify").Tag = "���"
                tlbTool.Buttons("Verify").ToolTipText = "���"
            Else
                mnuEditVerify.Caption = "����(&C)"
                tlbTool.Buttons("Verify").Caption = "����"
                tlbTool.Buttons("Verify").Tag = "����"
                tlbTool.Buttons("Verify").ToolTipText = "����"
                mnuEditPreparePhysic.Visible = zlStr.IsHavePrivs(mstrPrivs, "����")
            End If
            If mint�������� = 1 Then
                mnuEditStrike.Visible = True
                mnuEditWriteOff.Visible = False
                mnuEditStrike.Caption = "��˳���(&K)"
                tlbTool.Buttons("Strike").Visible = True
                tlbTool.Buttons("Strike").Caption = IIf(mnuViewToolText.Checked = False, "", "��˳���")
                tlbTool.Buttons("Strike").Tag = "��˳���"
                tlbTool.Buttons("Strike").ToolTipText = "��˳���"
                mnuEditLine0.Visible = True
                tlbTool.Buttons("VerifySeparate").Visible = True
            End If
                        
            mnuEditSendPhysic.Visible = mnuEditPreparePhysic.Visible
            mnuEditBack.Visible = mnuEditPreparePhysic.Visible
            mnuEditLine3.Visible = mnuEditPreparePhysic.Visible
            tlbTool.Buttons("PreparePhysic").Visible = mnuEditPreparePhysic.Visible
            tlbTool.Buttons("SendPhysic").Visible = mnuEditPreparePhysic.Visible
            tlbTool.Buttons("Back").Visible = mnuEditPreparePhysic.Visible
            tlbTool.Buttons("PrepareSeparate").Visible = mnuEditPreparePhysic.Visible
        Else
            If mlngMode = ģ���.ҩƷ���� Then
                tlbTool.Buttons("Add").Style = tbrDropdown
                tlbTool.Buttons("Add").ButtonMenus(1).Visible = True
                tlbTool.Buttons("Add").ButtonMenus(2).Visible = True
    
                If mblnStock Then
                    mnuEditVerify.Visible = zlStr.IsHavePrivs(mstrPrivs, "���")
                    mnuEditStrike.Visible = zlStr.IsHavePrivs(mstrPrivs, "����")
                    mnuEditWriteOff.Visible = zlStr.IsHavePrivs(mstrPrivs, "����")
                    If mint���ó������� = 1 Then
                        mnuEditStrike.Visible = False
                        mnuEditApplyStrike.Visible = zlStr.IsHavePrivs(mstrPrivs, "��������")
                        mnuEditVerifyStrike.Visible = zlStr.IsHavePrivs(mstrPrivs, "�������")
                    End If
                End If
            Else
                mnuEditVerify.Visible = zlStr.IsHavePrivs(mstrPrivs, "���")
                mnuEditStrike.Visible = zlStr.IsHavePrivs(mstrPrivs, "����")
                mnuEditWriteOff.Visible = zlStr.IsHavePrivs(mstrPrivs, "����")
            End If
            
            If mlngMode = ģ���.�⹺��� Then
                mnuEditLine0.Visible = True
                If zlStr.IsHavePrivs(mstrPrivs, "�˲�ɱ���") And mbln�˲� Then
                    mnuEditPrepare.Visible = True
                    mnuEditBack.Visible = True
                    tlbTool.Buttons("Prepare").Visible = True
                    tlbTool.Buttons("Back").Visible = True
                    tlbTool.Buttons("PrepareSeparate").Visible = True
                End If
                
                If zlStr.IsHavePrivs(mstrPrivs, "������") And gtype_UserSysParms.P173_������Ǹ������ܽ��и������ = 1 Then
                    mnuEditMark.Visible = True
                Else
                    mnuEditMark.Visible = False
                End If
            End If
            tlbTool.Buttons("Verify").Visible = mnuEditVerify.Visible
            tlbTool.Buttons("Strike").Visible = mnuEditStrike.Visible
            tlbTool.Buttons("ApplyStrike").Visible = mnuEditApplyStrike.Visible
            tlbTool.Buttons("VerifyStrike").Visible = mnuEditVerifyStrike.Visible
            tlbTool.Buttons("VerifySeparate").Visible = (mnuEditVerify.Visible Or mnuEditStrike.Visible)
        End If
    Else
        mnuEditVerify.Visible = zlStr.IsHavePrivs(mstrPrivs, "���")
        mnuEditStrike.Visible = zlStr.IsHavePrivs(mstrPrivs, "����")
        mnuEditWriteOff.Visible = zlStr.IsHavePrivs(mstrPrivs, "����")
        
        If mlngMode = ģ���.ҩƷ�ƿ� Then
            If mint�������� = 1 Then
                mnuEditStrike.Caption = "�������(&R)"
                tlbTool.Buttons("Strike").Caption = IIf(mnuViewToolText.Checked = False, "", "�������")
                tlbTool.Buttons("Strike").Tag = "�������"
                tlbTool.Buttons("Strike").ToolTipText = "�������"
            Else
                mnuEditStrike.Caption = "����(&K)"
                tlbTool.Buttons("Strike").Caption = IIf(mnuViewToolText.Checked = False, "", "����")
                tlbTool.Buttons("Strike").ToolTipText = "����"
            End If
        End If
        
        If mlngMode = ģ���.�⹺��� Then
            mnuEditLine0.Visible = True
        End If
        tlbTool.Buttons("Verify").Visible = mnuEditVerify.Visible
        tlbTool.Buttons("Strike").Visible = mnuEditStrike.Visible
        tlbTool.Buttons("VerifySeparate").Visible = True
    End If
    
    If mlngMode = ģ���.������� Or mlngMode = ģ���.��۵��� Then
        mnuEditWriteOff.Visible = False
    End If
End Sub

Private Sub cboStock_Click()

    Dim lng�ⷿID As Long
    Dim rsCheck As New ADODB.Recordset
    
    On Error GoTo errHandle
    lng�ⷿID = cboStock.ItemData(cboStock.ListIndex)

    If mlng�ⷿID <> lng�ⷿID Then
        mlng�ⷿID = Me.cboStock.ItemData(Me.cboStock.ListIndex)
        Call GetDrugDigit(mlng�ⷿID, Me.Tag, mintUnit, mintShowCostDigit, mintShowPriceDigit, mintShowNumberDigit, mintShowMoneyDigit)
        
        '������֯��ʽ����
        mstrCostFormat = "'999999999990." & String(mintShowCostDigit, "0") & "'"
        mstrPriceFormat = "'999999999990." & String(mintShowPriceDigit, "0") & "'"
        mstrNumberFormat = "'999999999990." & String(mintShowNumberDigit, "0") & "'"
        mstrMoneyFormat = "'999999999990." & String(mintShowMoneyDigit, "0") & "'"
    
        '���ÿⷿ�Ƿ�Ϊҩ�⣬ֻ��ҩ��������˻�
        gstrSQL = " SELECT DISTINCT 0 " & _
                  " FROM ��������˵�� " & _
                  " WHERE �������� LIKE '%ҩ��' " & _
                  " AND ����ID = [1]"
        Set rsCheck = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[��鵱ǰ�ⷿ�Ƿ�Ϊҩ��]", lng�ⷿID)
                  
        mnuEditRestore.Enabled = (rsCheck.RecordCount > 0)
'        mnuEditLine0.Enabled = (rsCheck.RecordCount = 0)
        
        If mblnBootUp Then mnuViewRefresh_Click
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cboStock_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim str�������� As String
    '��ȡ�ɲ����Ŀⷿ
    Select Case mlngMode
        Case ģ���.�⹺���
            If InStr(1, mstrPrivs, "����ҩ���⹺���") = 0 Then
                str�������� = "H,I,J"
            Else
                str�������� = "H,I,J,K,L,M,N"
            End If
        Case ģ���.�������
            str�������� = "H,I,J,K,L,M,N"
        Case ģ���.�������
            str�������� = "H,I,J,K,L,M,N"
        Case ģ���.��۵���
            str�������� = "H,I,J,K,L,M,N"
        Case ģ���.ҩƷ�ƿ�
            str�������� = "H,I,J,K,L,M,N"
        Case ģ���.ҩƷ����
            str�������� = "H,I,J,K,L,M,N"
        Case ģ���.��������
            str�������� = "H,I,J,K,L,M,N"
        Case Else
    End Select
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If cboStock.ListCount = 0 Then Call zlControl.ControlSetFocus(vsfList): Exit Sub
    
    If cboStock.ListIndex >= 0 Then
        If Val(cboStock.Tag) = cboStock.ItemData(cboStock.ListIndex) Then
            Call zlControl.ControlSetFocus(vsfList, True)
            Exit Sub
        End If
    End If
    
    If Select����ѡ����(Me, cboStock, Trim(cboStock.Text), str��������, mbln����Ա����) = False Then
        Exit Sub
    End If
    If cboStock.ListIndex >= 0 Then
        cboStock.Tag = cboStock.ItemData(cboStock.ListIndex)
    End If
End Sub

Private Sub cboStock_Validate(Cancel As Boolean)
    If cboStock.ListCount > 0 Then
        If cboStock.ListIndex = -1 Then
            MsgBox "��ѡ��һ��ҩ�����ҩ����", vbInformation, gstrSysName
            Cancel = True
        End If
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
    Me.Tag = strTitle
    
    int�ƿ⴦������ = Val(zlDataBase.GetPara("�ƿ�����", glngSys, ģ���.ҩƷ�ƿ�))
    mint�������� = Val(zlDataBase.GetPara("��������", glngSys, ģ���.ҩƷ�ƿ�))
    mint���ó������� = Val(zlDataBase.GetPara("��������", glngSys, ģ���.ҩƷ����))
    
    If mlngMode = ģ���.�⹺��� Or mlngMode = ģ���.������� Then
        If zlDataBase.GetPara("ѡ����", glngSys, mlngMode) = "" Then
            mstr������ = "���ۼ�|���۵�λ|���۽��|���۲��"
        Else
            mstr������ = zlDataBase.GetPara("������", glngSys, mlngMode)
        End If
    End If
    
    '���������Բ���
    If Not CheckDepend Then
        Unload Me
        Exit Sub
    End If
    
    'ʵ�����ɹ�ƽ̨�ӿ�
    If mlngMode = ģ���.�⹺��� Then
        On Error Resume Next
        If gobjDrugPurchase Is Nothing Then
            Set gobjDrugPurchase = CreateObject("zlDrugPurchase.clsDrugPurchase")
        End If
        Err.Clear
        On Error GoTo 0
        If Not gobjDrugPurchase Is Nothing Then
            mnuEditDeliveryInvoice.Visible = True
        End If
    End If
    
    mlng�ⷿID = Me.cboStock.ItemData(Me.cboStock.ListIndex)
    Call GetDrugDigit(mlng�ⷿID, Me.Tag, mintUnit, mintShowCostDigit, mintShowPriceDigit, mintShowNumberDigit, mintShowMoneyDigit)
    
    '��֯��ʽ����
    mstrCostFormat = "'999999999990." & String(mintShowCostDigit, "0") & "'"
    mstrPriceFormat = "'999999999990." & String(mintShowPriceDigit, "0") & "'"
    mstrNumberFormat = "'999999999990." & String(mintShowNumberDigit, "0") & "'"
    mstrMoneyFormat = "'999999999990." & String(mintShowMoneyDigit, "0") & "'"
    
    SetVisable  '����Ȩ�����ò�ͬ����ʾ��Ŀ
        
    dateCurrentDate = Sys.Currentdate

    int��ѯ���� = Val(zlDataBase.GetPara("��ѯ����", glngSys, mlngMode, 7)) - 1
    strStart = Format(DateAdd("d", -int��ѯ����, dateCurrentDate), "yyyy-MM-dd")
    strEnd = Format(dateCurrentDate, "yyyy-MM-dd")
    
    strVerifyStart = "1901-01-01"
    strVerifyEnd = "1901-01-01"
    
    strFind = " AND A.��¼״̬ = 1 And A.������� is Null And A.�������� Between [3] And [4] "
    SQLCondition.date����ʱ�俪ʼ = CDate(Format(strStart, "yyyy-mm-dd") & " 00:00:00")
    SQLCondition.date����ʱ����� = CDate(Format(strEnd, "yyyy-mm-dd") & " 23:59:59")
    
    mstrFind = strFind
    
    Call TabShow_Click(0)
    If mlngMode <> ģ���.ҩƷ�ƿ� Then GetList (mstrFind) '�г�����ͷ
    
    '�ָ��û����Ի�����
    RestoreWinState Me, App.ProductName, mstrTitle
    
    If mlngMode = ģ���.��۵��� Then
        vsfList.ColWidth(vsfList.Cols - 1) = 0
        vsfList.ColWidth(vsfList.Cols - 3) = 0
    End If
    If mlngMode = ģ���.�⹺��� Then
        vsfList.ColWidth(�⹺����.��¼״̬) = 0
        vsfList.ColWidth(�⹺����.����˵��) = 1000
    End If
    If mlngMode = ģ���.ҩƷ���� Then
        vsfList.ColWidth(vsfList.Cols - 4) = 0
        vsfList.ColWidth(vsfList.Cols - 3) = 1000
    End If
    '�û����Ի����ú���������Ȩ�޿��Ƶ����Ƿ���ʾ
    If mblnViewCost = False Then
        With vsfList
            Select Case mlngMode
                Case ģ���.�⹺���
                    .colHidden(�⹺����.�ɱ����) = True
                    .colHidden(�⹺����.��۽��) = True
                Case ģ���.�������
                    .colHidden(.ColIndex("�ɱ����")) = True
                    .colHidden(.ColIndex("��۽��")) = True
                Case ģ���.�������
                    .colHidden(.ColIndex("�ɱ����")) = True
                    .colHidden(.ColIndex("��۽��")) = True
                Case ģ���.ҩƷ�ƿ�
                    .colHidden(.ColIndex("�ɱ����")) = True
                    .colHidden(.ColIndex("��۽��")) = True
                Case ģ���.ҩƷ����
                    .colHidden(.ColIndex("�ɱ����")) = True
                    .colHidden(.ColIndex("��۽��")) = True
                Case ģ���.��������
                    .colHidden(.ColIndex("�ɱ����")) = True
                    .colHidden(.ColIndex("��۽��")) = True
            End Select
        End With
        
        With vsfDetail
            Select Case mlngMode
                Case ģ���.�⹺���
                    .colHidden(.ColIndex("�ɱ���")) = True
                    .colHidden(.ColIndex("�ɱ����")) = True
                    .colHidden(.ColIndex("���")) = True
                Case ģ���.�������
                    .colHidden(.ColIndex("�ɱ���")) = True
                    .colHidden(.ColIndex("�ɱ����")) = True
                    .colHidden(.ColIndex("���")) = True
                Case ģ���.�������
                    .colHidden(.ColIndex("�ɹ���")) = True
                    .colHidden(.ColIndex("�ɹ����")) = True
                    .colHidden(.ColIndex("���")) = True
                Case ģ���.ҩƷ�ƿ�
                    .colHidden(.ColIndex("�ɱ���")) = True
                    .colHidden(.ColIndex("�ɱ����")) = True
                    .colHidden(.ColIndex("���")) = True
                Case ģ���.ҩƷ����
                    .colHidden(.ColIndex("�ɱ���")) = True
                    .colHidden(.ColIndex("�ɱ����")) = True
                    .colHidden(.ColIndex("���")) = True
                Case ģ���.��������
                    .colHidden(.ColIndex("�ɱ���")) = True
                    .colHidden(.ColIndex("�ɱ����")) = True
                    .colHidden(.ColIndex("���")) = True
            End Select
        End With
    End If
            
    Call zlDataBase.ShowReportMenu(Me, glngSys, mlngMode, mstrPrivs)
    
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
    Dim rsDepend As New ADODB.Recordset
    Dim strStock As String, strCaption As String
    
    CheckDepend = False
    On Error GoTo errHandle
    
    '��ȡ�ɲ����Ŀⷿ
    Select Case mlngMode
        Case ģ���.�⹺���
            If InStr(1, mstrPrivs, "����ҩ���⹺���") = 0 Then
                strStock = "HIJ"
            Else
                strStock = "HIJKLMN"
            End If
        Case ģ���.�������
            strStock = "HIJKLMN"
        Case ģ���.�������
            strStock = "HIJKLMN"
        Case ģ���.��۵���
            strStock = "HIJKLMN"
        Case ģ���.ҩƷ�ƿ�
            strStock = "HIJKLMN"
        Case ģ���.ҩƷ����
            strStock = "HIJKLMN"
        Case ģ���.��������
            strStock = "HIJKLMN"
        Case Else
    End Select
    
    '�����ҩƷ���ã����鵱ǰ�����Ƿ������ò��ţ���������ⷿ��ҩ
    If mlngMode <> ģ���.ҩƷ���� Then
        gstrSQL = "SELECT DISTINCT a.id, a.���� " _
                & "FROM ��������˵�� c, �������ʷ��� b, ���ű� a " _
                & "Where (a.վ�� = [3] Or a.վ�� is Null) And c.�������� = b.���� " _
                & "  AND Instr([2],b.����,1) > 0 " _
                & "  AND a.id = c.����id " _
                & "  AND TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01'" _
                & IIf(zlStr.IsHavePrivs(gstrprivs, "���пⷿ"), "", " And a.ID IN (Select ����ID From ������Ա Where ��ԱID=[1])")
        Set rsDepend = zlDataBase.OpenSQLRecord(gstrSQL, mstrTitle, UserInfo.�û�ID, strStock, gstrNodeNo)
        
    Else
        '���ж��ǲ���ҩ����Աʹ�ñ�ģ��
        gstrSQL = "SELECT DISTINCT a.id, a.���� " _
                & "FROM ��������˵�� c, �������ʷ��� b, ���ű� a " _
                & "Where (a.վ�� = [3] Or a.վ�� is Null) And c.�������� = b.���� " _
                & "  AND Instr([2],b.����,1) > 0 " _
                & "  AND a.id = c.����id " _
                & "  AND TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01'" _
                & "  And a.ID IN (Select ����ID From ������Ա Where ��ԱID=[1])"
        Set rsDepend = zlDataBase.OpenSQLRecord(gstrSQL, mstrTitle, UserInfo.�û�ID, strStock, gstrNodeNo)
                  
        mblnStock = (rsDepend.RecordCount <> 0)
        
        If mblnStock Then
            '����Ȩ�����пⷿ��ȡ�ⷿ����
            gstrSQL = "SELECT DISTINCT a.id, a.���� " _
                    & "FROM ��������˵�� c, �������ʷ��� b, ���ű� a " _
                    & "Where (a.վ�� = [3] Or a.վ�� is Null) And c.�������� = b.���� " _
                    & "  AND Instr([2],b.����,1) > 0 " _
                    & "  AND a.id = c.����id " _
                    & "  AND TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01'" _
                    & IIf(zlStr.IsHavePrivs(gstrprivs, "���пⷿ"), "", " And a.ID IN (Select ����ID From ������Ա Where ��ԱID=[1])")
            Set rsDepend = zlDataBase.OpenSQLRecord(gstrSQL, mstrTitle, UserInfo.�û�ID, strStock, gstrNodeNo)
        Else
            '��ȡ����Ա���������ò���
            gstrSQL = " Select C.ID " & _
                      " From ��������˵�� A,�������ʷ��� B,���ű� C " & _
                      " Where (c.վ�� = [3] Or c.վ�� is Null) And A.��������=B.���� And A.����ID=C.ID " & _
                      "   AND TO_CHAR(C.����ʱ��, 'yyyy-MM-dd')='3000-01-01' And B.����='O'" & _
                      "   And C.ID IN (Select ����ID From ������Ա Where ��ԱID=[1])"
            Set rsDepend = zlDataBase.OpenSQLRecord(gstrSQL, mstrTitle, UserInfo.�û�ID, "O", gstrNodeNo)
                      
            If rsDepend.RecordCount = 0 Then
                MsgBox "�㲻����ҩ���ŵĲ�����Ա������ʹ�ñ�ģ�飡[���Ź���]", vbInformation, gstrSysName
                Exit Function
            End If
            
            '�ٸ���ҩƷ��ҩ���ƣ���ȡ��Щ��ҩ�����������ÿⷿ������
            gstrSQL = "SELECT DISTINCT a.id, a.���� " _
                    & "FROM ��������˵�� c, �������ʷ��� b, ���ű� a " _
                    & "Where (a.վ�� = [3] Or a.վ�� is Null) And c.�������� = b.���� " _
                    & "  AND Instr([2],b.����,1) > 0 " _
                    & "  AND a.id = c.����id " _
                    & "  AND TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01'" _
                    & "  And a.ID IN (Select �Է��ⷿID From ҩƷ���ÿ��� Where ���ò���ID IN (" & gstrSQL & "))"
            Set rsDepend = zlDataBase.OpenSQLRecord(gstrSQL, mstrTitle, UserInfo.�û�ID, strStock, gstrNodeNo)
        End If
    End If
        
    If rsDepend.EOF Then
        If mlngMode <> ģ���.ҩƷ���� Or mblnStock Then
            If mlngMode = ģ���.�⹺��� And InStr(1, mstrPrivs, "����ҩ���⹺���") = 0 Then
                MsgBox "����Ա�ޡ�����ҩ���⹺��⡱Ȩ�ޣ��������Ա��ϵ��", vbInformation, gstrSysName
            Else
                MsgBox "����Ӧ������һ������ҩ�����ʣ�ҩ�����ʣ������Ƽ������ʵĲ���,��鿴���Ź���", vbInformation, gstrSysName
            End If
        Else
            MsgBox "��û��Ȩ�����κοⷿ����ҩƷ����������������[������������]", vbInformation, gstrSysName
        End If
        If rsDepend.State = 1 Then rsDepend.Close
        Exit Function
    End If
    
    'װ��ⷿ����
    With cboStock
        .Clear
        Do While Not rsDepend.EOF
            .AddItem rsDepend!����
            .ItemData(.NewIndex) = rsDepend!id
            mStr�ⷿ = mStr�ⷿ & rsDepend!id & "," & rsDepend!���� & "|"
            If mlngMode <> ģ���.ҩƷ���� Or mblnStock Then
                If rsDepend!id = UserInfo.����ID Then
                    .ListIndex = .NewIndex
                End If
            End If
            rsDepend.MoveNext
        Loop
        If .ListIndex = -1 Then .ListIndex = 0
        rsDepend.Close
    End With
    
    '����Ƿ���Ҫ�˲黷�ڣ�������⹺��⣩
    strCaption = "�˲�"
    If mlngMode = ģ���.�⹺��� Then
        mbln�˲� = (gtype_UserSysParms.P75_�⹺�����Ҫ�˲� = 1)
                
        mnuEditPrepare.Caption = strCaption & "(&H)"
        tlbTool.Buttons("Prepare").Caption = strCaption
        tlbTool.Buttons("Prepare").Tag = strCaption
        tlbTool.Buttons("Prepare").ToolTipText = strCaption
    ElseIf mlngMode = ģ���.ҩƷ�ƿ� Then
        If int�ƿ⴦������ = 0 Then
            mnuEditVerify.Caption = "���(&C)"
            tlbTool.Buttons("Verify").Caption = "���"
            tlbTool.Buttons("Verify").Tag = "���"
            tlbTool.Buttons("Verify").ToolTipText = "���"
        Else
            mnuEditVerify.Caption = "����(&C)"
            tlbTool.Buttons("Verify").Caption = "����"
            tlbTool.Buttons("Verify").Tag = "����"
            tlbTool.Buttons("Verify").ToolTipText = "����"
        End If
        TabShow.Visible = True
    End If
    
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
    Dim strTransfer As String, strTransfer_Order As String
    Dim strsql As String
    Dim strSqlForm As String
    
    '����ͳ�ƺϼƽ��
    Dim dbl1 As Double
    Dim dbl2 As Double
    Dim dbl3 As Double
    Dim dbl4 As Double
    Dim n As Long
    Dim StrFormat As String
    
    StrFormat = "0.00##"
    On Error GoTo errHandle
    mlastRow = 0
    Call FS.ShowFlash("��������ҩƷ��¼,���Ժ� ...", Me)
    DoEvents
    Screen.MousePointer = vbHourglass
    
    If strFind = "" Then
        strFind = " AND A.��¼״̬ = 1 And A.������� is Null And A.�������� Between [3] And [4] "
        SQLCondition.date����ʱ�俪ʼ = CDate(Format(strStart, "yyyy-mm-dd") & " 00:00:00")
        SQLCondition.date����ʱ����� = CDate(Format(strEnd, "yyyy-mm-dd") & " 23:59:59")
    End If
    
    Select Case mlngMode
    Case ģ���.ҩƷ�ƿ�
        If TabShow.Tab = 0 Then
            strUserPart = " And A.�ⷿID+0=[16] "
            strTransfer = " A.����� AS ������,TO_CHAR(MIN(A.�������),'YYYY-MM-DD HH24:MI:SS') AS ��������,A.��ҩ�� AS ��ҩ��,TO_CHAR(MIN(A.��ҩ����),'YYYY-MM-DD HH24:MI:SS') AS ��������"
            strTransfer_Order = ",A.�����,A.��ҩ��"
        Else
            strUserPart = " And A.�Է�����ID+0=[16] "
            strTransfer = " A.��ҩ�� AS ��ҩ��,TO_CHAR(MIN(A.��ҩ����),'YYYY-MM-DD HH24:MI:SS') AS ��������,A.����� AS ������,TO_CHAR(MIN(A.�������),'YYYY-MM-DD HH24:MI:SS') AS ��������"
            strTransfer_Order = ",A.��ҩ��,A.�����"
        End If
    Case ģ���.ҩƷ����
        If mblnStock Then
            strUserPart = " And A.�ⷿID+0=[16] "
        Else
            strUserPart = " Select C.ID " & _
                      " From ��������˵�� A,�������ʷ��� B,���ű� C " & _
                      " Where (c.վ�� = [18] Or c.վ�� is Null) And A.��������=B.���� And A.����ID=C.ID " & _
                      " AND TO_CHAR(C.����ʱ��, 'yyyy-MM-dd')='3000-01-01' And B.����='O'" & _
                      " And C.ID IN (Select ����ID From ������Ա Where ��ԱID=[17])"

            strUserPart = " And A.�ⷿID+0=[16] And A.�Է�����ID+0 IN (" & strUserPart & ")"
        End If
    Case Else
        strUserPart = " And A.�ⷿID+0=[16] "
    End Select
    
    If mlngMode = ģ���.��۵��� Or mlngMode = ģ���.ҩƷ�ƿ� Or mlngMode = ģ���.ҩƷ���� Or mlngMode = ģ���.�������� Or mlngMode = ģ���.�⹺��� Then
        If SQLCondition.str���� <> "" And SQLCondition.lngҩƷ���� = 0 Then
            strSqlForm = " , ҩƷ��� F, ������ĿĿ¼ G, ҩƷ���� H"
            strFind = strFind & " and a.ҩƷid = f.ҩƷid And f.ҩ��id = g.Id And g.Id = h.ҩ��id(+) and h.ҩƷ���� in (select * from Table(Cast(f_Str2list([21]) As zlTools.t_Strlist))) and (g.���='5' or g.���='6' or g.���='7')"
        ElseIf SQLCondition.str���� = "" And SQLCondition.lngҩƷ���� <> 0 Then
            strSqlForm = " , ҩƷ��� F, ������ĿĿ¼ G"
            strFind = strFind & " and a.ҩƷid = f.ҩƷid And f.ҩ��id = g.Id And g.����id + 0=[22] and (g.���='5' or g.���='6' or g.���='7')"
        ElseIf SQLCondition.str���� <> "" And SQLCondition.lngҩƷ���� <> 0 Then
            strSqlForm = " , ҩƷ��� F, ������ĿĿ¼ G, ҩƷ���� H"
            strFind = strFind & " and a.ҩƷid = f.ҩƷid And f.ҩ��id = g.Id And g.Id = h.ҩ��id(+) and h.ҩƷ���� in (select * from Table(Cast(f_Str2list([21]) As zlTools.t_Strlist))) and (g.���='5' or g.���='6' or g.���='7') and g.����id + 0=[22]"
        End If
    End If

    vsfList.Redraw = flexRDNone
    Select Case mlngMode
        Case ģ���.�⹺���           'ҩƷ�⹺������
            
            If SQLCondition.int�������һ����ѯ = 0 Then
                gstrSQL = "SELECT A.NO, C.���� AS ��Ӧ��,LTRIM(TO_CHAR(SUM(A.�ɱ����)," & mstrMoneyFormat & ")) AS �ɱ����," & _
                    " LTRIM(TO_CHAR(SUM(A.���۽��)," & mstrMoneyFormat & " )) AS �ۼ۽��, LTRIM(TO_CHAR(SUM(A.���)," & mstrMoneyFormat & " )) AS ��۽��, " & _
                    " LTRIM(TO_CHAR(SUM(A.���۽��)," & mstrMoneyFormat & " )) AS ���۽��, LTRIM(TO_CHAR(SUM(A.���۽�� - A.�ɱ����)," & mstrMoneyFormat & " )) AS ���۲��, Nvl(A.����id, 0) ��������,A.������," & _
                    " TO_CHAR(MIN(A.��������), 'YYYY-MM-DD HH24:MI:SS') AS ��������,A.�޸���,TO_CHAR(MIN(A.�޸�����), 'YYYY-MM-DD HH24:MI:SS') AS �޸�����,A.��ҩ�� As �˲���,TO_CHAR(MIN(A.��ҩ����), 'YYYY-MM-DD HH24:MI:SS') As �˲�����,A.�����," & _
                    " TO_CHAR(MIN(A.�������), 'YYYY-MM-DD HH24:MI:SS') AS �������,A.��¼״̬,Decode(Nvl(A.��ҩ��ʽ, 0), 0, '��ⵥ', '�˿ⵥ') ����˵��,A.ժҪ " & _
                    " FROM ҩƷ�շ���¼ A, ���ű� B,��Ӧ�� C,Ӧ����¼ D " & strSqlForm & _
                    " WHERE A.�ⷿID + 0 = B.ID AND A.��ҩ��λID+0 = C.ID AND SUBSTR(C.����,1,1)=1 AND D.ϵͳ��ʶ(+)=1 AND D.��¼����(+)=0 " & _
                    " AND A.ID=D.�շ�ID(+) AND A.���� = 1 " & strUserPart & strFind & _
                    " GROUP BY A.NO,C.����,Nvl(A.����id, 0),A.������,A.�޸���,A.��ҩ��,A.��ҩ����,A.�����,A.��¼״̬, A.��ҩ��ʽ,A.ժҪ " & _
                    " ORDER BY NO DESC,�������� ASC"
            Else
                gstrSQL = "SELECT A.NO, C.���� AS ��Ӧ��,LTRIM(TO_CHAR(SUM(A.�ɱ����)," & mstrMoneyFormat & ")) AS �ɱ����," & _
                    " LTRIM(TO_CHAR(SUM(A.���۽��)," & mstrMoneyFormat & " )) AS �ۼ۽��, LTRIM(TO_CHAR(SUM(A.���)," & mstrMoneyFormat & " )) AS ��۽��, " & _
                    " LTRIM(TO_CHAR(SUM(A.���۽��)," & mstrMoneyFormat & " )) AS ���۽��, LTRIM(TO_CHAR(SUM(A.���۽�� - A.�ɱ����)," & mstrMoneyFormat & " )) AS ���۲��, Nvl(A.����id, 0) ��������,A.������," & _
                    " TO_CHAR(MIN(A.��������), 'YYYY-MM-DD HH24:MI:SS') AS ��������,A.�޸���,TO_CHAR(MIN(A.�޸�����), 'YYYY-MM-DD HH24:MI:SS') AS �޸�����,A.��ҩ�� As �˲���,TO_CHAR(MIN(A.��ҩ����), 'YYYY-MM-DD HH24:MI:SS') As �˲�����,A.�����," & _
                    " TO_CHAR(MIN(A.�������), 'YYYY-MM-DD HH24:MI:SS') AS �������,A.��¼״̬,Decode(Nvl(A.��ҩ��ʽ, 0), 0, '��ⵥ', '�˿ⵥ') ����˵��,A.ժҪ " & _
                    " FROM ҩƷ�շ���¼ A, ���ű� B,��Ӧ�� C,Ӧ����¼ D " & strSqlForm & _
                    " WHERE A.�ⷿID + 0 = B.ID AND A.��ҩ��λID+0 = C.ID AND SUBSTR(C.����,1,1)=1 AND D.ϵͳ��ʶ(+)=1 AND D.��¼����(+)=0 " & _
                    " AND A.ID=D.�շ�ID(+) AND A.���� = 1 " & strUserPart & strFind & " And (A.�������� Between [3] And [4]) And A.������� Is Null " & _
                    " GROUP BY A.NO,C.����,Nvl(A.����id, 0),A.������,A.�޸���,A.��ҩ��,A.��ҩ����,A.�����,A.��¼״̬,A.��ҩ��ʽ,A.ժҪ " & _
                    " Union All " & _
                    " SELECT A.NO, C.���� AS ��Ӧ��,LTRIM(TO_CHAR(SUM(A.�ɱ����)," & mstrMoneyFormat & ")) AS �ɱ����," & _
                    " LTRIM(TO_CHAR(SUM(A.���۽��)," & mstrMoneyFormat & " )) AS �ۼ۽��, LTRIM(TO_CHAR(SUM(A.���)," & mstrMoneyFormat & " )) AS ��۽��, " & _
                    " LTRIM(TO_CHAR(SUM(A.���۽��)," & mstrMoneyFormat & " )) AS ���۽��, LTRIM(TO_CHAR(SUM(A.���۽�� - A.�ɱ����)," & mstrMoneyFormat & " )) AS ���۲��, Nvl(A.����id, 0) ��������,A.������," & _
                    " TO_CHAR(MIN(A.��������), 'YYYY-MM-DD HH24:MI:SS') AS ��������,A.�޸���,TO_CHAR(MIN(A.�޸�����), 'YYYY-MM-DD HH24:MI:SS') AS �޸�����,A.��ҩ�� As �˲���,TO_CHAR(MIN(A.��ҩ����), 'YYYY-MM-DD HH24:MI:SS') As �˲�����,A.�����," & _
                    " TO_CHAR(MIN(A.�������), 'YYYY-MM-DD HH24:MI:SS') AS �������,A.��¼״̬,Decode(Nvl(A.��ҩ��ʽ, 0), 0, '��ⵥ', '�˿ⵥ') ����˵��,A.ժҪ " & _
                    " FROM ҩƷ�շ���¼ A, ���ű� B,��Ӧ�� C,Ӧ����¼ D " & strSqlForm & _
                    " WHERE A.�ⷿID + 0 = B.ID AND A.��ҩ��λID+0 = C.ID AND SUBSTR(C.����,1,1)=1 AND D.ϵͳ��ʶ(+)=1 AND D.��¼����(+)=0 " & _
                    " AND A.ID=D.�շ�ID(+) AND A.���� = 1 " & strUserPart & strFind & " And (A.������� Between [5] And [6]) " & _
                    " GROUP BY A.NO,C.����,Nvl(A.����id, 0),A.������,A.�޸���,A.��ҩ��,A.��ҩ����,A.�����,A.��¼״̬,A.��ҩ��ʽ,A.ժҪ " & _
                    " ORDER BY NO DESC,�������� ASC"
            End If
        Case ģ���.�������           'ҩƷ����������
            gstrSQL = "SELECT A.NO, C.���� AS �Ƽ���,LTRIM(TO_CHAR (SUM (A.�ɱ����), " & mstrMoneyFormat & ")) AS �ɱ����," & _
                " LTRIM(TO_CHAR ( (SUM (A.���۽��)), " & mstrMoneyFormat & ")) AS �ۼ۽��,  LTRIM(TO_CHAR((SUM(A.���۽�� - A.�ɱ����))," & mstrMoneyFormat & " )) AS ��۽��, A.������, " & _
                " TO_CHAR (MIN(A.��������), 'YYYY-MM-DD HH24:MI:SS') AS ��������, A.�����, " & _
                " TO_CHAR (MIN(A.�������), 'YYYY-MM-DD HH24:MI:SS') AS �������, A.��¼״̬, A.ժҪ " & _
                " FROM ҩƷ�շ���¼ A, ���ű� B ,���ű� C " & _
                " WHERE A.�ⷿID = B.ID AND A.�Է�����ID=C.ID AND A.���� = 2 AND A.���ϵ��=1 " & _
                strUserPart & strFind & _
                " GROUP BY A.NO,C.����,A.������,A.�����,A.��¼״̬,A.ժҪ " & _
                " ORDER BY NO DESC, �������� ASC "
    
        Case ģ���.�������           'ҩƷ����������
'            gstrSQL = "SELECT /*+ Rule*/ A.NO, C.���� AS ������,LTRIM(TO_CHAR (SUM (A.�ɱ����), " & mstrMoneyFormat & ")) AS �ɱ����," & _
'                " LTRIM(TO_CHAR (SUM (A.���۽��)-Sum(To_Number(Nvl(A.�÷�, 0))), " & mstrMoneyFormat & ")) AS �ۼ۽��,LTRIM(TO_CHAR(SUM(A.���۽�� - A.�ɱ����- To_Number(Nvl(A.�÷�, 0)))," & mstrMoneyFormat & " )) AS ��۽��, " & _
'                " LTRIM(TO_CHAR (SUM (A.���۽��), " & mstrMoneyFormat & ")) AS ���۽��,LTRIM(TO_CHAR(SUM(A.���۽�� - A.�ɱ����)," & mstrMoneyFormat & " )) AS ���۲��, " & _
'                " A.������,TO_CHAR (MIN(A.��������), 'YYYY-MM-DD HH24:MI:SS') AS ��������, A.�����," & _
'                " TO_CHAR (MIN(A.�������), 'YYYY-MM-DD HH24:MI:SS') AS �������, A.��¼״̬, A.ժҪ " & _
'                " FROM ҩƷ�շ���¼ A, ���ű� B,ҩƷ������ C " & _
'                " WHERE A.�ⷿID = B.ID AND A.������ID = C.ID AND A.���� = 4 " & _
'                strUserPart & StrFind & _
'                " GROUP BY A.NO,C.����,A.������,A.�����,A.��¼״̬,A.ժҪ " & _
'                " ORDER BY NO DESC,�������� ASC "
            gstrSQL = "SELECT  A.NO, C.���� AS ������,LTRIM(TO_CHAR (SUM (A.�ɱ����), " & mstrMoneyFormat & ")) AS �ɱ����," & _
                " LTRIM(TO_CHAR (SUM (A.���۽��), " & mstrMoneyFormat & ")) AS �ۼ۽��,LTRIM(TO_CHAR(SUM(A.���)," & mstrMoneyFormat & " )) AS ��۽��, " & _
                " LTRIM(TO_CHAR (SUM (A.���۽��), " & mstrMoneyFormat & ")) AS ���۽��,LTRIM(TO_CHAR(SUM(A.���۽�� - A.�ɱ����)," & mstrMoneyFormat & " )) AS ���۲��, " & _
                " A.������,TO_CHAR (MIN(A.��������), 'YYYY-MM-DD HH24:MI:SS') AS ��������,A.�޸���,TO_CHAR (MIN(A.�޸�����), 'YYYY-MM-DD HH24:MI:SS') AS �޸�����, A.�����," & _
                " TO_CHAR (MIN(A.�������), 'YYYY-MM-DD HH24:MI:SS') AS �������, A.��¼״̬, A.ժҪ " & _
                " FROM ҩƷ�շ���¼ A, ���ű� B,ҩƷ������ C " & _
                " WHERE A.�ⷿID = B.ID AND A.������ID = C.ID AND A.���� = 4 " & _
                strUserPart & strFind & _
                " GROUP BY A.NO,C.����,A.������,A.�޸���,A.�����,A.��¼״̬,A.ժҪ " & _
                " ORDER BY NO DESC,�������� ASC "
        Case ģ���.��۵���           '����۵�������
            gstrSQL = "SELECT  A.NO, LTRIM(TO_CHAR (SUM (A.���ۼ�), " & mstrMoneyFormat & ")) AS �����,LTRIM(TO_CHAR (SUM (A.�ɱ���),'9999999999999990.00000')) AS �����," & _
                " LTRIM(TO_CHAR ( (SUM (A.���)), '9999999999999990.00000')) AS ������, A.������,TO_CHAR (MIN(A.��������), 'YYYY-MM-DD HH24:MI:SS') AS ��������, A.�����," & _
                " TO_CHAR (MIN(A.�������), 'YYYY-MM-DD HH24:MI:SS') AS �������, A.��¼״̬, A.ժҪ,0 ��ҩ��ʽ " & _
                " FROM ҩƷ�շ���¼ A, ���ű� B  " & strSqlForm & _
                " WHERE A.�ⷿID = B.ID  AND A.���� = 5 And nvl(��ҩ��ʽ,0)=0 " & _
                strUserPart & strFind & _
                " GROUP BY A.NO,A.������,A.�����,A.��¼״̬,A.ժҪ " & _
                "UNION ALL " & _
                "SELECT A.NO, LTRIM(TO_CHAR (SUM (A.���ۼ�), " & mstrMoneyFormat & ")) AS �����,LTRIM(TO_CHAR (SUM (A.�ɱ���)," & mstrMoneyFormat & ")) AS �����," & _
                " LTRIM(TO_CHAR ( (SUM (A.���)), " & mstrMoneyFormat & ")) AS ������, A.������,TO_CHAR (MIN(A.��������), 'YYYY-MM-DD HH24:MI:SS') AS ��������, A.�����," & _
                " TO_CHAR (MIN(A.�������), 'YYYY-MM-DD HH24:MI:SS') AS �������, A.��¼״̬, A.ժҪ,1 ��ҩ��ʽ " & _
                " FROM ҩƷ�շ���¼ A, ���ű� B  " & strSqlForm & _
                " WHERE A.�ⷿID = B.ID  AND A.���� = 5 And  nvl(��ҩ��ʽ,0)=1  and a.�ⷿid=" & cboStock.ItemData(cboStock.ListIndex) & _
                strFind & _
                " GROUP BY A.NO,A.������,A.�����,A.��¼״̬,A.ժҪ " & _
                " ORDER BY NO DESC,�������� ASC "
        Case ģ���.ҩƷ�ƿ�           'ҩƷ�ƿ�������ҵĳ��⣬�������ޣ����ҵ���⣬��ֻ�ܿ���δ��ҩ���ѷ��͵ĵ��ݣ�
            gstrSQL = "SELECT A.NO," & IIf(TabShow.Tab = 0, "C.���� As ����ⷿ,", "B.���� AS �Ƴ��ⷿ,") & " LTRIM(TO_CHAR (SUM (A.�ɱ����), " & mstrMoneyFormat & ")) AS �ɱ����, " & _
                " LTRIM(TO_CHAR ( (SUM (A.���۽��)), " & mstrMoneyFormat & ")) AS �ۼ۽��, LTRIM(TO_CHAR((SUM(A.���۽�� - A.�ɱ����))," & mstrMoneyFormat & " )) AS ��۽��,A.������, " & _
                " TO_CHAR (MIN(A.��������), 'YYYY-MM-DD HH24:MI:SS') AS ��������,A.�޸���,TO_CHAR (MIN(A.�޸�����), 'YYYY-MM-DD HH24:MI:SS') AS �޸�����," & strTransfer & " ,A.��¼״̬, A.ժҪ " & _
                " FROM ҩƷ�շ���¼ A, ���ű� B ,���ű� C  " & strSqlForm & _
                " WHERE A.�ⷿID = B.ID AND A.�Է�����ID=C.ID AND A.���� = 6 AND  A.���ϵ��=-1" & _
                IIf(TabShow.Tab = 0, " ", " And (A.��ҩ�� Is NULL Or A.��ҩ���� Is Not NULL)") & _
                strUserPart & strFind & _
                " GROUP BY A.NO," & IIf(TabShow.Tab = 0, "C.����", "B.����") & ",A.������,A.�޸���" & strTransfer_Order & ",A.��¼״̬,A.ժҪ " & _
                " ORDER BY NO DESC, �������� ASC "
        Case ģ���.ҩƷ����           'ҩƷ���ù���
            gstrSQL = "SELECT A.NO, C.���� AS ���ò���,LTRIM(TO_CHAR (SUM (A.�ɱ����), " & mstrMoneyFormat & ")) AS �ɱ����, " & _
                " LTRIM(TO_CHAR ( (SUM (A.���۽��)), " & mstrMoneyFormat & ")) AS �ۼ۽��, LTRIM(TO_CHAR((SUM(A.���۽�� - A.�ɱ����))," & mstrMoneyFormat & " )) AS ��۽��,A.������, " & _
                " TO_CHAR (MIN(A.��������), 'YYYY-MM-DD HH24:MI:SS') AS ��������, A.�޸���,TO_CHAR (MIN(A.�޸�����), 'YYYY-MM-DD HH24:MI:SS') AS �޸�����, A.�����, " & _
                " TO_CHAR (MIN(A.�������), 'YYYY-MM-DD HH24:MI:SS') AS �������, A.��¼״̬, A.ժҪ,Decode(Nvl(A.��ҩ��ʽ,0),0,'��ⷿ��ҩ','��������ҩ') ��ҩ��ʽ,A.����ԭ��  " & _
                " FROM ҩƷ�շ���¼ A, ���ű� B ,���ű� C  " & strSqlForm & _
                " WHERE A.�ⷿID = B.ID AND A.�Է�����ID=C.ID AND A.���� = 7  " & _
                strUserPart & strFind & _
                " GROUP BY A.NO,C.����,A.������,A.�޸���,A.�����,A.��¼״̬,A.ժҪ,Nvl(A.��ҩ��ʽ,0),A.����ԭ�� " & _
                " ORDER BY NO DESC, �������� ASC "
        Case ģ���.��������          'ҩƷ�����������
            If SQLCondition.int�������һ����ѯ = 0 Then
                gstrSQL = "SELECT A.NO, C.���� AS ������,Decode(C.����, 'ҩƷ���', D.����, Decode(C.����, 'ҩƷ����', E.����, '')) AS �Է���λ,LTRIM(TO_CHAR (SUM (A.�ɱ����), " & mstrMoneyFormat & ")) AS �ɱ����, " & _
                    " LTRIM(TO_CHAR ( (SUM (A.���۽��)), " & mstrMoneyFormat & ")) AS �ۼ۽��, LTRIM(TO_CHAR((SUM(A.���۽�� - A.�ɱ����))," & mstrMoneyFormat & " )) AS ��۽��,LTRIM(TO_CHAR((SUM(A.���� * A.ʵ������))," & mstrMoneyFormat & " )) AS ����������,A.������," & _
                    " TO_CHAR (MIN(A.��������), 'YYYY-MM-DD HH24:MI:SS') AS ��������, A.�޸���,TO_CHAR (MIN(A.�޸�����), 'YYYY-MM-DD HH24:MI:SS') AS �޸�����, A.�����," & _
                    " TO_CHAR (MIN(A.�������), 'YYYY-MM-DD HH24:MI:SS') AS �������, A.��¼״̬, A.ժҪ " & _
                    " FROM ҩƷ�շ���¼ A, ���ű� B,ҩƷ������ C,ҩƷ�����λ D, ҩƷ������λ E  " & strSqlForm & _
                    " WHERE A.�ⷿID = B.ID AND A.������ID = C.ID AND A.��ҩ����=D.����(+) And A.��ҩ���� = E.����(+) AND A.���� = 11 " & _
                    strUserPart & strFind & _
                    " GROUP BY A.NO,C.����,D.����,E.����,A.������, A.�޸���,A.�����,A.��¼״̬,A.ժҪ " & _
                    " ORDER BY NO DESC,�������� ASC "
            Else
                gstrSQL = "SELECT A.NO, C.���� AS ������,Decode(C.����, 'ҩƷ���', D.����, Decode(C.����, 'ҩƷ����', E.����, '')) AS �Է���λ,LTRIM(TO_CHAR (SUM (A.�ɱ����), " & mstrMoneyFormat & ")) AS �ɱ����, " & _
                    " LTRIM(TO_CHAR ( (SUM (A.���۽��)), " & mstrMoneyFormat & ")) AS �ۼ۽��, LTRIM(TO_CHAR((SUM(A.���۽�� - A.�ɱ����))," & mstrMoneyFormat & " )) AS ��۽��,LTRIM(TO_CHAR((SUM(A.���� * A.ʵ������))," & mstrMoneyFormat & " )) AS ����������,A.������," & _
                    " TO_CHAR (MIN(A.��������), 'YYYY-MM-DD HH24:MI:SS') AS ��������, A.�޸���,TO_CHAR (MIN(A.�޸�����), 'YYYY-MM-DD HH24:MI:SS') AS �޸�����, A.�����," & _
                    " TO_CHAR (MIN(A.�������), 'YYYY-MM-DD HH24:MI:SS') AS �������, A.��¼״̬, A.ժҪ " & _
                    " FROM ҩƷ�շ���¼ A, ���ű� B,ҩƷ������ C,ҩƷ�����λ D, ҩƷ������λ E  " & strSqlForm & _
                    " WHERE A.�ⷿID = B.ID AND A.������ID = C.ID AND A.��ҩ����=D.����(+) And A.��ҩ���� = E.����(+) AND A.���� = 11 " & _
                    strUserPart & strFind & " And (A.�������� Between [3] And [4]) And A.������� Is Null " & _
                    " GROUP BY A.NO,C.����,D.����,E.����,A.������, A.�޸���,A.�����,A.��¼״̬,A.ժҪ " & _
                    " Union " & _
                    " SELECT A.NO, C.���� AS ������,Decode(C.����, 'ҩƷ���', D.����, Decode(C.����, 'ҩƷ����', E.����, '')) AS �Է���λ,LTRIM(TO_CHAR (SUM (A.�ɱ����), " & mstrMoneyFormat & ")) AS �ɱ����, " & _
                    " LTRIM(TO_CHAR ( (SUM (A.���۽��)), " & mstrMoneyFormat & ")) AS �ۼ۽��, LTRIM(TO_CHAR((SUM(A.���۽�� - A.�ɱ����))," & mstrMoneyFormat & " )) AS ��۽��,LTRIM(TO_CHAR((SUM(A.���� * A.ʵ������))," & mstrMoneyFormat & " )) AS ����������,A.������," & _
                    " TO_CHAR (MIN(A.��������), 'YYYY-MM-DD HH24:MI:SS') AS ��������, A.�޸���,TO_CHAR (MIN(A.�޸�����), 'YYYY-MM-DD HH24:MI:SS') AS �޸�����, A.�����," & _
                    " TO_CHAR (MIN(A.�������), 'YYYY-MM-DD HH24:MI:SS') AS �������, A.��¼״̬, A.ժҪ " & _
                    " FROM ҩƷ�շ���¼ A, ���ű� B,ҩƷ������ C,ҩƷ�����λ D, ҩƷ������λ E " & strSqlForm & _
                    " WHERE A.�ⷿID = B.ID AND A.������ID = C.ID AND A.��ҩ����=D.����(+) And A.��ҩ���� = E.����(+) AND A.���� = 11 " & _
                    strUserPart & strFind & " And (A.������� Between [5] And [6]) " & _
                    " GROUP BY A.NO,C.����,D.����,E.����,A.������, A.�޸���,A.�����,A.��¼״̬,A.ժҪ " & _
                    " ORDER BY NO DESC,�������� ASC "
            End If
    End Select

    Set rsList = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption, _
        SQLCondition.strNO��ʼ, _
        SQLCondition.strNO����, _
        SQLCondition.date����ʱ�俪ʼ, _
        SQLCondition.date����ʱ�����, _
        SQLCondition.date���ʱ�俪ʼ, _
        SQLCondition.date���ʱ�����, _
        SQLCondition.lngҩƷ, _
        SQLCondition.lng�ⷿ, _
        SQLCondition.str������, _
        SQLCondition.str�����, _
        SQLCondition.lng������, _
        SQLCondition.str����, _
        SQLCondition.str��Ʊ�ſ�ʼ, _
        SQLCondition.str��Ʊ�Ž���, _
        SQLCondition.lng������, _
        cboStock.ItemData(cboStock.ListIndex), _
        UserInfo.�û�ID, _
        gstrNodeNo, _
        SQLCondition.date��Ʊ������ڿ�ʼ, _
        SQLCondition.date��Ʊ������ڽ���, _
        SQLCondition.str����, _
        SQLCondition.lngҩƷ����)
    
    mblnBandEvent = True
    Set vsfList.DataSource = rsList
    mblnBandEvent = False
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
        
        For n = 0 To .Cols - 1
            .ColKey(n) = .TextMatrix(0, n)
            .FixedAlignment(n) = flexAlignCenterCenter
        Next
    End With
    SetListColWidth
    DoEvents
    
    'ͳ�ƺϼƽ��
    lbl1.Caption = ""
    lbl2.Caption = ""
    lbl3.Caption = ""
    If (Not rsList.EOF) And (Not rsList.BOF) Then
        rsList.MoveFirst
        Do While Not rsList.EOF
            Select Case mlngMode
                Case ģ���.�⹺���
                    dbl1 = dbl1 + IIf(IsNull(rsList!�ɱ����), 0, rsList!�ɱ����)
                    dbl2 = dbl2 + IIf(IsNull(rsList!�ۼ۽��), 0, rsList!�ۼ۽��)
                    dbl3 = dbl3 + IIf(IsNull(rsList!��۽��), 0, rsList!��۽��)
                Case ģ���.�������, ģ���.�������, ģ���.ҩƷ�ƿ�, ģ���.ҩƷ����
                    dbl1 = dbl1 + IIf(IsNull(rsList!�ɱ����), 0, rsList!�ɱ����)
                    dbl2 = dbl2 + IIf(IsNull(rsList!�ۼ۽��), 0, rsList!�ۼ۽��)
                    dbl3 = dbl3 + IIf(IsNull(rsList!��۽��), 0, rsList!��۽��)
                Case ģ���.��������
                    dbl1 = dbl1 + IIf(IsNull(rsList!�ɱ����), 0, rsList!�ɱ����)
                    dbl2 = dbl2 + IIf(IsNull(rsList!�ۼ۽��), 0, rsList!�ۼ۽��)
                    dbl3 = dbl3 + IIf(IsNull(rsList!��۽��), 0, rsList!��۽��)
                    dbl4 = dbl4 + IIf(IsNull(rsList!����������), 0, rsList!����������)
                Case ģ���.��۵���
                    dbl1 = dbl1 + IIf(IsNull(rsList!�����), 0, rsList!�����)
                    dbl2 = dbl2 + IIf(IsNull(rsList!�����), 0, rsList!�����)
                    dbl3 = dbl3 + IIf(IsNull(rsList!������), 0, rsList!������)
            End Select
            rsList.MoveNext
        Loop
        
        rsList.MoveFirst
        
        Select Case mlngMode
            Case ģ���.�⹺���
                lbl1.Caption = "�ɱ����ϼƣ�" & Format(dbl1, StrFormat)
                lbl2.Caption = "�ۼ۽��ϼƣ�" & Format(dbl2, StrFormat)
                lbl3.Caption = "��۽��ϼƣ�" & Format(dbl3, StrFormat)
            Case ģ���.�������, ģ���.�������, ģ���.ҩƷ�ƿ�, ģ���.ҩƷ����
                lbl1.Caption = "�ɱ����ϼƣ�" & Format(dbl1, StrFormat)
                lbl2.Caption = "�ۼ۽��ϼƣ�" & Format(dbl2, StrFormat)
                lbl3.Caption = "��۽��ϼƣ�" & Format(dbl3, StrFormat)
            Case ģ���.��������
                lbl1.Caption = "�ɱ����ϼƣ�" & Format(dbl1, StrFormat)
                lbl2.Caption = "�ۼ۽��ϼƣ�" & Format(dbl2, StrFormat)
                lbl3.Caption = "��۽��ϼƣ�" & Format(dbl3, StrFormat)
                lbl4.Caption = "���(��)���ϼƣ�" & Format(dbl4, StrFormat)
            Case ģ���.��۵���
                lbl1.Caption = "�����ϼƣ�" & Format(dbl1, StrFormat)
                lbl2.Caption = "����ۺϼƣ�" & Format(dbl2, StrFormat)
                lbl3.Caption = "������ϼƣ�" & Format(dbl3, StrFormat)
        End Select
    
    End If

    With vsfList
        If mintListRow >= .rows Then
            mintListRow = .rows - 1
        End If
        .Row = IIf(.rows > 1, IIf(mintListRow > 1, mintListRow, 1), 1)
        .Col = 0
'        .ColSel = .Cols - 1
    End With
    
    vsfList_EnterCell    '�г�������
    
    SetListColor
    
    SetEnable
    vsfList.Redraw = flexRDDirect
    Call FS.StopFlash

    Screen.MousePointer = vbDefault
    staThis.Panels(2).Text = "��ǰ����" & rsList.RecordCount & "�ŵ���"
    rsList.Close
    
    If vsfList.Visible = True Then
        vsfList.SetFocus
        vsfList.Row = IIf(vsfList.rows > 1, IIf(mintListRow > 1, mintListRow, 1), 1)
        vsfList.TopRow = vsfList.Row
    End If
    If mblnDo Then
        RestoreFlexState vsfList, App.ProductName & "\" & Me.Name & mstrTitle
        RestoreFlexState vsfDetail, App.ProductName & "\" & Me.Name & mstrTitle
    End If
    If mblnViewCost = False Then
        lbl1.Visible = False
        lbl3.Visible = False
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetListColor()
    Dim intStatus As Integer
    Dim intRow As Integer
    Dim intCol As Integer
    Dim int����״̬ As Integer      '0-������ĳ�����¼;1-����˵ĳ�����¼
    
    With vsfList
        If .rows <= 2 Then Exit Sub
        For intRow = 1 To .rows - 1
'            intStatus = .TextMatrix(intRow, IIf(mlngMode = ģ���.��۵��� Or mlngMode = ģ���.�⹺��� Or mlngMode = ģ���.ҩƷ����, .Cols - 3, .Cols - 2))
            intStatus = .TextMatrix(intRow, .ColIndex("��¼״̬"))
            If intStatus Mod 3 = 0 Then
                .Cell(flexcpForeColor, intRow, 0, intRow, .Cols - 1) = &H80000001
            End If
            If intStatus Mod 3 = 2 Then
                '�ƿ������������������ж�״̬
                If mlngMode = ģ���.ҩƷ�ƿ� Then
                    If Trim(.TextMatrix(intRow, .ColIndex("��������"))) <> "" Then
                        int����״̬ = 1
                    Else
                        int����״̬ = 0
                    End If
                End If
                
                '���ó�����������������ж�״̬
                If mlngMode = ģ���.ҩƷ���� Then
                    If Trim(.TextMatrix(intRow, .ColIndex("�������"))) <> "" Then
                        int����״̬ = 1
                    Else
                        int����״̬ = 0
                    End If
                End If
                
                If mlngMode = ģ���.�⹺��� Then
                    '�⹺����в������ʱ�����ĵ���Ϊǳ��ɫ��������ͨ��������Ϊ��ɫ
                    .Cell(flexcpForeColor, intRow, 0, intRow, .Cols - 1) = IIf(Val(.TextMatrix(intRow, �⹺����.��������)) = 1, &HFF00FF, &HFF)
                ElseIf mlngMode = ģ���.ҩƷ�ƿ� Or mlngMode = ģ���.ҩƷ���� Then
                    '�ƿ⡢�����������������Ϊǳ��ɫ���ѳ�������Ϊ��ɫ
                    .Cell(flexcpForeColor, intRow, 0, intRow, .Cols - 1) = IIf(int����״̬ = 0, &HFF00FF, &HFF)
                Else
                    '���������ѳ�������Ϊ��ɫ
                    .Cell(flexcpForeColor, intRow, 0, intRow, .Cols - 1) = &HFF
                End If
                
                '���ô��ڳ������ݣ�����ԭ������ʾ
                If mlngMode = ģ���.ҩƷ���� Then If vsfList.colHidden(.ColIndex("����ԭ��")) Then vsfList.colHidden(.ColIndex("����ԭ��")) = False
                
            End If
        Next
    End With
End Sub

'��ͷ�п��ʼ
Private Sub SetListColWidth()
    Dim intCol As Integer
    
    With vsfList
        Select Case mlngMode
            Case ģ���.�⹺���
                .ColAlignment(�⹺����.�ɱ����) = flexAlignRightCenter
                .ColAlignment(�⹺����.�ۼ۽��) = flexAlignRightCenter
                .ColAlignment(�⹺����.��۽��) = flexAlignRightCenter
                .ColAlignment(�⹺����.���۽��) = flexAlignRightCenter
                .ColAlignment(�⹺����.���۲��) = flexAlignRightCenter
            Case ģ���.�������
                .ColAlignment(2) = flexAlignRightCenter
                .ColAlignment(3) = flexAlignRightCenter
                .ColAlignment(4) = flexAlignRightCenter
            Case ģ���.�������
                .ColAlignment(2) = flexAlignRightCenter
                .ColAlignment(3) = flexAlignRightCenter
                .ColAlignment(4) = flexAlignRightCenter
                .ColAlignment(5) = flexAlignRightCenter
                .ColAlignment(6) = flexAlignRightCenter
            Case ģ���.��۵���
                .ColAlignment(1) = flexAlignRightCenter
                .ColAlignment(2) = flexAlignRightCenter
                .ColAlignment(3) = flexAlignRightCenter
            Case ģ���.ҩƷ�ƿ�
                .ColAlignment(2) = flexAlignRightCenter
                .ColAlignment(3) = flexAlignRightCenter
                .ColAlignment(4) = flexAlignRightCenter
            Case ģ���.ҩƷ����
                .ColAlignment(2) = flexAlignRightCenter
                .ColAlignment(3) = flexAlignRightCenter
                .ColAlignment(4) = flexAlignRightCenter
            Case ģ���.��������
                .ColAlignment(3) = flexAlignRightCenter         '�ۼ۽��
                .ColAlignment(4) = flexAlignRightCenter
                .ColAlignment(5) = flexAlignRightCenter
                .ColAlignment(6) = flexAlignRightCenter
        End Select
        
        For intCol = 1 To .Cols - 1
            If intCol = 1 Then
                If mlngMode = ģ���.��۵��� Then
                    .ColWidth(intCol) = 1000
                Else
                    .ColWidth(intCol) = 2000
                End If
            ElseIf intCol = .Cols - 2 Then
                  .ColWidth(intCol) = 0
            Else
                .ColWidth(intCol) = 1000
            End If
        Next
        If mlngMode = ģ���.��۵��� Then
            .ColWidth(.Cols - 1) = 0
            .ColWidth(.Cols - 3) = 0
        End If
        If mlngMode = ģ���.�⹺��� Then
'            .ColWidth(�⹺����.���۽��) = IIf(InStr("|" & mstr������ & "|", "|���۽��|") = 0, 1000, 0)
'            .ColWidth(�⹺����.���۲��) = IIf(InStr("|" & mstr������ & "|", "|���۲��|") = 0, 1000, 0)
            .ColWidth(�⹺����.���۽��) = 0
            .ColWidth(�⹺����.���۲��) = 0
            .ColWidth(�⹺����.��������) = 0
            .ColWidth(�⹺����.��¼״̬) = 0
            .ColWidth(�⹺����.����˵��) = 1000
            If mbln�˲� = True Then
                If .ColWidth(�⹺����.�˲���) = 0 Then .ColWidth(�⹺����.�˲���) = 1000
                If .ColWidth(�⹺����.�˲�����) = 0 Then .ColWidth(�⹺����.�˲�����) = 1000
            Else
                .ColWidth(�⹺����.�˲���) = 0
                .ColWidth(�⹺����.�˲�����) = 0
            End If
        End If
        If mlngMode = ģ���.������� Then
'            .ColWidth(5) = IIf(InStr("|" & mstr������ & "|", "|���۽��|") = 0, 1000, 0)
'            .ColWidth(6) = IIf(InStr("|" & mstr������ & "|", "|���۲��|") = 0, 1000, 0)
            .ColWidth(5) = 0
            .ColWidth(6) = 0
        End If
        If mlngMode = ģ���.ҩƷ���� Then
            .ColWidth(.Cols - 4) = 0
            .ColWidth(.Cols - 3) = 1000
            .ColWidth(.Cols - 2) = 1000
            
            .colHidden(.ColIndex("����ԭ��")) = True '����ԭ��Ĭ�ϲ���ʾ
        End If
        If mblnViewCost = False Then
            Select Case mlngMode
                Case ģ���.�⹺���
                    .colHidden(�⹺����.�ɱ����) = True
                    .colHidden(�⹺����.��۽��) = True
                Case ģ���.�������
                    .colHidden(.ColIndex("�ɱ����")) = True
                    .colHidden(.ColIndex("��۽��")) = True
                Case ģ���.�������
                    .colHidden(.ColIndex("�ɱ����")) = True
                    .colHidden(.ColIndex("��۽��")) = True
                Case ģ���.ҩƷ�ƿ�
                    .colHidden(.ColIndex("�ɱ����")) = True
                    .colHidden(.ColIndex("��۽��")) = True
                Case ģ���.ҩƷ����
                    .colHidden(.ColIndex("�ɱ����")) = True
                    .colHidden(.ColIndex("��۽��")) = True
                Case ģ���.��������
                    .colHidden(.ColIndex("�ɱ����")) = True
                    .colHidden(.ColIndex("��۽��")) = True
            End Select
        End If
    End With
End Sub

Private Sub SetDetailColWidth()
    Dim intCol As Integer
    Dim rsDetail As New Recordset
    Dim bln��ҩ�ⷿ As Boolean
    Dim str�ⷿ���� As String
    
    On Error GoTo errHandle

    With vsfDetail
        Select Case mlngMode
            Case ģ���.�⹺���
                .ColAlignment(.ColIndex("����")) = flexAlignRightCenter     '����
                .ColAlignment(.ColIndex("��λ")) = flexAlignCenterCenter    '��λ
                .ColAlignment(.ColIndex("�ɱ���")) = flexAlignRightCenter     '�ɱ���
                .ColAlignment(.ColIndex("�ɱ����")) = flexAlignRightCenter     '�ɱ����
                .ColAlignment(.ColIndex("����")) = flexAlignRightCenter     '����
                .ColAlignment(.ColIndex("�ۼ�")) = flexAlignRightCenter    '�ۼ�
                .ColAlignment(.ColIndex("�ۼ۽��")) = flexAlignRightCenter    '�ۼ۽��
                .ColAlignment(.ColIndex("���")) = flexAlignRightCenter    '���
                .ColAlignment(.ColIndex("��׼�ĺ�")) = flexAlignLeftCenter     '��׼�ĺ�
                .ColAlignment(.ColIndex("��Ʊ���")) = flexAlignRightCenter    '��Ʊ���
                .ColAlignment(.ColIndex("���������")) = flexAlignRightCenter    '���������
                .ColAlignment(.ColIndex("�������")) = flexAlignLeftCenter     '�������
                .ColAlignment(.ColIndex("���ۼ�")) = flexAlignRightCenter    '���ۼ�
                .ColAlignment(.ColIndex("���۽��")) = flexAlignRightCenter    '���۽��
                .ColAlignment(.ColIndex("���۲��")) = flexAlignRightCenter    '���۲��
            Case ģ���.�������
                .ColAlignment(.ColIndex("����")) = flexAlignRightCenter     '����
                .ColAlignment(.ColIndex("��λ")) = flexAlignCenterCenter    '��λ
                .ColAlignment(.ColIndex("�ɹ���")) = flexAlignRightCenter     '�ɱ���
                .ColAlignment(.ColIndex("�ɹ����")) = flexAlignRightCenter     '�ɱ����
                .ColAlignment(.ColIndex("�ۼ�")) = flexAlignRightCenter    '�ۼ�
                .ColAlignment(.ColIndex("�ۼ۽��")) = flexAlignRightCenter    '�ۼ۽��
                .ColAlignment(.ColIndex("���")) = flexAlignRightCenter    '���
            Case ģ���.�������
                .ColAlignment(.ColIndex("����")) = flexAlignRightCenter     '����
                .ColAlignment(.ColIndex("��λ")) = flexAlignCenterCenter    '��λ
                .ColAlignment(.ColIndex("�ɱ���")) = flexAlignRightCenter     '�ɱ���
                .ColAlignment(.ColIndex("�ɱ����")) = flexAlignRightCenter     '�ɱ����
                .ColAlignment(.ColIndex("�ۼ�")) = flexAlignRightCenter    '�ۼ�
                .ColAlignment(.ColIndex("�ۼ۽��")) = flexAlignRightCenter    '�ۼ۽��
                .ColAlignment(.ColIndex("���")) = flexAlignRightCenter    '���
                .ColAlignment(.ColIndex("���ۼ�")) = flexAlignRightCenter    '���ۼ�
                .ColAlignment(.ColIndex("���۽��")) = flexAlignRightCenter    '���۽��
                .ColAlignment(.ColIndex("���۲��")) = flexAlignRightCenter    '���۲��
            Case ģ���.��۵���
                .ColAlignment(.ColIndex("��λ")) = flexAlignCenterCenter    '��λ
                .ColAlignment(.ColIndex("�����")) = flexAlignRightCenter     '�����
                .ColAlignment(.ColIndex("�����")) = flexAlignRightCenter     '�����
                .ColAlignment(.ColIndex("������")) = flexAlignRightCenter     '������
                .ColAlignment(.ColIndex("�³ɱ���")) = flexAlignRightCenter     '�³ɱ���
            Case ģ���.ҩƷ�ƿ�
                .ColAlignment(.ColIndex("��д����")) = flexAlignRightCenter     '��д����
                .ColAlignment(.ColIndex("ʵ������")) = flexAlignRightCenter     'ʵ������
                .ColAlignment(.ColIndex("��λ")) = flexAlignCenterCenter    '��λ
                .ColAlignment(.ColIndex("�ɱ���")) = flexAlignRightCenter     '�ɱ���
                .ColAlignment(.ColIndex("�ɱ����")) = flexAlignRightCenter     '�ɱ����
                .ColAlignment(.ColIndex("�ۼ�")) = flexAlignRightCenter    '�ۼ�
                .ColAlignment(.ColIndex("�ۼ۽��")) = flexAlignRightCenter    '�ۼ۽��
                .ColAlignment(.ColIndex("���")) = flexAlignRightCenter    '���
            Case ģ���.ҩƷ����
                .ColAlignment(.ColIndex("��д����")) = flexAlignRightCenter     '��д����
                .ColAlignment(.ColIndex("ʵ������")) = flexAlignRightCenter     'ʵ������
                .ColAlignment(.ColIndex("��λ")) = flexAlignCenterCenter    '��λ
                .ColAlignment(.ColIndex("�ɱ���")) = flexAlignRightCenter     '�ɱ���
                .ColAlignment(.ColIndex("�ɱ����")) = flexAlignRightCenter     '�ɱ����
                .ColAlignment(.ColIndex("�ۼ�")) = flexAlignRightCenter    '�ۼ�
                .ColAlignment(.ColIndex("�ۼ۽��")) = flexAlignRightCenter    '�ۼ۽��
                .ColAlignment(.ColIndex("���")) = flexAlignRightCenter    '���
            Case ģ���.��������
                .ColAlignment(.ColIndex("����")) = flexAlignRightCenter     '����
                .ColAlignment(.ColIndex("��λ")) = flexAlignCenterCenter    '��λ
                .ColAlignment(.ColIndex("�ɱ���")) = flexAlignRightCenter     '�ɱ���
                .ColAlignment(.ColIndex("�ɱ����")) = flexAlignRightCenter     '�ɱ����
                .ColAlignment(.ColIndex("�ۼ�")) = flexAlignRightCenter    '�ۼ�
                .ColAlignment(.ColIndex("�ۼ۽��")) = flexAlignRightCenter    '�ۼ۽��
                .ColAlignment(.ColIndex("���")) = flexAlignRightCenter    '���
                If vsfList.TextMatrix(vsfList.Row, 1) = "ҩƷ����" Then
                    .ColAlignment(.ColIndex("������")) = flexAlignRightCenter    '�����/������
                    .ColAlignment(.ColIndex("�������")) = flexAlignRightCenter    '������/�������
                Else
                    .ColAlignment(.ColIndex("�����")) = flexAlignRightCenter    '�����/������
                    .ColAlignment(.ColIndex("������")) = flexAlignRightCenter    '������/�������
                End If
                .ColAlignment(.ColIndex("��ֵ˰��")) = flexAlignRightCenter    '��ֵ˰��
                .ColAlignment(.ColIndex("˰��")) = flexAlignRightCenter    '˰��
        End Select
        
        If mblnBootUp = False Then
            .ColWidth(0) = 500
            If mlngMode = ģ���.�⹺��� Then
                .ColWidth(1) = 1000
            Else
                .ColWidth(1) = 2500
            End If
            For intCol = 2 To .Cols - 1
                .ColWidth(intCol) = 1000
            Next
        End If
        
        str�ⷿ���� = ""
        gstrSQL = "Select a.�������� From ��������˵�� A Where a.����id =[1]"
        Set rsDetail = zlDataBase.OpenSQLRecord(gstrSQL, "�ж��ǿⷿ����", cboStock.ItemData(cboStock.ListIndex))
        Do While Not rsDetail.EOF
            str�ⷿ���� = str�ⷿ���� & "," & rsDetail!��������
            rsDetail.MoveNext
        Loop
        If str�ⷿ���� Like "*��ҩ*" Or str�ⷿ���� Like "*�Ƽ���*" Then bln��ҩ�ⷿ = True
        
        Select Case mlngMode
            Case ģ���.�⹺���
                .ColWidth(.ColIndex("�������")) = 0
                .ColWidth(.ColIndex("�б�ҩƷ")) = 0
                .ColWidth(.ColIndex("���������")) = 0
'                .ColWidth(.ColIndex("���ۼ�")) = IIf(InStr("|" & mstr������ & "|", "|���ۼ�|") = 0, 1000, 0)
'                .ColWidth(.ColIndex("���۽��")) = IIf(InStr("|" & mstr������ & "|", "|���۽��|") = 0, 1000, 0)
'                .ColWidth(.ColIndex("���۲��")) = IIf(InStr("|" & mstr������ & "|", "|���۲��|") = 0, 1000, 0)
                .ColWidth(.ColIndex("���ۼ�")) = 0
                .ColWidth(.ColIndex("���۽��")) = 0
                .ColWidth(.ColIndex("���۲��")) = 0
                If mblnViewCost = False Then
                    .colHidden(.ColIndex("�ɱ���")) = True
                    .colHidden(.ColIndex("�ɱ����")) = True
                    .colHidden(.ColIndex("���")) = True
                End If
                If bln��ҩ�ⷿ Then
                    .colHidden(.ColIndex("ԭ����")) = False
                Else
                    .colHidden(.ColIndex("ԭ����")) = True
                End If
            Case ģ���.�������
'                .ColWidth(.ColIndex("���ۼ�")) = IIf(InStr("|" & mstr������ & "|", "|���ۼ�|") = 0, 1000, 0)
'                .ColWidth(.ColIndex("���۽��")) = IIf(InStr("|" & mstr������ & "|", "|���۽��|") = 0, 1000, 0)
'                .ColWidth(.ColIndex("���۲��")) = IIf(InStr("|" & mstr������ & "|", "|���۲��|") = 0, 1000, 0)
                .ColWidth(.ColIndex("���ۼ�")) = 0
                .ColWidth(.ColIndex("���۽��")) = 0
                .ColWidth(.ColIndex("���۲��")) = 0
                If mblnViewCost = False Then
                    .colHidden(.ColIndex("�ɱ���")) = True
                    .colHidden(.ColIndex("�ɱ����")) = True
                    .colHidden(.ColIndex("���")) = True
                End If
                If bln��ҩ�ⷿ Then
                    .colHidden(.ColIndex("ԭ����")) = False
                Else
                    .colHidden(.ColIndex("ԭ����")) = True
                End If
            Case ģ���.�������
                If mblnViewCost = False Then
                    .colHidden(.ColIndex("�ɹ���")) = True
                    .colHidden(.ColIndex("�ɹ����")) = True
                    .colHidden(.ColIndex("���")) = True
                End If
            Case ģ���.��۵���
'                If mblnViewCost = False Then
'                    .ColHidden(.ColIndex("�����")) = True
'                    .ColHidden(.ColIndex("�³ɱ���")) = True
'                End If
                If bln��ҩ�ⷿ Then
                    .colHidden(.ColIndex("ԭ����")) = False
                Else
                    .colHidden(.ColIndex("ԭ����")) = True
                End If
            Case ģ���.ҩƷ�ƿ�
                If mblnViewCost = False Then
                    .colHidden(.ColIndex("�ɱ���")) = True
                    .colHidden(.ColIndex("�ɱ����")) = True
                    .colHidden(.ColIndex("���")) = True
                End If
                If bln��ҩ�ⷿ Then
                    .colHidden(.ColIndex("ԭ����")) = False
                Else
                    .colHidden(.ColIndex("ԭ����")) = True
                End If
            Case ģ���.ҩƷ����
                If mblnViewCost = False Then
                    .colHidden(.ColIndex("�ɱ���")) = True
                    .colHidden(.ColIndex("�ɱ����")) = True
                    .colHidden(.ColIndex("���")) = True
                End If
                .ColWidth(.ColIndex("����")) = 0
                If bln��ҩ�ⷿ Then
                    .colHidden(.ColIndex("ԭ����")) = False
                Else
                    .colHidden(.ColIndex("ԭ����")) = True
                End If
            Case ģ���.��������
                If mblnViewCost = False Then
                    .colHidden(.ColIndex("�ɱ���")) = True
                    .colHidden(.ColIndex("�ɱ����")) = True
                    .colHidden(.ColIndex("���")) = True
                End If
                .ColWidth(.ColIndex("����")) = 0
                
                If vsfList.TextMatrix(vsfList.Row, 1) = "ҩƷ���" Then
                   .ColWidth(.ColIndex("��ֵ˰��")) = 0
                   .ColWidth(.ColIndex("˰��")) = 0
                ElseIf vsfList.TextMatrix(vsfList.Row, 1) = "ҩƷ����" Then
                Else
                    .ColWidth(.ColIndex("�����")) = 0
                    .ColWidth(.ColIndex("������")) = 0
                    .ColWidth(.ColIndex("��ֵ˰��")) = 0
                    .ColWidth(.ColIndex("˰��")) = 0
                End If
                If bln��ҩ�ⷿ Then
                    .colHidden(.ColIndex("ԭ����")) = False
                Else
                    .colHidden(.ColIndex("ԭ����")) = True
                End If
        End Select
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
    '�⹺�������Ȩ�ޣ��������á����������пⷿ���Ǽǡ��޸ġ�ɾ�������ա����������ݴ�ӡ
    
    If zlStr.IsHavePrivs(mstrPrivs, "������") And gtype_UserSysParms.P173_������Ǹ������ܽ��и������ = 1 And mlngMode = ģ���.�⹺��� Then
        mnuEditMark.Visible = True
    Else
        mnuEditMark.Visible = False
    End If
    Select Case mlngMode
        Case ģ���.�⹺���, ģ���.�������, ģ���.�������, ģ���.��۵���, ģ���.ҩƷ�ƿ�, ģ���.ҩƷ����, ģ���.��������
            If Not zlStr.IsHavePrivs(mstrPrivs, "�Ǽ�") Then
                mnuEditAdd.Visible = False
                mnuEditRestore.Visible = False
                tlbTool.Buttons("Add").Visible = False
            Else
                mnuEditRestore.Visible = True
            End If
            
            If Not zlStr.IsHavePrivs(mstrPrivs, "�޸�") Then
                mnuEditModify.Visible = False
                tlbTool.Buttons("Modify").Visible = False
            End If
            
            If Not zlStr.IsHavePrivs(mstrPrivs, "ɾ��") Then
                mnuEditDel.Visible = False
                tlbTool.Buttons("Delete").Visible = False
                 '��û�����б༭Ȩ��ʱ���Ѳ˵��͹������ϵ���Ӧ�ķָ������Ρ�
                If mnuEditAdd.Visible = False And mnuEditModify.Visible = False Then
                    mnuEditLine1.Visible = False
                    tlbTool.Buttons("EditSeparate").Visible = False
                End If
            End If
            
            If mlngMode = ģ���.ҩƷ���� Then
                If Not mblnStock Or Not zlStr.IsHavePrivs(mstrPrivs, "���") Then
                    mnuEditVerify.Visible = False
                    mnuEditBill.Visible = False
                    tlbTool.Buttons("Verify").Visible = False
                End If
            Else
                If Not zlStr.IsHavePrivs(mstrPrivs, "���") Then
                    mnuEditVerify.Visible = False
                    mnuEditBill.Visible = False
                    tlbTool.Buttons("Verify").Visible = False
                End If
            End If
            
            If Not zlStr.IsHavePrivs(mstrPrivs, "����") Then
                mnuEditStrike.Visible = False
                mnuEditWriteOff.Visible = False
                tlbTool.Buttons("Strike").Visible = False
                
                If mnuEditVerify.Visible = False Then
                    mnuEditLine2.Visible = False
                    tlbTool.Buttons("VerifySeparate").Visible = False
                End If
            End If
            If Not zlStr.IsHavePrivs(mstrPrivs, "���ݴ�ӡ") Then
                mnuFileBillPrint.Visible = False
                mnuFileBillPreview.Visible = False
            End If
        Case Else
    End Select
    
    If mlngMode = ģ���.�⹺��� Then
        mnuEditVerifySelect.Visible = True
        If zlStr.IsHavePrivs(mstrPrivs, "���") Then
            mnuEditLine0.Visible = True
            mnuEditBill.Visible = True
        Else
            mnuEditLine0.Visible = False
        End If
        mnuEditRestore.Visible = zlStr.IsHavePrivs(mstrPrivs, "�˻�")
        If zlStr.IsHavePrivs(mstrPrivs, "�������") Then
            mnuEditLine0.Visible = True
            mnuEditAcc.Visible = True
        Else
            If mnuEditBill.Visible = False Then mnuEditLine0.Visible = False
            mnuEditAcc.Visible = False
        End If
        
        If zlStr.IsHavePrivs(mstrPrivs, "ҩƷ��ҩ�ƻ�") Then
            mnuEditHandBack.Visible = True
        Else
            mnuEditHandBack.Visible = False
        End If
        
        If zlStr.IsHavePrivs(mstrPrivs, "ҩƷ�ƻ�����������") Then
            mnuEditMediPlanImport.Visible = True
        Else
            mnuEditMediPlanImport.Visible = False
        End If
                
        If zlStr.IsHavePrivs(mstrPrivs, "�˲�ɱ���") And mbln�˲� Then
            mnuEditPrepare.Visible = True
            mnuEditBack.Visible = True
            mnuEditLine3.Visible = True
            Me.tlbTool.Buttons("Prepare").Visible = True
            Me.tlbTool.Buttons("Back").Visible = True
            Me.tlbTool.Buttons("PrepareSeparate").Visible = True
        End If
        mnuEditPrepare.Enabled = mbln�˲�
        mnuEditBack.Enabled = mbln�˲�
        Me.tlbTool.Buttons("Prepare").Enabled = mbln�˲�
        Me.tlbTool.Buttons("Back").Enabled = mbln�˲�
        
    ElseIf mlngMode = ģ���.ҩƷ�ƿ� Then
        mnuEditRestore.Visible = False
        mnuEditLine2.Visible = False
        If zlStr.IsHavePrivs(mstrPrivs, "����") Then
            mnuEditPreparePhysic.Visible = True
            mnuEditSendPhysic.Visible = True
            mnuEditBack.Visible = True
            mnuEditLine3.Visible = True
            tlbTool.Buttons("PreparePhysic").Visible = True
            tlbTool.Buttons("SendPhysic").Visible = True
            tlbTool.Buttons("Back").Visible = True
            tlbTool.Buttons("PrepareSeparate").Visible = True
        End If
        If Not zlStr.IsHavePrivs(mstrPrivs, "���") And _
           Not zlStr.IsHavePrivs(mstrPrivs, "����") Then
            TabShow.TabVisible(1) = False
        End If
    Else
        mnuEditBill.Visible = False
        mnuEditAcc.Visible = False
        mnuEditRestore.Visible = False
        mnuEditLine0.Visible = False
        mnuEditHandBack.Visible = False
    End If
End Sub

Private Sub Cmd����_Click()
    Call mnuEditDisplay_Click
End Sub

Private Sub Form_Activate()
    If vsfList.Visible = True Then
        vsfList.SetFocus
'        If vsfList.rows > 1 Then
'            vsfList.Row = 1
'        End If
        If vsfDetail.rows > 1 Then
            vsfDetail.Row = 1
        End If
    End If
End Sub

Private Sub Form_Load()
    '�ָ�����
    Dim dateCurrentDate As Date
    mbln����Ա���� = Not zlStr.IsHavePrivs(mstrPrivs, "���пⷿ")
    mblnViewCost = zlStr.IsHavePrivs(mstrPrivs, "�鿴�ɱ���")
    
    Me.Caption = mstrTitle
    dateCurrentDate = Sys.Currentdate
    lblRange.Caption = "��ѯ��Χ:" & Format(dateCurrentDate, "yyyy��MM��dd��") & "��" & Format(dateCurrentDate, "yyyy��MM��dd��")
    
    mnuViewLine3.Visible = (mlngMode = ģ���.�⹺��� Or mlngMode = ģ���.�������)
    mnuViewColDefine.Visible = (mlngMode = ģ���.�⹺��� Or mlngMode = ģ���.�������)
    
    mnuEditLine0.Visible = (mlngMode = ģ���.�⹺���)
    
    mnuFileCodePrint.Visible = (mlngMode = ģ���.�⹺��� Or mlngMode = ģ���.������� Or mlngMode = ģ���.ҩƷ���� Or mlngMode = ģ���.ҩƷ�ƿ� Or mlngMode = ģ���.��������)
    mnuEditCodePrintLine.Visible = mnuEditAllCodePrint.Visible
    
    TabShow.Visible = (mlngMode = ģ���.ҩƷ�ƿ�)
    mblnDo = Val(zlDataBase.GetPara("ʹ�ø��Ի����")) <> 0
    
    lbl1.Caption = ""
    lbl2.Caption = ""
    lbl3.Caption = ""
    lbl4.Caption = ""
    lbl2.Left = lbl1.Left + lbl1.Width + 2500
    lbl3.Left = lbl2.Left + lbl2.Width + 2500
    lbl4.Left = lbl3.Left + lbl3.Width + 2500
    If mblnViewCost = False Then
        lbl1.Visible = False
        lbl3.Visible = False
        lbl2.Left = lbl1.Left
        lbl4.Left = lbl2.Left + lbl2.Width + 2500
    End If
    
    Me.Top = (Screen.Height - Me.Height) / 2
    If mlngMode = ģ���.��۵��� Or mlngMode = ģ���.������� Then
        mnuEditWriteOff.Visible = False
    End If
    
    If mlngMode = ģ���.�⹺��� Then
        '�⹺ҵ����Ҳ���
        Call zlPlugIn_Ini(glngSys, glngModul, mobjPlugIn)
        
        '��Ҳ�������չ����
        Call zlPlugIn_SetVBMenu(glngSys, glngModul, mobjPlugIn, Me)
        
        '��Ҳ�������չ����
        Call zlPlugIn_SetVBToolbar(glngSys, glngModul, mobjPlugIn, Me, tlbTool, "PlugItem", "PlugInSeparator")
    End If
    
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
    
    With TabShow
        .Left = 0
        .Top = cbrTool.Height
    End With
    
    With vsfList
        .Top = IIf(cbrTool.Visible, cbrTool.Height, 0) + IIf(TabShow.Visible, TabShow.Height, 0)
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
    
    If mlngMode <> ģ���.�⹺��� And mlngMode <> ģ���.ҩƷ�ƿ� And mlngMode <> ģ���.ҩƷ���� Then
        picColor3.Visible = False
        lblColor3.Visible = False
        picColor.Width = lblColor2.Left + lblColor2.Width + 20
    Else
        If mlngMode = ģ���.ҩƷ�ƿ� Or mlngMode = ģ���.ҩƷ���� Then
            lblColor3.Caption = "δ��˳���"
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName, mstrTitle
    
    Call zlPlugIn_Unload(mobjPlugIn)
    
    mblnDo = False
End Sub

Private Sub mnuEditAcc_Click()
    Dim strNo As String
    Dim blnSuccess As Boolean
    Dim int��¼״̬ As Integer
    
    If cboStock.ListIndex = -1 Then
        MsgBox "��ѡ��ⷿ��", vbInformation, gstrSysName
        Exit Sub
    End If
    
    With vsfList
        strNo = .TextMatrix(.Row, 0)
        int��¼״̬ = .TextMatrix(.Row, .Cols - 3)
        frmPurchaseCard.ShowCard Me, strNo, �༭.�������, int��¼״̬, blnSuccess
        
        If blnSuccess = True Then
            mintListRow = vsfList.Row
            mnuViewRefresh_Click
        End If
    End With
End Sub

Private Sub mnuEditAdd_Click()
    Dim strNo As String
    Dim blnSuccess As Boolean
    Dim rsStock As ADODB.Recordset
    
    '��鱾���Ƿ��Ѿ���˽�棬���δ��˽�����ܽ�������ҵ�����
'    If CheckIsAccount(Val(cboStock.ItemData(cboStock.ListIndex))) = False Then
'        Exit Sub
'    End If
    If cboStock.ListIndex = -1 Then
        MsgBox "��ѡ��ⷿ��", vbInformation, gstrSysName
        Exit Sub
    End If
    strNo = ""
    '����
    Select Case mlngMode
        Case ģ���.�⹺���
            frmPurchaseCard.ShowCard Me, strNo, �༭.����, , blnSuccess
        Case ģ���.�������
            frmSelfMakeCard.ShowCard Me, strNo, �༭.����, , blnSuccess
        Case ģ���.�������
            frmOtherInputCard.ShowCard Me, strNo, �༭.����, , blnSuccess
        Case ģ���.��۵���
            frmDiffPriceAdjustCard.ShowCard Me, strNo, �༭.����, , blnSuccess
        Case ģ���.ҩƷ�ƿ�
            Set rsStock = ReturnSQL(Val(cboStock.ItemData(cboStock.ListIndex)), "mnuEditAdd_Click", True, ģ���.ҩƷ�ƿ�)
            If rsStock.EOF Then
                '�������������
                MsgBox "�ÿⷿδ����ҩƷ������ƣ����ܽ����������ݣ�", vbOKOnly, gstrSysName
                Exit Sub
            End If
            
            frmTransferCard.ShowCard Me, strNo, �༭.����, , blnSuccess
        Case ģ���.ҩƷ����
            frmDrawCard.ShowCard Me, strNo, �༭.����, mblnStock, , 0, blnSuccess
        Case ģ���.��������
            frmOtherOutputCard.ShowCard Me, strNo, �༭.����, , blnSuccess
    End Select
    
    If blnSuccess = True Then
        mintListRow = vsfList.Row + 1
        mnuViewRefresh_Click
    End If
End Sub

Private Sub mnuEditApplyStrike_Click()
    mnuEditApplyStrike.Tag = "1"
    Call mnuEditStrike_Click
End Sub

Private Sub mnuEditBack_Click()
    Dim strNo As String
    On Error GoTo ErrHand
    '1���ƿ⣺������һ��״̬�����δ��ҩֱ���˳���ֻ�ܴӷ��ͻ��˵���ҩ���ɱ�ҩ���˵��Ǳ�ҩ��
    '2���⹺�������˲�
    strNo = vsfList.TextMatrix(vsfList.Row, 0)
    
    If TestDelete(strNo) Then
        MsgBox "�õ����ѱ�ɾ����", vbInformation, gstrSysName
        mintListRow = vsfList.Row
        mnuViewRefresh_Click
        Exit Sub
    End If
    
    If TestVerify(strNo) Then
        MsgBox "�õ����ѱ���ˣ�", vbInformation, gstrSysName
        mintListRow = vsfList.Row + 1
        mnuViewRefresh_Click
        Exit Sub
    End If
    
    Select Case mlngMode
    Case ģ���.�⹺���   '�⹺��⳷���˲�
        gstrSQL = "Zl_ҩƷ�⹺_CancelCheck('" & strNo & "')"
    Case ģ���.ҩƷ�ƿ�   '�ƿ��˻�
        gstrSQL = "ZL_ҩƷ�ƿ�_BACK('" & strNo & "')"
    End Select
        
    Call zlDataBase.ExecuteProcedure(gstrSQL, "����")
    mintListRow = vsfList.Row
    Call mnuViewRefresh_Click
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub mnuEditBill_Click()
    Dim strNo As String
    Dim blnSuccess As Boolean
    Dim int��¼״̬ As Integer
    
    With vsfList
        strNo = .TextMatrix(.Row, 0)
        int��¼״̬ = .TextMatrix(.Row, .Cols - 3)
        frmPurchaseCard.ShowCard Me, strNo, �༭.�޸ķ�Ʊ, int��¼״̬, blnSuccess
        
        If blnSuccess = True Then
            mintListRow = vsfList.Row
            mnuViewRefresh_Click
        End If
    End With
    
End Sub

Private Sub mnuEditDeliveryInvoice_Click()
    gobjDrugPurchase.DeliveryInvoice gcnOracle
End Sub

Private Sub mnuEditHandBack_Click()
    frmHandBackPlan.ShowForm Me, mlng�ⷿID, mintUnit
End Sub

Private Sub mnuEditMark_Click()
    Call frmPurchaseMark.ShowME(mStr�ⷿ, cboStock.ListIndex, Me, mstrPrivs)
End Sub

Private Sub mnuEditMediPlanImport_Click()
    frmMediPlanImport.ShowCard Me, Val(cboStock.ItemData(cboStock.ListIndex))
    mnuViewRefresh_Click
End Sub

Private Sub mnuEditPrepare_Click()
    Dim strNo As String
    Dim blnSuccess As Boolean
    Dim rsTemp As New ADODB.Recordset
    '����⹺����Ǻ˲飨�������޸ĳɱ��ۣ�
    '����ƿⵥ������
    strNo = vsfList.TextMatrix(vsfList.Row, 0)
    If Trim(strNo) = "" Then Exit Sub
    
    Select Case mlngMode
    Case ģ���.�⹺���
        frmPurchaseCard.ShowCard Me, strNo, �༭.�˲�, vsfList.TextMatrix(vsfList.Row, �⹺����.��¼״̬), blnSuccess
    Case ģ���.ҩƷ�ƿ�
        If TestPrepare(strNo) Then
            MsgBox "���ƿⵥ[" & strNo & "]������ҩƷ�Ѿ����ͣ�", vbInformation, gstrSysName
            Exit Sub
        End If
        
    End Select
    
    If blnSuccess = True Then
        mintListRow = vsfList.Row
        mnuViewRefresh_Click
    End If
End Sub

Private Sub mnuEditPreparePhysic_Click()
    Dim strNo As String
    Dim strCheckString As String
    
    On Error GoTo ErrHand
    strNo = vsfList.TextMatrix(vsfList.Row, 0)
    If Trim(strNo) = "" Then Exit Sub
    
    strCheckString = CheckBill(Trim(strNo))
    If strCheckString <> "" Then
        MsgBox strCheckString, vbInformation, gstrSysName
        mintListRow = vsfList.Row + 1
        Call mnuViewRefresh_Click
        Exit Sub
    End If
    
    gstrSQL = "zl_ҩƷ�ƿ�_PREPARE('" & strNo & "','" & UserInfo.�û����� & "')"
    Call zlDataBase.ExecuteProcedure(gstrSQL, "��ҩ")
    
    mintListRow = vsfList.Row
    
    Call mnuViewRefresh_Click
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Public Function CheckBill(ByVal strNo As String) As String
    Dim rs As New ADODB.Recordset
    
    CheckBill = ""
    On Error GoTo errHandle
    gstrSQL = " Select �������,��ҩ����,��ҩ�� From ҩƷ�շ���¼ " & _
              " Where ����=6 And NO=[1] And ��¼״̬=1 And RowNum=1 "
    Set rs = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[��鵥��]", strNo)
            
    With rs
        '���ؿգ���ʾ�Ѿ�ɾ��
        If .EOF Then
            CheckBill = "�õ����Ѿ�����������Աɾ����"
        ElseIf Not IsNull(!�������) Then
            CheckBill = "�õ����Ѿ�����������Ա��ˣ�"
        ElseIf Not IsNull(!��ҩ����) Then
            CheckBill = "�õ����Ѿ�����������Ա���ͣ�"
        ElseIf Not IsNull(!��ҩ��) Then
            CheckBill = "�õ����Ѿ�����������Ա��ҩ��"
        End If
    End With
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Check�Ѹ����¼(ByVal strNo As String) As Boolean
    Dim strsql As String
    Dim rsCheck As ADODB.Recordset
    
    On Error GoTo errHandle
    strsql = "Select count(Id) �Ѹ��� From Ӧ����¼ Where �շ�id In(Select Id From ҩƷ�շ���¼ Where ����=5 And No=[1]) And nvl(�������,0)>0 "
    Set rsCheck = zlDataBase.OpenSQLRecord(strsql, Me.Caption & "[���Ӧ����¼]", strNo)
    
    Check�Ѹ����¼ = (rsCheck!�Ѹ��� > 0)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub mnuEditPric_Click()
    frmDiffPriceAdjustCard.ShowCard Me, "", �༭.����, , False, 2
End Sub

Private Sub mnuEditRestore_Click()
    Dim strNo As String
    Dim blnSuccess As Boolean
    
    If cboStock.ListIndex = -1 Then
        MsgBox "��ѡ��ⷿ��", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Call frmPurchaseCard.ShowCard(Me, strNo, �༭.ҩ���˻�, , blnSuccess)
    If blnSuccess Then
        mintListRow = vsfList.Row + 1
        mnuViewRefresh_Click
    End If
End Sub

Private Sub mnuEditSendPhysic_Click()
    Dim strNo As String
    Dim blnSuccess As Boolean
    On Error GoTo ErrHand
    strNo = vsfList.TextMatrix(vsfList.Row, 0)
    If Trim(strNo) = "" Then Exit Sub
    
    Call frmTransferCard.ShowCard(Me, strNo, �༭.����, 1, blnSuccess)
    If blnSuccess Then
        mintListRow = vsfList.Row
        mnuViewRefresh_Click
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub mnuEditVerify_Click()
    '����
    Dim strNo As String
    Dim blnSuccess As Boolean
    Dim int��˷�ʽ As Integer  '1-���������ĳ�������
    
    With vsfList
        strNo = .TextMatrix(.Row, 0)
        Select Case mlngMode
            Case ģ���.�⹺���
                If mbln�˲� Then
                    If Not TestPrepare(strNo) Then
                        MsgBox "�õ��ݻ�δͨ���˲飬��������ˣ�", vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
                frmPurchaseCard.ShowCard Me, strNo, �༭.���, .TextMatrix(.Row, �⹺����.��¼״̬), blnSuccess
            Case ģ���.�������
                frmSelfMakeCard.ShowCard Me, strNo, �༭.���, .TextMatrix(.Row, .Cols - 2), blnSuccess
            Case ģ���.�������
                frmOtherInputCard.ShowCard Me, strNo, �༭.���, .TextMatrix(.Row, .Cols - 2), blnSuccess
            Case ģ���.��۵���
                frmDiffPriceAdjustCard.ShowCard Me, strNo, �༭.���, .TextMatrix(.Row, .Cols - 3), blnSuccess, IIf(Val(vsfList.TextMatrix(vsfList.Row, vsfList.Cols - 1)) = 0, 1, 2)
            Case ģ���.ҩƷ�ƿ�
                frmTransferCard.ShowCard Me, strNo, �༭.���, .TextMatrix(.Row, .Cols - 2), blnSuccess
            Case ģ���.ҩƷ����
                frmDrawCard.ShowCard Me, strNo, �༭.���, mblnStock, .TextMatrix(.Row, .Cols - 4), 0, blnSuccess
            Case ģ���.��������
                frmOtherOutputCard.ShowCard Me, strNo, �༭.���, .TextMatrix(.Row, .Cols - 2), blnSuccess
        End Select
        
    End With
    If blnSuccess = True Then
        mintListRow = vsfList.Row
        mnuViewRefresh_Click
    End If
End Sub

Private Sub mnuEditDel_Click()
    'ɾ��
    Dim strBillNo As String
    Dim intRow As Integer
    Dim strTitle As String
    Dim intReturn As Integer
    Dim intRecord As Integer
    Dim rsCheck As New ADODB.Recordset
     
    With vsfList
        Select Case mlngMode
            Case ģ���.�⹺���
                strTitle = "�⹺��ⵥ"
            Case ģ���.�������
                strTitle = "������ⵥ"
            Case ģ���.�������
                strTitle = "������ⵥ"
            Case ģ���.��۵���
                strTitle = "����۵�����"
            Case ģ���.ҩƷ�ƿ�
                strTitle = "ҩƷ�ƿⵥ"
            Case ģ���.ҩƷ����
                strTitle = "ҩƷ���õ�"
            Case ģ���.��������
                strTitle = "ҩƷ�������ⵥ"
        End Select
        
        On Error GoTo errHandle
        intRow = .Row
        strBillNo = .TextMatrix(intRow, 0)
        intReturn = MsgBox("��ȷʵҪɾ�����ݺ�Ϊ��" & strBillNo & "����" & strTitle & "��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)
        intRecord = .rows - 1
        If intReturn = vbYes Then
            Select Case mlngMode
                Case ģ���.�⹺���
                    gstrSQL = "zl_ҩƷ�⹺_Delete('" & strBillNo & "')"
                Case ģ���.�������
                    gstrSQL = "zl_�������_Delete('" & strBillNo & "')"
                Case ģ���.�������
                    gstrSQL = "zl_ҩƷ�������_Delete('" & strBillNo & "')"
                Case ģ���.��۵���
                    gstrSQL = "zl_ҩƷ����۵���_Delete('" & strBillNo & "')"
                Case ģ���.ҩƷ�ƿ�
                    If Val(.TextMatrix(.Row, .Cols - 2)) = 1 Then
                        '�ѱ�ҩ����д����ҩ�ˣ����ѷ��͵ĵ��ݣ���������ⷽ�޸Ĵ��൥��
                        If TestDelete(strBillNo) Then
                            MsgBox "�õ����ѱ�ɾ����", vbInformation, gstrSysName
                            mintListRow = vsfList.Row
                            mnuViewRefresh_Click
                            Exit Sub
                        End If
                        If TestPrepare(strBillNo) Then
                            MsgBox "�ѱ�ҩ�ͷ��͵ĵ��ݲ�����ɾ����", vbInformation, gstrSysName
                            mintListRow = vsfList.Row + 1
                            mnuViewRefresh_Click
                            Exit Sub
                        End If
                    End If
                    
'                    If Is����(StrBillNo) Then
'                        If Not zlStr.IsHavePrivs(mstrPrivs, "�����޸����쵥") Then
'                            MsgBox "��û��Ȩ���޸����쵥��", vbInformation, gstrSysName
'                            Exit Sub
'                        End If
'                    End If
                    
                    '�����¼״̬����Ϊ�˿�����ɾ��δ��˵ĳ������뵥��
                    gstrSQL = "zl_ҩƷ�ƿ�_Delete('" & strBillNo & "'," & Val(.TextMatrix(.Row, .Cols - 2)) & " )"

                Case ģ���.ҩƷ����
                    gstrSQL = "zl_ҩƷ����_Delete('" & strBillNo & "'," & Val(.TextMatrix(.Row, .Cols - 4)) & " )"
                Case ģ���.��������
                    gstrSQL = "zl_ҩƷ��������_Delete('" & strBillNo & "')"
                Case Else
                
            End Select
            If gstrSQL = "" Then Exit Sub
            
            Call zlDataBase.ExecuteProcedure(gstrSQL, Me.Caption)
            
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
    If intRecord = 0 Then
        lbl1.Caption = ""
        lbl2.Caption = ""
        lbl3.Caption = ""
    End If
    mintListRow = vsfList.Row
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
        Select Case mlngMode
            Case ģ���.�⹺���
                frmPurchaseCard.ShowCard Me, strNo, �༭.����, .TextMatrix(.Row, �⹺����.��¼״̬)
                
            Case ģ���.�������
                frmSelfMakeCard.ShowCard Me, strNo, �༭.����, .TextMatrix(.Row, .Cols - 2)
            Case ģ���.�������
                frmOtherInputCard.ShowCard Me, strNo, �༭.����, .TextMatrix(.Row, .Cols - 2)
            Case ģ���.��۵���
                frmDiffPriceAdjustCard.ShowCard Me, strNo, �༭.����, .TextMatrix(.Row, .Cols - 3), False, IIf(Val(vsfList.TextMatrix(vsfList.Row, vsfList.Cols - 1)) = 0, 1, 2)
            Case ģ���.ҩƷ�ƿ�
                frmTransferCard.ShowCard Me, strNo, �༭.����, .TextMatrix(.Row, .Cols - 2)
            Case ģ���.ҩƷ����
                frmDrawCard.ShowCard Me, strNo, �༭.����, mblnStock, .TextMatrix(.Row, .Cols - 4), 0
            Case ģ���.��������
                frmOtherOutputCard.ShowCard Me, strNo, �༭.����, .TextMatrix(.Row, .Cols - 2)
            Case Else
        
        End Select
        
    End With
    
End Sub

Private Sub mnuEditStrike_Click()
    Dim blnPurchase As Boolean, blnRefresh As Boolean
    Dim strģ��� As String
    
    '������⹺(blnPurchaseΪ��)����ֱ�ӽ������
    'ѯ���Ƿ����(blnPurchaseΪ��ʾ�򷵻�ֵ)������������
    
    If cboStock.ListIndex = -1 Then
        MsgBox "��ѡ��ⷿ��", vbInformation, gstrSysName
        Exit Sub
    End If
    
    strģ��� = ģ���.�⹺��� & "," & ģ���.������� & "," & ģ���.ҩƷ�ƿ� & "," & ģ���.ҩƷ���� & "," & ģ���.��������
    blnPurchase = (InStr(1, strģ���, mlngMode) <> 0)
    With vsfList
        If Not blnPurchase Then
            blnPurchase = (MsgBox("��ȷʵҪȫ���������ݺ�Ϊ��" & .TextMatrix(.Row, 0) & "���ĵ�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes)
        End If
        If blnPurchase Then
            blnRefresh = StrikeSave
            If blnRefresh Then
                mintListRow = vsfList.Row
                mnuViewRefresh_Click
            End If
        End If
    End With
End Sub

Private Function StrikeSave() As Boolean
    Dim blnSuccess As Boolean
    Dim int����ʽ As Integer
    Dim n As Integer
    Dim int���� As Integer
    
    StrikeSave = False
    With vsfList
        Select Case mlngMode
            Case ģ���.�⹺���
                frmPurchaseCard.ShowCard Me, .TextMatrix(.Row, 0), �༭.����, vsfList.TextMatrix(vsfList.Row, �⹺����.��¼״̬), blnSuccess
                StrikeSave = blnSuccess
                Exit Function
            Case ģ���.�������
                mint����� = MediWork_GetCheckStockRule(cboStock.ItemData(cboStock.ListIndex))
                
                '�����������Ƿ��㹻����������Ϊ�������ʱ�����У����뵥�ݣ�ҩ��������飬���ݺţ���ţ�
                If mint����� <> 0 And .TextMatrix(.Row, 0) <> "" Then
                    For n = 1 To vsfDetail.rows - 1
                        If vsfDetail.TextMatrix(n, 0) <> "" Then
                            If CheckStrickUsable(���ݺ�.�������, 0, 0, vsfDetail.TextMatrix(n, vsfDetail.ColIndex("ҩƷ��Ϣ")), _
                                0, vsfDetail.TextMatrix(n, vsfDetail.ColIndex("����")), mint�����, Trim(.TextMatrix(.Row, 0)), vsfDetail.TextMatrix(n, vsfDetail.ColIndex("���"))) = False Then
                                Exit Function
                            End If
                        End If
                    Next
                End If
                gstrSQL = "zl_�������_Strike('" & .TextMatrix(.Row, 0) & "','" & UserInfo.�û����� & "')"
            Case ģ���.�������
                frmOtherInputCard.ShowCard Me, .TextMatrix(.Row, 0), �༭.����, vsfList.TextMatrix(vsfList.Row, vsfList.Cols - 2), blnSuccess
                StrikeSave = blnSuccess
                Exit Function
            Case ģ���.��۵���
                If Val(.TextMatrix(.Row, .Cols - 1)) = 1 Then
                    If Check�Ѹ����¼(.TextMatrix(.Row, 0)) Then
                        MsgBox "ҩƷ�Ѹ�����ܳ�����", vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
                gstrSQL = "zl_ҩƷ����۵���_Strike('" & .TextMatrix(.Row, 0) & "','" & UserInfo.�û����� & "')"
            Case ģ���.ҩƷ�ƿ�
                If mnuEditStrike.Caption = "�������(&R)" Then
                    int����ʽ = 1
                ElseIf mnuEditStrike.Caption = "��˳���(&K)" Then
                    int����ʽ = 2
                Else
                    int����ʽ = 0
                End If
               
                frmTransferCard.ShowCard Me, .TextMatrix(.Row, 0), �༭.����, vsfList.TextMatrix(vsfList.Row, vsfList.Cols - 2), blnSuccess, int����ʽ
                StrikeSave = blnSuccess
                Exit Function
            Case ģ���.ҩƷ����
                If mnuEditApplyStrike.Tag = "1" Then
                    int����ʽ = 1
                    mnuEditApplyStrike.Tag = "0"
                ElseIf mnuEditVerifyStrike.Tag = "1" Then
                    int����ʽ = 2
                    mnuEditVerifyStrike.Tag = "0"
                Else
                    int����ʽ = 0
                End If
            
                frmDrawCard.ShowCard Me, .TextMatrix(.Row, 0), �༭.����, mblnStock, vsfList.TextMatrix(vsfList.Row, vsfList.Cols - 4), 0, blnSuccess, int����ʽ
                StrikeSave = blnSuccess
                Exit Function
            Case ģ���.��������
                frmOtherOutputCard.ShowCard Me, .TextMatrix(.Row, 0), �༭.����, vsfList.TextMatrix(vsfList.Row, vsfList.Cols - 2), blnSuccess
                StrikeSave = blnSuccess
                Exit Function
            Case Else
            
        End Select
        
        On Error GoTo errHandle
        
        Call zlDataBase.ExecuteProcedure(gstrSQL, Me.Caption)
        
        If mlngMode = ģ���.������� Or mlngMode = ģ���.��۵��� Then
            If mlngMode = ģ���.������� Then
                int���� = ���ݺ�.�������
            ElseIf mlngMode = ģ���.��۵��� Then
                int���� = ���ݺ�.��۵���
            End If
            '��ʾͣ��ҩƷ
            Call CheckStopMedi(int���� & "|" & .TextMatrix(.Row, 0))
        End If
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
    Dim blnSuccess As Boolean
    
    blnSuccess = False
    With vsfList
        If cboStock.ListIndex = -1 Then
            MsgBox "��ѡ��ⷿ��", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        strNo = .TextMatrix(.Row, 0)
        
        Select Case mlngMode
            Case ģ���.�⹺���
                frmPurchaseCard.ShowCard Me, strNo, �༭.�޸�, vsfList.TextMatrix(vsfList.Row, �⹺����.��¼״̬), blnSuccess
            Case ģ���.�������
                frmSelfMakeCard.ShowCard Me, strNo, �༭.�޸�, vsfList.TextMatrix(vsfList.Row, vsfList.Cols - 2), blnSuccess
            Case ģ���.�������
                frmOtherInputCard.ShowCard Me, strNo, �༭.�޸�, vsfList.TextMatrix(vsfList.Row, vsfList.Cols - 2), blnSuccess
            Case ģ���.��۵���
                frmDiffPriceAdjustCard.ShowCard Me, strNo, �༭.�޸�, vsfList.TextMatrix(vsfList.Row, vsfList.Cols - 3), blnSuccess, IIf(Val(vsfList.TextMatrix(vsfList.Row, vsfList.Cols - 1)) = 0, 1, 2)
            Case ģ���.ҩƷ�ƿ�
                '�ѱ�ҩ����д����ҩ�ˣ����ѷ��͵ĵ��ݣ���������ⷽ�޸Ĵ��൥��
                If TabShow.Tab = 1 Then
                    If TestPrepare(strNo) Then
                        MsgBox "�ѷ��͵ĵ��ݲ������޸ģ�", vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
                frmTransferCard.ShowCard Me, strNo, �༭.�޸�, vsfList.TextMatrix(vsfList.Row, vsfList.Cols - 2), blnSuccess
            Case ģ���.ҩƷ����
                frmDrawCard.ShowCard Me, strNo, �༭.�޸�, mblnStock, vsfList.TextMatrix(vsfList.Row, vsfList.Cols - 4), 0, blnSuccess
            Case ģ���.��������
                frmOtherOutputCard.ShowCard Me, strNo, �༭.�޸�, vsfList.TextMatrix(vsfList.Row, vsfList.Cols - 2), blnSuccess
        End Select
        If blnSuccess = True Then
            mintListRow = vsfList.Row
            mnuViewRefresh_Click
        End If
    End With
End Sub

Private Sub mnuEditVerifySelect_Click()
    frmPurchaseVerifySelect.ShowME Me, mStr�ⷿ, cboStock.ListIndex
End Sub

Private Sub mnuEditVerifyStrike_Click()
    mnuEditVerifyStrike.Tag = "1"
    Call mnuEditStrike_Click
End Sub

Private Sub mnuEditWriteOff_Click()
    Dim strStock As String
    Dim i As Integer
    
    
    If cboStock.ListIndex = -1 Then
        MsgBox "��ѡ��ⷿ��", vbInformation, gstrSysName
        Exit Sub
    End If
    
    With Me.cboStock
        For i = 0 To .ListCount - 1
            strStock = strStock & .List(i) & "," & .ItemData(i) & "|"
        Next
    End With
    
    Call frm��������.ShowME(mlngMode, Me, strStock, Me.cboStock.ListIndex)
End Sub


Private Sub mnuFileAllCodePrint_Click()
    If Trim(vsfList.TextMatrix(vsfList.Row, 0)) = "" Or vsfList.rows <= 1 Then Exit Sub
    CodePrint vsfList.TextMatrix(vsfList.Row, 0)
End Sub

Private Sub mnuEditAllCodePrint_Click()
    If Trim(vsfList.TextMatrix(vsfList.Row, 0)) = "" Or vsfList.rows <= 1 Then Exit Sub
    CodePrint vsfList.TextMatrix(vsfList.Row, 0)
End Sub

Private Sub mnuFileSelCodePrint_Click()
    If Trim(vsfList.TextMatrix(vsfList.Row, 0)) = "" Or vsfList.rows <= 1 Then Exit Sub
    CodePrint Val(vsfDetail.TextMatrix(vsfDetail.Row, vsfDetail.ColIndex("ҩƷID")))
End Sub

Private Sub CodePrint(ByVal varPar As Variant)
'���ܣ���ӡҪƷ����
'���Σ�varPar��long�����ӡ��ӦҩƷ���룻��String������ݵ��ݺŴ�ӡ�����е�ҩƷ����
    Dim rsTemp As New ADODB.Recordset
    Dim int���� As Integer
    Dim strReport As String

    On Error GoTo errHandle
    
    If Not zlStr.IsHavePrivs(mstrPrivs, "ҩƷ�����ӡ") Then
        MsgBox "�Բ�����û�и�Ȩ�ޣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Select Case mlngMode
        Case ģ���.�⹺���
            int���� = 1
            strReport = "ZL1_INSIDE_1300_1"
        Case ģ���.�������
            int���� = 4
            strReport = "ZL1_INSIDE_1302_1"
        Case ģ���.��������
            int���� = 11
            strReport = "ZL1_INSIDE_1306_1"
        Case ģ���.ҩƷ�ƿ�
            int���� = 6
            strReport = "ZL1_INSIDE_1304_1"
        Case ģ���.ҩƷ����
            int���� = 7
            strReport = "ZL1_INSIDE_1305_2"
    End Select

    
    
    If TypeName(varPar) = "String" Then '��ӡ���ŵ�������
        gstrSQL = "select distinct ҩƷID from ҩƷ�շ���¼ where ���� = [2] and  NO = [1] order by ҩƷID"
        
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "ҩƷ�����ӡ", varPar, int����)
        
        Do While Not rsTemp.EOF
            ReportOpen gcnOracle, glngSys, strReport, Me, "ҩƷ=" & rsTemp!ҩƷid, 2
            rsTemp.MoveNext
        Loop
        
    Else '��ӡ��ӦҩƷ����
        If varPar = 0 Then Exit Sub
        ReportOpen gcnOracle, glngSys, strReport, Me, "ҩƷ=" & varPar, 2
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuFileBillPreview_Click()
    Dim int��λϵ�� As Integer
    Dim bln�˿ⵥ As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
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
        
        Select Case mlngMode
            Case ģ���.�⹺���
                '�ж��Ƿ����˿ⵥ
                gstrSQL = "Select Nvl(��ҩ��ʽ,0) ��־ From ҩƷ�շ���¼ Where NO=[1] And ��¼״̬=[2] And Rownum<2"
                Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[�ж��Ƿ����˿ⵥ]", .TextMatrix(.Row, 0), Val(.TextMatrix(.Row, �⹺����.��¼״̬)))
                
                bln�˿ⵥ = (rsTemp!��־ = 1)
                ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_" & ģ���.�⹺���, "zl8_bill_" & ģ���.�⹺���), Me, "���ݱ��=" & .TextMatrix(.Row, 0), "��¼״̬=" & .TextMatrix(.Row, �⹺����.��¼״̬), "��λϵ��=" & int��λϵ��, IIf(bln�˿ⵥ, "ҩƷ�˻���", "ҩƷ�⹺��ⵥ"), 1
            Case ģ���.�������
                ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_" & ģ���.�������, "zl8_bill_" & ģ���.�������), Me, "���ݱ��=" & .TextMatrix(.Row, 0), "��¼״̬=" & .TextMatrix(.Row, .Cols - 2), "��λϵ��=" & int��λϵ��, 1
            Case ģ���.�������
                ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_" & ģ���.�������, "zl8_bill_" & ģ���.�������), Me, "���ݱ��=" & .TextMatrix(.Row, 0), "��¼״̬=" & .TextMatrix(.Row, .Cols - 2), "��λϵ��=" & int��λϵ��, 1
            Case ģ���.��۵���
                ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_" & ģ���.��۵���, "zl8_bill_" & ģ���.��۵���), Me, "���ݱ��=" & .TextMatrix(.Row, 0), "��¼״̬=" & .TextMatrix(.Row, .Cols - 3), "��λϵ��=" & int��λϵ��, 1
            Case ģ���.ҩƷ�ƿ�
                ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_" & ģ���.ҩƷ�ƿ�, "zl8_bill_" & ģ���.ҩƷ�ƿ�), Me, "���ݱ��=" & .TextMatrix(.Row, 0), "��¼״̬=" & .TextMatrix(.Row, .Cols - 2), "��λϵ��=" & int��λϵ��, 1
            Case ģ���.ҩƷ����
                ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_" & ģ���.ҩƷ����, "zl8_bill_" & ģ���.ҩƷ����), Me, "���ݱ��=" & .TextMatrix(.Row, 0), "��¼״̬=" & .TextMatrix(.Row, .Cols - 4), "��λϵ��=" & int��λϵ��, 1
            Case ģ���.��������
                ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_" & ģ���.��������, "zl8_bill_" & ģ���.��������), Me, "���ݱ��=" & .TextMatrix(.Row, 0), "��¼״̬=" & .TextMatrix(.Row, .Cols - 2), "��λϵ��=" & int��λϵ��, 1
            Case Else
            
        End Select
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuFileBillPrint_Click()
    Dim int��λϵ�� As Integer
    Dim bln�˿ⵥ As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
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
        
        Select Case mlngMode
            Case ģ���.�⹺���
                '�ж��Ƿ����˿ⵥ
                gstrSQL = "Select Nvl(��ҩ��ʽ,0) ��־ From ҩƷ�շ���¼ Where NO=[1] And ��¼״̬=[2] And Rownum<2"
                Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[�ж��Ƿ����˿ⵥ]", .TextMatrix(.Row, 0), Val(.TextMatrix(.Row, .Cols - 3)))

                bln�˿ⵥ = (rsTemp!��־ = 1)
                ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_" & ģ���.�⹺���, "zl8_bill_" & ģ���.�⹺���), Me, "���ݱ��=" & .TextMatrix(.Row, 0), "��¼״̬=" & .TextMatrix(.Row, �⹺����.��¼״̬), "��λϵ��=" & int��λϵ��, IIf(bln�˿ⵥ, "ҩƷ�˻���", "ҩƷ�⹺��ⵥ"), 2
            Case ģ���.�������
                ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_" & ģ���.�������, "zl8_bill_" & ģ���.�������), Me, "���ݱ��=" & .TextMatrix(.Row, 0), "��¼״̬=" & .TextMatrix(.Row, .Cols - 2), "��λϵ��=" & int��λϵ��, 2
            Case ģ���.�������
                ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_" & ģ���.�������, "zl8_bill_" & ģ���.�������), Me, "���ݱ��=" & .TextMatrix(.Row, 0), "��¼״̬=" & .TextMatrix(.Row, .Cols - 2), "��λϵ��=" & int��λϵ��, 2
            Case ģ���.��۵���
                ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_" & ģ���.��۵���, "zl8_bill_" & ģ���.��۵���), Me, "���ݱ��=" & .TextMatrix(.Row, 0), "��¼״̬=" & .TextMatrix(.Row, .Cols - 3), "��λϵ��=" & int��λϵ��, 2
            Case ģ���.ҩƷ�ƿ�
                ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_" & ģ���.ҩƷ�ƿ�, "zl8_bill_" & ģ���.ҩƷ�ƿ�), Me, "���ݱ��=" & .TextMatrix(.Row, 0), "��¼״̬=" & .TextMatrix(.Row, .Cols - 2), "��λϵ��=" & int��λϵ��, 2
            Case ģ���.ҩƷ����
                ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_" & ģ���.ҩƷ����, "zl8_bill_" & ģ���.ҩƷ����), Me, "���ݱ��=" & .TextMatrix(.Row, 0), "��¼״̬=" & .TextMatrix(.Row, .Cols - 4), "��λϵ��=" & int��λϵ��, 2
            Case ģ���.��������
                ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_" & ģ���.��������, "zl8_bill_" & ģ���.��������), Me, "���ݱ��=" & .TextMatrix(.Row, 0), "��¼״̬=" & .TextMatrix(.Row, .Cols - 2), "��λϵ��=" & int��λϵ��, 2
            Case Else
            
        End Select
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
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
    Dim int��ѯ���� As Integer
    Dim dateCurrentDate As Date
    
    frm��������.���ò��� Me, mstrPrivs, Me.Tag
    
    Call GetDrugDigit(mlng�ⷿID, Me.Tag, mintUnit, mintShowCostDigit, mintShowPriceDigit, mintShowNumberDigit, mintShowMoneyDigit)
    
    '������֯��ʽ����
    mstrCostFormat = "'999999999990." & String(mintShowCostDigit, "0") & "'"
    mstrPriceFormat = "'999999999990." & String(mintShowPriceDigit, "0") & "'"
    mstrNumberFormat = "'999999999990." & String(mintShowNumberDigit, "0") & "'"
    mstrMoneyFormat = "'999999999990." & String(mintShowMoneyDigit, "0") & "'"
    
    int�ƿ⴦������ = Val(zlDataBase.GetPara("�ƿ�����", glngSys, ģ���.ҩƷ�ƿ�))
    mint�������� = Val(zlDataBase.GetPara("��������", glngSys, ģ���.ҩƷ�ƿ�))
    
    dateCurrentDate = Sys.Currentdate
    int��ѯ���� = Val(zlDataBase.GetPara("��ѯ����", glngSys, mlngMode, 1)) - 1
    strStart = Format(DateAdd("d", -int��ѯ����, dateCurrentDate), "yyyy-MM-dd")
    strEnd = Format(dateCurrentDate, "yyyy-MM-dd")
    
    SQLCondition.date����ʱ�俪ʼ = CDate(Format(strStart, "yyyy-mm-dd") & " 00:00:00")
    SQLCondition.date����ʱ����� = CDate(Format(strEnd, "yyyy-mm-dd") & " 23:59:59")
    
    Call SetMenu
    Call GetList(mstrFind)
End Sub

Private Sub mnuFilePreView_Click()
    '��ӡԤ��
    Dim lngCurRow As Long
    
    lngCurRow = vsfList.Row
    vsfList.Redraw = flexRDNone
    subPrint 2
    vsfList.Row = lngCurRow
    vsfList.Redraw = flexRDDirect
    vsfList.Col = 0
    vsfList.ColSel = vsfList.Cols - 1
    
End Sub

Private Sub mnuFilePrint_Click()
    '��ӡ
    Dim lngCurRow As Long
    
    lngCurRow = vsfList.Row
    vsfList.Redraw = flexRDNone
    subPrint 1
    vsfList.Row = lngCurRow
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
        Select Case mlngMode
            Case ģ���.�⹺���
                StrWinName = "frmMainList1"
            Case ģ���.�������
                StrWinName = "frmMainList2"
            Case ģ���.�������
                StrWinName = "frmMainList3"
            Case ģ���.��۵���
                StrWinName = "frmMainList4"
            Case ģ���.ҩƷ�ƿ�
                StrWinName = "frmMainList5"
            Case ģ���.ҩƷ����
                StrWinName = "frmMainList6"
            Case ģ���.��������
                StrWinName = "frmMainList7"
        End Select
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


Private Sub mnuPlugItem_Click(index As Integer)
    Call PlugInFun(mnuPlugItem(index).Tag)
End Sub

Private Sub mnuReportItem_Click(index As Integer)
    'Ĭ�ϲ�����ģ���.�⹺���(�⹺���)- ҩƷ=ҩƷid���ⷿ=�ⷿid����Ӧ��=��Ӧ��id������=�������ƣ�NO=��ⵥNO
    '          ģ���.�������(�������)- ҩƷ=ҩƷid���ⷿ=�ⷿid����ʼʱ��=���ƿ�ʼʱ�䣬����ʱ��=���ƽ���ʱ�䣬NO=��ⵥNO
    '          ģ���.�������(�������)- ҩƷ=ҩƷid���ⷿ=�ⷿid������=�������ƣ���ʼʱ��=���ƿ�ʼʱ�䣬����ʱ��=���ƽ���ʱ�䣬NO=��ⵥNO
    '          ģ���.��۵���(����۵���)- ҩƷ=ҩƷid���ⷿ=�ⷿid����ʼʱ��=���ƿ�ʼʱ�䣬����ʱ��=���ƽ���ʱ�䣬NO=��������NO
    '          ģ���.ҩƷ�ƿ�(ҩƷ�ƿ�)- ҩƷ=ҩƷid���ⷿ=�Ƴ��ⷿid������ⷿ=����ⷿid����ʼʱ��=���ƿ�ʼʱ�䣬����ʱ��=���ƽ���ʱ�䣬NO=�ƿⵥNO
    '          ģ���.ҩƷ����(ҩƷ����)- ҩƷ=ҩƷid���ⷿ=�ⷿid����ʼʱ��=���ƿ�ʼʱ�䣬����ʱ��=���ƽ���ʱ�䣬NO=���õ�NO
    '          ģ���.��������(��������)- ҩƷ=ҩƷid���ⷿ=�ⷿid����ʼʱ��=���ƿ�ʼʱ�䣬����ʱ��=���ƽ���ʱ�䣬NO=���ⵥNO
    Dim str��ʼʱ�� As String
    Dim str����ʱ�� As String
    Dim strNo As String
    Dim strReportName As String
    
    strReportName = Split(mnuReportItem(index).Tag, ",")(1)
    
    If strReportName = "ZL1_INSIDE_ģ���.ҩƷ����_1" Then
        ReportOpen gcnOracle, glngSys, "ZL1_INSIDE_ģ���.ҩƷ����_1", Me, "�ڼ�=" & Format(Sys.Currentdate, "YYYY"), "�ⷿ=" & cboStock.Text & "|" & cboStock.ItemData(cboStock.ListIndex), "��λ=סԺ��λ" & "|4"
    Else
        If vsfList.Row >= 1 And LTrim(vsfList.TextMatrix(vsfList.Row, 0)) <> "" Then
            strNo = vsfList.TextMatrix(vsfList.Row, 0)
        End If
        
        str��ʼʱ�� = IIf(Format(SQLCondition.date����ʱ�俪ʼ, "yyyy-mm-dd") = "1899-12-30", "", Format(SQLCondition.date����ʱ�俪ʼ, "yyyy-mm-dd"))
        str����ʱ�� = IIf(Format(SQLCondition.date����ʱ�����, "yyyy-mm-dd") = "1899-12-30", "", Format(SQLCondition.date����ʱ�����, "yyyy-mm-dd"))
            
        Call ReportOpen(gcnOracle, Split(mnuReportItem(index).Tag, ",")(0), Split(mnuReportItem(index).Tag, ",")(1), Me, _
            "ҩƷ=" & IIf(SQLCondition.lngҩƷ = 0, "", SQLCondition.lngҩƷ), _
            "�ⷿ=" & IIf(Val(cboStock.ItemData(cboStock.ListIndex)) = 0, "", Val(cboStock.ItemData(cboStock.ListIndex))), _
            "����ⷿ=" & IIf(SQLCondition.lng�ⷿ = 0, "", SQLCondition.lng�ⷿ), _
            "��Ӧ��=" & IIf(SQLCondition.lng������ = 0, "", SQLCondition.lng������), _
            "����=" & SQLCondition.str����, _
            "��ʼʱ��=" & str��ʼʱ��, _
            "����ʱ��=" & str����ʱ��, _
            "NO=" & strNo)
    End If
End Sub
Private Sub mnuViewColDefine_Click()
    Dim strColumn_All As String, strColumn_Select As String, strColumn_UnSelect As String
    Dim str��ѡ�� As String
    Dim str������ As String 'Ĭ��������
    Dim strAllCol As String
    Dim arr����, arr������
    Dim strChange As String
    Dim strOldColName As String, strNewColName As String
    Dim intCol As Integer
    
    On Error Resume Next
    
    Select Case mlngMode
    Case ģ���.�⹺���           'ҩƷ�⹺������
        strColumn_All = "ҩ��,0|ҩƷ��Դ,1|����ҩ��,1|ҩ�ۼ���,1|���,1|������,0|ԭ����,1|����,0|��������,1|Ч��,0|��λ,1|����,0|ָ��������,1|�ɹ���,1|����,1|" & _
                        "�ɱ���,0|�ɱ����,0|�ӳ���,1|�ۼ�,0|�ۼ۽��,0|���,0|���ۼ�,1|���۵�λ,1|���۽��,1|���۲��,1|��׼�ĺ�,1|���,1|" & _
                        "��Ʒ�ϸ�֤,1|�������,1|�������,1|���ս���,1|��Ʊ��,0|��Ʊ����,0|��Ʊ����,0|��Ʊ���,0"
        str��ѡ�� = "ҩ��|ҩƷ��Դ|����ҩ��|ҩ�ۼ���|���|������|ԭ����|����|��������|Ч��|��λ|����|ָ��������|�ɹ���|����|" & _
                        "�ɱ���|�ɱ����|�ӳ���|�ۼ�|�ۼ۽��|���|���ۼ�|���۵�λ|���۽��|���۲��|��׼�ĺ�|���|��Ʒ�ϸ�֤|�������|�������|���ս���|��Ʊ��|��Ʊ����|��Ʊ����|��Ʊ���"
        str������ = "���ۼ�|���۵�λ|���۽��|���۲��"
    Case ģ���.�������           'ҩƷ����������
    Case ģ���.�������           'ҩƷ����������
        strColumn_All = "ҩ��,0|ҩƷ��Դ,1|����ҩ��,1|���,1|������,0|ԭ����,1|����,0|��������,1|Ч��,0|��λ,1|����,0|��������,0|�ɱ���,1|�ɱ����,1|" & _
                        "�ۼ�,0|�ۼ۽��,0|���,0|���ۼ�,1|���۵�λ,1|���۽��,1|���۲��,1|��׼�ĺ�,1|���,1"
        str��ѡ�� = "ҩ��|ҩƷ��Դ|����ҩ��|���|������|ԭ����|����|��������|Ч��|��λ|����|��������|�ɱ���|�ɱ����|" & _
                        "�ۼ�|�ۼ۽��|���|���ۼ�|���۵�λ|���۽��|���۲��|��׼�ĺ�|���"
        str������ = "���ۼ�|���۵�λ|���۽��|���۲��"
    Case ģ���.��۵���           '����۵�������
    Case ģ���.ҩƷ�ƿ�           'ҩƷ�ƿ����
    Case ģ���.ҩƷ����           'ҩƷ���ù���
    Case ģ���.��������           'ҩƷ�����������
    End Select
    
    'ȡ��ѡ���е���Ϣ'Me.Caption
    strColumn_Select = zlDataBase.GetPara("ѡ����", glngSys, mlngMode, "")
    strColumn_UnSelect = zlDataBase.GetPara("������", glngSys, mlngMode, "")

    '�⹺��⡢�������Ĭ����"���ۼ�|���۵�λ|���۽��|���۲��"�⼸��
    If mlngMode = ģ���.�⹺��� Or mlngMode = ģ���.������� Then
        If strColumn_Select <> "" Then
            '�����ϰ汾���������Ʊ仯����ʽ��������,������|������,������...
            strChange = "����,������|�����,�ɱ���|������,�ɱ����"
        
            For intCol = 0 To UBound(Split(strChange, "|"))
                strOldColName = Split(Split(strChange, "|")(intCol), ",")(0)
                strNewColName = Split(Split(strChange, "|")(intCol), ",")(1)
                        
                If InStr(1, "|" & strColumn_Select & "|", "|" & strOldColName & "|") <> 0 Then
                    strColumn_Select = Replace("|" & strColumn_Select & "|", "|" & strOldColName & "|", "|" & strNewColName & "|")
                    strColumn_Select = Left(strColumn_Select, Len(strColumn_Select) - 1)
                    strColumn_Select = Mid(strColumn_Select, 2)
                End If
                
                If InStr(1, "|" & strColumn_UnSelect & "|", "|" & strOldColName & "|") <> 0 Then
                    strColumn_UnSelect = Replace("|" & strColumn_UnSelect & "|", "|" & strOldColName & "|", "|" & strNewColName & "|")
                    strColumn_UnSelect = Left(strColumn_UnSelect, Len(strColumn_UnSelect) - 1)
                    strColumn_UnSelect = Mid(strColumn_UnSelect, 2)
                End If
            Next
            
            If strColumn_UnSelect <> "" Then
                strAllCol = strColumn_Select & "|" & strColumn_UnSelect
            Else
                strAllCol = strColumn_Select
            End If
            arr���� = Split(str��ѡ��, "|")
            arr������ = Split(strAllCol, "|")
            
            If UBound(arr����) <> UBound(arr������) Or InStr(1, "|" & strColumn_Select & "|", "|������|") = 0 Or InStr(1, "|" & strColumn_UnSelect & "|", "|������|") <> 0 Or (mlngMode = ģ���.������� And (InStr(1, "|" & strColumn_Select & "|", "|�ɹ���|") <> 0 Or InStr(1, "|" & strColumn_UnSelect & "|", "|�ɹ���|") <> 0)) Then
                Select Case mlngMode
                Case ģ���.�⹺���
                    strColumn_Select = "ҩ��|ҩƷ��Դ|����ҩ��|ҩ�ۼ���|���|������|ԭ����|����|��������|Ч��|��λ|����|ָ��������|�ɹ���|����|�ɱ���|�ɱ����|�ӳ���|�ۼ�|�ۼ۽��|���|��׼�ĺ�|���|��Ʒ�ϸ�֤|�������|�������|��Ʊ��|��Ʊ����|��Ʊ����|��Ʊ���"
                    strColumn_UnSelect = "���ۼ�|���۵�λ|���۽��|���۲��"
                    zlDataBase.SetPara "ѡ����", strColumn_Select, glngSys, mlngMode
                    zlDataBase.SetPara "������", strColumn_UnSelect, glngSys, mlngMode
                Case ģ���.�������
                    strColumn_Select = "ҩ��|ҩƷ��Դ|����ҩ��|���|������|ԭ����|����|��������|Ч��|��λ|����|��������|�ɱ���|�ɱ����|�ۼ�|�ۼ۽��|���|��׼�ĺ�|���"
                    strColumn_UnSelect = "���ۼ�|���۵�λ|���۽��|���۲��"
                    zlDataBase.SetPara "ѡ����", strColumn_Select, glngSys, mlngMode
                    zlDataBase.SetPara "������", strColumn_UnSelect, glngSys, mlngMode
                End Select
            End If
        Else
            Select Case mlngMode
            Case ģ���.�⹺���
                strColumn_Select = "ҩ��|ҩƷ��Դ|����ҩ��|ҩ�ۼ���|���|������|ԭ����|����|��������|Ч��|��λ|����|ָ��������|�ɹ���|����|�ɱ���|�ɱ����|�ӳ���|�ۼ�|�ۼ۽��|���|��׼�ĺ�|���|��Ʒ�ϸ�֤|�������|�������|��Ʊ��|��Ʊ����|��Ʊ����|��Ʊ���"
                strColumn_UnSelect = "���ۼ�|���۵�λ|���۽��|���۲��"
                zlDataBase.SetPara "ѡ����", strColumn_Select, glngSys, mlngMode
                zlDataBase.SetPara "������", strColumn_UnSelect, glngSys, mlngMode
            Case ģ���.�������
                strColumn_Select = "ҩ��|ҩƷ��Դ|����ҩ��|���|������|ԭ����|����|��������|Ч��|��λ|����|��������|�ɱ���|�ɱ����|�ۼ�|�ۼ۽��|���|��׼�ĺ�|���"
                strColumn_UnSelect = "���ۼ�|���۵�λ|���۽��|���۲��"
                zlDataBase.SetPara "ѡ����", strColumn_Select, glngSys, mlngMode
                zlDataBase.SetPara "������", strColumn_UnSelect, glngSys, mlngMode
            End Select
        End If
    End If
    
    If Not frm������.ShowME(Me, strColumn_All, strColumn_Select) Then Exit Sub
    
    zlDataBase.SetPara "ѡ����", Split(strColumn_Select, "||")(0), glngSys, mlngMode
    zlDataBase.SetPara "������", Split(strColumn_Select, "||")(1), glngSys, mlngMode
End Sub

Private Sub mnuViewRefresh_Click()
    'ˢ��
    If cboStock.ListIndex = -1 Then
        MsgBox "��ѡ��ⷿ��", vbInformation, gstrSysName
        Exit Sub
    End If
    
    GetList mstrFind
End Sub

Private Sub mnuViewSearch_Click()
    '����
    Dim strFind As String
    
    If cboStock.ListIndex = -1 Then
        MsgBox "��ѡ��ⷿ��", vbInformation, gstrSysName
        Exit Sub
    End If
    Select Case mlngMode
        Case ģ���.��۵���, ģ���.ҩƷ�ƿ�, ģ���.ҩƷ����, ģ���.��������
            FrmTransferSearch.In_������� = IIf(TabShow.Tab = 0, -1, 1)
            strFind = FrmTransferSearch.GetSearch(Me, mlngMode, mlng�ⷿID, strStart, strEnd, strVerifyStart, strVerifyEnd, _
                SQLCondition.strNO��ʼ, _
                SQLCondition.strNO����, _
                SQLCondition.date����ʱ�俪ʼ, _
                SQLCondition.date����ʱ�����, _
                SQLCondition.date���ʱ�俪ʼ, _
                SQLCondition.date���ʱ�����, _
                SQLCondition.lngҩƷ, _
                SQLCondition.lng�ⷿ, _
                SQLCondition.str������, _
                SQLCondition.str�����, _
                SQLCondition.lngҩƷ����, _
                SQLCondition.str����, _
                SQLCondition.int�������һ����ѯ)
        Case ģ���.�⹺���
            strFind = FrmPurchaseSearch.GetSearch(Me, strStart, strEnd, strVerifyStart, strVerifyEnd, _
                SQLCondition.strNO��ʼ, _
                SQLCondition.strNO����, _
                SQLCondition.date����ʱ�俪ʼ, _
                SQLCondition.date����ʱ�����, _
                SQLCondition.date���ʱ�俪ʼ, _
                SQLCondition.date���ʱ�����, _
                SQLCondition.lngҩƷ, _
                SQLCondition.str������, _
                SQLCondition.str�����, _
                SQLCondition.lng������, _
                SQLCondition.str����, _
                SQLCondition.str��Ʊ�ſ�ʼ, _
                SQLCondition.str��Ʊ�Ž���, _
                SQLCondition.lngҩƷ����, _
                SQLCondition.str����, _
                SQLCondition.date��Ʊ������ڿ�ʼ, _
                SQLCondition.date��Ʊ������ڽ���, _
                SQLCondition.int�ޱ��, _
                SQLCondition.int�б��, _
                SQLCondition.int�޷�Ʊ, _
                SQLCondition.int�з�Ʊ, _
                SQLCondition.int�������һ����ѯ)
                
'                Call FrmPurchaseSearch.GetInfo(SQLCondition.int�ޱ��, SQLCondition.int�б��, SQLCondition.int�޷�Ʊ, SQLCondition.int�з�Ʊ)
        Case ģ���.�������
            strFind = FrmSelfMakeSearch.GetSearch(Me, strStart, strEnd, strVerifyStart, strVerifyEnd, _
                SQLCondition.strNO��ʼ, _
                SQLCondition.strNO����, _
                SQLCondition.date����ʱ�俪ʼ, _
                SQLCondition.date����ʱ�����, _
                SQLCondition.date���ʱ�俪ʼ, _
                SQLCondition.date���ʱ�����, _
                SQLCondition.lngҩƷ, _
                SQLCondition.str������, _
                SQLCondition.str�����)
        Case ģ���.�������
            strFind = FrmOtherInputSearch.GetSearch(Me, strStart, strEnd, strVerifyStart, strVerifyEnd, _
                SQLCondition.strNO��ʼ, _
                SQLCondition.strNO����, _
                SQLCondition.date����ʱ�俪ʼ, _
                SQLCondition.date����ʱ�����, _
                SQLCondition.date���ʱ�俪ʼ, _
                SQLCondition.date���ʱ�����, _
                SQLCondition.lngҩƷ, _
                SQLCondition.str������, _
                SQLCondition.str�����, _
                SQLCondition.str����, _
                SQLCondition.lng������)
    End Select
    
    If strFind <> "" Or SQLCondition.int�������һ����ѯ = 1 Then
        mstrFind = strFind
        vsfList.rows = 1
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
    Call SetMenu
        
    cbrTool.Bands(1).MinHeight = tlbTool.Height
    
    Form_Resize
End Sub





Private Sub vsfDetail_GotFocus()
    Call SetGridFocus(vsfDetail, True)
End Sub

Private Sub vsfDetail_LostFocus()
    Call SetGridFocus(vsfDetail, False)
End Sub

Private Sub vsfDetail_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    If mnuFileCodePrint.Visible = False Then Exit Sub
    
    PopupMenu mnuFileCodePrint, 2
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
    Dim intBill As Integer                      '��������  �磺1���⹺��⣻2��
    Dim strUnit As String                       '��λ����:�����ﵥλ��סԺ��λ��
    Dim str��װϵ�� As String
    Dim strOrder As String
    Dim strCompare As String
    Dim strSqlЧ�� As String
    Dim n As Long
    Dim strSqlҩ�� As String
    Dim intCol As Integer
    Dim intRow As Integer
    Dim strSqlOrder As String
    Dim strTemp As String
    
    On Error GoTo errHandle
    
    If mblnBandEvent = True Then Exit Sub
    If mlastRow = vsfList.Row Then
        SetEnable
        Exit Sub
    End If
    mlastRow = vsfList.Row
    
    With vsfList
        .Redraw = flexRDNone
        .ForeColorSel = .Cell(flexcpForeColor, mlastRow, 1)
        .Redraw = flexRDDirect
    End With
    
    strOrder = zlDataBase.GetPara("����", glngSys, mlngMode)
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
    ElseIf strCompare = "3" Then
        ''���ⷿ��λ����
        strSqlOrder = "�ⷿ��λ"
    End If
    
    strSqlOrder = strSqlOrder & IIf(Right(strOrder, 1) = "0", " ASC", " DESC") & ",ҩƷ��Ϣ,���"

    If vsfList.Row >= 1 And LTrim(vsfList.TextMatrix(vsfList.Row, 0)) <> "" Then
        vsfList.Col = 0
        vsfList.ColSel = vsfList.Cols - 1
        
        vsfDetail.Redraw = flexRDNone
        
        Select Case mintUnit
            Case mconint�ۼ۵�λ
                strUnit = "F.���㵥λ"
                strUnitQuantity = "LTRIM(to_char(A.ʵ������," & mstrNumberFormat & ")) AS ����," _
                    & "F.���㵥λ AS ��λ,"
                str��װϵ�� = "1"
            Case mconint���ﵥλ
                strUnit = "B.���ﵥλ"
                strUnitQuantity = "LTRIM(to_char(A.ʵ������ / B.�����װ," & mstrNumberFormat & ")) AS ����," _
                    & "B.���ﵥλ AS ��λ,"
                str��װϵ�� = "B.�����װ"
            Case mconintסԺ��λ
                strUnit = "B.סԺ��λ"
                strUnitQuantity = "LTRIM(to_char(A.ʵ������ / B.סԺ��װ," & mstrNumberFormat & ")) AS ����," _
                    & "B.סԺ��λ AS ��λ,"
                str��װϵ�� = "B.סԺ��װ"
            Case mconintҩ�ⵥλ
                strUnit = "B.ҩ�ⵥλ"
                strUnitQuantity = "LTRIM(to_char(A.ʵ������ / B.ҩ���װ," & mstrNumberFormat & ")) AS ����," _
                    & "B.ҩ�ⵥλ AS ��λ,"
                str��װϵ�� = "B.ҩ���װ"
        End Select
        
        strSqlЧ�� = IIf(gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1, "TO_CHAR(A.Ч��-1,'YYYY-MM-DD') AS ��Ч����", "TO_CHAR(A.Ч��,'YYYY-MM-DD') AS ʧЧ��")
        
        If gintҩƷ������ʾ = 0 Then
            strSqlҩ�� = ",('['||F.����||']'||F.����) AS ҩƷ��Ϣ"
        ElseIf gintҩƷ������ʾ = 1 Then
            strSqlҩ�� = ",('['||F.����||']'||NVL(E.����,F.����)) AS ҩƷ��Ϣ"
        Else
            strSqlҩ�� = ",('['||F.����||']'||F.����) AS ҩƷ��Ϣ,E.���� As ��Ʒ��"
        End If
        
        Select Case mlngMode
            Case ģ���.�⹺���
                intBill = 1
                strTemp = ""
                
                If SQLCondition.int�б�� = 1 And SQLCondition.int�ޱ�� = 0 Then
                    strTemp = strTemp & " and c.�����־=1"
                End If
                If SQLCondition.int�ޱ�� = 1 And SQLCondition.int�б�� = 0 Then
                    strTemp = strTemp & " and (c.�����־=0 or c.�����־ is null)"
                End If
                If SQLCondition.int�޷�Ʊ = 1 And SQLCondition.int�з�Ʊ = 0 Then
                    strTemp = strTemp & " and c.��Ʊ�� is null"
                End If
                If SQLCondition.int�з�Ʊ = 1 And SQLCondition.int�޷�Ʊ = 0 Then
                    strTemp = strTemp & " and c.��Ʊ�� is not null"
                End If
'                gstrSQL = "SELECT * FROM (SELECT DISTINCT A.���,decode(c.�����־,Null,'δ���',0,'δ���','�ѱ��') �����־ " & strSqlҩ�� & ",B.ҩƷ��Դ,B.����ҩ��,F.���, A.����, A.����, " & strSqlЧ�� & " ," & _
'                    strUnitQuantity & _
'                    " LTRIM(TO_CHAR(A.�ɱ���*" & str��װϵ�� & "," & mstrCostFormat & ")) AS �ɱ���, LTRIM(TO_CHAR(A.�ɱ����," & mstrMoneyFormat & ")) AS �ɱ����," & _
'                    " DECODE(A.����, NULL, 0, A.����) AS ����, " & _
'                    " LTRIM(TO_CHAR(Decode(To_Number(Nvl(A.�÷�, 0)), 0, A.���ۼ�, (A.���۽�� - To_Number(Nvl(A.�÷�, 0))) / A.ʵ������)*" & str��װϵ�� & "," & mstrPriceFormat & ")) AS �ۼ� , " & _
'                    " LTRIM(TO_CHAR(A.���۽��- To_Number(Nvl(A.�÷�, 0))," & mstrMoneyFormat & "))  AS �ۼ۽��, LTRIM(TO_CHAR(A.���- To_Number(Nvl(A.�÷�, 0))," & mstrMoneyFormat & ")) AS ���," & _
'                    " A.��׼�ĺ�, C.��Ʊ�� ,TO_CHAR(C.��Ʊ����,'YYYY-MM-DD') AS ��Ʊ����,NVL(C.�������,'0') AS �������, " & _
'                    " LTRIM(TO_CHAR(Decode(C.��Ʊ��,Null,0,C.��Ʊ���)," & mstrMoneyFormat & ")) AS ��Ʊ���,B.�б�ҩƷ,B.���������,C.�������, " & _
'                    " Decode(Nvl(F.�Ƿ���, 0) * Nvl(A.����, 0), 0, '',LTRIM(TO_CHAR(A.���ۼ�," & mstrPriceFormat & "))) As ���ۼ�," & _
'                    " Decode(Nvl(F.�Ƿ���, 0) * Nvl(A.����, 0), 0, '',LTRIM(TO_CHAR(A.���۽��," & mstrMoneyFormat & ")))  AS ���۽��, " & _
'                    " Decode(Nvl(F.�Ƿ���, 0) * Nvl(A.����, 0), 0, '',LTRIM(TO_CHAR(A.���," & mstrMoneyFormat & "))) AS ���۲�� " & _
'                    " FROM ҩƷ�շ���¼ A, ҩƷ��� B,�շ���ĿĿ¼ F,�շ���Ŀ���� E ,Ӧ����¼ C " & _
'                    " WHERE  A.ҩƷID = B.ҩƷID AND B.ҩƷID=F.ID " & _
'                    " AND A.ID = C.�շ�ID (+) AND C.ϵͳ��ʶ(+)=1 AND C.��¼����(+)<>-1 " & _
'                    " AND B.ҩƷID=E.�շ�ϸĿID(+) AND E.����(+)=3 " & _
'                    " AND A.���� = [1] AND A.��¼״̬ = [3] " & _
'                    " AND A.NO =[2] " & strTemp & _
'                    " ) ORDER BY " & strSqlOrder
                gstrSQL = "SELECT * FROM (SELECT DISTINCT A.���,decode(c.�����־,Null,'δ���',0,'δ���','�ѱ��') �����־ " & strSqlҩ�� & ",B.ҩƷ��Դ,B.����ҩ��,F.���, A.���� as ������,A.ԭ����, A.����, " & strSqlЧ�� & " ," & _
                    strUnitQuantity & _
                    " LTRIM(TO_CHAR(A.�ɱ���*" & str��װϵ�� & "," & mstrCostFormat & ")) AS �ɱ���, LTRIM(TO_CHAR(A.�ɱ����," & mstrMoneyFormat & ")) AS �ɱ����," & _
                    " DECODE(A.����, NULL, 0, A.����) AS ����, " & _
                    " LTRIM(TO_CHAR(A.���ۼ�*" & str��װϵ�� & "," & mstrPriceFormat & ")) AS �ۼ� , " & _
                    " LTRIM(TO_CHAR(A.���۽��," & mstrMoneyFormat & "))  AS �ۼ۽��, LTRIM(TO_CHAR(A.���," & mstrMoneyFormat & ")) AS ���," & _
                    " A.��׼�ĺ�, C.��Ʊ�� ,c.��Ʊ����,TO_CHAR(C.��Ʊ����,'YYYY-MM-DD') AS ��Ʊ����,NVL(C.�������,'0') AS �������, " & _
                    " LTRIM(TO_CHAR(Decode(C.��Ʊ��,Null,decode(c.��Ʊ����,Null,0,c.��Ʊ���),C.��Ʊ���)," & mstrMoneyFormat & ")) AS ��Ʊ���,B.�б�ҩƷ,B.���������,C.�������,TO_CHAR(C.�������,'YYYY-MM-DD') AS �������, " & _
                    " Decode(Nvl(F.�Ƿ���, 0) * Nvl(A.����, 0), 0, '',LTRIM(TO_CHAR(A.���ۼ�," & mstrPriceFormat & "))) As ���ۼ�," & _
                    " Decode(Nvl(F.�Ƿ���, 0) * Nvl(A.����, 0), 0, '',LTRIM(TO_CHAR(A.���۽��," & mstrMoneyFormat & ")))  AS ���۽��, " & _
                    " Decode(Nvl(F.�Ƿ���, 0) * Nvl(A.����, 0), 0, '',LTRIM(TO_CHAR(A.���," & mstrMoneyFormat & "))) AS ���۲��,F.ID ҩƷID " & _
                    " FROM ҩƷ�շ���¼ A, ҩƷ��� B,�շ���ĿĿ¼ F,�շ���Ŀ���� E ,Ӧ����¼ C " & _
                    " WHERE  A.ҩƷID = B.ҩƷID AND B.ҩƷID=F.ID " & _
                    " AND A.ID = C.�շ�ID (+) AND C.ϵͳ��ʶ(+)=1 AND C.��¼����(+)=0 " & _
                    " AND B.ҩƷID=E.�շ�ϸĿID(+) AND E.����(+)=3 " & _
                    " AND A.���� = [1] AND A.��¼״̬ = [3] " & _
                    " AND A.NO =[2] " & strTemp & _
                    " ) ORDER BY " & strSqlOrder
            Case ģ���.�������
                intBill = 2
                gstrSQL = " SELECT * FROM (SELECT DISTINCT A.���" & strSqlҩ�� & ",B.ҩƷ��Դ,B.����ҩ��,F.���,A.����, " & strSqlЧ�� & "," & _
                    strUnitQuantity & _
                    " LTRIM(TO_CHAR(A.�ɱ���*" & str��װϵ�� & "," & mstrCostFormat & ")) AS �ɹ���," & _
                    " LTRIM(TO_CHAR (A.�ɱ����, " & mstrMoneyFormat & ")) AS �ɹ����," & _
                    " LTRIM(TO_CHAR (A.���ۼ�*" & str��װϵ�� & ", " & mstrPriceFormat & ")) AS �ۼ�," & _
                    " LTRIM(TO_CHAR (A.���۽��, " & mstrMoneyFormat & ")) AS �ۼ۽��," & _
                    " LTRIM(TO_CHAR (A.���, " & mstrMoneyFormat & ")) AS ��� " & _
                    " FROM ҩƷ�շ���¼ A, ҩƷ��� B,�շ���ĿĿ¼ F, �շ���Ŀ���� E " & _
                    " WHERE A.ҩƷID = B.ҩƷID AND B.ҩƷID=F.ID " & _
                    " AND B.ҩƷID = E.�շ�ϸĿID (+) AND E.����(+)=3 " & _
                    " AND ��¼״̬ = [3] " & _
                    " AND A.���� = [1] AND ���ϵ��=1 " & _
                    " AND A.NO = [2] " & _
                    " ) ORDER BY " & strSqlOrder
            Case ģ���.�������
                intBill = 4
'                gstrSQL = "SELECT * FROM (SELECT DISTINCT A.���" & strSqlҩ�� & ",B.ҩƷ��Դ,B.����ҩ��," & _
'                    "F.���, A.����, A.����, " & strSqlЧ�� & "," & strUnitQuantity & _
'                    " LTRIM(TO_CHAR(A.�ɱ���*" & str��װϵ�� & "," & mstrCostFormat & ")) AS �ɱ���, LTRIM(TO_CHAR(A.�ɱ����," & mstrMoneyFormat & ")) AS �ɱ����," & _
'                    " LTRIM(TO_CHAR(Decode(To_Number(Nvl(A.�÷�, 0)), 0, A.���ۼ�, (A.���۽�� - To_Number(Nvl(A.�÷�, 0))) / A.ʵ������)*" & str��װϵ�� & "," & mstrPriceFormat & ")) AS �ۼ�," & _
'                    " LTRIM(TO_CHAR(A.���۽��- To_Number(Nvl(A.�÷�, 0))," & mstrMoneyFormat & "))  AS �ۼ۽��, LTRIM(TO_CHAR(A.���- To_Number(Nvl(A.�÷�, 0))," & mstrMoneyFormat & ")) AS ���, " & _
'                    " Decode(Nvl(F.�Ƿ���, 0) * Nvl(A.����, 0), 0, '',LTRIM(TO_CHAR(A.���ۼ�," & mstrPriceFormat & "))) As ���ۼ�," & _
'                    " Decode(Nvl(F.�Ƿ���, 0) * Nvl(A.����, 0), 0, '',LTRIM(TO_CHAR(A.���۽��," & mstrMoneyFormat & ")))  AS ���۽��, " & _
'                    " Decode(Nvl(F.�Ƿ���, 0) * Nvl(A.����, 0), 0, '',LTRIM(TO_CHAR(A.���," & mstrMoneyFormat & "))) AS ���۲�� " & _
'                    " FROM ҩƷ�շ���¼ A, ҩƷ��� B,�շ���ĿĿ¼ F,�շ���Ŀ���� E  " & _
'                    " WHERE  A.ҩƷID = B.ҩƷID AND B.ҩƷID=F.ID" & _
'                    " AND B.ҩƷID=E.�շ�ϸĿID(+) AND E.����(+)=3 " & _
'                    " AND ��¼״̬ = [3] " & _
'                    " AND A.���� = [1] " & _
'                    " AND A.NO =[2] " & _
'                    " ) ORDER BY " & strSqlOrder
                gstrSQL = "SELECT * FROM (SELECT DISTINCT A.���" & strSqlҩ�� & ",B.ҩƷ��Դ,B.����ҩ��," & _
                    "F.���, A.���� as ������,A.ԭ����, A.����, " & strSqlЧ�� & "," & strUnitQuantity & _
                    " LTRIM(TO_CHAR(A.�ɱ���*" & str��װϵ�� & "," & mstrCostFormat & ")) AS �ɱ���, LTRIM(TO_CHAR(A.�ɱ����," & mstrMoneyFormat & ")) AS �ɱ����," & _
                    " LTRIM(TO_CHAR(A.���ۼ�*" & str��װϵ�� & "," & mstrPriceFormat & ")) AS �ۼ�," & _
                    " LTRIM(TO_CHAR(A.���۽��," & mstrMoneyFormat & "))  AS �ۼ۽��, LTRIM(TO_CHAR(A.���," & mstrMoneyFormat & ")) AS ���, " & _
                    " Decode(Nvl(F.�Ƿ���, 0) * Nvl(A.����, 0), 0, '',LTRIM(TO_CHAR(A.���ۼ�," & mstrPriceFormat & "))) As ���ۼ�," & _
                    " Decode(Nvl(F.�Ƿ���, 0) * Nvl(A.����, 0), 0, '',LTRIM(TO_CHAR(A.���۽��," & mstrMoneyFormat & ")))  AS ���۽��, " & _
                    " Decode(Nvl(F.�Ƿ���, 0) * Nvl(A.����, 0), 0, '',LTRIM(TO_CHAR(A.���," & mstrMoneyFormat & "))) AS ���۲��,F.ID ҩƷID " & _
                    " FROM ҩƷ�շ���¼ A, ҩƷ��� B,�շ���ĿĿ¼ F,�շ���Ŀ���� E  " & _
                    " WHERE  A.ҩƷID = B.ҩƷID AND B.ҩƷID=F.ID" & _
                    " AND B.ҩƷID=E.�շ�ϸĿID(+) AND E.����(+)=3 " & _
                    " AND ��¼״̬ = [3] " & _
                    " AND A.���� = [1] " & _
                    " AND A.NO =[2] " & _
                    " ) ORDER BY " & strSqlOrder
            Case ģ���.��۵���
                intBill = 5
                gstrSQL = "SELECT * FROM (SELECT DISTINCT A.���" & strSqlҩ�� & ",B.ҩƷ��Դ,B.����ҩ��," & _
                    "F.���, A.���� as ������,A.ԭ����, A.����, " & strSqlЧ�� & "," & strUnit & _
                    " AS ��λ,LTRIM(TO_CHAR(A.���ۼ�," & mstrMoneyFormat & ")) AS �����,LTRIM(TO_CHAR(A.�ɱ���," & mstrMoneyFormat & ")) AS �����," & _
                    " LTRIM(TO_CHAR(A.���," & mstrMoneyFormat & "))  AS ������, " & _
                    " LTRIM(TO_CHAR(A.����*" & str��װϵ�� & "," & mstrCostFormat & ")) AS �³ɱ��� " & _
                    " FROM ҩƷ�շ���¼ A, ҩƷ��� B,�շ���ĿĿ¼ F,�շ���Ŀ���� E " & _
                    " WHERE A.ҩƷID = B.ҩƷID AND B.ҩƷID=F.ID " & _
                    " AND B.ҩƷID=E.�շ�ϸĿID(+) AND E.����(+)=3 " & _
                    " AND ��¼״̬ = [3] " & _
                    " AND A.���� =[1] " & _
                    " AND A.NO =[2] " & _
                    " ) ORDER BY " & strSqlOrder
                    
            Case ģ���.ҩƷ�ƿ�       'ҩƷ�ƿ����
                intBill = 6
                gstrSQL = " SELECT * FROM (SELECT DISTINCT A.���" & strSqlҩ�� & ",B.ҩƷ��Դ,B.����ҩ��,F.���,A.���� as ������,A.ԭ����, " & _
                    " A.����, " & strSqlЧ�� & ",LTRIM(TO_CHAR(A.��д���� /" & str��װϵ�� & "," & mstrNumberFormat & ")) AS ��д����," & _
                    " LTRIM(TO_CHAR(A.ʵ������ /" & str��װϵ�� & "," & mstrNumberFormat & ")) AS ʵ������," & strUnit & " AS ��λ," & _
                    " LTRIM(TO_CHAR (A.�ɱ���*" & str��װϵ�� & ", " & mstrCostFormat & ")) AS �ɱ���," & _
                    " LTRIM(TO_CHAR (A.�ɱ����, " & mstrMoneyFormat & ")) AS �ɱ����," & _
                    " LTRIM(TO_CHAR (A.���ۼ�*" & str��װϵ�� & ", " & mstrPriceFormat & ")) AS �ۼ�," & _
                    " LTRIM(TO_CHAR (A.���۽��, " & mstrMoneyFormat & ")) AS �ۼ۽��," & _
                    " LTRIM(TO_CHAR (A.���, " & mstrMoneyFormat & ")) AS ��� ,C.�ⷿ��λ,F.ID ҩƷID " & _
                    " FROM ҩƷ�շ���¼ A, ҩƷ��� B,�շ���ĿĿ¼ F,�շ���Ŀ���� E,ҩƷ�����޶� C " & _
                    " WHERE A.ҩƷID = B.ҩƷID AND B.ҩƷID=F.ID" & _
                    " AND B.ҩƷID = E.�շ�ϸĿID (+) AND E.����(+)=3 " & _
                    " AND ��¼״̬ = [3] " & _
                    " AND A.���� = [1] AND ���ϵ��=-1 " & _
                    " AND A.NO = [2] AND A.ҩƷID=C.ҩƷID(+) AND A.�ⷿID=C.�ⷿID(+)) " & _
                    " ORDER BY " & strSqlOrder
            Case ģ���.ҩƷ����
                intBill = 7
                gstrSQL = " SELECT * FROM (SELECT DISTINCT A.���" & strSqlҩ�� & ",B.ҩƷ��Դ,B.����ҩ��,F.���,A.���� as ������,A.ԭ����, " & _
                    " A.����, " & strSqlЧ�� & ",LTRIM(TO_CHAR(A.��д���� /" & str��װϵ�� & "," & mstrNumberFormat & ")) AS ��д����," & _
                    " LTRIM(TO_CHAR(A.ʵ������ /" & str��װϵ�� & "," & mstrNumberFormat & ")) AS ʵ������," & strUnit & " AS ��λ," & _
                    " LTRIM(TO_CHAR (A.�ɱ���*" & str��װϵ�� & ", " & mstrCostFormat & ")) AS �ɱ���," & _
                    " LTRIM(TO_CHAR (A.�ɱ����, " & mstrMoneyFormat & ")) AS �ɱ����," & _
                    " LTRIM(TO_CHAR (A.���ۼ�*" & str��װϵ�� & ", " & mstrPriceFormat & ")) AS �ۼ�," & _
                    " LTRIM(TO_CHAR (A.���۽��, " & mstrMoneyFormat & ")) AS �ۼ۽��," & _
                    " LTRIM(TO_CHAR (A.���, " & mstrMoneyFormat & ")) AS ��� ,C.�ⷿ��λ ,NVL(E.����,F.����) as ����,F.ID ҩƷID " & _
                    " FROM ҩƷ�շ���¼ A, ҩƷ��� B, �շ���ĿĿ¼ F,�շ���Ŀ���� E ,ҩƷ�����޶� C " & _
                    " WHERE A.ҩƷID = B.ҩƷID AND B.ҩƷID=F.ID " & _
                    " AND B.ҩƷID = E.�շ�ϸĿID (+) AND E.����(+)=3 " & _
                    " AND ��¼״̬ = [3] " & _
                    " AND A.���� = [1] " & _
                    " AND A.NO = [2] AND A.ҩƷID=C.ҩƷID(+) AND A.�ⷿID=C.�ⷿID(+))" & _
                    " ORDER BY " & strSqlOrder
            Case ģ���.��������
                intBill = 11
                gstrSQL = "SELECT * FROM (SELECT DISTINCT A.���" & strSqlҩ�� & ",B.ҩƷ��Դ,B.����ҩ��," & _
                        " F.���, A.���� as ������,A.ԭ����, A.����, " & strSqlЧ�� & "," & strUnitQuantity & _
                        " LTRIM(TO_CHAR(A.�ɱ���*" & str��װϵ�� & "," & mstrCostFormat & ")) AS �ɱ���, LTRIM(TO_CHAR(A.�ɱ����," & mstrMoneyFormat & ")) AS �ɱ����," & _
                        " LTRIM(TO_CHAR(A.���ۼ�*" & str��װϵ�� & "," & mstrPriceFormat & ")) AS �ۼ� , LTRIM(TO_CHAR(A.���۽��," & mstrMoneyFormat & "))  AS �ۼ۽��, LTRIM(TO_CHAR(A.���," & mstrMoneyFormat & ")) AS ���, " & _
                        " C.�ⷿ��λ ,NVL(E.����,F.����) as ���� ,F.ID ҩƷID, "
                    
                If vsfList.TextMatrix(vsfList.Row, 1) = "ҩƷ���" Then
                    gstrSQL = gstrSQL & " LTRIM(TO_CHAR(A.����*" & str��װϵ�� & "," & mstrPriceFormat & ")) AS �����,LTRIM(TO_CHAR(A.����*A.ʵ������," & mstrMoneyFormat & ")) AS ������,'' As ��ֵ˰��,'' As ˰�� "
                ElseIf vsfList.TextMatrix(vsfList.Row, 1) = "ҩƷ����" Then
                    gstrSQL = gstrSQL & " LTRIM(TO_CHAR(A.����*" & str��װϵ�� & "," & mstrPriceFormat & ")) AS ������,LTRIM(TO_CHAR(A.����*A.ʵ������," & mstrMoneyFormat & ")) AS �������,LTRIM(TO_CHAR(Nvl(A.Ƶ��,0)/100," & mstrMoneyFormat & ")) As ��ֵ˰��,LTRIM(TO_CHAR(A.����*A.ʵ������*(Nvl(A.Ƶ��,0)/100/(1+Nvl(A.Ƶ��,0)/100))," & mstrMoneyFormat & ")) As ˰�� "
                Else
                    gstrSQL = gstrSQL & " '' As �����,'' As ������,'' As ��ֵ˰��,'' As ˰�� "
                End If
                
                gstrSQL = gstrSQL & " FROM ҩƷ�շ���¼ A, ҩƷ��� B,�շ���ĿĿ¼ F,�շ���Ŀ���� E ,ҩƷ�����޶� C " & _
                    " WHERE  A.ҩƷID = B.ҩƷID AND B.ҩƷID=F.ID" & _
                    " AND B.ҩƷID=E.�շ�ϸĿID(+) AND E.����(+)=3 " & _
                    " AND ��¼״̬ = [3] " & _
                    " AND A.���� =[1] " & _
                    " AND A.NO =[2] AND A.ҩƷID=C.ҩƷID(+) AND A.�ⷿID=C.�ⷿID(+))" & _
                    " ORDER BY " & strSqlOrder
                
        End Select
        
        If mlngMode = ģ���.ҩƷ���� Then
            Set rsDetail = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption, intBill, vsfList.TextMatrix(vsfList.Row, 0), Val(vsfList.TextMatrix(vsfList.Row, vsfList.Cols - 4)))
        Else
            Set rsDetail = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption, intBill, vsfList.TextMatrix(vsfList.Row, 0), Val(vsfList.TextMatrix(vsfList.Row, vsfList.Cols - IIf(mlngMode = ģ���.��۵��� Or mlngMode = ģ���.�⹺���, 3, 2))))
        End If
        
        Set vsfDetail.DataSource = rsDetail
        With vsfDetail
            If rsDetail.RecordCount > 0 Then
                .Row = 1
            End If
            .Col = 0
            .ColSel = .Cols - 1
            
            For n = 0 To .Cols - 1
                .ColKey(n) = .TextMatrix(0, n)
                .FixedAlignment(n) = flexAlignCenterCenter
            Next
            
            .colHidden(.ColIndex("ҩƷID")) = True 'ҩƷID�в���ʾ
            
            '���¸��½�����ţ���ΪҩƷ�ƿ���2�����ݣ�����ֻ��ȡһ�������Ի����1��3��5��7����2��4��6��8�������Ҫ����ĳ�������
            If mlngMode = ģ���.ҩƷ�ƿ� Then
                For intRow = 0 To .rows - 1
                    If intRow <> 0 Then
                        .TextMatrix(intRow, vsfDetail.ColIndex("���")) = intRow
                    End If
                Next
            End If
        End With
        
        '�Ը��ֵ��ݵ�������ʽ��
        If rsDetail.RecordCount > 0 Then
            With vsfDetail
                Select Case mlngMode
                Case ģ���.�⹺���
                    For n = 1 To .rows - 1
                        .TextMatrix(n, .ColIndex("����")) = zlStr.FormatEx(.TextMatrix(n, .ColIndex("����")), mintShowNumberDigit, , True)
                    Next
                Case ģ���.�������
                    For n = 1 To .rows - 1
                        .TextMatrix(n, .ColIndex("����")) = zlStr.FormatEx(.TextMatrix(n, .ColIndex("����")), mintShowNumberDigit, , True)
                    Next
                Case ģ���.�������
                    For n = 1 To .rows - 1
                        .TextMatrix(n, .ColIndex("����")) = zlStr.FormatEx(.TextMatrix(n, .ColIndex("����")), mintShowNumberDigit, , True)
                    Next
                Case ģ���.ҩƷ�ƿ�
                    For n = 1 To .rows - 1
                        .TextMatrix(n, .ColIndex("��д����")) = zlStr.FormatEx(.TextMatrix(n, .ColIndex("��д����")), mintShowNumberDigit, , True)
                        .TextMatrix(n, .ColIndex("ʵ������")) = zlStr.FormatEx(.TextMatrix(n, .ColIndex("ʵ������")), mintShowNumberDigit, , True)
                    Next
                Case ģ���.ҩƷ����
                    For n = 1 To .rows - 1
                        .TextMatrix(n, .ColIndex("��д����")) = zlStr.FormatEx(.TextMatrix(n, .ColIndex("��д����")), mintShowNumberDigit, , True)
                        .TextMatrix(n, .ColIndex("ʵ������")) = zlStr.FormatEx(.TextMatrix(n, .ColIndex("ʵ������")), mintShowNumberDigit, , True)
                    Next
                Case ģ���.��������
                    For n = 1 To .rows - 1
                        .TextMatrix(n, .ColIndex("����")) = zlStr.FormatEx(.TextMatrix(n, .ColIndex("����")), mintShowNumberDigit, , True)
                    Next
                End Select
            End With
        End If
        
        vsfDetail.Redraw = flexRDDirect
    ElseIf LTrim(vsfList.TextMatrix(vsfList.Row, 0)) = "" Then
        With vsfDetail
            .Redraw = flexRDNone
            Select Case mlngMode
                Case ģ���.�⹺���
                    .Cols = IIf(gintҩƷ������ʾ = 2, 32, 31)
                    .rows = 2
                    .Clear
                    
                    intCol = 0
                    
                    .TextMatrix(0, intCol) = "���": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "�����־": intCol = intCol + 1
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
                    .TextMatrix(0, intCol) = IIf(gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1, "��Ч����", "ʧЧ��"): intCol = intCol + 1
                    .TextMatrix(0, intCol) = "����": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "��λ": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "�ɱ���": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "�ɱ����": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "����": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "�ۼ�": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "�ۼ۽��": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "���": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "��׼�ĺ�": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "��Ʊ��": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "��Ʊ����": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "��Ʊ����": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "�������": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "��Ʊ���": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "�б�ҩƷ": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "���������": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "�������": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "�������": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "���ۼ�": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "���۽��": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "���۲��": intCol = intCol + 1
                Case ģ���.�������
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
                    
                    .TextMatrix(0, intCol) = "����": intCol = intCol + 1
                    .TextMatrix(0, intCol) = IIf(gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1, "��Ч����", "ʧЧ��"): intCol = intCol + 1
                    .TextMatrix(0, intCol) = "����": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "��λ": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "�ɹ���": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "�ɹ����": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "�ۼ�": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "�ۼ۽��": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "���": intCol = intCol + 1
                Case ģ���.�������
                    .Cols = IIf(gintҩƷ������ʾ = 2, 20, 19)
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
                    .TextMatrix(0, intCol) = IIf(gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1, "��Ч����", "ʧЧ��"): intCol = intCol + 1
                    .TextMatrix(0, intCol) = "����": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "��λ": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "�ɱ���": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "�ɱ����": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "�ۼ�": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "�ۼ۽��": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "���": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "���ۼ�": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "���۽��": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "���۲��": intCol = intCol + 1
                Case ģ���.��۵���
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
                    .TextMatrix(0, intCol) = IIf(gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1, "��Ч����", "ʧЧ��"): intCol = intCol + 1
                    .TextMatrix(0, intCol) = "��λ": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "�����": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "�����": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "������": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "�³ɱ���": intCol = intCol + 1
                
                Case ģ���.ҩƷ�ƿ�
                    .Cols = IIf(gintҩƷ������ʾ = 2, 19, 18)
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
                    .TextMatrix(0, intCol) = IIf(gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1, "��Ч����", "ʧЧ��"): intCol = intCol + 1
                    .TextMatrix(0, intCol) = "��д����": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "ʵ������": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "��λ": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "�ɱ���": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "�ɱ����": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "�ۼ�": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "�ۼ۽��": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "���": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "�ⷿ��λ": intCol = intCol + 1
                Case ģ���.ҩƷ����
                    .Cols = IIf(gintҩƷ������ʾ = 2, 20, 19)
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
                    .TextMatrix(0, intCol) = IIf(gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1, "��Ч����", "ʧЧ��"): intCol = intCol + 1
                    .TextMatrix(0, intCol) = "��д����": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "ʵ������": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "��λ": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "�ɱ���": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "�ɱ����": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "�ۼ�": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "�ۼ۽��": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "���": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "�ⷿ��λ": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "����": intCol = intCol + 1
                Case ģ���.��������
                    .Cols = IIf(gintҩƷ������ʾ = 2, 23, 22)
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
                    .TextMatrix(0, intCol) = IIf(gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1, "��Ч����", "ʧЧ��"): intCol = intCol + 1
                    .TextMatrix(0, intCol) = "����": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "��λ": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "�ɱ���": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "�ɱ����": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "�ۼ�": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "�ۼ۽��": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "���": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "�ⷿ��λ": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "����": intCol = intCol + 1
                    
                    If vsfList.TextMatrix(vsfList.Row, 1) = "ҩƷ����" Then
                        .TextMatrix(0, intCol) = "������": intCol = intCol + 1
                        .TextMatrix(0, intCol) = "�������": intCol = intCol + 1
                    Else
                        .TextMatrix(0, intCol) = "�����": intCol = intCol + 1
                        .TextMatrix(0, intCol) = "������": intCol = intCol + 1
                    End If
                    
                    .TextMatrix(0, intCol) = "��ֵ˰��": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "˰��": intCol = intCol + 1
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
    SetEnable
    Call ShowColor(rsDetail)
    If mlngMode = ģ���.ҩƷ�ƿ� Then Call CheckNumber
    
    If mblnDo Then
        RestoreFlexState vsfDetail, App.ProductName & "\" & Me.Name & mstrTitle
    End If
    
    If vsfDetail.rows > 1 Then
        vsfDetail.Row = 1
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetListFocuse()
    Dim intStatus As Integer
    Dim lngForeColor As Long
    
    With vsfList
        .ForeColorFixed = glngFixedForeColorByFocus
        .BackColorSel = glngRowByFocus
    
'        If .Row > 0 Then
'            .ForeColorSel = .Cell(flexcpForeColor, .Row)
'        End If
    End With
    
    vsfDetail.ForeColorFixed = glngFixedForeColorNotFocus
    vsfDetail.BackColorSel = glngRowByNotFocus
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
    
    mnuEditAllCodePrint.Visible = True
    mnuEditAllCodePrint.Visible = mlngMode = 1300 Or mlngMode = 1302 Or mlngMode = 1304 Or mlngMode = 1305 Or mlngMode = 1306
    mnuEditCodePrintLine.Visible = mnuEditAllCodePrint.Visible
    PopupMenu mnuEdit, 2
    mnuEditAllCodePrint.Visible = False
    mnuEditCodePrintLine.Visible = mnuEditAllCodePrint.Visible
    
End Sub

Private Sub picSeparate_s_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    '�ָ�������
    
    If Button <> 1 Then Exit Sub
    With picSeparate_s
        If .Top + y < 2000 Then Exit Sub
        If .Top + y > ScaleHeight - 2000 Then Exit Sub
        .Move .Left, .Top + y
    End With
'    Call Form_Resize
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
    Me.Refresh
End Sub

Private Sub TabShow_Click(PreviousTab As Integer)
    If mlngMode <> ģ���.ҩƷ�ƿ� And mlngMode <> ģ���.ҩƷ���� Then Exit Sub
    
    lbl1.Caption = ""
    lbl2.Caption = ""
    lbl3.Caption = ""
    
    mintListRow = 1
    
    Call SetMenu
    Call GetList(mstrFind)
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
        Case "PreparePhysic"
            mnuEditPreparePhysic_Click
        Case "SendPhysic"
            mnuEditSendPhysic_Click
        Case "Back"
            mnuEditBack_Click
        Case "Verify"
            mnuEditVerify_Click
        Case "Strike"
            mnuEditStrike_Click
        Case "ApplyStrike"
            mnuEditApplyStrike_Click
        Case "VerifyStrike"
            mnuEditVerifyStrike_Click
        Case "Search"
            mnuViewSearch_Click
        Case "Refresh"
            mnuViewRefresh_Click
        Case "Help"
            mnuHelpTitle_Click
        Case "Exit"
            mnufileexit_Click
        Case Else
            'zlPlugIn��ҹ���
            If Button.Key Like "PlugItem*" Then
                Call PlugInFun(Button.Caption)
            End If
'        Case "Mark"
'            Call frmPurchaseMark.showMe(cboStock.ItemData(cboStock.ListIndex), Me)
    End Select
    
End Sub

Private Sub tlbTool_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Dim blnSuccess As Boolean
    Select Case ButtonMenu.Key
    Case "FromStore"
        frmDrawCard.ShowCard Me, "", �༭.����, mblnStock, , 0, blnSuccess
    Case "FromLeave"
        frmDrawCard.ShowCard Me, "", �༭.����, mblnStock, , 1, blnSuccess
    End Select
    
    If blnSuccess = True Then
        mintListRow = vsfList.Row + 1
        mnuViewRefresh_Click
    End If
End Sub
'���ò˵��͹��߰�ť�Ŀ�������
Private Sub SetEnable()
    Dim strVerify As String
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
            
            If mnuEditApplyStrike.Visible = True Then
                mnuEditApplyStrike.Enabled = False
                tlbTool.Buttons("ApplyStrike").Enabled = False
            End If
            
            If mnuEditVerifyStrike.Visible = True Then
                mnuEditVerifyStrike.Enabled = False
                tlbTool.Buttons("VerifyStrike").Enabled = False
            End If

            If mnuEditDisplay.Visible = True Then
                mnuEditDisplay.Enabled = False
            End If
            
            If mnuEditPrepare.Visible Then
                mnuEditPrepare.Enabled = False
                mnuEditBack.Enabled = False
                tlbTool.Buttons("Prepare").Enabled = False
                tlbTool.Buttons("Back").Enabled = False
            End If
            
            If mnuEditPreparePhysic.Visible Then
                mnuEditPreparePhysic.Enabled = False
                mnuEditSendPhysic.Enabled = False
                mnuEditBack.Enabled = False
                tlbTool.Buttons("PreparePhysic").Enabled = False
                tlbTool.Buttons("SendPhysic").Enabled = False
                tlbTool.Buttons("Back").Enabled = False
            End If
            
            If mnuEditBill.Visible = True Then
                mnuEditBill.Enabled = False
            End If
            If mnuEditAcc.Visible Then
                mnuEditAcc.Enabled = False
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
            If mnuEditAcc.Visible Then
                mnuEditAcc.Enabled = False
            End If
            
            If mlngMode = ģ���.ҩƷ�ƿ� Then
                If TabShow.Tab = 0 Then
                    strVerify = .TextMatrix(.Row, .Cols - 6)
                Else
                    strVerify = .TextMatrix(.Row, .Cols - 4)
                End If
            Else
                strVerify = .TextMatrix(.Row, .Cols - 4)
            End If
            
            If mlngMode = ģ���.ҩƷ���� Then strVerify = .TextMatrix(.Row, .Cols - 5)
            
            
            If strVerify = "" Then    'δ��˵�
                If mnuEditModify.Visible = True Then
                    mnuEditModify.Enabled = IIf(mlngMode = ģ���.ҩƷ�ƿ�, IIf(.TextMatrix(.Row, .Cols - 4) = "", True, False), True)
                    tlbTool.Buttons("Modify").Enabled = mnuEditModify.Enabled
                End If
                If mnuEditDel.Visible = True Then
                    mnuEditDel.Enabled = IIf(mlngMode = ģ���.ҩƷ�ƿ�, IIf(.TextMatrix(.Row, .Cols - 4) = "", True, False), True)
                    tlbTool.Buttons("Delete").Enabled = mnuEditDel.Enabled
                End If
                If mnuEditVerify.Visible = True Then
                    mnuEditVerify.Enabled = True
                    tlbTool.Buttons("Verify").Enabled = True
                End If
                If mnuEditStrike.Visible = True Then
                    mnuEditStrike.Enabled = False
                    tlbTool.Buttons("Strike").Enabled = False
                End If
                If mnuEditApplyStrike.Visible = True Then
                    mnuEditApplyStrike.Enabled = False
                    tlbTool.Buttons("ApplyStrike").Enabled = False
                End If
                If mnuEditVerifyStrike.Visible = True Then
                    mnuEditVerifyStrike.Enabled = False
                    tlbTool.Buttons("VerifyStrike").Enabled = False
                End If
                If mnuEditDisplay.Visible = True Then
                    mnuEditDisplay.Enabled = True
                End If
                
                'δ��˵ĵ��ݣ������ظ��˲�
                mnuEditPrepare.Enabled = True
                tlbTool.Buttons("Prepare").Enabled = True
                
                If mnuEditBack.Visible = True Then
                    If .TextMatrix(.Row, .Cols - 6) <> "" Then
                        mnuEditBack.Enabled = True
                        tlbTool.Buttons("Back").Enabled = True
                    Else
                        mnuEditBack.Enabled = False
                        tlbTool.Buttons("Back").Enabled = False
                    End If
                End If
                
                '������⹺��ⵥ������Ƿ�ͨ���˲飬δͨ���˲�ĵ��ݣ����������
                If mlngMode = ģ���.�⹺��� And mbln�˲� Then
                    mnuEditVerify.Enabled = TestPrepare(.TextMatrix(.Row, 0))
                    tlbTool.Buttons("Verify").Enabled = TestPrepare(.TextMatrix(.Row, 0))
                End If
                '�ƿⵥ�����ݵ�ǰѡ���ҳ�棬��ǰ�������ð�ť״̬
                If mlngMode = ģ���.ҩƷ�ƿ� Then
                    If TabShow.Tab = 0 Then
                        mnuEditPreparePhysic.Enabled = (.TextMatrix(.Row, .Cols - 4) = "")
                        mnuEditSendPhysic.Enabled = (.TextMatrix(.Row, .Cols - 4) <> "") And (.TextMatrix(.Row, .Cols - 3) = "")
                        mnuEditBack.Enabled = Not mnuEditPreparePhysic.Enabled
                        tlbTool.Buttons("PreparePhysic").Enabled = mnuEditPreparePhysic.Enabled
                        tlbTool.Buttons("SendPhysic").Enabled = mnuEditSendPhysic.Enabled
                        tlbTool.Buttons("Back").Enabled = mnuEditBack.Enabled
                        '����õ�������ˣ�������ҩ�뷢��
                        If TestVerify(vsfList.TextMatrix(vsfList.Row, 0)) Then
                            mnuEditPreparePhysic.Enabled = False
                            mnuEditSendPhysic.Enabled = False
                            mnuEditBack.Enabled = False
                            tlbTool.Buttons("PreparePhysic").Enabled = False
                            tlbTool.Buttons("SendPhysic").Enabled = False
                            tlbTool.Buttons("Back").Enabled = False
                        End If
                        
                        '�����������δ��ˣ���������˳�����
                        If mint�������� = 1 And Val(.TextMatrix(.Row, .Cols - 2)) Mod 3 = 2 Then
                            mnuEditPreparePhysic.Enabled = False
                            mnuEditSendPhysic.Enabled = False
                            mnuEditBack.Enabled = False
                            mnuEditDel.Enabled = False
                            tlbTool.Buttons("PreparePhysic").Enabled = False
                            tlbTool.Buttons("SendPhysic").Enabled = False
                            tlbTool.Buttons("Back").Enabled = False
                            tlbTool.Buttons("Delete").Enabled = False
                            
                            mnuEditStrike.Enabled = True
                            tlbTool.Buttons("Strike").Enabled = True
                            
                            mnuEditVerify.Enabled = False
                            tlbTool.Buttons("Verify").Enabled = mnuEditVerify.Enabled
                        End If
                    Else
                        If int�ƿ⴦������ = 1 Then
                            mnuEditVerify.Enabled = TestPrepare(.TextMatrix(.Row, 0))
                        Else
                            mnuEditVerify.Enabled = True
                        End If
                        tlbTool.Buttons("Verify").Enabled = mnuEditVerify.Enabled
                        
                        '�����������δ��ˣ�������ɾ��
                        If mint�������� = 1 And Val(.TextMatrix(.Row, .Cols - 2)) Mod 3 = 2 Then
                            mnuEditModify.Enabled = False
                            tlbTool.Buttons("Modify").Enabled = False
                            mnuEditVerify.Enabled = False
                            tlbTool.Buttons("Verify").Enabled = False
                            mnuEditStrike.Enabled = False
                            tlbTool.Buttons("Strike").Enabled = False
                            
                            mnuEditDel.Enabled = True
                            tlbTool.Buttons("Delete").Enabled = True
                        End If
                    End If
                End If
                If mlngMode = ģ���.ҩƷ���� Then
                     '�����������δ��ˣ���������˳�����
                    If mint���ó������� = 1 And Val(.TextMatrix(.Row, .Cols - 4)) Mod 3 = 2 Then
                        mnuEditDel.Enabled = True
                        tlbTool.Buttons("Delete").Enabled = True
                        mnuEditModify.Enabled = False
                        tlbTool.Buttons("Modify").Enabled = False
                        mnuEditVerify.Enabled = False
                        tlbTool.Buttons("Verify").Enabled = False
                        mnuEditApplyStrike.Enabled = False
                        tlbTool.Buttons("ApplyStrike").Enabled = False
                        mnuEditVerifyStrike.Enabled = True
                        tlbTool.Buttons("VerifyStrike").Enabled = True
                    End If
                End If
            ElseIf .TextMatrix(.Row, IIf(mlngMode = ģ���.��۵��� Or mlngMode = ģ���.�⹺��� Or mlngMode = ģ���.ҩƷ����, IIf(mlngMode = ģ���.ҩƷ����, .Cols - 4, .Cols - 3), .Cols - 2)) = 1 Then '��˵�
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
                
                '����۵����У�����ǳɱ��۵��������ܳ���
                If mlngMode = ģ���.��۵��� Then
                    If Val(vsfList.TextMatrix(vsfList.Row, vsfList.Cols - 1)) = 1 Then
                        mnuEditStrike.Enabled = False
                        tlbTool.Buttons("Strike").Enabled = False
                    End If
                End If
                
                If mlngMode = ģ���.ҩƷ���� Then
                    If mint���ó������� = 1 Then
                        mnuEditApplyStrike.Enabled = True
                        tlbTool.Buttons("ApplyStrike").Enabled = True
                        mnuEditVerifyStrike.Enabled = False
                        tlbTool.Buttons("VerifyStrike").Enabled = False
                    End If
                End If
                
                'ֻ���⹺��ⵥ����
                mnuEditBill.Enabled = True
                mnuEditAcc.Enabled = True
                If mnuEditPrepare.Visible Then
                    mnuEditPrepare.Enabled = False
                    mnuEditBack.Enabled = False
                    tlbTool.Buttons("Prepare").Enabled = False
                    tlbTool.Buttons("Back").Enabled = False
                End If
                If mlngMode = ģ���.ҩƷ�ƿ� Then
                    If TabShow.Tab = 0 Then
                        '����õ�������ˣ�������ҩ�뷢��
                        If TestVerify(vsfList.TextMatrix(vsfList.Row, 0)) Then
                            mnuEditPreparePhysic.Enabled = False
                            mnuEditSendPhysic.Enabled = False
                            mnuEditBack.Enabled = False
                            tlbTool.Buttons("PreparePhysic").Enabled = False
                            tlbTool.Buttons("SendPhysic").Enabled = False
                            tlbTool.Buttons("Back").Enabled = False
                        End If
                        If mint�������� = 1 Then
                            mnuEditStrike.Enabled = False
                            tlbTool.Buttons("Strike").Enabled = False
                        End If
                    End If
                End If
            Else   '2,3 ���������Ѹ���ĵ��ݲ����������ˣ�ͬ����������˺�ĵ��ݲ����������
                If .TextMatrix(.Row, IIf(mlngMode = ģ���.��۵��� Or mlngMode = ģ���.�⹺��� Or mlngMode = ģ���.ҩƷ����, IIf(mlngMode = ģ���.ҩƷ����, .Cols - 4, .Cols - 3), .Cols - 2)) Mod 3 = 0 Then
                    .ToolTipText = "�������ݵ�ԭ����"
                    If mnuEditStrike.Visible = True Then
                        mnuEditStrike.Enabled = True
                        tlbTool.Buttons("Strike").Enabled = True
                    End If
                    If mnuEditApplyStrike.Visible = True Then
                        mnuEditApplyStrike.Enabled = True
                        tlbTool.Buttons("ApplyStrike").Enabled = True
                    End If
                    '�����ֳ����ĵ��ݲ������
                    mnuEditAcc.Enabled = True
                    
                    '�����ֳ�����ԭ�����޸ķ�Ʊ��Ϣ
                    If mnuEditBill.Visible = True Then
                        mnuEditBill.Enabled = True
                    End If
                ElseIf .TextMatrix(.Row, IIf(mlngMode = ģ���.��۵��� Or mlngMode = ģ���.�⹺��� Or mlngMode = ģ���.ҩƷ����, IIf(mlngMode = ģ���.ҩƷ����, .Cols - 4, .Cols - 3), .Cols - 2)) Mod 3 = 2 Then
                    If mlngMode = ģ���.�⹺��� Then
                        If Val(.TextMatrix(.Row, �⹺����.��������)) = 1 Then
                            .ToolTipText = "������˳�������"
                        Else
                            .ToolTipText = "��������"
                        End If
                    Else
                        .ToolTipText = "��������"
                    End If
                    If mnuEditStrike.Visible = True Then
                        mnuEditStrike.Enabled = False
                        tlbTool.Buttons("Strike").Enabled = False
                    End If
                    If mnuEditApplyStrike.Visible = True Then
                        mnuEditApplyStrike.Enabled = False
                        tlbTool.Buttons("ApplyStrike").Enabled = False
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
                If mnuEditVerifyStrike.Visible = True Then
                    mnuEditVerifyStrike.Enabled = False
                    tlbTool.Buttons("VerifyStrike").Enabled = False
                End If
                If mnuEditPrepare.Visible Then
                    mnuEditPrepare.Enabled = False
                    mnuEditBack.Enabled = False
                    tlbTool.Buttons("Prepare").Enabled = False
                    tlbTool.Buttons("Back").Enabled = False
                End If
                If mnuEditDisplay.Visible = True Then
                    mnuEditDisplay.Enabled = True
                End If
                If mlngMode = ģ���.ҩƷ�ƿ� Then
                    If TabShow.Tab = 0 Then
                        mnuEditStrike.Enabled = False
                        tlbTool.Buttons("Strike").Enabled = False
                        mnuEditPreparePhysic.Enabled = False
                        mnuEditSendPhysic.Enabled = False
                        mnuEditBack.Enabled = False
                        tlbTool.Buttons("PreparePhysic").Enabled = False
                        tlbTool.Buttons("SendPhysic").Enabled = mnuEditSendPhysic.Enabled
                        tlbTool.Buttons("Back").Enabled = False
                        '����õ�������ˣ�������ҩ�뷢��
                        If TestVerify(vsfList.TextMatrix(vsfList.Row, 0)) Then
                            mnuEditPreparePhysic.Enabled = False
                            mnuEditSendPhysic.Enabled = False
                            mnuEditBack.Enabled = False
                            tlbTool.Buttons("PreparePhysic").Enabled = False
                            tlbTool.Buttons("SendPhysic").Enabled = False
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
    
    Select Case mlngMode
        Case ģ���.�⹺���       'ҩƷ�⹺������
            Set objRow = New zlTabAppRow
            objRow.Add "�ⷿ��" & Trim(cboStock.Text)
            objRow.Add "��Ӧ�̣�" & Trim(vsfList.TextMatrix(vsfList.Row, �⹺����.��Ӧ��))
            objPrint.UnderAppRows.Add objRow
                
        Case ģ���.�������       'ҩƷ����������
            Set objRow = New zlTabAppRow
            objRow.Add "�ⷿ��" & Trim(cboStock.Text)
            objRow.Add "�Ƽ��ң�" & Trim(vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "�Ƽ���")))
            objPrint.UnderAppRows.Add objRow
            
        Case ģ���.�������       'ҩƷ����������
            Set objRow = New zlTabAppRow
            objRow.Add "�ⷿ��" & Trim(cboStock.Text)
            objRow.Add "������" & Trim(vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "������")))
            objPrint.UnderAppRows.Add objRow
        Case ģ���.��۵���       '����۵�������
            Set objRow = New zlTabAppRow
            objRow.Add "�ⷿ��" & Trim(cboStock.Text)
            objPrint.UnderAppRows.Add objRow
            
        Case ģ���.ҩƷ�ƿ�       'ҩƷ�ƿ����
            Set objRow = New zlTabAppRow
            If TabShow.Tab = 0 Then
                objRow.Add "�Ƴ��ⷿ��" & Trim(cboStock.Text)
                objRow.Add "����ⷿ��" & Trim(vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "����ⷿ")))
            Else
                objRow.Add "����ⷿ��" & Trim(cboStock.Text)
                objRow.Add "�Ƴ��ⷿ��" & Trim(vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "�Ƴ��ⷿ")))
            End If
            objPrint.UnderAppRows.Add objRow
            
        Case ģ���.ҩƷ����       'ҩƷ���ù���
            Set objRow = New zlTabAppRow
            objRow.Add "��ҩ�ⷿ��" & Trim(cboStock.Text)
            objRow.Add "���ò��ţ�" & Trim(vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "���ò���")))
            objPrint.UnderAppRows.Add objRow
            
        Case ģ���.��������       'ҩƷ�����������
            Set objRow = New zlTabAppRow
            objRow.Add "�ⷿ��" & Trim(cboStock.Text)
            objRow.Add "������" & Trim(vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "������")))
            objPrint.UnderAppRows.Add objRow
    End Select
        
    Set objRow = New zlTabAppRow
    objRow.Add "ժҪ:" & vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "ժҪ"))
    objPrint.BelowAppRows.Add objRow
        
    Set objRow = New zlTabAppRow
    objRow.Add "������:" & vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "������")) & "  ��������:" & vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "��������"))
    
    If mlngMode = ģ���.ҩƷ�ƿ� Then
        objRow.Add "�����:" & vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "������")) & "  �������:" & vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "��������"))
    Else
        objRow.Add "�����:" & vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "�����")) & "  �������:" & vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "�������"))
    End If
    
    objPrint.BelowAppRows.Add objRow
    
    Set objPrint.Body = vsfDetail
    zlPrintOrView1Grd objPrint, bytMode
End Sub




Private Sub tlbTool_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu mnuViewTool
    End If
End Sub
'Ѱ����ĳһ����ȵ���
Private Function FindRow(ByVal FlexTemp As MSHFlexGrid, ByVal intTemp As Variant, ByVal intCol As Integer) As Integer
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

Private Function TestPrepare(ByVal strNo As String) As Boolean
    Dim intBill As Integer
    Dim rsTemp As New ADODB.Recordset
    '�����ҩ���Ƿ��Ѿ���д
    
    On Error GoTo errHandle
    Select Case mlngMode
    Case ģ���.�⹺���
        intBill = 1
    Case ģ���.ҩƷ�ƿ�
        intBill = 6
    Case Else
        Exit Function
    End Select
    
    gstrSQL = "Select ��ҩ�� From ҩƷ�շ���¼ Where ����=[1] And NO=[2] And Rownum<2"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "����Ƿ�ͨ���˲�", intBill, strNo)

    If Not IsNull(rsTemp!��ҩ��) Then
        TestPrepare = True
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function TestDelete(ByVal strNo As String) As Boolean
    Dim intBill As Integer
    Dim rsTemp As New ADODB.Recordset
    '��鵥���Ƿ�ɾ��
    On Error GoTo errHandle

    Select Case mlngMode
    Case ģ���.�⹺���
        intBill = 1
    Case ģ���.ҩƷ�ƿ�
        intBill = 6
    Case Else
        Exit Function
    End Select
    
    gstrSQL = "Select id From ҩƷ�շ���¼ Where ����=[1] And NO=[2] And Rownum<2"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "����Ƿ�ͨ���˲�", intBill, strNo)
    
    TestDelete = (rsTemp.RecordCount = 0)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function TestVerify(ByVal strNo As String) As Boolean
    Dim int���� As Integer
    Dim rsTemp As New ADODB.Recordset
    '���õ����Ƿ�ͨ�����
    On Error GoTo errHandle

    Select Case mlngMode
        Case ģ���.�⹺���
            int���� = 1
        Case ģ���.ҩƷ�ƿ�
            int���� = 6
    End Select
    
    gstrSQL = "Select ����� From ҩƷ�շ���¼ " & _
        " Where ����=[1] And NO=[2] And Rownum<2"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "����Ƿ�ͨ�����", int����, strNo)
    
    If Not IsNull(rsTemp!�����) Then
        TestVerify = True
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ShowColor(ByVal rsDetail As ADODB.Recordset)
    Dim lngCol As Long, lngCols As Long
    Dim lngRow As Long, lngRows As Long
    Dim bln�б�ҩƷ As Boolean
    Dim dbl��������� As Double
    
    'Ϊ�⹺��ⵥ��ɫ
    If mlngMode <> ģ���.�⹺��� Then Exit Sub
    If rsDetail.State = 0 Then Exit Sub
    If rsDetail.RecordCount = 0 Then Exit Sub
    vsfDetail.Redraw = flexRDNone
    lngRows = vsfDetail.rows - 1
    lngCols = vsfDetail.Cols - 1
    rsDetail.MoveFirst
    
    For lngRow = 1 To lngRows
        '�б�ҩƷ��Ҫ��ɫ
        vsfDetail.Row = lngRow
        bln�б�ҩƷ = (nvl(rsDetail!�б�ҩƷ, 0) = 1)
        dbl��������� = nvl(rsDetail!���������, 0)
        
        If bln�б�ҩƷ Then
            vsfDetail.Cell(flexcpForeColor, lngRow, 0, lngRow, lngCols) = IIf(dbl��������� = 0, &H800000, &H800080)
        Else
            vsfDetail.Cell(flexcpForeColor, lngRow, 0, lngRow, lngCols) = IIf(dbl��������� = 0, &H0, &H40&)
        End If

        rsDetail.MoveNext
    Next
    
    vsfDetail.Row = 1
    vsfDetail.Col = 0: vsfDetail.ColSel = lngCols
    vsfDetail.Redraw = flexRDDirect
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

Private Sub CheckNumber()
    '�����д������ʵ��������һ�£����ú�ɫ�����עʵ��������������
    Dim intRow As Integer
    Dim blnColor As Boolean

    With vsfDetail
        If .TextMatrix(1, 1) = "" Then Exit Sub
        For intRow = 1 To .rows - 1
            blnColor = False
            If .TextMatrix(intRow, .ColIndex("ҩƷID")) = "" Then Exit Sub
            If Val(.TextMatrix(intRow, .ColIndex("��д����"))) <> Val(.TextMatrix(intRow, .ColIndex("ʵ������"))) Then blnColor = True
            .Cell(flexcpForeColor, intRow, .ColIndex("ʵ������"), intRow, .ColIndex("ʵ������")) = IIf(blnColor, vbRed, vbBlack)
        Next
    End With
                
End Sub
