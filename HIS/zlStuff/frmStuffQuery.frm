VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmStuffQuery 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   Caption         =   "���Ŀ���ѯ"
   ClientHeight    =   7110
   ClientLeft      =   165
   ClientTop       =   825
   ClientWidth     =   9495
   Icon            =   "frmStuffQuery.frx":0000
   LockControls    =   -1  'True
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   7110
   ScaleWidth      =   9495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picColor 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   5160
      ScaleHeight     =   255
      ScaleWidth      =   2055
      TabIndex        =   12
      Top             =   6120
      Width           =   2055
      Begin VB.PictureBox picColor1 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   14
         Top             =   0
         Width           =   260
      End
      Begin VB.PictureBox picColor2 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   960
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   13
         Top             =   0
         Width           =   260
      End
      Begin VB.Label lblColor2 
         AutoSize        =   -1  'True
         Caption         =   "��Ч��"
         Height          =   180
         Left            =   1320
         TabIndex        =   16
         Top             =   30
         Width           =   540
      End
      Begin VB.Label lblColor1 
         AutoSize        =   -1  'True
         Caption         =   "ͣ��"
         Height          =   180
         Left            =   360
         TabIndex        =   15
         Top             =   30
         Width           =   360
      End
   End
   Begin VB.PictureBox picFind 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1470
      ScaleHeight     =   255
      ScaleWidth      =   2655
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   6810
      Width           =   2655
      Begin VB.TextBox txt������Ϣ 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   780
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   0
         Width           =   1245
      End
      Begin VB.Label lbl������Ϣ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "������Ϣ"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   30
         TabIndex        =   10
         Top             =   37
         Width           =   720
      End
   End
   Begin MSComctlLib.ImageList imglvw 
      Left            =   2985
      Top             =   2205
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
            Picture         =   "frmStuffQuery.frx":0982
            Key             =   "root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQuery.frx":268C
            Key             =   "child"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picVLine_S 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5460
      Left            =   2940
      MousePointer    =   9  'Size W E
      ScaleHeight     =   5460
      ScaleWidth      =   45
      TabIndex        =   1
      Top             =   1305
      Width           =   45
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   1125
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   1984
      BandCount       =   2
      _CBWidth        =   9495
      _CBHeight       =   1125
      _Version        =   "6.7.9782"
      Child1          =   "tbrThis"
      MinHeight1      =   720
      Width1          =   2730
      NewRow1         =   0   'False
      Caption2        =   "�ⷿ"
      Child2          =   "cob�ⷿ"
      MinHeight2      =   300
      Width2          =   6780
      NewRow2         =   -1  'True
      Begin VB.ComboBox cob�ⷿ 
         Height          =   300
         Left            =   585
         TabIndex        =   5
         Text            =   "cob�ⷿ"
         Top             =   780
         Width           =   8820
      End
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   720
         Left            =   165
         TabIndex        =   3
         Top             =   30
         Width           =   9240
         _ExtentX        =   16298
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imgTbrStard"
         HotImageList    =   "imgTbrHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   15
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ԥ��"
               Key             =   "Ԥ��"
               Object.ToolTipText     =   "Ԥ��"
               Object.Tag             =   "Ԥ��"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "��ӡ"
               Description     =   "��ӡ"
               Object.ToolTipText     =   "��ӡ"
               Object.Tag             =   "��ӡ"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "��������"
               Object.Tag             =   "����"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               ImageIndex      =   3
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ϸ"
               Key             =   "��ϸ"
               Object.ToolTipText     =   "������ϸ��"
               Object.Tag             =   "��ϸ"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "��������"
               Object.Tag             =   "����"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "��������"
               Object.Tag             =   "����"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ˢ��"
               Key             =   "ˢ��"
               Object.ToolTipText     =   "ˢ��"
               Object.Tag             =   "ˢ��"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Description     =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "�˳�"
               Description     =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imgTbrHot 
      Left            =   1425
      Top             =   780
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
            Picture         =   "frmStuffQuery.frx":4396
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQuery.frx":45B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQuery.frx":47CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQuery.frx":49E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQuery.frx":4C04
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQuery.frx":4E1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQuery.frx":5038
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQuery.frx":5254
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQuery.frx":5470
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQuery.frx":568C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQuery.frx":58A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQuery.frx":6180
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgTbrStard 
      Left            =   690
      Top             =   810
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
            Picture         =   "frmStuffQuery.frx":649A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQuery.frx":66B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQuery.frx":68D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQuery.frx":6AEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQuery.frx":6D08
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQuery.frx":6F22
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQuery.frx":713C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQuery.frx":7358
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQuery.frx":7574
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQuery.frx":7790
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQuery.frx":79AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQuery.frx":8284
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Msf������Ϣ_S 
      Height          =   2985
      Left            =   3000
      TabIndex        =   6
      Top             =   1260
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   5265
      _Version        =   393216
      BackColor       =   16777215
      FixedCols       =   0
      GridColor       =   -2147483631
      GridColorFixed  =   8421504
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      FillStyle       =   1
      GridLinesFixed  =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Msf�������_S 
      Height          =   870
      Left            =   3240
      TabIndex        =   7
      Top             =   5100
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   1535
      _Version        =   393216
      BackColor       =   16777215
      FixedCols       =   0
      GridColor       =   -2147483631
      GridColorFixed  =   8421504
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      FillStyle       =   1
      GridLinesFixed  =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   11
      Top             =   6744
      Width           =   9492
      _ExtentX        =   16748
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmStuffQuery.frx":83E6
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11668
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
   Begin MSComctlLib.TreeView tvwSection_S 
      Height          =   4350
      Left            =   60
      TabIndex        =   0
      Top             =   1275
      Width           =   2790
      _ExtentX        =   4921
      _ExtentY        =   7673
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "imglvw"
      Appearance      =   1
   End
   Begin VB.Label lbl����_S 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "�������"
      ForeColor       =   &H8000000E&
      Height          =   180
      Left            =   3270
      MousePointer    =   7  'Size N S
      TabIndex        =   4
      Top             =   4920
      Width           =   6585
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFilePrintSet 
         Caption         =   "��ӡ����(&S)"
      End
      Begin VB.Menu mnuFilePrintView 
         Caption         =   "��ӡԤ��(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "��ӡ(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuExcel 
         Caption         =   "�����&Excel"
      End
      Begin VB.Menu mnuFileBatch 
         Caption         =   "������ӡ��ϸ��(&B)"
      End
      Begin VB.Menu mnuViewLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "��������(&R)"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "�˳�(&X)"
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
      Begin VB.Menu mnuViewStatus 
         Caption         =   "״̬��(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFont 
         Caption         =   "����(&F)"
         Begin VB.Menu mnuViewFontSize 
            Caption         =   "С����"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuViewFontSize 
            Caption         =   "������"
            Index           =   1
         End
         Begin VB.Menu mnuViewFontSize 
            Caption         =   "������"
            Index           =   2
         End
      End
      Begin VB.Menu mnuViewForeColor 
         Caption         =   "ǰ��ɫ(&C)"
      End
      Begin VB.Menu mnuViewBackColor 
         Caption         =   "����ɫ(&B)"
      End
      Begin VB.Menu mnuviewLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "ˢ��(&R)"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "����(&F)"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuHelpTitle 
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
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)..."
      End
   End
   Begin VB.Menu mnuPopu 
      Caption         =   "����"
      Visible         =   0   'False
      Begin VB.Menu mnuPopuOpen 
         Caption         =   "��(&O)"
      End
      Begin VB.Menu mnuPopuBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopuFontSize 
         Caption         =   "С����"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuPopuFontSize 
         Caption         =   "������"
         Index           =   1
      End
      Begin VB.Menu mnuPopuFontSize 
         Caption         =   "������"
         Index           =   2
      End
   End
   Begin VB.Menu mnuReportBill 
      Caption         =   "����˵�"
      Visible         =   0   'False
      Begin VB.Menu mnuBill 
         Caption         =   "����(&D)"
      End
   End
End
Attribute VB_Name = "frmStuffQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'--------------------------------------------------------
Public mblnDo As Boolean
Public mbytUint As Byte '���ĵ�λ

Dim mintFont As Integer
Dim WithEvents mrsData  As ADODB.Recordset
Attribute mrsData.VB_VarHelpID = -1
Dim mrsTreeData As ADODB.Recordset

Dim mstrStartDate As String
Dim mstrEndDate As String

Dim mblnFirst As Boolean              'ȷ���Ƿ��һ��ʹ�ñ�ϵͳ
Dim mbln����� As Boolean
Dim mbln����ͣ�� As Boolean
Dim mintMonths As Integer
Dim mstrPrivs As String
Dim mblnColor As Boolean

Private mlngCardRow As Long
Private mlngRow As Long
Private mstrCardSort As String                 '������

Private mblnNoClick As Boolean

Private Const MLNG��ɫ As Long = &H80000005
Private Const MLNG��ɫ As Long = &H80000008
Private Const MLNG��ɫ As Long = &H8000000D
Private Const MLNGSEL As Long = &HA87B82
Private Const MLNG��ɫ As Long = &H8000000F
Private Const MLNG��ɫ As Long = &HC0C0C0
Private Const MLNG��ɫ As Long = &HC0           'ͣ��
Private mblnCostView As Boolean             '�鿴�ɱ��������Ϣ true-����鿴 false-������鿴
Private Const mstrCaption As String = "���Ŀ���ѯ"


Private mstrOthers() As String  '  0-����,1-����,2-����,3-���,4-����,5-ָ������
'----------------------
'���ű���ı�������
Public WithEvents mobjReport As zl9Report.clsReport
Attribute mobjReport.VB_VarHelpID = -1
Private mlngCurReport As Long
Private mobjCurSheet As Object
Dim mstrNoS As String
'-----------------------
Private mlngModule As Long
'----------------------------------------------------------------------------------------------------------
'���˺�:����С��λ���ĸ�ʽ��
'�޸�:2007/03/06
Private mFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------
Private Sub cbrThis_Resize()
    Form_Resize
End Sub

Private Sub cob�ⷿ_Click()
    If mblnNoClick Then Exit Sub
    If Me.tvwSection_S.Nodes.Count = 0 Then Exit Sub
    Me.tvwSection_S.Tag = ""
    ReFreshStuffData Me.cob�ⷿ.ItemData(Me.cob�ⷿ.ListIndex), mstrStartDate, mstrEndDate, IIf(Left(Me.tvwSection_S.SelectedItem.Key, 1) = "R", 0, Mid(Me.tvwSection_S.SelectedItem.Key, 2))
End Sub

Private Sub cob�ⷿ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If cob�ⷿ.ListCount = 0 Then Call zlControl.ControlSetFocus(Msf������Ϣ_S): Exit Sub
    
    If cob�ⷿ.ListIndex >= 0 Then
        If Val(cob�ⷿ.Tag) = cob�ⷿ.ItemData(cob�ⷿ.ListIndex) Then
            Call zlControl.ControlSetFocus(Msf������Ϣ_S, True)
            Exit Sub
        End If
    End If
    
    If Select����ѡ����(Me, cob�ⷿ, Trim(cob�ⷿ.Text), "V,K,12,W", IIf(InStr(1, mstrPrivs, "���пⷿ") = 0, True, False)) = False Then
        Exit Sub
    End If
    If cob�ⷿ.ListIndex >= 0 Then
        cob�ⷿ.Tag = cob�ⷿ.ItemData(cob�ⷿ.ListIndex)
    End If
End Sub


Private Sub cob�ⷿ_LostFocus()
    Dim i As Long
    
    If cob�ⷿ.ListCount = 0 Then Exit Sub
    If cob�ⷿ.ListIndex < 0 Then
        For i = 0 To cob�ⷿ.ListCount - 1
            If Val(cob�ⷿ.Tag) = cob�ⷿ.ItemData(i) Then
                mblnNoClick = True
                cob�ⷿ.ListIndex = i: Exit For
            End If
        Next
    End If
    mblnNoClick = False
End Sub


Private Sub mrsData_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If mblnColor Then Exit Sub
    If mrsData.RecordCount = 0 Then
        RefreshBatch Me.cob�ⷿ.ItemData(Me.cob�ⷿ.ListIndex), 0
        Exit Sub
    End If
    If mrsData.EOF Then mrsData.MoveFirst
    If mrsData.EOF Then Exit Sub
    RefreshBatch Me.cob�ⷿ.ItemData(Me.cob�ⷿ.ListIndex), mrsData.Fields("����id").Value
    If Me.tvwSection_S.Tag <> "T" Then Exit Sub
    err = 0
    On Error Resume Next
    Me.tvwSection_S.Nodes("_" & mrsData.Fields("����id").Value).Selected = True
    Me.tvwSection_S.Nodes("_" & mrsData.Fields("����id").Value).Expanded = True
End Sub

Private Sub Form_Activate()
    If Not mblnFirst Then Exit Sub
    
    mblnFirst = False
    
    tbrThis.Buttons("����").Visible = gblnCode
    
    If Not ReFreshTreeView() Then Unload Me: Exit Sub
    ReFreshStuffData Me.cob�ⷿ.ItemData(Me.cob�ⷿ.ListIndex), mstrStartDate, mstrEndDate, IIf(Left(Me.tvwSection_S.SelectedItem.Key, 1) = "R", 0, Mid(Me.tvwSection_S.SelectedItem.Key, 2))
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim strOthers(0 To 6) As String
    mlngModule = glngModul
    mstrPrivs = gstrPrivs
    mblnCostView = zlStr.IsHavePrivs(mstrPrivs, "�鿴�ɱ���")
    
    For i = 0 To 6
        strOthers(i) = ""
    Next
    mstrOthers = strOthers
    mblnFirst = True
    Call GetParaSet
    mnuViewForeColor.Visible = False
    mnuViewBackColor.Visible = False
    
    Call mnuViewFontSize_Click(mintFont)
    
    With Msf�������_S
        .Clear
        .Cols = IIf(gblnCode = True, 15, 13)
        .Rows = 2
        .FixedRows = 1
        .FixedCols = 0
        
        .ColAlignmentFixed(0) = 4
        .ColAlignmentFixed(1) = 4
        .ColAlignmentFixed(2) = 4
        .ColAlignmentFixed(3) = 4
        .ColAlignmentFixed(4) = 4
        .ColAlignmentFixed(5) = 4
        .ColAlignmentFixed(6) = 4
        .ColAlignmentFixed(7) = 4
        .ColAlignmentFixed(8) = 4
        .ColAlignmentFixed(9) = 4
        .ColAlignmentFixed(10) = 4
        
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .ColAlignment(3) = 1
        .ColAlignment(4) = 7
        .ColAlignment(5) = 7
        .ColAlignment(6) = 7
        .ColAlignment(7) = 7
        .ColAlignment(8) = 7
        .ColAlignment(9) = 7
        .ColAlignment(10) = 7

        .TextMatrix(0, 0) = "�ⷿ"
        .TextMatrix(0, 1) = "����"
        .TextMatrix(0, 2) = "ʧЧ��"
        .TextMatrix(0, 3) = "����"
        .TextMatrix(0, 4) = "���ÿ��"
        .TextMatrix(0, 5) = "�������"
        .TextMatrix(0, 6) = "�����"
        .TextMatrix(0, 7) = "�ɱ���"
        .TextMatrix(0, 8) = "�ɱ����"
        .TextMatrix(0, 9) = "�����"
        .TextMatrix(0, 10) = "������"
        
        .ColWidth(0) = 1000
        .ColWidth(1) = 1000
        .ColWidth(2) = 1200
        .ColWidth(3) = 1200
        .ColWidth(4) = 1500
        .ColWidth(5) = 1500
        .ColWidth(6) = 1500
        .ColWidth(10) = 1500
        
        If gblnCode = True Then
            .TextMatrix(0, 11) = "��Ʒ����"
            .TextMatrix(0, 12) = "�ڲ�����"
            .TextMatrix(0, 13) = "�ۼ�"
            .TextMatrix(0, 14) = "��Ӧ��"
            
            .ColAlignment(11) = 7
            .ColAlignment(12) = 7
            .ColAlignment(13) = 7
            .ColAlignment(14) = 1
            
            .ColAlignmentFixed(11) = 4
            .ColAlignmentFixed(12) = 4
            .ColAlignmentFixed(13) = 4
            .ColAlignmentFixed(14) = 4
            
            .ColWidth(11) = 2000
            .ColWidth(12) = 2000
            .ColWidth(13) = 1500
            .ColWidth(14) = 1500
        Else
            .TextMatrix(0, 11) = "�ۼ�"
            .TextMatrix(0, 12) = "��Ӧ��"
            .ColAlignment(11) = 7
            .ColAlignment(12) = 1
            .ColAlignmentFixed(11) = 4
            .ColAlignmentFixed(12) = 4
            .ColWidth(11) = 1500
            .ColWidth(12) = 1500
        End If
    End With
    Call SetFormat(True)
    RestoreWinState Me, App.ProductName, mstrCaption
    Msf�������_S.ColWidth(7) = IIf(mblnCostView = False, 0, 1500)
    Msf�������_S.ColWidth(8) = IIf(mblnCostView = False, 0, 1500)
    Msf�������_S.ColWidth(9) = IIf(mblnCostView = False, 0, 1500)
    
    '���ر���
    Set mobjReport = New zl9Report.clsReport

    '2006-04-25:���˺�,ͳһ���ӱ�������ģ��Ĺ���
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, gstrPrivs)
    
    Call ����Ȩ��
    
    Call SetParent(picFind.hwnd, stbThis.hwnd)
    picFind.Top = 80
    picFind.Left = stbThis.Panels(1).Width + 80
    
    stbThis.Panels(2).Picture = picColor
End Sub

Private Sub Form_Resize()
    Dim intTop As Integer, intButton As Integer
    If Me.WindowState = 1 Then Exit Sub
    intTop = IIf(Me.cbrThis.Visible, Me.cbrThis.Height, 0)
    intButton = IIf(Me.stbThis.Visible, Me.stbThis.Height, 0)
    
    On Error Resume Next
    Me.picVLine_S.Top = intTop + Me.ScaleTop
    Me.picVLine_S.Height = Me.ScaleHeight - Me.tvwSection_S.Top - intButton
    If Me.picVLine_S.Left < 500 Then Me.picVLine_S.Left = 500
    If Me.picVLine_S.Left > Me.ScaleWidth - 500 Then Me.picVLine_S.Left = Me.ScaleWidth - 500
    
    Me.tvwSection_S.Left = Me.ScaleLeft
    Me.tvwSection_S.Width = Me.picVLine_S.Left - Me.tvwSection_S.Left
    Me.tvwSection_S.Top = Me.ScaleTop + intTop
    
    If Me.ScaleWidth - Me.picVLine_S.Left - Me.picVLine_S.Width < 500 Then
        Me.Width = Me.picVLine_S.Left + Me.picVLine_S.Width + 500
    End If
    If Me.ScaleHeight - Me.lbl����_S.Top - Me.lbl����_S.Height < 500 Then
        Me.Height = Me.lbl����_S.Top + Me.lbl����_S.Height + 2000
    End If
    If Me.ScaleHeight < 500 Then
        Me.Height = 2000
    End If
    Me.tvwSection_S.Height = Me.ScaleHeight - tvwSection_S.Top - intButton
    
    Me.lbl����_S.Left = Me.picVLine_S.Left + Me.picVLine_S.Width
    Me.lbl����_S.Width = Me.ScaleWidth - Me.lbl����_S.Left
    With Me.Msf�������_S
        .Left = Me.lbl����_S.Left
        .Width = Me.lbl����_S.Width
    End With
    
    Me.Msf������Ϣ_S.Left = Me.lbl����_S.Left
    Me.Msf������Ϣ_S.Width = Me.lbl����_S.Width
        
    If Me.Msf�������_S.Visible Then
        With Me.Msf�������_S
            .Top = Me.lbl����_S.Top + Me.lbl����_S.Height
            .Height = Me.ScaleHeight - .Top - intButton
        End With
        Me.Msf������Ϣ_S.Top = intTop + 50
        Me.Msf������Ϣ_S.Height = Me.lbl����_S.Top - Me.Msf������Ϣ_S.Top
    Else
        Me.Msf������Ϣ_S.Top = intTop + 50
        Me.Msf������Ϣ_S.Height = Me.ScaleHeight - Me.Msf������Ϣ_S.Top - intButton
    End If
    
    With picColor
        .Top = Me.ScaleHeight - .Height - 30
        .Left = Me.ScaleWidth - stbThis.Panels(3).Width - stbThis.Panels(4).Width - .Width - 300
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    SaveWinState Me, App.ProductName, mstrCaption
End Sub

Private Sub lbl����_S_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        Me.lbl����_S.Top = Me.lbl����_S.Top + y
        If Me.lbl����_S.Top < 5000 Then Me.lbl����_S.Top = 5000
        If Me.Height - Me.lbl����_S.Top < 2000 Then Me.lbl����_S.Top = Me.Height - 2000
        Form_Resize
    End If
End Sub

Private Sub mnuEXCEL_Click()
    grdPrint 1
End Sub

Private Sub mnuFileBatch_Click()
    With FrmPrintList
        .mstrPrivs = mstrPrivs
        .Show 1, Me
    End With
End Sub

Private Sub mnufileexit_Click()
    Unload Me
End Sub
Private Sub GetParaSet()
    '����:��ȡ��������
    mbytUint = Val(zlDatabase.GetPara("���ĵ�λ", glngSys, mlngModule))
 
    '���˺�:����С����ʽ����
    With mFMT
        .FM_�ɱ��� = GetFmtString(mbytUint, g_�ɱ���)
        .FM_��� = GetFmtString(mbytUint, g_���)
        .FM_���ۼ� = GetFmtString(mbytUint, g_�ۼ�)
        .FM_���� = GetFmtString(mbytUint, g_����)
    End With
    
    
    mbln����� = IIf(Val(zlDatabase.GetPara("ֻ��ʾ�п������", glngSys, mlngModule)) = 1, 1, 0) = 1
    mintMonths = Val(zlDatabase.GetPara("��������", glngSys, mlngModule, 3))  '
    mbln����ͣ�� = IIf(Val(zlDatabase.GetPara("����ͣ������", glngSys, mlngModule)) = 1, 1, 0) = 1
    mintFont = Val(zlDatabase.GetPara("�����ֺ�", glngSys, mlngModule, 9))
   
End Sub
Private Sub mnuFileOpen_Click()
    mblnDo = False
    Call frmStuffQueryParaSet.��������(Me, mlngModule, mstrPrivs)
    If Not mblnDo Then Exit Sub
    If Me.tvwSection_S.Nodes.Count = 0 Then Exit Sub
    
    Call GetParaSet
    Me.tvwSection_S.Tag = ""
    ReFreshStuffData Me.cob�ⷿ.ItemData(Me.cob�ⷿ.ListIndex), mstrStartDate, mstrEndDate, IIf(Left(Me.tvwSection_S.SelectedItem.Key, 1) = "R", 0, Mid(Me.tvwSection_S.SelectedItem.Key, 2))
End Sub

Private Sub mnuFilePrint_Click()
    grdPrint 3
End Sub

Private Sub mnuFilePrintSet_Click()
     zlPrintSet
End Sub

Private Sub mnuFilePrintView_Click()
  grdPrint 0
End Sub
Private Sub grdPrint(blnIsPreview As Byte)
    '---------------------------------------------------
    '���ܣ�    ������Ļ��֯���ϸ�����Ŀ����ӡԤ��
    '������
    '     blnIsPreview: 0��ʾԤ�� 1��ʾ�����EXCEL ������ʾ��ӡ
    '���أ�
    '---------------------------------------------------
    Dim objPrint As New zlPrint1Grd
    Dim objRow As zlTabAppRow
    
    objPrint.Title.Text = "���Ŀ���ѯ"
    Set objRow = New zlTabAppRow
    objRow.Add "�ⷿ��" & Me.cob�ⷿ.Text
    objRow.Add "������;��" & Me.tvwSection_S.SelectedItem.Text
    objRow.Add "��ֹ���ڣ�" & Format(Sys.Currentdate, "yyyy��MM��DD��")
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ��:" & UserInfo.�û���
    objRow.Add "��ӡʱ��:" & Format(Sys.Currentdate, "yyyy��MM��DD�� HH:MM")
    objPrint.BelowAppRows.Add objRow
    Set objPrint.Body = Msf������Ϣ_S
    
    Call Msf������Ϣ_S_LostFocus
    If blnIsPreview = 0 Then
         zlPrintOrView1Grd objPrint, 2
    Else
      If blnIsPreview = 1 Then
            zlPrintOrView1Grd objPrint, 3
      Else
        Select Case zlPrintAsk(objPrint)
            Case 1
                 zlPrintOrView1Grd objPrint, 1
            Case 2
                zlPrintOrView1Grd objPrint, 2
            Case 3
                zlPrintOrView1Grd objPrint, 3
        End Select
      End If
    End If
    Set objPrint = Nothing
End Sub

Private Sub mnuHelpAbout_Click()
   ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpTitle_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
End Sub

Private Sub mnuViewFind_Click()
    Dim rsTmp As New ADODB.Recordset
    Dim strFind As String
    Me.tvwSection_S.Tag = ""
    FrmStuffQueryFind.Show 1, Me
    strFind = FrmStuffQueryFind.mstrTemp
    mstrOthers = FrmStuffQueryFind.mstrOthers
    
    Unload FrmStuffQueryFind
    If strFind = "" Then Exit Sub
    If Not ReFreshFilterData(cob�ⷿ.ItemData(cob�ⷿ.ListIndex), strFind) Then Exit Sub
    Me.tvwSection_S.Tag = "T"
End Sub

Private Sub mnuViewFontSize_Click(Index As Integer)
    Dim i As Integer
    For i = 0 To 2
        Me.mnuViewFontSize(i).Checked = False
    Next
    Me.mnuViewFontSize(Index).Checked = True

    Select Case Index
    Case 0
        Me.Msf������Ϣ_S.Font.Size = 9
        Me.tvwSection_S.Font.Size = 9
        Msf�������_S.Font.Size = 9
     Case 1
        Me.Msf������Ϣ_S.Font.Size = 11
        Me.tvwSection_S.Font.Size = 11
        Msf�������_S.Font.Size = 11
    Case 2
        Me.Msf������Ϣ_S.Font.Size = 15
        Me.tvwSection_S.Font.Size = 15
        Msf�������_S.Font.Size = 15
    End Select
    mintFont = Index
    Call zlDatabase.SetPara("�����ֺ�", mintFont, glngSys, mlngModule)
    Form_Resize
    Me.Refresh
End Sub

Private Sub mnuViewForeColor_Click()
    Dim lngForeColor As Long
    lngForeColor = zlGetColor(Me.Msf������Ϣ_S.ForeColor)
    Me.Msf������Ϣ_S.Redraw = False
    Me.Msf������Ϣ_S.ForeColor = lngForeColor
    Me.Msf������Ϣ_S.Redraw = True
    
End Sub
Private Sub mnuViewBackColor_Click()
    Dim lngBackColor As Long
    lngBackColor = zlGetColor(Me.Msf������Ϣ_S.BackColor)
    Me.Msf������Ϣ_S.BackColor = lngBackColor
    
End Sub

Private Sub showReportMXZ()
    If mrsData Is Nothing Then Exit Sub
    If Not (mrsData.State = 1) Then Exit Sub
    If mrsData.RecordCount = 0 Then Exit Sub
    If ISCheckReport("ZL1_INSIDE_1721_2") = False Then Exit Sub
    
    If cob�ⷿ.ItemData(cob�ⷿ.ListIndex) = 0 Then
        Call mobjReport.ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1721_2", Me, "����=" & mrsData.Fields("����").Value & "|" & mrsData.Fields("����id").Value, "�ⷿ=���пⷿ|is not null", "��λ=" & Choose(mbytUint, "ɢװ��λ", "��װ��λ") & "|" & mbytUint, "��ʼ����=" & Format(DateAdd("m", -1, Sys.Currentdate), "yyyy-MM-DD"), "��������=" & Format(Sys.Currentdate, "yyyy-MM-DD"))
    Else
        Call mobjReport.ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1721_2", Me, "����=" & mrsData.Fields("����").Value & "|" & mrsData.Fields("����id").Value, "�ⷿ=" & cob�ⷿ.Text & "|=  " & cob�ⷿ.ItemData(cob�ⷿ.ListIndex), "��λ=" & Choose(mbytUint, "ɢװ��λ", "��װ��λ") & "|" & mbytUint, "��ʼ����=" & Format(DateAdd("m", -1, Sys.Currentdate), "yyyy-MM-DD"), "��������=" & Format(Sys.Currentdate, "yyyy-MM-DD"), "��λ=" & mbytUint)      '"����δ��˵���=0| And A.����� Is Not NULL"
    End If
    
End Sub


Private Sub showReportCode()
    If mrsData Is Nothing Then Exit Sub
    If Not (mrsData.State = 1) Then Exit Sub
    If mrsData.RecordCount = 0 Then Exit Sub
    
    If cob�ⷿ.ItemData(cob�ⷿ.ListIndex) = 0 Then Exit Sub
    If Msf�������_S.Row = 0 Then Exit Sub
    If Msf�������_S.TextMatrix(Msf�������_S.Row, 11) = "" And Msf�������_S.TextMatrix(Msf�������_S.Row, 12) = "" Then Exit Sub
    
    Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1721_4", Me, "�ⷿ=" & cob�ⷿ.Text & "|=  " & cob�ⷿ.ItemData(cob�ⷿ.ListIndex), "��Ʒ����=" & Msf�������_S.TextMatrix(Msf�������_S.Row, 11), "�ڲ�����=" & Msf�������_S.TextMatrix(Msf�������_S.Row, 12))
        
End Sub

Private Sub mnuViewRefresh_Click()
    ReFreshStuffData Me.cob�ⷿ.ItemData(Me.cob�ⷿ.ListIndex), mstrStartDate, mstrEndDate, IIf(Left(Me.tvwSection_S.SelectedItem.Key, 1) = "R", 0, Mid(Me.tvwSection_S.SelectedItem.Key, 2))
End Sub

Private Sub ShowReportMXB()
    On Error Resume Next
    Call mobjReport.ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1721_3", Me, "�ⷿ=" & Me.cob�ⷿ.Text & "|" & IIf(Me.cob�ⷿ.ItemData(Me.cob�ⷿ.ListIndex) = 0, " is not null ", "=" & Me.cob�ⷿ.ItemData(Me.cob�ⷿ.ListIndex)), "��λ=" & Choose(mbytUint, "ɢװ��λ", "��װ��λ") & "|" & mbytUint)
End Sub

Private Sub mnuViewStatus_Click()
    Me.mnuViewStatus.Checked = Not Me.mnuViewStatus.Checked
    Me.stbThis.Visible = Me.mnuViewStatus.Checked
    Form_Resize
End Sub

Private Sub ShowReportSumAccount()
    err = 0
    On Error Resume Next
    Call mobjReport.ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1721_1", Me, "�ⷿ=" & Me.cob�ⷿ.Text & "|" & IIf(Me.cob�ⷿ.ItemData(Me.cob�ⷿ.ListIndex) = 0, " is not null ", "=" & Me.cob�ⷿ.ItemData(Me.cob�ⷿ.ListIndex)))
End Sub
Private Sub SetReportCtrlIndexEnabled()
    '����ָ�������Enable����
    Dim i As Long
    For i = 0 To mnuReportItem.UBound
        If Split(mnuReportItem(i).Tag & ",", ",")(1) = "ZL1_INSIDE_1721_2" Then
            tbrThis.Buttons("��ϸ").Enabled = mrsData.RecordCount <> 0
            mnuReportItem.Item(i).Enabled = mrsData.RecordCount <> 0
        End If
    Next
End Sub
Private Sub mnuReportItem_Click(Index As Integer)


    If Split(mnuReportItem(Index).Tag & ",", ",")(1) = "ZL1_INSIDE_1721_2" Then
        '��ϸ��
        Call showReportMXZ
        Exit Sub
    End If
    
    If Split(mnuReportItem(Index).Tag & ",", ",")(1) = "ZL1_INSIDE_1721_3" Then
        '��ϸ��
        Call ShowReportMXB
        Exit Sub
    End If
    
    
    If Split(mnuReportItem(Index).Tag & ",", ",")(1) = "ZL1_INSIDE_1721_1" Then
        '����
        Call ShowReportSumAccount
        Exit Sub
    End If
    
    
    Dim lng�ⷿID As Long, lng����id As Long, lng����ID As Long
    If cob�ⷿ.ListIndex < 0 Then
        lng�ⷿID = 0
    Else
        lng�ⷿID = cob�ⷿ.ItemData(cob�ⷿ.ListIndex)
    End If
    
    If Not tvwSection_S.SelectedItem Is Nothing Then
        lng����id = Val(Mid(tvwSection_S.SelectedItem.Key, 2))
    End If
    
    lng����ID = 0
    If Not mrsData Is Nothing Then
        If mrsData.State = 1 Then
            If Not mrsData.EOF Then
               lng����ID = mrsData.Fields("����id").Value
            End If
        End If
    End If
    
    '2006-04-25:���˺�:�����Զ��屨������ģ��Ĺ���
    Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, "�ⷿ=" & lng�ⷿID, "����=" & lng����id, "����=" & lng����ID)
    
End Sub

Private Sub mnuViewToolbarStAnd_Click()
    Dim intCount As Integer
    Me.mnuViewToolbarStand.Checked = Not Me.mnuViewToolbarStand.Checked
    Me.mnuViewToolbarText.Enabled = Me.mnuViewToolbarStand.Checked
    Me.cbrThis.Visible = Me.mnuViewToolbarStand.Checked
    
    If Me.mnuViewToolbarText.Checked Then
        For intCount = 1 To Me.tbrThis.Buttons.Count
            Me.tbrThis.Buttons(intCount).Caption = Me.tbrThis.Buttons(intCount).Tag
        Next
    Else
        For intCount = 1 To Me.tbrThis.Buttons.Count
            Me.tbrThis.Buttons(intCount).Caption = ""
        Next
    End If
    Me.cbrThis.Bands(1).MinHeight = Me.tbrThis.Height
    Me.cbrThis.Refresh
    Form_Resize

End Sub
Private Sub mnuViewToolbarText_Click()
    Dim intCount As Integer
    Me.mnuViewToolbarText.Checked = Not Me.mnuViewToolbarText.Checked
    If Me.mnuViewToolbarText.Checked Then
        For intCount = 1 To Me.tbrThis.Buttons.Count
            Me.tbrThis.Buttons(intCount).Caption = Me.tbrThis.Buttons(intCount).Tag
        Next
    Else
        For intCount = 1 To Me.tbrThis.Buttons.Count
            Me.tbrThis.Buttons(intCount).Caption = ""
        Next
    End If
    Me.cbrThis.Bands(1).MinHeight = Me.tbrThis.Height
    Me.cbrThis.Refresh
    Form_Resize

End Sub

Private Sub Msf�������_S_EnterCell()
    On Error Resume Next
    Dim intCol As Integer
    Dim lngColor As Long
    Dim LngSelectRow As Long
    
    With Msf�������_S
        .Redraw = False
        
        LngSelectRow = .Row     '���浱ǰѡ����
        If mlngRow <> 0 Then
            .Row = mlngRow       '����ϴ�ѡ����
            For intCol = 0 To .Cols - 1
                .Col = intCol
                .CellBackColor = MLNG��ɫ
                .CellForeColor = IIf(.RowData(.Row) = 0, MLNG��ɫ, glng����)
            Next
            .Col = 0
        End If
        
        mlngRow = LngSelectRow
        .Row = mlngRow     '���õ�ǰѡ����
        If Not Me.ActiveControl Is Nothing Then
            For intCol = 0 To .Cols - 1
                .Col = intCol
                .CellBackColor = IIf(Me.ActiveControl.Name = "Msf�������_S", MLNGSEL, MLNG��ɫ)
                If Me.ActiveControl.Name = "Msf�������_S" Then
                    lngColor = IIf(.RowData(.Row) = 0, MLNG��ɫ, glng����)
                Else
                    lngColor = IIf(.RowData(.Row) = 0, MLNG��ɫ, glng����)
                End If
                .CellForeColor = lngColor
            Next
        End If
        .Col = 0
        
        .Redraw = True
    End With
End Sub

Private Sub Msf�������_S_GotFocus()
    Dim intCol As Integer
    With Msf�������_S
        .GridColorFixed = MLNG��ɫ
        .GridColor = MLNG��ɫ
        For intCol = 0 To .Cols - 1
            .Redraw = False
            .Col = intCol
            .CellBackColor = MLNGSEL
            .CellForeColor = IIf(.RowData(.Row) = 0, MLNG��ɫ, glng����)
            .Redraw = True
        Next
        .Col = 0
    End With
    With Msf������Ϣ_S
        For intCol = 0 To .Cols - 1
            .Redraw = False
            .Col = intCol
            .CellBackColor = MLNG��ɫ
            .CellForeColor = IIf(Trim(.TextMatrix(.Row, 16)) = "", MLNG��ɫ, MLNG��ɫ)
            .Redraw = True
        Next
        .Col = 0
    End With
End Sub

Private Sub Msf�������_S_LostFocus()
    Dim intCol As Integer
    With Msf�������_S
        .GridColorFixed = &H80000011
        .GridColor = &H80000011
        For intCol = 0 To .Cols - 1
            .Redraw = False
            .Col = intCol
            .CellBackColor = MLNG��ɫ
            .CellForeColor = IIf(.RowData(.Row) = 0, MLNG��ɫ, glng����)
            .Redraw = True
        Next
        .Col = 0
    End With
End Sub

Private Sub Msf������Ϣ_S_DblClick()
    showReportMXZ
End Sub

Private Sub Msf������Ϣ_S_EnterCell()
    On Error Resume Next
    Dim intCol As Integer
    Dim lngColor As Long
    Dim LngSelectRow As Long
    
    With Msf������Ϣ_S
        .Redraw = False
        
        LngSelectRow = .Row     '���浱ǰѡ����
        If mlngCardRow <> 0 Then
            .Row = mlngCardRow       '����ϴ�ѡ����
            For intCol = 0 To .Cols - 1
                .Col = intCol
                .CellBackColor = MLNG��ɫ
                .CellForeColor = IIf(Trim(.TextMatrix(.Row, 16)) = "", MLNG��ɫ, MLNG��ɫ)
            Next
            .Col = 0
        End If
        
        mlngCardRow = LngSelectRow
        .Row = mlngCardRow       '���õ�ǰѡ����
        If Not ActiveControl Is Nothing Then
            For intCol = 0 To .Cols - 1
                .Col = intCol
                .CellBackColor = IIf(Me.ActiveControl.Name = "Msf������Ϣ_S", MLNGSEL, MLNG��ɫ)
                If Me.ActiveControl.Name = "Msf������Ϣ_S" Then
                    lngColor = IIf(Trim(.TextMatrix(.Row, 16)) = "", MLNG��ɫ, MLNG��ɫ)
                Else
                    lngColor = IIf(Trim(.TextMatrix(.Row, 16)) = "", MLNG��ɫ, MLNG��ɫ)
                End If
                .CellForeColor = lngColor
            Next
        End If
        .Col = 0
        
        .Redraw = True
        
        '��ȡ��������Ϣ
        If Val(.TextMatrix(.Row, 0)) <> 0 Then
            mrsData.MoveFirst
            mrsData.Find "����ID=" & Val(.TextMatrix(.Row, 0))
        End If
    End With
End Sub

Private Sub Msf������Ϣ_S_GotFocus()
    Dim intCol As Integer
    With Msf������Ϣ_S
        .GridColorFixed = MLNG��ɫ
        .GridColor = MLNG��ɫ
        For intCol = 0 To .Cols - 1
            .Redraw = False
            .Col = intCol
            .CellBackColor = MLNGSEL
            .CellForeColor = IIf(Trim(.TextMatrix(.Row, 16)) = "", MLNG��ɫ, MLNG��ɫ)
            .Redraw = True
        Next
        .Col = 0
    End With
    With Msf�������_S
        For intCol = 0 To .Cols - 1
            .Redraw = False
            .Col = intCol
            .CellBackColor = MLNG��ɫ
            .CellForeColor = IIf(.RowData(.Row) = 0, MLNG��ɫ, glng����)
            .Redraw = True
        Next
        .Col = 0
    End With
End Sub

Private Sub Msf������Ϣ_S_LostFocus()
    Dim intCol As Integer
    With Msf������Ϣ_S
        .GridColorFixed = &H80000011
        .GridColor = &H80000011
        For intCol = 0 To .Cols - 1
            .Redraw = False
            .Col = intCol
            .CellBackColor = MLNG��ɫ
            .CellForeColor = IIf(Trim(.TextMatrix(.Row, 16)) = "", MLNG��ɫ, MLNG��ɫ)
            .Redraw = True
        Next
        .Col = 0
    End With
End Sub

Private Sub Msf������Ϣ_S_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim StrHeader As String
    Dim intCol As Integer, intMouseCol As Integer

    'ʵ��������
    If Button = 1 Then
        With Msf������Ϣ_S
            If .MouseRow <> 0 Then Exit Sub
            If mrsData Is Nothing Then Exit Sub
            If mrsData.State = 0 Then Exit Sub
            If mrsData.EOF Then Exit Sub
            
            intMouseCol = .MouseCol
            StrHeader = .TextMatrix(0, intMouseCol)
            If StrHeader = "�������" Then
                StrHeader = "ʵ������"
            ElseIf StrHeader = "�����" Then
                StrHeader = "ʵ�ʽ��"
            ElseIf StrHeader = "�����" Then
                StrHeader = "ʵ�ʲ��"
            End If
            
            If Mid(mstrCardSort, 2) = StrHeader Then
                mstrCardSort = IIf(Mid(mstrCardSort, 1, 1) = "A", "D", "A") & StrHeader
                mrsData.Sort = StrHeader & IIf(Mid(mstrCardSort, 1, 1) = "D", " Desc", " Asc")
            Else
                mstrCardSort = "A" & StrHeader
                mrsData.Sort = StrHeader & " Asc"
            End If
            
            FS.ShowFlash ("���������У����Ժ�...")
            Call SetFormat(False)
            Call Msf������Ϣ_S_EnterCell
            FS.StopFlash
        End With
    Else
        PopupMenu mnuView
    End If
End Sub

Private Sub picVLine_S_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        Me.picVLine_S.Left = Me.picVLine_S.Left + x
        Form_Resize
    End If
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    With Button
        Select Case .Key
        Case "Ԥ��"
            mnuFilePrintView_Click
        Case "��ӡ"
            grdPrint 3
        Case "����"
            ShowReportSumAccount
        Case "��ϸ"
            showReportMXZ
        Case "����"
            showReportCode
        Case "����"
            mnuViewFind_Click
        Case "ˢ��"
            mnuViewRefresh_Click
        Case "����"
            ShowReportSumAccount
        Case "����"
             PopupMenu mnuViewFont
        Case "ǰ��ɫ"
            mnuViewForeColor_Click
        Case "����ɫ" '
            mnuViewBackColor_Click
        Case "����"
            mnuHelpTitle_Click
        Case "�˳�"
           mnufileexit_Click
        End Select
    End With
End Sub

Private Sub tbrThis_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu mnuViewToolbar
    End If
End Sub

Private Sub tvwSection_S_GotFocus()
    If Me.tvwSection_S.Tag = "T" Then Me.tvwSection_S.Tag = "F"
End Sub

Private Sub tvwSection_S_NodeClick(ByVal Node As MSComctlLib.Node)
    If Me.tvwSection_S.Tag = "T" Then Exit Sub
    ReFreshStuffData Me.cob�ⷿ.ItemData(Me.cob�ⷿ.ListIndex), mstrStartDate, mstrEndDate, IIf(Left(Me.tvwSection_S.SelectedItem.Key, 1) = "R", 0, Mid(Me.tvwSection_S.SelectedItem.Key, 2))
End Sub

Private Function ReFreshTreeView() As Boolean
    '-------------------------------------------------------------------------
    '--����:���»�ȡ�����ͽṹ����
    '--����:
    '--����:������ݿ�򿪳ɹ�,��True,���򷵻�False
    '-------------------------------------------------------------------------
    Dim objNode As Node
    Dim RecDept As New ADODB.Recordset
    Dim RecStuff As New ADODB.Recordset
    Dim str���� As String
    
    ReFreshTreeView = False
    
    On Error GoTo ErrHand
    
    gstrSQL = "" & _
        "   Select distinct a.ID,a.���� || '-' || a.���� As ���� " & _
        "   From ���ű� a,��������˵�� b,�������ʷ��� C " & _
        "   Where a.id=b.����id And b.��������=c.���� And C.���� In ('�Ƽ���', '���Ŀ�', '���ϲ���', '����ⷿ') " & _
        "       And (a.վ��=[2] or a.վ�� is null) " & _
                IIf(InStr(1, mstrPrivs, "���пⷿ") <> 0, "", " And A.id In (Select ����ID From ������Ա Where ��ԱID=[1])") & _
        " and (to_char(a.����ʱ��,'yyyy-mm-dd')='3000-01-01' or a.����ʱ�� is null) " & _
        " Order by a.���� || '-' || a.���� "
        
    Set RecDept = zlDatabase.OpenSQLRecord(gstrSQL, "���пⷿ", UserInfo.Id, gstrNodeNo)
    
    With RecDept
          
        If .RecordCount = 0 Then
            MsgBox "�ⷿ���ϲ�����ϵδ������Ȩ�޲��㣬����ִ�б�����!", vbInformation, gstrSysName
            Exit Function
        End If
        
        If InStr(1, mstrPrivs, "���пⷿ") <> 0 Then
            Me.cob�ⷿ.Clear
            Me.cob�ⷿ.AddItem "���пⷿ"
            Me.cob�ⷿ.ItemData(Me.cob�ⷿ.NewIndex) = 0
            Me.cob�ⷿ.ListIndex = Me.cob�ⷿ.NewIndex
        End If
        Do While Not .EOF
            Me.cob�ⷿ.AddItem .Fields("����").Value
            Me.cob�ⷿ.ItemData(Me.cob�ⷿ.NewIndex) = .Fields("ID").Value
            .MoveNext
        Loop
        Me.cob�ⷿ.ListIndex = 0
    End With
    
    Set mrsTreeData = New ADODB.Recordset
    gstrSQL = "" & _
        "   Select id,�ϼ�id,����,����" & _
        "   From ���Ʒ���Ŀ¼ " & _
        "   where ����=7" & _
        "   start with �ϼ�id is null " & _
        "   connect by prior id=�ϼ�id " & _
        "   Order by level,id"
    
    Set mrsTreeData = zlDatabase.OpenSQLRecord(gstrSQL, "���ķ���")
    
    With mrsTreeData
        If .RecordCount = 0 Then
            MsgBox "���ķ�����ϵδ����������ִ�б�����!", vbInformation, gstrSysName
            Exit Function
        End If
        Me.tvwSection_S.Nodes.Clear
        tvwSection_S.Nodes.Add , , "Root", "���з���", "root", "root"
        Do While Not .EOF
            
            If IsNull(.Fields("�ϼ�id").Value) Then
                Set objNode = Me.tvwSection_S.Nodes.Add("Root", 4, "_" & .Fields("id").Value, zlStr.NVL(!����) & " -" & .Fields("����").Value, "child")
            Else
                Set objNode = Me.tvwSection_S.Nodes.Add("_" & .Fields("�ϼ�id").Value, 4, "_" & .Fields("id").Value, zlStr.NVL(!����) & " -" & .Fields("����").Value, "child")
            End If
            .MoveNext
         Loop
         
         If tvwSection_S.Nodes(1).Children <> 0 Then
            tvwSection_S.Nodes(1).Child.Selected = True
         Else
            tvwSection_S.Nodes(1).Selected = True
         End If
    End With
    ReFreshTreeView = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Unload Me
End Function

Private Sub ReFreshStuffData(ByVal lngDeptId As Long, strStartDate As String, strEndDate As String, lngUseId As Long)
    '-------------------------------------------------------------------------
    '--����:���»�ȡ�Ĳ��Ͽ����
    '--����:
    '       lngDeptId:���Ϸ�id
    '       strStartDate:��ʼ����
    '       strEndDate:��������
    '       lngUseId:��;idֵ
    '--����:
    '-------------------------------------------------------------------------
    Dim gstrSQL1 As String, strSQL As String
    Dim intRow As Long
    Dim bln���� As Long
    Dim intCol As Long
    Dim ite As ListItem
    gstrSQL1 = ""
    
    Call FS.ShowFlash("���ڲ�������,���Ժ� ...", Me)
    DoEvents
    
      
   If lngDeptId = 0 Then
        Select Case mbytUint
        Case 0
            gstrSQL = ",Q.���㵥λ as ��λ,'' as �ϴβɹ���,decode(Q.�Ƿ���,1,decode(m.�ϴ��ۼ�,Null, m.ָ�����ۼ�,m.�ϴ��ۼ�),nvl(P.�ּ�,0)) as ����ۼ�," & _
            " 1 as ϵ��,Sum(B.��������) As ��������,Sum(B.ʵ������) As ʵ������" & _
                      ",Sum(B.ʵ�ʽ��) As ʵ�ʽ��,Sum(B.ʵ�ʲ��) As ʵ�ʲ��" & _
                      ",Decode(To_Char(Q.����ʱ��,'yyyy-MM-dd'),'3000-01-01',' ',To_Char(   Q.����ʱ��,'yyyy-MM-dd')) ����ʱ��,1 as ����, g.���� as ��Ӧ�� "
            
            gstrSQL1 = " Group by M.����ID,L.����id,Q.����,Q.����,Q.���,Q.�Ƿ���,Q.����,M.�ⷿ����,Q.���㵥λ,p.�ּ�,nvl(P.�ּ�,0)" & _
                       ",nvl(M.����ϵ��,0),m.�ϴ��ۼ�,m.ָ�����ۼ�,Decode(To_Char(Q.����ʱ��,'yyyy-MM-dd'),'3000-01-01',' ',To_Char(Q.����ʱ��,'yyyy-MM-dd')), g.���� order by q.���� "
            
        Case Else
            gstrSQL = ",M.��װ��λ as ��λ,'' as �ϴβɹ���,decode(Q.�Ƿ���,1,decode(m.�ϴ��ۼ�,Null, m.ָ�����ۼ�,m.�ϴ��ۼ�),nvl(P.�ּ�,0))*nvl(M.����ϵ��,0) as ����ۼ�, " & _
            " nvl(M.����ϵ��,0) as ϵ��" & _
                      ",Sum(B.��������/Decode(M.����ϵ��,0,1,null,1,M.����ϵ��)) as ��������" & _
                      ",Sum(B.ʵ������/Decode(M.����ϵ��,0,1,null,1,M.����ϵ��)) as ʵ������,Sum(B.ʵ�ʽ��) As ʵ�ʽ��" & _
                      ",Sum(B.ʵ�ʲ��) As ʵ�ʲ��,Decode(To_Char(q.����ʱ��,'yyyy-MM-dd'),'3000-01-01',' '" & _
                      ",To_Char(q.����ʱ��,'yyyy-MM-dd')) ����ʱ��,Decode(M.����ϵ��,0,1,null,1,M.����ϵ��) as ����, g.���� as ��Ӧ�� "
            
            gstrSQL1 = " Group by M.����ID,l.����id,Q.����,Q.����,Q.���,Q.�Ƿ���,Q.����,M.�ⷿ����,M.��װ��λ,p.�ּ�,nvl(P.�ּ�,0)*nvl(M.����ϵ��,0)" & _
                       ",nvl(M.����ϵ��,0),m.�ϴ��ۼ�,m.ָ�����ۼ�,Decode(M.����ϵ��,0,1,null,1,M.����ϵ��)" & _
                       ",Decode(To_Char(Q.����ʱ��,'yyyy-MM-dd'),'3000-01-01',' ',To_Char(Q.����ʱ��,'yyyy-MM-dd')), g.���� order by q.���� "
        End Select
    Else
        Select Case mbytUint
        Case 0
            gstrSQL = ",Q.���㵥λ as ��λ,Decode(M.�ⷿ����,0,Avg(S.�ϴβɹ���),Null) as �ϴβɹ���,decode(Q.�Ƿ���,1" & _
                      ",decode(m.�ϴ��ۼ�,Null, m.ָ�����ۼ�,m.�ϴ��ۼ�) ,nvl(P.�ּ�,0)) as ����ۼ�,1 as ϵ��" & _
                      ",Sum(S.��������) as ��������, Sum(S.ʵ������) as ʵ������,Sum(S.ʵ�ʽ��) as ʵ�ʽ��,Sum(S.ʵ�ʲ��) as ʵ�ʲ��" & _
                      ",Decode(To_Char(q.����ʱ��,'yyyy-MM-dd'),'3000-01-01',' ',To_Char(q.����ʱ��,'yyyy-MM-dd')) ����ʱ��,1 as ����, g.���� as ��Ӧ�� "
            
            gstrSQL1 = " Group by M.����ID,Q.����,L.����id,Q.����,Q.���,Q.�Ƿ���,Q.����,nvl(M.���Ч��,0),M.�ⷿ����,Q.���㵥λ,p.�ּ�" & _
                       ",nvl(P.�ּ�,0),nvl(M.����ϵ��,0),m.�ϴ��ۼ�,m.ָ�����ۼ�,Decode(To_Char(Q.����ʱ��,'yyyy-MM-dd'),'3000-01-01',' ',To_Char(Q.����ʱ��,'yyyy-MM-dd')), g.���� order by q.���� "
        Case Else
            gstrSQL = ",M.��װ��λ as ��λ,Decode(M.�ⷿ����,0,Avg(S.�ϴβɹ���*nvl(M.����ϵ��,0)),Null) as �ϴβɹ���,decode(Q.�Ƿ���,1," & _
                      "decode(m.�ϴ��ۼ�,Null, m.ָ�����ۼ�,m.�ϴ��ۼ�) ,nvl(P.�ּ�,0))*nvl(M.����ϵ��,0) as ����ۼ�," & _
                      "nvl(M.����ϵ��,0) as ϵ��,Sum(S.�������� /Decode(M.����ϵ��,0,1,null,1,M.����ϵ��)) as ��������" & _
                      ",Sum(S.ʵ������ /Decode(M.����ϵ��,0,1,null,1,M.����ϵ��)) as ʵ������,Sum(S.ʵ�ʽ��) as ʵ�ʽ��" & _
                      ",Sum(S.ʵ�ʲ��) as ʵ�ʲ��,Decode(To_Char(q.����ʱ��,'yyyy-MM-dd'),'3000-01-01',' ',To_Char(q.����ʱ��,'yyyy-MM-dd')) ����ʱ��" & _
                      ",Decode(M.����ϵ��,0,1,null,1,M.����ϵ��) as ����, g.���� as ��Ӧ�� "
            
            gstrSQL1 = " Group by M.����ID,Q.����,L.����id,Q.����,Q.���,Q.�Ƿ���,Q.����,nvl(M.���Ч��,0),M.�ⷿ����,M.��װ��λ,P.�ּ�" & _
                       ",M.����ϵ��,m.�ϴ��ۼ�,m.ָ�����ۼ�,nvl(P.�ּ�,0)*nvl(M.����ϵ��,0),Decode(To_Char(Q.����ʱ��,'yyyy-MM-dd'),'3000-01-01',' ',To_Char(Q.����ʱ��,'yyyy-MM-dd')), g.���� order by q.���� "
        End Select
    End If
    
    Set mrsData = New ADODB.Recordset
    On Error GoTo ErrHand:
    
    If lngDeptId = 0 Then
        strSQL = "Select  M.����ID,L.����id,Q.����,Q.���� as ����,Q.���,Q.����,Null as Ч��,Decode(M.�ⷿ����,1,'��','��') as �ⷿ���� " & gstrSQL & _
                " From �������� M,�շ���ĿĿ¼ Q,������ĿĿ¼ L ,�շѼ�Ŀ P ,��Ӧ�� G, " & _
                "     (SELECT a.�ⷿid, a.ҩƷid, a.����, a.Ч��, a.����, a.��������, a.ʵ������, a.ʵ�ʽ��, a.ʵ�ʲ��, a.�ϴι�Ӧ��id, a.�ϴβɹ���, a.�ϴ�����, a.�ϴ���������" & _
                "           ,a.�ϴβ���, a.���Ч��, a.��׼�ĺ�, a.���ۼ�, a.�ϴο��� " & _
                "       FROM ҩƷ��� A, �������� B, ������ĿĿ¼ C WHERE a.ҩƷid = b.����id And b.����id = c.Id And a.����=1 " & _
                IIf(tvwSection_S.SelectedItem.Key = "Root", "", "    And C.����id in ( Select id From ���Ʒ���Ŀ¼ Q start with Q.id= [1] connect by prior id=�ϼ�id)") & " ) B " & _
                " Where m.�ϴι�Ӧ��id = g.Id(+) And P.�շ�ϸĿId=M.����id and M.����ID=L.id And (Q.վ��=[3] or Q.վ�� is null) " & _
                "       And P.�շ�ϸĿid=Q.id " & IIf(mbln����ͣ��, "", " and (TO_CHAR(Q.����ʱ��, 'yyyy-mm-dd') = '3000-01-01' OR Q.����ʱ�� IS NULL) ") & _
                "       And sysdate between P.ִ������ And nvl(P.��ֹ����,To_Date('3000-01-01','yyyy-MM-DD')) and  M.����id=B.ҩƷid(+)  " & _
                GetPriceClassString("P") & _
                IIf(mbln�����, " And  B.ʵ������<>0 ", "") & IIf(tvwSection_S.SelectedItem.Key = "Root", "", "    And L.����id in ( Select id From ���Ʒ���Ŀ¼ Q start with Q.id= [1] connect by prior id=�ϼ�id)")
        strSQL = strSQL + gstrSQL1
    Else
        strSQL = "Select M.����ID,L.����id,Q.����,Q.���� as ����,Q.���,Q.����,nvl(M.���Ч��,0) as Ч��,Decode(M.�ⷿ����,1,'��','��') as �ⷿ���� " & gstrSQL & _
                " From �������� M,�շ���ĿĿ¼ Q,������ĿĿ¼ L,�շѼ�Ŀ P, ��Ӧ�� G, " & _
                "      (Select a.ҩƷID,a.�ϴβɹ���,sum(a.��������) as ��������,sum(a.ʵ������) as ʵ������,sum(a.ʵ�ʽ��) as ʵ�ʽ��,sum(a.ʵ�ʲ��) as ʵ�ʲ�� " & _
                "       From ҩƷ��� A, �������� B, ������ĿĿ¼ C Where a.ҩƷid = b.����id And b.����id = c.Id And a.�ⷿid=[2] And a.����=1 " & _
                IIf(tvwSection_S.SelectedItem.Key = "Root", "", "    And C.����id in ( Select id From ���Ʒ���Ŀ¼ Q start with Q.id= [1] connect by prior id=�ϼ�id)") & _
                " Group by a.ҩƷID,a.�ϴβɹ���) S,(Select Distinct �շ�ϸĿid, ִ�п���id From �շ�ִ�п��� Where ִ�п���id = [2]) K " & _
                " Where m.�ϴι�Ӧ��id = g.Id(+) And M.����id=P.�շ�ϸĿId and M.����ID=L.id  and P.�շ�ϸĿid=Q.id And (Q.վ��=[3] or Q.վ�� is null) And m.����id=k.�շ�ϸĿid " & _
                "       And M.����id=S.ҩƷid(+) " & IIf(mbln����ͣ��, "", " and (TO_CHAR(Q.����ʱ��, 'yyyy-mm-dd') = '3000-01-01' OR Q.����ʱ�� IS NULL)  ") & _
                "       And sysdate between P.ִ������ And nvl(P.��ֹ����,To_Date('3000-01-01','yyyy-MM-DD')) " & _
                GetPriceClassString("P") & _
                    IIf(mbln�����, " And S.ʵ������<>0 ", "") & IIf(tvwSection_S.SelectedItem.Key = "Root", "", " And L.����id in ( Select id From ���Ʒ���Ŀ¼ Q start with Q.id= [1] connect by prior id=�ϼ�id)")
        strSQL = strSQL + gstrSQL1
    End If
    gstrSQL = strSQL
        
    Set mrsData = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, lngUseId, lngDeptId, gstrNodeNo)
    
    With mrsData
        If .RecordCount = 0 Then
            mnuExcel.Enabled = False
            mnuFilePrint.Enabled = False
            mnuFilePrintView.Enabled = False
            mnuViewFind.Enabled = False
            tbrThis.Buttons.Item(1).Enabled = False
            tbrThis.Buttons.Item(2).Enabled = False
            tbrThis.Buttons.Item(6).Enabled = False
            'tbrThis.Buttons.Item(7).Enabled = False
        Else
            mnuExcel.Enabled = True
            mnuFilePrint.Enabled = True
            mnuFilePrintView.Enabled = True
            mnuViewFind.Enabled = True
            tbrThis.Buttons.Item(1).Enabled = True
            tbrThis.Buttons.Item(2).Enabled = True
            tbrThis.Buttons.Item(6).Enabled = True
           ' tbrThis.Buttons.Item(7).Enabled = True
        End If
        Call SetReportCtrlIndexEnabled
        
        Call FS.StopFlash
        Call SetFormat(False)
    End With
    
    With Msf������Ϣ_S
        .Row = 1
        Call Msf������Ϣ_S_EnterCell
    End With
    Exit Sub
ErrHand:
    Call FS.StopFlash
    If ErrCenter() = 1 Then
        Resume
    End If
    Exit Sub
End Sub

Private Sub RefreshBatch(lng�ⷿID As Long, lng����ID As Long)
    '-------------------------------------------------------------------------
    '--����:���»�ȡ�����ķ��������
    '--����:
    '       lng�ⷿId:���Ϸ�id
    '       lng����ID:��;idֵ
    '--����:
    '-------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim intRow As Long
    Dim intCol As Long
    Dim lngColor As Long
    
    Dim int���� As Integer
    Dim int���� As Integer
     
    On Error GoTo ErrHand
    Me.Msf�������_S.Redraw = False
    Me.Msf�������_S.Rows = 1
    gstrSQL = "Select 1 From ��������˵�� Where ����id=[1] And ��������  like '���ϲ���'"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���ϲ����ж�", lng�ⷿID)
            
    If rsTemp.EOF Then
        int���� = 0
    Else
        int���� = 1
    End If

    gstrSQL = "" & _
        "   Select Decode(nvl(�ⷿ����,0),1,Decode(Nvl(���÷���,0),1,2,1),0) As ���� " & _
        "   From �������� " & _
        "   Where ����id=[1]"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��������", lng����ID)
        
        
    '����ⷿ���������÷�����int����=2�������ⷿ������int����=1������������int����=0��
    If Not rsTemp.EOF And Not rsTemp.BOF Then
        int���� = Val(zlStr.NVL(rsTemp!����))
        If lng�ⷿID = 0 Or (int���� = 1 And int���� = 2) Or (int���� = 0 And int���� <> 0) Then
            '�����пⷿ ���� �ǿⷿ�ҿ�����������ʾ�ֿⷿ�������
            If lng�ⷿID = 0 Then
                gstrSQL = "" & _
                    "   Select D.���� as �ⷿ,Null As ����,Null As ʧЧ��,0 ����,Null As ����,Null As ������," & _
                    "        Sum(S.��������)/" & mrsData("����").Value & " as ��������," & _
                    "        Sum(S.ʵ������)/" & mrsData("����").Value & " as ʵ������," & _
                    "        Sum(S.ʵ�ʽ��) As ʵ�ʽ��," & _
                    "        Sum(S.ʵ�ʲ��) As ʵ�ʲ��," & _
                    "       avg(s.ƽ���ɱ���)* " & mrsData("����").Value & " as ƽ���ɱ���," & _
                    "        Null As ��������,Null As No,Null as ������λ, " & _
                    "        decode(sum(s.ʵ������),0,0,sum(s.ʵ�ʽ��)/sum(s.ʵ������)) as �ۼ� " & _
                    "   From ҩƷ��� S,���ű� D  " & _
                    "   Where S.�ⷿid=D.id And S.����=1 And S.ҩƷid=[2]" & _
                    "       And (S.ʵ������<>0 or S.ʵ�ʽ��<>0 or S.ʵ�ʲ��<>0)" & _
                    " Group By D.���� "
            Else
               gstrSQL = "Select D.���� as �ⷿ,s.�ϴ����� As ����, s.Ч�� as ʧЧ��, s.�ϴβ��� As ����,Decode(sign(Add_Months(Sysdate," & mintMonths & ")-s.Ч��),-1,0,1) ����," & _
                        "        S.��������/" & mrsData("����").Value & " as ��������,S.ʵ������/" & mrsData("����").Value & " as ʵ������,S.ʵ�ʽ��,S.ʵ�ʲ��," & _
                        "        S.�ϴβɹ���*" & mrsData("����").Value & " as ������,S.��Ʒ����,S.�ڲ�����,s.ƽ���ɱ���*" & mrsData("����").Value & "ƽ���ɱ���," & _
                        "decode(c.�Ƿ���,1,decode(s.����,null,decode(s.ʵ������,0,0,s.ʵ�ʽ��/s.ʵ������),s.���ۼ� ),b.�ּ�) * " & mrsData("����").Value & " as �ۼ� , g.���� As ��Ӧ�� " & _
                        " From ҩƷ��� S,���ű� D,�������� A,�շѼ�Ŀ B,�շ���ĿĿ¼ C, ��Ӧ�� G " & _
                        " Where s.�ϴι�Ӧ��id = g.Id(+) And S.�ⷿid=D.id  And S.ҩƷid=A.����id" & _
                        "       And S.ҩƷid=[2] And S.����=1 And S.�ⷿid=[1] and b.�շ�ϸĿid=a.����id and B.�շ�ϸĿID=c.id  and sysdate between b.ִ������ and b.��ֹ���� " & _
                        GetPriceClassString("B") & " And (S.ʵ������<>0 or S.ʵ�ʽ��<>0 or S.ʵ�ʲ��<>0)" & _
                        " order by D.����"
            End If
            
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, lng�ⷿID, lng����ID)
            
            With rsTemp
                Me.Msf�������_S.Rows = .RecordCount + 1
                Do While Not .EOF
                    Me.Msf�������_S.TextMatrix(.AbsolutePosition, 0) = !�ⷿ
                    Me.Msf�������_S.TextMatrix(.AbsolutePosition, 1) = IIf(IsNull(!����), "", !����)
                    Me.Msf�������_S.TextMatrix(.AbsolutePosition, 2) = Format(!ʧЧ��, "yyyy��MM��dd��")
                    Me.Msf�������_S.TextMatrix(.AbsolutePosition, 3) = IIf(IsNull(!����), "", !����)
                    Me.Msf�������_S.TextMatrix(.AbsolutePosition, 4) = Format(!��������, mFMT.FM_����)
                    Me.Msf�������_S.TextMatrix(.AbsolutePosition, 5) = Format(!ʵ������, mFMT.FM_����)
                    Me.Msf�������_S.TextMatrix(.AbsolutePosition, 6) = Format(!ʵ�ʽ��, mFMT.FM_���)
'                    Me.Msf�������_S.TextMatrix(.AbsolutePosition, 7) = Format(((NVL(!ʵ�ʽ��, 0) - NVL(!ʵ�ʲ��, 0)) / IIf(NVL(!ʵ������, 0) = 0, 1, NVL(!ʵ������, 1))), mFMT.FM_�ɱ���)
'                    Me.Msf�������_S.TextMatrix(.AbsolutePosition, 8) = Format((NVL(!ʵ�ʽ��, 0) - NVL(!ʵ�ʲ��, 0)), mFMT.FM_���)
                    Me.Msf�������_S.TextMatrix(.AbsolutePosition, 7) = Format(!ƽ���ɱ���, mFMT.FM_�ɱ���) '�ɱ���
                    Me.Msf�������_S.TextMatrix(.AbsolutePosition, 8) = Format(!ƽ���ɱ��� * !ʵ������, mFMT.FM_���) '�ɱ����
                    Me.Msf�������_S.TextMatrix(.AbsolutePosition, 9) = Format(!ʵ�ʲ��, mFMT.FM_���)
                    Me.Msf�������_S.TextMatrix(.AbsolutePosition, 10) = Format(!������, mFMT.FM_�ɱ���)
                    
                    If gblnCode = True And lng�ⷿID > 0 Then
                        Me.Msf�������_S.TextMatrix(.AbsolutePosition, 11) = IIf(IsNull(!��Ʒ����), "", !��Ʒ����)
                        Me.Msf�������_S.TextMatrix(.AbsolutePosition, 12) = IIf(IsNull(!�ڲ�����), "", !�ڲ�����)
                    End If
                    If gblnCode = True Then
                        Me.Msf�������_S.TextMatrix(.AbsolutePosition, 13) = Format(!�ۼ�, mFMT.FM_���ۼ�)
                        If lng�ⷿID <> 0 Then
                            Me.Msf�������_S.ColWidth(14) = 1500
                            Me.Msf�������_S.TextMatrix(.AbsolutePosition, 14) = IIf(IsNull(!��Ӧ��), "", !��Ӧ��)
                        Else
                            Me.Msf�������_S.ColWidth(14) = 0
                        End If
                    Else
                        Me.Msf�������_S.TextMatrix(.AbsolutePosition, 11) = Format(!�ۼ�, mFMT.FM_���ۼ�)
                        If lng�ⷿID <> 0 Then
                            Me.Msf�������_S.ColWidth(12) = 1500
                            Me.Msf�������_S.TextMatrix(.AbsolutePosition, 12) = IIf(IsNull(!��Ӧ��), "", !��Ӧ��)
                         Else
                            Me.Msf�������_S.ColWidth(12) = 0
                        End If
                    End If
                    
                    Me.Msf�������_S.RowData(.AbsolutePosition) = !����
                    '���ݼ�¼״̬�Ĳ�ͬ��������ɫ
                    lngColor = IIf(!���� = 0, glng����, glng����)
                    For intCol = 0 To Msf�������_S.Cols - 1
                        Msf�������_S.Col = intCol
                        Msf�������_S.Row = .AbsolutePosition
                        Msf�������_S.CellForeColor = lngColor
                    Next
                    .MoveNext
                Loop
            End With
        End If
    End If
    If lng�ⷿID = 0 Then
        Me.Msf�������_S.ColWidth(0) = 1000
        Me.Msf�������_S.ColWidth(1) = 0
        Me.Msf�������_S.ColWidth(2) = 0
        Me.Msf�������_S.ColWidth(3) = 0
        Me.Msf�������_S.ColWidth(10) = 0
        Me.Msf�������_S.ColWidth(11) = 0
        Me.Msf�������_S.ColWidth(12) = 0
    Else
        Me.Msf�������_S.ColWidth(0) = 0
        Me.Msf�������_S.ColWidth(1) = 1500
        Me.Msf�������_S.ColWidth(2) = 1500
        Me.Msf�������_S.ColWidth(3) = 1500
        Me.Msf�������_S.ColWidth(10) = 0
        Me.Msf�������_S.ColWidth(11) = 1800
        Me.Msf�������_S.ColWidth(12) = 1800
    End If
    If mblnCostView = False Then
        Me.Msf�������_S.ColWidth(7) = 0
        Me.Msf�������_S.ColWidth(8) = 0
        Me.Msf�������_S.ColWidth(9) = 0
        Me.Msf�������_S.ColWidth(10) = 0
    End If
    If Me.Msf�������_S.Rows = 1 Then
        Me.Msf�������_S.Visible = False
        Me.lbl����_S.Visible = False
        Me.Msf�������_S.Rows = 2
    Else
        Me.Msf�������_S.Visible = True
        Me.lbl����_S.Visible = True
    End If
    Me.Msf�������_S.FixedRows = 1
    Me.Msf�������_S.Redraw = True
    Call Form_Resize
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Exit Sub
End Sub

Private Function ReFreshFilterData(ByVal lngDeptId As Long, strFind As String) As Boolean
    '-------------------------------------------------------------------------
    '--����:������ָ��������
    '--����:
    '       lngDeptId:���Ϸ�id
    '       strFind:�������
    '--����:
    '-------------------------------------------------------------------------
    Dim gstrSQL1 As String
    Dim intRow As Long
    Dim bln���� As Long
    Dim intCol As Long
    Dim rsTemp As New ADODB.Recordset
    Dim str���� As String
    Dim ite As ListItem
    gstrSQL1 = ""
    On Error GoTo ErrHand:
    
    Call FS.ShowFlash("���ڲ�������,���Ժ� ...", Me)
    DoEvents
    
    ReFreshFilterData = False
    If lngDeptId = 0 Then
        Select Case mbytUint
            Case 0
                gstrSQL = ",Q.���㵥λ as ��λ,0 as �ϴβɹ���,Decode(q.�Ƿ���, 1,decode(m.�ϴ��ۼ�,Null, m.ָ�����ۼ�,m.�ϴ��ۼ�), Nvl(p.�ּ�, 0)) as ����ۼ�,nvl(M.����ϵ��,0) as ϵ��,(B.��������) as ��������" & _
                          ",(B.ʵ������) as ʵ������,(B.ʵ�ʽ��) as ʵ�ʽ��,(B.ʵ�ʲ��) as ʵ�ʲ��" & _
                          ",Decode(To_Char(Q.����ʱ��,'yyyy-MM-dd'),'3000-01-01',' ',To_Char(Q.����ʱ��,'yyyy-MM-dd')) ����ʱ��,1 as ����, g.���� as ��Ӧ�� "
            Case Else
                gstrSQL = ",M.��װ��λ as ��λ,0 as �ϴβɹ���,Decode(q.�Ƿ���, 1,decode(m.�ϴ��ۼ�,Null, m.ָ�����ۼ�,m.�ϴ��ۼ�), Nvl(p.�ּ�, 0))*nvl(M.����ϵ��,0) as ����ۼ�,nvl(M.����ϵ��,0) as ϵ��" & _
                          ",(B.��������/Decode(M.����ϵ��,0,1,null,1,M.����ϵ��)) as ��������" & _
                          ",(B.ʵ������/Decode(M.����ϵ��,0,1,null,1,M.����ϵ��)) as ʵ������" & _
                          ",(B.ʵ�ʽ��) As ʵ�ʽ��,(B.ʵ�ʲ��) As ʵ�ʲ��" & _
                          ",Decode(To_Char(Q.����ʱ��,'yyyy-MM-dd'),'3000-01-01',' ',To_Char(Q.����ʱ��,'yyyy-MM-dd')) ����ʱ��" & _
                          ",Decode(M.����ϵ��,0,1,null,1,M.����ϵ��) as ����, g.���� as ��Ӧ�� "
        End Select
    Else
        Select Case mbytUint
        Case 0
            gstrSQL = ",Q.���㵥λ as ��λ,S.�ϴβɹ���,Decode(q.�Ƿ���, 1,decode(m.�ϴ��ۼ�,Null, m.ָ�����ۼ�,m.�ϴ��ۼ�), Nvl(p.�ּ�, 0)) as ����ۼ�,nvl(M.����ϵ��,0) as ϵ��,(S.��������) as ��������" & _
                      ",(S.ʵ������) as ʵ������,(S.ʵ�ʽ��) as ʵ�ʽ��,(S.ʵ�ʲ��) as ʵ�ʲ��" & _
                      ",Decode(To_Char(Q.����ʱ��,'yyyy-MM-dd'),'3000-01-01',' ',To_Char(Q.����ʱ��,'yyyy-MM-dd')) ����ʱ��,1 as ����, g.���� as ��Ӧ��  "
        Case Else
            gstrSQL = ",M.��װ��λ as ��λ,S.�ϴβɹ���*nvl(M.����ϵ��,0) as �ϴβɹ���,Decode(q.�Ƿ���, 1,decode(m.�ϴ��ۼ�,Null, m.ָ�����ۼ�,m.�ϴ��ۼ�), Nvl(p.�ּ�, 0))*nvl(M.����ϵ��,0) as ����ۼ�" & _
                      ",nvl(M.����ϵ��,0) as ϵ��,S.�������� /Decode(M.����ϵ��,0,1,null,1,M.����ϵ��) as ��������" & _
                      ",S.ʵ������ /Decode(M.����ϵ��,0,1,null,1,M.����ϵ��) as ʵ������,S.ʵ�ʽ�� as ʵ�ʽ��,S.ʵ�ʲ�� as ʵ�ʲ��" & _
                      ",Decode(To_Char(Q.����ʱ��,'yyyy-MM-dd'),'3000-01-01',' ',To_Char(Q.����ʱ��,'yyyy-MM-dd')) ����ʱ��" & _
                      ",Decode(M.����ϵ��,0,1,null,1,M.����ϵ��) as ����, g.���� as ��Ӧ�� "
        End Select
    End If
    
    If lngDeptId = 0 Then
       gstrSQL = "" & _
            "   Select distinct B.�ⷿID,M.����ID,B.����,L.����id,Q.����,Q.���� as ����,Q.���,Q.����,nvl(M.���Ч��,0) as Ч��" & _
            "       ,Decode(M.�ⷿ����,1,'��','��') as �ⷿ���� " & gstrSQL & _
            "   From �������� M,�շ���ĿĿ¼ Q,������ĿĿ¼ L ,�շѼ�Ŀ P ,��Ӧ�� G, " & _
            "       (select �ⷿid, ҩƷid, ����, Ч��, ����, ��������, ʵ������, ʵ�ʽ��, ʵ�ʲ��, �ϴι�Ӧ��id, �ϴβɹ���" & _
            "             ,�ϴ�����, �ϴ���������, �ϴβ���, ���Ч��, ��׼�ĺ�, ���ۼ�, �ϴο��� from ҩƷ��� where ����=1) B " & _
            "   Where m.�ϴι�Ӧ��id = g.Id(+) And P.�շ�ϸĿId=M.����id  and M.����id=L.id and P.�շ�ϸĿid=Q.id And (Q.վ��=[6] or Q.վ�� is null) And M.����id=B.ҩƷid(+) " & _
                        IIf(mbln����ͣ��, "", " and (TO_CHAR(Q.����ʱ��, 'yyyy-mm-dd') = '3000-01-01' OR Q.����ʱ�� IS NULL) ") & _
            "           And sysdate between P.ִ������ And nvl(P.��ֹ����,To_Date('3000-01-01','yyyy-MM-DD'))  " & _
            GetPriceClassString("P") & _
                        IIf(mbln�����, "  And B.ʵ������<>0 ", "") & IIf(strFind = "", "", " And " & strFind)
        
        gstrSQL = "" & _
            "   SELECT ����ID,����id,����,����,���,����,Ч��,�ⷿ����,��λ,max(�ϴβɹ���) �ϴβɹ���,����ۼ�," & _
            "           ϵ��,SUM(��������) AS �������� , SUM(ʵ������) AS ʵ������ ,SUM(ʵ�ʽ��) AS ʵ�ʽ��," & _
            "           SUM(ʵ�ʲ��) AS ʵ�ʲ��,����ʱ��,����,��Ӧ�� " & _
            "   From (" & gstrSQL & ") " & _
            " GROUP BY ����ID,����id,����,����,���,����,Ч��,�ⷿ����,����ۼ�,ϵ�� ,����ʱ��, ����,��λ,��Ӧ�� Order By ���� "
    Else
       gstrSQL = "" & _
            "   Select Distinct M.����ID,S.����,Q.����,L.����id,Q.���� as ����,Q.���,Q.����,nvl(M.���Ч��,0) as Ч��" & _
            "       ,Decode(M.�ⷿ����,1,'��','��') as �ⷿ���� " & gstrSQL & _
            "   From �������� M,�շ���ĿĿ¼ Q,������ĿĿ¼ L ,�շѼ�Ŀ P ,��Ӧ�� G, " & _
            "       (Select �ⷿID,ҩƷid ����ID,����,�ϴβɹ���,sum(��������) as ��������, sum(ʵ������) as ʵ������,sum(ʵ�ʽ��) as ʵ�ʽ��" & _
            "           ,sum(ʵ�ʲ��) as ʵ�ʲ�� From ҩƷ��� Where �ⷿid+0=[1] And ����=1 " & _
            "         Group by �ⷿID,ҩƷID,����,�ϴβɹ���) S " & _
            "   Where m.�ϴι�Ӧ��id = g.Id(+) And P.�շ�ϸĿId=M.����id and M.����id=L.id and P.�շ�ϸĿid=Q.id And (Q.վ��=[6] or Q.վ�� is null) " & _
            "       And M.����id=S.����id(+)" & IIf(mbln����ͣ��, "", " and (TO_CHAR(Q.����ʱ��, 'yyyy-mm-dd') = '3000-01-01' OR Q.����ʱ�� IS NULL) ") & _
            "       And sysdate between P.ִ������ And nvl(P.��ֹ����,To_Date('3000-01-01','yyyy-MM-DD')) " & _
            GetPriceClassString("P") & _
                    IIf(mbln�����, " And S.ʵ������<>0 And S.ʵ������ is not null ", "") & " And " & strFind
        
       gstrSQL = "" & _
            "   SELECT ����ID,����id,����,����,���,����,Ч��,�ⷿ����,��λ,Decode(�ⷿ����,'��',Avg(�ϴβɹ���),Null) �ϴβɹ���,����ۼ�," & _
             "      ϵ��,SUM(��������) AS �������� , SUM(ʵ������) AS ʵ������ ,SUM(ʵ�ʽ��) AS ʵ�ʽ��," & _
            "       SUM(ʵ�ʲ��) AS ʵ�ʲ��,����ʱ��,����,��Ӧ�� " & _
            "   from (" & gstrSQL & ") " & _
            "   GROUP BY ����ID,����,����id,����,���,����,Ч��,�ⷿ����,����ۼ�,ϵ�� ,����ʱ��, ����,��λ ,��Ӧ�� Order By ���� "

    End If
        
   '0-����,1-����,2-����,3-���,4-����,5-ָ������,6-վ��
    '����:[1]�ⷿ,[2]-���� ,[3]-����,[4]-����,[5]-���,[6]-վ��,[7]-ָ������
        
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, lngDeptId, mstrOthers(0), mstrOthers(1), mstrOthers(2), mstrOthers(3), gstrNodeNo, mstrOthers(5))
    
    Call FS.StopFlash
    With rsTemp
        If .RecordCount = 0 Then
            MsgBox "��ָ�������е����ݲ�����!", vbInformation, gstrSysName
            Exit Function
        End If
        ReFreshFilterData = True
        Set mrsData = rsTemp
      
        Call SetFormat(False)
    End With
    
    With Msf������Ϣ_S
        .Row = 1
        Call Msf������Ϣ_S_EnterCell
    End With
    Exit Function
ErrHand:
    Call FS.StopFlash
    If ErrCenter() = 1 Then Resume
    ReFreshFilterData = False
    Exit Function
End Function


'-------------------------------
'
'����
'
'''''''''''''''''''''''''''''''''

Private Sub mnuBill_Click()
    Dim strNo As String
    Dim byt���� As Integer
    Dim byt��¼״̬ As Integer
          
    Select Case Mid(mstrNoS, 4)
        Case "_INSIDE_1721_1"  '����
            strNo = Mid(Trim(mobjCurSheet.TextMatrix(mobjCurSheet.Row, 3)), 3)
            byt���� = Val(mobjCurSheet.TextMatrix(mobjCurSheet.Row, 1))
            byt��¼״̬ = Val(mobjCurSheet.TextMatrix(mobjCurSheet.Row, 4))
        Case "_INSIDE_1721_2"  '��ϸ��
            strNo = Trim(mobjCurSheet.TextMatrix(mobjCurSheet.Row, 3))
            byt���� = Val(mobjCurSheet.TextMatrix(mobjCurSheet.Row, 2))
            byt��¼״̬ = Val(mobjCurSheet.TextMatrix(mobjCurSheet.Row, 1))
        Case "_INSIDE_1721_3"  '��ϸ��
        
    End Select
    
    If strNo = "" Or byt���� = 0 Or byt��¼״̬ = 99 Then Exit Sub
    If byt���� = 0 Then Exit Sub
    ShowBill Me, strNo, byt��¼״̬, byt����
End Sub

Private Sub mobjReport_ReportActive(ByVal strNo As String, Form As Object)
    mlngCurReport = Form.hwnd
    mstrNoS = strNo
End Sub


Private Sub mobjReport_SheetDblClick(ByVal strNo As String, Sheet As Object, frmParent As Object)
    mlngCurReport = frmParent.hwnd
    mstrNoS = strNo
    Set mobjCurSheet = Sheet
    If Mid(UCase(strNo), 4) = "_INSIDE_1723_3" Then Exit Sub
    mnuBill_Click
End Sub

Private Sub mobjReport_SheetMouseDown(ByVal strNo As String, Button As Integer, Shift As Integer, x As Single, y As Single, Sheet As Object, frmParent As Object)
    mlngCurReport = frmParent.hwnd
    mstrNoS = strNo
    Set mobjCurSheet = Sheet
    If Mid(UCase(strNo), 4) <> "_INSIDE_1723_3" Then
        If Button = 2 Then PopupMenu mnuReportBill, 2
    End If
End Sub

Private Sub SetMenu(ByVal intState As Integer)
    If intState = 0 Then mnuReportBill.Visible = False: Exit Sub
End Sub

Private Sub ShowBill(frmObject As Object, strNo As String, int��¼״̬ As Integer, int���� As Integer, Optional bln���� As Boolean = False)
    '--------------------------------------------------------------------------------------
    '����:��ʾָ������
    '����:
    '       frmObject:����
    '           strNo:���ݺ�
    '     int��¼״̬:����״̬(mod(��¼״̬,3)=1-������¼;mod(��¼״̬,3)=2-������¼;mod(��¼״̬,3)=0-�Ѿ������ļ�¼)
    '         int����:�������( �ⷿ:1-�⹺��ⵥ;2-�������;3-�ƿⵥ;4-����;5-��������;6-�̴�;7-������;
    '                           ����:1-����;2-����;3-���ϵ�;4-Ȩ�����)
    '                           15-�����⹺���,16-�����������,17-�����������,18-���ϲ�۵���,19-�����ƿ�,20-���Ų�������,21-������������,22-�����̵㣬23-�����̵��¼����24-�շѴ������ϣ�25-���ʵ��������ϣ�26-���ʱ������ϣ�
    '--------------------------------------------------------------------------------------
    Dim strPrivsTemp As String
    
    On Error GoTo ErrHandle
    Select Case int����
        Case 15
            strPrivsTemp = GetPrivFunc(glngSys, 1712)
            frmPurchaseCard.ShowCard frmObject, strNo, 4, int��¼״̬, strPrivsTemp
        Case 16
            strPrivsTemp = GetPrivFunc(glngSys, 1713)
            frmSelfMakeCard.ShowCard frmObject, strNo, 4, int��¼״̬, strPrivsTemp
        Case 17
            strPrivsTemp = GetPrivFunc(glngSys, 1714)
            frmOtherInputCard.ShowCard frmObject, strNo, 4, int��¼״̬, strPrivsTemp
        Case 18
            strPrivsTemp = GetPrivFunc(glngSys, 1715)
            frmDiffPriceAdjustCard.ShowCard frmObject, strNo, 4, int��¼״̬, strPrivsTemp
        Case 19
            strPrivsTemp = GetPrivFunc(glngSys, 1716)
            frmTransferCard.ShowCard frmObject, strNo, 4, int��¼״̬, strPrivsTemp
        Case 20
            strPrivsTemp = GetPrivFunc(glngSys, 1717)
            frmDrawCard.ShowCard frmObject, strNo, 4, int��¼״̬, strPrivsTemp
        Case 21
            strPrivsTemp = GetPrivFunc(glngSys, 1718)
            frmOtherOutputCard.ShowCard frmObject, strNo, 4, int��¼״̬, strPrivsTemp
        Case 22
            strPrivsTemp = GetPrivFunc(glngSys, 1719)
            frmCheckCard.ShowCard frmObject, strNo, 4, int��¼״̬, strPrivsTemp
        Case 13
            Dim rsTemp As New ADODB.Recordset
            gstrSQL = "Select id,����,NO,nvl(�۸�id,0) as �۸�id" & _
                " From ҩƷ�շ���¼" & _
                " Where No=[1] And ����=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ�۸��¼ID", strNo, int����)
            
            With rsTemp
                If .EOF Or .BOF Then Exit Sub
            End With
            gstrUserName = UserInfo.�û���
            Call frmStuffPrice.ShowBill(frmObject, B_����, Val(zlStr.NVL(rsTemp!�۸�id)), 0)
'            With frmStuffPrice
'                .mlngBillId = rsTemp!�۸�id
'                .mlngStuffId = 1
'                .Show 1, frmObject
'            End With
        Case Else
            With Frm����See
                .int��¼״̬ = int��¼״̬
                .byt���� = int����
                .strNo = strNo
                .Show 1, frmObject
            End With
    End Select
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ����Ȩ��()

    Dim lngCount As Long
    Dim i As Long
    
    lngCount = 0        'ͳ���Ƿ������ز˵�����
    tbrThis.Buttons("��ϸ").Visible = False
    tbrThis.Buttons("����").Visible = False
    For i = 0 To mnuReportItem.UBound
        If Split(mnuReportItem(i).Tag & ",", ",")(1) = "ZL1_INSIDE_1721_2" Then
            tbrThis.Buttons("��ϸ").Visible = True
            lngCount = lngCount + 1
        End If
        If Split(mnuReportItem(i).Tag & ",", ",")(1) = "ZL1_INSIDE_1721_1" Then
            tbrThis.Buttons("����").Visible = True
            lngCount = lngCount + 1
        End If
    Next
End Sub
Private Function ISCheckReport(ByVal strReportCode As String) As Boolean
    '����:���ָ�������Ƿ���Ȩ��
    '����:strReportCode-������
    Dim i As Long
    
    For i = 0 To mnuReportItem.UBound
        If Split(mnuReportItem(i).Tag & ",", ",")(1) = strReportCode Then
            ISCheckReport = mnuReportItem(i).Enabled And mnuReport.Visible
            Exit Function
        End If
    Next
    ISCheckReport = False
End Function

Private Sub SetFormat(ByVal BlnSetHeader As Boolean)
    On Error Resume Next
    
    Dim intCol As Integer
    With Msf������Ϣ_S
        .Clear
        .Rows = 2
        .Cols = 19
        .TextMatrix(0, 0) = "����ID"
        .TextMatrix(0, 1) = "����ID"
        .TextMatrix(0, 2) = "����"
        .TextMatrix(0, 3) = "����"
        .TextMatrix(0, 4) = "���"
        .TextMatrix(0, 5) = "����"
        .TextMatrix(0, 6) = "Ч��"
        .TextMatrix(0, 7) = "�ⷿ����"
        .TextMatrix(0, 8) = "��λ"
        .TextMatrix(0, 9) = "�ϴβɹ���"
        .TextMatrix(0, 10) = "����ۼ�"
        .TextMatrix(0, 11) = "ϵ��"
        .TextMatrix(0, 12) = "��������"
        .TextMatrix(0, 13) = "�������"
        .TextMatrix(0, 14) = "�����"
        .TextMatrix(0, 15) = "�����"
        .TextMatrix(0, 16) = "����ʱ��"
        .TextMatrix(0, 17) = "����"
        .TextMatrix(0, 18) = "�ϴι�Ӧ��"
        If Not BlnSetHeader Then
            If mrsData.RecordCount = 0 Then Exit Sub
            Call DataBound
        End If
        
        For intCol = 0 To .Cols - 1
            .ColAlignmentFixed(intCol) = 4
        Next
        
        If BlnSetHeader Then
            If mblnFirst Then
                .ColWidth(0) = 0
                .ColWidth(1) = 0
                .ColWidth(2) = 1000
                .ColWidth(3) = 2000
                .ColWidth(4) = 900
                .ColWidth(5) = 1400
                .ColWidth(6) = 0
                .ColWidth(7) = 800
                .ColWidth(8) = 800
                If Me.cob�ⷿ.ItemData(Me.cob�ⷿ.ListIndex) = -1 Or Me.cob�ⷿ.ItemData(Me.cob�ⷿ.ListIndex) = 0 Then
                    .ColWidth(9) = 0
                Else
                    .ColWidth(9) = IIf(mblnCostView = False, 0, 1000)
                End If
                .ColWidth(10) = 1000
                .ColWidth(11) = 0
                .ColWidth(12) = 1000
                .ColWidth(13) = 1000
                .ColWidth(14) = 1000
                .ColWidth(15) = IIf(mblnCostView = False, 0, 1000)
                .ColWidth(16) = 0
                .ColWidth(17) = 0
                .ColWidth(18) = 1500
            End If
        Else
            .ColWidth(0) = 0
            .ColWidth(1) = 0
            .ColWidth(6) = 0
            If Me.cob�ⷿ.ItemData(Me.cob�ⷿ.ListIndex) = -1 Or Me.cob�ⷿ.ItemData(Me.cob�ⷿ.ListIndex) = 0 Then
                .ColWidth(9) = 0
            Else
                .ColWidth(9) = IIf(mblnCostView = False, 0, 1000)
            End If
            .ColWidth(11) = 0
            .ColWidth(15) = IIf(mblnCostView = False, 0, 1000)
            .ColWidth(16) = 0
            .ColWidth(17) = 0
            
            .ColAlignment(2) = 1
            .ColAlignment(3) = 1
            .ColAlignment(4) = 1
            .ColAlignment(9) = 7
            .ColAlignment(10) = 7
            .ColAlignment(11) = 7
            .ColAlignment(12) = 7
            .ColAlignment(13) = 7
            .ColAlignment(14) = 7
            .ColAlignment(15) = 7
        End If
        .Row = 1
    End With
End Sub

Private Sub DataBound()
    Dim lngColor As Long
    Dim lngRow As Long, lngCol As Long
    
    If mrsData.RecordCount <> 0 Then mrsData.MoveFirst
    With Msf������Ϣ_S
        .Redraw = False
        mblnColor = True
        Do While Not mrsData.EOF
            If mrsData.AbsolutePosition > .Rows - 1 Then .Rows = .Rows + 1
            .Row = mrsData.AbsolutePosition
            '�������
            .TextMatrix(.Row, 0) = mrsData!����ID
            .TextMatrix(.Row, 1) = mrsData!����id
            .TextMatrix(.Row, 2) = mrsData!����
            .TextMatrix(.Row, 3) = mrsData!����
            .TextMatrix(.Row, 4) = zlStr.NVL(mrsData!���, "")
            .TextMatrix(.Row, 5) = zlStr.NVL(mrsData!����, "")
            .TextMatrix(.Row, 6) = zlStr.NVL(mrsData!Ч��, "")
            .TextMatrix(.Row, 7) = zlStr.NVL(mrsData!�ⷿ����, "��")
            .TextMatrix(.Row, 8) = zlStr.NVL(mrsData!��λ, "")
            .TextMatrix(.Row, 9) = Format(mrsData!�ϴβɹ���, mFMT.FM_�ɱ���)
            .TextMatrix(.Row, 10) = Format(mrsData!����ۼ�, mFMT.FM_���ۼ�)
            .TextMatrix(.Row, 11) = Format(mrsData!ϵ��, GFM_VBXS)
            .TextMatrix(.Row, 12) = Format(mrsData!��������, mFMT.FM_����)
            .TextMatrix(.Row, 13) = Format(mrsData!ʵ������, mFMT.FM_����)
            .TextMatrix(.Row, 14) = Format(mrsData!ʵ�ʽ��, mFMT.FM_���)
            .TextMatrix(.Row, 15) = Format(mrsData!ʵ�ʲ��, mFMT.FM_���)
            .TextMatrix(.Row, 16) = zlStr.NVL(mrsData!����ʱ��, "")
            .TextMatrix(.Row, 17) = zlStr.NVL(mrsData!����, 1)
            .TextMatrix(.Row, 18) = zlStr.NVL(mrsData!��Ӧ��, "")
            '��ɫ
            If mbln����ͣ�� Then
                lngColor = IIf(Trim(.TextMatrix(.Row, 16)) = "", MLNG��ɫ, MLNG��ɫ)
                For lngCol = 0 To .Cols - 1
                    .Col = lngCol
                    .CellForeColor = lngColor
                Next
            End If
            mrsData.MoveNext
        Loop
        .Redraw = True
        mblnColor = False
    End With
End Sub

Private Sub txt������Ϣ_GotFocus()
    Call zlControl.TxtSelAll(txt������Ϣ)
End Sub

Private Sub txt������Ϣ_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strFind As String
    Dim strTemp As String
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(txt������Ϣ.Text) = "" Then Exit Sub
    
    '0-����,1-����,2-����,3-���,4-����,5-ָ������
    '����:[1]�ⷿ,[2]-���� ,[3]-����,[4]-����,[5]-���,[6]-����,[7]-ָ������

    txt������Ϣ.Text = Replace(txt������Ϣ.Text, "'", "")
    strTemp = GetMatchingSting(txt������Ϣ.Text)
    mstrOthers(0) = strTemp
    mstrOthers(1) = strTemp
    mstrOthers(2) = strTemp
    mstrOthers(3) = strTemp
'    strFind = "(Q.���� like [3] "
'    strFind = strFind & " Or Q.���� like [2] "
'    strFind = strFind & " Or M.����id in (Select �շ�ϸĿID from �շ���Ŀ����  where ���� like [4] ))"
    
    strFind = " M.����id in (Select Distinct a.Id " & _
        " From �շ���ĿĿ¼ A, �շ���Ŀ���� B " & _
        " Where a.Id = b.�շ�ϸĿid And (a.���� Like [3] Or a.���� Like [2] Or ���� Like [4]) "
    
    If gblnCode = True Then
        strFind = strFind & " Union All " & _
        " Select ҩƷid From ҩƷ��� " & _
        " Where ���� = 1 And �ⷿid + 0 = [1] And (��Ʒ���� Like [2] Or �ڲ����� Like [2])) "
    Else
        strFind = strFind & ") "
    End If
    
    If Not ReFreshFilterData(cob�ⷿ.ItemData(cob�ⷿ.ListIndex), strFind) Then Exit Sub
    
    Me.tvwSection_S.Tag = "T"
End Sub



Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub



