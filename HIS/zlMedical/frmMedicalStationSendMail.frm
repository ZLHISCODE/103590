VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMedicalStationSendMail 
   Caption         =   "���ͱ����ʼ�"
   ClientHeight    =   5760
   ClientLeft      =   2775
   ClientTop       =   4050
   ClientWidth     =   11220
   Icon            =   "frmMedicalStationSendMail.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   11220
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   11220
      _ExtentX        =   19791
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   11220
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbrThis"
      MinHeight1      =   720
      Width1          =   8370
      Key1            =   "only"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   720
         Left            =   30
         TabIndex        =   25
         Top             =   30
         Width           =   11100
         _ExtentX        =   19579
         _ExtentY        =   1270
         ButtonWidth     =   1429
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ilsMenu"
         HotImageList    =   "ilsHotMenu"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   8
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&S.����"
               Key             =   "����"
               Object.ToolTipText     =   "����(Alt+S)"
               Object.Tag             =   "&S.����"
               ImageKey        =   "SendMail"
               Style           =   5
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&O.���"
               Key             =   "���"
               Object.ToolTipText     =   "���ΪHtml��ʽ�ļ�"
               Object.Tag             =   "&O.���"
               ImageKey        =   "Html"
               Style           =   5
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&A.ȫѡ"
               Key             =   "ȫѡ"
               Object.ToolTipText     =   "ȫѡ(Alt+A)"
               Object.Tag             =   "&A.ȫѡ"
               ImageKey        =   "SelectAll"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&C.ȫ��"
               Key             =   "ȫ��"
               Object.ToolTipText     =   "ȫ��(Alt+C)"
               Object.Tag             =   "&C.ȫ��"
               ImageKey        =   "ClearAll"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&H.����"
               Key             =   "����"
               Object.ToolTipText     =   "����(Alt+H)"
               Object.Tag             =   "&H.����"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&X.�˳�"
               Key             =   "�˳�"
               Object.ToolTipText     =   "�˳�(Alt+X)"
               Object.Tag             =   "&X.�˳�"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fra2 
      Height          =   3825
      Left            =   3315
      TabIndex        =   14
      Top             =   810
      Width           =   7875
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1260
         Left            =   1065
         TabIndex        =   20
         Top             =   2475
         Width           =   6735
         _cx             =   11880
         _cy             =   2222
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
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
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
         Begin VB.Line lnX 
            Index           =   0
            Visible         =   0   'False
            X1              =   -4635
            X2              =   -2850
            Y1              =   -1695
            Y2              =   -1695
         End
         Begin VB.Line lnY 
            Index           =   0
            Visible         =   0   'False
            X1              =   270
            X2              =   270
            Y1              =   420
            Y2              =   1635
         End
      End
      Begin VB.TextBox txt 
         Height          =   1890
         Index           =   7
         Left            =   1080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Top             =   525
         Width           =   6705
      End
      Begin VB.TextBox txt 
         ForeColor       =   &H80000006&
         Height          =   300
         Index           =   8
         Left            =   1065
         TabIndex        =   16
         Top             =   165
         Width           =   3750
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&M.�ʼ�����"
         Height          =   180
         Index           =   8
         Left            =   60
         TabIndex        =   17
         Top             =   555
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&P.�����Ա"
         Height          =   180
         Index           =   9
         Left            =   105
         TabIndex        =   19
         Top             =   2505
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&E.�����ʼ�"
         Height          =   180
         Index           =   10
         Left            =   60
         TabIndex        =   15
         Top             =   225
         Width           =   900
      End
   End
   Begin VB.Frame fraInfo 
      Height          =   630
      Left            =   3105
      TabIndex        =   30
      Top             =   4620
      Width           =   3165
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   9
         Left            =   1980
         TabIndex        =   22
         Top             =   225
         Width           =   1140
      End
      Begin VB.CommandButton cmdMenu 
         Height          =   270
         Left            =   675
         Picture         =   "frmMedicalStationSendMail.frx":076A
         Style           =   1  'Graphical
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   240
         Width           =   285
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Index           =   7
         Left            =   180
         TabIndex        =   32
         Tag             =   "����"
         Top             =   285
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&6.����"
         Height          =   180
         Index           =   11
         Left            =   1020
         TabIndex        =   21
         Tag             =   "����"
         Top             =   285
         Width           =   540
      End
   End
   Begin VB.TextBox txtInfo 
      Height          =   2295
      Left            =   11520
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   29
      TabStop         =   0   'False
      Text            =   "frmMedicalStationSendMail.frx":09F0
      Top             =   1605
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSWinsockLib.Winsock sckMail 
      Left            =   3960
      Top             =   5040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtHead 
      Height          =   2295
      Left            =   11790
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   26
      TabStop         =   0   'False
      Text            =   "frmMedicalStationSendMail.frx":11E8
      Top             =   2895
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   23
      Top             =   5400
      Width           =   11220
      _ExtentX        =   19791
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMedicalStationSendMail.frx":1964
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14711
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
   Begin MSComctlLib.ImageList ilsMenu 
      Left            =   7950
      Top             =   1050
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationSendMail.frx":21F8
            Key             =   "SelectAll"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationSendMail.frx":2972
            Key             =   "ClearAll"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationSendMail.frx":30EC
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationSendMail.frx":3306
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationSendMail.frx":3526
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationSendMail.frx":3746
            Key             =   "PrintSet"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationSendMail.frx":3960
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationSendMail.frx":3B7A
            Key             =   "SendMail"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationSendMail.frx":42F4
            Key             =   "Html"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsHotMenu 
      Left            =   8625
      Top             =   1050
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationSendMail.frx":450E
            Key             =   "SelectAll"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationSendMail.frx":4C88
            Key             =   "ClearAll"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationSendMail.frx":5402
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationSendMail.frx":561C
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationSendMail.frx":583C
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationSendMail.frx":5A5C
            Key             =   "PrintSet"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationSendMail.frx":5C76
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationSendMail.frx":5E90
            Key             =   "SendMail"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationSendMail.frx":660A
            Key             =   "Html"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fra 
      Height          =   4425
      Left            =   45
      TabIndex        =   0
      Top             =   870
      Width           =   2820
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   6
         Left            =   2145
         TabIndex        =   13
         Text            =   "30"
         Top             =   3915
         Width           =   525
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   5
         Left            =   2145
         TabIndex        =   11
         Text            =   "5"
         Top             =   3540
         Width           =   525
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   0
         Left            =   120
         TabIndex        =   27
         Text            =   "25"
         Top             =   1050
         Width           =   2580
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   4
         Left            =   120
         TabIndex        =   2
         Top             =   435
         Width           =   2580
      End
      Begin VB.CheckBox chk 
         Alignment       =   1  'Right Justify
         Caption         =   "&6.���淢��������"
         Height          =   255
         Left            =   90
         TabIndex        =   9
         Top             =   3240
         Width           =   1845
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   2865
         Width           =   2580
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   2235
         Width           =   2580
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   1635
         Width           =   2580
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&8.�ȴ�����Ӧ����(��)"
         Height          =   180
         Index           =   6
         Left            =   120
         TabIndex        =   12
         Top             =   3960
         Width           =   1980
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&7.���������ʼ����(��)"
         Height          =   180
         Index           =   5
         Left            =   120
         TabIndex        =   10
         Top             =   3600
         Width           =   1980
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&2.�˿ں�"
         Height          =   180
         Index           =   0
         Left            =   105
         TabIndex        =   28
         Top             =   795
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&1.�ʼ�������"
         Height          =   180
         Index           =   4
         Left            =   90
         TabIndex        =   1
         Top             =   195
         Width           =   1080
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&5.��  ��"
         Height          =   180
         Index           =   3
         Left            =   105
         TabIndex        =   7
         Top             =   2625
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&4.�û���"
         Height          =   180
         Index           =   2
         Left            =   105
         TabIndex        =   5
         Top             =   2010
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&3.�����˵�ַ"
         Height          =   180
         Index           =   1
         Left            =   105
         TabIndex        =   3
         Top             =   1425
         Width           =   1080
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFileMail 
         Caption         =   "���͸��˱���(&M)"
      End
      Begin VB.Menu mnuFileMailGroup 
         Caption         =   "�������屨��(&E)"
      End
      Begin VB.Menu mnuFileOut 
         Caption         =   "������˱���(&O)"
      End
      Begin VB.Menu mnuFileOutGroup 
         Caption         =   "������屨��(&U)"
      End
      Begin VB.Menu mnuFile_0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSelectAll 
         Caption         =   "ȫѡ(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuFileClearAll 
         Caption         =   "ȫ��(&C)"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "�˳�(&X)"
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
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuHelpTopic 
         Caption         =   "��������(&T)"
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
      Begin VB.Menu mnuHelp_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)..."
      End
   End
End
Attribute VB_Name = "frmMedicalStationSendMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'���������弶��������**************************************************************************************************
Private mblnStartUp As Boolean                          '����������־
Private mblnOK As Boolean
Private mfrmMain As Object
Private mlngKey As Long
Private mblnChanged As Boolean
Private mblnMaining As Boolean
Private mlng����id As Long
Private mblnDataMoved As Boolean
Private Enum mCol
    ѡ�� = 0
    ����
    �����
    �Ա�
    ��������
    ����״��
    �����ʼ�
    ״̬
End Enum

Public WithEvents mobjPopMenu As clsPopMenu                '�Զ��嵯���˵�����
Attribute mobjPopMenu.VB_VarHelpID = -1
Public mbytPopMenu As Byte

'�������Զ�����̻���************************************************************************************************
Private Property Let EditChanged(ByVal vData As Boolean)
    '------------------------------------------------------------------------------------------------------------------
    '����:
    'ֵ��:
    '------------------------------------------------------------------------------------------------------------------
    
    mnuFileMail.Enabled = True
        
    If vData = False Then mnuFileMail.Enabled = False
        
    tbrThis.Buttons("����").Enabled = mnuFileMail.Enabled
    
End Property

Private Function CreateTmpFile(Optional ByVal strFile As String) As String
    '------------------------------------------------------------------------------------------------------------------
    '
    '����:
    '
    '------------------------------------------------------------------------------------------------------------------
    Dim strFileTemp As String
    Dim lngTemp As Long
    
    strFileTemp = Space(256)
    lngTemp = GetTempPath(256, strFileTemp)
    
    strFileTemp = Mid(strFileTemp, 1, InStr(strFileTemp, Chr(0)) - 1)
    
    strFileTemp = strFileTemp & strFile
    
    CreateTmpFile = strFileTemp
    
End Function

Private Function ClearData(Optional ByVal strMenuItem As String = "") As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long

    On Error Resume Next



    On Error GoTo 0

    Call InitData

    EditChanged = True


End Function

Public Function ShowEdit(ByVal frmMain As Object, ByRef lngKey As Long, Optional lng����id As Long = 0) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��ʾ�༭���壬������ô���Ľӿں���
    '����:  frmMain         ���ô������
    '       lngKey          ԤԼ�Ǽ�id
    '����:  True
    '       False
    '------------------------------------------------------------------------------------------------------------------
    mblnStartUp = True
    mblnOK = False

    mlngKey = lngKey
        
    Set mfrmMain = frmMain
    mlng����id = lng����id
    
    If InitData = False Then Exit Function
    If ReadData(mlngKey, lng����id) = False Then Exit Function
    
    EditChanged = (Val(vsf.RowData(1)) > 0)

    Me.Show 1, frmMain
    
    ShowEdit = mblnOK

End Function

Private Function ReadData(ByVal lngKey As Long, ByVal lng����id As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��ȡ����
    '����:  lngKey      ����������
    '����:  True        ��ȡ�ɹ�
    '       False       ��ȡʧ��
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset

    On Error GoTo errHand
    
    gstrSQL = "SELECT 1 AS ѡ��,A.����id AS ID,A.����,B.�����,b.������,b.���￨��,a.�����,b.���֤��,B.�Ա�,B.����״��,TO_CHAR(B.��������,'yyyy-mm-dd') AS ��������,A.�����ʼ�,'' as ״̬ " & _
                "FROM �����Ա���� A,������Ϣ B " & _
                "WHERE A.��챨��=1 AND A.���״̬ IN (4,5) AND A.����id=B.����id and A.�Ǽ�id=[1] "
    If lng����id > 0 Then gstrSQL = gstrSQL & " AND B.����id=[2] "
    
    gstrSQL = gstrSQL & " Order By B.�����"
    
    mblnDataMoved = DataMove(lngKey)
    If mblnDataMoved Then
        gstrSQL = Replace(gstrSQL, "�����Ա����", "H�����Ա����")
    End If
    
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey, lng����id)
    If rs.BOF = False Then
        Call FillGrid(vsf, rs)
        Call AppendRows(vsf, lnX, lnY)
    End If
    
    gstrSQL = "SELECT A.�����ʼ� FROM ��Լ��λ A,���ǼǼ�¼ B WHERE A.ID=B.��Լ��λid AND B.ID=[1]"
    
    If mblnDataMoved Then
        gstrSQL = Replace(gstrSQL, "���ǼǼ�¼", "H���ǼǼ�¼")
    End If
    
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
    If rs.BOF = False Then txt(8).Text = zlCommFun.NVL(rs("�����ʼ�"))
        
    ReadData = True

    Exit Function

errHand:
    If ErrCenter = 1 Then Resume

End Function

Private Function InitData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��ʼ������
    '����:  True        ��ʼ���ɹ�
    '       False       ��ʼ��ʧ��
    '------------------------------------------------------------------------------------------------------------------
    Dim strVsf As String
    
    On Error GoTo errHand
    
    strVsf = "ѡ��,450,1,1,1,;����,1080,1,1,1,;�����,810,7,1,1,;������,810,7,1,1,;���￨��,0,1,1,1,;�����,990,1,1,1,;���֤��,1200,1,1,0,;�Ա�,600,1,1,1,;��������,990,1,1,1,;����״��,900,1,1,1,;�����ʼ�,1800,1,1,1,;״̬,750,1,1,1,"
    
    Call CreateVsf(vsf, strVsf)
    vsf.Cols = vsf.Cols + 1
    vsf.ColWidth(vsf.Cols - 1) = 15
    vsf.ColDataType(0) = flexDTBoolean
    vsf.Editable = True
    
    Call AppendRows(vsf, lnX, lnY)
    
    If mlng����id > 0 Then
        
        mnuFileMailGroup.Visible = False
        mnuFileOutGroup.Visible = False
        
        txt(8).Visible = False
        lbl(10).Visible = False
    End If
    
    InitData = True

    Exit Function

errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function ValidEdit() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  У�����ݵ���Ч��
    '����:  True        ������Ч
    '       False       ������Ч
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long


    ValidEdit = True

End Function


Private Sub chk_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub


Private Sub cmdMenu_Click()
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(cmdMenu.hWnd, objPoint)
    
    mbytPopMenu = 3
    Set mobjPopMenu = New clsPopMenu
    Call mobjPopMenu.ShowPopupMenu(objPoint.X * Screen.TwipsPerPixelX, objPoint.Y * Screen.TwipsPerPixelY - 255 * 8 - 300)
    
    txt(9).Text = ""
    LocationObj txt(9)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 4 Then
        Select Case KeyCode
        Case vbKeyA
            If tbrThis.Buttons("ȫѡ").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("ȫѡ"))
        Case vbKeyC
            If tbrThis.Buttons("ȫ��").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("ȫ��"))
        Case vbKeyS
            If tbrThis.Buttons("����").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("����"))
            
        Case vbKeyO
            If tbrThis.Buttons("���").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("���"))
            
        Case vbKeyH
            If tbrThis.Buttons("����").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("����"))
        Case vbKeyX
            If tbrThis.Buttons("�˳�").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("�˳�"))
        End Select
    ElseIf Shift = 0 Then
        If KeyCode = vbKeyEscape Then
            If tbrThis.Buttons("�˳�").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("�˳�"))
        End If
    End If
End Sub

'���������弰��ؼ����¼�����******************************************************************************************
Private Sub Form_Load()
    
    txt(0).Text = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "������", txt(0).Text)
    txt(1).Text = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "�����˵�ַ", txt(1).Text)
    txt(2).Text = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "�û���", txt(2).Text)
    txt(3).Text = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "����", txt(3).Text)
    
    txt(4).Text = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "�ʼ�������", txt(4).Text)
    
    txt(5).Text = Val(GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "�������", txt(5).Text))
    txt(6).Text = Val(GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "�ȴ����", txt(6).Text))
    
    chk.Value = Val(GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "�Ƿ񱣴�����", chk.Value))
    txt(7).Text = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "�ʼ�����", txt(7).Text)
    
    Call RestoreWinState(Me, App.ProductName)
    
    
    If Val(GetSetting("ZLSOFT", "˽��ȫ��\" & gstrDBUser, "ʹ�ø��Ի����", "0")) = 1 Then
        'ʹ�ø��Ի�����
      
        lbl(11).Caption = "&6." & (GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "������Ϣ", "����"))
        lbl(11).Tag = Mid(lbl(11).Caption, 4)
    End If
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next

    With fra
        .Left = 0
        .Top = IIf(cbrThis.Visible, cbrThis.Height, 0) - 90
        .Height = Me.ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
    End With
    
    With fra2
        .Left = fra.Left + fra.Width + 15
        .Top = fra.Top
        .Width = Me.ScaleWidth - .Left
        .Height = fra.Height - fraInfo.Height + 90
    End With
    
    With fraInfo
        .Left = fra2.Left
        .Top = fra2.Top + fra2.Height - 90
        .Width = fra2.Width
    End With
    
    txt(8).Width = fra2.Width - txt(8).Left - 60
    With txt(7)
        .Width = fra2.Width - txt(7).Left - 60
    End With
    
    If mlng����id > 0 Then
        txt(7).Top = txt(8).Top
        lbl(8).Top = lbl(10).Top
        vsf.Top = txt(7).Top + txt(7).Height + 30
    End If
    
    lbl(9).Top = vsf.Top + 60
    With vsf
        .Width = fra2.Width - .Left - 60
        .Height = fra2.Height - .Top - 60
    End With
    
    Call AppendRows(vsf, lnX, lnY)
End Sub

Private Sub Form_Unload(Cancel As Integer)
        
    If mblnMaining Then
        Cancel = True
        Exit Sub
    End If
    
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "������", txt(0).Text)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "�����˵�ַ", txt(1).Text)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "�û���", txt(2).Text)
    
    If chk.Value = 1 Then
        Call SaveSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "����", txt(3).Text)
    Else
        Call SaveSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "����", "")
    End If
    
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "�ʼ�������", txt(4).Text)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "�������", Val(txt(5).Text))
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "�ȴ����", Val(txt(6).Text))
    
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "�Ƿ񱣴�����", chk.Value)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "�ʼ�����", txt(7).Text)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "������Ϣ", lbl(11).Tag)

    Call SaveWinState(Me, App.ProductName)
    
End Sub

Private Sub mnuFileClearAll_Click()
    Dim lngLoop As Long
    
    For lngLoop = 1 To vsf.Rows - 1
        If Val(vsf.RowData(lngLoop)) > 0 Then
            vsf.TextMatrix(lngLoop, mCol.ѡ��) = 0
        End If
    Next
    
    EditChanged = False
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileMail_Click()
    Dim objMail As clsMail
    Dim blnSuccess As Boolean
    Dim strMessage As String
    Dim lngLoop As Long
    
    Dim strFile As String
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    
    '���
    If ValidData = False Then Exit Sub
    
    Set objMail = New clsMail
    Set objMail.WinSockObj = sckMail
    
    mblnMaining = True
    
    tbrThis.Buttons("����").Enabled = False
    tbrThis.Buttons("ȫ��").Enabled = False
    tbrThis.Buttons("ȫѡ").Enabled = False
    tbrThis.Buttons("����").Enabled = False
    tbrThis.Buttons("�˳�").Enabled = False
    
    vsf.Editable = flexEDNone
    mnuFile.Enabled = False
    mnuView.Enabled = False
    mnuHelp.Enabled = False
    
    vsf.Cell(flexcpText, 1, mCol.״̬, vsf.Rows - 1, mCol.״̬) = ""
    vsf.Cell(flexcpForeColor, 1, mCol.״̬, vsf.Rows - 1, mCol.״̬) = COLOR.��ɫ
    
    frmWait.OpenWait Me, "���͵����ʼ�"
    frmWait.WaitInfo = "���������ʼ�������..."
    
    objMail.ResponseInternal = Val(txt(6).Text)
    
    If objMail.OpenMailServer(txt(4).Text, txt(2).Text, txt(3).Text, Val(txt(0).Text)) Then
'    If objMail.OpenOutLookExMail() Then
        
        For lngLoop = 1 To vsf.Rows - 1
            If Val(vsf.RowData(lngLoop)) > 0 And Abs(Val(vsf.TextMatrix(lngLoop, mCol.ѡ��))) = 1 And Trim(vsf.TextMatrix(lngLoop, mCol.�����ʼ�)) <> "" Then
                
                frmWait.WaitInfo = "���ڷ��͡�" & vsf.TextMatrix(lngLoop, mCol.����) & "������챨���ʼ�..."
                
                txtInfo.Text = txtHead.Text
                Call GetReportMessageHtml(mlngKey, Val(vsf.RowData(lngLoop)))
           
'                strFile = CreateTmpFile("��챨��.htm")
'                Set objText = objFile.CreateTextFile(strFile, True)
'                objText.Write txtInfo.Text
'                objText.Close

                blnSuccess = objMail.SendHead(vsf.TextMatrix(lngLoop, mCol.�����ʼ�), txt(2).Text, txt(1).Text, "������챨��", vbMultipartAlternative)
                blnSuccess = objMail.SendMessage(txtInfo.Text, vbTextHtml)
                blnSuccess = objMail.SendOver
'                blnSuccess = objMail.SendOutLookExMail(vsf.TextMatrix(lngLoop, mCol.�����ʼ�), "������챨��", txt(7).Text, strFile)
                
                If blnSuccess Then
                    vsf.TextMatrix(lngLoop, mCol.״̬) = "�ѷ���"
                Else
                    vsf.TextMatrix(lngLoop, mCol.״̬) = "ʧ  ��"
                    vsf.Cell(flexcpForeColor, lngLoop, mCol.״̬) = COLOR.��ɫ
                End If
                
                Sleep Val(txt(5).Text) * 1000
                
            End If
        Next
    End If
    
    frmWait.WaitInfo = "���ڹر��ʼ�������..."
    
    Call objMail.CloseMailServer
'    Call objMail.CloseOutLookExMail
    
    tbrThis.Buttons("����").Enabled = True
    tbrThis.Buttons("ȫ��").Enabled = True
    tbrThis.Buttons("ȫѡ").Enabled = True
    tbrThis.Buttons("����").Enabled = True
    tbrThis.Buttons("�˳�").Enabled = True
    
    vsf.Editable = flexEDKbdMouse
    mnuFile.Enabled = True
    mnuView.Enabled = True
    mnuHelp.Enabled = True
    mblnMaining = False
    
    frmWait.CloseWait
    
End Sub

Private Function GetReportMessageHtml(ByVal lngKey As Long, ByVal lng����id As Long) As String
    '------------------------------------------------------------------------------------------------------------------
    '
    '����:������Ա��챨��Html��ʽ,�����ʼ����ͣ�ע��˸�ʽ�ǹ̶���
    '
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim rs3 As New ADODB.Recordset
    Dim lngLoop1 As Long
    Dim lngLoop2 As Long
    Dim lngLoop3 As Long
    Dim strTmp1 As String
    Dim strTmp2 As String
    
    Dim strSQL As String
    
    On Error GoTo errHand
    
    strTmp1 = txt(7).Text
    strTmp1 = ReplaceAll(strTmp1, vbCrLf, "<br>")
    
    txtInfo.Text = txtInfo.Text & vbCrLf & _
        "<BODY BGCOLOR=#FFFFFF>" & vbCrLf & _
        "<table x:str border=0 cellpadding=5 cellspacing=0 width=728 style='border-collapse:collapse;table-layout:fixed;width:548pt'>" & vbCrLf & _
        "<col style='mso-width-source:userset;mso-width-alt:512;width:150pt'>" & vbCrLf & _
        "<col style='mso-width-source:userset;mso-width-alt:512;width:150pt'>" & vbCrLf & _
        "<col style='mso-width-source:userset;mso-width-alt:512;width:120pt'>" & vbCrLf & _
        "<col style='mso-width-source:userset;mso-width-alt:512;width:40pt'>"

    txtInfo.Text = txtInfo.Text & vbCrLf & _
            "<tr><td colspan=4 class=xl39 style='font-weight:300'>" & strTmp1 & "<br></td></tr>"
    
    txtInfo.Text = txtInfo.Text & vbCrLf & _
        "<tr><td colspan=4 class=xlTitle style='width:536pt'>" & GetUnitName & "��챨�浥</td></tr>"
                        
    strTmp1 = ""
    
    strSQL = "SELECT A.����,C.�����,B.���ʱ��,C.����,B.��첡��id,B.����ʱ��,D.��д��,C.�����,E.���� " & _
                "FROM ���ǼǼ�¼ A,�����Ա���� B,������Ϣ C,���˲�����¼ D,��Լ��λ E " & _
                "WHERE A.��Լ��λID=E.ID(+) AND D.ID(+)=B.��첡��id AND C.����id=B.����id AND A.ID=B.�Ǽ�id AND A.ID=[1] AND B.����id=[2] "
    
    If mblnDataMoved Then
        strSQL = Replace(strSQL, "���ǼǼ�¼", "H���ǼǼ�¼")
        strSQL = Replace(strSQL, "�����Ա����", "H�����Ա����")
        strSQL = Replace(strSQL, "���˲�����¼", "H���˲�����¼")
    End If
    
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngKey, lng����id)
    
    If rs.BOF Then Exit Function
    
    txtInfo.Text = txtInfo.Text & _
        "<tr><td class=xl39 style='font-weight:700'>�ܼ쵥λ��<font class=" & Chr(34) & "font8" & Chr(34) & ">" & zlCommFun.NVL(rs("����")) & "</td></tr>" & _
        "<tr><td class=xl39 style='font-weight:700'>�ܼ���Ա��<font class=" & Chr(34) & "font8" & Chr(34) & ">" & zlCommFun.NVL(rs("����")) & "</td></tr>" & _
        "<tr><td class=xl39 style='font-weight:700'>������ڣ�<font class=" & Chr(34) & "font8" & Chr(34) & ">" & Format(zlCommFun.NVL(rs("���ʱ��")), "YYYY-MM-DD") & "</td></tr>" & _
        "<tr><td class=xl39 style='font-weight:700'>�� �� �ţ�<font class=" & Chr(34) & "font8" & Chr(34) & ">" & zlCommFun.NVL(rs("�����")) & "</td></tr>"
        
        
    '�ܼ�
    '------------------------------------------------------------------------------------------------------------------
    strTmp1 = ""
    strTmp2 = ""
    
    strSQL = "SELECT * FROM �����Ա���� WHERE ����id in (select id from ���˲������� where ������¼id=[1]) ORDER BY ��¼����,��¼���"
    If mblnDataMoved Then
        strSQL = Replace(strSQL, "�����Ա����", "H�����Ա����")
    End If
    Set rs3 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(zlCommFun.NVL(rs("��첡��id"))))
    If rs3.BOF = False Then
        For lngLoop3 = 1 To rs3.RecordCount
            
            If zlCommFun.NVL(rs3("��¼����"), 0) = 0 Then strTmp1 = strTmp1 & zlCommFun.NVL(rs3("��������")) & vbCrLf
            If zlCommFun.NVL(rs3("��¼����"), 0) = 1 Then strTmp2 = zlCommFun.NVL(rs3("�ο�����"))
            
            rs3.MoveNext
        Next
    End If
            
    txtInfo.Text = txtInfo.Text & vbCrLf & _
        "<tr><td colspan=2 class=xl39 style='font-weight:700'>һ���ܼ챨��</td>" & vbCrLf & _
        "<td colspan=2 class=xl39 style='text-align:right'>�ܼ�ҽ����<font class=" & Chr(34) & "font8" & Chr(34) & ">" & zlCommFun.NVL(rs("��д��")) & "</td></tr>"
    
    strTmp1 = ReplaceAll(strTmp1, vbCrLf, "<br>")
    strTmp2 = ReplaceAll(strTmp2, vbCrLf, "<br>")
    
    txtInfo.Text = txtInfo.Text & vbCrLf & _
        "<tr><td colspan=4 class=xl25 style='text-align:left'>���ۣ�<font class=" & Chr(34) & "font8" & Chr(34) & ">" & strTmp1 & "</td></tr>" & vbCrLf & _
        "<tr><td colspan=4 class=xl25 style='text-align:left'>���飺<font class=" & Chr(34) & "font8" & Chr(34) & ">" & strTmp2 & "</td></tr>" & vbCrLf & _
        "<tr><td colspan=4 class=xl25 style='text-align:left'>���飺<font class=" & Chr(34) & "font8" & Chr(34) & ">" & Format(zlCommFun.NVL(rs("����ʱ��")), "yyyy-MM-dd") & "</td></tr>"
                    
    '�����Ŀ����
    '------------------------------------------------------------------------------------------------------------------
    txtInfo.Text = txtInfo.Text & _
        "<tr><td colspan=4 class=xl39 style='font-weight:700'>������Ŀ����</td></tr>"

    '1.����
    strSQL = _
        "Select c.����, c.Id" & vbNewLine & _
        "From ���ű� c," & vbNewLine & _
        "        (Select b.ִ�п���id, Max(Nvl(s.����˳��, 0)) As ����˳��" & vbNewLine & _
        "            From �����Ŀҽ�� a, �����Ŀ�嵥 b, �����Ŀ���� s" & vbNewLine & _
        "            Where b.�Ǽ�id = [1] And a.����id = [2] And a.�嵥id = b.Id And s.������Ŀid(+) = b.������Ŀid And s.��������(+) = 1" & vbNewLine & _
        "            Group By b.ִ�п���id) b" & vbNewLine & _
        "Where c.Id = b.ִ�п���id" & vbNewLine & _
        "Order By Decode(b.����˳��, 0, 9999999)"
    
    If mblnDataMoved Then
        strSQL = Replace(strSQL, "�����Ŀҽ��", "H�����Ŀҽ��")
        strSQL = Replace(strSQL, "�����Ŀ�嵥", "H�����Ŀ�嵥")
    End If
    
    Set rs1 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngKey, lng����id)
    If rs1.BOF Then Exit Function
    
    For lngLoop1 = 1 To rs1.RecordCount
        
        '2.�����Ŀ(��д�˱����)
        strSQL = "select C.����,B.����id,D.��д�� " & _
                        "from ( " & _
                             "SELECT * FROM ����ҽ����¼ WHERE ����id=[2] AND �Һŵ�=[1] AND ִ�п���id=[3] AND ������Դ=4 AND ҽ��״̬<>4 AND �������='D' AND ���id IS NULL " & _
                             "Union All " & _
                             "SELECT * FROM ����ҽ����¼ WHERE ����id=[2] AND �Һŵ�=[1] AND ִ�п���id=[3] AND ������Դ=4 AND ҽ��״̬<>4 AND �������='C' AND ���id>0 " & _
                             ") A, " & _
                             "����ҽ������ B, " & _
                             "������ĿĿ¼ C, " & _
                             "���˲�����¼ D,�����Ŀ���� S " & _
                        "Where A.ID = B.ҽ��id " & _
                              "AND B.����id>0 " & _
                              "AND C.ID=A.������ĿID " & _
                              "AND D.ID=B.����id And s.������Ŀid(+)=a.������ĿID And s.��������(+)=1 " & _
                        "Order By Nvl(s.����˳��,9999999)"
        
        If mblnDataMoved Then
            strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
            strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
            strSQL = Replace(strSQL, "���˲�����¼", "H���˲�����¼")
        End If
    
        Set rs2 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CStr(zlCommFun.NVL(rs("����"))), lng����id, Val(zlCommFun.NVL(rs1("ID"))))
        If rs2.BOF = False Then
                
            txtInfo.Text = txtInfo.Text & "<tr><td colspan=4 class=xl39 style='font-weight:700'>�� " & zlCommFun.NVL(rs1("����")) & "</td></tr>"
            
            txtInfo.Text = txtInfo.Text & "<tr>"
            
            For lngLoop2 = 1 To rs2.RecordCount
                
                txtInfo.Text = txtInfo.Text & "<td colspan=2 class=xl39 style='font-weight:600'>�� " & zlCommFun.NVL(rs2("����")) & "</td>"
                txtInfo.Text = txtInfo.Text & "<td colspan=2 class=xl39 style='text-align:right'>���ҽ����<font class=" & Chr(34) & "font8" & Chr(34) & ">" & zlCommFun.NVL(rs2("��д��")) & "</td>"
                txtInfo.Text = txtInfo.Text & "</tr>"
                
                txtInfo.Text = txtInfo.Text & _
                            "<tr><td class=xl25>��Ŀ����</td>" & vbCrLf & _
                            "<td class=xl25>�����</td>" & vbCrLf & _
                            "<td class=xl25>�ο���Χ</td>" & vbCrLf & _
                            "<td class=xl25>��ʾ</td></tr>"
                
                '��������Ŀ�����
                strSQL = _
                    "SELECT * FROM ( " & _
                        "SELECT " & _
                               "��Ŀ, " & _
                               "����, " & _
                               "�ο�, " & _
                               "Decode(��־,Null,'', '����', '', '�쳣', '(+)', 'ƫ��', '��', 'ƫ��', '��',��־) As ��ʾ," & _
                               "�������, " & _
                               "Ԫ������� " & _
                        "FROM ( " & _
                        "SELECT " & _
                               "��Ŀ, " & _
                               "����, " & _
                               "DECODE(SIGN(INSTR(�ο�,'''')),1,SUBSTR(�ο�,1,INSTR(�ο�,'''')-1),'') AS ��־, " & _
                               "DECODE(SIGN(INSTR(�ο�,'''')),1,SUBSTR(�ο�,INSTR(�ο�,'''')+1,1000),'') AS �ο�, " & _
                               "�������, " & _
                               "Ԫ������� " & _
                        "FROM ( " & _
                        "SELECT " & _
                               "��Ŀ, " & _
                               "DECODE(SIGN(INSTR(����,'''')),1,SUBSTR(����,1,INSTR(����,'''')-1),����) AS ����, " & _
                               "DECODE(SIGN(INSTR(����,'''')),1,SUBSTR(����,INSTR(����,'''')+1,1000),'') AS �ο�, " & _
                               "�������, " & _
                               "Ԫ������� "
                strSQL = strSQL & _
                        "FROM ( " & _
                        "SELECT C.������ AS ��Ŀ,DECODE(A.��������,NULL,NULL,A.��������||' '||DECODE(C.��λ,NULL,'',C.��λ)) AS ����,B.�������,A.�ؼ��� AS Ԫ������� FROM ���˲��������� A,���˲������� B,����������Ŀ C " & _
                        "Where A.����ID = B.ID " & _
                              "AND B.������¼ID=[1] " & _
                              "AND C.ID=A.������ID " & _
                        "))) " & _
                        "Union All " & _
                        "SELECT B.�����ı� AS ��Ŀ,A.����,'' AS �ο�,'' AS ��ʾ,B.�������,0 AS Ԫ������� FROM ���˲����ı��� A,���˲������� B " & _
                        "Where A.����ID = B.ID " & _
                                "And B.������¼ID =[1] " & _
                              "AND Ԫ������ IN (0,-5) " & _
                        ") ORDER BY �������,Ԫ�������"
                        
                If mblnDataMoved Then
                    strSQL = Replace(strSQL, "���˲����ı���", "H���˲����ı���")
                    strSQL = Replace(strSQL, "���˲�������", "H���˲�������")
                    strSQL = Replace(strSQL, "���˲���������", "H���˲���������")
                End If
                Set rs3 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(zlCommFun.NVL(rs2("����id"))))
                If rs3.BOF = False Then
                    For lngLoop3 = 1 To rs3.RecordCount
                        txtInfo.Text = txtInfo.Text & vbCrLf & _
                                "<tr><td class=xl28>" & zlCommFun.NVL(rs3("��Ŀ")) & "</td>" & vbCrLf & _
                                "<td class=xl28><font class=" & Chr(34) & "font8" & Chr(34) & ">" & zlCommFun.NVL(rs3("����")) & "</td>" & vbCrLf & _
                                "<td class=xl28><font class=" & Chr(34) & "font8" & Chr(34) & ">" & zlCommFun.NVL(rs3("�ο�")) & "</td>" & vbCrLf & _
                                "<td class=xl28><font class=" & Chr(34) & "font8" & Chr(34) & ">" & zlCommFun.NVL(rs3("��ʾ")) & "</td></tr>"
                        rs3.MoveNext
                    Next
                Else
                    txtInfo.Text = txtInfo.Text & vbCrLf & _
                                "<tr><td class=xl28 style='mso-height-source:userset;height:15.0pt'></td>" & vbCrLf & _
                                "<td class=xl28><font class=" & Chr(34) & "font8" & Chr(34) & "></td>" & vbCrLf & _
                                "<td class=xl28><font class=" & Chr(34) & "font8" & Chr(34) & "></td>" & vbCrLf & _
                                "<td class=xl28><font class=" & Chr(34) & "font8" & Chr(34) & "></td></tr>"
                End If
                                        
                strTmp1 = ""
                strTmp2 = ""
                
                strSQL = "SELECT * FROM �����Ա���� WHERE ����id in (select id from ���˲������� where ������¼id=[1]) ORDER BY ��¼����,��¼���"
                
                If mblnDataMoved Then
                    strSQL = Replace(strSQL, "�����Ա����", "H�����Ա����")
                End If
                
                Set rs3 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(zlCommFun.NVL(rs2("����id"))))
                If rs3.BOF = False Then
                    For lngLoop3 = 1 To rs3.RecordCount
                        
                        If zlCommFun.NVL(rs3("��¼����"), 0) = 0 Then strTmp1 = strTmp1 & zlCommFun.NVL(rs3("��������")) & vbCrLf
                        If zlCommFun.NVL(rs3("��¼����"), 0) = 1 Then strTmp2 = zlCommFun.NVL(rs3("�ο�����"))
                        
                        rs3.MoveNext
                    Next
                End If
                
                txtInfo.Text = txtInfo.Text & vbCrLf & _
                    "<tr><td colspan=4 class=xl28 style='font-weight:600'>���ۣ�<font class=" & Chr(34) & "font8" & Chr(34) & ">" & strTmp1 & "</td></tr>" & vbCrLf & _
                    "<tr><td colspan=4 class=xl28 style='font-weight:600'>���飺<font class=" & Chr(34) & "font8" & Chr(34) & ">" & strTmp2 & "</td></tr>"
                    
                txtInfo.Text = txtInfo.Text & vbCrLf & "<tr><td class=xl39 style='mso-height-source:userset;height:15.0pt'></td></tr>"
                
                rs2.MoveNext
            Next
        End If
        
        rs1.MoveNext
    Next
                
    '���
    txtInfo.Text = txtInfo.Text & vbCrLf & "</tr></table></BODY></HTML>"
    
    GetReportMessageHtml = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function GetGroupReportMessageHtml(ByVal lngKey As Long) As String
    '------------------------------------------------------------------------------------------------------------------
    '
    '����:����������챨��Html��ʽ,�����ʼ�����
    '
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim intCount As Integer
    Dim strSQL As String
    Dim strTmp1 As String
    
    strTmp1 = txt(7).Text
    strTmp1 = ReplaceAll(strTmp1, vbCrLf, "<br>")
    
    txtInfo.Text = txtInfo.Text & vbCrLf & _
        "<BODY BGCOLOR=#FFFFFF>" & vbCrLf & _
        "<table x:str border=0 cellpadding=5 cellspacing=0 style='border-collapse:collapse;table-layout:fixed;width:400pt'>" & vbCrLf & _
        "<col style='mso-width-source:userset;mso-width-alt:512;width:25pt'>" & vbCrLf & _
        "<col style='mso-width-source:userset;mso-width-alt:512;width:25pt'>" & vbCrLf & _
        "<col style='mso-width-source:userset;mso-width-alt:512;width:25pt'>" & vbCrLf & _
        "<col style='mso-width-source:userset;mso-width-alt:512;width:25pt'>" & vbCrLf & _
        "<col style='mso-width-source:userset;mso-width-alt:512;width:25pt'>" & vbCrLf & _
        "<col style='mso-width-source:userset;mso-width-alt:512;width:25pt'>" & vbCrLf & _
        "<col style='mso-width-source:userset;mso-width-alt:512;width:25pt'>" & vbCrLf & _
        "<col style='mso-width-source:userset;mso-width-alt:512;width:25pt'>" & vbCrLf & _
        "<col style='mso-width-source:userset;mso-width-alt:512;width:25pt'>" & vbCrLf & _
        "<col style='mso-width-source:userset;mso-width-alt:512;width:25pt'>"
        
    txtInfo.Text = txtInfo.Text & vbCrLf & _
            "<tr><td colspan=10 class=xl39 style='font-weight:300'>" & strTmp1 & "<br></td></tr>"
    
    txtInfo.Text = txtInfo.Text & vbCrLf & _
        "<tr><td colspan=10 class=xlTitle>������챨�浥</td></tr>"
                        
    strTmp1 = ""
    
    strSQL = "SELECT A.����,A.���ʱ��,B.���� FROM ���ǼǼ�¼ A,��Լ��λ B WHERE B.ID=A.��Լ��λid AND A.ID=[1]"
    If mblnDataMoved Then
        strSQL = Replace(strSQL, "���ǼǼ�¼", "H���ǼǼ�¼")
    End If
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngKey)
    If rs.BOF Then Exit Function
    
    txtInfo.Text = txtInfo.Text & _
        "<tr><td class=xl39 colspan=10 style='font-weight:700'>������壺<font class=font8>" & zlCommFun.NVL(rs("����")) & "</td></tr>" & _
        "<tr><td class=xl39 colspan=10 style='font-weight:700'>������ڣ�<font class=font8>" & Format(zlCommFun.NVL(rs("���ʱ��")), "YYYY-MM-DD") & "</td></tr>" & _
        "<tr><td class=xl39 colspan=10 style='font-weight:700'>��쵥�ţ�<font class=font8>" & zlCommFun.NVL(rs("����")) & "</td></tr>"
        
    '1.�������
            
    strSQL = _
        "SELECT " & _
            "DECODE(��������,0,NULL,��������) AS ��������, " & _
            "DECODE(Ů������,0,NULL,Ů������) AS Ů������, " & _
            "DECODE(����,0,NULL,����) AS ����, " & _
            "DECODE(�Ѽ���������,0,NULL,�Ѽ���������) AS �Ѽ���������, " & _
            "DECODE(�Ѽ�Ů������,0,NULL,�Ѽ�Ů������) AS �Ѽ�Ů������, " & _
            "DECODE(�Ѽ�����,0,NULL,�Ѽ�����) AS �Ѽ�����, " & _
            "DECODE(δ����������,0,NULL,δ����������) AS δ����������, " & _
            "DECODE(δ��Ů������,0,NULL,δ��Ů������) AS δ��Ů������, " & _
            "DECODE(δ������, 0, Null, δ������) As δ������ " & _
        "From " & _
        "( " & _
        "SELECT A.��������, " & _
               "A.Ů������, " & _
               "nvl(A.��������,0)+nvl(A.Ů������,0) AS ����, " & _
               "A.�Ѽ���������, " & _
               "A.�Ѽ�Ů������, " & _
               "nvl(A.�Ѽ���������,0)+nvl(A.�Ѽ�Ů������,0) AS �Ѽ�����, " & _
               "nvl(A.��������,0)-nvl(A.�Ѽ���������,0) AS δ����������, " & _
               "nvl(A.Ů������,0)-nvl(A.�Ѽ�Ů������,0) AS δ��Ů������, " & _
               "(nvl(A.��������,0)-nvl(A.�Ѽ���������,0))+(nvl(A.Ů������,0)-nvl(A.�Ѽ�Ů������,0)) AS δ������ "
               
    strSQL = strSQL & _
        "From " & _
        "( " & _
        "select SUM(DECODE(sign(instr(B.�Ա�,'Ů')-0),1,0,1)) AS ��������, " & _
               "SUM(DECODE(sign(instr(B.�Ա�,'Ů')-0),1,1,0)) AS Ů������, " & _
               "SUM(DECODE(sign(0 - NVL(B.��첡��ID,0)),-1, DECODE(SIGN(instr(B.�Ա�,'Ů')-0),1,0,1),0)) AS �Ѽ���������, " & _
               "SUM(DECODE(sign(0 - NVL(B.��첡��ID,0)),-1, DECODE(SIGN(instr(B.�Ա�,'Ů')-0),1,1,0),0)) AS �Ѽ�Ů������ " & _
        "from ���ǼǼ�¼ A, " & _
             "�����Ա���� B " & _
        "Where A.ID = B.�Ǽ�ID " & _
              "AND A.ID=[1] " & _
        ") A " & _
        ")"
        
    If mblnDataMoved Then
        strSQL = Replace(strSQL, "���ǼǼ�¼", "H���ǼǼ�¼")
        strSQL = Replace(strSQL, "�����Ա����", "H�����Ա����")
    End If
    
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngKey)
    If rs.BOF Then Exit Function
    
    intCount = intCount + 1
    txtInfo.Text = txtInfo.Text & "<tr><td colspan=10 class=xl39 style='font-weight:700'>" & intCount & ".�������</td></tr>"
        
    txtInfo.Text = txtInfo.Text & _
        "<tr>" & _
        "<td rowspan=2 class=xl25></td>" & _
        "<td colspan=3 class=xl25>����</td>" & _
        "<td colspan=3 class=xl25>�Ѽ�����</td>" & _
        "<td colspan=3 class=xl25>δ������</td>" & _
        "</tr>" & _
        "<tr>" & _
        "<td class=xl25>����</td>" & _
        "<td class=xl25>Ů��</td>" & _
        "<td class=xl25>�ϼ�</td>" & _
        "<td class=xl25>����</td>" & _
        "<td class=xl25>Ů��</td>" & _
        "<td class=xl25>�ϼ�</td>" & _
        "<td class=xl25>����</td>" & _
        "<td class=xl25>Ů��</td>" & _
        "<td class=xl25>�ϼ�</td>" & _
        "</tr>"
    
    txtInfo.Text = txtInfo.Text & _
        "<tr>" & _
        "<td class=xl25>����</td>" & _
        "<td class=xl29><font class=font8>" & zlCommFun.NVL(rs("��������")) & "</td>" & _
        "<td class=xl29><font class=font8>" & zlCommFun.NVL(rs("Ů������")) & "</td>" & _
        "<td class=xl29><font class=font8>" & zlCommFun.NVL(rs("����")) & "</td>" & _
        "<td class=xl29><font class=font8>" & zlCommFun.NVL(rs("�Ѽ���������")) & "</td>" & _
        "<td class=xl29><font class=font8>" & zlCommFun.NVL(rs("�Ѽ�Ů������")) & "</td>" & _
        "<td class=xl29><font class=font8>" & zlCommFun.NVL(rs("�Ѽ�����")) & "</td>" & _
        "<td class=xl29><font class=font8>" & zlCommFun.NVL(rs("δ����������")) & "</td>" & _
        "<td class=xl29><font class=font8>" & zlCommFun.NVL(rs("δ��Ů������")) & "</td>" & _
        "<td class=xl29><font class=font8>" & zlCommFun.NVL(rs("δ������")) & "</td>" & _
        "</tr>"
    
    '2.�������
    strSQL = _
        "Select ��������,count(����id) As ����,100*Count(����id)/Decode(�Ѽ�������,Null,1,0,1,�Ѽ�������) As ���� From " & _
        "( " & _
        "Select Distinct ��������,����id From �����Ա���� " & _
        "Where ��¼���� = 0 " & _
              "And ����id in " & _
                  "( " & _
                   "Select A.ID From ���˲������� A,����Ԫ��Ŀ¼ B " & _
                   "Where A.Ԫ�ر���=B.���� AND upper(B.����)='ZL9CISCORE.USRMEDICALSUM' " & _
                         "And A.������¼id In " & _
                             "( " & _
                              "Select ��첡��id From �����Ա���� Where �Ǽ�id=[1] " & _
                             ") " & _
                  ") " & _
        ") A, " & _
        "(Select Count(1) As �Ѽ������� From �����Ա���� Where ��첡��id>0 And �Ǽ�id=[1]) B " & _
        "Group by ��������,�Ѽ�������"
        
    If mblnDataMoved Then
        strSQL = Replace(strSQL, "���˲�������", "H���˲�������")
        strSQL = Replace(strSQL, "�����Ա����", "H�����Ա����")
        strSQL = Replace(strSQL, "�����Ա����", "H�����Ա����")
    End If
    
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngKey)
    If rs.BOF Then Exit Function
    
    intCount = intCount + 1
    txtInfo.Text = txtInfo.Text & "<tr><td colspan=10 class=xl39 style='font-weight:700'><br>" & intCount & ".�������</td></tr>"
        
    txtInfo.Text = txtInfo.Text & _
        "<tr>" & _
        "<td colspan=6 class=xl25>��������</td>" & _
        "<td colspan=2 class=xl25>����</td>" & _
        "<td colspan=2 class=xl25>����</td>" & _
        "</tr>"
        
    Do While Not rs.EOF
        
        txtInfo.Text = txtInfo.Text & _
            "<tr>" & _
            "<td colspan=6 class=xl28><font class=font8>" & zlCommFun.NVL(rs("��������")) & "</td>" & _
            "<td colspan=2 class=xl29><font class=font8>" & zlCommFun.NVL(rs("����")) & "</td>" & _
            "<td colspan=2 class=xl29><font class=font8>" & Format(zlCommFun.NVL(rs("����")), "0.00") & "%</td>" & _
            "</tr>"
            
        rs.MoveNext
    Loop
    
    strSQL = _
        "Select Count(����id) As ����,100*Count(����id)/Decode(�Ѽ�������,Null,1,0,1,�Ѽ�������) As ���� From " & _
        "( " & _
        "Select Distinct ��������,����id From �����Ա���� " & _
        "Where ��¼���� = 0 " & _
              "And ����id in " & _
                  "( " & _
                   "Select A.ID From ���˲������� A,����Ԫ��Ŀ¼ B " & _
                   "Where A.Ԫ�ر���=B.���� AND upper(B.����)='ZL9CISCORE.USRMEDICALSUM' " & _
                         "And A.������¼id In " & _
                             "( " & _
                              "Select ��첡��id From �����Ա���� Where �Ǽ�id=[1] " & _
                             ") " & _
                  ") " & _
        ") A, " & _
        "(Select Count(1) As �Ѽ������� From �����Ա���� Where ��첡��id>0 And �Ǽ�id=[1]) B " & _
        "Group by �Ѽ�������"
        
    If mblnDataMoved Then
        strSQL = Replace(strSQL, "���˲�������", "H���˲�������")
        strSQL = Replace(strSQL, "�����Ա����", "H�����Ա����")
        strSQL = Replace(strSQL, "�����Ա����", "H�����Ա����")
    End If
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngKey)
    If rs.BOF Then Exit Function
    txtInfo.Text = txtInfo.Text & _
        "<tr>" & _
        "<td colspan=6 class=xl25>�ϼ�</td>" & _
        "<td colspan=2 class=xl29><font class=font8>" & zlCommFun.NVL(rs("����")) & "</td>" & _
        "<td colspan=2 class=xl29><font class=font8>" & Format(zlCommFun.NVL(rs("����")), "0.00") & "%</td>" & _
        "</tr>"
    
    '3.��������
    strSQL = _
        "Select Distinct A.��������,B.���� " & _
        "From �����Ա���� A, " & _
             "������Ϣ B " & _
        "Where A.��¼���� = 0 " & _
              "And ����id in " & _
                  "( " & _
                   "Select A.ID From ���˲������� A,����Ԫ��Ŀ¼ B " & _
                   "Where A.Ԫ�ر���=B.���� AND upper(B.����)='ZL9CISCORE.USRMEDICALSUM' " & _
                         "And A.������¼id In " & _
                             "( " & _
                              "Select ��첡��id From �����Ա���� Where �Ǽ�id=[1]" & _
                             ") " & _
                  ") " & _
               "And B.����id=A.����id " & _
        "Group By A.��������,B.����"
        
    If mblnDataMoved Then
        strSQL = Replace(strSQL, "���˲�������", "H���˲�������")
        strSQL = Replace(strSQL, "�����Ա����", "H�����Ա����")
        strSQL = Replace(strSQL, "�����Ա����", "H�����Ա����")
    End If
    
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngKey)
    If rs.BOF Then Exit Function
    
    intCount = intCount + 1
    txtInfo.Text = txtInfo.Text & "<tr><td colspan=10 class=xl39 style='font-weight:700'><br>" & intCount & ".��������</td></tr>"
        
    txtInfo.Text = txtInfo.Text & _
        "<tr>" & _
        "<td colspan=6 class=xl25>��������</td>" & _
        "<td colspan=4 class=xl25>����</td>" & _
        "</tr>"
    
    Dim strSvrName As String
    Dim strList As String
    
    Do While Not rs.EOF
        
        If strSvrName <> "" Then
            If strSvrName <> zlCommFun.NVL(rs("��������")) Then
                
                If strList <> "" Then strList = Mid(strList, 2)
                
                txtInfo.Text = txtInfo.Text & _
                    "<tr>" & _
                    "<td colspan=6 class=xl28><font class=font8>" & strSvrName & "</td>" & _
                    "<td colspan=4 class=xl28><font class=font8>" & strList & "</td>" & _
                    "</tr>"
                    
                strList = ""
            End If
        End If
        
        strList = strList & "��" & zlCommFun.NVL(rs("����"))
        strSvrName = zlCommFun.NVL(rs("��������"))
        
        rs.MoveNext
    Loop
                   
    If strSvrName <> "" Then
            
        If strList <> "" Then strList = Mid(strList, 2)
        
        txtInfo.Text = txtInfo.Text & _
            "<tr>" & _
            "<td colspan=6 class=xl28><font class=font8>" & strSvrName & "</td>" & _
            "<td colspan=4 class=xl28><font class=font8>" & strList & "</td>" & _
            "</tr>"
            
        strList = ""
    
    End If
                
    '���
    txtInfo.Text = txtInfo.Text & vbCrLf & "</table></BODY></HTML>"
End Function


Private Function ValidData() As Boolean
    '���
    If Trim(txt(4).Text) = "" Then
        MsgBox "����ȷ���ʼ���������"
        LocationObj txt(4)
        Exit Function
    End If
    
    If Val(txt(0).Text) = 0 Then
        MsgBox "�����ʼ��˿ںţ�һ��Ϊ25����"
        LocationObj txt(0)
        Exit Function
    End If
    
    If Trim(txt(1).Text) = "" Then
        MsgBox "����ȷ�������˵ĵ����ʼ���ַ��"
        LocationObj txt(1)
        Exit Function
    End If
    
    
    If Trim(txt(2).Text) = "" Then
        MsgBox "����ȷ���û�����"
        LocationObj txt(2)
        Exit Function
    End If
    
    If Trim(txt(8).Text) = "" And mlng����id = 0 Then
        MsgBox "����ȷ����������ʼ���ַ��"
        LocationObj txt(8)
        Exit Function
    End If
    
    ValidData = True
    
End Function

Private Sub mnuFileMailGroup_Click()
    Dim objMail As clsMail
    Dim blnSuccess As Boolean
    Dim strMessage As String
    Dim lngLoop As Long
    
    Dim strFile As String
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    
    '���
    If ValidData = False Then Exit Sub
    
    Set objMail = New clsMail
    Set objMail.WinSockObj = sckMail
    
    mblnMaining = True
    
    tbrThis.Buttons("����").Enabled = False
    tbrThis.Buttons("ȫ��").Enabled = False
    tbrThis.Buttons("ȫѡ").Enabled = False
    tbrThis.Buttons("����").Enabled = False
    tbrThis.Buttons("�˳�").Enabled = False
    
    vsf.Editable = flexEDNone
    mnuFile.Enabled = False
    mnuView.Enabled = False
    mnuHelp.Enabled = False
    
    vsf.Cell(flexcpText, 1, mCol.״̬, vsf.Rows - 1, mCol.״̬) = ""
    vsf.Cell(flexcpForeColor, 1, mCol.״̬, vsf.Rows - 1, mCol.״̬) = COLOR.��ɫ
    
    frmWait.OpenWait Me, "���͵����ʼ�"
    frmWait.WaitInfo = "���������ʼ�������..."
    
    objMail.ResponseInternal = Val(txt(6).Text)
    
    If objMail.OpenMailServer(txt(4).Text, txt(2).Text, txt(3).Text, Val(txt(0).Text)) Then
'    If objMail.OpenOutLookExMail() Then
        
        frmWait.WaitInfo = "���ڷ������屨���ʼ�..."
        
        txtInfo.Text = txtHead.Text
        Call GetGroupReportMessageHtml(mlngKey)
        
'        strFile = CreateTmpFile("������챨��.htm")
'        Set objText = objFile.CreateTextFile(strFile, True)
'        objText.Write txtInfo.Text
                
        blnSuccess = objMail.SendHead(Trim(txt(8).Text), txt(2).Text, txt(1).Text, "������챨��", vbMultipartAlternative)
        blnSuccess = objMail.SendMessage(txt(7).Text, vbTextPlain)
        blnSuccess = objMail.SendMessage(txtInfo.Text, vbTextHtml)
        blnSuccess = objMail.SendOver
                
'        objText.Close
        
'        blnSuccess = objMail.SendOutLookExMail(txt(8).Text, "������챨��", txt(7).Text, strFile)
                
        If blnSuccess = False Then ShowSimpleMsg "���屨�淢��ʧ�ܣ�"
                
    End If
    
    frmWait.WaitInfo = "���ڹر��ʼ�������..."
    
    Call objMail.CloseMailServer
'    Call objMail.CloseOutLookExMail
    
    tbrThis.Buttons("����").Enabled = True
    tbrThis.Buttons("ȫ��").Enabled = True
    tbrThis.Buttons("ȫѡ").Enabled = True
    tbrThis.Buttons("����").Enabled = True
    tbrThis.Buttons("�˳�").Enabled = True
    
    vsf.Editable = flexEDKbdMouse
    mnuFile.Enabled = True
    mnuView.Enabled = True
    mnuHelp.Enabled = True
    mblnMaining = False
    
    frmWait.CloseWait
    
End Sub

Private Sub mnuFileOut_Click()
    Dim objMail As clsMail
    Dim blnSuccess As Boolean
    Dim strMessage As String
    Dim lngLoop As Long
    
    Dim strFile As String
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    
    Dim strPath
    
    On Error GoTo errHand
    
    strPath = zlCommFun.OpenDir(Me.hWnd, "ָ�������ļ������Ŀ¼")
    
    If Trim(strPath) = "" Then Exit Sub
    
    If Right(strPath, 1) <> "\" Then strPath = strPath & "\"
        
    mblnMaining = True
    
    tbrThis.Buttons("����").Enabled = False
    tbrThis.Buttons("ȫ��").Enabled = False
    tbrThis.Buttons("ȫѡ").Enabled = False
    tbrThis.Buttons("����").Enabled = False
    tbrThis.Buttons("�˳�").Enabled = False
    
    vsf.Editable = flexEDNone
    mnuFile.Enabled = False
    mnuView.Enabled = False
    mnuHelp.Enabled = False
    
    vsf.Cell(flexcpText, 1, mCol.״̬, vsf.Rows - 1, mCol.״̬) = ""
    vsf.Cell(flexcpForeColor, 1, mCol.״̬, vsf.Rows - 1, mCol.״̬) = COLOR.��ɫ
    
    frmWait.OpenWait Me, "������챨��"
    frmWait.WaitInfo = "��������Html��ʽ����챨��..."
            
    For lngLoop = 1 To vsf.Rows - 1
        If Val(vsf.RowData(lngLoop)) > 0 And Abs(Val(vsf.TextMatrix(lngLoop, mCol.ѡ��))) = 1 Then
            
            frmWait.WaitInfo = "�������ɡ�" & vsf.TextMatrix(lngLoop, mCol.����) & "����Html��ʽ��챨��..."
            
            txtInfo.Text = txtHead.Text
            Call GetReportMessageHtml(mlngKey, Val(vsf.RowData(lngLoop)))
            
            strFile = strPath & "��챨��(" & vsf.TextMatrix(lngLoop, mCol.�����) & "_" & vsf.TextMatrix(lngLoop, mCol.����) & ").htm"
            Set objText = objFile.CreateTextFile(strFile, True)
            objText.Write txtInfo.Text
            objText.Close
                        
        End If
    Next
    
errHand:
    
    tbrThis.Buttons("����").Enabled = True
    tbrThis.Buttons("ȫ��").Enabled = True
    tbrThis.Buttons("ȫѡ").Enabled = True
    tbrThis.Buttons("����").Enabled = True
    tbrThis.Buttons("�˳�").Enabled = True
    
    vsf.Editable = flexEDKbdMouse
    mnuFile.Enabled = True
    mnuView.Enabled = True
    mnuHelp.Enabled = True
    mblnMaining = False
    
    On Error Resume Next
    
    frmWait.CloseWait
End Sub

Private Sub mnuFileOutGroup_Click()
    Dim objMail As clsMail
    Dim blnSuccess As Boolean
    Dim strMessage As String
    Dim lngLoop As Long
    
    Dim strFile As String
    Dim objFile As New FileSystemObject
    Dim objText As TextStream

    Dim strPath
    
    strPath = zlCommFun.OpenDir(Me.hWnd, "ָ�������ļ������Ŀ¼")
    
    If Trim(strPath) = "" Then Exit Sub
    If Right(strPath, 1) <> "\" Then strPath = strPath & "\"
    
    mblnMaining = True
    
    tbrThis.Buttons("����").Enabled = False
    tbrThis.Buttons("ȫ��").Enabled = False
    tbrThis.Buttons("ȫѡ").Enabled = False
    tbrThis.Buttons("����").Enabled = False
    tbrThis.Buttons("�˳�").Enabled = False
    
    vsf.Editable = flexEDNone
    mnuFile.Enabled = False
    mnuView.Enabled = False
    mnuHelp.Enabled = False
    
    vsf.Cell(flexcpText, 1, mCol.״̬, vsf.Rows - 1, mCol.״̬) = ""
    vsf.Cell(flexcpForeColor, 1, mCol.״̬, vsf.Rows - 1, mCol.״̬) = COLOR.��ɫ
    
    frmWait.OpenWait Me, "������챨��"
    frmWait.WaitInfo = "��������Html��ʽ����챨��..."
    
    txtInfo.Text = txtHead.Text
    Call GetGroupReportMessageHtml(mlngKey)
    
'        strFile = CreateTmpFile("������챨��.htm")
    strFile = strPath & "������챨��" & mlngKey & ".htm"
    
    Set objText = objFile.CreateTextFile(strFile, True)
    objText.Write txtInfo.Text
    
    tbrThis.Buttons("����").Enabled = True
    tbrThis.Buttons("ȫ��").Enabled = True
    tbrThis.Buttons("ȫѡ").Enabled = True
    tbrThis.Buttons("����").Enabled = True
    tbrThis.Buttons("�˳�").Enabled = True
    
    vsf.Editable = flexEDKbdMouse
    mnuFile.Enabled = True
    mnuView.Enabled = True
    mnuHelp.Enabled = True
    mblnMaining = False
    
    frmWait.CloseWait
End Sub

Private Sub mnuFileSelectAll_Click()
    Dim lngLoop As Long
    
    For lngLoop = 1 To vsf.Rows - 1
        If Val(vsf.RowData(lngLoop)) > 0 Then
            vsf.TextMatrix(lngLoop, mCol.ѡ��) = 1
            EditChanged = True
        End If
    Next
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
        If mnuFileMail.Visible Then mobjPopMenu.Add 1, mnuFileMail.Caption, , , mnuFileMail.Enabled
        If mnuFileMailGroup.Visible Then mobjPopMenu.Add 2, mnuFileMailGroup.Caption, , , mnuFileMailGroup.Enabled
    Case 2
        
        If mnuFileOut.Visible Then mobjPopMenu.Add 1, mnuFileOut.Caption, , , mnuFileOut.Enabled
        If mnuFileOutGroup.Visible Then mobjPopMenu.Add 2, mnuFileOutGroup.Caption, , , mnuFileOutGroup.Enabled
    Case 3
        
        mobjPopMenu.Add 1, "&1.����", , , True, , (lbl(11).Tag = "����")
        mobjPopMenu.Add 2, "&2.�����", , , True, , (lbl(11).Tag = "�����")
        mobjPopMenu.Add 3, "&3.������", , , True, , (lbl(11).Tag = "������")
        mobjPopMenu.Add 4, "&4.���￨��", , , True, , (lbl(11).Tag = "���￨��")
        mobjPopMenu.Add 5, "&5.����ƴ��", , , True, , (lbl(11).Tag = "����ƴ��")
        mobjPopMenu.Add 6, "&6.�������", , , True, , (lbl(11).Tag = "�������")
        mobjPopMenu.Add 7, "&7.���֤��", , , True, , (lbl(11).Tag = "���֤��")
        mobjPopMenu.Add 8, "&8.�����", , , True, , (lbl(11).Tag = "�����")
        
    End Select
    
End Sub

Private Sub mobjPopMenu_MenuClick(ByVal Key As Long, ByVal Caption As String)
    Select Case mbytPopMenu
    Case 1
        Select Case Key
        Case 1
            Call mnuFileMail_Click
        Case 2
            Call mnuFileMailGroup_Click
        End Select
    Case 2
        Select Case Key
        Case 1
            Call mnuFileOut_Click
        Case 2
            Call mnuFileOutGroup_Click
        End Select
    Case 3
    
        Caption = Mid(Caption, 4)
        
        lbl(11).Caption = "&6." & Left(Trim(Caption), Len(Trim(Caption)) - 1)
        lbl(11).Tag = Left(Trim(Caption), Len(Trim(Caption)) - 1)
        
    End Select
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(tbrThis.hWnd, objPoint)
    
    Select Case Button.Key
    Case "ȫѡ"
        Call mnuFileSelectAll_Click
    Case "ȫ��"
        Call mnuFileClearAll_Click
    Case "����"
    
        mbytPopMenu = 1
        Set mobjPopMenu = New clsPopMenu
        Call mobjPopMenu.ShowPopupMenu(objPoint.X * 15 + Button.Left - 15, objPoint.Y * 15 + Button.Top + Button.Height + 15)
        
    Case "���"
        
        mbytPopMenu = 2
        Set mobjPopMenu = New clsPopMenu
        Call mobjPopMenu.ShowPopupMenu(objPoint.X * 15 + Button.Left - 15, objPoint.Y * 15 + Button.Top + Button.Height + 15)
        
    Case "����"
        Call mnuHelpTopic_Click
    Case "�˳�"
        Call mnuFileExit_Click
    End Select
End Sub

Private Sub tbrThis_ButtonDropDown(ByVal Button As MSComctlLib.Button)
    Call tbrThis_ButtonClick(Button)
End Sub

Private Sub tbrThis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuViewTool
End Sub

Private Sub txt_Change(Index As Integer)
    If Index = 2 Then txt(2).Tag = "Changed"
End Sub

Private Sub txt_GotFocus(Index As Integer)
    If Index <> 7 Then zlControl.TxtSelAll txt(Index)
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim rs As New ADODB.Recordset
    Dim lngLoop As Long
    Dim strCol As String
    Dim lngCol As Long
    Dim lngRow As Long
    
    Dim blnCard As Boolean

    
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
        
    strCol = Mid(lbl(11).Caption, 4)
    lngCol = GetCol(vsf, strCol)
            
    If strCol = "���￨��" And KeyAscii <> vbKeyReturn And Index = 9 Then
        '���￨�ţ��Զ�ʶ��

        blnCard = InputIsCard(txt(Index).Text, KeyAscii)

        If blnCard And Len(txt(Index).Text) = ParamInfo.���￨���볤�� - 1 And KeyAscii <> 8 And txt(Index).Text <> "" Then
            If KeyAscii <> 13 Then
                txt(Index).Text = txt(Index).Text & Chr(KeyAscii)
                txt(Index).SelStart = Len(txt(Index).Text)
            End If
            KeyAscii = vbKeyReturn
        End If
    End If
        
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0

        If Index = 9 And Trim(txt(Index).Text) <> "" Then
            
            strCol = Mid(lbl(11).Caption, 4)
            
            Select Case strCol
            Case "����ƴ��"
                lngCol = GetCol(vsf, "����")
            Case "�������"
                lngCol = GetCol(vsf, "����")
            Case Else
                lngCol = GetCol(vsf, strCol)
            End Select
'            lngCol = GetCol(vsf, strCol)

            If lngCol < 0 Then Exit Sub
            
            lngRow = 0
            If vsf.Row + 1 <= vsf.Rows - 1 Then
                For lngLoop = vsf.Row + 1 To vsf.Rows - 1
                
                    lngRow = 0
                    Select Case strCol
                    Case "�����"
                        If UCase(vsf.TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "������"
                        If UCase(vsf.TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "���￨��"
                        If UCase(vsf.TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "���֤��"
                        If UCase(vsf.TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "����"
                        If UCase(vsf.TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "����ƴ��"
                        If zlGetSymbol(UCase(vsf.TextMatrix(lngLoop, lngCol))) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "�������"
                        If zlGetSymbol(UCase(vsf.TextMatrix(lngLoop, lngCol)), 1) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case Else
                        If UCase(vsf.TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    End Select
            
                    If lngRow > 0 Then Exit For
 
                Next
            End If
            
            If lngRow = 0 Then
                For lngLoop = 1 To vsf.Row

                    lngRow = 0
                    Select Case strCol
                    Case "�����"
                        If UCase(vsf.TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "������"
                        If UCase(vsf.TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "���￨��"
                        If UCase(vsf.TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "���֤��"
                        If UCase(vsf.TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "����"
                        If UCase(vsf.TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "����ƴ��"
                        If zlGetSymbol(UCase(vsf.TextMatrix(lngLoop, lngCol))) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "�������"
                        If zlGetSymbol(UCase(vsf.TextMatrix(lngLoop, lngCol)), 1) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case Else
                        If UCase(vsf.TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    End Select
            
                    If lngRow > 0 Then Exit For
                    
                Next
            End If
            
            If lngRow <= 0 Then
                ShowSimpleMsg "û���ҵ�����Ҫ�����Ϣ��"
                txt(Index).Text = ""
            Else
                vsf.ShowCell lngRow, vsf.Col
                vsf.Row = lngRow
            End If
            
            txt(Index).SetFocus
            zlControl.TxtSelAll txt(Index)
    
        Else
            zlCommFun.PressKey vbKeyTab
        End If
    End If
    
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lngLoop As Long
    
    If Abs(Val(vsf.TextMatrix(Row, mCol.ѡ��))) = 1 Then
        EditChanged = True
        Exit Sub
    End If
        
    For lngLoop = 1 To vsf.Rows - 1
        If Abs(Val(vsf.TextMatrix(lngLoop, mCol.ѡ��))) = 1 Then
            EditChanged = True
            Exit Sub
        End If
    Next
    
    If lngLoop = vsf.Rows Then EditChanged = False
    
End Sub

Private Sub vsf_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Call AppendRows(vsf, lnX, lnY)
End Sub

Private Sub vsf_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call AppendRows(vsf, lnX, lnY)
End Sub

Private Sub vsf_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> mCol.ѡ�� Or Val(vsf.RowData(Row)) <= 0 Then
        Cancel = True
    End If
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

