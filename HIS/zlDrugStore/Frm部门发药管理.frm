VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm���ŷ�ҩ���� 
   Caption         =   "ҩƷ���ŷ�ҩ"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   585
   ClientWidth     =   11760
   DrawMode        =   14  'Copy Pen
   Icon            =   "Frm���ŷ�ҩ����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8160
   ScaleWidth      =   11760
   Begin VB.TextBox txt������ 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0.00000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   0
      EndProperty
      Height          =   225
      Left            =   2640
      MaxLength       =   20
      TabIndex        =   23
      Text            =   "####"
      Top             =   4200
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CheckBox Chk��ʾ��ҩ�������� 
      Appearance      =   0  'Flat
      Caption         =   "��ʾ��ҩ��������"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   8520
      TabIndex        =   14
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton cmdAlley 
      Caption         =   "����ʷ/����״̬"
      Height          =   270
      Left            =   6840
      TabIndex        =   13
      Top             =   4440
      Width           =   1530
   End
   Begin VB.CheckBox Chk�嵥 
      Appearance      =   0  'Flat
      Caption         =   "��ʾ���й��̵���"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   8520
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   4440
      Width           =   1845
   End
   Begin VB.TextBox TxtInput 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   240
      TabIndex        =   11
      Text            =   "####"
      Top             =   4200
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.ComboBox Cbo���� 
      Height          =   300
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   4080
      Visible         =   0   'False
      Width           =   1005
   End
   Begin MSComctlLib.ImageList ImgTbarBlack 
      Left            =   3990
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImgTbarColor 
      Left            =   4560
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin TabDlg.SSTab TabShow 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   4440
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   5530
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      WordWrap        =   0   'False
      TabCaption(0)   =   "δ��ҩ�嵥(&N)"
      TabPicture(0)   =   "Frm���ŷ�ҩ����.frx":1CFA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl��ҩ��"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl��ҩ����ʽ"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Dtp��ѯ����"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Billδ��ҩ�嵥"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cbo��ҩ��"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cbo��ҩ����ʽ"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "���ܷ�ҩ(&T)"
      TabPicture(1)   =   "Frm���ŷ�ҩ����.frx":1D16
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Bill���ܷ�ҩ"
      Tab(1).Control(1)=   "Bill��ҩ����"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "ȱҩ�嵥(&Q)"
      TabPicture(2)   =   "Frm���ŷ�ҩ����.frx":1D32
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Billȱҩ�嵥"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "�ܷ�ҩ�嵥(&D)"
      TabPicture(3)   =   "Frm���ŷ�ҩ����.frx":1D4E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Bill�ܷ�ҩ�嵥"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "����ҩ�嵥(&A)"
      TabPicture(4)   =   "Frm���ŷ�ҩ����.frx":1D6A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Bill�ѷ�ҩ�嵥"
      Tab(4).ControlCount=   1
      Begin VB.ComboBox cbo��ҩ����ʽ 
         Height          =   300
         Left            =   8300
         Style           =   2  'Dropdown List
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1900
      End
      Begin VB.ComboBox cbo��ҩ�� 
         Height          =   300
         Left            =   720
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   "cbo��ҩ��"
         Top             =   2760
         Width           =   1900
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Billδ��ҩ�嵥 
         Height          =   2145
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "���ո��������Ҽ��л�����״̬"
         Top             =   360
         Width           =   6795
         _ExtentX        =   11986
         _ExtentY        =   3784
         _Version        =   393216
         FixedCols       =   0
         GridColor       =   -2147483631
         GridColorFixed  =   8421504
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         FillStyle       =   1
         GridLinesFixed  =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Bill���ܷ�ҩ 
         Height          =   1335
         Left            =   -74880
         TabIndex        =   5
         Top             =   360
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   2355
         _Version        =   393216
         FixedCols       =   0
         GridColor       =   -2147483631
         GridColorFixed  =   8421504
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         FillStyle       =   1
         GridLinesFixed  =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Billȱҩ�嵥 
         Height          =   2055
         Left            =   -74880
         TabIndex        =   6
         ToolTipText     =   "���ո��������Ҽ��л�����״̬"
         Top             =   360
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   3625
         _Version        =   393216
         FixedCols       =   0
         GridColor       =   -2147483631
         GridColorFixed  =   8421504
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         FillStyle       =   1
         GridLinesFixed  =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Bill�ܷ�ҩ�嵥 
         Height          =   2055
         Left            =   -74880
         TabIndex        =   7
         ToolTipText     =   "���ո��������Ҽ��л�����״̬"
         Top             =   360
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   3625
         _Version        =   393216
         FixedCols       =   0
         GridColor       =   -2147483631
         GridColorFixed  =   8421504
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         FillStyle       =   1
         GridLinesFixed  =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Bill�ѷ�ҩ�嵥 
         Height          =   2055
         Left            =   -74880
         TabIndex        =   8
         ToolTipText     =   "���ո��������Ҽ��л�����״̬"
         Top             =   360
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   3625
         _Version        =   393216
         FixedCols       =   0
         GridColor       =   -2147483631
         GridColorFixed  =   8421504
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         FillStyle       =   1
         GridLinesFixed  =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSComCtl2.DTPicker Dtp��ѯ���� 
         Height          =   315
         Left            =   8160
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   960
         Visible         =   0   'False
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   99680259
         CurrentDate     =   36985
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Bill��ҩ���� 
         Height          =   1335
         Left            =   -74880
         TabIndex        =   24
         Top             =   1680
         Visible         =   0   'False
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   2355
         _Version        =   393216
         FixedCols       =   0
         GridColor       =   -2147483631
         GridColorFixed  =   8421504
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         FillStyle       =   1
         GridLinesFixed  =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label lbl��ҩ����ʽ 
         AutoSize        =   -1  'True
         Caption         =   "��ҩ����ʽ"
         Height          =   180
         Left            =   7300
         TabIndex        =   22
         Top             =   2820
         Width           =   900
      End
      Begin VB.Label lbl��ҩ�� 
         AutoSize        =   -1  'True
         Caption         =   "��ҩ��"
         Height          =   180
         Left            =   120
         TabIndex        =   20
         Top             =   2820
         Width           =   540
      End
   End
   Begin ComCtl3.CoolBar Cbar 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   1164
      BandCount       =   2
      _CBWidth        =   11760
      _CBHeight       =   660
      _Version        =   "6.7.8988"
      Child1          =   "Tbar"
      MinWidth1       =   4005
      MinHeight1      =   600
      Width1          =   4770
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Caption2        =   "��ҩҩ��"
      Child2          =   "Cbo��ҩҩ��"
      MinWidth2       =   1695
      MinHeight2      =   300
      Width2          =   1695
      NewRow2         =   0   'False
      Begin VB.ComboBox Cbo��ҩҩ�� 
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   9975
         Style           =   2  'Dropdown List
         TabIndex        =   65
         Top             =   180
         Width           =   1695
      End
      Begin MSComctlLib.Toolbar Tbar 
         Height          =   600
         Left            =   165
         TabIndex        =   2
         Top             =   30
         Width           =   8805
         _ExtentX        =   15531
         _ExtentY        =   1058
         ButtonWidth     =   820
         ButtonHeight    =   1058
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   11
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ԥ��"
               Key             =   "Preview"
               Object.ToolTipText     =   "Ԥ��"
               Object.Tag             =   "Ԥ��"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "Print"
               Object.ToolTipText     =   "��ӡ"
               Object.Tag             =   "��ӡ"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ҩ"
               Key             =   "Consignment"
               Object.ToolTipText     =   "��ҩ"
               Object.Tag             =   "��ҩ"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "����"
               Key             =   "Desire"
               Object.ToolTipText     =   "ȱҩ����"
               Object.Tag             =   "����"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�ܷ�"
               Key             =   "Handback"
               Object.ToolTipText     =   "�ܷ�"
               Object.Tag             =   "�ܷ�"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ҩ"
               Key             =   "Restore"
               Object.ToolTipText     =   "��ҩ"
               Object.Tag             =   "��ҩ"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "ReVerify"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Edit1"
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Exit"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
            EndProperty
         EndProperty
         Begin VB.Timer TimerAuto 
            Enabled         =   0   'False
            Interval        =   10000
            Left            =   6960
            Top             =   240
         End
         Begin MSComctlLib.ImageList imgPass 
            Left            =   5415
            Top             =   0
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   14
            ImageHeight     =   14
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   5
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Frm���ŷ�ҩ����.frx":1D86
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Frm���ŷ�ҩ����.frx":2040
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Frm���ŷ�ҩ����.frx":22FA
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Frm���ŷ�ҩ����.frx":25B4
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Frm���ŷ�ҩ����.frx":286E
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   7800
      Width           =   11760
      _ExtentX        =   20743
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
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15663
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
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
            Picture         =   "Frm���ŷ�ҩ����.frx":2B28
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm���ŷ�ҩ����.frx":2E42
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView Lvw���� 
      Height          =   465
      Left            =   10680
      TabIndex        =   62
      Top             =   4680
      Visible         =   0   'False
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   820
      View            =   2
      Arrange         =   1
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "����"
         Object.Width           =   5292
      EndProperty
   End
   Begin MSComctlLib.ListView Lvw��ҩ;�� 
      Height          =   345
      Left            =   10680
      TabIndex        =   63
      Top             =   5400
      Visible         =   0   'False
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   609
      View            =   2
      Arrange         =   1
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "����"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSComctlLib.ListView Lvw���� 
      Height          =   345
      Left            =   10680
      TabIndex        =   64
      Top             =   6000
      Visible         =   0   'False
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   609
      View            =   2
      Arrange         =   1
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "����"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Frame fraCondition 
      Height          =   3000
      Left            =   20
      TabIndex        =   15
      Top             =   620
      Visible         =   0   'False
      Width           =   12975
      Begin VB.Frame frmLine1 
         Height          =   30
         Left            =   0
         TabIndex        =   25
         Top             =   2300
         Width           =   12945
      End
      Begin VB.Frame fraConRequest 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   20
         TabIndex        =   57
         Top             =   2280
         Visible         =   0   'False
         Width           =   10770
         Begin MSComCtl2.DTPicker Dtp���ʽ���ʱ�� 
            Height          =   315
            Left            =   4200
            TabIndex        =   58
            Top             =   255
            Width           =   2430
            _ExtentX        =   4286
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy��MM��dd�� HH:mm:ss"
            Format          =   99680259
            CurrentDate     =   36985
         End
         Begin MSComCtl2.DTPicker Dtp���ʿ�ʼʱ�� 
            Height          =   315
            Left            =   1320
            TabIndex        =   59
            Top             =   255
            Width           =   2430
            _ExtentX        =   4286
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy��MM��dd�� HH:mm:ss"
            Format          =   99680259
            CurrentDate     =   36985
         End
         Begin VB.Label lblS1 
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Left            =   3840
            TabIndex        =   61
            Top             =   315
            Width           =   180
         End
         Begin VB.Label lblTimeRequest 
            AutoSize        =   -1  'True
            Caption         =   "��������ʱ��"
            Height          =   180
            Left            =   120
            TabIndex        =   60
            Top             =   315
            Width           =   1080
         End
      End
      Begin VB.Frame frmLine 
         Height          =   30
         Left            =   0
         TabIndex        =   16
         Top             =   1440
         Width           =   12945
      End
      Begin VB.Frame fraConNormal 
         BorderStyle     =   0  'None
         Height          =   1320
         Left            =   20
         TabIndex        =   26
         Top             =   100
         Width           =   10800
         Begin VB.CheckBox chkSend 
            Caption         =   "Ժ����ҩ"
            Height          =   180
            Index           =   0
            Left            =   960
            TabIndex        =   34
            Top             =   960
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox chkSend 
            Caption         =   "��ȡҩ"
            Height          =   180
            Index           =   2
            Left            =   3240
            TabIndex        =   33
            Top             =   960
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.Frame fraTypeLine 
            Height          =   400
            Left            =   4200
            TabIndex        =   32
            Top             =   795
            Width           =   30
         End
         Begin VB.TextBox txt���� 
            Height          =   300
            Left            =   3640
            TabIndex        =   31
            Top             =   545
            Width           =   6705
         End
         Begin VB.TextBox txtPati 
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   8790
            TabIndex        =   30
            Top             =   180
            Width           =   1905
         End
         Begin VB.CheckBox chkSendType 
            Caption         =   "��ҩ���ͣ���̬����"
            Height          =   180
            Index           =   0
            Left            =   4320
            TabIndex        =   29
            Top             =   963
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.CheckBox chkSend 
            Caption         =   "��Ժ��ҩ"
            Height          =   180
            Index           =   1
            Left            =   2160
            TabIndex        =   28
            Top             =   960
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CommandButton cmd�������� 
            Caption         =   "��"
            Height          =   300
            Left            =   10320
            TabIndex        =   27
            Top             =   530
            Width           =   375
         End
         Begin MSComCtl2.DTPicker Dtp����ʱ�� 
            Height          =   315
            Left            =   3640
            TabIndex        =   35
            Top             =   180
            Width           =   2430
            _ExtentX        =   4286
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy��MM��dd�� HH:mm:ss"
            Format          =   99680259
            CurrentDate     =   36985
         End
         Begin MSComCtl2.DTPicker Dtp��ʼʱ�� 
            Height          =   315
            Left            =   960
            TabIndex        =   36
            Top             =   180
            Width           =   2430
            _ExtentX        =   4286
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy��MM��dd�� HH:mm:ss"
            Format          =   99680259
            CurrentDate     =   36985
         End
         Begin MSComctlLib.TabStrip tbsType 
            Height          =   255
            Left            =   960
            TabIndex        =   37
            Top             =   570
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   450
            MultiRow        =   -1  'True
            Style           =   2
            HotTracking     =   -1  'True
            Separators      =   -1  'True
            _Version        =   393216
            BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
               NumTabs         =   3
               BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "�ٴ�"
                  Key             =   "T1"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "ҽ��"
                  Key             =   "T2"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "����"
                  Key             =   "T3"
                  ImageVarType    =   2
               EndProperty
            EndProperty
         End
         Begin VB.Label lblPatiInputType 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "סԺ�š�"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   8040
            TabIndex        =   43
            Top             =   240
            Width           =   720
         End
         Begin VB.Label lblInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "������Ϣ"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   7200
            TabIndex        =   42
            Top             =   240
            Width           =   720
         End
         Begin VB.Label lblS 
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Left            =   3400
            TabIndex        =   41
            Top             =   240
            Width           =   180
         End
         Begin VB.Label lblDepType 
            AutoSize        =   -1  'True
            Caption         =   "��ҩ����"
            Height          =   180
            Left            =   120
            TabIndex        =   40
            Top             =   600
            Width           =   720
         End
         Begin VB.Label lbl��ҩ���� 
            AutoSize        =   -1  'True
            Caption         =   "��ҩ����"
            Height          =   180
            Left            =   120
            TabIndex        =   39
            Top             =   963
            Width           =   720
         End
         Begin VB.Label lblTime 
            AutoSize        =   -1  'True
            Caption         =   "ʱ�䷶Χ"
            Height          =   180
            Left            =   120
            TabIndex        =   38
            Top             =   247
            Width           =   720
         End
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "ˢ��(&R)"
         Height          =   345
         Left            =   10850
         TabIndex        =   18
         Top             =   960
         Width           =   900
      End
      Begin VB.CommandButton cmdOtherCon 
         Caption         =   "ȫ������(&C)"
         Height          =   345
         Left            =   11750
         TabIndex        =   17
         Top             =   960
         Width           =   1140
      End
      Begin VB.Frame fraConExpand 
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   20
         TabIndex        =   44
         Top             =   1320
         Visible         =   0   'False
         Width           =   10800
         Begin VB.CheckBox chkType 
            Caption         =   "Ӥ��ҩƷ"
            Height          =   180
            Index           =   1
            Left            =   9600
            TabIndex        =   67
            Top             =   675
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox chkType 
            Caption         =   "����ҩƷ"
            Height          =   180
            Index           =   0
            Left            =   8400
            TabIndex        =   66
            Top             =   675
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CommandButton cmdҩƷ���� 
            Caption         =   "��"
            Height          =   300
            Left            =   10320
            TabIndex        =   52
            Top             =   240
            Width           =   375
         End
         Begin VB.CommandButton cmd��ҩ;�� 
            Caption         =   "��"
            Height          =   300
            Left            =   4575
            TabIndex        =   51
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtҩƷ���� 
            Height          =   300
            Left            =   6120
            TabIndex        =   50
            Top             =   255
            Width           =   4215
         End
         Begin VB.TextBox txt��ҩ;�� 
            Height          =   300
            Left            =   960
            TabIndex        =   49
            Top             =   255
            Width           =   3615
         End
         Begin VB.ComboBox Cboҽ������ 
            Height          =   300
            Left            =   6120
            Style           =   2  'Dropdown List
            TabIndex        =   48
            Top             =   615
            Width           =   1815
         End
         Begin VB.OptionButton opt��Χ 
            Caption         =   "��ҩ����"
            Height          =   225
            Index           =   2
            Left            =   3840
            TabIndex        =   47
            Top             =   675
            Width           =   1125
         End
         Begin VB.OptionButton opt��Χ 
            Caption         =   "��ҩ����"
            Height          =   225
            Index           =   1
            Left            =   2400
            TabIndex        =   46
            Top             =   675
            Width           =   1125
         End
         Begin VB.OptionButton opt��Χ 
            Caption         =   "��������"
            Height          =   225
            Index           =   0
            Left            =   960
            TabIndex        =   45
            Top             =   675
            Value           =   -1  'True
            Width           =   1125
         End
         Begin VB.Label lbl��ҩ;�� 
            AutoSize        =   -1  'True
            Caption         =   "��ҩ;��"
            Height          =   180
            Left            =   120
            TabIndex        =   56
            Top             =   315
            Width           =   720
         End
         Begin VB.Label lblҩƷ���� 
            AutoSize        =   -1  'True
            Caption         =   "ҩƷ����"
            Height          =   180
            Left            =   5280
            TabIndex        =   55
            Top             =   315
            Width           =   720
         End
         Begin VB.Label lbl�������� 
            AutoSize        =   -1  'True
            Caption         =   "����Χ"
            Height          =   180
            Left            =   120
            TabIndex        =   54
            Top             =   675
            Width           =   720
         End
         Begin VB.Label Lblҽ������ 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "ҽ������"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   5280
            TabIndex        =   53
            Top             =   675
            Width           =   720
         End
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu MnuFileSet 
         Caption         =   "��ӡ����(&S)"
      End
      Begin VB.Menu MnuFilePreview 
         Caption         =   "��ӡԤ��(&V)"
      End
      Begin VB.Menu MnuFilePrint 
         Caption         =   "��ӡ(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu MnuFileExcel 
         Caption         =   "�����&Excel"
      End
      Begin VB.Menu MnuFile1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuFileBillprint 
         Caption         =   "���ݴ�ӡ(&B)"
         Shortcut        =   ^B
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFilePrintTotal 
         Caption         =   "��ӡ�����嵥(&C)"
      End
      Begin VB.Menu mnuFileRestore 
         Caption         =   "��ӡ��ҩ֪ͨ��(&R)"
      End
      Begin VB.Menu mnuFileWait 
         Caption         =   "��ӡҩƷ��ҩ��(&W)"
      End
      Begin VB.Menu MnuFile2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuFilePara 
         Caption         =   "��������(&A)"
         Shortcut        =   {F12}
      End
      Begin VB.Menu MnuFile3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuFileExit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu MnuEditVerify 
         Caption         =   "��ҩ(&V)"
      End
      Begin VB.Menu MnuEditDesire 
         Caption         =   "ȱҩ����(&D)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditHandback 
         Caption         =   "�ܷ�ȷ��(&H)"
      End
      Begin VB.Menu MnuEditRestore 
         Caption         =   "��ҩ(&R)"
      End
      Begin VB.Menu mnuEditHandbackBatch 
         Caption         =   "������ҩ���Ĵ���(&T)"
      End
      Begin VB.Menu mnuline1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReVerify 
         Caption         =   "ҩƷ��ҩ����(&B)"
      End
      Begin VB.Menu mnuFlag 
         Caption         =   "ֹͣ��ҩ���(&S)"
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
      Begin VB.Menu MnuViewTool 
         Caption         =   "������(&T)"
         Begin VB.Menu MnuViewToolS 
            Caption         =   "��׼��ť(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu MnuViewToolT 
            Caption         =   "�ı���ǩ(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu MnuViewState 
         Caption         =   "״̬��(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu MnuView1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFont 
         Caption         =   "����(&Z)"
         Begin VB.Menu mnuViewFontSet 
            Caption         =   "С����(&S)"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuViewFontSet 
            Caption         =   "������(&M)"
            Index           =   1
         End
         Begin VB.Menu mnuViewFontSet 
            Caption         =   "������(&B)"
            Index           =   2
         End
      End
      Begin VB.Menu MnuView3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuViewLocate 
         Caption         =   "����(&L)"
         Shortcut        =   ^L
      End
      Begin VB.Menu MnuViewLocateNext 
         Caption         =   "������һ��(&N)"
         Enabled         =   0   'False
         Shortcut        =   {F3}
      End
      Begin VB.Menu MnuView4 
         Caption         =   "-"
      End
      Begin VB.Menu MnuViewTotal 
         Caption         =   "ȫѡ(&A)"
      End
      Begin VB.Menu MnuViewNone 
         Caption         =   "ȫ��(&C)"
      End
      Begin VB.Menu MnuView5 
         Caption         =   "-"
      End
      Begin VB.Menu MnuViewRefresh 
         Caption         =   "ˢ��(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu MnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu MnuHelpTitle 
         Caption         =   "��������(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu MnuHelpWeb 
         Caption         =   "Web�ϵ�����(&W)"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "������ҳ(&H)"
         End
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "������̳(&F)"
         End
         Begin VB.Menu MnuHelpWebM 
            Caption         =   "���ͷ���(&E)..."
         End
      End
      Begin VB.Menu MnuHelp1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuHelpAbout 
         Caption         =   "����(&A)..."
      End
   End
   Begin VB.Menu PopMenu_1 
      Caption         =   "PopMenuδ��ҩ"
      Visible         =   0   'False
      Begin VB.Menu Consignment 
         Caption         =   "��ҩ(&C)"
         Checked         =   -1  'True
      End
      Begin VB.Menu HandBack 
         Caption         =   "�ܷ�(&H)"
         Checked         =   -1  'True
      End
      Begin VB.Menu Lack 
         Caption         =   "ȱҩ(&L)"
         Checked         =   -1  'True
         Enabled         =   0   'False
      End
      Begin VB.Menu Nop_1 
         Caption         =   "������(&N)"
         Checked         =   -1  'True
      End
      Begin VB.Menu Split_1 
         Caption         =   "-"
      End
      Begin VB.Menu ConsignmentALL 
         Caption         =   "ȫ����ҩ(&S)"
      End
      Begin VB.Menu HandBackALL 
         Caption         =   "ȫ���ܷ�(&J)"
      End
      Begin VB.Menu Nop_ALL 
         Caption         =   "ȫ��������(&B)"
      End
   End
   Begin VB.Menu PopMenu_2 
      Caption         =   "PopMenu�ѷ�ҩ"
      Visible         =   0   'False
      Begin VB.Menu Restore 
         Caption         =   "��ҩ(&R)"
         Checked         =   -1  'True
      End
      Begin VB.Menu Nop_2 
         Caption         =   "������(&N)"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu PopMenu_3 
      Caption         =   "PopMenu�ܷ�ҩ"
      Visible         =   0   'False
      Begin VB.Menu ResumeDo 
         Caption         =   "�ָ�(&R)"
         Checked         =   -1  'True
      End
      Begin VB.Menu Nop_3 
         Caption         =   "������(&N)"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuPass 
      Caption         =   "Pass"
      Visible         =   0   'False
      Begin VB.Menu mnuPassItem 
         Caption         =   "ҩ���ٴ���Ϣ�ο�(&C)"
         Index           =   0
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "ҩƷ˵����(&D)"
         Index           =   1
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "�й�ҩ��(&N)"
         Index           =   2
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "������ҩ����(&S)"
         Index           =   3
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "����ֵ(&T)"
         Index           =   4
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "ר����Ϣ(&P)"
         Index           =   6
         Begin VB.Menu mnuPassSpec 
            Caption         =   "ҩ��-ҩ���໥����(&D)"
            Index           =   0
         End
         Begin VB.Menu mnuPassSpec 
            Caption         =   "ҩ��-ʳ���໥����(&F)"
            Index           =   1
         End
         Begin VB.Menu mnuPassSpec 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu mnuPassSpec 
            Caption         =   "����ע�������(&M)"
            Index           =   3
         End
         Begin VB.Menu mnuPassSpec 
            Caption         =   "����ע�������(&T)"
            Index           =   4
         End
         Begin VB.Menu mnuPassSpec 
            Caption         =   "-"
            Index           =   5
         End
         Begin VB.Menu mnuPassSpec 
            Caption         =   "����֢(&C)"
            Index           =   6
         End
         Begin VB.Menu mnuPassSpec 
            Caption         =   "������(&S)"
            Index           =   7
         End
         Begin VB.Menu mnuPassSpec 
            Caption         =   "-"
            Index           =   8
         End
         Begin VB.Menu mnuPassSpec 
            Caption         =   "��������ҩ(&G)"
            Index           =   9
         End
         Begin VB.Menu mnuPassSpec 
            Caption         =   "��ͯ��ҩ(&P)"
            Index           =   10
         End
         Begin VB.Menu mnuPassSpec 
            Caption         =   "��������ҩ(&E)"
            Index           =   11
         End
         Begin VB.Menu mnuPassSpec 
            Caption         =   "��������ҩ(&L)"
            Index           =   12
         End
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "ҽҩ��Ϣ����(&I)"
         Index           =   8
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "ҩƷ�����Ϣ(&M)"
         Index           =   10
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "��ҩ;�������Ϣ(&R)"
         Index           =   11
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "ҽԺҩƷ��Ϣ(&F)"
         Index           =   12
      End
   End
   Begin VB.Menu mnuColHide 
      Caption         =   "ColHide"
      Visible         =   0   'False
      Begin VB.Menu mnuDrugCodeName 
         Caption         =   "ҩƷ(���������)"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuDrugCodeName 
         Caption         =   "ҩƷ(������)"
         Checked         =   -1  'True
         Index           =   1
      End
      Begin VB.Menu mnuDrugCodeName 
         Caption         =   "ҩƷ(������)"
         Checked         =   -1  'True
         Index           =   2
      End
      Begin VB.Menu mnuColHideLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "������"
         Index           =   0
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "Ӣ����"
         Index           =   1
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "����"
         Index           =   2
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "����ҽ��"
         Index           =   3
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "״̬"
         Index           =   4
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "����"
         Index           =   5
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "NO"
         Index           =   6
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "����Ա"
         Index           =   7
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "����"
         Index           =   8
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "����"
         Index           =   9
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "סԺ��"
         Index           =   10
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "���"
         Index           =   11
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "����"
         Index           =   12
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "����"
         Index           =   13
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "��"
         Index           =   14
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "����"
         Index           =   15
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "������"
         Index           =   16
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "׼����"
         Index           =   17
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "��ҩ��"
         Index           =   18
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "����"
         Index           =   19
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "���"
         Index           =   20
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "����"
         Index           =   21
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "Ƶ��"
         Index           =   22
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "�÷�"
         Index           =   23
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "����ʱ��"
         Index           =   24
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "˵��"
         Index           =   25
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "����Ա"
         Index           =   26
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "��ҩʱ��"
         Index           =   27
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "��/��ҩ��"
         Index           =   28
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "�ⷿ��λ"
         Index           =   29
      End
   End
   Begin VB.Menu mnuPatiInfo 
      Caption         =   "������Ϣ"
      Visible         =   0   'False
      Begin VB.Menu mnuInfoItem 
         Caption         =   "סԺ��(&0)"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuInfoItem 
         Caption         =   "����(&1)"
         Index           =   1
      End
      Begin VB.Menu mnuInfoItem 
         Caption         =   "����(&2)"
         Index           =   2
      End
      Begin VB.Menu mnuInfoItem 
         Caption         =   "���ݺ�(&3)"
         Index           =   3
      End
      Begin VB.Menu mnuInfoItem 
         Caption         =   "����ID(&4)"
         Index           =   4
      End
      Begin VB.Menu mnuInfoItem 
         Caption         =   "���￨(&5)"
         Index           =   5
      End
   End
   Begin VB.Menu mnuType 
      Caption         =   "����˵��"
      Visible         =   0   'False
      Begin VB.Menu mnuTypeItem 
         Caption         =   "-"
         Index           =   0
      End
   End
End
Attribute VB_Name = "Frm���ŷ�ҩ����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'--�������--

'Public strPart As String                                'ѡ����,������ʾ���Զ��屨����
Public BlnSetPara As Boolean                            '�������ô����Ƿ�ȷ�����˳�
Public BlnRefresh As Boolean                            '���������Ƿ���������,����ˢ��


'--��ѯ��������--
Private mstr��ʼ����_δ�� As String                        '��ǰʱ��
Private mstr��������_δ�� As String                       '����ʱ��
Private mstr��ʼ����_�ѷ� As String
Private mstr��������_�ѷ� As String
Private mlng����ID As Long
Private mstrסԺ�� As String
Private mstr�������� As String
Private mstrSerchNO As String
Private mstr��ʼNO As String
Private mstr����NO As String
Private mstrDrug As String                                'ҩƷ����
Private mstrUse As String                                 '��ҩ;��
Private mstr���� As String                                'ѡ����
Private mstr�������� As String
Private mint���� As Integer                               'ѡ��������
Private mint��Χ As String                                'ѡ��ҩ��Χ
Private mstr���� As String                                '����
Private mstr��ҩ���� As String
Private mint�������� As Integer                           '0-����;1-Ӥ��;2�����˺�Ӥ��

'--��������--
Private IntCheckStock As Integer                        '�����
Private Int����δ��˴�����ҩ As Integer                'δ����Ƿ�����ҩ
Private lngҩ��ID As Long                               '���ŷ�ҩ
Private Lng����ģʽ As Long                             '��������������ҩ��������
Private Lngҽ������ As Long                             '������Lng����ģʽ��������Lng����ģʽ=������������ʱ������������Ч�����С����������������ʵ�������ҽ����
Private int��Ժ��ҩ As Integer                          '0-����,1-������Ժ��ҩ,2-������Ժ��ҩ,3-������ȡҩ,4-������ȡҩ,5-Ժ����ҩ(��������Ժ��ҩ����ȡҩ),6-��Ժ��ҩ����ȡҩ
Private Lng������ʾ As Long                             '�Ƿ񰴿��һ�����ʾ�����嵥
Private Lng�Զ���ӡ As Long                             '��ҩ���Ƿ��Զ���ӡ
Private Lngȱҩ��� As Long                             '�������ȱҩ���,���޿��ҩƷ������ҩ
Private mlng�������� As Long                            '�Ƿ���ʾ��ҩ��������
Private intDays As Integer                              '��ѯ����
'Private IntSendAfterDosage As Integer                  'δ��ҩ�����Ƿ�����ҩ(ֻ�������ﴦ��)
Private intFont As Integer                              '����
Private StrFindStyle As String                          '����ƥ��
Private lngδ��ҩ��¼ As Long
Private BlnEnterCell As Boolean                         '�Ƿ񼤻�ENTERCELL()�¼�
Private intҩƷ���� As Integer                          'ҩƷ������ʾ��ʽ��0-����������;1-������;2-������
Private Lng��ҩ��ǩ�� As Long
Private Lng��ҩ��ǩ�� As Long
Private str������ As String
Private mblnStarPass As Boolean                         '���ú�����ҩ(PASS)
Private int��ҩ���� As Integer                          '0-ȫ��ʵ�� 1-��ʵ�� 2-���ַ�������
Private int����λ�� As Integer                      '���ý���λ��
Private int��˻��۵� As Integer                        'ִ�к��Զ���˻��۵�
Private mstr������� As String
Private mstr��ֵ���� As String
Private mstr������ҩ��ʽ As String                      '��������
Private mint�Զ�ˢ��δ��ҩ�嵥 As Integer               '0-���Զ�ˢ��
Private mdate�ϴ�ˢ��ʱ�� As Date                       '��¼�ϴ�ˢ��ʱϵͳʱ��
Private mblnAllConditon As Boolean                      '����״̬��
Private mblnҩƷ���� As Boolean                         '�Ƿ���ʾ�ⷿ��λ�����������ʾ
Private mbln��ʾ����ҩ�� As Boolean                     '�Ƿ���ʾ��ҩ����ҩ��
Private mbln���ܷ�ҩ As Boolean                         '���ܷ�ҩʱ�Ƿ�һ��������ҩ���ʼ�¼

'--������ʹ�ñ���--
Private strUnit As String                               '��λ��
Private BlnStartUp As Boolean                           '�����ɹ�
Private BlnFirstStart As Boolean                        '��һ������
Private mblnFirstSended As Boolean
Private Blnˢ��δ��ҩ�嵥 As Boolean                    '����δ��ҩ�嵥�����Ƿ�����ʾǰˢ��
Private Bln����� As Boolean
Private BlnInRefresh As Boolean                         '������ˢ��״̬
Private str����_δ��ҩ As String                        '������
Private str����_����ҩ As String                        '������
Private blnҽ������ As Boolean                          'δ����ҽ���Ƿ�������ҩ
Private str��ҩ�� As String
Private str��ҩ�� As String
Private mstr�۸�ʧЧ��ʾ As String
Private LngLastRow As Long
Private lngLastCol As Long
Private blnҩƷ���������� As Boolean
Private mstrDrawDept As String                          '��ʱ��¼��ҩ����
Private mbln�ͷֱ��� As Boolean                         '�ж��Ƿ��ǵͷֱ��ʣ�800��600��
Private mlng���ܷ�ҩ�� As Variant
Private Const mstrAllType As String = "�ٴ�,����,���,����,����,����,Ӫ��"
Private mbln�Ƿ��������� As Boolean                      'ҩ���Ƿ���С��������ġ�����
Private mblnCard As Boolean                             '�Ƿ�ˢ���￨

Private mstrNo As String
Private mInt���� As Integer

Private mdblConditonHeight As Double
Private mintLastTab As Integer
Private mintLastDeptType As Integer
Private mstrSendDrugId As String

Private mblnDrop As Boolean                     '��KeyDown���ж������б��Ƿ񵯳�

Private Const CB_GETDROPPEDSTATE = &H157
Private Const CB_SHOWDROPDOWN = &H14F

Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private mlngMyWindow As Long

'PASS
Private mlngPatiID As Long
Private mlngPassPati As Long
Private mlng��ҳID As Long
Private mstr�Һŵ� As String
'--������ʹ�õļ�¼��--
Private RecBillData As New adodb.Recordset              'δ��������¼���ѷ�������¼��

Private mrsPASS As New adodb.Recordset                  'PASS�����ݼ�

'--�ڲ���¼��--
Private RecChangeData As adodb.Recordset                '������ʾ��ҳ����(δ��)
Private RecChangeSendedData As adodb.Recordset          '������ʾ�ѷ�ҩ�嵥ҳ�������
Private RecRefreshCompare As adodb.Recordset            '����ˢ��ʱʹ�ã��ָ��ϴ�δ��ҩ�嵥�и���¼���趨״̬��
Private rs��� As New adodb.Recordset
Private mrsRequest As New adodb.Recordset               '������ʾ���������¼
Private mrsRequestMain As New adodb.Recordset

'--���Ҽ�¼��--
Private strFind As String
Private Recδ�� As adodb.Recordset
Private Rec�ѷ� As adodb.Recordset

'--����--
Private Const mlng��ɫ As Long = &HC000C0
Private Const gIntδ��ҩ�嵥ȱҩ As Integer = 0
Private Const gIntδ��ҩ�嵥��ҩ As Integer = 1
Private Const gIntδ��ҩ�嵥�ܷ� As Integer = 2
Private Const gIntδ��ҩ�嵥������ As Integer = 3
Private Const gInt�ѷ�ҩ�嵥��ҩ As Integer = 3
Private Const gInt�ѷ�ҩ�嵥������ As Integer = 1
Private Const gstr�������� As String = "|����|NO|����|����|ҩƷ����|"

'����ɫ
Private Const glngOtherBlkColor As Long = &H80000005        'һ��״̬����ɫ
Private Const glngSendBlkColor As Long = &HFFC0C0           '��ҩ״̬��ǳ��ɫ
Private Const glngSelectBlkColor As Long = &HC0C0C0         '��ǰѡ�񣺻�ɫ

Private mstrPrivs As String                              'Ȩ�޴�
Private mlngMode As Long

Private lng�����嵥���� As Long

Private Enum PatiInfo
    סԺ�� = 0
    ���� = 1
    ���� = 2
    ���ݺ� = 3
    ����ID = 4
    ���￨ = 5
End Enum

Private Type PrivDetail
    Priv_ҽ����ѯ As Boolean
End Type

Private UserPrivDetail As PrivDetail

Private Type CellInfo
    Col As Long
    Row As Long
    CellLeft As Single
    CellTop As Single
    CellHeight As Single
    CellWidth As Single
End Type
Private CurCell As CellInfo

'ҽ���ӿ�
Private gclsInsure As New clsInsure

Private Type TYPE_MedicarePAR
    �������� As Boolean
    �����ϴ� As Boolean
    ������ɺ��ϴ� As Boolean
    ���������ϴ� As Boolean
End Type
Private MCPAR As TYPE_MedicarePAR

Private Const mconstRequest = "����,4,0|NO,7,1200|ҩƷID,7,0|����ʱ��,1,2000|�շ����,7,0|����,1,2000|����,1,1000|Ч��,1,1500|׼������,7,1000|��������,7,1000|��װ,1,0|��λ,1,1000"

Private Enum ����_δ��ҩ�嵥
    ����� = 0
    ����� = 1
    ���� = 2
    ����ҽ�� = 3
    ״̬ = 4
    ���� = 5
    NO = 6
    ����Ա = 7
    ���� = 8
    ���� = 9
    סԺ�� = 10
    ҩƷ���� = 11
    ������ = 12
    Ӣ���� = 13
    ��� = 14
    ���� = 15
    ���� = 16
    �� = 17
    ���� = 18
    ���� = 19
    ��� = 20
    ���� = 21
    Ƶ�� = 22
    �÷� = 23
    ����ʱ�� = 24
    ˵�� = 25
    ���� = 26
    ҽ��id = 27
    ��ҩ�� = 28
    �ⷿ��λ = 29
    ���ID = 30
    ҩƷID = 31
    ������λ = 32
    ��ҩ���� = 33
    ��ҩ����id = 34
    
    ���� = 35
End Enum

Private Enum ����_�ѷ�ҩ�嵥
    ����� = 0
    ����� = 1
    ���� = 2
    ״̬ = 3
    ���� = 4
    NO = 5
    ���� = 6
    ���� = 7
    סԺ�� = 8
    ҩƷ���� = 9
    ������ = 10
    Ӣ���� = 11
    ��� = 12
    ���� = 13
    ���� = 14
    �� = 15
    ���� = 16
    ������ = 17
    ׼���� = 18
    ��ҩ�� = 19
    ���� = 20
    ��� = 21
    ���� = 22
    Ƶ�� = 23
    �÷� = 24
    ����Ա = 25
    ��ҩʱ�� = 26
    ���� = 27
    ҽ��id = 28
    ��ҩ�� = 29
    �ⷿ��λ = 30
    ���ID = 31
    ҩƷID = 32
    ������λ = 33
    
    ���� = 34
End Enum

Private Enum ����_�����嵥
    ҩƷ���� = 0
    ��� = 1
    ���� = 2
    ���� = 3
    ���� = 4
    ��λ = 5
    ���� = 6
    ��� = 7
        
    ���� = 8
End Enum

Private Enum ����_���һ����嵥
    ���� = 0
    ҩƷ���� = 1
    ��� = 2
    ���� = 3
    ���� = 4
    Ӧ������ = 5
    �������� = 6
    �������� = 7
    ʵ������ = 8
    ��λ = 9
    ���� = 10
    ��� = 11
    ���� = 12
    ����ID = 13
    ҩƷID = 14
    ��ҩ���� = 15
    ��ҩ����id = 16
    
    ���� = 17
End Enum

Private Enum �����б�
    ���� = 0
    NO = 1
    ҩƷID = 2
    ����ʱ�� = 3
    �շ���� = 4
    ���� = 5
    ���� = 6
    Ч�� = 7
    ׼������ = 8
    �������� = 9
    ��װ = 10
    ��λ = 11
  
    ���� = 12
End Enum

Private Function CheckGroupSend(ByVal lng���ID As Long) As Boolean
    '���ͬ��ҩƷ�Ƿ��ܹ�����
    'ǰ����ҩ������������������
    'ͬ��ҩƷ��ֻ�е����ж��Ƿ�ҩ״̬����������ȱҩ���ܷ������������ܷ�ҩ
    Dim rsGroupRec As adodb.Recordset     'Ϊ��ҩ���ݼ�RecChangeData��һ������
    Dim i As Integer

    'Ĭ��������
    CheckGroupSend = True
    
    '���������������޸ù���
    If mbln�Ƿ��������� = False Then Exit Function
    
    '�޷���Ĳ���
    If lng���ID = 0 Then Exit Function
    
    '������ҩ���ݼ��ĸ���
    Set rsGroupRec = RecChangeData.Clone
    
    '���ݴ����NO�����ID���ж��Ƿ����ҩƷ���ܷ�ҩ
    With rsGroupRec
        .Filter = "���ID=" & lng���ID
        
        If .EOF Then Exit Function
        
        Do While Not .EOF
            'ֻҪ����ִ��״̬��Ϊ1���Ͳ��ܷ�ҩ
            If !ִ��״̬ <> 1 Then
                CheckGroupSend = False
                Exit Function
            End If
            .MoveNext
        Loop
    End With
End Function

Private Function CheckIsCenter(ByVal lngStockId As Long) As Boolean
    '����ҩ���Ƿ���С��������ġ�����
    Dim rsTmp As adodb.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "Select 1 From ��������˵�� Where �������� = '��������' And ����id = [1]"
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ����������������", lngStockId)
    
    If Not rsTmp.EOF Then CheckIsCenter = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


'�������������
Private Function GetDepend() As Boolean
    Dim rsDepend As New Recordset
    Dim int��ϵ�� As Integer, int��ϵ�� As Integer
    
    On Error GoTo errHandle
    GetDepend = False
    gstrSQL = "SELECT B.Id,b.ϵ��, b.���� " _
        & " FROM ҩƷ�������� A, ҩƷ������ B " _
        & "Where A.���id = B.ID " _
      & "AND A.���� = 27  "
    Call SQLTest(App.Title, "ҩƷ���ŷ�ҩ", gstrSQL)
    Set rsDepend = zldatabase.OpenSQLRecord(gstrSQL, "GetDepend")

    Call SQLTest
    
    If rsDepend.EOF Then
        rsDepend.Close
        Exit Function
    End If
    rsDepend.Close
    
    GetDepend = True
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetPatiInfo(ByVal intType As Integer, ByVal strInfo As String) As String
    'intType��PatiInfo����Ŀֵ
    '���ز�����Ϣ����ǰ������ID�Ͳ������ƣ���������Ϣ��ID��������
    '��ʽ��13,һ����|1,����
    Dim rsTemp As adodb.Recordset
    Dim vRect As RECT, sngX As Single, sngY As Single
    Dim lngH As Long
    Dim blnCancel As Boolean
    
    On Error GoTo errHandle
    If intType = PatiInfo.סԺ�� Then
        If Not IsNumeric(strInfo) Then Exit Function
        
        gstrSQL = "Select A.��ǰ����id As ����id, C.���� || '-' || C.���� As ��������, B.����id, B.���� As �������� " & _
            " From ������ҳ A, ������Ϣ B, ���ű� C " & _
            " Where A.����id = B.����id And A.��ҳid = B.סԺ���� And A.��ǰ����id = C.ID And B.סԺ�� = [1]"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "ȡ������Ϣ", strInfo)
    ElseIf intType = PatiInfo.����ID Then
        If Not IsNumeric(strInfo) Then Exit Function
        
        gstrSQL = "Select A.��ǰ����id As ����id, C.���� || '-' || C.���� As ��������, B.����id, B.���� As �������� " & _
            " From ������ҳ A, ������Ϣ B, ���ű� C " & _
            " Where A.����id = B.����id And A.��ҳid = B.סԺ���� And A.��ǰ����id = C.ID And A.����id = [1]"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "ȡ������Ϣ", Val(strInfo))
    ElseIf intType = PatiInfo.���ݺ� Then
        gstrSQL = "Select Distinct A.���˲���id As ����id, B.���� || '-' || B.���� As ��������, A.����id, A.���� As �������� " & _
            " From ���˷��ü�¼ A, ���ű� B " & _
            " Where A.���˲���id = B.ID And NO = [1] "
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "ȡ������Ϣ", strInfo)
    ElseIf intType = PatiInfo.���� Then
        gstrSQL = "Select A.��ǰ����id As ����id, C.���� || '-' || C.���� As ��������, B.����id, B.���� As �������� " & _
            " From ������ҳ A, ������Ϣ B, ���ű� C " & _
            " Where A.����id = B.����id And A.��ҳid = B.סԺ���� And A.��ǰ����id = C.ID And B.��ǰ���� = [1]"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "ȡ������Ϣ", strInfo)
    ElseIf intType = PatiInfo.���� Then
        If mblnCard = True Then
            gstrSQL = "Select A.��ǰ����id As ����id, C.���� || '-' || C.���� As ��������, B.����id, B.���� As �������� " & _
                " From ������ҳ A, ������Ϣ B, ���ű� C " & _
                " Where A.����id = B.����id And A.��ҳid = B.סԺ���� And A.��ǰ����id = C.ID And B.���￨�� = [1]"
            Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "ȡ������Ϣ", strInfo)
        Else
            '�������ƿ��ܻ����ظ��������б�ѡ��
            gstrSQL = "Select Rownum As ID, ��������, ����id, ��������, ����id" & _
                " From (Select Distinct B.���� As ��������, B.����id, A.��ǰ����id As ����id, C.���� || '-' || C.���� As �������� " & _
                " From ������ҳ A, ������Ϣ B, ���ű� C " & _
                " Where A.����id = B.����id And A.��ҳid = B.סԺ���� And A.��ǰ����id = C.ID And B.���� Like [1])"
            
            vRect = GetControlRect(txtPati.hWnd)
            lngH = txtPati.Height
            sngX = vRect.Left - 15
            sngY = vRect.Top
            
            Set rsTemp = zldatabase.ShowSQLSelect(Me, gstrSQL, 0, "ȡ������Ϣ", False, "", "", False, False, True, sngX, sngY, lngH, blnCancel, False, False, "%" & strInfo & "%")
            If blnCancel = True Then Exit Function
        End If
    ElseIf intType = PatiInfo.���￨ Then
        gstrSQL = "Select A.��ǰ����id As ����id, C.���� || '-' || C.���� As ��������, B.����id, B.���� As �������� " & _
            " From ������ҳ A, ������Ϣ B, ���ű� C " & _
            " Where A.����id = B.����id And A.��ҳid = B.סԺ���� And A.��ǰ����id = C.ID And B.���￨�� = [1]"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "ȡ������Ϣ", UCase(strInfo))
    End If
    
    If rsTemp.EOF Then Exit Function
    
    GetPatiInfo = rsTemp!����id & "," & rsTemp!�������� & "|" & rsTemp!����ID & "," & rsTemp!��������
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub GetPrivs()
    With UserPrivDetail
        .Priv_ҽ����ѯ = IsHavePrivs(mstrPrivs, "ҽ����ѯ")
    End With
End Sub

Private Function GetSumSended(ByVal int���� As Integer, ByVal strNo As String, ByVal lngҩƷID As Long, ByVal int��� As Integer)
    Dim rsTmp As adodb.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "Select Sum(Nvl(����, 1) * ʵ������) �ѷ����� From ҩƷ�շ���¼ Where ���� = [1] And NO = [2] And ҩƷID+0 = [3] And ��� = [4]"
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "�����ѷ�����", int����, strNo, lngҩƷID, int���)
    
    If Not rsTmp.EOF Then
        GetSumSended = rsTmp!�ѷ�����
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Get��ҩ����ʽ()
    Dim rsTemp As adodb.Recordset
    Dim n As Integer
    
    On Error GoTo errHandle
    gstrSQL = "Select ˵�� From zltools.zlRPTFMTs Where ����id = (Select ID From zltools.zlReports Where ��� = 'ZL1_BILL_1342') Order By ���"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "ȡ��ҩ����ʽ")
    
    If rsTemp.RecordCount > 0 Then
        For n = 0 To rsTemp.RecordCount - 1
            cbo��ҩ����ʽ.AddItem rsTemp!˵��
            rsTemp.MoveNext
        Next
          
        cbo��ҩ����ʽ.ListIndex = 0
        
        If rsTemp.RecordCount = 1 Then
            cbo��ҩ����ʽ.Enabled = False
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Get����()
    Dim bln��ҩ�ⷿ As Boolean
    Dim rsTmp As adodb.Recordset
    
    On Error GoTo errHandle
    '��ȡ���м���
    bln��ҩ�ⷿ = False
    gstrSQL = "Select 1 From ��������˵�� " & _
         " Where �������� Like '��ҩ%' And ����ID=[1]"
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[��鲿������]", lngҩ��ID)
    
    If Not rsTmp.EOF Then bln��ҩ�ⷿ = True
    
    gstrSQL = "Select Distinct J.����||'-'||J.���� ����" & _
         " From ����ִ�п��� A,ҩƷ���� B,ҩƷ���� J " & _
         " Where A.������ĿID=B.ҩ��ID And B.ҩƷ����=J.����" & _
         " And A.ִ�п���ID=[1]"
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[��ȡ�ÿⷿ���ڼ���]", lngҩ��ID)
    
    With rsTmp
        Lvw����.ListItems.Clear
        Lvw����.ListItems.Add , "_" & Lvw����.ListItems.Count + 1, "����ҩƷ����", 1, 1
        Lvw����.ListItems(Lvw����.ListItems.Count).Checked = True
        Do While Not .EOF
            Lvw����.ListItems.Add , "_" & Lvw����.ListItems.Count + 1, !����, 1, 1
            Lvw����.ListItems(Lvw����.ListItems.Count).Checked = True
            .MoveNext
        Loop
        If bln��ҩ�ⷿ Then
           Lvw����.ListItems.Add , "_" & Lvw����.ListItems.Count + 1, "0-����", 1, 1
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Get��������(ByVal lng����id As Long, ByVal lngҩƷID As Long) As Double
    Dim dblSum As Double
    
    With mrsRequest
        .Filter = "��ҩ����id=" & lng����id & " And ҩƷID=" & lngҩƷID & " And ��˱�־ = 1"
        If .EOF Then Exit Function
        
        Do While Not .EOF
            dblSum = dblSum + !�������� / !��װ
            .MoveNext
        Loop
    End With
    
    Get�������� = dblSum
End Function

Private Sub IniConditon()
    Dim dateCurDate As Date
    Dim rsTmp As New adodb.Recordset
    Dim n As Integer
    Const cstÿ���ֽڿ�� = 128
    
    On Error GoTo errHandle
    dateCurDate = zldatabase.Currentdate()
    Me.Dtp��ʼʱ��.Value = Format(DateAdd("d", -1 * intDays, dateCurDate), "yyyy-MM-dd 00:00:00")
    Me.Dtp����ʱ��.Value = Format(dateCurDate, "yyyy-MM-dd") & " 23:59:59"
    
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "���ŷ�ҩ����\�ѷ�ҩ�嵥����", "ʱ�䷶Χ", Me.Dtp��ʼʱ��.Value & ";" & Me.Dtp����ʱ��.Value
    
    Me.Dtp���ʿ�ʼʱ��.Value = Me.Dtp��ʼʱ��.Value
    Me.Dtp���ʽ���ʱ��.Value = Me.Dtp����ʱ��.Value
    
    'Ĭ�ϵ���ҩ���������ǲ���
    mintLastDeptType = 2
    tbsType.Tabs(3).Selected = True
    
    '��ȡ����
    Call Get����
    
    '��ȡ��ҩ���ͣ�����̬���ӷ�ҩ����ѡ���
    gstrSQL = "Select ���� From ��ҩ���� Order By ����"
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[��ȡ��ҩ����]")
    
    If rsTmp.RecordCount > 0 Then
        chkSendType(0).Visible = True
        chkSendType(0).Caption = rsTmp!����
        chkSendType(0).Width = 150 + LenB(chkSendType(0).Caption) * 128
        If rsTmp.RecordCount > 1 Then
            rsTmp.MoveNext
            For n = 2 To rsTmp.RecordCount
                Load chkSendType(n - 1)
                chkSendType(n - 1).Visible = True
                chkSendType(n - 1).Caption = rsTmp!����
                chkSendType(n - 1).Width = 150 + LenB(chkSendType(n - 1).Caption) * 128
                rsTmp.MoveNext
            Next
        End If
        
        Call ResizeCheckControl
    End If
    
    '����ҽ������
    With Cboҽ������
        .Clear
        .AddItem "0-�������е���"
        .AddItem "1-��������ҽ��"
        .AddItem "2-������ʱҽ��"
        .AddItem "3-��ͨ���ʵ���"
        .AddItem "4-��������ҽ��"
        .ListIndex = Lngҽ������
    End With
    
    '��ȡ���и�ҩ;��
    gstrSQL = "Select ���� as �÷� ,�걾��λ As ���� From ������ĿĿ¼ Where ���='E' And ��������='2'And (�������=2 Or �������=3) " & _
            " And (����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd') Or ����ʱ�� Is Null) Order by ���� "
    Call zldatabase.OpenRecordset(rsTmp, gstrSQL, Me.Caption)
    With rsTmp
        Lvw��ҩ;��.ListItems.Add , "_" & Lvw��ҩ;��.ListItems.Count + 1, "���и�ҩ;��", 1, 1
        Lvw��ҩ;��.ListItems(Lvw��ҩ;��.ListItems.Count).Checked = True
        Do While Not .EOF
            Lvw��ҩ;��.ListItems.Add , "_" & Lvw��ҩ;��.ListItems.Count + 1, !�÷�, 1, 1
            Lvw��ҩ;��.ListItems(Lvw��ҩ;��.ListItems.Count).Checked = True
            Lvw��ҩ;��.ListItems(Lvw��ҩ;��.ListItems.Count).Tag = !����
            .MoveNext
        Loop
    End With
    
    '���ø�ҩ;������
    gstrSQL = "Select Distinct �걾��λ As ���� From ������ĿĿ¼ Where ��� = 'E' And �������� = '2' And �걾��λ Is Not Null"
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "ȡ��ҩ;������")
    
    If rsTmp.RecordCount = 0 Then Exit Sub
    
    mnuTypeItem.Item(0).Caption = rsTmp!����
    
    If rsTmp.RecordCount > 1 Then
        rsTmp.MoveNext
        For n = 2 To rsTmp.RecordCount
            Load mnuTypeItem.Item(n - 1)
            mnuTypeItem.Item(n - 1).Caption = rsTmp!����
            mnuTypeItem.Item(n - 1).Visible = True
            rsTmp.MoveNext
        Next
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function SelectDept(ByVal intType As Integer, ByVal strInput As String) As adodb.Recordset
    Dim strSQL As String
    Dim dblX As Double
    Dim dblY As Double
    Dim DblHeight As Double
    
    dblX = fraCondition.Left + fraConNormal.Left + txt����.Left
    dblY = fraCondition.Top + fraConNormal.Top + txt����.Top + txt����.Height
    DblHeight = 5000
    
    If intType = 0 Then
        strSQL = " Select ID, ����||'-'||���� ���� From ���ű� " & _
                 " Where (վ�� = '" & gstrNodeNo & "' Or վ�� is Null) And ID in (Select ����ID From ��������˵�� Where ��������='�ٴ�' And ������� IN(2,3))" & _
                 " And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','yyyy-MM-dd')) "
    ElseIf intType = 1 Then
        strSQL = " Select ID, ����||'-'||���� ���� From ���ű� " & _
                 " Where (վ�� = '" & gstrNodeNo & "' Or վ�� is Null) And ID in (Select ����ID From ��������˵�� Where �������� In ('���','����','����','����') And ������� IN(2,3))" & _
                 " And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','yyyy-MM-dd')) "
    Else
        strSQL = " Select ID, ����||'-'||���� ���� From ���ű� " & _
                 " Where (վ�� = '" & gstrNodeNo & "' Or վ�� is Null) And ID in (Select ����ID From ��������˵�� Where ��������='����' And ������� IN(2,3))" & _
                 " And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','yyyy-MM-dd')) "
    End If
    
    strSQL = strSQL & " And (Upper(����) Like '" & UCase(strInput) & "%'" & _
            " Or Upper(����) Like '" & StrFindStyle & UCase(strInput) & "%'" & _
            " Or Upper(����) Like '" & StrFindStyle & UCase(strInput) & "%')"
            
    strSQL = strSQL & " Order By ����||'-'||����"
    
    Set SelectDept = zldatabase.ShowSelect(Me, strSQL, 0, "�����б�", , , , , True, , dblX, dblY, DblHeight)
End Function
Private Sub Get��ҩ��()
    Dim strSQL As String
    Dim rsTemp As New adodb.Recordset
    
    On Error GoTo errHandle
    '���ü�����
    gstrSQL = "Select Distinct A.����||'-'||A.���� As ����" & _
             " From ��Ա�� A,������Ա B,��������˵�� C,��Ա����˵�� D " & _
             " Where (A.վ�� = '" & gstrNodeNo & "' Or A.վ�� is Null) And A.Id=B.��Աid And B.����id=C.����Id And D.��Աid=A.Id And D.��Ա���� = 'ҩ����ҩ��' " & _
             " And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) AND B.����id=[1] " & _
             " ORDER BY ���� "

    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "��ȡҩ����Ա", lngҩ��ID)
    
    cbo��ҩ��.Clear
    Do While Not rsTemp.EOF
        cbo��ҩ��.AddItem rsTemp!����
        rsTemp.MoveNext
    Loop
    
    cbo��ҩ��.Text = gstrUserAbbr & "-" & gstrUserName
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub LoadCondition(ByVal intType As Integer)
    Dim strPath As String
    Dim strTemp As String
    Dim strBegin As String
    Dim strEnd As String
    Dim dateCurDate As Date
    Dim n As Integer
    Dim i As Integer
    Dim arrStr
    
    If BlnFirstStart = False And gblnMyStyle = False Then
        dateCurDate = zldatabase.Currentdate()
        Me.Dtp��ʼʱ��.Value = Format(DateAdd("d", -1 * intDays, dateCurDate), "yyyy-MM-dd 00:00:00")
        Me.Dtp����ʱ��.Value = Format(dateCurDate, "yyyy-MM-dd") & " 23:59:59"
        
        Me.Dtp���ʿ�ʼʱ��.Value = Me.Dtp��ʼʱ��.Value
        Me.Dtp���ʽ���ʱ��.Value = Me.Dtp����ʱ��.Value
        
        Cboҽ������.ListIndex = Lngҽ������
        
        '0-����,1-������Ժ��ҩ,2-������Ժ��ҩ,3-������ȡҩ,4-������ȡҩ,5-Ժ����ҩ(��������Ժ��ҩ����ȡҩ),6-��Ժ��ҩ����ȡҩ
        If int��Ժ��ҩ = 1 Then
            chkSend(0).Value = 1
            chkSend(1).Value = 0
            chkSend(2).Value = 1
        ElseIf int��Ժ��ҩ = 2 Then
            chkSend(0).Value = 0
            chkSend(1).Value = 1
            chkSend(2).Value = 0
        ElseIf int��Ժ��ҩ = 3 Then
            chkSend(0).Value = 1
            chkSend(1).Value = 1
            chkSend(2).Value = 0
        ElseIf int��Ժ��ҩ = 4 Then
            chkSend(0).Value = 0
            chkSend(1).Value = 0
            chkSend(2).Value = 1
        ElseIf int��Ժ��ҩ = 5 Then
            chkSend(0).Value = 1
            chkSend(1).Value = 0
            chkSend(2).Value = 0
        ElseIf int��Ժ��ҩ = 6 Then
            chkSend(0).Value = 0
            chkSend(1).Value = 1
            chkSend(2).Value = 1
        Else
            chkSend(0).Value = 1
            chkSend(1).Value = 1
            chkSend(2).Value = 1
        End If
        Exit Sub
    End If
    
    If intType = 0 Then
        strPath = "δ��ҩ�嵥����"
    Else
        strPath = "�ѷ�ҩ�嵥����"
    End If
    
    '����
    mblnAllConditon = (Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "���ŷ�ҩ����\" & strPath, "ȫ������", "0")) = 1)
    If mblnAllConditon = True Then
        cmdOtherCon.Caption = "��Ҫ����(&C)"
    Else
        cmdOtherCon.Caption = "ȫ������(&C)"
    End If
    Call ResizeCondition
    
    'ʱ�䷶Χ
    strTemp = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "���ŷ�ҩ����\" & strPath, "ʱ�䷶Χ", "")
    If strTemp = "" Or InStr(strTemp, ";") = 0 Then
        dateCurDate = zldatabase.Currentdate()
        strTemp = Format(DateAdd("d", -1 * intDays, dateCurDate), "yyyy-MM-dd 00:00:00") & ";" & Format(dateCurDate, "yyyy-MM-dd") & " 23:59:59"
    Else
        strBegin = Split(strTemp, ";")(0)
        strEnd = Split(strTemp, ";")(1)
        
        If Not IsDate(strBegin) Then
            dateCurDate = zldatabase.Currentdate()
            strBegin = Format(DateAdd("d", -1 * intDays, dateCurDate), "yyyy-MM-dd 00:00:00")
        End If
        
        If Not IsDate(strEnd) Then
            dateCurDate = zldatabase.Currentdate()
            strEnd = Format(dateCurDate, "yyyy-MM-dd") & " 23:59:59"
        End If
        
        strTemp = strBegin & ";" & strEnd
    End If
    
    Dtp��ʼʱ��.Value = Split(strTemp, ";")(0)
    Dtp����ʱ��.Value = Split(strTemp, ";")(1)
        
    '������Ϣ
    strTemp = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "���ŷ�ҩ����\" & strPath, "������Ϣ", "0;")
    If Val(Split(strTemp, ";")(0)) < 0 Or Val(Split(strTemp, ";")(0)) > 4 Then
        Call mnuInfoItem_Click(0)
    Else
        Call mnuInfoItem_Click(Val(Split(strTemp, ";")(0)))
    End If
    txtPati.Text = Split(strTemp, ";")(1)
    
    '����
    strTemp = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "���ŷ�ҩ����\" & strPath, "����", "")
    If strTemp = "" Or InStr(strTemp, ";") = 0 Then
        tbsType.Tabs(3).Selected = True
        txt����.Text = ""
    Else
        If Val(Split(strTemp, ";")(0)) < 0 Or Val(Split(strTemp, ";")(0)) > 2 Then
            tbsType.Tabs(3).Selected = True
        Else
            tbsType.Tabs(Val(Split(strTemp, ";")(0)) + 1).Selected = True
        End If
        txt����.Tag = Split(strTemp, ";")(1)
        txt����.Text = Split(strTemp, ";")(2)
    End If
    
    '��ҩ����
    strTemp = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "���ŷ�ҩ����\" & strPath, "��ҩ����", "0;")
    If strTemp = "" Or InStr(strTemp, ";") = 0 Then
        chkSend(0).Value = 1
        chkSend(1).Value = 1
        chkSend(2).Value = 1
        
        If chkSendType(0).Visible = True Then
            For n = 0 To chkSendType.UBound
                chkSendType(n).Value = 0
            Next
        End If
    ElseIf Val(Split(strTemp, ";")(0)) < 0 Or Val(Split(strTemp, ";")(0)) > 5 Then
        chkSend(0).Value = 1
        chkSend(1).Value = 1
        chkSend(2).Value = 1
    Else
        '0-����,1-������Ժ��ҩ,2-������Ժ��ҩ,3-������ȡҩ,4-������ȡҩ,5-Ժ����ҩ����������Ժ��ҩ����ȡҩ��
        If Val(Split(strTemp, ";")(0)) = 1 Then
            chkSend(0).Value = 1
            chkSend(1).Value = 0
            chkSend(2).Value = 1
        ElseIf Val(Split(strTemp, ";")(0)) = 2 Then
            chkSend(0).Value = 0
            chkSend(1).Value = 1
            chkSend(2).Value = 0
        ElseIf Val(Split(strTemp, ";")(0)) = 3 Then
            chkSend(0).Value = 1
            chkSend(1).Value = 1
            chkSend(2).Value = 0
        ElseIf Val(Split(strTemp, ";")(0)) = 4 Then
            chkSend(0).Value = 0
            chkSend(1).Value = 0
            chkSend(2).Value = 1
        ElseIf Val(Split(strTemp, ";")(0)) = 5 Then
            chkSend(0).Value = 1
            chkSend(1).Value = 0
            chkSend(2).Value = 0
        ElseIf Val(Split(strTemp, ";")(0)) = 6 Then
            chkSend(0).Value = 0
            chkSend(1).Value = 1
            chkSend(2).Value = 1
        Else
            chkSend(0).Value = 1
            chkSend(1).Value = 1
            chkSend(2).Value = 1
        End If
        
        If Split(strTemp, ";")(1) <> "" Then
            arrStr = Split(Split(strTemp, ";")(1), ",")
            For n = 0 To UBound(arrStr)
                For i = 0 To chkSendType.UBound
                    If arrStr(n) = chkSendType(i).Caption Then
                        chkSendType(i).Value = 1
                    End If
                Next
            Next
        End If
    End If
    
    '��ҩ;��
    strTemp = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "���ŷ�ҩ����\" & strPath, "��ҩ;��", "���и�ҩ;��")
    txt��ҩ;��.Text = strTemp
    
    'ҩƷ����
    strTemp = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "���ŷ�ҩ����\" & strPath, "ҩƷ����", "����ҩƷ����")
    txtҩƷ����.Text = strTemp
    
    '����Χ
    strTemp = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "���ŷ�ҩ����\" & strPath, "����Χ", "0"))
    If Val(strTemp) < 0 Or Val(strTemp) > 2 Then
        strTemp = "0"
    End If
    opt��Χ(Val(strTemp)).Value = True
    
    'ҽ������
    strTemp = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "���ŷ�ҩ����\" & strPath, "ҽ������", "0")
    If Val(strTemp) < 0 Or Val(strTemp) > 4 Then
        strTemp = "0"
    End If
    Cboҽ������.ListIndex = Val(strTemp)
    
    
End Sub


Private Function Get�����嵥() As Boolean
    Dim strSubUnit As String
    Dim rsTemp As adodb.Recordset
    Dim strCon As String
    Dim strTmpCon As String
    Dim str����ʱ�� As String
    Dim lng��ҩ����ID As Long
    Dim lngҩƷID As Long
    Dim lng����id As Long
    Dim dbl׼������ As Double

    '��λ����װ����
    On Error GoTo errHandle
    Select Case strUnit
    Case "�ۼ۵�λ"
        strSubUnit = "X.���㵥λ ��λ,1 ��װ,C.ʵ������ As ׼������,A.���� As ��������"
    Case "���ﵥλ"
        strSubUnit = "D.���ﵥλ ��λ,D.�����װ ��װ,C.ʵ������ As ׼������,A.���� As ��������"
    Case "סԺ��λ"
        strSubUnit = "D.סԺ��λ ��λ,D.סԺ��װ ��װ,C.ʵ������ As ׼������,A.���� As ��������"
    Case "ҩ�ⵥλ"
        strSubUnit = "D.ҩ�ⵥλ ��λ,D.ҩ���װ ��װ,C.ʵ������ As ׼������,A.���� As ��������"
    End Select
    
    If mint���� = 0 Then
        strCon = " And H.Id = B.���˿���id "
    ElseIf mint���� = 1 Then
        strCon = " And H.Id = B.��������id "
    Else
        strCon = " And H.Id = B.���˲���ID "
    End If

    If mstrSerchNO <> "" Then
    ElseIf mstrסԺ�� <> "" Then
        strCon = strCon & " And B.��ʶ��=[4] "
    ElseIf mstr�������� <> "" Then
        strCon = strCon & " And B.���� Like [5] "
    ElseIf mlng����ID <> 0 Then
        strCon = strCon & " And B.����ID=[6] "
    ElseIf mstr���� <> "" Then
        strCon = strCon & " And B.���� = [7] "
    End If
    
    gstrSQL = "Select /*+rule*/ Distinct '['||X.����||']'||" & IIf(mblnTradeName, "NVL(K.����,X.����)", "X.����") & " As ҩƷ����, " & _
        " C.ID As �շ�ID, C.ҩƷID, C.����, C.NO, C.��� As �շ����, C.����, C.����, C.Ч��, F.����, P.���� As ��������,H.���� As ��ҩ����,H.Id As ��ҩ����Id, " & _
        " A.����id, B.��� As �������, B.��¼����, B.��ҳID, A.����ʱ��, " & strSubUnit & " " & _
        " From ���˷������� A, ���˷��ü�¼ B," & _
        " (Select A.ID, A.����, A.NO, A.���, A.ҩƷid, A.����, A.����, A.Ч��, A.����id, B.ʵ������ " & _
            " From ҩƷ�շ���¼ A, " & _
            " (Select C.����, C.NO, C.���, C.ҩƷid, Sum(Nvl(C.����, 1) * C.ʵ������) As ʵ������ " & _
            " From ҩƷ�շ���¼ C, ���˷������� A, ���˷��ü�¼ B " & _
            " Where A.����id = B.ID And B.NO = C.NO And B.ID = C.����id And A.״̬ = 0 " & _
            " And C.���� In (9, 10) And C.������� Is Not Null And C.�ⷿid = [1] And Instr([3], ',' || A.�շ�ϸĿid || ',') > 0 " & strTmpCon & _
            " Group By C.����, C.NO, C.���, C.ҩƷid " & _
            " Having Sum(Nvl(C.����, 1) * C.ʵ������) > 0) B" & _
            " Where A.NO = B.NO And A.���� = B.���� And A.ҩƷid + 0 = B.ҩƷid And A.��� = B.��� And A.����� Is Not Null " & _
            " And (A.��¼״̬ = 1 Or Mod(A.��¼״̬, 3) = 0))C, " & _
        " ҩƷ��� D, �շ���ĿĿ¼ X, �շ���Ŀ���� K, ���ű� P, ������ҳ F, ���ű� E,���ű� H " & _
        " Where A.����id = B.ID And B.NO = C.NO And B.ID = C.����id And B.��������id = P.ID And B.�շ�ϸĿid = D.ҩƷid And B.�շ�ϸĿid = X.ID And B.����id = F.����id And B.��ҳid = F.��ҳid  And F.��Ժ���� Is Null And A.���벿��id = E.ID " & strCon & _
        " And X.Id = K.�շ�ϸĿID(+) AND K.����(+)=3  And B.ִ�в���id = [1] And Instr([2], ',' || A.���벿��id || ',') > 0 And A.����� Is Null And A.״̬ = 0 " & _
        " Order By A.����ʱ��, C.����, C.NO, C.��� Desc "
    
    'And Instr([3], ',' || A.�շ�ϸĿid || ',') > 0 " & _
    '" And A.����ʱ�� Between [3] And [4] "
    'Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "��ȡ������ϸ", lngҩ��ID, "," & mstrDrawDept & ",", Dtp���ʿ�ʼʱ��.Value, Dtp���ʽ���ʱ��.Value)
    
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "��ȡ������ϸ", lngҩ��ID, "," & mstrDrawDept & ",", "," & mstrSendDrugId & ",", mstrסԺ��, mstr��������, mlng����ID, mstr����)
    
    If rsTemp.EOF Then
        Exit Function
    End If
    
    Do While Not rsTemp.EOF
        With mrsRequest
            .AddNew
            !ҩƷ���� = rsTemp!ҩƷ����
            !��ҩ���� = rsTemp!��ҩ����
            !��ҩ����id = rsTemp!��ҩ����id
            !���� = rsTemp!����
            !NO = rsTemp!NO
            !ҩƷID = rsTemp!ҩƷID
            !����ʱ�� = Format(rsTemp!����ʱ��, "yyyy-mm-dd hh:mm:ss")
            !�շ���� = rsTemp!�շ����
            !���� = rsTemp!����
            !���� = rsTemp!����
            !Ч�� = rsTemp!Ч��
            
            If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1 And NVL(!Ч��) <> "" Then
                '����Ϊ��Ч��
                !Ч�� = Format(DateAdd("D", -1, !Ч��), "yyyy-mm-dd")
            End If
            
            !׼������ = rsTemp!׼������
            !�������� = rsTemp!��������
            !��װ = rsTemp!��װ
            !��λ = rsTemp!��λ
            !�շ�ID = rsTemp!�շ�ID
            !��ҳid = IIf(IsNull(rsTemp!��ҳid), 0, rsTemp!��ҳid)
            !������� = rsTemp!�������
            !���� = rsTemp!����
            !����id = rsTemp!����id
            !��¼���� = rsTemp!��¼����
            !��˱�־ = 0
            .Update
        End With
        
        With mrsRequestMain
            dbl׼������ = dbl׼������ + rsTemp!׼������
            If lng��ҩ����ID <> rsTemp!��ҩ����id And str����ʱ�� <> Format(rsTemp!����ʱ��, "yyyy-mm-dd hh:mm:ss") And lng����id <> rsTemp!����id Then
                .AddNew
                !��ҩ����id = rsTemp!��ҩ����id
                !ҩƷID = rsTemp!ҩƷID
                !����ʱ�� = Format(rsTemp!����ʱ��, "yyyy-mm-dd hh:mm:ss")
                !����id = rsTemp!����id
                !׼������ = dbl׼������
                !�������� = rsTemp!��������
                
                .Update
                
                dbl׼������ = 0
            End If
            lng��ҩ����ID = rsTemp!��ҩ����id
            str����ʱ�� = Format(rsTemp!����ʱ��, "yyyy-mm-dd hh:mm:ss")
            lngҩƷID = rsTemp!ҩƷID
            lng����id = rsTemp!����id
        End With
        
        rsTemp.MoveNext
    Loop
    
    'ֻ����ҩ�嵥��Ӧ��ҩƷ������ҩ����ID��ҩƷIDΪ׼��
    mrsRequest.MoveFirst
    Do While Not mrsRequest.EOF
        RecBillData.MoveFirst
        Do While Not RecBillData.EOF
            If mrsRequest!��ҩ����id = RecBillData!��ҩ����id And mrsRequest!ҩƷID = RecBillData!ҩƷID Then
                mrsRequest!��˱�־ = 1
                mrsRequest.Update
            End If
            RecBillData.MoveNext
        Loop
        mrsRequest.MoveNext
    Loop
    
    RecBillData.MoveFirst
    mrsRequest.MoveFirst
    
    Call AutoExpendQuantity
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Function LoadDataInBill�����嵥(ByVal lng��ҩ����ID As Integer, ByVal lngҩƷID As Long) As Boolean
    Dim dblSumNum As Double
    
    With mrsRequest
        Call ClearBill(Bill��ҩ����)
        
        .Filter = "��ҩ����id=" & lng��ҩ����ID & " And ҩƷID=" & lngҩƷID & " And ��˱�־ = 1"
        .Sort = "NO,�շ���� Desc"
        
        If .EOF Then Exit Function
        
        Do While Not .EOF
            Bill��ҩ����.TextMatrix(Bill��ҩ����.rows - 1, �����б�.����) = !����
            Bill��ҩ����.TextMatrix(Bill��ҩ����.rows - 1, �����б�.NO) = !NO
            Bill��ҩ����.TextMatrix(Bill��ҩ����.rows - 1, �����б�.ҩƷID) = !ҩƷID
            Bill��ҩ����.TextMatrix(Bill��ҩ����.rows - 1, �����б�.����ʱ��) = Format(!����ʱ��, "yyyy-mm-dd hh:mm:ss")
            Bill��ҩ����.TextMatrix(Bill��ҩ����.rows - 1, �����б�.����) = IIf(IsNull(!����), "", !����)
            Bill��ҩ����.TextMatrix(Bill��ҩ����.rows - 1, �����б�.����) = IIf(IsNull(!����), "", !����)
            Bill��ҩ����.TextMatrix(Bill��ҩ����.rows - 1, �����б�.Ч��) = Format(!Ч��, "yyyy-mm-dd")
            Bill��ҩ����.TextMatrix(Bill��ҩ����.rows - 1, �����б�.׼������) = FormatEx(!׼������ / !��װ, 5)
            Bill��ҩ����.TextMatrix(Bill��ҩ����.rows - 1, �����б�.��������) = FormatEx(!�������� / !��װ, 5)
            Bill��ҩ����.TextMatrix(Bill��ҩ����.rows - 1, �����б�.��װ) = IIf(IsNull(!��װ), "", !��װ)
            Bill��ҩ����.TextMatrix(Bill��ҩ����.rows - 1, �����б�.��λ) = IIf(IsNull(!��λ), "", !��λ)
            Bill��ҩ����.TextMatrix(Bill��ҩ����.rows - 1, �����б�.�շ����) = IIf(IsNull(!�շ����), "", !�շ����)
            Bill��ҩ����.rows = Bill��ҩ����.rows + 1
            
            dblSumNum = dblSumNum + !�������� / !��װ
            
           .MoveNext
        Loop
        
        Bill��ҩ����.TextMatrix(Bill��ҩ����.rows - 1, �����б�.NO) = "�ϼ�"
        Bill��ҩ����.TextMatrix(Bill��ҩ����.rows - 1, �����б�.��������) = FormatEx(dblSumNum, 5)
        
        Bill��ҩ����.Row = Bill��ҩ����.rows - 1
        Bill��ҩ����.Col = �����б�.NO
        Bill��ҩ����.CellForeColor = glng��ҩ
        
        Bill��ҩ����.Col = �����б�.��������
        Bill��ҩ����.CellForeColor = glng��ҩ
    End With
    
    LoadDataInBill�����嵥 = True
End Function
Private Sub AutoExpendQuantity()
    '���ǵ�ͬһ����ID��Ӧ����շ�ID���������Ҫ�����������ֽ⵽����շ���¼��
    '�ֽ��ԭ���ǰ���Ŵ�����ȷ��䣨�Ѱ���Ž�������
    Dim n As Integer
    Dim dbl׼������ As Double
    Dim dblʣ������ As Double
    Dim int�շ���� As Integer
    Dim lng����id As Long
    Dim lngҩƷID As Long
    Dim str����ʱ�� As String
    
    With mrsRequest
        If .RecordCount > 0 Then .MoveFirst
        For n = 1 To .RecordCount
            dbl׼������ = !׼������

            If lng����id = !����id And lngҩƷID = !ҩƷID And str����ʱ�� = !����ʱ�� Then

            Else
                dblʣ������ = !��������
            End If

            If dblʣ������ >= dbl׼������ Then
                dblʣ������ = dblʣ������ - dbl׼������
                !�������� = dbl׼������
            Else
                !�������� = dblʣ������
                dblʣ������ = 0
            End If

            lng����id = !����id
            lngҩƷID = !ҩƷID
            str����ʱ�� = !����ʱ��

            .Update
            .MoveNext
        Next
    End With
    
    '��������������׼�����������־Ϊ�ܾ����
    With mrsRequestMain
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            mrsRequest.Filter = "ҩƷID=" & !ҩƷID & _
                " And ����ID=" & !����id & _
                " And ����ʱ��='" & !����ʱ�� & "'"
            If mrsRequest.RecordCount > 0 Then
                If !׼������ < !�������� Then
                    Do While Not mrsRequest.EOF
                        mrsRequest!��˱�־ = 2
                        mrsRequest.Update
                        mrsRequest.MoveNext
                    Loop
                End If
            End If
            .MoveNext
        Loop
    End With
End Sub
Private Sub ResizeCheckControl()
    '������ҩ����ѡ���λ��
    Dim n As Integer
    Dim dbl����� As Double
    Dim dblTmp As Double
    Dim dblSumTmp As Double
    Dim int���� As Integer
    Const cst������ = 50
    Const cst�о� = 50
    
    If chkSendType.UBound > 0 Then
        dbl����� = fraConNormal.Width - fraTypeLine.Left - 150
        
        int���� = 0
        dblSumTmp = chkSendType(0).Width + cst������
        For n = 1 To chkSendType.UBound
            dblTmp = chkSendType(n).Width + dblSumTmp
            
            If dblTmp <= dbl����� Then
                chkSendType(n).Top = chkSendType(n - 1).Top
                chkSendType(n).Left = chkSendType(n - 1).Left + chkSendType(n - 1).Width + cst������
                dblSumTmp = dblSumTmp + chkSendType(n).Width + cst������
            Else
                '�����У������������ؼ�λ��
                int���� = int���� + 1
                chkSendType(n).Left = chkSendType(0).Left
                chkSendType(n).Top = chkSendType(0).Top + (chkSendType(0).Height + cst�о�) * int����
                dblSumTmp = chkSendType(n).Width + cst������
                
                fraTypeLine.Height = fraTypeLine.Height + chkSendType(0).Height * int����
                
                fraConNormal.Height = fraConNormal.Height + chkSendType(0).Height + cst�о�
                fraCondition.Height = fraCondition.Height + chkSendType(0).Height + cst�о�
                
                frmLine.Top = frmLine.Top + chkSendType(0).Height + cst�о�
                
                fraConExpand.Top = fraConExpand.Top + chkSendType(0).Height + cst�о�
                fraConRequest.Top = fraConRequest.Top + chkSendType(0).Height + cst�о�
                
                If mblnAllConditon = True Then
                    cmdRefresh.Top = cmdRefresh.Top + chkSendType(0).Height + cst�о�
                    cmdOtherCon.Top = cmdRefresh.Top
                End If

                TabShow.Height = TabShow.Height - (chkSendType(0).Height + cst�о�)
            End If
        Next
    End If
End Sub


Private Sub ResizeCondition()
    Dim dblDistance As Double
    Dim n As Integer
    Dim DblHeight As Double, DblWidth As Double
    
    fraCondition.Top = IIf(Cbar.Visible, Cbar.Height, 0)
    fraCondition.Width = Me.ScaleWidth - 20
    fraCondition.Visible = True
    
    fraConExpand.Visible = False
    fraConRequest.Visible = False
    
    frmLine.Visible = False
    frmLine1.Visible = False
    
    fraConNormal.Top = 100
    fraConNormal.Left = 20
    fraConExpand.Left = 20
    fraConRequest.Left = 20
    
    frmLine.Width = fraCondition.Width
    frmLine1.Width = fraCondition.Width
    frmLine.ZOrder 0
    frmLine1.ZOrder 0
    
    cmdRefresh.Top = fraConNormal.Top + (fraConNormal.Height - cmdRefresh.Height) - 50
    cmdOtherCon.Top = cmdRefresh.Top
    cmdRefresh.Left = fraConNormal.Left + fraConNormal.Width + 10
    cmdOtherCon.Left = cmdRefresh.Left + cmdRefresh.Width + 10
    
    If mblnAllConditon = True Then
        fraConExpand.Visible = True
        frmLine.Visible = True

        frmLine.Top = fraConNormal.Top + fraConNormal.Height + 20

        fraConExpand.Top = frmLine.Top - 120
    Else
        fraConExpand.Visible = False
        frmLine.Visible = False
    End If

    If fraConExpand.Visible = True Then
        frmLine1.Top = fraConExpand.Top + fraConExpand.Height
    Else
        frmLine1.Top = fraConNormal.Top + fraConNormal.Height
    End If
    fraCondition.Height = frmLine1.Top + 50
    
    With TabShow
        .Top = IIf(Cbar.Visible, Cbar.Height + fraCondition.Height, fraCondition.Height)
        .Left = 0
        .Height = Me.ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
        .Width = Me.ScaleWidth
    End With
    
    DblHeight = TabShow.Height - TabShow.TabHeight - 120
    DblWidth = TabShow.Width - 150
    With Billδ��ҩ�嵥
        .Height = DblHeight
        .Width = DblWidth
    End With
    With Bill���ܷ�ҩ
        .Height = DblHeight
        .Width = DblWidth
    End With
    With Bill�ܷ�ҩ�嵥
        .Height = DblHeight
        .Width = DblWidth
    End With
    With Billȱҩ�嵥
        .Height = DblHeight
        .Width = DblWidth
    End With
    With Bill�ѷ�ҩ�嵥
        .Height = DblHeight
        .Width = DblWidth
    End With
    
    If Bill��ҩ����.Visible = True Then
        Bill���ܷ�ҩ.Height = DblHeight - Bill��ҩ����.Height - 25
        Bill��ҩ����.Top = Bill���ܷ�ҩ.Top + Bill���ܷ�ҩ.Height + 25
    End If
    
    '������ҩ�˺ͷ�ҩ����ӡ��ʽ
    If TabShow.Tab = 0 Then
        lbl��ҩ��.Top = TabShow.Height - lbl��ҩ��.Height - 120
        cbo��ҩ��.Top = lbl��ҩ��.Top - 60
        Billδ��ҩ�嵥.Height = TabShow.Height - TabShow.TabHeight - 120 - lbl��ҩ��.Height - 150
        
        lbl��ҩ����ʽ.Top = lbl��ҩ��.Top
        cbo��ҩ����ʽ.Top = cbo��ҩ��.Top
        cbo��ҩ����ʽ.Left = TabShow.Width - cbo��ҩ����ʽ.Width - 50
        lbl��ҩ����ʽ.Left = cbo��ҩ����ʽ.Left - 50 - lbl��ҩ����ʽ.Width
    End If
    
    '�����ؼ�����
    Chk�嵥.Top = TabShow.Top + 70
    Chk��ʾ��ҩ��������.Top = Chk�嵥.Top
    cmdAlley.Top = TabShow.Top + 30
End Sub

Private Function NeedName(strList As String) As String
    NeedName = Mid(strList, InStr(strList, "-") + 1)
End Function
Private Sub SaveCondition(ByVal intType As Integer)
    Dim strPath As String
    Dim strBegin As String
    Dim strEnd As String
    
    '�������������ע���
    
    If intType = 0 Then
        strPath = "δ��ҩ�嵥����"
        strBegin = mstr��ʼ����_δ��
        strEnd = mstr��������_δ��
    Else
        strPath = "�ѷ�ҩ�嵥����"
        strBegin = mstr��ʼ����_�ѷ�
        strEnd = mstr��������_�ѷ�
    End If
    
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "���ŷ�ҩ����\" & strPath, "ȫ������", IIf(mblnAllConditon = True, 1, 0)
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "���ŷ�ҩ����\" & strPath, "ʱ�䷶Χ", strBegin & ";" & strEnd
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "���ŷ�ҩ����\" & strPath, "������Ϣ", Val(lblPatiInputType.Tag) & ";" & Trim(txtPati.Text)
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "���ŷ�ҩ����\" & strPath, "����", mint���� & ";" & mstr���� & ";" & mstr��������
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "���ŷ�ҩ����\" & strPath, "��ҩ����", int��Ժ��ҩ & ";" & mstr��ҩ����
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "���ŷ�ҩ����\" & strPath, "��ҩ;��", mstrUse
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "���ŷ�ҩ����\" & strPath, "ҩƷ����", mstrDrug
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "���ŷ�ҩ����\" & strPath, "����Χ", mint��Χ
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "���ŷ�ҩ����\" & strPath, "ҽ������", Lngҽ������
End Sub


Private Sub ClearCondition(ByVal intType As Integer)
    Dim strPath As String
    Dim strBegin As String
    Dim strEnd As String
    
    '���������������Ҫ�����Ҫ���沿����������ɾ�������
    On Error Resume Next
    
    DeleteSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "���ŷ�ҩ����\δ��ҩ�嵥����"
    DeleteSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "���ŷ�ҩ����\�ѷ�ҩ�嵥����"
    
    '�������������ע���
    If intType = 0 Then
        strPath = "δ��ҩ�嵥����"
        strBegin = mstr��ʼ����_δ��
        strEnd = mstr��������_δ��
    Else
        strPath = "�ѷ�ҩ�嵥����"
        strBegin = mstr��ʼ����_�ѷ�
        strEnd = mstr��������_�ѷ�
    End If
    
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "���ŷ�ҩ����\" & strPath, "��ҩ����", int��Ժ��ҩ & ";" & mstr��ҩ����
End Sub
Private Sub SetCondition(ByVal intType As Integer)
    Dim n As Integer
    
    'ʱ�䷶Χ
    If intType = 1 Then
        mstr��ʼ����_�ѷ� = Format(Dtp��ʼʱ��.Value, "yyyy-MM-dd hh:mm:ss")
        mstr��������_�ѷ� = Format(Dtp����ʱ��.Value, "yyyy-MM-dd hh:mm:ss")
    Else
        mstr��ʼ����_δ�� = Format(Dtp��ʼʱ��.Value, "yyyy-MM-dd hh:mm:ss")
        mstr��������_δ�� = Format(Dtp����ʱ��.Value, "yyyy-MM-dd hh:mm:ss")
    End If
    
    '������Ϣ
    mstrסԺ�� = ""
    mstr�������� = ""
    mstr���� = ""
    mstrSerchNO = ""
    mlng����ID = 0
        
    If Trim(txtPati.Text) <> "" Then
        Select Case Val(lblPatiInputType.Tag)
            Case PatiInfo.סԺ��
                If InStr(txtPati.Text, "-") > 0 Then
                    mstrסԺ�� = Mid(Trim(txtPati.Text), 1, InStr(txtPati.Text, "-") - 1)
                Else
                    mstrסԺ�� = Trim(txtPati.Text)
                End If
            Case PatiInfo.����
                If mblnCard = True Then
                    mlng����ID = Val(txtPati.Tag)
                Else
                    mstr�������� = Trim(txtPati.Text)
                End If
            Case PatiInfo.����
                If InStr(txtPati.Text, "-") > 0 Then
                    mstr���� = Mid(Trim(txtPati.Text), 1, InStr(txtPati.Text, "-") - 1)
                Else
                    mstr���� = Trim(txtPati.Text)
                End If
            Case PatiInfo.���ݺ�
                If InStr(txtPati.Text, "-") > 0 Then
                    mstrSerchNO = Mid(Trim(txtPati.Text), 1, InStr(txtPati.Text, "-") - 1)
                Else
                    mstrSerchNO = Trim(txtPati.Text)
                End If
            Case PatiInfo.����ID
                If InStr(txtPati.Text, "-") > 0 Then
                    mlng����ID = Mid(Trim(txtPati.Text), 1, InStr(txtPati.Text, "-") - 1)
                Else
                    mlng����ID = Val(Trim(txtPati.Text))
                End If
            Case PatiInfo.���￨
                mlng����ID = Val(txtPati.Tag)
        End Select
    End If
    
    '�������ͺͲ���ID
    mstr���� = ""
    mstr�������� = ""
    mint���� = (tbsType.SelectedItem.Index - 1)
    If Trim(txt����.Text) <> "" Then
        mstr���� = txt����.Tag
        mstr�������� = txt����.Text
    End If
        
    '����
    If Trim(txtҩƷ����.Text) = "" Or InStr(Trim(txtҩƷ����.Text), "����ҩƷ����") > 0 Then
        mstrDrug = ""
    Else
        mstrDrug = Trim(txtҩƷ����.Text)
    End If
    
    '��ҩ;��
    If Trim(txt��ҩ;��.Text) = "" Or InStr(Trim(txt��ҩ;��.Text), "���и�ҩ;��") > 0 Then
        mstrUse = ""
    Else
        mstrUse = Trim(txt��ҩ;��.Text)
    End If
    
    '����Χ
    If Me.opt��Χ(1).Value = True Then
        mint��Χ = 1
    ElseIf Me.opt��Χ(2).Value = True Then
        mint��Χ = 2
    Else
        mint��Χ = 0
    End If
    
    '��ҩ����
    '0-����,1-������Ժ��ҩ,2-������Ժ��ҩ,3-������ȡҩ,4-������ȡҩ,5-Ժ����ҩ(��������Ժ��ҩ����ȡҩ),6-��Ժ��ҩ����ȡҩ
    If chkSend(0).Value = 1 And chkSend(1).Value = 1 And chkSend(2).Value = 1 Then
        int��Ժ��ҩ = 0
    ElseIf chkSend(0).Value = 1 And chkSend(2).Value = 1 Then
        int��Ժ��ҩ = 1
    ElseIf chkSend(0).Value = 1 And chkSend(1).Value = 1 Then
        int��Ժ��ҩ = 3
    ElseIf chkSend(1).Value = 1 And chkSend(2).Value = 1 Then
        int��Ժ��ҩ = 6
    ElseIf chkSend(0).Value = 1 Then
        int��Ժ��ҩ = 5
    ElseIf chkSend(1).Value = 1 Then
        int��Ժ��ҩ = 2
    ElseIf chkSend(2).Value = 1 Then
        int��Ժ��ҩ = 4
    End If
    
    mstr��ҩ���� = ""
    If chkSendType(0).Visible = True Then
        For n = 0 To chkSendType.UBound
            If chkSendType(n).Value = 1 Then
                mstr��ҩ���� = IIf(mstr��ҩ���� = "", "", mstr��ҩ���� & ",") & chkSendType(n).Caption
            End If
        Next
    End If
    
    'ҽ������
    Lngҽ������ = Cboҽ������.ListIndex
    
    '��������
    If chkType(0).Value = 1 And chkType(1).Value = 1 Then
        mint�������� = 2
    ElseIf chkType(1).Value = 1 Then
        mint�������� = 1
    Else
        mint�������� = 0
    End If
    
End Sub

Private Sub SetGroup(ByVal Bill As MSHFlexGrid, ByVal bln�Ƿ���� As Boolean)
    Dim n As Integer
    Dim lng�������ID As Long
    Dim lng�������ID As Long
    Dim lng�������ID As Long
    Dim int����_���ID As Integer
    Dim int����_����� As Integer
    Dim bln�Ƿ���ڷ��� As Boolean
    Dim bln�����з���� As Boolean
    
    '�Ʊ������ �� ��
    
    '������С������ʱû�б�Ҫ����
    If Bill.rows < 3 Then Exit Sub
    
    lng�������ID = -1
        
    '�����ID����
    With Bill
        Select Case .Name
        Case "Billδ��ҩ�嵥"
            int����_���ID = ����_δ��ҩ�嵥.���ID
            int����_����� = ����_δ��ҩ�嵥.�����
        Case "Bill�ѷ�ҩ�嵥"
            int����_���ID = ����_�ѷ�ҩ�嵥.���ID
            int����_����� = ����_�ѷ�ҩ�嵥.�����
        End Select
        
        .Redraw = False
        For n = 1 To .rows - 1
             .TextMatrix(n, int����_�����) = ""
             .RowHeight(n) = 220
        Next
                
        If Not bln�Ƿ���� Then
            .ColWidth(int����_�����) = 0
            .Redraw = True
            Exit Sub
        Else
            .ColWidth(int����_�����) = 250
        End If
        
        For n = 1 To .rows - 1
            .Row = n
            .Col = int����_�����
            If .TextMatrix(n, int����_���ID) <> "" Then
                lng�������ID = .TextMatrix(n, int����_���ID)
                If n + 1 <= .rows - 1 Then
                    If .TextMatrix(n + 1, int����_���ID) <> "" Then    '�������Ϊ��¼��ʱ
                        lng�������ID = IIf(.TextMatrix(n + 1, int����_���ID) = 0, -1, .TextMatrix(n + 1, int����_���ID))
                    ElseIf n + 2 <= .rows - 1 Then  '�������Ϊ��������ʱ
                        If .TextMatrix(n + 2, int����_���ID) <> "" Then    '���������Ϊ��¼��ʱ
                            lng�������ID = IIf(.TextMatrix(n + 2, int����_���ID) = 0, -1, .TextMatrix(n + 2, int����_���ID))
                        Else
                            lng�������ID = -1
                        End If
                    Else
                        lng�������ID = -1
                    End If
                Else
                    lng�������ID = -1
                End If
                
                If lng�������ID = lng�������ID Then
                    If lng�������ID = lng�������ID Then
                        .TextMatrix(n, int����_�����) = "��"
                        .RowHeight(n) = 220
                    Else
                        .TextMatrix(n, int����_�����) = "��"
                        .CellAlignment = flexAlignLeftTop
                    End If
                ElseIf lng�������ID = lng�������ID Then
                    .TextMatrix(n, int����_�����) = "��"
                    .CellAlignment = flexAlignLeftBottom
                    bln�Ƿ���ڷ��� = True
                End If
            
                lng�������ID = IIf(lng�������ID = 0, -1, lng�������ID)
            Else
                '��������ǻ����У���Ҫ�������е����ID�жϷ������
                If n + 1 <= .rows - 1 Then
                    If .TextMatrix(n + 1, int����_���ID) <> "" Then
                        If lng�������ID <> -1 And lng�������ID = IIf(.TextMatrix(n + 1, int����_���ID) = 0, -1, .TextMatrix(n + 1, int����_���ID)) Then
                            .TextMatrix(n, int����_�����) = "��"
                            .RowHeight(n) = 220
                        End If
                    End If
                End If
            End If
        Next
        
        If Not bln�Ƿ���ڷ��� Then .ColWidth(int����_�����) = 0
        
        .Redraw = True

    End With
    
End Sub

Private Sub Bill��ҩ����_EnterCell()
    Call SetSelectColor(Bill��ҩ����)
End Sub


Private Sub Bill��ҩ����_GotFocus()
    Call Bill��ҩ����_EnterCell
End Sub
Private Sub Billδ��ҩ�嵥_Scroll()
    Cbo����.Visible = False
End Sub

Private Sub Bill�ѷ�ҩ�嵥_Scroll()
    TxtInput.Visible = False
End Sub

Private Sub Cbo��ҩҩ��_Click()
    lngҩ��ID = Cbo��ҩҩ��.ItemData(Cbo��ҩҩ��.ListIndex)
    
    str������ = "���м�����"

    If lngҩ��ID <> Val(Cbo��ҩҩ��.Tag) Then
        Cbo��ҩҩ��.Tag = lngҩ��ID
        strUnit = GetSpecUnit(lngҩ��ID, gintסԺҩ��)
        mbln�Ƿ��������� = CheckIsCenter(lngҩ��ID)
        
        Call Get��ҩ��
        Call Get����
        
        DoEvents
        
        Call mnuViewRefresh_Click
    End If

End Sub
Private Sub cbo��ҩ��_Click()
'    Exit Sub
End Sub

Private Sub cbo��ҩ��_KeyDown(KeyCode As Integer, Shift As Integer)
    mblnDrop = False
    If KeyCode = 13 Then mblnDrop = SendMessage(cbo��ҩ��.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 1
End Sub

Private Sub cbo��ҩ��_KeyPress(KeyAscii As Integer)
    Dim i As Long, intIdx As Integer
    Dim strText As String, strResult As String, strFilter As String

    If KeyAscii = 13 Then
        strText = UCase(cbo��ҩ��.Text)
        If cbo��ҩ��.ListIndex <> -1 Then
            '�����б�ʱ,�����ı�������������
            If strText <> cbo��ҩ��.List(cbo��ҩ��.ListIndex) Then Call zlControl.CboSetIndex(cbo��ҩ��.hWnd, -1)
        End If
        If strText = "" Then
            cbo��ҩ��.ListIndex = -1
        ElseIf cbo��ҩ��.ListIndex = -1 Then
            intIdx = -1

            For i = 1 To cbo��ҩ��.ListCount - 1
                If Mid(cbo��ҩ��.List(i), 1, InStr(1, cbo��ҩ��.List(i), "-") - 1) = strText _
                    Or Mid(cbo��ҩ��.List(i), InStr(1, cbo��ҩ��.List(i), "-")) = strText Then
                    intIdx = i
                    Exit For
                End If
            Next

            If intIdx = -1 Then
                For i = 1 To cbo��ҩ��.ListCount - 1
                    If UCase(cbo��ҩ��.List(i)) Like strText & "*" Then
                        intIdx = i
                    End If
                Next
            End If

            cbo��ҩ��.ListIndex = intIdx
            SendMessage cbo��ҩ��.hWnd, CB_SHOWDROPDOWN, True, 0
        ElseIf Not mblnDrop Then
            '�س���꾭��
            Call cbo��ҩ��_Click
            Exit Sub
        End If
        If cbo��ҩ��.ListIndex = -1 Then
            cbo��ҩ��.ListIndex = 0
        Else
            If intIdx <> -1 And mblnDrop Then
                '�����س�-ǿ�м���Click
                Call cbo��ҩ��_Click
            ElseIf intIdx <> cbo��ҩ��.ListIndex And intIdx <> -1 Then
                '������ѡ��-�Զ�����Click
                cbo��ҩ��.SetFocus
                Exit Sub
            ElseIf intIdx <> -1 Then
                'һ��������-ǿ�м���Click
                Call cbo��ҩ��_Click
            End If
        End If
    End If
End Sub

Private Sub chkSend_Click(Index As Integer)
    Dim i As Integer
    Dim blnAllUnCheck As Boolean
    
    If chkSend(Index).Value = 0 Then
        blnAllUnCheck = True
        For i = 0 To chkSend.Count - 1
            If chkSend(i).Value = 1 Then
                blnAllUnCheck = False
                Exit For
            End If
        Next
        If blnAllUnCheck = True Then chkSend(Index).Value = 1
    End If
End Sub

Private Sub chkType_Click(Index As Integer)
    If Index = 0 Then
        If chkType(1).Value <> 1 Then chkType(0).Value = 1
    Else
        If chkType(0).Value <> 1 Then chkType(1).Value = 1
    End If
End Sub
Private Sub Chk��ʾ��ҩ��������_Click()
    mlng�������� = Chk��ʾ��ҩ��������.Value
    Call mnuViewRefresh_Click
End Sub

Private Sub cmdOtherCon_Click()
    If cmdOtherCon.Caption = "ȫ������(&C)" Then
        cmdOtherCon.Caption = "��Ҫ����(&C)"
        mblnAllConditon = True
    Else
        cmdOtherCon.Caption = "ȫ������(&C)"
        mblnAllConditon = False
    End If
    
    Call ResizeCondition
End Sub
Private Sub cmdRefresh_Click()
    ''''ˢ��
    BlnInRefresh = False
    
    If TabShow.Tab = 4 Then
        mblnFirstSended = False
    End If
    
    Call mnuViewRefresh_Click
End Sub

Private Sub cmd��������_Click()
    Dim rsTemp As adodb.Recordset
    Dim rsCount As adodb.Recordset
    Dim str����id() As String
    Dim n As Integer
    Dim i As Integer
    Dim strCond��ҩ���� As String
    
    On Error GoTo errHandle
    If Me.Lvw����.Tag <> "" Then
        If Me.Lvw����.Tag <> tbsType.TabIndex Then
            Me.txt����.Tag = ""
            Me.txt����.Text = ""
        End If
    End If
    
    If TabShow.Tab = 4 Then
        If tbsType.SelectedItem.Index - 1 = 0 Then
            gstrSQL = " Select ����||'-'||���� ����,ID From ���ű� " & _
                     " Where (վ�� = '" & gstrNodeNo & "' Or վ�� is Null) And ID in (Select ����ID From ��������˵�� Where ��������='�ٴ�' And ������� IN(2,3))" & _
                     " And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','yyyy-MM-dd')) " & _
                     " Order By ����||'-'||���� "
        ElseIf tbsType.SelectedItem.Index - 1 = 1 Then
            gstrSQL = " Select ����||'-'||���� ����,ID From ���ű� " & _
                     " Where (վ�� = '" & gstrNodeNo & "' Or վ�� is Null) And ID in (Select ����ID From ��������˵�� Where �������� In ('���','����','����','����','Ӫ��') And ������� IN(2,3))" & _
                     " And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','yyyy-MM-dd')) " & _
                     " Order By ����||'-'||���� "
        Else
            gstrSQL = " Select ����||'-'||���� ����,ID From ���ű� " & _
                     " Where (վ�� = '" & gstrNodeNo & "' Or վ�� is Null) And ID in (Select ����ID From ��������˵�� Where ��������='����' And ������� IN(2,3))" & _
                     " And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','yyyy-MM-dd')) " & _
                     " Order By ����||'-'||���� "
        End If
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "ȡ���ſ���")
        
                
        With rsTemp
            If .EOF Then
                MsgBox "û�����ø��ಿ�ţ������Ź���", vbInformation, gstrSysName
                Exit Sub
            End If
            Me.Lvw����.ListItems.Clear
            Me.Lvw����.Tag = tbsType.TabIndex
            Do While Not .EOF
                Me.Lvw����.ListItems.Add , "_" & !Id, !����, 1, 1
                .MoveNext
            Loop
        End With
    Else
        If tbsType.SelectedItem.Index - 1 = 0 Then
            gstrSQL = "Select Distinct A.���� || '-' || A.���� ����, A.ID " & _
                " From ���ű� A, ��������˵�� B, δ��ҩƷ��¼ C, ���˷��ü�¼ D " & _
                " Where (A.վ�� = '" & gstrNodeNo & "' Or A.վ�� Is Null) And B.�������� ='�ٴ�' And B.������� In (2, 3) And A.ID = B.����id And " & _
                " (A.����ʱ�� Is Null Or A.����ʱ�� = To_Date('3000-01-01', 'yyyy-MM-dd')) And C.�ⷿid = [1] And C.���� In (9,10) And " & _
                " C.�������� Between [2] And [3] And C.NO = D.NO And C.�ⷿid = D.ִ�в���id And A.ID = D.��������id And D.���˿���id = D.��������id " & _
                " Order By A.���� || '-' || A.���� "
        ElseIf tbsType.SelectedItem.Index - 1 = 1 Then
            gstrSQL = "Select Distinct A.���� || '-' || A.���� ����, A.ID " & _
                " From ���ű� A, ��������˵�� B, δ��ҩƷ��¼ C, ���˷��ü�¼ D " & _
                " Where (A.վ�� = '" & gstrNodeNo & "' Or A.վ�� Is Null) And B.�������� In ('���','����','����','����','Ӫ��') And B.������� In (2, 3) And A.ID = B.����id And " & _
                " (A.����ʱ�� Is Null Or A.����ʱ�� = To_Date('3000-01-01', 'yyyy-MM-dd')) And C.�ⷿid = [1] And C.���� In (9,10) And " & _
                " C.�������� Between [2] And [3] And C.NO = D.NO And C.�ⷿid = D.ִ�в���id And A.ID = D.��������id And D.���˿���id <> D.��������id " & _
                " Order By A.���� || '-' || A.���� "
        Else
            gstrSQL = "Select Distinct A.���� || '-' || A.���� ����, A.ID " & _
                " From ���ű� A, ��������˵�� B, δ��ҩƷ��¼ C, ���˷��ü�¼ D " & _
                " Where (A.վ�� = '" & gstrNodeNo & "' Or A.վ�� Is Null) And B.�������� = '����' And B.������� In (2, 3) And A.ID = B.����id And " & _
                " (A.����ʱ�� Is Null Or A.����ʱ�� = To_Date('3000-01-01', 'yyyy-MM-dd')) And C.�ⷿid = [1] And C.���� In (9,10) And " & _
                " C.�������� Between [2] And [3] And C.NO = D.NO And C.�ⷿid = D.ִ�в���id And A.ID = D.���˲���id "
                
            If mstr������ҩ��ʽ = "" Then
                gstrSQL = gstrSQL & " And D.���˿���id = D.��������id "
            End If
            
            gstrSQL = gstrSQL & " Order By A.���� || '-' || A.���� "
        End If
        
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "ȡ���ſ���", lngҩ��ID, CDate(Format(Dtp��ʼʱ��.Value, "yyyy-MM-dd hh:mm:ss")), CDate(Format(Dtp����ʱ��.Value, "yyyy-MM-dd hh:mm:ss")))
        
        With rsTemp
            If .EOF Then
                Exit Sub
            End If
            Me.Lvw����.ListItems.Clear
            Me.Lvw����.Tag = tbsType.TabIndex
            
            Call SetCondition(IIf(TabShow.Tab = 4, 1, 0))
            '��Ժ��ҩ:'0-����,1-������Ժ��ҩ,2-������Ժ��ҩ,3-������ȡҩ,4-������ȡҩ,5-Ժ����ҩ(��������Ժ��ҩ����ȡҩ),6-��Ժ��ҩ����ȡҩ
            If int��Ժ��ҩ = 0 Then
                strCond��ҩ���� = ""
            ElseIf int��Ժ��ҩ = 1 Then
                strCond��ҩ���� = " And Not Ltrim(To_Char(Nvl(A.����,0),'00')) Like '_3'"
            ElseIf int��Ժ��ҩ = 2 Then
                strCond��ҩ���� = " And Ltrim(To_Char(Nvl(A.����,0),'00')) Like '_3'"
            ElseIf int��Ժ��ҩ = 3 Then
                strCond��ҩ���� = " And Not Ltrim(To_Char(Nvl(A.����,0),'00')) Like '_4'"
            ElseIf int��Ժ��ҩ = 4 Then
                strCond��ҩ���� = " And Ltrim(To_Char(Nvl(A.����,0),'00')) Like '_4'"
            ElseIf int��Ժ��ҩ = 5 Then
                strCond��ҩ���� = " And Not Ltrim(To_Char(Nvl(A.����,0),'00')) Like '_3' And Not Ltrim(To_Char(Nvl(A.����,0),'00')) Like '_4'"
            ElseIf int��Ժ��ҩ = 6 Then
                strCond��ҩ���� = " And (Ltrim(To_Char(Nvl(A.����,0),'00')) Like '_3' Or Ltrim(To_Char(Nvl(A.����,0),'00')) Like '_4')"
            End If
            
'            IIf(mstr��ҩ���� = "", "", " And Instr([15],',' || D.��ҩ���� || ',') > 0")
            If mstr��ҩ���� <> "" Then mstr��ҩ���� = "," & mstr��ҩ���� & ","
                
            Do While Not .EOF
                gstrSQL = "Select Count(Distinct A.ҩƷid) As ҩƷ " & _
                    " From ҩƷ�շ���¼ A, δ��ҩƷ��¼ B, ���˷��ü�¼ C " & IIf(mstr��ҩ���� = "", "", " ,ҩƷ��� D") & _
                    " Where A.���� = B.���� And A.NO = B.NO And A.����� Is Null And A.NO = C.NO And B.�ⷿid = C.ִ�в���id " & strCond��ҩ���� & IIf(mstr��ҩ���� = "", "", " And A.ҩƷID = D.ҩƷID And Instr([5],',' || D.��ҩ���� || ',') > 0") & _
                    " And B.�ⷿid = [2] And B.���� In (9,10) And B.�������� Between [3] And [4] "
                    
                If tbsType.SelectedItem.Index - 1 = 0 Then
                    gstrSQL = gstrSQL & " And C.��������id = [1] And C.���˿���id=C.��������id "
                ElseIf tbsType.SelectedItem.Index - 1 = 1 Then
                    gstrSQL = gstrSQL & " And C.��������id = [1] And C.���˿���id<>C.��������id "
                Else
                    If mstr������ҩ��ʽ = "" Then
                        gstrSQL = gstrSQL & " And C.���˲���id = [1] And C.���˿���id=C.��������id "
                    Else
                        gstrSQL = gstrSQL & " And C.���˲���id = [1] "
                    End If
                End If
                
                Set rsCount = zldatabase.OpenSQLRecord(gstrSQL, "ȡ���ſ���", CLng(!Id), lngҩ��ID, CDate(Format(Dtp��ʼʱ��.Value, "yyyy-MM-dd hh:mm:ss")), CDate(Format(Dtp����ʱ��.Value, "yyyy-MM-dd hh:mm:ss")), mstr��ҩ����)
                
                Me.Lvw����.ListItems.Add , "_" & !Id, !���� & "(" & rsCount!ҩƷ & "��ҩƷ������", 1, 1
                .MoveNext
            Loop
        End With
    End If
    
    Lvw����.Move fraCondition.Left + fraConNormal.Left + txt����.Left - 10, fraCondition.Top + txt����.Top + txt����.Height + 60, txt����.Width, 4000
    Lvw����.Visible = True
    Lvw����.SetFocus
    Lvw����.ZOrder 0
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

    
Private Sub cmd��ҩ;��_Click()
    Lvw��ҩ;��.Move fraCondition.Left + fraConExpand.Left + txt��ҩ;��.Left - 10, fraCondition.Top + fraConExpand.Top + txt��ҩ;��.Top + txt��ҩ;��.Height + 60, txt��ҩ;��.Width, 3000
    Lvw��ҩ;��.Visible = True
    Lvw��ҩ;��.SetFocus
    Lvw��ҩ;��.ZOrder 0
End Sub


Private Sub cmdҩƷ����_Click()
    Lvw����.Move fraCondition.Left + fraConExpand.Left + txtҩƷ����.Left - 10, fraCondition.Top + fraConExpand.Top + txtҩƷ����.Top + txtҩƷ����.Height + 60, txtҩƷ����.Width, 3000
    Lvw����.Visible = True
    Lvw����.SetFocus
    Lvw����.ZOrder 0
End Sub

Private Sub lblPatiInputType_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        PopupMenu mnuPatiInfo, 2, fraCondition.Left + fraConNormal.Left + lblPatiInputType.Left - 30, fraCondition.Top + fraConNormal.Top + lblPatiInputType.Top + lblPatiInputType.Height + 30
    End If
End Sub
Private Sub Lvw��ҩ;��_DblClick()
    Dim n As Integer
    
    With Lvw��ҩ;��
        If .SelectedItem Is Nothing Then Exit Sub
        Me.txt��ҩ;��.Tag = ""
        Me.txt��ҩ;��.Text = ""
        
        '���ѡ����ȫѡ������ȡ���и�ҩ;����
        If .ListItems(1).Checked Then
            Me.txt��ҩ;��.Tag = ""
            Me.txt��ҩ;��.Text = "���и�ҩ;��"
            .Visible = False
            Exit Sub
        End If
        For n = 1 To .ListItems.Count
            If .ListItems(n).Checked Then
                Me.txt��ҩ;��.Tag = IIf(Me.txt��ҩ;��.Tag = "", Mid(.ListItems(n).Key, 2), Me.txt��ҩ;��.Tag & "," & Mid(.ListItems(n).Key, 2))
                Me.txt��ҩ;��.Text = IIf(Me.txt��ҩ;��.Text = "", .ListItems(n).Text, Me.txt��ҩ;��.Text & "," & .ListItems(n).Text)
            End If
        Next
    
        '�����ǰ˫���ĸ�ҩ;��δ��ѡ�ϣ�����ǰ˫���ĸ�ҩ;��Ҳ���뵽�༭����
        If .SelectedItem.Checked = False Then
            .SelectedItem.Checked = True
            Me.txt��ҩ;��.Tag = IIf(Me.txt��ҩ;��.Tag = "", Mid(.SelectedItem.Key, 2), Me.txt��ҩ;��.Tag & "," & Mid(.SelectedItem.Key, 2))
            Me.txt��ҩ;��.Text = IIf(Me.txt��ҩ;��.Text = "", .SelectedItem.Text, Me.txt��ҩ;��.Text & "," & .SelectedItem.Text)
        End If
        .Visible = False
    End With
End Sub

Private Sub Lvw��ҩ;��_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim n As Integer
    Dim blnAllChecked As Boolean
    
    With Lvw��ҩ;��
        For n = 1 To .ListItems.Count
            .ListItems(n).Selected = False
        Next
        Item.Selected = True
        If Item.Text = "���и�ҩ;��" Then
            If Item.Checked Then
                blnAllChecked = True
            End If
                
            For n = 1 To .ListItems.Count
                .ListItems(n).Checked = blnAllChecked
            Next
        Else
            If Item.Checked = False Then
                .ListItems(1).Checked = False
            End If
        End If
    End With
End Sub

Private Sub Lvw��ҩ;��_LostFocus()
    Lvw��ҩ;��.Visible = False
End Sub

Private Sub Lvw��ҩ;��_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        If mnuTypeItem.Item(0).Caption <> "-" Then
            PopupMenu mnuType, 2
        End If
    End If
End Sub

Private Sub Lvw����_DblClick()
    Dim n As Integer
    
    With Lvw����
        If .SelectedItem Is Nothing Then Exit Sub
        Me.txtҩƷ����.Text = ""
        
        '���ѡ����ȫѡ������ȡ���и�ҩ;����
        If .ListItems(1).Checked Then
             Me.txtҩƷ����.Text = "����ҩƷ����"
            .Visible = False
            Exit Sub
        End If
        For n = 1 To .ListItems.Count
            If .ListItems(n).Checked Then
                Me.txtҩƷ����.Text = IIf(Me.txtҩƷ����.Text = "", Mid(.ListItems(n).Text, InStr(1, .ListItems(n).Text, "-") + 1), Me.txtҩƷ����.Text & "," & Mid(.ListItems(n).Text, InStr(1, .ListItems(n).Text, "-") + 1))
            End If
        Next
    
        '�����ǰ˫���ĸ�ҩ;��δ��ѡ�ϣ�����ǰ˫���ĸ�ҩ;��Ҳ���뵽�༭����
        If .SelectedItem.Checked = False Then
            .SelectedItem.Checked = True
            Me.txtҩƷ����.Text = IIf(Me.txtҩƷ����.Text = "", Mid(.SelectedItem.Text, InStr(1, .SelectedItem.Text, "-") + 1), Me.txtҩƷ����.Text & "," & Mid(.SelectedItem.Text, InStr(1, .SelectedItem.Text, "-") + 1))
        End If
        .Visible = False
    End With
End Sub


Private Sub Lvw����_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim n As Integer
    Dim blnAllChecked As Boolean
    
    With Lvw����
        For n = 1 To .ListItems.Count
            .ListItems(n).Selected = False
        Next
        Item.Selected = True
        If Item.Text = "����ҩƷ����" Then
            If Item.Checked Then
                blnAllChecked = True
            End If
                
            For n = 1 To .ListItems.Count
                .ListItems(n).Checked = blnAllChecked
            Next
        Else
            If Item.Checked = False Then
                .ListItems(1).Checked = False
            End If
        End If
    End With
End Sub


Private Sub Lvw����_LostFocus()
    Lvw����.Visible = False
End Sub
Private Sub Lvw����_DblClick()
    Dim n As Integer
    
    With Me.Lvw����
        If .SelectedItem Is Nothing Then Exit Sub
        Me.txt����.Tag = ""
        Me.txt����.Text = ""
        For n = 1 To .ListItems.Count
            If .ListItems(n).Checked Then
                Me.txt����.Tag = IIf(Me.txt����.Tag = "", Mid(.ListItems(n).Key, 2), Me.txt����.Tag & "," & Mid(.ListItems(n).Key, 2))
                Me.txt����.Text = IIf(Me.txt����.Text = "", .ListItems(n).Text, Me.txt����.Text & "," & .ListItems(n).Text)
            End If
        Next
    
        '�����ǰ˫���Ŀ���δ��ѡ�ϣ�����ǰ˫���Ŀ���Ҳ���뵽�Է����ұ༭����
        If .SelectedItem.Checked = False Then
            .SelectedItem.Checked = True
            Me.txt����.Tag = IIf(Me.txt����.Tag = "", Mid(.SelectedItem.Key, 2), Me.txt����.Tag & "," & Mid(.SelectedItem.Key, 2))
            Me.txt����.Text = IIf(Me.txt����.Text = "", .SelectedItem.Text, Me.txt����.Text & "," & .SelectedItem.Text)
        End If
        .Visible = False
        txt����.SetFocus
    End With
End Sub





Private Sub Lvw����_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim n As Integer
    
    For n = 1 To Lvw����.ListItems.Count
        Lvw����.ListItems(n).Selected = False
    Next
    
    Item.Selected = True
End Sub


Private Sub Lvw����_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn, vbKeySpace
        If Me.Lvw����.SelectedItem Is Nothing Then Exit Sub
        Call Lvw����_DblClick
    End Select
End Sub


Private Sub Lvw����_LostFocus()
    Me.Lvw����.Visible = False
    txt����.SetFocus
End Sub
Private Sub mnuBillItem_Click(Index As Integer)
    mnuBillItem(Index).Checked = Not mnuBillItem(Index).Checked
    
    If (Me.mnuBillItem(Index).Caption = "��/��ҩ��" Or Me.mnuBillItem(Index).Caption = "��ҩ��") Then
        mbln��ʾ����ҩ�� = (Me.mnuBillItem(Index).Checked)
    End If
                        
    Call SetColHideByMenu(mnuBillItem(Index), IIf(TabShow.Tab = 0, Billδ��ҩ�嵥, Bill�ѷ�ҩ�嵥))
End Sub
Private Sub mnuDrugCodeName_Click(Index As Integer)
    Dim n As Integer
    Dim strSave As String
    
    If mnuDrugCodeName(Index).Checked = True Then Exit Sub
    
    For n = 0 To mnuDrugCodeName.Count - 1
        mnuDrugCodeName(n).Checked = False
    Next
    
    mnuDrugCodeName(Index).Checked = True
    
    '��������
    intҩƷ���� = Index
    
    strSave = Index & "|" & "ҩƷ����"
    
    For n = 0 To mnuBillItem.Count - 1
        strSave = strSave & "," & IIf(mnuBillItem(n).Checked, "0", "1") & "|" & mnuBillItem(n).Caption
    Next
    
    zldatabase.SetPara "������", strSave, glngSys, 1342
    
    '��������
    mnuViewRefresh_Click
End Sub

Private Sub mnuInfoItem_Click(Index As Integer)
    Dim strItem As String, i As Long
    
    For i = 0 To mnuInfoItem.UBound
        mnuInfoItem(i).Checked = (i = Index)
    Next
    
    strItem = Split(mnuInfoItem(Index).Caption, "(")(0)
    lblPatiInputType.Caption = strItem & "��"
    lblPatiInputType.Tag = Index
    
    txtPati.Text = ""
    txtPati.PasswordChar = ""
    txtPati.MaxLength = 0
    
    If Val(lblPatiInputType.Tag) = PatiInfo.���￨ Then
        If gtype_UserSysParms.P12_���￨�Ƿ�������ʾ Then
            txtPati.PasswordChar = "*"
        End If
        txtPati.MaxLength = gtype_UserSysParms.P20_���￨�ų���
    End If
    
'    txtPati.SetFocus
    
End Sub
Private Sub mnuPassItem_Click(Index As Integer)
    '���ܣ�ִ��PASS����
    'Pass
    Select Case Index
    Case 0 'ҩ���ٴ���Ϣ�ο�
        Call PassDoCommand(101)
    Case 1 'ҩƷ˵����
        Call PassDoCommand(102)
    Case 2 '�й�ҩ��
        Call PassDoCommand(107)
    Case 3 '������ҩ����
        Call PassDoCommand(103)
    Case 4 '����ֵ
        Call PassDoCommand(104)
    Case 8 'ҽҩ��Ϣ����
        Call PassDoCommand(106)
    Case 10 'ҩƷ�����Ϣ
        Call PassDoCommand(13)
    Case 11 '��ҩ;�������Ϣ
        Call PassDoCommand(14)
    Case 12 'ҽԺҩƷ��Ϣ
        Call PassDoCommand(105)
    End Select
End Sub

Private Function AdviceCheckWarn(ByVal lngCmd As Long, Optional ByVal lngRow As Long) As Long
'���ܣ�����Passϵͳ��ع���
'������lngCmd=
'        0-�������PASS�˵�״̬
'        21-����״̬/����ʷ����(ֻ��)
'      lngRow=��ǰҩƷҽ�����кţ�lngCmd=0ʱ��Ҫ
'���أ����PASS�˵�ʱ������>=0��ʾ���Ե����˵�,��������-1
'˵������ҩ�о����漰�������е�ҽ��(���Դ����ݿ��,Ҫ�󱣴�)
'      ��ҩ���棺Ӧ����ҩ����֮����е���(�о���ֵ)
    Dim rsTmp As New adodb.Recordset
    Dim strҩƷ As String, str�÷� As String, lngҩƷID As Long, str������λ As String
    Dim strSQL As String, i As Long, k As Long
    
    AdviceCheckWarn = -1
    
    On Error GoTo errH
    Screen.MousePointer = 11
    
    If mstrNo = "" Then Exit Function
        
        
    '����PASS����״̬
    '-------------------------------------------------------------
    If PassGetState("PassEnable") = 0 Then
        MsgBox "��ǰ������ҩ���ϵͳ�����ã�������������Ƿ���ȷ��", vbInformation, gstrSysName
        Screen.MousePointer = 0: Exit Function
    End If
    
    '�ж���סԺ�������ﲡ�ˣ����û���ҵ���¼����ҽ�������˳�
    strSQL = "Select distinct B.����id,nvl(B.��ҳid,0) ��ҳid,nvl(C.�Һŵ�,'') �Һŵ� " & _
        " From ҩƷ�շ���¼ A,���˷��ü�¼ B,����ҽ����¼ C " & _
        " Where A.����id=B.Id And b.ҽ�����=c.Id And nvl(B.ҽ�����,0)<>0 And C.������� IN('5','6','7')" & _
        " And A.����=[2] And A.no=[1] "
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, mstrNo, mInt����)
    
    If rsTmp.RecordCount = 0 Then
        rsTmp.Close
        Exit Function
    End If
    
    mlngPatiID = rsTmp!����ID
    mstr�Һŵ� = NVL(rsTmp!�Һŵ�)
    mlng��ҳID = rsTmp!��ҳid
    
    '���벡�˾�����Ϣ(PASS��Ҫ�Ļ�������,ͬһ���˿ɲ��ظ�����)
    '-------------------------------------------------------------
    If mlngPatiID <> mlngPassPati Then
        If mstr�Һŵ� <> "" Then               '���ﲡ��
            strSQL = "Select ����ID,Count(Distinct Trunc(�Ǽ�ʱ��)) as ������� From ���˹Һż�¼ Where ����ID=[1] Group by ����ID"
            strSQL = "Select D.�������,A.����,A.�Ա�,A.��������," & _
                " C.���� as ������,C.���� as ������,E.��� as ҽ����,E.���� as ҽ����" & _
                " From ������Ϣ A,���˹Һż�¼ B,���ű� C,(" & strSQL & ") D,��Ա�� E" & _
                " Where A.����ID=B.����ID And B.ִ�в���ID=C.ID And A.����ID=D.����ID" & _
                " And B.ִ����=E.����(+) And A.����ID=[1] And B.NO=[2]"
            Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, mlngPatiID, mstr�Һŵ�)
            If rsTmp.EOF Then Screen.MousePointer = 0: Exit Function
        
            Call PassSetPatientInfo(mlngPatiID, rsTmp!�������, rsTmp!����, NVL(rsTmp!�Ա�), Format(rsTmp!��������, "yyyy-MM-dd"), "", "", _
                rsTmp!������ & "/" & rsTmp!������, IIf(Not IsNull(rsTmp!ҽ����), NVL(rsTmp!ҽ����) & "/" & NVL(rsTmp!ҽ����), ""), "")
        Else                                    'סԺ����
            strSQL = _
                " Select A.����,A.�Ա�,A.��������,B.��Ժ����,B.��Ժ����," & _
                " C.���� as ������,C.���� as ������,D.��� as ҽ����,D.���� as ҽ����" & _
                " From ������Ϣ A,������ҳ B,���ű� C,��Ա�� D" & _
                " Where A.����ID=B.����ID And B.��Ժ����ID=C.ID" & _
                " And B.סԺҽʦ=D.����(+) And A.����ID=[1] And B.��ҳID=[2]"
            Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, mlngPatiID, mlng��ҳID)
            If rsTmp.EOF Then Screen.MousePointer = 0: Exit Function
        
            Call PassSetPatientInfo(mlngPatiID, mlng��ҳID, rsTmp!����, NVL(rsTmp!�Ա�), Format(rsTmp!��������, "yyyy-MM-dd"), "", "", _
                rsTmp!������ & "/" & rsTmp!������, IIf(Not IsNull(rsTmp!ҽ����), NVL(rsTmp!ҽ����) & "/" & NVL(rsTmp!ҽ����), ""), _
                IIf(IsNull(rsTmp!��Ժ����), "", Format(rsTmp!��Ժ����, "yyyy-MM-dd")))
        End If
        mlngPassPati = mlngPatiID
    End If
    
    'PASS�Զ���˵����
    '-------------------------------------------------------------
    If lngCmd = 0 Then
        If TabShow = 0 Then
           'ȡҩƷ����
            strҩƷ = Billδ��ҩ�嵥.TextMatrix(lngRow, ����_δ��ҩ�嵥.ҩƷ����)
            lngҩƷID = Billδ��ҩ�嵥.TextMatrix(lngRow, ����_δ��ҩ�嵥.ҩƷID)
            str������λ = Billδ��ҩ�嵥.TextMatrix(lngRow, ����_δ��ҩ�嵥.������λ)
            'ȡҩƷ��ҩ;��
            str�÷� = Billδ��ҩ�嵥.TextMatrix(lngRow, ����_δ��ҩ�嵥.�÷�)
        Else
            'ȡҩƷ����
            strҩƷ = Bill�ѷ�ҩ�嵥.TextMatrix(lngRow, ����_�ѷ�ҩ�嵥.ҩƷ����)
            lngҩƷID = Bill�ѷ�ҩ�嵥.TextMatrix(lngRow, ����_�ѷ�ҩ�嵥.ҩƷID)
            str������λ = Bill�ѷ�ҩ�嵥.TextMatrix(lngRow, ����_�ѷ�ҩ�嵥.������λ)
            'ȡҩƷ��ҩ;��
            str�÷� = Bill�ѷ�ҩ�嵥.TextMatrix(lngRow, ����_�ѷ�ҩ�嵥.�÷�)
        End If
        
        If InStr(strҩƷ, " ") > 0 Then strҩƷ = Left(strҩƷ, InStr(strҩƷ, " ") - 1)
        If InStr(strҩƷ, "(") > 0 Then strҩƷ = Left(strҩƷ, InStr(strҩƷ, "(") - 1)
        '�����ѯҩƷ��Ϣ
        Call PassSetQueryDrug(lngҩƷID, strҩƷ, str������λ, str�÷�)
            
        '���ò˵�����״̬
        Call SetPassMenuState
        
        AdviceCheckWarn = 1 '��ʾ���Ե����˵�

        Screen.MousePointer = 0: Exit Function
    End If
    
    'ִ����Ӧ������
    '-------------------------------------------------------------
    Call PassDoCommand(lngCmd)
    Screen.MousePointer = 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function



Private Sub SetPassMenuState()
    '���ܣ�����Pass�˵�����״̬
    'Pass
    'һ���˵�
    'ҩ���ٴ���Ϣ�ο�
    mnuPassItem(0).Enabled = PassGetState("CPRRes") = 1
    'ҩƷ˵����
    mnuPassItem(1).Enabled = PassGetState("Directions") = 1
    '�й�ҩ��
    mnuPassItem(2).Enabled = PassGetState("Chp") = 1
    '������ҩ����
    mnuPassItem(3).Enabled = PassGetState("CPERes") = 1
    '����ֵ
    mnuPassItem(4).Enabled = PassGetState("CheckRes") = 1
    'ר����Ϣ
    'mnuPassItem(6).Enabled = PassGetState("") = 1
    'ҽҩ��Ϣ����
    mnuPassItem(8).Enabled = PassGetState("MEDInfo") = 1
    'ҩƷ�����Ϣ
    mnuPassItem(10).Enabled = PassGetState("MATCH-DRUG") = 1
    '��ҩ;�������Ϣ
    mnuPassItem(11).Enabled = PassGetState("MATCH-ROUTE") = 1
    'ҽԺҩƷ��Ϣ
    mnuPassItem(12).Enabled = PassGetState("HisDrugInfo") = 1
    
    '���˲˵�
    'ҩ��-ҩ���໥����
    mnuPassSpec(0).Enabled = PassGetState("DDIM") = 1
    'ҩ��-ʳ���໥ʹ��
    mnuPassSpec(1).Enabled = PassGetState("DFIM") = 1
    '����ע�����������
    mnuPassSpec(3).Enabled = PassGetState("MatchRes") = 1
    '����ע�����������
    mnuPassSpec(4).Enabled = PassGetState("TriessRes") = 1
    '����֢
    mnuPassSpec(6).Enabled = PassGetState("DDCM") = 1
    '������
    mnuPassSpec(7).Enabled = PassGetState("SIDE") = 1
    '��������ҩ
    mnuPassSpec(9).Enabled = PassGetState("GERI") = 1
    '��ͯ��ҩ
    mnuPassSpec(10).Enabled = PassGetState("PEDI") = 1
    '��������ҩ
    mnuPassSpec(11).Enabled = PassGetState("PREG") = 1
    '��������ҩ
    mnuPassSpec(12).Enabled = PassGetState("LACT") = 1
End Sub
Private Sub LoadPASS(ByVal BillStyle As Integer, ByVal BillNo As String)
    Dim strSQL As String
    Dim rs As New adodb.Recordset
    Dim n As Integer
    Dim strCondition As String
    On Error GoTo errHandle

    strSQL = "Select distinct B.����id,nvl(B.��ҳid,0) ��ҳid,nvl(C.�Һŵ�,0) �Һŵ�,B.ҽ����� " & _
        " From ҩƷ�շ���¼ A,���˷��ü�¼ B,����ҽ����¼ C " & _
        " Where A.����id=B.Id And b.ҽ�����=c.Id And nvl(B.ҽ�����,0)<>0 And C.������� IN('5','6','7')" & _
        " And A.����=[2] And A.no=[1] "
    Set rs = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, BillNo, BillStyle)

    If rs!�Һŵ� <> 0 Then
        strSQL = "Select A.ID,B.���� as �÷� From ����ҽ����¼ A,������ĿĿ¼ B" & _
        " Where A.�������='E' And B.�������� IN('2','4') And A.������ĿID=B.ID And A.����ID=[1] And A.�Һŵ�=[2] "
        strSQL = _
            " Select A.ID,A.���ID,Nvl(A.Ӥ��,0) as Ӥ��,A.�շ�ϸĿID,A.�������,A.ҽ������," & _
            " A.��������,B.���㵥λ,C.�÷�,A.Ƶ�ʴ���,A.����ҽ��,A.����ʱ��,A.ִ����ֹʱ��," & _
            " nvl(A.�����,-1) �����,nvl(A.��ҳid,0) ��ҳid,nvl(A.�Һŵ�,'') �Һŵ�,A.����id " & _
            " From ����ҽ����¼ A,������ĿĿ¼ B,(" & strSQL & ") C" & _
            " Where A.������ĿID=B.ID And A.���ID=C.ID And A.������� IN('5','6','7') And A.�շ�ϸĿID is Not Null" & _
            " And A.ҽ��״̬<>4 And (A.ҽ��״̬ Not IN(8,9) Or A.ҽ����Ч=1) " & _
            " And A.����ID=[1] And A.�Һŵ�=[2] " & _
            " And A.��ʼִ��ʱ�� is Not NULL" & _
            " Order by Nvl(A.Ӥ��,0),A.���"
        Set mrsPASS = zldatabase.OpenSQLRecord(strSQL, Me.Caption, CLng(rs!����ID), CStr(rs!�Һŵ�))
        
    ElseIf rs!��ҳid <> 0 Then
        strSQL = "Select A.ID,B.���� as �÷� From ����ҽ����¼ A,������ĿĿ¼ B" & _
        " Where A.�������='E' And B.�������� IN('2','4') And A.������ĿID=B.ID And A.����ID=[1] And A.��ҳID=[2] "
        strSQL = _
            " Select A.ID,A.���ID,Nvl(A.Ӥ��,0) as Ӥ��,A.�շ�ϸĿID,A.�������,A.ҽ������," & _
            " A.��������,B.���㵥λ,C.�÷�,A.Ƶ�ʴ���,A.����ҽ��,A.����ʱ��,A.ִ����ֹʱ��," & _
            " nvl(A.�����,-1) �����,nvl(A.��ҳid,0) ��ҳid,nvl(A.�Һŵ�,'') �Һŵ�,A.����id " & _
            " From ����ҽ����¼ A,������ĿĿ¼ B,(" & strSQL & ") C" & _
            " Where A.������ĿID=B.ID And A.���ID=C.ID And A.������� IN('5','6','7') And A.�շ�ϸĿID is Not Null" & _
            " And A.ҽ��״̬<>4 And (A.ҽ��״̬ Not IN(8,9) Or A.ҽ����Ч=1) " & _
            " And A.����ID=[1] And A.��ҳID=[2] " & _
            " And A.��ʼִ��ʱ�� is Not NULL" & _
            " Order by Nvl(A.Ӥ��,0),A.���"
        Set mrsPASS = zldatabase.OpenSQLRecord(strSQL, Me.Caption, CLng(rs!����ID), CLng(rs!��ҳid))
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub GetDrugFormat()
    Dim intCol As Integer
    Dim strSave As String
    Dim intRow As Integer, intRows As Integer
    Dim arrColumn
    
    'ȡ��ҩƷ���Ƶĸ�ʽ��ʽ
    strSave = zldatabase.GetPara("������", glngSys, 1342)
    If strSave = "" Then strSave = "0|ҩƷ����,0|������,0|Ӣ����,0|����,0|����ҽ��,0|״̬,0|����,0|NO,0|����Ա,0|����,0|����,0|סԺ��,0|���,0|����,0|����,0|��,0|����,0|������,0|׼����,0|��ҩ��,0|����,0|���,0|����,0|Ƶ��,0|�÷�,0|����ʱ��,0|˵��,0|����Ա,0|��ҩʱ��,0|��/��ҩ��,0|�ⷿ��λ"
    arrColumn = Split(strSave, ",")
    intҩƷ���� = Val(Split(arrColumn(0), "|")(0))
End Sub

Private Function GetColDefaultWidth(ByVal Bill As MSHFlexGrid, ByVal Col As Integer) As Integer
    Dim intCol As Integer
    Dim strSave As String
    Dim intRow As Integer, intRows As Integer
    Dim arrColumn
    
    '����ָ�����ָ���е�Ĭ�Ͽ��
    strSave = GetSetting("ZLSOFT", "����ģ��\����\" & App.ProductName & "\" & Me.Name, Bill.Name & "��Ĭ�Ͽ��", "")
    arrColumn = Split(strSave, ",")
    intRows = UBound(arrColumn)
    For intRow = 0 To intRows
'        intCol = GetDetailCol(Split(arrColumn(intRow), "|")(1), Bill)
        If Split(arrColumn(intRow), "|")(0) = Bill.TextMatrix(0, Col) Then
            GetColDefaultWidth = Split(arrColumn(intRow), "|")(1)
            Exit For
        End If
    Next
End Function

Private Sub SaveColDefaultWidth(ByVal Bill As MSHFlexGrid)
    '�����е�Ĭ�Ͽ��
    Dim strSave As String
    Dim i As Integer
    
    For i = 0 To Bill.Cols - 1
        strSave = strSave & Bill.TextMatrix(0, i) & "|" & Bill.ColWidth(i) & ","
    Next
    SaveSetting "ZLSOFT", "����ģ��\����\" & App.ProductName & "\" & "frm���ŷ�ҩ����", Bill.Name & "��Ĭ�Ͽ��", strSave
    
End Sub

Private Sub SetColHide(ByVal Bill As MSHFlexGrid)
    Dim intCol As Integer
    Dim strSave As String
    Dim intRow As Integer, intRows As Integer
    Dim arrColumn
    
    '�����û��ı��ز������ã���ʾ�������ز�����
    strSave = zldatabase.GetPara("������", glngSys, 1342)

    If strSave = "" Then strSave = "0|ҩƷ����,0|������,0|Ӣ����,0|����,0|����ҽ��,0|״̬,0|����,0|NO,0|����Ա,0|����,0|����,0|סԺ��,0|���,0|����,0|����,0|��,0|����,0|������,0|׼����,0|��ҩ��,0|����,0|���,0|����,0|Ƶ��,0|�÷�,0|����ʱ��,0|˵��,0|����Ա,0|��ҩʱ��,0|��/��ҩ��,0|�ⷿ��λ"
    arrColumn = Split(strSave, ",")
    intRows = UBound(arrColumn)
    For intRow = 0 To intRows
        intCol = GetDetailCol(Split(arrColumn(intRow), "|")(1), Bill)
        If intCol > -1 Then
            If Split(arrColumn(intRow), "|")(1) = "ҩƷ����" Then
                intҩƷ���� = Val(Split(arrColumn(intRow), "|")(0))
            Else
                If Val(Split(arrColumn(intRow), "|")(0)) = 1 Then
                    Bill.ColWidth(intCol) = 0
                ElseIf Bill.ColWidth(intCol) = 0 Then       '���Ҫ��ʾ���п�Ϊ0����ȡĬ�ϵ��п�
                    Bill.ColWidth(intCol) = GetColDefaultWidth(Bill, intCol)
                End If
            End If
        End If
    Next
    
    '������Ҫ��Ȩ��Ӱ�죬��ʱҪ����Ȩ����ȷ���Ƿ���ʾ
    If Bill.Name = "Billδ��ҩ�嵥" Then
        If UserPrivDetail.Priv_ҽ����ѯ = False Then
            Bill.ColWidth(����_δ��ҩ�嵥.����ҽ��) = 0
        Else
            Bill.ColWidth(����_δ��ҩ�嵥.����ҽ��) = 1100
        End If
    End If
End Sub


Private Sub SetColHideByMenu(ByVal MenuObj As Menu, ByVal Bill As MSHFlexGrid)
    Dim intCol As Integer
    Dim strSave As String
    Dim n As Integer
        
    intCol = GetDetailCol(MenuObj.Caption, Bill)
    If intCol > -1 Then
        If MenuObj.Checked = False Then
            Bill.ColWidth(intCol) = 0
        Else
            Bill.ColWidth(intCol) = GetColDefaultWidth(Bill, intCol)
        End If
    End If
    
    '��������
    strSave = intҩƷ���� & "|" & "ҩƷ����"
    
    For n = 0 To mnuBillItem.Count - 1
        If mnuBillItem(n).Caption = "��ҩ��" Then
            strSave = strSave & "," & IIf(mnuBillItem(n).Checked, "0", "1") & "|" & "��/��ҩ��"
        Else
            strSave = strSave & "," & IIf(mnuBillItem(n).Checked, "0", "1") & "|" & mnuBillItem(n).Caption
        End If
    Next
    
    zldatabase.SetPara "������", strSave, glngSys, 1342
End Sub
Private Sub SetColMenu()
    Dim strSave As String
    Dim intRow As Integer, intRows As Integer
    Dim arrColumn
    Dim n As Integer
    
    'ȡ����ע�������������Ŀ���Ʋ˵�
    strSave = zldatabase.GetPara("������", glngSys, 1342)
    
    If strSave = "" Then strSave = "0|ҩƷ����,0|������,0|Ӣ����,0|����,0|����ҽ��,0|״̬,0|����,0|NO,0|����Ա,0|����,0|����,0|סԺ��,0|���,0|����,0|����,0|��,0|����,0|������,0|׼����,0|��ҩ��,0|����,0|���,0|����,0|Ƶ��,0|�÷�,0|����ʱ��,0|˵��,0|����Ա,0|��ҩʱ��,0|��/��ҩ��,0|�ⷿ��λ"
    arrColumn = Split(strSave, ",")
    intRows = UBound(arrColumn)
    
    For n = 0 To Me.mnuDrugCodeName.Count - 1
        Me.mnuDrugCodeName(n).Checked = False
    Next
    
    For n = 0 To Me.mnuBillItem.Count - 1
        Me.mnuBillItem(n).Checked = False
    Next
    
    mbln��ʾ����ҩ�� = False
    
    For intRow = 0 To intRows
        If Split(arrColumn(intRow), "|")(1) = "ҩƷ����" Then
            Me.mnuDrugCodeName(Val(Split(arrColumn(intRow), "|")(0))).Checked = True
        Else
            For n = 0 To Me.mnuBillItem.Count - 1
                If Me.mnuBillItem(n).Caption = Split(arrColumn(intRow), "|")(1) And Me.mnuBillItem(n).Visible = True Then
                    If Val(Split(arrColumn(intRow), "|")(0)) = 0 Then
                        Me.mnuBillItem(n).Checked = True
                        If Me.mnuBillItem(n).Caption = "��/��ҩ��" Or Me.mnuBillItem(n).Caption = "��ҩ��" Then
                            mbln��ʾ����ҩ�� = True
                        End If
                    End If
                End If
            Next
        End If
    Next
    If UserPrivDetail.Priv_ҽ����ѯ = False Then
        mnuBillItem(1).Visible = False
    Else
        mnuBillItem(1).Visible = True
    End If
    
    If mblnҩƷ���� = False Then
        Me.mnuBillItem(29).Visible = False
    End If
End Sub
Private Sub Bill���ܷ�ҩ_EnterCell()
    Dim Col As Integer
    Dim lngTop As Long
    Dim lngLeft As Long
    Dim lngWidth As Long
    
    Bill��ҩ����.Visible = False
    With Bill���ܷ�ҩ
        .Height = TabShow.Height - TabShow.TabHeight - 120
        .Width = TabShow.Width - 150
    End With
    
    With Bill��ҩ����
        .Visible = False
        .Left = Bill���ܷ�ҩ.Left
        .Height = 1400
        .Width = TabShow.Width - 150
    End With
    
    Call SetSelectColor(Bill���ܷ�ҩ)
    
    If Lng������ʾ = 0 Then Exit Sub
    
    If txt������.Visible Then
        txt������_LostFocus
        txt������.Visible = False
    End If
    If CurCell.Row = 0 Or CurCell.Row >= Bill���ܷ�ҩ.rows - 2 Or Bill���ܷ�ҩ.TextMatrix(CurCell.Row, 0) = "С��" Then
        Exit Sub
    End If
    
    DoEvents
    
    If mbln���ܷ�ҩ = True Then
        If LoadDataInBill�����嵥(Val(Bill���ܷ�ҩ.TextMatrix(CurCell.Row, ����_���һ����嵥.��ҩ����id)), Val(Bill���ܷ�ҩ.TextMatrix(CurCell.Row, ����_���һ����嵥.ҩƷID))) = True Then
            Bill���ܷ�ҩ.Height = Bill���ܷ�ҩ.Height - Bill��ҩ����.Height - 25
            
            Bill��ҩ����.Visible = True
            Bill��ҩ����.Top = Bill���ܷ�ҩ.Top + Bill���ܷ�ҩ.Height + 25
        End If
    End If
        
    DoEvents
    
    '��������
    If IsHavePrivs(mstrPrivs, "�޸���������") = False Then Exit Sub
    
    If CurCell.Col <> ����_���һ����嵥.�������� And CurCell.Col <> ����_���һ����嵥.ʵ������ Then
        Exit Sub
    End If
    
    If Val(Bill���ܷ�ҩ.TextMatrix(CurCell.Row, ����_���һ����嵥.ʵ������)) < 0 Then
        Exit Sub
    End If
    
    LngLastRow = CurCell.Row
    lngLastCol = CurCell.Col
    
    lngLeft = TabShow.Left + Bill���ܷ�ҩ.Left + CurCell.CellLeft - 20
    lngTop = TabShow.Top + Bill���ܷ�ҩ.Top + CurCell.CellTop + 20
    
    lngWidth = CurCell.CellWidth - 20

    With txt������
        If .Visible = False Then
            .Alignment = 1
            .Move lngLeft, lngTop, lngWidth
            .Visible = True
            .ZOrder 0
            .SetFocus
            .Text = FormatEx(Val(Bill���ܷ�ҩ.TextMatrix(CurCell.Row, CurCell.Col)), 5)
        End If
    End With
    Call SelAll(txt������)
End Sub

Private Sub Bill���ܷ�ҩ_GotFocus()
    Bill���ܷ�ҩ_EnterCell
End Sub

Private Sub Bill�ܷ�ҩ�嵥_DblClick()
    Call Bill�ܷ�ҩ�嵥_KeyDown(vbKeySpace, 0)
End Sub

Private Sub Bill�ܷ�ҩ�嵥_EnterCell()
    Call SetSelectColor(Bill�ܷ�ҩ�嵥)
End Sub

Private Sub Bill�ܷ�ҩ�嵥_GotFocus()
    Bill�ܷ�ҩ�嵥_EnterCell
End Sub

Private Sub Bill�ܷ�ҩ�嵥_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeySpace Then Exit Sub
    With Bill�ܷ�ҩ�嵥
        If Trim(.TextMatrix(.Row, 1)) = "" Then Exit Sub
        
        Select Case Trim(.TextMatrix(.Row, 1))
        Case "�ָ�"
            Call UpdateRsByMenu(Nop_3, 3)
        Case "������"
            Call UpdateRsByMenu(ResumeDo, 3)
        End Select
    End With
End Sub

Private Sub Bill�ܷ�ҩ�嵥_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Dim MenuDefault As Menu
        With Bill�ܷ�ҩ�嵥
            If Trim(.TextMatrix(.Row, 1)) = "" Then Exit Sub
            If Trim(.TextMatrix(.Row, 1)) = "�ϼ�" Then Exit Sub
            
            Set MenuDefault = SetMenuCheck(PopMenu_3)
            PopupMenu PopMenu_3, 2, , , MenuDefault
        End With
    End If
End Sub

Private Sub Billȱҩ�嵥_EnterCell()
    Call SetSelectColor(Billȱҩ�嵥)
End Sub

Private Sub Billȱҩ�嵥_GotFocus()
    Billȱҩ�嵥_EnterCell
End Sub

Private Sub Billδ��ҩ�嵥_DblClick()
    Call Billδ��ҩ�嵥_KeyDown(vbKeySpace, 0)
End Sub

Private Sub Billδ��ҩ�嵥_EnterCell()
    Dim rsTmp As adodb.Recordset
    Dim rs���� As New adodb.Recordset
    Dim lng���� As Long, lngҩƷID As Long, Dbl���� As Double, blnAllow As Boolean
    Dim ArrayPhysic
    
    On Error GoTo errHandle
    If Not BlnEnterCell Then Exit Sub
    Call SetSelectColor(Billδ��ҩ�嵥)
    Cbo����.Clear
    Cbo����.Visible = False
    
    mstrNo = Billδ��ҩ�嵥.TextMatrix(Billδ��ҩ�嵥.Row, ����_δ��ҩ�嵥.NO)
    mInt���� = Val(Billδ��ҩ�嵥.TextMatrix(Billδ��ҩ�嵥.Row, ����_δ��ҩ�嵥.����))
            
    '����cmdAlley��ť״̬
    If mblnStarPass Then
        '�ж���סԺ�������ﲡ�ˣ����û���ҵ���¼����ҽ�����Ͳ���ʾcmdAlley��ť
        gstrSQL = "Select distinct B.����id,nvl(B.��ҳid,0) ��ҳid,nvl(C.�Һŵ�,'') �Һŵ� " & _
            " From ҩƷ�շ���¼ A,���˷��ü�¼ B,����ҽ����¼ C " & _
            " Where A.����id=B.Id And b.ҽ�����=c.Id And nvl(B.ҽ�����,0)<>0 And C.������� IN('5','6','7')" & _
            " And A.����=[2] And A.no=[1] "
        Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, mstrNo, mInt����)
        If rsTmp.RecordCount = 0 Then
            If cmdAlley.Visible Then cmdAlley.Visible = False
        Else
            If Not cmdAlley.Visible Then cmdAlley.Visible = True
        End If
    End If
    
    
    '�����ҩƷҩ���������㣬����ȡ��ҩƷ�������ι��û�ѡ��
    '�ۼ���ͬ(ָʱ��ҩƷ)�ҿ����㣬������ҩ����Ա��������
    With Billδ��ҩ�嵥
        If Not (Trim(.TextMatrix(.Row, ����_δ��ҩ�嵥.״̬)) = "��ҩ") Then Exit Sub
    End With
    
    If CurCell.Col = ����_δ��ҩ�嵥.���� Then
        RecChangeData.MoveFirst
        RecChangeData.Find "λ��=" & CurCell.Row
        If RecChangeData.EOF Then Exit Sub
        If RecChangeData!���� = 0 Then Exit Sub
        lng���� = RecChangeData!����
        lngҩƷID = RecChangeData!ҩƷID
        Dbl���� = FormatEx(RecChangeData!ʵ������, 5)
        ArrayPhysic = Split(GetPhysicDict(lngҩƷID), "^")        '��ȡ��ҩƷ�������Ϣ
        
        '������ڷ�ҩ��¼�Ҳ�����ҩ���������޸�������Ϣ
        blnAllow = False
        
        gstrSQL = " Select count(*) Records From ҩƷ�շ���¼ A,ҩƷ�շ���¼ B " & _
        " Where (Mod(A.��¼״̬,3)=0 or A.��¼״̬=1) And A.����� Is Not NULL And B.ID=[1] " & _
        " And A.NO=B.NO And A.����=B.���� And A.ҩƷID=B.ҩƷID And Nvl(A.����,0)=Nvl(B.����,0)"
        Set rs���� = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, CLng(RecChangeData!Id))
        
        
        blnAllow = (rs����!Records = 0)
        
        '��ȡ����������Ϣ
        gstrSQL = " SELECT B.�ϴ����� ����,B.����,ROUND(B.ʵ������/" & ArrayPhysic(3) & ",2) ����" & _
         " FROM ҩƷ��� A,ҩƷ��� B,�շѼ�Ŀ C,�շ���ĿĿ¼ F" & _
         " WHERE A.ҩƷID = B.ҩƷID AND b.ҩƷID=F.ID" & _
         " AND B.�ⷿID = [1] AND B.ҩƷID=[2] AND A.ҩƷID = C.�շ�ϸĿID" & _
         " AND ((SYSDATE BETWEEN C.ִ������ AND C.��ֹ����) OR C.��ֹ���� IS NULL)" & _
         " AND NVL(����,0)<>0 AND NVL(ʵ������,0)<>0 AND ����=1" & _
         " AND ROUND(DECODE(F.�Ƿ���,NULL,C.�ּ�,0,C.�ּ�,B.ʵ�ʽ��/B.ʵ������),2)=" & _
         "     (SELECT ROUND(DECODE(F.�Ƿ���,NULL,C.�ּ�,0,C.�ּ�,B.ʵ�ʽ��/B.ʵ������),2) ����" & _
         "     FROM ҩƷ��� A,ҩƷ��� B,�շѼ�Ŀ C,�շ���ĿĿ¼ F" & _
         "     WHERE A.ҩƷID = B.ҩƷID AND b.ҩƷID=f.ID " & _
         "     AND B.�ⷿID = [1] AND B.ҩƷID=[2] AND A.ҩƷID = C.�շ�ϸĿID" & _
         "     AND ((SYSDATE BETWEEN C.ִ������ AND C.��ֹ����) OR C.��ֹ���� IS NULL)" & _
         "     AND NVL(����,0)<>0 AND NVL(ʵ������,0)<>0 AND ����=1 AND NVL(����,0)=[3])" & _
         " AND ROUND(B.ʵ������/" & ArrayPhysic(3) & ",2)>=[4] AND (NVL(A.ҩ������,0)=0 OR (NVL(A.ҩ������,0)=1 AND (Ч�� IS NULL OR Ч��>TRUNC(SYSDATE))))" & _
         " ORDER BY B.����"
        Set rs���� = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngҩ��ID, lngҩƷID, lng����, Dbl����)
        
        With rs����
            Do While Not .EOF
                If (!���� <> lng���� And blnAllow) Or !���� = lng���� Then
                    Cbo����.AddItem IIf(IsNull(!����), "", !����) & "(" & !���� & ")"
                    Cbo����.ItemData(Cbo����.NewIndex) = !����
                End If
                .MoveNext
            Loop
        End With
        Call LocateCboItemData(Cbo����, lng����)
        Call ShowCbo
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Billδ��ҩ�嵥_GotFocus()
    Billδ��ҩ�嵥_EnterCell
End Sub

Private Sub Billδ��ҩ�嵥_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeySpace Then Exit Sub
    
    CurCell.Col = 0
    Cbo����.Visible = False
    
    With Billδ��ҩ�嵥
        If Trim(.TextMatrix(.Row, ����_δ��ҩ�嵥.״̬)) = "" Then Exit Sub
        If Trim(.TextMatrix(.Row, ����_δ��ҩ�嵥.״̬)) = "ȱҩ" Then Exit Sub
            
        RecChangeData.MoveFirst
        RecChangeData.Find "λ��=" & Billδ��ҩ�嵥.Row
        If RecChangeData.EOF Then Exit Sub
        If RecChangeData!����� = "" And Int����δ��˴�����ҩ = 0 Then
            Select Case Trim(.TextMatrix(.Row, ����_δ��ҩ�嵥.״̬))
            Case "�ܷ�"
                Call UpdateRsByMenu(Nop_1, 1)
            Case "������"
                Call UpdateRsByMenu(HandBack, 1)
            End Select
            Exit Sub
        End If
        
        Select Case Trim(.TextMatrix(.Row, ����_δ��ҩ�嵥.״̬))
        Case "��ҩ"
            Call UpdateRsByMenu(HandBack, 1)
        Case "�ܷ�"
            Call UpdateRsByMenu(Lack, 1)
        Case "ȱҩ"
            Call UpdateRsByMenu(Nop_1, 1)
        Case "������"
            Call UpdateRsByMenu(Consignment, 1)
        End Select
    End With
End Sub

Private Sub Billδ��ҩ�嵥_LostFocus()
    Call Cbo����_LostFocus
End Sub

Private Sub Billδ��ҩ�嵥_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim strҩƷ As String
    Dim intCurRow As Integer
    Dim intCurCol As Integer
    
    intCurRow = Billδ��ҩ�嵥.MouseRow
    intCurCol = Billδ��ҩ�嵥.MouseCol
    
    If Button = 2 Then
        If intCurRow = 0 Then
            PopupMenu mnuColHide, 2
            Exit Sub
        End If
        If intCurCol > 0 Then
            Dim MenuDefault As Menu
        
            CurCell.Col = 1
            Cbo����.Visible = False
            
            With Billδ��ҩ�嵥
                Consignment.Enabled = True
                If Trim(.TextMatrix(.Row, ����_δ��ҩ�嵥.״̬)) = "" Then Exit Sub
                If Trim(.TextMatrix(.Row, ����_δ��ҩ�嵥.״̬)) = "�ϼ�" Then Exit Sub
                If Trim(.TextMatrix(.Row, ����_δ��ҩ�嵥.״̬)) = "ȱҩ" Then Exit Sub
                If RecChangeData.RecordCount = 0 Then Exit Sub
                
                RecChangeData.MoveFirst
                RecChangeData.Find "λ��=" & Billδ��ҩ�嵥.Row
                If RecChangeData.EOF Then Exit Sub
                If RecChangeData!����� = "" And Int����δ��˴�����ҩ = 0 Then Consignment.Enabled = False
                
                Set MenuDefault = SetMenuCheck(PopMenu_1)
                PopupMenu PopMenu_1, 2, , , MenuDefault
            End With
        ElseIf intCurCol = 0 Then
            mstrNo = Billδ��ҩ�嵥.TextMatrix(Billδ��ҩ�嵥.Row, ����_δ��ҩ�嵥.NO)
            mInt���� = Val(Billδ��ҩ�嵥.TextMatrix(Billδ��ҩ�嵥.Row, ����_δ��ҩ�嵥.����))
            
            '���Pass״̬
            If AdviceCheckWarn(0, Billδ��ҩ�嵥.Row) >= 0 Then PopupMenu mnuPass, 2
        End If
            
    End If
End Sub

Private Sub Billδ��ҩ�嵥_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim strColumn As String
    Dim bln�Ƿ���ʾ���� As Boolean
        
    '����������
    With Billδ��ҩ�嵥
        If Button <> 1 Then Exit Sub
        If .MouseRow <> 0 Then Exit Sub

        strColumn = .TextMatrix(.MouseRow, .MouseCol)
        If InStr(1, gstr��������, "|" & strColumn & "|") = 0 Then Exit Sub
        If strColumn = "ҩƷ����" Then strColumn = "Ʒ��"
        
        'ֻ�а�NO����ʱ����ʾ����
        If strColumn = "NO" Then bln�Ƿ���ʾ���� = True

        '���������ͬ����ı�����ʽ����������ʽ
        If str����_δ��ҩ Like "*" & strColumn & "*" Then
            str����_δ��ҩ = ExchangeOrder(str����_δ��ҩ)
        Else
            str����_δ��ҩ = strColumn & strAsc
        End If
    End With

    '������ʾδ��ҩ�嵥
    Call ClearCons
    Call LoadDataInBillδ��ҩ�嵥
    Call SetGroup(Billδ��ҩ�嵥, bln�Ƿ���ʾ����)
    
End Sub

Private Sub Bill�ѷ�ҩ�嵥_DblClick()
    Call Bill�ѷ�ҩ�嵥_KeyDown(vbKeySpace, 0)
End Sub

Private Sub Bill�ѷ�ҩ�嵥_EnterCell()
    Dim rsTmp As adodb.Recordset
    Dim rs���� As New adodb.Recordset
    Dim lng���� As Long, lngҩƷID As Long, Dbl���� As Double
    Dim ArrayPhysic
    
    On Error GoTo errHandle
    mnuFileRestore = False
    If TxtInput.Visible Then
        Call TxtInput_LostFocus
        TxtInput.Visible = False
    End If
    
    mstrNo = Bill�ѷ�ҩ�嵥.TextMatrix(Bill�ѷ�ҩ�嵥.Row, ����_�ѷ�ҩ�嵥.NO)
    mInt���� = Val(Bill�ѷ�ҩ�嵥.TextMatrix(Bill�ѷ�ҩ�嵥.Row, ����_�ѷ�ҩ�嵥.����))
            
    '����cmdAlley��ť״̬
    If mblnStarPass Then
        '�ж���סԺ�������ﲡ�ˣ����û���ҵ���¼����ҽ�����Ͳ���ʾcmdAlley��ť
        gstrSQL = "Select distinct B.����id,nvl(B.��ҳid,0) ��ҳid,nvl(C.�Һŵ�,'') �Һŵ� " & _
            " From ҩƷ�շ���¼ A,���˷��ü�¼ B,����ҽ����¼ C " & _
            " Where A.����id=B.Id And b.ҽ�����=c.Id And nvl(B.ҽ�����,0)<>0 And C.������� IN('5','6','7')" & _
            " And A.����=[2] And A.no=[1] "
        Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, mstrNo, mInt����)
        If rsTmp.RecordCount = 0 Then
            If cmdAlley.Visible Then cmdAlley.Visible = False
        Else
            If Not cmdAlley.Visible Then cmdAlley.Visible = True
        End If
    End If
    
    If Not IsHavePrivs(mstrPrivs, "��ҩ") Then Exit Sub
    
    '��ʾ��ҩ���ı���ȱʡΪ��ǰ��λ�����ݣ������û��޸ġ�
    '�������ֵ�Ƿ����㡢�ո񡢷Ƿ���������ȫ��������������ȱʡΪȫ��
    With Bill�ѷ�ҩ�嵥
        .Col = ����_�ѷ�ҩ�嵥.��ҩ��
        Call SetSelectColor(Bill�ѷ�ҩ�嵥)
    End With
    
    'ǿ���趨����Ϊ��ҩ����
    With RecChangeSendedData
        If CurCell.Col = ����_�ѷ�ҩ�嵥.��ҩ�� And Bill�ѷ�ҩ�嵥.ColWidth(����_�ѷ�ҩ�嵥.��ҩ��) > 0 Then
            If .RecordCount = 0 Then Exit Sub
            .MoveFirst
            .Find "λ��=" & CurCell.Row
            If .EOF Then Exit Sub
            If !�ɲ��� = 0 Then Exit Sub        '��ʾ�ü�¼�Ƿ���ԭʼ��¼

            '��֤ÿ��EnterCell�¼�����ʱ���������˲˵�"��ӡ��ҩ֪ͨ��"
            mnuFileRestore = (!�ɲ��� = 3)
            If Not (Trim(Bill�ѷ�ҩ�嵥.TextMatrix(Bill�ѷ�ҩ�嵥.Row, ����_�ѷ�ҩ�嵥.״̬)) = "��ҩ") Then Exit Sub

            TxtInput.Tag = Val(!׼����)
            TxtInput.Text = FormatEx(Val(Bill�ѷ�ҩ�嵥.TextMatrix(Bill�ѷ�ҩ�嵥.Row, ����_�ѷ�ҩ�嵥.��ҩ��)), 5)
            Call ShowTxt
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Bill�ѷ�ҩ�嵥_GotFocus()
    Bill�ѷ�ҩ�嵥_EnterCell
End Sub

Private Sub Bill�ѷ�ҩ�嵥_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not (KeyCode = vbKeySpace) Then Exit Sub
    If Not IsHavePrivs(mstrPrivs, "��ҩ") Then Exit Sub
    CurCell.Col = 0
    TxtInput.Visible = False
    
    With Bill�ѷ�ҩ�嵥
        If Trim(.TextMatrix(.Row, ����_�ѷ�ҩ�嵥.״̬)) = "" Then Exit Sub
        With RecChangeSendedData
            If .RecordCount = 0 Then
                MsgErr "�����б仯����ˢ�º����ԣ�"
                Exit Sub
            End If
            .MoveFirst
            .Find "λ��=" & Bill�ѷ�ҩ�嵥.Row
            If .EOF Then Exit Sub
            If !�ɲ��� <> 1 Then Exit Sub
        End With
        
        Select Case Trim(.TextMatrix(.Row, ����_�ѷ�ҩ�嵥.״̬))
        Case "��ҩ"
            Call UpdateRsByMenu(Nop_1, 2)
        Case "������"
            If .ColWidth(����_�ѷ�ҩ�嵥.��ҩ��) = 0 Then Exit Sub
            Call UpdateRsByMenu(Restore, 2)
        End Select
    End With
    Call Bill�ѷ�ҩ�嵥_EnterCell
End Sub

Private Sub Bill�ѷ�ҩ�嵥_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim strҩƷ As String
    Dim intCurRow As Integer
    Dim intCurCol As Integer
    
    intCurRow = Bill�ѷ�ҩ�嵥.MouseRow
    intCurCol = Bill�ѷ�ҩ�嵥.MouseCol

    If CurCell.Col <> ����_�ѷ�ҩ�嵥.��ҩ�� Then
        CurCell.Col = 0
        TxtInput.Visible = False
    End If

    If Button = 2 Then
        If intCurRow = 0 Then
            PopupMenu mnuColHide, 2
            Exit Sub
        End If
        If intCurCol > 0 Then
            Dim MenuDefault As Menu
            With Bill�ѷ�ҩ�嵥
                If Trim(.TextMatrix(.Row, ����_�ѷ�ҩ�嵥.״̬)) = "" Then Exit Sub
                If Trim(.TextMatrix(.Row, ����_�ѷ�ҩ�嵥.״̬)) = "�ϼ�" Then Exit Sub
                With RecChangeSendedData
                    If .RecordCount = 0 Then
                        MsgErr "�����б仯����ˢ�º����ԣ�"
                        Exit Sub
                    End If
                    .MoveFirst
                    .Find "λ��=" & Bill�ѷ�ҩ�嵥.Row
                    If .EOF Then Exit Sub
                    If !�ɲ��� <> 1 Then Exit Sub
                End With
                
                Set MenuDefault = SetMenuCheck(PopMenu_2)
                PopupMenu PopMenu_2, 2, , , MenuDefault
            End With
        ElseIf intCurCol = 0 Then
            mstrNo = Bill�ѷ�ҩ�嵥.TextMatrix(Bill�ѷ�ҩ�嵥.Row, ����_�ѷ�ҩ�嵥.NO)
            mInt���� = Val(Bill�ѷ�ҩ�嵥.TextMatrix(Bill�ѷ�ҩ�嵥.Row, ����_�ѷ�ҩ�嵥.����))

            '���Pass״̬
            If AdviceCheckWarn(0, Bill�ѷ�ҩ�嵥.Row) >= 0 Then PopupMenu mnuPass, 2
        End If
    End If
End Sub

Private Sub Bill�ѷ�ҩ�嵥_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim strColumn As String
    Dim bln�Ƿ���ʾ���� As Boolean
    
    '����������
    With Bill�ѷ�ҩ�嵥
        If Button <> 1 Then Exit Sub
        If .MouseRow <> 0 Then Exit Sub
'        If Chk�嵥.Value = 1 Then Exit Sub
        
        strColumn = .TextMatrix(.MouseRow, .MouseCol)
        If InStr(1, gstr��������, "|" & strColumn & "|") = 0 Then Exit Sub
        If strColumn = "ҩƷ����" Then strColumn = "Ʒ��"
        
        'ֻ�в���ʾ���̵��ݲ����ǰ�NO����ʱ����ʾ����
        If Chk�嵥.Value = 0 And strColumn = "NO" Then bln�Ƿ���ʾ���� = True
        
        '���������ͬ����ı�����ʽ����������ʽ
        If str����_����ҩ Like "*" & strColumn & "*" Then
            str����_����ҩ = ExchangeOrder(str����_����ҩ)
        Else
            str����_����ҩ = strColumn & strAsc
        End If
    End With
    
    '������ʾδ��ҩ�嵥
    Call ClearBill(Bill�ѷ�ҩ�嵥)
    Call LoadDataInBill�ѷ�ҩ�嵥
    Call SetGroup(Bill�ѷ�ҩ�嵥, bln�Ƿ���ʾ����)
End Sub

Private Sub Cbo����_Click()
    RecChangeData.MoveFirst
    RecChangeData.Find "λ��=" & CurCell.Row
    If RecChangeData.EOF Then Exit Sub
    
    With RecChangeData
        If !���� = 0 Then Exit Sub
        If !���� = Cbo����.ItemData(Cbo����.ListIndex) Then Exit Sub
        !���� = Cbo����.ItemData(Cbo����.ListIndex)
        !���� = Cbo����.Text
        .Update
    End With
    With Billδ��ҩ�嵥
        .TextMatrix(.Row, ����_δ��ҩ�嵥.����) = Cbo����
    End With
End Sub

Private Sub Cbo����_LostFocus()
    On Error Resume Next
    
    If InStr(1, "Billδ��ҩ�嵥,Cbo����", ActiveControl.Name) = 0 Then
        CurCell.Col = 0
        Cbo����.Visible = False
    End If
End Sub

Private Sub Chk�嵥_Click()
    '��¼��ǰҳ��CHECK���״̬�����л�����Իָ�ԭ����״̬
    If TabShow.Tab = 1 Then
        Chk�嵥.Tag = Chk�嵥.Value & Mid(Chk�嵥.Tag, 2, 1)
    ElseIf TabShow.Tab = 4 Then
        Chk�嵥.Tag = Mid(Chk�嵥.Tag, 1, 1) & Chk�嵥.Value
    End If
    Call mnuViewRefresh_Click
End Sub

Private Sub cmdAlley_Click()
    '���ܣ��Բ��˹���ʷ/����״̬���й���
    'Pass
    Call AdviceCheckWarn(21)
End Sub

Private Sub Consignment_Click()
    Call UpdateRsByMenu(Consignment, 1)
End Sub

Private Sub ConsignmentALL_Click()
    'ȫ����ҩ
    Dim Strִ��״̬ As String
    Dim intCol As Integer
    Dim lngColor As Long
    
    Billδ��ҩ�嵥.Redraw = False
    With RecChangeData
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            If !ִ��״̬ <> 0 Then
                If Not (!����� = "" And Int����δ��˴�����ҩ = 0) Then
                    !ִ��״̬ = 1
                    Strִ��״̬ = IIf(!ִ��״̬ = 0, "ȱҩ", IIf(!ִ��״̬ = 1, "��ҩ", IIf(!ִ��״̬ = 2, "�ܷ�", "������")))
                    !״̬ = Strִ��״̬
                    .Update
                
                    '����ü�¼����䵽�������������
                    With Billδ��ҩ�嵥
                        If .rows - 1 >= RecChangeData!λ�� Then .TextMatrix(RecChangeData!λ��, ����_δ��ҩ�嵥.״̬) = Strִ��״̬
                        
                        lngColor = IIf(Strִ��״̬ = "��ҩ", glngSendBlkColor, glngOtherBlkColor)
                        
                        .Row = RecChangeData!λ��
                        For intCol = 0 To .Cols - 1
                            .Col = intCol
                            .CellBackColor = lngColor
                        Next
                    End With
                End If
            End If
            
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
    End With
    
    Billδ��ҩ�嵥.Redraw = True
    
    '���ò˵������߰�ť��״̬
    Call SetMenuAndToolbarState
End Sub

Private Sub Form_Activate()
    Dim dateCurDate As Date
    
    On Error Resume Next
    
    If BlnStartUp = False Then
        Unload Me
        Exit Sub
    End If
    
    mblnFirstSended = True
    
    If BlnFirstStart = False Then
'        mnuViewRefresh_Click
    End If
   
    Form_Resize
    BlnFirstStart = True
    TimerAuto.Enabled = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    If TabShow.Tab = 4 And ActiveControl.Name = "TxtInput" Then
        If (KeyCode = vbKeySpace Or KeyCode = vbKeyReturn Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) Then
            If KeyCode = vbKeySpace Then
                Call Bill�ѷ�ҩ�嵥_KeyDown(KeyCode, 0)
            Else
                If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
                    If Bill�ѷ�ҩ�嵥.Row + 1 < Bill�ѷ�ҩ�嵥.rows - 1 Then Bill�ѷ�ҩ�嵥.Row = Bill�ѷ�ҩ�嵥.Row + 1
                Else
                    If Bill�ѷ�ҩ�嵥.Row - 1 > 0 Then Bill�ѷ�ҩ�嵥.Row = Bill�ѷ�ҩ�嵥.Row - 1
                End If
            End If
            Call Bill�ѷ�ҩ�嵥_EnterCell
            Bill�ѷ�ҩ�嵥.SetFocus
        End If
    End If
    
    If Lvw��ҩ;��.Visible = True Then
        If KeyCode = 102 Or KeyCode = 65 Then
            If Shift = vbCtrlMask Then   'Ctrl+A
                Call SelectAllCheck(Lvw��ҩ;��)
            End If
        End If
        
        If KeyCode = 102 Or KeyCode = 82 Then
            If Shift = vbCtrlMask Then   'Ctrl+R
                Call UnSelectAllCheck(Lvw��ҩ;��)
            End If
        End If
    End If
    
    If Lvw����.Visible = True Then
        If KeyCode = 102 Or KeyCode = 65 Then
            If Shift = vbCtrlMask Then   'Ctrl+A
                Call SelectAllCheck(Lvw����)
            End If
        End If
        
        If KeyCode = 102 Or KeyCode = 82 Then
            If Shift = vbCtrlMask Then   'Ctrl+R
                Call UnSelectAllCheck(Lvw����)
            End If
        End If
    End If
    
    
    If Lvw����.Visible = True Then
        If KeyCode = 102 Or KeyCode = 65 Then
            If Shift = vbCtrlMask Then   'Ctrl+A
                Call SelectAllCheck(Lvw����)
            End If
        End If
        
        If KeyCode = 102 Or KeyCode = 82 Then
            If Shift = vbCtrlMask Then   'Ctrl+R
                Call UnSelectAllCheck(Lvw����)
            End If
        End If
        
        If KeyCode = vbKeyEscape Then
            Call Lvw����_LostFocus
        End If
    End If
    
    err = 0
End Sub

Private Sub SelectAllCheck(ByVal UserListView As ListView)
    Dim n As Integer
    
    For n = 1 To UserListView.ListItems.Count
        UserListView.ListItems(n).Checked = True
    Next
End Sub

Private Sub UnSelectAllCheck(ByVal UserListView As ListView)
    Dim n As Integer
    
    For n = 1 To UserListView.ListItems.Count
        UserListView.ListItems(n).Checked = False
    Next
End Sub
Private Sub Form_Load()
    Dim dblAdjustWidth As Double
    Dim dblAdjustWidth1 As Double
    
    BlnEnterCell = False
    str����_δ��ҩ = "NO " & strAsc
    str����_����ҩ = "NO " & strAsc
    
    '��ʼ������
    BlnStartUp = False
    BlnFirstStart = False
    Blnˢ��δ��ҩ�嵥 = True
    Bln����� = True
    mdblConditonHeight = 3000
    
    If Screen.Width \ Screen.TwipsPerPixelX <= 800 Then
        mbln�ͷֱ��� = True
    End If
    
    '�ͷֱ���ʱ�������ֿؼ��Ŀ�Ȼ�λ��
    If mbln�ͷֱ��� Then
        dblAdjustWidth = lblInfo.Left - Dtp����ʱ��.Left - Dtp����ʱ��.Width - 100
        
        fraConNormal.Width = fraConNormal.Width - dblAdjustWidth
        fraConExpand.Width = fraConNormal.Width
        fraConRequest.Width = fraConNormal.Width
        
        lblInfo.Left = lblInfo.Left - dblAdjustWidth
        lblPatiInputType.Left = lblPatiInputType.Left - dblAdjustWidth
        txtPati.Left = txtPati.Left - dblAdjustWidth
        
        txt����.Width = txt����.Width - dblAdjustWidth
        cmd��������.Left = cmd��������.Left - dblAdjustWidth
        
        
        dblAdjustWidth1 = lblҩƷ����.Left - cmd��ҩ;��.Left - 100
        
        If dblAdjustWidth > dblAdjustWidth1 Then
            txt��ҩ;��.Width = txt��ҩ;��.Width - (dblAdjustWidth - dblAdjustWidth1) / 2 - 150
            txtҩƷ����.Width = txtҩƷ����.Width - (dblAdjustWidth - dblAdjustWidth1) / 2 - 150
        End If
        
        cmd��ҩ;��.Left = txt��ҩ;��.Left + txt��ҩ;��.Width + 10
        
        lblҩƷ����.Left = cmd��ҩ;��.Left + cmd��ҩ;��.Width + 100
        txtҩƷ����.Left = lblҩƷ����.Left + lblҩƷ����.Width + 100
        cmdҩƷ����.Left = txtҩƷ����.Left + txtҩƷ����.Width + 10
        Lblҽ������.Left = lblҩƷ����.Left
        Cboҽ������.Left = txtҩƷ����.Left
        opt��Χ(1).Left = opt��Χ(1).Left - 150
        opt��Χ(2).Left = opt��Χ(2).Left - 150
    End If
    
    mlngMode = glngModul
    mstrPrivs = gstrprivs
    
    If gstrUserName = "" Then
        MsgBox "��Ϊ��ǰ�û����ö�Ӧ�Ĳ���Ա����ʹ�ñ�ģ�飡", vbInformation, gstrSysName
        Exit Sub
    End If
    
    blnҩƷ���������� = GetDepend
    
    Call GetSysParms
    
    Call GetPrivs
    
    Call TradeName
    
    'Ϊ���ؼ�װ��ͼ��
    If LoadInIcon = False Then Exit Sub
    '�������ݼ��
    If DependOnCheck = False Then Exit Sub
    
    '��ʼ��������
    Call IniConditon
    
      
    Call LoadCondition(IIf(TabShow.Tab = 4, 1, 0))
    
    '��ʼ����¼��
    Call InitRec
    Call InitRefreshRec
    '���ø��ؼ�����ʽ
    Call SetFormat
    
    Call Ȩ�޿���
    
    BlnStartUp = True
    BlnEnterCell = True
    RestoreWinState Me, App.ProductName
    '�ָ����Ի����ú��м���ʼ�ղ�������
    If Billδ��ҩ�嵥.ColWidth(����_δ��ҩ�嵥.״̬) < 200 Then Billδ��ҩ�嵥.ColWidth(����_δ��ҩ�嵥.״̬) = 700
    If Billδ��ҩ�嵥.ColWidth(����_δ��ҩ�嵥.����) < 200 Then Billδ��ҩ�嵥.ColWidth(����_δ��ҩ�嵥.����) = 1500
    If Bill�ѷ�ҩ�嵥.ColWidth(����_�ѷ�ҩ�嵥.״̬) < 200 Then Bill�ѷ�ҩ�嵥.ColWidth(����_�ѷ�ҩ�嵥.״̬) = 700
    If Bill�ѷ�ҩ�嵥.ColWidth(����_�ѷ�ҩ�嵥.��ҩ��) < 200 Then Bill�ѷ�ҩ�嵥.ColWidth(����_�ѷ�ҩ�嵥.��ҩ��) = 1000
    '��ʾ�и��ݲ����������Ƿ���ʾ
    Billδ��ҩ�嵥.ColWidth(����_δ��ҩ�嵥.�����) = IIf(Not mblnStarPass, 0, 240)
    Bill�ѷ�ҩ�嵥.ColWidth(����_�ѷ�ҩ�嵥.�����) = IIf(Not mblnStarPass, 0, 240)
    
    Call zldatabase.ShowReportMenu(Me, glngSys, mlngMode, mstrPrivs, "ZL1_INSIDE_1342_1")
    
    'ȡͨ�����ز������õ��Ƿ���ʾ����ֵ
    Call SetColMenu
    Call SetColHide(Billδ��ҩ�嵥)
    Call SetColHide(Bill�ѷ�ҩ�嵥)
    
    'ȡҩ����Ա
    Call Get��ҩ��
    
    'ȡ��ҩ����ʽ
    Call Get��ҩ����ʽ
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    mlngMyWindow = 0
    
    If TabShow.Tab = 0 Then
        SaveSetting "ZLSOFT", "����ģ��\����\" & App.ProductName & "\Frm���ŷ�ҩ����", "��ʾ��ҩ��������", mlng��������
    End If
    
'    '���淢ҩ����
    Call SetCondition(IIf(TabShow.Tab = 4, 1, 0))
    Call SaveCondition(IIf(TabShow.Tab = 4, 1, 0))
    Call ClearCondition(IIf(TabShow.Tab = 4, 1, 0))
    
    mintLastTab = 0
    
    '�����δ��ҩ�嵥����ҩ�� �����򱣴�������
    If Billδ��ҩ�嵥.ColWidth(����_δ��ҩ�嵥.״̬) < 200 Then Billδ��ҩ�嵥.ColWidth(����_δ��ҩ�嵥.״̬) = 700
    If Billδ��ҩ�嵥.ColWidth(����_δ��ҩ�嵥.����) < 200 Then Billδ��ҩ�嵥.ColWidth(����_δ��ҩ�嵥.����) = 1500
    If Bill�ѷ�ҩ�嵥.ColWidth(����_�ѷ�ҩ�嵥.״̬) < 200 Then Bill�ѷ�ҩ�嵥.ColWidth(����_�ѷ�ҩ�嵥.״̬) = 700
    If Bill�ѷ�ҩ�嵥.ColWidth(����_�ѷ�ҩ�嵥.��ҩ��) < 200 Then Bill�ѷ�ҩ�嵥.ColWidth(����_�ѷ�ҩ�嵥.��ҩ��) = 1000
    
    Billδ��ҩ�嵥.Tag = "": Bill�ѷ�ҩ�嵥.Tag = ""
    Billȱҩ�嵥.Tag = "": Bill�ܷ�ҩ�嵥.Tag = "": Bill���ܷ�ҩ.Tag = ""
    Call SaveFlexState(Billδ��ҩ�嵥, "δ��ҩ�嵥")
    Call SaveFlexState(Bill�ѷ�ҩ�嵥, "�ѷ�ҩ�嵥")
    Call SaveFlexState(Bill���ܷ�ҩ, "���ܷ�ҩ" & Lng������ʾ)
    Call SaveFlexState(Billȱҩ�嵥, "ȱҩ�嵥")
    Call SaveFlexState(Bill�ܷ�ҩ�嵥, "�ܷ�ҩ�嵥")
    SaveWinState Me, App.ProductName
End Sub

Private Sub Form_Resize()
    Dim DblHeight As Double, DblWidth As Double
    Dim dblMaxWidth As Double
    
    On Error Resume Next
    
    dblMaxWidth = IIf(mbln�ͷֱ���, 12240, 13275)
    
    If Me.WindowState = 1 Then Exit Sub
    
    If BlnFirstStart = False Then
        Cbar.Align = 1
        With Cbar
            Set .Bands(1).Child = Tbar
            .Bands(1).MinHeight = Tbar.Height
        End With
    End If
    
    If Me.Height < 8500 Then Me.Height = 8500
    If Me.Width < dblMaxWidth Then Me.Width = dblMaxWidth
    
        
    '����������
    Call ResizeCondition
    
End Sub

Private Sub HandBack_Click()
    Call UpdateRsByMenu(HandBack, 1)
End Sub

Private Sub HandBackALL_Click()
    Dim Strִ��״̬ As String
    Dim intCol As Integer
    Dim lngColor As Long
    
    'ȫ���ܷ�
    Billδ��ҩ�嵥.Redraw = False
    With RecChangeData
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            If !ִ��״̬ <> 0 Then
                !ִ��״̬ = 2
                Strִ��״̬ = IIf(!ִ��״̬ = 0, "ȱҩ", IIf(!ִ��״̬ = 1, "��ҩ", IIf(!ִ��״̬ = 2, "�ܷ�", "������")))
                !״̬ = Strִ��״̬
                .Update
                
                '����ü�¼����䵽�������������
                With Billδ��ҩ�嵥
                    If .rows - 1 >= RecChangeData!λ�� Then .TextMatrix(RecChangeData!λ��, ����_δ��ҩ�嵥.״̬) = Strִ��״̬
                    
                    lngColor = IIf(Strִ��״̬ = "��ҩ", glngSendBlkColor, glngOtherBlkColor)
                    
                    .Row = RecChangeData!λ��
                    For intCol = 0 To .Cols - 1
                        .Col = intCol
                        .CellBackColor = lngColor
                    Next
                End With
            End If
            
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
    End With
    Billδ��ҩ�嵥.Redraw = True
    
    '���ò˵������߰�ť��״̬
    Call SetMenuAndToolbarState
    
End Sub

Private Sub Lack_Click()
    Call UpdateRsByMenu(Lack, 1)
End Sub

Private Sub mnuEditHandbackBatch_Click()
    TimerAuto.Enabled = False
    If Not frm������ҩ.ShowEditor(Me, lngҩ��ID, False, int����λ��) Then Exit Sub
    mnuViewRefresh_Click
    
    DoEvents
    TimerAuto.Enabled = True
End Sub
Private Sub mnuFilePrintTotal_Click()
    Dim strҩ�� As String, str���� As String
    Dim rsTmp As New adodb.Recordset
    Dim str��ʾ As String
    Dim n As Integer
    
    On Error GoTo errHandle
    gstrSQL = "Select ����,���� From ���ű� Where ID=[1]"
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[��ȡ��ǰҩ��������]", lngҩ��ID)
    
    If Not rsTmp.RecordCount <= 0 Then strҩ�� = "(" & rsTmp!���� & ")" & rsTmp!����
    
    str��ʾ = ""
    If InStr(mstr����, ",") > 0 Then
        gstrSQL = "Select ID,���� From ���ű� Where ID In(" & mstr���� & ") Order by ����"
        Call zldatabase.OpenRecordset(rsTmp, gstrSQL, "��ȡ��������")
    Else
        gstrSQL = "Select ID,���� From ���ű� Where ID = [1] Order by ����"
        Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "��ȡ��������", mstr����)
    End If
    
    If Not rsTmp.RecordCount <= 0 Then
        For n = 1 To rsTmp.RecordCount
            str��ʾ = str��ʾ & "," & rsTmp!����
            rsTmp.MoveNext
        Next
    End If
    
    str��ʾ = Mid(str��ʾ, 2)
    
    Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1342", Me, _
        "��ҩ�ⷿ=" & strҩ�� & "|" & lngҩ��ID, _
        "��������=" & IIf(mint���� = 0, "�ٴ�����", IIf(mint���� = 1, "ҽ������", "���˲���")) & "|" & mint����, _
        "��ҩ����=" & str��ʾ & "|" & " IN (" & mstr���� & ")", "��װϵ��=" & IIf(strUnit = "���ﵥλ", "S.�����װ", "S.סԺ��װ"), "ReportFormat=" & IIf(cbo��ҩ����ʽ.ListIndex = -1, 1, cbo��ҩ����ʽ.ListIndex + 1))
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuFileRestore_Click()
'���ܣ���ӡ��ҩ֪ͨ��
    Dim StrDate As String
    
    If Trim(Bill�ѷ�ҩ�嵥.TextMatrix(Bill�ѷ�ҩ�嵥.Row, ����_�ѷ�ҩ�嵥.ҩƷID)) = "" Then Exit Sub
    With RecChangeSendedData
        If .RecordCount <> 0 Then
            .MoveFirst
            .Find "λ��=" & Bill�ѷ�ҩ�嵥.Row
        End If
        If .EOF Then Exit Sub
        If !�ɲ��� <> 3 Then Exit Sub
        StrDate = Format(!��ҩʱ��, "yyyy-MM-dd HH:mm:ss")
    End With
    
    Call ReportOpen(gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "ZL1_BILL_1342_1", "ZL8_BILL_1342_1"), Me, "��ҩʱ��=" & StrDate, "��װϵ��=" & IIf(strUnit = "���ﵥλ", "C.�����װ", "C.סԺ��װ"), 2)
End Sub

Private Sub mnuFileWait_Click()
    Dim rsTmp As New adodb.Recordset
    Dim str��ʾ As String, str�� As String
    Dim strҩ�� As String, i As Long
    Dim n As Integer
    
    On Error GoTo errHandle
    If glngSys \ 100 = 1 Then
        '�ⷿ����
        gstrSQL = "Select ���� From ���ű� Where ID=[1]"
        Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[��ȡ��ǰҩ��������]", lngҩ��ID)
        
        strҩ�� = rsTmp!���� & "|" & lngҩ��ID
            
        str��ʾ = ""
        If InStr(mstr����, ",") > 0 Then
            gstrSQL = "Select ID,���� From ���ű� Where ID In(" & mstr���� & ") Order by ����"
            Call zldatabase.OpenRecordset(rsTmp, gstrSQL, "��ȡ��������")
        Else
            gstrSQL = "Select ID,���� From ���ű� Where ID = [1] Order by ����"
            Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "��ȡ��������", mstr����)
        End If
        If Not rsTmp.RecordCount <= 0 Then
            For n = 1 To rsTmp.RecordCount
                str��ʾ = str��ʾ & "," & rsTmp!����
                rsTmp.MoveNext
            Next
        End If
        str��ʾ = Mid(str��ʾ, 2)
        str�� = mstr����
    
        Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1342_1", Me, _
            "סԺҩ��=" & strҩ��, "סԺ����=" & str��ʾ & "|" & " IN (" & str�� & ")", _
            "��ʼʱ��=" & mstr��ʼ����_δ��, "����ʱ��=" & mstr��������_δ��, 1)

    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuFlag_Click()
    Dim frmFlag As New Frm���ٷ�ҩ������־
    
    TimerAuto.Enabled = False
    BlnRefresh = False
    
    frmFlag.gstrParentName = Me.Name
    frmFlag.Show vbModal
    
    If BlnRefresh Then
        Call mnuViewRefresh_Click
    End If
    
    DoEvents
    TimerAuto.Enabled = True
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hWnd)
End Sub

Private Sub mnuPassSpec_Click(Index As Integer)
    '���ܣ�ִ��ר��PASS����
    'Pass
    Select Case Index
    Case 0 'ҩ��-ҩ���໥����
        Call PassDoCommand(201)
    Case 1 'ҩ��-ʳ���໥ʹ��
        Call PassDoCommand(202)
    Case 3 '����ע�������
        Call PassDoCommand(203)
    Case 4 '����ע�������
        Call PassDoCommand(204)
    Case 6 '����֢
        Call PassDoCommand(205)
    Case 7 '������
        Call PassDoCommand(206)
    Case 9 '��������ҩ
        Call PassDoCommand(207)
    Case 10 '��ͯ��ҩ
        Call PassDoCommand(208)
    Case 11 '��������ҩ
        Call PassDoCommand(209)
    Case 12 '��������ҩ
        Call PassDoCommand(210)
    End Select
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    'Ĭ�ϲ�����ҩƷ=ҩƷid��ҩ��=ҩ��id������ID=����id��סԺ��=סԺ�ţ�NO=����NO����������=ҩƷ�շ���¼.����
    Dim lngҩƷID As Long
    
    If TabShow.Tab = 0 Then
        If Billδ��ҩ�嵥.Row > 0 Then
            If Val(Billδ��ҩ�嵥.TextMatrix(Billδ��ҩ�嵥.Row, ����_δ��ҩ�嵥.ҩƷID)) > 0 Then
                lngҩƷID = Val(Billδ��ҩ�嵥.TextMatrix(Billδ��ҩ�嵥.Row, ����_δ��ҩ�嵥.ҩƷID))
            End If
        End If
    ElseIf TabShow.Tab = 4 Then
        If Bill�ѷ�ҩ�嵥.Row > 0 Then
            If Val(Bill�ѷ�ҩ�嵥.TextMatrix(Bill�ѷ�ҩ�嵥.Row, ����_�ѷ�ҩ�嵥.ҩƷID)) > 0 Then
                lngҩƷID = Val(Bill�ѷ�ҩ�嵥.TextMatrix(Bill�ѷ�ҩ�嵥.Row, ����_�ѷ�ҩ�嵥.ҩƷID))
            End If
        End If
    End If
    
    Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
        "ҩƷ=" & IIf(lngҩƷID = 0, "", lngҩƷID), _
        "ҩ��=" & IIf(lngҩ��ID = 0, "", lngҩ��ID), _
        "����ID=" & IIf(mlng����ID = 0, "", mlng����ID), _
        "סԺ��=" & mstrסԺ��, _
        "NO=" & mstrNo, _
        "��������=" & IIf(mInt���� = 0, "", mInt����))
        
End Sub

Private Sub mnuReVerify_Click()
    TimerAuto.Enabled = False
    BlnRefresh = False
    
    FrmҩƷ����.ShowForm Me, lngҩ��ID, strUnit, intҩƷ����, int����λ��
    
    If BlnRefresh Then
        Call mnuViewRefresh_Click
    End If
    
    DoEvents
    TimerAuto.Enabled = True
End Sub

Private Sub mnuTypeItem_Click(Index As Integer)
    Dim n As Integer
    Dim strType As String
    
    With mnuTypeItem
        .Item(Index).Checked = Not .Item(Index).Checked
        For n = 0 To .Count - 1
            If .Item(n).Checked = True Then
                strType = strType & ";" & .Item(n).Caption & ";"
            End If
        Next
    End With
    
    With Lvw��ҩ;��
        For n = 1 To .ListItems.Count
            If InStr(1, strType, ";" & .ListItems(n).Tag & ";") > 0 Then
                .ListItems(n).Checked = True
            Else
                .ListItems(n).Checked = False
            End If
        Next
    End With
End Sub

Private Sub mnuViewFontSet_Click(Index As Integer)
    Dim i As Integer
    For i = 0 To 2
        Me.mnuViewFontSet(i).Checked = False
    Next
    Me.mnuViewFontSet(Index).Checked = True
    
    Me.Billδ��ҩ�嵥.Font.Size = IIf(Index = 0, 9, IIf(Index = 1, 11, 15))
    Me.Billȱҩ�嵥.Font.Size = IIf(Index = 0, 9, IIf(Index = 1, 11, 15))
    Me.Bill�ѷ�ҩ�嵥.Font.Size = IIf(Index = 0, 9, IIf(Index = 1, 11, 15))
    Me.Bill�ܷ�ҩ�嵥.Font.Size = IIf(Index = 0, 9, IIf(Index = 1, 11, 15))
    Me.Bill���ܷ�ҩ.Font.Size = IIf(Index = 0, 9, IIf(Index = 1, 11, 15))
    
    zldatabase.SetPara "����", Index, glngSys, 1342
    
    Form_Resize
    Me.Refresh
End Sub

Private Sub MnuViewLocate_Click()
    MnuViewLocateNext.Enabled = False
    MnuViewLocateNext.Tag = 0
    TimerAuto.Enabled = False
    strFind = Frm���ŷ�ҩ��λ.ShowME(lngҩ��ID, Me, mstrPrivs)
    If strFind = "" Then
        TimerAuto.Enabled = True
        Exit Sub
    End If
    
    '��ʼ����¼��
    Set Recδ�� = New adodb.Recordset
    Set Rec�ѷ� = New adodb.Recordset
    With Recδ��
        If .State = 1 Then .Close
        .Fields.Append "λ��", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "NO", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "ҩƷID", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    With Rec�ѷ�
        If .State = 1 Then .Close
        .Fields.Append "λ��", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "NO", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "ҩƷID", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    '����δ��ҩƷ��¼
    With RecChangeData
        If .RecordCount <> 0 Then
            .Filter = strFind
            Do While Not .EOF
                Recδ��.AddNew
                Recδ��!λ�� = !λ��
                Recδ��!���� = !����
                Recδ��!���� = !����
                Recδ��!NO = !NO
                Recδ��!���� = !����
                Recδ��!���� = !����
                Recδ��!ҩƷID = !ҩƷID
                Recδ��.Update
                .MoveNext
            Loop
            .Filter = 0
        End If
    End With
    '�����ѷ�ҩƷ��¼
    With RecChangeSendedData
        If .RecordCount <> 0 Then
            .Filter = strFind
            Do While Not .EOF
                Rec�ѷ�.AddNew
                Rec�ѷ�!λ�� = !λ��
                Rec�ѷ�!���� = !����
                Rec�ѷ�!���� = !����
                Rec�ѷ�!NO = !NO
                Rec�ѷ�!���� = !����
                Rec�ѷ�!���� = !����
                Rec�ѷ�!ҩƷID = !ҩƷID
                Rec�ѷ�.Update
                .MoveNext
            Loop
            .Filter = 0
        End If
    End With
    
    Call FindRecord
    
    DoEvents
    TimerAuto.Enabled = True
End Sub

Private Sub MnuViewLocateNext_Click()
    Call FindRecord(False)
End Sub

Private Sub MnuViewNone_Click()
    Call UpdateState(False, False)
End Sub

Private Sub MnuViewTotal_Click()
    Call UpdateState(False, True)
End Sub

Private Sub Nop_1_Click()
    Call UpdateRsByMenu(Nop_1, 1)
End Sub

Private Sub Nop_2_Click()
    Call UpdateRsByMenu(Nop_2, 2)
End Sub

Private Sub Nop_3_Click()
    Call UpdateRsByMenu(Nop_3, 3)
End Sub

Private Sub Nop_ALL_Click()
    Dim Strִ��״̬ As String
    Dim intCol As Integer
    Dim lngColor As Long
    
    'ȫ��������
    Billδ��ҩ�嵥.Redraw = False
    With RecChangeData
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            If !ִ��״̬ <> 0 Then
                !ִ��״̬ = 3
                Strִ��״̬ = IIf(!ִ��״̬ = 0, "ȱҩ", IIf(!ִ��״̬ = 1, "��ҩ", IIf(!ִ��״̬ = 2, "�ܷ�", "������")))
                !״̬ = Strִ��״̬
                .Update
                
                '����ü�¼����䵽�������������
                With Billδ��ҩ�嵥
                    If .rows - 1 >= RecChangeData!λ�� Then .TextMatrix(RecChangeData!λ��, ����_δ��ҩ�嵥.״̬) = Strִ��״̬
                    
                    lngColor = IIf(Strִ��״̬ = "��ҩ", glngSendBlkColor, glngOtherBlkColor)

                    .Row = RecChangeData!λ��
                    For intCol = 0 To .Cols - 1
                        .Col = intCol
                        .CellBackColor = lngColor
                    Next
                End With
            End If
            
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
    End With
    Billδ��ҩ�嵥.Redraw = False
    
    '���ò˵������߰�ť��״̬
    Call SetMenuAndToolbarState
    
End Sub

Private Sub Restore_Click()
    Call UpdateRsByMenu(Restore, 2)
End Sub

Private Sub ResumeDo_Click()
    Call UpdateRsByMenu(ResumeDo, 3)
End Sub

Private Sub Tbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "Preview"
        mnuFilePreView_Click
    Case "Print"
        mnuFilePrint_Click
    Case "Consignment"
        MnuEditVerify_Click
    Case "Desire"
        MnuEditDesire_Click
    Case "Handback"
        MnuEditHandback_Click
    Case "Restore"
        MnuEditRestore_Click
    Case "ReVerify"
        mnuReVerify_Click
    Case "Help"
        mnuHelpTitle_Click
    Case "Exit"
        mnufileexit_Click
    End Select
End Sub

Private Sub Cbar_Resize()
    Form_Resize
End Sub

Private Sub MnuEditDesire_Click()
    '
End Sub

Private Sub MnuEditHandback_Click()
    Dim IntSet As Integer
    On Error GoTo ErrHand
    '��ҩƷID˳�����
    
    '�Ȼָ��ܷ�ҩ������¼Ϊ������¼
    gcnOracle.BeginTrans
    With Bill�ܷ�ҩ�嵥
        For IntSet = 1 To .rows - 1
            If Trim(.TextMatrix(IntSet, 1)) = "�ָ�" Then
                gstrSQL = "zl_ҩƷ�շ���¼_���Żָ�(" & .RowData(IntSet) & ")"
                Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-�ָ��ܷ�ҩƷ")
            End If
        Next
    End With
    
    '�����û����õ�ǰ������¼Ϊ�ܷ�ҩ
    With RecChangeData
        If .RecordCount <> 0 Then
            .MoveFirst
            .Sort = "ҩƷID Asc"
        End If
        Do While Not .EOF
            If !ִ��״̬ = 2 Then
                If CheckBill(0, !Id) <> 0 Then gcnOracle.RollbackTrans: Exit Sub
                gstrSQL = "zl_ҩƷ�շ���¼_���žܷ�(" & !Id & ")"
                Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-���þܷ�ҩƷ")
            End If
            .MoveNext
        Loop
    End With
    
    'ˢ��
    gcnOracle.CommitTrans
    Set RecRefreshCompare = CopyNewRec(RecChangeData)
    mnuViewRefresh_Click
    Call InitRefreshRec
    Exit Sub
ErrHand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    If RecChangeData.RecordCount <> 0 Then RecChangeData.Sort = "NO Asc"
End Sub

Private Sub MnuEditRestore_Click()
    Dim StrDate As String
    Dim lng���� As Long, lng���� As Long, lngRow As Long
    Dim strShow As String, strReturn As String, blnInput As Boolean, strSubSql As String
    Dim sig��ҩ�� As Single
    Dim RecRecord As New adodb.Recordset
    Dim rsTemp As New adodb.Recordset
    Dim bln�Ƿ�����ҩ As Boolean
    Dim strҩƷid As String
     
    On Error GoTo ErrHand
    
    If TxtInput.Visible Then
        Call TxtInput_LostFocus
        TxtInput.Visible = False
    End If
    
    '��ҩƷID˳�����
    StrDate = Format(zldatabase.Currentdate(), "yyyy-MM-dd HH:mm:ss")
        
    With RecChangeSendedData
        If .RecordCount <> 0 Then .MoveFirst
        If .RecordCount = 0 Then Exit Sub
        If .EOF Then Exit Sub
        
        Call BuildRecord(False)
        If Not CheckCorrelation Then Exit Sub
        
        If MsgBox("��ȷ��Ҫ��ҩ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        
        '��ҩ��ǩ��
        str��ҩ�� = ""
        If Lng��ҩ��ǩ�� = 1 Then
            str��ҩ�� = zldatabase.UserIdentify(Me, "��ҩ��ǩ��", glngSys, 1342, "��ҩ")
            If str��ҩ�� = "" Then
                Exit Sub
            End If
        End If
        
        .Sort = "ҩƷID Asc"
        
        Do While Not .EOF
            If !ִ��״̬ = 3 Then
                '�ȼ���Ƿ�������ҩ��ҽ����
                If blnҽ������ = False Then
                    gstrSQL = "select ���� From ҩƷ�շ���¼ Where ID=[1]"
                    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[����Ƿ�������]", CLng(!Id))
                    
                    If (rsTemp!���� Like "1*") Then       '����
                        gstrSQL = "Select Nvl(ҽ�����,0) ҽ�����,Nvl(�����־,1) �����־ From ���˷��ü�¼ Where ID=(Select ����ID From ҩƷ�շ���¼ Where ID=[1])"
                        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[����Ƿ���ҽ��]", CLng(!Id))
                        
                        If Not rsTemp.EOF Then
                            If (rsTemp!�����־ = 1 Or rsTemp!�����־ = 4) And rsTemp!ҽ����� <> 0 Then
                                gstrSQL = "Select decode(ҽ��״̬,4,1,0) ���� From ����ҽ����¼ Where ID=[1]"
                                Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[�жϸ�ҽ���Ƿ�����]", CLng(rsTemp!ҽ�����))
                                
                                If rsTemp!���� = 0 Then
                                    MsgBox "��" & !λ�� & "��ҩƷ��Ӧ��ҽ����δ���ϣ�������ҩ��", vbInformation, gstrSysName
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                End If
                
                lngRow = !λ��
                lng���� = IIf(IsNull(!����), 0, !����)
                lng���� = IIf(IsNull(!����), 0, !����)
                '���ԭ�������������ڷ���
                If lng���� = 0 And lng���� = 1 Then
                    '������Ż�Ч��Ϊ�գ�����ȡ���û�����
                    blnInput = IIf(IsNull(!����), True, False)
                    If Not blnInput Then blnInput = (Trim(!����) = "")
                    If blnInput Then
                        strShow = Bill�ѷ�ҩ�嵥.TextMatrix(lngRow, ����_�ѷ�ҩ�嵥.����) & "|" & Bill�ѷ�ҩ�嵥.TextMatrix(lngRow, ����_�ѷ�ҩ�嵥.����) & _
                        "|" & Bill�ѷ�ҩ�嵥.TextMatrix(lngRow, ����_�ѷ�ҩ�嵥.����) & "|" & Bill�ѷ�ҩ�嵥.TextMatrix(lngRow, ����_�ѷ�ҩ�嵥.ҩƷ����) & "|" & !ҩƷID
                        strReturn = Frm��ҩ����.ShowME(Me, strShow)
                        If strReturn = "" Then Exit Sub
                        '�������š�Ч�ڼ�����
                        !���� = Split(strReturn, "|")(0)
                        !Ч�� = Split(strReturn, "|")(1)
                        !���� = Split(strReturn, "|")(2)
                        .Update
                    End If
                End If
            End If
            .MoveNext
        Loop
        .MoveFirst
        
        gcnOracle.BeginTrans
        Do While Not .EOF
            If !ִ��״̬ = 3 Then
                If CheckBill(2, !Id) <> 0 Then gcnOracle.RollbackTrans: Exit Sub
                
                'modified.by.zyb ���ﵥλ��סԺ��λ��һ��ʱ����ҩδ���� 2003-01-10
                Select Case strUnit
                Case "�ۼ۵�λ"
                    strSubSql = "*1"
                Case "���ﵥλ"
                    strSubSql = "*Decode(�����װ,Null,1,0,1,�����װ)"
                Case "סԺ��λ"
                    strSubSql = "*Decode(סԺ��װ,Null,1,0,1,סԺ��װ)"
                Case "ҩ�ⵥλ"
                    strSubSql = "*Decode(ҩ���װ,Null,1,0,1,ҩ���װ)"
                End Select
                    
                sig��ҩ�� = !��ҩ��
                
                gstrSQL = " Select round(" & sig��ҩ�� & strSubSql & ",5) ���� From ҩƷ���" & _
                         " Where ҩƷID=[1]"
                Set RecRecord = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, CLng(RecChangeSendedData!ҩƷID))
                
                With RecRecord
                    sig��ҩ�� = !����
                End With
                If Val(!׼����) = Val(!��ҩ��) Then
                    sig��ҩ�� = Val(!ʵ������)
                End If
                
                If sig��ҩ�� <> 0 Then
                    If CheckPrice(!Id, mstr�۸�ʧЧ��ʾ) = False Then
                        If MsgBox("ҩƷ[" & !Ʒ�� & "(" & !��� & ")]" & mstr�۸�ʧЧ��ʾ, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                            gstrSQL = "zl_ҩƷ�շ���¼_������ҩ(" & !Id & ",'" & gstrUserName & "',To_Date('" & StrDate & "','yyyy-MM-dd hh24:mi:ss')," & _
                            IIf(IsNull(!����), "NULL", IIf(Mid(!����, 1, 1) = "(", "NULL", "'" & Mid(!����, 1, 8) & "'")) & "," & _
                            IIf(IsNull(!Ч��), "NULL", IIf(!Ч�� = "", "NULL", "To_Date('" & Format(!Ч��, "yyyy-MM-dd") & "','yyyy-MM-dd')")) & "," & _
                            IIf(IsNull(!����), "NULL", "'" & !���� & "'") & "," & sig��ҩ�� & ",NULL,'" & str��ҩ�� & "'," & int����λ�� & ")"
                            Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-ҩƷ��ҩ")
                            bln�Ƿ�����ҩ = True
                            
                            If InStr("," & strҩƷid & ",", "," & !ҩƷID & ",") = 0 Then
                                strҩƷid = IIf(strҩƷid = "", "", strҩƷid & ",") & !ҩƷID
                            End If
                        End If
                    Else
                        gstrSQL = "zl_ҩƷ�շ���¼_������ҩ(" & !Id & ",'" & gstrUserName & "',To_Date('" & StrDate & "','yyyy-MM-dd hh24:mi:ss')," & _
                        IIf(IsNull(!����), "NULL", IIf(Mid(!����, 1, 1) = "(", "NULL", "'" & Mid(!����, 1, 8) & "'")) & "," & _
                        IIf(IsNull(!Ч��), "NULL", IIf(!Ч�� = "", "NULL", "To_Date('" & Format(!Ч��, "yyyy-MM-dd") & "','yyyy-MM-dd')")) & "," & _
                        IIf(IsNull(!����), "NULL", "'" & !���� & "'") & "," & sig��ҩ�� & ",NULL,'" & str��ҩ�� & "'," & int����λ�� & ")"
                        Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-ҩƷ��ҩ")
                        bln�Ƿ�����ҩ = True
                        
                        If InStr("," & strҩƷid & ",", "," & !ҩƷID & ",") = 0 Then
                            strҩƷid = IIf(strҩƷid = "", "", strҩƷid & ",") & !ҩƷID
                        End If
                    End If
                End If
            End If
            .MoveNext
        Loop
    End With
    
    gcnOracle.CommitTrans
    
    '��ӡ��ҩ��
    If bln�Ƿ�����ҩ = True Then
        If MsgBox("����Ҫ��ӡ��ҩ�嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Call ReportOpen(gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "ZL1_BILL_1342_1", "ZL8_BILL_1342_1"), Me, "��ҩʱ��=" & StrDate, "��װϵ��=" & IIf(strUnit = "���ﵥλ", "C.�����װ", "C.סԺ��װ"), 2)
        End If
        
        '��ʾͣ��ҩƷ
        If strҩƷid <> "" Then
            Call CheckStopMedi(strҩƷid)
        End If
    Else
        MsgBox "����û����ҩ��"
        Exit Sub
    End If
    
    'ˢ��
    mnuViewRefresh_Click
    Exit Sub
ErrHand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    If RecChangeSendedData.RecordCount <> 0 Then RecChangeSendedData.Sort = "NO Asc"
End Sub

Private Sub MnuEditVerify_Click()
    Dim StrCurDate As String
    Dim LngLocate As Long
    Dim strRecipeKey As String              '���汾�η�ҩ������ID
    Dim blnUpdate As Boolean
    Dim str��ʾ As String
    Dim n As Integer
    Dim rsTmp As New adodb.Recordset
    Dim str�ڼ� As String
    Dim lngPatId As Long
    Dim blnBeginTrans As Boolean
    Dim strId���� As String
    Dim strID As String
    Dim strDept As String
    Dim lngPre����id As Long
    Dim strPreNo As String
    Dim lngPre������� As Long
    Dim dblSum As Double
    Dim strҩƷid As String
    Dim dbl�������� As Double
    Dim dblPrice As Double
    Dim strSubSql As String
    Dim RecRecord As adodb.Recordset
        
    On Error GoTo ErrHand
    
    If txt������.Visible Then
        txt������_LostFocus
        txt������.Visible = False
    End If
    
    mlng���ܷ�ҩ�� = Val(zldatabase.GetNextNo(20))
    
    '������ID��������
    StrCurDate = Format(zldatabase.Currentdate(), "yyyy-MM-dd HH:mm:ss")
    str�ڼ� = Format(StrCurDate, "yyyy")
    
    '���洢�ⷿ
    If CheckDrugStock = False Then Exit Sub
    
    '���´���ҩ------------------------------------------------------------------------------------------------------
    With RecChangeData
        If .RecordCount <> 0 Then .MoveFirst
        If .RecordCount = 0 Then Exit Sub
        If .EOF Then Exit Sub

        Call BuildRecord(True)
        If Not CheckCorrelation Then Exit Sub
        
        If MsgBox("��ȷ��Ҫ��ҩ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        
        '��ҩ��ǩ��
        TimerAuto.Enabled = False
        str��ҩ�� = ""
        If Lng��ҩ��ǩ�� = 1 Then
            str��ҩ�� = zldatabase.UserIdentify(Me, "��ҩ��ǩ��", glngSys, 1342, "")
            If str��ҩ�� = "" Then
                TimerAuto.Enabled = True
                Exit Sub
            End If
        End If
        TimerAuto.Enabled = True
        
        '���밴����ID��ҩƷID����
        .Sort = "����ID Asc ,ҩƷID Asc"
        
        Do While Not .EOF
            'ִ��״̬Ϊ1����ͨ�����ݼ��ſ��Է���
            If !ִ��״̬ = 1 And CheckBill(1, !Id) = 0 And CheckGroupSend(!���ID) = True Then
                If lngPatId = 0 Then
                    lngPatId = !����ID
                End If
                
                '����ID��ͬʱ��
                If lngPatId = !����ID Then
                    '���������ַ�������3950ʱ���ύ��������ַ���Ϊ4000��
                    If LenB(strId����) > 3950 Then
                        gcnOracle.BeginTrans
                        blnBeginTrans = True
                        
                        gstrSQL = "Zl_ҩƷ�շ���¼_������ҩ('" & strId���� & "'," & lngҩ��ID & ",'" & gstrUserName & "',To_Date('" & StrCurDate & "','yyyy-MM-dd hh24:mi:ss') ,3,'" & str��ҩ�� & "'," & mlng���ܷ�ҩ�� & "," & int����λ�� & ",'" & NeedName(cbo��ҩ��.Text) & "') "
                        gcnOracle.BeginTrans
                        Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-����������ҩ")
                                                
                        If int��˻��۵� = 1 Then
                            gstrSQL = "Zl_סԺ���ʼ�¼_��ҩ���('" & strID & "','" & gstrUserCode & "','" & gstrUserName & "',To_Date('" & StrCurDate & "','yyyy-MM-dd hh24:mi:ss'))"
                            Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-סԺ�������")
                        End If
                        gcnOracle.CommitTrans
                                                                    
                        blnBeginTrans = False
                        blnUpdate = True
                        lngPatId = 0
                        strId���� = !Id & "," & NVL(!����, 0)
                        strID = !Id
                    Else
                        strId���� = IIf(strId���� = "", !Id & "," & NVL(!����, 0), strId���� & "|" & !Id & "," & NVL(!����, 0))
                        strID = IIf(strID = "", !Id, strID & "," & !Id)
                    End If
                Else
                    '�������ID��ͬ���ύ����
                    blnBeginTrans = True
                    
                    gstrSQL = "Zl_ҩƷ�շ���¼_������ҩ('" & strId���� & "'," & lngҩ��ID & ",'" & gstrUserName & "',To_Date('" & StrCurDate & "','yyyy-MM-dd hh24:mi:ss') ,3,'" & str��ҩ�� & "'," & mlng���ܷ�ҩ�� & "," & int����λ�� & ",'" & NeedName(cbo��ҩ��.Text) & "') "
                    gcnOracle.BeginTrans
                    Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-����������ҩ")
                                        
                    If int��˻��۵� = 1 Then
                        gstrSQL = "Zl_סԺ���ʼ�¼_��ҩ���('" & strID & "','" & gstrUserCode & "','" & gstrUserName & "',To_Date('" & StrCurDate & "','yyyy-MM-dd hh24:mi:ss'))"
                        Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-סԺ�������")
                    End If
                    gcnOracle.CommitTrans
                    
                    blnBeginTrans = False
                    blnUpdate = True
                    lngPatId = !����ID
                    strId���� = !Id & "," & NVL(!����, 0)
                    strID = !Id
                End If
            End If
            .MoveNext
            
            '�������û�м�¼���Ҵ����ַ�����Ϊ�գ����ύ����
            If .EOF And strId���� <> "" Then
                blnBeginTrans = True
                
                gstrSQL = "Zl_ҩƷ�շ���¼_������ҩ('" & strId���� & "'," & lngҩ��ID & ",'" & gstrUserName & "',To_Date('" & StrCurDate & "','yyyy-MM-dd hh24:mi:ss') ,3,'" & str��ҩ�� & "'," & mlng���ܷ�ҩ�� & "," & int����λ�� & ",'" & NeedName(cbo��ҩ��.Text) & "') "
                gcnOracle.BeginTrans
                Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-����������ҩ")
                                    
                If int��˻��۵� = 1 Then
                    gstrSQL = "Zl_סԺ���ʼ�¼_��ҩ���('" & strID & "','" & gstrUserCode & "','" & gstrUserName & "',To_Date('" & StrCurDate & "','yyyy-MM-dd hh24:mi:ss'))"
                    Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-סԺ�������")
                End If
                gcnOracle.CommitTrans
                blnUpdate = True
                
                blnBeginTrans = False
            End If
        Loop
    End With
    '���ϴ���ҩ----------------------------------------------------------------------------------------------------------
    
    
    '���´���ҩƷ����-------------------------------------------------------------------------------------------------------
    gcnOracle.BeginTrans
    'ǰ�������ǰ����һ���
    If TabShow.Tab = 1 And Lng������ʾ = 1 Then
        For n = 1 To Bill���ܷ�ҩ.rows - 3
            If Val(Bill���ܷ�ҩ.TextMatrix(n, ����_���һ����嵥.��������)) <> 0 Then
                If Not blnҩƷ���������� Then
                    gcnOracle.RollbackTrans
                    MsgBox "û������ҩƷ���������������ҩƷ������࣡���ΰ�ȫʵ������", vbInformation + vbOKOnly, gstrSysName
                    If blnUpdate Then GoTo RefData
                    Exit Sub
                End If
                
                dbl�������� = Val(Bill���ܷ�ҩ.TextMatrix(n, ����_���һ����嵥.��������))
                dblPrice = Val(Bill���ܷ�ҩ.TextMatrix(n, ����_���һ����嵥.����))
                
                Select Case strUnit
                Case "�ۼ۵�λ"
                    strSubSql = "round(" & dbl�������� & ",5) As ����, round(" & dblPrice & ",5) As ����"
                Case "���ﵥλ"
                    strSubSql = "round(" & dbl�������� & " * Decode(�����װ,Null,1,0,1,�����װ) ,5) As ����, round(" & dblPrice & " /Decode(�����װ,Null,1,0,1,�����װ) ,5) As ���� "
                Case "סԺ��λ"
                    strSubSql = "round(" & dbl�������� & " * Decode(סԺ��װ,Null,1,0,1,סԺ��װ) ,5) As ����, round(" & dblPrice & " /Decode(סԺ��װ,Null,1,0,1,סԺ��װ) ,5) As ���� "
                Case "ҩ�ⵥλ"
                    strSubSql = "round(" & dbl�������� & " * Decode(ҩ���װ,Null,1,0,1,ҩ���װ) ,5) As ����, round(" & dblPrice & " /Decode(ҩ���װ,Null,1,0,1,ҩ���װ) ,5) As ���� "
                End Select
                    
                gstrSQL = " Select " & strSubSql & " From ҩƷ���" & _
                         " Where ҩƷID=[1]"
                Set RecRecord = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Bill���ܷ�ҩ.TextMatrix(n, ����_���һ����嵥.ҩƷID)))
                dbl�������� = RecRecord!����
                dblPrice = RecRecord!����
                
                gstrSQL = "ZL_ҩƷ�����¼_INSERT(" & str�ڼ� & "," & mlng���ܷ�ҩ�� & "," & lngҩ��ID & "," & Val(Bill���ܷ�ҩ.TextMatrix(n, ����_���һ����嵥.����ID)) & "," & Val(Bill���ܷ�ҩ.TextMatrix(n, ����_���һ����嵥.ҩƷID)) & "," & Val(Bill���ܷ�ҩ.TextMatrix(n, ����_���һ����嵥.����)) & ", " & _
                " " & dbl�������� & "," & dblPrice & ", '" & gstrUserName & "',To_Date('" & StrCurDate & "','yyyy-MM-dd hh24:mi:ss') ," & Val(Bill���ܷ�ҩ.TextMatrix(n, ����_���һ����嵥.��ҩ����id)) & ") "
                Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-��������")
            End If
        Next
    End If
    gcnOracle.CommitTrans
    '���ϴ���ҩƷ����-------------------------------------------------------------------------------------------------------
    
    
    '���´�����������------------------------------------------------------------------------------------------------------
    Dim i As Integer
    Dim strMCNO As String, arrMCRec As Variant, arrMCPar As Variant
    Dim int��˱�־ As Integer
    Dim bln�Ƿ�����ҩ As Boolean
    Dim str������� As String
    
    'ǰ�������ǻ������ʼ�¼һ����ҩ
    If TabShow.Tab = 1 And mbln���ܷ�ҩ = True Then
        If mrsRequest.State <> 0 Then
            mrsRequest.Filter = ""
            mrsRequest.Sort = "No,����id,�շ�id"
            If mrsRequest.RecordCount > 0 Then
                With mrsRequest
                    gcnOracle.BeginTrans
                    blnBeginTrans = True
                    gclsInsure.InitOracle gcnOracle
                    Do While Not .EOF
                        If !��˱�־ <> 0 Then
                            If lngPre����id <> !����id Then
                                '�������ʼ�¼����
                                gstrSQL = "zl_���˷�������_Audit(" & !����id & ",To_Date('" & !����ʱ�� & "','YYYY-MM-DD HH24:MI:SS'),'" & _
                                               gstrUserName & "',To_Date('" & StrCurDate & "','yyyy-MM-dd hh24:mi:ss')," & !��˱�־ & ")"
                                Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-���²��˷������ʼ�¼")
                                lngPre����id = !����id
                            End If
                        End If
                        
                        '��ҩ����
                        If !��˱�־ = 1 And !�������� <> 0 Then
                            gstrSQL = "zl_ҩƷ�շ���¼_������ҩ(" & !�շ�ID & ",'" & gstrUserName & "',To_Date('" & StrCurDate & "','yyyy-MM-dd hh24:mi:ss')," & _
                                IIf(IsNull(!����), "NULL", IIf(Mid(!����, 1, 1) = "(", "NULL", "'" & Mid(!����, 1, 8) & "'")) & "," & _
                                IIf(IsNull(!Ч��), "NULL", IIf(!Ч�� = "", "NULL", "To_Date('" & Format(!Ч��, "yyyy-MM-dd") & "','yyyy-MM-dd')")) & "," & _
                                IIf(IsNull(!����), "NULL", "'" & !���� & "'") & "," & !�������� & ",NULL,'" & gstrUserName & "'," & int����λ�� & "," & mlng���ܷ�ҩ�� & ")"
            
                            Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-ҩƷ��ҩ����")
                            bln�Ƿ�����ҩ = True
                            
                            If InStr("," & strҩƷid & ",", "," & !ҩƷID & ",") = 0 Then
                                strҩƷid = IIf(strҩƷid = "", "", strҩƷid & ",") & !ҩƷID
                            End If
                        
                            '���ʴ���
                            strPreNo = !NO
                            lngPre������� = !�������
                            dblSum = dblSum + !��������
                            
                            .MoveNext
                            If .EOF Then
                                .MovePrevious
                                str������� = !������� & ":" & dblSum
                
                                gstrSQL = "ZL_סԺ���ʼ�¼_Delete('" & !NO & "','" & str������� & "','" & gstrUserCode & "','" & gstrUserName & "'," & !��¼���� & ")"
                                Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-ɾ�����ʼ�¼")
                
                                'ҽ������
                                If Not IsNull(!����) And InStr(1, strMCNO, !NO) = 0 Then
                                    MCPAR.���������ϴ� = gclsInsure.GetCapability(support���������ϴ�, , Val(!����))
                                    MCPAR.������ɺ��ϴ� = gclsInsure.GetCapability(support������ɺ��ϴ�, , Val(!����))
                                    strMCNO = strMCNO & IIf(strMCNO = "", "", "|") & !NO & "," & !���� & _
                                            "," & IIf(MCPAR.���������ϴ�, "1", "0") & "," & IIf(MCPAR.������ɺ��ϴ�, "1", "0")
                                End If
                                .MoveNext
                            Else
                                If strPreNo <> !NO Or (strPreNo = !NO And lngPre������� <> !�������) Then
                                    .MovePrevious
                                    str������� = !������� & ":" & dblSum
                                    
                                    gstrSQL = "ZL_סԺ���ʼ�¼_Delete('" & !NO & "','" & str������� & "','" & gstrUserCode & "','" & gstrUserName & "'," & !��¼���� & ")"
                                    Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-ɾ�����ʼ�¼")
                    
                                    'ҽ������
                                    If Not IsNull(!����) And InStr(1, strMCNO, !NO) = 0 Then
                                        MCPAR.���������ϴ� = gclsInsure.GetCapability(support���������ϴ�, , Val(!����))
                                        MCPAR.������ɺ��ϴ� = gclsInsure.GetCapability(support������ɺ��ϴ�, , Val(!����))
                                        strMCNO = strMCNO & IIf(strMCNO = "", "", "|") & !NO & "," & !���� & _
                                                "," & IIf(MCPAR.���������ϴ�, "1", "0") & "," & IIf(MCPAR.������ɺ��ϴ�, "1", "0")
                                    End If
                                    
                                    dblSum = 0
                                    .MoveNext
                                End If
                            End If
                            .MovePrevious
                        End If
                        .MoveNext
                    Loop
                End With
            
                'ҽ�������������ϴ�������ʱ�ϴ�
                If strMCNO <> "" Then
                    arrMCRec = Split(strMCNO, "|")
                    For i = 0 To UBound(arrMCRec)
                        arrMCPar = Split(arrMCRec(i), ",")
                        If arrMCPar(2) = 1 And arrMCPar(3) = 0 Then
                            If Not gclsInsure.TranChargeDetail(2, CStr(arrMCPar(0)), 2, 2, "", , Val(arrMCPar(1))) Then
                                gcnOracle.RollbackTrans
                                GoTo RefData
                            End If
                        End If
                    Next
                End If
                                        
                gcnOracle.CommitTrans
                blnBeginTrans = False
                
                'ҽ�������������ϴ�����ɺ��ϴ�
                If strMCNO <> "" Then
                    For i = 0 To UBound(arrMCRec)
                        arrMCPar = Split(arrMCRec(i), ",")
                        If arrMCPar(2) = 1 And arrMCPar(3) = 1 Then
                            If Not gclsInsure.TranChargeDetail(2, CStr(arrMCPar(0)), 2, 2, "", , Val(arrMCPar(1))) Then
                                MsgBox "����""" & CStr(arrMCPar(0)) & """������������ҽ������ʧ�ܣ��õ��������ʡ�", vbInformation, gstrSysName
                            End If
                        End If
                    Next
                End If
                
                If bln�Ƿ�����ҩ = True Then
                    If MsgBox("����Ҫ��ӡ��ҩ�嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                        Call ReportOpen(gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "ZL1_BILL_1342_1", "ZL8_BILL_1342_1"), Me, "��ҩʱ��=" & StrCurDate, "��װϵ��=" & IIf(strUnit = "���ﵥλ", "C.�����װ", "C.סԺ��װ"), 2)
                    End If
                End If
            End If
        End If
    End If
    
    '��ʾͣ��ҩƷ
    If strҩƷid <> "" Then
        Call CheckStopMedi(strҩƷid)
    End If
    
    '���ϴ�����������------------------------------------------------------------------------------------------------------
    
    blnBeginTrans = False
RefData:
    'ˢ��
    If blnUpdate Then
        strDept = mstrDrawDept
        Set RecRefreshCompare = CopyNewRec(RecChangeData)
        mnuViewRefresh_Click
        Call InitRefreshRec
        
        '��ӡ���ܵ���
        If Lng�Զ���ӡ = 1 Then
            str��ʾ = ""
            
            If InStr(strDept, ",") > 0 Then
                gstrSQL = "Select ID,���� From ���ű� Where ID In(" & strDept & ") Order by ����"
                Call zldatabase.OpenRecordset(rsTmp, gstrSQL, "��ȡ��������")
            Else
                gstrSQL = "Select ID,���� From ���ű� Where ID = [1] Order by ����"
                Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "��ȡ��������", strDept)
            End If
            
            If Not rsTmp.RecordCount <= 0 Then
                For n = 1 To rsTmp.RecordCount
                    str��ʾ = str��ʾ & "," & rsTmp!����
                    rsTmp.MoveNext
                Next
            End If
            
            str��ʾ = Mid(str��ʾ, 2)
            
            Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1342", Me, _
                "��ҩ�ⷿ=" & lngҩ��ID, _
                "��������=" & mint����, _
                "��ҩ����=" & str��ʾ & "|" & " IN (" & strDept & ")", _
                "��װϵ��=" & IIf(strUnit = "���ﵥλ", "S.�����װ", "S.סԺ��װ"), _
                "��ҩ��=" & mlng���ܷ�ҩ��, "ReportFormat=" & IIf(cbo��ҩ����ʽ.ListIndex = -1, 1, cbo��ҩ����ʽ.ListIndex + 1), "PrintEmpty=0", 2)

        End If
    End If
    
    blnUpdate = False
        
    Exit Sub
ErrHand:
    '����ѿ������񣬲���δ�ύ�������ʱ�ع�����
    If blnBeginTrans Then
        gcnOracle.RollbackTrans
    End If
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    If RecChangeData.RecordCount <> 0 Then RecChangeData.Sort = "NO Asc"
    If blnUpdate Then GoTo RefData
End Sub



Private Sub MnuFileBillprint_Click()
    '
End Sub

Private Sub mnuFileExcel_Click()
    subPrint 3
End Sub

Private Sub mnufileexit_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub MnuFilePara_Click()
    Dim intFixedCol As Integer
    Dim dateCurDate As Date
    
    BlnSetPara = False
    TimerAuto.Enabled = False
    With Frm���ŷ�ҩ��������
        .strPrivs = mstrPrivs
        .Show 1, Me
    End With
    
    '��ע����ж�ȡ��ز�������
    If BlnSetPara Then
        '���»�ȡע���
        Call ReadFromReg
        
        Call Get��ҩ��
        Call Get����

        '�������û����嵥�ĸ�ʽ
        With Bill���ܷ�ҩ
            .rows = 2
            .Cols = IIf(Lng������ʾ = 1, ����_���һ����嵥.����, ����_�����嵥.����)
        
            If Lng������ʾ = 0 Then
                .TextMatrix(0, ����_�����嵥.ҩƷ����) = "ҩƷ����"
                .TextMatrix(0, ����_�����嵥.���) = "���"
                .TextMatrix(0, ����_�����嵥.����) = "����"
                .TextMatrix(0, ����_�����嵥.����) = "����"
                .TextMatrix(0, ����_�����嵥.����) = "����"
                .TextMatrix(0, ����_�����嵥.��λ) = "��λ"
                .TextMatrix(0, ����_�����嵥.����) = "����"
                .TextMatrix(0, ����_�����嵥.���) = "���"
                            
                .ColWidth(����_�����嵥.ҩƷ����) = 2000
                .ColWidth(����_�����嵥.���) = 1500
                .ColWidth(����_�����嵥.����) = 1500
                .ColWidth(����_�����嵥.����) = 1200
                .ColWidth(����_�����嵥.����) = 1200
                .ColWidth(����_�����嵥.��λ) = 500
                .ColWidth(����_�����嵥.����) = 1200
                .ColWidth(����_�����嵥.���) = 1200
            Else
                .TextMatrix(0, ����_���һ����嵥.����) = "����"
                .TextMatrix(0, ����_���һ����嵥.ҩƷ����) = "ҩƷ����"
                .TextMatrix(0, ����_���һ����嵥.���) = "���"
                .TextMatrix(0, ����_���һ����嵥.����) = "����"
                .TextMatrix(0, ����_���һ����嵥.����) = "����"
                .TextMatrix(0, ����_���һ����嵥.Ӧ������) = "Ӧ������"
                .TextMatrix(0, ����_���һ����嵥.��������) = "��������"
                .TextMatrix(0, ����_���һ����嵥.��������) = "��������"
                .TextMatrix(0, ����_���һ����嵥.ʵ������) = "ʵ������"
                .TextMatrix(0, ����_���һ����嵥.��λ) = "��λ"
                .TextMatrix(0, ����_���һ����嵥.����) = "����"
                .TextMatrix(0, ����_���һ����嵥.���) = "���"
                .TextMatrix(0, ����_���һ����嵥.����) = "����"
                .TextMatrix(0, ����_���һ����嵥.����ID) = "����ID"
                .TextMatrix(0, ����_���һ����嵥.ҩƷID) = "ҩƷID"
                
                .ColWidth(����_���һ����嵥.����) = 1200
                .ColWidth(����_���һ����嵥.ҩƷ����) = 2000
                .ColWidth(����_���һ����嵥.���) = 1500
                .ColWidth(����_���һ����嵥.����) = 1500
                .ColWidth(����_���һ����嵥.����) = 1200
                .ColWidth(����_���һ����嵥.Ӧ������) = 1200
                .ColWidth(����_���һ����嵥.��������) = 1200
                .ColWidth(����_���һ����嵥.��������) = IIf(mbln���ܷ�ҩ = True, 1200, 0)
                .ColWidth(����_���һ����嵥.ʵ������) = 1200
                .ColWidth(����_���һ����嵥.��λ) = 500
                .ColWidth(����_���һ����嵥.����) = 1200
                .ColWidth(����_���һ����嵥.���) = 1200
                .ColWidth(����_���һ����嵥.����) = 0
                .ColWidth(����_���һ����嵥.����ID) = 0
                .ColWidth(����_���һ����嵥.ҩƷID) = 0
            End If
        
            For intFixedCol = 0 To .Cols - 1
                .ColAlignmentFixed(intFixedCol) = 4
            Next
            .ColAlignment(IIf(Lng������ʾ = 1, ����_���һ����嵥.���, ����_�����嵥.���)) = 1
            .ColAlignment(IIf(Lng������ʾ = 1, ����_���һ����嵥.����, ����_�����嵥.����)) = 1
        End With
        Call RestoreFlexState(Bill���ܷ�ҩ, "���ܷ�ҩ" & Lng������ʾ)
        Bill���ܷ�ҩ.ColWidth(����_���һ����嵥.��������) = IIf(mbln���ܷ�ҩ = True, 1200, 0)
        
        'ˢ������
        Call mnuViewRefresh_Click
    End If
    
    Billδ��ҩ�嵥.ColWidth(����_δ��ҩ�嵥.�����) = IIf(Not mblnStarPass, 0, 240)
    Bill�ѷ�ҩ�嵥.ColWidth(����_�ѷ�ҩ�嵥.�����) = IIf(Not mblnStarPass, 0, 240)
End Sub

Private Sub mnuFilePreView_Click()
    subPrint 2
End Sub

Private Sub mnuFilePrint_Click()
    subPrint 1
End Sub

Private Sub mnuFileset_Click()
    zlPrintSet
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub MnuViewState_Click()
    MnuViewState.Checked = MnuViewState.Checked Xor True
    stbThis.Visible = MnuViewState.Checked
    Form_Resize
End Sub

Private Sub MnuViewToolS_Click()
    MnuViewToolS.Checked = MnuViewToolS.Checked Xor True
    Cbar.Visible = MnuViewToolS.Checked
    MnuViewToolT.Enabled = MnuViewToolS.Checked
    
    Form_Resize
End Sub

Private Sub MnuViewToolT_Click()
    MnuViewToolT.Checked = MnuViewToolT.Checked Xor True
    If MnuViewToolT.Checked Then
        Tbar.Buttons("Preview").Caption = "Ԥ��"
        Tbar.Buttons("Print").Caption = "��ӡ"
        Tbar.Buttons("Consignment").Caption = "��ҩ"
        Tbar.Buttons("Desire").Caption = "����"
        Tbar.Buttons("Handback").Caption = "�ܷ�"
        Tbar.Buttons("Restore").Caption = "��ҩ"
        Tbar.Buttons("ReVerify").Caption = "����"
        Tbar.Buttons("Help").Caption = "����"
        Tbar.Buttons("Exit").Caption = "�˳�"
    Else
        Tbar.Buttons("Preview").Caption = ""
        Tbar.Buttons("Print").Caption = ""
        Tbar.Buttons("Consignment").Caption = ""
        Tbar.Buttons("Desire").Caption = ""
        Tbar.Buttons("Handback").Caption = ""
        Tbar.Buttons("Restore").Caption = ""
        Tbar.Buttons("ReVerify").Caption = ""
        Tbar.Buttons("Help").Caption = ""
        Tbar.Buttons("Exit").Caption = ""
    End If
    
    Cbar.Bands(1).MinHeight = Tbar.Height
    Form_Resize
End Sub

Private Sub mnuHelpTitle_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub MnuHelpWebM_Click()
    Call zlMailTo(Me.hWnd)
End Sub

Private Sub mnuViewRefresh_Click()
    If BlnStartUp = False Then Exit Sub
    If BlnInRefresh Then Exit Sub
    
    BlnInRefresh = True
    Blnˢ��δ��ҩ�嵥 = True
    Bln����� = True
    
    '������ҩ����д��ˢ�º���б���
    If TxtInput.Visible Then
        TxtInput.Visible = False
        Call TxtInput_LostFocus
    End If
    
    '����ˢ�º󣬱����������ò�������
    strFind = ""
    MnuViewLocateNext.Enabled = False
    MnuViewLocateNext.Tag = 0
    
    Call AviShow
    ''''��������
    Call SetCondition(IIf(TabShow.Tab = 4, 1, 0))
    Call RefreshData
    Call AviShow(False)
    
    mdate�ϴ�ˢ��ʱ�� = zldatabase.Currentdate
    
    BlnInRefresh = False
End Sub

Private Sub tabShow_Click(PreviousTab As Integer)
    On Error Resume Next
    
    If TabShow.Tab = 4 Then
        mnuBillItem(2).Visible = False
        mnuBillItem(6).Visible = False
        mnuBillItem(23).Visible = False
        mnuBillItem(24).Visible = False
    ElseIf TabShow.Tab = 0 Then
        mnuBillItem(2).Visible = UserPrivDetail.Priv_ҽ����ѯ
        mnuBillItem(6).Visible = True
        mnuBillItem(23).Visible = True
        mnuBillItem(24).Visible = True
    End If
    If TabShow.Tab = 0 Or TabShow.Tab = 4 Then
        cmdAlley.Visible = mblnStarPass
    Else
        cmdAlley.Visible = False
    End If
    
    If TabShow.Tab = 0 Then
        Chk��ʾ��ҩ��������.Visible = True
    Else
        Chk��ʾ��ҩ��������.Visible = False
    End If
    
    Cbo����.Visible = False
    
    '��ҳ��ʱ�򱣴�ͻָ���������
    If mintLastTab <> TabShow.Tab Then
        '�����ϸ�ҳ������
        Call SetCondition(IIf(mintLastTab = 4, 1, 0))
        
        Call SaveCondition(IIf(mintLastTab = 4, 1, 0))
        
        '�ָ���ǰҳ������
        Call LoadCondition(IIf(TabShow.Tab = 4, 1, 0))
        
        Call SetCondition(IIf(TabShow.Tab = 4, 1, 0))
    End If
    mintLastTab = TabShow.Tab
    
    '�����δ��ҩ�嵥����ҩ�嵥���򱣴�������
    Dim strTag As String
    If Billδ��ҩ�嵥.ColWidth(����_δ��ҩ�嵥.״̬) < 200 Then Billδ��ҩ�嵥.ColWidth(����_δ��ҩ�嵥.״̬) = 700
    If Billδ��ҩ�嵥.ColWidth(����_δ��ҩ�嵥.����) < 200 Then Billδ��ҩ�嵥.ColWidth(����_δ��ҩ�嵥.����) = 1500
    If Bill�ѷ�ҩ�嵥.ColWidth(����_�ѷ�ҩ�嵥.״̬) < 200 Then Bill�ѷ�ҩ�嵥.ColWidth(����_�ѷ�ҩ�嵥.״̬) = 700
    If Bill�ѷ�ҩ�嵥.ColWidth(����_�ѷ�ҩ�嵥.��ҩ��) < 200 Then Bill�ѷ�ҩ�嵥.ColWidth(����_�ѷ�ҩ�嵥.��ҩ��) = 1000
    
    Bill���ܷ�ҩ.ColWidth(����_���һ����嵥.��������) = IIf(mbln���ܷ�ҩ = True, 1200, 0)
    
    If PreviousTab = 0 Then
        strTag = Billδ��ҩ�嵥.Tag: Billδ��ҩ�嵥.Tag = ""
        Call SaveFlexState(Billδ��ҩ�嵥, "δ��ҩ�嵥")
        Billδ��ҩ�嵥.Tag = strTag
    End If
    If PreviousTab = 4 Then
        strTag = Bill�ѷ�ҩ�嵥.Tag: Bill�ѷ�ҩ�嵥.Tag = ""
        Call SaveFlexState(Bill�ѷ�ҩ�嵥, "�ѷ�ҩ�嵥")
        Bill�ѷ�ҩ�嵥.Tag = strTag
    End If
    If PreviousTab = 1 Then
        strTag = Bill���ܷ�ҩ.Tag: Bill���ܷ�ҩ.Tag = ""
        Call SaveFlexState(Bill���ܷ�ҩ, "���ܷ�ҩ" & Lng������ʾ)
        Bill���ܷ�ҩ.Tag = strTag
    End If
    If PreviousTab = 2 Then
        strTag = Billȱҩ�嵥.Tag: Billȱҩ�嵥.Tag = ""
        Call SaveFlexState(Billȱҩ�嵥, "ȱҩ�嵥")
        Billȱҩ�嵥.Tag = strTag
    End If
    If PreviousTab = 3 Then
        strTag = Bill�ܷ�ҩ�嵥.Tag: Bill�ܷ�ҩ�嵥.Tag = ""
        Call SaveFlexState(Bill�ܷ�ҩ�嵥, "�ܷ�ҩ�嵥")
        Bill�ܷ�ҩ�嵥.Tag = strTag
    End If
    
    Chk�嵥.Visible = (TabShow.Tab = 1 Or TabShow.Tab = 4)
    Chk�嵥.Enabled = Chk�嵥.Visible
    Chk�嵥.Caption = IIf(TabShow.Tab = 4, "��ʾ���й��̵���", "��ҩƷ���λ���")
    If Chk�嵥.Tag <> "" Then
        If Chk�嵥.Value <> IIf(TabShow.Tab = 1, Mid(Chk�嵥.Tag, 1, 1), Mid(Chk�嵥.Tag, 2, 1)) Then
            Chk�嵥.Value = IIf(TabShow.Tab = 1, Mid(Chk�嵥.Tag, 1, 1), Mid(Chk�嵥.Tag, 2, 1))
            Exit Sub
        End If
    Else
        Chk�嵥.Tag = "00"
    End If
    
    Call RefreshDataBaseOnPage
    
    MnuViewTotal.Enabled = (TabShow.Tab = 4)
    MnuViewNone.Enabled = (TabShow.Tab = 4)
    MnuViewLocate.Enabled = (TabShow.Tab = 0 Or TabShow.Tab = 4)
    MnuViewLocateNext.Enabled = (MnuViewLocate.Enabled And Val(MnuViewLocateNext.Tag))
    
    mnuFileRestore = False
    Select Case TabShow.Tab
    Case 0
        TxtInput.Visible = False
        Billδ��ҩ�嵥.Col = ����_δ��ҩ�嵥.״̬
        Billδ��ҩ�嵥.SetFocus
        Call SetMenu(Trim(Billδ��ҩ�嵥.TextMatrix(1, ����_δ��ҩ�嵥.����)) <> "")
    Case 1
        TxtInput.Visible = False
        Bill���ܷ�ҩ.SetFocus
        Call SetMenu(Trim(Bill���ܷ�ҩ.TextMatrix(1, 0)) <> "")
    Case 2
        TxtInput.Visible = False
        Billȱҩ�嵥.SetFocus
        Call SetMenu(Trim(Billȱҩ�嵥.TextMatrix(1, 0)) <> "")
    Case 3
        TxtInput.Visible = False
        Bill�ܷ�ҩ�嵥.SetFocus
        Call SetMenu(Trim(Bill�ܷ�ҩ�嵥.TextMatrix(1, 0)) <> "")
    Case 4
        If Not BlnInRefresh Then
            '�ѷ�������¼��
            Set RecChangeSendedData = New adodb.Recordset
            With RecChangeSendedData
                If .State = 1 Then .Close
                .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
                .Fields.Append "����", adLongVarChar, 20, adFldIsNullable
                .Fields.Append "ID", adDouble, 18, adFldIsNullable
                .Fields.Append "ҩƷID", adDouble, 18, adFldIsNullable
                .Fields.Append "ִ��״̬", adDouble, 1, adFldIsNullable
                .Fields.Append "NO", adLongVarChar, 8, adFldIsNullable
                .Fields.Append "����", adDouble, 18, adFldIsNullable
                .Fields.Append "����ID", adDouble, 18, adFldIsNullable
                .Fields.Append "���", adDouble, 18, adFldIsNullable
                .Fields.Append "����", adLongVarChar, 10, adFldIsNullable
                .Fields.Append "����", adLongVarChar, 20, adFldIsNullable
                .Fields.Append "סԺ��", adLongVarChar, 20, adFldIsNullable
                .Fields.Append "Ʒ��", adLongVarChar, 50, adFldIsNullable
                .Fields.Append "���", adLongVarChar, 50, adFldIsNullable
                .Fields.Append "�������", adLongVarChar, 10, adFldIsNullable
                .Fields.Append "����", adDouble, 18, adFldIsNullable
                .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
                .Fields.Append "Ч��", adLongVarChar, 50, adFldIsNullable
                .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
                .Fields.Append "����", adDouble, 2, adFldIsNullable
                .Fields.Append "��", adDouble, 18, adFldIsNullable
                .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
                .Fields.Append "������", adLongVarChar, 50, adFldIsNullable
                .Fields.Append "׼����", adLongVarChar, 50, adFldIsNullable
                .Fields.Append "��ҩ��", adDouble, 18, adFldIsNullable
                .Fields.Append "�ɲ���", adDouble, 2, adFldIsNullable
                .Fields.Append "��λ", adLongVarChar, 50, adFldIsNullable
                .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
                .Fields.Append "���", adLongVarChar, 50, adFldIsNullable
                .Fields.Append "����", adLongVarChar, 20, adFldIsNullable
                .Fields.Append "Ƶ��", adLongVarChar, 20, adFldIsNullable
                .Fields.Append "�÷�", adLongVarChar, 30, adFldIsNullable
                .Fields.Append "˵��", adLongVarChar, 40, adFldIsNullable
                .Fields.Append "����Ա", adLongVarChar, 20, adFldIsNullable
                .Fields.Append "��ҩʱ��", adLongVarChar, 40, adFldIsNullable
                .Fields.Append "λ��", adDouble, 18, adFldIsNullable
                
                .CursorLocation = adUseClient
                .CursorType = adOpenStatic
                .LockType = adLockOptimistic
                .Open
            End With
        End If
        
        Call SetCondition(1)
        If mblnFirstSended = False Then
            Call zlCommFun.ShowFlash
            If RefreshSendedData = False Then Call zlCommFun.StopFlash: Exit Sub
            Bill�ѷ�ҩ�嵥.Col = 1
            Bill�ѷ�ҩ�嵥.SetFocus
            Call SetMenu(Trim(Bill�ѷ�ҩ�嵥.TextMatrix(1, ����_�ѷ�ҩ�嵥.����)) <> "")
            Call zlCommFun.StopFlash
        End If
'        mblnFirstSended = False
    End Select
    
    If TabShow.Tab = 0 Then
        Call SetColHide(Billδ��ҩ�嵥)
    ElseIf TabShow.Tab = 4 Then
        Call SetColHide(Bill�ѷ�ҩ�嵥)
    End If
End Sub

Private Function GetDetailCol(ByVal strText As String, ByVal Bill As MSHFlexGrid) As Integer
    Dim intCol As Integer, intCols As Integer
    intCols = Bill.Cols - 1
    If strText = "����" Then strText = "����"
    If strText = "��/��ҩ��" Then
        If TabShow.Tab = 0 Then
            strText = "��ҩ��"
        End If
    End If
            
    For intCol = 0 To intCols
        If Trim(Bill.TextMatrix(0, intCol)) = strText Then
            GetDetailCol = intCol
            Exit Function
        End If
    Next
    GetDetailCol = -1
End Function

Private Sub SetFormat()
    Dim intFixedCol As Integer
    Dim strArr As Variant
    Dim strTemp As Variant
    Dim i As Integer
    
    '���ø��б�ؼ��ĸ�ʽ
    With Billδ��ҩ�嵥
        .rows = 2
        .Cols = ����_δ��ҩ�嵥.����
        
        .TextMatrix(0, ����_δ��ҩ�嵥.�����) = "��"
        .TextMatrix(0, ����_δ��ҩ�嵥.�����) = "��"
        .TextMatrix(0, ����_δ��ҩ�嵥.����) = "����"
        .TextMatrix(0, ����_δ��ҩ�嵥.����ҽ��) = "����ҽ��"
        .TextMatrix(0, ����_δ��ҩ�嵥.״̬) = "״̬"
        .TextMatrix(0, ����_δ��ҩ�嵥.����) = "����"
        .TextMatrix(0, ����_δ��ҩ�嵥.NO) = "NO"
        .TextMatrix(0, ����_δ��ҩ�嵥.����Ա) = "����Ա"
        .TextMatrix(0, ����_δ��ҩ�嵥.����) = "����"
        .TextMatrix(0, ����_δ��ҩ�嵥.����) = "����"
        .TextMatrix(0, ����_δ��ҩ�嵥.סԺ��) = "סԺ��"
        .TextMatrix(0, ����_δ��ҩ�嵥.ҩƷ����) = "ҩƷ����"
        .TextMatrix(0, ����_δ��ҩ�嵥.������) = "������"
        .TextMatrix(0, ����_δ��ҩ�嵥.Ӣ����) = "Ӣ����"
        .TextMatrix(0, ����_δ��ҩ�嵥.���) = "���"
        .TextMatrix(0, ����_δ��ҩ�嵥.����) = "����"
        .TextMatrix(0, ����_δ��ҩ�嵥.����) = "����"
        .TextMatrix(0, ����_δ��ҩ�嵥.��) = "��"
        .TextMatrix(0, ����_δ��ҩ�嵥.����) = "����"
        .TextMatrix(0, ����_δ��ҩ�嵥.����) = "����"
        .TextMatrix(0, ����_δ��ҩ�嵥.���) = "���"
        .TextMatrix(0, ����_δ��ҩ�嵥.����) = "����"
        .TextMatrix(0, ����_δ��ҩ�嵥.Ƶ��) = "Ƶ��"
        .TextMatrix(0, ����_δ��ҩ�嵥.�÷�) = "�÷�"
        .TextMatrix(0, ����_δ��ҩ�嵥.����ʱ��) = "����ʱ��"
        .TextMatrix(0, ����_δ��ҩ�嵥.˵��) = "˵��"
        .TextMatrix(0, ����_δ��ҩ�嵥.����) = "����"
        .TextMatrix(0, ����_δ��ҩ�嵥.ҽ��id) = "ҽ��id"
        .TextMatrix(0, ����_δ��ҩ�嵥.��ҩ��) = "��ҩ��"
        .TextMatrix(0, ����_δ��ҩ�嵥.�ⷿ��λ) = "�ⷿ��λ"
        .TextMatrix(0, ����_δ��ҩ�嵥.���ID) = ""
        .TextMatrix(0, ����_δ��ҩ�嵥.ҩƷID) = ""
        .TextMatrix(0, ����_δ��ҩ�嵥.������λ) = ""
        .TextMatrix(0, ����_δ��ҩ�嵥.��ҩ����) = "��ҩ����"
        .TextMatrix(0, ����_δ��ҩ�嵥.��ҩ����id) = ""
                
        .ColWidth(����_δ��ҩ�嵥.�����) = IIf(Not mblnStarPass, 0, 240)
        .ColWidth(����_δ��ҩ�嵥.�����) = 0
        .ColWidth(����_δ��ҩ�嵥.����) = 1000
        .ColWidth(����_δ��ҩ�嵥.����ҽ��) = 1100
        .ColWidth(����_δ��ҩ�嵥.״̬) = 700
        .ColWidth(����_δ��ҩ�嵥.����) = 900
        .ColWidth(����_δ��ҩ�嵥.NO) = 900
        .ColWidth(����_δ��ҩ�嵥.����Ա) = 800
        .ColWidth(����_δ��ҩ�嵥.����) = 600
        .ColWidth(����_δ��ҩ�嵥.����) = 700
        .ColWidth(����_δ��ҩ�嵥.סԺ��) = 1200
        .ColWidth(����_δ��ҩ�嵥.ҩƷ����) = 2000
        .ColWidth(����_δ��ҩ�嵥.������) = 2000
        .ColWidth(����_δ��ҩ�嵥.Ӣ����) = 2000
        .ColWidth(����_δ��ҩ�嵥.���) = 1500
        .ColWidth(����_δ��ҩ�嵥.����) = 1500
        .ColWidth(����_δ��ҩ�嵥.����) = 1500
        .ColWidth(����_δ��ҩ�嵥.��) = 300
        .ColWidth(����_δ��ҩ�嵥.����) = 1200
        .ColWidth(����_δ��ҩ�嵥.����) = 1200
        .ColWidth(����_δ��ҩ�嵥.����) = 1200
        .ColWidth(����_δ��ҩ�嵥.����) = 500
        .ColWidth(����_δ��ҩ�嵥.Ƶ��) = 500
        .ColWidth(����_δ��ҩ�嵥.�÷�) = 800
        .ColWidth(����_δ��ҩ�嵥.˵��) = 1200
        .ColWidth(����_δ��ҩ�嵥.����ʱ��) = 1800
        .ColWidth(����_δ��ҩ�嵥.����) = 0
        .ColWidth(����_δ��ҩ�嵥.ҽ��id) = 0
        .ColWidth(����_δ��ҩ�嵥.��ҩ��) = 1000
        .ColWidth(����_δ��ҩ�嵥.�ⷿ��λ) = 1200
        .ColWidth(����_δ��ҩ�嵥.���ID) = 0
        .ColWidth(����_δ��ҩ�嵥.ҩƷID) = 0
        .ColWidth(����_δ��ҩ�嵥.������λ) = 0
        .ColWidth(����_δ��ҩ�嵥.��ҩ����) = 1000
        .ColWidth(����_δ��ҩ�嵥.��ҩ����id) = 0
        
        For intFixedCol = 0 To .Cols - 1
            .ColAlignmentFixed(intFixedCol) = 4
        Next
        .ColAlignment(����_δ��ҩ�嵥.���) = 1
        .ColAlignment(����_δ��ҩ�嵥.����) = 1
        .ColAlignment(����_δ��ҩ�嵥.����ʱ��) = 1
        .ColAlignment(����_δ��ҩ�嵥.������) = 1
        .ColAlignment(����_δ��ҩ�嵥.Ӣ����) = 1
    End With
    
    With Bill���ܷ�ҩ
        .rows = 2
        .Cols = IIf(Lng������ʾ = 1, ����_���һ����嵥.����, ����_�����嵥.����)
        
        If Lng������ʾ = 0 Then
            .TextMatrix(0, ����_�����嵥.ҩƷ����) = "ҩƷ����"
            .TextMatrix(0, ����_�����嵥.���) = "���"
            .TextMatrix(0, ����_�����嵥.����) = "����"
            .TextMatrix(0, ����_�����嵥.����) = "����"
            .TextMatrix(0, ����_�����嵥.����) = "����"
            .TextMatrix(0, ����_�����嵥.��λ) = "��λ"
            .TextMatrix(0, ����_�����嵥.����) = "����"
            .TextMatrix(0, ����_�����嵥.���) = "���"
                        
            .ColWidth(����_�����嵥.ҩƷ����) = 2000
            .ColWidth(����_�����嵥.���) = 1500
            .ColWidth(����_�����嵥.����) = 1500
            .ColWidth(����_�����嵥.����) = 1200
            .ColWidth(����_�����嵥.����) = 1200
            .ColWidth(����_�����嵥.��λ) = 500
            .ColWidth(����_�����嵥.����) = 1200
            .ColWidth(����_�����嵥.���) = 1200
        Else
            .TextMatrix(0, ����_���һ����嵥.����) = "��������"
            .TextMatrix(0, ����_���һ����嵥.ҩƷ����) = "ҩƷ����"
            .TextMatrix(0, ����_���һ����嵥.���) = "���"
            .TextMatrix(0, ����_���һ����嵥.����) = "����"
            .TextMatrix(0, ����_���һ����嵥.����) = "����"
            .TextMatrix(0, ����_���һ����嵥.Ӧ������) = "Ӧ������"
            .TextMatrix(0, ����_���һ����嵥.��������) = "��������"
            .TextMatrix(0, ����_���һ����嵥.��������) = "��������"
            .TextMatrix(0, ����_���һ����嵥.ʵ������) = "ʵ������"
            .TextMatrix(0, ����_���һ����嵥.��λ) = "��λ"
            .TextMatrix(0, ����_���һ����嵥.����) = "����"
            .TextMatrix(0, ����_���һ����嵥.���) = "���"
            .TextMatrix(0, ����_���һ����嵥.����) = "����"
            .TextMatrix(0, ����_���һ����嵥.����ID) = "����ID"
            .TextMatrix(0, ����_���һ����嵥.ҩƷID) = "ҩƷID"
            .TextMatrix(0, ����_���һ����嵥.��ҩ����) = "��ҩ����"
            .TextMatrix(0, ����_���һ����嵥.��ҩ����id) = ""
            
            
            .ColWidth(����_���һ����嵥.����) = 1200
            .ColWidth(����_���һ����嵥.ҩƷ����) = 2000
            .ColWidth(����_���һ����嵥.���) = 1500
            .ColWidth(����_���һ����嵥.����) = 1500
            .ColWidth(����_���һ����嵥.����) = 1200
            .ColWidth(����_���һ����嵥.Ӧ������) = 1200
            .ColWidth(����_���һ����嵥.��������) = 1200
            .ColWidth(����_���һ����嵥.��������) = IIf(mbln���ܷ�ҩ = True, 1200, 0)
            .ColWidth(����_���һ����嵥.ʵ������) = 1200
            .ColWidth(����_���һ����嵥.��λ) = 500
            .ColWidth(����_���һ����嵥.����) = 1200
            .ColWidth(����_���һ����嵥.���) = 1200
            .ColWidth(����_���һ����嵥.����) = 0
            .ColWidth(����_���һ����嵥.����ID) = 0
            .ColWidth(����_���һ����嵥.ҩƷID) = 0
            .ColWidth(����_���һ����嵥.��ҩ����) = 1200
            .ColWidth(����_���һ����嵥.��ҩ����id) = 0
        End If
    
        For intFixedCol = 0 To .Cols - 1
            .ColAlignmentFixed(intFixedCol) = 4
        Next
        .ColAlignment(IIf(Lng������ʾ = 1, ����_���һ����嵥.���, ����_�����嵥.���)) = 1
        .ColAlignment(IIf(Lng������ʾ = 1, ����_���һ����嵥.����, ����_�����嵥.����)) = 1
    End With

    With Billȱҩ�嵥
        .rows = 2
        .Cols = 12
        
        .TextMatrix(0, 0) = "����"
        .TextMatrix(0, 1) = "NO"
        .TextMatrix(0, 2) = "����"
        .TextMatrix(0, 3) = "����"
        .TextMatrix(0, 4) = "����"
        .TextMatrix(0, 5) = "ҩƷ����"
        .TextMatrix(0, 6) = "���"
        .TextMatrix(0, 7) = "����"
        .TextMatrix(0, 8) = "����"
        .TextMatrix(0, 9) = "����"
        .TextMatrix(0, 10) = "����"
        .TextMatrix(0, 11) = "���"
        
        .ColWidth(0) = 1200
        .ColWidth(1) = 800
        .ColWidth(2) = 900
        .ColWidth(3) = 800
        .ColWidth(4) = 1000
        .ColWidth(5) = 2000
        .ColWidth(6) = 1500
        .ColWidth(7) = 1500
        .ColWidth(8) = 1000
        .ColWidth(9) = 1200
        .ColWidth(10) = 1200
        .ColWidth(11) = 1200
    
        For intFixedCol = 0 To .Cols - 1
            .ColAlignmentFixed(intFixedCol) = 4
        Next
        .ColAlignment(6) = 1
    End With

    With Bill�ܷ�ҩ�嵥
        .rows = 2
        .Cols = 13
        
        .TextMatrix(0, 0) = "����"
        .TextMatrix(0, 1) = "״̬"
        .TextMatrix(0, 2) = "NO"
        .TextMatrix(0, 3) = "����"
        .TextMatrix(0, 4) = "����"
        .TextMatrix(0, 5) = "����"
        .TextMatrix(0, 6) = "ҩƷ����"
        .TextMatrix(0, 7) = "���"
        .TextMatrix(0, 8) = "����"
        .TextMatrix(0, 9) = "����"
        .TextMatrix(0, 10) = "����"
        .TextMatrix(0, 11) = "����"
        .TextMatrix(0, 12) = "���"
        
        .ColWidth(0) = 1200
        .ColWidth(1) = 700
        .ColWidth(2) = 800
        .ColWidth(3) = 900
        .ColWidth(4) = 800
        .ColWidth(5) = 1000
        .ColWidth(6) = 2000
        .ColWidth(7) = 1500
        .ColWidth(8) = 1500
        .ColWidth(9) = 1500
        .ColWidth(10) = 1200
        .ColWidth(11) = 1200
        .ColWidth(12) = 1200
    
        For intFixedCol = 0 To .Cols - 1
            .ColAlignmentFixed(intFixedCol) = 4
        Next
        .ColAlignment(7) = 1
        .ColAlignment(8) = 1
    End With

    With Bill�ѷ�ҩ�嵥
        .rows = 2
        .Cols = ����_�ѷ�ҩ�嵥.����
        
        .TextMatrix(0, ����_�ѷ�ҩ�嵥.�����) = "��"
        .TextMatrix(0, ����_�ѷ�ҩ�嵥.�����) = "��"
        .TextMatrix(0, ����_�ѷ�ҩ�嵥.����) = "����"
        .TextMatrix(0, ����_�ѷ�ҩ�嵥.״̬) = "״̬"
        .TextMatrix(0, ����_�ѷ�ҩ�嵥.����) = "����"
        .TextMatrix(0, ����_�ѷ�ҩ�嵥.NO) = "NO"
        .TextMatrix(0, ����_�ѷ�ҩ�嵥.����) = "����"
        .TextMatrix(0, ����_�ѷ�ҩ�嵥.����) = "����"
        .TextMatrix(0, ����_�ѷ�ҩ�嵥.סԺ��) = "סԺ��"
        .TextMatrix(0, ����_�ѷ�ҩ�嵥.ҩƷ����) = "ҩƷ����"
        .TextMatrix(0, ����_�ѷ�ҩ�嵥.������) = "������"
        .TextMatrix(0, ����_�ѷ�ҩ�嵥.Ӣ����) = "Ӣ����"
        .TextMatrix(0, ����_�ѷ�ҩ�嵥.���) = "���"
        .TextMatrix(0, ����_�ѷ�ҩ�嵥.����) = "����"
        .TextMatrix(0, ����_�ѷ�ҩ�嵥.����) = "����"
        .TextMatrix(0, ����_�ѷ�ҩ�嵥.��) = "��"
        .TextMatrix(0, ����_�ѷ�ҩ�嵥.����) = "����"
        .TextMatrix(0, ����_�ѷ�ҩ�嵥.������) = "������"
        .TextMatrix(0, ����_�ѷ�ҩ�嵥.׼����) = "׼����"
        .TextMatrix(0, ����_�ѷ�ҩ�嵥.��ҩ��) = "��ҩ��"
        .TextMatrix(0, ����_�ѷ�ҩ�嵥.����) = "����"
        .TextMatrix(0, ����_�ѷ�ҩ�嵥.���) = "���"
        .TextMatrix(0, ����_�ѷ�ҩ�嵥.����) = "����"
        .TextMatrix(0, ����_�ѷ�ҩ�嵥.Ƶ��) = "Ƶ��"
        .TextMatrix(0, ����_�ѷ�ҩ�嵥.�÷�) = "�÷�"
        .TextMatrix(0, ����_�ѷ�ҩ�嵥.����Ա) = "����Ա"
        .TextMatrix(0, ����_�ѷ�ҩ�嵥.��ҩʱ��) = "��ҩʱ��"
        .TextMatrix(0, ����_�ѷ�ҩ�嵥.����) = "����"
        .TextMatrix(0, ����_�ѷ�ҩ�嵥.ҽ��id) = "ҽ��id"
        .TextMatrix(0, ����_�ѷ�ҩ�嵥.��ҩ��) = "��/��ҩ��"
        .TextMatrix(0, ����_�ѷ�ҩ�嵥.�ⷿ��λ) = "�ⷿ��λ"
        .TextMatrix(0, ����_�ѷ�ҩ�嵥.���ID) = ""
        .TextMatrix(0, ����_�ѷ�ҩ�嵥.ҩƷID) = ""
        .TextMatrix(0, ����_�ѷ�ҩ�嵥.������λ) = ""
                
        .ColWidth(����_�ѷ�ҩ�嵥.�����) = IIf(Not mblnStarPass, 0, 240)
        .ColWidth(����_�ѷ�ҩ�嵥.�����) = 0
        .ColWidth(����_�ѷ�ҩ�嵥.����) = 1200
        .ColWidth(����_�ѷ�ҩ�嵥.״̬) = 700
        .ColWidth(����_�ѷ�ҩ�嵥.����) = 900
        .ColWidth(����_�ѷ�ҩ�嵥.NO) = 900
        .ColWidth(����_�ѷ�ҩ�嵥.����) = 600
        .ColWidth(����_�ѷ�ҩ�嵥.����) = 700
        .ColWidth(����_�ѷ�ҩ�嵥.סԺ��) = 1200
        .ColWidth(����_�ѷ�ҩ�嵥.ҩƷ����) = 2000
        .ColWidth(����_�ѷ�ҩ�嵥.������) = 2000
        .ColWidth(����_�ѷ�ҩ�嵥.Ӣ����) = 2000
        .ColWidth(����_�ѷ�ҩ�嵥.���) = 1500
        .ColWidth(����_�ѷ�ҩ�嵥.����) = 1500
        .ColWidth(����_�ѷ�ҩ�嵥.����) = 1500
        .ColWidth(����_�ѷ�ҩ�嵥.��) = 300
        .ColWidth(����_�ѷ�ҩ�嵥.����) = 1000
        .ColWidth(����_�ѷ�ҩ�嵥.������) = 1000
        .ColWidth(����_�ѷ�ҩ�嵥.׼����) = 1000
        .ColWidth(����_�ѷ�ҩ�嵥.��ҩ��) = 1000
        .ColWidth(����_�ѷ�ҩ�嵥.����) = 1000
        .ColWidth(����_�ѷ�ҩ�嵥.���) = 1000
        .ColWidth(����_�ѷ�ҩ�嵥.����) = 500
        .ColWidth(����_�ѷ�ҩ�嵥.Ƶ��) = 500
        .ColWidth(����_�ѷ�ҩ�嵥.�÷�) = 800
        .ColWidth(����_�ѷ�ҩ�嵥.����Ա) = 800
        .ColWidth(����_�ѷ�ҩ�嵥.��ҩʱ��) = 1500
        .ColWidth(����_�ѷ�ҩ�嵥.����) = 0
        .ColWidth(����_�ѷ�ҩ�嵥.ҽ��id) = 0
        .ColWidth(����_�ѷ�ҩ�嵥.��ҩ��) = 1000
        .ColWidth(����_�ѷ�ҩ�嵥.�ⷿ��λ) = 1200
        .ColWidth(����_�ѷ�ҩ�嵥.���ID) = 0
        .ColWidth(����_�ѷ�ҩ�嵥.ҩƷID) = 0
        .ColWidth(����_�ѷ�ҩ�嵥.������λ) = 0
                
        For intFixedCol = 0 To .Cols - 1
            .ColAlignmentFixed(intFixedCol) = 4
        Next
        .ColAlignment(����_�ѷ�ҩ�嵥.���) = 1
        .ColAlignment(����_�ѷ�ҩ�嵥.����) = 1
        .ColAlignment(����_�ѷ�ҩ�嵥.�÷�) = 1
        .ColAlignment(����_�ѷ�ҩ�嵥.������) = 7
        .ColAlignment(����_�ѷ�ҩ�嵥.׼����) = 7
        .ColAlignment(����_�ѷ�ҩ�嵥.��ҩ��) = 7
        .ColAlignment(����_�ѷ�ҩ�嵥.��ҩʱ��) = 1
        .ColAlignment(����_�ѷ�ҩ�嵥.������) = 1
        .ColAlignment(����_�ѷ�ҩ�嵥.Ӣ����) = 1
    End With
  
    '��ʼ�����б�
    strTemp = Split(mconstRequest, "|")
    With Bill��ҩ����
        .Redraw = False
        .rows = 2
        .Cols = �����б�.����
        .SelectionMode = flexSelectionByRow
        For i = 0 To .Cols - 1
            strArr = Split(strTemp(i), ",")
            
            If strArr(0) = "Ч��" Then
                .TextMatrix(0, i) = IIf(gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1, "��Ч����", "ʧЧ��")
            Else
                .TextMatrix(0, i) = strArr(0)
            End If
            
            .ColAlignment(i) = strArr(1)
            .ColWidth(i) = strArr(2)
            
            .FixedAlignment(i) = 4
        Next
        .Redraw = True
    End With
  
    '�����ʼ���п�
    Call SaveColDefaultWidth(Billδ��ҩ�嵥)
    Call SaveColDefaultWidth(Bill�ѷ�ҩ�嵥)
  
    Call RestoreFlexState(Bill���ܷ�ҩ, "���ܷ�ҩ" & Lng������ʾ)
    Call RestoreFlexState(Billȱҩ�嵥, "ȱҩ�嵥")
    Call RestoreFlexState(Bill�ܷ�ҩ�嵥, "�ܷ�ҩ�嵥")
    Call RestoreFlexState(Billδ��ҩ�嵥, "δ��ҩ�嵥")
    Call RestoreFlexState(Bill�ѷ�ҩ�嵥, "�ѷ�ҩ�嵥")
    '�ָ����Ի����ú��м���ʼ�ղ�������
    If Billδ��ҩ�嵥.ColWidth(����_δ��ҩ�嵥.״̬) < 200 Then Billδ��ҩ�嵥.ColWidth(����_δ��ҩ�嵥.״̬) = 700
    If Billδ��ҩ�嵥.ColWidth(����_δ��ҩ�嵥.����) < 200 Then Billδ��ҩ�嵥.ColWidth(����_δ��ҩ�嵥.����) = 1500
    If Bill�ѷ�ҩ�嵥.ColWidth(����_�ѷ�ҩ�嵥.״̬) < 200 Then Bill�ѷ�ҩ�嵥.ColWidth(����_�ѷ�ҩ�嵥.״̬) = 700
    If Bill�ѷ�ҩ�嵥.ColWidth(����_�ѷ�ҩ�嵥.��ҩ��) < 200 Then Bill�ѷ�ҩ�嵥.ColWidth(����_�ѷ�ҩ�嵥.��ҩ��) = 1000
    '��ʾ�и��ݲ����������Ƿ���ʾ
    Billδ��ҩ�嵥.ColWidth(����_δ��ҩ�嵥.�����) = IIf(Not mblnStarPass, 0, 240)
    Bill�ѷ�ҩ�嵥.ColWidth(����_�ѷ�ҩ�嵥.�����) = IIf(Not mblnStarPass, 0, 240)
    
    Bill���ܷ�ҩ.ColWidth(����_���һ����嵥.��������) = IIf(mbln���ܷ�ҩ = True, 1200, 0)
End Sub

Private Function LoadInIcon() As Boolean
    '--Ϊ���ؼ�װ��ͼ��--
    On Error Resume Next
    err = 0
    LoadInIcon = False
    
    '������
    With ImgTbarBlack
        .ImageHeight = 24
        .ImageWidth = 24
        .ListImages.Add , , LoadResPicture("BPREVIEW", vbResIcon)
        .ListImages.Add , , LoadResPicture("BPRINT", vbResIcon)
        .ListImages.Add , , LoadResPicture("BDOSAGE", vbResIcon)
        .ListImages.Add , , LoadResPicture("BDOSAGE", vbResIcon)
        .ListImages.Add , , LoadResPicture("BSEND", vbResIcon)
        .ListImages.Add , , LoadResPicture("BSEND", vbResIcon)
        .ListImages.Add , , LoadResPicture("BHELP", vbResIcon)
        .ListImages.Add , , LoadResPicture("BEXIT", vbResIcon)
        .ListImages.Add , , LoadResPicture("BBACKSTRICK", vbResIcon)
    End With
    With ImgTbarColor
        .ImageHeight = 24
        .ImageWidth = 24
        .ListImages.Add , , LoadResPicture("CPREVIEW", vbResIcon)
        .ListImages.Add , , LoadResPicture("CPRINT", vbResIcon)
        .ListImages.Add , , LoadResPicture("CDOSAGE", vbResIcon)
        .ListImages.Add , , LoadResPicture("CDOSAGE", vbResIcon)
        .ListImages.Add , , LoadResPicture("CSEND", vbResIcon)
        .ListImages.Add , , LoadResPicture("CSEND", vbResIcon)
        .ListImages.Add , , LoadResPicture("CHELP", vbResIcon)
        .ListImages.Add , , LoadResPicture("CEXIT", vbResIcon)
        .ListImages.Add , , LoadResPicture("CBACKSTRICK", vbResIcon)
    End With
    With Tbar
        Set .ImageList = ImgTbarBlack
        Set .HotImageList = ImgTbarColor

        .Buttons("Preview").Image = 1
        .Buttons("Print").Image = 2
        .Buttons("Consignment").Image = 3
        .Buttons("Desire").Image = 4
        .Buttons("Handback").Image = 5
        .Buttons("Restore").Image = 6
        .Buttons("Help").Image = 7
        .Buttons("Exit").Image = 8
        .Buttons("ReVerify").Image = 9
    End With
    Cbar.Bands(1).MinHeight = Tbar.Height
    
    If err <> 0 Then
        MsgBox "�����Դ�ļ���ʧ�����������������ϵ��", vbInformation, gstrSysName
        Exit Function
    End If
    LoadInIcon = True
End Function

Private Function InitRefreshRec()
    
    '����ִ�й��ܣ���ҩ���ܷ����󣬽��ϴ��趨�ķǷ�ҩ��ȱҩ�ļ�¼��ִ��״̬�ָ�
    Set RecRefreshCompare = New adodb.Recordset
    With RecRefreshCompare
        If .State = 1 Then .Close
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "״̬", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "NO", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "Ʒ��", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "���", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "��", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����", adDouble, 18, adFldIsNullable
        .Fields.Append "���", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "Ƶ��", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "�÷�", adLongVarChar, 30, adFldIsNullable
        .Fields.Append "ID", adDouble, 18, adFldIsNullable
        .Fields.Append "ҩƷID", adDouble, 18, adFldIsNullable
        .Fields.Append "���շ�", adDouble, 2, adFldIsNullable
        .Fields.Append "λ��", adDouble, 18, adFldIsNullable
        .Fields.Append "ִ��״̬", adDouble, 1, adFldIsNullable
        .Fields.Append "ʵ������", adDouble, 18, adFldIsNullable            '�жϿ����
        .Fields.Append "˵��", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "����ʱ��", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "��ҩ��", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "�����", adLongVarChar, 20, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Function

Private Function InitRec()
    '������:����
    '��������:2000-11-02
    
    'δ��������¼��
    If Blnˢ��δ��ҩ�嵥 = True Then
        Set RecChangeData = New adodb.Recordset
        With RecChangeData
            If .State = 1 Then .Close
            .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "����ҽ��", adLongVarChar, 20, adFldIsNullable
            .Fields.Append "״̬", adLongVarChar, 10, adFldIsNullable
            .Fields.Append "����", adLongVarChar, 20, adFldIsNullable
            .Fields.Append "NO", adLongVarChar, 8, adFldIsNullable
            .Fields.Append "����", adDouble, 18, adFldIsNullable
            .Fields.Append "����ID", adDouble, 18, adFldIsNullable
            .Fields.Append "���", adDouble, 18, adFldIsNullable
            .Fields.Append "����", adLongVarChar, 10, adFldIsNullable
            .Fields.Append "����", adLongVarChar, 20, adFldIsNullable
            .Fields.Append "סԺ��", adLongVarChar, 20, adFldIsNullable
            .Fields.Append "Ʒ��", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "������", adLongVarChar, 80, adFldIsNullable
            .Fields.Append "Ӣ����", adLongVarChar, 80, adFldIsNullable
            .Fields.Append "���", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "�������", adLongVarChar, 10, adFldIsNullable
            .Fields.Append "��ֵ����", adLongVarChar, 10, adFldIsNullable
            .Fields.Append "����", adDouble, 18, adFldIsNullable
            .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "����", adDouble, 2, adFldIsNullable
            .Fields.Append "��", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "���", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "����Ա", adLongVarChar, 20, adFldIsNullable
            .Fields.Append "����", adLongVarChar, 20, adFldIsNullable
            .Fields.Append "Ƶ��", adLongVarChar, 20, adFldIsNullable
            .Fields.Append "�÷�", adLongVarChar, 30, adFldIsNullable
            .Fields.Append "ID", adDouble, 18, adFldIsNullable
            .Fields.Append "ҩƷID", adDouble, 18, adFldIsNullable
            .Fields.Append "���շ�", adDouble, 2, adFldIsNullable
            .Fields.Append "λ��", adDouble, 18, adFldIsNullable
            .Fields.Append "ִ��״̬", adDouble, 1, adFldIsNullable
            .Fields.Append "ʵ������", adDouble, 18, adFldIsNullable            '�жϿ����
            .Fields.Append "��������", adDouble, 18, adFldIsNullable
            .Fields.Append "˵��", adLongVarChar, 40, adFldIsNullable
            .Fields.Append "����ʱ��", adLongVarChar, 20, adFldIsNullable
            .Fields.Append "��ҩ��", adLongVarChar, 20, adFldIsNullable
            .Fields.Append "�����", adLongVarChar, 20, adFldIsNullable
            .Fields.Append "�����", adDouble, 18, adFldIsNullable
            .Fields.Append "ҽ��id", adDouble, 18, adFldIsNullable
            .Fields.Append "��ҩ��", adLongVarChar, 20, adFldIsNullable
            .Fields.Append "�ⷿ��λ", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "���id", adDouble, 18, adFldIsNullable
            .Fields.Append "����ID", adDouble, 18, adFldIsNullable
            .Fields.Append "������λ", adLongVarChar, 20, adFldIsNullable
            .Fields.Append "�������", adDouble, 18, adFldIsNullable
            .Fields.Append "��ҩ����", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "��ҩ����ID", adDouble, 18, adFldIsNullable
                        
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .Open
        End With
        
        Set mrsRequest = New adodb.Recordset
        With mrsRequest
            If .State = 1 Then .Close
            .Fields.Append "��ҩ����", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "��ҩ����ID", adDouble, 18, adFldIsNullable
            .Fields.Append "����", adDouble, 18, adFldIsNullable
            .Fields.Append "NO", adLongVarChar, 20, adFldIsNullable
            .Fields.Append "ҩƷID", adDouble, 18, adFldIsNullable
            .Fields.Append "����ʱ��", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "�շ����", adDouble, 18, adFldIsNullable
            .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "Ч��", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "׼������", adDouble, 18, adFldIsNullable
            .Fields.Append "��������", adDouble, 18, adFldIsNullable
            .Fields.Append "��װ", adDouble, 18, adFldIsNullable
            .Fields.Append "��λ", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "�շ�ID", adDouble, 18, adFldIsNullable
            .Fields.Append "��ҳID", adDouble, 18, adFldIsNullable
            .Fields.Append "�������", adDouble, 18, adFldIsNullable
            .Fields.Append "����", adDouble, 18, adFldIsNullable
            .Fields.Append "����ID", adDouble, 18, adFldIsNullable
            .Fields.Append "��¼����", adDouble, 18, adFldIsNullable
            .Fields.Append "��˱�־", adDouble, 18, adFldIsNullable
            .Fields.Append "ҩƷ����", adLongVarChar, 100, adFldIsNullable
            
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .Open
        End With
        
        Set mrsRequestMain = New adodb.Recordset
        With mrsRequestMain
            If .State = 1 Then .Close
            .Fields.Append "��ҩ����ID", adDouble, 18, adFldIsNullable
            .Fields.Append "ҩƷID", adDouble, 18, adFldIsNullable
            .Fields.Append "����ʱ��", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "׼������", adDouble, 18, adFldIsNullable
            .Fields.Append "��������", adDouble, 18, adFldIsNullable
            .Fields.Append "����ID", adDouble, 18, adFldIsNullable
            
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .Open
        End With
    End If
    
    '�ѷ�������¼��
    Set RecChangeSendedData = New adodb.Recordset
    With RecChangeSendedData
        If .State = 1 Then .Close
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "ID", adDouble, 18, adFldIsNullable
        .Fields.Append "ҩƷID", adDouble, 18, adFldIsNullable
        .Fields.Append "ִ��״̬", adDouble, 1, adFldIsNullable
        .Fields.Append "NO", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "����", adDouble, 18, adFldIsNullable
        .Fields.Append "����ID", adDouble, 18, adFldIsNullable
        .Fields.Append "���", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "סԺ��", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "Ʒ��", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "������", adLongVarChar, 80, adFldIsNullable
        .Fields.Append "Ӣ����", adLongVarChar, 80, adFldIsNullable
        .Fields.Append "���", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "�������", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "����", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "Ч��", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����", adDouble, 2, adFldIsNullable
        .Fields.Append "��", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "������", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "׼����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "��ҩ��", adDouble, 18, adFldIsNullable
        .Fields.Append "�ɲ���", adDouble, 2, adFldIsNullable
        .Fields.Append "��λ", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "���", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "Ƶ��", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "�÷�", adLongVarChar, 30, adFldIsNullable
        .Fields.Append "˵��", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "����Ա", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "��ҩʱ��", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "λ��", adDouble, 18, adFldIsNullable
        .Fields.Append "�����", adDouble, 18, adFldIsNullable
        .Fields.Append "ҽ��id", adDouble, 18, adFldIsNullable
        .Fields.Append "��ҩ��", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "ʵ������", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "�ⷿ��λ", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "���id", adDouble, 18, adFldIsNullable
        .Fields.Append "������λ", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "ת��", adDouble, 1, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Function
Private Function DependOnCheck() As Boolean
    Dim strSQL As String
    Dim BlnInҩ�� As Boolean
    Dim n As Integer
    
    On Error GoTo errHandle
    '�������ݼ��
    DependOnCheck = False
    
   '���ҩ�����÷�(��ҩ������ҩ������ҩ��)
    If IsHavePrivs(mstrPrivs, "����ҩ��") Then
        strSQL = "(Select Distinct ����ID From ��������˵�� Where �������� Like '%ҩ��' And ������� IN (2,3))"
    Else
        strSQL = "(Select distinct A.����ID From ������Ա A,��������˵�� B " & _
                 " Where A.��ԱID=[1] And A.����ID=B.����ID And B.�������� Like '%ҩ��' And B.������� IN (2,3))"
    End If
    gstrSQL = " Select Distinct P.ID,P.���� From ���ű� P " & _
             " Where (P.վ�� = '" & gstrNodeNo & "' Or P.վ�� is Null) And P.ID In " & strSQL & _
             " And (P.����ʱ�� Is Null Or P.����ʱ��=To_Date('3000-01-01','yyyy-MM-dd'))"
    Set RecBillData = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, glngUserId)
    
    With RecBillData
        If .EOF Then
           If IsHavePrivs(mstrPrivs, "����ҩ��") Then
               strSQL = "���ʼ��ҩ�������Ź���"
           Else
               strSQL = "�㲻��ҩ��������Ա�����ܲ�����ģ�飡"
           End If
           MsgBox strSQL, vbInformation, gstrSysName
           Exit Function
        Else
            Cbo��ҩҩ��.Clear
            Do While Not .EOF
                Cbo��ҩҩ��.AddItem !����
                Cbo��ҩҩ��.ItemData(Cbo��ҩҩ��.NewIndex) = !Id
                .MoveNext
            Loop
            Cbo��ҩҩ��.ListIndex = 0
       End If
       
       Call ReadFromReg
       Call mnuViewFontSet_Click(intFont)
       
       If lngҩ��ID <> 0 Then
           .MoveFirst
           .Find "ID=" & lngҩ��ID
           BlnInҩ�� = (.EOF <> True)
       End If
       
       '���ö�Ӧ��ҩ��
       If lngҩ��ID = 0 Or BlnInҩ�� = False Then
           '�����ô���
            With Frm���ŷ�ҩ��������
                .strPrivs = mstrPrivs
                .Show 1, Me
            End With
            Call ReadFromReg
            
            If lngҩ��ID = 0 Then
                MsgBox "����������ҩ��������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
                Exit Function
            End If
           
           '��δ����ҩ�����˳�
           If lngҩ��ID = 0 Then Exit Function
           .MoveFirst
           .Find "ID=" & lngҩ��ID
           BlnInҩ�� = (.EOF <> True)
           If Not BlnInҩ�� Then Exit Function
       End If
    End With
    
    DependOnCheck = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ReadFromReg() As Boolean
    Dim strTemp As String
    Dim RecRead As New adodb.Recordset
    Dim dateCurDate As Date
    Dim strArr
    Dim n As Integer
    
    On Error GoTo errHandle
    'ȡ������˽�в���
    '˽��ģ��
    intFont = Val(zldatabase.GetPara("����", glngSys, 1342))
    StrFindStyle = IIf(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0) = "0", "%", "")
    
    '����ģ��
    intDays = Val(zldatabase.GetPara("��ѯ����", glngSys, 1342)) - 1
    int��ҩ���� = Val(zldatabase.GetPara("��ҩ����", glngSys, 1342))
    Lng��ҩ��ǩ�� = Val(zldatabase.GetPara("��ҩ��ǩ��", glngSys, 1342))
    Lngȱҩ��� = Val(zldatabase.GetPara("ȱҩ���", glngSys, 1342))
    Lng��ҩ��ǩ�� = Val(zldatabase.GetPara("��ҩ��ǩ��", glngSys, 1342))
    mint�Զ�ˢ��δ��ҩ�嵥 = Val(zldatabase.GetPara("�Զ�ˢ��δ��ҩ�嵥", glngSys, 1342))
    mstr������ҩ��ʽ = zldatabase.GetPara("������ҩ��ʽ", glngSys, 1342, "�ٴ�,����,���,����,����,����,Ӫ��")
    mblnҩƷ���� = (Val(zldatabase.GetPara("�ⷿ��λ�����������ʾ", glngSys, 1342, 0)) = 1)
    mbln���ܷ�ҩ = (Val(zldatabase.GetPara("��ҩʱ������ҩ���ʼ�¼", glngSys, 1342, 0)) = 1)

    Lng����ģʽ = Val(zldatabase.GetPara("����ģʽ", glngSys, 1342))
    Lngҽ������ = Val(zldatabase.GetPara("ҽ������", glngSys, 1342))
    int��Ժ��ҩ = Val(zldatabase.GetPara("��Ժ��ҩ", glngSys, 1342))
    Lng������ʾ = Val(zldatabase.GetPara("�����һ�����ʾ�����嵥", glngSys, 1342))
    str������ = zldatabase.GetPara("������", glngSys, 1342, "���м�����")
    mstr������� = zldatabase.GetPara("�������", glngSys, 1342)
    mstr��ֵ���� = zldatabase.GetPara("��ֵ����", glngSys, 1342)
    
    lngҩ��ID = Val(zldatabase.GetPara("��ҩҩ��", glngSys, 1342))
    Lng�Զ���ӡ = Val(zldatabase.GetPara("�Զ���ӡ", glngSys, 1342))
    
    mlng�������� = GetSetting("ZLSOFT", "����ģ��\����\" & App.ProductName & "\Frm���ŷ�ҩ����", "��ʾ��ҩ��������", 1)
    

    '���ݲ�������
    Chk��ʾ��ҩ��������.Value = mlng��������
    
    '[ҩƷ��������]ϵͳ����
    gstrSQL = " Select Nvl(��鷽ʽ,0) ����� From ҩƷ������ Where �ⷿID=[1]"
    Set RecRead = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngҩ��ID)
    
    If Not RecRead.EOF Then
        IntCheckStock = RecRead!�����
    End If
    
    '����ϵͳ�����趨�ĵ�λ��ʾ����
    strUnit = GetSpecUnit(lngҩ��ID, gintסԺҩ��)
    
    '���õ�ǰҩ��
    If lngҩ��ID > 0 And Cbo��ҩҩ��.ListCount > 0 Then
        Cbo��ҩҩ��.Tag = lngҩ��ID
        For n = 0 To Cbo��ҩҩ��.ListCount - 1
            If lngҩ��ID = Cbo��ҩҩ��.ItemData(n) Then
                Cbo��ҩҩ��.ListIndex = n
                Exit For
            End If
        Next
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub GetSysParms()
    Int����δ��˴�����ҩ = gtype_UserSysParms.P6_δ��˼��ʴ�����ҩ
    
    blnҽ������ = (gtype_UserSysParms.P68_����ҩ�������Ϻ���ҩ = 0)          'Ϊ���ʾ������ҩ
    
    '��ȡ���С��λ��
    int����λ�� = gtype_UserSysParms.P9_���ý���λ��
    
    '�жϻ��۵���ҩ���Ƿ��Զ����Ϊ���ʵ�
    int��˻��۵� = gtype_UserSysParms.P81_ִ�к��Զ���˻��۵�
End Sub
Private Function RefreshData() As Boolean
    Dim strCond As String, strSubSql As String
    Dim strName As String
    Dim str���˼����� As String
    Dim strSql��ҩ�� As String
    Dim strSql�������� As String
    
    RefreshData = False
    On Error GoTo errHandle
    '��Ҫ�������
    If mstrSerchNO = "" And mstr���� = "" Then
'        MsgBox "��ѡ����ҩ���ţ�", vbInformation, gstrSysName
        Call ClearCons
        Exit Function
    End If
    
    str���˼����� = IIf(str������ <> "���м�����", " AND S.������=[1] ", "")
    
    '����:bit1=0-����,1-������bit2:3-��Ժ��ҩ
    '����ģʽ:0-����,1-���ʵ�,2-���ʱ�
    If Lng����ģʽ = 0 Then
        strCond = " And S.���� IN(9,10)"
    ElseIf Lng����ģʽ = 1 Then
        strCond = " And S.����=9"
    ElseIf Lng����ģʽ = 2 Then
        strCond = " And S.����=10"
    End If
    
    'ҽ������:0-����,1-����,2-����,3-��ͨ
    '�õ����Ƿ���д�����Ƿ�ҽ��������ҩƷ����
    If Lngҽ������ = 0 Then
    ElseIf Lngҽ������ = 1 Then
        strCond = strCond & " And S.���� Is Not Null And Ltrim(To_Char(Nvl(S.����,0),'00')) Like '0_' And Nvl(C.ҽ�����,0) + 0 >0 "
    ElseIf Lngҽ������ = 2 Then
        strCond = strCond & " And S.���� Is Not Null And Ltrim(To_Char(Nvl(S.����,0),'00')) Like '1_' And Nvl(C.ҽ�����,0) + 0 >0 "
    ElseIf Lngҽ������ = 3 Then
        strCond = strCond & " And (Nvl(C.ҽ�����,0) + 0 =0 Or S.���� Is Null) "
    ElseIf Lngҽ������ = 4 Then
        strCond = strCond & " And S.���� Is Not Null And (Ltrim(To_Char(Nvl(S.����,0),'00')) Like '0_' Or Ltrim(To_Char(Nvl(S.����,0),'00')) Like '1_') And Nvl(C.ҽ�����,0) + 0 > 0 "
    End If
    
    '��Ժ��ҩ:'0-����,1-������Ժ��ҩ,2-������Ժ��ҩ,3-������ȡҩ,4-������ȡҩ,5-Ժ����ҩ(��������Ժ��ҩ����ȡҩ),6-��Ժ��ҩ����ȡҩ
    If int��Ժ��ҩ = 0 Then
    ElseIf int��Ժ��ҩ = 1 Then
        strCond = strCond & " And Not Ltrim(To_Char(Nvl(S.����,0),'00')) Like '_3'"
    ElseIf int��Ժ��ҩ = 2 Then
        strCond = strCond & " And Ltrim(To_Char(Nvl(S.����,0),'00')) Like '_3'"
    ElseIf int��Ժ��ҩ = 3 Then
        strCond = strCond & " And Not Ltrim(To_Char(Nvl(S.����,0),'00')) Like '_4'"
    ElseIf int��Ժ��ҩ = 4 Then
        strCond = strCond & " And Ltrim(To_Char(Nvl(S.����,0),'00')) Like '_4'"
    ElseIf int��Ժ��ҩ = 5 Then
        strCond = strCond & " And Not Ltrim(To_Char(Nvl(S.����,0),'00')) Like '_3' And Not Ltrim(To_Char(Nvl(S.����,0),'00')) Like '_4'"
    ElseIf int��Ժ��ҩ = 6 Then
        strCond = strCond & " And (Ltrim(To_Char(Nvl(S.����,0),'00')) Like '_3' Or Ltrim(To_Char(Nvl(S.����,0),'00')) Like '_4')"
    End If
    
    If mint���� = 0 Then
        strCond = strCond & " And H.Id = C.���˿���id "
    ElseIf mint���� = 1 Then
        strCond = strCond & " And H.Id = C.��������id "
    Else
        strCond = strCond & " And H.Id = C.���˲���ID "
    End If
    
    '��λ����
    Select Case strUnit
    Case "�ۼ۵�λ"
        strSubSql = "X.���㵥λ ��λ,1 ��װ,"
    Case "���ﵥλ"
        strSubSql = "D.���ﵥλ ��λ,D.�����װ ��װ,"
    Case "סԺ��λ"
        strSubSql = "D.סԺ��λ ��λ,D.סԺ��װ ��װ,"
    Case "ҩ�ⵥλ"
        strSubSql = "D.ҩ�ⵥλ ��λ,D.ҩ���װ ��װ,"
    End Select
    
    '�õ�ҩƷ���ƴ�
    Call GetDrugFormat
    Select Case intҩƷ����
    Case 0  'ҩƷ����������
        strName = "'['||X.����||']'||" & IIf(mblnTradeName, "NVL(E.����,X.����)", "X.����") & " As Ʒ��,"
    Case 1  'ҩƷ����
        strName = "X.����" & " As Ʒ��,"
    Case 2  'ҩƷ����
        strName = IIf(mblnTradeName, "NVL(E.����,X.����)", "X.����") & " As Ʒ��,"
    End Select
    
    strName = strName & IIf(Not mblnTradeName, "NVL(E.����,'')", "Decode(E.����,Null,'',X.����)") & " As ������, "
    
    '�������ͣ����˻�Ӥ��
    If mint�������� = 0 Then
        strSql�������� = " And Nvl(C.Ӥ����,0)=0 "
    ElseIf mint�������� = 1 Then
        strSql�������� = " And Nvl(C.Ӥ����,0)>0 "
    End If
    
    gstrSQL = "SELECT A.*, Nvl(C.��������,0) As �������� " & IIf(mbln��ʾ����ҩ�� = True, ", B.��ҩ��", "") & " FROM " & _
             " (SELECT DISTINCT S.ID,S.ҩƷID,NVL(N.���շ�,0) ���շ�,P.���� ����,S.��ҩ��,C.������ ����ҽ��,C.����Ա���� �����,S.����,S.����," & _
             " S.NO,S.���,C.����ID,C.����,C.����,C.�����־,C.��ʶ��,C.����Ա����," & strName & " S.���� ��,S.ʵ������ ����," & _
             " NVL(D.ҩ������,0) ����,X.���,T.�������,T.��ֵ����,C.�Ǽ�ʱ��,H.���� As ��ҩ����,H.Id As ��ҩ����Id," & _
             strSubSql & _
             " S.���ۼ� ����,S.���۽�� ���,S.����,S.Ƶ��,S.�÷�,S.ժҪ ˵��,DECODE(S.����,NULL,'',S.����)||DECODE(S.����,NULL,'',0,'','('||S.����||')') ����,NVL(S.����,0) ����,C.ҽ�����,I.���㵥λ,NVL(S.����,NVL(X.����,'')) ����,nvl(M.�����,-1) �����,nvl(C.ҽ�����,-1) ҽ��id," & IIf(mblnҩƷ���� = True, "L.", "'' ") & "�ⷿ��λ,M.���ID,S.�Է�����id As ����ID,C.��� �������," & IIf(mblnҩƷ���� = True, "Decode(Sign(Nvl(K.�������, 0) - Nvl(L.����, 0)), -1, 0, 1) ", "0 ") & " �������, Z.���� As Ӣ���� " & _
             " FROM ҩƷ�շ���¼ S,���˷��ü�¼ C,δ��ҩƷ��¼ N,���ű� P,���ű� H,�շ���Ŀ���� E,�շ���ĿĿ¼ X,ҩƷ��� D,ҩƷ���� T,������ĿĿ¼ I,����ҽ����¼ M," & IIf(mblnҩƷ���� = True, "ҩƷ�����޶� L,", "") & "������Ŀ���� Z "
             
    If mblnҩƷ���� = True Then
        gstrSQL = gstrSQL & ",(Select �ⷿid, ҩƷid, Nvl(Sum(ʵ������), 0) ������� From ҩƷ��� Where ���� = 1 And �ⷿid = [2] Group By �ⷿid, ҩƷid) K "
    End If
             
    gstrSQL = gstrSQL & " WHERE S.ҩƷID=D.ҩƷID AND d.ҩƷID=X.ID and D.ҩ��ID=T.ҩ��ID AND D.ҩ��ID=I.ID and C.ҽ�����=M.ID(+) " & _
             " And D.ҩ��id = Z.������Ŀid(+) And Z.����(+) = 2 " & IIf(mblnҩƷ���� = True, " And S.ҩƷID=L.ҩƷID(+) And Nvl(S.�ⷿID,[2])=L.�ⷿID(+) ", "") & _
             " AND D.ҩƷID=E.�շ�ϸĿID(+) AND E.����(+)=3 " & IIf(mstr���� = "", "", " And C.����=[11] ") & _
             " AND S.NO=N.NO AND S.����=N.���� AND NVL(S.�ⷿID,[2])+0=NVL(N.�ⷿID,[2]) AND S.����ID=C.ID " & _
             IIf(Val(mlng����ID) = 0, "", " AND C.����ID=[3]") & IIf(Trim(mstrסԺ��) = "", "", " AND C.��ʶ��=[4]") & IIf(mstr�������� = "", "", " AND C.���� LIKE [5] ") & _
             " AND S.�Է�����ID+0=P.ID AND S.����� IS NULL " & IIf(mstr��ʼNO = "", "", " AND S.NO>=[6] ") & IIf(mstr����NO = "", "", " AND S.NO<=[7] ") & _
             " AND NVL(S.�ⷿID,[2])+0=[2] AND N.�������� BETWEEN [8] AND [9] " & IIf(mstrDrug = "", "", " And Instr([14],',' || T.ҩƷ���� || ',') > 0") & IIf(mstr��ҩ���� = "", "", " And Instr([15],',' || D.��ҩ���� || ',') > 0") & _
             " AND NVL(LTRIM(RTRIM(S.ժҪ)),'С��')<>'�ܷ�' and nvl(S.��ҩ��ʽ,-999)<>-1 " & strSql��������
    
    If mblnҩƷ���� = True Then
        gstrSQL = gstrSQL & " And Nvl(S.�ⷿid, [2]) + 0 = K.�ⷿid(+) And S.ҩƷid = K.ҩƷid(+) "
    End If
             
    Select Case mint��Χ
    Case 1
        gstrSQL = gstrSQL & " And S.ʵ������>=0"
    Case 2
        gstrSQL = gstrSQL & " And S.ʵ������<0"
    End Select
    
    If Trim(mstr����) <> "" Then
        If mint���� = 0 Then
            gstrSQL = gstrSQL & " And Instr([10], ',' || C.��������id || ',') > 0 And C.���˿���id=C.��������id"
        ElseIf mint���� = 1 Then
            gstrSQL = gstrSQL & " And Instr([10], ',' || C.��������id || ',') > 0 And C.���˿���id<>C.��������id"
        Else
            If mstr������ҩ��ʽ = "" Then
                gstrSQL = gstrSQL & " And Instr([10], ',' || C.���˲���ID || ',') > 0 And C.���˿���id=C.��������id"
            Else
                gstrSQL = gstrSQL & " And Instr([10], ',' || C.���˲���ID || ',') > 0 "
                If mstr������ҩ��ʽ <> mstrAllType Then
                    gstrSQL = gstrSQL & " And C.��������id Not In (Select Distinct ����id From ��������˵�� " & _
                        " Where Instr([16],',' || �������� || ',') > 0) "
                End If
            End If
        End If
    Else
        If mint���� = 0 Then
            gstrSQL = gstrSQL & " And C.���˿���id=C.��������id"
        ElseIf mint���� = 1 Then
            gstrSQL = gstrSQL & " And C.���˿���id<>C.��������id"
        Else
            If mstr������ҩ��ʽ = "" Then
                gstrSQL = gstrSQL & " And C.���˿���id=C.��������id"
            Else
                If mstr������ҩ��ʽ <> mstrAllType Then
                    gstrSQL = gstrSQL & " And C.��������id Not In (Select Distinct ����id From ��������˵�� " & _
                        " Where Instr([16],',' || �������� || ',') > 0) "
                End If
            End If
        End If
    End If
    
    If mlng�������� = 0 Then
        gstrSQL = gstrSQL & " And S.��¼״̬ = 1"
    Else
        gstrSQL = gstrSQL & " And Mod(S.��¼״̬,3)=1"
    End If
    gstrSQL = gstrSQL & strCond & IIf(mstrUse = "", "", " And Instr([13],',' || S.�÷� || ',') > 0") & str���˼����� & IIf(mstrSerchNO = "", "", " AND S.NO=[12] ") & " Order By S.No,S.����) A "
    
    gstrSQL = gstrSQL & ", (Select ҩƷid,�ⷿid,����id,�������� From ҩƷ����ƻ�  Where ״̬=0) C "
    
    '�����һ����ҩ����ҩ��
    If mbln��ʾ����ҩ�� = True Then
        strSql��ҩ�� = ",(Select a.���� ,a.No,a.���,a.������ ��ҩ�� From ҩƷ�շ���¼ a," & _
                " (Select s.����,s.No,s.���, Max(s.��¼״̬) ��¼״̬ " & _
                " From ҩƷ�շ���¼ s, δ��ҩƷ��¼ n " & _
                " Where s.No = n.No And s.���� = n.���� And Nvl(s.�ⷿid, [2]) + 0 = Nvl(n.�ⷿid, [2]) And " & _
                " Nvl(s.�ⷿid, [2]) + 0 = [2] " & _
                " AND N.�������� BETWEEN [8] AND [9] And Nvl(s.��ҩ��ʽ, -999) <> -1 And " & _
                " Mod(s.��¼״̬, 3) = 2 And s.���� In (9, 10) " & _
                " Group By s.����,s.No,s.���) b " & _
                " Where a.����=b.���� And a.No=b.No And a.���=b.��� And a.��¼״̬=b.��¼״̬) B "
        gstrSQL = gstrSQL & strSql��ҩ��
    End If
    
    gstrSQL = gstrSQL & " Where A.��ҩ����id = C.����id(+) And C.�ⷿid(+) = [2] And A.ҩƷid = C.ҩƷid(+) "
    
    If mbln��ʾ����ҩ�� = True Then
        gstrSQL = gstrSQL & " And A.���� = B.����(+) And A.No = B.No(+) And A.��� = B.���(+) "
    End If
    
    gstrSQL = gstrSQL & "  Order By a.No,a.������� "
        
    '--ˢ������--
'    on error Resume Next
    err = 0
    
    '��ʼ����¼��
    Call InitRec
    
    'δ��������¼

    Set RecBillData = zldatabase.OpenSQLRecord(gstrSQL, "��ȡ������Ϣ", _
        str������, _
        lngҩ��ID, _
        mlng����ID, _
        mstrסԺ��, _
        "%" & mstr�������� & "%", _
        mstr��ʼNO, _
        mstr����NO, _
        CDate(mstr��ʼ����_δ��), _
        CDate(mstr��������_δ��), _
        "," & mstr���� & ",", _
        mstr����, _
        mstrSerchNO, _
        "," & mstrUse & ",", _
        "," & mstrDrug & ",", _
        "," & mstr��ҩ���� & ",", _
        "," & mstr������ҩ��ʽ & ",")
    
    With RecBillData
        MnuFilePreview.Enabled = Not (.EOF)
        MnuFilePrint.Enabled = Not (.EOF)
        MnuFileExcel.Enabled = Not (.EOF)
        Tbar.Buttons("Preview").Enabled = Not (.EOF)
        Tbar.Buttons("Print").Enabled = Not (.EOF)
    End With
    
    '���ܷ�ҩʱ��һЩ����
    If Not RecBillData.EOF Then
        mstrDrawDept = ""
        mstrSendDrugId = ""
        
        'ȡ��ҩ�嵥�е���ҩ���ź�ҩƷ����Ϊ������ȡ��ҩ�����嵥
        RecBillData.MoveFirst
        Do While Not RecBillData.EOF
            If InStr("," & mstrDrawDept & ",", "," & RecBillData!��ҩ����id & ",") = 0 Then
                mstrDrawDept = IIf(mstrDrawDept = "", "", mstrDrawDept & ",") & RecBillData!��ҩ����id
            End If
            
            If InStr("," & mstrSendDrugId & ",", "," & RecBillData!ҩƷID & ",") = 0 Then
                mstrSendDrugId = IIf(mstrSendDrugId = "", "", mstrSendDrugId & ",") & RecBillData!ҩƷID
            End If
            
            RecBillData.MoveNext
        Loop
        RecBillData.MoveFirst
        
        Call Get�����嵥
    End If
    
    If ProduceInsideRecordset = False Then Exit Function
    If RefreshDataBaseOnPage(True) = False Then Exit Function
    
    Call ClearBill(Bill�ܷ�ҩ�嵥)
    Call Load�ܷ�
    Call SetMenuAndToolbarState
    
    If err <> 0 Then
        MsgBox "ˢ��ʱ����������Ԥ֪�Ĵ���", vbInformation, gstrSysName
        Exit Function
    End If
    
    Call tabShow_Click(TabShow.Tab)
    RefreshData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function ProduceInsideRecordset() As Boolean
    Dim ArrayPhysic
    Dim IntArray As Integer, lngState As Long
    
    '--�����ڲ���¼��(δ��)--
    On Error GoTo ErrHand
    err = 0
    ProduceInsideRecordset = False
   
    With RecBillData
        Do While Not .EOF
            RecChangeData.AddNew
            RecChangeData!Id = !Id
            RecChangeData!״̬ = "��ҩ"
            RecChangeData!���� = !����
            RecChangeData!��ҩ���� = !��ҩ����
            RecChangeData!��ҩ����id = !��ҩ����id
            RecChangeData!����ҽ�� = !����ҽ��
            RecChangeData!���� = IIf(NVL(!ҽ�����, 0) = 0, IIf(!�����־ = 1 Or !�����־ = 4, "������ʵ�", IIf(!���� = 9, "סԺ���ʵ�", "סԺ���ʱ�")), IIf(IsNull(!����) = True, "סԺ���ʵ�", IIf(!���� Like "0*", "����", IIf(!���� Like "1*", "����", "���ʱ�"))))
            RecChangeData!ҩƷID = !ҩƷID
            RecChangeData!λ�� = .AbsolutePosition
            RecChangeData!NO = !NO
            RecChangeData!���� = !����
            RecChangeData!����ID = !����ID
            RecChangeData!��� = !���
            RecChangeData!���� = !����
            RecChangeData!���� = IIf(IsNull(!����), "", !����)
            RecChangeData!סԺ�� = NVL(!��ʶ��)
            RecChangeData!Ʒ�� = !Ʒ��
            RecChangeData!������ = !������
            RecChangeData!Ӣ���� = !Ӣ����
            RecChangeData!��� = IIf(IsNull(!���), "", !���)
            RecChangeData!���� = IIf(IsNull(!����), "", !����)
            RecChangeData!������� = IIf(IsNull(!�������), "", !�������)
            RecChangeData!��ֵ���� = IIf(IsNull(!��ֵ����), "", !��ֵ����)
            RecChangeData!���� = IIf(IsNull(!����), 0, !����)
            RecChangeData!���� = IIf(IsNull(!����), "", !����)
            RecChangeData!���� = IIf(IsNull(!����), 0, !����)
            RecChangeData!�� = IIf(IsNull(!��), 1, !��)
            RecChangeData!ʵ������ = FormatEx(IIf(IsNull(!����), 1, !����) / !��װ, 5)
            RecChangeData!�������� = FormatEx(IIf(IsNull(!��������), 0, !��������) / !��װ, 5)
            RecChangeData!���� = FormatEx(IIf(IsNull(!����), 1, !����) / !��װ, 5) & !��λ
            RecChangeData!���� = FormatEx(!���� * !��װ, 5)
            RecChangeData!��� = Format(!���, "#####0.00;-#####0.00; ;")
            RecChangeData!����Ա = IIf(IsNull(!����Ա����), "", !����Ա����)
            RecChangeData!���� = IIf(IsNull(!����), "", FormatEx(!����, 5) & NVL(!���㵥λ))
            RecChangeData!������λ = NVL(!���㵥λ)
            RecChangeData!Ƶ�� = IIf(IsNull(!Ƶ��), "", !Ƶ��)
            RecChangeData!�÷� = IIf(IsNull(!�÷�), "", !�÷�)
            RecChangeData!˵�� = IIf(IsNull(!˵��), "", !˵��)
            If IsNull(!�Ǽ�ʱ��) Then
                RecChangeData!����ʱ�� = ""
            Else
                RecChangeData!����ʱ�� = Format(!�Ǽ�ʱ��, "yyyy-MM-dd HH:mm:ss")
            End If
            RecChangeData!��ҩ�� = IIf(IsNull(!��ҩ��), "", !��ҩ��)
            RecChangeData!����� = IIf(IsNull(!�����), "", !�����)
            RecChangeData!���շ� = !���շ�                          'δ�շѻ���ʴ�����������ҩ
            RecChangeData!����� = !�����
            RecChangeData!ҽ��id = !ҽ��id
            If mbln��ʾ����ҩ�� = True Then
                RecChangeData!��ҩ�� = !��ҩ��
            Else
                RecChangeData!��ҩ�� = ""
            End If
            RecChangeData!�ⷿ��λ = IIf(IsNull(!�ⷿ��λ), "", !�ⷿ��λ)
            RecChangeData!���ID = IIf(IsNull(!���ID), 0, !���ID)
            RecChangeData!����ID = IIf(IsNull(!����ID), 0, !����ID)
            RecChangeData!������� = !�������
            '����Ƿ�����ҩ
            lngState = 1
            If RecChangeData!���շ� = 0 Then lngState = 3
            '20020903 Modified by zyb
            '���˵���Ǿܷ����������ҩƷ�Ѿܷ���ͬʱ������ִ��״̬
            '--Begin
            If Not IsNull(!˵��) Then
                lngState = IIf(!˵�� = "�ܷ�", 2, lngState)
            End If
            '--End
            If Int����δ��˴�����ҩ = 0 Then
                If IsNull(RecChangeData!�����) Then
                    lngState = 3
                Else
                    If Trim(RecChangeData!�����) = "" Then lngState = 3
                End If
            Else
                lngState = 1
            End If
            
            RecChangeData!ִ��״̬ = lngState                        'ȱʡΪ��ҩ
            RecChangeData.Update
            If err <> 0 Then GoTo ErrHand
            .MoveNext
        Loop
    End With
    
    If err <> 0 Then
ErrHand:
        MsgBox "�����ڲ���¼��ʱ����������Ԥ֪�Ĵ���", vbInformation, gstrSysName
        InitRec
        Exit Function
    End If
    
    ProduceInsideRecordset = True
End Function

Private Function ProduceInsideSendedRecordset() As Boolean
    Dim ArrayPhysic
    Dim IntArray As Integer
    Dim dblSumSended As Double '�ѷ�����
    
    '--�����ڲ���¼��(�ѷ�)--
'    on error Resume Next
    err = 0
    ProduceInsideSendedRecordset = False
    
    With RecBillData
        Do While Not .EOF
            RecChangeSendedData.AddNew
            RecChangeSendedData!Id = !Id
            RecChangeSendedData!ҩƷID = !ҩƷID
            RecChangeSendedData!λ�� = .AbsolutePosition
            RecChangeSendedData!���� = !����
            RecChangeSendedData!���� = IIf(NVL(!ҽ�����, 0) = 0, IIf(!�����־ = 1 Or !�����־ = 4, "������ʵ�", IIf(!���� = 9, "סԺ���ʵ�", "סԺ���ʱ�")), IIf(IsNull(!����) = True, "סԺ���ʵ�", IIf(!���� Like "0*", "����", IIf(!���� Like "1*", "����", "���ʱ�"))))
            RecChangeSendedData!ִ��״̬ = 1                        'ȱʡΪ������
            RecChangeSendedData!NO = !NO
            RecChangeSendedData!���� = !����
            RecChangeSendedData!��� = !���
            RecChangeSendedData!����ID = !����ID
            RecChangeSendedData!���� = !����
            RecChangeSendedData!���� = IIf(IsNull(!����), "", !����)
            RecChangeSendedData!סԺ�� = NVL(!��ʶ��)
            RecChangeSendedData!Ʒ�� = !Ʒ��
            RecChangeSendedData!������ = !������
            RecChangeSendedData!Ӣ���� = !Ӣ����
            RecChangeSendedData!��� = IIf(IsNull(!���), "", !���)
            RecChangeSendedData!���� = IIf(IsNull(!����), "", !����)
            RecChangeSendedData!������� = NVL(!�������)
            RecChangeSendedData!���� = IIf(IsNull(!����), 0, !����)
            RecChangeSendedData!���� = IIf(IsNull(!����), 0, !����)
            RecChangeSendedData!���� = IIf(IsNull(!����), "", !����)
            RecChangeSendedData!Ч�� = IIf(IsNull(!Ч��), "", !Ч��)
            RecChangeSendedData!�� = IIf(IsNull(!��), 1, !��)
            RecChangeSendedData!���� = FormatEx(IIf(IsNull(!����), 1, !����) / !��װ, 5) & !��λ
            If Chk�嵥.Value = 0 Or !�ɲ��� <> 1 Then
                RecChangeSendedData!������ = FormatEx(IIf(IsNull(!��������), 1, !��������) / !��װ, 5)
                RecChangeSendedData!׼���� = FormatEx(IIf(IsNull(!׼����), 1, !׼����) / !��װ, 5)
                RecChangeSendedData!��ҩ�� = FormatEx(IIf(IsNull(!׼����), 1, !׼����) / !��װ, 5)
            Else
                dblSumSended = GetSumSended(!����, !NO, !ҩƷID, !���)
                RecChangeSendedData!������ = FormatEx((IIf(IsNull(!����), 1, !����) - dblSumSended) / !��װ, 5)
                RecChangeSendedData!׼���� = FormatEx(dblSumSended / !��װ, 5)
                RecChangeSendedData!��ҩ�� = FormatEx(dblSumSended / !��װ, 5)
            End If
            RecChangeSendedData!��λ = !��λ
            RecChangeSendedData!���� = FormatEx(!���� * !��װ, 5)
            RecChangeSendedData!��� = !���
            RecChangeSendedData!���� = IIf(IsNull(!����), "", FormatEx(!����, 5) & NVL(!���㵥λ))
            RecChangeSendedData!������λ = NVL(!���㵥λ)
            RecChangeSendedData!Ƶ�� = IIf(IsNull(!Ƶ��), "", !Ƶ��)
            RecChangeSendedData!�÷� = IIf(IsNull(!�÷�), "", !�÷�)
            RecChangeSendedData!˵�� = IIf(IsNull(!˵��), "", !˵��)
            RecChangeSendedData!����Ա = IIf(IsNull(!�����), "", !�����)
            RecChangeSendedData!��ҩʱ�� = IIf(IsNull(!��ҩʱ��), "", !��ҩʱ��)
            If Val(!ת��) = 1 Then
                RecChangeSendedData!�ɲ��� = -1
            Else
                RecChangeSendedData!�ɲ��� = IIf(IsNull(!�ɲ���), 0, !�ɲ���)
            End If
            RecChangeSendedData!����� = !�����
            RecChangeSendedData!ҽ��id = !ҽ��id
            RecChangeSendedData!��ҩ�� = !��ҩ��
            RecChangeSendedData!ʵ������ = !׼����
            RecChangeSendedData!�ⷿ��λ = IIf(IsNull(!�ⷿ��λ), "", !�ⷿ��λ)
            RecChangeSendedData!ת�� = Val(!ת��)
            If Chk�嵥.Value = 1 Then
                RecChangeSendedData!���ID = 0
            Else
                RecChangeSendedData!���ID = IIf(IsNull(!���ID), 0, !���ID)
            End If
            
            .MoveNext
        Loop
    End With
    
    If err <> 0 Then
        MsgBox "�����ڲ���¼��ʱ����������Ԥ֪�Ĵ���", vbInformation, gstrSysName
        Call InitRec
        Exit Function
    End If
    ProduceInsideSendedRecordset = True
End Function

Private Function RefreshDataBaseOnPage(Optional ByVal BlnRefsh As Boolean = False) As Boolean
    Dim lngRows As Long
    Dim strCaption As String
    '--�����û�ѡ���ҳ�������ʾ����--
    On Error Resume Next
    err = 0
    RefreshDataBaseOnPage = False
    
    '��ս�Ҫ��ʾ�Ŀؼ��е�����
    If InStr(1, "1,2,3", TabShow.Tab) <> 0 Then ClearCons
    If Blnˢ��δ��ҩ�嵥 And (TabShow.Tab = 0 Or TabShow.Tab = 1) Then        '��һЩ׼��
        If Bln����� Then
            '����湻��
            Call CheckStock
        End If
        Bln����� = False
    End If
    
    '����ҳ���ʼ��������
    stbThis.Panels(2).Text = ""
    Select Case TabShow.Tab
    Case 0
        If Blnˢ��δ��ҩ�嵥 Then
            Call ClearCons
            If LoadDataInBillδ��ҩ�嵥 = False Then GoTo ErrHand
            Call SetGroup(Billδ��ҩ�嵥, True)
        End If
        lngRows = lngδ��ҩ��¼
        strCaption = "��δ��ҩƷ��¼"
    Case 1
        If LoadDataInBill�����嵥 = False Then GoTo ErrHand
'        lngRows = IIf(Billδ��ҩ�嵥.TextMatrix(Billδ��ҩ�嵥.Rows - 1, 0) = "", Billδ��ҩ�嵥.Rows - 2, Billδ��ҩ�嵥.Rows - 1)
'        strCaption = "��δ��ҩƷ��¼"
        lngRows = IIf(Bill���ܷ�ҩ.TextMatrix(Bill���ܷ�ҩ.rows - 1, 0) = "", 0, lng�����嵥����)
        strCaption = "��ҩƷ���ܼ�¼"
    Case 2
        If LoadDataInBillȱҩ�嵥 = False Then GoTo ErrHand
        lngRows = IIf(Billȱҩ�嵥.TextMatrix(Billȱҩ�嵥.rows - 1, 0) = "", Billȱҩ�嵥.rows - 2, Billȱҩ�嵥.rows - 1)
        strCaption = "��ȱҩ��¼"
    Case 3
        If LoadDataInBill�ܷ��嵥 = False Then GoTo ErrHand
        lngRows = IIf(Bill�ܷ�ҩ�嵥.TextMatrix(Bill�ܷ�ҩ�嵥.rows - 1, 0) = "", Bill�ܷ�ҩ�嵥.rows - 2, Bill�ܷ�ҩ�嵥.rows - 1)
        strCaption = "���ܷ�ҩƷ��¼"
    Case 4
        '�ֹ�Ԥ���
        lngRows = IIf(Bill�ѷ�ҩ�嵥.TextMatrix(Bill�ѷ�ҩ�嵥.rows - 1, 0) = "", Bill�ѷ�ҩ�嵥.rows - 2, Bill�ѷ�ҩ�嵥.rows - 1)
        strCaption = "���ѷ�ҩƷ��¼"
    End Select
    
    If err <> 0 Then
ErrHand:
        MsgBox "��ʾ[" & TabShow.TabCaption(TabShow.Tab) & "]ҳ�������ʱ����������Ԥ֪�Ĵ���", vbInformation, gstrSysName
        Exit Function
    End If
    If lngRows <> 0 Then stbThis.Panels(2).Text = "��ǰ����" & lngRows & strCaption
    If mlng���ܷ�ҩ�� > 0 Then
        stbThis.Panels(2).Text = stbThis.Panels(2).Text & "[�ϴη�ҩ�ţ�" & mlng���ܷ�ҩ�� & "]"
    End If
    RefreshDataBaseOnPage = True
End Function

Private Function ClearCons()
    '--����ҳ�������ؿؼ�����ʾ����--
    Select Case TabShow.Tab
    Case 0
        Call ClearBill(Billδ��ҩ�嵥)
    Case 1
        Call ClearBill(Bill���ܷ�ҩ)
        Call ClearBill(Bill��ҩ����)
        Bill��ҩ����.Visible = False
        Bill���ܷ�ҩ.Height = TabShow.Height - TabShow.TabHeight - 120
    Case 2
        Call ClearBill(Billȱҩ�嵥)
    Case 3
        'ֻ�����ѡ��δ�ܷ���ҩƷ��¼
        Dim i As Integer, j As Integer
        With Bill�ܷ�ҩ�嵥
            For i = 1 To .rows - 1
                If Trim(.TextMatrix(i, 1)) = "" Then
                    If i = .rows - 1 Then
                        For j = 0 To .Cols - 1
                            .TextMatrix(i, j) = ""
                        Next
                    Else
                        .RemoveItem i: i = i - 1
                    End If
                End If
            Next
        End With
    Case 4
        '�ֹ�����
        Call ClearBill(Bill�ѷ�ҩ�嵥)
    End Select
End Function

Private Function LoadDataInBillδ��ҩ�嵥() As Boolean
    Dim blnEnable As Boolean, lngRow As Long, intCol As Long
    Dim strCompare As String, strColumn As String, strValue As String
    Dim dbl�ϼƽ�� As Double, dblС�ƽ�� As Double
    Dim lngColor As Long
    Dim strOrder As String
    
    '--���δ��ҩ�嵥--
    On Error Resume Next
    err = 0
    blnEnable = False
    LoadDataInBillδ��ҩ�嵥 = False
    
    With Billδ��ҩ�嵥
        .MousePointer = 11
        .Redraw = False
    End With
    
    dblС�ƽ�� = 0: dbl�ϼƽ�� = 0
    If InStr(1, str����_δ��ҩ, strAsc) <> 0 Then
        strColumn = Mid(str����_δ��ҩ, 1, InStr(1, str����_δ��ҩ, strAsc) - 1)
    Else
        strColumn = Mid(str����_δ��ҩ, 1, InStr(1, str����_δ��ҩ, strDesc) - 1)
    End If
    strColumn = Trim(strColumn)
    
    '���������н�������
    With RecChangeData
        lngδ��ҩ��¼ = .RecordCount
        If .RecordCount <> 0 Then .MoveFirst
        strOrder = GetOrder(str����_δ��ҩ)
        '�����NO������ͬʱ�����ID���򣬱������÷���
        strOrder = IIf(InStr(strOrder, "NO") > 0, strOrder & ",���ID" & IIf(InStr(strOrder, " ASC") > 0, " ASC", " DESC"), strOrder)
        .Sort = strOrder
        Do While Not .EOF
'            Strִ��״̬ = IIf(!ִ��״̬ = 0, "ȱҩ", IIf(!ִ��״̬ = 1, "��ҩ", IIf(!ִ��״̬ = 2, "�ܷ�", "������")))
            If !˵�� <> "�ܷ�" Then
                blnEnable = True
                Billδ��ҩ�嵥.MergeRow(Billδ��ҩ�嵥.rows - 1) = False
                
                strValue = IIf(IsNull(.Fields(strColumn).Value), "", .Fields(strColumn).Value)
                If strCompare <> strValue And strCompare <> "" Then
                    '���Ӻϼ���
                    Call AddCollect(dblС�ƽ��, "С��")
                    dblС�ƽ�� = 0
                End If
                
                '��ֵ
                strCompare = IIf(IsNull(.Fields(strColumn).Value), "", .Fields(strColumn).Value)
                
                '�������
                Billδ��ҩ�嵥.TextMatrix(Billδ��ҩ�嵥.rows - 1, ����_δ��ҩ�嵥.����) = !����
                Billδ��ҩ�嵥.TextMatrix(Billδ��ҩ�嵥.rows - 1, ����_δ��ҩ�嵥.����ҽ��) = IIf(IsNull(!����ҽ��), "", !����ҽ��)
                Billδ��ҩ�嵥.TextMatrix(Billδ��ҩ�嵥.rows - 1, ����_δ��ҩ�嵥.״̬) = !״̬
                Billδ��ҩ�嵥.TextMatrix(Billδ��ҩ�嵥.rows - 1, ����_δ��ҩ�嵥.����) = !����
                Billδ��ҩ�嵥.TextMatrix(Billδ��ҩ�嵥.rows - 1, ����_δ��ҩ�嵥.NO) = !NO
                Billδ��ҩ�嵥.TextMatrix(Billδ��ҩ�嵥.rows - 1, ����_δ��ҩ�嵥.����Ա) = !����Ա
                Billδ��ҩ�嵥.TextMatrix(Billδ��ҩ�嵥.rows - 1, ����_δ��ҩ�嵥.����) = IIf(IsNull(!����), "", !����)
                Billδ��ҩ�嵥.TextMatrix(Billδ��ҩ�嵥.rows - 1, ����_δ��ҩ�嵥.����) = !����
                Billδ��ҩ�嵥.TextMatrix(Billδ��ҩ�嵥.rows - 1, ����_δ��ҩ�嵥.סԺ��) = !סԺ��
                Billδ��ҩ�嵥.TextMatrix(Billδ��ҩ�嵥.rows - 1, ����_δ��ҩ�嵥.ҩƷ����) = !Ʒ��
                Billδ��ҩ�嵥.TextMatrix(Billδ��ҩ�嵥.rows - 1, ����_δ��ҩ�嵥.������) = IIf(IsNull(!������), "", !������)
                Billδ��ҩ�嵥.TextMatrix(Billδ��ҩ�嵥.rows - 1, ����_δ��ҩ�嵥.Ӣ����) = IIf(IsNull(!Ӣ����), "", !Ӣ����)
                Billδ��ҩ�嵥.TextMatrix(Billδ��ҩ�嵥.rows - 1, ����_δ��ҩ�嵥.���) = IIf(IsNull(!���), "", !���)
                Billδ��ҩ�嵥.TextMatrix(Billδ��ҩ�嵥.rows - 1, ����_δ��ҩ�嵥.����) = !����
                Billδ��ҩ�嵥.TextMatrix(Billδ��ҩ�嵥.rows - 1, ����_δ��ҩ�嵥.����) = !����
                Billδ��ҩ�嵥.TextMatrix(Billδ��ҩ�嵥.rows - 1, ����_δ��ҩ�嵥.��) = !��
                Billδ��ҩ�嵥.TextMatrix(Billδ��ҩ�嵥.rows - 1, ����_δ��ҩ�嵥.����) = !����
                Billδ��ҩ�嵥.TextMatrix(Billδ��ҩ�嵥.rows - 1, ����_δ��ҩ�嵥.����) = !����
                Billδ��ҩ�嵥.TextMatrix(Billδ��ҩ�嵥.rows - 1, ����_δ��ҩ�嵥.���) = Format(!���, "#####0.00;-#####0.00; ;")
                Billδ��ҩ�嵥.TextMatrix(Billδ��ҩ�嵥.rows - 1, ����_δ��ҩ�嵥.����) = !����
                Billδ��ҩ�嵥.TextMatrix(Billδ��ҩ�嵥.rows - 1, ����_δ��ҩ�嵥.Ƶ��) = !Ƶ��
                Billδ��ҩ�嵥.TextMatrix(Billδ��ҩ�嵥.rows - 1, ����_δ��ҩ�嵥.�÷�) = !�÷�
                Billδ��ҩ�嵥.TextMatrix(Billδ��ҩ�嵥.rows - 1, ����_δ��ҩ�嵥.����ʱ��) = Format(!����ʱ��, "yyyy-MM-dd HH:mm:ss")
                Billδ��ҩ�嵥.TextMatrix(Billδ��ҩ�嵥.rows - 1, ����_δ��ҩ�嵥.˵��) = !˵��
                Billδ��ҩ�嵥.TextMatrix(Billδ��ҩ�嵥.rows - 1, ����_δ��ҩ�嵥.����) = !����
                Billδ��ҩ�嵥.TextMatrix(Billδ��ҩ�嵥.rows - 1, ����_δ��ҩ�嵥.ҽ��id) = !ҽ��id
                Billδ��ҩ�嵥.TextMatrix(Billδ��ҩ�嵥.rows - 1, ����_δ��ҩ�嵥.��ҩ��) = IIf(IsNull(!��ҩ��), "", !��ҩ��)
                Billδ��ҩ�嵥.TextMatrix(Billδ��ҩ�嵥.rows - 1, ����_δ��ҩ�嵥.�ⷿ��λ) = IIf(IsNull(!�ⷿ��λ), "", !�ⷿ��λ)
                Billδ��ҩ�嵥.TextMatrix(Billδ��ҩ�嵥.rows - 1, ����_δ��ҩ�嵥.���ID) = !���ID
                Billδ��ҩ�嵥.TextMatrix(Billδ��ҩ�嵥.rows - 1, ����_δ��ҩ�嵥.ҩƷID) = !ҩƷID
                Billδ��ҩ�嵥.TextMatrix(Billδ��ҩ�嵥.rows - 1, ����_δ��ҩ�嵥.��ҩ����) = !��ҩ����
                Billδ��ҩ�嵥.TextMatrix(Billδ��ҩ�嵥.rows - 1, ����_δ��ҩ�嵥.��ҩ����id) = !��ҩ����id
                                
                Billδ��ҩ�嵥.ColWidth(����_δ��ҩ�嵥.�����) = IIf(Not mblnStarPass, 0, 240)
                Billδ��ҩ�嵥.ColWidth(����_δ��ҩ�嵥.����) = 0
                Billδ��ҩ�嵥.ColWidth(����_δ��ҩ�嵥.ҽ��id) = 0
                
                If !����� <> -1 Then
                    BlnEnterCell = False
                    Billδ��ҩ�嵥.Row = Billδ��ҩ�嵥.rows - 1
                    Billδ��ҩ�嵥.Col = 0
                    Set Billδ��ҩ�嵥.CellPicture = imgPass.ListImages(Val(!�����) + 1).Picture
                    Billδ��ҩ�嵥.CellPictureAlignment = 4
                    BlnEnterCell = True
                End If

                '����ҩƷ������ʾ
                BlnEnterCell = False
                If InStr(";����ҩ;����ҩ;����I��;����II��;", NVL(!�������)) > 0 And NVL(!�������) <> "" Then
                    Billδ��ҩ�嵥.Col = ����_δ��ҩ�嵥.ҩƷ����
                    Billδ��ҩ�嵥.Row = Billδ��ҩ�嵥.rows - 1
                    Billδ��ҩ�嵥.CellFontBold = True
                End If

                If mblnҩƷ���� = True Then
                    If !������� = 0 Then
                        Billδ��ҩ�嵥.Row = Billδ��ҩ�嵥.rows - 1
                        For intCol = 0 To Billδ��ҩ�嵥.Cols - 1
                            Billδ��ҩ�嵥.Col = intCol
                            Billδ��ҩ�嵥.CellForeColor = mlng��ɫ
                        Next
                        Billδ��ҩ�嵥.RowData(Billδ��ҩ�嵥.Row) = mlng��ɫ
                    End If
                End If
                
                '���÷�ҩ״̬�ı���ɫ
                lngColor = IIf(!״̬ = "��ҩ", glngSendBlkColor, glngOtherBlkColor)
                Billδ��ҩ�嵥.Row = Billδ��ҩ�嵥.rows - 1
                For intCol = 0 To Billδ��ҩ�嵥.Cols - 1
                    Billδ��ҩ�嵥.Col = intCol
                    Billδ��ҩ�嵥.CellBackColor = lngColor
                Next
                
                BlnEnterCell = True
                
                dblС�ƽ�� = dblС�ƽ�� + Val(!���)
                dbl�ϼƽ�� = dbl�ϼƽ�� + Val(!���)
                !λ�� = Billδ��ҩ�嵥.rows - 1
                .Update
                Billδ��ҩ�嵥.rows = Billδ��ҩ�嵥.rows + 1
                
                Billδ��ҩ�嵥.ColAlignment(����_δ��ҩ�嵥.ҩƷ����) = 1
                
            End If
            .MoveNext
        Loop
        If .RecordCount <> 0 Then
            If strCompare <> "" Then
                '���Ӻϼ���
                Call AddCollect(dblС�ƽ��, "С��")
                Call AddCollect(dbl�ϼƽ��)
            End If
            .MoveFirst
        End If
        
        '�ϲ�
        Billδ��ҩ�嵥.MergeCells = flexMergeFree
        For lngRow = 0 To Billδ��ҩ�嵥.rows - 1
            If InStr(1, "С��,�ϼ�", Billδ��ҩ�嵥.TextMatrix(lngRow, ����_δ��ҩ�嵥.����)) <> 0 Then
                Billδ��ҩ�嵥.MergeRow(lngRow) = True
            End If
        Next
        
        Call SetMenu(blnEnable)
    End With
    
    With Billδ��ҩ�嵥
        .MousePointer = 0
        .Redraw = True
        .Row = 1
        .Col = 1
    End With
    
    If err <> 0 Then Exit Function
    
    LoadDataInBillδ��ҩ�嵥 = True
    Blnˢ��δ��ҩ�嵥 = False
End Function

Private Function LoadDataInBill�����嵥() As Boolean
    Dim LngFindPhysicID As Long, strPartName As String
    Dim LngLocate As Long, blnEnable As Boolean
    Dim dbl�ϼƽ�� As Double, dbl���Һϼ� As Double
    Dim lng���� As Long
    Dim n As Integer
    '--�������嵥--
'    on error Resume Next
    err = 0
    blnEnable = False
    LoadDataInBill�����嵥 = False
    Bill���ܷ�ҩ.Redraw = False
    
    LngFindPhysicID = 0
    lng���� = 0
    strPartName = ""
    dbl���Һϼ� = 0
    dbl�ϼƽ�� = 0
    
    With RecChangeData
        If .RecordCount = 0 Then
            LoadDataInBill�����嵥 = True
            Call SetMenu(blnEnable)
            Bill���ܷ�ҩ.Redraw = True
            Exit Function
        End If
        
        .MoveFirst
        
        If Chk�嵥.Value = 0 Then   '��ҩƷ���ƻ���
            .Sort = IIf(Lng������ʾ = 1, "��ҩ���� Asc,", "") & "ҩƷID Asc"
        Else    '��ҩƷ���λ���
            .Sort = IIf(Lng������ʾ = 1, "��ҩ���� Asc,", "") & "ҩƷID Asc" & ",���� Asc"
        End If
        '�ֹ��������ʾ
        Do While Not .EOF
            If Lng������ʾ = 1 Then
                If !ִ��״̬ = 1 And !��ҩ���� <> strPartName And CheckGroupSend(!���ID) = True Then
                    LngLocate = !λ��
                    blnEnable = True
                    strPartName = !��ҩ����
                    If LngFindPhysicID <> 0 And IIf(Chk�嵥.Value = 0, True, lng���� <> 0) Then
                        Bill���ܷ�ҩ.rows = Bill���ܷ�ҩ.rows + 1
                        Call AddCollect(dbl���Һϼ�, "С��")
                        dbl���Һϼ� = 0
                        LngFindPhysicID = 0
                        lng���� = 0
                    End If
                End If
            End If
            lng�����嵥���� = 0
            If !ִ��״̬ = 1 And (!ҩƷID <> LngFindPhysicID Or IIf(Chk�嵥.Value = 0, False, !���� <> lng����)) And CheckGroupSend(!���ID) = True Then 'ֻ���ܷ�ҩ�ļ�¼
                LngLocate = !λ��
                blnEnable = True
                With Bill���ܷ�ҩ
                    If Trim(.TextMatrix(.rows - 1, 0)) <> "" Then .rows = .rows + 1
                    .MergeRow(.rows - 1) = False
                    If Lng������ʾ = 1 Then
                        .Row = .rows - 1: .Col = 1: .CellAlignment = 1
                        .TextMatrix(.rows - 1, ����_���һ����嵥.����) = RecChangeData!����
                        .TextMatrix(.rows - 1, ����_���һ����嵥.ҩƷ����) = RecChangeData!Ʒ��
                        .TextMatrix(.rows - 1, ����_���һ����嵥.���) = RecChangeData!���
                        .TextMatrix(.rows - 1, ����_���һ����嵥.����) = RecChangeData!����
                        .TextMatrix(.rows - 1, ����_���һ����嵥.����) = RecChangeData!����
                        .TextMatrix(.rows - 1, ����_���һ����嵥.Ӧ������) = FormatEx(RecChangeData!ʵ������ * RecChangeData!��, 5)
                        .TextMatrix(.rows - 1, ����_���һ����嵥.��λ) = Right(RecChangeData!����, 1)
                        .TextMatrix(.rows - 1, ����_���һ����嵥.����) = FormatEx(RecChangeData!����, 5)
                        .TextMatrix(.rows - 1, ����_���һ����嵥.���) = Format(RecChangeData!���, "#####0.00;-#####0.00; ;")
                        .TextMatrix(.rows - 1, ����_���һ����嵥.����) = RecChangeData!����
                        .TextMatrix(.rows - 1, ����_���һ����嵥.����ID) = RecChangeData!����ID
                        .TextMatrix(.rows - 1, ����_���һ����嵥.ҩƷID) = RecChangeData!ҩƷID
                        .TextMatrix(.rows - 1, ����_���һ����嵥.��ҩ����) = RecChangeData!��ҩ����
                        .TextMatrix(.rows - 1, ����_���һ����嵥.��ҩ����id) = RecChangeData!��ҩ����id
                        .TextMatrix(.rows - 1, ����_���һ����嵥.��������) = FormatEx(RecChangeData!��������, 5)
                        
                        If mbln���ܷ�ҩ = True Then
                            .TextMatrix(.rows - 1, ����_���һ����嵥.��������) = FormatEx(Get��������(RecChangeData!��ҩ����id, RecChangeData!ҩƷID), 5)
                        End If
                        
                    Else
                        .Row = .rows - 1: .Col = 0: .CellAlignment = 1
                        .TextMatrix(.rows - 1, ����_�����嵥.ҩƷ����) = RecChangeData!Ʒ��
                        .TextMatrix(.rows - 1, ����_�����嵥.���) = RecChangeData!���
                        .TextMatrix(.rows - 1, ����_�����嵥.����) = RecChangeData!����
                        .TextMatrix(.rows - 1, ����_�����嵥.����) = RecChangeData!����
                        .TextMatrix(.rows - 1, ����_�����嵥.����) = FormatEx(RecChangeData!ʵ������ * RecChangeData!��, 5)
                        .TextMatrix(.rows - 1, ����_�����嵥.��λ) = Right(RecChangeData!����, 1)
                        .TextMatrix(.rows - 1, ����_�����嵥.����) = FormatEx(RecChangeData!����, 5)
                        .TextMatrix(.rows - 1, ����_�����嵥.���) = Format(RecChangeData!���, "#####0.00;-#####0.00; ;")
                    End If
                    
                    '����ҩƷ������ʾ
                    If InStr(";����ҩ;����ҩ;����I��;����II��;", NVL(RecChangeData!�������)) > 0 And NVL(RecChangeData!�������) <> "" Then
                        .Row = .rows - 1
                        .Col = IIf(Lng������ʾ = 1, ����_���һ����嵥.ҩƷ����, ����_�����嵥.ҩƷ����)
                        .CellFontBold = True
                    End If
                    
                    dbl���Һϼ� = dbl���Һϼ� + Val(RecChangeData!���)
                    dbl�ϼƽ�� = dbl�ϼƽ�� + Val(RecChangeData!���)
                    
                End With
                LngFindPhysicID = !ҩƷID
                lng���� = !����
                
                '����
                If Not .EOF Then .MoveNext
                Do While Not .EOF
                    If LngFindPhysicID = !ҩƷID And IIf(Chk�嵥.Value = 0, True, lng���� = !����) And !ִ��״̬ = 1 And CheckGroupSend(!���ID) = True Then
                        If strPartName <> !��ҩ���� And Lng������ʾ = 1 Then Exit Do
                        With Bill���ܷ�ҩ
                            If Lng������ʾ = 0 Then
                                .TextMatrix(.rows - 1, ����_�����嵥.����) = FormatEx(Val(.TextMatrix(.rows - 1, ����_�����嵥.����)) + (RecChangeData!ʵ������ * RecChangeData!��), 5)
                                .TextMatrix(.rows - 1, ����_�����嵥.���) = Format(Val(.TextMatrix(.rows - 1, ����_�����嵥.���)) + Val(RecChangeData!���), "#####0.00;-#####0.00; ;")
                            Else
                                .TextMatrix(.rows - 1, ����_���һ����嵥.Ӧ������) = FormatEx(Val(.TextMatrix(.rows - 1, ����_���һ����嵥.Ӧ������)) + (RecChangeData!ʵ������ * RecChangeData!��), 5)
                                .TextMatrix(.rows - 1, ����_���һ����嵥.���) = Format(Val(.TextMatrix(.rows - 1, ����_���һ����嵥.���)) + Val(RecChangeData!���), "#####0.00;-#####0.00; ;")
                            End If
                            dbl���Һϼ� = dbl���Һϼ� + Val(RecChangeData!���)
                            dbl�ϼƽ�� = dbl�ϼƽ�� + Val(RecChangeData!���)
                        End With
                    End If
                    .MoveNext
                Loop
                .MoveFirst
                .Find "λ��=" & LngLocate
            End If
            
            If Not .EOF Then
                .MoveNext
            Else
                Exit Do
            End If
        Loop
        
        'ͳ��ʵ�ʷ�ҩ����
        If Lng������ʾ = 1 Then
            For n = 1 To Bill���ܷ�ҩ.rows - 1
                With Bill���ܷ�ҩ
                    If .TextMatrix(n, 0) <> "С��" Then
                        'Ӧ������С��������������ʵ��Ϊ��������ʾ���ҽ�ʵ����ҩ����������Ϊ0
                        If Val(.TextMatrix(n, ����_���һ����嵥.Ӧ������)) - Val(.TextMatrix(n, ����_���һ����嵥.��������)) < 0 Then
                            .TextMatrix(n, ����_���һ����嵥.ʵ������) = FormatEx(Val(.TextMatrix(n, ����_���һ����嵥.Ӧ������)) - Val(.TextMatrix(n, ����_���һ����嵥.��������)), 5)
                            .TextMatrix(n, ����_���һ����嵥.��������) = 0
                        Else
                            '�����������Ϊ0��ʵ�������򰴲���������䣬������ʵ��Ӧ���������㣨ʵ��Ӧ����Ӧ������������������
                            If Val(.TextMatrix(n, ����_���һ����嵥.��������)) = 0 Then
                                If int��ҩ���� = 0 Then
                                    .TextMatrix(n, ����_���һ����嵥.ʵ������) = FormatEx(Val(.TextMatrix(n, ����_���һ����嵥.Ӧ������)) - Val(.TextMatrix(n, ����_���һ����嵥.��������)), 5)
                                ElseIf int��ҩ���� = 1 Then
                                    .TextMatrix(n, ����_���һ����嵥.ʵ������) = 0
                                Else
                                    .TextMatrix(n, ����_���һ����嵥.ʵ������) = FormatEx(Int(Val(.TextMatrix(n, ����_���һ����嵥.Ӧ������)) - Val(.TextMatrix(n, ����_���һ����嵥.��������))), 5)
                                End If
                                .TextMatrix(n, ����_���һ����嵥.��������) = FormatEx(Val(.TextMatrix(n, ����_���һ����嵥.Ӧ������)) - Val(.TextMatrix(n, ����_���һ����嵥.��������)) - Val(.TextMatrix(n, ����_���һ����嵥.ʵ������)), 5)
                            Else
                            '�������������Ϊ0����ҩƷ����ƻ�ȡֵ��������ʵ��Ӧ���������㣨ʵ��Ӧ����Ӧ������������������
                                If Val(.TextMatrix(n, ����_���һ����嵥.��������)) > Val(.TextMatrix(n, ����_���һ����嵥.Ӧ������)) - Val(.TextMatrix(n, ����_���һ����嵥.��������)) Then
                                    '��������������ʵ��Ӧ��������������������ʵ��Ӧ������
                                    .TextMatrix(n, ����_���һ����嵥.��������) = FormatEx(Val(.TextMatrix(n, ����_���һ����嵥.Ӧ������)) - Val(.TextMatrix(n, ����_���һ����嵥.��������)), 5)
                                End If
                                
                                'ʵ��������Ӧ��������������������������
                                .TextMatrix(n, ����_���һ����嵥.ʵ������) = FormatEx(Int(Val(.TextMatrix(n, ����_���һ����嵥.Ӧ������)) - Val(.TextMatrix(n, ����_���һ����嵥.��������)) - Val(.TextMatrix(n, ����_���һ����嵥.��������))), 5)
                            End If
                        End If
                        
                        .Row = n
                        .Col = ����_���һ����嵥.ʵ������
                        .CellFontBold = True
                        If Val(.TextMatrix(n, ����_���һ����嵥.ʵ������)) < 0 Then
                            .CellForeColor = vbRed
                        ElseIf Val(.TextMatrix(n, ����_���һ����嵥.ʵ������)) > 0 Then
                            .CellForeColor = vbBlue
                        End If
                    End If
                End With
            Next
        End If
                
        lng�����嵥���� = Bill���ܷ�ҩ.rows - 1
        If Lng������ʾ = 1 And dbl���Һϼ� <> 0 Then
            Bill���ܷ�ҩ.rows = Bill���ܷ�ҩ.rows + 1
            Call AddCollect(dbl���Һϼ�, "С��")
        End If
        Call SetMenu(blnEnable)
        
        .Sort = "NO Asc"
        
        '������ǰ����λ��ܷ�ҩ���Ͳ���ʾ������
        With Bill���ܷ�ҩ
            .ColWidth(IIf(Lng������ʾ = 1, ����_���һ����嵥.����, ����_�����嵥.����)) = IIf(Chk�嵥.Value = 1, 1200, 0)
        End With
        
    End With
    
    With Bill���ܷ�ҩ
        If .TextMatrix(.rows - 1, 0) <> "" Then .rows = .rows + 1
        If dbl�ϼƽ�� <> 0 Then Call AddCollect(dbl�ϼƽ��)
        For LngLocate = 1 To .rows - 1
            If InStr(1, "С��,�ϼ�", .TextMatrix(LngLocate, 0)) <> 0 Then
                .MergeCells = flexMergeFree
                .MergeRow(LngLocate) = True
            End If
        Next
        .Row = 1: .Col = 0
    End With
    
    
    Bill���ܷ�ҩ.Redraw = True
    If err <> 0 Then Exit Function
    LoadDataInBill�����嵥 = True
End Function

Private Function LoadDataInBillȱҩ�嵥() As Boolean
    Dim LngRecords As Long, blnEnable As Boolean
    '--���ȱҩ�嵥--
    Debug.Print Now
    On Error Resume Next
    err = 0
    blnEnable = False
    LoadDataInBillȱҩ�嵥 = False
    
    With RecChangeData
        If .RecordCount <> 0 Then
            .MoveFirst
        Else
            LoadDataInBillȱҩ�嵥 = True: Call SetMenu(blnEnable): Exit Function
        End If
        
        '�ֹ��������ʾ
        LngRecords = 0
        Do While Not .EOF
            If !ִ��״̬ = 0 Then   'ֻ��ʾȱҩ��¼
                If Not IsNull(!��ҩ��) Then
                    If !��ҩ�� <> "���ŷ�ҩ" Then
                        blnEnable = True
                        With Billȱҩ�嵥
                            If Trim(.TextMatrix(.rows - 1, 0)) <> "" Then .rows = .rows + 1
                            .TextMatrix(.rows - 1, 0) = RecChangeData!����
                            .TextMatrix(.rows - 1, 1) = RecChangeData!NO
                            .TextMatrix(.rows - 1, 2) = RecChangeData!����
                            .TextMatrix(.rows - 1, 3) = IIf(IsNull(RecChangeData!����), "", RecChangeData!����)
                            .TextMatrix(.rows - 1, 4) = IIf(IsNull(RecChangeData!����), "", RecChangeData!����)
                            .TextMatrix(.rows - 1, 5) = RecChangeData!Ʒ��
                            .TextMatrix(.rows - 1, 6) = IIf(IsNull(RecChangeData!���), "", RecChangeData!���)
                            .TextMatrix(.rows - 1, 7) = IIf(IsNull(RecChangeData!����), "", RecChangeData!����)
                            .TextMatrix(.rows - 1, 8) = IIf(IsNull(RecChangeData!����), "", RecChangeData!����)
                            .TextMatrix(.rows - 1, 9) = FormatEx(RecChangeData!ʵ������ * RecChangeData!��, 5) & Right(RecChangeData!����, 1)
                            .TextMatrix(.rows - 1, 10) = FormatEx(RecChangeData!����, 5)
                            .TextMatrix(.rows - 1, 11) = Format(RecChangeData!���, "#####0.00;-#####0.00; ;")
                            
                            '����ҩƷ������ʾ
                            If InStr(";����ҩ;����ҩ;����I��;����II��;", NVL(RecChangeData!�������)) > 0 And NVL(RecChangeData!�������) <> "" Then
                                .Row = .rows - 1: .Col = 5
                                .CellFontBold = True
                            End If
                        End With
                        LngRecords = LngRecords + 1
                    End If
                End If
            End If
            If Not .EOF Then
                .MoveNext
            Else
                Exit Do
            End If
        Loop
        Call SetMenu(blnEnable)
    End With
    
    If err <> 0 Then Exit Function
    LoadDataInBillȱҩ�嵥 = True
End Function

Private Function LoadDataInBill�ܷ��嵥() As Boolean
    Dim lngRow As Long, blnEnable As Boolean
    
    '--���ܷ��嵥--
    On Error Resume Next
    err = 0
    blnEnable = False
    LoadDataInBill�ܷ��嵥 = False
    
    'װ���趨Ϊ�ܷ����嵥(δ�������ݿ�)
    With RecChangeData
        If .RecordCount <> 0 Then
            .MoveFirst
        Else
            LoadDataInBill�ܷ��嵥 = True: Call SetMenu(blnEnable): Exit Function
        End If
        
        '�ֹ��������ʾ
        Do While Not .EOF
            If !ִ��״̬ = 2 Then   'ֻ��ʾ�ܷ���¼
                blnEnable = True
                With Bill�ܷ�ҩ�嵥
                    If Trim(.TextMatrix(.rows - 1, 1)) <> "" Then .rows = .rows + 1
                    .TextMatrix(.rows - 1, 0) = RecChangeData!����
                    .TextMatrix(.rows - 1, 1) = ""
                    .TextMatrix(.rows - 1, 2) = RecChangeData!NO
                    .TextMatrix(.rows - 1, 3) = RecChangeData!����
                    .TextMatrix(.rows - 1, 4) = IIf(IsNull(RecChangeData!����), "", RecChangeData!����)
                    .TextMatrix(.rows - 1, 5) = IIf(IsNull(RecChangeData!����), "", RecChangeData!����)
                    .TextMatrix(.rows - 1, 6) = RecChangeData!Ʒ��
                    .TextMatrix(.rows - 1, 7) = IIf(IsNull(RecChangeData!���), "", RecChangeData!���)
                    .TextMatrix(.rows - 1, 8) = IIf(IsNull(RecChangeData!����), "", RecChangeData!����)
                    .TextMatrix(.rows - 1, 9) = IIf(IsNull(RecChangeData!����), "", RecChangeData!����)
                    .TextMatrix(.rows - 1, 10) = FormatEx(RecChangeData!ʵ������ * RecChangeData!��, 5) & Right(RecChangeData!����, 1)
                    .TextMatrix(.rows - 1, 11) = FormatEx(RecChangeData!����, 5)
                    .TextMatrix(.rows - 1, 12) = Format(RecChangeData!���, "#####0.00;-#####0.00; ;")
                    .RowData(.rows - 1) = 0
                    
                    '����ҩƷ������ʾ
                    If InStr(";����ҩ;����ҩ;����I��;����II��;", NVL(RecChangeData!�������)) > 0 And NVL(RecChangeData!�������) <> "" Then
                        .Row = .rows - 1: .Col = 6
                        .CellFontBold = True
                    End If
                    
                    .rows = .rows + 1
                End With
            End If
            If Not .EOF Then
                .MoveNext
            Else
                Exit Do
            End If
        Loop
        Call SetMenu(blnEnable)
    End With
    lngRow = IIf(Trim(Bill�ܷ�ҩ�嵥.TextMatrix(Bill�ܷ�ҩ�嵥.rows - 1, 0)) <> "", Bill�ܷ�ҩ�嵥.rows - 1, Bill�ܷ�ҩ�嵥.rows - 2)

    If err <> 0 Then Exit Function
    LoadDataInBill�ܷ��嵥 = True
End Function

Private Function LoadDataInBill�ѷ�ҩ�嵥() As Boolean
    Dim Strִ��״̬ As String, blnEnable As Boolean, lngColor As Long, intCol As Integer
    
    '--����ѷ�ҩ�嵥--
    On Error Resume Next
    err = 0
    blnEnable = False
    LoadDataInBill�ѷ�ҩ�嵥 = False
    
    With Bill�ѷ�ҩ�嵥
        .MousePointer = 11
        .Redraw = False
    End With
    
    '���������н�������
    With RecChangeSendedData
        If .RecordCount <> 0 Then .MoveFirst
'        If Chk�嵥.Value = 0 Then .Sort = GetOrder(str����_����ҩ)
        .Sort = GetOrder(str����_����ҩ)
        Do While Not .EOF
            blnEnable = True
            
            '������ϸ�Ƿ���ת��
            If Val(!ת��) = 0 Then
                Strִ��״̬ = IIf(!ִ��״̬ = 3, "��ҩ", "������")
            Else
                Strִ��״̬ = "������"
            End If
            Bill�ѷ�ҩ�嵥.TextMatrix(Bill�ѷ�ҩ�嵥.rows - 1, ����_�ѷ�ҩ�嵥.����) = !����
            Bill�ѷ�ҩ�嵥.TextMatrix(Bill�ѷ�ҩ�嵥.rows - 1, ����_�ѷ�ҩ�嵥.״̬) = Strִ��״̬
            Bill�ѷ�ҩ�嵥.TextMatrix(Bill�ѷ�ҩ�嵥.rows - 1, ����_�ѷ�ҩ�嵥.����) = !����
            Bill�ѷ�ҩ�嵥.TextMatrix(Bill�ѷ�ҩ�嵥.rows - 1, ����_�ѷ�ҩ�嵥.NO) = !NO
            Bill�ѷ�ҩ�嵥.TextMatrix(Bill�ѷ�ҩ�嵥.rows - 1, ����_�ѷ�ҩ�嵥.����) = IIf(IsNull(!����), "", !����)
            Bill�ѷ�ҩ�嵥.TextMatrix(Bill�ѷ�ҩ�嵥.rows - 1, ����_�ѷ�ҩ�嵥.����) = !����
            Bill�ѷ�ҩ�嵥.TextMatrix(Bill�ѷ�ҩ�嵥.rows - 1, ����_�ѷ�ҩ�嵥.סԺ��) = !סԺ��
            Bill�ѷ�ҩ�嵥.TextMatrix(Bill�ѷ�ҩ�嵥.rows - 1, ����_�ѷ�ҩ�嵥.ҩƷ����) = !Ʒ��
            Bill�ѷ�ҩ�嵥.TextMatrix(Bill�ѷ�ҩ�嵥.rows - 1, ����_�ѷ�ҩ�嵥.������) = IIf(IsNull(!������), "", !������)
            Bill�ѷ�ҩ�嵥.TextMatrix(Bill�ѷ�ҩ�嵥.rows - 1, ����_�ѷ�ҩ�嵥.Ӣ����) = IIf(IsNull(!Ӣ����), "", !Ӣ����)
            Bill�ѷ�ҩ�嵥.TextMatrix(Bill�ѷ�ҩ�嵥.rows - 1, ����_�ѷ�ҩ�嵥.���) = IIf(IsNull(!���), "", !���)
            Bill�ѷ�ҩ�嵥.TextMatrix(Bill�ѷ�ҩ�嵥.rows - 1, ����_�ѷ�ҩ�嵥.����) = IIf(IsNull(!����), "", !����)
            Bill�ѷ�ҩ�嵥.TextMatrix(Bill�ѷ�ҩ�嵥.rows - 1, ����_�ѷ�ҩ�嵥.����) = IIf(IsNull(!����), "", !����)
            Bill�ѷ�ҩ�嵥.TextMatrix(Bill�ѷ�ҩ�嵥.rows - 1, ����_�ѷ�ҩ�嵥.��) = !��
            Bill�ѷ�ҩ�嵥.TextMatrix(Bill�ѷ�ҩ�嵥.rows - 1, ����_�ѷ�ҩ�嵥.����) = !����
            Bill�ѷ�ҩ�嵥.TextMatrix(Bill�ѷ�ҩ�嵥.rows - 1, ����_�ѷ�ҩ�嵥.������) = !������
            Bill�ѷ�ҩ�嵥.TextMatrix(Bill�ѷ�ҩ�嵥.rows - 1, ����_�ѷ�ҩ�嵥.׼����) = !׼����
            Bill�ѷ�ҩ�嵥.TextMatrix(Bill�ѷ�ҩ�嵥.rows - 1, ����_�ѷ�ҩ�嵥.��ҩ��) = ""
            Bill�ѷ�ҩ�嵥.TextMatrix(Bill�ѷ�ҩ�嵥.rows - 1, ����_�ѷ�ҩ�嵥.����) = !����
            Bill�ѷ�ҩ�嵥.TextMatrix(Bill�ѷ�ҩ�嵥.rows - 1, ����_�ѷ�ҩ�嵥.���) = Format(!���, "#####0.00;-#####0.00; ;")
            Bill�ѷ�ҩ�嵥.TextMatrix(Bill�ѷ�ҩ�嵥.rows - 1, ����_�ѷ�ҩ�嵥.����) = !����
            Bill�ѷ�ҩ�嵥.TextMatrix(Bill�ѷ�ҩ�嵥.rows - 1, ����_�ѷ�ҩ�嵥.Ƶ��) = !Ƶ��
            Bill�ѷ�ҩ�嵥.TextMatrix(Bill�ѷ�ҩ�嵥.rows - 1, ����_�ѷ�ҩ�嵥.�÷�) = !�÷�
            Bill�ѷ�ҩ�嵥.TextMatrix(Bill�ѷ�ҩ�嵥.rows - 1, ����_�ѷ�ҩ�嵥.����Ա) = !����Ա
            Bill�ѷ�ҩ�嵥.TextMatrix(Bill�ѷ�ҩ�嵥.rows - 1, ����_�ѷ�ҩ�嵥.��ҩʱ��) = !��ҩʱ��
            Bill�ѷ�ҩ�嵥.TextMatrix(Bill�ѷ�ҩ�嵥.rows - 1, ����_�ѷ�ҩ�嵥.����) = !����
            Bill�ѷ�ҩ�嵥.TextMatrix(Bill�ѷ�ҩ�嵥.rows - 1, ����_�ѷ�ҩ�嵥.ҽ��id) = !ҽ��id
            Bill�ѷ�ҩ�嵥.TextMatrix(Bill�ѷ�ҩ�嵥.rows - 1, ����_�ѷ�ҩ�嵥.��ҩ��) = IIf(IsNull(!��ҩ��), "", !��ҩ��)
            Bill�ѷ�ҩ�嵥.TextMatrix(Bill�ѷ�ҩ�嵥.rows - 1, ����_�ѷ�ҩ�嵥.�ⷿ��λ) = IIf(IsNull(!�ⷿ��λ), "", !�ⷿ��λ)
            Bill�ѷ�ҩ�嵥.TextMatrix(Bill�ѷ�ҩ�嵥.rows - 1, ����_�ѷ�ҩ�嵥.���ID) = !���ID
            Bill�ѷ�ҩ�嵥.TextMatrix(Bill�ѷ�ҩ�嵥.rows - 1, ����_�ѷ�ҩ�嵥.ҩƷID) = !ҩƷID
            
            Bill�ѷ�ҩ�嵥.ColWidth(����_�ѷ�ҩ�嵥.�����) = IIf(Not mblnStarPass, 0, 240)
            Bill�ѷ�ҩ�嵥.ColWidth(����_�ѷ�ҩ�嵥.����) = 0
            Bill�ѷ�ҩ�嵥.ColWidth(����_�ѷ�ҩ�嵥.ҽ��id) = 0
            
            If !����� <> -1 Then
                BlnEnterCell = False
                Bill�ѷ�ҩ�嵥.Row = Bill�ѷ�ҩ�嵥.rows - 1
                Bill�ѷ�ҩ�嵥.Col = 0
                Set Bill�ѷ�ҩ�嵥.CellPicture = imgPass.ListImages(Val(!�����) + 1).Picture
                Bill�ѷ�ҩ�嵥.CellPictureAlignment = 4
                BlnEnterCell = True
            End If
            
            !λ�� = Bill�ѷ�ҩ�嵥.rows - 1
            .Update
            
            '���ݼ�¼״̬�Ĳ�ͬ��������ɫ���ɲ�����1-ԭʼ��¼��2-��ҩ��3-��ҩ��-1-������ת���������������
            lngColor = IIf(!�ɲ��� = 2, glng��ҩ, IIf(!�ɲ��� = 3, glng��ҩ, glng����))
            Bill�ѷ�ҩ�嵥.Row = Bill�ѷ�ҩ�嵥.rows - 1
            For intCol = 0 To Bill�ѷ�ҩ�嵥.Cols - 1
                Bill�ѷ�ҩ�嵥.Col = intCol
                Bill�ѷ�ҩ�嵥.CellForeColor = lngColor
            Next
            Billδ��ҩ�嵥.RowData(Billδ��ҩ�嵥.Row) = lngColor
            
            '����ҩƷ������ʾ
            If InStr(";����ҩ;����ҩ;����I��;����II��;", NVL(!�������)) > 0 And NVL(!�������) <> "" Then
                Bill�ѷ�ҩ�嵥.Col = ����_�ѷ�ҩ�嵥.ҩƷ����
                Bill�ѷ�ҩ�嵥.CellFontBold = True
            End If
            
            Bill�ѷ�ҩ�嵥.rows = Bill�ѷ�ҩ�嵥.rows + 1
            
            Bill�ѷ�ҩ�嵥.ColAlignment(����_�ѷ�ҩ�嵥.ҩƷ����) = 1
            
            
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
        Me.stbThis.Panels(2) = "��ǰ����" & .RecordCount & "���ѷ�ҩƷ��¼"
        Call SetMenu(blnEnable)
    End With
    
    With Bill�ѷ�ҩ�嵥
        .MousePointer = 0
        .Redraw = True
        .Row = 1
        .Col = 1
    End With
    
    Me.MousePointer = 0
    If err <> 0 Then Exit Function
    LoadDataInBill�ѷ�ҩ�嵥 = True
End Function

Private Function CheckStock(Optional ByVal lngҩƷID As Long = 0)
    Dim RecCheckStock As New adodb.Recordset            '������¼��
    Dim dblStock As Double                              '�����
    Dim LngPhysicID As Long                             '��ǰҩƷID
    Dim DblCompare As Double                            '���ڱȽϵ�����
    Dim Strִ��״̬ As String
    Dim lngState As Long, LngLocate As Long
    Dim BlnSet As Boolean, BlnRestore As Boolean, blnEof As Boolean, blnGetData As Boolean
    Dim strSubSql As String
    Dim rsStock As adodb.Recordset
    Dim blnFlag As Boolean
    Dim intCol As Integer
    
    '--���ݿ��״̬��ʾȱҩ�����¼�¼��--
    '�п��ܼ�����¼����ͬһ�ֹ���ҩƷ�������ͳ�Ƴ��ӿ�ʼһֱͳ�Ƶ���ǰλ�õ�������
    '�����ˢ���ü�¼��Ϊ�ա��ж�Ӧ��¼��ִ��״̬��Ϊ��ҩ��ȱҩ���򲻱ؼ���棬ֱ�ӻָ��ϴ��趨��ִ��״̬
    '--Modified by ZYB 20021009
    '--����lngҩƷID���ڣ���ĳ�ʼ�¼��״̬�����ı�ʱ��ֻ�ж�ʹ�ø�ҩƷID�ļ�¼�Ŀ��
    
    On Error GoTo errHandle
    '��λ����
    Select Case strUnit
    Case "�ۼ۵�λ"
        strSubSql = "/1"
    Case "���ﵥλ"
        strSubSql = "/Decode(Nvl(�����װ,0),0,1,�����װ)"
    Case "סԺ��λ"
        strSubSql = "/Decode(Nvl(סԺ��װ,0),0,1,סԺ��װ)"
    Case "ҩ�ⵥλ"
        strSubSql = "/Decode(Nvl(ҩ���װ,0),0,1,ҩ���װ)"
    End Select
    
    Set rsStock = New adodb.Recordset
    With rsStock
        If .State = 1 Then .Close
        .Fields.Append "ҩƷID", adDouble, 18
        .Fields.Append "����", adDouble, 18
        .Fields.Append "���", adDouble, 18
        .Fields.Append "����", adDouble, 18
        .Fields.Append "���", adDouble, 5
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    LngPhysicID = 0
    BlnSet = (RecRefreshCompare.RecordCount <> 0)
    
    With RecChangeData
        If .RecordCount <> 0 Then .MoveFirst
        If .RecordCount = 0 Then Exit Function
        If lngҩƷID <> 0 Then
            .Filter = "ҩƷID=" & lngҩƷID
            If .RecordCount = 0 Then .Filter = 0: Exit Function
        End If
        
        With Billδ��ҩ�嵥
            .Redraw = False
        End With
        
        Do While Not .EOF
            If BlnSet Then  '����Ƿ��ж�Ӧ��¼
                With RecRefreshCompare
                    .MoveFirst
                    .Find "ID=" & RecChangeData!Id
                    BlnRestore = (.EOF Xor True)
                    If BlnRestore Then BlnRestore = (!ִ��״̬ >= 2)
                End With
            End If
            
            If BlnSet And BlnRestore Then
                !ִ��״̬ = RecRefreshCompare!ִ��״̬
                .Update
            Else
                If !ִ��״̬ <= 1 Then  '���ȱҩ�뷢ҩ�ļ�¼
                    blnGetData = False
                    If LngPhysicID <> !ҩƷID Then
                        LngPhysicID = !ҩƷID
                        blnEof = True
                        blnGetData = True
                    End If
                    With rsStock
                        If .RecordCount <> 0 Then
                            .Filter = "ҩƷID=" & LngPhysicID & " And ����=" & IIf(IsNull(RecChangeData!����), 0, RecChangeData!����)
                            blnEof = .EOF
                            blnGetData = blnEof
                            If .RecordCount <> 0 Then LngLocate = !���
                            .Filter = 0
                        End If
                    End With
                    
                    If blnGetData Then
                        If blnEof Then
                            LngLocate = rsStock.RecordCount + 1

                            gstrSQL = " Select nvl(F.�Ƿ���,0) ���,nvl(A.ʵ������,0)" & strSubSql & " ����" & _
                                         " From ҩƷ��� B,�շ���ĿĿ¼ F," & _
                                         "      (Select * From ҩƷ��� " & _
                                         "      Where ����=1 And �ⷿID=[1] And ҩƷID=[2] And nvl(����,0)=[3]) A" & _
                                         " Where B.ҩƷID=F.ID And A.ҩƷID(+)=B.ҩƷID And B.ҩƷID=[2]"
                            Set RecCheckStock = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngҩ��ID, CLng(RecChangeData!ҩƷID), CLng(IIf(IsNull(RecChangeData!����), 0, RecChangeData!����)))
                            
                           With RecCheckStock
                                If .EOF Then
                                    dblStock = 0
                                Else
                                    dblStock = !����
                                End If
                                
                                '������Ӧ�Ŀ���¼
                                With rsStock
                                    .AddNew
                                    !ҩƷID = LngPhysicID
                                    !���� = IIf(IsNull(RecChangeData!����), 0, RecChangeData!����)
                                    !��� = RecCheckStock!���
                                    !���� = dblStock
                                    !��� = LngLocate
                                    .Update
                                End With
                            End With
                        End If
                    End If
                    rsStock.MoveFirst
                    rsStock.Find "���=" & LngLocate
                    dblStock = rsStock!����         'ȡ��ǰ���
                    DblCompare = !ʵ������
                    
                    If dblStock < DblCompare Then
                        '���øü�¼Ϊȱҩ״̬(��״̬�������û��޸�)
                        !ִ��״̬ = IIf(Lngȱҩ��� = 1 Or rsStock!���� <> 0 Or rsStock!��� = 1, 0, !ִ��״̬)
                        .Update
                    ElseIf !ִ��״̬ = 0 Then
                        '����Ƿ�����ҩ
                        lngState = 1
                        If !���շ� = 0 Then lngState = 3
                        If Int����δ��˴�����ҩ = 0 Then
                            If IsNull(!�����) Then
                                lngState = 3
                            Else
                                If Trim(!�����) = "" Then lngState = 3
                            End If
                        End If
                        !ִ��״̬ = lngState                        'ȱʡΪ��ҩ
                        .Update
                    End If
                    
                    '���ִ��״̬Ϊ��ҩ��������
                    If !ִ��״̬ = 1 Then
                        With rsStock
                            !���� = !���� - DblCompare
                            .Update
                        End With
                    End If
                End If
            End If
            
            '���û��ָ��ҩƷ����������ҩƷ��״̬ʱ�������ҵ�ǰҩƷ��ִ��״̬����3ʱ������ݲ����Զ������Ƿ񡰲�����
            If lngҩƷID = 0 And !ִ��״̬ <> 3 Then
                If mstr������� <> "" And !������� <> "" Then
                    If InStr("," & mstr������� & ",", "," & !������� & ",") > 0 Then
                        !ִ��״̬ = 3
                    End If
                End If
                If mstr��ֵ���� <> "" And !��ֵ���� <> "" Then
                    If InStr("," & mstr��ֵ���� & ",", "," & !��ֵ���� & ",") > 0 Then
                        !ִ��״̬ = 3
                    End If
                End If
            End If
            
            Strִ��״̬ = IIf(!ִ��״̬ = 0, "ȱҩ", IIf(!ִ��״̬ = 1, "��ҩ", IIf(!ִ��״̬ = 2, "�ܷ�", "������")))
            !״̬ = Strִ��״̬
            .Update
            
            '����ü�¼����䵽�������������
            With Billδ��ҩ�嵥
                If .rows - 1 >= RecChangeData!λ�� Then
                    .TextMatrix(RecChangeData!λ��, ����_δ��ҩ�嵥.״̬) = Strִ��״̬
                    If Strִ��״̬ = "��ҩ" Then
                        .Row = RecChangeData!λ��
                        For intCol = 0 To .Cols - 1
                            .Col = intCol
                            .CellBackColor = glngSendBlkColor
                        Next
                    End If
                End If
            End With
            .MoveNext
        Loop
        
        With Billδ��ҩ�嵥
            .Redraw = True
        End With
        If lngҩƷID <> 0 Then .Filter = 0
        If .RecordCount <> 0 Then .MoveFirst
    End With
    
    Set rsStock = Nothing
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckDrugStock() As Boolean
    Dim rsTmp As adodb.Recordset
    Dim lngRow As Integer
    Dim lngҩƷID As Long
    
    On Error GoTo errHandle
    CheckDrugStock = True
    With RecChangeData
        If .RecordCount <> 0 Then .MoveFirst
        If .RecordCount = 0 Then Exit Function
        If .EOF Then Exit Function
        
        .Sort = "ҩƷID Asc"
        
        Do While Not .EOF
            If lngҩƷID <> !ҩƷID Then
                If !ִ��״̬ = 1 Then
                    gstrSQL = "Select �շ�ϸĿid From �շ�ִ�п��� Where ִ�п���id = [1] And �շ�ϸĿid = [2]"
                    Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "���ҩƷ�洢�ⷿ", lngҩ��ID, Val(!ҩƷID))
                    
                    If rsTmp.EOF Then
                        MsgBox !Ʒ�� & "δ���ô洢�ⷿ�����ܷ�ҩ��", vbInformation, gstrSysName
                        CheckDrugStock = False
                        Exit Function
                    End If
                    
                    lngҩƷID = !ҩƷID
                Else
                    lngҩƷID = 0
                End If
            End If
            .MoveNext
        Loop
    End With
    
    CheckDrugStock = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function RefreshSendedData() As Boolean
    Dim strCond As String, strSubSql As String
    Dim strName As String
    Dim str���˼����� As String
    Dim strSql�������� As String
    
    On Error GoTo errHandle
    '��Ҫ�������
    If mstrSerchNO = "" And mstr���� = "" Then
'        MsgBox "��ѡ����ҩ���ţ�", vbInformation, gstrSysName
        Call ClearCons
        Exit Function
    End If
    
    mblnFirstSended = False
    
    str���˼����� = IIf(str������ <> "���м�����", " AND A.������=[1] ", "")
    
    '����:bit1=0-����,1-����
    '����ģʽ:0-����,1-���ʵ�,2-���ʱ�
    If Lng����ģʽ = 0 Then
        strCond = " And S.���� IN(9,10)"
    ElseIf Lng����ģʽ = 1 Then
        strCond = " And S.����=9"
    ElseIf Lng����ģʽ = 2 Then
        strCond = " And S.����=10"
    End If
    'ҽ������:0-����,1-����,2-����,3-��ͨ
    '�õ����Ƿ���д�����Ƿ�ҽ��������ҩƷ����
    If Lngҽ������ = 0 Then
    ElseIf Lngҽ������ = 1 Then
        strCond = strCond & " And S.���� Is Not Null And Ltrim(To_Char(Nvl(S.����,0),'00')) Like '0_' And Nvl(C.ҽ�����,0) + 0 >0 "
    ElseIf Lngҽ������ = 2 Then
        strCond = strCond & " And S.���� Is Not Null And Ltrim(To_Char(Nvl(S.����,0),'00')) Like '1_' And Nvl(C.ҽ�����,0) + 0 >0 "
    ElseIf Lngҽ������ = 3 Then
        strCond = strCond & " And (Nvl(C.ҽ�����,0) + 0 =0 Or S.���� Is Null) "
    ElseIf Lngҽ������ = 4 Then
        strCond = strCond & " And S.���� Is Not Null And (Ltrim(To_Char(Nvl(S.����,0),'00')) Like '0_' Or Ltrim(To_Char(Nvl(S.����,0),'00')) Like '1_') And Nvl(C.ҽ�����,0) + 0 > 0 "
    End If
    
    '��λ����
    Select Case strUnit
    Case "�ۼ۵�λ"
        strSubSql = "X.���㵥λ ��λ,1 ��װ,"
    Case "���ﵥλ"
        strSubSql = "D.���ﵥλ ��λ,D.�����װ ��װ,"
    Case "סԺ��λ"
        strSubSql = "D.סԺ��λ ��λ,D.סԺ��װ ��װ,"
    Case "ҩ�ⵥλ"
        strSubSql = "D.ҩ�ⵥλ ��λ,D.ҩ���װ ��װ,"
    End Select
    
    '�õ�ҩƷ���ƴ�
    Call GetDrugFormat
    Select Case intҩƷ����
    Case 0  'ҩƷ����������
        strName = "'['||X.����||']'||" & IIf(mblnTradeName, "NVL(A.����,X.����)", "X.����") & " As Ʒ��,"
    Case 1  'ҩƷ����
        strName = "X.���� As Ʒ��,"
    Case 2  'ҩƷ����
        strName = IIf(mblnTradeName, "NVL(A.����,X.����)", "X.����") & " As Ʒ��,"
    End Select
    
    strName = strName & IIf(Not mblnTradeName, "NVL(A.����,'')", "Decode(A.����,Null,'',X.����)") & " As ������, "
    
    '�������ͣ����˻�Ӥ��
    If mint�������� = 0 Then
        strSql�������� = " And Nvl(C.Ӥ����,0)=0 "
    ElseIf mint�������� = 1 Then
        strSql�������� = " And Nvl(C.Ӥ����,0)>0 "
    End If
    
    
    If Chk�嵥.Value = 0 Then
        '##################������ʾÿ�ʼ�¼�������˶���##################
        gstrSQL = " SELECT DISTINCT S.ID,S.����,S.ҩƷID,S.NO,S.���,S.����,P.���� ����,C.�����־,C.��ʶ��,C.����ID,C.����,C.����," & _
            strName & _
            " NVL(D.ҩ������,0) ����,X.���,T.�������," & _
            strSubSql & _
            " S.���� ��,S.ʵ������ ����,S.��������,S.�ѷ����� ׼����,DECODE(S.����,NULL,'',S.����)||DECODE(S.����,NULL,'',0,'','('||S.����||')') ����,NVL(S.����,0) ����,S.Ч��," & _
            " S.���ۼ� ����,S.���۽�� ���,S.����,S.Ƶ��,S.�÷�,S.ժҪ ˵��,S.�����,TO_CHAR(S.�������,'YYYY-MM-DD HH24:MI:SS') ��ҩʱ��,1 �ɲ���,C.ҽ�����,I.���㵥λ,NVL(S.����,NVL(X.����,'')) ����,nvl(M.�����,-1) �����,nvl(C.ҽ�����,-1) ҽ��id,S.��ҩ��," & IIf(mblnҩƷ���� = True, "L.", "'' ") & "�ⷿ��λ,M.���ID,c.��� �������, Z.���� As Ӣ����,0 As ת�� " & _
            " FROM " & _
            "      (SELECT A.ID,A.NO,A.����,A.���,A.ҩƷID,A.����ID,A.����,A.����,A.Ч��,A.����," & _
            "          NVL(A.����,1) ����,A.ʵ������ ʵ������,NVL(A.����,1)*A.ʵ������-B.�ѷ����� ��������,B.�ѷ�����,A.��¼״̬," & _
            "          A.���ۼ�,A.���۽��,A.����,A.Ƶ��,A.�÷�,A.ժҪ,A.�����,A.�������,A.�Է�����ID,A.�ⷿID,A.����,decode(NVL(A.������,''),'','','(��)'||A.������) ��ҩ�� " & _
            "      FROM ҩƷ�շ���¼ A," & _
            "          (SELECT A.NO,A.����,A.ҩƷID,A.���,SUM(NVL(A.����,1)*A.ʵ������) �ѷ�����" & _
            "          FROM ҩƷ�շ���¼ A" & _
            "          WHERE A.����� IS NOT NULL" & _
            "          AND A.�ⷿID+0=[2] " & _
            "          AND A.������� BETWEEN [8] AND [9] " & str���˼����� & _
            "          GROUP BY A.NO,A.����,A.ҩƷID,A.���) B" & _
            "      WHERE A.NO = B.NO AND A.���� = B.���� AND A.ҩƷID+0 = B.ҩƷID AND A.��� = B.��� AND B.�ѷ�����<>0 And A.����� IS NOT NULL AND (A.��¼״̬=1 OR MOD(A.��¼״̬,3)=0)) S,"
        gstrSQL = gstrSQL & "" & _
            "      ���˷��ü�¼ C,���ű� P,ҩƷ��� D,�շ���ĿĿ¼ X,�շ���Ŀ���� A,ҩƷ���� T,������ĿĿ¼ I,����ҽ����¼ M," & IIf(mblnҩƷ���� = True, "ҩƷ�����޶� L,", "") & "������Ŀ���� Z" & _
            " WHERE S.ҩƷID=D.ҩƷID AND S.�Է�����ID+0=P.ID AND D.ҩ��ID=T.ҩ��ID AND d.ҩƷID=X.ID AND D.ҩ��ID=I.ID and C.ҽ�����=M.ID(+) " & _
            " And D.ҩ��id = Z.������Ŀid(+) And Z.����(+) = 2" & IIf(mblnҩƷ���� = True, " And S.ҩƷID=L.ҩƷID(+) And Nvl(S.�ⷿID,[2])=L.�ⷿID(+) ", "") & _
            " AND D.ҩƷID=A.�շ�ϸĿID(+) AND a.����(+)=3 " & strCond & IIf(mstr���� = "", "", " And C.����=[11] ") & _
            " AND S.����ID=C.ID " & IIf(Val(mlng����ID) = 0, "", " AND C.����ID=[3] ") & IIf(Trim(mstrסԺ��) = "", "", " AND C.��ʶ��=[4] ") & IIf(mstr�������� = "", "", " AND C.���� LIKE [5] ") & _
            " AND (S.��¼״̬=1 OR MOD(S.��¼״̬,3)=0)" & _
            " AND S.����� IS NOT NULL AND S.�ⷿID+0=[2] " & IIf(mstrDrug = "", "", " And Instr([14],',' || T.ҩƷ���� || ',') > 0") & IIf(mstr��ҩ���� = "", "", " And Instr([15],',' || D.��ҩ���� || ',') > 0") & strCond & strSql�������� & _
            IIf(mstr��ʼNO = "", "", " AND S.NO>=[6] ") & IIf(mstr����NO = "", "", " AND S.NO<=[7] ") & " AND Abs(S.ʵ������*S.����)>Abs(S.��������) " & IIf(mstrUse = "", "", " And Instr([13],',' || S.�÷� || ',') > 0") & IIf(mstrSerchNO = "", "", " AND S.NO=[12] ")
        If Trim(mstr����) <> "" Then
            If mint���� = 0 Then
                gstrSQL = gstrSQL & " And Instr([10], ',' || C.��������id || ',') > 0 And C.���˿���id=C.��������id"
            ElseIf mint���� = 1 Then
                gstrSQL = gstrSQL & " And Instr([10], ',' || C.��������id || ',') > 0 And C.���˿���id<>C.��������id"
            Else
                If mstr������ҩ��ʽ = "" Then
                    gstrSQL = gstrSQL & " And Instr([10], ',' || C.���˲���ID || ',') > 0 And C.���˿���id=C.��������id"
                Else
                    gstrSQL = gstrSQL & " And Instr([10], ',' || C.���˲���ID || ',') > 0 "
                    If mstr������ҩ��ʽ <> mstrAllType Then
                        gstrSQL = gstrSQL & " And C.��������id Not In (Select Distinct ����id From ��������˵�� " & _
                            " Where Instr([16],',' || �������� || ',') > 0) "
                    End If
                End If
            End If
        Else
            If mint���� = 0 Then
                gstrSQL = gstrSQL & " And C.���˿���id=C.��������id"
            ElseIf mint���� = 1 Then
                gstrSQL = gstrSQL & " And C.���˿���id<>C.��������id"
            Else
                If mstr������ҩ��ʽ = "" Then
                    gstrSQL = gstrSQL & " And C.���˿���id=C.��������id"
                Else
                    If mstr������ҩ��ʽ <> mstrAllType Then
                        gstrSQL = gstrSQL & " And C.��������id Not In (Select Distinct ����id From ��������˵�� " & _
                            " Where Instr([16],',' || �������� || ',') > 0) "
                    End If
                End If
            End If
        End If
    Else
        '##################�嵥��ʾÿ�ʲ�������##################
        gstrSQL = " SELECT DISTINCT S.ID,S.����,S.ҩƷID,S.NO,S.���,S.����,P.���� ����,C.�����־,C.��ʶ��,C.����ID,C.����,C.����," & strName & _
                 " NVL(D.ҩ������,0) ����,X.���,T.�������," & _
                 strSubSql & _
                 " S.���� ��,S.ʵ������ ����,S.��������,S.�ѷ����� ׼����,DECODE(S.����,NULL,'',S.����)||DECODE(S.����,NULL,'',0,'','('||S.����||')') ����,NVL(S.����,0) ����,S.Ч��," & _
                 " S.���ۼ� ����,S.���۽�� ���,S.����,S.Ƶ��,S.�÷�,S.ժҪ ˵��,TO_CHAR(S.�������,'YYYY-MM-DD HH24:MI:SS') ��ҩʱ��,S.�����,S.�������,�ɲ���,C.ҽ�����,I.���㵥λ,NVL(S.����,NVL(X.����,'')) ����,nvl(M.�����,-1) �����,nvl(C.ҽ�����,-1) ҽ��id,S.��ҩ��," & IIf(mblnҩƷ���� = True, "L.", "'' ") & "�ⷿ��λ, Z.���� As Ӣ����,0 As ת�� " & _
                 " FROM "
        gstrSQL = gstrSQL & _
                 "          (SELECT A.ID,A.NO,A.����,A.���,A.ҩƷID,A.����ID,A.����,A.����,A.Ч��,A.����," & _
                 "              NVL(A.����,1) ����,A.ʵ������,NVL(A.����,1)*A.ʵ������-B.�ѷ����� ��������,B.�ѷ�����,A.��¼״̬," & _
                 "              A.���ۼ� , A.���۽��, A.����, A.Ƶ��, A.�÷�, A.ժҪ, A.�����, A.�������, A.�Է�����ID, A.�ⷿID,1 �ɲ���,A.����," & _
                 "              decode(nvl(A.������,''),'','',Decode(A.��¼״̬,1,'(��)'||A.������," & _
                 "              decode(Mod(A.��¼״̬,3),0,'(��)'||A.������,1,'(��)'||A.������,2,'(��)'||A.������))) ��ҩ�� " & _
                 "          FROM ҩƷ�շ���¼ A," & _
                 "          (SELECT A.NO,A.����,A.ҩƷID,A.���,SUM(NVL(A.����,1)*A.ʵ������) �ѷ�����" & _
                 "          FROM ҩƷ�շ���¼ A" & _
                 "          WHERE A.����� IS NOT NULL" & _
                 "          AND A.�ⷿID+0=[2] " & _
                 "          AND A.������� BETWEEN [8] AND [9] " & str���˼����� & _
                 "          GROUP BY A.NO,A.����,A.ҩƷID,A.���) B" & _
                 "          WHERE A.NO = B.NO AND A.���� = B.���� AND A.ҩƷID+0 = B.ҩƷID AND A.��� = B.��� And A.����� IS NOT NULL AND (A.��¼״̬=1 OR MOD(A.��¼״̬,3)=0)"
        gstrSQL = gstrSQL & _
                 "          UNION" & _
                 "          SELECT A.ID,A.NO,A.����,A.���,A.ҩƷID,A.����ID,A.����,A.����,A.Ч��,A.����," & _
                 "          NVL(A.����,1) ����,A.ʵ������,0 ������,0 �ѷ�����,A.��¼״̬," & _
                 "          A.���ۼ� , A.���۽��, A.����, A.Ƶ��, A.�÷�, A.ժҪ, A.�����, A.�������, A.�Է�����ID, A.�ⷿID," & _
                 "          DECODE(A.��¼״̬,1,1,DECODE(MOD(A.��¼״̬,3),0,1,MOD(A.��¼״̬,3)+1)) �ɲ���,A.����," & _
                 "          decode(nvl(A.������,''),'','',Decode(A.��¼״̬,1,'(��)'||A.������," & _
                 "          decode(Mod(A.��¼״̬,3),0,'(��)'||A.������,1,'(��)'||A.������,2,'(��)'||A.������))) ��ҩ�� " & _
                 "          FROM ҩƷ�շ���¼ A" & _
                 "          WHERE A.����� IS NOT NULL AND NOT (��¼״̬=1 OR MOD(��¼״̬,3)=0)" & _
                 "          AND A.�ⷿID+0=[2] " & _
                 "          AND A.������� BETWEEN [8] AND [9] " & str���˼����� & _
                 "          ) S,"
        gstrSQL = gstrSQL & "" & _
                 "      ���˷��ü�¼ C,���ű� P,ҩƷ��� D,�շ���ĿĿ¼ X,�շ���Ŀ���� A,ҩƷ���� T,������ĿĿ¼ I,����ҽ����¼ M," & IIf(mblnҩƷ���� = True, "ҩƷ�����޶� L,", "") & "������Ŀ���� Z " & _
                 " WHERE S.ҩƷID=D.ҩƷID AND D.ҩ��ID=T.ҩ��ID AND d.ҩƷID=x.ID AND S.�Է�����ID+0=P.ID AND D.ҩ��ID=I.ID and C.ҽ�����=M.ID(+) " & _
                 " And D.ҩ��id = Z.������Ŀid(+) And Z.����(+) = 2 " & IIf(mblnҩƷ���� = True, " And S.ҩƷID=L.ҩƷID(+) And Nvl(S.�ⷿID,[2])=L.�ⷿID(+) ", "") & _
                 " AND D.ҩƷID=A.�շ�ϸĿID(+) AND A.����(+)=3 " & IIf(mstr���� = "", "", " And C.����=[11] ") & _
                 " AND S.����ID=C.ID " & strCond & strSql�������� & _
                 " AND S.����� IS NOT NULL" & IIf(mstrDrug = "", "", " And Instr([14],',' || T.ҩƷ���� || ',') > 0") & IIf(mstr��ҩ���� = "", "", " And Instr([15],',' || D.��ҩ���� || ',') > 0") & _
                 IIf(mstr��ʼNO = "", "", " AND S.NO>=[6] ") & IIf(mstr����NO = "", "", " AND S.NO<=[7] ") & IIf(mstrUse = "", "", " And Instr([13],',' || S.�÷� || ',') > 0") & IIf(mstrSerchNO = "", "", " AND S.NO=[12] ") & _
                 IIf(Val(mlng����ID) = 0, "", " AND C.����ID=[3] ") & IIf(Trim(mstrסԺ��) = "", "", " AND C.��ʶ��=[4] ") & IIf(mstr�������� = "", "", " AND C.���� LIKE [5] ")
        If Trim(mstr����) <> "" Then
            If mint���� = 0 Then
                gstrSQL = gstrSQL & " And Instr([10], ',' || C.��������id || ',') > 0 And C.���˿���id=C.��������id"
            ElseIf mint���� = 1 Then
                gstrSQL = gstrSQL & " And Instr([10], ',' || C.��������id || ',') > 0 And C.���˿���id<>C.��������id"
            Else
                If mstr������ҩ��ʽ = "" Then
                    gstrSQL = gstrSQL & " And Instr([10], ',' || C.���˲���ID || ',') > 0 And C.���˿���id=C.��������id"
                Else
                    gstrSQL = gstrSQL & " And Instr([10], ',' || C.���˲���ID || ',') > 0 "
                    If mstr������ҩ��ʽ <> mstrAllType Then
                        gstrSQL = gstrSQL & " And C.��������id Not In (Select Distinct ����id From ��������˵�� " & _
                            " Where Instr([16],',' || �������� || ',') > 0) "
                    End If
                End If
            End If
        Else
            If mint���� = 0 Then
                gstrSQL = gstrSQL & " And C.���˿���id=C.��������id"
            ElseIf mint���� = 1 Then
                gstrSQL = gstrSQL & " And C.���˿���id<>C.��������id"
            Else
                If mstr������ҩ��ʽ = "" Then
                    gstrSQL = gstrSQL & " And C.���˿���id=C.��������id"
                Else
                    If mstr������ҩ��ʽ <> mstrAllType Then
                        gstrSQL = gstrSQL & " And C.��������id Not In (Select Distinct ����id From ��������˵�� " & _
                            " Where Instr([16],',' || �������� || ',') > 0) "
                    End If
                End If
            End If
        End If
    End If
    
    Dim blnMoved As Boolean
    Dim strSQL As String
    '�ж��Ƿ���ڲ���������ת��
    blnMoved = zldatabase.DateMoved(mstr��ʼ����_�ѷ�)
    If blnMoved Then
        'SQL����¼��Ż��ܣ����κ�һ����ϸҪô���ߣ�Ҫô�󱸣���ˣ���UNION��ʽ����
        strSQL = gstrSQL
        strSQL = Replace(strSQL, "ҩƷ�շ���¼", "HҩƷ�շ���¼")
        strSQL = Replace(strSQL, "���˷��ü�¼", "H���˷��ü�¼")
        strSQL = Replace(strSQL, "0 As ת��", "1 As ת��")
        
        gstrSQL = gstrSQL & " UNION ALL " & strSQL
    End If
    
    If Chk�嵥.Value = 0 Then
        gstrSQL = gstrSQL & " Order By No,����,������� "
    Else
        gstrSQL = gstrSQL & " Order By No,����,�������"
    End If
    
    '--ˢ���ѷ�ҩ�嵥--
'    on error Resume Next
'    err = 0
    
    '��ʼ����¼��
    Call InitRec
    
    RefreshSendedData = False
    
    '�ѷ�������¼
    Set RecBillData = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
        str������, _
        lngҩ��ID, _
        mlng����ID, _
        mstrסԺ��, _
        mstr��������, _
        mstr��ʼNO, _
        mstr����NO, _
        CDate(mstr��ʼ����_�ѷ�), _
        CDate(mstr��������_�ѷ�), _
        "," & mstr���� & ",", _
        mstr����, _
        mstrSerchNO, _
        "," & mstrUse & ",", _
        "," & mstrDrug & ",", _
        "," & mstr��ҩ���� & ",", _
        "," & mstr������ҩ��ʽ & ",")
        
    '�ֹ�Ԥ���
    Call ClearBill(Bill�ѷ�ҩ�嵥)
    If ProduceInsideSendedRecordset = False Then Exit Function
    If LoadDataInBill�ѷ�ҩ�嵥 = False Then
        MsgBox "����ѷ�ҩ�嵥ʱ����������Ԥ֪�Ĵ���", vbInformation, gstrSysName
        Exit Function
    End If
    Call SetGroup(Bill�ѷ�ҩ�嵥, Chk�嵥.Value = 0)
    
    If err <> 0 Then
        MsgBox "��ȡ�ѷ�ҩ����ʱ����������Ԥ֪�Ĵ���", vbInformation, gstrSysName
        Exit Function
    End If
    RefreshSendedData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetPhysicDict(ByVal lngҩƷID As String) As String
    Dim str��λ As String, strϵ�� As String
    Dim rsTemp As New adodb.Recordset
    '--��ȡָ��ҩƷID��Ʒ������񡢵�λ����װ--
    On Error GoTo errHandle
    GetPhysicDict = " ^ ^ ^ "
    gstrSQL = " SELECT A.ҩƷID,A.ҩ��ID,NVL(A.ҩ������,0) ����," & _
              " DECODE(B.���,NULL,B.����,DECODE(B.����,NULL,B.���,B.���||'|'||B.����)) ���," & _
              " A.���ﵥλ,A.�����װ,A.סԺ��λ,A.סԺ��װ,A.ҩ�ⵥλ,ҩ���װ,B.���㵥λ �ۼ۵�λ,1 �ۼ۰�װ " & _
              " FROM ҩƷ��� A,�շ���ĿĿ¼ B" & _
              " WHERE A.ҩƷID=B.ID AND A.ҩƷID=[1]"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[��ȡָ��ҩƷID��Ʒ������񡢵�λ����װ]", lngҩƷID)
    
    If rsTemp.EOF Then Exit Function
    
    Select Case strUnit
    Case "�ۼ۵�λ"
        str��λ = rsTemp!�ۼ۵�λ
        strϵ�� = rsTemp!�ۼ۰�װ
    Case "���ﵥλ"
        str��λ = rsTemp!���ﵥλ
        strϵ�� = rsTemp!�����װ
    Case "סԺ��λ"
        str��λ = rsTemp!סԺ��λ
        strϵ�� = rsTemp!סԺ��װ
    Case "ҩ�ⵥλ"
        str��λ = rsTemp!ҩ�ⵥλ
        strϵ�� = rsTemp!ҩ���װ
    End Select
    
    GetPhysicDict = "С��"
    GetPhysicDict = GetPhysicDict & "^" & IIf(IsNull(rsTemp!���), " ", rsTemp!���) & "^" & _
    str��λ & "^" & strϵ�� & "^" & rsTemp!����
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub SetMenuAndToolbarState()
    Dim LngCurLocate As Long                    '��ǰλ��
    '--���ò˵������߰�ť��״̬--
    With RecChangeData
        If .RecordCount <> 0 Then .MoveFirst
        .Find "ִ��״̬=0"                      'ȱҩ
        MnuEditDesire.Enabled = (.EOF Xor True)
        Tbar.Buttons("Desire").Enabled = (.EOF Xor True)
        If .RecordCount <> 0 Then .MoveFirst
        
        .Find "ִ��״̬=1"                      '��ҩ
        MnuEditVerify.Enabled = (.EOF Xor True)
        Tbar.Buttons("Consignment").Enabled = (.EOF Xor True)
        If .RecordCount <> 0 Then .MoveFirst
        
        .Find "ִ��״̬=2"                      '�ܷ�
        mnuEditHandback.Enabled = (.EOF Xor True)
        Tbar.Buttons("Handback").Enabled = (.EOF Xor True)
        If .RecordCount <> 0 Then .MoveFirst
    End With
    With RecChangeSendedData
        If .RecordCount <> 0 Then .MoveFirst
        .Find "ִ��״̬=3"                      '��ҩ
        MnuEditRestore.Enabled = (.EOF Xor True)
        Tbar.Buttons("Restore").Enabled = (.EOF Xor True)
    End With
    
    If mnuEditHandback.Enabled = False Then
        With Bill�ܷ�ҩ�嵥
            For LngCurLocate = 1 To .rows - 1
                If Trim(.TextMatrix(LngCurLocate, 1)) = "�ָ�" Then
                    mnuEditHandback.Enabled = True
                    Tbar.Buttons("Handback").Enabled = True
                    Exit For
                End If
            Next
        End With
    End If
End Sub

Private Function AviShow(Optional ByVal BlnShow As Boolean = True)
    '����Flash����
    DoEvents
    
    If BlnShow Then
        zlCommFun.ShowFlash "���ڲ�������,���Ժ�...", Me
    Else
        zlCommFun.StopFlash
    End If
    
    DoEvents
End Function

Private Function CheckBill(ByVal IntOper As Integer, ByVal LngID As Long) As Integer
    Dim RecCheck As New adodb.Recordset
    
    '--���ݽ�Ҫִ�еĲ������ж��Ƿ�����--
    '0-�ܷ�;1-��ҩ;2-��ҩ
    '����:
    '0-�������
    '1-�ѷ�ҩ
    '2-��ɾ��
    '3-δ��ҩ
    On Error GoTo errHandle
    gstrSQL = " Select A.NO,Nvl(B.��¼״̬,0) AS ��˱�־,A.�����,Decode(Nvl(A.ժҪ,'С��'),'�ܷ�',3,B.ִ��״̬) ִ��״̬ From ҩƷ�շ���¼ A,���˷��ü�¼ B " & _
             " Where A.����ID=B.ID And A.ID=[1] "
    If IntOper = 2 Then
        gstrSQL = gstrSQL & " And ����� IS Not Null"
    Else
        gstrSQL = gstrSQL & " And ����� IS Null"
    End If
    Set RecCheck = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, LngID)
    
    With RecCheck
        If .EOF Then CheckBill = 2: MsgBox "δ�ҵ�ָ������,�����Ѿ�����������Ա����,����������ֹ��", vbInformation, gstrSysName: Exit Function
        If Not IsNull(!�����) Then
            If IntOper <> 2 Then CheckBill = 1: MsgBox "�ô���[" & !NO & "]�ѱ���������Ա��ҩ������������ֹ��", vbInformation, gstrSysName: Exit Function
        Else
            If IntOper = 2 Then CheckBill = 3: MsgBox "�ô���[" & !NO & "]��δ��ҩ������������ֹ��", vbInformation, gstrSysName: Exit Function
        End If
        If IntOper = 1 Then
            If !ִ��״̬ = 3 Then CheckBill = 2: MsgBox "�ô���[" & !NO & "]�Ѿܷ�������������ֹ��", vbInformation, gstrSysName: Exit Function
            If !��˱�־ = 0 And Int����δ��˴�����ҩ = 0 Then
                CheckBill = 4: MsgBox "�ô���[" & !NO & "]��δ��ˣ�����������ֹ��", vbInformation, gstrSysName
            End If
        End If
    End With
    
    CheckBill = 0
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog

End Function

Private Function subPrint(ByVal bytMode As Byte)
    '--��ӡ--
    Dim ObjThis As Object
    Dim objPrint As New zlPrint1Grd
    Dim ObjAppRow As New zlTabAppRow
    Dim intCol As Integer
    
    Select Case TabShow.Tab
    Case 0
        Set ObjThis = Billδ��ҩ�嵥
    Case 1
        Set ObjThis = Bill���ܷ�ҩ
    Case 2
        Set ObjThis = Billȱҩ�嵥
    Case 3
        Set ObjThis = Bill�ܷ�ҩ�嵥
    Case 4
        Set ObjThis = Bill�ѷ�ҩ�嵥
    End Select
    
    '�ָ�����ǰ��ɫ
    With ObjThis
        .Redraw = False
        For intCol = 0 To .Cols - 1
            .Col = intCol
            .CellForeColor = &H80000008
        Next
        .Col = 0
        .Redraw = True
    End With
    
    Set ObjAppRow = New zlTabAppRow
    ObjAppRow.Add "��ӡ��:" & gstrUserName
    ObjAppRow.Add "��ӡ����:" & Format(zldatabase.Currentdate, "yyyy-MM-dd")
    objPrint.BelowAppRows.Add ObjAppRow
    
    Set ObjAppRow = New zlTabAppRow
    ObjAppRow.Add "��ʼʱ��:" & Format(IIf(TabShow.Tab = 4, mstr��ʼ����_�ѷ�, mstr��ʼ����_δ��), "yyyy-MM-dd HH:mm:ss")
    ObjAppRow.Add "����ʱ��:" & Format(IIf(TabShow.Tab = 4, mstr��������_�ѷ�, mstr��������_δ��), "yyyy-MM-dd HH:mm:ss")
    objPrint.UnderAppRows.Add ObjAppRow
    
    objPrint.Title.Text = TabShow.TabCaption(TabShow.Tab)
    Set objPrint.Body = ObjThis
    
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
    
    '�ָ�ѡ��״̬������ǰ��ɫ
    Call SetSelectColor(ObjThis)
End Function

Private Sub Tbar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu MnuViewTool, 2
End Sub

Private Sub Ȩ�޿���()
    '��������
    '��ҩ
    '�ܷ�
    '��ҩ

    If Not IsHavePrivs(mstrPrivs, "��ҩ") Then
        MnuEditVerify.Visible = False
        Tbar.Buttons("Consignment").Visible = False
    End If
    If Not IsHavePrivs(mstrPrivs, "�ܷ�") Then
        mnuEditHandback.Visible = False
        Tbar.Buttons("Handback").Visible = False
    End If
    If Not IsHavePrivs(mstrPrivs, "��ҩ") Then
        If MnuEditVerify.Visible = False And mnuEditHandback.Visible = False Then
            mnuEdit.Visible = False
            Tbar.Buttons("Edit1").Visible = False
        Else
            MnuEditRestore.Visible = False
        End If
        Tbar.Buttons("Restore").Visible = False
    End If
    mnuFilePrintTotal.Visible = IsHavePrivs(mstrPrivs, "���ܴ�ӡ")
    mnuFileRestore.Visible = IsHavePrivs(mstrPrivs, "��ӡ������ҩ��ϸ")
    If Not mnuFileRestore.Visible Then MnuFile2.Visible = mnuFilePrintTotal.Visible
    If Not IsHavePrivs(mstrPrivs, "������ҩ���Ĵ���") Then
        mnuEditHandbackBatch.Visible = False
    End If
    If gblnPass And IsHavePrivs(mstrPrivs, "������ҩ���") Then
        mblnStarPass = True
    End If
    If Not IsHavePrivs(mstrPrivs, "��ҩ����") Then
        mnuReVerify.Visible = False
        Tbar.Buttons("ReVerify").Visible = False
    End If
End Sub

Private Sub ClearBill(ByVal MsfObj As MSHFlexGrid)
    '����ؼ�����
    Dim i As Long, j As Long
    
    MsfObj.Redraw = False
    For i = 1 To MsfObj.rows - 1
        For j = 0 To MsfObj.Cols - 1
            MsfObj.TextMatrix(i, j) = ""
        Next
    Next
    
    MsfObj.rows = 2
    MsfObj.Row = 1: MsfObj.Col = 0
    MsfObj.Redraw = True
End Sub

Private Sub SetSelectColor(ByVal MsfObj As MSHFlexGrid)
    Dim LngSelectRow As Long, intCol As Integer, lngColor As Long
    Dim strCompare As String
    
    On Error Resume Next
    
    With MsfObj
        '����������λ
        CurCell.Col = .Col
        CurCell.Row = .Row
        CurCell.CellHeight = .CellHeight
        CurCell.CellLeft = .CellLeft
        CurCell.CellTop = .CellTop - 30
        CurCell.CellWidth = .CellWidth
        
        .Redraw = False
        LngSelectRow = .Row         '���浱ǰѡ����
        If Val(.Tag) <> 0 Then
            .Row = Val(.Tag)        '����ϴ�ѡ����
            strCompare = IIf(.TextMatrix(.Row, 0) = "", "С��", .TextMatrix(.Row, 0))
            
            Select Case .Name
            Case "Bill�ѷ�ҩ�嵥"
                With RecChangeSendedData
                    If .RecordCount <> 0 Then
                        .MoveFirst
                        .Find "λ��=" & Val(MsfObj.Tag)
                    End If
                    If .EOF Then
                        lngColor = &H80000008
                    Else
                        lngColor = IIf(!�ɲ��� = 1, glng����, IIf(!�ɲ��� = 2, glng��ҩ, glng��ҩ))
                    End If
                End With
            Case "Billδ��ҩ�嵥"
                With RecChangeData
                    If .RecordCount <> 0 Then
                        .MoveFirst
                        .Find "λ��=" & Val(MsfObj.Tag)
                    End If
                    If .EOF Then
                        lngColor = glngOtherBlkColor
                    Else
                        lngColor = IIf(!ִ��״̬ = 1, glngSendBlkColor, glngOtherBlkColor)
                    End If
                End With
'                lngColor = IIf(InStr(1, "�ϼ�,С��", strCompare) <> 0, glng��ҩ, glng����)
            Case "Bill���ܷ�ҩ"
                lngColor = IIf(InStr(1, "�ϼ�,С��", strCompare) <> 0, glng��ҩ, glng����)
            Case Else
                lngColor = glng����
            End Select
            
            For intCol = 0 To .Cols - 1
                .Col = intCol
                .CellBackColor = IIf(.Name = "Billδ��ҩ�嵥", lngColor, glngOtherBlkColor)
            Next
            .Col = 0
        End If
        
        .Tag = LngSelectRow
        .Row = .Tag                 '���õ�ǰѡ����
        strCompare = IIf(.TextMatrix(.Row, 0) = "", "С��", .TextMatrix(.Row, 0))
        
        Select Case .Name
        Case "Bill�ѷ�ҩ�嵥"
            With RecChangeSendedData
                If .RecordCount <> 0 Then
                    .MoveFirst
                    .Find "λ��=" & LngSelectRow
                End If
                If .EOF Then
                    lngColor = &H8000000D
                Else
                    lngColor = IIf(!�ɲ��� = 1, glng����, IIf(!�ɲ��� = 2, glng��ҩ, glng��ҩ))
                End If
            End With
        Case "Billδ��ҩ�嵥"
            lngColor = IIf(InStr(1, "�ϼ�,С��", strCompare) <> 0, glng��ҩ, glng����)
        Case "Bill���ܷ�ҩ"
            lngColor = IIf(InStr(1, "�ϼ�,С��", strCompare) <> 0, glng��ҩ, glng����)
        Case Else
            lngColor = glng����
        End Select
        
        For intCol = 0 To .Cols - 1
            .Col = intCol
            .CellBackColor = &HC0C0C0
        Next
        .Col = 0
        .Redraw = True
    End With
End Sub

Private Function SetMenuCheck(ByVal MenuObj As Menu) As Menu
    Dim MenuCheck As Menu, strState As String
    '���ö�Ӧ�˵���ѡ��״̬,������
    
    Select Case MenuObj.Name
    Case "PopMenu_1"
        Consignment.Checked = False
        Lack.Checked = False
        HandBack.Checked = False
        Nop_1.Checked = False
        
        strState = Billδ��ҩ�嵥.TextMatrix(Billδ��ҩ�嵥.Row, ����_δ��ҩ�嵥.״̬)
        Select Case strState
        Case "��ҩ"
            Set MenuCheck = Consignment
        Case "ȱҩ"
            Set MenuCheck = Lack
        Case "�ܷ�"
            Set MenuCheck = HandBack
        Case "������"
            Set MenuCheck = Nop_1
        End Select
        MenuCheck.Checked = True
    Case "PopMenu_2"
        Restore.Checked = False
        Nop_2.Checked = False
        
        strState = Bill�ѷ�ҩ�嵥.TextMatrix(Bill�ѷ�ҩ�嵥.Row, ����_�ѷ�ҩ�嵥.״̬)
        Select Case strState
        Case "��ҩ"
            Set MenuCheck = Restore
        Case "������"
            Set MenuCheck = Nop_2
        End Select
        MenuCheck.Checked = True
    Case "PopMenu_3"
        ResumeDo.Checked = False
        Nop_3.Checked = False
        
        strState = Bill�ܷ�ҩ�嵥.TextMatrix(Bill�ܷ�ҩ�嵥.Row, 1)
        Select Case strState
        Case "�ָ�"
            Set MenuCheck = ResumeDo
        Case "������"
            Set MenuCheck = Nop_3
        End Select
        MenuCheck.Checked = True
    End Select
    Set SetMenuCheck = MenuCheck
End Function

Private Sub UpdateRsByMenu(ByVal MenuObj As Menu, Optional ByVal IntStyle As Integer = 1)
    Dim lngFind As Long
    '1:δ��ҩ
    '3:�ܷ�ҩ
    '2:�ѷ�ҩ
    
    '--�����ڲ���¼��--
    Select Case IntStyle
    Case 1
        With Billδ��ҩ�嵥
            lngFind = .Row
        End With
        With RecChangeData
            If .RecordCount <> 0 Then .MoveFirst
            .Find "λ��=" & lngFind
            If .EOF Then
                MsgBox "δ�ҵ��ü�¼��", vbInformation, gstrSysName
                Exit Sub
            End If

            Select Case MenuObj.Name
            Case "Consignment"
                lngFind = 1
            Case "HandBack"
                lngFind = 2
            Case Else
                lngFind = 3
            End Select
            !ִ��״̬ = lngFind
            .Update
        End With
    
        '������ؼ�¼��ִ��״̬
        Call CheckStock(RecChangeData!ҩƷID)
    Case 2
        With Bill�ѷ�ҩ�嵥
            lngFind = .Row
        End With
        With RecChangeSendedData
            If .RecordCount <> 0 Then .MoveFirst
            .Find "λ��=" & lngFind
            If .EOF Then
                MsgBox "δ�ҵ��ü�¼��", vbInformation, gstrSysName
                Exit Sub
            End If

            Select Case MenuObj.Name
            Case "Restore"
                lngFind = 3
                Bill�ѷ�ҩ�嵥.TextMatrix(Bill�ѷ�ҩ�嵥.Row, ����_�ѷ�ҩ�嵥.״̬) = "��ҩ"
                Bill�ѷ�ҩ�嵥.TextMatrix(Bill�ѷ�ҩ�嵥.Row, ����_�ѷ�ҩ�嵥.��ҩ��) = Val(Bill�ѷ�ҩ�嵥.TextMatrix(Bill�ѷ�ҩ�嵥.Row, ����_�ѷ�ҩ�嵥.׼����))
            Case Else
                lngFind = 1
                Bill�ѷ�ҩ�嵥.TextMatrix(Bill�ѷ�ҩ�嵥.Row, ����_�ѷ�ҩ�嵥.״̬) = "������"
                Bill�ѷ�ҩ�嵥.TextMatrix(Bill�ѷ�ҩ�嵥.Row, ����_�ѷ�ҩ�嵥.��ҩ��) = ""
            End Select
            !ִ��״̬ = lngFind
            .Update
        End With
        Call Bill�ѷ�ҩ�嵥_EnterCell
    Case 3
        Select Case MenuObj.Name
        Case "ResumeDo"
            Bill�ܷ�ҩ�嵥.TextMatrix(Bill�ܷ�ҩ�嵥.Row, 1) = "�ָ�"
        Case Else
            Bill�ܷ�ҩ�嵥.TextMatrix(Bill�ܷ�ҩ�嵥.Row, 1) = "������"
        End Select
    End Select
    
    '���ò˵������߰�ť��״̬
    Call SetMenuAndToolbarState
End Sub

Private Function Load�ܷ�()
    Dim ArrayPhysic As Variant
    Dim rsRefuse As New adodb.Recordset
    Dim strCond As String, strSubSql As String
    
    '��װ����ʵʵ���ھܷ��Ĵ����嵥

    '����:bit1=0-����,1-������bit2:3-��Ժ��ҩ
    '����ģʽ:0-����,1-���ʵ�,2-���ʱ�
    On Error GoTo errHandle
    If Lng����ģʽ = 0 Then
        strCond = " And S.���� IN(9,10)"
    ElseIf Lng����ģʽ = 1 Then
        strCond = " And S.����=9"
    ElseIf Lng����ģʽ = 2 Then
        strCond = " And S.����=10"
    End If
    'ҽ������:0-����,1-����,2-����,3-��ͨ
    '�õ����Ƿ���д�����Ƿ�ҽ��������ҩƷ����
    '��ҽ��������ж��Ƿ�Ϊҽ�� by lyq 2005-05-18
    Dim strҽ����� As String
    If Lngҽ������ = 0 Then
    ElseIf Lngҽ������ = 1 Then
        strCond = strCond & " And S.���� Is Not Null And Ltrim(To_Char(Nvl(S.����,0),'00')) Like '0_' "
        strҽ����� = " And Nvl(C.ҽ�����,0) + 0 > 0 "
    ElseIf Lngҽ������ = 2 Then
        strCond = strCond & " And S.���� Is Not Null And Ltrim(To_Char(Nvl(S.����,0),'00')) Like '1_' "
        strҽ����� = " And Nvl(C.ҽ�����,0) + 0 > 0 "
    ElseIf Lngҽ������ = 3 Then
        strCond = strCond
        strҽ����� = " And (Nvl(C.ҽ�����,0) + 0 = 0 Or S.���� Is Null) "
    ElseIf Lngҽ������ = 4 Then
        strCond = strCond & " And S.���� Is Not Null And (Ltrim(To_Char(Nvl(S.����,0),'00')) Like '0_' Or Ltrim(To_Char(Nvl(S.����,0),'00')) Like '1_') "
        strҽ����� = " And Nvl(C.ҽ�����,0) + 0 > 0 "
    End If
    '��Ժ��ҩ:'0-����,1-������Ժ��ҩ,2-������Ժ��ҩ,3-������ȡҩ,4-������ȡҩ,5-Ժ����ҩ(��������Ժ��ҩ����ȡҩ),6-��Ժ��ҩ����ȡҩ
    If int��Ժ��ҩ = 0 Then
    ElseIf int��Ժ��ҩ = 1 Then
        strCond = strCond & " And Not Ltrim(To_Char(Nvl(S.����,0),'00')) Like '_3'"
    ElseIf int��Ժ��ҩ = 2 Then
        strCond = strCond & " And Ltrim(To_Char(Nvl(S.����,0),'00')) Like '_3'"
    ElseIf int��Ժ��ҩ = 3 Then
        strCond = strCond & " And Not Ltrim(To_Char(Nvl(S.����,0),'00')) Like '_4'"
    ElseIf int��Ժ��ҩ = 4 Then
        strCond = strCond & " And Ltrim(To_Char(Nvl(S.����,0),'00')) Like '_4'"
    ElseIf int��Ժ��ҩ = 5 Then
        strCond = strCond & " And Not Ltrim(To_Char(Nvl(S.����,0),'00')) Like '_3' And Not Ltrim(To_Char(Nvl(S.����,0),'00')) Like '_4'"
    ElseIf int��Ժ��ҩ = 6 Then
        strCond = strCond & " And (Ltrim(To_Char(Nvl(S.����,0),'00')) Like '_3' Or Ltrim(To_Char(Nvl(S.����,0),'00')) Like '_4')"
    End If
    
    '��λ����
    Select Case strUnit
    Case "�ۼ۵�λ"
        strSubSql = "X.���㵥λ ��λ,1 ��װ,"
    Case "���ﵥλ"
        strSubSql = "D.���ﵥλ ��λ,D.�����װ ��װ,"
    Case "סԺ��λ"
        strSubSql = "D.סԺ��λ ��λ,D.סԺ��װ ��װ,"
    Case "ҩ�ⵥλ"
        strSubSql = "D.ҩ�ⵥλ ��λ,D.ҩ���װ ��װ,"
    End Select

    gstrSQL = " SELECT DISTINCT S.ID,S.ҩƷID,P.���� ����,S.��ҩ��,C.����Ա���� �����,S.����,NVL(S.����,0) ����," & _
             " S.NO,C.����,C.����,C.�����־,'['||X.����||']'||" & IIf(mblnTradeName, "NVL(A.����,X.����)", "X.����") & " Ʒ��,S.���� ��,S.ʵ������ ����," & _
             " NVL(D.ҩ������,0) ����,X.���," & _
             strSubSql & _
             " DECODE(S.����,NULL,'',S.����) ����,NVL(S.����,0) ����,T.�������," & _
             " S.���ۼ� ����,S.���۽�� ���,S.����,S.Ƶ��,S.�÷�,S.ժҪ ˵��,C.ҽ�����" & _
             " FROM " & _
             "      (SELECT * FROM ҩƷ�շ���¼ S " & _
             "      WHERE MOD(��¼״̬,3)=1 AND NVL(LTRIM(RTRIM(ժҪ)),'С��')='�ܷ�' " & _
             "      AND ����� IS NULL" & _
             "      AND (�ⷿID+0=[1] OR �ⷿID IS NULL) AND �������� BETWEEN [2] AND [3] " & strCond & IIf(mstrUse = "", "", " And Instr([6],',' || S.�÷� || ',') > 0")
    gstrSQL = gstrSQL & ") S,���˷��ü�¼ C,���ű� P,ҩƷ��� D,�շ���ĿĿ¼ X,�շ���Ŀ���� A,ҩƷ���� T " & _
             " WHERE S.ҩƷID=D.ҩƷID AND D.ҩƷID=X.ID and D.ҩ��ID=T.ҩ��ID" & _
             " AND D.ҩƷID=A.�շ�ϸĿID(+) AND A.����(+)=3 " & IIf(mstrDrug = "", "", " And Instr([7],',' || T.ҩƷ���� || ',') > 0") & IIf(mstr��ҩ���� = "", "", " And Instr([8],',' || D.��ҩ���� || ',') > 0") & _
             " AND S.�Է�����ID=P.ID AND S.NO=C.NO AND S.����ID=C.ID " & strҽ����� & IIf(mstr���� = "", "", " And C.����=[5] ")
             
    Select Case mint��Χ
    Case 1
        gstrSQL = gstrSQL & " And S.ʵ������>=0"
    Case 2
        gstrSQL = gstrSQL & " And S.ʵ������<0"
    End Select
    
    If Trim(mstr����) <> "" Then
        If mint���� = 0 Then
            gstrSQL = gstrSQL & " And Instr([4], ',' || C.��������id || ',') > 0 And C.���˿���id=C.��������id"
        ElseIf mint���� = 1 Then
            gstrSQL = gstrSQL & " And Instr([4], ',' || C.��������id || ',') > 0 And C.���˿���id<>C.��������id"
        Else
            If mstr������ҩ��ʽ = "" Then
                gstrSQL = gstrSQL & " And Instr([4], ',' || C.���˲���ID || ',') > 0 And C.���˿���id=C.��������id"
            Else
                gstrSQL = gstrSQL & " And Instr([4], ',' || C.���˲���ID || ',') > 0 "
                If mstr������ҩ��ʽ <> mstrAllType Then
                    gstrSQL = gstrSQL & " And C.��������id Not In (Select Distinct ����id From ��������˵�� " & _
                        " Where Instr([9],',' || �������� || ',') > 0) "
                End If
            End If
        End If
    Else
        If mint���� = 0 Then
            gstrSQL = gstrSQL & " And C.���˿���id=C.��������id"
        ElseIf mint���� = 1 Then
            gstrSQL = gstrSQL & " And C.���˿���id<>C.��������id"
        Else
            If mstr������ҩ��ʽ = "" Then
                gstrSQL = gstrSQL & " And C.���˿���id=C.��������id"
            Else
                If mstr������ҩ��ʽ <> mstrAllType Then
                    gstrSQL = gstrSQL & " And C.��������id Not In (Select Distinct ����id From ��������˵�� " & _
                        " Where Instr([9],',' || �������� || ',') > 0) "
                End If
            End If
        End If
    End If
    gstrSQL = gstrSQL & " Order By S.No,S.����"
    
    '--���ܷ��嵥--
'    On Error Resume Next
'    err = 0
    
    Set rsRefuse = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
        lngҩ��ID, _
        CDate(mstr��ʼ����_δ��), _
        CDate(mstr��������_δ��), _
        "," & mstr���� & ",", _
        mstr����, _
        "," & mstrUse & ",", _
        "," & mstrDrug & ",", _
        "," & mstr��ҩ���� & ",", _
        "," & mstr������ҩ��ʽ & ",")
    
    With rsRefuse
        Do While Not .EOF
            With Bill�ܷ�ҩ�嵥
                If Trim(.TextMatrix(.rows - 1, 0)) <> "" Then .rows = .rows + 1
                .TextMatrix(.rows - 1, 0) = rsRefuse!����
                .TextMatrix(.rows - 1, 1) = "������"
                .TextMatrix(.rows - 1, 2) = rsRefuse!NO
                .TextMatrix(.rows - 1, 3) = IIf(NVL(rsRefuse!����, 0) = 0, IIf(rsRefuse!�����־ = 1 Or rsRefuse!�����־ = 4, "������ʵ�", IIf(rsRefuse!���� = 9, "סԺ���ʵ�", "סԺ���ʱ�")), IIf(IsNull(rsRefuse!����) = True, "סԺ���ʵ�", IIf(rsRefuse!���� Like "0*", "����", IIf(rsRefuse!���� Like "1*", "����", "���ʱ�"))))
                .TextMatrix(.rows - 1, 4) = IIf(IsNull(rsRefuse!����), "", rsRefuse!����)
                .TextMatrix(.rows - 1, 5) = IIf(IsNull(rsRefuse!����), "", rsRefuse!����)
                '��ȡ��ҩƷ�������Ϣ
                .TextMatrix(.rows - 1, 6) = rsRefuse!Ʒ��
                .TextMatrix(.rows - 1, 7) = IIf(IsNull(rsRefuse!���), "", rsRefuse!���)
                .TextMatrix(.rows - 1, 8) = IIf(IsNull(rsRefuse!����), "", rsRefuse!����)
                .TextMatrix(.rows - 1, 9) = FormatEx(rsRefuse!���� * rsRefuse!�� / rsRefuse!��װ, 5) & rsRefuse!��λ
                .TextMatrix(.rows - 1, 10) = FormatEx(rsRefuse!���� * rsRefuse!��װ, 5)
                .TextMatrix(.rows - 1, 11) = Format(rsRefuse!���, "#####0.00;-#####0.00; ;")
                .RowData(.rows - 1) = rsRefuse!Id
            End With
            If Not .EOF Then
                .MoveNext
            Else
                Exit Do
            End If
        Loop
        .Close
    End With
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetMenu(Optional ByVal blnEnable As Boolean = False)
    MnuFilePrint.Enabled = blnEnable
    MnuFilePreview.Enabled = blnEnable
    MnuFileExcel.Enabled = blnEnable
    Tbar.Buttons("Preview").Enabled = blnEnable
    Tbar.Buttons("Print").Enabled = blnEnable
End Sub

Private Sub LocateCboItemData(ByVal cboObj As ComboBox, ByVal lngItem As Long)
    Dim LngLocate As Long
    With cboObj
        If .ListCount = 0 Then Exit Sub
        For LngLocate = 0 To .ListCount - 1
            If .ItemData(LngLocate) = lngItem Then
                .ListIndex = LngLocate
                Exit Sub
            End If
        Next
        .ListIndex = 0
    End With
End Sub

Private Sub ShowCbo()
    On Error Resume Next
    
    With Cbo����
        If .ListCount = 0 Then Exit Sub
        .Left = Billδ��ҩ�嵥.Left + TabShow.Left + CurCell.CellLeft
        .Top = Billδ��ҩ�嵥.Top + TabShow.Top + CurCell.CellTop
        .Width = CurCell.CellWidth
        .Visible = True
        .ZOrder 0
    End With
End Sub

Private Sub ShowTxt(Optional ByVal ���뷽ʽ As Integer = 1)
    '0-�����;1-�Ҷ���;2-���ж���
    On Error Resume Next
    With TxtInput
        .Alignment = ���뷽ʽ
        .Left = Bill�ѷ�ҩ�嵥.Left + TabShow.Left + CurCell.CellLeft
        .Top = Bill�ѷ�ҩ�嵥.Top + TabShow.Top + CurCell.CellTop + 20
        .Width = CurCell.CellWidth - 20
        .Visible = True
        .ZOrder 0
        .SetFocus
    End With
    Call SelAll(TxtInput)
End Sub

Private Sub tbsType_Click()
    If mintLastDeptType <> tbsType.SelectedItem.Index - 1 Then
        txt����.Tag = ""
        txt����.Text = ""
        mintLastDeptType = tbsType.SelectedItem.Index - 1
    End If
    
End Sub

Private Sub TimerAuto_Timer()
    '�Զ�ˢ��ֻ���δ��ҩƷ�嵥
    Dim dateCurr As Date
        
    '���������С��ʱ�˳�
    If Me.WindowState = 1 Then Exit Sub
    
    '�������ڲ��ǵ�ǰ����ʱ�˳�
    If mlngMyWindow = 0 Then
        mlngMyWindow = GetActiveWindow()
    Else
        If mlngMyWindow <> GetActiveWindow() Then Exit Sub
    End If
    
    '�������δ��ҩ��������Զ�ˢ�²���Ϊ0ʱ�˳�
    If TabShow.Tab <> 0 Or mint�Զ�ˢ��δ��ҩ�嵥 = 0 Then Exit Sub
    
    '���ݵ�ǰʱ�����ϴ�ˢ��ʱ�����������Ƿ�ˢ��
    dateCurr = zldatabase.Currentdate
    If DateDiff("s", mdate�ϴ�ˢ��ʱ��, dateCurr) < mint�Զ�ˢ��δ��ҩ�嵥 * 60 Then Exit Sub
    
    TimerAuto.Enabled = False
    DoEvents
    Call mnuViewRefresh_Click

'    MsgBox "Ok��" & "[" & Format(dateCurr, "yyyy-mm-dd hh:mm:ss") & "]" & "[" & Format(mdate�ϴ�ˢ��ʱ��, "yyyy-mm-dd hh:mm:ss") & "]"
'    mdate�ϴ�ˢ��ʱ�� = zldatabase.Currentdate
    
    DoEvents
    TimerAuto.Enabled = True
End Sub
Private Sub TxtInput_LostFocus()
    Dim blnUnValid As Boolean, dblCount As Double
    Dim lngҽ����� As Long
    Dim rsTemp As New adodb.Recordset
'    On Error Resume Next
    On Error GoTo errHandle
    If Not TxtInput.Visible Then Exit Sub
    blnUnValid = False
    TxtInput = Trim(TxtInput)
    If TxtInput = "" Then TxtInput = 0
    
    blnUnValid = Not IsNumeric(TxtInput)
    If Not blnUnValid Then blnUnValid = Not ((Abs(TxtInput) <= Abs(TxtInput.Tag)) And ((Val(TxtInput) >= 0 And Val(TxtInput.Tag) >= 0) Or (Val(TxtInput) <= 0 And Val(TxtInput.Tag) <= 0)))
    If blnUnValid Then TxtInput = Val(TxtInput.Tag)
    
    With RecChangeSendedData
        .MoveFirst
        .Find "λ��=" & CurCell.Row
        If .EOF Then Exit Sub
        
        '�ȼ���Ƿ���ҽ��������ҩƷ��¼
        '��������򲻹�
        '����ǣ����ϵͳ�����Ƿ�����δ����ҽ����ҩ�������������ҩ��Ϊ��
        '��������򲻹�
        dblCount = FormatEx(TxtInput.Text, 5)
        If dblCount <> 0 And blnҽ������ = False Then
            gstrSQL = "select ���� From ҩƷ�շ���¼ Where ID=[1]"
            Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[����Ƿ�������]", CLng(!Id))

            If (rsTemp!���� Like "1*") Then       '����
                gstrSQL = "Select nvl(ҽ�����,0) ҽ�����,Nvl(�����־,1) �����־ From ���˷��ü�¼ Where ID=(Select ����ID From ҩƷ�շ���¼ Where ID=[1])"
                Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[����Ƿ���ҽ��]", CLng(!Id))
            
                If Not rsTemp.EOF Then
                    If (rsTemp!�����־ = 1 Or rsTemp!�����־ = 4) And rsTemp!ҽ����� <> 0 Then
                        gstrSQL = "Select decode(ҽ��״̬,4,1,0) ���� From ����ҽ����¼ Where ID=[1]"
                        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[�жϸ�ҽ���Ƿ�����]", CLng(rsTemp!ҽ�����))
                        
                        If rsTemp!���� = 0 Then
                            dblCount = 0
                            'MsgBox "�ñ�ҽ����δ���ϣ�������ҩ��", vbInformation, gstrSysName
                        End If
                    End If
                End If
            End If
        End If
        
        Bill�ѷ�ҩ�嵥.TextMatrix(CurCell.Row, ����_�ѷ�ҩ�嵥.��ҩ��) = FormatEx(dblCount, 5)
        !��ҩ�� = Val(TxtInput.Text)
        .Update
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub MsgErr(ByVal strMsg As String)
    MsgBox strMsg, vbInformation, gstrSysName
End Sub

Private Sub AddCollect(ByVal dbl�ϼƽ�� As Double, Optional ByVal str���� As String = "�ϼ�")
    Dim intCol As Integer, str�ϼ� As String
    
    dbl�ϼƽ�� = Val(Format(dbl�ϼƽ��, "#####0.00;-#####0.00; ;"))
    str�ϼ� = zlCommFun.UppeMoney(dbl�ϼƽ��)
    
    Select Case TabShow.Tab
    Case 0
        With Billδ��ҩ�嵥
            .TextMatrix(.rows - 1, ����_δ��ҩ�嵥.����) = str����
            .TextMatrix(.rows - 1, ����_δ��ҩ�嵥.����ҽ��) = str����
            .TextMatrix(.rows - 1, ����_δ��ҩ�嵥.״̬) = Format(dbl�ϼƽ��, "#####0.00;-#####0.00; ;")
            .TextMatrix(.rows - 1, ����_δ��ҩ�嵥.����) = Format(dbl�ϼƽ��, "#####0.00;-#####0.00; ;")
            .TextMatrix(.rows - 1, ����_δ��ҩ�嵥.NO) = str�ϼ�
            .TextMatrix(.rows - 1, ����_δ��ҩ�嵥.����Ա) = str�ϼ�
            .TextMatrix(.rows - 1, ����_δ��ҩ�嵥.����) = str�ϼ�
            .TextMatrix(.rows - 1, ����_δ��ҩ�嵥.����) = str�ϼ�
            .TextMatrix(.rows - 1, ����_δ��ҩ�嵥.סԺ��) = str�ϼ�
            .TextMatrix(.rows - 1, ����_δ��ҩ�嵥.ҩƷ����) = str�ϼ�
            .TextMatrix(.rows - 1, ����_δ��ҩ�嵥.���) = str�ϼ�
            .TextMatrix(.rows - 1, ����_δ��ҩ�嵥.����) = str�ϼ�
            .TextMatrix(.rows - 1, ����_δ��ҩ�嵥.����) = str�ϼ�
            .TextMatrix(.rows - 1, ����_δ��ҩ�嵥.��) = str�ϼ�
            .TextMatrix(.rows - 1, ����_δ��ҩ�嵥.����) = str�ϼ�
            .TextMatrix(.rows - 1, ����_δ��ҩ�嵥.����) = str�ϼ�
            .TextMatrix(.rows - 1, ����_δ��ҩ�嵥.���) = str�ϼ�
            .TextMatrix(.rows - 1, ����_δ��ҩ�嵥.����) = str�ϼ�
            .TextMatrix(.rows - 1, ����_δ��ҩ�嵥.Ƶ��) = str�ϼ�
            .TextMatrix(.rows - 1, ����_δ��ҩ�嵥.�÷�) = str�ϼ�
            .TextMatrix(.rows - 1, ����_δ��ҩ�嵥.����ʱ��) = str�ϼ�
            .TextMatrix(.rows - 1, ����_δ��ҩ�嵥.˵��) = str�ϼ�
            
            .Row = .rows - 1
            .Col = 0: .CellAlignment = 4
'            .Col = 1: .CellAlignment = 4
            For intCol = 0 To .Cols - 1
                .Col = intCol
                .CellForeColor = glng��ҩ
            Next
            .RowData(.Row) = glng��ҩ
            .rows = .rows + 1
        End With
    Case 1
        If dbl�ϼƽ�� = 0 Then Exit Sub
        With Bill���ܷ�ҩ
            .TextMatrix(.rows - 1, 0) = str����
            .TextMatrix(.rows - 1, 1) = Format(dbl�ϼƽ��, "#####0.00;-#####0.00; ;")
            .TextMatrix(.rows - 1, 2) = str�ϼ�
            .TextMatrix(.rows - 1, 3) = str�ϼ�
            .TextMatrix(.rows - 1, 4) = str�ϼ�
            If Lng������ʾ = 1 Then
                .TextMatrix(.rows - 1, 5) = str�ϼ�
                .TextMatrix(.rows - 1, 6) = str�ϼ�
                .TextMatrix(.rows - 1, 7) = str�ϼ�
            End If
            
            .Row = .rows - 1
            .Col = 0: .CellAlignment = 4
            .Col = 1: .CellAlignment = 7
            For intCol = 0 To .Cols - 1
                .Col = intCol
                .CellForeColor = glng��ҩ
            Next
            .RowData(.Row) = glng��ҩ
        End With
    Case Else
        Exit Sub
    End Select
End Sub

Private Sub UpdateState(ByVal blnδ�� As Boolean, Optional ByVal blnȫѡ As Boolean = True)
    Dim intState As Integer, strState As String, lngλ�� As Long
    '����δ��ҩ�嵥���ѷ�ҩ�嵥��״̬
    
    intState = IIf(blnδ��, IIf(blnȫѡ, gIntδ��ҩ�嵥��ҩ, gIntδ��ҩ�嵥������), _
                            IIf(blnȫѡ, gInt�ѷ�ҩ�嵥��ҩ, gInt�ѷ�ҩ�嵥������))
    strState = IIf(blnδ��, IIf(blnȫѡ, "��ҩ", "������"), _
                            IIf(blnȫѡ, "��ҩ", "������"))
    
    If blnδ�� Then
'        With RecChangeData
'            If .RecordCount = 0 Then Exit Sub
'            .MoveFirst
'
'            Do While Not .EOF
'
'                .MoveNext
'            Loop
'        End With
    Else
        With RecChangeSendedData
            If TxtInput.Visible Then TxtInput.Visible = False
            
            If .RecordCount = 0 Then Exit Sub
            .MoveFirst
            
            Do While Not .EOF
                If !�ɲ��� = 1 Then
                    lngλ�� = !λ��
                    Bill�ѷ�ҩ�嵥.TextMatrix(lngλ��, ����_�ѷ�ҩ�嵥.״̬) = strState
                    If intState = 3 Then
                        Bill�ѷ�ҩ�嵥.TextMatrix(lngλ��, ����_�ѷ�ҩ�嵥.��ҩ��) = Bill�ѷ�ҩ�嵥.TextMatrix(lngλ��, ����_�ѷ�ҩ�嵥.׼����)
                    Else
                        Bill�ѷ�ҩ�嵥.TextMatrix(lngλ��, ����_�ѷ�ҩ�嵥.��ҩ��) = ""
                    End If
                    
                    !ִ��״̬ = intState
                    .Update
                End If
                .MoveNext
            Loop
        End With
    End If
    Call SetMenuAndToolbarState
End Sub

Private Sub FindRecord(Optional ByVal BlnFirst As Boolean = True)
    Dim RecObject As adodb.Recordset
    Static lngδ�� As Long, lng�ѷ� As Long
    Dim lngRecord As Long
    Dim strMsg As String
    Dim blnExist As Boolean
    'lngLocate:��ʼ����̬��������ҳ�淢���ı��ˢ��ʱ��
    
    MnuViewLocateNext.Enabled = False
    MnuViewLocateNext.Tag = 0
    If strFind = "" Then Exit Sub
    
    '����ָ�����ݵļ�¼
    Select Case TabShow.Tab
    Case 0      'δ��ҩ�嵥
        lngRecord = lngδ��
        Set RecObject = Recδ��.Clone
    Case 4      '�ѷ�ҩ�嵥
        lngRecord = lng�ѷ�
        Set RecObject = Rec�ѷ�.Clone
    End Select
    
    '�Ϸ�����֤
    If RecObject Is Nothing Then Exit Sub
    If RecObject.State = 0 Then Exit Sub
    If RecObject.RecordCount = 0 Then Exit Sub
    
    RecObject.MoveFirst
    If Not BlnFirst Then
        RecObject.Find "λ��=" & lngRecord
        If RecObject.EOF Then RecObject.MoveFirst
        RecObject.MoveNext
    End If
    
    Do While Not RecObject.EOF
        '���Ҹü�¼���Ƿ����ڲ�ӳ���¼����
        Select Case TabShow.Tab
        Case 0
            With RecChangeData
                If .RecordCount = 0 Then Exit Sub
                .Filter = strFind
                If .RecordCount <> 0 Then
                    .Find "λ��=" & RecObject!λ��
                    blnExist = Not (.EOF)
                End If
                .Filter = 0
            End With
        Case 4
            With RecChangeSendedData
                If .RecordCount = 0 Then Exit Sub
                .Filter = strFind
                If .RecordCount <> 0 Then
                    .Find "λ��=" & RecObject!λ��
                    blnExist = Not (.EOF)
                End If
                .Filter = 0
            End With
        End Select
        If blnExist Then Exit Do
        RecObject.MoveNext
    Loop
    If Not blnExist Then
        If MsgBox("���ҽ������Ƿ��ͷ����һ�飿", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Call FindRecord(True)
        End If
        Exit Sub
    End If
    
    MnuViewLocateNext.Enabled = True
    MnuViewLocateNext.Tag = 1
    'ѡ����ľ�����
    Select Case TabShow.Tab
    Case 0
        lngδ�� = RecObject!λ��
        Billδ��ҩ�嵥.Row = RecObject!λ��
        Billδ��ҩ�嵥_EnterCell
    Case 4
        lng�ѷ� = RecObject!λ��
        Bill�ѷ�ҩ�嵥.Row = RecObject!λ��
        Bill�ѷ�ҩ�嵥_EnterCell
    End Select
End Sub

Private Function CheckSpec(ByVal strRecipeKey As String) As Boolean
    Dim strNote As String
    Dim rsTemp As New adodb.Recordset
    '�Զ�����ҩƷ���м��
    On Error GoTo errHandle
    gstrSQL = "SELECT Distinct '['||C.����||']'||NVL(L.����,C.����) Ʒ��,X.������� " & _
             "   FROM (Select ҩ��ID,ҩƷID From ҩƷ��� Where ҩƷID IN (" & strRecipeKey & ")) B, " & _
             "        �շ���ĿĿ¼ C, " & _
             "        �շ���Ŀ���� L, " & _
             "        ҩƷ����     X " & _
             "  WHERE X.ҩ��ID = B.ҩ��ID And B.ҩƷID = C.ID  " & _
             "        AND C.ID = L.�շ�ϸĿID(+) AND L.����(+) = 3 AND L.����(+) = 1  " & _
             "        AND X.������� <> '��ͨҩ' " & _
             "  Order by X.�������"
    Call zldatabase.OpenRecordset(rsTemp, gstrSQL, "�Զ�����ҩƷ���м��")
    If rsTemp.RecordCount = 0 Then
        CheckSpec = True
        Exit Function
    End If
    
    With rsTemp
        Do While Not .EOF
            strNote = strNote & vbCrLf & Space(4) & !������� & "-" & !Ʒ��
            .MoveNext
        Loop
    End With
    If MsgBox("�Ƿ�����¶����顢������ҩƷ���з�ҩ��" & strNote, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    CheckSpec = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub BuildRecord(Optional ByVal bln��ҩ As Boolean = True)
    Dim intRow As Integer, intRows As Integer
    Dim strNo As String, lng���� As Long, str��� As String, lng����ID As Long
    Dim blnAdd As Boolean
    
    Call InitCheckRec
    '���ݴ���ҩ������ҩ�嵥�������ݻ�ȡ��ϸ���
    If bln��ҩ Then
        intRows = RecChangeData.RecordCount
        If RecChangeData.RecordCount <> 0 Then RecChangeData.MoveFirst
        For intRow = 1 To intRows
            If Val(RecChangeData!ִ��״̬) = 1 Then
                strNo = RecChangeData!NO
                lng���� = Val(RecChangeData!����)
                lng����ID = RecChangeData!����ID
                
                rs���.Filter = "���ݱ�ʶ='" & strNo & "|" & lng���� & "'"
                blnAdd = (rs���.RecordCount = 0)
                If Not blnAdd Then
                    rs���.Find "����ID=" & lng����ID
                    blnAdd = rs���.EOF
                End If
                
                If blnAdd Then rs���.AddNew
                rs���!���ݱ�ʶ = strNo & "|" & lng����
                rs���!����ID = lng����ID
                
                str��� = NVL(rs���!���)
                If InStr(1, "," & str��� & ",", "," & Val(RecChangeData!���) & ",") = 0 Then
                    If str��� = "" Then
                        str��� = Val(RecChangeData!���)
                    Else
                        str��� = str��� & "," & Val(RecChangeData!���)
                    End If
                    rs���!��� = str���
                End If
                rs���.Update
                rs���.Filter = 0
            End If
            RecChangeData.MoveNext
        Next
        If RecChangeData.RecordCount <> 0 Then RecChangeData.MoveFirst
    Else
        intRows = RecChangeSendedData.RecordCount
        If RecChangeSendedData.RecordCount <> 0 Then RecChangeSendedData.MoveFirst
        For intRow = 1 To intRows
            If Val(RecChangeSendedData!ִ��״̬) = 3 Then
                If Val(NVL(RecChangeSendedData!��ҩ��, 0)) <> 0 Then
                    strNo = RecChangeSendedData!NO
                    lng���� = Val(RecChangeSendedData!����)
                    lng����ID = RecChangeSendedData!����ID
                    
                    rs���.Filter = "���ݱ�ʶ='" & strNo & "|" & lng���� & "'"
                    blnAdd = (rs���.RecordCount = 0)
                    If Not blnAdd Then
                        rs���.Find "����ID=" & lng����ID
                        blnAdd = rs���.EOF
                    End If
                    
                    If blnAdd Then rs���.AddNew
                    rs���!���ݱ�ʶ = strNo & "|" & lng����
                    rs���!����ID = lng����ID
                    
                    str��� = NVL(rs���!���)
                    If InStr(1, "," & str��� & ",", "," & Val(RecChangeSendedData!���) & ",") = 0 Then
                        If str��� = "" Then
                            str��� = Val(RecChangeSendedData!���)
                        Else
                            str��� = str��� & "," & Val(RecChangeSendedData!���)
                        End If
                        rs���!��� = str���
                    End If
                    rs���.Update
                    rs���.Filter = 0
                End If
            End If
            RecChangeSendedData.MoveNext
        Next
        If RecChangeSendedData.RecordCount <> 0 Then RecChangeSendedData.MoveFirst
    End If

    '��ӡ
    intRows = rs���.RecordCount
    If rs���.RecordCount <> 0 Then rs���.MoveFirst
    For intRow = 1 To intRows
        Debug.Print rs���!���ݱ�ʶ & "," & rs���!����ID & "," & rs���!���
        rs���.MoveNext
    Next
    If rs���.RecordCount <> 0 Then rs���.MoveFirst
End Sub

Private Function CheckCorrelation() As Boolean
    Dim strNo As String, lng���� As Long, str��� As String, lng����ID As Long
    '��鴦���Ƿ��ѽ��ʡ����ò����Ƿ��ѳ�Ժ������Ȩ�޽��м��
    With rs���
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            strNo = !���ݱ�ʶ
            lng���� = Split(strNo, "|")(1)
            strNo = Split(strNo, "|")(0)
            lng����ID = !����ID
            str��� = NVL(!���)
            If Not IsReceiptBalance_Charge(mstrPrivs, lng����, strNo, str���) Then Exit Function
            If Not IsOutPatient(mstrPrivs, lng����, strNo, lng����ID) Then Exit Function
            .MoveNext
        Loop
    End With
    
    CheckCorrelation = True
End Function

Private Sub InitCheckRec()
    Set rs��� = New adodb.Recordset
    With rs���
        If .State = 1 Then .Close
        .Fields.Append "���ݱ�ʶ", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "����ID", adDouble, 18, adFldIsNullable
        .Fields.Append "���", adLongVarChar, 500, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub











Private Sub txtPati_GotFocus()
    Call SelAll(txtPati)
    
    txtPati.PasswordChar = ""
    txtPati.MaxLength = 0
    
    If Val(lblPatiInputType.Tag) = PatiInfo.���￨ Then
        If gtype_UserSysParms.P12_���￨�Ƿ�������ʾ Then
            txtPati.PasswordChar = "*"
        End If
        txtPati.MaxLength = gtype_UserSysParms.P20_���￨�ų���
    End If
End Sub
Private Sub txtPati_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
         Call txtPati_Validate(True)
    End If
End Sub

Private Sub txtPati_KeyPress(KeyAscii As Integer)
    mblnCard = False
    
    If Val(lblPatiInputType.Tag) = PatiInfo.סԺ�� Or Val(lblPatiInputType.Tag) = PatiInfo.����ID Then
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyEscape Or KeyAscii = vbKeyBack Then Exit Sub
        KeyAscii = 0
    ElseIf Val(lblPatiInputType.Tag) = PatiInfo.���� Then
        mblnCard = zlCommFun.InputIsCard(txtPati, KeyAscii, glngSys)
    ElseIf Val(lblPatiInputType.Tag) = PatiInfo.���￨ Then
        mblnCard = (KeyAscii <> 8 And Len(txtPati.Text) = gtype_UserSysParms.P20_���￨�ų��� - 1 And txtPati.SelLength <> Len(txtPati.Text))
    End If
End Sub

    
Private Sub txtPati_Validate(Cancel As Boolean)
    Dim strDeptInfo As String
    Dim strInput As String
    
    'ȡ�������ƣ����˵�ǰ����������ȡ������¼
    '��ȡ��������Ϣ�󣬷���������ʽ��������Ϣ-��������
    If InStr(Trim(txtPati.Text), "-") > 0 Then
        'ȡ��-��ǰ���������Ϣ
        strInput = Mid(Trim(txtPati.Text), 1, InStr(Trim(txtPati.Text), "-") - 1)
    Else
        strInput = Trim(txtPati.Text)
    End If
    
    If strInput = "" Then Exit Sub
    
    If Val(lblPatiInputType.Tag) = PatiInfo.���ݺ� Then
        If IsNumeric(strInput) Then
            strInput = GetFullNO(strInput, 14)
        End If
    End If
    
    strDeptInfo = GetPatiInfo(Val(lblPatiInputType.Tag), strInput)
    
    If strDeptInfo <> "" Then
        mintLastDeptType = 2
        tbsType.Tabs(3).Selected = True
        
        txt����.Text = Mid(Split(strDeptInfo, "|")(0), InStr(Split(strDeptInfo, "|")(0), ",") + 1)
        txt����.Tag = Mid(Split(strDeptInfo, "|")(0), 1, InStr(Split(strDeptInfo, "|")(0), ",") - 1)
        
        Select Case Val(lblPatiInputType.Tag)
        Case PatiInfo.����
            If mblnCard = True Then
                txtPati.Text = UCase(strInput)
                txtPati.Tag = Mid(Split(strDeptInfo, "|")(1), 1, InStr(Split(strDeptInfo, "|")(1), ",") - 1)
            Else
                txtPati.Text = Mid(Split(strDeptInfo, "|")(1), InStr(Split(strDeptInfo, "|")(1), ",") + 1)
            End If
        Case PatiInfo.���￨
            txtPati.PasswordChar = ""
            txtPati.MaxLength = 0
            txtPati.Text = Mid(Split(strDeptInfo, "|")(1), InStr(Split(strDeptInfo, "|")(1), ",") + 1)
            txtPati.Tag = Mid(Split(strDeptInfo, "|")(1), 1, InStr(Split(strDeptInfo, "|")(1), ",") - 1)
        Case Else
            txtPati.Text = strInput & "-" & Mid(Split(strDeptInfo, "|")(1), InStr(Split(strDeptInfo, "|")(1), ",") + 1)
        End Select
        
        DoEvents
        
        Call cmdRefresh_Click
    End If
End Sub

Private Sub txt��ҩ;��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub


Private Sub txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTemp As adodb.Recordset
    If KeyCode = vbKeyReturn Then
        If Trim(txt����.Text) = "" Then
            txt����.Tag = ""
            Exit Sub
        End If
        
        Set rsTemp = SelectDept(tbsType.SelectedItem.Index - 1, Trim(txt����.Text))
        
        If Not rsTemp Is Nothing Then
            txt����.Tag = rsTemp("ID")
            txt����.Text = rsTemp("����")
        End If
    End If
End Sub
Private Sub txt����_Validate(Cancel As Boolean)
    If Trim(txt����.Text) = "" Then
        txt����.Tag = ""
        Exit Sub
    End If
End Sub


Private Sub txt������_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call txt������_LostFocus
        txt������.Visible = False
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Or KeyAscii = Asc("-") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub








Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub


Private Sub txt������_LostFocus()
    Dim dbl������ As Double
    Dim dblӦ���� As Double
    Dim dblʵ���� As Double
    
    dblӦ���� = Val(Bill���ܷ�ҩ.TextMatrix(LngLastRow, ����_���һ����嵥.Ӧ������)) - Val(Bill���ܷ�ҩ.TextMatrix(LngLastRow, ����_���һ����嵥.��������))
    
    If lngLastCol = ����_���һ����嵥.ʵ������ Then
        dblʵ���� = Val(txt������.Text)
        If dblʵ���� > dblӦ���� Or dblʵ���� < 0 Then
            Bill���ܷ�ҩ.TextMatrix(LngLastRow, ����_���һ����嵥.ʵ������) = FormatEx(dblӦ����, 5)
            Bill���ܷ�ҩ.TextMatrix(LngLastRow, ����_���һ����嵥.��������) = 0
        Else
            Bill���ܷ�ҩ.TextMatrix(LngLastRow, ����_���һ����嵥.ʵ������) = FormatEx(dblʵ����, 5)
            Bill���ܷ�ҩ.TextMatrix(LngLastRow, ����_���һ����嵥.��������) = FormatEx(dblӦ���� - Val(Bill���ܷ�ҩ.TextMatrix(LngLastRow, ����_���һ����嵥.ʵ������)), 5)
        End If
    ElseIf lngLastCol = ����_���һ����嵥.�������� Then
        dbl������ = Val(txt������.Text)
        If dbl������ > dblӦ���� Or dbl������ < 0 Then
            Bill���ܷ�ҩ.TextMatrix(LngLastRow, ����_���һ����嵥.ʵ������) = FormatEx(dblӦ����, 5)
            Bill���ܷ�ҩ.TextMatrix(LngLastRow, ����_���һ����嵥.��������) = 0
        Else
            Bill���ܷ�ҩ.TextMatrix(LngLastRow, ����_���һ����嵥.ʵ������) = FormatEx(dblӦ���� - Val(Bill���ܷ�ҩ.TextMatrix(LngLastRow, ����_���һ����嵥.��������)), 5)
            Bill���ܷ�ҩ.TextMatrix(LngLastRow, ����_���һ����嵥.��������) = FormatEx(dbl������, 5)
        End If
    End If
            
    DoEvents
    
    Bill���ܷ�ҩ.Row = LngLastRow
    Bill���ܷ�ҩ.Col = ����_���һ����嵥.ʵ������
    If Val(Bill���ܷ�ҩ.TextMatrix(LngLastRow, ����_���һ����嵥.ʵ������)) < 0 Then
        Bill���ܷ�ҩ.CellForeColor = vbRed
    ElseIf Val(Bill���ܷ�ҩ.TextMatrix(LngLastRow, ����_���һ����嵥.ʵ������)) > 0 Then
        Bill���ܷ�ҩ.CellForeColor = vbBlue
    End If
End Sub


