VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Begin VB.Form frmManageBed 
   AutoRedraw      =   -1  'True
   Caption         =   "������λ����"
   ClientHeight    =   5880
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8715
   Icon            =   "frmManageBed.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   5880
   ScaleWidth      =   8715
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lvwBeds 
      Height          =   4560
      Left            =   15
      TabIndex        =   3
      Tag             =   "�ɱ仯��"
      Top             =   870
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   8043
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "img32"
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   1376
      BandCount       =   2
      _CBWidth        =   8715
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbr"
      MinHeight1      =   720
      Width1          =   7635
      NewRow1         =   0   'False
      BandForeColor2  =   8388608
      Caption2        =   "��ǰ����"
      Child2          =   "cboUnit"
      MinWidth2       =   1995
      MinHeight2      =   300
      Width2          =   1215
      UseCoolbarColors2=   0   'False
      NewRow2         =   0   'False
      Begin VB.ComboBox cboUnit 
         Height          =   300
         Left            =   6630
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   1995
      End
      Begin MSComctlLib.Toolbar tbr 
         Height          =   720
         Left            =   165
         TabIndex        =   2
         Top             =   30
         Width           =   5460
         _ExtentX        =   9631
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imgGray"
         HotImageList    =   "imgColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   14
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ԥ��"
               Key             =   "Preview"
               Description     =   "Ԥ��"
               Object.ToolTipText     =   "Ԥ��"
               Object.Tag             =   "Ԥ��"
               ImageKey        =   "Preview"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "Print"
               Description     =   "��ӡ"
               Object.ToolTipText     =   "��ӡ"
               Object.Tag             =   "��ӡ"
               ImageKey        =   "Print"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Add"
               Description     =   "����"
               Object.ToolTipText     =   "���Ӳ���"
               Object.Tag             =   "����"
               ImageKey        =   "New"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Modi"
               Description     =   "����"
               Object.ToolTipText     =   "��������"
               Object.Tag             =   "����"
               ImageKey        =   "Modi"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Del"
               Description     =   "����"
               Object.ToolTipText     =   "��������"
               Object.Tag             =   "����"
               ImageKey        =   "Del"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Remedy"
               Description     =   "����"
               Object.ToolTipText     =   "���մ�תΪ���ɴ�"
               Object.Tag             =   "����"
               ImageKey        =   "Remedy"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�մ�"
               Key             =   "Empty"
               Description     =   "�մ�"
               Object.ToolTipText     =   "���޺õĴ�תΪ�մ�"
               Object.Tag             =   "�մ�"
               ImageKey        =   "Empty"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�б�"
               Key             =   "View"
               Description     =   "�б�"
               Object.ToolTipText     =   "��λ�б���ʾ��ʽ"
               Object.Tag             =   "�б�"
               ImageKey        =   "View"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Icon"
                     Object.Tag             =   "��ͼ��(&G)"
                     Text            =   "��ͼ��(&G)"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Small"
                     Object.Tag             =   "Сͼ��(&M)"
                     Text            =   "Сͼ��(&M)"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "List"
                     Object.Tag             =   "�б�(&L)"
                     Text            =   "�б�(&L)"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Detail"
                     Object.Tag             =   "��ϸ����(&D)"
                     Text            =   "��ϸ����(&D)"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Description     =   "����"
               Object.ToolTipText     =   "��ǰ��������"
               Object.Tag             =   "����"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Quit"
               Description     =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   5520
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmManageBed.frx":030A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10292
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
            AutoSize        =   2
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
   Begin MSComctlLib.ImageList imgColor 
      Left            =   60
      Top             =   450
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
            Picture         =   "frmManageBed.frx":0B9E
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":0DB8
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":0FD2
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":11EC
            Key             =   "Modi"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":1406
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":1620
            Key             =   "View"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":183A
            Key             =   "Remedy"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":1A54
            Key             =   "Empty"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":1C6E
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":1E88
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   645
      Top             =   450
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
            Picture         =   "frmManageBed.frx":20A2
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":22BC
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":24D6
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":26F0
            Key             =   "Modi"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":290A
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":2B24
            Key             =   "View"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":2D3E
            Key             =   "Remedy"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":2F58
            Key             =   "Empty"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":3172
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":338C
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img32 
      Left            =   2760
      Top             =   495
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":35A6
            Key             =   "Empty"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":38C0
            Key             =   "M_Empty"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":3BDA
            Key             =   "F_Empty"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":3EF4
            Key             =   "Holding"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":420E
            Key             =   "Remedy"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":4528
            Key             =   "MASK_�Ӵ�"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":4842
            Key             =   "MASK_�Ǳ�"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":4B5C
            Key             =   "MASK_����"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":4E76
            Key             =   "MASK_����_�Ӵ�"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":5190
            Key             =   "MASK_����_�Ǳ�"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   3345
      Top             =   495
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":54AA
            Key             =   "Empty"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":57C4
            Key             =   "M_Empty"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":5ADE
            Key             =   "F_Empty"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":5DF8
            Key             =   "Holding"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":6112
            Key             =   "Remedy"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":642C
            Key             =   "MASK_�Ӵ�"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":6586
            Key             =   "MASK_�Ǳ�"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":66E0
            Key             =   "MASK_����"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":683A
            Key             =   "MASK_����_�Ӵ�"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":6994
            Key             =   "MASK_����_�Ǳ�"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFile_PrintSet 
         Caption         =   "��ӡ����(&S)"
      End
      Begin VB.Menu mnuFile_Preview 
         Caption         =   "��ӡԤ��(&V)"
      End
      Begin VB.Menu mnuFile_Print 
         Caption         =   "��ӡ(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFile_Excel 
         Caption         =   "�����&Excel"
      End
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_quit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuEdit_Add 
         Caption         =   "����(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEdit_Modi 
         Caption         =   "����(&M)"
      End
      Begin VB.Menu mnuEdit_Del 
         Caption         =   "����(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEdit_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Remedy 
         Caption         =   "ת����(&R)"
      End
      Begin VB.Menu mnuEdit_Empty 
         Caption         =   "ת�մ�(&E)"
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
         Begin VB.Menu mnuViewToolUnit 
            Caption         =   "����ѡ��(&U)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewTool_1 
            Caption         =   "-"
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
      Begin VB.Menu mnuView_5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewSelCol 
         Caption         =   "ѡ����(&C)"
      End
      Begin VB.Menu mnuEdit_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuView_ListView 
         Caption         =   "��ͼ��(&G)"
         Index           =   0
      End
      Begin VB.Menu mnuView_ListView 
         Caption         =   "Сͼ��(&M)"
         Index           =   1
      End
      Begin VB.Menu mnuView_ListView 
         Caption         =   "�б�(&L)"
         Index           =   2
      End
      Begin VB.Menu mnuView_ListView 
         Caption         =   "��ϸ����(&D)"
         Checked         =   -1  'True
         Index           =   3
      End
      Begin VB.Menu mnuView_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuView_reFlash 
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
Attribute VB_Name = "frmManageBed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������

Private mblnUnload As Boolean
Private mintEmpty As Integer, intHolding, intRemedy As Integer
Private Const STR_HEAD = "����,600,0,1;����,1200,0,2;�����,800,0,2;״̬,600,0,2;�Ա����,1000,0,2;�ȼ�,1000,0,2;��λ����,1000,0,2;����,1000,0,0;�Ա�,600,0,0;����,600,0,0"
Private mstrPrivs As String

Private Sub cboUnit_Click()
    Call ReadBeds(cboUnit.ItemData(cboUnit.ListIndex))
    Call SetMenuState
    Me.Refresh
End Sub

Private Sub cbr_Resize()
    Form_Resize
End Sub

Private Sub Form_Activate()
    If mblnUnload Then Unload Me
End Sub

Private Sub Form_Load()
    
    Call RestoreWinState(Me, App.ProductName)
    If lvwBeds.ColumnHeaders.Count = 0 Then
        Call zlcontrol.LvwSelectColumns(lvwBeds, STR_HEAD, True)
    End If
    
    mstrPrivs = gstrPrivs
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs)
    
    '���ݱ����б�ʽ���ò˵�
     Call SetView(lvwBeds.View)
     
     Call MakeBedIcon
        
    '��ȡ����
    If Not InitUnits Then mblnUnload = True: Exit Sub
    If cboUnit.ListIndex = -1 Then
        MsgBox "�㲻�������в�����Ȩ��,���Ҳ���ȷ������������,����ʹ�ô�λ����", vbExclamation, gstrSysName
        mblnUnload = True: Exit Sub
    End If
    
    If Not ReadBeds(cboUnit.ItemData(cboUnit.ListIndex)) Then
        mblnUnload = True: Exit Sub
    End If
    Call SetMenuState
End Sub

Private Sub Form_Resize()
    Dim cbrH As Long '������ռ�ø߶�
    Dim staH As Long '״̬��ռ�ø߶�
    
    On Error Resume Next
    
    If WindowState = 1 Then Exit Sub
    
    '����ؼ���Ⱥ͸߶�
    cbrH = IIf(cbr.Visible, cbr.Height, 0)
    staH = IIf(sta.Visible, sta.Height, 0)
    
    With lvwBeds
        .Left = Me.ScaleLeft
        .Top = Me.ScaleTop + cbrH
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - cbrH - staH
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnUnload = False
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub lvwBeds_DblClick()
    mnuEdit_Modi_Click
End Sub

Private Sub lvwBeds_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Call SetMenuState
End Sub

Private Sub lvwBeds_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And Not lvwBeds.SelectedItem Is Nothing Then
        mnuEdit_Modi_Click
    End If
End Sub

Private Sub lvwBeds_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objItem As ListItem
    
    If Button = 2 Then
        PopupMenu mnuEdit, 2
    ElseIf lvwBeds.View <> lvwReport Then
        Set objItem = lvwBeds.HitTest(X, Y)
        If Not objItem Is Nothing Then
            With objItem
                sta.Panels(2) = "����[" & Trim(.Text) & "]" & _
                    " ״̬:" & .SubItems(lvwBeds.ColumnHeaders("_״̬").Index - 1) & _
                    " �Ա����:" & .SubItems(lvwBeds.ColumnHeaders("_�Ա����").Index - 1) & _
                    " ����:" & .SubItems(lvwBeds.ColumnHeaders("_����").Index - 1) & _
                    " �ȼ�:" & .SubItems(lvwBeds.ColumnHeaders("_�ȼ�").Index - 1)
            End With
        Else
            sta.Panels(2) = "��ǰ������ " & lvwBeds.ListItems.Count & " �Ų���,���в���ռ�� " & intHolding & " ��,�մ� " & mintEmpty & " ��,�������� " & intRemedy & " �ţ�"
        End If
    Else
        sta.Panels(2) = "��ǰ������ " & lvwBeds.ListItems.Count & " �Ų���,���в���ռ�� " & intHolding & " ��,�մ� " & mintEmpty & " ��,�������� " & intRemedy & " �ţ�"
    End If
End Sub

Private Sub mnuEdit_Del_Click()
    Dim intIdx As Integer, strSQL As String
    
    If lvwBeds.SelectedItem Is Nothing Then
        MsgBox "��ѡ��Ҫ�����Ĳ�����", vbExclamation, gstrSysName: Exit Sub
    End If
    If lvwBeds.SelectedItem.SubItems(lvwBeds.ColumnHeaders("_״̬").Index - 1) = "ռ��" Then
        MsgBox "�ò����ѱ�����ռ��,���ڲ��ܳ�����", vbExclamation, gstrSysName: Exit Sub
    End If
    If MsgBox("ȷʵҪ��������" & Mid(lvwBeds.SelectedItem.Key, 2) & " ��", vbExclamation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    On Error GoTo errH
    intIdx = lvwBeds.SelectedItem.Index
    
    strSQL = "zl_��λ״����¼_Delete('" & Mid(lvwBeds.SelectedItem.Key, 2) & "'," & cboUnit.ItemData(cboUnit.ListIndex) & ")"
    
    On Error GoTo errH
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    On Error GoTo 0
    
    lvwBeds.ListItems.Remove intIdx
    If lvwBeds.ListItems.Count <> 0 Then
        If intIdx <= lvwBeds.ListItems.Count Then
            lvwBeds.ListItems(intIdx).Selected = True
        Else
            lvwBeds.ListItems(lvwBeds.ListItems.Count).Selected = True
        End If
        lvwBeds.SelectedItem.EnsureVisible
    End If
    Call SetBedNOLen
    Call SetMenuState
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEdit_Empty_Click()
    Dim strSQL As String
    
    If lvwBeds.SelectedItem Is Nothing Then
        MsgBox "��ѡ���Ѿ����ɺõĲ�����", vbExclamation, gstrSysName: Exit Sub
    End If
    If lvwBeds.SelectedItem.SubItems(lvwBeds.ColumnHeaders("_״̬").Index - 1) <> "����" Then
        MsgBox "�ò���û�н�������,����ִ�иò�����", vbExclamation, gstrSysName: Exit Sub
    End If
    
    strSQL = "zl_��λ״����¼_REUSE('" & Mid(lvwBeds.SelectedItem.Key, 2) & "'," & cboUnit.ItemData(cboUnit.ListIndex) & ")"
    
    On Error GoTo errH
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    On Error GoTo 0
    
    lvwBeds.SelectedItem.SubItems(lvwBeds.ColumnHeaders("_״̬").Index - 1) = "�մ�"
    If lvwBeds.SelectedItem.SubItems(lvwBeds.ColumnHeaders("_�Ա����").Index - 1) = "�д�" Then
        lvwBeds.SelectedItem.Icon = "M_Empty"
        lvwBeds.SelectedItem.SmallIcon = "M_Empty"
    ElseIf lvwBeds.SelectedItem.SubItems(lvwBeds.ColumnHeaders("_�Ա����").Index - 1) = "Ů��" Then
        lvwBeds.SelectedItem.Icon = "F_Empty"
        lvwBeds.SelectedItem.SmallIcon = "F_Empty"
    Else
        lvwBeds.SelectedItem.Icon = "Empty"
        lvwBeds.SelectedItem.SmallIcon = "Empty"
    End If
    
    Call SetBedIcon(lvwBeds, lvwBeds.SelectedItem)
    
    Call SetMenuState
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEdit_Modi_Click()
    If lvwBeds.SelectedItem Is Nothing Then
        MsgBox "��ѡ��Ҫ�����Ĳ�����", vbExclamation, gstrSysName: Exit Sub
    End If
    If lvwBeds.SelectedItem.SubItems(lvwBeds.ColumnHeaders("_״̬").Index - 1) = "ռ��" Then
        MsgBox "�ò����ѱ�����ռ��,���ڲ��ܽ��е�����", vbExclamation, gstrSysName: Exit Sub
    End If
    If lvwBeds.SelectedItem.SubItems(lvwBeds.ColumnHeaders("_״̬").Index - 1) = "����" Then
        MsgBox "�ò�����������,���ڲ��ܽ��е�����", vbExclamation, gstrSysName: Exit Sub
    End If
    
    On Error Resume Next
    Err.Clear
    
    frmEditBed.mblnModi = True
    Set frmEditBed.mlvwBeds = lvwBeds
    Set frmEditBed.mobjSta = sta
    frmEditBed.mlngUnit = cboUnit.ItemData(cboUnit.ListIndex)
    frmEditBed.Show 1, Me
    
    If gblnOK Then Call SetMenuState
End Sub

Private Sub mnuEdit_Add_Click()
    On Error Resume Next
    Err.Clear
    
    frmEditBed.mblnModi = False
    frmEditBed.mlngUnit = cboUnit.ItemData(cboUnit.ListIndex)
    Set frmEditBed.mlvwBeds = lvwBeds
    Set frmEditBed.mobjSta = sta
    frmEditBed.Show 1, Me
End Sub

Private Sub mnuEdit_Remedy_Click()
    Dim strSQL As String
    
    If lvwBeds.SelectedItem Is Nothing Then
        MsgBox "��ѡ��Ҫ���ɵĲ�����", vbExclamation, gstrSysName: Exit Sub
    End If
    If lvwBeds.SelectedItem.SubItems(lvwBeds.ColumnHeaders("_״̬").Index - 1) <> "�մ�" Then
        MsgBox "�ò������ǿմ�,����ִ�иò�����", vbExclamation, gstrSysName: Exit Sub
    End If
    
    strSQL = "zl_��λ״����¼_STOP('" & Mid(lvwBeds.SelectedItem.Key, 2) & "'," & cboUnit.ItemData(cboUnit.ListIndex) & ")"
    
    On Error GoTo errH
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    lvwBeds.SelectedItem.Icon = "Remedy"
    lvwBeds.SelectedItem.SmallIcon = "Remedy"
    lvwBeds.SelectedItem.SubItems(lvwBeds.ColumnHeaders("_״̬").Index - 1) = "����"
    
    Call SetBedIcon(lvwBeds, lvwBeds.SelectedItem)
    
    Call SetMenuState
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuFile_quit_Click()
    Unload Me
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpTitle_Click()
ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    Dim lngUnitID As Long, str���� As Long
        
    If cboUnit.ListIndex <> -1 Then lngUnitID = cboUnit.ItemData(cboUnit.ListIndex)
    If Not lvwBeds.SelectedItem Is Nothing Then str���� = Trim(lvwBeds.SelectedItem.Text)
    
    Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
        "����=" & lngUnitID, _
        "����=" & str����)
End Sub

Private Sub mnuView_ListView_Click(Index As Integer)
    Call SetView(CByte(Index))
End Sub

Private Sub mnuView_reFlash_Click()
    Call ReadBeds(cboUnit.ItemData(cboUnit.ListIndex))
    Call SetMenuState
    Me.Refresh
End Sub

Private Sub mnuViewSelCol_Click()
    If zlcontrol.LvwSelectColumns(lvwBeds, STR_HEAD) Then
        mnuView_reFlash_Click
        Call SetView(3)
    End If
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    sta.Visible = Not sta.Visible
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim i As Integer
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For i = 1 To tbr.Buttons.Count
        tbr.Buttons(i).Caption = IIf(mnuViewToolText.Checked, tbr.Buttons(i).Tag, "")
    Next
    cbr.Bands(1).MinHeight = tbr.ButtonHeight
    Form_Resize
End Sub

Private Sub mnuViewToolUnit_Click()
    mnuViewToolUnit.Checked = Not mnuViewToolUnit.Checked
    If mnuViewToolButton.Checked Then cbr.Bands(1).Visible = False
    cbr.Bands(2).Visible = Not cbr.Bands(2).Visible
    If mnuViewToolButton.Checked Then cbr.Bands(1).Visible = True
    cbr.Visible = cbr.Bands(2).Visible Or cbr.Bands(1).Visible
    Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    cbr.Bands(1).Visible = Not cbr.Bands(1).Visible
    mnuViewToolText.Enabled = Not mnuViewToolText.Enabled
    cbr.Visible = cbr.Bands(2).Visible Or cbr.Bands(1).Visible
    Form_Resize
End Sub

Private Sub tbr_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Quit"
            mnuFile_quit_Click
        Case "View"
            Call SetView((lvwBeds.View + 1) Mod 4)
        Case "Add"
            mnuEdit_Add_Click
        Case "Modi"
            mnuEdit_Modi_Click
        Case "Del"
            mnuEdit_Del_Click
        Case "Empty"
            mnuEdit_Empty_Click
        Case "Remedy"
            mnuEdit_Remedy_Click
        Case "Print"
            mnuFile_Print_Click
        Case "Preview"
            mnuFile_PreView_Click
        Case "Help"
            mnuHelpTitle_Click
    End Select
End Sub

Private Function InitUnits() As Boolean
'���ܣ���ʼ��סԺ����
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer, lngUnitID As Long, blnLimitUnit As Boolean
    Dim strUnitIDs As String
    
    On Error GoTo errH
        
    '��������۲���
    blnLimitUnit = InStr(mstrPrivs, "���в���") = 0
    '����30922 by lesfeng 2010-06-18 b
    If blnLimitUnit Then strUnitIDs = UserInfo.ID
    'by lesfeng 2010-1-8 �����Ż�
    gstrSQL = _
        " Select A.ID,A.����,A.����" & _
        " From ���ű� A,��������˵�� B" & IIf(blnLimitUnit, ",������Ա C ", "") & _
        " Where B.����ID = A.ID" & _
        " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        " And B.������� IN(1,2,3) And B.��������='����'" & _
        IIf(blnLimitUnit, " And A.ID = C.����ID And C.��ԱID In ([1])", "") & _
        " And (A.վ��=[2] Or A.վ�� is Null)" & _
        " Order by A.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(strUnitIDs), gstrNodeNo)
    
'    If blnLimitUnit Then strUnitIDs = GetUserUnits
'    'by lesfeng 2010-1-8 �����Ż�
'    gstrSQL = _
'        " Select A.ID,A.����,A.����" & _
'        " From ���ű� A,��������˵�� B" & _
'        " Where B.����ID = A.ID" & _
'        " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
'        " And B.������� IN(1,2,3) And B.��������='����'" & _
'        IIf(blnLimitUnit, " And A.ID In ([1])", "") & _
'        " And (A.վ��=[2] Or A.վ�� is Null)" & _
'        " Order by A.����"
'    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(strUnitIDs), gstrNodeNo)
    '����30922 by lesfeng 2010-06-18 e
    If Not rsTmp.EOF Then
        lngUnitID = UserInfo.����ID
        For i = 1 To rsTmp.RecordCount
            cboUnit.AddItem rsTmp!���� & "-" & rsTmp!����
            cboUnit.ItemData(cboUnit.NewIndex) = rsTmp!ID
            If rsTmp!ID = lngUnitID And cboUnit.ListIndex = -1 Then cboUnit.ListIndex = cboUnit.NewIndex
            rsTmp.MoveNext
        Next
        If cboUnit.ListIndex = -1 And cboUnit.ListCount > 0 Then cboUnit.ListIndex = 0
    ElseIf InStr(";" & mstrPrivs, "���в���") > 0 Then
        MsgBox "û�����ò���,�����ȵ����Ź��������ù�������Ϊ����Ĳ��ţ�", vbExclamation, gstrSysName
        Exit Function
    Else
        MsgBox "��û�� [���в���] ��Ȩ��,���������ڲ��Ų��ǲ�����", vbExclamation, gstrSysName
        Exit Function
    End If
    InitUnits = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ReadBeds(lngUnitID As Long) As Boolean
'���ܣ���ȡָ�������Ĵ�λ�б�
    Dim i As Integer, j As Integer
    Dim objItem As ListItem
    Dim intBedLen As Integer
    Dim mrsBeds As ADODB.Recordset
    
    On Error GoTo errH
    intBedLen = GetMaxBedLen(lngUnitID)
    gstrSQL = _
        " Select LPAD(A.����,[1],' ') ����,A.����ID," & _
        " A.�����,A.�Ա����,A.��λ����,A.�ȼ�ID,A.״̬,A.����ID,A.����," & _
        " Nvl(B.����,Decode(A.����,1,'<���ò���>',NULL)) as ����," & _
        " A.����ID,C.���� as �ȼ�,D.����,D.�Ա�,D.����" & _
        " From ��λ״����¼ A,���ű� B,�շ���ĿĿ¼ C,������Ϣ D" & _
        " Where A.����ID=B.ID(+) And A.�ȼ�ID=C.ID(+)" & _
        " And A.����ID=D.����ID(+) And A.����ID=[2] " & _
        " Order by LPAD(A.����,[1],' ')"
    Set mrsBeds = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, intBedLen, lngUnitID)
    
    lvwBeds.ListItems.Clear
    intHolding = 0: mintEmpty = 0: intRemedy = 0
    
    If Not mrsBeds.EOF Then
        For i = 1 To mrsBeds.RecordCount
            Select Case mrsBeds!״̬
                Case "�մ�"
                    If mrsBeds!�Ա���� = "�д�" Then
                        Set objItem = lvwBeds.ListItems.Add(, "_" & Trim(mrsBeds!����), mrsBeds!����, "M_Empty", "M_Empty")
                    ElseIf mrsBeds!�Ա���� = "Ů��" Then
                        Set objItem = lvwBeds.ListItems.Add(, "_" & Trim(mrsBeds!����), mrsBeds!����, "F_Empty", "F_Empty")
                    Else
                        Set objItem = lvwBeds.ListItems.Add(, "_" & Trim(mrsBeds!����), mrsBeds!����, "Empty", "Empty")
                    End If
                    mintEmpty = mintEmpty + 1
                Case "ռ��"
                    Set objItem = lvwBeds.ListItems.Add(, "_" & Trim(mrsBeds!����), mrsBeds!����, "Holding", "Holding")
                    intHolding = intHolding + 1
                Case "����"
                    Set objItem = lvwBeds.ListItems.Add(, "_" & Trim(mrsBeds!����), mrsBeds!����, "Remedy", "Remedy")
                    intRemedy = intRemedy + 1
                Case Else '��������
                    Set objItem = lvwBeds.ListItems.Add(, "_" & Trim(mrsBeds!����), mrsBeds!����, "Remedy", "Remedy")
                    mintEmpty = mintEmpty + 1
            End Select
            For j = 2 To lvwBeds.ColumnHeaders.Count
                objItem.SubItems(j - 1) = IIf(IsNull(mrsBeds.Fields(lvwBeds.ColumnHeaders(j).Text).Value), "", mrsBeds.Fields(lvwBeds.ColumnHeaders(j).Text).Value)
            Next
            objItem.Tag = IIf(IsNull(mrsBeds!����ID), 0, mrsBeds!����ID)
            objItem.ListSubItems(1).Tag = IIf(IsNull(mrsBeds!����), 0, mrsBeds!����) '��¼�Ƿ��ò���
            
            Call SetBedIcon(lvwBeds, objItem)
            
            mrsBeds.MoveNext
        Next
    End If
    Call SetBedNOLen
    ReadBeds = True
    sta.Panels(2) = "��ǰ������ " & lvwBeds.ListItems.Count & " �Ų���,���в���ռ�� " & intHolding & " ��,�մ� " & mintEmpty & " ��,�������� " & intRemedy & " �ţ�"
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SetView(bytStyle As Byte)
'���ܣ�������λ�б���ʾ��ʽ
'������bytstyle=0-��ͼ��,1-Сͼ��,2-�б�,3-��ϸ����
    mnuView_ListView(0).Checked = False
    mnuView_ListView(1).Checked = False
    mnuView_ListView(2).Checked = False
    mnuView_ListView(3).Checked = False
    mnuView_ListView(bytStyle).Checked = True
    lvwBeds.View = bytStyle
End Sub

Private Sub tbr_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
        Case "Icon"
            Call SetView(0)
        Case "Small"
            Call SetView(1)
        Case "List"
            Call SetView(2)
        Case "Detail"
            Call SetView(3)
    End Select
End Sub

Private Sub lvwBeds_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static blnDesc As Boolean
    Static intIdx As Integer
    
    If intIdx = ColumnHeader.Index Then
        blnDesc = Not blnDesc
    Else
        blnDesc = False
    End If
    lvwBeds.SortKey = ColumnHeader.Index - 1
    If blnDesc Then
        lvwBeds.SortOrder = lvwDescending
    Else
        lvwBeds.SortOrder = lvwAscending
    End If
    lvwBeds.Sorted = True
    intIdx = ColumnHeader.Index
    
    If Not lvwBeds.SelectedItem Is Nothing Then lvwBeds.SelectedItem.EnsureVisible
End Sub

Private Sub tbr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuViewTool, 2
End Sub

Private Sub mnuFile_Excel_Click()
    If lvwBeds.ListItems.Count > 100 Then
        If MsgBox("�����Excel�����ݹ���,�⽫�ķ����ʱ��,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    Call OutputList(3)
End Sub

Private Sub mnuFile_PreView_Click()
    Call OutputList(2)
End Sub

Private Sub mnuFile_Print_Click()
    Call OutputList(1)
End Sub

Private Sub mnuFile_PrintSet_Click()
    Call zlPrintSet
End Sub

Private Sub OutputList(bytStyle As Byte)
'���ܣ�������б�
'������bytStyle=1-��ӡ,2-Ԥ��,3-�����Excel
    Dim objOut As New zlPrintLvw
    Dim bytR As Byte
    
    On Error GoTo errH
    
    '��ͷ
    objOut.Title.Text = "סԺ�����嵥"
    objOut.Title.Font.Name = "����_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '����
    objOut.UnderAppItems.Add "����:" & NeedName(cboUnit.Text)
    objOut.BelowAppItems.Add "��ӡ�ˣ�" & UserInfo.����
    objOut.BelowAppItems.Add "��ӡ���ڣ�" & Format(zlDatabase.Currentdate(), "yyyy��MM��dd��")
    
    '����
    Set objOut.Body.objData = lvwBeds
    
    '���
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        If bytR <> 0 Then zlPrintOrViewLvw objOut, bytR
    Else
        zlPrintOrViewLvw objOut, bytStyle
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuHelpWebHome_Click()
    zlHomePage hwnd
End Sub

Private Sub mnuHelpWebMail_Click()
    zlMailTo hwnd
End Sub

Public Sub SetMenuState()
'���ܣ����ݲ�����λ���ȷ�����ܵ�ʹ��״̬
    If lvwBeds.SelectedItem Is Nothing Then
        mnuFile_Print.Enabled = False
        mnuFile_Preview.Enabled = False
        mnuFile_Excel.Enabled = False
        
        tbr.Buttons("Print").Enabled = False
        tbr.Buttons("Preview").Enabled = False
        
        mnuEdit_Modi.Enabled = False
        mnuEdit_Del.Enabled = False
        mnuEdit_Remedy.Enabled = False
        mnuEdit_Empty.Enabled = False
        
        tbr.Buttons("Modi").Enabled = False
        tbr.Buttons("Del").Enabled = False
        tbr.Buttons("Remedy").Enabled = False
        tbr.Buttons("Empty").Enabled = False
    Else
        mnuFile_Print.Enabled = True
        mnuFile_Preview.Enabled = True
        mnuFile_Excel.Enabled = True
        
        tbr.Buttons("Print").Enabled = True
        tbr.Buttons("Preview").Enabled = True
        
        Select Case lvwBeds.SelectedItem.SubItems(lvwBeds.ColumnHeaders("_״̬").Index - 1)
            Case "ռ��"
                mnuEdit_Modi.Enabled = False
                mnuEdit_Del.Enabled = False
                mnuEdit_Remedy.Enabled = False
                mnuEdit_Empty.Enabled = False
                tbr.Buttons("Modi").Enabled = False
                tbr.Buttons("Del").Enabled = False
                tbr.Buttons("Remedy").Enabled = False
                tbr.Buttons("Empty").Enabled = False
            Case "�մ�"
                mnuEdit_Modi.Enabled = True
                mnuEdit_Del.Enabled = True
                mnuEdit_Remedy.Enabled = True
                mnuEdit_Empty.Enabled = False
                tbr.Buttons("Modi").Enabled = True
                tbr.Buttons("Del").Enabled = True
                tbr.Buttons("Remedy").Enabled = True
                tbr.Buttons("Empty").Enabled = False
            Case "����"
                mnuEdit_Modi.Enabled = False
                mnuEdit_Del.Enabled = False
                mnuEdit_Remedy.Enabled = False
                mnuEdit_Empty.Enabled = True
                tbr.Buttons("Modi").Enabled = False
                tbr.Buttons("Del").Enabled = False
                tbr.Buttons("Remedy").Enabled = False
                tbr.Buttons("Empty").Enabled = True
        End Select
    End If
End Sub

Public Sub SetBedNOLen()
    Dim bytLen As Byte, i As Integer
    
    If lvwBeds.ListItems.Count = 0 Then Exit Sub
    
    bytLen = GetMaxBedLen(cboUnit.ItemData(cboUnit.ListIndex))
    
    For i = 1 To lvwBeds.ListItems.Count
        lvwBeds.ListItems(i).Text = Space(bytLen - Len(CStr(Trim(lvwBeds.ListItems(i).Text)))) & Trim(lvwBeds.ListItems(i).Text)
    Next
End Sub

Private Sub MakeBedIcon()
    Dim i As Integer, k As Integer
    
    k = img32.ListImages.Count
    For i = 1 To img32.ListImages.Count
        If Not img32.ListImages(i).Key Like "MASK_*" Then
            img32.ListImages.Add , "�Ӵ�_" & img32.ListImages(i).Key, img32.Overlay("MASK_�Ӵ�", i)
            img32.ListImages.Add , "�Ǳ�_" & img32.ListImages(i).Key, img32.Overlay("MASK_�Ǳ�", i)
            img32.ListImages.Add , "����_" & img32.ListImages(i).Key, img32.Overlay("MASK_����", i)
            img32.ListImages.Add , "����_�Ӵ�_" & img32.ListImages(i).Key, img32.Overlay("MASK_����_�Ӵ�", i)
            img32.ListImages.Add , "����_�Ǳ�_" & img32.ListImages(i).Key, img32.Overlay("MASK_����_�Ǳ�", i)
        End If
    Next
    
    k = img16.ListImages.Count
    For i = 1 To img16.ListImages.Count
        If Not img16.ListImages(i).Key Like "MASK_*" Then
            img16.ListImages.Add , "�Ӵ�_" & img16.ListImages(i).Key, img16.Overlay("MASK_�Ӵ�", i)
            img16.ListImages.Add , "�Ǳ�_" & img16.ListImages(i).Key, img16.Overlay("MASK_�Ǳ�", i)
            img16.ListImages.Add , "����_" & img16.ListImages(i).Key, img16.Overlay("MASK_����", i)
            img16.ListImages.Add , "����_�Ӵ�_" & img16.ListImages(i).Key, img16.Overlay("MASK_����_�Ӵ�", i)
            img16.ListImages.Add , "����_�Ǳ�_" & img16.ListImages(i).Key, img16.Overlay("MASK_����_�Ǳ�", i)
        End If
    Next
End Sub

Private Sub SetBedIcon(objLvw As Object, objItem As ListItem)
    If objItem.SubItems(objLvw.ColumnHeaders("_��λ����").Index - 1) = "�Ӵ�" Then
        objItem.Icon = "�Ӵ�_" & objItem.Icon
        objItem.SmallIcon = "�Ӵ�_" & objItem.SmallIcon
    ElseIf objItem.SubItems(objLvw.ColumnHeaders("_��λ����").Index - 1) = "�Ǳ�" Then
        objItem.Icon = "�Ǳ�_" & objItem.Icon
        objItem.SmallIcon = "�Ǳ�_" & objItem.SmallIcon
    End If
    
    If Val(objItem.ListSubItems(1).Tag) <> 0 Then
        objItem.Icon = "����_" & objItem.Icon
        objItem.SmallIcon = "����_" & objItem.SmallIcon
    End If
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

