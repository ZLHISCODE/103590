VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmPresManage 
   Caption         =   "��Ա����"
   ClientHeight    =   6960
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   9645
   Icon            =   "frmPresManage.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   9645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picPres 
      BackColor       =   &H80000005&
      Height          =   1950
      Left            =   2895
      ScaleHeight     =   1890
      ScaleWidth      =   6435
      TabIndex        =   14
      Top             =   4050
      Width           =   6495
      Begin VB.PictureBox pic˵�� 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1365
         Left            =   3930
         MousePointer    =   9  'Size W E
         ScaleHeight     =   1365
         ScaleWidth      =   45
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   420
         Width           =   45
      End
      Begin VB.PictureBox pic��Ƭ 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1395
         Left            =   1815
         MousePointer    =   9  'Size W E
         ScaleHeight     =   1395
         ScaleWidth      =   45
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   390
         Width           =   45
      End
      Begin VB.PictureBox pic���� 
         BackColor       =   &H00FFFFFF&
         Height          =   1440
         Left            =   1980
         ScaleHeight     =   1380
         ScaleWidth      =   1830
         TabIndex        =   16
         Top             =   360
         Width           =   1890
         Begin VB.Image img��Ƭ 
            Height          =   1035
            Left            =   90
            Stretch         =   -1  'True
            Top             =   240
            Width           =   945
         End
      End
      Begin VB.TextBox txt˵�� 
         Height          =   1455
         Left            =   4065
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   360
         Width           =   2325
      End
      Begin MSComctlLib.ListView lvw��Ա����_S 
         Height          =   1440
         Left            =   30
         TabIndex        =   19
         Top             =   330
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   2540
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "��Ա����"
            Object.Tag             =   "��Ա����"
            Text            =   "��Ա����"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "˵��"
            Object.Tag             =   "˵��"
            Text            =   "˵��"
            Object.Width           =   14111
         EndProperty
      End
      Begin VB.Label lbl���� 
         BackColor       =   &H00808080&
         Caption         =   " ���˼��"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   2
         Left            =   4065
         TabIndex        =   22
         Top             =   45
         Width           =   2325
      End
      Begin VB.Label lbl���� 
         BackColor       =   &H00808080&
         Caption         =   " ��Ƭ"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   1
         Left            =   1965
         TabIndex        =   21
         Top             =   45
         Width           =   1905
      End
      Begin VB.Label lbl���� 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   " ��������"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   0
         Left            =   75
         TabIndex        =   20
         Top             =   75
         Width           =   1680
      End
   End
   Begin VB.PictureBox pic֤�� 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   6360
      MousePointer    =   9  'Size W E
      ScaleHeight     =   645
      ScaleWidth      =   45
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2880
      Width           =   45
   End
   Begin MSComctlLib.ImageList Ilsrw 
      Left            =   2835
      Top             =   6045
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
            Picture         =   "frmPresManage.frx":030A
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":052A
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":074A
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":096A
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":0B8A
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":0DAA
            Key             =   "View"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":0FC6
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":11E0
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":1400
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":1620
            Key             =   "start"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":183A
            Key             =   "stop"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":1A54
            Key             =   "sign"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList Ilscolor 
      Left            =   3810
      Top             =   6045
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
            Picture         =   "frmPresManage.frx":232E
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":254E
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":276E
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":298E
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":2BAE
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":2DCE
            Key             =   "View"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":2FEA
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":3204
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":3424
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":3644
            Key             =   "start"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":385E
            Key             =   "stop"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":3A78
            Key             =   "sign"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   9645
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "Toolbar1"
      MinHeight1      =   720
      Width1          =   8370
      Key1            =   "only"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   720
         Left            =   30
         TabIndex        =   6
         Top             =   30
         Width           =   9525
         _ExtentX        =   16801
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "Ilsrw"
         HotImageList    =   "Ilscolor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   20
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ԥ��"
               Key             =   "Preview"
               Object.ToolTipText     =   "Ԥ��"
               Object.Tag             =   "Ԥ��"
               ImageKey        =   "Preview"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "Print"
               Object.ToolTipText     =   "��ӡ"
               Object.Tag             =   "��ӡ"
               ImageKey        =   "Print"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "New"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageKey        =   "New"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�޸�"
               Key             =   "Modify"
               Object.ToolTipText     =   "�޸�"
               Object.Tag             =   "�޸�"
               ImageKey        =   "Modify"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ɾ��"
               Key             =   "Delete"
               Object.ToolTipText     =   "ɾ��"
               Object.Tag             =   "ɾ��"
               ImageKey        =   "Delete"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "sdf"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Start"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageKey        =   "start"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ͣ��"
               Key             =   "Stop"
               Object.ToolTipText     =   "ͣ��"
               Object.Tag             =   "ͣ��"
               ImageKey        =   "stop"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "SignLine"
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "֤��"
               Key             =   "Sign"
               Object.ToolTipText     =   "����֤����ͣ��"
               Object.Tag             =   "֤��"
               ImageKey        =   "sign"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "SignOn"
                     Object.Tag             =   "����֤������"
                     Text            =   "����֤������"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "SignOff"
                     Object.Tag             =   "����֤��ͣ��"
                     Text            =   "����֤��ͣ��"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "sgf1"
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�鿴"
               Key             =   "View"
               Object.ToolTipText     =   "��Ա�鿴��ʽ"
               Object.Tag             =   "�鿴"
               ImageKey        =   "View"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "  ��ͼ��"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "  Сͼ��"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "  �б�"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "  ��ϸ����"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Find"
               Description     =   "����"
               Object.ToolTipText     =   "��Ա����"
               Object.Tag             =   "����"
               ImageKey        =   "Find"
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��չ"
               Key             =   "plugIn"
               Object.ToolTipText     =   "��չ����"
               Object.Tag             =   "��չ"
               ImageKey        =   "View"
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "plugInS"
               Style           =   3
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Quit"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
         Begin VB.TextBox txtFind 
            Height          =   300
            Left            =   8040
            MaxLength       =   10
            TabIndex        =   11
            Tag             =   "����"
            Top             =   120
            Width           =   1425
         End
         Begin VB.PictureBox picFind 
            BorderStyle     =   0  'None
            Height          =   300
            Left            =   7440
            ScaleHeight     =   285.714
            ScaleMode       =   0  'User
            ScaleWidth      =   495
            TabIndex        =   9
            Top             =   120
            Width           =   495
            Begin VB.Label lbl���� 
               Caption         =   "����"
               Height          =   255
               Left            =   120
               TabIndex        =   10
               Top             =   75
               Width           =   495
            End
         End
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   7
      Top             =   6600
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   635
      SimpleText      =   $"frmPresManage.frx":4752
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPresManage.frx":4799
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11933
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
   Begin MSComctlLib.ListView lvwMain 
      Height          =   1695
      Left            =   2835
      TabIndex        =   1
      Top             =   780
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   2990
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ils32"
      SmallIcons      =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.PictureBox picSplitV 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4830
      Left            =   2790
      MousePointer    =   9  'Size W E
      ScaleHeight     =   4830
      ScaleWidth      =   45
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   780
      Width           =   45
   End
   Begin VB.PictureBox picSplitH 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   2970
      MousePointer    =   7  'Size N S
      ScaleHeight     =   33.75
      ScaleMode       =   0  'User
      ScaleWidth      =   6165
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3945
      Width           =   6165
   End
   Begin MSComctlLib.ImageList ils32 
      Left            =   900
      Top             =   1950
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":502D
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":5349
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":5995
            Key             =   "Dept"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":5CB1
            Key             =   "Dept_No"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":5FD1
            Key             =   "Item_G"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":C833
            Key             =   "Item_W"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   1575
      Top             =   1980
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":13095
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":136E1
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":139FD
            Key             =   "Dept"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":13D19
            Key             =   "Dept_No"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":14039
            Key             =   "Cert"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":14193
            Key             =   "Item_G"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":1A9F5
            Key             =   "Item_W"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":21257
            Key             =   "SignOn"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":217A9
            Key             =   "SignOff"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwMain_S 
      Height          =   4815
      Left            =   60
      TabIndex        =   0
      Top             =   780
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   8493
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "ils16"
      Appearance      =   1
   End
   Begin MSComctlLib.ListView lvwCert 
      Height          =   1155
      Left            =   2835
      TabIndex        =   2
      Top             =   2775
      Visible         =   0   'False
      Width           =   3420
      _ExtentX        =   6033
      _ExtentY        =   2037
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ע��ʱ��"
         Object.Width           =   3651
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "���к�"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ʹ����"
         Object.Width           =   6174
      EndProperty
   End
   Begin MSComctlLib.ListView lvwLogOnOff 
      Height          =   1155
      Left            =   6480
      TabIndex        =   12
      Top             =   2760
      Visible         =   0   'False
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   2037
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "StopTime"
         Text            =   "ͣ��ʱ��"
         Object.Width           =   3651
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "StartTime"
         Text            =   "����ʱ��"
         Object.Width           =   3651
      EndProperty
   End
   Begin XtremeSuiteControls.TabControl tbcPres 
      Height          =   645
      Left            =   1065
      TabIndex        =   23
      Top             =   5805
      Width           =   1065
      _Version        =   589884
      _ExtentX        =   1879
      _ExtentY        =   1138
      _StockProps     =   64
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   " ����֤��"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   3
      Left            =   2850
      MousePointer    =   7  'Size N S
      TabIndex        =   8
      Top             =   2550
      Visible         =   0   'False
      Width           =   6285
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFileSet 
         Caption         =   "��ӡ����(&S)"
      End
      Begin VB.Menu mnuFilePreview 
         Caption         =   "��ӡԤ��(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "��ӡ(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "�����&Excel"
      End
      Begin VB.Menu mnuSplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileReport 
         Caption         =   "��������(&R)"
      End
      Begin VB.Menu mnuFileFile 
         Caption         =   "���������ļ�(&F)"
      End
      Begin VB.Menu mnusplit2 
         Caption         =   "-"
      End
      Begin VB.Menu mnufileexit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuEditNew 
         Caption         =   "������Ա��Ϣ(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "�޸���Ա��Ϣ(&M)"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "ɾ����Ա��Ϣ(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuPlugIn 
         Caption         =   "��չ(&E)"
         Begin VB.Menu mnuPlugItem 
            Caption         =   "����"
            Index           =   0
         End
      End
      Begin VB.Menu mnuEdit_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditAdjust 
         Caption         =   "��Ա���ŵ���(&J)"
      End
      Begin VB.Menu mnuEditRole 
         Caption         =   "��Ա��ɫ����(&O)"
      End
      Begin VB.Menu mnuEditDeptRole 
         Caption         =   "������ɫ����"
      End
      Begin VB.Menu mnuEditExtend 
         Caption         =   "��չ��Ϣά��(&E)"
      End
      Begin VB.Menu mnuEditSplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditStart 
         Caption         =   "����(&S)"
      End
      Begin VB.Menu mnuEditStop 
         Caption         =   "ͣ��(&T)"
      End
      Begin VB.Menu mnuEdit_1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditRegCert 
         Caption         =   "����֤��ע��(&R)"
         Shortcut        =   {F2}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditViewCert 
         Caption         =   "�鿴����֤��(&V)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditImportCertPic 
         Caption         =   "����ǩ��ͼƬ(&P)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEdit_2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditDelCert 
         Caption         =   "ȡ��֤��ע��"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditSignLine1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditSignOn 
         Caption         =   "����֤������(&B)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditSignOff 
         Caption         =   "����֤��ͣ��(&E)"
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
         Begin VB.Menu mnuViewToolspilt1 
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
      Begin VB.Menu mnuviewsplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "��ͼ��(&G)"
         Index           =   0
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "Сͼ��(&M)"
         Index           =   1
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "�б�(&L)"
         Index           =   2
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "��ϸ����(&D)"
         Checked         =   -1  'True
         Index           =   3
      End
      Begin VB.Menu mnuViewSplit5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewStretch 
         Caption         =   "��Ƭ�Զ�����(&E)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewSplit3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "��Ա����(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewColumn 
         Caption         =   "ѡ����(&C)"
      End
      Begin VB.Menu mnuViewSplit6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewShowStopDept 
         Caption         =   "��ʾͣ�ò���(&E)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewSplit4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewShowStop 
         Caption         =   "��ʾͣ����Ա(&P)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewShow 
         Caption         =   "ֻ��ʾֱ����Ա(&H)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewReflash 
         Caption         =   "ˢ��(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuHelpTopic 
         Caption         =   "��������(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "Web�ϵ�����"
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
      Begin VB.Menu mnuHelpSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)..."
      End
   End
   Begin VB.Menu mnuShort 
      Caption         =   "��ݲ˵�"
      Visible         =   0   'False
      Begin VB.Menu mnuShortMenu 
         Caption         =   "����(&A)"
         Index           =   1
      End
      Begin VB.Menu mnuShortMenu 
         Caption         =   "�޸�(&M)"
         Index           =   2
      End
      Begin VB.Menu mnuShortMenu 
         Caption         =   "ɾ��(&D)"
         Index           =   3
      End
      Begin VB.Menu mnuShortMenu 
         Caption         =   "��չ(&E)"
         Index           =   4
         Begin VB.Menu mnuShortPlugInItem 
            Caption         =   "����"
            Index           =   0
         End
      End
      Begin VB.Menu mnuShortMenu 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuShortMenu 
         Caption         =   "��Ա���ŵ���(&J)"
         Index           =   6
      End
      Begin VB.Menu mnuShortMenu 
         Caption         =   "��Ա��ɫ����(&O)"
         Index           =   7
      End
      Begin VB.Menu mnuShortsplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShortIcon 
         Caption         =   "��ͼ��(&G)"
         Index           =   0
      End
      Begin VB.Menu mnuShortIcon 
         Caption         =   "Сͼ��(&M)"
         Index           =   1
      End
      Begin VB.Menu mnuShortIcon 
         Caption         =   "�б�(&L)"
         Index           =   2
      End
      Begin VB.Menu mnuShortIcon 
         Caption         =   "��ϸ����(&D)"
         Checked         =   -1  'True
         Index           =   3
      End
      Begin VB.Menu mnuShortSign 
         Caption         =   "����֤������(&B)"
         Index           =   0
      End
      Begin VB.Menu mnuShortSign 
         Caption         =   "����֤��ͣ��(&E)"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmPresManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private msngStartX As Single, msngStartY As Single    '�ƶ�ǰ����λ��
Private mblnItem As Boolean  'Ϊ���ʾ������ListViewĳһ����
Private mintColumn As Integer
Private mblnLoad As Boolean
Private mstrKey As String    '��һ�ε�Node�ؼ�ֵ
Private Const mstrLvw As String = "����,2000,0,1;���,800,0,2;����ְ��,1400,0,0;רҵ����ְ��,1400,0,0;סԺ����ҩ��Ȩ��,1400,0,0;���￹��ҩ��Ȩ��,1400,0,0;" & _
                                  "�����ȼ�,1000,0,0;Ƹ�μ���ְ��,1400,0,0;��������,1200,0,0;���֤��,1800,0,0;�Ա�,800,0,0;����,800,0,0;ѧ��,800,0,0;" & _
                                  "�칫�ҵ绰,1400,0,0;�ƶ��绰,1400,0,0;�����ʼ�,1400,0,0;����,600,0,0;��������,2000,0,0;����ʱ��,1440,0,0;����ʱ��,1440,0,0"

Private mobjESign As Object                 '����ǩ���ӿ�
Private mintCA As Integer                   '����ǩ����֤����
Private mlngMode As Long
Public mstrPrivs As String                  'Ȩ�޴�
Private Declare Function SetParent Lib "user32 " (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private mrsFind As ADODB.Recordset          '�������ѯ
Private mstrFindValue As String             '��¼��ѯ�ı����ֵ
Private mrsPersonProper As ADODB.Recordset  '��������
Private mblnCAOnOff As Boolean              '����֤����ͣ��Ȩ��
Private mobjForm As frmDeptExtend
Private mblnPACSInterface As Boolean        '����Ӱ����Ϣϵͳ�ӿ�
Private Sub Form_Activate()
    If mblnLoad = True Then
        Call Form_Resize 'Ϊ����ȷ����coolbar�ĸ߶�
        If InStr(mstrPrivs, "���в���") = 0 Then
            Call FillTreePrivs
        Else
            If FillTree = False Then
                Unload Me
            End If
        End If
    End If
    mblnLoad = False
End Sub

Private Sub Form_Load()
    Dim rsTemp As ADODB.Recordset
    
    mblnLoad = True
        
    mlngMode = glngModul
    mstrPrivs = gstrPrivs
    Call InitTabControl
    
    mblnCAOnOff = InStr(mstrPrivs, ";����֤����ͣ��;") > 0
    mblnPACSInterface = (Val(zlDatabase.GetPara(255, glngSys, , "0")) = 1)
    
    '���������ɾ����ListView�������
    lvwMain.Tag = "�ɱ仯��"
    '-----------
    RestoreWinState Me, App.ProductName
    lvw��Ա����_S.Visible = True
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngMode, mstrPrivs, "ZL3_INSIDE_222_1")
    
    mnuViewShow.Checked = (Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "ֻ��ʾֱ����Ա", 0)) = 1)
    mnuViewStretch.Checked = (Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "��Ƭ�Զ�����", 1)) = 1)
    mnuViewShowStop.Checked = (Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "��ʾͣ��", 0)) = 1)
    
    If InStr(1, mstrPrivs, ";�޸�ʱ���޶���Ա����;") = 0 Then
        gstrSQL = "Select 1 From ��Ա����˵�� Where ��Աid =  [1] and rownum>=1"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����Ա�Ƿ�����Ȩ��", glngUserId)
        
        If rsTemp.RecordCount > 0 Then
            gstrSQL = "Select Distinct ��Աid From ��Ա����˵�� Where ��Ա���� In (Select ��Ա���� From ��Ա����˵�� Where ��Աid = [1])"
        Else
            '��ѯ�͵�ǰ����Ա������ͬ���ʵ���Ա
            gstrSQL = "Select ID As ��Աid" & vbNewLine & _
                "From ��Ա��" & vbNewLine & _
                "Where ID Not In (Select Distinct ��Աid From ��Ա����˵��)"
        End If
        Set mrsPersonProper = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯ����Ա��������", glngUserId)
    End If
    
    Call Set��Ƭ����
    Call SetȨ�޿��� '�����Ե���ǩ���ӿڵĳ�ʼ�����ж�
    
    '���ListView�Ļ�δ�����ã������һ��ʹ�ã��Ǿ͵���ȱʡ�ĳ�ʼ��
'    If lvwMain.ColumnHeaders.Count = 0 Then
        zlControl.LvwSelectColumns lvwMain, mstrLvw, True
'    End If
    
    If gobjPlugIn Is Nothing Then
        On Error Resume Next
        Set gobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
        If Not gobjPlugIn Is Nothing Then
            Call gobjPlugIn.Initialize(gcnOracle, glngSys, glngModul)
            If InStr(",438,0,", "," & Err.Number & ",") = 0 Then
                MsgBox "zlPlugIn ��Ҳ���ִ�� Initialize ʱ����" & vbCrLf & Err.Number & vbCrLf & Err.Description, vbInformation, gstrSysName
            End If
        End If
        Err.Clear: On Error GoTo 0
    End If
     
    Call LoadPlugInMnu(Not gobjPlugIn Is Nothing)
    
    '����LvwMain��ʾ���ö�Ӧ�˵�
     mnuViewIcon_Click lvwMain.View
     
    '��ʼ������RIS�ӿ�
    If mblnPACSInterface Then
        Call IniRIS
    End If
End Sub

Private Sub InitTabControl()
    '��ʼ��Tabcontrol�ؼ�
    With Me.tbcPres
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = False
            .ShowIcons = True
        End With
        
        Set mobjForm = New frmDeptExtend
        Call SetFormVisible(mobjForm.hwnd) '�����������С������

        .InsertItem(0, "��Ա����", picPres.hwnd, 0).Tag = "��Ա����"
        .InsertItem(1, "��չ��Ϣ", mobjForm.hwnd, 0).Tag = "��չ��Ϣ"
        
        .Item(1).Selected = True
        .Item(0).Selected = True
    End With
End Sub

Private Sub LoadPlugInMnu(ByVal blnHave As Boolean)
'������blnHave true ��ʾ����������
    Dim strTmp As String
    Dim arrTmp As Variant
    Dim i As Integer
    
    mnuPlugIn.Visible = blnHave
    mnuShortMenu(4).Visible = blnHave
    Toolbar1.Buttons("plugIn").Visible = blnHave
    Toolbar1.Buttons("plugInS").Visible = blnHave
 
    If blnHave Then
        'blnHave Ϊtrue ʱ����ȷ�� gobjPlugIn ����Ϊ Nothing
        On Error Resume Next
        strTmp = gobjPlugIn.GetFuncNames(glngSys, glngModul)
        If InStr(",438,0,", "," & Err.Number & ",") = 0 Then
            MsgBox "zlPlugIn ��Ҳ���ִ�� GetFuncNames ʱ����" & vbCrLf & Err.Number & vbCrLf & Err.Description, vbInformation, gstrSysName
        End If
        Err.Clear: On Error GoTo 0
        
        If strTmp = "" Then Exit Sub
        
        strTmp = Replace(strTmp, "Auto:", "")
        arrTmp = Split(strTmp, ",")
        For i = 0 To UBound(arrTmp)
            If i <> 0 Then
                Load mnuPlugItem(i)
                Load mnuShortPlugInItem(i)
            End If
            
            mnuPlugItem(i).Caption = CStr(arrTmp(i))
            mnuPlugItem(i).Tag = CStr(arrTmp(i))
            mnuShortPlugInItem(i).Caption = CStr(arrTmp(i))
            mnuShortPlugInItem(i).Tag = CStr(arrTmp(i))
            
            If i <= 9 Then
                mnuPlugItem(i).Caption = CStr(arrTmp(i)) & "(&" & IIF(i = 9, 0, i + 1) & ")"
                mnuShortPlugInItem(i).Caption = mnuPlugItem(i).Caption
            End If
        Next
    End If
End Sub

Private Sub lvwCert_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If lvwCert.ListItems.Count <= 0 Then Exit Sub
    
    Dim lngId As Long
    
    lngId = Val(Mid(lvwCert.ListItems(lvwCert.SelectedItem.Index).Key, 2))
    Call FillLogOnOff(lngId)
End Sub

Private Sub lvwCert_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        '���õ����˵�
        If lvwCert.ListItems.Count <= 0 Then Exit Sub
        
        Dim i As Integer
        
        '���������˵���
        For i = mnuShortIcon.LBound To mnuShortIcon.UBound
            mnuShortIcon(i).Visible = False
        Next
        For i = mnuShortMenu.LBound To mnuShortMenu.UBound
            mnuShortMenu(i).Visible = False
        Next
        For i = mnuShortSign.LBound To mnuShortSign.UBound
            mnuShortSign(i).Visible = True
        Next
        mnuShortsplit1.Visible = False
        
        mnuShortSign(0).Enabled = mnuEditSignOn.Enabled
        mnuShortSign(1).Enabled = mnuEditSignOff.Enabled
        
        '�����˵�
        PopupMenu mnuShort
        
        '�ָ������˵���
        For i = mnuShortIcon.LBound To mnuShortIcon.UBound
            mnuShortIcon(i).Visible = True
        Next
        For i = mnuShortMenu.LBound To mnuShortMenu.UBound
            mnuShortMenu(i).Visible = True
        Next
        For i = mnuShortSign.LBound To mnuShortSign.UBound
            mnuShortSign(i).Visible = False
        Next
        mnuShortsplit1.Visible = True
    End If
End Sub

Private Sub mnuEditExtend_Click()
    Dim strKey As String
    Dim strName As String
    
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    With lvwMain.SelectedItem
        strKey = Mid(.Key, 2)
        strName = .Text
    End With
    
    Call frmDeptExtend.ShowMe(Me, strKey, strName, 1, 1)
    Call mobjForm.initVSf(Val(strKey), 1)
End Sub

Private Sub mnuEditImportCertPic_Click()
    Dim arrData As Variant
    
    On Error GoTo errH
    
    If Not mobjESign Is Nothing Then
        If mobjESign.RegisterCertificate(arrData) Then
            If arrData(0) <> lvwMain.SelectedItem.Text Then
                If MsgBox("������֤���ǰ䷢��""" & arrData(0) & """������ǰע����ԱΪ""" & lvwMain.SelectedItem.Text & """��Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            End If
            
            '����ǩ��ͼƬ
            If UBound(arrData) > 4 Then
                If arrData(5) <> "" Then
                    If SaveSignPIC(Mid(lvwMain.SelectedItem.Key, 2), arrData(5)) = False Then
                        GoTo errH
                    End If
                End If
            End If
            
            Call ShowAttribe
            Call SetMenu
            
            MsgBox lvwMain.SelectedItem.Text & "��ǩ��ͼƬ���³ɹ���", vbInformation, gstrSysName
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditSignOff_Click()
    Dim lngId As Long
    Dim strSQL As String
    
    If lvwCert.ListItems.Count <= 0 Then Exit Sub
    If lvwCert.SelectedItem.Index <= 0 Then
        MsgBox "��ѡ��һ������֤�飡", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If MsgBox("�Ƿ�ȷ����ͣ������ǩ����������", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    
    lngId = Val(Mid(lvwCert.SelectedItem.Key, 2))
    
    strSQL = "Zl_��Ա֤���¼_Esignswitch(" & lngId & ",1,null)"
    Call zlDatabase.ExecuteProcedure(strSQL, "ͣ������ǩ��")
    
    'Call FillLogOnOff(lngId)
    lngId = lvwCert.SelectedItem.Index
    Call ShowAttribe
    lvwCert.ListItems(lngId).Selected = True
    Call lvwCert_ItemClick(lvwCert.SelectedItem)
End Sub

Private Sub mnuEditSignOn_Click()
    Dim lngId As Long
    Dim strSQL As String, strStop As String
    
    If lvwCert.ListItems.Count <= 0 Then Exit Sub
    If lvwCert.SelectedItem.Index <= 0 Then
        MsgBox "��ѡ��һ������֤�飡", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If MsgBox("�Ƿ�ȷ������������ǩ����������", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    
    lngId = Val(Mid(lvwCert.SelectedItem.Key, 2))
    strStop = lvwLogOnOff.ListItems(1).Text
    
    strSQL = "Zl_��Ա֤���¼_Esignswitch(" & lngId & ",0,to_date('" & strStop & "', 'yyyy-mm-dd hh24:mi:ss'))"
    Call zlDatabase.ExecuteProcedure(strSQL, "��������ǩ��")
    
    'Call FillLogOnOff(lngID)
    lngId = lvwCert.SelectedItem.Index
    Call ShowAttribe
    lvwCert.ListItems(lngId).Selected = True
    Call lvwCert_ItemClick(lvwCert.SelectedItem)
End Sub

Private Sub mnuPlugItem_Click(Index As Integer)
    Call ExcPlugInFun(mnuPlugItem(Index).Tag)
End Sub

Private Sub mnuShortPlugInItem_Click(Index As Integer)
    Call ExcPlugInFun(mnuShortPlugInItem(Index).Tag)
End Sub

Private Sub ExcPlugInFun(ByVal strFunName As String)
    Dim lng��Աid As Long
    
    If Not lvwMain.SelectedItem Is Nothing Then
        With lvwMain.SelectedItem
            lng��Աid = Val(Mid(.Key, 2))
        End With
    End If
    
    On Error Resume Next
    
    If gobjPlugIn Is Nothing Then
        On Error Resume Next
        Set gobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
        If Not gobjPlugIn Is Nothing Then
            Call gobjPlugIn.Initialize(gcnOracle, glngSys, glngModul)
            If InStr(",438,0,", "," & Err.Number & ",") = 0 Then
                MsgBox "zlPlugIn ��Ҳ���ִ�� Initialize ʱ����" & vbCrLf & Err.Number & vbCrLf & Err.Description, vbInformation, gstrSysName
            End If
        End If
        Err.Clear: On Error GoTo 0
    End If
    
    If Not gobjPlugIn Is Nothing Then
        On Error Resume Next
        Call gobjPlugIn.ExecuteFunc(glngSys, glngModul, strFunName, lng��Աid, 0, 0)
        If InStr(",438,0,", "," & Err.Number & ",") = 0 Then
            MsgBox "zlPlugIn ��Ҳ���ִ�� ExecuteFunc ʱ����" & vbCrLf & Err.Number & vbCrLf & Err.Description, vbInformation, gstrSysName
        End If
        Err.Clear: On Error GoTo 0
    End If
End Sub

Private Sub Form_Resize()
    Dim cbrH As Long, staH As Long
    Dim lngCert As Long
    
    On Error Resume Next
    
    If WindowState = 1 Then Exit Sub
    cbrH = IIF(CoolBar1.Visible, CoolBar1.Height, 0)
    staH = IIF(stbThis.Visible, stbThis.Height, 0)
    lngCert = IIF(lvwCert.Visible, lvwCert.Height + lbl����(3).Height + 30, 0)
    
    tvwMain_S.Left = 0
    tvwMain_S.Top = cbrH
    tvwMain_S.Height = Me.ScaleHeight - cbrH - staH
    
    picSplitV.Top = tvwMain_S.Top
    picSplitV.Left = tvwMain_S.Left + tvwMain_S.Width
    picSplitV.Height = tvwMain_S.Height
    
    lvwMain.Top = cbrH
    lvwMain.Left = picSplitV.Left + picSplitV.Width
    lvwMain.Width = Me.ScaleWidth - picSplitV.Left - picSplitV.Width
    lvwMain.Height = Me.ScaleHeight - cbrH - staH - lngCert - picSplitH.Height - 2000
    
    lbl����(3).Top = lvwMain.Top + lvwMain.Height + 15
    lbl����(3).Left = lvwMain.Left + 15
    lbl����(3).Width = lvwMain.Width - 30
    
    lvwLogOnOff.Top = lbl����(3).Top + lbl����(3).Height + 15
    If mblnLoad Then lvwLogOnOff.Width = lvwMain.Width \ 2
    lvwLogOnOff.Left = lvwMain.Left + lvwMain.Width - lvwLogOnOff.Width - pic֤��.Width + 30
    lvwLogOnOff.Height = lvwCert.Height
    
    pic֤��.Top = lvwLogOnOff.Top
    pic֤��.Left = lvwLogOnOff.Left - pic֤��.Width
    pic֤��.Height = lvwCert.Height
    
    lvwCert.Left = lvwMain.Left
    lvwCert.Top = lbl����(3).Top + lbl����(3).Height + 15
    lvwCert.Width = lvwMain.Width - lvwLogOnOff.Width - pic֤��.Width
    
    picSplitH.Left = lvwMain.Left
    picSplitH.Top = lvwMain.Top + lvwMain.Height + lngCert
    picSplitH.Width = lvwMain.Width
    
    tbcPres.Move picSplitV.Left + picSplitV.Width + 30, picSplitH.Top + picSplitH.Height, Me.ScaleWidth - picSplitV.Left - picSplitV.Width - 30, Me.ScaleHeight - picSplitH.Top - picSplitH.Height - staH
    picPres.Move 0, 360, tbcPres.Width, tbcPres.Height - 360
    
    lbl����(0).Left = 0
    lbl����(0).Top = 0
    lbl����(0).Width = 1800
    lvw��Ա����_S.Left = 0
    lvw��Ա����_S.Top = lbl����(0).Top + lbl����(0).Height + 15
    lvw��Ա����_S.Width = lbl����(0).Width
    lvw��Ա����_S.Height = picPres.ScaleHeight - lvw��Ա����_S.Top
    
    pic��Ƭ.Left = lvw��Ա����_S.Left + lvw��Ա����_S.Width
    pic��Ƭ.Top = lbl����(0).Top
    pic��Ƭ.Height = picPres.ScaleHeight
    
    lbl����(1).Top = lbl����(0).Top
    lbl����(1).Left = pic��Ƭ.Left + pic��Ƭ.Width
    lbl����(1).Width = 1800
    pic����.Left = lbl����(1).Left
    pic����.Top = lvw��Ա����_S.Top
    pic����.Width = lbl����(1).Width
    pic����.Height = lvw��Ա����_S.Height
    
    pic˵��.Top = pic��Ƭ.Top
    pic˵��.Left = pic����.Left + pic����.Width
    pic˵��.Height = pic��Ƭ.Height
    
    lbl����(2).Top = lbl����(0).Top
    lbl����(2).Left = pic˵��.Left + pic˵��.Width
    lbl����(2).Width = picPres.ScaleWidth - lbl����(2).Left
    
    txt˵��.Left = lbl����(2).Left
    txt˵��.Top = lvw��Ա����_S.Top
    txt˵��.Height = lvw��Ա����_S.Height
    txt˵��.Width = lbl����(2).Width
    
    SetParent txtFind.hwnd, Toolbar1.hwnd
    SetParent picFind.hwnd, Toolbar1.hwnd
    txtFind.Left = Me.Width - txtFind.Width
    picFind.Left = txtFind.Left - 100 - picFind.Width
    
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "ֻ��ʾֱ����Ա", IIF(mnuViewShow.Checked, 1, 0)
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "��ʾͣ��", IIF(mnuViewShowStop.Checked, 1, 0)
        
    SaveWinState Me, App.ProductName
    
    Set mobjESign = Nothing
    If Not mobjForm Is Nothing Then Set mobjForm = Nothing
    
End Sub

Private Sub img��Ƭ_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        msngStartX = X
        msngStartY = Y
    End If
End Sub

Private Sub img��Ƭ_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sngLeft As Single
    Dim sngTop As Single
    
    '����״̬������
    If mnuViewStretch.Checked = True Then Exit Sub
    If Button = 1 Then
        '����������ܵ�
        sngLeft = img��Ƭ.Left + X - msngStartX
        sngTop = img��Ƭ.Top + Y - msngStartY
        
        '���ÿ��ܵ���߾�
        If img��Ƭ.Width < pic����.ScaleWidth Or sngLeft > pic����.ScaleLeft Then
            sngLeft = pic����.ScaleLeft
        Else
            If sngLeft + img��Ƭ.Width < pic����.ScaleWidth Then
                sngLeft = pic����.ScaleWidth - img��Ƭ.Width
            End If
        End If
        '���ÿ��ܵĶ��߾�
        If img��Ƭ.Height < pic����.ScaleHeight Or sngTop > pic����.ScaleTop Then
            sngTop = pic����.ScaleTop
        Else
            If sngTop + img��Ƭ.Height < pic����.ScaleHeight Then
                sngTop = pic����.ScaleHeight - img��Ƭ.Height
            End If
        End If
        img��Ƭ.Left = sngLeft
        img��Ƭ.Top = sngTop
    End If
End Sub

Private Sub lbl����_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And Index = 3 Then msngStartY = Y
End Sub

Private Sub lbl����_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And Index = 3 Then
        If lvwMain.Height + Y - msngStartY < 2000 Or lvwCert.Height - (Y - msngStartY) < 500 Then Exit Sub
        lbl����(3).Top = lbl����(3).Top + Y - msngStartY
        lvwMain.Height = lvwMain.Height + Y - msngStartY
        lvwCert.Top = lvwCert.Top + Y - msngStartY
        lvwCert.Height = lvwCert.Height - (Y - msngStartY)
        lvwLogOnOff.Top = lvwCert.Top
        lvwLogOnOff.Height = lvwCert.Height
        pic֤��.Top = lvwCert.Top
        pic֤��.Height = lvwLogOnOff.Height
    End If
End Sub

Private Sub lvwCert_DblClick()
    Call lvwCert_KeyPress(13)
End Sub

Private Sub lvwCert_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Not lvwCert.SelectedItem Is Nothing Then
            If mnuEditViewCert.Enabled And mnuEditViewCert.Visible Then
                Call mnuEditViewCert_Click
            End If
        End If
    End If
End Sub

Private Sub lvwMain_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If mintColumn = ColumnHeader.Index - 1 Then '���Ǹղ�����
        lvwMain.SortOrder = IIF(lvwMain.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintColumn = ColumnHeader.Index - 1
        lvwMain.SortKey = mintColumn
        lvwMain.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwMain_DblClick()
    If mblnItem = True And mnuEditModify.Enabled And mnuEditModify.Visible Then mnuEditModify_Click
End Sub

Private Sub lvwMain_GotFocus()
    SetMenu
End Sub

Public Sub lvwMain_ItemClick(ByVal Item As MSComctlLib.ListItem)
    ShowAttribe
    Call mobjForm.initVSf(Val(Mid(Item.Key, 2)), 1)
    
    mblnItem = True
    
    Call lvwMain_GotFocus
End Sub

Private Sub lvwMain_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If mnuEditModify.Enabled And mnuEditModify.Visible Then mnuEditModify_Click
    End If
End Sub

Private Sub lvwMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnItem = False
End Sub

Private Sub lvwMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    If Button = 2 Then
        mnuShortMenu(1).Enabled = mnuEditNew.Enabled
        mnuShortMenu(2).Enabled = mnuEditModify.Enabled
        mnuShortMenu(3).Enabled = mnuEditDelete.Enabled
        mnuShortMenu(5).Enabled = mnuEditAdjust.Enabled
        mnuShortMenu(6).Enabled = mnuEditRole.Enabled
        For i = 0 To 3
            mnuShortIcon(i).Checked = mnuViewIcon(i).Checked
        Next
        PopupMenu mnuShort, vbPopupMenuRightButton
    End If
End Sub

Private Sub mnuEditAdjust_Click()
    '��Ա���ŵ���
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    If Mid(lvwMain.SelectedItem.Key, 2) = glngUserId And InStr(mstrPrivs, "���в���") = 0 Then
        MsgBox "������Ե�ǰ��¼��Ա���С���Ա���ŵ�������", vbInformation, gstrSysName
        Exit Sub
    End If
    With frmPresAdjust
        .EntryPort Mid(lvwMain.SelectedItem.Key, 2), mstrPrivs
        .Show vbModal, Me
        mstrKey = ""
        Call tvwMain_S_NodeClick(tvwMain_S.SelectedItem)
    End With
End Sub

Private Sub mnuEditDelete_Click()
    Dim intIndex As Integer
    Dim blnRisTrans As Boolean
    
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    If InStr(mstrPrivs, "���в���") = 0 Then
        If glngUserId = Val(Mid(lvwMain.SelectedItem.Key, 2)) Then
            MsgBox "����ɾ����ǰ��¼�û���Ӧ����Ա��", vbInformation, gstrSysName
            Exit Sub
        End If
        'If CheckDeptPermission(glngUserId, Val(Mid(lvwMain.SelectedItem.Key, 2))) = False Then Exit Sub
    End If
    
    If MsgBox("��ȷ��Ҫɾ������Ϊ��" & lvwMain.SelectedItem.Text & "������Ա��", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
        On Error GoTo ErrHandle
        
        '����RIS�ӿڣ�ɾ����Ա��Ϣ����׼�棬���ò�������������Ϊ����顱�Ĳ�����Ա���ӿڲ�����Ч��ǰ����
        If Int(glngSys / 100) = 1 And mblnPACSInterface = True Then
            If IsCheckDeptPres(Val(Mid(lvwMain.SelectedItem.Key, 2))) Then
                If Not gobjRIS Is Nothing Then
                    If gobjRIS.HISBasicDictTable(RISBaseItemType.Personnel, RISBaseItemOper.Delete, Val(Mid(lvwMain.SelectedItem.Key, 2))) <> 1 Then
                        '����ʱ��ʾ�ӿڴ�����Ϣ
                        If gobjRIS.LastErrorInfo <> "" Then
                            MsgBox gobjRIS.LastErrorInfo, vbInformation, gstrSysName
                        Else
                            MsgBox "����RIS�ӿڴ��󣬲��ܼ�����ǰ����������ϵͳ����Ա��ϵ", vbInformation, gstrSysName
                        End If

                        Exit Sub
                    End If
                    
                    blnRisTrans = True
                Else
                    '�ӿڲ�����Чʱ��ֹ����ʾ
                    MsgBox "RIS�ӿڴ���ʧ�ܣ����ܼ�����ǰ�����������ǽӿ��ļ���װ��ע�᲻����������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
                    
                    Exit Sub
                End If
            End If
        End If
        
        gstrSQL = "zl_��Ա��_delete(" & Mid(lvwMain.SelectedItem.Key, 2) & ")"
        Call SQLTest(App.ProductName, Me.Caption, gstrSQL)
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        Call SQLTest
        
        blnRisTrans = False
        
        With lvwMain
            intIndex = .SelectedItem.Index
            .ListItems.Remove .SelectedItem.Key
            If .ListItems.Count > 0 Then
                intIndex = IIF(.ListItems.Count > intIndex, intIndex, .ListItems.Count)
                .ListItems(intIndex).Selected = True
                .ListItems(intIndex).EnsureVisible
            End If
        End With
        
        Call ShowAttribe
        If lvwMain.SelectedItem Is Nothing Then
            Call mobjForm.initVSf(0, 1)
        Else
            Call mobjForm.initVSf(Val(Mid(lvwMain.SelectedItem.Key, 2)), 1)
        End If
        Call SetMenu
    End If
    Exit Sub
ErrHandle:
    'Ris�ӿں�HIS��ͬ��ʱ��д������־
    If blnRisTrans = True And Not gobjRIS Is Nothing Then
        MsgBox "HISɾ����Ա��Ϣ����RIS�ӿں�HIS���ݲ�ͬ��������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
        
        On Error Resume Next
        Call gobjRIS.WriteCommLog("frmPresManage��mnuEditDelete_Click", "HISɾ����Ա��Ϣ����RIS�ӿں�HIS���ݲ�ͬ��", "��ԱID=" & Val(Mid(lvwMain.SelectedItem.Key, 2)), 0)
    End If
    
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function IsCheckDeptPres(ByVal lngPres As Long) As Boolean
    '�Ƿ��������Ա
    Dim rsData  As ADODB.Recordset
    
    gstrSQL = "Select 1 From ������Ա A, ��������˵�� B Where a.����id = b.����id And �������� = '���' And a.��Աid = [1] "
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "IsCheckDeptPres", lngPres)
    
    IsCheckDeptPres = Not rsData.EOF
End Function
Private Sub mnuEditDeptRole_Click()
    Dim frmTmp As frmPresRoleBat
    If tvwMain_S.SelectedItem.Key = "Root" Then Exit Sub
    Set frmTmp = New frmPresRoleBat
    frmTmp.ShowMe Me, Val(Mid(tvwMain_S.SelectedItem.Key, 2)), tvwMain_S.SelectedItem.Text
End Sub

Private Sub mnuEditModify_Click()
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    With lvwMain.SelectedItem
'        If InStr(mstrPrivs, "���в���") = 0 Then
'            If CheckDeptPermission(glngUserId, Val(Mid(.Key, 2))) Then
               frmPresSet.�༭��Ա Mid(.Key, 2)
'            End If
'        Else
'            frmPresSet.�༭��Ա Mid(.Key, 2)
'        End If

        If InStr(1, mstrPrivs, ";�޸�ʱ���޶���Ա����;") = 0 Then
        gstrSQL = "Select 1 From ��Ա����˵�� Where ��Աid =  [1] and rownum>=1"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����Ա�Ƿ�����Ȩ��", glngUserId)
        
        If rsTemp.RecordCount > 0 Then
            gstrSQL = "Select Distinct ��Աid From ��Ա����˵�� Where ��Ա���� In (Select ��Ա���� From ��Ա����˵�� Where ��Աid = [1])"
        Else
            '��ѯ�͵�ǰ����Ա������ͬ���ʵ���Ա
            gstrSQL = "Select ID As ��Աid" & vbNewLine & _
                "From ��Ա��" & vbNewLine & _
                "Where ID Not In (Select Distinct ��Աid From ��Ա����˵��)"
        End If
        Set mrsPersonProper = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯ����Ա��������", glngUserId)
    End If
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuEditNew_Click()
    On Error GoTo ErrHandle
    If tvwMain_S.SelectedItem Is Nothing Then Exit Sub
    With tvwMain_S.SelectedItem
        If InStr(mstrPrivs, "���в���") = 0 And .ForeColor <> vbBlack Then
            MsgBox "�㲻���ڡ�" & .Text & "����������Ա��Ϣ��", vbInformation, gstrSysName
            Exit Sub
        End If
        frmPresSet.�༭��Ա , Mid(.Key, 2)
        
        '��ѯ�͵�ǰ����Ա������ͬ���ʵ���Ա
        gstrSQL = "Select ID As ��Աid" & vbNewLine & _
                "From ��Ա��" & vbNewLine & _
                "Where ID Not In (Select Distinct ��Աid From ��Ա����˵��)" & vbNewLine & _
                "Union" & vbNewLine & _
                "Select Distinct ��Աid From ��Ա����˵�� Where ��Ա���� In(Select ��Ա���� From ��Ա����˵�� Where ��Աid = [1])"
        Set mrsPersonProper = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯ����Ա��������", glngUserId)
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuEditRegCert_Click()
    Dim arrData As Variant
    
    On Error GoTo errH
    
    If Not mobjESign Is Nothing Then
        If mobjESign.RegisterCertificate(arrData, Val(Mid(lvwMain.SelectedItem.Key, 2))) Then
            If arrData(0) <> lvwMain.SelectedItem.Text Then
                If MsgBox("������֤���ǰ䷢��""" & arrData(0) & """������ǰע����ԱΪ""" & lvwMain.SelectedItem.Text & """��Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            End If
            
            '����ǩ��ͼƬ
            gcnOracle.BeginTrans
            If UBound(arrData) > 4 Then
                If arrData(5) <> "" Then
                    If SaveSignPIC(Mid(lvwMain.SelectedItem.Key, 2), arrData(5)) = False Then
                        GoTo errH
                    End If
                End If
            End If
            
            gstrSQL = "zl_��Ա֤���¼_Insert(" & _
                Val(Mid(lvwMain.SelectedItem.Key, 2)) & "," & _
                "'" & Replace(arrData(1), "'", "''") & "'," & _
                "'" & Replace(arrData(2), "'", "''") & "'," & _
                "'" & Replace(arrData(3), "'", "''") & "'," & _
                "'" & Replace(arrData(4), "'", "''") & "'," & _
                "'" & Replace(arrData(6), "'", "''") & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            If CStr(arrData(7)) <> "" Then
                If Not Sys.SaveLob(glngSys, 14, Val(Mid(lvwMain.SelectedItem.Key, 2)) & "," & Trim(arrData(2)), CStr(arrData(7)), 1) Then
                    GoTo errH
                End If
            End If
            Call ShowAttribe
            Call SetMenu
            gcnOracle.CommitTrans
            
            MsgBox "����֤��ע��ɹ���""" & lvwMain.SelectedItem.Text & """��������������ʹ�ø�֤����е���ǩ����", vbInformation, gstrSysName
        End If
    End If
    Exit Sub
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditDelCert_Click()
    Dim arrData As Variant
    
    On Error GoTo errH
    
    If MsgBox("ȷʵҪȡ����Ա""" & lvwMain.SelectedItem.Text & """��ǰѡ�������֤��ע����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    gstrSQL = "zl_��Ա֤���¼_Delete(" & Val(Mid(lvwCert.SelectedItem.Key, 2)) & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    Call ShowAttribe
    Call SetMenu
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditRole_Click()
    Dim frmTmp As frmPresRole
    '��Ա��ɫ����
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    If Mid(lvwMain.SelectedItem.Key, 2) = glngUserId Then
        MsgBox "������Ե�ǰ��¼��Ա���С���Ա��ɫ���䡱��", vbInformation, gstrSysName
        Exit Sub
    End If
    If CheckIsUser(Mid(lvwMain.SelectedItem.Key, 2)) = False Then
        Exit Sub
    End If
    Set frmTmp = New frmPresRole
    Call frmTmp.ShowMe(Me, Val(Mid(lvwMain.SelectedItem.Key, 2)))
End Sub

Private Sub mnuEditStart_Click()
    Dim strKey As String
    
    On Error GoTo ErrHandle
    
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    If tvwMain_S.SelectedItem.Image = "Dept_No" Then
        MsgBox "����Ա�������Ż���ͣ��״̬���뵽���Ź��������ö��ڲ��ţ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    gstrSQL = "Zl_��Ա��_����(" & Mid(lvwMain.SelectedItem.Key, 2) & ")"
            
    'ִ�����ù���
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    mstrKey = ""
    Call tvwMain_S_NodeClick(tvwMain_S.SelectedItem)
'    If InStr(mstrPrivs, "���в���") = 0 Then
'        Call FillTreePrivs
'    Else
'        Call FillTree
'    End If
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub mnuEditStop_Click()
    Dim strKey As String
    
    On Error GoTo ErrHandle
    
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    If InStr(mstrPrivs, "���в���") = 0 And Mid(lvwMain.SelectedItem.Key, 2) = glngUserId Then
        MsgBox "�ܾ�ͣ���û���Ӧ����Ա��", vbInformation, gstrSysName
        Exit Sub
    End If
    
    frmPresStop.�༭��Ա (Mid(lvwMain.SelectedItem.Key, 2))
    mstrKey = ""
    Call tvwMain_S_NodeClick(tvwMain_S.SelectedItem)
    
'    If InStr(mstrPrivs, "���в���") = 0 Then
'        Call FillTreePrivs
'    Else
'        Call FillTree
'    End If
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuEditViewCert_Click()
    If Not mobjESign Is Nothing Then
        If mobjESign.ViewCertificate(Val(Mid(lvwCert.SelectedItem.Key, 2))) Then
            
        End If
    End If
End Sub

Private Sub mnuFileFile_Click()
'    Call ���鱨��(Me)
    Dim objNow As Object
    
    On Error Resume Next
    
    Set objNow = CreateObject("zl9MedRec.ClsMedRec")
    Call objNow.���鱨��(Me)
End Sub

Private Sub mnuFileReport_Click()
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    
    On Error Resume Next
    ReportOpen gcnOracle, glngSys, "ZL1_INSIDE_1002_1", Me, "��ԱID=" & Mid(lvwMain.SelectedItem.Key, 2), 1
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    'Ĭ�ϲ���������=����id����Ա=��Աid
    Dim lng����ID As Long
    Dim lng��Աid As Long
    
    If Not tvwMain_S.SelectedItem Is Nothing Then
        If tvwMain_S.SelectedItem.Key <> "Root" Then
            lng����ID = Mid(tvwMain_S.SelectedItem.Key, 2)
        End If
    End If
    
    If Not lvwMain.SelectedItem Is Nothing Then
        lng��Աid = Mid(lvwMain.SelectedItem.Key, 2)
    End If
    
    Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
        "����=" & IIF(lng����ID = 0, "", lng����ID), _
        "��Ա=" & IIF(lng��Աid = 0, "", lng��Աid))
End Sub

Private Sub mnuShortSign_Click(Index As Integer)
    Select Case Index
        Case 0      '����֤������
            Call mnuEditSignOn_Click
        Case 1      '����֤��ͣ��
            Call mnuEditSignOff_Click
    End Select
End Sub

Private Sub mnuViewColumn_Click()
    If zlControl.LvwSelectColumns(lvwMain, mstrLvw) = True Then
        '���б仯��Ҫ����ˢ��
        FillList tvwMain_S.SelectedItem.Key
    End If
End Sub

Private Sub mnuViewFind_Click()
    frmPresFind.ShowOfType Me, 0, mnuViewShowStop.Checked
End Sub

Private Sub mnuViewReflash_Click()
    If InStr(mstrPrivs, "���в���") = 0 Then
        FillTreePrivs
    Else
        FillTree
    End If
End Sub

Private Sub mnuViewShow_Click()
    mnuViewShow.Checked = Not mnuViewShow.Checked
    FillList tvwMain_S.SelectedItem.Key
End Sub

Private Sub mnuViewShowStop_Click()
    mnuViewShowStop.Checked = Not mnuViewShowStop.Checked
    FillList tvwMain_S.SelectedItem.Key
End Sub

Private Sub mnuViewShowStopDept_Click()
    mnuViewShowStopDept.Checked = Not mnuViewShowStopDept.Checked
    Call FillTree
End Sub

Private Sub mnuViewStretch_Click()
    mnuViewStretch.Checked = Not mnuViewStretch.Checked
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "��Ƭ�Զ�����", IIF(mnuViewStretch.Checked, 1, 0)
    Call Set��Ƭ����
End Sub

Private Sub picSplitV_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = 1 Then
        If tbcPres.Width - X < 2000 Or tvwMain_S.Width + X < 2000 Then Exit Sub
        picSplitV.Left = picSplitV.Left + X
        tvwMain_S.Width = tvwMain_S.Width + X
        lvwMain.Left = lvwMain.Left + X
        lvwMain.Width = lvwMain.Width - X
        
        lbl����(3).Left = lbl����(3).Left + X
        lbl����(3).Width = lbl����(3).Width - X
        lvwCert.Left = lvwCert.Left + X
        lvwCert.Width = lvwCert.Width - X
        
        picSplitH.Left = picSplitH.Left + X
        picSplitH.Width = picSplitH.Width - X
        
        tbcPres.Left = tbcPres.Left + X
        tbcPres.Width = tbcPres.Width - X
        picPres.Width = picPres.Width - X
        
        lbl����(2).Width = lbl����(2).Width - X
        txt˵��.Width = txt˵��.Width - X
        
        tvwMain_S.SetFocus
    End If
End Sub

Private Sub pic����_Resize()
    If mnuViewStretch.Checked = True Then
        '����
        img��Ƭ.Width = pic����.ScaleWidth
        img��Ƭ.Height = pic����.ScaleHeight
    Else
        '��������
        If pic����.ScaleWidth > img��Ƭ.Width Then
            img��Ƭ.Left = pic����.ScaleLeft
        Else
            If img��Ƭ.Left + img��Ƭ.Width < pic����.ScaleWidth Then
                img��Ƭ.Left = pic����.ScaleWidth - img��Ƭ.Width
            End If
        End If
        
        If pic����.ScaleHeight > img��Ƭ.Height Then
            img��Ƭ.Top = pic����.ScaleTop
        Else
            If img��Ƭ.Top + img��Ƭ.Height < pic����.ScaleHeight Then
                img��Ƭ.Top = pic����.ScaleHeight - img��Ƭ.Height
            End If
        End If
    End If
End Sub

Private Sub pic��Ƭ_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        msngStartX = X
    End If
End Sub

Private Sub pic��Ƭ_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = 1 Then
        If lvw��Ա����_S.Width + X < 1000 Or pic����.Width - X < 1000 Then Exit Sub
        
        lvw��Ա����_S.Width = lvw��Ա����_S.Width + X
        lbl����(0).Width = lvw��Ա����_S.Width
        
        pic��Ƭ.Left = pic��Ƭ.Left + X
        
        pic����.Left = pic����.Left + X
        pic����.Width = pic����.Width - X
        lbl����(1).Left = pic����.Left
        lbl����(1).Width = pic����.Width
        
        lvwMain.SetFocus
    End If
End Sub

Private Sub pic˵��_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        msngStartX = X
    End If
End Sub

Private Sub pic˵��_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = 1 Then
        If pic����.Width + X < 1000 Or txt˵��.Width - X < 1000 Then Exit Sub
        
        pic����.Width = pic����.Width + X
        lbl����(1).Width = pic����.Width
        
        pic˵��.Left = pic˵��.Left + X
        
        txt˵��.Left = txt˵��.Left + X
        txt˵��.Width = txt˵��.Width - X
        lbl����(2).Left = txt˵��.Left
        lbl����(2).Width = txt˵��.Width
        
        lvwMain.SetFocus
    End If
End Sub

Private Sub picSplitH_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = 1 Then
        If lvwCert.Visible Then
            If tbcPres.Height - Y < 800 Or lvwCert.Height + Y < 500 Then Exit Sub
        Else
            If tbcPres.Height - Y < 800 Or lvwMain.Height + Y < 2000 Then Exit Sub
        End If
        
        picSplitH.Top = picSplitH.Top + Y
                
        If lvwCert.Visible Then
            lvwCert.Height = lvwCert.Height + Y
            pic֤��.Height = lvwCert.Height
            lvwLogOnOff.Height = lvwCert.Height
        Else
            lvwMain.Height = lvwMain.Height + Y
        End If
        
        tbcPres.Top = tbcPres.Top + Y
        tbcPres.Height = tbcPres.Height - Y
        picPres.Height = picPres.Height - Y

        lvw��Ա����_S.Height = lvw��Ա����_S.Height - Y

        pic����.Height = pic����.Height - Y
        txt˵��.Height = txt˵��.Height - Y
        pic��Ƭ.Height = pic��Ƭ.Height - Y
        pic˵��.Height = pic˵��.Height - Y
        
        lvwMain.SetFocus
    End If
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileExcel_Click()
    subPrint 3
End Sub

Private Sub mnufilePreview_Click()
    subPrint 2
End Sub

Private Sub mnuFilePrint_Click()
    subPrint 1
End Sub

Private Sub mnufileset_Click()
    zlPrintSet
End Sub

Private Sub pic֤��_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        msngStartX = X
    End If
End Sub

Private Sub pic֤��_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sngTemp As Single
    If Button = 1 Then
        sngTemp = pic֤��.Left + X - msngStartX
        If sngTemp - lvwCert.Left > 1000 And ScaleWidth - (sngTemp + pic֤��.Width) > 1000 Then
            pic֤��.Left = sngTemp
            lvwCert.Width = pic֤��.Left - lvwCert.Left
            lvwLogOnOff.Left = pic֤��.Left + pic֤��.Width
            lvwLogOnOff.Width = lvwMain.Left + lvwMain.Width - pic֤��.Width - lvwCert.Width - lvwCert.Left
        End If
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "New"
            mnuEditNew_Click
        Case "Modify"
            mnuEditModify_Click
        Case "Delete"
            mnuEditDelete_Click
        Case "Start"
            mnuEditStart_Click
        Case "Stop"
            mnuEditStop_Click
        Case "Quit"
            mnuFileExit_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Preview"
            mnufilePreview_Click
        Case "Help"
            mnuhelptopic_Click
        Case "Find"
            mnuViewFind_Click
        Case "View"
            mnuViewIcon(lvwMain.View).Checked = False
            If lvwMain.View = 3 Then
                mnuViewIcon(0).Checked = True
                lvwMain.View = 0
            Else
                mnuViewIcon(lvwMain.View + 1).Checked = True
                lvwMain.View = lvwMain.View + 1
            End If
        Case "plugIn"
            PopupMenu mnuPlugIn, vbPopupMenuRightButton
        Case "Sign"
            '���ڴ˴�����
    End Select
End Sub

Private Sub Toolbar1_ButtonDropDown(ByVal Button As MSComctlLib.Button)
    If Button.Key = "Sign" Then
        Button.ButtonMenus("SignOn").Enabled = mnuEditSignOn.Enabled
        Button.ButtonMenus("SignOff").Enabled = mnuEditSignOff.Enabled
    End If
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Dim i As Integer
    
    If ButtonMenu.Key = "SignOn" Then
        Call mnuEditSignOn_Click
    ElseIf ButtonMenu.Key = "SignOff" Then
        Call mnuEditSignOff_Click
    Else
        For i = 0 To 3
            mnuViewIcon(i).Checked = False
            Toolbar1.Buttons("View").ButtonMenus(i + 1).Text = Replace(Toolbar1.Buttons("View").ButtonMenus(i + 1).Text, "��", "  ")
        Next
        mnuViewIcon(ButtonMenu.Index - 1).Checked = True
        Toolbar1.Buttons("View").ButtonMenus(ButtonMenu.Index).Text = Replace(Toolbar1.Buttons("View").ButtonMenus(ButtonMenu.Index).Text, "  ", "��")
        lvwMain.View = ButtonMenu.Index - 1
    End If
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    Me.mnuViewToolText.Enabled = mnuViewToolButton.Checked
    CoolBar1.Visible = mnuViewToolButton.Checked
    CoolBar1.Bands("only").MinHeight = Toolbar1.Height
    Form_Resize
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = mnuViewStatus.Checked
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim buttTemp As Button
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For Each buttTemp In Toolbar1.Buttons
        If mnuViewToolText.Checked Then
            buttTemp.Caption = buttTemp.Tag
        Else
            buttTemp.Caption = ""
        End If
    Next
    CoolBar1.Bands("only").MinHeight = Toolbar1.Height
    Form_Resize
End Sub

Private Sub mnuViewIcon_Click(Index As Integer)
    Dim i As Integer
    For i = 0 To 3
        mnuViewIcon(i).Checked = False
        Toolbar1.Buttons("View").ButtonMenus(i + 1).Text = Replace(Toolbar1.Buttons("View").ButtonMenus(i + 1).Text, "��", "  ")
    Next
    mnuViewIcon(Index).Checked = True
    Toolbar1.Buttons("View").ButtonMenus(Index + 1).Text = Replace(Toolbar1.Buttons("View").ButtonMenus(Index + 1).Text, "  ", "��")
    lvwMain.View = Index
End Sub

Private Sub mnuShortMenu_Click(Index As Integer)
    Select Case Index
        Case 1
            mnuEditNew_Click
        Case 2
            mnuEditModify_Click
        Case 3
            mnuEditDelete_Click
        Case 6
            mnuEditAdjust_Click
        Case 7
            mnuEditRole_Click
    End Select
End Sub

Private Sub mnuShortIcon_Click(Index As Integer)
    mnuViewIcon_Click Index
End Sub

Private Sub mnuhelptopic_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hwnd)
End Sub

Private Sub Toolbar1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuViewTool
    End If
End Sub

Private Sub tvwMain_S_GotFocus()
    SetMenu
    If mnuViewShow.Checked = True Then
        stbThis.Panels(2).Text = "�ò�������Ա" & lvwMain.ListItems.Count & "���������¼����ţ���"
    Else
        stbThis.Panels(2).Text = "�ò�������Ա" & lvwMain.ListItems.Count & "����"
    End If
End Sub

Private Sub tvwMain_S_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim objNode As Node
    If tvwMain_S.SelectedItem Is Nothing Then Exit Sub
    Set objNode = tvwMain_S.SelectedItem
    If mstrKey = objNode.Key Then Exit Sub
    mstrKey = objNode.Key
    
    If objNode.ForeColor = &H8000000C Then
        lvwMain.ListItems.Clear
    Else
        FillList objNode.Key
    End If
    If mnuViewShow.Checked = True Then
        stbThis.Panels(2).Text = "�ò�������Ա" & lvwMain.ListItems.Count & "���������¼����ţ���"
    Else
        stbThis.Panels(2).Text = "�ò�������Ա" & lvwMain.ListItems.Count & "����"
    End If
    SetMenu
End Sub


Private Sub subPrint(bytMode As Byte)
'����:���д�ӡ,Ԥ���������EXCEL
'����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    If gstrUserName = "" Then Call GetUserInfo
    Dim objPrint As New zlPrintLvw
    Dim str��λ As String
    
    str��λ = GetUnitName
    objPrint.Title.Text = str��λ & "��Ա��"
    Set objPrint.Body.objData = lvwMain
    objPrint.BelowAppItems.Add "��ӡ�ˣ�" & gstrUserName
    objPrint.BelowAppItems.Add "��ӡʱ�䣺" & Format(Sys.Currentdate, "yyyy��MM��dd��")
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrViewLvw objPrint, 1
          Case 2
              zlPrintOrViewLvw objPrint, 2
          Case 3
              zlPrintOrViewLvw objPrint, 3
      End Select
    Else
        zlPrintOrViewLvw objPrint, bytMode
    End If

End Sub

Private Sub FillTreePrivs()
'����:װ���������ŵ�tvwMain_S
    Dim nodTmp As Node
    Dim rsDeptID As ADODB.Recordset
    Dim strTemp As String
    strTemp = "Dept"
    
    On Error GoTo ErrHandle
    gstrSQL = "Select Max(Level) as ��,A.ID,A.�ϼ�ID,A.����,'��'||A.����||'��' ���� " & _
              "From ���ű� A Start With ID IN(Select ����ID From ������Ա Where ��ԱID=[1]) Connect by Prior �ϼ�ID=ID " & _
              "Group by A.ID,A.�ϼ�ID,A.����,A.���� " & _
              "Order by A.����,�� Desc"
    Set rsDeptID = zlDatabase.OpenSQLRecord(gstrSQL, Caption, glngUserId)
    With tvwMain_S
        .Sorted = True
        .Nodes.Clear
        Do While Not rsDeptID.EOF
            If IIF(IsNull(rsDeptID!�ϼ�id), 0, rsDeptID!�ϼ�id) = 0 Then
                If .Nodes.Count > 0 Then
                    If FindKey("C" & rsDeptID!ID) = False Then
                        Set nodTmp = .Nodes.Add(, , "C" & rsDeptID!ID, rsDeptID!���� & rsDeptID!����, strTemp, strTemp)
                    Else
                        Set nodTmp = .Nodes("C" & rsDeptID!ID)
                    End If
                Else
                    Set nodTmp = .Nodes.Add(, , "C" & rsDeptID!ID, rsDeptID!���� & rsDeptID!����, strTemp, strTemp)
                End If
            Else
                If FindKey("C" & rsDeptID!ID) = False Then
                    Set nodTmp = .Nodes.Add("C" & rsDeptID!�ϼ�id, tvwChild, "C" & rsDeptID!ID, rsDeptID!���� & rsDeptID!����, strTemp, strTemp)
                Else
                    Set nodTmp = .Nodes("C" & rsDeptID!ID)
                End If
            End If
            nodTmp.ForeColor = &H8000000C
            rsDeptID.MoveNext
        Loop
        rsDeptID.Close
    End With
    '�����ӽ��
    gstrSQL = "Select ID,�ϼ�ID,'��'||����||'��' ����,���� " & _
              "From ���ű� A " & _
              "Start With ID IN(Select ����ID From ������Ա Where ��ԱID=[1]) Connect by Prior ID=�ϼ�ID"
    Set rsDeptID = zlDatabase.OpenSQLRecord(gstrSQL, Caption, glngUserId)
    With tvwMain_S
        Do While Not rsDeptID.EOF
            If IIF(IsNull(rsDeptID!�ϼ�id), 0, rsDeptID!�ϼ�id) = 0 Then
                If .Nodes.Count > 0 Then
                    If FindKey("C" & rsDeptID!ID) = False Then
                        Set nodTmp = .Nodes.Add(, , "C" & rsDeptID!ID, rsDeptID!���� & rsDeptID!����, strTemp, strTemp)
                    Else
                        Set nodTmp = .Nodes("C" & rsDeptID!ID)
                    End If
                Else
                    Set nodTmp = .Nodes.Add(, , "C" & rsDeptID!ID, rsDeptID!���� & rsDeptID!����, strTemp, strTemp)
                End If
            Else
                If FindKey("C" & rsDeptID!ID) = False Then
                    Set nodTmp = .Nodes.Add("C" & rsDeptID!�ϼ�id, tvwChild, "C" & rsDeptID!ID, rsDeptID!���� & rsDeptID!����, strTemp, strTemp)
                Else
                    Set nodTmp = .Nodes("C" & rsDeptID!ID)
                End If
            End If
            nodTmp.ForeColor = vbBlack
            rsDeptID.MoveNext
        Loop
        rsDeptID.Close
    End With
    
    If tvwMain_S.Nodes.Count > 0 Then tvwMain_S.Nodes(1).Selected = True
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function FillTree() As Boolean
'����:װ�����в��ŵ�tvwMain_S
'����:
    Dim strTemp As String
    Dim strKey As String
    Dim rs���� As New ADODB.Recordset
    
    
    mstrKey = ""
    
    On Error GoTo ErrHandle
    rs����.CursorLocation = adUseClient
    
    If Not tvwMain_S.SelectedItem Is Nothing Then
        strKey = tvwMain_S.SelectedItem.Key
    End If
    If mnuViewShowStopDept.Checked = True Then
        strTemp = ""
    Else
        strTemp = " where (����ʱ�� = to_date('3000-01-01','YYYY-MM-DD') or ����ʱ�� is null ) "
    End If
    
    gstrSQL = "select id,�ϼ�id,���� ,����,to_char(����ʱ��,'YYYY-MM-DD') as ����ʱ��  from ���ű� " & strTemp & " start with �ϼ�id is null connect by prior id =�ϼ�id "
    Set rs���� = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)

    tvwMain_S.Nodes.Clear
    tvwMain_S.Nodes.Add , , "Root", "���в���", "Root", "Root"
    tvwMain_S.Nodes("Root").Sorted = True
'    strTemp = "Dept"
    Do Until rs����.EOF
        If CDate(IIF(IsNull(rs����("����ʱ��")), CDate("3000/1/1"), rs����("����ʱ��"))) = CDate("3000/1/1") Then
            strTemp = "Dept"
        Else
            strTemp = "Dept_No"
        End If
        If IsNull(rs����("�ϼ�id")) Then
            tvwMain_S.Nodes.Add "Root", tvwChild, "C" & rs����("id"), "��" & rs����("����") & "��" & rs����("����"), strTemp, strTemp
        Else
            tvwMain_S.Nodes.Add "C" & rs����("�ϼ�id"), tvwChild, "C" & rs����("id"), "��" & rs����("����") & "��" & rs����("����"), strTemp, strTemp
        End If
        tvwMain_S.Nodes("C" & rs����("id")).Sorted = True
        rs����.MoveNext
    Loop
    If tvwMain_S.Nodes.Count = 1 Then
        MsgBox "������Ϣ��ȫ������ʹ�ñ�ģ�顣" & vbCrLf & "������Ϣ���ڡ����Ź���������", vbInformation, gstrSysName
        FillTree = False
        Exit Function
    End If
    
    
    Dim nod As Node
    On Error Resume Next
    Set nod = tvwMain_S.Nodes(strKey)
    If Err <> 0 Then
        Set nod = tvwMain_S.Nodes(2) '�ܿ����ڵ�
        nod.Selected = True
        nod.EnsureVisible
        tvwMain_S_NodeClick nod
    Else
        Err.Clear
        nod.Selected = True
        If nod.Key = "Root" Then nod.Expanded = True
        nod.EnsureVisible
        tvwMain_S_NodeClick nod
    End If
    FillTree = True
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub FillList(ByVal str����ID As String)
'����:װ���Ӧ���ŵ���Ա��lvwMain
'����:str����ID ���ŵı�ʶ

    Dim rs��Ա As New ADODB.Recordset
    Dim lst As ListItem
    Dim i As Integer, varValue As Variant
    Dim strKey As String
    Dim stroldkey As String
    Dim rsTemp As ADODB.Recordset
    Dim bln�������� As Boolean
    
    On Error GoTo ErrHandle
    If Not lvwMain.SelectedItem Is Nothing Then
        '����ԭ�м�ֵ
        strKey = lvwMain.SelectedItem.Key
    End If
    rs��Ա.CursorLocation = adUseClient
    rs��Ա.CursorType = adOpenKeyset
    rs��Ա.LockType = adLockReadOnly
    
    Call FS.ShowFlash("���ڻ�ȡ��Ա��Ϣ����,���Ժ� ...", Me)
    If str����ID = "Root" Then
        gstrSQL = "select a.ID,C.����ID,a.����,a.���,to_char(A.��������,'yyyy-MM-dd') as ��������,A.���֤��,A.�Ա�,A.����,a.���� ,b.���� as �������� " & _
                    "   ,A.���˼��,A.רҵ����ְ��,A.����ְ��,Decode(D.����,1,'������ʹ��',2,'����ʹ��',3,'����ʹ��','') as סԺ����ҩ��Ȩ��" & _
                    "   ,Decode(D.����,1,'������ʹ��',2,'����ʹ��',3,'����ʹ��','') as ���￹��ҩ��Ȩ�� " & _
                    "   ,A.�����ȼ�,A.�칫�ҵ绰,A.�ƶ��绰,A.�����ʼ�,A.ѧ��,decode(A.Ƹ�μ���ְ��,1,'����',2,'����',3,'�м�',4,'����/ʦ��',5,'Ա/ʿ',9,'��Ƹ') as Ƹ�μ���ְ��" & _
                    "   ,to_char(A.����ʱ��,'YYYY-MM-DD') as ����ʱ��,Nvl(To_Char(A.����ʱ��, 'YYYY-MM-DD'), '3000-01-01') as ����ʱ��" & _
                    " from ��Ա�� a,���ű� b,������Ա C,��Ա����ҩ��Ȩ�� D, ��Ա����ҩ��Ȩ�� E " & _
                    " where a.id = C.��Աid  and C.����ID=B.ID And D.��ԱID(+)=a.id And (D.��¼״̬=1 or D.��¼״̬ is null) and d.����(+) = 1 " & _
                    "   and e.��Աid(+)=a.id and (e.��¼״̬=1 or e.��¼״̬ is null) and e.����(+) = 2 " & _
                    "   and B.ID in (select ID from ���ű� start with �ϼ�ID is null connect by prior id=�ϼ�ID)"
    Else
        If mnuViewShow.Checked = True Then
            gstrSQL = "select a.ID,C.����ID,a.����,a.���,to_char(A.��������,'yyyy-MM-dd') as ��������,A.���֤��,A.�Ա�,A.����,a.���� ,b.���� as �������� " & _
                        " ,A.���˼��,A.רҵ����ְ��,A.����ְ��,Decode(F.����,1,'������ʹ��',2,'����ʹ��',3,'����ʹ��','') as סԺ����ҩ��Ȩ��" & _
                        " ,Decode(g.����,1,'������ʹ��',2,'����ʹ��',3,'����ʹ��','') as ���￹��ҩ��Ȩ��" & _
                        " ,A.�����ȼ�,A.�칫�ҵ绰,A.�ƶ��绰,A.�����ʼ�,A.ѧ��,decode(A.Ƹ�μ���ְ��,1,'����',2,'����',3,'�м�',4,'����/ʦ��',5,'Ա/ʿ',9,'��Ƹ') as Ƹ�μ���ְ��" & _
                        " ,to_char(A.����ʱ��,'YYYY-MM-DD') as ����ʱ��,Nvl(To_Char(A.����ʱ��, 'YYYY-MM-DD'), '3000-01-01') as ����ʱ��" & _
                        " from ��Ա�� a,���ű� b,������Ա C, " & _
                        "(select distinct ��ԱID From ������Ա where ����id =[1]) D,��Ա����ҩ��Ȩ�� F, ��Ա����ҩ��Ȩ�� G  " & _
                        " where A.id = C.��Աid and C.����ID=B.ID and A.ID=D.��ԱID And F.��ԱID(+)=a.id And g.��ԱID(+)=a.id " & _
                        "   And (F.��¼״̬=1 or F.��¼״̬ is null) and f.����(+)=1 " & _
                        "   And (g.��¼״̬=1 or g.��¼״̬ is null) and g.����(+)=2 "
        Else
            gstrSQL = "select a.ID,C.����ID,a.����,a.���,to_char(A.��������,'yyyy-MM-dd') as ��������,A.���֤��,A.�Ա�,A.����,a.���� ,b.���� as �������� " & _
                        " ,A.���˼��,A.רҵ����ְ��,A.����ְ��,Decode(F.����,1,'������ʹ��',2,'����ʹ��',3,'����ʹ��','') as סԺ����ҩ��Ȩ��" & _
                        " ,Decode(g.����,1,'������ʹ��',2,'����ʹ��',3,'����ʹ��','') as ���￹��ҩ��Ȩ�� " & _
                        " ,A.�����ȼ�,A.�칫�ҵ绰,A.�ƶ��绰,A.�����ʼ�,A.ѧ��,decode(A.Ƹ�μ���ְ��,1,'����',2,'����',3,'�м�',4,'����/ʦ��',5,'Ա/ʿ',9,'��Ƹ') as Ƹ�μ���ְ��" & _
                        " ,to_char(A.����ʱ��,'YYYY-MM-DD') as ����ʱ��,Nvl(To_Char(A.����ʱ��, 'YYYY-MM-DD'), '3000-01-01') as ����ʱ��" & _
                        " from ��Ա�� a,���ű� b,������Ա C, " & _
                        "(select distinct ��ԱID From ������Ա where ����id in " & _
                        "(select distinct id  From ���ű� start with ID=[1] connect by prior id=�ϼ�id)) D " & _
                        "  ,��Ա����ҩ��Ȩ�� F, ��Ա����ҩ��Ȩ�� G " & _
                        " where A.id = C.��Աid and C.����ID=B.ID and A.ID=D.��ԱID And F.��ԱID(+)=a.id And (f.��¼״̬=1 or f.��¼״̬ is null) and f.����(+)=1 " & _
                        "   and g.��Աid(+)=a.id and (g.��¼״̬=1 or g.��¼״̬ is null) and g.����(+)=2 "
        End If
    End If
    
    If mnuViewShowStop.Checked = False Then
        gstrSQL = gstrSQL & " and (a.����ʱ�� = to_date('3000-01-01','YYYY-MM-DD') or a.����ʱ�� is null ) "
    End If
    
    gstrSQL = gstrSQL & " order by a.id"
    Set rs��Ա = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Mid(str����ID, 2)))
        
    Dim lng�������� As Long
    For i = 2 To lvwMain.ColumnHeaders.Count
        If lvwMain.ColumnHeaders(i).Text = "��������" Then lng�������� = i
    Next
    
    zlControl.FormLock lvwMain.hwnd
    lvwMain.ListItems.Clear
        
    Do Until rs��Ա.EOF
        bln�������� = False
        If InStr(1, mstrPrivs, ";�޸�ʱ���޶���Ա����;") = 0 Then
            mrsPersonProper.MoveFirst
            Do Until mrsPersonProper.EOF
                If rs��Ա!ID = mrsPersonProper!��ԱID Then
                    bln�������� = True
                    Exit Do
                Else
                    mrsPersonProper.MoveNext
                End If
            Loop
        Else
            bln�������� = True
        End If
        If bln�������� = True Then
            If stroldkey <> "C" & rs��Ա("ID") Then
                stroldkey = "C" & rs��Ա("ID")
                Set lst = lvwMain.ListItems.Add(, "C" & rs��Ա("ID"), rs��Ա("����"), IIF(rs��Ա!�Ա� = "��", "Item", IIF(rs��Ա!�Ա� = "Ů", "Item_G", "Item_W")), IIF(rs��Ա!�Ա� = "��", "Item", IIF(rs��Ա!�Ա� = "Ů", "Item_G", "Item_W")))
                lst.Tag = IIF(IsNull(rs��Ա("���˼��")), "", rs��Ա("���˼��"))
    
                For i = 2 To lvwMain.ColumnHeaders.Count
                    varValue = rs��Ա(lvwMain.ColumnHeaders(i).Text).value
                    lst.SubItems(i - 1) = IIF(IsNull(varValue), "", varValue)
                Next
                
                If Format(rs��Ա!����ʱ��, "YYYY-MM-DD") <> "3000-01-01" Then
                    lst.ForeColor = &HFF&
                    For i = 1 To Me.lvwMain.ColumnHeaders.Count - 1
                        lst.ListSubItems(i).ForeColor = &HFF&
                    Next
                End If
            Else
                '����Ա�Ѿ�����
                If lng�������� > 1 Then
                    '���������ʾ���Ǿ�׷�ӵ����
                    lvwMain.ListItems("C" & rs��Ա("ID")).SubItems(lng�������� - 1) = lvwMain.ListItems("C" & rs��Ա("ID")).SubItems(lng�������� - 1) & "," & rs��Ա("��������")
                End If
                Err.Clear
            End If
        End If
            
        rs��Ա.MoveNext
    Loop
    zlControl.FormLock 0
    
    If lvwMain.ListItems.Count > 0 Then
        Dim Item As ListItem
        Err.Clear
        On Error Resume Next
        Set Item = lvwMain.ListItems(strKey)
        If Err <> 0 Then
            Set Item = lvwMain.ListItems(1)
            Item.Selected = True
            Item.EnsureVisible
        Else
            Err.Clear
            Item.Selected = True
            Item.EnsureVisible
        End If
    End If
    Call ShowAttribe
    If lvwMain.SelectedItem Is Nothing Then
        Call mobjForm.initVSf(0, 1)
    Else
        Call mobjForm.initVSf(Val(Mid(lvwMain.SelectedItem.Key, 2)), 1)
    End If
    Call SetMenu
    Call FS.StopFlash
    
    Exit Sub

ErrHandle:
    Call FS.StopFlash
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ShowAttribe()
    Dim rsTemp As New ADODB.Recordset
    Dim strTempFile As String
    Dim ObjItem As ListItem
    
    On Error GoTo ErrHandle
    lvwCert.ListItems.Clear
    lvwCert.Sorted = False
    lvw��Ա����_S.ListItems.Clear
    Set img��Ƭ.Picture = Nothing
    img��Ƭ.ToolTipText = "����Ƭ"
    txt˵��.Text = ""
    
    If lvwMain.SelectedItem Is Nothing Then
        '���û��ѡ���������������б�
        Exit Sub
    End If
    rsTemp.CursorLocation = adUseClient
    
    '��ʾ��Ա��֤���¼
    gstrSQL = "Select ID,CertDN,CertSN,SignCert,EncCert,ע��ʱ��,�Ƿ�ͣ�� From ��Ա֤���¼" & _
        " Where ��ԱID=[1] Order by ע��ʱ�� Desc,ID Desc"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Mid(lvwMain.SelectedItem.Key, 2)))
    
    Do Until rsTemp.EOF
        Set ObjItem = lvwCert.ListItems.Add(, "_" & rsTemp!ID, Format(rsTemp!ע��ʱ��, "yyyy-MM-dd HH:mm:ss"), , IIF(NVL(rsTemp!�Ƿ�ͣ��, 0) = 0, "SignOn", "SignOff"))
        ObjItem.SubItems(1) = "" & rsTemp!CertSN
        ObjItem.SubItems(2) = "" & rsTemp!CertDN
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    If lvwCert.ListItems.Count > 0 Then
        lvwCert.ListItems(1).Selected = True
        '�������ǩ����ͣ�ü�¼
        Call FillLogOnOff(Val(Mid(lvwCert.SelectedItem.Key, 2)))
    Else
        Call FillLogOnOff(0)
    End If
    
    '��ʾָ����Ա������
    gstrSQL = "select A.��Ա����,B.˵�� from ��Ա����˵�� A,��Ա���ʷ��� B where A.��Ա����=B.���� and A.��ԱID=[1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Mid(lvwMain.SelectedItem.Key, 2)))
    
    Do Until rsTemp.EOF
        lvw��Ա����_S.ListItems.Add , "C" & rsTemp("��Ա����"), rsTemp("��Ա����")
        lvw��Ա����_S.ListItems("C" & rsTemp("��Ա����")).SubItems(1) = IIF(IsNull(rsTemp("˵��")), "", rsTemp("˵��"))
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    
    '��ʾ��Ƭ
    strTempFile = Sys.ReadLobV2("��Ա��Ƭ", "��Ƭ", "��ԱID=[1]", "", Val(Mid(lvwMain.SelectedItem.Key, 2)))
    img��Ƭ.Picture = LoadPicture(strTempFile)
    img��Ƭ.ToolTipText = GetPictureInfo(img��Ƭ.Picture)
    'ɾ������ʱ�ļ�
    If img��Ƭ.ToolTipText <> "����Ƭ" Then
        Kill strTempFile
    End If
    
    img��Ƭ.Left = pic����.ScaleLeft
    img��Ƭ.Top = pic����.ScaleTop
    
    '��ʾ���˼��
    txt˵��.Text = lvwMain.SelectedItem.Tag
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SetMenu()
    Dim blnEnabled As Boolean
    Dim lng����ʱ�� As Long
    Dim i As Long
    
    blnEnabled = Not tvwMain_S.SelectedItem Is Nothing
    If blnEnabled = True Then
        blnEnabled = tvwMain_S.SelectedItem.Key <> "Root"
    End If
    Toolbar1.Buttons("New").Enabled = blnEnabled
    mnuEditNew.Enabled = blnEnabled
    
    blnEnabled = Not (lvwMain.ListItems.Count = 0 Or lvwMain.SelectedItem Is Nothing)
    Toolbar1.Buttons("Modify").Enabled = blnEnabled
    Toolbar1.Buttons("Delete").Enabled = blnEnabled
'    Toolbar1.Buttons("Start").Enabled = blnEnabled
'    Toolbar1.Buttons("Stop").Enabled = blnEnabled
    mnuEditDelete.Enabled = blnEnabled
    mnuEditModify.Enabled = blnEnabled
    mnuEditExtend.Enabled = blnEnabled
    mnuFileReport.Enabled = blnEnabled
    mnuEditAdjust.Enabled = blnEnabled
    mnuEditRole.Enabled = blnEnabled
    mnuEditStart.Enabled = blnEnabled
    mnuEditStop.Enabled = blnEnabled
    
    mnuPlugIn.Enabled = blnEnabled
    mnuShortMenu(4).Enabled = blnEnabled
    Toolbar1.Buttons("plugIn").Enabled = blnEnabled
    
    If Not lvwMain.SelectedItem Is Nothing Then
        For i = 2 To lvwMain.ColumnHeaders.Count
            If lvwMain.ColumnHeaders(i).Text = "����ʱ��" Then
                lng����ʱ�� = i
                Exit For
            End If
        Next
        If lvwMain.SelectedItem.ListSubItems(lng����ʱ�� - 1) <> "3000-01-01" Then
            mnuEditStart.Enabled = True
            Toolbar1.Buttons("Start").Enabled = True
            
            mnuEditDelete.Enabled = False
            mnuEditModify.Enabled = False
            mnuEditExtend.Enabled = False
            mnuEditStop.Enabled = False
            mnuEditAdjust.Enabled = False
            mnuEditRole.Enabled = False
            Toolbar1.Buttons("Stop").Enabled = False
            Toolbar1.Buttons("Modify").Enabled = False
            Toolbar1.Buttons("Delete").Enabled = False
        Else
            mnuEditStart.Enabled = False
            Toolbar1.Buttons("Start").Enabled = False
            
            mnuEditDelete.Enabled = True
            mnuEditModify.Enabled = True
            mnuEditExtend.Enabled = True
            mnuEditStop.Enabled = True
            mnuEditAdjust.Enabled = True
            mnuEditRole.Enabled = True
            Toolbar1.Buttons("Modify").Enabled = True
            Toolbar1.Buttons("Delete").Enabled = True
            Toolbar1.Buttons("Stop").Enabled = True
        End If
    End If
    
    EnablePrint lvwMain.ListItems.Count <> 0
    
    '����֤�鹦��
    mnuEditRegCert.Enabled = blnEnabled
    mnuEditImportCertPic.Enabled = Not lvwCert.SelectedItem Is Nothing
    mnuEditViewCert.Enabled = Not lvwCert.SelectedItem Is Nothing
    mnuEditDelCert.Enabled = Not lvwCert.SelectedItem Is Nothing
        
    stbThis.Panels(2).Text = "��Ա�б���ʾ��" & lvwMain.ListItems.Count & "����Ա��"
End Sub

Private Sub EnablePrint(ByVal blnEnabled As Boolean)
'����:���ô�ӡ��Ԥ����ť����Чֵ
'����:blnEnabled ��Чֵ
    Toolbar1.Buttons("Print").Enabled = blnEnabled
    Toolbar1.Buttons("Preview").Enabled = blnEnabled
    mnuFilePreview.Enabled = blnEnabled
    mnuFilePrint.Enabled = blnEnabled
    mnuFileExcel.Enabled = blnEnabled
End Sub

Private Sub SetȨ�޿���()
'����:1.�����е��û�Ȩ�޲���,��ʹһЩ�˵����ť���ɼ�
'     2.����ǩ�����ƺͳ�ʼ��
    Dim rsTmp As New ADODB.Recordset
    Dim lngSys As Long
    
    '��ȡʹ�õĵ���ǩ����֤����
    On Error GoTo ErrHandle
    mintCA = 0
    gstrSQL = "Select ����ֵ From Zlparameters Where ϵͳ = " & glngSys & " And Nvl(˽��, 0) = 0 And ģ�� Is Null  And ������=25"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If Not rsTmp.EOF Then
        mintCA = Val(NVL(rsTmp!����ֵ))
    End If
    
    mnuShortSign(0).Visible = False
    mnuShortSign(1).Visible = False
    
    '����ǩ���ӿڿ���
    If mintCA <> 0 And InStr(mstrPrivs, "����֤��ע��") > 0 Then
        On Error Resume Next
        Set mobjESign = CreateObject("zl9ESign.clsESign")
        Err.Clear: On Error GoTo 0
        If Not mobjESign Is Nothing Then
            If mobjESign.Initialize(gcnOracle, glngSys) Then
                mnuEdit_1.Visible = True
                mnuEditRegCert.Visible = True
                mnuEditViewCert.Visible = True
                mnuEditImportCertPic.Visible = True
                
                mnuEdit_2.Visible = True
                mnuEditDelCert.Visible = True
                
                '����֤����ͣ��
                mnuEditSignLine1.Visible = True
                mnuEditSignOn.Visible = True
                mnuEditSignOff.Visible = True
                
                lbl����(3).Visible = True
                lvwCert.Visible = True
            Else
                Set mobjESign = Nothing
            End If
        End If
    End If
    
    '����֤����ͣ��
    Toolbar1.Buttons("SignLine").Visible = mnuEditSignOn.Visible
    Toolbar1.Buttons("Sign").Visible = mnuEditSignOn.Visible
    pic֤��.Visible = mnuEditSignOn.Visible
    lvwLogOnOff.Visible = mnuEditSignOn.Visible
    
    'Ȩ�޿���
    If InStr(mstrPrivs, "��ɾ��") = 0 And InStr(mstrPrivs, "����֤��ע��") = 0 Then
        mnuEdit.Visible = False
        mnuEditModify.Visible = False
        mnuEditViewCert.Visible = False
        mnuEditImportCertPic.Visible = False
        
        mnuShortMenu(1).Visible = False
        mnuShortMenu(2).Visible = False
        mnuShortMenu(3).Visible = False
        mnuShortsplit1.Visible = False
        
        Toolbar1.Buttons("Split").Visible = False
        Toolbar1.Buttons("New").Visible = False
        Toolbar1.Buttons("Modify").Visible = False
        Toolbar1.Buttons("Delete").Visible = False
    ElseIf InStr(mstrPrivs, "��ɾ��") = 0 Then
        mnuEditNew.Visible = False
        mnuEditModify.Visible = False
        mnuEditDelete.Visible = False
        mnuEdit_1.Visible = False
        
        mnuShortMenu(1).Visible = False
        mnuShortMenu(2).Visible = False
        mnuShortMenu(3).Visible = False
        mnuShortsplit1.Visible = False
        
        Toolbar1.Buttons("Split").Visible = False
        Toolbar1.Buttons("New").Visible = False
        Toolbar1.Buttons("Modify").Visible = False
        Toolbar1.Buttons("Delete").Visible = False
    ElseIf InStr(mstrPrivs, "����֤��ע��") = 0 Then
        mnuEdit_1.Visible = False
        mnuEditRegCert.Visible = False
        mnuEditViewCert.Visible = False
        mnuEditImportCertPic.Visible = False
        mnuEdit_2.Visible = False
        mnuEditDelCert.Visible = False
    ElseIf InStr(mstrPrivs, "���в���") = 0 Then
        mnuViewShow.Checked = True
        mnuViewShow.Enabled = False
    End If
    
    If InStr(mstrPrivs, ";��չ��Ϣά��;") = 0 Then
        mnuEditExtend.Visible = False
    End If
    
    gstrSQL = "Select ��� from zlSystems"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    rsTmp.Filter = "���=300"
    If Not rsTmp.EOF Then
'        mnuFileReport.Enabled = True
'        mnuFileReport.Visible = True
        mnuFileFile.Enabled = True
        mnuFileFile.Visible = True
        mnuSplit1.Visible = True
    Else
        '�ǲ���ϵͳ��������ʾ����
'        mnuFileReport.Enabled = False
'        mnuFileReport.Visible = False
        mnuFileFile.Enabled = False
        mnuFileFile.Visible = False
'        mnusplit1.Visible = False
    End If
    
    
    If glngSys = 100 Then
        mnuFileReport.Enabled = True
        mnuFileReport.Visible = True
    Else
        mnuFileReport.Enabled = False
        mnuFileReport.Visible = False
    End If
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Set��Ƭ����()
'���ܣ�������Ա��Ƭ�Ƿ����ž����Զ��仯
    Dim bln���� As Boolean
    
    bln���� = mnuViewStretch.Checked
    img��Ƭ.Stretch = bln����
    
    img��Ƭ.Left = pic����.ScaleLeft
    img��Ƭ.Top = pic����.ScaleTop
    
    If bln���� = True Then
        '����Ҫ����λ��
        img��Ƭ.MousePointer = vbArrow
        img��Ƭ.Width = pic����.ScaleWidth
        img��Ƭ.Height = pic����.ScaleHeight
    Else
        '��Ҫ����λ��
        img��Ƭ.MousePointer = vbSizeAll
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

Private Function FindKey(ByVal strKey As String) As Boolean
    Dim nodTmp As Node
    For Each nodTmp In tvwMain_S.Nodes
        If nodTmp.Key = strKey Then
            FindKey = True
            Exit Function
        End If
    Next
End Function

Private Function SaveSignPIC(ByVal lng��Աid As Long, ByVal strFileName As String) As Boolean
    Dim rsTemp As New ADODB.Recordset, blnOk As Boolean
    
    On Error GoTo ErrHandle
    blnOk = Sys.SaveLob(100, 15, lng��Աid, strFileName)
    SaveSignPIC = blnOk
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckIsUser(ByVal lngUserID As Long) As Boolean
'��鵱ǰ��Ա���޶�Ӧ�û���
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    
    On Error GoTo ErrHandle
    strTmp = "select count(��Աid) rec from �ϻ���Ա�� where ��Աid=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strTmp, Caption, lngUserID)
    If rsTmp!Rec = 1 Then
        CheckIsUser = True
    ElseIf rsTmp!Rec > 1 Then
        MsgBox "����Ա���ڶ����¼�˻����ϼ���Ա�����ݴ������⣬����ϵ����Ա���д���", vbInformation, gstrSysName
    Else
        MsgBox "����Ա��δ�����û����뵽�������д�������Ա�ĵ�¼�û���", vbInformation, gstrSysName
    End If
    rsTmp.Close
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub txtFind_GotFocus()
    zlControl.TxtSelAll txtFind
    OS.OpenIme True
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    Dim strTemp As String
    
    On Error GoTo ErrHandle
    If KeyAscii = vbKeyReturn Then
        If txtFind.Text = "" Then Exit Sub
        If mstrFindValue <> txtFind.Text And txtFind.Text <> "" Then
            mstrFindValue = txtFind.Text
            Set mrsFind = Nothing
            strTemp = " and (a.����ʱ�� = to_date('3000-01-01','YYYY-MM-DD') or a.����ʱ�� is null ) "
            gstrSQL = "Select a.Id, a.����, b.����id" & _
                       " From ��Ա�� A, ������Ա B " & _
                       " Where a.Id = b.��Աid And b.ȱʡ = 1 and (a.��� like [1] or a.���� like [1] or a.���� like [1]) "
            
            If mnuViewShowStop.Checked = False Then
                gstrSQL = gstrSQL & strTemp
            End If
            Set mrsFind = zlDatabase.OpenSQLRecord(gstrSQL, "��Ա��ѯ", UCase(txtFind.Text) & "%")
            Call LocateItem
        Else
            If Not mrsFind.EOF Then
                mrsFind.MoveNext
                Call LocateItem
            ElseIf mrsFind.RecordCount <> 0 And mrsFind.EOF Then
                mrsFind.MoveFirst
                Call LocateItem
            End If
        End If
    End If
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LocateItem()
    Dim strTemp As String
    
    txtFind.SetFocus
    txtFind.SelStart = 0
    txtFind.SelLength = Len(txtFind.Text)
    If mrsFind.RecordCount = 0 Then
        MsgBox " û���ҵ�������������Ϣ��", vbInformation, gstrSysName
        txtFind.SetFocus
        Exit Sub
    End If
    If mrsFind.EOF = True Then
        MsgBox " �Ѿ���λ�������ҵ�����Ϣ������������������", vbInformation, gstrSysName
        txtFind.SetFocus
        Exit Sub
    End If
    
    With frmPresManage.tvwMain_S
        .Nodes("C" & mrsFind("����ID")).Selected = True
        .SelectedItem.EnsureVisible
        frmPresManage.FillList "C" & mrsFind("����ID")
    End With
        
    With frmPresManage.lvwMain
        .ListItems("C" & mrsFind("ID")).Selected = True
        .SelectedItem.EnsureVisible
        frmPresManage.lvwMain_ItemClick .SelectedItem
    End With
End Sub

Private Sub FillLogOnOff(ByVal lngId As Long)
'���ܣ��������֤����ͣ�ü�¼
'������
'  lngID��֤��ID

    Dim ObjItem As ListItem
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strTmp As String
    Dim i As Integer
    
    On Error GoTo ErrHandle
    'XMLType�ֶζ�ȡ
    If lngId <> 0 Then
        strSQL = "Select b.Stop_Time, b.Start_Time " & vbCr & _
                 "From ��Ա֤���¼ A, " & vbCr & _
                 "     Xmltable('/root/records' Passing a.ͣ�ü�¼ Columns Stop_Time Varchar2(30) Path '/records/stop_time'," & vbCr & _
                 "              Start_Time Varchar2(30) Path '/records/start_time') B " & vbCr & _
                 "Where a.Id = [1] And Nvl(a.�Ƿ�ͣ��,0) =1 " & vbCr & _
                 "Order By To_Date(b.Stop_Time, 'yyyy-mm-dd hh24:mi:ss') Desc "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��Ա֤���¼����ͣ�ü�¼", lngId)
    End If
    With Me.lvwLogOnOff
        .ListItems.Clear
        i = 1
        If Not rsTmp Is Nothing Then
            Do While rsTmp.EOF = False
                strTmp = NVL(rsTmp!Stop_Time)
                Set ObjItem = .ListItems.Add(, "_" & i, Format(strTmp, "yyyy-mm-dd hh:MM:ss"))
                strTmp = NVL(rsTmp!Start_Time, "1")
                If CDate(strTmp) >= CDate("1990-01-01 00:00:00") Then
                    ObjItem.SubItems(1) = Format(strTmp, "yyyy-mm-dd hh:MM:ss")
                End If
                i = i + 1
                rsTmp.MoveNext
            Loop
        End If
        If .ListItems.Count > 0 Then .ListItems(1).Selected = True
    End With
    
    '����˵�״̬
    Toolbar1.Buttons("Sign").Enabled = lvwCert.ListItems.Count > 0 And mblnCAOnOff
    If lvwCert.ListItems.Count <= 0 Then
        mnuEditSignOn.Enabled = False
        mnuEditSignOff.Enabled = False
        Exit Sub
    End If
    If mblnCAOnOff = False Or lvwCert.SelectedItem.Index <= 0 Or lvwCert.SelectedItem.Index > 1 Then
        mnuEditSignOn.Enabled = False
        mnuEditSignOff.Enabled = False
        Exit Sub
    End If
    If lvwLogOnOff.ListItems.Count <= 0 Then
        'Ĭ��Ϊ����״̬��ֻ��ͣ�ò���
        mnuEditSignOn.Enabled = False
        mnuEditSignOff.Enabled = True
    Else
        strTmp = (lvwLogOnOff.ListItems(1).SubItems(1))
        mnuEditSignOn.Enabled = strTmp = ""
        mnuEditSignOff.Enabled = Not mnuEditSignOn.Enabled
    End If
    
    Exit Sub
    
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub

