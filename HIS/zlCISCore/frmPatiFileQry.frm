VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmPatiFileQry 
   Caption         =   "��������"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10740
   Icon            =   "frmPatiFileQry.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   10740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   WindowState     =   2  'Maximized
   Begin VB.PictureBox imgY 
      BackColor       =   &H00808080&
      Height          =   3375
      Index           =   0
      Left            =   4080
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3375
      ScaleWidth      =   45
      TabIndex        =   7
      Top             =   720
      Width           =   45
   End
   Begin MSComctlLib.ImageList iLsTree32 
      Left            =   840
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":058A
            Key             =   "Class"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":0E64
            Key             =   "Attr"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilstbrMain 
      Left            =   1200
      Top             =   5800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   22
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":117E
            Key             =   "Ԥ��"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":139A
            Key             =   "��ӡ"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":15B6
            Key             =   "����"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":17D2
            Key             =   "����"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":19EE
            Key             =   "�޸�"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":1C0A
            Key             =   "ɾ��"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":1E26
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":2042
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":225E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":2478
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":2694
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":28B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":2AD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":2CEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":2F0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":3604
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":3CFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":3F18
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":4132
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":48AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":4AC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":4CE0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilstbrMainHot 
      Left            =   3000
      Top             =   5920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   22
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":4EFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":511A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":533A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":555A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":577A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":599A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":5BBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":5DDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":5FFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":6214
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":6434
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":6654
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":6874
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":6A8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":6CAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":73A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":7AA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":7CBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":7ED6
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":8650
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":886A
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":8A84
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrMain 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10740
      _ExtentX        =   18944
      _ExtentY        =   1270
      BandCount       =   1
      _CBWidth        =   10740
      _CBHeight       =   720
      _Version        =   "6.7.8988"
      Child1          =   "tbrMain"
      MinHeight1      =   660
      Width1          =   9000
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrMain 
         Height          =   660
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   10620
         _ExtentX        =   18733
         _ExtentY        =   1164
         ButtonWidth     =   820
         ButtonHeight    =   1164
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ilstbrMain"
         HotImageList    =   "ilstbrMainHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   11
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ԥ��"
               Key             =   "Ԥ��"
               Object.ToolTipText     =   "��ӡԤ��ҽ����"
               Object.Tag             =   "Ԥ��"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "��ӡ"
               Object.ToolTipText     =   "��ӡҽ����"
               Object.Tag             =   "��ӡ"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "ɾ��"
               Key             =   "ɾ������"
               Description     =   "����"
               Object.ToolTipText     =   "ɾ����ǰ����"
               Object.Tag             =   "ɾ��"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "�鵵"
               Key             =   "�鵵"
               Description     =   "����"
               Object.ToolTipText     =   "�������鵵����"
               Object.Tag             =   "�鵵"
               ImageIndex      =   19
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "���ҵ�ǰ�ƵĲ���"
               Object.Tag             =   "����"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "Split_4"
               Description     =   "����"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "�鿴"
               Key             =   "�鿴"
               Description     =   "����"
               Object.ToolTipText     =   "�任��ʾͼ�귽ʽ"
               Object.Tag             =   "�鿴"
               ImageIndex      =   12
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Object.Visible         =   0   'False
                     Text            =   "��ͼ��(&G)"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Сͼ��(&M)"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "�б�(&L)"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "��ϸ����(&D)"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_5"
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "��������"
               Object.Tag             =   "����"
               ImageIndex      =   13
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "�˳�"
               Object.ToolTipText     =   "�˳�ҽ������"
               Object.Tag             =   "�˳�"
               ImageIndex      =   14
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList iLsTree 
      Left            =   0
      Top             =   5000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":8C9E
            Key             =   "����"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":9238
            Key             =   "סԺ"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":97D2
            Key             =   "����"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":9D6C
            Key             =   "����"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":A306
            Key             =   "����"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar prbRefresh 
      Height          =   200
      Left            =   2280
      TabIndex        =   6
      Top             =   6840
      Visible         =   0   'False
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   7365
      Width           =   10740
      _ExtentX        =   18944
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPatiFileQry.frx":A8A0
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11324
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin MSComctlLib.ListView lvwItem 
      Height          =   2295
      Left            =   360
      TabIndex        =   0
      Top             =   1680
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   4048
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDragMode     =   1
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "iLsTree"
      SmallIcons      =   "iLsTree"
      ColHdrIcons     =   "iLsTree"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDragMode     =   1
      NumItems        =   0
   End
   Begin VB.PictureBox picFile 
      Height          =   3735
      Left            =   4800
      ScaleHeight     =   3675
      ScaleWidth      =   5475
      TabIndex        =   4
      Top             =   1440
      Width           =   5535
      Begin zl9CISCore.ctrlPatientFile ProFile1 
         Height          =   5175
         Left            =   600
         TabIndex        =   5
         Top             =   120
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   9128
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFind 
         Caption         =   "���Ҳ���(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintSet 
         Caption         =   "��ӡ����(&S)"
      End
      Begin VB.Menu mnuPreview 
         Caption         =   "��ӡԤ��(&V)"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "��ӡ(&P)"
      End
      Begin VB.Menu mnuExcel 
         Caption         =   "�����&Excel"
      End
      Begin VB.Menu mnuFile_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_Preview 
         Caption         =   "����Ԥ��(&L)"
      End
      Begin VB.Menu mnuFile_Print 
         Caption         =   "������ӡ(&Y)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuParamSet 
         Caption         =   "��������(&M)"
         Shortcut        =   {F12}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuPatiRec 
      Caption         =   "����(&R)"
      Visible         =   0   'False
      Begin VB.Menu mnuOrder_Edit 
         Caption         =   "�޸Ĳ���(&E)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOrder_Delete 
         Caption         =   "ɾ������(&D)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOrder_2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOrder_File 
         Caption         =   "�����鵵(&F)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOrder_Undo 
         Caption         =   "��������(&U)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOrder_Print 
         Caption         =   "������ӡ(&P)"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "�鿴(&V)"
      Begin VB.Menu mnuToolbar 
         Caption         =   "������(&T)"
         Begin VB.Menu mnuToolbarStand 
            Caption         =   "��׼��ť(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuToolbarText 
            Caption         =   "�ı���ǩ(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuStatus 
         Caption         =   "״̬��(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu v1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuIconOrder 
         Caption         =   "�鿴��ʽ(&I)"
         Visible         =   0   'False
         Begin VB.Menu mnuIcon 
            Caption         =   "��ͼ��(&G)"
            Index           =   0
         End
         Begin VB.Menu mnuIcon 
            Caption         =   "Сͼ��(&M)"
            Index           =   1
         End
         Begin VB.Menu mnuIcon 
            Caption         =   "�б�(&L)"
            Index           =   2
         End
         Begin VB.Menu mnuIcon 
            Caption         =   "��ϸ����(&D)"
            Checked         =   -1  'True
            Index           =   3
         End
      End
      Begin VB.Menu v7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewInfo 
         Caption         =   "������Ϣ(&I)"
         Shortcut        =   ^I
         Visible         =   0   'False
      End
      Begin VB.Menu mnuViewHist 
         Caption         =   "��ʷ����(&A)"
      End
      Begin VB.Menu v6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "ˢ��(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuHelpTitle 
         Caption         =   "��������(&H)"
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
      Begin VB.Menu h1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)..."
      End
   End
End
Attribute VB_Name = "frmPatiFileQry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strPrivs As String       '�û����б�����ľ���Ȩ��

'��ѯ����������������||��������||�Ա�||��С����||�������||ҽ��||��������||��������||��������
Private strQuery As String
Private WithEvents objParentForm As Form
Attribute objParentForm.VB_VarHelpID = -1

Public Sub ShowMe(frmParent As Object, Optional ByVal ModalWindow As Boolean = True)
    On Error Resume Next
    Set objParentForm = frmParent
    Me.Show IIf(ModalWindow, 1, 0), frmParent
End Sub

Private Sub Form_Activate()
    If Me.Tag = "" Then Exit Sub
    
    Me.Tag = ""
    If Len(strQuery) > 0 Then
        ListItem
    End If
    
    If lvwItem.ListItems.Count > 0 Then lvwItem.ListItems(1).Selected = True: lvwItem_ItemClick lvwItem.SelectedItem
End Sub

Private Sub mnuExcel_Click()
    zlRptPrint 3
End Sub

Private Sub mnuFile_Preview_Click()
    Dim frmPreview As frmCasePrint
    Dim FileID As Long, PatientID As String, CheckID As Variant, FileType As Integer
    Dim rsTmp As New ADODB.Recordset
    
    Dim intPage As Integer
    
    If Me.lvwItem.SelectedItem Is Nothing Then Exit Sub
    FileID = Mid(Me.lvwItem.SelectedItem.Key, 4)
    On Error Resume Next
    zlDatabase.OpenRecordset rsTmp, "Select ����ID,��ҳID,�Һŵ�,�������� From ���˲�����¼ Where ID=" & FileID, Me.Caption
    If rsTmp.EOF Then Exit Sub
    PatientID = rsTmp(0): FileType = rsTmp(3)
        
    Set frmPreview = New frmCasePrint
    PrintOutCase Me, frmPreview, FileType, True, -1 * FileID, CLng(PatientID), "", False, 0, 1
    frmPreview.Preview Me, FileType, True, -1 * FileID, CLng(PatientID), "", False, 0, 1
End Sub

Private Sub mnuFile_Print_Click()
    Dim frmPreview As frmCasePrint
    Dim FileID As Long, PatientID As String, CheckID As Variant, FileType As Integer
    Dim rsTmp As New ADODB.Recordset
    
    Dim intPage As Integer
    
    If Me.lvwItem.SelectedItem Is Nothing Then Exit Sub
    FileID = Mid(Me.lvwItem.SelectedItem.Key, 4)
    On Error Resume Next
    zlDatabase.OpenRecordset rsTmp, "Select ����ID,��ҳID,�Һŵ�,�������� From ���˲�����¼ Where ID=" & FileID, Me.Caption
    If rsTmp.EOF Then Exit Sub
    PatientID = rsTmp(0): FileType = rsTmp(3)
        
    intPage = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��ӡ����", "ֽ��", Printer.PaperSize)
    If IsWindowsNT And intPage = 256 Then DelCustomPaper
    
    If Not InitPrint(Me) Then
        MsgBox "��ӡ����ʼ��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    PrintOutCase Me, Printer, FileType, True, -1 * FileID, CLng(PatientID), "", False, 0, 1
    'WinNT�Զ���ֽ�Ŵ���
    If IsWindowsNT And intPage = 256 Then DelCustomPaper

    Call InitPrint(Me)
End Sub

Private Sub mnuFind_Click()
    Dim strTmp As String
    
    strTmp = strQuery
    frmPatiFileQry1.GetQueryString Me, strTmp
    If Len(strTmp) > 0 Then
        strQuery = strTmp
        
        ListItem
        If lvwItem.ListItems.Count > 0 Then lvwItem.ListItems(1).Selected = True
        lvwItem_ItemClick lvwItem.SelectedItem
    End If
End Sub

Private Sub mnuPreview_Click()
    zlRptPrint 2
End Sub

Private Sub mnuPrint_Click()
    zlRptPrint 1
End Sub

Private Sub mnuPrintSet_Click()
    frmPrintSet.Show vbModal, Me
End Sub

Private Sub mnuRefresh_Click()
    On Error Resume Next
    
    ListItem
End Sub

Private Sub objParentForm_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub picFile_Resize()
    On Error Resume Next
    With ProFile1
        .Left = 0: .Top = 0
        .Width = picFile.ScaleWidth
        .Height = picFile.ScaleHeight
        
        If .Width > picFile.ScaleWidth Then Me.Width = .Width
        If .Height > picFile.ScaleHeight Then Me.Height = .Height + picFile.Top
    End With
End Sub

Private Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Ԥ��"
            mnuPreview_Click
        Case "��ӡ"
            mnuPrint_Click
        Case "����"
            mnuFind_Click
        Case "����"
            mnuHelpTitle_Click
        Case "�˳�"
            mnuExit_Click
    End Select
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    With Me.lvwItem.ColumnHeaders
        .Clear
        .Add , "_����", "����", 2000
        .Add , "_ID", "ID", 0
        .Add , "_����", "����", 1200
        .Add , "_����", "����", 800
        .Add , "_�Ա�", "�Ա�", 500
    End With
    With Me.lvwItem
        .ColumnHeaders("_����").Position = 3
        .SortKey = .ColumnHeaders("_����").Index - 1
        .SortOrder = lvwAscending
    End With
    
    Call RestoreWinState(Me, App.ProductName)
    
    '---------Ȩ�޿���-------------
    strPrivs = gstrPrivs
    
    '��ȡ����Ĳ�ѯ��������
    strQuery = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "��ѯ����", "")
    ShowQryString strQuery
    
    Me.Tag = "Loading"
End Sub

Private Sub Form_Resize()
    Dim lngTools As Single, lngStatus As Single
    
    If WindowState = 1 Then Exit Sub
    lngTools = IIf(Me.cbrMain.Visible, Me.cbrMain.Height, 0)
    lngStatus = IIf(Me.stbThis.Visible, Me.stbThis.Height, 0)
    
    On Error Resume Next
    imgY(0).Top = lngTools
    imgY(0).Height = Me.ScaleHeight - lngStatus - imgY(0).Top
    
    With lvwItem
        .Left = 0
        .Top = imgY(0).Top
        .Width = imgY(0).Left
        .Height = imgY(0).Height
    End With
    With picFile
        .Left = imgY(0).Left + imgY(0).Width: .Top = imgY(0).Top
        .Height = Me.ScaleHeight - lngStatus - .Top: .Width = Me.ScaleWidth - .Left
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '�����ѯ��������
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "��ѯ����", strQuery
    
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub imgY_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    On Error Resume Next
    imgY(Index).Left = imgY(Index).Left + x
End Sub

Private Sub imgY_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    On Error Resume Next
    Select Case Index
        Case 0
            If imgY(0).Left < 2000 Then imgY(0).Left = 2000
            If Me.ScaleWidth - imgY(0).Left < 4000 Then imgY(0).Left = Me.ScaleWidth - 4000
    End Select

    Form_Resize
End Sub

Private Sub lvwItem_DblClick()
    If lvwItem.SelectedItem Is Nothing Then Exit Sub
End Sub

Private Sub lvwItem_ItemClick(ByVal Item As MSComctlLib.ListItem)
    ShowMenu
    
    Me.MousePointer = vbHourglass
    BeginShowProgress "��ʾ������"
    If Item Is Nothing Then
        ProFile1.ShowFile "", , , , -1, , , Me.prbRefresh '�����������
    Else
        ProFile1.ShowFile Mid(Item.Key, 4), , , , , , , Me.prbRefresh
    End If
    Me.prbRefresh.Visible = False
    Me.MousePointer = vbDefault
    
    ShowQryString strQuery
End Sub

Private Sub lvwItem_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lvwItem
        .SortKey = ColumnHeader.Index - 1: .SortOrder = IIf(.SortOrder = lvwDescending, lvwAscending, lvwDescending)
    End With
End Sub

Private Sub mnuHelpAbout_Click()
    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
End Sub

Private Sub mnuHelpTitle_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hwnd)
End Sub

Private Sub mnuIcon_Click(Index As Integer)
'    lvwItem.View = Index
'    SetViewCheck lvwItem.View
End Sub

Private Sub mnuStatus_Click()
    Me.mnuStatus.Checked = Not Me.mnuStatus.Checked
    Me.stbThis.Visible = Me.mnuStatus.Checked
    Form_Resize
End Sub

Private Sub mnuToolbarStand_Click()
    Me.mnuToolbarStand.Checked = Not Me.mnuToolbarStand.Checked
    Me.cbrMain.Visible = Me.mnuToolbarStand.Checked
    Form_Resize
End Sub

Private Sub mnuToolbarText_Click()
    Dim i As Integer
    Me.mnuToolbarText.Checked = Not Me.mnuToolbarText.Checked
    If Me.mnuToolbarText.Checked Then
        For i = 1 To Me.tbrMain.Buttons.Count
            Me.tbrMain.Buttons(i).Caption = Me.tbrMain.Buttons(i).Tag
        Next
    Else
        For i = 1 To Me.tbrMain.Buttons.Count
            Me.tbrMain.Buttons(i).Caption = ""
        Next
    End If
    Me.cbrMain.Bands(1).MINHEIGHT = Me.tbrMain.ButtonHeight
    Form_Resize
End Sub

Private Sub mnuViewHist_Click()
    On Error Resume Next
    
    If lvwItem.SelectedItem Is Nothing Then Exit Sub
    Call frmRcdAnalyse.ShowMe(1, Me, CLng(lvwItem.SelectedItem.SubItems(1)))
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub tbrMain_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Text
    Case "��ͼ��(&G)"
        mnuIcon_Click 0
    Case "Сͼ��(&M)"
        mnuIcon_Click 1
    Case "�б�(&L)"
        mnuIcon_Click 2
    Case "��ϸ����(&D)"
        mnuIcon_Click 3
    End Select
End Sub

Private Sub ListItem()
    Dim rsTmp As New ADODB.Recordset
    Dim tmpItem As MSComctlLib.ListItem
    Dim iNum As Long
    Dim strWhereClause As String
    Dim aQryString() As String
    
    lvwItem.ListItems.Clear
    
    strWhereClause = ""
    If Len(strQuery) > 0 Then
        aQryString = Split(strQuery, "||")
        
        If Len(aQryString(0)) > 0 And aQryString(0) <> "0" Then strWhereClause = strWhereClause + " And a.��������=" + aQryString(0)
        If Len(aQryString(1)) > 0 Then strWhereClause = strWhereClause + " And b.���� Like '%" + Replace(aQryString(1), "'", "''") + "%'"
        If Len(aQryString(2)) > 0 And aQryString(2) <> "0" Then strWhereClause = strWhereClause + " And b.�Ա�='" + aQryString(2) + "'"
        If Len(aQryString(3)) > 0 Then strWhereClause = strWhereClause + " And b.����>=" + aQryString(3)
        If Len(aQryString(4)) > 0 Then strWhereClause = strWhereClause + " And b.����<=" + aQryString(4)
        If Len(aQryString(5)) > 0 Then strWhereClause = strWhereClause + " And a.��д�� Like '%" + Replace(aQryString(5), "'", "''") + "%'"
        If Len(aQryString(6)) > 0 Then strWhereClause = strWhereClause + " And a.��д����>=To_Date('" + aQryString(6) + "','yyyy-mm-dd')"
        If Len(aQryString(7)) > 0 Then strWhereClause = strWhereClause + " And a.��д����<=To_Date('" + aQryString(7) + " 23:59:59','yyyy-mm-dd hh24:mi:ss')"
'        If Len(aQryString(8)) > 0 Then strWhereClause = strWhereClause + " And (d.���� Like '%" + aQryString(8) + "%' or " + _
'            "e.�������� Like '%" + aQryString(8) + "%')"
            
        If Len(strWhereClause) > 0 Then strWhereClause = Mid(strWhereClause, 6)
    End If
    
    On Error Resume Next
    
    Me.MousePointer = vbHourglass
    BeginShowProgress "���ڼ�����"
    If Len(aQryString(8)) = 0 Then
        On Error GoTo QryError
        zlDatabase.OpenRecordset rsTmp, "Select a.ID,a.����ID,To_Char(a.��д����,'yyyy-mm-dd'),b.����,nvl(b.�Ա�,' ')," + _
            "decode(a.��������,1,'����',2,'סԺ',3,'����',4,'����','����'),a.�������� From ���˲�����¼ a,������Ϣ b Where " + _
            IIf(Len(strWhereClause) = 0, "", strWhereClause + " And ") + "a.��������>0 And a.����ID=b.����ID Order By a.��д����,a.����ID", Me.Caption
    Else
        On Error GoTo QryError
        zlDatabase.OpenRecordset rsTmp, "Select Distinct a.ID,a.����ID,To_Char(a.��д����,'yyyy-mm-dd'),b.����,nvl(b.�Ա�,' ')," + _
            "decode(a.��������,1,'����',2,'סԺ',3,'����',4,'����','����'),a.�������� From ���˲�����¼ a,������Ϣ b,���˲������� c Where " + _
            IIf(Len(strWhereClause) = 0, "", strWhereClause + " And ") + "a.��������>0 And a.����ID=b.����ID And a.ID=c.������¼ID And c.ID In " + _
            "(Select ����ID From ���˲����ı��� Where ���� Like '%" + Replace(aQryString(8), "'", "''") + _
            "%' Union Select ����ID From ���˲��������� Where �ؼ��� In (2,Null) And �������� Like '%" + Replace(aQryString(8), "'", "''") + _
            "%') Order By To_Char(a.��д����,'yyyy-mm-dd'),b.����", Me.Caption
    End If
    
    prbRefresh.Value = 50
    iNum = 0
    Do While Not rsTmp.EOF
        Set tmpItem = lvwItem.ListItems.Add(, "Key" & rsTmp(0), rsTmp(6))
        tmpItem.SubItems(Me.lvwItem.ColumnHeaders("_ID").Index - 1) = rsTmp(1)
        tmpItem.SubItems(Me.lvwItem.ColumnHeaders("_����").Index - 1) = rsTmp(2)
        tmpItem.SubItems(Me.lvwItem.ColumnHeaders("_����").Index - 1) = rsTmp(3)
        tmpItem.SubItems(Me.lvwItem.ColumnHeaders("_�Ա�").Index - 1) = rsTmp(4)
        tmpItem.Icon = CStr(rsTmp(5)): tmpItem.SmallIcon = CStr(rsTmp(5))
        
        iNum = iNum + 1
        prbRefresh.Value = 50 + CLng(50 * iNum / rsTmp.RecordCount)
        rsTmp.MoveNext
    Loop
    Me.stbThis.Panels(3).Text = "������¼��" + IIf(iNum = 0, "��", iNum & "��")
    prbRefresh.Visible = False
    Me.MousePointer = vbDefault
    
    ShowQryString strQuery
    
    ShowMenu
    Exit Sub
QryError:
    If ErrCenter = 1 Then Resume
    SaveErrLog
End Sub

Private Sub ShowMenu()
'    mnuOrder_Jz.Enabled = False
'    mnuOrder_Qx.Enabled = False
'    mnuOrder_Wc.Enabled = False
'    mnuOrder_Hf.Enabled = False
'    Select Case iPatiType
'        Case 0
'            If Not lvwItem(0).SelectedItem Is Nothing Then mnuOrder_Jz.Enabled = True
'        Case 1
'            If Not lvwItem(1).SelectedItem Is Nothing Then
'                mnuOrder_Qx.Enabled = True
'                mnuOrder_Wc.Enabled = True
'            End If
'        Case 2
'            If Not lvwItem(2).SelectedItem Is Nothing Then mnuOrder_Hf.Enabled = True
'    End Select
'
'    tbrMain.Buttons("����").Enabled = mnuOrder_Jz.Enabled
'    tbrMain.Buttons("���").Enabled = mnuOrder_Wc.Enabled
End Sub

Private Sub zlRptPrint(ByVal bytMode As Byte)
    '-------------------------------------------------
    '����:��¼���ӡ
    '����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    '-------------------------------------------------
    Dim objPrint As New zlPrintLvw
    On Error Resume Next
    Set objPrint.Body.objData = Me.lvwItem
    objPrint.Title.Text = "�����嵥"
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrViewLvw objPrint, bytMode
    Else
        zlPrintOrViewLvw objPrint, bytMode
    End If
End Sub
'��ʾ��ѯ����
Private Sub ShowQryString(ByVal strQry As String)
'��ѯ����������������||��������||�Ա�||��С����||�������||ҽ��||��������||��������||��������
    Dim aQryString() As String
    
    If Len(Trim(strQry)) = 0 Then Me.stbThis.Panels(2).Text = "��ѯ����  δ����": Exit Sub
    
    aQryString = Split(strQry, "||")
    With Me.stbThis.Panels(2)
        .Text = ""
        If Len(aQryString(0)) > 0 And aQryString(0) <> "0" Then
            Select Case aQryString(0)
                Case 1
                    .Text = .Text + "���ﲡ����"
                Case 2
                    .Text = .Text + "סԺ������"
                Case 3
                    .Text = .Text + "�����¼��"
                Case 4
                    .Text = .Text + "������飬"
                Case 5
                    .Text = .Text + "���Ƶ��ݣ�"
            End Select
        End If
        If Len(aQryString(1)) > 0 Then .Text = .Text + "������" + aQryString(1) + "��"
        If Len(aQryString(2)) > 0 And aQryString(2) <> "0" Then .Text = .Text + "�Ա�" + aQryString(2) + "��"
        If Len(aQryString(3)) > 0 Then .Text = .Text + "���䣺" + aQryString(3) + "��"
        If Len(aQryString(4)) > 0 Then
            If Len(aQryString(3)) = 0 Then .Text = .Text + "���䣺��"
            .Text = .Text + aQryString(4) + "��"
        Else
            If Len(aQryString(3)) > 0 Then .Text = .Text + "��"
        End If
        If Len(aQryString(5)) > 0 Then .Text = .Text + "ҽ����" + aQryString(5) + "��"
        If Len(aQryString(6)) > 0 Then .Text = .Text + "���ڣ�" + aQryString(6) + "��"
        If Len(aQryString(7)) > 0 Then
            If Len(aQryString(6)) = 0 Then .Text = .Text + "���ڣ���"
            .Text = .Text + aQryString(7) + "��"
        End If
        If Len(aQryString(8)) > 0 Then .Text = .Text + "���ݰ�����" + aQryString(8) + "��"
        
        If Len(.Text) > 0 Then
            .Text = "��ѯ����  " + Mid(.Text, 1, Len(.Text) - 1)
        Else
            .Text = "��ѯ����  δ����"
        End If
    End With
End Sub

Private Sub BeginShowProgress(ByVal strCaption As String)
    With prbRefresh
        .Left = stbThis.Panels(2).Left + Me.TextWidth(strCaption) + 200
        .Top = stbThis.Top + (stbThis.Height - .Height) / 2
        .Width = stbThis.Panels(2).Width + stbThis.Panels(2).Left - .Left
        
        stbThis.Panels(2).Text = strCaption
        .Visible = True: Me.Refresh
    End With
End Sub

Private Sub tbrMain_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu Me.mnuToolbar, 2
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

