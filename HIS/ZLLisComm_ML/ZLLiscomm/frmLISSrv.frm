VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmLISSrv 
   AutoRedraw      =   -1  'True
   Caption         =   "�������ݽ���"
   ClientHeight    =   7305
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   10995
   Icon            =   "frmLISSrv.frx":0000
   KeyPreview      =   -1  'True
   ScaleHeight     =   7305
   ScaleWidth      =   10995
   StartUpPosition =   2  '��Ļ����
   WindowState     =   1  'Minimized
   Begin VB.Timer timConn 
      Interval        =   30000
      Left            =   9930
      Top             =   3030
   End
   Begin Zl9LISComm.ctrlComm DevComm 
      Height          =   495
      Index           =   0
      Left            =   8280
      TabIndex        =   10
      Top             =   1785
      Visible         =   0   'False
      Width           =   510
      _ExtentX        =   900
      _ExtentY        =   873
   End
   Begin VB.PictureBox picTmp 
      Height          =   1080
      Left            =   8535
      ScaleHeight     =   1020
      ScaleWidth      =   1170
      TabIndex        =   9
      Top             =   5340
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.PictureBox picIcon 
      Height          =   285
      Left            =   8790
      ScaleHeight     =   225
      ScaleWidth      =   255
      TabIndex        =   8
      Top             =   4170
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Frame fraUD_s 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   30
      Left            =   240
      MousePointer    =   7  'Size N S
      TabIndex        =   4
      Top             =   3480
      Width           =   8745
   End
   Begin MSComctlLib.ListView lvwResult 
      Height          =   2730
      Left            =   150
      TabIndex        =   6
      Top             =   3720
      Width           =   8280
      _ExtentX        =   14605
      _ExtentY        =   4815
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvwLISRec 
      Height          =   2415
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   4260
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   735
      Top             =   2190
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
            Picture         =   "frmLISSrv.frx":08CA
            Key             =   "_0"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":0E64
            Key             =   "_1"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   1376
      BandCount       =   2
      _CBWidth        =   10995
      _CBHeight       =   780
      _Version        =   "6.7.8988"
      Child1          =   "tbrMain"
      MinWidth1       =   4995
      MinHeight1      =   720
      NewRow1         =   0   'False
      Caption2        =   "����"
      Child2          =   "cboDev"
      MinWidth2       =   3795
      MinHeight2      =   300
      Width2          =   5940
      NewRow2         =   0   'False
      AllowVertical2  =   0   'False
      Begin VB.ComboBox cboDev 
         Height          =   300
         Left            =   5805
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   240
         Width           =   5100
      End
      Begin MSComctlLib.Toolbar tbrMain 
         Height          =   720
         Left            =   165
         TabIndex        =   2
         Top             =   30
         Width           =   4995
         _ExtentX        =   8811
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
            NumButtons      =   8
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ԥ��"
               Key             =   "Ԥ��"
               Object.ToolTipText     =   "Ԥ��"
               Object.Tag             =   "Ԥ��"
               ImageKey        =   "Ԥ��"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "��ӡ"
               Object.ToolTipText     =   "��ӡ"
               Object.Tag             =   "��ӡ"
               ImageKey        =   "��ӡ"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "���ӵ�ǰ����׼����������"
               Object.Tag             =   "����"
               ImageKey        =   "����"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�Ͽ�"
               Key             =   "�Ͽ�"
               Object.ToolTipText     =   "�뵱ǰ�����Ͽ�����"
               Object.Tag             =   "�Ͽ�"
               ImageKey        =   "�Ͽ�"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageKey        =   "����"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "�˳�"
               Key             =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageKey        =   "�˳�"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   6945
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmLISSrv.frx":13FE
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11748
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
   Begin VB.Frame fraLR_s 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6045
      Left            =   3330
      MousePointer    =   9  'Size W E
      TabIndex        =   3
      Top             =   750
      Visible         =   0   'False
      Width           =   30
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   1335
      Top             =   2160
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
            Picture         =   "frmLISSrv.frx":1C92
            Key             =   "Ԥ��"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":1EAC
            Key             =   "��ӡ"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":20C6
            Key             =   "����"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":22E0
            Key             =   "�˳�"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":24FA
            Key             =   "��¼"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":2BF4
            Key             =   "����"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":32EE
            Key             =   "���"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":39E8
            Key             =   "����"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":40E2
            Key             =   "����"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":47DC
            Key             =   "�ķ�"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":4ED6
            Key             =   "ɾ��"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":55D0
            Key             =   "����"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":5CCA
            Key             =   "�޸�"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":63C4
            Key             =   "ɾ��"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":6ABE
            Key             =   "����"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":71B8
            Key             =   "����"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":7932
            Key             =   "�Ͽ�"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   1935
      Top             =   2160
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
            Picture         =   "frmLISSrv.frx":80AC
            Key             =   "Ԥ��"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":82C6
            Key             =   "��ӡ"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":84E0
            Key             =   "����"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":86FA
            Key             =   "�˳�"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":8914
            Key             =   "��¼"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":900E
            Key             =   "����"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":9708
            Key             =   "���"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":9E02
            Key             =   "����"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":A4FC
            Key             =   "����"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":ABF6
            Key             =   "�ķ�"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":B2F0
            Key             =   "ɾ��"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":B9EA
            Key             =   "����"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":C0E4
            Key             =   "�޸�"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":C7DE
            Key             =   "ɾ��"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":CED8
            Key             =   "����"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":D5D2
            Key             =   "����"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":DD4C
            Key             =   "�Ͽ�"
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock WinsockS 
      Left            =   9840
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFilePrintSet 
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
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSetup 
         Caption         =   "��������(&S)"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFtpSet 
         Caption         =   "FTP����(&F)"
      End
      Begin VB.Menu mnuFile_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "����(&C)"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "�Ͽ�(&D)"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuFile_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDevExp 
         Caption         =   "������������(&P)"
      End
      Begin VB.Menu mnuFile_4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileQuit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "�鿴(&V)"
      Begin VB.Menu mnuViewTool 
         Caption         =   "������(&T)"
         Begin VB.Menu mnuViewToolItem 
            Caption         =   "��׼��ť(&S)"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuViewToolItem 
            Caption         =   "����ѡ��(&D)"
            Checked         =   -1  'True
            Index           =   1
            Visible         =   0   'False
         End
         Begin VB.Menu mnuViewTool_1 
            Caption         =   "-"
            Visible         =   0   'False
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
      Begin VB.Menu mnuView_4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewCharge 
         Caption         =   "ֻ��ʾ�Ѿ��շѵĲ���(&P)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuView_2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuViewFilter 
         Caption         =   "���ݹ���(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuView_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewComm 
         Caption         =   "ͨѶ���(&C)"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuReLoad 
         Caption         =   "����ͨѶ(&L)"
      End
      Begin VB.Menu mnuView_5 
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
Attribute VB_Name = "frmLISSrv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit 'Ҫ���������
Private Const COLOR_LOST = &HFFEBD7
Private Const COLOR_FOCUS = &HFFCC99

Private lngPort As Long
Private lngErrCounts As Long, strWhere As String, strBeginDate As String
Private strDevIDs As String '���ӵ��豸ID

'********************���ظ���ʦվ����Ϣ*****************************
Private Const strSend_Refresh = "Refresh"      '�ѱ������ݿ���ˢ��
Private Const strSend_True = "True"            '�Ѳ����ɹ�
Private Const strSend_False = "False"          '����ʧ��
Private Const strSend_AutoCompute = "AutoCompute" '�������
'*******************************************************************

Private WithEvents frmView As frmViewComm
Attribute frmView.VB_VarHelpID = -1
Private mblnOwner As Boolean '�Ƿ������ߵ�¼
Private mfsoTmp As New FileSystemObject  '�ļ�����

Private Sub cboDev_Click()
    ListSeq strWhere
    ShowMenu
End Sub

Private Sub DevComm_DevDecode(Index As Integer, ByVal commport As String, ByVal str��� As String)
    Dim strCOM As String
    If frmView Is Nothing Then Exit Sub
    
    If InStr(commport, ".") <= 0 Then strCOM = "COM" & Val(commport)
    
    If InStr(cboDev.List(cboDev.ListIndex), strCOM) > 0 Then
        If str��� <> "" Then
         '  ��ʾ�յ��Ľ������
            Call frmView.ShowDecode(0, str���)
        End If
    End If
End Sub

Private Sub DevComm_DevRefresh(Index As Integer, ByVal lngID As Long)
    '����ˢ����Ϣ��LISWORK
    If lngID <> 0 Then
        Me.WinsockS.SendData Me.WinsockS.LocalIP & ";" & strSend_Refresh & ";" & lngID
    End If
End Sub

Private Sub DevComm_ItemUnknown(Index As Integer, ByVal commport As String, ByVal strItems As String)
    Dim strCOM As String
    If frmView Is Nothing Then Exit Sub
    
    If InStr(commport, ".") <= 0 Then strCOM = "COM" & Val(commport)
    
    If InStr(cboDev.List(cboDev.ListIndex), strCOM) > 0 Then
        If strItems <> "" Then
         '  ��ʾ�յ���δ֪��
            Call frmView.ShowDecode(1, strItems)
        End If
    End If
End Sub

Private Sub DevComm_ReturnCompute(Index As Integer, ByVal strReturn As String)
    
    If DevComm(Index) Is Nothing Then Exit Sub
    If blnDataReceived Then Exit Sub
    blnDataReceived = True
    With Me.WinsockS
        .SendData .LocalIP & ";" & "AutoQCCompute|" & strReturn
    End With
    blnDataReceived = False
End Sub

Private Sub Form_Activate()
    ListSeq strWhere
    ShowMenu
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        Me.WindowState = vbMinimized
    End If
End Sub

Private Sub fraUD_s_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    
    fraUD_s.BackColor = RGB(0, 0, 0)
    On Error Resume Next
    If fraUD_s.Top + y < 2000 Then
        fraUD_s.Top = 2000
    ElseIf Me.ScaleHeight - fraUD_s.Top - y < 4000 Then
        fraUD_s.Top = Me.ScaleHeight - 4000
    Else
        fraUD_s.Top = fraUD_s.Top + y
    End If
End Sub

Private Sub fraUD_s_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    fraUD_s.BackColor = Me.BackColor
    Form_Resize
End Sub

Private Sub frmView_CloseWindow()
    mnuViewComm.Checked = False
End Sub

Private Sub lvwLISRec_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call gobjControl.LvwSortColumn(lvwLISRec, ColumnHeader.Index)
End Sub

Private Sub lvwLISRec_ItemClick(ByVal Item As MSComctlLib.ListItem)
    ListResult
End Sub

Private Sub mnuDevExp_Click()
    Dim strFile As String
    strFile = ExpLisDevData
    If strFile <> "" Then
        MsgBox "�ѵ����� " & strFile & "��"
    End If
    
End Sub

Private Sub mnuFileClose_Click()
    If Me.cboDev.ListIndex = -1 Then Exit Sub
    
'    Me.DevComm(Me.cboDev.ListIndex + 1).ClosePort
    ShowMenu
End Sub

Private Sub mnuFileOpen_Click()
    If Me.cboDev.ListIndex = -1 Then Exit Sub
    If gblnFromDB Then
        mMakeNoRule = gobjDatabase.GetPara("�걾������ɹ���", glngSys, 1208, "��  ��")
    Else
        mMakeNoRule = GetSetting("ZLSOFT", "����ģ��\zl9LisWork\frmLabMain", "�걾������ɹ���", "��  ��")
    End If
'    Me.DevComm(Me.cboDev.ListIndex + 1).OpenPort
    ShowMenu
End Sub

Private Sub mnuFileSetup_Click()

    If frmParaSet.ShowMe(Me) Then
        If Not ReadPara("ResetExe") Then Unload Me
    End If
End Sub

Private Sub mnuFtpSet_Click()
    frmFtpSet.Show vbModal
End Sub

Private Sub mnuReLoad_Click()
    Dim intTime As Integer
    Dim tsmTmp As TextStream
    Dim objWait As New clsLISComm
    On Error GoTo errH

'    For inttime = LBound(g����) To UBound(g����)
'        If mfsoTmp.FileExists(g����(inttime).ͨѶĿ¼ & "\Lock.txt") Then
'            Set tsmTmp = mfsoTmp.CreateTextFile(g����(inttime).ͨѶĿ¼ & "\Send\CloseExe.txt")
'            tsmTmp.WriteLine Format(Now, "yyyy-MM-dd HH:mm:ss")
'            tsmTmp.Close
'            Set tsmTmp = Nothing
'        End If
'    Next
    
    Call KillProc("zlLisReceiveSend.exe")
    For intTime = LBound(g����) To UBound(g����)
        If Dir(g����(intTime).ͨѶĿ¼ & "\Lock.txt") <> "" Then Kill g����(intTime).ͨѶĿ¼ & "\Lock.txt"
    Next
    '��ʱ1.5���������ӿ�
    objWait.Wait 1500
    Set objWait = Nothing
    Call ReadPara("")
    Exit Sub
errH:
    WriteLog "mnuReload", LOG_������־, Err.Number, Err.Description
End Sub

Private Sub mnuViewComm_Click()
    
    If mnuViewComm.Checked Then
        If Not frmView Is Nothing Then Unload frmView
    Else
        mnuViewComm.Checked = True
        If frmView Is Nothing Then Set frmView = New frmViewComm
        Call frmView.ShowMe(cboDev.List(cboDev.ListIndex), Me.DevComm(Me.cboDev.ListIndex + 1).CommSetting, Me.DevComm(Me.cboDev.ListIndex + 1).DevProgName)
    End If
End Sub

Private Sub mnuViewRefresh_Click()
    ListSeq strWhere
End Sub

Private Sub picIcon_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    '--------------------------------------------------------------------------------------------------
    '����:  ����ͼ��ĸ��ִ����¼�
    '--------------------------------------------------------------------------------------------------
    On Error Resume Next
    Select Case Button '
        Case vbLeftButton
            Me.Show
            Me.WindowState = vbNormal
        Case vbRightButton
            ModifyIcon picIcon.hwnd, Me.Icon, , False
            Me.PopupMenu Me.mnuFile
            ModifyIcon picIcon.hwnd, Me.Icon
    End Select '
End Sub

Private Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "����"
            mnuFileOpen_Click
        Case "�Ͽ�"
            mnuFileClose_Click
        Case "�˳�"
            mnuFileQuit_Click
        Case "��ӡ"
            mnuFilePrint_Click
        Case "Ԥ��"
            mnuFilePreview_Click
        Case "����"
            mnuHelpTitle_Click
    End Select
End Sub

Private Sub mnuFilePrintSet_Click()
'���ܣ���ӡ����
    Call gobjPrintMode.zlPrintSet
End Sub

Private Sub mnuFileExcel_Click()
'���ܣ������Excel
    Call OutputList(3)
End Sub

Private Sub mnuFilePreview_Click()
'���ܣ���ӡԤ��
    Call OutputList(2)
End Sub

Private Sub mnuFilePrint_Click()
'���ܣ���ӡ
    Call OutputList(1)
End Sub

Private Sub mnuHelpTitle_Click()
'���ܣ����ð�������
    gobjComLib.ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub mnuFileQuit_Click()
    If MsgBox("�˳��󽫲��ܽ����������ݣ��Ƿ�ȷ��Ҫ�˳���", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then
        End
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim rsTmp As ADODB.Recordset
    
    '�ͼ�ʦվ��ͨѶ�ӿ�
    With Me.WinsockS
        .Protocol = sckUDPProtocol
        .RemoteHost = "Localhost"
        .RemotePort = 1001
        .Bind 1000
    End With
    
    strBeginDate = Format(date & " " & Time, "yyyy-MM-dd hh:mm:ss")
    lngErrCounts = 0
    strWhere = ""
    
    With lvwLISRec
        With .ColumnHeaders
            .Clear
            .Add , , "��������", 1000
            .Add , , "�걾��", 800, 1
            .Add , , "�걾", 1000
            .Add , , "������Ŀ", 2500
            .Add , , "�������", 1200
            .Add , , "ҽ��", 1000
            .Add , , "����ʱ��", 2000
            .Add , , "������", 1000
            .Add , , "����ʱ��", 2000
            .Add , , "�ʿ�Ʒ", 800
        End With
        .ListItems.Add , , "Temp", , 1
        .ListItems.Clear
    End With
    With lvwResult
        With .ColumnHeaders
            .Clear
            .Add , , "������Ŀ", 2000
            .Add , , "������", 1200, 1
            .Add , , "��־", 1000
        End With
        .ListItems.Add , , "Temp", , 1
        .ListItems.Clear
    End With
    
    
    '��ȡ���ӵ��豸
    If Dir(App.Path & "\zlLisReceiveSend.exe") = "" Then
        MsgBox "ȱ��ͨѶ����zlLisReceiveSend.Exe�����������У�", vbQuestion, "zl9LisComm"
        End
    Else
        If Not ReadPara("") Then End
    End If
    

    Me.stbThis.Panels(3).Text = "����" & Format(lngErrCounts, "@@@@@@") & "��"
    
    If gblnFromDB Then
        mMakeNoRule = gobjDatabase.GetPara("�걾������ɹ���", glngSys, 1208, "��  ��")
    Else
        mMakeNoRule = GetSetting("ZLSOFT", "����ģ��\zl9LisWork\frmLabMain", "�걾������ɹ���", "��  ��")
    End If
    gstrSQL = "Select ������ from zlsystems where ���=[1] And ������=[2]"
    Set rsTmp = gobjDatabase.OpenSqlRecord(gstrSQL, Me.Caption, glngSys, UCase(gstrDbUser))
    mblnOwner = False
    If rsTmp.RecordCount > 0 Then
        mblnOwner = True
    End If

    Call mnuReLoad_Click
End Sub

Private Sub cbr_Resize()
    Call Form_Resize
End Sub

Private Sub mnuHelpAbout_Click()
    gobjComLib.ShowAbout Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = Not stbThis.Visible
    Form_Resize
End Sub

Private Sub mnuViewToolItem_Click(Index As Integer)
    Dim blnEnabled As Boolean, blnVisible As Boolean, i As Integer
    
    mnuViewToolItem(Index).Checked = Not mnuViewToolItem(Index).Checked
    cbr.Bands(Index + 1).Visible = Not cbr.Bands(Index + 1).Visible

    blnEnabled = False: blnVisible = False
    For i = 1 To cbr.Bands.Count
        'ֻ����һ��ToolBar�ɼ�,��"��ʾ�ı�"�˵��ɼ�
        If TypeName(cbr.Bands(i).Child) = "Toolbar" Then
            If cbr.Bands(i).Visible Then
                blnEnabled = True
            End If
        End If
        'ֻҪ��һ��Band�ɼ�,��CoolBar�ɼ�
        If cbr.Bands(i).Visible Then
            blnVisible = True
        End If
    Next
    mnuViewToolText.Enabled = blnEnabled
    cbr.Visible = blnVisible
    
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim i As Integer, j As Integer
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For i = 1 To cbr.Bands.Count
        If TypeName(cbr.Bands(i).Child) = "Toolbar" Then
            For j = 1 To cbr.Bands(i).Child.Buttons.Count
                cbr.Bands(i).Child.Buttons(j).Caption = IIf(mnuViewToolText.Checked, cbr.Bands(i).Child.Buttons(j).Tag, "")
            Next
            If Not mnuViewToolText.Checked Then
                cbr.Bands(i).Child.TextAlignment = tbrTextAlignBottom
            End If
            cbr.Bands(i).MinHeight = cbr.Bands(i).Child.ButtonHeight
            cbr.Bands(i).Child.Refresh
        End If
    Next
End Sub

Private Sub mnuHelpWebHome_Click()
    gobjComLib.zlHomePage hwnd
End Sub

Private Sub mnuHelpWebMail_Click()
    gobjComLib.zlMailTo hwnd
End Sub

Private Sub Form_Resize()
    Dim cbrH As Long, staH As Long, i As Long

    On Error Resume Next
    
    Select Case WindowState
        Case vbMinimized
            Me.Hide
            AddIcon picIcon.hwnd, Me.Icon
        Case Else
            RemoveIcon picIcon.hwnd
    End Select

    If WindowState = 1 Then Exit Sub
    
    cbrH = IIf(cbr.Visible, cbr.Height, 0)
    staH = IIf(stbThis.Visible, stbThis.Height, 0)
    
    With Me.fraUD_s
        If .Top > Me.ScaleHeight Then .Top = cbrH + (Me.ScaleHeight - cbrH) / 2
        .Left = 0: .Width = Me.ScaleWidth
    End With
    
    With lvwResult
        .Left = 0: .Top = fraUD_s.Top + fraUD_s.Height
        .Width = Me.ScaleWidth: .Height = Me.ScaleHeight - staH - .Top
    End With
    
    With lvwLISRec
        .Left = 0
        .Top = cbrH
        .Height = fraUD_s.Top - .Top
        .Width = Me.ScaleWidth
    End With
    
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    For i = 1 To DevComm.UBound
        Unload DevComm(i)
    Next
    RemoveIcon picIcon.hwnd
    Call gobjComLib.SaveWinState(Me, App.ProductName)
End Sub

Private Sub tbrMain_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuViewTool, 2
End Sub

Private Sub OutputList(bytStyle As Byte)
'����: ������б�
'������bytStyle=1-��ӡ,2-Ԥ��,3-�����Excel
    Dim objOut As Object
    Set objOut = CreateObject("zl9PrintMode.zlPrintLvw")
    On Error Resume Next
    
    
    If lvwLISRec.SelectedItem Is Nothing Then Exit Sub
    
    Set objOut.Body.objData = Me.lvwLISRec
    objOut.Title.Text = "�����¼"
    objOut.UnderAppItems.Add ""
    objOut.UnderAppItems.Add "ʱ�䣺" & strBeginDate & " - " & Format(date & " " & Time, "yyyy-MM-dd HH:mm:SS")
    If bytStyle = 1 Then
        bytStyle = gobjPrintMode.zlPrintAsk(objOut)
        If bytStyle <> 0 Then gobjPrintMode.zlPrintOrViewLvw objOut, bytStyle
    Else
        gobjPrintMode.zlPrintOrViewLvw objOut, bytStyle
    End If
End Sub

Private Sub ListSeq(ByVal strWhere As String)
    Dim rsTmp As New ADODB.Recordset
    Dim strCurKey As String
    Dim tmpItem As MSComctlLib.ListItem
    Dim strIDs As String '�걾��¼IDö��
    Dim aDevIDs() As String
    
    If cboDev.ListIndex = -1 Then Me.lvwLISRec.ListItems.Clear: Exit Sub
    On Error GoTo DBError
    If Not lvwLISRec.SelectedItem Is Nothing Then strCurKey = lvwLISRec.SelectedItem.Key

    Me.lvwLISRec.ListItems.Clear
    If Len(strDevIDs) > 0 Then
        aDevIDs = Split(strDevIDs, ",")
        gstrSQL = "Select Distinct A.ID,A.�걾���,A.�걾����,A.����ʱ��,A.������,A.����ʱ��,A.�Ƿ��ʿ�Ʒ," & _
            "C.����,B.ҽ������,D.����,B.����ҽ�� " & _
            "From ����걾��¼ A,����ҽ����¼ B,������Ϣ C,���ű� D " & _
            "Where A.ҽ��ID=B.ID(+) And B.����ID=C.����ID(+) And B.��������ID=D.ID(+) " & _
            " And A.����ʱ�� Between [1] And Sysdate" & _
            " And A.����ID =[2]"
        If rsTmp.State <> adStateClosed Then rsTmp.Close
        Set rsTmp = gobjDatabase.OpenSqlRecord(gstrSQL, Me.Caption, CDate(strBeginDate), CLng(aDevIDs(Me.cboDev.ListIndex)))
        Do While Not rsTmp.EOF
            Set tmpItem = lvwLISRec.ListItems.Add(, "_" & rsTmp("ID"), Nvl(rsTmp("����")))
            With tmpItem
                .SubItems(1) = Nvl(rsTmp("�걾���"))
                .SubItems(2) = Nvl(rsTmp("�걾����"))
                .SubItems(3) = Nvl(rsTmp("ҽ������"))
                .SubItems(4) = Nvl(rsTmp("����"))
                .SubItems(5) = Nvl(rsTmp("����ҽ��"))
                .SubItems(6) = Nvl(rsTmp("����ʱ��"))
                .SubItems(7) = Nvl(rsTmp("������"))
                .SubItems(8) = Nvl(rsTmp("����ʱ��"))
                .SubItems(9) = IIf(Nvl(rsTmp("�Ƿ��ʿ�Ʒ"), 0) = 0, "  ", "��")

                If .Key = strCurKey Then .Selected = True
            End With

            rsTmp.MoveNext
        Loop
    End If
    Exit Sub
DBError:
'    If gobjComLib.ErrCenter() = 1 Then Resume
    lngErrCounts = lngErrCounts + 1
    Me.stbThis.Panels(3).Text = "����" & Format(lngErrCounts, "@@@@@@") & "��"
    Call WriteLog("frmLISSrv.ListSeq", LOG_������־, Err.Number, Err.Description)
End Sub

Private Sub ListResult()
    Dim rsTmp As New ADODB.Recordset
    Dim tmpItem As MSComctlLib.ListItem
    lvwResult.ListItems.Clear
    If lvwLISRec.SelectedItem Is Nothing Then Exit Sub
    
    On Error GoTo DBError
    gstrSQL = "Select A.ID,B.������,A.������,Decode(A.�����־,Null,'����',1,'����',2,'ƫ��',3,'ƫ��') As �����־ " & _
        "From ������ͨ��� A,����������Ŀ B,����걾��¼ C " & _
        "Where A.������ĿID=B.ID And A.����걾ID=C.ID And A.��¼����=C.������ And A.����걾ID=[1]"
    Set rsTmp = gobjDatabase.OpenSqlRecord(gstrSQL, Me.Caption, CLng(Mid(lvwLISRec.SelectedItem.Key, 2)))
    Do While Not rsTmp.EOF
        Set tmpItem = lvwResult.ListItems.Add(, "_" & rsTmp("ID"), Nvl(rsTmp("������")))
        With tmpItem
            .SubItems(1) = Nvl(rsTmp("������"))
            .SubItems(2) = Nvl(rsTmp("�����־"))
        End With

        rsTmp.MoveNext
    Loop
    Exit Sub
DBError:
    lngErrCounts = lngErrCounts + 1
    Me.stbThis.Panels(3).Text = "����" & Format(lngErrCounts, "@@@@@@") & "��"
    Call WriteLog("frmLISSrv.ListResult", LOG_������־, Err.Number, Err.Description)
End Sub

Private Function ReadPara(ByVal strCmd As String) As Boolean
    '
    '��ȡ���ӵ��豸���򿪴���
    'strCmd= ResetExe -�����ӿ� CloseExe-�رսӿ�
    
    Dim aDevices As Variant, i As Integer, y As Integer
    Dim iData As Integer, blnNextIP As Boolean
    Dim strSQL As String, strSet As String, strTmp As String, varSet As Variant, lngID As Long, lngSaveAsID As Long
    Dim strCOM As String, rsTmp As ADODB.Recordset
    Dim aPorts As Variant, str�����б� As String, blnAdd As Boolean
    '����ؼ�
    On Error GoTo errH
    
    For i = 1 To Me.DevComm.Count - 1
        Unload Me.DevComm(i)
    Next
    Me.cboDev.Clear
    strDevIDs = ""

    ReDim g����(1)

    
    If gblnFromDB Then
        gblnClearData = gobjDatabase.GetPara("��ս�����־", glngSys, 1208, 1)
        '�����ݿ������
        '���ø�ʽ:  ����id,����,COM��,������,����λ,У��λ,ֹͣλ,����,TCPIP�˿�,IP��ַ,�ַ�ģʽ,���Ϊ������ID,����,�Զ�Ӧ��,�ɷ��Ѻ˱걾,ͨѶĿ¼,�Զ������,�Զ������ʿ�,���Ϊͨ����
       
        strSet = Trim(gobjDatabase.GetPara("������������", glngSys, 1208, ""))
        If strSet = "" Then
            ShowMenu
            'û�������κ�����
        Else
            varSet = Split(strSet, ";")
            
            ReDim g����(UBound(varSet))
            For i = LBound(g����) To UBound(g����)
                g����(i).ID = 0
                g����(i).IP = ""
                g����(i).IP�˿� = 6666
                g����(i).SaveAsID = 0
                g����(i).������ = 9600
                g����(i).���� = 0
                g����(i).COM�� = 0
                g����(i).����λ = 0
                g����(i).ֹͣλ = 0
                g����(i).���� = 0
                g����(i).У��λ = "N"
                g����(i).�ַ�ģʽ = 0
                g����(i).���� = 0
                g����(i).�������� = ""
                g����(i).�Զ�Ӧ�� = "0"
                g����(i).�ɷ��Ѻ˱걾 = 1
                g����(i).ͨѶĿ¼ = ""
                g����(i).ͨѶ���� = ""
                g����(i).�Զ������ = ""
                g����(i).�Զ������ʿ� = 0
                g����(i).���Ϊͨ���� = 0
            Next
            
            str�����б� = ""
            If gstr�������� <> "" Then
                If Val(gstr��������) <> 0 Then
                    Set rsTmp = GetDevices
                    Do Until rsTmp.EOF
                        str�����б� = str�����б� & "," & rsTmp!ID
                        rsTmp.MoveNext
                    Loop
                End If
            End If
            
            For i = LBound(varSet) To UBound(varSet)
                
                If varSet(i) <> "" Then
                    lngID = Val(Split(varSet(i), ",")(0))
                    If lngID > 0 Then
                        blnAdd = True

                        If gstr�������� <> "" Then
                            If Val(gstr��������) <> 0 Then
                                If InStr("," & str�����б� & ",", "," & lngID & ",") <= 0 Then
                                    blnAdd = False
                                End If
                            Else
                                blnAdd = False
                            End If
                        End If
                        
                        If blnAdd Then
                            strCOM = Split(varSet(i), ",")(1)
                            If Val(strCOM) = 0 Then
                                strTmp = "COM" & Split(varSet(i), ",")(2)
                            Else
                                strTmp = Split(varSet(i), ",")(9) & ":" & Trim(Split(varSet(i), ",")(8))
                            End If
                            
                            Set rsTmp = gobjDatabase.OpenSqlRecord("Select ����,����, ͨѶ������ From �������� where ID=[1]", "ȡ����������", lngID)
                            Do Until rsTmp.EOF
                                strTmp = strTmp & " " & rsTmp!����
                                g����(i).�������� = "(" & rsTmp!���� & ")" & rsTmp!����
                                g����(i).ͨѶ���� = Trim("" & rsTmp!ͨѶ������)
                                rsTmp.MoveNext
                            Loop
                            If g����(i).�������� <> "" Then
                                strDevIDs = strDevIDs & "," & lngID
                                lngSaveAsID = Split(varSet(i), ",")(11)
                                Set rsTmp = gobjDatabase.OpenSqlRecord("Select ���� From �������� where ID=[1]", "ȡ������������", lngSaveAsID)
                                Do Until rsTmp.EOF
                                    strTmp = strTmp & " -> " & rsTmp!����
                                    rsTmp.MoveNext
                                Loop
                                 If strTmp <> "" Then Me.cboDev.AddItem strTmp
                            
                                 With g����(i)
                                     .ID = lngID
                                     .���� = Trim(Split(varSet(i), ",")(1))
                                     .COM�� = Trim(Split(varSet(i), ",")(2))
                                     .������ = Trim(Split(varSet(i), ",")(3))
                                     .����λ = Trim(Split(varSet(i), ",")(4))
                                     .У��λ = Trim(Split(varSet(i), ",")(5))
                                     .ֹͣλ = Trim(Split(varSet(i), ",")(6))
                                     .���� = Trim(Split(varSet(i), ",")(7))
                                     .IP�˿� = Trim(Split(varSet(i), ",")(8))
                                     .IP = Trim(Split(varSet(i), ",")(9))
                                     .�ַ�ģʽ = Trim(Split(varSet(i), ",")(10))
                                     .SaveAsID = lngSaveAsID
                                     .���� = Trim(Split(varSet(i), ",")(12))
                                     .�Զ�Ӧ�� = Trim(Split(varSet(i), ",")(13))
                                     If UBound(Split(varSet(i), ",")) >= 14 Then
                                        .�ɷ��Ѻ˱걾 = Val(Split(varSet(i), ",")(14))
                                     End If
                                     If UBound(Split(varSet(i), ",")) >= 15 Then
                                        .ͨѶĿ¼ = Split(varSet(i), ",")(15)
                                     End If
                                     If UBound(Split(varSet(i), ",")) >= 16 Then
                                        .�Զ������ = Split(varSet(i), ",")(16)
                                     End If
                                     If UBound(Split(varSet(i), ",")) >= 17 Then
                                        .�Զ������ʿ� = Split(varSet(i), ",")(17)
                                     End If
                                     If UBound(Split(varSet(i), ",")) >= 18 Then
                                        .���Ϊͨ���� = Split(varSet(i), ",")(18)
                                     End If
                                 End With
                                
                                 Load Me.DevComm(Me.DevComm.Count)
                                 Me.DevComm(Me.DevComm.Count - 1).InitContrl i, strCmd
    
                            End If
                        End If '���Լ�
                    End If
                End If
            Next
            If Len(strDevIDs) > 0 Then strDevIDs = Mid(strDevIDs, 2)
            If Me.cboDev.ListCount > 0 Then Me.cboDev.ListIndex = 0
        End If
    Else
        gblnClearData = GetSetting("ZLSOFT", "����ģ��\ZlLISSrv", "��ս�����־", 1)
        '��ע��������
        Err = 0: On Error Resume Next
        aPorts = GetAllSettings("ZLSOFT", "����ģ��\ZlLISSrv")
        On Error GoTo errH
        If IsEmpty(aPorts) Then
            ReDim aPorts(8, 0)
            For i = LBound(aPorts) To UBound(aPorts)
                aPorts(i, 0) = "COM" & (i + 1)
            Next
        End If
        
        If IsEmpty(aPorts) Then
            ShowMenu
    '        MsgBox "û�������κ�������ϵͳ�����ܽ��ռ������ݣ������������ý��д���", vbInformation, gstrSysName
        Else
            ReDim g����(UBound(aPorts))
            For i = LBound(g����) To UBound(g����)
                g����(i).ID = 0
                g����(i).IP = ""
                g����(i).IP�˿� = 6666
                g����(i).SaveAsID = 0
                g����(i).������ = 9600
                g����(i).���� = 0
                g����(i).COM�� = 0
                g����(i).����λ = 0
                g����(i).ֹͣλ = 0
                g����(i).���� = 0
                g����(i).У��λ = "N"
                g����(i).�ַ�ģʽ = 0
                g����(i).���� = 0
                g����(i).�������� = ""
                g����(i).�Զ�Ӧ�� = "0"
                g����(i).�ɷ��Ѻ˱걾 = "1"
                g����(i).ͨѶĿ¼ = ""
                g����(i).ͨѶ���� = ""
                g����(i).�Զ������ = ""
                g����(i).�Զ������ʿ� = 0
                g����(i).���Ϊͨ���� = 0
            Next
            
            For i = LBound(aPorts) To UBound(aPorts)
                lngID = Val(GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "Device", 0))
                If lngID > 0 Then
                    
                    strCOM = IIf(aPorts(i, 0) Like "COM*", 0, 1)
                    If strCOM = 0 Then
                    
                        strTmp = aPorts(i, 0)
                    Else
                        strTmp = GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "IP", "127.0.0.1") & ":" & _
                                 GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "Port", "6666")
                        
                    End If
                
                    Set rsTmp = gobjDatabase.OpenSqlRecord("Select ����,����,ͨѶ������ From �������� where ID=[1]", "ȡ����������", lngID)
                    Do Until rsTmp.EOF
                        strTmp = strTmp & " " & rsTmp!����
                        g����(i).�������� = "(" & rsTmp!���� & ")" & rsTmp!����
                        g����(i).ͨѶ���� = Trim("" & rsTmp!ͨѶ������)
                        rsTmp.MoveNext
                    Loop
                    
                    If g����(i).�������� <> "" Then
                        strDevIDs = strDevIDs & "," & lngID
                        lngSaveAsID = Val(GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "SaveAs", "0"))
                        
                        Set rsTmp = gobjDatabase.OpenSqlRecord("Select ���� From �������� where ID=[1]", "ȡ������������", lngSaveAsID)
                        Do Until rsTmp.EOF
                            strTmp = strTmp & " -> " & rsTmp!����
                            rsTmp.MoveNext
                        Loop
                        If strTmp <> "" Then Me.cboDev.AddItem strTmp
                    
                        With g����(i)
                            .ID = lngID
                            .���� = strCOM
                            .COM�� = IIf(strCOM = 0, Replace(aPorts(i, 0), "COM", ""), "0")
                            .������ = Val(GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "Speed", "9600"))
                            .����λ = Val(GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "DataBit", "8"))
                            .У��λ = GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "Parity", "N")
                            .ֹͣλ = Val(GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "StopBit", "1"))
                            .���� = Val(GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "HandShaking", "0"))
                            .IP�˿� = Val(GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "Port", "6666"))
                            .IP = GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "IP", "127.0.0.1")
                            .�ַ�ģʽ = IIf(strCOM = 0, Val(GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "InputMode", "0")), _
                                        Val(GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "InMode", "0")))
                            .SaveAsID = lngSaveAsID
                            .���� = Val(GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "Host", "0"))
                            .�Զ�Ӧ�� = GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "Auto", "0")
                            .�ɷ��Ѻ˱걾 = Val(GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "blnSend", "1"))
                            .ͨѶĿ¼ = GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "ReceiveDir", "")
                            .�Զ������ = GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "AutoCheckMan", "")
                            .�Զ������ʿ� = Val(GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "AutoQCCalc", 0))
                            .���Ϊͨ���� = Val(GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "SaveAsTonDao", 0))
                        End With
                        Load Me.DevComm(Me.DevComm.Count)
                        Me.DevComm(Me.DevComm.Count - 1).InitContrl i, strCmd
                    End If
                End If
            Next
            If Len(strDevIDs) > 0 Then strDevIDs = Mid(strDevIDs, 2)
            If Me.cboDev.ListCount > 0 Then Me.cboDev.ListIndex = 0
        End If
        
    End If ''�Ƿ�����ݿ��ȡ����
    ReadPara = True
    Exit Function
errH:
    MsgBox Err.Description
    
End Function

Private Sub ShowMenu()
    Dim blnEnabled As Boolean
    If Me.cboDev.ListIndex = -1 Then
        Me.mnuFileOpen.Enabled = False
        Me.mnuFileClose.Enabled = False
        Me.tbrMain.Buttons("����").Enabled = False
        Me.tbrMain.Buttons("�Ͽ�").Enabled = False
    Else
        blnEnabled = Me.DevComm(cboDev.ListIndex + 1).PortOpened
        
        Me.mnuFileOpen.Enabled = Not blnEnabled
        Me.mnuFileClose.Enabled = blnEnabled
        Me.tbrMain.Buttons("����").Enabled = Not blnEnabled
        Me.tbrMain.Buttons("�Ͽ�").Enabled = blnEnabled
    End If
    Me.mnuDevExp.Enabled = mblnOwner
End Sub

Public Function SendSample(ByVal lngDeviceID As Long, ByVal strSampleDate As String, ByVal lngSampleNO As Long, Optional strAdviceIDs As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As Boolean
'���ͱ걾��¼������
    Dim i As Integer
    SendSample = True
    
    For i = 1 To DevComm.UBound
        If DevComm(i).DeviceID = lngDeviceID Then
            SendSample = DevComm(i).SendSample(lngDeviceID, strSampleDate, lngSampleNO, strAdviceIDs, blnUndo, iType)
            Exit For
        End If
    Next
End Function

Private Sub timConn_Timer()
    Dim dateNow As Date, i As Integer
    Dim strSQL As String, rsTmp As ADODB.Recordset
    On Error GoTo errH
    strSQL = "Select 1 From dual"
    
    Set rsTmp = gcnOracle.Execute(strSQL)
    
    Exit Sub
errH:
    WriteLog "���ݿ����ӶϿ�", LOG_������־, Err.Number, Err.Description
    i = 0
    Do While i <= 30
        If Err.Number <> 0 Then
            Err.Clear
            If gcnOracle.State = 1 Then gcnOracle.Close
            gcnOracle.Open mstrConn
        Else
            WriteLog "���ݿ������ѻָ�", LOG_������־, 0, "�������Դ���=" & i
            Exit Do
        End If
        i = i + 1
    Loop
End Sub

Private Sub timLsn_Timer()
    Call mnuReLoad_Click
End Sub

Private Sub WinsockS_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    Dim aItem() As String
    On Error Resume Next
    With Me.WinsockS
        .GetData strData
    End With
    If Len(Trim(strData)) = 0 Then Exit Sub
    aItem = Split(strData, ",")
    If UBound(aItem) <= 0 Then Exit Sub
    If aItem(1) <> Me.WinsockS.LocalIP Then Exit Sub            '����ͬһIPʱ�˳�
    Select Case aItem(0)
        Case "SendSample"
            SendSample aItem(2), aItem(3), aItem(4), IIf(aItem(5) = "", "", _
            Replace(aItem(5), ";", ",")), IIf(aItem(6) = "", "", aItem(6)), IIf(aItem(7) = "", "", aItem(7))
        Case "ResultFromFile"
            ResultFromFile aItem(2), aItem(3), aItem(4), aItem(5), aItem(6)
    End Select
    '���ز������ʱ
    With Me.WinsockS
        .SendData .LocalIP & ";" & strSend_True
    End With
End Sub

Private Function ExpLisDevData() As String
    '������������������
    
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim objStream As TextStream
    Dim objFileSystem As New FileSystemObject
    
    Dim strPath As String, strFileName As String
    Dim strLog As String '��¼��������
    
    Dim strTable As String, strFiled As String
    
    On Error GoTo errH
    
    
    
    strTable = "������Ŀ"
    strFiled = "�ٴ�����,VARCHAR2(4000)|��˽��Ŀ,NUMBER(1)"
    If Not CheckFiled(strTable, strFiled) Then Exit Function
    
    strTable = "����������Ŀ"
    strFiled = "��������Ŀ,NUMBER(1)"
    If Not CheckFiled(strTable, strFiled) Then Exit Function
    
    
    '------------------------------------------------------------------------------
    strFileName = App.Path & "\zlLis��������_" & Format(date, "yyyyMMdd") & ".txt"
    
    If objFileSystem.FileExists(strFileName) Then Kill strFileName
    Call objFileSystem.CreateTextFile(strFileName)
    Set objStream = objFileSystem.OpenTextFile(strFileName, ForAppending)
    
    objStream.WriteLine "[�û�����]"
    objStream.WriteLine gstr��λ����
    
    strSQL = "Select ID||Chr(9)||����||Chr(9)||����||Chr(9)||ͨѶ������||Chr(9)||�������� As Line From �������� Where ΢���� <> 1"
    Call WritData("[����]", strSQL, objStream)
    
    strSQL = "Select Distinct B.��Ŀid||Chr(9)||D.����||Chr(9)||A.��д||Chr(9)||A.��Ŀ���||Chr(9)||A.�������||Chr(9)||D.�����Ա�||Chr(9)||D.��������||Chr(9)||D.�걾��λ||Chr(9)||A.��λ||Chr(9)||A.Ĭ��ֵ||Chr(9)||" & vbNewLine & _
            "                A.ȡֵ����||Chr(9)||A.��˽��Ŀ||Chr(9)||A.���Թ�ʽ||Chr(9)||A.�����Թ�ʽ||Chr(9)||A.Cutoff��ʽ||Chr(9)||A.���㹫ʽ||Chr(9)||A.���鷽��||Chr(9)||A.�ٴ����� As Line" & vbNewLine & _
            "From ������ĿĿ¼ D, ���鱨����Ŀ C, ������Ŀ A, ����������Ŀ B, �������� E" & vbNewLine & _
            "Where D.�����Ŀ <> 1 And C.������Ŀid = D.ID And C.������Ŀid = B.��Ŀid And A.������Ŀid = B.��Ŀid And E.ID = B.����id And E.΢���� <> 1"
    Call WritData("[��Ŀ]", strSQL, objStream)

    strSQL = "Select Distinct A.��Ŀid||Chr(9)||A.�걾����||Chr(9)||A.�Ա���||Chr(9)||A.��������||Chr(9)||A.��������||Chr(9)||A.���䵥λ||Chr(9)||A.�ο���ֵ||Chr(9)||A.�ο���ֵ||Chr(9)||A.�ٴ����� As Line" & vbNewLine & _
            "From ������Ŀ�ο� A, ����������Ŀ B, �������� E" & vbNewLine & _
            "Where A.��Ŀid = B.��Ŀid And (Nvl(A.��������, 0) <> 0 Or Nvl(A.��������, 0) <> 0 Or Nvl(A.�ο���ֵ, 0) <> 0 Or Nvl(A.�ο���ֵ, 0) <> 0) And" & vbNewLine & _
            "      E.ID = B.����id And E.΢���� <> 1"
    Call WritData("[��Ŀ�ο�]", strSQL, objStream)
    
    strSQL = "Select ����id||Chr(9)||��Ŀid||Chr(9)||ͨ������||Chr(9)||С��λ��||Chr(9)||��������Ŀ as Line" & vbNewLine & _
            "From ����������Ŀ A, �������� E" & vbNewLine & _
            "Where E.ID = A.����id And E.΢���� <> 1"
    Call WritData("[������Ŀ]", strSQL, objStream)
    
    '΢����--�ֵ��������
    
    strSQL = "select ����||chr(9)||����||chr(9)||���� as Line from ����ϸ������"
    Call WritData("[����ϸ������]", strSQL, objStream)
    
    strSQL = "select ����||chr(9)||����||chr(9)||����||chr(9)||ȱʡ��־ as Line from ����ϸ�����"
    Call WritData("[����ϸ�����]", strSQL, objStream)
    
    strSQL = "select ����||chr(9)||����||chr(9)||����||chr(9)||ȱʡ��־ as Line from ����Ⱦɫ����"
    Call WritData("[����Ⱦɫ����]", strSQL, objStream)
    
    strSQL = "select ����||chr(9)||����||chr(9)||���� as Line from ϸ����ⷽ��"
    Call WritData("[ϸ����ⷽ��]", strSQL, objStream)
    
    '΢����--����
    strSQL = "select id||chr(9)||����||chr(9)||����||chr(9)||Ӣ��||chr(9)||���� as Line from ���鿹������"
    Call WritData("[���鿹������]", strSQL, objStream)
    
    strSQL = "select id||chr(9)||����||chr(9)||������||chr(9)||Ӣ����||chr(9)||����||chr(9)||˵��||chr(9)||ҩ������||chr(9)||whonet��||chr(9)||�÷�����1||chr(9)||ѪҩŨ��1||chr(9)||��ҩŨ��1||chr(9)||�÷�����2||chr(9)||ѪҩŨ��2||chr(9)||��ҩŨ��2 as Line from �����ÿ�����"
    Call WritData("[�����ÿ�����]", strSQL, objStream)
    
    strSQL = "select ������id||chr(9)||�����ط���id as Line From ���鿹������ҩ"
    Call WritData("[���鿹������ҩ]", strSQL, objStream)
    
    strSQL = "select id||chr(9)||����||chr(9)||��������||chr(9)||Ӣ������||chr(9)||���� as Line from ����ϸ������"
    Call WritData("[����ϸ������]", strSQL, objStream)
    
    strSQL = "select id||chr(9)||����||chr(9)||������||chr(9)||Ӣ����||chr(9)||����id||chr(9)||����||chr(9)||Ĭ��ҩ��||chr(9)||Ĭ�Ϸ���||chr(9)||whonet��||chr(9)||Ĭ�Ͻ��||chr(9)||ϸ�����||chr(9)||ϸ������||chr(9)||�����Ϸ��� as Line from ����ϸ��"
    Call WritData("[����ϸ��]", strSQL, objStream)
    
    strSQL = "select ϸ��id||chr(9)||�����ط���id||chr(9)||ȱʡ��־ as Line From ����ϸ��������"
    Call WritData("[����ϸ��������]", strSQL, objStream)
    
    strSQL = "select ϸ��id||chr(9)||�����ط���id||chr(9)||������id||chr(9)||ҩ������||chr(9)||�ο���ֵ||chr(9)||�ο���ֵ||chr(9)||�жϷ�ʽ||chr(9)||��ע as Line From ����ϸ�������زο�"
    Call WritData("[����ϸ�������زο�]", strSQL, objStream)
    
    '΢����--������������Ŀ����
    strSQL = "Select ID||Chr(9)||����||Chr(9)||����||Chr(9)||ͨѶ������||Chr(9)||�������� As Line  From �������� Where ΢����=1"
    Call WritData("[΢��������]", strSQL, objStream)
    
    strSQL = "Select ����id||Chr(9)||ͨ������||Chr(9)||ϸ��id||Chr(9)||������id as Line From ����ϸ������"
    Call WritData("[����ϸ������]", strSQL, objStream)
    
    objStream.Close
    Set objStream = Nothing
    ExpLisDevData = strFileName
    Exit Function
errH:
    MsgBox "��������ʱ���ִ���" & Err.Description & vbCrLf & strLog & vbCrLf & strSQL

End Function

Private Sub WritData(ByVal strHead As String, ByVal str_Sql As String, objStream As TextStream)
    Dim rsTmp As ADODB.Recordset
    
    If str_Sql = "" Then Exit Sub
    If objStream Is Nothing Then Exit Sub
    If InStr(UCase(str_Sql), UCase(" as Line")) <= 0 Then Exit Sub
    
    objStream.WriteLine strHead
    Set rsTmp = gobjDatabase.OpenSqlRecord(str_Sql, Me.Caption)
    With rsTmp
        Do Until .EOF
            objStream.WriteLine "" & !Line
            .MoveNext
        Loop
    End With
    objStream.WriteLine ""
    
End Sub

Private Function CheckFiled(ByVal strTable As String, ByVal strFileds As String) As Boolean
    '������ݽṹ����Ļ��ͼӡ�
    
    Dim rsTmp As ADODB.Recordset, strSQL As String, i As Integer
    Dim strName As String, strTypeLen As String
    Dim varFiled As Variant
    strSQL = "Select Data_Type As ����, Data_Precision As ����, Data_Scale As С��, Data_Length As ����" & vbNewLine & _
            "From User_Tab_Columns" & vbNewLine & _
            "Where Table_Name = [1] And Column_Name = [2]"
    
    varFiled = Split(strFileds, "|")
    
    strSQL = "Select upper(Column_Name) as �ֶ��� " & vbNewLine & _
    "From User_Tab_Columns" & vbNewLine & _
    "Where Table_Name = [1]"
    Set rsTmp = gobjDatabase.OpenSqlRecord(strSQL, Me.Caption, strTable)
    If rsTmp.RecordCount <= 0 Then
        MsgBox "ȱ�ٱ�" & strTable & "��,���ܵ�������!"
        Exit Function
    End If
    
    For i = LBound(varFiled) To UBound(varFiled)
        strName = UCase(Split(varFiled(i), ",")(0))
        strTypeLen = UCase(Split(varFiled(i), ",")(1))
        rsTmp.Filter = "�ֶ���= '" & strName & "'"
        If rsTmp.EOF Then
            strSQL = "Alter Table " & strTable & " Add " & strName & " " & strTypeLen
            gcnOracle.Execute strSQL
        End If
    Next

    CheckFiled = True
    
End Function
