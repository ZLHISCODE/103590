VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frm��ҩ���� 
   Caption         =   "��ҩ���ڹ���"
   ClientHeight    =   4980
   ClientLeft      =   168
   ClientTop       =   456
   ClientWidth     =   6648
   Icon            =   "frm��ҩ����.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4980
   ScaleWidth      =   6648
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Tag             =   "93"
   Begin MSComctlLib.ImageList LvwBlack 
      Left            =   2730
      Top             =   2280
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��ҩ����.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��ҩ����.frx":075C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList TvwImg 
      Left            =   2580
      Top             =   1020
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��ҩ����.frx":0A76
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��ҩ����.frx":0D90
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList LvwColor 
      Left            =   2700
      Top             =   1680
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��ҩ����.frx":10AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��ҩ����.frx":13C4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView Lvw 
      Height          =   2235
      Left            =   3120
      TabIndex        =   1
      Top             =   1380
      Width           =   2595
      _ExtentX        =   4572
      _ExtentY        =   3937
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "LvwBlack"
      SmallIcons      =   "LvwColor"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "����"
         Text            =   "����"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "����"
         Text            =   "����"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "״̬"
         Text            =   "״̬"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "ҩ��"
         Text            =   "ҩ��"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "ר��"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "�кŴ���"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.TreeView Tree 
      Height          =   3345
      Left            =   240
      TabIndex        =   0
      Top             =   990
      Width           =   2445
      _ExtentX        =   4318
      _ExtentY        =   5906
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "TvwImg"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList Ilscolor 
      Left            =   4290
      Top             =   660
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��ҩ����.frx":16DE
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��ҩ����.frx":18FE
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��ҩ����.frx":1B1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��ҩ����.frx":1D38
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��ҩ����.frx":1F58
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��ҩ����.frx":2178
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��ҩ����.frx":2398
            Key             =   "Start"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��ҩ����.frx":25B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��ҩ����.frx":2ACA
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��ҩ����.frx":2CEA
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList Ilsrw 
      Left            =   5040
      Top             =   630
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��ҩ����.frx":2F0A
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��ҩ����.frx":312A
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��ҩ����.frx":334A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��ҩ����.frx":3564
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��ҩ����.frx":3784
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��ҩ����.frx":39A4
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��ҩ����.frx":3BC4
            Key             =   "Start"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��ҩ����.frx":3DE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��ҩ����.frx":42F6
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��ҩ����.frx":4516
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar Cbar 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6648
      _ExtentX        =   11726
      _ExtentY        =   1164
      BandCount       =   1
      _CBWidth        =   6648
      _CBHeight       =   660
      _Version        =   "6.7.9782"
      Child1          =   "Tbar"
      MinHeight1      =   612
      Width1          =   8376
      Key1            =   "only"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar Tbar 
         Height          =   612
         Left            =   24
         TabIndex        =   4
         Top             =   24
         Width           =   6552
         _ExtentX        =   11557
         _ExtentY        =   1080
         ButtonWidth     =   783
         ButtonHeight    =   1080
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "Ilsrw"
         HotImageList    =   "Ilscolor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   14
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ԥ��"
               Key             =   "Preview"
               Object.ToolTipText     =   "Ԥ��"
               Object.Tag             =   "Ԥ��"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "Print"
               Object.ToolTipText     =   "��ӡ"
               Object.Tag             =   "��ӡ"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Add"
               Object.ToolTipText     =   "���ӷ�ҩ����"
               Object.Tag             =   "����"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�޸�"
               Key             =   "Modify"
               Object.ToolTipText     =   "�޸�"
               Object.Tag             =   "�޸�"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ɾ��"
               Key             =   "Delete"
               Object.ToolTipText     =   "ɾ��"
               Object.Tag             =   "ɾ��"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "split1"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�ϰ�"
               Key             =   "Start"
               Object.ToolTipText     =   "�ϰ�"
               Object.Tag             =   "�ϰ�"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�°�"
               Key             =   "Stop"
               Object.ToolTipText     =   "�°�"
               Object.Tag             =   "�°�"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "split2"
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�鿴"
               Key             =   "View"
               Object.ToolTipText     =   "�鿴"
               Object.Tag             =   "�鿴"
               ImageIndex      =   8
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
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
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Quit"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageIndex      =   10
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox picSplit 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3225
      Left            =   3120
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3228
      ScaleWidth      =   48
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1590
      Width           =   45
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   4620
      Width           =   6645
      _ExtentX        =   11726
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2350
            MinWidth        =   882
            Picture         =   "frm��ҩ����.frx":4736
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6689
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
      Begin VB.Menu mnuFileset 
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
      Begin VB.Menu mnusplit2 
         Caption         =   "-"
      End
      Begin VB.Menu mnufileexit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuEditAdd 
         Caption         =   "����(&A)"
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "�޸�(&M)"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "ɾ��(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditSplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditStart 
         Caption         =   "�ϰ�(&S)"
      End
      Begin VB.Menu mnuEditStop 
         Caption         =   "�°�(&T)"
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
   Begin VB.Menu mnuview 
      Caption         =   "�鿴(&V)"
      Begin VB.Menu mnuViewTool 
         Caption         =   "������(&T)"
         Begin VB.Menu mnuViewButton 
            Caption         =   "��׼��ť(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuviewspilt1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewText 
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
      Begin VB.Menu mnuViewSplit3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewShow 
         Caption         =   "�����°ര��(&H)"
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuviewr 
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
         Caption         =   "WEB�ϵ�����(&W)"
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
End
Attribute VB_Name = "frm��ҩ����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RecData As New ADODB.Recordset  '��ҩ���ڼ�¼��
Private BlnStartUp As Boolean
Private blnFirst As Boolean
Private LngLastRoot As Long
Private mlngMode As Long
Private mstrPrivs As String             '��ǰ�û����еĵ�ǰģ��Ĺ���
Private mstrҩ�� As String
Dim mbln���в��� As Boolean             '�Ƿ�������в���Ȩ��
Private Sub Form_Activate()
    If BlnStartUp = False Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    BlnStartUp = False
    blnFirst = True
    LngLastRoot = 0
    mlngMode = glngModul
    mstrPrivs = gstrprivs
    
    RestoreWinState Me, App.ProductName
    Call zldatabase.ShowReportMenu(Me, glngSys, mlngMode, mstrPrivs)
    
    mbln���в��� = IsHavePrivs(mstrPrivs, "���в���")
   
    mnuViewIcon_Click Me.Lvw.View
    If LoadInTree = False Then Exit Sub
    Ȩ�޿���
    
    BlnStartUp = True
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
    If Me.Width < 8000 Then
        Me.Width = 8000
        Exit Sub
    End If
    
    If blnFirst Then picSplit.Left = 3100
    With picSplit
        .Top = IIf(Cbar.Visible = False, 0, Cbar.Height)
        .Height = Me.ScaleHeight - IIf(stbThis.Visible = False, 0, stbThis.Height) - .Top
    End With
    
    With Tree
        .Top = picSplit.Top
        .Width = picSplit.Left
        .Height = picSplit.Height
        .Left = 0
    End With
    
    With Lvw
        .Top = picSplit.Top
        .Left = picSplit.Left + picSplit.Width
        .Width = Me.ScaleWidth - .Left
        .Height = picSplit.Height
    End With
    blnFirst = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Sub Lvw_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With Lvw
        .Sorted = False
        .SortKey = ColumnHeader.index - 1
        .SortOrder = IIf(.SortOrder = lvwAscending, lvwDescending, lvwAscending)
        .Sorted = True
    End With
End Sub

Private Sub Lvw_DblClick()
    If Lvw.ListItems.count = 0 Then Exit Sub
    If Lvw.SelectedItem Is Nothing Then Exit Sub
    If Tbar.Buttons("Add").Visible = False Then Exit Sub
    
    mnuEditModify_Click
End Sub

Private Sub Lvw_ItemClick(ByVal Item As MSComctlLib.listItem)
     mnuEditStop.Enabled = IIf(Item.SubItems(2) = "�ϰ�", True, False)
     mnuEditStart.Enabled = mnuEditStop.Enabled Xor True
     Tbar.Buttons("Start").Enabled = mnuEditStart.Enabled
     Tbar.Buttons("Stop").Enabled = mnuEditStop.Enabled
End Sub

Private Sub Lvw_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Lvw.ListItems.count = 0 Then Exit Sub
    If Lvw.SelectedItem Is Nothing Then Exit Sub
    If Tbar.Buttons("Add").Visible = False Then Exit Sub
    
    mnuEditModify_Click
End Sub

Private Sub Lvw_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    If Tbar.Buttons("Add").Visible Or Tbar.Buttons("Start").Visible Then
        PopupMenu mnuEdit, 2
    End If
End Sub

Private Sub mnuEditAdd_Click()
    With Frm��ҩ���ڱ༭
        .EditState = 1
        .Show 1, Me
    End With
    mnuviewr_Click
End Sub

Private Sub mnuEditDelete_Click()
    Dim dteBegin As Date, dteEnd As Date
    Dim lngҩ��ID As Long
    Dim strWinName As String
    Dim rsData As ADODB.Recordset
    Dim strDeptNode As String
    Dim blnAllStop As Boolean
    
    On Error GoTo ErrHand
    
    dteEnd = zldatabase.Currentdate
    dteBegin = DateAdd("D", -3, dteEnd)
    lngҩ��ID = Val(Mid(frm��ҩ����.Lvw.SelectedItem.Key, 3, InStr(1, frm��ҩ����.Lvw.SelectedItem.Key, ",") - 3))
    strWinName = Lvw.SelectedItem.Text
    strDeptNode = GetDeptStationNode(lngҩ��ID)
    
    gstrSQL = "Select 1 From δ��ҩƷ��¼ Where �ⷿid = [1] And ��ҩ���� = [2] And �������� Between [3] And [4] And Rownum < 2 "
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "mnuEditStop_Click", lngҩ��ID, strWinName, dteBegin, dteEnd)
    
    '�������δ��ҩ�ľʹ򿪵�����ҩ���ڴ���
    If rsData.EOF Then
        If MsgBox("��ȷ��Ҫɾ���÷�ҩ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    Else
        MsgBox "�ô��ڻ���δ��ҩ�������뽫��������������ڡ�", vbInformation, gstrSysName
        If frm������ҩ����.ShowMe(lngҩ��ID, Me, dteBegin, dteEnd, strDeptNode, strWinName) = False Then
            If MsgBox("δ��δ������������������ҩ���ڣ��Ƿ���ɾ���÷�ҩ���ڣ�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
    End If
    
    gstrSQL = "zl_��ҩ����_delete("
    '����
    gstrSQL = gstrSQL & "'" & Lvw.SelectedItem.SubItems(1) & "'"
    '�ⷿID
    gstrSQL = gstrSQL & "," & Mid(frm��ҩ����.Lvw.SelectedItem.Key, 3, InStr(1, frm��ҩ����.Lvw.SelectedItem.Key, ",") - 3)
    gstrSQL = gstrSQL & ")"
    
    On Error GoTo ErrHand
    Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    mnuviewr_Click
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub mnuEditModify_Click()
    With Frm��ҩ���ڱ༭
        .EditState = 2
        .Show 1, Me
    End With
    mnuviewr_Click
End Sub

Private Sub mnuEditStart_Click()
    gstrSQL = "zl_��ҩ����_setwork("
    '����
    gstrSQL = gstrSQL & "'" & Lvw.SelectedItem.SubItems(1) & "'"
    '�ⷿID
    gstrSQL = gstrSQL & "," & Mid(frm��ҩ����.Lvw.SelectedItem.Key, 3, InStr(1, frm��ҩ����.Lvw.SelectedItem.Key, ",") - 3)
    '�Ƿ��ϰ�
    gstrSQL = gstrSQL & ",1"
    gstrSQL = gstrSQL & ")"

    On Error GoTo ErrHand
    Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    mnuviewr_Click
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditStop_Click()
    Dim dteBegin As Date, dteEnd As Date
    Dim lngҩ��ID As Long
    Dim strWinName As String
    Dim rsData As ADODB.Recordset
    Dim strDeptNode As String
    Dim blnAllStop As Boolean
    
    On Error GoTo ErrHand
    
    dteEnd = zldatabase.Currentdate
    dteBegin = DateAdd("D", -3, dteEnd)
    lngҩ��ID = Val(Mid(frm��ҩ����.Lvw.SelectedItem.Key, 3, InStr(1, frm��ҩ����.Lvw.SelectedItem.Key, ",") - 3))
    strWinName = Lvw.SelectedItem.Text
    strDeptNode = GetDeptStationNode(lngҩ��ID)
    
    gstrSQL = "select 1 from ��ҩ���� where ҩ��id=[1] and �ϰ��=1 And ����<>[2] "
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "mnuEditStop_Click", lngҩ��ID, strWinName)
    If rsData.RecordCount = 0 Then blnAllStop = True
    
    gstrSQL = "Select 1 From δ��ҩƷ��¼ Where �ⷿid = [1] And ��ҩ���� = [2] And �������� Between [3] And [4] And Rownum < 2 "
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "mnuEditStop_Click", lngҩ��ID, strWinName, dteBegin, dteEnd)
    
    '�������δ��ҩ�ľʹ򿪵�����ҩ���ڴ���
    If rsData.RecordCount > 0 Then
        If blnAllStop Then
            If MsgBox("�ô��ڻ���δ��ҩ�������������ڶ����°࣬�Ƿ��ֽ���ǰ��������Ϊ�°ࣿ", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        Else
            If MsgBox("�ô��ڻ���δ��ҩ�������Ƿ��ֽ���ǰ��������Ϊ�°ࣿ", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            gstrSQL = "zl_��ҩ����_setwork("
            '����
            gstrSQL = gstrSQL & "'" & Lvw.SelectedItem.SubItems(1) & "'"
            '�ⷿID
            gstrSQL = gstrSQL & "," & Mid(frm��ҩ����.Lvw.SelectedItem.Key, 3, InStr(1, frm��ҩ����.Lvw.SelectedItem.Key, ",") - 3)
            '�Ƿ��ϰ�
            gstrSQL = gstrSQL & ",0"
            gstrSQL = gstrSQL & ")"
        
            Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            
            MsgBox "�뽫δ��ҩ�����������������ڡ�", vbInformation, gstrSysName
            Call frm������ҩ����.ShowMe(lngҩ��ID, Me, dteBegin, dteEnd, strDeptNode, strWinName)
            mnuviewr_Click
            Exit Sub
        End If
    End If
    
    gstrSQL = "zl_��ҩ����_setwork("
    '����
    gstrSQL = gstrSQL & "'" & Lvw.SelectedItem.SubItems(1) & "'"
    '�ⷿID
    gstrSQL = gstrSQL & "," & Mid(frm��ҩ����.Lvw.SelectedItem.Key, 3, InStr(1, frm��ҩ����.Lvw.SelectedItem.Key, ",") - 3)
    '�Ƿ��ϰ�
    gstrSQL = gstrSQL & ",0"
    gstrSQL = gstrSQL & ")"

    Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    mnuviewr_Click
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuFileExcel_Click()
    subPrint 3
End Sub

Private Sub mnufileexit_Click()
    Unload Me
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

Private Sub mnuHelpTitle_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub mnuHelpWebHome_Click()
    zlHomePage hWnd
End Sub

Private Sub mnuHelpWebMail_Click()
    zlMailTo hWnd
End Sub

Private Sub mnuReportItem_Click(index As Integer)
    'Ĭ�ϲ�����ҩ��=ҩ��id����ҩ����=��ҩ��������
    Dim Str���� As String
    If Not Me.Lvw.SelectedItem Is Nothing Then
        Str���� = Me.Lvw.SelectedItem.Text
    End If
    
    Call ReportOpen(gcnOracle, Split(mnuReportItem(index).Tag, ",")(0), Split(mnuReportItem(index).Tag, ",")(1), Me, _
        "ҩ��=" & IIf(LngLastRoot = 0, "", LngLastRoot), _
        "��ҩ����=" & Str����)
End Sub

Private Sub mnuViewButton_Click()
    mnuViewButton.Checked = Not mnuViewButton.Checked
    Cbar.Visible = mnuViewButton.Checked
    mnuViewText.Enabled = mnuViewButton.Checked
    Cbar.Bands("only").MinHeight = Tbar.Height
    Form_Resize
End Sub

Private Sub mnuViewIcon_Click(index As Integer)
    mnuViewIcon(0).Checked = False
    mnuViewIcon(1).Checked = False
    mnuViewIcon(2).Checked = False
    mnuViewIcon(3).Checked = False
    
    mnuViewIcon(index).Checked = True
    
    Select Case index
        Case 0
            Lvw.View = lvwIcon
        Case 1
            Lvw.View = lvwSmallIcon
        Case 2
            Lvw.View = lvwList
        Case 3
            Lvw.View = lvwReport
    End Select
End Sub

Private Sub mnuviewr_Click()
    If LoadInTree = False Then
        BlnStartUp = False
        Form_Activate
    End If
End Sub

Private Sub mnuViewShow_Click()
    mnuViewShow.Checked = mnuViewShow.Checked Xor True
    mnuviewr_Click
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = mnuViewStatus.Checked
    Form_Resize
End Sub

Private Sub mnuViewText_Click()
    Dim buttTemp As Button
    
    mnuViewText.Checked = Not mnuViewText.Checked
    For Each buttTemp In Tbar.Buttons
        If mnuViewText.Checked Then
            buttTemp.Caption = buttTemp.Tag
        Else
            buttTemp.Caption = ""
        End If
    Next
    Cbar.Bands("only").MinHeight = Tbar.Height
    Form_Resize
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub subPrint(ByVal bytMode As Byte)
    Dim objPrint As New zlPrintLvw
    objPrint.Title.Text = "��ҩ����"
    Set objPrint.Body.objData = Lvw
    objPrint.BelowAppItems.Add "��ӡ�ˣ�" & gstrUserName
    objPrint.BelowAppItems.Add "��ӡʱ�䣺" & Format(zldatabase.Currentdate, "yyyy��MM��dd��")

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

Private Function LoadInLvw(Optional ByVal Blnҩ��ID As Long = 0)
    Dim strCon As String
    
    If Not mbln���в��� Then
        strCon = " And Id In (Select ����id From ������Ա Where ��Աid = [1]) "
    End If
    
    On Error GoTo errHandle
    gstrSQL = "Select A.����,A.����,A.�ϰ��,A.ҩ��ID,A.ר��,B.���� as ҩ��,A.�кŴ��� From ��ҩ���� A,���ű� B" & _
    " Where A.ҩ��ID=B.ID " & strCon
    If Blnҩ��ID <> 0 Then gstrSQL = gstrSQL & " And ҩ��ID=[2]"
    If mnuViewShow.Checked = False Then gstrSQL = gstrSQL & " And �ϰ��=1"
    gstrSQL = gstrSQL & " Order by A.����"
   Set RecData = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, glngUserId, Blnҩ��ID)
   
   With RecData
        Lvw.ListItems.Clear
        Do While Not .EOF
            If !�ϰ�� = 1 Then
                Lvw.ListItems.Add , "K_" & !ҩ��ID & "," & !����, !����, 1, 1
            Else
                Lvw.ListItems.Add , "K_" & !ҩ��ID & "," & !����, !����, 2, 2
            End If
            Lvw.ListItems("K_" & !ҩ��ID & "," & !����).SubItems(1) = !����
            Lvw.ListItems("K_" & !ҩ��ID & "," & !����).SubItems(2) = IIf(!�ϰ�� = 1, "�ϰ�", "�°�")
            Lvw.ListItems("K_" & !ҩ��ID & "," & !����).SubItems(3) = !ҩ��
            Lvw.ListItems("K_" & !ҩ��ID & "," & !����).SubItems(4) = IIf(IsNull(!ר��), "", IIf(!ר�� = 1, "��", ""))
            Lvw.ListItems("K_" & !ҩ��ID & "," & !����).SubItems(5) = zlStr.Nvl(!�кŴ���)
            .MoveNext
        Loop
        If Blnҩ��ID <> 0 Then
            Lvw.ColumnHeaders(4).Width = 0
        Else
            Lvw.ColumnHeaders(4).Width = 1500
        End If
        
        If .RecordCount = 0 Then
            mnuEditModify.Enabled = False
            mnuEditDelete.Enabled = False
            mnuEditStart.Enabled = False
            mnuEditStop.Enabled = False
            Tbar.Buttons("Modify").Enabled = False
            Tbar.Buttons("Delete").Enabled = False
            Tbar.Buttons("Start").Enabled = False
            Tbar.Buttons("Stop").Enabled = False
        Else
            mnuEditModify.Enabled = True
            mnuEditDelete.Enabled = True
            Tbar.Buttons("Modify").Enabled = True
            Tbar.Buttons("Delete").Enabled = True
        End If
    End With
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function LoadInTree() As Boolean
    Dim strCon As String
    
    LoadInTree = False
    On Error GoTo errHandle
    If Not mbln���в��� Then
        strCon = " And Id In (Select ����id From ������Ա Where ��Աid = [1]) "
    End If
    
    gstrSQL = " Select ID,����,���� From ���ű� Where (վ�� = '" & gstrNodeNo & "' Or վ�� is Null) And ID in (" & _
          " Select distinct ����ID From ��������˵��" & _
          " Where �������� Like '%ҩ��')" & strCon & _
          " And To_Char(����ʱ��,'yyyy-MM-dd')='3000-01-01' Order by ����"
       
    Set RecData = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, glngUserId)
        
    With RecData
        If .EOF Then
            MsgBox "ҩ��ҩ����Ϣ��ȫ�����Ź��������㲻��ҩ����Ա��", vbInformation, gstrSysName
            Exit Function
        End If
        
        Tree.Nodes.Clear
        Tree.Nodes.Add , , "R", "����ҩ��", 1, 1
        
        Do While Not .EOF
            Tree.Nodes.Add "R", 4, "K_" & !Id, "��" & !���� & "��" & !����, 2, 2
            .MoveNext
        Loop
        If LngLastRoot <> 0 Then
            Tree.Nodes("K_" & LngLastRoot).Selected = True
        Else
            Tree.Nodes("R").Selected = True
        End If
        Tree.SelectedItem.Selected = True
        Tree.SelectedItem.Expanded = True
        Tree_NodeClick Tree.SelectedItem
    End With
    
    LoadInTree = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Ȩ�޿���()
    If Not IsHavePrivs(mstrPrivs, "��ɾ��") Then
        mnuEditAdd.Visible = False
        mnuEditModify.Visible = False
        mnuEditDelete.Visible = False
        mnuEditSplit1.Visible = False

        Tbar.Buttons("Add").Visible = False
        Tbar.Buttons("Modify").Visible = False
        Tbar.Buttons("Delete").Visible = False
        Tbar.Buttons("split1").Visible = False
    End If
    If Not IsHavePrivs(mstrPrivs, "���°�") Then
        If Not IsHavePrivs(mstrPrivs, "��ɾ��") Then
            mnuEdit.Visible = False
        Else
            mnuEditStart.Visible = False
            mnuEditStop.Visible = False
            mnuEditSplit1.Visible = False
        End If
        Tbar.Buttons("Start").Visible = False
        Tbar.Buttons("Stop").Visible = False
        Tbar.Buttons("split2").Visible = False
    End If
End Function

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    
    If picSplit.Left + x < 3000 Then Exit Sub
    If picSplit.Left + x > Me.ScaleWidth - 3000 Then Exit Sub
    
    picSplit.Left = picSplit.Left + x
    Form_Resize
End Sub

Private Sub Tbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Print"
            mnuFilePrint_Click
        Case "Preview"
            mnuFilePreView_Click
        Case "Add"
            mnuEditAdd_Click
        Case "Modify"
            mnuEditModify_Click
        Case "Delete"
            mnuEditDelete_Click
        Case "Start"
            mnuEditStart_Click
        Case "Stop"
            mnuEditStop_Click
        Case "View"
            If Lvw.View < lvwReport Then
                mnuViewIcon_Click Lvw.View + 1
            Else
                mnuViewIcon_Click 0
            End If
        Case "Help"
            mnuHelpTitle_Click
        Case "Quit"
            Unload Me
            Exit Sub
    End Select
End Sub

Private Sub Tbar_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    mnuViewIcon_Click ButtonMenu.index - 1
End Sub

Private Sub Tbar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuview, 2
End Sub

Private Sub Tree_NodeClick(ByVal Node As MSComctlLib.Node)
    If Node.Key = "R" Then
        mnuEditAdd.Enabled = False
        Tbar.Buttons("Add").Enabled = False
        LngLastRoot = 0
        mstrҩ�� = ""
    Else
        mnuEditAdd.Enabled = True
        Tbar.Buttons("Add").Enabled = True
        LngLastRoot = Mid(Node.Key, 3)
        mstrҩ�� = Mid(Node.Text, InStr(1, Node.Text, "��") + 1)
    End If
    
    LoadInLvw IIf(Node.Key = "R", 0, Mid(Node.Key, 3))
    If Lvw.ListItems.count > 0 Then
        Lvw.ListItems(1).Selected = True
        Lvw.SelectedItem.Selected = True
        Lvw_ItemClick Lvw.SelectedItem
        
        mnuFilePrint.Enabled = True
        mnuFilePreview.Enabled = True
        mnuFileExcel.Enabled = True
        Tbar.Buttons("Preview").Enabled = True
        Tbar.Buttons("Print").Enabled = True
    Else
        mnuFilePrint.Enabled = False
        mnuFilePreview.Enabled = False
        mnuFileExcel.Enabled = False
        Tbar.Buttons("Preview").Enabled = False
        Tbar.Buttons("Print").Enabled = False
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

