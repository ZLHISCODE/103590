VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "PACS���ݹ���"
   ClientHeight    =   6375
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   9210
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   9210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Timer timAutoPolicy 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7260
      Top             =   720
   End
   Begin MSComctlLib.ListView LivMain 
      Height          =   4845
      Left            =   210
      TabIndex        =   3
      Top             =   1020
      Width           =   8565
      _ExtentX        =   15108
      _ExtentY        =   8546
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList Ilscolor 
      Left            =   7710
      Top             =   570
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":030A
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":052A
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":074A
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":096A
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0B8A
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0DAA
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0FCA
            Key             =   "Start"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11EA
            Key             =   "View"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1406
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1626
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1846
            Key             =   "Hand"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B60
            Key             =   "Back"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D7A
            Key             =   "Auto"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1F94
            Key             =   "Filtrate"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList Ilsrw 
      Left            =   8310
      Top             =   570
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":21AE
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":23CE
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":25EE
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":280E
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2A2E
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2C4E
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2E6E
            Key             =   "Start"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":308E
            Key             =   "View"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":32AA
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":34CA
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":36EA
            Key             =   "Hand"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3A04
            Key             =   "Back"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3C1E
            Key             =   "Auto"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3E38
            Key             =   "Filtrate"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9210
      _ExtentX        =   16245
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   9210
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbrMain"
      MinHeight1      =   720
      Width1          =   8370
      Key1            =   "only"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrMain 
         Height          =   720
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   9090
         _ExtentX        =   16034
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "Ilsrw"
         HotImageList    =   "Ilscolor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   12
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
               Caption         =   "�ֶ�"
               Key             =   "ManualArchive"
               Object.ToolTipText     =   "�ֶ��鵵"
               Object.Tag             =   "�ֶ�"
               ImageKey        =   "Hand"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "ManualDeArchive"
               Object.ToolTipText     =   "�ֶ����鵵"
               Object.Tag             =   "����"
               ImageKey        =   "Back"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�Զ�"
               Key             =   "AutoArchiveSetup"
               Object.ToolTipText     =   "�Զ��鵵"
               Object.Tag             =   "�Զ�"
               ImageKey        =   "Auto"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split1"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Filtrate"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageKey        =   "Filtrate"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�鿴"
               Key             =   "View"
               Object.ToolTipText     =   "��Ա�鿴��ʽ"
               Object.Tag             =   "�鿴"
               ImageKey        =   "View"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "��ͼ��"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Сͼ��"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "�б�"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "��ϸ����"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Quit"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   6015
      Width           =   9210
      _ExtentX        =   16245
      _ExtentY        =   635
      SimpleText      =   $"frmMain.frx":4052
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMain.frx":4099
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11165
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
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   6750
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
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
      Begin VB.Menu mnusplit3 
         Caption         =   "-"
      End
      Begin VB.Menu mnufileexit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuEditManualArchive 
         Caption         =   "�ֶ��鵵(&M)"
      End
      Begin VB.Menu mnuEditManualDeArchive 
         Caption         =   "�ֶ����鵵(&D)"
      End
      Begin VB.Menu mnuEditAutoArchiveSetup 
         Caption         =   "�Զ��鵵(&A)"
      End
      Begin VB.Menu mnuEditCollate 
         Caption         =   "У��(&L)"
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
      Begin VB.Menu mnuViewSplit3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFilter 
         Caption         =   "����(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewSplit4 
         Caption         =   "-"
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
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public LastState As Integer

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long

Private Const WM_SYSCOMMAND = &H112
Private Const SC_MOVE = &HF010&
Private Const SC_RESTORE = &HF120&
Private Const SC_SIZE = &HF000&


Dim rsData As New ADODB.Recordset

Private WithEvents mobjIcon As clsTaskIcon  '������
Attribute mobjIcon.VB_VarHelpID = -1


Sub InitLiv()
    With LivMain
        .ColumnHeaders.Add , "A", "����"
        .ColumnHeaders.Add , "B", "Ӱ�����"
        .ColumnHeaders.Add , "C", "����"
        .ColumnHeaders.Add , "D", "��������"
        .ColumnHeaders.Add , "E", "Ӣ����"
        .ColumnHeaders.Add , "F", "�ձ�"
        .ColumnHeaders.Add , "G", "������"
        .ColumnHeaders.Add , "H", "ͼ����"
        .ColumnHeaders.Add , "I", "λ��һ"
        .ColumnHeaders.Add , "J", "λ�ö�"
        .ColumnHeaders.Add , "K", "���UID"
'        .ColumnHeaders.Add , "L", "����״̬"   '�鵵ʱ����û�и����б��е����ݹ鵵�����ǲ�ѯ�����ݿ�
    End With
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    With LivMain
        .Top = IIf(Me.CoolBar1.Visible, Me.CoolBar1.Height, 0)
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - IIf(Me.CoolBar1.Visible, Me.CoolBar1.Height, 0) - IIf(Me.stbThis.Visible, Me.stbThis.Height, 0)
    End With
    
    If WindowState <> vbMinimized Then
        LastState = WindowState
    End If
End Sub
Private Sub Form_Load()
    Dim strSQL As String
    Dim tmpset As ADODB.Recordset

    If WindowState = vbMinimized Then
        LastState = vbNormal
    Else
        LastState = WindowState
    End If

    Call RestoreWinState(Me, App.EXEName)

    '----------��������ͼ��
    Set mobjIcon = New clsTaskIcon
    mobjIcon.frmHwnd = tbrMain.hwnd ' hwnd
    mobjIcon.Icon = Icon.Handle
    mobjIcon.Message = "PACS���ݹ������"
    mobjIcon.AddIcon
    '----------��������ͼ��
    
'    AddToTray Me

    ''''gcnOracle
    '��ʹ���б�ؼ�
    Call InitLiv
    '��ע����ȡ�Զ��鵵����
    subReadPolicy
    '���ñ���������������
    beginDay = Date
    timAutoPolicy.Enabled = True
    Call ShowChkRecord

'    SetTrayTip "PACS���ݹ������"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    timAutoPolicy.Enabled = False
    
'    RemoveFromTray
    '�������ͼ��
    mobjIcon.DelIcon
    Set mobjIcon = Nothing
    
    Call SaveWinState(Me, App.EXEName)
End Sub

Private Sub mobjIcon_MouseLeftDBClick()
On Error Resume Next
    '����������ݿ����ʾ��־��ģʽ�����Ѿ����򿪣����˳���������ִ���
'    If mfrmUpdateDB Is Nothing And mfrmShowLog Is Nothing Then
        If WindowState <> 1 Then
            WindowState = vbMinimized
            Me.Hide
        Else
            WindowState = vbNormal
            Me.Show
        End If
'    End If
    Err.Clear
End Sub

Public Sub ShowChkRecord()
    Dim strSQL As String
    Dim tmpset As New ADODB.Recordset
    Dim strStorePlace As String         'ʹ�á�λ��һ������λ�ö�����ɸѡ
    Dim objItem As ListItem             '�б����
    '��ע�����ȡ��������������ʱ����
    Dim mDevice As String
    Dim mStorageDevice As String
    Dim mFStudy As String
    Dim mEStudy As String
    Dim mFTime As String
    Dim mETime As String
    Dim mArchiveState As String
    
    mDevice = GetSetting("ZLSOFT", "����ģ��\�鵵����\����", "Ӱ������", "��������")
    mStorageDevice = GetSetting("ZLSOFT", "����ģ��\�鵵����\����", "�����豸", cAllStorageDevice)
    mFStudy = GetSetting("ZLSOFT", "����ģ��\�鵵����\����", "��ʼ����", "")
    mEStudy = GetSetting("ZLSOFT", "����ģ��\�鵵����\����", "��������", "")
    mFTime = GetSetting("ZLSOFT", "����ģ��\�鵵����\����", "��ʼʱ��", zlDatabase.Currentdate - 90)
    mETime = GetSetting("ZLSOFT", "����ģ��\�鵵����\����", "����ʱ��", zlDatabase.Currentdate - 30)
    mArchiveState = GetSetting("ZLSOFT", "����ģ��\�鵵����\����", "�鵵״̬", "δ�鵵")
    
    With frmFilter
        strSQL = "Select Ӱ�����,����,��������,����,Ӣ����,�Ա�,Sum(1) As ������,Sum(ͼ����) As ͼ����,λ��һ,λ�ö�,b.���UID From" & _
            " (Select a.���UID,b.����UID,Sum(1) As ͼ���� from Ӱ�����¼ a,Ӱ�������� b,Ӱ����ͼ�� c" & _
            " Where a.���UID=b.���UID And b.����UID=c.����UID " & _
            IIf(mStorageDevice = cAllStorageDevice, "", IIf(mArchiveState = "δ�鵵", " And a.λ��һ=" & mStorageDevice, " And a.λ�ö�=" & mStorageDevice)) & _
            IIf(mArchiveState = "δ�鵵", " And a.λ�ö� is null ", IIf(mArchiveState = "�ѹ鵵��ɾ��", " And a.λ��һ is null ", " And a.λ��һ is not null And a.λ�ö� is not null")) & _
            IIf(mDevice = "��������", "", " And a.Ӱ�����='" & mDevice & "'") & _
            IIf(mFTime = "3000-01-01", "", " And a.��������>=to_Date('" & Format(mFTime, "yyyy-MM-dd HH:mm:SS") & "','YYYY-MM-DD HH24:Mi:SS')") & _
            IIf(mETime = "3000-01-01", "", " And a.��������<=to_Date('" & Format(mETime, "yyyy-MM-dd HH:mm:SS") & "','YYYY-MM-DD HH24:Mi:SS')") & _
            IIf(Len(Trim(mFStudy)) = 0 Or Not IsNumeric(mFStudy), "", " And a.����>=" & mFStudy) & _
            IIf(Len(Trim(mEStudy)) = 0 Or Not IsNumeric(mEStudy), "", " And a.����<=" & mEStudy) & _
            " Group By a.���UID,b.����UID) a, Ӱ�����¼ b Where a.���UID=b.���UID Group By Ӱ�����,����,��������,����,Ӣ����,�Ա�,λ��һ,λ�ö�,b.���UID"
    End With
    On Error GoTo errH
    zlDatabase.OpenRecordset rsData, strSQL, Me.Caption
    LivMain.ListItems.Clear
    Do Until rsData.EOF
        With LivMain
            Set objItem = .ListItems.Add(, "A" & rsData("����") & "UID:" & rsData("���UID"), ZlCommFun.NVL(rsData("����")))
            objItem.SubItems(1) = IIf(IsNull(rsData("����")), "", rsData("����"))
            objItem.SubItems(2) = rsData("��������")
            objItem.SubItems(3) = rsData("Ӱ�����")
            objItem.SubItems(4) = ZlCommFun.NVL(rsData("Ӣ����"), "UnKnow")
            objItem.SubItems(5) = ZlCommFun.NVL(rsData("�Ա�"), "δ֪")
            objItem.SubItems(6) = rsData("������")
            objItem.SubItems(7) = rsData("ͼ����")
            objItem.SubItems(8) = IIf(IsNull(rsData("λ��һ")), "", rsData("λ��һ"))
            objItem.SubItems(9) = IIf(IsNull(rsData("λ�ö�")), "", rsData("λ�ö�"))
            objItem.SubItems(10) = rsData("���UID")
        End With
        rsData.MoveNext
    Loop
    
    '��ʾ
    Me.stbThis.Panels(2).Text = "��ǰ<" & mArchiveState & ">״̬������" & rsData.RecordCount & "����¼��"
    
    '���ι������Ͳ˵���ť
    If Me.LivMain.ListItems.Count > 0 Then
        Select Case mArchiveState
            Case "δ�鵵"
                Me.mnuEditManualArchive.Enabled = True
                Me.mnuEditManualDeArchive.Enabled = False
                Me.tbrMain.Buttons("ManualArchive").Enabled = True
                Me.tbrMain.Buttons("ManualDeArchive").Enabled = False
            Case "�ѹ鵵��ɾ��"
                Me.mnuEditManualArchive.Enabled = False
                Me.mnuEditManualDeArchive.Enabled = True
                Me.tbrMain.Buttons("ManualArchive").Enabled = False
                Me.tbrMain.Buttons("ManualDeArchive").Enabled = True
            Case "�ѹ鵵δɾ��"
                Me.mnuEditManualArchive.Enabled = True
                Me.mnuEditManualDeArchive.Enabled = True
                Me.tbrMain.Buttons("ManualArchive").Enabled = True
                Me.tbrMain.Buttons("ManualDeArchive").Enabled = True
        End Select
    Else
        Me.mnuEditManualArchive.Enabled = False
        Me.mnuEditManualDeArchive.Enabled = False
        Me.tbrMain.Buttons("ManualArchive").Enabled = False
        Me.tbrMain.Buttons("ManualDeArchive").Enabled = False
    End If
        
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Function funcdoArchiveJob(Optional lngJobID As Long = 0) As Boolean
    'ִ�й鵵��ҵ
    ''''''''''''''''''''''''''''''''''''''
    ''' ����:lngJobID----��Ҫִ�е���ҵID,�������0�������ҵ���ݿ��м�������һ����ҵ��ִ��
    ''''''''''''''''''''''''''''''''''''''
    
    Dim tmpset As ADODB.Recordset
    Dim dsArchive As ADODB.Recordset
    Dim strSQL As String
    Dim strSourceDev As String, strDestinationDev As String, strAppointDev As String
    Dim bMove As Boolean, bDelete As Boolean, bAutoBackup As Boolean
    Dim strFilter As String
    
    funcdoArchiveJob = False
    '����Զ�ִ����ҵ���������ȡ���ݿ⣬��ȡ�鵵����
    If lngJobID = 0 Then        '����Ӱ��鵵��ҵ����ȡ��һ����ִ�е���ҵID
        strSQL = "Select ���� From Ӱ��鵵��ҵ where ִ�й��� = 0 "
        Set tmpset = gcnOracle.Execute(strSQL)
        If Not tmpset.EOF Then
            lngJobID = tmpset!����
        Else
            Exit Function       'û�п�ִ�е���ҵ���˳�����
        End If
    End If
    '��д���ݿ⣬��д��ʼ�鵵��ʶ�Ϳ�ʼʱ��
    strSQL = "update Ӱ��鵵��ҵ set ��ʼʱ�� = to_date('" & Date & " " & Time & "','yyyy-mm-dd hh24:mi:ss')," & _
             "ִ�й��� = 1 where ����=" & lngJobID
    gcnOracle.Execute (strSQL)
    
    '��ȡ�鵵����
    strSQL = "select Դ�豸,Ŀ���豸,ָ���豸,�Ƿ�Ǩ��,�Ƿ�ɾ��,�Զ�����,�������� from Ӱ��鵵��ҵ where ����=" & lngJobID
    Set tmpset = gcnOracle.Execute(strSQL)
    strSourceDev = tmpset!Դ�豸
    strDestinationDev = tmpset!Ŀ���豸
    If IsNull(tmpset!ָ���豸) Then
        strAppointDev = vbNullString
    Else
        strAppointDev = tmpset!ָ���豸
    End If
    bMove = tmpset!�Ƿ�Ǩ��
    bDelete = tmpset!�Ƿ�ɾ��
    bAutoBackup = tmpset!�Զ�����
    If IsNull(tmpset!��������) Then
        strFilter = vbNullString
    Else
        strFilter = tmpset!��������
    End If
    
    '�鵵����
    '��ȡ�鵵���ݼ�¼
    If bAutoBackup Then     '�Զ��鵵����Ҫ�����������м���
        strSQL = "Select ��������,λ��һ,λ�ö�,���UID From Ӱ�����¼ where not λ��һ is null and λ�ö� is null"
        Set dsArchive = gcnOracle.Execute(strSQL)
        '��Ҫ�Լ����������н���
    Else                    '�ֶ��鵵��ֱ�ӻ�ȡԭ�������ļ�¼
        Set dsArchive = rsData
    End If
    
    '���ú���ִ�й鵵����
    funcdoArchiveJob = funcArchiveExec(strSourceDev, strDestinationDev, strAppointDev, bMove, bDelete, dsArchive)
    
    '�鵵��ɣ���д���ݿ⣬��д�鵵��ɱ�ʶ�����ʱ��
    
    If funcdoArchiveJob = True Then
        strSQL = "update Ӱ��鵵��ҵ set ����ʱ�� = to_date('" & Date & " " & Time & "','yyyy-mm-dd hh24:mi:ss')," & _
                 "ִ�й��� = 2 where ����=" & lngJobID
    
    Else            '��д��ʶ��ִ������ʧ��
        strSQL = "update Ӱ��鵵��ҵ set ����ʱ�� = to_date('" & Date & " " & Time & "','yyyy-mm-dd hh24:mi:ss')," & _
                 "ִ�й��� = 3 where ����=" & lngJobID
    End If
    gcnOracle.Execute (strSQL)
End Function

Private Function funcArchiveExec(strSourceDev As String, strDestinationDev As String, strAppointDevID As String, _
                                 bMove As Boolean, bDelete As Boolean, dsArchive As ADODB.Recordset) As Boolean
    Dim cDevice As New Collection       '����ȫ���鵵�豸��Ϣ�ļ���
    Dim clsOneDevice As clsBakDevice    '�ݴ�һ���豸��Ϣ����
    Dim strSQL As String                '������Ҫִ�е���ʱSQL���
    Dim tmpset As ADODB.Recordset       '����SQL����ִ�н�����ݼ�
    Dim strTempDir As String            '������ʱĿ¼
    Dim strLocalIP As String            '����IP��ַ
    Dim strLocalDirDest As String       'Ŀ���豸����Ŀ¼
    Dim lngResult As Long               '�洢����ִ�еķ�����Ϣ
    Dim strDestDevID As String          'Ŀ���豸��ID
    Dim i As Integer                    'ͨ��ѭ��������
    
    If dsArchive.RecordCount <= 0 Then      '���ݼ�Ϊ�գ�ֱ���˳���������ʶ�������
        funcArchiveExec = True
        Exit Function
    End If
    funcArchiveExec = False
    
    'ʹ��һ�����������汻ʹ�õ���Դ�豸IP���û���������
    strSQL = "select �豸��,�豸��,����,IP��ַ,FTPĿ¼,ftp�û���,ftp����,״̬,����Ŀ¼ from Ӱ���豸Ŀ¼ WHERE " & _
             "���� = 1"
    Set tmpset = gcnOracle.Execute(strSQL)
    While Not tmpset.EOF
        With tmpset
            Set clsOneDevice = New clsBakDevice
            clsOneDevice.strDevID = IIf(IsNull(!�豸��), "", !�豸��)
            clsOneDevice.strDevName = IIf(IsNull(!�豸��), "", !�豸��)
            clsOneDevice.lngType = IIf(IsNull(!����), "", !����)
            clsOneDevice.strIP = IIf(IsNull(!ip��ַ), "", !ip��ַ)
            clsOneDevice.strPasswd = IIf(IsNull(!ftp����), "", !ftp����)
            clsOneDevice.strUser = IIf(IsNull(!ftp�û���), "", !ftp�û���)
            clsOneDevice.strVirtualPath = IIf(IsNull(!FTPĿ¼), "", !FTPĿ¼)
            clsOneDevice.strLocalPath = IIf(IsNull(!����Ŀ¼), "", !����Ŀ¼)
            clsOneDevice.lngStatus = IIf(IsNull(!״̬), "", !״̬)
            cDevice.Add clsOneDevice, clsOneDevice.strDevID
            .MoveNext
        End With
    Wend
    
    '��ȡ��������ʱ·��
    strTempDir = Environ("TEMP") & "\zlPacs"
    If Dir(strTempDir, vbDirectory) = vbNullString Then
        MkDir strTempDir
    End If
    '��ȡ����IP��ַ
    strLocalIP = Winsock1.LocalIP
    '�趨Ŀ���豸��
    If strAppointDevID = vbNullString Then '�Զ�ѡ��
        For i = 1 To cDevice.Count
            If cDevice(i).lngStatus = 1 Then
                strDestDevID = cDevice(i).strDevID
                Exit For
            End If
        Next
    Else
        strDestDevID = strAppointDevID
    End If
    '���Ŀ���豸�Ǳ������趨����Ŀ���豸Ŀ¼
    strLocalDirDest = vbNullString
    If UCase(cDevice(strDestDevID).strIP) = "LOCALHOST" Or cDevice(strDestDevID).strIP = strLocalIP Then
        strLocalDirDest = cDevice(strDestDevID).strLocalPath
    End If
    
    '��ʼ���й鵵����
    dsArchive.MoveFirst
    While Not dsArchive.EOF
        
        ''''���ö�һ����¼���й鵵��ʵ�ʲ���'''''
        lngResult = funcArchiveOneRecord(dsArchive, cDevice, strLocalIP, strTempDir, strLocalDirDest, _
                             strSourceDev, strDestinationDev, strDestDevID, bMove, bDelete)
        Select Case lngResult
        Case 0            '�ɹ���ɣ�����ƶ�һ����¼
            dsArchive.MoveNext
        Case 1, 2            '1--FTP����ʧ�ܣ�ת�Ƶ���һ���豸;2--'FTPǨ��ʧ�ܣ���ǵ�ǰ�豸����ת����һ���豸
            If lngResult = 2 Then           '��ʶ�豸��
                strSQL = "update Ӱ���豸Ŀ¼ set Ŀ¼�� = 1 where �豸��='" & strDestDevID & "'"
                gcnOracle.Execute strSQL
            End If
            'ת�Ƶ���һ���豸
            strDestDevID = vbNullString
            For i = 1 To cDevice.Count
                If cDevice(i).lngStatus = 1 Then '�����豸
                    strDestDevID = cDevice(i).strDevID
                    '���Ŀ���豸�Ǳ������趨����Ŀ���豸Ŀ¼
                    If UCase(cDevice(strDestDevID).strIP) = "LOCALHOST" Or cDevice(strDestDevID).strIP = strLocalIP Then
                        strLocalDirDest = cDevice(strDestDevID).strLocalPath
                    Else
                        strLocalDirDest = vbNullString
                    End If
                    Exit For
                End If
            Next
            If strDestDevID = vbNullString Then         '���豸���ã���ʾ���˳�
                If lngResult = 1 Then
                    MsgBox "û�п������ӵ��豸���������������ӹ��ϣ������豸�����"
                Else
                    MsgBox "û�п������ӵ��豸���������豸�洢�������������豸�����"
                End If
                Exit Function
            End If
        Case 3             'FTPɾ��ʧ�ܣ���¼����־�У�����ƶ�һ����¼
            '''''''''��¼����־��'''''''''''''''
            dsArchive.MoveNext
        Case 4             'δ֪����ֱ����ʾ�û����жϲ���
            MsgBox "�����쳣���󣬹鵵������ֹ,�����Ǳ��������������㣬�������ú����ԡ�"
            Exit Function
        Case 5
            dsArchive.MoveNext
        End Select
    Wend
    funcArchiveExec = True
End Function
              
Private Function funcArchiveOneRecord(dsArchive As ADODB.Recordset, cDevice As Collection, _
                 strLocalIP As String, strTempDir As String, strLocalDirDest As String, _
                 strSourceDev As String, strDestinationDev As String, strDestDevID As String, _
                 bMove As Boolean, bDelete As Boolean) As Long
''''''''''''''''''''''''''''''''''''''''''''''''''
'''���ܣ���һ����¼ִ�й鵵��Ǩ�ƺ�ɾ��������
'''���أ��ɹ�����0������ʧ�ܷ���1��Ǩ��ʧ�ܷ���2��ɾ��ʧ�ܷ���3����������4
''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim i As Integer                    '������
    Dim iImgCount As Integer            'ͼ�������
    Dim objFileSystem As Object         '�������е��ļ����ƺ���
    Set objFileSystem = CreateObject("Scripting.FileSystemObject")
    Dim aImageFiles() As String         'ͼ���ļ�������
    Dim strImgPath As String            'ͼ���м�·������ ���ɼ�����\���UID�����
    Dim strFTPImgPath As String         'FTPʹ�õ�ͼ���м�·������ ���ɼ�����/���UID�����
    Dim clsFTPsubs As New clsFTP        'Ϊ�˵���clsFTP������ĺ���
    Dim strSQL As String                '������Ҫִ�е���ʱSQL���
    Dim tmpset As ADODB.Recordset       '����SQL����ִ�н�����ݼ�
    Dim strLocalDirSource As String     'Դ�豸����Ŀ¼
    Dim lngResult As Long               '���淵��ֵ
    Dim strSourceDevID As String        'Դ�豸��ID
    Dim lngFilesCount As Long           '������Ҫ�鵵���ƶ����ļ�����Ŀ
    Dim aRptImgFiles() As String        '����ͼ���ļ�������
    Dim aOtherFiles() As String         '���������ļ������飬��������ͼ��¼�������
    
    funcArchiveOneRecord = 1
    '�ж�Դ�豸���Ƿ��б����豸��ͨ���Ƚ�IP��ַʵ��
    strSourceDevID = IIf(strSourceDev = "1", IIf(IsNull(dsArchive!λ��һ), "", dsArchive!λ��һ), IIf(IsNull(dsArchive!λ�ö�), "", dsArchive!λ�ö�))
    If UCase(cDevice(strSourceDevID).strIP) = "LOCALHOST" Or cDevice(strSourceDevID).strIP = strLocalIP Then
        '����б����豸����ֱ��ʹ�ñ���Ŀ¼
        strLocalDirSource = cDevice(strSourceDevID).strLocalPath
    Else
        strLocalDirSource = vbNullString
    End If
    
    '������ͼ��·��
    strFTPImgPath = Format(dsArchive!��������, "yyyymmdd") & "/" & dsArchive!���uid
    strImgPath = Format(dsArchive!��������, "yyyymmdd") & "\" & dsArchive!���uid
    
    '��ѯ���ݿ⣬��ȡͼ���ļ�Ŀ¼���ļ���
    strSQL = "select ͼ��UID from Ӱ����ͼ�� a ,Ӱ�������� b where b.����UID = a.����UID and b.���UID = '" & _
             dsArchive!���uid & "'"
    Set tmpset = gcnOracle.Execute(strSQL)
    ReDim aImageFiles(tmpset.RecordCount) As String
    i = 1
    While Not tmpset.EOF
        aImageFiles(i) = tmpset!ͼ��uid
        i = i + 1
        tmpset.MoveNext
    Wend
    
    '��ѯ���ݿ⣬��ȡ����ͼ���ļ���,¼�������ļ����������ļ���
    strSQL = "select ����ͼ�� from Ӱ�����¼ where ���UID = '" & dsArchive!���uid & "'"
    Set tmpset = gcnOracle.Execute(strSQL)
    If Not IsNull(tmpset!����ͼ��) Then
        aOtherFiles = Split(tmpset!����ͼ��, "|")   '���ֳ�����ͼ���¼������
        If UBound(aOtherFiles) > -1 Then            '�б���ͼ���¼�����棬�����ļ�����ӵ�ͼ���ļ���������
            For i = 0 To UBound(aOtherFiles)
                aRptImgFiles = Split(aOtherFiles(i), ";")
                lngFilesCount = UBound(aImageFiles)
                If UBound(aRptImgFiles) > -1 Then   '�б���ͼ���¼������
                    ReDim Preserve aImageFiles(lngFilesCount + UBound(aRptImgFiles) + 1) As String
                    For iImgCount = 0 To UBound(aRptImgFiles)
                        aImageFiles(iImgCount + 1 + lngFilesCount) = Trim(aRptImgFiles(iImgCount))
                    Next
                End If
            Next
        End If
    End If
    
    If bMove Then       '���й鵵����
        If strLocalDirSource <> vbNullString And strLocalDirDest <> vbNullString Then   '�������ļ�����
            '��Ŀ���豸����Ŀ¼
            If Dir(strLocalDirDest & "\" & Left(strImgPath, InStr(strImgPath, "\") - 1), vbDirectory) = vbNullString Then
                MkDir strLocalDirDest & "\" & Left(strImgPath, InStr(strImgPath, "\") - 1)
            End If
            If Dir(strLocalDirDest & "\" & strImgPath, vbDirectory) = vbNullString Then
                MkDir strLocalDirDest & "\" & strImgPath
            End If
            For i = 1 To UBound(aImageFiles)
                objFileSystem.CopyFile strLocalDirSource & "\" & strImgPath & "\" & aImageFiles(i), _
                                       strLocalDirDest & "\" & strImgPath & "\"
            Next
            '�����ļ�
        ElseIf strLocalDirSource <> vbNullString Then       '�ӱ���Ŀ¼ֱ���ϴ�
            clsFTPsubs.strIPAddress = cDevice(strDestDevID).strIP
            clsFTPsubs.strPsw = cDevice(strDestDevID).strPasswd
            clsFTPsubs.strUser = cDevice(strDestDevID).strUser
            lngResult = clsFTPsubs.FuncFtpMkDir(cDevice(strDestDevID).strVirtualPath, strFTPImgPath)
            lngResult = clsFTPsubs.FuncUpLoadFiles(cDevice(strDestDevID).strVirtualPath & "/" & strFTPImgPath, strLocalDirSource & "\" & strImgPath, aImageFiles)
            If lngResult <> 0 Then
                funcArchiveOneRecord = lngResult
                Exit Function
            End If
        ElseIf strLocalDirDest <> vbNullString Then         'ֱ�����ص�����
            '��Ŀ���豸����Ŀ¼
            MkDir strLocalDirDest & "\" & Left(strImgPath, InStr(strImgPath, "\") - 1)
            MkDir strLocalDirDest & "\" & strImgPath
            clsFTPsubs.strIPAddress = cDevice(strSourceDevID).strIP
            clsFTPsubs.strPsw = cDevice(strSourceDevID).strPasswd
            clsFTPsubs.strUser = cDevice(strSourceDevID).strUser
            lngResult = clsFTPsubs.FuncDownLoadFiles(cDevice(strSourceDevID).strVirtualPath & "/" & strFTPImgPath, strLocalDirDest & "\" & strImgPath, aImageFiles)
            If lngResult <> 0 Then
                If lngResult = 1 Then
                    funcArchiveOneRecord = lngResult
                Else
                    funcArchiveOneRecord = 4        '�����������󣬿����Ǳ���������������
                End If
                Exit Function
            End If
        Else        'ʹ����ʱĿ¼����ת
            '���û�б����豸����ʹ��ϵͳ��ʱĿ¼��Ϊ��ת
            clsFTPsubs.strIPAddress = cDevice(strSourceDevID).strIP
            clsFTPsubs.strPsw = cDevice(strSourceDevID).strPasswd
            clsFTPsubs.strUser = cDevice(strSourceDevID).strUser
            lngResult = clsFTPsubs.FuncDownLoadFiles(cDevice(strSourceDevID).strVirtualPath & "/" & strFTPImgPath, strTempDir, aImageFiles)
            If lngResult <> 0 Then
                If lngResult = 1 Then
                    funcArchiveOneRecord = lngResult
                Else
                    funcArchiveOneRecord = 5        '�����������󣬿���δ�ܶ�ȡԴ�ļ�
                End If
                Exit Function
            End If
            clsFTPsubs.strIPAddress = cDevice(strDestDevID).strIP
            clsFTPsubs.strPsw = cDevice(strDestDevID).strPasswd
            clsFTPsubs.strUser = cDevice(strDestDevID).strUser
            lngResult = clsFTPsubs.FuncFtpMkDir(cDevice(strDestDevID).strVirtualPath, strFTPImgPath)
            lngResult = clsFTPsubs.FuncUpLoadFiles(cDevice(strDestDevID).strVirtualPath & "/" & strFTPImgPath, strTempDir, aImageFiles)
            If lngResult <> 0 Then
                funcArchiveOneRecord = lngResult
                Exit Function
            End If
            For i = 1 To UBound(aImageFiles) 'ɾ����ת�ļ�
                objFileSystem.DeleteFile strTempDir & "\" & aImageFiles(i)
            Next
        End If
    End If
    
    '����ɾ������
    If bDelete Then
        If strLocalDirSource <> vbNullString Then       'ɾ�������ļ���Ŀ¼
            For i = 1 To UBound(aImageFiles)        'ɾ���ļ�
                objFileSystem.DeleteFile strLocalDirSource & "\" & strImgPath & "\" & aImageFiles(i)
            Next
            'ɾ��Ŀ¼
            '���Ŀ¼�Ƿ�Ϊ��,��ɾ��
            If (Dir(strLocalDirSource & "\" & strImgPath & "\*.*") = vbNullString) Then
                objFileSystem.DeleteFolder strLocalDirSource & "\" & strImgPath
            End If
            If (Dir(strLocalDirSource & "\" & Left(strImgPath, InStr(strImgPath, "\") - 1) & "\*.*") = vbNullString) Then
                objFileSystem.DeleteFolder strLocalDirSource & "\" & Left(strImgPath, InStr(strImgPath, "\") - 1)
            End If
        Else            'ɾ��FTP�ļ���Ŀ¼
            clsFTPsubs.strIPAddress = cDevice(strSourceDevID).strIP
            clsFTPsubs.strPsw = cDevice(strSourceDevID).strPasswd
            clsFTPsubs.strUser = cDevice(strSourceDevID).strUser
            'ɾ���ļ�
            lngResult = clsFTPsubs.FuncDelFiles(cDevice(strSourceDevID).strVirtualPath & "/" & strFTPImgPath, aImageFiles)
            If lngResult <> 0 Then
                If lngResult = 1 Then
                    funcArchiveOneRecord = lngResult
                Else
                    funcArchiveOneRecord = 3        '����ɾ��ʧ��
                End If
                Exit Function
            End If
            'ɾ��Ŀ¼
            lngResult = clsFTPsubs.FuncFtpDelDir(cDevice(strSourceDevID).strVirtualPath, strFTPImgPath)
            lngResult = clsFTPsubs.FuncFtpDelDir(cDevice(strSourceDevID).strVirtualPath, Left(strFTPImgPath, InStr(strFTPImgPath, "/") - 1))
        End If
    End If
    '��Ӱ�����¼������д�鵵���
    If bMove = True And bDelete = True Then
        strSQL = IIf(strDestinationDev = "2", "λ�ö� = '" & strDestDevID & "'", "λ��һ = '" & strDestDevID & "'") & _
                 IIf(strSourceDev = "2", " , λ�ö� =null", " , λ��һ = null")
    ElseIf bMove = True Then
        strSQL = IIf(strDestinationDev = "2", "λ�ö� = '" & strDestDevID & "'", "λ��һ = '" & strDestDevID & "'")
    ElseIf bDelete = True Then
        strSQL = IIf(strSourceDev = "2", " λ�ö� =null", " λ��һ = null")
    End If
    strSQL = "update Ӱ�����¼ set  " & strSQL & " where ���UID = '" & dsArchive!���uid & "'"
    gcnOracle.Execute (strSQL)
    funcArchiveOneRecord = 0        '������������������
End Function

Private Sub subManualArchive()
    '�ֶ��鵵����
    Dim tmpset As ADODB.Recordset
    Dim strSQL As String
    frmManualArchive.Caption = "�ֶ��鵵"
    '���鵵�豸
    frmManualArchive.cobDevice.Clear
    strSQL = "Select �豸��,�豸�� From Ӱ���豸Ŀ¼ where ���� = 1 and ״̬ = 1 and NVL(Ŀ¼��,0) = 0"
    Set tmpset = gcnOracle.Execute(strSQL)
    With tmpset
        While Not .EOF
            frmManualArchive.cobDevice.AddItem !�豸�� & "-" & !�豸��
            .MoveNext
        Wend
    End With
    frmManualArchive.cobDevice.ListIndex = IIf(frmManualArchive.cobDevice.ListCount > 0, 0, -1)
    frmManualArchive.cobManualMoveDelete.ListIndex = 0
    frmManualArchive.sstabManualArchive.Tab = 0
    frmManualArchive.cmdStep3.Caption = "��ʼ�鵵"
    frmManualArchive.bArchive = True
    frmManualArchive.Show 1, Me
End Sub

Private Sub subManualDeArchive()
    
    '�ַ����鵵����
    Dim tmpset As ADODB.Recordset
    Dim strSQL As String
    frmManualArchive.Caption = "�ֶ����鵵"
    '���鵵�豸
    frmManualArchive.cobDevice.Clear
    strSQL = "Select �豸��,�豸�� From Ӱ���豸Ŀ¼ where ���� = 1 and NVL(Ŀ¼��,0) = 0"
    Set tmpset = gcnOracle.Execute(strSQL)
    With tmpset
        While Not .EOF
            frmManualArchive.cobDevice.AddItem !�豸�� & "-" & !�豸��
            'frmManualArchive.cobDevice.ItemData(frmManualArchive.cobDevice.NewIndex) = !�豸��
            .MoveNext
        Wend
    End With
    frmManualArchive.cobDevice.ListIndex = IIf(frmManualArchive.cobDevice.ListCount > 0, 0, -1)
    frmManualArchive.cobManualMoveDelete.ListIndex = 0
    frmManualArchive.sstabManualArchive.Tab = 0
    frmManualArchive.cmdStep3.Caption = "��ʼ���鵵"
    frmManualArchive.bArchive = False   '��ʶ���鵵
    frmManualArchive.Show 1, Me
End Sub

Private Sub mnuEditAutoArchiveSetup_Click()
    Call subAutoArchive
End Sub

Private Sub mnuEditCollate_Click()
    frmCollate.Show vbModal, Me
End Sub

Private Sub mnuEditFilter_Click()
    If frmFilter.ShowMe(Me) = True Then
        Call ShowChkRecord
    End If
End Sub

Private Sub mnuEditManualArchive_Click()
    Call subManualArchive
End Sub

Private Sub mnuEditManualDeArchive_Click()
    Call subManualDeArchive
End Sub

Private Sub mnuFileExcel_Click()
    subPrint 3
End Sub

Private Sub mnufileexit_Click()
    Unload Me
End Sub

Private Sub mnuFilePreview_Click()
    subPrint 2
End Sub

Private Sub mnuFilePrint_Click()
    subPrint 1
End Sub

Private Sub mnuFileSet_Click()
    zlPrintSet
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpTopic_Click()
'    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hwnd)
End Sub

Private Sub mnuViewIcon_Click(Index As Integer)
    Dim i As Integer
    For i = 0 To 3
        mnuViewIcon(i).Checked = False
    Next
    mnuViewIcon(Index).Checked = True
    LivMain.View = Index
    
End Sub

Private Sub mnuViewReflash_Click()
    ShowChkRecord
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    CoolBar1.Visible = mnuViewToolButton.Checked
    CoolBar1.Bands("only").MinHeight = tbrMain.Height
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim buttTemp As Button

    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For Each buttTemp In tbrMain.Buttons
        If mnuViewToolText.Checked Then
            buttTemp.Caption = buttTemp.Tag
        Else
            buttTemp.Caption = ""
        End If
    Next
    CoolBar1.Bands("only").MinHeight = tbrMain.Height
    Form_Resize
End Sub

Private Sub timAutoPolicy_Timer()
    Dim aTempPolicy() As String         '�����������Ե�����
    Dim strSQL As String                '�ݴ�SQL���
    Dim tmpset As ADODB.Recordset       '�ݴ��ѯ���ݼ�
    Dim lngJobNum As Long               '�µĹ鵵��ҵ��
    
    '���ʱ��
    If strTimePolicy = vbNullString Then Exit Sub
    
    aTempPolicy = Split(strTimePolicy, ",")
    If aTempPolicy(1) <> "N/A" Then         '��ʱ����ԣ�����ʱ�����
        '�ж��Ƿ�����鵵����
        If UCase(aTempPolicy(1)) = "DAY" Then      '����ÿ����ԣ������������ʱ��
            If (Date - beginDay) >= aTempPolicy(2) Then     '��⵱ǰʱ��
                If Time = CDate(aTempPolicy(3)) Then
                    '���һ�����ݼ�¼
                    strSQL = "select Ӱ��鵵��ҵ_ID.nextval as JobID from dual"
                    Set tmpset = gcnOracle.Execute(strSQL)
                    lngJobNum = tmpset!JobID
                    strSQL = "Insert into Ӱ��鵵��ҵ (����,����,ִ��ʱ��,Դ�豸,Ŀ���豸,ָ���豸,�Ƿ�Ǩ��,�Ƿ�ɾ��,�Զ�����,ִ�й���) values (" & _
                             lngJobNum & ",'�Զ�" & lngJobNum & "',to_date('" & Date & " " & Time & "','yyyy-mm-dd hh24:mi:ss') " & _
                             ",'1','2',''," & aTempPolicy(5) & "," & aTempPolicy(4) & ",1,0)"
                    gcnOracle.Execute (strSQL)
        
                    '֪ͨ���ݳ���    'ִ�й鵵��ҵ
                    frmMain.funcdoArchiveJob lngJobNum
                    '�޸�beginDayΪ����
                    beginDay = Date
                End If
            End If
        ElseIf UCase(aTempPolicy(1)) = "MONTH" Then     '����ÿ�²��ԣ��������Ƿ�鵵����
            If Day(Date) = aTempPolicy(2) Then     '���ǹ鵵����
                If Time = CDate(aTempPolicy(3)) Then    '���ǹ鵵ʱ��
                    '���һ�����ݼ�¼
                    strSQL = "select Ӱ��鵵��ҵ_ID.nextval as JobID from dual"
                    Set tmpset = gcnOracle.Execute(strSQL)
                    lngJobNum = tmpset!JobID
                    strSQL = "Insert into Ӱ��鵵��ҵ (����,����,ִ��ʱ��,Դ�豸,Ŀ���豸,ָ���豸,�Ƿ�Ǩ��,�Ƿ�ɾ��,�Զ�����,ִ�й���) values (" & _
                             lngJobNum & ",'�Զ�" & lngJobNum & "',to_date('" & Date & " " & Time & "','yyyy-mm-dd hh24:mi:ss') " & _
                             ",'1','2',''," & aTempPolicy(5) & "," & aTempPolicy(4) & ",1,0)"
                    gcnOracle.Execute (strSQL)
        
                    '֪ͨ���ݳ���    'ִ�й鵵��ҵ
                    frmMain.funcdoArchiveJob lngJobNum
                End If
            End If
        End If
    End If
End Sub

Public Sub subReadPolicy()
    strTimePolicy = GetSetting("ZLSOFT", "����ģ��\�鵵����", "ʱ��鵵����")
    bAutoArchive = IIf(GetSetting("ZLSOFT", "����ģ��\�鵵����", "ʹ���Զ��鵵", "False") = "True", True, False)
End Sub

Private Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "ManualArchive"
            '�ֶ��鵵
            mnuEditManualArchive_Click
        Case "ManualDeArchive"
            '�ֶ����鵵
            mnuEditManualDeArchive_Click
        Case "AutoArchiveSetup"
            '�Զ��鵵
            mnuEditAutoArchiveSetup_Click
        Case "Filtrate"
            '����
            mnuEditFilter_Click
        Case "Preview"
            'Ԥ��
            subPrint 2
        Case "Print"
            '��ӡ
            subPrint 1
        Case "View"
            '�鿴
            mnuViewIcon(LivMain.View).Checked = False
            If LivMain.View = 3 Then
                mnuViewIcon(0).Checked = True
                LivMain.View = 0
            Else
                mnuViewIcon(LivMain.View + 1).Checked = True
                LivMain.View = LivMain.View + 1
            End If
        Case "Help"
            '����
            mnuHelpTopic_Click
        Case "Quit"
            '�˳�
            mnufileexit_Click
    End Select
End Sub
Private Sub subPrint(bytMode As Byte)
'����:���д�ӡ,Ԥ���������EXCEL
'����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
'    If gstrUserName = "" Then Call GetUserInfo
    Dim objPrint As Object
    
    Set objPrint = New zlPrintLvw
    objPrint.Title.Text = "PACS���ݹ���"
    Set objPrint.Body.objData = LivMain
    objPrint.BelowAppItems.Add "��ӡ�ˣ�" & gstrUserName
    objPrint.BelowAppItems.Add "��ӡʱ�䣺" & Format(zlDatabase.Currentdate, "yyyy��MM��dd��")
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
Private Sub subAutoArchive()
    '��ʾ�Զ��������ý���
    frmAutoArchive.Show 1, Me
End Sub

Private Sub tbrMain_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Dim i As Integer
    For i = 0 To 3
        mnuViewIcon(i).Checked = False
    Next
    mnuViewIcon(ButtonMenu.Index - 1).Checked = True
    LivMain.View = ButtonMenu.Index - 1
End Sub

Private Sub tbrMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mobjIcon.MouseState X
End Sub
