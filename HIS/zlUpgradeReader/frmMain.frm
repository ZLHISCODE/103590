VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   Caption         =   "����˵���Ķ���"
   ClientHeight    =   7815
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   13575
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmMain.frx":6852
   ScaleHeight     =   7815
   ScaleWidth      =   13575
   StartUpPosition =   2  '��Ļ����
   WindowState     =   2  'Maximized
   Begin XtremeReportControl.ReportControl rptList 
      Height          =   5310
      Left            =   2655
      TabIndex        =   1
      Top             =   2145
      Width           =   7260
      _Version        =   589884
      _ExtentX        =   12806
      _ExtentY        =   9366
      _StockProps     =   0
      ShowGroupBox    =   -1  'True
      AutoColumnSizing=   0   'False
   End
   Begin VB.PictureBox pic�������� 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2010
      Left            =   11550
      ScaleHeight     =   2010
      ScaleWidth      =   2550
      TabIndex        =   18
      Top             =   5730
      Width           =   2550
      Begin RichTextLib.RichTextBox txt�������� 
         Height          =   525
         Left            =   510
         TabIndex        =   26
         Top             =   510
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   926
         _Version        =   393217
         Appearance      =   0
         TextRTF         =   $"frmMain.frx":D0A4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.PictureBox pic˵�� 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2010
      Left            =   10305
      ScaleHeight     =   2010
      ScaleWidth      =   2550
      TabIndex        =   17
      Top             =   4605
      Width           =   2550
      Begin RichTextLib.RichTextBox txt˵�� 
         Height          =   525
         Left            =   270
         TabIndex        =   25
         Top             =   240
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   926
         _Version        =   393217
         Appearance      =   0
         TextRTF         =   $"frmMain.frx":D133
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.PictureBox picRight 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4875
      Left            =   11145
      ScaleHeight     =   4875
      ScaleWidth      =   3075
      TabIndex        =   16
      Top             =   2565
      Width           =   3075
      Begin RichTextLib.RichTextBox txt���� 
         Height          =   525
         Left            =   420
         TabIndex        =   24
         Top             =   45
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   926
         _Version        =   393217
         Appearance      =   0
         TextRTF         =   $"frmMain.frx":D1C2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.ComboBox cboϵͳ 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmMain.frx":D251
      Left            =   11430
      List            =   "frmMain.frx":D253
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   165
      Width           =   2955
   End
   Begin VB.Frame fraFind 
      Caption         =   "��ѯ����"
      Height          =   1125
      Left            =   2670
      TabIndex        =   3
      Top             =   720
      Width           =   10710
      Begin VB.ComboBox cboӰ������ 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6975
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   600
         Width           =   2055
      End
      Begin VB.CommandButton cmdReLoad 
         Caption         =   "ˢ��(&R)"
         Height          =   350
         Left            =   9495
         TabIndex        =   23
         Top             =   180
         Width           =   1100
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "����(&S)"
         Height          =   350
         Left            =   8310
         TabIndex        =   22
         Top             =   180
         Width           =   1100
      End
      Begin VB.ComboBox cbo�û� 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1050
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   240
         Width           =   2715
      End
      Begin VB.ComboBox cbo��ѵ 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4695
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   600
         Width           =   1350
      End
      Begin VB.ComboBox cbo�Ƿ��Ķ� 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2970
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   600
         Width           =   800
      End
      Begin VB.ComboBox cbo���յȼ� 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1050
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   600
         Width           =   1000
      End
      Begin VB.ComboBox cbo�����汾 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6825
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   225
         Width           =   1350
      End
      Begin VB.ComboBox cbo��ʼ�汾 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4695
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   1350
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Ӱ������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   6090
         TabIndex        =   13
         Top             =   675
         Width           =   840
      End
      Begin VB.Label lblA 
         AutoSize        =   -1  'True
         Caption         =   "#"
         Height          =   180
         Left            =   9120
         TabIndex        =   21
         Top             =   675
         Width           =   90
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "�Ǽ��û�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   19
         Top             =   300
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "�Ƿ���ѵ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3810
         TabIndex        =   11
         Top             =   675
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "�Ƿ��Ķ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2055
         TabIndex        =   9
         Top             =   675
         Width           =   840
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         Caption         =   "���յȼ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   7
         Top             =   675
         Width           =   840
      End
      Begin VB.Label lbl��ʼ�汾 
         AutoSize        =   -1  'True
         Caption         =   "�汾��Χ               ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3810
         TabIndex        =   4
         Top             =   300
         Width           =   2625
      End
   End
   Begin MSComctlLib.ImageList ImgList 
      Left            =   9615
      Top             =   4110
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D255
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D7EF
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":DD89
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E123
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E6BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":EC57
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11039
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1341B
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":157FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17BDF
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17F79
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picTree 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6435
      Left            =   105
      ScaleHeight     =   6435
      ScaleWidth      =   2565
      TabIndex        =   0
      Top             =   675
      Width           =   2565
      Begin MSComctlLib.ImageList imgTree 
         Left            =   585
         Top             =   4320
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   393216
      End
      Begin MSComctlLib.TreeView tvwLeft 
         Height          =   6090
         Left            =   60
         TabIndex        =   2
         Top             =   60
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   10742
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   88
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "imgTree"
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   7050
      Top             =   135
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin XtremeCommandBars.ImageManager imgMenu 
      Left            =   900
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmMain.frx":18313
   End
   Begin XtremeCommandBars.CommandBars cbsMenu 
      Left            =   480
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmMain.frx":1BCC3
      Left            =   1650
      Top             =   345
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrsSheet As ADODB.Recordset
Private mlngItemID As Long

Private Const Dkp_ID_Tree As Integer = 1                          '�б�
Private Const Dkp_ID_Find As Integer = 2                          '����
Private Const Dkp_ID_Rept As Integer = 3
Private Const Dkp_ID_Right As Integer = 4
Private Const Dkp_ID_˵�� As Integer = 5
Private Const Dkp_ID_���� As Integer = 6

Private mIntType As Integer '��ʾ��ʽ  0-δ��¼��ʽ 1-�ѵ�¼��ʽ

Enum �û�Ӱ��
    δ��д
    ��������
    ��������
    ��Ӱ��
End Enum

Private ItemHot As ReportRecordItem         '��ǰ��ѵ���
Private rowLink As ReportRow        '��ǰ�����ӽ�����
Private mblnEdit As Boolean                 '�Ƿ��޸Ĺ���Ŀֵ
Private mstrFileName As String
Private mLastFileName As String
Private Type T����
    �û�     As String
    ģ��     As String
    ��ʼ�汾 As String
    �����汾 As String
    ���յȼ� As String
    �Ƿ��Ķ� As String
    �Ƿ���ѵ As String
    Ӱ������ As String
End Type

Private m���� As T����
Private mstr�������� As String
Private mstr���� As String
Private mLastNode As Node

Private Sub cbo�����汾_Click()
    If cbo�����汾.List(cbo�����汾.ListIndex) < cbo��ʼ�汾.List(cbo��ʼ�汾.ListIndex) Then
        
    End If
End Sub

Private Sub cbo�����汾_Validate(Cancel As Boolean)
    Dim intIndex As Integer
    If cbo�����汾.List(cbo�����汾.ListIndex) < cbo��ʼ�汾.List(cbo��ʼ�汾.ListIndex) Then
        For intIndex = 0 To cbo�����汾.ListCount - 1
            If cbo�����汾.List(intIndex) > cbo��ʼ�汾.List(cbo��ʼ�汾.ListIndex) Then
                cbo�����汾.ListIndex = intIndex
                Exit For
            End If
        Next
    End If
End Sub

Private Sub cbsMenu_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)

    Select Case Control.Id
        Case conMenu_File_Save
            Control.Enabled = mblnEdit
            cmdSave.Enabled = Control.Enabled
        Case conMenu_View_ShowPrivewText                '��ʾ�û�����
            Control.Checked = rptList.PreviewMode
        Case conMenu_View_ShowGroupBox                  '��ʾ�����
            Control.Checked = rptList.ShowGroupBox
        Case conMenu_View_ShowRelation
            Control.Enabled = mstr�������� <> ""         '��ʾ��������
    End Select
End Sub


Private Sub cmdReLoad_Click()
    Call ReLoad
End Sub

Private Sub cmdSave_Click()
     Call SaveItem: mblnEdit = False
End Sub

Private Sub Form_Load()
    
    Call initCommbar    '��ʼ���˵�
    Call initDockPane   '��ʼ�����������
    
    Call LoadInitIcon   'װ��ϵͳͼ��

    Call initRptList(rptList, ImgList, lbl����.Font, True)     '��ʼ�������б�
    
    Call initSYS        '��ʼ����ѡϵͳ
    '����Ĭ��ϵͳ
    
        
    ' ���±���
    If mIntType = 1 Then
        Me.Caption = Me.Caption & "��" & "�Ķ��ߣ�" & gstrDBUser
        
    ElseIf mIntType = 0 Then
        Me.Caption = Me.Caption & "��(δ��¼) "
    End If
    mstr���� = Me.Caption
    
    txt˵��.Locked = True
    txt����.Locked = True
    txt��������.Locked = True
    
    mblnEdit = False

    '�ָ���������
    Dim strTmp As String
    strTmp = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & TypeName(rptList), rptList.Name, "")
    If strTmp <> "" Then rptList.LoadSettings strTmp
    If cboϵͳ.ListCount > 0 Then cboϵͳ.ListIndex = 0

End Sub

Private Sub cboϵͳ_Click()
    Dim intI As Integer
    If mblnEdit Then
        If MsgBox("�������޸�δ���棬�Ƿ������", vbYesNo + vbDefaultButton2, gstrSysname) = vbNo Then Exit Sub
        mblnEdit = False
    End If
    mLastFileName = mstrFileName
    mstrFileName = ""
    mblnEdit = False
    
    cbo���յȼ�.Clear
    cbo���յȼ�.AddItem "����"
    
    cbo���յȼ�.AddItem "��"
    cbo���յȼ�.AddItem "��"
    cbo���յȼ�.AddItem "��"
    cbo���յȼ�.AddItem "��ȷ��"
    
    cbo�Ƿ��Ķ�.Clear
    cbo�Ƿ��Ķ�.AddItem "����"
    cbo�Ƿ��Ķ�.AddItem "����"
    cbo�Ƿ��Ķ�.AddItem "δ��"
    
    cbo��ѵ.Clear
    cbo��ѵ.AddItem "����"
    cbo��ѵ.AddItem "δ��д"
    cbo��ѵ.AddItem "����ѵ"
    cbo��ѵ.AddItem "������ѵ"

    m����.���յȼ� = GetSetting("ZLSOFT", "����ģ��\UpgradeReader", "���յȼ�", "����")
    m����.�Ƿ���ѵ = GetSetting("ZLSOFT", "����ģ��\UpgradeReader", "�Ƿ���ѵ", "����")
    m����.�Ƿ��Ķ� = GetSetting("ZLSOFT", "����ģ��\UpgradeReader", "�Ƿ��Ķ�", "����")
    m����.�û� = GetSetting("ZLSOFT", "����ģ��\UpgradeReader", "�û�", "����")
    m����.ģ�� = GetSetting("ZLSOFT", "����ģ��\UpgradeReader", "ģ��", "����ģ��")
    m����.��ʼ�汾 = GetSetting("ZLSOFT", "����ģ��\UpgradeReader", "��ʼ�汾", "0.0.0")
    m����.�����汾 = GetSetting("ZLSOFT", "����ģ��\UpgradeReader", "�����汾", "100.100.100")
    m����.Ӱ������ = GetSetting("ZLSOFT", "����ģ��\UpgradeReader", "Ӱ������", "����")
    
    cbo���յȼ�.ListIndex = 0
    If m����.���յȼ� <> "����" Then
        For intI = 0 To cbo���յȼ�.ListCount - 1
            If m����.���յȼ� = cbo���յȼ�.List(intI) Then
                cbo���յȼ�.ListIndex = intI
                Exit For
            End If
        Next
    End If
    
    cbo�Ƿ��Ķ�.ListIndex = 0
    If m����.�Ƿ��Ķ� <> "����" Then
        For intI = 0 To cbo�Ƿ��Ķ�.ListCount - 1
            If m����.�Ƿ��Ķ� = cbo�Ƿ��Ķ�.List(intI) Then
                cbo�Ƿ��Ķ�.ListIndex = intI
                Exit For
            End If
        Next
    End If
    
    cbo��ѵ.ListIndex = 0
    If m����.�Ƿ���ѵ <> "����" Then
        For intI = 0 To cbo��ѵ.ListCount - 1
            If m����.�Ƿ���ѵ = cbo��ѵ.List(intI) Then
                cbo��ѵ.ListIndex = intI
                Exit For
            End If
        Next
    End If
    
    Call initTree       '��ʼ��ģ���б�
    If mLastFileName = "" Then mLastFileName = mstrFileName
    Me.Caption = mstr���� & " " & mstrFileName
End Sub

Private Sub cbsMenu_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl
    Dim objRow As ReportRow, i As Long


    Select Case Control.Id

    Case conMenu_View_Expend_CurCollapse                            '�۵���ǰ��
        If rptList.SelectedRows.Count > 0 Then
            If rptList.SelectedRows(0).GroupRow Then
                rptList.SelectedRows(0).Expanded = False
            ElseIf Not rptList.SelectedRows(0).ParentRow Is Nothing Then
                If rptList.SelectedRows(0).ParentRow.GroupRow Then
                    rptList.SelectedRows(0).ParentRow.Expanded = False
                End If
            End If
        End If
        '���۵���λ��������,�����Զ�������¼�
        Call rptList_SelectionChanged

    Case conMenu_View_Expend_CurExpend                              'չ����ǰ��
        If rptList.SelectedRows.Count > 0 Then
            rptList.SelectedRows(0).Expanded = True
        End If
    Case conMenu_View_Expend_AllCollapse                            '�۵�������
        For Each objRow In rptList.Rows
            If objRow.GroupRow Then objRow.Expanded = False
        Next
        '���۵���λ��������,�����Զ�������¼�
        Call rptList_SelectionChanged
    Case conMenu_View_Expend_AllExpend                              'չ��������
        For Each objRow In rptList.Rows
            If objRow.GroupRow Then objRow.Expanded = True
        Next
    Case conMenu_View_ShowPrivewText                                '��ʾ�û�����
        rptList.PreviewMode = Not rptList.PreviewMode
    Case conMenu_View_ShowGroupBox
        rptList.ShowGroupBox = Not rptList.ShowGroupBox             '��ʾ�����
    Case conMenu_View_ShowRelation
        Call frmRelation.ShowRelation(mstrFileName, mstr��������)   '��ʾ��������
        
    Case conMenu_File_Save                                          '����
        Call SaveItem: mblnEdit = False
    Case conMenu_File_Exit        '�˳�
        Unload Me
        
    Case conMenu_View_Refresh     'ˢ��
        Call ReLoad
    End Select
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.Id = Dkp_ID_Tree Then
        Item.Handle = picTree.hwnd
    ElseIf Item.Id = Dkp_ID_Find Then
        Item.Handle = fraFind.hwnd
    ElseIf Item.Id = Dkp_ID_Rept Then
        Item.Handle = rptList.hwnd
    ElseIf Item.Id = Dkp_ID_Right Then
        Item.Handle = picRight.hwnd
    ElseIf Item.Id = Dkp_ID_˵�� Then
        Item.Handle = pic˵��.hwnd
    ElseIf Item.Id = Dkp_ID_���� Then
        Item.Handle = pic��������.hwnd
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnEdit Then
        If MsgBox("�������޸�δ���棬�Ƿ������", vbYesNo + vbDefaultButton2, gstrSysname) = vbNo Then Exit Sub
        mblnEdit = False
    End If
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & TypeName(rptList), rptList.Name, rptList.SaveSettings
    '�رչ��������Ĵ���
    CloseWindows
    '�ر�Ӧ�ù��߰������Ĵ���
    mclsAppTool.CloseWindows
End Sub

Private Sub picRight_Resize()
    On Error Resume Next
    With Me.txt����
        .Left = picRight.ScaleLeft
        .Top = picRight.ScaleTop
        .Width = picRight.ScaleWidth - 45
        .Height = picRight.ScaleHeight - 45
    End With
End Sub

Private Sub picTree_Resize()
    On Error Resume Next
    Me.tvwLeft.Left = 0
    Me.tvwLeft.Top = 0
    Me.tvwLeft.Width = picTree.ScaleWidth
    Me.tvwLeft.Height = picTree.ScaleHeight - Me.tvwLeft.Top
End Sub

Private Sub pic��������_Resize()
    On Error Resume Next
    With Me.txt��������
        .Left = pic��������.ScaleLeft
        .Top = pic��������.ScaleTop
        .Width = pic��������.ScaleWidth - 45
        .Height = pic��������.ScaleHeight - 45
    End With
End Sub

Private Sub pic˵��_Resize()
    On Error Resume Next
    With Me.txt˵��
        .Left = pic˵��.ScaleLeft
        .Top = pic˵��.ScaleTop
        .Width = pic˵��.ScaleWidth - 45
        .Height = pic˵��.ScaleHeight - 45
    End With
End Sub

Private Sub RichTextBox1_Change()

End Sub

Private Sub rptList_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
     Dim RecordItem As ReportRecordItem
    If (Row.Record(mCol.Ӱ������).Value = δ��д) Then
        For Each RecordItem In Row.Record
            RecordItem.Bold = True
        Next
    Else
        For Each RecordItem In Row.Record
            RecordItem.Bold = False
        Next
    End If
        
    If (Item.Index = mCol.����) Then
        Select Case Item.Value
            Case 0: Item.Icon = ICON_Unknown    '��ȷ��
            Case 1: Item.Icon = ICON_Low        '��
            Case 2: Item.Icon = ICON_Center     '��
            Case 3: Item.Icon = ICON_High       '��
        End Select
    End If
    
    If (Item.Index = mCol.���) Then
        If Row.Record(mCol.����).Value = "��" Then
            Set Metrics.Font = fntUnderLine
            Metrics.ForeColor = vbBlue
        End If
    End If
End Sub

Private Sub rptList_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim strLinkFile As String
    If Button = 1 Then
'        If (Not ItemHot Is Nothing) Then
'            If ItemHot.Value = "������ѵ" Then Exit Sub '������ѵ
'             ItemHot.Value = IIf(ItemHot.Value = "����ѵ", "��", "����ѵ") '�޸���ѵ״̬
'            'If ItemHot.Icon = -1 Then Exit Sub
''            ItemHot.Icon = IIf(ItemHot.Icon = ICON_Train, 6, 5)
''            ItemHot.Value = IIf(ItemHot.Icon = ICON_Train, 6, 5)
'            mblnEdit = True
'        End If
        
        If (Not rowLink Is Nothing) Then
            If rowLink.Record(mCol.����).Value = "��" Then
                strLinkFile = Mid(mstrFileName, 1, InStrRev(mstrFileName, "\")) & "Document\" & rowLink.Record(mCol.���).Value & ".htm"
                If Dir(strLinkFile) <> "" Then
                    Call ShellExecute(Me.hwnd, "open", "file:///" & Replace(strLinkFile, "\", "/"), vbNullString, vbNullString, 1)
                Else
                    MsgBox "δ�ҵ���Ӧ��html�ļ������ļ�ʧ�ܣ�", vbInformation, gstrSysname
                End If
            End If
        End If
    End If
End Sub

Private Sub rptList_MouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
    
    Dim htInfo As ReportHitTestInfo
    Set htInfo = rptList.HitTest(X, Y)
    
    Dim Item As ReportRecordItem
    Dim objRow As ReportRow

    If (Not htInfo.Item Is Nothing) Then
        If (htInfo.Item.Index = mCol.��ѵ) Then
            Set Item = htInfo.Item
        End If
        
        If (htInfo.Item.Index = mCol.���) Then
            Set objRow = htInfo.Row
        End If
    End If

    If (Not objRow Is rowLink) Then
        If (Not objRow Is Nothing) Then
            If objRow.Record(mCol.����).Value = "��" Then
                objRow.Record(mCol.���).BackColor = RGB(255, 238, 99)
            End If
            
        End If
        
        If (Not rowLink Is Nothing) Then
            rowLink.Record(mCol.���).BackColor = -1
        End If
        
        Set rowLink = objRow
        rptList.Redraw
    End If
    
    If (Not Item Is ItemHot) Then
        If (Not Item Is Nothing) Then
            If Item.Value = "������ѵ" Then Exit Sub '������ѵ
'            If Item.Icon = -1 Then Exit Sub
'            Item.BackColor = IIf(Item.Icon = ICON_Train, RGB(207, 93, 96), RGB(255, 238, 194))
            Item.BackColor = IIf(Item.Value = "��", RGB(207, 93, 96), RGB(255, 238, 194))
        End If

        If (Not ItemHot Is Nothing) Then
            If ItemHot.Value = "" Then Exit Sub
            ItemHot.BackColor = -1
        End If
        Set ItemHot = Item
        rptList.Redraw
    End If
    

End Sub

Private Sub rptList_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim objPopup As CommandBarPopup
    If Button = 2 Then
        Set objPopup = cbsMenu.ActiveMenuBar.FindControl(, conMenu_View)
        If Not objPopup Is Nothing Then
            objPopup.CommandBar.ShowPopup
        End If
    End If
End Sub

Private Sub rptList_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
        
        If (Not ItemHot Is Nothing) Then
            If ItemHot.Value = "������ѵ" Then Exit Sub '������ѵ
             ItemHot.Value = IIf(ItemHot.Value = "����ѵ", "��", "����ѵ") '�޸���ѵ״̬
            'If ItemHot.Icon = -1 Then Exit Sub
'            ItemHot.Icon = IIf(ItemHot.Icon = ICON_Train, 6, 5)
'            ItemHot.Value = IIf(ItemHot.Icon = ICON_Train, 6, 5)
            mblnEdit = True
            rptList.Redraw
        End If
End Sub

Private Sub rptList_SelectionChanged()
    '#
    txt˵�� = ""
    txt���� = ""
    txt�������� = ""
    mstr�������� = ""
    If rptList.FocusedRow Is Nothing Then Exit Sub
    If Not rptList.FocusedRow.GroupRow Then
        txt˵�� = rptList.FocusedRow.Record(mCol.˵��).Value
        txt���� = rptList.FocusedRow.Record(mCol.����).Value
        txt�������� = rptList.FocusedRow.Record(mCol.��������).Value
        mstr�������� = Trim(rptList.FocusedRow.Record(mCol.��������).Value)
    End If
    
End Sub

Private Sub rptList_ValueChanged(ByVal Row As XtremeReportControl.IReportRow, ByVal Column As XtremeReportControl.IReportColumn, ByVal Item As XtremeReportControl.IReportRecordItem)
    
    If (Item.Index = mCol.Ӱ������) Then
        Dim ItemRead As ReportRecordItem
        Set ItemRead = Item.Record(mCol.�Ķ�)
        
        If (Item.Value = δ��д) Then
            ItemRead.Icon = ICON_NoRead
        Else
            ItemRead.Icon = ICON_Read
        End If
        Item.Record(mCol.�޸�).Value = "1"
        mblnEdit = True
    
    End If

End Sub

Private Sub tvwLeft_NodeClick(ByVal Node As MSComctlLib.Node)
    '����
    If mstrFileName <> "" Then
        If mblnEdit Then
            If MsgBox("�������޸�δ���棬�Ƿ������", vbYesNo + vbDefaultButton2, gstrSysname) = vbNo Then
                If Not mLastNode Is Nothing Then tvwLeft.SelectedItem = mLastNode
                Exit Sub
            End If
        End If
        Call LoadSheet(mstrFileName)
        Set mLastNode = Node
    End If
    
End Sub

'-----------------�����Ǳ������Զ������
Public Sub Show_me(ByVal intType As Integer)
    mIntType = intType
    Me.Show
    
End Sub

Private Sub OpenExcel(ByVal strϵͳ As String)
    Dim strSheet As String
    Dim strFilename As String
    Dim strPath As String
   
    If mstrFileName = "" Then
        
        strPath = App.Path & "\"
        strPath = Mid(strPath, 1, InStrRev(strPath, "\"))
        strPath = strPath & ReadFromIni(App.Path & "\" & App.EXEName & ".ini", strϵͳ, "Path")
        strFilename = Dir(strPath & "\*.xls")
        
        If strFilename = "" Then
            strPath = GetSetting("ZLSOFT", "����ȫ��", "����·��", App.Path & "\")
            strPath = Mid(strPath, 1, InStrRev(strPath, "\"))
            strPath = strPath & ReadFromIni(App.Path & "\" & App.EXEName & ".ini", strϵͳ, "Path")
            strFilename = Dir(strPath & "\*.xls")
        End If
        
        If strFilename = "" Then
            If MsgBox("��Ĭ��Ŀ¼δ�ҵ�EXCEL�ļ������ֹ�ָ������˵���ļ���", vbYesNo + vbDefaultButton1, gstrSysname) = vbYes Then
                dlgFile.DialogTitle = "��һ������˵���ļ��������"
                dlgFile.InitDir = App.Path
                dlgFile.Filename = ""
                dlgFile.Filter = "����˵���ļ�|*.xls"
                dlgFile.ShowOpen
                If dlgFile.Filename = "" Then Exit Sub
                strFilename = dlgFile.Filename
            End If
        Else
            strFilename = strPath & "\" & strFilename
        End If
        mstrFileName = strFilename
    End If
    
    
    
End Sub

Private Function initSYS() As Boolean
    '���ϵͳ
    Dim strIniFileName As String, i As Integer, strRemoveItem As String
    Dim strϵͳ As String, varTmp As Variant, str�Ѱ�װϵͳ As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    strϵͳ = Trim(ReadFromIni(App.Path & "\" & App.EXEName & ".ini", "ϵͳ", "����"))
    If strϵͳ = "" Then
        MsgBox "�����ļ���ʧ�����ܼ������У�", vbQuestion, gstrSysname
        Exit Function
    End If
    
    cboϵͳ.Clear
    
    If mIntType = 0 Then
        'δ��¼��ʽ����ʾ�����ļ��е�ϵͳ
        If InStr(strϵͳ, "|") > 0 Then
            varTmp = Split(strϵͳ, "|")
            For i = LBound(varTmp) To UBound(varTmp)
                If Trim(varTmp(i)) <> "" Then
                    cboϵͳ.AddItem varTmp(i)
                End If
            Next
        Else
            cboϵͳ.AddItem strϵͳ
        End If
    ElseIf mIntType = 1 Then
        '������ѵ�¼�ģ���ʾ��Ȩ���ʵ�ϵͳ
        str�Ѱ�װϵͳ = ""
        Set rsTmp = gcnOracle.Execute("Select ����,��� From zlsystems Where ��� In(" & gstrSystems & ") Order by ���")
        Do Until rsTmp.EOF
            str�Ѱ�װϵͳ = Trim("" & rsTmp!����)
            If str�Ѱ�װϵͳ <> "" Then
                cboϵͳ.AddItem str�Ѱ�װϵͳ
            End If
            rsTmp.MoveNext
        Loop
    End If
    
    
    initSYS = True
    Exit Function
errHandle:
    initSYS = False
    MsgBox Err.Number & " " & Err.Description, vbInformation, gstrSysname
End Function


Private Sub initCommbar()
    '��ʼ���˵�����������
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom
    Dim objPopup As CommandBarPopup
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMenu.VisualTheme = xtpThemeOffice2003
    With Me.cbsMenu.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '����VisualTheme����Ч
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = False
        .SetIconSize False, 16, 16
    End With
    cbsMenu.EnableCustomization False
    Set cbsMenu.Icons = imgMenu.Icons
    
    '�˵�����:������������
    '    ���xtpControlPopup���͵�����ID���¸�ֵ
    '-----------------------------------------------------
    cbsMenu.ActiveMenuBar.Title = "�˵�"
    cbsMenu.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    
    Set objMenu = cbsMenu.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_File, "�ļ�(&F)", -1, False)
    objMenu.Id = conMenu_File
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Save, "����(&S)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)")
    End With


    
    Set objMenu = cbsMenu.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_View, "�鿴(&V)", -1, False)
    objMenu.Id = conMenu_View
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_Expend, "չ��/�۵���(&X)")
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_AllCollapse, "�۵�������(&L)", -1, False)
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_AllExpend, "չ��������(&X)", -1, False)
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_CurCollapse, "�۵���ǰ��(&C)", -1, False): objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_CurExpend, "չ����ǰ��(&E)", -1, False)
        End With
        Set objControl = .Add(xtpControlButton, conMenu_View_ShowGroupBox, "��ʾ�����(&S)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_ShowPrivewText, "Ԥ���û�����(&P)")
        Set objControl = .Add(xtpControlButton, conMenu_View_ShowRelation, "�鿴��������(&R)")
        
'        Set objControl = .Add(xtpControlButton, conMenu_View_Find, "����(&F)"): objControl.BeginGroup = True
'        Set objControl = .Add(xtpControlButton, conMenu_View_FindNext, "����һ�¸�(&N)")
'        Set objControl = .Add(xtpControlButton, conMenu_View_Filter, "ɸѡ(&I)"): objControl.BeginGroup = True
'        Set objControl = .Add(xtpControlButton, conMenu_View_RecordPrev, "ǰһ����¼(&P)")
'        Set objControl = .Add(xtpControlButton, conMenu_View_RecordNext, "��һ����¼(&D)")
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)"): objControl.BeginGroup = True

    End With
    

    
'    Set objMenu = cbsMenu.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_Help, "����(&H)", -1, False)
'    objMenu.Id = conMenu_Help
'    With objMenu.CommandBar.Controls
'        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)..."): objControl.BeginGroup = True
'    End With

    '����������:������������
    '-----------------------------------------------------
    Set objBar = cbsMenu.Add("������", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Save, "����"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_ShowGroupBox, "�����"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_ShowPrivewText, "Ԥ��")
        Set objControl = .Add(xtpControlButton, conMenu_View_ShowRelation, "����")
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��"): objControl.BeginGroup = True
        
'        Set objControl = .Add(xtpControlButton, conMenu_View_Find, "����"): objControl.BeginGroup = True
'        Set objControl = .Add(xtpControlButton, conMenu_View_Filter, "ɸѡ"): objControl.BeginGroup = True
'        Set objControl = .Add(xtpControlButton, conMenu_View_RecordPrev, "ǰһ��")
'        Set objControl = .Add(xtpControlButton, conMenu_View_RecordNext, "��һ��")
        
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�"): objControl.BeginGroup = True
        
        Set objControl = .Add(xtpControlLabel, conMenu_Custom_System - 1, "ϵͳ")
        objControl.Flags = xtpFlagRightAlign
    
        Set objCustom = .Add(xtpControlCustom, conMenu_Custom_System, "ϵͳ")
        objCustom.ShortcutText = "ϵͳ"
        objCustom.Handle = Me.cboϵͳ.hwnd
        objCustom.Flags = xtpFlagRightAlign
        objCustom.Style = xtpButtonIconAndCaption
    End With
    
    'ǰһ������һ����¼����ʾ����˵��
    For Each objControl In objBar.Controls
        If objControl.Id <> conMenu_View_RecordPrev And objControl.Id <> conMenu_View_RecordNext And objControl.Id <> conMenu_Custom_System And objControl.Id <> conMenu_Custom_System - 1 Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    '����Ŀ����:���������������Ѵ���
    '-----------------------------------------------------
    With cbsMenu.KeyBindings
        .Add FCONTROL, vbKeyS, conMenu_File_Save
'        .Add FCONTROL, vbKeyF, conMenu_View_Find
'        .Add 0, vbKeyF3, conMenu_View_FindNext
'        .Add FCONTROL, vbKeyI, conMenu_View_Filter
'        .Add FCONTROL, vbKeyLeft, conMenu_View_RecordPrev
'        .Add FCONTROL, vbKeyRight, conMenu_View_RecordNext
        .Add 0, vbKeyF5, conMenu_View_Refresh
    End With
    
    'MDI Tab
'    '-----------------------------------------------------
'    cbsMenu.ActiveMenuBar.SetFlags xtpFlagHideMDIButtons, 0
'    Set mWorkSpace = cbsMenu.ShowTabWorkspace(True)
'    cbsMenu.TabWorkspace.AutoTheme = False
'    cbsMenu.TabWorkspace.PaintManager.Appearance = xtpTabAppearanceVisualStudio
'    cbsMenu.TabWorkspace.PaintManager.Color = xtpTabColorOffice2003
'    cbsMenu.TabWorkspace.PaintManager.ClientFrame = xtpTabFrameSingleLine
    
    '״̬��
    '-----------------------------------------------------
'    cbsMenu.StatusBar.Visible = True
'    cbsMenu.StatusBar.AddPane 1
'    cbsMenu.StatusBar.SetPaneStyle 1, SBPS_STRETCH
'    cbsMenu.StatusBar.SetPaneText 1, ""
'    cbsMenu.StatusBar.AddPane 2
'    cbsMenu.StatusBar.SetPaneWidth 2, 100
'    cbsMenu.StatusBar.SetPaneText 2, ""
'    cbsMenu.StatusBar.AddPane 3
'    cbsMenu.StatusBar.SetPaneWidth 3, 60
'    cbsMenu.StatusBar.SetPaneText 3, ""
'    cbsMenu.StatusBar.IdleText = ""
    
    picRight.BackColor = cbsMenu.GetSpecialColor(STDCOLOR_BTNFACE)
    picTree.BackColor = cbsMenu.GetSpecialColor(STDCOLOR_BTNFACE)
    pic˵��.BackColor = cbsMenu.GetSpecialColor(STDCOLOR_BTNFACE)
    pic��������.BackColor = cbsMenu.GetSpecialColor(STDCOLOR_BTNFACE)
End Sub

Private Sub initDockPane()
    Dim paneTree As Pane, paneFind As Pane, paneEdit As Pane, paneList As Pane, paneRight As Pane, pane˵�� As Pane, pane���� As Pane
    
    Me.dkpMain.SetCommandBars Me.cbsMenu
    Me.dkpMain.Options.UseSplitterTracker = False 'ʵʱ�϶�
    Me.dkpMain.Options.ThemedFloatingFrames = True
    Me.dkpMain.Options.AlphaDockingContext = True
    
    Me.dkpMain.Options.HideClient = True
    
    Set paneList = Me.dkpMain.CreatePane(Dkp_ID_Rept, 900, 700, DockTopOf, Nothing)
    paneList.Title = "�����嵥"
    paneList.Handle = Me.rptList.hwnd
    paneList.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set paneTree = Me.dkpMain.CreatePane(Dkp_ID_Tree, 180, 90, DockLeftOf, Nothing)
    paneTree.Title = "ģ��"
    paneTree.Handle = Me.picTree.hwnd
    
    Set paneFind = Me.dkpMain.CreatePane(Dkp_ID_Find, 50, 180, DockTopOf, paneList)
    paneFind.Title = "��ѯ����"
    paneFind.Handle = Me.fraFind.hwnd
    paneFind.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set pane˵�� = Me.dkpMain.CreatePane(Dkp_ID_˵��, 800, 500, DockBottomOf, paneList)
    pane˵��.Title = "�޸�˵��"
    pane˵��.Handle = Me.pic˵��.hwnd
    pane˵��.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set paneRight = Me.dkpMain.CreatePane(Dkp_ID_Right, 100, 300, DockBottomOf, pane˵��)
    paneRight.Title = "�û�����"
    paneRight.Handle = Me.picRight.hwnd
    paneRight.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set pane���� = Me.dkpMain.CreatePane(Dkp_ID_����, 100, 300, DockRightOf, paneRight)
    pane����.Title = "��������"
    pane����.Handle = Me.pic��������.hwnd
    pane����.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
End Sub

Private Sub initTree()
    '
    Dim objNode As Node, intGrant As Integer
    Dim objRootNode As Node
    Dim strSheet As String, varSheet As Variant, i As Integer, intList As Integer
    Dim strNodeText As String, strģ�� As String
    Dim blnAdd As Boolean
    Dim str��Ͱ汾 As String, str��߰汾 As String, strTmp�汾 As String, str�û� As String, strӰ��ģ�� As String, varӰ��ģ�� As Variant, strӰ������ As String
    Dim str�û��汾 As String, rsTmp As ADODB.Recordset, StrSQL As String
    
    On Error Resume Next
    
    If cboϵͳ.ListIndex < 0 Then Exit Sub
    Call OpenExcel(cboϵͳ.List(cboϵͳ.ListIndex))
    If mstrFileName = "" Then
        mstrFileName = mLastFileName
        Exit Sub
    End If
    
    Me.tvwLeft.Nodes.Clear
    Set objRootNode = Me.tvwLeft.Nodes.Add(, , "Root", "����ģ��", "K_" & 141)
    
    cbo��ʼ�汾.Clear
    'cbo��ʼ�汾.AddItem "0.0.0"
    cbo�����汾.Clear
    'cbo�����汾.AddItem "100.100.100"
    cbo�û�.Clear
    cbo�û�.AddItem "����"
    If mIntType = 1 And gstr�û���λ���� <> "" Then
        cbo�û�.AddItem gstr�û���λ����
        cbo�û�.ListIndex = cbo�û�.ListCount - 1
    End If
    
    If mstrFileName <> "" Then
        strSheet = OpenExcelFile(mstrFileName)
        If InStr(strSheet, "|") <= 0 Then Exit Sub
    End If
    
    cboӰ������.Clear
    cboӰ������.AddItem "����"
    cboӰ������.ListIndex = 0
    
    '----
    varSheet = Split(strSheet, "|")
    For i = LBound(varSheet) To UBound(varSheet)
        Set mrsSheet = OpenExcelSheet(varSheet(i))
        Do Until mrsSheet.EOF
            'ȡ�汾
            str�û� = Trim("" & mrsSheet(Excel_Col.�Ǽ��û�).Value)
'            If mIntType = 0 Or (mIntType = 1 And str�û� = gstr�û���λ����) Then
                strTmp�汾 = "" & mrsSheet(Excel_Col.�����汾).Value
                If strTmp�汾 <> "" Then
                    If str��Ͱ汾 = "" Then
                        cbo��ʼ�汾.AddItem strTmp�汾
                        str��Ͱ汾 = strTmp�汾
                    Else
                        If str��Ͱ汾 > strTmp�汾 Then
                            str��Ͱ汾 = strTmp�汾
                        End If
                        
                        blnAdd = True
                        For intList = 0 To cbo��ʼ�汾.ListCount - 1
                            If cbo��ʼ�汾.List(intList) = strTmp�汾 Then
                                blnAdd = False
                            End If
                        Next
                        
                        If blnAdd Then
                            cbo��ʼ�汾.AddItem strTmp�汾
                        End If
                    
                    End If
                
                    If str��߰汾 = "" Then
                        cbo�����汾.AddItem strTmp�汾
                        str��߰汾 = strTmp�汾
                    Else
                        If str��߰汾 < strTmp�汾 Then
                            str��߰汾 = strTmp�汾
                        End If
                        blnAdd = True
                        For intList = 0 To cbo�����汾.ListCount - 1
                            If cbo�����汾.List(intList) = strTmp�汾 Then blnAdd = False
                        Next
                        If blnAdd Then
                            cbo�����汾.AddItem strTmp�汾
                        End If
                    End If
                End If 'end If strTmp�汾 <> ""
'            End If
            
            'ȡӰ������
            strӰ������ = Trim("" & mrsSheet(Excel_Col.Ӱ������).Value)
            If strӰ������ <> "" Then
                blnAdd = True
                For intList = 0 To cboӰ������.ListCount - 1
                    If cboӰ������.List(intList) = strӰ������ Then blnAdd = False
                Next
                If blnAdd Then cboӰ������.AddItem strӰ������
            End If
            
            If mIntType = 0 Then
                'ȡ�û���
                str�û� = Trim("" & mrsSheet(Excel_Col.�Ǽ��û�).Value)
                If str�û� <> "" Then
                    blnAdd = True
                    For intList = 0 To cbo�û�.ListCount - 1
                        If cbo�û�.List(intList) = str�û� Then blnAdd = False
                    Next
                    If blnAdd Then cbo�û�.AddItem str�û�
                End If
                

            End If
            mrsSheet.MoveNext
        Loop
    Next
    '����̶�ģ��
    Dim str�̶�ģ�� As String, var�̶�ģ�� As Variant, int�̶� As Integer
    str�̶�ģ�� = Trim(ReadFromIni(App.Path & "\" & App.EXEName & ".ini", "�̶�ģ��", "ģ��"))
    var�̶�ģ�� = Split(str�̶�ģ��, "|")
    For int�̶� = LBound(var�̶�ģ��) To UBound(var�̶�ģ��)
        If var�̶�ģ��(int�̶�) <> "" Then
            Set objNode = Me.tvwLeft.Nodes.Add(, , "G" & Format(int�̶� + 1, "000"), var�̶�ģ��(int�̶�), "K_" & 99)
            Call addModleToTree(var�̶�ģ��(int�̶�), "G" & Format(int�̶� + 1, "000"))
        End If
    Next
    
    If mIntType = 0 Then
        'δ��¼ ��ģ��
        varSheet = Split(strSheet, "|")
        For i = LBound(varSheet) To UBound(varSheet)
            Set mrsSheet = OpenExcelSheet(varSheet(i))
            Do Until mrsSheet.EOF
                Call AddExcelToTree("")
                mrsSheet.MoveNext
            Loop
        Next

    '-----
    ElseIf mIntType = 1 Then
'        If gblnOwner Then
'            '��ϵͳ������
'            Set objNode = Me.tvwLeft.Nodes.Add(, , "G01", "������", "K_" & 207)
'            Set objNode = Me.tvwLeft.Nodes.Add("G01", 4, "G0101", "װж����", "K_" & 99)
'            Set objNode = Me.tvwLeft.Nodes.Add("G0101", 4, "G010101", "ϵͳװж����", "K_" & 99)
'            Set objNode = Me.tvwLeft.Nodes.Add("G0101", 4, "G010102", "ϵͳ��Ǩ����", "K_" & 99)
'            Set objNode = Me.tvwLeft.Nodes.Add("G0101", 4, "G010103", "�������޸�", "K_" & 99)
'            Set objNode = Me.tvwLeft.Nodes.Add("G0101", 4, "G010104", "�û���װ�ű�", "K_" & 99)
'            Set objNode = Me.tvwLeft.Nodes.Add("G0101", 4, "G010105", "������Ч����", "K_" & 99)
'
'            Set objNode = Me.tvwLeft.Nodes.Add("G01", 4, "G0102", "���ݹ���", "K_" & 99)
'            Set objNode = Me.tvwLeft.Nodes.Add("G0102", 4, "G010201", "����ת��", "K_" & 99)
'            Set objNode = Me.tvwLeft.Nodes.Add("G0102", 4, "G010202", "���ݵ���", "K_" & 99)
'            Set objNode = Me.tvwLeft.Nodes.Add("G0102", 4, "G010203", "���ݵ���", "K_" & 99)
'            Set objNode = Me.tvwLeft.Nodes.Add("G0102", 4, "G010204", "���ݵ���", "K_" & 99)
'            Set objNode = Me.tvwLeft.Nodes.Add("G0102", 4, "G010205", "���ݵ���", "K_" & 99)
'            Set objNode = Me.tvwLeft.Nodes.Add("G0102", 4, "G010206", "�������", "K_" & 99)
'
'            Set objNode = Me.tvwLeft.Nodes.Add("G01", 4, "G0103", "���й���", "K_" & 99)
'            Set objNode = Me.tvwLeft.Nodes.Add("G0103", 4, "G010301", "�û�ע�����", "K_" & 99)
'            Set objNode = Me.tvwLeft.Nodes.Add("G0103", 4, "G010302", "����״̬���", "K_" & 99)
'            Set objNode = Me.tvwLeft.Nodes.Add("G0103", 4, "G010303", "��̨��ҵ����", "K_" & 99)
'            Set objNode = Me.tvwLeft.Nodes.Add("G0103", 4, "G010304", "������־����", "K_" & 99)
'            Set objNode = Me.tvwLeft.Nodes.Add("G0103", 4, "G010305", "������־����", "K_" & 99)
'            Set objNode = Me.tvwLeft.Nodes.Add("G0103", 4, "G010306", "ϵͳ����ѡ��", "K_" & 99)
'            Set objNode = Me.tvwLeft.Nodes.Add("G0103", 4, "G010307", "վ�㲿������", "K_" & 99)
'            Set objNode = Me.tvwLeft.Nodes.Add("G0103", 4, "G010308", "վ�����п���", "K_" & 99)
'            Set objNode = Me.tvwLeft.Nodes.Add("G0103", 4, "G010309", "վ���ļ��ռ�", "K_" & 99)
'
'            Set objNode = Me.tvwLeft.Nodes.Add("G01", 4, "G0104", "Ȩ�޹���", "K_" & 99)
'            Set objNode = Me.tvwLeft.Nodes.Add("G0104", 4, "G010401", "��ɫ��Ȩ����", "K_" & 99)
'            Set objNode = Me.tvwLeft.Nodes.Add("G0104", 4, "G010402", "�û���Ȩ����", "K_" & 99)
'            Set objNode = Me.tvwLeft.Nodes.Add("G0104", 4, "G010403", "�˵�����滮", "K_" & 99)
'
'            Set objNode = Me.tvwLeft.Nodes.Add("G01", 4, "G0105", "ר���", "K_" & 99)
'            Set objNode = Me.tvwLeft.Nodes.Add("G0105", 4, "G010501", "�������", "K_" & 99)
'            Set objNode = Me.tvwLeft.Nodes.Add("G0105", 4, "G010502", "��������", "K_" & 99)
'        End If
        
'        Set objNode = Me.tvwLeft.Nodes.Add(, , "T" & �����嵥.���������嵥, "���������嵥", "K_" & 210)
'
'        intGrant = zlRegTool
'        If ((intGrant And 4) = 4) Then
'            If InStr(1, GetPrivFunc(0, �����嵥.��Ϣ�շ�����), "����") <> 0 Then
'                Set objNode = Me.tvwLeft.Nodes.Add(, , "T" & �����嵥.��Ϣ�շ�����, "��Ϣ�շ�����", "K_" & 145)
'            End If
'        End If
'        If ((intGrant And 8) = 8) Then
'            If InStr(1, GetPrivFunc(0, �����嵥.EXCEL������), "����") Then
'                Set objNode = Me.tvwLeft.Nodes.Add(, , "T" & �����嵥.EXCEL������, "EXCEL������", "K_" & 217)
'            End If
'        End If
'
'        If InStr(1, GetPrivFunc(0, �����嵥.���ز�������), "����") Then
'            Set objNode = Me.tvwLeft.Nodes.Add(, , "T" & �����嵥.���ز�������, "���ز�������", "K_" & 135)
'        End If
'
'        If InStr(1, GetPrivFunc(0, �����嵥.ϵͳѡ������), "����") Then
'            Set objNode = Me.tvwLeft.Nodes.Add(, , "T" & �����嵥.ϵͳѡ������, "ϵͳѡ������", "K_" & 147)
'        End If
'
'        If InStr(1, GetPrivFunc(0, �����嵥.�ֵ������), "����") Then
'            Set objNode = Me.tvwLeft.Nodes.Add(, , "T" & �����嵥.�ֵ������, "�ֵ������", "K_" & 144)
'        End If
        
        '�ѵ�¼
        If cboϵͳ.List(cboϵͳ.ListIndex) = "������ϵͳ" Then
            Set rsTmp = rsMenuPEIS
        Else
            Set rsTmp = rsMenu
        End If
        With rsTmp
            If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
            Do While Not .EOF
                'On Error Resume Next
                If UCase(.Fields("����").Value) <> UCase("zl9Report") Then
                    If .Fields("ģ��").Value = 0 Then
                        If .Fields("�ϼ�") = 0 Then
                            Set objNode = Me.tvwLeft.Nodes.Add(, , "_" & .Fields("���").Value, .Fields("����").Value, "K_" & IIf(!ͼ�� = 0, 99, !ͼ��))
                        Else
                            Set objNode = Me.tvwLeft.Nodes.Add("_" & .Fields("�ϼ�").Value, 4, "_" & .Fields("���").Value, .Fields("����").Value, "K_" & IIf(!ͼ�� = 0, 99, !ͼ��))
                        End If
                    Else
                        '�ӹ���
                        If .Fields("�ϼ�") = 0 Then
                            Set objNode = Me.tvwLeft.Nodes.Add(, , "_" & .Fields("���").Value, .Fields("����").Value, "K_" & IIf(!ͼ�� = 0, 99, !ͼ��))
                        Else
                            Set objNode = Me.tvwLeft.Nodes.Add("_" & .Fields("�ϼ�").Value, 4, "_" & .Fields("���").Value, .Fields("����").Value, "K_" & IIf(!ͼ�� = 0, 99, !ͼ��))
                        End If
                    End If
                End If
                .MoveNext
            Loop
        End With
        
        '�ѵ�¼ δ�ҵ���ģ��ӵ��������С�
        Set objNode = Me.tvwLeft.Nodes.Add(, , "Q01", "����", "K_" & 172)
        varSheet = Split(strSheet, "|")
        For i = LBound(varSheet) To UBound(varSheet)
            Set mrsSheet = OpenExcelSheet(varSheet(i))
            Do Until mrsSheet.EOF
                Call AddExcelToTree("Q01")
                mrsSheet.MoveNext
            Loop
        Next
        
        
        'ȡ�û��汾
        StrSQL = "Select �汾�� From zlsystems Where ����='" & cboϵͳ.List(cboϵͳ.ListIndex) & "'"
        Set rsTmp = gcnOracle.Execute(StrSQL)
        Do Until rsTmp.EOF
            str�û��汾 = "" & rsTmp("�汾��").Value
            rsTmp.MoveNext
        Loop
    End If
    
    '--��ʼ�汾
    If cbo��ʼ�汾.ListCount > 0 Then
        For intList = 0 To cbo��ʼ�汾.ListCount - 1
            If cbo��ʼ�汾.List(intList) = str��Ͱ汾 Then
               cbo��ʼ�汾.ListIndex = intList: Exit For
            End If
        Next
    End If
    For intList = 0 To cbo��ʼ�汾.ListCount - 1
        If cbo��ʼ�汾.List(intList) = m����.��ʼ�汾 Then
             cbo��ʼ�汾.ListIndex = intList: Exit For
        End If
    Next
    If str�û��汾 <> "" Then
        If str�û��汾 < str��߰汾 Then
            For intList = 0 To cbo��ʼ�汾.ListCount - 1
                If cbo��ʼ�汾.List(intList) = str�û��汾 Then
                     cbo��ʼ�汾.ListIndex = intList: Exit For
                End If
            Next
        End If
    End If
    '--�����汾
    If cbo�����汾.ListCount > 0 Then
        For intList = 0 To cbo�����汾.ListCount - 1
            If cbo�����汾.List(intList) = str��߰汾 Then
                cbo�����汾.ListIndex = intList: Exit For
            End If
        Next
    End If
    For intList = 0 To cbo�����汾.ListCount - 1
        If cbo�����汾.List(intList) = m����.�����汾 Then
            cbo�����汾.ListIndex = intList: Exit For
        End If
    Next
    
    '--�û�
    If cbo�û�.ListCount > 0 Then cbo�û�.ListIndex = 0
    For intList = 0 To cbo�û�.ListCount - 1
        If cbo�û�.List(intList) = m����.�û� Then
           cbo�û�.ListIndex = intList: Exit For
        End If
    Next
    
    cboӰ������.ListIndex = 0
    If m����.Ӱ������ <> "����" Then
        For intList = 0 To cboӰ������.ListCount - 1
            If m����.Ӱ������ = cboӰ������.List(intList) Then
                cboӰ������.ListIndex = intList
                Exit For
            End If
        Next
    End If
    '����ģ��
    For intList = 1 To tvwLeft.Nodes.Count
        If tvwLeft.Nodes(intList).Text = m����.ģ�� Then
            'Set tvwLeft.SelectedItem = tvwLeft.Nodes(intList)
            Set objRootNode = tvwLeft.Nodes(intList)
            
            Exit For
        End If
    Next
        
    Set tvwLeft.SelectedItem = objRootNode
    Call tvwLeft_NodeClick(tvwLeft.SelectedItem)
    Call ReLoad
End Sub

Private Sub LoadSheet(ByVal Filename As String)

    Dim rptRcd As ReportRecord
    Dim rptItem As ReportRecordItem
    Dim rptItem1 As ReportRecordItem
    Dim rptItem2 As ReportRecordItem
    Dim rptRow As ReportRow
    Dim rptColum As ReportColumn
    
    Dim strSheet As String, varSheet As Variant, i As Integer, intNode As Integer, strģ�� As String
    Dim lngCount As Long
    Dim strTxt As String
    On Error GoTo errHandle
    If Filename = "" Then Exit Sub
    strSheet = OpenExcelFile(Filename)
    
    If strSheet = "" Then Exit Sub
    
    '����б��е�ģ�����
    
    For i = 1 To tvwLeft.Nodes.Count
        tvwLeft.Nodes(i).Tag = 0
        strTxt = Replace(tvwLeft.Nodes(i).Text, "(�°�)", "")
        If InStr(strTxt, "(") > 0 Then strTxt = Mid(strTxt, 1, InStr(strTxt, "(") - 1)
        tvwLeft.Nodes(i).Text = strTxt
    Next
    
    rptList.Records.DeleteAll '���ԭ�б�
    lblA.Caption = ""
    If InStr(strSheet, "|") <= 0 Then Exit Sub
    
    varSheet = Split(strSheet, "|")
    For i = LBound(varSheet) To UBound(varSheet)
        Set mrsSheet = OpenExcelSheet(varSheet(i))
        lngCount = 0
        Do Until mrsSheet.EOF
            
            '������ϸ
            With rptList
                    '���ģ�������
                If IsAdd(False) Then
                    strģ�� = Replace("" & mrsSheet(Excel_Col.�Ǽ�ģ��).Value, "(�°�)", "")
                    For intNode = 1 To tvwLeft.Nodes.Count
                        strTxt = tvwLeft.Nodes(intNode).Text
                        If InStr(strTxt, "(") > 0 Then strTxt = Mid(strTxt, 1, InStr(strTxt, "(") - 1)
                        If strģ�� = strTxt Then
                            tvwLeft.Nodes(intNode).Tag = Val(tvwLeft.Nodes(intNode).Tag) + 1
                            tvwLeft.Nodes(intNode).Text = strTxt & "(" & Val(tvwLeft.Nodes(intNode).Tag) & ")"
                            Call AddParent(intNode)
                        End If
                    Next
                End If
                
                If IsAdd(True) Then      '�������������

                    lngCount = lngCount + 1
                    Set rptRcd = Me.rptList.Records.Add()
                    
                    '�Ѷ� = 0: ����: ��ѵ: �汾: ����: ���: ģ��: Ӱ��ģ��: ��������: �û�: ����: ˵��: ��������: ��ע: Ӱ������: ����
                    Set rptItem = rptRcd.AddItem(""): rptItem.Focusable = False
                    If Val("" & mrsSheet(Excel_Col.���û�Ӱ������).Value) = 0 Then
                        rptItem.Icon = ICON_NoRead
                    Else
                        rptItem.Icon = ICON_Read
                    End If
                        
                    If "" & mrsSheet(Excel_Col.�������).Value = "��" Then
                        Set rptItem1 = rptRcd.AddItem(3)
                    ElseIf "" & mrsSheet(Excel_Col.�������).Value = "��" Then
                        Set rptItem1 = rptRcd.AddItem(2)
                    ElseIf "" & mrsSheet(Excel_Col.�������).Value = "��" Then
                        Set rptItem1 = rptRcd.AddItem(1)
                    Else
                        Set rptItem1 = rptRcd.AddItem(0)
                    End If
                    rptItem1.Caption = " ": rptItem1.Focusable = False

                    Set rptItem = rptRcd.AddItem(CStr("" & mrsSheet(Excel_Col.�����汾).Value)): rptItem.Focusable = False
                    
                    Set rptItem = rptRcd.AddItem(CStr(Replace(varSheet(i), "$", ""))): rptItem.Focusable = False
                    Set rptItem = rptRcd.AddItem(CStr("" & mrsSheet(Excel_Col.������).Value)): rptItem.Focusable = False
                   
                    Set rptItem = rptRcd.AddItem(CStr("" & mrsSheet(Excel_Col.�Ǽ�ģ��).Value)): rptItem.Focusable = False
                    
                    Set rptItem = rptRcd.AddItem(CStr("" & mrsSheet(Excel_Col.Ӱ��ģ��).Value)): rptItem.Focusable = False
                    Set rptItem = rptRcd.AddItem(CStr("" & mrsSheet(Excel_Col.Ӱ������).Value)): rptItem.Focusable = False
                    Set rptItem = rptRcd.AddItem(CStr("" & mrsSheet(Excel_Col.��������˵��).Value)): rptItem.Focusable = False
                    Set rptItem = rptRcd.AddItem(CStr("" & mrsSheet(Excel_Col.�Ǽ��û�).Value)): rptItem.Focusable = False
                    Set rptItem = rptRcd.AddItem(CStr("" & mrsSheet(Excel_Col.�û�����).Value)): rptItem.Focusable = False
                    Set rptItem = rptRcd.AddItem(CStr("" & mrsSheet(Excel_Col.�޸�˵��).Value)): rptItem.Focusable = False
                    Set rptItem = rptRcd.AddItem(CStr("" & mrsSheet(Excel_Col.�������).Value)): rptItem.Focusable = False
                    Set rptItem = rptRcd.AddItem(CStr("" & mrsSheet(Excel_Col.��ע).Value)): rptItem.Focusable = False
                    
                    '---- �û�������
'                    Set rptItem2 = rptRcd.AddItem("")
                    If "" & mrsSheet(Excel_Col.�Ƿ���Ҫ��ѵ) = "��" Then
'                         rptItem2.Icon=-1
                        Set rptItem2 = rptRcd.AddItem("������ѵ") '������ѵ
                        
                    Else
                        If "" & mrsSheet(Excel_Col.������ѵ���).Value = "" Then
'                            rptItem2.Icon = ICON_UnTrain
                            Set rptItem2 = rptRcd.AddItem("��")

                        Else
'                            rptItem2.Icon = ICON_Train
                            Set rptItem2 = rptRcd.AddItem("����ѵ")
       
                        End If
                    End If
                    
                    rptRcd.AddItem Val(CStr("" & mrsSheet(Excel_Col.���û�Ӱ������).Value))   '0-δ��д 1-����Ӱ�� 2-����Ӱ�� 3-��Ӱ��
                    
                    '--- ������,�Ƿ��޸�
                    rptRcd.AddItem CStr("" & mrsSheet(Excel_Col.�Ƿ���HTML�ĵ�).Value)
                    rptRcd.AddItem CStr("0")
                    rptRcd.PreviewText = "" & mrsSheet(Excel_Col.�û�����).Value
                End If
                
            End With
            
            mrsSheet.MoveNext
        Loop
        If lngCount > 0 Then lblA.Caption = Trim(lblA.Caption & " " & Replace(varSheet(i), "$", "") & "(" & lngCount & ")")
    Next
    Set mrsSheet = Nothing
        
    '��λ���ϴ�ѡ����
    If mlngItemID <> 0 Then
        For Each rptRow In Me.rptList.Rows
            If rptRow.GroupRow = False Then
                If Val(rptRow.Record(mCol.���).Value) = mlngItemID Then
                    Set Me.rptList.FocusedRow = rptRow
                    Exit For
                End If
            End If
        Next
    End If
    
    'չ��ѡ����
    If Me.rptList.FocusedRow Is Nothing And Me.rptList.Rows.Count > 0 Then
        If Me.rptList.Rows(0).GroupRow Then
            Set Me.rptList.FocusedRow = Me.rptList.Rows(0).Childs(0)
        Else
            Set Me.rptList.FocusedRow = Me.rptList.Rows(0)
        End If
    End If
    
    rptList.Populate
    Call rptList_SelectionChanged '����ѡ���¼�
    
    '�����ѯ����
    m����.���յȼ� = cbo���յȼ�.List(cbo���յȼ�.ListIndex)
    m����.�Ƿ��Ķ� = cbo�Ƿ��Ķ�.List(cbo�Ƿ��Ķ�.ListIndex)
    m����.�Ƿ���ѵ = cbo��ѵ.List(cbo��ѵ.ListIndex)
    m����.ģ�� = Replace(tvwLeft.SelectedItem.Text, "(�°�)", "")
    If InStr(m����.ģ��, "(") > 0 Then m����.ģ�� = Mid(m����.ģ��, 1, InStr(m����.ģ��, "(") - 1)
    
    m����.��ʼ�汾 = cbo��ʼ�汾.List(cbo��ʼ�汾.ListIndex)
    m����.�����汾 = cbo�����汾.List(cbo�����汾.ListIndex)
    m����.�û� = cbo�û�.List(cbo�û�.ListIndex)
    
    m����.Ӱ������ = cboӰ������.List(cboӰ������.ListIndex)
    Call SaveSetting("ZLSOFT", "����ģ��\UpgradeReader", "���յȼ�", m����.���յȼ�)
    Call SaveSetting("ZLSOFT", "����ģ��\UpgradeReader", "�Ƿ��Ķ�", m����.�Ƿ��Ķ�)
    Call SaveSetting("ZLSOFT", "����ģ��\UpgradeReader", "�Ƿ���ѵ", m����.�Ƿ���ѵ)
    Call SaveSetting("ZLSOFT", "����ģ��\UpgradeReader", "��ʼ�汾", m����.��ʼ�汾)
    Call SaveSetting("ZLSOFT", "����ģ��\UpgradeReader", "�����汾", m����.�����汾)
    Call SaveSetting("ZLSOFT", "����ģ��\UpgradeReader", "ģ��", m����.ģ��)
    Call SaveSetting("ZLSOFT", "����ģ��\UpgradeReader", "�û�", m����.�û�)
    Call SaveSetting("ZLSOFT", "����ģ��\UpgradeReader", "Ӱ������", m����.Ӱ������)

'    If int���� > 0 Then
'        '������һ��ԭ�����������ʧЧ����Ҫ����ˢ��
'        Call ReLoad
'    End If
    
    Exit Sub
errHandle:
    MsgBox Err.Number & " " & Err.Description, vbQuestion, gstrSysname
    
End Sub

Private Function LoadInitIcon()
    'װ��ICON�� imgTree�ؼ�
    Dim intIcon As Integer
    Dim strIcon As String
    
    strIcon = ","
    With imgTree
        .ListImages.Clear
        .ImageHeight = 16
        .ImageWidth = 16
    End With
    
    For intIcon = 99 To 240
        imgTree.ListImages.Add , "K_" & intIcon, mclsAppTool.GetIcon(intIcon)
    Next
    Set Me.tvwLeft.ImageList = imgTree

End Function

Private Sub SaveItem()
'�����޸Ľ��
    Dim str��� As String, strSheet As String, strӰ������ As String, str��ѵ As String
    Dim i As Long
    With rptList
        For i = 0 To .Records.Count - 1
            
            If .Records(i).Item(mCol.�޸�).Value = "1" Or .Records(i).Item(mCol.��ѵ).Value <> " " Then
                str��� = .Records(i).Item(mCol.���).Value
                strSheet = .Records(i).Item(mCol.����).Value & "$"
                strӰ������ = Val(.Records(i).Item(mCol.Ӱ������).Value)
                str��ѵ = .Records(i).Item(mCol.��ѵ).Value
                If str��� <> "" Then
                
                    Set mrsSheet = OpenExcelSheet(strSheet)
                    mrsSheet.Filter = 0
                    
                    If mrsSheet.RecordCount > 0 Then mrsSheet.MoveFirst
                    
                    mrsSheet.Filter = "" & mrsSheet(Excel_Col.������).Name & " = " & str���
                    Do Until mrsSheet.EOF
                        mrsSheet(Excel_Col.���û�Ӱ������) = Switch(strӰ������ = 0, "0-δ��д", strӰ������ = 1, "1-��������", strӰ������ = 2, "2-��������", strӰ������ = 3, "3-��Ӱ��")
                        mrsSheet(Excel_Col.�Ķ���¼) = gstrDBUser
                        If mrsSheet(Excel_Col.�Ƿ���Ҫ��ѵ) <> "��" And str��ѵ <> "������ѵ" Then
                            mrsSheet(Excel_Col.������ѵ���) = IIf(str��ѵ = "����ѵ", "����ѵ", "")
                        End If
                        mrsSheet.Update
                        mrsSheet.MoveNext
                    Loop
                    
                    Set mrsSheet = Nothing
                End If
            End If
        Next
    End With
    
End Sub

Private Function IsAdd(ByVal bln����ģ������ As Boolean) As Boolean
    '   �Ƿ�������������Ϸ���true
    Dim strģ�� As String, str��ʼ�汾 As String, str�����汾 As String, strTmp As String
    Dim str�û� As String, strӰ��ģ�� As String
    Dim str���յȼ� As String, str�Ƿ��Ķ� As String, str�Ƿ���ѵ As String, strӰ������ As String
    IsAdd = False
        
    '
    If Val("" & mrsSheet(Excel_Col.������).Value) = 0 Then Exit Function
    If Trim("" & mrsSheet(Excel_Col.�����汾).Value) = "" Then Exit Function
    '0-δ��¼ ��ʾ�����û�
    '1-�ѵ�¼ ֻ�ܵ�ǰ�û���
'    If mIntType = 1 And gstr�û���λ���� <> "" Then
'        If "" & mrsSheet(Excel_Col.�Ǽ��û�) <> gstr�û���λ���� Then Exit Function
'    ElseIf mIntType = 0 Then
        
        If mstrFileName <> mLastFileName Then
            '���ε��ã�ȡע����е�����
            str�û� = m����.�û�
        Else
            If cbo�û�.ListIndex >= 0 Then str�û� = cbo�û�.List(cbo�û�.ListIndex)
        End If

        
        If str�û� <> "����" Then
            strTmp = Trim("" & mrsSheet(Excel_Col.�Ǽ��û�))
            If str�û� <> strTmp Then Exit Function
        End If
'    End If
    
    '-- ģ�� ,    '-- Ӱ��ģ��

    If bln����ģ������ Then '���ģ�����ʱ��������ģ��
         
        If mLastFileName <> mstrFileName Then
            strģ�� = m����.ģ��
        Else
            If Not tvwLeft.SelectedItem Is Nothing Then
                If tvwLeft.SelectedItem.Key <> "Root" Then
                    strģ�� = tvwLeft.SelectedItem.Text
                Else
                    strģ�� = "����ģ��"
                End If
            Else
                strģ�� = "����ģ��"
            End If
        End If
        
        If InStr(strģ��, "����ģ��") <= 0 Then
            strģ�� = Trim(Replace(strģ��, "(�°�)", ""))
            If InStr(strģ��, "(") > 0 Then strģ�� = Mid(strģ��, 1, InStr(strģ��, "(") - 1)
            strTmp = Trim(Replace("" & mrsSheet(Excel_Col.�Ǽ�ģ��), "(�°�)", ""))
            strӰ��ģ�� = Trim(Replace("" & mrsSheet(Excel_Col.Ӱ��ģ��), "(�°�)", ""))
            
            If strTmp <> strģ�� Then
                If strӰ��ģ�� <> "" Then
                    If InStr(strӰ��ģ��, strģ��) <= 0 Then Exit Function
                Else
                    Exit Function
                End If
            End If
        End If
    End If
    
    '-- �汾
    
        If mLastFileName <> mstrFileName Then
            str��ʼ�汾 = m����.��ʼ�汾
            str�����汾 = m����.�����汾
        Else
            If cbo��ʼ�汾.ListIndex >= 0 And cbo�����汾.ListIndex >= 0 Then
                str��ʼ�汾 = cbo��ʼ�汾.List(cbo��ʼ�汾.ListIndex)
                str�����汾 = cbo�����汾.List(cbo�����汾.ListIndex)
            Else
                str��ʼ�汾 = "10.19.0"
                str�����汾 = "90.19.0"
            End If
        End If
        strTmp = Trim("" & mrsSheet(Excel_Col.�����汾))

        If strTmp = "" Then Exit Function
        If strTmp < str��ʼ�汾 Or strTmp > str�����汾 Then Exit Function
    
    
    '-- ���յȼ�
    If cbo���յȼ�.ListIndex >= 0 Then
        str���յȼ� = cbo���յȼ�.List(cbo���յȼ�.ListIndex)
        If str���յȼ� <> "����" Then
            strTmp = Trim("" & mrsSheet(Excel_Col.�������))
            If strTmp <> str���յȼ� Then Exit Function
        End If
    End If
    
    '-- �Ƿ��Ķ�
    If cbo�Ƿ��Ķ�.ListIndex >= 0 Then
        str�Ƿ��Ķ� = cbo�Ƿ��Ķ�.List(cbo�Ƿ��Ķ�.ListIndex)
        If str�Ƿ��Ķ� <> "����" Then
            strTmp = Trim("" & mrsSheet(Excel_Col.���û�Ӱ������))
            If str�Ƿ��Ķ� = "����" Then
                If Val(strTmp) = 0 Then Exit Function
            Else
                If Val(strTmp) <> 0 Then Exit Function
            End If
        End If
    End If
    
    '-- �Ƿ���ѵ
    If cbo��ѵ.ListIndex >= 0 Then
        str�Ƿ���ѵ = cbo��ѵ.List(cbo��ѵ.ListIndex)
        If str�Ƿ���ѵ <> "����" Then
            strTmp = Trim("" & mrsSheet(Excel_Col.�Ƿ���Ҫ��ѵ))
            
            If strTmp = "��" Then
                strTmp = "" & mrsSheet(Excel_Col.������ѵ���)
            Else
                strTmp = "������ѵ"
            End If
            
            If strTmp = "" Then strTmp = "��" 'δ�����δ��ѵ
            If InStr(strTmp, str�Ƿ���ѵ) <= 0 Then Exit Function
        End If
    End If
    
    '-- Ӱ������
    If cboӰ������.ListIndex >= 0 Then
        strӰ������ = cboӰ������.List(cboӰ������.ListIndex)
        If strӰ������ <> "����" Then
            strTmp = Trim("" & mrsSheet(Excel_Col.Ӱ������))
            If strTmp <> strӰ������ Then Exit Function
        End If
    End If

    If mLastFileName <> mstrFileName Then mLastFileName = mstrFileName
    IsAdd = True
    
End Function

Private Sub AddExcelToTree(ByVal strNodeKey As String)

    '���Excel�е�ģ�鵽ģ���б���
    Dim str��� As String, strģ�� As String, blnAdd As Boolean, IntCount As Integer
    Dim objNote As Node, strNoteTxt As String
    strģ�� = Replace(Trim("" & mrsSheet(Excel_Col.�Ǽ�ģ��).Value), "(�°�)", "")
    str��� = Val(Trim("" & mrsSheet(Excel_Col.������)))
    
    If strģ�� = "" Or Val(str���) = 0 Then Exit Sub
    With tvwLeft
        If tvwLeft.Nodes.Count > 0 Then
            blnAdd = True
            For IntCount = 1 To tvwLeft.Nodes.Count
                strNoteTxt = tvwLeft.Nodes(IntCount).Text
                
                If InStr(strNoteTxt, "(") > 0 Then strNoteTxt = Mid(strNoteTxt, 1, InStr(strNoteTxt, "(") - 1)
                If strģ�� = strNoteTxt Then
                    blnAdd = False
                    Exit For
                End If
            Next
            
            If blnAdd Then
                If strNodeKey = "" Then
                    Set objNote = tvwLeft.Nodes.Add(, , "E" & Val(str���), strģ��, "K_" & 170)
                Else
                    Set objNote = tvwLeft.Nodes.Add(strNodeKey, 4, "E" & Val(str���), strģ��, "K_" & 170)
                End If
            End If
        End If
    End With
End Sub

Private Sub ReLoad()
    'ˢ��
    
    If mblnEdit = True Then
        If MsgBox("�������޸�δ���棬�Ƿ������", vbYesNo + vbDefaultButton2, gstrSysname) = vbNo Then Exit Sub
        mblnEdit = False
    End If
    'Call initTree
    Dim objNode As Node
    If Not tvwLeft.SelectedItem Is Nothing Then
        Set objNode = tvwLeft.SelectedItem
        Call tvwLeft_NodeClick(objNode)
    End If
End Sub
Private Sub AddParent(ByVal intIndex As Integer)
    '���¸��ڵ��ģ�����
    Dim strTxt As String
    If tvwLeft.Nodes(intIndex).Parent Is Nothing Then Exit Sub
    strTxt = tvwLeft.Nodes(intIndex).Parent.Text
    If InStr(strTxt, "(") > 0 Then strTxt = Mid(strTxt, 1, InStr(strTxt, "(") - 1)
    tvwLeft.Nodes(intIndex).Parent.Tag = Val(tvwLeft.Nodes(intIndex).Parent.Tag) + 1
    tvwLeft.Nodes(intIndex).Parent.Text = strTxt & "(" & Val(tvwLeft.Nodes(intIndex).Parent.Tag) & ")"
    Call AddParent(tvwLeft.Nodes(intIndex).Parent.Index)
End Sub

Private Sub addModleToTree(ByVal str�ϼ�ģ�� As String, ByVal strKey As String)
    '����ģ�鵽�б�
    Dim strģ�� As String, varģ�� As Variant, intģ�� As Integer
    Dim objNode As Node
    If str�ϼ�ģ�� = "" Then Exit Sub
    
    strģ�� = Trim(ReadFromIni(App.Path & "\" & App.EXEName & ".ini", str�ϼ�ģ��, "ģ��"))
    varģ�� = Split(strģ��, "|")
    For intģ�� = LBound(varģ��) To UBound(varģ��)
        If varģ��(intģ��) <> "" Then
            Set objNode = Me.tvwLeft.Nodes.Add(strKey, 4, strKey & Format(intģ�� + 1, "000"), varģ��(intģ��), "K_" & 99)
            Call addModleToTree(varģ��(intģ��), strKey & Format(intģ�� + 1, "000"))
        End If
    Next
End Sub

