VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmWin 
   BackColor       =   &H80000008&
   BorderStyle     =   0  'None
   ClientHeight    =   4395
   ClientLeft      =   -30
   ClientTop       =   -315
   ClientWidth     =   8040
   ControlBox      =   0   'False
   Icon            =   "Frm������.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   8040
   WindowState     =   2  'Maximized
   Begin MSWinsockLib.Winsock winSock 
      Left            =   5520
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmrUpdateConnect 
      Enabled         =   0   'False
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer TimeToolTipText 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3840
      Top             =   690
   End
   Begin VB.PictureBox PicToolTipText 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FFFF&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4860
      ScaleHeight     =   225
      ScaleWidth      =   1485
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1500
      Visible         =   0   'False
      Width           =   1515
      Begin VB.Label LblToolTipText 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ʾ��Ϣ"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   30
         TabIndex        =   22
         Top             =   30
         Width           =   720
      End
   End
   Begin VB.PictureBox PicBackBitmap 
      AutoRedraw      =   -1  'True
      Height          =   585
      Left            =   360
      Picture         =   "Frm������.frx":1CFA
      ScaleHeight     =   525
      ScaleWidth      =   1605
      TabIndex        =   19
      Top             =   600
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.PictureBox PicRollUp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   165
      Index           =   0
      Left            =   390
      ScaleHeight     =   165
      ScaleWidth      =   2505
      TabIndex        =   18
      Top             =   1680
      Visible         =   0   'False
      Width           =   2505
      Begin VB.Image ImgRollUp 
         Height          =   240
         Index           =   0
         Left            =   1110
         Picture         =   "Frm������.frx":DEB2
         Stretch         =   -1  'True
         Top             =   -30
         Width           =   240
      End
   End
   Begin VB.PictureBox PicRollDown 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   165
      Index           =   0
      Left            =   390
      ScaleHeight     =   165
      ScaleWidth      =   2505
      TabIndex        =   17
      Top             =   1890
      Visible         =   0   'False
      Width           =   2505
      Begin VB.Image ImgRollDown 
         Height          =   240
         Index           =   0
         Left            =   1110
         Picture         =   "Frm������.frx":DFFC
         Stretch         =   -1  'True
         Top             =   -30
         Width           =   240
      End
   End
   Begin VB.Timer TimePass 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   0
      Top             =   0
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   3780
      Top             =   1950
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer SetZorder 
      Interval        =   10
      Left            =   3420
      Top             =   1950
   End
   Begin VB.PictureBox Pic������ 
      AutoRedraw      =   -1  'True
      Height          =   435
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   7965
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3600
      Width           =   8025
      Begin VB.PictureBox Pic���� 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   345
         Index           =   0
         Left            =   4650
         ScaleHeight     =   285
         ScaleWidth      =   2055
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   25
         Visible         =   0   'False
         Width           =   2115
         Begin VB.Label Lbl���� 
            Height          =   165
            Index           =   0
            Left            =   480
            TabIndex        =   11
            Top             =   90
            Width           =   1365
         End
         Begin VB.Image Img���� 
            Height          =   285
            Index           =   0
            Left            =   90
            Stretch         =   -1  'True
            Top             =   60
            Width           =   285
         End
      End
      Begin VB.PictureBox Pic�ָ� 
         AutoRedraw      =   -1  'True
         Height          =   315
         Left            =   990
         MousePointer    =   9  'Size W E
         ScaleHeight     =   255
         ScaleWidth      =   0
         TabIndex        =   5
         Top             =   60
         Width           =   60
      End
      Begin VB.PictureBox Pic��ʼ 
         AutoRedraw      =   -1  'True
         Height          =   345
         Left            =   30
         ScaleHeight     =   285
         ScaleWidth      =   795
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   30
         Width           =   855
         Begin VB.PictureBox PicImg 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   30
            ScaleHeight     =   255
            ScaleWidth      =   405
            TabIndex        =   3
            Top             =   30
            Width           =   405
            Begin VB.Image Img��ʼ 
               Height          =   240
               Left            =   60
               Picture         =   "Frm������.frx":E146
               Stretch         =   -1  'True
               Top             =   30
               Width           =   270
            End
         End
         Begin VB.Label Lbl��ʼ 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��ʼ"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   435
            TabIndex        =   4
            Top             =   90
            Width           =   360
         End
      End
      Begin ComctlLib.StatusBar Sbar 
         Height          =   375
         Left            =   6810
         TabIndex        =   12
         Top             =   0
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   661
         SimpleText      =   ""
         _Version        =   327682
         BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
            NumPanels       =   1
            BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Style           =   5
               Alignment       =   1
               Object.Width           =   2117
               MinWidth        =   2117
               TextSave        =   "15:49"
               Key             =   ""
               Object.Tag             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.PictureBox Pic���ù��� 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   0
         Left            =   570
         ScaleHeight     =   345
         ScaleWidth      =   345
         TabIndex        =   20
         Top             =   30
         Visible         =   0   'False
         Width           =   345
         Begin VB.Image Img���ù��� 
            Height          =   285
            Index           =   0
            Left            =   30
            Stretch         =   -1  'True
            Top             =   30
            Width           =   285
         End
      End
      Begin VB.Line LineRight 
         BorderColor     =   &H80000005&
         X1              =   990
         X2              =   990
         Y1              =   60
         Y2              =   360
      End
      Begin VB.Line LineLeft 
         BorderColor     =   &H80000003&
         X1              =   960
         X2              =   960
         Y1              =   60
         Y2              =   360
      End
   End
   Begin VB.Frame FraSplit 
      Height          =   30
      Index           =   0
      Left            =   360
      TabIndex        =   13
      Top             =   2370
      Visible         =   0   'False
      Width           =   2565
   End
   Begin VB.PictureBox PicBackDesktop 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   555
      Index           =   0
      Left            =   390
      ScaleHeight     =   555
      ScaleWidth      =   2505
      TabIndex        =   7
      Top             =   2550
      Visible         =   0   'False
      Width           =   2500
      Begin VB.PictureBox Pic�˵� 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   465
         Index           =   0
         Left            =   30
         ScaleHeight     =   465
         ScaleWidth      =   1995
         TabIndex        =   8
         TabStop         =   0   'False
         Tag             =   ";FileExit"
         Top             =   30
         Visible         =   0   'False
         Width           =   1995
         Begin VB.Label Lbl��� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   0
            Left            =   930
            TabIndex        =   15
            Top             =   150
            Width           =   90
         End
         Begin VB.Image Img�˵� 
            Height          =   480
            Index           =   0
            Left            =   60
            Stretch         =   -1  'True
            Top             =   60
            Width           =   480
         End
         Begin VB.Label Lbl�˵� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   0
            Left            =   615
            TabIndex        =   9
            Top             =   150
            Width           =   180
         End
         Begin VB.Image Img�˵�ָʾ 
            Height          =   150
            Index           =   0
            Left            =   1200
            Picture         =   "Frm������.frx":E710
            Stretch         =   -1  'True
            Top             =   210
            Visible         =   0   'False
            Width           =   120
         End
      End
   End
   Begin VB.PictureBox Pic��ʶ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   3645
      Left            =   30
      ScaleHeight     =   3645
      ScaleWidth      =   345
      TabIndex        =   6
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   0
      Visible         =   0   'False
      Width           =   345
      Begin VB.Label Lbl��ʶ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   60
         TabIndex        =   16
         Top             =   2460
         Width           =   90
      End
      Begin VB.Image Img��ʶ 
         Height          =   3495
         Left            =   30
         Stretch         =   -1  'True
         Tag             =   "1"
         Top             =   120
         Width           =   285
      End
   End
   Begin MSComctlLib.ImageList ImgLvw 
      Left            =   3180
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Timer RefreshMenu 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1860
      Top             =   0
   End
   Begin MSComctlLib.ListView LvwList 
      Height          =   3795
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   7905
      _ExtentX        =   13944
      _ExtentY        =   6694
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   0
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   14
      Top             =   4035
      Width           =   8040
      _ExtentX        =   14182
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2223
            MinWidth        =   882
            Picture         =   "Frm������.frx":E85A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8996
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1429
            MinWidth        =   1411
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
            Object.ToolTipText     =   "��ǰ���ּ�״̬"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1411
            MinWidth        =   1411
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
            Object.ToolTipText     =   "��ǰ��д��״̬"
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
   Begin VB.Menu MnuRightMenu 
      Caption         =   "�Ҽ��˵�(&R)"
      Visible         =   0   'False
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)"
      End
      Begin VB.Menu MnuBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolTester 
         Caption         =   "ʹ��SQL�ٶȲ��Թ���(&U)"
         Checked         =   -1  'True
         Visible         =   0   'False
      End
      Begin VB.Menu MnuToolIndividuation 
         Caption         =   "ʹ�ø��Ի�����(&I)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuToolNotify 
         Caption         =   "��Ϣ֪ͨ(&N)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuToolShowDisReport 
         Caption         =   "��ʾͣ�ñ���(&P)"
         Checked         =   -1  'True
      End
      Begin VB.Menu MnuBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolDictonary 
         Caption         =   "�ֵ������(&D)"
      End
      Begin VB.Menu mnuToolMessage 
         Caption         =   "��Ϣ�շ�����(&M)"
      End
      Begin VB.Menu mnuToolNotice 
         Caption         =   "������Ϣ����(&T)"
      End
      Begin VB.Menu mnuToolStyle 
         Caption         =   "ϵͳѡ��(&S)"
      End
      Begin VB.Menu mnuToolExcel 
         Caption         =   "����&EXCEL����"
      End
      Begin VB.Menu MnuBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolHistory 
         Caption         =   "�����ʷ��¼(&H)"
      End
      Begin VB.Menu MnuBar5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolSetColor 
         Caption         =   "����������ɫ(&O)"
      End
      Begin VB.Menu mnuToolSelBackBmp 
         Caption         =   "ѡ�񱳾�ͼƬ(&B)"
      End
      Begin VB.Menu mnuToolOutTool 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolOutToolSet 
         Caption         =   "��ӹ�������(&O)��"
      End
      Begin VB.Menu mnuToolOutToolList 
         Caption         =   "��ӹ���(&G)"
         Visible         =   0   'False
         Begin VB.Menu mnuToolOutToolExecute 
            Caption         =   "����(&1)"
            Index           =   0
         End
      End
      Begin VB.Menu MnuBar6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRepairIndividuationClear 
         Caption         =   "������������쳣(&L)"
      End
      Begin VB.Menu mnuRepairComponent 
         Caption         =   "��ⰲװ����(&C)"
      End
      Begin VB.Menu mnuRepairClientUpdate 
         Caption         =   "�ͻ����޸�(&U)"
      End
      Begin VB.Menu MnuBar7 
         Caption         =   "-"
      End
      Begin VB.Menu MnuReg 
         Caption         =   "ע��(&R)"
      End
      Begin VB.Menu MnuExit 
         Caption         =   "�˳�(&X)"
      End
   End
End
Attribute VB_Name = "FrmWin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnPress As Boolean                            '��ʼ��ť�Ƿ���ѹ��״̬
Private mintLevel As Integer                            '��ǰ���ڵڼ���
Private mintLast��� As Integer                         '��һ�ε����
Private mlngSelectModul As Long                         'ѡ���ģ��
Private mlngSelectUsual As Long                         'ѡ��ĳ��ù���
Private mblnMenuOpened As Boolean                       '��ѡ��Ĳ˵��Ƿ��Ѿ���
Private mdblMenuWidth As Double                         '��ǰ�˵��ĸ߶�
Private mdblMenuHeight As Double                        '��ǰ�˵��������
Private mblnFirst As Boolean                            '��һ�������ɹ�
Private mblnShow As Boolean
Private mstrLastSelectCaption As String                 '�ϴ���ѡ����ı���
Private mlngLastSelectIndex As Long                     '�ϴ���ѡ�����Ӧ��������������
Private mFrmChildObj As Form                            '�Ӵ������
Private mCurTime As Date                                '��ǰԤ����ʱ�����.
Private mblnAdjustPost As Boolean
Private mcllTemp As Collection
Private marrRoll(256) As String                         '--����ÿ���˵���������˵��������
Private mstrTitle As String                             '��Ʒ����
Private mblnHide As Boolean                             '�Ƿ���ʾ������
Private Const M_INT_RPTDISABLED As Integer = 242        '���ñ���ͼ��
Private WithEvents mclsMipModule As zl9ComLib.clsMipModule
Attribute mclsMipModule.VB_VarHelpID = -1
Public mclsAppTool As New zl9AppTool.clsAppTool
Private mblnRemote As Boolean '�Ƿ���Զ��

Public Property Get frmHide() As Boolean
    frmHide = mblnHide
End Property

Public Property Get ObjLogin() As Object
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.���
    Set ObjLogin = gobjRelogin
End Property

Public Property Get mobjEmr() As Object
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.���
    Set mobjEmr = gobjRelogin.EMR
End Property

Private Sub Form_Activate()
    If Not mblnFirst Then Exit Sub
    
    mblnFirst = False
    mblnMenuOpened = False
    
    '--װ��˵���--
    If LoadLvw = False Then
        Unload Me
        Exit Sub
    End If
    Call LoadUsual
    
    '�˶α����ڴ���ͬ��ʺ�(����Ϣ֪ͨ����ZlAppTool����,ִ���亯��--GetUserInfoʱ����)
    MnuToolIndividuation.Checked = IIf(Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 0, False, True)
    mnuToolNotify.Checked = IIf(Val(zlDatabase.GetPara("�����ʼ���Ϣ")) = 0, True, False)
    mnuToolTester.Checked = IIf(GetSetting("ZLSOFT", "����ȫ��", "SQLTest", 0) = 0, False, True)
    mnuToolShowDisReport.Checked = IIf(Val(zlDatabase.GetPara("��ʾͣ�ñ���")) = 0, False, True)
    mnuToolNotify_Click
    Call SetMainForm(Me)
    Call InitEvn
    
    '���ֻ��һ����ģ��,���
    On Error Resume Next
    With grsMenus
        .Filter = "ģ��<>0 And ����=0"
        If Not .EOF Then
            If .RecordCount = 1 Then
                Call ExecuteFunc(.Fields("ϵͳ").Value, IIf(IsNull(.Fields("����").Value), "", .Fields("����").Value), .Fields("ģ��").Value)
            End If
        End If
        .Filter = 0
    End With
    
    If mblnShow = False Then MsgBox "��ʾ����ͼƬʱ���������󣡣��ָ�ΪȱʡͼƬ��", vbInformation, gstrSysName
    Call LoadOutTools(False)
    
    '������Ϣ����ƽ̨�ͻ����շ�����
    '------------------------------------------------------------------------------------------------------------------
    If ConnectMip(Me.hwnd) = True Then
        Set mclsMipModule = New zl9ComLib.clsMipModule
        Call mclsMipModule.InitMessage(0, 0, "")
        Call AddMipModule(mclsMipModule)
    End If
    
    '�����Զ����ѷ���
    mclsAppTool.CodeMan 0, 5, gcnOracle, Me, gstrDbUser
    If mblnHide Then Me.Hide '���ⲿ���ã�����������,by �¶�
End Sub

Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Static StrPass As String                                '��������(Open zlReport.ReportMan )
    Dim objItem As ListItem, blnExist As Boolean
    
    If mblnPress And (KeyCode >= vbKeyA And KeyCode < vbKeyZ) Then
        Call FindMenu(KeyCode)
        Exit Sub
    End If
    
    '--���ز˵�--
    If KeyCode = vbKeyEscape And mblnPress Then ShowMenu
    If KeyCode = vbKeyW And Shift = vbCtrlMask Then Pic��ʼ_MouseDown 1, 0, 0, 0
    If KeyCode = vbKeyF4 And Shift = vbAltMask Then MnuExit_Click: Exit Sub
    
    TimePass.Enabled = False
    If KeyCode = vbKeyF12 And Shift = 7 Then
        StrPass = ""
        Exit Sub
    End If
    
    If KeyCode <> vbKeyReturn Then
        If InStr(1, "1234567890 ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Chr(KeyCode))) <> 0 Then StrPass = StrPass & UCase(Chr(KeyCode))
        
        If StrPass = "OPEN ZLREPORT REPORTMAN" Then
            If OwnerUser(gstrDbUser) Then
                StrPass = ""
                
                If FindWindow(vbNullString, "�������") <> 0 Then Exit Sub
                If MsgBox("��ȷ��Ҫ�����Զ��屨������", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                Call ExecuteFunc(0, "ZL9REPORT", 99999901)
                SetParent FindWindow(vbNullString, "�������"), Me.hwnd
            End If
        End If
    End If
    TimePass.Enabled = True
End Sub

Private Sub Form_Load()
    Dim IntKind As Integer
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim LngHdl As Long                                  '���������
    Dim intGrant As Integer
    mblnFirst = True
    mblnAdjustPost = False
    Dim strTitle As String, strTag As String
    
    On Error Resume Next
    'ȡϵͳ������
    LngHdl = FindWindow("Shell_traywnd", "")
    Call SetWindowPos(LngHdl, 0, 0, 0, 0, 0, TOGGLE_HIDEWINDOW)
    
    '�ж��Ƿ���Ȩ��ʹ����Ϣ�շ�����
    Call CheckTools
    
    'ʹ�øú�����Ŀ�ľ����������Զ��޸Ĳ˵�(���������ڵ�)
    RestoreWinState Me
    
    '--����Ƿ�Ϊ���ð棨�Ĳ˵��ı�ʶͼƬ��--
    IntKind = IIf(GetSetting("ZLSOFT", "ע����Ϣ", "Kind", "") = "����", -1, 0)
    Set gcllCollMap = New Collection
    
    Me.WindowState = 2
    '���û�׼�˵�
    �˵���׼.���ܲ˵� = 90000001
    �˵���׼.���ڲ˵� = 99990001
    �˵���׼.�������ܲ˵� = 99999901
    �˵���׼.�ָ��˵� = 99999999
    
    Call CheckWinVersion
    
    strTitle = zlRegInfo("��Ʒ����")
    strTag = ""
    If strTitle <> "" Then
        If InStr(strTitle, "-") > 0 Then
            If Split(strTitle, "-")(1) = "Ultimate" Then
                strTag = "�콢��"
            ElseIf Split(strTitle, "-")(1) = "Professional" Then
                strTag = "רҵ��"
            End If
        End If
    End If
    strTitle = Split(strTitle, "-")(0)
    mstrTitle = strTitle & IIf(strTag = "", "", "(" & strTag & ")")
    '�������ݿ����Ӹ���ӡ����
    IniPrintMode gcnOracle, gstrDbUser
    
    '�����жϻỰ���Ƿ�����Ϣ����������
    'select ����ֵ from zloptions where ������ =17
    strSQL = "select ����ֵ from zloptions where ������ =17"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�ж���ѯ�������Ƿ���")
    If rsTemp.RecordCount = 1 Then
        If NVL(rsTemp!����ֵ) <> "" Then
            '������ѯ������,�ر�TIME
            tmrUpdateConnect.Enabled = False
        Else
            'û����ѯ������,ʹ��TIME���� Ԥ�������
            tmrUpdateConnect.Enabled = True
            tmrUpdateConnect.Interval = 30000
            mCurTime = Now
        End If
    Else
        'û����ѯ������,ʹ��TIME���� Ԥ�������
        tmrUpdateConnect.Enabled = True
        tmrUpdateConnect.Interval = 30000
        mCurTime = Now
    End If
    
    '�ⲿ���õĴ���,by �¶�
    mblnHide = False
    If gstrCommand <> "" Then Call DoCommand
    If gobjPlugIn Is Nothing Then
        On Error Resume Next
        Set gobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
        If Not gobjPlugIn Is Nothing Then
            Call gobjPlugIn.Initialize(gcnOracle, 0, 0)
            If InStr(",438,0,", "," & Err.Number & ",") = 0 Then
                MsgBox "zlPlugIn ��Ҳ���ִ�� Initialize ʱ����" & vbCrLf & Err.Number & vbCrLf & Err.Description, vbInformation, gstrSysName
            End If
        End If
        Err.Clear: On Error GoTo 0
    End If
    
    If Not gobjPlugIn Is Nothing Then
        On Error Resume Next
        Call gobjPlugIn.LogInAfter
        If InStr(",438,0,", "," & Err.Number & ",") = 0 Then
            MsgBox "zlPlugIn ��Ҳ���ִ�� LogInAfter ʱ����" & vbCrLf & Err.Number & vbCrLf & Err.Description, vbInformation, gstrSysName
        End If
        Err.Clear: On Error GoTo 0
    End If

    '��ʼ������
    InitWinsock
End Sub

Private Sub Form_Resize()
    Me.WindowState = 2
    With LvwList
        .Top = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - 405
    End With
    With PicBackBitmap
        .Top = 0
        .Width = Me.ScaleWidth
        .Left = 0
        .Height = Me.ScaleHeight - 405
    End With
    
    With Pic������  'Windows ������
        .Left = Me.ScaleLeft
        .Width = Me.ScaleWidth
        .Top = LvwList.Top + LvwList.Height
    End With
    With Sbar
        .Left = Pic������.Width - .Width - 50
    End With
    
    zlControl.PicShowFlat Pic������, 2, , taCenterAlign
    zlControl.PicShowFlat Pic��ʼ, 2, , taCenterAlign
    zlControl.PicShowFlat Pic�ָ�, 2, , taCenterAlign
    
    gLngFormID = Me.hwnd
    Dim StrCaption As String
    StrCaption = mstrTitle ' zlProductTitle(GetUnitInfo("������"))
    
    '--���ô������--
    Call SetWindowText(Me.hwnd, StrCaption)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim LngHdl As Long
    Dim blnCloaseWin As Boolean
    
    On Error Resume Next
    blnCloaseWin = Val(zlDatabase.GetPara("�ر�Windows")) <> 0
    'ȡϵͳ������
    LngHdl = FindWindow("Shell_traywnd", "")
    Call SetWindowPos(LngHdl, 0, 0, 0, 0, 0, TOGGLE_UNHIDEWINDOW)
    '�������ҽ�����Լ�ҵ����
    Call CloseChildWindows(Me)
    '������Ϣ����
    Call DisConnectMip
    If Not (mclsMipModule Is Nothing) Then
        mclsMipModule.CloseMessage
        Call DelMipModule(mclsMipModule)
        Set mclsMipModule = Nothing
    End If
    Call gobjRelogin.Dispose '��Ҫ��ж�ض���
    Set gobjRelogin = Nothing
    SaveSetting "ZLSOFT", "����ȫ��", "SQLTest", 0
    '�������Ĳ���ֵ
    zlDatabase.ClearParaCache
    Call ShutDown(blnCloaseWin)
    If Not gcnOracle Is Nothing Then
        If gcnOracle.State = 1 Then gcnOracle.Close
        Set gcnOracle = Nothing
    End If
    ReDim Preserve gobjCls(0)
    ReDim Preserve gstrObj(0)
End Sub

Private Sub ImgRollDown_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PicRollDown_MouseMove(Index, Button, Shift, x, y)
End Sub

Private Sub ImgRollDown_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PicRollDown_MouseUp(Index, Button, Shift, 0, 0)
End Sub

Private Sub ImgRollUp_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PicRollUp_MouseMove(Index, Button, Shift, x, y)
End Sub

Private Sub ImgRollUp_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PicRollUp_MouseUp(Index, Button, Shift, x, y)
End Sub

Private Sub Img���ù���_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call Pic���ù���_MouseDown(Index, Button, Shift, x, y)
End Sub

Private Sub Img���ù���_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call Pic���ù���_MouseMove(Index, Button, Shift, x, y)
End Sub

Private Sub Img���ù���_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call Pic���ù���_MouseUp(Index, Button, Shift, x, y)
End Sub

Private Sub Img����_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Pic����_MouseDown Index, Button, Shift, x, y
End Sub

Private Sub LblToolTipText_Click()
    Call PicToolTipText_Click
End Sub

Private Sub LblToolTipText_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PicToolTipText_MouseMove(Button, Shift, x, y)
End Sub

Private Sub Lbl����_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Pic����_MouseDown Index, Button, Shift, x, y
End Sub

Private Sub LvwList_DblClick()
    Dim LngFindWindows As Long                          'Ŀ�괰��
    
    If LvwList.ListItems.Count = 0 Then Exit Sub
    If LvwList.SelectedItem.Key = "HELP" Then
        Shell "hh.exe  zl9start.chm", vbNormalFocus
        Exit Sub
    End If
    If LvwList.SelectedItem.Key = "EXIT" Then
        MnuExit_Click
        Exit Sub
    End If
    
    If LvwList.SelectedItem.Tag = -1 Then
        '--ִ�и�ģ��--
        With grsMenus
            .MoveFirst
            .Find "���='" & Mid(LvwList.SelectedItem.Key, 3) & "'"
            
            Call ExecuteFunc(.Fields("ϵͳ").Value, IIf(IsNull(.Fields("����").Value), "", .Fields("����").Value), .Fields("ģ��").Value)
        End With
    Else
        '--�򿪸�ģ��--
        Call OpenWindow(Mid(LvwList.SelectedItem.Key, 3), LvwList.SelectedItem.Text)
    End If
End Sub

Public Function LoadLvw() As Boolean
    LoadLvw = False
    
    With grsMenus
        .Filter = "�ϼ�=0"
        LvwList.ListItems.Clear
        If .EOF Then
            MsgBox "��û�в�����ϵͳ��Ȩ�ޣ�"
            grsMenus.Filter = 0
            Exit Function
        End If
        
        On Error Resume Next
        Do While Not .EOF 'ΪImageListװ��ͼ��
            ImgLvw.ListImages.Add ImgLvw.ListImages.Count + 1, "K_" & ImgLvw.ListImages.Count + 1, GetPicDisp(!ͼ��, False)
            .MoveNext
        Loop
        
        ImgLvw.ListImages.Add ImgLvw.ListImages.Count + 1, "HELP", GetPicDisp(-1)
        ImgLvw.ListImages.Add ImgLvw.ListImages.Count + 1, "EXIT", GetPicDisp(-3)
        
        Set LvwList.Icons = ImgLvw
        .MoveFirst
        Do While Not .EOF
            LvwList.ListItems.Add , "K_" & !���, !����, .AbsolutePosition
            LvwList.ListItems("K_" & !���).Tag = IIf((!ģ��) = 0, 0, -1)
            .MoveNext
        Loop
        LvwList.ListItems.Add , "HELP", "����", "HELP"
        LvwList.ListItems.Add , "EXIT", "�˳�ϵͳ", "EXIT"
        
        .Filter = 0
    End With
    
    LoadLvw = True
End Function

Private Sub LvwList_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    LvwList.Drag 0
End Sub

Private Sub LvwList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then LvwList_DblClick
End Sub

Private Sub LvwList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If mblnPress Then ShowMenu
    Call Find����(-99999999)
    If Button = 2 Then PopupMenu MnuRightMenu, 2
End Sub

Private Sub LvwList_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then LvwList.Drag 1
End Sub

Private Sub mclsMipModule_ConnectStateChanged(ByVal IsConnected As Boolean)
    '����״̬�Ѿ��仯
    If IsConnected Then
        tmrUpdateConnect.Enabled = False
    Else
        tmrUpdateConnect.Enabled = True
        tmrUpdateConnect.Interval = 30000
        mCurTime = Now
    End If
End Sub

Private Sub mclsMipModule_OpenModule(ByVal lngSystem As Long, ByVal lngModule As Long, ByVal strPara As String)
    Call RunModual(lngSystem, lngModule, strPara)
End Sub

Private Sub mclsMipModule_OpenReport(ByVal lngSystem As Long, ByVal lngModule As Long, ByVal strPara As String)
    Call RunModual(lngSystem, lngModule, strPara, True)
End Sub

Private Sub mclsMipModule_ReceiveMessage(ByVal strMessageItemKey As String, ByVal strMessageConent As String)
    Select Case UCase(strMessageItemKey)
    '--------------------------------------------------------------------------------------------------------------
    Case "ZLHIS_PUB_005"            '��Ʒ����֪ͨ
        Call gobjRelogin.UpdateClient
    End Select

End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuRepairClientUpdate_Click()
    If MsgBox("�����������¼�Ȿ�������������Ա����������������޸������޸�������в�����������ע�ᡣ��ȷ��Ҫ���пͻ����޸���", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
        Call gobjRelogin.UpdateClient(True)
    End If
End Sub

Public Sub mnuRepairComponent_Click()
    '--���ע���[��������]--
    SaveSetting "ZLSOFT", "ע����Ϣ", "��������", ""
    MsgBox "���������ϣ����иĶ������µ�¼����Ч��", vbInformation, gstrSysName
End Sub

Private Sub mnuRepairIndividuationClear_Click()
    Dim strSQL As String, rsTmp As Recordset
    Dim strAnalyseComputer As String
    
    If MsgBox("�����������ZLHIS��ص�ע���������Լ����ݿ��д洢�ı��ˡ�������������Ʒ��ع��ܽ�������ȱʡֵ���У���ȷ��Ҫ������", vbYesNo + vbDefaultButton2 + vbQuestion, "������������쳣") = vbYes Then
        strSQL = "Select Distinct ���� From zlPrograms Where ���� Is Not Null"
        On Error GoTo ErrHand
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "������������쳣")
        Do While Not rsTmp.EOF
            Call DelWinState(Me, rsTmp!���� & "")
            rsTmp.MoveNext
        Loop
        strAnalyseComputer = OS.ComputerName
        strSQL = "Zl_zluserparas_Clear('" & gstrDbUser & "','" & strAnalyseComputer & "')"
        Call zlDatabase.ExecuteProcedure(strSQL, gstrSysName)
        MsgBox "����ɹ�����رճ������½��룬ȷ���Ƿ��������쳣���⡣", vbInformation, "������������쳣"
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub mnuToolDictonary_Click()
    mclsAppTool.CodeMan 0, 1, gcnOracle, Me, gstrDbUser
End Sub

Private Sub mnuToolExcel_Click()
    Dim ObjExcel As Object, strHaveSys As String
    
    If gstrUserName = "" Then
        MsgBox "��Ϊ����Ա���ö�Ӧ���û�����ʹ�ñ����ܣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    strHaveSys = gobjRelogin.Systems
    On Error Resume Next
    Err = 0
    Set ObjExcel = CreateObject("Zl9Excel.ClsExcel")
    If Err <> 0 Then
        MsgBox "�޷�����EXCEL��������������ʹ��EXCEL����", vbInformation, gstrSysName
        Exit Sub
    End If
    Call ObjExcel.CodeMan(0, 0, gcnOracle, Me, gstrDbUser)
    Call ObjExcel.SetHaveSys(strHaveSys)
    Call ObjExcel.ExcelReportMain
    Set ObjExcel = Nothing
End Sub

Private Sub mnuToolHistory_Click()
    Call zlDatabase.SetPara("���ʹ��ģ��", "")
End Sub

Private Sub MnuToolIndividuation_Click()
    MnuToolIndividuation.Checked = MnuToolIndividuation.Checked Xor True
    Call zlDatabase.SetPara("ʹ�ø��Ի����", IIf(MnuToolIndividuation.Checked, "1", "0"))
    SaveSetting "ZLSOFT", "˽��ȫ��\" & gstrDbUser, "ʹ�ø��Ի����", IIf(MnuToolIndividuation.Checked, "1", "0")
End Sub

Private Sub mnuToolMessage_Click()
    mclsAppTool.CodeMan 0, 2, gcnOracle, Me, gstrDbUser
End Sub

Private Sub MnuReg_Click()
    If MsgBox("��ȷ��Ҫע����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    Call ReLogin
End Sub

Private Sub MnuExit_Click()
    Dim intStyle As Integer
    
    If Frm�ر�.ShowMe(intStyle) Then
        If intStyle = 0 Then
            ReLogin
        Else
            Unload Me
        End If
    End If
End Sub

Private Sub mnuToolNotice_Click()
    mclsAppTool.CodeMan 0, 6, gcnOracle, Me, gstrDbUser
End Sub

Private Sub mnuToolNotify_Click()
    mnuToolNotify.Checked = Not mnuToolNotify.Checked
    Call zlDatabase.SetPara("�����ʼ���Ϣ", IIf(mnuToolNotify.Checked, "1", "0"))
    mclsAppTool.CodeMan 0, 4, gcnOracle, Me, gstrDbUser, IIf(mnuToolNotify.Checked = True, "Open", "Close")
End Sub

Private Sub mnuToolOutToolExecute_Click(Index As Integer)
    '���˺�:2007/08/22
    '���Ӷ��ⲿ���ߵ�ִ��
    Call ExeCuteToolFile(mnuToolOutToolExecute(Index).Tag)
End Sub

Private Sub mnuToolOutToolSet_Click()
    Dim blnApply As Boolean
    '���˺�:2007/08/22
    '�����ⲿ���ߵ�����
     Call frm��������.ShowEdit(Me, blnApply)
    If blnApply = False Then Exit Sub
    Call LoadOutTools(False)
End Sub

Private Sub mnuToolShowDisReport_Click()
    mnuToolShowDisReport.Checked = Not mnuToolShowDisReport.Checked
    Call zlDatabase.SetPara("��ʾͣ�ñ���", IIf(mnuToolShowDisReport.Checked, 1, 0))
End Sub

Private Sub mnuToolSelBackBmp_Click()
    Dim BlnShow As Boolean              '�ܷ�������ʾ
    Dim StrPicPath As String            '����ͼƬ·��
    '--���û�ѡ�񱳾�ͼƬ--
    On Error GoTo ErrHand
    With Dialog
        .CancelError = True
        .Filter = "����ͼƬ (*.bmp;*.jpg)|*.bmp;*.jpg"
        .ShowOpen
        
        '�û�ѡ��ͼƬ,�����Ƿ�����
        On Error Resume Next
        Err = 0
        BlnShow = False
        StrPicPath = .FileName
        Img��ʶ.Picture = LoadPicture(StrPicPath)
        If Err <> 0 Then
            MsgBox "����ѡ���ͼƬ�ļ���������ʾ��", vbInformation, gstrSysName
        Else
            BlnShow = True
        End If
    End With
    
    PicBackBitmap.Picture = LoadPicture(StrPicPath)
    Call PicBackBitmap.PaintPicture(PicBackBitmap.Picture, 0, 0, PicBackBitmap.Width, PicBackBitmap.Height, _
                    0, 0, PicBackBitmap.Picture.Width * 0.57, PicBackBitmap.Picture.Height * 0.57)
    LvwList.Picture = PicBackBitmap.Image
    '����ͼƬλ�ù��´���ȡ
    Call zlDatabase.SetPara("zlWinBackPic", StrPicPath)
    '�ָ�ԭ�����õ�ͼƬ
    Img��ʶ.Picture = LoadResPicture(101, 0) '�˵���ʶ
ErrHand:
End Sub

Private Sub mnuToolSetColor_Click()
    '--���û�ѡ��������ɫ--
    On Error GoTo ErrHand
    With Dialog
        .CancelError = True
        .ShowColor
        
        LvwList.ForeColor = .Color
        
        '��������ɫ���´���ȡ
        Call zlDatabase.SetPara("zlWinFontColor", .Color)
    End With
ErrHand:
End Sub

Public Sub mnuToolStyle_Click()
    mclsAppTool.CodeMan 0, 3, gcnOracle, Me, gstrDbUser, gstrMenuSys
    If Val(zlDatabase.GetPara("����Զ�̿���")) <> winSock.LocalPort Then
        Call InitWinsock
    End If
    If mclsAppTool.IsRestart Then
        mclsAppTool.IsRestart = False
        Call ReLogin
    Else
        '���¼��س��ù���
        Call ShutUsual
        Pic�ָ�.Left = 990
        Call LoadUsual
    End If
End Sub

Private Sub mnuToolTester_Click()
    mnuToolTester.Checked = mnuToolTester.Checked Xor True
    SaveSetting "ZLSOFT", "����ȫ��", "SQLTest", IIf(mnuToolTester.Checked, 1, 0)
End Sub

Private Sub PicRollDown_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim MouseOver As Boolean
    On Error Resume Next
    
    '--�жϵ�ǰ���λ���Ƿ��ڲ˵���--
    MouseOver = (0 <= x) And (x <= PicRollDown(Index).Width) And (0 <= y) And (y <= PicRollDown(Index).Height)
    If MouseOver Then
        Call ShutMenu(mintLevel)
        Call zlControl.PicShowFlat(PicRollDown(Index), -1, , taCenterAlign)
        Call SetCapture(PicRollDown(Index).hwnd)
    Else
        Call zlControl.PicShowFlat(PicRollDown(Index), 0, , taCenterAlign)
        Call ReleaseCapture
    End If
End Sub

Private Sub PicRollDown_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    Call RollUpMenu(PicRollDown(Index).Tag, 1)
End Sub

Private Sub PicRollUp_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim MouseOver As Boolean
    On Error Resume Next
    
    '--�жϵ�ǰ���λ���Ƿ��ڲ˵���--
    MouseOver = (0 <= x) And (x <= PicRollUp(Index).Width) And (0 <= y) And (y <= PicRollUp(Index).Height)
    If MouseOver Then
        Call ShutMenu(mintLevel)
        Call zlControl.PicShowFlat(PicRollUp(Index), -1, , taCenterAlign)
        Call SetCapture(PicRollUp(Index).hwnd)
    Else
        Call zlControl.PicShowFlat(PicRollUp(Index), 0, , taCenterAlign)
        Call ReleaseCapture
    End If
End Sub

Private Sub PicRollUp_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    Call RollUpMenu(PicRollUp(Index).Tag, 2)
End Sub

Private Sub PicToolTipText_Click()
    PicToolTipText.Visible = False
    TimeToolTipText.Enabled = False
End Sub

Private Sub PicToolTipText_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PicToolTipText_Click
End Sub

Private Sub Pic�˵�_DblClick(Index As Integer)
    '--�����¼��˵����--
    Call PicToolTipText_Click
    If Img�˵�ָʾ(Index).Visible Then
        '--�򿪸�ģ��--
        Call OpenWindow(Img�˵�(Index).Tag, Mid(Lbl�˵�(Index).Caption, 1, IIf(InStr(1, Lbl�˵�(Index).Caption, "(") <> 0, Len(Lbl�˵�(Index).Caption) - 3, Len(Lbl�˵�(Index).Caption))))
        
        If mblnPress Then ShowMenu
    End If
End Sub

Private Sub Pic�˵�_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call ReleaseCapture
    mlngSelectModul = 0
End Sub

Private Sub Pic���ù���_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    '����Ϊ����
    If Button <> 1 Then Exit Sub
    mlngSelectUsual = Index
    Call zlControl.PicShowFlat(Pic���ù���(Index), -2, , taCenterAlign)
End Sub

Private Sub Pic���ù���_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    '����Ϊ͹��
    Dim MouseOver As Boolean
    '--�жϵ�ǰ���λ���Ƿ��ڲ˵���--
    
    If Button = 1 Then Exit Sub
    MouseOver = (0 <= x) And (x <= Pic���ù���(Index).Width) And (0 <= y) And (y <= Pic���ù���(Index).Height)
    If MouseOver Then
        Call zlControl.PicShowFlat(Pic���ù���(Index), 2, , taCenterAlign)
        Call SetCapture(Pic���ù���(Index).hwnd)
        Call ShowToolTipText(Pic���ù���(Index))
    Else
        Call zlControl.PicShowFlat(Pic���ù���(Index), 0, , taCenterAlign)
        Call ReleaseCapture
        Call ShowToolTipText(Pic���ù���(Index), False)
    End If
End Sub

Private Sub Pic���ù���_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lngϵͳ As Long, str���� As String, lngģ�� As Long
    If Button <> 1 Then Exit Sub
    If Not (Index = mlngSelectUsual) Then Exit Sub
    
    mlngSelectUsual = 0
    '�ȷֽ����
    str���� = ""
    lngϵͳ = Split(Pic���ù���(Index).Tag, "��")(0)
    lngģ�� = Split(Pic���ù���(Index).Tag, "��")(1)
    
    grsMenus.Filter = "ϵͳ=" & lngϵͳ & " And ģ��=" & lngģ��
    If grsMenus.RecordCount <> 0 Then str���� = IIf(IsNull(grsMenus!����), "", grsMenus!����)
    grsMenus.Filter = 0
    If str���� = "" Then Exit Sub
    
    '���и�ģ��
    Call zlControl.PicShowFlat(Pic���ù���(Index), 0, , taCenterAlign)
    Call ExecuteFunc(lngϵͳ, str����, lngģ��)
End Sub

Private Sub Pic����_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim IntChange As Integer
    Dim LngActiveWindow As Long
    '���ڱ��漯���е�����
    Dim IntIndex As Integer, LngHdl As Long, intStyle As Integer
    
    '����������ģʽ
    For IntChange = 1 To gcllCollMap.Count
        IntIndex = gcllCollMap("K_" & IntChange)(0)
        LngHdl = gcllCollMap("K_" & IntChange)(1)
        intStyle = gcllCollMap("K_" & IntChange)(2)
        
        gcllCollMap.Remove "K_" & IntChange
        
        If IntIndex <> Index Then
            gcllCollMap.Add Array(IntIndex, LngHdl, 0), "K_" & IntChange
            Call zlControl.PicShowFlat(Pic����(IntIndex), 2, , taCenterAlign)
        Else
            '���ǰ����
            If IsIconic(LngHdl) Then
                gcllCollMap.Add Array(IntIndex, LngHdl, 1), "K_" & IntChange
                Call ShowWindow(LngHdl, 9)            '��ԭָ������Ϊԭ��С
                Call zlControl.PicShowFlat(Pic����(Index), -2, , taCenterAlign)
            Else
                If intStyle = 0 Then
                    gcllCollMap.Add Array(IntIndex, LngHdl, 1), "K_" & IntChange
                    Call zlControl.PicShowFlat(Pic����(Index), -2, , taCenterAlign)
                Else
                    gcllCollMap.Add Array(IntIndex, LngHdl, 0), "K_" & IntChange
                    Call CloseWindow(LngHdl)
                    Call zlControl.PicShowFlat(Pic����(Index), 2, , taCenterAlign)
                End If
            End If
            If Not IsIconic(LngHdl) Then Call SetActiveWindow(LngHdl)
        End If
    Next
End Sub

Private Sub Img�˵�_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Pic�˵�_MouseDown Index, Button, Shift, x, y
End Sub

Private Sub Img�˵�_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Pic�˵�_MouseMove Index, Button, Shift, x, y
End Sub

Private Sub Img�˵�ָʾ_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Pic�˵�_MouseDown Index, Button, Shift, x, y
End Sub

Private Sub Img�˵�ָʾ_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Pic�˵�_MouseMove Index, Button, Shift, x, y
End Sub

Private Sub Img��ʼ_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Pic��ʼ_MouseDown Button, Shift, x, y
End Sub

Private Sub Lbl�˵�_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Pic�˵�_MouseDown Index, Button, Shift, x, y
End Sub

Private Sub Lbl�˵�_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Pic�˵�_MouseMove Index, Button, Shift, x, y
End Sub

Private Sub Lbl��ʼ_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Pic��ʼ_MouseDown Button, Shift, x, y
End Sub

Private Sub PicBackDesktop_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    mintLevel = Index + 1
End Sub

Private Sub PicImg_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Pic��ʼ_MouseDown Button, Shift, x, y
End Sub

Private Sub Pic�˵�_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim strCode As String                               'ģ����
    Dim LngFindWindows As Long                          'Ŀ�괰��
    
    If Button = 1 Then
        Call PicToolTipText_Click
        If Img�˵�ָʾ(Index).Visible Then
            '--�����¼��˵�--
            If mblnMenuOpened = False Then
                LoadMenu mintLevel, Index
                mblnMenuOpened = True
            End If
            
        Else
            '--ִ��ģ�鹦��--
            strCode = Img�˵�(Index).Tag
            If mblnPress Then ShowMenu

            Select Case Index
            Case "9000"                             '����
                Shell "hh.exe  zl9start.chm", vbNormalFocus
            Case "9001"                             'ע��
                MnuReg_Click
            Case "9002"                             '�˳�
                MnuExit_Click
            Case "9100"                             '�ֵ����
                mnuToolDictonary_Click
            Case "9101"                             '��Ϣ�շ�
                mnuToolMessage_Click
            Case "9102"
                mnuRepairComponent_Click
            Case "9103"
                mnuToolStyle_Click
            Case "9104"
                mnuToolExcel_Click
            Case "9105"                             '������Ϣ
                mnuToolNotice_Click
            Case Is >= 9300 And Index <= 9500
                '���˺�:Ŀǰ�ݶ�200������
                'С����������,��ʾ���е��ⲿ���ߵ���
                If Index = 9301 Then
                    Dim blnApply As Boolean
                    '���˺�:2007/08/22
                    '�����ⲿ���ߵ�����
                     Call frm��������.ShowEdit(Me, blnApply)
                    If blnApply = False Then Exit Sub
                    Call LoadOutTools(False)
                Else
                    Err = 0: On Error Resume Next
                    Call ExeCuteToolFile(mcllTemp("K" & Index))
                End If
            Case Else                               '���˵�
                    DoEvents
                    With grsMenus
                        .MoveFirst
                        .Find "���='" & strCode & "'"
                        Call ExecuteFunc(.Fields("ϵͳ").Value, IIf(IsNull(.Fields("����").Value), "", .Fields("����").Value), .Fields("ģ��").Value)
                    End With
 
            End Select
        End If
    End If
End Sub

Private Sub Pic�˵�_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim MouseOver As Boolean
    
    '--������������Ϊƽ��
    On Error Resume Next
    Call zlControl.PicShowFlat(PicRollUp(mintLevel - 1), 0, , taCenterAlign)
    Call zlControl.PicShowFlat(PicRollDown(mintLevel - 1), 0, , taCenterAlign)
    
    '--�жϵ�ǰ���λ���Ƿ��ڲ˵���--
    MouseOver = (0 <= x) And (x <= Pic�˵�(Index).Width) And (0 <= y) And (y <= Pic�˵�(Index).Height)
    If MouseOver Then
        If mlngSelectModul = Index Then Exit Sub
        Pic�˵�(Index).BackColor = &H8000000D
        Lbl�˵�(Index).ForeColor = &H80000005
        Call SetCapture(Pic�˵�(Index).hwnd)
        If mlngSelectModul <> Index Then mblnMenuOpened = False
        mlngSelectModul = Index
        If Button <> 88 Then RefreshMenu.Enabled = True
        Call PicBackDesktop_MouseMove(Pic�˵�(Index).Container.Index, Button, Shift, x, y)
        Call ShowToolTipText(Pic�˵�(Index))
    Else
        Pic�˵�(Index).BackColor = &H8000000F
        Lbl�˵�(Index).ForeColor = &H80000008
        Call ReleaseCapture
        mlngSelectModul = 0
        RefreshMenu.Enabled = False
        Call ShowToolTipText(Pic�˵�(Index), False)
    End If
End Sub

Private Sub Pic�ָ�_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If mblnPress Then ShowMenu
End Sub

Private Sub Pic�ָ�_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim DblTotalWidth As Double
    
    If Button = 1 Then
        With Pic�ָ�
            
            DblTotalWidth = 2000
            DblTotalWidth = DblTotalWidth * (Pic����.Count - 1)
            
            If DblTotalWidth > Sbar.Left - Pic�ָ�.Left - x - Pic�ָ�.Width - 1000 Then
                '--�����Ⱥʹ��ڿ����ɵĿռ�--
                DblTotalWidth = ((Sbar.Left - Pic�ָ�.Left - x - Pic�ָ�.Width - 100) / IIf(Pic����.Count - 1 = 0, 1, Pic����.Count - 1)) - 50
                If DblTotalWidth > 2000 Then DblTotalWidth = 2000
            Else
                DblTotalWidth = 2000
            End If
            If DblTotalWidth < 800 Then Exit Sub
            
            If .Left + x > Sbar.Left - 3000 Then .Left = Sbar.Left - 3000: Pic�ָ�_MouseMove Button, Shift, 0, 0: Exit Sub
            If .Left + x < Pic���ù���(Pic���ù���.Count - 1).Left + Pic���ù���(Pic���ù���.Count - 1).Width + 100 Then .Left = Pic���ù���(Pic���ù���.Count - 1).Left + Pic���ù���(Pic���ù���.Count - 1).Width + 100: Pic�ָ�_MouseMove Button, Shift, 0, 0: Exit Sub
            
            .Move .Left + x
        End With
        
        With LineRight
            .X1 = Pic�ָ�.Left - 25
            .X2 = .X1
        End With
        
        With LineLeft
            .X1 = LineRight.X1 - 25
            .X2 = .X1
        End With
        
        Call AdjustPost
    End If
End Sub

Private Sub Pic�ָ�_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then Call AdjustPost
End Sub

Private Sub Pic��ʼ_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If mblnFirst Then Exit Sub
    mintLevel = 1
    If Button = 1 Then Call ShowMenu
End Sub

Public Function ShowMenu()
    Dim IntDelLevel As Integer
    '--���¿�ʼ��ť,����ʾ�����ز˵�,�����������--
    
    mblnPress = mblnPress Xor True
    Call zlControl.PicShowFlat(Pic��ʼ, IIf(mblnPress, -2, 2), , taCenterAlign)
    
    '--��ʾ���������в˵�--
    If mblnPress Then
        Call LoadMenu(-1)
    Else
        Call ShutMenu
        mlngSelectModul = 0
    End If
    Pic��ʶ.Visible = mblnPress
End Function

Private Sub Pic������_Click()
    If mblnPress Then ShowMenu
    Call Find����(-99999999)
End Sub

Private Sub Pic������_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Pic������.ZOrder 0
End Sub

Private Sub RefreshMenu_Timer()
    Dim Index As Integer
    
    On Error Resume Next
    RefreshMenu.Enabled = False
    If mlngSelectModul = 0 Then Exit Sub
    
    If Img�˵�ָʾ(mlngSelectModul).Visible = False Then
        Call ShutMenu(mintLevel)
        Exit Sub
    End If
    
    Index = mlngSelectModul
    Pic�˵�_MouseMove Index, 88, 0, 1, 1
    Pic�˵�_MouseDown Index, 1, 0, 1, 1
End Sub

Private Sub LoadMenu(ByVal IntState As Integer, Optional ByVal IntIndex As Integer = 0)
    Dim DblWidthTmp As Double, lngIndexThis As Long, IntDelLevel As Integer
    Dim strϵͳ As String, str��� As String
    Dim intϵͳ_Cur As Integer, intϵͳ_Max As Integer
    Dim int���_Cur As Integer, int���_Max As Integer
    Dim arrϵͳ, arr���
    Dim blnRight As Boolean
    '--��ж���¼��˵�����װ��ָ���˵����������伶��--
    
    With grsMenus
        If IntIndex = 0 Then
            .Filter = "�ϼ�=0"
        Else
            .Filter = "�ϼ�='" & Img�˵�(IntIndex).Tag & "'"
        End If
        If .EOF And Img�˵�(IntIndex).Tag <> "9003" And Img�˵�(IntIndex).Tag <> "9004" Then Exit Sub
        
        On Error Resume Next
        '--��ж����װ��˵�--
        Call ShutMenu(mintLevel)
        Load PicBackDesktop(mintLevel)
        
        mintLast��� = 0
        mdblMenuHeight = 0
        mdblMenuWidth = 0
        
        If Img�˵�(IntIndex).Tag <> "9003" And Img�˵�(IntIndex).Tag <> "9004" Then             '���ǹ��߲˵�����ʷʹ�ü�¼
            '--װ��ϵͳ�˵�--
            Do While Not .EOF  'ͳ�Ʋ˵��������
                DblWidthTmp = 95 * PicBackDesktop(mintLevel).Font.Size / 9 * LenB(StrConv(!����, vbFromUnicode)) + 800
                If mdblMenuWidth < DblWidthTmp Then mdblMenuWidth = DblWidthTmp
                .MoveNext
            Loop
            mdblMenuWidth = mdblMenuWidth + Img�˵�(0).Width + 400 + Img�˵�ָʾ(0).Width
        
            .MoveFirst
            marrRoll(mintLevel) = ""
            Do While Not .EOF
                lngIndexThis = Img�˵�.Count
                marrRoll(mintLevel) = marrRoll(mintLevel) & IIf(marrRoll(mintLevel) = "", "", ",") & lngIndexThis
                If !���� = 1 And Val(!�Ƿ�ͣ��) = 1 Then
                    If mnuToolShowDisReport.Checked Then
                        Call SetMenuState(lngIndexThis, !���, M_INT_RPTDISABLED, !����, IIf(IsNull(!���), "", !���), 1, IIf(!ģ�� = 0, False, True), IIf(IsNull(!˵��), "", !˵��))
                    End If
                Else
                    Call SetMenuState(lngIndexThis, !���, !ͼ��, !����, IIf(IsNull(!���), "", !���), 1, IIf(!ģ�� = 0, False, True), IIf(IsNull(!˵��), "", !˵��))
                End If
                .MoveNext
            Loop
        Else
            DblWidthTmp = 95 * PicBackDesktop(mintLevel).Font.Size / 9 * LenB(StrConv("С��С��С��С��", vbFromUnicode)) + 800
            mdblMenuWidth = DblWidthTmp
            mdblMenuWidth = DblWidthTmp + Img�˵�(0).Width + 400 + Img�˵�ָʾ(0).Width
            If Img�˵�(IntIndex).Tag = "9003" Then
                '--װ�빤�߲˵�--
                'װ���ֵ����˵�
                If mnuToolDictonary.Visible Then Call SetMenuState(9100, 9100, -5, "�ֵ������", "D", 1, True)
                'װ����Ϣ�շ��˵�
                If mnuToolMessage.Visible Then Call SetMenuState(9101, 9101, -5, "��Ϣ�շ�����", "M", 1, True)
                'װ��������Ϣ���Ĳ˵�
                If mnuToolNotice.Visible Then Call SetMenuState(9105, 9105, -5, "������Ϣ����", "R", 1, True)
                'װ�������ѡ��
                If mnuToolStyle.Visible Then Call SetMenuState(9103, 9103, -5, "ϵͳѡ��", "S", 1, True)
                'װ������EXCEL����
                If mnuToolExcel.Visible Then Call SetMenuState(9104, 9104, -5, "����EXCEL����", "E", 1, True)
                'װ���ⰲװ����
                Call SetMenuState(9102, 9102, -5, "��ⰲװ����", "C", 1, True)
                
                'װ�빤�߲˵�
                Call LoadOutTools(True)
                
            Else
                '--װ����ʷʹ�ü�¼�˵�--
                Call LoadHistory
            End If
        End If
        
        If IntState = -1 Then
            Load FraSplit(1)
            'װ�빤�߲˵�
            Call SetMenuState(9003, 9003, -4, "���湤��", "T", 1, False)
            If Trim(zlDatabase.GetPara("���ʹ��ģ��")) <> "" Then
                'װ����ʷʹ�ü�¼�˵�
                Call SetMenuState(9004, 9004, -4, "��ʷ��¼", "O", 1, False)
            End If
            'װ������˵�
            Call SetMenuState(9000, 9000, -1, "����", "H", 1, True)
            Load FraSplit(2)
            'װ��ע���˵�
            Call SetMenuState(9001, 9001, -2, "ע��", "R", 1, True)
            'װ���˳��˵�
            Call SetMenuState(9002, 9002, -3, "�˳�", "X", 1, True)
        End If
        
        '--����װ��������λ��--
        With PicBackDesktop(mintLevel)
            If mintLevel = 1 Then
                .Left = Pic��ʶ.Left + Pic��ʶ.Width - 20
            Else
                '��.Left���м��
                blnRight = True
                If mintLevel >= 3 Then
                    blnRight = (PicBackDesktop(mintLevel - 2).Left < PicBackDesktop(mintLevel - 1).Left)
                End If
                If blnRight Then
                    .Left = PicBackDesktop(mintLevel - 1).Left + PicBackDesktop(mintLevel - 1).Width - 100
                    If .Left + mdblMenuWidth > Me.ScaleWidth Then
                        .Left = PicBackDesktop(mintLevel - 1).Left - mdblMenuWidth + 100
                        blnRight = blnRight Xor True
                    End If
                Else
                    .Left = PicBackDesktop(mintLevel - 1).Left - mdblMenuWidth + 100
                    If .Left < 0 Then
                        .Left = PicBackDesktop(mintLevel - 1).Left + PicBackDesktop(mintLevel - 1).Width - 100
                        blnRight = blnRight Xor True
                    End If
                End If
            End If
            .Width = mdblMenuWidth + 50
            .Height = mdblMenuHeight + IIf(mintLevel = 1, 150, 90)
            
            If mintLevel = 1 Then
                '--����ǵ�һ������ԭֵ--
                Pic��ʶ.Height = .Height - 50
                Img��ʶ.Height = Pic��ʶ.Height
                .Top = Pic������.Top - .Height
            Else
                '--����¼��˵���ȱʡTopΪ�ϼ��˵��ĸ߶�--
                Dim DblTop
                DblTop = PicBackDesktop(mintLevel - 1).Top + Pic�˵�(IntIndex).Top
                '��.Top���м��
                If Pic������.Top - mdblMenuHeight - 50 < DblTop Then
                    '�������(������ʾ)
                    .Top = Pic������.Top - mdblMenuHeight - 50
                Else
                    .Top = DblTop - 50
                End If
            End If
            .Tag = mintLevel
            .Visible = IIf(.Height > 100, True, False)
            .ZOrder 0
        End With
        
        If mintLevel = 1 Then
            '--����ǵ�һ���˵����������ʶ--
            With Pic��ʶ
                .Height = mdblMenuHeight + 150
                .Top = Pic������.Top - .Height
                .ZOrder 0
            End With
    
            With Img��ʶ
                .Top = Pic��ʶ.Height - .Height
            End With
            
            '�����������
            With Lbl��ʶ
                .AutoSize = True
                .Caption = mstrTitle 'zlProductTitle(GetUnitInfo("������"))
                .AutoSize = False
                .Height = .Width
                .Width = 200
                .Left = Img��ʶ.Left + 80
                .Top = Pic��ʶ.Height - .Height - 100
                .ForeColor = IIf(GetSetting("ZLSOFT", "ע����Ϣ", "Kind", "") = "����", &HFF, &HFFFFFF)
            End With
            
            zlControl.PicShowFlat Pic��ʶ, 2, , taCenterAlign
        End If
        
        Call AdjustMenu(mintLevel)
        Call zlControl.PicShowFlat(PicBackDesktop(mintLevel), 2, , taCenterAlign)
        PicBackDesktop(mintLevel).ZOrder 0
    End With
    grsMenus.Filter = 0
End Sub

Private Function SetMenuState(ByVal lngCurID As Long, ByVal strCode As String, ByVal LngIcon As Long, ByVal StrCaption As String, _
ByVal BytLink As String, ByVal intType As Integer, Optional ByVal BlnEndMenu As Boolean = True, Optional ByVal strNote As String = "")
    '--�������˵����Ŀ�ȵ�����--
    'IntType:�Ƿ�������� 1-����
    
    Load Img�˵�(lngCurID)
    Load Img�˵�ָʾ(lngCurID)
    Load Lbl�˵�(lngCurID)
    Load Pic�˵�(lngCurID)
    Load Lbl���(lngCurID)
        
    With Img�˵�(lngCurID)
        .Left = 100
        .Top = -10
        .Tag = strCode
        Set .Container = Pic�˵�(lngCurID)
        .Visible = True
        .Picture = GetPicDisp(LngIcon, BlnEndMenu) '�˵���ʶ
    End With
    
    With Lbl�˵�(lngCurID)
        .Left = Img�˵�(lngCurID).Left + Img�˵�(lngCurID).Width + 200
        Set .Container = Pic�˵�(lngCurID)
        .Visible = True
        .Caption = StrCaption
        If BytLink <> "" Then .Caption = .Caption & "(" & BytLink & ")"
    End With
    
    With Lbl���(lngCurID)
        .Left = Lbl�˵�(lngCurID).Left + Lbl�˵�(lngCurID).Width - 180
        Set .Container = Pic�˵�(lngCurID)
        .Caption = BytLink
        .Visible = (BytLink <> "")
    End With
    
    If LngIcon < 0 Then
        Select Case lngCurID
        Case 9003
            With FraSplit(1)
                .Visible = True
                Set .Container = PicBackDesktop(mintLevel)
                .Left = 0
                .Top = Pic�˵�(mintLast���).Top + Pic�˵�(mintLast���).Height
                .Width = Pic�˵�(mintLast���).Width
            End With
        Case 9001
            With FraSplit(2)
                .Visible = True
                Set .Container = PicBackDesktop(mintLevel)
                .Left = 0
                .Top = Pic�˵�(mintLast���).Top + Pic�˵�(mintLast���).Height
                .Width = Pic�˵�(mintLast���).Width
            End With
        End Select
    End If
    
    With Pic�˵�(lngCurID)
        .Left = 25
        If mintLast��� = 0 Then
            .Top = 50
        Else
            If LngIcon < 0 Then
                Select Case lngCurID
                Case 9003
                    .Top = FraSplit(1).Top + FraSplit(1).Height
                Case 9001
                    .Top = FraSplit(2).Top + FraSplit(2).Height
                Case Else
                    .Top = Pic�˵�(mintLast���).Top + Pic�˵�(mintLast���).Height
                End Select
            Else
                .Top = Pic�˵�(mintLast���).Top + Pic�˵�(mintLast���).Height
            End If
        End If
        .Tag = mintLevel
        Set .Container = PicBackDesktop(mintLevel)
        .Visible = True
        .Width = mdblMenuWidth - 50
        mdblMenuHeight = mdblMenuHeight + .Height
    End With
    
    With Img�˵�ָʾ(lngCurID)
        .Left = Pic�˵�(lngCurID).Width - .Width - 50
        .Top = (Pic�˵�(lngCurID).Height - .Height) / 2
        Set .Container = Pic�˵�(lngCurID)
        .Visible = BlnEndMenu Xor True
    End With
    
    With Lbl�˵�(lngCurID)
        .Width = Img�˵�ָʾ(lngCurID).Left - .Left - 100
    End With
    
    If mintLast��� <> lngCurID Then mintLast��� = lngCurID
    
    '������ʾ��Ϣ
    Call SetToolTipText(Img�˵�, lngCurID, 0, strNote)
    Call SetToolTipText(Lbl�˵�, lngCurID, 0, strNote)
    Call SetToolTipText(Lbl���, lngCurID, 0, strNote)
    Call SetToolTipText(Pic�˵�, lngCurID, 0, strNote)
    Call SetToolTipText(Img�˵�ָʾ, lngCurID, 0, strNote)
End Function

Public Sub ShutMenu(Optional ByVal Level As Integer = 0)
    Dim ObjShut As Object, LngUnloadObjs As Long
    Dim IntDelLevel As Integer
    
    On Error Resume Next
    LngUnloadObjs = 0
    RefreshMenu.Enabled = False
    For Each ObjShut In Me.Controls 'ɾ���ؼ�
        Err = 0
        Select Case TypeName(ObjShut)
            Case "Image", "Label", "PictureBox"
                If ObjShut.Index <> 0 And Not (ObjShut.Name Like "*���ù���") Then
                    If Err = 0 Then
                        If Val(ObjShut.Container.Tag) >= Level Then
                            Unload ObjShut
                        End If
                    End If
                End If
            Case "Frame"
                If Level <= 1 Then
                    If ObjShut.Index <> 0 And Not (ObjShut.Name Like "*���ù���") Then
                        If Err = 0 Then
                            If Val(ObjShut.Container.Tag) >= Level Then
                                Unload ObjShut
                            End If
                        End If
                    End If
                End If
        End Select
    Next

    For Each ObjShut In Me.Controls '�������ϴ��޷�ɾ�����ٴ�ִ��
        Err = 0
        Select Case TypeName(ObjShut)
            Case "PictureBox"
                If ObjShut.Index <> 0 And Not (ObjShut.Name Like "*���ù���") Then
                    If Err = 0 Then
                        If Val(ObjShut.Tag) >= Level Then
                            ObjShut.Visible = False
                            Unload ObjShut
                        End If
                    End If
                End If
        End Select
    Next
    
    For IntDelLevel = mintLevel To PicBackDesktop.Count - 1
        PicBackDesktop(IntDelLevel).Visible = False
        Unload PicBackDesktop(IntDelLevel)
    Next
End Sub

Private Sub FindMenu(ByVal IntKey As Integer)
    Dim ObjShut As Control
    
    On Error Resume Next
    '--�ڵ�ǰ�����в���ָ����ݼ��Ĳ˵�--
    For Each ObjShut In Me.Controls '�������ϴ��޷�ɾ�����ٴ�ִ��
        Err = 0
        If TypeName(ObjShut) = "PictureBox" Then
            If ObjShut.Index <> 0 Then
                If Err = 0 Then
                    If Val(ObjShut.Container.Index) = mintLevel Then
                        If Err = 0 Then
                            If Lbl���(ObjShut.Index).Caption = UCase(Chr(IntKey)) Then
                                '���ò˵�Ϊѡ��״̬��ִ��
                                RefreshMenu.Enabled = False
                                Pic�˵�_MouseMove ObjShut.Index, 88, 0, 0, 0
                                Pic�˵�_MouseDown ObjShut.Index, 1, 0, 0, 0
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next
    
End Sub

Public Sub OpenWindow(ByVal strCode As String, ByVal StrCaption As String)
    Dim FrmMainOpen As FrmExecute, FrmTest As Form
    Dim IntChange As Integer, LngHdl As Long
    Dim strBuffer As String * 256, IntBuffer As Integer
    
    IntBuffer = 255
    '����������ģʽ
    For IntChange = 1 To gcllCollMap.Count
        LngHdl = gcllCollMap("K_" & IntChange)(1)
        Call GetWindowText(LngHdl, strBuffer, IntBuffer)
        If StrCaption = Trim(Replace(strBuffer, Chr(0), "")) Then Exit For
    Next
    
    If StrCaption = Trim(Replace(strBuffer, Chr(0), "")) Then
        '����������ģʽ
        For IntChange = 1 To gcllCollMap.Count
            If gcllCollMap("K_" & IntChange)(1) <> LngHdl Then
                Call zlControl.PicShowFlat(Pic����(gcllCollMap("K_" & IntChange)(0)), 2, , taCenterAlign)
            Else
                Call zlControl.PicShowFlat(Pic����(gcllCollMap("K_" & IntChange)(0)), -2, , taCenterAlign)
            End If
        Next
        
        '���ǰ����
        If IsIconic(LngHdl) Then Call ShowWindow(LngHdl, 9)            '��ԭָ������Ϊԭ��С
        Call SetActiveWindow(LngHdl)
    Else
        Set FrmMainOpen = New FrmExecute
        With FrmMainOpen
            .�������� = StrCaption
            .Str��� = strCode
            Set .mrsMenus = grsMenus.Clone
            Call SetParent(.hwnd, gLngFormID)
            .Show 0
        End With
    End If
End Sub

Public Function Show����(ByVal ChildObj As Object, Optional ByVal strCode As String = "", Optional ByVal StrCaption As String = "")
    Dim LngIcon As Long
    Dim IntIndex As Integer
    Dim DblTotalWidth As Double
    
    If grsMenus.State = 0 Then Exit Function
    If grsMenus.EOF Then Exit Function
    With grsMenus
        .MoveFirst
        .Find "����='" & ChildObj.Caption & "'"
        If .EOF Then
            .MoveFirst
            
            '������ڹ���
            If ChildObj.Caption = "" Then Exit Function
            If InStr(1, "�Զ��屨�����,�ֵ������,��Ϣ�շ�����", ChildObj.Caption) <> 0 Then GoTo Normal
            Exit Function
        End If
    End With
    
Normal:                                                     '��������
    Set mFrmChildObj = ChildObj
    Call SetParent(ChildObj.hwnd, Me.hwnd)
    
    If StrCaption = "" Then StrCaption = mFrmChildObj.Caption
    If gcllCollMap.Count = 0 Then
        IntIndex = 1
    Else
        IntIndex = gcllCollMap(gcllCollMap.Count)(0) + 1
    End If
    Load Pic����(IntIndex)
    Load Lbl����(IntIndex)
    Load Img����(IntIndex)
    
    With Pic����(IntIndex)
        .Tag = -IntIndex
        .Visible = True
'        Set .Container = Pic������
    End With
    With Lbl����(IntIndex)
        .Tag = -IntIndex
        .Caption = StrCaption
        Set .Container = Pic����(IntIndex)
        .Visible = True
    End With
    With Img����(IntIndex)
        .Width = 240
        .Height = 240
        .Tag = -IntIndex
        .Picture = ChildObj.Icon
        Set .Container = Pic����(IntIndex)
        .Visible = True
    End With
    '����˵��          ����       ������       �Ƿ񼤻�
    gcllCollMap.Add Array(IntIndex, mFrmChildObj.hwnd, 1), "K_" & gcllCollMap.Count + 1
    Call AdjustPost
    
    If grsMenus!ģ�� <> 0 Then Call AddHistory(grsMenus!ϵͳ & "," & grsMenus!ģ��)
End Function

Public Sub Shut����(ByVal ObjFrm As Object)
    Dim IntChange As Integer, IntDelete As Integer
    On Error Resume Next
    
    '�ҵ�������
    IntDelete = 0
    For IntChange = 1 To gcllCollMap.Count
        If gcllCollMap(IntChange)(1) = ObjFrm.hwnd Then
            'ɾ��������
            IntDelete = IntChange
            Call Find����(gcllCollMap(IntChange)(0), True)
            Exit For
        End If
    Next
    
    If IntDelete = 0 Then Exit Sub
    '�����޸ĺ������
    For IntChange = IntDelete To gcllCollMap.Count - 1
        gcllCollMap.Remove "K_" & IntChange
        gcllCollMap.Add gcllCollMap("K_" & IntChange + 1), "K_" & IntChange
    Next
    gcllCollMap.Remove "K_" & gcllCollMap.Count
    
    Call AdjustPost
End Sub

Public Sub Find����(ByVal Index As Long, Optional BlnDel As Boolean = False, Optional BlnState As Boolean = False)
    Dim ObjShut As Object
    
    On Error Resume Next
    
    Call AdjustPost
    For Each ObjShut In Me.Controls 'ɾ���ؼ�
        Err = 0
        Select Case TypeName(ObjShut)
            Case "Image", "Label", "PictureBox"
                If ObjShut.Index <> 0 Then
                    If Err = 0 Then
                        If ObjShut.Tag < 0 Then
                            If BlnDel And ObjShut.Tag = -Index Then
                                Unload ObjShut
                                Call AdjustPost
                            End If
                        End If
                    End If
                End If
        End Select
    Next
    For Each ObjShut In Me.Controls 'ɾ���ؼ�
        Err = 0
        Select Case TypeName(ObjShut)
            Case "PictureBox"
                If ObjShut.Index <> 0 Then
                    If Err = 0 Then
                        If ObjShut.Tag < 0 Then
                            If BlnDel And ObjShut.Tag = -Index Then
                                Unload ObjShut
                                Call AdjustPost
                            ElseIf BlnState And ObjShut.Tag = -Index Then
                                Call zlControl.PicShowFlat(ObjShut, -2, , taCenterAlign)
                            Else
                                Call zlControl.PicShowFlat(ObjShut, 2, , taCenterAlign)
                            End If
                        End If
                    End If
                End If
        End Select
    Next
    
    If Index = -99999999 Then
        Dim IntChange As Integer
        Dim IntIndex As Integer, LngThisHdl As Long, intStyle As Integer
        mstrLastSelectCaption = ""
        mlngLastSelectIndex = 0
        
        '����������ģʽ
        For IntChange = 1 To gcllCollMap.Count
            IntIndex = gcllCollMap("K_" & IntChange)(0)
            LngThisHdl = gcllCollMap("K_" & IntChange)(1)
            intStyle = gcllCollMap("K_" & IntChange)(2)
            
            gcllCollMap.Remove "K_" & IntChange
            
            gcllCollMap.Add Array(IntIndex, LngThisHdl, 0), "K_" & IntChange
            Call zlControl.PicShowFlat(Pic����(IntIndex), 2, , taCenterAlign)
        Next
    End If
End Sub

Public Sub AdjustPost()
    Dim IntReAdjust As Integer, DblTotalWidth As Double
    Dim DblPicResizeWidth As Double, DblPicResizeLeft As Double
    Dim DblLblResizeLeft As Double, DblLblResizeWidth As Double
    Dim DblPicResizeTop As Double, DblPicResizeHeight As Double
    On Error Resume Next
    
    DblTotalWidth = 2000
    DblTotalWidth = DblTotalWidth * (Pic����.Count - 1)
    
    If DblTotalWidth > Sbar.Left - Pic�ָ�.Left - Pic�ָ�.Width - 1000 Then
        '--�����Ⱥʹ��ڿ����ɵĿռ�--
        DblTotalWidth = ((Sbar.Left - Pic�ָ�.Left - Pic�ָ�.Width - 100) / IIf(Pic����.Count - 1 = 0, 1, Pic����.Count - 1)) - 50
        If DblTotalWidth > 2000 Then DblTotalWidth = 2000
    Else
        DblTotalWidth = 2000
    End If
    
    DblPicResizeTop = Pic����(0).Top / Screen.TwipsPerPixelX
    DblPicResizeHeight = Pic����(0).Height / Screen.TwipsPerPixelX
    
    DblPicResizeLeft = (Pic�ָ�.Left + Pic�ָ�.Width + 50) / Screen.TwipsPerPixelX
    DblPicResizeWidth = (DblTotalWidth) / Screen.TwipsPerPixelX
    Lbl����(gcllCollMap("K_1")(0)).Width = DblTotalWidth - Lbl����(gcllCollMap("K_1")(0)).Left - 100
    Call MoveWindow(Pic����(gcllCollMap("K_1")(0)).hwnd, DblPicResizeLeft, DblPicResizeTop, DblPicResizeWidth, DblPicResizeHeight, 0)
    Call zlControl.PicShowFlat(Pic����(gcllCollMap("K_1")(0)), 2, , taCenterAlign)
    
    For IntReAdjust = 3 To Pic����.Count
        DblPicResizeLeft = (Pic����(gcllCollMap("K_" & IntReAdjust - 2)(0)).Left + Pic����(gcllCollMap("K_" & IntReAdjust - 2)(0)).Width + 50) / Screen.TwipsPerPixelX
        DblPicResizeWidth = DblTotalWidth / Screen.TwipsPerPixelX
        Lbl����(gcllCollMap("K_" & IntReAdjust - 1)(0)).Width = DblTotalWidth - Lbl����(gcllCollMap("K_" & IntReAdjust - 1)(0)).Left - 100
        Call MoveWindow(Pic����(gcllCollMap("K_" & IntReAdjust - 1)(0)).hwnd, DblPicResizeLeft, DblPicResizeTop, DblPicResizeWidth, DblPicResizeHeight, 0)
        Call zlControl.PicShowFlat(Pic����(gcllCollMap("K_" & IntReAdjust - 1)(0)), 2, , taCenterAlign)
    Next
    Pic������.Refresh
    
    For IntReAdjust = 1 To gcllCollMap.Count
        gcllCollMap("K_" & IntReAdjust)(2) = 0
    Next
End Sub

Private Function SetToolTipText(ByVal ObjCon As Object, ByVal NewIndex As Long, ByVal intStyle As Integer, ByVal strNote As String)
    '������:����
    '��������:2000-11-21
    '����:���ӿؼ��İ���˵��

    Select Case intStyle
    Case 0
        ObjCon(NewIndex).ToolTipText = strNote
    Case -1
        ObjCon(NewIndex).ToolTipText = "��ȡ�������ϵͳ�����߰�����"
    Case -2
        ObjCon(NewIndex).ToolTipText = "�˳��������ϵͳ��"
    End Select
End Function

Private Sub SetZorder_Timer()
    Dim LngHdl As Long
    Dim FrmTest As Form
    Dim StrCaption As String * 255
    Dim StrTran As String
    Dim LngCount As Long
    Dim IntChange As Integer
    '���ڱ��漯���е�����
    Dim IntIndex As Integer, LngThisHdl As Long, intStyle As Integer
    
    Pic������.ZOrder
    PicToolTipText.ZOrder

    '��ȡ��ǰ�����,�������������,�򽫲˵��ر�
    LngHdl = GetActiveWindow()
    If LngHdl <> Me.hwnd Then
        '�رղ˵�
        mblnPress = True
        mblnAdjustPost = False
        Call ShowMenu
        
        '���Ҹ��Ӵ����Ӧ��������
        On Error Resume Next
        LngCount = 254
        Call GetWindowText(LngHdl, StrCaption, LngCount)
        StrTran = Trim(Replace(StrCaption, Chr(0), ""))
        If StrTran <> mstrLastSelectCaption Then
            mstrLastSelectCaption = StrTran

            '����������ģʽ
            For IntChange = 1 To gcllCollMap.Count
                IntIndex = gcllCollMap("K_" & IntChange)(0)
                LngThisHdl = gcllCollMap("K_" & IntChange)(1)
                intStyle = gcllCollMap("K_" & IntChange)(2)
                
                gcllCollMap.Remove "K_" & IntChange
                
                If LngThisHdl <> LngHdl Then
                    gcllCollMap.Add Array(IntIndex, LngThisHdl, 0), "K_" & IntChange
                    Call zlControl.PicShowFlat(Pic����(IntIndex), 2, , taCenterAlign)
                Else
                    gcllCollMap.Add Array(IntIndex, LngThisHdl, 1), "K_" & IntChange
                    Call zlControl.PicShowFlat(Pic����(IntIndex), -2, , taCenterAlign)
                End If
            Next

            '���ǰ����
            If IsWindowVisible(LngHdl) <> 0 Then
'                If IsIconic(LngHdl) Then Call ShowWindow(LngHdl, 9)                       '��ԭָ������Ϊԭ��С
'                Call SetActiveWindow(LngHdl)
                If Not IsIconic(LngHdl) Then Call SetActiveWindow(LngHdl)
            End If
        End If
    Else
        If mblnAdjustPost = False Then
            Call AdjustPost
            mblnAdjustPost = True
        End If
    End If
End Sub

Private Sub InitEvn()
    Dim StrPicPath As String
    Dim LngColor As Long
    
    '--��ʼװ��ͼ��,ͼƬ--
    Img��ʶ.Picture = LoadResPicture(101, 0) '�˵���ʶ
    StrPicPath = zlDatabase.GetPara("zlWinBackPic")
    
    If Trim(StrPicPath) <> "" Then
        '�û�ѡ��ͼƬ,�����Ƿ�����
        On Error Resume Next
        Err = 0
        
        Img��ʶ.Picture = LoadPicture(StrPicPath)
        mblnShow = (Err = 0)
        
        If mblnShow Then PicBackBitmap.Picture = LoadPicture(StrPicPath)
    Else
        mblnShow = True
    End If
    Call PicBackBitmap.PaintPicture(PicBackBitmap.Picture, 0, 0, PicBackBitmap.Width, PicBackBitmap.Height, _
                    0, 0, PicBackBitmap.Picture.Width * 0.57, PicBackBitmap.Picture.Height * 0.57)
    LvwList.Picture = PicBackBitmap.Image
    '�ָ�ԭ�����õ�ͼƬ
    Img��ʶ.Picture = LoadResPicture(101, 0) '�˵���ʶ
    
    'ȡ����ɫ
    LngColor = Val(zlDatabase.GetPara("zlWinFontColor"))
    If LngColor <> -1 Then
        LvwList.ForeColor = LngColor
    End If
End Sub

Private Sub TimePass_Timer()
    Call Form_KeyDown(vbKeyF12, 7)  '�����̬����
End Sub

Private Sub AdjustMenu(ByVal intLevel As Integer)
    Dim lngMin As Long, lngMax As Long, arrIndex
    Dim lngHeight As Long, lngRollHeight As Long, blnVisible As Boolean '�˵���ĸ߶�,�ǹ̶��˵�����ܸ߶�,�Ƿ���ʾ�ò˵�
    Dim lngEnd As Long '���һ����ʾ�Ĳ˵���
    '�����˵������ڳ�����ʾ�Ĳ˵����ֹ��ʾ��������ӹ�����
    
    If marrRoll(intLevel) = "" Then Exit Sub
    If PicBackDesktop(intLevel).Top >= 0 Then Exit Sub
    
    arrIndex = Split(marrRoll(intLevel), ",")
    lngMax = UBound(arrIndex)
    lngHeight = (-1 * PicBackDesktop(intLevel).Top) \ Pic�˵�(0).Height + 2
    lngHeight = lngHeight * Pic�˵�(0).Height
    
    '��ȡ�ǹ̶��˵�����ܸ߶�
    lngEnd = 0
    lngRollHeight = 0
    For lngMin = 0 To lngMax
        lngRollHeight = lngRollHeight + Pic�˵�(arrIndex(lngMin)).Height
    Next
    lngRollHeight = lngRollHeight - lngHeight + PicRollUp(0).Height
    
    '�����˵�����ʶ���߶�
    If intLevel = 1 Then
        Pic��ʶ.Height = Pic��ʶ.Height - lngHeight + PicRollUp(0).Height
        Pic��ʶ.Top = Pic������.Top - Pic��ʶ.Height
        Pic��ʶ.Visible = True
        
        With Img��ʶ
            .Top = 50
        End With
        
        '�����������
        With Lbl��ʶ
            .AutoSize = True
            .Caption = mstrTitle ' zlProductTitle(GetUnitInfo("������"))
            .AutoSize = False
            .Height = .Width
            .Width = 200
            .Left = Img��ʶ.Left + 80
            .Top = Pic��ʶ.Height - .Height - 100
            .ForeColor = IIf(GetSetting("ZLSOFT", "ע����Ϣ", "Kind", "") = "����", &HFF, &HFFFFFF)
        End With
    End If
    PicBackDesktop(intLevel).Height = PicBackDesktop(intLevel).Height - lngHeight + PicRollUp(0).Height
    PicBackDesktop(intLevel).Top = Pic������.Top - PicBackDesktop(intLevel).Height
    PicBackDesktop(intLevel).ZOrder 0
    
    Call AddRollMenu(intLevel)
    
    '�����ǹ̶��˵���
    For lngMin = 0 To lngMax
        blnVisible = True
        If lngMin <> 0 Then
            blnVisible = (Pic�˵�(arrIndex(lngMin - 1)).Top + Pic�˵�(arrIndex(lngMin - 1)).Height < lngRollHeight - PicRollUp(0).Height)
        End If
        Pic�˵�(arrIndex(lngMin)).Visible = blnVisible
        If Not blnVisible And lngEnd = 0 Then lngEnd = lngMin - 1
    Next
    PicRollDown(intLevel).Top = Pic�˵�(arrIndex(lngEnd)).Top + Pic�˵�(arrIndex(lngEnd)).Height
    
    If intLevel <> 1 Then Exit Sub
    
    '����ϵͳ�̶��˵���
    Dim blnHistory As Boolean
    FraSplit(1).Top = PicRollDown(intLevel).Top + PicRollDown(intLevel).Height
    Pic�˵�(9003).Top = FraSplit(1).Top + FraSplit(1).Height
    blnHistory = Trim(zlDatabase.GetPara("���ʹ��ģ��")) <> ""
    If blnHistory Then Pic�˵�(9004).Top = Pic�˵�(9003).Top + Pic�˵�(9003).Height
    Pic�˵�(9000).Top = Pic�˵�(IIf(blnHistory, 9004, 9003)).Top + Pic�˵�(IIf(blnHistory, 9004, 9003)).Height
    FraSplit(2).Top = Pic�˵�(9000).Top + Pic�˵�(9000).Height
    Pic�˵�(9001).Top = FraSplit(2).Top + FraSplit(2).Height
    Pic�˵�(9002).Top = Pic�˵�(9001).Top + Pic�˵�(9001).Height
End Sub

Private Sub AddRollMenu(ByVal intLevel As Integer)
    '���ӹ�����
    
    Load PicRollDown(intLevel)
    Load PicRollUp(intLevel)
    Load ImgRollDown(intLevel)
    Load ImgRollUp(intLevel)
    
    With PicRollUp(intLevel)
        Set .Container = PicBackDesktop(intLevel)
        .Left = 50
        .Top = 50
        .Width = PicBackDesktop(intLevel).Width - 80
        .Tag = intLevel
    End With
    With PicRollDown(intLevel)
        Set .Container = PicBackDesktop(intLevel)
        .Left = 50
        .Width = PicBackDesktop(intLevel).Width - 80
        .Visible = True
        .Tag = intLevel
    End With
    With ImgRollDown(intLevel)
        Set .Container = PicRollDown(intLevel)
        .Left = PicRollDown(intLevel).Width / 2 - .Width
        .Visible = True
        .Tag = intLevel
    End With
    With ImgRollUp(intLevel)
        Set .Container = PicRollUp(intLevel)
        .Left = PicRollUp(intLevel).Width / 2 - .Width
        .Visible = True
        .Tag = intLevel
    End With
End Sub

Private Sub RollUpMenu(ByVal intLevel As Integer, Optional ByVal intWay As Integer = 1)
    Dim lngMin As Long, lngMax As Long, lngCur As Long
    Dim lngStart As Long, lngEnd As Long, blnVisible As Boolean
    Dim arrIndex
    '�����˵�
    'intWay-��������:1-����;2-����
    'lngStart����VisibleΪ��ĵ�һ���˵�������
    'lngEnd����VisibleΪ������һ���˵�������
    
    If marrRoll(intLevel) = "" Then Exit Sub
    
    arrIndex = Split(marrRoll(intLevel), ",")
    lngMax = UBound(arrIndex)
    blnVisible = False
    
    '����intWay���ҵ���һ���˵�����һ���˵���
    lngCur = 0
    For lngMin = 0 To lngMax
        If Pic�˵�(arrIndex(lngMin)).Visible Then
            If Not blnVisible Then
                lngStart = lngMin
                blnVisible = True
            End If
        Else
            If blnVisible Then
                lngEnd = lngCur
                Exit For
            End If
        End If
        lngCur = lngMin
    Next
    If lngEnd = 0 Then lngEnd = lngMax
    
    '���Ų˵�
    If (lngStart = 0 And intWay <> 1) Or (lngEnd = lngMax And intWay = 1) Then Exit Sub
    lngStart = lngStart + IIf(intWay = 1, 1, -1)
    lngEnd = lngEnd + IIf(intWay = 1, 1, -1)
    PicRollUp(intLevel).Visible = Not (lngStart = 0)
    PicRollDown(intLevel).Visible = Not (lngEnd = lngMax)
    
    For lngMin = 0 To lngMax
        blnVisible = (lngMin >= lngStart And lngMin <= lngEnd)
        Pic�˵�(arrIndex(lngMin)).Visible = blnVisible
    Next
    For lngMin = lngStart To lngEnd
        If lngMin = lngStart Then
            Pic�˵�(arrIndex(lngMin)).Top = IIf(PicRollUp(intLevel).Visible, PicRollUp(intLevel).Top + PicRollUp(intLevel).Height, 50)
        Else
            Pic�˵�(arrIndex(lngMin)).Top = Pic�˵�(arrIndex(lngMin - 1)).Top + Pic�˵�(arrIndex(lngMin - 1)).Height
        End If
    Next
    
    Call zlControl.PicShowFlat(PicRollUp(intLevel), 0, , taCenterAlign)
    Call zlControl.PicShowFlat(PicRollDown(intLevel), 0, , taCenterAlign)
    If intWay <> 1 Then
        Call zlControl.PicShowFlat(PicRollUp(intLevel), -1, , taCenterAlign)
    Else
        Call zlControl.PicShowFlat(PicRollDown(intLevel), -1, , taCenterAlign)
    End If
End Sub

Private Sub LoadHistory()
    Dim strϵͳ As String, str��� As String
    Dim arrϵͳ As Variant, arr��� As Variant
    Dim intϵͳ_Cur As Integer, int���_Cur As Integer
    Dim intϵͳ_Max As Integer, int���_Max As Integer
    Dim strValue As String
    
    '����ʷ��¼װ��˵�
    strValue = zlDatabase.GetPara("���ʹ��ģ��")
    If UBound(Split(strValue, "|")) < 1 Then Exit Sub
    strϵͳ = Trim(Split(strValue, "|")(0))
    str��� = Trim(Split(strValue, "|")(1))
    If strϵͳ = "" Or str��� = "" Then Exit Sub
    
    arrϵͳ = Split(strϵͳ, ",")
    arr��� = Split(str���, ",")
    intϵͳ_Max = UBound(arrϵͳ)
    int���_Max = UBound(arr���)
    If intϵͳ_Max > 8 Then intϵͳ_Max = 8 '���˸���ʷ��¼
    
    For intϵͳ_Cur = 0 To intϵͳ_Max
        int���_Cur = intϵͳ_Cur
        If int���_Cur > int���_Max Then Exit For
        
        With grsMenus
            .Filter = "ϵͳ=" & IIf(arrϵͳ(intϵͳ_Cur) = "", 0, arrϵͳ(intϵͳ_Cur)) & " And ģ��=" & arr���(int���_Cur)
            If .RecordCount <> 0 Then
                '����ȱʡֵ
                Call SetMenuState(9200 + intϵͳ_Cur, !���, !ͼ��, !����, intϵͳ_Cur + 1, 1, IIf(!ģ�� = 0, False, True), IIf(IsNull(!˵��), "", !˵��))
            End If
            .Filter = 0
        End With
    Next
End Sub

Private Sub LoadUsual()
    Dim strϵͳ As String, str��� As String, strͼ�� As String, str���� As String
    Dim arrϵͳ As Variant, arr��� As Variant, arrͼ�� As Variant, arr���� As Variant
    Dim intϵͳ_Cur As Integer, int���_Cur As Integer, intͼ��_Cur As Integer, int����_Cur As Integer
    Dim intϵͳ_Max As Integer, int���_Max As Integer, intͼ��_Max As Integer, int����_Max As Integer
    Dim strValue As String
    
    '���ӳ��ù���
    strValue = zlDatabase.GetPara("���ù���ģ��")
    If UBound(Split(strValue, "|")) < 3 Then Exit Sub
    strϵͳ = Trim(Split(strValue, "|")(0))
    str��� = Trim(Split(strValue, "|")(1))
    strͼ�� = Trim(Split(strValue, "|")(2))
    str���� = Trim(Split(strValue, "|")(3))
    If strϵͳ = "" Or str��� = "" Then Exit Sub
    
    arrϵͳ = Split(strϵͳ, ",")
    arr��� = Split(str���, ",")
    arrͼ�� = Split(strͼ��, ",")
    arr���� = Split(str����, ",")
    intϵͳ_Max = UBound(arrϵͳ)
    int���_Max = UBound(arr���)
    intͼ��_Max = UBound(arrͼ��)
    int����_Max = UBound(arr����)
    
    For intϵͳ_Cur = 0 To intϵͳ_Max
        int���_Cur = intϵͳ_Cur
        intͼ��_Cur = intϵͳ_Cur
        int����_Cur = intϵͳ_Cur
        If int���_Cur > int���_Max Then Exit For
        
        With grsMenus
            .Filter = "ϵͳ=" & arrϵͳ(intϵͳ_Cur) & " And ģ��=" & arr���(int���_Cur)
            If .RecordCount <> 0 Then
                '����ȱʡֵ
                strϵͳ = !ϵͳ
                str��� = !ģ��
                If int����_Cur <= int����_Max Then
                    str���� = arr����(int����_Cur)
                Else
                    str���� = !����
                End If
                If intͼ��_Cur <= intͼ��_Max Then
                    strͼ�� = arrͼ��(intͼ��_Cur)
                Else
                    strͼ�� = !ͼ��
                End If
                Call AddUsualModul(strϵͳ & "��" & str��� & "��" & str���� & "��" & strͼ��)
                Call Pic�ָ�_MouseMove(1, 0, 0, 0)
                If Pic�ָ�.Left + Pic���ù���(0).Width >= Sbar.Left - 3000 Then Exit Sub
            End If
            .Filter = 0
        End With
    Next
End Sub

Private Sub AddUsualModul(ByVal strModul As String)
    Dim lngAdd As Long
    Dim lngϵͳ As Long, lngģ�� As Long, lngͼ�� As Long, str���� As String
    '����ָ���ĳ��ù��ܿؼ�
    
    '�ȷֽ����
    lngϵͳ = Split(strModul, "��")(0)
    lngģ�� = Split(strModul, "��")(1)
    str���� = Split(strModul, "��")(2)
    lngͼ�� = Split(strModul, "��")(3)
    
    '���ؿؼ�
    lngAdd = Pic���ù���.Count
    Load Pic���ù���(lngAdd)
    Load Img���ù���(lngAdd)
    With Pic���ù���(lngAdd)
        Set .Container = Pic������
        .Left = Pic���ù���(lngAdd - 1).Left + Pic���ù���(lngAdd - 1).Width
        .Tag = strModul
        .Visible = True
    End With
    With Img���ù���(lngAdd)
        Set .Container = Pic���ù���(lngAdd)
        .Picture = GetPicDisp(lngͼ��)
        .Visible = True
    End With
    Call SetToolTipText(Pic���ù���, lngAdd, 0, str����)
    Call SetToolTipText(Img���ù���, lngAdd, 0, str����)
End Sub

Private Sub ShutUsual()
    'ɾ�����г��ù���
    Dim ObjShut As Object, LngUnloadObjs As Long
    
    On Error Resume Next
    LngUnloadObjs = 0
    For Each ObjShut In Me.Controls 'ɾ���ؼ�
        Err = 0
        Select Case TypeName(ObjShut)
            Case "Image", "PictureBox"
                If (ObjShut.Name Like "*���ù���") Then
                    If Err = 0 Then
                        Unload ObjShut
                    End If
                End If
        End Select
    Next

    For Each ObjShut In Me.Controls '�������ϴ��޷�ɾ�����ٴ�ִ��
        Err = 0
        Select Case TypeName(ObjShut)
            Case "PictureBox"
                If (ObjShut.Name Like "*���ù���") Then
                    If Err = 0 Then
                        ObjShut.Visible = False
                        Unload ObjShut
                    End If
                End If
        End Select
    Next
End Sub

Private Sub ShowToolTipText(ByVal ObjCon As Object, Optional ByVal blnVisible As Boolean = True)
    Static ObjCon_Last As Object
    '1���ӳٺ���ʾ��4���ӳٺ���ʧ
    'blnVisible:ǿ�ƹر�
    On Error Resume Next
    
    If Not blnVisible Then
        TimeToolTipText.Enabled = False
        PicToolTipText.Visible = False
        Exit Sub
    End If
    
    '�������ͬ�������¼�ʱ��1�룩
    If Not ObjCon_Last Is Nothing Then
        If ObjCon_Last.Name = ObjCon.Name And ObjCon_Last.Index = ObjCon.Index Then
            '�п����ϴ�ѡ����ǲ˵������˵������Ѿ��رգ������ж��Ƿ�������
            If Err = 0 Then
                TimeToolTipText.Interval = 1000
                TimeToolTipText.Enabled = True
                Exit Sub
            End If
        End If
    End If
    Set ObjCon_Last = ObjCon
    
    PicToolTipText.Visible = False
    With TimeToolTipText
        .Enabled = False
        .Interval = 1000
        .Enabled = True
        .Tag = ObjCon.Container.Left + ObjCon.Left & "��" & ObjCon.Container.Top + ObjCon.Top & "��" & ObjCon.ToolTipText
        Exit Sub
    End With
End Sub

Private Sub TimeToolTipText_Timer()
    Dim BlnShowToolTipText As Boolean
    'BlnShowToolTipText:��-��ʧ;��-��ʾ
    
    With TimeToolTipText
        BlnShowToolTipText = (.Interval = 1000)
        .Enabled = False
        If BlnShowToolTipText Then
            .Interval = 4000
            .Enabled = True
        End If
    End With
    If Trim(TimeToolTipText.Tag) = "" Then Exit Sub
    LblToolTipText.Caption = Split(TimeToolTipText.Tag, "��")(2)
    If Trim(LblToolTipText.Caption) = "" Then Exit Sub
    With PicToolTipText
        .Visible = BlnShowToolTipText
        .Left = Split(TimeToolTipText.Tag, "��")(0) + 250
        .Top = Split(TimeToolTipText.Tag, "��")(1) + 500
        .Width = LblToolTipText.Width + 80
        .ZOrder
        '��������߽磬��ȡ
        If .Left < 0 Then
            .Left = 0
        End If
        If .Left > Me.Width Then
            .Left = Me.Width - .Width
        End If
        If .Top < 0 Then
            .Top = 0
        End If
        If .Top > Me.Height Then
            .Top = Split(TimeToolTipText.Tag, "��")(1) - .Height
        End If
    End With
End Sub

Private Sub CheckTools()
    Dim blnSplit As Boolean         '�Ƿ���ʾ�ָ���
    '��Ϣ�շ���EXCEL�����Ȩ�޿��ƣ�
    '1�������Ȩ���к��д˹���
    '2��������û�ӵ�д�Ȩ��
    '3����ʾ����������
    '��������ģ����жϸ��û��Ƿ�ӵ�д�Ȩ��
    
    '���߶�Ӧ˵��
    '��ӡ��Ԥ�������EXCEL  ,10,'���������嵥','����'
    'mnuToolDictonary       ,11,'�ֵ������','����'
    'mnuToolMessage         ,12,'��Ϣ�շ�����','����,������Ϣ'
    'mnuTooleSelect         ,13,'ϵͳѡ������','����'
    'mnuToolExcel           ,14,'EXCEL������','����,������ɾ,�������,����ϵͳ'
    'mnuToolUp              ,15,'���ز����ϴ�' ,'����'
    
    Dim intGrant As Integer
    
    '���������嵥
    'Excel������
    mnuToolExcel.Visible = False
    '��Ϣ�շ�����
    mnuToolMessage.Visible = False
    mnuToolNotify.Visible = False
    'ϵͳѡ������
    mnuToolStyle.Visible = False
    '�ֵ������
    mnuToolDictonary.Visible = False
    
    '��Ȼ,�ָ���һ����Ҫ��ֹ��,ֻҪ��������һ�����ܣ��ֵ������Ϣ�շ���EXCEL�����ϵͳѡ�������Ҫ��ʾ�ָ���
    blnSplit = False
    
    intGrant = zlRegTool '(GetUnitInfo("ע����"))
    If ((intGrant And 4) = 4) Then
        If InStr(1, GetPrivFunc(0, �����嵥.��Ϣ�շ�����), "����") <> 0 Then
            mnuToolMessage.Visible = True
            mnuToolNotify.Visible = True
            blnSplit = True
        Else
            Call zlDatabase.SetPara("�����ʼ���Ϣ", "0")
        End If
    End If
    If ((intGrant And 8) = 8) Then
        If InStr(1, GetPrivFunc(0, �����嵥.EXCEL������), "����") Then
            mnuToolExcel.Visible = True
            blnSplit = True
        End If
    End If
    
    
    If InStr(1, GetPrivFunc(0, �����嵥.ϵͳѡ������), "����") Then
        mnuToolStyle.Visible = True
        blnSplit = True
    End If
    If InStr(1, GetPrivFunc(0, �����嵥.�ֵ������), "����") Then
        mnuToolDictonary.Visible = True
        blnSplit = True
    End If
    MnuBar3.Visible = blnSplit
End Sub

Public Sub RunModual(ByVal lngSys As Long, ByVal lngModual As Long, ByVal strPara As String, Optional ByVal blnReport As Boolean)
    '------------------------------------------------------------------------------------------------------
    '����:����ִ�б���,�˹�����Ϊ�Զ����ѵ��ö�д,by �¸���
    '����:lngSys ϵͳ���;lngModual ģ���
    '------------------------------------------------------------------------------------------------------
    
    On Error GoTo ErrHand
    
    With grsMenus
        If blnReport Then
            .Filter = "ϵͳ=" & lngSys & " AND ģ��=" & lngModual & " And ����=1"
        Else
            .Filter = "ϵͳ=" & lngSys & " AND ģ��=" & lngModual
        End If
        If .RecordCount = 0 Then .Filter = 0: Exit Sub
        If .Fields("ģ��").Value <> 0 Then
            Call ExecuteFunc(.Fields("ϵͳ").Value, IIf(IsNull(.Fields("����").Value), "", .Fields("����").Value), .Fields("ģ��").Value, strPara)
        End If
        .Filter = 0
    End With
    
ErrHand:
    
End Sub

Private Function LoadOutTools(ByVal blnMenu As Boolean) As Boolean
    '-----------------------------------------------------------------------------------
    '����:�����ⲿ����
    '����:blnMenu-�����ʼ�˵���ʾ�б�
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2007/08/22
    '-----------------------------------------------------------------------------------
    Dim i As Long
    Dim strReg As String, arrTemp As Variant, ArrTool As Variant
    Dim objButton As ButtonMenu
    Err = 0: On Error Resume Next
    
    '������ⲿ���߲˵�
    
    For i = 1 To mnuToolOutToolExecute.UBound
        Unload mnuToolOutToolExecute(i)
    Next
    mnuToolOutToolList.Visible = False
    
    '���ع��߲˵�
    strReg = GetSetting("ZLSOFT", "����ȫ��\TOOLS", "TOOLFILES", "")
    Set mcllTemp = New Collection
    
    
    If strReg = "" Then Exit Function
    
    ArrTool = Split(strReg, "|")
    If blnMenu = True Then
        Call SetMenuState(9301, 9301, -4, "��ӹ�������", i, 1, True, "��ӹ�������")
    End If
    
    For i = 0 To UBound(ArrTool)
        arrTemp = Split(ArrTool(i) & ",", ",")
        If arrTemp(0) <> "" And arrTemp(1) <> "" And i <= 199 Then
            If i = 0 Then
                With mnuToolOutToolExecute(0)
                    .Caption = arrTemp(0) & "(&1)"
                    .Tag = arrTemp(1)
                    .Visible = True
                    mnuToolOutToolList.Visible = True
                End With
            Else
                Load mnuToolOutToolExecute(i)
                With mnuToolOutToolExecute(i)
                    .Caption = arrTemp(0) & IIf(i + 1 > 9, "", "(&" & i + 1 & ")")
                    .Tag = arrTemp(1)
                    .Visible = True
                End With
            End If
            If blnMenu = True Then
                Call SetMenuState(9300 + i + 2, -1 * 9300 + i + 2, 0, arrTemp(0), i, 1, True, arrTemp(1))
            End If
            mcllTemp.Add arrTemp(1), "K" & 9300 + i + 2
        End If
    Next
    LoadOutTools = True
End Function


Private Sub ExeCuteToolFile(ByVal strFile As String)
    '-----------------------------------------------------------------------------------
    '����:ִ�й����ļ�
    '����:strFile-�ļ���
    '����:���˺�
    '����:2007/08/22
    '-----------------------------------------------------------------------------------
    Dim objFile As New FileSystemObject
    Err = 0: On Error GoTo ErrHand:
    If objFile.FileExists(strFile) = False Then
        MsgBox "�����ļ�:" & strFile & vbCrLf & "������,�����ѱ�ɾ��,����!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Sub
    End If
    Shell strFile, vbNormalFocus
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Public Function GetCommand() As String
    '����:����ҵ�񲿼���ȡ�����в���,by �¶�
    '����:��
    GetCommand = gstrCommand
End Function

Private Sub DoCommand()
    '���ܣ��ⲿ���õ���̨ʱ�����ݴ����������ҵ�񲿼���,by �¶�
    '��������
    Dim i As Integer, lngModual As Long
    Dim varCmd As Variant
    On Error GoTo errH
    varCmd = Split(gstrCommand, " ")
    For i = LBound(varCmd) To UBound(varCmd)
        If UCase(varCmd(i)) Like "PROGRAM=*" Then
            lngModual = Val(Split(varCmd(i), "=")(1))
            grsMenus.Filter = "ģ��=" & lngModual
            If Not grsMenus.EOF Then
                Call RunModual(grsMenus!ϵͳ, lngModual, "")
                mblnHide = True
            End If
            grsMenus.Filter = 0
        End If
    Next
    Exit Sub
errH:
    
End Sub

Public Sub UnloadForm()
    '���ܣ��ⲿ���õ���̨����ҵ�񲿼���ҵ�񲿼����˳�ʱ��Ҫ���ô˺����رյ���̨��by �¶�
    '��������
    Unload Me
End Sub

Private Sub tmrUpdateConnect_Timer()
    'Ԥ��������
    If DateAdd("n", -30, Now) >= mCurTime Then '30���Ӽ��һ��
        tmrUpdateConnect.Enabled = False
        Call gobjRelogin.UpdateClient
        mCurTime = Now
        tmrUpdateConnect.Enabled = True
    End If
End Sub

Public Function CloseChildWindows(ByVal frmMain As Object) As Boolean
     '����:�ر������Ӵ���
    Dim FrmThis     As Form, ClsClose As Object, IntCount As Integer, LngErr As Long
    Dim objInsure   As Object
    Dim blnOK       As Boolean
    
    On Error Resume Next
    blnOK = True
    If gobjPlugIn Is Nothing Then
        On Error Resume Next
        Set gobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
        If Not gobjPlugIn Is Nothing Then
            Call gobjPlugIn.Initialize(gcnOracle, 0, 0)
            If InStr(",438,0,", "," & Err.Number & ",") = 0 Then
                blnOK = False
                MsgBox "zlPlugIn ��Ҳ���ִ�� Initialize ʱ����" & vbCrLf & Err.Number & vbCrLf & Err.Description, vbInformation, gstrSysName
            End If
        End If
        Err.Clear: On Error GoTo 0
    End If
    
    If Not gobjPlugIn Is Nothing Then
        On Error Resume Next
        Call gobjPlugIn.LogOutBefore
        If InStr(",438,0,", "," & Err.Number & ",") = 0 Then
            blnOK = False
            MsgBox "zlPlugIn ��Ҳ���ִ�� LogOutBefore ʱ����" & vbCrLf & Err.Number & vbCrLf & Err.Description, vbInformation, gstrSysName
        End If
        Err.Clear: On Error GoTo 0
    End If
    On Error Resume Next
    For Each FrmThis In Forms
        If FrmThis.Caption <> frmMain.Caption Then Unload FrmThis
    Next
    '�ر����в����Ĵ���
    If Err.Number <> 0 Then Err.Clear
    LngErr = UBound(gstrObj)
    If Err.Number = 0 Then
        For IntCount = 0 To LngErr
            Set ClsClose = gobjCls(IntCount)
            blnOK = blnOK And ClsClose.CloseWindows
            Set gobjCls(IntCount) = Nothing
        Next
    End If
    '�ر�Ӧ�ù��߰������Ĵ���
    blnOK = blnOK And mclsAppTool.CloseWindows
    '�رչ��������Ĵ���
    blnOK = blnOK And CloseWindows
    Set objInsure = GetObject("", "zl9Insure.clsInsure")
    Call objInsure.Releaseme
    If Err.Number <> 0 Then Err.Clear
    CloseChildWindows = blnOK
End Function

Public Function GetPicDisp(Optional ByVal intIcon As Long = 0, Optional ByVal Blnģ�� As Boolean = True) As IPictureDisp
    '������:����
    '��������:2000-12-12
    '�õ�ͼƬ����

    On Error Resume Next
    If intIcon = 0 Then intIcon = IIf(Blnģ��, -5, -4)
    Select Case intIcon
    Case -1
        Set GetPicDisp = LoadResPicture("HELP", 1)
    Case -2
        Set GetPicDisp = LoadResPicture("RELOGIN", 1)
    Case -3
        Set GetPicDisp = LoadResPicture("EXIT", 1)
    Case -4
        Set GetPicDisp = LoadResPicture("DIRECTORY", 1)
    Case -5
        Set GetPicDisp = LoadResPicture("MODUL", 1)
    Case Else
        Set GetPicDisp = mclsAppTool.GetIcon(intIcon)
    End Select
End Function

Private Sub InitWinsock()
'����:��ȡ����,��ʼ��������
    Dim lngPort As Long
            
    On Error Resume Next
    
    lngPort = Val(zlDatabase.GetPara("����Զ�̿���"))
    mblnRemote = Not lngPort = -1
    winSock.Tag = "1"
    With winSock
        If mblnRemote Then
            .LocalPort = IIf(Val(lngPort) = 0, "1001", Val(lngPort))
            .Listen
        Else
            If .State <> sckClosed Then .Close
        End If
    End With
    winSock.Tag = ""
End Sub

Private Sub winSock_Close()
    If winSock.Tag = "" Then
        If winSock.State <> sckClosed And mblnRemote Then winSock.Close: winSock.Listen  '���¼���
    End If
End Sub

Private Sub winSock_ConnectionRequest(ByVal requestID As Long)
    If winSock.State <> sckClosed Then winSock.Close
    winSock.Accept requestID
End Sub

Private Sub winSock_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    Dim strMsg  As String
    
    winSock.GetData strData
    
    On Error GoTo errH
    If strData = "����Զ��" Then
                RunCommand "REG ADD HKLM\SYSTEM\CurrentControlSet\Control\Terminal"" ""Server /v fDenyTSConnections /t REG_DWORD /d 0 /f"
                winSock.SendData "YES"
    End If
    Exit Sub
errH:
    MsgBox Err.Description
End Sub

Private Sub winSock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    winSock.Close: winSock.Listen
    If winSock.Tag = "" Then
        Select Case Number
            Case 10053
                MsgBox "���ڳ�ʱ��û�в����������Զ��жϡ�", vbInformation, gstrSysName
            Case Else
                MsgBox Number & Description, vbInformation, gstrSysName
         End Select
    Else
        winSock.Tag = ""
    End If
End Sub



