VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4365
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   6480
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   4365
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   1005
      TabIndex        =   8
      Top             =   3720
      Width           =   6135
   End
   Begin VB.Image imgPic 
      Height          =   2700
      Left            =   150
      Picture         =   "frmSplash.frx":5D0A2
      Top             =   420
      Width           =   1200
   End
   Begin VB.Label LblProductName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   1590
      TabIndex        =   7
      Top             =   1350
      Width           =   4650
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ʹ��Ȩ���ڣ�"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1650
      TabIndex        =   6
      Top             =   2205
      Width           =   1080
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��Ʒ�����̣�"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1650
      TabIndex        =   5
      Top             =   3030
      Width           =   1080
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����֧���̣�"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1650
      TabIndex        =   4
      Top             =   2610
      Width           =   1080
   End
   Begin VB.Label lblGrant 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2745
      TabIndex        =   3
      Top             =   2205
      Width           =   90
   End
   Begin VB.Label lbl����֧���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2745
      TabIndex        =   2
      Top             =   2610
      Width           =   90
   End
   Begin VB.Label lbl������ 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2745
      TabIndex        =   1
      Top             =   3030
      Width           =   90
   End
   Begin VB.Image ImgIndicate 
      Appearance      =   0  'Flat
      Height          =   720
      Left            =   165
      Picture         =   "frmSplash.frx":5DB05
      Stretch         =   -1  'True
      Top             =   3390
      Width           =   720
   End
   Begin VB.Label lblWarning 
      BackStyle       =   0  'Transparent
      Caption         =   "���棺���������������������ʹ�����֤������δ����Ȩ��ɣ��κ��˲��ø��ơ����ۼ����ܴ���������򽫳е�ȫ���������Ρ�"
      Height          =   465
      Left            =   1095
      TabIndex        =   0
      Top             =   3825
      Width           =   5550
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ShowSplash()
    Dim StrUnitName As String
    Dim intCount As Integer
    
    Load frmSplash 'ǿ��װ�룬�Է������ĳ���ֱ��ж��
    '��ע����л�ȡ�û�ע�������Ϣ,����û���λ���Ʋ�Ϊ��,����ʾ���ִ���
    StrUnitName = GetSetting("ZLSOFT", "ע����Ϣ", "��λ����", "")
    If StrUnitName <> "" Then
        lblGrant = StrUnitName
        LblProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒȫ��", "")
        gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "") & "���"
        lbl����֧���� = GetSetting("ZLSOFT", "ע����Ϣ", "����֧����", "")
        lbl������ = ""
        StrUnitName = GetSetting("ZLSOFT", "ע����Ϣ", "������", "")
        
        gstrProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "")
        Call ApplyOEM_Picture(ImgIndicate, "Picture")
        Call ApplyOEM_Picture(imgPic, "PictureB")
        
        If Trim(StrUnitName) = "" Then
            Label3.Visible = False
            lbl������.Visible = False
        Else
            For intCount = 0 To UBound(Split(StrUnitName, ";"))
                lbl������.Caption = lbl������.Caption & Split(StrUnitName, ";")(intCount) & vbCrLf
            Next
        End If
        frmSplash.Show
        DoEvents
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    gdtStart = 0
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    gdtStart = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    If gcnOracle.State = adStateOpen Then
        If RegCheck = False Then
            Exit Sub
        End If
        gstrProductTitle = zlReginfo("��Ʒ����")
        gstrProductName = zlReginfo("��Ʒ����")
        gstrDevelopers = zlReginfo("��Ʒ������")
        gstrSustainer = zlReginfo("����֧����")
        gstrWebSustainer = zlReginfo("֧���̼���")
        gstrWebURL = zlReginfo("֧����URL")
        gstrWebEmail = zlReginfo("֧����MAIL")
        '���û�ע�������Ϣд��ע���,���´�����ʱ��ʾ
        SaveSetting "ZLSOFT", "ע����Ϣ", "��Ʒȫ��", gstrProductTitle
        SaveSetting "ZLSOFT", "ע����Ϣ", "��Ʒ����", gstrProductName
        SaveSetting "ZLSOFT", "ע����Ϣ", "����֧����", gstrSustainer
        SaveSetting "ZLSOFT", "ע����Ϣ", "������", gstrDevelopers
        SaveSetting "ZLSOFT", "ע����Ϣ", "WEB֧���̼���", gstrWebSustainer
        SaveSetting "ZLSOFT", "ע����Ϣ", "WEB֧��EMAIL", gstrWebEmail
        SaveSetting "ZLSOFT", "ע����Ϣ", "WEB֧��URL", gstrWebURL
        SaveSetting "ZLSOFT", "ע����Ϣ", "��λ����", zlReginfo("��λ����")
    End If
End Sub

'�����˳�
Private Sub ImgIndicate_Click()
    gdtStart = 0
End Sub

Private Sub LblProductName_Click()
    gdtStart = 0
End Sub

Private Sub Label1_Click()
    gdtStart = 0
End Sub

Private Sub Label2_Click()
    gdtStart = 0
End Sub

Private Sub Label3_Click()
    gdtStart = 0
End Sub

Private Sub lblWarning_Click()
    gdtStart = 0
End Sub


