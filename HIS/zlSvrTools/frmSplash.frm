VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4395
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   6480
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   4395
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picHos 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2750
      Left            =   1650
      ScaleHeight     =   2745
      ScaleWidth      =   4830
      TabIndex        =   10
      Top             =   0
      Width           =   4835
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   1005
      TabIndex        =   8
      Top             =   3720
      Width           =   6135
   End
   Begin VB.Image imgPic 
      Height          =   2745
      Left            =   150
      Picture         =   "frmSplash.frx":5D0A2
      Top             =   420
      Width           =   1260
   End
   Begin VB.Label lblTag 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   4785
      TabIndex        =   9
      Top             =   1815
      Width           =   615
   End
   Begin VB.Label LblProductName 
      Alignment       =   2  'Center
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
      Left            =   1515
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
      Top             =   2430
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
      Top             =   3255
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
      Top             =   2835
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
      Top             =   2430
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
      Top             =   2835
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
      Top             =   3255
      Width           =   90
   End
   Begin VB.Image ImgIndicate 
      Appearance      =   0  'Flat
      Height          =   720
      Left            =   165
      Picture         =   "frmSplash.frx":5D923
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
    Dim objPic As IPictureDisp
    
    Load frmSplash 'ǿ��װ�룬�Է������ĳ���ֱ��ж��
    '��ע����л�ȡ�û�ע�������Ϣ,����û���λ���Ʋ�Ϊ��,����ʾ���ִ���
    StrUnitName = GetSetting("ZLSOFT", "ע����Ϣ", "��λ����", "")
    If StrUnitName <> "" Then
        lblGrant = StrUnitName
        LblProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒȫ��", "")
        If Len(LblProductName.Caption) > 10 Then
            LblProductName.FontSize = 15.75 '����
        Else
            LblProductName.FontSize = 21.75  '����
        End If
        gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "") & "���"
        lbl����֧���� = GetSetting("ZLSOFT", "ע����Ϣ", "����֧����", "")
        lbl������ = ""
        StrUnitName = GetSetting("ZLSOFT", "ע����Ϣ", "������", "")
    
        gstrProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "")
        gstrUltimatetag = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒϵ��", "")
        lblTag = gstrUltimatetag
        
        Call ApplyOEM_Picture(ImgIndicate, "Picture")
        Call ApplyOEM_Picture(imgPic, "PictureB")
        If gobjFile Is Nothing Then Set gobjFile = New FileSystemObject
        If gstrAppsoft = "" Then
            gstrAppsoft = App.Path
            If gblnInIDE Then
                gstrAppsoft = "C:\APPSOFT"
            End If
        End If

        If gobjFile.FileExists(gstrAppsoft & "\�����ļ�\logo_login.jpg") Then
            Set objPic = LoadPicture(gstrAppsoft & "\�����ļ�\logo_login.jpg")
            picHos.Visible = True
            picHos.Height = IIf(objPic.Height < 2745, objPic.Height, 2745) '183����
            picHos.Width = IIf(objPic.Width < 4845, objPic.Width, 4845) '323����
            picHos.PaintPicture objPic, 0, 0, picHos.Width, picHos.Height
        Else
            picHos.Visible = False
        End If
        If Trim$(lbl����֧����.Caption) = "" Then
            Label1.Visible = False
            lbl����֧����.Visible = False
        Else
            Label1.Visible = True
            lbl����֧����.Visible = True
        End If
        
        If Trim(StrUnitName) = "" Then
            Label3.Visible = False
            lbl������.Visible = False
        Else
            Label3.Visible = True
            lbl������.Visible = True
            lbl������.Caption = ""
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

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    gdtStart = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strCode As String
       
    If gcnOracle.State = adStateOpen Then
        If gblnCreate Then
            gstrUltimatetag = ""
            gstrProductTitle = gobjRegister.zlRegInfo("��Ʒ����")
            If gstrProductTitle <> "" Then
                If InStr(gstrProductTitle, "-") > 0 Then
                    If Split(gstrProductTitle, "-")(1) = "Ultimate" Then
                        gstrUltimatetag = "�콢��"
                    ElseIf Split(gstrProductTitle, "-")(1) = "Professional" Then
                        gstrUltimatetag = "רҵ��"
                    End If
                End If
            End If
            gstrProductTitle = Split(gstrProductTitle, "-")(0)
            
            gstrProductName = gobjRegister.zlRegInfo("��Ʒ����")
            gstrDevelopers = gobjRegister.zlRegInfo("��Ʒ������", , -1)
            gstrSustainer = gobjRegister.zlRegInfo("����֧����", , -1)
            gstrWebSustainer = gobjRegister.zlRegInfo("֧���̼���")
            gstrWebURL = gobjRegister.zlRegInfo("֧����URL")
            gstrWebEmail = gobjRegister.zlRegInfo("֧����MAIL")
            
            '���û�ע�������Ϣд��ע���,���´�����ʱ��ʾ
            SaveSetting "ZLSOFT", "ע����Ϣ", "��Ʒȫ��", gstrProductTitle
            SaveSetting "ZLSOFT", "ע����Ϣ", "��Ʒ����", gstrProductName
            SaveSetting "ZLSOFT", "ע����Ϣ", "����֧����", gstrSustainer
            SaveSetting "ZLSOFT", "ע����Ϣ", "������", gstrDevelopers
            SaveSetting "ZLSOFT", "ע����Ϣ", "WEB֧���̼���", gstrWebSustainer
            SaveSetting "ZLSOFT", "ע����Ϣ", "WEB֧��EMAIL", gstrWebEmail
            SaveSetting "ZLSOFT", "ע����Ϣ", "WEB֧��URL", gstrWebURL
            SaveSetting "ZLSOFT", "ע����Ϣ", "��λ����", gobjRegister.zlRegInfo("��λ����", , -1)
            SaveSetting "ZLSOFT", "ע����Ϣ", "��Ʒϵ��", gstrUltimatetag
            
        End If
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
