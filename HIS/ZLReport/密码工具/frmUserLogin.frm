VERSION 5.00
Begin VB.Form frmUserLogin 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����Ա��¼"
   ClientHeight    =   2205
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   4170
   Icon            =   "frmUserLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   4170
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   -360
      TabIndex        =   8
      Top             =   1455
      Width           =   5025
   End
   Begin VB.CommandButton CMD���� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   2865
      TabIndex        =   4
      Top             =   1710
      Width           =   1100
   End
   Begin VB.CommandButton CDMȷ�� 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   1755
      TabIndex        =   3
      Top             =   1710
      Width           =   1100
   End
   Begin VB.TextBox TXT���� 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1950
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   630
      Width           =   1920
   End
   Begin VB.TextBox txt���ݿ� 
      Height          =   300
      Left            =   1950
      TabIndex        =   2
      Top             =   1050
      Width           =   1920
   End
   Begin VB.TextBox txt�û� 
      Height          =   300
      Left            =   1950
      TabIndex        =   0
      Top             =   195
      Width           =   1920
   End
   Begin VB.Image imgFlag 
      Height          =   720
      Left            =   315
      Picture         =   "frmUserLogin.frx":1CFA
      Top             =   240
      Width           =   720
   End
   Begin VB.Label Lbl������ 
      AutoSize        =   -1  'True
      Caption         =   "������"
      Height          =   180
      Left            =   1320
      TabIndex        =   7
      Top             =   1110
      Width           =   540
   End
   Begin VB.Label Lbl���� 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Left            =   1500
      TabIndex        =   6
      Top             =   690
      Width           =   360
   End
   Begin VB.Label Lbl�û��� 
      AutoSize        =   -1  'True
      Caption         =   "�û���"
      Height          =   180
      Left            =   1320
      TabIndex        =   5
      Top             =   255
      Width           =   540
   End
End
Attribute VB_Name = "frmUserLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim intTimes As Integer
Dim strNote As String
Dim strUserName As String
Dim strServerName As String
Dim strPassword As String

Private Sub CDMȷ��_Click()
    SetConState False
    intTimes = intTimes + 1
    
    '------�����û��Ƿ�oracle�Ϸ��û�----------------
    strUserName = Trim(txt�û�.Text)
    strServerName = Trim(txt���ݿ�.Text)
    
    '��Ч�ַ���Ч��
    If Len(Trim(txt�û�)) = 0 Then
        strNote = "�������û���"
        txt�û�.SetFocus
        GoTo InputError
    End If
    
    If Len(strUserName) <> 1 Then
        If Mid(strUserName, 1, 1) = "/" Or Mid(strUserName, 1, 1) = "@" Or Mid(strUserName, Len(strUserName) - 1, 1) = "/" Or Mid(strUserName, Len(strUserName) - 1, 1) = "@" Then
            txt�û�.SetFocus
            strNote = "�û�������"
            SetConState
            Exit Sub
        End If
    End If
    If Trim(strPassword) <> "" And Len(strPassword) <> 1 Then
        If Mid(strPassword, Len(strPassword) - 1, 1) = "/" Or Mid(strPassword, Len(strPassword) - 1, 1) = "@" Or Mid(strPassword, 1, 1) = "/" Or Mid(strPassword, 1, 1) = "@" Then
            If TXT����.Enabled Then TXT����.SetFocus
            strNote = "�������"
            GoTo InputError
        End If
    End If
    If Trim(strServerName) <> "" Then
        If Mid(strServerName, Len(strServerName) - 1, 1) = "/" Or Mid(strServerName, Len(strServerName) - 1, 1) = "@" Or Mid(strServerName, 1, 1) = "/" Or Mid(strServerName, 1, 1) = "@" Then
            strNote = "�������Ӵ�����"
            txt���ݿ�.SetFocus
            GoTo InputError
        End If
    End If
    
    '�����ַ���
    Dim intPos As Integer
    strPassword = TXT����.Text
    
    intPos = InStr(1, strUserName, "@", vbTextCompare)
    If intPos > 0 Then
        strServerName = Mid(strUserName, intPos + 1)
        strUserName = Mid(strUserName, 1, intPos - 1)
    End If
    
    intPos = InStr(1, strUserName, "/", vbTextCompare)
    If intPos > 0 Then
        strPassword = Mid(strUserName, intPos + 1)
        strUserName = Mid(strUserName, 1, intPos - 1)
    End If
    
    intPos = InStr(1, strPassword, "@", vbTextCompare)
    If intPos > 0 Then
        strServerName = Mid(strPassword, intPos + 1)
        strPassword = Mid(strPassword, 1, intPos - 1)
    End If
    
    
    If Len(Trim(strPassword)) = 0 Then
        strNote = "����������"
        GoTo InputError
    End If
    
    If Not OraDataOpen(strServerName, strUserName, IIf(UCase(strUserName) = "SYS" Or UCase(strUserName) = "SYSTEM", strPassword, TranPasswd(strPassword))) Then
        TXT����.Text = ""
        If TXT����.Enabled Then TXT����.SetFocus
        SetConState
        Exit Sub
    End If
    
    '�޸�ע���
    SaveSetting "ZLSOFT", "����", "USER", strUserName
    SaveSetting "ZLSOFT", "����", "SERVER", strServerName
    
    '������ݷ�ʽ��
    SaveSetting "ZLSOFT", "����", "����·��", App.Path & "\" & App.EXEName & ".exe"
    
    Unload Me
    Exit Sub
InputError:
    If intTimes > 3 Then
        MsgBox "�������ε�¼ʧ�ܣ�ϵͳ���Զ��˳�", vbExclamation, App.Title
        CMD����_Click
    Else
        If strNote <> "" Then
            MsgBox strNote, vbExclamation, App.Title
        End If
        SetConState
        Exit Sub
    End If
End Sub

Private Sub CMD����_Click()
    Set gcnOracle = Nothing
    Unload Me
End Sub

Private Sub Form_Activate()
    If TXT����.Enabled Then
        TXT����.SetFocus
    Else
        CDMȷ��.SetFocus
    End If
End Sub

Private Sub Form_Load()
    intTimes = 1
    txt�û�.Text = GetSetting(appName:="ZLSOFT", Section:="����", Key:="USER", Default:="")
    txt���ݿ�.Text = GetSetting(appName:="ZLSOFT", Section:="����", Key:="SERVER", Default:="")
End Sub

Private Sub GetFocus(ByVal TxtBox As TextBox)
    With TxtBox
        .SelStart = 0
        .SelLength = LenB(StrConv(.Text, vbFromUnicode))
    End With
End Sub

Private Sub txt�û�_GotFocus()
    GetFocus txt�û�
End Sub

Private Sub TXT����_GotFocus()
    GetFocus TXT����
End Sub

Private Sub txt���ݿ�_GotFocus()
    GetFocus txt���ݿ�
End Sub

Private Sub SetConState(Optional ByVal BlnState As Boolean = True)
    CMD����.Enabled = BlnState
    CDMȷ��.Enabled = BlnState
End Sub
