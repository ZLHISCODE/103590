VERSION 5.00
Begin VB.Form frmˢ�� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ˢ��"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4050
   Icon            =   "frmˢ��.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   4050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmd���� 
      Caption         =   "����"
      Height          =   350
      Left            =   3300
      TabIndex        =   5
      Top             =   585
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.CommandButton cmd���� 
      Caption         =   "����"
      Height          =   350
      Left            =   3300
      TabIndex        =   7
      ToolTipText     =   "�����������"
      Top             =   1050
      Width           =   675
   End
   Begin VB.OptionButton opt����� 
      Caption         =   "�ſ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   360
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Value           =   -1  'True
      Width           =   885
   End
   Begin VB.OptionButton opt����� 
      Caption         =   "IC��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   360
      Index           =   1
      Left            =   1350
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   945
   End
   Begin VB.OptionButton opt����� 
      Caption         =   "���֤��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   360
      Index           =   2
      Left            =   2460
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   1365
   End
   Begin VB.TextBox txt���� 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1020
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   1050
      Width           =   2295
   End
   Begin VB.CommandButton cmdCancle 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   2790
      TabIndex        =   10
      Top             =   1620
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   1560
      TabIndex        =   9
      Top             =   1620
      Width           =   1100
   End
   Begin VB.TextBox txt���� 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1020
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   495
      TabIndex        =   6
      Top             =   1110
      Width           =   510
   End
   Begin VB.Label lbl���� 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   495
      TabIndex        =   3
      Top             =   660
      Width           =   510
   End
End
Attribute VB_Name = "frmˢ��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrCard As String
Private mstrPass As String
Private mblnInit As Boolean

Public Function ShowME() As String
    mstrCard = ""
    mstrPass = ""
    mblnInit = False
    Me.Show 1
    ShowME = mstrCard & "|" & mstrPass
End Function

Private Sub cmdCancle_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Trim(txt����.Text) = "" Then
        MsgBox "��ˢ����", vbInformation, gstrSysName
        If txt����.Enabled = True Then txt����.SetFocus
        Exit Sub
    End If
    
    If gintType = 3 Then gstrIDNO = txt����.Text
    mstrCard = txt����.Text
    mstrPass = txt����.Text
    Unload Me
End Sub

Private Sub Form_Activate()
    If mblnInit Then Exit Sub
    mblnInit = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    mblnInit = False
    Select Case gintType
        Case 1
            opt�����.Item(0).Value = True
        Case 2
            opt�����.Item(1).Value = True
        Case Else
            opt�����.Item(2).Value = True
            txt����.Text = gstrIDNO
    End Select
    gstrIDNO = ""
End Sub

Private Sub opt�����_Click(Index As Integer)
    gintType = Index + 1
    txt����.Enabled = (Index <> 1)
    cmd����.Visible = (Index = 1)
    cmd����.Visible = (Index <> 1)
    Select Case Index
    Case 0
        lbl����.Caption = "����"
    Case 1
        lbl����.Caption = "IC����"
    Case 2
        lbl����.Caption = "���֤��"
    End Select
    If Index <> 1 And mblnInit = True And txt����.Enabled = True Then txt����.SetFocus
End Sub

Private Sub txt����_GotFocus()
    If gblnLED Then
        zl9LedVoice.Speak "#5"
    End If
End Sub

Private Sub txt����_GotFocus()
    If gblnLED Then
        zl9LedVoice.Speak "#0"
    End If
End Sub

Private Sub cmd����_Click()
    Dim intPort As Integer, intCardType As Integer
    Dim strPort As String
    Dim STRERR As String, strPin As String, strPass As String, strAddr As String
    Dim lngHandle As Long
    On Error GoTo errHand
    
    strPort = GetSetting("ZLSOFT", "����ģ��\������ҽ��", "�˿�", "COM1")
    If strPort = "USB" Then
        intPort = 100
    Else
        intPort = Right(strPort, 1)
    End If
    
    '�򿪶�����
    STRERR = Space(2000)
    If SGZ_IFD_Open(intPort, lngHandle, STRERR) <> 0 Then
        MsgBox STRERR, vbInformation, gstrSysName
        Exit Sub
    End If
    
    '��ȡPSAMоƬ����
    STRERR = Space(2000)
    gstrPSAMNO = Space(2000)
    If SGZ_SAM_ReadNmuber(lngHandle, gstrPSAMNO, STRERR) <> 0 Then
        MsgBox STRERR, vbInformation, gstrSysName
        GoTo Exith
    End If
    gstrPSAMNO = TruncZero(gstrPSAMNO)
    
    '��ȡ�籣����
    STRERR = Space(2000)
    gstrSNO = Space(2000)
    strPin = "000000"
    strAddr = "MF|EF05|07|$MF|EF06|01|$"
    If SGZ_ICC_ReadCardInfo(lngHandle, intCardType, strPin, strAddr, gstrSNO, STRERR) <> 0 Then
        MsgBox STRERR, vbInformation, gstrSysName
        GoTo Exith
    End If
    gstrSNO = TruncZero(gstrSNO)
    gstrIDNO = Split(gstrSNO, "|")(5)
    gstrSNO = Split(gstrSNO, "|")(2)
    txt����.Text = gstrSNO
    
    STRERR = Space(2000)
    strPass = Space(2000)
    If SGZ_IFD_GetPIN(lngHandle, intPort, strPass, STRERR) <> 0 Then
        MsgBox STRERR, vbInformation, gstrSysName
        GoTo Exith
    End If
    strPass = TruncZero(strPass)
    txt����.Text = strPass
    
Exith:
    STRERR = Space(2000)
    If lngHandle > 0 Then Call SGZ_IFD_Close(lngHandle, STRERR)
    Call cmdOK_Click
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    GoTo Exith
End Sub

Private Sub cmd����_Click()
        Dim intPort As Integer, intCardType As Integer
    Dim strPort As String
    Dim STRERR As String, strPin As String, strPass As String, strAddr As String
    Dim lngHandle As Long
    On Error GoTo errHand
    
    strPort = GetSetting("ZLSOFT", "����ģ��\������ҽ��", "�˿�", "COM1")
    If strPort = "USB" Then
        intPort = 100
    Else
        intPort = Right(strPort, 1)
    End If
    
    '�򿪶�����
    STRERR = Space(2000)
    If SGZ_IFD_Open(intPort, lngHandle, STRERR) <> 0 Then
        MsgBox STRERR, vbInformation, gstrSysName
        Exit Sub
    End If
    
    '��ȡPSAMоƬ����
    STRERR = Space(2000)
    gstrPSAMNO = Space(2000)
    If SGZ_SAM_ReadNmuber(lngHandle, gstrPSAMNO, STRERR) <> 0 Then
        MsgBox STRERR, vbInformation, gstrSysName
        GoTo Exith
    End If
    gstrPSAMNO = TruncZero(gstrPSAMNO)
    
    STRERR = Space(2000)
    strPass = Space(2000)
    If SGZ_IFD_GetPIN(lngHandle, intPort, strPass, STRERR) <> 0 Then
        MsgBox STRERR, vbInformation, gstrSysName
        GoTo Exith
    End If
    strPass = TruncZero(strPass)
    txt����.Text = strPass
    
Exith:
    STRERR = Space(2000)
    If lngHandle > 0 Then Call SGZ_IFD_Close(lngHandle, STRERR)
    Call cmdOK_Click
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    GoTo Exith
End Sub
