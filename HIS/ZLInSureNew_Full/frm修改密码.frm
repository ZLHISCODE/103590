VERSION 5.00
Begin VB.Form frm�޸����� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�޸�����"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4920
   Icon            =   "frm�޸�����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
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
      Left            =   2340
      PasswordChar    =   "*"
      TabIndex        =   12
      Top             =   255
      Width           =   1875
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
      Left            =   660
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   480
      Width           =   945
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
      Left            =   645
      TabIndex        =   10
      Top             =   75
      Value           =   -1  'True
      Width           =   885
   End
   Begin VB.CommandButton cmd���� 
      Caption         =   "����"
      Height          =   350
      Left            =   4185
      TabIndex        =   9
      Top             =   240
      Width           =   675
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1665
      TabIndex        =   7
      Top             =   2580
      Width           =   1380
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3195
      TabIndex        =   8
      Top             =   2580
      Width           =   1380
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   -225
      TabIndex        =   6
      Top             =   2400
      Width           =   5865
   End
   Begin VB.TextBox txtȷ�������� 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   2355
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1890
      Width           =   2265
   End
   Begin VB.TextBox txt������ 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   2355
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1440
      Width           =   2265
   End
   Begin VB.TextBox txtԭ���� 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   2355
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   990
      Width           =   2265
   End
   Begin VB.Label lbl���� 
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
      Left            =   1755
      TabIndex        =   13
      Top             =   315
      Width           =   510
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   285
      Picture         =   "frm�޸�����.frx":000C
      Top             =   1080
      Width           =   480
   End
   Begin VB.Label lblȷ�������� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ȷ��������(&V)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   585
      TabIndex        =   4
      Top             =   1950
      Width           =   1680
   End
   Begin VB.Label lbl������ 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "������(&N)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1095
      TabIndex        =   2
      Top             =   1500
      Width           =   1170
   End
   Begin VB.Label lblԭ���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ԭ����(&O)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1095
      TabIndex        =   0
      Top             =   1050
      Width           =   1170
   End
End
Attribute VB_Name = "frm�޸�����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstr������ As String
Private mstr������ As String
Private mintLen As Integer
Private mintType As Integer
Private mstrSNO As String, mstrIDNO As String
Private mblnGo As Boolean
Private mblnKey As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Trim(TXT������.Text) = "" Then
        MsgBox "���벻��Ϊ�գ�", vbInformation, gstrSysName
        TXT������.SetFocus
        Exit Sub
    End If
    
    If TXT������.Text <> txtȷ��������.Text Then
        MsgBox "��������������벻һ�£������䣡", vbInformation, gstrSysName
        TXT������.SetFocus
        Exit Sub
    End If
    mstr������ = TXT������.Text
    mstr������ = TXTԭ����.Text
    Unload Me
    Exit Sub
End Sub

Private Sub cmd����_Click()
    mblnGo = True
    mblnGo = ReadCard(TXTԭ����, 0)
End Sub

Private Function ReadCard(ByVal objTxt As TextBox, ByVal intType As Integer) As Boolean
    '
    Dim IntPort As Integer, intCardType As Integer
    Dim strPort As String
    Dim STRERR As String, strPin As String, strPass As String, strAddr As String, strSNO As String
    Dim lngHandle As Long
    Dim blnNO As Boolean
    
    On Error GoTo errHand
    
    If Not mblnGo Then Exit Function
    
    strPort = GetSetting("ZLSOFT", "����ģ��\������ҽ��", "�˿�", "COM1")
    If strPort = "USB" Then
        IntPort = 100
    Else
        IntPort = Right(strPort, 1)
    End If

    '�򿪶�����
    STRERR = Space(2000)
    If SGZ_IFD_Open(IntPort, lngHandle, STRERR) <> 0 Then
        MsgBox STRERR, vbInformation, gstrSysName
        Exit Function
    End If

    '��ȡPSAMоƬ����
    STRERR = Space(2000)
    gstrPSAMNO = Space(2000)
    If SGZ_SAM_ReadNmuber(lngHandle, gstrPSAMNO, STRERR) <> 0 Then
        MsgBox STRERR, vbInformation, gstrSysName
        blnNO = True
        GoTo Exith
    End If
    gstrPSAMNO = TruncZero(gstrPSAMNO)

    If mintType = 2 And intType = 0 Then
         '��ȡ�籣����
        STRERR = Space(2000)
        mstrSNO = Space(2000)
        strPin = "000000"
        strAddr = "MF|EF05|07|$MF|EF06|01|$"
        If SGZ_ICC_ReadCardInfo(lngHandle, intCardType, strPin, strAddr, mstrSNO, STRERR) <> 0 Then
            MsgBox STRERR, vbInformation, gstrSysName
            blnNO = True
            GoTo Exith
        End If
        mstrSNO = TruncZero(mstrSNO)
        mstrIDNO = Split(mstrSNO, "|")(5)
        mstrSNO = Split(mstrSNO, "|")(2)
        txt����.Text = mstrSNO
    End If
    If objTxt.Enabled = True Then objTxt.SetFocus
    STRERR = Space(2000)
    strPass = Space(2000)
    If SGZ_IFD_GetPIN(lngHandle, IntPort, strPass, STRERR) <> 0 Then
        MsgBox STRERR, vbInformation, gstrSysName
        blnNO = True
        GoTo Exith
    End If
    strPass = TruncZero(strPass)
    objTxt.Text = strPass
    
Exith:
    STRERR = Space(2000)
    If lngHandle > 0 Then Call SGZ_IFD_Close(lngHandle, STRERR)
    If blnNO = True Then Exit Function
    Select Case objTxt.Name
        Case "txtԭ����"
            Call txtԭ����_KeyDown(vbKeyReturn, 0)
        Case "txt������"
            Call TXT������_KeyDown(vbKeyReturn, 0)
        Case "txtȷ��������"
            Call txtȷ��������_KeyDown(vbKeyReturn, 0)
    End Select
    ReadCard = True
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    GoTo Exith
End Function

Private Sub Form_Activate()
    'Modified by ZYB �Ͻ�
    If TXTԭ����.Text <> "" Then Me.TXT������.SetFocus

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab): mblnKey = True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    TXTԭ����.Text = mstr������
    mstr������ = ""
    mblnGo = False
    mblnKey = False
    If gintType = 1 Then
        opt�����.Item(0).Value = True
        mintType = 1
    Else
        opt�����.Item(1).Value = True
        mintType = 2
    End If
    mstrSNO = "": mstrIDNO = ""
    Me.txtȷ��������.MaxLength = mintLen
    Me.TXT������.MaxLength = mintLen
    Me.TXTԭ����.MaxLength = mintLen
End Sub

Private Sub opt�����_Click(Index As Integer)
    If Index = 0 Then
        lbl����.Caption = "����"
        cmd����.Caption = "����"
    Else
        lbl����.Caption = "IC��"
        cmd����.Caption = "����"
    End If
    mintType = Index + 1
End Sub

Private Sub txtȷ��������_GotFocus()
    txtȷ��������.SelStart = 0
    txtȷ��������.SelLength = mintLen
End Sub

Private Sub txtȷ��������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And mblnKey = False Then zlCommFun.PressKey (vbKeyTab)
    mblnKey = False
End Sub

Private Sub txt������_GotFocus()
    TXT������.SelStart = 0
    TXT������.SelLength = mintLen
End Sub

Private Sub TXT������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And mblnKey = False Then zlCommFun.PressKey (vbKeyTab)
    mblnKey = False
    Call ReadCard(txtȷ��������, 1)
End Sub

Private Sub txtԭ����_Change()
    cmdOK.Enabled = (Len(TXTԭ����.Text) <> 0)
End Sub

Private Sub txtԭ����_GotFocus()
    TXTԭ����.SelStart = 0
    TXTԭ����.SelLength = mintLen
End Sub

Public Function ChangePassword(ByVal strPass As String, Optional ByRef strOldPassWord As String = "", Optional ByVal intLen As Integer = 8) As String
    mintLen = intLen
    mstr������ = strPass
    Me.Show 1
    strOldPassWord = mstr������
    ChangePassword = mstr������
    
End Function

Private Sub txtԭ����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And mblnKey = False Then zlCommFun.PressKey (vbKeyTab)
    mblnKey = False
    Call ReadCard(TXT������, 1)
End Sub
