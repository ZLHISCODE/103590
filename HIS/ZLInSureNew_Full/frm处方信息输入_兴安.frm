VERSION 5.00
Begin VB.Form frm������Ϣ����_�˰� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������Ϣ����"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5805
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txtEdit 
      Height          =   315
      Index           =   5
      Left            =   3975
      MaxLength       =   8
      TabIndex        =   11
      Tag             =   "��������"
      Top             =   1868
      Width           =   1575
   End
   Begin VB.TextBox txtEdit 
      Height          =   315
      Index           =   4
      Left            =   945
      MaxLength       =   8
      TabIndex        =   9
      Tag             =   "��������"
      Top             =   1868
      Width           =   1575
   End
   Begin VB.TextBox txtEdit 
      Height          =   315
      Index           =   3
      Left            =   3975
      MaxLength       =   8
      TabIndex        =   7
      Tag             =   "������"
      Top             =   1433
      Width           =   1575
   End
   Begin VB.TextBox txtEdit 
      Height          =   315
      Index           =   2
      Left            =   945
      MaxLength       =   8
      TabIndex        =   5
      Tag             =   "�ε�λ"
      Top             =   1433
      Width           =   1530
   End
   Begin VB.TextBox txtEdit 
      Height          =   315
      Index           =   1
      Left            =   3975
      MaxLength       =   8
      TabIndex        =   3
      Tag             =   "������"
      Top             =   983
      Width           =   1575
   End
   Begin VB.TextBox txtEdit 
      Height          =   315
      Index           =   0
      Left            =   945
      MaxLength       =   3
      TabIndex        =   1
      Tag             =   "�մ���"
      Top             =   983
      Width           =   1530
   End
   Begin VB.Frame fra 
      Height          =   45
      Index           =   1
      Left            =   15
      TabIndex        =   15
      Top             =   705
      Width           =   7110
   End
   Begin VB.Frame fra 
      Height          =   45
      Index           =   0
      Left            =   -1110
      TabIndex        =   14
      Top             =   2430
      Width           =   9300
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   3150
      TabIndex        =   12
      Top             =   2640
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4575
      TabIndex        =   13
      Top             =   2640
      Width           =   1100
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "��������"
      Height          =   180
      Index           =   5
      Left            =   3240
      TabIndex        =   10
      Top             =   1935
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "��������"
      Height          =   180
      Index           =   4
      Left            =   210
      TabIndex        =   8
      Top             =   1935
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "������"
      Height          =   180
      Index           =   3
      Left            =   3420
      TabIndex        =   6
      Top             =   1500
      Width           =   540
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�ε�λ"
      Height          =   180
      Index           =   2
      Left            =   390
      TabIndex        =   4
      Tag             =   "�ε�λ"
      Top             =   1500
      Width           =   540
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "������"
      Height          =   180
      Index           =   6
      Left            =   3420
      TabIndex        =   2
      Top             =   1050
      Width           =   540
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�մ���"
      Height          =   180
      Index           =   1
      Left            =   390
      TabIndex        =   0
      Top             =   1050
      Width           =   540
   End
   Begin VB.Label lbl 
      Caption         =   "�����봦���������Ϣ�����մ���,�������ȡ�"
      Height          =   165
      Index           =   0
      Left            =   945
      TabIndex        =   16
      Top             =   375
      Width           =   4965
   End
   Begin VB.Image img 
      Height          =   555
      Left            =   180
      Picture         =   "frm������Ϣ����_�˰�.frx":0000
      Stretch         =   -1  'True
      Top             =   75
      Width           =   615
   End
End
Attribute VB_Name = "frm������Ϣ����_�˰�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mstrInfor As String
Dim mblnOK As Boolean
Dim mblnChange As Boolean


Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub
Private Function IsValid() As Boolean
    '������֤
    Dim i As Long
    Dim strTemp As String
    IsValid = False
    For i = 0 To 5
        strTemp = txtEdit(i).Text
        
        If LenB(StrConv(strTemp, vbFromUnicode)) > txtEdit(i).MaxLength Then
            ShowMsgbox txtEdit(i).Tag & "����,���������" & txtEdit(i).MaxLength / 2 & "�����ֻ�" & txtEdit(i).MaxLength & "���ַ�!"
            If txtEdit(i).Enabled Then txtEdit(i).SetFocus
            Exit Function
        End If
        If InStr(1, strTemp, "'") <> 0 Then
            ShowMsgbox txtEdit(i).Tag & "�������뵥����!"
            If txtEdit(i).Enabled Then txtEdit(i).SetFocus
            Exit Function
        End If
        If InStr(1, strTemp, ":") <> 0 Then
            ShowMsgbox txtEdit(i).Tag & "��������ð��!"
            If txtEdit(i).Enabled Then txtEdit(i).SetFocus
            Exit Function
        End If
        If InStr(1, strTemp, vbTab) <> 0 Then
            ShowMsgbox txtEdit(i).Tag & "���������Ʊ��!"
            If txtEdit(i).Enabled Then txtEdit(i).SetFocus
            Exit Function
        End If
    Next
    
    IsValid = True
End Function
Private Sub cmdOK_Click()
    Dim i As Long
    Dim strInfor As String
    
    If IsValid = False Then Exit Sub
    strInfor = ""
    For i = 0 To 5
        If i = 5 Then
            strInfor = strInfor & txtEdit(i).Tag & ":" & txtEdit(i).Text
        Else
            strInfor = strInfor & txtEdit(i).Tag & ":" & txtEdit(i).Text & vbTab
        End If
    Next
    mstrInfor = strInfor
    mblnOK = True
    Unload Me
End Sub

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
    SetOkCtl
End Sub
Private Sub SetOkCtl()
    cmdOK.Enabled = mblnChange
End Sub

Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    Dim strTag As String
    strTag = txtEdit(Index).Tag
    Select Case strTag
    Case "�մ���"
        txtEdit(3).Text = Val(txtEdit(0).Text) * Val(txtEdit(1).Text)
    Case "������"
        txtEdit(3).Text = Val(txtEdit(0).Text) * Val(txtEdit(1).Text)
    Case "������"
        txtEdit(5).Text = Val(txtEdit(3).Text) * Val(txtEdit(4).Text)
    Case "��������"
        txtEdit(5).Text = Val(txtEdit(3).Text) * Val(txtEdit(4).Text)
    End Select
    zlCommFun.PressKey vbKeyTab
End Sub


Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If txtEdit(Index).Tag = "�ε�λ" Then
        If InStr(1, ";:", Chr(KeyAscii)) <> 0 Then
            KeyAscii = 0
            Exit Sub
        End If
        zlControl.TxtCheckKeyPress txtEdit(Index), KeyAscii, m�ı�ʽ
    Else
        zlControl.TxtCheckKeyPress txtEdit(Index), KeyAscii, m���ʽ
    End If
End Sub
Public Function EditCard(strInfor As String) As Boolean
    '���ܣ��༭��������
    '����:strInfor-�����Ϣ,�Էֺŷָ�vbkeytab�ָ�
    Dim strArr
    Dim i As Integer
    Dim intMouse As Integer
    intMouse = Screen.MousePointer
    Screen.MousePointer = 1
    
    If strInfor <> "" Then
        Err = 0
        On Error Resume Next
        strArr = Split(strInfor, vbTab)
        For i = 0 To 5
            If i > UBound(strArr) Then Exit For
            txtEdit(i).Text = Split(strArr(i), ":")(1)
        Next
    End If
    Me.Show 1
    If mblnOK Then
        strInfor = mstrInfor
    Else
        strInfor = ""
    End If
    EditCard = mblnOK
    Screen.MousePointer = intMouse
End Function
