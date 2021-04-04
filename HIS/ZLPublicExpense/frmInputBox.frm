VERSION 5.00
Begin VB.Form frmInputBox 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2376
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5448
   Icon            =   "frmInputBox.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2376
   ScaleWidth      =   5448
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picCmd 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   690
      Left            =   0
      ScaleHeight     =   696
      ScaleWidth      =   5448
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1680
      Width           =   5445
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
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
         Height          =   420
         Left            =   3720
         TabIndex        =   2
         Top             =   165
         Width           =   1500
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
         Height          =   420
         Left            =   2115
         TabIndex        =   1
         Top             =   165
         Width           =   1500
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000010&
         X1              =   0
         X2              =   6000
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         X1              =   -60
         X2              =   6000
         Y1              =   15
         Y2              =   15
      End
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Left            =   135
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   420
      Width           =   5175
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������XXX:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   135
      TabIndex        =   4
      Top             =   120
      Width           =   1200
   End
End
Attribute VB_Name = "frmInputBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Private mintLen As Integer
Private mblnNULL As Boolean
Private mblnEnter As Boolean
Private mstrText As String
Private mblnOK As Boolean

Public Function InputBox(frmParent As Object, ByVal ���� As String, ByVal ��ʾ As String, _
    ByVal ���� As Integer, ByVal ���� As Byte, ByVal ����� As Boolean, _
    ByVal ����س� As Boolean, ByRef ���� As String) As Boolean
    
    Dim lngOutClient As Long
    
    Load Me
    
    Caption = ����
    lblInfo.Caption = ��ʾ
    txt.MaxLength = ����
    
    txt.Height = IIf(���� = 0, 1, ����) * 300
    lngOutClient = (GetSystemMetrics(SM_CYBORDER) + GetSystemMetrics(SM_CYFRAME)) * 2 * 15 + GetSystemMetrics(SM_CYSMCAPTION) * 15
    Me.Height = txt.Top + txt.Height + picCmd.Height + 200 + lngOutClient
    
    txt.Text = ����
    
    mintLen = ����
    mblnNULL = �����
    mblnEnter = ����س�
    
    
    Me.Show 1, frmParent
    
    If Not mblnOK Then ���� = "": Exit Function
    ���� = mstrText
    InputBox = True
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Not mblnNULL And txt.Text = "" Then
        MsgBox "��������" & Caption & "��", vbInformation, gstrSysName
        txt.SetFocus: Exit Sub
    End If
    If mintLen > 0 Then
        If gobjCommFun.ActualLen(txt.Text) > mintLen Then
            MsgBox Caption & "����������� " & mintLen & " ���ַ��� " & mintLen \ 2 & " �����֣�", vbInformation, gstrSysName
            txt.SetFocus: Exit Sub
        End If
    End If

    mstrText = txt.Text
    
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("';" & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    mblnOK = False
End Sub

Private Sub txt_GotFocus()
    gobjControl.TxtSelAll txt
End Sub

Private Sub txt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt.Text = "" Then
            cmdOK_Click
        ElseIf Not mblnEnter Then
            cmdOK.SetFocus
        End If
    End If
End Sub

Private Sub txt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txt.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txt.hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub
