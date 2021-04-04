VERSION 5.00
Begin VB.Form frmSet_NBYKT 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
   Icon            =   "frmSet_NBYKT.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CheckBox chk��Ϣת�� 
      Caption         =   "ͨ��ǰ�û�������Ϣת��"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   780
      TabIndex        =   6
      Top             =   1590
      Width           =   2625
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3960
      TabIndex        =   7
      Top             =   360
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3960
      TabIndex        =   8
      Top             =   810
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3255
      Left            =   3690
      TabIndex        =   9
      Top             =   -150
      Width           =   45
   End
   Begin VB.TextBox txt·�� 
      Height          =   300
      Left            =   1125
      TabIndex        =   5
      Text            =   "/wsdl"
      Top             =   1110
      Width           =   2235
   End
   Begin VB.TextBox txt�˿ں� 
      Height          =   300
      Left            =   1125
      TabIndex        =   3
      Text            =   "8080"
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox txtServerIP 
      Height          =   300
      Left            =   1125
      TabIndex        =   1
      Text            =   "schemas.xmlsoap.org"
      Top             =   330
      Width           =   2235
   End
   Begin VB.Label lbl·�� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Ŀ¼"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   690
      TabIndex        =   4
      Top             =   1170
      Width           =   360
   End
   Begin VB.Label lbl�˿ں� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�˿ں�"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   510
      TabIndex        =   2
      Top             =   780
      Width           =   540
   End
   Begin VB.Label lblServerIP 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "IP����ַ"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   330
      TabIndex        =   0
      Top             =   390
      Width           =   720
   End
End
Attribute VB_Name = "frmSet_NBYKT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mint���� As Integer
Private mblnOK As Boolean
Private Const strFile As String = "C:\APPSOFT\NBYKT.INI"

Public Function ��������(ByVal int���� As Integer) As Boolean
    mblnOK = False
    mint���� = int����
    Me.Show 1
    �������� = mblnOK
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHand
    Dim blnOpen As Boolean
    Dim objStream As TextStream
    Dim objFileSys As New FileSystemObject
    
    If objFileSys.FileExists(strFile) = False Then Call objFileSys.CreateTextFile(strFile, True)
    Set objStream = objFileSys.OpenTextFile(strFile, ForWriting)
    blnOpen = True
    objStream.WriteLine "��ַ=" & Me.txtServerIP.Text
    objStream.WriteLine "�˿ں�=" & Me.txt�˿ں�.Text
    objStream.WriteLine "Ŀ¼=" & Me.txt·��.Text
    objStream.WriteLine "��Ϣת��=" & Me.chk��Ϣת��.Value
     objStream.WriteLine "������=NBYKT"
      objStream.WriteLine "�û�=system"
       objStream.WriteLine "����=abc123"
    objStream.Close
    
    mblnOK = True
    blnOpen = False
    
    Unload Me
    Exit Sub
errHand:
    MsgBox Err.Description
    If blnOpen Then objStream.Close
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call gobjCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim strLine As String
    Dim strName As String
    Dim strValue As String
    Dim objStream As TextStream
    Dim objFileSys As New FileSystemObject
    
    If objFileSys.FileExists(strFile) = False Then Exit Sub
    Set objStream = objFileSys.OpenTextFile(strFile, ForReading)
    Do While Not objStream.AtEndOfStream
        strLine = objStream.ReadLine
        strName = Trim(Split(strLine, "=")(0))
        strValue = Trim(Split(strLine, "=")(1))
        Select Case strName
        Case "��ַ"
            txtServerIP.Text = strValue
        Case "�˿ں�"
            txt�˿ں�.Text = strValue
        Case "Ŀ¼"
            txt·��.Text = strValue
        Case "��Ϣת��"
            Me.chk��Ϣת��.Value = Val(strValue)
        End Select
    Loop
End Sub

Private Sub txtServerIP_GotFocus()
    Call gobjControl.TxtSelAll(txtServerIP)
End Sub

Private Sub txt�˿ں�_GotFocus()
    Call gobjControl.TxtSelAll(txt�˿ں�)
End Sub

Private Sub txt·��_GotFocus()
    Call gobjControl.TxtSelAll(txt·��)
End Sub
