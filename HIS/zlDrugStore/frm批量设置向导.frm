VERSION 5.00
Begin VB.Form frm���������� 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����������"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6885
   Icon            =   "frm����������.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   6885
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��"
      Height          =   350
      Left            =   5640
      TabIndex        =   13
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��"
      Height          =   350
      Left            =   4200
      TabIndex        =   12
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Frame fra�´��� 
      Caption         =   "�滻Ϊ�´���"
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   6615
      Begin VB.ComboBox cbo���� 
         Height          =   300
         Left            =   120
         TabIndex        =   6
         Text            =   "Combo1"
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame fra 
      Height          =   45
      Left            =   0
      TabIndex        =   2
      Top             =   1850
      Width           =   6855
   End
   Begin VB.Frame fra���� 
      Caption         =   "ѡ�񴦷�����"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   6615
      Begin VB.CheckBox chk���� 
         Caption         =   "����II��"
         Height          =   255
         Index           =   5
         Left            =   5450
         TabIndex        =   11
         Top             =   320
         Width           =   1095
      End
      Begin VB.CheckBox chk���� 
         Caption         =   "����I��"
         Height          =   255
         Index           =   4
         Left            =   4192
         TabIndex        =   10
         Top             =   320
         Width           =   975
      End
      Begin VB.CheckBox chk���� 
         Caption         =   "����"
         Height          =   255
         Index           =   3
         Left            =   3174
         TabIndex        =   9
         Top             =   320
         Width           =   735
      End
      Begin VB.CheckBox chk���� 
         Caption         =   "����"
         Height          =   255
         Index           =   2
         Left            =   2156
         TabIndex        =   8
         Top             =   320
         Width           =   735
      End
      Begin VB.CheckBox chk���� 
         Caption         =   "����"
         Height          =   255
         Index           =   1
         Left            =   1138
         TabIndex        =   7
         Top             =   320
         Width           =   735
      End
      Begin VB.CheckBox chk���� 
         Caption         =   "��ͨ"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   320
         Width           =   735
      End
   End
   Begin VB.Frame fra�ִ��� 
      Caption         =   "ѡ���ִ���"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      Begin VB.CheckBox chk���� 
         Caption         =   "��̬��Ӵ���"
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   320
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frm����������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngҩ��ID As Long
Private mstr���� As String
Private mstrConWin As String   'ѡ����ִ���
Private mstr���� As String
Private mstr�´��� As String
Private mstr�ִ��� As String

Private Sub LoadOldWin()
    Dim i As Integer
    
    If mstr���� = "" Then Exit Sub
    
    Me.chk����(0).Caption = Split(mstr����, ",")(0)
    chk����(0).Width = 150 + LenB(chk����(0).Caption) * 128
    For i = 1 To UBound(Split(mstr����, ",")) - 1
        Load chk����(i)
        chk����(i).Visible = True
        chk����(i).Caption = Split(mstr����, ",")(i)
        chk����(i).Width = 150 + LenB(chk����(i - 1).Caption) * 128
        chk����(i).Left = chk����(i - 1).Left + chk����(i - 1).Width + 100
    Next
End Sub
Private Sub LoadNewWin()
    Dim strSql As String
    Dim rsRecord As Recordset
    
    On Error GoTo errHandle
    
    strSql = "select ����,���� from ��ҩ���� where ҩ��id=[1] and �ϰ��=1"
    If mstr�ִ��� <> "" Then strSql = strSql & " And ����<>[2] "
    Set rsRecord = zldatabase.OpenSQLRecord(strSql, "Init����", mlngҩ��ID, mstr�ִ���)
    
    If Not (rsRecord Is Nothing) Then
        Do While Not rsRecord.EOF
            Me.cbo����.AddItem rsRecord!����
            mstr���� = mstr���� & rsRecord!���� & ","
            rsRecord.MoveNext
        Loop
    End If
    
    If mstr�ִ��� <> "" Then mstr���� = mstr�ִ���
    
    Me.cbo����.ListIndex = 0
    Exit Sub
errHandle:
    If errcenter() = 1 Then Resume
    Call saveerrlog
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOK_Click()
    Dim i As Integer
    For i = 0 To Me.chk����.UBound
        If Me.chk����(i).Value = 1 Then
            mstrConWin = mstrConWin & Me.chk����(i).Caption & ","
        End If
    Next
    
    For i = 0 To Me.chk����.UBound
        If Me.chk����(i).Value = 1 Then
            mstr���� = mstr���� & Me.chk����(i).Caption & ","
        End If
    Next
    
    mstr�´��� = Me.cbo����.Text
    
    Unload Me
End Sub

Private Sub Form_Load()
    LoadNewWin
    LoadOldWin
End Sub

Public Sub ShowME(ByVal lngҩ��ID As Long, ByRef strConWin As String, ByRef str���� As String, ByRef str�´��� As String, Optional str�ִ��� As String = "")
    mlngҩ��ID = lngҩ��ID
    mstr�ִ��� = str�ִ���
    Me.Show 1
    
    strConWin = mstrConWin
    str���� = mstr����
    str�´��� = mstr�´���
    
    
    mstrConWin = ""
    mstr���� = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstr���� = ""
End Sub
