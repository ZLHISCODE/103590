VERSION 5.00
Begin VB.Form Frm�ر� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�ر�"
   ClientHeight    =   2412
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4284
   Icon            =   "Frm�ر�.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2412
   ScaleWidth      =   4284
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton Cmdȷ�� 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   1560
      TabIndex        =   1
      Top             =   1950
      Width           =   1100
   End
   Begin VB.CommandButton Cmdȡ�� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   2880
      TabIndex        =   2
      Top             =   1950
      Width           =   1100
   End
   Begin VB.ComboBox Cbo�ر� 
      Height          =   300
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   720
      Width           =   2685
   End
   Begin VB.Label LblNote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   1080
      TabIndex        =   4
      Top             =   1110
      Width           =   2625
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ϣ���������ʲô:"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1080
      TabIndex        =   3
      Top             =   420
      Width           =   1530
   End
   Begin VB.Image Image1 
      Height          =   192
      Left            =   300
      Picture         =   "Frm�ر�.frx":27A2
      Top             =   240
      Width           =   192
   End
End
Attribute VB_Name = "Frm�ر�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mintStyle As Integer         '-1=�ر�;0=ע��

Public Function ShowMe(ByRef intStyle As Integer) As Boolean
'������
'���أ�intStyle=-1=�ر�;0=ע��
'      �Ƿ���ȷ���ر�
    mintStyle = 0
    Me.Show vbModal
    ShowMe = mblnOK
    intStyle = mintStyle
End Function

Private Sub Cbo�ر�_Click()
    With Cbo�ر�
        Select Case .ItemData(.ListIndex)
        Case -1
            LblNote.Caption = "�ر�ϵͳ���ص�Windows���档"
            mintStyle = -1
        Case 0
            LblNote.Caption = "�������û���������µ�¼��"
            mintStyle = 0
        End Select
    End With
End Sub

Private Sub Cmdȡ��_Click()
    mblnOK = False
    mintStyle = 0
    Unload Me
End Sub

Private Sub Cmdȷ��_Click()
    mblnOK = True
    With Cbo�ر�
        Select Case .ItemData(.ListIndex)
        Case -1
            mintStyle = -1
        Case 0
            mintStyle = 0
        End Select
    End With
    Unload Me
End Sub

Private Sub Form_Load()
    With Cbo�ر�
        .Clear
        .AddItem "�ر�ϵͳ"
        .ItemData(.NewIndex) = -1
        .AddItem "ע��"
        .ItemData(.NewIndex) = 0
        .ListIndex = 0
    End With
End Sub
