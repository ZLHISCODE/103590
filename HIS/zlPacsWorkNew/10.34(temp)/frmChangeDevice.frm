VERSION 5.00
Begin VB.Form frmChangeDevice 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�����豸����"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3795
   Icon            =   "frmChangeDevice.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   3795
   StartUpPosition =   1  '����������
   Begin VB.ComboBox cboDeviceType 
      Height          =   300
      ItemData        =   "frmChangeDevice.frx":0CCA
      Left            =   960
      List            =   "frmChangeDevice.frx":0CDA
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��"
      Height          =   350
      Left            =   2160
      TabIndex        =   1
      Top             =   1080
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��"
      Height          =   350
      Left            =   480
      TabIndex        =   0
      Top             =   1080
      Width           =   1100
   End
   Begin VB.Label Label1 
      Caption         =   "�ѵ�ǰ�豸���͸�����"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "frmChangeDevice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strDeviceType As String

Public Sub ShowMe(strOldType As String, Parentform As Object)
    Me.Label1.Caption = "��ǰ�豸�����ǣ�" & strOldType & "�������ɣ�"
    Me.cboDeviceType.ListIndex = 0
    Me.Show 1, Parentform
End Sub

Private Sub cmdCancel_Click()
    strDeviceType = ""
    Unload Me
End Sub

Private Sub cmdOK_Click()
    strDeviceType = cboDeviceType.Text
    Unload Me
End Sub

Private Sub Form_Load()
    strDeviceType = ""
End Sub
