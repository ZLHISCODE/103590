VERSION 5.00
Begin VB.Form frmErrAsk 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��ʾ"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4395
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   4395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox txtHelp 
      Height          =   1260
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "frmErrAsk.frx":0000
      Top             =   1845
      Width           =   4275
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   3120
      TabIndex        =   6
      Top             =   1380
      Width           =   900
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   2175
      TabIndex        =   5
      Top             =   1380
      Width           =   900
   End
   Begin VB.CommandButton cmdRetry 
      Caption         =   "����(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   1230
      TabIndex        =   4
      Top             =   1380
      Width           =   900
   End
   Begin VB.Label lblAsk 
      AutoSize        =   -1  'True
      Caption         =   "����һ����"
      Height          =   180
      Left            =   975
      TabIndex        =   3
      Top             =   1050
      Width           =   1080
   End
   Begin VB.Label lblNote 
      Caption         =   "    �����������û��Ķ�ռ�����°�װ�˲���ϵͳ�����Ĵ����ų���ռʹ�������Բ������У����貿����װ��ϵͳ��"
      Height          =   585
      Left            =   975
      TabIndex        =   2
      Top             =   360
      Width           =   3390
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblScrip 
      AutoSize        =   -1  'True
      Caption         =   "˵����"
      Height          =   180
      Left            =   975
      TabIndex        =   1
      Top             =   150
      Width           =   540
   End
   Begin VB.Label lblNumber 
      AutoSize        =   -1  'True
      Caption         =   "��ţ�"
      Height          =   180
      Left            =   3150
      TabIndex        =   0
      Top             =   150
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   285
      Picture         =   "frmErrAsk.frx":0065
      Top             =   165
      Width           =   480
   End
End
Attribute VB_Name = "frmErrAsk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytReturn As Byte

Private Sub cmdCancel_Click()
    mbytReturn = 0
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    Height = Height + txtHelp.Height + 100
    cmdHelp.Enabled = False
End Sub

Private Sub cmdRetry_Click()
    mbytReturn = 1
    Unload Me
End Sub

Public Function ShowEdit(lngErrNum As Long, strNote As String, strErrInfo As String) As Byte
'���ܣ���ʾ������ʾ���ڣ�����ѡ������
'������lngErrNum   ������
'      strNote     ��������
'      strErrInfo  ��ϸ�Ĵ�����Ϣ
'���أ���һ����������ʾ��1-���ԣ�0-ȡ��
    mbytReturn = 0
        
    lblNumber.Caption = "��ţ�" & lngErrNum
    lblNote.Caption = Space(4) & strNote
    txtHelp.Text = strErrInfo
    
    frmErrAsk.Show vbModal
    ShowEdit = mbytReturn
End Function
