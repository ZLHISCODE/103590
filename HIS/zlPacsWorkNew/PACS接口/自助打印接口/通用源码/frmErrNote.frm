VERSION 5.00
Begin VB.Form frmErrNote 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ע��"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4020
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   4020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox txtHelp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "frmErrNote.frx":0000
      Top             =   1605
      Width           =   3915
   End
   Begin VB.CommandButton cmdCopyScreen 
      Caption         =   "��ͼ(&S)"
      Height          =   350
      Left            =   2310
      TabIndex        =   4
      Top             =   1155
      Width           =   1080
   End
   Begin VB.CommandButton cmdOk 
      Cancel          =   -1  'True
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   1050
      TabIndex        =   3
      Top             =   1155
      Width           =   1080
   End
   Begin VB.PictureBox picS 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   2730
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   156
      TabIndex        =   6
      Top             =   1470
      Visible         =   0   'False
      Width           =   2340
   End
   Begin VB.Label lblNote 
      Caption         =   "    �����������û��Ķ�ռ�����°�װ�˲���ϵͳ�����Ĵ����ų���ռʹ�������Բ������У����貿����װ��ϵͳ��"
      Height          =   585
      Left            =   900
      TabIndex        =   2
      Top             =   465
      Width           =   3075
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblScrip 
      AutoSize        =   -1  'True
      Caption         =   "˵����"
      Height          =   180
      Left            =   900
      TabIndex        =   1
      Top             =   180
      Width           =   540
   End
   Begin VB.Label lblNumber 
      AutoSize        =   -1  'True
      Caption         =   "��ţ�"
      Height          =   180
      Left            =   2805
      TabIndex        =   0
      Top             =   210
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmErrNote.frx":006B
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmErrNote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCopyScreen_Click()
'    Call SaveScreen(txtHelp.Text, picS)
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Public Sub ShowEdit(lngErrNum As Long, strNote As String, strErrInfo As String)
'���ܣ���ʾ������ʾ����
'������lngErrNum   ������
'      strNote     ��������
'      strErrInfo  ��ϸ�Ĵ�����Ϣ
    
    lblNumber.Caption = "��ţ�" & lngErrNum
    lblNote.Caption = Space(4) & strNote
    txtHelp.Text = strErrInfo
    
    frmErrNote.Show vbModal
End Sub

