VERSION 5.00
Begin VB.Form frmErrNote 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ע��"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4020
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   4020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdOk 
      Cancel          =   -1  'True
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   1020
      TabIndex        =   0
      Top             =   1155
      Width           =   1080
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   2280
      TabIndex        =   1
      Top             =   1155
      Width           =   1080
   End
   Begin VB.TextBox txtHelp 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "frmErrNote.frx":0000
      Top             =   1605
      Width           =   3885
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   210
      Picture         =   "frmErrNote.frx":006B
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblNumber 
      AutoSize        =   -1  'True
      Caption         =   "��ţ�"
      Height          =   180
      Left            =   2775
      TabIndex        =   5
      Top             =   210
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label lblScrip 
      AutoSize        =   -1  'True
      Caption         =   "˵����"
      Height          =   180
      Left            =   870
      TabIndex        =   4
      Top             =   180
      Width           =   540
   End
   Begin VB.Label lblNote 
      Caption         =   "    �����������û��Ķ�ռ�����°�װ�˲���ϵͳ�����Ĵ����ų���ռʹ�������Բ������У����貿����װ��ϵͳ��"
      Height          =   585
      Left            =   870
      TabIndex        =   3
      Top             =   465
      Width           =   3075
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmErrNote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdHelp_Click()
    Height = Height + txtHelp.Height + 100
    cmdHelp.Enabled = False
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

