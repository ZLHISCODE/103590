VERSION 5.00
Begin VB.Form frmCaseTendBodyPara 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���µ�ѡ��"
   ClientHeight    =   1485
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3525
   Icon            =   "frmCaseTendBodyPara.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   3525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame Fra 
      Height          =   750
      Left            =   135
      TabIndex        =   0
      Top             =   90
      Width           =   3225
      Begin VB.CheckBox chk 
         Caption         =   "���µ����ʱ����ʾƤ�Խ��"
         Height          =   210
         Index           =   1
         Left            =   270
         TabIndex        =   1
         Top             =   315
         Width           =   2790
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   1095
      TabIndex        =   2
      Top             =   990
      Width           =   1100
   End
   Begin VB.CommandButton cmdCanc 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   2250
      TabIndex        =   3
      Top             =   990
      Width           =   1100
   End
End
Attribute VB_Name = "frmCaseTendBodyPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mfrmMain As Object
Private mblnOK As Boolean
Private mstrPrivs As String

Public Function ShowPara(ByVal frmMain As Object, ByVal strPrivs As String) As Boolean
    
    mblnOK = False
    mstrPrivs = strPrivs
    
    Set mfrmMain = frmMain
   
    chk(1).Value = Val(zlDatabase.GetPara("���µ���ʾƤ�Խ��", glngSys, 1255, "0", Array(chk(1)), InStr(mstrPrivs, "����ѡ������") > 0))
    
    Me.Show 1, mfrmMain
    ShowPara = mblnOK
End Function

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmdOK_Click()
    Call zlDatabase.SetPara("���µ���ʾƤ�Խ��", chk(1).Value, glngSys, 1255)
    mblnOK = True
    Unload Me
End Sub

Private Sub cmdCanc_Click()
    Unload Me
End Sub
