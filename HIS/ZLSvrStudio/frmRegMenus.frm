VERSION 5.00
Begin VB.Form frmRegMenus 
   Caption         =   "�û�ע�ḽ�Ӵ���"
   ClientHeight    =   2400
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   2895
   Icon            =   "frmRegMenus.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   2895
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Menu TrackMenu 
      Caption         =   "�����˵�"
      Begin VB.Menu MnuDeleteCur 
         Caption         =   "ɾ����ǰ��־(&D)"
      End
      Begin VB.Menu MnuDeleteAll 
         Caption         =   "ɾ��������־(&A)"
      End
   End
End
Attribute VB_Name = "frmRegMenus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Bln��־ As Boolean                               'Ϊ��,������־����;���������־����
Public FrmObj As Form

Private Sub MnuDeleteAll_Click()
    Call DeleteAllLog(FrmObj, Bln��־)
End Sub

Private Sub MnuDeleteCur_Click()
    Call DeleteCurLog(FrmObj, Bln��־)
End Sub
