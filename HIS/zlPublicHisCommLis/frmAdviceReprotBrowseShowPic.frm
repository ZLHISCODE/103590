VERSION 5.00
Begin VB.Form frmAdviceReprotBrowseShowPic 
   Caption         =   "ͼƬ�鿴"
   ClientHeight    =   6405
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8160
   Icon            =   "frmAdviceReprotBrowseShowPic.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6405
   ScaleWidth      =   8160
   StartUpPosition =   1  '����������
   Begin VB.Image imgMain 
      Height          =   1665
      Left            =   1920
      Stretch         =   -1  'True
      ToolTipText     =   "˫���鿴��ͼ"
      Top             =   1980
      Width           =   2025
   End
End
Attribute VB_Name = "frmAdviceReprotBrowseShowPic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2019-08-01
'��    ��:  �鿴��ͼ
'��    ��:
'           objFrm      ���ô���
'           strPath     ͼƬ·��
'��    ��:
'��    ��:
'����Ӱ��:
'����ע��:
'---------------------------------------------------------------------------------------
Public Sub showMe(objFrm As Object, ByVal strPath As String)
    If strPath <> "" Then
        Me.imgMain.Picture = LoadPicture(strPath)
        Me.Show vbModal, objFrm
    Else
        Unload Me
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With imgMain
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight
    End With
End Sub
