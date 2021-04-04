VERSION 5.00
Begin VB.Form frmAdviceReprotBrowseShowPic 
   Caption         =   "图片查看"
   ClientHeight    =   6405
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8160
   Icon            =   "frmAdviceReprotBrowseShowPic.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6405
   ScaleWidth      =   8160
   StartUpPosition =   1  '所有者中心
   Begin VB.Image imgMain 
      Height          =   1665
      Left            =   1920
      Stretch         =   -1  'True
      ToolTipText     =   "双击查看大图"
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
'编    码:蔡青松
'编码时间:2019-08-01
'功    能:  查看大图
'入    参:
'           objFrm      调用窗体
'           strPath     图片路径
'出    参:
'返    回:
'调整影响:
'调用注意:
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
