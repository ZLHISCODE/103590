VERSION 5.00
Begin VB.Form frmExamPicEdit 
   Caption         =   "frmExamPicEdit"
   ClientHeight    =   6195
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8100
   LinkTopic       =   "Form1"
   ScaleHeight     =   6195
   ScaleWidth      =   8100
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   465
      Left            =   3960
      TabIndex        =   3
      Top             =   450
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      Height          =   330
      Left            =   5445
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   90
      Width           =   1410
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   3960
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   90
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Height          =   3030
      Left            =   90
      Picture         =   "frmExamPicEdit.frx":0000
      ScaleHeight     =   2970
      ScaleWidth      =   3690
      TabIndex        =   0
      Top             =   90
      Width           =   3750
   End
End
Attribute VB_Name = "frmExamPicEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function �����Զ��滻ͼƬ(ByVal lngPara1 As Long, ByVal lngPara2 As Long) As String
    Dim strF As String
    strF = App.Path & "\TMP.jpg"
    SavePicture Me.Picture1.Picture, strF
    �����Զ��滻ͼƬ = strF
End Function

Public Function �޸��Զ��滻ͼƬ(ByVal lngPara1 As Long, ByVal lngPara2 As Long) As String
    Dim strF As String
    strF = App.Path & "\TMP.jpg"
    SavePicture Me.Picture1.Picture, strF
    �޸��Զ��滻ͼƬ = strF
End Function

Public Function ���ɲ���ͼƬ(ByRef lngPara1 As Long, ByRef lngPara2 As Long) As String
    Dim strF As String
    strF = App.Path & "\TMP.jpg"
    SavePicture Me.Picture1.Picture, strF
    ���ɲ���ͼƬ = strF
End Function

Public Function �޸�����ͼƬ(ByRef lngPara1 As Long, ByRef lngPara2 As Long) As String
    Dim strF As String
    strF = App.Path & "\TMP.jpg"
    SavePicture Me.Picture1.Picture, strF
    �޸�����ͼƬ = strF
End Function

