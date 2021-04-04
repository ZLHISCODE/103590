VERSION 5.00
Begin VB.Form frmClose 
   BackColor       =   &H8000000D&
   BorderStyle     =   0  'None
   Caption         =   "挂号成功"
   ClientHeight    =   4200
   ClientLeft      =   0
   ClientTop       =   -15
   ClientWidth     =   6285
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmClose.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Time 
      Interval        =   2000
      Left            =   0
      Top             =   0
   End
   Begin zl9NewQuery.ctlButton ctlClose 
      Height          =   720
      Left            =   2385
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3225
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   1270
      Caption         =   " 关闭 "
      AutoSize        =   0   'False
      ButtonHeight    =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "挂号成功"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   2025
      TabIndex        =   4
      Top             =   180
      Width           =   1860
   End
   Begin VB.Label lblMsg 
      BackStyle       =   0  'Transparent
      Caption         =   "  请取走挂号凭据，根据凭单到指定科室就诊！如凭单打坏或不出凭单，请立即与挂号处联系。谢谢！"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1845
      Left            =   2235
      TabIndex        =   3
      Top             =   1305
      Width           =   3990
   End
   Begin VB.Label LblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   525
      Left            =   105
      TabIndex        =   1
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Image Imgbak 
      Height          =   2055
      Left            =   165
      Picture         =   "frmClose.frx":1E26
      Stretch         =   -1  'True
      Top             =   810
      Width           =   1920
   End
   Begin VB.Label Lblname 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "一一一"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   2250
      TabIndex        =   0
      Top             =   945
      Width           =   765
   End
End
Attribute VB_Name = "frmClose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrName As String
Private mstrNO As String
Private mstrChargeNo As String

Dim mlngTime As Long

Public Function ShowForm(ByVal frmMain As Object, ByVal strName As String, ByVal strNo As String, Optional ByVal strChargeNo As String) As Boolean
    '******************************************************************************************************************
    '
    '
    '
    '******************************************************************************************************************
    
    mstrName = strName
    mstrNO = strNo
    mstrChargeNo = strChargeNo
    
    mlngTime = Val(GetPara("密码验证窗体停留时间")) / 2
    LblName.Caption = mstrName + "同志："
    lblInfo.Caption = "挂号单号：" + mstrNO
    
    If mstrChargeNo = "" Then
        lblMsg.Caption = "  请取走挂号凭据，根据凭单到指定科室就诊！" & vbCrLf & "  如凭单打坏或不出凭单，请立即与挂号处联系。谢谢！"
    Else
        lblMsg.Caption = "  请直接到指定科室等候就诊！" & vbCrLf & "   谢谢！"
    End If
    If Not frmMain Is frmFreeRegist Then
        ctlClose.Picture = frmselectinfo.ilsImage.ListImages("close")
    Else
        ctlClose.ShowPicture = False
    End If
    Me.Show 1, frmMain
    
End Function

Private Sub ctlClose_CommandClick()
    Unload Me
End Sub

Private Sub Form_Load()
    If Dir(App.Path & "\图形\挂号成功背景.pic") <> "" Then
        Imgbak.Picture = LoadPicture(App.Path & "\图形\挂号成功背景.pic")
    End If
End Sub

Private Sub Form_Paint()
    Call DrawColorToColor(Me, Me.BackColor, &HFFC0C0, , True)
End Sub

Private Sub Time_Timer()
   On Error Resume Next
   mlngTime = mlngTime - 1
   If mlngTime = 0 Then Unload Me
End Sub
