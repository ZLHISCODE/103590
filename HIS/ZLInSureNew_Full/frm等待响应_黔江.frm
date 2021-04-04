VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm等待响应_黔江 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "等待响应..."
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5355
   ControlBox      =   0   'False
   Icon            =   "frm等待响应_黔江.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   5355
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   1560
      Top             =   1530
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "中断"
      Height          =   375
      Left            =   4095
      TabIndex        =   4
      Top             =   1560
      Width           =   1035
   End
   Begin VB.Timer TimeAvi 
      Interval        =   50
      Left            =   2700
      Top             =   690
   End
   Begin VB.Timer TimeSearch 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2070
      Top             =   660
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Left            =   30
      Picture         =   "frm等待响应_黔江.frx":000C
      ScaleHeight     =   5
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   355
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1140
      Width           =   5325
   End
   Begin MSComCtl2.Animation Avi 
      Height          =   765
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   1349
      _Version        =   393216
      AutoPlay        =   -1  'True
      BackColor       =   -2147483643
      FullWidth       =   61
      FullHeight      =   51
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "因中心方有已结算的可能,中断操作可能造成双方帐目不平"
      Height          =   180
      Left            =   382
      TabIndex        =   3
      Top             =   1260
      Width           =   4590
   End
   Begin VB.Label LblNote 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "  已提交请求，正在等待医保服务器响应..."
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1380
      TabIndex        =   0
      Top             =   480
      Width           =   3510
   End
End
Attribute VB_Name = "frm等待响应_黔江"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnReturn As Boolean           '是否成功返回
Private mstrFileName As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call Avi_Play
    TimeSearch.Enabled = True
    TimeSearch.Interval = 5000
    mblnReturn = False
End Sub

Private Sub Avi_Play()
    On Error Resume Next
    With Avi
        .Open gstrAviPath & "\FINDFILE.AVI"
        .AutoPlay = True
        .Play
    End With
End Sub

Private Sub Avi_Stop()
    Avi.Stop
End Sub

Public Function ShowME(strFileName As String) As Boolean
    mblnReturn = False
    mstrFileName = strFileName
    Me.Show 1
    ShowME = mblnReturn
End Function

Private Sub Timer1_Timer()
    Unload Me
End Sub

Private Sub TimeSearch_Timer()
    Dim Temp
    On Error Resume Next
    Temp = FileDateTime(mstrFileName)
    If Err = 0 Then
        mblnReturn = True
        Timer1.Enabled = True
        TimeSearch.Enabled = False
    Else
        Timer1.Enabled = False
    End If
    On Error GoTo 0
End Sub

Private Sub TimeAvi_Timer()
    Static i As Long
    i = i + 20
    If i >= Picture1.ScaleWidth Then i = 1
    
    Picture1.PaintPicture Picture1.Picture, i, 0, Picture1.ScaleWidth - i, Picture1.ScaleHeight, 0, 0, Picture1.ScaleWidth - i, Picture1.ScaleHeight
    Picture1.PaintPicture Picture1.Picture, 0, 0, i, Picture1.ScaleHeight, Picture1.ScaleWidth - i, 0, i, Picture1.ScaleHeight
End Sub

