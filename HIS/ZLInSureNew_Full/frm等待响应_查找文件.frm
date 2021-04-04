VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm等待响应_查找文件 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "等待服务器返回应答文件"
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5325
   Icon            =   "frm等待响应_查找文件.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Left            =   0
      Picture         =   "frm等待响应_查找文件.frx":000C
      ScaleHeight     =   5
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   355
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1020
      Width           =   5325
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4020
      TabIndex        =   0
      Top             =   1200
      Width           =   1100
   End
   Begin VB.Timer TimeSearch 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2040
      Top             =   540
   End
   Begin VB.Timer TimeAvi 
      Interval        =   50
      Left            =   2040
      Top             =   540
   End
   Begin MSComCtl2.Animation Avi 
      Height          =   765
      Left            =   90
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   1349
      _Version        =   393216
      AutoPlay        =   -1  'True
      BackColor       =   -2147483643
      FullWidth       =   61
      FullHeight      =   51
   End
   Begin VB.Label LblNote 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "  已提交请求，正在等待医保服务器响应..."
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1350
      TabIndex        =   3
      Top             =   360
      Width           =   3510
   End
End
Attribute VB_Name = "frm等待响应_查找文件"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnReturn As Boolean           '是否成功返回
Private mblnFind As Boolean             '是否找到文件
Private mintWait As Integer             '找到文件、读文件之间的间隔秒数
Private mintWaited As Integer           '从找到文件起累计已等待时间
Private mstrFile As String              '

Private Sub cmdCancel_Click()
    If mblnFind Then
        If MsgBox("    已找到应答文件，但医保商要求再等待" & mintWait & "秒后读文件，你确认要退出吗？" & _
            vbCrLf & "医保交易已完成，如果退出，HIS交易将不会保存，这会造成医保与医院报表不一符！", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    
    Unload Me
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Form_Load()
    Call Avi_Play
    TimeSearch.Enabled = True
End Sub

Private Sub Avi_Play()
    On Error Resume Next
    With Avi
        .Open mstrFile
        .AutoPlay = True
        .Play
    End With
End Sub

Private Sub Avi_Stop()
    Avi.Stop
End Sub

Public Function ShowME(ByVal strFile As String, Optional ByVal intWait As Integer = 0) As Boolean
    mblnReturn = False
    mstrFile = strFile
    mintWait = intWait
    
    Me.Show 1
    
    ShowME = mblnReturn
End Function

Private Sub TimeSearch_Timer()
    If Not mblnFind Then
        If Not SearchFile Then Exit Sub
        mblnFind = True
        mintWaited = 0
    Else
        mintWaited = mintWaited + 1
    End If
    
    If mintWaited < mintWait Then Exit Sub
    
    TimeSearch.Enabled = False
    mblnReturn = True
    
    Unload Me
    Exit Sub
End Sub

Private Sub TimeAvi_Timer()
    Static i As Long
    i = i + 20
    If i >= Picture1.ScaleWidth Then i = 1
    
    Picture1.PaintPicture Picture1.Picture, i, 0, Picture1.ScaleWidth - i, Picture1.ScaleHeight, 0, 0, Picture1.ScaleWidth - i, Picture1.ScaleHeight
    Picture1.PaintPicture Picture1.Picture, 0, 0, i, Picture1.ScaleHeight, Picture1.ScaleWidth - i, 0, i, Picture1.ScaleHeight
End Sub

Private Function SearchFile() As Boolean
    Dim objFileSys As New FileSystemObject
    SearchFile = False
    
    If Not objFileSys.FileExists(mstrFile) Then Exit Function
    SearchFile = True
End Function
