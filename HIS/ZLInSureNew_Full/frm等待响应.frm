VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm等待响应 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "等待响应..."
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5355
   ControlBox      =   0   'False
   Icon            =   "frm等待响应.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   5355
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer TimeAvi 
      Interval        =   50
      Left            =   2070
      Top             =   660
   End
   Begin VB.Timer TimeSearch 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2070
      Top             =   660
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4050
      TabIndex        =   3
      Top             =   1320
      Width           =   1100
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Left            =   30
      Picture         =   "frm等待响应.frx":000C
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
Attribute VB_Name = "frm等待响应"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mint操作方式 As Integer         '操作方式
Private mint请求目的 As Integer         '请求目的
Private mlng病人ID As Long              '病人ID
Private mlng结帐ID As Long              '结帐ID
Private mint险类 As Integer
Private mblnReturn As Boolean           '是否成功返回
Private mstrFile As String              '

Private Sub cmdCancel_Click()
    On Error GoTo ErrHand
    Dim objFileSys As New FileSystemObject
    
    TimeSearch.Enabled = False
    '删除请求文件
    If objFileSys.FileExists(mstrPath_福建巨龙 & mint险类 & "\" & mstrRequest_福建巨龙) Then
        Call objFileSys.DeleteFile(mstrPath_福建巨龙 & mint险类 & "\" & mstrRequest_福建巨龙, True)
    End If
    '先检查应答文件是否存在，如果存在则提示
    If objFileSys.FileExists(mstrPath_福建巨龙 & mint险类 & "\" & mstrReply_福建巨龙) Then
        If MsgBox("服务器已经响应请求，你确定要退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            TimeSearch.Enabled = True
            Exit Sub
        End If
        Call objFileSys.DeleteFile(mstrPath_福建巨龙 & mint险类 & "\" & mstrReply_福建巨龙, True)
    End If
    
    Unload Me
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Form_Activate()
    Dim objFileSys As New FileSystemObject
    
    '如果指定目录不存在，则创建
    If mint操作方式 <> 操作方式.验证 Then
        If Not objFileSys.FolderExists(mstrPath_福建巨龙 & mint险类) Then objFileSys.CreateFolder (mstrPath_福建巨龙 & mint险类)
        LblNote.Caption = "  正在向医保服务器发送请求..."
        If SendRequest(mint操作方式, mint请求目的, mlng病人ID, mlng结帐ID, mint险类) = False Then
            mblnReturn = False
            Unload Me
            Exit Sub
        End If
    End If
    
    TimeSearch.Enabled = True
    LblNote.Caption = "  已提交请求，正在等待医保服务器响应..."
End Sub

Private Sub Form_Load()
    Dim strCaption As String
    mstrFile = gstrAviPath & "\" & mstrSearch_福建巨龙
    
    LblNote.Caption = "  正在检查相关环境..."
    
    '分析标题
    Select Case mint请求目的
    Case 请求目的.申请
        Select Case mint操作方式
        Case 操作方式.挂号
            strCaption = "挂号请求..."
        Case 操作方式.收费
            strCaption = "门诊收费请求..."
        Case 操作方式.结帐
            strCaption = "住院结算请求..."
        Case 操作方式.入院
            strCaption = "入院请求..."
        Case 操作方式.出院
            strCaption = "出院请求..."
        Case 操作方式.登录
            strCaption = "登录请求..."
        End Select
    Case 请求目的.冲销
        Select Case mint操作方式
        Case 操作方式.挂号
            strCaption = "挂号冲销请求..."
        Case 操作方式.收费
            strCaption = "门诊收费冲销请求..."
        Case 操作方式.结帐
            strCaption = "住院结算冲销请求..."
        Case 操作方式.入院
            strCaption = "撤销入院请求..."
        Case 操作方式.出院
            strCaption = "撤销出院请求..."
        Case 操作方式.登录
            strCaption = "退出请求..."
        End Select
    Case 请求目的.刷卡
        strCaption = "请刷卡..."
    End Select
    
    Me.Caption = strCaption
    Call Avi_Play
    mblnReturn = False
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

Public Function ShowME(ByVal intInsure As Integer, ByVal 操作 As Integer, ByVal 目的 As Integer, Optional ByVal 病人ID As Long = 0, _
        Optional ByVal 结帐ID As Long = 0) As Boolean
    mint操作方式 = 操作
    mint请求目的 = 目的
    mlng病人ID = 病人ID
    mlng结帐ID = 结帐ID
    mint险类 = intInsure
    Me.Show 1
    ShowME = mblnReturn
End Function

Private Sub TimeSearch_Timer()
    Dim intResult As Integer
    intResult = SearchFile
    If intResult = 0 Then Exit Sub
    
    mblnReturn = (intResult = 1)
    If mblnReturn Then
        If mint请求目的 = 请求目的.刷卡 And (mint操作方式 = 操作方式.挂号 _
        Or mint操作方式 = 操作方式.收费 Or mint操作方式 = 操作方式.入院 Or mint操作方式 = 操作方式.验证) Then
            mblnReturn = frmIdentify福建巨龙.ShowCard(获取病人ID(mint险类), mint险类)
        End If
    End If
    
    TimeSearch.Enabled = False
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

Private Function SearchFile() As Integer
    Dim objFileSys As New FileSystemObject
    SearchFile = False
    
    If mint操作方式 <> 操作方式.验证 Then
        If Not objFileSys.FileExists(mstrPath_福建巨龙 & mint险类 & "\" & mstrReply_福建巨龙) Then Exit Function
        SearchFile = AnalyseReply(mint操作方式, mint请求目的, mint险类)
    Else
        SearchFile = 1
    End If
End Function
