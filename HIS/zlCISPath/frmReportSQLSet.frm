VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmReportSQLSet 
   Caption         =   "编辑SQL"
   ClientHeight    =   8175
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11835
   Icon            =   "frmReportSQLSet.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8175
   ScaleWidth      =   11835
   StartUpPosition =   1  '所有者中心
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picInfo 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1460
      Left            =   0
      ScaleHeight     =   1455
      ScaleWidth      =   11835
      TabIndex        =   4
      Top             =   0
      Width           =   11835
      Begin VB.Label lblParaRemark 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmReportSQLSet.frx":6852
         Height          =   540
         Left            =   1080
         TabIndex        =   7
         Top             =   795
         Width           =   10170
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   240
         Picture         =   "frmReportSQLSet.frx":69A5
         Top             =   120
         Width           =   480
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   2
         X1              =   0
         X2              =   11760
         Y1              =   1450
         Y2              =   1450
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   3
         X1              =   0
         X2              =   11760
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label lblNote 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmReportSQLSet.frx":6CAF
         Height          =   360
         Left            =   1080
         TabIndex        =   6
         Top             =   360
         Width           =   10245
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单元格取数SQL编辑"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   1095
         TabIndex        =   5
         Top             =   120
         Width           =   1680
      End
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F0F4E4&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   11835
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   7500
      Width           =   11835
      Begin VB.CommandButton cmdOK 
         Caption         =   "保存(&O)"
         Height          =   350
         Left            =   9360
         TabIndex        =   1
         Top             =   240
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   10560
         TabIndex        =   2
         Top             =   240
         Width           =   1100
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   0
         X2              =   11760
         Y1              =   30
         Y2              =   30
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   0
         X2              =   11760
         Y1              =   45
         Y2              =   45
      End
   End
   Begin RichTextLib.RichTextBox rtbSQL 
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   10398
      _Version        =   393217
      TextRTF         =   $"frmReportSQLSet.frx":6D8A
   End
End
Attribute VB_Name = "frmReportSQLSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrSQL文本 As String
Private mlng行号 As Long

Private Sub cmdCancel_Click()
    Dim strSQL文本 As String
    
    strSQL文本 = Trim(rtbSQL.Text)
    If strSQL文本 <> mstrSQL文本 Then
        If MsgBox("修改的SQL未保存，你确定要放弃本次修改的内容吗？", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
    
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strSQL文本 As String, strSQL As String
    
    strSQL文本 = Trim(rtbSQL.Text)
    If strSQL文本 = "" And mstrSQL文本 <> "" Then
        If MsgBox("你确定要清除当前单元格读取数据的SQL吗？", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
        
    On Error GoTo errH
    strSQL = "Zl_路径报表结构_Update(1," & mlng行号 & ",'" & strSQL文本 & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        
    mstrSQL文本 = strSQL文本
    Unload Me
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Dim strParaPatiAndPage As String
    Dim strParaPati As String
    Dim strParaPatiNum As String
    Dim strParaPathID As String
    Dim strParaBeginDate As String
    Dim strParaEndDate As String
    Dim strParaMoney As String
    Dim strDays As String
    
    '8种参数分别对应的参数说明
    strParaPatiAndPage = "病人及主页ID列表，如：“病人ID1:主页ID1,病人ID2:主页ID2,……”。"
    strParaPati = "病人列表，如：“病人ID1,病人ID2,病人ID3,……”。"
    strParaPatiNum = "满足条件的病人人数。"
    strParaPathID = "当前选择的路径的ID。"
    strParaBeginDate = "当前报表的统计期间的开始时间。"
    strParaEndDate = "当前报表的统计期间的结束时间。"
    strParaMoney = "本次统计的所有病人的费用合计。"
    strDays = "本次统计的所有病人的住院天数合计。"

    rtbSQL.Text = mstrSQL文本
    lblParaRemark.Visible = True
    Select Case mlng行号
        
        Case 2, 3, 6, 7, 14, 15, 16, 24
            lblParaRemark.Caption = "参数[1]:" & strParaPatiAndPage
            picInfo.Height = lblNote.Top + lblNote.Height + 130 + 200
            rtbSQL.Top = rtbSQL.Top - 610 + 200
            rtbSQL.Height = rtbSQL.Height + 610 - 200
            Line1(3).Y1 = lblNote.Top + lblNote.Height + 110 + 200
            Line1(3).Y2 = lblNote.Top + lblNote.Height + 110 + 200
        Case 5, 8, 9, 11, 12, 13, 23, 25
            lblParaRemark.Caption = "参数[1]:" & strParaPatiAndPage & vbCrLf & "参数[2]:" & strParaPatiNum
            picInfo.Height = lblNote.Top + lblNote.Height + 130 + 400
            rtbSQL.Top = rtbSQL.Top - 610 + 400
            rtbSQL.Height = rtbSQL.Height + 610 - 400
            Line1(3).Y1 = lblNote.Top + lblNote.Height + 110 + 400
            Line1(3).Y2 = lblNote.Top + lblNote.Height + 110 + 400
        Case 10
            lblParaRemark.Caption = "参数[1]:" & strParaPati & vbCrLf & "参数[2]:" & strParaPatiAndPage & vbCrLf & "参数[3]:" & strParaPatiNum
        Case 18
            lblParaRemark.Caption = "参数[1]:" & strParaPathID & vbCrLf & "参数[2]:" & strParaBeginDate & vbCrLf & "参数[3]:" & strParaEndDate
        Case 19, 20, 21
            lblParaRemark.Caption = "参数[1]:" & strParaPatiAndPage & vbCrLf & "参数[2]:" & strParaPathID
            picInfo.Height = lblNote.Top + lblNote.Height + 130 + 400
            rtbSQL.Top = rtbSQL.Top - 610 + 400
            rtbSQL.Height = rtbSQL.Height + 610 - 400
            Line1(3).Y1 = lblNote.Top + lblNote.Height + 110 + 400
            Line1(3).Y2 = lblNote.Top + lblNote.Height + 110 + 400
        Case 26
            lblParaRemark.Caption = "参数[1]:" & strParaPatiAndPage & vbCrLf & "参数[2]:" & strDays
            picInfo.Height = lblNote.Top + lblNote.Height + 130 + 400
            rtbSQL.Top = rtbSQL.Top - 610 + 400
            rtbSQL.Height = rtbSQL.Height + 610 - 400
            Line1(3).Y1 = lblNote.Top + lblNote.Height + 110 + 400
            Line1(3).Y2 = lblNote.Top + lblNote.Height + 110 + 400
        Case 27, 28, 29
            lblParaRemark.Caption = "参数[1]:" & strParaPatiAndPage & vbCrLf & "参数[2]:" & strParaMoney
            picInfo.Height = lblNote.Top + lblNote.Height + 130 + 400
            rtbSQL.Top = rtbSQL.Top - 610 + 400
            rtbSQL.Height = rtbSQL.Height + 610 - 400
            Line1(3).Y1 = lblNote.Top + lblNote.Height + 110 + 400
            Line1(3).Y2 = lblNote.Top + lblNote.Height + 110 + 400
        Case Else
           lblParaRemark.Visible = False
           picInfo.Height = lblNote.Top + lblNote.Height + 130
           rtbSQL.Top = rtbSQL.Top - 610
           rtbSQL.Height = rtbSQL.Height + 610
           Line1(3).Y1 = lblNote.Top + lblNote.Height + 110
           Line1(3).Y2 = lblNote.Top + lblNote.Height + 110
    End Select
End Sub


Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    'On Error Resume Next    '避免最小化时出错
    rtbSQL.Width = Me.ScaleWidth - 120
    rtbSQL.Left = Me.ScaleLeft + 60
    rtbSQL.Height = Me.ScaleHeight - picBottom.Height - picInfo.Height - 120
    
    Line1(0).X2 = Me.ScaleWidth
    Line1(1).X2 = Me.ScaleWidth
    Line1(2).X2 = Me.ScaleWidth
    Line1(3).X2 = Me.ScaleWidth
    
    cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width - 120
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 120
End Sub


Public Function ShowMe(frmMain As Object, ByVal lng行号 As Long, ByRef strSQL文本 As String) As Boolean
'参数：strSQL文本:传回修改后的SQL文本

    mstrSQL文本 = strSQL文本
    mlng行号 = lng行号
    
    Me.Show 1, frmMain
    strSQL文本 = mstrSQL文本
End Function

