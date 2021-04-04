VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ctlDefaultFrame 
   Appearance      =   0  'Flat
   BackColor       =   &H0000C000&
   ClientHeight    =   5145
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10380
   ControlContainer=   -1  'True
   ScaleHeight     =   5145
   ScaleWidth      =   10380
   Begin zl9NewQuery.ctlButton usrMoveUp 
      Height          =   420
      Left            =   405
      TabIndex        =   23
      Top             =   570
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   741
      Caption         =   "向上移动"
      BackColor       =   16761024
      ForeColor       =   0
      AutoSize        =   0   'False
      DrawColor       =   0   'False
   End
   Begin zl9NewQuery.ctlButton usrMoveDown 
      Height          =   420
      Left            =   390
      TabIndex        =   22
      Top             =   3300
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   741
      Caption         =   "向下移动"
      BackColor       =   16761024
      ForeColor       =   0
      AutoSize        =   0   'False
      DrawColor       =   0   'False
   End
   Begin VB.Timer tmrMusic 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   4335
      Top             =   3930
   End
   Begin VB.Timer tmrLoop 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3780
      Top             =   3990
   End
   Begin VB.Timer tmrMsg 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3270
      Top             =   3930
   End
   Begin zl9NewQuery.ctlPicture UsrPic 
      Height          =   885
      Index           =   2
      Left            =   300
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3915
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   1561
      Border          =   0
   End
   Begin zl9NewQuery.ctlPicture UsrPic 
      Height          =   705
      Index           =   1
      Left            =   3060
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   60
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   1244
      Border          =   0
   End
   Begin VB.Timer tmrPlay 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4950
      Top             =   4050
   End
   Begin zl9NewQuery.UsrHome pagHome 
      Height          =   1260
      Left            =   7395
      TabIndex        =   0
      Top             =   705
      Visible         =   0   'False
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   2223
   End
   Begin zl9NewQuery.UsrTodayQuery pagTodayQuery 
      Height          =   1065
      Left            =   1140
      TabIndex        =   6
      Top             =   6255
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1879
   End
   Begin zl9NewQuery.UsrChargeQuery pagChargeQuery 
      Height          =   1320
      Left            =   7080
      TabIndex        =   5
      Top             =   6405
      Visible         =   0   'False
      Width           =   2670
      _ExtentX        =   4710
      _ExtentY        =   2328
   End
   Begin zl9NewQuery.UsrPriceQuery pagPriceQuery 
      Height          =   2910
      Left            =   3735
      TabIndex        =   4
      Top             =   5910
      Visible         =   0   'False
      Width           =   3090
      _ExtentX        =   5450
      _ExtentY        =   5133
   End
   Begin zl9NewQuery.ctlQueryItem usrQueryItem 
      Height          =   1365
      Left            =   3210
      TabIndex        =   3
      Top             =   765
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   2408
   End
   Begin VB.PictureBox picState 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1350
      Left            =   1560
      ScaleHeight     =   1350
      ScaleWidth      =   9660
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3600
      Width           =   9660
      Begin zl9NewQuery.ctlButton UsrEdit 
         Height          =   570
         Left            =   45
         TabIndex        =   24
         Top             =   90
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   1005
         Caption         =   "查询维护"
         FontSize        =   10.5
         AutoSize        =   0   'False
         ButtonHeight    =   450
      End
      Begin VB.PictureBox PicOEM 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   570
         Left            =   7440
         ScaleHeight     =   570
         ScaleWidth      =   1485
         TabIndex        =   20
         Top             =   720
         Width           =   1485
         Begin VB.Image imgFlag 
            Height          =   300
            Left            =   60
            Picture         =   "ctlDefaultFrame.ctx":0000
            Stretch         =   -1  'True
            Top             =   135
            Width           =   390
         End
         Begin VB.Label lblOEM 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "中联软件"
            BeginProperty Font 
               Name            =   "黑体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   480
            TabIndex        =   21
            Top             =   240
            Width           =   900
         End
      End
      Begin VB.PictureBox picMsg 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   570
         Left            =   975
         ScaleHeight     =   570
         ScaleWidth      =   4755
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   735
         Width           =   4755
         Begin VB.Label lblMsg 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "当前没有公告信息！"
            BeginProperty Font 
               Name            =   "黑体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   -90
            TabIndex        =   14
            Top             =   225
            Width           =   2160
         End
      End
      Begin zl9NewQuery.ctlButton UsrBack 
         Height          =   570
         Left            =   6915
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   45
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   1005
         Caption         =   "后退"
         BackColor       =   16777215
         FontName        =   "黑体"
         FontSize        =   10.5
         FontBold        =   -1  'True
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrHome 
         Height          =   570
         Left            =   7935
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   45
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   1005
         Caption         =   "主页"
         BackColor       =   16777215
         FontName        =   "黑体"
         FontSize        =   10.5
         FontBold        =   -1  'True
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   4
         Left            =   5925
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   45
         Visible         =   0   'False
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   1005
         Caption         =   "目录"
         BackColor       =   16777215
         FontName        =   "黑体"
         FontSize        =   10.5
         FontBold        =   -1  'True
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   0
         Left            =   1680
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   30
         Visible         =   0   'False
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   1005
         Caption         =   "左翻"
         BackColor       =   16777215
         FontName        =   "黑体"
         FontSize        =   10.5
         FontBold        =   -1  'True
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   1
         Left            =   2685
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   45
         Visible         =   0   'False
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   1005
         Caption         =   "右翻"
         BackColor       =   16777215
         FontName        =   "黑体"
         FontSize        =   10.5
         FontBold        =   -1  'True
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   2
         Left            =   3615
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   45
         Visible         =   0   'False
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   1005
         Caption         =   "上翻"
         BackColor       =   16777215
         FontName        =   "黑体"
         FontSize        =   10.5
         FontBold        =   -1  'True
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   3
         Left            =   4680
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   45
         Visible         =   0   'False
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   1005
         Caption         =   "下翻"
         BackColor       =   16777215
         FontName        =   "黑体"
         FontSize        =   10.5
         FontBold        =   -1  'True
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrGuoHao 
         Height          =   570
         Left            =   5730
         TabIndex        =   25
         Top             =   675
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   1005
         Caption         =   "自助挂号"
         FontSize        =   10.5
         AutoSize        =   0   'False
         ButtonHeight    =   450
      End
      Begin zl9NewQuery.ctlButton ctlshowwww 
         Height          =   570
         Left            =   45
         TabIndex        =   26
         Top             =   720
         Visible         =   0   'False
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   1005
         Caption         =   "医院信息"
         FontSize        =   10.5
         AutoSize        =   0   'False
         ButtonHeight    =   450
      End
      Begin zl9NewQuery.ctlButton Usr自助打印 
         Height          =   570
         Left            =   6630
         TabIndex        =   27
         Top             =   690
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   1005
         Caption         =   "自助打印"
         FontSize        =   10.5
         AutoSize        =   0   'False
         ButtonHeight    =   450
      End
      Begin zl9NewQuery.ctlButton usr免费号 
         Height          =   570
         Left            =   0
         TabIndex        =   28
         Top             =   0
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   1005
         Caption         =   "简易挂号"
         FontSize        =   10.5
         AutoSize        =   0   'False
         ButtonHeight    =   450
      End
   End
   Begin VB.PictureBox picMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      FillColor       =   &H80000002&
      FillStyle       =   3  'Vertical Line
      ForeColor       =   &H80000008&
      Height          =   2145
      Left            =   45
      ScaleHeight     =   2145
      ScaleWidth      =   1770
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1035
      Width           =   1770
      Begin zl9NewQuery.ctlButton lblMenu 
         Height          =   480
         Index           =   0
         Left            =   30
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   75
         Visible         =   0   'False
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   847
         Caption         =   "按钮按钮按钮按钮"
         BackColor       =   16777215
         FontSize        =   10.5
         AutoSize        =   0   'False
         ButtonHeight    =   360
         TextAligment    =   0
      End
   End
   Begin MSComctlLib.ImageList ilsImage 
      Left            =   5250
      Top             =   4470
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlDefaultFrame.ctx":0ECA
            Key             =   "folder"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlDefaultFrame.ctx":1264
            Key             =   "page"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlDefaultFrame.ctx":15FE
            Key             =   "home"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlDefaultFrame.ctx":1998
            Key             =   "back"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlDefaultFrame.ctx":1D32
            Key             =   "up"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlDefaultFrame.ctx":20CC
            Key             =   "down"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlDefaultFrame.ctx":2466
            Key             =   "menu"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlDefaultFrame.ctx":2800
            Key             =   "list"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlDefaultFrame.ctx":2B9A
            Key             =   "pagedefault"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlDefaultFrame.ctx":48A4
            Key             =   "folderdefault"
         EndProperty
      EndProperty
   End
   Begin zl9NewQuery.ctlPicture UsrPic 
      Height          =   540
      Index           =   0
      Left            =   1200
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   450
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   953
      Border          =   0
   End
   Begin MSComctlLib.ImageList ilsTmp 
      Left            =   2895
      Top             =   2655
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlDefaultFrame.ctx":75F6
            Key             =   "folder"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlDefaultFrame.ctx":7990
            Key             =   "page"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlDefaultFrame.ctx":7D2A
            Key             =   "home"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlDefaultFrame.ctx":80C4
            Key             =   "back"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlDefaultFrame.ctx":845E
            Key             =   "up"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlDefaultFrame.ctx":87F8
            Key             =   "down"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlDefaultFrame.ctx":8B92
            Key             =   "menu"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlDefaultFrame.ctx":8F2C
            Key             =   "list"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEditTable 
         Caption         =   "用户表格定义(&1)"
      End
      Begin VB.Menu mnuEditPicture 
         Caption         =   "查询图形设置(&2)"
      End
      Begin VB.Menu mnuEditProfession 
         Caption         =   "专家介绍清单(&3)"
      End
      Begin VB.Menu mnuEditAdvice 
         Caption         =   "播放广告设置(&4)"
      End
      Begin VB.Menu mnuEdit_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditPage 
         Caption         =   "查询页面定义(&5)"
      End
      Begin VB.Menu mnuEditFolder 
         Caption         =   "查询目录规划(&6)"
      End
      Begin VB.Menu mnuEdit_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditParam 
         Caption         =   "查询参数设置(&7)"
      End
      Begin VB.Menu mnuEdit_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditPriter 
         Caption         =   "自助打印报告(&8)"
      End
   End
End
Attribute VB_Name = "ctlDefaultFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mvarAdvicePlayInternal As Long
Private mvarAdvicePlayLong As Long
'zyk add 200410
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Event KeyPress(KeyAscii As Integer)
Public Event ShowPage(ByVal PageNo As Long, ByVal CusomFormat As String)
Public Event MenuClick(ByVal Key As String, ByVal ParentKey As String)
Public Event BackClick()
Public Event HomeClick()
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event ConnectClick(ByVal PageNo As Long, ByVal OrderNo As Long)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)

Private mvar页面栈(1 To 20) As String
Private mvar栈顶 As Long

Private mvarKey As String
Private mvarListKey As Long

Private blnBack As Boolean
Private mstrMusicFile As String

Private mvarPlayOrder As Long               '当前广告播放的序号

Event ExitNewQuery(blnCancel As Boolean)      '退咨询系统事件

Public Sub RefreshFolder()
    '功能:刷新目录
    
    Call LoadMenuList(mvarListKey)
    Call RefreshPage
    
End Sub

Public Function RefreshPage() As Boolean
    '功能:刷新当前查询页面内容
    
    Dim intLoop As Integer
    
    RefreshPage = True
    
    Call CheckPicture
    
    'mvarKey为当前页面关键字
    If mvarKey <> "" And mvarKey <> "0;0" Then
        For intLoop = 1 To lblMenu.UBound
            If lblMenu(intLoop).Key = Split(mvarKey, ";")(0) Then
                mvarKey = ""
                Call lblMenu_CommandClick(intLoop)
                Exit Function
            End If
        Next
        
        If intLoop > lblMenu.UBound Then
            '主页
            RaiseEvent HomeClick
            Call ShowHome
        End If
    ElseIf mvarKey = "0;0" Then
        '主页
        RaiseEvent HomeClick
        Call ShowHome
    End If
    
    
End Function

Private Sub CalcMoveState()
    '---------------------------------------------------------------------------------------------
    '计算是否可以向上移动或者向下移动
    '---------------------------------------------------------------------------------------------
    
    If lblMenu.Count > 1 Then
        usrMoveDown.Enabled = (lblMenu(lblMenu.Count - 1).Top + lblMenu(lblMenu.Count - 1).Height > picMenu.Height)
        usrMoveUp.Enabled = (lblMenu(1).Top < 0)
    End If
    
End Sub

'zyk add 200410
Private Sub ctlshowwww_CommandClick()
    Dim wwwurl As String
    wwwurl = GetPara("医院主页")
    If wwwurl = "" Then
        MsgBox "医院网址没有设置,请在参数设置允许""使用医院网站"",并输入医院网址", vbOKOnly
    Else
        ShellExecute hwnd, "open", "iexplore.exe", "-k " & wwwurl, "", 1
        'Sleep 5000   'API函数延时5000毫秒
    End If
End Sub


Private Sub lblMenu_CommandClick(Index As Integer)
    Dim lngNO As Long
    
    lngNO = Val(Mid(lblMenu(Index).Key, 2))
    
    If mvarKey <> (lblMenu(Index).Key & ";" & lblMenu(Index).Tag) Then
        Call SelectMenu(Index)
        Call LocationPage(lblMenu(Index).Key, lblMenu(Index).Tag)
    End If
    
    '执行命令行
    Dim rs As New ADODB.Recordset
    
    gstrSQL = "select A.页面,B.命令参数 from 咨询页面排列 A,咨询页面目录 B where A.页面=B.页面序号(+) and A.序号=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "页面框架", lngNO)
    If rs.BOF = False Then
        If IIf(IsNull(rs!页面), 0, rs!页面) > 0 Then
            If IsNull(rs("命令参数").Value) = False Then
            
                If Trim(rs("命令参数").Value) <> "" Then
                    On Error Resume Next
                    Call Shell(rs("命令参数").Value, vbNormalFocus)
                End If
                
            End If
        End If
    End If
    
End Sub

Private Sub lblMenu_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub lblMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub lblMsg_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub mnuEditAdvice_Click()
    RunMudal 4
End Sub

Private Sub mnuEditFolder_Click()
    RunMudal 6
End Sub

Private Sub mnuEditPage_Click()
    RunMudal 5
End Sub

Private Sub mnuEditParam_Click()
    Call frmParameter.ShowDialog(gfrmMain, gstrPrivs)
End Sub

Private Sub mnuEditPicture_Click()
    RunMudal 2
End Sub

Private Sub mnuEditPriter_Click()
    RunMudal 9
End Sub

Private Sub mnuEditProfession_Click()
    RunMudal 3
End Sub

Private Sub mnuEditTable_Click()
    RunMudal 1
End Sub

Private Sub pagChargeQuery_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub pagChargeQuery_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub pagHome_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub pagHome_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub pagPriceQuery_ClickOK(ByVal strQuery As String, blnCancel As Boolean)
    '修改编号2667
    '判断输入的查询内容是不是AdminExitNewQuery,如果是就激活退出事件
    Dim blnTmp As Boolean
    
    If UCase(strQuery) = UCase("AdminExitQuery") Then
        RaiseEvent ExitNewQuery(blnCancel)
    End If
End Sub

Private Sub pagPriceQuery_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub pagPriceQuery_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub pagTodayQuery_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub pagTodayQuery_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub picMenu_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub picMenu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub picMenu_Paint()
    DrawColorToColor picMenu, picMenu.BackColor, &HFFC0C0
End Sub

Private Sub picMsg_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub picMsg_Paint()
    DrawColorToColor picMsg, picMsg.BackColor, &HFFC0C0
End Sub

Private Sub picMsg_Resize()
    lblMsg.Top = (picMsg.ScaleHeight - lblMsg.Height) / 2
End Sub

Private Sub PicOEM_Paint()
    DrawColorToColor PicOEM, picMsg.BackColor, &HFFC0C0
End Sub

Private Sub PicOEM_Resize()
    imgFlag.Top = (PicOEM.ScaleHeight - imgFlag.Height) / 2
    lblOEM.Top = (PicOEM.ScaleHeight - lblOEM.Height) / 2
End Sub

Private Sub picState_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub picState_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub picState_Paint()
    DrawColorToColor picState, picState.BackColor, &HFFC0C0
End Sub

Private Sub tmrLoop_Timer()
    If lblMsg.Left - 300 + lblMsg.Width < 600 Then
        tmrMsg.Interval = 1
        lblMsg.Left = picMsg.Width
    Else
        lblMsg.Left = lblMsg.Left - 15
    End If
End Sub

Private Sub tmrMsg_Timer()
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
        
    lblMsg.Caption = GetPara("公告信息")
    lblMsg.Caption = IIf(Trim(lblMsg.Caption) = "", "", lblMsg.Caption)
    If lblMsg.Caption = "" Then
        tmrLoop.Interval = 0
        tmrLoop.Enabled = False
        lblMsg.Left = 600
    Else
        tmrLoop.Interval = 50
        tmrLoop.Enabled = True
    End If
    
    tmrMsg.Interval = 30000
    
    Exit Sub
errHand:
End Sub

Private Sub tmrMusic_Timer()
    '
    If MusicPlayStatus = False And mstrMusicFile <> "" Then
        Call MusicClose
        Call MusicPlay(mstrMusicFile)
    End If
End Sub

Private Sub tmrPlay_Timer()
    Dim W As Single
    Dim H As Single
    Dim vSQL As String
    Dim vFileName As String
    Dim vRs As New ADODB.Recordset
    Dim vMovieHead As FLASHHEADER
                        
    On Error GoTo errHand
    
    mvarAdvicePlayInternal = mvarAdvicePlayInternal + 1000
    If mvarAdvicePlayInternal < mvarAdvicePlayLong Then Exit Sub
    mvarAdvicePlayInternal = 0
    
    vSQL = "Select A.序号,A.图片序号 From  咨询广告序列 A where A.序号>[1]  order by A.序号"
    Set vRs = zlDatabase.OpenSQLRecord(vSQL, "页面框架", mvarPlayOrder)
    If vRs.BOF = False Then
        '下一幅广告存在,则直接取之
        vFileName = GetFileName(IIf(IsNull(vRs!图片序号), 0, vRs!图片序号), W, H)
        mvarPlayOrder = IIf(IsNull(vRs!序号), 0, vRs!序号)
    Else
        '下一幅广告不存在,则返回到第一条
        vSQL = "Select B.类型,B.名称,A.序号,A.图片序号 From  咨询广告序列 A,咨询图片元素 B where A.图片序号=B.序号 order by A.序号"
        Set vRs = zlDatabase.OpenSQLRecord(vSQL, "页面框架")
        If vRs.BOF = False Then
            vFileName = GetFileName(IIf(IsNull(vRs!图片序号), 0, vRs!图片序号), W, H)
            mvarPlayOrder = IIf(IsNull(vRs!序号), 0, vRs!序号)
        End If
    End If
    
    '设置缺省的广告播放时间
    mvarAdvicePlayLong = GetInterval * 1000
    
    If vFileName <> "" And Dir(vFileName) <> "" Then
        If Right(vFileName, 3) = "swf" Then
            '如果是Flash电影,则计算实际的播放时间
            vMovieHead = GetFlashHeader(vFileName)
            If vMovieHead.intIsFlashMovie = 1 And vMovieHead.intMRate > 0 Then
                mvarAdvicePlayLong = IIf((1000 * (vMovieHead.intMTotalFrames / vMovieHead.intMRate)) < mvarAdvicePlayLong, mvarAdvicePlayLong, 1000 * (vMovieHead.intMTotalFrames / vMovieHead.intMRate))
            End If
        End If
        '刷新广告播放
        Call UsrPic(2).ShowPictureByFile(vFileName)
    End If
    CloseRecord vRs
    Exit Sub
errHand:
    CloseRecord vRs
'    If ErrCenter() = -1 Then Resume
    Call SaveErrLog
End Sub

Private Sub UserControl_Initialize()
    
    '最小宽度1800,最大宽度2400
    
    'UsrPic(0).Width = 1800
    UsrPic(0).Width = 2400
    
    UsrPic(0).Height = 945
    
    picMenu.Width = UsrPic(0).Width
    UsrPic(2).Height = 1200
    
    lblMenu(0).Left = 30
    lblMenu(0).Width = UsrPic(0).Width - 60
    
    lblMenu(0).Font.Name = "黑体"
    lblMenu(0).Font.Size = 12
    
    usrMoveUp.Font.Name = "黑体"
    usrMoveDown.Font.Name = "黑体"
    
    usrMoveUp.Font.Size = 12
    usrMoveDown.Font.Size = 12
    
    usrMoveUp.Font.Bold = True
    usrMoveDown.Font.Bold = True
    
    usrMoveUp.ForeColor = &HFF&
    usrMoveDown.ForeColor = &HFF&
    
    usrMoveUp.Picture = ilsTmp.ListImages("up")
    usrMoveDown.Picture = ilsTmp.ListImages("down")
    
    usrMoveUp.TextAligment = 1
    usrMoveDown.TextAligment = 1
        
    '江磊改于2003-8-12 编号：2491    目的：将右下面的按钮上移
    picMsg.Height = lblMsg.Height + 100 'UsrHome.Top + UsrHome.Height + 45
    picState.Height = picState.Height + picMsg.Height
    
    UsrBack.Picture = ilsImage.ListImages("back")
    UsrHome.Picture = ilsImage.ListImages("home")
    
    UsrEdit.ShowPicture = False
    UsrGuoHao.ShowPicture = False
    Usr自助打印.ShowPicture = False
    usr免费号.ShowPicture = False
    ctlshowwww.ShowPicture = False
    ctlshowwww.Width = UsrEdit.Width
    
    mvarKey = ""
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub



Private Sub UserControl_Resize()
    On Error Resume Next
        
    Call ResizeControl(UsrPic(0), 0, 0, UsrPic(0).Width, UsrPic(0).Height)
        
    Call ResizeControl(UsrPic(1), UsrPic(0).Left + UsrPic(0).Width + 45, 0, UserControl.Width - UsrPic(0).Left - UsrPic(0).Width - 45, UsrPic(0).Height)
    
    Call ResizeControl(usrMoveUp, 0, UsrPic(0).Top + UsrPic(0).Height + 45, picMenu.Width, usrMoveUp.Height)
    
    Call ResizeControl(picMenu, 0, usrMoveUp.Top + usrMoveUp.Height, picMenu.Width, UserControl.Height - usrMoveUp.Top - usrMoveUp.Height - usrMoveDown.Height - UsrPic(2).Height - 45)
    
    Call ResizeControl(usrMoveDown, 0, picMenu.Top + picMenu.Height, picMenu.Width, usrMoveDown.Height)
            
    Call ResizeControl(UsrPic(2), 0, usrMoveDown.Top + usrMoveDown.Height + 45, picMenu.Width, UsrPic(2).Height)
        
    Call ResizeControl(usrQueryItem, picMenu.Left + picMenu.Width + 45, usrMoveUp.Top, UserControl.Width - picMenu.Left - picMenu.Width - 45, picMenu.Height + usrMoveUp.Height + usrMoveDown.Height)
    
    Call ResizeControl(picState, usrQueryItem.Left, UsrPic(2).Top, usrQueryItem.Width, UserControl.Height - UsrPic(2).Top)
        
    Call ResizeControl(UsrHome, picState.ScaleWidth - UsrHome.Width - 450, 45, UsrHome.Width, UsrHome.Height)
    Call ResizeControl(UsrBack, UsrHome.Left - UsrBack.Width - 30, UsrHome.Top, UsrBack.Width, UsrBack.Height)
    Call ResizeControl(UsrCmd(4), UsrBack.Left - UsrCmd(4).Width - 30, UsrBack.Top, UsrCmd(4).Width, UsrCmd(4).Height)
    Call ResizeControl(UsrCmd(3), UsrCmd(4).Left - UsrCmd(3).Width - 30, UsrCmd(4).Top, UsrCmd(3).Width, UsrCmd(3).Height)
    Call ResizeControl(UsrCmd(2), UsrCmd(3).Left - UsrCmd(2).Width - 30, UsrCmd(3).Top, UsrCmd(2).Width, UsrCmd(2).Height)
    Call ResizeControl(UsrCmd(1), UsrCmd(2).Left - UsrCmd(1).Width - 30, UsrCmd(2).Top, UsrCmd(1).Width, UsrCmd(1).Height)
    Call ResizeControl(UsrCmd(0), UsrCmd(1).Left - UsrCmd(0).Width - 30, UsrCmd(1).Top, UsrCmd(0).Width, UsrCmd(0).Height)
    
    Call ResizeControl(UsrEdit, 30, UsrCmd(1).Top, UsrEdit.Width, UsrEdit.Height)
    Call ResizeControl(UsrGuoHao, IIf(UsrEdit.Visible, UsrEdit.Left + UsrEdit.Width + 30, 30), UsrCmd(1).Top, UsrEdit.Width, UsrEdit.Height)
    Call ResizeControl(Usr自助打印, IIf(UsrGuoHao.Visible, UsrGuoHao.Left + UsrGuoHao.Width + 30, 30), UsrCmd(1).Top, UsrEdit.Width, UsrEdit.Height)
    
     'zyk
    If UsrEdit.Visible Then
       Call ResizeControl(ctlshowwww, IIf(UsrGuoHao.Visible, UsrGuoHao.Left + UsrGuoHao.Width + 30, UsrEdit.Left + UsrEdit.Width + 30), UsrCmd(1).Top, UsrGuoHao.Width, UsrGuoHao.Height)
       Call ResizeControl(ctlshowwww, IIf(Usr自助打印.Visible, Usr自助打印.Left + Usr自助打印.Width + 30, UsrGuoHao.Left + UsrGuoHao.Width + 30), UsrCmd(1).Top, Usr自助打印.Width, Usr自助打印.Height)
       
    Else
        Call ResizeControl(ctlshowwww, 30, UsrCmd(1).Top, UsrEdit.Width, UsrEdit.Height)
    End If
  
    '以下是固定查询页面
    Call ResizeControl(pagPriceQuery, usrQueryItem.Left, usrQueryItem.Top, usrQueryItem.Width, usrQueryItem.Height)
    Call ResizeControl(pagChargeQuery, usrQueryItem.Left, usrQueryItem.Top, usrQueryItem.Width, usrQueryItem.Height)
    Call ResizeControl(pagTodayQuery, usrQueryItem.Left, usrQueryItem.Top, usrQueryItem.Width, usrQueryItem.Height)
    Call ResizeControl(pagHome, usrQueryItem.Left, usrQueryItem.Top, usrQueryItem.Width, usrQueryItem.Height)
    
    '江磊改于2003-8-12 编号：2491    目的：将右下面的按钮上移
    Call ResizeControl(PicOEM, 0, UsrHome.Top + UsrHome.Height + 100, PicOEM.ScaleWidth, picState.ScaleHeight - (UsrHome.Top + UsrHome.Height + 100))
    Call ResizeControl(picMsg, PicOEM.Left + PicOEM.Width, PicOEM.Top, picState.ScaleWidth - PicOEM.Width, PicOEM.Height)
    Call ResizeControl(lblMsg, picMsg.ScaleWidth, (picMsg.ScaleHeight - lblMsg.Height) / 2, lblMsg.Width, lblMsg.Height)
    
End Sub

Public Sub AddMenuItem(ByVal Index As Long, Title As String, Key2 As String, Key As String, Optional ByVal ParentKey As String = "", Optional ByVal IconFile As String = "", Optional ByVal FontName As String = "黑体", Optional ByVal FontSize As Single = 12, Optional ByVal FontForm As Byte = 1, Optional ByVal FontColor As Long = &HFF0000)
    '
    Load lblMenu(Index)
            
    lblMenu(Index).ZOrder
    lblMenu(Index).Font.Name = FontName
    lblMenu(Index).Font.Size = FontSize
    
    lblMenu(Index).Font.Bold = False
    lblMenu(Index).Font.Italic = False
        
    Select Case FontForm
    Case 2
        lblMenu(Index).Font.Italic = True
    Case 3
        lblMenu(Index).Font.Bold = True
    Case 4
        lblMenu(Index).Font.Bold = True
        lblMenu(Index).Font.Italic = True
    End Select
    
    lblMenu(Index).ForeColor = FontColor
    
    
    
    lblMenu(Index).Caption = Title
    lblMenu(Index).Key2 = Key2
    lblMenu(Index).Key = Key
    lblMenu(Index).Tag = ParentKey
    lblMenu(Index).Left = 30
    lblMenu(Index).Top = 60 * Index + lblMenu(Index).Height * (Index - 1)
    lblMenu(Index).Width = lblMenu(0).Width
        
    If IconFile = "" Then
        If Val(Key2) > 0 Then
            lblMenu(Index).Picture = ilsImage.ListImages("pagedefault")
        Else
            lblMenu(Index).Picture = ilsImage.ListImages("folderdefault")
        End If
    Else
        ilsTmp.ListImages.Clear
        ilsTmp.ImageWidth = 16
        ilsTmp.ImageHeight = 16
        ilsTmp.ListImages.Add 1, , VB.LoadPicture(IconFile)
        
        lblMenu(Index).Picture = ilsTmp.ListImages(1)
    End If
    lblMenu(Index).Visible = True
        
End Sub

Public Sub ClearAllMenuItem()
    '清除所有的功能菜单项
    Dim i As Long
    
    For i = lblMenu.UBound To 1 Step -1
        Unload lblMenu(i)
    Next
End Sub

Public Property Get ClientWidth() As Single
    ClientWidth = usrQueryItem.Width
End Property
    
Public Property Get ClientObj() As ctlQueryItem
    Set ClientObj = usrQueryItem
End Property

Private Sub UserControl_Show()
    Call tmrPlay_Timer
End Sub

Private Sub UsrBack_CommandClick()
    Dim i As Integer
    
    On Error GoTo errHand
    
    blnBack = True
    mvarKey = OutFront
    If mvarKey <> "" Then
                
        '1.检查上一页面是否在当前一层中
        For i = 1 To lblMenu.UBound
            If lblMenu(i).Key = Split(mvarKey, ";")(0) Then
                mvarKey = ""
                Call lblMenu_CommandClick(i)
                blnBack = False
                Exit Sub
            End If
        Next
        
        '2.上一页不在当前一层中，直接找出他所在的一层，并显示所在层的功能菜单体系
        gstrSQL = "select nvl(父序号,0) as 父序号 from 咨询页面排列 where 序号=[1]"
        Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, "页面框架", Val(Mid(Split(mvarKey, ";")(0), 2)))
        If gRs.BOF = False Then
            Call LoadMenuList(gRs!父序号)
            For i = 1 To lblMenu.UBound
                If lblMenu(i).Key = Split(mvarKey, ";")(0) Then
                    mvarKey = ""
                    Call lblMenu_CommandClick(i)
                    blnBack = False
                    Exit Sub
                End If
            Next
        End If
        
        '3.上一页可能是主页面
        Call ShowHome
        
    End If
    blnBack = False
    Exit Sub
errHand:
    blnBack = False
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub UsrBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UsrCmd_CommandClick(Index As Integer)
    Select Case Index
    Case 0
        Call usrQueryItem.TurnToLeftPage
    Case 1
        Call usrQueryItem.TurnToRightPage
    Case 2
        Call usrQueryItem.TurnToLastPage
    Case 3
        Call usrQueryItem.TurnToNextPage
    Case 4
        Call usrQueryItem.ShowTreeList
    End Select
End Sub

Private Sub UsrCmd_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UsrEdit_CommandClick()

    UserControl.PopupMenu UserControl.mnuEdit, , picState.Left + UsrEdit.Left, picState.Top + UsrEdit.Top - 30
    
    EnterFocus picMenu
    
End Sub

Private Sub UsrGuoHao_CommandClick()
    Call InitLocPar
    Call InitSysPar
    
    On Error Resume Next
    
    frmselectinfo.Show , gfrmMain
    
    EnterFocus picMenu
    
End Sub

Private Sub UsrHome_CommandClick()
    
    RaiseEvent HomeClick
    If mvarKey <> "0;0" Then Call ShowHome
    
End Sub

Private Sub UsrHome_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub usrMoveDown_CommandClick()
    '向下移动一行
    
    Dim lngLoop As Long
    
    If lblMenu.Count < 1 Then Exit Sub
    
    For lngLoop = 1 To lblMenu.Count - 1
        lblMenu(lngLoop).Top = lblMenu(lngLoop).Top - lblMenu(0).Height - 60
    Next
    
    Call CalcMoveState
    
End Sub

Private Sub usrMoveDown_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub usrMoveUp_CommandClick()
    Dim lngLoop As Long
    
    If lblMenu.Count < 1 Then Exit Sub
    
    For lngLoop = 1 To lblMenu.Count - 1
        lblMenu(lngLoop).Top = lblMenu(lngLoop).Top + lblMenu(0).Height + 60
    Next
    
    Call CalcMoveState
End Sub

Private Sub usrMoveUp_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UsrPic_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UsrPic_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub usrQueryItem_ChangeNavigator()
    '功能:根据页面的范围控制按钮是否有效

    UsrCmd(0).Enabled = True
    UsrCmd(1).Enabled = True
    UsrCmd(2).Enabled = True
    UsrCmd(3).Enabled = True
        
    If usrQueryItem.ValueVsb = 0 Then UsrCmd(2).Enabled = False
    If usrQueryItem.ValueVsb = usrQueryItem.MaxVsb Then UsrCmd(3).Enabled = False
    
    If usrQueryItem.ValueHsb = 0 Then UsrCmd(0).Enabled = False
    If usrQueryItem.ValueHsb = usrQueryItem.MaxHsb Then UsrCmd(1).Enabled = False
    
End Sub

Private Sub usrQueryItem_ConnectClick(ByVal PageNo As Long, ByVal OrderNo As Long)
    Dim i As Integer
    
    If PageNo > 0 Then
                
        '1.检查上一页面是否在当前一层中
        For i = 1 To lblMenu.UBound
            If Val(lblMenu(i).Key2) = PageNo Then
                blnBack = False
                Call lblMenu_CommandClick(i)
                GoTo EndHand
            End If
        Next
        
        '2.上一页不在当前一层中，直接找出他所在的一层，并显示所在层的功能菜单体系
        
        gstrSQL = "select 父序号 from 咨询页面排列 where 页面=" & PageNo
        
        gstrSQL = "select nvl(父序号,0) as 父序号 from 咨询页面排列 where 页面=[1]"
        Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, "页面框架", PageNo)
        If gRs.BOF = False Then
            Call LoadMenuList(gRs!父序号)
            For i = 1 To lblMenu.UBound
                If Val(lblMenu(i).Key2) = PageNo Then
                    blnBack = False
                    Call lblMenu_CommandClick(i)
                    GoTo EndHand
                End If
            Next
        End If
    End If
    Exit Sub
EndHand:
    Call usrQueryItem.GoPageItemByOrder(OrderNo)
End Sub

Public Property Get MusicFile() As String
    MusicFile = mstrMusicFile
End Property

Public Property Let MusicFile(ByVal vData As String)
    mstrMusicFile = vData
End Property

Public Property Let PageType(ByVal vData As Byte)
    Select Case vData
    Case 99         '主页面
        pagHome.Visible = True
        pagTodayQuery.Visible = False
        pagPriceQuery.Visible = False
        pagChargeQuery.Visible = False
        usrQueryItem.Visible = False
                
        usrQueryItem.Enabled = False
        pagTodayQuery.Enabled = False
        pagChargeQuery.Enabled = False
        pagPriceQuery.Enabled = False
        pagHome.Enabled = True
    Case 0          '自定义页面
        pagHome.Visible = False
        pagTodayQuery.Visible = False
        pagPriceQuery.Visible = False
        pagChargeQuery.Visible = False
        usrQueryItem.Visible = True
        
        usrQueryItem.Enabled = True
        pagTodayQuery.Enabled = False
        pagChargeQuery.Enabled = False
        pagPriceQuery.Enabled = False
        pagHome.Enabled = False
    Case 1          '收费价格查询页面
        pagHome.Visible = False
        pagTodayQuery.Visible = False
        pagPriceQuery.Visible = True
        pagChargeQuery.Visible = False
        usrQueryItem.Visible = False
        
        usrQueryItem.Enabled = False
        pagTodayQuery.Enabled = False
        pagChargeQuery.Enabled = False
        pagPriceQuery.Enabled = True
        pagHome.Enabled = False
        
        Call pagPriceQuery.InitLoad
        
    Case 2          '病人费用查询页面
        pagHome.Visible = False
        pagTodayQuery.Visible = False
        pagPriceQuery.Visible = False
        pagChargeQuery.Visible = True
        usrQueryItem.Visible = False
        
        usrQueryItem.Enabled = False
        pagTodayQuery.Enabled = False
        pagChargeQuery.Enabled = True
        pagPriceQuery.Enabled = False
        pagHome.Enabled = False
        
        Call pagChargeQuery.InitLoad
        
    Case 3          '专家介绍页面
        pagHome.Visible = False
        pagTodayQuery.Visible = False
        pagPriceQuery.Visible = False
        pagChargeQuery.Visible = False
        usrQueryItem.Visible = True
        
        usrQueryItem.Enabled = True
        pagTodayQuery.Enabled = False
        pagChargeQuery.Enabled = False
        pagPriceQuery.Enabled = False
        pagHome.Enabled = False
        
    Case 4          '今日就诊页面
        pagHome.Visible = False
        pagTodayQuery.Visible = True
        pagPriceQuery.Visible = False
        pagChargeQuery.Visible = False
        usrQueryItem.Visible = False
        
        usrQueryItem.Enabled = False
        pagTodayQuery.Enabled = True
        pagChargeQuery.Enabled = False
        pagPriceQuery.Enabled = False
        pagHome.Enabled = False
        
        Call pagTodayQuery.InitLoad
    End Select
End Property

Public Sub SelectMenu(ByVal Index As Integer)
    Dim i As Long
    
'    For i = 1 To lblMenu.UBound
'        lblMenu(i).State = 0
'    Next
'    lblMenu(Index).State = -1
    
End Sub


Private Sub InFront(ByVal Key As String)
'功能:将选择的查询页面进行入栈操作
'参数:页面关键字
    Dim i As Long
    
    If mvar栈顶 >= 20 Then
        For i = 2 To 20
            mvar页面栈(i - 1) = mvar页面栈(i)
        Next
        mvar栈顶 = mvar栈顶 - 1
    End If
    mvar栈顶 = mvar栈顶 + 1
    mvar页面栈(mvar栈顶) = Key
    
End Sub

Private Function OutFront() As String
'功能:已查询页面的出栈操作
'返回:页面关键字

    If mvar栈顶 <= 1 Then Exit Function
    mvar栈顶 = mvar栈顶 - 1
    If mvar栈顶 > 0 Then OutFront = mvar页面栈(mvar栈顶)

End Function

Public Sub ShowHome()
'功能:显示主页面
    Dim W As Single
    Dim H As Single
        
    Call LoadMenuList(0)
    PageType = 99
    
    Call pagHome.InitLoad
    Call InitNavigator(0, 0)
    DoEvents
    
    On Error GoTo errHand
    
    gstrSQL = "select A.宣传标语,A.页面背景 from 咨询页面目录 A where A.页面序号=0"
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, "页面框架")
    If gRs.BOF = False Then
        UsrPic(1).Tag = GetFileName(IIf(IsNull(gRs!宣传标语), 0, gRs!宣传标语), W, H)
        Call UsrPic(1).ShowPictureByFile(UsrPic(1).Tag)
    End If
    
    gstrSQL = "select A.插图序号 from 咨询段落目录 A where A.页面序号=0 and A.段落序号=2"
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, "页面框架")
    If gRs.BOF = False Then
        UsrPic(0).Tag = GetFileName(IIf(IsNull(gRs!插图序号), 0, gRs!插图序号), W, H)
        Call UsrPic(0).ShowPictureByFile(UsrPic(0).Tag)
    End If
    
    '获取背景音乐文件
    MusicFile = ""
    gstrSQL = "select B.名称 from 咨询页面目录 A,咨询图片元素 B where A.背景音乐=B.序号 and A.页面序号=0"
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, "主页显示")
    If gRs.BOF = False Then
        If IsNull(gRs!名称) = False Then MusicFile = App.Path & "\图形\" & gRs!名称 & ".mid"
    End If
    tmrMusic.Enabled = False
    Call MusicClose
    Call MusicPlay(mstrMusicFile)
    If mstrMusicFile <> "" Then tmrMusic.Enabled = True
        
    If mvarKey <> "0;0" And blnBack = False Then
        mvarKey = "0;0"
        Call InFront(mvarKey)
    End If
    
    DoEvents
    
    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function LoadMenuList(ByVal lngUpKey As Long) As Long
    '功能:装载某一层的功能菜单列表
    '参数:上一级的功能菜单项关键字
    Dim i As Long
    Dim lngLoop As Long
    Dim W As Single
    Dim H As Single
    Dim vFileName As String
    Dim sglMax As Single
    Dim sglTmp As Single
    
    On Error GoTo errHand
    
    gstrSQL = "select A.* from 咨询页面排列 A where " & IIf(lngUpKey = 0, "(A.父序号=0 or A.父序号 is null)", "A.父序号=[1]")
    gstrSQL = gstrSQL & " and (A.页面>0 or A.序号 in (select B.父序号 from 咨询页面排列 B where B.父序号=A.序号)) order by A.序号"
    
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, "页面框架", lngUpKey)
    
    Call ClearAllMenuItem
    
    i = 1
    If gRs.BOF = False Then
        mvarListKey = lngUpKey
        While Not gRs.EOF
            vFileName = ""
            If IIf(IsNull(gRs!页面图标), 0, gRs!页面图标) > 0 Then vFileName = GetFileName(IIf(IsNull(gRs!页面图标), 0, gRs!页面图标), W, H)
                        
            Call AddMenuItem(i, IIf(IsNull(gRs!名称), "", gRs!名称), IIf(IsNull(gRs!页面), "", gRs!页面), "K" & gRs!序号, CStr(lngUpKey), vFileName, IIf(IsNull(gRs!字体), "黑体", gRs!字体), IIf(IsNull(gRs!大小), 12, gRs!大小), IIf(IsNull(gRs!字形), 1, gRs!字形), IIf(IsNull(gRs!颜色), &HFF0000, gRs!颜色))
            
'            With lblMenu(i)
'
'                picMenu.Font.Name = .Font.Name
'                picMenu.Font.Size = .Font.Size
'                picMenu.Font.Bold = .Font.Bold
'                picMenu.Font.Italic = .Font.Italic
'
'                sglTmp = picMenu.TextWidth(IIf(IsNull(gRs!名称), "", gRs!名称))
'                If sglTmp > sglMax Then sglMax = sglTmp
'
'            End With
            
            gRs.MoveNext
            i = i + 1
        Wend
    End If
    
'    If sglMax < 1800 Then sglMax = 1800
'    If sglMax > 3000 Then sglMax = 3000
    
    'UsrPic(0).Width = sglMax + 360
    
    'Call UserControl_Initialize
    
'    For lngLoop = 1 To lblMenu.UBound
'        lblMenu(lngLoop).Width = lblMenu(0).Width
'    Next
    
    LoadMenuList = i - 1
    
    
    
    '计算是否可以向上移动或者向下移动
    Call CalcMoveState
      
    
    Exit Function
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub LocationPage(ByVal Key As String, ByVal ParentKey As String)
    Dim W As Single
    Dim H As Single
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
       
    gstrSQL = "select A.页面,B.固定页面,B.页面名称,B.宣传标语,B.页面背景,B.命令参数 from 咨询页面排列 A,咨询页面目录 B where A.页面=B.页面序号(+) and A.序号=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "页面框架", Val(Mid(Key, 2)))
    If rs.BOF = False Then
        If IIf(IsNull(rs!页面), 0, rs!页面) > 0 Then
            '选中的是具体的查询页面项
                        
            Call usrQueryItem.ClearAllPageItem
                        
            
            Call InitNavigator(0, 0)
            
            If IIf(IsNull(rs!固定页面), 0, rs!固定页面) = 1 Then
                '系统固定页面的显示
                AdviceMovie = GetFileName(IIf(IsNull(rs!宣传标语), 0, rs!宣传标语), W, H)
                
                Select Case IIf(IsNull(rs!页面名称), "", rs!页面名称)
                Case "收费价格"
                    PageType = 1
                Case "病人费用"
                    PageType = 2
                Case "专家介绍"
                    PageType = 0
                    RaiseEvent ShowPage(IIf(IsNull(rs!页面), 0, rs!页面), "专家介绍")
                Case "今日就诊"
                    PageType = 4
                End Select
            Else
                '用户自定义页面的显示
                PageType = 0
                RaiseEvent ShowPage(IIf(IsNull(rs!页面), 0, rs!页面), "")
                                
            End If
                                               
            '获取背景音乐文件
            MusicFile = ""
            gstrSQL = "select B.名称 from 咨询页面目录 A,咨询图片元素 B where A.背景音乐=B.序号 and A.页面序号=[1]"
            Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, "查询显示", Val(IIf(IsNull(rs!页面), 0, rs!页面)))
            If gRs.BOF = False Then
                If IsNull(gRs!名称) = False Then MusicFile = App.Path & "\图形\" & gRs!名称 & ".mid"
            End If
            tmrMusic.Enabled = False
            Call MusicClose
            Call MusicPlay(mstrMusicFile)
            If mstrMusicFile <> "" Then tmrMusic.Enabled = True
            
            mvarKey = Key & ";" & ParentKey
            
            '如果不是后退操作同时也不是目录项，则进行进栈操作保存页面序号
            If Key <> "" And blnBack = False Then Call InFront(mvarKey)
            
        Else
            '选中的是查询目录项,将显示此查询目录下的页面清单
            If LoadMenuList(Val(Mid(Key, 2))) > 0 Then Call lblMenu_CommandClick(1)
        End If
    End If
    CloseRecord rs
    Exit Sub
errHand:
    CloseRecord rs
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Property Let HospitalMovie(ByVal vData As String)
    UsrPic(0).Tag = vData
End Property

Public Property Get HospitalMovie() As String
    HospitalMovie = UsrPic(0).Tag
End Property

Public Property Let AdviceMovie(ByVal vData As String)
    UsrPic(1).Tag = vData
    UsrPic(1).ShowPictureByFile (vData)
End Property

Public Property Get AdviceMovie() As String
    AdviceMovie = UsrPic(1).Tag
End Property

Public Property Let AllowEdit(vData As Boolean)
    UsrEdit.Visible = vData
End Property

Public Property Let AllowFreeRegist(vData As Boolean)
    usr免费号.Visible = vData
     If vData Then
        usr免费号.Top = UsrEdit.Top
        If Usr自助打印.Visible Then
             usr免费号.Left = Usr自助打印.Left + Usr自助打印.Width
        ElseIf UsrGuoHao.Visible Then
             usr免费号.Left = UsrGuoHao.Left + UsrGuoHao.Width
        ElseIf UsrEdit.Visible Then
             usr免费号.Left = UsrEdit.Left + UsrEdit.Width: usr免费号.Top = UsrEdit.Top
        Else
            usr免费号.Left = 150
        End If
     End If
End Property

'zyk add 200412
Public Sub showwww()
    If GetPara("医院主页") = "" Then
        ctlshowwww.Visible = False
    Else
        ctlshowwww.Visible = True
    End If
End Sub


Public Property Let AllowSelfRegist(vData As Boolean)
    UsrGuoHao.Visible = vData
    If UsrEdit.Visible = False Then UsrGuoHao.Left = 150
End Property

Public Property Let AllowSelfPrint(vData As Boolean)
    Usr自助打印.Visible = vData
    If UsrEdit.Visible = False And UsrGuoHao.Visible = False Then
        Usr自助打印.Left = 150
    ElseIf UsrEdit.Visible = True And UsrGuoHao.Visible = False Then
        Usr自助打印.Left = UsrGuoHao.Left
    ElseIf UsrEdit.Visible = False And UsrGuoHao.Visible = True Then
        Usr自助打印.Left = UsrGuoHao.Left + UsrGuoHao.Width
    End If
End Property

Private Sub DoSoftFlag()
    Dim strTmp As String
    Dim strOEM As String
    On Error Resume Next
    Err.Clear
    
    strTmp = zlRegInfo("产品简名")
    If strTmp <> "-" Then
        lblOEM.Caption = strTmp & "软件"
        '处理状态栏图标的OEM策略
        If strTmp = "中联" Then
            Set imgFlag.Picture = LoadCustomPicture("Logo")
        Else
            strOEM = GetOEM(strTmp)
            Set imgFlag.Picture = LoadCustomPicture(strOEM)
            If Err <> 0 Then
                Err.Clear
                Set imgFlag.Picture = LoadCustomPicture("Logo")
            End If
        End If
        lblOEM.ToolTipText = ""
    End If
End Sub

Private Function GetOEM(ByVal strAsk As String) As String
    '-------------------------------------------------------------
    '功能：返回每个字线的ASCII码
    '参数：
    '返回：
    '-------------------------------------------------------------
    Dim intBit As Integer, iCount As Integer, blnCan As Boolean
    Dim strCode As String
    
    strCode = "OEM_"
    For intBit = 1 To Len(strAsk)
        '取每个字的ASCII码
        strCode = strCode & Hex(Asc(Mid(strAsk, intBit, 1)))
    Next
    GetOEM = strCode
End Function


Public Sub InitLoad()
    UsrCmd(0).Picture = ilsImage.ListImages("back")
    UsrCmd(1).Picture = ilsImage.ListImages("menu")
    UsrCmd(2).Picture = ilsImage.ListImages("up")
    UsrCmd(3).Picture = ilsImage.ListImages("down")
    UsrCmd(4).Picture = ilsImage.ListImages("list")
    
    tmrLoop.Enabled = True
    tmrMsg.Enabled = True
    tmrPlay.Enabled = True
    lblMsg.Caption = ""
                
    Call InitCommon(gcnOracle)
    Call DoSoftFlag
    
    'UsrHome.SetFocus
    EnterFocus UsrHome
    
End Sub

Private Sub usrQueryItem_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub usrQueryItem_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Public Sub InitNavigator(ByVal W As Single, ByVal H As Single)
    '功能:重新计算页面范围，并设置滚动标志
    
    UsrCmd(0).Visible = False
    UsrCmd(1).Visible = False
    UsrCmd(2).Visible = False
    UsrCmd(3).Visible = False
    UsrCmd(4).Visible = False
            
    usrQueryItem.MaxVsb = 0
    usrQueryItem.MaxHsb = 0
    
    usrQueryItem.ValueVsb = 0
    usrQueryItem.ValueHsb = 0
    
    If W > usrQueryItem.Width Then
        UsrCmd(0).Visible = True
        UsrCmd(1).Visible = True
                        
        usrQueryItem.MaxHsb = 0 - Int(0 - (W - usrQueryItem.Width) / 600)
    End If
    
    If H > usrQueryItem.Height Then
        UsrCmd(2).Visible = True
        UsrCmd(3).Visible = True
        UsrCmd(4).Visible = True
        
        usrQueryItem.MaxVsb = 0 - Int(0 - (H - usrQueryItem.Height) / 600)
    End If
        
    Call usrQueryItem_ChangeNavigator
    Call RefreshPostion
    
    usrQueryItem.FactWidth = W
    usrQueryItem.FactHeight = H
End Sub

Private Sub RefreshPostion()
'功能:计算页面范围导航器的状态，包括宽度
    Dim i As Long
    Dim vTmp As Single
            
    '江磊改于2003-8-12 编号：2491    目的：将右下面的按钮上移
    If UsrCmd(4).Visible Then vTmp = UsrCmd(4).Width + 120
    For i = 3 To 0 Step -1
        If UsrCmd(i).Visible Then
            UsrCmd(i).Left = UsrCmd(i + 1).Left - UsrCmd(i).Width - 30
            vTmp = vTmp + UsrCmd(i).Width + 30
        End If
    Next
    On Error Resume Next
    Call ResizeControl(PicOEM, 0, UsrHome.Top + UsrHome.Height + 100, PicOEM.ScaleWidth, picState.ScaleHeight - (UsrHome.Top + UsrHome.Height + 100))
    Call ResizeControl(picMsg, PicOEM.Left + PicOEM.Width, PicOEM.Top, picState.ScaleWidth - PicOEM.Width, PicOEM.Height)
    'Call ResizeControl(lblMsg, picMsg.ScaleWidth, (picMsg.ScaleHeight - lblMsg.Height) / 2, lblMsg.Width, lblMsg.Height)
    
End Sub

Private Sub usrQueryItem_RefreshNavigator(ByVal W As Single, ByVal H As Single)
    Call InitNavigator(W, H)
End Sub

Public Sub ShowSpecPage(ByVal PageNo As Long)
    '
    Call usrQueryItem_ConnectClick(PageNo, 0)
    
End Sub

Public Sub FirstChar(ByVal ch As String)
    Call pagChargeQuery.FirstChar(ch)
End Sub

Private Sub Usr自助打印_CommandClick()
    frmLisPrinter.Show , Me
End Sub

Private Sub usr免费号_CommandClick()
     frmFreeRegist.ShowMe Me
End Sub
