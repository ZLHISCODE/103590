VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "CODEJO~1.OCX"
Begin VB.Form frmClientUpgradeManage 
   BackColor       =   &H80000005&
   Caption         =   "客户端升级管理"
   ClientHeight    =   6405
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11715
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Picture         =   "frmClientUpgradeManage.frx":0000
   ScaleHeight     =   6405
   ScaleWidth      =   11715
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PicBar 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1650
      Left            =   0
      Picture         =   "frmClientUpgradeManage.frx":803A
      ScaleHeight     =   1650
      ScaleWidth      =   37500
      TabIndex        =   3
      Top             =   570
      Width           =   37500
      Begin VB.PictureBox picNowTag 
         AutoSize        =   -1  'True
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   180
         Left            =   780
         Picture         =   "frmClientUpgradeManage.frx":D1724
         ScaleHeight     =   180
         ScaleWidth      =   315
         TabIndex        =   7
         Top             =   1470
         Width           =   315
      End
      Begin VB.Image imgArrow 
         Height          =   1125
         Index           =   2
         Left            =   8250
         Picture         =   "frmClientUpgradeManage.frx":D1A66
         Top             =   195
         Width           =   1125
      End
      Begin VB.Label lblPic 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客户端升级结果"
         Height          =   180
         Index           =   3
         Left            =   9555
         TabIndex        =   8
         Top             =   1155
         Width           =   1260
      End
      Begin VB.Image imgBtn 
         Height          =   960
         Index           =   3
         Left            =   9705
         Picture         =   "frmClientUpgradeManage.frx":D5D76
         Top             =   210
         Width           =   960
      End
      Begin VB.Image imgArrow 
         Height          =   1125
         Index           =   1
         Left            =   5085
         Picture         =   "frmClientUpgradeManage.frx":DAB21
         Top             =   195
         Width           =   1125
      End
      Begin VB.Image imgArrow 
         Height          =   1125
         Index           =   0
         Left            =   1965
         Picture         =   "frmClientUpgradeManage.frx":DEE31
         Top             =   195
         Width           =   1125
      End
      Begin VB.Image imgBtn 
         Height          =   825
         Index           =   0
         Left            =   600
         Picture         =   "frmClientUpgradeManage.frx":E3141
         ToolTipText     =   "进行服务器参数设置"
         Top             =   240
         Width           =   825
      End
      Begin VB.Image imgBtn 
         Height          =   825
         Index           =   1
         Left            =   3645
         Picture         =   "frmClientUpgradeManage.frx":E559D
         ToolTipText     =   "升级部件上传管理"
         Top             =   240
         Width           =   825
      End
      Begin VB.Image imgBtn 
         Height          =   825
         Index           =   2
         Left            =   6720
         Picture         =   "frmClientUpgradeManage.frx":E79F9
         ToolTipText     =   "对客户端升级参数设置"
         Top             =   240
         Width           =   825
      End
      Begin VB.Label lblPic 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "文件升级管理"
         Height          =   180
         Index           =   1
         Left            =   3528
         TabIndex        =   6
         Top             =   1152
         Width           =   1080
      End
      Begin VB.Label lblPic 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客户端升级配置"
         Height          =   180
         Index           =   2
         Left            =   6495
         TabIndex        =   5
         Top             =   1150
         Width           =   1260
      End
      Begin VB.Label lblPic 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "文件服务器配置"
         Height          =   180
         Index           =   0
         Left            =   372
         TabIndex        =   4
         Top             =   1152
         Width           =   1260
      End
   End
   Begin VB.Frame fraCaption 
      BackColor       =   &H00FFFFFF&
      Height          =   120
      Left            =   -135
      TabIndex        =   1
      Top             =   1050
      Width           =   10305
   End
   Begin XtremeSuiteControls.TabControl tbcThis 
      Height          =   1380
      Left            =   -15
      TabIndex        =   2
      Top             =   1860
      Width           =   1275
      _Version        =   589884
      _ExtentX        =   2249
      _ExtentY        =   2434
      _StockProps     =   64
   End
   Begin ComctlLib.ImageList imgList 
      Left            =   10710
      Top             =   195
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   12
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmClientUpgradeManage.frx":E9E55
            Key             =   "服务器配置-正常"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmClientUpgradeManage.frx":EC2BF
            Key             =   "服务器配置-高亮"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmClientUpgradeManage.frx":EE729
            Key             =   "服务器配置-按下"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmClientUpgradeManage.frx":F0B93
            Key             =   "升级文件管理-正常"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmClientUpgradeManage.frx":F2FFD
            Key             =   "升级文件管理-高亮"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmClientUpgradeManage.frx":F5467
            Key             =   "升级文件管理-按下"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmClientUpgradeManage.frx":F78D1
            Key             =   "客户端升级设置-正常"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmClientUpgradeManage.frx":F9D3B
            Key             =   "客户端升级管理-高亮"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmClientUpgradeManage.frx":FC1A5
            Key             =   "客户端升级管理-按下"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmClientUpgradeManage.frx":FE60F
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmClientUpgradeManage.frx":101661
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmClientUpgradeManage.frx":1046B3
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblcaption 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "客户端升级管理"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   285
      TabIndex        =   0
      Top             =   150
      Width           =   1470
   End
End
Attribute VB_Name = "frmClientUpgradeManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mobjPage(4) As Object
Private mblnMove As Boolean '控制图标按钮显示状态
Private mpeSelect As PageEnum '当前选中功能模块 0-服务器配置 1-升级文件管理 2-客户端升级设置
Private mblnLoad As Boolean '加载判断值 ture - 正在加载  false - 加载完成
Private mintPage As Integer
Private mstrFunc As String '记录模块功能权限字符串

'页面索引
Private Enum PageEnum
    PE_文件服务器配置 = 0
    PE_文件升级管理 = 1
    PE_客户端升级配置 = 2
    PE_客户端升级概况 = 3
End Enum

'按钮状态
Private Enum ImageState
    IS_正常 = 1
    IS_高亮 = 2
    IS_按下 = 3
End Enum

Private Enum PageBack
    PB_文件服务器配置 = 10
    PB_文件升级管理 = 11
    PB_客户端升级配置 = 12
End Enum

Public Function SupportPrint() As Boolean
'返回本窗口是否支持打印，供主窗口调用
    SupportPrint = mobjPage(mpeSelect).SupportPrint
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'供主窗口调用，实现具体的打印工作
'如果没有可打印的，就留下一个空的接口
    '参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    Call mobjPage(mpeSelect).SubPrint(bytMode)
End Sub

Private Sub InitTbcthis()
    On Error GoTo errH:
    mblnLoad = True
    With tbcThis
        .RemoveAll
        .InsertItem PE_文件服务器配置, "服务器配置", mobjPage(PE_文件服务器配置).hwnd, PE_文件服务器配置 * 3 + IS_正常
        .InsertItem PE_文件升级管理, "升级文件管理", mobjPage(PE_文件升级管理).hwnd, PE_文件升级管理 * 3 + IS_正常
        .InsertItem PE_客户端升级配置, "客户端设置", mobjPage(PE_客户端升级配置).hwnd, PE_客户端升级配置 * 3 + IS_正常
        .InsertItem PE_客户端升级概况, "客户端升级概况", mobjPage(PE_客户端升级概况).hwnd, PE_客户端升级概况 * 3 + IS_正常
    End With
    mblnLoad = False
    Exit Sub
errH:
End Sub

Private Sub Form_Load()
    On Error GoTo errH:
    Set mobjPage(PE_文件服务器配置) = New frmClientUpgradeSeverConfigure
    Set mobjPage(PE_文件升级管理) = New frmClientUpgradeFileManage
    Set mobjPage(PE_客户端升级配置) = New frmClientUpgradeConfigure
    Set mobjPage(PE_客户端升级概况) = New frmClientUpgradeProfile
    
    '获取当前用户拥有的功能权限
    mstrFunc = GetProgFuncs("0307")
    
    If Not CheckAndAdjustMustTable("ZLUPGRADESERVER", , False, , False) Then
        MsgBox "请将数据库升级至10.35.40以后的版本再使用该功能！", vbInformation, gstrSysName
        Me.Tag = "HIDE"
        Exit Sub
    End If
    
    mintPage = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Caption, "选择页签", "0"))
    Call InitTbcthis
    Call imgBtn_Click(mintPage) '默认显示服务器配置页面
    Exit Sub
errH:
    Me.Tag = "HIDE"
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    If Me.Tag = "HIDE" Then Me.Hide
    tbcThis.Top = PicBar.Top + PicBar.Height - 350
    tbcThis.Width = Me.ScaleWidth + 50
    tbcThis.Height = Me.ScaleHeight - tbcThis.Top + 10
    imgBtn(0).Top = PicBar.Height / 2 - imgBtn(0).Height / 2 - 180
    imgBtn(1).Top = PicBar.Height / 2 - imgBtn(1).Height / 2 - 180
    imgBtn(2).Top = PicBar.Height / 2 - imgBtn(2).Height / 2 - 180
    lblPic(0).Top = imgBtn(0).Top + imgBtn(0).Height + 100
    lblPic(0).Left = imgBtn.Item(0).Left + (imgBtn.Item(0).Width / 2) - (lblPic(0).Width / 2)
    lblPic(1).Top = lblPic(0).Top
    lblPic(1).Left = imgBtn.Item(1).Left + (imgBtn.Item(1).Width / 2) - (lblPic(1).Width / 2)
    lblPic(2).Top = lblPic(0).Top
    lblPic(2).Left = imgBtn.Item(2).Left + (imgBtn.Item(2).Width / 2) - (lblPic(2).Width / 2)
    picNowTag.Top = PicBar.Height - picNowTag.Height
    fraCaption.Width = Me.ScaleWidth + 200
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mobjPage(PE_文件服务器配置) Is Nothing Then Unload mobjPage(PE_文件服务器配置)
    If Not mobjPage(PE_文件升级管理) Is Nothing Then Unload mobjPage(PE_文件升级管理)
    If Not mobjPage(PE_客户端升级配置) Is Nothing Then Unload mobjPage(PE_客户端升级配置)
    If Not mobjPage(PE_客户端升级概况) Is Nothing Then Unload mobjPage(PE_客户端升级概况)
    Set mobjPage(PE_文件服务器配置) = Nothing
    Set mobjPage(PE_文件升级管理) = Nothing
    Set mobjPage(PE_客户端升级配置) = Nothing
    Set mobjPage(PE_客户端升级概况) = Nothing
End Sub

Private Sub imgBtn_Click(Index As Integer)
    imgBtn.Item(mpeSelect).Picture = imgList.ListImages.Item(mpeSelect * 3 + IS_正常).Picture
    lblPic.Item(mpeSelect).Font.Bold = False
    Select Case Index
        Case 0, 1, 2, 3
            imgBtn.Item(Index).Picture = imgList.ListImages.Item(Index * 3 + IS_高亮).Picture   '图标按钮状态切换
            picNowTag.Left = imgBtn.Item(Index).Left + (imgBtn.Item(Index).Width / 2) - (picNowTag.Width / 2)
            lblPic.Item(Index).Font.Bold = True
            tbcThis.Item(Index).Selected = True
            mpeSelect = Index
    End Select
End Sub

Private Sub imgBtn_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If mpeSelect <> Index Then
        imgBtn.Item(Index).Picture = imgList.ListImages.Item(Index * 3 + IS_按下).Picture
    End If
End Sub

Private Sub imgBtn_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If mblnMove = False And mpeSelect <> Index Then
        imgBtn.Item(Index).Picture = imgList.ListImages.Item(Index * 3 + IS_高亮).Picture
        mblnMove = True
    End If
End Sub

Private Sub lblPic_Click(Index As Integer)
    Call imgBtn_Click(Index)
End Sub

Private Sub PicBar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Integer
    
    For i = 0 To 3
        If mpeSelect <> i Then
            imgBtn.Item(i).Picture = imgList.ListImages.Item(i * 3 + IS_正常).Picture
            lblPic.Item(i).Font.Bold = False
        End If
    Next
    mblnMove = False
End Sub

Private Sub tbcThis_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If mblnLoad And mintPage <> 0 Then Exit Sub
    Me.Refresh
    Call mobjPage(Item.Index).RefreshData
    Call mobjPage(Item.Index).SetMenu
    '若功能字符串为空，则表示拥有全部权限，将不再做权限控制
    If mstrFunc <> "" Then
        Call mobjPage(Item.Index).SetControlEnable(mstrFunc)
    End If
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Caption, "选择页签", Item.Index
End Sub

