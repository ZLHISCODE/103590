VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmUpgradeManage 
   BackColor       =   &H80000005&
   Caption         =   "客户端升级管理"
   ClientHeight    =   5925
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9765
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "frmUpgradeManage.frx":0000
   ScaleHeight     =   5925
   ScaleWidth      =   9765
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PicBar 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1650
      Left            =   0
      Picture         =   "frmUpgradeManage.frx":803A
      ScaleHeight     =   1650
      ScaleWidth      =   37500
      TabIndex        =   5
      Top             =   1155
      Width           =   37500
      Begin VB.PictureBox picNowTag 
         AutoSize        =   -1  'True
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   180
         Left            =   780
         Picture         =   "frmUpgradeManage.frx":D1724
         ScaleHeight     =   180
         ScaleWidth      =   315
         TabIndex        =   9
         Top             =   1470
         Width           =   315
      End
      Begin VB.Image imgArrow 
         Height          =   1125
         Index           =   1
         Left            =   5085
         Picture         =   "frmUpgradeManage.frx":D1A66
         Top             =   195
         Width           =   1125
      End
      Begin VB.Image imgArrow 
         Height          =   1125
         Index           =   0
         Left            =   1965
         Picture         =   "frmUpgradeManage.frx":D5D76
         Top             =   195
         Width           =   1125
      End
      Begin VB.Image imgBtn 
         Height          =   825
         Index           =   0
         Left            =   600
         Picture         =   "frmUpgradeManage.frx":DA086
         Top             =   240
         Width           =   825
      End
      Begin VB.Image imgBtn 
         Height          =   825
         Index           =   1
         Left            =   3645
         Picture         =   "frmUpgradeManage.frx":DC4E2
         Top             =   240
         Width           =   825
      End
      Begin VB.Image imgBtn 
         Height          =   825
         Index           =   2
         Left            =   6720
         Picture         =   "frmUpgradeManage.frx":DE93E
         Top             =   225
         Width           =   825
      End
      Begin VB.Label lblPic 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "文件升级管理"
         Height          =   180
         Index           =   1
         Left            =   3528
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   6
         Top             =   1152
         Width           =   1260
      End
   End
   Begin VB.Frame fraCaption 
      BackColor       =   &H00FFFFFF&
      Height          =   120
      Left            =   -135
      TabIndex        =   3
      Top             =   1050
      Width           =   10305
   End
   Begin XtremeSuiteControls.TabControl tbcThis 
      Height          =   1380
      Left            =   -15
      TabIndex        =   4
      Top             =   2445
      Width           =   1275
      _Version        =   589884
      _ExtentX        =   2249
      _ExtentY        =   2434
      _StockProps     =   64
   End
   Begin ComctlLib.ImageList imgList 
      Left            =   8685
      Top             =   15
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
            Picture         =   "frmUpgradeManage.frx":E0D9A
            Key             =   "服务器配置-正常"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmUpgradeManage.frx":E3204
            Key             =   "服务器配置-高亮"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmUpgradeManage.frx":E566E
            Key             =   "服务器配置-按下"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmUpgradeManage.frx":E7AD8
            Key             =   "升级文件管理-正常"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmUpgradeManage.frx":E9F42
            Key             =   "升级文件管理-高亮"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmUpgradeManage.frx":EC3AC
            Key             =   "升级文件管理-按下"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmUpgradeManage.frx":EE816
            Key             =   "客户端升级设置-正常"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmUpgradeManage.frx":F0C80
            Key             =   "客户端升级管理-高亮"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmUpgradeManage.frx":F30EA
            Key             =   "客户端升级管理-按下"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmUpgradeManage.frx":F5554
            Key             =   "背景1"
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmUpgradeManage.frx":1BEC4E
            Key             =   "背景2"
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmUpgradeManage.frx":288348
            Key             =   "背景3"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblcaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " 客户端升级设置：对客户端升级参数设置"
      Height          =   180
      Index           =   2
      Left            =   900
      TabIndex        =   2
      Top             =   840
      Width           =   3330
   End
   Begin VB.Label lblcaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " 服务器配置：进行服务器参数设置       升级文件管理：升级部件上传管理 "
      Height          =   180
      Index           =   1
      Left            =   900
      TabIndex        =   1
      Top             =   600
      Width           =   6210
   End
   Begin VB.Image imgCaption 
      Height          =   480
      Left            =   225
      Picture         =   "frmUpgradeManage.frx":351A42
      Stretch         =   -1  'True
      Top             =   555
      Width           =   480
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
Attribute VB_Name = "frmUpgradeManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mobjPage(3) As Object
Private mblnMove As Boolean '控制图标按钮显示状态
Private mpeSelect As PageEnum '当前选中功能模块 0-服务器配置 1-升级文件管理 2-客户端升级设置
Private mblnLoad As Boolean '加载判断值 ture - 正在加载  false - 加载完成
Private mintPage As Integer

'页面索引
Private Enum PageEnum
    PE_文件服务器配置 = 0
    PE_文件升级管理 = 1
    PE_客户端升级配置 = 2
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
    End With
    mblnLoad = False
    Exit Sub
errH:
End Sub

Private Sub Form_Load()
    On Error GoTo errH:
    Set mobjPage(PE_文件服务器配置) = New frmFilesSeverConfigure
    Set mobjPage(PE_文件升级管理) = New frmFilesUpgradeManage
    Set mobjPage(PE_客户端升级配置) = New frmClientsUpgradeConfigure
    
    If CheckSystem = False Then
        MsgBox "请将数据库升级至10.35.40以后的版本再使用该功能！"
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

    Set mobjPage(PE_文件服务器配置) = Nothing
    Set mobjPage(PE_文件升级管理) = Nothing
    Set mobjPage(PE_客户端升级配置) = Nothing
End Sub

Private Sub imgBtn_Click(Index As Integer)
    Dim i As Integer
    
    For i = 0 To 2
        If i = Index Then
            imgBtn.Item(i).Picture = imgList.ListImages.Item(i * 3 + IS_高亮).Picture   '图标按钮状态切换
'            PicBar.Picture = imgList.ListImages.Item(PB_文件服务器配置 + i).Picture '背景切换
            picNowTag.Left = imgBtn.Item(i).Left + (imgBtn.Item(i).Width / 2) - (picNowTag.Width / 2)
            lblPic.Item(i).Font.Bold = True
            lblPic.Item(i).Left = imgBtn.Item(i).Left + (imgBtn.Item(i).Width / 2) - (lblPic.Item(i).Width / 2)
            tbcThis.Item(i).Selected = True
            mpeSelect = i
        Else
            imgBtn.Item(i).Picture = imgList.ListImages.Item(i * 3 + IS_正常).Picture
            lblPic.Item(i).Font.Bold = False
            lblPic.Item(i).Left = imgBtn.Item(i).Left + (imgBtn.Item(i).Width / 2) - (lblPic.Item(i).Width / 2)
        End If
    Next
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
    
    For i = 0 To 2
        If mpeSelect <> i Then
            imgBtn.Item(i).Picture = imgList.ListImages.Item(i * 3 + IS_正常).Picture
            lblPic.Item(i).Font.Bold = False
        End If
    Next
    mblnMove = False
End Sub

Private Sub tbcThis_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
'    If mobjPage(Item.Index).blnRefreshData = True Then
'        mobjPage(Item.Index).RefreshData
'        mobjPage(Item.Index).blnRefreshData = False
'    End If
    If mblnLoad And mintPage <> 0 Then Exit Sub
    Me.Refresh
    Call mobjPage(Item.Index).RefreshData
    Call mobjPage(Item.Index).SetMenu
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Caption, "选择页签", Item.Index
End Sub

Public Function SetFormRefresh(intPage As Integer, Optional blnRefresh As Boolean = True)
    mobjPage(intPage).blnRefreshData = blnRefresh
End Function

Private Function CheckSystem() As Boolean
'    检查ZLUPGRADESERVER表是否存在，存在即可使用升级管理工具
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errH:
    
    strSQL = "select count(*) as 存在  from all_tables where table_name = 'ZLUPGRADESERVER' and owner = 'ZLTOOLS'"
    Call OpenRecordset(rsTemp, strSQL, Me.Caption)
    
    If rsTemp!存在 = "0" Then
        CheckSystem = False
    Else
        CheckSystem = True
    End If

    Exit Function
errH:
    MsgBox err.Description, vbInformation, "中联软件"
    If False Then
        Resume
    End If
End Function
