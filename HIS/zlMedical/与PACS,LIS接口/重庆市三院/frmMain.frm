VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmMain 
   Caption         =   "重庆市第三人民医院检验接口平台"
   ClientHeight    =   7815
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   11880
   StartUpPosition =   1  '所有者中心
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picContainer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000C&
      Height          =   1665
      Left            =   555
      ScaleHeight     =   1605
      ScaleWidth      =   3090
      TabIndex        =   0
      Top             =   1200
      Width           =   3150
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   705
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1244
      BandCount       =   1
      _CBWidth        =   11880
      _CBHeight       =   705
      _Version        =   "6.7.8988"
      Child1          =   "tbrThis"
      MinWidth1       =   4995
      MinHeight1      =   645
      Width1          =   8370
      Key1            =   "only"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   645
         Left            =   30
         TabIndex        =   2
         Top             =   30
         Width           =   11760
         _ExtentX        =   20743
         _ExtentY        =   1138
         ButtonWidth     =   1138
         ButtonHeight    =   1138
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ilsMenuHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   8
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   " 注销 "
               Key             =   "注销"
               Object.ToolTipText     =   "注销"
               Object.Tag             =   " 注销 "
               ImageKey        =   "Login"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "对码"
               Key             =   "对码"
               Object.ToolTipText     =   "对码"
               Object.Tag             =   "对码"
               ImageKey        =   "combo"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "接受"
               Key             =   "接受"
               Object.ToolTipText     =   "接受"
               Object.Tag             =   "接受"
               ImageKey        =   "Accept"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_1"
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "帮助"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "帮助"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ilsMenuHot 
      Left            =   8040
      Top             =   900
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6852
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6A72
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6C92
            Key             =   "Accept"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D4F4
            Key             =   "Send"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":DC6E
            Key             =   "Login"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":DF88
            Key             =   "Dept"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E702
            Key             =   "Diag"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14F64
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B7C6
            Key             =   "combo"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   7455
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMain.frx":22028
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15875
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFileLogin 
         Caption         =   "注销(&L)..."
      End
      Begin VB.Menu mnuFile_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "退出(&X)"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuOper 
      Caption         =   "操作(&O)"
      Begin VB.Menu mnuCodeCode 
         Caption         =   "检验项目对码(&1)"
         Index           =   0
      End
      Begin VB.Menu mnuOper_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOperAccept 
         Caption         =   "接受检验数据(&A)"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "查看(&V)"
      Begin VB.Menu mnuViewTool 
         Caption         =   "工具栏(&T)"
         Begin VB.Menu mnuViewToolButton 
            Caption         =   "标准按钮(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewToolText 
            Caption         =   "文本标签(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "状态栏(&S)"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuHelpTopic 
         Caption         =   "帮助主题(&T)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelp_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'（１）窗体级变量定义**************************************************************************************************
Private mblnStartUp As Boolean                          '窗体启动标志
Private mlngLoop As Long


Private Sub Form_Activate()
    
    If mblnStartUp = False Then Exit Sub
    mblnStartUp = False
        
    
    stbThis.Panels(2).Text = "用户：" & gstrDBUser
    If GetSetting("ZLSOFT", "注册信息\登陆信息", "SERVER", "") <> "" Then
         stbThis.Panels(2).Text = stbThis.Panels(2).Text & "@" & GetSetting("ZLSOFT", "注册信息\登陆信息", "SERVER", "")
    End If
    
    If InStr(1, gstrPrive, ";数据对码;") <= 0 And InStr(1, gstrPrive, ";接受结果;") <= 0 Then
        mnuOper.Visible = False
    Else
    
        If InStr(1, gstrPrive, ";数据对码;") <= 0 Then
            mnuCodeCode(0).Visible = False
            mnuOper_1.Visible = False
        End If
        
        If InStr(1, gstrPrive, ";接受结果;") <= 0 Then
            mnuOperAccept.Visible = False
            mnuOper_1.Visible = False
        End If
    End If
    
    stbThis.Panels(2).Text = stbThis.Panels(2).Text & "  姓名：" & UserInfo.姓名
    stbThis.Panels(2).Text = stbThis.Panels(2).Text & "  部门：" & UserInfo.部门
        
End Sub

Private Sub Form_Load()
    
    mblnStartUp = True
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    
    With picContainer
        .Left = 0
        .Top = IIf(cbrThis.Visible, cbrThis.Height, 0)
        .Width = Me.ScaleWidth - .Left
        .Height = Me.ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim frmThis As Form
    
    On Error Resume Next
    '关闭本部件窗体
    For Each frmThis In Forms
        If frmThis.Caption <> Me.Caption Then Unload frmThis
    Next

End Sub

Private Sub mnuCodeCode_Click(Index As Integer)
    Dim LngStyle As Long
    Dim blnShowed As Boolean

    
    blnShowed = False
    picContainer.BorderStyle = 0
    
    Select Case Index
    Case 0
        frmCombo.Show , Me
    End Select
    
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileLogin_Click()
    If MsgBox("你确定真的要注销吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    If CloseChildWindows(Me, Me) = False Then
        MsgBox "无法关闭部分窗体，注销操作中止！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Unload Me
    
    Call Main
    
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show 1, Me
End Sub

Private Sub mnuHelpTopic_Click()
    Call ShowHelp(Me.hWnd, Me.Name)
End Sub

Private Sub mnuOperAccept_Click()
    frmTaskAccept.Show 1, Me
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = mnuViewStatus.Checked
    Call Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    cbrThis.Visible = mnuViewToolButton.Checked
    Call Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim intLoop As Integer
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For intLoop = 1 To tbrThis.Buttons.Count
        tbrThis.Buttons(intLoop).Caption = IIf(mnuViewToolText.Checked, tbrThis.Buttons(intLoop).Tag, "")
    Next
    cbrThis.Bands(1).MinHeight = tbrThis.Height
    Call Form_Resize
    
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "注销"
        Call mnuFileLogin_Click
    Case "对码"
        Call mnuCodeCode_Click(0)
    Case "接受"
        Call mnuOperAccept_Click
    Case "帮助"
        Call mnuHelpTopic_Click
    Case "退出"
        Call mnuFileExit_Click
    End Select
End Sub
