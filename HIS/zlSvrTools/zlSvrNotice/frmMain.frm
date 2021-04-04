VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmMain 
   Caption         =   "自动提醒服务"
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8175
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5190
   ScaleWidth      =   8175
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picNotify 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   5100
      ScaleHeight     =   345
      ScaleWidth      =   225
      TabIndex        =   2
      Top             =   4110
      Visible         =   0   'False
      Width           =   225
   End
   Begin MSWinsockLib.Winsock winSock 
      Left            =   3915
      Top             =   4095
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "FRCHEN"
   End
   Begin VB.Timer tmrMessage 
      Interval        =   1
      Left            =   225
      Top             =   4005
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   4830
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMain.frx":038A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8017
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "2017/7/28"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1323
            MinWidth        =   1323
            TextSave        =   "13:21"
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
   Begin MSComctlLib.ListView lvw 
      Height          =   2685
      Left            =   210
      TabIndex        =   1
      Top             =   1125
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   4736
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ils16"
      SmallIcons      =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "工作站"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "端口号"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "用户姓名"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "数据库用户"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "IP地址"
         Object.Width           =   2540
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   8175
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbrThis"
      MinHeight1      =   720
      Width1          =   315
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   720
         Left            =   30
         TabIndex        =   4
         Top             =   30
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ilsMenu"
         HotImageList    =   "ilsHotMenu"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "启动"
               Key             =   "启动"
               Object.ToolTipText     =   "启动提醒服务"
               Object.Tag             =   "启动"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "停止"
               Key             =   "停止"
               Object.ToolTipText     =   "停止提醒服务"
               Object.Tag             =   "停止"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "参数"
               Key             =   "参数"
               Object.ToolTipText     =   "参数设置"
               Object.Tag             =   "参数"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "帮助"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "帮助"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "退出"
               Object.ToolTipText     =   "退出服务"
               Object.Tag             =   "退出"
               ImageIndex      =   4
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   7485
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0C1E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsMenu 
      Left            =   6825
      Top             =   705
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2A28
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":31A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":391C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4096
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":42B0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsHotMenu 
      Left            =   6210
      Top             =   705
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":44CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4C44
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":53BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5B38
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5D52
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image imgNotify 
      Height          =   240
      Left            =   5295
      Picture         =   "frmMain.frx":5F6C
      Top             =   90
      Width           =   240
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFileStart 
         Caption         =   "启动服务(&S)"
      End
      Begin VB.Menu mnuFileStop 
         Caption         =   "停止服务(&D)"
      End
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileParm 
         Caption         =   "参数设置(&P)"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFile_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileLogin 
         Caption         =   "重新登录(&L)"
      End
      Begin VB.Menu mnuFile_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileHide 
         Caption         =   "隐藏提醒服务(&H)"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "退出服务(&X)"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuHelpTopic 
         Caption         =   "帮助主题(&T)"
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "&Web上的中联"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "中联主页(&H)"
         End
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "中联论坛(&F)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "发送反馈(&K)..."
         End
      End
      Begin VB.Menu mnuHelp_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mblnStartUp As Boolean
Private mlngCount As Long
Private mlngPort As Long
Private mstrLocalIP As String
Private mblnCancel As Boolean
Private mblnTest As Boolean

Private Sub cmdFunc_Click(Index As Integer)
    Select Case Index
    Case 0
        If tbrThis.Buttons("注销").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("注销"))
    Case 1
        If tbrThis.Buttons("启动").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("启动"))
    Case 2
        If tbrThis.Buttons("停止").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("停止"))
    Case 3
        If tbrThis.Buttons("退出").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("退出"))
    Case 4
        If tbrThis.Buttons("端口").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("端口"))
    Case 5
        '
    End Select
End Sub

Private Sub Form_Activate()
    Dim rs As New ADODB.Recordset
    Dim varParam As Variant
    Dim strSQL As String
    
    If mblnStartUp = False Then Exit Sub
    mblnStartUp = False
                                        
    
    Me.Caption = Me.Caption & " - [" & gstrUserName & IIf(gstrServer = "", "", "@" & gstrServer) & "]"
    
    gstrSysName = gstrProductName & "软件"
    SaveSetting "ZLSOFT", "注册信息", UCase("gstrSysName"), gstrSysName
    
    Call ApplyOEM(stbThis)
    Call ApplyOEM_Picture(Me, "Icon")
    
   
    '格式:服务器;端口号;状态
    strSQL = "SELECT 参数值 FROM zloptions WHERE 参数号=[1]"
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, 7)
    
    If rs.BOF = False Then
        varParam = Split(zlCommFun.NVL(rs("参数值").Value, ""), ";")
        If UBound(varParam) < 2 Then
            mlngPort = 9999
            mstrLocalIP = winSock.LocalIP
        Else
            mlngPort = Val(varParam(1))            '取服务器配置的端口号
            mstrLocalIP = Trim(varParam(0))
        End If
    Else
        mlngPort = 9999
        mstrLocalIP = winSock.LocalIP
    End If

    '启动服务
    Call tbrThis_ButtonClick(tbrThis.Buttons("启动"))
        
    DoEvents
    
    Call AddIcon(picNotify.hWnd, imgNotify.Picture, Me.Caption)
    
    If rs.State = adStateOpen Then rs.Close
    
    Call mnuFileHide_Click
    
End Sub

Private Sub Form_Load()
    
    tmrMessage.Interval = 1
    tmrMessage.Enabled = False
    mblnStartUp = True
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Select Case Me.WindowState
    Case 1
        mnuFileHide.Caption = "显示提醒服务(&O)"
        Me.Hide
        Exit Sub
    End Select
    
    With lvw
        .Left = 0
        .Top = cbrThis.Height
        .Height = Me.ScaleHeight - .Top - stbThis.Height
        .Width = Me.ScaleWidth - .Left
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    mblnCancel = False
    
    If MsgBox("你是否真的要退出自动提醒服务？", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then
        Cancel = True
        mblnCancel = True
        Exit Sub
    End If

    On Error Resume Next
    If gcnOracle.State <> adStateClosed Then gcnOracle.Close
    Call RemoveIcon(picNotify.hWnd)
    
End Sub

Private Sub lvw_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If lvw.SortKey = ColumnHeader.Index - 1 Then
        lvw.SortOrder = IIf(lvw.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        lvw.SortKey = ColumnHeader.Index - 1
        lvw.SortOrder = lvwAscending
    End If
End Sub


Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileHide_Click()
    If mnuFileHide.Caption = "隐藏提醒服务(&H)" Then
        Me.WindowState = 1
        Me.Hide
        mnuFileHide.Caption = "显示提醒服务(&O)"
    Else
        Me.WindowState = 0
        Me.Show
        mnuFileHide.Caption = "隐藏提醒服务(&H)"
    End If
End Sub

Private Sub mnuFileLogin_Click()
    'If MsgBox("你是否真的要注销消息提醒服务？", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then Exit Sub
    
    Unload Me
    If mblnCancel = False Then Call Main
End Sub

Private Sub mnuFileParm_Click()
    mnuFileHide.Enabled = False
    
    If frmParam.ShowEdit(Me, mlngPort, mstrLocalIP, IIf(tbrThis.Buttons("启动").Enabled, 0, 1)) Then
            
    End If
    mnuFileHide.Enabled = True
End Sub


Private Sub mnuFileStart_Click()
    Select Case StartServer(mlngPort, mstrLocalIP)
    Case 0
        Call AdjustEnabledState(1)
    Case 1
        stbThis.Panels(2).Text = "错误：端口地址正在使用！"
    End Select
End Sub

Private Sub mnuFileStop_Click()
    winSock.Close
    tmrMessage.Enabled = False
    stbThis.Panels(2).Text = "提醒服务已停止。"
    lvw.ListItems.Clear

    Call AdjustEnabledState(2)
End Sub


Private Sub mnuHelpAbout_Click()
    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
End Sub

Private Sub mnuHelpTopic_Click()
    Shell "hh.exe  zl9SvrNotice.chm", vbNormalFocus
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hWnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hWnd)
End Sub

Private Sub picNotify_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    '--------------------------------------------------------------------------------------------------
    '功能:  处理picNotify的各种处理事件,主要用于自动提醒相关功能(陈渝编写)
    '--------------------------------------------------------------------------------------------------

    Select Case Hex(x) '
        Case "1E3C"     'Right-Button-Down
        Case "1E4B"     'Right-Button-Up
            Me.PopupMenu mnuFile
        Case "1830"     'Right-Button-Down LARGE FONTS '
        Case "1E1E"     'Left-Button-up
        Case "1E0F"     'Left-Button-Down '
        Case "1E2D"     'Left-Button-Double-Click '
            If mnuFileHide.Enabled Then Call mnuFileHide_Click
        Case "1824"     'Left-Button-Double-Click LARGE FONTS
        Case "1E5A"     'Right-Button-Double-Click '
    End Select '

End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "启动"
        Call mnuFileStart_Click
    Case "停止"
        Call mnuFileStop_Click
    Case "参数"
        Call mnuFileParm_Click
    Case "退出"
        Call mnuFileExit_Click
    End Select
End Sub

Private Sub tmrMessage_Timer()
    
    If tmrMessage.Interval = 1 Then tmrMessage.Interval = 60000
    'Call ShowAlert(Me)
    
    tmrMessage.Enabled = False
    Call CheckNoticeAll
    tmrMessage.Enabled = True
    
End Sub

Private Sub winSock_DataArrival(ByVal bytesTotal As Long)
    Dim varItem As Variant
    Dim strData As String
    Dim varData As Variant
    Dim lngLoop As Long
    Dim strSQL As String
    Dim objItem As ListItem
    Dim strTmp As String
    
    On Error Resume Next
    
    winSock.GetData strData
    
    lngLoop = InStr(strData, "]")
    If lngLoop = 0 Then Exit Sub
    
    strTmp = Mid(strData, 2, lngLoop - 2)
    strData = Mid(strData, lngLoop + 1)
    
    Select Case strTmp
    Case "SYS-COMPUTER"             '工作站回答有关连接情况
        
        '格式:机器名;端号号;用户名;用户编码
        
        varData = Split(strData, ";")
        
        If UCase(varData(5)) = UCase(mstrLocalIP) Then
            
            For lngLoop = lvw.ListItems.Count To 1 Step -1
                
                Set objItem = lvw.ListItems(lngLoop)
                If UCase(objItem.Text) = UCase(varData(0)) And _
                    UCase(objItem.SubItems(1)) = UCase(varData(1)) And _
                    UCase(objItem.SubItems(2)) = UCase(varData(2)) And _
                    UCase(objItem.SubItems(3)) = UCase(varData(3)) Then
                    
                    lvw.ListItems.Remove lngLoop
                    Exit For
                    
                End If
            Next
            
            Set objItem = lvw.ListItems.Add(, , varData(0), 1, 1)
            objItem.SubItems(1) = varData(1)
            objItem.SubItems(2) = varData(2)
            objItem.SubItems(3) = varData(3)
            objItem.SubItems(4) = GetClientIP(objItem.Text)
            objItem.ListSubItems(1).Tag = Val(varData(4))
            
        End If
        
    Case "SYS-DISCONNECT"           '工作站发来的下线消息
    
        '格式:机器名;端号号;用户名;用户编码
        
        varData = Split(strData, ";")
        If UCase(varData(5)) = UCase(mstrLocalIP) Then
            For lngLoop = lvw.ListItems.Count To 1 Step -1
                
                Set objItem = lvw.ListItems(lngLoop)
                If UCase(objItem.Text) = UCase(varData(0)) And _
                    UCase(objItem.SubItems(1)) = UCase(varData(1)) And _
                    UCase(objItem.SubItems(2)) = UCase(varData(2)) And _
                    UCase(objItem.SubItems(3)) = UCase(varData(3)) Then
                    
                    lvw.ListItems.Remove lngLoop
                    Exit For
                    
                End If
            Next
        End If
        
    Case "SYS-STARTUP"              '工作站启动时发来的请求进行启动提醒消息
    
        '格式:机器名;端号号;用户名;部门名称
        
        varData = Split(strData, ";")
        If UCase(varData(5)) = UCase(mstrLocalIP) Then
            For lngLoop = lvw.ListItems.Count To 1 Step -1
                
                Set objItem = lvw.ListItems(lngLoop)
                If UCase(objItem.Text) = UCase(varData(0)) And _
                    UCase(objItem.SubItems(1)) = UCase(varData(1)) And _
                    UCase(objItem.SubItems(2)) = UCase(varData(2)) And _
                    UCase(objItem.SubItems(3)) = UCase(varData(3)) Then
                            
                    '工作站进行启动检查请求
                    tmrMessage.Enabled = False
                    Call CheckNoticeOne(objItem, True)
                    tmrMessage.Enabled = True
                    
                    Exit For
                    
                End If
            Next
        End If
    Case "SYS-READED"               '工作站发来的回置消息已读标志
    
        '格式:提醒序号;用户名
        
        varData = Split(strData, ";")
        
        If UCase(varData(5)) = UCase(mstrLocalIP) Then
            strSQL = "Zl_Zlnoticerec_Edit(1," & Val(varData(0)) & ",'" & varData(1) & "',Null,Null,Null,1,Null)"
            On Error GoTo errHand
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        End If
    Case "SYS-TEST"
        mblnTest = True
    End Select
    
    Exit Sub
    
errHand:
'    gcnOracle.RollbackTrans
End Sub

Private Function StartServer(Optional ByVal Port As Long = 1024, Optional ByVal LocalIP As String = "") As Long
    '--------------------------------------------------------------------------------------------------------
    '功能:启动自动提醒服务
    '--------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    
    '初始化
    winSock.Close
    winSock.Protocol = sckUDPProtocol
    
    stbThis.Panels(2).Text = "正在启动自动提醒服务...."
    DoEvents
    
    '开始启动
    Err = 0
    On Error Resume Next
    
'    winSock.Bind Port
    winSock.LocalPort = Port
    winSock.SendData ""
    
    If Err <> 0 Then
        MsgBox Err.Description, , gstrSysName
        
        stbThis.Panels(2).Text = ""
        StartServer = 1
                
        Exit Function
    End If
    On Error GoTo 0
    
    stbThis.Panels(2).Text = "提醒服务已启动(IP:" & LocalIP & " Port:" & Port & ")。"
    
    StartServer = 0
    
End Function

Private Sub AdjustEnabledState(ByVal bytMode As Byte)
    If bytMode = 1 Then
        '已经启动
        tmrMessage.Enabled = True
        tbrThis.Buttons("启动").Enabled = False
        tbrThis.Buttons("停止").Enabled = True
        
        mnuFileStart.Enabled = False
        mnuFileStop.Enabled = True
    Else
        '已经停止
        tmrMessage.Enabled = False
        mnuFileStart.Enabled = True
        mnuFileStop.Enabled = False
    End If
    
    
    tbrThis.Buttons("停止").Enabled = mnuFileStop.Enabled
    tbrThis.Buttons("启动").Enabled = mnuFileStart.Enabled
End Sub

Public Function UpdateRefresh(ByVal lngNewPort As Long, ByVal strLocalIP As String)
    '如果已启动服务，则必须先停止
    If tbrThis.Buttons("启动").Enabled = False Then
        winSock.Close
        tmrMessage.Enabled = False
    End If
    
    '检查新端口号是否有效
    If StartServer(lngNewPort, strLocalIP) <> 0 Then
        MsgBox "设置的端口号无效或冲突！", vbOKOnly, gstrSysName
        
        '如果原来处于启动状态，出错后还要重新启动
        If tbrThis.Buttons("启动").Enabled = False Then
            Call tbrThis_ButtonClick(tbrThis.Buttons("启动"))
        End If
        
        Exit Function
    End If
    
    '如果原来处于启动状态，则要按新端口号启动服务
    If tbrThis.Buttons("启动").Enabled = False Then
        Call AdjustEnabledState(1)
    End If
    
    mlngPort = lngNewPort
    
    UpdateRefresh = True
    
End Function

Private Function CheckNoticeAll() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:执行提醒检查,并发送提醒
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    Dim strMachine As String
    
    For lngLoop = 1 To lvw.ListItems.Count
        lvw.ListItems(lngLoop).ListSubItems(2).Tag = ""
    Next
    
    strSQL = "SELECT DISTINCT USERNAME,TERMINAL FROM GV$Session WHERE USERNAME IS NOT NULL"
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rs.BOF = False Then
        Do While Not rs.EOF
            strMachine = Trim(NVL(rs("TERMINAL").Value))
            If InStr(strMachine, "\") > 0 Then strMachine = Mid(strMachine, InStr(strMachine, "\") + 1)
            
            If strMachine <> "" Then

                If Chr(0) = Right(strMachine, 1) Then
                    strMachine = Mid(strMachine, 1, Len(strMachine) - 1)
                End If
                
                For lngLoop = 1 To lvw.ListItems.Count
                    If lvw.ListItems(lngLoop).ListSubItems(2).Tag = "" Then
                        If UCase(lvw.ListItems(lngLoop).Text) = UCase(strMachine) And _
                            UCase(lvw.ListItems(lngLoop).SubItems(3)) = UCase(NVL(rs("USERNAME").Value)) Then
                            
                            lvw.ListItems(lngLoop).ListSubItems(2).Tag = "CHECKED"
                            Call CheckNoticeOne(lvw.ListItems(lngLoop))
                            
                        End If
                    End If
                Next
            End If
            rs.MoveNext
        Loop
    End If
    
    '删除已经异常终止的工作站
    For lngLoop = lvw.ListItems.Count To 1 Step -1
        If lvw.ListItems(lngLoop).ListSubItems(2).Tag = "" Then
            lvw.ListItems.Remove lngLoop
        End If
    Next
    
End Function

Private Function CheckNoticeOne(ByVal Item As MSComctlLib.ListItem, Optional ByVal StartUp As Boolean = False) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:检查指定用户的提醒消息,并发送提醒
    '------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim lngLoop As Long
    Dim str提醒内容 As String
    Dim strSyss As String
    Dim strDepts As String
    
    mblnTest = False
    Item.Tag = ""
    
    '1.检查此人是否还处于登录状态(因为工作站异常退出时)
    
    '2.找出用户所属的系统及部门名称
    strSQL = "Select 编号, 名称, 共享号, 所有者, 安装日期, 正常安装, 版本号 From zlSystems "
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rs.BOF Then Exit Function
    
    strSyss = "0"
    strDepts = "''"
    
    Do While Not rs.EOF
    
        On Error Resume Next
        
        Err = 0
        strSQL = "Select R.部门id " & _
                        " From " & rs("所有者").Value & ".上机人员表 U," & rs("所有者").Value & ".人员表 P," & rs("所有者").Value & ".部门人员 R" & _
                        " Where U.人员ID = P.ID And P.ID=R.人员ID and U.用户名=[1] And (P.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or P.撤档时间 Is Null) and R.缺省=1"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Item.SubItems(3))
        If Err = 0 Then
            If rsTmp.BOF = False Then
                strSyss = strSyss & "," & rs("编号").Value
                
                strSQL = "SELECT 名称 FROM " & rs("所有者").Value & ".部门表 START WITH ID=[1] CONNECT BY PRIOR 上级id=ID"
                Set rs2 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, rsTmp("部门id").Value)
                If rs2.BOF = False Then
                    Do While Not rs2.EOF
                        strDepts = strDepts & ",'" & UCase(rs2("名称").Value) & "'"
                        rs2.MoveNext
                    Loop
                End If
            End If
        End If
        
        On Error GoTo 0
        
        rs.MoveNext
    Loop
        
    '3.找出本用户的提醒消息
    strSQL = "SELECT A.序号, A.系统, A.提醒条件, A.提醒内容, A.提醒报表, A.提醒声音, A.提醒窗口, A.提醒顺序, A.检查周期, A.提醒周期, A.开始时间, A.终止时间, B.程序ID AS 模块,B.名称 AS 报表名称,B.系统 As 报表系统 " & _
                "FROM zlNotices A," & _
                    "(SELECT Id , 编号, 名称, 说明, 密码, 打印机, 进纸, 票据, 打印方式, 系统, 程序id, 功能, 修改时间, 发布时间, 禁止开始时间, 禁止结束时间, 执行开始时间, 执行结束时间, 执行人员, 最后执行时间 FROM zlReports WHERE 发布时间 IS NOT NULL) B " & _
                "WHERE (A.序号 IN (" & _
                                    "SELECT 提醒序号 FROM zlNoticeUsr WHERE 提醒对象=0 " & _
                                    "Union " & _
                                    "SELECT 提醒序号 FROM zlNoticeUsr WHERE 提醒对象=1 AND UPPER(对象名称)=[1] " & _
                                    "Union " & _
                                    "SELECT 提醒序号 FROM zlNoticeUsr WHERE 提醒对象=2 AND UPPER(对象名称) IN (" & strDepts & ") " & _
                                    "Union " & _
                                    "SELECT 提醒序号 FROM zlNoticeUsr WHERE 提醒对象=3 AND UPPER(对象名称)=[2] " & _
                                ") " & _
                        "OR NOT EXISTS (SELECT 提醒序号 FROM zlNoticeUsr C WHERE C.提醒序号=A.序号))" & _
                    "AND B.编号(+) = A.提醒报表 " & _
                    "AND (A.系统 IN (" & strSyss & ") OR A.系统 IS NULL) " & _
                    "AND A.开始时间 <= SYSDATE And (A.终止时间 >= SYSDATE Or A.终止时间 Is Null) " & _
                    "AND " & IIf(StartUp = False, " A.检查周期 IS NOT NULL", " A.检查周期 IS NULL")
                    
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UCase(Item.SubItems(2)), UCase(Item.Text))
    If rs.BOF = False Then
    
        Item.Tag = ""
        
        Do While Not rs.EOF
            
            If CheckNotice(Item.SubItems(2), rs("序号").Value, str提醒内容) Then
                '要提醒,产生提醒信息
                                
                '--检查用户是否已经提醒-----------------------------------------------------------------------------------------------------------
                strSQL = "SELECT 提醒序号, 用户名, 检查时间, 检查结果, 提醒标志, 提醒时间, 已读标志, 提醒内容 FROM zlNoticeRec WHERE 提醒序号=[1] AND 用户名=[2] AND 已读标志<>1"
                
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, rs("序号").Value, Item.SubItems(2))
                If rsTmp.BOF = False Then

                    '提醒过,再检查是否新消息
                    If rsTmp("提醒标志").Value > 0 Then
                        Item.Tag = Item.Tag & "[INFOITEM-BEGIN']" & NVL(rsTmp("提醒内容").Value, 0) & "[''']" & NVL(rs("提醒声音").Value, 0) & "[''']" & NVL(rs("报表名称").Value, "") & "[''']" & NVL(rs("系统").Value, 0) & "[''']" & NVL(rs("模块").Value, 0) & "[''']" & winSock.LocalHostName & "[''']" & NVL(rs("提醒窗口").Value, 0) & "[''']" & NVL(rs("报表系统").Value, 0)
                    End If
                    
                End If
                '--检查用户是否已经提醒-----------------------------------------------------------------------------------------------------------
                
            
            End If
            
            rs.MoveNext
        Loop
        
        '发送提醒
        If Item.Tag <> "" Then
            If Left(Item.Tag, 17) = "[INFOITEM-BEGIN']" Then Item.Tag = Mid(Item.Tag, 18)
            Call SendMessage(Item.SubItems(4), Val(Item.SubItems(1)), Item.SubItems(2), Item.Tag)
        End If
    End If
End Function

Private Function CheckNotice(ByVal strUser As String, ByVal lngNo As Long, ByRef str提醒内容 As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:检查指定的提醒消息
    '------------------------------------------------------------------------------------------------------------------
    
    Dim strSQL As String
    Dim blnNotice As Boolean
    Dim blnCheck As Boolean
    Dim blnHaveResult As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim str提醒条件 As String
    Dim lngPos As Long
    Dim lngPosField As Long
    Dim strTmp As String
    Dim strTmpField As String
    Dim strField As String
    Dim strFieldType As String
    Dim strResult As String
    Dim lngLoop As Long

    
    strSQL = "SELECT A.序号, A.系统, A.提醒条件, A.提醒内容, A.提醒报表, A.提醒声音, A.提醒窗口, A.提醒顺序, A.检查周期, A.提醒周期, A.开始时间, A.终止时间,B.检查时间,B.提醒时间," & _
                    "DECODE(B.检查时间,NULL,NULL,DECODE(检查周期,NULL,NULL,SYSDATE - (B.检查时间+检查周期/(24*60)))) AS 检查期限," & _
                    "DECODE(B.提醒时间,NULL,NULL,DECODE(提醒周期,NULL,NULL,SYSDATE - (B.提醒时间+提醒周期/(24*60)))) AS 提醒期限 " & _
                    "FROM zlNotices A,(SELECT 提醒序号, 用户名, 检查时间, 检查结果, 提醒标志, 提醒时间, 已读标志, 提醒内容 FROM zlNoticeRec WHERE 用户名=[1])B " & _
                    "WHERE A.序号=B.提醒序号(+) AND A.序号=" & lngNo

    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strUser)
    If rs.BOF Then Exit Function
            
    blnCheck = False
        
    If IsNull(rs("检查时间").Value) Or IsNull(rs("检查周期").Value) Then
        blnCheck = True
    Else
        If IsNull(rs("检查期限").Value) = False Then
            If rs("检查期限").Value >= 0 Then blnCheck = True
        End If
    End If
    
    If blnCheck Then
    
        blnHaveResult = False
        str提醒内容 = ""
        
        If IsNull(rs("提醒条件").Value) = False Then
            
            str提醒条件 = UCase(rs("提醒条件").Value)
            
            '替换提醒条件中的固定变量[USER]
            
            str提醒条件 = ReplaceAll(str提醒条件, "[USER]", "[1]")
                
            '因为此时可能执行SQL不成功,没有权限,则不对本消息进行处理
            Err = 0
            On Error GoTo errHand
            
            Set rsTmp = zlDatabase.OpenSQLRecord(str提醒条件, Me.Caption, strUser)
            
            str提醒内容 = ""
            
            If Err = 0 Then
                If rsTmp.BOF = False Then
                    blnHaveResult = True
                    
                    'str提醒内容 = NVL(rsTmp("结果").Value)
                                        
                Else
                    '提醒条件不成立,删除记录
                    strSQL = "Zl_Zlnoticerec_Edit(2," & lngNo & ",'" & strUser & "',Null,Null,Null,Null,Null)"
                    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
                    
                    Exit Function
                End If
            Else
                Err = 0
                
                Exit Function
                
            End If
            
            
            '生成提醒内容,即替换引用的字段值
            
            str提醒内容 = NVL(rs("提醒内容").Value)
            
            'strTmp格式:如'[姓名];varchar2|[性别];date'
            strTmp = NVL(rs("提醒顺序").Value) & "|"
                           
            Do While strTmp <> "|" And strTmp <> ""
                
                lngPos = InStr(strTmp, "|")
                strTmpField = Mid(strTmp, 1, lngPos - 1)
                lngPosField = InStr(strTmpField, ";")
                
                If lngPosField > 0 Then
                
                    strField = Mid(strTmpField, 1, lngPosField - 1)
                    strFieldType = Trim(UCase(Mid(strTmpField, lngPosField + 1)))
                    
                    strTmp = Trim(Mid(strTmp, lngPos + 1))
                    
                    lngPos = InStr(str提醒内容, strField)
                    
                    If lngPos > 0 Then
                        
                        strResult = Trim(Mid(strField, 2))
                        strResult = Mid(strResult, 1, Len(strResult) - 1)
                        
                        On Error Resume Next
                        Err = 0
                        If strFieldType = "NUMBER" Then
                        
                            strResult = NVL(rsTmp(strResult).Value)
                            
'                            strResult = "to_char(" & strResult & ")"
                        ElseIf strFieldType = "DATE" Then
'                            strResult = "to_char(" & strResult & ",'yyyy-mm-dd')"
                            strResult = Format(NVL(rsTmp(strResult).Value), "yyyy-MM-dd")
                        Else
                            strResult = NVL(rsTmp(strResult).Value)
                        End If
                        
                        'str提醒内容 = Trim(Mid(str提醒内容, 1, lngPos - 1) & "'||" & strResult & "||'" & Mid(str提醒内容, lngPos + Len(strField)))
                        
                        If Err = 0 Then
                            str提醒内容 = Trim(Mid(str提醒内容, 1, lngPos - 1) & strResult & Mid(str提醒内容, lngPos + Len(strField)))
                        End If
                        
                        On Error GoTo errHand
                        
                    End If
                    
                End If
                
            Loop
            
'            lngPos = InStr(str提醒条件, " FROM ")
'
'            If lngPos > 0 Then
'                strSQL = Trim("SELECT '" & str提醒内容 & "' AS 结果 " & Mid(str提醒条件, lngPos))
'
'                '因为此时可能执行SQL不成功,没有权限,则不对本消息进行处理
'                Err = 0
'                On Error GoTo errHand
'
'                If rsTmp.State = adStateOpen Then rsTmp.Close
'                rsTmp.Open strSQL, gcnOracle
'
'
'                str提醒内容 = ""
'
'                If Err = 0 Then
'                    If rsTmp.BOF = False Then
'                        blnHaveResult = True
'                        str提醒内容 = NVL(rsTmp("结果").Value)
'                    Else
'                        '提醒条件不成立,删除记录
'                        strSQL = "DELETE FROM zlNoticeRec WHERE 提醒序号=" & lngNo & " AND upper(用户名)='" & UCase(strUser) & "'"
'                        gcnOracle.Execute strSQL
'
'                        Exit Function
'                    End If
'                Else
'                    Err = 0
'
'                    Exit Function
'
'                End If
'            End If
        Else
            blnHaveResult = True
            str提醒内容 = NVL(rs("提醒内容").Value)
        End If
                            
        strSQL = "Zl_Zlnoticerec_Edit(2," & lngNo & ",'" & strUser & "',Null,Null,Null,Null,Null)"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        
        strSQL = "Zl_Zlnoticerec_Edit(0," & lngNo & ",'" & UCase(strUser) & "',SYSDATE," & IIf(blnHaveResult, 1, 0) & ",0,0,'" & str提醒内容 & "')"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    End If
    
    '提醒检查
    blnNotice = False
        
    If IsNull(rs("检查周期").Value) Then
        blnNotice = True
    Else
        If IsNull(rs("提醒时间").Value) Then
            blnNotice = True
        Else
            If IsNull(rs("提醒期限").Value) = False Then
                If rs("提醒期限").Value >= 0 Then blnNotice = True
            End If
        End If
    End If
    
    If blnNotice Then
        strSQL = "Zl_Zlnoticerec_Edit(1," & lngNo & ",'" & strUser & "',Null,Null,1,Null,Null)"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    End If
            
    CheckNotice = blnNotice
    
    Exit Function
    
errHand:
    
End Function

Private Function GetClientIP(ByVal strWorkStation As String) As String
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    GetClientIP = strWorkStation
    
    strSQL = "Select IP From zlClients Where 工作站=[1] And IP Is Not Null"
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UCase(strWorkStation))
    If rs.BOF = False Then
        GetClientIP = rs("IP").Value
    End If
    
End Function
Private Function SendMessage(ByVal str工作站 As String, _
                            ByVal lng端口号 As Long, _
                            ByVal str用户名 As String, _
                            ByVal str消息串 As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:发送提醒消息
    '------------------------------------------------------------------------------------------------------------------
    
    Dim strTmp As String
    Dim varTmp As Variant
    Dim varTmp2 As Variant
    Dim lngCount As Long
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHand
    
    If Trim(str消息串) = "" Then Exit Function
'
'    '要发送的客户端机器名称(IP)和端口号
    winSock.RemoteHost = str工作站
    winSock.RemotePort = lng端口号
    
    '发送消息
    winSock.SendData str消息串
    
    SendMessage = True
    
    Exit Function
    
errHand:
    
End Function


Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

