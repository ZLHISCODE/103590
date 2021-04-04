VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmExecute 
   BackColor       =   &H8000000A&
   ClientHeight    =   4455
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7290
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4455
   ScaleWidth      =   7290
   ShowInTaskbar   =   0   'False
   Tag             =   "Execute"
   Begin MSComctlLib.ImageList ImgLvwSmall 
      Left            =   3360
      Top             =   1950
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.PictureBox PicBackGroud 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000009&
      Height          =   4455
      Left            =   0
      ScaleHeight     =   4395
      ScaleWidth      =   7215
      TabIndex        =   0
      Top             =   0
      Width           =   7275
      Begin MSComctlLib.ImageList ImgLvw 
         Left            =   2130
         Top             =   540
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   393216
      End
      Begin MSComctlLib.ListView LvwList 
         Height          =   4395
         Left            =   2760
         TabIndex        =   1
         Top             =   0
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   7752
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483639
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "标题"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "说明"
            Object.Width           =   10583
         EndProperty
      End
      Begin VB.Label Lbl说明 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "#说明#"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   60
         TabIndex        =   4
         Top             =   1590
         Width           =   540
      End
      Begin VB.Label Lbl标题 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "#标题#"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   60
         TabIndex        =   3
         Top             =   1320
         Width           =   600
      End
      Begin VB.Image ImgLine 
         Height          =   45
         Left            =   30
         Picture         =   "FrmExecute.frx":0000
         Top             =   1140
         Width           =   2760
      End
      Begin VB.Image ImgIcon 
         Height          =   435
         Left            =   210
         Top             =   180
         Width           =   495
      End
      Begin VB.Label LblCaption 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "#控制面板#"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   30
         TabIndex        =   2
         Top             =   810
         Width           =   1590
      End
      Begin VB.Image ImgBackGroud 
         Height          =   705
         Left            =   60
         Picture         =   "FrmExecute.frx":06BA
         Top             =   30
         Width           =   1755
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFilePrintset 
         Caption         =   "打印设置(&S)"
      End
      Begin VB.Menu mnuFilePreview 
         Caption         =   "打印预览(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "打印(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "输出到&Excel"
      End
      Begin VB.Menu mnuFile2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "查看(&V)"
      Begin VB.Menu mnuViewMode 
         Caption         =   "大图标(&G)"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuViewMode 
         Caption         =   "小图标(&M)"
         Checked         =   -1  'True
         Index           =   1
      End
      Begin VB.Menu mnuViewMode 
         Caption         =   "列表(&L)"
         Checked         =   -1  'True
         Index           =   2
      End
      Begin VB.Menu mnuViewMode 
         Caption         =   "详细资料(&D)"
         Checked         =   -1  'True
         Index           =   3
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuHelpTitle 
         Caption         =   "帮助主题(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "WEB上的中联(&W)"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "中联主页(&H)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "发送反馈(&K)..."
         End
      End
      Begin VB.Menu mnuHelp1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)..."
      End
   End
End
Attribute VB_Name = "FrmExecute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnShowMe As Boolean                               '是否由最小化显示
Private mblnStartUp As Boolean                              '启动成功
Private mstrCode As String                                  '菜单编号
Private mstrCaption As String                               '窗体名称
Public mrsMenus As New ADODB.Recordset

Public Property Get Str编号() As String
    Str编号 = mstrCode
End Property

Public Property Let Str编号(ByVal vNewValue As String)
    mstrCode = vNewValue
End Property

Private Sub LvwList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With LvwList
        .Sorted = False
        .SortKey = ColumnHeader.Index - 1
        .SortOrder = IIf(.SortOrder = 0, 1, 0)
        .Sorted = True
    End With
End Sub

Private Sub LvwList_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    LvwList.Drag 0
End Sub

Private Sub LvwList_GotFocus()
    mnuViewMode_Click LvwList.View
End Sub

Private Sub LvwList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim LvwThis As ListItem, IntLen As Integer
    Dim Str说明 As String
    
    With LvwList
        If .ListItems.Count = 0 Then Exit Sub
        If .SelectedItem Is Nothing Then Exit Sub
        Set LvwThis = .SelectedItem
    End With
    
    '--填充数据至左上角说明文字--
    Lbl标题 = LvwThis                                   '应该不会超过10个汉字
    Str说明 = LvwThis.SubItems(1)                       '超过则添加回车符
    Lbl说明 = ""
    
    If Len(Str说明) > 10 Then
        For IntLen = 1 To (Len(Str说明) \ 14) + 1
            Lbl说明 = Lbl说明 & IIf(IntLen = 1, Space(4), vbCrLf) & Mid(Str说明, IIf(IntLen = 1, 1, IntLen * 14 - 15), IIf(IntLen = 1, 12, 14))
        Next
    Else
        Lbl说明 = Str说明
    End If
End Sub

Private Sub LvwList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    With LvwList
        If .ListItems.Count = 0 Then Exit Sub
    End With
    If Button = 2 Then PopupMenu mnuView, 2
End Sub

Private Sub LvwList_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then LvwList.Drag 1
End Sub

Private Sub mnuFileExcel_Click()
    SubPrint 3
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFilePreview_Click()
    SubPrint 2
End Sub

Private Sub mnuFilePrint_Click()
    SubPrint 1
End Sub

Private Sub mnuFilePrintset_Click()
    zlPrintSet
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuhelpTitle_Click()
    'Shell "hh.exe " & App.Path & "\zlBaseCode.chm::/基础数据管理/部门部门管理.htm", vbNormalFocus
End Sub

Public Property Let 窗体名称(ByVal vNewValue As String)
    mstrCaption = vNewValue
End Property

Private Sub Form_Activate()
    If mblnStartUp = False Then
        Unload Me
        Exit Sub
    End If
    On Error Resume Next
    LvwList.SetFocus
End Sub

Private Sub Form_Deactivate()
    FrmWin.Find任务 -99999999
End Sub

Private Sub Form_Load()
    Dim LngIcon As Long, LngModul As Long
    
    mblnShowMe = False
    mblnStartUp = False
    
    LngIcon = 0
    If mstrCode <> "9003" Then
        With mrsMenus
            .MoveFirst
            .Find "编号='" & mstrCode & "'"
            LngIcon = !图标
            LngModul = !模块
        End With
    End If
    
    Icon = FrmWin.GetPicDisp(LngIcon, LngModul <> 0)
    ImgIcon = FrmWin.GetPicDisp(LngIcon, LngModul <> 0)
    Caption = mstrCaption
    LblCaption = mstrCaption
    
    If LoadLvw = False Then Exit Sub
    
    mblnStartUp = True
    RestoreWinState Me, , Me.Caption
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = 1 Then
        FrmWin.Find任务 -99999999
        Me.Hide
    End If
    
    With PicBackGroud
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight
    End With
    
    With LvwList
        .Width = Me.ScaleWidth - .Left - 50
        .Height = Me.ScaleHeight - 50
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, , Me.Caption
End Sub

Private Sub LvwList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then LvwList_DblClick
End Sub

Private Sub LvwList_DblClick()
    Dim LngFindWindows As Long
    
    If LvwList.ListItems.Count = 0 Then Exit Sub
    If LvwList.SelectedItem.Tag = -1 Then
        
        If mstrCode <> "9003" Then
            '执行该模块
            With mrsMenus
                .MoveFirst
                .Find "编号=" & Mid(LvwList.SelectedItem.Key, 3)
                
                Call ExecuteFunc(.Fields("系统").Value, IIf(IsNull(.Fields("部件").Value), "", .Fields("部件").Value), .Fields("模块").Value)
            End With
        Else
            Select Case Mid(LvwList.SelectedItem.Key, 3)
            Case "9100"
                FrmWin.mclsAppTool.CodeMan 0, 1, gcnOracle, FrmWin, gstrDbUser
            Case "9101"
                FrmWin.mclsAppTool.CodeMan 0, 2, gcnOracle, FrmWin, gstrDbUser
            Case "9102"
                FrmWin.mnuRepairComponent_Click
            Case "9103"
                FrmWin.mnuToolStyle_Click
            Case "9104"
                Dim ObjExcel As Object, strHaveSys As String
                
                strHaveSys = gobjRelogin.Systems
                On Error Resume Next
                Err = 0
                Set ObjExcel = CreateObject("Zl9Excel.ClsExcel")
                If Err <> 0 Then
                    MsgBox "无法创建EXCEL部件，您将不能使用EXCEL报表！", vbInformation, gstrSysName
                    Exit Sub
                End If
                Call ObjExcel.CodeMan(0, 0, gcnOracle, Me, gstrDbUser)
                Call ObjExcel.SetHaveSys(strHaveSys)
                Call ObjExcel.ExcelReportMain
                Set ObjExcel = Nothing
            End Select
        End If
    Else
        '打开该模块
        FrmWin.OpenWindow Mid(LvwList.SelectedItem.Key, 3), LvwList.SelectedItem.Text
    End If
End Sub

Public Property Let ShowMe(ByVal vNewValue As Boolean)
    mblnShowMe = vNewValue
End Property

Private Function LoadLvw() As Boolean
    LoadLvw = False
    
    If mstrCode = "9003" Then
        With ImgLvw
            .ImageHeight = 32
            .ImageWidth = 32
            .ListImages.Add .ListImages.Count + 1, "K_" & .ListImages.Count + 1, FrmWin.GetPicDisp(-5)
            .ListImages.Add .ListImages.Count + 1, "K_" & .ListImages.Count + 1, FrmWin.GetPicDisp(-5)
            .ListImages.Add .ListImages.Count + 1, "K_" & .ListImages.Count + 1, FrmWin.GetPicDisp(-5)
            .ListImages.Add .ListImages.Count + 1, "K_" & .ListImages.Count + 1, FrmWin.GetPicDisp(-5)
            .ListImages.Add .ListImages.Count + 1, "K_" & .ListImages.Count + 1, FrmWin.GetPicDisp(-5)
        End With
        With ImgLvwSmall
            .ImageHeight = 16
            .ImageWidth = 16
            .ListImages.Add .ListImages.Count + 1, "K_" & .ListImages.Count + 1, FrmWin.GetPicDisp(-5)
            .ListImages.Add .ListImages.Count + 1, "K_" & .ListImages.Count + 1, FrmWin.GetPicDisp(-5)
            .ListImages.Add .ListImages.Count + 1, "K_" & .ListImages.Count + 1, FrmWin.GetPicDisp(-5)
            .ListImages.Add .ListImages.Count + 1, "K_" & .ListImages.Count + 1, FrmWin.GetPicDisp(-5)
            .ListImages.Add .ListImages.Count + 1, "K_" & .ListImages.Count + 1, FrmWin.GetPicDisp(-5)
        End With
        Set LvwList.Icons = ImgLvw
        Set LvwList.SmallIcons = ImgLvwSmall
        
        With LvwList
            If FrmWin.mnuToolDictonary.Visible Then
                .ListItems.Add , "K_9100", "字典管理工具", 1, 1
                .ListItems("K_9100").SubItems(1) = "对本系统的基础数据进行管理"
                .ListItems("K_9100").Tag = -1
            End If
            If FrmWin.mnuToolMessage.Visible Then
                .ListItems.Add , "K_9101", "消息收发管理", 2, 2
                .ListItems("K_9101").SubItems(1) = "对本系统内的消息进行收发管理"
                .ListItems("K_9101").Tag = -1
            End If
            
            If FrmWin.mnuToolMessage.Visible Then
                .ListItems.Add , "K_9105", "提醒消息查阅", 2, 2
                .ListItems("K_9105").SubItems(1) = "对本用户的提醒消息进行查阅"
                .ListItems("K_9105").Tag = -1
            End If
            
            If FrmWin.mnuToolStyle.Visible Then
                .ListItems.Add , "K_9103", "系统选项", 3, 3
                .ListItems("K_9103").SubItems(1) = "根据自己的喜好选择导航台,使工作更顺心"
                .ListItems("K_9103").Tag = -1
            End If
            If FrmWin.mnuToolExcel.Visible Then
                .ListItems.Add , "K_9104", "启动EXCEL报表", 4, 4
                .ListItems("K_9104").SubItems(1) = "随心所欲定义自己所需要的EXCEL报表"
                .ListItems("K_9104").Tag = -1
            End If
            .ListItems.Add , "K_9102", "检测安装部件", 5, 5
            .ListItems("K_9102").SubItems(1) = "检测本机安装的部件是否发生变动"
            .ListItems("K_9102").Tag = -1
        End With
    Else
        With mrsMenus
            .Filter = "上级='" & mstrCode & "'"
            LvwList.ListItems.Clear
            If .EOF Then Exit Function
            
            On Error Resume Next
            With ImgLvw
                .ImageHeight = 32
                .ImageWidth = 32
            End With
            Do While Not .EOF '为ImageList装入图标
                ImgLvw.ListImages.Add ImgLvw.ListImages.Count + 1, "K_" & ImgLvw.ListImages.Count + 1, FrmWin.GetPicDisp(!图标, !模块 <> 0)
                .MoveNext
            Loop
            
            .MoveFirst
            With ImgLvwSmall
                .ImageHeight = 16
                .ImageWidth = 16
            End With
            Do While Not .EOF '为ImageList装入图标
                ImgLvwSmall.ListImages.Add ImgLvwSmall.ListImages.Count + 1, "K_" & ImgLvwSmall.ListImages.Count + 1, FrmWin.GetPicDisp(!图标, !模块 <> 0)
                .MoveNext
            Loop
            
            Set LvwList.Icons = ImgLvw
            Set LvwList.SmallIcons = ImgLvwSmall
            
            .MoveFirst
            Do While Not .EOF
                LvwList.ListItems.Add , "K_" & !编号, !标题, .AbsolutePosition, .AbsolutePosition
                LvwList.ListItems("K_" & !编号).SubItems(1) = IIf(IsNull(!说明), "", !说明)
                LvwList.ListItems("K_" & !编号).Tag = IIf(!模块 = 0, 0, -1)
                .MoveNext
            Loop
        End With
    End If
    
    If LvwList.ListItems.Count <> 0 Then
        LvwList.ListItems(1).Selected = True
        LvwList.SelectedItem.Selected = True
        LvwList_ItemClick LvwList.SelectedItem
    End If
    LoadLvw = True
End Function

Private Function SubPrint(ByVal BytMode As Byte)
    Dim objPrint As New zlPrintLvw
    
    objPrint.Title.Text = Caption
    Set objPrint.Body.objData = LvwList
    objPrint.BelowAppItems.Add "打印人：" & gstrUserName
    objPrint.BelowAppItems.Add "打印时间：" & Format(zlDatabase.Currentdate, "yyyy年MM月dd日")

    If BytMode = 1 Then
        Select Case zlPrintAsk(objPrint)
        Case 1
            zlPrintOrViewLvw objPrint, 1
        Case 2
            zlPrintOrViewLvw objPrint, 2
        Case 3
            zlPrintOrViewLvw objPrint, 3
        End Select
    Else
        zlPrintOrViewLvw objPrint, BytMode
    End If
End Function

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hwnd)
End Sub

Private Sub mnuViewMode_Click(Index As Integer)
    Dim IntCheck As Integer
    
    LvwList.View = Index
    For IntCheck = 0 To 3
        If IntCheck <> Index Then
            mnuViewMode(IntCheck).Checked = False
        Else
            mnuViewMode(IntCheck).Checked = True
        End If
    Next
End Sub
