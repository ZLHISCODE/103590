VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Begin VB.Form frmMedInOutClass 
   AutoRedraw      =   -1  'True
   Caption         =   "入出分类"
   ClientHeight    =   7170
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   11610
   FillColor       =   &H00FF0000&
   Icon            =   "frmMedInOutClass.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "frmMedInOutClass.frx":1CFA
   ScaleHeight     =   7170
   ScaleWidth      =   11610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Tag             =   "15"
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   6810
      Width           =   11610
      _ExtentX        =   20479
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMedInOutClass.frx":2004
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15399
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
   Begin MSComctlLib.ImageList ImgLvw单据Small 
      Left            =   4890
      Top             =   2010
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImgLvwBig 
      Left            =   4350
      Top             =   780
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImgLvwSmall 
      Left            =   4890
      Top             =   780
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImgTbarBlack 
      Left            =   6750
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImgTbarColor 
      Left            =   6180
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ListView Lvw所属单据 
      Height          =   1725
      Left            =   2040
      TabIndex        =   4
      Top             =   1980
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   3043
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "单据名称"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "说明"
         Object.Width           =   9701
      EndProperty
   End
   Begin MSComctlLib.ListView Lvw入出类别 
      Height          =   1185
      Left            =   2010
      TabIndex        =   3
      Top             =   750
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   2090
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "名称"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "编码"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "类别"
         Object.Width           =   2540
      EndProperty
   End
   Begin ComCtl3.CoolBar Cbar 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11610
      _ExtentX        =   20479
      _ExtentY        =   1164
      BandCount       =   1
      FixedOrder      =   -1  'True
      _CBWidth        =   11610
      _CBHeight       =   660
      _Version        =   "6.7.8988"
      Child1          =   "Tbar"
      MinHeight1      =   600
      Width1          =   4995
      Key1            =   "Common"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar Tbar 
         Height          =   600
         Left            =   30
         TabIndex        =   2
         Top             =   30
         Width           =   11490
         _ExtentX        =   20267
         _ExtentY        =   1058
         ButtonWidth     =   820
         ButtonHeight    =   1058
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   11
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预览"
               Key             =   "Preview"
               Object.ToolTipText     =   "预览"
               Object.Tag             =   "预览"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打印"
               Key             =   "Print"
               Object.ToolTipText     =   "打印"
               Object.Tag             =   "打印"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split1"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "增加"
               Key             =   "Add"
               Object.ToolTipText     =   "增加"
               Object.Tag             =   "增加"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "修改"
               Key             =   "Modify"
               Object.ToolTipText     =   "修改"
               Object.Tag             =   "修改"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "删除"
               Key             =   "Delete"
               Object.ToolTipText     =   "删除"
               Object.Tag             =   "删除"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split2"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "查看"
               Key             =   "View"
               Object.ToolTipText     =   "查看"
               Object.Tag             =   "查看"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Big"
                     Text            =   "大图标(&G)"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Small"
                     Text            =   "小图标(&M)"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "List"
                     Text            =   "列表(&L)"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Detail"
                     Text            =   "详细资料(&D)"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split3"
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "帮助"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Exit"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
            EndProperty
         EndProperty
         Begin VB.TextBox txtFind 
            Height          =   300
            Left            =   10080
            MaxLength       =   10
            TabIndex        =   7
            Tag             =   "简码"
            Top             =   210
            Width           =   1425
         End
         Begin VB.PictureBox picFind 
            BorderStyle     =   0  'None
            Height          =   300
            Left            =   9480
            ScaleHeight     =   285.714
            ScaleMode       =   0  'User
            ScaleWidth      =   495
            TabIndex        =   5
            Top             =   210
            Width           =   495
            Begin VB.Label lbl查找 
               Caption         =   "查找"
               Height          =   255
               Left            =   120
               TabIndex        =   6
               Top             =   74
               Width           =   495
            End
         End
      End
   End
   Begin VB.Image ImgUpDown_S 
      Height          =   45
      Left            =   2040
      MousePointer    =   7  'Size N S
      Top             =   1920
      Width           =   3345
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
      Begin VB.Menu mnuFile1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEditAdd 
         Caption         =   "增加(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "修改(&M)"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "删除(&D)"
         Shortcut        =   {DEL}
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "报表(&R)"
      Visible         =   0   'False
      Begin VB.Menu mnuReportItem 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "查看(&V)"
      Begin VB.Menu mnuViewTool 
         Caption         =   "工具栏(&T)"
         Begin VB.Menu mnuViewToolS 
            Caption         =   "标准按钮(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewToolT 
            Caption         =   "文本标签(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "状态栏(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuView1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "大图标(&G)"
         Checked         =   -1  'True
         Index           =   1
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "小图标(&M)"
         Checked         =   -1  'True
         Index           =   2
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "列表(&L)"
         Checked         =   -1  'True
         Index           =   3
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "详细资料(&D)"
         Checked         =   -1  'True
         Index           =   4
      End
      Begin VB.Menu mnuView2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewBill 
         Caption         =   "按单据显示药品入出类别(&B)"
      End
      Begin VB.Menu mnuView3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "刷新(&R)"
         Shortcut        =   {F5}
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
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "中联论坛(&F)"
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
Attribute VB_Name = "frmMedInOutClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private BlnStartUp As Boolean
Private RecClient As New ADODB.Recordset
Private strSQL As String                    '书写SQL语句
Private BlnEditReturn As Boolean            '修改成功与否
Private mlngMode As Long
Private mstrPrivs As String                              '权限串
Private Declare Function SetParent Lib "user32 " (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private mrsFind As ADODB.Recordset
Private mstrFindValue As String             '记录查询的值

Private Sub Form_Activate()
    If BlnStartUp = False Then
        Unload Me
        Exit Sub
    End If
    Form_Resize
End Sub

Private Sub Form_Load()
    BlnStartUp = False
    
    mlngMode = glngModul
    mstrPrivs = gstrPrivs
    '--恢复窗体及控件相关状态--
    RestoreWinState Me, App.ProductName
    Call zldatabase.ShowReportMenu(Me, glngSys, mlngMode, mstrPrivs)
    
    '--更新相关菜单--
    Lvw所属单据.View = lvwReport
    mnuViewToolT.Enabled = mnuViewToolS.Checked
    ClearViewState Lvw入出类别.View + 1
    
    If LoadInIcon = False Then Exit Sub
    Call LoadInLvw
    Call 权限控制
    
    BlnStartUp = True
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    SetParent txtFind.hwnd, Tbar.hwnd
    SetParent picFind.hwnd, Tbar.hwnd
    txtFind.Left = Me.Width - txtFind.Width
    picFind.Left = txtFind.Left - 100 - picFind.Width
    
    If Me.WindowState = 1 Then Exit Sub
    
    With Cbar
        .Bands(1).MinHeight = Tbar.Height
        Set .Bands(1).Child = Tbar
    End With
    
    With ImgUpDown_S
        .Left = 0
        .Top = (Me.ScaleHeight - stbThis.Height - Cbar.Top + Cbar.Height) * 0.5
        .Width = Me.ScaleWidth - .Left
    End With
    
    With Lvw入出类别
        .Left = 0
        .Top = IIF(Cbar.Visible, Cbar.Height, 0)
        .Width = ImgUpDown_S.Width
        .Height = ImgUpDown_S.Top - .Top
    End With
    
    With Lvw所属单据
        .Left = 0
        .Top = ImgUpDown_S.Top + ImgUpDown_S.Height
        .Width = ImgUpDown_S.Width
        .Height = Me.ScaleHeight - IIF(stbThis.Visible, stbThis.Height, 0) - .Top
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Function LoadInIcon() As Boolean
    '--为各控件装载图标--
    On Error Resume Next
    Err = 0
    LoadInIcon = False
    
    '--工具栏Tbar--
    With ImgTbarBlack
        .ImageHeight = 24
        .ImageWidth = 24
        .ListImages.Add 1, , LoadResPicture("BPREVIEW", vbResIcon)
        .ListImages.Add 2, , LoadResPicture("BPRINT", vbResIcon)
        .ListImages.Add 3, , LoadResPicture("BADD", vbResIcon)
        .ListImages.Add 4, , LoadResPicture("BMODIFY", vbResIcon)
        .ListImages.Add 5, , LoadResPicture("BDELETE", vbResIcon)
        .ListImages.Add 6, , LoadResPicture("BVIEW", vbResIcon)
        .ListImages.Add 7, , LoadResPicture("BHELP", vbResIcon)
        .ListImages.Add 8, , LoadResPicture("BEXIT", vbResIcon)
    End With
    With ImgTbarColor
        .ImageHeight = 24
        .ImageWidth = 24
        .ListImages.Add 1, , LoadResPicture("CPREVIEW", vbResIcon)
        .ListImages.Add 2, , LoadResPicture("CPRINT", vbResIcon)
        .ListImages.Add 3, , LoadResPicture("CADD", vbResIcon)
        .ListImages.Add 4, , LoadResPicture("CMODIFY", vbResIcon)
        .ListImages.Add 5, , LoadResPicture("CDELETE", vbResIcon)
        .ListImages.Add 6, , LoadResPicture("CVIEW", vbResIcon)
        .ListImages.Add 7, , LoadResPicture("CHELP", vbResIcon)
        .ListImages.Add 8, , LoadResPicture("CEXIT", vbResIcon)
    End With
    With Tbar
        Set .ImageList = ImgTbarBlack
        Set .HotImageList = ImgTbarColor
        .Buttons("Preview").Image = 1
        .Buttons("Print").Image = 2
        .Buttons("Add").Image = 3
        .Buttons("Modify").Image = 4
        .Buttons("Delete").Image = 5
        .Buttons("View").Image = 6
        .Buttons("Help").Image = 7
        .Buttons("Exit").Image = 8
    End With
    Cbar.Bands("Common").MinHeight = Tbar.Height
    
    '--列表Lvw入出类别--
    With ImgLvwBig
        .ImageHeight = 24
        .ImageWidth = 24
        .ListImages.Add , , LoadResPicture("CLASS", vbResIcon)
    End With
    With ImgLvwSmall
        .ImageHeight = 16
        .ImageWidth = 16
        .ListImages.Add , , LoadResPicture("CLASS", vbResIcon)
    End With
    With Lvw入出类别
        Set .SmallIcons = ImgLvwSmall
        Set .Icons = ImgLvwBig
    End With
    
    With ImgLvw单据Small
        .ImageHeight = 16
        .ImageWidth = 16
        .ListImages.Add , , LoadResPicture("BILL1", vbResIcon)
    End With
    With Lvw所属单据
        Set .SmallIcons = ImgLvw单据Small
    End With
    
    If Err <> 0 Then
        MsgBox "相关资源文件丢失，请与软件开发商联系！", vbInformation, gstrSysName
        Exit Function
    End If
    LoadInIcon = True
End Function

Private Sub ImgUpDown_S_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    
    With ImgUpDown_S
        If .Top + Y < 2000 Then Exit Sub
        If .Top + Y > Me.ScaleHeight - 2500 Then Exit Sub
        
        .Move .Left, .Top + Y
    End With
    
    With Lvw入出类别
        .Left = 0
        .Top = IIF(Cbar.Visible, Cbar.Height, 0)
        .Width = ImgUpDown_S.Width
        .Height = ImgUpDown_S.Top - .Top
    End With
    
    With Lvw所属单据
        .Left = 0
        .Top = ImgUpDown_S.Top + ImgUpDown_S.Height
        .Width = ImgUpDown_S.Width
        .Height = Me.ScaleHeight - IIF(stbThis.Visible, stbThis.Height, 0) - .Top
    End With
End Sub

Private Sub Lvw入出类别_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With Lvw入出类别
        .Sorted = False
        
        .SortKey = ColumnHeader.Index - 1
        .SortOrder = IIF(.SortOrder = lvwAscending, lvwDescending, lvwAscending)
        .Sorted = True
    End With
End Sub

Private Sub Lvw入出类别_DblClick()
    If Lvw入出类别.ListItems.Count = 0 Then Exit Sub
    
    If InStr(1, mstrPrivs, "增删改") <> 0 Then mnuEditModify_Click
End Sub

Private Sub Lvw入出类别_ItemClick(ByVal Item As MSComctlLib.ListItem)
    '--读出该入出类别属于哪些单据--
    
    On Error GoTo ErrHandle
    strSQL = "Select 编码,名称,说明 From 药品单据分类 Where 编码 In" & _
             " (Select 单据 From 药品单据性质 Where 类别ID=[1]) Order by 编码"
    Set RecClient = zldatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Mid(Item.Key, 3)))

    LoadInBill
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Lvw入出类别_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Lvw入出类别_DblClick
End Sub

Private Sub Lvw入出类别_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Lvw入出类别.ListItems.Count = 0 Then Exit Sub
    If Button <> 2 Then Exit Sub
    Dim ItemThis As ListItem
    
    On Error Resume Next
    Err = 0
    
    With Lvw入出类别
        Set ItemThis = .HitTest(X, Y)
        If Err <> 0 Then Exit Sub
        
        ItemThis.Selected = True
        .SelectedItem.Selected = True
    End With
    Lvw入出类别_ItemClick Lvw入出类别.SelectedItem
    If InStr(1, mstrPrivs, "增删改") <> 0 Then PopupMenu mnuEdit, 2
End Sub

Private Sub Lvw所属单据_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With Lvw所属单据
        .Sorted = False
        
        .SortKey = ColumnHeader.Index - 1
        .SortOrder = IIF(.SortOrder = lvwAscending, lvwDescending, lvwAscending)
        .Sorted = True
    End With
End Sub

Private Sub mnuEditAdd_Click()
    With frmEditInOutClass
        .EditState = 1
        .系数 = 1
        .Show 1, Me
    End With
    If BlnEditReturn Then mnuViewRefresh_Click
End Sub

Private Sub mnuEditDelete_Click()
    On Error GoTo ErrHand
    If MsgBox("你确认要删除该入出类别吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    gstrSQL = "zl_药品入出类别_delete (" & Mid(Lvw入出类别.SelectedItem.Key, 3) & ")"
    Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-删除药品入出类别")
    
    mnuViewRefresh_Click
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditModify_Click()
    With frmEditInOutClass
        .EditState = 2
        .类别ID = Mid(Lvw入出类别.SelectedItem.Key, 3)
        .名称 = Lvw入出类别.SelectedItem
        .编码 = Lvw入出类别.SelectedItem.SubItems(1)
        .系数 = IIF(Lvw入出类别.SelectedItem.SubItems(2) = "入库", 1, -1)
        .Show 1, Me
    End With
    If BlnEditReturn Then mnuViewRefresh_Click
End Sub

Private Sub mnuFileExcel_Click()
    subPrint 3
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub mnufilePreview_Click()
    subPrint 2
End Sub

Private Sub mnuFilePrint_Click()
    subPrint 1
End Sub

Private Sub mnuFilePrintset_Click()
    zlPrintSet
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me)
End Sub

Private Sub mnuViewBill_Click()
    With frmByBillShow
        .Show 1, Me
    End With
End Sub

Private Sub mnuViewRefresh_Click()
    LoadInLvw
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuhelpTitle_Click()
ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hwnd)
End Sub

Private Sub mnuViewIcon_Click(Index As Integer)
    ClearViewState Index
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = mnuViewStatus.Checked Xor True
    stbThis.Visible = mnuViewStatus.Checked
    
    Form_Resize
End Sub

Private Sub MnuViewToolS_Click()
    mnuViewToolS.Checked = mnuViewToolS.Checked Xor True
    Cbar.Visible = mnuViewToolS.Checked
    mnuViewToolT.Enabled = mnuViewToolS.Checked
    
    Form_Resize
End Sub

Private Sub MnuViewToolT_Click()
    mnuViewToolT.Checked = mnuViewToolT.Checked Xor True
    If mnuViewToolT.Checked Then
        With Tbar
            .Buttons("Preview").Caption = .Buttons("Preview").Tag
            .Buttons("Print").Caption = .Buttons("Print").Tag
            .Buttons("Add").Caption = .Buttons("Add").Tag
            .Buttons("Modify").Caption = .Buttons("Modify").Tag
            .Buttons("Delete").Caption = .Buttons("Delete").Tag
            .Buttons("View").Caption = .Buttons("View").Tag
            .Buttons("Help").Caption = .Buttons("Help").Tag
            .Buttons("Exit").Caption = .Buttons("Exit").Tag
        End With
    Else
        With Tbar
            .Buttons("Preview").Caption = ""
            .Buttons("Print").Caption = ""
            .Buttons("Add").Caption = ""
            .Buttons("Modify").Caption = ""
            .Buttons("Delete").Caption = ""
            .Buttons("View").Caption = ""
            .Buttons("Help").Caption = ""
            .Buttons("Exit").Caption = ""
        End With
    End If
    Cbar.Bands(1).MinHeight = Tbar.Height
    
    Form_Resize
End Sub

Private Sub Tbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "Preview"
        mnufilePreview_Click
    Case "Print"
        mnuFilePrint_Click
    Case "Add"
        mnuEditAdd_Click
    Case "Modify"
        mnuEditModify_Click
    Case "Delete"
        mnuEditDelete_Click
    Case "View"
        ClearViewState IIF(Lvw入出类别.View < lvwReport, Lvw入出类别.View + 2, 1)
    Case "Help"
        mnuhelpTitle_Click
    Case "Exit"
        mnuFileExit_Click
    End Select
End Sub

Private Sub Tbar_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    ClearViewState ButtonMenu.Index
End Sub

Private Sub Tbar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuViewTool, 2
End Sub

Private Function LoadInLvw() As Boolean
    '--装入ListView--
    Dim ItemThis As ListItem
    
    Lvw入出类别.ListItems.Clear
    stbThis.Panels(2) = ""
    
    On Error GoTo ErrHandle
    
'        If .State = 1 Then .Close
    strSQL = "Select ID,编码,名称,Decode(系数,1,'入库','出库') 系数 From 药品入出类别 Order by 系数"
        
'        Call SQLTest(App.Title, Me.Caption, strSQL)
    Set RecClient = zldatabase.OpenSQLRecord(strSQL, "LoadInLvw")
'        Call SQLTest
    With RecClient
        Do While Not .EOF
            Set ItemThis = Lvw入出类别.ListItems.Add(, "K_" & !ID, !名称, 1, 1)
            ItemThis.SubItems(1) = !编码
            ItemThis.SubItems(2) = !系数
            .MoveNext
        Loop
        
    End With
    
    With Lvw入出类别
        If .ListItems.Count <> 0 Then
            .ListItems(1).Selected = True
            .SelectedItem.Selected = True
            Lvw入出类别_ItemClick Lvw入出类别.SelectedItem
        End If
    End With
    
    SetMenuAndButton
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Function LoadInBill() As Boolean
    '--装入ListView--
    Dim ItemThis As ListItem
    
    Lvw所属单据.ListItems.Clear
    With RecClient
        If .EOF Then Exit Function
        
        Do While Not .EOF
            Set ItemThis = Lvw所属单据.ListItems.Add(, "K_" & !编码, !名称, , 1)
            ItemThis.SubItems(1) = !说明
            .MoveNext
        Loop
        
    End With
    
    With Lvw所属单据
        .ListItems(1).Selected = True
        .SelectedItem.Selected = True
    End With
End Function

Private Function ClearViewState(ByVal Index As Integer)
    '--设置显示状态--
    Dim intIndex As Integer
    For intIndex = 1 To 4
        mnuViewIcon(intIndex).Checked = False
    Next
    mnuViewIcon(Index).Checked = True
    
    Lvw入出类别.View = Index - 1
End Function

Public Function EditReturn(ByVal EditValue As Boolean)
    BlnEditReturn = EditValue
End Function
    
Public Function subPrint(ByVal bytMode As Byte)
    Dim objPrint As New zlPrintLvw
    
    objPrint.Title.Text = "药品入出类别"
    Set objPrint.Body.objData = Lvw入出类别
    objPrint.BelowAppItems.Add "打印人：" & gstrUserName
    objPrint.BelowAppItems.Add "打印时间：" & Format(zldatabase.Currentdate, "yyyy年MM月dd日")

    If bytMode = 1 Then
        Select Case zlPrintAsk(objPrint)
        Case 1
            zlPrintOrViewLvw objPrint, 1
        Case 2
            zlPrintOrViewLvw objPrint, 2
        Case 3
            zlPrintOrViewLvw objPrint, 3
        End Select
    Else
        zlPrintOrViewLvw objPrint, bytMode
    End If
End Function

Private Sub SetMenuAndButton()
    
    '设置按钮与工具栏
    mnuFilePrint.Enabled = (Lvw入出类别.ListItems.Count <> 0)
    mnuFilePreview.Enabled = mnuFilePrint.Enabled
    mnuFileExcel.Enabled = mnuFilePrint.Enabled
    Tbar.Buttons("Preview").Enabled = mnuFilePrint.Enabled
    Tbar.Buttons("Print").Enabled = mnuFilePrint.Enabled
    
    mnuEditModify.Enabled = (Lvw入出类别.ListItems.Count <> 0)
    mnuEditDelete.Enabled = (Lvw入出类别.ListItems.Count <> 0)
    Tbar.Buttons("Modify").Enabled = mnuEditModify.Enabled
    Tbar.Buttons("Delete").Enabled = mnuEditDelete.Enabled
End Sub

Private Sub 权限控制()
    If InStr(1, mstrPrivs, "增删改") = 0 Then
        mnuEdit.Visible = False
        Tbar.Buttons("Add").Visible = False
        Tbar.Buttons("Modify").Visible = False
        Tbar.Buttons("Delete").Visible = False
        Tbar.Buttons("Split1").Visible = False
    End If
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrHandle
    If KeyAscii = vbKeyReturn Then
        If txtFind.Text = "" Then Exit Sub
        If mstrFindValue <> txtFind.Text And txtFind.Text <> "" Then
            mstrFindValue = txtFind.Text
            Set mrsFind = Nothing
            
            gstrSQL = "select * from 药品入出类别 where 编码 like [1] or 名称 like [1]"
            Set mrsFind = zldatabase.OpenSQLRecord(gstrSQL, "药品入出类别查询", txtFind.Text & "%")
            Call LocateItem
        Else
            If Not mrsFind.EOF Then
                mrsFind.MoveNext
                Call LocateItem
            ElseIf mrsFind.RecordCount <> 0 And mrsFind.EOF Then
                mrsFind.MoveFirst
                Call LocateItem
            End If
        End If
    End If
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LocateItem()
    Dim strTemp As String
    
    txtFind.SetFocus
    txtFind.SelStart = 0
    txtFind.SelLength = Len(txtFind.Text)
    If mrsFind.RecordCount = 0 Then
        MsgBox " 没有找到符合条件的信息！", vbInformation, gstrSysName
        txtFind.SetFocus
        Exit Sub
    End If
    If mrsFind.EOF = True Then
        MsgBox " 已经定位完所有找到的信息，请重新输入条件！", vbInformation, gstrSysName
        txtFind.SetFocus
        Exit Sub
    End If
    
    Lvw入出类别.ListItems("K_" & mrsFind("ID")).Selected = True
    Lvw入出类别.SelectedItem.EnsureVisible
End Sub
