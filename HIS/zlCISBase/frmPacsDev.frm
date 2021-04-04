VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Begin VB.Form frmPacsDev 
   BackColor       =   &H00C0C0C0&
   Caption         =   "影像设备目录"
   ClientHeight    =   7365
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   8010
   Icon            =   "frmPacsDev.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7365
   ScaleWidth      =   8010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Visible         =   0   'False
   Begin ComCtl3.CoolBar clbThis 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8010
      _ExtentX        =   14129
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   8010
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tlbThis"
      MinWidth1       =   24000
      MinHeight1      =   720
      Width1          =   8730
      FixedBackground1=   0   'False
      Key1            =   "Comm"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tlbThis 
         Height          =   720
         Left            =   30
         TabIndex        =   4
         Top             =   30
         Width           =   24000
         _ExtentX        =   42333
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imgGray"
         HotImageList    =   "imgColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   9
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预览"
               Key             =   "Preview"
               Description     =   "预览"
               Object.ToolTipText     =   "预览当前表"
               Object.Tag             =   "预览"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打印"
               Key             =   "Print"
               Description     =   "打印"
               Object.ToolTipText     =   "打印当前表"
               Object.Tag             =   "打印"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split1"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "增加"
               Key             =   "Add"
               Description     =   "增加"
               Object.ToolTipText     =   "新文件"
               Object.Tag             =   "增加"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "修改"
               Key             =   "Mod"
               Description     =   "修改"
               Object.ToolTipText     =   "修改文件"
               Object.Tag             =   "修改"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "删除"
               Key             =   "Del"
               Description     =   "删除"
               Object.ToolTipText     =   "删除文件"
               Object.Tag             =   "删除"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split2"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Description     =   "帮助"
               Object.ToolTipText     =   "当前帮助主题"
               Object.Tag             =   "帮助"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Exit"
               Description     =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageIndex      =   11
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   6990
      Width           =   8010
      _ExtentX        =   14129
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPacsDev.frx":058A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9049
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
   Begin MSComctlLib.ImageList imgKind 
      Left            =   2220
      Top             =   6120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsDev.frx":0E1C
            Key             =   "Server"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsDev.frx":0F76
            Key             =   "Gate"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsDev.frx":1510
            Key             =   "Printer"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsDev.frx":1AAA
            Key             =   "影像设备"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picLine 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5895
      Left            =   2040
      MousePointer    =   9  'Size W E
      ScaleHeight     =   5895
      ScaleWidth      =   30
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   30
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   7080
      Top             =   525
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsDev.frx":1DC4
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsDev.frx":1FDE
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsDev.frx":21F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsDev.frx":2412
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsDev.frx":262C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsDev.frx":2846
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsDev.frx":2A60
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsDev.frx":2C7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsDev.frx":2E94
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsDev.frx":30AE
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsDev.frx":32CE
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   6315
      Top             =   435
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsDev.frx":34EE
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsDev.frx":370E
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsDev.frx":392E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsDev.frx":3B48
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsDev.frx":3D62
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsDev.frx":3F7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsDev.frx":4196
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsDev.frx":43B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsDev.frx":45CA
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsDev.frx":47E4
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsDev.frx":4A04
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwItem 
      Height          =   5385
      Left            =   2130
      TabIndex        =   0
      Top             =   930
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   9499
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgKind"
      SmallIcons      =   "imgKind"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFilePrintSet 
         Caption         =   "打印设置(&U)"
      End
      Begin VB.Menu mnuFilePreview 
         Caption         =   "预览(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "打印(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "输出到&Excel"
      End
      Begin VB.Menu mnuFileLine0 
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
      Begin VB.Menu mnuEditMod 
         Caption         =   "修改(&M)"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuEditDel 
         Caption         =   "删除(&D)"
         Shortcut        =   {DEL}
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "查看(&Q)"
      Begin VB.Menu mnuViewTools 
         Caption         =   "工具栏(&T)"
         Begin VB.Menu mnuViewToolsButton 
            Caption         =   "标准按钮(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewToolsText 
            Caption         =   "文本标签(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "状态栏(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewSpt1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "刷新(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuHelpHelp 
         Caption         =   "帮助主题(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "Web上的中联(&W)"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "中联主页(&H)"
         End
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "中联论坛(&F)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "发送反馈(&E)..."
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
Attribute VB_Name = "frmPacsDev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsTemp As New ADODB.Recordset
Dim objItem As ListItem
Dim intCount As Integer       '行列自由记数器

Private Sub Form_Load()
    '界面恢复
    Me.lvwItem.ColumnHeaders.Clear
    With Me.lvwItem.ColumnHeaders
        .Clear
        .Add , "_名称", "名称", 2500
        .Add , "_设备号", "设备号", 1000
        .Add , "_类型", "类型", 1200
        .Add , "_IP地址", "IP地址", 1500
        .Add , "_端口号", "端口号", 900
        .Add , "_Ftp目录", "Ftp目录", 3000
        .Add , "_用户名", "用户名", 1200
        .Add , "_Ftp本地路径", "Ftp本地路径", 2000
        .Add , "_本地AE", "本地AE", 2000
        .Add , "_设备AE", "设备AE", 2000
    End With
    With Me.lvwItem
        .SortKey = .ColumnHeaders("_设备号").Index - 1: .SortOrder = lvwAscending
    End With
    lvwItem.ListItems.Add , , , , 1
    lvwItem.ListItems.Clear
    
    Call RestoreWinState(Me, App.ProductName)
    Me.lvwItem.ColumnHeaders("_设备号").Position = 1
    
    '权限控制
    If InStr(1, gstrPrivs, "增删改") = 0 Then
        Me.mnuEdit.Enabled = False
        Me.mnuEditAdd.Enabled = False
        Me.mnuEditMod.Enabled = False
        Me.mnuEditDel.Enabled = False
        Me.tlbThis.Buttons("Add").Enabled = False
        Me.tlbThis.Buttons("Mod").Enabled = False
        Me.tlbThis.Buttons("Del").Enabled = False
    End If
    
    '装入数据
    Call zlRefItems
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    '-------------------------------------------------
    '根据窗体变化，调整各个部件的位置
    '-------------------------------------------------
    Dim lngHeightTools As Long, lngHeightState As Long
    lngHeightTools = IIf(Me.clbThis.Visible, Me.clbThis.Height, 0)
    lngHeightState = IIf(Me.stbThis.Visible, Me.stbThis.Height, 0)
    
    Me.picLine.Top = 0
    Me.picLine.Height = Me.ScaleHeight
    On Error Resume Next
    If Me.picLine.Left < 1000 Then Me.picLine.Left = 1000
    If Me.picLine.Left > Me.ScaleWidth - 2600 Then Me.picLine.Left = Me.ScaleWidth - 2600
    
    With Me.lvwItem
        .Left = Me.ScaleLeft
        .Width = Me.ScaleWidth - .Left
        .Top = Me.ScaleTop + lngHeightTools
        .Height = Me.ScaleHeight - .Top - lngHeightState
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub lvwItem_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With Me.lvwItem
        .SortKey = ColumnHeader.Index - 1
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
        Else
            .SortOrder = lvwAscending
        End If
    End With
End Sub

Private Sub lvwItem_DblClick()
    If Me.mnuEditMod.Enabled Then Call mnuEditMod_Click
End Sub

Private Sub lvwItem_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Me.mnuEditMod.Enabled Then Call mnuEditMod_Click
End Sub

Private Sub lvwItem_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And Me.mnuEdit.Enabled Then PopupMenu Me.mnuEdit, 2
End Sub

Private Sub mnuEditAdd_Click()
    If frmPACSDevEdit.ShowMe(Me, "") Then zlRefItems
End Sub

Private Sub mnuEditDel_Click()
    If Me.lvwItem.SelectedItem Is Nothing Then Exit Sub
    If MsgBox("真的将“" & Me.lvwItem.SelectedItem.Text & "”从影像设备目录中删除吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    gstrSql = "zl_影像设备目录_Delete('" & Mid(Me.lvwItem.SelectedItem.Key, 2) & "')"
    Err = 0: On Error GoTo ErrHand
    Call SQLTest(App.ProductName, Me.Caption, gstrSql): gcnOracle.Execute gstrSql, , adCmdStoredProc: Call SQLTest
    Call Me.lvwItem.ListItems.Remove(Me.lvwItem.SelectedItem.Key)
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume Next
    Call SaveErrLog
End Sub

Private Sub mnuEditMod_Click()
    If Me.lvwItem.SelectedItem Is Nothing Then Exit Sub
    If frmPACSDevEdit.ShowMe(Me, Mid(Me.lvwItem.SelectedItem.Key, 2)) Then zlRefItems
End Sub

Private Sub mnuFileExcel_Click()
    Call RptPrint(3)
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFilePreview_Click()
    Call RptPrint(0)
End Sub

Private Sub mnuFilePrint_Click()
    Call RptPrint(1)
End Sub

Private Sub mnuFilePrintSet_Click()
    Call zlPrintSet
End Sub

Private Sub mnuhelpAbout_Click()
    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
End Sub

Private Sub mnuHelpHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hWnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hWnd)
End Sub

Private Sub mnuViewRefresh_Click()
    Call zlRefItems
End Sub

Private Sub mnuViewStatus_Click()
    Me.mnuViewStatus.Checked = Not Me.mnuViewStatus.Checked
    Me.stbThis.Visible = Me.mnuViewStatus.Checked
    Call Form_Resize
End Sub

Private Sub mnuViewToolsButton_Click()
    Me.mnuViewToolsButton.Checked = Not Me.mnuViewToolsButton.Checked
    Me.clbThis.Visible = Me.mnuViewToolsButton.Checked
    Form_Resize
End Sub

Private Sub mnuViewToolsText_Click()
    Dim i As Integer
    Me.mnuViewToolsText.Checked = Not Me.mnuViewToolsText.Checked
    If Me.mnuViewToolsText.Checked Then
        For i = 1 To Me.tlbThis.Buttons.Count
            Me.tlbThis.Buttons(i).Caption = Me.tlbThis.Buttons(i).Tag
        Next
    Else
        For i = 1 To Me.tlbThis.Buttons.Count
            Me.tlbThis.Buttons(i).Caption = ""
        Next
    End If
    Me.clbThis.Bands(1).MinHeight = Me.tlbThis.Height
    Me.clbThis.Refresh
    Form_Resize
End Sub

Private Sub picLine_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        Me.picLine.Left = Me.picLine.Left + x
    End If
End Sub

Private Sub picLine_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call Form_Resize
End Sub

Private Sub tlbThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case UCase(Button.Key)
    Case UCase("Preview")
        Call mnuFilePreview_Click
    Case UCase("Print")
        Call mnuFilePrint_Click
    Case UCase("Add")
        Call mnuEditAdd_Click
    Case UCase("Mod")
        Call mnuEditMod_Click
    Case UCase("Del")
        Call mnuEditDel_Click
    Case UCase("Help")
        Call mnuHelpHelp_Click
    Case UCase("Exit")
        Call mnuFileExit_Click
    End Select
End Sub

Private Sub RptPrint(ByVal bytMode As Byte)
    '-------------------------------------------------
    '功能:记录表打印
    '参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    '-------------------------------------------------
    Dim objPrint As New zlPrintLvw
    Dim bytR As Byte
    On Error Resume Next
    
    Set objPrint.Body.objData = Me.lvwItem
    objPrint.Title.Text = "影像设备目录"
    objPrint.UnderAppItems.Add ""
    objPrint.BelowAppItems.Add "打印时间：" & Now
    
    If bytMode = 1 Then
        bytR = zlPrintAsk(objPrint)
        If bytR <> 0 Then zlPrintOrViewLvw objPrint, bytR
    Else
        zlPrintOrViewLvw objPrint, bytMode
    End If
End Sub

Public Sub zlRefItems()
    '-------------------------------------------------
    '功能:刷新当前的项目列表
    '-------------------------------------------------
    Dim strCurrKey As String
    If Not lvwItem.SelectedItem Is Nothing Then strCurrKey = lvwItem.SelectedItem.Key
    gstrSql = "Select 设备号,设备名,Decode(Nvl(类型,1),1,'存储',2,'影像接收',3,'胶片打印',4,'影像设备') As 设备类型," & _
        "Nvl(类型,1) As 类型,IP地址,端口号,Ftp目录,用户名,密码,本地AE,设备AE,本机目录 From 影像设备目录"
    
    Err = 0: On Error GoTo ErrHand
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.ProductName, Me.Caption, gstrSql)
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "zlRefItems")
'        Call SQLTest
    Me.lvwItem.ListItems.Clear
    With rsTemp
        Do While Not .EOF
            Set objItem = Me.lvwItem.ListItems.Add(, "_" & !设备号, !设备名, Val(!类型), Val(!类型))
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_设备号").Index - 1) = !设备号
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_类型").Index - 1) = !设备类型
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_IP地址").Index - 1) = Nvl(!IP地址)
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_端口号").Index - 1) = Nvl(!端口号)
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_Ftp目录").Index - 1) = Nvl(!ftp目录)
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_用户名").Index - 1) = Nvl(!用户名)
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_Ftp本地路径").Index - 1) = Nvl(!本机目录)
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_本地AE").Index - 1) = Nvl(!本地AE)
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_设备AE").Index - 1) = Nvl(!设备AE)
            objItem.Tag = Nvl(!密码)
            .MoveNext
        Loop
    End With
    If Me.lvwItem.ListItems.Count > 0 Then
        Err = 0: On Error Resume Next
        lvwItem.ListItems(strCurrKey).Selected = True
        Me.lvwItem.SelectedItem.EnsureVisible
    End If
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub
