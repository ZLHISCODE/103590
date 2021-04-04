VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Begin VB.Form frmRadLists 
   BackColor       =   &H00C0C0C0&
   Caption         =   "影像检查项目"
   ClientHeight    =   7365
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   8010
   Icon            =   "frmRadLists.frx":0000
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
      TabIndex        =   4
      Top             =   0
      Width           =   8010
      _ExtentX        =   14129
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   8010
      _CBHeight       =   780
      _Version        =   "6.7.8988"
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
         TabIndex        =   5
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
      TabIndex        =   3
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
            Picture         =   "frmRadLists.frx":08CA
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
            Enabled         =   0   'False
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
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadLists.frx":115C
            Key             =   "kind"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadLists.frx":16F6
            Key             =   "item"
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
      TabIndex        =   2
      Top             =   960
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
            Picture         =   "frmRadLists.frx":1C90
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadLists.frx":1EAA
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadLists.frx":20C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadLists.frx":22DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadLists.frx":24F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadLists.frx":2712
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadLists.frx":292C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadLists.frx":2B46
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadLists.frx":2D60
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadLists.frx":2F7A
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadLists.frx":319A
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
            Picture         =   "frmRadLists.frx":33BA
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadLists.frx":35DA
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadLists.frx":37FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadLists.frx":3A14
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadLists.frx":3C2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadLists.frx":3E48
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadLists.frx":4062
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadLists.frx":427C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadLists.frx":4496
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadLists.frx":46B0
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadLists.frx":48D0
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwKind 
      Height          =   5625
      Left            =   15
      TabIndex        =   0
      Top             =   945
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   9922
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "imgKind"
      SmallIcons      =   "imgKind"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "名称"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "简称"
         Object.Width           =   1411
      EndProperty
   End
   Begin MSComctlLib.ListView lvwItem 
      Height          =   5385
      Left            =   2130
      TabIndex        =   1
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
Attribute VB_Name = "frmRadLists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsTemp As New ADODB.Recordset
Dim objItem As ListItem
Dim intCount As Integer       '行列自由记数器

Private Sub Form_Activate()
    If Me.lvwKind.ListItems.count = 0 Then
        MsgBox "影像检查类别数据丢失！(联系管理员)", vbInformation, gstrSysName
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    '界面恢复
    Me.lvwItem.ColumnHeaders.Clear
    With Me.lvwItem.ColumnHeaders
        .Clear
        .Add , "_名称", "名称", 2500
        .Add , "_编码", "编码", 1000
        .Add , "_部位", "部位", 900
        .Add , "_单位", "单位", 600
        .Add , "_可行病检", "可行病检", 1000
        .Add , "_可发胶片", "可发胶片", 1000
        .Add , "_报告图象", "报告图象", 900
        .Add , "_检查准备", "检查准备", 2000
    End With
    With Me.lvwItem
        .SortKey = .ColumnHeaders("_编码").Index - 1: .SortOrder = lvwAscending
    End With
    
    Call RestoreWinState(Me, App.ProductName)
    Me.lvwKind.View = lvwReport
    Me.lvwItem.ColumnHeaders("_编码").Position = 1
    
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
    gstrSql = "Select * From 影像检查类别 Order By 排列"
    Err = 0: On Error GoTo ErrHand
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.ProductName, Me.Caption, gstrSql)
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "Form_Load")
'        Call SQLTest
    With rsTemp
        Me.lvwKind.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwKind.ListItems.Add(, "_" & !编码, !名称, "kind", "kind")
            objItem.SubItems(1) = !编码
            .MoveNext
        Loop
    End With
    Err = 0: On Error GoTo 0
    If Me.lvwKind.ListItems.count > 0 Then
        Me.lvwKind.ListItems(1).Selected = True
        Me.lvwKind.SelectedItem.EnsureVisible
        Call zlRefItems
    End If
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
    
    With Me.lvwKind
        .Left = Me.ScaleLeft
        .Width = Me.picLine.Left - .Left
        .Top = Me.ScaleTop + lngHeightTools
        .Height = Me.ScaleHeight - .Top - lngHeightState
    End With
    
    With Me.lvwItem
        .Left = Me.picLine.Left + Me.picLine.Width
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

Private Sub lvwKind_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Call zlRefItems
End Sub

Private Sub mnuEditAdd_Click()
    frmRadNew.Show 1, Me
End Sub

Private Sub mnuEditDel_Click()
    If Me.lvwItem.SelectedItem Is Nothing Then Exit Sub
    If MsgBox("真的将“" & Me.lvwItem.SelectedItem.Text & "”从影像检查项目中删除吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    gstrSql = "zl_影像检查项目_Delete(" & Mid(Me.lvwItem.SelectedItem.Key, 2) & ")"
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
    With frmRadMod
        .lblBaseInfo.Tag = Mid(Me.lvwItem.SelectedItem.Key, 2)
        .Show 1, Me
    End With
    Call zlRefItems(Mid(Me.lvwItem.SelectedItem.Key, 2))
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
    If Me.lvwItem.SelectedItem Is Nothing Then
        Call zlRefItems
    Else
        Call zlRefItems(Mid(Me.lvwItem.SelectedItem.Key, 2))
    End If
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
        For i = 1 To Me.tlbThis.Buttons.count
            Me.tlbThis.Buttons(i).Caption = Me.tlbThis.Buttons(i).Tag
        Next
    Else
        For i = 1 To Me.tlbThis.Buttons.count
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
    objPrint.Title.Text = Me.lvwKind.SelectedItem.Text & "检查项目"
    objPrint.UnderAppItems.Add ""
    objPrint.BelowAppItems.Add "打印时间：" & Now
    
    If bytMode = 1 Then
        bytR = zlPrintAsk(objPrint)
        If bytR <> 0 Then zlPrintOrViewLvw objPrint, bytR
    Else
        zlPrintOrViewLvw objPrint, bytMode
    End If
End Sub

Public Sub zlRefItems(Optional lngItemId As Long)
    '-------------------------------------------------
    '功能:刷新当前的项目列表
    '-------------------------------------------------
    If Me.lvwKind.SelectedItem Is Nothing Then Exit Sub
    
    gstrSql = "Select I.ID,I.编码, I.名称,I.标本部位, I.计算单位,R.可行病检,R.可发胶片,R.报告图象,R.检查准备" & _
            "  From 诊疗项目目录 I, 影像检查项目 R" & _
            " Where I.ID = R.诊疗项目id And R.影像类别=[1] "
    
    Err = 0: On Error GoTo ErrHand
    
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, Mid(Me.lvwKind.SelectedItem.Key, 2))
        
    With rsTemp
        Me.lvwItem.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwItem.ListItems.Add(, "_" & !ID, !名称, "item", "item")
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_编码").Index - 1) = !编码
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_部位").Index - 1) = IIf(IsNull(!标本部位), "", !标本部位)
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_单位").Index - 1) = IIf(IsNull(!计算单位), "", !计算单位)
            Select Case !可行病检
            Case 1
                objItem.SubItems(Me.lvwItem.ColumnHeaders("_可行病检").Index - 1) = "1-必须"
            Case 2
                objItem.SubItems(Me.lvwItem.ColumnHeaders("_可行病检").Index - 1) = "2-选择进行"
            Case Else
                objItem.SubItems(Me.lvwItem.ColumnHeaders("_可行病检").Index - 1) = "0-不可能"
            End Select
            Select Case !可发胶片
            Case 1
                objItem.SubItems(Me.lvwItem.ColumnHeaders("_可发胶片").Index - 1) = "1-必须"
            Case 2
                objItem.SubItems(Me.lvwItem.ColumnHeaders("_可发胶片").Index - 1) = "2-选择发放"
            Case Else
                objItem.SubItems(Me.lvwItem.ColumnHeaders("_可发胶片").Index - 1) = "0-不可能"
            End Select
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_报告图象").Index - 1) = IIf(IsNull(!报告图象), "", !报告图象)
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_检查准备").Index - 1) = IIf(IsNull(!检查准备), "", !检查准备)
            .MoveNext
        Loop
    End With
    If Me.lvwItem.ListItems.count > 0 Then
        Err = 0: On Error Resume Next
        Me.lvwItem.ListItems("_" & lngItemId).Selected = True
        If Me.lvwItem.SelectedItem Is Nothing Then Me.lvwItem.ListItems(1).Selected = True
        Me.lvwItem.SelectedItem.EnsureVisible
        Me.stbThis.Panels(2).Text = "该类别共" & Me.lvwItem.ListItems.count & "个项目"
    Else
        Me.stbThis.Panels(2).Text = ""
    End If
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
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

