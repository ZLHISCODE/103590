VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmCollectionManage 
   Caption         =   "收藏管理"
   ClientHeight    =   7920
   ClientLeft      =   120
   ClientTop       =   345
   ClientWidth     =   12000
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCollectionManage.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7920
   ScaleWidth      =   12000
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5775
      Left            =   120
      ScaleHeight     =   5775
      ScaleWidth      =   12015
      TabIndex        =   0
      Top             =   1080
      Width           =   12015
      Begin zl9PacsControl.ucSplitter ucSplitter 
         Height          =   5775
         Left            =   3255
         TabIndex        =   4
         Top             =   0
         Width           =   90
         _ExtentX        =   159
         _ExtentY        =   10186
         BackColor       =   -2147483644
         SplitWidth      =   90
         SplitLevel      =   3
         Con1MinSize     =   2250
         Con2MinSize     =   2430
         Control1Name    =   "PicTvw"
         Control2Name    =   "PicData"
      End
      Begin VB.PictureBox PicData 
         BorderStyle     =   0  'None
         Height          =   5775
         Left            =   3345
         ScaleHeight     =   5775
         ScaleWidth      =   8670
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   0
         Width           =   8670
         Begin VSFlex8Ctl.VSFlexGrid vfgCollectionData 
            Height          =   4455
            Left            =   1560
            TabIndex        =   3
            Top             =   600
            Width           =   5535
            _cx             =   9763
            _cy             =   7858
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483636
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   3
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   ""
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   0
            AutoSearch      =   0
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   7
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   0
            ShowComboButton =   1
            WordWrap        =   0   'False
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            DataMode        =   0
            VirtualData     =   -1  'True
            DataMember      =   ""
            ComboSearch     =   3
            AutoSizeMouse   =   -1  'True
            FrozenRows      =   0
            FrozenCols      =   0
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
      End
      Begin VB.PictureBox PicTvw 
         BorderStyle     =   0  'None
         Height          =   5775
         Left            =   0
         ScaleHeight     =   5775
         ScaleWidth      =   3255
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   3255
         Begin MSComctlLib.ImageList imgList 
            Left            =   2640
            Top             =   120
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   3
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCollectionManage.frx":6852
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCollectionManage.frx":6BEC
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCollectionManage.frx":6F86
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.TreeView tvwCollectionType 
            Height          =   5295
            Left            =   240
            TabIndex        =   6
            Top             =   1320
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   9340
            _Version        =   393217
            Indentation     =   494
            LabelEdit       =   1
            Style           =   7
            ImageList       =   "imgList"
            Appearance      =   1
         End
         Begin VB.Image imgTree 
            Height          =   720
            Index           =   3
            Left            =   2040
            Picture         =   "frmCollectionManage.frx":7320
            Top             =   1680
            Width           =   720
         End
         Begin VB.Image imgTree 
            Height          =   720
            Index           =   1
            Left            =   480
            Picture         =   "frmCollectionManage.frx":DB72
            Top             =   1680
            Width           =   720
         End
         Begin VB.Image imgTree 
            Height          =   720
            Index           =   2
            Left            =   1320
            Picture         =   "frmCollectionManage.frx":143C4
            Top             =   1680
            Width           =   720
         End
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   7560
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmCollectionManage.frx":1AC16
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14288
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin XtremeCommandBars.ImageManager imgPopup 
      Left            =   1320
      Top             =   360
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmCollectionManage.frx":1B4AA
   End
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   480
      Top             =   360
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmCollectionManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strSql As String
Private mstrNodeKey As String
Private mstrNodeName As String
Private mobjNode As Node
Private mrsTvwData As ADODB.Recordset
Private mobjSourceNode As Object
'菜单
Private Enum popMenus
    conMenu_Edit_Add = 100
    conMenu_Edit_Rename = 101
    conMenu_Edit_Del = 102
    conMenu_Edit_DelColl = 103
    conMenu_Edit_Share = 104
End Enum

Public Sub ShowCollectionManageWind(Optional owner As Form = Nothing)
'显示收藏管理窗口
    Call Me.Show(1, owner)
End Sub

Private Sub cbrMain_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
    On Error GoTo ErrorHand
    
    Select Case control.ID
        Case conMenu_Edit_Add
            Call Menu_Edit_Add
            
        Case conMenu_Edit_Rename
            Call Menu_Edit_Rename
            
        Case conMenu_Edit_Del
            Call Menu_Edit_Del
            
        Case conMenu_Edit_DelColl
            Call Menu_Edit_DelColl
            
        Case conMenu_Edit_Share
            Call Menu_Edit_Share(control)

        Case conMenu_File_Exit
            Call Menu_File_Exit
            
        Case conMenu_View_Refresh
            Call LoadTreeView
'---------------------------查看----------------
        Case conMenu_View_ToolBar_Button '工具栏
            Call Menu_View_ToolBar_Button_click(control)

        Case conMenu_View_ToolBar_Text '按钮文字
            Call Menu_View_ToolBar_Text_click(control)

        Case conMenu_View_StatusBar '状态栏
            Call Menu_View_StatusBar_click(control)
            
'--------------------------帮助-----------------
        Case conMenu_Help_Help
            Call Menu_Help_Help_click

        Case conMenu_Help_Web_Forum
            Call Menu_Help_Web_Forum_click

        Case conMenu_Help_Web_Home
            Call Menu_Help_Web_Home_click

        Case conMenu_Help_Web_Mail
            Call Menu_Help_Web_Mail_click

        Case conMenu_Help_About
            Call Menu_Help_About_click
    End Select
    Call Form_Resize
    Exit Sub
ErrorHand:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_File_Exit()
    Unload Me
End Sub

Private Sub Menu_Help_About_click()
    ShowAbout Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision
End Sub

Private Sub Menu_Help_Web_Mail_click()
    zlMailTo hWnd
End Sub

Private Sub Menu_Help_Web_Home_click()
    zlHomePage hWnd
End Sub

Private Sub Menu_Help_Web_Forum_click()
    Call zlWebForum(Me.hWnd)
End Sub

Private Sub Menu_View_ToolBar_Button_click(ByVal control As XtremeCommandBars.ICommandBarControl)
Dim i As Integer
    For i = 2 To cbrMain.Count
        Me.cbrMain(i).Visible = Not Me.cbrMain(i).Visible
    Next

    control.Checked = Not control.Checked
    Me.cbrMain.RecalcLayout
End Sub

Private Sub Menu_View_ToolBar_Text_click(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error GoTo ErrorHand
    Dim i As Integer, cbrControl As CommandBarControl
    Dim intStyle As Integer

    For i = 2 To cbrMain.Count
        If Me.cbrMain(i).Controls.Count >= 1 Then
            intStyle = Me.cbrMain(i).Controls(i).Style
            If intStyle = xtpButtonIconAndCaption Then
                intStyle = xtpButtonIcon
                Me.cbrMain(i).ShowTextBelowIcons = False
            Else
                intStyle = xtpButtonIconAndCaption
                Me.cbrMain(i).ShowTextBelowIcons = True
            End If
        End If
        
        For Each cbrControl In Me.cbrMain(i).Controls
            cbrControl.Style = intStyle
        Next
    Next
    
    control.Checked = Not control.Checked
    Me.cbrMain.RecalcLayout
    
    Exit Sub
ErrorHand:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_View_StatusBar_click(ByVal control As XtremeCommandBars.ICommandBarControl)
    Me.stbThis.Visible = Not Me.stbThis.Visible
    control.Checked = Not control.Checked
    Me.cbrMain.RecalcLayout
End Sub

Private Sub Menu_Help_Help_click()
    '功能：调用帮助主题
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub cbrMain_Update(ByVal control As XtremeCommandBars.ICommandBarControl)
    Dim rsTemp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo ErrorHand
    If tvwCollectionType.SelectedItem Is Nothing Then Exit Sub
    control.Enabled = True
    Select Case control.ID
        Case conMenu_Edit_Add
            
            
        Case conMenu_Edit_Rename
            If tvwCollectionType.SelectedItem.Text = "收藏类别" Then control.Enabled = False
            
        Case conMenu_Edit_Del
            If tvwCollectionType.SelectedItem.Text = "收藏类别" Then control.Enabled = False
            
        Case conMenu_Edit_DelColl
            
            
        Case conMenu_Edit_Share
            If tvwCollectionType.SelectedItem.Text = "收藏类别" Then
                control.Enabled = False
                Exit Sub
            End If
            
            control.Caption = IIf(tvwCollectionType.SelectedItem.Tag = 3, "取消共享", "设置共享")
            If control.Caption = "取消共享" Then control.ToolTipText = "取消此收藏为共享"
        Case conMenu_File_Exit
            
        
    End Select
    Exit Sub
ErrorHand:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Load()
On Error GoTo errHand
    
    InitCommandBars
    '加载TreeView数据
    Call LoadTreeView
     
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub InitVfgData()
    Dim i As Integer
'初始化数据控件
    With vfgCollectionData
        .Clear
        .FixedRows = 1
        .Cols = 13
        .ColWidth(0) = 500
        .ColWidth(1) = 900
        .ColWidth(2) = 900
        .ColWidth(3) = 900
        .ColWidth(4) = 1000
        .ColWidth(5) = 700
        .ColWidth(6) = 700
        .ColWidth(7) = 2500
        .ColWidth(8) = 3000
        .ColWidth(9) = 1000
        .ColWidth(10) = 1200
        .ColWidth(11) = 1200
        .TextMatrix(0, 1) = "检查号"
        .TextMatrix(0, 2) = "门诊号"
        .TextMatrix(0, 3) = "住院号"
        .TextMatrix(0, 4) = "姓名"
        .TextMatrix(0, 5) = "性别"
        .TextMatrix(0, 6) = "年龄"
        .TextMatrix(0, 7) = "医嘱内容"
        .TextMatrix(0, 8) = "部位方法"
        .TextMatrix(0, 9) = "开嘱医生"
        .TextMatrix(0, 10) = "开嘱时间"
        .TextMatrix(0, 11) = "收藏时间"
        .TextMatrix(0, 12) = "关联ID"
        For i = 0 To 12
            .ColAlignment(i) = flexAlignLeftCenter
        Next
        .AllowSelection = True
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        .ColHidden(12) = True
        .AllowUserResizing = flexResizeColumns
    End With

End Sub

Private Sub LoadTreeView()
'加载TreeView数据方法
    Dim strCurrKey As String
    Dim rsTemp As ADODB.Recordset
    Dim dtServicesTime As Date
    Dim strSql As String
err = 0: On Error GoTo errHand

    strSql = "select id from 影像收藏类别 where 收藏类别='收藏类别' "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    
    '如果数据库中没有顶级节点数据，则插入顶级节点数据
    If rsTemp.RecordCount <= 0 Then
         '当前服务器时间
        dtServicesTime = zlDatabase.Currentdate
        
        strSql = "select Zl_影像收藏类别_新增([1],[2],[3],[4],[5]) as 返回值 from dual"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, _
                                "", _
                                "收藏类别", _
                                0, _
                                "", _
                                dtServicesTime)
    
    End If

    strSql = "select ID,上级ID,收藏类别,是否共享 from 影像收藏类别 where 创建人= '" & UserInfo.姓名 & "' or 创建人 is null Start With 上级id Is Null Connect By Prior ID = 上级id"
    Set mrsTvwData = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    
    With mrsTvwData
        Me.tvwCollectionType.Nodes.Clear
        
        If Not tvwCollectionType.SelectedItem Is Nothing Then strCurrKey = tvwCollectionType.SelectedItem.Key
        Do While Not .EOF
            If IsNull(!上级ID) Then
                Set mobjNode = Me.tvwCollectionType.Nodes.Add(, , "_" & Nvl(!ID), Nvl(!收藏类别), IIf(!是否共享 = 0, 1, 3), IIf(Nvl(!是否共享) = 0, 2, 3))
                mobjNode.Tag = IIf(Nvl(!是否共享) = 0, 2, 3)
            Else
                Set mobjNode = Me.tvwCollectionType.Nodes.Add("_" & Nvl(!上级ID), tvwChild, "_" & Nvl(!ID), Nvl(!收藏类别), IIf(Nvl(!是否共享) = 0, 1, 3), IIf(Nvl(!是否共享) = 0, 2, 3))
                mobjNode.Tag = IIf(Nvl(!是否共享) = 0, 2, 3)
            End If
            
            mobjNode.Sorted = True
            mobjNode.Expanded = True
            If strCurrKey = mobjNode.Key Then mobjNode.Selected = True
            .MoveNext
        Loop
    End With
    
    '如果加载时自动选中则加载 右侧收藏的数据
err = 0: On Error GoTo 0
    If Me.tvwCollectionType.Nodes.Count > 0 Then
        If tvwCollectionType.SelectedItem Is Nothing Then Me.tvwCollectionType.Nodes(1).Selected = True
        tvwCollectionType.SelectedItem.EnsureVisible
        Call tvwCollectionType_NodeClick(Me.tvwCollectionType.SelectedItem)
    End If

    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub tvwCollectionType_AfterLabelEdit(Cancel As Integer, NewString As String)
    Dim rsTemp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errHand
    tvwCollectionType.DragIcon = imgTree(1).Picture
    If NewString = "" Then NewString = tvwCollectionType.SelectedItem.Text
    If tvwCollectionType.SelectedItem.Text = NewString Then Exit Sub
    '判断修改类型是否重复
    strSql = "select 收藏类别 from 影像收藏类别 where 创建人= [1] and 收藏类别=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UserInfo.姓名, NewString)
    
    If rsTemp.RecordCount > 0 Then
        Call MsgBoxD(Me, "收藏类型重复。", vbOKOnly, Me.Caption)
        '收藏类型重复则保留原来的名字
        NewString = tvwCollectionType.SelectedItem.Text
        tvwCollectionType.SelectedItem.Selected = True
        Exit Sub
    End If
    
    strSql = "Zl_影像收藏类别_更新(" & Val(Mid(tvwCollectionType.SelectedItem.Key, 2)) & ",'" & _
                        Decode(Trim(NewString), "", tvwCollectionType.SelectedItem.Text, Trim(NewString)) & "',2)"

    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub tvwCollectionType_DragDrop(Source As control, X As Single, Y As Single)
    Dim strCurParent As String
    Dim lngNodesKey As Long
    
On Error Resume Next
    If mobjSourceNode Is Nothing Then Exit Sub
    strCurParent = mobjSourceNode.Parent.Text
    If strCurParent = "" Then Exit Sub
    If Not (tvwCollectionType.DropHighlight Is Nothing) Then
        Set mobjSourceNode.Parent = tvwCollectionType.DropHighlight
        Set tvwCollectionType.DropHighlight = Nothing
        
        If strCurParent <> mobjSourceNode.Parent.Text Then
            strSql = "Zl_影像收藏类别_更新分类(" & Mid(mobjSourceNode.Key, 2) & _
                                          "," & Val(Mid(tvwCollectionType.SelectedItem.Parent.Key, 2)) & ")"
            Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
        End If
    End If
    
    Set mobjSourceNode = Nothing
    
End Sub

Private Sub tvwCollectionType_DragOver(Source As control, X As Single, Y As Single, State As Integer)
    Dim objNode As Node
    Dim objTargetNode As Object
    
    On Error GoTo errHand

    If mobjSourceNode Is Nothing Then Exit Sub
    
    Set objNode = tvwCollectionType.HitTest(X, Y)
    
    If objNode Is objTargetNode Then Exit Sub
    Set objTargetNode = objNode
    
    Set tvwCollectionType.DropHighlight = objTargetNode
    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub tvwCollectionType_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errHand
    If KeyCode = vbKeyF2 Then
        If tvwCollectionType.SelectedItem.Text <> "收藏类别" Then tvwCollectionType.StartLabelEdit
    ElseIf KeyCode = vbKeyF5 Then
        Call LoadTreeView
    End If
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub tvwCollectionType_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo errHand
    If tvwCollectionType.HitTest(X, Y) Is Nothing Then Exit Sub
    Set mobjSourceNode = tvwCollectionType.HitTest(X, Y)
    tvwCollectionType.SelectedItem = tvwCollectionType.HitTest(X, Y)
    If mobjSourceNode.Text = "收藏类别" Then Set mobjSourceNode = Nothing
    '刷新按钮状态
    cbrMain.RecalcLayout
    If tvwCollectionType.HitTest(X, Y).Text <> "收藏类别" Then tvwCollectionType_NodeClick mobjSourceNode
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub tvwCollectionType_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If mobjSourceNode Is Nothing Then Exit Sub
    
    tvwCollectionType.DragIcon = IIf(tvwCollectionType.SelectedItem.Tag = 2, imgTree(2).Picture, imgTree(3).Picture)
    If Button = vbLeftButton Then
        Set tvwCollectionType.SelectedItem = mobjSourceNode
        tvwCollectionType.Drag vbBeginDrag
    End If
    
    Exit Sub
End Sub

Private Sub tvwCollectionType_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'弹出右键菜单
On Error GoTo ErrHandle
    If Button = 2 Then
        Dim objPopup As CommandBar
        Dim objControl As CommandBarControl

        Set cbrMain.Icons = imgPopup.Icons
        Set objPopup = cbrMain.Add("右键菜单", xtpBarPopup)
        With objPopup.Controls
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Add, "新增类别(&A)")
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Rename, "重命名(&U)")
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Del, "删除类别(&D)")
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Share, "设置共享(&S)")
        End With
        objPopup.ShowPopup
    End If
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub tvwCollectionType_NodeClick(ByVal Node As MSComctlLib.Node)
On Error GoTo errHand
'加载选中节点下的数据
    Dim rsCollectionData As ADODB.Recordset
    Dim rsIsShare As ADODB.Recordset
    Dim i As Long
    Dim strSql As String
    Dim strAdviceTemp As String
    
    On Error GoTo errHand
    Set mobjNode = Node
    '得到节点的Key
    mstrNodeKey = Mid(Node.Key, 2)
    
    strSql = "select distinct e.id,nvl(c.姓名,d.姓名) 姓名,nvl(c.年龄,d.年龄) 年龄,nvl(c.性别,d.性别) 性别,a.医嘱内容,a.开嘱医生,a.开嘱时间,c.门诊号,c.住院号,d.检查号,f.是否共享,e.收藏时间 " & _
            " from 病人医嘱记录 a,病人医嘱发送 b,病人信息 c,影像检查记录 d,影像收藏内容 e,影像收藏类别 f" & _
            " where a.id = b.医嘱id and b.医嘱ID=d.医嘱ID(+)" & _
            " and a.病人ID=c.病人ID and a.相关id is null" & _
            " and b.医嘱id = e.医嘱id and e.收藏id = f.id and f.创建人='" & UserInfo.姓名 & _
            "' and f.id in (select distinct id from 影像收藏类别 start with id = " & mstrNodeKey & " connect by prior id=上级id) order by e.id"

    Set rsCollectionData = zlDatabase.OpenSQLRecord(strSql, Me.Caption)

    With vfgCollectionData
        .Clear
        
        If rsCollectionData.RecordCount = 0 Then
            .Rows = 1
        Else
            .Rows = rsCollectionData.RecordCount + 1
        End If
        
        '初始化数据显示控件
        Call InitVfgData
        
        For i = 1 To rsCollectionData.RecordCount
        
            strAdviceTemp = Nvl(rsCollectionData!医嘱内容)
            
            .TextMatrix(i, 0) = i
            .TextMatrix(i, 1) = Nvl(rsCollectionData!检查号)
            .TextMatrix(i, 2) = Nvl(rsCollectionData!门诊号)
            .TextMatrix(i, 3) = Nvl(rsCollectionData!住院号)
            .TextMatrix(i, 4) = Nvl(rsCollectionData!姓名)
            .TextMatrix(i, 5) = Nvl(rsCollectionData!性别)
            .TextMatrix(i, 6) = Nvl(rsCollectionData!年龄)
            
            '处理医嘱内容的医嘱部分和部位方法部分
            .TextMatrix(i, 7) = Mid(strAdviceTemp, 1, InStr(strAdviceTemp, ":") - 1)
            .TextMatrix(i, 8) = Mid(strAdviceTemp, InStr(strAdviceTemp, ":") + 1, Len(strAdviceTemp))
            
            .TextMatrix(i, 9) = Nvl(rsCollectionData!开嘱医生)
            .TextMatrix(i, 10) = Format(Nvl(rsCollectionData!开嘱时间), "yyyy-mm-dd")
            .TextMatrix(i, 11) = Format(Nvl(rsCollectionData!收藏时间), "yyyy-mm-dd")
            .TextMatrix(i, 12) = Nvl(rsCollectionData!ID)

            If Not rsCollectionData.EOF Then rsCollectionData.MoveNext
        Next
    End With
    
    stbThis.Panels(2).Text = "当前收藏类别下有 " & rsCollectionData.RecordCount & " 个收藏"
    
    cbrMain.RecalcLayout
    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Menu_Edit_Add()
'新增收藏类型
On Error GoTo errHand
    Dim lngNodesKey As Long
    Dim rsTemp As ADODB.Recordset
    Dim dtServicesTime As Date
    Dim strSql As String
      
    '当前服务器时间
    dtServicesTime = zlDatabase.Currentdate

    strSql = "select Zl_影像收藏类别_新增([1],[2],[3],[4],[5]) as 返回值 from dual"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, _
                                Val(Mid(tvwCollectionType.SelectedItem.Key, 2)), _
                                "新建节点" & GetNextIndex, _
                                0, _
                                UserInfo.姓名, _
                                dtServicesTime)

    If rsTemp.RecordCount > 0 Then lngNodesKey = Nvl(rsTemp!返回值)
    
    '在treeView控件上添加新增节点
    Set mobjNode = Me.tvwCollectionType.Nodes.Add("_" & Val(Mid(tvwCollectionType.SelectedItem.Key, 2)), tvwChild, "_" & lngNodesKey, "新建节点" & GetNextIndex, 1, 2)
    mobjNode.Selected = True
    mobjNode.Tag = 2
    tvwCollectionType.StartLabelEdit
                
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function GetNextIndex() As String
    Dim i                   As Integer
    Dim j                   As Integer
    Dim strName()           As String
    Dim strTemp             As String

    On Error GoTo ErrorHand
    
    mstrNodeName = ""
    Call GetAllNode(mobjNode.Root)
    strName() = Split(mstrNodeName, "|")
    For i = 0 To UBound(strName()) - 1
        For j = i + 1 To UBound(strName()) - 1
            If CInt(Mid(strName(i), 5)) > CInt(Mid(strName(j), 5)) Then
                strTemp = strName(i)
                strName(i) = strName(j)
                strName(j) = strTemp
            End If
        Next
    Next
    For i = 0 To UBound(strName()) - 1
        If "新建节点" & i + 1 <> strName(i) Then
            GetNextIndex = i + 1
            Exit Function
        End If
    Next
    GetNextIndex = i + 1
    Exit Function
ErrorHand:
    GetNextIndex = i + 1
End Function

Private Sub GetAllNode(ByVal Node As MSComctlLib.Node)
    Dim objNode As Node
    
    If Node.Children > 0 Then
        Set objNode = Node.Child
        Do While Not objNode Is Nothing
            If InStr(objNode.Text, "新建节点") > 0 Then
                If objNode.Text <> "新建节点" Then mstrNodeName = mstrNodeName & objNode.Text & "|"
            End If
            Call GetAllNode(objNode)
            Set objNode = objNode.Next
        Loop
    End If
End Sub

Private Sub Menu_Edit_Del()
'删除收藏类型(级联删除)
On Error GoTo errHand
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset

    If tvwCollectionType.SelectedItem.Children <> 0 Then
        Call MsgBoxD(Me, "该类型下有子类型，不能删除。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If vfgCollectionData.Rows > 1 Then
        If MsgBoxD(Me, "确定删除吗？(删除该类型会删除收藏信息)", vbOKCancel, Me.Caption) = 2 Then Exit Sub
    End If
    
    strSql = "Zl_影像收藏类别_删除(" & Val(Mid(tvwCollectionType.SelectedItem.Key, 2)) & ")"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    '刷新
    tvwCollectionType_NodeClick tvwCollectionType.SelectedItem.Parent
    '在TreeView控件中删除选中节点
    tvwCollectionType.Nodes.Remove (tvwCollectionType.SelectedItem.Key)
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Menu_Edit_Rename()
'更新收藏类型
On Error GoTo errHand
Dim strSql As String
Dim strCurNodeText As String
    
    Set mobjNode = Me.tvwCollectionType.SelectedItem
    strCurNodeText = mobjNode.Text
    mobjNode.Selected = True
    tvwCollectionType.DragIcon = imgTree(1).Picture
    tvwCollectionType.StartLabelEdit
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Edit_Share(ByVal control As XtremeCommandBars.ICommandBarControl)
'更新收藏共享状态
On Error GoTo errHand
Dim strSql As String

    '只更新共享状态
    strSql = "Zl_影像收藏类别_更新(" & Val(Mid(tvwCollectionType.SelectedItem.Key, 2)) & ",null," & IIf(control.Caption = "取消共享", 0, 1) & ")"

    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    tvwCollectionType.SelectedItem.SelectedImage = IIf(control.Caption = "取消共享", 2, 3)
    tvwCollectionType.SelectedItem.Image = IIf(control.Caption = "取消共享", 1, 3)
    tvwCollectionType.SelectedItem.Tag = IIf(control.Caption = "取消共享", 2, 3)
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub Menu_Edit_DelColl()
'删除收藏检查关联
On Error GoTo errHand
Dim strSql As String
Dim i As Integer

    If Me.vfgCollectionData.SelectedRows = 0 Then
       Call MsgBoxD(Me, "请先选择要删除的收藏数据。", vbOKOnly, Me.Caption)
       Exit Sub
    End If
    
    If MsgBoxD(Me, "您确定要删除所选择的收藏数据吗？", vbOKCancel, Me.Caption) = 2 Then Exit Sub
    With vfgCollectionData
        For i = 0 To .SelectedRows - 1
            strSql = "Zl_影像收藏内容_删除(" & .TextMatrix(.SelectedRow(0), 12) & ")"
            Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
            '删除VSFlexGrid数据
            vfgCollectionData.RemoveItem (vfgCollectionData.SelectedRow(0))
        Next
    End With
     
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub InitCommandBars()
    '功能创建工具条
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrPopControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    
    '设置菜单栏和工具栏风格
    With cbrMain.Options
        .ShowExpandButtonAlways = False                         '总是在工具栏右侧显示选项按钮,即使窗体宽度足够。
        .ToolBarAccelTips = True                                '显示按钮提示
        .AlwaysShowFullMenus = False                            '不常用的菜单项先隐藏
        .UseFadedIcons = False                                  '图标显示为褪色效果
        .IconsWithShadow = True                                 '鼠标指向的命令图标显示阴影效果
        .UseDisabledIcons = True                                '工具栏按钮禁用时图标显示为禁用样式
        .LargeIcons = True                                      '工具栏显示为大图标
        .SetIconSize True, 24, 24                               '设置大图标的尺寸
        .SetIconSize False, 16, 16                              '设置小图标的尺寸
    End With
    With cbrMain
        .VisualTheme = xtpThemeOffice2003                      '设置控件显示风格
        .EnableCustomization False                              '是否允许自定义设置
        Set .Icons = imgPopup.Icons                           '设置关联的图标控件
    End With

    Me.cbrMain.EnableCustomization False
    Me.cbrMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    
    '菜单定义
'Begin------------------------编辑菜单--------------------------------------默认可见
    cbrMain.ActiveMenuBar.Title = "菜单"
    
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)")
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, conMenu_File_Exit, "退出(&Q)")
    
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)")
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Add, "新增类别(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Rename, "重命名(&U)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Del, "删除类别(&D)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_DelColl, "删除收藏(&M)")
        cbrControl.BeginGroup = True
    End With
    
    'Begin----------------------查看菜单--------------------------------------
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(V)")
    With cbrMenuBar.CommandBar
        Set cbrControl = .Controls.Add(xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(T)")
        cbrControl.ID = conMenu_View_ToolBar
            With cbrControl.CommandBar '二级菜单
                Set cbrPopControl = .Controls.Add(xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(0)"): cbrPopControl.Checked = True
                Set cbrPopControl = .Controls.Add(xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(1)"): cbrPopControl.Checked = True
            End With
        Set cbrControl = .Controls.Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(S)"): cbrControl.Checked = True
        Set cbrControl = .Controls.Add(xtpControlButton, conMenu_View_Refresh, "刷新(R)")
    End With

    'Begin----------------------帮助菜单--------------------------------------默认可见
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(H)")
    With cbrMenuBar.CommandBar
        Set cbrControl = .Controls.Add(xtpControlButton, conMenu_Help_Help, "帮助主题(M)")
        Set cbrControl = .Controls.Add(xtpControlButtonPopup, conMenu_Help_Web, "WEB上的中联(W)")
            With cbrControl.CommandBar
                Set cbrPopControl = .Controls.Add(xtpControlButton, conMenu_Help_Web_Forum, "中联论坛(0)")
                Set cbrPopControl = .Controls.Add(xtpControlButton, conMenu_Help_Web_Home, "中联主页(1)")
                Set cbrPopControl = .Controls.Add(xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(2)")
            End With
        Set cbrControl = .Controls.Add(xtpControlButton, conMenu_Help_About, "关于…(A)")
    End With
    '---------------------工具栏定义------------------------------------------
    Set cbrToolBar = Me.cbrMain.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = True
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, conMenu_Edit_Add, "新增类别", "新增类别")
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, conMenu_Edit_Rename, "重命名", "重命名")
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, conMenu_Edit_Del, "删除类别", "删除类别")
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, conMenu_Edit_Share, "设置共享")
    cbrControl.ToolTipText = "将此收藏设置为共享"
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, conMenu_Edit_DelColl, "删除收藏", "删除收藏")
    cbrControl.BeginGroup = True
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, conMenu_View_Refresh, "刷新", "刷新")
    cbrControl.BeginGroup = True
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, conMenu_File_Exit, "退出", "退出")
    cbrControl.BeginGroup = True
    
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Dim lngLeft As Long, lngTop As Long, lngRight As Long, lngBottom As Long
    
    cbrMain.GetClientRect lngLeft, lngTop, lngRight, lngBottom
    
    picMain.Left = 0
    picMain.Top = lngTop
    picMain.Width = Me.ScaleWidth
    If stbThis.Visible = True Then
        picMain.Height = Me.ScaleHeight - picMain.Top - stbThis.Height
    Else
        picMain.Height = Me.ScaleHeight - picMain.Top
    End If
    
    '窗体改变,调用用户控件自动调整方法
    ucSplitter.RePaint
End Sub

Private Sub PicTvw_Resize()
    On Error Resume Next
    
    tvwCollectionType.Top = 0
    tvwCollectionType.Left = 60
    tvwCollectionType.Height = PicTvw.Height
    tvwCollectionType.Width = PicTvw.Width - 60
End Sub

Private Sub PicData_Resize()
    On Error Resume Next
    
    vfgCollectionData.Top = 0
    vfgCollectionData.Left = 0
    vfgCollectionData.Height = PicData.Height
    vfgCollectionData.Width = PicData.Width - 60
End Sub

Private Sub vfgCollectionData_Click()
    On Error GoTo errHand
    tvwCollectionType.HideSelection = False
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
End Sub




