VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmPatholMeal 
   Caption         =   "套餐维护"
   ClientHeight    =   8190
   ClientLeft      =   75
   ClientTop       =   405
   ClientWidth     =   11295
   Icon            =   "frmPatholMeal.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8190
   ScaleWidth      =   11295
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picDatas 
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   240
      ScaleHeight     =   3735
      ScaleWidth      =   10455
      TabIndex        =   6
      Top             =   360
      Width           =   10455
      Begin VB.Frame framMeals 
         Caption         =   "套餐记录"
         Height          =   3495
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   10215
         Begin VB.ComboBox cboMealClass 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   320
            Width           =   1935
         End
         Begin zl9PACSWork.ucFlexGrid ufgMeal 
            Height          =   2415
            Left            =   120
            TabIndex        =   9
            Top             =   840
            Width           =   9975
            _ExtentX        =   17595
            _ExtentY        =   4260
            IsKeepRows      =   0   'False
            DisCellColor    =   16777215
            IsCopyAdoMode   =   0   'False
            IsEjectConfig   =   -1  'True
            Editable        =   0
            HeadFontCharset =   134
            HeadFontWeight  =   400
            DataFontCharset =   134
            DataFontWeight  =   400
         End
         Begin VB.Label lblMealClass 
            Caption         =   "套餐类别："
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   360
            Width           =   975
         End
      End
   End
   Begin VB.PictureBox picMealLink 
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   240
      ScaleHeight     =   3735
      ScaleWidth      =   10455
      TabIndex        =   0
      Top             =   4080
      Width           =   10455
      Begin VB.Frame framAntibody 
         Caption         =   "抗体明细"
         Height          =   3495
         Left            =   0
         TabIndex        =   1
         Top             =   120
         Width           =   10215
         Begin VB.CheckBox chkRowFilter 
            Caption         =   "只显示勾选行"
            Height          =   180
            Left            =   240
            TabIndex        =   3
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox txtFind 
            Height          =   300
            Left            =   3480
            TabIndex        =   2
            ToolTipText     =   "根据抗体名称进行快速定位。"
            Top             =   300
            Width           =   1695
         End
         Begin zl9PACSWork.ucFlexGrid ufgMealLink 
            Height          =   2535
            Left            =   120
            TabIndex        =   4
            Top             =   840
            Width           =   9975
            _ExtentX        =   17595
            _ExtentY        =   4471
            IsKeepRows      =   0   'False
            BackColor       =   12648447
            IsCopyAdoMode   =   0   'False
            IsEjectConfig   =   -1  'True
            Editable        =   0
            HeadFontCharset =   134
            HeadFontWeight  =   400
            DataFontCharset =   134
            DataFontWeight  =   400
         End
         Begin VB.Label lblFind 
            Caption         =   "快速查找："
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2400
            TabIndex        =   5
            ToolTipText     =   "根据抗体名称进行快速定位。"
            Top             =   360
            Width           =   1095
         End
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   11
      Top             =   7830
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPatholMeal.frx":179A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13044
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
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   600
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMeal 
      Bindings        =   "frmPatholMeal.frx":202E
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmPatholMeal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrPrivs As String
Private mblnEdit As Boolean
Private mblnIsUpdate As Boolean
Private mblnCurModifyState As Boolean

Public Sub ShowMealWindow(ByVal strPrivs As String, owner As Form)
'显示套餐维护窗口
    mstrPrivs = strPrivs
    
    Call ConfigPopedom
    
    Call Me.Show(1, owner)
End Sub


Private Sub ConfigPopedom()
'配置权限
    Dim blnIsAllowMeal As Boolean
    
    blnIsAllowMeal = CheckPopedom(mstrPrivs, "套餐维护")

    mblnCurModifyState = blnIsAllowMeal
    ufgMealLink.Enabled = False
    ufgMealLink.DataGrid.Enabled = True
    ufgMealLink.DataGrid.BackColor = &H8000000F
    mblnIsUpdate = False
End Sub

Private Sub InitMealList()
'初始化套餐显示列表
    Dim strTemp As String
    
     '判断数据库参数表是否有数据 有则读取数据库参数  没有则加载默认
    strTemp = zlDatabase.GetPara("抗体套餐列表配置", glngSys, G_LNG_PATHOLSYS_NUM, "")
    ufgMeal.DefaultColNames = gstrAntibodyMealCols
     
    If strTemp = "" Then
        ufgMeal.ColNames = gstrAntibodyMealCols
    Else
        ufgMeal.ColNames = strTemp
    End If
    
    ufgMeal.IsCopyMode = True
    '禁止右键弹出列表配置窗口
    ufgMeal.IsEjectConfig = False
    '设置行数
    ufgMeal.GridRows = glngStandardRowCount
    '设置行高
    ufgMeal.RowHeightMin = glngStandardRowHeight
    ufgMeal.ColConvertFormat = gstrAntibodyMealConvertFormat
    ufgMeal.IsShowPopupMenu = False
End Sub

Private Sub InitMealLinkList()
'初始化套餐明细列表
    Dim strTemp As String
    
     '判断数据库参数表是否有数据 有则读取数据库参数  没有则加载默认
    strTemp = zlDatabase.GetPara("套餐明细列表配置", glngSys, G_LNG_PATHOLSYS_NUM, "")
    ufgMealLink.DefaultColNames = gstrAntibodyMealLinkCols
     
    If strTemp = "" Then
        ufgMealLink.ColNames = gstrAntibodyMealLinkCols
    Else
        ufgMealLink.ColNames = strTemp
    End If
    
    '禁止右键弹出列表配置窗口
    ufgMealLink.IsEjectConfig = False
      '设置行数
    ufgMealLink.GridRows = glngStandardRowCount
    '设置行高
    ufgMealLink.RowHeightMin = glngStandardRowHeight
    ufgMealLink.ColConvertFormat = gstrAntibodyMealLinkConvertFormat
    ufgMealLink.IsShowPopupMenu = False
End Sub

Private Sub cbrMain_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
    On Error GoTo ErrorHand
    
    Select Case control.ID
        Case conMenu_PatholMeal_Save              '保存
            Call Menu_PatholMeal_Save

        Case conMenu_PatholMeal_Cancel            '取消
            Call Menu_PatholMeal_Cancel

        Case conMenu_PatholMeal_AddRecord         '新增
            Call Menu_PatholMeal_AddMeal
            
        Case conMenu_PatholMeal_ModRecord         '修改
            Call Menu_PatholMeal_ModMeal
            
        Case conMenu_PatholMeal_DelRecord         '删除
            Call Menu_PatholMeal_DelRecord
            
        Case conMenu_PatholMeal_UpRow             '上移
            Call Menu_PatholMeal_UpRow
            
        Case conMenu_PatholMeal_DownRow           '下移
            Call Menu_PatholMeal_DownRow
            
        Case conMenu_File_Exit                    '退出
            Call Menu_File_Exit
        
        '---------------------------查看----------------
        Case conMenu_View_ToolBar_Button          '工具栏
            Call Menu_View_ToolBar_Button_click(control)

        Case conMenu_View_ToolBar_Text            '按钮文字
            Call Menu_View_ToolBar_Text_click(control)

        Case conMenu_View_StatusBar               '状态栏
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
    Exit Sub
ErrorHand:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_PatholMeal_Save()
    '保存前得先让处于编辑状态的单元格失去焦点，否则检测不到输入的值
    ufgMeal.DataGrid.Col = 5
    ufgMeal.DataGrid.Row = ufgMeal.SelectionRow
    ufgMeal.DataGrid.SetFocus
    
    '检查套餐名称是否为空
    If ufgMeal.Text(ufgMeal.SelectionRow, gstrAntibodyMeal_套餐名称) = "" Then
        MsgBoxD Me, "未能通过验证，原因是套餐名称不能为空！", vbExclamation, Me.Caption
        ufgMeal.LocateRow ufgMeal.SelectionRow
        ufgMeal.DataGrid.EditCell
        Exit Sub
    Else
        '检查套餐名称是否重复
        If Not CheckMealName Then Exit Sub
    End If
    
    Call Menu_PatholMeal_SaveMeal
    Call Menu_PatholMeal_SureSelected
    Call LoadMealClass

    ufgMeal.DataGrid.Cell(flexcpText, ufgMeal.SelectionRow, 5) = Format(ufgMeal.DataGrid.Cell(flexcpText, ufgMeal.SelectionRow, 5), "yyyy-mm-dd")

    If ufgMeal.AdoData.RecordCount > 0 Then Call ConfigMealLink(Val(ufgMeal.KeyValue(ufgMeal.SelectionRow)))
    
    mblnEdit = False
    mblnIsUpdate = False
    cboMealClass.Enabled = True
    lblMealClass.Enabled = True
    
    If ufgMealLink.DataGrid.Row > 0 Then ufgMealLink.LocateRow (1)
    ufgMealLink.Enabled = False
    ufgMealLink.DataGrid.Enabled = True
    ufgMealLink.DataGrid.BackColor = &H8000000F
    ufgMealLink.ReadOnly = False
    
    stbThis.Panels(2).Text = "当前套餐名称为：" & ufgMeal.DataGrid.Cell(flexcpText, ufgMeal.SelectionRow, 2)
End Sub

Private Function CheckMealName() As Boolean
'检查套餐名称是否重复
    Dim i As Integer
    
    CheckMealName = False
    For i = 1 To ufgMeal.GridRows - 1
        If Not ufgMeal.RowState(i) = TDataRowState.Del Then
            If Not mblnIsUpdate Then
                If (Not ufgMeal.RowState(i) = TDataRowState.Add) And (Not ufgMeal.RowHidden(i)) Then
                    If ufgMeal.Text(i, gstrAntibodyMeal_套餐名称) = ufgMeal.DataGrid.Cell(flexcpText, ufgMeal.SelectionRow, 2) Then
                        MsgBoxD Me, "抗体名称重复。", vbExclamation, Me.Caption

                        ufgMeal.LocateRow ufgMeal.SelectionRow
                        ufgMeal.DataGrid.EditCell
                        Exit Function
                    End If
                End If
            Else
                If Not ufgMeal.SelectionRow = i And (Not ufgMeal.RowHidden(i)) Then
                    If ufgMeal.Text(i, gstrAntibodyMeal_套餐名称) = ufgMeal.DataGrid.Cell(flexcpText, ufgMeal.SelectionRow, 2) Then
                        MsgBoxD Me, "抗体名称重复。", vbExclamation, Me.Caption

                        ufgMeal.LocateRow ufgMeal.SelectionRow
                        ufgMeal.DataGrid.EditCell
                        Exit Function
                    End If
                End If
            End If
        End If
    Next i
    
    CheckMealName = True
End Function

Private Sub Menu_PatholMeal_Cancel()
    Dim i As Integer
    
    '取消前得先让处于编辑状态的单元格失去焦点，否则不能恢复当前行单元格的数据信息
    ufgMeal.DataGrid.Col = 5
    ufgMeal.DataGrid.Row = ufgMeal.SelectionRow
    ufgMeal.DataGrid.SetFocus
    
    mblnEdit = False
    mblnIsUpdate = False
    cboMealClass.Enabled = True
    lblMealClass.Enabled = True
    mblnCurModifyState = True
    
    If ufgMeal.CurKeyValue = "" Then ufgMeal.DelCurRow False
    
    ufgMealLink.HeadCheckValue = False
    
    If ufgMeal.AdoData.RecordCount > 0 Then
        Call ufgMeal.RestoreCurRowText
        Call ufgMeal.LocateRow(ufgMeal.SelectionRow)
        Call ConfigMealLink(Val(ufgMeal.KeyValue(ufgMeal.SelectionRow)))
    End If
    
    For i = 1 To ufgMeal.AdoData.RecordCount - 1
        ufgMeal.DataGrid.TextMatrix(i, 5) = Format(ufgMeal.DataGrid.Cell(flexcpText, i, 5), "yyyy-mm-dd")
    Next
    If ufgMealLink.DataGrid.Row > 0 Then ufgMealLink.LocateRow (1)
    ufgMealLink.Enabled = False
    ufgMealLink.DataGrid.Enabled = True
    ufgMealLink.DataGrid.BackColor = &H8000000F
    ufgMealLink.ReadOnly = False
    
    stbThis.Panels(2).Text = "当前套餐名称为：" & ufgMeal.DataGrid.Cell(flexcpText, ufgMeal.SelectionRow, 2)
End Sub

Private Sub cbrMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible = True Then Bottom = stbThis.Height
End Sub

Private Sub cbrMain_Update(ByVal control As XtremeCommandBars.ICommandBarControl)
    Dim blnHasRecord As Boolean
    
On Error GoTo errHandle
    blnHasRecord = ufgMeal.IsSelectionRow
    
    Select Case control.ID
        Case conMenu_PatholMeal_Save
            control.Enabled = (Not mblnCurModifyState) And blnHasRecord
            
        Case conMenu_PatholMeal_Cancel
            control.Enabled = (Not mblnCurModifyState) And blnHasRecord
            
        Case conMenu_PatholMeal_AddRecord
            control.Enabled = mblnCurModifyState Or (Not blnHasRecord)
            
        Case conMenu_PatholMeal_ModRecord
            control.Enabled = mblnCurModifyState And blnHasRecord

        Case conMenu_PatholMeal_DelRecord
            control.Enabled = mblnCurModifyState And blnHasRecord

        Case conMenu_PatholMeal_UpRow
            control.Enabled = mblnEdit

        Case conMenu_PatholMeal_DownRow
            control.Enabled = mblnEdit

        Case conMenu_File_Exit
            
    End Select
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgMeal_OnBeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '在非新增和修改的情况下，不允许编辑单元格
    If Not mblnEdit Then Cancel = True
End Sub

Private Sub ufgMeal_OnBeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If Not mblnCurModifyState And OldRow <> NewRow Then Cancel = True
End Sub

Private Sub ufgMeal_OnColFormartChange()
 '保存列表参数
    zlDatabase.SetPara "抗体套餐列表配置", ufgMeal.GetColsString(ufgMeal), glngSys, G_LNG_PATHOLSYS_NUM
End Sub

Private Sub ufgMeal_OnMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'弹出右键菜单
On Error GoTo errHandle
    If Button = 2 Then
        Dim objPopup As CommandBar
        Dim objControl As CommandBarControl

        Set cbrMain.Icons = zlCommFun.GetPubIcons
        Set objPopup = cbrMain.Add("右键菜单", xtpBarPopup)
        With objPopup.Controls
            Set objControl = .Add(xtpControlButton, conMenu_PatholMeal_AddRecord, "新增套餐(&A)")
            Set objControl = .Add(xtpControlButton, conMenu_PatholMeal_ModRecord, "修改套餐(&M)")
            Set objControl = .Add(xtpControlButton, conMenu_PatholMeal_DelRecord, "删除套餐(&D)")
            
            Set objControl = .Add(xtpControlButton, conMenu_PatholMeal_Save, "保存(&S)"): objControl.IconId = 3091
            Set objControl = .Add(xtpControlButton, conMenu_PatholMeal_Cancel, "撤销(&R)"): objControl.IconId = 3565
        End With
        objPopup.ShowPopup
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgMealLink_OnCheckChanging(ByVal Row As Long, ByVal Col As Long, AllowChange As Boolean)
    If Not ufgMealLink.ReadOnly Then AllowChange = False
End Sub

Private Sub ufgMealLink_OnColFormartChange()
    zlDatabase.SetPara "套餐明细列表配置", ufgMealLink.GetColsString(ufgMealLink), glngSys, G_LNG_PATHOLSYS_NUM
End Sub

Private Sub InitFace()
'初始化功能界面
    Dim Pane1 As Pane, Pane2 As Pane

    With dkpMeal
        .CloseAll
        .SetCommandBars cbrMain
        .Options.HideClient = True
        .Options.UseSplitterTracker = False '实时拖动
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = True
    End With

    Set Pane1 = dkpMeal.CreatePane(1, 0, Round(Me.Width / 2), DockLeftOf)
    Pane1.Title = "套餐记录"
    Pane1.Handle = picDatas.hWnd
    Pane1.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    Pane1.MinTrackSize.Width = 50

    Set Pane2 = dkpMeal.CreatePane(2, 0, Round(Me.Width / 2), DockRightOf)
    Pane2.Title = "抗体明细"
    Pane2.Handle = picMealLink.hWnd
    Pane2.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    Pane2.MinTrackSize.Width = 50
End Sub

Private Sub LoadMealData()
'载入套餐数据
    Dim i As Integer
    Dim strSql As String
    Dim rsMeal As ADODB.Recordset
    
    strSql = "select 套餐ID,套餐名称,套餐类别,套餐说明,创建时间,创建人 from 病理套餐信息"
      
    Set ufgMeal.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    
    Call ufgMeal.RefreshData
    For i = 1 To ufgMeal.AdoData.RecordCount
        ufgMeal.DataGrid.Cell(flexcpText, i, 5) = Format(ufgMeal.DataGrid.Cell(flexcpText, i, 5), "yyyy-mm-dd")
    Next
End Sub

Private Sub LoadAntibodyData()
'读取抗体数据（排除禁用的抗体）
    Dim strSql As String
    Dim rsAntibody As ADODB.Recordset

    strSql = "select '' as 关联ID,抗体ID,抗体名称,克隆性,作用对象,理化性质,应用情况,备注, '' as 抗体顺序 from 病理抗体信息 where 使用状态 = 1 order by 抗体ID"
      
    Set ufgMealLink.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    
    Call ufgMealLink.RefreshData
End Sub

Private Sub cboMealClass_Click()
On Error GoTo errHandle
    '过滤套餐信息
    
    If cboMealClass.Text = "" Then
        ufgMeal.AdoData.Filter = ""
    Else
        ufgMeal.AdoData.Filter = "套餐类别='" & cboMealClass.Text & "'"
    End If
    
    Call ufgMeal.RefreshData
    
    If ufgMeal.DataGrid.Row <= 0 Then Exit Sub
    Call ConfigMealLink(Val(ufgMeal.KeyValue(ufgMeal.DataGrid.Row)))
    
    stbThis.Panels(2).Text = "当前套餐名称为：" & ufgMeal.DataGrid.Cell(flexcpText, ufgMeal.SelectionRow, 2)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub chkRowFilter_Click()
On Error GoTo errHandle
    If chkRowFilter.value = 1 Then
        Call ufgMealLink.ShowCheckRows
    Else
        Call ufgMealLink.ShowAllRows
    End If
Exit Sub
errHandle:
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

Private Sub Menu_PatholMeal_DelRecord()
'删除套餐
On Error GoTo errHandle
    If Not ufgMeal.IsSelectionRow Then
        Call MsgBoxD(Me, "请选择需要进行删除的套餐记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If MsgBoxD(Me, "确认要删除该套餐吗？", vbYesNo, Me.Caption) = vbNo Then Exit Sub
    
    Call ufgMeal.DelRow(ufgMeal.SelectionRow, False)
    
    Call SaveMealData(True)
    
    If ufgMeal.ShowingDataRowCount <= 0 Then
        '当没有套餐数据时，清除套餐关联
        Call ufgMealLink.ClearCellCheck(ufgMealLink.GetColIndexWithRowCheck)
        
        Call ReinitMealLinkData
    Else
        '配置下一套餐关联
        Call ConfigMealLink(Val(ufgMeal.KeyValue(ufgMeal.SelectionRow)))
    End If
    
    If ufgMeal.IsSelectionRow Then
        If ufgMeal.IsEmptyKey(ufgMeal.SelectionRow) Then
            Call ConfigButState(False)
        End If
    End If
    
    Call LoadMealClass
    
    stbThis.Panels(2).Text = "当前套餐名称为：" & ufgMeal.DataGrid.Cell(flexcpText, ufgMeal.SelectionRow, 2)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_PatholMeal_DownRow()
On Error GoTo errHandle
    Call ufgMealLink.MoveDown(ufgMealLink.SelectionRow)
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub SaveMealData(Optional ByVal blnIsSaveOnlyDel As Boolean = False)
'保存套餐数据
    Dim i As Long
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim dtServicesTime As Date
    
    For i = 1 To ufgMeal.GridRows - 1
        Select Case ufgMeal.RowState(i)
            Case IIf(blnIsSaveOnlyDel, -1, TDataRowState.Add)
                dtServicesTime = zlDatabase.Currentdate
                
                '添加新的套餐
                strSql = "select Zl_病理套餐_新增([1],[2],[3],[4],[5]) as 返回值 from dual"
                
                Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, _
                                                    ufgMeal.Text(i, gstrAntibodyMeal_套餐名称), _
                                                    ufgMeal.Text(i, gstrAntibodyMeal_套餐类别), _
                                                    ufgMeal.Text(i, gstrAntibodyMeal_套餐说明), _
                                                    CDate(Format(dtServicesTime, "yyyy-mm-dd")), _
                                                    UserInfo.姓名)
                
                If rsData.RecordCount <= 0 Then
                    Call err.Raise(0, "SaveMealData", "未成功获取新增后的套餐ID,处理失败。")
                    Exit Sub
                End If
                
                ufgMeal.Text(i, gstrAntibodyMeal_套餐ID) = rsData!返回值
                ufgMeal.Text(i, gstrAntibodyMeal_创建人) = UserInfo.姓名
                ufgMeal.Text(i, gstrAntibodyMeal_创建时间) = dtServicesTime
                
                ufgMeal.SyncRowDataToAdo i
            Case TDataRowState.Del
                '删除套餐(会级联删除关联数据)
                
                strSql = "Zl_病理套餐_删除(" & Val(ufgMeal.KeyValue(i)) & ")"
                Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
                
                If ufgMeal.ShowingDataRowCount <= 0 Then
                    '清除选择
                    Call ufgMealLink.ClearCellCheck(ufgMealLink.GetColIndexWithRowCheck)
                    '将关联ID设置为空
                    Call ReinitMealLinkData
                End If
                
                ufgMeal.SyncRowDataToAdo i
            Case IIf(blnIsSaveOnlyDel, -1, TDataRowState.Update)
                '更新套餐
                strSql = "Zl_病理套餐_更新(" & Val(ufgMeal.KeyValue(i)) & ",'" & _
                                            ufgMeal.Text(i, gstrAntibodyMeal_套餐名称) & "','" & _
                                            ufgMeal.Text(i, gstrAntibodyMeal_套餐类别) & "','" & _
                                            ufgMeal.Text(i, gstrAntibodyMeal_套餐说明) & "')"
                
                Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
                
                ufgMeal.SyncRowDataToAdo i
        End Select
        
        '更新行状态
        ufgMeal.RowState(i) = TDataRowState.Normal
    Next i
  
End Sub

Private Sub SaveMealLinkData(ByVal lngMealId As Long)
'保存套餐对应的抗体数据
    Dim i As Long
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim lngMealLinkId As Long
    Dim lngAntibodyOrder As Long
    
    lngAntibodyOrder = 0
    
    '删除套餐关联
    strSql = "Zl_病理套餐关联_删除(" & lngMealId & ")"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
        
    For i = 1 To ufgMealLink.GridRows - 1

        If ufgMealLink.GetRowCheck(i) Then
            '判断是否有关联ID,如果没有，则是增加关联，如果有则不做处理
            strSql = "select Zl_病理套餐关联_新增([1],[2],[3]) as 返回值 from dual"
            Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngMealId, Val(ufgMealLink.KeyValue(i)), lngAntibodyOrder)
            
            If rsData.RecordCount <= 0 Then
                Call err.Raise(0, "SaveMealLinkData", "未成功获取新增后的套餐抗体关联ID,处理失败。")
                Exit Sub
            End If
            
            '设置关联ID
            ufgMealLink.Text(i, gstrAntibodyMealLink_关联ID) = rsData!返回值
            
            lngAntibodyOrder = lngAntibodyOrder + 1
        Else
            ufgMealLink.Text(i, gstrAntibodyMealLink_关联ID) = ""
        End If
        
'        If ufgMealLink.RowState(i) = TDataRowState.Update Then
'            lngMealLinkId = Val(ufgMealLink.Text(i, gstrAntibodyMealLink_关联ID))
'
'            If ufgMealLink.GetRowChecked(i) Then
'                '判断是否有关联ID,如果没有，则是增加关联，如果有则不做处理
'                If lngMealLinkId <= 0 Then
'                    strSQL = "select Zl_病理套餐关联_新增([1],[2]) as 返回值 from dual"
'                    Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngMealId, Val(ufgMealLink.GetKeyValue(i)))
'
'                    If rsData.RecordCount <= 0 Then
'                        Call err.Raise(0, "SaveMealLinkData", "未成功获取新增后的套餐抗体关联ID,处理失败。")
'                        Exit Sub
'                    End If
'
'                    '设置关联ID
'                    Call ufgMealLink.SetText(i, gstrAntibodyMealLink_关联ID, rsData!返回值)
'                End If
'            Else
'                '判读是否有关联ID,如果有，则删除关联，如果没有则不做处理
'                If lngMealLinkId > 0 Then
'                    strSQL = "Zl_病理套餐关联_删除1(" & lngMealLinkId & ")"
'                    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
'
'                    '清除套餐关联ID
'                    Call ufgMealLink.SetText(i, gstrAntibodyMealLink_关联ID, "")
'                End If
'            End If
'
'            '恢复行状态
'            ufgMealLink.RowState(i) = TDataRowState.Normal
'        End If
    Next i
End Sub

Private Sub Menu_PatholMeal_AddMeal()
    Dim i As Integer
    
On Error GoTo errHandle
    mblnEdit = True
    mblnIsUpdate = False
    cboMealClass.Enabled = False
    lblMealClass.Enabled = False
    ufgMealLink.ReadOnly = True
    
    For i = 1 To ufgMealLink.DataGrid.Rows - 1
        ufgMealLink.SetRowCheck i, False
    Next
    ufgMealLink.Enabled = True
    ufgMealLink.DataGrid.BackColor = vbWhite
    
    ufgMeal.Editable = flexEDKbd
    ufgMeal.NewRow
    ufgMeal.DataGrid.EditCell

    mblnCurModifyState = False
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub
            
Private Sub Menu_PatholMeal_ModMeal()
On Error GoTo errHandle
    mblnEdit = True
    mblnIsUpdate = True
    cboMealClass.Enabled = False
    lblMealClass.Enabled = False
    ufgMealLink.Enabled = True
    ufgMealLink.DataGrid.BackColor = vbWhite
    ufgMealLink.ReadOnly = True
    
    ufgMeal.Editable = flexEDKbd
    ufgMeal.LocateRow ufgMeal.SelectionRow
    ufgMeal.DataGrid.EditCell

    mblnCurModifyState = False
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_PatholMeal_SaveMeal()
'保存套餐信息
On Error GoTo errHandle
    Dim blnValid As Boolean
    
    blnValid = Not ufgMeal.IsErrColorWithList
    If Not blnValid Then
        Call MsgBoxD(Me, "检测到套餐列表中存在无效数据，请确认相关数据是否正确完整的录入，“红色”标记的单元格为必录数据。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    '保存套餐信息
    Call SaveMealData
    
    If ufgMeal.IsSelectionRow Then
        If Not ufgMeal.IsEmptyKey(ufgMeal.SelectionRow) Then
            Call ConfigButState(True)
        End If
    End If
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_PatholMeal_SureSelected()
'确认关联选择
On Error GoTo errHandle

    If Not ufgMeal.IsSelectionRow Then
        Call MsgBoxD(Me, "请选择所对应的套餐信息。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    Call SaveMealLinkData(Val(ufgMeal.KeyValue(ufgMeal.SelectionRow)))
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_PatholMeal_UpRow()
On Error GoTo errHandle
    Call ufgMealLink.MoveUp(ufgMealLink.SelectionRow)

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Load()
On Error GoTo errHandle
'    InitDebugObject 1294, Me, "zlhis", "his"
    Call InitCommandBars

    Call RestoreWinState(Me, App.ProductName)
    
    Call InitFace
    
    '初始化列表
    Call InitMealList
    Call InitMealLinkList
    
    '载入数据
    Call LoadMealData
    Call LoadAntibodyData
    Call LoadMealClass
    
    '如果选择了第一行，则自动加载配置数据
    If ufgMeal.IsSelectionRow And Trim(ufgMeal.KeyValue(ufgMeal.SelectionRow)) <> "" Then
        Call ConfigMealLink(Val(ufgMeal.KeyValue(ufgMeal.SelectionRow)))
    End If
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub LoadMealClass()
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim strMealClass As String
    
    strSql = "select distinct 套餐类别 from 病理套餐信息"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    
    cboMealClass.Clear
    cboMealClass.AddItem ""
    
    While Not rsData.EOF
        If Nvl(rsData!套餐类别) <> "" Then
            cboMealClass.AddItem Nvl(rsData!套餐类别)
            strMealClass = strMealClass & "|" & Nvl(rsData!套餐类别)
        End If
        rsData.MoveNext
    Wend
    
    ufgMeal.ComboxListFormat(ufgMeal.GetColIndex(gstrAntibodyMeal_套餐类别)) = strMealClass
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
        Set .Icons = zlCommFun.GetPubIcons                          '设置关联的图标控件
    End With

    Me.cbrMain.EnableCustomization False
    Me.cbrMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    
    '菜单定义
'Begin------------------------编辑菜单--------------------------------------默认可见
    cbrMain.ActiveMenuBar.Title = "菜单"
    
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)")
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_PatholMeal_Save, "保存(&S)")
        cbrControl.IconId = 3091
        Set cbrControl = .Add(xtpControlButton, conMenu_PatholMeal_Cancel, "撤销(&R)")
        cbrControl.IconId = 3565
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&Q)")
        cbrControl.BeginGroup = True
    End With
    
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)")
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_PatholMeal_AddRecord, "新增套餐(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_PatholMeal_ModRecord, "修改套餐(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_PatholMeal_DelRecord, "删除套餐(&D)")
        Set cbrControl = .Add(xtpControlButton, conMenu_PatholMeal_UpRow, "上移(&U)")
        cbrControl.IconId = 21802
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_PatholMeal_DownRow, "下移(&D)")
        cbrControl.IconId = 21801
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
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_PatholMeal_Save, "保存")
        cbrControl.IconId = 3091
        Set cbrControl = .Add(xtpControlButton, conMenu_PatholMeal_Cancel, "撤销")
        cbrControl.IconId = 3565
        Set cbrControl = .Add(xtpControlButton, conMenu_PatholMeal_AddRecord, "新增套餐")
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_PatholMeal_ModRecord, "修改套餐")
        Set cbrControl = .Add(xtpControlButton, conMenu_PatholMeal_DelRecord, "删除套餐")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_PatholMeal_UpRow, "上移")
        cbrControl.IconId = 21802
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_PatholMeal_DownRow, "下移")
        cbrControl.IconId = 21801
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
        cbrControl.BeginGroup = True
    End With
    
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub picDatas_Resize()
'调整套餐界面布局
On Error Resume Next
    framMeals.Left = 120
    framMeals.Top = 120
    framMeals.Width = picDatas.Width - 120
    framMeals.Height = picDatas.Height - 240
    
    ufgMeal.Left = 120
    ufgMeal.Top = lblMealClass.Top + 360
    ufgMeal.Width = framMeals.Width - 240
    ufgMeal.Height = framMeals.Height - lblMealClass.Height - lblMealClass.Top - 240
End Sub

Private Sub ConfigMealLink(ByVal lngMealId As Long)
'配置套餐关联(查询该套餐所属的抗体，然后再抗体列表的对应checked上设置为True)
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    strSql = "select ID, 抗体ID,抗体顺序  from 病理套餐关联 where 套餐ID=[1] order by 抗体顺序"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngMealId)
    
    '清除选择
    Call ufgMealLink.ClearCellCheck(ufgMealLink.GetColIndexWithRowCheck())
    
    '将关联ID设置为空
    Call ReinitMealLinkData
        
    If rsData.RecordCount <= 0 Then Exit Sub

    Do While Not rsData.EOF
        Call SetMealAntibodyLink(Val(Nvl(rsData!抗体ID)), Val(Nvl(rsData!ID)), Val(Nvl(rsData!抗体顺序)))
        
        rsData.MoveNext
    Loop
    
'    Call ufgMealLink.Sort(ufgMealLink.vfgHelper.GetColumnIndex(ufgMealLink.vfgHelper.CheckColName))
    Call ufgMealLink.Sort(ufgMealLink.GetColIndex(gstrAntibodyMealLink_抗体顺序))
End Sub

Private Sub SetMealAntibodyLink(ByVal lngAntibodyId As Long, ByVal lngMealLinkId As Long, ByVal lngAntibodyOrder As Long)
'设置套餐抗体关联
    Dim i As Long
    
    ufgMealLink.ReadOnly = True
    For i = 1 To ufgMealLink.GridRows - 1
        If Val(ufgMealLink.KeyValue(i)) = lngAntibodyId Then
            ufgMealLink.Text(i, gstrAntibodyMealLink_关联ID) = lngMealLinkId
            ufgMealLink.Text(i, gstrAntibodyMealLink_抗体顺序) = String(4 - Len("" & lngAntibodyOrder & ""), "0") & lngAntibodyOrder
            Call ufgMealLink.SetRowCheck(i, True)
            ufgMealLink.ReadOnly = False
            Exit Sub
        End If
    Next i
    
End Sub

Private Sub ReinitMealLinkData()
'重新初始化套餐关联数据
    Dim i As Long

    For i = 1 To ufgMealLink.GridRows - 1
        ufgMealLink.Text(i, gstrAntibodyMealLink_关联ID) = ""
        ufgMealLink.Text(i, gstrAntibodyMealLink_抗体顺序) = "9999"
        ufgMealLink.RowState(i) = TDataRowState.Normal
    Next i
End Sub

Private Sub picMealLink_Resize()
'调整套餐明细界面布局
On Error Resume Next
    framAntibody.Left = 120
    framAntibody.Top = 120
    framAntibody.Width = picMealLink.Width - 240
    framAntibody.Height = picMealLink.Height - 240
    
    ufgMealLink.Left = 120
    ufgMealLink.Top = chkRowFilter.Top + 360
    ufgMealLink.Width = framAntibody.Width - 240
    ufgMealLink.Height = framAntibody.Height - lblFind.Height - lblFind.Top - 240
End Sub

Private Sub txtFind_Change()
On Error GoTo errHandle
    Dim lngFindIndex As Long
    
    If Trim(txtFind.Text) = "" Then Exit Sub
    
    lngFindIndex = ufgMealLink.FindRowIndex(txtFind.Text, gstrAntibodyMealLink_抗体名称)
    
    If lngFindIndex > 0 Then Call ufgMealLink.LocateRow(lngFindIndex)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub txtFind_GotFocus()
On Error Resume Next
    txtFind.SelStart = 0
    txtFind.SelLength = Len(txtFind.Text)
End Sub

Private Sub ufgMeal_OnAfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim iCol As Long
    Dim i As Long
    
    If ufgMeal.IsNullRow(Row) Then
        ufgMeal.RowState(Row) = TDataRowState.Normal
        Call ufgMeal.SetRowColor(Row, ufgMeal.BackColor)
        
        Exit Sub
    End If
        
    '如果未录入标本名称，则显示淡红色
    iCol = ufgMeal.GetColIndex(gstrAntibodyMeal_套餐名称)
    
    ufgMeal.CellColor(Row, iCol) = IIf(ufgMeal.Text(Row, gstrAntibodyMeal_套餐名称) = "", ufgMeal.ErrCellColor, ufgMeal.BackColor)
End Sub

Private Sub ConfigButState(ByVal blnEnable As Boolean)
    mblnCurModifyState = blnEnable
End Sub

Private Sub ufgMeal_OnClick()
On Error GoTo errHandle
    If Not mblnCurModifyState Then
        ufgMeal.Editable = flexEDKbd
        ufgMeal.DataGrid.EditCell
        
        Exit Sub
    End If
    
    If ufgMeal.ShowingDataRowCount <= 0 Then
        Call ufgMealLink.ClearCellCheck(ufgMealLink.GetColIndexWithRowCheck)
        Call ReinitMealLinkData
        Call ConfigButState(False)
        
        Exit Sub
    End If

    If ufgMeal.MouseRowIndex <= 0 Then Exit Sub
    
    If Trim(ufgMeal.KeyValue(ufgMeal.MouseRowIndex)) = "" Then
        Call ufgMealLink.ClearCellCheck(ufgMealLink.GetColIndexWithRowCheck)
        Call ReinitMealLinkData
        Call ConfigButState(False)
        
        Exit Sub
    End If

    '配置关联
    Call ConfigMealLink(Val(ufgMeal.KeyValue(ufgMeal.MouseRowIndex)))
    Call ConfigButState(True)
    
    If chkRowFilter.value = 1 Then
        Call ufgMealLink.ShowCheckRows
    End If
    stbThis.Panels(2).Text = "当前套餐名称为：" & ufgMeal.DataGrid.Cell(flexcpText, ufgMeal.SelectionRow, 2)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgMealLink_OnClick()
On Error GoTo errHandle
    stbThis.Panels(2).Text = "当前抗体名称为：" & ufgMealLink.DataGrid.Cell(flexcpText, ufgMealLink.SelectionRow, 3)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgMeal_OnStartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Row > 0 Then
        If ufgMeal.Text(Row, gstrAntibodyMeal_套餐类别) = "" And cboMealClass.Text <> "" Then
            ufgMeal.Text(Row, gstrAntibodyMeal_套餐类别) = cboMealClass.Text
        End If
    End If
End Sub

Private Sub ufgMealLink_OnAfterEdit(ByVal Row As Long, ByVal Col As Long)
    '将行修改为更新状态
    If ufgMealLink.IsComboboxCol(Col) Then
        ufgMealLink.RowState(Row) = TDataRowState.Update
    End If
End Sub


Private Sub ufgMealLink_OnMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'弹出右键菜单
On Error GoTo errHandle
    If Button = 2 Then
        Dim objPopup As CommandBar
        Dim objControl As CommandBarControl

        Set cbrMain.Icons = zlCommFun.GetPubIcons
        Set objPopup = cbrMain.Add("右键菜单", xtpBarPopup)
        With objPopup.Controls
            Set objControl = .Add(xtpControlButton, conMenu_PatholMeal_UpRow, "上移(&U)"): objControl.IconId = 21802
            Set objControl = .Add(xtpControlButton, conMenu_PatholMeal_DownRow, "下移(&D)"): objControl.IconId = 21801
        End With
        objPopup.ShowPopup
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub
