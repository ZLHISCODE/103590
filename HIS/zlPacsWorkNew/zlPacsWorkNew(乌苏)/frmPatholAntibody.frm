VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPatholAntibody 
   Caption         =   "抗体维护"
   ClientHeight    =   7755
   ClientLeft      =   75
   ClientTop       =   405
   ClientWidth     =   10155
   Icon            =   "frmPatholAntibody.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7755
   ScaleWidth      =   10155
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picFeedback 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   120
      ScaleHeight     =   2985
      ScaleWidth      =   9705
      TabIndex        =   5
      Top             =   4200
      Width           =   9735
      Begin zl9PACSWork.ucFlexGrid ufgFeedback 
         Height          =   2175
         Left            =   0
         TabIndex        =   6
         Top             =   720
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   3836
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
      Begin VB.Label labSubTitle 
         Caption         =   "抗体反馈记录"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.Line linFlag 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         X1              =   0
         X2              =   9840
         Y1              =   360
         Y2              =   360
      End
   End
   Begin VB.PictureBox picData 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   120
      ScaleHeight     =   3585
      ScaleWidth      =   9705
      TabIndex        =   0
      Top             =   480
      Width           =   9735
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   1080
         TabIndex        =   1
         ToolTipText     =   "根据抗体名称进行快速定位。"
         Top             =   120
         Width           =   1695
      End
      Begin TabDlg.SSTab tsFilter 
         Height          =   330
         Left            =   0
         TabIndex        =   2
         Top             =   600
         Width           =   6525
         _ExtentX        =   11509
         _ExtentY        =   582
         _Version        =   393216
         Tabs            =   4
         Tab             =   3
         TabsPerRow      =   4
         TabHeight       =   520
         TabMaxWidth     =   2822
         WordWrap        =   0   'False
         TabCaption(0)   =   "所有抗体(0)"
         TabPicture(0)   =   "frmPatholAntibody.frx":179A
         Tab(0).ControlEnabled=   0   'False
         Tab(0).ControlCount=   0
         TabCaption(1)   =   "过期抗体(0)"
         TabPicture(1)   =   "frmPatholAntibody.frx":17B6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).ControlCount=   0
         TabCaption(2)   =   "低量抗体(0)"
         TabPicture(2)   =   "frmPatholAntibody.frx":17D2
         Tab(2).ControlEnabled=   0   'False
         Tab(2).ControlCount=   0
         TabCaption(3)   =   "禁用抗体(0)"
         TabPicture(3)   =   "frmPatholAntibody.frx":17EE
         Tab(3).ControlEnabled=   -1  'True
         Tab(3).ControlCount=   0
      End
      Begin zl9PACSWork.ucFlexGrid ufgData 
         Height          =   2655
         Left            =   0
         TabIndex        =   3
         Top             =   960
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   4683
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
      Begin VB.Label labFind 
         AutoSize        =   -1  'True
         Caption         =   "快速查找："
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "根据抗体名称进行快速定位"
         Top             =   120
         Width           =   900
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   7
      Top             =   7395
      Width           =   10155
      _ExtentX        =   17912
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPatholAntibody.frx":180A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11033
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
      Left            =   1200
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmPatholAntibody.frx":209E
      Left            =   480
      Top             =   120
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmPatholAntibody"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngAntibodyLowCount As Long
Private mstrPrivs As String

Private mblnDataModifyState As Boolean
Private mblnFeedbackModifyState As Boolean

'菜单类型枚举定义
Private Enum TMenuType
    mtAntibodyAdd = 1
    mtAntibodyMod = 2
    mtAntibodyDel = 3
    mtAntibodyStatus = 4
    mtFeedbackAdd = 5
    mtFeedbackMod = 6
    mtFeedbackDel = 7
End Enum

'抗体数据的显示类型
Private Enum TAntibodyDataShowType
    stAll = 0   '所有抗体
    stOverdue = 1 '过期抗体
    stLow = 2 '低量抗体
    stDisable = 3 '禁用抗体
End Enum

Public Sub ShowAntibodyManageWind(ByVal strPrivs As String, Optional owner As Form = Nothing)
'显示抗体管理窗口
    mstrPrivs = strPrivs
    
    Call ConfigPopedom
    
    Call Me.Show(1, owner)
End Sub

Private Sub ConfigPopedom()
'配置权限
    mblnDataModifyState = InStr(mstrPrivs, "抗体管理") > 0
    mblnFeedbackModifyState = InStr(mstrPrivs, "抗体反馈") > 0
End Sub

Private Sub InitAntibodyList()
'初始化抗体显示列
    Dim strTemp As String

     '判断数据库参数表是否有数据 有则读取数据库参数  没有则加载默认
    strTemp = zlDatabase.GetPara("抗体信息列表配置", glngSys, G_LNG_PATHOLSYS_NUM, "")
     
    If strTemp = "" Then
        ufgData.ColNames = gstrAntibodyCols
    Else
        ufgData.ColNames = strTemp
    End If
    
     '设置行数
    ufgData.GridRows = glngStandardRowCount
    '设置行高
    ufgData.RowHeightMin = glngStandardRowHeight
    ufgData.DefaultColNames = gstrAntibodyCols
    ufgData.ColConvertFormat = gstrAntibodyConvertFormat
    ufgData.IsShowPopupMenu = False
End Sub

Private Sub InitFeedbackList()
'初始化抗体反馈显示列
    Dim strTemp As String
    
     '判断数据库参数表是否有数据 有则读取数据库参数  没有则加载默认
    strTemp = zlDatabase.GetPara("抗体反馈列表配置", glngSys, G_LNG_PATHOLSYS_NUM, "")
     
    If strTemp = "" Then
        ufgFeedback.ColNames = gstrAntibodyFeedbackCols
    Else
        ufgFeedback.ColNames = strTemp
    End If
    
     '设置行数
    'ufgFeedback.GridRows = glngStandardRowCount
    
    '禁止右键弹出列表配置窗口
    ufgFeedback.IsEjectConfig = False
    
    '设置行高
    ufgFeedback.RowHeightMin = glngStandardRowHeight
    
    ufgFeedback.DefaultColNames = gstrAntibodyFeedbackCols
    ufgFeedback.ColConvertFormat = gstrAntibodyFeedbackConvertFormat
    ufgFeedback.IsShowPopupMenu = False
End Sub

Private Sub cbrMain_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
    On Error GoTo ErrorHand
    
    Select Case control.ID
        Case TMenuType.mtAntibodyAdd                '新增抗体
            Menu_Antibody_Add
            
        Case TMenuType.mtAntibodyMod                '修改抗体
            Menu_Antibody_Mod
            
        Case TMenuType.mtAntibodyDel                '删除抗体
            Menu_Antibody_Del
            
        Case TMenuType.mtAntibodyStatus
            If InStr(control.Caption, "启用抗体") > 0 Then
                Menu_Antibody_Enable                '启用抗体
                control.Caption = "禁用抗体"
            Else
                Menu_Antibody_Disable               '禁用抗体
                control.Caption = "启用抗体"
            End If
        
        Case TMenuType.mtFeedbackAdd                '新增反馈
            Menu_Feedback_Add
            
        Case TMenuType.mtFeedbackMod                '修改反馈
            Menu_Feedback_Mod
            
        Case TMenuType.mtFeedbackDel                '删除反馈
            Menu_Feedback_Del
            
        Case conMenu_File_Exit                      '退出
            Call Menu_File_Exit
        
        '---------------------------查看----------------
        Case conMenu_View_ToolBar_Button            '工具栏
            Call Menu_View_ToolBar_Button_click(control)

        Case conMenu_View_ToolBar_Text              '按钮文字
            Call Menu_View_ToolBar_Text_click(control)

        Case conMenu_View_StatusBar                 '状态栏
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

Private Sub cbrMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible = True Then Bottom = stbThis.Height
End Sub

Private Sub cbrMain_Update(ByVal control As XtremeCommandBars.ICommandBarControl)
    Dim blnHasDataRecord As Boolean
    Dim blnHasFeedBackRecord As Boolean
    
    On Error GoTo ErrorHand
    
    '如果没有记录或者没有选中行，菜单和工具栏则不可用
    blnHasDataRecord = ufgData.IsSelectionRow
    blnHasFeedBackRecord = ufgFeedback.IsSelectionRow
    
    Select Case control.ID
        Case TMenuType.mtAntibodyAdd
            control.Enabled = mblnDataModifyState
            
        Case TMenuType.mtAntibodyMod
            control.Enabled = mblnDataModifyState And blnHasDataRecord
        
        Case TMenuType.mtAntibodyDel
            control.Enabled = mblnDataModifyState And blnHasDataRecord
            
        Case TMenuType.mtAntibodyStatus
            If control.Parent.type = xtpControlPopup Then
                control.Caption = IIf(ufgData.CurText("使用状态") = "使用中", "禁用抗体(&I)", "启用抗体(&S)")
            Else
                control.Caption = IIf(ufgData.CurText("使用状态") = "使用中", "禁用抗体", "启用抗体")
            End If
            
            control.IconId = IIf(ufgData.CurText("使用状态") = "使用中", 3006, 3009)
            
            control.Enabled = mblnDataModifyState And blnHasDataRecord
            control.Enabled = Not control.Enabled
            control.Enabled = Not control.Enabled
            
        Case TMenuType.mtFeedbackAdd
            control.Enabled = mblnDataModifyState And blnHasDataRecord
            
        Case TMenuType.mtFeedbackMod
            control.Enabled = mblnFeedbackModifyState And blnHasFeedBackRecord
        
        Case TMenuType.mtFeedbackDel
            control.Enabled = mblnFeedbackModifyState And blnHasFeedBackRecord
            
    End Select
    Exit Sub
ErrorHand:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub PicData_Resize()
    On Error Resume Next
    
    tsFilter.Left = 120
    tsFilter.Top = 120
    
    labFind.Left = tsFilter.Left + tsFilter.Width + 240
    labFind.Top = 160
    
    txtFind.Left = labFind.Left + labFind.Width
    txtFind.Top = labFind.Top - 40
    
    ufgData.Left = 120
    ufgData.Top = tsFilter.Top + tsFilter.Height
    ufgData.Height = picData.Height - ufgData.Top - 60
    ufgData.Width = picData.Width - 240
End Sub

Private Sub picFeedback_Resize()
    On Error Resume Next
    
    linFlag.X1 = 0
    linFlag.X2 = picFeedback.Width
    linFlag.Y1 = 200
    linFlag.Y2 = 200
    
    labSubTitle.Top = 110
    
    ufgFeedback.Left = 120
    ufgFeedback.Top = labSubTitle.Top + labSubTitle.Height + 40
    ufgFeedback.Height = picFeedback.Height - ufgFeedback.Top
    ufgFeedback.Width = picFeedback.Width - 240
End Sub

Private Sub ufgData_OnColFormartChange()
'保存数据列表配置
    zlDatabase.SetPara "抗体信息列表配置", ufgData.GetColsString(ufgData), glngSys, G_LNG_PATHOLSYS_NUM
End Sub

Private Sub ufgData_OnColsNameReSet()
On Error GoTo errHandle
   Call LoadAntibodyData(0)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgData_OnMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'弹出右键菜单
On Error GoTo errHandle
    If Button = 2 Then
        Dim objPopup As CommandBar
        Dim objControl As CommandBarControl

        Set cbrMain.Icons = zlCommFun.GetPubIcons
        Set objPopup = cbrMain.Add("右键菜单", xtpBarPopup)
        With objPopup.Controls
            Set objControl = .Add(xtpControlButton, TMenuType.mtAntibodyAdd, "新增抗体(&A)"): objControl.IconId = 4112
            Set objControl = .Add(xtpControlButton, TMenuType.mtAntibodyMod, "修改抗体(&U)"): objControl.IconId = 4113
            Set objControl = .Add(xtpControlButton, TMenuType.mtAntibodyDel, "删除抗体(&D)"): objControl.IconId = 4114
            
            Set objControl = .Add(xtpControlButton, TMenuType.mtAntibodyStatus, "启用抗体(&S)"): objControl.IconId = 3009
            objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, TMenuType.mtFeedbackAdd, "新增反馈(&N)"): objControl.IconId = 4010
            objControl.BeginGroup = True
        End With
        objPopup.ShowPopup
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgData_OnSelChange()
    ufgData_OnClick
End Sub

Private Sub ufgFeedback_OnClick()
    mblnDataModifyState = False
    mblnFeedbackModifyState = True
End Sub

Private Sub ufgFeedback_OnColFormartChange()
'保存数据列表配置
    zlDatabase.SetPara "抗体反馈列表配置", ufgFeedback.GetColsString(ufgFeedback), glngSys, G_LNG_PATHOLSYS_NUM
End Sub

Private Sub LoadFeedbackData(ByVal lngAntibodyId As Long)
'载入抗体反馈数据
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    strSql = "select ID,参考病理号,实验类型,反馈意见,抗体评价,反馈医生,反馈时间 from 病理抗体反馈 where 抗体ID=[1] order by id"
    Set ufgFeedback.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, ufgData.KeyValue(lngAntibodyId))
    
    Call ufgFeedback.RefreshData
End Sub

'Private Function GetTestTypeValue(ByVal lngTestType As Long) As String
''获取实验类型的取值
'    GetTestTypeValue = ""
'
'    Select Case lngTestType
'        Case 0
'            GetTestTypeValue = "免疫组化"
'        Case 1
'            GetTestTypeValue = "特殊染色"
'        Case 2
'            GetTestTypeValue = "分子病理"
'        Case 3
'            GetTestTypeValue = "其他"
'    End Select
'End Function

'
'Private Sub FillFeedbackDataToList(rsData As ADODB.Recordset)
''添加抗体反馈记录
'
'    '如果超过当前所显示的行数，则退出
'    If rsData.AbsolutePosition >= vfgData.Rows Then Exit Sub
'
'    Call mclsVFGFeedback.SetTextWithColName(rsData.AbsolutePosition, gstrAntibodyFeedback_ID, Nvl(rsData!ID))
'    Call mclsVFGFeedback.SetTextWithColName(rsData.AbsolutePosition, gstrAntibodyFeedback_参考病理号, Nvl(rsData!参考病理号))
'    Call mclsVFGFeedback.SetTextWithColName(rsData.AbsolutePosition, gstrAntibodyFeedback_实验类型, GetTestTypeValue(Val(Nvl(rsData!实验类型))))
'    Call mclsVFGFeedback.SetTextWithColName(rsData.AbsolutePosition, gstrAntibodyFeedback_抗体评价, Nvl(rsData!抗体评价))
'    Call mclsVFGFeedback.SetTextWithColName(rsData.AbsolutePosition, gstrAntibodyFeedback_反馈意见, Nvl(rsData!反馈意见))
'    Call mclsVFGFeedback.SetTextWithColName(rsData.AbsolutePosition, gstrAntibodyFeedback_反馈医生, Nvl(rsData!反馈医生))
'    Call mclsVFGFeedback.SetTextWithColName(rsData.AbsolutePosition, gstrAntibodyFeedback_反馈时间, Format(Nvl(rsData!反馈时间), gstrFullDateTimeFormat))
'
'End Sub

Private Sub LoadAntibodyData(iShowType As TAntibodyDataShowType)
'读取抗体数据
    Dim strSql As String
    Dim rsAntibody As ADODB.Recordset
    
    Select Case iShowType
        Case TAntibodyDataShowType.stAll:
            strSql = "select 抗体ID,抗体名称,使用人份,已用人份,生产日期,有效期,过期日期,克隆性,作用对象,理化性质,应用情况,登记人,登记时间,使用状态,备注 from 病理抗体信息 order by 抗体ID"
        Case TAntibodyDataShowType.stOverdue:
            strSql = "select 抗体ID,抗体名称,使用人份,已用人份,生产日期,有效期,过期日期,克隆性,作用对象,理化性质,应用情况,登记人,登记时间,使用状态,备注 from 病理抗体信息 where 过期日期<sysdate order by 抗体ID"
        Case TAntibodyDataShowType.stLow:
            strSql = "select 抗体ID,抗体名称,使用人份,已用人份,生产日期,有效期,过期日期,克隆性,作用对象,理化性质,应用情况,登记人,登记时间,使用状态,备注 from 病理抗体信息 where 使用人份-已用人份 < " & mlngAntibodyLowCount & " order by 抗体ID"
        Case TAntibodyDataShowType.stDisable:
            strSql = "select 抗体ID,抗体名称,使用人份,已用人份,生产日期,有效期,过期日期,克隆性,作用对象,理化性质,应用情况,登记人,登记时间,使用状态,备注 from 病理抗体信息 where 使用状态=0 order by 抗体ID"
    End Select
    
    Set ufgData.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    
    Call ufgData.RefreshData
End Sub

'Private Sub FillAntibodyDataToList(rsData As ADODB.Recordset)
'    '添加抗体记录
'
'    '如果超过当前所显示的行数，则退出
'    If rsData.AbsolutePosition >= vfgData.Rows Then Exit Sub
'
'    Call mclsVFGAntibody.SetTextWithColName(rsData.AbsolutePosition, gstrAntibody_抗体ID, Nvl(rsData!抗体ID))
'    Call mclsVFGAntibody.SetTextWithColName(rsData.AbsolutePosition, gstrAntibody_抗体名称, Nvl(rsData!抗体名称))
'    Call mclsVFGAntibody.SetTextWithColName(rsData.AbsolutePosition, gstrAntibody_使用人份, Val(Nvl(rsData!使用人份)))
'    Call mclsVFGAntibody.SetTextWithColName(rsData.AbsolutePosition, gstrAntibody_已用人份, Val(Nvl(rsData!已用人份)))
'    Call mclsVFGAntibody.SetTextWithColName(rsData.AbsolutePosition, gstrAntibody_生产日期, Format(Nvl(rsData!生产日期), gstrDateFormat))
'    Call mclsVFGAntibody.SetTextWithColName(rsData.AbsolutePosition, gstrAntibody_有效期, Val(Nvl(rsData!有效期)) & "月")
'    Call mclsVFGAntibody.SetTextWithColName(rsData.AbsolutePosition, gstrAntibody_过期日期, Format(Nvl(rsData!过期日期), gstrDateFormat))
'    Call mclsVFGAntibody.SetTextWithColName(rsData.AbsolutePosition, gstrAntibody_克隆性, IIf(Val(Nvl(rsData!克隆性)) = 1, "单克隆", "多克隆"))
'    Call mclsVFGAntibody.SetTextWithColName(rsData.AbsolutePosition, gstrAntibody_作用对象, Nvl(rsData!作用对象))
'    Call mclsVFGAntibody.SetTextWithColName(rsData.AbsolutePosition, gstrAntibody_理化性质, Nvl(rsData!理化性质))
'    Call mclsVFGAntibody.SetTextWithColName(rsData.AbsolutePosition, gstrAntibody_应用情况, Nvl(rsData!应用情况))
'    Call mclsVFGAntibody.SetTextWithColName(rsData.AbsolutePosition, gstrAntibody_使用状态, IIf(Val(Nvl(rsData!使用状态)) = 0, "已禁用", "使用中"))
'    Call mclsVFGAntibody.SetTextWithColName(rsData.AbsolutePosition, gstrAntibody_登记人, Nvl(rsData!登记人))
'    Call mclsVFGAntibody.SetTextWithColName(rsData.AbsolutePosition, gstrAntibody_登记时间, Format(Nvl(rsData!登记时间), gstrFullDateTimeFormat))
'    Call mclsVFGAntibody.SetTextWithColName(rsData.AbsolutePosition, gstrAntibody_备注, Nvl(rsData!备注))
'
'End Sub

Private Sub RefreshAntibodyCount()
'刷新抗体数量显示
    Dim strSql As String
    Dim rsAntibodyCount As ADODB.Recordset
    
    strSql = "select " & _
             " (select count(1)  from 病理抗体信息) as 所有, " & _
             " (select count(1)  from 病理抗体信息 where 使用状态=0) as 禁用, " & _
             " (select count(1)  from 病理抗体信息 where (使用人份-已用人份) < " & mlngAntibodyLowCount & ") as 低量, " & _
             " (select count(1)  from 病理抗体信息 where 过期日期 < sysdate) as 过期 " & _
             " from dual"
             
    Set rsAntibodyCount = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    If rsAntibodyCount.RecordCount <= 0 Then
        tsFilter.TabCaption(0) = "所有抗体(0)"
        tsFilter.TabCaption(1) = "过期抗体(0)"
        tsFilter.TabCaption(2) = "低量抗体(0)"
        tsFilter.TabCaption(3) = "禁用抗体(0)"
    Else
        tsFilter.TabCaption(0) = "所有抗体(" & Val(Nvl(rsAntibodyCount!所有)) & ")"
        tsFilter.TabCaption(1) = "过期抗体(" & Val(Nvl(rsAntibodyCount!过期)) & ")"
        tsFilter.TabCaption(2) = "低量抗体(" & Val(Nvl(rsAntibodyCount!低量)) & ")"
        tsFilter.TabCaption(3) = "禁用抗体(" & Val(Nvl(rsAntibodyCount!禁用)) & ")"
    End If
End Sub

Private Sub Menu_Antibody_Add()
On Error GoTo errHandle
    Dim blnOk As Boolean
    
    blnOk = ShowUpdateWindow(True)
    
    If blnOk Then RefreshAntibodyCount
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Function CheckAntibodyIsUsed(lngAntibodyId As Long) As Boolean
'检查抗体是否已经使用
    Dim strSql As String
    Dim rsAntibody As ADODB.Recordset
    
    CheckAntibodyIsUsed = False
    
    strSql = "select 1 from 病理特检信息 where 抗体ID=[1]"
    Set rsAntibody = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngAntibodyId)
    
    If rsAntibody.RecordCount > 0 Then CheckAntibodyIsUsed = True
End Function

Private Sub DelAntibodyData(lngAntibodyId As Long)
'删除抗体数据记录
    Dim strSql As String
    
    strSql = "zl_病理抗体_删除(" & lngAntibodyId & ")"
    
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
End Sub

Private Function ShowUpdateFeedbackWindow(Optional ByVal isNew As Boolean = False) As Boolean
    Dim frmUpdate As New frmPatholAntibody_FeedbackUpdate
    On Error GoTo errFree
        If isNew Then
            ShowUpdateFeedbackWindow = frmUpdate.ShowAddAntibodyFeedback(Val(ufgData.KeyValue(ufgData.SelectionRow)), ufgFeedback, Me)
        Else
            ShowUpdateFeedbackWindow = frmUpdate.ShowUpdateAntibodyFeedback(ufgFeedback, Me)
        End If
errFree:
    Unload frmUpdate
    Set frmUpdate = Nothing
End Function

Private Sub Menu_Feedback_Add()
On Error GoTo errHandle
    If Not ufgData.IsSelectionRow Then
        Call MsgBoxD(Me, "请选择需要进行反馈的抗体记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    Call ShowUpdateFeedbackWindow(True)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Antibody_Del()
On Error GoTo errHandle
    '需要判断当前抗体是否已经使用过，已经使用的抗体不能执行删除
    If Not ufgData.IsSelectionRow Then
        Call MsgBoxD(Me, "请选择需要删除的抗体记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    Dim lngCurAntibodyId As Long

    lngCurAntibodyId = ufgData.KeyValue(ufgData.SelectionRow)
    
    If CheckAntibodyIsUsed(lngCurAntibodyId) Then
        Call MsgBoxD(Me, "抗体已被使用不能删除。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If MsgBoxD(Me, "确认要删除当前选择的抗体记录吗？", vbYesNo, Me.Caption) = vbNo Then Exit Sub
    
    Call DelAntibodyData(lngCurAntibodyId)
    
    '清空对应的抗体反馈记录
    Call ufgFeedback.ClearListData
    
    Call ufgData.DelRow(ufgData.SelectionRow, False)
    
    Call RefreshAntibodyCount
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub EnableOrDisableAntibody(lngAntibodyId As Long, blnIsEnable As Boolean)
'禁用或启用抗体
    Dim strSql As String
    
    strSql = "Zl_病理抗体_使用状态(" & lngAntibodyId & "," & IIf(blnIsEnable, 1, 0) & ")"
                                   
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
End Sub

Private Sub DelFeedbackData(lngFeedbackId As Long)
'删除抗体数据记录
    Dim strSql As String
    
    strSql = "Zl_病理抗体反馈_删除(" & lngFeedbackId & ")"
    
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
End Sub

Private Sub Menu_Feedback_Del()
'删除抗体反馈记录
On Error GoTo errHandle

    If Not ufgFeedback.IsSelectionRow Then
        Call MsgBoxD(Me, "请选择需要删除的反馈记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    Dim lngCurFeedbackId As Long
    
    lngCurFeedbackId = ufgFeedback.KeyValue(ufgFeedback.SelectionRow)
    
    If MsgBoxD(Me, "确认要删除当前选择的抗体记录吗？", vbYesNo, Me.Caption) = vbNo Then Exit Sub
    
    '删除抗体反馈数据
    Call DelFeedbackData(lngCurFeedbackId)

    Call ufgFeedback.DelRow(ufgFeedback.SelectionRow, False)
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Antibody_Disable()
On Error GoTo errHandle
    
    If Not ufgData.IsSelectionRow Then
        Call MsgBoxD(Me, "请选择需要禁用的抗体记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If

    If ufgData.Text(ufgData.SelectionRow, gstrAntibody_使用状态) = "已禁用" Then
        Call MsgBoxD(Me, "抗体已被禁用。", vbOKOnly, Me.Caption)
        Exit Sub
    End If

    Dim lngCurAntibodyId As Long

    lngCurAntibodyId = ufgData.KeyValue(ufgData.SelectionRow)
    Call EnableOrDisableAntibody(lngCurAntibodyId, False)

    '更新数据显示列表
    ufgData.Text(ufgData.SelectionRow, gstrAntibody_使用状态) = "已禁用"

    Call RefreshAntibodyCount
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Antibody_Enable()
On Error GoTo errHandle

    If Not ufgData.IsSelectionRow Then
        Call MsgBoxD(Me, "请选择需要启用的抗体记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If

    If ufgData.Text(ufgData.SelectionRow, gstrAntibody_使用状态) = "使用中" Then
        Call MsgBoxD(Me, "抗体已处于使用中。", vbOKOnly, Me.Caption)
        Exit Sub
    End If

    Dim lngCurAntibodyId As Long

    lngCurAntibodyId = ufgData.KeyValue(ufgData.SelectionRow)
    Call EnableOrDisableAntibody(lngCurAntibodyId, True)

    '更新数据显示列表
    ufgData.Text(ufgData.SelectionRow, gstrAntibody_使用状态) = "使用中"

    Call RefreshAntibodyCount
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Function ShowUpdateWindow(Optional ByVal isNew As Boolean = False) As Boolean
    Dim frmUpdate As New frmPatholAntibody_AntiUpdate
    On Error GoTo errFree
        If isNew Then
            ShowUpdateWindow = frmUpdate.ShowAddAntibodyWindow(ufgData, Me)
        Else
            ShowUpdateWindow = frmUpdate.ShowUpdateAntibodyWindow(ufgData, Me)
        End If
errFree:
    Unload frmUpdate
    Set frmUpdate = Nothing
    
End Function

Private Sub Menu_Antibody_Mod()
'抗体更新
On Error GoTo errHandle
    
    If Not ufgData.IsSelectionRow Then
        Call MsgBoxD(Me, "请选择需要更新的抗体记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If

    Dim blnOk As Boolean

    blnOk = ShowUpdateWindow(False)

    If blnOk Then RefreshAntibodyCount

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Feedback_Mod()
On Error GoTo errHandle

    If Not ufgFeedback.IsSelectionRow Then
        Call MsgBoxD(Me, "请选择需要更新的反馈记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If

    Call ShowUpdateFeedbackWindow(False)

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Initialize()
    mlngAntibodyLowCount = 3
'    tsFilter.Tab = 0
End Sub

Private Sub Form_Load()
On Error GoTo errHandle
'    InitDebugObject 1294, Me, "zlhis", "his"
    Call InitCommandBars
    
    Call RestoreWinState(Me, App.ProductName)
    
    Call InitFace
    '列表初始化
    Call InitAntibodyList
    Call InitFeedbackList
    
'    Call LoadAntibodyData(stAll)
    tsFilter.Tab = 0 '对该属性赋值时，会触发载入事件
    
    Call RefreshAntibodyCount
    
    '如果选择了第一行，则自动加载配置数据
    If ufgData.IsSelectionRow And Trim(ufgData.KeyValue(ufgData.SelectionRow)) <> "" Then
        Call LoadFeedbackData(Val(ufgData.SelectionRow))
    End If
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Exit Sub
End Sub

Private Sub InitFace()
'初始化功能界面
    Dim Pane1 As Pane, Pane2 As Pane

    With dkpMain
        .CloseAll
        .SetCommandBars cbrMain
        .Options.HideClient = True
        .Options.UseSplitterTracker = False '实时拖动
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = True
    End With

    Set Pane1 = dkpMain.CreatePane(1, 0, Round(Me.Height * 3 / 5), DockTopOf)
    Pane1.Title = "套餐记录"
    Pane1.Handle = picData.hWnd
    Pane1.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    Pane1.MinTrackSize.Height = 100
    
    Set Pane2 = dkpMain.CreatePane(2, 0, Round(Me.Height * 2 / 5), DockBottomOf, Pane1)
    Pane2.Title = "抗体明细"
    Pane2.Handle = picFeedback.hWnd
    Pane2.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    Pane2.MinTrackSize.Height = 100
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
        .VisualTheme = xtpThemeOffice2003                       '设置控件显示风格
        .EnableCustomization False                              '是否允许自定义设置
        Set .Icons = zlCommFun.GetPubIcons                      '设置关联的图标控件
    End With

    Me.cbrMain.EnableCustomization False
    Me.cbrMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    
    '菜单定义
'Begin------------------------编辑菜单--------------------------------------默认可见
    cbrMain.ActiveMenuBar.Title = "菜单"
    
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)")
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&Q)")
        cbrControl.BeginGroup = True
    End With
    
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)")
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtAntibodyAdd, "新增抗体(&A)"): cbrControl.IconId = 4112
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtAntibodyMod, "修改抗体(&U)"): cbrControl.IconId = 4113
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtAntibodyDel, "删除抗体(&D)"): cbrControl.IconId = 4114
        
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtAntibodyStatus, "启用抗体(&S)"): cbrControl.IconId = 3009
        cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtFeedbackAdd, "新增反馈(&N)"): cbrControl.IconId = 4010
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtFeedbackMod, "修改反馈(&U)"): cbrControl.IconId = 3003
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtFeedbackDel, "删除反馈(&C)"): cbrControl.IconId = 4008
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
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtAntibodyAdd, "新增抗体"): cbrControl.IconId = 4112
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtAntibodyMod, "修改抗体"): cbrControl.IconId = 4113
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtAntibodyDel, "删除抗体"): cbrControl.IconId = 4114
        
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtAntibodyStatus, "启用抗体"): cbrControl.IconId = 3009
        cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtFeedbackAdd, "新增反馈"): cbrControl.IconId = 4010
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtFeedbackMod, "修改反馈"): cbrControl.IconId = 3003
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtFeedbackDel, "删除反馈"): cbrControl.IconId = 4008
        
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

Private Sub tsFilter_Click(PreviousTab As Integer)
    Dim i As Integer
On Error GoTo errHandle
    If PreviousTab = tsFilter.Tab Then Exit Sub
    
    Select Case tsFilter.Tab
        Case 0: '所有抗体
            Call LoadAntibodyData(stAll)
        Case 1: '过期抗体
            Call LoadAntibodyData(stOverdue)
        Case 2: '低量抗体
            Call LoadAntibodyData(stLow)
        Case 3: '禁用抗体
            Call LoadAntibodyData(stDisable)
    End Select
    
    For i = 1 To ufgData.DataGrid.Rows - 1
        ufgData.Text(i, "有效期") = ufgData.Text(i, "有效期") & "月"
    Next
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub txtFind_Change()
On Error GoTo errHandle
    Dim lngFindIndex As Long
    
    If Trim(txtFind.Text) = "" Then Exit Sub
    
    lngFindIndex = ufgData.FindRowIndex(txtFind.Text, gstrAntibody_抗体名称)
    
    If lngFindIndex > 0 Then Call ufgData.LocateRow(lngFindIndex)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub txtFind_GotFocus()
On Error Resume Next
    txtFind.SelStart = 0
    txtFind.SelLength = Len(txtFind.Text)
End Sub

Private Sub ufgData_OnClick()
On Error GoTo errHandle
    mblnDataModifyState = True
    mblnFeedbackModifyState = False
    
    ufgFeedback.ClearListData
    If ufgData.GridRows <= 1 Then Exit Sub
    If Not ufgData.IsSelectionRow Then Exit Sub
    
    Call LoadFeedbackData(ufgData.SelectionRow)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgData_OnDblClick()
On Error GoTo errHandle
    Dim blnOk As Boolean
    
    If ufgData.GridRows <= 1 Then Exit Sub
    If ufgData.MouseRowIndex <= 0 Then Exit Sub
        

    blnOk = ShowUpdateWindow(False)
    If blnOk Then RefreshAntibodyCount
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgFeedback_OnDblClick()
On Error GoTo errHandle
    
    If ufgFeedback.GridRows <= 1 Then Exit Sub
    If ufgFeedback.MouseRowIndex <= 0 Then Exit Sub
        
    Call ShowUpdateFeedbackWindow(False)
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgFeedback_OnMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'弹出右键菜单
On Error GoTo errHandle
    If Button = 2 Then
        Dim objPopup As CommandBar
        Dim objControl As CommandBarControl

        Set cbrMain.Icons = zlCommFun.GetPubIcons
        Set objPopup = cbrMain.Add("右键菜单", xtpBarPopup)
        With objPopup.Controls
            Set objControl = .Add(xtpControlButton, TMenuType.mtFeedbackMod, "修改反馈(&U)"): objControl.IconId = 3003
            Set objControl = .Add(xtpControlButton, TMenuType.mtFeedbackDel, "删除反馈(&C)"): objControl.IconId = 4008
        End With
        objPopup.ShowPopup
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub
