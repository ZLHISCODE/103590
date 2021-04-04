VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmPatholConsultation 
   Caption         =   "病理会诊"
   ClientHeight    =   7980
   ClientLeft      =   75
   ClientTop       =   405
   ClientWidth     =   9135
   Icon            =   "frmPatholConsultation.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7980
   ScaleWidth      =   9135
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6975
      Left            =   0
      ScaleHeight     =   6975
      ScaleWidth      =   9135
      TabIndex        =   0
      Top             =   480
      Width           =   9135
      Begin VB.Frame framRequisition 
         Caption         =   "会诊记录"
         Height          =   6855
         Left            =   240
         TabIndex        =   1
         Top             =   0
         Width           =   8655
         Begin VB.TextBox txtAdvice 
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   240
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   3
            Top             =   5400
            Width           =   8175
         End
         Begin VB.TextBox txtResult 
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   240
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   2
            Top             =   3840
            Width           =   8175
         End
         Begin zl9PACSWork.ucFlexGrid ufgData 
            Height          =   3135
            Left            =   240
            TabIndex        =   4
            Top             =   240
            Width           =   8175
            _ExtentX        =   14420
            _ExtentY        =   5530
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
         Begin VB.Label labAdvice 
            Caption         =   "会诊意见："
            Height          =   255
            Left            =   240
            TabIndex        =   6
            Top             =   5160
            Width           =   1095
         End
         Begin VB.Label labResult 
            Caption         =   "会诊结果："
            Height          =   255
            Left            =   240
            TabIndex        =   5
            Top             =   3480
            Width           =   1095
         End
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   7
      Top             =   7620
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPatholConsultation.frx":179A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9234
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
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmPatholConsultation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private mlngCurAdviceId As Long     '当前医嘱ID
Private mstrPrivs As String         '当前权限串
Private mblnMoved As Boolean        '是否转存

Private mlngCurDepartmentId As Long

Private mrecStudyInf As TStudyStateInf
Private mblnIsDoFeedback As Boolean

Private mblnDataModifyState As Boolean
Private mblnViewState As Boolean
Private mblnFeedBackState As Boolean

Private Enum TMenuType
    mtFeedback = 1      '反馈
    mtCancle = 2        '撤回
    mtView = 3          '已阅
    
    mtAddCon = 4        '添加会诊
    mtDelCon = 5        '删除会诊
End Enum

Public Sub zlRefresh(lngAdviceID As Long, ByVal blnReadOnly As Boolean, _
    strPrivs As String, ByVal blnMoved As Boolean, _
    ByVal lngCurDepartmentId As Long, ByVal blnIsDoFeedback As Boolean, Optional owner As Form = Nothing)
'病理会诊

    If lngAdviceID <= 0 Then
        Call ConfigConsultationFace(False, "医嘱ID无效请检查。")
        Exit Sub
    End If
    
'    If mlngCurAdviceId = lngAdviceId Then Exit Sub

    mlngCurAdviceId = lngAdviceID
    mstrPrivs = strPrivs
    mblnMoved = blnMoved
    mlngCurDepartmentId = lngCurDepartmentId
    mblnIsDoFeedback = blnIsDoFeedback
    
    
    '设置窗口标题
    If mblnIsDoFeedback Then
        Me.Caption = "病理会诊-反馈"
    Else
        Me.Caption = "病理会诊-申请"
    End If
    
    Call GetPatholStudyState(lngAdviceID, mrecStudyInf)
        
   
    If mrecStudyInf.strPatholNumber = "" Then
        Call ConfigConsultationFace(False, "该检查尚未生成有效的病理号，请确认该检查是否已被核收。")
        
        If Not (owner Is Nothing) Then
            Call MsgBoxD(Me, "该检查尚未生成有效的病理号，请确认该检查是否已被核收。", vbOKOnly, Me.Caption)
        End If
        
        Exit Sub
    Else
        Call ConfigConsultationFace(True)
    End If
    
    '载入会诊数据
    Call LoadConsultationData
    
    '配置权限
    Call ConfigPopedom(blnReadOnly)
    
    '如果有拥有者，则弹出窗口
    If Not (owner Is Nothing) Then
        Call Me.Show(0, owner)
    End If
End Sub


Private Sub ConfigPopedom(ByVal blnIsReadOnly As Boolean)
'配置权限
    Dim blnIsAllowConRequest As Boolean
    Dim blnIsAllowConFeedback As Boolean
    
    blnIsAllowConRequest = CheckPopedom(mstrPrivs, "会诊申请")
    blnIsAllowConFeedback = CheckPopedom(mstrPrivs, "会诊反馈")
    
    mblnDataModifyState = blnIsAllowConRequest And Not mblnIsDoFeedback And Not blnIsReadOnly
    mblnViewState = blnIsAllowConRequest And Not mblnIsDoFeedback And Not blnIsReadOnly
    mblnFeedBackState = blnIsAllowConFeedback And mblnIsDoFeedback And Not blnIsReadOnly

    ufgData.ReadOnly = blnIsReadOnly
End Sub


Private Sub ConfigConsultationFace(ByVal blnIsValid As Boolean, Optional ByVal strHintInf As String = "")
'配置特检界面

    mblnDataModifyState = blnIsValid
    mblnViewState = blnIsValid
    mblnFeedBackState = blnIsValid
    
    If blnIsValid Then
        Call ufgData.CloseHintInf
    Else
        Call ufgData.ShowHintInf(strHintInf)
    End If
End Sub



Private Sub LoadConsultationData()
'载入会诊数据到列表
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    strSql = "select Id, 申请医师, 会诊单位, 会诊医师, 会诊类型, 会诊时间, 截止时间, 检查描述,诊断结果,诊断意见,备注,当前状态, 完成时间 " & _
            " from 病理会诊信息 where 病理医嘱ID=[1] order by 当前状态,会诊时间,截止时间,完成时间"
            
'    If mblnMoved Then strSql = GetMovedDataSql(strSql)
        
    Set ufgData.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mrecStudyInf.lngPatholAdviceId)
    
    Call ufgData.RefreshData
    
    If ufgData.ShowingDataRowCount > 0 Then
        Call LoadConContext(1)
    End If
End Sub



Private Sub InitConsultationList()
'初始化会诊显示列表
    Dim strTemp As String
    
   
    
    '判断数据库参数表是否有数据 有则读取数据库参数  没有则加载默认
    strTemp = zlDatabase.GetPara("病理会诊列表配置", glngSys, G_LNG_PATHOLSYS_NUM, "")
    ufgData.DefaultColNames = gstrConsultationCols
     
    If strTemp = "" Then
        ufgData.ColNames = gstrConsultationCols
    Else
        ufgData.ColNames = strTemp
    End If
     '设置行数
    ufgData.GridRows = glngStandardRowCount
    '设置行高
    ufgData.RowHeightMin = glngStandardRowHeight
    ufgData.ColConvertFormat = gstrConsultationConvertFormat
    ufgData.IsShowPopupMenu = False
End Sub

Private Sub cbrMain_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error GoTo ErrorHand
    Select Case control.ID
        Case TMenuType.mtFeedback                   '反馈
            Call Menu_Edit_Feedback
        
        Case TMenuType.mtCancle                     '撤回
            Call Menu_Edit_Cancel
        
        Case TMenuType.mtView                       '已阅
            Call Menu_Edit_View
            
        Case TMenuType.mtAddCon                     '添加会诊
            Call Menu_Edit_AddCon
        
        Case TMenuType.mtDelCon                     '删除会诊
            Call Menu_Edit_DelCon
        
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

Private Sub cbrMain_ResizeClient(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
On Error Resume Next
    picBack.Left = Left
    picBack.Top = Top
    picBack.Width = Right - Left
    picBack.Height = Bottom - Top
End Sub

Private Sub cbrMain_Update(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error GoTo ErrorHand
    Select Case control.ID
        Case TMenuType.mtFeedback               '反馈
            control.Enabled = mblnFeedBackState
            
        Case TMenuType.mtCancle                 '撤回
            control.Enabled = mblnDataModifyState And ufgData.IsSelectionRow
            
        Case TMenuType.mtView                   '已阅
            control.Enabled = mblnViewState And ufgData.IsSelectionRow
            
        Case TMenuType.mtAddCon                 '添加会诊
            control.Enabled = mblnDataModifyState
        
        Case TMenuType.mtDelCon                 '删除会诊
            control.Enabled = mblnDataModifyState
            
    End Select
    Exit Sub
ErrorHand:
    If ErrCenter() = 1 Then Exit Sub
End Sub

Private Sub ufgData_OnColFormartChange()
'保存列表参数
     zlDatabase.SetPara "病理会诊列表配置", ufgData.GetColsString(ufgData), glngSys, G_LNG_PATHOLSYS_NUM
End Sub

Private Sub picBack_Resize()
'调整界面布局
On Error Resume Next
    framRequisition.Left = 0
    framRequisition.Top = 60
    framRequisition.Width = picBack.Width
    framRequisition.Height = picBack.Height - 60
    
    ufgData.Left = 120
    ufgData.Top = 240
    ufgData.Width = framRequisition.Width - 240
    ufgData.Height = framRequisition.Height - txtResult.Height - txtAdvice.Height - labResult.Height * 2 - 840
    
    labResult.Left = 120
    labResult.Top = ufgData.Top + ufgData.Height + 240
    
    txtResult.Left = 120
    txtResult.Top = labResult.Top + labResult.Height + 60
    txtResult.Width = ufgData.Width
    
    labAdvice.Left = 120
    labAdvice.Top = txtResult.Top + txtResult.Height + 120
    
    txtAdvice.Left = 120
    txtAdvice.Top = labAdvice.Top + labAdvice.Height + 60
    txtAdvice.Width = ufgData.Width
End Sub


Private Sub ShowNewConsultationWindow()
'显示会诊新增窗口
Dim frmConsultation As New frmPatholConsultation_New
On Error GoTo errFree
    Call frmConsultation.ShowConsultationWindow(ufgData, mrecStudyInf.lngPatholAdviceId, mlngCurDepartmentId, Me)
errFree:
    Call Unload(frmConsultation)
    Set frmConsultation = Nothing
End Sub


Private Sub Menu_Edit_AddCon()
'添加会诊
On Error GoTo errHandle

'    If mlngStudyProcedure <> TStudyProcedure.spDiagnose Then
'        Call MsgBoxD(Me, "当前病理执行过程处于非诊断阶段，不能进行会诊申请。", vbOKOnly, Me.Caption)
'        Exit Sub
'    End If
    
    Call ShowNewConsultationWindow
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Function CheckAllowUpdateConsultation(ByVal lngConsultationRow As Long) As Boolean
'检查是否允许更新会诊记录
    CheckAllowUpdateConsultation = IIf(ufgData.Text(lngConsultationRow, gstrConsultation_当前状态) <> "已查阅" And ufgData.Text(lngConsultationRow, gstrConsultation_当前状态) <> "已反馈", True, False)
    If Not CheckAllowUpdateConsultation Then
        Call MsgBoxD(Me, "该会诊已反馈或已查阅，不能执行此操作。", vbOKOnly, Me.Caption)
        Exit Function
    End If
    
    CheckAllowUpdateConsultation = IIf(ufgData.Text(lngConsultationRow, gstrConsultation_申请医师) = UserInfo.姓名, True, False)
    If Not CheckAllowUpdateConsultation Then
        Call MsgBoxD(Me, "该会诊只能由申请医师 [" & ufgData.Text(lngConsultationRow, gstrConsultation_申请医师) & "] 进行修改。", vbOKOnly, Me.Caption)
        Exit Function
    End If
End Function


Private Sub DelConsultationData(ByVal lngConsultationRow As Long)
'删除会诊记录
    Dim strSql As String
    
    strSql = "Zl_病理会诊_删除(" & ufgData.KeyValue(lngConsultationRow) & ")"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    Call ufgData.DelRow(lngConsultationRow, False)
End Sub


Private Sub CancelConsultationFinish(ByVal lngConsultationRow As Long)
'撤销会诊完成
    Dim strSql As String
    
    strSql = "Zl_病理会诊_状态(" & ufgData.KeyValue(lngConsultationRow) & ",1)"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    ufgData.Text(lngConsultationRow, gstrConsultation_当前状态) = "已撤销"

End Sub


Private Function CheckAllowCancelConsultation(ByVal lngConsultationRow As Long) As Boolean
'检查是否允许撤销查阅
    CheckAllowCancelConsultation = IIf(ufgData.Text(lngConsultationRow, gstrConsultation_申请医师) = UserInfo.姓名, True, False)
    If Not CheckAllowCancelConsultation Then
        Call MsgBoxD(Me, "该会诊只能由申请医师 [" & ufgData.Text(lngConsultationRow, gstrConsultation_申请医师) & "] 进行修改。", vbOKOnly, Me.Caption)
        Exit Function
    End If
End Function



Private Sub Menu_Edit_Cancel()
'撤销完成
On Error GoTo errHandle
    If Not ufgData.IsSelectionRow Then Exit Sub
    
    If ufgData.IsEmptyKey(ufgData.SelectionRow) Then
        Call MsgBoxD(Me, "请选择需要撤销的会诊记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    '检查是否允许撤销
    If Not CheckAllowCancelConsultation(ufgData.SelectionRow) Then Exit Sub
    
    If MsgBoxD(Me, "确认要撤销该会诊记录吗？", vbYesNo, Me.Caption) = vbNo Then Exit Sub
    
    '撤销会诊完成
    Call CancelConsultationFinish(ufgData.SelectionRow)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Edit_DelCon()
'删除会诊记录
On Error GoTo errHandle
    If Not ufgData.IsSelectionRow Then Exit Sub
    
    If ufgData.IsEmptyKey(ufgData.SelectionRow) Then
        Call MsgBoxD(Me, "请选择需要删除的会诊记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If Not CheckAllowUpdateConsultation(ufgData.SelectionRow) Then Exit Sub
    
    If MsgBoxD(Me, "确认要删除该会诊记录吗？", vbYesNo, Me.Caption) = vbNo Then Exit Sub
    
    '删除会诊记录
    Call DelConsultationData(ufgData.SelectionRow)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub



Private Sub ShowConsultationFeedback(ByVal lngConsultationRow As Long)
'显示会诊反馈窗口
Dim frmFeedback As New frmPatholConsultation_Feedback
On Error GoTo errFree
    Call frmFeedback.ShowFeedbackWindow(ufgData, Val(ufgData.KeyValue(lngConsultationRow)), mlngCurDepartmentId, Me)
    
    
errFree:
'    会诊窗口的显示使用非模态窗口，因此这里不能进行释放
'    Call Unload(frmFeedback)
'    Set frmFeedback = Nothing
End Sub


Private Sub Menu_Edit_Feedback()
'会诊反馈
On Error GoTo errHandle
    If Not ufgData.IsSelectionRow Then Exit Sub
    
    If ufgData.IsEmptyKey(ufgData.SelectionRow) Then
        Call MsgBoxD(Me, "请选择需要反馈的会诊记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    '判断当前反馈用户是否为该记录的会诊医生
    If UserInfo.姓名 <> ufgData.Text(ufgData.SelectionRow, gstrConsultation_会诊医师) Then
        Call MsgBoxD(Me, "当前用户与该记录的会诊医师不同，不能反馈，请选择会诊医师属于 [ " & UserInfo.姓名 & "] 的会诊记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    Call ShowConsultationFeedback(ufgData.SelectionRow)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Function CheckAllowViewConsultation(ByVal lngConsultationRow As Long)
'检查是否允许查看会诊记录
    CheckAllowViewConsultation = IIf(ufgData.Text(lngConsultationRow, gstrConsultation_当前状态) <> "已申请", True, False)
    
    If Not CheckAllowViewConsultation Then
        Call MsgBoxD(Me, "该会诊尚未反馈，不能查阅。", vbOKOnly, Me.Caption)
        Exit Function
    End If
    
    CheckAllowViewConsultation = IIf(ufgData.Text(lngConsultationRow, gstrConsultation_申请医师) = UserInfo.姓名, True, False)
    If Not CheckAllowViewConsultation Then
        Call MsgBoxD(Me, "该会诊只能由申请医师 [" & ufgData.Text(lngConsultationRow, gstrConsultation_申请医师) & "] 进行修改。", vbOKOnly, Me.Caption)
        Exit Function
    End If
End Function



Private Sub ShowFeedbackViewWindow(ByVal lngConsultationRow As Long)
'显示会诊反馈窗口
'Dim frmFeedbackView As New frmPatholConsultation_Feedback
'On Error GoTo errFree
'    Call frmFeedbackView.ShowFeedbackViewWindow(ufgData, Me)
    
    '修改会诊记录状态
    Call ViewConsultation(lngConsultationRow)
'errFree:
''    Call Unload(frmFeedback)
''    Set frmFeedback = Nothing
End Sub



Private Sub ViewConsultation(ByVal lngConsultationRow As Long)
'会诊查阅
    Dim strSql As String

    strSql = "Zl_病理会诊_状态(" & ufgData.KeyValue(lngConsultationRow) & ",3)"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)

    ufgData.Text(lngConsultationRow, gstrConsultation_当前状态) = "已查阅"
End Sub


Private Sub Menu_Edit_View()
'查看报告
On Error GoTo errHandle
    If Not ufgData.IsSelectionRow Then Exit Sub
    
    If ufgData.IsEmptyKey(ufgData.SelectionRow) Then
        Call MsgBoxD(Me, "请选择需要查看的会诊记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    '检查是否允许查看会诊记录
    If Not CheckAllowViewConsultation(ufgData.SelectionRow) Then Exit Sub
    
    Call ShowFeedbackViewWindow(ufgData.SelectionRow)
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub



Private Sub Form_Load()
On Error GoTo errHandle
    Call InitCommandBars
    
    Call RestoreWinState(Me, App.ProductName)
    
    '该窗口使用的非模式窗口显示，因此需要置前
    SetWindowPos Me.hWnd, -1, Me.CurrentX, Me.CurrentY, Me.ScaleWidth, Me.ScaleHeight, 3 '将窗口置顶
    
    Call InitConsultationList
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
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
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtFeedback, "反馈(&F)"): cbrControl.IconId = 9022
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtCancle, "撤回(&C)"): cbrControl.IconId = 3014
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtView, "已阅(&S)"): cbrControl.IconId = 225

        Set cbrControl = .Add(xtpControlButton, TMenuType.mtAddCon, "添加会诊(&A)"): cbrControl.IconId = 4112
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtDelCon, "删除会诊(&D)"): cbrControl.IconId = 4114
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
'        Set cbrControl = .Controls.Add(xtpControlButton, conMenu_Help_About, "关于…(A)")
    End With
    '---------------------工具栏定义------------------------------------------
    Set cbrToolBar = Me.cbrMain.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = True
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtFeedback, "反馈"): cbrControl.IconId = 9022
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtCancle, "撤回"): cbrControl.IconId = 3014
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtView, "已阅"): cbrControl.IconId = 225

        Set cbrControl = .Add(xtpControlButton, TMenuType.mtAddCon, "添加会诊"): cbrControl.IconId = 4112
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtDelCon, "删除会诊"): cbrControl.IconId = 4114
        
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
        cbrControl.BeginGroup = True
    End With
    
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
End Sub

Private Sub ClearConContext()
'清除会诊显示内容
    txtAdvice.Text = ""
    txtResult.Text = ""
End Sub


Private Sub LoadConContext(ByVal lngRow As Long)
'加载会诊报告内容
    txtResult.Text = ufgData.Text(lngRow, gstrConsultation_诊断结果)
    txtAdvice.Text = ufgData.Text(lngRow, gstrConsultation_诊断意见)
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
End Sub


Private Sub ufgData_OnClick()
On Error GoTo errHandle
    If Not ufgData.IsSelectionRow Then Exit Sub
    
    If ufgData.IsEmptyKey(ufgData.SelectionRow) Then
        Call ClearConContext
    Else
        Call LoadConContext(ufgData.SelectionRow)
    End If
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub



Private Sub ufgData_OnColsNameReSet()
On Error GoTo errHandle

    Call LoadConsultationData
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgData_OnDblClick()
'查阅会诊反馈
On Error GoTo errHandle
    Call ViewFeedback
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub ViewFeedback()
'显示会诊反馈窗口
Dim frmFeedbackView As New frmPatholConsultation_Feedback
On Error GoTo errFree
    If Not ufgData.IsSelectionRow Then Exit Sub
    
    If ufgData.IsEmptyKey(ufgData.SelectionRow) Then
        Call MsgBoxD(Me, "请选择需要查看的会诊记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    Call frmFeedbackView.ShowFeedbackViewWindow(ufgData, Me)

errFree:
'    Call Unload(frmFeedback)
'    Set frmFeedback = Nothing
End Sub


