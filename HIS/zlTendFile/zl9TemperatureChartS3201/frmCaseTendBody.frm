VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmCaseTendBody 
   Caption         =   "体温作图"
   ClientHeight    =   5280
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9630
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCaseTendBody.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5280
   ScaleWidth      =   9630
   WindowState     =   2  'Maximized
   Begin zl9TemperatureChartS3201.usrBodyEditor BodyEdit 
      Height          =   3495
      Left            =   720
      TabIndex        =   1
      Top             =   720
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   6165
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   4920
      Width           =   9630
      _ExtentX        =   16986
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmCaseTendBody.frx":08CA
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14076
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
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   360
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmCaseTendBody"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'***************************************************************
'病人基本信息
'***************************************************************
Private Type type_Patient
    lng病人id As Long
    lng主页id As Long
    lng病区ID As Long
    lng科室ID As Long
    lng出院 As Long
    lng婴儿 As Long
    lng编辑 As Long
    lng护理等级 As Long
    lng文件ID As Long
    lng原始大小 As Long
    lngPage As Long
End Type

Private T_Info As type_Patient

Private mblnChildForm As Boolean
Private mcbrToolBar As CommandBar
Private mcbr查看 As CommandBarControl
Private mstrPrivs As String
Private mstrSQL As String
Private mblnShowing As Boolean
Private mblnChanged As Boolean
Private mfrmMain As Form
Private mIntDataEditor As Integer
Private mblnMove As Boolean
Private mfrmTendBody As Object

Public Event AfterPrint()
Public Event CmdClick(ByVal strParam As String)

'######################################################################################################################
'自定义函数、过程区域

Public Function ShowEdit(ByVal frmMain As Object, strParam As String, Optional ByVal bytMode As Byte = 1, Optional strPrivs As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim RS As New ADODB.Recordset
    Dim varParam As Variant
    Dim strPar As String
    Dim strTmp As String
    Dim blnShowing As Boolean
    
    mblnMove = False
    mblnChildForm = True
    mblnChanged = False
    mstrPrivs = strPrivs
    
    blnShowing = mblnShowing
    
    
    mblnShowing = True
    
    If strParam <> "" Then varParam = Split(strParam, ";")
    
    If blnShowing Then
        If Val(varParam(0)) = T_Info.lng病人id Or Val(varParam(1)) = T_Info.lng主页id And T_Info.lng科室ID = Val(varParam(2)) Then
            Call ShowWindow(Me.hWnd, SW_RESTORE)
            Call BringWindowToTop(Me.hWnd)
            Exit Function
        End If
    End If
    
    Set BodyEdit.ParentForm = Me
    Set mfrmMain = frmMain

    '参数格式：病人ID;主页ID;病区ID;文件ID;出院;编辑;婴儿;是否更具窗体大小自动校正体温单格式(1 否 0 校正)页号(默认显示第几页,如果页号超出范围就按缺省显示,0按缺省显示)
    
    '初始化参数
    
    T_Info.lng婴儿 = 0
    T_Info.lng出院 = 0
    T_Info.lng编辑 = 0
    T_Info.lng原始大小 = 0
    T_Info.lngPage = 0
    
    T_Info.lng病人id = Val(varParam(0))
    T_Info.lng主页id = Val(varParam(1))
    T_Info.lng病区ID = Val(varParam(2))
    T_Info.lng科室ID = Val(varParam(2))
    T_Info.lng文件ID = Val(varParam(3))
    
    If UBound(varParam) > 3 Then T_Info.lng出院 = Val(varParam(4))
    If UBound(varParam) > 4 Then T_Info.lng编辑 = Val(varParam(5))
    If InStr(1, ";" & mstrPrivs & ";", ";体温单作图;") = 0 Then
        T_Info.lng编辑 = 0
    Else
        T_Info.lng编辑 = 1
    End If
    If UBound(varParam) > 5 Then T_Info.lng婴儿 = Val(varParam(6))
    If UBound(varParam) > 6 Then T_Info.lng原始大小 = Val(varParam(7))
    If UBound(varParam) > 7 Then
        T_Info.lngPage = Val(varParam(8))
    Else
        T_Info.lngPage = glngCurPage
    End If
    
    If blnShowing = False Then Call InitMenuBar
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select 出院科室ID,nvl(数据转出,0) 转出  from 病案主页 Where 病人id=[1] And 主页id=[2] "
    Set RS = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, T_Info.lng病人id, T_Info.lng主页id)
    If RS.BOF = False Then
        T_Info.lng科室ID = Val(zlCommFun.Nvl(RS("出院科室ID").Value))
        If T_Info.lng出院 = 1 Then mblnMove = (Val(RS("转出")) <> 0)
    End If
    
    
    '------------------------------------------------------------------------------------------------------------------
    If OpenPatientMap = False Then
        Unload Me
        Exit Function
    End If
    
    Call GetTendEidor
    
    If blnShowing = False Then
        Hook Me
        
        If bytMode = 1 Then
            Me.Show , mfrmMain
        Else
            Me.Show 1, mfrmMain
        End If
        
        ShowEdit = mblnChanged
    End If
End Function

Public Function zlInit() As Boolean
    mblnChildForm = True
End Function

Public Function GetCurvePage() As Long
   GetCurvePage = BodyEdit.intPage
End Function

Public Sub zlDataEditor(ByVal intDataEditor As Integer)
    BodyEdit.DateEditor = intDataEditor
End Sub

Public Function zlRefresh(ByVal frmParent As Form, strParam As String, Optional strPrivs As String) As Boolean

   '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim RS As New ADODB.Recordset
    Dim varParam As Variant
    Dim strPar As String
    Dim strTmp As String
    Dim intBaby As Integer
    
    mblnMove = False
    mstrPrivs = strPrivs
    mblnChildForm = True
    stbThis.Visible = Not mblnChildForm
    cbsThis.ActiveMenuBar.Visible = False
    cbsThis.RecalcLayout
    
    mblnChanged = False
    
    Set BodyEdit.ParentForm = frmParent
    
    If strParam <> "" Then varParam = Split(strParam, ";")
    
    '参数格式：病人ID;主页ID;病区ID;文件ID;出院;编辑;婴儿;是否更具窗体大小自动校正体温单格式(1 否 0校正);页号(默认显示第几页,如果页号超出范围就按缺省显示,0按缺省显示)
    If Val(varParam(3)) <> T_Info.lng文件ID Then
        glngCurPage = 0
    Else
        If UBound(varParam) > 5 Then
            intBaby = Val(varParam(6))
        Else
            intBaby = 0
        End If
        
        If T_Info.lng婴儿 <> intBaby Then
            glngCurPage = 0
        End If
    End If
    
    '初始化参数
    T_Info.lng婴儿 = 0
    T_Info.lng出院 = 0
    T_Info.lng编辑 = 0
    T_Info.lng原始大小 = 0
    T_Info.lngPage = 0
    
    T_Info.lng病人id = Val(varParam(0))
    T_Info.lng主页id = Val(varParam(1))
    T_Info.lng病区ID = Val(varParam(2))
    T_Info.lng科室ID = Val(varParam(2))
    T_Info.lng文件ID = Val(varParam(3))
    
    If UBound(varParam) > 3 Then T_Info.lng出院 = Val(varParam(4))
    If UBound(varParam) > 4 Then T_Info.lng编辑 = Val(varParam(5))
    If InStr(1, ";" & mstrPrivs & ";", ";体温单作图;") = 0 Then
        T_Info.lng编辑 = 0
    Else
        T_Info.lng编辑 = 1
    End If
    If UBound(varParam) > 5 Then T_Info.lng婴儿 = Val(varParam(6))
    If UBound(varParam) > 6 Then T_Info.lng原始大小 = Val(varParam(7))
    If UBound(varParam) > 7 Then
        T_Info.lngPage = Val(varParam(8))
    Else
        T_Info.lngPage = glngCurPage
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select 出院科室ID,nvl(数据转出,0) 转出 from 病案主页 Where 病人id=[1] And 主页id=[2] "
    Set RS = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, T_Info.lng病人id, T_Info.lng主页id)
    If RS.BOF = False Then
        T_Info.lng科室ID = Val(zlCommFun.Nvl(RS("出院科室ID").Value))
        If T_Info.lng出院 = 1 Then mblnMove = (Val(RS("转出")) <> 0)
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    If OpenPatientMap = False Then
        Unload Me
        Exit Function
    End If
    
    Call GetTendEidor
    
    zlRefresh = True
End Function

Private Function OpenPatientMap() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim RS As New ADODB.Recordset
    Dim strParam As String
    
    T_Info.lng护理等级 = 3
    gstrSQL = "Select zl_PatitTendGrade([1],[2]) As 护理等级 From dual"
    Set RS = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, T_Info.lng病人id, T_Info.lng主页id)
    If RS.BOF = False Then T_Info.lng护理等级 = zlCommFun.Nvl(RS("护理等级"), 3)
    
    '重新提取文件ID
    gstrSQL = "select A.ID from 病人护理文件 A,病历文件列表 B" & _
       "    where A.病人ID=[1] and A.主页Id=[2] and A.婴儿=[3] and A.科室ID=[4] and A.格式ID=B.ID and B.种类=3 and B.保留=-1"
    Set RS = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, T_Info.lng病人id, T_Info.lng主页id, T_Info.lng婴儿, T_Info.lng科室ID)
    If mblnMove = True Then
        gstrSQL = Replace(gstrSQL, "病人护理文件", "H病人护理文件")
    End If
    
    If RS.BOF = False Then T_Info.lng文件ID = Val(zlCommFun.Nvl(RS("ID")))
    '初始化曲线菜单
    If InitBodyLine = False Then Exit Function
    
    '参数：病人ID;主页ID;病区ID;文件ID;出院标志;编辑标志;婴儿;护理等级;原始大小;页号(默认显示第几页,如果页号超出范围就按缺省显示,0按缺省显示)
    strParam = T_Info.lng病人id & ";" & T_Info.lng主页id & ";" & T_Info.lng病区ID & ";" & T_Info.lng文件ID & ";" & _
        T_Info.lng出院 & ";" & T_Info.lng编辑 & ";" & T_Info.lng婴儿 & ";" & T_Info.lng护理等级 & ";" & T_Info.lng原始大小 & ";" & T_Info.lngPage
    Call BodyEdit.zlMenuClick("初始化", strParam)
        
    OpenPatientMap = True
    
End Function

Private Function InitBodyLine() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rsTmp As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim cbrItem As CommandBarControl
    Dim strSQL As String
    
    On Error GoTo errHand
    
    '--曲线设置检查
    mstrSQL = "SELECT A.记录名,A.项目序号 FROM 体温记录项目 A,护理记录项目 B " & _
            "WHERE A.记录法 =1 And A.项目序号=B.项目序号 AND B.护理等级>=[1]  And Nvl(b.应用方式,0)=1 " & _
            "ORDER BY A.排列序号"
            
    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, T_Info.lng护理等级)
    If rsTmp.BOF Then
        MsgBox "无体温单操作曲线项目，请在护理项目管理中设置！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '--记录频次时间段设置检查
    mstrSQL = " Select Distinct nvl(记录频次,2) 频次  From 体温记录项目 A,护理记录项目 B" & _
            "   WHERE A.记录法 =2 And A.项目序号<>3 And  项目表示<>4 And A.项目序号=B.项目序号 AND B.护理等级>=[1] And Nvl(b.应用方式,0)=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, T_Info.lng护理等级)
    
    Do While Not rsTmp.EOF
        strSQL = "select Count(*) 记录数 From 护理项目频次 where 频次=[1]"
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsTmp!频次))
        If Val(rsData!记录数) < Val(rsTmp!频次) Then
            MsgBox "护理项目记录频次时段设置不完整，请在护理项目管理中设置！", vbInformation, gstrSysName
            Exit Function
        End If
    rsTmp.MoveNext
    Loop
    '--汇总项目时间段设置检查
    mstrSQL = "select count(*) 记录数 from 护理汇总时段 Where 单据=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption)
    If Val(rsTmp!记录数) < 3 Then
        MsgBox "护理汇总时段设置不完整，请在护理项目管理中设置！", vbInformation, gstrSysName
        Exit Function
    End If
    
    InitBodyLine = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function PrintData(ByVal bytMode As Byte, Optional ByVal strPrintDevice As String = "") As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim blnCur As Boolean
    Dim lngBeginY As Long
    Dim intBeginPage As Integer
    Dim intPrintRange As Integer
    Dim strPage  As String, strParam As String
    
    '传入了打印机名称,说明是批量打印,自动从第1页开始打印,不进行任何询问
    '返回:0-取消,2-预览,1-打印
    
    frmCaseTendBodyPrintSet.cmdPrint.Visible = (bytMode = 1)
    frmCaseTendBodyPrintSet.cmdPreview.Visible = (bytMode = 2)
    
    
    If strPrintDevice = "" Then
        'strParam = T_Info.lng文件ID & ";" & T_Info.lng病人ID & ";" & T_Info.lng主页ID & ";" & T_Info.lng科室Id & ";" & T_Info.lng科室Id
        strParam = T_Info.lng文件ID & ";" & Me.BodyEdit.AllPage
        bytMode = frmCaseTendBodyPrintSet.PrintSet(Me, True, strParam, intPrintRange, lngBeginY, intBeginPage, strPage, mstrPrivs)
    Else
        bytMode = 2
        intPrintRange = 2
    End If
    If bytMode = 0 Then Exit Function
    If intBeginPage <= 0 Then intBeginPage = -1
    
    '打印当前页传入当前页号
    If intPrintRange = 0 Then
        strPage = Me.BodyEdit.intPage - 1
    End If
    
    Select Case bytMode
    Case 2  '打印
        Call BodyEdit.PrintState(intPrintRange, True, lngBeginY, intBeginPage, strPrintDevice, strPage)
    Case 1  '预览
        Call BodyEdit.PrintState(intPrintRange, False, lngBeginY, intBeginPage, strPrintDevice, strPage)
    End Select

End Function

Public Function zlPrintBody(Optional ByVal bytMode As Byte = 2, Optional ByVal strPrintDevice As String) As Long
    '入参:1-预览,2-打印
    '返回值:0-失败;1-成功;2-打印
    gblnPrinted = False
    Call PrintData(IIf(bytMode = 1, 2, 1), strPrintDevice)
    zlPrintBody = IIf(gblnPrinted, 2, 1)
End Function

Private Function InitMenuBar() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    
    On Error GoTo errHand
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    cbsThis.ActiveMenuBar.Title = "菜单栏"
    
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    With cbsThis.Options
        .AlwaysShowFullMenus = False
        .ShowExpandButtonAlways = False
        .UseDisabledIcons = True
        .SetIconSize True, 24, 24
        .LargeIcons = True
    End With

    '------------------------------------------------------------------------------------------------------------------
    '菜单定义
    cbsThis.ActiveMenuBar.Title = "菜单"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagStretched)
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&E)")
                
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
       
'        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存数据(&S)"): cbrControl.BeginGroup = True
'        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "恢复数据(&R)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)")
        cbrControl.BeginGroup = True
    End With


    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Adjust, "设置记录(&E)")
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Show, "设置显示(&D)")
    End With

    Set mcbr查看 = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    With mcbr查看.CommandBar.Controls
                
'       Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
'
'       cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
'       cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
                
    End With
    
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "论坛(&F)", -1, False  '固有
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)..."):
        cbrControl.BeginGroup = True
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    '工具栏定义
'    Set mcbrToolBar = cbsThis.Add("标准", xtpBarTop)
'    mcbrToolBar.ShowTextBelowIcons = False
'    mcbrToolBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
'    mcbrToolBar.ModifyStyle XTP_CBRS_GRIPPER, 0
'
'    With mcbrToolBar.Controls
'        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
'        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
'        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存")
'        cbrControl.BeginGroup = True
'        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "恢复")
'        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助")
'        cbrControl.BeginGroup = True
'        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
'    End With
'
'    '定位工具栏
'    '------------------------------------------------------------------------------------------------------------------
'
'    For Each cbrControl In mcbrToolBar.Controls
'        If cbrControl.Type <> xtpControlCustom And cbrControl.Type <> xtpControlLabel Then
'            cbrControl.Style = xtpButtonIconAndCaption
'        End If
'    Next
    
     '快键绑定
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("E"), conMenu_Edit_Adjust
        .Add FCONTROL, Asc("D"), conMenu_View_Show
        .Add FCONTROL, Asc("S"), conMenu_Edit_Save
        .Add FCONTROL, Asc("R"), conMenu_Edit_Reuse
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_F5, conMenu_View_Refresh
    End With
    
    InitMenuBar = True
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub BodyEditCur(ByVal intDataEditor As Integer, Optional ByVal strParam As String = "")
    Call GetTendEidor
    If intDataEditor = 0 Then
        Call BodyEdit.zlMenuClick("体温数据编辑", strParam)
    ElseIf intDataEditor = 1 Then
         Call BodyEdit.zlMenuClick("体温数据显示设置", strParam)
    End If
End Sub

Private Sub BodyEdit_DbClickCur(ByVal intDataEditor As Integer)
    Call BodyEditCur(intDataEditor)
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strKey As String
    Dim lngLoop As Long
    Dim strSQL() As String
    Dim blnTran As Boolean
    Dim lngIndex As Long
    Dim cbrControl As CommandBarControl
    Dim lngKey As Long
    
    On Error GoTo errHand
    
    ReDim Preserve strSQL(1 To 1)
        
    Select Case Control.Id
        Case conMenu_File_PrintSet   '打印设置
            
            On Error Resume Next
            Call frmPrintSet.ShowMe(Me, 1)
            
        Case conMenu_File_Preview  '打印预览
            
            Call PrintData(2)
            
        Case conMenu_File_Print  '打印
        
            Call PrintData(1)
        
        Case conMenu_View_ToolBar_Button

'            cbsThis(2).Visible = Not cbsThis(2).Visible
'            cbsThis.RecalcLayout

        Case conMenu_View_ToolBar_Text

'            For Each cbrControl In cbsThis(1).Controls
'                cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
'            Next
'
'            cbsThis.RecalcLayout
            
        Case conMenu_View_StatusBar
        
            stbThis.Visible = Not stbThis.Visible
            cbsThis.RecalcLayout
            
        Case conMenu_Edit_Adjust '设置记录
            Call BodyEditCur(0)
            
        Case conMenu_View_Show '设置显示
            Call BodyEditCur(1)
            
        Case conMenu_Edit_Save '保存数据
            
        Case conMenu_Edit_Reuse '数据恢复
            
        Case conMenu_Help_Help
        
            Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
        
        Case conMenu_Help_About
            
            Call ShowAbout(Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision)
            
        Case conMenu_Help_Web_Home
            
            Call zlHomePage(Me.hWnd)
            
        Case conMenu_Help_Web_Forum '中联论坛
            Call zlWebForum(Me.hWnd)
            
        Case conMenu_Help_Web_Mail
            
            Call zlMailTo(Me.hWnd)
        
        Case conMenu_File_Exit
            Unload Me
            Exit Sub
    End Select
    
    Exit Sub
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTran Then gcnOracle.RollbackTrans
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Control.Type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup
        End Select
    End If
    
    Err = 0
    On Error Resume Next
    
    Select Case Control.Id

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Adjust, conMenu_View_Show
        
        Control.Enabled = (T_Info.lng编辑 = 1)
        
    Case conMenu_View_ToolBar_Button
    
        Control.Checked = Me.cbsThis(2).Visible
        
    Case conMenu_View_ToolBar_Text
    
        Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
        
    Case conMenu_View_ToolBar_Size
    
        Control.Checked = Me.cbsThis.Options.LargeIcons
        
    Case conMenu_View_StatusBar
    
        Control.Checked = Me.stbThis.Visible
        
    End Select
End Sub

Private Sub BodyEdit_zlAfterPrint()
    gblnPrinted = True
    RaiseEvent AfterPrint
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub cbsThis_Resize()
    On Error Resume Next
    
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngRight As Long
    Dim lngBottom As Long  '客户区域的大小

    Call cbsThis.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    With BodyEdit
        .mblnResize = True
        .Left = lngLeft
        .Width = lngRight - lngLeft
        .Top = lngTop
        .mblnResize = False
        .Height = lngBottom - lngTop
    End With
End Sub

Private Sub Form_Load()
    If Not mblnChildForm Then
         Call RestoreWinState(Me, App.ProductName)
    End If
End Sub

Private Sub GetTendEidor()
    If Not gobjTendEditor Is Nothing Then Set gobjTendEditor = Nothing
    Set gobjTendEditor = Me
End Sub

Private Sub BodyEdit_CmdClick(ByVal strParam As String)
    Dim arrParam() As String
    If mfrmTendBody Is Nothing Then Set mfrmTendBody = New frmCaseTendBody
    
    If mfrmTendBody.ShowEdit(BodyEdit.ParentForm, strParam, 0, mstrPrivs) Then
        arrParam = Split(strParam, ";")
        If UBound(arrParam) > 6 Then arrParam(7) = 0
        If UBound(arrParam) > 7 Then
            strParam = arrParam(0) & ";" & arrParam(1) & ";" & arrParam(2) & ";" & arrParam(3) & ";" & arrParam(4) & ";" & arrParam(5) & ";" & arrParam(6) & ";" & arrParam(7)
        Else
            strParam = Join(arrParam, ";")
        End If
        
        '刷新体温单页面
        Call zlRefresh(BodyEdit.ParentForm, strParam, mstrPrivs)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnHook Me
    
    mblnShowing = False
    Set mfrmTendBody = Nothing
    
    If Not mblnChildForm Then
        Call SaveWinState(Me, App.ProductName)
    Else
        mblnChanged = True
    End If
    If Not gobjTendEditor Is Nothing Then Set gobjTendEditor = Nothing
    
    '卸载用户控件对象 （窗体关闭时用户控件的 UserControl_Terminate 事件无法进入 所以放在父窗体关闭执行 ）
    Call BodyEdit.ReleaseObj
End Sub

