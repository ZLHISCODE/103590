VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmPatholRequisition 
   Caption         =   "检查申请"
   ClientHeight    =   6825
   ClientLeft      =   75
   ClientTop       =   405
   ClientWidth     =   10350
   Icon            =   "frmPatholRequisition.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6825
   ScaleWidth      =   10350
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picRequest 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   240
      ScaleHeight     =   3255
      ScaleWidth      =   9615
      TabIndex        =   3
      Top             =   360
      Width           =   9615
      Begin zl9PACSWork.ucFlexGrid ufgRequest 
         Height          =   2895
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   5106
         GridRows        =   21
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
   End
   Begin VB.PictureBox picRequestContext 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   240
      ScaleHeight     =   3015
      ScaleWidth      =   9615
      TabIndex        =   1
      Top             =   3480
      Width           =   9615
      Begin zl9PACSWork.ucFlexGrid ufgContext 
         Height          =   2415
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   4260
         GridRows        =   21
         IsKeepRows      =   0   'False
         BackColor       =   12648447
         IsEnterNextCell =   0   'False
         IsBtnNextCell   =   0   'False
         IsCopyAdoMode   =   0   'False
         IsEjectConfig   =   -1  'True
         Editable        =   1
         HeadFontCharset =   134
         HeadFontWeight  =   400
         DataFontCharset =   134
         DataFontWeight  =   400
      End
      Begin VB.Label labSubTitle 
         AutoSize        =   -1  'True
         Caption         =   "申请项目"
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
         Height          =   180
         Left            =   240
         TabIndex        =   5
         Top             =   120
         Width           =   780
      End
      Begin VB.Line linSubTitle 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         X1              =   0
         X2              =   9840
         Y1              =   240
         Y2              =   240
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   6465
      Width           =   10350
      _ExtentX        =   18256
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPatholRequisition.frx":179A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11377
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
   Begin XtremeDockingPane.DockingPane dkpRequest 
      Bindings        =   "frmPatholRequisition.frx":202E
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmPatholRequisition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnSpecialExamState As Boolean
Private mblnSpecialExam As Boolean
Private mblnSlices As Boolean
Private mblnSlicesState As Boolean

Private mblnReqSpecialExam As Boolean
Private mblnReqSlices As Boolean
Private mblnReqGet As Boolean
Private mblnDelReq As Boolean

Private mlngCurAdviceId As Long
Private mstrPrivs As String
Private mblnMoved As Boolean
Private mlngRequestType As Long

Private mlngCurDepartmentId As Long

Private mrecStudyInf As TStudyStateInf

Private Enum TMenuType
    mtReqGet = 1         '补取申请
    mtReqSlices = 2      '制片申请
    mtReqSpecialExam = 3 '特检申请
    mtDelReq = 4         '撤销申请
    
    mtAddSpeExamPro = 5  '添加
    mtNewSE = 6          '补做
    mtRedoSE = 7         '重做
    mtDelSE = 8          '删除
    
'    mtAddSlicesPro = 9   '增加
'    mtDelSlicesPro = 10  '删除
End Enum

Public blnIsUpdate As Boolean


Public Sub zlRefresh(lngAdviceID As Long, ByVal blnReadOnly As Boolean, strPrivs As String, ByVal blnMoved As Boolean, _
    ByVal lngCurDepartmentId As Long, Optional owner As Form = Nothing)
On Error GoTo errHandle

    If lngAdviceID <= 0 Then
        Call ConfigRequisitionFace(False, "医嘱ID无效请检查。")
        Exit Sub
    End If
    
'    If mlngCurAdviceId = lngAdviceId Then Exit Sub

    mlngCurAdviceId = lngAdviceID
    mstrPrivs = strPrivs
    mblnMoved = blnMoved
    mlngCurDepartmentId = lngCurDepartmentId
    blnIsUpdate = False
    
    Call GetPatholStudyState(lngAdviceID, mrecStudyInf)
    
   
    If mrecStudyInf.strPatholNumber = "" Then
        Call ConfigRequisitionFace(False, "该检查尚未生成有效的病理号，请确认该检查是否已被核收。")
        
        If Not (owner Is Nothing) Then
            Call MsgBoxD(Me, "该检查尚未生成有效的病理号，请确认该检查是否已被核收。", vbOKOnly, Me.Caption)
        End If
        
        Exit Sub
    Else
        '读取申请信息
        Call LoadRequestInf(mrecStudyInf.lngPatholAdviceId)
        
        '载入申请明细
        Call ufgRequest_OnClick
        
        Call ConfigRequisitionFace(True)
    End If
    
    Call ConfigPopedom(blnReadOnly)
    
    If Not (owner Is Nothing) Then
        Call Me.Show(1, owner)
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub ConfigPopedom(ByVal blnIsReadOnly As Boolean)
'配置权限
    Dim blnSpeExamPopedom As Boolean
    Dim blnSlicesPopedom As Boolean
    Dim blnMaterialPopedom As Boolean
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    blnSpeExamPopedom = CheckPopedom(mstrPrivs, "特检申请")
    blnSlicesPopedom = CheckPopedom(mstrPrivs, "制片申请")
    blnMaterialPopedom = CheckPopedom(mstrPrivs, "补取申请")
    
    mblnReqSpecialExam = blnSpeExamPopedom And Not blnIsReadOnly
    mblnReqSlices = blnSlicesPopedom And Not blnIsReadOnly
    mblnReqGet = blnMaterialPopedom And Not blnIsReadOnly
    mblnDelReq = (blnSpeExamPopedom Or blnSlicesPopedom Or blnMaterialPopedom) And Not blnIsReadOnly
    
    mblnSpecialExam = blnSpeExamPopedom And Not blnIsReadOnly
    
    mblnSlices = blnSlicesPopedom And Not blnIsReadOnly
    
    ufgRequest.ReadOnly = blnIsReadOnly
    ufgContext.ReadOnly = blnIsReadOnly
    
    '得到制片状态从而用来控制 “特检申请”和“制片申请” 两个按钮
    strSql = "select distinct 当前状态 from 病理检查信息 a,病理制片信息 b where a.病理医嘱id = b.病理医嘱id and a.医嘱id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "得到制片状态", mlngCurAdviceId)
    
    If rsTemp.RecordCount < 1 Then
        mblnReqSlices = False
        mblnReqSpecialExam = False
        Exit Sub
    End If
    
    mblnReqSlices = IIf(Nvl(rsTemp!当前状态, 0) = 2, True, False)
    mblnReqSpecialExam = IIf(Nvl(rsTemp!当前状态, 0) = 2, True, False)

End Sub



Private Sub ConfigRequisitionFace(ByVal blnIsValid As Boolean, Optional ByVal strHintInf As String = "")
'配置申请界面
    mblnReqSpecialExam = blnIsValid
    mblnReqSlices = blnIsValid
    mblnReqGet = blnIsValid
    mblnDelReq = blnIsValid

    mblnSpecialExam = blnIsValid

    mblnSlices = blnIsValid
    
    If blnIsValid Then
        Call ufgRequest.CloseHintInf
        Call ufgContext.CloseHintInf
    Else
        Call ufgRequest.ShowHintInf(strHintInf)
        Call ufgContext.ShowHintInf(strHintInf)
    End If
End Sub


Private Sub InitFace()
'初始化界面布局
    Dim Pane1 As Pane, Pane2 As Pane

    With dkpRequest
        .CloseAll
        .SetCommandBars cbrMain
        .Options.HideClient = True
        .Options.UseSplitterTracker = False '实时拖动
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = True
    End With

    Set Pane1 = dkpRequest.CreatePane(1, 0, Round(Me.Height * 3 / 5), DockTopOf, Nothing)
    Pane1.Title = "申请记录"
    Pane1.Handle = picRequest.hWnd
    Pane1.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    Pane1.MinTrackSize.Width = 50
    Pane1.MinTrackSize.Height = 100

    Set Pane2 = dkpRequest.CreatePane(2, 0, Round(Me.Height * 2 / 5), DockBottomOf, Pane1)
    Pane2.Title = "申请明细"
    Pane2.Handle = picRequestContext.hWnd
    Pane2.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    Pane2.MinTrackSize.Width = 50
    Pane2.MinTrackSize.Height = 100
End Sub


Private Sub InitRequisitionList()
'初始化申请列表
    Dim strTemp As String
    
    '判断数据库参数表是否有数据 有则读取数据库参数  没有则加载默认
    strTemp = zlDatabase.GetPara("检查申请列表配置", glngSys, G_LNG_PATHOLSYS_NUM, "")
     
    If strTemp = "" Then
        ufgRequest.ColNames = gstrRequisitionCols
    Else
        ufgRequest.ColNames = strTemp
    End If
    
    '设置行数
    ufgRequest.GridRows = glngStandardRowCount
    '设置行高
    ufgRequest.RowHeightMin = glngStandardRowHeight
    
    ufgRequest.DefaultColNames = gstrRequisitionCols
    ufgRequest.ColConvertFormat = gstrRequisitionConvertFormat
    ufgRequest.IsShowPopupMenu = False
End Sub

Private Sub InitRequestContextList(ByVal lngRequestType As Long)
'初始化申请项目明细列表
    Dim strTemp As String
    
    mlngRequestType = lngRequestType
    
    Select Case lngRequestType
        Case 0, 1, 2
        
            '判断数据库参数表是否有数据 有则读取数据库参数  没有则加载默认
            strTemp = zlDatabase.GetPara("特检申请列表配置", glngSys, G_LNG_PATHOLSYS_NUM, "")
             
            If strTemp = "" Then
                ufgContext.ColNames = gstrRequest_SpeExam_Cols
            Else
                ufgContext.ColNames = strTemp
            End If
                   '禁止右键弹出列表配置窗口
            ufgContext.IsEjectConfig = False
            ufgContext.DefaultColNames = gstrRequest_SpeExam_Cols
            ufgContext.ColConvertFormat = gstrRequest_SpeExamConvertFormat
            
        Case 3
            
            '判断数据库参数表是否有数据 有则读取数据库参数  没有则加载默认
            strTemp = zlDatabase.GetPara("制片申请列表配置", glngSys, G_LNG_PATHOLSYS_NUM, "")
             
            If strTemp = "" Then
                ufgContext.ColNames = gstrRequest_Slices_Cols
            Else
                ufgContext.ColNames = strTemp
            End If
                   '禁止右键弹出列表配置窗口
            ufgContext.IsEjectConfig = False
            ufgContext.DefaultColNames = gstrRequest_Slices_Cols
            ufgContext.ColConvertFormat = gstrRequest_SlicesConvertFormat
        Case 4, 5
            
            '判断数据库参数表是否有数据 有则读取数据库参数  没有则加载默认
            strTemp = zlDatabase.GetPara("补取申请列表配置", glngSys, G_LNG_PATHOLSYS_NUM, "")
             
            If strTemp = "" Then
                ufgContext.ColNames = gstrRequest_Material_Cols
            Else
                ufgContext.ColNames = strTemp
            End If
                   '禁止右键弹出列表配置窗口
            ufgContext.IsEjectConfig = False
            ufgContext.DefaultColNames = gstrRequest_Material_Cols
            '设置行数
            ufgContext.GridRows = glngStandardRowCount
            '设置行高
            ufgContext.RowHeightMin = glngStandardRowHeight
        
            ufgContext.ColConvertFormat = gstrRequest_MaterialConvertFormat
    End Select
    
    ufgContext.IsShowPopupMenu = False
End Sub

Private Sub cbrMain_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
    Dim blnHasRequestRecord As Boolean
    Dim blnHasContextRecord As Boolean
    
On Error GoTo ErrorHand
    blnHasRequestRecord = ufgRequest.IsSelectionRow
    blnHasContextRecord = ufgContext.IsSelectionRow
    
    Select Case control.ID
        Case TMenuType.mtReqGet                       '补取申请
            Call Menu_Edit_ReqGet
        
        Case TMenuType.mtReqSlices                    '制片申请
            Call Menu_Edit_ReqSlices
        
        Case TMenuType.mtReqSpecialExam               '特检申请
            Call Menu_Edit_ReqSpecialExam
        
        Case TMenuType.mtDelReq                       '撤销申请
            Call Menu_Edit_DelReq
            
        Case TMenuType.mtAddSpeExamPro                '添加
            If mblnSpecialExamState And mblnSpecialExam Then
                Call Menu_Edit_AddSpeExamPro
            ElseIf mblnSlicesState And mblnSlices Then
                Call Menu_Edit_AddSlicesPro
            End If
        
        Case TMenuType.mtNewSE                        '补做
            Call Menu_Edit_NewSE
        
        Case TMenuType.mtRedoSE                       '重做
            Call Menu_Edit_RedoSE
        
        Case TMenuType.mtDelSE                        '删除
            If mblnSpecialExamState And mblnSpecialExam And blnHasContextRecord Then
                Call Menu_Edit_DelSE
            ElseIf mblnSlicesState And mblnSlices And blnHasContextRecord Then
                Call Menu_Edit_DelSlicesPro
            End If
            
'        Case TMenuType.mtAddSlicesPro                 '增加
'            Call Menu_Edit_AddSlicesPro
'
'        Case TMenuType.mtDelSlicesPro                 '删除
'            Call Menu_Edit_DelSlicesPro

        Case conMenu_File_Exit                        '退出
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
    Dim blnHasRequestRecord As Boolean
    Dim blnHasContextRecord As Boolean
    
On Error GoTo ErrorHand
    blnHasRequestRecord = ufgRequest.IsSelectionRow
    blnHasContextRecord = ufgContext.IsSelectionRow
    
    Select Case control.ID
        Case TMenuType.mtReqGet                         '补取申请
            control.Enabled = mblnReqGet
        
        Case TMenuType.mtReqSlices                      '制片申请
            control.Enabled = mblnReqSlices
        
        Case TMenuType.mtReqSpecialExam                 '特检申请
            control.Enabled = mblnReqSpecialExam
        
        Case TMenuType.mtDelReq                         '撤销申请
            control.Enabled = mblnDelReq
            
        Case TMenuType.mtAddSpeExamPro                  '添加
            control.Enabled = (mblnSpecialExamState And mblnSpecialExam) Or (mblnSlicesState And mblnSlices)
        
        Case TMenuType.mtNewSE                          '补做
            control.Enabled = mblnSpecialExamState And mblnSpecialExam
        
        Case TMenuType.mtRedoSE                         '重做
            control.Enabled = mblnSpecialExamState And mblnSpecialExam And blnHasContextRecord
        
        Case TMenuType.mtDelSE                          '删除
            control.Enabled = ((mblnSpecialExamState And mblnSpecialExam) Or (mblnSlicesState And mblnSlices)) And blnHasContextRecord
            
'        Case TMenuType.mtAddSlicesPro                   '添加
'            control.Enabled = mblnSlicesState And mblnSlices
'
'        Case TMenuType.mtDelSlicesPro                   '删除
'            control.Enabled = mblnSlicesState And mblnSlices And blnHasContextRecord

    End Select
    Exit Sub
ErrorHand:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgContext_OnColFormartChange()
    '窗体改变时保存列表配置
    
    Select Case mlngRequestType
        Case 0, 1, 2
        
            zlDatabase.SetPara "特检申请列表配置", ufgContext.GetColsString(ufgContext), glngSys, G_LNG_PATHOLSYS_NUM
            
        Case 3
        
            zlDatabase.SetPara "制片申请列表配置", ufgContext.GetColsString(ufgContext), glngSys, G_LNG_PATHOLSYS_NUM
           
        Case 4, 5
            
            zlDatabase.SetPara "补取申请列表配置", ufgContext.GetColsString(ufgContext), glngSys, G_LNG_PATHOLSYS_NUM
            
    End Select

End Sub

Private Sub ufgRequest_OnColFormartChange()
'窗体改变时保存列表配置
     zlDatabase.SetPara "检查申请列表配置", ufgRequest.GetColsString(ufgRequest), glngSys, G_LNG_PATHOLSYS_NUM
     
End Sub


Public Sub ChangeControlFace(ByVal lngRequestType As Long)
'改变控制界面
    mblnSpecialExamState = IIf("0,1,2" Like "*" & lngRequestType & "*", True, False)
    
    mblnSlicesState = IIf("3" Like "*" & lngRequestType & "*", True, False)
End Sub


Private Sub ShowSpecialExamRequestWindow()
'显示特检申请窗口
    Dim frmSpeExamRequest As New frmPatholRequisition_SpeExam
    On Error GoTo errFree
      
        
        '显示特检申请窗口
        Call frmSpeExamRequest.ShowSpeExamRequestWindow(ufgRequest, ufgContext, mrecStudyInf.lngPatholAdviceId, -1, Me)
        
        blnIsUpdate = frmSpeExamRequest.blnIsOk
errFree:
    Call Unload(frmSpeExamRequest)
    Set frmSpeExamRequest = Nothing
    
End Sub


Private Sub AddSpeExamProject(ByVal lngCurRequestId As Long, Optional ByVal blnIsBuZuo As Boolean = False)
'添加特检项目
    Dim frmSpeExamRequest As New frmPatholRequisition_SpeExam
    On Error GoTo errFree
    
        '显示特检申请窗口
        Call frmSpeExamRequest.ShowSpeExamRequestWindow(ufgRequest, ufgContext, mrecStudyInf.lngPatholAdviceId, lngCurRequestId, Me, blnIsBuZuo)
        
        blnIsUpdate = frmSpeExamRequest.blnIsOk
errFree:
    Call Unload(frmSpeExamRequest)
    Set frmSpeExamRequest = Nothing
End Sub


Private Sub Menu_Edit_AddSlicesPro()
'增加制片申请项目
On Error GoTo errHandle
    If Not ufgRequest.IsSelectionRow Then
        Call MsgBoxD(Me, "请选择所属的申请记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgRequest.IsEmptyKey(ufgRequest.SelectionRow) Then
        Call MsgBoxD(Me, "请选择有效的申请记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    '显示制片新增窗口
    Call ShowSlicesRequestWindow(ufgRequest.KeyValue(ufgRequest.SelectionRow))
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Edit_AddSpeExamPro()
'添加特检项目
On Error GoTo errHandle
    If Not ufgRequest.IsSelectionRow Then
        Call MsgBoxD(Me, "请选择所属的申请记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgRequest.IsEmptyKey(ufgRequest.SelectionRow) Then
        Call MsgBoxD(Me, "请选择有效的申请记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    '新增特检项目
    Call AddSpeExamProject(ufgRequest.KeyValue(ufgRequest.SelectionRow), False)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub DelSpeExamProject(ByVal lngSpeExamRow As Long)
'删除特检项目
    Dim strSql As String
    
    strSql = "Zl_病理申请_特检项目_删除(" & ufgContext.KeyValue(lngSpeExamRow) & ")"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    Call ufgContext.RemoveRow(lngSpeExamRow)
End Sub


Private Function CheckAllowDelSpeExam(ByVal lngSpeExamRow As Long) As Boolean
'判断是否允许删除特检项目
    CheckAllowDelSpeExam = IIf(ufgContext.Text(lngSpeExamRow, gstrRequest_SpeExam_当前状态) = "已申请", True, False)
    
    If Not CheckAllowDelSpeExam Then
        Call MsgBoxD(Me, "该特检项目已被接受或完成，不能执行删除操作。", vbOKOnly, Me.Caption)
    End If
    
End Function


Private Sub CancelRequest(ByVal lngRequestRow As Long)
'撤销申请
    Dim strSql As String
    
    strSql = "Zl_病理申请_删除(" & Val(ufgRequest.KeyValue(lngRequestRow)) & ")"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    Call ufgRequest.RemoveRow(lngRequestRow)
End Sub



Private Function CheckAllowDelRequest(ByVal lngRequestRow As Long)
'检查申请是否允许删除
    CheckAllowDelRequest = IIf(ufgRequest.Text(lngRequestRow, gstrRequisition_当前状态) = "已申请", True, False)
    
    If Not CheckAllowDelRequest Then
        Call MsgBoxD(Me, "该申请项目已被接受或执行，不能删除。", vbOKOnly, Me.Caption)
    End If
End Function


Private Sub Menu_Edit_DelReq()
'删除申请
On Error GoTo errHandle
    If Not ufgRequest.IsSelectionRow Then
        Call MsgBoxD(Me, "请选择需要删除的申请记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgRequest.IsEmptyKey(ufgRequest.SelectionRow) Then
        Call MsgBoxD(Me, "请选择有效的申请记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If Not CheckAllowDelRequest(ufgRequest.SelectionRow) Then Exit Sub
    
    If MsgBoxD(Me, "确认要删除该申请项目吗？", vbYesNo, Me.Caption) = vbNo Then Exit Sub
    
    '删除特检项目
    Call CancelRequest(ufgRequest.SelectionRow)
    
    If ufgRequest.ShowingDataRowCount > 0 Then
        Call ufgRequest.DataGrid.Select(ufgRequest.GridRows - 1, 0)
    End If
    
    Call ufgRequest_OnClick
    
    
    blnIsUpdate = True
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Edit_DelSE()
'删除特检项目
On Error GoTo errHandle
    If Not ufgContext.IsSelectionRow Then
        Call MsgBoxD(Me, "请选择需要删除的特检项目。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgContext.IsEmptyKey(ufgContext.SelectionRow) Then
        Call MsgBoxD(Me, "请选择有效的特检项目。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If Not CheckAllowDelSpeExam(ufgContext.SelectionRow) Then Exit Sub
    
    If MsgBoxD(Me, "确认要删除该特检项目吗？", vbYesNo, Me.Caption) = vbNo Then Exit Sub
    
    '删除特检项目
    Call DelSpeExamProject(ufgContext.SelectionRow)
    
    blnIsUpdate = True
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Function CheckAllowDelSlices(ByVal lngSlicesRow As Long) As Boolean
'检查是否允许删除制片项目
    CheckAllowDelSlices = IIf(ufgContext.Text(lngSlicesRow, gstrRequest_Slices_当前状态) = "已申请", True, False)
    
    If Not CheckAllowDelSlices Then
        Call MsgBoxD(Me, "该制片项目已被接受或完成，不能执行删除操作。", vbOKOnly, Me.Caption)
    End If
End Function


Private Sub DelSlicesProject(ByVal lngSlicesRow As Long)
'删除制片项目
    Dim strSql As String
    
    strSql = "Zl_病理申请_制片项目_删除(" & ufgContext.KeyValue(lngSlicesRow) & ")"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    Call ufgContext.RemoveRow(lngSlicesRow)
End Sub


Private Sub Menu_Edit_DelSlicesPro()
'删除制片项目
On Error GoTo errHandle
    If Not ufgContext.IsSelectionRow Then
        Call MsgBoxD(Me, "请选择需要删除的制片项目。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgContext.IsEmptyKey(ufgContext.SelectionRow) Then
        Call MsgBoxD(Me, "请选择有效的制片项目。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If Not CheckAllowDelSpeExam(ufgContext.SelectionRow) Then Exit Sub
    
    If MsgBoxD(Me, "确认要删除该制片项目吗？", vbYesNo, Me.Caption) = vbNo Then Exit Sub
    
    '删除特检项目
    Call DelSlicesProject(ufgContext.SelectionRow)
    
    blnIsUpdate = True
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Menu_Edit_NewSE()
'补做，与添加新的特检项目相同，只不过制作类型为补做
On Error GoTo errHandle
    If Not ufgRequest.IsSelectionRow Then
        Call MsgBoxD(Me, "请选择所属的申请记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgRequest.IsEmptyKey(ufgRequest.SelectionRow) Then
        Call MsgBoxD(Me, "请选择有效的申请记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    '补做特检项目
    Call AddSpeExamProject(ufgRequest.KeyValue(ufgRequest.SelectionRow), True)
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub





Private Function GetRedoCount(ByVal strMaterialId As String, ByVal strAntibodyName As String)
'获取特检项目的重做次数
    Dim i As Long
    Dim lngCount As Long
    
    lngCount = 0
    For i = 1 To ufgContext.GridRows - 1
        If Not ufgContext.IsEmptyKey(i) Then
            If Val(ufgContext.Text(i, gstrRequest_SpeExam_材块号)) = Val(strMaterialId) And _
                UCase(ufgContext.Text(i, gstrRequest_SpeExam_抗体名称)) = UCase(strAntibodyName) Then
                If lngCount < GetNumber(ufgContext.Text(i, gstrRequest_SpeExam_制作类型)) Then lngCount = GetNumber(ufgContext.Text(i, gstrRequest_SpeExam_制作类型))
            End If
        End If
    Next i
    
    GetRedoCount = lngCount

End Function


Private Sub RedoSpeExamProject(ByVal lngSpeExamRow As Long)
'项目重做
    Dim lngNewRow As Long
    Dim lngRedoCount As Long
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    strSql = "select Zl_病理申请_特检项目_重做([1]) as 返回值 from dual"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(ufgContext.KeyValue(lngSpeExamRow)))
    
    If rsData.RecordCount <= 0 Then
        Call err.Raise(0, "RedoSpeExamProject", "未成功获取重做后的特检项目ID,处理失败。")
        Exit Sub
    End If
    
    '增加特检记录到列表
    lngNewRow = ufgContext.NewRow
    
    lngRedoCount = GetRedoCount(ufgContext.Text(lngSpeExamRow, gstrRequest_SpeExam_材块号), _
                                ufgContext.Text(lngSpeExamRow, gstrRequest_SpeExam_抗体名称))
    
    '复制行数据
    Call ufgContext.CopyRowData(lngSpeExamRow, lngNewRow)
    
    ufgContext.Text(lngNewRow, gstrRequest_SpeExam_ID) = rsData!返回值
    ufgContext.Text(lngNewRow, gstrRequest_SpeExam_制作类型) = "第" & lngRedoCount + 1 & "次重做"
    ufgContext.Text(lngNewRow, gstrRequest_SpeExam_当前状态) = "已申请"
    ufgContext.Text(lngNewRow, gstrRequest_SpeExam_操作人) = ""
    ufgContext.Text(lngNewRow, gstrRequest_SpeExam_完成时间) = ""
    ufgContext.Text(lngNewRow, gstrRequest_SpeExam_项目结果) = ""
    
    Call ufgContext.LocateRow(lngNewRow)
End Sub



Private Function GetAntibodyUseCount(ByVal lngAntibodyId As String) As Long
'获取抗体的可用份数
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    GetAntibodyUseCount = 0
    
    strSql = "select 使用人份-已用人份 as 可用人份 from 病理抗体信息 where 抗体ID=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngAntibodyId)
    
    If rsData.RecordCount > 0 Then GetAntibodyUseCount = Val(Nvl(rsData!可用人份))
End Function

Private Sub Menu_Edit_RedoSE()
'特检项目重做
On Error GoTo errHandle
    If Not ufgContext.IsSelectionRow Then
        Call MsgBoxD(Me, "请选择需要重做的特检项目。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgContext.IsEmptyKey(ufgContext.SelectionRow) Then
        Call MsgBoxD(Me, "请选择有效的特检项目。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgContext.Text(ufgContext.SelectionRow, gstrRequest_SpeExam_当前状态) <> "已完成" Then
        Call MsgBoxD(Me, "该项目尚未完成，不能进行重做。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If GetAntibodyUseCount(Val(ufgContext.Text(ufgContext.SelectionRow, gstrRequest_SpeExam_抗体ID))) <= 0 Then
        If MsgBoxD(Me, "抗体 [" & ufgContext.Text(ufgContext.SelectionRow, gstrRequest_SpeExam_抗体名称) & "] 已无可用人份，是否继续添加该项目？", vbYesNo, Me.Caption) <> vbYes Then
            Exit Sub
        End If
    End If
    
    '特检项目重做
    Call RedoSpeExamProject(ufgContext.SelectionRow)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then
    Resume
    End If
End Sub


Private Sub ShowSlicesRequestWindow(Optional ByVal lngCurRequestId As Long = -1)
'显示制片申请窗口
Dim frmSlicesRequest As New frmPatholRequisition_Slices
On Error GoTo errFree
    
    Call frmSlicesRequest.ShowSlicesRequestWindow(ufgRequest, ufgContext, mrecStudyInf.lngPatholAdviceId, lngCurRequestId, mlngRequestType, Me)
    
    blnIsUpdate = frmSlicesRequest.blnIsOk
errFree:
    Call Unload(frmSlicesRequest)
    Set frmSlicesRequest = Nothing
End Sub


Private Sub ShowSupMaterialRequestWindow()
'显示补取申请窗口
    Dim frmSupMateriasRequest As New frmPatholRequisition_SupMaterial
    On Error GoTo errFree
        
        Call frmSupMateriasRequest.ShowSupMaterialWindow(ufgRequest, ufgContext, mrecStudyInf.lngPatholAdviceId, Me)
        
        blnIsUpdate = frmSupMateriasRequest.blnIsOk
errFree:
    Call Unload(frmSupMateriasRequest)
    Set frmSupMateriasRequest = Nothing
    
End Sub



Private Sub Menu_Edit_ReqGet()
'补取材申请
On Error GoTo errHandle
    If Not CheckAllowNewRequest(TRequestType.rtMaterial) Then Exit Sub
     
    Call ShowSupMaterialRequestWindow
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Edit_ReqSlices()
On Error GoTo errHandle
    '显示制片申请
    If Not CheckAllowNewRequest(TRequestType.rtSlices) Then Exit Sub

    Call ShowSlicesRequestWindow
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Function CheckAllowNewRequest(ByVal rtRequestType As TRequestType) As Boolean
'判断是否允许进行新的申请
    Dim strRequestType As String
    Dim i As Integer
    
    CheckAllowNewRequest = True


    strRequestType = "Null"
    '如存在申请类型相同且申请尚未完成时，则不能进行申请
    Select Case rtRequestType
        Case TRequestType.rtMianyi
            strRequestType = "免疫组化"
        Case TRequestType.rtTeran
            strRequestType = "特殊染色"
        Case TRequestType.rtFenzi
            strRequestType = "分子病理"
        Case TRequestType.rtSlices
            strRequestType = "再制片"
        Case TRequestType.rtMaterial
            strRequestType = "补取材"
    End Select
    
    For i = 1 To ufgRequest.GridRows - 1
        If ufgRequest.Text(i, gstrRequisition_当前状态) = "已申请" _
            And ufgRequest.Text(i, gstrRequisition_申请类型) Like "*" & strRequestType & "*" Then
            CheckAllowNewRequest = False
            Exit For
        End If
    Next i

    If Not CheckAllowNewRequest Then
        Call MsgBoxD(Me, "该检查存在尚未完成的申请，不能执行次操作。", vbOKOnly, Me.Caption)
    End If
End Function


Private Sub Menu_Edit_ReqSpecialExam()
On Error GoTo errHandle
    '显示特检申请
    If Not CheckAllowNewRequest(-1) Then Exit Sub
    
    Call ShowSpecialExamRequestWindow
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Form_Load()
On Error GoTo errHandle
    Call InitCommandBars
    
    Call RestoreWinState(Me, App.ProductName)
    
    Call InitFace
    
    '初始化申请列表
    Call InitRequisitionList
    
    Call ChangeControlFace(-1)
    
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
    With cbrMenuBar.CommandBar
        Set cbrControl = .Controls.Add(xtpControlButtonPopup, conMenu_View_ToolBar, "申请"): cbrControl.IconId = 3903
        With cbrControl.CommandBar '二级菜单
            Set cbrPopControl = .Controls.Add(xtpControlButton, TMenuType.mtReqGet, "补取申请(&G)"): cbrPopControl.IconId = 10016
            Set cbrPopControl = .Controls.Add(xtpControlButton, TMenuType.mtReqSlices, "制片申请(&S)"): cbrPopControl.IconId = 10017
            Set cbrPopControl = .Controls.Add(xtpControlButton, TMenuType.mtReqSpecialExam, "特检申请(&T)"): cbrPopControl.IconId = 10018
        End With
        Set cbrControl = .Controls.Add(xtpControlButton, TMenuType.mtDelReq, "撤销申请(&D)"): cbrControl.IconId = 3565
        
        Set cbrControl = .Controls.Add(xtpControlButton, TMenuType.mtAddSpeExamPro, "添加(&A)"): cbrControl.IconId = 4010
        cbrControl.BeginGroup = True
        Set cbrControl = .Controls.Add(xtpControlButton, TMenuType.mtDelSE, "删除(&C)"): cbrControl.IconId = 4008
        Set cbrControl = .Controls.Add(xtpControlButton, TMenuType.mtNewSE, "补做(&N)"): cbrControl.IconId = 3082
        Set cbrControl = .Controls.Add(xtpControlButton, TMenuType.mtRedoSE, "重做(&U)"): cbrControl.IconId = 3945
        
'        Set cbrControl = .Controls.Add(xtpControlButton, TMenuType.mtAddSlicesPro, "增加(&A)"): cbrControl.IconId = 4112
'        cbrControl.BeginGroup = True
'        Set cbrControl = .Controls.Add(xtpControlButton, TMenuType.mtDelSlicesPro, "删除(&C)"): cbrControl.IconId = 4114
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
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "申请"): cbrControl.IconId = 3903
        With cbrControl.CommandBar '二级菜单
            Set cbrPopControl = .Controls.Add(xtpControlButton, TMenuType.mtReqGet, "补取申请(&G)"): cbrPopControl.IconId = 10016
            Set cbrPopControl = .Controls.Add(xtpControlButton, TMenuType.mtReqSlices, "制片申请(&S)"): cbrPopControl.IconId = 10017
            Set cbrPopControl = .Controls.Add(xtpControlButton, TMenuType.mtReqSpecialExam, "特检申请(&T)"): cbrPopControl.IconId = 10018
        End With
        
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtDelReq, "撤销申请"): cbrControl.IconId = 3565
        
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtAddSpeExamPro, "添加"): cbrControl.IconId = 4010
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtDelSE, "删除"): cbrControl.IconId = 4008
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtNewSE, "补做"): cbrControl.IconId = 3082
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtRedoSE, "重做"): cbrControl.IconId = 3945
        
'        Set cbrControl = .Add(xtpControlButton, TMenuType.mtAddSlicesPro, "增加"): cbrControl.IconId = 4112
'        cbrControl.BeginGroup = True
'        Set cbrControl = .Add(xtpControlButton, TMenuType.mtDelSlicesPro, "删除"): cbrControl.IconId = 4114
        
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

Private Sub picRequest_Resize()
On Error Resume Next
    '调整界面布局
    ufgRequest.Left = 120
    ufgRequest.Top = 0
    ufgRequest.Width = picRequest.Width - 240
    ufgRequest.Height = picRequest.Height - 60
End Sub


Private Sub picRequestContext_Resize()
On Error Resume Next
    '调整picRequestContext的内容
    labSubTitle.Left = 120
    labSubTitle.Top = 120
    
    linSubTitle.X1 = 0
    linSubTitle.Y1 = labSubTitle.Top + 90
    linSubTitle.X2 = picRequestContext.Width
    linSubTitle.Y2 = labSubTitle.Top + 90
    
    ufgContext.Left = 120
    ufgContext.Top = 400
    ufgContext.Width = picRequestContext.Width - 240
    ufgContext.Height = picRequestContext.Height - 360
End Sub


Private Sub LoadRequestInf(ByVal lngPatholAdviceId As Long)
'载入申请信息
    Dim strSql As String
    
    strSql = "select 申请ID,申请人,申请类型,补费状态,申请细目,申请时间,申请状态,申请描述,完成时间 from 病理申请信息 where 病理医嘱ID=[1] and (申请类型=-1 "
'    If mblnMoved Then strSql = GetMovedDataSql(strSql)
    
    If CheckPopedom(mstrPrivs, "特检申请") Then strSql = strSql & " or 申请类型<=2"
    If CheckPopedom(mstrPrivs, "制片申请") Then strSql = strSql & " or 申请类型=3"
    If CheckPopedom(mstrPrivs, "补取申请") Then strSql = strSql & " or 申请类型=4"
    
    strSql = strSql & ")order by 申请类型,申请时间"
    
    Set ufgRequest.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPatholAdviceId)
    
    Call ufgRequest.RefreshData
End Sub


Private Sub LoadSpeExamRequestContext(ByVal lngRequestId As Long)
'读取特检申请内容
    Dim strSql As String
    
    strSql = "select a.ID,a.材块ID,b.序号,b.标本名称,c.抗体ID, b.标本名称,c.抗体名称,a.制作类型,a.当前状态,a.项目结果,a.完成时间,a.特检医师 " & _
                " from 病理特检信息 a, 病理取材信息 b, 病理抗体信息 c " & _
                " where a.材块id = b.材块id and a.抗体id=c.抗体id and a.申请id=[1] order by a.制作类型, a.材块ID, c.抗体名称"
'    If mblnMoved Then strSql = GetMovedDataSql(strSql)
    
    Set ufgContext.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngRequestId)
    
    Call ufgContext.RefreshData
End Sub




Private Sub LoadSlicesRequestContext(ByVal lngRequestId As Long)
'读取制片申请内容
    Dim strSql As String
    
    strSql = "select a.ID,a.材块ID,b.序号,b.标本名称,a.制片类型,a.制片方式,a.制片数,a.当前状态,a.制片时间,a.制片人 " & _
            " from 病理制片信息 a, 病理取材信息 b " & _
            " where a.材块id=b.材块id and a.申请id=[1] order by a.当前状态, b.标本名称,a.材块ID"
'    If mblnMoved Then strSql = GetMovedDataSql(strSql)
    
    Set ufgContext.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngRequestId)
    
    Call ufgContext.RefreshData
End Sub




Private Sub LoadSupMaterialRequestContext(ByVal lngRequestId As Long)
'读取取材的完成内容
    Dim strSql As String
    
    strSql = "select 材块ID,序号,标本名称,标本量,蜡块数,取材时间,主取医师,副取医师,记录医师 " & _
            " from 病理取材信息 where  申请id=[1] order by 取材时间"
'    If mblnMoved Then strSql = GetMovedDataSql(strSql)
    
    Set ufgContext.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngRequestId)
    
    Call ufgContext.RefreshData
End Sub


Private Sub ShowAntibodyInf(ByVal lngAntibodyRow As Long)
'显示抗体明细信息
    Dim frmAntibodyInf As New frmPatholRequisition_AntibodyInf
    On Error GoTo errFree
        Call frmAntibodyInf.ShowAntibodyInf(ufgContext.Text(lngAntibodyRow, gstrRequest_SpeExam_抗体ID), Me)
errFree:
    Call Unload(frmAntibodyInf)
    Set frmAntibodyInf = Nothing
    
End Sub



Private Sub ufgContext_OnCellButtonClick(ByVal Row As Long, ByVal Col As Long)
On Error GoTo errHandle
    Call ShowAntibodyInf(Row)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub



Private Sub ufgRequest_OnClick()
'读取申请内容
On Error GoTo errHandle
    Dim strRequestType As String
    
    '清除申请项目明细
    Call ufgContext.ClearListData
    
    If Not ufgRequest.IsSelectionRow Then Exit Sub
    If ufgRequest.IsEmptyKey(ufgRequest.SelectionRow) Then Exit Sub
    
    strRequestType = ufgRequest.Text(ufgRequest.SelectionRow, gstrRequisition_申请类型)
    
    Select Case strRequestType
        Case "免疫组化", "分子病理", "特殊染色"
        
            Call InitRequestContextList(0)
            Call ChangeControlFace(0)
            
            '读取特检项目明细
            Call LoadSpeExamRequestContext(ufgRequest.KeyValue(ufgRequest.SelectionRow))
            
            Case "再制片", "重切", "深切", "连切", "白片", "重染", "薄片"
            
            Call InitRequestContextList(3)
            Call ChangeControlFace(3)
             
            '读取制片项目明细
            Call LoadSlicesRequestContext(ufgRequest.KeyValue(ufgRequest.SelectionRow))

        Case "重取材", "补取材"
            
            Call InitRequestContextList(4)
            Call ChangeControlFace(4)
            
            '读取取材项目明细
            Call LoadSupMaterialRequestContext(ufgRequest.KeyValue(ufgRequest.SelectionRow))
            
    End Select
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgRequest_OnColsNameReSet()
On Error GoTo errHandle

   '读取申请信息
    Call LoadRequestInf(mrecStudyInf.lngPatholAdviceId)
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub ufgRequest_OnSelChange()
    ufgRequest_OnClick
End Sub
