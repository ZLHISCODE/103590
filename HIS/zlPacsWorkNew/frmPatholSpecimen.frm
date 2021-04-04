VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPatholSpecimen 
   Caption         =   "标本核收"
   ClientHeight    =   8805
   ClientLeft      =   75
   ClientTop       =   405
   ClientWidth     =   10140
   Icon            =   "frmPatholSpecimen.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8805
   ScaleWidth      =   10140
   StartUpPosition =   3  '窗口缺省
   Begin MSComctlLib.ImageList imgTbrS 
      Left            =   1845
      Top             =   795
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholSpecimen.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholSpecimen.frx":0C7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholSpecimen.frx":18F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholSpecimen.frx":2562
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholSpecimen.frx":31D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholSpecimen.frx":3E46
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholSpecimen.frx":4AB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholSpecimen.frx":572A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholSpecimen.frx":639C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholSpecimen.frx":700E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "历史核收记录"
      Height          =   2175
      Left            =   120
      TabIndex        =   2
      Top             =   5475
      Width           =   9855
      Begin RichTextLib.RichTextBox txtHistoryRecord 
         Height          =   1815
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   3201
         _Version        =   393217
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         Appearance      =   0
         TextRTF         =   $"frmPatholSpecimen.frx":7788
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "送检标本记录"
      Height          =   4455
      Left            =   30
      TabIndex        =   1
      Top             =   675
      Width           =   9855
      Begin VB.ListBox lstPartment 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2790
         ItemData        =   "frmPatholSpecimen.frx":7825
         Left            =   7320
         List            =   "frmPatholSpecimen.frx":7827
         TabIndex        =   7
         Top             =   1080
         Width           =   2415
      End
      Begin VB.ComboBox cbxSpecimentPartment 
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
         Left            =   7320
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   600
         Width           =   2415
      End
      Begin zl9PACSWork.ucFlexGrid ufgData 
         Height          =   3735
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   6588
         DefaultCols     =   ""
         GridRows        =   21
         IsCopyAdoMode   =   0   'False
         IsEjectConfig   =   -1  'True
         HeadFontCharset =   134
         HeadFontWeight  =   400
         DataFontCharset =   134
         DataFontWeight  =   400
         ExtendLastCol   =   -1  'True
      End
      Begin VB.Label labSpecimenName 
         Caption         =   "标本部位名称选择："
         Height          =   255
         Left            =   7440
         TabIndex        =   5
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label labRecordInf 
         Caption         =   "标本总数：0"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   4080
         Width           =   3375
      End
   End
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   240
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmPatholSpecimen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IWorkMenu


Private Const M_STR_HINT_NoSelectData As String = "无效的检查数据，请选择需要执行的检查记录。"
Private Const M_STR_MODULE_MENU_TAG As String = "标本"

Private WithEvents zlReport As zl9Report.clsReport
Attribute zlReport.VB_VarHelpID = -1

Private mlngModule As Long
Private mstrPrivs As String              '模块权限
Private mlngCurDeptId As Long          '当前科室
Private mobjOwner As Object

Private mlngAdviceID As Long
Private mlngSendNo As Long
Private mblnMoved As Boolean
Private mlngStudyState As Long

Private mblnReadOnly As Boolean

Private mlngTmpAdviceId As Long
Private mlngTmpSendNo As Long


Private mrsSpecimenPartData As ADODB.Recordset


Private mObjActiveMenuBar As CommandBar

Private mblnRefreshState As Boolean
Private mbytFontSize As Byte '字号    9--小字体    12--大字体
Private mstrFormats As String 'rtf格式，用于改变字号
Private mblLordingOrRefreshing As Boolean '是否正在加载或者刷新

Private mblnShowSentInfo As Boolean    '是否启用显示送检信息



'获取菜单接口对象
Property Get zlMenu() As IWorkMenu
    Set zlMenu = Me
End Property


Public Sub NotificationRefresh()
'通知刷新
    mblnRefreshState = False
End Sub

Private Sub cbrMain_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error GoTo ErrorHand
    Select Case control.ID
        Case conMenu_PatholSpecimen_PreviewLab, conMenu_PatholSpecimen_LAB
            Call PrintSpecimenLabel(False)
            
        Case conMenu_PatholSpecimen_PrintLab
            Call PrintSpecimenLabel(True)
            
        Case conMenu_PatholSpecimen_PreviewAccept, conMenu_PatholSpecimen_ACP
            Call PrintAcceptNotification(False)
            
        Case conMenu_PatholSpecimen_PrintAccept
            Call PrintAcceptNotification(True)
        
        Case conMenu_PatholSpecimen_Get
            '调用自动提取信息方法
            Call AutoGetSpecimenInf
            
        Case conMenu_PatholSpecimen_Del
            '删除标本
            Call DelSelectionSpecimen
            
        Case conMenu_PatholSpecimen_Save
            '保存标本
            Call SaveCurSpecimenInf
            
        Case conMenu_PatholSpecimen_Accept
            '标本接收
            Call SpecimenAccept
            
        Case conMenu_PatholSpecimen_Reject
            '标本拒收
            Call SpecimenReject
            
        Case conMenu_PatholSpecimen_Cancel
            '标本回退
            Call CancelSelectionSpecimen
    End Select
    Exit Sub
ErrorHand:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cbrMain_Update(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error GoTo ErrorHand
    Dim blnIsAllowAccept As Boolean
    
    blnIsAllowAccept = CheckPopedom(mstrPrivs, "标本核收") And mlngAdviceID > 0
    
    Select Case control.ID
        Case conMenu_PatholSpecimen_ACP, conMenu_PatholSpecimen_LAB, conMenu_PatholSpecimen_PreviewLab, _
        conMenu_PatholSpecimen_PreviewAccept, conMenu_PatholSpecimen_PrintLab, conMenu_PatholSpecimen_PrintAccept
            control.Enabled = blnIsAllowAccept
                   
        Case conMenu_PatholSpecimen_Del, conMenu_PatholSpecimen_Save, conMenu_PatholSpecimen_Accept, _
        conMenu_PatholSpecimen_Reject, conMenu_PatholSpecimen_Cancel

            control.Enabled = blnIsAllowAccept And Not mblnReadOnly
    End Select
    Exit Sub
ErrorHand:
    If ErrCenter() = 1 Then Resume
End Sub


'接口实现部分*********************************************************************************

Public Function IWorkMenu_zlIsModuleMenu(ByVal objControlMenu As XtremeCommandBars.ICommandBarControl) As Boolean
'判断菜单是否属于该模块菜单
    IWorkMenu_zlIsModuleMenu = IIf(objControlMenu.Category = M_STR_MODULE_MENU_TAG, True, False)
End Function


Public Sub IWorkMenu_zlCreateMenu(objMenuBar As Object)
'创建影像记录对应的菜单
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrPopControl As CommandBarControl
    Dim cbrControl As CommandBarControl
    
    Dim str3DFuncs() As String
    Dim i As Long
    Dim lng3DFunc As Long
    
    
    Set mObjActiveMenuBar = objMenuBar.ActiveMenuBar
    
    If Not HasMenu(objMenuBar, conMenu_PatholSpecimen) Then
        Set cbrMenuBar = mObjActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_PatholSpecimen, "标本(&P)", GetPatholMenuIndex(objMenuBar) + 1, False)
        cbrMenuBar.ID = conMenu_PatholSpecimen
        cbrMenuBar.Category = M_STR_MODULE_MENU_TAG
        
        With cbrMenuBar.CommandBar
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlPopup, conMenu_PatholSpecimen_LAB, "标签打印(&L)", "", 1, False)
            With cbrControl
                Set cbrPopControl = CreateModuleMenu(.CommandBar.Controls, xtpControlButton, conMenu_PatholSpecimen_PreviewLab, "预览(V)", "", 1, False)
                Set cbrPopControl = CreateModuleMenu(.CommandBar.Controls, xtpControlButton, conMenu_PatholSpecimen_PrintLab, "打印(P)", "", 1, False)
            End With
                
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlPopup, conMenu_PatholSpecimen_ACP, "凭单打印(&A)", "", 1, False)
            With cbrControl
                Set cbrPopControl = CreateModuleMenu(.CommandBar.Controls, xtpControlButton, conMenu_PatholSpecimen_PreviewAccept, "预览(V)", "", 1, False)
                Set cbrPopControl = CreateModuleMenu(.CommandBar.Controls, xtpControlButton, conMenu_PatholSpecimen_PrintAccept, "打印(P)", "", 1, False)
            End With
            
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholSpecimen_Get, "标本提取(&G)", "", 1, True)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholSpecimen_Del, "标本删除(&D)", "", 1, False)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholSpecimen_Save, "标本保存(&S)", "", 1, True)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholSpecimen_Accept, "标本接收(&R)", "", 0, True)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholSpecimen_Reject, "标本拒收(&J)", "", 0, False)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholSpecimen_Cancel, "标本回退(&H)", "", 0, True)
            
        End With
    End If
End Sub


Public Sub IWorkMenu_zlCreateToolBar(objToolBar As Object)
'创建工具栏
    Exit Sub
End Sub

Public Sub IWorkMenu_zlClearMenu()
'清除所创建的菜单
    Exit Sub
End Sub


Public Sub IWorkMenu_zlClearToolBar()
'清除创建的工具栏
    Exit Sub
End Sub


Public Sub IWorkMenu_zlExecuteMenu(ByVal lngMenuId As Long)
'根据菜单ID执行对应功能
    Dim objCbrControl As XtremeCommandBars.CommandBarControl
    
    Select Case lngMenuId
        Case conMenu_PatholSpecimen_PreviewLab      '预览标签
            Call PrintSpecimenLabel(False)
            
        Case conMenu_PatholSpecimen_PrintLab        '打印标签
            Call PrintSpecimenLabel(True)
            
        Case conMenu_PatholSpecimen_PreviewAccept    '核收单预览
            Call PrintAcceptNotification(False)
            
        Case conMenu_PatholSpecimen_PrintAccept     '核收单打印
            Call PrintAcceptNotification(True)
        
        Case conMenu_PatholSpecimen_Get             '材块提取
            Call AutoGetSpecimenInf
            
        Case conMenu_PatholSpecimen_Del             '删除选择的标本
            Call DelSelectionSpecimen
            
        Case conMenu_PatholSpecimen_Save            '保存当前标本信息
            Call SaveCurSpecimenInf
            
        Case conMenu_PatholSpecimen_Accept          '标本核收
            Call SpecimenAccept
            
        Case conMenu_PatholSpecimen_Reject          '标本拒收
            Call SpecimenReject
            
        Case conMenu_PatholSpecimen_Cancel        '标本回退
            Call CancelSelectionSpecimen
    End Select
End Sub


Public Sub IWorkMenu_zlUpdateMenu(ByVal control As XtremeCommandBars.ICommandBarControl)
'更新菜单
    Dim blnIsAllowAccept As Boolean

    If Not Me.Visible Then
        control.Enabled = False
        Exit Sub
    End If
    
    blnIsAllowAccept = CheckPopedom(mstrPrivs, "标本核收") And mlngAdviceID > 0
    
    Select Case control.ID
        Case conMenu_PatholSpecimen_ACP, conMenu_PatholSpecimen_LAB, conMenu_PatholSpecimen_PreviewLab, _
        conMenu_PatholSpecimen_PreviewAccept, conMenu_PatholSpecimen_PrintLab, conMenu_PatholSpecimen_PrintAccept
            control.Enabled = blnIsAllowAccept
                   
        Case conMenu_PatholSpecimen_Del, conMenu_PatholSpecimen_Save, conMenu_PatholSpecimen_Accept, _
        conMenu_PatholSpecimen_Reject, conMenu_PatholSpecimen_Cancel
            control.Enabled = blnIsAllowAccept And Not mblnReadOnly
    End Select
    
End Sub


Public Sub IWorkMenu_zlPopupMenu(objPopup As XtremeCommandBars.ICommandBar)
'配置右键菜单
    Exit Sub
End Sub

Public Sub IWorkMenu_zlRefreshSubMenu(objMenuBar As Object)
'刷新弹出的子菜单
    Exit Sub
End Sub
'*************************************************************************************************


Private Function CreateModuleMenu(objMenuControl As CommandBarControls, _
    ByVal lngType As XTPControlType, ByVal lngID As Long, ByVal strCaption As String, _
    Optional strToolTip As String = "", Optional lngIconId As Long = 0, Optional blnStartGroup As Boolean = False) As CommandBarControl
'创建该模块内的菜单
    
    Set CreateModuleMenu = objMenuControl.Add(lngType, lngID, strCaption)
    CreateModuleMenu.ID = lngID '如果这里不指定id，则不能将有些菜单添加到右键菜单中
    
    If lngIconId <> 0 Then CreateModuleMenu.IconId = lngIconId
    If blnStartGroup Then CreateModuleMenu.BeginGroup = True
    If strToolTip <> "" Then CreateModuleMenu.ToolTipText = strToolTip
    
    CreateModuleMenu.Category = M_STR_MODULE_MENU_TAG
End Function


Public Sub zlInitModule(ByVal lngModule As Long, ByVal strPrivs As String, ByVal lngDepartId As Long, Optional owner As Object = Nothing)
'初始化模块参数
    mlngModule = lngModule
    mstrPrivs = strPrivs
    mlngCurDeptId = lngDepartId
    
    If Not owner Is Nothing Then Set mobjOwner = owner
End Sub


Public Sub zlUpdateAdviceInf(ByVal lngAdviceID As Long, ByVal lngSendNO As Long, _
    ByVal lngStudyState As Long, ByVal blnMoved As Boolean)
'更新医嘱信息
    mlngAdviceID = lngAdviceID
    mlngSendNo = lngSendNO
    mblnMoved = blnMoved
    mlngStudyState = lngStudyState
    mblnReadOnly = False
    mblnRefreshState = True
    
    '数据被转移时，没有权限时，状态为指定状态时，该模块为只读
    If blnMoved Or lngStudyState = 6 Or lngStudyState = -2 Or lngStudyState = 5 Then
        mblnReadOnly = True
    End If

End Sub

Public Sub ReSetFormFontSize(ByVal bytFontSize As Byte)
'功能:重新设置工作站窗体的字体大小
On Error GoTo ErrHandle

    Dim objCtrl As control
    Dim CtlFont As StdFont
    Dim strFontType  As String
    
    
    mbytFontSize = bytFontSize
    
    
    Set CtlFont = New StdFont
    Me.FontSize = bytFontSize
    strFontType = IIf(IsUseClearType = True, "微软雅黑", "宋体")
    
    CtlFont.Name = strFontType
    CtlFont.Size = bytFontSize
    
    For Each objCtrl In Me.Controls
        Select Case UCase(TypeName(objCtrl))
        Case UCase("Frame")
            objCtrl.Font.Size = bytFontSize
            objCtrl.Font.Name = strFontType
        Case UCase("TabStrip") '页面控件
            objCtrl.Font.Size = bytFontSize
            objCtrl.Font.Name = strFontType
        Case UCase("Label")
            objCtrl.FontSize = bytFontSize
            objCtrl.FontName = strFontType
            objCtrl.Height = TextHeight("测") + 150
        Case UCase("vsFlexGrid")
            objCtrl.Cell(flexcpFont, 0, 0, 0, objCtrl.DataGrid.Cols - 1) = CtlFont
            objCtrl.Font = CtlFont
            objCtrl.RowHeight(0) = TextHeight("测") + 150
         Case UCase("ucFlexGrid")
            objCtrl.DataGrid.Cell(flexcpFont, 0, 0, 0, objCtrl.DataGrid.Cols - 1) = CtlFont
            objCtrl.DataGrid.Font = CtlFont
            objCtrl.DataGrid.RowHeight(0) = TextHeight("测") + 150
        Case UCase("ComboBox")
            objCtrl.FontSize = bytFontSize
            objCtrl.FontName = strFontType
        Case UCase("OptionButton")
            objCtrl.FontSize = bytFontSize
            objCtrl.FontName = strFontType
            objCtrl.Width = TextWidth("测试" & objCtrl.Caption)
        Case UCase("listBox")
            objCtrl.FontSize = bytFontSize
            objCtrl.Font = CtlFont
        Case UCase("CheckBox")
            objCtrl.FontSize = bytFontSize
            objCtrl.FontName = strFontType
            objCtrl.Width = TextWidth("测试" & objCtrl.Caption)
        Case UCase("DTPicker")
            objCtrl.Font.Size = bytFontSize
            objCtrl.FontN.ame = strFontType
            objCtrl.Width = TextWidth("2012-01-01 23:59:59") * 1.25
            objCtrl.Height = TextHeight("测") * 1.5
        Case UCase("textBox")
            objCtrl.FontSize = bytFontSize
            objCtrl.FontName = strFontType
        Case UCase("ReportControl")
            Set objCtrl.PaintManager.CaptionFont = CtlFont
            Set objCtrl.PaintManager.TextFont = CtlFont
            objCtrl.Redraw
        Case UCase("DockingPane")
            Set objCtrl.PaintManager.CaptionFont = CtlFont
        Case UCase("CommandBars")
            Set objCtrl.Options.Font = CtlFont
        Case UCase("TabControl")
            Set objCtrl.PaintManager.Font = CtlFont
        Case UCase("CommandButton")
            objCtrl.FontName = strFontType
            objCtrl.FontSize = bytFontSize
        Case UCase("richtextbox")
        
            If bytFontSize = C_INT_FONTSISE_SMALL Then
                mstrFormats = "{\rtf1\ansi\ansicpg936\deff0\deflang1033\deflangfe2052{\fonttbl{\f0\fnil\fcharset134 \'cb\'ce\'cc\'e5;}}" & _
                        "{\colortbl ;\red255\green104\blue104;\red19\green164\blue251;}" & _
                        "{\*\generator Msftedit 5.41.21.2509;}\viewkind4\uc1\sl276\slmult1\lang2052\b\f0\fs18 "
            ElseIf bytFontSize = C_INT_FONTSISE_MEDIUM Then
                mstrFormats = "{\rtf1\ansi\ansicpg936\deff0\deflang1033\deflangfe2052{\fonttbl{\f0\fnil\fcharset134 \'cb\'ce\'cc\'e5;}}" & _
                        "{\colortbl ;\red255\green104\blue104;\red19\green164\blue251;}" & _
                        "{\*\generator Msftedit 5.41.21.2509;}\viewkind4\uc1\sl276\slmult1\lang2052\b\f0\fs24 "
            ElseIf bytFontSize = C_INT_FONTSISE_BIG Then
                mstrFormats = "{\rtf1\ansi\ansicpg936\deff0\deflang1033\deflangfe2052{\fonttbl{\f0\fnil\fcharset134 \'cb\'ce\'cc\'e5;}}" & _
                        "{\colortbl ;\red255\green104\blue104;\red19\green164\blue251;}" & _
                        "{\*\generator Msftedit 5.41.21.2509;}\viewkind4\uc1\sl276\slmult1\lang2052\b\f0\fs30 "
            End If
            
            txtHistoryRecord.Text = ""
            Call LoadSpecimenAcceptOrRejectHistoryData
        End Select
    Next
    
    Call AdjustFace
    
    Exit Sub
ErrHandle:
End Sub




Public Sub zlRefreshFace(Optional blnForceRefresh As Boolean = False)
'刷新界面数据
    Dim lngNewAdviceId As Long

    If (mlngTmpAdviceId = mlngAdviceID And mlngTmpSendNo = mlngSendNo And mblnRefreshState) And Not blnForceRefresh Then Exit Sub
    
    lngNewAdviceId = mlngAdviceID
    mblnRefreshState = True
    
    If mlngTmpAdviceId <> mlngAdviceID And mlngTmpAdviceId > 0 Then
        '判断取材是否需要进行确认
        If IsNeedSaveSpecimen Then
            If MsgBoxD(Me, "尚未对录入的标本进行保存，是否需要保存？", vbYesNo, Me.Caption) = vbYes Then
                mlngAdviceID = mlngTmpAdviceId
                
                Call SaveCurSpecimenInf
            End If
        End If
    End If
        
    mlngAdviceID = lngNewAdviceId
    mlngTmpAdviceId = lngNewAdviceId
    mlngTmpSendNo = mlngSendNo
    
    If mlngAdviceID <= 0 Then
        Call ConfigSpecimenFace(False, "医嘱ID无效请检查。")
        Exit Sub
    Else
        Call ConfigSpecimenFace(True)
    End If
    
    mblLordingOrRefreshing = True
    '载入标本数据
    Call LoadSpecimenData
    
    
    '读取标本核收记录
    txtHistoryRecord.Text = ""
    Call LoadSpecimenAcceptOrRejectHistoryData
    
    '刷新标本数量
    Call RefreshSpecimenCount
    
    Call ConfigPopedom(mblnReadOnly)
    mblLordingOrRefreshing = False
'    If Not (owner Is Nothing) Then
'        Call Me.Show(1, owner)
'    End If
    
End Sub

Public Sub zlRefresh(lngAdviceID As Long, ByVal blnReadOnly As Boolean, _
    strPrivs As String, ByVal blnMoved As Boolean, _
    ByVal lngCurDepartmentId As Long, Optional owner As Form = Nothing)
    
    If lngAdviceID <= 0 Then
        Call ConfigSpecimenFace(False, "医嘱ID无效请检查。")
        Exit Sub
    Else
        Call ConfigSpecimenFace(True)
    End If
    
    
    If lngAdviceID <> mlngAdviceID And mlngAdviceID > 0 Then
        '判断取材是否需要进行确认
        If IsNeedSaveSpecimen Then
            If MsgBoxD(Me, "尚未对标本进行保存操作，是否需要保存？", vbYesNo, Me.Caption) = vbYes Then
                Call SaveCurSpecimenInf
            End If
        End If
    End If
    
    
'    If mlngCurAdviceId = lngAdviceID Then Exit Sub
    
    mlngAdviceID = lngAdviceID
    mstrPrivs = strPrivs
    mblnMoved = blnMoved
    mlngCurDeptId = lngCurDepartmentId
    
'    mlngStudyProcedure = GetStudyProcedure
    
    mblLordingOrRefreshing = True
    '载入标本数据
    Call LoadSpecimenData
    
    
    '读取标本核收记录
    txtHistoryRecord.Text = ""
    Call LoadSpecimenAcceptOrRejectHistoryData
    
    '刷新标本数量
    Call RefreshSpecimenCount
    
    Call ConfigPopedom(blnReadOnly)
    mblLordingOrRefreshing = False
    If Not (owner Is Nothing) Then
        Call Me.Show(1, owner)
    End If
End Sub


Public Sub LoadSpecimenAcceptOrRejectHistoryData()
    Dim strSql As String
    Dim rsHistory As ADODB.Recordset
    Dim strRecord As String
    Dim lngStart As Long
    Dim strFormats As String
    
    strSql = "select 送检单位,送检科室,送检人,送检日期,联系方式,登记人,核收状态,拒收原因,通知人,备注 from 病理送检信息 where 医嘱ID=[1] and" _
               & " 送检日期<>to_date('1000/10/10 10:10:10','yyyy/mm/dd hh24:mi:ss')  order by 送检日期 "
'    If mblnMoved Then strSql = GetMovedDataSql(strSql)
    
    Set rsHistory = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngAdviceID)
    
    If rsHistory.RecordCount <= 0 Then Exit Sub
               
    strFormats = mstrFormats
    txtHistoryRecord.Text = ""
    While Not rsHistory.EOF
        If Val(Nvl(rsHistory!核收状态)) = 1 Then
            strRecord = Nvl(rsHistory!送检日期) & "：由[ " & Nvl(rsHistory!送检人) & " ]从[ " & Nvl(rsHistory!送检单位) & Nvl(rsHistory!送检科室) & " ]送检的标本已被[ " & Nvl(rsHistory!登记人) & " ]核收。"
            
            strFormats = strFormats & "\cf2 " & strRecord & "\par"
        Else
            strRecord = Nvl(rsHistory!送检日期) & "：由[ " & Nvl(rsHistory!送检人) & " ]从[ " & Nvl(rsHistory!送检单位) & Nvl(rsHistory!送检科室) & " ]送检的标本已被[ " & Nvl(rsHistory!登记人) & " ]拒收。已通知[ " & Nvl(rsHistory!通知人) & " ] 联系方式[ " & Nvl(rsHistory!联系方式) & " ]"
            
            strFormats = strFormats & "\cf1 " & strRecord & "\par"
        End If
        
        rsHistory.MoveNext
    Wend
    
    txtHistoryRecord.SelRTF = strFormats & "}"
End Sub


Private Sub ConfigPopedom(ByVal blnIsReadOnly As Boolean)
'配置权限
    Dim blnIsAllowAccept As Boolean
    
    blnIsAllowAccept = CheckPopedom(mstrPrivs, "标本核收")
    
    ufgData.ReadOnly = blnIsReadOnly
    
    
    lstPartment.Enabled = blnIsAllowAccept
    cbxSpecimentPartment.Enabled = blnIsAllowAccept
    
    If blnIsReadOnly Then
        cbxSpecimentPartment.BackColor = Me.BackColor
        lstPartment.BackColor = Me.BackColor
    Else
        cbxSpecimentPartment.BackColor = vbWhite
        lstPartment.BackColor = vbWhite
    End If

End Sub



Private Sub LoadSpecimenData()
'读取接收的标本信息
    Dim strSql As String
    Dim rsSpecimen As ADODB.Recordset
    
    strSql = "select 标本ID,送检ID,标本名称,标本类型,采集部位,数量,材料类别,存放位置,原有编号,接收日期,备注,case when nvl(送检ID,0)<=0 then '未核收' else '已核收' end as 核收状态 " & _
             "from 病理标本信息 where 医嘱id=[1] order by 标本类型,材料类别,接收日期,标本ID"
'    If mblnMoved Then strSql = GetMovedDataSql(strSql)
    
    Set ufgData.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngAdviceID)
    
    Call ufgData.RefreshData
End Sub



Private Sub AdjustFace()
    '调整界面布局
    Frame1.Left = 0
    If mbytFontSize = C_INT_FONTSISE_SMALL Then
        Frame1.Top = 800
    ElseIf mbytFontSize = C_INT_FONTSISE_MEDIUM Then
        Frame1.Top = 850
    Else
        Frame1.Top = 900
    End If
    Frame1.Width = Me.Width - 0
    Frame1.Height = Me.Height - Frame2.Height - 1000
    
    ufgData.Left = 120
    ufgData.Top = 240 + IIf(mbytFontSize = C_INT_FONTSISE_SMALL, 0, 120)
    ufgData.Width = Frame1.Width - lstPartment.Width - 360
    ufgData.Height = Frame1.Height - labRecordInf.Height - 480
    
    labRecordInf.Left = 120
    labRecordInf.Top = Frame1.Height - labRecordInf.Height - 120 + IIf(mbytFontSize = C_INT_FONTSISE_SMALL, 0, 85)


    
    '调整frame2的内容
    Frame2.Left = 0
    Frame2.Top = Frame1.Top + Frame1.Height + 120
    Frame2.Width = Frame1.Width
    
    txtHistoryRecord.Left = 120
    txtHistoryRecord.Top = 240 + IIf(mbytFontSize = C_INT_FONTSISE_SMALL, 0, 120)
    txtHistoryRecord.Width = Frame2.Width - 240
    txtHistoryRecord.Height = Frame2.Height - 360 + IIf(mbytFontSize = C_INT_FONTSISE_SMALL, 0, -120)
    
    labSpecimenName.Left = ufgData.Left + ufgData.Width + 120
    labSpecimenName.Top = ufgData.Top + IIf(mbytFontSize = C_INT_FONTSISE_SMALL, 0, 120)

    cbxSpecimentPartment.Left = labSpecimenName.Left
    cbxSpecimentPartment.Top = labSpecimenName.Top + labSpecimenName.Height + 120

    lstPartment.Left = labSpecimenName.Left
    lstPartment.Top = cbxSpecimentPartment.Top + cbxSpecimentPartment.Height + 120
    lstPartment.Height = ufgData.Height - labSpecimenName.Height - cbxSpecimentPartment.Height - 240
    
    
End Sub




Private Sub RefreshSpecimenCount()
    '刷新标本数量
    Dim Count As Long
    Dim lngTotal As Long
    Dim i As Long
    
    lngTotal = 0
    
    For i = 1 To ufgData.GridRows - 1
        If Not ufgData.RowHidden(i) Then
            If Not ufgData.IsNullRow(i) Then
                lngTotal = lngTotal + Val(ufgData.Text(i, gSpecimen_数量))
            End If
        End If
    Next i
    
    labRecordInf.Caption = "标本总数：" & lngTotal
End Sub




Private Function IsNeedSaveSpecimen() As Boolean
'是否需要取材确认
    Dim i As Long
    
    IsNeedSaveSpecimen = False
    
    For i = 1 To ufgData.GridRows - 1
        If ufgData.RowState(i) = TDataRowState.Add And Not ufgData.RowHidden(i) Then
            IsNeedSaveSpecimen = True
            Exit For
        End If
    Next i
End Function



Private Sub ConfigSpecimenFace(ByVal blnIsValid As Boolean, Optional ByVal strHintInf As String = "")
'配置核收界面
    
    lstPartment.Enabled = blnIsValid
    cbxSpecimentPartment.Enabled = blnIsValid
    
    If blnIsValid Then
        Call ufgData.CloseHintInf
        
        cbxSpecimentPartment.BackColor = Me.BackColor
        lstPartment.BackColor = Me.BackColor
    Else
        Call ufgData.ShowHintInf(strHintInf)
        labRecordInf.Caption = ""
        
        cbxSpecimentPartment.BackColor = vbWhite
        lstPartment.BackColor = vbWhite
    End If
End Sub



Private Sub cbxSpecimentPartment_Click()
On Error GoTo ErrHandle
    Dim i As Integer
    Dim strSql As String
    
    '清空ListBox
    lstPartment.Clear
    
    If Trim(cbxSpecimentPartment.Text) <> "" Then
        mrsSpecimenPartData.Filter = "标本部位='" & cbxSpecimentPartment.Text & "'"
    Else
        mrsSpecimenPartData.Filter = ""
    End If
    
    
    While Not mrsSpecimenPartData.EOF
       '加载具体检查标本名称
        lstPartment.AddItem Nvl(mrsSpecimenPartData!标本名称)       '移到下一条数据
        mrsSpecimenPartData.MoveNext
    Wend
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub AutoGetSpecimenInf()
'自动从医嘱中提取标本信息
    Dim strSql As String
    Dim lngRow As Long
    Dim i As Long
    Dim blnFind As Boolean
    Dim objCheck As CheckState
    Dim strSpecimenType As String
    Dim lngImgIndex As Long
    Dim rsAdviceRecord As ADODB.Recordset
    
    '已经核收的标本不能进行信息提取
    If Not Val(ufgData.Text(1, gSpecimen_送检ID)) <= 0 Then
        Call MsgBoxD(Me, "标本已被核收，不能进行提取。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    strSql = "select a.标本部位 as 标本名称,a.检查方法,b.标本部位,b.标本类型 from 病人医嘱记录 a,病理检查标本 b where a.标本部位=b.标本名称(+) and 相关ID=[1]"
    Set rsAdviceRecord = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngAdviceID)
    
    If rsAdviceRecord.RecordCount <= 0 Then Exit Sub
    
    Do While Not rsAdviceRecord.EOF
        
        For i = 1 To ufgData.GridRows - 1
            If (ufgData.Text(i, gstrMaterial_标本名称) = Nvl(rsAdviceRecord!标本名称)) And Not ufgData.RowHidden(i) Then GoTo continue
        Next i
        
        lngRow = ufgData.GetNullRowIndex
        
        '给各项赋值
        ufgData.Text(lngRow, gSpecimen_标本名称) = Nvl(rsAdviceRecord!标本名称)
        
        If Trim(ufgData.Text(lngRow, gSpecimen_标本类型)) = "" Then ufgData.Text(lngRow, gSpecimen_标本类型) = "0-手术标本"
        
        Call ufgData.GetFieldDisplayText(gSpecimen_标本类型, Val(Nvl(rsAdviceRecord!标本类型)), blnFind, objCheck, strSpecimenType, lngImgIndex)
        ufgData.Text(lngRow, gSpecimen_标本类型) = Val(Nvl(rsAdviceRecord!标本类型)) & "-" & strSpecimenType
        
        ufgData.Text(lngRow, gSpecimen_采集部位) = Nvl(rsAdviceRecord!标本部位)
        ufgData.Text(lngRow, gSpecimen_数量) = 1
        ufgData.Text(lngRow, gSpecimen_材料类别) = Decode(Nvl(rsAdviceRecord!检查方法), "标本", 0, "蜡块", 1, "玻片", 2, "白片", 3, "其他", 4) & "-" & Nvl(rsAdviceRecord!检查方法)
          
          
        Call ufgData.DataGrid.Select(lngRow, ufgData.GetColIndex(gSpecimen_标本名称))
        Call ufgData.DataGrid.EditCell
        
continue:
        rsAdviceRecord.MoveNext
    Loop
    
End Sub


Private Sub CancelSelectionSpecimen()
'回退标本

    Dim Row As Integer
    Dim ID As String
    Dim strSql As String
    
    If ufgData.ShowingRowCount <= 0 Then Exit Sub
    
    If Not ufgData.IsSelectionRow Then
        Call MsgBoxD(Me, "请选择需要删除的标本记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgData.IsNullRow(ufgData.SelectionRow) Then
        Call MsgBoxD(Me, "请选择需要删除的标本记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    Row = ufgData.SelectionRow
    ID = ufgData.Text(Row, gSpecimen_标本ID)
    
    If ufgData.IsSelectionRow = False Then
        Call MsgBoxD(Me, "请选择需要回退的项目。")
        Exit Sub
    End If
    
    
    If CheckAllowUpdateSpecimen(ID) = False Then
        Call MsgBoxD(Me, "该标本关联了对应的取材记录，不能进行回退，请先删除对应的取材信息。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    strSql = "Zl_病理标本_退回(" & ID & ")"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    Call ufgData.SyncData(Row, gSpecimen_接收日期, Null, True)
    Call ufgData.SyncData(Row, gSpecimen_送检ID, Null, True)
    Call ufgData.SyncText(Row, gSpecimen_核收状态, "未核收", True)
                
End Sub


Private Sub DelSelectionSpecimen()
'删除选择的标本
    If ufgData.ShowingRowCount <= 0 Then Exit Sub
    
    If Not ufgData.IsSelectionRow Then
        Call MsgBoxD(Me, "请选择需要删除的标本记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgData.IsNullRow(ufgData.SelectionRow) Then
        Call MsgBoxD(Me, "请选择需要删除的标本记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    '已经核收的标本不能进行删除。
    If Not Val(ufgData.Text(ufgData.SelectionRow, gSpecimen_送检ID)) <= 0 Then
        Call MsgBoxD(Me, "标本已被核收不能进行删除,请先进行回退处理。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If MsgBoxD(Me, "确认要删除选择的标本数据吗？", vbYesNo, Me.Caption) = vbNo Then Exit Sub
    
    
    '删除行
    Call ufgData.DelCurRow
    
    '保存删除的标本数据（已被核收的标本不能进行删除，核收后送检ID不为空）
    Call SaveSpecimenData(False, True)
    
    '刷新标本数量
    Call RefreshSpecimenCount
End Sub


'Private Sub cmdReload_Click()
'On Error GoTo errHandle
'    '恢复列表数据
'    mclsVFGSpecimen.RestoreList
'
'    '刷新标本数量
'    Call RefreshSpecimenCount
'
'    Exit Sub
'errHandle:
'    If ErrCenter() = 1 Then Resume
'End Sub



Private Function GetPatholNum(ByVal lngAdviceID As Long) As String
'获取病理号等相关信息
    Dim strSql As String
    Dim rsPatholNum As ADODB.Recordset
    
    
    GetPatholNum = ""
    strSql = "select 病理号 from 病理检查信息 where 医嘱id=[1]"
    
    Set rsPatholNum = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngAdviceID)
    
    If rsPatholNum.RecordCount <= 0 Then Exit Function
    
    GetPatholNum = Nvl(rsPatholNum!病理号)
End Function


Private Function CheckNewSpecimenInf() As Boolean
'检查是否有新的需要核收的标本信息
    Dim i As Long
    
    CheckNewSpecimenInf = False
    For i = 1 To ufgData.GridRows - 1
        If Not ufgData.IsNullRow(i) And ufgData.IsEmptyKey(i) Then
            CheckNewSpecimenInf = True
            Exit Function
        End If
    Next i
End Function


Private Sub SpecimenAccept()
'标本核收
    Dim blnValid As Boolean
    
    '标本核收
    If ufgData.ShowingDataRowCount <= 0 Then
        Call MsgBoxD(Me, "没有找到需要核收的标本信息，请检查标本是否正确录入。", vbOKOnly, Me.Caption)
        Exit Sub
    End If

'    '判断是否有需要核收的标本信息
'    If Not CheckNewSpecimenInf() Then
'        Call MsgBoxD(Me, "没有找到需要再次核收的标本信息，请检查标本是否正确录入。", vbOKOnly, Me.Caption)
'        Exit Sub
'    End If
    
    
    blnValid = Not ufgData.IsErrColorWithList
    If Not blnValid Then
        Call MsgBoxD(Me, "检测到标本列表存在无效数据，请确认是否正确完整的录入，“红色”标记的单元格为必录数据。", vbOKOnly, Me.Caption)
        Exit Sub
    End If


    Dim blnIsSucceed As Boolean
    blnIsSucceed = frmPatholSpecimen_AcceptOrReject.ShowAcceptOrRejectSpecimenWindow(mlngAdviceID, _
                                    mlngCurDeptId, txtHistoryRecord, False, Me, mstrPrivs, mblnShowSentInfo)
    
    If blnIsSucceed Then

        '更新核收状态
        Call UpdateAcceptState
    
        '这里执行核收事件
        Call SendMsgToMainWindow(Me, wetSpecimenAccept, mlngAdviceID, GetPatholNum(mlngAdviceID))
        
        Call ufgData.SetMenuState(False)
    End If
    
'    '刷新标本数量
'    Call RefreshSpecimenCount

End Sub


Private Sub UpdateAcceptState()
'更新核收状态
    Dim i As Long
    
    For i = 1 To ufgData.GridRows - 1
        If Not ufgData.IsNullRow(i) Then
            ufgData.Text(i, gSpecimen_核收状态) = "已核收"
        End If
    Next i
End Sub


Private Sub PrintSpecimenLabel(Optional ByVal blnIsPrint As Boolean = True)
'打印特检项目标签
    Dim i As Long
    Dim j As Long
    Dim strValue(5) As String
    
    j = 0
    strValue(0) = "0": strValue(1) = "0": strValue(2) = "0": strValue(3) = "0": strValue(4) = "0": strValue(5) = "0"
    For i = 1 To ufgData.GridRows - 1
        If Not ufgData.IsEmptyKey(i) Then
            If zlCommFun.ActualLen(strValue(j)) > 2000 Then
                j = j + 1
                strValue(j) = ""
            End If
    
            If strValue(j) <> "" Then strValue(j) = strValue(j) & ","
    
            strValue(j) = strValue(j) & ufgData.KeyValue(i)
        End If
    Next i
    
    
    '判断 是否打印参数值 如果值为0 则表示没有勾选 是否直接打印参数
    '最后附加参数:0=缺省值,可不传,表示正常(含报表及预览),1=直接到预览,2=直接打印,3-输出到Excel
    Call zlReport.ReportOpen(gcnOracle, 100, "ZL1_Inside_1294_12", Me, "标本ID1=" & strValue(0), "标本ID2=" & strValue(1), "标本ID3=" & strValue(2), "标本ID4=" & strValue(3), "标本ID5=" & strValue(4), "标本ID6=" & strValue(5), IIf(blnIsPrint, 2, 1))
End Sub


Private Sub PrintSelectSpecimenLabel(Optional ByVal blnIsPrint As Boolean = True)
'打印选择的材块标签
On Error GoTo ErrHandle
    Dim strValue(5) As String
    
    
    If Not ufgData.IsSelectionRow Then
        Call MsgBoxD(Me, "请选择需要打印的标本记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgData.IsNullRow(ufgData.SelectionRow) Then
        Call MsgBoxD(Me, "请选择需要打印的标本记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    strValue(0) = ufgData.KeyValue(ufgData.SelectionRow)
    strValue(1) = "0": strValue(2) = "0": strValue(3) = "0": strValue(4) = "0": strValue(5) = "0"

    '最后附加参数:0=缺省值,可不传,表示正常(含报表及预览),1=直接到预览,2=直接打印,3-输出到Excel
    Call zlReport.ReportOpen(gcnOracle, 100, "ZL1_Inside_1294_12", Me, "标本ID1=" & strValue(0), "标本ID2=" & strValue(1), "标本ID3=" & strValue(2), "标本ID4=" & strValue(3), "标本ID5=" & strValue(4), "标本ID6=" & strValue(5), IIf(blnIsPrint, 2, 1))
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub PrintAcceptNotification(Optional ByVal blnIsPrint As Boolean = True)
'打印标本核收通知单
    '最后附加参数:0=缺省值,可不传,表示正常(含报表及预览),1=直接到预览,2=直接打印,3-输出到Excel
    Call zlReport.ReportOpen(gcnOracle, 100, "ZL1_Inside_1294_13", Me, "医嘱ID=" & mlngAdviceID, IIf(blnIsPrint, 2, 1))
End Sub



Private Sub SpecimenReject()
'拒收标本
    Dim blnIsSucceed As Boolean
    blnIsSucceed = frmPatholSpecimen_AcceptOrReject.ShowAcceptOrRejectSpecimenWindow(mlngAdviceID, "", txtHistoryRecord, True, Me, mstrPrivs, mblnShowSentInfo)

    If blnIsSucceed Then
        '这里可以执行事件...

    End If

End Sub



Public Sub SaveSpecimenData(ByVal blnSetFocus As Boolean, Optional ByVal blnIsSaveOnlyDel As Boolean = False)
'保存标本数据
'blnSetFocus:用于设置一次焦点，处理109548,本窗体调用时设置为TRUE
    Dim i As Long
    Dim strSql As String
    Dim rsReturn As ADODB.Recordset
    Dim lngSpecimenID As Long
    Dim dtServicesTime As String
    
    If blnSetFocus Then ufgData.SetFocus
    
    For i = 1 To ufgData.GridRows - 1
        If ufgData.RowState(i) = TDataRowState.Add And Not blnIsSaveOnlyDel Then
                dtServicesTime = zlDatabase.Currentdate
                
                '添加标本数据
                strSql = "select Zl_病理标本_新增([1],[2],[3],[4],[5],[6],[7],[8],[9],[10]) as 返回值 from dual"
                
                    Set rsReturn = zlDatabase.OpenSQLRecord(strSql, Me.Caption, _
                            mlngAdviceID, _
                            ufgData.Text(i, gSpecimen_标本名称), _
                            Val(ufgData.Text(i, gSpecimen_标本类型)), _
                            ufgData.Text(i, gSpecimen_采集部位), _
                            Val(ufgData.Text(i, gSpecimen_数量)), _
                            Val(ufgData.Text(i, gSpecimen_材料类别)), _
                            ufgData.Text(i, gSpecimen_原有编号), _
                            ufgData.Text(i, gSpecimen_存放位置), _
                            CDate(dtServicesTime), _
                            ufgData.Text(i, gSpecimen_备注) _
                            )
                            
                    If rsReturn.RecordCount <= 0 Then
                        Call err.Raise(0, "SaveSpecimenData", "未成功获取新增后的标本ID,处理失败。")
                        Exit Sub
                    End If
                    
                    lngSpecimenID = rsReturn!返回值
                    
                    '保存新增的标本ID
                    ufgData.Text(i, gSpecimen_标本ID) = lngSpecimenID
                    ufgData.Text(i, gSpecimen_接收日期) = dtServicesTime
                    ufgData.Text(i, gSpecimen_核收状态) = "未核收"
                    
            ElseIf ufgData.RowState(i) = TDataRowState.Del Then
                '删除标本，当标本被核收后，不允许删除
                strSql = "Zl_病理标本_删除(" & Val(ufgData.KeyValue(i)) & ")"
                Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
            ElseIf ufgData.RowState(i) = TDataRowState.Update And Not blnIsSaveOnlyDel Then
                '更新标本
                lngSpecimenID = Val(ufgData.KeyValue(i))
        
                strSql = "Zl_病理标本_更新(" & lngSpecimenID & ",'" & _
                        ufgData.Text(i, gSpecimen_标本名称) & "'," & _
                        Val(ufgData.Text(i, gSpecimen_标本类型)) & ",'" & _
                        ufgData.Text(i, gSpecimen_采集部位) & "'," & _
                        Val(ufgData.Text(i, gSpecimen_数量)) & "," & _
                        Val(ufgData.Text(i, gSpecimen_材料类别)) & ",'" & _
                        ufgData.Text(i, gSpecimen_原有编号) & "','" & _
                        ufgData.Text(i, gSpecimen_存放位置) & "','" & _
                        ufgData.Text(i, gSpecimen_备注) & "')"
        
                Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
        End If
        
        '更新行状态
        ufgData.RowState(i) = TDataRowState.Normal
    Next i
End Sub

Private Sub SaveCurSpecimenInf()
'保存当前标本信息
    Dim blnValid As Boolean
    
    '标本核收
    If ufgData.ShowingDataRowCount <= 0 Then
        Call MsgBoxD(Me, "没有找到需要保存的标本信息，请检查标本是否正确录入。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    blnValid = Not ufgData.IsErrColorWithList
    If Not blnValid Then
        Call MsgBoxD(Me, "检测到标本列表中存在无效数据，请确认是否正确完整的录入，“红色”标记的单元格为必录数据。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    Call SaveSpecimenData(True)
    
    Call SendMsgToMainWindow(Me, wetSpecimenSave, mlngAdviceID)
    
    Call MsgBoxD(Me, "数据已成功保存。", vbOKOnly, Me.Caption)
    
'    '刷新标本数量
'    Call RefreshSpecimenCount
End Sub


Private Sub Form_Initialize()
    Set zlReport = New zl9Report.clsReport
End Sub


Private Sub LoadSpecimenPart()
'加载标本检查部位
    Dim i As Integer
    Dim strSql As String
    Dim rsSpecimenPart As ADODB.Recordset
    
    
    strSql = "select 标本名称,简码,标本部位,标本类型 from 病理检查标本 order by 标本部位,标本名称"
    Set mrsSpecimenPartData = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    
    
    strSql = "select distinct 标本部位 from 病理检查标本 order by 标本部位"
    Set rsSpecimenPart = zlDatabase.OpenSQLRecord(strSql, Me.Caption)

    cbxSpecimentPartment.Clear
    Call cbxSpecimentPartment.AddItem("")
    
    While Not rsSpecimenPart.EOF
       '添加标本检查部位
       cbxSpecimentPartment.AddItem Nvl(rsSpecimenPart!标本部位) '移到下一条数据
       rsSpecimenPart.MoveNext
    Wend
    
    If cbxSpecimentPartment.ListCount > 0 Then cbxSpecimentPartment.ListIndex = 0
End Sub


Private Sub Form_Load()
On Error GoTo ErrHandle
   Dim strTemp As String
   
   '设置行数
    ufgData.GridRows = glngStandardRowCount
    '设置行高
    ufgData.RowHeightMin = glngStandardRowHeight
    ufgData.IsCopyMode = True
    
    Set mrsSpecimenPartData = Nothing
    '判断数据库参数表是否有数据 有则读取数据库参数  没有则加载默认
    strTemp = zlDatabase.GetPara("标本核收列表配置", glngSys, G_LNG_PATHOLSYS_NUM, "")
    mblnShowSentInfo = Val(zlDatabase.GetPara("录入外院信息", glngSys, G_LNG_PATHOLSYS_NUM, 0)) '是否显示送检信息
    
    ufgData.DefaultColNames = gstrSpecimenCols

    If strTemp = "" Then
        ufgData.ColNames = gstrSpecimenCols
    Else
        ufgData.ColNames = strTemp
    End If
    
    ufgData.ColConvertFormat = gstrSpecimenConvertFormat
    
    Call InitCommandBars
    '加载标本检查部位
    Call LoadSpecimenPart
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgData_OnColFormartChange()
'保存列表配置
    zlDatabase.SetPara "标本核收列表配置", ufgData.GetColsString(ufgData), glngSys, G_LNG_PATHOLSYS_NUM
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Call AdjustFace
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Set zlReport = Nothing
    Set mrsSpecimenPartData = Nothing
End Sub


Private Sub lstPartment_DblClick()
On Error GoTo ErrHandle
    Dim strPartSuperadd As String
    Dim strSpeciPartName As String
    Dim strSpeciName As String
    Dim blnFind As Boolean
    Dim objCheck As CheckState
    Dim lngImgIndex As Long
    Dim strSpecimenType As String

    
    If Not CheckAllowUpdateSpecimen(Val(ufgData.KeyValue(ufgData.SelectionRow))) Then
        Call MsgBoxD(Me, "该检查已进行取材，不能进行编辑。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    '判断当前选中是不是列头，是 则自动选择下一行，不是 就跳过
    If ufgData.SelectionRow = 0 Then
        Call ufgData.EditNextCell(1)
    End If
    
    strSpeciPartName = ufgData.Text(ufgData.SelectionRow, gSpecimen_采集部位)
    strSpeciName = ufgData.Text(ufgData.SelectionRow, gSpecimen_标本名称)
    
    mrsSpecimenPartData.Filter = "标本名称='" & lstPartment.Text & "'"
    
     '采集部位判断
    If strSpeciPartName <> "" And strSpeciPartName <> Nvl(mrsSpecimenPartData!标本部位) Then
        Call MsgBoxD(Me, "采集部位不一致，不能进行修改。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    
    ufgData.Text(ufgData.SelectionRow, gSpecimen_采集部位) = Nvl(mrsSpecimenPartData!标本部位)
    
    Call ufgData.GetFieldDisplayText(gSpecimen_标本类型, Val(Nvl(mrsSpecimenPartData!标本类型)), blnFind, objCheck, strSpecimenType, lngImgIndex)
    ufgData.Text(ufgData.SelectionRow, gSpecimen_标本类型) = Val(Nvl(mrsSpecimenPartData!标本类型)) & "-" & strSpecimenType
    
    '如果存在相同的名称，则结束
    If strSpeciName Like "*" & lstPartment.Text & "*" Then
        Call ufgData_OnAfterEdit(ufgData.DataGrid.Row, ufgData.DataGrid.Col)
        
        Call ufgData.DataGrid.Select(ufgData.SelectionRow, ufgData.GetColIndex(gSpecimen_标本名称))
        Call ufgData.DataGrid.EditCell
        Exit Sub
    End If
    
    
    '标本名称判断,有数据追加 无数据新增
    If strSpeciName <> "" Then
        strPartSuperadd = strSpeciName & "," & lstPartment.Text
    Else
        strPartSuperadd = lstPartment.Text
    End If
    
    ufgData.Text(ufgData.SelectionRow, gSpecimen_标本名称) = strPartSuperadd
    
    Call ufgData_OnAfterEdit(ufgData.DataGrid.Row, ufgData.DataGrid.Col)
    
    Call ufgData.DataGrid.Select(ufgData.SelectionRow, ufgData.GetColIndex(gSpecimen_标本名称))
    Call ufgData.DataGrid.EditCell
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgData_OnAfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim iCol As Long
    Dim i As Long
    Dim iCount As Long
    Dim strNewSpecimenName As String
    Dim lngCode As String
    Dim blnFind As Boolean
    Dim objCheck As CheckState
    Dim lngImgIndex As Long
    Dim strSpecimenType As String
    
    If ufgData.IsNullRow(Row) Then
        ufgData.RowState(Row) = TDataRowState.Normal
        Call ufgData.SetRowColor(Row, ufgData.BackColor)
        
        Exit Sub
    End If
    
    If Col = ufgData.GetColIndex(gSpecimen_标本名称) Then
    '检查标本名称是否重复
        lngCode = ufgData.Text(Row, gSpecimen_标本名称)
        If lngCode <> "" Then
            mrsSpecimenPartData.Filter = "简码='" & lngCode & "'"
            If mrsSpecimenPartData.RecordCount > 0 Then
                ufgData.Text(Row, gSpecimen_标本名称) = mrsSpecimenPartData!标本名称
                ufgData.Text(Row, gSpecimen_采集部位) = mrsSpecimenPartData!标本部位
                
                Call ufgData.GetFieldDisplayText(gSpecimen_标本类型, Val(Nvl(mrsSpecimenPartData!标本类型)), blnFind, objCheck, strSpecimenType, lngImgIndex)
                ufgData.Text(Row, gSpecimen_标本类型) = Val(Nvl(mrsSpecimenPartData!标本类型)) & "-" & strSpecimenType
            End If
        End If
    
        strNewSpecimenName = ufgData.CheckEquateValue(Row, Col)
        If strNewSpecimenName <> "" Then
            Call MsgBoxD(Me, "标本名称 [" & ufgData.Text(Row, gSpecimen_标本名称) & "]已经存在。", vbOKOnly, Me.Caption)
            
            ufgData.Text(Row, gSpecimen_标本名称) = strNewSpecimenName
        End If
    End If
    
    '如果未录入标本名称，则显示淡红色
    iCol = ufgData.GetColIndex(gSpecimen_标本名称)
    
    ufgData.CellColor(Row, iCol) = IIf(ufgData.Text(Row, gSpecimen_标本名称) = "", ufgData.ErrCellColor, ufgData.BackColor)
       
    
    '如果未录入标本类型，则显示淡红色
    iCol = ufgData.GetColIndex(gSpecimen_标本类型)
    
    ufgData.CellColor(Row, iCol) = IIf(ufgData.Text(Row, gSpecimen_标本类型) = "", ufgData.ErrCellColor, ufgData.BackColor)
    
    
    '如果未录入标本数量，则显示淡红色
    iCol = ufgData.GetColIndex(gSpecimen_数量)
    
    ufgData.CellColor(Row, iCol) = IIf(Val(ufgData.Text(Row, gSpecimen_数量)) <= 0, ufgData.ErrCellColor, ufgData.BackColor)
    
    
    
    '如果未录入材料，则显示淡红色
    iCol = ufgData.GetColIndex(gSpecimen_材料类别)
    
    ufgData.CellColor(Row, iCol) = IIf(ufgData.Text(Row, gSpecimen_材料类别) = "", ufgData.ErrCellColor, ufgData.BackColor)
    
    
    '当标本数量改变时，刷新标本数量的显示
    If Col = ufgData.GetColIndex(gSpecimen_数量) Then
        Call RefreshSpecimenCount
    End If
End Sub

'Private Function CheckIsMaterials(ByVal lngSpecimenID As Long) As Boolean
''检查是否进行取材处理
'
'    Dim strSql As String
'    Dim rsData As ADODB.Recordset
'
'    CheckIsMaterials = False
'
'    If lngSpecimenID <= 0 Then Exit Function
'
'    strSql = "select 材块ID from 病理取材信息 where  标本ID=[1]"
'    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngSpecimenID)
'
'    If rsData.RecordCount > 0 Then CheckAllowUpdateSpecimen = False
'
'End Function


Private Function CheckAllowUpdateSpecimen(ByVal lngSpecimenID As Long) As Boolean
'检查是否允许更新
'未制片的材块均可进行更新,通过检查病理制片信息表，可判断材块是否已制片(如果当前状态不为0，则已制片)

    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    CheckAllowUpdateSpecimen = True
    
    If lngSpecimenID <= 0 Then Exit Function
    
    strSql = "select 材块ID from 病理取材信息 where  标本ID=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngSpecimenID)
    
    If rsData.RecordCount > 0 Then CheckAllowUpdateSpecimen = False
End Function




Private Sub ufgData_OnColsNameReSet()
On Error GoTo ErrHandle

    Call LoadSpecimenData
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgData_OnSelChange()
    If mblLordingOrRefreshing Then Exit Sub
    If ufgData.SelectionRow = 0 Then Exit Sub
    If Not CheckAllowUpdateSpecimen(Val(ufgData.KeyValue(ufgData.SelectionRow))) Then
        Call ufgData.SetMenuState(False)
    Else
        Call ufgData.SetMenuState(True)
    End If
End Sub


Private Sub ufgData_OnStartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'接收日期读取服务器中的日期
    Dim dtServices As Date
    
    '当检查过程大于1时，说明已进行取材操作
    If Not CheckAllowUpdateSpecimen(Val(ufgData.KeyValue(ufgData.SelectionRow))) Then
        Cancel = True
        Call MsgBoxD(Me, "该检查已进行取材，不能进行编辑。", vbOKOnly, Me.Caption)
    End If
        
    
    'If Col = ufgData.vfgHelper.GetColumnIndex(gSpecimen_数量) And Row > 0 Then
        If Val(ufgData.Text(Row, gSpecimen_数量)) <= 0 Then ufgData.Text(Row, gSpecimen_数量) = "1"
    '    Exit Sub
    'End If
    
    
    'If Col = ufgData.vfgHelper.GetColumnIndex(gSpecimen_标本类型) And Row > 0 Then
        If Trim(ufgData.Text(Row, gSpecimen_标本类型)) = "" Then ufgData.Text(Row, gSpecimen_标本类型) = "0-手术标本"
    '    Exit Sub
    'End If
    
    'If Col = ufgData.vfgHelper.GetColumnIndex(gSpecimen_材料类别) And Row > 0 Then
        If Trim(ufgData.Text(Row, gSpecimen_材料类别)) = "" Then ufgData.Text(Row, gSpecimen_材料类别) = "0-标本"
    '    Exit Sub
    'End If
End Sub

Private Sub InitCommandBars()
On Error GoTo errH
    '功能创建工具条
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    Dim cbrPopControl As CommandBarControl
    Dim intTMP As Integer
    Dim cbrEdit As CommandBarEdit
                                
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbrMain.VisualTheme = xtpThemeOffice2003
    Set Me.cbrMain.Icons = zlCommFun.GetPubIcons
    
    With Me.cbrMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
    End With
    Me.cbrMain.EnableCustomization False
    Me.cbrMain.ActiveMenuBar.Visible = False
    
    '采集工具栏定义
    Set cbrToolBar = Me.cbrMain.Add("工具栏", xtpBarTop)
    cbrToolBar.EnableDocking xtpFlagStretched
    cbrToolBar.ShowTextBelowIcons = True
    cbrToolBar.Closeable = False
    With cbrToolBar.Controls

        Set cbrControl = .Add(xtpControlSplitButtonPopup, conMenu_PatholSpecimen_LAB, "标签打印"): cbrControl.IconId = 5001: cbrControl.ToolTipText = "标签打印"
            With cbrControl
                Set cbrPopControl = CreateModuleMenu(.CommandBar.Controls, xtpControlButton, conMenu_PatholSpecimen_PreviewLab, "预览", "", 0, False)
                Set cbrPopControl = CreateModuleMenu(.CommandBar.Controls, xtpControlButton, conMenu_PatholSpecimen_PrintLab, "打印", "", 0, False)
            End With
        Set cbrControl = .Add(xtpControlSplitButtonPopup, conMenu_PatholSpecimen_ACP, "凭单打印"): cbrControl.IconId = 5002: cbrControl.ToolTipText = "凭单打印"
            With cbrControl
                Set cbrPopControl = CreateModuleMenu(.CommandBar.Controls, xtpControlButton, conMenu_PatholSpecimen_PreviewAccept, "预览", "", 0, False)
                Set cbrPopControl = CreateModuleMenu(.CommandBar.Controls, xtpControlButton, conMenu_PatholSpecimen_PrintAccept, "打印", "", 0, False)
            End With

        Set cbrControl = .Add(xtpControlButton, conMenu_PatholSpecimen_Get, "提取标本"): cbrControl.IconId = 5003: cbrControl.ToolTipText = "提取标本"
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_PatholSpecimen_Del, "删除标本"): cbrControl.IconId = 5004: cbrControl.ToolTipText = "删除标本"
        
        
        Set cbrControl = .Add(xtpControlButton, conMenu_PatholSpecimen_Save, "保存标本"): cbrControl.IconId = 5005: cbrControl.ToolTipText = "保存标本"
        cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_PatholSpecimen_Accept, "接收标本"): cbrControl.IconId = 5006: cbrControl.ToolTipText = "接收标本"
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_PatholSpecimen_Reject, "拒收标本"): cbrControl.IconId = 5007: cbrControl.ToolTipText = "拒收标本"
        
        Set cbrControl = .Add(xtpControlButton, conMenu_PatholSpecimen_Cancel, "标本回退"): cbrControl.IconId = 5019: cbrControl.ToolTipText = "标本回退"
        cbrControl.BeginGroup = True
        
        
        

    End With
    Exit Sub
errH:
End Sub
