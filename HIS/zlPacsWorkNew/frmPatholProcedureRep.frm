VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "*\A..\zl9PacsControl\zl9PacsControl.vbp"
Begin VB.Form frmPatholProcedureRep 
   Caption         =   "冰冻特检报告"
   ClientHeight    =   8235
   ClientLeft      =   75
   ClientTop       =   405
   ClientWidth     =   11385
   Icon            =   "frmPatholProcedureRep.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   11385
   StartUpPosition =   1  '所有者中心
   Begin MSComctlLib.ImageList imgTbrS 
      Left            =   9045
      Top             =   540
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholProcedureRep.frx":000C
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholProcedureRep.frx":0C7E
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholProcedureRep.frx":18F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholProcedureRep.frx":2562
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholProcedureRep.frx":31D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholProcedureRep.frx":3E46
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholProcedureRep.frx":4AB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholProcedureRep.frx":572A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholProcedureRep.frx":639C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picBack 
      BorderStyle     =   0  'None
      Height          =   8040
      Left            =   30
      ScaleHeight     =   8040
      ScaleWidth      =   11325
      TabIndex        =   0
      Top             =   780
      Width           =   11325
      Begin zl9PacsControl.ucSplitter ucSplitter1 
         Height          =   8040
         Left            =   3255
         TabIndex        =   20
         Top             =   0
         Width           =   85
         _ExtentX        =   159
         _ExtentY        =   14182
         BackColor       =   -2147483633
         SplitWidth      =   85
         SplitLevel      =   3
         Control1Name    =   "picWordModule"
         Control2Name    =   "picReportEdit"
      End
      Begin VB.PictureBox picWordModule 
         BorderStyle     =   0  'None
         Height          =   8040
         Left            =   0
         ScaleHeight     =   8040
         ScaleWidth      =   3255
         TabIndex        =   16
         Top             =   0
         Width           =   3255
         Begin VB.Frame framWord 
            Height          =   7215
            Left            =   0
            TabIndex        =   17
            Top             =   0
            Width           =   3255
            Begin zl9PACSWork.WordInputModule wimWord 
               Height          =   4335
               Left            =   120
               TabIndex        =   18
               Top             =   240
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   7646
               CurDepartId     =   0
            End
            Begin zl9PACSWork.ucFlexGrid ufgData 
               Height          =   2415
               Left            =   120
               TabIndex        =   19
               Top             =   4680
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   4260
               DefaultCols     =   ""
               GridRows        =   21
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
      End
      Begin VB.PictureBox picReportEdit 
         BorderStyle     =   0  'None
         Height          =   8040
         Left            =   3340
         ScaleHeight     =   8040
         ScaleWidth      =   7980
         TabIndex        =   1
         Top             =   0
         Width           =   7985
         Begin VB.Frame framReport 
            Height          =   7455
            Left            =   45
            TabIndex        =   2
            Top             =   0
            Width           =   7695
            Begin VB.ComboBox cbxReportType 
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
               ItemData        =   "frmPatholProcedureRep.frx":700E
               Left            =   1020
               List            =   "frmPatholProcedureRep.frx":7010
               Style           =   2  'Dropdown List
               TabIndex        =   8
               Top             =   240
               Width           =   1545
            End
            Begin VB.ComboBox cbxSpecimenName 
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
               Left            =   6000
               Style           =   2  'Dropdown List
               TabIndex        =   7
               Top             =   240
               Width           =   1545
            End
            Begin VB.ComboBox cbxReportSub 
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
               Left            =   3600
               Style           =   2  'Dropdown List
               TabIndex        =   4
               Top             =   240
               Width           =   2025
            End
            Begin RichTextLib.RichTextBox txtAdvice 
               Height          =   1815
               Left            =   120
               TabIndex        =   3
               Top             =   3360
               Width           =   7335
               _ExtentX        =   12938
               _ExtentY        =   3201
               _Version        =   393217
               BorderStyle     =   0
               ScrollBars      =   2
               AutoVerbMenu    =   -1  'True
               TextRTF         =   $"frmPatholProcedureRep.frx":7012
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin RichTextLib.RichTextBox txtResult 
               Height          =   2055
               Left            =   120
               TabIndex        =   5
               Top             =   960
               Width           =   7335
               _ExtentX        =   12938
               _ExtentY        =   3625
               _Version        =   393217
               BorderStyle     =   0
               Enabled         =   -1  'True
               ScrollBars      =   2
               AutoVerbMenu    =   -1  'True
               TextRTF         =   $"frmPatholProcedureRep.frx":70AF
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin zl9PACSWork.ReportImage rpImage 
               Height          =   1935
               Left            =   120
               TabIndex        =   6
               Top             =   5280
               Width           =   7335
               _ExtentX        =   12938
               _ExtentY        =   3413
               ShowPhotoCount  =   3
               BackColor       =   4210752
            End
            Begin MSComCtl2.DTPicker dtpReportTime 
               Height          =   300
               Left            =   5640
               TabIndex        =   9
               Top             =   3050
               Width           =   1785
               _ExtentX        =   3149
               _ExtentY        =   529
               _Version        =   393216
               CustomFormat    =   "yyyy-MM-dd HH:mm"
               Format          =   110493699
               CurrentDate     =   40646.4399652778
            End
            Begin VB.Label labReportType 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "报告类型："
               ForeColor       =   &H80000008&
               Height          =   180
               Left            =   120
               TabIndex        =   15
               Top             =   300
               Width           =   900
            End
            Begin VB.Label labSpecimenName 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "标本名："
               ForeColor       =   &H80000008&
               Height          =   180
               Left            =   5280
               TabIndex        =   14
               Top             =   300
               Width           =   720
            End
            Begin VB.Line Line1 
               X1              =   110
               X2              =   7440
               Y1              =   650
               Y2              =   650
            End
            Begin VB.Label labResult 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "检查结果："
               ForeColor       =   &H80000008&
               Height          =   180
               Left            =   120
               TabIndex        =   13
               Top             =   720
               Width           =   900
            End
            Begin VB.Label labAdvice 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "病理诊断："
               ForeColor       =   &H80000008&
               Height          =   180
               Left            =   120
               TabIndex        =   12
               Top             =   3120
               Width           =   900
            End
            Begin VB.Label labReportSub 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "报告子项："
               ForeColor       =   &H80000008&
               Height          =   180
               Left            =   2640
               TabIndex        =   11
               Top             =   300
               Width           =   900
            End
            Begin VB.Label labReportTime 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "报告时间："
               ForeColor       =   &H80000008&
               Height          =   180
               Left            =   4800
               TabIndex        =   10
               Top             =   3090
               Width           =   900
            End
         End
      End
   End
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   0
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmPatholProcedureRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IWorkMenu

Private Const M_REPORTSTATE_NORMAL As Long = 0  '未打印
Private Const M_REPORTSTATE_VIEW As Long = 1    '已阅
Private Const M_REPORTSTATE_CANCEL As Long = 2  '已撤回
Private Const M_REPORTSTATE_PRINT As Long = 3   '已打印

Private Const M_STR_HINT_NoSelectData As String = "无效的检查数据，请选择需要执行的检查记录。"
Private Const M_STR_MODULE_MENU_TAG As String = "过程"


Dim WithEvents zlReport As zl9Report.clsReport
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
Private mrecStudyInf As TStudyStateInf

Private mlngTmpAdviceId As Long
Private mlngTmpSendNo As Long


Private mSelMiniImg As DicomImage

Private mintShowPhotoNumber As Integer
Private mintCurImgIndex As Integer
Private strCurTempReportPath As String
Private mblnEditState As Boolean

Private mCurEditText As RichTextBox


Private mblnIsAllowSpeExam As Boolean
Private mblnIsAllowWriteReport As Boolean

Private mbytFontSize As Byte '字号    9--小字体    12--大字体


Private mObjActiveMenuBar As CommandBar

Private mblnRefreshState As Boolean
Private mblnIsLoading As Boolean '用于判断是否正在加载数据
Private mlngRow As Long '保存列表数据row



'获取菜单接口对象
Property Get zlMenu() As IWorkMenu
    Set zlMenu = Me
End Property


Public Sub NotificationRefresh()
'通知刷新
    mblnRefreshState = False
End Sub

Private Sub cbxReportSub_LostFocus()
    If mblnIsLoading = True Then Exit Sub

    mblnEditState = True
End Sub

Private Sub cbxReportType_LostFocus()
    If mblnIsLoading = True Then Exit Sub
    
    mblnEditState = True
End Sub


Private Sub Form_Resize()
On Error GoTo ErrHandle
    picBack.Left = 0
    If mbytFontSize = C_INT_FONTSISE_SMALL Then
        picBack.Top = 800
    ElseIf mbytFontSize = C_INT_FONTSISE_MEDIUM Then
        picBack.Top = 850
    Else
        picBack.Top = 900
    End If
    picBack.Width = Me.ScaleWidth
    picBack.Height = Me.ScaleHeight - 1000

    Call ucSplitter1.RePaint(False)
Exit Sub
ErrHandle:
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


    If Not HasMenu(objMenuBar, conMenu_PatholProRep) Then
        Set cbrMenuBar = mObjActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_PatholProRep, "过程报告(&O)", GetPatholMenuIndex(objMenuBar) + 1, False)
        cbrMenuBar.ID = conMenu_PatholProRep
        cbrMenuBar.Category = M_STR_MODULE_MENU_TAG
                
        With cbrMenuBar.CommandBar
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlPopup, conMenu_PatholProRep_Report, "报告打印(&V)", "", 1, False)
            With cbrControl
                Set cbrPopControl = CreateModuleMenu(.CommandBar.Controls, xtpControlButton, conMenu_PatholProRep_Preview, "预览(&V)", "", 1, False)
                Set cbrPopControl = CreateModuleMenu(.CommandBar.Controls, xtpControlButton, conMenu_PatholProRep_Print, "打印(&P)", "", 1, False)
            End With
            
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholProRep_Already, "报告查阅(&A)", "", 1, False)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholProRep_Back, "报告撤回(&C)", "", 1, False)
            
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholProRep_Clear, "清除内容(&R)", "", 1, True)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholProRep_Input, "特检项目录入(&I)", "", 1, False)
            
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholProRep_New, "新增报告(&N)", "", 1, True)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholProRep_Del, "删除报告(&D)", "", 1, False)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholProRep_Save, "保存报告(&S)", "", 1, True)
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
        Case conMenu_PatholProRep_Preview
            Call PrintCurProcedureRep(False)
            
        Case conMenu_PatholProRep_Print
            Call PrintCurProcedureRep(True)
            
        Case conMenu_PatholProRep_Already
            Call MakeSureShowMsg(True)
            Call UpdateCurProcedureRepState(M_REPORTSTATE_VIEW)
            
        Case conMenu_PatholProRep_Back
            Call UpdateCurProcedureRepState(M_REPORTSTATE_CANCEL)
            
        Case conMenu_PatholProRep_Clear
            Call ClearReportContext
            
        Case conMenu_PatholProRep_Input
            Call GetSpeExamResult
            
        Case conMenu_PatholProRep_New
            Call NewProcedureRep
            
        Case conMenu_PatholProRep_Del
            Call DelCurProcedureRep
            
        Case conMenu_PatholProRep_Save
            Call SaveCurProcedureRep
    End Select
End Sub


Public Sub IWorkMenu_zlUpdateMenu(ByVal control As XtremeCommandBars.ICommandBarControl)
'更新菜单
    Dim blnIsPopedom As Boolean
    Dim blHavePatholNumber As Boolean
    Dim blGetCurRepAllowAuditing As Boolean
    
    If Not Me.Visible Then
        control.Enabled = False
        Exit Sub
    End If
    
    Select Case Val(cbxReportType.Text)
        Case 0:
            blnIsPopedom = CheckPopedom(mstrPrivs, "冰冻报告")
        Case 1
            blnIsPopedom = CheckPopedom(mstrPrivs, "免疫报告")
        Case 2, 3
            blnIsPopedom = CheckPopedom(mstrPrivs, "分子报告")
        Case 4
            blnIsPopedom = CheckPopedom(mstrPrivs, "特染报告")
    End Select
    
    blGetCurRepAllowAuditing = GetCurRepAllowAuditing
    blHavePatholNumber = (Len(mrecStudyInf.strPatholNumber) > 0)
    
    Select Case control.ID
        Case conMenu_PatholProRep_Report
            control.Enabled = (mblnIsAllowSpeExam Or mblnIsAllowWriteReport) And blHavePatholNumber And ufgData.ShowingRowCount > 0
            
        Case conMenu_PatholProRep_Preview, conMenu_PatholProRep_Print
            control.Enabled = (mblnIsAllowSpeExam Or mblnIsAllowWriteReport) And blHavePatholNumber
            
        Case conMenu_PatholProRep_Already
            control.Enabled = mblnIsAllowWriteReport And Not mblnReadOnly And Not blGetCurRepAllowAuditing And blHavePatholNumber And ufgData.ShowingRowCount > 0
            
        Case conMenu_PatholProRep_Back
            control.Enabled = mblnIsAllowWriteReport And Not mblnReadOnly And blGetCurRepAllowAuditing And blHavePatholNumber And ufgData.ShowingRowCount > 0
            
        Case conMenu_PatholProRep_Input
            control.Enabled = Not (mblnReadOnly Or blGetCurRepAllowAuditing) And blnIsPopedom And Val(cbxReportType.Text) > 0 And blHavePatholNumber
            
        Case conMenu_PatholProRep_Clear
            control.Enabled = Not (mblnReadOnly Or blGetCurRepAllowAuditing) And blnIsPopedom And blHavePatholNumber
            
        Case conMenu_PatholProRep_New
            control.Enabled = Not mblnReadOnly And blnIsPopedom And blHavePatholNumber
            
        Case conMenu_PatholProRep_Del
            control.Enabled = Not (mblnReadOnly Or blGetCurRepAllowAuditing) And blnIsPopedom And blHavePatholNumber
            
        Case conMenu_PatholProRep_Save
            control.Enabled = Not (mblnReadOnly Or blGetCurRepAllowAuditing) And blnIsPopedom And blHavePatholNumber And mblnEditState
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
    If blnMoved Or lngStudyState = 6 Or lngStudyState = 5 Or lngStudyState = 0 Or lngStudyState = 1 Or lngStudyState = -2 Then
        mblnReadOnly = True
    End If

End Sub

Public Sub zlRefreshFace(Optional blnForceRefresh As Boolean = False)
'刷新界面数据
    Call MakeSureShowMsg
    
    If (mlngTmpAdviceId = mlngAdviceID And mlngTmpSendNo = mlngSendNo And mblnRefreshState) And Not blnForceRefresh Then Exit Sub
    
    mlngTmpAdviceId = mlngAdviceID
    mlngTmpSendNo = mlngSendNo
    mblnRefreshState = True
    
    Call ClearReportContext
    
    If mlngAdviceID <= 0 Then
        Call ConfigProcedureReportFace(False, "医嘱ID无效请检查。")
        Exit Sub
    End If
    
    Call LoadReportModule
        
    mblnIsAllowSpeExam = CheckPopedom(mstrPrivs, "冰冻报告") Or CheckPopedom(mstrPrivs, "免疫报告") Or CheckPopedom(mstrPrivs, "分子报告") Or CheckPopedom(mstrPrivs, "特染报告")
    mblnIsAllowWriteReport = CheckPopedom(mstrPrivs, "冰冻特检报告查阅")
    
    
    Call GetPatholStudyState(mlngAdviceID, mrecStudyInf)

    
    If mrecStudyInf.strPatholNumber = "" Then
        Call ConfigProcedureReportFace(False, "该检查尚未生成有效的病理号，请确认该检查是否已被核收。")
        Call rpImage.ReInit
        
        Exit Sub
    Else
        Call ConfigReportType
        Call LoadReportSub(Val(cbxReportType.Text))
        Call ConfigProcedureReportFace(True)
        
 
        '载入报告图像
        Call rpImage.LoadReportImages(mlngAdviceID, mblnMoved, Me)
        '配置标本录入
        Call ConfigSpecimenName(mlngAdviceID)
        '读取过程报告记录
        Call LoadProcedureRepData(mblnReadOnly)
    End If

    Call EnableReportWithSpeExamType(Val(cbxReportType.Text), mblnReadOnly, GetCurRepAllowAuditing)
'    If Not (owner Is Nothing) Then
'        Call Me.Show(1, owner)
'    End If
End Sub


Public Sub zlRefresh(lngAdviceID As Long, ByVal blnReadOnly As Boolean, _
    strPrivs As String, ByVal blnMoved As Boolean, _
    ByVal lngCurDepartmentId As Long, Optional owner As Form = Nothing)
'刷新数据
    Call ClearReportContext
    
        
    If lngAdviceID <= 0 Then
        Call ConfigProcedureReportFace(False, "医嘱ID无效请检查。")
        Exit Sub
    End If

'    If mlngCurAdviceId = lngAdviceID Then Exit Sub
    
    mlngAdviceID = lngAdviceID
    mstrPrivs = strPrivs
    mblnMoved = blnMoved
    mlngCurDeptId = lngCurDepartmentId
'    mblnReadOnly = blnReadOnly
    
    Call LoadReportModule
        
    mblnIsAllowSpeExam = CheckPopedom(mstrPrivs, "冰冻报告") Or CheckPopedom(mstrPrivs, "免疫报告") Or CheckPopedom(mstrPrivs, "分子报告") Or CheckPopedom(mstrPrivs, "特染报告")
    mblnIsAllowWriteReport = CheckPopedom(mstrPrivs, "冰冻特检报告查阅")
    
    
    Call GetPatholStudyState(lngAdviceID, mrecStudyInf)

    
    If mrecStudyInf.strPatholNumber = "" Then
        Call ConfigProcedureReportFace(False, "该检查尚未生成有效的病理号，请确认该检查是否已被核收。")
        Call rpImage.ReInit
        
        Exit Sub
    Else
        Call ConfigReportType
        Call LoadReportSub(Val(cbxReportType.Text))
        Call ConfigProcedureReportFace(True)
        
 
        '载入报告图像
        Call rpImage.LoadReportImages(mlngAdviceID, mblnMoved, Me)
        '配置标本录入
        Call ConfigSpecimenName(mlngAdviceID)
        '读取过程报告记录
        Call LoadProcedureRepData(blnReadOnly)
    End If

    Call EnableReportWithSpeExamType(Val(cbxReportType.Text), blnReadOnly, GetCurRepAllowAuditing)

    
    If Not (owner Is Nothing) Then
        Call Me.Show(1, owner)
    End If
End Sub



Private Function GetCurRepAllowAuditing() As Boolean
'判断当前过程报告是否允许查阅
    GetCurRepAllowAuditing = False
    
    If ufgData.ShowingDataRowCount <= 0 Then Exit Function
    If ufgData.IsEmptyKey(ufgData.SelectionRow) Then Exit Function
    
    GetCurRepAllowAuditing = ("已打印,已查阅" Like "*" & ufgData.Text(ufgData.SelectionRow, gstrProcedureRep_当前状态) & "*")
End Function



Private Sub ConfigProcedureReportFace(ByVal blnIsValid As Boolean, Optional ByVal strHintInf As String = "")
'配置特检界面
    Call EnableReportWithSpeExamType(Val(cbxReportType.Text), Not blnIsValid, Not blnIsValid)
    
    If blnIsValid Then
        Call ufgData.CloseHintInf
    Else
    
        Call ufgData.ShowHintInf(strHintInf)
    End If
    
    cbxReportSub.Enabled = blnIsValid
    dtpReportTime.Enabled = blnIsValid
End Sub



Private Sub InitProcedureRepList()
'初始化过程报告列表
    Dim strTemp As String
    
    ufgData.IsKeepRows = False


     '判断数据库参数表是否有数据 有则读取数据库参数  没有则加载默认
    strTemp = zlDatabase.GetPara("过程报告列表配置", glngSys, G_LNG_PATHOLSYS_NUM, "")
    ufgData.DefaultColNames = gstrProcedureRepCols
     
    If strTemp = "" Then
        ufgData.ColNames = gstrProcedureRepCols
    Else
        ufgData.ColNames = strTemp
    End If
    
    '禁止右键弹出列表配置窗口
    ufgData.IsEjectConfig = False
    '设置行数
    ufgData.GridRows = glngStandardRowCount
    '设置行高
    ufgData.RowHeightMin = glngStandardRowHeight
    ufgData.ColConvertFormat = gstrProcedureRepConvertFormat
    Call ufgData.SetMenuState(False)
End Sub

Private Sub ufgData_OnColFormartChange()
'关闭窗口时保存列表配置
    zlDatabase.SetPara "过程报告列表配置", ufgData.GetColsString(ufgData), glngSys, G_LNG_PATHOLSYS_NUM
End Sub


Private Sub LoadProcedureRepData(ByVal blnReadOnly As Boolean)
'读取过程报告数据
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim blnIsAllowWriteReport As Boolean
    
    blnIsAllowWriteReport = CheckPopedom(mstrPrivs, "冰冻特检报告查阅")
    
    strSql = "select ID,标本名称,报告类型,报告子项,检查结果,检查意见,报告图像,报告医师,报告日期,当前状态,备注 from 病理过程报告 where 病理医嘱ID=[1] and (报告类型=-1"
    
    If CheckPopedom(mstrPrivs, "冰冻报告") Or blnIsAllowWriteReport Then strSql = strSql & " or 报告类型=0"
    If CheckPopedom(mstrPrivs, "免疫报告") Or blnIsAllowWriteReport Then strSql = strSql & " or 报告类型=1"
    If CheckPopedom(mstrPrivs, "分子报告") Or blnIsAllowWriteReport Then strSql = strSql & " or 报告类型=2"
    If CheckPopedom(mstrPrivs, "特染报告") Or blnIsAllowWriteReport Then strSql = strSql & " or 报告类型=3"
    
    strSql = strSql & ")"
    
    Set ufgData.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mrecStudyInf.lngPatholAdviceId)

    Call ufgData.RefreshData
    
    If ufgData.ShowingDataRowCount >= 1 Then
        Call LoadReportContext(1)
    
        Call EnableReportWithSpeExamType(Val(cbxReportType.Text), blnReadOnly, GetCurRepAllowAuditing)
    End If
End Sub



Private Sub EnableReportWithSpeExamType(ByVal lngSpeExamType As Long, ByVal blnStudyFinal As Boolean, _
    ByVal blnRepAuditing As Boolean, Optional ByVal blnShowHint As Boolean = True)
'配置报告的编辑状态
    Dim blnIsPopedom As Boolean
    
    Select Case lngSpeExamType
        Case 0:
            blnIsPopedom = CheckPopedom(mstrPrivs, "冰冻报告")
        Case 1
            blnIsPopedom = CheckPopedom(mstrPrivs, "免疫报告")
        Case 2, 3
            blnIsPopedom = CheckPopedom(mstrPrivs, "分子报告")
        Case 4
            blnIsPopedom = CheckPopedom(mstrPrivs, "特染报告")
    End Select
        
    cbxReportType.Enabled = mblnIsAllowSpeExam And Not blnStudyFinal
    cbxSpecimenName.Enabled = mblnIsAllowSpeExam And Not blnStudyFinal
    
    txtResult.Locked = Not blnIsPopedom Or (blnStudyFinal Or blnRepAuditing)
    txtAdvice.Locked = Not blnIsPopedom Or (blnStudyFinal Or blnRepAuditing)
    
    txtResult.BackColor = IIf(Not (blnStudyFinal Or blnRepAuditing) And blnIsPopedom, vbWhite, Me.BackColor)
    txtAdvice.BackColor = IIf(Not (blnStudyFinal Or blnRepAuditing) And blnIsPopedom, vbWhite, Me.BackColor)
    
    txtResult.Enabled = Not (blnStudyFinal Or blnRepAuditing) And blnIsPopedom
    txtAdvice.Enabled = txtResult.Enabled
    
    rpImage.Enable = Not (blnStudyFinal Or blnRepAuditing) And blnIsPopedom
    
    If ufgData.ShowingRowCount < 1 Then '若无项目存在
        txtResult.Enabled = False
        txtAdvice.Enabled = txtResult.Enabled
        Exit Sub
    End If
        
End Sub


Private Sub LoadReportType()
'载入报告类型
    Dim lngIndex As Long
    
    Call cbxReportType.Clear
    
    Call cbxReportType.AddItem("0-冰冻报告")
    Call cbxReportType.AddItem("1-免疫报告")
    Call cbxReportType.AddItem("2-分子报告")
    Call cbxReportType.AddItem("3-特染报告")
    
    cbxReportType.ListIndex = 1
End Sub


Private Sub LoadReportSub(ByVal lngReportType As Long)
'载入报告子项
    mblnIsLoading = True
    cbxReportSub.Clear
    
'    Call cbxReportSub.AddItem("")
    
    If lngReportType = 1 Then
        Call cbxReportSub.AddItem("1-免疫(鉴别)")
        Call cbxReportSub.AddItem("2-免疫(多药耐药)")
        
        cbxReportSub.ListIndex = 1
    ElseIf lngReportType = 2 Then
        Call cbxReportSub.AddItem("1-分子(荧光)")  '对应 3
        Call cbxReportSub.AddItem("2-分子(普通)")  '对应 4
        
        cbxReportSub.ListIndex = 0
    End If
    
    mblnIsLoading = False
    
End Sub



Public Sub ReSetFormFontSize(ByVal bytFontSize As Byte)
'功能:重新设置工作站窗体的字体大小
On Error GoTo ErrHandle

    Dim objCtrl As control
    Dim CtlFont As StdFont
    Dim strFontType  As String
    
    mblnIsLoading = True
    mbytFontSize = bytFontSize
    
    cbxReportType.Left = cbxReportType.Left + IIf(cbxReportType.FontSize = bytFontSize, 0, IIf(bytFontSize = C_INT_FONTSISE_SMALL, -100, 100))
    cbxReportSub.Left = cbxReportSub.Left + IIf(cbxReportSub.FontSize = bytFontSize, 0, IIf(bytFontSize = C_INT_FONTSISE_SMALL, -100, 100))
    cbxSpecimenName.Left = cbxSpecimenName.Left + IIf(cbxSpecimenName.FontSize = bytFontSize, 0, IIf(bytFontSize = C_INT_FONTSISE_SMALL, -100, 100))
    
    
    Set CtlFont = New StdFont
    Me.FontSize = bytFontSize
    strFontType = IIf(IsUseClearType = True, "微软雅黑", "宋体")
    
    CtlFont.Name = strFontType
    CtlFont.Size = bytFontSize
    
    For Each objCtrl In Me.Controls
        Select Case UCase(TypeName(objCtrl))
        Case UCase("Frame")
            If objCtrl.Name = "FramCheck" Then
                objCtrl.Height = TextHeight("测") * 1.7
            End If
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
            objCtrl.Height = TextHeight("测") + 100
        Case UCase("listBox")
            objCtrl.Font = CtlFont
        Case UCase("CheckBox")
            objCtrl.FontSize = bytFontSize
            objCtrl.FontName = strFontType
            objCtrl.Width = TextWidth("测试" & objCtrl.Caption)
        Case UCase("DTPicker")
            objCtrl.Font.Size = bytFontSize
            objCtrl.Font.Name = strFontType
            objCtrl.Width = TextWidth("2012-01-01 23:59:59") * 1.25
            objCtrl.Height = TextHeight("测") * 1.7
        Case UCase("textBox")
            objCtrl.FontSize = bytFontSize
            objCtrl.FontName = strFontType
        Case UCase("richtextbox")
            objCtrl.Font.Size = bytFontSize
            objCtrl.SelFontSize = bytFontSize
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
        End Select
    Next
    
    ''''''''''''''''''''''''''一些单独的界面控件调整107522
    cbxReportType.Width = TextWidth("设计六个字加") * 1.25
    cbxReportSub.Width = TextWidth("设计八个字符加加") * 1.25
    cbxSpecimenName.Width = TextWidth("设计六个字加") * 1.25
    
    Call picReportEdit_Resize
    
    mblnIsLoading = False
    Exit Sub
ErrHandle:
End Sub






Private Sub ConfigReportType()
'配置当前检查的报告类型
    '判断检查类型，如果是冰冻检查，则默认设置为冰冻检查报告，
    If mrecStudyInf.lngStudyType = 1 Then
        If CheckPopedom(mstrPrivs, "冰冻报告") Then
            cbxReportType.ListIndex = 0
            Exit Sub
        End If
    End If
    
    
    If mrecStudyInf.lngMianYiStep > 0 Then
        If CheckPopedom(mstrPrivs, "免疫报告") Then
            cbxReportType.ListIndex = 1
            Exit Sub
        End If
    End If
 
    
    If mrecStudyInf.lngTeRanStep > 0 Then
        If CheckPopedom(mstrPrivs, "特染报告") Then
            cbxReportType.ListIndex = 3
            Exit Sub
        End If
    End If
    
    If mrecStudyInf.lngFenZiStep > 0 Then
        If CheckPopedom(mstrPrivs, "分子报告") Then
            cbxReportType.ListIndex = 2
            Exit Sub
        End If
    End If
    
    
    
    
    If CheckPopedom(mstrPrivs, "免疫报告") Then
        cbxReportType.ListIndex = 1
        Exit Sub
    End If
    
    
    If CheckPopedom(mstrPrivs, "分子报告") Then
        cbxReportType.ListIndex = 2
        Exit Sub
    End If
    
    
    If CheckPopedom(mstrPrivs, "特染报告") Then
        cbxReportType.ListIndex = 3
        Exit Sub
    End If
End Sub


Private Sub ConfigSpecimenName(ByVal lngAdviceID As String)
'读取标本名称
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    strSql = "select 标本名称 from 病理标本信息 where 医嘱ID=[1] and 送检ID > 0"
'    If mblnMoved Then strSql = GetMovedDataSql(strSql)
    
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngAdviceID)
    
    mblnIsLoading = True
    Call cbxSpecimenName.Clear
    
    If rsData.RecordCount < 0 Then Exit Sub

    Call cbxSpecimenName.AddItem("")
    While Not rsData.EOF
        Call cbxSpecimenName.AddItem(Nvl(rsData!标本名称))
                
        rsData.MoveNext
    Wend

    cbxSpecimenName.ListIndex = 0
    mblnIsLoading = False
End Sub


Private Sub ShowReportImageWindow()
'
    Dim frmImage As New frmPatholProcedureRep_Image
    On Error GoTo errFree
        Call frmImage.ShowImageWindow(mlngAdviceID, mblnMoved, Me)
errFree:
    Call Unload(frmImage)
    Set frmImage = Nothing
End Sub


Private Sub cmdAddRepImage_Click()
On Error GoTo ErrHandle
    Call ShowReportImageWindow
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub



Private Function GetSelectReportImgs() As String
'获取选择的报告图像
    Dim i As Long
    Dim j As Long
    Dim objLabs As DicomLabels
    Dim strUids As String
    
    strUids = ""
    For i = 1 To rpImage.dcmViewer.Images.Count
        Set objLabs = rpImage.dcmViewer.Images(i).Labels
        
        For j = 1 To objLabs.Count
            If objLabs(j).tag = rpImage.SelectTag Then
                If Not objLabs(j).Transparent Then
                    If strUids <> "" Then strUids = strUids & ";"
                    strUids = strUids & rpImage.dcmViewer.Images(i).InstanceUID
                End If
            End If
        Next j
    Next i
    
    GetSelectReportImgs = strUids
End Function



Private Function GetReportTypeValue(ByVal strCode As String) As String
'获取报告类型
    Dim blnFind As Boolean
    Dim chkState As CheckState
    Dim strValue As String
    Dim lngImgIndex As Long
    
    Call ufgData.GetFieldDisplayText(gstrProcedureRep_报告类型, strCode, blnFind, chkState, strValue, lngImgIndex)
    GetReportTypeValue = IIf(blnFind, strValue, strCode)
End Function

Private Function GetReportSubValue(ByVal strCode As String) As String
'获取报告子项
    Dim blnFind As Boolean
    Dim chkState As CheckState
    Dim strValue As String
    Dim lngImgIndex As Long
    
    Call ufgData.GetFieldDisplayText(gstrProcedureRep_报告子项, strCode, blnFind, chkState, strValue, lngImgIndex)
    GetReportSubValue = IIf(blnFind, strValue, strCode)
End Function



Private Function GetReportTypeCode(ByVal strValue As String) As String
'获取报告类型
    Dim blnFind As Boolean
    Dim strCode As String
    
    strCode = ufgData.GetFieldDataValue(gstrProcedureRep_报告类型, strValue, blnFind)
    GetReportTypeCode = IIf(blnFind, strCode, strValue)
End Function


Private Function GetReportSubCode(ByVal strValue As String) As String
'获取报告子项
    Dim blnFind As Boolean
    Dim strCode As String
    
    strCode = ufgData.GetFieldDataValue(gstrProcedureRep_报告子项, strValue, blnFind)
    GetReportSubCode = IIf(blnFind, strCode, strValue)
End Function


Private Sub NewProcedureRep()
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim strRepImages As String
    Dim lngNewRow As Long
    Dim dtServicesTime As Date
    Dim lngSpeexamDetails As Long
    
    
    lngSpeexamDetails = 0
    
    '获取当前特检细目
    Select Case Val(cbxReportType.Text)
        Case 1
            lngSpeexamDetails = Val(cbxReportSub.Text)
        Case 3
            lngSpeexamDetails = 0
        Case 2
            lngSpeexamDetails = IIf(Val(cbxReportSub.Text) > 0, Val(cbxReportSub.Text) + 2, 0)
    End Select
    
    
    strRepImages = GetSelectReportImgs()
    dtServicesTime = dtpReportTime.value  ' zlDatabase.Currentdate
    
    strSql = "select Zl_病理过程报告_新增([1],[2],[3],[4],[5],[6],[7],[8],[9],[10]) as 返回值 from dual"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, _
                                            mrecStudyInf.lngPatholAdviceId, _
                                            cbxSpecimenName.Text, _
                                            Val(cbxReportType.Text), _
                                            lngSpeexamDetails, _
                                            txtAdvice.Text, _
                                            txtResult.Text, _
                                            UserInfo.姓名, _
                                            CDate(dtServicesTime), _
                                            strRepImages, _
                                            "")
                                            
    If rsData.RecordCount <= 0 Then
        Call err.Raise(0, "NewProcedureRep", "未成功获取新增后的报告ID,处理失败。")
        Exit Sub
    End If
    
    '填充过程报告数据列表
    lngNewRow = ufgData.NewRow ' ufgData.GetNullRowIndex
    
    ufgData.Text(lngNewRow, gstrProcedureRep_ID) = rsData!返回值
    ufgData.Text(lngNewRow, gstrProcedureRep_报告图像) = strRepImages
    ufgData.Text(lngNewRow, gstrProcedureRep_标本名称) = cbxSpecimenName.Text
    ufgData.Text(lngNewRow, gstrProcedureRep_报告类型) = GetReportTypeValue(Val(cbxReportType.Text))
    ufgData.Text(lngNewRow, gstrProcedureRep_报告子项) = GetReportSubValue(IIf(Val(cbxReportType.Text) = 1, Val(cbxReportSub.Text), IIf(Val(cbxReportType.Text) = 2, Val(cbxReportSub.Text) + 2, 0)))
    ufgData.Text(lngNewRow, gstrProcedureRep_检查结果) = txtResult.Text
    ufgData.Text(lngNewRow, gstrProcedureRep_检查意见) = txtAdvice.Text
    ufgData.Text(lngNewRow, gstrProcedureRep_报告人) = UserInfo.姓名
    ufgData.Text(lngNewRow, gstrProcedureRep_报告日期) = dtServicesTime
    ufgData.Text(lngNewRow, gstrProcedureRep_当前状态) = "未打印"
'    ufgData.text(lngNewRow, gstrProcedureRep_备注)=txtMemo.Text

'    Call ufgData.LocateRow(lngNewRow)
    Call ufgData_OnSelChange

    'Call MsgBoxD(Me, Decode(Val(cbxReportType.Text), 0, "冰冻", 1, "免疫", 2, "分子", 3, "特染", "") & "报告已成功添加。", vbOKOnly, Me.Caption)
End Sub



Private Sub LoadReportContext(ByVal lngRow As Long)
'载入报告内容
    Dim i As Long
    Dim strRepImages As String
    Dim lngReportSub As Long
    
    mblnIsLoading = True
    txtResult.Text = ufgData.Text(lngRow, gstrProcedureRep_检查结果)
    txtAdvice.Text = ufgData.Text(lngRow, gstrProcedureRep_检查意见)
    
    dtpReportTime.value = ufgData.Text(lngRow, gstrProcedureRep_报告日期)
    
    '读取标本名称
    For i = 0 To cbxSpecimenName.ListCount - 1
        If cbxSpecimenName.list(i) = ufgData.Text(lngRow, gstrProcedureRep_标本名称) Then
            cbxSpecimenName.ListIndex = i
            Exit For
        End If
    Next i
    
    '读取报告类型
    cbxReportType.ListIndex = GetReportTypeCode(ufgData.Text(lngRow, gstrProcedureRep_报告类型))
    
    '读取报告子项
    lngReportSub = GetReportSubCode(ufgData.Text(lngRow, gstrProcedureRep_报告子项))
    cbxReportSub.ListIndex = IIf(lngReportSub > 2, lngReportSub - 2, lngReportSub) - 1
    
    '配置图像的选择状态
    strRepImages = ufgData.Text(lngRow, gstrProcedureRep_报告图像)
    
    For i = 1 To rpImage.dcmViewer.Images.Count
        If InStr(1, strRepImages, rpImage.dcmViewer.Images(i).InstanceUID) > 0 Then
            rpImage.ItemSelected(i) = True
        Else
            rpImage.ItemSelected(i) = False
        End If
    Next i
    
    mblnIsLoading = False
End Sub


Private Sub ClearReportContext()
'清除报告编辑器内容
    mblnIsLoading = True
    txtResult.Text = ""
    txtAdvice.Text = ""
    
    If cbxSpecimenName.ListCount > 0 Then cbxSpecimenName.ListIndex = 0
    
    mblnIsLoading = False
    Call rpImage.ClearSelected
    
End Sub



Private Sub cbxReportType_Click()
On Error GoTo ErrHandle
    Call LoadReportSub(Val(cbxReportType.Text))
    Call LoadReportModule(True)
    
    Call EnableReportWithSpeExamType(Val(cbxReportType.Text), mblnReadOnly, GetCurRepAllowAuditing)
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cbxSpecimenName_LostFocus()
    If mblnIsLoading Then Exit Sub

    mblnEditState = True
End Sub


Private Sub SaveCurProcedureRep()
'保存过程报告更新
    Dim strSql As String
    Dim strSelectRpImages As String
    Dim lngSpeexamDetails As Long
    Dim dtServicesTime As Date
    Dim lngRow As Long
    
    If Not ufgData.IsSelectionRow Then Exit Sub
    
    If ufgData.IsNullRow(ufgData.SelectionRow) Then
        Call MsgBoxD(Me, "当前报告记录不能进行保存，请尝试“新增报告”。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgData.IsEmptyKey(ufgData.SelectionRow) Then
        Call NewProcedureRep
        
        Exit Sub
    End If
    
    lngSpeexamDetails = 0
    
    '获取当前特检细目
    Select Case Val(cbxReportType.Text)
        Case 1
            lngSpeexamDetails = Val(cbxReportSub.Text)
        Case 3
            lngSpeexamDetails = 0
        Case 2
            lngSpeexamDetails = IIf(Val(cbxReportSub.Text) > 0, Val(cbxReportSub.Text) + 2, 0)
    End Select
    
    
    dtServicesTime = dtpReportTime.value
    strSelectRpImages = GetSelectReportImgs()
    
    lngRow = mlngRow
    
    strSql = "Zl_病理过程报告_更新(" & ufgData.KeyValue(lngRow) & ",'" & _
                                        cbxSpecimenName.Text & "'," & _
                                        Val(cbxReportType.Text) & "," & _
                                        lngSpeexamDetails & ",'" & _
                                        txtAdvice.Text & "','" & _
                                        txtResult.Text & "'," & _
                                        zlStr.To_Date(dtServicesTime) & ",'" & _
                                        strSelectRpImages & "',Null)"
                                        
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    mblnEditState = False
    
    '更新数据列表
    ufgData.Text(lngRow, gstrProcedureRep_报告图像) = strSelectRpImages
    ufgData.Text(lngRow, gstrProcedureRep_标本名称) = cbxSpecimenName.Text
    ufgData.Text(lngRow, gstrProcedureRep_报告类型) = GetReportTypeValue(Val(cbxReportType.Text))
    ufgData.Text(lngRow, gstrProcedureRep_报告子项) = GetReportSubValue(IIf(Val(cbxReportType.Text) = 1, Val(cbxReportSub.Text), IIf(Val(cbxReportType.Text) = 2, Val(cbxReportSub.Text) + 2, 0)))
    ufgData.Text(lngRow, gstrProcedureRep_检查结果) = txtResult.Text
    ufgData.Text(lngRow, gstrProcedureRep_检查意见) = txtAdvice.Text
    ufgData.Text(lngRow, gstrProcedureRep_报告日期) = dtServicesTime
    ufgData.Text(lngRow, gstrProcedureRep_当前状态) = "未打印"
    
    mblnEditState = False
End Sub



Private Sub DelCurProcedureRep()
'删除过程报告
    Dim strSql As String
    
    mblnEditState = False
    
    If Not ufgData.IsSelectionRow Then Exit Sub
    
    If ufgData.IsEmptyKey(ufgData.SelectionRow) Then
        If MsgBoxD(Me, "确认要删除所选择的" & Decode(Val(cbxReportType.Text), 0, "冰冻", 1, "免疫", 2, "分子", 3, "特染", "") & "报告吗？删除后报告将不能恢复。", vbYesNo, Me.Caption) = vbNo Then Exit Sub
        Call ufgData.DelRow(ufgData.SelectionRow, False)
        If ufgData.ShowingRowCount <= 1 Or Not ufgData.IsSelectionRow Then Call ClearReportContext
        Exit Sub
    End If
    
    If MsgBoxD(Me, "确认要删除所选择的" & Decode(Val(cbxReportType.Text), 0, "冰冻", 1, "免疫", 2, "分子", 3, "特染", "") & "报告吗？删除后报告将不能恢复。", vbYesNo, Me.Caption) = vbNo Then Exit Sub
    
    strSql = "Zl_病理过程报告_删除(" & ufgData.KeyValue(ufgData.SelectionRow) & ")"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    Call ufgData.DelRow(ufgData.SelectionRow, False)
    
    ClearReportContext
    
    Call EnableReportWithSpeExamType(Val(cbxReportType.Text), mblnReadOnly, GetCurRepAllowAuditing)
    
    '如果有其他过程报告，则载入其他过程报告，否则清除数据
    If Not ufgData.IsSelectionRow Then Exit Sub
    
    If ufgData.IsEmptyKey(ufgData.SelectionRow) Then Exit Sub
    
    Call LoadReportContext(ufgData.SelectionRow)

End Sub


Private Function GetSubReportFormat(ByVal strReportFmt As String, ByVal strRepTag As String) As String
'根据报告tag获取格式名称
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    GetSubReportFormat = ""
    
    strSql = "select b.序号 from zlreports a, zlrptfmts b " & _
                " where a.id = b.报表id and a.编号=upper([1]) and b.说明 like [2]"
                
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strReportFmt, "%" & strRepTag & "%")
    
    If rsData.RecordCount <= 0 Then Exit Function
    
    GetSubReportFormat = rsData!序号
    
End Function


Private Sub SetDcmLabesVisible(dcmImage As DicomImage, ByVal blnVisible As Boolean)
    Dim i As Long
    
    For i = 1 To dcmImage.Labels.Count
        dcmImage.Labels(i).Visible = blnVisible
    Next i
End Sub


Private Function GetReportImageFile() As String
'取得报告图像文件
    Dim i As Long
    Dim strImageFiles As String
    Dim objCurDcmImage As DicomImage
    
    For i = 1 To rpImage.dcmViewer.Images.Count
        If rpImage.ItemSelected(i) Then
            Set objCurDcmImage = rpImage.dcmViewer.Images(i)
           
            '隐藏lab标签
            Call SetDcmLabesVisible(objCurDcmImage, False)
            
            Call objCurDcmImage.FileExport(strCurTempReportPath & objCurDcmImage.InstanceUID & ".jpg", "JPG")
            
            '显示标签
            Call SetDcmLabesVisible(objCurDcmImage, True)
            
            If strImageFiles <> "" Then strImageFiles = strImageFiles & ";"
            strImageFiles = strImageFiles & strCurTempReportPath & objCurDcmImage.InstanceUID & ".jpg"
        End If
    Next i
    
    GetReportImageFile = strImageFiles
End Function


Private Sub PrintCurProcedureRep(Optional ByVal blnIsPrint As Boolean = True)
'打印过程报告
    Dim lngReportType As Long
    Dim lngReportSub As Long
    Dim lngReportID As Long
    Dim strReportFormat As String
    Dim strSubFormat As String
    Dim lngSelectImgCount As Long
    Dim strReportImgFiles As String
    Dim aryImageFiles() As String
    
    If Not ufgData.IsSelectionRow Then Exit Sub
    
    If ufgData.IsEmptyKey(ufgData.SelectionRow) Then
        Call MsgBoxD(Me, "请选择需要打印的报告记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    lngReportID = ufgData.Text(ufgData.SelectionRow, gstrProcedureRep_ID)
    lngReportType = GetReportTypeCode(ufgData.Text(ufgData.SelectionRow, gstrProcedureRep_报告类型))
    lngReportSub = GetReportSubCode(ufgData.Text(ufgData.SelectionRow, gstrProcedureRep_报告子项))
    
    Select Case lngReportType
        Case 0 '冰冻报告
            strReportFormat = "ZL1_Inside_1294_03"
        Case 1 '免疫组化
            If lngReportSub = 1 Then
                '鉴别
                strReportFormat = "ZL1_Inside_1294_15"
            ElseIf lngReportSub = 2 Then
                '多药耐药
                strReportFormat = "ZL1_Inside_1294_04"
            End If
        Case 2 '分子病理
            If lngReportSub = 3 Then
                '荧光
                strReportFormat = "ZL1_Inside_1294_05"
            ElseIf lngReportSub = 4 Then
                '普通
                strReportFormat = "ZL1_Inside_1294_06"
            End If
        Case 3 '特殊染色
            strReportFormat = "ZL1_Inside_1294_14"
    End Select
    
    
    lngSelectImgCount = rpImage.SelectedCount()
    
    strSubFormat = GetSubReportFormat(strReportFormat, lngSelectImgCount & "幅")
    
    If strSubFormat <> "" Then strSubFormat = "ReportFormat=" & strSubFormat
    
    strReportImgFiles = GetReportImageFile()
    
    aryImageFiles = Split(strReportImgFiles & ";;;;;;;;", ";")
    
    Call zlReport.ReportOpen(gcnOracle, 100, strReportFormat, Me, strSubFormat, _
                            "病理号=" & mrecStudyInf.strPatholNumber & "", "过程报告ID=" & lngReportID, _
                            "pic1=" & aryImageFiles(0), _
                            "pic2=" & aryImageFiles(1), _
                            "pic3=" & aryImageFiles(2), _
                            "pic4=" & aryImageFiles(3), _
                            "pic5=" & aryImageFiles(4), _
                            "pic6=" & aryImageFiles(5), _
                            "pic7=" & aryImageFiles(6), _
                            "pic8=" & aryImageFiles(7), IIf(blnIsPrint, 2, 1))
End Sub

Private Sub GetSpeExamResult()
'提取特检结果
If mCurEditText Is Nothing Then Exit Sub
If mCurEditText.Locked Then Exit Sub

Dim frmResultGet As New frmPatholResultGet
On Error GoTo errFree
    Select Case Val(cbxReportType.Text)
        Case 1  '免疫结果
            Call frmResultGet.ShowResultGetWind(mrecStudyInf.lngPatholAdviceId, 0, mstrPrivs, Me)
        Case 2  '分子结果
            Call frmResultGet.ShowResultGetWind(mrecStudyInf.lngPatholAdviceId, 2, mstrPrivs, Me)
        Case 3  '特染结果
            Call frmResultGet.ShowResultGetWind(mrecStudyInf.lngPatholAdviceId, 1, mstrPrivs, Me)
    End Select
    
    If frmResultGet.IsOk Then
        mCurEditText.SelText = frmResultGet.txtResult.Text
    End If
    
errFree:
    Call Unload(frmResultGet)
    Set frmResultGet = Nothing
End Sub



Private Sub UpdateCurProcedureRepState(ByVal lngRPState As Long)
'更新过程报告状态
    Dim strSql As String
    Dim strRPState As String
    
    If Not ufgData.IsSelectionRow Then Exit Sub
        
    If ufgData.IsEmptyKey(ufgData.SelectionRow) Then
        Call MsgBoxD(Me, "请选择需要进行该操作的过程报告。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If lngRPState = M_REPORTSTATE_CANCEL Then
        If MsgBoxD(Me, "确认要撤回该报告吗？", vbYesNo, Me.Caption) = vbNo Then Exit Sub
    End If
    
    strSql = "Zl_病理过程报告_状态(" & ufgData.KeyValue(ufgData.SelectionRow) & "," & lngRPState & ")"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    '同步更新列表状态显示
    strRPState = ""
    Select Case lngRPState
        Case M_REPORTSTATE_NORMAL
            strRPState = "未打印"
        Case M_REPORTSTATE_VIEW
            strRPState = "已查阅"
            
            Call EnableReportWithSpeExamType(Val(cbxReportType.Text), mblnReadOnly, True)
        Case M_REPORTSTATE_CANCEL
            strRPState = "已撤回"
            
            Call EnableReportWithSpeExamType(Val(cbxReportType.Text), mblnReadOnly, False)
        Case M_REPORTSTATE_PRINT
            strRPState = "已打印"
            
            Call EnableReportWithSpeExamType(Val(cbxReportType.Text), mblnReadOnly, True)
    End Select
    
    ufgData.Text(ufgData.SelectionRow, gstrProcedureRep_当前状态) = strRPState
End Sub


Private Sub Form_Initialize()
    Set mCurEditText = txtResult
    Set zlReport = New zl9Report.clsReport
    
    mblnEditState = False
    
    
    strCurTempReportPath = IIf(Len(App.Path) > 3, App.Path & "\", App.Path) & "TmpReportImg\"
    
    '如果目录存在，则删除临时报告目录
    If Dir(strCurTempReportPath, vbDirectory) <> "" Then
      Call mdlDir.DeleteFolder(strCurTempReportPath)
    End If
    
    '判断临时报告目录是否存在，如补存在则创建
    If Dir(strCurTempReportPath, vbDirectory) = "" Then
        Call MkDir(strCurTempReportPath)
    End If
End Sub

Private Sub LoadReportModule(Optional blnRefresh As Boolean = False)
'载入报告模板
    Dim strLinkClassName As String
    
    If mlngCurDeptId = wimWord.CurDepartId And Not blnRefresh Then Exit Sub
    
    Select Case Val(cbxReportType.Text)
        Case 0
            strLinkClassName = zlDatabase.GetPara("常规报告模板", glngSys, glngModul, "")
        Case 1
            strLinkClassName = zlDatabase.GetPara("免疫报告模板", glngSys, glngModul, "")
        Case 2
            strLinkClassName = zlDatabase.GetPara("分子报告模板", glngSys, glngModul, "")
        Case 3
            strLinkClassName = zlDatabase.GetPara("特染报告模板", glngSys, glngModul, "")
    End Select
    
    wimWord.ModuleName = strLinkClassName
    wimWord.CurDepartId = mlngCurDeptId
    
    Call wimWord.LoadWordModel
End Sub


Private Sub Form_Load()
On Error GoTo ErrHandle

    dtpReportTime.value = zlDatabase.Currentdate
    Call InitCommandBars
    
    '初始化列表
    Call InitProcedureRepList
    
    '载入报告类型
    Call LoadReportType
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call MakeSureShowMsg
    Set zlReport = Nothing
End Sub

Private Sub picReportEdit_Resize()
On Error Resume Next
    Dim lngAvgHeight As Long

    framReport.Left = 0
    framReport.Top = 0
    framReport.Width = picReportEdit.Width
    framReport.Height = picReportEdit.Height - 120
    
    lngAvgHeight = Round((framReport.Height - cbxReportType.Height - 120 * 10) / 3)
    
    labReportType.Left = 120
    labReportType.Top = 300
    
    cbxReportType.Left = labReportType.Left + labReportType.Width + 120
    cbxReportType.Top = 240
    
    labReportSub.Left = cbxReportType.Left + cbxReportType.Width + 240
    labReportSub.Top = labReportType.Top
    
    cbxReportSub.Left = labReportSub.Left + labReportSub.Width + 120
    cbxReportSub.Top = 240
    
    labSpecimenName.Left = cbxReportSub.Left + cbxReportSub.Width + 240
    labSpecimenName.Top = labReportType.Top
    
    cbxSpecimenName.Left = labSpecimenName.Left + labSpecimenName.Width + 120
    cbxSpecimenName.Top = 240
    
    Line1.X1 = 120
    Line1.X2 = framReport.Width - 120
    
    labResult.Left = 120
    
    '下面是一个公式，字号为9,12,15时 Y分别为650,715,780
    Line1.Y1 = 455 + (65 / 3) * mbytFontSize
    Line1.Y2 = Line1.Y1
    labResult.Top = Line1.Y1 + 70
    
    txtResult.Width = framReport.Width - 240
    txtResult.Top = labResult.Top + labResult.Height + 30
    txtResult.Width = framReport.Width - 240
    txtResult.Height = lngAvgHeight
    
    labAdvice.Left = 120
    labAdvice.Top = txtResult.Top + txtResult.Height + 120
    
    txtAdvice.Left = 120
    txtAdvice.Top = labAdvice.Top + labAdvice.Height + 60
    txtAdvice.Width = framReport.Width - 240
    txtAdvice.Height = lngAvgHeight - 260
    
    labReportTime.Left = txtAdvice.Width - dtpReportTime.Width - labReportTime.Width
    labReportTime.Top = labAdvice.Top
    
    dtpReportTime.Left = labReportTime.Left + labReportTime.Width + 120
    dtpReportTime.Top = labReportTime.Top - 60

    
    rpImage.Left = 120
    rpImage.Top = txtAdvice.Top + txtAdvice.Height + 120
    rpImage.Width = framReport.Width - 240
    rpImage.Height = lngAvgHeight
    
    '下面的调整防止图像下端超过界面
    If mbytFontSize = C_INT_FONTSISE_SMALL Then
        rpImage.Height = lngAvgHeight
    ElseIf mbytFontSize = C_INT_FONTSISE_MEDIUM Then
        rpImage.Height = lngAvgHeight - 75
    Else
        rpImage.Height = lngAvgHeight - 150
    End If
End Sub


Private Sub picWordModule_Resize()
On Error Resume Next
    framWord.Left = 0
    framWord.Top = 0
    framWord.Width = picWordModule.Width
    framWord.Height = picWordModule.Height - 120
    
    wimWord.Left = 120
    wimWord.Top = 240
    wimWord.Width = framWord.Width - 240
    wimWord.Height = Round(framWord.Height / 3 * 2) - 240
    
    ufgData.Left = 120
    ufgData.Top = wimWord.Top + wimWord.Height + 120
    ufgData.Width = framWord.Width - 240
    ufgData.Height = Round(framWord.Height / 3) - 240
End Sub



Private Sub rpImage_SelectedChange(ByVal lngSelectedIndex As Long, ByVal blnSelected As Boolean)
    If mblnIsLoading = True Then Exit Sub
    mblnEditState = True
End Sub

Private Sub txtAdvice_Change()
    If mblnIsLoading = True Then Exit Sub
        
    mblnEditState = True
End Sub

Private Sub txtMemo_Change()
    If mblnIsLoading = True Then Exit Sub
    
    mblnEditState = True
End Sub

Private Sub txtAdvice_GotFocus()
    Set mCurEditText = txtAdvice
End Sub

Private Sub txtResult_Change()
    If mblnIsLoading = True Then Exit Sub
    
    mblnEditState = True
End Sub

Private Sub txtResult_GotFocus()
    Set mCurEditText = txtResult
End Sub

Private Sub ufgData_OnSelChange()
'载入报告内容
On Error GoTo ErrHandle

    If ufgData.ShowingRowCount < 1 Or Not ufgData.IsSelectionRow Then
        
        Exit Sub
    End If
    
    Call MakeSureShowMsg
    mlngRow = ufgData.SelectionRow

    Call ClearReportContext
    Call EnableReportWithSpeExamType(Val(cbxReportType.Text), mblnReadOnly, GetCurRepAllowAuditing, False)
    
    If ufgData.SelectionRow = 0 Then Exit Sub
    If ufgData.IsEmptyKey(ufgData.SelectionRow) Then Exit Sub
    
    Call LoadReportContext(ufgData.SelectionRow)
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub wimWord_OnWordDbClickEvent(ByVal strWord As String)
'载入词句
On Error GoTo ErrHandle
    If Not mCurEditText.Locked Then mCurEditText.SelText = strWord
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub zlReport_AfterPrint(ByVal ReportNum As String)
On Error GoTo ErrHandle
    
    '打印后保存已打印的报告
    If mblnEditState Then Call SaveCurProcedureRep
    
    '修改当前报告状态
    Call UpdateCurProcedureRepState(M_REPORTSTATE_PRINT)
    
    '打印后不允许编辑
    Call EnableReportWithSpeExamType(Val(cbxReportType.Text), mblnReadOnly, True)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub



Private Sub cbrMain_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error GoTo ErrorHand
    Select Case control.ID
        Case conMenu_PatholProRep_Report, conMenu_PatholProRep_Preview
            '预览过程报告
            Call PrintCurProcedureRep(False)
            
        Case conMenu_PatholProRep_Print
            '打印过程报告
            Call PrintCurProcedureRep(True)
            
        Case conMenu_PatholProRep_Already
            '过程报告查阅
            Call MakeSureShowMsg(True)
            Call UpdateCurProcedureRepState(M_REPORTSTATE_VIEW)
            
        Case conMenu_PatholProRep_Back
            '撤回过程报告
            Call UpdateCurProcedureRepState(M_REPORTSTATE_CANCEL)
            
        Case conMenu_PatholProRep_Clear
            '清除录入内容,
            Call ClearReportContext
            
        Case conMenu_PatholProRep_Input
            '项目录入
            Call GetSpeExamResult
            
        Case conMenu_PatholProRep_New
            '新增报告
            If mblnEditState Then
                If PromptSave() Then
                    Call SaveCurProcedureRep
                End If
            End If
            
            Call NewProcedureRep
            
        Case conMenu_PatholProRep_Del
            '删除报告（执行删除后更新查阅按钮状态）
            Call DelCurProcedureRep
            
        Case conMenu_PatholProRep_Save
            '保存报告
            Call SaveCurProcedureRep
    End Select
    Exit Sub
ErrorHand:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cbrMain_Update(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error GoTo ErrorHand
    Dim blnIsPopedom As Boolean
    Dim blGetCurRepAllowAuditing As Boolean     '    是否“已查阅、或者已打印”
    Dim blHavePatholNumber     '是否有病理号
    
    Select Case Val(cbxReportType.Text)
        Case 0:
            blnIsPopedom = CheckPopedom(mstrPrivs, "冰冻报告")
        Case 1
            blnIsPopedom = CheckPopedom(mstrPrivs, "免疫报告")
        Case 2, 3
            blnIsPopedom = CheckPopedom(mstrPrivs, "分子报告")
        Case 4
            blnIsPopedom = CheckPopedom(mstrPrivs, "特染报告")
    End Select
    
    blGetCurRepAllowAuditing = GetCurRepAllowAuditing
    blHavePatholNumber = (Len(mrecStudyInf.strPatholNumber) > 0)
    
    Select Case control.ID
        Case conMenu_PatholProRep_Report
            control.Enabled = (mblnIsAllowSpeExam Or mblnIsAllowWriteReport) And blHavePatholNumber And ufgData.ShowingRowCount > 0
            
        Case conMenu_PatholProRep_Preview
            control.Enabled = (mblnIsAllowSpeExam Or mblnIsAllowWriteReport) And blHavePatholNumber
            
        Case conMenu_PatholProRep_Print
            control.Enabled = (mblnIsAllowSpeExam Or mblnIsAllowWriteReport) And blHavePatholNumber
            
        Case conMenu_PatholProRep_Already
            control.Enabled = mblnIsAllowWriteReport And Not mblnReadOnly And Not blGetCurRepAllowAuditing And blHavePatholNumber And ufgData.ShowingRowCount > 0
            
        Case conMenu_PatholProRep_Back
            control.Enabled = mblnIsAllowWriteReport And Not mblnReadOnly And blGetCurRepAllowAuditing And blHavePatholNumber And ufgData.ShowingRowCount > 0
            
        Case conMenu_PatholProRep_Input
            control.Enabled = Not (mblnReadOnly Or blGetCurRepAllowAuditing) And blnIsPopedom And Val(cbxReportType.Text) > 0 And blHavePatholNumber
            
        Case conMenu_PatholProRep_Clear
            control.Enabled = Not (mblnReadOnly Or blGetCurRepAllowAuditing) And blnIsPopedom And blHavePatholNumber
            
        Case conMenu_PatholProRep_New
            control.Enabled = Not mblnReadOnly And blnIsPopedom And blHavePatholNumber
            
        Case conMenu_PatholProRep_Del
            control.Enabled = Not (mblnReadOnly Or blGetCurRepAllowAuditing) And blnIsPopedom And blHavePatholNumber
            
        Case conMenu_PatholProRep_Save
            control.Enabled = Not (mblnReadOnly Or blGetCurRepAllowAuditing) And blnIsPopedom And blHavePatholNumber And mblnEditState
    End Select
    
    Exit Sub
ErrorHand:
    If ErrCenter() = 1 Then Resume
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

        Set cbrControl = .Add(xtpControlSplitButtonPopup, conMenu_PatholProRep_Report, "报告打印"): cbrControl.IconId = 5012: cbrControl.ToolTipText = "报告打印"
            With cbrControl
                Set cbrPopControl = CreateModuleMenu(.CommandBar.Controls, xtpControlButton, conMenu_PatholProRep_Preview, "预览", "", 0, False)
                Set cbrPopControl = CreateModuleMenu(.CommandBar.Controls, xtpControlButton, conMenu_PatholProRep_Print, "打印", "", 0, False)
            End With
        Set cbrControl = .Add(xtpControlButton, conMenu_PatholProRep_Already, "报告查阅"): cbrControl.IconId = 5013: cbrControl.ToolTipText = "报告查阅"
        Set cbrControl = .Add(xtpControlButton, conMenu_PatholProRep_Back, "报告撤回"): cbrControl.IconId = 5014: cbrControl.ToolTipText = "报告撤回"
        
        
        Set cbrControl = .Add(xtpControlButton, conMenu_PatholProRep_Clear, "清空内容"): cbrControl.IconId = 5015: cbrControl.ToolTipText = "清空内容"
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_PatholProRep_Input, "项目录入"): cbrControl.IconId = 5016: cbrControl.ToolTipText = "项目录入"
        
        
        Set cbrControl = .Add(xtpControlButton, conMenu_PatholProRep_New, "新增报告"): cbrControl.IconId = 5017: cbrControl.ToolTipText = "新增报告"
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_PatholProRep_Del, "删除报告"): cbrControl.IconId = 5018: cbrControl.ToolTipText = "删除报告"
        
        
        Set cbrControl = .Add(xtpControlButton, conMenu_PatholProRep_Save, "保存报告"): cbrControl.IconId = 5005: cbrControl.ToolTipText = "保存报告"
        cbrControl.BeginGroup = True
    End With
    Exit Sub
errH:
End Sub

Private Function PromptSave() As Boolean
'提示保存
    If Len(txtAdvice.Text) > 0 Or Len(txtResult.Text) > 0 Then
        PromptSave = (MsgBoxD(Me, "操作前是否保存刚才被修改的报告？", vbYesNo, Me.Caption) = vbYes)
        mblnEditState = False
    End If
End Function

Private Sub MakeSureShowMsg(Optional ByVal blIsAlReady As Boolean)
'判断是否需要弹出是否保存对话
'分为新增部分和修改部分新增数据没有过程ID，修改数据有过程ID
'blIsAlReady 是否查阅功能，若是，不保存会重新加载显示修改前的内容
    Dim blNeed As Boolean
    Dim intNeed As Integer
  
    If Val(ufgData.Text(mlngRow, gstrProcedureRep_ID)) > 0 And mblnEditState = True Then
        mblnEditState = False
        intNeed = MsgBoxD(Me, "刚才打开的报告有所改变，是否保存？", vbYesNo, gstrSysName)
        If intNeed = vbYes Then
            Call SaveCurProcedureRep
        Else
            If blIsAlReady Then Call LoadReportContext(mlngRow)
        End If
    End If
End Sub
