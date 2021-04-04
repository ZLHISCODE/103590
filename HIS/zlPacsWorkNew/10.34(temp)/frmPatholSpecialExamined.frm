VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPatholSpecialExamined 
   Caption         =   "特殊检查"
   ClientHeight    =   9525
   ClientLeft      =   75
   ClientTop       =   405
   ClientWidth     =   10875
   Icon            =   "frmPatholSpecialExamined.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9525
   ScaleWidth      =   10875
   StartUpPosition =   3  '窗口缺省
   Begin MSComctlLib.ImageList imgTbrS 
      Left            =   9195
      Top             =   675
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholSpecialExamined.frx":000C
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholSpecialExamined.frx":0C7E
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholSpecialExamined.frx":18F0
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholSpecialExamined.frx":2562
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholSpecialExamined.frx":31D4
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholSpecialExamined.frx":3E46
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholSpecialExamined.frx":4AB8
            Key             =   "IMG7"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrMain 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   1270
      ButtonWidth     =   1455
      ButtonHeight    =   1270
      Style           =   1
      ImageList       =   "imgTbrS"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "标签打印"
            Key             =   "tbLAB"
            ImageIndex      =   4
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "tbLabPreview"
                  Text            =   "预览"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "tblabPrint"
                  Text            =   "打印"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "清单打印"
            Key             =   "tbList"
            ImageIndex      =   5
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "tbListPreview"
                  Text            =   "预览"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "tbListPrint"
                  Text            =   "打印"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "-"
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "申请查看"
            Key             =   "tbViewRequest"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "特检接受"
            Key             =   "tbAcceptSpeExam"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "特检完成"
            Key             =   "tbEndSpeExam"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin VB.Frame framSpeExam 
      Caption         =   "特检记录"
      Height          =   7215
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   9975
      Begin VB.Frame FramCheck 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   5760
         TabIndex        =   6
         Top             =   6840
         Width           =   4095
         Begin VB.CheckBox chkYSQ 
            Caption         =   "已申请"
            Height          =   255
            Left            =   1080
            TabIndex        =   9
            Top             =   -7
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chkYJS 
            Caption         =   "已接受"
            Height          =   180
            Left            =   2160
            TabIndex        =   8
            Top             =   30
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chkYWC 
            Caption         =   "已完成"
            Height          =   180
            Left            =   3120
            TabIndex        =   7
            Top             =   30
            Width           =   855
         End
      End
      Begin VB.OptionButton optFenZi 
         Caption         =   "分子病理"
         Height          =   255
         Left            =   3360
         TabIndex        =   2
         Tag             =   "2"
         Top             =   0
         Width           =   1095
      End
      Begin VB.OptionButton optTeShu 
         Caption         =   "特殊染色"
         Height          =   255
         Left            =   2160
         TabIndex        =   1
         Tag             =   "1"
         Top             =   0
         Width           =   1095
      End
      Begin VB.OptionButton optMianYi 
         Caption         =   "免疫组化"
         Height          =   255
         Left            =   960
         TabIndex        =   0
         Tag             =   "0"
         Top             =   0
         Value           =   -1  'True
         Width           =   1095
      End
      Begin zl9PACSWork.ucFlexGrid ufgData 
         Height          =   6015
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   10610
         GridRows        =   21
         IsBtnNextCell   =   0   'False
         IsCopyAdoMode   =   0   'False
         IsEjectConfig   =   -1  'True
         HeadFontCharset =   134
         HeadFontWeight  =   400
         DataFontCharset =   134
         DataFontWeight  =   400
         ExtendLastCol   =   -1  'True
      End
      Begin VB.Label labRecordInf 
         Caption         =   "当前总项目数：0    当前需检查项目数：0"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   6840
         Width           =   5295
      End
   End
End
Attribute VB_Name = "frmPatholSpecialExamined"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IWorkMenu

Private Const C_INT_MIANYI As Integer = 0
Private Const C_INT_TESHU As Integer = 1
Private Const C_INT_FENZI As Integer = 2

Private Const M_STR_HINT_NoSelectData As String = "无效的检查数据，请选择需要执行的检查记录。"
Private Const M_STR_MODULE_MENU_TAG As String = "特检"

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

Private mrecStudyInf As TStudyStateInf
Private mblnReadOnly As Boolean

Private mblnAutoAcceptOfAfterPrint As Boolean '打印后自动接受

Private mlngTmpAdviceId As Long
Private mlngTmpSendNo As Long


Private mObjActiveMenuBar As CommandBar
Private mbytFontSize As Byte '字号    9--小字体    12--大字体

Private mblnRefreshState As Boolean

Private mKeyCode As Long
Private mKeyShift As Long



'获取菜单接口对象
Property Get zlMenu() As IWorkMenu
    Set zlMenu = Me
End Property


Public Sub NotificationRefresh()
'通知刷新
    mblnRefreshState = False
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

    If Not HasMenu(objMenuBar, conMenu_PatholSpeExam) Then
        Set cbrMenuBar = mObjActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_PatholSpeExam, "特检(&T)", GetPatholMenuIndex(objMenuBar) + 1, False)
        cbrMenuBar.ID = conMenu_PatholSpeExam
        cbrMenuBar.Category = M_STR_MODULE_MENU_TAG
        
        
        With cbrMenuBar.CommandBar
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlPopup, conMenu_PatholSpeExam_LAB, "标签打印(&B)", "", 1, False)
                Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_PatholSpeExam_PreviewLAB, "预览(V)", "", 1, False)
                Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_PatholSpeExam_PrintLab, "打印(P)", "", 1, False)
                
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlPopup, conMenu_PatholSpeExam_List, "清单打印(&L)", "", 1, False)
                Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_PatholSpeExam_PreviewList, "预览(V)", "", 1, False)
                Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_PatholSpeExam_PrintList, "打印(P)", "", 1, False)
            
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholSpeExam_RequestView, "申请查看(&Q)", "", 1, True)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholSpeExam_Accept, "特检接受(&R)", "", 1, False)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholSpeExam_Finish, "特检完成(&F)", "", 1, False)
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
        Case conMenu_PatholSpeExam_PreviewLAB
            Call PrintSpeExamLabel(False)
            
        Case conMenu_PatholSpeExam_PrintLab
            Call PrintSpeExamLabel(True)
            
        Case conMenu_PatholSpeExam_PreviewList
            Call PrintWorkList(False)
            
        Case conMenu_PatholSpeExam_PrintList
            Call PrintWorkList(True)
            
        Case conMenu_PatholSpeExam_RequestView
            Call ShowSpeExamRequest
            
        Case conMenu_PatholSpeExam_Accept
            Call SpeExamined_Accept
            
        Case conMenu_PatholSpeExam_Finish
            Call SpeExamined_Sure
    End Select
End Sub


Public Sub IWorkMenu_zlUpdateMenu(ByVal control As XtremeCommandBars.ICommandBarControl)
'更新菜单
    Dim blnIsAllowSpeExam As Boolean

    If Not Me.Visible Then
        control.Enabled = False
        Exit Sub
    End If
    
    blnIsAllowSpeExam = CheckPopedom(mstrPrivs, "免疫组化") Or CheckPopedom(mstrPrivs, "分子病理") Or CheckPopedom(mstrPrivs, "特殊染色") And mlngAdviceID > 0
    
    Select Case control.ID
        Case conMenu_PatholSpeExam_LAB
            control.Enabled = blnIsAllowSpeExam And mrecStudyInf.strPatholNumber <> ""
            
        Case conMenu_PatholSpeExam_List
            control.Enabled = blnIsAllowSpeExam And mrecStudyInf.strPatholNumber <> ""
            
        Case conMenu_PatholSpeExam_RequestView
            control.Enabled = blnIsAllowSpeExam And mrecStudyInf.strPatholNumber <> ""
            
        Case conMenu_PatholSpeExam_Accept
            control.Enabled = blnIsAllowSpeExam And Not mblnReadOnly
            
        Case conMenu_PatholSpeExam_Finish
            control.Enabled = blnIsAllowSpeExam And Not mblnReadOnly
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
    If blnMoved Or blnMoved Or lngStudyState = 6 Or lngStudyState = 5 Or lngStudyState = 0 Or lngStudyState = 1 Or lngStudyState = -2 Then
        mblnReadOnly = True
    End If

End Sub

Public Sub zlRefreshFace(Optional blnForceRefresh As Boolean = False)
'刷新界面数据
    If (mlngTmpAdviceId = mlngAdviceID And mlngTmpSendNo = mlngSendNo And mblnRefreshState) And Not blnForceRefresh Then Exit Sub
    
    mlngTmpAdviceId = mlngAdviceID
    mlngTmpSendNo = mlngSendNo
    mblnRefreshState = True
    
    If mlngAdviceID <= 0 Then
        Call ConfigSpeExamFace(False, "医嘱ID无效请检查。")
        Exit Sub
    End If
    
   Call GetPatholStudyState(mlngAdviceID, mrecStudyInf)
        
    '清除显示数据
    Call ufgData.ClearListData
       
    If mrecStudyInf.strPatholNumber = "" Then
        Call ConfigSpeExamFace(False, "该检查尚未生成有效的病理号，请确认该检查是否已被核收。")
        Exit Sub
    Else
        Call ConfigSpeExamFace(True)
    End If
    
    Call ConfigSpeExamType(mrecStudyInf.lngPatholAdviceId)
    
    Call QuerySpeExamData(mrecStudyInf.lngPatholAdviceId)
    
    Call RefreshSpeExamInf
    
    Call ConfigPopedom(mblnReadOnly)
    
'    If Not (owner Is Nothing) Then
'        Call Me.Show(1, owner)
'    End If
End Sub

Public Sub zlRefresh(lngAdviceID As Long, ByVal blnReadOnly As Boolean, _
    strPrivs As String, ByVal blnMoved As Boolean, _
    ByVal lngCurDepartmentId As Long, Optional owner As Form = Nothing)
    
    If lngAdviceID <= 0 Then
        Call ConfigSpeExamFace(False, "医嘱ID无效请检查。")
        Exit Sub
    End If
    
'    If mlngCurAdviceId = lngAdviceID Then
'        Call ConfigSpeExamFace(False, "该检查尚未生成有效的病理号，请确认该检查是否已被核收。")
'        Exit Sub
'    End If
    
    mlngAdviceID = lngAdviceID
    mstrPrivs = strPrivs
    mblnMoved = blnMoved
    mlngCurDeptId = lngCurDepartmentId
    
    Call GetPatholStudyState(lngAdviceID, mrecStudyInf)
        
    '清除显示数据
    Call ufgData.ClearListData
       
    If mrecStudyInf.strPatholNumber = "" Then
        Call ConfigSpeExamFace(False, "该检查尚未生成有效的病理号，请确认该检查是否已被核收。")
        Exit Sub
    Else
        Call ConfigSpeExamFace(True)
    End If
    
    Call ConfigSpeExamType(mrecStudyInf.lngPatholAdviceId)
    
    Call QuerySpeExamData(mrecStudyInf.lngPatholAdviceId)
    
    Call RefreshSpeExamInf
    
    Call ConfigPopedom(blnReadOnly)
    
    If Not (owner Is Nothing) Then
        Call Me.Show(1, owner)
    End If
End Sub




Private Sub RefreshSpeExamInf()
'刷新制片记录数量
    Dim i As Long
    Dim lngNeedCount As Long
    Dim lngTotal As Long
    
    lngNeedCount = 0
    lngTotal = 0
    
    
    For i = 1 To ufgData.GridRows - 1
        If Not ufgData.IsNullRow(i) Then
        
            
            lngTotal = lngTotal + 1
                
            If ufgData.Text(i, gstrSpeExam_当前状态) <> "已完成" Then
                lngNeedCount = lngNeedCount + 1
            End If
        End If
    Next i
    
    labRecordInf.Caption = "当前总项目数：" & lngTotal & "    当前需检查项目数：" & lngNeedCount
    
End Sub


Private Sub ConfigPopedom(ByVal blnIsReadOnly As Boolean)
'配置权限
    Dim blnIsAllowSpeExam As Boolean
    
    blnIsAllowSpeExam = CheckPopedom(mstrPrivs, "免疫组化") Or CheckPopedom(mstrPrivs, "分子病理") Or CheckPopedom(mstrPrivs, "特殊染色")
    
    tbrMain.Buttons("tbAcceptSpeExam").Enabled = blnIsAllowSpeExam And Not blnIsReadOnly
    tbrMain.Buttons("tbEndSpeExam").Enabled = blnIsAllowSpeExam And Not blnIsReadOnly
    
    tbrMain.Buttons("tbViewRequest").Enabled = blnIsAllowSpeExam
    tbrMain.Buttons("tbLAB").Enabled = blnIsAllowSpeExam
    tbrMain.Buttons("tbList").Enabled = blnIsAllowSpeExam
    
    ufgData.ReadOnly = blnIsReadOnly
    
    optMianYi.Enabled = CheckPopedom(mstrPrivs, "免疫组化")
    optTeShu.Enabled = CheckPopedom(mstrPrivs, "特殊染色")
    optFenZi.Enabled = CheckPopedom(mstrPrivs, "分子病理")
End Sub

Private Sub ConfigSpeExamFace(ByVal blnIsValid As Boolean, Optional ByVal strHintInf As String = "")
'配置特检界面
    tbrMain.Buttons("tbAcceptSpeExam").Enabled = blnIsValid
    tbrMain.Buttons("tbEndSpeExam").Enabled = blnIsValid
    
    tbrMain.Buttons("tbViewRequest").Enabled = blnIsValid
    tbrMain.Buttons("tbLAB").Enabled = blnIsValid
    tbrMain.Buttons("tbList").Enabled = blnIsValid

    optFenZi.Enabled = blnIsValid
    optMianYi.Enabled = blnIsValid
    optTeShu.Enabled = blnIsValid

    chkYSQ.Enabled = blnIsValid
    chkYJS.Enabled = blnIsValid
    chkYWC.Enabled = blnIsValid
    
    If blnIsValid Then
        Call ufgData.CloseHintInf
    Else
        Call ufgData.ShowHintInf(strHintInf)
        labRecordInf.Caption = ""
    End If
End Sub


Private Function GetSelectSpeExamType()
'取得选择的特检类型
    If optMianYi.value Then GetSelectSpeExamType = C_INT_MIANYI
    If optTeShu.value Then GetSelectSpeExamType = C_INT_TESHU
    If optFenZi.value Then GetSelectSpeExamType = C_INT_FENZI
End Function


Private Sub AdjustFace()
    '调整界面布局
    framSpeExam.Left = 0
    framSpeExam.Top = tbrMain.Top + tbrMain.Height + 120
    framSpeExam.Width = Me.Width - 0
    framSpeExam.Height = Me.Height - tbrMain.Height - 240
    
    
    ufgData.Left = 120
    ufgData.Top = 280 + IIf(mbytFontSize = 9, 0, 120)
    ufgData.Width = framSpeExam.Width - 240
    ufgData.Height = framSpeExam.Height - labRecordInf.Height - 600
    
    labRecordInf.Left = 120
    labRecordInf.Top = framSpeExam.Height - labRecordInf.Height - 120
    
    
    '调整FrameCheck位置
    
     FramCheck.Top = framSpeExam.Height - labRecordInf.Height - 120
     FramCheck.Left = framSpeExam.Width - FramCheck.Width - 200
     
     chkYJS.Top = 0
     chkYSQ.Top = 0
     chkYWC.Top = 0
End Sub

Private Sub ConfigSpeExamType(ByVal strPatholNum As String)
'配置当前特检类型
    Dim lngType As Long
    
    lngType = GetCurSpeExamType(strPatholNum)
    
    Select Case lngType
        Case C_INT_MIANYI
            optMianYi.value = True
        Case C_INT_TESHU
            optTeShu.value = True
        Case C_INT_FENZI
            optFenZi.value = True
    End Select

End Sub


Private Sub InitSpeExamList()
'初始化特检列表
    Dim strTemp As String
    
    ufgData.IsKeepRows = True
    ufgData.GridRows = glngMaxRowCount
    '设置行高
    ufgData.RowHeightMin = glngStandardRowHeight
    
    ufgData.IsCopyMode = True
    
    '判断数据库参数表是否有数据 有则读取数据库参数  没有则加载默认
    strTemp = zlDatabase.GetPara("特检信息列表配置", glngSys, G_LNG_PATHOLSYS_NUM, "")
     
    If strTemp = "" Then
        ufgData.ColNames = gstrSpeExamCols
    Else
        ufgData.ColNames = strTemp
    End If
    
    ufgData.DefaultColNames = gstrSpeExamCols
    ufgData.ColConvertFormat = gstrSpeExamConvertFormat
End Sub

Private Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo errHandle
    Call ExecuteTbrOperation(Button.Key)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub tbrMain_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
On Error GoTo errHandle
    Call ExecuteTbrOperation(ButtonMenu.Key)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ExecuteTbrOperation(ByVal strButtonKey As String)
    Dim strKey As String
    
    strKey = UCase(strButtonKey)
    
    Select Case strKey
        Case UCase("tbLab"), UCase("tbLabPreview")
            '预览标签
            Call PrintSpeExamLabel(False)
            
        Case UCase("tbLabPrint")
            '打印标签
            Call PrintSpeExamLabel(True)
        
        Case UCase("tbList"), UCase("tbListPreview")
            '预览清单
            Call PrintWorkList(False)
        
        Case UCase("tbListPrint")
            '打印清单
            Call PrintWorkList(True)
        
        Case UCase("tbViewRequest")
            '查看申请
            Call ShowSpeExamRequest
        
        Case UCase("tbAcceptSpeExam")
            '特检接受
            Call SpeExamined_Accept
        
        Case UCase("tbEndSpeExam")
            '特检完成
            Call SpeExamined_Sure
            
    End Select
End Sub


Private Sub ufgData_OnColFormartChange()
 '保存列表参数
     zlDatabase.SetPara "特检信息列表配置", ufgData.GetColsString(ufgData), glngSys, G_LNG_PATHOLSYS_NUM
End Sub


Private Function GetCurSpeExamType(ByVal lngPatholAdviceId As Long) As Long
'取得当前特检类型
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    
    
    strSql = "select 免疫过程,分子过程,特染过程 from 病理检查信息 where 病理医嘱ID=[1]"
'    If mblnMoved Then strSql = GetMovedDataSql(strSql)
    
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPatholAdviceId)
    
    If rsData.RecordCount > 0 Then
    
        If Nvl(rsData!免疫过程) > 0 Then
            GetCurSpeExamType = 0
            If CheckPopedom(mstrPrivs, "免疫组化") Then Exit Function
        End If
        
        If Nvl(rsData!特染过程) > 0 Then
            GetCurSpeExamType = 1
            If CheckPopedom(mstrPrivs, "特殊染色") Then Exit Function
        End If
        
        If Nvl(rsData!分子过程) > 0 Then
            GetCurSpeExamType = 2
            If CheckPopedom(mstrPrivs, "分子病理") Then Exit Function
        End If
    End If
    
    
    
    If CheckPopedom(mstrPrivs, "免疫组化") Then
        GetCurSpeExamType = 0
        Exit Function
    End If
    
    If CheckPopedom(mstrPrivs, "特殊染色") Then
        GetCurSpeExamType = 1
        Exit Function
    End If
    
    If CheckPopedom(mstrPrivs, "分子病理") Then
        GetCurSpeExamType = 2
        Exit Function
    End If
End Function


Private Sub QuerySpeExamData(ByVal lngPatholAdviceId As Long)
'载入特检数据（包括免疫组化，分子病理，特殊染色）
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    strSql = "select to_Number(a.ID) as ID,a.材块ID,to_Number(b.序号) as 序号,b.标本名称,to_Number(a.申请ID) as 申请ID,to_Number(a.抗体ID) as 抗体ID,c.抗体名称,to_Number(a.制作类型) as 制作类型,to_Number(a.当前状态) as 当前状态,a.项目结果, d.申请时间, a.完成时间,a.特检医师,to_Number(a.清单状态) as 清单状态,to_Number(a.特检类型) as 特检类型,to_Number(a.特检细目) as 特检细目" & _
            " from 病理特检信息 a, 病理取材信息 b, 病理抗体信息 c, 病理申请信息 d " & _
            " where a.材块id=b.材块id and a.抗体id=c.抗体id and a.申请ID=d.申请ID and b.病理医嘱ID=[1] and (特检类型=-1"
'    If mblnMoved Then strSql = GetMovedDataSql(strSql)
    
    If CheckPopedom(mstrPrivs, "免疫组化") Then strSql = strSql & " or 特检类型=0"
    If CheckPopedom(mstrPrivs, "特殊染色") Then strSql = strSql & " or 特检类型=1"
    If CheckPopedom(mstrPrivs, "分子病理") Then strSql = strSql & " or 特检类型=2"
        
    strSql = strSql & ") order by 特检类型,当前状态,序号,ID"
            
    Set ufgData.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPatholAdviceId)
    
    '读取对应特检项目数据 并过滤数据
    Call FilterData
End Sub



Public Sub ReSetFormFontSize(ByVal bytFontSize As Byte)
'功能:重新设置工作站窗体的字体大小
On Error GoTo errHandle

    Dim objCtrl As control
    Dim CtlFont As StdFont
    Dim strFontType  As String
    
    
    
    mbytFontSize = bytFontSize
    
    optMianYi.Left = optMianYi.Left + IIf(optMianYi.FontSize = bytFontSize, 0, IIf(bytFontSize = 9, -300, 300))
    optTeShu.Left = optTeShu.Left + IIf(optTeShu.FontSize = bytFontSize, 0, IIf(bytFontSize = 9, -500, 500))
    optFenZi.Left = optFenZi.Left + IIf(optFenZi.FontSize = bytFontSize, 0, IIf(bytFontSize = 9, -700, 700))
    
    chkYSQ.Left = chkYSQ.Left + IIf(chkYSQ.FontSize = bytFontSize, 0, IIf(bytFontSize = 9, 500, -500))
    chkYJS.Left = chkYJS.Left + IIf(chkYSQ.FontSize = bytFontSize, 0, IIf(bytFontSize = 9, 300, -300))
    
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
        End Select
    Next
    
    Call AdjustFace
    
    Exit Sub
errHandle:
End Sub



Private Sub FilterData()
'过滤数据
     Dim strFilter As String
     Dim lngCurSpeExamType As Long
     
    If ufgData.AdoData Is Nothing Then Exit Sub
            
    If optMianYi.value Then lngCurSpeExamType = 0
    If optTeShu.value Then lngCurSpeExamType = 1
    If optFenZi.value Then lngCurSpeExamType = 2
            
    '判断当前状态，根据复选框显示数据
    If chkYSQ.value <> 0 Then
        If strFilter <> "" Then strFilter = strFilter & " or "
        strFilter = strFilter & "(当前状态=0 and 特检类型=" & lngCurSpeExamType & ")"
    End If
    
    If chkYJS.value <> 0 Then
         If strFilter <> "" Then strFilter = strFilter & " or "
         strFilter = strFilter & "(当前状态=1 and 特检类型=" & lngCurSpeExamType & ")"
    End If
    
    If chkYWC.value <> 0 Then
         If strFilter <> "" Then strFilter = strFilter & " or "
         strFilter = strFilter & "(当前状态=2 and 特检类型=" & lngCurSpeExamType & ")"
    End If
    
    If strFilter = "" Then
         strFilter = "(当前状态=9 and 特检类型=" & lngCurSpeExamType & ")"
    End If
    
    ufgData.AdoData.Filter = strFilter
    
    '刷新数据
    Call ufgData.RefreshData

    Call RefreshSpeExamInf
    
End Sub

Private Sub chkYSQ_Click()
On Error GoTo errHandle
    '调用过滤数据方法
    Call FilterData
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub chkYJS_Click()
On Error GoTo errHandle

    Call FilterData
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub chkYWC_Click()
On Error GoTo errHandle

    Call FilterData
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub SpeExamined_Accept()
'特检接收
    Dim strSql As String
    Dim i As Long
    Dim blnIsExit As Boolean
    
    blnIsExit = False
    Select Case GetSelectSpeExamType
        Case C_INT_MIANYI
            If mrecStudyInf.lngMianYiStep <> TExecuteStep.NeedDo And mrecStudyInf.lngMianYiStep <> TExecuteStep.AcceptDo Then blnIsExit = True
            
        Case C_INT_FENZI
            If mrecStudyInf.lngFenZiStep <> TExecuteStep.NeedDo And mrecStudyInf.lngFenZiStep <> TExecuteStep.AcceptDo Then blnIsExit = True
            
        Case C_INT_TESHU
            If mrecStudyInf.lngTeRanStep <> TExecuteStep.NeedDo And mrecStudyInf.lngTeRanStep <> TExecuteStep.AcceptDo Then blnIsExit = True
    End Select
    
    '非特检阶段，不能进行接受
    If blnIsExit Then
        
        Call MsgBoxD(Me, "尚未进入特检阶段，不能进行特检确认操作。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    
    If Not CheckAllowAccept Then
        Call MsgBoxD(Me, "没有需要进行接受的特检项目。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    strSql = "Zl_病理特检_接受('" & mrecStudyInf.lngPatholAdviceId & "'," & GetSelectSpeExamType & ",'" & UserInfo.姓名 & "')"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    Select Case GetSelectSpeExamType
        Case C_INT_MIANYI
            mrecStudyInf.lngMianYiStep = TExecuteStep.AcceptDo
        Case C_INT_FENZI
            mrecStudyInf.lngFenZiStep = TExecuteStep.AcceptDo
        Case C_INT_TESHU
            mrecStudyInf.lngTeRanStep = TExecuteStep.AcceptDo
    End Select
    
    For i = 1 To ufgData.GridRows - 1
        If Not ufgData.IsEmptyKey(i) Then
            If ufgData.Text(i, gstrSpeExam_当前状态) = "已申请" Then
                Call ufgData.SyncText(i, gstrSpeExam_当前状态, "已接受", True)
                Call ufgData.SyncText(i, gstrSpeExam_特检医师, UserInfo.姓名, True)
            End If
        End If
    Next i
    
    Call MsgBoxD(Me, "已接受" & Decode(GetSelectSpeExamType, 0, "免疫组化", 1, "特殊染色", 2, "分子病理") & "检查。", vbOKOnly, Me.Caption)
End Sub


Private Function CheckAllowAccept() As Boolean
    Dim i As Long
    
    CheckAllowAccept = False
    
    For i = 1 To ufgData.GridRows - 1
        If Not ufgData.IsEmptyKey(i) Then
            If Trim(ufgData.Text(i, gstrSpeExam_当前状态)) = "已申请" Then
                CheckAllowAccept = True
                Exit Function
            End If
        End If
    Next i
End Function



Private Sub SpeExamined_Sure()
'特检确认
    Dim strSql As String
    Dim i As Long
    Dim dtServicesTime As Date
    Dim blnIsExit As Boolean
    
    blnIsExit = False
    Select Case GetSelectSpeExamType
        Case C_INT_MIANYI
            If mrecStudyInf.lngMianYiStep <> TExecuteStep.NeedDo And mrecStudyInf.lngMianYiStep <> TExecuteStep.AcceptDo Then blnIsExit = True
            
        Case C_INT_FENZI
            If mrecStudyInf.lngFenZiStep <> TExecuteStep.NeedDo And mrecStudyInf.lngFenZiStep <> TExecuteStep.AcceptDo Then blnIsExit = True
            
        Case C_INT_TESHU
            If mrecStudyInf.lngTeRanStep <> TExecuteStep.NeedDo And mrecStudyInf.lngTeRanStep <> TExecuteStep.AcceptDo Then blnIsExit = True
    End Select
    
    '非特检阶段，不能进行确认
    If blnIsExit Then
        
        Call MsgBoxD(Me, "尚未进入特检阶段，不能进行特检确认操作。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    
    If Not CheckAllowSpeExamSure Then
        Call MsgBoxD(Me, "已接受的特检项目结果尚未完全录入，不能进行确认。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
'    '保存项目结果
'    Call SpeExamined_Save

    dtServicesTime = zlDatabase.Currentdate
    
    strSql = "Zl_病理特检_确认('" & mrecStudyInf.lngPatholAdviceId & "'," & GetSelectSpeExamType & "," & To_Date(dtServicesTime) & ")"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    Select Case GetSelectSpeExamType
        Case C_INT_MIANYI
            mrecStudyInf.lngMianYiStep = TExecuteStep.AlreadDo
        Case C_INT_FENZI
            mrecStudyInf.lngFenZiStep = TExecuteStep.AlreadDo
        Case C_INT_TESHU
            mrecStudyInf.lngTeRanStep = TExecuteStep.AlreadDo
    End Select
    
    
    For i = 1 To ufgData.GridRows - 1
        If Not ufgData.IsEmptyKey(i) Then
            If ufgData.Text(i, gstrSpeExam_当前状态) = "已接受" Then
                Call ufgData.SyncText(i, gstrSpeExam_当前状态, "已完成", True)
                Call ufgData.SyncText(i, gstrSpeExam_完成时间, dtServicesTime, True)
            End If
            
            If ufgData.Text(i, gstrSpeExam_当前状态) = "已申请" Then
                Select Case GetSelectSpeExamType
                    Case C_INT_MIANYI
                        mrecStudyInf.lngMianYiStep = TExecuteStep.NeedDo
                    Case C_INT_FENZI
                        mrecStudyInf.lngFenZiStep = TExecuteStep.NeedDo
                    Case C_INT_TESHU
                        mrecStudyInf.lngTeRanStep = TExecuteStep.NeedDo
                End Select
            End If
        End If
    Next i
    
    '触发特检确认事件
    Call SendMsgToMainWindow(Me, wetSpeExamSure, mlngAdviceID)
    
    Call MsgBoxD(Me, "已完成对" & Decode(GetSelectSpeExamType, 0, "免疫组化", 1, "特殊染色", 2, "分子病理") & "检查的确认。", vbOKOnly, Me.Caption)
End Sub


Public Function CheckAllowSpeExamSure() As Boolean
'是否允许特检确认(如果存在已申请或者项目结果为空的项目则不能进行确认)
    Dim i As Long
    
    CheckAllowSpeExamSure = True
    
'    For i = 1 To ufgData.GridRows - 1
'        If Not ufgData.KeyEmpty(i) Then
'            If Trim(ufgData.Text(i, gstrSpeExam_项目结果)) = "" And ufgData.Text(i, gstrSpeExam_当前状态) = "已接受" Then
'                CheckAllowSpeExamSure = False
'                Exit Function
'            End If
'        End If
'    Next i
End Function






Private Sub PrintSpeExamLabel(Optional ByVal blnIsPrint As Boolean = True)
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
    
    Call zlReport.ReportOpen(gcnOracle, 100, "ZL1_Inside_1294_11", Me, "项目ID1=" & strValue(0), "项目ID2=" & strValue(1), "项目ID3=" & strValue(2), "项目ID4=" & strValue(3), "项目ID5=" & strValue(4), "项目ID6=" & strValue(5), IIf(blnIsPrint, 2, 1))
End Sub


Private Sub PrintSelectSpeExamLabel(Optional ByVal blnIsPrint As Boolean = True)
'打印选择的材块标签
On Error GoTo errHandle
    Dim strValue(5) As String
    
    
    If Not ufgData.IsSelectionRow Then
        Call MsgBoxD(Me, "请选择需要打印的特检记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgData.IsNullRow(ufgData.SelectionRow) Then
        Call MsgBoxD(Me, "请选择需要打印的特检记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    strValue(0) = ufgData.KeyValue(ufgData.SelectionRow)
    strValue(1) = "0": strValue(2) = "0": strValue(3) = "0": strValue(4) = "0": strValue(5) = "0"

    
    Call zlReport.ReportOpen(gcnOracle, 100, "ZL1_Inside_1294_11", Me, "项目ID1=" & strValue(0), "项目ID2=" & strValue(1), "项目ID3=" & strValue(2), "项目ID4=" & strValue(3), "项目ID5=" & strValue(4), "项目ID6=" & strValue(5), IIf(blnIsPrint, 2, 1))
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub SpeExamined_Save()
'保存特检项目
    Dim i As Long
    Dim strSql As String
    
    For i = 1 To ufgData.GridRows - 1
        If Not ufgData.IsEmptyKey(i) Then
            If ufgData.Text(i, gstrSpeExam_当前状态) = "已接受" Then
                strSql = "Zl_病理特检_项目录入(" & ufgData.KeyValue(i) & ",null)"
        
                Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
            End If
        End If
    Next i
    
End Sub

Private Sub cmdSave_Click()
'保存项目结果
On Error GoTo errHandle
    If Not CheckAllowSpeExamSure Then
        Call MsgBoxD(Me, "项目结果尚未完全录入，不能进行保存。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    Call SpeExamined_Save
    
    Call MsgBoxD(Me, "项目结果已保存。", vbOKOnly, Me.Caption)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub PrintWorkList(Optional ByVal blnIsPrint As Boolean = True)
'打印特检工作列表
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
    
    '对于清单的打印，使用带报表预览的方式
    Call zlReport.ReportOpen(gcnOracle, 100, "ZL1_Inside_1294_10", Me, "项目ID=" & strValue(0), "项目ID1=" & strValue(1), "项目ID2=" & strValue(2), "项目ID3=" & strValue(3), "项目ID4=" & strValue(4), "项目ID5=" & strValue(5), IIf(blnIsPrint, 2, 1))
    
End Sub


Private Sub ShowSpeExamRequest()
'显示特检申请
Dim frmRequestView As New frmPatholRequisition_View
On Error GoTo errFree
    Call frmRequestView.ShowRequestViewWind(mrecStudyInf.lngPatholAdviceId, GetSelectSpeExamType, mblnMoved, Me)
errFree:
    Call Unload(frmRequestView)
    Set frmRequestView = Nothing
End Sub

Private Sub Form_Initialize()
    mKeyCode = -1
    mKeyShift = -1
    
    Set zlReport = New zl9Report.clsReport
    mblnAutoAcceptOfAfterPrint = False
End Sub

Private Sub Form_Load()
On Error GoTo errHandle
    '初始化特检列表
    Call InitSpeExamList
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub



Private Sub Form_Resize()
On Error Resume Next
    Call AdjustFace
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Set zlReport = Nothing
End Sub

'Private Sub vfgData_KeyDown(KeyCode As Integer, Shift As Integer)
'    mKeyCode = KeyCode
'    mKeyShift = Shift
'End Sub
'
'Private Sub vfgData_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
'    mKeyCode = KeyCode
'    mKeyShift = Shift
'End Sub

Private Function GetRomanAscii(ByVal lngNum As Long) As Integer
'将数字转换成罗马数字的ascii
    GetRomanAscii = Decode(lngNum, 49, -23823, 50, -23822, 51, 23821, 52, -23820, 53, -23819, 54, -23818, 55, -23817, 56, -23816, 57, -23815)
End Function


'
'Private Sub vfgData_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
'    '将小键盘的*替换为%符合
'    If KeyAscii = 42 And mKeyShift <> 1 Then KeyAscii = 37
'
''    If mKeyShift = 2 Then
''        If KeyAscii >= 49 And KeyAscii <= 57 Then KeyAscii = GetRomanAscii(KeyAscii)
''    End If
'End Sub
'
'Private Sub vfgData_KeyUp(KeyCode As Integer, Shift As Integer)
'    mKeyCode = -1
'    mKeyShift = -1
'End Sub

'Private Sub vfgData_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
'    mKeyCode = -1
'    mKeyShift = -1
'End Sub

'Private Sub vfgData_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
''除项目结果外，其他列不允许编辑
'    If Col <> mvfgSpeExam.GetColumnIndex(gstrSpeExam_项目结果) Then Cancel = True
'End Sub

Private Sub UpdateWorkListPrintState()
'在打印后，更新工作清单的打印状态
    Dim strSql As String
    Dim i As Long
        
    For i = 1 To ufgData.GridRows - 1
        If Not ufgData.IsEmptyKey(i) Then
            strSql = "Zl_病理特检_清单打印(" & ufgData.KeyValue(i) & ")"
            Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)

            Call ufgData.SyncText(i, gstrSpeExamWork_清单状态, "已打印", True)
        End If
    Next i
End Sub



Private Sub OptFenZi_Click()
'过滤指定特检类型的数据
On Error GoTo errHandle
     Call FilterData
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub optMianYi_Click()
'过滤指定特检类型的数据
On Error GoTo errHandle
    Call FilterData
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub OptTeShu_Click()
'过滤指定特检类型的数据
On Error GoTo errHandle
    Call FilterData
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ShowAntibodyInf(ByVal lngAntibodyRow As Long)
'显示抗体明细信息
    Dim frmAntibodyInf As New frmPatholRequisition_AntibodyInf
    On Error GoTo errFree
        Call frmAntibodyInf.ShowAntibodyInf(ufgData.Text(lngAntibodyRow, gstrSpeExam_抗体ID), Me)
errFree:
    Call Unload(frmAntibodyInf)
    Set frmAntibodyInf = Nothing
    
End Sub



Private Sub ufgData_OnCellButtonClick(ByVal Row As Long, ByVal Col As Long)
On Error GoTo errHandle
    Call ShowAntibodyInf(Row)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub



Private Sub ufgData_OnColsNameReSet()
On Error GoTo errHandle

    Call QuerySpeExamData(mrecStudyInf.lngPatholAdviceId)
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgData_OnStartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'只有特检项目被接收后，才能进行编辑
    Dim strState As String
    
    strState = ufgData.Text(Row, gstrSpeExam_当前状态)
    
    If Col = ufgData.GetColIndex(gstrSpeExam_项目结果) Then
        Cancel = IIf(Trim(strState) = "" Or strState = "已申请", True, False)
        
        If Cancel Then
            Call MsgBoxD(Me, "特检项目未被接受，不能进行录入。", vbOKOnly, Me.Caption)
        End If
    End If
End Sub

Private Sub zlReport_AfterPrint(ByVal ReportNum As String)
On Error GoTo errHandle
    '如果不是特检清单打印，则直接退出
    If ReportNum <> "ZL1_PATHOLSPEEXAM_01" Then Exit Sub
    
    Call UpdateWorkListPrintState
    
    If mblnAutoAcceptOfAfterPrint Then
    '打印后自动接受
        Call SpeExamined_Accept
    End If
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

