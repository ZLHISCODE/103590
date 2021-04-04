VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmPatholSlices 
   Caption         =   "病理制片"
   ClientHeight    =   8955
   ClientLeft      =   75
   ClientTop       =   405
   ClientWidth     =   10665
   Icon            =   "frmPatholSlices.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8955
   ScaleWidth      =   10665
   StartUpPosition =   3  '窗口缺省
   Begin MSComctlLib.ImageList imgTbrS 
      Left            =   8415
      Top             =   765
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
            Picture         =   "frmPatholSlices.frx":000C
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholSlices.frx":0C7E
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholSlices.frx":18F0
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholSlices.frx":2562
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholSlices.frx":31D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholSlices.frx":3E46
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholSlices.frx":4AB8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame framSlices 
      Caption         =   "制片记录"
      Height          =   7215
      Left            =   240
      TabIndex        =   1
      Top             =   795
      Width           =   9975
      Begin VB.Frame FramCheck 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   6120
         TabIndex        =   3
         Top             =   6840
         Width           =   3735
         Begin VB.CheckBox chkYWC 
            Caption         =   "已完成"
            Height          =   180
            Left            =   2760
            TabIndex        =   6
            Top             =   30
            Width           =   855
         End
         Begin VB.CheckBox chkYJS 
            Caption         =   "已接受"
            Height          =   180
            Left            =   1800
            TabIndex        =   5
            Top             =   30
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chkWCL 
            Caption         =   "未处理"
            Height          =   255
            Left            =   720
            TabIndex        =   4
            Top             =   0
            Value           =   1  'Checked
            Width           =   855
         End
      End
      Begin zl9PACSWork.ucFlexGrid ufgData 
         Height          =   6255
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   11033
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
         ExtendLastCol   =   -1  'True
      End
      Begin VB.Label labRecordInf 
         Caption         =   "当前总制片数：0    当前需制片数：0"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   6840
         Width           =   4695
      End
   End
   Begin MSComctlLib.Toolbar tbrMain 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   10665
      _ExtentX        =   18812
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
            Caption         =   "制片接受"
            Key             =   "tbAcceptSlices"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "制片完成"
            Key             =   "tbEndSlices"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPatholSlices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IWorkMenu

Private Const M_STR_HINT_NoSelectData As String = "无效的检查数据，请选择需要执行的检查记录。"
Private Const M_STR_MODULE_MENU_TAG As String = "制片"


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

Private mrecStudy As TStudyStateInf
Private mblnReadOnly As Boolean

Private mlngTmpAdviceId As Long
Private mlngTmpSendNo As Long

Private mblnAutoAcceptOfAfterPrint As Boolean
Private mbytFontSize As Byte '字号    9--小字体    12--大字体


Private mObjActiveMenuBar As CommandBar

Private mblnRefreshState As Boolean


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

    If Not HasMenu(objMenuBar, conMenu_PatholSlices) Then
        Set cbrMenuBar = mObjActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_PatholSlices, "制片(&L)", GetPatholMenuIndex(objMenuBar) + 1, False)
        cbrMenuBar.ID = conMenu_PatholSlices
        cbrMenuBar.Category = M_STR_MODULE_MENU_TAG
                
        With cbrMenuBar.CommandBar
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlPopup, conMenu_PatholSlices_LAB, "标签打印(&B)", "", 1, False)
                Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_PatholSlices_PreviewLAB, "预览(V)", "", 1, False)
                Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_PatholSlices_PrintLAB, "打印(P)", "", 1, False)
                
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlPopup, conMenu_PatholSlices_List, "清单打印(&T)", "", 1, False)
                Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_PatholSlices_PreviewList, "预览(V)", "", 1, False)
                Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_PatholSlices_PrintList, "打印(P)", "", 1, False)
            
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholSlices_RequestView, "申请查看(&Q)", "", 1, True)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholSlices_Accept, "制片接受(&R)", "", 1, False)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholSlices_Finish, "制片完成(&F)", "", 1, False)
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
        Case conMenu_PatholSlices_PreviewLAB
            Call PrintSlicesLabel(False)
            
        Case conMenu_PatholSlices_PrintLAB
            Call PrintSlicesLabel(True)
            
        Case conMenu_PatholSlices_PreviewList
            Call PrintWorkList(False)
            
        Case conMenu_PatholSlices_PrintList
            Call PrintWorkList(True)
            
        Case conMenu_PatholSlices_RequestView
            Call ShowSlicesRequest
            
        Case conMenu_PatholSlices_Accept
            Call Slices_Accept
            
        Case conMenu_PatholSlices_Finish
            Call Slices_Sure
    End Select
End Sub


Public Sub IWorkMenu_zlUpdateMenu(ByVal control As XtremeCommandBars.ICommandBarControl)
'更新菜单
    Dim blnIsAllowSlices As Boolean
    
    If Not Me.Visible Then
        control.Enabled = False
        Exit Sub
    End If
    
    blnIsAllowSlices = CheckPopedom(mstrPrivs, "病理制片") And mlngAdviceID > 0
    
    Select Case control.ID
        Case conMenu_PatholSlices_LAB
            control.Enabled = blnIsAllowSlices And mrecStudy.strPatholNumber <> ""
            
        Case conMenu_PatholSlices_List
            control.Enabled = blnIsAllowSlices And mrecStudy.strPatholNumber <> ""
            
        Case conMenu_PatholSlices_RequestView
            control.Enabled = blnIsAllowSlices And mrecStudy.strPatholNumber <> ""
            
        Case conMenu_PatholSlices_Accept
            control.Enabled = blnIsAllowSlices And Not mblnReadOnly
            
        Case conMenu_PatholSlices_Finish
            control.Enabled = blnIsAllowSlices And Not mblnReadOnly
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
    If (mlngTmpAdviceId = mlngAdviceID And mlngTmpSendNo = mlngSendNo And mblnRefreshState) And Not blnForceRefresh Then Exit Sub
    
    mlngTmpAdviceId = mlngAdviceID
    mlngTmpSendNo = mlngSendNo
    mblnRefreshState = True
        
    If mlngAdviceID <= 0 Then
        Call ConfigSlicesFace(False, "医嘱ID无效请检查。")
        Exit Sub
    End If
    
    Call GetPatholStudyState(mlngAdviceID, mrecStudy)
    
    If Trim(mrecStudy.strPatholNumber) = "" Then
        Call ConfigSlicesFace(False, "该检查尚未生成有效的病理号，请确认该检查是否已被核收。")
        
        Exit Sub
    Else
        Call ConfigSlicesFace(True)
    End If

    
    '读取制片数据
    Call LoadSlicesData
    
    '刷新材块数量
    Call RefreshSilcesCount
    
    Call ConfigPopedom(mblnReadOnly)
    
'    If Not (owner Is Nothing) Then
'        Call Me.Show(1, owner)
'    End If
End Sub


Public Sub zlRefresh(ByVal lngAdviceID As Long, ByVal blnReadOnly As Boolean, _
    ByVal strPrivs As String, ByVal blnMoved As Boolean, _
    ByVal lngCurDepartmentId As Long, Optional owner As Form = Nothing)
'刷新取材模块
'如果同时有取材功能，则添加取材记录后，制片需要刷新
'    If lngAdviceID = mlngCurAdviceId Then  Exit Sub
        
    If lngAdviceID <= 0 Then
        Call ConfigSlicesFace(False, "医嘱ID无效请检查。")
        Exit Sub
    End If
    
    mlngAdviceID = lngAdviceID              '医嘱ID
    mstrPrivs = strPrivs                    '执行权限
    mblnMoved = blnMoved                    '是否转储
    mlngCurDeptId = lngCurDepartmentId      '部门编号
    
   

    Call GetPatholStudyState(mlngAdviceID, mrecStudy)
    
    If Trim(mrecStudy.strPatholNumber) = "" Then
        Call ConfigSlicesFace(False, "该检查尚未生成有效的病理号，请确认该检查是否已被核收。")
        
        Exit Sub
    Else
        Call ConfigSlicesFace(True)
    End If

    
    '读取制片数据
    Call LoadSlicesData
    
    '刷新材块数量
    Call RefreshSilcesCount
    
    Call ConfigPopedom(blnReadOnly)
    
    If Not (owner Is Nothing) Then
        Call Me.Show(1, owner)
    End If
End Sub


Private Sub RefreshSilcesCount()
'刷新制片记录数量
    Dim i As Long
    Dim lngRecord As Long
    Dim lngTotal As Long
    Dim lngSlices As Long
    
    lngTotal = 0
    lngSlices = 0
    
    
    For i = 1 To ufgData.GridRows - 1
        If Not ufgData.IsNullRow(i) Then
            
            lngTotal = lngTotal + Val(ufgData.Text(i, gstrSlices_制片数))
            
            If ufgData.Text(i, gstrSlices_当前状态) <> "已完成" Then
                lngSlices = lngSlices + Val(ufgData.Text(i, gstrSlices_制片数))
            End If
        End If
    Next i
    
    labRecordInf.Caption = "当前总制片数：" & lngTotal & "    当前需制片数：" & lngSlices
    
End Sub

Private Sub ConfigPopedom(ByVal blnIsReadOnly As Boolean)
'配置权限
    Dim blnIsAllowSlices As Boolean
    
    blnIsAllowSlices = CheckPopedom(mstrPrivs, "病理制片")
    
    tbrMain.Buttons("tbAcceptSlices").Enabled = blnIsAllowSlices And Not blnIsReadOnly
    tbrMain.Buttons("tbEndSlices").Enabled = blnIsAllowSlices And Not blnIsReadOnly

    
    tbrMain.Buttons("tbLAB").Enabled = blnIsAllowSlices
    tbrMain.Buttons("tbList").Enabled = blnIsAllowSlices
    tbrMain.Buttons("tbViewRequest").Enabled = blnIsAllowSlices
    
    ufgData.ReadOnly = blnIsReadOnly
End Sub


Private Sub AdjustFace()
    '调整界面布局
    framSlices.Left = 0
    framSlices.Top = tbrMain.Top + tbrMain.Height + 120
    framSlices.Width = Me.Width - 0
    framSlices.Height = Me.Height - tbrMain.Height - 240
    
    ufgData.Left = 120
    ufgData.Top = 240 + IIf(mbytFontSize = 9, 0, 120)
    ufgData.Width = framSlices.Width - 240
    ufgData.Height = framSlices.Height - labRecordInf.Height - 480
    
    labRecordInf.Left = 120
    labRecordInf.Top = framSlices.Height - labRecordInf.Height - 120 + IIf(mbytFontSize = 9, 0, 85)

    
    '调整FrameCheck位置
     FramCheck.Top = framSlices.Height - labRecordInf.Height - 120 + IIf(mbytFontSize = 9, 0, 70)
     FramCheck.Left = framSlices.Width - FramCheck.Width - 200
     
     chkWCL.Top = 0
     chkYJS.Top = 0
     chkYWC.Top = 0
End Sub

Public Sub ReSetFormFontSize(ByVal bytFontSize As Byte)
'功能:重新设置工作站窗体的字体大小
On Error GoTo errHandle

    Dim objCtrl As control
    Dim CtlFont As StdFont
    Dim strFontType  As String
    
    '先移动控件位置
    mbytFontSize = bytFontSize
    
    '再设置字体
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



Private Sub LoadSlicesData()
'读取制片信息
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    strSql = "select a.ID,a.材块ID,b.序号,b.取材位置, b.标本名称,a.制片数,a.制片类型, a.制片方式,a.制片时间,a.制片人 as 制片技师,a.当前状态,a.清单状态" & _
            " from 病理制片信息 a, 病理取材信息 b " & _
            " where a.材块id=b.材块id and b.确认状态=1 and b.病理医嘱ID = [1] order by a.当前状态,b.序号,a.ID"
'    If mblnMoved Then strSql = GetMovedDataSql(strSql)
    
    Set ufgData.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mrecStudy.lngPatholAdviceId)
    
    Call FilterData

End Sub

Private Sub FilterData()
'过滤数据
     Dim strFilter As String
    
    '判断当前状态，根据复选框显示数据
    If chkWCL.value <> 0 Then
        If strFilter = "" Then
            strFilter = "当前状态=0"
        Else
             strFilter = strFilter & " or " & "当前状态=0"
        End If
        
    End If
    
    If chkYJS.value <> 0 Then
        If strFilter = "" Then
            strFilter = "当前状态=1"
        Else
             strFilter = strFilter & " or " & "当前状态=1"
        End If
    End If
    
    If chkYWC.value <> 0 Then
        If strFilter = "" Then
            strFilter = "当前状态=2"
        Else
             strFilter = strFilter & " or " & "当前状态=2"
        End If
    End If
    
     If strFilter = "" Then
            strFilter = "当前状态=9"
    End If
    
    ufgData.AdoData.Filter = strFilter
    '刷新数据
    Call ufgData.RefreshData

    Call RefreshSilcesCount
End Sub

Private Sub chkWCL_Click()
On Error Resume Next
    Call FilterData

End Sub

Private Sub chkYJS_Click()
On Error Resume Next
    Call FilterData

End Sub

Private Sub chkYWC_Click()
On Error Resume Next
    Call FilterData

End Sub



Private Sub InitSlicesList()
'初始化制片列表
    Dim strTemp As String
    
    ufgData.IsKeepRows = True
    ufgData.GridRows = glngMaxRowCount
    '设置行高
    ufgData.RowHeightMin = glngStandardRowHeight
    
     '判断数据库参数表是否有数据 有则读取数据库参数  没有则加载默认
    strTemp = zlDatabase.GetPara("病理制片列表配置", glngSys, G_LNG_PATHOLSYS_NUM, "")
     
    If strTemp = "" Then
        ufgData.ColNames = gstrSlicesCols
    Else
        ufgData.ColNames = strTemp
    End If
    
    ufgData.DefaultColNames = gstrSlicesCols
    ufgData.ColConvertFormat = gstrSlicesConvertFormat
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
        Case UCase("tbLAB"), UCase("tbLABPreview")
            '预览标签
            Call PrintSlicesLabel(False)
            
        Case UCase("tbLABPrint")
            '打印标签
            Call PrintSlicesLabel(True)
        
        Case UCase("tbList"), UCase("tbListPreview")
            '预览清单
            Call PrintWorkList(False)
            
        Case UCase("tbListPrint")
            '打印清单
            Call PrintWorkList(True)
            
        Case UCase("tbAcceptSlices")
            '制片接收
            Call Slices_Accept
            
        Case UCase("tbEndSlices")
            '制片完成
            Call Slices_Sure
            
        Case UCase("tbViewRequest")
            '查看申请单
            ShowSlicesRequest
    End Select
End Sub

Private Sub ufgData_OnColFormartChange()
'关闭窗口时保存列表配置
    zlDatabase.SetPara "病理制片列表配置", ufgData.GetColsString(ufgData), glngSys, G_LNG_PATHOLSYS_NUM
End Sub




Private Sub ConfigSlicesFace(ByVal blnIsValid As Boolean, Optional ByVal strHintInf As String = "")
'配置特检界面

    tbrMain.Buttons("tbAcceptSlices").Enabled = blnIsValid
    tbrMain.Buttons("tbEndSlices").Enabled = blnIsValid
    tbrMain.Buttons("tbLAB").Enabled = blnIsValid
    tbrMain.Buttons("tbList").Enabled = blnIsValid
    tbrMain.Buttons("tbViewRequest").Enabled = blnIsValid
    
    
    chkWCL.Enabled = blnIsValid
    chkYJS.Enabled = blnIsValid
    chkYWC.Enabled = blnIsValid
    
    If blnIsValid Then
        Call ufgData.CloseHintInf
    Else
        Call ufgData.ShowHintInf(strHintInf)
        labRecordInf.Caption = ""
    End If
End Sub


Private Sub Slices_Accept()
'制片接收
    Dim strSql As String
    Dim i As Long
    
    '非制片阶段，不能进行接受
    If mrecStudy.lngSlicesStep <> TExecuteStep.NeedDo And mrecStudy.lngSlicesStep <> TExecuteStep.AcceptDo Then
        Call MsgBoxD(Me, "尚未进入制片阶段，不能进行制片接受操作。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
       
    
    If Not HasNeedState("未处理") Then
        Call MsgBoxD(Me, "没有需要进行接受的制片信息，请确认是否存在未处理的制片信息。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    strSql = "Zl_病理制片_接受('" & mrecStudy.lngPatholAdviceId & "','" & UserInfo.姓名 & "')"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    mrecStudy.lngSlicesStep = 2
    
    For i = 1 To ufgData.GridRows - 1
        If Not ufgData.IsEmptyKey(i) Then
            If ufgData.Text(i, gstrSlices_当前状态) = "未处理" Then
                ufgData.Text(i, gstrSlices_当前状态) = "已接受"
                ufgData.Text(i, gstrSlices_制片人) = UserInfo.姓名
            End If
        End If
    Next i
    
    Call MsgBoxD(Me, "已接受制片。", vbOKOnly, Me.Caption)
End Sub


Private Function HasNeedState(ByVal strState As String) As Boolean
'判断是否需要进行核收
    Dim i As Long
    
    HasNeedState = False
    For i = 1 To ufgData.GridRows - 1
        If Not ufgData.IsEmptyKey(i) Then
            If ufgData.Text(i, gstrSlices_当前状态) = strState Then
                HasNeedState = True
                Exit Function
            End If
        End If
    Next i
End Function


Private Sub Slices_Sure()
'制片确认
    Dim strSql As String
    Dim i As Long
    Dim j As Long
    Dim lngSlicesCount As Long
    Dim strTemp As String
    Dim rsData As ADODB.Recordset
    Dim dtServicesTime As Date
    
    '非制片阶段，不能进行确认
    If mrecStudy.lngSlicesStep <> TExecuteStep.NeedDo And mrecStudy.lngSlicesStep <> TExecuteStep.AcceptDo Then
        Call MsgBoxD(Me, "尚未进入制片阶段，不能进行制片确认操作。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If Not HasNeedState("已接受") Then
        Call MsgBoxD(Me, "没有需要进行确认的制片信息，请确认是否存在已被接受的制片信息。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    dtServicesTime = zlDatabase.Currentdate
    
    strSql = "Zl_病理制片_确认('" & mrecStudy.lngPatholAdviceId & "'," & To_Date(dtServicesTime) & ")"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    mrecStudy.lngSlicesStep = 3
    
    For i = 1 To ufgData.GridRows - 1
    
        If Not ufgData.IsEmptyKey(i) Then
            If ufgData.Text(i, gstrSlices_当前状态) = "已接受" Then
                ufgData.Text(i, gstrSlices_当前状态) = "已完成"
                ufgData.Text(i, gstrSlices_制片时间) = dtServicesTime
            End If
            
            If ufgData.Text(i, gstrSlices_当前状态) = "未处理" Then
                mrecStudy.lngSlicesStep = 1
            End If
        End If
        
    Next i
    
    '触发制片确认事件
    Call SendMsgToMainWindow(Me, wetSlicesSure, mlngAdviceID)
    
    Call MsgBoxD(Me, "已确认制片。", vbOKOnly, Me.Caption)
End Sub





Private Sub PrintSlicesLabel(Optional ByVal blnIsPrint As Boolean = True)
'打印特检项目标签
    Dim i As Long
    Dim j As Long
    Dim strValue(5) As String
    
    Dim strSliceId As String
    Dim k As Long
    Dim lngCount As Long
    
    j = 0
    strValue(0) = "0": strValue(1) = "0": strValue(2) = "0": strValue(3) = "0": strValue(4) = "0": strValue(5) = "0"
    For i = 1 To ufgData.GridRows - 1
        If Not ufgData.IsEmptyKey(i) Then
            If zlCommFun.ActualLen(strValue(j)) > 2000 Then
                j = j + 1
                strValue(j) = ""
            End If
            
            strSliceId = ufgData.KeyValue(i)
            lngCount = Val(ufgData.Text(i, gstrSlices_制片数))
    
            If strValue(j) <> "" Then strValue(j) = strValue(j) & ","
    
            strValue(j) = strValue(j) & strSliceId
            
            If lngCount > 1 Then
                For k = 1 To lngCount - 1
                    strValue(j) = strValue(j) & "," & strSliceId
                Next k
            End If
            
        End If
    Next i
    
    Call zlReport.ReportOpen(gcnOracle, 100, "ZL1_Inside_1294_09", Me, "制片ID1=" & strValue(0), "制片ID2=" & strValue(1), "制片ID3=" & strValue(2), "制片ID4=" & strValue(3), "制片ID5=" & strValue(4), "制片ID6=" & strValue(5), IIf(blnIsPrint, 2, 1))
End Sub



Private Sub PrintSelectSlicesLabel(Optional ByVal blnIsPrint As Boolean = True)
'打印选择的材块标签
On Error GoTo errHandle
    Dim strValue(5) As String
    Dim strSliceId As String
    Dim i As Long
    Dim lngCount As Long
    
    
    If Not ufgData.IsSelectionRow Then
        Call MsgBoxD(Me, "请选择需要打印的制片记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgData.IsNullRow(ufgData.SelectionRow) Then
        Call MsgBoxD(Me, "请选择需要打印的制片记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    strSliceId = ufgData.KeyValue(ufgData.SelectionRow)
    lngCount = Val(ufgData.Text(ufgData.SelectionRow, gstrSlices_制片数))
    
    strValue(0) = strSliceId
    If lngCount > 1 Then
    '当制片数大于1时，则传递相同数量的ID
        For i = 1 To lngCount - 1
            strValue(0) = strValue(0) & "," & strSliceId
        Next i
    End If
    
    strValue(1) = "0": strValue(2) = "0": strValue(3) = "0": strValue(4) = "0": strValue(5) = "0"

    
    Call zlReport.ReportOpen(gcnOracle, 100, "ZL1_Inside_1294_09", Me, "制片ID1=" & strValue(0), "制片ID2=" & strValue(1), "制片ID3=" & strValue(2), "制片ID4=" & strValue(3), "制片ID5=" & strValue(4), "制片ID6=" & strValue(5), IIf(blnIsPrint, 2, 1))
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub PrintWorkList(Optional ByVal blnIsPrint As Boolean = True)
'打印制片工作列表
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
    Call zlReport.ReportOpen(gcnOracle, 100, "ZL1_Inside_1294_08", Me, "制片ID1=" & strValue(0), "制片ID2=" & strValue(1), "制片ID3=" & strValue(2), "制片ID4=" & strValue(3), "制片ID5=" & strValue(4), "制片ID6=" & strValue(5), IIf(blnIsPrint, 2, 1))
    
End Sub


Private Sub ShowSlicesRequest()
'显示制片申请
Dim frmRequestView As New frmPatholRequisition_View
On Error GoTo errFree
    Call frmRequestView.ShowRequestViewWind(mrecStudy.lngPatholAdviceId, 3, mblnMoved, Me)
errFree:
    Call Unload(frmRequestView)
    Set frmRequestView = Nothing
End Sub

Private Sub Form_Initialize()
    Set zlReport = New zl9Report.clsReport
    mblnAutoAcceptOfAfterPrint = False
End Sub

Private Sub Form_Load()
On Error GoTo errHandle
    '初始化制片显示列表
    Call InitSlicesList

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Call AdjustFace
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    
    Set zlReport = Nothing
End Sub


Private Sub UpdateWorkListPrintState()
'在打印后，更新工作清单的打印状态
    Dim strSql As String
    Dim i As Long
        
    For i = 1 To ufgData.GridRows - 1
        If Not ufgData.IsEmptyKey(i) Then
            strSql = "Zl_病理制片_清单打印(" & ufgData.KeyValue(i) & ")"
            Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)

            Call ufgData.SyncText(i, gstrSlices_清单状态, "已打印", True)
        End If
    Next i
End Sub


Private Sub ufgData_OnColsNameReSet()
On Error GoTo errHandle

    Call LoadSlicesData
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub zlReport_AfterPrint(ByVal ReportNum As String)
On Error GoTo errHandle
    '如果不是制片清单打印，则直接退出
    If ReportNum <> "ZL1_PATHOLSLICES_01" Then Exit Sub
    
    Call UpdateWorkListPrintState
    
    If mblnAutoAcceptOfAfterPrint Then
    '打印后自动接受
        Call Slices_Accept
    End If
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

