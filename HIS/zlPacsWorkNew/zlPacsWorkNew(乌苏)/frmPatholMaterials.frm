VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{84865D89-6B2D-42E2-98C7-18F4206945F5}#2.0#0"; "zl9PacsControl.ocx"
Begin VB.Form frmPatholMaterials 
   Caption         =   "取材登记"
   ClientHeight    =   9405
   ClientLeft      =   75
   ClientTop       =   405
   ClientWidth     =   10560
   Icon            =   "frmPatholMaterials.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9405
   ScaleWidth      =   10560
   StartUpPosition =   3  '窗口缺省
   Begin zl9PacsControl.ucSplitter ucSplitter1 
      Height          =   135
      Left            =   0
      TabIndex        =   11
      Top             =   4320
      Width           =   10560
      _ExtentX        =   18627
      _ExtentY        =   238
      MousePointer    =   7
      SplitType       =   0
      SplitLevel      =   3
      StartDistance   =   840
      Control1Name    =   "picMaterial"
      Control2Name    =   "picPane1"
   End
   Begin VB.PictureBox picMaterial 
      BorderStyle     =   0  'None
      Height          =   3480
      Left            =   0
      ScaleHeight     =   3480
      ScaleWidth      =   10560
      TabIndex        =   12
      Top             =   840
      Width           =   10560
      Begin VB.Frame framMaterial 
         Caption         =   "取材记录"
         Height          =   3495
         Left            =   105
         TabIndex        =   13
         Top             =   15
         Width           =   10080
         Begin VB.CommandButton cmdAutoInputMaterials 
            Caption         =   "录入取材信息(&W)"
            Height          =   400
            Left            =   9600
            TabIndex        =   20
            ToolTipText     =   "将取材记录信息录入到巨检描述中"
            Top             =   2880
            Width           =   1575
         End
         Begin VB.TextBox txtPos 
            Height          =   375
            Left            =   3000
            TabIndex        =   15
            ToolTipText     =   "在这里输入取材后剩余标本所存放的位置。"
            Top             =   2880
            Width           =   2535
         End
         Begin VB.ComboBox cbxSpecimenProcess 
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
            ItemData        =   "frmPatholMaterials.frx":000C
            Left            =   7320
            List            =   "frmPatholMaterials.frx":0022
            TabIndex        =   14
            Text            =   "常规保留"
            Top             =   2880
            Width           =   2295
         End
         Begin zl9PACSWork.ucFlexGrid ufgData 
            Height          =   2535
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   9735
            _ExtentX        =   17171
            _ExtentY        =   4471
            DefaultCols     =   ""
            GridRows        =   51
            IsCopyAdoMode   =   0   'False
            IsEjectConfig   =   -1  'True
            HeadFontCharset =   134
            HeadFontWeight  =   400
            DataFontCharset =   134
            DataFontWeight  =   400
            ExtendLastCol   =   -1  'True
         End
         Begin VB.Label labInf 
            Caption         =   "剩余存放位置："
            Height          =   255
            Left            =   1680
            TabIndex        =   19
            Top             =   3000
            Width           =   1695
         End
         Begin VB.Label labRecordInf 
            Caption         =   "材块总数：0"
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   120
            TabIndex        =   18
            Top             =   3000
            Width           =   1335
         End
         Begin VB.Label labSpecimenProcess 
            Caption         =   "标本后续处理："
            Height          =   255
            Left            =   6000
            TabIndex        =   17
            Top             =   3000
            Width           =   1815
         End
      End
   End
   Begin MSComctlLib.ImageList imgTbrS 
      Left            =   9885
      Top             =   540
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholMaterials.frx":005E
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholMaterials.frx":0CD0
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholMaterials.frx":1942
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholMaterials.frx":25B4
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholMaterials.frx":3226
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholMaterials.frx":3E98
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholMaterials.frx":4B0A
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholMaterials.frx":577C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picPane1 
      BorderStyle     =   0  'None
      Height          =   4950
      Left            =   0
      ScaleHeight     =   4950
      ScaleWidth      =   10560
      TabIndex        =   6
      Top             =   4455
      Width           =   10560
      Begin VB.Frame framWordEdit 
         Height          =   3135
         Left            =   3255
         TabIndex        =   9
         Top             =   -450
         Width           =   9855
         Begin zl9PACSWork.WordInputText wtDescription 
            Height          =   2895
            Left            =   600
            TabIndex        =   0
            Top             =   0
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   5106
            DepartId        =   0
         End
      End
      Begin TabDlg.SSTab tsFilter 
         Height          =   330
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   9885
         _ExtentX        =   17436
         _ExtentY        =   582
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   4
         TabHeight       =   520
         TabMaxWidth     =   2822
         WordWrap        =   0   'False
         TabCaption(0)   =   "巨检描述(&D)"
         TabPicture(0)   =   "frmPatholMaterials.frx":63EE
         Tab(0).ControlEnabled=   0   'False
         Tab(0).ControlCount=   0
         TabCaption(1)   =   "脱钙管理(&C)"
         TabPicture(1)   =   "frmPatholMaterials.frx":640A
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).ControlCount=   0
      End
      Begin VB.Frame framDecalin 
         Height          =   3135
         Left            =   15
         TabIndex        =   8
         Top             =   480
         Width           =   9855
         Begin VB.CommandButton cmdChange 
            Caption         =   "换 缸(&G)"
            Height          =   400
            Left            =   5880
            TabIndex        =   3
            Top             =   2520
            Width           =   1215
         End
         Begin VB.CommandButton cmdDecalin 
            Caption         =   "脱 钙(&T)"
            Height          =   400
            Left            =   4560
            TabIndex        =   2
            Top             =   2520
            Width           =   1215
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "撤 销(&R)"
            Height          =   400
            Left            =   7200
            TabIndex        =   4
            Top             =   2520
            Width           =   1215
         End
         Begin VB.CommandButton cmdSucceed 
            Caption         =   "完 成(&O)"
            Height          =   400
            Left            =   8520
            TabIndex        =   5
            Top             =   2520
            Width           =   1215
         End
         Begin zl9PACSWork.ucFlexGrid ufgDecalin 
            Height          =   2535
            Left            =   120
            TabIndex        =   1
            Top             =   240
            Width           =   9735
            _ExtentX        =   17171
            _ExtentY        =   4471
            DefaultCols     =   ""
            GridRows        =   21
            IsCopyAdoMode   =   0   'False
            IsEjectConfig   =   -1  'True
            HeadFontCharset =   134
            HeadFontWeight  =   400
            DataFontCharset =   134
            DataFontWeight  =   400
         End
      End
   End
   Begin MSComctlLib.Toolbar tbrMain 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   10560
      _ExtentX        =   18627
      _ExtentY        =   1270
      ButtonWidth     =   1455
      ButtonHeight    =   1270
      Style           =   1
      ImageList       =   "imgTbrS"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "标签打印"
            Key             =   "tbLAB"
            ImageIndex      =   7
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
            Caption         =   "-"
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "查看申请"
            Key             =   "tbViewRequest"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "提取材块"
            Key             =   "tbGetMaterials"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "删除材块"
            Key             =   "tbDelMaterials"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "保存材块"
            Key             =   "tbSaveMaterials"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "确认取材"
            Key             =   "tbSureMaterials"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPatholMaterials"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IWorkMenu

Private Const mMustColor As Long = &HC0C0FF

Private Const M_STR_HINT_NoSelectData As String = "无效的检查数据，请选择需要执行的检查记录。"
Private Const M_STR_MODULE_MENU_TAG As String = "取材"


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

Private mStrTemp As String

Private mlngTmpAdviceId As Long
Private mlngTmpSendNo As Long

Private mrecStudy As TStudyStateInf
Attribute mrecStudy.VB_VarHelpID = -1
Private mblnReadOnly As Boolean

Private mObjActiveMenuBar As CommandBar

Private mblnRefreshState As Boolean
Private mbytFontSize As Byte '字号    9--小字体    12--大字体



'获取菜单接口对象
Property Get zlMenu() As IWorkMenu
    Set zlMenu = Me
End Property


Public Sub NotificationRefresh()
'通知刷新
    mblnRefreshState = False
End Sub




Private Sub cmdAutoInputMaterials_Click()
'在巨捡描述中快速录入取材记录信息
    Dim i As Integer, j As Integer
    Dim strTemp As String, strMaterials As String
On Error GoTo errHandle
    
    For i = 1 To ufgData.DataGrid.Rows - 1
        If Not ufgData.IsNullRow(i) Then
            For j = 0 To ufgData.DataGrid.Cols - 1
                If Not ufgData.DataGrid.ColHidden(j) And ufgData.Text(0, ufgData.GetColName(j)) <> "≡" Then
                    strTemp = strTemp & ", " & ufgData.Text(0, ufgData.GetColName(j)) & ":" & ufgData.Text(i, ufgData.GetColName(j))
                End If
            Next
            
            If strTemp <> "" Then
                strMaterials = strMaterials & Mid(strTemp, 3) & vbCrLf
                strTemp = ""
            End If
        End If
    Next
    
    If strMaterials <> "" Then wtDescription.WordText = strMaterials & wtDescription.WordText
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Resize()
On Error GoTo errHandle
    Call ucSplitter1.RePaint(False)
Exit Sub
errHandle:
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
    
    If Not HasMenu(objMenuBar, conMenu_PatholMaterial) Then
        Set cbrMenuBar = mObjActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_PatholMaterial, "取材(&I)", GetPatholMenuIndex(objMenuBar) + 1, False)
        cbrMenuBar.ID = conMenu_PatholMaterial
        cbrMenuBar.Category = M_STR_MODULE_MENU_TAG
        
        
        With cbrMenuBar.CommandBar
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholMaterial_PreviewAll, "标签预览(&V)", "", 0, False)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholMaterial_PrintAll, "标签打印(&P)", "", 0, False)
'            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholMaterial_PreviewSingle, "预览选中标签(&E)", "", 0, False)
'            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholMaterial_PrintSingle, "打印选中标签(&I)", "", 0, False)
            
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholMaterial_RequestView, "申请查看(&R)", "", 0, True)
            
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholMaterial_Get, "材块提取(&G)", "", 0, True)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholMaterial_Del, "材块删除(&D)", "", 0, False)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholMaterial_Save, "材块保存(&S)", "", 0, False)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholMaterial_Sure, "确认取材(&U)", "", 0, False)
            
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholMaterial_Decalcification, "脱钙(&F)", "", 0, True)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholMaterial_ChangeVat, "换缸(&A)", "", 0, False)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholMaterial_CancelVat, "撤销(&C)", "", 0, False)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PahtolMaterial_Finish, "完成(&F)", "", 0, False)
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
        Case conMenu_PatholMaterial_PreviewAll      '预览所有取材标签
            Call PrintMaterialLabel(False)
            
        Case conMenu_PatholMaterial_PrintAll        '打印所有取材标签
            Call PrintMaterialLabel(True)
            
        Case conMenu_PatholMaterial_PreviewSingle   '预览选中标签
            Call PrintSelectMaterialLabel(False)
            
        Case conMenu_PatholMaterial_PrintSingle     '打印选中标签
            Call PrintSelectMaterialLabel(True)
            
        Case conMenu_PatholMaterial_RequestView     '申请查看
            Call ShowMaterialRequest
            
        Case conMenu_PatholMaterial_Get             '材块提取
            Call MaterialGet
            
        Case conMenu_PatholMaterial_Del             '删除选中材块
            Call DelSelectionMaterial
            
        Case conMenu_PatholMaterial_Save            '保存当前取材信息
            Call SaveCurMaterialInf
        
        Case conMenu_PatholMaterial_Sure            '确认取材
            Call SureCurMaterialInf
            
        Case conMenu_PatholMaterial_Decalcification '脱钙
            Call Decalcification
            
        Case conMenu_PatholMaterial_ChangeVat       '换缸
            Call ChangeVat
            
        Case conMenu_PatholMaterial_CancelVat       '撤销
            Call CancelVat
            
        Case conMenu_PahtolMaterial_Finish          '完成
            Call Finish
    End Select
End Sub


Public Sub IWorkMenu_zlUpdateMenu(ByVal control As XtremeCommandBars.ICommandBarControl)
'更新菜单
    Dim blnIsAllowMaterial As Boolean
    
    If Not Me.Visible Then
        control.Enabled = False
        Exit Sub
    End If
    
    blnIsAllowMaterial = CheckPopedom(mstrPrivs, "病理取材") And mlngAdviceID > 0
    
    Select Case control.ID
        Case conMenu_PatholMaterial_PreviewAll
            control.Enabled = blnIsAllowMaterial And mrecStudy.strPatholNumber <> ""
            
        Case conMenu_PatholMaterial_PrintAll
            control.Enabled = blnIsAllowMaterial And mrecStudy.strPatholNumber <> ""
            
        Case conMenu_PatholMaterial_PreviewSingle
            control.Enabled = blnIsAllowMaterial And mrecStudy.strPatholNumber <> ""
            
        Case conMenu_PatholMaterial_PrintSingle
            control.Enabled = blnIsAllowMaterial And mrecStudy.strPatholNumber <> ""
            
        Case conMenu_PatholMaterial_RequestView
            control.Enabled = blnIsAllowMaterial And mrecStudy.strPatholNumber <> ""
        
        Case conMenu_PatholMaterial_Get
            control.Enabled = blnIsAllowMaterial And Not mblnReadOnly
            
        Case conMenu_PatholMaterial_Del
            control.Enabled = blnIsAllowMaterial And Not mblnReadOnly
            
        Case conMenu_PatholMaterial_Save
            control.Enabled = blnIsAllowMaterial And Not mblnReadOnly
            
        Case conMenu_PatholMaterial_Sure
            control.Enabled = blnIsAllowMaterial And Not mblnReadOnly
            
        Case conMenu_PatholMaterial_Decalcification
            control.Enabled = blnIsAllowMaterial And Not mblnReadOnly
            
        Case conMenu_PatholMaterial_ChangeVat
            control.Enabled = blnIsAllowMaterial And Not mblnReadOnly
            
        Case conMenu_PatholMaterial_CancelVat
            control.Enabled = blnIsAllowMaterial And Not mblnReadOnly
            
        Case conMenu_PahtolMaterial_Finish
            control.Enabled = blnIsAllowMaterial And Not mblnReadOnly
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
    Dim lngNewAdviceId As Long
    
    
    If (mlngTmpAdviceId = mlngAdviceID And mlngTmpSendNo = mlngSendNo And mblnRefreshState) And Not blnForceRefresh Then Exit Sub
    
    lngNewAdviceId = mlngAdviceID
    mblnRefreshState = True
    
    If mlngTmpAdviceId <> mlngAdviceID And mlngTmpAdviceId > 0 Then
        '判断取材是否需要进行确认
        If IsNeedMaterialSure Then
            If MsgBoxD(Me, "尚未对检查 [" & mrecStudy.strPatholNumber & "] 进行取材确认，是否需执行确认？", vbYesNo, Me.Caption) = vbYes Then
                mlngAdviceID = mlngTmpAdviceId
                
                Call ExecuteTbrOperation("tbSureMaterials")
            End If
        End If
    End If
    
    mlngAdviceID = lngNewAdviceId
    mlngTmpAdviceId = lngNewAdviceId
    mlngTmpSendNo = mlngSendNo
    
    If lngNewAdviceId <= 0 Then
        Call ConfigMaterialFace(False, "医嘱ID无效请检查。")
        Exit Sub
    Else
        Call ConfigMaterialFace(True)
    End If
    
    Call LoadReportModule
    
    Call GetPatholStudyState(mlngAdviceID, mrecStudy)
    
    
    If Trim(mrecStudy.strPatholNumber) = "" Then
        Call ConfigMaterialFace(False, "该检查尚未生成有效的病理号，请确认该检查是否已被核收。")
        
'        If Not (mobjOwner Is Nothing) Then
'            Call MsgBoxD(Me, "该检查尚未生成有效的病理号，请确认该检查是否已被核收。", vbOKOnly, Me.Caption)
'        End If
        
        Exit Sub
    Else
        Call ConfigMaterialFace(True)
    End If
    
    '判断 如果病人检查类型 是 "细胞" "快速石蜡"，就隐藏脱钙界面
    tsFilter.TabVisible(1) = IIf(mrecStudy.lngStudyType = 2 Or mrecStudy.lngStudyType = 5, False, True)
    
    '配置取材列表
    Call ConfigMaterialList(mrecStudy.lngStudyType)
    '配置脱钙列表
    Call ConfigDecalinList

    
    '配置输入
    Call ConfigGridInput(mrecStudy.lngStudyType)
    
        
    '读取材块记录
    Call LoadMaterialData(mrecStudy.lngPatholAdviceId)
    
    '读取脱钙记录
    Call LoadDecalinData(mlngAdviceID)
    
    '读取巨检描述
    Call LoadDescriptionInf(mrecStudy.lngPatholAdviceId)
    
    
    '刷新材块数量
    Call RefreshMaterialCount
    
    Call ConfigPopedom(mblnReadOnly)
    
'    If Not (owner Is Nothing) Then
'        Call Me.Show(1, owner)
'    End If
End Sub

Public Sub zlRefresh(ByVal lngAdviceID As Long, ByVal blnReadOnly As Boolean, _
    ByVal strPrivs As String, ByVal blnMoved As Boolean, _
    ByVal lngCurDepartmentId As Long, Optional owner As Form = Nothing)
'刷新取材模块
    If lngAdviceID <= 0 Then
        Call ConfigMaterialFace(False, "医嘱ID无效请检查。")
        Exit Sub
    Else
        Call ConfigMaterialFace(True)
    End If
    
    If lngAdviceID <> mlngAdviceID And mlngAdviceID > 0 Then
        '判断取材是否需要进行确认
        If IsNeedMaterialSure Then
            If MsgBoxD(Me, "尚未对检查 [" & mrecStudy.strPatholNumber & "] 进行取材确认，是否需执行确认？", vbYesNo, Me.Caption) = vbYes Then
                Call ExecuteTbrOperation("tbSureMaterials")
            End If
        End If
    End If
    
    
    mlngAdviceID = lngAdviceID
    mstrPrivs = strPrivs
    mblnMoved = blnMoved
    mlngCurDeptId = lngCurDepartmentId
    
    Call LoadReportModule
    
    Call GetPatholStudyState(lngAdviceID, mrecStudy)
    
    
    If Trim(mrecStudy.strPatholNumber) = "" Then
        Call ConfigMaterialFace(False, "该检查尚未生成有效的病理号，请确认该检查是否已被核收。")
        
        If Not (owner Is Nothing) Then
            Call MsgBoxD(Me, "该检查尚未生成有效的病理号，请确认该检查是否已被核收。", vbOKOnly, Me.Caption)
        End If
        
        Exit Sub
    Else
        Call ConfigMaterialFace(True)
    End If
    
    '判断 如果病人检查类型 是 "细胞" "快速石蜡"，就隐藏脱钙界面
    tsFilter.TabVisible(1) = IIf(mrecStudy.lngStudyType = 2 Or mrecStudy.lngStudyType = 5, False, True)
    
    '配置取材列表
    Call ConfigMaterialList(mrecStudy.lngStudyType)
    '配置脱钙列表
    Call ConfigDecalinList

    
    '配置输入
    Call ConfigGridInput(mrecStudy.lngStudyType)

    
    
    '读取材块记录
    Call LoadMaterialData(mrecStudy.lngPatholAdviceId)
    
    '读取脱钙记录
    Call LoadDecalinData(mlngAdviceID)
    
    '读取巨检描述
    Call LoadDescriptionInf(mrecStudy.lngPatholAdviceId)
    
    
    '刷新材块数量
    Call RefreshMaterialCount
    
    Call ConfigPopedom(blnReadOnly)
    
    If Not (owner Is Nothing) Then
        Call Me.Show(1, owner)
    End If
End Sub


Private Sub ConfigPopedom(ByVal blnIsReadOnly As Boolean)
'配置权限
    Dim blnIsAllowMaterial As Boolean
    
    blnIsAllowMaterial = CheckPopedom(mstrPrivs, "病理取材")
    
    tbrMain.Buttons("tbGetMaterials").Enabled = blnIsAllowMaterial And Not blnIsReadOnly
    tbrMain.Buttons("tbDelMaterials").Enabled = blnIsAllowMaterial And Not blnIsReadOnly
    tbrMain.Buttons("tbSaveMaterials").Enabled = blnIsAllowMaterial And Not blnIsReadOnly
    tbrMain.Buttons("tbSureMaterials").Enabled = blnIsAllowMaterial And Not blnIsReadOnly
    tbrMain.Buttons("tbLAB").Enabled = blnIsAllowMaterial And Not blnIsReadOnly
    tbrMain.Buttons("tbViewRequest").Enabled = blnIsAllowMaterial And Not blnIsReadOnly
    
    cmdDecalin.Enabled = blnIsAllowMaterial And Not blnIsReadOnly
    cmdChange.Enabled = blnIsAllowMaterial And Not blnIsReadOnly
    cmdCancel.Enabled = blnIsAllowMaterial And Not blnIsReadOnly
    cmdSucceed.Enabled = blnIsAllowMaterial And Not blnIsReadOnly
    
    
    txtPos.Locked = blnIsReadOnly
    txtPos.BackColor = IIf(blnIsReadOnly, Me.BackColor, vbWhite)
    
    cbxSpecimenProcess.Enabled = blnIsAllowMaterial And Not blnIsReadOnly
    
    wtDescription.ReadOnly = blnIsReadOnly
    
    ufgData.ReadOnly = blnIsReadOnly
    ufgDecalin.ReadOnly = blnIsReadOnly
End Sub



Private Sub ConfigMaterialFace(ByVal blnIsValid As Boolean, Optional ByVal strHintInf As String = "")
'配置取材界面
    tbrMain.Buttons("tbGetMaterials").Enabled = blnIsValid
    tbrMain.Buttons("tbDelMaterials").Enabled = blnIsValid
    tbrMain.Buttons("tbSaveMaterials").Enabled = blnIsValid
    tbrMain.Buttons("tbSureMaterials").Enabled = blnIsValid
    tbrMain.Buttons("tbLAB").Enabled = blnIsValid
    tbrMain.Buttons("tbViewRequest").Enabled = blnIsValid
    
    cmdDecalin.Enabled = blnIsValid
    cmdChange.Enabled = blnIsValid
    cmdCancel.Enabled = blnIsValid
    cmdSucceed.Enabled = blnIsValid
    
    txtPos.Enabled = blnIsValid
    txtPos.BackColor = IIf(Not blnIsValid, Me.BackColor, vbWhite)
    
    cbxSpecimenProcess.Enabled = blnIsValid
    cbxSpecimenProcess.BackColor = IIf(Not blnIsValid, Me.BackColor, vbWhite)
    
    If blnIsValid Then
        Call ufgData.CloseHintInf
        Call ufgDecalin.CloseHintInf
    Else
        Call ufgData.ShowHintInf(strHintInf)
        Call ufgDecalin.ShowHintInf(strHintInf)
        
        wtDescription.WordText = ""
        txtPos.Text = ""
    End If
End Sub


Private Function IsNeedMaterialSure() As Boolean
'是否需要取材确认
    Dim i As Long
    
    IsNeedMaterialSure = False
    
    For i = 1 To ufgData.GridRows - 1
        If ufgData.RowState(i) = TDataRowState.Add And Not ufgData.RowHidden(i) Then
            IsNeedMaterialSure = True
            Exit For
        End If
    Next i
End Function


Private Sub ConfigMaterialList(ByVal lngStudyType As Long)
'配置材块显示列表
    Dim strTemp As String
    
    '设置行数
    ufgData.GridRows = glngStandardRowCount
    '设置行高
    ufgData.RowHeightMin = glngStandardRowHeight
    
    ufgData.ColConvertFormat = gstrMaterialConvertFormat
    
    Select Case lngStudyType
    Case 0, 3, 4, 5
        '判断数据库参数表是否有数据 有则读取数据库参数  没有则加载默认
        strTemp = zlDatabase.GetPara("常规取材列表配置", glngSys, G_LNG_PATHOLSYS_NUM, "")
         
        If strTemp = "" Then
            ufgData.ColNames = gstrNormalMaterialCols
        Else
            ufgData.ColNames = strTemp
        End If
        
        ufgData.DefaultColNames = gstrNormalMaterialCols
    Case 1
        strTemp = zlDatabase.GetPara("冰冻取材列表配置", glngSys, G_LNG_PATHOLSYS_NUM, "")
         
        If strTemp = "" Then
            ufgData.ColNames = gstrIceMaterialCols
        Else
            ufgData.ColNames = strTemp
        End If
        
        ufgData.DefaultColNames = gstrIceMaterialCols
    Case 2
        
        strTemp = zlDatabase.GetPara("细胞取材列表配置", glngSys, G_LNG_PATHOLSYS_NUM, "")
         
        If strTemp = "" Then
            ufgData.ColNames = gstrCellMaterialCols
        Else
            ufgData.ColNames = strTemp
        End If
        
        ufgData.DefaultColNames = gstrCellMaterialCols
    End Select

End Sub

Private Sub ConfigDecalinList()
'配置脱钙显示列表
    ufgDecalin.ColConvertFormat = gstrDecalinConvertFormat
    
    '设置行数
    ufgDecalin.GridRows = glngStandardRowCount
    '设置行高
    ufgDecalin.RowHeightMin = glngStandardRowHeight
    
    Dim strTemp As String
    
    '判断数据库参数表是否有数据 有则读取数据库参数  没有则加载默认
    strTemp = zlDatabase.GetPara("病理脱钙列表配置", glngSys, G_LNG_PATHOLSYS_NUM, "")
     
    If strTemp = "" Then
        '初始化标本显示列表
        ufgDecalin.ColNames = gstrDecalinCols
    Else
        ufgDecalin.ColNames = strTemp
    End If
    
    '禁止右键弹出列表配置窗口
    ufgDecalin.IsEjectConfig = False
    ufgDecalin.DefaultColNames = gstrDecalinCols
    
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
            Call PrintMaterialLabel(False)

        Case UCase("tbLabPrint")
            '打印标签
            Call PrintMaterialLabel(True)
            
        Case UCase("tbGetMaterials")
            '材块提取
            Call MaterialGet
            
        Case UCase("tbDelMaterials")
            '删除材块
            Call DelSelectionMaterial
            
        Case UCase("tbSaveMaterials")
            '保存材块
            Call SaveCurMaterialInf
            
        Case UCase("tbSureMaterials")
            '确认取材
            Call SureCurMaterialInf
            
        Case UCase("tbViewRequest")
            '查看申请
            Call ShowMaterialRequest
            
    End Select
End Sub

Private Sub ufgData_OnColFormartChange()
'根据不同的取材类型保存不同的列头参数

    Select Case mrecStudy.lngStudyType
        Case 0, 3, 4, 5
        
            zlDatabase.SetPara "常规取材列表配置", ufgData.GetColsString(ufgData), glngSys, G_LNG_PATHOLSYS_NUM
            
        Case 1
        
            zlDatabase.SetPara "冰冻取材列表配置", ufgData.GetColsString(ufgData), glngSys, G_LNG_PATHOLSYS_NUM
            
        Case 2
        
            zlDatabase.SetPara "细胞取材列表配置", ufgData.GetColsString(ufgData), glngSys, G_LNG_PATHOLSYS_NUM
            
    End Select

End Sub


Private Sub ufgData_OnColsNameReSet()
On Error GoTo errHandle

    '配置输入
    Call ConfigGridInput(mrecStudy.lngStudyType)
    '读取材块记录
    Call LoadMaterialData(mrecStudy.lngPatholAdviceId)
    '刷新材块数量
    Call RefreshMaterialCount
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgDecalin_OnColFormartChange()
'保存脱钙列表参数
 zlDatabase.SetPara "病理脱钙列表配置", ufgDecalin.GetColsString(ufgDecalin), glngSys, G_LNG_PATHOLSYS_NUM

End Sub

Private Sub ConfigGridInput(ByVal lngStudyType As Long)
'配置输入列表
    Dim strSql As String
    Dim strUsers As String
    Dim strSpecimenName As String
    Dim rsData As ADODB.Recordset
    
    
    '读取主取医师
    strSql = "select a.姓名 from 人员表 a, 部门人员 b where a.id=b.人员ID and b.部门ID=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UserInfo.部门ID)
    
    ufgData.ComboxListFormat(ufgData.GetColIndex(gstrMaterial_主取医师)) = " "
    If rsData.RecordCount > 0 Then
        strUsers = ""
        While Not rsData.EOF
            If Trim(strUsers) <> "" Then strUsers = strUsers & "|"
            
            strUsers = strUsers & Nvl(rsData!姓名)
            
            rsData.MoveNext
        Wend
        
        ufgData.ComboxListFormat(ufgData.GetColIndex(gstrMaterial_主取医师)) = strUsers
        ufgData.ComboxListFormat(ufgData.GetColIndex(gstrMaterial_副取医师)) = " |" & strUsers
        ufgDecalin.ComboxListFormat(ufgDecalin.GetColIndex(gstrDecalin_操作员)) = strUsers
    End If
    
    
    '读取标本名称
    strSql = "select 标本ID, 标本名称,材料类别 from 病理标本信息 where 送检ID>0 and 医嘱ID=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngAdviceID)
    
    ufgData.ComboxListFormat(ufgData.GetColIndex(gstrMaterial_标本名称)) = " "
    ufgDecalin.ComboxListFormat(ufgDecalin.GetColIndex(gstrDecalin_标本名称)) = " "
    If rsData.RecordCount > 0 Then
        strSpecimenName = ""
        
        While Not rsData.EOF
            '判断如果材料类别是 蜡块 并且 检查类型是 会诊 或者 材料类别是 标本的才配置输入
            If Nvl(rsData!材料类别) = 1 And mrecStudy.lngStudyType = 3 Or Nvl(rsData!材料类别) = 0 Then
                If Trim(strSpecimenName) <> "" Then strSpecimenName = strSpecimenName & "|"
                strSpecimenName = strSpecimenName & "#" & Nvl(rsData!标本ID) & ";" & Nvl(rsData!标本名称)
            End If
            
            rsData.MoveNext
        Wend
        
        ufgData.ComboxListFormat(ufgData.GetColIndex(gstrMaterial_标本名称)) = strSpecimenName
        
        '检查类型为 “细胞”“快速石蜡”的情况，没有脱钙功能 即不用加载脱钙列表的标本名称，“常规”“冰冻”“尸检”“会诊”才加载。
        If lngStudyType = 0 Or lngStudyType = 1 Or lngStudyType = 3 Or lngStudyType = 4 Then
            ufgDecalin.ComboxListFormat(ufgDecalin.GetColIndex(gstrDecalin_标本名称)) = strSpecimenName
        End If
    End If
End Sub



Public Sub ReSetFormFontSize(ByVal bytFontSize As Byte)
'功能:重新设置工作站窗体的字体大小
 On Error GoTo errHandle
 
    Dim objCtrl As control
    Dim CtlFont As StdFont
    Dim strFontType  As String
    
    '先移动控件位置
    mbytFontSize = bytFontSize

    
    '再设置字体大小
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
            objCtrl.Height = TextHeight("罗") + 114
        Case UCase("vsFlexGrid")
            objCtrl.Cell(flexcpFont, 0, 0, 0, objCtrl.DataGrid.Cols - 1) = CtlFont
            objCtrl.Font = CtlFont
            objCtrl.RowHeight(0) = TextHeight("罗") + 150
         Case UCase("ucFlexGrid")
            objCtrl.DataGrid.Cell(flexcpFont, 0, 0, 0, objCtrl.DataGrid.Cols - 1) = CtlFont
            objCtrl.DataGrid.Font = CtlFont
            objCtrl.DataGrid.RowHeight(0) = TextHeight("罗") + 150
        Case UCase("ComboBox")
            objCtrl.FontSize = bytFontSize
            objCtrl.FontName = strFontType
        Case UCase("OptionButton")
            objCtrl.FontSize = bytFontSize
            objCtrl.FontName = strFontType
            objCtrl.Width = TextWidth("罗冠" & objCtrl.Caption)
        Case UCase("CheckBox")
            objCtrl.FontSize = bytFontSize
            objCtrl.FontName = strFontType
            objCtrl.Width = TextWidth("罗冠" & objCtrl.Caption)
        Case UCase("DTPicker")
            objCtrl.Font.Size = bytFontSize
            objCtrl.FontN.ame = strFontType
            objCtrl.Width = TextWidth("2012-01-01 23:59:59") * 1.25
            objCtrl.Height = TextHeight("罗") * 1.5
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
    
    Call picMaterial_Resize
    Call picPane1_Resize
    
    Exit Sub
errHandle:
End Sub

'Private Sub LoadMaterialParameter()
'    wtDescription.ModuleHeight = zlDatabase.GetPara("Material_ModuleHeight", glngSys, glngModul, 0)
'    wtDescription.WordWidth = zlDatabase.GetPara("Material_WordWidth", glngSys, glngModul, 0)
'
'    mlngMaterialListHeight = zlDatabase.GetPara("Material_ListHeight", glngSys, glngModul, 0)
'End Sub


'Private Sub SaveMaterialParameter()
'    Call zlDatabase.SetPara("Material_ModuleHeight", wtDescription.ModuleHeight, glngSys, glngModul, True)
'    Call zlDatabase.SetPara("Material_WordWidth", wtDescription.WordWidth, glngSys, glngModul, True)
'    Call zlDatabase.SetPara("Material_ListHeight", mlngMaterialListHeight, glngSys, glngModul, True)
'
'End Sub

Private Sub SwitchWork(ByVal blnIsChangeDescription As Boolean)
'切换工作页面
    framDecalin.Visible = Not blnIsChangeDescription
    framWordEdit.Visible = blnIsChangeDescription
End Sub


Private Sub LoadDecalinData(ByVal lngAdviceID As Long)
'载入脱钙信息
    Dim strSql As String
    Dim rsDecalin As ADODB.Recordset
    
    strSql = "select a.ID,a.标本ID,b.标本名称,a.开始时间, case when a.所需时长 / 60 < 1 then '0' else '' end || to_char(a.所需时长 / 60) as 所需时长, a.开始时间 + a.所需时长/60/24 as 结束时间, a.当前缸次,a.完成状态,a.操作员" & _
                " from 病理脱钙信息 a, 病理标本信息 b " & _
                " where a.标本id = b.标本id and b.医嘱ID =[1] order by a.完成状态, a.开始时间,a.Id"
    
'    If mblnMoved Then strSql = GetMovedDataSql(strSql)
    
    Set ufgDecalin.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngAdviceID)
    
    Call ufgDecalin.RefreshData
End Sub


Private Sub LoadMaterialData(ByVal lngPatholAdviceId As Long)
'载入材块记录信息
    Dim strSql As String
    Dim rsMaterial As ADODB.Recordset
    
    strSql = "select a.材块ID,a.序号, a.标本ID as 标本名称,a.是否脱钙,a.申请ID,a.是否蜡块, a.确认状态,case when a.申请ID>0 then '补取材' else '常规取材' end as 取材类型, a.标本ID,a.取材位置,a.形状,a.颜色,a.性质,a.标本量,a.蜡块数,b.制片数,a.是否冰余,a.主取医师,a.副取医师,a.记录医师,a.取材时间 " & _
                " from 病理取材信息 a,病理制片信息 b" & _
                " where a.病理医嘱ID =[1] and a.材块ID = b.材块ID and (b.申请ID is null or a.申请ID=b.申请ID) " & _
                " union all " & _
                "  select a.材块ID,a.序号, a.标本ID as 标本名称,a.是否脱钙,a.申请ID, a.是否蜡块,a.确认状态,'会诊核收' as 取材类型, " & _
                "  a.标本ID,a.取材位置,a.形状,a.颜色,a.性质,a.标本量,a.蜡块数,0 as 制片数,a.是否冰余,a.主取医师,a.副取医师,'' as 记录医师,a.取材时间 " & _
                "  from 病理取材信息 a,病理检查信息 b, 病理标本信息 c " & _
                "  Where a.病理医嘱ID = [1] And a.病理医嘱ID = b.病理医嘱ID And a.标本ID = c.标本ID And c.材料类别 = 1 And b.检查类型 = 3"

    strSql = "select * from (" & strSql & ") order by 取材类型,序号"

'    If mblnMoved Then strSql = GetMovedDataSql(strSql)
    
    Set ufgData.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPatholAdviceId)
    
    Call ufgData.RefreshData
End Sub


Private Sub LoadDescriptionInf(ByVal lngPatholAdviceId As Long)
'载入巨检描述等信息
    Dim strSql As String
    Dim rsDescription As ADODB.Recordset
    
    strSql = "select 巨检描述,剩余位置,后续处理 from 病理检查信息 where 病理医嘱ID=[1]"
    
'    If mblnMoved Then strSql = GetMovedDataSql(strSql)
    
    Set rsDescription = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPatholAdviceId)
    
    wtDescription.WordText = ""
    txtPos.Text = ""
    
    If rsDescription.RecordCount <= 0 Then Exit Sub
    
    wtDescription.WordText = Nvl(rsDescription("巨检描述").value)
    txtPos.Text = Nvl(rsDescription("剩余位置").value)
    cbxSpecimenProcess.Text = Nvl(rsDescription("后续处理").value)
End Sub


Private Function Decalin_Start(ByVal lngDecalinRowIndex As Long) As String
'开始脱钙
    Dim strSql As String
    Dim lngTimeLen As Long
    Dim dtEndTime As Date
    Dim rsDecalin As ADODB.Recordset
    
    Decalin_Start = ""
    
    strSql = "select Zl_病理脱钙_开始([1],[2],[3],[4]) as 返回值 from dual"
    
    lngTimeLen = Fix(Val(ufgDecalin.Text(lngDecalinRowIndex, gstrDecalin_所需时长)) * 60)
    dtEndTime = CDate(ufgDecalin.Text(lngDecalinRowIndex, gstrDecalin_开始时间))
    
    Set rsDecalin = zlDatabase.OpenSQLRecord(strSql, Me.Caption, _
                                                ufgDecalin.Text(lngDecalinRowIndex, gstrDecalin_标本名称), _
                                                dtEndTime, _
                                                lngTimeLen, _
                                                ufgDecalin.Text(lngDecalinRowIndex, gstrDecalin_操作员) _
                                                )
                                                
    If rsDecalin.RecordCount <= 0 Then
        Decalin_Start = "脱钙执行失败，未能返回有效的脱钙ID。"
        Exit Function
    End If
    
    
    '更新脱钙显示列表
    ufgDecalin.RowState(lngDecalinRowIndex) = TDataRowState.Normal
    
    ufgDecalin.Text(lngDecalinRowIndex, gstrDecalin_ID) = rsDecalin!返回值
    ufgDecalin.Text(lngDecalinRowIndex, gstrDecalin_结束时间) = DateAdd("n", lngTimeLen, dtEndTime)
    ufgDecalin.Text(lngDecalinRowIndex, gstrDecalin_当前缸次) = 1
    ufgDecalin.Text(lngDecalinRowIndex, gstrDecalin_操作员) = UserInfo.姓名
    ufgDecalin.Text(lngDecalinRowIndex, gstrDecalin_当前状态) = "未完成"
    
End Function



Private Sub Decalin_Succed()
'完成脱钙
    Dim strSql As String
    Dim lngDecalinId As Long
    
    lngDecalinId = ufgDecalin.KeyValue(ufgDecalin.SelectionRow)
    
    strSql = "Zl_病理脱钙_完成(" & lngDecalinId & ")"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    '更新脱钙显示列表
    ufgDecalin.Text(ufgDecalin.SelectionRow, gstrDecalin_当前状态) = "已完成"
End Sub



Private Sub Decalin_Change(ByVal dtStart As Date, ByVal lngTimeLen As Double)
'脱钙换缸
    Dim strSql As String
    Dim lngDecalinId As Long
    
    lngDecalinId = ufgDecalin.KeyValue(ufgDecalin.SelectionRow)
    
    strSql = "Zl_病理脱钙_换缸(" & lngDecalinId & "," & To_Date(dtStart) & "," & Fix(lngTimeLen * 60) & ")"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    '更新脱钙显示列表
    ufgDecalin.Text(ufgDecalin.SelectionRow, gstrDecalin_当前缸次) = Val(ufgDecalin.Text(ufgDecalin.SelectionRow, gstrDecalin_当前缸次)) + 1
    ufgDecalin.Text(ufgDecalin.SelectionRow, gstrDecalin_开始时间) = dtStart
    ufgDecalin.Text(ufgDecalin.SelectionRow, gstrDecalin_所需时长) = Format$(lngTimeLen, "0.0")
    ufgDecalin.Text(ufgDecalin.SelectionRow, gstrDecalin_结束时间) = DateAdd("n", lngTimeLen * 60, dtStart)
End Sub



Private Sub Decalin_Cancel()
'撤销脱钙
    Dim strSql As String
    Dim lngDecalinId As Long

    lngDecalinId = Val(ufgDecalin.KeyValue(ufgDecalin.SelectionRow))

    If Trim(lngDecalinId) > 0 Then
        strSql = "Zl_病理脱钙_撤销(" & lngDecalinId & ")"
        Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    End If

    '删除脱钙显示列表
    Call ufgDecalin.DelCurRow
End Sub

Private Sub CancelVat()
'撤销脱钙

    If ufgData.ShowingDataRowCount <= 0 Then Exit Sub
    
    If Not ufgDecalin.IsSelectionRow Then
        Call MsgBoxD(Me, "请选择需要撤销脱钙的记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgDecalin.IsNullRow(ufgDecalin.SelectionRow) Then
        Call MsgBoxD(Me, "请选择需要撤销脱钙的记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    
    '判断当前记录是否已经完成脱钙，已完成的脱钙任务不能进行撤销
    If ufgDecalin.Text(ufgDecalin.SelectionRow, gstrDecalin_当前状态) = "已完成" Then
        Call MsgBoxD(Me, "该标本已完成脱钙，不能进行撤销。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    
    If MsgBoxD(Me, "确认要删除当前未完成的脱钙任务吗？", vbYesNo, Me.Caption) = vbNo Then Exit Sub
    
    Call Decalin_Cancel
    
    Call ConfigDecalcificationBut
End Sub

Private Sub cmdCancel_Click()
On Error GoTo errHandle
    Call CancelVat
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub ChangeVat()
'换缸
    Dim frmChangeInput As frmPatholMaterials_Change
    
    If Not ufgDecalin.IsSelectionRow Then
        Call MsgBoxD(Me, "请选择需要换缸的记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgDecalin.IsNullRow(ufgDecalin.SelectionRow) Then
        Call MsgBoxD(Me, "请选择需要换缸的记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    '判断当前记录是否已经开始脱钙
    If ufgDecalin.IsEmptyKey(ufgDecalin.SelectionRow) Then
        Call MsgBoxD(Me, "该标本尚未开始脱钙，不能执行换缸操作，请先执行脱钙。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgDecalin.Text(ufgDecalin.SelectionRow, gstrDecalin_当前状态) = "已完成" Then
        Call MsgBoxD(Me, "脱钙任务已完成，不能进行换缸操作。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    Set frmChangeInput = New frmPatholMaterials_Change
    
On Error GoTo errFree
    
    Call frmChangeInput.ShowChangeWindow(Me)
        
    If Not frmChangeInput.IsSure Then Exit Sub
    
    '换缸
    Call Decalin_Change(frmChangeInput.StartTime, frmChangeInput.TimeLen)
errFree:
    Unload frmChangeInput
    Set frmChangeInput = Nothing
End Sub


Private Sub cmdChange_Click()
On Error GoTo errHandle
    Call ChangeVat
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Decalcification()
'脱钙

    Dim strErr As String
    Dim blnValid As Boolean
    
    If Not ufgDecalin.IsSelectionRow Then
        Call MsgBoxD(Me, "请录入需要脱钙的相关信息。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgDecalin.IsNullRow(ufgDecalin.SelectionRow) Then
        Call MsgBoxD(Me, "请录入需要脱钙的相关信息。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    '判断当前记录是否已经开始脱钙
    If Not ufgDecalin.IsEmptyKey(ufgDecalin.SelectionRow) Then
        Call MsgBoxD(Me, "该标本已开始执行脱钙操作。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    
    '检查数据是否有效
    blnValid = Not ufgDecalin.IsErrColorWithRow(ufgDecalin.SelectionRow)
    If Not blnValid Then
        Call MsgBoxD(Me, "检测到脱钙列表存在无效数据，请确认是否正确完整的录入，“红色”标记的单元格为必录数据。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    
    '开始脱钙
    strErr = Decalin_Start(ufgDecalin.SelectionRow)
    If Trim(strErr) <> "" Then
        Call MsgBoxD(Me, strErr, vbOKOnly, Me.Caption)
    End If
    
    
    Call ConfigDecalcificationBut
    
End Sub


Private Sub cmdDecalin_Click()
On Error GoTo errHandle
    Call Decalcification
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub DelSelectionMaterial()
'删除选中的材块

    If ufgData.ShowingRowCount <= 0 Then Exit Sub
    
    If Not ufgData.IsSelectionRow Then
        Call MsgBoxD(Me, "请选择需要删除的材块记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgData.IsNullRow(ufgData.SelectionRow) Then
        Call MsgBoxD(Me, "请选择需要删除的材块记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    '判断该材块是否已制片
    If Not ufgData.IsEmptyKey(ufgData.SelectionRow) Then
        If Not CheckAllowUpdate(ufgData.KeyValue(ufgData.SelectionRow)) Then
            Call MsgBoxD(Me, "该材块记录已执行制片处理，不能进行删除。", vbOKOnly, Me.Caption)
            Exit Sub
        End If
    End If
    
    '判断该材块是否已申请制片
    If Not ufgData.IsEmptyKey(ufgData.SelectionRow) Then
        If CheckRequisitionSlices(ufgData.KeyValue(ufgData.SelectionRow)) Then
            Call MsgBoxD(Me, "该材块记录已申请制片处理，不能进行删除。", vbOKOnly, Me.Caption)
            Exit Sub
        End If
    End If
    
    '判断该材块是否已申请特检
    If Not ufgData.IsEmptyKey(ufgData.SelectionRow) Then
        If CheckRequisitionSpeExam(ufgData.KeyValue(ufgData.SelectionRow)) Then
            Call MsgBoxD(Me, "该材块记录已申请特检处理，不能进行删除。", vbOKOnly, Me.Caption)
            Exit Sub
        End If
    End If
    
    If MsgBoxD(Me, "确认要删除选择的材块数据吗？", vbYesNo, Me.Caption) = vbNo Then Exit Sub
    
    '删除行
    Call ufgData.DelCurRow
    
    '保存删除的材块数据
    Call SaveMaterialData(True)
    
    '刷新材块数量
    Call RefreshMaterialCount
End Sub

Private Sub RefreshMaterialCount()
    '刷新材块数量
    Dim lngTotal As Long
    Dim i As Long
    
    lngTotal = 0
    For i = 1 To ufgData.GridRows - 1
        If Not ufgData.RowHidden(i) Then
            If Not ufgData.IsNullRow(i) Then
                
                Select Case mrecStudy.lngStudyType
                    Case StudyType.stNormal, StudyType.stMeet, StudyType.stAutopsy, StudyType.stIce, StudyType.stSpeed
                        lngTotal = lngTotal + Val(ufgData.Text(i, gstrMaterial_蜡块数))
                    Case StudyType.stCell
                        lngTotal = lngTotal + Val(ufgData.Text(i, gstrMaterial_细胞块数))
                End Select
            End If
        End If
    Next i
    
    labRecordInf.Caption = "材块总数：" & lngTotal
End Sub


Private Sub AutoGetMaterialInf()
'自动提取材块信息
    Dim strSql As String
    Dim rsSpeciman As ADODB.Recordset
    Dim lngRow As Long
    Dim i As Long
    Dim dtServicesTime As Date
    Dim strComboboxText As String
    
    strSql = "select a.标本ID,a.标本名称,a.材料类别,b.默认标本量,b.默认制片数 from 病理标本信息 a,病理检查标本 b where a.标本名称 = b.标本名称(+) and a.医嘱ID=[1]"
'    If mblnMoved Then strSql = GetMovedDataSql(strSql)
    Set rsSpeciman = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngAdviceID)
    
    If rsSpeciman.RecordCount <= 0 Then Exit Sub
    
    lngRow = 0
    dtServicesTime = zlDatabase.Currentdate
    
    Do While Not rsSpeciman.EOF
        
        '如果核收材料不为标本类型，则不加载数据，直接跳过
        If Nvl(rsSpeciman!材料类别) <> 0 Then GoTo continue
        
        For i = 1 To ufgData.GridRows - 1
            If (ufgData.Text(i, gstrMaterial_标本名称) = rsSpeciman("标本ID").value Or ufgData.Text(i, gstrMaterial_标本名称) = rsSpeciman("标本名称").value) _
                And Not ufgData.RowHidden(i) Then GoTo continue
            
            If ufgData.IsNullRow(i) And Not ufgData.RowHidden(i) Then
                lngRow = i
                Exit For
            End If
        Next i
        If lngRow = 0 Then Exit Do
        
        ufgData.Text(lngRow, gSpecimen_标本名称) = Nvl(rsSpeciman!标本ID)
        
        ufgData.Text(lngRow, gstrMaterial_制片数) = IIf(Nvl(rsSpeciman!默认制片数) = "", 1, Nvl(rsSpeciman!默认制片数))
        
        '判断如果检查类型为2则读取默认标本量
        If mrecStudy.lngStudyType = 2 Then
            ufgData.Text(lngRow, gstrMaterial_标本量) = Nvl(rsSpeciman!默认标本量)
        End If

        '在材块提取中不需要取材时间,取材时间是确认取材后才插入
        'ufgData.Text(lngRow, gstrMaterial_取材时间) = dtServicesTime
        
        If mrecStudy.lngStudyType <> StudyType.stCell And mrecStudy.lngStudyType <> StudyType.stIce Then
            ufgData.Text(lngRow, gstrMaterial_蜡块数) = "1"
            If ufgData.Text(lngRow, gstrMaterial_是否蜡块) = "" Then ufgData.Text(lngRow, gstrMaterial_是否蜡块) = "1-是"
            
        Else
            If mrecStudy.lngStudyType = StudyType.stIce Then
                ufgData.Text(lngRow, gstrMaterial_是否冰余) = "0-否"
                ufgData.Text(lngRow, gstrMaterial_蜡块数) = "0"
            Else
                ufgData.Text(lngRow, gstrMaterial_细胞块数) = "0"
            End If
            
            If ufgData.Text(lngRow, gstrMaterial_是否蜡块) = "" Then ufgData.Text(lngRow, gstrMaterial_是否蜡块) = "0-否"
        End If
        
        If Not IsDate(ufgData.Text(lngRow, gstrMaterial_取材时间)) Then
            ufgData.Text(lngRow, gstrMaterial_取材时间) = dtServicesTime
        End If
        
        
        If ufgData.Text(lngRow, gstrMaterial_主取医师) = "" Then
            If lngRow - 1 > 0 Then
                If ufgData.Text(lngRow - 1, gstrMaterial_主取医师) <> "" Then
                    ufgData.Text(lngRow, gstrMaterial_主取医师) = ufgData.Text(lngRow - 1, gstrMaterial_主取医师)
                End If
            End If
            
            If ufgData.Text(lngRow, gstrMaterial_主取医师) = "" Then
                strComboboxText = ufgData.DataGrid.ColComboList(ufgData.GetColIndex(gstrMaterial_主取医师))
                
                If strComboboxText <> "" Then
                    If InStr(strComboboxText, "|") > 0 Then
                        strComboboxText = Mid(strComboboxText, 1, InStr(strComboboxText, "|") - 1)
                    End If
                    ufgData.Text(lngRow, gstrMaterial_主取医师) = strComboboxText
                    
                End If
            End If
        End If
        
        '更新当前行状态为添加
        ufgData.RowState(lngRow) = TDataRowState.Add
        
        Call ufgData_OnAfterEdit(lngRow, ufgData.GetColIndex(gSpecimen_标本名称))
        
continue:
        rsSpeciman.MoveNext
    Loop
    
    '提示用户
    If ufgData.ShowingDataRowCount = 0 Then
        Call MsgBoxD(Me, "核收材料不为标本类型，不能自动提取材块信息。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
End Sub


Private Sub MaterialGet()
'提取材块
    '非取材阶段，不能进行确认
    If mrecStudy.lngMaterialStep <> TExecuteStep.NeedDo Then
        Call MsgBoxD(Me, "尚未进入取材阶段，不能自动提取材块信息。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    '自动提取材块信息
    Call AutoGetMaterialInf
    
    '刷新材块数量
    Call RefreshMaterialCount
End Sub


Private Sub PrintMaterialLabel(Optional ByVal blnIsPrint As Boolean = True)
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
    
    Call zlReport.ReportOpen(gcnOracle, 100, "ZL1_Inside_1294_01", Me, "材块ID1=" & strValue(0), "材块ID2=" & strValue(1), "材块ID3=" & strValue(2), "材块ID4=" & strValue(3), "材块ID5=" & strValue(4), "材块ID6=" & strValue(5), IIf(blnIsPrint, 2, 1))
End Sub



Private Sub PrintSelectMaterialLabel(Optional ByVal blnIsPrint As Boolean = True)
'打印选择的材块标签
On Error GoTo errHandle
    Dim strValue(5) As String
    
    
    If Not ufgData.IsSelectionRow Then
        Call MsgBoxD(Me, "请选择需要打印的材块记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgData.IsNullRow(ufgData.SelectionRow) Then
        Call MsgBoxD(Me, "请选择需要打印的材块记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    strValue(0) = ufgData.KeyValue(ufgData.SelectionRow)
    strValue(1) = "0": strValue(2) = "0": strValue(3) = "0": strValue(4) = "0": strValue(5) = "0"

    
    Call zlReport.ReportOpen(gcnOracle, 100, "ZL1_Inside_1294_01", Me, "材块ID1=" & strValue(0), "材块ID2=" & strValue(1), "材块ID3=" & strValue(2), "材块ID4=" & strValue(3), "材块ID5=" & strValue(4), "材块ID6=" & strValue(5), IIf(blnIsPrint, 2, 1))
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub ShowMaterialRequest()
'显示特检申请
Dim frmRequestView As New frmPatholRequisition_View
On Error GoTo errFree
    Call frmRequestView.ShowRequestViewWind(mrecStudy.lngPatholAdviceId, 4, mblnMoved, Me)
errFree:
    Call Unload(frmRequestView)
    Set frmRequestView = Nothing
End Sub


'Private Sub CmdRefresh_Click()
'On Error GoTo errHandle
'    '恢复列表数据
'    Call ufgData.RefreshData
'
'    '刷新材块数量
'    Call RefreshMaterialCount
'
'    Exit Sub
'errHandle:
'    If ErrCenter() = 1 Then Resume
'End Sub

Private Sub SaveCurMaterialInf()
'保存当前取材信息

    Dim blnValid As Boolean
    
    '材块保存
    If ufgData.ShowingDataRowCount <= 0 Then
        Call MsgBoxD(Me, "没有找到需要保存的材块信息，请录入材块数据。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    blnValid = Not ufgData.IsErrColorWithList
    If Not blnValid Then
        Call MsgBoxD(Me, "检测到取材列表中存在无效数据，请确认相关数据是否正确完整的录入，“红色”标记的单元格为必录数据。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    Call SaveMaterialData
    
    Call SendMsgToMainWindow(Me, wetMaterialSave, mlngAdviceID)
    
    Call MsgBoxD(Me, "数据已成功保存。", vbOKOnly, Me.Caption)
    
    '刷新材块数量
    Call RefreshMaterialCount
    
End Sub


Private Sub Finish()
'完成脱钙

    If ufgDecalin.ShowingRowCount <= 0 Then Exit Sub

    If Not ufgDecalin.IsSelectionRow Then
        Call MsgBoxD(Me, "请选择需要完成脱钙的记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgDecalin.IsNullRow(ufgDecalin.SelectionRow) Then
        Call MsgBoxD(Me, "请选择需要完成脱钙的记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    '判断当前记录是否已经开始脱钙
    If ufgDecalin.IsEmptyKey(ufgDecalin.SelectionRow) Then
        Call MsgBoxD(Me, "该标本尚未开始脱钙，不能执行该操作，请先执行脱钙。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    Call Decalin_Succed
End Sub

Private Sub cmdSucceed_Click()
On Error GoTo errHandle
    Call Finish

Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub



Private Function GetRequisitionId(Optional ByVal lngRequisitionType As Long = 4) As Long
'获取申请ID
'lngRequisitionType:默认为4，表示补取材
'根据申请类型获取尚未执行的申请ID
'如果没有补取申请记录，则返回空申请ID

    Dim strSql As String
    Dim rsRequisition As ADODB.Recordset
    
    GetRequisitionId = -1
    
    strSql = "select 申请ID from 病理申请信息 where 申请状态=0 and 病理医嘱ID=[1] and 申请类型=[2]"
'    If mblnMoved Then strSql = GetMovedDataSql(strSql)
    
    Set rsRequisition = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mrecStudy.lngPatholAdviceId, lngRequisitionType)
    
    If rsRequisition.RecordCount > 0 Then GetRequisitionId = rsRequisition("申请ID").value
End Function



Private Function CheckAllowUpdate(ByVal strMaterialId As String) As Boolean
'检查是否允许更新
'未制片的材块均可进行更新,通过检查病理制片信息表，可判断材块是否已制片(如果当前状态不为0，则已制片)

    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    CheckAllowUpdate = True
    
    If Trim(strMaterialId) = "" Then Exit Function
    
    strSql = "select ID from 病理制片信息 where  材块ID=[1] and 当前状态<>0"
'    If mblnMoved Then strSql = GetMovedDataSql(strSql)
    
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strMaterialId)
    
    If rsData.RecordCount > 0 Then CheckAllowUpdate = False
End Function

Private Function CheckRequisitionSlices(ByVal strMaterialId As String) As Boolean
'检查是否已申请制片
'已执行制片申请的材块取材界面不能删除材块

    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    CheckRequisitionSlices = False
    
    If Trim(strMaterialId) = "" Then Exit Function
    
    strSql = "select 申请ID from 病理制片信息 where 申请ID is not null and 材块ID=[1]"
'    If mblnMoved Then strSql = GetMovedDataSql(strSql)
    
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strMaterialId)
    
    If rsData.RecordCount > 0 Then CheckRequisitionSlices = True
End Function

Private Function CheckRequisitionSpeExam(ByVal strMaterialId As String) As Boolean
'检查是否已申请特检
'已申请特检处理的材块取材界面不能删除材块

    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    CheckRequisitionSpeExam = False
    
    If Trim(strMaterialId) = "" Then Exit Function
    
    strSql = "select 申请ID from 病理特检信息 where  材块ID=[1]"
'    If mblnMoved Then strSql = GetMovedDataSql(strSql)
    
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strMaterialId)
    
    If rsData.RecordCount > 0 Then CheckRequisitionSpeExam = True
End Function


Public Sub SureMaterialData()
'确认取材数据
    Dim strSql As String
    Dim i As Long
    
    strSql = "Zl_病理取材_确认('" & mrecStudy.lngPatholAdviceId & "')"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    For i = 1 To ufgData.GridRows - 1
        If Not ufgData.IsNullRow(i) Then
            ufgData.Text(i, gstrMaterial_确认状态) = "已确认"
        End If
    Next i
    
    
    mrecStudy.lngSlicesStep = 1
    mrecStudy.lngMaterialStep = 2
End Sub


Private Sub SplitMaterialNumber(ByVal strDataValue As String, ByRef strID As String, ByRef strSeq As String)
'分解由执行过程返回的材块号码
    Dim lngFind As Long
    
    lngFind = InStr(strDataValue, "-")
    
    If lngFind <= 0 Then Exit Sub
    
    strID = Mid(strDataValue, 1, lngFind - 1)
    strSeq = Mid(strDataValue, lngFind + 1, 18)
End Sub


Public Sub SaveMaterialData(Optional ByVal blnIsSaveOnlyDel As Boolean = False)
'------------------------------------------------------------------------------
'blnIsSaveOnlyDel:是否仅仅保存删除的数据
'------------------------------------------------------------------------------


'取材确认保存
'如果没有新的材块，则只保存巨检描述和剩余位置
'如果已制片，则不能进行更新操作


    Dim i As Long
    Dim strSql As String
    Dim rsResult As ADODB.Recordset
    Dim lngRequisitionId As Long
    Dim dtSerivcesTime As Date
    Dim strNewId As String
    Dim strNewSeq As String
    Dim lngCount As Long
    
    
    '获取补取的申请ID，如果没有补取申请记录，则返回空申请ID
    lngRequisitionId = GetRequisitionId
    
    For i = 1 To ufgData.GridRows - 1
        If ufgData.RowState(i) = TDataRowState.Add And Not blnIsSaveOnlyDel Then
            
            dtSerivcesTime = ufgData.Text(i, gstrMaterial_取材时间) 'zlDatabase.Currentdate
            
            '添加取材记录
            Select Case mrecStudy.lngStudyType
                Case StudyType.stNormal, StudyType.stMeet, StudyType.stAutopsy, StudyType.stSpeed
                    strSql = "select Zl_病理取材_常规([1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14])  as 返回值 from dual"
                    
                    '如果蜡块数显示在界面中，则以具体的录入数量为准
                    lngCount = 1
                    If Not ufgData.DataGrid.ColHidden(ufgData.GetColIndex(gstrMaterial_蜡块数)) Then
                        lngCount = Val(ufgData.Text(i, gstrMaterial_蜡块数))
                    End If
                    
                    Set rsResult = zlDatabase.OpenSQLRecord(strSql, Me.Caption, _
                                                            mrecStudy.lngPatholAdviceId, _
                                                            IIf(lngRequisitionId <= 0, "", lngRequisitionId), _
                                                            ufgData.Text(i, gstrMaterial_标本名称), _
                                                            ufgData.DisplayText(i, gstrMaterial_标本名称), _
                                                            ufgData.Text(i, gstrMaterial_取材位置), _
                                                            ufgData.Text(i, gstrMaterial_形状), _
                                                            lngCount, _
                                                            Val(ufgData.Text(i, gstrMaterial_制片数)), _
                                                            ufgData.Text(i, gstrMaterial_主取医师), _
                                                            ufgData.Text(i, gstrMaterial_副取医师), _
                                                            Val(ufgData.Text(i, gstrMaterial_是否蜡块)), _
                                                            Val(ufgData.Text(i, gstrMaterial_是否脱钙)), _
                                                            UserInfo.姓名, _
                                                            CDate(dtSerivcesTime))
                                                            
                                                            
                Case StudyType.stIce
                    strSql = "select Zl_病理取材_冰冻([1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14])  as 返回值 from dual"
                    
                    '如果蜡块数显示在界面中，则以具体的录入数量为准
                    lngCount = 1
                    If Not ufgData.DataGrid.ColHidden(ufgData.GetColIndex(gstrMaterial_蜡块数)) Then
                        lngCount = Val(ufgData.Text(i, gstrMaterial_蜡块数))
                    End If
                    
                    Set rsResult = zlDatabase.OpenSQLRecord(strSql, Me.Caption, _
                                                            mrecStudy.lngPatholAdviceId, _
                                                            IIf(lngRequisitionId <= 0, "", lngRequisitionId), _
                                                            ufgData.Text(i, gstrMaterial_标本名称), _
                                                            ufgData.DisplayText(i, gstrMaterial_标本名称), _
                                                            ufgData.Text(i, gstrMaterial_取材位置), _
                                                            ufgData.Text(i, gstrMaterial_形状), _
                                                            Val(ufgData.Text(i, gstrMaterial_是否蜡块)), _
                                                            Val(ufgData.Text(i, gstrMaterial_是否冰余)), _
                                                            lngCount, _
                                                            Val(ufgData.Text(i, gstrMaterial_制片数)), _
                                                            ufgData.Text(i, gstrMaterial_主取医师), _
                                                            ufgData.Text(i, gstrMaterial_副取医师), _
                                                            UserInfo.姓名, _
                                                            CDate(dtSerivcesTime))
                Case StudyType.stCell
                    strSql = "select Zl_病理取材_细胞([1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14]) as 返回值 from dual"
                    
                    Set rsResult = zlDatabase.OpenSQLRecord(strSql, Me.Caption, _
                                                            mrecStudy.lngPatholAdviceId, _
                                                            IIf(lngRequisitionId <= 0, "", lngRequisitionId), _
                                                            ufgData.Text(i, gstrMaterial_标本名称), _
                                                            ufgData.DisplayText(i, gstrMaterial_标本名称), _
                                                            ufgData.Text(i, gstrMaterial_颜色), _
                                                            ufgData.Text(i, gstrMaterial_性质), _
                                                            ufgData.Text(i, gstrMaterial_标本量), _
                                                            Val(ufgData.Text(i, gstrMaterial_制片数)), _
                                                            Val(ufgData.Text(i, gstrMaterial_是否蜡块)), _
                                                            Val(ufgData.Text(i, gstrMaterial_细胞块数)), _
                                                            ufgData.Text(i, gstrMaterial_主取医师), _
                                                            ufgData.Text(i, gstrMaterial_副取医师), _
                                                            UserInfo.姓名, _
                                                            CDate(dtSerivcesTime))
            End Select
            
            
            If rsResult.RecordCount <= 0 Then
                Call err.Raise(0, "SaveMaterialData", "未成功获取新增后的材块号,处理失败。")
                Exit Sub
            End If
            
            Call SplitMaterialNumber(rsResult("返回值").value, strNewId, strNewSeq)
            
            '更新材块列表
            ufgData.Text(i, gstrMaterial_材块ID) = strNewId
            ufgData.Text(i, gstrMaterial_材块号) = strNewSeq
            ufgData.Text(i, gstrMaterial_记录医师) = UserInfo.姓名
            ufgData.Text(i, gstrMaterial_取材时间) = dtSerivcesTime
            ufgData.Text(i, gstrMaterial_取材类型) = IIf(lngRequisitionId > 0, "补取材", "常规取材")
            ufgData.Text(i, gstrMaterial_确认状态) = "未确认"
            
        ElseIf ufgData.RowState(i) = TDataRowState.Update And Not blnIsSaveOnlyDel Then
            dtSerivcesTime = ufgData.Text(i, gstrMaterial_取材时间)
            
            '更新取材记录
            Select Case mrecStudy.lngStudyType
                Case StudyType.stNormal, StudyType.stMeet, StudyType.stAutopsy, StudyType.stSpeed
                    strSql = "Zl_病理取材_常规更新('" & ufgData.KeyValue(i) & "','" & _
                                                        ufgData.Text(i, gstrMaterial_取材位置) & "','" & _
                                                        ufgData.Text(i, gstrMaterial_形状) & "'," & _
                                                        Val(ufgData.Text(i, gstrMaterial_蜡块数)) & "," & _
                                                        Val(ufgData.Text(i, gstrMaterial_制片数)) & ",'" & _
                                                        ufgData.Text(i, gstrMaterial_主取医师) & "','" & _
                                                        ufgData.Text(i, gstrMaterial_副取医师) & "'," & _
                                                        Val(ufgData.Text(i, gstrMaterial_是否蜡块)) & "," & _
                                                        Val(ufgData.Text(i, gstrMaterial_是否脱钙)) & "," & _
                                                        To_Date(dtSerivcesTime) & ")"
                    
                    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
                                                            
                                                            
                Case StudyType.stIce
                    strSql = "Zl_病理取材_冰冻更新('" & ufgData.KeyValue(i) & "','" & _
                                                        ufgData.Text(i, gstrMaterial_取材位置) & "','" & _
                                                        ufgData.Text(i, gstrMaterial_形状) & "'," & _
                                                        Val(ufgData.Text(i, gstrMaterial_是否冰余)) & "," & _
                                                        Val(ufgData.Text(i, gstrMaterial_是否蜡块)) & "," & _
                                                        Val(ufgData.Text(i, gstrMaterial_蜡块数)) & "," & _
                                                        Val(ufgData.Text(i, gstrMaterial_制片数)) & ",'" & _
                                                        ufgData.Text(i, gstrMaterial_主取医师) & "','" & _
                                                        ufgData.Text(i, gstrMaterial_副取医师) & "'," & _
                                                        To_Date(dtSerivcesTime) & ")"
                    
                    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
                    
                Case StudyType.stCell
                    strSql = "Zl_病理取材_细胞更新('" & ufgData.KeyValue(i) & "','" & _
                                                        ufgData.Text(i, gstrMaterial_颜色) & "','" & _
                                                        ufgData.Text(i, gstrMaterial_性质) & "','" & _
                                                        ufgData.Text(i, gstrMaterial_标本量) & "'," & _
                                                        Val(ufgData.Text(i, gstrMaterial_制片数)) & "," & _
                                                        Val(ufgData.Text(i, gstrMaterial_是否蜡块)) & "," & _
                                                        Val(ufgData.Text(i, gstrMaterial_细胞块数)) & ",'" & _
                                                        ufgData.Text(i, gstrMaterial_主取医师) & "','" & _
                                                        ufgData.Text(i, gstrMaterial_副取医师) & "'," & _
                                                        To_Date(dtSerivcesTime) & ")"
                    
                    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
            End Select
        
            
        ElseIf ufgData.RowState(i) = TDataRowState.Del Then
            '删除取材记录
            If Trim(ufgData.KeyValue(i)) <> "" Then
                strSql = "Zl_病理取材_删除('" & ufgData.KeyValue(i) & "')"
                Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
            End If
        End If
        
        
        '更新行状态
        ufgData.RowState(i) = TDataRowState.Normal
    Next i
    
    
    '保存取材信息(巨检描述，剩余位置等)
    strSql = "Zl_病理取材_信息保存('" & mrecStudy.lngPatholAdviceId & "','" & wtDescription.WordText & "','" & txtPos.Text & "','" & cbxSpecimenProcess.Text & "')"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    

End Sub


Private Sub SureCurMaterialInf()
'确认当前取材信息
    Dim i As Integer
    Dim blnValid As Boolean
    
    '如果有脱钙未完成的，不能进行确认
    If Not ufgDecalin.IsNullRow(i) And ufgDecalin.RowState(i) <> Del Then
        For i = 1 To ufgDecalin.DataGrid.Rows - 1
            If ufgDecalin.Text(i, gstrDecalin_当前状态) <> "" And ufgDecalin.Text(i, gstrDecalin_当前状态) <> "已完成" Then
                Call MsgBoxD(Me, "还有脱钙未完成，不能进行取材确认操作。", vbOKOnly, Me.Caption)
                Exit Sub
            End If
        Next
    End If
    
    '非取材阶段，不能进行确认
    If mrecStudy.lngMaterialStep <> TExecuteStep.NeedDo Then
        Call MsgBoxD(Me, "尚未进入取材阶段，不能进行取材确认操作。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    '材块保存
    If ufgData.ShowingDataRowCount <= 0 Then
        Call MsgBoxD(Me, "没有找到需要确认的材块信息，请录入材块数据。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    blnValid = Not ufgData.IsErrColorWithList
    If Not blnValid Then
        Call MsgBoxD(Me, "检测到取材列表中存在无效数据，请确认相关数据是否正确完整的录入，“红色”标记的单元格为必录数据。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    '保存材块记录(先保存材块记录)
    Call SaveMaterialData
    
    '确认取材
    Call SureMaterialData
    
    '触发事件
    Call SendMsgToMainWindow(Me, wetMaterialSure, mlngAdviceID)
    
    Call MsgBoxD(Me, "已完成取材确认。", vbOKOnly, Me.Caption)

End Sub


Private Sub Form_Initialize()
    Dim strRegPath As String
    Set zlReport = New zl9Report.clsReport
    
    strRegPath = "公共模块\" & App.ProductName & "\" & Me.Name
    picMaterial.Height = Val(GetSetting("ZLSOFT", strRegPath, "MaterialListHeight", picMaterial.Height))
End Sub


Private Sub LoadReportModule()
'载入报告模板
    Dim strLinkClassName As String
    
    If mlngCurDeptId = wtDescription.CurDepartId Then Exit Sub
    
    strLinkClassName = zlDatabase.GetPara("巨检描述模板", glngSys, glngModul, "")
    
    wtDescription.ModuleName = strLinkClassName
    wtDescription.CurDepartId = mlngCurDeptId
    
    Call wtDescription.LoadWordModel
End Sub


Private Sub Form_Load()
On Error GoTo errHandle
    
    Call SwitchWork(True)
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Dim strRegPath As String
    
'    '保存参数
'    Call SaveMaterialParameter
    strRegPath = "公共模块\" & App.ProductName & "\" & Me.Name
    Call SaveSetting("ZLSOFT", strRegPath, "MaterialListHeight", picMaterial.Height)
    
    Set zlReport = Nothing
End Sub



Private Sub picMaterial_Resize()
On Error Resume Next
    framMaterial.Left = 0
    framMaterial.Top = 0
    framMaterial.Width = picMaterial.Width
    framMaterial.Height = picMaterial.Height
    
    txtPos.Height = IIf(mbytFontSize = 9, 315, 330)
    txtPos.Top = framMaterial.Height - txtPos.Height - 120 + IIf(mbytFontSize = 9, 0, 30)
    
    
    ufgData.Left = 120
    ufgData.Top = 240 + IIf(mbytFontSize = 9, 0, 120)
    ufgData.Width = framMaterial.Width - 240
    ufgData.Height = framMaterial.Height - txtPos.Height - 480

    labRecordInf.Left = 120
    labRecordInf.Top = txtPos.Top + 60
    
    labInf.Left = labRecordInf.Left + labRecordInf.Width + 400
    labInf.Top = labRecordInf.Top
    
    txtPos.Left = labInf.Left + labInf.Width
    
    labSpecimenProcess.Left = txtPos.Left + txtPos.Width + 900
    labSpecimenProcess.Top = labInf.Top
    
    cbxSpecimenProcess.Left = labSpecimenProcess.Left + labSpecimenProcess.Width - 150
    cbxSpecimenProcess.Top = txtPos.Top
    
    cmdAutoInputMaterials.Left = cbxSpecimenProcess.Left + cbxSpecimenProcess.Width + 400
    cmdAutoInputMaterials.Top = cbxSpecimenProcess.Top
End Sub


Private Sub picPane1_Resize()
On Error Resume Next

    tsFilter.Left = 0
    tsFilter.Top = 0
    tsFilter.Width = picPane1.Width

    '脱钙管理----------------------------------------------------
    framDecalin.Left = 0
    framDecalin.Top = tsFilter.Top + tsFilter.Height + 10
    framDecalin.Width = picPane1.Width
    framDecalin.Height = picPane1.Height - tsFilter.Height - 120

    ufgDecalin.Left = 120
    ufgDecalin.Top = 240
    ufgDecalin.Width = framDecalin.Width - 240
    ufgDecalin.Height = framDecalin.Height - cmdSucceed.Height - 480

    cmdSucceed.Left = framDecalin.Width - cmdSucceed.Width - 120
    cmdSucceed.Top = framDecalin.Height - cmdSucceed.Height - 120

    cmdCancel.Left = cmdSucceed.Left - cmdCancel.Width - 120
    cmdCancel.Top = cmdSucceed.Top

    cmdChange.Left = cmdCancel.Left - cmdChange.Width - 120
    cmdChange.Top = cmdSucceed.Top

    cmdDecalin.Left = cmdChange.Left - cmdDecalin.Width - 120
    cmdDecalin.Top = cmdSucceed.Top



    '巨检描述----------------------------------------------------
    framWordEdit.Left = 0
    framWordEdit.Top = tsFilter.Top + tsFilter.Height + 10
    framWordEdit.Width = picPane1.Width
    framWordEdit.Height = picPane1.Height - tsFilter.Height - 120


    wtDescription.Left = 0
    wtDescription.Top = 0
    wtDescription.Width = framWordEdit.Width
    wtDescription.Height = framWordEdit.Height
End Sub

Private Sub tsFilter_Click(PreviousTab As Integer)
On Error GoTo errHandle
    Call SwitchWork(IIf(tsFilter.Tab = 0, True, False))
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgData_OnChangeEdit()
'如果是标本名称改变 才执行跳到下个编辑框
    If ufgData.DataGrid.Col = ufgData.GetColIndex(gstrMaterial_标本名称) Then
        ufgData.EditNextCellWithCurRow
    End If

End Sub


Private Sub ufgData_OnAfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim iCol As Long
    Dim i As Long
    Dim iCount As Long
    Dim strNewSpecimenName As String
    Dim strSql As String
    Dim rsInspectionSpe As ADODB.Recordset
    Dim rsMaterislType As ADODB.Recordset
    
    '判断该行是不是会诊核收并且已确认，是则不执行删除
    If ufgData.DataGrid.Cell(flexcpTextDisplay, Row, ufgData.GetColIndex(gstrMaterial_确认状态)) <> "已确认" And ufgData.DataGrid.Cell(flexcpTextDisplay, Row, ufgData.GetColIndex(gstrMaterial_取材类型)) <> "会诊核收" Then
        '判断该行是否隐藏，是则表示改行以被删除 不执行删除*
        If Not ufgData.DataGrid.RowHidden(Row) Then
            strSql = "select 材料类别 from 病理标本信息 where 标本名称=[1] and 医嘱ID=[2]"
            Set rsMaterislType = zlDatabase.OpenSQLRecord(strSql, Me.Caption, ufgData.DataGrid.Cell(flexcpTextDisplay, Row, ufgData.GetColIndex(gstrMaterial_标本名称)), mlngAdviceID)
            '判断数据集是否有数据，没有则不执行删除
            If rsMaterislType.RecordCount > 0 Then
                '判断材料类别是否为蜡块 是则删除数据行
                If Nvl(rsMaterislType!材料类别) = 1 Then
                    Call MsgBoxD(Me, "该标本的材料类型为蜡块，不能进行取材操作。", vbOKOnly, Me.Caption)
    
                    '删除数据行
                    Call ufgData.DelCurRow
                    Exit Sub
                End If
            End If
        End If
    End If

    If ufgData.IsNullRow(Row) Then
        ufgData.RowState(Row) = TDataRowState.Normal
        
        Call ufgData.SetRowColor(Row, ufgData.BackColor)
        Exit Sub
    End If
    
    If ufgData.Text(Row, gstrMaterial_制片数) = "" Or ufgData.Text(Row, gstrMaterial_标本量) = "" Then
        '如果制片数或者标本量有一个为空则执行查询
        strSql = "select 标本名称,默认标本量,默认制片数 from 病理检查标本"
        Set rsInspectionSpe = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    End If
    
    
      '如果制片数为空则执行读取数据  不为空则跳过
    If ufgData.Text(Row, gstrMaterial_制片数) = "" Then
        If rsInspectionSpe.RecordCount > 0 Then
            rsInspectionSpe.MoveFirst
            Do While Not rsInspectionSpe.EOF
                '如果显示的标本名称与 病理检查标本中的匹配则读取默认制片数
                If ufgData.DataGrid.Cell(flexcpTextDisplay, Row, ufgData.GetColIndex(gstrMaterial_标本名称)) = Nvl(rsInspectionSpe!标本名称) Then
                    ufgData.Text(Row, gstrMaterial_制片数) = Nvl(rsInspectionSpe!默认制片数)
                    Exit Do
                Else
                    ufgData.Text(Row, gstrMaterial_制片数) = 1
                End If
                rsInspectionSpe.MoveNext
            Loop
        End If
    End If

    
    Select Case mrecStudy.lngStudyType
        Case StudyType.stNormal, StudyType.stMeet, StudyType.stAutopsy, StudyType.stSpeed  '常规,会诊,尸体检查
        
            If Val(ufgData.Text(Row, gstrMaterial_蜡块数)) < 1 And ufgData.Text(Row, gstrMaterial_是否蜡块) = "1-是" Then
                ufgData.Text(Row, gstrMaterial_蜡块数) = "1"
            End If
        
            '如果界面中显示了蜡块数，则蜡块数必须录入，否则修改单元格颜色
            If Not ufgData.DataGrid.ColHidden(ufgData.GetColIndex(gstrMaterial_蜡块数)) Then
                
                '如果未录入蜡块数量，则显示淡红色
                iCol = ufgData.GetColIndex(gstrMaterial_蜡块数)
                
                ufgData.CellColor(Row, iCol) = IIf(Val(ufgData.Text(Row, gstrMaterial_蜡块数)) <= 0, ufgData.ErrCellColor, ufgData.BackColor)
            End If
            
        Case StudyType.stIce  '冰冻检查
        
            '如果蜡块数为空则默认等于0  不为空则跳过
            If ufgData.Text(Row, gstrMaterial_蜡块数) = "" Then
                ufgData.Text(Row, gstrMaterial_蜡块数) = "0"
            End If
        
            If Val(ufgData.Text(Row, gstrMaterial_蜡块数)) < 1 And ufgData.Text(Row, gstrMaterial_是否蜡块) = "1-是" Then
                ufgData.Text(Row, gstrMaterial_蜡块数) = "1"
            End If
            
            '如果界面中显示了蜡块数，则蜡块数必须录入，否则修改单元格颜色
            If Not ufgData.DataGrid.ColHidden(ufgData.GetColIndex(gstrMaterial_蜡块数)) Then
                If InStr(ufgData.Text(Row, gstrMaterial_是否冰余), "是") > 0 Then
                
                    '如果未录入蜡块数量，则显示淡红色(冰冻检查也需要录入蜡块数，确定制片数量)
                    iCol = ufgData.GetColIndex(gstrMaterial_蜡块数)
                    
                    ufgData.CellColor(Row, iCol) = IIf(Val(ufgData.Text(Row, gstrMaterial_蜡块数)) <= 0, ufgData.ErrCellColor, ufgData.BackColor)
                Else
                    iCol = ufgData.GetColIndex(gstrMaterial_蜡块数)
                    
                    ufgData.CellColor(Row, iCol) = ufgData.BackColor
                End If
            End If
            
            
            '如果界面中显示了冰余，则冰余必须录入，否则修改单元格颜色
            If Not ufgData.DataGrid.ColHidden(ufgData.GetColIndex(gstrMaterial_是否冰余)) Then
                '如果未录入是否冰余，则显示淡红色
                iCol = ufgData.GetColIndex(gstrMaterial_是否冰余)
            
                ufgData.CellColor(Row, iCol) = IIf(ufgData.Text(Row, gstrMaterial_是否冰余) = "", ufgData.ErrCellColor, ufgData.BackColor)
            End If
        
        Case StudyType.stCell  '细胞检查
        
            '如果细胞块数为空则默认等于0  不为空则跳过
            If ufgData.Text(Row, gstrMaterial_细胞块数) = "" Then
                ufgData.Text(Row, gstrMaterial_细胞块数) = "0"
            End If
            
            If Val(ufgData.Text(Row, gstrMaterial_细胞块数)) < 1 And ufgData.Text(Row, gstrMaterial_是否蜡块) = "1-是" Then
                ufgData.Text(Row, gstrMaterial_细胞块数) = "1"
            End If
            
            
             '如果标本量为空则执行读取数据  不为空则跳过
            If ufgData.Text(Row, gstrMaterial_标本量) = "" Then
                If rsInspectionSpe.RecordCount > 0 Then
                    rsInspectionSpe.MoveFirst
                    Do While Not rsInspectionSpe.EOF
                        '如果显示的标本名称与 病理检查标本中的匹配则读取默认制片数
                        If ufgData.DataGrid.Cell(flexcpTextDisplay, Row, ufgData.GetColIndex(gstrMaterial_标本名称)) = rsInspectionSpe("标本名称").value Then
                            ufgData.Text(Row, gstrMaterial_标本量) = rsInspectionSpe("默认标本量").value
                            Exit Do
                        End If
                        rsInspectionSpe.MoveNext
                    Loop
                End If
            End If
            
        
            '如果未录入标本量，则显示淡红色
            iCol = ufgData.GetColIndex(gstrMaterial_标本量)
            
            ufgData.CellColor(Row, iCol) = IIf(Val(ufgData.Text(Row, gstrMaterial_标本量)) <= 0, ufgData.ErrCellColor, ufgData.BackColor)
                    
    End Select
    
    
    '如果未录入标本名称，则显示淡红色
    iCol = ufgData.GetColIndex(gstrMaterial_标本名称)

    ufgData.CellColor(Row, iCol) = IIf(ufgData.Text(Row, gstrMaterial_标本名称) = "", ufgData.ErrCellColor, ufgData.BackColor)
                 
    
    '如果未录入主取医师，则显示淡红色
    iCol = ufgData.GetColIndex(gstrMaterial_主取医师)
    
    ufgData.CellColor(Row, iCol) = IIf(ufgData.Text(Row, gstrMaterial_主取医师) = "", ufgData.ErrCellColor, ufgData.BackColor)
    
    
    '如果未录入主取医师，则显示淡红色
    iCol = ufgData.GetColIndex(gstrMaterial_取材时间)
    
    ufgData.CellColor(Row, iCol) = IIf(Not IsDate(ufgData.Text(Row, gstrMaterial_取材时间)), ufgData.ErrCellColor, ufgData.BackColor)

    '如果为会诊，则制片数不允许编辑
    If mrecStudy.lngStudyType = stMeet Then
        ufgData.Text(Row, gstrMaterial_制片数) = ""
    End If

       
End Sub

Private Sub ufgData_OnStartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim dtServicesTime As Date
    Dim strComboboxText As String
    Dim strSql As String
    Dim rsInspectionSpe As ADODB.Recordset

    
    '判断是否允许进行新的取材
    If ufgData.IsNullRow(Row) Then
        If mrecStudy.lngMaterialStep <> TExecuteStep.NeedDo Then
            Cancel = True
            Call MsgBoxD(Me, "非取材阶段，不能进行新的取材操作，请申请。", vbOKOnly, Me.Caption)
        
            Exit Sub
        End If
    End If

    
    '判断是否允许更新
    If Not CheckAllowUpdate(ufgData.KeyValue(Row)) Then
        Cancel = True
        Call MsgBoxD(Me, "该材块记录已执行制片处理，不能进行更新。", vbOKOnly, Me.Caption)
        
        Exit Sub
    End If
    
    
    If Not IsDate(ufgData.Text(Row, gstrMaterial_取材时间)) Then
        ufgData.Text(Row, gstrMaterial_取材时间) = zlDatabase.Currentdate
    End If
    
    
    If Row > 0 Then
    
            '加载标本名称
            'If Col = ufgData.vfgHelper.GetColumnIndex(gstrMaterial_标本名称) Then
                If ufgData.Text(Row, gstrMaterial_标本名称) = "" Then
                    If Row - 1 > 0 Then
                        If ufgData.Text(Row - 1, gstrMaterial_标本名称) <> "" Then
                            ufgData.Text(Row, gstrMaterial_标本名称) = ufgData.Text(Row - 1, gstrMaterial_标本名称)
                        End If
                    End If
                    
                    If ufgData.Text(Row, gstrMaterial_标本名称) = "" Then
                        strComboboxText = ufgData.DataGrid.ColComboList(ufgData.GetColIndex(gstrMaterial_标本名称))
                        
                        If strComboboxText <> "" Then
                            If InStr(strComboboxText, ";") > 0 Then
                                strComboboxText = Mid(strComboboxText, 1, InStr(strComboboxText, ";") - 1)
                            End If
                            ufgData.Text(Row, gstrMaterial_标本名称) = Mid(strComboboxText, InStr(strComboboxText, "#") + 1, 255)
                            
                        End If
                    End If
                End If
    
    
        If mrecStudy.lngStudyType = StudyType.stIce Then
        
            If Not ufgData.DataGrid.ColHidden(ufgData.GetColIndex(gstrMaterial_是否冰余)) Then
                If ufgData.Text(Row, gstrMaterial_是否冰余) = "" Then ufgData.Text(Row, gstrMaterial_是否冰余) = "0-否"
            End If
            
            
            
            '如果不是冰余，蜡块数则不允许进行编辑
            'If Col = ufgData.vfgHelper.GetColumnIndex(gstrMaterial_蜡块数) Then
            If Not ufgData.DataGrid.ColHidden(ufgData.GetColIndex(gstrMaterial_蜡块数)) Then
                If InStr(ufgData.Text(Row, gstrMaterial_是否冰余), "是") > 0 Then
                    If Val(ufgData.Text(Row, gstrMaterial_蜡块数)) <= 0 Then ufgData.Text(Row, gstrMaterial_蜡块数) = "1"
                Else
                    If Col = ufgData.GetColIndex(gstrMaterial_蜡块数) Then Cancel = True
                End If
            End If
            'End If
            If ufgData.Text(Row, gstrMaterial_是否蜡块) = "" Then ufgData.Text(Row, gstrMaterial_是否蜡块) = "0-否"
            
        ElseIf mrecStudy.lngStudyType = StudyType.stCell Then
            'If Col = ufgData.vfgHelper.GetColumnIndex(gstrMaterial_细胞块数) Then
            'ufgData.Text(Row, gstrMaterial_细胞块数) = "0"
            
            If UCase(ufgData.DisplayText(Row, gstrMaterial_标本名称)) = UCase("TCT") Then
                ufgData.Text(Row, gstrMaterial_标本量) = "20ml"
            ElseIf ufgData.DisplayText(Row, gstrMaterial_标本名称) Like "*痰*" Then
                ufgData.Text(Row, gstrMaterial_标本量) = "1.5ml"
            End If
            
            If ufgData.Text(Row, gstrMaterial_是否蜡块) = "" Then ufgData.Text(Row, gstrMaterial_是否蜡块) = "0-否"
        Else
            '设置蜡块数
            'If Col = ufgData.vfgHelper.GetColumnIndex(gstrMaterial_蜡块数) Then
                If Not ufgData.DataGrid.ColHidden(ufgData.GetColIndex(gstrMaterial_蜡块数)) Then
                    If Val(ufgData.Text(Row, gstrMaterial_蜡块数)) <= 0 Then ufgData.Text(Row, gstrMaterial_蜡块数) = "1"
                End If
                
                If ufgData.Text(Row, gstrMaterial_是否蜡块) = "" Then ufgData.Text(Row, gstrMaterial_是否蜡块) = "1-是"
       
            
            'End If
        End If
        

        
        'If Col = ufgData.vfgHelper.GetColumnIndex(gstrMaterial_主取医师) Then
            If ufgData.Text(Row, gstrMaterial_主取医师) = "" Then
                If Row - 1 > 0 Then
                    If ufgData.Text(Row - 1, gstrMaterial_主取医师) <> "" Then
                        ufgData.Text(Row, gstrMaterial_主取医师) = ufgData.Text(Row - 1, gstrMaterial_主取医师)
                    End If
                End If
                
                If ufgData.Text(Row, gstrMaterial_主取医师) = "" Then
                    strComboboxText = ufgData.DataGrid.ColComboList(ufgData.GetColIndex(gstrMaterial_主取医师))
                    
                    If strComboboxText <> "" Then
                        If InStr(strComboboxText, "|") > 0 Then
                            strComboboxText = Mid(strComboboxText, 1, InStr(strComboboxText, "|") - 1)
                        End If
                        ufgData.Text(Row, gstrMaterial_主取医师) = strComboboxText
                        
                    End If
                End If
            End If
        'End If

         
    End If
End Sub

Private Sub ufgDecalin_OnAfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim iCol As Long
    Dim i As Long
    Dim iCount As Long
    Dim strNewSpecimenName As String

    If ufgDecalin.IsNullRow(Row) Then
        ufgDecalin.RowState(Row) = TDataRowState.Normal
        Call ufgDecalin.SetRowColor(Row, ufgDecalin.BackColor)
        
        Exit Sub
    End If



    '如果未录入标本名称，则显示淡红色
    iCol = ufgDecalin.GetColIndex(gstrDecalin_标本名称)
    
    ufgDecalin.CellColor(Row, iCol) = IIf(ufgDecalin.Text(Row, gstrDecalin_标本名称) = "", ufgDecalin.ErrCellColor, ufgDecalin.BackColor)
           


    '如果未录入开始时间，则显示淡红色
    iCol = ufgDecalin.GetColIndex(gstrDecalin_开始时间)
    
    ufgDecalin.CellColor(Row, iCol) = IIf(Not IsDate(ufgDecalin.Text(Row, gstrDecalin_开始时间)), ufgDecalin.ErrCellColor, ufgDecalin.BackColor)
          
    

    '如果未录入所需时长，则显示淡红色
    iCol = ufgDecalin.GetColIndex(gstrDecalin_所需时长)
    
    ufgDecalin.CellColor(Row, iCol) = IIf(Val(ufgDecalin.Text(Row, gstrDecalin_所需时长)) <= 0, ufgDecalin.ErrCellColor, ufgDecalin.BackColor)
    
    
    '如果未录入操作员，则显示淡红色
    iCol = ufgDecalin.GetColIndex(gstrDecalin_操作员)
    
    ufgDecalin.CellColor(Row, iCol) = IIf(ufgDecalin.Text(Row, gstrDecalin_操作员) = "", ufgDecalin.ErrCellColor, ufgDecalin.BackColor)
    
End Sub


Private Sub ConfigDecalcificationBut()
'配置脱钙按钮
    If Not ufgDecalin.IsSelectionRow Then
        cmdDecalin.Enabled = False
        cmdChange.Enabled = False
        cmdCancel.Enabled = False
        cmdSucceed.Enabled = False
    End If
    
    cmdDecalin.Enabled = Not ufgDecalin.IsNullRow(ufgDecalin.SelectionRow) And ufgDecalin.IsEmptyKey(ufgDecalin.SelectionRow)
    cmdChange.Enabled = Not ufgDecalin.IsEmptyKey(ufgDecalin.SelectionRow)
    cmdCancel.Enabled = Not ufgDecalin.IsNullRow(ufgDecalin.SelectionRow)
    cmdSucceed.Enabled = Not ufgDecalin.IsEmptyKey(ufgDecalin.SelectionRow)
End Sub



Private Sub ufgDecalin_OnClick()
On Error GoTo errHandle
    Call ConfigDecalcificationBut
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub






Private Sub ufgDecalin_OnStartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim dtServicesTime As Date
    
    If Col = ufgDecalin.GetColIndex(gstrDecalin_开始时间) And Row > 0 Then
        
        dtServicesTime = zlDatabase.Currentdate
        ufgDecalin.Text(Row, gstrDecalin_开始时间) = dtServicesTime
        
        Exit Sub
    End If
End Sub


