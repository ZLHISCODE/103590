VERSION 5.00
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmPatholPrice 
   Caption         =   "检查补费"
   ClientHeight    =   7080
   ClientLeft      =   75
   ClientTop       =   405
   ClientWidth     =   9195
   Icon            =   "frmPatholPrice.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7080
   ScaleWidth      =   9195
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picRequest 
      BorderStyle     =   0  'None
      Height          =   6855
      Left            =   120
      ScaleHeight     =   6855
      ScaleWidth      =   5175
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      Begin VB.Frame framDetails 
         Caption         =   "申请明细"
         Height          =   3015
         Left            =   120
         TabIndex        =   6
         Top             =   3240
         Width           =   4575
         Begin zl9PACSWork.ucFlexGrid ufgContext 
            Height          =   2535
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   4471
            DefaultCols     =   ""
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
      End
      Begin VB.Frame framRequest 
         Caption         =   "申请记录"
         Height          =   3015
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   4575
         Begin zl9PACSWork.ucFlexGrid ufgRequest 
            Height          =   2655
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   4335
            _ExtentX        =   16325
            _ExtentY        =   3413
            DefaultCols     =   ""
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
      Begin VB.PictureBox picControl 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   -240
         ScaleHeight     =   495
         ScaleWidth      =   4935
         TabIndex        =   1
         Top             =   6360
         Width           =   4935
         Begin VB.CommandButton cmdAlreadyPrice 
            Caption         =   "完成补费(&F)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   240
            TabIndex        =   9
            Top             =   0
            Width           =   4575
         End
         Begin VB.CommandButton cmdTempPrice 
            Caption         =   "零 耗 (&T)"
            Enabled         =   0   'False
            Height          =   400
            Left            =   240
            TabIndex        =   4
            Top             =   0
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CommandButton cmdBill 
            Caption         =   "记 账(&M)"
            Enabled         =   0   'False
            Height          =   400
            Left            =   240
            TabIndex        =   3
            Top             =   0
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CommandButton cmdAccept 
            Caption         =   "收 费(&R)"
            Enabled         =   0   'False
            Height          =   400
            Left            =   120
            TabIndex        =   2
            Top             =   0
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CheckBox chkAutoExecute 
            Caption         =   "补费后费用自动执行"
            Height          =   255
            Left            =   2760
            TabIndex        =   10
            Top             =   120
            Visible         =   0   'False
            Width           =   1935
         End
      End
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   5400
      Top             =   240
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmPatholPrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private mlngModule As Long
Private mstrPrivs As String
Private mlngCurDepartmentId As Long
Private mobjOwner As Object

Private mlngCurAdviceId As Long
Private mlngSendNo As Long
Private mblnMoved As Boolean
Private mblnReadOnly As Boolean


Private mlngRequestType As Long

Private mlngCurRequestId As Long
Private mblnButtonEvent As Boolean

Private mrecStudyInf As TStudyStateInf

Private mobjExpense As zlPublicExpense.clsDockExpense        '费用对象


Public Sub zlInitModule(ByVal lngModule As Long, ByVal strPrivs As String, ByVal lngDepartId As Long, Optional owner As Object = Nothing)
'初始化模块参数
    mlngModule = lngModule
    mstrPrivs = strPrivs
    mlngCurDepartmentId = lngDepartId
    
    If Not owner Is Nothing Then Set mobjOwner = owner
End Sub

Public Sub zlRefresh(ByVal lngCurDepartmentId As Long, lngAdviceID As Long, ByVal lngSendNO As Long, ByVal blnMoved As Boolean)
    
On Error GoTo errHandle
    If lngAdviceID <= 0 Then
        Call ConfigPriceFace(False, "医嘱ID无效请检查。")
        Exit Sub
    End If
    
    mlngCurAdviceId = lngAdviceID
    mblnMoved = blnMoved
    mlngSendNo = lngSendNO
    mlngCurDepartmentId = lngCurDepartmentId
    mblnReadOnly = blnMoved
    
    
    Call GetPatholStudyState(lngAdviceID, mrecStudyInf)
    
    
    If mrecStudyInf.strPatholNumber = "" Then
        Call RefreshPrice(lngCurDepartmentId, lngAdviceID, lngSendNO, blnMoved)
        
        Call ConfigPriceFace(False, "该检查尚未生成有效的病理号，请确认该检查是否已被核收。")
        
'        If Not (mobjOwner Is Nothing) Then
'            Call MsgBoxD(Me, "该检查尚未生成有效的病理号，请确认该检查是否已被核收。", vbOKOnly, Me.Caption)
'        End If
        
        Exit Sub
    End If
    
    
    '读取申请信息
    Call LoadRequestInf(mrecStudyInf.lngPatholAdviceId)
    
    '载入申请明细
    Call ufgRequest_OnClick
    
    Call ConfigPriceFace(True)

    
    Call ConfigPopedom(mblnReadOnly)
    
'    '配置病人来源界面
'    Call ConfigPatientSource(lngAdviceID)
    
    Call RefreshPrice(lngCurDepartmentId, lngAdviceID, lngSendNO, blnMoved)
    
    If ufgRequest.ShowingDataRowCount > 0 Then
        Call ufgRequest.LocateRow(1)
        Call ConfigPriceState(ufgRequest.Text(1, gstrRequisition_补费状态) = "需补费")
    Else
        Call ConfigPriceState(False)
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub ConfigPatientSource(ByVal lngAdviceID As Long)
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    strSql = "select 病人来源 from 病人医嘱记录 where ID=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngAdviceID)
    
    If rsData.RecordCount <= 0 Then Exit Sub
    
    If Nvl(rsData!病人来源) = 2 Then
        cmdAccept.Enabled = False
    End If

End Sub

Private Sub LoadRequestInf(ByVal lngPatholAdviceId As Long)
'载入申请信息
    Dim strSql As String
    
    strSql = "select 申请ID,申请人,申请类型,补费状态,申请细目,申请时间,申请状态,申请描述,完成时间 " & _
        " from 病理申请信息 where 病理医嘱ID=[1] order by 申请类型,申请时间"
'    If mblnMoved Then strSql = GetMovedDataSql(strSql)
    
    Set ufgRequest.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPatholAdviceId)
    
    Call ufgRequest.RefreshData
End Sub



Private Sub ConfigPriceFace(ByVal blnIsValid As Boolean, Optional ByVal strHintInf As String = "")
'配置补费界面

    cmdAccept.Enabled = blnIsValid
    cmdBill.Enabled = blnIsValid
    cmdTempPrice.Enabled = blnIsValid
    cmdAlreadyPrice.Enabled = blnIsValid
    
    If blnIsValid Then
        Call ufgRequest.CloseHintInf
        Call ufgContext.CloseHintInf
    Else
        Call ufgRequest.ShowHintInf(strHintInf)
        Call ufgContext.ShowHintInf(strHintInf)
    End If
End Sub


Private Sub ConfigPopedom(ByVal blnIsReadOnly As Boolean)
'配置权限
    Dim blnIsAllowPrice As Boolean
    
    blnIsAllowPrice = IIf(GetInsidePrivs(p医嘱附费管理, True) <> "", True, False)
    
    
    cmdAccept.Enabled = blnIsAllowPrice And Not blnIsReadOnly
    cmdBill.Enabled = blnIsAllowPrice And Not blnIsReadOnly
    cmdTempPrice.Enabled = blnIsAllowPrice And Not blnIsReadOnly
    
    
    ufgRequest.ReadOnly = blnIsReadOnly
    ufgContext.ReadOnly = blnIsReadOnly
End Sub


Private Sub InitFace()
'初始化界面布局
    Dim Pane1 As Pane, Pane2 As Pane

    If Not mobjExpense Is Nothing Then
        With dkpMain
            .CloseAll
            .Options.HideClient = True
            .Options.UseSplitterTracker = False '实时拖动
            .Options.ThemedFloatingFrames = True
            .Options.AlphaDockingContext = True
        End With
    
        Set Pane1 = dkpMain.CreatePane(1, 0, Round(Me.Height / 2), DockLeftOf, Nothing)
        Pane1.Title = "申请记录"
        Pane1.Handle = picRequest.hWnd
        Pane1.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        Pane1.MinTrackSize.Width = 50
        Pane1.MinTrackSize.Height = 50
    
        Set Pane2 = dkpMain.CreatePane(2, 0, Round(Me.Height / 2), DockRightOf, Pane1)
        Pane2.Title = "费用记录"
        Pane2.Handle = mobjExpense.zlGetForm.hWnd
        Pane2.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        Pane2.MinTrackSize.Width = 50
        Pane2.MinTrackSize.Height = 50
    Else
        picRequest.Width = Me.ScaleWidth - 240
        picRequest.Height = Me.ScaleHeight
    End If
End Sub

Private Sub cmdAccept_Click()
'On Error GoTo errHandle
'
'    mblnButtonEvent = True
'
'
'    '收费单据
'    Call mobjExpense.zlExecuteCommandBars1(1)
'
'    mblnButtonEvent = False
'Exit Sub
'errHandle:
'    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdAlreadyPrice_Click()
On Error GoTo errHandle
    Dim i As Long
    Dim strSql As String
    Dim lngRequestId As String
    
    If ufgRequest.ShowingDataRowCount <= 0 Then
        Call MsgBoxD(Me, "没有需要执行补费的记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    
    '执行补费
    For i = 1 To ufgRequest.GridRows - 1
        If ufgRequest.Text(ufgRequest.SelectionRow, gstrRequisition_补费状态) = "需补费" Then
            lngRequestId = Val(ufgRequest.KeyValue(ufgRequest.SelectionRow))
    
            strSql = "zl_病理申请_补费(" & lngRequestId & ")"
            Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
        
            '更新申请列表的补费状态
            ufgRequest.Text(ufgRequest.SelectionRow, gstrRequisition_补费状态) = "已补费"
        End If
    Next i

    
    '更新费用状态为正在执行
    If chkAutoExecute.value <> 0 Then
        Call ExecuteStudyMoney
    End If

        
    Call MsgBoxD(Me, "已完成补费操作。", vbOKOnly, Me.Caption)
    
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub



'费用执行
Private Sub ExecuteStudyMoney()
    On Error GoTo errHandle
      
    
    gstrSQL = "Zl_影像费用执行(" & mlngCurAdviceId & "," & mlngSendNo & ",2,Null,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & mlngCurDepartmentId & ")"
    zlDatabase.ExecuteProcedure gstrSQL, "费用执行"
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub



Private Sub cmdBill_Click()
'On Error GoTo errHandle
'    mblnButtonEvent = True
'
'    '记账单据
'    Call mobjExpense.zlExecuteCommandBars1(2)
'
'    mblnButtonEvent = False
'Exit Sub
'errHandle:
'    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdTempPrice_Click()
'On Error GoTo errHandle
'    mblnButtonEvent = True
'
'    '零耗费用
'    Call mobjExpense.zlExecuteCommandBars1(3)
'
'    mblnButtonEvent = False
'Exit Sub
'errHandle:
'    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Load()
    Dim objTmp As zlPublicExpense.clsPublicExpense
On Error GoTo errHandle
    Call RestoreWinState(Me, App.ProductName)
    
    mblnButtonEvent = False
    mlngCurRequestId = -1
    
    If mlngModule > -1 Then
        Set objTmp = New zlPublicExpense.clsPublicExpense
        Call objTmp.zlInitCommon(glngSys, gcnOracle, gstrDBUser)
        Set mobjExpense = objTmp.zlDockExpense
    End If
    
    Call InitFace
    
        '初始化申请列表
    Call InitRequisitionList
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Public Sub SetFontSize(ByVal bytFontSize As Byte, ByVal bytSize As Byte)
On Error GoTo errHandle

    Call ReSetFormFontSize(bytFontSize)
    
    If Not mobjExpense Is Nothing Then
        Call mobjExpense.SetFontSize(bytSize)
    End If
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Public Sub ReSetFormFontSize(ByVal bytFontSize As Byte)
'功能:重新设置工作站窗体的字体大小
    
    Dim objCtrl As Control
    Dim CtlFont As StdFont
    
    Me.FontSize = bytFontSize
    For Each objCtrl In Me.Controls
        Select Case UCase(TypeName(objCtrl))
        Case UCase("Frame")
            objCtrl.Font.Size = bytFontSize
        Case UCase("Label")
            objCtrl.FontSize = bytFontSize
            objCtrl.Height = TextHeight("罗") + 20
        Case UCase("ucFlexGrid")
            objCtrl.DataGrid.Cell(flexcpFontSize, 0, 0, 0, objCtrl.DataGrid.Cols - 1) = bytFontSize
            objCtrl.DataGrid.FontSize = bytFontSize
        Case UCase("CheckBox")
            objCtrl.FontSize = bytFontSize
            objCtrl.Width = TextWidth("罗冠" & objCtrl.Caption)
        Case UCase("DockingPane")
            Set CtlFont = objCtrl.PaintManager.CaptionFont
            If CtlFont Is Nothing Then
                Set CtlFont = Me.Font
            End If
            CtlFont.Size = bytFontSize
            Set objCtrl.PaintManager.CaptionFont = CtlFont
        Case UCase("CommandButton")
            objCtrl.FontSize = bytFontSize
        End Select
    Next
    
End Sub

Private Sub InitRequisitionList()
'初始化申请列表
    Dim strTemp As String
    

    
    '判断数据库参数表是否有数据 有则读取数据库参数  没有则加载默认
    strTemp = zlDatabase.GetPara("检查申请列表配置", glngSys, G_LNG_PATHOLSYS_NUM, "")
    ufgRequest.DefaultColNames = gstrRequisitionCols
     
    If strTemp = "" Then
        ufgRequest.ColNames = gstrRequisitionCols
    Else
        ufgRequest.ColNames = strTemp
    End If
        '设置行数
    ufgRequest.GridRows = glngStandardRowCount
    '设置行高
    ufgRequest.RowHeightMin = glngStandardRowHeight
    ufgRequest.ColConvertFormat = gstrRequisitionConvertFormat
End Sub

Private Sub InitRequestContextList(ByVal lngRequestType As Long)
'初始化申请项目明细列表
    Dim strTemp As String
    

    
    mlngRequestType = lngRequestType
    
    Select Case lngRequestType
        Case 0, 1, 2
        
            '判断数据库参数表是否有数据 有则读取数据库参数  没有则加载默认
            strTemp = zlDatabase.GetPara("特检申请列表配置", glngSys, G_LNG_PATHOLSYS_NUM, "")
            ufgContext.DefaultColNames = gstrRequest_SpeExam_Cols
                        
            If strTemp = "" Then
                ufgContext.ColNames = gstrRequest_SpeExam_Cols
            Else
                ufgContext.ColNames = strTemp
            End If
            
            ufgContext.ColConvertFormat = gstrRequest_SpeExamConvertFormat
            
        Case 3
            
            '判断数据库参数表是否有数据 有则读取数据库参数  没有则加载默认
            strTemp = zlDatabase.GetPara("制片申请列表配置", glngSys, G_LNG_PATHOLSYS_NUM, "")
            ufgContext.DefaultColNames = gstrRequest_Slices_Cols
             
            If strTemp = "" Then
                ufgContext.ColNames = gstrRequest_Slices_Cols
            Else
                ufgContext.ColNames = strTemp
            End If
            
            ufgContext.ColConvertFormat = gstrRequest_SlicesConvertFormat
        Case 4, 5
            
            '判断数据库参数表是否有数据 有则读取数据库参数  没有则加载默认
            strTemp = zlDatabase.GetPara("补取申请列表配置", glngSys, G_LNG_PATHOLSYS_NUM, "")
            ufgContext.DefaultColNames = gstrRequest_Material_Cols
             
            If strTemp = "" Then
                ufgContext.ColNames = gstrRequest_Material_Cols
            Else
                ufgContext.ColNames = strTemp
            End If
            
            ufgContext.ColConvertFormat = gstrRequest_MaterialConvertFormat
    End Select
        '设置行数
    ufgContext.GridRows = glngStandardRowCount
    '设置行高
    ufgContext.RowHeightMin = glngStandardRowHeight
End Sub

Private Sub Form_Resize()
On Error Resume Next
    If mlngModule = -1 Then
        picRequest.Width = Me.ScaleWidth - 240
        picRequest.Height = Me.ScaleHeight
    End If
err.Clear
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


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
    
    If Not mobjExpense Is Nothing Then
        Unload mobjExpense.zlGetForm
    End If
    
    Set mobjExpense = Nothing
End Sub



'Private Sub mobjExpense_OnPriceEvent(ByVal lngAdviceID As Long, ByVal lngPriceType As Long)
'    Dim strSQL As String
'
'    If mblnButtonEvent And mlngCurRequestId > 0 Then
'        strSQL = "zl_病理申请_补费(" & mlngCurRequestId & ")"
'        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
'
'        '更新申请列表的补费状态
'        ufgRequest.Text(ufgRequest.SelectRowIndex, gstrRequisition_补费状态) = "已补费"
'
'        Call ConfigPriceState(ufgRequest.Text(ufgRequest.SelectRowIndex, gstrRequisition_补费状态) = "需补费")
'    End If
'End Sub

Private Sub picRequest_Resize()
On Error Resume Next
     Call AdjustFace
End Sub



Private Sub AdjustFace()
    Dim lngAvgHeight As Long
    
    lngAvgHeight = Fix((picRequest.Height - picControl.Height) / 2)
    
    framRequest.Left = 0
    framRequest.Top = 0
    framRequest.Width = picRequest.Width
    framRequest.Height = lngAvgHeight - 120
    
    ufgRequest.Left = 120
    ufgRequest.Top = 240
    ufgRequest.Width = framRequest.Width - 240
    ufgRequest.Height = framRequest.Height - 360
    
    
    framDetails.Left = 0
    framDetails.Top = framRequest.Top + framRequest.Height + 60
    framDetails.Width = picRequest.Width
    framDetails.Height = lngAvgHeight - 120
    
    
    ufgContext.Left = 120
    ufgContext.Top = 240
    ufgContext.Width = framDetails.Width - 240
    ufgContext.Height = framDetails.Height - 360
    
    
    picControl.Left = 0
    picControl.Top = framDetails.Top + framDetails.Height + 60
    picControl.Width = framRequest.Width
    
    
    cmdAlreadyPrice.Left = 0 'picControl.Width - cmdAlreadyPrice.Width
    cmdAlreadyPrice.Top = 60
    cmdAlreadyPrice.Width = framDetails.Width
    
'    chkAutoExecute.Left = 0 'cmdAlreadyPrice.Left + cmdAlreadyPrice.Width + 120
'    chkAutoExecute.Top = cmdAlreadyPrice.Top
    
    
'    cmdTempPrice.Left = picControl.Width - cmdTempPrice.Width
'    cmdTempPrice.Top = 0
'
'    cmdBill.Left = cmdTempPrice.Left - cmdBill.Width - 120
'    cmdBill.Top = 0
'
'    cmdAccept.Left = cmdBill.Left - cmdAccept.Width - 120
'    cmdAccept.Top = 0
    

    
End Sub

Private Sub ConfigPriceState(ByVal blnIsPrice As Boolean)
'配置补费按钮状态
    cmdAccept.Enabled = blnIsPrice
    cmdBill.Enabled = blnIsPrice
    cmdTempPrice.Enabled = blnIsPrice
    cmdAlreadyPrice.Enabled = blnIsPrice
End Sub



Private Sub ufgRequest_OnClick()
'读取申请内容
On Error GoTo errHandle
    Dim strRequestType As String
    
    mlngCurRequestId = -1
    
    '清除申请项目明细
    Call ufgContext.ClearListData
    Call ConfigPriceState(False)
    
    If Not ufgRequest.IsSelectionRow Then Exit Sub
    If ufgRequest.IsEmptyKey(ufgRequest.SelectionRow) Then Exit Sub
    
    Call ConfigPriceState(ufgRequest.Text(ufgRequest.SelectionRow, gstrRequisition_补费状态) = "需补费")
    
    strRequestType = ufgRequest.Text(ufgRequest.SelectionRow, gstrRequisition_申请类型)
    mlngCurRequestId = Val(ufgRequest.KeyValue(ufgRequest.SelectionRow))
    
    Select Case strRequestType
        Case "免疫组化", "分子病理", "特殊染色"
        
            Call InitRequestContextList(0)
            
            '读取特检项目明细
            Call LoadSpeExamRequestContext(ufgRequest.KeyValue(ufgRequest.SelectionRow))
            
        Case "再制片", "重切", "深切", "连切", "白片"
            
            Call InitRequestContextList(3)
             
            '读取制片项目明细
            Call LoadSlicesRequestContext(ufgRequest.KeyValue(ufgRequest.SelectionRow))

        Case "重取材", "补取材"
            
            Call InitRequestContextList(4)
            
            '读取取材项目明细
            Call LoadSupMaterialRequestContext(ufgRequest.KeyValue(ufgRequest.SelectionRow))
            
    End Select
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
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


Private Sub LoadSupMaterialRequestContext(ByVal lngRequestId As Long)
'读取取材的完成内容
    Dim strSql As String
    
    strSql = "select 材块ID,序号,标本名称,标本量,蜡块数,取材时间,主取医师,副取医师,记录医师 " & _
            " from 病理取材信息 where  申请id=[1] order by 取材时间"
'    If mblnMoved Then strSql = GetMovedDataSql(strSql)
    
    Set ufgContext.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngRequestId)
    
    Call ufgContext.RefreshData
End Sub







Public Sub zlExecuteCommandBars(Control As Object)
    If mobjExpense Is Nothing Then Exit Sub
    
    Call mobjExpense.zlExecuteCommandBars(Control)
End Sub

Public Sub zlDefCommandBars(frmParent As Object, CommandBars As Object)
    If mobjExpense Is Nothing Then Exit Sub
    
    Call mobjExpense.zlDefCommandBars(frmParent, CommandBars)
End Sub


Public Sub zlPopupCommandBars(CommandBar As Object)
    If mobjExpense Is Nothing Then Exit Sub
    
    Call mobjExpense.zlPopupCommandBars(CommandBar)
End Sub

Public Sub zlUpdateCommandBars(Control As Object)
    If mobjExpense Is Nothing Then Exit Sub
    
    Call mobjExpense.zlUpdateCommandBars(Control)
End Sub

Public Function zlGetForm() As Object
    Set zlGetForm = Me
End Function

Private Sub RefreshPrice(lng科室ID As Long, lng医嘱ID As Long, lng发送号 As Long, _
    Optional ByVal blnMoved As Boolean, Optional ByVal bln单独执行 As Boolean)
    If mobjExpense Is Nothing Then Exit Sub
    
'    Call mobjExpense.zlRefresh(lng科室ID, lng医嘱ID, lng发送号, blnMoved, bln单独执行)
    Call mobjExpense.zlRefresh(lng科室ID, lng医嘱ID & ":" & lng发送号 & ":" & IIf(bln单独执行 = True, 1, 0), blnMoved)
End Sub

