VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmEventMsgEditSegment 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4875
   Icon            =   "frmEventMsgEditSegment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   4875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtLocation 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   1470
      TabIndex        =   8
      Top             =   75
      Width           =   1575
   End
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   2190
      Index           =   0
      Left            =   0
      ScaleHeight     =   2190
      ScaleWidth      =   4710
      TabIndex        =   0
      Top             =   405
      Width           =   4710
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   1800
         TabIndex        =   7
         Top             =   1680
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(C)"
         Height          =   350
         Left            =   3135
         TabIndex        =   6
         Top             =   1725
         Width           =   1100
      End
      Begin VB.Frame fra 
         Height          =   1530
         Left            =   45
         TabIndex        =   1
         Top             =   -90
         Width           =   4665
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   0
            Left            =   960
            TabIndex        =   3
            Top             =   285
            Width           =   2790
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   0
            Left            =   960
            TabIndex        =   2
            Top             =   690
            Width           =   2790
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "段名"
            Height          =   180
            Index           =   0
            Left            =   150
            TabIndex        =   5
            Top             =   360
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "重复"
            Height          =   180
            Index           =   1
            Left            =   150
            TabIndex        =   4
            Top             =   795
            Width           =   360
         End
      End
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmEventMsgEditSegment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mfrmParent As Object
Private mbytMode As Byte
Private mblnDataChanged As Boolean
Private mblnReading As Boolean
Private mobjFindKey As CommandBarControl
Private mstrFindKey As String
Private mrsPara As ADODB.Recordset
Private mstrDataKey As String
Private mlngModualCode As Long
Private mblnContiune As Boolean

Public Event BeforeNewData(ByVal DataKey As String)
Public Event AfterNewData(ByVal DataKey As String)
Public Event AfterModifyData(ByVal DataKey As String)
Public Event AfterDeleteData(ByVal DataKey As String)
Public Event Forward(ByRef DataKey As String, ByRef Cancel As Boolean)
Public Event Backward(ByRef DataKey As String, ByRef Cancel As Boolean)

'######################################################################################################################

Public Function InitDialog(ByVal frmParent As Object) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Set mfrmParent = frmParent
    
    InitDialog = True
    
End Function

Public Sub NewData()
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mbytMode = 1
    Me.Caption = "新增消息段"
    mstrDataKey = ""
    
    Call InitData
    Call InitCommandBar
    
    txt(1).Text = gclsBusiness.GetMaxCode("m_Event", "code")
    
    mblnDataChanged = False
        
    Me.Show 1, mfrmParent
    
End Sub

Public Sub ModifyData(ByVal strDataKey As String)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************

    
    mbytMode = 2
    mstrDataKey = strDataKey
    
    Me.Caption = "修改消息段"
    
    Call InitData
    Call InitCommandBar
    
    Call ReadData(mstrDataKey)
    
    Me.Show 1, mfrmParent
    
End Sub

'######################################################################################################################
Private Function InitData() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rsTmp As ADODB.Recordset
    
    mblnContiune = False
    
    Set rsTmp = gclsBusiness.EventStruct()
    If Not (rsTmp Is Nothing) Then
        txt(0).MaxLength = rsTmp("title").DefinedSize
        txt(1).MaxLength = rsTmp("code").Precision
        txt(2).MaxLength = rsTmp("app").DefinedSize
        txt(3).MaxLength = rsTmp("device").DefinedSize
        txt(4).MaxLength = rsTmp("note").DefinedSize
    End If
    
    Set rsTmp = gclsBusiness.ReadEventKind
    If rsTmp.BOF = False Then
        Do While Not rsTmp.EOF
            cbo(0).AddItem rsTmp("kind").Value
            rsTmp.MoveNext
        Loop
    End If
    
    InitData = True
End Function

Private Function ReadData(ByVal strDataKey As String) As Boolean

    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    Dim rsTmp As ADODB.Recordset
    Dim rsCondition As ADODB.Recordset
    
    Set rsCondition = zlCommFun.CreateCondition
    Call zlCommFun.SetCondition(rsCondition, "id", strDataKey)
    
    mblnReading = True
    Set rsTmp = gclsBusiness.EventRead("id", rsCondition)
    If rsTmp.BOF = False Then
        txt(0).Text = zlCommFun.NVL(rsTmp("title").Value)
        txt(1).Text = zlCommFun.NVL(rsTmp("code").Value)
        txt(2).Text = zlCommFun.NVL(rsTmp("app").Value)
        txt(3).Text = zlCommFun.NVL(rsTmp("device").Value)
        cbo(0).Text = zlCommFun.NVL(rsTmp("kind").Value)
        txt(4).Text = zlCommFun.NVL(rsTmp("note").Value)
    End If
    mblnReading = False
    mblnDataChanged = False
    
    ReadData = True
    
End Function

Private Function InitCommandBar() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom
    Dim objFindKey As CommandBarControl
    
    On Error GoTo errHand
    
    '------------------------------------------------------------------------------------------------------------------
    '初始设置
    Call zlCommFun.CommandBarInit(cbsMain)
    cbsMain.VisualTheme = xtpThemeNativeWinXP
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    cbsMain.Options.LargeIcons = False
    
    '------------------------------------------------------------------------------------------------------------------
    '菜单定义:包括公共部份，请对xtpControlPopup类型的命令ID重新赋值

    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.ActiveMenuBar.Visible = False
    
    '------------------------------------------------------------------------------------------------------------------
    '工具栏定义:包括公共部份
    
    
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched
        
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_View_Option, IIf(mbytMode = 1, "确定之继续新增", "确定之继续修改"), True)
    objControl.IconId = conMenu_View_UnCheck
    
    mstrFindKey = zlDataBase.GetPara("定位依据", ParamInfo.系统号, mlngModualCode, "名称")
    If mstrFindKey = "" Then mstrFindKey = "名称"

    Set mobjFindKey = zlCommFun.NewToolBar(objBar, xtpControlPopup, conMenu_View_LocationItem, mstrFindKey, True, , xtpButtonIconAndCaption)
    mobjFindKey.IconId = conMenu_View_Find
    mobjFindKey.Flags = xtpFlagRightAlign
    mobjFindKey.Style = xtpButtonIconAndCaption
    Set objControl = mobjFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_View_LocationItem, "&1.名称"): objControl.Parameter = "名称"
    objControl.IconId = 1
    Set objControl = mobjFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_View_LocationItem, "&2.编码"): objControl.Parameter = "编码"
    objControl.IconId = 1

    Set cbrCustom = zlCommFun.NewToolBar(objBar, xtpControlCustom, 0, "")
    cbrCustom.Handle = txtLocation.hWnd
        
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_View_Filter, "搜索")
    
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_View_Forward, "上一条")
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_View_Backward, "下一条")
        
    
    txtLocation.Visible = (mbytMode = 2)
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
    
End Function

Private Function ValidData() As Boolean
    '******************************************************************************************************************
    '功能：校验编辑数据的有效性
    '参数：
    '返回：
    '******************************************************************************************************************
        
        
    If Len(txt(0).Text) = 0 Then
        ShowSimpleMsg "业务事件的名称不能为空！"
        Call LocationObj(txt(0))
        Exit Function
    End If
    
    If Len(txt(1).Text) = 0 Then
        ShowSimpleMsg "业务事件的编码不能为空！"
        Call LocationObj(txt(1))
        Exit Function
    End If
    
    '检查编码是否为数字字符
    If zlCommFun.CheckStrType(Trim(txt(1).Text), 99, "0123456789") = False Then
        ShowSimpleMsg "业务事件的编码必须为数字字符！"
        LocationObj txt(1)
        Exit Function
    End If
    
    ValidData = True
    
End Function

Private Function SaveData(ByRef strDataKey As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rsPara As ADODB.Recordset
    
    On Error GoTo errHand
    
    Set rsPara = zlCommFun.CreateParameter
    
    Call zlCommFun.SetParameter(rsPara, "id", strDataKey)
    Call zlCommFun.SetParameter(rsPara, "title", Trim(txt(0).Text))
    Call zlCommFun.SetParameter(rsPara, "code", Trim(txt(1).Text))
    Call zlCommFun.SetParameter(rsPara, "app", Trim(txt(2).Text))
    Call zlCommFun.SetParameter(rsPara, "device", Trim(txt(3).Text))
    Call zlCommFun.SetParameter(rsPara, "kind", Trim(cbo(0).Text))
    Call zlCommFun.SetParameter(rsPara, "note", Trim(txt(4).Text))
        
    Select Case mbytMode
    '------------------------------------------------------------------------------------------------------------------
    Case 1          '新增
        strDataKey = zlCommFun.GetGUID
        Call zlCommFun.SetParameter(rsPara, "id", strDataKey)
        
        SaveData = gclsBusiness.EventEdit("INSERT", rsPara)
    '------------------------------------------------------------------------------------------------------------------
    Case 2          '修改
        SaveData = gclsBusiness.EventEdit("UPDATE", rsPara)
    End Select
    
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub cbo_Change(Index As Integer)
    mblnDataChanged = True
End Sub

Private Sub cbo_Click(Index As Integer)
    mblnDataChanged = True
End Sub

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    
    Dim blnCancel As Boolean
    Dim strDataKey As String
    
    Select Case Control.ID
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Forward               '上一条
        
        strDataKey = mstrDataKey
        
        RaiseEvent Forward(strDataKey, blnCancel)
        If blnCancel = False Then
        
            mstrDataKey = strDataKey
            Call ReadData(mstrDataKey)
    
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Backward               '下一条
        
        strDataKey = mstrDataKey
        
        RaiseEvent Backward(strDataKey, blnCancel)
        If blnCancel = False Then
            
            mstrDataKey = strDataKey
            Call ReadData(mstrDataKey)
            
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_LocationItem
        mstrFindKey = Control.Parameter
        mobjFindKey.Caption = mstrFindKey
        cbsMain.RecalcLayout
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Filter
        
        Dim strText As String
        Dim rsCondition As ADODB.Recordset
        Dim rsData As ADODB.Recordset
        Dim rs As ADODB.Recordset
        
        If txtLocation.Text <> "" Then
            
            txtLocation.Tag = ""
            
            
            Set rsCondition = zlCommFun.CreateCondition
            
            Call zlCommFun.SetCondition(rsCondition, "FilterStyle", mstrFindKey)
            Call zlCommFun.SetCondition(rsCondition, "FilterText", txtLocation.Text)
            Set rsData = gclsBusiness.EventRead("FilterData", rsCondition)
            
            If zlCommFun.ShowPubSelect(Me, txtLocation, 2, "名称,1500,0,1;编码,1500,0,0;程序,1500,0,0;设备,1500,0,0", Me.Name & "\业务事件过滤", "请从下表中选择一个业务事件", rsData, rs, , , , , , True) = 1 Then
                mstrDataKey = rs("id").Value
                Call ReadData(mstrDataKey)
                txtLocation.Tag = ""
            Else
                txtLocation.Tag = ""
                Call LocationObj(txtLocation, True)
                Exit Sub
            End If
                        
            Call LocationObj(txtLocation, True)
        End If
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_LocationItem
    
        mstrFindKey = Control.Parameter
        mobjFindKey.Caption = mstrFindKey
        cbsMain.RecalcLayout
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Option
        mblnContiune = Not mblnContiune
    End Select
    
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long
    Dim lngTop  As Long
    Dim lngRight  As Long
    Dim lngBottom  As Long

    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)

    On Error Resume Next
    
    '窗体其它控件Resize处理
    picPane(0).Move lngLeft, lngTop, lngRight - lngLeft, lngBottom - lngTop
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_View_Filter, conMenu_View_LocationItem, conMenu_View_Backward, conMenu_View_Forward, 0
        Control.Visible = (mbytMode = 2)
        Control.Enabled = Not mblnDataChanged
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Option
        Control.Checked = mblnContiune
        Control.IconId = IIf(mblnContiune = True, conMenu_View_Check, conMenu_View_UnCheck)
    End Select
End Sub

Private Sub cmdCancel_Click()
    '
    Unload Me
End Sub

Private Sub cmdOK_Click()
        
    If mblnDataChanged = True Then
        If ValidData = True Then
                
            RaiseEvent BeforeNewData(mstrDataKey)
            
            If SaveData(mstrDataKey) = True Then
                
                Select Case mbytMode
                Case 1
                    RaiseEvent AfterNewData(mstrDataKey)
                Case 2
                    RaiseEvent AfterModifyData(mstrDataKey)
                End Select
                
                If mblnContiune = False Then
                    mblnDataChanged = False
                    Unload Me
                Else
                    '重置环境，进入下一次新增状态
                    If mbytMode = 1 Then
                        mstrDataKey = ""
                        txt(0).Text = ""
                        txt(1).Text = gclsBusiness.GetMaxCode("m_Event", "code")
                        txt(2).Text = ""
                        txt(3).Text = ""
                        txt(4).Text = ""
                    End If
                    Call LocationObj(txt(0))
                    mblnDataChanged = False
                End If
                
            End If
        End If
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnDataChanged Then
        Cancel = (MsgBox("新增或修改的数据必须保存后才生效，是否不保存就退出？", vbYesNo + vbQuestion + vbDefaultButton2, ParamInfo.系统名称) = vbNo)
        If Cancel Then Exit Sub
    End If
End Sub

Private Sub txt_Change(Index As Integer)
    
    If mblnReading Then Exit Sub
    
    mblnDataChanged = True
End Sub

Private Sub txt_GotFocus(Index As Integer)
    
    zlControl.TxtSelAll txt(Index)
    
    Select Case Index
    Case 4
        zlCommFun.OpenIme True
    End Select
        
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        
        '
        
    End If
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim strText As String
    Dim strTmp As String
    Dim bytMode As Byte
    Dim rsData As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
        zlCommFun.PressKey vbKeyTab
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
        Select Case Index
        Case 1
            If zlCommFun.FilterKeyAscii(KeyAscii, 99, "0123456789") = 0 Then KeyAscii = 0
        End Select
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)

    Select Case Index
    Case 4
        zlCommFun.OpenIme False
    End Select

End Sub

Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        glngTXTProc = GetWindowLong(txt(Index).hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not zlCommFun.StrIsValid(txt(Index).Text, txt(Index).MaxLength)
End Sub

Private Sub txtLocation_Change()
    txtLocation.Tag = "Changed"
End Sub

Private Sub txtLocation_GotFocus()
    zlControl.TxtSelAll txtLocation
End Sub

Private Sub txtLocation_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyDelete Then
        KeyCode = 0
        txtLocation.Text = ""
        txtLocation.Tag = ""
    End If

End Sub

Private Sub txtLocation_KeyPress(KeyAscii As Integer)
    Dim strText As String
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim rsCondition As ADODB.Recordset
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0

        If txtLocation.Text <> "" Then
            txtLocation.Tag = ""
            
            Dim obj As CommandBarControl
            
            Set obj = cbsMain.FindControl(, conMenu_View_Filter, True)
            If obj.Enabled = True Then
                Call cbsMain_Execute(obj)
            End If

        End If
        txtLocation.Tag = ""
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End If
End Sub

Private Sub txtLocation_Validate(Cancel As Boolean)
    If (txtLocation.Tag = "Changed") Then
        txtLocation.Tag = ""
    End If
End Sub


