VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmServiceEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "#"
   ClientHeight    =   4590
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6885
   Icon            =   "frmServiceEdit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtLocation 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   1470
      TabIndex        =   14
      Top             =   60
      Width           =   1575
   End
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   4155
      Index           =   0
      Left            =   30
      ScaleHeight     =   4155
      ScaleWidth      =   7215
      TabIndex        =   12
      Top             =   510
      Width           =   7215
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   4200
         TabIndex        =   10
         Top             =   3570
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(C)"
         Height          =   350
         Left            =   5535
         TabIndex        =   11
         Top             =   3570
         Width           =   1100
      End
      Begin VB.Frame fra 
         Height          =   3465
         Left            =   15
         TabIndex        =   13
         Top             =   -15
         Width           =   6840
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   0
            Left            =   4575
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   240
            Width           =   2145
         End
         Begin VB.TextBox txt 
            Height          =   315
            Index           =   5
            Left            =   615
            TabIndex        =   1
            Top             =   240
            Width           =   3420
         End
         Begin VB.TextBox txt 
            Height          =   315
            Index           =   2
            Left            =   615
            TabIndex        =   5
            Top             =   630
            Width           =   3420
         End
         Begin VB.TextBox txt 
            Height          =   315
            Index           =   3
            Left            =   4575
            TabIndex        =   7
            Top             =   630
            Width           =   2145
         End
         Begin VB.TextBox txt 
            Height          =   705
            Index           =   4
            Left            =   615
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   9
            Top             =   2670
            Width           =   6060
         End
         Begin VB.Frame Frame1 
            Height          =   1635
            Left            =   615
            TabIndex        =   16
            Top             =   960
            Width           =   6105
            Begin VB.OptionButton opt 
               Caption         =   "WebService"
               Height          =   225
               Index           =   1
               Left            =   1155
               TabIndex        =   18
               Top             =   30
               Width           =   1335
            End
            Begin VB.OptionButton opt 
               Caption         =   "Socket"
               Height          =   225
               Index           =   0
               Left            =   75
               TabIndex        =   17
               Top             =   30
               Value           =   -1  'True
               Width           =   1155
            End
            Begin VB.PictureBox picBack 
               BorderStyle     =   0  'None
               Height          =   1275
               Index           =   0
               Left            =   30
               ScaleHeight     =   1275
               ScaleWidth      =   6045
               TabIndex        =   24
               Top             =   330
               Width           =   6045
               Begin VB.TextBox txt 
                  Height          =   345
                  Index           =   10
                  Left            =   3030
                  TabIndex        =   34
                  Top             =   885
                  Width           =   2895
               End
               Begin VB.TextBox txt 
                  Height          =   345
                  Index           =   9
                  Left            =   885
                  TabIndex        =   33
                  Top             =   900
                  Width           =   1245
               End
               Begin VB.TextBox txt 
                  Height          =   345
                  Index           =   8
                  Left            =   900
                  TabIndex        =   29
                  Top             =   495
                  Width           =   2070
               End
               Begin VB.TextBox txt 
                  Height          =   345
                  Index           =   7
                  Left            =   3825
                  TabIndex        =   28
                  Top             =   495
                  Width           =   2100
               End
               Begin VB.TextBox txt 
                  Height          =   315
                  Index           =   6
                  Left            =   885
                  TabIndex        =   25
                  Top             =   120
                  Width           =   5025
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  Caption         =   "空间地址"
                  Height          =   180
                  Index           =   12
                  Left            =   2190
                  TabIndex        =   32
                  Top             =   960
                  Width           =   720
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  Caption         =   "空间标识"
                  Height          =   180
                  Index           =   11
                  Left            =   120
                  TabIndex        =   31
                  Top             =   960
                  Width           =   720
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  Caption         =   "调用方法"
                  Height          =   180
                  Index           =   10
                  Left            =   105
                  TabIndex        =   30
                  Top             =   570
                  Width           =   720
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  Caption         =   "方法参数"
                  Height          =   180
                  Index           =   9
                  Left            =   3030
                  TabIndex        =   27
                  Top             =   570
                  Width           =   720
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  Caption         =   "调用地址"
                  Height          =   180
                  Index           =   7
                  Left            =   105
                  TabIndex        =   26
                  Top             =   180
                  Width           =   720
               End
            End
            Begin VB.PictureBox picBack 
               BorderStyle     =   0  'None
               Height          =   855
               Index           =   1
               Left            =   90
               ScaleHeight     =   855
               ScaleWidth      =   5970
               TabIndex        =   19
               Top             =   390
               Width           =   5970
               Begin VB.TextBox txt 
                  Height          =   315
                  Index           =   1
                  Left            =   450
                  TabIndex        =   21
                  Top             =   510
                  Width           =   5310
               End
               Begin VB.TextBox txt 
                  Height          =   315
                  Index           =   0
                  Left            =   450
                  TabIndex        =   20
                  Top             =   90
                  Width           =   5310
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  Caption         =   "PORT"
                  Height          =   180
                  Index           =   1
                  Left            =   30
                  TabIndex        =   23
                  Top             =   570
                  Width           =   360
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  Caption         =   "IP"
                  Height          =   180
                  Index           =   0
                  Left            =   45
                  TabIndex        =   22
                  Top             =   150
                  Width           =   180
               End
            End
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "接口"
            Height          =   180
            Index           =   8
            Left            =   165
            TabIndex        =   15
            Top             =   1050
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "类型"
            Height          =   180
            Index           =   6
            Left            =   4170
            TabIndex        =   2
            Top             =   300
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "名称"
            Height          =   180
            Index           =   5
            Left            =   165
            TabIndex        =   0
            Top             =   315
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "程序"
            Height          =   195
            Index           =   2
            Left            =   165
            TabIndex        =   4
            Top             =   705
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "设备"
            Height          =   195
            Index           =   3
            Left            =   4170
            TabIndex        =   6
            Top             =   690
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "说明"
            Height          =   195
            Index           =   4
            Left            =   165
            TabIndex        =   8
            Top             =   2655
            Width           =   360
         End
      End
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   -15
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmServiceEdit"
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
Private mblnContiune As Boolean

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
    Me.Caption = "新增服务配置"
    mstrDataKey = ""
    
    Call InitData
    Call InitCommandBar
    
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
    
    Me.Caption = "修改服务配置"

    Call InitData
    Call InitCommandBar
    
    Call ReadData(mstrDataKey)
    
    Me.Show 1, mfrmParent
    
End Sub

Public Sub DeleteData(ByVal strDataKey As String)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mbytMode = 3
    If strDataKey = "" Then Exit Sub
    mstrDataKey = strDataKey
    
    Set mrsPara = zlCommFun.CreateParameter
    Call zlCommFun.SetParameter(mrsPara, "ID", mstrDataKey)
        
    If gclsMsgBase.EditService("Delete", mrsPara) Then
        RaiseEvent AfterDeleteData(mstrDataKey)
    End If
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
    
    Set rsTmp = gclsMsgBase.GetServiceStruct()
    If Not (rsTmp Is Nothing) Then
        txt(0).MaxLength = 15
        txt(1).MaxLength = 5
        txt(2).MaxLength = rsTmp("app").DefinedSize
        txt(3).MaxLength = rsTmp("device").DefinedSize
        txt(4).MaxLength = rsTmp("note").DefinedSize
        txt(5).MaxLength = rsTmp("title").DefinedSize
        txt(6).MaxLength = rsTmp("interface_para").DefinedSize
        txt(7).MaxLength = 50
        txt(8).MaxLength = 100
        txt(9).MaxLength = 20
        txt(10).MaxLength = 50
    End If
    
    With cbo(0)
        .AddItem "1-目标服务"
        .ItemData(.NewIndex) = 1
        
        .AddItem "2-接收服务"
        .ItemData(.NewIndex) = 2
        
        .ListIndex = 0
    End With
    
    picBack(0).Visible = False
    picBack(1).Visible = True
        
    InitData = True
End Function

Private Function ReadData(ByVal strDataKey As String) As Boolean

    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strTemp As String
    Dim rsTmp As ADODB.Recordset
    Dim rsCondition As ADODB.Recordset
    Dim varTemp As Variant
    
    Set rsCondition = zlCommFun.CreateCondition
    Call zlCommFun.SetCondition(rsCondition, "id", strDataKey)
    
    mblnReading = True
    Set rsTmp = gclsMsgBase.GetService("id", rsCondition)
    If rsTmp.BOF = False Then

        txt(2).Text = zlCommFun.NVL(rsTmp("app").Value)
        txt(3).Text = zlCommFun.NVL(rsTmp("device").Value)
        txt(4).Text = zlCommFun.NVL(rsTmp("note").Value)
        txt(5).Text = zlCommFun.NVL(rsTmp("title").Value)
        strTemp = zlCommFun.NVL(rsTmp("interface_para").Value)
        Select Case zlCommFun.NVL(rsTmp("interface_type").Value)
        Case 1
            opt(0).Value = True
            If InStr(strTemp, "/") > 0 Then
                txt(0).Text = Mid(strTemp, 1, InStr(strTemp, "/") - 1)
                txt(1).Text = Mid(strTemp, InStr(strTemp, "/") + 1)
            End If
        Case 2
            '<root><address></address><method></method></root>
            
            varTemp = Split(strTemp, vbCrLf)
    
            txt(6).Text = varTemp(0)
            If UBound(varTemp) > 0 Then txt(8).Text = varTemp(1)
            If UBound(varTemp) > 1 Then txt(7).Text = varTemp(2)
            If UBound(varTemp) > 2 Then txt(9).Text = varTemp(3)
            If UBound(varTemp) > 3 Then txt(10).Text = varTemp(4)
            
            opt(1).Value = True
            
        End Select
        
        Call zlControl.CboLocate(cbo(0), zlCommFun.NVL(rsTmp("kind").Value), True)
        
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
    
    Select Case mbytMode
    Case 1
        Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_View_Option, "确定后继续新增", True)
    Case 2
        Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_View_Option, "确定后继续修改", True)
    End Select
    objControl.IconId = conMenu_View_UnCheck
    
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlLabel, conMenu_View_LocationItem, "名称", True, , xtpButtonIconAndCaption)
    objControl.IconId = conMenu_View_Find

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
    
    If Len(Trim(txt(5).Text)) = 0 Then
        ShowSimpleMsg "服务的名称不能为空！"
        Call LocationObj(txt(5))
        Exit Function
    End If
    
    If zlCommFun.CheckValid(txt(0).Text, emDataFormat.IP4) = False Then
        ShowSimpleMsg "服务地址无效，服务地址格式为：[0-255].[0-255].[0-255].[0-255]，如192.168.4.51！"
        Call LocationObj(txt(0))
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
    
    If opt(0).Value Then
        Call zlCommFun.SetParameter(rsPara, "interface_type", 1)
        Call zlCommFun.SetParameter(rsPara, "interface_para", Trim(txt(0).Text) & "/" & Val(txt(1).Text))
    Else
        Call zlCommFun.SetParameter(rsPara, "interface_type", 2)
        Call zlCommFun.SetParameter(rsPara, "interface_para", Trim(txt(6).Text) & vbCrLf & Trim(txt(8).Text) & vbCrLf & Trim(txt(7).Text) & vbCrLf & Trim(txt(9).Text) & vbCrLf & Trim(txt(10).Text))
    End If
    Call zlCommFun.SetParameter(rsPara, "app", Trim(txt(2).Text))
    Call zlCommFun.SetParameter(rsPara, "device", Trim(txt(3).Text))
    Call zlCommFun.SetParameter(rsPara, "note", Trim(txt(4).Text))
    Call zlCommFun.SetParameter(rsPara, "title", Trim(txt(5).Text))
    Call zlCommFun.SetParameter(rsPara, "kind", Trim(cbo(0).ItemData(cbo(0).ListIndex)))
        
    Select Case mbytMode
    '------------------------------------------------------------------------------------------------------------------
    Case 1          '新增
        strDataKey = zlCommFun.GetGUID
        Call zlCommFun.SetParameter(rsPara, "id", strDataKey)
        
        SaveData = gclsMsgBase.EditService("INSERT", rsPara)
    '------------------------------------------------------------------------------------------------------------------
    Case 2          '修改
        SaveData = gclsMsgBase.EditService("UPDATE", rsPara)
    End Select
    
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub cbo_Click(Index As Integer)
    mblnDataChanged = True
End Sub

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
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
    Case conMenu_View_Filter
        
        Dim strText As String
        Dim rsCondition As ADODB.Recordset
        Dim rsData As ADODB.Recordset
        Dim rs As ADODB.Recordset
        
        If txtLocation.Text <> "" Then
            
            txtLocation.Tag = ""
            
            strText = UCase(txtLocation.Text)
                        
            Set rsCondition = zlCommFun.CreateCondition
            
            Call zlCommFun.SetCondition(rsCondition, "server_ip", "%" & strText & "%")
            Set rsData = gclsMsgBase.GetService("title", rsCondition)
            
            If zlCommFun.ShowPubSelect(Me, txtLocation, 2, "类型,1200,0,1;名称,2400,0,1;地址,1500,0,1;端口,600,0,0;程序,900,0,0;设备,900,0,0", Me.Name & "\服务配置过滤", "请从下表中选择一个服务配置", rsData, rs, , , , , , True) = 1 Then
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
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Option
        Control.Checked = mblnContiune
        Control.IconId = IIf(mblnContiune = True, conMenu_View_Check, conMenu_View_UnCheck)
    End Select
    
End Sub

Private Sub chk_Click(Index As Integer)
    mblnDataChanged = True
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmdCancel_Click()
    '
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strOldDataKey As String
    Dim rsTmp As ADODB.Recordset
    
    If mblnDataChanged = True Then
        If ValidData = True Then
    
            If SaveData(mstrDataKey) = True Then
                
                If strOldDataKey <> "" Then
                    RaiseEvent AfterModifyData(strOldDataKey)
                End If
                
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
                        txt(1).Text = ""
                        txt(2).Text = ""
                        txt(3).Text = ""
                        txt(4).Text = ""
                        txt(5).Text = ""
                        txt(6).Text = ""
                        txt(7).Text = ""
                        txt(8).Text = ""
                        txt(9).Text = ""
                        txt(10).Text = ""
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

Private Sub opt_Click(Index As Integer)
    
    Select Case Index
    Case 0
        picBack(0).Visible = False
        picBack(1).Visible = True
    Case 1
        picBack(0).Visible = True
        picBack(1).Visible = False
    End Select
    
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
