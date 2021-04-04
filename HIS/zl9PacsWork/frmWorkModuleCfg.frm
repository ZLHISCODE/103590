VERSION 5.00
Begin VB.Form frmWorkModuleCfg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "站点模式配置"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3495
   Icon            =   "frmWorkModuleCfg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.ListBox lstModule 
      Height          =   2580
      ItemData        =   "frmWorkModuleCfg.frx":000C
      Left            =   120
      List            =   "frmWorkModuleCfg.frx":000E
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   120
      Width           =   3255
   End
   Begin VB.CheckBox chkModule 
      Caption         =   "排队叫号已在流程管理设置中启用"
      Enabled         =   0   'False
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Value           =   1  'Checked
      Width           =   3255
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取 消(&C)"
      Height          =   400
      Left            =   2160
      TabIndex        =   1
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdSure 
      Caption         =   "确 定(&S)"
      Height          =   400
      Left            =   840
      TabIndex        =   0
      Top             =   3120
      Width           =   1215
   End
End
Attribute VB_Name = "frmWorkModuleCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngModule As Long

Public blnIsUseQueue As Boolean
Public blnIsOk As Boolean
Public strWorkModule As String


Public Sub ShowWorkModuleCfg(ByVal lngModule As Long, owner As Object)
    mlngModule = lngModule
    
    chkModule(5).Visible = IIf(mlngModule = G_LNG_PATHOLSYS_NUM, False, True)
    
    If blnIsUseQueue Then
        chkModule(5).value = 1
        chkModule(5).Caption = "排队叫号已在流程管理设置中(启用)"
    Else
        chkModule(5).value = 0
        chkModule(5).Caption = "排队叫号已在流程管理设置中(禁用)"
    End If
    
    Call Me.LoadDefaultModuleConfig
    Call Me.ReadWorkModuleCfg
    
    blnIsOk = False
    
    Me.Show 1, owner
End Sub


Public Sub LoadDefaultModuleConfig()
    Dim strModuleName As String
    Dim strModuleItem() As String
    Dim i As Long
    
    Select Case mlngModule
        Case G_LNG_PACSSTATION_MODULE
            strModuleName = "影像图像模块;影像报告模块;病历记录模块;费用记录模块;医嘱记录模块;电子病历模块;"
            
        Case G_LNG_VIDEOSTATION_MODULE
            strModuleName = "影像采集模块;影像报告模块;病历记录模块;费用记录模块;医嘱记录模块;电子病历模块;"
            
        Case G_LNG_PATHOLSYS_NUM
            strModuleName = "影像采集模块;标本核收模块;病理取材模块;病理制片模块;病理特检模块;过程报告模块;病理诊断模块;病历记录模块;费用记录模块;医嘱记录模块;电子病历模块;"
        Case Else
            Exit Sub
    End Select
    
    strModuleItem = Split(strModuleName, ";")
    
    lstModule.Clear
    
    For i = LBound(strModuleItem) To UBound(strModuleItem)
        If strModuleItem(i) <> "" Then lstModule.AddItem strModuleItem(i)
    Next i
    
End Sub


Public Sub ReadWorkModuleCfg()
    Dim i As Long
    Dim blnAll As Boolean

    '读取站点参数设置，如为空则默认勾选 除排队叫号所有模块
    strWorkModule = zlDatabase.GetPara("站点模块", glngSys, mlngModule, "")
    
    If strWorkModule <> "" Then strWorkModule = ";" & strWorkModule & ";"
    
    For i = 0 To lstModule.ListCount - 1
        If strWorkModule = "" And Not blnIsUseQueue Then
            lstModule.Selected(i) = True
        Else
            lstModule.Selected(i) = IIf(InStr(strWorkModule, ";" & lstModule.list(i) & ";") > 0, True, False)
        End If
    Next i
    
End Sub


Private Sub cmdCancel_Click()
On Error Resume Next
    blnIsOk = False
    
    Call Me.Hide
    Call err.Clear
End Sub

Private Sub cmdSure_Click()
On Error GoTo errHandle
    Dim i As Long
    
    strWorkModule = ""
    For i = 0 To lstModule.ListCount - 1
        If lstModule.Selected(i) Then
            If strWorkModule <> "" Then strWorkModule = strWorkModule & ";"
            strWorkModule = strWorkModule & lstModule.list(i)
        End If
    Next i
    
    If strWorkModule = "" And chkModule(5).value = 0 Then
        Call MsgBoxD(Me, "请至少选择一种工作模块。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    strWorkModule = IIf(strWorkModule <> "", ";" & strWorkModule & ";", ";NULL;")
    
    Call zlDatabase.SetPara("站点模块", strWorkModule, glngSys, mlngModule)
    
    blnIsOk = True
    
    Call Me.Hide
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Load()
On Error GoTo errHandle
    Call RestoreWinState(Me, App.ProductName)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub




Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
    Call err.Clear
End Sub
