VERSION 5.00
Begin VB.Form frm参数设置 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7320
   Icon            =   "frm参数设置.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fra是否按批次显示 
      Caption         =   "当前库房药品是否按批次显示(当按批次申领时)"
      Height          =   735
      Left            =   180
      TabIndex        =   24
      Top             =   4440
      Width           =   6975
      Begin VB.OptionButton opt按批次显示 
         Caption         =   "按出库批次显示当前库房的库存数量"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Value           =   -1  'True
         Width           =   3375
      End
      Begin VB.OptionButton opt按批次显示 
         Caption         =   "按当前库房该药品总数量显示"
         Height          =   180
         Index           =   1
         Left            =   3480
         TabIndex        =   25
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.Frame frm当前库存 
      Caption         =   "填单时当前库房库存显示方式"
      Height          =   735
      Left            =   180
      TabIndex        =   21
      Top             =   3480
      Width           =   3270
      Begin VB.OptionButton opt当前库存 
         Caption         =   "显示实际数量"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton opt当前库存 
         Caption         =   "显示可用数量"
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   22
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame frm对方库存 
      Caption         =   "填单时对方库房库存显示方式"
      Height          =   735
      Left            =   3570
      TabIndex        =   18
      Top             =   3480
      Width           =   3615
      Begin VB.OptionButton opt对方库存 
         Caption         =   "显示实际数量"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton opt对方库存 
         Caption         =   "显示可用数量"
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   19
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame fra库存显示控制 
      Caption         =   "库存显示控制"
      Height          =   855
      Left            =   180
      TabIndex        =   15
      Top             =   2520
      Width           =   6975
      Begin VB.CheckBox chkShow 
         Caption         =   "显示无库存的药品"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lbComment 
         Caption         =   "说明：申领时药品选择器中是否显示无库存的药品记录"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   240
         TabIndex        =   17
         Top             =   495
         Width           =   6420
      End
   End
   Begin VB.TextBox txt查询天数 
      Height          =   300
      Left            =   4395
      TabIndex        =   12
      Text            =   "1"
      Top             =   2130
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Frame fraSort 
      Caption         =   "排序方式"
      Height          =   1770
      Left            =   3510
      TabIndex        =   8
      Top             =   240
      Width           =   3675
      Begin VB.ComboBox Cbo列名 
         Height          =   300
         ItemData        =   "frm参数设置.frx":000C
         Left            =   120
         List            =   "frm参数设置.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   390
         Width           =   2415
      End
      Begin VB.ComboBox Cbo方向 
         Height          =   300
         Left            =   2700
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   390
         Width           =   885
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "    本参数的设置，将影响所有编辑窗体中单据的显示内容的排序方式。缺省：按用户输入的顺序显示各单据的内容"
         ForeColor       =   &H80000008&
         Height          =   705
         Left            =   180
         TabIndex        =   11
         Top             =   930
         Width           =   3345
      End
   End
   Begin VB.ComboBox Cbo指定单位 
      Height          =   300
      Left            =   1020
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.Frame Frame3 
      Height          =   1935
      Left            =   180
      TabIndex        =   2
      Top             =   480
      Width           =   3255
      Begin VB.ComboBox cbo报表 
         Height          =   300
         Left            =   120
         TabIndex        =   29
         Text            =   "Combo1"
         Top             =   1080
         Width           =   3015
      End
      Begin VB.CommandButton cmd打印设置 
         Caption         =   "打印设置(&P)"
         Height          =   315
         Left            =   120
         TabIndex        =   28
         Top             =   1440
         Width           =   2985
      End
      Begin VB.CheckBox chkPrintCode 
         Caption         =   "存盘或审核后打印药品条码"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   480
         Width           =   2535
      End
      Begin VB.CheckBox chkVerifyPrint 
         Caption         =   "审核打印单据"
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         Top             =   120
         Width           =   1455
      End
      Begin VB.CheckBox chkSavePrint 
         Caption         =   "存盘打印单据"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   90
      TabIndex        =   7
      Top             =   5400
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   4590
      TabIndex        =   5
      Top             =   5400
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5940
      TabIndex        =   6
      Top             =   5400
      Width           =   1100
   End
   Begin VB.Label lbl查询天数 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "查询天数"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3600
      TabIndex        =   14
      Top             =   2190
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label lbl天数 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "天"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   5340
      TabIndex        =   13
      Top             =   2190
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label lbl药品单位 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "药品单位"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   720
   End
End
Attribute VB_Name = "frm参数设置"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrFunction As String
Private mlngMode As Long
Dim mstrPrivs As String
Private mblnSetPara As Boolean                          '是否具有参数设置权限

Private Const M_LNG_FRMWIDTH_1 = 3800
Private Const M_LNG_FRMWIDTH_2 = 7500
Private Const M_LNG_FRMHEIGHT_1 = 3200
Private Const M_LNG_FRMHEIGHT_2 = 6315


Private Sub Cbo列名_Click()
    If Cbo方向.ListCount < 1 Then Exit Sub
    Cbo方向.Enabled = Not (Cbo列名.ListIndex = 0)
    If Not Cbo方向.Enabled Then Cbo方向.ListIndex = 0
End Sub


Private Sub chkSavePrint_Click()
    chkPrintCode.Enabled = chkVerifyPrint.Value = 1 Or chkSavePrint.Value = 1
End Sub

Private Sub chkVerifyPrint_Click()
    chkPrintCode.Enabled = chkVerifyPrint.Value = 1 Or chkSavePrint.Value = 1
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub cmdOK_Click()
    If mlngMode = 1343 Then
        If Trim(txt查询天数.Text) = "" Then
            MsgBox "请输入查询天数（1天-365天）！", vbInformation, gstrSysName
            txt查询天数.SetFocus
            Exit Sub
        End If
        If Not IsNumeric(txt查询天数.Text) Then
            MsgBox "查询天数中含有非法字符！", vbInformation, gstrSysName
            txt查询天数.SetFocus
            Exit Sub
        End If
        If Val(txt查询天数.Text) < 1 Or Val(txt查询天数.Text) > 365 Then
            MsgBox "查询天数不能小于1天或大于365天！", vbInformation, gstrSysName
            txt查询天数.SetFocus
            Exit Sub
        End If
    End If
    
    On Error Resume Next
    
    Select Case mlngMode
        Case 1343   '药品申领
'            zldatabase.SetPara "是否选择库房", IIf(chkStock.Value = 1, "1", "0"), glngSys, mlngMode
            zldatabase.SetPara "药品单位", Cbo指定单位.ListIndex, glngSys, mlngMode
            zldatabase.SetPara "排序", CStr(Cbo列名.ListIndex) & CStr(Cbo方向.ListIndex), glngSys, mlngMode
            zldatabase.SetPara "存盘打印", IIf(chkSavePrint.Value = 1, "1", "0"), glngSys, mlngMode
            zldatabase.SetPara "审核打印", IIf(chkVerifyPrint.Value = 1, "1", "0"), glngSys, mlngMode
            zldatabase.SetPara "打印药品条码", IIf(chkPrintCode.Value = 1, "1", "0"), glngSys, mlngMode
            zldatabase.SetPara "查询天数", Val(txt查询天数.Text), glngSys, mlngMode
            zldatabase.SetPara "显示无库存药品", chkShow.Value, glngSys, mlngMode
            zldatabase.SetPara "填单时当前库房库存显示方式", IIf(opt当前库存(0).Value = True, 0, 1), glngSys, mlngMode
            zldatabase.SetPara "填单时对方库房库存显示方式", IIf(opt对方库存(0).Value = True, 0, 1), glngSys, mlngMode
            zldatabase.SetPara "当前库房药品数量是否按批次显示", IIf(opt按批次显示(0).Value = True, 0, 1), glngSys, mlngMode
        Case 1344   '协定入库
'            zldatabase.SetPara "是否选择库房", IIf(chkStock.Value = 1, "1", "0"), glngSys, mlngMode
            zldatabase.SetPara "存盘打印", IIf(chkSavePrint.Value = 1, "1", "0"), glngSys, mlngMode
            zldatabase.SetPara "审核打印", IIf(chkVerifyPrint.Value = 1, "1", "0"), glngSys, mlngMode
            zldatabase.SetPara "药品单位", Cbo指定单位.ListIndex, glngSys, mlngMode
    End Select
    
    Unload Me
End Sub

Public Sub 设置参数(frmParent As Object, ByVal strPrivs As String, Optional ByVal intMode As Integer = 1344, Optional ByVal strFunction As String = "")
    mstrFunction = strFunction
    mlngMode = intMode
    mstrPrivs = strPrivs
    
    Dim int是否选择库房 As Integer
    Dim int药品单位 As Integer
    Dim str排序 As String
    Dim int存盘打印 As Integer
    Dim int审核打印 As Integer
    Dim int打印药品条码 As Integer
    Dim int查询天数 As Integer
    Dim int显示无库存药品 As Integer
    Dim int当前库存显示方式 As Integer
    Dim int对方库存显示方式 As Integer
    Dim int当前库存按批次显示 As Integer
    
    mblnSetPara = IsHavePrivs(mstrPrivs, "参数设置")
    
    '取公共及私有参数
    Select Case mlngMode
        Case 1343   '药品申领
            int药品单位 = Val(zldatabase.GetPara("药品单位", glngSys, mlngMode, 0, Array(lbl药品单位, Cbo指定单位), mblnSetPara))
            str排序 = zldatabase.GetPara("排序", glngSys, mlngMode, "00", Array(fraSort, Cbo列名, Cbo方向, Label5), mblnSetPara)
            int存盘打印 = Val(zldatabase.GetPara("存盘打印", glngSys, mlngMode, 0, Array(chkSavePrint), mblnSetPara))
            int审核打印 = Val(zldatabase.GetPara("审核打印", glngSys, mlngMode, 0, Array(chkVerifyPrint), mblnSetPara))
            int打印药品条码 = Val(zldatabase.GetPara("打印药品条码", glngSys, mlngMode, 0, Array(chkPrintCode), mblnSetPara))
            int查询天数 = Val(zldatabase.GetPara("查询天数", glngSys, mlngMode, 7, Array(lbl查询天数, txt查询天数, lbl天数), mblnSetPara))
            int显示无库存药品 = Val(zldatabase.GetPara("显示无库存药品", glngSys, mlngMode, 0, Array(fra库存显示控制, chkShow), mblnSetPara))
            int当前库存显示方式 = Val(zldatabase.GetPara("填单时当前库房库存显示方式", glngSys, mlngMode, 0, Array(opt当前库存(0), opt当前库存(1)), mblnSetPara))
            int对方库存显示方式 = Val(zldatabase.GetPara("填单时对方库房库存显示方式", glngSys, mlngMode, 0, Array(opt对方库存(0), opt对方库存(1)), mblnSetPara))
            int当前库存按批次显示 = Val(zldatabase.GetPara("当前库房药品数量是否按批次显示", glngSys, mlngMode, 0, Array(opt按批次显示(0), opt按批次显示(1)), mblnSetPara))
        Case 1344   '协定入库
'            int是否选择库房 = Val(zldatabase.GetPara("是否选择库房", glngSys, mlngMode, 0, Array(chkStock, Label2), mblnSetPara))
            int存盘打印 = Val(zldatabase.GetPara("存盘打印", glngSys, mlngMode, 0, Array(chkSavePrint), mblnSetPara))
            int审核打印 = Val(zldatabase.GetPara("审核打印", glngSys, mlngMode, 0, Array(chkVerifyPrint), mblnSetPara))
            int药品单位 = Val(zldatabase.GetPara("药品单位", glngSys, mlngMode, 0, Array(lbl药品单位, Cbo指定单位), mblnSetPara))
    End Select
    
    '根据参数值设置
'    If int是否选择库房 = 0 Then
'        chkStock.Value = 0
'    Else
'        chkStock.Value = 1
'    End If
    If int存盘打印 = 0 Then
        chkSavePrint.Value = 0
    Else
        chkSavePrint.Value = 1
    End If
    
    If int审核打印 = 0 Then
        chkVerifyPrint.Value = 0
    Else
        chkVerifyPrint.Value = 1
    End If
    
    chkPrintCode.Enabled = chkVerifyPrint.Value = 1 Or chkSavePrint.Value = 1
    
    If int打印药品条码 = 0 Then
        chkPrintCode.Value = 0
    Else
        chkPrintCode.Value = 1
    End If
    
    With Cbo指定单位
        .Clear
        .AddItem "缺省（当前库房对应的单位）"
        If glngSys \ 100 = 8 Then
            .AddItem "采购单位"
            .AddItem "售价单位"
        Else
            .AddItem "药库单位"
            .AddItem "门诊单位"
            .AddItem "住院单位"
            .AddItem "售价单位"
        End If
        .ListIndex = int药品单位
    End With
    
    fra库存显示控制.Visible = False
    
    Select Case mlngMode
        Case 1343   '申领
            fra库存显示控制.Visible = True
'            Frame3.Top = Frame2.Top
'            Frame2.Visible = True
'            chkVerifyPrint.Visible = False
'            Label3.Caption = Replace(Label3.Caption, "审核打印与此同理。", "")
            lbl药品单位.Visible = True
            Cbo指定单位.Visible = True
            
            fraSort.Visible = True
            Me.Width = M_LNG_FRMWIDTH_2
            Me.Height = M_LNG_FRMHEIGHT_2
            
            CmdCancel.Top = Me.Height - CmdCancel.Height - 500
            CmdOK.Top = CmdCancel.Top
            CmdHelp.Top = CmdCancel.Top
            
            CmdCancel.Left = M_LNG_FRMWIDTH_2 - CmdCancel.Width - 400
            CmdOK.Left = CmdCancel.Left - CmdOK.Width - 200
            
            Dim strValue As String
            mstrFunction = strFunction
            
            '装入缺省数据
            With Cbo列名
                .Clear
                .AddItem "输入顺序"
                .ItemData(.NewIndex) = 0
                .AddItem "编码"
                .ItemData(.NewIndex) = 1
                .AddItem "药品名称"
                .ItemData(.NewIndex) = 2
                .AddItem "库房货位"
                .ItemData(.NewIndex) = 3
                .ListIndex = 0
            End With
            With Cbo方向
                .Clear
                .AddItem "升序"
                .ItemData(.NewIndex) = 0
                .AddItem "降序"
                .ItemData(.NewIndex) = 1
                .ListIndex = 0
            End With
            
            '取排序字段及方向，如果为缺省，则置cbo方向.Enabled=False
            strValue = str排序
            Cbo列名.ListIndex = Mid(strValue, 1, 1)
            Cbo方向.ListIndex = Right(strValue, 1)
            Cbo方向.Enabled = Not (Cbo列名.ListIndex = 0)
            
            lbl查询天数.Visible = True
            txt查询天数.Visible = True
            lbl天数.Visible = True
            
            txt查询天数.Text = int查询天数
            
            chkShow.Value = IIf(int显示无库存药品 = 1, 1, 0)
            
            If int当前库存显示方式 = 1 Then
                opt当前库存(1).Value = True
            Else
                opt当前库存(0).Value = True
            End If
            
            If int对方库存显示方式 = 1 Then
                opt对方库存(1).Value = True
            Else
                opt对方库存(0).Value = True
            End If
            
            If int当前库存按批次显示 = 1 Then
                opt按批次显示(1).Value = True
            Else
                opt按批次显示(0).Value = True
            End If
            
        Case 1344   '协定
'            Frame3.Top = Frame2.Top + Frame2.Height + cmd打印设置.Height + 200
            cmd打印设置.Top = Cbo指定单位.Top + Cbo指定单位.Height + 200
             Frame3.Top = cmd打印设置.Top + cmd打印设置.Height - 300
'            Me.Height = 4000

            fraSort.Visible = False
            Me.Width = M_LNG_FRMWIDTH_1
            Me.Height = 3800
            CmdCancel.Top = Frame3.Top + Frame3.Height + 300
            CmdCancel.Left = M_LNG_FRMWIDTH_1 - CmdCancel.Width - 200
            CmdOK.Top = CmdCancel.Top
            CmdOK.Left = CmdCancel.Left - CmdOK.Width - 50
            CmdHelp.Top = CmdCancel.Top
    End Select
'    cmd打印设置.Top = IIf(mlngMode = 1343, cmd打印设置.Top, Cbo指定单位.Top)
    
    frm参数设置.Show vbModal, frmParent
End Sub

Private Sub cmd打印设置_Click()
    Dim strBill As String
    Select Case mstrFunction
    Case "药品申领管理"
        strBill = Split(cbo报表.Text, "(")(0)
    Case "协定药品入库"
        strBill = "ZL1_BILL_1344"
    End Select
    Call ReportPrintSet(gcnOracle, glngSys, strBill, Me)
End Sub

Private Sub Form_Load()
    Me.cmd打印设置.Caption = "票据《" & Replace(mstrFunction, "管理", "") & "单》打印设置"
    chkPrintCode.Visible = mlngMode = 1343
    '下拉列表设置
    cbo报表.Visible = mlngMode = 1343
    cbo报表.AddItem "ZL1_BILL_1304(单据打印)"
    cbo报表.AddItem "ZL1_INSIDE_1343_1(药品条码打印)"
    cbo报表.ListIndex = 0
End Sub

