VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmParaset 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6870
   Icon            =   "frmParaset.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5400
      TabIndex        =   25
      Top             =   5760
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   4200
      TabIndex        =   24
      Top             =   5760
      Width           =   1100
   End
   Begin TabDlg.SSTab tabMain 
      Height          =   5535
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   9763
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "基本(&0)"
      TabPicture(0)   =   "frmParaset.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra其他控制"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fra排序"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fra卫材单位"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fra打印控制"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fra其他"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.Frame fra其他 
         Caption         =   " 其他控制"
         Height          =   4920
         Left            =   4080
         TabIndex        =   18
         Top             =   480
         Width           =   2400
         Begin VB.CheckBox chk移库流程 
            Caption         =   "移库启用备料、发送、接收环节"
            Height          =   375
            Left            =   120
            TabIndex        =   26
            Top             =   720
            Visible         =   0   'False
            Width           =   2865
         End
         Begin VB.Frame fra查询天数 
            BorderStyle     =   0  'None
            Height          =   450
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   1955
            Begin VB.TextBox txt查询天数 
               Height          =   300
               Left            =   840
               TabIndex        =   20
               Text            =   "7"
               Top             =   60
               Width           =   300
            End
            Begin MSComCtl2.UpDown upd查询天数 
               Height          =   300
               Left            =   1080
               TabIndex        =   21
               Top             =   60
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   529
               _Version        =   393216
               Value           =   1
               BuddyControl    =   "txt查询天数"
               BuddyDispid     =   196614
               OrigLeft        =   1800
               OrigTop         =   360
               OrigRight       =   2055
               OrigBottom      =   735
               Max             =   90
               Min             =   1
               SyncBuddy       =   -1  'True
               BuddyProperty   =   65547
               Enabled         =   -1  'True
            End
            Begin VB.Label lbl天数 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "天"
               ForeColor       =   &H80000008&
               Height          =   180
               Left            =   1440
               TabIndex        =   23
               Top             =   120
               Width           =   180
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
               Left            =   0
               TabIndex        =   22
               Top             =   120
               Width           =   720
            End
         End
         Begin VB.Label lbl移库流程 
            Caption         =   "注意：如果不打勾，那么在填写移库单后，增加一个审核操作，审核后自动完成备料、发送、接收这一过程"
            ForeColor       =   &H00000080&
            Height          =   900
            Left            =   120
            TabIndex        =   27
            Top             =   1095
            Visible         =   0   'False
            Width           =   2865
         End
      End
      Begin VB.Frame fra打印控制 
         Caption         =   " 打印控制"
         Height          =   1215
         Left            =   120
         TabIndex        =   12
         Top             =   4200
         Width           =   3675
         Begin VB.CheckBox chkVerifyPrint 
            Caption         =   "审核后打印"
            Height          =   255
            Left            =   1560
            TabIndex        =   15
            Top             =   240
            Width           =   1455
         End
         Begin VB.CheckBox chkSavePrint 
            Caption         =   "存盘后打印"
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdPrintSet 
            Caption         =   "单据打印设置(&S)"
            Height          =   350
            Left            =   360
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   600
            Width           =   2925
         End
      End
      Begin VB.Frame fra卫材单位 
         Caption         =   " 卫材单位"
         Height          =   1665
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   3675
         Begin VB.ComboBox cboUnit 
            Height          =   300
            Left            =   870
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   390
            Width           =   2655
         End
         Begin VB.ComboBox CboUnit1 
            Height          =   300
            Left            =   870
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   780
            Width           =   2655
         End
         Begin VB.Label Label2 
            Caption         =   "注：请选择一种卫材单位，所有卫材将使用该单位进行包装显示和包装换算"
            ForeColor       =   &H00000080&
            Height          =   405
            Left            =   120
            TabIndex        =   11
            Top             =   1170
            Width           =   3315
         End
         Begin VB.Label lbl盘点表 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "盘点表"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   120
            TabIndex        =   10
            Top             =   450
            Width           =   540
         End
         Begin VB.Label lbl盘点单 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "盘点单"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   120
            TabIndex        =   9
            Top             =   840
            Width           =   540
         End
      End
      Begin VB.Frame fra排序 
         Caption         =   " 排序方式"
         Height          =   1785
         Left            =   120
         TabIndex        =   2
         Top             =   2250
         Width           =   3675
         Begin VB.ComboBox cbo列名 
            Height          =   300
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   390
            Width           =   2415
         End
         Begin VB.ComboBox cbo方向 
            Height          =   300
            Left            =   2580
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   390
            Width           =   885
         End
         Begin VB.Label lbl排序说明 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "注：本参数的设置，将影响所有编辑窗体中单据的显示内容的排序方式。缺省：按用户输入的顺序显示各单据的内容"
            ForeColor       =   &H00000080&
            Height          =   600
            Left            =   120
            TabIndex        =   5
            Top             =   960
            Width           =   3345
         End
      End
      Begin VB.Frame fra其他控制 
         Caption         =   " 其他控制"
         Height          =   1785
         Left            =   120
         TabIndex        =   16
         Top             =   2250
         Width           =   3675
         Begin VB.CheckBox chk申领核查 
            Caption         =   "申领需要核查后才能移库"
            Height          =   375
            Left            =   120
            TabIndex        =   17
            Top             =   360
            Width           =   3105
         End
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   360
      TabIndex        =   0
      Top             =   5760
      Width           =   1100
   End
End
Attribute VB_Name = "frmParaset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrFunction As String
Private mlngModule As Long '
Private mstrPrivs As String '
Private mblnHavePriv As Boolean
Private mblnFirstLoad As Boolean    '记录是否第一次加载
Private mfrmMain As Object '父窗体


Private Sub Cbo列名_Click()
    If cbo方向.ListCount < 1 Then Exit Sub
    cbo方向.Enabled = Not (cbo列名.ListIndex = 0)
    If Not cbo方向.Enabled Then cbo方向.ListIndex = 0
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
End Sub

Private Function SaveSet() As Boolean
    '------------------------------------------------------------------------------------------
    '功能:向数据库保存参数设置
    '返回:保存成功返回True,否则返回False
    '编制:刘兴宏
    '日期:2007/12/24
    '------------------------------------------------------------------------------------------
    err = 0: On Error GoTo ErrHand:
    If txt查询天数.Text > 7 Then
        If MsgBox("查询天数大于7天了，可能进入主页面会很慢，是否继续？", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Function
        End If
    End If
    
    gcnOracle.BeginTrans
    
    Call zlDatabase.SetPara(IIf(mlngModule = 1719, "盘点表单位", "卫材单位"), cboUnit.ListIndex, glngSys, mlngModule)
    If CboUnit1.Visible Then
        Call zlDatabase.SetPara("记录单单位", CboUnit1.ListIndex, glngSys, mlngModule)
    End If
    
    '申领没有单据排序参数
    If mlngModule <> 1722 Then
        Call zlDatabase.SetPara("单据排序", CStr(cbo列名.ListIndex) & CStr(cbo方向.ListIndex), glngSys, mlngModule)
    End If
    
    Call zlDatabase.SetPara("存盘打印", IIf(chkSavePrint.Value = 1, 1, 0), glngSys, mlngModule)
    
    '申领没有审核打印参数
    If mlngModule <> 1722 Then
        Call zlDatabase.SetPara("审核打印", IIf(chkVerifyPrint.Value = 1, 1, 0), glngSys, mlngModule)
    End If
    
    '卫材申领管理单独的参数
    If mlngModule = 1722 Then
        Call zlDatabase.SetPara("申领需要核查后才能移库", IIf(chk申领核查.Value = 1, 1, 0), glngSys, mlngModule)
    End If
    zlDatabase.SetPara "查询天数", Val(txt查询天数.Text), glngSys, mlngModule
    
    '卫材移库管理单独的参数
    If mlngModule = 1716 Then
        Call zlDatabase.SetPara("移库流程", IIf(chk移库流程.Value = 1, 1, 0), glngSys, mlngModule, , mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex))
    End If
    
    gcnOracle.CommitTrans
    SaveSet = True
    Exit Function
ErrHand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
End Function

Private Sub cmdOk_Click()
    If SaveSet = False Then Exit Sub
    Unload Me
End Sub

Private Sub initPara()
    '-------------------------------------------------------------------------------------------------------------
    '功能:初始化参数设置
    '返回:
    '编制:刘兴宏
    '修改:2007/12/24
    '-------------------------------------------------------------------------------------------------------------
    Dim strValue As String
    Dim strBidMess As String
    Dim int查询天数 As Integer
    
    '装入缺省数据
    With cbo列名
        .Clear
        .AddItem "输入顺序"
        .ItemData(.NewIndex) = 0
        .AddItem "编码"
        .ItemData(.NewIndex) = 1
        .AddItem "卫材名称"
        .ItemData(.NewIndex) = 2
        If mstrFunction = "卫材盘点管理" Then
            .AddItem "库房货位"
            .ItemData(.NewIndex) = 3
        End If
        .ListIndex = 0
    End With
    
    With cbo方向
        .Clear
        .AddItem "升序"
        .ItemData(.NewIndex) = 0
        .AddItem "降序"
        .ItemData(.NewIndex) = 1
        .ListIndex = 0
    End With
    
    If mlngModule <> 1722 Then
        strValue = zlDatabase.GetPara("单据排序", glngSys, mlngModule, "00", Array(cbo列名, cbo方向, fra排序, lbl排序说明), mblnHavePriv)
        strValue = IIf(strValue = "", "00", strValue)
        cbo列名.ListIndex = Val(Mid(strValue, 1, 1))
        cbo方向.ListIndex = Val(Right(strValue, 1))
        cbo方向.Enabled = Not (cbo列名.ListIndex = 0)
    End If
    
    chkSavePrint.Value = IIf(Val(zlDatabase.GetPara("存盘打印", glngSys, mlngModule, "0", Array(chkSavePrint), mblnHavePriv)) = 1, 1, 0)
    chkVerifyPrint.Value = IIf(Val(zlDatabase.GetPara("审核打印", glngSys, mlngModule, "0", Array(chkVerifyPrint), mblnHavePriv)) = 1, 1, 0)
    
    With CboUnit1
        .Clear
        .AddItem "散装单位"
        .AddItem "包装单位"
    End With

    With cboUnit
        .Clear
        .AddItem "散装单位"
        .AddItem "包装单位"
    End With
    cboUnit.ListIndex = IIf(Val(zlDatabase.GetPara(IIf(mlngModule = 1719, "盘点表单位", "卫材单位"), glngSys, mlngModule, "0", Array(cboUnit, lbl盘点表), mblnHavePriv)) = 1, 1, 0)
    If mstrFunction <> "卫材盘点管理" Then
        CboUnit1.Visible = False
        lbl盘点表.Visible = False
        lbl盘点单.Visible = False
        cboUnit.Left = lbl盘点表.Left
        Label2.Top = lbl盘点单.Top
    Else
        CboUnit1.ListIndex = IIf(Val(zlDatabase.GetPara("记录单单位", glngSys, mlngModule, "0", Array(CboUnit1, lbl盘点单), mblnHavePriv)) = 1, 1, 0)
    End If
    
    int查询天数 = Val(zlDatabase.GetPara("查询天数", glngSys, mlngModule, 1))
    txt查询天数.Text = int查询天数
    
    fra其他控制.Visible = False
    Select Case mstrFunction
        Case "卫材移库管理"
            Me.Width = Me.Width + 700 '改变宽度
            tabMain.Width = tabMain.Width + 700
            fra其他.Width = fra其他.Width + 700
            
            cmdOK.Left = cmdOK.Left + 700
            cmdCancel.Left = cmdCancel.Left + 700
            
            '设置可见
            chk移库流程.Visible = True
            lbl移库流程.Visible = True
            
            chk移库流程.Value = Val(zlDatabase.GetPara("移库流程", glngSys, mlngModule, "0", , , , mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)))
            
        Case "卫材盘点管理"

        Case "卫材外购入库管理"

        Case "卫材计划管理"
        
        Case "卫材申购管理"
            
        Case "卫材领用管理"
        
        Case "卫材申领管理"
            fra排序.Visible = False
            fra其他控制.Visible = True
            fra打印控制.Top = fra排序.Top
            fra其他控制.Top = fra打印控制.Top + fra打印控制.Height + 150
            chkVerifyPrint.Visible = False
            chk申领核查.Value = IIf((zlDatabase.GetPara("申领需要核查后才能移库", glngSys, mlngModule, "0")) = 0, 0, 1)
        Case "卫材调价管理"
            fra排序.Visible = False
            fra打印控制.Visible = False
            
            fra其他.Height = fra卫材单位.Height
            tabMain.Height = fra其他.Top + fra其他.Height + 200
            
            Me.Height = tabMain.Top + tabMain.Height + cmdHelp.Height + 650
            cmdHelp.Top = tabMain.Top + tabMain.Height + 100
            cmdCancel.Top = cmdHelp.Top
            cmdCancel.Left = Me.Width - cmdCancel.Width - 200
            cmdOK.Top = cmdHelp.Top
            cmdOK.Left = cmdCancel.Left - cmdOK.Width - 100
        Case Else
    End Select

    Me.cmdPrintSet.Enabled = InStr(1, gstrPrivs, ";单据打印;") <> 0

End Sub

Public Sub 设置参数(ByVal lngModule As Long, ByVal strPrivs As String, ByVal frmMain As Form, Optional ByVal strFunction As String = "")
    '-------------------------------------------------------------------------------------------------------------
    '功能:设置相关单据操作的控制参数
    '参数:lngModule-模块号
    '     str权限串-权限串
    '     frmMain-调用的主窗体
    '     strFunction-功能说明
    '返回:
    '编制:刘兴宏
    '修改:2007/12/24
    '-------------------------------------------------------------------------------------------------------------
    mstrPrivs = strPrivs: mlngModule = lngModule: mstrFunction = strFunction
    mblnHavePriv = zlStr.IsHavePrivs(mstrPrivs, "参数设置")
    Set mfrmMain = frmMain
    
    Call initPara
    frmParaset.Show vbModal, frmMain
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub cmdPrintSet_Click()
    Dim strBill As String
    strBill = "ZL1_BILL_" & glngModul
    Call ReportPrintSet(gcnOracle, glngSys, strBill, Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnFirstLoad = False
End Sub

