VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm参数设置 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8700
   Icon            =   "frm参数设置.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleMode       =   0  'User
   ScaleWidth      =   8700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin TabDlg.SSTab tabMain 
      Height          =   5895
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   10398
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "基本(&0)"
      TabPicture(0)   =   "frm参数设置.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra排序方式"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fra其他"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fra药品单位"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fra打印控制"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin VB.Frame fra打印控制 
         Caption         =   " 打印控制"
         Height          =   1980
         Left            =   120
         TabIndex        =   20
         Top             =   3840
         Width           =   4000
         Begin VB.ComboBox cbo报表 
            Height          =   300
            Left            =   480
            TabIndex        =   49
            Text            =   "Combo1"
            Top             =   1200
            Width           =   3135
         End
         Begin VB.CheckBox chkPrintCode 
            Caption         =   "存盘或审核后打印药品条码"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   840
            Width           =   3015
         End
         Begin VB.CommandButton cmd打印设置 
            Caption         =   "打印设置(&P)"
            Height          =   315
            Left            =   480
            TabIndex        =   24
            Top             =   1560
            Width           =   3135
         End
         Begin VB.CheckBox chkSendPrint 
            Caption         =   "发送后打印单据"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   480
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.CheckBox chkSavePrint 
            Caption         =   "存盘后打印单据"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   1635
         End
         Begin VB.CheckBox chkVerifyPrint 
            Caption         =   "审核后打印单据"
            Height          =   255
            Left            =   2040
            TabIndex        =   21
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame fra药品单位 
         Caption         =   " 药品单位"
         Height          =   1785
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   4000
         Begin VB.ComboBox cboUnit 
            Height          =   300
            Left            =   750
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   390
            Width           =   2655
         End
         Begin VB.ComboBox CboUnit1 
            Height          =   300
            Left            =   750
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   780
            Width           =   2655
         End
         Begin VB.Label lblUnitComment 
            Caption         =   "注：请选择一种药品单位，所有药品将使用该单位进行包装显示和包装换算"
            ForeColor       =   &H00000080&
            Height          =   405
            Left            =   120
            TabIndex        =   18
            Top             =   1200
            Width           =   3315
         End
         Begin VB.Label lbl盘点表 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "大包装"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   120
            TabIndex        =   17
            Top             =   450
            Width           =   540
         End
         Begin VB.Label lbl盘点单 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "小包装"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   120
            TabIndex        =   16
            Top             =   840
            Width           =   540
         End
      End
      Begin VB.Frame fra其他 
         Caption         =   " 其他控制"
         Height          =   5355
         Left            =   4200
         TabIndex        =   8
         Top             =   480
         Width           =   4200
         Begin VB.Frame frm对方库存 
            Caption         =   "填单时对方库房库存显示方式"
            Height          =   735
            Left            =   120
            TabIndex        =   45
            Top             =   1920
            Width           =   3975
            Begin VB.OptionButton opt对方库存 
               Caption         =   "显示可用数量"
               Height          =   255
               Index           =   1
               Left            =   1680
               TabIndex        =   47
               Top             =   360
               Width           =   1575
            End
            Begin VB.OptionButton opt对方库存 
               Caption         =   "显示实际数量"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   46
               Top             =   360
               Value           =   -1  'True
               Width           =   1575
            End
         End
         Begin VB.Frame frm当前库存 
            Caption         =   "填单时当前库房库存显示方式"
            Height          =   735
            Left            =   120
            TabIndex        =   42
            Top             =   960
            Width           =   3975
            Begin VB.OptionButton opt当前库存 
               Caption         =   "显示可用数量"
               Height          =   255
               Index           =   1
               Left            =   1680
               TabIndex        =   44
               Top             =   360
               Width           =   1455
            End
            Begin VB.OptionButton opt当前库存 
               Caption         =   "显示实际数量"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   43
               Top             =   360
               Value           =   -1  'True
               Width           =   1575
            End
         End
         Begin VB.CheckBox chkALLPlanPoint 
            Caption         =   "全院计划不管站点"
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   4200
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.Frame fra药品计划供应商设置 
            Caption         =   " 药品采购计划供应商设置"
            Height          =   1725
            Left            =   120
            TabIndex        =   35
            Top             =   2400
            Width           =   3960
            Begin VB.ComboBox cbo供应商范围 
               Height          =   300
               Left            =   960
               Style           =   2  'Dropdown List
               TabIndex        =   37
               Top             =   660
               Width           =   2700
            End
            Begin VB.ComboBox cbo供应商选择 
               Height          =   300
               Left            =   960
               Style           =   2  'Dropdown List
               TabIndex        =   36
               Top             =   300
               Width           =   2700
            End
            Begin VB.Label Label3 
               Caption         =   "注：药品采购计划编辑界面中药品供应商的默认处理，以及手工选择供应商时的可选范围"
               ForeColor       =   &H00000080&
               Height          =   495
               Left            =   120
               TabIndex        =   40
               Top             =   1080
               Width           =   3765
            End
            Begin VB.Label lbl供应商范围 
               AutoSize        =   -1  'True
               Caption         =   "选择范围"
               Height          =   180
               Left            =   120
               TabIndex        =   39
               Top             =   720
               Width           =   720
            End
            Begin VB.Label lbl供应商选择 
               AutoSize        =   -1  'True
               Caption         =   "默认选择"
               Height          =   180
               Left            =   120
               TabIndex        =   38
               Top             =   360
               Width           =   720
            End
         End
         Begin VB.Frame fra查询天数 
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            Height          =   450
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Width           =   2175
            Begin VB.ComboBox cboDay 
               Height          =   300
               Left            =   840
               Style           =   2  'Dropdown List
               TabIndex        =   50
               Top             =   60
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.TextBox txt查询天数 
               Height          =   300
               Left            =   840
               TabIndex        =   30
               Text            =   "7"
               Top             =   60
               Width           =   300
            End
            Begin MSComCtl2.UpDown upd查询天数 
               Height          =   300
               Left            =   1140
               TabIndex        =   31
               Top             =   60
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   529
               _Version        =   393216
               Value           =   1
               BuddyControl    =   "txt查询天数"
               BuddyDispid     =   196636
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
               TabIndex        =   33
               Top             =   120
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
               Left            =   1440
               TabIndex        =   32
               Top             =   120
               Width           =   180
            End
         End
         Begin VB.Frame fra盘点时间范围 
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   0  'None
            Height          =   450
            Left            =   120
            TabIndex        =   25
            Top             =   720
            Visible         =   0   'False
            Width           =   1700
            Begin VB.TextBox txt盘点时间 
               Height          =   300
               Left            =   840
               TabIndex        =   26
               Text            =   "3"
               Top             =   60
               Width           =   300
            End
            Begin MSComCtl2.UpDown UpD盘点时间 
               Height          =   300
               Left            =   1140
               TabIndex        =   27
               Top             =   60
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   529
               _Version        =   393216
               BuddyControl    =   "txt盘点时间"
               BuddyDispid     =   196641
               OrigLeft        =   1800
               OrigTop         =   360
               OrigRight       =   2055
               OrigBottom      =   735
               Max             =   90
               SyncBuddy       =   -1  'True
               BuddyProperty   =   65547
               Enabled         =   -1  'True
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "盘点时间"
               ForeColor       =   &H80000008&
               Height          =   180
               Left            =   0
               TabIndex        =   34
               Top             =   120
               Width           =   720
            End
            Begin VB.Label lblday 
               BackStyle       =   0  'Transparent
               Caption         =   "天"
               Height          =   195
               Left            =   1440
               TabIndex        =   28
               Top             =   120
               Width           =   255
            End
         End
         Begin VB.CheckBox chk留存领用 
            Caption         =   "按月留存领用"
            Height          =   255
            Left            =   2280
            TabIndex        =   19
            Top             =   360
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Frame fra药品计划价格显示方式 
            Caption         =   " 药品采购计划价格显示方式"
            Height          =   735
            Left            =   120
            TabIndex        =   9
            Top             =   1320
            Visible         =   0   'False
            Width           =   3960
            Begin VB.OptionButton Opt混合 
               Caption         =   "成本价和售价"
               Height          =   180
               Left            =   2160
               TabIndex        =   11
               Top             =   375
               Width           =   1400
            End
            Begin VB.OptionButton Opt成本价 
               Caption         =   "成本价"
               Height          =   180
               Left            =   120
               TabIndex        =   12
               Top             =   375
               Width           =   900
            End
            Begin VB.OptionButton Opt售价 
               Caption         =   "售价"
               Height          =   180
               Left            =   1200
               TabIndex        =   10
               Top             =   375
               Width           =   720
            End
         End
      End
      Begin VB.Frame fra排序方式 
         Caption         =   " 排序方式"
         Height          =   1515
         Left            =   120
         TabIndex        =   4
         Top             =   2280
         Width           =   4000
         Begin VB.ComboBox Cbo列名 
            Height          =   300
            ItemData        =   "frm参数设置.frx":0028
            Left            =   120
            List            =   "frm参数设置.frx":002A
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   390
            Width           =   2415
         End
         Begin VB.ComboBox Cbo方向 
            Height          =   300
            Left            =   2580
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   390
            Width           =   885
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "注：本参数的设置，将影响所有编辑窗体中单据的显示内容的排序方式。缺省：按用户输入的顺序显示各单据的内容"
            ForeColor       =   &H00000080&
            Height          =   675
            Left            =   120
            TabIndex        =   7
            Top             =   840
            Width           =   3345
         End
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7200
      TabIndex        =   2
      Top             =   6240
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   6000
      TabIndex        =   1
      Top             =   6240
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   240
      TabIndex        =   0
      Top             =   6240
      Width           =   1100
   End
End
Attribute VB_Name = "frm参数设置"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrFunction As String
Dim mstrPrivs As String
Dim mlngModul As Long
Dim mblnSetPara As Boolean      '是否具有参数设置权限
Private mint盘点时间 As Integer  '用来记录设置的盘点时间范围

Private Sub Cbo列名_Click()
    If Cbo方向.ListCount < 1 Then Exit Sub
    Cbo方向.Enabled = Not (Cbo列名.ListIndex = 0)
    If Not Cbo方向.Enabled Then Cbo方向.ListIndex = 0
End Sub

Private Sub chkSavePrint_Click()
    chkPrintCode.Enabled = chkVerifyPrint.Value = 1 Or chkSavePrint.Value = 1 Or chkSendPrint.Value = 1
End Sub

Private Sub chkSendPrint_Click()
    chkPrintCode.Enabled = chkVerifyPrint.Value = 1 Or chkSavePrint.Value = 1 Or chkSendPrint.Value = 1
End Sub

Private Sub chkVerifyPrint_Click()
    chkPrintCode.Enabled = chkVerifyPrint.Value = 1 Or chkSavePrint.Value = 1 Or chkSendPrint.Value = 1
End Sub

Private Sub chk留存领用_Click()
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errH
    
    If chk留存领用.Value = 0 Then
        gstrSQL = "Select 期间 From 药品留存 Where Length(期间) > 4"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        If rsTemp.RecordCount > 0 Then
            MsgBox "按月留存模式下已经产生数据，不能修改！", vbInformation, gstrSysName
            chk留存领用.Value = 1
        End If
    End If
    Exit Sub
errH:
If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub cmdOK_Click()
    On Error Resume Next
    
    If ISValid = False Then Exit Sub
    
    Select Case mlngModul
        Case 1300   '药品外购入库管理
            zldatabase.SetPara "排序", CStr(Cbo列名.ListIndex) & CStr(Cbo方向.ListIndex), glngSys, mlngModul
            zldatabase.SetPara "存盘打印", IIf(chkSavePrint.Value = 1, "1", "0"), glngSys, mlngModul
            zldatabase.SetPara "审核打印", IIf(chkVerifyPrint.Value = 1, "1", "0"), glngSys, mlngModul
            zldatabase.SetPara "打印药品条码", IIf(chkPrintCode.Value = 1, "1", "0"), glngSys, mlngModul
            zldatabase.SetPara "药品单位", cboUnit.ListIndex, glngSys, mlngModul
            
            zldatabase.SetPara "查询天数", Val(txt查询天数.Text), glngSys, mlngModul
        Case 1301   '药品自制入库管理
            zldatabase.SetPara "排序", CStr(Cbo列名.ListIndex) & CStr(Cbo方向.ListIndex), glngSys, mlngModul
            zldatabase.SetPara "存盘打印", IIf(chkSavePrint.Value = 1, "1", "0"), glngSys, mlngModul
            zldatabase.SetPara "审核打印", IIf(chkVerifyPrint.Value = 1, "1", "0"), glngSys, mlngModul
            zldatabase.SetPara "药品单位", cboUnit.ListIndex, glngSys, mlngModul
            zldatabase.SetPara "查询天数", Val(txt查询天数.Text), glngSys, mlngModul
        Case 1302   '药品其他入库管理
            zldatabase.SetPara "排序", CStr(Cbo列名.ListIndex) & CStr(Cbo方向.ListIndex), glngSys, mlngModul
            zldatabase.SetPara "存盘打印", IIf(chkSavePrint.Value = 1, "1", "0"), glngSys, mlngModul
            zldatabase.SetPara "审核打印", IIf(chkVerifyPrint.Value = 1, "1", "0"), glngSys, mlngModul
            zldatabase.SetPara "打印药品条码", IIf(chkPrintCode.Value = 1, "1", "0"), glngSys, mlngModul
            zldatabase.SetPara "药品单位", cboUnit.ListIndex, glngSys, mlngModul
            zldatabase.SetPara "查询天数", Val(txt查询天数.Text), glngSys, mlngModul
        Case 1303   '药品库存差价调整管理
            zldatabase.SetPara "排序", CStr(Cbo列名.ListIndex) & CStr(Cbo方向.ListIndex), glngSys, mlngModul
            zldatabase.SetPara "存盘打印", IIf(chkSavePrint.Value = 1, "1", "0"), glngSys, mlngModul
            zldatabase.SetPara "审核打印", IIf(chkVerifyPrint.Value = 1, "1", "0"), glngSys, mlngModul
            zldatabase.SetPara "药品单位", cboUnit.ListIndex, glngSys, mlngModul
            zldatabase.SetPara "查询天数", Val(txt查询天数.Text), glngSys, mlngModul
        Case 1304   '药品移库管理
            zldatabase.SetPara "药品单位", cboUnit.ListIndex, glngSys, mlngModul
            zldatabase.SetPara "排序", CStr(Cbo列名.ListIndex) & CStr(Cbo方向.ListIndex), glngSys, mlngModul
            zldatabase.SetPara "存盘打印", IIf(chkSavePrint.Value = 1, "1", "0"), glngSys, mlngModul
            zldatabase.SetPara "审核打印", IIf(chkVerifyPrint.Value = 1, "1", "0"), glngSys, mlngModul
            zldatabase.SetPara "发送打印", IIf(chkSendPrint.Value = 1, "1", "0"), glngSys, mlngModul
            zldatabase.SetPara "打印药品条码", IIf(chkPrintCode.Value = 1, "1", "0"), glngSys, mlngModul
            zldatabase.SetPara "查询天数", Val(txt查询天数.Text), glngSys, mlngModul
            zldatabase.SetPara "填单时当前库房库存显示方式", IIf(opt当前库存(0).Value = True, 0, 1), glngSys, mlngModul
            zldatabase.SetPara "填单时对方库房库存显示方式", IIf(opt对方库存(0).Value = True, 0, 1), glngSys, mlngModul
        Case 1305   '药品领用管理
            zldatabase.SetPara "药品单位", cboUnit.ListIndex, glngSys, mlngModul
            zldatabase.SetPara "排序", CStr(Cbo列名.ListIndex) & CStr(Cbo方向.ListIndex), glngSys, mlngModul
            zldatabase.SetPara "存盘打印", IIf(chkSavePrint.Value = 1, "1", "0"), glngSys, mlngModul
            zldatabase.SetPara "审核打印", IIf(chkVerifyPrint.Value = 1, "1", "0"), glngSys, mlngModul
            zldatabase.SetPara "打印药品条码", IIf(chkPrintCode.Value = 1, "1", "0"), glngSys, mlngModul
            zldatabase.SetPara "按月留存领用", IIf(chk留存领用.Value = 1, "1", "0"), glngSys, mlngModul
            zldatabase.SetPara "查询天数", Val(txt查询天数.Text), glngSys, mlngModul
        Case 1306   '药品其他出库管理
            zldatabase.SetPara "排序", CStr(Cbo列名.ListIndex) & CStr(Cbo方向.ListIndex), glngSys, mlngModul
            zldatabase.SetPara "存盘打印", IIf(chkSavePrint.Value = 1, "1", "0"), glngSys, mlngModul
            zldatabase.SetPara "审核打印", IIf(chkVerifyPrint.Value = 1, "1", "0"), glngSys, mlngModul
            zldatabase.SetPara "打印药品条码", IIf(chkPrintCode.Value = 1, "1", "0"), glngSys, mlngModul
            zldatabase.SetPara "药品单位", cboUnit.ListIndex, glngSys, mlngModul
            zldatabase.SetPara "查询天数", Val(txt查询天数.Text), glngSys, mlngModul
        Case 1307   '药品盘点管理
            zldatabase.SetPara "排序", CStr(Cbo列名.ListIndex) & CStr(Cbo方向.ListIndex), glngSys, mlngModul
            zldatabase.SetPara "小包装单位", CboUnit1.ListIndex, glngSys, mlngModul
            zldatabase.SetPara "存盘打印", IIf(chkSavePrint.Value = 1, "1", "0"), glngSys, mlngModul
            zldatabase.SetPara "审核打印", IIf(chkVerifyPrint.Value = 1, "1", "0"), glngSys, mlngModul
            zldatabase.SetPara "查询天数", Val(cboDay.ItemData(cboDay.ListIndex)), glngSys, mlngModul

            zldatabase.SetPara "盘点时间范围设置", txt盘点时间.Text, glngSys, mlngModul
        Case 1330   '药品计划管理
            zldatabase.SetPara "排序", CStr(Cbo列名.ListIndex) & CStr(Cbo方向.ListIndex), glngSys, mlngModul
            zldatabase.SetPara "价格显示方式", IIf(Opt成本价.Value = True, "0", IIf(Opt售价.Value = True, "1", "2")), glngSys, mlngModul
            zldatabase.SetPara "存盘打印", IIf(chkSavePrint.Value = 1, "1", "0"), glngSys, mlngModul
            zldatabase.SetPara "审核打印", IIf(chkVerifyPrint.Value = 1, "1", "0"), glngSys, mlngModul
            zldatabase.SetPara "供应商默认选择", cbo供应商选择.ListIndex, glngSys, mlngModul
            zldatabase.SetPara "供应商选择范围", cbo供应商范围.ListIndex, glngSys, mlngModul
            zldatabase.SetPara "药品单位", cboUnit.ListIndex, glngSys, mlngModul
            zldatabase.SetPara "查询天数", Val(txt查询天数.Text), glngSys, mlngModul
            zldatabase.SetPara "全院计划不管站点", IIf(chkALLPlanPoint.Value = 1, "1", "0"), glngSys, mlngModul
        Case 1331   '药品质量管理
            zldatabase.SetPara "查询天数", Val(txt查询天数.Text), glngSys, mlngModul
        Case 1333 '药品调价管理
            zldatabase.SetPara "排序", CStr(Cbo列名.ListIndex) & CStr(Cbo方向.ListIndex), glngSys, mlngModul
            zldatabase.SetPara "药品单位", cboUnit.ListIndex, glngSys, mlngModul
            zldatabase.SetPara "查询天数", Val(txt查询天数.Text), glngSys, mlngModul
    End Select
           
    Unload Me
End Sub

Private Function ISValid() As Boolean
    Dim i As Integer
    
    If Val(txt查询天数.Text) > 7 Then
        If MsgBox("查询时间大于7天可能会导致查询很慢，是否继续？", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            txt查询天数.SetFocus
            zlControl.TxtSelAll txt查询天数
            Exit Function
        End If
    End If
    If Val(txt查询天数.Text) = 0 Then
        MsgBox "查询时间必须大于0，请重新输入！", vbInformation, gstrSysName
        txt查询天数.SetFocus
        zlControl.TxtSelAll txt查询天数
        Exit Function
    End If
    
    ISValid = True
End Function

Private Sub loadCboDay()
    With cboDay
        .AddItem "显示今日"
        .ItemData(.NewIndex) = 1
        .AddItem "显示7天之内"
        .ItemData(.NewIndex) = 7
        
        .Visible = True
        lbl查询天数.Caption = "查询范围"
        txt查询天数.Visible = False
    End With
End Sub

Public Sub 设置参数(frmParent As Object, ByVal strPrivs As String, Optional ByVal strFunction As String = "")
    mstrFunction = strFunction
    mstrPrivs = strPrivs
    mlngModul = glngModul
    Dim str单据打印 As String
    
    '通用（私有模块）
    Dim str排序 As String
    Dim int存盘打印 As Integer
    Dim int审核打印 As Integer
    Dim int打印药品条码 As Integer
        
    '用于主要流通模块（私有模块）
    Dim int药品单位 As Integer
    Dim int成本价来源 As Integer
    Dim int查询天数 As Integer
        
    '用于盘点（私有模块）
    Dim int小包装单位 As Integer
        
    '用于药品计划（私有模块）
    Dim int价格显示方式 As Integer
    Dim int供应商选择 As Integer
    Dim int供应商范围 As Integer
    Dim intPlanPoint As Integer
    
    '用于移库(私有)
    Dim int发送打印 As Integer
    Dim int当前库存 As Integer
    Dim int对方库存 As Integer
    
    '用于领用
    Dim int留存领用 As Integer
    
    Dim i As Integer
    
    On Error Resume Next
    
    mblnSetPara = zlStr.IsHavePrivs(mstrPrivs, "参数设置")
    
    '取参数值
    Select Case mlngModul
        Case 1300   '药品外购入库管理
            str排序 = zldatabase.GetPara("排序", glngSys, mlngModul, "00", Array(fra排序方式, Cbo列名, Cbo方向, Label5), mblnSetPara)
            int存盘打印 = Val(zldatabase.GetPara("存盘打印", glngSys, mlngModul, 0, Array(chkSavePrint), mblnSetPara))
            int审核打印 = Val(zldatabase.GetPara("审核打印", glngSys, mlngModul, 0, Array(chkVerifyPrint), mblnSetPara))
            int打印药品条码 = Val(zldatabase.GetPara("打印药品条码", glngSys, mlngModul, 0, Array(chkPrintCode), mblnSetPara))
            int药品单位 = Val(zldatabase.GetPara("药品单位", glngSys, mlngModul, 0, Array(lbl盘点表, cboUnit), mblnSetPara))
            int查询天数 = Val(zldatabase.GetPara("查询天数", glngSys, mlngModul, 7))
        Case 1301   '药品自制入库管理
            str排序 = zldatabase.GetPara("排序", glngSys, mlngModul, "00", Array(fra排序方式, Cbo列名, Cbo方向, Label5), mblnSetPara)
            int存盘打印 = Val(zldatabase.GetPara("存盘打印", glngSys, mlngModul, 0, Array(chkSavePrint), mblnSetPara))
            int审核打印 = Val(zldatabase.GetPara("审核打印", glngSys, mlngModul, 0, Array(chkVerifyPrint), mblnSetPara))
            int药品单位 = Val(zldatabase.GetPara("药品单位", glngSys, mlngModul, 0, Array(lbl盘点表, cboUnit), mblnSetPara))
            int查询天数 = Val(zldatabase.GetPara("查询天数", glngSys, mlngModul, 7))
        Case 1302   '药品其他入库管理
            str排序 = zldatabase.GetPara("排序", glngSys, mlngModul, "00", Array(fra排序方式, Cbo列名, Cbo方向, Label5), mblnSetPara)
            int存盘打印 = Val(zldatabase.GetPara("存盘打印", glngSys, mlngModul, 0, Array(chkSavePrint), mblnSetPara))
            int审核打印 = Val(zldatabase.GetPara("审核打印", glngSys, mlngModul, 0, Array(chkVerifyPrint), mblnSetPara))
            int打印药品条码 = Val(zldatabase.GetPara("打印药品条码", glngSys, mlngModul, 0, Array(chkPrintCode), mblnSetPara))
            int药品单位 = Val(zldatabase.GetPara("药品单位", glngSys, mlngModul, 0, Array(lbl盘点表, cboUnit), mblnSetPara))
            int查询天数 = Val(zldatabase.GetPara("查询天数", glngSys, mlngModul, 7))
        Case 1303   '药品库存差价调整管理
            str排序 = zldatabase.GetPara("排序", glngSys, mlngModul, "00", Array(fra排序方式, Cbo列名, Cbo方向, Label5), mblnSetPara)
            int存盘打印 = Val(zldatabase.GetPara("存盘打印", glngSys, mlngModul, 0, Array(chkSavePrint), mblnSetPara))
            int审核打印 = Val(zldatabase.GetPara("审核打印", glngSys, mlngModul, 0, Array(chkVerifyPrint), mblnSetPara))
            int药品单位 = Val(zldatabase.GetPara("药品单位", glngSys, mlngModul, 0, Array(lbl盘点表, cboUnit), mblnSetPara))
            int查询天数 = Val(zldatabase.GetPara("查询天数", glngSys, mlngModul, 7))
        Case 1304   '药品移库管理
            int药品单位 = Val(zldatabase.GetPara("药品单位", glngSys, mlngModul, 0, Array(lbl盘点表, cboUnit), mblnSetPara))
            str排序 = zldatabase.GetPara("排序", glngSys, mlngModul, "00", Array(fra排序方式, Cbo列名, Cbo方向, Label5), mblnSetPara)
            int存盘打印 = Val(zldatabase.GetPara("存盘打印", glngSys, mlngModul, 0, Array(chkSavePrint), mblnSetPara))
            int审核打印 = Val(zldatabase.GetPara("审核打印", glngSys, mlngModul, 0, Array(chkVerifyPrint), mblnSetPara))
            int发送打印 = Val(zldatabase.GetPara("发送打印", glngSys, mlngModul, 0, Array(chkSendPrint), mblnSetPara))
            int打印药品条码 = Val(zldatabase.GetPara("打印药品条码", glngSys, mlngModul, 0, Array(chkPrintCode), mblnSetPara))
            int查询天数 = Val(zldatabase.GetPara("查询天数", glngSys, mlngModul, 7))
            int当前库存 = Val(zldatabase.GetPara("填单时当前库房库存显示方式", glngSys, mlngModul, 0, Array(opt当前库存(0), opt当前库存(1)), mblnSetPara))
            int对方库存 = Val(zldatabase.GetPara("填单时对方库房库存显示方式", glngSys, mlngModul, 0, Array(opt对方库存(0), opt对方库存(1)), mblnSetPara))
        Case 1305   '药品领用管理
            int药品单位 = Val(zldatabase.GetPara("药品单位", glngSys, mlngModul, 0, Array(lbl盘点表, cboUnit), mblnSetPara))
            str排序 = zldatabase.GetPara("排序", glngSys, mlngModul, "00", Array(fra排序方式, Cbo列名, Cbo方向, Label5), mblnSetPara)
            int存盘打印 = Val(zldatabase.GetPara("存盘打印", glngSys, mlngModul, 0, Array(chkSavePrint), mblnSetPara))
            int审核打印 = Val(zldatabase.GetPara("审核打印", glngSys, mlngModul, 0, Array(chkVerifyPrint), mblnSetPara))
            int打印药品条码 = Val(zldatabase.GetPara("打印药品条码", glngSys, mlngModul, 0, Array(chkPrintCode), mblnSetPara))
            int留存领用 = Val(zldatabase.GetPara("按月留存领用", glngSys, mlngModul, 0, Array(chk留存领用), mblnSetPara))
            int查询天数 = Val(zldatabase.GetPara("查询天数", glngSys, mlngModul, 7))
        Case 1306   '药品其他出库管理
            str排序 = zldatabase.GetPara("排序", glngSys, mlngModul, "00", Array(fra排序方式, Cbo列名, Cbo方向, Label5), mblnSetPara)
            int存盘打印 = Val(zldatabase.GetPara("存盘打印", glngSys, mlngModul, 0, Array(chkSavePrint), mblnSetPara))
            int审核打印 = Val(zldatabase.GetPara("审核打印", glngSys, mlngModul, 0, Array(chkVerifyPrint), mblnSetPara))
            int打印药品条码 = Val(zldatabase.GetPara("打印药品条码", glngSys, mlngModul, 0, Array(chkPrintCode), mblnSetPara))
            int药品单位 = Val(zldatabase.GetPara("药品单位", glngSys, mlngModul, 0, Array(lbl盘点表, cboUnit), mblnSetPara))
            int查询天数 = Val(zldatabase.GetPara("查询天数", glngSys, mlngModul, 7))
        Case 1307   '药品盘点管理
            str排序 = zldatabase.GetPara("排序", glngSys, mlngModul, "00", Array(fra排序方式, Cbo列名, Cbo方向, Label5), mblnSetPara)
            int小包装单位 = Val(zldatabase.GetPara("小包装单位", glngSys, mlngModul, 0, Array(lbl盘点单, CboUnit1), mblnSetPara))
            int存盘打印 = Val(zldatabase.GetPara("存盘打印", glngSys, mlngModul, 0, Array(chkSavePrint), mblnSetPara))
            int审核打印 = Val(zldatabase.GetPara("审核打印", glngSys, mlngModul, 0, Array(chkVerifyPrint), mblnSetPara))
                        
            mint盘点时间 = Val(zldatabase.GetPara("盘点时间范围设置", glngSys, mlngModul, 30))
            txt盘点时间.Text = mint盘点时间
            UpD盘点时间.Value = mint盘点时间
            int查询天数 = Val(zldatabase.GetPara("查询天数", glngSys, mlngModul, 7))
        Case 1330   '药品计划管理
            str排序 = zldatabase.GetPara("排序", glngSys, mlngModul, "00", Array(fra排序方式, Cbo列名, Cbo方向, Label5), mblnSetPara)
            int价格显示方式 = Val(zldatabase.GetPara("价格显示方式", glngSys, mlngModul, 1, Array(fra药品计划价格显示方式, Opt成本价, Opt售价, Opt混合), mblnSetPara))
            int存盘打印 = Val(zldatabase.GetPara("存盘打印", glngSys, mlngModul, 0, Array(chkSavePrint), mblnSetPara))
            int审核打印 = Val(zldatabase.GetPara("审核打印", glngSys, mlngModul, 0, Array(chkVerifyPrint), mblnSetPara))
            int供应商选择 = Val(zldatabase.GetPara("供应商默认选择", glngSys, mlngModul, 0, Array(cbo供应商选择), mblnSetPara))
            int供应商范围 = Val(zldatabase.GetPara("供应商选择范围", glngSys, mlngModul, 0, Array(cbo供应商范围), mblnSetPara))
            int药品单位 = Val(zldatabase.GetPara("药品单位", glngSys, mlngModul, 0, Array(lbl盘点表, cboUnit), mblnSetPara))
            int查询天数 = Val(zldatabase.GetPara("查询天数", glngSys, mlngModul, 7))
            intPlanPoint = Val(zldatabase.GetPara("全院计划不管站点", glngSys, mlngModul, 0, Array(chkALLPlanPoint), mblnSetPara))
            chkALLPlanPoint.Value = intPlanPoint
        Case 1331  '药品质量管理
            int查询天数 = Val(zldatabase.GetPara("查询天数", glngSys, mlngModul, 7))
        Case 1333 '药品调价管理
            str排序 = zldatabase.GetPara("排序", glngSys, mlngModul, "00", Array(fra排序方式, Cbo列名, Cbo方向, Label5), mblnSetPara)
            int药品单位 = Val(zldatabase.GetPara("药品单位", glngSys, mlngModul, 0, Array(lbl盘点表, cboUnit), mblnSetPara))
            int查询天数 = Val(zldatabase.GetPara("查询天数", glngSys, mlngModul, 7))
    End Select
    
    If mlngModul <> 1307 Then
        txt查询天数.Text = int查询天数
    Else '盘点
        loadCboDay
        int查询天数 = IIf(int查询天数 <> 1 And int查询天数 <> 7, 7, int查询天数)
        For i = 0 To cboDay.ListCount - 1
            If int查询天数 = cboDay.ItemData(i) Then cboDay.ListIndex = i
        Next
    End If
    
    If strFunction = "药品计划管理" Then
        str单据打印 = "采购计划打印"
    Else
        str单据打印 = "单据打印"
    End If
    
    '装入缺省数据
    With Cbo列名
        .Clear
        .AddItem "输入顺序"
        .ItemData(.NewIndex) = 0
        .AddItem "编码"
        .ItemData(.NewIndex) = 1
        .AddItem "药品名称"
        .ItemData(.NewIndex) = 2
        
        If InStr("药品盘点管理/药品移库管理/药品领用管理/药品其他出库管理", strFunction) > 0 Then
            .AddItem "库房货位"
            .ItemData(.NewIndex) = 3
        End If
     
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
    Cbo列名.ListIndex = Mid(str排序, 1, 1)
    Cbo方向.ListIndex = Right(str排序, 1)
    Cbo方向.Enabled = Not (Cbo列名.ListIndex = 0)
    
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
    
    If int打印药品条码 = 0 Then
        chkPrintCode.Value = 0
    Else
        chkPrintCode.Value = 1
    End If
    
    chkPrintCode.Enabled = chkVerifyPrint.Value = 1 Or chkSavePrint.Value = 1
    
    If int留存领用 = 0 Then
        chk留存领用.Value = 0
    Else
        chk留存领用.Value = 1
    End If

    If mstrFunction = "药品盘点管理" Then
        If glngSys \ 100 = 8 Then
            With CboUnit1
                .AddItem "采购单位"
                .AddItem "售价单位"
            End With
        Else
            With CboUnit1
                .AddItem "和大包装相同"
                .AddItem "药库单位"
                .AddItem "门诊单位"
                .AddItem "住院单位"
                .AddItem "售价单位"
            End With
        End If
        CboUnit1.ListIndex = int小包装单位
        lblUnitComment.Caption = "    请选择盘点时的大小包装，盘点单及盘点表编辑时按所选包装进行盘点。"
    Else
        CboUnit1.Visible = False
        lbl盘点表.Visible = False
        lbl盘点单.Visible = False
        cboUnit.Left = lbl盘点表.Left
    End If
    
    With cboUnit
        .Clear
        If glngSys \ 100 = 8 Then
            .AddItem "缺省（当前库房对应的单位）"
            .AddItem "采购单位"
            .AddItem "售价单位"
        Else
            If mlngModul <> 1333 Then   '调价不需要库房
                .AddItem "缺省（当前库房对应的单位）"
            End If
            .AddItem "药库单位"
            .AddItem "门诊单位"
            .AddItem "住院单位"
            .AddItem "售价单位"
        End If
        .ListIndex = int药品单位
    End With
    
    '界面调整，根据模块显示或隐藏不同的模块参数设置
    chkSendPrint.Visible = False
    If strFunction = "药品移库管理" Then
        chkSendPrint.Value = IIf(int发送打印 = 1, 1, 0)
        chkSendPrint.Visible = True
        
        chkPrintCode.Enabled = chkPrintCode.Enabled Or chkSendPrint.Value = 1
        
        If int当前库存 = 1 Then
            opt当前库存(1).Value = True
        Else
            opt当前库存(0).Value = True
        End If
        
        If int对方库存 = 1 Then
            opt对方库存(1).Value = True
        Else
            opt对方库存(0).Value = True
        End If
    Else
        frm当前库存.Visible = False
        frm对方库存.Visible = False
    End If
    
    fra药品计划价格显示方式.Visible = False
    fra药品计划供应商设置.Visible = False
    chkALLPlanPoint.Visible = False
    If strFunction = "药品计划管理" Then
        If int价格显示方式 = 0 Then
            Opt成本价.Value = True
        ElseIf int价格显示方式 = 1 Then
            Opt售价.Value = True
        Else
            Opt混合.Value = True
        End If
        
        chkALLPlanPoint.Visible = True
        cbo供应商选择.Clear
        cbo供应商选择.AddItem "1-取上次入库供应商"
        cbo供应商选择.AddItem "2-取合同单位"
        cbo供应商选择.ListIndex = IIf(int供应商选择 < 0 Or int供应商选择 > 1, 0, int供应商选择)
        
        cbo供应商范围.Clear
        cbo供应商范围.AddItem "1-所有供应商"
        cbo供应商范围.AddItem "2-中标单位"
        cbo供应商范围.ListIndex = IIf(int供应商范围 < 0 Or int供应商范围 > 1, 0, int供应商范围)
        
        fra药品计划价格显示方式.Visible = True
        fra药品计划供应商设置.Visible = True
        
        fra药品计划价格显示方式.Top = fra查询天数.Top + fra查询天数.Height + 100
        fra药品计划价格显示方式.Left = fra查询天数.Left
        
        fra药品计划供应商设置.Top = fra药品计划价格显示方式.Top + fra药品计划价格显示方式.Height + 150
        fra药品计划供应商设置.Left = fra查询天数.Left
        
        chkALLPlanPoint.Top = fra药品计划供应商设置.Top + fra药品计划供应商设置.Height + 150
        chkALLPlanPoint.Left = fra药品计划供应商设置.Left
    End If

    fra盘点时间范围.Visible = False
    If strFunction = "药品盘点管理" Then
        cboUnit.Enabled = False
        fra盘点时间范围.Visible = True
    End If
    
    If strFunction = "药品自制入库管理" Then

    End If
    
    If strFunction = "药品调价管理" Then
        fra排序方式.Visible = False
        fra打印控制.Visible = False
        fra药品计划价格显示方式.Visible = False
        fra药品计划供应商设置.Visible = False
        chk留存领用.Visible = False
        fra盘点时间范围.Visible = False
        chkALLPlanPoint.Visible = False
        
        fra其他.Height = fra药品单位.Height
        
        tabMain.Height = fra其他.Top + fra其他.Height + 200
        tabMain.Width = fra其他.Left + fra其他.Width + 200
        
        Me.Height = tabMain.Top + tabMain.Height + cmdHelp.Height + 650
        Me.Width = tabMain.Left + tabMain.Width + 200
        
        cmdHelp.Top = tabMain.Top + tabMain.Height + 100
        CmdCancel.Top = cmdHelp.Top
        CmdCancel.Left = Me.Width - CmdCancel.Width - 200
        cmdOK.Top = cmdHelp.Top
        cmdOK.Left = CmdCancel.Left - cmdOK.Width - 50
    End If
    
    If strFunction = "药品质量管理" Then
        fra药品单位.Visible = False
        fra排序方式.Visible = False
        fra打印控制.Visible = False
        fra药品计划价格显示方式.Visible = False
        fra药品计划供应商设置.Visible = False
        chk留存领用.Visible = False
        fra盘点时间范围.Visible = False
        chkALLPlanPoint.Visible = False
        
        fra其他.Move fra药品单位.Left, fra药品单位.Top, fra药品单位.Width, fra药品单位.Height
        
        tabMain.Height = fra其他.Top + fra其他.Height + 200
        tabMain.Width = fra其他.Left + fra其他.Width + 200
        
        Me.Height = tabMain.Top + tabMain.Height + cmdHelp.Height + 650
        Me.Width = tabMain.Left + tabMain.Width + 200
        
        cmdHelp.Top = tabMain.Top + tabMain.Height + 100
        CmdCancel.Top = cmdHelp.Top
        CmdCancel.Left = Me.Width - CmdCancel.Width - 200
        cmdOK.Top = cmdHelp.Top
        cmdOK.Left = CmdCancel.Left - cmdOK.Width - 50

    End If
    
    
    chk留存领用.Visible = False
    If strFunction = "药品领用管理" Then
        chk留存领用.Visible = True
    End If
    
    If mlngModul = 1302 Or mlngModul = 1303 Or mlngModul = 1306 Then
        '1302 :其他入库;1303:库存差价; 1306：其他出库
    End If
    
    frm参数设置.Show vbModal, frmParent
End Sub
Private Sub cmd打印设置_Click()
    Dim strBill As String
    Select Case mstrFunction
    Case "药品外购入库管理"
        strBill = Split(cbo报表.Text, "(")(0)
    Case "药品其他入库管理"
        strBill = Split(cbo报表.Text, "(")(0)
    Case "药品自制入库管理"
        strBill = "ZL1_BILL_1301"
    Case "库存差价调整管理"
        strBill = "ZL1_BILL_1303"
    Case "药品移库管理"
        strBill = Split(cbo报表.Text, "(")(0)
    Case "药品领用管理"
        strBill = Split(cbo报表.Text, "(")(0)
    Case "药品其他出库管理"
        strBill = Split(cbo报表.Text, "(")(0)
    Case "药品盘点管理"
        strBill = "ZL1_BILL_1307"
    Case "药品计划管理"
        strBill = "zl1_bill_1330"
    Case "药品调价管理"
        strBill = "ZL1_BILL_1333"
    End Select
    Call ReportPrintSet(gcnOracle, glngSys, strBill, Me)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Me.cmd打印设置.Caption = "票据《" & Mid(mstrFunction, 1, Len(mstrFunction) - 2) & "单》打印设置"
    
    '更改设计时的部分状态
    fra查询天数.BackColor = &H8000000F
    fra盘点时间范围.BackColor = &H8000000F
    
    chkPrintCode.Visible = True
    cbo报表.Visible = True
    Select Case mlngModul
        Case 1300
            '下拉列表设置
            cbo报表.AddItem "ZL1_BILL_1300(单据打印)"
            cbo报表.AddItem "ZL1_INSIDE_1300_1(药品条码打印)"
            cbo报表.ListIndex = 0
        Case 1302
            '下拉列表设置
            cbo报表.AddItem "ZL1_BILL_1302(单据打印)"
            cbo报表.AddItem "ZL1_INSIDE_1302_1(药品条码打印)"
            cbo报表.ListIndex = 0
        Case 1304
            chkPrintCode.Caption = "存盘或审核(发送)后打印药品条码"
            '下拉列表设置
            cbo报表.AddItem "ZL1_BILL_1304(单据打印)"
            cbo报表.AddItem "ZL1_INSIDE_1304_1(药品条码打印)"
            cbo报表.ListIndex = 0
        Case 1305
            '下拉列表设置
            cbo报表.AddItem "ZL1_BILL_1305(单据打印)"
            cbo报表.AddItem "ZL1_INSIDE_1305_2(药品条码打印)"
            cbo报表.ListIndex = 0
        Case 1306
            '下拉列表设置
            cbo报表.AddItem "ZL1_BILL_1306(单据打印)"
            cbo报表.AddItem "ZL1_INSIDE_1306_1(药品条码打印)"
            cbo报表.ListIndex = 0
        Case Else
            chkPrintCode.Visible = False
            cbo报表.Visible = False
    End Select
End Sub

Private Sub txt查询天数_KeyPress(KeyAscii As Integer)
    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Then Exit Sub
    KeyAscii = 0
End Sub


Private Sub txt查询天数_Validate(Cancel As Boolean)
    If Val(txt查询天数.Text) > 7 Then
        If MsgBox("查询时间大于7天可能会导致查询很慢，是否继续？", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = False
            txt查询天数.SetFocus
            zlControl.TxtSelAll txt查询天数
        End If
    End If
    If Val(txt查询天数.Text) = 0 Then
        MsgBox "查询时间必须大于0，请重新输入！", vbInformation, gstrSysName
        Cancel = False
        txt查询天数.SetFocus
        zlControl.TxtSelAll txt查询天数
    End If
End Sub


Private Sub txt盘点时间_Change()
    UpD盘点时间.Value = Val(txt盘点时间.Text)
End Sub

Private Sub txt盘点时间_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyBack Then
        If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
            KeyAscii = 0
        End If
        If Val(txt盘点时间.Text & Chr(KeyAscii)) > 90 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txt盘点时间_Validate(Cancel As Boolean)
    If Val(txt盘点时间.Text) > 90 Then
        MsgBox "盘点时间范围不能大于3个月！", vbInformation, gstrSysName
        Cancel = True
    End If
End Sub

Private Sub UpD盘点时间_Change()
    txt盘点时间.Text = UpD盘点时间.Value
End Sub


