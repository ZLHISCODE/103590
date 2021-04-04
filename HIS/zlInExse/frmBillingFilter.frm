VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBillingFilter 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "过滤设置"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7110
   ControlBox      =   0   'False
   Icon            =   "frmBillingFilter.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5880
      TabIndex        =   20
      Top             =   600
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5880
      TabIndex        =   21
      Top             =   1020
      Width           =   1100
   End
   Begin VB.CommandButton cmdDef 
      Caption         =   "缺省(&D)"
      Height          =   350
      Left            =   5880
      TabIndex        =   22
      Top             =   1800
      Width           =   1100
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   5741
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "基本(&0)"
      TabPicture(0)   =   "frmBillingFilter.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblDate"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label7"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lbl操作员"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "dtpBegin"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "dtpEnd"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "chkType(1)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "chkType(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "chkBill(0)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "chkBill(2)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "chkBill(1)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txt住院号"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "chkBill(3)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cbo操作员"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtNOBegin"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtNoEnd"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txt姓名"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).ControlCount=   20
      TabCaption(1)   =   "收费项目(&1)"
      TabPicture(1)   =   "frmBillingFilter.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lbl收入项目(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "tlbOpt(0)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "txtInput(0)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "ListFeeItem(0)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "收入项目(&2)"
      TabPicture(2)   =   "frmBillingFilter.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lbl收入项目(1)"
      Tab(2).Control(1)=   "tlbOpt(1)"
      Tab(2).Control(2)=   "txtInput(1)"
      Tab(2).Control(3)=   "ListFeeItem(1)"
      Tab(2).ControlCount=   4
      Begin VB.ListBox ListFeeItem 
         Height          =   1950
         Index           =   1
         Left            =   -73680
         Style           =   1  'Checkbox
         TabIndex        =   19
         ToolTipText     =   "Ctrl+A全选,Ctrl+C全消,如果一个都未选则表示不限制"
         Top             =   960
         Width           =   4095
      End
      Begin VB.TextBox txtInput 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   1
         Left            =   -73680
         MaxLength       =   40
         TabIndex        =   18
         ToolTipText     =   "最多匹配100项搜索结果"
         Top             =   480
         Width           =   2160
      End
      Begin VB.ListBox ListFeeItem 
         Height          =   1950
         Index           =   0
         Left            =   -73680
         Style           =   1  'Checkbox
         TabIndex        =   16
         ToolTipText     =   "Ctrl+A全选,Ctrl+C全消,如果一个都未选则表示不限制"
         Top             =   960
         Width           =   4095
      End
      Begin VB.TextBox txtInput 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   0
         Left            =   -73680
         MaxLength       =   40
         TabIndex        =   15
         ToolTipText     =   "最多匹配100项搜索结果"
         Top             =   480
         Width           =   2160
      End
      Begin VB.TextBox txt姓名 
         Height          =   300
         IMEMode         =   1  'ON
         Left            =   3420
         MaxLength       =   100
         TabIndex        =   13
         Top             =   2490
         Width           =   1470
      End
      Begin VB.TextBox txtNoEnd 
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3420
         MaxLength       =   8
         TabIndex        =   11
         Top             =   2070
         Width           =   1470
      End
      Begin VB.TextBox txtNOBegin 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1035
         MaxLength       =   8
         TabIndex        =   10
         Top             =   2070
         Width           =   1470
      End
      Begin VB.ComboBox cbo操作员 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1035
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1575
         Width           =   1470
      End
      Begin VB.CheckBox chkBill 
         Caption         =   "临嘱记帐"
         Height          =   210
         Index           =   3
         Left            =   3840
         TabIndex        =   9
         Top             =   1770
         Value           =   1  'Checked
         Width           =   1020
      End
      Begin VB.TextBox txt住院号 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1035
         MaxLength       =   18
         TabIndex        =   12
         Top             =   2490
         Width           =   1470
      End
      Begin VB.CheckBox chkBill 
         Caption         =   "自动记帐"
         Height          =   210
         Index           =   1
         Left            =   3840
         TabIndex        =   7
         Top             =   1230
         Width           =   1020
      End
      Begin VB.CheckBox chkBill 
         Caption         =   "长嘱记帐"
         Height          =   195
         Index           =   2
         Left            =   3840
         TabIndex        =   8
         Top             =   1515
         Value           =   1  'Checked
         Width           =   1020
      End
      Begin VB.CheckBox chkBill 
         Caption         =   "普通记帐"
         Height          =   195
         Index           =   0
         Left            =   3840
         TabIndex        =   6
         Top             =   975
         Value           =   1  'Checked
         Width           =   1020
      End
      Begin VB.CheckBox chkType 
         Caption         =   "记帐单据"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   0
         Left            =   3240
         TabIndex        =   4
         Top             =   630
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkType 
         Caption         =   "销帐单据"
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   1
         Left            =   4380
         TabIndex        =   5
         Top             =   630
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   1035
         TabIndex        =   2
         Top             =   1020
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   113704963
         CurrentDate     =   36588
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Left            =   1035
         TabIndex        =   1
         Top             =   600
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   113704963
         CurrentDate     =   36588
      End
      Begin MSComctlLib.Toolbar tlbOpt 
         Height          =   1800
         Index           =   0
         Left            =   -74760
         TabIndex        =   30
         Top             =   1080
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   3175
         ButtonWidth     =   1614
         ButtonHeight    =   1058
         Style           =   1
         ImageList       =   "ils16"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "移除(&M)"
               Key             =   "Delete"
               Object.ToolTipText     =   "移除当前选择的列表项"
               ImageKey        =   "Delete"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "清空(&T)"
               Key             =   "Clear"
               Object.ToolTipText     =   "清空列表项目"
               ImageKey        =   "Clear"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "保存(&S)"
               Key             =   "Save"
               Object.ToolTipText     =   "保存选择的列表项目"
               ImageKey        =   "Save"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbOpt 
         Height          =   525
         Index           =   1
         Left            =   -74760
         TabIndex        =   31
         Top             =   1080
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   926
         ButtonWidth     =   1455
         ButtonHeight    =   926
         Style           =   1
         ImageList       =   "ils16"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "移除(&M)"
               Key             =   "Delete"
               Object.ToolTipText     =   "移除当前选择的列表项"
               ImageKey        =   "Delete"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "清空(&T)"
               Key             =   "Clear"
               Object.ToolTipText     =   "清空列表项目"
               ImageKey        =   "Clear"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "保存(&S)"
               Key             =   "Save"
               Object.ToolTipText     =   "保存选择的列表项目"
               ImageKey        =   "Save"
            EndProperty
         EndProperty
      End
      Begin VB.Label lbl收入项目 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "收入项目(&R)"
         Height          =   180
         Index           =   1
         Left            =   -74760
         TabIndex        =   17
         Top             =   540
         Width           =   990
      End
      Begin VB.Label lbl收入项目 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "收费项目(&F)"
         Height          =   180
         Index           =   0
         Left            =   -74760
         TabIndex        =   14
         Top             =   540
         Width           =   990
      End
      Begin VB.Label lbl操作员 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "操作员"
         Height          =   180
         Left            =   420
         TabIndex        =   29
         Top             =   1650
         Width           =   540
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名"
         Height          =   180
         Left            =   3015
         TabIndex        =   28
         Top             =   2550
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单据号"
         Height          =   180
         Left            =   420
         TabIndex        =   27
         Top             =   2130
         Width           =   540
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "至"
         Height          =   180
         Left            =   3180
         TabIndex        =   26
         Top             =   2130
         Width           =   180
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结束时间"
         Height          =   180
         Left            =   240
         TabIndex        =   25
         Top             =   1080
         Width           =   720
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "开始时间"
         Height          =   180
         Left            =   240
         TabIndex        =   24
         Top             =   660
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院号"
         Height          =   180
         Left            =   420
         TabIndex        =   23
         Top             =   2550
         Width           =   540
      End
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   6240
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillingFilter.frx":0060
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillingFilter.frx":03FA
            Key             =   "Insert"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillingFilter.frx":0794
            Key             =   "Clear"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillingFilter.frx":0B2E
            Key             =   "Delete"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmBillingFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Public mstrFilter As String
Public mlngDeptID As Long   '当前主界面科室
Public mlngUnitID As Long   '科室所属病区
Public mstrPrivs As String
Public mblnDateMoved As Boolean '当前所选条件的数据是否在后备数据表中

'传出参数
Public mstrFeeItems As String '收费项目ID串
Public mstrIncomeItems As String '收入项目ID串

Private mintTab As Integer

Private Enum chkTypes
    记帐单据 = 0
    销帐单据 = 1
End Enum
Public Enum chkBills
    普通记帐 = 0
    自动记帐 = 1
    长嘱记帐 = 2
    临嘱记帐 = 3
End Enum

Private Sub cbo操作员_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii >= 32 Then
        lngIdx = zlControl.CboMatchIndex(cbo操作员.hWnd, KeyAscii)
        If lngIdx = -1 And cbo操作员.ListCount > 0 Then lngIdx = 0
        cbo操作员.ListIndex = lngIdx
    End If
End Sub


Private Sub chkBill_Click(Index As Integer)
    Dim i As Integer, j As Integer
    
    j = 0
    For i = 0 To chkBill.UBound
        If chkBill(i).Value = 0 Then j = j + 1
    Next
    If j = i Then
        If Index = chkBills.自动记帐 And Not (frmManageBilling.tbs.SelectedItem.Key = "Auditing") Then
            '划价禁用自动记帐
            chkBill(chkBills.普通记帐).Value = 1
        Else
            chkBill(Index).Value = 1  '最后的i是加了1的
        End If
    End If
    
End Sub

Private Sub chkType_Click(Index As Integer)
    If chkType(0).Value = 0 And chkType(1).Value = 0 Then
        chkType((Index + 1) Mod 2).Value = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    gblnOK = False
    Hide
End Sub

Private Sub cmdDef_Click()
    Form_Load
End Sub



Private Sub cmdOK_Click()
    If txtNOBegin.Text <> "" And txtNoEnd.Text <> "" Then
        If txtNoEnd.Text < txtNOBegin.Text Then
            MsgBox "结束单据号不能小于开始单据号！", vbInformation, gstrSysName
            txtNoEnd.SetFocus: Exit Sub
        End If
    End If
    If cbo操作员.ListIndex = -1 Then
        MsgBox "不能确定要过滤的操作员！", vbInformation, gstrSysName
        cbo操作员.SetFocus: Exit Sub
    End If
    
    Call MakeFilter
    
    gblnOK = True
    Hide
End Sub

Private Sub dtpEnd_Change()
    dtpBegin.MaxDate = dtpEnd.Value
End Sub

Private Sub Form_Activate()
    Dim i As Integer
    Dim bln划价 As Boolean
    
    dtpBegin.SetFocus
    
    '划价记帐单时,禁自动记帐,禁销帐单据
    bln划价 = Not (frmManageBilling.tbs.SelectedItem.Key = "Auditing")
    If bln划价 Then
        chkType(chkTypes.记帐单据).Value = 1
        chkType(chkTypes.销帐单据).Value = 0
        chkBill(chkBills.自动记帐).Value = 0    '值不同时,会调用click事件
    End If
    chkType(chkTypes.记帐单据).Enabled = Not bln划价
    chkType(chkTypes.销帐单据).Enabled = Not bln划价
    chkBill(chkBills.自动记帐).Enabled = Not bln划价
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If mintTab = 1 Or mintTab = 2 Then txtInput(mintTab - 1).SetFocus
    ElseIf KeyCode = vbKeyReturn And Not (mintTab = 1 Or mintTab = 2) Then
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf Shift = 2 Then
        If mintTab = 1 Or mintTab = 2 Then
            Dim i As Integer, Index As Integer
            
            Index = mintTab - 1
            If UCase(Chr(KeyCode)) = "A" Then
                For i = 0 To ListFeeItem(Index).ListCount - 1
                    ListFeeItem(Index).Selected(i) = True
                Next
            ElseIf UCase(Chr(KeyCode)) = "C" Then
                For i = 0 To ListFeeItem(Index).ListCount - 1
                    ListFeeItem(Index).Selected(i) = False
                Next
            End If
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Public Sub LoadOper()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    
    Err = 0: On Error GoTo errH:
    cbo操作员.Clear
    
    If InStr(mstrPrivs, ";所有操作员;") > 0 Then
        cbo操作员.AddItem "所有操作员"
        cbo操作员.ListIndex = 0
    End If
    
    strSQL = "Select Distinct A.ID,A.编号,A.姓名,A.简码" & _
        " From 人员表 A,部门人员 C Where A.ID=C.人员ID And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & vbNewLine & _
        IIf(InStr(mstrPrivs, ";所有操作员;") = 0 And mlngDeptID > 0, " And C.部门ID IN([1],[2])", "") & _
        IIf(InStr(mstrPrivs, ";所有操作员;") = 0, " And A.ID=[3]", "") & _
        " Order by A.简码"
            
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngDeptID, mlngUnitID, UserInfo.ID)
    For i = 1 To rsTmp.RecordCount
        cbo操作员.AddItem rsTmp!简码 & "-" & rsTmp!姓名
        cbo操作员.ItemData(cbo操作员.NewIndex) = rsTmp!ID
        If rsTmp!ID = UserInfo.ID Then cbo操作员.ListIndex = cbo操作员.NewIndex
        rsTmp.MoveNext
    Next
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Dim Curdate As Date, i As Long, Index As Integer
    Dim strListFeeItem As String
    Dim arrItem As Variant
    On Error GoTo errH
    
    gblnOK = False
    
    txtNOBegin.Text = ""
    txtNoEnd.Text = ""
    txt住院号.Text = ""
    txt姓名.Text = ""
    
    mstrFeeItems = ""
    mstrIncomeItems = ""
     
    '设置初始值
    
    Curdate = zlDatabase.Currentdate
    dtpBegin.MaxDate = Format(Curdate, "yyyy-MM-dd 23:59:59")
    dtpBegin.Value = Format(Curdate, "yyyy-MM-dd 00:00:00")
    dtpEnd.Value = dtpBegin.MaxDate
    
    
    Call LoadOper
    
    Call SSTab1_Click(0)
    
    If InStr(1, mstrPrivs, ";明细项目过滤;") = 0 Then
        SSTab1.TabVisible(1) = False
        SSTab1.TabVisible(2) = False
    Else
        For Index = 0 To 1
            strListFeeItem = ""
            ListFeeItem(Index).Clear
            
            Call GetRegisterItem(g私有模块, Me.Name & "\" & ListFeeItem(0).Name, IIf(Index = 0, "收费项目列表", "收入项目列表"), strListFeeItem)
            If strListFeeItem <> "" Then
                arrItem = Split(strListFeeItem, ";")
                
                For i = 0 To UBound(arrItem)
                    ListFeeItem(Index).AddItem Split(arrItem(i), ",")(0)
                    ListFeeItem(Index).ItemData(ListFeeItem(Index).NewIndex) = Val(Split(arrItem(i), ",")(1))
                    ListFeeItem(Index).Selected(ListFeeItem(Index).NewIndex) = IIf(Val(Split(arrItem(i), ",")(2)) = 1, True, False)
                Next
            End If
        Next
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ListFeeItem_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If ListFeeItem(Index).ListIndex >= 0 Then
            ListFeeItem(Index).RemoveItem ListFeeItem(Index).ListIndex
        End If
    End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    
    Select Case SSTab1.Caption
        Case "基本(&0)"
           mintTab = 0
        Case "收费项目(&1)"
            mintTab = 1
        Case "收入项目(&2)"
            mintTab = 2
    End Select
    
End Sub


Private Sub tlbOpt_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Delete"
            If ListFeeItem(Index).ListIndex >= 0 Then
                Call ListFeeItem(Index).RemoveItem(ListFeeItem(Index).ListIndex)
            End If
        Case "Clear"
            ListFeeItem(Index).Clear
        Case "Save"
            Dim strTmp As String, i As Long
            With ListFeeItem(Index)
                For i = 0 To .ListCount - 1
                    strTmp = strTmp & ";" & .List(i) & "," & .ItemData(i) & "," & IIf(.Selected(i), 1, 0)
                Next
            End With
            strTmp = Mid(strTmp, 2)
            Call SaveRegisterItem(g私有模块, Me.Name & "\" & ListFeeItem(0).Name, IIf(Index = 0, "收费项目列表", "收入项目列表"), strTmp)
    End Select
End Sub

Private Sub txtInput_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txtInput(Index))
End Sub

Private Sub txtInput_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim strSQL As String, strInput As String, strMatch As String, strIF As String
    Dim rsTmp As ADODB.Recordset
    Dim blnCancel As Boolean, i As Long
    Dim vRect As RECT
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        strInput = UCase(Trim(txtInput(Index).Text))
        If strInput = "" Then Exit Sub
        strMatch = IIf(Len(strInput) < 3, "", gstrLike)
        
        If Index = 0 Then
        '收费项目
            If zlCommFun.IsNumOrChar(strInput) Then
                strIF = " And (A.编码 like [1] Or B.简码 like [1] And B.码类 in(3," & gbytCode + 1 & "))"
            Else
                strIF = " And B.名称 like [1]"
            End If
            strSQL = "Select Distinct A.ID, A.编码, B.名称 ,A.规格, A.产地, A.计算单位 " & _
                  " From 收费项目目录 A,收费项目别名 B Where A.id=B.收费细目ID " & strIF & _
                  " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & vbNewLine & _
                  " And rownum<101 Order by 名称"
        Else
        '收入项目
            If zlCommFun.IsNumOrChar(strInput) Then
                If IsNumeric(strInput) Then
                    strIF = " And 编码 like [1]"
                Else
                    strIF = " And 简码 like [1]"
                End If
            Else
                strIF = " And 名称 like [1]"
            End If
            
            strSQL = "Select ID, 编码, 名称 From 收入项目 Where 末级=1 " & strIF & _
                " And rownum<101 Order by 名称"
        End If
        
        On Error GoTo errH
        vRect = zlControl.GetControlRect(txtInput(Index).hWnd)
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "项目选择", 1, "", "请选择", False, False, True, vRect.Left, vRect.Top, txtInput(Index).Height, blnCancel, False, True, strMatch & strInput & "%")
        If Not rsTmp Is Nothing Then
            With ListFeeItem(Index)
                For i = 0 To .ListCount - 1
                    If .ItemData(i) = rsTmp!ID Then
                        txtInput(Index).SetFocus
                        txtInput(Index).SelStart = 0
                        txtInput(Index).SelLength = Len(txtInput(Index).Text)
                        Exit Sub
                    End If
                Next
                If .ListCount < 100 Then
                    If Index = 0 Then
                        .AddItem rsTmp!编码 & "-" & rsTmp!名称 & "(" & rsTmp!规格 & ")"
                    Else
                        .AddItem rsTmp!编码 & "-" & rsTmp!名称
                    End If
                    .ItemData(.NewIndex) = rsTmp!ID
                    .Selected(.NewIndex) = True
                Else
                    MsgBox "出于性能考虑,搜索项目最多只允许添加100项!", vbInformation, gstrSysName
                End If
            End With
        End If
        
        txtInput(Index).SetFocus
        txtInput(Index).SelStart = 0
        txtInput(Index).SelLength = Len(txtInput(Index).Text)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtNOBegin_Change()
    txtNoEnd.Enabled = Not (Trim(txtNOBegin.Text) = "")
    If Trim(txtNOBegin.Text = "") Then txtNoEnd.Text = ""
End Sub

Private Sub txtNOBegin_GotFocus()
    zlControl.TxtSelAll txtNOBegin
End Sub

Private Sub txtNOBegin_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    '46516
    zlControl.TxtCheckKeyPress txtNOBegin, KeyAscii, m文本式
 
End Sub

Private Sub txtNOBegin_LostFocus()
    If txtNOBegin.Text <> "" Then txtNOBegin.Text = GetFullNO(txtNOBegin.Text, 14)
End Sub

Private Sub txtNOEnd_LostFocus()
    If txtNoEnd.Text <> "" Then txtNoEnd.Text = GetFullNO(txtNoEnd.Text, 14)
End Sub

Private Sub txtNoEnd_GotFocus()
    zlControl.TxtSelAll txtNoEnd
End Sub

Private Sub txtNoEnd_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    '46516
    zlControl.TxtCheckKeyPress txtNoEnd, KeyAscii, m文本式
End Sub

Private Sub MakeFilter()
    Dim bln普通记帐 As Boolean
    Dim i As Long, Index As Integer
    Dim strIDs As String
    
    mstrFilter = " And 登记时间 Between [1] And [2]"
    
    If frmManageBilling.tbs.SelectedItem.Key = "Auditing" Then
        mblnDateMoved = zlDatabase.DateMoved(Format(IIf(dtpBegin.Value < dtpEnd.Value, dtpBegin.Value, dtpEnd.Value), dtpBegin.CustomFormat), , , Me.Caption)
    Else
        '划价单筛选时,不用从后备数据表取
        mblnDateMoved = False
    End If
    
    If txtNOBegin.Text <> "" And txtNoEnd.Text <> "" Then
        mstrFilter = mstrFilter & " And NO Between [3] And [4]"
    ElseIf txtNOBegin.Text <> "" Then
        mstrFilter = mstrFilter & " And NO=[3]"
    End If
    
    If IsNumeric(txt住院号.Text) Then
        mstrFilter = mstrFilter & " And 病人ID = (Select Distinct 病人ID From 病案主页 Where 住院号 = [5])"
    End If
    
    If txt姓名.Text <> "" Then
        If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Left(txt姓名.Text, 1))) > 0 Then
            mstrFilter = mstrFilter & " And Upper(姓名) Like [6]"
        Else
            mstrFilter = mstrFilter & " And 姓名 Like [6]"
        End If
    End If
    
    
    If frmManageBilling.tbs.SelectedItem.Key = "Auditing" Then
        '自动记帐
        bln普通记帐 = chkBill(chkBills.普通记帐).Value = 1 Or chkBill(chkBills.临嘱记帐).Value = 1 Or chkBill(chkBills.长嘱记帐).Value = 1
        If chkBill(chkBills.自动记帐).Value = 1 And bln普通记帐 Then
            mstrFilter = mstrFilter & " And 记录性质 IN(2,3)"
        ElseIf chkBill(chkBills.自动记帐).Value = 0 And bln普通记帐 Then
            mstrFilter = mstrFilter & " And 记录性质=2"
        ElseIf chkBill(chkBills.自动记帐).Value = 1 And Not bln普通记帐 Then
            mstrFilter = mstrFilter & " And 记录性质=3"
        End If
        
        '记帐或销帐
        If chkType(chkTypes.记帐单据).Value = 1 And chkType(chkTypes.销帐单据).Value = 1 Then
            mstrFilter = mstrFilter & " And 记录状态 IN(1,2,3)"
        ElseIf chkType(chkTypes.记帐单据).Value = 1 Then
            mstrFilter = mstrFilter & " And 记录状态 IN(1,3)"
        ElseIf chkType(chkTypes.销帐单据).Value = 1 Then
            mstrFilter = mstrFilter & " And 记录状态=2"
        End If
    Else
        mstrFilter = mstrFilter & " And 记录性质=2 And 记录状态=0"
    End If
    
    If InStr(1, mstrPrivs, ";明细项目过滤;") > 0 Then
        For Index = 0 To 1
            strIDs = ""
            For i = 0 To ListFeeItem(Index).ListCount - 1
                If ListFeeItem(Index).Selected(i) Then
                    strIDs = strIDs & "," & ListFeeItem(Index).ItemData(i)
                End If
            Next
            If strIDs <> "" Then
                strIDs = Mid(strIDs, 2)
                If Index = 0 Then
                    mstrFeeItems = strIDs
                    mstrFilter = mstrFilter & " And Instr(','||[9]||',',','||收费细目ID||',')>0"
                Else
                    mstrIncomeItems = strIDs
                    mstrFilter = mstrFilter & " And Instr(','||[10]||',',','||收入项目ID||',')>0"
                End If
            End If
        Next
    End If
    
    
    '医嘱的判断在主界面做
    '因为可能已审及未审在主界面变化,所以操作员在主界面取
End Sub



Private Sub txt住院号_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt姓名_GotFocus()
    zlControl.TxtSelAll txt姓名
End Sub

Private Sub txt住院号_GotFocus()
    zlControl.TxtSelAll txt住院号
End Sub
