VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRegistPlanFilter 
   BorderStyle     =   0  'None
   Caption         =   "过滤设置"
   ClientHeight    =   5820
   ClientLeft      =   0
   ClientTop       =   -45
   ClientWidth     =   4365
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fra 
      Height          =   5040
      Index           =   0
      Left            =   120
      TabIndex        =   23
      Top             =   45
      Width           =   4170
      Begin VB.CheckBox chkShowExpiredPlan 
         Caption         =   "显示失效的计划"
         Height          =   270
         Left            =   285
         TabIndex        =   25
         Top             =   1830
         Width           =   2235
      End
      Begin VB.CommandButton cmdDoct 
         Height          =   240
         Left            =   2970
         Picture         =   "frmRegistPlanFilter.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "选择项目(F4)"
         Top             =   720
         Width           =   255
      End
      Begin VB.CheckBox chk有效期 
         Caption         =   "有效期"
         Height          =   195
         Left            =   285
         TabIndex        =   4
         Top             =   1148
         Width           =   840
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Left            =   1185
         TabIndex        =   5
         Top             =   1095
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   115736579
         CurrentDate     =   38091
      End
      Begin VB.TextBox txtDoct 
         Height          =   300
         Left            =   1185
         MaxLength       =   8
         TabIndex        =   3
         Top             =   690
         Width           =   2070
      End
      Begin VB.ComboBox cbo科室 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1185
         TabIndex        =   1
         Text            =   "cbo科室"
         Top             =   285
         Width           =   2070
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   1185
         TabIndex        =   6
         Top             =   1440
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   115736579
         CurrentDate     =   38091
      End
      Begin VB.PictureBox picCon 
         BorderStyle     =   0  'None
         Height          =   2865
         Left            =   210
         ScaleHeight     =   2865
         ScaleWidth      =   3885
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   2070
         Width           =   3885
         Begin VB.CheckBox chk计划 
            Caption         =   "仅显示未审计划"
            Height          =   270
            Index           =   1
            Left            =   405
            TabIndex        =   20
            Top             =   2655
            Width           =   2235
         End
         Begin VB.CheckBox chkDate 
            Caption         =   "按生效日期查找"
            Height          =   375
            Index           =   0
            Left            =   75
            TabIndex        =   7
            Top             =   0
            Value           =   1  'Checked
            Width           =   1665
         End
         Begin VB.CheckBox chkDate 
            Caption         =   "按安排时间查找"
            Height          =   375
            Index           =   1
            Left            =   75
            TabIndex        =   11
            Top             =   825
            Width           =   1665
         End
         Begin VB.CheckBox chkDate 
            Caption         =   "按审核时间查找"
            Height          =   375
            Index           =   2
            Left            =   75
            TabIndex        =   15
            Top             =   1575
            Width           =   1665
         End
         Begin MSComCtl2.DTPicker dtpStartDate 
            Height          =   315
            Index           =   0
            Left            =   615
            TabIndex        =   8
            Top             =   375
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   115736579
            CurrentDate     =   37007
         End
         Begin MSComCtl2.DTPicker dtpEndDate 
            Height          =   315
            Index           =   0
            Left            =   2430
            TabIndex        =   10
            Top             =   375
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   115736579
            CurrentDate     =   37007
         End
         Begin MSComCtl2.DTPicker dtpStartDate 
            Height          =   315
            Index           =   1
            Left            =   615
            TabIndex        =   12
            Top             =   1185
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   115736579
            CurrentDate     =   37007
         End
         Begin MSComCtl2.DTPicker dtpEndDate 
            Height          =   315
            Index           =   1
            Left            =   2430
            TabIndex        =   14
            Top             =   1185
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   115736579
            CurrentDate     =   37007
         End
         Begin MSComCtl2.DTPicker dtpStartDate 
            Height          =   315
            Index           =   2
            Left            =   615
            TabIndex        =   16
            Top             =   1935
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   115736579
            CurrentDate     =   37007
         End
         Begin MSComCtl2.DTPicker dtpEndDate 
            Height          =   315
            Index           =   2
            Left            =   2430
            TabIndex        =   18
            Top             =   1935
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   115736579
            CurrentDate     =   37007
         End
         Begin VB.CheckBox chk计划 
            Caption         =   "仅显示未生效的计划"
            Height          =   270
            Index           =   0
            Left            =   420
            TabIndex        =   19
            Top             =   2415
            Width           =   2235
         End
         Begin VB.Label lbl至 
            AutoSize        =   -1  'True
            Caption         =   "至"
            Height          =   180
            Index           =   5
            Left            =   2070
            TabIndex        =   17
            Top             =   1995
            Width           =   180
         End
         Begin VB.Label lbl至 
            AutoSize        =   -1  'True
            Caption         =   "至"
            Height          =   180
            Index           =   4
            Left            =   2070
            TabIndex        =   13
            Top             =   1245
            Width           =   180
         End
         Begin VB.Label lbl至 
            AutoSize        =   -1  'True
            Caption         =   "至"
            Height          =   180
            Index           =   3
            Left            =   2070
            TabIndex        =   9
            Top             =   435
            Width           =   180
         End
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "挂号科室"
         Height          =   180
         Index           =   0
         Left            =   390
         TabIndex        =   0
         Top             =   345
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "挂号医生"
         Height          =   180
         Index           =   1
         Left            =   390
         TabIndex        =   2
         Top             =   735
         Width           =   720
      End
   End
   Begin VB.CommandButton cmd刷新 
      Caption         =   "过滤(&O)"
      Height          =   350
      Left            =   3135
      TabIndex        =   21
      Top             =   5175
      Width           =   1100
   End
End
Attribute VB_Name = "frmRegistPlanFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnPlanPage As Boolean '计划安排页面
Private mArrFilter As Variant, mrs科室 As ADODB.Recordset
Private mlngModule As Long, mstrPrivs As String
Private mblnShowStoped As Boolean '显示停用安排
Private mblnShowDel As Boolean '实现删除安排
'--------------------------------------------------------------------------------------------------------
Public Event zlRefreshCon(ByVal ArrFilter As Variant)


Private Sub cbo科室_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If cbo科室.Text = "" Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If cbo科室.ListIndex >= 0 Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If Select科室(Me, mlngModule, mrs科室, cbo科室, cbo科室.Text) = True Then
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    End If
    If cbo科室.Enabled Then cbo科室.SetFocus
    zlcontrol.TxtSelAll cbo科室
    
End Sub

Private Sub cmd刷新_Click()
    Call GetBuildingtCondition
    RaiseEvent zlRefreshCon(mArrFilter)
End Sub

Private Sub InitCon()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化缺省条件
    '编制:刘兴洪
    '日期:2009-09-16 15:00:39
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim i As Long, bln所属部门 As Boolean
    dtpBegin.Value = Format(zlDatabase.Currentdate, "yyyy-MM-dd 00:00:00")
    dtpEnd.Value = CDate("3000-01-01")
 
    '检查门诊临床科室
    If zlStr.IsHavePrivs(mstrPrivs, "所有科室") Then
        cbo科室.AddItem "所有科室"
        cbo科室.ListIndex = 0
    Else
        cbo科室.AddItem "操作员所属科室"
        cbo科室.ListIndex = 0
        cbo科室.ItemData(cbo科室.NewIndex) = -1
        bln所属部门 = True
    End If
    Set mrs科室 = GetDepartments("'临床'", "1,3", bln所属部门)
    
    If Not mrs科室 Is Nothing Then
        For i = 1 To mrs科室.RecordCount
            cbo科室.AddItem "[" & mrs科室!编码 & "]" & mrs科室!名称
            cbo科室.ItemData(cbo科室.NewIndex) = Val(Nvl(mrs科室!ID))
            mrs科室.MoveNext
        Next
    End If

    chkShowExpiredPlan.Value = IIf(zlDatabase.GetPara("显示失效计划", glngSys, mlngModule, "0", Array(chkShowExpiredPlan)) = "1", 1, 0)
    '初始化条件
   ' dtpEndDate(0).MaxDate =
    dtpEndDate(1).MaxDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
    dtpEndDate(2).MaxDate = dtpEndDate(1).MaxDate
    
    dtpEndDate(0).Value = Format(DateAdd("m", 1, dtpEndDate(1).MaxDate), "yyyy-mm-dd")
    dtpEndDate(1).Value = dtpEndDate(1).MaxDate
    dtpEndDate(2).Value = dtpEndDate(1).MaxDate
    
    dtpStartDate(0).Value = dtpEndDate(1).MaxDate
    dtpStartDate(1).Value = Format(DateAdd("m", -1, dtpEndDate(1).MaxDate), "yyyy-mm-dd")
    dtpStartDate(2).Value = dtpStartDate(1).Value
End Sub
Private Sub GetBuildingtCondition()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取条件
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-09-15 09:52:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllFilter As New Collection
    If cbo科室.ListIndex < 0 Then cbo科室.ListIndex = 0
    If cbo科室.ItemData(cbo科室.ListIndex) <> 0 Then
        cllFilter.Add cbo科室.ItemData(cbo科室.ListIndex), "科室ID"
    Else
        cllFilter.Add 0, "科室ID"
    End If
    If txtDoct.Text <> "" Then
        If Val(txtDoct.Tag) <> 0 Then
            cllFilter.Add Array(Val(txtDoct.Tag), "ID"), "医生ID"
        ElseIf InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Left(txtDoct.Text, 1))) > 0 Then
            cllFilter.Add Array(Val(txtDoct.Tag), "UPR"), "医生ID"
        Else
            cllFilter.Add Array(Val(txtDoct.Tag), "NONE"), "医生ID"
        End If
    Else
        cllFilter.Add Array(Val(txtDoct.Tag), "NOT"), "医生ID"
    End If
    If chk有效期.Value = 1 Then
        cllFilter.Add Array(Format(dtpBegin.Value, "yyyy-mm-dd HH:MM:SS"), Format(dtpEnd.Value, "yyyy-mm-dd HH:MM:SS")), "有效期"
    Else
        cllFilter.Add Array("1901-01-01", "1901-01-01"), "有效期"
    End If
     
     If chkDate(0).Value = 1 Then
        cllFilter.Add Array(Format(dtpStartDate(0).Value, "yyyy-mm-dd") & " 00:00:00", Format(dtpEndDate(0).Value, "yyyy-mm-dd") & " 23:59:59"), "生效时间"
    Else
        cllFilter.Add Array("1901-01-01", "1901-01-01"), "生效时间"
    End If
    
    If chkDate(1).Value = 1 Then
        cllFilter.Add Array(Format(dtpStartDate(1).Value, "yyyy-mm-dd") & " 00:00:00", Format(dtpEndDate(1).Value, "yyyy-mm-dd") & " 23:59:59"), "安排时间"
    Else
        cllFilter.Add Array("1901-01-01", "1901-01-01"), "安排时间"
    End If
    
    If chkDate(2).Value = 1 Then
        cllFilter.Add Array(Format(dtpStartDate(2).Value, "yyyy-mm-dd") & " 00:00:00", Format(dtpEndDate(2).Value, "yyyy-mm-dd") & " 23:59:59"), "审核时间"
    Else
        cllFilter.Add Array("1901-01-01", "1901-01-01"), "审核时间"
    End If
    cllFilter.Add IIf(chk计划(0).Value = 1, 1, 0), "仅显未生效计划"
    cllFilter.Add IIf(chk计划(1).Value = 1, 1, 0), "仅显示未审计划"
    '38505
    cllFilter.Add IIf(mblnShowStoped, 1, 0), "显示停用安排"
    
    cllFilter.Add IIf(mblnShowDel, 1, 0), "显示删除安排"
    Set mArrFilter = cllFilter
 End Sub

Public Property Get GetCondition() As Variant
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取条件
    '返回:相关条件
    '编制:刘兴洪
    '日期:2009-09-15 09:54:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '获取条件
    Call GetBuildingtCondition
    Set GetCondition = mArrFilter
End Property

Private Sub cbo科室_Click()
    txtDoct.Text = ""
    txtDoct.Tag = ""
End Sub

Private Sub chk有效期_Click()
    dtpBegin.Enabled = chk有效期.Value = 1
    dtpEnd.Enabled = chk有效期.Value = 1
    
    If Visible And dtpBegin.Enabled Then
        dtpBegin.SetFocus
    End If
End Sub

'Private Sub cmdCancel_Click()
'    mstrFilter = ""
'    mblnOK = False
'    Me.Hide
'End Sub

Private Sub cmdDoct_Click()
    Dim rsTemp As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    Dim strSQL As String, lngDeptID As Long
    
    
    If cbo科室.ItemData(cbo科室.ListIndex) <= 0 Then
        lngDeptID = 0
        '所有门诊临床科室的医生
        strSQL = "" & _
        " Select Distinct A.ID From 部门表 A,部门性质说明 B" & _
        " Where A.ID=B.部门ID And B.工作性质='临床' And B.服务对象 IN(1,3) " & _
                    IIf(zlStr.IsHavePrivs(mstrPrivs, "所有科室") = False, " And Exists(Select 1 From 部门人员 where 人员id=[1] and C.部门id=部门id) ", "")
        strSQL = "And C.部门ID IN(" & strSQL & ")"
    Else
        '当前选择的科室的医生
        lngDeptID = cbo科室.ItemData(cbo科室.ListIndex)
        strSQL = "And C.部门ID =[2]"
    End If
    
    strSQL = "Select Distinct A.ID,A.编号,A.姓名,A.简码" & _
        " From 人员表 A,人员性质说明 B,部门人员 C" & _
        " Where A.ID=B.人员ID And B.人员性质='医生' And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) " & _
        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
        " And A.ID=C.人员ID " & strSQL & _
        " Order by A.编号"
    vRect = zlcontrol.GetControlRect(txtDoct.hWnd)
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "医生", False, "", "", False, False, True, vRect.Left, vRect.Top, txtDoct.Height, blnCancel, False, True, UserInfo.ID, lngDeptID)
    If rsTemp Is Nothing Then
        If Not blnCancel Then
            MsgBox "没有设置门诊医生，请先到人员管理中设置。", vbInformation, gstrSysName
        End If
        txtDoct.SetFocus
    Else
        txtDoct.Text = rsTemp!姓名
        txtDoct.Tag = rsTemp!ID
        zlcontrol.ControlSetFocus chk有效期, True
    End If
End Sub


Private Sub dtpBegin_Change()
    Err = 0: On Error Resume Next
    If dtpEnd.Value < dtpBegin.Value Then dtpEnd.Value = dtpBegin.Value
End Sub

Private Sub dtpEnd_Change()
    Err = 0: On Error Resume Next
    If dtpBegin.Value > dtpEnd.Value Then dtpBegin.Value = dtpEnd.Value
End Sub

Private Sub dtpEndDate_Change(Index As Integer)
    Err = 0: On Error Resume Next
    If Index <> 0 Then
        If dtpEndDate(Index).Value > dtpStartDate(Index).MaxDate Then dtpEndDate(Index).Value = dtpStartDate(Index).MaxDate
    End If
    If dtpEndDate(Index).Value < dtpStartDate(Index).Value Then dtpStartDate(Index).Value = dtpEndDate(Index).Value
End Sub

Private Sub dtpStartDate_Change(Index As Integer)
    Err = 0: On Error Resume Next
    If Index <> 0 Then
        If dtpStartDate(Index).Value > dtpEndDate(Index).MaxDate Then dtpStartDate(Index).Value = dtpEndDate(Index).MaxDate
    End If
    If dtpStartDate(Index).Value > dtpEndDate(Index).Value Then dtpEndDate(Index).Value = dtpStartDate(Index).Value
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Me.ActiveControl Is cbo科室 Then Exit Sub
    
    Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Or KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    mlngModule = glngModul: mstrPrivs = gstrPrivs
    zlcontrol.CboSetHeight cbo科室, 5000
    Call InitCon
    Call SetControlLoacle
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    cmd刷新.Left = Me.ScaleLeft + Me.ScaleWidth - cmd刷新.Width - 100
End Sub

Private Sub Form_Unload(Cancel As Integer)
    zlDatabase.SetPara "显示失效计划", chkShowExpiredPlan.Value, glngSys, mlngModule
End Sub

Private Sub txtDoct_Change()
    If txtDoct.Text = "" Then txtDoct.Tag = ""
End Sub

Private Sub txtDoct_GotFocus()
    zlcontrol.TxtSelAll txtDoct
End Sub

Private Sub txtDoct_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
        Call cmdDoct_Click
    End If
End Sub
Private Sub chkDate_Click(Index As Integer)
    dtpStartDate(Index).Enabled = chkDate(Index).Value = 1
    dtpEndDate(Index).Enabled = chkDate(Index).Value = 1
    If Index = 2 Then
        chk计划(1).Value = 0: chk计划(1).Enabled = chkDate(Index).Value <> 1
    End If
End Sub

 

Private Sub txtDoct_Validate(Cancel As Boolean)
    Dim rsTemp As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    Dim strSQL As String, lngDeptID As Long
    
    If txtDoct.Text <> "" Then
        If cbo科室.ListIndex < 0 Then cbo科室.ListIndex = 0
        If cbo科室.ItemData(cbo科室.ListIndex) <= 0 Then
            '所有门诊临床科室的医生
            strSQL = "Select Distinct A.ID From 部门表 A,部门性质说明 B" & _
                " Where A.ID=B.部门ID And B.工作性质='临床' And B.服务对象 IN(1,3) " & _
            IIf(zlStr.IsHavePrivs(mstrPrivs, "所有科室") = False, " And Exists(Select 1 From 部门人员 where 人员id=[1] and A.id=部门id) ", "")
            strSQL = "And C.部门ID IN(" & strSQL & ")"
        Else
            '当前选择的科室的医生
            lngDeptID = cbo科室.ItemData(cbo科室.ListIndex)
            strSQL = " And C.部门ID=[2] "
        End If
        
        strSQL = "Select Distinct A.ID,A.编号,A.姓名,A.简码" & _
        " From 人员表 A,人员性质说明 B,部门人员 C" & _
        " Where A.ID=B.人员ID And B.人员性质='医生' And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) " & _
        "       And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
        "       And A.ID=C.人员ID  " & strSQL & "" & _
        "       And (A.编号 Like [3] Or A.姓名 Like [4] Or A.简码 Like [4] )" & _
        " Order by A.编号"
        
        vRect = zlcontrol.GetControlRect(txtDoct.hWnd)
        Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "医生", False, "", "", False, False, True, vRect.Left, vRect.Top, txtDoct.Height, blnCancel, False, True, UserInfo.ID, lngDeptID, UCase(txtDoct.Text) & "%", "%" & UCase(txtDoct.Text) & "%")
        If rsTemp Is Nothing Then
            If Not blnCancel Then txtDoct.Tag = ""
            
        Else
            txtDoct.Text = rsTemp!姓名
            txtDoct.Tag = rsTemp!ID
            zlcontrol.ControlSetFocus chk有效期, True
         End If
    End If
End Sub

Private Sub SetControlLoacle()
    Err = 0: On Error Resume Next
    '设置控件的位置
    If mblnPlanPage Then
        fra(0).Height = 2300 + picCon.Height
        picCon.Visible = True
    Else
        fra(0).Height = 2200
        picCon.Visible = False
    End If
    cmd刷新.Top = fra(0).Top + fra(0).Height + 50
End Sub

Public Property Get zlblnShowPlanCon() As Boolean
    zlblnShowPlanCon = mblnPlanPage
End Property

Public Property Let zlblnShowPlanCon(ByVal vNewValue As Boolean)
    mblnPlanPage = vNewValue
    Call SetControlLoacle
End Property

Public Property Get zlGet科室ID() As Long
    '获取科室ID
    If cbo科室.ListIndex < 0 Then
        zlGet科室ID = 0
    Else
        zlGet科室ID = cbo科室.ItemData(cbo科室.ListIndex)
    End If
End Property
 
Public Property Let ShowStop(ByVal vNewValue As Boolean)
'显示停用安排 属性
      If Not mblnShowStoped = vNewValue Then
          mblnShowStoped = vNewValue
          'RaiseEvent zlRefreshCon(mArrFilter)
     End If
End Property

Public Property Let ShowDel(ByVal vNewValue As Boolean)
'显示删除安排 属性
      If Not mblnShowDel = vNewValue Then
          mblnShowDel = vNewValue
     End If
End Property
