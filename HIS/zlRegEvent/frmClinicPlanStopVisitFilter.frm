VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmClinicPlanStopVisitFilter 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "过滤设置"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdAudit 
      Height          =   240
      Left            =   2670
      Picture         =   "frmClinicPlanStopVisitFilter.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "选择项目(F4)"
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton cmdApply 
      Height          =   240
      Left            =   2670
      Picture         =   "frmClinicPlanStopVisitFilter.frx":00F6
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "选择项目(F4)"
      Top             =   330
      Width           =   255
   End
   Begin VB.TextBox txtAudit 
      Height          =   300
      Left            =   885
      MaxLength       =   8
      TabIndex        =   4
      Top             =   810
      Width           =   2070
   End
   Begin VB.CheckBox chkAudit 
      Caption         =   "已审批"
      Height          =   255
      Index           =   1
      Left            =   3300
      TabIndex        =   5
      Top             =   833
      Value           =   1  'Checked
      Width           =   885
   End
   Begin VB.CheckBox chkAudit 
      Caption         =   "未审批"
      Height          =   255
      Index           =   0
      Left            =   3300
      TabIndex        =   2
      Top             =   323
      Value           =   1  'Checked
      Width           =   855
   End
   Begin VB.Frame frmSplit 
      Height          =   25
      Left            =   -270
      TabIndex        =   11
      Top             =   2280
      Width           =   6525
   End
   Begin VB.CheckBox chkShowInvalid 
      Caption         =   "显示已失效停诊安排"
      Height          =   255
      Left            =   885
      TabIndex        =   6
      Top             =   1320
      Width           =   2025
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4410
      TabIndex        =   13
      Top             =   2520
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3150
      TabIndex        =   12
      Top             =   2520
      Width           =   1100
   End
   Begin MSComCtl2.DTPicker dtpStopBegin 
      Height          =   300
      Left            =   885
      TabIndex        =   8
      Top             =   1755
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      CalendarTitleBackColor=   -2147483647
      CalendarTitleForeColor=   -2147483634
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   159252483
      CurrentDate     =   36588
   End
   Begin MSComCtl2.DTPicker dtpStopEnd 
      Height          =   300
      Left            =   3300
      TabIndex        =   10
      Top             =   1755
      Width           =   2070
      _ExtentX        =   3651
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      CalendarTitleBackColor=   -2147483647
      CalendarTitleForeColor=   -2147483634
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   159252483
      CurrentDate     =   36588
   End
   Begin VB.TextBox txtApply 
      Height          =   300
      Left            =   885
      MaxLength       =   8
      TabIndex        =   1
      Top             =   300
      Width           =   2070
   End
   Begin VB.Label lblStopTimeRange 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "停诊时间"
      Height          =   180
      Left            =   90
      TabIndex        =   7
      Top             =   1815
      Width           =   720
   End
   Begin VB.Label lblApply 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "申请人"
      Height          =   180
      Left            =   270
      TabIndex        =   0
      Top             =   360
      Width           =   540
   End
   Begin VB.Label lblAudit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "审批人"
      Height          =   180
      Left            =   270
      TabIndex        =   3
      Top             =   870
      Width           =   540
   End
   Begin VB.Label lbl审核时间范围 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "至"
      Height          =   180
      Left            =   3000
      TabIndex        =   9
      Top             =   1815
      Width           =   180
   End
End
Attribute VB_Name = "frmClinicPlanStopVisitFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Public mstrFilter As String
Public mblnOK As Boolean
Private mrsApply As ADODB.Recordset
Private mrsAudit As ADODB.Recordset

Private Sub chkAudit_Click(index As Integer)
    If chkAudit(0).Value = vbUnchecked And chkAudit(1).Value = vbUnchecked Then
        chkAudit(IIf(index = 0, 1, 0)).Value = vbChecked
    End If
    txtAudit.Enabled = chkAudit(1).Value = vbChecked
    txtAudit.BackColor = IIf(txtAudit.Enabled, vbWindowBackground, vbButtonFace)
    cmdAudit.Enabled = chkAudit(1).Value = vbChecked
End Sub

Private Sub chkAudit_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chkShowInvalid_Click()
    dtpStopBegin.Enabled = chkShowInvalid.Value = vbChecked
    dtpStopEnd.Enabled = chkShowInvalid.Value = vbChecked
End Sub

Private Sub chkShowInvalid_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False
    If txtApply.Visible And txtApply.Enabled Then txtApply.SetFocus
    Me.Hide '隐藏窗口
End Sub

Private Sub cmdOK_Click()
    Err = 0: On Error GoTo ErrHandler
    If zlControl.FormCheckInput(Me) = False Then Exit Sub
    If chkShowInvalid.Value = vbChecked Then
        If DateDiff("s", dtpStopBegin.Value, dtpStopEnd.Value) <= 0 Then
            MsgBox "停诊时间的结束时间必须大于开始时间！", vbInformation, gstrSysName
            If dtpStopBegin.Visible And dtpStopBegin.Enabled Then dtpStopBegin.SetFocus
            Exit Sub
        End If
    End If
    
    Call MakeFilter
    mblnOK = True
    
    If txtApply.Visible And txtApply.Enabled Then txtApply.SetFocus
    Me.Hide '隐藏窗口
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub dtpStopBegin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub dtpStopEnd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub Form_Load()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Integer, lngOldID As Long, strListFeeItem As String
    Dim Curdate As Date, index As Integer, arrItem As Variant
    
    Err = 0: On Error GoTo ErrHandler
    mblnOK = False
    Curdate = zlDatabase.Currentdate
    
    dtpStopBegin.Value = Format(DateAdd("m", -1, Curdate), "yyyy-MM-dd 00:00:00")
    dtpStopEnd.Value = Format(DateAdd("m", 1, Curdate), "yyyy-MM-dd 23:59:59")
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub MakeFilter()
    Err = 0: On Error GoTo ErrHandler
    
    mstrFilter = ""
    If Trim(txtApply.Text) <> "" Then mstrFilter = mstrFilter & " And 申请人=[1]"
    If Trim(txtAudit.Text) <> "" And chkAudit(1).Value = vbChecked Then mstrFilter = mstrFilter & " And 审批人=[2] "
    
    If chkAudit(0).Value = vbChecked And chkAudit(1).Value = vbChecked Then
    ElseIf chkAudit(0).Value = vbChecked Then '仅未审批
        mstrFilter = mstrFilter & " And (审批人 Is Null Or 取消人 Is Not Null)"
    Else '仅已审批
        mstrFilter = mstrFilter & " And 审批人 Is Not Null And 取消人 Is Null"
    End If
    
    If chkShowInvalid.Value = vbChecked Then
        mstrFilter = mstrFilter & " And 开始时间<=[4] And Nvl(失效时间,终止时间)>=[3]"
    Else
        mstrFilter = mstrFilter & " And Nvl(失效时间,终止时间)>sysdate"
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mrsApply Is Nothing Then Set mrsApply = Nothing
    If Not mrsAudit Is Nothing Then Set mrsAudit = Nothing
End Sub

Private Sub txtApply_GotFocus()
    zlControl.TxtSelAll txtApply
End Sub

Private Sub txtApply_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtApply_Validate(Cancel As Boolean)
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    
    Err = 0: On Error GoTo ErrHandler
    If Trim(txtApply.Text) = "" Then Exit Sub
    
    strSQL = "Select Distinct A.ID,A.编号,A.姓名,A.简码" & vbNewLine & _
            " From 人员表 A,人员性质说明 B" & vbNewLine & _
            " Where A.ID=B.人员ID And B.人员性质='医生'" & vbNewLine & _
            "       And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) " & vbNewLine & _
            "       And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & vbNewLine & _
            "       And (A.编号 Like [1] Or Upper(A.姓名) Like [2] Or A.简码 Like [2] )" & vbNewLine & _
            " Order by A.编号"
    
    vRect = zlControl.GetControlRect(txtApply.hwnd)
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "申请人", False, "", "", False, False, True, _
        vRect.Left, vRect.Top, txtApply.Height, blnCancel, False, True, Trim(txtApply.Text) & "%", gstrLike & UCase(Trim(txtApply.Text)) & "%")
    If rsTemp Is Nothing Then Cancel = True: Exit Sub
    If rsTemp.EOF Then Cancel = True: Exit Sub
    If blnCancel Then Cancel = True: Exit Sub
    
    txtApply.Text = rsTemp!姓名
    If chkAudit(0).Visible And chkAudit(0).Enabled Then chkAudit(0).SetFocus
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdApply_Click()
    Dim rsTemp As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    Dim strSQL As String, lngDeptID As Long
    
    Err = 0: On Error GoTo ErrHandler
    strSQL = "Select Distinct A.ID,A.编号,A.姓名,A.简码" & vbNewLine & _
            " From 人员表 A,人员性质说明 B" & vbNewLine & _
            " Where A.ID=B.人员ID And B.人员性质='医生'" & vbNewLine & _
            "       And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) " & vbNewLine & _
            "       And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & vbNewLine & _
            " Order by A.编号"
    
    vRect = zlControl.GetControlRect(txtApply.hwnd)
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "申请人", False, "", "", False, False, True, _
        vRect.Left, vRect.Top, txtApply.Height, blnCancel, False, True)
    If rsTemp Is Nothing Then Exit Sub
    If rsTemp.EOF Then Exit Sub
    If blnCancel Then Exit Sub
    
    txtApply.Text = rsTemp!姓名
    If chkAudit(0).Visible And chkAudit(0).Enabled Then chkAudit(0).SetFocus
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtAudit_GotFocus()
    zlControl.TxtSelAll txtAudit
End Sub

Private Sub txtAudit_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtAudit_Validate(Cancel As Boolean)
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    
    Err = 0: On Error GoTo ErrHandler
    If Trim(txtAudit.Text) = "" Then Exit Sub
    
    strSQL = "Select Distinct d.ID, d.编号, d.简码, d.姓名" & vbNewLine & _
            " From 人员表 D,上机人员表 C, Zluserroles B, zlRoleGrant A" & vbNewLine & _
            " Where A.角色 = B.角色 And B.用户 = C.用户名 And C.人员id = D.ID" & vbNewLine & _
            "       And A.系统 = 100 And A.序号 = 1114 And A.功能 = '停诊审批'" & vbNewLine & _
            "       And (d.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or d.撤档时间 Is Null) " & vbNewLine & _
            "       And (d.站点='" & gstrNodeNo & "' Or d.站点 is Null)" & vbNewLine & _
            "       And (d.编号 Like [1] Or Upper(d.姓名) Like [2] Or d.简码 Like [2] )" & vbNewLine & _
            " Order by d.编号"

    vRect = zlControl.GetControlRect(txtAudit.hwnd)
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "审批人", False, "", "", False, False, True, _
        vRect.Left, vRect.Top, txtAudit.Height, blnCancel, False, True, Trim(txtAudit.Text) & "%", gstrLike & UCase(Trim(txtAudit.Text)) & "%")
    If rsTemp Is Nothing Then Cancel = True: Exit Sub
    If rsTemp.EOF Then Cancel = True: Exit Sub
    If blnCancel Then Cancel = True: Exit Sub
    
    txtAudit.Text = rsTemp!姓名
    If chkAudit(1).Visible And chkAudit(1).Enabled Then chkAudit(1).SetFocus
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdAudit_Click()
    Dim rsTemp As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    Dim strSQL As String, lngDeptID As Long
    
    Err = 0: On Error GoTo ErrHandler
    strSQL = "Select Distinct d.ID, d.编号, d.简码, d.姓名" & vbNewLine & _
            " From 人员表 D,上机人员表 C, Zluserroles B, zlRoleGrant A" & vbNewLine & _
            " Where A.角色 = B.角色 And B.用户 = C.用户名 And C.人员id = D.ID" & vbNewLine & _
            "       And A.系统 = 100 And A.序号 = 1114 And A.功能 = '停诊审批'" & vbNewLine & _
            "       And (d.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or d.撤档时间 Is Null) " & vbNewLine & _
            "       And (d.站点='" & gstrNodeNo & "' Or d.站点 is Null)" & vbNewLine & _
            " Order by d.编号"
    
    vRect = zlControl.GetControlRect(txtAudit.hwnd)
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "审批人", False, "", "", False, False, True, _
        vRect.Left, vRect.Top, txtAudit.Height, blnCancel, False, True)
    If rsTemp Is Nothing Then Exit Sub
    If rsTemp.EOF Then Exit Sub
    If blnCancel Then Exit Sub
    
    txtAudit.Text = rsTemp!姓名
    If chkAudit(1).Visible And chkAudit(1).Enabled Then chkAudit(1).SetFocus
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

