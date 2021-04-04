VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmIDCardFilter 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "过滤设置"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6195
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdDef 
      Caption         =   "缺省(&D)"
      Height          =   350
      Left            =   4905
      TabIndex        =   24
      Top             =   1665
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   2625
      Left            =   105
      TabIndex        =   13
      Top             =   15
      Width           =   4635
      Begin VB.TextBox txt姓名 
         Height          =   300
         IMEMode         =   1  'ON
         Left            =   2880
         TabIndex        =   7
         Top             =   1440
         Width           =   1470
      End
      Begin VB.TextBox txt住院号 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   960
         MaxLength       =   18
         TabIndex        =   6
         Top             =   1440
         Width           =   1245
      End
      Begin VB.TextBox txtNOBegin 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   960
         TabIndex        =   2
         Top             =   660
         Width           =   1245
      End
      Begin VB.TextBox txtNOEnd 
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   2880
         TabIndex        =   3
         Top             =   660
         Width           =   1470
      End
      Begin VB.TextBox txtCardBegin 
         BackColor       =   &H00EBFFFF&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   960
         TabIndex        =   4
         Top             =   1050
         Width           =   1245
      End
      Begin VB.TextBox txtCardEnd 
         BackColor       =   &H00EBFFFF&
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   2880
         TabIndex        =   5
         Top             =   1050
         Width           =   1470
      End
      Begin VB.ComboBox cbo操作员 
         Height          =   300
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1845
         Width           =   1485
      End
      Begin VB.ComboBox cboType 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1845
         Width           =   1275
      End
      Begin VB.CheckBox chkCancel 
         Caption         =   "退卡记录"
         Height          =   210
         Left            =   960
         TabIndex        =   10
         Top             =   2265
         Width           =   1020
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   2880
         TabIndex        =   1
         Top             =   270
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CheckBox        =   -1  'True
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   112525315
         CurrentDate     =   36992
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Left            =   960
         TabIndex        =   0
         Top             =   270
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   112525315
         CurrentDate     =   36992
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名"
         Height          =   180
         Left            =   2460
         TabIndex        =   23
         Top             =   1500
         Width           =   360
      End
      Begin VB.Label lbl住院号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院号"
         Height          =   180
         Left            =   345
         TabIndex        =   22
         Top             =   1500
         Width           =   540
      End
      Begin VB.Label lbl时间 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "发卡时间"
         Height          =   180
         Left            =   165
         TabIndex        =   21
         Top             =   330
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "至"
         Height          =   180
         Left            =   2460
         TabIndex        =   20
         Top             =   330
         Width           =   180
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单据号"
         Height          =   180
         Left            =   345
         TabIndex        =   19
         Top             =   720
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "至"
         Height          =   180
         Left            =   2460
         TabIndex        =   18
         Top             =   720
         Width           =   180
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "卡号"
         Height          =   180
         Left            =   525
         TabIndex        =   17
         Top             =   1110
         Width           =   360
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "至"
         Height          =   180
         Left            =   2460
         TabIndex        =   16
         Top             =   1110
         Width           =   180
      End
      Begin VB.Label lbl操作员 
         Caption         =   "发卡人"
         Height          =   165
         Left            =   2310
         TabIndex        =   15
         Top             =   1905
         Width           =   540
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "类型"
         Height          =   180
         Left            =   510
         TabIndex        =   14
         Top             =   1905
         Width           =   360
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4905
      TabIndex        =   12
      Top             =   795
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4905
      TabIndex        =   11
      Top             =   390
      Width           =   1100
   End
End
Attribute VB_Name = "frmIDCardFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Public mstrFilter As String
Public mblnDateMoved As Boolean 'in/Out
Public mcllFilter As Collection

Private Sub cboType_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cboType.hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cboType.hwnd, KeyAscii)
    If lngIdx <> -2 Then cboType.ListIndex = lngIdx
    If cboType.ListIndex = -1 And cboType.ListCount <> 0 Then cboType.ListIndex = 0
End Sub

Private Sub cbo操作员_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo操作员.hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo操作员.hwnd, KeyAscii)
    If lngIdx <> -2 Then cbo操作员.ListIndex = lngIdx
    If cbo操作员.ListIndex = -1 And cbo操作员.ListCount <> 0 Then cbo操作员.ListIndex = 0
End Sub

Private Sub chkCancel_Click()
    If chkCancel.Value = 0 Then
        lbl操作员.Caption = "发卡人"
        lbl时间.Caption = "发卡时间"
    Else
        lbl操作员.Caption = "退卡人"
        lbl时间.Caption = "退卡时间"
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
    If Not IsNull(dtpEnd.Value) Then
        If dtpEnd.Value < dtpBegin.Value Then
            MsgBox "结束时间不能小于开始时间！", vbInformation, gstrSysName
            dtpEnd.SetFocus: Exit Sub
        End If
    End If
    If txtNOBegin.Text <> "" And txtNoEnd.Text <> "" Then
        If txtNoEnd.Text < txtNOBegin.Text Then
            MsgBox "结束单据号不能小于开始单据号！", vbInformation, gstrSysName
            txtNoEnd.SetFocus: Exit Sub
        End If
    End If
    If txtCardBegin.Text <> "" And txtCardEnd.Text <> "" Then
        If txtCardEnd.Text < txtCardBegin.Text Then
            MsgBox "结束卡号不能小于开始卡号！", vbInformation, gstrSysName
            txtCardEnd.SetFocus: Exit Sub
        End If
    End If
    
    Call MakeFilter
    
    gblnOK = True
    Hide
End Sub

Private Sub Form_Activate()
    dtpBegin.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':：;；?？", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    Dim curDate As Date, i As Integer
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errH
    
    gblnOK = False
    
    If glngSys Like "8??" Then
        lbl住院号.Visible = False
        txt住院号.Visible = False
    End If
    
    '设置初始值
    If Not gblnShowCard Then
        txtCardBegin.PasswordChar = "*"
        txtCardEnd.PasswordChar = "*"
    End If
    'txtCardBegin.MaxLength = gbytCardNOLen
    'txtCardEnd.MaxLength = gbytCardNOLen
    
    curDate = zldatabase.Currentdate
    dtpBegin.Value = Format(DateAdd("m", -1, curDate), "yyyy-MM-dd 00:00")
    dtpEnd.Value = Format(curDate, "yyyy-MM-dd 23:59")
    dtpBegin.MaxDate = dtpEnd.Value
    
    cboType.Clear
    cboType.AddItem "0-所有"
    cboType.AddItem "1-发卡"
    cboType.AddItem "2-补卡"
    cboType.AddItem "3-换卡"
    cboType.ListIndex = 0
    
    cbo操作员.Clear
    cbo操作员.AddItem "所有操作员"
    cbo操作员.ListIndex = 0
        
    Set rsTmp = GetPersonnel("发卡登记人", True)
    For i = 1 To rsTmp.RecordCount
        cbo操作员.AddItem rsTmp!简码 & "-" & rsTmp!姓名
        If rsTmp!ID = UserInfo.ID Then cbo操作员.ListIndex = cbo操作员.NewIndex
        rsTmp.MoveNext
    Next
    Exit Sub
errH:
    If errCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtCardBegin_GotFocus()
    SelAll txtCardBegin
End Sub

Private Sub txtCardBegin_KeyPress(KeyAscii As Integer)
'    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If InStr(":：;；?？'‘||", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    
    If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
        If Len(txtCardBegin.Text) = gbytCardNOLen - 1 Then
            txtCardBegin.Text = txtCardBegin.Text & Chr(KeyAscii)
            KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
        End If
    ElseIf KeyAscii = vbKeyReturn Then
        KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txtCardEnd_GotFocus()
    SelAll txtCardEnd
End Sub

Private Sub txtCardEnd_KeyPress(KeyAscii As Integer)
'    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If InStr(":：;；?？'‘||", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    
    If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
        If Len(txtCardEnd.Text) = gbytCardNOLen - 1 Then
            txtCardEnd.Text = txtCardEnd.Text & Chr(KeyAscii)
            KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
        End If
    ElseIf KeyAscii = vbKeyReturn Then
        KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txtCardBegin_Change()
    txtCardEnd.Enabled = Not (Trim(txtCardBegin.Text) = "")
    If Trim(txtCardBegin.Text = "") Then txtCardEnd.Text = ""
End Sub

Private Sub txtNOBegin_Change()
    txtNoEnd.Enabled = Not (Trim(txtNOBegin.Text) = "")
    If Trim(txtNOBegin.Text = "") Then txtNoEnd.Text = ""
End Sub

Private Sub txtNOBegin_GotFocus()
    SelAll txtNOBegin
End Sub

Private Sub txtNOBegin_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii <> 13 Then
        If Not (txtNOBegin.Text = "" Or txtNOBegin.SelLength = Len(txtNOBegin.Text)) And _
            InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0: Beep: Exit Sub
        End If
    End If
End Sub

Private Sub txtNOBegin_LostFocus()
    If txtNOBegin.Text <> "" Then txtNOBegin.Text = GetFullNO(txtNOBegin.Text, 16)
End Sub

Private Sub txtNOEnd_LostFocus()
    If txtNoEnd.Text <> "" Then txtNoEnd.Text = GetFullNO(txtNoEnd.Text, 16)
End Sub

Private Sub txtNoEnd_GotFocus()
    SelAll txtNoEnd
End Sub

Private Sub txtNoEnd_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii <> 13 Then
        If Not (txtNoEnd.Text = "" Or txtNoEnd.SelLength = Len(txtNoEnd.Text)) And _
            InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0: Beep: Exit Sub
        End If
    End If
End Sub

Private Sub MakeFilter()
    Dim strSQL As String
    Dim strSQLtmp As String
    Dim strTmp As String
    'by lesfeng 2010-03-08 性能优化
    Set mcllFilter = New Collection
    mcllFilter.Add Array("1901-01-01 00:00:00", "1901-01-01 00:00:00"), "登记时间"
    mcllFilter.Add Array("", ""), "单据号"
    mcllFilter.Add Array("", ""), "票据号"
    mcllFilter.Add "", "住院号"
    mcllFilter.Add "", "姓名"
    mcllFilter.Add "", "记录状态"
    mcllFilter.Add "", "附加标志"
    mcllFilter.Add "", "收款人"
    
    mstrFilter = ""
    mstrFilter = mstrFilter & " And (登记时间  Between [1] And [2]) "
   
    If IsNull(dtpEnd.Value) Then
        '该天内
        mcllFilter.Remove "登记时间"
        mcllFilter.Add Array(Format(dtpBegin.Value, "yyyy-MM-dd") & " 00:00:00", Format(dtpBegin.Value, "yyyy-MM-dd") & " 23:59:59"), "登记时间"
    Else
        '范围内
        mcllFilter.Remove "登记时间"
        mcllFilter.Add Array(Format(dtpBegin.Value, "yyyy-MM-dd") & " 00:00:00", Format(dtpEnd.Value, "yyyy-MM-dd") & " 23:59:59"), "登记时间"
    End If
    
    mblnDateMoved = zldatabase.DateMoved(Format(IIf(dtpBegin.Value < dtpEnd.Value, dtpBegin.Value, dtpEnd.Value), dtpBegin.CustomFormat), , , Me.Caption)
    
    If txtNOBegin.Text <> "" And txtNoEnd.Text <> "" Then
        mstrFilter = mstrFilter & " And NO Between [3] And [4] "
    ElseIf txtNOBegin.Text <> "" Then
        mstrFilter = mstrFilter & " And NO=[3] "
    End If
    
    mcllFilter.Remove "单据号"
    mcllFilter.Add Array(Trim(txtNOBegin.Text), Trim(txtNoEnd.Text)), "单据号"
    
    If txtCardBegin.Text <> "" And txtCardEnd.Text <> "" Then
        mstrFilter = mstrFilter & " And 实际票号 Between [5] And [6] "
    ElseIf txtCardBegin.Text <> "" Then
        mstrFilter = mstrFilter & " And 实际票号=[5] "
    End If
    
    mcllFilter.Remove "票据号"
    mcllFilter.Add Array(Trim(txtCardBegin.Text), Trim(txtCardEnd.Text)), "票据号"
    
    If txt住院号.Text <> "" Then
        mstrFilter = mstrFilter & " And 标识号=[7]"
    End If
    If txt姓名.Text <> "" Then
        mstrFilter = mstrFilter & " And Upper(姓名) Like [8]"
    End If
    If chkCancel.Value = Checked Then
        '直接查看退卡记录
        mstrFilter = mstrFilter & " And 记录状态=[9]"
        strTmp = 2
    Else
        mstrFilter = mstrFilter & " And 记录状态=[9]"
        strTmp = 1
    End If
    If cboType.ListIndex <> 0 Then mstrFilter = mstrFilter & " And Nvl(附加标志,0)=[10]"
    If cbo操作员.ListIndex <> 0 Then mstrFilter = mstrFilter & " And 操作员姓名=[11]"
    
    mcllFilter.Remove "住院号"
    mcllFilter.Add Trim(txt住院号.Text), "住院号"
    mcllFilter.Remove "姓名"
    mcllFilter.Add "%" & Trim(UCase(txt姓名.Text)) & "%", "姓名"
    mcllFilter.Remove "记录状态"
    mcllFilter.Add strTmp, "记录状态"
    mcllFilter.Remove "附加标志"
    mcllFilter.Add cboType.ListIndex - 1, "附加标志"
    mcllFilter.Remove "收款人"
    mcllFilter.Add Trim(NeedName(cbo操作员.Text)), "收款人"
End Sub

Private Sub txt姓名_GotFocus()
    SelAll txt姓名
End Sub

Private Sub txt住院号_GotFocus()
    SelAll txt住院号
End Sub

Private Sub txt住院号_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

