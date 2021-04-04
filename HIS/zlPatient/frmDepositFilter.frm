VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDepositFilter 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "过滤设置"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdDef 
      Caption         =   "缺省(&D)"
      Height          =   350
      Left            =   4800
      TabIndex        =   13
      Top             =   2130
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4800
      TabIndex        =   11
      Top             =   330
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4800
      TabIndex        =   12
      Top             =   750
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   3945
      Left            =   105
      TabIndex        =   14
      Top             =   15
      Width           =   4590
      Begin VB.TextBox txtClinicNO 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   2955
         MaxLength       =   18
         TabIndex        =   26
         Top             =   1920
         Width           =   1455
      End
      Begin VB.CheckBox chk收款 
         Caption         =   "收款记录"
         Height          =   210
         Left            =   2760
         TabIndex        =   25
         Top             =   3180
         Width           =   1020
      End
      Begin VB.TextBox txt结算号码 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1005
         MaxLength       =   30
         TabIndex        =   8
         Top             =   2730
         Width           =   3240
      End
      Begin VB.TextBox txtFactEnd 
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   2955
         TabIndex        =   5
         Top             =   1500
         Width           =   1455
      End
      Begin VB.TextBox txtFactBegin 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1005
         TabIndex        =   4
         Top             =   1500
         Width           =   1260
      End
      Begin VB.TextBox txt住院号 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1005
         MaxLength       =   18
         TabIndex        =   6
         Top             =   1920
         Width           =   1260
      End
      Begin VB.TextBox txt姓名 
         Height          =   300
         IMEMode         =   1  'ON
         Left            =   1005
         TabIndex        =   7
         Top             =   2325
         Width           =   1260
      End
      Begin VB.TextBox txtNoEnd 
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   2955
         TabIndex        =   3
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txtNOBegin 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1005
         TabIndex        =   2
         Top             =   1080
         Width           =   1260
      End
      Begin VB.ComboBox cbo操作员 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1005
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   3135
         Width           =   1275
      End
      Begin VB.CheckBox chkCancel 
         Caption         =   "退款记录"
         Height          =   210
         Left            =   2760
         TabIndex        =   10
         Top             =   3525
         Width           =   1020
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   1005
         TabIndex        =   1
         Top             =   675
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   105185283
         CurrentDate     =   36588
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Left            =   1005
         TabIndex        =   0
         Top             =   270
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   105185283
         CurrentDate     =   36588
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "门诊号"
         Height          =   180
         Left            =   2310
         TabIndex        =   27
         Top             =   1980
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结算号码"
         Height          =   180
         Left            =   180
         TabIndex        =   24
         Top             =   2790
         Width           =   720
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "票据号"
         Height          =   180
         Left            =   360
         TabIndex        =   23
         Top             =   1575
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "至"
         Height          =   180
         Left            =   2490
         TabIndex        =   22
         Top             =   1560
         Width           =   180
      End
      Begin VB.Label lbl操作员 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "收款人"
         Height          =   180
         Left            =   360
         TabIndex        =   21
         Top             =   3195
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院号"
         Height          =   180
         Left            =   360
         TabIndex        =   20
         Top             =   1965
         Width           =   540
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名"
         Height          =   180
         Left            =   540
         TabIndex        =   19
         Top             =   2385
         Width           =   360
      End
      Begin VB.Label lblBegin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "开始时间"
         Height          =   180
         Left            =   180
         TabIndex        =   18
         Top             =   330
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单据号"
         Height          =   180
         Left            =   360
         TabIndex        =   17
         Top             =   1155
         Width           =   540
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "至"
         Height          =   180
         Left            =   2490
         TabIndex        =   16
         Top             =   1140
         Width           =   180
      End
      Begin VB.Label lblEnd 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结束时间"
         Height          =   180
         Left            =   180
         TabIndex        =   15
         Top             =   735
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmDepositFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Public mstrFilter As String
Public mblnDateMoved As Boolean 'in/Out
Public mcllFilter As Collection

Private Sub cbo操作员_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo操作员.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = cbo.MatchIndex(cbo操作员.hWnd, KeyAscii, 0.5)
    If lngIdx <> -2 Then cbo操作员.ListIndex = lngIdx
    If cbo操作员.ListIndex = -1 And cbo操作员.ListCount <> 0 Then cbo操作员.ListIndex = 0
End Sub

Private Sub chkCancel_Click()
    If chkCancel.Value = 1 Then
        lbl操作员.Caption = "退款人"
    Else
        lbl操作员.Caption = "收款人"
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
    If dtpEnd.Value < dtpBegin.Value Then
        MsgBox "结束时间不能小于开始时间！", vbInformation, gstrSysName
        dtpEnd.SetFocus: Exit Sub
    End If

    If txtNOBegin.Text <> "" And txtNoEnd.Text <> "" Then
        If txtNoEnd.Text < txtNOBegin.Text Then
            MsgBox "结束单据号不能小于开始单据号！", vbInformation, gstrSysName
            txtNoEnd.SetFocus: Exit Sub
        End If
    End If
    If txtFactBegin.Text <> "" And txtFactEnd.Text <> "" Then
        If txtFactEnd.Text < txtFactBegin.Text Then
            MsgBox "结束票据号不能小于开始票据号！", vbInformation, gstrSysName
            txtFactEnd.SetFocus: Exit Sub
        End If
    End If
    
    If chk收款.Value = 0 And chkCancel.Value = 0 Then
        MsgBox "收款记录、退款记录选项至少应该选择一个！", vbInformation, gstrSysName
        chk收款.SetFocus
        Exit Sub
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
    If KeyAscii = Asc("'") Then KeyAscii = 0
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim Curdate As Date, i As Integer
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errH
    
    gblnOK = False
    
    txtNOBegin.Text = ""
    txtNoEnd.Text = ""
    txtFactBegin.Text = ""
    txtFactEnd.Text = ""
    txt姓名.Text = ""
    txt住院号.Text = ""
    chkCancel.Value = 0
    chk收款.Value = 1
    '设置初始值
    'txtFactBegin.MaxLength = gbyt预交
    'txtFactEnd.MaxLength = gbyt预交
    
    Curdate = zlDatabase.Currentdate
    dtpBegin.Value = Format(Curdate, "yyyy-MM-dd 00:00:00")
    dtpEnd.Value = Format(Curdate, "yyyy-MM-dd 23:59:59")
    dtpBegin.MaxDate = dtpEnd.Value
    
    cbo操作员.Clear
    cbo操作员.AddItem "所有操作员"
    cbo操作员.ListIndex = 0
    
    Set rsTmp = GetPersonnel("预交收款员", True)
    For i = 1 To rsTmp.RecordCount
        cbo操作员.AddItem rsTmp!简码 & "-" & rsTmp!姓名
        If rsTmp!ID = UserInfo.ID Then cbo操作员.ListIndex = cbo操作员.NewIndex
        rsTmp.MoveNext
    Next
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtFactBegin_GotFocus()
    zlControl.TxtSelAll txtFactBegin
End Sub

Private Sub txtFactBegin_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtFactEnd_GotFocus()
    zlControl.TxtSelAll txtFactEnd
End Sub

Private Sub txtFactEnd_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtFactBegin_Change()
    txtFactEnd.Enabled = Not (Trim(txtFactBegin.Text) = "")
    If Trim(txtFactBegin.Text = "") Then txtFactEnd.Text = ""
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
    If txtNOBegin.Text <> "" Then txtNOBegin.Text = GetFullNO(txtNOBegin.Text, 11)
End Sub

Private Sub txtNOEnd_LostFocus()
    If txtNoEnd.Text <> "" Then txtNoEnd.Text = GetFullNO(txtNoEnd.Text, 11)
End Sub

Private Sub txtNoEnd_GotFocus()
    zlControl.TxtSelAll txtNoEnd
End Sub

Private Sub txtNoEnd_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
   '46516
   zlControl.TxtCheckKeyPress txtNOBegin, KeyAscii, m文本式
End Sub

Private Sub MakeFilter()
    Dim strSQL As String
    Dim strSQLtmp As String
    
    'by lesfeng 2010-03-06 性能优化
    Set mcllFilter = New Collection
    mcllFilter.Add Array("1901-01-01 00:00:00", "1901-01-01 00:00:00"), "收款时间"
    mcllFilter.Add Array("", ""), "单据号"
    mcllFilter.Add Array("", ""), "票据号"
    mcllFilter.Add "", "门诊号"
    mcllFilter.Add "", "住院号"
    mcllFilter.Add "", "姓名"
    mcllFilter.Add "", "结算号码"
    mcllFilter.Add "", "收款人"
    
    mstrFilter = ""
    mstrFilter = mstrFilter & " And (收款时间  Between [1] And [2]) "
    mcllFilter.Remove "收款时间"
    mcllFilter.Add Array(Format(dtpBegin.Value, "yyyy-MM-dd HH:mm:ss"), Format(dtpEnd.Value, "yyyy-MM-dd HH:mm:ss")), "收款时间"
      
    mblnDateMoved = zlDatabase.DateMoved(Format(IIf(dtpBegin.Value < dtpEnd.Value, dtpBegin.Value, dtpEnd.Value), dtpBegin.CustomFormat), , , Me.Caption)
    
    If txtNOBegin.Text <> "" And txtNoEnd.Text <> "" Then
        mstrFilter = mstrFilter & " And NO Between [3] And [4] "
    ElseIf txtNOBegin.Text <> "" Then
        mstrFilter = mstrFilter & " And NO=[3] "
    End If
    
    mcllFilter.Remove "单据号"
    mcllFilter.Add Array(Trim(txtNOBegin.Text), Trim(txtNoEnd.Text)), "单据号"
    
    If (txtFactBegin.Text <> "" And txtFactEnd.Text <> "") Or (txtFactBegin.Text <> "" And txtFactEnd.Text = "") Then
        '无需根据票据号判断,直接根据单据的登记时间判断
        strSQLtmp = IIf(txtFactEnd.Text = "", " =[5]", " Between [5] And [6]")
        If mblnDateMoved Then
            strSQL = "(Select A.NO" & _
            " From 票据打印内容 A,票据使用明细 B" & _
            " Where A.数据性质=2 And A.ID=B.打印ID And B.票种=2 And B.性质=1" & _
            " And B.号码 " & strSQLtmp & ") Union All" & _
            " (Select A.NO " & _
            " From H票据打印内容 A,H票据使用明细 B" & _
            " Where A.数据性质=2 And A.ID=B.打印ID And B.票种=2 And B.性质=1" & _
            " And B.号码 " & strSQLtmp & ")"
        Else
            strSQL = "Select A.NO" & _
            " From 票据打印内容 A,票据使用明细 B" & _
            " Where A.数据性质=2 And A.ID=B.打印ID And B.票种=2 And B.性质=1" & _
            " And B.号码 " & strSQLtmp
        End If
    End If
    
    If strSQL <> "" Then mstrFilter = mstrFilter & " And NO IN(" & strSQL & ")"
    
    mcllFilter.Remove "票据号"
    mcllFilter.Add Array(Trim(txtFactBegin.Text), Trim(txtFactEnd.Text)), "票据号"
    
    If chkCancel.Value = 1 And chk收款.Value = 0 Then
        mstrFilter = mstrFilter & " And (记录状态=2 or 记录状态=3)"
    ElseIf (chk收款.Value = 1 And chkCancel.Value = 0) Or (chk收款.Value = 0 And chkCancel.Value = 0) Then
        mstrFilter = mstrFilter & " And 记录状态=1"
    End If
    
    If txt住院号.Text <> "" Then
        mstrFilter = mstrFilter & " And B.病人ID = (Select Nvl(Max(病人ID),0) as 病人ID From 病案主页 Where 住院号=[7])"
    End If
    
    If txtClinicNO.Text <> "" Then
        mstrFilter = mstrFilter & " And B.门诊号=[11]"
    End If
    
    If txt姓名.Text <> "" Then
        mstrFilter = mstrFilter & " And Upper(B.姓名) Like [8]"
    End If
    
    If txt结算号码.Text <> "" Then mstrFilter = mstrFilter & " And 结算号码=[9]"
    
    If cbo操作员.ListIndex <> 0 Then mstrFilter = mstrFilter & " And 操作员姓名=[10]"
    
    mcllFilter.Remove "门诊号"
    mcllFilter.Add Trim(txtClinicNO.Text), "门诊号"
    mcllFilter.Remove "住院号"
    mcllFilter.Add Trim(txt住院号.Text), "住院号"
    mcllFilter.Remove "姓名"
    mcllFilter.Add "%" & Trim(UCase(txt姓名.Text)) & "%", "姓名"
    mcllFilter.Remove "结算号码"
    mcllFilter.Add Trim(txt结算号码.Text), "结算号码"
    mcllFilter.Remove "收款人"
    mcllFilter.Add Trim(zlCommFun.GetNeedName(cbo操作员.Text)), "收款人"
    
End Sub

Private Sub txt结算号码_GotFocus()
    zlControl.TxtSelAll txt结算号码
End Sub

Private Sub txt结算号码_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckInputLen txt结算号码, KeyAscii
End Sub

Private Sub txt姓名_GotFocus()
    zlControl.TxtSelAll txt姓名
End Sub

Private Sub txt住院号_GotFocus()
    zlControl.TxtSelAll txt住院号
End Sub

Private Sub txt住院号_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

