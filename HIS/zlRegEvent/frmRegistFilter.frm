VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#9.0#0"; "zlIDKind.ocx"
Begin VB.Form frmRegistFilter 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "过滤设置"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7155
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   7155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdDef 
      Caption         =   "缺省(&D)"
      Height          =   350
      Left            =   5880
      TabIndex        =   20
      Top             =   1560
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   4005
      Left            =   120
      TabIndex        =   21
      Top             =   0
      Width           =   5640
      Begin VB.CheckBox chkFilter 
         Caption         =   "预约接收单按预约时间显示"
         Height          =   210
         Left            =   690
         TabIndex        =   17
         Top             =   3645
         Width           =   2580
      End
      Begin VB.OptionButton optRegistRecord 
         Caption         =   "挂号及退号记录"
         Height          =   315
         Index           =   2
         Left            =   3480
         TabIndex        =   16
         Top             =   3240
         Width           =   1575
      End
      Begin VB.OptionButton optRegistRecord 
         Caption         =   "退号记录"
         Height          =   315
         Index           =   1
         Left            =   2115
         TabIndex        =   15
         Top             =   3240
         Width           =   1305
      End
      Begin VB.OptionButton optRegistRecord 
         Caption         =   "挂号记录"
         Height          =   315
         Index           =   0
         Left            =   690
         TabIndex        =   14
         Top             =   3240
         Value           =   -1  'True
         Width           =   1365
      End
      Begin zlIDKind.IDKindNew IDKind 
         Height          =   300
         Left            =   960
         TabIndex        =   12
         Top             =   2880
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   529
         Appearance      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   12
         FontName        =   "宋体"
         IDKind          =   -1
         AllowAutoICCard =   -1  'True
         AllowAutoIDCard =   -1  'True
         BackColor       =   -2147483633
      End
      Begin VB.ComboBox cbo费别 
         Height          =   300
         Left            =   1005
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1920
         Width           =   1830
      End
      Begin VB.ComboBox cbo号类 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3585
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1920
         Width           =   1830
      End
      Begin VB.TextBox txt医生 
         Height          =   300
         IMEMode         =   1  'ON
         Left            =   3585
         MaxLength       =   15
         TabIndex        =   7
         Top             =   1500
         Width           =   1830
      End
      Begin VB.ComboBox cbo科室 
         Height          =   300
         Left            =   1005
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1500
         Width           =   1830
      End
      Begin VB.ComboBox cbo操作员 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1005
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2400
         Width           =   1830
      End
      Begin VB.TextBox txtNOBegin 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1005
         MaxLength       =   8
         TabIndex        =   2
         Top             =   675
         Width           =   1815
      End
      Begin VB.TextBox txtNOEnd 
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3585
         MaxLength       =   8
         TabIndex        =   3
         Top             =   675
         Width           =   1830
      End
      Begin VB.TextBox txtFactBegin 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1005
         TabIndex        =   4
         Top             =   1095
         Width           =   1815
      End
      Begin VB.TextBox txtFactEnd 
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3585
         TabIndex        =   5
         Top             =   1095
         Width           =   1830
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   3360
         TabIndex        =   1
         Top             =   270
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   136970243
         CurrentDate     =   36588
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Left            =   1005
         TabIndex        =   0
         Top             =   270
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   136970243
         CurrentDate     =   36588
      End
      Begin VB.TextBox txtPatient 
         Height          =   300
         IMEMode         =   1  'ON
         Left            =   1560
         MaxLength       =   100
         TabIndex        =   13
         Top             =   2880
         Width           =   3855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "号类"
         Height          =   180
         Left            =   3000
         TabIndex        =   32
         Top             =   1980
         Width           =   360
      End
      Begin VB.Label lbl费别 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "费别"
         Height          =   180
         Left            =   585
         TabIndex        =   31
         Top             =   1980
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "至"
         Height          =   180
         Left            =   3120
         TabIndex        =   30
         Top             =   1155
         Width           =   180
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "至"
         Height          =   180
         Left            =   3120
         TabIndex        =   29
         Top             =   735
         Width           =   180
      End
      Begin VB.Label lblData_ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "至"
         Height          =   180
         Left            =   3120
         TabIndex        =   28
         Top             =   330
         Width           =   180
      End
      Begin VB.Label lbl医生 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "医生"
         Height          =   180
         Left            =   3030
         TabIndex        =   27
         Top             =   1560
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "身份类别"
         Height          =   180
         Left            =   225
         TabIndex        =   11
         Top             =   2925
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单据号"
         Height          =   180
         Left            =   405
         TabIndex        =   26
         Top             =   735
         Width           =   540
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "挂号时间"
         Height          =   180
         Left            =   225
         TabIndex        =   25
         Top             =   330
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "科室"
         Height          =   180
         Left            =   585
         TabIndex        =   24
         Top             =   1560
         Width           =   360
      End
      Begin VB.Label lbl操作员 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "挂号员"
         Height          =   180
         Left            =   405
         TabIndex        =   23
         Top             =   2460
         Width           =   540
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "票据号"
         Height          =   180
         Left            =   405
         TabIndex        =   22
         Top             =   1155
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5880
      TabIndex        =   19
      Top             =   600
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5880
      TabIndex        =   18
      Top             =   120
      Width           =   1100
   End
End
Attribute VB_Name = "frmRegistFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit    '要求变量声明
Private mbytType As Byte     '0-挂号清单条件,1-预约清单条件,2-接收清单条件
Public mlngModule As Long
Public mstrFilter As String
Public mstrSectName As String   '用来指定当前默认的科室
Public mblnDateMoved As Boolean    'Out
Private mstrCardStr As String    '用来保存启用的卡
Private Const mstrIDKind = "1-姓名;2-就诊卡;3-门诊号;4-医保号;5-身份证号;6-IC卡号"

Private mblnNotClick As Boolean
Private mblnUnChange As Boolean
Private mrsInfo As ADODB.Recordset
Private mbln允许住院病人挂号 As Boolean
Private mblnOlnyBJYB As Boolean
Public mlngPrePatient As Long
Private mblnKeyReturn As Boolean
Private mblnValid As Boolean
Private Sub cbo操作员_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo操作员.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo操作员.Hwnd, KeyAscii)
    If lngIdx <> -2 Then cbo操作员.ListIndex = lngIdx
    If cbo操作员.ListIndex = -1 And cbo操作员.ListCount <> 0 Then cbo操作员.ListIndex = 0
End Sub

Private Sub cbo科室_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo科室.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo科室.Hwnd, KeyAscii)
    If lngIdx <> -2 Then cbo科室.ListIndex = lngIdx
    If cbo科室.ListIndex = -1 And cbo科室.ListCount <> 0 Then cbo科室.ListIndex = 0
End Sub

Private Sub cmdCancel_Click()
    gblnOk = False
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
    If txtFactBegin.Text <> "" And txtFactEnd.Text <> "" Then
        If txtFactEnd.Text < txtFactBegin.Text Then
            MsgBox "结束票据号不能小于开始票据号！", vbInformation, gstrSysName
            txtFactEnd.SetFocus: Exit Sub
        End If
    End If
    '74237:预约时间的查询范围提示
    If mbytType = 1 Then
        If dtpEnd.Value - dtpBegin.Value > gint预约天数 + 1 Then
            If MsgBox("当前预约时间范围过大(超过" & gint预约天数 & "天),可能会导致读取和加载时间过长,你是否还需要继续查询?", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
    End If
    Call MakeFilter

    gblnOk = True
    Hide
End Sub

Private Sub Form_Activate()
    txtNOBegin.SetFocus
    Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And Not ActiveControl Is txtPatient Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr(1, "'[]", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii = 13 And Not ActiveControl Is txtPatient Then KeyAscii = 0
    If InStr("':：;；?？", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0: Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Dim Curdate As Date, i As Integer
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strTmp As String
    On Error GoTo errH

    gblnOk = False

    txtNOBegin.Text = ""
    txtNoEnd.Text = ""
    txtFactBegin.Text = ""
    txtFactEnd.Text = ""
    optRegistRecord(0).Value = True
    txtPatient.Text = ""
    txt医生.Text = ""
    '47928
    InitIDKind
    'dtpBegin.Enabled = mbytType <> 2:问题44946
    'dtpEnd.Enabled = mbytType <> 2:问题44946
    txtFactBegin.Enabled = mbytType = 0
    txtFactEnd.Enabled = mbytType = 0
    txtFactBegin.BackColor = IIf(mbytType = 0, txtPatient.BackColor, Me.BackColor)
    txtFactEnd.BackColor = IIf(mbytType = 0, txtPatient.BackColor, Me.BackColor)
    optRegistRecord(0).Enabled = mbytType = 0 Or mbytType = 1
    optRegistRecord(1).Enabled = mbytType = 0 Or mbytType = 1
    optRegistRecord(2).Enabled = mbytType = 0 Or mbytType = 1
    dtpEnd.MinDate = CDate("1905-01-01")
    dtpBegin.MinDate = CDate("1905-01-01")
    Curdate = zlDatabase.Currentdate
    If mbytType = 0 Then    '挂号
        lblDate.Caption = "挂号时间"
        '缺省时间为当天内
        dtpBegin.Value = Format(Curdate, "yyyy-MM-dd 00:00:00")
        dtpEnd.Value = Format(Curdate, "yyyy-MM-dd 23:59:59")
    ElseIf mbytType = 1 Then    '预约
        '缺省为预约时间未失效的单据
        lblDate.Caption = "预约时间"
        dtpBegin.Value = Format(Curdate, "yyyy-MM-dd 00:00:00")
        dtpEnd.Value = Format(Curdate + gint预约天数, "yyyy-MM-dd 23:59:59")
    ElseIf mbytType = 2 Then    '接收
        '不需要设置时间
        lblDate.Caption = "预约时间"
        dtpBegin.Value = Format(Curdate, "yyyy-MM-dd 00:00:00")
        dtpEnd.Value = Format(Curdate, "yyyy-MM-dd 23:59:59")
        dtpEnd.MinDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
        dtpBegin.MinDate = dtpEnd.MinDate
    End If

    '挂号员
    cbo操作员.Clear
    cbo操作员.AddItem "所有挂号员"
    cbo操作员.ListIndex = 0

    Set rsTmp = GetPersonnel("门诊挂号员", True)
    If rsTmp.RecordCount > 0 Then
        For i = 1 To rsTmp.RecordCount
            cbo操作员.AddItem rsTmp!简码 & "-" & rsTmp!姓名
            If rsTmp!id = UserInfo.id Then cbo操作员.ListIndex = cbo操作员.NewIndex
            rsTmp.MoveNext
        Next
    End If

    '挂号科室
    Set rsTmp = GetDepartments("'临床'", "1,3")
    cbo科室.Clear
    cbo科室.AddItem "所有科室"
    cbo科室.ListIndex = 0

    Do While Not rsTmp.EOF
        cbo科室.AddItem rsTmp!编码 & "-" & rsTmp!名称
        cbo科室.ItemData(cbo科室.NewIndex) = rsTmp!id
        If mstrSectName = rsTmp!名称 Then cbo科室.ListIndex = cbo科室.NewIndex
        rsTmp.MoveNext
    Loop

    '费别
    cbo费别.Clear
    cbo费别.AddItem "所有费别"
    cbo费别.ListIndex = 0
    strSQL = "Select 编码,名称,简码,Nvl(缺省标志,0) as 缺省 From 费别 Where Nvl(服务对象,3) IN(1,3) Order by 编码"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cbo费别.AddItem rsTmp!编码 & "-" & rsTmp!名称
            rsTmp.MoveNext
        Next
    End If

    '号类
    cbo号类.Clear
    cbo号类.AddItem "所有号类"
    cbo号类.ListIndex = 0
    strSQL = "Select 编码,名称,简码,缺省标志,说明 From 号类 Order by 编码"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    For i = 1 To rsTmp.RecordCount
        cbo号类.AddItem rsTmp!名称
        rsTmp.MoveNext
    Next
    mbln允许住院病人挂号 = zlDatabase.GetPara("允许住院病人挂号", glngSys, mlngModule, 0) = "1"
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub optRegistRecord_Click(index As Integer)
    If optRegistRecord(1).Value = True Then
        lbl操作员.Caption = "退号员"
        lblDate.Caption = "退号时间"
    Else
        lbl操作员.Caption = "挂号员"
        lblDate.Caption = "挂号时间"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mbytType = 0
    Set mrsInfo = Nothing
    mlngPrePatient = 0
    IDKind.SetAutoReadCard False
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

Private Sub txtPatient_Change()
    txtPatient.Tag = "": mlngPrePatient = 0
    If Me.ActiveControl Is txtPatient Then
        'If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "")
       ' If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
        IDKind.SetAutoReadCard txtPatient.Text = ""
    End If
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim lngID As Long, lngUnit As Long, i As Long
    Dim rsTmp As ADODB.Recordset, strInfo As String
    Dim strSQL As String, curTotal As Currency
    Dim blnCard As Boolean, blnICCard As Boolean

    On Error GoTo errH
    If txtPatient.Locked Then Exit Sub
    mblnKeyReturn = KeyAscii = 13
    If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub

    If IDKind.GetCurCard.名称 Like "姓名*" Then
        blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, IDKind.ShowPassText)
    ElseIf IDKind.IDKind = IDKind.GetKindIndex("门诊号") Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            If Not (IsNumeric(Chr(KeyAscii)) Or Chr(KeyAscii) = "-") Then KeyAscii = 0: Exit Sub
        End If
        txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
    End If
    If blnCard And Len(txtPatient.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Trim(txtPatient.Text) <> "" Then
        If KeyAscii <> 13 Then
            txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
            txtPatient.SelStart = Len(txtPatient.Text)
        End If
        KeyAscii = 0
        Call FindPati(IDKind.GetCurCard, blnCard, txtPatient.Text)
    End If
    If Me.ActiveControl Is txtPatient And mblnKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog    '
End Sub

Private Sub FindPati(ByVal objCard As Card, ByVal blnCard As Boolean, ByVal strInput As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:查找病人
    '编制:刘兴洪
    '日期:2012-08-29 17:53:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnICCard As Boolean, blnIDCard As Boolean
   '读取病人信息
    Call GetPatient(objCard, txtPatient.Text, blnCard)
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
    '46512
    zlControl.TxtCheckKeyPress txtNOBegin, KeyAscii, m文本式
End Sub

Private Sub txtNOBegin_LostFocus()
    If txtNOBegin.Text <> "" Then txtNOBegin.Text = GetFullNO(txtNOBegin.Text, 12)
End Sub


Private Sub txtNOEnd_LostFocus()
    If txtNoEnd.Text <> "" Then txtNoEnd.Text = GetFullNO(txtNoEnd.Text, 12)
End Sub

Private Sub txtNoEnd_GotFocus()
    zlControl.TxtSelAll txtNoEnd
End Sub


Private Sub txtNoEnd_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    '46512
    zlControl.TxtCheckKeyPress txtNoEnd, KeyAscii, m文本式
End Sub

Private Sub MakeFilter()
    Dim strSQL As String
    Dim strSQLtmp As String

    mstrFilter = " And 1=1"

    If mbytType = 0 Then    '挂号
        If chkFilter.Value = 0 Then
            mstrFilter = " And A.登记时间 Between [1] And [2]"
        Else
            mstrFilter = " And A.发生时间 Between [1] And [2]"
        End If
    ElseIf mbytType = 1 Then    '预约
        mstrFilter = " And A.发生时间 Between [1] And [2]"
    ElseIf mbytType = 2 Then    '接收
        '不需要另设时间条件
    End If

    mblnDateMoved = zlDatabase.DateMoved(Format(IIf(dtpBegin.Value < dtpEnd.Value, dtpBegin.Value, dtpEnd.Value), dtpBegin.CustomFormat), , , Me.Caption)

    If txtNOBegin.Text <> "" And txtNoEnd.Text <> "" Then
        mstrFilter = mstrFilter & " And A.NO Between [3] And [4]"
    ElseIf txtNOBegin.Text <> "" Then
        mstrFilter = mstrFilter & " And A.NO=[3]"
    ElseIf txtNoEnd.Text <> "" Then
        mstrFilter = mstrFilter & " And A.NO=[4]"
    End If

    If cbo操作员.ListIndex > 0 Then mstrFilter = mstrFilter & " And A.操作员姓名||''=[5]"

    If txtPatient.Text <> "" And mlngPrePatient <> 0 And Not mrsInfo Is Nothing Then
        If Val(Nvl(mrsInfo!id)) = mlngPrePatient Then
            mstrFilter = mstrFilter & " And D.病人ID=[6]"
        End If
    ElseIf txtPatient.Text <> "" And mrsInfo Is Nothing Then
        If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Left(txtPatient.Text, 1))) > 0 Then
            mstrFilter = mstrFilter & " And Upper(A.姓名) Like [13]"
            txtPatient.Text = UCase(txtPatient.Text)
        Else
            mstrFilter = mstrFilter & " And A.姓名 Like [13]"
        End If
    End If


    If txt医生.Text <> "" Then
        mstrFilter = mstrFilter & " And A.执行人 Like [7]"
    End If

    If (txtFactBegin.Text <> "" And txtFactEnd.Text <> "") Or (txtFactBegin.Text <> "" And txtFactEnd.Text = "") Then
        '无需根据票据号判断,直接根据单据的登记时间判断
        strSQLtmp = IIf(txtFactEnd.Text = "", " =[8] ", " Between [8] And [9] ")

        If mblnDateMoved Then
            strSQL = "(Select A.NO" & _
                   " From 票据打印内容 A,票据使用明细 B" & _
                   " Where A.数据性质=4 And A.ID=B.打印ID And B.性质=1" & _
                   " And B.号码 " & strSQLtmp & ") Union All" & _
                   " (Select A.NO " & _
                   " From H票据打印内容 A,H票据使用明细 B" & _
                   " Where A.数据性质=4 And A.ID=B.打印ID And B.性质=1" & _
                   " And B.号码 " & strSQLtmp & ")"
        Else
            strSQL = "Select A.NO" & _
                   " From 票据打印内容 A,票据使用明细 B" & _
                   " Where A.数据性质=4 And A.ID=B.打印ID And B.性质=1" & _
                   " And B.号码 " & strSQLtmp
        End If
    End If

    If strSQL <> "" Then mstrFilter = mstrFilter & " And A.NO IN(" & strSQL & ")"

    '挂号科室(执行科室)
    If cbo科室.ListIndex > 0 Then
        mstrFilter = mstrFilter & " And A.执行部门ID+0=[10]"
    End If

    If cbo费别.ListIndex > 0 Then
        mstrFilter = mstrFilter & " And (F.费别 = [11] or F.费别 is Null)"
    End If

    If cbo号类.ListIndex > 0 Then
        mstrFilter = mstrFilter & " And B.号类 = [12]"
    End If

End Sub

Private Sub txtPatient_GotFocus()
    Call zlControl.TxtSelAll(txtPatient)
    Call zlCommFun.OpenIme(True)
    If txtPatient.Text = "" And ActiveControl Is txtPatient Then
'        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "")
'        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
        IDKind.SetAutoReadCard txtPatient.Text = ""
    End If
End Sub

Private Sub txtPatient_LostFocus()
    Call zlCommFun.OpenIme
    IDKind.SetAutoReadCard False
End Sub

 

Private Sub txt医生_GotFocus()
    Call zlControl.TxtSelAll(txt医生)
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt医生_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txt医生_Validate(Cancel As Boolean)
    Dim strDoctor As String
    strDoctor = UCase(Trim(txt医生.Text))
    If strDoctor <> "" Then
        If zlCommFun.IsNumOrChar(strDoctor) Then
            strDoctor = GetDoctorName(strDoctor)
            If strDoctor = "" Then Cancel = True
        End If
    End If
    txt医生.Text = strDoctor
End Sub

Private Function GetDoctorName(ByVal strCode As String) As String
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strIF As String, lngDept As Long, blnCancel As Boolean, vRect As RECT

    On Error GoTo Hd
    If zlCommFun.IsCharAlpha(strCode) Then
        strIF = " And A.简码 Like [1]"
        strCode = strCode & "%"
    Else
        strIF = " And (A.简码 = [1] Or A.编号 = [1])"
    End If
    If cbo科室.ListIndex > 0 Then
        strIF = strIF & " And B.部门ID = [2]"
        lngDept = cbo科室.ItemData(cbo科室.ListIndex)
    End If
    strSQL = "Select Distinct A.Id,A.姓名 From 人员表 A, 部门人员 B,人员性质说明 C" & vbCrLf & _
             "Where A.id=B.人员id And A.id=C.人员id  And C.人员性质='医生'" & strIF

    vRect = zlControl.GetControlRect(txt医生.Hwnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "选择医生", 1, "", "请选择医生", False, False, True, vRect.Left, vRect.Top, txt医生.Height, blnCancel, False, True, strCode, lngDept)
    If Not rsTmp Is Nothing Then
        GetDoctorName = rsTmp!姓名
    End If
    Exit Function
Hd:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Function

'初始化IDKIND
Private Function InitIDKind() As Boolean
    Dim objCard As Card, rsTmp As ADODB.Recordset
    Dim lngCardID As Long, strSQL As String
    Call IDKind.zlInit(Me, glngSys, mlngModule, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "", txtPatient)
    lngCardID = Val(zlDatabase.GetPara("缺省医疗卡类别", glngSys, mlngModule, 0))
    '72936:刘尔旋,2014-05-13,缺省发卡类型被停用后报错的问题
    If lngCardID <> 0 Then
        strSQL = "Select 1 From 医疗卡类别 Where ID=[1] And Nvl(是否启用,0)=1"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngCardID)
        If Not rsTmp.EOF Then IDKind.DefaultCardType = lngCardID
    End If
    Set objCard = IDKind.GetfaultCard
    If IDKind.Cards.按缺省卡查找 And Not objCard Is Nothing Then
        gobjSquare.bln缺省卡号密文 = objCard.卡号密文规则 <> ""
        gobjSquare.int缺省卡号长度 = objCard.卡号长度
        Set gobjSquare.objDefaultCard = objCard

    Else
        gobjSquare.bln缺省卡号密文 = IDKind.Cards.加密显示
        gobjSquare.int缺省卡号长度 = 100
    End If
End Function
'获取默认IDKind索引
Private Function IDKindDefaultKind() As Long
    Dim lngIndex As Long
    'IDkind的默认Kind
    If IDKind.DefaultCardType = "" Then
        lngIndex = -1
    Else
        If IsNumeric(IDKind.DefaultCardType) Then
            lngIndex = IDKind.GetKindIndex(IDKind.GetfaultCard.名称)
        Else
            lngIndex = IDKind.GetKindIndex(IDKind.DefaultCardType)
        End If
    End If
    IDKindDefaultKind = lngIndex
End Function


'控件名称是否匹配
Private Function IsCardType(ByVal IDKindCtl As IDKindNew, ByVal strCardName As String) As Boolean
    If IDKindCtl Is Nothing Then Exit Function
    If UCase(TypeName(IDKindCtl)) <> "IDKINDNEW" Then Exit Function
    Select Case strCardName
    Case "姓名", "姓名或就诊卡"
        IsCardType = IDKindCtl.GetCurCard.名称 Like "姓名*"
    Case "身份证", "身份证号", "二代身份证"
        IsCardType = IDKindCtl.GetCurCard.名称 Like "*身份证*"
    Case "IC卡号", "IC卡"
        IsCardType = IDKindCtl.GetCurCard.名称 Like "IC卡*"
    Case "医保号"
        IsCardType = IDKindCtl.GetCurCard.名称 = "医保号"
    Case "门诊号"
        IsCardType = IDKindCtl.GetCurCard.名称 = "门诊号"
    Case Else
        If IDKindCtl.GetCurCard Is Nothing Then Exit Function
        If Not IsNumeric(strCardName) Or Val(strCardName) <= 0 Then Exit Function
        If IDKindCtl.GetCurCard.接口序号 <= 0 Then Exit Function
        IsCardType = IDKindCtl.GetCurCard.接口序号 = Val(strCardName)
    End Select
End Function

Private Sub IDKind_ItemClick(index As Integer, objCard As zlIDKind.Card)
    Set gobjSquare.objCurCard = objCard
    If txtPatient.Text <> "" Then txtPatient.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    zlControl.TxtSelAll txtPatient
End Sub

Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng卡类别ID As Long, strOutCardNO As String, strExpand
    Dim strOutPatiInforXML As String
    
    If objCard.名称 Like "IC卡*" And objCard.系统 Then
'        If mobjICCard Is Nothing Then
'            Set mobjICCard = CreateObject("zlICCard.clsICCard")
'            Set mobjICCard.gcnOracle = gcnOracle
'        End If
'        If mobjICCard Is Nothing Then Exit Sub
'        txtPatient.Text = mobjICCard.Read_Card()
'        If txtPatient.Text <> "" Then
'            Call FindPati(objCard, True, txtPatient.Text)
'        End If
        Exit Sub
    End If
    
   lng卡类别ID = objCard.接口序号
    If lng卡类别ID <= 0 Then Exit Sub
    '    zlReadCard(frmMain As Object, _
    '    ByVal lngModule As Long, _
    '    ByVal lngCardTypeID As Long, _
    '    ByVal blnOlnyCardNO As Boolean, _
    '    ByVal strExpand As String, _
    '    ByRef strOutCardNO As String, _
    '    ByRef strOutPatiInforXML As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '功能:读卡接口
    '    '入参:frmMain-调用的父窗口
    '    '       lngModule-调用的模块号
    '    '       strExpand-扩展参数,暂无用
    '    '       blnOlnyCardNO-仅仅读取卡号
    '    '出参:strOutCardNO-返回的卡号
    '    '       strOutPatiInforXML-(病人信息返回.XML串)
    '    '返回:函数返回    True:调用成功,False:调用失败\
    If gobjSquare.objSquareCard.zlReadCard(Me, mlngModule, lng卡类别ID, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
    txtPatient.Text = strOutCardNO
    If txtPatient.Text <> "" Then
        Call FindPati(objCard, True, txtPatient.Text)
    End If
End Sub
Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    If txtPatient.Locked Then Exit Sub
    txtPatient.Text = objPatiInfor.卡号
    Call FindPati(objCard, True, txtPatient.Text)
End Sub
 


Private Sub GetPatient(ByVal objCard As Card, ByVal strInput As String, Optional blnCard As Boolean)
    '------------------------------------------------------------------------------------------------------------------------
    '功能：获取病人信息
    '入参：blnCard=是否就诊卡刷卡
    '编制：刘兴洪
    '日期：2010-07-16 14:24:14
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strTemp As String
    Dim blnSame As Boolean, blnCancel As Boolean
    Dim cur余额 As Currency, curMoney As Currency
    Dim i As Integer, strPati As String
    Dim vRect As RECT, str非在院 As String
    Dim strSQL As String, lng卡类别ID As Long, strPassWord As String, strErrMsg As String
    Dim strTmp As String
    Dim lng病人ID As Long, blnHavePassWord As Boolean

    On Error GoTo errH
    If Not mbln允许住院病人挂号 Then
        str非在院 = " And Not Exists(Select 1 From 病案主页 Where 病人ID=B.病人ID And 主页ID=B.主页ID And Nvl(病人性质,0)=0 And 出院日期 is Null)"
    End If

    strSQL = ""
    If blnCard = True And objCard.名称 Like "姓名*" Then    '刷卡
        If IDKind.Cards.按缺省卡查找 And Not IDKind.GetfaultCard Is Nothing Then
            lng卡类别ID = IDKind.GetfaultCard.接口序号
        Else
            lng卡类别ID = "-1"
        End If
        '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|是否存在帐户(1-存在帐户;0-不存在帐户);…
        If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
        If lng病人ID <= 0 Then GoTo NotFoundPati:
        strInput = "-" & lng病人ID
        blnHavePassWord = True
        strSQL = strSQL & " And B.病人ID=[2] " & str非在院
        
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then
        '门诊号
        strSQL = strSQL & " And B.门诊号=[2]" & str非在院
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then
        '病人ID
        strSQL = strSQL & " And B.病人ID=[2]" & str非在院
    Else
        Select Case objCard.名称
        Case "姓名", "姓名或就诊卡"
            txtPatient.Tag = strInput
            Set mrsInfo = Nothing: Exit Sub
            zlCommFun.PressKey vbKeyTab
        Case "医保号"
            strInput = UCase(strInput)
            If mblnOlnyBJYB And zlCommFun.ActualLen(strInput) >= 9 Then
                '仅北京医保才有效:见问题:问题:26982
                strSQL = strSQL & " And B.医保号 like [3] " & str非在院
                strTemp = Left(strInput, 9) & "%"
            Else
                strSQL = strSQL & " And B.医保号=[1]" & str非在院
            End If
        Case "身份证号", "身份证", "二代身份证"
            strInput = UCase(strInput)
            If gobjSquare.objSquareCard.zlGetPatiID("身份证", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
            strSQL = strSQL & " And B.病人ID=[2]" & str非在院
            strInput = "-" & lng病人ID
        Case "IC卡号", "IC卡"
            strInput = UCase(strInput)
            If gobjSquare.objSquareCard.zlGetPatiID("IC卡", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
            strSQL = strSQL & " And B.病人ID=[2]" & str非在院
            strInput = "-" & lng病人ID
        Case "门诊号"
            If Not IsNumeric(strInput) Then strInput = "0"
            strSQL = strSQL & " And B.门诊号=[1]" & str非在院
        Case Else
            '其他类别的,获取相关的病人ID
            If Val(objCard.接口序号) > 0 Then
                lng卡类别ID = Val(objCard.接口序号)
                If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
                If lng病人ID = 0 Then lng病人ID = 0
            Else
                If gobjSquare.objSquareCard.zlGetPatiID(objCard.名称, strInput, False, lng病人ID, _
                                                        strPassWord, strErrMsg) = False Then lng病人ID = 0
            End If
            If lng病人ID <= 0 Then lng病人ID = 0
            strSQL = strSQL & " And B.病人ID=[2]" & str非在院
            strInput = "-" & lng病人ID
            blnHavePassWord = True
        End Select
    End If
    strSQL = "" & _
    "   Select distinct  B.病人id As ID, Decode(sign(nvl(X.病人id,0)),0,'','√') as 三方账户,  " & _
    "           B.病人id,B.姓名, B.性别, B.年龄, B.门诊号, B.出生日期, B.身份证号, B.家庭地址, B.工作单位," & _
    "            A.名称 险类名称" & _
    "   From 病人信息 B, 保险类别 A,医疗卡类别 Y,病人医疗卡信息 X" & _
    "   Where B.险类 = A.序号(+) and b.病人id=X.病人id(+)  " & _
    "               And X.状态(+)=0 and  X.卡类别id=Y.id(+)  and Y.是否自制(+)=0 And B.停用时间 Is Null   " & _
                    strSQL
    On Error GoTo errH
    vRect = zlControl.GetControlRect(txtPatient.Hwnd)
    Set mrsInfo = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "病人查找", 1, "√", "请选择病人", False, False, True, vRect.Left, vRect.Top, txtPatient.Height, blnCancel, False, True, strInput, CStr(Mid(strInput, 2)), strInput & "%", dtpBegin.Value, dtpEnd.Value)
    
    If blnCancel Or mrsInfo Is Nothing Then
        Set mrsInfo = Nothing: txtPatient.Text = "": Exit Sub
    End If
    
    If mrsInfo!id = 0 Then    '没有找到病人信息
        Set mrsInfo = Nothing: txtPatient.Text = "": Exit Sub
    End If
    
    txtPatient.MaxLength = zlGetPatiInforMaxLen.intPatiName
    txtPatient.Text = Nvl(mrsInfo!姓名)
    Me.txtPatient.Tag = Nvl(mrsInfo!id)
    mlngPrePatient = Val(Nvl(mrsInfo!id))
    zlCommFun.PressKey vbKeyTab
    Exit Sub
    
NotFoundPati:
    Set mrsInfo = Nothing: txtPatient.Text = "": Exit Sub
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



'过滤方式
Public Property Let bytType(ByVal vNewValue As Byte)
    mbytType = vNewValue
    chkFilter.Visible = mbytType = 0
End Property
