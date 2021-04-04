VERSION 5.00
Begin VB.Form frmLabSamplingRegister 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "登记"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8040
   Icon            =   "frmLabSamplingRegister.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   345
      Left            =   6540
      TabIndex        =   26
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "保存(&S)"
      Height          =   345
      Left            =   5100
      TabIndex        =   13
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CheckBox chkConcatenation 
      Caption         =   "保存当前项目连续输入"
      Height          =   225
      Left            =   30
      TabIndex        =   14
      Top             =   2220
      Width           =   2295
   End
   Begin VB.Frame FraPatientInfo 
      Caption         =   "病人信息"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2145
      Left            =   30
      TabIndex        =   15
      Top             =   30
      Width           =   7965
      Begin VB.TextBox txtUnit 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   900
         TabIndex        =   11
         Top             =   1350
         Width           =   4455
      End
      Begin VB.ComboBox cbo执行科室 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6240
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   960
         Width           =   1515
      End
      Begin VB.TextBox txt姓名 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   900
         MaxLength       =   20
         TabIndex        =   0
         ToolTipText     =   "数字为就诊卡号、“－”打头为病人ID、“＋”住院号、“*”门诊号、“.”挂号单号、“/”收费单据号"
         Top             =   210
         Width           =   1635
      End
      Begin VB.ComboBox cbo性别 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmLabSamplingRegister.frx":6852
         Left            =   3210
         List            =   "frmLabSamplingRegister.frx":6854
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   210
         Width           =   675
      End
      Begin VB.TextBox txt年龄 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4320
         MaxLength       =   5
         TabIndex        =   2
         Top             =   210
         Width           =   435
      End
      Begin VB.ComboBox cboAge 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmLabSamplingRegister.frx":6856
         Left            =   4770
         List            =   "frmLabSamplingRegister.frx":6869
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   210
         Width           =   750
      End
      Begin VB.TextBox txtBed 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6720
         TabIndex        =   5
         Top             =   210
         Width           =   1035
      End
      Begin VB.TextBox txtPatientDept 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3570
         TabIndex        =   7
         Top             =   600
         Width           =   4185
      End
      Begin VB.TextBox txtID 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   600
         Width           =   1635
      End
      Begin VB.TextBox txt医嘱内容 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   900
         MaxLength       =   1000
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   1710
         Width           =   6525
      End
      Begin VB.ComboBox cbo开单科室 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmLabSamplingRegister.frx":6885
         Left            =   900
         List            =   "frmLabSamplingRegister.frx":6887
         TabIndex        =   8
         Top             =   960
         Width           =   1635
      End
      Begin VB.ComboBox cbo医生 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3570
         TabIndex        =   9
         Top             =   960
         Width           =   1785
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "…"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7440
         TabIndex        =   16
         Top             =   1710
         Width           =   285
      End
      Begin VB.TextBox txt年龄1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5550
         MaxLength       =   5
         TabIndex        =   4
         Top             =   210
         Width           =   555
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单        位"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   10
         Left            =   150
         TabIndex        =   28
         Top             =   1380
         Width           =   720
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "执行科室"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   6
         Left            =   5460
         TabIndex        =   27
         Top             =   990
         Width           =   720
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   3915
         TabIndex        =   25
         Top             =   255
         Width           =   360
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   2790
         TabIndex        =   24
         Top             =   255
         Width           =   360
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "所在科室"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   2790
         TabIndex        =   23
         Top             =   645
         Width           =   720
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "床号"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   5
         Left            =   6330
         TabIndex        =   22
         Top             =   255
         Width           =   360
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "标  识 号"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   180
         TabIndex        =   21
         Top             =   645
         Width           =   675
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓       名"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   180
         TabIndex        =   20
         Top             =   255
         Width           =   675
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "申请项目"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   7
         Left            =   150
         TabIndex        =   19
         Top             =   1740
         Width           =   720
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "申请科室"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   8
         Left            =   150
         TabIndex        =   18
         Top             =   990
         Width           =   720
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "申请医生"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   9
         Left            =   2790
         TabIndex        =   17
         Top             =   990
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmLabSamplingRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rsRelativeAdvice As ADODB.Recordset                             '登记的相关医嘱
Private PatientType As Integer, mlng病人ID As Long, mstrNO As String    '门诊收费单据号
Private mlngCapID As Long                                               '采集项目ID
Private mlngReqDept As Long, mstrReqDoctor As String                    '默认的登记科室和医生
Private mlngKey As Long                                                 'ID
Private mblnSaveAdvice As Boolean                                       '是否需要保存医嘱，用于修改在院病人标本信息
Private mstrKeys As String                                              '当前核收的申请医嘱ID
Private mblnBarCode As Boolean                                          '条码
Private iInputType As Integer
Private mstrExtData  As String                                           '登记的申请项目信息
Private mbln微生物项目 As Boolean
Private mlngDeptID As Long                                              '科室ID

Private Sub cboAge_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cbo开单科室_Click()
    If cbo开单科室.ListIndex > -1 Then InitDoctors cbo开单科室.ItemData(cbo开单科室.ListIndex)
End Sub

Private Sub cbo开单科室_GotFocus()
    Call zlControl.TxtSelAll(cbo开单科室)
End Sub

Private Sub cbo开单科室_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cbo开单科室_Validate(Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, intIdx As Long, i As Long
    Dim strInput As String
    Dim vRect As RECT, blnCancel As Boolean
        
    If cbo开单科室.ListIndex <> -1 Then mlngReqDept = Me.cbo开单科室.ItemData(Me.cbo开单科室.ListIndex): Exit Sub '已选中
    If cbo开单科室.Text = "" Then '无输入
        Exit Sub
    End If
    
    strInput = UCase(NeedName(cbo开单科室.Text))
    '全院临床科室
    strSQL = _
        " Select Distinct A.ID,A.编码,A.名称,A.简码" & _
        " From 部门表 A,部门性质说明 B " & _
        " Where B.部门ID = A.ID " & _
        " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) " & _
        " And (B.工作性质 IN('临床','体检'))" & _
        " And (Upper(A.编码) Like [1] Or Upper(A.名称) Like [2] Or Upper(A.简码) Like [2])" & _
        " Order by A.编码"
    
    On Error GoTo errH
    vRect = GetControlRect(cbo医生.Hwnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "开嘱科室", False, "", "", False, False, _
        True, vRect.Left, vRect.Top, cbo开单科室.Height, blnCancel, False, True, strInput & "%", strInput & "%")
    If Not rsTmp Is Nothing Then
        If Not zlControl.CboLocate(cbo开单科室, rsTmp!名称) Then
            cbo开单科室.Text = ""
        End If
    Else
        If Not blnCancel Then
            MsgBox "未找到对应的科室。", vbInformation, gstrSysName
        End If
        Cancel = True: Exit Sub
    End If
    If Me.cbo开单科室.ListIndex > -1 Then mlngReqDept = Me.cbo开单科室.ItemData(Me.cbo开单科室.ListIndex)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbo性别_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
        Exit Sub
    End If
End Sub

Private Sub cbo医生_Click()
    Call zlControl.TxtSelAll(cbo医生)
End Sub

Private Sub cbo医生_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cbo医生_Validate(Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, intIdx As Long, i As Long
    Dim strInput As String
    Dim vRect As RECT, blnCancel As Boolean
        
    If cbo医生.ListIndex <> -1 Then mstrReqDoctor = Me.cbo医生.Text: Exit Sub '已选中
    If cbo医生.Text = "" Then '无输入
        Exit Sub
    End If
    
    strInput = UCase(NeedName(cbo医生.Text))
    '全院医生
    strSQL = "Select Distinct 部门ID From 部门性质说明 Where 服务对象 IN(1,2,3)"
    strSQL = "Select Distinct A.ID,A.编号,A.姓名,A.简码" & _
        " From 人员表 A,部门人员 B,人员性质说明 C" & _
        " Where A.ID=B.人员ID And A.ID=C.人员ID And C.人员性质='医生'" & _
        " And B.部门ID IN(" & strSQL & ")" & _
        " And (Upper(A.编号) Like [1] Or Upper(A.姓名) Like [2] Or Upper(A.简码) Like [2])" & _
        " And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) " & _
        " Order by A.简码"
    
    On Error GoTo errH
    vRect = GetControlRect(cbo医生.Hwnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "开嘱医生", False, "", "", False, False, _
        True, vRect.Left, vRect.Top, cbo医生.Height, blnCancel, False, True, strInput & "%", strInput & "%")
    If Not rsTmp Is Nothing Then
        cbo医生.Text = rsTmp!姓名
    Else
        If Not blnCancel Then
            MsgBox "未找到对应的医生。", vbInformation, gstrSysName
        End If
        Cancel = True: Exit Sub
    End If
    If Len(Trim(Me.cbo医生.Text)) > 0 Then mstrReqDoctor = Me.cbo医生.Text
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbo执行科室_Click()
    mlngDeptID = cbo执行科室.ItemData(cbo执行科室.ListIndex)
End Sub

Private Sub cbo执行科室_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Not ValidAdvice Then Exit Sub
        
    mlngKey = SaveAdviceData
    If mlngKey = 0 Then
        MsgBox "保存失败", vbInformation, gstrSysName
        Exit Sub
    Else
        If Me.chkConcatenation.Value = 1 Then
'            Me.txt姓名.Text = "": Me.txt姓名.Tag = "":
            Me.txt姓名.SetFocus
            If Me.cbo开单科室.ListIndex > -1 Then
                mlngReqDept = Me.cbo开单科室.ItemData(Me.cbo开单科室.ListIndex)
            End If
            If Me.cbo医生.ListIndex > -1 Then
                mstrReqDoctor = Me.cbo医生.ItemData(Me.cbo医生.ListIndex)
            End If
        Else
'            Me.txt姓名.Text = "": Me.txt姓名.Tag = "":
            txtUnit.Text = "": Me.txt医嘱内容.Text = "": Me.txt医嘱内容.Tag = "": Me.txt姓名.SetFocus
        End If
    End If
End Sub

Private Sub cmdSelect_Click()
    Dim strExtData As String
    Dim rsTmp As New ADODB.Recordset
    
    strExtData = frmLabSamplingSelect.ShowMe(Me, mlngDeptID)
    If strExtData <> "" Then
        '获取采集方式
        Set rsTmp = SelectCap(Split(Split(strExtData, ";")(0), ",")(0))
        If rsTmp Is Nothing Then
            MsgBox "没有定义标本采集方式，请到诊疗项目管理中设置。", vbInformation, gstrSysName
            Exit Sub
        End If
        mlngCapID = rsTmp("ID")
        Call AdviceSet检查手术(3, strExtData)
        txt医嘱内容.Text = Get检查手术名称(2, "")
        txt医嘱内容.Text = txt医嘱内容.Text & "(" & Split(strExtData, ";")(1) & ")"
    End If
End Sub

Private Sub Form_Load()
    InitDepts                     '取得科室和性别
End Sub

Private Sub Form_Unload(Cancel As Integer)
    zlDatabase.SetPara "采集工作站登记", chkConcatenation.Value, 100, 1211
End Sub

Private Sub Text1_Change()

End Sub

Private Sub txtUnit_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txt年龄_GotFocus()
    zlControl.TxtSelAll txt年龄
End Sub

Private Sub txt年龄_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
    
        If Len(Trim(Me.cbo开单科室.Text)) <= 0 Then
            Me.cbo开单科室.SetFocus
        ElseIf Len(Trim(Me.cbo医生.Text)) <= 0 Then
            Me.cbo医生.SetFocus
'        ElseIf Len(Trim(Me.cbo执行科室.Text)) <= 0 Then
'            Me.cbo执行科室.SetFocus
        ElseIf Len(Trim(Me.txt医嘱内容.Text)) <= 0 Then
            Me.txt医嘱内容.SetFocus
        Else
            Me.cmdOK.SetFocus
        End If
    Else
        KeyAscii = FilterKeyAscii(KeyAscii, 99, "0123456789")
    End If
End Sub

Private Sub txt姓名_GotFocus()
    zlControl.TxtSelAll txt姓名
End Sub

Private Sub txt姓名_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then
        KeyCode = Asc(UCase(Chr(KeyCode)))
    Else
        zlCommFun.PressKey vbKeyTab
    End If
End Sub


Private Sub txt姓名_Validate(Cancel As Boolean)
    Dim strInput As String
    Dim rsTmp As New ADODB.Recordset, i As Integer
    Dim strField As String
    Dim strBarCode As String
    Dim rsDept As ADODB.Recordset, strSQL As String
    Dim intSelect As Integer
    Dim strAge As String
    Dim aAge() As String
    
    If Len(Trim(txt姓名)) = 0 Then Exit Sub
    If txt姓名 = txt姓名.Tag Then Exit Sub
    
    Call AdjustEditState(True)

    
    mblnSaveAdvice = True
    Cancel = Not StrIsValid(txt姓名.Text, txt姓名.MaxLength)
    
    Me.cbo开单科室.ListIndex = -1
    Me.cbo医生.ListIndex = -1
'    Me.txt医嘱内容.Text = ""
    
    '初始病人信息
    Set rsTmp = GetPatient(txt姓名)
    strBarCode = txt姓名
    If rsTmp.EOF Then
        mlng病人ID = 0
        '登记新病人
        mstrKeys = ""
        Me.txt年龄 = "": Me.txt年龄1 = "": Me.cboAge.ListIndex = 0
        Me.txtPatientDept = "": Me.txtPatientDept.Tag = 0
        Me.txtID = "": Me.txtBed = ""
        '如果想输入院内病人，则不允许继续
        If InStr("+-*./", Left(Me.txt姓名.Text, 1)) > 0 Or mblnBarCode Then
            Me.txt姓名.Text = "": Cancel = True
            Exit Sub
        End If
        PatientType = 1
        '处理登记的默认科室、医生
        If mlngReqDept > 0 Then
            cbo开单科室.ListIndex = FindComboItem(cbo开单科室, mlngReqDept)
            Me.cbo医生.Text = mstrReqDoctor
        End If
    Else
        On Error Resume Next
        Me.txt姓名.Text = Nvl(rsTmp("姓名"))
        Me.txt年龄.Text = "": Me.txt年龄1.Text = ""
        strAge = IIf(IsNull(rsTmp("年龄")), "", rsTmp("年龄")): If Me.txt年龄 = "0" Then Me.txt年龄 = ""
        
        strAge = Replace(strAge, "小时", "时")
        strAge = Replace(strAge, "分钟", "分")

        If Trim(Replace(Replace(Replace(Replace(Replace(strAge, "岁", ""), "月", ""), "天", ""), "时", ""), "分", "")) <> "" Then
            If InStr(strAge, "成人") > 0 Or InStr(strAge, "婴儿") > 0 Then
                Me.txt年龄.Text = ""
                Me.cboAge.Text = Trim(strAge)
            Else
                strAge = Replace(Replace(Replace(Replace(Replace(strAge, "岁", "岁;"), "月", "月;"), "天", "天;"), "时", "时;"), "分", "分;")
                aAge = Split(strAge, ";")
                If UBound(aAge) = 1 Then
                    Me.txt年龄.Text = Val(aAge(0))
                    Me.cboAge.Text = Replace(Replace(Right(aAge(0), 1), "分", "分钟"), "时", "小时")
                Else
                    Me.txt年龄.Text = Val(aAge(0))
                    Me.cboAge.Text = Replace(Replace(Right(aAge(0), 1), "分", "分钟"), "时", "小时")
                    Me.txt年龄1.Text = Val(aAge(1)) & Replace(Replace(Right(aAge(1), 1), "分", "分钟"), "时", "小时")
                End If
            End If
        Else
            Me.txt年龄.Text = ""
            Me.cboAge.ListIndex = 0
        End If
'        Me.txt年龄 = IIf(IsNull(rsTmp("年龄")), "", Val(rsTmp("年龄"))): If Me.txt年龄 = "0" Then Me.txt年龄 = ""
'        Me.cboAge.Text = IIf(IsNull(rsTmp("年龄")), "岁", Replace(rsTmp("年龄"), Val(rsTmp("年龄")), ""))
        If cboAge.ListIndex = -1 Then cboAge.ListIndex = 0
        Me.cbo性别 = Nvl(rsTmp("性别")) ' CombIndex(cbo性别, Nvl(rsTmp("性别")))
        
        mlng病人ID = Nvl(rsTmp("病人ID"), 0): PatientType = Nvl(rsTmp("PatientType"), 1)
            
        '设置默认开单科室、医生
        cbo开单科室.ListIndex = FindComboItem(cbo开单科室, Nvl(rsTmp("病人科室"), 0))
        
        '病人单位
        txtUnit.Text = Nvl(rsTmp("工作单位"))
        DoEvents
        
        strField = ""
        strField = rsTmp.Fields("医生").Name
        If strField = "医生" Then
            Me.cbo医生.Text = Nvl(rsTmp("医生"))
            For i = 0 To Me.cbo医生.ListCount - 1
                If Me.cbo医生.List(i) Like Nvl(rsTmp("医生")) Then
                    Me.cbo医生.ListIndex = i
                    Exit For
                End If
            Next
        End If
        '显示病人科室
        strSQL = "Select 名称 From 部门表 Where ID=[1]"
        Set rsDept = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CLng(Nvl(rsTmp("病人科室"), 0)))
        If rsDept.EOF Then
            Me.txtPatientDept = "": Me.txtPatientDept.Tag = 0
        Else
            Me.txtPatientDept.Text = rsDept("名称"): Me.txtPatientDept.Tag = Nvl(rsTmp("病人科室"), 0)
        End If
        Me.txtID = Nvl(rsTmp("住院号")): If Len(Me.txtID) = 0 Then Me.txtID = Nvl(rsTmp("门诊号"))
        Me.txtBed = Nvl(rsTmp("当前床号"))
    
        '处理登记的默认科室、医生
        If Me.cbo开单科室.ListIndex = -1 And mlngReqDept > 0 Then
            cbo开单科室.ListIndex = FindComboItem(cbo开单科室, mlngReqDept)
            Me.cbo医生.Text = mstrReqDoctor
        End If
    End If
    txt姓名.Tag = txt姓名.Text
    Me.cbo性别.Tag = "新增"
End Sub

Private Sub txt医嘱内容_GotFocus()
    Call zlControl.TxtSelAll(txt医嘱内容)
End Sub

Private Sub txt医嘱内容_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        If txt医嘱内容.Text = txt医嘱内容.Tag Then
            zlCommFun.PressKey vbKeyTab
            Exit Sub
        End If
        
        With txt医嘱内容
            Set rsTmp = SelectDiagItem()
        End With
        
        If rsTmp Is Nothing Then '取消或无数据
            '恢复原值
            txt医嘱内容.Text = txt医嘱内容.Tag
            zlControl.TxtSelAll txt医嘱内容
            txt医嘱内容.SetFocus: Exit Sub
        End If
        '新项目的录入
        '根据选择项目设置缺省医嘱信息
        If AdviceInput(rsTmp) Then
            DoEvents
            '显示已缺省设置的值
            txt医嘱内容.Tag = txt医嘱内容.Text
            Me.cmdOK.SetFocus
        Else
            DoEvents
            '恢复原值
            txt医嘱内容.Text = txt医嘱内容.Tag
            zlControl.TxtSelAll txt医嘱内容

            txt医嘱内容.SetFocus: Exit Sub
        End If
    ElseIf KeyAscii = Asc("*") Then
        KeyAscii = 0
    ElseIf KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
End Sub
Private Sub InitDoctors(ByVal lng科室ID As Long)
'功能：读取当前开单科室中包含的所有人员
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    
    Me.cbo医生.Clear
    
    '科室医生或护士
    strSQL = _
        "Select Distinct A.ID,B.部门ID,A.编号,A.姓名,Upper(A.简码) as 简码," & _
        " C.人员性质,Nvl(A.聘任技术职务,0) as 职务" & _
        " From 人员表 A,部门人员 B,人员性质说明 C" & _
        " Where A.ID=B.人员ID And A.ID=C.人员ID" & _
        " And C.人员性质 IN('医生') And B.部门ID=[1] " & _
        " And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) "
        
    strSQL = strSQL & " Order by 简码,人员性质 Desc"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng科室ID)
    
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cbo医生.AddItem rsTmp!姓名
            cbo医生.ItemData(cbo医生.ListCount - 1) = rsTmp!部门ID
            
            If rsTmp!ID = UserInfo.ID And cbo医生.ListIndex = -1 Then cbo医生.ListIndex = cbo医生.NewIndex
            rsTmp.MoveNext
        Next
        
        If cbo医生.ListCount = 1 And cbo医生.ListIndex = -1 Then cbo医生.ListIndex = 0
    End If
End Sub
Public Sub ShowMe(Objfrm As Object)
    Me.Show vbModal, Objfrm
End Sub

Private Function InitDepts() As Boolean
'功能：初始化住院临床科室
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim strOldText As String
    Dim intLoop As Integer
    
    On Error GoTo errH
    
    strSQL = _
        " Select Distinct A.ID,A.编码,A.名称" & _
        " From 部门表 A,部门性质说明 B " & _
        " Where B.部门ID = A.ID " & _
        " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) " & _
        " And (B.工作性质 IN('检验'))" & _
        " Order by A.编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    With Me.cbo执行科室
        .AddItem ""
        Do While Not rsTmp.EOF
            .AddItem Nvl(rsTmp("名称"))
            .ItemData(.NewIndex) = rsTmp("ID")
            rsTmp.MoveNext
        Loop
        If .ListCount > 0 And .ListIndex < 0 Then
            .ListIndex = 0
        End If
    End With
    
    
    strOldText = Me.cbo开单科室.Text
    Me.cbo开单科室.Clear
    
    strSQL = _
        " Select Distinct A.ID,A.编码,A.名称" & _
        " From 部门表 A,部门性质说明 B " & _
        " Where B.部门ID = A.ID " & _
        " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) " & _
        " And (B.工作性质 IN('临床','体检'))" & _
        " Order by A.编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    For i = 1 To rsTmp.RecordCount
        cbo开单科室.AddItem rsTmp!名称
        cbo开单科室.ItemData(cbo开单科室.NewIndex) = rsTmp!ID
        
        rsTmp.MoveNext
    Next
    
    On Error Resume Next
    Me.cbo开单科室.Text = strOldText
    If cbo开单科室.ListCount > 0 And Me.cbo开单科室.ListIndex = -1 Then cbo开单科室.ListIndex = 0
    
    
    
    
     '性别
    Set rsTmp = Nothing
    Set rsTmp = GetDictData("性别")
    cbo性别.Clear
    If Not rsTmp Is Nothing Then
        For intLoop = 1 To rsTmp.RecordCount
            cbo性别.AddItem rsTmp!名称
            If rsTmp!缺省 = 1 Then
                cbo性别.ItemData(cbo性别.NewIndex) = 1
                cbo性别.ListIndex = cbo性别.NewIndex
            End If
            rsTmp.MoveNext
        Next
    End If
    
    chkConcatenation.Value = zlDatabase.GetPara("采集工作站登记", 100, 1211, 0)
    
    InitDepts = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub AdjustEditState(blEnable As Boolean)
    '功能:              调整编辑状态
    'Me.txt姓名.Enabled = blEnable
    cbo性别.Enabled = blEnable
    txt年龄.Enabled = blEnable
    txt年龄1.Enabled = blEnable
    cboAge.Enabled = blEnable
    cbo开单科室.Enabled = blEnable
    cbo医生.Enabled = blEnable
    txt医嘱内容.Enabled = blEnable
    cmdSelect.Enabled = blEnable
End Sub
Private Function GetPatient(strCode As String) As ADODB.Recordset
'功能：读取病人信息，并显示该病人存在的医嘱时间
    Dim strSQL As String, i As Long
    Dim strNO As String, str姓名 As String, lng病人ID As Long
    Dim strSeek As String
    
    On Error GoTo errH
    
    If BlnIsNumber(strCode) Then
    '预置条码单独处理
        mblnBarCode = True
        strSQL = "Select Decode(A.当前科室id,Null,1,2) As PatientType,B.主页ID,B.病人科室id As 病人科室,B.开嘱医生 As 医生," & gConst_病人信息_列名 & _
            " From 病人信息 A,病人医嘱记录 B,病人医嘱发送 C Where A.病人ID=B.病人ID+0 And B.ID=C.医嘱ID+0" & _
            " And C.样本条码=[1]"
        Set GetPatient = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strCode)
        Exit Function
    End If
    mblnBarCode = False
    
    strSeek = strCode
    '判断当前输入模式
    If IsNumeric(strCode) And IsNumeric(Left(strCode, 1)) And iInputType = -1 Then '刷卡
        iInputType = 0
    ElseIf (Left(strCode, 1) = "A" Or Left(strCode, 1) = "-") And IsNumeric(Mid(strCode, 2)) Then '病人ID
        iInputType = 1
        strSeek = Mid(strCode, 2)
    ElseIf (Left(strCode, 1) = "B" Or Left(strCode, 1) = "+") And IsNumeric(Mid(strCode, 2)) Then '住院号
        iInputType = 2
        strSeek = Mid(strCode, 2)
    ElseIf (Left(strCode, 1) = "D" Or Left(strCode, 1) = "*") And IsNumeric(Mid(strCode, 2)) Then '门诊号
        iInputType = 3
        strSeek = Mid(strCode, 2)
    ElseIf Left(strCode, 1) = "G" Or Left(strCode, 1) = "." Then '挂号单
        iInputType = 4
        strSeek = Mid(strCode, 2)
    ElseIf Left(strCode, 1) = "/" Then '收费单据号
        iInputType = 5
        strSeek = Mid(strCode, 2)
    ElseIf Not IsNumeric(Mid(strCode, 2)) Then '当作姓名
        iInputType = 6
        strSeek = Replace(strCode, "(婴儿)", "")
    End If
    
    If iInputType = 0 Then '刷卡
        strSQL = "Select Decode(A.当前科室id,Null,1,2) As PatientType,A.主页ID,Decode(A.当前科室id,Null,Nvl(B.执行部门ID,0),A.当前科室id) As 病人科室,B.执行人 As 医生," & gConst_病人信息_列名 & _
            " From 病人信息 A,病人挂号记录 B Where A.就诊卡号=[1] And A.病人ID=B.病人ID(+) And A.门诊号=B.门诊号(+) and (b.病人ID is null or (b.记录性质 =1 and b.记录状态 =1)) "
'            " And (A.当前科室id IS NOT NULL Or NVL(B.执行状态,1) IN (0,2))"
    ElseIf iInputType = 1 Then '病人ID
        strSQL = "Select Decode(A.当前科室id,Null,1,2) As PatientType,A.主页ID,Nvl(A.当前科室id,0) As 病人科室," & gConst_病人信息_列名 & _
            " From 病人信息 A Where A.病人ID=[2]"
    ElseIf iInputType = 2 Then '住院号
        strSQL = "Select Decode(A.当前科室id,Null,1,2) As PatientType,A.主页ID,Decode(A.当前科室id,Null,Nvl(B.入院科室ID,0),A.当前科室id) As 病人科室,B.住院医师 As 医生," & gConst_病人信息_列名 & _
            " From 病人信息 A,病案主页 B Where A.住院号=[2] And A.病人ID=B.病人ID" ' And A.当前科室id IS NOT NULL And B.出院日期 Is NULL"
    ElseIf iInputType = 3 Then '门诊号
        strSQL = "Select Decode(A.当前科室id,Null,1,2) As PatientType,A.主页ID,Decode(A.当前科室id,Null,Nvl(B.执行部门ID,0),A.当前科室id) As 病人科室,B.执行人 As 医生," & gConst_病人信息_列名 & _
            " From 病人信息 A,病人挂号记录 B Where A.门诊号=[2] And A.病人ID=B.病人ID(+) And A.门诊号=B.门诊号(+) and (b.病人ID is null or (b.记录性质 =1 and b.记录状态 =1)) "
'            " And (A.当前科室id IS NOT NULL Or NVL(B.执行状态,1) IN (0,2))"
    ElseIf iInputType = 4 Then '挂号单
        strNO = GetFullNO(strSeek, 12)
        strSQL = "Select 1 As PatientType,0 As 主页ID,Nvl(B.执行部门ID,0) As 病人科室,B.执行人 As 医生," & gConst_病人信息_列名 & _
            " From 病人信息 A,门诊费用记录 B " & _
            " Where B.记录性质=4 And B.记录状态 IN(1,3) And B.NO=[3] And B.病人ID=A.病人ID"
    ElseIf iInputType = 5 Then '收费单据号
        strNO = GetFullNO(strSeek, 13): mstrNO = strNO
        
        strSQL = "Select 1 As PatientType,0 As 主页ID,B.开单部门ID As 病人科室,B.开单人 As 医生,B.姓名,B.性别,B.年龄," & _
            "A.病人ID,A.单位电话,A.工作单位,A.单位邮编,A.家庭地址,A.家庭电话,A.家庭地址邮编,A.门诊号,A.身份证号,A.费别,A.医疗付款方式," & _
            "A.国籍,A.婚姻状况,A.民族,A.职业 From 病人信息 A,门诊费用记录 B" & _
            " Where Mod(B.记录性质,10)=1 And B.记录状态 IN(1,3) And B.NO=[3] And B.病人ID=A.病人ID(+) Order By B.病人ID" ' And B.医嘱序号 Is Null"
    Else '当作姓名
        strSQL = "Select Decode(A.当前科室id,Null,1,2) As PatientType,A.主页ID,Nvl(A.当前科室id,0) As 病人科室," & gConst_病人信息_列名 & _
            " From 病人信息 A Where A.姓名=[1] and 1 = 2 " '所有输入姓名的病人当新病人处理
    End If
    
    Set GetPatient = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strSeek, Val(strSeek), strNO)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function SelectDiagItem() As ADODB.Recordset
'选择检验项目
    Dim strSQL As String
    Dim objPoint As POINTAPI
    
    strSQL = "Select Distinct A.ID,A.编码,A.名称,nvl(A.计算单位,'次') As 计算单位,nvl(A.标本部位,' ') As 标本部位," + _
        "Decode(A.类别,'H',Decode(A.操作类型,'1','护理等级','护理常规')," + _
        "'E',Decode(A.操作类型,'1','过敏试验','2','给药途径','3','中药煎法',4,'中药用法','其它')," + _
        "'Z',Decode(A.操作类型,'1','留观','2','住院','3','转科','4','术后','5','出院','6','转院','其它'),A.操作类型) As 项目特性,A.类别 As 类别ID,A.ID As 诊疗项目ID,nvl(执行频率,0) As 执行频率ID,nvl(计算方式,0) As 计算方式ID,nvl(执行安排,0) As 执行安排ID,nvl(计价性质,0) As 计价性质ID,nvl(执行科室,0) As 执行科室ID "
    strSQL = strSQL + "From 诊疗项目目录 A,诊疗项目别名 C,诊疗执行科室 D Where A.ID=C.诊疗项目ID And A.ID=D.诊疗项目ID And A.类别='C' "   'And D.执行科室ID=" & mlngDeptID
    strSQL = strSQL + " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL) " + _
        "And A.服务对象 IN(" & PatientType & ",3) And Nvl(A.单独应用,0)=1 And Nvl(A.适用性别,0) IN (" + _
        IIf(Me.cbo性别.Text Like "*男*", "1,0)", "2,0)") + _
        " And Nvl(A.执行频率,0) IN(0,1)" + _
        " And (A.编码 Like '" + txt医嘱内容 + "%' Or Upper(A.名称) Like '" + txt医嘱内容 + "%' Or Upper(C.简码) Like '" + UCase(txt医嘱内容) + "%')"
            
    Call ClientToScreen(txt医嘱内容.Hwnd, objPoint)
    Set SelectDiagItem = zlDatabase.ShowSelect(Me, strSQL, 0, "选择申请项目", True, Me.txt医嘱内容.Text, "", True, True, True, objPoint.x * 15, objPoint.Y * 15, Me.txt医嘱内容.Height, False, True)
End Function
Private Function AdviceInput(Optional rsInput As ADODB.Recordset = Nothing) As Boolean
'功能：根据新输的诊疗项目(新增或更换)设置缺省的医嘱数据
'参数：rsInput=输入或选择返回的记录集
'返回：本次录入是否有效
    Dim rsTmp As ADODB.Recordset
    Dim strHelpText As String
    Dim strSQL As String
    Dim t_Pati As TYPE_PatiInfoEx
    Dim blnOk As Boolean
    Dim strExtData As String
    
    On Error GoTo errH

    '项目附加数据输入及输入合法性检查
    '---------------------------------------------------------------------------------------------------------------
    If Not rsInput Is Nothing Then txt医嘱内容.Text = rsInput!名称    '暂时显示

    '需要输入更多数据的一些项目
    '---------------------------------------------------------------------------------------------------------------
    '检验项目选择检验标本
    strHelpText = "检验项目"
    If Not rsInput Is Nothing Then
        strExtData = rsInput!诊疗项目id & ";" & rsInput!标本部位    '新输入项目
    Else
        strExtData = mstrExtData    '新输入项目
    End If
    
    On Error Resume Next
    '接口改造：int场合没有传，现在传为0， bytUseType 以前没传，现在传为0
    blnOk = frmAdviceEditEx.ShowMe(Me, Me.txt医嘱内容.Hwnd, t_Pati, 0, 4, 0, 1, PatientType, , , , 0, strExtData, , , , , True, mlngDeptID)
    On Error GoTo errH

    If Not blnOk Then Exit Function
    If strExtData = "" Or Mid(strExtData, 1, 1) = ";" Then Exit Function
    
    '获取采集方式
    Set rsTmp = SelectCap(Split(Split(strExtData, ";")(0), ",")(0))
    If rsTmp Is Nothing Then
        MsgBox "没有定义标本采集方式，请到诊疗项目管理中设置。", vbInformation, gstrSysName
        Exit Function
    End If
    mlngCapID = rsTmp("ID")
    
    strSQL = "Select C.项目类别 From 诊疗项目目录 A,检验报告项目 B,检验项目 C " & _
        "Where A.ID=B.诊疗项目ID And B.报告项目ID=C.诊治项目ID And A.ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Split(Split(strExtData, ";")(0), ",")(0))
    If rsTmp.EOF Then
        mbln微生物项目 = False
    Else
        mbln微生物项目 = IIf(Nvl(rsTmp("项目类别"), 0) = 2, True, False)
    End If
    
    mstrExtData = strExtData
    
    
    Call AdviceSet检查手术(3, mstrExtData)
    txt医嘱内容.Text = Get检查手术名称(2, "")
    txt医嘱内容.Text = txt医嘱内容.Text & "(" & Split(mstrExtData, ";")(1) & ")"
    
    '开嘱医生
    On Error Resume Next
    If Me.cbo医生.Text = "" Then Me.cbo医生.ListIndex = 0
    
    AdviceInput = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function GetDictData(strDict As String) As ADODB.Recordset
'功能：从指定的字典中读取数据
'参数：strDict=字典对应的表名
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
        
    strSQL = "Select 编码,名称,Nvl(缺省标志,0) as 缺省 From " & strDict & " Order by 编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If Not rsTmp.EOF Then Set GetDictData = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function GetFullNO(ByVal strNO As String, ByVal intNum As Integer) As String
'功能：由用户输入的部份单号，返回全部的单号。
'参数：intNum=项目序号,为0时固定按年产生
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, intType As Integer
    Dim curDate As Date
    
    If Len(strNO) >= 8 Then
        GetFullNO = Right(strNO, 8)
        Exit Function
    ElseIf Len(strNO) = 7 Then
        GetFullNO = PreFixNO & strNO
        Exit Function
    ElseIf intNum = 0 Then
        GetFullNO = PreFixNO & Format(Right(strNO, 7), "0000000")
        Exit Function
    End If
    GetFullNO = strNO
    
    strSQL = "Select 编号规则,Sysdate as 日期 From 号码控制表 Where 项目序号=" & intNum
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If Not rsTmp.EOF Then
        intType = Nvl(rsTmp!编号规则, 0)
        curDate = rsTmp!日期
    End If

    If intType = 1 Then
        '按日编号
        strSQL = Format(CDate("1992-" & Format(rsTmp!日期, "MM-dd")) - CDate("1992-01-01") + 1, "000")
        GetFullNO = PreFixNO & strSQL & Format(Right(strNO, 4), "0000")
    Else
        '按年编号
        GetFullNO = PreFixNO & Format(Right(strNO, 7), "0000000")
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function SelectCap(Optional ByVal lngItemID As Long = 0) As ADODB.Recordset
'获取采集方式
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim tmpRect As RECT
    
    On Error GoTo DBError
        
    strSQL = "Select Distinct A.ID,A.编码,A.名称 " + _
        "From 诊疗项目目录 A,诊疗用法用量 D Where A.ID=D.用法ID" + _
        " And A.类别='E' And A.操作类型='6'" & _
        " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL) " + _
        " And A.服务对象 IN(" & PatientType & ",3) And Nvl(A.适用性别,0) IN (" + _
        IIf(Me.cbo性别.Text Like "*男*", "1,0)", "2,0)") + _
        " And Nvl(A.执行频率,0) IN(0,1)" + _
        " And D.项目ID=" & lngItemID
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rsTmp.EOF Then
        strSQL = "Select Distinct A.ID,A.编码,A.名称 " + _
            "From 诊疗项目目录 A Where " + _
            " A.类别='E' And A.操作类型='6'" & _
            " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL) " + _
            " And A.服务对象 IN(" & PatientType & ",3) And Nvl(A.适用性别,0) IN (" + _
            IIf(Me.cbo性别.Text Like "*男*", "1,0)", "2,0)") + _
            " And Nvl(A.执行频率,0) IN(0,1)"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If Not rsTmp.EOF Then Set SelectCap = rsTmp
    
    Exit Function
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub AdviceSet检查手术(ByVal int类型 As Integer, ByVal strDataIDs As String)
'功能：1.重新设置指定检查组合项目的部位行,用于新输入检查组合项目或修改部位
'      2.重新设置指定手术项目的附加手术及麻醉项目行,用于新输入手术项目或手术项目的附加手术及麻醉项目
'参数：int类型=1=处理检查部位项目,2=处理附加手术及麻醉项目
'      strDataIDs=检查:包含检查部位信息,手术:包含附加手术及麻醉项目信息,其中可能没有附加手术和麻醉
    Dim strSQL As String, i As Long
    Dim arrIDs As Variant
    
    On Error GoTo errH
            
    '处理检验项目
    strDataIDs = Mid(strDataIDs, 1, InStr(strDataIDs, ";") - 1)
    
    If strDataIDs <> "" Then
        If Not rsRelativeAdvice Is Nothing Then
            rsRelativeAdvice.Close
        Else
            Set rsRelativeAdvice = New ADODB.Recordset
        End If
        strSQL = "Select ID,编码,名称,nvl(标本部位,' ') As 标本部位," + _
        "类别,nvl(计价性质,0) As 计价性质,nvl(执行科室,0) As 执行科室,操作类型 From 诊疗项目目录 Where ID IN(" & strDataIDs & ")"
        Set rsRelativeAdvice = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Else
        If Not rsRelativeAdvice Is Nothing Then rsRelativeAdvice.Close: Set rsRelativeAdvice = Nothing
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Function Get检查手术名称(ByVal int类型 As Integer, ByVal txtMainAdvice As String) As String
'功能：重新生成检查手术内容的医嘱内容
'参数：int类型=1=处理检查部位项目,2=处理附加手术及麻醉项目
    Dim lngBegin As Long, i As Long
    Dim str麻醉 As String, strTmp As String
    Dim strDate As String
    
    If rsRelativeAdvice Is Nothing Or int类型 = 1 Then Get检查手术名称 = txtMainAdvice: Exit Function
        
    rsRelativeAdvice.MoveFirst
    Do While Not rsRelativeAdvice.EOF
        If Len(Trim(rsRelativeAdvice("名称"))) > 0 Then
            strTmp = strTmp & "," & rsRelativeAdvice("名称")
        End If
        
        rsRelativeAdvice.MoveNext
    Loop
    
    If strTmp <> "" Then
        Get检查手术名称 = IIf(Len(Trim(txtMainAdvice)) = 0, "", txtMainAdvice & " 及 ") & Mid(strTmp, 2)
    Else
        Get检查手术名称 = txtMainAdvice
    End If
End Function
'检查医嘱内容的合法性
Private Function ValidAdvice() As Boolean
    ValidAdvice = True
    
    On Error Resume Next
    If txt姓名.Text = "" Then
        ValidAdvice = False
        MsgBox "请输入病人的姓名！", vbInformation, gstrSysName: DoEvents
'        mintFocusItem = FocusItem.姓名
        txt姓名.SetFocus: Exit Function
    End If
    
    If Len(Trim(Me.txt医嘱内容)) = 0 Then
        ValidAdvice = False
        MsgBox "必须输入申请项目！", vbInformation, gstrSysName: DoEvents
'        mintFocusItem = FocusItem.医嘱内容
        Me.txt医嘱内容.SetFocus: Exit Function
    End If
    If Me.cbo开单科室.ListIndex = -1 Then
        ValidAdvice = False
        MsgBox "请指定开单科室！", vbInformation, gstrSysName: DoEvents
'        mintFocusItem = FocusItem.开单科室
        Me.cbo开单科室.SetFocus: Exit Function
    End If
'    If Me.cbo执行科室.ListIndex = -1 Then
'        ValidAdvice = False
'        MsgBox "请指定执行科室!", vbInformation, gstrSysName: DoEvents
'        Me.cbo执行科室.SetFocus: Exit Function
'    End If
    If Len(Trim(Me.cbo医生.Text)) = 0 Then
        ValidAdvice = False
        MsgBox "请指定开单医生！", vbInformation, gstrSysName: DoEvents
'        mintFocusItem = FocusItem.医生
        Me.cbo医生.SetFocus: Exit Function
    End If
End Function
Private Function SaveAdviceData() As Long
    Dim strSQL As String, strDate As String, strNO As String
    Dim lngAdviceID As Long, lngTmpID As Long, lngSendNO As Long
    Dim iMaxSeq As Integer, iSendSeq As Integer
    Dim rsTmp As New ADODB.Recordset
    Dim lng开嘱科室ID As Long, lng病人ID As Long, strDoctor As String, i As Integer
    Dim str执行科室ID As String, str执行科室ID1 As String, lngDept As Long
    Dim rsCard As ADODB.Recordset
    Dim tmpstr类别 As String, tmplngClinicID As Long, tmpint计价特性 As Integer, tmpint执行性质 As Integer
    Dim rsDept As ADODB.Recordset
    Dim intPatientSource As Integer                     '病人来源
    Dim lngJ As Long, strCostType As String
    
    Dim strAge As String
    Dim strInfo As String
    Dim lngTmp As Long
    
    On Error GoTo ErrHand
    gcnOracle.BeginTrans
    
    '保存病人信息
    strDate = "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    If PatientType = 1 Then '门诊病人
        If mlng病人ID > 0 Then '已有的病人
'            strSQL = _
                "zl_挂号病人病案_INSERT(3," & mlng病人ID & ",Null," & _
                "'',''," & _
                "'" & txt姓名.Text & "','" & NeedName(cbo性别.Text) & "','" & txt年龄.Text & Me.cboAge.Text & Me.txt年龄1.Text & "'," & _
                "'自费','自费'," & _
                "'','',''," & _
                "'','','',0,'','','','',''," & strDate & ",NULL)"
        Else '新病人
            If txt年龄.Locked = False Then
                strAge = txt年龄.Text
                If IsNumeric(strAge) Then strAge = strAge & cboAge.Text & txt年龄1.Text
                strInfo = CheckAge(strAge)
                If InStr(1, strInfo, "|") > 0 Then
                    lngTmp = Val(Split(strInfo, "|")(0)) '1禁止,0提示
                    strInfo = Split(strInfo, "|")(1)
                    If lngTmp = 1 Then
                        MsgBox strInfo, vbInformation, gstrSysName
                        gcnOracle.RollbackTrans
                        If txt年龄.Enabled And txt年龄.Visible Then txt年龄.SetFocus: Exit Function
                    End If
                End If
            End If
            '添加获取默认费别
            strSQL = "select 名称,缺省标志 from 费别 order by 编码"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlLisWork")
            Do While Not rsTmp.EOF
                lngJ = lngJ + 1
                If lngJ = 1 Then
                    strCostType = rsTmp("名称")
                End If
                If rsTmp("缺省标志") = 1 Then
                    strCostType = rsTmp("名称")
                    Exit Do
                End If
                rsTmp.MoveNext
            Loop
            If strCostType = "" Then strCostType = "自费"
            
            mlng病人ID = zlDatabase.GetNextNo(1)
            strSQL = _
                "zl_挂号病人病案_INSERT(1," & mlng病人ID & ",Null," & _
                "'',''," & _
                "'" & txt姓名.Text & "','" & NeedName(cbo性别.Text) & "','" & txt年龄.Text & Me.cboAge.Text & Me.txt年龄1.Text & "'," & _
                "'" & strCostType & "','" & strCostType & "'," & _
                "'','',''," & _
                "'','','" & Me.txtUnit.Text & "',0,'','','','',''," & strDate & ",NULL)"
            zlDatabase.ExecuteProcedure strSQL, "病人信息保存"
        End If
    End If
    '保存医嘱并发送
    lngAdviceID = zlDatabase.GetNextId("病人医嘱记录")
    iMaxSeq = 0
    
    lng开嘱科室ID = Me.cbo开单科室.ItemData(Me.cbo开单科室.ListIndex)
    strDoctor = NeedName(Me.cbo医生.Text)
    
    If rsRelativeAdvice.RecordCount = 0 Then
        str执行科室ID = mlngDeptID
    Else
        'PatientType
        If mlng病人ID > 0 Then
            strSQL = "select  执行科室ID from  诊疗执行科室 where 病人来源 = [1] and 诊疗项目ID = [2] "
        Else
            strSQL = "select 执行科室id from 诊疗执行科室 where 诊疗项目id = [2]"
        End If
        rsRelativeAdvice.MoveFirst
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, PatientType, CLng(rsRelativeAdvice("Id")))
        str执行科室ID = Val(Nvl(rsTmp("执行科室ID")))
    End If
    
    '选择了执行科室按执行科室进行
    If Me.cbo执行科室.Text <> "" Then
        str执行科室ID = Me.cbo执行科室.ItemData(Me.cbo执行科室.ListIndex)
    End If
    
    iSendSeq = 1
    '检验项目将采集方式作为主医嘱
    tmplngClinicID = mlngCapID
    '取采集方式的执行部门
    str执行科室ID1 = UserInfo.部门ID
    
    lngSendNO = zlDatabase.GetNextNo(10)
    strNO = zlDatabase.GetNextNo(IIf(PatientType = 2, 14, 13))
    
    '保存相关医嘱
    If Not rsRelativeAdvice Is Nothing Then
        i = 2
        rsRelativeAdvice.MoveFirst
        Do While Not rsRelativeAdvice.EOF
            lngTmpID = zlDatabase.GetNextId("病人医嘱记录")
            With rsRelativeAdvice
                strSQL = "ZL_病人医嘱记录_Insert(" & lngTmpID & "," & lngAdviceID & "," & _
                    (iMaxSeq + i) & ",3," & mlng病人ID & ",NULL," & _
                    "0,1," & _
                    "1,'" & .Fields("类别") & "'," & _
                    .Fields("ID") & ",NULL,NULL,NULL,NULL," & _
                    "'" & Replace(.Fields("名称"), "'", "''") & "',''," & _
                    "'" & .Fields("标本部位") & "','一次性',NULL,NULL,'',NULL," & _
                    .Fields("计价性质") & "," & _
                    str执行科室ID & "," & _
                    .Fields("执行科室") & ",0," & strDate & ",NULL," & _
                    IIf(Me.txtPatientDept.Tag = 0, lng开嘱科室ID, Me.txtPatientDept.Tag) & "," & lng开嘱科室ID & ",'" & strDoctor & "'," & _
                    "Sysdate,'',Null)"
                    zlDatabase.ExecuteProcedure strSQL, Me.Caption
                iSendSeq = iSendSeq + 1
                strSQL = "ZL_病人医嘱发送_Insert(" & _
                    lngTmpID & "," & lngSendNO & "," & PatientType & ",'" & strNO & "'," & _
                    iSendSeq & ",NULL,NULL,NULL," & _
                    "Sysdate+1/(24*3600)," & _
                    "0," & str执行科室ID & ",0,0)"
                zlDatabase.ExecuteProcedure strSQL, Me.Caption
                i = i + 1
                .MoveNext
            End With
        Loop
    End If
    '检验申请的采集方式放到最后
    iMaxSeq = iMaxSeq + 1
    strSQL = "ZL_病人医嘱记录_Insert(" & lngAdviceID & ",NULL," & _
        iMaxSeq & ",3," & mlng病人ID & ",NULL," & _
        "0,1," & _
        "1,'E'," & mlngCapID & ",NULL,NULL,NULL,NULL," & _
        "'" & Replace(Me.txt医嘱内容, "'", "''") & "',''," & _
        "'','一次性',NULL,NULL,'',NULL,2," & _
        str执行科室ID1 & ",3,0," & strDate & ",NULL," & _
        IIf(Me.txtPatientDept.Tag = 0, lng开嘱科室ID, Me.txtPatientDept.Tag) & "," & lng开嘱科室ID & ",'" & strDoctor & "'," & _
        "Sysdate,'',Null)"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    iSendSeq = iSendSeq + 1
    '发送主医嘱
    strSQL = "ZL_病人医嘱发送_Insert(" & _
        lngAdviceID & "," & lngSendNO & "," & PatientType & ",'" & strNO & "'," & _
        iSendSeq & ",NULL,NULL,NULL," & _
        "Sysdate+1/(24*3600)," & _
        "0," & str执行科室ID & ",0,1)"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    SaveAdviceData = mlng病人ID
    gcnOracle.CommitTrans
    
    Exit Function
ErrHand:
    mlng病人ID = 0
    gcnOracle.RollbackTrans
'    Err.Raise Err.Number, "标本核收"
    Exit Function
End Function

Private Function PreFixNO(Optional curDate As Date = #1/1/1900#) As String
'功能：返回大写的单据号年前缀
    If curDate = #1/1/1900# Then
        PreFixNO = CStr(CInt(Format(zlDatabase.Currentdate, "YYYY")) - 1990)
    Else
        PreFixNO = CStr(CInt(Format(curDate, "YYYY")) - 1990)
    End If
    PreFixNO = IIf(CInt(PreFixNO) < 10, PreFixNO, Chr(55 + CInt(PreFixNO)))
End Function
