VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPACSFilter 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "过滤条件"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
   ControlBox      =   0   'False
   Icon            =   "frmPACSFilter.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox cboPart 
      Height          =   300
      Left            =   1170
      Style           =   2  'Dropdown List
      TabIndex        =   34
      Top             =   3540
      Width           =   3870
   End
   Begin VB.CommandButton cmdClear 
      Height          =   300
      Left            =   4725
      Picture         =   "frmPACSFilter.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "清除报告内容列表"
      Top             =   4305
      Width           =   300
   End
   Begin VB.ComboBox cboContent 
      Height          =   300
      Left            =   1170
      TabIndex        =   13
      Top             =   4305
      Width           =   3570
   End
   Begin VB.ComboBox cboItem 
      Height          =   300
      Left            =   1170
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   3915
      Width           =   3870
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Index           =   3
      Left            =   -15
      TabIndex        =   30
      Top             =   4665
      Width           =   5775
   End
   Begin VB.OptionButton optAdviceTime 
      Caption         =   "按开嘱时间查找&A(直接在医技站中完成的，可用此项查找)"
      Height          =   180
      Left            =   75
      TabIndex        =   5
      Top             =   1995
      Width           =   5130
   End
   Begin VB.OptionButton optCheckTime 
      Caption         =   "按检查时间查找&T(推荐)"
      Height          =   270
      Left            =   75
      TabIndex        =   4
      Top             =   1680
      Value           =   -1  'True
      Width           =   2490
   End
   Begin VB.TextBox txtChkNO 
      Height          =   300
      Left            =   1170
      MaxLength       =   10
      TabIndex        =   8
      Top             =   2655
      Width           =   795
   End
   Begin VB.TextBox txt姓名 
      Height          =   300
      Left            =   1170
      TabIndex        =   2
      Top             =   1215
      Width           =   1185
   End
   Begin VB.TextBox txtNO 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   3420
      MaxLength       =   10
      TabIndex        =   3
      Top             =   1215
      Width           =   1185
   End
   Begin VB.TextBox txt就诊卡 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   3420
      MaxLength       =   10
      TabIndex        =   1
      Top             =   855
      Width           =   1185
   End
   Begin VB.TextBox txt标识号 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1170
      MaxLength       =   10
      TabIndex        =   0
      Top             =   855
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Index           =   2
      Left            =   0
      TabIndex        =   24
      Top             =   1620
      Width           =   5775
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Index           =   1
      Left            =   0
      TabIndex        =   22
      Top             =   720
      Width           =   5775
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Index           =   0
      Left            =   -105
      TabIndex        =   21
      Top             =   3435
      Width           =   5775
   End
   Begin VB.CommandButton cmdDefault 
      Cancel          =   -1  'True
      Caption         =   "缺省(&D)"
      Height          =   350
      Left            =   330
      TabIndex        =   20
      Top             =   4785
      Width           =   1100
   End
   Begin VB.CheckBox chk来源 
      Caption         =   "住院病人"
      Height          =   195
      Index           =   1
      Left            =   4020
      TabIndex        =   10
      Top             =   2730
      Value           =   1  'Checked
      Width           =   1020
   End
   Begin VB.CheckBox chk来源 
      Caption         =   "门诊病人"
      Height          =   195
      Index           =   0
      Left            =   2910
      TabIndex        =   9
      Top             =   2715
      Value           =   1  'Checked
      Width           =   1020
   End
   Begin VB.ComboBox cboDept 
      Height          =   300
      Left            =   1170
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   3060
      Width           =   3870
   End
   Begin MSComCtl2.DTPicker dtpEnd 
      Height          =   300
      Left            =   3210
      TabIndex        =   7
      Top             =   2295
      Width           =   1830
      _ExtentX        =   3228
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   84738051
      CurrentDate     =   38082
   End
   Begin MSComCtl2.DTPicker dtpBegin 
      Height          =   300
      Left            =   1170
      TabIndex        =   6
      Top             =   2295
      Width           =   1830
      _ExtentX        =   3228
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   84738051
      CurrentDate     =   38082
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2775
      TabIndex        =   15
      Top             =   4785
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3945
      TabIndex        =   16
      Top             =   4785
      Width           =   1100
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "检查部位"
      Height          =   180
      Left            =   330
      TabIndex        =   33
      Top             =   3600
      Width           =   720
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "报告内容"
      Height          =   180
      Left            =   330
      TabIndex        =   32
      Top             =   4365
      Width           =   720
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "报告项目"
      Height          =   180
      Left            =   330
      TabIndex        =   31
      Top             =   3975
      Width           =   720
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "检查号"
      Height          =   285
      Left            =   510
      TabIndex        =   29
      Top             =   2745
      Width           =   705
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "姓名(&3)"
      Height          =   180
      Left            =   510
      TabIndex        =   28
      Top             =   1275
      Width           =   630
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "单据号(&4)"
      Height          =   180
      Left            =   2595
      TabIndex        =   27
      Top             =   1275
      Width           =   810
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "就诊卡(&2)"
      Height          =   180
      Left            =   2610
      TabIndex        =   25
      Top             =   915
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "标识号(&1)"
      Height          =   180
      Left            =   330
      TabIndex        =   26
      Top             =   915
      Width           =   810
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   270
      Picture         =   "frmPACSFilter.frx":0596
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblnote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    通过过滤条件设置以便准确查找执行记录；建议时间范围尽量精确，以查找保证效率。"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   915
      TabIndex        =   23
      Top             =   180
      Width           =   4035
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "病人来源"
      Height          =   180
      Left            =   2070
      TabIndex        =   19
      Top             =   2745
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "病人科室"
      Height          =   180
      Left            =   330
      TabIndex        =   18
      Top             =   3120
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "时间范围                      ～"
      Height          =   180
      Left            =   330
      TabIndex        =   17
      Top             =   2355
      Width           =   2880
   End
End
Attribute VB_Name = "frmPACSFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mstrFilter As String
Public mstrPati As String
Private mblnLoad As Boolean
Public mBeforeDays As Integer '默认查询的天数
Public mblnOK As Boolean
Public FindType As Integer '时间查找方式：1＝按检查时间、2＝按开嘱时间

Private Sub cboContent_LostFocus()
    Dim i As Integer
    With cboContent
        If Len(Trim(.Text)) = 0 Then Exit Sub
        
        For i = 0 To .ListCount - 1
            If Trim(.Text) = .List(i) Then Exit For
        Next
        If i > .ListCount - 1 Then .AddItem .Text
    End With
End Sub

Private Sub chk来源_Click(Index As Integer)
    If chk来源(0).Value = 0 And chk来源(1).Value = 0 Then
        chk来源((Index + 1) Mod 2).Value = 1
    End If
    Call LoadDept
End Sub

Private Sub cmdCancel_Click()
    mstrFilter = ""
    mstrPati = ""
    mblnOK = False
    Call SaveSetting("ZLSOFT", "公共模块\" & App.ProductName & "\frmPACSFilter", "过滤报告项目", cboItem.Text)
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\frmPACSFilter", "过滤病人科室", cboDept.Text)
    Me.Hide
End Sub

Private Sub cmdClear_Click()
    cboContent.Clear
End Sub

Private Sub cmdDefault_Click()
    Me.optCheckTime.Value = True: mBeforeDays = 2
    Call Form_Load
End Sub

Private Sub cmdOK_Click()
    Call txtNO_Validate(False)
    Call MakeFilter(mstrFilter, mstrPati)
    
    mBeforeDays = dtpEnd.Value - dtpBegin.Value
    FindType = IIf(Me.optCheckTime.Value, 1, 2)
    mblnOK = True
    Call SaveSetting("ZLSOFT", "公共模块\" & App.ProductName & "\frmPACSFilter", "过滤报告项目", cboItem.Text)
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\frmPACSFilter", "过滤病人科室", cboDept.Text)
    Me.Hide
End Sub

Private Sub MakeFilter(strFilter As String, strPati As String)
'功能：产生条件(病人医嘱发送 A,病人医嘱记录 B)
    Dim strTmp As String
    
    '发送时间
    If Format(dtpEnd.Value, "yyyy-MM-dd HH:mm") = Format(dtpEnd.Tag, "yyyy-MM-dd HH:mm") Then
        strFilter = " And A.发送时间 Between To_Date('" & Format(dtpBegin.Value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS') And Sysdate"
    Else
        strFilter = " And A.发送时间 Between To_Date('" & Format(dtpBegin.Value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')" & _
            " And To_Date('" & Format(dtpEnd.Value, "yyyy-MM-dd HH:mm:59") & "','YYYY-MM-DD HH24:MI:SS')"
    End If
    
    '单据号
    If txtNO.Text <> "" Then
        strFilter = strFilter & " And A.NO='" & txtNO.Text & "'"
    End If
    
    '病人科室
    If cboDept.ListIndex <> 0 Then
        strFilter = strFilter & " And B.病人科室ID+0=" & cboDept.ItemData(cboDept.ListIndex)
    End If
    
    '病人来源
    strFilter = strFilter & " And Nvl(B.病人来源,0) IN(3," & IIf(chk来源(0).Value, "1,4", "-1") & "," & IIf(chk来源(1).Value, 2, -1) & ")"
        
    '标本部位
    If Trim(Me.cboPart) <> "" Then
        strFilter = strFilter & " And B.标本部位 = '" & Me.cboPart.Text & "' "
    End If
    
    '病人标识
    strPati = ""
    If txt标识号.Text <> "" Then
        strPati = strPati & " And Decode(B.病人来源,1,D.门诊号,2,D.住院号,NULL)=" & txt标识号.Text
    End If
    If txt就诊卡.Text <> "" Then
        strPati = strPati & " And D.就诊卡号||''='" & txt就诊卡.Text & "'"
    End If
    If txt姓名.Text <> "" Then
        strPati = strPati & " And D.姓名||''='" & txt姓名.Text & "'"
    End If
    If txtChkNO.Text <> "" Then
        strPati = strPati & " And H.检查号=" & txtChkNO.Text
    End If
    
    
End Sub

Private Sub Form_Activate()
    Dim curDate As Date
    
    '如果上一次是取的当前时间,则重新设置时刷新结果时间为当前时间
    If Not mblnLoad Then
        If Format(dtpEnd.Value, "yyyy-MM-dd HH:mm") = Format(dtpEnd.Tag, "yyyy-MM-dd HH:mm") Then
            curDate = zlDatabase.Currentdate
            dtpEnd.MaxDate = curDate: dtpBegin.MaxDate = curDate
            dtpEnd.Value = Format(curDate, "yyyy-MM-dd HH:mm")
            dtpEnd.Tag = Format(dtpEnd.Value, "yyyy-MM-dd HH:mm")
        End If
    End If
    If mblnLoad Then mblnLoad = False
        
    '自动定位
    dtpBegin.SetFocus
    If txtNO.Text <> "" Then
        txtNO.Text = "": txtNO.SetFocus
    End If
    If txt姓名.Text <> "" Then
        txt姓名.Text = "": txt姓名.SetFocus
    End If
    If txt就诊卡.Text <> "" Then
        txt就诊卡.Text = "": txt就诊卡.SetFocus
    End If
    If txt标识号.Text <> "" Then
        txt标识号.Text = "": txt标识号.SetFocus
    End If
    If txtChkNO.Text <> "" Then
        txtChkNO.Text = "": txtChkNO.SetFocus
    End If
    '报告项目
    Call InitRptItem
    On Error Resume Next
    cboItem.Text = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\frmPACSFilter", "过滤报告项目", "")
    cboContent.Text = ""
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim curDate As Date
    Dim aContent() As String, i As Long
    Dim strTmp As String
    
    mblnLoad = True
    
    txtNO.Text = ""
    txt标识号.Text = ""
    txt姓名.Text = ""
    txt就诊卡.Text = ""
    txt就诊卡.PasswordChar = IIf(gblnCardHide, "*", "")
    txtChkNO.Text = ""
    
    '来源和状态
    chk来源(0).Value = 1
    chk来源(1).Value = 1
    
    '发送时间
    curDate = zlDatabase.Currentdate
    If mBeforeDays <= 0 Then mBeforeDays = 2 '默认查询3天前的申请
    dtpEnd.MaxDate = curDate: dtpBegin.MaxDate = curDate
    dtpBegin.Value = Format(curDate - mBeforeDays, "yyyy-MM-dd 00:00")
    dtpEnd.Value = Format(curDate, "yyyy-MM-dd HH:mm")
    dtpEnd.Tag = Format(dtpEnd.Value, "yyyy-MM-dd HH:mm")
        
    '病人科室
    Call LoadDept
    strTmp = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\frmPACSFilter", "过滤病人科室", "")
    On Error Resume Next
    If strTmp <> "" Then Me.cboDept.Text = strTmp
    On Error GoTo 0
    '初始报告内容选项
    aContent = Split(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\frmPACSFilter", "过滤报告内容", ""), "|")
    With cboContent
        .Clear: .AddItem ""
        For i = 0 To UBound(aContent)
            .AddItem aContent(i)
        Next
    End With
    
    mstrFilter = ""
    mstrPati = ""
    mblnOK = False
End Sub

Private Function LoadDept() As Boolean
'功能：根据病人来源读取病人科室
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim lngPre As Long
    
    If cboDept.ListIndex <> -1 Then
        lngPre = cboDept.ItemData(cboDept.ListIndex)
    End If
    strSQL = "Select Distinct A.ID,A.编码,A.名称,B.服务对象" & _
        " From 部门表 A,部门性质说明 B" & _
        " Where A.ID=B.部门ID And B.工作性质 IN('临床','手术')" & _
        " And B.服务对象 IN(3," & IIf(chk来源(0).Value, 1, -1) & "," & IIf(chk来源(1).Value, 2, -1) & ")" & _
        " And (A.撤档时间 is NULL Or A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
        " Order by A.编码"
    On Error GoTo errH
    Call OpenRecord(rsTmp, strSQL, Me.Caption)
    On Error GoTo 0
    cboDept.Clear
    cboDept.AddItem "所有科室"
    cboDept.ListIndex = 0
    For i = 1 To rsTmp.RecordCount
        cboDept.AddItem rsTmp!编码 & "-" & rsTmp!名称
        cboDept.ItemData(cboDept.NewIndex) = rsTmp!ID
        If rsTmp!ID = lngPre Then cboDept.ListIndex = cboDept.NewIndex
        rsTmp.MoveNext
    Next
    LoadDept = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function InitRptItem() As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    
    InitRptItem = True
    On Error GoTo errH
    strSQL = "Select Distinct D.标题文本 " & _
        "From 诊疗执行科室 B, 诊疗单据应用 C, 病历文件组成 D " & _
        "Where B.诊疗项目id = C.诊疗项目id AND C.病历文件ID=D.病历文件ID And B.执行科室id = [1] AND D.填写时机=2"
    With frmPACStation
        Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, .cboDept.ItemData(.cboDept.ListIndex))
    End With
    With cboItem
        .Clear: .AddItem " "
        Do While Not rsTmp.EOF
            .AddItem Nvl(rsTmp("标题文本"))
        
            rsTmp.MoveNext
        Loop
    End With
    
    strSQL = "Select Distinct  标本部位  From 诊疗项目目录 Where 类别 = 'D' And 标本部位 Is Not Null"
    zlDatabase.OpenRecordset rsTmp, strSQL, gstrSysName
    With Me.cboPart
        .Clear: .AddItem ""
        Do While Not rsTmp.EOF
            .AddItem Nvl(rsTmp("标本部位"))
            rsTmp.MoveNext
        Loop
    End With
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    InitRptItem = False
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer, strContent As String
    
    '保存报告内容
    strContent = ""
    With cboContent
        For i = 0 To .ListCount - 1
            If Len(Trim(.List(i))) > 0 Then strContent = strContent & "|" & .List(i)
        Next
    End With
    If Len(strContent) > 0 Then strContent = Mid(strContent, 2)
    Call SaveSetting("ZLSOFT", "公共模块\" & App.ProductName & "\frmPACSFilter", "过滤报告内容", strContent)
End Sub

Private Sub optAdviceTime_Click()
    Me.dtpBegin.SetFocus
End Sub

Private Sub optCheckTime_Click()
    Me.dtpBegin.SetFocus
End Sub

Private Sub txtChkNO_GotFocus()
    Call zlControl.TxtSelAll(txtChkNO)
End Sub

Private Sub txtChkNO_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtNO_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii <> 13 Then
        If Not (txtNO.Text = "" Or txtNO.SelLength = Len(txtNO.Text)) _
            And InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtNO_Validate(Cancel As Boolean)
    If IsNumeric(txtNO.Text) Then
        txtNO.Text = GetFullNO(txtNO.Text, 0)
    End If
End Sub

Private Sub txt就诊卡_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txt就诊卡.Hwnd, GWL_WNDPROC)
        Call SetWindowLong(txt就诊卡.Hwnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt就诊卡_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Call SetWindowLong(txt就诊卡.Hwnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt姓名_GotFocus()
    Call zlControl.TxtSelAll(txt姓名)
End Sub

Private Sub txtNO_GotFocus()
    Call zlControl.TxtSelAll(txtNO)
End Sub

Private Sub txt就诊卡_GotFocus()
    Call zlControl.TxtSelAll(txt就诊卡)
End Sub

Private Sub txt就诊卡_KeyPress(KeyAscii As Integer)
    Dim blnCard As Boolean
    
    '去掉磁卡的其他的特殊字符
    If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    
    blnCard = InputIsCard(Me.txt就诊卡, KeyAscii)
    
    '刷卡完成或确认输入
    If blnCard And Len(Me.txt就诊卡.Text) = gbytCardLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Me.txt就诊卡.Text <> "" Then
        If KeyAscii <> 13 Then
            Me.txt就诊卡.Text = Me.txt就诊卡.Text & Chr(KeyAscii)
            Me.txt就诊卡.SelStart = Len(Me.txt就诊卡.Text)
        End If
        KeyAscii = 0
        Me.txt就诊卡.Text = UCase(Me.txt就诊卡)
        Me.txt就诊卡.SetFocus
    End If
End Sub

Private Sub txt标识号_GotFocus()
    Call zlControl.TxtSelAll(txt标识号)
End Sub

Private Sub txt标识号_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
