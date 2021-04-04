VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBillPrivilege 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "单据操作权限设置"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5235
   Icon            =   "frmBillPrivilege.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3990
      TabIndex        =   11
      Top             =   240
      Width           =   1100
   End
   Begin VB.Frame fra基本信息 
      Caption         =   "操作权限"
      Height          =   2385
      Left            =   180
      TabIndex        =   0
      Top             =   150
      Width           =   3675
      Begin VB.TextBox txt金额上限 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
         Height          =   270
         Left            =   2280
         TabIndex        =   10
         Top             =   1980
         Width           =   1035
      End
      Begin VB.ComboBox cmb人员 
         Height          =   300
         Left            =   1410
         TabIndex        =   2
         Text            =   "cmb人员"
         Top             =   360
         Width           =   1935
      End
      Begin VB.ComboBox cmb单据 
         Height          =   300
         Left            =   1410
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   810
         Width           =   1935
      End
      Begin VB.TextBox txtUD 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   1
         Left            =   2670
         TabIndex        =   6
         Text            =   "0"
         Top             =   1260
         Width           =   390
      End
      Begin VB.CheckBox chk修改 
         Caption         =   "准许操作他人单据(&T)"
         Height          =   210
         Left            =   270
         TabIndex        =   8
         Top             =   1665
         Value           =   1  'Checked
         Width           =   2040
      End
      Begin MSComCtl2.UpDown ud 
         Height          =   300
         Left            =   3060
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1260
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txtUD(1)"
         BuddyDispid     =   196614
         BuddyIndex      =   1
         OrigLeft        =   3105
         OrigTop         =   1260
         OrigRight       =   3345
         OrigBottom      =   1560
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "允许操作的金额上限(&M)"
         Height          =   180
         Index           =   3
         Left            =   285
         TabIndex        =   9
         Top             =   2025
         Width           =   1890
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "单据类型(&S)"
         Height          =   180
         Index           =   1
         Left            =   270
         TabIndex        =   3
         Top             =   870
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "操作员(&N)"
         Height          =   180
         Index           =   0
         Left            =   270
         TabIndex        =   1
         Top             =   420
         Width           =   810
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "允许操作单据的历史天数(&D)"
         Height          =   180
         Index           =   2
         Left            =   270
         TabIndex        =   5
         Top             =   1320
         Width           =   2250
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3990
      TabIndex        =   12
      Top             =   750
      Width           =   1100
   End
End
Attribute VB_Name = "frmBillPrivilege"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstr姓名 As String, mstr人员ID As String, mstr单据 As String
Private mlng单据 As Long, mlng天数 As Long, mbln修改他人 As Boolean
Private mdbl金额上限 As Double

Dim mblnChange As Boolean     '是否改变了
Dim mblnOk As Boolean
Dim mstrLike As String
Private Sub chk修改_Click()
    mblnChange = True
End Sub

Private Sub cmb单据_Click()
    mblnChange = True
    If Mid(cmb单据.Text, 1, 1) = 2 Or Mid(cmb单据.Text, 1, 1) = 4 Or Mid(cmb单据.Text, 1, 1) = 5 Or Mid(cmb单据.Text, 1, 1) = 9 Then
        Me.txt金额上限.Enabled = True
    Else
        Me.txt金额上限.Text = "0.00"
        Me.txt金额上限.Enabled = False
    End If
End Sub

Private Sub cmb人员_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strTmp As String
    Dim vRect As RECT, blnCancel As Boolean
    Dim intIdx As Integer
    
    On Local Error Resume Next
    
    If Visible Then mblnChange = True
    
    If cmb人员.ItemData(cmb人员.ListIndex) = -1 And Visible Then
        strSQL = "Select ID,简码,姓名 From 人员表 Where 撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null Order by 简码"

        vRect = GetControlRect(cmb人员.hwnd)
        Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "操作员", , , , , , True, vRect.Left, vRect.Top, cmb人员.Height, blnCancel, , True)
        If Not rsTmp Is Nothing Then
            intIdx = SeekCboIndex(cmb人员, rsTmp!ID)
            If intIdx <> -1 Then
                cmb人员.ListIndex = intIdx
            Else
                cmb人员.AddItem Nvl(rsTmp!简码) & "-" & rsTmp!姓名, cmb人员.ListCount - 1
                cmb人员.ItemData(cmb人员.NewIndex) = rsTmp!ID
                cmb人员.ListIndex = cmb人员.NewIndex
            End If
        Else
            If Not blnCancel Then
                MsgBox "没有操作员数据，请先到部门/人员管理中设置。", vbInformation, gstrSysName
            End If
            '恢复成现有的人员(不引发Click)
            intIdx = SeekCboIndex(cmb人员, cmb人员.Tag)
            Call zlControl.CboSetIndex(cmb人员.hwnd, intIdx)
        End If
    Else
        cmb人员.Tag = cmb人员.Text
    End If
End Sub

Private Function SeekCboIndex(objCbo As Object, varData As Variant) As Long
'功能：由ItemData或Text查找ComboBox的索引值
    Dim strType As String, i As Integer
    
    SeekCboIndex = -1
    
    strType = TypeName(varData)
    If strType = "Field" Then
        If IsType(varData.Type, adVarChar) Then strType = "String"
    End If
    
    If strType = "String" Then
        If varData <> "" Then
            '先精确查找
            For i = 0 To objCbo.ListCount - 1
                If objCbo.List(i) = varData Then
                    SeekCboIndex = i: Exit Function
                ElseIf NeedName(objCbo.List(i)) = varData And varData <> "" Then
                    SeekCboIndex = i: Exit Function
                End If
            Next
            '再模糊查找
            For i = 0 To objCbo.ListCount - 1
                If InStr(objCbo.List(i), varData) > 0 And varData <> "" Then
                    SeekCboIndex = i: Exit Function
                End If
            Next
        End If
    Else
        If varData <> 0 Then
            For i = 0 To objCbo.ListCount - 1
                If objCbo.ItemData(i) = varData Then
                    SeekCboIndex = i: Exit Function
                End If
            Next
        End If
    End If
End Function

Public Function NeedName(strList As String) As String
    If InStr(strList, Chr(13)) > 0 Then
        NeedName = LTrim(Mid(strList, InStr(strList, Chr(13)) + 1))
    ElseIf InStr(strList, "]") > 0 And InStr(strList, "-") = 0 Then
        NeedName = LTrim(Mid(strList, InStr(strList, "]") + 1))
    ElseIf InStr(strList, ")") > 0 And InStr(strList, "-") = 0 Then
        NeedName = LTrim(Mid(strList, InStr(strList, ")") + 1))
    Else
        NeedName = LTrim(Mid(strList, InStr(strList, "-") + 1))
    End If
End Function
Private Function IsType(ByVal varType As DataTypeEnum, ByVal varBase As DataTypeEnum) As Boolean
'功能：判断某个ADO字段数据类型是否与指定字段类型是同一类(如数字,日期,字符,二进制)
    Dim intA As Integer, intB As Integer
    
    Select Case varBase
        Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
            intA = -1
        Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
            intA = -2
        Case adDBTimeStamp, adDBTime, adDBDate, adDate
            intA = -3
        Case adBinary, adVarBinary, adLongVarBinary
            intA = -4
        Case Else
            intA = varBase
    End Select
    Select Case varType
        Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
            intB = -1
        Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
            intB = -2
        Case adDBTimeStamp, adDBTime, adDBDate, adDate
            intB = -3
        Case adBinary, adVarBinary, adLongVarBinary
            intB = -4
        Case Else
            intB = varType
    End Select
    
    IsType = intA = intB
End Function
Private Function GetControlRect(ByVal lngHwnd As Long) As RECT
'功能：获取指定控件在屏幕中的位置(Twip)
    Dim vRect As RECT
    Call GetWindowRect(lngHwnd, vRect)
    vRect.Left = vRect.Left * Screen.TwipsPerPixelX
    vRect.Right = vRect.Right * Screen.TwipsPerPixelX
    vRect.Top = vRect.Top * Screen.TwipsPerPixelY
    vRect.Bottom = vRect.Bottom * Screen.TwipsPerPixelY
    GetControlRect = vRect
End Function

Private Sub cmb人员_GotFocus()
    If cmb人员.Style = 0 Then
        Call zlControl.TxtSelAll(cmb人员)
    End If
End Sub

Private Sub cmb人员_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If cmb人员.Style = 2 And cmb人员.ListIndex <> -1 Then
            cmb人员.ListIndex = -1
        End If
    End If
End Sub


Private Sub cmb人员_KeyPress(KeyAscii As Integer)
'    Dim lngIdx As Long
'
'    lngIdx = MatchIndex(cmb人员.hwnd, KeyAscii)
'    If lngIdx <> -2 Then cmb人员.ListIndex = lngIdx
    
    Dim lngIdx As Long
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf KeyAscii >= 32 Then
        If Not cmb人员.Locked And cmb人员.Style = 2 Then
            lngIdx = zlControl.CboMatchIndex(cmb人员.hwnd, KeyAscii)
            If lngIdx = -1 And cmb人员.ListCount > 0 Then lngIdx = 0
            cmb人员.ListIndex = lngIdx
        End If
    End If
End Sub

Private Sub cmb人员_Validate(Cancel As Boolean)
    '功能：根据输入的内容,自动匹配执行科室
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, intIdx As Long, i As Long
    Dim strInput As String
    Dim vRect As RECT, blnCancel As Boolean
        
    If cmb人员.ListIndex <> -1 Then Exit Sub '已选中
    If cmb人员.Text = "" Then cmb人员.Tag = "": Exit Sub '无输入
    
    strInput = UCase(NeedName(cmb人员.Text))
    strSQL = "Select ID,简码,姓名 From 人员表 Where (撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null) Order by 简码"
    strSQL = Replace(UCase(strSQL), UCase("Order by"), " And (Upper(编号) Like [1] Or Upper(姓名) Like [2] Or Upper(简码) Like [2]) Order by")
    
    On Error GoTo errH
    vRect = GetControlRect(cmb人员.hwnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "操作员", False, "", "", False, False, _
        True, vRect.Left, vRect.Top, cmb人员.Height, blnCancel, False, True, strInput & "%", mstrLike & strInput & "%")
    If Not rsTmp Is Nothing Then
        intIdx = SeekCboIndex(cmb人员, rsTmp!ID)
        If intIdx <> -1 Then
            cmb人员.ListIndex = intIdx
        Else
            cmb人员.AddItem Nvl(rsTmp!简码) & "-" & Chr(13) & rsTmp!姓名, cmb人员.ListCount - 1
            cmb人员.ItemData(cmb人员.NewIndex) = rsTmp!ID
            cmb人员.ListIndex = cmb人员.NewIndex
        End If
    Else
        If Not blnCancel Then
            MsgBox "未找到对应的操作员。", vbInformation, gstrSysName
        End If
        Cancel = True: Exit Sub
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
    mstrLike = IIF(Val(zlDatabase.GetPara("输入匹配")) = 0, "%", "")
End Sub


Private Sub txtUD_Change(Index As Integer)
    If Index = 1 Then
        If Val(txtUD(Index).Text) > 100 Then txtUD(Index).Text = 100
        If Val(txtUD(Index).Text) < 0 Then txtUD(Index).Text = 0
    End If
End Sub

Private Sub txtUD_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txtUD(1))
End Sub


Private Sub txtUD_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 1 Then
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
        KeyAscii = 0
    End If
End Sub


Private Sub txt金额上限_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Trim(Me.txt金额上限.Text) = "" Then
            Me.txt金额上限.Text = "0.00"
        End If
        If Not IsNumeric(Me.txt金额上限.Text) Then
            MsgBox "输入的金额格式不正确。"
            Me.txt金额上限.SetFocus
            Exit Sub
        ElseIf Val(Me.txt金额上限.Text) > 10000000 Then
            MsgBox "金额不能超过7位整数。"
            Me.txt金额上限.SetFocus
            Exit Sub
        End If
        Me.txt金额上限.Text = Format(Me.txt金额上限.Text, "0.00")
    Else
        If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If

End Sub


Private Sub ud_Change()
    mblnChange = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认退出？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    If IsValid() = False Then Exit Sub
    
    mstr姓名 = Replace(GetTextFromCombo(cmb人员, True), "'", "")
    mstr人员ID = cmb人员.ItemData(cmb人员.ListIndex)
    
    mstr单据 = Mid(cmb单据.Text, 3)
    mlng单据 = Left(cmb单据.Text, 1)
    mlng天数 = Val(txtUD(1).Text)
    mbln修改他人 = (chk修改.Value = 1)
    mdbl金额上限 = Val(txt金额上限.Text)
    
    mblnOk = True
    mblnChange = False
    Unload Me
End Sub

Private Function IsValid() As Boolean
    If cmb人员.ListIndex < 0 Then
        MsgBox "请选择操作员。", vbInformation, gstrSysName
        cmb人员.SetFocus
        Exit Function
    End If
    If Trim(cmb人员.Text) = "" Then
        MsgBox "请选择操作员。", vbInformation, gstrSysName
        cmb人员.SetFocus
        Exit Function
    End If
    If cmb单据.Text = "" Then
        MsgBox "请选择操作的单据类型。", vbInformation, gstrSysName
        cmb单据.SetFocus
        Exit Function
    End If
    
    If Not IsNumeric(Me.txt金额上限.Text) Then
        MsgBox "输入的金额格式不正确。"
        Me.txt金额上限.SetFocus
        Exit Function
    ElseIf Val(Me.txt金额上限.Text) > 10000000 Then
        MsgBox "金额不能超过7位整数。"
        Me.txt金额上限.SetFocus
        Exit Function
    End If
    Me.txt金额上限.Text = Format(Me.txt金额上限.Text, "0.00")
    
'    If ud.Value = 0 And chk修改.Value = 1 Then
'        MsgBox "对单据的操作日期和操作人都没有限制，无需保存。", vbInformation, gstrSysName
'        chk修改.SetFocus
'        Exit Function
'    End If
    
    IsValid = True
End Function

Public Function 编辑权限(str姓名 As String, str人员ID As String, str单据 As String, lng单据 As Long, lng天数 As Long, bln修改他人 As Boolean, dbl金额上限 As Double _
                        , frmParent As Form) As Boolean
'功能：作为接口函数
    Dim rsTemp As New ADODB.Recordset, str简码 As String
    
    On Error GoTo ErrHandle
    gstrSQL = "Select ID,简码,姓名 From 人员表 Where 撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null Order by 简码"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    cmb人员.Clear
    Do Until rsTemp.EOF
        If IsNull(rsTemp("简码")) Then
            str简码 = zlStr.GetCodeByVB(rsTemp("姓名"))
        Else
            str简码 = rsTemp("简码")
        End If
        cmb人员.AddItem str简码 & "-" & rsTemp("姓名")
        cmb人员.ItemData(cmb人员.NewIndex) = rsTemp("ID")
        rsTemp.MoveNext
    Loop
    
    cmb单据.Clear
    
    If glngSys \ 100 = 8 Then
        '药店系统处理的单据有限制
        cmb单据.AddItem "2.收费单"
        cmb单据.AddItem "8.会员卡"
    Else
        cmb单据.AddItem "1.挂号单据"
        cmb单据.AddItem "2.收费单"
        cmb单据.AddItem "3.划价单"
        cmb单据.AddItem "4.门诊记帐"
        cmb单据.AddItem "5.住院记帐"
        cmb单据.AddItem "6.预交款"
        cmb单据.AddItem "7.结帐单据"
        cmb单据.AddItem "8.就诊卡"
        cmb单据.AddItem "9.处方"
    End If
    
    mstr人员ID = str人员ID
    SetComboByText cmb人员, str姓名, True
    '修改编号2779
    If cmb人员.List(cmb人员.ListCount - 1) = "" Then
        '删除那个空行
        cmb人员.RemoveItem cmb人员.ListCount - 1
    End If
    '----------------------------------
    If cmb人员.ListIndex >= 0 Then
        '如果是新增原列表中不存在的人员，则设置其ID
        If cmb人员.ItemData(cmb人员.ListIndex) = 0 Then cmb人员.ItemData(cmb人员.ListIndex) = Val(str人员ID)
    End If
    
    ud.Value = lng天数
    chk修改.Value = IIF(bln修改他人, 1, 0)
    txt金额上限.Text = Format(dbl金额上限, "0.00")
    SetComboByText cmb单据, lng单据, False, "."
    
    mblnChange = False
    mblnOk = False
    frmBillPrivilege.Show vbModal, frmParent
    
    
    If mblnOk = True Then
        str姓名 = mstr姓名
        str人员ID = mstr人员ID
        str单据 = mstr单据
        lng单据 = mlng单据
        lng天数 = mlng天数
        bln修改他人 = mbln修改他人
        dbl金额上限 = mdbl金额上限
    End If
    编辑权限 = mblnOk
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
