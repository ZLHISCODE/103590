VERSION 5.00
Begin VB.Form frmModiPatiBaseInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病人基本信息调整"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4380
   Icon            =   "frmModiPatiBaseInfo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   4380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox cboAge 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3135
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1215
      Width           =   705
   End
   Begin VB.TextBox txtAge 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   2  'OFF
      Left            =   2115
      TabIndex        =   5
      Top             =   1215
      Width           =   1020
   End
   Begin VB.ComboBox cboSex 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmModiPatiBaseInfo.frx":030A
      Left            =   2115
      List            =   "frmModiPatiBaseInfo.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   690
      Width           =   1725
   End
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2115
      MaxLength       =   64
      TabIndex        =   1
      Top             =   210
      Width           =   1725
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   1215
      TabIndex        =   7
      Top             =   1995
      Width           =   1450
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   2685
      TabIndex        =   8
      Top             =   1995
      Width           =   1450
   End
   Begin VB.Frame Frame2 
      Height          =   120
      Left            =   30
      TabIndex        =   9
      Top             =   1710
      Width           =   5100
   End
   Begin VB.Image imgFlag 
      Height          =   480
      Left            =   495
      Picture         =   "frmModiPatiBaseInfo.frx":030E
      Top             =   600
      Width           =   480
   End
   Begin VB.Label lblAge 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "年龄"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1500
      TabIndex        =   4
      Top             =   1275
      Width           =   480
   End
   Begin VB.Label lblSex 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "性别"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1485
      TabIndex        =   2
      Top             =   750
      Width           =   480
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "姓名"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1530
      TabIndex        =   0
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmModiPatiBaseInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnOK As Boolean
Private mlng病人ID As Long
Private mstr就诊ID As String
Private mstr模块 As String
Private mint场合 As Integer

Public Function ShowMe(ByVal lng病人ID As Long, ByVal str就诊ID As String, ByVal str模块 As String, ByVal int场合 As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:程序入口
    '入参:lng病人ID-病人ID
    '     str就诊ID=批量修改为空，门诊病人为挂号ID，住院病人为主页ID，外来病人根据业务自行决定，如：医嘱ID，体检病人为任务单号
    '     str模块=调用该功能的模块描述，如"门诊挂号"，"检查报到"。
    '     int场合=0-批量,1-门诊,2-住院,3-外来病人,4-体检病人
    '出参:
    '返回:
    '编制:刘鹏飞
    '日期:2013-10-22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlng病人ID = lng病人ID
    mstr就诊ID = str就诊ID
    mstr模块 = str模块
    mint场合 = int场合
    
    mblnOK = False
    '获取病人基本信息
    If Not LoadPatiBaseInfo Then ShowMe = False: Exit Function
    
    Me.Show 1
    ShowMe = mblnOK
End Function

Private Sub InitDicts()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHand
    
    txtName.Text = ""
    txtName.MaxLength = GetColumnLength("病人信息", "姓名")
    txtAge.Text = ""
    cboAge.Clear
    cboAge.AddItem "岁"
    cboAge.AddItem "月"
    cboAge.AddItem "天"
    cboAge.ListIndex = 0
    txtAge.MaxLength = GetColumnLength("病人信息", "年龄")
    
    cboSex.Clear
    
    strSQL = "Select 编码,名称,简码,Nvl(缺省标志,0) as 缺省 From 性别 Order by 编码"
    Call gobjComLib.zlDatabase.OpenRecordset(rsTmp, strSQL, "性别")
    Do While Not rsTmp.EOF
        cboSex.AddItem rsTmp!编码 & "-" & rsTmp!名称
        If rsTmp!缺省 = 1 Then
            cboSex.ListIndex = cboSex.NewIndex
            cboSex.ItemData(cboSex.NewIndex) = 1
        End If
    rsTmp.MoveNext
    Loop
    
    Exit Sub
errHand:
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Sub

Private Function LoadPatiBaseInfo() As Boolean
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim lngIndex As Long
    
    On Error GoTo errHand
    
    If mint场合 = 1 Then '门诊病人
        strSQL = "Select 姓名,性别,年龄 from 病人挂号记录 where 病人ID=[1] And ID=[2]"
    ElseIf mint场合 = 2 Then '住院病人
        strSQL = " Select Nvl(a.姓名, b.姓名) 姓名, Nvl(a.性别, b.性别) 性别, a.年龄" & vbNewLine & _
                " From 病案主页 a, 病人信息 b" & vbNewLine & _
                " Where a.病人id = b.病人id And a.病人id = [1] And a.主页id = [2]"
    Else
        strSQL = "Select 姓名,性别,年龄 From 病人信息 Where 病人ID=[1]"
    End If
    
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "提取病人基本信息", mlng病人ID, Val(mstr就诊ID))
    
    If Not rsTmp.EOF Then
        If mint场合 = 0 Then
            Me.Caption = "病人历次基本信息调整"
        Else
            Me.Caption = "病人基本信息调整"
        End If
        '基本信息初始化
        Call InitDicts
        
        txtName.Text = gobjComLib.zlCommFun.NVL(rsTmp!姓名)
        txtName.Tag = txtName.Text
        cboSex.Tag = gobjComLib.zlCommFun.NVL(rsTmp!性别)
        lngIndex = GetCboIndex(cboSex, gobjComLib.zlCommFun.NVL(rsTmp!性别))
        If lngIndex <> -1 Then cboSex.ListIndex = lngIndex
        Call LoadOldData("" & rsTmp!年龄, txtAge, cboAge)
        txtAge.Tag = gobjComLib.zlCommFun.NVL(rsTmp!年龄)
    Else
        MsgBox "获取病人基本信息失败,请您确认要进行信息调整的病人！", vbInformation, gstrSysName
        Exit Function
    End If
    
    LoadPatiBaseInfo = True
    Exit Function
errHand:
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
    gobjComLib.SaveErrLog
End Function

Private Sub LoadOldData(ByVal strOld As String, ByRef txtAge As TextBox, ByRef cboAge As ComboBox)
'功能:将数据库中保存的年龄按规范的格式加载到界面,不规范的原样显示
    Dim strTmp As String, lngIdx As Long
    
    If Trim(strOld) = "" Then Exit Sub
    
    lngIdx = -1
    strTmp = strOld
    If InStr(strOld, "岁") > 0 Then
        If InStr(strOld, "岁") = Len(strOld) Then
            strTmp = Mid(strOld, 1, InStr(strOld, "岁") - 1)
            lngIdx = 0
        End If
    ElseIf InStr(strOld, "月") > 0 Then
        If InStr(strOld, "月") = Len(strOld) Then
            strTmp = Mid(strOld, 1, InStr(strOld, "月") - 1)
            lngIdx = 1
        End If
    ElseIf InStr(strOld, "天") > 0 Then
        If InStr(strOld, "天") = Len(strOld) Then
            strTmp = Mid(strOld, 1, InStr(strOld, "天") - 1)
            lngIdx = 2
        End If
    ElseIf IsNumeric(strOld) Then
        lngIdx = 0
    End If
    txtAge.Text = strTmp
    If cboAge.ListCount > 0 Then Call gobjComLib.zlControl.CboSetIndex(cboAge.hwnd, lngIdx)
    If lngIdx = -1 Then
        cboAge.Visible = False
    Else
        If cboAge.Visible = False Then cboAge.Visible = True
    End If
End Sub

Private Function CheckOldData(ByRef txtAge As TextBox, ByRef cboAge As ComboBox) As Boolean
'功能：检查年龄输入值的有效性
'返回：
    If Not IsNumeric(txtAge.Text) Then CheckOldData = True: Exit Function
    
    Select Case cboAge.Text
        Case "岁"
            If Val(txtAge.Text) > 200 Then
                MsgBox "年龄不能大于200岁!", vbInformation, gstrSysName
                If txtAge.Enabled And txtAge.Visible Then txtAge.SetFocus
                CheckOldData = False: Exit Function
            End If
        Case "月"
            If Val(txtAge.Text) > 2400 Then
                MsgBox "年龄不能大于2400月!", vbInformation, gstrSysName
                If txtAge.Enabled And txtAge.Visible Then txtAge.SetFocus
                CheckOldData = False: Exit Function
            End If
        Case "天"
            If Val(txtAge.Text) > 73000 Then
                MsgBox "年龄不能大于73000天!", vbInformation, gstrSysName
                If txtAge.Enabled And txtAge.Visible Then txtAge.SetFocus
                CheckOldData = False: Exit Function
            End If
    End Select
    CheckOldData = True
End Function

Private Sub CheckInputLen(txt As Object, KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: Exit Sub
    If KeyAscii < 32 And KeyAscii >= 0 Then Exit Sub
    If txt.MaxLength = 0 Then Exit Sub
    If gobjComLib.zlCommFun.ActualLen(txt.Text & Chr(KeyAscii)) > txt.MaxLength Then KeyAscii = 0
End Sub

Private Function CheckTextLength(strName As String, txtObj As TextBox) As Boolean
'功能:检查并提示文本框输入长度是否超限

    CheckTextLength = True
    If gobjComLib.zlCommFun.ActualLen(txtObj.Text) > txtObj.MaxLength Then
        MsgBox strName & "输入过长，只允许输入 " & txtObj.MaxLength & " 个字符或 " & txtObj.MaxLength \ 2 & " 个汉字。", vbInformation, gstrSysName
        If txtObj.Enabled And txtObj.Visible Then txtObj.SetFocus
        CheckTextLength = False
    End If
End Function

Private Function GetColumnLength(strTable As String, strColumn As String) As Long
'功能：获取指定表中指定字段的长度
    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "Select Nvl(Data_Precision, Data_Length) collen From All_Tab_Columns Where Table_Name = [1] And Column_Name = [2]"
    On Error GoTo errH
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, App.ProductName, strTable, strColumn)
    GetColumnLength = Val("" & rsTmp!collen)
    
    Exit Function
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Private Function GetCboIndex(cbo As ComboBox, strFind As String, _
    Optional blnKeep As Boolean, _
    Optional blnLike As Boolean, Optional strSplit As String = "-") As Long
'功能：由字符串在ComboBox中查找索引
    Dim i As Long
    If strFind = "" Then GetCboIndex = -1: Exit Function
    '先精确查找
    For i = 0 To cbo.ListCount - 1
        If InStr(cbo.List(i), strSplit) > 0 Then
            If NeedName(cbo.List(i)) = strFind Then GetCboIndex = i: Exit Function
        Else
            If cbo.List(i) = strFind Then GetCboIndex = i: Exit Function
        End If
    Next
    '最后模糊查找
    If blnLike Then
        For i = 0 To cbo.ListCount - 1
            If InStr(cbo.List(i), strFind) > 0 Then GetCboIndex = i: Exit Function
        Next
    End If
    If Not blnKeep Then GetCboIndex = -1
End Function

Private Function NeedName(strList As String) As String
    If InStr(strList, Chr(&HA)) > 0 Then
        NeedName = Trim(Mid(strList, InStr(strList, Chr(&HA)) + 1))
    Else
        NeedName = Trim(Mid(strList, InStr(strList, "-") + 1))
    End If
    If InStr(NeedName, Chr(&HD)) > 0 Then
        NeedName = Replace(NeedName, Chr(&HD), "")
    End If
End Function

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
'功能：完成数据校验和保存
    Dim strSQL As String
    Dim str年龄 As String
    
    '第一步：数据合法性校验
    If Trim(txtName.Text) = "" Then
        MsgBox "必须输入病人姓名！", vbInformation, gstrSysName
        txtName.SetFocus: Exit Sub
    End If
    If cboSex.ListIndex = -1 Then
        MsgBox "必须确定病人性别！", vbInformation, gstrSysName
        cboSex.SetFocus: Exit Sub
    End If
    If Trim(txtAge.Text) = "" Then
        MsgBox "必须输入病人年龄！", vbInformation, gstrSysName
        txtAge.SetFocus: Exit Sub
    End If
    
    
    If Not CheckTextLength("姓名", txtName) Then Exit Sub
    If Not CheckTextLength("年龄", txtAge) Then Exit Sub
    If Not CheckOldData(txtAge, cboAge) Then Exit Sub
    
    str年龄 = Trim(txtAge.Text)
    If IsNumeric(str年龄) Then str年龄 = str年龄 & cboAge.Text
    
    '第二步：数据保存
    On Error GoTo errHand
    strSQL = "Zl_病人信息_基本信息调整("
'   病人id_In 病人信息变动.病人id%Type,
    strSQL = strSQL & "" & mlng病人ID & ","
'   就诊id_In Number := Null,
    strSQL = strSQL & "'" & mstr就诊ID & "',"
'   模块_In   病人信息变动.变动模块%Type,
    strSQL = strSQL & "'" & mstr模块 & "',"
'   姓名_In   病人信息.姓名%Type,
    strSQL = strSQL & "'" & Trim(txtName.Text) & "',"
'   性别_In   病人信息.性别%Type,
    strSQL = strSQL & "'" & Split(cboSex.Text, "-")(1) & "',"
'   年龄_In   病人信息.年龄%Type
    strSQL = strSQL & "'" & str年龄 & "',"
'   场合_In   number(1)
    strSQL = strSQL & "" & mint场合 & ")"
    
    Call gobjComLib.zlDatabase.ExecuteProcedure(strSQL, "Zl_病人信息_基本信息调整")
    
    mblnOK = True
    Unload Me
    Exit Sub
errHand:
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
    gobjComLib.SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyReturn Then
        If ActiveControl.Name <> txtName.Name And ActiveControl.Name <> txtAge.Name Then
            Call gobjComLib.zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtAge_GotFocus()
    Call gobjComLib.zlCommFun.OpenIme
    gobjComLib.zlControl.TxtSelAll txtAge
End Sub

Private Sub txtAge_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cboAge.Visible = False And IsNumeric(txtAge.Text) Then
            Call txtAge_Validate(False)
            Call cboAge.SetFocus
        Else
            Call gobjComLib.zlCommFun.PressKey(vbKeyTab)
        End If
        If Not IsNumeric(txtAge.Text) Then Call gobjComLib.zlCommFun.PressKey(vbKeyTab)
    Else
        If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Chr(KeyAscii))) > 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txtAge_Validate(Cancel As Boolean)
    If Not IsNumeric(txtAge.Text) And Trim(txtAge.Text) <> "" Then
        cboAge.ListIndex = -1: cboAge.Visible = False
    ElseIf cboAge.Visible = False Then
        cboAge.ListIndex = 0: cboAge.Visible = True
    End If
End Sub

Private Sub txtName_GotFocus()
    gobjComLib.zlControl.TxtSelAll txtName
    Call gobjComLib.zlCommFun.OpenIme(True)
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then
        If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then
            KeyAscii = 0
        Else
            Call CheckInputLen(txtName, KeyAscii)
        End If
    Else
        If Trim(txtName.Text) = "" Then
            Exit Sub
        Else
            gobjComLib.zlCommFun.PressKey (vbKeyTab)
        End If
    End If
End Sub

Private Sub txtName_LostFocus()
    Call gobjComLib.zlCommFun.OpenIme
    txtName.Text = Trim(txtName.Text)
End Sub
