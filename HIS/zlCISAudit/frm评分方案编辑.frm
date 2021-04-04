VERSION 5.00
Begin VB.Form frm评分方案编辑 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "评分方案编辑"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5670
   Icon            =   "frm评分方案编辑.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin zl9CISAudit.tipPopup tipPopup1 
      Height          =   420
      Left            =   1890
      Top             =   2610
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   4635
      Top             =   945
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4080
      TabIndex        =   7
      Top             =   2970
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2775
      TabIndex        =   6
      Top             =   2970
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   465
      TabIndex        =   8
      Top             =   2970
      Width           =   1100
   End
   Begin VB.TextBox txt下值 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   1185
      MaxLength       =   10
      TabIndex        =   3
      Top             =   1458
      Width           =   2040
   End
   Begin VB.TextBox txt总分 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   1185
      MaxLength       =   10
      TabIndex        =   1
      Text            =   "100"
      Top             =   576
      Width           =   2040
   End
   Begin VB.CheckBox chk选用 
      Caption         =   "选用(&S)"
      Height          =   285
      Left            =   1185
      TabIndex        =   5
      Top             =   2340
      Width           =   1095
   End
   Begin VB.ComboBox cmb分制 
      Height          =   300
      Left            =   1185
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1899
      Width           =   2040
   End
   Begin VB.TextBox txt上值 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   1185
      MaxLength       =   10
      TabIndex        =   2
      Top             =   1017
      Width           =   2040
   End
   Begin VB.TextBox txt名称 
      Height          =   300
      Left            =   1185
      LinkTimeout     =   25
      MaxLength       =   25
      TabIndex        =   0
      Top             =   135
      Width           =   3870
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   11520
      Y1              =   2745
      Y2              =   2745
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   15
      X2              =   11520
      Y1              =   2775
      Y2              =   2775
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "分"
      Height          =   180
      Left            =   3300
      TabIndex        =   16
      Top             =   1518
      Width           =   180
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "分"
      Height          =   180
      Left            =   3300
      TabIndex        =   15
      Top             =   1077
      Width           =   180
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "分"
      Height          =   180
      Left            =   3300
      TabIndex        =   14
      Top             =   636
      Width           =   180
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "总分(&M)"
      Height          =   180
      Left            =   480
      TabIndex        =   13
      Top             =   636
      Width           =   630
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "分制(&F)"
      Height          =   180
      Left            =   480
      TabIndex        =   12
      Top             =   1959
      Width           =   630
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "下值(&B)"
      Height          =   180
      Left            =   480
      TabIndex        =   11
      Top             =   1518
      Width           =   630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "上值(&U)"
      Height          =   180
      Left            =   480
      TabIndex        =   10
      Top             =   1077
      Width           =   630
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "名称(&N)"
      Height          =   180
      Left            =   480
      TabIndex        =   9
      Top             =   195
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   900
      Left            =   4230
      Picture         =   "frm评分方案编辑.frx":000C
      Top             =   1665
      Width           =   900
   End
End
Attribute VB_Name = "frm评分方案编辑"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Private m_lngID                 As Long     '当前编辑的评分方案的ID号
Private m_strEditMode           As String   '窗口编辑模式：Add.新增 Mod.修改
Private m_blnModed              As Boolean
Private zlCheck                 As New clsCheck

Public Property Get Moded() As Boolean
   Moded = m_blnModed
End Property

Public Property Let Moded(ByVal blnModed As Boolean)
    m_blnModed = blnModed
End Property

'==============================================================================
'=功能： 公共接口函数：用于传入初始化参数:ID；ID的值为0表示添加标准；
'==============================================================================
Public Sub ShowForm(Optional ID As Long = 0)
    On Error GoTo ErrH
    
    m_lngID = ID          '为0表示新增
    m_blnModed = False
    '先填入选择下拉框中数据
    Call FillCmbs
    
    If ID <= 0 Then
        m_strEditMode = "Add"
        Me.Caption = "新增方案"
    Else
        m_strEditMode = "Mod"
        Me.Caption = "修改方案"
        FillInitData
    End If
    Me.Show 1
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 根据ID填入初始数据
'==============================================================================
Private Sub FillInitData()
    Dim rs      As ADODB.Recordset
    
    On Error GoTo ErrH
    
    gstrSQL = "select 名称,总分,上值,下值,分制,选用 from 病案评分方案 where ID = [1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, m_lngID)
    If Not rs.EOF Then
        txt名称 = IIf(IsNull(rs.Fields("名称")), "", rs.Fields("名称"))
        txt总分 = NVL(rs.Fields("总分"), 0)
        txt上值 = NVL(rs.Fields("上值"), 0)
        txt下值 = NVL(rs.Fields("下值"), 0)

        If rs.Fields("分制") = "加分制" Then
            cmb分制.ListIndex = 1
        Else
            cmb分制.ListIndex = 0
        End If
        If rs.Fields("选用") = 0 Then
            chk选用.Value = vbUnchecked
        Else
            chk选用.Value = vbChecked
        End If
        zlControl.TxtSelAll txt名称

    Else
        Unload Me
        MsgBox "初始化数据错误，没有发现该条评分方案！请重试。", vbOKOnly + vbInformation, "参数错误"
        Exit Sub
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 添加分制
'==============================================================================
Private Sub FillCmbs()
    On Error GoTo ErrH
    
    cmb分制.AddItem "扣分制"
    cmb分制.AddItem "加分制"
    cmb分制.ListIndex = 0
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 选用提示
'==============================================================================
Private Sub chk选用_GotFocus()
    On Error GoTo ErrH
    ShowTips chk选用, "每个类型下最多只能有一个方案被选用。", "选用"
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 选用回车相当于确定
'==============================================================================
Private Sub chk选用_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrH
    If KeyAscii = vbKeyReturn Then
        Call CmdOK_Click
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 分制提示
'==============================================================================
Private Sub cmb分制_GotFocus()
    On Error GoTo ErrH
    ShowTips cmb分制, "可以选择“加分制”与“扣分制”两类分制，默认为“扣分制”。", "方案分制"
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 取消编辑
'==============================================================================
Private Sub CmdCancel_Click()
    On Error GoTo ErrH
    Moded = False
    Unload Me
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 点击帮助
'==============================================================================
Private Sub cmdHelp_Click()
    On Error GoTo ErrH
    ShowHelp App.ProductName, Me.Hwnd, Me.Name, 3
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：确定保存数据
'==============================================================================
Private Sub CmdOK_Click()
    Dim strT As String
    
    On Error GoTo ErrH
    
    If m_strEditMode = "Add" Then
        strT = "ZL_病案评分方案_Insert"
        gstrSQL = strT & _
                "(" & zlDatabase.GetNextId("病案评分方案") & ",'" & txt名称 & "'," & CStr(Val(txt总分)) & "," & CStr(Val(txt上值)) & "," & CStr(Val(txt下值)) & _
                ",'住院','" & cmb分制.Text & "'," & CStr(IIf(chk选用.Value = vbChecked, 1, 0)) & _
                ",NULL,NULL" & _
                ")"
    Else
        strT = "ZL_病案评分方案_Update"
        gstrSQL = strT & _
                "(" & CStr(m_lngID) & ",'" & txt名称 & "'," & CStr(Val(txt总分)) & "," & CStr(Val(txt上值)) & "," & CStr(Val(txt下值)) & _
                ",'住院','" & cmb分制.Text & "'," & CStr(IIf(chk选用.Value = vbChecked, 1, 0)) & _
                ",NULL,NULL" & _
                ")"
    End If
    '检查分数合法性
    If IsValid() = False Then Exit Sub
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Moded = True
    MsgBox "评分方案保存成功！", vbOKOnly + vbInformation, gstrSysName
    Unload Me
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：分析输入项目的内容是否有效
'=返回：有效返回True,否则为False
'==============================================================================
Private Function IsValid() As Boolean
    On Error GoTo ErrH
    '必填字段检查
    IsValid = False
    '调用StrIsValid函数来确保字符串格式正确，注意：长度使用的是lenB值（对应数据表定义中的值）
    If Len(Trim(txt名称)) = 0 Then
        MsgBox "请输入方案名称！", vbInformation, gstrSysName
        zlControl.TxtSelAll txt名称: txt名称.SetFocus
        Exit Function
    End If
    If zlCommFun.StrIsValid(txt名称.Text, txt名称.MaxLength * 2) = False Then
        MsgBox "输入的名称太长，请重新输入！", vbInformation, gstrSysName
        zlControl.TxtSelAll txt名称: txt名称.SetFocus
        Exit Function
    End If
    If Len(Trim(txt总分)) = 0 Then
        MsgBox "请输入方案总分！", vbInformation, gstrSysName
        zlControl.TxtSelAll txt总分: txt总分.SetFocus
        Exit Function
    End If
    If Len(Trim(txt上值)) = 0 Then
        MsgBox "请输入方案中甲级病案的分数线！", vbInformation, gstrSysName
        zlControl.TxtSelAll txt上值: txt上值.SetFocus
        Exit Function
    End If
    If Len(Trim(txt下值)) = 0 Then
        MsgBox "请输入方案中乙级病案的分数线！", vbInformation, gstrSysName
        zlControl.TxtSelAll txt下值: txt下值.SetFocus
        Exit Function
    End If
    If Len(Trim(txt总分)) > 0 Then
        If Not IsNumeric(txt总分) Then
            MsgBox "请输入方案总分！", vbInformation, gstrSysName
            zlControl.TxtSelAll txt总分: txt总分.SetFocus
            Exit Function
        End If
        If Val(txt总分.Text) > 9999# Then
            MsgBox "方案总分中输入的数据太大！", vbInformation, gstrSysName
            zlControl.TxtSelAll txt总分: txt总分.SetFocus
            Exit Function
        End If
    End If
    If Len(Trim(txt上值)) > 0 Then
        If Not IsNumeric(txt上值) Then
            MsgBox "请输入甲级病案的分数线！", vbInformation, gstrSysName
            zlControl.TxtSelAll txt上值: txt上值.SetFocus
            Exit Function
        End If
        If Val(txt上值.Text) > 9999# Or Val(txt上值.Text) > Val(txt总分.Text) Then
            MsgBox "甲级病案的分数线中输入的数据太大！", vbInformation, gstrSysName
            zlControl.TxtSelAll txt上值: txt上值.SetFocus
            Exit Function
        End If
    End If
    If Len(Trim(txt下值)) > 0 Then
        If Not IsNumeric(txt下值) Then
            MsgBox "请输入乙级病案的分数线！", vbInformation, gstrSysName
            zlControl.TxtSelAll txt下值: txt下值.SetFocus
            Exit Function
        End If
        If Val(txt下值.Text) > 9999# Or Val(txt下值.Text) > Val(txt总分.Text) Or Val(txt下值.Text) > Val(txt上值.Text) Then
            MsgBox "乙级病案的分数线中输入的数据太大！", vbInformation, gstrSysName
            zlControl.TxtSelAll txt下值: txt下值.SetFocus
            Exit Function
        End If
        If Val(txt下值.Text) < 0 Then
            MsgBox "乙级病案的分数线中输入的数据太小！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    IsValid = True
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'==============================================================================
'=功能：页面初始化
'==============================================================================
Private Sub Form_Initialize()
    On Error GoTo ErrH
    Call InitCommonControls
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：页面初始化
'==============================================================================
Private Sub Form_Load()
    On Error GoTo ErrH
    zlCheck.Sys_System Me
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Sub

'==============================================================================
'=功能：名称对 ' | 录入的控制
'==============================================================================
Private Sub txt名称_Change()
    On Error GoTo ErrH
    If InStr(txt名称, "'") <> 0 Then txt名称 = Replace(txt名称, "'", "")
    If InStr(txt名称, "|") <> 0 Then txt名称 = Replace(txt名称, "|", "")
    txt名称.SelStart = Len(txt名称)
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：名称得到焦点提示
'==============================================================================
Private Sub txt名称_GotFocus()
    On Error GoTo ErrH
    zlControl.TxtSelAll txt名称
    Call zlCommFun.OpenIme(True)
    ShowTips txt名称, "输入方案名称，长度在25个字符内。", "方案名称"
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：名称失去焦点检测
'==============================================================================
Private Sub txt名称_LostFocus()
    On Error GoTo ErrH
    Call zlCommFun.OpenIme(False)
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：名称取消提示
'==============================================================================
Private Sub txt名称_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo ErrH
    tipPopup1.Hide
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：上值得到焦点提示
'==============================================================================
Private Sub txt上值_GotFocus()
    On Error GoTo ErrH
    zlControl.TxtSelAll txt上值
    ShowTips txt上值, "甲级病案分数线", "方案上值", 5000
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：上值按键F1提示
'==============================================================================
Private Sub txt上值_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrH
    If KeyCode = vbKeyF1 Then
        ShowTips txt上值, "输入方案上值。大于等于上值的病案为甲级，介于下值与上值间的病案为乙级，低于下值的病案为丙级。", "方案上值"
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 下值得到焦点提示
'==============================================================================
Private Sub txt下值_GotFocus()
    On Error GoTo ErrH
    zlControl.TxtSelAll txt下值
    ShowTips txt下值, "乙级病案分数线", "方案下值", 5000
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：下值按键F1提示
'==============================================================================
Private Sub txt下值_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrH
    If KeyCode = vbKeyF1 Then
        '显示Tips
        ShowTips txt下值, "输入方案下值。大于等于上值的病案为甲级，介于下值与上值间的病案为乙级，低于下值的病案为丙级。", "方案下值"
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 总分得到焦点提示
'==============================================================================
Private Sub txt总分_GotFocus()
    On Error GoTo ErrH
    zlControl.TxtSelAll txt总分
    ShowTips txt总分, "输入方案标准总分，默认为100分。", "方案总分"
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 提示
'==============================================================================
Private Sub ShowTips(ctl As Control, str内容 As String, Optional str标题 As String = "提示信息", Optional lng时间 As Long = 2500, Optional 错误提示 As Boolean = False)
    Dim X       As Single
    Dim Y       As Single
    On Error GoTo ErrH
    
    X = (ctl.Left + ctl.Width / 2) / Screen.TwipsPerPixelX
    Y = (ctl.Top + ctl.Height) / Screen.TwipsPerPixelY
    If Len(str内容) > 0 Then
        tipPopup1.Hide
        tipPopup1.StandardIcon = IDI_INFORMATION
        tipPopup1.ShowCloseButton = True
        
        tipPopup1.TimeOut = lng时间
        tipPopup1.Title = str标题
        tipPopup1.Text = str内容
        tipPopup1.Show Me.Hwnd, X, Y
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
