VERSION 5.00
Begin VB.Form frm审查方案编辑 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "评分方案编辑"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5670
   Icon            =   "frm审查方案编辑.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txt说明 
      Height          =   900
      Left            =   1185
      MaxLength       =   200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1395
      Width           =   3885
   End
   Begin VB.TextBox txt分段线 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   1185
      MaxLength       =   10
      TabIndex        =   2
      ToolTipText     =   "如方案总分>=90为合格"
      Top             =   975
      Width           =   2040
   End
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
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4080
      TabIndex        =   6
      Top             =   2970
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2775
      TabIndex        =   5
      Top             =   2970
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   465
      TabIndex        =   7
      Top             =   2970
      Width           =   1100
   End
   Begin VB.TextBox txt总分 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   1185
      MaxLength       =   10
      TabIndex        =   1
      Text            =   "100"
      ToolTipText     =   "方案扣分时最高分"
      Top             =   576
      Width           =   2040
   End
   Begin VB.CheckBox chk选用 
      Caption         =   "选用(&S)"
      Height          =   285
      Left            =   1185
      TabIndex        =   4
      Top             =   2340
      Width           =   1095
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "说明(&B)"
      Height          =   180
      Left            =   495
      TabIndex        =   13
      Top             =   1410
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "分"
      Height          =   180
      Left            =   3300
      TabIndex        =   12
      Top             =   1020
      Width           =   180
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
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "分"
      Height          =   180
      Left            =   3300
      TabIndex        =   11
      Top             =   636
      Width           =   180
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "总分(&M)"
      Height          =   180
      Left            =   480
      TabIndex        =   10
      Top             =   636
      Width           =   630
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "分段线(&F)"
      Height          =   180
      Left            =   300
      TabIndex        =   9
      Top             =   1020
      Width           =   810
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "名称(&N)"
      Height          =   180
      Left            =   480
      TabIndex        =   8
      Top             =   195
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   900
      Left            =   15
      Picture         =   "frm审查方案编辑.frx":000C
      Top             =   1740
      Width           =   900
   End
End
Attribute VB_Name = "frm审查方案编辑"
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
    
    gstrSQL = "Select 名称,总分,分段线,启用时间,停用时间,说明 From 病案审查方案 where ID = [1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, m_lngID)
    If Not rs.EOF Then
        txt名称 = IIf(IsNull(rs.Fields("名称")), "", rs.Fields("名称"))
        txt总分 = NVL(rs.Fields("总分"), 0)
        txt分段线 = NVL(rs.Fields("分段线"), 0)
        txt说明 = NVL(rs.Fields("说明"))
         
        If NVL(rs.Fields("启用时间")) = "" Then
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
'=功能： 选用提示
'==============================================================================
Private Sub chk选用_GotFocus()
    On Error GoTo ErrH
    ShowTips chk选用, "最多只能有一个方案被选用。", "选用"
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
    ShowHelp App.ProductName, Me.hWnd, Me.Name, 3
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
        strT = "ZL_病案审查方案_Insert"
        gstrSQL = strT & _
                "(" & zlDatabase.GetNextId("病案审查方案") & ",'" & txt名称 & "'," & txt总分.Text & "," & txt分段线.Text & "," & CStr(IIf(chk选用.Value = vbChecked, 1, 0)) & _
                ",'" & txt说明.Text & "'" & _
                ")"
    Else
        strT = "ZL_病案审查方案_Update"
        gstrSQL = strT & _
                       "(" & m_lngID & ",'" & txt名称 & "'," & txt总分.Text & "," & txt分段线.Text & "," & CStr(IIf(chk选用.Value = vbChecked, 1, 0)) & _
                       ",'" & txt说明.Text & "'" & _
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

    If Len(Trim(txt总分)) > 0 Then
        If Not IsNumeric(txt总分) Then
            MsgBox "请输入有效的方案总分值！", vbInformation, gstrSysName
            zlControl.TxtSelAll txt总分: txt总分.SetFocus
            Exit Function
        End If
        If Val(txt总分.Text) > 9999# Then
            MsgBox "方案总分中输入的数据太大！", vbInformation, gstrSysName
            zlControl.TxtSelAll txt总分: txt总分.SetFocus
            Exit Function
        End If
    End If

    If Len(Trim(txt分段线)) > 0 Then
        If Not IsNumeric(txt分段线) Then
            MsgBox "请输入正确的分段线数值！", vbInformation, gstrSysName
            zlControl.TxtSelAll txt分段线: txt分段线.SetFocus
            Exit Function
        End If
        If Val(txt分段线.Text) > 9999# Or Val(txt分段线.Text) > Val(txt总分.Text) Then
            MsgBox "分段线中输入的数据太大！", vbInformation, gstrSysName
            zlControl.TxtSelAll txt分段线: txt分段线.SetFocus
            Exit Function
        End If
        If Val(txt分段线.Text) < 0 Then
            MsgBox "分段线中输入的数据太小！", vbInformation, gstrSysName
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
Private Sub txt分段线_GotFocus()
    On Error GoTo ErrH
    zlControl.TxtSelAll txt分段线
    ShowTips txt分段线, "甲级病案分段线", "分段线值", 5000
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：上值按键F1提示
'==============================================================================
Private Sub txt分段线_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrH
    If KeyCode = vbKeyF1 Then
        ShowTips txt分段线, "如:分段线>=90为合格,否则为不合格，比如入院后24小时未书写入院病历。"
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
        tipPopup1.Show Me.hWnd, X, Y
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
