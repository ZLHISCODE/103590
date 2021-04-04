VERSION 5.00
Begin VB.Form frmInput 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "中联软件"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5130
   ControlBox      =   0   'False
   Icon            =   "frmInput.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3480
      TabIndex        =   2
      Top             =   1530
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2220
      TabIndex        =   1
      Top             =   1530
      Width           =   1100
   End
   Begin VB.TextBox txtInput 
      Height          =   300
      Left            =   1980
      MaxLength       =   18
      TabIndex        =   0
      Top             =   795
      Width           =   2025
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      X1              =   0
      X2              =   6000
      Y1              =   1335
      Y2              =   1335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      X1              =   0
      X2              =   6000
      Y1              =   1350
      Y2              =   1350
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   240
      Picture         =   "frmInput.frx":000C
      Top             =   90
      Width           =   480
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "在留观病人 钟无艳 转为住院病人之前，请先为该病人确定一个住院号。"
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   975
      TabIndex        =   4
      Top             =   165
      Width           =   3825
   End
   Begin VB.Label lblInput 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "住院号"
      Height          =   180
      Left            =   1380
      TabIndex        =   3
      Top             =   855
      Width           =   540
   End
End
Attribute VB_Name = "frmInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Private mblnIme As Boolean
Private mbytType As Byte
Private mblnAllowNull As Boolean
Private mblnUcase As Boolean
Private mstrInput As String
Private mblnOK As Boolean
Private mblnPassInput As Boolean
Private mblnAffirmPass As Boolean
Private mobjKeyboard As Object
Public Function InputVal(ByVal frmParent As Object, ByVal strItem As String, _
    ByVal strNote As String, ByRef strInput As String, ByVal BytType As Byte, _
    Optional ByVal intMax As Integer, Optional ByVal blnAllowNull As Boolean, _
    Optional ByVal blnUCase As Boolean, Optional ByVal blnIme As Boolean, _
    Optional ByVal PassChar As String, Optional blnPassInput As Boolean = False, _
    Optional blnAffirmPass As Boolean = False) As Boolean
'功能：显示一个输入框,类似VB的InputBox函数
'参数：frmParent=父窗体
'      strItem=要输入的项目名称
'      strNote=输入框中的提示。
'      strInput=入/出参数:初始显示及返回的值
'      bytType=数据类型:0-字符串,1-数字,2-日期,3-字母和数字
'      intMax=输入长度限制
'      blnAllowNull=是否允许输入空
'      blnUCase=输入是否全部大写
'      blnIme=是否自动打开输入法
'      blnPassInput-是否密码输入
'      blnAffirmPass-是否输入的确认密码
'返回：输入确定返回True,取消返回Fasle
    mblnPassInput = blnPassInput: mblnAffirmPass = blnAffirmPass
    Load Me
    Me.Caption = gstrSysName
    Me.lblNote.Caption = strNote
    Me.lblInput.Caption = strItem
    Me.txtInput.Text = strInput
    Me.txtInput.MaxLength = intMax
    Me.txtInput.PasswordChar = PassChar
    
    mblnIme = blnIme
    mbytType = BytType
    mblnUcase = blnUCase
    mblnAllowNull = blnAllowNull
    
    Me.Show 1, frmParent
    
    strInput = mstrInput
    InputVal = mblnOK
End Function

Private Sub cmdCancel_Click()
    mstrInput = ""
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Trim(txtInput.Text) = "" And Not mblnAllowNull Then
        MsgBox "必须输入" & lblInput.Caption & "！", vbInformation, gstrSysName
        txtInput.SetFocus: Exit Sub
    End If
    If txtInput.MaxLength <> 0 Then
        If zlCommFun.ActualLen(txtInput.Text) > txtInput.MaxLength Then
            MsgBox "最多允许输入 " & txtInput.MaxLength & " 个字符或 " & txtInput.MaxLength \ 2 & " 个汉字！", vbInformation, gstrSysName
            txtInput.SetFocus: Exit Sub
        End If
    End If
    If mbytType = 1 Then
        If Not IsNumeric(txtInput.Text) Then
            MsgBox "输入内容不是合法的数字！", vbInformation, gstrSysName
            txtInput.SetFocus: Exit Sub
        End If
    ElseIf mbytType = 2 Then
        If Not IsNumeric(txtInput.Text) Then
            MsgBox "输入内容不是合法的日期！", vbInformation, gstrSysName
            txtInput.SetFocus: Exit Sub
        End If
    End If
    
    '留观病人转住院病人时，frmManageCourse
    If lblInput.Caption = "住院号" Then
        If ExistInPatiNO(txtInput.Text) Then
            MsgBox "发现当前住院号已经被其它病人使用,系统将自动更换一个不重复的住院号！", vbInformation, gstrSysName
            txtInput.Text = zlDatabase.GetNextNo(2)
            txtInput.SetFocus: Exit Sub
        End If
    End If
    
    mstrInput = txtInput.Text
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_Load()
    If mblnPassInput Then Call CreateObjectKeyboard
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set mobjKeyboard = Nothing
End Sub
Private Sub txtInput_GotFocus()
    zlControl.TxtSelAll txtInput
    If mblnIme Then Call OpenIme(gstrIme)
    If mblnPassInput = False Then Exit Sub
    Call OpenPassKeyboard(txtInput, mblnAffirmPass)
End Sub
Private Sub txtInput_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call cmdOK_Click
    Else
        If mbytType = 1 Then '数字
            If InStr("0123456789" & Chr(27), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
        ElseIf mbytType = 2 Then '日期
            If InStr("0123456789:-" & Chr(27), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
        ElseIf mbytType = 3 Then '字母和数字
            If InStr("0123456789abcdefghijklmnopqrstuvwxyz" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
        End If
        If mblnUcase Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtInput_LostFocus()
    txtInput.Text = Trim(txtInput.Text)
    If mblnIme Then Call OpenIme
    If mblnPassInput Then Call ClosePassKeyboard(txtInput)
End Sub

Private Function OpenPassKeyboard(ctlText As Control, Optional bln确认密码 As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打开密码键盘输入
    '返回:打成成功,返回true,否者False
    '编制:刘兴洪
    '日期:2011-07-25 00:04:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjKeyboard Is Nothing Then Exit Function
    If mobjKeyboard.OpenPassKeyoardInput(Me, ctlText, bln确认密码) = False Then Exit Function
    OpenPassKeyboard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
Private Function ClosePassKeyboard(ctlText As Control) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打开密码键盘输入
    '返回:打成成功,返回true,否者False
    '编制:刘兴洪
    '日期:2011-07-25 00:04:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjKeyboard Is Nothing Then Exit Function
    If mobjKeyboard.ColsePassKeyoardInput(Me, ctlText) = False Then Exit Function
    ClosePassKeyboard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Private Function CreateObjectKeyboard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建密码创建
    '返回:创建成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-24 23:59:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    Set mobjKeyboard = CreateObject("zl9Keyboard.clsKeyboard")
    If Err <> 0 Then Exit Function
    Err = 0
    CreateObjectKeyboard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

