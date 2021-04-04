VERSION 5.00
Begin VB.Form frmModiPass 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "修改密码"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5685
   Icon            =   "frmModiPass.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   300
      TabIndex        =   7
      Top             =   2595
      Width           =   1500
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
      Height          =   420
      Left            =   3900
      TabIndex        =   6
      Top             =   2595
      Width           =   1500
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
      Height          =   420
      Left            =   2385
      TabIndex        =   5
      Top             =   2595
      Width           =   1500
   End
   Begin VB.Frame Frame2 
      Height          =   45
      Left            =   -45
      TabIndex        =   15
      Top             =   2325
      Width           =   7170
   End
   Begin VB.TextBox txtAudi 
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
      IMEMode         =   3  'DISABLE
      Left            =   3315
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1710
      Width           =   1605
   End
   Begin VB.TextBox txtPass 
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
      IMEMode         =   3  'DISABLE
      Left            =   825
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1710
      Width           =   1605
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
      IMEMode         =   3  'DISABLE
      Left            =   4380
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1080
      Width           =   525
   End
   Begin VB.TextBox txtSex 
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
      IMEMode         =   3  'DISABLE
      Left            =   3330
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1080
      Width           =   405
   End
   Begin VB.TextBox txtPati 
      BackColor       =   &H00EBFFFF&
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
      IMEMode         =   3  'DISABLE
      Left            =   840
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1080
      Width           =   1605
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   -225
      TabIndex        =   8
      Top             =   810
      Width           =   7170
   End
   Begin VB.Image imgFlag 
      Height          =   720
      Left            =   270
      Picture         =   "frmModiPass.frx":000C
      Top             =   60
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "验证"
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
      Left            =   2775
      TabIndex        =   14
      Top             =   1770
      Width           =   480
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "密码"
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
      Left            =   270
      TabIndex        =   13
      Top             =   1770
      Width           =   480
   End
   Begin VB.Label Label4 
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
      Left            =   3825
      TabIndex        =   12
      Top             =   1140
      Width           =   480
   End
   Begin VB.Label Label3 
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
      Left            =   2790
      TabIndex        =   11
      Top             =   1140
      Width           =   480
   End
   Begin VB.Label lbl病人 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "病人"
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
      Left            =   285
      TabIndex        =   10
      Top             =   1140
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "请将就诊卡从刷卡器上轻轻划过，  然后连续两次输入相同的密码！"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1260
      TabIndex        =   9
      Top             =   180
      Width           =   3810
   End
End
Attribute VB_Name = "frmModiPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Private mblnDO As Boolean
Private mobjKeyboard As Object

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

Private Sub cmdCancel_Click()
    If txtPati.Text <> "" And Val(txtPati.Tag) <> 0 Then
        Call ClearFace: txtPati.SetFocus: Exit Sub
    End If
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub cmdOK_Click()
    Dim strSQL As String
    Dim blnTrans As Boolean
    
    If Val(txtPati.Tag) = 0 Or txtPati.Text = "" Then
        If glngSys Like "8??" Then
            MsgBox "不能读取客户信息，请确定是否正确刷卡！", vbInformation, gstrSysName
        Else
            MsgBox "不能读取病人信息，请确定是否正确刷卡！", vbInformation, gstrSysName
        End If
        Call ClearFace: txtPati.SetFocus: Exit Sub
    End If
        
    If txtPass.Text <> txtAudi.Text Then
        MsgBox "两次输入的密码不一致，请重新输入！", vbInformation, gstrSysName
        txtPass.Text = "": txtAudi.Text = ""
        txtPass.SetFocus: Exit Sub
    End If
    
    If txtPass.Text = "" Then
        If MsgBox("当前设置的密码为空，确实要这样设置吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    
    strSQL = "zl_就诊卡记录_MODIPASS(" & Val(txtPati.Tag) & ",'" & txtPass.Text & "')"
    
    On Error GoTo errH
    
    gcnOracle.BeginTrans: blnTrans = True
'    Call SQLTest(App.ProductName, Me.Caption, strSQL)
'    gcnOracle.Execute strSQL
'    Call SQLTest
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    gcnOracle.CommitTrans: blnTrans = False
    
    gblnOK = True
    Unload Me
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    If glngSys Like "8??" Then lbl病人.Caption = "客户"
    Call CreateObjectKeyboard
    gblnOK = False
    Call ClearFace
    mblnDO = True
End Sub

Private Sub txtAudi_GotFocus()
    zlControl.TxtSelAll txtAudi
    OpenPassKeyboard txtAudi, True
End Sub

Private Sub txtAudi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: cmdOK.SetFocus
    Else
        If InStr("';" & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txtAudi_LostFocus()
   ClosePassKeyboard txtPass
End Sub

Private Sub txtPass_GotFocus()
    zlControl.TxtSelAll txtPass
    OpenPassKeyboard txtPass, False
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txtPass.Text = "" And txtAudi.Text = "" Then
            cmdOK.SetFocus
        Else
            txtAudi.SetFocus
        End If
    Else
        If InStr("';" & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txtPass_LostFocus()
    ClosePassKeyboard txtPass
End Sub

Private Sub txtPati_Change()
    If Not mblnDO Then Exit Sub
    If gblnShowCard Then
        txtPati.PasswordChar = ""
    Else
        txtPati.PasswordChar = "*"
    End If
End Sub

Private Sub txtPati_GotFocus()
    zlControl.TxtSelAll txtPati
End Sub

Private Sub txtPati_KeyPress(KeyAscii As Integer)
    If (Len(txtPati.Text) = gbytCardNOLen - 1 And KeyAscii <> 8) Or (KeyAscii = 13 And Trim(txtPati.Text) <> "") Then
        If KeyAscii <> 13 Then
            txtPati.Text = txtPati.Text & Chr(KeyAscii)
            txtPati.SelStart = Len(txtPati.Text)
        End If
        KeyAscii = 0
                
        If Not GetPatiFromCard(txtPati.Text) Then
            Call ClearFace
            If glngSys Like "8??" Then
                MsgBox "不能读取客户信息，请确定是否正确刷卡！", vbInformation, gstrSysName
            Else
                MsgBox "不能读取病人信息，请确定是否正确刷卡！", vbInformation, gstrSysName
            End If
            txtPati.SetFocus: Exit Sub
        End If
        
        txtPass.SetFocus
    End If
End Sub

Private Function GetPatiFromCard(strCard As String) As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    'by lesfeng 2010-03-08 性能优化 select *及绑定
    strSQL = "Select 病人ID,门诊号,住院号,就诊卡号,姓名,性别,年龄,病人类型,险类" & _
             "  From 病人信息 Where 就诊卡号=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strCard)
    
    If rsTmp.EOF Then Exit Function
    
    txtPati.PasswordChar = ""
    mblnDO = False
    txtPati.Text = IIf(IsNull(rsTmp!姓名), "", rsTmp!姓名)
    '74426:李南春,2014-7-9,病人姓名显示颜色处理
    Call SetPatiColor(txtPati, Nvl(rsTmp!病人类型), IIf(IsNull(rsTmp!险类), Me.ForeColor, vbRed))
    mblnDO = True
    txtPati.Tag = rsTmp!病人ID
    
    txtSex.Text = IIf(IsNull(rsTmp!性别), "", rsTmp!性别)
    txtAge.Text = IIf(IsNull(rsTmp!年龄), "", rsTmp!年龄)
    txtPass.Text = ""
    txtAudi.Text = ""
    
    GetPatiFromCard = True
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub ClearFace()
    If gblnShowCard Then
        txtPati.PasswordChar = ""
    Else
        txtPati.PasswordChar = "*"
    End If
    
    txtPati.Tag = ""
    
    txtPati.Text = ""
    txtSex.Text = ""
    txtAge.Text = ""
    txtPass.Text = ""
    txtAudi.Text = ""
End Sub
