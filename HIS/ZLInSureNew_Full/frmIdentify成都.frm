VERSION 5.00
Begin VB.Form frmIdentify成都 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "医保病人身份识别"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6165
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmIdentify成都.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox txtCard 
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1965
      MaxLength       =   30
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   675
      Width           =   3765
   End
   Begin VB.TextBox txtPwd 
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1965
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1125
      Width           =   3765
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Enabled         =   0   'False
      Height          =   405
      Left            =   2745
      TabIndex        =   2
      Top             =   2220
      Width           =   1305
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   405
      Left            =   4200
      TabIndex        =   3
      Top             =   2220
      Width           =   1305
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   75
      Left            =   -180
      TabIndex        =   4
      Top             =   2025
      Width           =   6660
   End
   Begin VB.Label lblNote 
      Caption         =   "请在正确刷卡之后，输入个人密码。"
      Height          =   255
      Left            =   900
      TabIndex        =   8
      Top             =   165
      Width           =   3645
   End
   Begin VB.Label lblCard 
      AutoSize        =   -1  'True
      Caption         =   "卡号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1290
      TabIndex        =   7
      Top             =   735
      Width           =   510
   End
   Begin VB.Label lblPwd 
      AutoSize        =   -1  'True
      Caption         =   "密码"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1290
      TabIndex        =   6
      Top             =   1185
      Width           =   510
   End
   Begin VB.Label lblPatiInfo 
      AutoSize        =   -1  'True
      Caption         =   "病人信息"
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   195
      TabIndex        =   5
      Top             =   1740
      Width           =   960
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   255
      Picture         =   "frmIdentify成都.frx":030A
      Top             =   345
      Width           =   480
   End
End
Attribute VB_Name = "frmIdentify成都"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mstrPatiInfo As String
Public mcur余额 As Currency
'200308z012:住院病人使用
Public mcur住院基数 As Currency
Public mcur住院限额 As Currency
Public mcur报销比例 As Currency

Private mstr医保号 As String
Private mstr卡号 As String

Private mintTimes As Integer
Private mintCardLen As Integer

Private Sub cmdCancel_Click()
    mstrPatiInfo = ""
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    If mstr医保号 = "" And mstr卡号 = "" Then
        MsgBox "未正确地刷卡,不能通过验证！", vbInformation, gstrSysName
        Exit Sub
    End If
    Me.Hide
End Sub

Private Sub Form_Load()
    mstrPatiInfo = ""
    mcur余额 = 0
    mcur住院基数 = 0
    mcur住院限额 = 0
    mcur报销比例 = 0
    
    mintTimes = 0
    Me.lblPatiInfo.Caption = ""
    mintCardLen = GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("CardNoLength"), 26)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstr医保号 = "": mstr卡号 = ""
End Sub

Private Sub txtCard_GotFocus()
    zlControl.TxtSelAll txtCard
    If gblnLED And txtCard.Text = "" Then
        zl9LedVoice.Speak "#5"
    End If
End Sub

Private Sub txtCard_KeyPress(KeyAscii As Integer)
'功能：刷卡并分析个个编码、卡号
    Dim str医保号 As String, str卡号 As String
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call ExecuteZ015(txtCard.Text, str医保号, str卡号)
        If str医保号 = "" And str卡号 = "" Then
            MsgBox "刷卡分析失败，请重试！", vbInformation, gstrSysName
            txtCard.Text = "": txtCard.SetFocus: Exit Sub
        Else
            mstr医保号 = str医保号
            mstr卡号 = str卡号
            txtPwd.SetFocus: Exit Sub
        End If
    End If
    If txtCard.SelLength = Len(txtCard.Text) Then txtCard.Text = ""
    
    If Len(txtCard.Text) + 1 = mintCardLen Then
        txtCard.Text = txtCard.Text & Chr(KeyAscii)
        KeyAscii = 0
        Call ExecuteZ015(txtCard.Text, str医保号, str卡号)
        If str医保号 = "" And str卡号 = "" Then
            MsgBox "刷卡分析失败，请重试！", vbInformation, gstrSysName
            txtCard.Text = "": txtCard.SetFocus: Exit Sub
        Else
            mstr医保号 = str医保号
            mstr卡号 = str卡号
            txtPwd.SetFocus: Exit Sub
        End If
    End If
    
    Me.cmdOK.Enabled = False
    Me.lblPatiInfo.Caption = ""
    Me.txtPwd.Text = ""
End Sub

Private Sub txtPwd_GotFocus()
    zlControl.TxtSelAll txtPwd
    If gblnLED And txtPwd.Text = "" Then
        zl9LedVoice.Speak "#0"
    End If
End Sub

Private Sub txtPwd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtPwd.Text = "" Then Exit Sub
        KeyAscii = 0: Call zlControl.TxtSelAll(txtPwd)
        If mstr医保号 = "" And mstr卡号 = "" Then Exit Sub
        
        Call ThisIdentify: Exit Sub
    End If
    
    If txtPwd.SelLength = Len(txtPwd.Text) Then txtPwd.Text = ""
    
    If Len(txtPwd.Text) + 1 = txtPwd.MaxLength Then
        txtPwd.Text = txtPwd.Text & Chr(KeyAscii)
        KeyAscii = 0: Call zlControl.TxtSelAll(txtPwd)
        If mstr医保号 = "" And mstr卡号 = "" Then Exit Sub
        
        Call ThisIdentify: Exit Sub
    End If
End Sub

Private Sub ThisIdentify()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    Dim strSelfNo As String, strSelfPwd As String, strSerial As String, strKH As String
    Dim strSwapNo As String         '交易顺序号
    
    strSelfNo = mstr医保号
    strKH = mstr卡号
    strSelfPwd = TrimStr(txtPwd.Text)

    mintTimes = mintTimes + 1
    strSQL = "select 部门表_id.nextval||'1' from dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    With rsTmp
        strSwapNo = .Fields(0).Value
        strSerial = getSerial(strSelfNo)
        
        'New:交易编号,客户机编号,交易顺序号,密码,操作员编号,就诊登记号,医保号,医院编码,交易时间,数据批号,支付类别,卡号
        strSQL = "z001('z001','" & UserInfo.站点 & "','" & strSwapNo & "','" & strSelfPwd & "','" & UserInfo.编号 & "'," & _
            "'" & strSerial & "','" & strSelfNo & "','" & Trim(gstr医院编码) & "','" & DateStr & "','" & strSwapNo & "','" & IIf(Me.Tag = 0, "11", "31") & "','" & strKH & "')"
        gcnSybase.Execute strSQL, , adCmdStoredProc
        
        If .State = adStateOpen Then .Close
        .Open "select code from zjycl  where jysxh='" & strSwapNo & "' and jybh='z001' order by jyend desc", gcnSybase, adOpenStatic, adLockReadOnly
        If Trim(.Fields(0).Value) <> "0000" Then
            MsgBox "交易""z001""出现错误""" & !CODE & """:" & vbCrLf & String(2, "　") & GetErrInfo(!CODE, TYPE_成都市) & String(2, vbTab), vbInformation, gstrSysName
            If mintTimes > 6 Then
                MsgBox "无法识别你的身份，请确认你的卡和密码后再来！", vbExclamation, gstrSysName
                mstrPatiInfo = "": Me.Hide: Exit Sub
            End If
            
            Me.lblNote.Caption = "无法识别身份，请重新刷卡！"
            Me.txtPwd.Text = ""
            Me.cmdOK.Enabled = False
            Me.txtPwd.SetFocus
            mstrPatiInfo = ""
        Else
            strSQL = "select * from grjbxx where grbm='" & strSelfNo & "'"
            If .State = adStateOpen Then .Close
            .CursorLocation = adUseClient
            .Open strSQL, gcnSybase, adOpenKeyset
            If Not .EOF Then
                'New:0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码)
                mstrPatiInfo = strKH & ";" & strSelfNo & ";" & strSelfPwd & ";" & _
                        TrimStr(.Fields("xm").Value) & ";" & _
                        IIf(TrimStr(Nvl(.Fields("xb").Value)) = "1", "男", "女") & ";" & _
                        TrimStr(Nvl(.Fields("csrq").Value)) & ";" & _
                        TrimStr(Nvl(.Fields("sfz").Value)) & ";" & _
                        TrimStr(Nvl(.Fields("dwmc").Value)) & "(" & Trim(Nvl(.Fields("dwbm").Value)) & ")"
                mcur余额 = IIf(IsNull(!grzhlnye), 0, !grzhlnye) + IIf(IsNull(!grzhbnye), 0, !grzhbnye)
                '200308z012:住院病人使用
                If Val(Me.Tag) <> 0 Then
                    mcur住院基数 = IIf(IsNull(!zyjs), 0, !zyjs)
                    mcur报销比例 = IIf(IsNull(!tcbxbl), 0, !tcbxbl)
                    mcur住院限额 = IIf(IsNull(!zyxe), 0, !zyxe)
                End If
                
                Me.lblNote.Caption = "已经正确完成身份识别。"
                Me.lblPatiInfo.Caption = "病人:" & Trim(.Fields("xm").Value) & "  " & IIf(Trim(Nvl(.Fields("xb").Value)) = "1", "男", "女") & "  " & Trim(Nvl(.Fields("csrq").Value)) & ",请确认！"
                
                '曾明春（2005-10-14） 身份验证成功后，进行余额提示。
                If gblnLED Then
                   zl9LedVoice.Speak "#26 " & mcur余额
                End If
                
                Me.cmdOK.Enabled = True
                Me.cmdOK.SetFocus
            Else
                Me.lblNote.Caption = "无法识别身份，请重新刷卡！"
                Me.txtPwd.Text = ""
                Me.cmdOK.Enabled = False
                Me.txtPwd.SetFocus
                mstrPatiInfo = ""
            End If
        End If
    End With
End Sub
