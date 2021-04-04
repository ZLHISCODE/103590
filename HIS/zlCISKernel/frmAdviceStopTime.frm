VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAdviceStopTime 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "停止医嘱"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4170
   Icon            =   "frmAdviceStopTime.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraTZYY 
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   1305
      TabIndex        =   6
      Top             =   990
      Width           =   2205
      Begin VB.TextBox txtTZYY 
         Height          =   300
         Left            =   0
         MaxLength       =   200
         TabIndex        =   1
         Top             =   240
         Width           =   1800
      End
      Begin VB.CommandButton cmdSel 
         Caption         =   "…"
         Height          =   265
         Left            =   1800
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   240
         Width           =   285
      End
      Begin VB.Label lblTZYY 
         AutoSize        =   -1  'True
         Caption         =   "执行终止原因"
         Height          =   180
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   1080
      End
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Left            =   -255
      TabIndex        =   4
      Top             =   1800
      Width           =   4845
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   2625
      TabIndex        =   3
      Top             =   1920
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   1485
      TabIndex        =   2
      Top             =   1920
      Width           =   1100
   End
   Begin MSComCtl2.DTPicker dtpTime 
      Height          =   300
      Left            =   1320
      TabIndex        =   0
      Top             =   645
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   232980483
      UpDown          =   -1  'True
      CurrentDate     =   39668.3388888889
   End
   Begin VB.Image imgCharge 
      Height          =   240
      Left            =   240
      Picture         =   "frmAdviceStopTime.frx":058A
      Top             =   1170
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image ImgAudit 
      Height          =   720
      Left            =   360
      Picture         =   "frmAdviceStopTime.frx":6DDC
      Stretch         =   -1  'True
      Top             =   255
      Width           =   720
   End
   Begin VB.Label lblBT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "执行终止时间"
      Height          =   180
      Left            =   1320
      TabIndex        =   5
      Top             =   375
      Width           =   1080
   End
   Begin VB.Image ImgStop 
      Height          =   720
      Left            =   360
      Picture         =   "frmAdviceStopTime.frx":7166
      Top             =   255
      Width           =   720
   End
End
Attribute VB_Name = "frmAdviceStopTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrsAdvice As ADODB.Recordset
Private mlng医嘱ID As Long
Private mblnOK As Boolean

Private mstrTime As String
Private mintMode As Integer '0-医嘱停止，1－医嘱审核的窗体，2－护士输液配药记录销帐
Private mdatRegister As Date
Private mstr原因 As String '当 mintMode=0 医嘱停止时要求录入停止原因。
Private mlng科室ID As Long

Public Function ShowMe(frmParent As Object, ByVal lng医嘱ID As Long, ByVal lng科室ID As Long, Optional ByVal intMode As Integer = 0, Optional ByVal datRegister As Date = 0, Optional ByRef str原因 As String) As String
     '******************************************************************************************************************
    '参数：intMode,为1的话表示是弹出医嘱审核的窗体
    '      datRegister,医嘱执行的登记时间
    '说明：返回选择的时间的字符串
    '******************************************************************************************************************
    mlng医嘱ID = lng医嘱ID
    mlng科室ID = lng科室ID
    mintMode = intMode
    mdatRegister = datRegister
    Me.Show 1, frmParent
    If mblnOK Then
        str原因 = mstr原因
        ShowMe = mstrTime
    End If
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
   '检查合法性
    If cmdOK.Visible And cmdOK.Enabled Then cmdOK.SetFocus
    
    If mintMode = 0 Then
        '必须大于开始执行时间
        If Format(dtpTime.value, "yyyy-MM-dd HH:mm") <= Format(mrsAdvice!开始执行时间, "yyyy-MM-dd HH:mm") Then
            MsgBox "输入的执行终止时间必须大于医嘱的开始执行时间 " & Format(mrsAdvice!开始执行时间, "yyyy-MM-dd HH:mm") & "。", vbInformation, gstrSysName
            dtpTime.SetFocus: Exit Sub
        End If
        '登记执行时间>上次执行时间
        mstrTime = GetAdviceStopTime(mlng医嘱ID)
        If mstrTime <> "" Then
            If Format(dtpTime.value, "yyyy-MM-dd HH:mm") < Format(mstrTime, "yyyy-MM-dd HH:mm") Then
                MsgBox "不能停止到执行时间 " & mstrTime & " 之前，请调整停止时间，如果确实要停止到执行时间之前，请先取消执行登记。", vbInformation, gstrSysName
                dtpTime.SetFocus: Exit Sub
            End If
        End If
        '不应小于上次执行时间
        If Not IsNull(mrsAdvice!上次执行时间) Then
            If Format(dtpTime.value, "yyyy-MM-dd HH:mm") < Format(mrsAdvice!上次执行时间, "yyyy-MM-dd HH:mm") Then
                If MsgBox("输入的执行终止时间小于医嘱的上次执行时间 " & Format(mrsAdvice!上次执行时间, "yyyy-MM-dd HH:mm") & "，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    dtpTime.SetFocus: Exit Sub
                End If
            End If
        End If
        
        '未填写终止原因
        If gbln医嘱终止原因 Then
            If Trim(txtTZYY.Text) = "" And InStr(gstr可不填停嘱原因科室, "," & mlng科室ID & ",") = 0 Then
                MsgBox "请录入终止原因。", vbInformation, gstrSysName
                txtTZYY.SetFocus: Exit Sub
            If zlCommFun.ActualLen(txtTZYY.Text) > txtTZYY.MaxLength Then
                    MsgBox "终止原因内容太长，最多允许 " & txtTZYY.MaxLength / 2 & " 个汉字或 " & txtTZYY.MaxLength & " 个字符。", vbInformation, gstrSysName
                    txtTZYY.SetFocus: Exit Sub
                End If
            End If
            mstr原因 = Trim(txtTZYY.Text)
        End If
    ElseIf mintMode = 1 Then
        '必须大于执行时间
        If Format(dtpTime.value, "yyyy-MM-dd HH:mm") < Format(mdatRegister, "yyyy-MM-dd HH:mm") Then
            MsgBox "输入的核对时间不能够小于医嘱执行的登记时间 " & Format(mdatRegister, "yyyy-MM-dd HH:mm") & "。", vbExclamation, gstrSysName
            dtpTime.SetFocus: Exit Sub
        End If
    ElseIf mintMode = 2 Then
        If Trim(txtTZYY.Text) = "" Then
            MsgBox "请录入销帐原因。", vbInformation, gstrSysName
            txtTZYY.SetFocus: Exit Sub
        If zlCommFun.ActualLen(txtTZYY.Text) > txtTZYY.MaxLength Then
                MsgBox "销帐原因内容太长，最多允许 " & txtTZYY.MaxLength / 2 & " 个汉字或 " & txtTZYY.MaxLength & " 个字符。", vbInformation, gstrSysName
                txtTZYY.SetFocus: Exit Sub
            End If
        End If
        mstr原因 = Trim(txtTZYY.Text)
    End If
    mstrTime = Format(dtpTime.value, "yyyy-MM-dd HH:mm")
    mblnOK = True
    Unload Me
End Sub

Private Sub cmdSel_Click()
'功能：弹出选择器
    Call GetItem原因(1)
End Sub

Private Sub dtpTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call cmdOK_Click
End Sub

Private Sub Form_Activate()
    If dtpTime.Enabled Then dtpTime.SetFocus
    Me.Refresh
    zlCommFun.PressKey vbKeyRight
    zlCommFun.PressKey vbKeyRight
    zlCommFun.PressKey vbKeyRight
End Sub

Private Sub Form_Load()
    Dim datCurr As Date
    Dim strSQL As String
    
    mblnOK = False
    datCurr = zlDatabase.Currentdate
    
    On Error GoTo errH
    If mintMode = 0 Then
        ImgAudit.Visible = False
        ImgStop.Visible = True
        fraTZYY.Visible = gbln医嘱终止原因
        
        Set Me.Icon = ImgStop.Picture
        
        lblBT.Caption = "执行终止时间"
        Me.Caption = "停止医嘱"
        
        strSQL = "Select 开始执行时间,执行终止时间,上次执行时间,开嘱时间 From 病人医嘱记录 Where ID=[1]"
        Set mrsAdvice = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng医嘱ID)
        
        If gbln长期医嘱次日生效 Then
            dtpTime.value = CDate(Format(datCurr + 1, "yyyy-MM-dd 00:00"))
        Else
            dtpTime.value = CDate(Format(datCurr, "yyyy-MM-dd HH:mm"))
        End If
        
        If Not IsNull(mrsAdvice!上次执行时间) Then
            If Format(dtpTime.value, "yyyy-MM-dd HH:mm") < Format(mrsAdvice!上次执行时间, "yyyy-MM-dd HH:mm") Then
                dtpTime.value = Format(mrsAdvice!上次执行时间, "yyyy-MM-dd HH:mm")
            End If
        End If
    ElseIf mintMode = 1 Then
        ImgAudit.Visible = True
        ImgStop.Visible = False
        fraTZYY.Visible = False
        Set Me.Icon = ImgAudit.Picture
        
        lblBT.Caption = "核对时间"
        Me.Caption = "医嘱核对"
        dtpTime.value = CDate(Format(datCurr, "yyyy-MM-dd HH:mm"))
    ElseIf mintMode = 2 Then
        Set Me.Icon = imgCharge.Picture
        Me.Caption = "输液配药记录销帐"
        lblTZYY.Caption = "销帐原因"
        Set ImgAudit.Picture = imgCharge.Picture
        ImgStop.Visible = False
        dtpTime.Enabled = False
        strSQL = "select 操作说明 from 病人医嘱状态 where 医嘱id=[1] and 操作类型=8 and 操作说明 is not null"
        Set mrsAdvice = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng医嘱ID)
        If Not mrsAdvice.EOF Then
            txtTZYY.Text = mrsAdvice!操作说明 & ""
        End If
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    
    If mintMode = 1 Or mintMode = 0 And Not gbln医嘱终止原因 Then
        fraLine.Top = 1800 - fraTZYY.Height
        cmdOK.Top = 1920 - fraTZYY.Height
        cmdCancel.Top = 1920 - fraTZYY.Height
        Me.Height = 2775 - fraTZYY.Height
    ElseIf mintMode = 2 Then
        fraTZYY.Top = lblBT.Top
        
        fraLine.Top = 1800 - fraTZYY.Height
        cmdOK.Top = 1920 - fraTZYY.Height
        cmdCancel.Top = 1920 - fraTZYY.Height
        Me.Height = 2775 - fraTZYY.Height
        
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mrsAdvice Is Nothing Then
        If mrsAdvice.State = 1 Then mrsAdvice.Close
        Set mrsAdvice = Nothing
    End If
End Sub

Private Sub GetItem原因(ByVal intType As Integer)
'功能：选择停嘱原因
'参数：intType =0 KeyPress调用，=1 下拉按钮调用
    Dim strSQL As String, rsTmp As Recordset
    Dim blnCancel As Boolean, vRect As RECT
    Dim strMatch As String
    Dim strInput As String
    
    On Error GoTo errH
    
    If intType = 0 Then
       strInput = txtTZYY.Text
       If IsNumeric(strInput) Then '10,11.输入全是数字时只匹配编码
           If Mid(gstrMatchMode, 1, 1) = "1" Then strMatch = " where  A.编码 Like [1]"
       ElseIf zlCommFun.IsCharAlpha(strInput) Then '01,11.输入全是字母时只匹配简码
           If Mid(gstrMatchMode, 2, 1) = "1" Then strMatch = " where  a.简码 Like [1]"
       ElseIf zlCommFun.IsCharChinese(strInput) Then
           strMatch = " where  a.名称 Like [1]"
       End If
    End If
    
    strSQL = "select a.编码 as id, a.编码,a.名称,a.简码 from 停嘱原因 a  " & strMatch & " order by a.编码"
    vRect = zlControl.GetControlRect(txtTZYY.hwnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, Me.Caption, False, "", "", False, False, True, _
        vRect.Left, vRect.Top, txtTZYY.Height, blnCancel, False, True, UCase(txtTZYY.Text) & "%")

    If Not rsTmp Is Nothing Then
''        If Not blnCancel Then
''            MsgBox "未找到匹配的项目。", vbInformation, gstrSysName
''        End If
''        Call zlControl.TxtSelAll(txtTZYY)
''        txtTZYY.SetFocus: Exit Sub
'    Else
        txtTZYY.Text = rsTmp!名称 & ""
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtTZYY_GotFocus()
    Call zlControl.TxtSelAll(txtTZYY)
End Sub

Private Sub txtTZYY_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call GetItem原因(0)
    Else
        If KeyAscii = 39 Then KeyAscii = 0 '单引号
    End If
End Sub
