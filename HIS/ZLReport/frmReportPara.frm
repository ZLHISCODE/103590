VERSION 5.00
Begin VB.Form frmReportPara 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "报表参数设置"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5205
   Icon            =   "frmReportPara.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      Begin VB.CheckBox chkReportUse 
         Caption         =   "更新报表使用状态"
         Height          =   180
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtBegin 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   2070
         MaxLength       =   7
         TabIndex        =   5
         Text            =   "3000"
         Top             =   975
         Width           =   765
      End
      Begin VB.TextBox txtEnd 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   3150
         MaxLength       =   7
         TabIndex        =   6
         Text            =   "1000000"
         Top             =   975
         Width           =   765
      End
      Begin VB.CheckBox chkMiddle 
         Caption         =   "性能问题判断时包含中型表"
         Height          =   180
         Left            =   240
         TabIndex        =   3
         Top             =   760
         Width           =   2600
      End
      Begin VB.CheckBox chkReportLog 
         Caption         =   "开启报表运行日志"
         Height          =   180
         Left            =   240
         TabIndex        =   2
         Top             =   510
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "~"
         Height          =   180
         Left            =   2940
         TabIndex        =   9
         Top             =   1065
         Width           =   90
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "中型表记录数范围"
         Height          =   180
         Left            =   495
         TabIndex        =   4
         Top             =   1035
         Width           =   1440
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   2760
      TabIndex        =   7
      Top             =   1800
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3960
      TabIndex        =   8
      Top             =   1800
      Width           =   1100
   End
End
Attribute VB_Name = "frmReportPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnState As Boolean

Public Function ShowMe(ByVal frmOwner As Form) As Boolean
    Me.Show vbModal, frmOwner
    ShowMe = mblnState
End Function

Private Sub chkMiddle_Click()
    txtBegin.Enabled = chkMiddle.Value
    txtEnd.Enabled = chkMiddle.Value
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strSQL As String
    Dim strBegin As String
    Dim strEnd As String
    
    On Error GoTo ErrHand
    
    If chkMiddle.Value = 1 Then
        strBegin = Trim(txtBegin.Text)
        strEnd = Trim(txtEnd.Text)
        If Val(strBegin) = Val(strEnd) Then
            MsgBox "请检查中型表记录数范围（范围值相同）！", vbInformation, App.Title
            Exit Sub
        End If
        If Val(strBegin) > Val(strEnd) Then
            MsgBox "请检查中型表记录数范围（范围值颠倒）！", vbInformation, App.Title
            Exit Sub
        End If
    End If
    
    strSQL = "Zl_Parameters_Update('记录报表使用痕迹', " & chkReportUse.Value & ", 0, 0)"
    gcnOracle.Execute strSQL
    
    strSQL = "Zl_Parameters_Update('开启报表运行日志', " & chkReportLog.Value & ", 0, 0)"
    gcnOracle.Execute strSQL
    
    If chkMiddle.Value = 1 Then
        strSQL = "Zl_Parameters_Update('检查中型表', '" & strBegin & "," & strEnd & "', 0, 0)"
    Else
        strSQL = "Zl_Parameters_Update('检查中型表', '0,0', 0, 0)"
    End If
    gcnOracle.Execute strSQL
    
    mblnState = True
    Unload Me
    Exit Sub
    
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    '读取参数
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim i As Integer
    
    On Error GoTo errH
    
    mblnState = False
    
    strSQL = _
        "Select 28 参数号, zl_GetSysParameter('记录报表使用痕迹', 0, 0) 参数值 From Dual " & vbNewLine & _
        "Union All " & vbNewLine & _
        "Select 26, zl_GetSysParameter('开启报表运行日志', 0, 0) From Dual " & vbNewLine & _
        "Union All " & vbNewLine & _
        "Select 27, zl_GetSysParameter('检查中型表', 0, 0) From Dual "
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption)
    Do While rsTmp.EOF = False
        Select Case rsTmp("参数号").Value
        Case Val("26-开启报表运行日志")
            chkReportLog.Value = Val(Nvl(rsTmp("参数值").Value))
        Case Val("27-检查中型表")
            chkMiddle.Value = IIF(Nvl(rsTmp("参数值").Value, "0,0") = "0,0", 0, 1)
            If chkMiddle.Value = 1 Then
                txtBegin.Text = Split(Nvl(rsTmp("参数值").Value), ",")(0)
                txtEnd.Text = Split(Nvl(rsTmp("参数值").Value) & ",", ",")(1)
            Else
                txtBegin.Text = "3000"
                txtEnd.Text = "1000000"
            End If
            txtBegin.Tag = txtBegin.Text
            txtEnd.Tag = txtEnd.Text
        Case Val("28-记录报表使用痕迹")
            chkReportUse.Value = Val(Nvl(rsTmp("参数值").Value))
        End Select
        
        rsTmp.MoveNext
    Loop
    rsTmp.Close
    
    Call chkMiddle_Click
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtBegin_KeyPress(KeyAscii As Integer)
    If InStr("1234567890", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then KeyAscii = 0
End Sub

Private Sub txtBegin_Validate(Cancel As Boolean)
    If Val(txtEnd.Text) <= Val(txtBegin.Text) Then
        MsgBox "最小记录数应该比最大记录数小，请检查。", vbInformation, App.Title
        Cancel = True
    End If
    If Val(txtBegin.Text) < 1000 Then
        MsgBox "中型表记录数应该大于等于1000条记录数。", vbInformation, App.Title
        Cancel = True
    End If
    If Cancel = False Then
        txtBegin.Tag = txtBegin.Text
    Else
        txtBegin.Text = txtBegin.Tag
    End If
End Sub

Private Sub txtEnd_KeyPress(KeyAscii As Integer)
    If InStr("1234567890", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then KeyAscii = 0
End Sub

Private Sub txtEnd_Validate(Cancel As Boolean)
    If Val(txtEnd.Text) <= Val(txtBegin.Text) Then
        MsgBox "最小记录数应该比最大记录数小，请检查。", vbInformation, App.Title
        Cancel = True
        txtEnd.Text = txtEnd.Tag
    End If
    If Cancel = False Then
        txtEnd.Tag = txtEnd.Text
    End If
End Sub



