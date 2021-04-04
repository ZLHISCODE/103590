VERSION 5.00
Begin VB.Form frmEInvoiceInsureSet 
   Caption         =   "支付类别对照设置"
   ClientHeight    =   2505
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5385
   Icon            =   "frmEInvoiceInsureSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   5385
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   3
      Left            =   960
      MaxLength       =   25
      TabIndex        =   3
      Top             =   1920
      Width           =   2955
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   2
      Left            =   960
      MaxLength       =   10
      TabIndex        =   2
      Top             =   1390
      Width           =   2955
   End
   Begin VB.TextBox txtEdit 
      Enabled         =   0   'False
      Height          =   300
      Index           =   1
      Left            =   960
      TabIndex        =   1
      Top             =   860
      Width           =   2955
   End
   Begin VB.TextBox txtEdit 
      Enabled         =   0   'False
      Height          =   300
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   330
      Width           =   2955
   End
   Begin VB.Frame fra 
      Height          =   3400
      Left            =   4000
      TabIndex        =   6
      Top             =   -120
      Width           =   15
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4140
      TabIndex        =   5
      Top             =   840
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4140
      TabIndex        =   4
      Top             =   345
      Width           =   1100
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "大类名称"
      Height          =   180
      Index           =   3
      Left            =   180
      TabIndex        =   10
      Top             =   1950
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "大类编码"
      Height          =   180
      Index           =   2
      Left            =   180
      TabIndex        =   9
      Top             =   1425
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "支付名称"
      Height          =   180
      Index           =   1
      Left            =   180
      TabIndex        =   8
      Top             =   885
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "保险类别"
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   7
      Top             =   360
      Width           =   720
   End
End
Attribute VB_Name = "frmEInvoiceInsureSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum TXT_Idex
    Idex_保险名称 = 0
    Idex_支付名称 = 1
    Idex_大类编码 = 2
    Idex_大类名称 = 3
End Enum
Private mByteMode As Byte      '修改：1，新增：0
Private mblnOK As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdOK_Click()
    If IsValid = False Then Exit Sub
    If Save支付类别对照 = False Then Exit Sub
    mblnOK = True
    Unload Me
End Sub

Private Function Save支付类别对照() As Boolean
    Dim strSQL As String

    On Error GoTo errHandle
    '新增或修改收费渠道对照
    strSQL = "Zl_支付类别对照_Update("
    '操作类型_In In Number,
    strSQL = strSQL & mByteMode & ","
    '保险大类id_In In 支付类别对照.保险大类id%Type,
    strSQL = strSQL & Val(txtEdit(Idex_支付名称).Tag) & ","
    '大类编码_In   In 支付类别对照.大类编码%Type := Null,
    strSQL = strSQL & "'" & txtEdit(Idex_大类编码).Text & "',"
    '大类名称_In   In 支付类别对照.大类名称%Type := Null
    strSQL = strSQL & "'" & txtEdit(Idex_大类名称).Text & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, "收费渠道对照")
    
    Save支付类别对照 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = Idex_大类编码 Then
        If Not ((KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 8 Or KeyAscii = 13) Then KeyAscii = 0
    End If
End Sub

Public Sub ShowMe(ByVal frmMain As Object, ByVal bytMode As Byte, ByVal str险类名称 As String, ByVal lng保险大类ID As Long, _
                                ByVal str支付名称 As String, Optional ByVal str大类编码 As String, Optional ByVal str大类名称 As String, _
                                Optional blnRefresh As Boolean)
    mblnOK = False
    txtEdit(Idex_保险名称).Text = str险类名称
    txtEdit(Idex_支付名称).Tag = lng保险大类ID
    txtEdit(Idex_支付名称).Text = str支付名称
    txtEdit(Idex_大类编码).Text = str大类编码
    txtEdit(Idex_大类名称).Text = str大类名称
    mByteMode = bytMode
    Me.Show 1, frmMain
    blnRefresh = mblnOK
End Sub

Private Function IsValid() As Boolean
    On Error GoTo errHandle

    If Len(txtEdit(Idex_大类编码).Text) = 0 Then
        MsgBox "大类编码不能为空。", vbExclamation, gstrSysName
        zlControl.ControlSetFocus txtEdit(Idex_大类编码)
        Exit Function
    End If

    If Not IsNumeric(txtEdit(Idex_大类编码).Text) Or InStr(txtEdit(Idex_大类编码).Text, ",") > 0 Or InStr(txtEdit(Idex_大类编码).Text, ".") > 0 Then
        MsgBox "编码应由数字组成。", vbExclamation, gstrSysName
        zlControl.ControlSetFocus txtEdit(Idex_大类编码)
        Exit Function
    End If
    
    If Len(txtEdit(Idex_大类名称).Text) = 0 Then
        MsgBox "大类名称不能为空。", vbExclamation, gstrSysName
        zlControl.ControlSetFocus txtEdit(Idex_大类名称)
        Exit Function
    End If
    
    IsValid = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function



