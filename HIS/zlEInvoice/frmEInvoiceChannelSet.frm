VERSION 5.00
Begin VB.Form frmEInvoiceChannelSet 
   Caption         =   "收费渠道对照设置"
   ClientHeight    =   2160
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5385
   Icon            =   "frmEInvoiceChannelSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   5385
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4080
      TabIndex        =   5
      Top             =   340
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4080
      TabIndex        =   4
      Top             =   840
      Width           =   1100
   End
   Begin VB.Frame fra 
      Height          =   3400
      Left            =   3840
      TabIndex        =   3
      Top             =   -120
      Width           =   15
   End
   Begin VB.TextBox txtEdit 
      Enabled         =   0   'False
      Height          =   300
      Index           =   0
      Left            =   1200
      TabIndex        =   2
      Top             =   330
      Width           =   2475
   End
   Begin VB.TextBox txtEdit 
      Enabled         =   0   'False
      Height          =   300
      Index           =   1
      Left            =   1200
      MaxLength       =   20
      TabIndex        =   1
      Top             =   915
      Width           =   2475
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   2
      Left            =   1200
      MaxLength       =   20
      TabIndex        =   0
      Top             =   1455
      Width           =   2475
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "卡类别名称"
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   8
      Top             =   360
      Width           =   900
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "结算方式"
      Height          =   180
      Index           =   1
      Left            =   360
      TabIndex        =   7
      Top             =   945
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "渠道编码"
      Height          =   180
      Index           =   2
      Left            =   360
      TabIndex        =   6
      Top             =   1485
      Width           =   720
   End
End
Attribute VB_Name = "frmEInvoiceChannelSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum TXT_Idex
    Idex_卡类别名称 = 0
    Idex_结算方式 = 1
    Idex_渠道编码 = 2
End Enum
Private mByteMode As Byte      '修改：1，新增：0
Private mblnOK As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If IsValid = False Then Exit Sub
    If Save收费渠道对照 = False Then Exit Sub
    mblnOK = True
    Unload Me
End Sub

Private Function Save收费渠道对照() As Boolean
    Dim strSQL As String

    On Error GoTo errHandle
    '新增或修改收费渠道对照
    strSQL = "Zl_收费渠道对照_Update("
    '操作类型_In In Number,
    strSQL = strSQL & mByteMode & ","
    '结算方式_In In 收费渠道对照.结算方式%Type,
    strSQL = strSQL & "'" & txtEdit(Idex_结算方式).Text & "',"
    '卡类别id_In In 收费渠道对照.卡类别id%Type,
    strSQL = strSQL & ZVal(txtEdit(Idex_卡类别名称).Tag) & ","
    '渠道编码_In In 收费渠道对照.渠道编码%Type
    strSQL = strSQL & "'" & txtEdit(Idex_渠道编码).Text & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, "收费渠道对照")
    
    Save收费渠道对照 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If Not ((KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 8 Or KeyAscii = 13) Then KeyAscii = 0
End Sub

Public Sub ShowMe(ByVal frmMain As Object, ByVal bytMode As Byte, ByVal lng卡类别ID As Long, ByVal str卡类别名称 As String, _
                                ByVal str结算方式 As String, Optional ByVal str渠道编码 As String, Optional blnRefresh As Boolean)
    mblnOK = False
    txtEdit(Idex_结算方式).Text = str结算方式
    txtEdit(Idex_卡类别名称).Tag = lng卡类别ID
    txtEdit(Idex_卡类别名称).Text = IIf(str卡类别名称 = "", "无", str卡类别名称)
    txtEdit(Idex_渠道编码).Text = str渠道编码
    mByteMode = bytMode
    Me.Show 1, frmMain
    blnRefresh = mblnOK
End Sub

Private Function IsValid() As Boolean
    On Error GoTo errHandle

    If Len(txtEdit(Idex_渠道编码).Text) = 0 Then
        MsgBox "渠道编码不能为空。", vbExclamation, gstrSysName
        zlControl.ControlSetFocus txtEdit(Idex_渠道编码)
        Exit Function
    End If

    If Not IsNumeric(txtEdit(Idex_渠道编码).Text) Or InStr(txtEdit(Idex_渠道编码).Text, ",") > 0 Or InStr(txtEdit(Idex_渠道编码).Text, ".") > 0 Then
        MsgBox "编码应由数字组成。", vbExclamation, gstrSysName
        zlControl.ControlSetFocus txtEdit(Idex_渠道编码)
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

