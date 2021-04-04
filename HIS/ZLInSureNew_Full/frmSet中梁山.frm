VERSION 5.00
Begin VB.Form frmSet中梁山 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "保险参数设置"
   ClientHeight    =   2295
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3660
   Icon            =   "frmSet中梁山.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   3660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   1110
      TabIndex        =   5
      Top             =   1740
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   2395
      TabIndex        =   6
      Top             =   1740
      Width           =   1100
   End
   Begin VB.Frame fraTop 
      Caption         =   "运行参数"
      Height          =   1410
      Left            =   150
      TabIndex        =   0
      Top             =   180
      Width           =   3345
      Begin VB.TextBox txtEdit 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   0
         Left            =   2400
         MaxLength       =   2
         TabIndex        =   2
         Top             =   360
         Width           =   645
      End
      Begin VB.TextBox txtEdit 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   1
         Left            =   2400
         MaxLength       =   2
         TabIndex        =   4
         Top             =   735
         Width           =   645
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   240
         Picture         =   "frmSet中梁山.frx":000C
         Top             =   360
         Width           =   480
      End
      Begin VB.Label lblEdit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "医保卡号长度(&R)"
         Height          =   180
         Index           =   0
         Left            =   990
         TabIndex        =   1
         Top             =   420
         Width           =   1350
      End
      Begin VB.Label lblEdit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "退休证号长度(&T)"
         Height          =   180
         Index           =   1
         Left            =   990
         TabIndex        =   3
         Top             =   795
         Width           =   1350
      End
   End
End
Attribute VB_Name = "frmSet中梁山"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum enum编辑
    Text卡号长度 = 0
    Text退休证号 = 1
End Enum

Dim mlng险类 As Long, mlng中心 As Long
Dim mblnOK As Boolean
Dim mblnChange As Boolean     '是否改变了

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认退出？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim colPara As New Collection
    Dim lngCount As Long
    
    If Val(TxtEdit(Text卡号长度).Text) > 25 Then
        MsgBox "卡号长度不能超过25位。", vbInformation, gstrSysName
        TxtEdit(Text卡号长度).SetFocus
        Exit Sub
    End If
    
    If Val(TxtEdit(Text退休证号).Text) > 26 Then
        MsgBox "退休证号长度不能超过26位。", vbInformation, gstrSysName
        TxtEdit(Text退休证号).SetFocus
        Exit Sub
    End If
    
    gcnOracle.BeginTrans
    On Error GoTo errHandle
    
    '删除已经数据
    gstrSQL = "zl_保险参数_Delete(" & mlng险类 & ",null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Delete(" & mlng险类 & "," & mlng中心 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '新增参数数据
    colPara.Add mlng中心 & ",'卡号长度','" & Int(Val(TxtEdit(Text卡号长度).Text))
    colPara.Add mlng中心 & ",'退休证长度','" & Int(Val(TxtEdit(Text退休证号).Text))
    
    For lngCount = 1 To colPara.Count
        gstrSQL = "zl_保险参数_Insert(" & mlng险类 & "," & colPara(lngCount) & "'," & lngCount & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Next
    
    gcnOracle.CommitTrans
    mblnChange = False
    mblnOK = True
    Unload Me
    Exit Sub

errHandle:
    If ErrCenter() = 1 Then Resume
    gcnOracle.RollbackTrans
    Call SaveErrLog
End Sub

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll TxtEdit(Index)
    zlCommFun.OpenIme False
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Public Function 参数设置() As Boolean
'功能：设置我们中联医保所需要的参数
    Dim rsTemp As New ADODB.Recordset
    Dim str参数值 As String
    
    mblnOK = False
    mlng险类 = TYPE_重庆中梁山
    mlng中心 = 0
    
    rsTemp.CursorLocation = adUseClient
    gstrSQL = "select 参数名,参数值 from 保险参数 where 险类=[1] and (中心 is null or 中心=[2])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng险类, mlng中心)
    
    Do Until rsTemp.EOF
        Select Case rsTemp("参数名")
            Case "卡号长度"
                TxtEdit(Text卡号长度).Text = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
            Case "退休证长度"
                TxtEdit(Text退休证号).Text = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
        End Select
        
        rsTemp.MoveNext
    Loop
    
    mblnChange = False
    frmSet中梁山.Show vbModal, frm医保类别
    参数设置 = mblnOK
End Function
