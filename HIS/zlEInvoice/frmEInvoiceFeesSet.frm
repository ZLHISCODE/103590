VERSION 5.00
Begin VB.Form frmEInvoiceFeeseSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "收据费目对照"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5100
   Icon            =   "frmEInvoiceFeesSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.OptionButton Option场合 
      Caption         =   "住院"
      Height          =   255
      Index           =   2
      Left            =   2760
      TabIndex        =   13
      Top             =   2040
      Width           =   700
   End
   Begin VB.CommandButton cmd费目 
      Caption         =   "…"
      Height          =   250
      Left            =   3120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   360
      Width           =   255
   End
   Begin VB.OptionButton Option场合 
      Caption         =   "门诊"
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   5
      Top             =   2040
      Width           =   700
   End
   Begin VB.OptionButton Option场合 
      Caption         =   "不区分"
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   4
      Top             =   2040
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   2
      Left            =   960
      TabIndex        =   3
      Top             =   1462
      Width           =   2475
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   1
      Left            =   960
      MaxLength       =   20
      TabIndex        =   2
      Top             =   920
      Width           =   2475
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   335
      Width           =   2115
   End
   Begin VB.Frame fra 
      Height          =   3400
      Left            =   3600
      TabIndex        =   8
      Top             =   -120
      Width           =   10
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3840
      TabIndex        =   7
      Top             =   840
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3840
      TabIndex        =   6
      Top             =   360
      Width           =   1100
   End
   Begin VB.Label lbl 
      Caption         =   "费用场合"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   12
      Top             =   2040
      Width           =   730
   End
   Begin VB.Label lbl 
      Caption         =   "名    称"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   1485
      Width           =   730
   End
   Begin VB.Label lbl 
      Caption         =   "编    码"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   943
      Width           =   730
   End
   Begin VB.Label lbl 
      Caption         =   "收据费目"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   358
      Width           =   730
   End
End
Attribute VB_Name = "frmEInvoiceFeeseSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum TXT_Idex
    Idex_费目 = 0
    Idex_编码 = 1
    Idex_名称 = 2
End Enum
Private mlngID As Long      '收据费目对照.ID，修改是传入，新增是为0
Private mblnOK As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If IsValid = False Then Exit Sub
    If Save收据费目对照 = False Then Exit Sub
    mblnOK = True
    Unload Me
End Sub

Private Function Save收据费目对照() As Boolean
    Dim strSQL As String
    Dim lngNewID As Long, int场合 As Integer
    
    On Error GoTo errHandle
    int场合 = IIf(Option场合(0).Value = True, 0, IIf(Option场合(1).Value = True, 1, 2))
    If mlngID = 0 Then
        '新增收据费目对照
        lngNewID = zlDatabase.GetNextId("收据费目对照")
        strSQL = "Zl_收据费目对照_Update("
'        操作类型_In In Number,
        strSQL = strSQL & 0 & ","
'        Id_In       In 收据费目对照.Id%Type,
        strSQL = strSQL & lngNewID & ","
'        收据费目_In In 收据费目对照.收据费目%Type,
        strSQL = strSQL & "'" & txtEdit(Idex_费目).Text & "',"
'        名称_In     In 收据费目对照.名称%Type,
        strSQL = strSQL & "'" & txtEdit(Idex_名称).Text & "',"
'        编码_In     In 收据费目对照.编码%Type,
        strSQL = strSQL & "'" & txtEdit(Idex_编码).Text & "',"
'        费用场合_In In 收据费目对照.费用场合%Type
        strSQL = strSQL & int场合 & ")"
    Else
        '修改收据费目对照
        strSQL = "Zl_收据费目对照_Update("
'        操作类型_In In Number,
        strSQL = strSQL & 1 & ","
'        Id_In       In 收据费目对照.Id%Type,
        strSQL = strSQL & mlngID & ","
'        收据费目_In In 收据费目对照.收据费目%Type,
        strSQL = strSQL & "'" & txtEdit(Idex_费目).Text & "',"
'        名称_In     In 收据费目对照.名称%Type,
        strSQL = strSQL & "'" & txtEdit(Idex_名称).Text & "',"
'        编码_In     In 收据费目对照.编码%Type,
        strSQL = strSQL & "'" & txtEdit(Idex_编码).Text & "',"
'        费用场合_In In 收据费目对照.费用场合%Type
        strSQL = strSQL & int场合 & ")"
    End If
    Call zlDatabase.ExecuteProcedure(strSQL, "收据费目对照")
    
    Save收据费目对照 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cmd费目_Click()
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim vRect As RECT
    
    strSQL = "Select Rownum As id, 编码, Upper(名称) as 名称,Upper(简码) as 简码 From 收据费目 Order  By 编码 "
    vRect = zlControl.GetControlRect(txtEdit(Idex_费目).hWnd)
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "获取收据费目", False, "", "", False, False, _
            True, vRect.Left, vRect.Top, txtEdit(Idex_费目).Height, True, False, False)
     If Not rsTemp Is Nothing Then
        txtEdit(Idex_费目).Text = rsTemp("名称")
    End If
    zlControl.ControlSetFocus txtEdit(Idex_费目)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':：;；?？", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0: Exit Sub
    End If
    
    If KeyAscii = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub Form_Load()
    Call Load收据费目FromID(mlngID)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mlngID = 0
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim vRect As RECT
    
    If Index = Idex_编码 Then
        If Not ((KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 8 Or KeyAscii = 13) Then KeyAscii = 0
    ElseIf Index = Idex_名称 Then
        If InStr("':：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    ElseIf Index = Idex_费目 Then
        If InStr("':：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0
        If KeyAscii <> vbKeyReturn Then Exit Sub
        strSQL = "Select Rownum As id, 编码, Upper(名称) as 名称,Upper(简码) as 简码 From 收据费目 " & _
                  "Where 编码 Like Upper([1]) Or Upper(名称) Like Upper([1]) Or Upper(简码)  Like Upper([1]) " & _
                  "   Or Upper(zlPinYinCode(名称)) Like Upper([1]) Order By 编码 "
                  
        vRect = zlControl.GetControlRect(txtEdit(Idex_费目).hWnd)
        Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "获取客户端", False, "", "", False, False, _
                True, vRect.Left, vRect.Top, txtEdit(Idex_费目).Height, True, False, False, "%" & txtEdit(Idex_费目).Text & "%")
         If Not rsTemp Is Nothing Then
            txtEdit(Idex_费目).Text = rsTemp("名称")
            zlControl.ControlSetFocus txtEdit(Idex_费目)
        Else
            MsgBox "根据输入的信息未找到有效的收据费目，请重试！", vbInformation, gstrSysName
            txtEdit(Idex_费目).Text = ""
            zlControl.ControlSetFocus txtEdit(Idex_费目)
        End If
    End If
End Sub

Public Sub ShowMe(ByVal frmMain As Object, Optional ByVal lngID As Long, Optional blnRefresh As Boolean)
    mlngID = lngID
    mblnOK = False
    Me.Show 1, frmMain
    blnRefresh = mblnOK
End Sub

Private Function IsValid() As Boolean
    On Error GoTo errHandle

    If Len(Trim(txtEdit(Idex_费目).Text)) = 0 Then
        MsgBox "收据费目不能为空。", vbExclamation, gstrSysName
        zlControl.ControlSetFocus txtEdit(Idex_费目)
        Exit Function
    End If
    
    If Len(txtEdit(Idex_编码).Text) = 0 Then
        MsgBox "编码不能为空。", vbExclamation, gstrSysName
        zlControl.ControlSetFocus txtEdit(Idex_编码)
        Exit Function
    End If

    If Not IsNumeric(txtEdit(Idex_编码).Text) Or InStr(txtEdit(Idex_编码).Text, ",") > 0 Or InStr(txtEdit(Idex_编码).Text, ".") > 0 Then
        MsgBox "编码应由数字组成。", vbExclamation, gstrSysName
        zlControl.ControlSetFocus txtEdit(Idex_编码)
        Exit Function
    End If
    
    If Len(Trim(txtEdit(Idex_名称).Text)) = 0 Then
        MsgBox "名称不能为空。", vbExclamation, gstrSysName
        txtEdit(Idex_名称).Text = ""
        zlControl.ControlSetFocus txtEdit(Idex_名称)
        Exit Function
    End If
    
    If LenB(StrConv(txtEdit(Idex_名称).Text, vbFromUnicode)) > 20 Then
        MsgBox "名称长度不能超过10个汉字或者20个字符，请重新录入！", vbInformation, gstrSysName
        zlControl.ControlSetFocus txtEdit(Idex_名称)
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

Private Sub Load收据费目FromID(ByVal lngID As Long)
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim int场合 As Integer
    
    If lngID = 0 Then Exit Sub
    strSQL = "Select 收据费目,名称,编码,费用场合 From 收据费目对照 where ID =[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngID)
    If rsTemp.EOF Then Exit Sub
    With rsTemp
        txtEdit(Idex_费目).Text = NVL(!收据费目)
        txtEdit(Idex_编码).Text = NVL(!编码)
        txtEdit(Idex_名称).Text = NVL(!名称)
        int场合 = Val(!费用场合)
        Option场合(int场合).Value = True
    End With
End Sub


