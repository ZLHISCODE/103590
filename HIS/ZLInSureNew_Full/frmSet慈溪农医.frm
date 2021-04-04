VERSION 5.00
Begin VB.Form frmSet慈溪农医 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4050
   Icon            =   "frmSet慈溪农医.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   4050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtIC端口号 
      Height          =   300
      Left            =   1410
      MaxLength       =   1
      TabIndex        =   6
      Text            =   "1"
      Top             =   1560
      Width           =   465
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   2790
      TabIndex        =   9
      Top             =   2100
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   1560
      TabIndex        =   8
      Top             =   2100
      Width           =   1100
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "测试(&T)"
      Height          =   350
      Left            =   180
      TabIndex        =   7
      Top             =   2100
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Caption         =   "前置机IP及端口设置"
      Height          =   1275
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.TextBox txt端口号 
         Height          =   300
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   4
         Text            =   "8801"
         Top             =   750
         Width           =   495
      End
      Begin VB.TextBox txtIP地址 
         Height          =   300
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   2
         Text            =   "192.168.168.168"
         Top             =   360
         Width           =   1485
      End
      Begin VB.Label lbl端口号 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "端口号(&P)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   540
         TabIndex        =   3
         Top             =   810
         Width           =   810
      End
      Begin VB.Label lblIP地址 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "IP地址(&A)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   540
         TabIndex        =   1
         Top             =   420
         Width           =   810
      End
   End
   Begin VB.Label lblIC端口号 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "IC端口号(&R)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   270
      TabIndex        =   5
      Top             =   1620
      Width           =   990
   End
End
Attribute VB_Name = "frmSet慈溪农医"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnReturn As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHand
    If Trim(txtIC端口号.Text) = "" Then
        MsgBox "端口号不能为空！", vbInformation, gstrSysName
        txt端口号.SetFocus
        Exit Sub
    End If
    If Val(txtIC端口号.Text) < 1 Or Val(txtIC端口号.Text) > 5 Then
        MsgBox "端口号不能小于1或者大于5", vbInformation, gstrSysName
        txt端口号.SetFocus
        Exit Sub
    End If
    
    gcnOracle.BeginTrans
    
    '删除已经数据
    gstrSQL = "zl_保险参数_Delete(" & TYPE_慈溪农医 & ",null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '新增参数数据
    gstrSQL = "zl_保险参数_Insert(" & TYPE_慈溪农医 & ",NULL,'IP地址','" & txtIP地址.Text & "',1)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_慈溪农医 & ",NULL,'端口号','" & txt端口号.Text & "',2)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_慈溪农医 & ",NULL,'IC端口号','" & txtIC端口号.Text & "',3)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    gcnOracle.CommitTrans
    mblnReturn = True
    
    Unload Me
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
    gcnOracle.RollbackTrans
End Sub

Private Sub cmdTest_Click()
    Dim lngPort As Long
    Dim strIP As String
    '测试是否连接的通
    strIP = txtIP地址.Text
    lngPort = Val(txt端口号.Text)
    Call CXNY_SetRemoteServerAddr(lngPort, strIP)
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "Select 参数名,参数值 From 保险参数 Where 险类=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取保险参数", TYPE_慈溪农医)
    
    With rsTemp
        Do While Not .EOF
            Select Case !参数名
            Case "IP地址"
                txtIP地址.Text = Nvl(!参数值, "127.0.0.1")
            Case "端口号"
                txt端口号.Text = Nvl(!参数值, "8801")
            Case "IC端口号"
                txtIC端口号.Text = Nvl(!参数值, 1)
            End Select
            .MoveNext
        Loop
    End With
End Sub

Private Sub txtIP地址_GotFocus()
    txtIP地址.SelStart = 0
    txtIP地址.SelLength = Len(txtIP地址.Text)
End Sub

Private Sub txtIP地址_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Or KeyAscii = 46) Then KeyAscii = 0
End Sub

Private Sub txtIP地址_Validate(Cancel As Boolean)
    Dim arrIP
    Dim intCOUNT As Integer, intDO As Integer
    '检查IP输入是否合法
    If Trim(txtIP地址.Text) = "" Then Exit Sub
    
    On Error GoTo errHand
    arrIP = Split(txtIP地址.Text, ".")
    intCOUNT = UBound(arrIP)
    If intCOUNT > 3 Then
        MsgBox "请输入合法的IP地址！", vbInformation, gstrSysName
        Cancel = True
        Exit Sub
    End If
    
    For intDO = 0 To 3
        If Val(arrIP(intDO)) > 255 Then
            MsgBox "请输入合法的IP地址！", vbInformation, gstrSysName
            Cancel = True
            Exit Sub
        End If
    Next
    Exit Sub
errHand:
    Cancel = True
    MsgBox "请输入合法的IP地址！", vbInformation, gstrSysName
End Sub

Private Sub Txt端口号_GotFocus()
    txt端口号.SelStart = 0
    txt端口号.SelLength = Len(txt端口号.Text)
End Sub

Private Sub txt端口号_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Public Function ShowME() As Boolean
    mblnReturn = False
    Me.Show 1
    ShowME = mblnReturn
End Function
