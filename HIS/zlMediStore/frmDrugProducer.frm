VERSION 5.00
Begin VB.Form frmDrugProducer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "药品生产商"
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7545
   Icon            =   "frmDrugProducer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   7545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   375
      Left            =   6360
      TabIndex        =   8
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   375
      Left            =   6360
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
   Begin VB.Frame frmLine 
      Height          =   1575
      Left            =   6000
      TabIndex        =   6
      Top             =   -120
      Width           =   30
   End
   Begin VB.TextBox txtProducer 
      Enabled         =   0   'False
      Height          =   270
      Index           =   2
      Left            =   1080
      TabIndex        =   5
      Top             =   915
      Width           =   4695
   End
   Begin VB.TextBox txtProducer 
      Height          =   270
      Index           =   1
      Left            =   1080
      MaxLength       =   60
      TabIndex        =   4
      Top             =   555
      Width           =   4695
   End
   Begin VB.TextBox txtProducer 
      Enabled         =   0   'False
      Height          =   270
      Index           =   0
      Left            =   1080
      TabIndex        =   3
      Top             =   195
      Width           =   1215
   End
   Begin VB.Label lblProducer 
      AutoSize        =   -1  'True
      Caption         =   "简码(&3)"
      Height          =   180
      Index           =   2
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   630
   End
   Begin VB.Label lblProducer 
      AutoSize        =   -1  'True
      Caption         =   "名称(&2)"
      Height          =   180
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   630
   End
   Begin VB.Label lblProducer 
      AutoSize        =   -1  'True
      Caption         =   "编码(&1)"
      Height          =   180
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   630
   End
End
Attribute VB_Name = "frmDrugProducer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim rsMaxs As New Recordset
    Dim ints编码 As Integer, strCodes As String, strSpecifys As String
    
    On Error GoTo errHandle
    
    If Trim(txtProducer(1).Text) = "" Then
        MsgBox "名称不能为空，请检查！", vbExclamation, gstrSysName
        txtProducer(1).SetFocus
        Exit Sub
    End If
    
    If LenB(StrConv(txtProducer(1).Text, vbFromUnicode)) > txtProducer(1).MaxLength Then
        MsgBox "所输入内容不能超过" & Int(txtProducer(1).MaxLength / 2) & "个汉字或" & txtProducer(1).MaxLength & "个字符!", vbExclamation + vbOKOnly, gstrSysName
        txtProducer(1).SetFocus
        Exit Sub
    End If
    
    '保存
    gstrSQL = "ZL_药品生产商_INSERT('" & txtProducer(0).Text & "','" & txtProducer(1).Text & "',substr('" & txtProducer(2).Text & "',0,10))"
    Call zldatabase.ExecuteProcedure(gstrSQL, "")
    
    '刷新界面，可以再次新增
    gstrSQL = "SELECT Nvl(MAX(LENGTH(编码)),2) As Length FROM 药品生产商"
    Set rsMaxs = zldatabase.OpenSQLRecord(gstrSQL, "" & "-药品生产商编码长度")
    ints编码 = rsMaxs!length
    
    gstrSQL = "SELECT Nvl(MAX(LPAD(编码," & ints编码 & ",'0')),'00') As Code FROM 药品生产商"
    Set rsMaxs = zldatabase.OpenSQLRecord(gstrSQL, "" & "-药品生产商编码")
    strCodes = rsMaxs!Code
    
    ints编码 = Len(strCodes)
    strCodes = strCodes + 1
    If ints编码 >= Len(strCodes) Then
        strCodes = String(ints编码 - Len(strCodes), "0") & strCodes
    End If
    
    txtProducer(0).Text = strCodes
    txtProducer(1).Text = ""
    txtProducer(2).Text = ""
    
    txtProducer(1).SetFocus
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Dim rsMaxs As New Recordset
    Dim ints编码 As Integer, strCodes As String, strSpecifys As String

    On Error GoTo errHandle
    
    Call GetDefineSize
    
    gstrSQL = "SELECT Nvl(MAX(LENGTH(编码)),2) As Length FROM 药品生产商"
    Set rsMaxs = zldatabase.OpenSQLRecord(gstrSQL, "" & "-药品生产商编码长度")
    ints编码 = rsMaxs!length
    
    gstrSQL = "SELECT Nvl(MAX(LPAD(编码," & ints编码 & ",'0')),'00') As Code FROM 药品生产商"
    Set rsMaxs = zldatabase.OpenSQLRecord(gstrSQL, "" & "-药品生产商编码")
    strCodes = rsMaxs!Code
    
    ints编码 = Len(strCodes)
    strCodes = strCodes + 1
    If ints编码 >= Len(strCodes) Then
        strCodes = String(ints编码 - Len(strCodes), "0") & strCodes
    End If
    
    txtProducer(0).Text = strCodes
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtProducer_Change(Index As Integer)
    txtProducer(2).Text = zlStr.GetCodeByVB(txtProducer(1).Text)
End Sub

Private Sub txtProducer_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 3 Then
        If InStr("0123456789", Chr(KeyAscii)) > 0 Or KeyAscii = 8 Then
            KeyAscii = KeyAscii
        Else
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub GetDefineSize()
    '功能：得到数据库的表字段的长度
    On Error GoTo errHandle
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
     
    strSQL = "Select t.上次产地 as 生产商 From 药品规格 T Where Rownum < 1"
    Call zldatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    
    txtProducer(1).MaxLength = rsTmp.Fields("生产商").DefinedSize
   
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
