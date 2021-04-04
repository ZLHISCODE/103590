VERSION 5.00
Begin VB.Form frmCommenLogSetDetailEdit 
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "日志记录规则"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5520
   Icon            =   "frmCommenLogSetDetailEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox cboStation 
      Height          =   300
      Left            =   1200
      TabIndex        =   2
      Top             =   1875
      Width           =   2295
   End
   Begin VB.ComboBox cboIP 
      Height          =   300
      Left            =   1200
      TabIndex        =   1
      Top             =   1425
      Width           =   2295
   End
   Begin VB.ComboBox cboName 
      Height          =   300
      Left            =   1200
      TabIndex        =   0
      Top             =   990
      Width           =   2295
   End
   Begin VB.TextBox txtCalls 
      Height          =   320
      Left            =   1200
      MaxLength       =   2000
      TabIndex        =   6
      Top             =   3660
      Width           =   3975
   End
   Begin VB.TextBox txtFunctions 
      Height          =   320
      Left            =   1200
      MaxLength       =   2000
      TabIndex        =   5
      Top             =   3210
      Width           =   3975
   End
   Begin VB.TextBox txtModules 
      Height          =   320
      Left            =   1200
      MaxLength       =   2000
      TabIndex        =   4
      Top             =   2760
      Width           =   3975
   End
   Begin VB.TextBox txtComponents 
      Height          =   320
      Left            =   1200
      MaxLength       =   2000
      TabIndex        =   3
      Top             =   2310
      Width           =   3975
   End
   Begin VB.Frame fraEnd 
      Height          =   45
      Index           =   1
      Left            =   -90
      TabIndex        =   11
      Top             =   720
      Width           =   6195
   End
   Begin VB.Frame fraEnd 
      Height          =   45
      Index           =   0
      Left            =   0
      TabIndex        =   9
      Top             =   4200
      Width           =   6075
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4005
      TabIndex        =   8
      Top             =   4365
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "保存(&S)"
      Height          =   350
      Left            =   2745
      TabIndex        =   7
      Top             =   4365
      Width           =   1100
   End
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   645
      Left            =   0
      ScaleHeight     =   645
      ScaleWidth      =   5520
      TabIndex        =   10
      Top             =   0
      Width           =   5520
      Begin VB.Label lblEXP 
         BackStyle       =   0  'Transparent
         Caption         =   "设置在哪些条件下记录日志，未设置时不限制。          多个部件、模块、功能、服务之间以半角分号（;）分隔。"
         Height          =   660
         Index           =   2
         Left            =   240
         TabIndex        =   19
         Top             =   120
         Width           =   4770
      End
   End
   Begin VB.Label lblCalls 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "适用服务"
      Height          =   180
      Left            =   360
      TabIndex        =   18
      Top             =   3735
      Width           =   720
   End
   Begin VB.Label lblFunctions 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "适用功能"
      Height          =   180
      Left            =   360
      TabIndex        =   17
      Top             =   3285
      Width           =   720
   End
   Begin VB.Label lblModules 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "适用模块"
      Height          =   180
      Left            =   360
      TabIndex        =   16
      Top             =   2835
      Width           =   720
   End
   Begin VB.Label lblComponents 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "适用部件"
      Height          =   180
      Left            =   360
      TabIndex        =   15
      Top             =   2385
      Width           =   720
   End
   Begin VB.Label lblStation 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "适用客户端"
      Height          =   180
      Left            =   180
      TabIndex        =   14
      Top             =   1935
      Width           =   900
   End
   Begin VB.Label lblIP 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "适用IP地址"
      Height          =   180
      Left            =   180
      TabIndex        =   13
      Top             =   1485
      Width           =   900
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "适用用户"
      Height          =   180
      Left            =   360
      TabIndex        =   12
      Top             =   1050
      Width           =   720
   End
End
Attribute VB_Name = "frmCommenLogSetDetailEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnOK          As Boolean
Private mlngId          As Long
Private mlngCategoryID  As Long
Private mcnOracle       As ADODB.Connection

Public Function ShowMe(ByRef cnOracle As ADODB.Connection, ByVal lngCategoryID As Long, Optional ByVal lngID As Long) As Boolean
    mblnOK = False
    mlngId = lngID
    mlngCategoryID = lngCategoryID
    Set mcnOracle = cnOracle
    Me.Show vbModal, frmMDIMain
    ShowMe = mblnOK
End Function

Private Sub cboIP_KeyPress(KeyAscii As Integer)
    If Not (Chr(KeyAscii) Like "#" Or Chr(KeyAscii) = ".") Then
        If Chr(KeyAscii) <> vbBack Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub cboName_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strSQL As String, strVal As String
    Dim arrTmp As Variant
    Dim i As Long, lngID As Long
    
    On Error GoTo errH
    
    strVal = cboName.Text
    strVal = strVal & cboIP.Text
    strVal = strVal & cboStation.Text
    strVal = strVal & txtComponents.Text
    strVal = strVal & txtModules.Text
    strVal = strVal & txtFunctions.Text
    strVal = strVal & txtCalls.Text
    If Trim$(strVal) = "" Then
        MsgBox "未填写数据！", vbInformation, gstrSysName
        cboName.SetFocus
        Exit Sub
    End If
        
    strVal = CheckIP("请输入有效的日志客户端IP地址！", cboIP.Text)
    If strVal <> "" Then
        MsgBox strVal, vbInformation, gstrSysName
        cboIP.SetFocus
        Exit Sub
    End If
        
    strSQL = "Zllogset_Edit(" & IIf(mlngId = 0, 0, 1) & _
        ", " & mlngId & _
        ", '" & UCase(Trim(cboName.Text)) & "'" & _
        ", '" & Trim(cboStation.Text) & "'" & _
        ", '" & Trim(cboIP.Text) & "'" & _
        ", "
    strSQL = strSQL & SQLAdjust(UCase(Trim(txtComponents.Text))) & _
        ", " & SQLAdjust(UCase(Trim(txtModules.Text))) & _
        ", " & SQLAdjust(UCase(Trim(txtFunctions.Text))) & _
        ", " & SQLAdjust(UCase(Trim(txtCalls.Text))) & _
        ", " & mlngCategoryID & _
        ")"
        
    Call ExecuteProcedure(strSQL, Me.Caption, mcnOracle)
    mblnOK = True
    Unload Me
    Exit Sub
    
errH:
    MsgBox "保存日志适用条件限制出错：" & err.Description, vbInformation, gstrSysName
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    Dim strSQL          As String, rsTmp            As ADODB.Recordset
    Dim dtCurrent       As Date
    
    If mlngId = 0 Then
        Me.Caption = Me.Caption & "(新增)"
    Else
        Me.Caption = Me.Caption & "(修改)"
    End If

    On Error GoTo errH
    
    lblName.Tag = "1"
    If mlngId > 0 Then
        strSQL = _
            "Select ID, Category_Id, User_Name, Station, Ip, Component_Names, Module_Names, Function_Names" & vbCr & _
            "  , Call_Names" & vbCr & _
            "From ZllogSet " & vbCr & _
            "Where Id = [1]"
        Set rsTmp = gclsBase.OpenSQLRecord(mcnOracle, strSQL, Me.Caption, mlngId)
        
        cboName.Text = rsTmp!User_Name & ""
        cboStation.Text = rsTmp!Station & ""
        cboIP.Text = rsTmp!IP & ""
        txtComponents.Text = rsTmp!Component_Names & ""
        txtModules.Text = rsTmp!Module_Names & ""
        txtFunctions.Text = rsTmp!Function_Names & ""
        txtCalls.Text = rsTmp!Call_Names & ""
    End If
    
    '使用当前库的用户，因为当前库是业务库
    On Error Resume Next
    strSQL = "Select Distinct 用户名 From 上机人员表 Order By 1"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
    cboName.addItem ""
    If Not rsTmp Is Nothing Then
        Do While Not rsTmp.EOF
            cboName.addItem rsTmp!用户名
            rsTmp.MoveNext
        Loop
    End If
    If mlngId = 0 Or cboName.Text = "" Then cboName.ListIndex = 0
    
    '使用当前库的客户端，因为当前库是业务库
    On Error Resume Next
    strSQL = "Select Distinct Trim(工作站) 工作站 From zlClients Order By 1"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
    cboStation.addItem ""
    If Not rsTmp Is Nothing Then
        Do While Not rsTmp.EOF
            cboStation.addItem rsTmp!工作站
            rsTmp.MoveNext
        Loop
    End If
    If mlngId = 0 Or cboStation.Text = "" Then cboStation.ListIndex = 0
    
    '使用当前库的IP，因为当前库是业务库
    On Error Resume Next
    strSQL = "Select Distinct Trim(IP) IP From zlClients Order By 1"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
    cboIP.addItem ""
    If Not rsTmp Is Nothing Then
        Do While Not rsTmp.EOF
            cboIP.addItem rsTmp!IP
            rsTmp.MoveNext
        Loop
    End If
    If mlngId = 0 Or cboIP.Text = "" Then cboIP.ListIndex = 0
    lblName.Tag = ""
    Exit Sub
    
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mcnOracle = Nothing
End Sub

Private Sub txtCalls_KeyPress(KeyAscii As Integer)
    If InStr("`~!@#$%^&*()=+|\{}[]':"",./<>?", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtComponents_KeyPress(KeyAscii As Integer)
    If InStr("`~!@#$%^&*()=+|\{}[]':"",./<>?", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtFunctions_KeyPress(KeyAscii As Integer)
    If InStr("`~!@#$%^&*()=+|\{}[]':"",./<>?", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtModules_KeyPress(KeyAscii As Integer)
    If InStr("`~!@#$%^&*()=+|\{}[]':"",./<>?", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    End If
End Sub
