VERSION 5.00
Begin VB.Form frmOneCard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "一卡通配置"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5805
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   5805
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox cbo启用 
      Height          =   300
      Left            =   1200
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1890
      Width           =   2025
   End
   Begin VB.Frame Frame1 
      Height          =   4485
      Left            =   4320
      TabIndex        =   13
      Top             =   -120
      Width           =   30
   End
   Begin VB.TextBox txtNO 
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   300
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   1
      TabStop         =   0   'False
      Tag             =   "编码"
      Text            =   "11"
      Top             =   240
      Width           =   1125
   End
   Begin VB.TextBox txtOrgCode 
      Height          =   300
      Left            =   1200
      MaxLength       =   6
      TabIndex        =   7
      Tag             =   "简码"
      Top             =   1470
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4530
      TabIndex        =   10
      Top             =   240
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4530
      TabIndex        =   11
      Top             =   690
      Width           =   1100
   End
   Begin VB.TextBox txtName 
      Height          =   300
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   3
      Tag             =   "名称"
      Top             =   600
      Width           =   2925
   End
   Begin VB.ComboBox cboPayType 
      Height          =   300
      Left            =   1200
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1050
      Width           =   2025
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   4530
      TabIndex        =   12
      Top             =   1560
      Width           =   1100
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "启用(&E)"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   1950
      Width           =   630
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "医院编码(&O)"
      Height          =   180
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   1530
      Width           =   990
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "名称(&N)"
      Height          =   180
      Index           =   2
      Left            =   480
      TabIndex        =   2
      Top             =   660
      Width           =   630
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "编号(&U)"
      Height          =   180
      Index           =   1
      Left            =   480
      TabIndex        =   0
      Top             =   300
      Width           =   630
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "结算方式(&P)"
      Height          =   180
      Index           =   5
      Left            =   120
      TabIndex        =   4
      Top             =   1110
      Width           =   990
   End
End
Attribute VB_Name = "frmOneCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mbytInFun As Byte '0-新增,1-修改

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, "frmOneCard", Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    Dim strSQL As String
    
    If cboPayType.ListIndex = -1 Then
        MsgBox "请选择结算方式!", vbInformation, gstrSysName
        cboPayType.SetFocus
        Exit Sub
    End If
    If txtName.Text = "" Then
        MsgBox "请输入一卡通接口名称!", vbInformation, gstrSysName
        txtName.SetFocus
        Exit Sub
    End If
    If txtOrgCode.Text = "" Then
        MsgBox "请输入医院编码!", vbInformation, gstrSysName
        txtOrgCode.SetFocus
        Exit Sub
    End If
    
    If zlCommFun.ActualLen(txtName.Text) > txtName.MaxLength Then
        MsgBox "名称不能超过" & txtName.MaxLength & "个字符!", vbInformation, gstrSysName
        txtName.SetFocus
        Exit Sub
    End If
    If zlCommFun.ActualLen(txtOrgCode.Text) > txtOrgCode.MaxLength Then
        MsgBox "医院编码不能超过" & txtOrgCode.MaxLength & "个字符!", vbInformation, gstrSysName
        txtOrgCode.SetFocus
        Exit Sub
    End If

    '该操作一般为系统管理员进行,所以忽略并发控制, 由数据结构控制
    strSQL = "Zl_一卡通目录_Update(" & txtNO.Text & ",'" & txtName.Text & "','" & cboPayType.Text & "','" & _
            txtOrgCode.Text & "'," & cbo启用.ListIndex & "," & mbytInFun & ")"
    
    On Error GoTo errH
    Call zlDatabase.ExecuteProcedure(strSQL, App.ProductName)
    
    Unload Me
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    txtName.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Not Me.ActiveControl Is cmdOK Then Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Function GetPayType() As ADODB.Recordset
    Dim strSQL As String
 
    strSQL = "Select 名称,编码 From 结算方式 Where 性质=7"
    On Error GoTo errH
    Set GetPayType = zlDatabase.OpenSQLRecord(strSQL, App.ProductName)

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetOneCardMaxNO() As String
    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "Select Nvl(max(编号),0) 编号 From 一卡通目录"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName)
    GetOneCardMaxNO = Val(rsTmp!编号) + 1
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub ShowMe(objParent As Form, Optional intNO As Integer, Optional strName As String, _
    Optional strPayType As String, Optional strOrgCode As String, Optional intState As Integer)
    Dim rsTmp As ADODB.Recordset
    
    Set rsTmp = GetPayType
    If rsTmp.RecordCount = 0 Then
        MsgBox "没有找到性质为一卡通的结算方式,请先到[结算方式管理]中配置。", vbInformation
        Exit Sub
    End If
    Call zlControl.CboAddData(cboPayType, rsTmp, True)
    
    With Me.cbo启用
        .Clear
        .AddItem "停用", 0
        .AddItem "启用:仅涉及扣卡", 1
        .AddItem "启用:标准一卡通", 2
        .ListIndex = 0
    End With
    
    If mbytInFun = 0 Then
        cboPayType.ListIndex = 0
        txtNO.Text = GetOneCardMaxNO
    Else
        txtNO.Text = intNO
        txtName.Text = strName
        Call zlControl.CboLocate(cboPayType, strPayType)
        txtOrgCode.Text = strOrgCode
        cbo启用.ListIndex = intState
    End If
    
    Me.Show 1, objParent
End Sub

Public Sub DelOneCardRec(intNO As Integer)
    Dim strSQL As String
    
    strSQL = "Zl_一卡通目录_Update(" & intNO & ",null,null,null,null,2)"
    
    On Error GoTo errH
    Call zlDatabase.ExecuteProcedure(strSQL, App.ProductName)
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
    txtName.Text = Trim(txtName.Text)
End Sub

Private Sub txtOrgCode_Change()
    txtOrgCode.Text = Trim(txtOrgCode.Text)
End Sub
