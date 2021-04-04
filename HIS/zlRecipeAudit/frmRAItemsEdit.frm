VERSION 5.00
Begin VB.Form frmRAItemsEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "自定义审查项目"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7110
   Icon            =   "frmRAItemsEdit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   360
      Left            =   4200
      TabIndex        =   8
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   360
      Left            =   5520
      TabIndex        =   9
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CheckBox chkAdd 
      Caption         =   "连续新增自定义审查项目(&T)"
      Height          =   180
      Left            =   240
      TabIndex        =   7
      Top             =   2950
      Width           =   2775
   End
   Begin VB.Frame fraSplit 
      Height          =   30
      Left            =   120
      TabIndex        =   6
      Top             =   2640
      Width           =   6855
   End
   Begin VB.TextBox txtContent 
      Appearance      =   0  'Flat
      Height          =   1215
      Left            =   1800
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1200
      Width           =   4815
   End
   Begin VB.TextBox txtSimName 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   1800
      TabIndex        =   3
      Top             =   720
      Width           =   4815
   End
   Begin VB.TextBox txtCode 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label lblContent 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "内容描述(&M)"
      Height          =   180
      Left            =   600
      TabIndex        =   4
      Top             =   1230
      Width           =   990
   End
   Begin VB.Label lblSimName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "简称(&I)"
      Height          =   180
      Left            =   960
      TabIndex        =   2
      Top             =   750
      Width           =   630
   End
   Begin VB.Label lblCode 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "编码(&C)"
      Height          =   180
      Left            =   960
      TabIndex        =   0
      Top             =   270
      Width           =   630
   End
End
Attribute VB_Name = "frmRAItemsEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbytMode As Byte                '窗体模式；1-新增；2-编辑
Private mlngID As Long
'Private mblnOutPati As Boolean          '门诊启用；编辑模式下用到
'Private mblnInPati As Boolean           '住院启用；编辑模式下用到
Private mfrmOwner As Form

Public Sub ShowMe(ByVal bytMode As Byte, ByVal lngID As Long, ByVal frmOwner As Form)
'功能：上层代码显示本窗体的接口
'参数：
'  bytMode：窗体模式；1-新增；2-编辑
'  lngID：自定义项目的ID值，可以不传入，表示新增模式
'  frmOwner：宿主窗体对象

    If bytMode < 1 Or bytMode > 2 Then
        MsgBox "窗体模式参数不正确！", vbInformation, gstrSysName
        Exit Sub
    End If

    mbytMode = bytMode
    mlngID = lngID
    Set mfrmOwner = frmOwner
    
    InitCard
    
    Show vbModal, frmOwner

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

    '逻辑检查
    If Validate() = False Then Exit Sub
    '保存
    If Save() = False Then Exit Sub
    
    If mbytMode = 1 And chkAdd.Value Then
        '连续新增
        txtCode.Text = ""
        txtSimName.Text = ""
        txtContent.Text = ""
        txtCode.SetFocus
    Else
        Unload Me
    End If
    
End Sub

Private Function Validate() As Boolean
'功能：保存前的逻辑验证
'返回：True成功；False失败

    If Trim(txtCode.Text) = "" Then
        MsgBox "“编码”未填写！", vbInformation, gstrSysName
        txtCode.SetFocus
        Exit Function
    End If
    If Trim(txtSimName.Text) = "" Then
        MsgBox "“简称”未填写！", vbInformation, gstrSysName
        txtSimName.SetFocus
        Exit Function
    End If
    If Trim(txtContent.Text) = "" Then
        MsgBox "“内容描述”未填写！", vbInformation, gstrSysName
        txtContent.SetFocus
        Exit Function
    End If

    If Len(txtCode.Text) > txtCode.MaxLength Then
        MsgBox FormatEx("“编码”超长，最多能输入[1]个汉字或[2]个字符！", txtCode.MaxLength \ 2, txtCode.MaxLength), vbInformation, gstrSysName
        txtCode.SetFocus
        Exit Function
    End If
    
    If Len(txtSimName.Text) > txtSimName.MaxLength Then
        MsgBox FormatEx("“简称”超长，最多能输入[1]个汉字或[2]个字符！", txtSimName.MaxLength \ 2, txtSimName.MaxLength), vbInformation, gstrSysName
        txtSimName.SetFocus
        Exit Function
    End If
    
    If Len(txtContent.Text) > txtContent.MaxLength Then
        MsgBox FormatEx("“内容描述”超长，最多能输入[1]个汉字或[2]个字符！", txtContent.MaxLength \ 2, txtContent.MaxLength), vbInformation, gstrSysName
        txtContent.SetFocus
        Exit Function
    End If

    Validate = True
    
End Function

Private Function Save() As Boolean
'功能：保存数据
'返回：True成功；False失败

    Dim rsTmp As ADODB.Recordset
    Dim lngRow As Long

    On Error GoTo errHandle
    If mbytMode = 1 Then
        '新增时，获取ID值
        gstrSQL = "Select 处方审查项目_ID.Nextval as ID From Dual"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "获取处方审查项目新ID")
        If rsTmp.EOF = False Then
            mlngID = rsTmp!ID
        End If
        rsTmp.Close
    End If
    
    With mfrmOwner.vsfItems
        .Redraw = False
        
        If mbytMode = 1 Then
            '新增
            .Rows = .Rows + 1
            lngRow = .Rows - 1
            .TextMatrix(lngRow, .ColIndex("新增")) = "1"
        Else
            '编辑
            lngRow = .Row
        End If
        .TextMatrix(lngRow, .ColIndex("ID")) = CStr(mlngID)
        .TextMatrix(lngRow, .ColIndex("类别")) = "4-自定义"
        .TextMatrix(lngRow, .ColIndex("编码")) = txtCode.Text
        .TextMatrix(lngRow, .ColIndex("简称")) = txtSimName.Text
        .TextMatrix(lngRow, .ColIndex("内容描述")) = txtContent.Text
        .TextMatrix(lngRow, .ColIndex("服务对象")) = "2"
        
        .Redraw = True
    End With
    
    Save = True
    Exit Function
    
errHandle:
    If zl9ComLib.ErrCenter = 1 Then Resume
End Function

Private Sub InitCard()
'功能：初始化编辑卡片

    Dim rsTmp As ADODB.Recordset
    
    '设置TextBox的MaxLength
    SetTextMaxLen txtCode, "处方审查项目.编码"
    SetTextMaxLen txtSimName, "处方审查项目.简称"
    SetTextMaxLen txtContent, "处方审查项目.内容"
    
    If mbytMode = 1 Then Exit Sub          '新增不初始化
    
    chkAdd.Visible = False
    
    On Error GoTo errHandle
    gstrSQL = "Select 编码, 简称, 内容, 是否门诊启用, 是否住院启用 From 处方审查项目 Where ID = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "获取处方审查项目", mlngID)
    If rsTmp.EOF = False Then
        '编辑数据库的项目
        With rsTmp
            '加载数据
            txtCode.Text = !编码
            txtSimName.Text = !简称
            txtContent.Text = !内容
'            mblnOutPati = Val(zlCommFun.NVL(!是否门诊启用)) = 1
'            mblnInPati = Val(zlCommFun.NVL(!是否住院启用)) = 1
        End With
    Else
        '编辑新增未保存到数据库的项目
        With mfrmOwner.vsfItems
            '加载数据
            txtCode.Text = .TextMatrix(.Row, .ColIndex("编码"))
            txtSimName.Text = .TextMatrix(.Row, .ColIndex("简称"))
            txtContent.Text = .TextMatrix(.Row, .ColIndex("内容描述"))
'            mblnOutPati = .TextMatrix(.Row, .ColIndex("审查门诊"))
'            mblnInPati = .TextMatrix(.Row, .ColIndex("审查住院"))
        End With
    End If
    rsTmp.Close

    Exit Sub
    
errHandle:
    If zl9ComLib.ErrCenter = 1 Then Resume
End Sub

Private Sub txtCode_GotFocus()
    Call zlControl.TxtSelAll(txtCode)
End Sub

Private Sub txtCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    If InStr("~`!@#$%^&*()+={[}]|\:;""'<,>.?/", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtContent_GotFocus()
    Call zlControl.TxtSelAll(txtContent)
End Sub

Private Sub txtContent_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub txtContent_KeyPress(KeyAscii As Integer)
    If InStr("""'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtSimName_GotFocus()
    Call zlControl.TxtSelAll(txtSimName)
End Sub

Private Sub txtSimName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub txtSimName_KeyPress(KeyAscii As Integer)
    If InStr("~`!@#$%^&*()+={[}]|\:;""'<,>.?/", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub
