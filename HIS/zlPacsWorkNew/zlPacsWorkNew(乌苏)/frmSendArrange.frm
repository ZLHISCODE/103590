VERSION 5.00
Begin VB.Form frmSendArrange 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "发送安排"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5040
   Icon            =   "frmSendArrange.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton CmdCancle 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   2520
      Width           =   1125
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "确定(&O)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      ToolTipText     =   "保存(F2)"
      Top             =   2520
      Width           =   1125
   End
   Begin VB.Frame Frame1 
      Height          =   25
      Left            =   0
      TabIndex        =   8
      Top             =   2280
      Width           =   5055
   End
   Begin VB.ComboBox cboRoom 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1680
      Width           =   2685
   End
   Begin VB.ComboBox cbo技师一 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   2685
   End
   Begin VB.ComboBox cbo技师二 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1020
      Width           =   2685
   End
   Begin VB.Label lblRoom 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "执  行  间"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   7
      Top             =   1750
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "检查技师二"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   6
      Top             =   1075
      Width           =   1425
   End
   Begin VB.Label Label30 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "检查技师一"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Top             =   400
      Width           =   1425
   End
End
Attribute VB_Name = "frmSendArrange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngCurDeptId As Long
Private mlngAdviceId As Long
Private mlngSendNo As Long

Public Sub ShowMe(objParent As Object, ByVal lngCurDeptId As Long, ByVal lngAdviceId As Long, ByVal lngSendNo As Long)
    mlngCurDeptId = lngCurDeptId
    mlngAdviceId = lngAdviceId
    mlngSendNo = lngSendNo
    
    Me.Show 1, objParent
End Sub

Private Sub CmdCancle_Click()
    Unload Me
    
End Sub

Private Sub CmdOK_Click()
On Error GoTo ErrorHand

    Dim strSql As String
    
    strSql = "ZL_影像检查记录_发送安排(" & mlngAdviceId & "," & mlngSendNo & ",1," & "'" & NeedName(cbo技师一.Text) & "','" & NeedName(cbo技师二.Text) & "','" & NeedNo(cboRoom.Text) & "')"
    Call zlDatabase.ExecuteProcedure(strSql, "发送安排")
    
    '保存本次的选择
    Call SaveSetting("ZLSOFT", "公共模块\" & App.ProductName, "检查技师一", cbo技师一.Text)
    Call SaveSetting("ZLSOFT", "公共模块\" & App.ProductName, "检查技师二", cbo技师二.Text)
    Call SaveSetting("ZLSOFT", "公共模块\" & App.ProductName, "执行间", cboRoom.Text)
    
    Unload Me
    
    Exit Sub
ErrorHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
On Error GoTo ErrorHand
    Dim str检查技师一 As String
    Dim str检查技师二 As String
    Dim strRoom As String

    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    str检查技师一 = GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "检查技师一")
    str检查技师二 = GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "检查技师二")
    strRoom = GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "执行间")
    
    '加载检查技师
    strSql = "Select " & vbNewLine & _
                "Distinct b.id,b.姓名, Upper(b.简码) As 简码" & vbNewLine & _
                " From 部门人员 a, 人员表 b " & vbNewLine & _
                " Where a.人员id = b.Id And " & vbNewLine & _
                "      (b.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or b.撤档时间 Is Null) and a.部门id = [1] " & vbNewLine & _
                " Order By 简码 Desc"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngCurDeptId)
    
    '加载检查技师一
    cbo技师一.Clear
    Do Until rsTmp.EOF
        cbo技师一.AddItem rsTmp!简码 & "-" & rsTmp!姓名
        
        If rsTmp!简码 & "-" & rsTmp!姓名 = str检查技师一 Then
            cbo技师一.ListIndex = cbo技师一.NewIndex
        End If
        
        If cbo技师一.ListIndex = -1 And rsTmp!ID = UserInfo.ID Then
            cbo技师一.ListIndex = cbo技师一.NewIndex
        End If
        
        rsTmp.MoveNext
    Loop
    If cbo技师一.ListCount > 0 And cbo技师一.ListIndex = -1 Then cbo技师一.ListIndex = 0
    
    '加载检查技师二
    cbo技师二.Clear
    If rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
        Do Until rsTmp.EOF
            cbo技师二.AddItem rsTmp!简码 & "-" & rsTmp!姓名
            
            If rsTmp!简码 & "-" & rsTmp!姓名 = str检查技师二 Then
                cbo技师二.ListIndex = cbo技师二.NewIndex
            End If
            
            If cbo技师二.ListIndex = -1 And rsTmp!ID = UserInfo.ID Then
                cbo技师二.ListIndex = cbo技师二.NewIndex
            End If
            
            rsTmp.MoveNext
        Loop
        
        If cbo技师二.ListCount > 0 And cbo技师二.ListIndex = -1 Then cbo技师二.ListIndex = 0
    End If
    
    '执行间
    strSql = "Select 执行间,检查设备 From 医技执行房间 Where 科室id = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngCurDeptId)
    
    cboRoom.Clear
    Do While Not rsTmp.EOF
        cboRoom.AddItem rsTmp!执行间 & "-" & Nvl(rsTmp!检查设备)
        
        If Nvl(rsTmp!执行间) & "-" & Nvl(rsTmp!检查设备) = strRoom Then
            cboRoom.ListIndex = cboRoom.NewIndex
        End If
        
        rsTmp.MoveNext
    Loop
    If cboRoom.ListCount > 0 And cboRoom.ListIndex = -1 Then cboRoom.ListIndex = 0

    Exit Sub
ErrorHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
