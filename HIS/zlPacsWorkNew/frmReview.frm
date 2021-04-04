VERSION 5.00
Begin VB.Form frmReview 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "附加信息"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5655
   Icon            =   "frmReview.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   5655
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox cboDiagnosisType 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1080
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   2520
      Width           =   4455
   End
   Begin VB.CommandButton cmdCancelReview 
      Caption         =   "取消随访"
      Height          =   350
      Left            =   3240
      TabIndex        =   4
      Top             =   3120
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
      Height          =   350
      Left            =   4440
      TabIndex        =   3
      Top             =   3120
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   350
      Left            =   2040
      TabIndex        =   2
      Top             =   3120
      Width           =   1100
   End
   Begin VB.TextBox txtReview 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   480
      Width           =   5415
   End
   Begin VB.Label Label2 
      Caption         =   "诊断分类"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2565
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "随访描述"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmReview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngOrderID As Long     '医嘱ID
Private mlngSendNo As Long      '发送号
Private mstrReview As String    '随访描述
Private mModifyReview As Boolean '是否修改随访描述

Public Function ShowMe(lngOrderID As Long, lngSendNO As Long, frmParent As Object, _
    strDeptName As String, strReview As String) As Boolean
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    mlngOrderID = lngOrderID
    mlngSendNo = lngSendNO
    
    Me.cboDiagnosisType.Clear
    strSQL = "select 名称 from 影像诊断分类 Where 科室名称= [1] order by 编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "选择诊断分类", strDeptName)
    While Not rsTemp.EOF
        Me.cboDiagnosisType.AddItem rsTemp!名称
        rsTemp.MoveNext
    Wend
    
    strSQL = "Select 随访描述,诊断分类 From 影像检查记录 Where 医嘱id=[1] And 发送号 = [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngOrderID, mlngSendNo)
    
    If Not rsTemp.EOF Then
        Me.txtReview.Text = Nvl(rsTemp!随访描述)
        Me.cboDiagnosisType.Text = Nvl(rsTemp!诊断分类)
    Else
        Me.txtReview.Text = ""
        Me.cboDiagnosisType.Text = ""
    End If
    
    Me.Show 1, frmParent
    
    strReview = mstrReview
    ShowMe = mModifyReview
End Function

Private Sub cmdCancel_Click()
    mModifyReview = False
    Unload Me
End Sub

Private Sub cmdCancelReview_Click()
    Dim strSQL As String
    If MsgBoxD(Me, "是否清空随访记录？", vbOKCancel) = vbOK Then
        strSQL = "Zl_影像随访_Update(" & mlngOrderID & ",'')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        mstrReview = ""
        mModifyReview = True
        
        Unload Me
   End If
End Sub

Private Sub cmdOK_Click()
    Dim strSQL As String
    Dim arrSQL() As Variant
    Dim blnInTrans As Boolean
    Dim i As Integer
    
    arrSQL = Array()
    
    On Error GoTo errHandle
    
    strSQL = "Zl_影像诊断分类_Update(" & mlngOrderID & ",'" & cboDiagnosisType.Text & "')"
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = strSQL
    
    strSQL = "Zl_影像随访_Update(" & mlngOrderID & ",'" & txtReview.Text & "')"
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = strSQL
        
    gcnOracle.BeginTrans        '----------更新诊断分类和随访
    blnInTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "更新诊断分类和随访")
    Next i
    gcnOracle.CommitTrans
    blnInTrans = False
    
    mstrReview = Me.txtReview.Text
    mModifyReview = True
    
    Unload Me
    
    Exit Sub
errHandle:
    If blnInTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
        Call SaveErrLog
End Sub

