VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmScheme_BaseQueryCfg 
   BorderStyle     =   0  'None
   Caption         =   "基础查询设置"
   ClientHeight    =   7620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   11340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraDetail 
      Height          =   30
      Left            =   2160
      TabIndex        =   8
      Top             =   3720
      Width           =   7215
   End
   Begin VB.Frame fraCheckSQL 
      Height          =   30
      Left            =   2160
      TabIndex        =   5
      Top             =   600
      Width           =   6255
   End
   Begin VB.CommandButton cmdSQLInsert 
      Caption         =   "插入参数"
      Height          =   350
      Left            =   8880
      TabIndex        =   3
      Top             =   960
      Width           =   1100
   End
   Begin VB.CommandButton cmdSQLVerify 
      Caption         =   "查询验证"
      Height          =   350
      Left            =   8880
      TabIndex        =   2
      Top             =   1560
      Width           =   1100
   End
   Begin VB.CommandButton cmdDetailInsert 
      Caption         =   "插入参数"
      Height          =   350
      Left            =   8880
      TabIndex        =   1
      Top             =   4200
      Width           =   1100
   End
   Begin VB.CommandButton cmdDetailVerify 
      Caption         =   "查询验证"
      Height          =   350
      Left            =   8880
      TabIndex        =   0
      Top             =   4800
      Width           =   1100
   End
   Begin RichTextLib.RichTextBox rtbCheckSQL 
      Height          =   2610
      Left            =   840
      TabIndex        =   4
      Top             =   720
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   4604
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmScheme_BaseQueryCfg.frx":0000
   End
   Begin RichTextLib.RichTextBox rtbDetail 
      Height          =   3135
      Left            =   840
      TabIndex        =   7
      Top             =   4080
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   5530
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmScheme_BaseQueryCfg.frx":009D
   End
   Begin VB.Label lblDetail 
      AutoSize        =   -1  'True
      Caption         =   "数据明细语句"
      Height          =   180
      Left            =   840
      TabIndex        =   9
      Top             =   3600
      Width           =   1080
   End
   Begin VB.Label lblCheckSQL 
      AutoSize        =   -1  'True
      Caption         =   "数据查询语句"
      Height          =   180
      Left            =   1080
      TabIndex        =   6
      Top             =   480
      Width           =   1080
   End
End
Attribute VB_Name = "frmScheme_BaseQueryCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public mblnIsEdit As Boolean

Private mstrQuerySql As String

Private Sub cmdDetailInsert_Click()
'插入参数
On Error GoTo errHandle
    Dim strPar As String
    Dim frmPar As New frmSetPara
    
    strPar = frmPar.ShowParameterWindow(False, Me)
    If strPar <> "" Then
        rtbDetail.SelText = strPar
    End If
    
Exit Sub
errHandle:
    MsgBox Err.Description, vbOKOnly, Me.Caption
End Sub

Private Sub cmdDetailVerify_Click()
    Dim strResult As String

    On Error GoTo errHandle
    
    strResult = SqlVerify(rtbDetail.Text)
    
    If Len(strResult) = 0 Then
        MsgBox "验证成功！", vbInformation, Me.Caption
    Else
        MsgBox "验证失败，原因是：" & strResult, vbInformation, Me.Caption
        rtbDetail.SetFocus
    End If
    
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub cmdSQLInsert_Click()
'插入参数
    Dim strPar As String
    Dim frmPar As New frmSetPara
    
    On Error GoTo errHandle
    
    strPar = frmPar.ShowParameterWindow(False, Me)
    If strPar <> "" Then
        rtbCheckSQL.SelText = strPar
    End If
    
    Set frmPar = Nothing
    Exit Sub
errHandle:
    MsgBox Err.Description, vbOKOnly, Me.Caption
End Sub

Private Sub cmdSQLVerify_Click()
    Dim strResult As String
    Dim objSqlParse As New clsSqlParse
    Dim objQuery As New clsPacsQuery
    Dim rsRecord As Recordset
    Dim strItem As String
    Dim strPar As String
    Dim i As Long
    
    On Error GoTo errHandle
    
    strResult = SqlVerify(rtbCheckSQL.Text)
    
    If Len(strResult) = 0 Then
        strResult = IsHaveID(rtbCheckSQL.Text)
    End If
    
    If Len(strResult) = 0 Then
        MsgBox "验证成功！", vbInformation, Me.Caption
    Else
        MsgBox "验证失败，原因是：" & strResult, vbInformation, Me.Caption
        rtbCheckSQL.SetFocus
    End If
    
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub Form_Load()
    Call RefreshWindowState(False)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    '查询语句部分
    lblCheckSQL.Move Me.ScaleLeft + 100, Me.ScaleTop + 100
    fraCheckSQL.Move lblCheckSQL.Left + lblCheckSQL.Width, lblCheckSQL.Top + lblCheckSQL.Height / 2, Me.ScaleWidth - fraCheckSQL.Left
        rtbCheckSQL.Move Me.ScaleLeft + 100, fraCheckSQL.Top + 200, Me.ScaleWidth - 300 - cmdSQLInsert.Width, (Me.ScaleHeight - rtbCheckSQL.Top * 2 - 300) / 2
    cmdSQLInsert.Move rtbCheckSQL.Left + rtbCheckSQL.Width + 100, rtbCheckSQL.Top
    cmdSQLVerify.Move cmdSQLInsert.Left, cmdSQLInsert.Top + cmdSQLInsert.Height + 200

    '检查明细部分
    lblDetail.Move lblCheckSQL.Left, rtbCheckSQL.Top + rtbCheckSQL.Height + 100
    fraDetail.Move lblDetail.Left + lblDetail.Width, lblDetail.Top + lblDetail.Height / 2, Me.ScaleWidth - fraDetail.Left
    rtbDetail.Move rtbCheckSQL.Left, fraDetail.Top + 200, rtbCheckSQL.Width, rtbCheckSQL.Height
    cmdDetailInsert.Move cmdSQLInsert.Left, rtbDetail.Top
    cmdDetailVerify.Move cmdSQLInsert.Left, cmdDetailInsert.Top + cmdDetailInsert.Height + 200
End Sub

Public Sub ShowQuerySet(objSqlScheme As clsSqlScheme)
    rtbCheckSQL.Text = objSqlScheme.Query
    rtbDetail.Text = objSqlScheme.Detail
End Sub

Private Sub rtbCheckSQL_Change()
    mstrQuerySql = rtbCheckSQL.Text
    mblnIsEdit = True
    
    If cmdSQLInsert.Enabled Then
        If Len(Trim(rtbCheckSQL.Text)) = 0 Then
            cmdSQLVerify.Enabled = False
        Else
            cmdSQLVerify.Enabled = True
        End If
    Else
        cmdSQLVerify.Enabled = False
    End If
End Sub

Public Function GetQuerySql() As String
    GetQuerySql = mstrQuerySql
End Function

Public Sub RefreshWindowState(blnState As Boolean)
    rtbCheckSQL.Locked = Not blnState
    rtbDetail.Locked = Not blnState
    
    If blnState Then
        rtbCheckSQL.BackColor = &H80000005
        rtbDetail.BackColor = &H80000005
    Else
        rtbCheckSQL.BackColor = &H8000000F
        rtbDetail.BackColor = &H8000000F
    End If
    
    cmdDetailInsert.Enabled = blnState
    cmdSQLInsert.Enabled = blnState
    
    If Len(Trim(rtbCheckSQL.Text)) = 0 Or Not blnState Then
        cmdSQLVerify.Enabled = False
    Else
        cmdSQLVerify.Enabled = True
    End If
    
    If Len(Trim(rtbDetail.Text)) = 0 Or Not blnState Then
        cmdDetailVerify.Enabled = False
    Else
        cmdDetailVerify.Enabled = True
    End If
End Sub

Public Sub SetQueryCfg(objSqlScheme As clsSqlScheme)
    objSqlScheme.Query = rtbCheckSQL.Text
    objSqlScheme.Detail = rtbDetail.Text
End Sub

Private Sub rtbDetail_Change()
    mblnIsEdit = True
    If cmdDetailInsert.Enabled Then
        If Len(Trim(rtbDetail.Text)) = 0 Then
            cmdDetailVerify.Enabled = False
        Else
            cmdDetailVerify.Enabled = True
        End If
    Else
        cmdDetailVerify.Enabled = False
    End If
End Sub

Public Function IsEnabledToSave() As Boolean
    IsEnabledToSave = False
    
    If Len(Trim(rtbCheckSQL.Text)) = 0 Then
        Exit Function
    End If
    
    IsEnabledToSave = True
End Function

Public Sub rtbCheckSQLSetFocus()
    rtbCheckSQL.SetFocus
End Sub
