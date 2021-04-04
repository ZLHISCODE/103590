VERSION 5.00
Object = "{CA73588D-282F-4592-9369-A61CC244FADA}#15.3#0"; "Codejock.SyntaxEdit.v15.3.1.ocx"
Begin VB.Form frmProcEditCommon 
   BackColor       =   &H00FFFFFF&
   Caption         =   "过程编辑"
   ClientHeight    =   7260
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   9090
   Icon            =   "frmProcEditCommon.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7260
   ScaleWidth      =   9090
   StartUpPosition =   1  '所有者中心
   WindowState     =   2  'Maximized
   Begin XtremeSyntaxEdit.SyntaxEdit txtEdit 
      Height          =   5655
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   8655
      _Version        =   983043
      _ExtentX        =   15266
      _ExtentY        =   9975
      _StockProps     =   84
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   1
      EnableSyntaxColorization=   -1  'True
      ShowLineNumbers =   -1  'True
      ShowSelectionMargin=   0   'False
      ShowScrollBarVert=   -1  'True
      ShowScrollBarHorz=   -1  'True
      EnableVirtualSpace=   0   'False
      EnableAutoIndent=   -1  'True
      ShowWhiteSpace  =   0   'False
      ShowCollapsibleNodes=   -1  'True
      AutoCompleteWndWidth=   160
      EnableEditAccelerators=   -1  'True
   End
   Begin VB.PictureBox pctBottom 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   240
      ScaleHeight     =   465
      ScaleWidth      =   8610
      TabIndex        =   3
      Top             =   6600
      Width           =   8610
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   1200
         TabIndex        =   2
         Top             =   5
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Appearance      =   0  'Flat
         Caption         =   "保存(&O)"
         Height          =   350
         Left            =   0
         TabIndex        =   1
         Top             =   5
         Width           =   1095
      End
   End
   Begin VB.Label lblProc 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "过程名称"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "frmProcEditCommon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngId As Long
Private mstrProcDec As String
Private mstrUser As String

Private mblnSave As Boolean
Private mblnErr As Boolean

Public Function ShowMe(ByVal lngID As Long, ByVal strProcName As String, ByVal strProcTxt As String, _
                                        ByVal strProcSys As String, ByVal strProcDec As String, ByVal strUser As String, _
                                        Optional ByVal bytType As Byte, Optional lngLine As Long) As Boolean
    'bytType:1-过程检查结果窗体调用，默认过程变动窗体调用
    '设置控件格式
    With txtEdit
        '设置控件的显示颜色方案为：SQL
        .SyntaxSet "[Schemes]" & vbCrLf & "SQL" & vbCrLf & "[Themes]" & vbCrLf & "Default" & vbCrLf & "Alternative" & vbCrLf
        .SyntaxScheme = GetSqlColor
        .Text = strProcTxt
    End With
    
    If bytType = 0 Then
        mlngId = lngID
        mstrProcDec = strProcDec
        mstrUser = strUser
        Me.Caption = "过程编辑"
        cmdCancel.Caption = "取消(&C)"
        lblProc.Caption = strProcName & "(" & strProcSys & ")"
    Else
        Me.Caption = "过程查看"
        cmdCancel.Caption = "退出(&E)"
        lblProc.Caption = strProcName
        cmdSave.Visible = False
        txtEdit.CurrPos.Row = lngLine
        txtEdit.ReadOnly = True
    End If
    Me.Show 1
    If bytType = 0 Then
        ShowMe = mblnSave    '如果是修改保存关闭窗口就返回True,否则返回Fasle
    End If
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim strSQL As String
    
    On Error Resume Next    '为了避免编译出错,这里要屏蔽错误
    strSQL = txtEdit.Text
    '填写修改说明
    If Not frmProcEditor.ShowMe(mstrUser, mstrProcDec) Then Exit Sub
    
    '保存过程文本,执行一次即可
    gcnOldOra.Execute strSQL '这里用微软的驱动执行,防止因为单引号等特殊字符保存失败
    If err.Number <> 0 Then
        mblnErr = True
        MsgBox "过程编译出错，请检查后重试" & vbNewLine & err.Description, , "错误"
        Exit Sub
    ElseIf gcnOracle.Errors.Count > 1 Then
        mblnErr = True
        MsgBox "过程编译出错，请检查后重试", , "错误"
        Exit Sub
    Else
        mblnErr = False
    End If

    '更新zlProcedure数据
    strSQL = "Update zlProcedure Set 说明 = '" & mstrProcDec & "',修改人员='" & mstrUser & "',修改时间= Sysdate " & vbNewLine & _
                ",上次修改人员 = 修改人员, 上次修改时间 = 修改时间  Where ID = " & mlngId
    gcnOracle.Execute strSQL, "保存变动过程。"
    If err <> 0 Then
        mblnErr = True
        MsgBox "过程保存失败，请检查后重试" & vbNewLine & err.Description, , "提示"
        Exit Sub
    End If
    
    mblnSave = True
    Unload Me
    
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    
    pctBottom.Top = Me.ScaleHeight - pctBottom.Height - 100
    pctBottom.Width = Me.ScaleWidth
    
    cmdCancel.Left = pctBottom.ScaleWidth - cmdCancel.Width - 360
    cmdSave.Left = cmdCancel.Left - cmdSave.Width - 120
    
    txtEdit.Width = pctBottom.Width - txtEdit.Left - 240
    txtEdit.Height = pctBottom.Top - txtEdit.Top - 120
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strMsg As String
    
    strMsg = "编辑后的过程存在错误，不加处理会导致该过程失效，是否继续退出？"
    
    If mblnErr Then
        If MsgBox(strMsg, vbOKCancel, "退出确认") = vbNo Then
            Cancel = 1
        End If
    End If
End Sub

