VERSION 5.00
Object = "{CA73588D-282F-4592-9369-A61CC244FADA}#15.3#0"; "Codejock.SyntaxEdit.v15.3.1.ocx"
Begin VB.Form frmDiffCommon 
   BackColor       =   &H00FFFFFF&
   Caption         =   "文本对比"
   ClientHeight    =   5325
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   8040
   Icon            =   "frmDiffCommon.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   8040
   StartUpPosition =   3  '窗口缺省
   WindowState     =   2  'Maximized
   Begin XtremeSyntaxEdit.SyntaxEdit txtRight 
      Height          =   3495
      Left            =   4680
      TabIndex        =   6
      Top             =   720
      Width           =   3015
      _Version        =   983043
      _ExtentX        =   5318
      _ExtentY        =   6165
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
      ReadOnly        =   -1  'True
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
   Begin XtremeSyntaxEdit.SyntaxEdit txtLeft 
      Height          =   3495
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   3015
      _Version        =   983043
      _ExtentX        =   5318
      _ExtentY        =   6165
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
      ReadOnly        =   -1  'True
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
   Begin VB.CommandButton cmdNext 
      Caption         =   "下一个过程(&N)"
      Height          =   350
      Left            =   1560
      TabIndex        =   14
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "上一个过程(&P)"
      Height          =   350
      Left            =   120
      TabIndex        =   13
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton cmdExchange 
      Caption         =   "从标准过程还原(&E)"
      Height          =   350
      Left            =   4440
      TabIndex        =   12
      Top             =   4680
      Width           =   1950
   End
   Begin VB.PictureBox pctOpt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   3720
      ScaleHeight     =   3255
      ScaleWidth      =   615
      TabIndex        =   9
      Top             =   1320
      Width           =   615
      Begin VB.Image imgUp 
         Height          =   240
         Left            =   75
         Picture         =   "frmDiffCommon.frx":6852
         Top             =   480
         Width           =   240
      End
      Begin VB.Image imgDown 
         Height          =   240
         Left            =   75
         Picture         =   "frmDiffCommon.frx":D0A4
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label lblsta 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "F8"
         ForeColor       =   &H00404000&
         Height          =   180
         Index           =   6
         Left            =   360
         TabIndex        =   11
         Top             =   2310
         Width           =   180
      End
      Begin VB.Label lblsta 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "F7"
         ForeColor       =   &H00404000&
         Height          =   180
         Index           =   7
         Left            =   360
         TabIndex        =   10
         Top             =   510
         Width           =   180
      End
   End
   Begin VB.Timer Timer 
      Interval        =   1
      Left            =   3240
      Top             =   600
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "关闭(&C)"
      Height          =   350
      Left            =   6480
      TabIndex        =   8
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label lblPgs 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "第2/13个过程"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   3000
      TabIndex        =   15
      Top             =   4770
      Width           =   1080
   End
   Begin VB.Label lblRight 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "文本1"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   4680
      TabIndex        =   7
      Top             =   480
      Width           =   450
   End
   Begin VB.Label lblLeft 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "文本1"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   450
   End
   Begin VB.Label lblsta 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "蓝色"
      ForeColor       =   &H00C00000&
      Height          =   180
      Index           =   4
      Left            =   3480
      TabIndex        =   2
      Top             =   120
      Width           =   360
   End
   Begin VB.Label lblsta 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "红色"
      ForeColor       =   &H000000C0&
      Height          =   180
      Index           =   3
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   360
   End
   Begin VB.Label lblsta 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "表示新增的代码,    表示删除的代码,    表示修改代码"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   2
      Left            =   480
      TabIndex        =   3
      Top             =   120
      Width           =   4500
   End
   Begin VB.Label lblsta 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "绿色"
      ForeColor       =   &H0000C000&
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   360
   End
End
Attribute VB_Name = "frmDiffCommon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mcolDiff As New Collection
Private mlngLeftRow As Long
Private mlngRightRow As Long
Private mintLast As Long

Private marrIds() As String '所有传入的ID集合
Private mlngIdx As Long '当前过程ID

Private Enum 颜色
    白色 = &HFFFFFF
    背景色 = &HC9C9CD
    绿色 = &H106E2A
    黑色 = &H0&
    红色 = &H4040FF
    蓝色 = vbBlue
End Enum

Public Sub ShowMe(ByVal arrIds As Variant, ByVal lngIdx As Long)
    '窗体显示
    'arrIds-传入的过程ID数组    lngIdx-当前ID在数组中的下标

    marrIds = arrIds
    mlngIdx = lngIdx
    If LoadProc Then
        Me.Show 1
    Else
        Unload Me
    End If
    
    Unload Me
End Sub

Private Sub cmdExchange_Click()
    Dim strMsg As String
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    
    On Error GoTo errH
    strMsg = "从标准过程还原会丢失用户变动的过程记录，并将数据库中的过程还原为产品标准过程，是否继续？"
    If MsgBox(strMsg, vbYesNo, "确认") = vbNo Then Exit Sub
    
    '从zlProcedureText表中重新提取过程的标准定义
    strSQL = "Select 内容 From zlproceduretext Where 过程ID=[1]  And 性质=1 Order by 序号"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "读取过程文本", marrIds(mlngIdx))
    
    strSQL = ""
    If rsTmp.RecordCount = 0 Then
        strSQL = ""
    Else
        Do While Not rsTmp.EOF
            strSQL = IIf(strSQL = "", rsTmp!内容, strSQL & vbNewLine & rsTmp!内容)
            rsTmp.MoveNext
        Loop
    End If
    gcnOldOra.Execute strSQL, , adCmdText
    
    MsgBox "过程还原成功!", , "提示"
    Exit Sub
errH:
    MsgBox "保存过程中发生错误" & vbNewLine & err.Description, , "错误"
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    txtLeft.Width = Me.ScaleWidth / 2 - 500
    txtRight.Width = Me.ScaleWidth / 2 - 500
    
    pctOpt.Top = Me.ScaleHeight / 2 - pctOpt.Height / 2
    pctOpt.Left = txtLeft.Left + txtLeft.Width
    
    txtRight.Left = pctOpt.Left + pctOpt.Width
    lblRight.Left = txtRight.Left
    
    cmdExit.Top = Me.ScaleHeight - cmdExit.Height - 120
    cmdExit.Left = txtRight.Left + txtRight.Width - cmdExit.Width
    cmdExchange.Top = cmdExit.Top
    cmdExchange.Left = cmdExit.Left - cmdExchange.Width - 60
    cmdPrevious.Top = cmdExit.Top
    cmdNext.Top = cmdExit.Top
    lblPgs.Top = cmdNext.Top + cmdNext.Height / 2 - lblPgs.Height / 2
    
    txtLeft.Height = cmdExit.Top - txtLeft.Top - 60
    txtRight.Height = txtLeft.Height
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 118 Then '按下F7
        Call imgUp_Click
    ElseIf KeyCode = 119 Then '按下F8
        Call imgDown_Click
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mcolDiff = Nothing
End Sub

Private Sub txtLeft_GotFocus()
    mintLast = 1
End Sub

Private Sub txtRight_GotFocus()
    mintLast = 2
End Sub

Private Sub imgDown_Click()
    Dim i As Long

    If mintLast = 2 Then
        For i = txtRight.CurrPos.Row + 1 To txtRight.RowsCount - 1
            If GetValueFromCol(mcolDiff, "_" & i) <> "" And GetValueFromCol(mcolDiff, "_" & i - 1) = "" Then
                txtRight.CurrPos.Row = i
                txtRight.TopRow = i - 50
                Exit Sub
            End If
        Next
    Else
        '左侧控件 ,需要遍历集合
        For i = txtLeft.CurrPos.Row + 1 To txtLeft.RowsCount - 1
            If GetValueFromCol(mcolDiff, "_" & i) <> "" And GetValueFromCol(mcolDiff, "_" & i - 1) = "" Then
                txtLeft.CurrPos.Row = i
                txtLeft.TopRow = i - 50
                Exit Sub
            End If
        Next
    End If
End Sub

Private Sub imgUp_Click()
    Dim i As Long
    
    If mintLast = 2 Then
        For i = txtRight.CurrPos.Row - 1 To 1 Step -1
            If GetValueFromCol(mcolDiff, "_" & i) <> "" And GetValueFromCol(mcolDiff, "_" & i - 1) = "" Then
                txtRight.CurrPos.Row = i
                txtRight.TopRow = i - 50
                Exit Sub
            End If
        Next
    Else
        '左侧控件 ,需要遍历集合
        For i = txtLeft.CurrPos.Row - 1 To 1 Step -1
            If GetValueFromCol(mcolDiff, "_" & i) <> "" And GetValueFromCol(mcolDiff, "_" & i - 1) = "" Then
                txtLeft.CurrPos.Row = i
                txtLeft.TopRow = i - 50
                Exit Sub
            End If
        Next
    End If
End Sub

Private Sub lblsta_Click(Index As Integer)
    Select Case Index
        Case 6
            Call imgDown_Click
        Case 7
            Call imgUp_Click
    End Select
End Sub

Private Sub txtLeft_CurPosChanged(ByVal nNewRow As Long, ByVal nNewCol As Long)
    With txtLeft
        .SetRowBkColor mlngLeftRow, 白色
        .SetRowBkColor nNewRow, 背景色
        mlngLeftRow = nNewRow
    End With
End Sub

Private Sub txtLeft_LostFocus()
    txtLeft.SetRowBkColor mlngLeftRow, 白色
End Sub

Private Sub txtRight_LostFocus()
    txtRight.SetRowBkColor mlngRightRow, 白色
End Sub

Private Sub txtRight_CurPosChanged(ByVal nNewRow As Long, ByVal nNewCol As Long)
    With txtRight
        .SetRowBkColor mlngRightRow, 白色
        .SetRowBkColor nNewRow, 背景色
        mlngRightRow = nNewRow
    End With
End Sub

Private Sub Timer_Timer()
    If Me.ActiveControl.Name = "txtRight" Then
        txtLeft.TopRow = txtRight.TopRow
    ElseIf Me.ActiveControl.Name = "txtLeft" Then
        txtRight.TopRow = txtLeft.TopRow
    End If
End Sub

Private Function LoadProc() As Boolean
    Dim strTxt1 As String, strTxt2 As String
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim lngID As Long, strName As String
    
    On Error GoTo errH
    ShowFlash "正在对比..."
    lngID = Split(marrIds(mlngIdx), ":")(0)
    strName = Split(marrIds(mlngIdx), ":")(1)
    
    '设置对比控件上方的label
    strSQL = "Select 升级前版本 From zlProcedure where ID = [1]"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "获取版本", lngID)
    If IsNull(rsTmp!升级前版本) Then
        lblLeft.Caption = "产品标准过程"
    Else
        lblLeft.Caption = "产品标准过程(" & rsTmp!升级前版本 & ")"
    End If
    lblRight.Caption = "用户变动过程"
    lblPgs.Caption = "第" & mlngIdx + 1 & "/" & UBound(marrIds) + 1 & "个过程"
    
    '获取过程定义并对比
    strSQL = "Select 内容 From zlproceduretext Where 过程ID=[1]  And 性质=1 Order by 序号"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "读取过程文本", lngID)
    
    If rsTmp.RecordCount = 0 Then
        strTxt1 = ""
    Else
        Do While Not rsTmp.EOF
            strTxt1 = IIf(strTxt1 = "", rsTmp!内容, strTxt1 & vbNewLine & rsTmp!内容)
            rsTmp.MoveNext
        Loop
    End If
    strTxt2 = LoadBaseProcs(strName)
    
    CompareIt strTxt1, strTxt2: MergeDiff strTxt1, strTxt2
    MergeDiffInto2SynEdit strTxt1, strTxt2, txtLeft, txtRight, mcolDiff
    
    ShowFlash ""
    LoadProc = True
    Exit Function
errH:
    ShowFlash ""
    MsgBox "还原过程中发生错误:" & vbNewLine & err.Description
End Function

Private Sub cmdNext_Click()
    If mlngIdx = UBound(marrIds) Then
        MsgBox "当前已经是最后一个过程。", , gstrSysName
        Exit Sub
    End If
    
    mlngIdx = mlngIdx + 1
    Call LoadProc
End Sub

Private Sub cmdPrevious_Click()
    If mlngIdx = 0 Then
        MsgBox "当前已经是第一个过程。", , gstrSysName
        Exit Sub
    End If
    
    mlngIdx = mlngIdx - 1
    Call LoadProc
End Sub
