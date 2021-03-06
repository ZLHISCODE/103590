VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmShowLog 
   Caption         =   "通讯日志"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9390
   Icon            =   "frmShowLog.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5760
   ScaleWidth      =   9390
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdClose 
      Caption         =   "关闭"
      Height          =   300
      Left            =   6720
      TabIndex        =   3
      Top             =   5280
      Width           =   1100
   End
   Begin VB.CommandButton cmdChangeDate 
      Caption         =   "后一天"
      Height          =   300
      Index           =   1
      Left            =   4200
      TabIndex        =   2
      Top             =   5280
      Width           =   1100
   End
   Begin VB.CommandButton cmdChangeDate 
      Caption         =   "前一天"
      Height          =   300
      Index           =   0
      Left            =   1320
      TabIndex        =   1
      Top             =   5280
      Width           =   1100
   End
   Begin MSComctlLib.ListView listLog 
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   8916
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmShowLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intLogType As Integer    '日志类型
Private intDay As Integer       '间隔天数

Private Sub cmdChangeDate_Click(Index As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    
    If Index = 0 Then '前一天
        If intLogType = 1 Then  '通讯和处理日志
            strSQL = "SELECT datediff('d',通讯时间 , date()) as differ FROM HL7通讯日志 where datediff('d',通讯时间 , date()) >" & _
                 intDay & " order by 通讯时间 desc"
        ElseIf intLogType = 2 Then  '显示消息记录
            strSQL = "SELECT datediff('d',通讯时间 , date()) as differ FROM HL7消息记录 where datediff('d',通讯时间 , date()) >" & _
                 intDay & " order by 通讯时间 desc"
        ElseIf intLogType = 3 Then  '错误日志
            strSQL = "SELECT datediff('d',产生时间 , date()) as differ FROM 错误日志 where datediff('d',产生时间 , date()) >" & _
                 intDay & " order by 产生时间 desc"
        End If
        Set rsTmp = gcnAccess.Execute(strSQL)
        If Not rsTmp.EOF Then
            intDay = rsTmp!differ
        Else
            cmdChangeDate(Index).Enabled = False
        End If
        cmdChangeDate(1).Enabled = True
        subShowList
    Else '后一天
        If intLogType = 1 Then      '通讯和处理日志
            strSQL = "SELECT datediff('d',通讯时间 ,date()) as differ FROM HL7通讯日志 where datediff('d',通讯时间 ,date()) <" & _
                 intDay & " order by 通讯时间 desc"
        ElseIf intLogType = 2 Then  '显示消息记录
            strSQL = "SELECT datediff('d',通讯时间 ,date()) as differ FROM HL7消息记录 where datediff('d',通讯时间 ,date()) <" & _
                 intDay & " order by 通讯时间 desc"
        ElseIf intLogType = 3 Then  '错误日志
            strSQL = "SELECT datediff('d',产生时间 ,date()) as differ FROM 错误日志 where datediff('d',产生时间 ,date()) <" & _
                 intDay & " order by 产生时间 desc"
        End If
        Set rsTmp = gcnAccess.Execute(strSQL)
        If Not rsTmp.EOF Then
            intDay = rsTmp!differ
        Else
            cmdChangeDate(Index).Enabled = False
        End If
        cmdChangeDate(0).Enabled = True
        subShowList
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    If intLogType = 1 Then
        Me.Caption = "通讯和处理日志"
    ElseIf intLogType = 2 Then
        Me.Caption = "消息记录"
    ElseIf intLogType = 3 Then
        Me.Caption = "错误日志"
    ElseIf intLogType = 4 Then
        Me.Caption = "服务日志"
    End If
    
    With listLog
        With .ColumnHeaders
            .Clear
            If intLogType = 1 Then              '通讯和处理日志
                .Add , , "ID", 400
                .Add , , "通讯时间", 1800
                .Add , , "通讯函数", 1700
                .Add , , "记录标题", 1700
                .Add , , "记录内容", 3400
            ElseIf intLogType = 2 Then          '显示消息记录
                .Add , , "ID", 400
                .Add , , "通讯时间", 1800
                .Add , , "消息类型", 1700
                .Add , , "消息内容", 3400
            ElseIf intLogType = 3 Then          '显示错误日志
                .Add , , "ID", 400
                .Add , , "产生时间", 1800
                .Add , , "错误号", 1700
                .Add , , "错误类型", 1700
                .Add , , "错误信息", 3400
            ElseIf intLogType = 4 Then          '显示服务日志
                .Add , , "序号", 400
                .Add , , "发送程序名称", 1000
                .Add , , "发送设备名称", 1000
                .Add , , "接收程序名称", 1000
                .Add , , "接收设备名称", 1000
                .Add , , "IP地址", 1200
                .Add , , "服务端口", 900
                .Add , , "服务类型", 900
                .Add , , "服务状态", 800
            End If
        End With
        .ListItems.Add , , "Temp"
        .ListItems.Clear
    End With
    
    If intLogType = 1 Or intLogType = 2 Or intLogType = 3 Then
        intDay = 0
        cmdChangeDate(0).Visible = True
        cmdChangeDate(1).Visible = True
        cmdChangeDate_Click (1)
    ElseIf intLogType = 4 Then      '服务日志，不显示“前一天，后一天”的按钮
        cmdChangeDate(0).Visible = False
        cmdChangeDate(1).Visible = False
    End If
    Call subShowList
    gzlComLib.RestoreWinState Me, App.ProductName
End Sub


Private Sub subShowList()
    Dim rsTmp As New ADODB.Recordset
    Dim strCurKey As String
    Dim tmpItem As MSComctlLib.ListItem
    Dim strSQL As String
    Dim strContent As String
    Dim i As Integer
    
    If gcnAccess.State = adStateOpen Then
        If intLogType = 1 Then      '通讯和处理日志
            strSQL = "SELECT ID,通讯时间,通讯函数,记录标题,记录内容 FROM HL7通讯日志 where datediff('d',通讯时间 ,date()) <=" & _
                     intDay & " and datediff('d',通讯时间 ,date()) > " & intDay - 1
            Set rsTmp = gcnAccess.Execute(strSQL)
            
            Me.listLog.ListItems.Clear
            Do While Not rsTmp.EOF
                Set tmpItem = listLog.ListItems.Add(, "_" & rsTmp("ID"), Nvl(rsTmp("ID")))
                With tmpItem
                    .SubItems(1) = Nvl(rsTmp("通讯时间"))
                    .SubItems(2) = Nvl(rsTmp("通讯函数"))
                    .SubItems(3) = Nvl(rsTmp("记录标题"))
                    strContent = rsTmp.Fields("记录内容")
                    .SubItems(4) = Nvl(strContent)
                End With
                rsTmp.MoveNext
            Loop
        ElseIf intLogType = 2 Then      '显示消息记录
            strSQL = "SELECT ID,通讯时间,消息类型,消息内容 FROM HL7消息记录 where datediff('d',通讯时间 ,date()) <=" & _
                     intDay & " and datediff('d',通讯时间 ,date()) > " & intDay - 1
            Set rsTmp = gcnAccess.Execute(strSQL)
            
            Me.listLog.ListItems.Clear
            Do While Not rsTmp.EOF
                Set tmpItem = listLog.ListItems.Add(, "_" & rsTmp("ID"), Nvl(rsTmp("ID")))
                With tmpItem
                    .SubItems(1) = Nvl(rsTmp("通讯时间"))
                    .SubItems(2) = Nvl(rsTmp("消息类型"))
                    strContent = rsTmp.Fields("消息内容")
                    .SubItems(3) = Nvl(strContent)
                End With
                rsTmp.MoveNext
            Loop
        ElseIf intLogType = 3 Then      '错误日志
            strSQL = "select ID,产生时间,错误号,错误类型,错误信息 from 错误日志 where datediff('d',产生时间 , date()) <=" & _
                     intDay & " and datediff('d',产生时间 , date()) > " & intDay - 1
            Set rsTmp = gcnAccess.Execute(strSQL)
            
            Me.listLog.ListItems.Clear
            Do While Not rsTmp.EOF
                Set tmpItem = listLog.ListItems.Add(, "_" & rsTmp("ID"), Nvl(rsTmp("ID")))
                With tmpItem
                    .SubItems(1) = Nvl(rsTmp("产生时间"))
                    .SubItems(2) = Nvl(rsTmp("错误号"))
                    .SubItems(3) = Nvl(rsTmp("错误类型"))
                    strContent = rsTmp.Fields("错误信息")
                    .SubItems(4) = Nvl(strContent)
                End With
                rsTmp.MoveNext
            Loop
        ElseIf intLogType = 4 Then      '服务日志
            Me.listLog.ListItems.Clear
            For i = 1 To UBound(HL7Services)
                Set tmpItem = listLog.ListItems.Add(, "_" & i, i)
                With tmpItem
                    .SubItems(1) = HL7Services(i).strSendApp
                    .SubItems(2) = HL7Services(i).strSendFacility
                    .SubItems(3) = HL7Services(i).strReceiveApp
                    .SubItems(4) = HL7Services(i).strReceiveFacility
                    .SubItems(5) = HL7Services(i).strIP
                    .SubItems(6) = HL7Services(i).lngPort
                    .SubItems(7) = IIf(HL7Services(i).intServiceType = 1, "接收", "发送")
                    If .SubItems(7) = "发送" Then
                        .SubItems(8) = IIf(gblnServiceStart = True, "已启动", "已停止")
                    Else
                        .SubItems(8) = IIf(HL7Services(i).Started = True, "已启动", "已停止")
                    End If
                End With
            Next i
        End If
    End If
    Exit Sub
End Sub


Private Sub Form_Resize()
    Me.listLog.Left = 0
    Me.listLog.Top = 0
    Me.listLog.Width = Abs(Me.Width - 100)
    Me.listLog.Height = Abs(Me.Height - 1200)
    Me.cmdChangeDate(0).Top = Me.listLog.Height + 200
    Me.cmdChangeDate(1).Top = Me.cmdChangeDate(0).Top
    Me.cmdClose.Top = Me.cmdChangeDate(0).Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
    gzlComLib.SaveWinState Me, App.ProductName
End Sub

Private Sub listLog_DblClick()
    If intLogType = 1 Or intLogType = 3 Then
        MsgBox Me.listLog.SelectedItem.SubItems(4)
    ElseIf intLogType = 2 Then
        MsgBox Me.listLog.SelectedItem.SubItems(3)
    End If
End Sub
