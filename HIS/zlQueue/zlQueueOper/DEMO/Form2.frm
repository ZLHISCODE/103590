VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{1B83D023-3CA6-4181-A286-20352E645AE2}#2.2#0"; "zlQueueOper.ocx"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   6930
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13155
   LinkTopic       =   "Form2"
   ScaleHeight     =   6930
   ScaleWidth      =   13155
   StartUpPosition =   3  '窗口缺省
   Begin zlQueueOper.UcQueue UcQueueStation1 
      Height          =   4770
      Left            =   90
      TabIndex        =   8
      Top             =   210
      Width           =   12690
      _ExtentX        =   22384
      _ExtentY        =   8414
      Interval        =   30000
      ValidDays       =   0
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1650
      Left            =   105
      ScaleHeight     =   1590
      ScaleWidth      =   12675
      TabIndex        =   0
      Top             =   5160
      Width           =   12735
      Begin VB.CommandButton Command2 
         Caption         =   "配置测试数据"
         Height          =   390
         Left            =   8190
         TabIndex        =   9
         Top             =   405
         Width           =   1500
      End
      Begin VB.CommandButton Command3 
         Caption         =   "刷新数据"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6765
         TabIndex        =   7
         Top             =   420
         Width           =   1200
      End
      Begin VB.CommandButton Command1 
         Caption         =   "站点设置"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5430
         TabIndex        =   6
         Top             =   390
         Width           =   1005
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3945
         TabIndex        =   5
         Text            =   "Text2"
         Top             =   405
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1650
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   405
         Width           =   1245
      End
      Begin VB.Label Label3 
         Caption         =   "呼叫站点"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2955
         TabIndex        =   4
         Top             =   480
         Width           =   1080
      End
      Begin VB.Label Label2 
         Caption         =   "本地站点"
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
         Left            =   750
         TabIndex        =   3
         Top             =   480
         Width           =   900
      End
      Begin XtremeCommandBars.CommandBars cbrMain 
         Left            =   90
         Top             =   15
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   465
         Left            =   60
         TabIndex        =   1
         Top             =   1110
         Width           =   12570
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cbrMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call UcQueueStation1.zlExecuteCommandBars(Control)
End Sub

Private Sub Command1_Click()
    UcQueueStation1.QueueOper.LocalStation = Text1.Text
    UcQueueStation1.QueueOper.PlayStation = Text2.Text
End Sub

Private Sub Command2_Click()
    Call ConfigDefaultQueueData
End Sub

Private Sub Command3_Click()
    Call UcQueueStation1.RefreshQueueData
End Sub

Private Sub Form_Load()
    Dim cbrMenuBar As CommandBarPopup
    
    Call OraDataOpen("002133-1033ORCLMSG", "ZLHIS", TranPasswd("aqa"))
    
    
    Call InitCommon(gcnOracle)
    Call SetDbUser("ZLHIS")
    
    Call ConnectMip(Me.hWnd)
    
    '启用消息中心
    Call UcQueueStation1.UseMsgCenter(100, 1290)
    
    UcQueueStation1.QueryQueueNames = "测试队列,QUEUE1"           '如果未设置此属性，则显示该业务类型下的所有队列数据
    UcQueueStation1.ReportNum = "Test"
    
    UcQueueStation1.DataFields = "ID,队列名称,排队序号,排队号码,排队状态,患者姓名,诊室,医生姓名,排队时间,备注,测试1,呼叫医生,呼叫时间,优先,测试1"
    UcQueueStation1.DisplayQueueFields = "排队号码,患者姓名,诊室,医生姓名,排队时间,备注,测试1,呼叫时间,排队状态"
    UcQueueStation1.DisplayCallFields = "排队号码,患者姓名,呼叫时间"    '如果没有设置此属性，则显示所有字段
        
'    UcQueueStation1.CustomOrderField = "排队序号" ' "患者姓名"  '设置队列的排序字段，如果未设置，则使用数据库的默认排序方式
    
    UcQueueStation1.GroupField = "队列名称"                      '设置排队叫号的分组方式
    
'    UcQueueStation1.IsShowBars = True                           '设置是否显示排队叫号的操作工具栏，如果不设置，则默认显示工具栏
    
'    UcQueueStation1.IsShowCalledQueue = True                    '设置是否显示已呼叫队列,如果不设置，则默认显示工具栏
    
    UcQueueStation1.FindWayEx = "门诊号,住院号,就诊号,社保卡,其他"
    
    Call UcQueueStation1.InitQueue(gcnOracle, 1, Me, App.ProductName, "ZLHIS", ",打号,顺呼,直呼,广播,优先,插队,重排,接诊,暂停,弃号,恢复,完成,刷新,过滤,定位,查找,修改,设置,")
    
    UcQueueStation1.QueueOper.VoiceType = 1
    
    
    Call UcQueueStation1.ApplyVoiceConfig
    
    Call UcQueueStation1.StartVoice
    
    Set UcQueueStation1.Font = Me.Font
    
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "排队")
    
    Call UcQueueStation1.zlCreateMenuBars(cbrMenuBar, True)
End Sub

Private Sub ConfigDefaultQueueData()
    Dim objQueue As clsQueueOperation
    Dim i As Long
    Dim lngQueueId As Long
    Dim strNewQueueNo As String
    
    Set objQueue = UcQueueStation1.QueueOper

    For i = 1 To 10
        lngQueueId = objQueue.InsertQueue("测试队列", , , "刘" & i & Format(Now, "hh:mm:ss") & Rnd)
        Call objQueue.LineQueue(lngQueueId, strNewQueueNo)
    Next i

    Call UcQueueStation1.RefreshQueueData
    
End Sub

Private Sub Form_Resize()
On Error Resume Next
    UcQueueStation1.Left = 0
    UcQueueStation1.Top = 0
    UcQueueStation1.Width = Me.ScaleWidth
    UcQueueStation1.Height = Me.ScaleHeight - Picture1.Height
    
    Picture1.Top = UcQueueStation1.Height
    Picture1.Left = 0
    Picture1.Width = Me.ScaleWidth
Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call UcQueueStation1.StopVoice
    
    Call DisConnectMip
End Sub

Private Sub UcQueueStation1_OnCallPreBefore(ByVal lngQueueId As Long, ByVal lngCallWay As TCallWay, strCallContext As String, blnCancel As Boolean)
'    MsgBox "呼叫操作：" & lngQueueId
End Sub

Private Sub UcQueueStation1_OnCmdBarExecute(objControl As Object, ByRef blnUseCustom As Boolean)
    Dim strName As String
    Dim lngRowIndex As Long
    Dim ID As Long
    
    Select Case objControl.ID
        Case 7890
            blnUseCustom = True
            
            lngRowIndex = UcQueueStation1.GetCalledQueueIndex()
            
            If lngRowIndex >= 0 Then
                strName = UcQueueStation1.GetListValue(qftCalledQueue, lngRowIndex, "患者姓名")
            End If
            
            '获取队列ID后，呼叫需候诊患者
            ID = UcQueueStation1.GetListValue(qftCalledQueue, lngRowIndex, "ID")
            Call UcQueueStation1.QueueOper.WaitRoomCall(ID)
            
            MsgBox "测试按钮被按下  姓名：" & strName
    End Select
End Sub

Private Sub UcQueueStation1_OnCmdBarInit(CmdBar As Object)
    CmdBar.Controls.Add(1, 7890, "测试", 7).IconId = 721
End Sub


Private Sub UcQueueStation1_OnCustomFindButton(ByVal lngQueueId As Long)
    MsgBox "自定义查找执行：" & lngQueueId
End Sub

Private Sub UcQueueStation1_OnCmdBarUpdate(objControl As Object)
    If objControl.ID = 7890 Then
        objControl.Enabled = UcQueueStation1.CurQueueType = qftCalledQueue
    End If
End Sub

Private Sub UcQueueStation1_OnColumnInit(objQueueList As Object, objReportColumn As Object)
    If objReportColumn.Caption = "排队时间" Then
'        objReportColumn.Width = 200
        objReportColumn.Icon = 721
    End If
End Sub

Private Sub UcQueueStation1_OnFindData(ByVal strFindWay As String, ByVal strFindValue As String, txtFind As Object, rsData As ADODB.Recordset, blnUseCustom As Boolean)
    Dim strSql As String
    
    Select Case strFindWay
        Case "门诊号"
            strSql = "select * from 排队叫号队列"
        Case "住院号"
            strSql = "select * from 排队叫号队列"
        Case Else
            Exit Sub
    End Select
    
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "检索排队数据", strFindValue)
    
    blnUseCustom = True
End Sub

Private Sub UcQueueStation1_OnItemDblClick(ByVal lngListType As zlQueueOper.TQueueFromType, ByVal lngQueueId As Long, objReoprtRow As Object, objReportItem As Object)
    MsgBox UcQueueStation1.GetListValue(lngListType, objReoprtRow.Index, "患者姓名")
End Sub



Private Sub UcQueueStation1_OnPlayVoiceAfter(ByVal lngCallId As Long, ByVal lngQueueId As Long, ByVal strCallContext As String)
    Label1.Caption = "完成呼叫，呼叫内容为[" & strCallContext & "]"
End Sub



Private Sub UcQueueStation1_OnQueryQueueData(rsData As ADODB.Recordset, blnUseCustom As Boolean)
    Dim strSql As String
    '排队号码,患者姓名,优先,诊室,医生姓名,排队状态,排队时间,备注
    
    strSql = "select " & UcQueueStation1.GetValidCols("ID,排队号码,患者姓名,优先,诊室,医生姓名,排队状态,排队时间,备注, 排队序号, '测试1' as 测试1,'测试2' as 测试2, 呼叫时间", "") & _
            " from 排队叫号队列 where 业务类型=" & UcQueueStation1.WorkType & _
          " and 排队时间 between sysdate - 1 and sysdate and 队列名称 in ('" & Replace(UcQueueStation1.QueryQueueNames, ",", "','") & "') order by 排队序号 "
            
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "获取排队数据")
    
    blnUseCustom = True
    
        
End Sub




Private Sub UcQueueStation1_OnReadAfter(rsData As ADODB.Recordset, ByVal lngListType As TQueueFromType, objReportRecord As Object)
    If InStr(rsData!排队序号, ".") > 0 Then
        objReportRecord(0).Icon = 3560
        
    End If

    Label1.Caption = "OnReadAfter事件执行：队列类型为" & lngListType
End Sub

Private Sub UcQueueStation1_OnSelectionChanged(ByVal lngListType As zlQueueOper.TQueueFromType, ByVal lngQueueId As Long, objQueueList As Object, objReportRow As Object)
    If objReportRow.GroupRow = True Then Exit Sub
    
    Label1.Caption = "OnSelectionChanged事件执行：" & objReportRow.Record(1).Value
End Sub

Private Sub UcQueueStation1_OnWorkBefore(ByVal lngQueueId As Long, ByVal lngOperationType As TOperationType, blnCancel As Boolean)
    Select Case lngOperationType
        Case TOperationType.otDiagnose
            Label1.Caption = "OnWorkBefore事件执行：操作类型为接诊"
            
            Call UcQueueStation1.QueueOper.AbstainQueue(lngQueueId)
        Case TOperationType.otAbstain
            Label1.Caption = "OnWorkBefore事件执行：操作类型为弃号"
        Case TOperationType.otPause
            Label1.Caption = "OnWorkBefore事件执行：操作类型为暂停"
        Case TOperationType.otComplete
            Label1.Caption = "OnWorkBefore事件执行：操作类型为完成"
        Case TOperationType.otRestore
            Label1.Caption = "OnWorkBefore事件执行：操作类型为重排"
        Case TOperationType.otStart
            Label1.Caption = "OnWorkBefore事件执行：操作类型为恢复"
        Case TOperationType.otPrintNo
            MsgBox "执行打号处理：ID" & lngQueueId
        Case Else
            Label1.Caption = "OnWorkBefore事件执行：操作类型为" & lngOperationType
    End Select
End Sub
