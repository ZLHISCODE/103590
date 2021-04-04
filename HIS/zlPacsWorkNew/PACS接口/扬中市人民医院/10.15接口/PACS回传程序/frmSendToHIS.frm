VERSION 5.00
Begin VB.Form frmSendToHIS 
   Caption         =   "PACS回传服务"
   ClientHeight    =   2265
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5100
   LinkTopic       =   "Form1"
   ScaleHeight     =   2265
   ScaleWidth      =   5100
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdStart 
      Caption         =   "启动服务"
      Height          =   350
      Left            =   3960
      TabIndex        =   6
      Top             =   1440
      Width           =   1100
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "停止服务"
      Height          =   350
      Left            =   3960
      TabIndex        =   5
      Top             =   1080
      Width           =   1100
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "退出"
      Height          =   350
      Left            =   3960
      TabIndex        =   4
      Top             =   1800
      Width           =   1100
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存设置"
      Height          =   350
      Left            =   3960
      TabIndex        =   3
      Top             =   720
      Width           =   1100
   End
   Begin VB.TextBox txtInterval 
      Height          =   285
      Left            =   4080
      TabIndex        =   2
      Top             =   360
      Width           =   495
   End
   Begin VB.Timer tmlistener 
      Interval        =   5000
      Left            =   3600
      Top             =   720
   End
   Begin VB.Label Label2 
      Caption         =   "秒"
      Height          =   255
      Left            =   4680
      TabIndex        =   7
      Top             =   375
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "时间间隔："
      Height          =   255
      Left            =   4080
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblStatus 
      Height          =   2145
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmSendToHIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim lngInterval As Long
    
    If Val(Me.txtInterval.Text) < 1 Then
        lngInterval = 1
    ElseIf Val(Me.txtInterval.Text) > 65 Then
        lngInterval = 65
    Else
        lngInterval = Val(Me.txtInterval.Text)
    End If
    Me.txtInterval.Text = lngInterval
    
    Me.tmlistener.Enabled = False
    Me.tmlistener.Interval = lngInterval * 1000
    Me.tmlistener.Enabled = True
    glngInterval = lngInterval
    MsgBox "时间间隔设置保存成功！"
End Sub

Private Sub cmdStart_Click()
    Me.tmlistener.Enabled = True
    Me.cmdStop.Enabled = True
    Me.cmdStart.Enabled = False
End Sub

Private Sub cmdStop_Click()
    Me.tmlistener.Enabled = False
    Me.cmdStart.Enabled = True
    Me.cmdStop.Enabled = False
    
End Sub


Private Sub Form_Load()
    If glngInterval < 1 Or glngInterval > 65 Then glngInterval = 1
    Me.tmlistener.Interval = glngInterval * 1000
    Me.txtInterval.Text = Me.tmlistener.Interval / 1000
    Me.cmdStart.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "ZLSOFT", "PACS回传", "HISIP地址", gstrHISIP
    SaveSetting "ZLSOFT", "PACS回传", "HIS用户名", gstrUser
    SaveSetting "ZLSOFT", "PACS回传", "HIS密码", gstrPassw
    SaveSetting "ZLSOFT", "PACS回传", "监听间隔", glngInterval
    
    
    SaveSetting "ZLSOFT", "PACS回传", "PACS用户名", gstrPACSUser
    SaveSetting "ZLSOFT", "PACS回传", "PACS密码", gstrPACSPassw
    SaveSetting "ZLSOFT", "PACS回传", "PACSsid", gstrPACSsid

End Sub

Private Sub tmlistener_Timer()
    subSendDataToHIS
End Sub

Private Sub subSendDataToHIS()
    Dim dsReport As New ADODB.Recordset
    Dim dsAdvice As New ADODB.Recordset
    Dim dsState As New ADODB.Recordset
    Dim strSQL As String
    Dim strCheckDocID As String     '审核医生
    Dim strWriteDocID As String     '诊断医生
    Dim intRowCount As Integer      '成功修改的记录数量
    
    On Error GoTo errLog
    '查询临时表 PACS_TMP病人病历记录，如果有内容，则开始数据回传
    
    strSQL = "select id,报告ID,病人ID,科室ID,病历名称,书写人ID,书写人,to_date(书写日期) as 书写日期,审阅人ID,审阅人,审阅日期,记录类型 from PACS_TMP病人病历记录"
    Set dsReport = gcnOracle.Execute(strSQL)
    If dsReport.EOF Then        '没有数据，退出程序
        Exit Sub
    Else
        dsReport.MoveFirst
        While Not dsReport.EOF
            '根据 PACS_TMP病人病历记录.id=病人医嘱发送.报告id，查询对应的“医嘱ID”
            strSQL = "select 医嘱ID from 病人医嘱发送 where 病人医嘱发送.报告id = " & dsReport!报告ID
            Set dsAdvice = gcnOracle.Execute(strSQL)
            
            If Not dsAdvice.EOF Then
                '对于有对应医嘱ID的报告，根据PACS_TMP病人病历记录.记录类型判断回传“书写人”还是“审阅人”，回传后删除记录
                'PACS的医嘱ID对应HIS的检查号
                lblStatus.Caption = Date & " " & Time & vbCrLf & " 正在回传:ID--" & dsReport!报告ID & " 书写人--" & dsReport!书写人 & " 审阅人--" & dsReport!审阅人
                strWriteDocID = Format(dsReport!书写人ID, "0000")
                strCheckDocID = Format(dsReport!审阅人ID, "0000")
                
                If Len(strWriteDocID) > 4 Then strWriteDocID = Left(strWriteDocID, 4)
                If Len(strCheckDocID) > 4 Then strCheckDocID = Left(strCheckDocID, 4)
                
                If dsReport!记录类型 = 1 Then   '报告书写完成，回传书写人，书写时间和报告完成标志=1
                
                    strSQL = "update pacs_bldak set zdys ='" & strWriteDocID & "', bgdate = '" _
                        & Format(dsReport!书写日期, "YYYY-MM-DD HH:mm:ss") & "',bgzt='1' where jcdh = '" _
                        & dsAdvice!医嘱ID & "'"
                    gcnSQL2K.Execute strSQL, intRowCount
                    
                    'HIS中没有记录被更改，将报告人和报告时间写入日志
                    If intRowCount <= 0 Then
                        subLogErr 100, "报告人写入错误,书写人=" & dsReport!书写人 & " 书写日期=" _
                                       & dsReport!书写日期 & " 检查号 = " & dsAdvice!医嘱ID
                    End If
                    
                ElseIf dsReport!记录类型 = 2 Then   '报告被审核，回传审核人
                    
                    '检查报告是否审核状态，如果已经审核，则回传审核人
                    '否则是报告被修改，直接删除该条记录
                    strSQL = "select 执行状态,执行过程 from 病人医嘱发送 where 医嘱ID =" & dsAdvice!医嘱ID
                    Set dsState = gcnOracle.Execute(strSQL)
                    
                    If dsState!执行状态 = 1 And dsState!执行过程 = 6 Then
                        
                        strSQL = "update pacs_bldak set zdys ='" & strWriteDocID & "', bgdate = '" _
                            & Format(dsReport!书写日期, "YYYY-MM-DD HH:mm:ss") & "',shys ='" _
                            & strCheckDocID & "' where jcdh = '" & dsAdvice!医嘱ID & "'"
                        gcnSQL2K.Execute strSQL, intRowCount
                        
                        If intRowCount <= 0 Then    'HIS中没有记录被更改，将报告人和报告时间写入日志
                            subLogErr 101, "审核人写入错误,书写人=" & dsReport!书写人 & " 书写日期=" _
                                       & dsReport!书写日期 & " 审核人 = " & dsReport!审阅人 _
                                       & " 检查号 = " & dsAdvice!医嘱ID
                        End If
                        
                    End If
                ElseIf dsReport!记录类型 = 3 Then   '报告被驳回，回传报告完成标志=0,目前不需要这个状态
                
                    'strSQL = "update pacs_bldak set bgzt='1' where jcdh = '" & dsAdvice!医嘱ID & "'"
                End If
            
            End If
            '没有对应医嘱的报告为申请单，直接删除记录
            '报告人存储完成后，直接删除记录
            strSQL = "delete from PACS_TMP病人病历记录 where id = " & dsReport!id
            gcnOracle.Execute (strSQL)
            
            lblStatus.Caption = Date & " " & Time & vbCrLf & " 回传完成:ID--" & dsReport!报告ID
            dsReport.MoveNext
        Wend
    End If
    
    Exit Sub
errLog:
    subLogErr Err.Number, Err.Description
End Sub


Private Sub subLogErr(lngErrNo As Long, strDesc As String)
    On Error Resume Next
    Dim lngID As Long
    Dim strSQL As String
    Dim dsRecord As ADODB.Recordset
    
    Me.lblStatus.Caption = Date & " " & Time & vbCrLf & " 发生错误，错误代码：" & lngErrNo & " 错误描述：" & strDesc
    strSQL = "SELECT MAX(ID) as mID FROM PACS_ERR"
    Set dsRecord = gcnOracle.Execute(strSQL)
    If Not dsRecord.EOF Then
        lngID = dsRecord!Mid + 1
    End If
    strSQL = "insert into PACS_ERR (ID,错误号,错误描述,错误时间) values(" & lngID & "," _
            & lngErrNo & ",'" & strDesc & "',sysdate)"
    gcnOracle.Execute strSQL
End Sub


