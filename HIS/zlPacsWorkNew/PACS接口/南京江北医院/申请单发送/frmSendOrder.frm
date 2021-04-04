VERSION 5.00
Begin VB.Form frmSendOrder 
   Caption         =   "HIS－电子申请单－发送服务"
   ClientHeight    =   2265
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5100
   Icon            =   "frmSendOrder.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2265
   ScaleWidth      =   5100
   StartUpPosition =   2  '屏幕中心
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
Attribute VB_Name = "frmSendOrder"
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
    
    SaveSetting "ZLSOFT", gstrRegPath, "HIS用户名", gstrHISUser
    SaveSetting "ZLSOFT", gstrRegPath, "HIS密码", gstrHISPassw
    SaveSetting "ZLSOFT", gstrRegPath, "HISsid", gstrHISsid
     
    SaveSetting "ZLSOFT", gstrRegPath, "监听间隔", glngInterval
    
    SaveSetting "ZLSOFT", gstrRegPath, "PACSIP地址", gstrPACSIP
    SaveSetting "ZLSOFT", gstrRegPath, "PACS用户名", gstrPACSUser
    SaveSetting "ZLSOFT", gstrRegPath, "PACS密码", gstrPACSPassw
    SaveSetting "ZLSOFT", gstrRegPath, "PACSsid", gstrPACSsid
    SaveSetting "ZLSOFT", gstrRegPath, "gstrPACSport", gstrPACSport

End Sub

Private Sub tmlistener_Timer()
    subSendDataToPACS
End Sub

Private Sub subSendDataToPACS()
    Dim adoCmd As New ADODB.Command
    Dim dsOrder As ADODB.Recordset
    Dim adoParaReturn As Parameter  '存储返回值
    Dim strSQL As String
    
    
    
    On Error GoTo errLog
    '查询临时表 PACS_TMP病人病历记录，如果有内容，则开始数据回传
    
    strSQL = "Select ID,操作类型,病人类别,医嘱ID,标识号,病人ID,姓名,英文名,性别,出生日期,身份证号," & _
             "家庭电话,家庭地址,病区,床号,影像类别,检查项目代码,检查项目描述,开嘱医生,开嘱科室,病史," & _
             "临床诊断,注意事项,备注 From ZLPACS接口KODAK where 操作类型=1 or 操作类型=2"
    Set dsOrder = gcnHIS.Execute(strSQL)
    If dsOrder.EOF Then        '没有数据，退出程序
        Exit Sub
    Else
        dsOrder.MoveFirst
        While Not dsOrder.EOF
            '把每一条记录的结果组织成电子申请单，通过SP_EOrder_For_Kodak发送给柯达RIS
            
            '调用带返回值的存储过程
    
            adoCmd.CommandText = "ZLHIS.SP_EOrder_For_Kodak"
            adoCmd.CommandType = adCmdStoredProc
            adoCmd.ActiveConnection = gcnPACS
            
            
            adoCmd.Parameters.Append adoCmd.CreateParameter("操作类型", adInteger, adParamInput, , Nvl(dsOrder!操作类型))
            adoCmd.Parameters.Append adoCmd.CreateParameter("病人类别", adInteger, adParamInput, , Nvl(dsOrder!病人类别))
            adoCmd.Parameters.Append adoCmd.CreateParameter("医嘱ID", adVarChar, adParamInput, 32, Nvl(dsOrder!医嘱ID))
            adoCmd.Parameters.Append adoCmd.CreateParameter("标识号", adVarChar, adParamInput, 64, Nvl(dsOrder!标识号))
            adoCmd.Parameters.Append adoCmd.CreateParameter("病人ID", adVarChar, adParamInput, 64, Nvl(dsOrder!病人ID))
            adoCmd.Parameters.Append adoCmd.CreateParameter("姓名", adVarChar, adParamInput, 64, Nvl(dsOrder!姓名))
            adoCmd.Parameters.Append adoCmd.CreateParameter("英文名", adVarChar, adParamInput, 64, Nvl(dsOrder!英文名))
            adoCmd.Parameters.Append adoCmd.CreateParameter("性别", adVarChar, adParamInput, 1, Nvl(dsOrder!性别))
            adoCmd.Parameters.Append adoCmd.CreateParameter("出生日期", adVarChar, adParamInput, 16, Nvl(dsOrder!出生日期))
            adoCmd.Parameters.Append adoCmd.CreateParameter("身份证号", adVarChar, adParamInput, 32, Nvl(dsOrder!身份证号))
            adoCmd.Parameters.Append adoCmd.CreateParameter("家庭电话", adVarChar, adParamInput, 128, Nvl(dsOrder!家庭电话))
            adoCmd.Parameters.Append adoCmd.CreateParameter("家庭地址", adVarChar, adParamInput, 256, Nvl(dsOrder!家庭地址))
            adoCmd.Parameters.Append adoCmd.CreateParameter("病区", adVarChar, adParamInput, 32, Nvl(dsOrder!病区))
            adoCmd.Parameters.Append adoCmd.CreateParameter("床号", adVarChar, adParamInput, 32, Nvl(dsOrder!床号))
            adoCmd.Parameters.Append adoCmd.CreateParameter("影像类别", adVarChar, adParamInput, 128, Nvl(dsOrder!影像类别))
            adoCmd.Parameters.Append adoCmd.CreateParameter("检查项目代码", adVarChar, adParamInput, 1024, Nvl(dsOrder!检查项目代码))
            adoCmd.Parameters.Append adoCmd.CreateParameter("检查项目描述", adVarChar, adParamInput, 1024, Nvl(dsOrder!检查项目描述))
            adoCmd.Parameters.Append adoCmd.CreateParameter("开嘱医生", adVarChar, adParamInput, 128, Nvl(dsOrder!开嘱医生))
            adoCmd.Parameters.Append adoCmd.CreateParameter("开嘱科室", adVarChar, adParamInput, 128, Nvl(dsOrder!开嘱科室))
            adoCmd.Parameters.Append adoCmd.CreateParameter("病史", adVarChar, adParamInput, 1024, Nvl(dsOrder!病史))
            adoCmd.Parameters.Append adoCmd.CreateParameter("临床诊断", adVarChar, adParamInput, 1024, Nvl(dsOrder!临床诊断))
            adoCmd.Parameters.Append adoCmd.CreateParameter("注意事项", adVarChar, adParamInput, 1024, Nvl(dsOrder!注意事项))
            adoCmd.Parameters.Append adoCmd.CreateParameter("备注", adVarChar, adParamInput, 156, Nvl(dsOrder!备注))
            
            '返回值
           ' Set adoParaReturn = adoCmd.CreateParameter("返回值", adVarChar, adParamOutput, 4)
            '如果对方存储过程使用的是返回值而不是输出参数，使用这一句读取返回值
             Set adoParaReturn = adoCmd.CreateParameter("返回值", adVarChar, adParamReturnValue, 4)
            adoCmd.Parameters.Append adoParaReturn
            
            adoCmd.Execute
    
            '判断返回值 1001：执行成功,1002：执行失败，此医嘱号已存在,1003：执行失败，未知理由
            '如果操作失败，则记录错误日志
            If adoParaReturn.Value = "1002" Or adoParaReturn.Value = "1003" Then
                subLogErr adoParaReturn.Value, IIf(adoParaReturn.Value = "1002", "执行失败，此医嘱号已存在", "执行失败，未知理由") & _
                    " ,操作类型 = " & Nvl(dsOrder!操作类型) & " , " & _
                    "医嘱ID = " & Nvl(dsOrder!医嘱ID) & " , " & "姓名 = " & Nvl(dsOrder!姓名)
            End If
            
            '操作完成后，删除电子申请单
            
            strSQL = "delete from ZLPACS接口KODAK where ID = " & dsOrder!ID
            gcnHIS.Execute (strSQL)
            
            dsOrder.MoveNext
        Wend
    End If
    
    Exit Sub
errLog:
    subLogErr Err.Number, Err.Description
End Sub

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Private Sub subLogErr(lngErrNo As Long, strDesc As String)
    On Error Resume Next
    Dim lngID As Long
    Dim strSQL As String
    Dim dsRecord As ADODB.Recordset
    
    Me.lblStatus.Caption = Date & " " & Time & vbCrLf & " 发生错误，错误代码：" & lngErrNo & " 错误描述：" & strDesc
    strSQL = "SELECT MAX(ID) as mID FROM ZLPACS接口KODAK_ERR"
    Set dsRecord = gcnHIS.Execute(strSQL)
    If Not dsRecord.EOF Then
        lngID = dsRecord!Mid + 1
    End If
 '   strSQL = "insert into ZLPACS接口KODAK_ERR (ID,错误号,错误描述,错误时间) values(" & lngID & "," _
 '           & lngErrNo & ",'" & Replace(strDesc, "'", "''") & "',sysdate)"
 '   gcnHIS.Execute strSQL
     '查找目录，如果不存在则创建
     '错误存储到磁盘以日期方式存储
    If Dir("D:\ZLPACS接口KODAK_ERR", vbDirectory) = "" Then
        MkDir "D:\ZLPACS接口KODAK_ERR\"
    End If
    
    '创建文件
    Err = 0
    Open "D:\ZLPACS接口KODAK_ERR\" & Date & ".txt" For Append As #1
    Print #1, Date & " " & Time & vbCrLf & "医嘱id:" & lngID & " 发生错误，错误代码：" & lngErrNo & " 错误描述：" & strDesc
    Close #1

End Sub


