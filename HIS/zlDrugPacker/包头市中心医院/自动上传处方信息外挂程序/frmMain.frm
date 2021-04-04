VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HIS数据上传 v1.0"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7110
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   7110
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdStore 
      Caption         =   "获取设备药品库存"
      Height          =   360
      Left            =   2400
      TabIndex        =   1
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton cmdDept 
      Caption         =   "部门数据上传(&D)"
      Height          =   345
      Left            =   4200
      TabIndex        =   2
      Tag             =   "0"
      Top             =   5760
      Width           =   1575
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "退出(&E)"
      Height          =   350
      Left            =   5880
      TabIndex        =   3
      Top             =   5760
      Width           =   1100
   End
   Begin VB.Timer TimerTrans 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame fraH 
      Height          =   45
      Left            =   120
      TabIndex        =   4
      Top             =   5520
      Width           =   6885
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "开始上传(&S)"
      Height          =   350
      Left            =   120
      TabIndex        =   0
      Tag             =   "0"
      Top             =   5760
      Width           =   1335
   End
   Begin VB.ListBox lstLog 
      Height          =   5280
      ItemData        =   "frmMain.frx":030A
      Left            =   120
      List            =   "frmMain.frx":030C
      TabIndex        =   5
      Top             =   120
      Width           =   6855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlng轮询间隔 As Long
Private mint查询天数 As Integer
Private mstr开始时间 As String
Private mstr结束时间 As String
Private mstr药房id As String
Private mblnPackerConnect As Boolean
Private mblnExit As Boolean
Private mstrUserCode As String
Private mstrUserName As String

Private Sub AutoTrans()
    Dim rsData As ADODB.Recordset
    Dim strUserCode As String
    Dim strUserName As String
    Dim strReturn As String
    
    On Error GoTo errHandle
    
    '更新日期范围
    Call UpdateDateValue
    
    Me.cmdStart.Enabled = False
    
    Call OutputLog("开始读取处方信息")
        
    '按NO分批上传数据，包含住院数据（门诊用单据=8和已收费=1区分，住院用单据=9区分）
    gstrSql = "Select 单据, NO " & vbNewLine & _
        " From 未发药品记录 " & vbNewLine & _
        " Where (单据 = 8 And 已收费 = 1 Or 单据 = 9) And Nvl(是否上传, 0) = 0 And 填制日期 Between [1] And [2] And " & vbNewLine & _
        " 库房id In (Select * From Table(Cast(f_Num2list([3], ';') As Zltools.t_Numlist))) "
    Set rsData = OpenSQLRecord(gstrSql, "AutoTrans", CDate(mstr开始时间), CDate(mstr结束时间), mstr药房id)
    
    If Not gobjPacker Is Nothing And mblnPackerConnect = True Then
        If rsData.EOF = False Then
            Do While Not rsData.EOF
                Call gobjPacker.DYEY_MZ_TransRecipeDetail(1, mstrUserCode, mstrUserName, 0, rsData!单据 & "," & rsData!NO, strReturn)
                
                LogListItem "处方上传成功：" & rsData!NO
                Call OutputLog("处方上传成功：" & rsData!NO)
                rsData.MoveNext
            Loop
            LogListItem "本次上传数据完成！" & Now
            Call OutputLog("本次上传数据完成！" & Now)
        Else
            LogListItem "本次无数据上传！" & Now
            Call OutputLog("本次无数据上！" & Now)
        End If
    Else
        LogListItem "WebService地址不正确！" & Now
        Call OutputLog("WebService地址不正确！" & Now)
    End If
  
    Me.cmdStart.Enabled = True
    
    Exit Sub
    
errHandle:
    Me.cmdStart.Enabled = True
    LogListItem Err.Description
End Sub

Private Sub cmdDept_Click()
    Dim strMsg As String
    
    If gobjPacker Is Nothing Then Exit Sub
    
    On Error GoTo hErr
    
    cmdDept.Enabled = False
    If gobjPacker.DYEY_MZ_TransDept("", mstrUserCode, mstrUserName, strMsg) Then
        LogListItem "部门数据上传成功！" & Now
    Else
        LogListItem "部门数据上传失败，请查看日志文件确定原因！" & Now
    End If
    cmdDept.Enabled = True
    Exit Sub
    
hErr:
    cmdDept.Enabled = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdStart_Click()
    Dim strMsg As String
    
    If cmdStart.Tag = "0" Then
        If mblnPackerConnect = False Then
            '重新初始化
            TimerTrans.Enabled = False
            mblnPackerConnect = gobjPacker.DYEY_MZ_IniSoap(False, strMsg, , gcnOracle, 1)
            If mblnPackerConnect = False Then
                MsgBox "初始化接口部件失败：Soap初始化失败！", vbInformation, "提示信息"
                Call OutputLog("初始化接口部件失败：Soap初始化失败")
                Exit Sub
            End If
        End If
                
        cmdStart.Tag = "1"
        cmdStart.Caption = "停止上传(&S)"
        
        '开始上传
        TimerTrans.Enabled = True
        
        LogListItem "开始上传：" & Now
        
    Else

        cmdStart.Tag = "0"
        cmdStart.Caption = "开始上传(&S)"
        
        '停止上传
        TimerTrans.Enabled = False
        
        LogListItem "停止上传" & Now
        
    End If
    
    cmdDept.Enabled = cmdStart.Tag = "0"

End Sub

Private Sub cmdStore_Click()
    Dim strStore As String
    
    cmdStore.Enabled = False
    Call ReadDeviceStore(strStore)
    Call WriteZLHIS(strStore)
    cmdStore.Enabled = True
End Sub

Private Sub Form_Activate()
    If mblnExit Then Unload Me
End Sub

Private Sub Form_Load()
    Dim strMsg As String
    
    '读取注册表参数
    mlng轮询间隔 = Val(GetSetting("ZLSOFT", "公共模块\自动发药机", "轮询间隔", 60))
    mint查询天数 = Val(GetSetting("ZLSOFT", "公共模块\自动发药机", "查询天数", 0))
    mstr药房id = GetSetting("ZLSOFT", "公共模块\自动发药机", "门诊药房", "")
    mblnExit = False
    
    If mstr药房id = "" Then
        MsgBox "未注册药房信息，初始化失败！", vbInformation, ""
        Call OutputLog("未注册药房信息，初始化失败！")
        Unload Me
    End If
    
    If mlng轮询间隔 > 60 Then
        mlng轮询间隔 = 60
    End If
    TimerTrans.Interval = mlng轮询间隔 * 1000
    
    '取用户信息
    Call GetUserInfo
    
    '更新日期范围
    Call UpdateDateValue
    
    '自动发药机接口
    On Error Resume Next
    Set gobjPacker = CreateObject("zlDrugPacker.clsDrugPacker")
    Err.Clear
    If gobjPacker Is Nothing Then
        MsgBox "初始化自动发药机接口部件失败！", vbInformation, "提示信息"
        Call OutputLog("初始化接口部件失败：创建接口部件失败")
        mblnExit = True
        Exit Sub
    Else
        mblnPackerConnect = gobjPacker.DYEY_MZ_IniSoap(False, strMsg, , gcnOracle, 1)
        If mblnPackerConnect = False Then
            MsgBox "初始化接口部件失败：Soap初始化失败！", vbInformation, "提示信息"
            Call OutputLog("初始化接口部件失败：Soap初始化失败")
            Exit Sub
        End If
    End If
    
    Call OutputLog("读取参数：" & "轮询间隔=" & mlng轮询间隔 & "," & "查询天数=" & mint查询天数 & "," & "药房ID=" & mstr药房id)
    Call OutputLog("初始化接口部件成功")
End Sub

Private Sub UpdateDateValue()
    If mint查询天数 = 0 Then
        '默认是当天
        mstr开始时间 = Format(Currentdate, "YYYY-MM-DD")
        mstr结束时间 = Format(Currentdate, "YYYY-MM-DD 23:59:59")
    Else
        '指定天数内
        If mint查询天数 > 3 Then
            mint查询天数 = 3
        End If
        
        mstr开始时间 = Format(DateAdd("d", -mint查询天数, Currentdate), "YYYY-MM-DD")
        mstr结束时间 = Format(Currentdate, "YYYY-MM-DD 23:59:59")
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If cmdStart.Enabled = False Then Cancel = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    Set gobjComLib = Nothing
End Sub

Private Sub TimerTrans_Timer()
    TimerTrans.Enabled = False
    
    On Error GoTo errHandle
    
    '检查连接
    If gcnOracle.State <> adStateOpen Then
        gcnOracle.Open
    End If
    
    '调用自动上传程序
    DoEvents
    Call AutoTrans
    DoEvents
    
    TimerTrans.Enabled = True
    
    Exit Sub
    
errHandle:
    Call LogListItem("异常：" & Err.Description)
    TimerTrans.Enabled = True
End Sub

Private Sub LogListItem(ByVal strLog As String)
    Const INT_MAX_LINES As Integer = 200

    Me.lstLog.AddItem strLog
    Me.lstLog.Selected(Me.lstLog.ListCount - 1) = True
    Me.lstLog.TopIndex = Me.lstLog.ListCount - 1
    If lstLog.ListCount >= INT_MAX_LINES Then lstLog.RemoveItem 0

End Sub

Private Sub GetUserInfo()
'取用户信息

    Dim rsData As ADODB.Recordset

    On Error GoTo hErr
    gstrSql = "Select a.编号, a.姓名, b.用户名 From 人员表 A, 上机人员表 B Where a.Id = b.人员id And Upper(用户名) = Upper([1])"
    Set rsData = OpenSQLRecord(gstrSql, "用户信息", gstrUser)
    If rsData.RecordCount > 0 Then
        mstrUserCode = rsData!编号
        mstrUserName = rsData!姓名
    Else
        mstrUserCode = ""
        mstrUserName = ""
    End If
    rsData.Close
    Exit Sub
    
hErr:
    mstrUserCode = ""
    mstrUserName = ""
    MsgBox Err.Description, vbInformation, "提示"
End Sub

Private Sub ReadDeviceStore(ByRef strVar As String)
'读取设备的库存

    Dim strMsg As String

    strVar = ""
    If gobjPacker Is Nothing Then Exit Sub
    
    On Error GoTo hErr
    
    '读取设备库存
    If gobjPacker.DYEY_MZ_TransStockDevice(mstrUserCode, mstrUserName, strVar, strMsg) Then
        LogListItem "设备库存数据下载成功！" & Now
    Else
        LogListItem "设备库存数据下载失败，请查看日志文件确定原因！" & Now
        LogListItem strMsg
    End If
    
    Exit Sub
    
hErr:
    MsgBox Err.Description, vbInformation, "提示信息"
End Sub

Private Sub WriteZLHIS(ByVal strVar As String)
'写入ZLHIS数据表
    
    Dim strSQL As String, strTmp As String
    Dim intReturn As Integer
    Dim strDisp As String, strCode As String
    Dim dblQTY As Double
    Dim cmdInsert As ADODB.Command
    
'<ROOT>
'    <RETCODE>1</RETCODE>
'    <CONSIS_DRUG_BATCHVW>
'        <DISPENSARY>320</DISPENSARY>
'        <DRUG_CODE>1-1009</DRUG_CODE>
'        <QUANTITY>9</QUANTITY>
'    </CONSIS_DRUG_BATCHVW>
'    <CONSIS_DRUG_BATCHVW>
'        <DISPENSARY>86</DISPENSARY>
'        <DRUG_CODE>1-1015</DRUG_CODE>
'        <QUANTITY>12</QUANTITY>
'    </CONSIS_DRUG_BATCHVW>
'</ROOT>
    
    If strVar = "" Then Exit Sub
    
    '解析XML
    If InStr(strVar, "<RETCODE>") <= 0 Then Exit Sub
    
    intReturn = Val(Mid(strVar, InStr(strVar, "<RETCODE>") + 9))
    If intReturn = 1 Then
        '调用成功
        
        On Error GoTo hErr
        
        ''清空数据表
        gcnOracle.Execute "Delete DrugDeviceStoreTemp"
        
        ''回写
        Do While InStr(strVar, "<CONSIS_DRUG_BATCHVW>") > 0
            strVar = Mid(strVar, InStr(strVar, "<CONSIS_DRUG_BATCHVW>") + 29)
            
            'DISPENSARY
            If InStr(strVar, "<DISPENSARY>") > 0 Then
                strDisp = Mid(strVar, InStr(strVar, "<DISPENSARY>") + 12)
                strDisp = Left(strDisp, InStr(strDisp, "</") - 1)
            Else
                strDisp = ""
            End If
            
            'DRUG_CODE
            If InStr(strVar, "<DRUG_CODE>") > 0 Then
                strCode = Mid(strVar, InStr(strVar, "<DRUG_CODE>") + 11)
                strCode = Left(strCode, InStr(strCode, "</") - 1)
            Else
                strCode = ""
            End If
            
            'QUANTITY
            If InStr(strVar, "<QUANTITY>") > 0 Then
                dblQTY = Val(Mid(strVar, InStr(strVar, "<QUANTITY>") + 10))
            Else
                dblQTY = 0
            End If
                        
            '写ZLHIS数据表
            If strCode <> "" Then
                strSQL = "insert into DrugDeviceStoreTemp (库房ID,药品编码,库存数量) values "
                strSQL = strSQL & "(" & IIf(strDisp = "", "null", strDisp) & ","
                strSQL = strSQL & "'" & strCode & "',"
                strSQL = strSQL & dblQTY & ")"
                
                Set cmdInsert = New ADODB.Command
                With cmdInsert
                    .ActiveConnection = gcnOracle
                    .CommandText = strSQL
                    .Execute
                End With
            Else
                Debug.Print "无药品编码！"
            End If
        Loop
    ElseIf intReturn = 0 Then
        LogListItem "设备无库存数据！" & Now
    Else
        LogListItem "设备库存数据下载异常（WebService）！" & Now
    End If
    
    Exit Sub
    
hErr:
    MsgBox Err.Description, vbInformation, "提示信息"
End Sub
