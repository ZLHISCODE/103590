VERSION 5.00
Begin VB.UserControl ctrlComm 
   ClientHeight    =   1560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1710
   Picture         =   "ctrlComm.ctx":0000
   ScaleHeight     =   1560
   ScaleWidth      =   1710
   ToolboxBitmap   =   "ctrlComm.ctx":0842
   Begin VB.Timer timInData 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   15
      Top             =   525
   End
End
Attribute VB_Name = "ctrlComm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strBuffer As String '数据缓冲区
'strSampleInfo：发送的标本信息
'iSendStep：发送步骤。从1开始递增，0代表不执行发送
Private strSampleInfo As String, iSendStep As Integer, dtSendTime As Date, mblnUndo As Boolean, miType As Integer

'Public Event DataReceived()
'Public Event DevOnComm(ByVal comPort As String, ByVal lngEvent As Long, ByVal strR As String)  ' 显示日志事件
'Public Event DevSenComm(ByVal comPort As String, ByVal strR As String, ByVal intErr As Integer)
Public Event DevDecode(ByVal commport As String, ByVal str结果 As String)
Public Event DevRefresh(ByVal lngID As Long)

Public Event ItemUnknown(ByVal commport As String, ByVal strItems As String) '返回未知项
Public Event ReturnCompute(ByVal strReturn As String)  '返回自动计算结果

Private mstrReceiveDir As String  '通讯程序目录
Private mfsoTmp As New FileSystemObject  '文件对象

Private mintIndex As Integer     '这个控件的索引号，可以从g仪器设置中取得设置信息
Private mintMicrobe As Integer    '是否是微生物 1= 微生物
'Private mItem() As Variant       '存入通道码
Private mlng允许发送已核收标本 As Long   '双向通讯中使用的一个参数

Private mlngManID As Long        '主仪器ID,另存为状态时有用

Private mlngDeviceID As Long     '仪器ID,如果用了另存为，是另存为仪器的ID
Private mlngExeDeptID As Long    '检验小组ID
Private mstrAutoCheckMan As String '自动审核人
Private mintAutoQCCalc  As Integer  '自动计算质控 0-不计算 1-要计算

Private int体检处理方式 As Integer  '1-提示，2-修正，3-不修正
Private int门诊处理方式 As Integer  '1-提示，2-修正，3-不修正
Private int住院处理方式 As Integer  '1-提示，2-修正，3-不修正
Private int院外处理方式 As Integer  '1-提示，2-修正，3-不修正

Private mItem() As Variant

Public Property Get CommSetting() As String
    '串口设置参数
    CommSetting = ""
End Property

Public Property Get DevProgName() As String
    '串口设置参数
    DevProgName = ""
End Property

Public Function SendSample(ByVal lngDeviceID As Long, ByVal strSampleDate As String, ByVal strSampleNO As String, Optional strAdviceIDs As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As Boolean
    '主动向仪器发送标本信息
    
    Dim strSendData As String
    On Error GoTo errH
   
    strSampleInfo = GetSampleInfo(lngDeviceID, mlngManID, strSampleDate, strSampleNO, "", strAdviceIDs, iType)
    If strSampleInfo <> "" Then
        strSampleInfo = strSampleInfo & ";" & IIf(blnUndo, 1, 0) & ";" & iType
        Call WriteToSendDir(strSampleInfo)
    End If
    SendSample = True
    
    If strSampleInfo <> "" Then
        gstrSQL = "ZL_检验标本记录_传送(" & lngDeviceID & ",To_Date('" & strSampleDate & "','yyyy-MM-dd HH24:mi:ss'),'" & strSampleNO & "',1," & iType & ")"
        gobjDatabase.ExecuteProcedure gstrSQL, "打传送标志"
    End If
    
    Exit Function
errH:
    WriteLog "SendSample", LOG_错误日志, Err.Number, Err.Description
End Function

Public Property Get PortOpened() As Boolean
    PortOpened = False
End Property

Public Property Get DeviceID() As Long
    DeviceID = mlngDeviceID
End Property

Public Sub InitContrl(ByVal intIndex As Integer, Optional strCmd As String = "")
    '初始化控件
    '   intindex: 索引
    '   strCmd  : 实始化后，要发给通讯程序的指令，现有在ResetExe-重启通讯程序，CloseExe-关闭通讯程序
        Dim tsmTmp As TextStream
        Dim lngSaveAsID As Long, rsTmp As adodb.Recordset, strSQL As String
        Dim strVer As String, strDevVer As String
        On Error GoTo errH
    
100     ReDim mItem(1, 2) As Variant
102     mItem(1, 0) = -1
104     mItem(1, 1) = 0
106     mItem(1, 2) = 2
    
108     timInData.Enabled = False
    
110     If g仪器(intIndex).ID > 0 Then
            '检查仪器目录是否存在，不存在则创建，并拷通讯程序过去
112         mintIndex = intIndex
114         mstrReceiveDir = g仪器(intIndex).通讯目录
116         mstrAutoCheckMan = Trim(g仪器(intIndex).自动审核人)
118         If mstrAutoCheckMan <> "" Then
120             int体检处理方式 = Val(gobjDatabase.GetPara("体检病人信息不一致的处理方式", glngSys, 1208, True, 1))
122             int院外处理方式 = Val(gobjDatabase.GetPara("院外病人信息不一致的处理方式", glngSys, 1208, True, 1))
124             int住院处理方式 = Val(gobjDatabase.GetPara("住院病人信息不一致的处理方式", glngSys, 1208, True, 1))
126             int门诊处理方式 = Val(gobjDatabase.GetPara("门诊病人信息不一致的处理方式", glngSys, 1208, True, 1))
            End If
128         mintAutoQCCalc = Val(g仪器(intIndex).自动计算质控)
130         If Not mstrReceiveDir Like "?:\*" Then mstrReceiveDir = App.Path & "\Dev_" & g仪器(intIndex).ID
132         g仪器(intIndex).通讯目录 = mstrReceiveDir
134         If Not mfsoTmp.FolderExists(mstrReceiveDir) Then
136             Call mfsoTmp.CreateFolder(mstrReceiveDir)
            End If
        
138         If Dir(mstrReceiveDir & "\zlLisReceiveSend.exe") = "" Then mfsoTmp.CopyFile App.Path & "\zlLisReceiveSend.exe", mstrReceiveDir & "\"
        
140         If Dir(mstrReceiveDir & "\ReceiveSend.ini") = "" Then
142             Set tsmTmp = mfsoTmp.CreateTextFile(mstrReceiveDir & "\ReceiveSend.ini")
144             tsmTmp.WriteLine "[RECEIVE_SET]"
146             tsmTmp.WriteLine "类型 = " & g仪器(intIndex).类型
            
148             tsmTmp.WriteLine "COM端口 = " & g仪器(intIndex).COM口
150             tsmTmp.WriteLine "波特率 = " & g仪器(intIndex).波特率
152             tsmTmp.WriteLine "数据位 = " & g仪器(intIndex).数据位
154             tsmTmp.WriteLine "停止位 = " & g仪器(intIndex).停止位
156             tsmTmp.WriteLine "校验位 = " & g仪器(intIndex).校验位
158             tsmTmp.WriteLine "握手 = " & g仪器(intIndex).握手
160             tsmTmp.WriteLine "缓冲大小 = 2048"
            
162             tsmTmp.WriteLine "IP = " & g仪器(intIndex).IP
164             tsmTmp.WriteLine "IP端口 = " & g仪器(intIndex).IP端口
166             tsmTmp.WriteLine "主机 = " & g仪器(intIndex).主机
            
168             tsmTmp.WriteLine "自动应答 = " & g仪器(intIndex).自动应答
170             tsmTmp.WriteLine "字符模式 = " & g仪器(intIndex).字符模式
172             tsmTmp.WriteLine "通讯程序 = " & g仪器(intIndex).通讯程序
174             tsmTmp.WriteLine "通讯周期 = 0.5"
176             tsmTmp.Close
178             Set tsmTmp = Nothing
            End If
            '检查是否启动，没启动则启动
180         If Dir(mstrReceiveDir & "\Lock.txt") = "" Then
                '自动升级部件后再启动
182             strVer = mfsoTmp.GetFileVersion(App.Path & "\zlLisReceiveSend.exe")
184             strDevVer = mfsoTmp.GetFileVersion(mstrReceiveDir & "\zlLisReceiveSend.exe")
186             If strVer > strDevVer And strVer <> "" And strDevVer <> "" Then
188                 mfsoTmp.CopyFile App.Path & "\zlLisReceiveSend.exe", mstrReceiveDir & "\"
                End If
            
190             If strCmd = "" Then Call Shell(mstrReceiveDir & "\zlLisReceiveSend.exe", vbNormalNoFocus)

            Else
                '已启动，则写需要receive执行的命令,如重启接口，关闭等.
192             If strCmd <> "" Then
194                 Set tsmTmp = mfsoTmp.CreateTextFile(mstrReceiveDir & "\Send\" & strCmd & ".txt")
196                 tsmTmp.WriteLine Format(Now, "yyyy-MM-dd HH:mm:ss")
198                 tsmTmp.Close
200                 Set tsmTmp = Nothing
                End If
            End If
        
202         timInData.Enabled = False
204         If mfsoTmp.FolderExists(mstrReceiveDir & "\Result") Then
            
206             mlng允许发送已核收标本 = g仪器(intIndex).可发已核标本
208             mlngManID = g仪器(intIndex).ID
210             mlngDeviceID = mlngManID

            
                '初始化通讯程序,始终从主仪器取
212             strSQL = "Select 通讯程序名,nvl(微生物,0) as 微生物,使用小组ID From 检验仪器 Where ID=[1]"
214             Set rsTmp = gobjDatabase.OpenSqlRecord(strSQL, "取是否微生物", mlngManID)
            
216             If Not rsTmp.EOF Then
218                 mintMicrobe = Nvl(rsTmp(1), 0)
220                 mlngExeDeptID = Nvl(rsTmp(2), 0)
222                 glngExeDeptID = mlngExeDeptID
                End If


                '-----
                '如果另存为模式启用，则换为 关联仪器 的ID
224             If g仪器(intIndex).SaveAsID > 0 Then mlngDeviceID = g仪器(intIndex).SaveAsID
            
                '根据参数设置来决定主仪器的通道码是从哪个地方取
226             If g仪器(intIndex).另存为通道码 = 0 Then mlngManID = mlngDeviceID
                '------
            
228             If mintMicrobe = 1 Then
230                 strSQL = "Select 通道编码,抗生素ID As 项目ID, 2 as 小数位数,b.编码||nvl(b.简码,b.中文名) as 名称 From 仪器细菌对照 A, 检验用抗生素 B Where a.抗生素id = b.Id And a.仪器id = [1] "
                Else
232                 strSQL = "Select a.通道编码, a.项目id, Nvl(a.小数位数, 2) As 小数位数, b.编码 || '-' || Nvl(b.英文名, b.中文名) As 名称," & vbNewLine & _
                            "       LPad(Decode(c.排列序号, Null, b.编码, c.排列序号), 10, '0') As 排列" & vbNewLine & _
                            "From 检验项目 C, 诊治所见项目 B, 检验仪器项目 A" & vbNewLine & _
                            "Where a.项目id = b.Id And a.项目id = c.诊治项目id And a.仪器id = [1] " & vbNewLine & _
                            "Order By LPad(Decode(c.排列序号, Null, b.编码, c.排列序号), 10, '0')"

                
                    '2011-12-07 死锁问题修改，3/5 - 指标排序
                End If
            
234             Set rsTmp = gobjDatabase.OpenSqlRecord(strSQL, "取通道码", mlngManID)
236             If Not rsTmp.EOF Then
238                 mItem = rsTmp.GetRows
                End If

            
                ' 启动计时器，开始监测返回数据
                timInData.Interval = 500
240             timInData.Enabled = True
            End If
        End If
        Exit Sub
errH:
242     WriteLog "initcontrl", LOG_错误日志, Err.Number, CStr(Erl()) & "行," & Err.Description
End Sub

Private Sub timInData_Timer()
    '每两秒触发保存数据的过程
    Dim strData As String
    Dim ii As Long
    On Error GoTo errH
    
    '进入定时器的时候先关闭定时器
    timInData.Enabled = False
    
    '优先处理 样本请求文件
    Call ReadResultDirFileIQ(True)

    '处理结果文件
    Call ReadResultDirFileRE(True)
        
    '保存完所有数据之后再开启定时器
    timInData.Enabled = True
        
    Exit Sub
errH:
    'Resume
    If Err.Number <> 9 Then
        WriteLog "timInData", LOG_错误日志, Err.Number, Err.Description
    End If
End Sub

Private Function ReadResultDirFileIQ(ByVal blnDelete As Boolean) As String
    '读返回文件
    Dim objFolder As Folder
    Dim objStream As TextStream
    Dim objFiles As Files
    Dim objOneFile As File
    Dim strFolder As String
    Dim strFileName As String
    Dim i As Long
    
 
    Dim strLine As String, lngCount As Long
    
    On Error GoTo errH
    
    strFolder = mstrReceiveDir & "\Result"
    Set objFolder = mfsoTmp.GetFolder(strFolder)
    Set objFiles = objFolder.Files
    For Each objOneFile In objFiles
        '为了防止文件更改被创建还未写入数据就被读取,文件的创建时间小于当前时间10秒的文件才会被读取
        If objOneFile.Name Like "IQ" & Format(Now, "yyyyMMdd") & "_*.txt" And objOneFile.DateCreated < Format(Now - 0.0001, "yyyy-mm-dd hh:mm:ss") Then
            '如果找到了匹配的文件就读取该文件
            strFileName = objOneFile.Path
            If mfsoTmp.FileExists(strFileName) Then
                Set objStream = mfsoTmp.OpenTextFile(strFileName, ForReading)
                strLine = ""
                Do
                    If objStream.AtEndOfStream Then Exit Do
                    strLine = strLine & objStream.ReadLine
                Loop
                objStream.Close
                Set objStream = Nothing
                If mfsoTmp.FileExists(strFileName) And blnDelete = True Then mfsoTmp.DeleteFile (strFileName)
'                    ReadResultDirFile = strLine
                '将读取的文件内容保存到数据库
                If strLine <> "" Then
                    WriteLog "timInData-发送请求", LOG_通讯日志, 0, strLine
                    Call WriteSampleInfo(strLine)   '= True Then RaiseEvent DataReceived
                End If
            End If
        End If
        DoEvents
    Next
    Set objFolder = Nothing
    Set objFiles = Nothing
    
    Exit Function
errH:
    WriteLog "ReadResultDirFile", LOG_错误日志, Err.Number, Err.Description
End Function

Private Function ReadResultDirFileRE(ByVal blnDelete As Boolean) As String
    '读返回文件
    Dim objFolder As Folder
    Dim objStream As TextStream
    Dim objFiles As Files
    Dim objOneFile As File
    Dim strFolder As String
    Dim strFileName As String
    Dim ii As Long
    Dim i As Long
    
 
    Dim strLine As String, lngCount As Long
    
    On Error GoTo errH
    
    strFolder = mstrReceiveDir & "\Result"
    Set objFolder = mfsoTmp.GetFolder(strFolder)
    Set objFiles = objFolder.Files
    For Each objOneFile In objFiles
        '为了防止文件更改被创建还未写入数据就被读取,文件的创建时间小于当前时间10秒的文件才会被读取
        If objOneFile.Name Like "RE" & Format(Now, "yyyyMMdd") & "_*.txt" And objOneFile.DateCreated < Format(Now - 0.0001, "yyyy-mm-dd hh:mm:ss") Then
            strFileName = objOneFile.Path
            If mfsoTmp.FileExists(strFileName) Then
                Set objStream = mfsoTmp.OpenTextFile(strFileName, ForReading)
                strLine = ""
                Do
                    If objStream.AtEndOfStream Then Exit Do
                    strLine = strLine & objStream.ReadLine
                Loop
                objStream.Close
                Set objStream = Nothing
                If mfsoTmp.FileExists(strFileName) And blnDelete = True Then mfsoTmp.DeleteFile (strFileName)
                If strLine <> "" Then
                    ii = UBound(mItem, 2)
                    strLine = Replace(strLine, "CHR(10) CHR(13)", vbCrLf)
                    WriteLog "TimInData-保存数据", LOG_通讯日志, 0, strLine
                    Call InDataBase(strLine)  '= True ' Then RaiseEvent DataReceived
                End If
            End If
        End If
        DoEvents
    Next
    Set objFolder = Nothing
    Set objFiles = Nothing
        
    Exit Function
errH:
    WriteLog "ReadResultDirFile", LOG_错误日志, Err.Number, Err.Description
End Function

Private Sub UserControl_Initialize()
    strBuffer = ""
    iSendStep = 0 '开始不执行发送
End Sub

Private Sub Return_Decode(ByVal strDecode As String)
    '返回解码结果
    If strDecode = "" Then Exit Sub
    If g仪器(mintIndex).类型 = 0 Then
        RaiseEvent DevDecode(g仪器(mintIndex).COM口, strDecode)
    Else
        RaiseEvent DevDecode(g仪器(mintIndex).IP, strDecode)
    End If
End Sub

Private Function WriteSampleInfo(ByVal strResult As String) As Boolean
    '双向通讯中，取得标本信息,然后写到通讯目录
    
    Dim aSamples() As String, aSampleInfo() As String, i As Integer, strSampleInfo As String, iType As Integer
    Dim strSampleNO As String, aTmp() As String, strBarcode As String
    
    On Error GoTo errH
    
    If Len(strResult) > 0 Then '要向仪器发送标本信息
        aSamples = Split(strResult, "||")
        strSampleInfo = "": miType = 0
        For i = 0 To UBound(aSamples)
            aSampleInfo = Split(aSamples(i), "|")
            If UBound(aSampleInfo) > 0 Then
                aTmp = Split(aSampleInfo(1), "^")
                If UBound(aTmp) = 0 Then
                    strSampleNO = Val(aTmp(0)): miType = 0: strBarcode = ""
                Else
                    strSampleNO = Val(aTmp(0)): miType = Val(aTmp(1)): strBarcode = ""
                    If UBound(aTmp) > 1 Then
                        strBarcode = Trim(aTmp(2))
                    End If
                End If
                
                '写数据到 通讯程序目录
                strSampleInfo = GetSampleInfo(mlngDeviceID, mlngManID, Format(aSampleInfo(0), "yyyy-MM-dd"), strSampleNO, strBarcode, , miType)
                If strSampleInfo <> "" Then
                    strSampleInfo = strSampleInfo & ";0;" & miType
                    Call WriteToSendDir(strSampleInfo)
                    If UBound(Split(strSampleInfo, "|")) > 2 Then strSampleNO = Split(strSampleInfo, "|")(1)
                    gstrSQL = "ZL_检验标本记录_传送(" & mlngDeviceID & ",To_Date('" & Format(aSampleInfo(0), "yyyy-MM-dd") & "','yyyy-MM-dd'),'" & strSampleNO & "',1," & miType & ")"
                    WriteLog "写传送标志", LOG_通讯日志, 0, gstrSQL
                    gobjDatabase.ExecuteProcedure gstrSQL, "打传送标志"
                End If
                
            End If
        Next

    End If
    Exit Function
errH:
    WriteLog "WriteSampleInfo", LOG_错误日志, Err.Number, Err.Description
End Function

Private Sub WriteToSendDir(ByVal strInput As String, Optional ByVal strFileType As String)
    'strFileType: 写入要发送目录的文件类型，SendSample为要发送给仪器的指令。
    
    
    Dim strFileName As String
    Dim lngCount As Long, lngFileNo As Long
    On Error GoTo errH
    
    If mfsoTmp.FolderExists(mstrReceiveDir & "\Send") = False Then mfsoTmp.CreateFolder (mstrReceiveDir & "\Send")
    lngCount = 0
    If strFileType = "" Then strFileType = "SendSample"
    strFileName = Dir(mstrReceiveDir & "\Send\" & strFileType & "_*.txt")
    If strFileName <> "" Then lngCount = Val(Split(strFileName, "_")(1))
    Do While lngCount < 1000
        lngCount = lngCount + 1
        strFileName = mstrReceiveDir & "\Send\" & strFileType & "_" & Format(lngCount, "000") & ".txt"
        If mfsoTmp.FileExists(strFileName) = False Then
            lngFileNo = FreeFile
            Open strFileName For Binary Access Read Write Lock Read Write As lngFileNo
            Put lngFileNo, , strInput
            Close lngFileNo
            Exit Do
        End If
    Loop
    Exit Sub
errH:
    WriteLog "WriteToSendDir", LOG_错误日志, Err.Number, Err.Description
End Sub

Private Function GetSampleInfo(ByVal lngDeviceID As Long, ByVal lngMainID As Long, ByVal strSampleDate As String, ByVal strSampleNO As String, ByVal strBarcode As String, Optional strAdviceIDs As String = "", Optional ByVal iType As Integer = 0) As String
        '获取需要向仪器发送的标本信息
        '返回：标本信息。
        '   标本之间以||分隔
        '   元素之间以|分隔
        '   第0个元素：检验时间
        '   第1个元素：样本序号
        '   第2个元素：病人姓名
        '   第3个元素：标本类型
        '   第4个元素：急诊标志
        '   第5个元素：样本条码
        '   第6个元素：盘号，杯号
        '   第7个元素：病人ID^性别^出生日期^年龄^姓名全拼^稀释倍数(预留，暂传空)     2013年11月07日 modify by 陈东,
        '   第8～9元素：系统保留
        '   从第10个元素开始为需要的检验项目。
        '  lngDeviceID = 仪器ID
        '  strSampleDate = 日期 格式为 YYYY-MM-DD
        '  strSampleNO = 标本号
        '  strBarcode = 条码
        '  strAdviceIDs =???
        '  iType = 标本类别
        Dim objDevice As Object
        Dim rsTmp As New adodb.Recordset
        Dim lngAdviceID As Long, aAdviceIDs() As String, i As Integer
        Dim bln发送时指定杯号 As Boolean
        Dim strAddInfo As String
        Dim str标本号 As String, int_急诊 As Integer
    
        On Error GoTo DBErr
        '发往仪器的数据，正常情况下是不核收，病人医嘱发送中执行状态为0 ，而指定了杯号的仪器需要先核收,填写杯号后再发送，所以指定杯号的仪器就不管执行状态。
    
100     If mlng允许发送已核收标本 = 0 Then
102         bln发送时指定杯号 = False
104         gstrSQL = "Select 发送时指定杯号 From 检验仪器 Where Id = [1]"
106         Set rsTmp = gobjDatabase.OpenSqlRecord(gstrSQL, "取仪器属性", lngDeviceID)
108         Do Until rsTmp.EOF
110             bln发送时指定杯号 = Val("" & rsTmp!发送时指定杯号) = 1
112             rsTmp.MoveNext
            Loop
        Else
114         bln发送时指定杯号 = True
        End If
    
116     If Len(strAdviceIDs) = 0 Or Val(strAdviceIDs) = 0 Then
118         If Len(Trim(strBarcode)) = 0 Then
                '按标本序号查询, 2013-11-12 加上检验结果为空的不传
120             gstrSQL = "Select TO_CHAR(A.核收时间, 'MM-DD HH24:MI') AS 标本时间,A.标本序号 AS 标本号,F.姓名,D.中文名,D.英文名,C.通道编码,A.标本类型,A.样本条码, A.杯号, E.紧急标志, A.标本类别 ,A.出生日期,A.性别,A.年龄,A.病人id" & _
                    " From 检验标本记录 A,检验普通结果 B,检验仪器项目 C,诊治所见项目 D,病人医嘱记录 E,病人信息 F,检验项目 G,病人医嘱发送 H " & _
                    " Where A.ID+0=B.检验标本ID And A.报告结果=B.记录类型 And B.检验项目ID+0=C.项目ID And C.仪器ID=[6] And B.检验项目ID+0=D.ID" & _
                    " And A.医嘱ID+0=E.ID And E.病人ID+0=F.病人ID And D.ID=G.诊治项目ID And A.仪器ID=[1]" & _
                    " And A.核收时间 BETWEEN [2] AND [3] And E.id=H.医嘱ID " & IIf(bln发送时指定杯号 = True, "", " And H.执行状态 = 0 ") & _
                    " And B.检验结果 Is Null And A.标本序号=[4] And G.项目类别<>3 And C.通道编码<>'0' " & IIf(gblnEmerge, " and nvl(a.标本类别,0)  = [5] ", "")
122             Set rsTmp = gobjDatabase.OpenSqlRecord(gstrSQL, "发送仪器数据", lngDeviceID, CDate(Format(strSampleDate, "yyyy-mm-dd") & " 00:00:00"), _
                    CDate(Format(strSampleDate, "yyyy-mm-dd") & " 23:59:59"), Val(strSampleNO), iType, lngMainID)
124             Call WriteLog("GetSampleInfo", LOG_通讯日志, 0, "按标本查:" & lngDeviceID & "," & strSampleNO & "," & strSampleDate & "," & iType & "," & mlng允许发送已核收标本)
            
            Else
                '按条码查找  2013-11-12 加上检验结果为空的不传
126             gstrSQL = "Select TO_CHAR(A.核收时间, 'MM-DD HH24:MI') AS 标本时间,A.标本序号 AS 标本号,F.姓名,D.中文名,D.英文名,C.通道编码,A.标本类型,A.样本条码, A.杯号, E.紧急标志, A.标本类别 ,A.出生日期,A.性别,A.年龄,A.病人id " & _
                    " From 检验标本记录 A,检验普通结果 B,检验仪器项目 C,诊治所见项目 D,病人医嘱记录 E,病人信息 F,检验项目 G,病人医嘱发送 H" & _
                    " Where A.ID+0=B.检验标本ID And A.报告结果=B.记录类型 And B.检验项目ID+0=C.项目ID And C.仪器ID=[7] And B.检验项目ID+0=D.ID" & _
                    " And A.医嘱ID+0=E.ID And E.病人ID+0=F.病人ID And D.ID=G.诊治项目ID And A.仪器ID=[1]" & _
                    " And A.核收时间 BETWEEN [2] AND [3] And E.id=H.医嘱ID " & IIf(bln发送时指定杯号 = True, "", " And H.执行状态 = 0 ") & _
                    " And B.检验结果 Is Null And A.样本条码=[5] And G.项目类别<>3 And C.通道编码<>'0' "
128             Set rsTmp = gobjDatabase.OpenSqlRecord(gstrSQL, "发送仪器数据", lngDeviceID, CDate(Format(strSampleDate, "yyyy-mm-dd") & " 00:00:00"), _
                    CDate(Format(strSampleDate, "yyyy-mm-dd") & " 23:59:59"), Val(strSampleNO), strBarcode, iType, lngMainID)
130             If rsTmp.EOF Then
                    '查找医嘱  医嘱状态=8 检验医嘱都是临嘱，发送后就是已停止
132                 gstrSQL = "Select TO_CHAR(F.发送时间, 'MM-DD HH24:MI') AS 标本时间,0 AS 标本号,A.紧急标志," & _
                        "C.姓名||Decode(A.婴儿,0,'',Null,'','(婴儿)') As 姓名,Y.通道编码,A.标本部位 As 标本类型,F.样本条码,'' as 杯号, 0 as 标本类别,C.出生日期,C.性别,C.年龄,C.病人ID " & _
                        "FROM 病人医嘱记录 A," & _
                        "病人信息 C,病人医嘱发送 F,检验报告项目 G,检验项目 I,检验仪器项目 Y " & _
                        "WHERE A.诊疗类别 = 'C' " & _
                        "AND A.病人ID=C.病人ID " & IIf(bln发送时指定杯号 = True, "", " And F.执行状态 = 0 ") & _
                        "AND A.相关id IS NOT NULL " & _
                        "AND A.医嘱状态=8 AND A.ID=F.医嘱id " & _
                        "AND A.诊疗项目id=G.诊疗项目id AND G.细菌ID Is Null AND G.报告项目id=Y.项目id " & _
                        "AND G.报告项目ID=I.诊治项目ID " & _
                        "AND Y.仪器ID+0=[1] " & _
                        "And F.样本条码=[2] " & _
                        " And Y.通道编码<>'0' "

134                 Set rsTmp = gobjDatabase.OpenSqlRecord(gstrSQL, "发送仪器数据", lngMainID, strBarcode, iType)
                End If
            End If
136         GetSampleInfo = ""
138         If Not rsTmp.EOF Then
140             int_急诊 = IIf(gblnEmerge, Val("" & rsTmp!紧急标志), 0)

142             GetSampleInfo = Format(rsTmp("标本时间"), "yyyy-MM-dd HH:mm:ss")
144             GetSampleInfo = GetSampleInfo & "|" & Nvl(rsTmp("标本号"), " ")
146             GetSampleInfo = GetSampleInfo & "|" & Nvl(rsTmp("姓名"), " ")
148             GetSampleInfo = GetSampleInfo & "|" & Nvl(rsTmp("标本类型"), " ") & "|" & int_急诊
                '2013-11-07 add by cd 医大二院西门子流水线需要的信息添加
                 
150             strAddInfo = Val("" & rsTmp!病人id) & "^" & Trim("" & rsTmp!性别) & "^" & Format(rsTmp!出生日期, "YYYY-MM-DD") & "^" & Trim("" & rsTmp!年龄) & "^" & _
                             gobjCommFun.mGetFullPY("" & rsTmp("姓名")) & "^"  '此处稀释倍数为空，兼容新版LIS
152             GetSampleInfo = GetSampleInfo & "|" & Nvl(rsTmp("样本条码"), " ") & "|" & IIf(Trim("" & rsTmp("杯号")) = "", " ", Trim("" & rsTmp("杯号"))) & "|" & strAddInfo & "| | "
154             Do While Not rsTmp.EOF
156                 GetSampleInfo = GetSampleInfo & "|" & Nvl(rsTmp("通道编码"), " ")
158                 rsTmp.MoveNext
                Loop
            End If
        Else '按医嘱ID查询
160         aAdviceIDs = Split(strAdviceIDs, ",")
162         GetSampleInfo = ""
164         For i = 0 To UBound(aAdviceIDs)
166             lngAdviceID = Val(aAdviceIDs(i))
        
168             gstrSQL = "Select TO_CHAR(A.核收时间, 'MM-DD HH24:MI') AS 标本时间,A.标本序号 AS 标本号,F.姓名,D.中文名,D.英文名,C.通道编码,A.标本类型,'' As 样本条码, A.杯号, E.紧急标志, A.标本类别,A.出生日期,A.性别,A.年龄,A.病人ID " & _
                    " From 检验标本记录 A,检验项目分布 B,检验仪器项目 C,诊治所见项目 D,病人医嘱记录 E,病人信息 F,检验项目 G,病人医嘱发送 H " & _
                    " Where A.ID=B.标本ID+0 And B.项目ID+0=C.项目ID And C.仪器ID=[4] And B.检验项目ID+0=D.ID" & _
                    " And B.医嘱ID=E.ID And E.病人ID+0=F.病人ID And D.ID=G.诊治项目ID And A.仪器ID=[1] And E.id=H.医嘱ID " & IIf(bln发送时指定杯号 = True, "", " And H.执行状态 = 0 ") & _
                    " And B.医嘱ID=[2] And G.项目类别<>3 And C.通道编码<>'0' " & IIf(gblnEmerge, " and nvl(a.标本类别,0)  = [3] ", "")
170             Set rsTmp = gobjDatabase.OpenSqlRecord(gstrSQL, "发送仪器数据", lngDeviceID, lngAdviceID, iType, lngMainID)
172             If Not rsTmp.EOF Then
174                 If Len(GetSampleInfo) = 0 Then
176                     int_急诊 = Val("" & rsTmp!紧急标志)

178                     GetSampleInfo = Format(rsTmp("标本时间"), "yyyy-MM-dd HH:mm:ss")
180                     GetSampleInfo = GetSampleInfo & "|" & Nvl(rsTmp("标本号"), " ")
182                     GetSampleInfo = GetSampleInfo & "|" & Nvl(rsTmp("姓名"), " ")
184                     GetSampleInfo = GetSampleInfo & "|" & Nvl(rsTmp("标本类型"), " ") & "|" & int_急诊
                        '2013-11-07 add by cd 医大二院西门子流水线需要的信息添加
186                     strAddInfo = Val("" & rsTmp!病人id) & "^" & Trim("" & rsTmp!性别) & "^" & Format(rsTmp!出生日期, "YYYY-MM-DD") & "^" & Trim("" & rsTmp!年龄) & _
                                     "^" & gobjCommFun.mGetFullPY("" & rsTmp("姓名")) & "^"   '此处稀释倍数为空，兼容新版LIS'此处稀释倍数为空，兼容新版LIS
188                     GetSampleInfo = GetSampleInfo & "|" & Nvl(rsTmp("样本条码"), " ") & "|" & IIf(Trim("" & rsTmp("杯号")) = "", " ", Trim("" & rsTmp("杯号"))) & "|" & strAddInfo & "| | "
                    End If
190                 Do While Not rsTmp.EOF
192                     GetSampleInfo = GetSampleInfo & "|" & Nvl(rsTmp("通道编码"), " ")
                
194                     rsTmp.MoveNext
                    Loop
                End If
            Next
        End If
196     Call WriteLog("getSampleInfo", LOG_通讯日志, 0, GetSampleInfo)
        Exit Function
DBErr:
198     GetSampleInfo = ""
200     Call WriteLog("GetSampleInfo", LOG_错误日志, Err.Number, CStr(Erl()) & "行出现错误：  " & Err.Description)
End Function

Private Function InDataBase(ByVal strResult As String) As Boolean
        '保存数据到数据库
        Dim strErr As String, lngErr As Long, strQCComputeInfo As String, strUnkonw As String
        Dim strIDs As String, varIds As Variant, i As Integer
        Dim strlogs As String
        On Error GoTo hErr
    
100     If SaveToDataBase(mlngDeviceID, mlngManID, mlngExeDeptID, mintMicrobe, mintAutoQCCalc, mstrAutoCheckMan, strResult, mItem, strUnkonw, strQCComputeInfo, lngErr, strErr, strIDs, strlogs) = True Then
102         InDataBase = True
104         If strIDs <> "" Then
106             varIds = Split(strIDs, ",")
108             For i = LBound(varIds) To UBound(varIds)
110                 If Val("" & varIds(i)) <> 0 Then
112                     RaiseEvent DevRefresh(Val("" & varIds(i)))
                    End If
                Next
            End If
114         If strQCComputeInfo <> "" Then
116             RaiseEvent ReturnCompute(strQCComputeInfo)
            End If
118         If strUnkonw <> "" Then
120             If g仪器(mintIndex).类型 = 0 Then
122                 RaiseEvent ItemUnknown(g仪器(mintIndex).COM口, strUnkonw)
                Else
124                 RaiseEvent ItemUnknown(g仪器(mintIndex).IP, strUnkonw)
                End If
            End If
126         If strlogs <> "" Then
128             Call WriteToSendDir(strlogs, "SaveDataLog")
            End If
        End If
        Exit Function
hErr:
130     Call WriteLog("GetSampleInfo", LOG_错误日志, Err.Number, CStr(Erl()) & "行出现错误：  " & Err.Description)
End Function








