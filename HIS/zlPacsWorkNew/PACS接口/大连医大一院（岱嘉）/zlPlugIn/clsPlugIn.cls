VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPlugIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'---------------------------------------------------------------------------
'功能：外挂功能接口部件，用于一些工作站集成外挂调用功能，并传入相应的业务参数
'编制人：周韬
'编制时间：2010-07-22
'说明：
'  2010-07-23:集成到了门诊/住院医嘱虚拟模块中
'  2010-11-25:增加了对自动调用功能的支持,以及对医技站前提ID的传入
'---------------------------------------------------------------------------
Private mblnInit As Boolean

Public Sub Initialize(ByVal cnOracle As ADODB.Connection, ByVal lngSys As Long, ByVal lngModual As Long)
'功能：接口初始化
'参数：cnOracle=活动连接
'      lngSys,lngModual=当前调用接口的主程序系统号及模块号
    If Not mblnInit Then
        Set gcnOracle = cnOracle
        Set gComLib = CreateObject("zlComFunc.clsComLib")
        Set gDatabase = CreateObject("zlComFunc.clsDatabase")
       
        gComLib.InitCommon cnOracle
        
        mblnInit = True
    End If
End Sub

Public Function GetFuncNames(ByVal lngSys As Long, ByVal lngModual As Long, Optional ByVal strReserve As String) As String
'功能：获取外挂功能列表；该接口一般在初始化之后调用一次，调用时机和频度可能因不同的主程序而不同
'参数：lngSys,lngModual=当前调用接口的上级系统号及模块号
'      strReserve=保留参数,用于扩展使用
'        医技站中的调用时,传入为当前执行的前提医嘱ID
'返回：外挂功能列表，格式为"功能名1,功能名2,功能名3,..."
'      如果某个功能需要在病人切换(或刷新)时自动调用，则在功能名前面加上"Auto:"前缀，这种功能只允许有一个。
    Select Case lngModual
        Case 1252
            GetFuncNames = "Auto:门诊医生自动提示," & gstrFunc_PACS影像调阅
        Case 1253
            GetFuncNames = "Auto:住院医生自动提示,护理病历," & gstrFunc_PACS影像调阅
        Case 1254
            GetFuncNames = "Auto:住院护士自动提示,护理病历"
    End Select
End Function
    
Public Sub ExecuteFunc(ByVal lngSys As Long, ByVal lngModual As Long, _
    ByVal strFuncName As String, ByVal lngPatiID As Long, ByVal varRecId As Variant, _
    ByVal varKeyId As Variant, Optional ByVal strReserve As String)
'功能：执行某个外挂功能
'参数：lngSys,lngModual=当前调用接口的上级系统号及模块号
'      strFuncName=要执行的功能名，该功能名之前由GetFuncNames函数返回
'      lngPatiID=当前病人ID
'      varRecId=数字或者字符串；对门诊病人，为当前挂号单号或者挂号ID；对住院病人，为当前住院主页ID
'      varKeyId=数字或者字符串；当前的关键业务数据唯一标识ID，如医嘱ID
'      strReserve=保留参数,用于扩展使用

'156 1260    门诊医生工作站  门诊医生对接诊病人进行写病历和下医嘱等相关工作事务的处理    100 zl9CISJob
'157 1261    住院医生工作站  医生对在院病人写病历和下医嘱等事务和查阅护理记录的处理  100 zl9CISJob
'158 1262    住院护士工作站  护士进行护理记录、医嘱处理与执行以及病历查看等工作的处理    100 zl9CISJob
On Error GoTo ErrH
    Select Case lngModual
        Case 1252
            '门诊医生站
            Select Case strFuncName
                Case "门诊医生自动提示"
                    With frm门诊
                        .PatiID = lngPatiID
                        .RecId = varRecId
                        .KeyId = varKeyId
                        .Reserve = strReserve
                        .RefreshData
                    End With
                Case gstrFunc_PACS影像调阅
                    Call ShowPacsViewer(varKeyId)
                Case Else '其它功能
                    
            End Select
        Case 1253, 1254
            Select Case strFuncName
                Case "住院医生自动提示", "住院护士自动提示"
                    With frm住院
                        .PatiID = lngPatiID
                        .RecId = varRecId
                        .KeyId = varKeyId
                        .Reserve = strReserve
                        .RefreshData
                    End With
                Case "护理病历" '其它功能
                    Dim StrParm As String, strSQL As String, rsTmp As ADODB.Recordset
                    Dim objFile As TextStream, strLine As String, strCmd As String
                    Dim lngR  As Long
                    Dim objStream As TextStream
                    Dim strLog As String
                    
                    On Error GoTo ErrhLis
                    
102                 If lngPatiID = 0 Or Val(varRecId) = 0 Then
104                     MsgBox "请选择一个病人之后再执行此操作！"
                        Exit Sub
                    End If
106                 If gFso.FileExists(App.Path & "\zlPlugin.ini") Then
108                     Set objFile = gFso.OpenTextFile(App.Path & "\zlPlugin.ini")
110                     Do Until objFile.AtEndOfLine
112                         strLine = objFile.ReadLine
114                         If UCase(strLine) Like "CMD*=*" Then
116                             strCmd = Trim(Split(strLine, "=")(1))
118                         ElseIf UCase(strLine) Like "DEBUG*=*" Then
120                             gblnDebug = Val(Split(strLine, "=")(1))
122                         ElseIf UCase(strLine) Like "PARM*=*" Then
124                             StrParm = Trim(Split(strLine, "=")(1))
                            End If
                        Loop
                    End If
126                 If strCmd = "" Then strCmd = "C:\Program Files\嘉和电子病历平台软件\vbRunEmr.exe"
        
128                 If Not gFso.FileExists(strCmd) Then
130                     MsgBox "缺少文件" & strCmd
                        Exit Sub
                    End If
132                 If gblnDebug And StrParm <> "" Then
                        '调试模式，手工指定了参数，直接调用
134                    WriteLog "调试模式：" & strCmd & " " & StrParm
        
136                    lngR = Shell(Chr(34) & strCmd & Chr(34) & " " & StrParm, vbNormalFocus)
138                    If lngR = 0 Then WriteLog "启动失败！"
                       Exit Sub
                    End If
140                 strSQL = "Select User As 编号, a.姓名, b.Id As 编码, b.名称" & vbNewLine & _
                        "From 人员表 A, 部门表 B, 部门人员 C" & vbNewLine & _
                        "Where a.Id = c.人员id And c.部门id = b.Id And (b.撤档时间 Is Null Or b.撤档时间 = To_Date('3000-01-01', 'yyyy-MM-dd')) And" & vbNewLine & _
                        "      (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'yyyy-MM-dd')) And c.缺省 = 1 And a.姓名 = Zl_Username"
        
142                 Set rsTmp = gcnOracle.Execute(strSQL)
                    
144                 StrParm = ""
146                 strLog = ""
148                 Do Until rsTmp.EOF
150                     StrParm = Trim("" & rsTmp!编号) & " " & Replace(Trim("" & rsTmp!姓名), " ", "") & " "
152                     StrParm = StrParm & Trim("" & rsTmp!编码) & " " & Replace(Trim("" & rsTmp!名称), " ", "") & " "
154                     strLog = "获取当前操作员成功！" & StrParm
156                     rsTmp.MoveNext
                    Loop
158                 If StrParm <> "" Then
160                     strSQL = "Select a.住院号, a.出院科室id As 科室代码 From 病案主页 A Where a.病人id = " & lngPatiID & " And a.主页id = " & Val(varRecId)
162                     Set rsTmp = gcnOracle.Execute(strSQL)
164                     StrParm = StrParm & lngPatiID & " " & Val(varRecId) & " "
                        
166                     strLog = strLog & vbNewLine & "获取病人ID，住院次数！" & lngPatiID & Val(varRecId)
168                     Do Until rsTmp.EOF
170                         StrParm = StrParm & rsTmp!住院号 & " " & rsTmp!科室代码 & " " & "*"
172                         strLog = strLog & vbNewLine & "获取病人住院号，科室码成功！" & StrParm
174                         rsTmp.MoveNext
                        Loop
                    End If
176                 If Right(StrParm, 1) = "*" Then
178                     WriteLog "执行命令：" & strCmd & " " & StrParm
180                     If Shell(Chr(34) & strCmd & Chr(34), vbNormalFocus) = 0 Then
182                         If gblnDebug Then WriteLog "启动失败！"
                        End If
                    Else
184                     If gblnDebug Then WriteLog "参数不全：" & strLog
                    End If
                Case gstrFunc_PACS影像调阅
                    If lngModual = 1253 Then    ' 住院医生站调用PACS影像
                        Call ShowPacsViewer(varKeyId)
                    End If
            End Select
    End Select
    Exit Sub
ErrH:
    MsgBox err.Description, vbCritical, "系统消息"
    Exit Sub
ErrhLis:
186 MsgBox CStr(Erl()) & "行出现错误，" & err.Description
188 WriteLog CStr(Erl()) & "行," & err.Description
    Exit Sub
End Sub

Public Sub Terminate(ByVal lngSys As Long, ByVal lngModual As Long)
'功能：接口终止
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号
    Select Case lngModual
        Case 1252
            If Not frm门诊 Is Nothing Then Unload frm门诊
        Case 1253
            If Not frm住院 Is Nothing Then Unload frm住院
        Case 1254
            If Not frm住院 Is Nothing Then Unload frm住院
        Case Else
            
    End Select
End Sub


