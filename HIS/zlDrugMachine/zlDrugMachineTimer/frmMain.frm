VERSION 5.00
Begin VB.Form frmMain 
   ClientHeight    =   2295
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3315
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2295
   ScaleWidth      =   3315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer timTransmit 
      Enabled         =   0   'False
      Left            =   360
      Top             =   240
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event AfterTransmit(ByVal strLog As String, ByVal blnTransmitting As Boolean)
Private mobjComLib As Object
Private mobjMachine As Object
Private mstrBaseData As String                              '基础数据传送；详见clsDataTransmit.Transmit的参数说明
Private mclsLog As clsLog
Private mtypParams As TYPE_PARAMS
Private mblnTransmitting As Boolean                         'True正在传送数据
Private mstrUser As String

'基础数据传送
Public Property Get BaseData() As String
    BaseData = mstrBaseData
End Property
Public Property Let BaseData(ByVal strValue As String)
    mstrBaseData = strValue
End Property

'目前支持定时传送的业务数据，供管理工具选择
'如果有调整定时业务数据传送，请同步修改 timTransmit_Timer() 事件等
Public Property Get SupportData() As String
    SupportData = "门诊收费|门诊退费（整单）|门诊发药|住院发药" '& "|窗口状态（蝶和）"
End Property

Public Function ShowMe(ByVal strUser As String, ByVal objComLib As Object, ByVal clsVar As clsLog) As Boolean
    Dim strMsg As String, strTmp As String

    mstrUser = strUser
    Set mclsLog = clsVar
    
    If mobjMachine Is Nothing Then
        On Error Resume Next
        Set mobjMachine = CreateObject("zlDrugMachine.clsDrugMachine")
        If Err.Number <> 0 Then
            strTmp = "请确保zlDrugMachine部件已注册！"
            mclsLog.Add strTmp
            mclsLog.Save
            RaiseEvent AfterTransmit(strTmp, mblnTransmitting)
            Exit Function
        End If
        On Error GoTo 0
    End If
    
    strTmp = "初始化药品自动化设备接口部件。"
    mclsLog.Add vbNewLine & "" & Now
    mclsLog.Add strTmp
    RaiseEvent AfterTransmit(strTmp, mblnTransmitting)
    
    On Error GoTo hErr

    Set mobjComLib = objComLib

    If mobjMachine.Init(Val("2-管理工具"), mobjComLib, strMsg) = False Then
        mclsLog.Add strMsg
        RaiseEvent AfterTransmit(strMsg, mblnTransmitting)
        strTmp = "初始化zlDrugMachine部件失败！"
        mclsLog.Add strTmp
        mclsLog.Save
        RaiseEvent AfterTransmit(strTmp, mblnTransmitting)
        Exit Function
    End If
    mclsLog.Save

    Me.Show
    Me.Visible = False
    
    ShowMe = True

    Exit Function
    
hErr:
    strTmp = Err.Description
    mclsLog.Add strTmp
    mclsLog.Save
    RaiseEvent AfterTransmit(strTmp, mblnTransmitting)
End Function

Public Sub ReadParams()
'功能：读取参数，并保存到变量
    
    Dim objXML As New clsXML
    Dim strFile As String, strTmp As String
    
    If LCase(App.Path) Like "*\apply" Then
        strFile = App.Path & "\zlDrugMachine.cfg"
    ElseIf LCase(App.Path) Like "*zldrugmachinemanage*" Or LCase(App.Path) Like "*zldrugmachine\*" Or LCase(App.Path) Like "*zldrugmachine" Then
        strFile = Replace(App.Path, App.EXEName, "") & "zlDrugMachineManage\zlDrugMachine.cfg"
    End If
    
    If objXML.OpenXMLFile(strFile) = False Then
        strTmp = "管理工具的参数文件不正确！"
        RaiseEvent AfterTransmit(strTmp, mblnTransmitting)
    End If

    With mtypParams
        .定时周期 = Val(GetParameter(objXML, "cycle", "5"))
        .有效天数 = Val(GetParameter(objXML, "validdays", "2"))
        .显示最大行数 = Val(GetParameter(objXML, "viewlines", "200"))
        .输出日志 = Val(GetParameter(objXML, "output", "0")) = 1
        .详细日志 = Val(GetParameter(objXML, "detailed", "0")) = 1
        .保存日志天数 = Val(GetParameter(objXML, "savedays", "7"))
        .业务数据 = Trim(GetParameter(objXML, "businessdata", ""))
    End With
    
    objXML.CloseXMLDocument
    Set objXML = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobjMachine = Nothing
    Set mobjComLib = Nothing
End Sub

Private Sub timTransmit_Timer()
    Dim strSQL As String, strTmp As String, strINF As String
    Dim rsTemp As ADODB.Recordset

    '由于Timer最大支持65535毫秒，因此，通过变通方式实现大小65秒的定时事务
    If (Timer() - Val(Me.Tag)) \ 60 < mtypParams.定时周期 Then Exit Sub
    
    timTransmit.Tag = "1"   '开始定时传送
    timTransmit.Enabled = False
    
    On Error GoTo hErr
    
    mclsLog.Add vbNewLine & "" & Now
    
    mblnTransmitting = True
    
    strTmp = "开始业务数据定时传送"
    mclsLog.Add strTmp
    RaiseEvent AfterTransmit(strTmp, mblnTransmitting)
    
    strTmp = "|" & mtypParams.业务数据 & "|"
    
    strSQL = _
        "Select Distinct a.NO 处方号, a.单据, a.库房id  " & vbNewLine & _
        "From 药品收发记录 A, 药品收发门诊标志 B " & vbNewLine & _
        "Where a.No = b.处方号 And a.单据 = b.单据 " & vbNewLine & _
        "    And a.填制日期 Between Sysdate - [2] And Sysdate And b.业务分类 = [1] And Instr(';0;11;12;', ';' || Nvl(b.标志, 0) || ';') > 0 "
    
    '门诊收费
    If InStr(strTmp, "|门诊收费|") > 0 Then
        Call TransBusiness(21, strSQL, "1-门诊收费")
    End If
    
    '门诊退费（整单）
    If InStr(strTmp, "|门诊退费（整单）|") > 0 Then
        Call TransBusiness(25, strSQL, "2-门诊退费（整单）")
    End If
    
    '门诊发药
    If InStr(strTmp, "|门诊发药|") > 0 Then
        Call TransBusiness(22, strSQL, "3-门诊发药")
    End If
    
    '住院发药
    If InStr(strTmp, "|住院发药|") > 0 Then
        strSQL = _
            "Select b.收发id " & vbNewLine & _
            "From 药品收发记录 A, 药品收发住院标志 B " & vbNewLine & _
            "Where a.Id = b.收发id And a.填制日期 Between Sysdate - [2] And Sysdate And b.业务分类 = [1] And b.标志 > 10 "

        Call TransBusiness(21, strSQL, "4-住院发药")
    End If
    
    '蝶和设备需要的药房窗口状态通知
    If InStr(strTmp, "|窗口状态（蝶和）|") > 0 Then
        '所有药房的发药窗口状态
        strSQL = _
            "Select f_List2str(Cast(Collect(编号) As t_Strlist), ';') 接口编号 " & vbNewLine & _
            "From 药品设备接口 " & vbNewLine & _
            "Where 停用日期 Is Null "
        Set rsTemp = mobjComLib.zlDatabase.OpenSQLRecord(strSQL, "获取药品设备接口")
        If rsTemp.RecordCount > 0 Then
            strINF = IIf(IsNull(rsTemp!接口编号), "", rsTemp!接口编号)
        End If
        rsTemp.Close

        Call TransBase(strINF & "|5")
    End If
    
    '检查有无基础数据传送
    If mstrBaseData <> "" Then
        strTmp = "开始传送基础数据"
        mclsLog.Add strTmp
        RaiseEvent AfterTransmit(strTmp, mblnTransmitting)
        
        Call TransBase(mstrBaseData)
        
        mblnTransmitting = False
        strTmp = "完成传送基础数据"
        mclsLog.Add strTmp
        mclsLog.Save
        RaiseEvent AfterTransmit(strTmp, mblnTransmitting)
        
        mstrBaseData = ""
    Else
        mblnTransmitting = False
        strTmp = "完成业务数据定时传送"
        mclsLog.Add strTmp
        mclsLog.Save
        RaiseEvent AfterTransmit(strTmp, mblnTransmitting)
    End If
    
    Me.Tag = Timer()
    timTransmit.Enabled = True
    timTransmit.Tag = ""    '完成定时传送
    
    Exit Sub
    
hErr:
    mblnTransmitting = False
    timTransmit.Enabled = True
    timTransmit.Tag = ""
End Sub

Private Sub TransBusiness(ByVal intType As Integer, ByVal strSQL As String, ByVal strInfo As String)
'功能：传送业务数据
'参数：
'  intType：业务类别
'           21-配药[门诊和住院处方明细上传]；
'           22-开始发药；
'           23-完成发药；
'           25-处方完整退药；
'  strInfo：1-门诊收费；2-门诊退费（整单）；3-门诊发药；4-住院发药
    
    Dim rsTemp As ADODB.Recordset
    Dim strData As String, strBill As String, strMsg As String, strTmp As String
    
    On Error GoTo hErr
    
    strTmp = "获取“" & strInfo & "”数据"
    mclsLog.Add strTmp
    mclsLog.Add strSQL, 1, Val("1-详细日志")
    RaiseEvent AfterTransmit(strTmp, mblnTransmitting)
    
'    With mclsLog
'        .Add "查询变量：", 1, Val("1-详细日志")
'        .Add strInfo, 1, Val("1-详细日志")
'        .Add mtypParams.有效天数, 1, Val("1-详细日志")
'    End With
    
    Set rsTemp = mobjComLib.zlDatabase.OpenSQLRecord(strSQL, strInfo, intType - 20, mtypParams.有效天数)
    Do While rsTemp.EOF = False
        If Val(strInfo) = Val("2-门诊退费（整单）") Then
            strBill = strBill & ";" & rsTemp!单据 & "," & rsTemp!处方号 & "," & IIf(IsNull(rsTemp!库房id), "", rsTemp!库房id)
        Else
            strBill = strBill & ";" & rsTemp!单据 & "," & rsTemp!处方号
        End If
        
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    
    mclsLog.Add strBill, 1, Val("1-详细日志")
    
    If Left(strBill, 1) = ";" Then strBill = Mid(strBill, 2)
    
    Select Case Val(strInfo)
    Case Val("1-门诊收费"), Val("3-门诊发药")
        strData = "1|" & strBill
        
    Case Val("2-门诊退费（整单）")
        strData = strBill
        
    Case Val("4-住院发药")
        strData = "2|" & strBill
        
    End Select
    
    If strBill = "" Then
        strTmp = "“" & strInfo & "”数据暂无"
        mclsLog.Add strTmp
        mclsLog.Save
        RaiseEvent AfterTransmit(strTmp, mblnTransmitting)
        Exit Sub
    End If
    
    strTmp = "开始传送“" & strInfo & "”数据"
    mclsLog.Add strTmp
    RaiseEvent AfterTransmit(strTmp, mblnTransmitting)
    
    '传送数据
    If mobjMachine.Operation(mstrUser, intType, strData, strMsg) Then
        '正常日志输出
        strTmp = "传送“" & strInfo & "”数据成功"
    Else
        '异常日志输出
        strTmp = "传送“" & strInfo & "”数据失败"
    End If
    mclsLog.Add strTmp
    mclsLog.Save
    RaiseEvent AfterTransmit(strTmp, mblnTransmitting)
    
    Exit Sub
    
hErr:
    RaiseEvent AfterTransmit(Err.Description, mblnTransmitting)
    mclsLog.Add Err.Description
    mclsLog.Save
End Sub

Public Sub TransBase(ByVal strData As String)
'功能：基础数据传送
'参数：
'  strData：详见clsDataTransmit.Transmit
'  strData：发送的数据
'    格式：[接口编号]|[业务类别]|[业务数据]
'    接口编号：接口编号1[;接口编号...]
'    业务类别：
'        1-部门信息；
'        2-人员信息；
'        3-药品目录；
'        4-药品库存与库位；
'        5-发药窗口；
'    业务数据：
'        1-部门：部门性质1;部门性质2;…
'        2-人员：人员性质1;人员性质2;…
'        3-药品：剂型1;剂型2;…
'        4-库存：库房id1;库房id2;…
'        5-窗口：库房id1;库房id2;…

    Dim strINF As String, strClass As String, strDetail As String, strTrans As String
    Dim strMsg As String, strTmp As String
    Dim arrItems As Variant, arrINF As Variant
    Dim i As Integer, j As Integer
    
    arrItems = Split(strData, "|")
    
    mclsLog.Add "" & Now, Val("1-层")
    mclsLog.Add "开始基础数据传送", Val("1-层")
    
    On Error GoTo hErr
    
    '接口编号
    arrINF = Split(arrItems(0), ";")
    For i = LBound(arrINF) To UBound(arrINF)
    
        strINF = Trim(arrINF(i))
        If strINF = "" Then GoTo Continue
        
        '业务类型
        strClass = arrItems(1)
        If UBound(arrItems) > 1 Then
            strDetail = arrItems(2)
        Else
            strDetail = ""
        End If
        strTrans = strINF & "|" & strDetail
                
        Select Case Val(strClass)
        Case 1
            strTmp = "向“" & strINF & "”接口传送“部门信息”"
        Case 2
            strTmp = "向“" & strINF & "”接口传送“人员信息”"
        Case 3
            strTmp = "向“" & strINF & "”接口传送“药品目录”"
        Case 4
            strTmp = "向“" & strINF & "”接口传送“库存信息”"
        Case 5
            strTmp = "向“" & strINF & "”接口传送“发药窗口”"
        End Select
        
        '传送数据
        mclsLog.Add strTrans, Val("1-层"), Val("1-详细日志")
        
        If mobjMachine.Operation(mstrUser, Val(strClass), strTrans, strMsg) Then
            '正常日志输出
            strTmp = strTmp & "成功"
        Else
            '异常日志输出
            strTmp = strTmp & "失败"
        End If
        
        mclsLog.Add strTmp, Val("1-层")
        mclsLog.Save
        RaiseEvent AfterTransmit(strTmp, mblnTransmitting)
        
Continue:
    Next
    
    Erase arrINF
    Erase arrItems
    
    Exit Sub
    
hErr:
    mclsLog.Add Err.Number & ":" & Err.Description, Val("1-层")
    mclsLog.Save
End Sub


