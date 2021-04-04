VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsComPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mblnLogin As Boolean    '是否已经登录
Private mfrmPrint As frmMain    '调用打印部件要用的窗体
Private mobjPrint As Object     '打印部件

Public Function zlLogin(ByVal strUserName As String, ByVal strPwd As String, ByVal strSvr As String) As String
    '功能:      登录
    '参数:
    '       UserName    用户名
    '       strPwd      登录中联导航台的密码
    '       strSvr      服务器
    '返回
    '       OK 表示成功;其他则为提示信息
    Dim strErr As String
    
    mblnLogin = False
    zlLogin = OraDataOpen(strSvr, strUserName, TranPasswd(strPwd))
    
    If zlLogin = "OK" Then
        strErr = ""
        If ComInitComLib(strErr) Then
            mblnLogin = True
        Else
            If strErr = "" Then strErr = "初始化公共部件失败！"
            zlLogin = strErr
        End If
    End If
    
End Function
Public Function zlGetZyPrintList(ByVal strInPatientNo As String, ByVal strOutDate As String) As String
    '功能：  获取住院患者可打印报告
    '参数：  strInPatientNo    String  住院号
    '        strOutDate        strdate 出院日期，用于确定病人住院次数
    
    '返回：  strOut          string  OK|标本ID|标本号|姓名|性别|年龄|申请科室|
    '                                   申请人|申请时间|执行科室|核收日期|申请项目|打印次数;
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim dateS As Date, dateE As Date, strReturn As String
    On Error GoTo hErr
    ShowLog LOG_PRINTSVR, LOG_INFO, "获取住院报告清单", 100, "入参 " & strInPatientNo & "," & strOutDate
    If Not mblnLogin Then
        zlGetZyPrintList = "未登录!"
        Exit Function
    End If
    If strOutDate = "" Then
        zlGetZyPrintList = "出院日期为空!"
        Exit Function
    End If
    If Not IsDate(strOutDate) Then
        zlGetZyPrintList = "传入的出院日期参数不是日期!"
        Exit Function
    Else
        dateS = Format(CDate(strOutDate), "yyyy-MM-dd 00:00:00")
        dateE = Format(CDate(strOutDate), "yyyy-MM-dd 23:59:59")
    End If
    strSQL = "Select a.Id, a.标本序号, a.姓名, a.性别, a.年龄, b.名称 As 申请科室, a.申请人, a.申请时间, c.名称 As 执行科室, a.审核时间, a.核收时间, a.检验项目, a.打印次数 " & vbNewLine & _
            "From 检验标本记录 A, 部门表 B, 部门表 C, 病案主页 D" & vbNewLine & _
            "Where a.申请科室id = b.Id(+) And a.执行科室id = c.Id(+) And a.审核时间 Is Not Null And 病人来源 = 2 And" & vbNewLine & _
            "      a.主页id = d.主页id And a.病人id = d.病人id And d.住院号 = [1] And d.出院日期 Between [2] And [3]"

    Set rsTmp = ComOpenSQL(strSQL, "", strInPatientNo, dateS, dateE)
    strReturn = ""
    Do Until rsTmp.EOF
        strReturn = strReturn & "^" & rsTmp!ID & "|" & rsTmp!标本序号 & "|" & rsTmp!姓名 & "|" & rsTmp!性别 & "|" & rsTmp!年龄 & "|" & rsTmp!申请科室 & "|" & rsTmp!申请人
        strReturn = strReturn & "|" & Format(rsTmp!申请时间, "yyyy-MM-dd HH:mm:ss") & "|" & rsTmp!执行科室 & "|" & Format(rsTmp!核收时间, "yyyy-MM-dd HH:mm:ss") & "|" & rsTmp!检验项目
        strReturn = strReturn & "|" & Val("" & rsTmp!打印次数)
        rsTmp.MoveNext
    Loop
    If strReturn <> "" Then
        zlGetZyPrintList = "OK" & strReturn
    Else
        zlGetZyPrintList = "无住院号为" & strInPatientNo & ", 出院日期为" & strOutDate & "的报告！"
    End If
    
    ShowLog LOG_PRINTSVR, LOG_INFO, "获取住院报告清单", 100, "出参 " & zlGetZyPrintList
    Exit Function
hErr:
    strReturn = CStr(Erl()) & "行," & Err.Description
    zlGetZyPrintList = "zlGetZyPrintList " & strReturn
    ShowLog LOG_PRINTSVR, LOG_ERR, "获取住院报告清单", Err.Number, strReturn
    
End Function

Public Function zlGetPrintList(ByVal strOutPatientNO As String, Optional ByVal intDays As Integer = 7) As String
    '功能：  获取患者可打印报告
    '参数：  OutPatientNO    String  门诊号
    '        intDays         integer 天数，表示取最近几天内的报告，默认取7天内的报告
    '返回：  strOut          string  OK|标本ID|标本号|姓名|性别|年龄|申请科室|
    '                                   申请人|申请时间|执行科室|核收日期|申请项目;
    Dim strSQL As String, dNow As Date, dateS As Date, dateE As Date
    Dim strKeyName As String, rsTmp As ADODB.Recordset
    Dim strReturn As String
    On Error GoTo hErr
    
    ShowLog LOG_PRINTSVR, LOG_INFO, "获取报告清单", 100, "入参 " & strOutPatientNO & "," & intDays
    If Not mblnLogin Then
        zlGetPrintList = "未登录！"
        Exit Function
    End If
    If Not (intDays >= 1 And intDays <= 100) Then intDays = 7
    
    dNow = ComCurrDate
    
    dateS = Format(dNow - intDays, "yyyy-MM-dd")
    dateE = Format(dNow, "yyyy-MM-dd 23:59:59")
    
     
    strSQL = "Select a.Id, a.标本序号, a.姓名, a.性别, a.年龄, b.名称 As 申请科室, a.申请人, a.申请时间, c.名称 As 执行科室, a.审核时间, a.核收时间, a.检验项目" & vbNewLine & _
            "From 检验标本记录 A, 部门表 B, 部门表 C" & vbNewLine & _
            "Where a.申请科室id = b.Id(+) And a.执行科室id = c.Id(+) And a.审核时间 Is Not Null And Nvl(A.打印次数,0)<=0 And 病人来源<>2 And a.门诊号 = [1] And" & vbNewLine & _
            "      a.核收时间 Between [2] And [3]"

    '
    Set rsTmp = ComOpenSQL(strSQL, "", Val(strOutPatientNO), dateS, dateE)
    strReturn = ""
    Do Until rsTmp.EOF
        strReturn = strReturn & "^" & rsTmp!ID & "|" & rsTmp!标本序号 & "|" & rsTmp!姓名 & "|" & rsTmp!性别 & "|" & rsTmp!年龄 & "|" & rsTmp!申请科室 & "|" & rsTmp!申请人
        strReturn = strReturn & "|" & Format(rsTmp!申请时间, "yyyy-MM-dd HH:mm:ss") & "|" & rsTmp!执行科室 & "|" & Format(rsTmp!核收时间, "yyyy-MM-dd HH:mm:ss") & "|" & rsTmp!检验项目
        rsTmp.MoveNext
    Loop
    If strReturn <> "" Then
        zlGetPrintList = "OK" & strReturn
    Else
        zlGetPrintList = "最近" & intDays & "天内，没有门诊号为" & strOutPatientNO & "的可打印报告！"
    End If
    ShowLog LOG_PRINTSVR, LOG_INFO, "获取报告清单", 100, "出参 " & zlGetPrintList
    Exit Function
hErr:
    strReturn = CStr(Erl()) & "行," & Err.Description
    zlGetPrintList = "zlGetPrintList " & strReturn
    ShowLog LOG_PRINTSVR, LOG_ERR, "获取报告清单", Err.Number, strReturn
    
End Function

Public Function zlPrintReport(ByVal lngSampleID As Long) As String
    '功能:   打印患者的指定报告
    '参数：  SampleID        long       传入从zlGetPrintList中返回的标本ID
    '返回：  strOut          string     OK-打印成功;否则为失败的提示信息
    If mfrmPrint Is Nothing Then Set mfrmPrint = New frmMain
    If mobjPrint Is Nothing Then Set mobjPrint = CreateObject("zl9Report.clsReport")
    If Not mblnLogin Then
        zlPrintReport = "未登录！"
        Exit Function
    End If
    
    zlPrintReport = ReportPrint(lngSampleID, mfrmPrint, mobjPrint, True)

End Function

Public Function zlLoginOut() As String
    '功能:   退出中联系统
    '参数:   无
    '返回：  strOut      string  OK-成功;否则为失败的提示信息
    
    mblnLogin = False
    Call OraDataClose
    Call FreeImageObj
    
    If Not mfrmPrint Is Nothing Then Unload mfrmPrint
    zlLoginOut = "OK"
End Function

