Attribute VB_Name = "mdl成都"
Option Explicit
Public gcnSybase As New ADODB.Connection
Public g成都结算信息 As String

Public Function 医保设置_成都() As Boolean
'功能： 该方法用于供相关应用部件调用配置连接医保数据服务器的连接串
'返回：接口配置成功，返回true；否则，返回false
    Dim strConn As String
    
    If frmSet成都.ShowSet(TYPE_成都市) = False Then
        Exit Function
    End If
    
    strConn = GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("ConnectionString"), "dsn=cnnSyb;uID=face;pwd=facepass")
    '重新建立到医保服务器的公共连接
    If gcnSybase.State = adStateClosed Then
        On Error Resume Next
        gcnSybase.Open strConn
        If Err = 0 Then
            医保设置_成都 = True
        Else
            Err.Clear
        End If
    Else
        医保设置_成都 = True
    End If
End Function

Public Function 医保初始化_成都() As Boolean
'功能：传递应用部件已经建立的ORacle连接，同时根据配置信息建立与医保服务器的连接。
'返回：初始化成功，返回true；否则，返回false

    '建立到医保服务器的公共连接
    Dim strConn As String
    strConn = GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("ConnectionString"), "")
    Err = 0
    On Error Resume Next
    With gcnSybase
        If .State = 1 Then .Close
        .ConnectionString = strConn
        .Open
        If Err <> 0 Then
            MsgBox "不能建立到医保服务器的连接，无法执行医保交易", vbExclamation, gstrSysName
            Exit Function
        End If
    End With
    
    医保初始化_成都 = True
End Function

Public Function 身份标识_成都2(ByVal strCard As String, ByVal strPass As String, Optional bytType As Byte, Optional lng病人ID As Long) As String
'功能：识别指定人员是否为参保病人，返回病人的信息
'参数：strCard-刷卡得到；strPass-病人密码；bytType-识别类型，0-门诊，1-住院
'返回：空或信息串
'注意：1)主要利用接口的身份识别交易；
'      2)如果识别错误，在此函数内直接提示错误信息；
'      3)识别正确，而个人信息缺少某项，必须以空格填充；
'权限：部门表_ID,病人信息,保险帐户,zl_病人信息_Insert,zl_病人信息_Update,zl_保险帐户_insert,zl_帐户年度信息_Insert
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strInfo As String
    Dim str医保号 As String, str卡号 As String
    Dim strSerial As String, strSwapNo As String '交易顺序号
    Dim cur余额 As Currency
    Dim cur住院基数 As Currency, cur报销比例 As Currency, cur住院限额 As Currency
    
    If strCard = "" Then Exit Function
    
    '解析出医保号和卡号
    Call ExecuteZ015(strCard, str医保号, str卡号)
    If str医保号 = "" And str卡号 = "" Then
        MsgBox "刷卡解析失败，请重试！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '验证身份
    With rsTmp
        If .State = 1 Then .Close
        strSQL = "select 部门表_id.nextval||'1' from dual"
        .CursorLocation = adUseClient
    End With
    
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "成都医保")
    
    With rsTmp
        strSwapNo = .Fields(0).Value
        strSerial = getSerial(str医保号)
        
        'New:交易编号,客户机编号,交易顺序号,密码,操作员编号,就诊登记号,医保号,医院编码,交易时间,数据批号,支付类别,卡号
        strSQL = "z001('z001','" & UserInfo.站点 & "','" & strSwapNo & "','" & strPass & "','" & UserInfo.编号 & "'," & _
            "'" & strSerial & "','" & str医保号 & "','" & Trim(gstr医院编码) & "','" & DateStr & "','" & strSwapNo & "','" & IIf(bytType = 0, "11", "31") & "','" & str卡号 & "')"
        gcnSybase.Execute strSQL, , adCmdStoredProc
        
        strSQL = "select code from zjycl  where jysxh='" & strSwapNo & "' and jybh='z001'"
        If .State = 1 Then .Close
        .Open strSQL, gcnSybase, adOpenStatic, adLockReadOnly
        If Trim(.Fields(0).Value) <> "0000" Then
            MsgBox "交易""z001""出现错误""" & !CODE & """:" & vbCrLf & String(2, "　") & GetErrInfo(!CODE, TYPE_成都市) & String(2, vbTab), vbInformation, gstrSysName
            Exit Function
        Else
            strSQL = "select * from grjbxx where grbm='" & str医保号 & "'"
            If .State = 1 Then .Close
            .Open strSQL, gcnSybase, adOpenKeyset, adLockReadOnly
            If Not .EOF Then
                'New:0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码)
                strInfo = str卡号 & ";" & str医保号 & ";" & strPass & ";" & _
                        TrimStr(.Fields("xm").Value) & ";" & _
                        IIf(TrimStr(.Fields("xb").Value) = "1", "男", "女") & ";" & _
                        TrimStr(.Fields("csrq").Value) & ";" & _
                        TrimStr(.Fields("sfz").Value) & ";" & _
                        TrimStr(.Fields("dwmc").Value) & "(" & Trim(.Fields("dwbm").Value) & ")"
                
                cur余额 = IIf(IsNull(!grzhlnye), 0, !grzhlnye) + IIf(IsNull(!grzhbnye), 0, !grzhbnye)
                '200308z012
                If bytType <> 0 Then
                    cur住院基数 = IIf(IsNull(!zyjs), 0, !zyjs)
                    cur报销比例 = IIf(IsNull(!tcbxbl), 0, !tcbxbl)
                    cur住院限额 = IIf(IsNull(!zyxe), 0, !zyxe)
                End If
                
                lng病人ID = BuildPatiInfo(bytType, strInfo & ";;;;" & cur余额 & ";;;;;;;" & _
                    cur余额 & ";;;;;;" & cur住院基数 & ";" & cur报销比例 & ";" & cur住院限额, lng病人ID, TYPE_成都市)
                
                '返回格式:中间插入病人ID
                身份标识_成都2 = strInfo & ";" & lng病人ID & ";;;;" & cur余额 & ";;;;;;;" & cur余额 & ";;;;;"
            End If
        End If
    End With
End Function

Public Function 身份标识_成都(Optional bytType As Byte, Optional lng病人ID As Long) As String
'功能：识别指定人员是否为参保病人，返回病人的信息
'参数：bytType-识别类型，0-门诊，1-住院
'返回：空或信息串
'注意：1)主要利用接口的身份识别交易；
'      2)如果识别错误，在此函数内直接提示错误信息；
'      3)识别正确，而个人信息缺少某项，必须以空格填充；
    Dim frmIDentified As New frmIdentify成都
    Dim strPatiInfo As String, cur余额 As Currency
    Dim cur住院基数 As Currency, cur住院限额 As Currency, cur报销比例 As Currency
    
    frmIDentified.Tag = bytType
    frmIDentified.Show 1
    'New:0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码)
    strPatiInfo = frmIDentified.mstrPatiInfo
    cur余额 = frmIDentified.mcur余额
    cur住院基数 = frmIDentified.mcur住院基数
    cur报销比例 = frmIDentified.mcur报销比例
    cur住院限额 = frmIDentified.mcur住院限额
    Unload frmIDentified
    
    If strPatiInfo <> "" Then
        '建立病人档案信息，传入格式：
        '0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);8中心;9.顺序号;
        '10人员身份;11帐户余额;12当前状态;13病种ID;14在职(0,1);15退休证号;16年龄段;17灰度级
        '18帐户增加累计;19帐户支出累计;20进入统筹累计;21统筹报销累计;22住院次数累计;23就诊类型 (1、急诊门诊);
        '24本次起付线;25起付线累计;26基本统筹限额
        
        '200308z012
        lng病人ID = BuildPatiInfo(bytType, strPatiInfo & ";;;;" & cur余额 & ";;;;;;;" & _
            cur余额 & ";;;;;;" & cur住院基数 & ";" & cur报销比例 & ";" & cur住院限额, lng病人ID, TYPE_成都市)
        If lng病人ID = 0 Then Exit Function
        '返回格式:中间插入病人ID
        strPatiInfo = strPatiInfo & ";" & lng病人ID & ";;;;" & cur余额 & ";;;;;;;" & cur余额 & ";;;;;"
    End If
    身份标识_成都 = strPatiInfo
End Function

Public Function 个人余额_成都(strSelfNo As String, Optional bytYear As Byte) As Currency
'功能: 提取参保病人个人帐户余额
'参数: strSelfNO-病人个人编号,bytYear-余额类型,0-所有余额,1-本年余额,2-往年余额
'返回: 返回个人帐户余额的金额
    Dim rsTmp As New ADODB.Recordset
    
    On Error Resume Next
    With rsTmp
        gstrSQL = "Select grzhlnye,grzhbnye From grjbxx Where grbm='" & strSelfNo & "'"
        .CursorLocation = adUseClient
        .Open gstrSQL, gcnSybase, adOpenKeyset
        If .RecordCount > 0 Then
            Select Case bytYear
            Case 1
                个人余额_成都 = .Fields(1).Value
            Case 2
                个人余额_成都 = .Fields(0).Value
            Case Else
                个人余额_成都 = .Fields(0).Value + .Fields(1).Value
            End Select
        Else
            个人余额_成都 = 0
        End If
    End With
End Function

Public Function 门诊结算_成都(lng结帐ID As Long, lng病人ID As Long, str医保号 As String, str密码 As String, str卡号 As String, cur全自付 As Currency) As Boolean
'功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
'参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
'      cur支付金额   从个人帐户中支出的金额
'      str医保号     医保号
'返回：交易成功返回true；否则，返回false
'注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
'      2)理论上，由于我们保证了个人帐户结算金额不大于个人帐户余额，因此交易必然成功。但从安全角度考虑；
'        当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，
'        需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
    Dim strSerial As String, lngCount As Long, cur余额 As Currency
    Dim rsList As New ADODB.Recordset, rsTmp As New ADODB.Recordset
    Dim cur个帐支付 As Currency, cur发生费用 As Currency, cur首先自付 As Currency
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer, curDate As Date
    Dim cur本次起付线 As Currency, cur起付线累计 As Currency, cur基本统筹限额 As Currency
On Error GoTo ErrH
    strSerial = getSerial(str医保号)
    
    '此时所有收费细目必然有对应的医保编码
    gstrSQL = "Select * From 门诊费用记录 Where Nvl(附加标志,0)<>9 And 结帐ID=[1]"
    gstrSQL = "Select A.NO,A.登记时间,A.开单人 as 医生," & _
            "   A.数次*A.付数 as 数量,Round(A.结帐金额/(A.数次*A.付数),2) as 实际价格,A.结帐金额," & _
            "   D.项目编码 as 收费项目,B.名称 as 项目名称," & _
            "   decode(Instr(B.规格,'┆'),0,B.规格,substr(B.规格,1,Instr(B.规格,'┆')-1)) as 规格," & _
            "   decode(Instr(B.规格,'┆'),0,'',substr(B.规格,Instr(B.规格,'┆')+1)) as 产地," & _
            "   C.名称 as 科室名称" & _
            " From (" & gstrSQL & ") A,收费细目 B,部门表 C,保险支付项目 D " & _
            " Where A.收费细目ID=B.ID And A.开单部门ID=C.ID And A.收费细目ID=D.收费细目ID And D.险类=[2]" & _
            " Order by A.ID"
    Set rsList = zlDatabase.OpenSQLRecord(gstrSQL, "成都医保", CLng(lng结帐ID), TYPE_成都市)
    With rsList
        If .RecordCount = 0 Then
            Err.Raise 9000, gstrSysName, "没有填写收费记录。"
            Exit Function
        End If
        
        '插入费用明细(Z003)
        Dim strFeeKind As String
        lngCount = 0
        Do While Not .EOF
            lngCount = lngCount + 1
            gstrSQL = "Select sfdlmc From sfxmdl Where sfdlbm='" & Left(!收费项目, 3) & "'"
            With rsTmp
                If .State = 1 Then .Close
                .CursorLocation = adUseClient
                .Open gstrSQL, gcnSybase, adOpenKeyset
                strFeeKind = .Fields(0).Value
            End With
            gstrSQL = "insert into zfymx(jysxh,sfsj,pcno,grbm," & _
                    "   sfdlbm,sfxmbm,sl,sjjg," & _
                    "   cd,gg,yfyl,fyze,zfbl," & _
                    "   txbz,bpbz,qzfbf,ggzfbf,yxbxbf,fyshbz," & _
                    "   sfy,jbr,bz,sfdlmc,sfxmmc," & _
                    "   sjph,xh,yybm,ksbm,fylx," & _
                    "   tjdm,ysxm,ksmc,blh,zyh) " & _
                    " values ('" & lng结帐ID & "3',getdate(),'" & UserInfo.站点 & "','" & str医保号 & "'," & _
                    "   '" & Left(!收费项目, 3) & "','" & !收费项目 & "'," & !数量 & "," & !实际价格 & "," & _
                    "   '" & !产地 & "','" & !规格 & "',''," & !结帐金额 & ",0," & _
                    "   '','',0,0,0,''," & _
                    "   '" & UserInfo.姓名 & "','" & UserInfo.姓名 & "','','" & strFeeKind & "','" & !项目名称 & "'," & _
                    "   '" & lng结帐ID & "3','" & lngCount & "','" & Trim(gstr医院编码) & "','',''," & _
                    "   '','" & !医生 & "','" & !科室名称 & "','" & !NO & "','')"
            gcnSybase.Execute gstrSQL
            
            cur发生费用 = cur发生费用 + !结帐金额
            .MoveNext
        Loop
    End With
    
    'New:交易编号,客户机编号,交易顺序号,密码,操作员编号,就诊登记号,医保号,医院编码,交易时间,数据批号,支付类别,卡号
    gstrSQL = "z003('z003','" & UserInfo.站点 & "','" & lng结帐ID & "3','" & str密码 & "','" & UserInfo.编号 & "'," & _
        "'" & strSerial & "','" & str医保号 & "','" & Trim(gstr医院编码) & "','" & DateStr & "','" & lng结帐ID & "3','11','" & str卡号 & "')"
    gcnSybase.Execute gstrSQL, , adCmdStoredProc

    '检查是否正确(zjycl)
    With rsTmp
        gstrSQL = "select code from zjycl where jysxh='" & lng结帐ID & "3' And jybh='z003' order by jyend"
        If .State = 1 Then .Close
        .Open gstrSQL, gcnSybase, adOpenKeyset
        If Trim(.Fields(0).Value) <> "0000" Then
            Err.Raise 9000, gstrSysName, "交易""z003""出现错误""" & !CODE & """:" & vbCrLf & String(2, "　") & GetErrInfo(!CODE, TYPE_成都市) & String(2, vbTab)
            门诊结算_成都 = False
            Exit Function
        End If
    End With
    
    'New:交易编号,客户机编号,交易顺序号,密码,操作员编号,就诊登记号,医保号,医院编码,交易时间,数据批号,支付类别,卡号
    gstrSQL = "z008('z008','" & UserInfo.站点 & "','" & lng结帐ID & "8','" & str密码 & "','" & UserInfo.编号 & "'," & _
        "'" & strSerial & "','" & str医保号 & "','" & Trim(gstr医院编码) & "','" & DateStr & "','','11','" & str卡号 & "')"
    gcnSybase.Execute gstrSQL, , adCmdStoredProc
    
    With rsTmp
        '检查是否正确(zjycl)
        gstrSQL = "select code from zjycl where jysxh='" & lng结帐ID & "8' And jybh='z008' order by jyend"
        If .State = 1 Then .Close
        .Open gstrSQL, gcnSybase, adOpenKeyset
        If Trim(.Fields(0).Value) <> "0000" Then
            Err.Raise 9000, gstrSysName, "交易""z008""出现错误""" & !CODE & """:" & vbCrLf & String(2, "　") & GetErrInfo(!CODE, TYPE_成都市) & String(2, vbTab)
            门诊结算_成都 = False: Exit Function
        End If
        '---------------------------------------------------------------------------------------------
        '填写结算表
        curDate = zlDatabase.Currentdate
                
        cur余额 = 个人余额_成都(str医保号)
    End With
    
    '求个人帐户支付金额
    gstrSQL = "Select 冲预交 From 病人预交记录 Where 结算方式='个人帐户' And 记录性质 Not In (11,1) And 结帐ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "成都医保", lng结帐ID)
        
    With rsTmp
        If Not .EOF Then cur个帐支付 = IIf(IsNull(!冲预交), 0, !冲预交)
                
        '帐户年度信息
        Call Get帐户信息(TYPE_成都市, lng病人ID, Year(curDate), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计, cur本次起付线, cur起付线累计, cur基本统筹限额)
                        
        '200308z012:"本次起付线=住院基数","基本统筹限额=报销比例"
        gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & TYPE_成都市 & "," & Year(curDate) & "," & _
            cur余额 & "," & cur帐户支出累计 & "," & cur进入统筹累计 & "," & _
            cur统筹报销累计 & "," & int住院次数累计 & "," & cur本次起付线 & "," & cur起付线累计 & "," & cur基本统筹限额 & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "成都医保")
        
        '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
        gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & TYPE_成都市 & "," & lng病人ID & "," & _
            Year(curDate) & "," & cur余额 & "," & cur帐户支出累计 & "," & cur进入统筹累计 & "," & _
            cur统筹报销累计 & "," & int住院次数累计 & ",NULL,NULL,NULL," & cur发生费用 & "," & _
            cur全自付 & "," & cur首先自付 & ",NULL,NULL,NULL,NULL," & cur个帐支付 & ",NULL)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "成都医保")
        '---------------------------------------------------------------------------------------------
        '曾明春(2005-10-13)进行语音提示
        If gblnLED Then
           zl9LedVoice.Speak "#25 " & cur发生费用
           If cur个帐支付 < cur发生费用 Then
              zl9LedVoice.Speak "#27 " & cur发生费用 - cur个帐支付
           Else
              zl9LedVoice.Speak "#26 " & cur余额
           End If
        End If
    End With
    
    门诊结算_成都 = True
    Exit Function
ErrH:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 个人帐户转预交_成都(lng预交ID As Long, curMoney As Currency, rs预交记录 As ADODB.Recordset) As Boolean
'功能：将需要从个人帐户余额转入预交款的数据记录发送医保前置服务器确认；
'参数：lng预交ID-当前预交记录的ID，从预交记录中可以检索医保号和密码
'返回：交易成功返回true；否则，返回false
    Dim str医保号 As String, str密码 As String, strSerial As String, str卡号 As String
    Dim lng病人ID As Long, lng主页ID As Long, cur余额 As Currency, cur金额 As Currency
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer, cur本次起付线 As Currency
    Dim cur起付线累计 As Currency, cur基本统筹限额 As Currency
    Dim rsTmp As New ADODB.Recordset, curDate As Date
    Dim strDJZT As String
On Error GoTo ErrH
    With rs预交记录
        lng病人ID = rs预交记录!病人ID
        lng主页ID = IIf(IsNull(rs预交记录!主页ID), 0, rs预交记录!主页ID)
        str卡号 = TrimStr(IIf(IsNull(!卡号), "", !卡号))
        str医保号 = TrimStr(IIf(IsNull(!医保号), str卡号, !医保号))
        str密码 = TrimStr(IIf(IsNull(!密码), "", !密码))
        strSerial = getSerial(str医保号)
        
        cur金额 = !金额
        cur余额 = 个人余额_成都(str医保号, 1) '取本年余额,所有余额肯定大于下帐金额
    End With
    
    strDJZT = Trim(GetGrjbxx(str医保号, "djzt"))
    If strDJZT <> "120" Then
        Err.Raise 9000, gstrSysName, "该医保病人尚未入院,不能执行个人帐户转预交交易！"
        Exit Function
    End If
    
    '插入数据到个人帐户支付表
    gstrSQL = "insert into zgrzhzf(jysxh,pcno,grbm," & _
            "   yybm,zfsj,bnzhzf,lnzhzf,jbr,zfyy,bz)" & _
            " values ('" & lng预交ID & "A','" & UserInfo.站点 & "','" & str医保号 & "'," & _
            "   '" & Trim(gstr医院编码) & "',getdate()," & _
            IIf(cur余额 >= cur金额, cur金额, cur余额) & "," & _
            IIf(cur余额 >= cur金额, 0, cur金额 - cur余额) & "," & _
            "   '" & UserInfo.姓名 & "','','')"
    gcnSybase.Execute gstrSQL
    
    'New:交易编号,客户机编号,交易顺序号,密码,操作员编号,就诊登记号,医保号,医院编码,交易时间,数据批号,支付类别,卡号
    gstrSQL = "z010('z010','" & UserInfo.站点 & "','" & lng预交ID & "A','" & str密码 & "','" & UserInfo.编号 & "'," & _
        "'" & strSerial & "','" & str医保号 & "','" & Trim(gstr医院编码) & "','" & DateStr & "','" & lng预交ID & "A'," & _
        IIf(lng主页ID = 0, "'11'", "'31'") & ",'" & str卡号 & "')"
    gcnSybase.Execute gstrSQL, , adCmdStoredProc
    
    With rsTmp
        '检查是否正确(zjycl)
        gstrSQL = "Select code From zjycl Where jysxh='" & lng预交ID & "A' And jybh='z010'"
        If .State = 1 Then .Close
        .CursorLocation = adUseClient
        .Open gstrSQL, gcnSybase, adOpenKeyset
        If Trim(.Fields(0).Value) <> "0000" Then
            Err.Raise 9000, gstrSysName, "交易""z010""出现错误""" & !CODE & """:" & vbCrLf & String(2, "　") & GetErrInfo(!CODE, TYPE_成都市) & String(2, vbTab), vbInformation, gstrSysName
            个人帐户转预交_成都 = False: Exit Function
        End If
        '---------------------------------------------------------------------------------------------
        '填写结算表
        curDate = zlDatabase.Currentdate
        
        '帐户年度信息
        Call Get帐户信息(TYPE_成都市, lng病人ID, Year(curDate), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计, cur本次起付线, cur起付线累计, cur基本统筹限额)
        If int住院次数累计 = 0 Then int住院次数累计 = Get住院次数(lng病人ID)
        
        cur余额 = 个人余额_成都(str医保号) '取所有余额
        
        '200308z012:"本次起付线=住院基数","基本统筹限额=报销比例"
        gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & TYPE_成都市 & "," & Year(curDate) & "," & _
            cur余额 & "," & cur帐户支出累计 & "," & cur进入统筹累计 & "," & _
            cur统筹报销累计 & "," & int住院次数累计 & "," & cur本次起付线 & "," & cur起付线累计 & "," & cur基本统筹限额 & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "成都医保")
        
        '保险结算记录(因为"性质,记录ID"唯一,所以本次新预交ID肯定为插入)
        gstrSQL = "zl_保险结算记录_insert(3," & lng预交ID & "," & TYPE_成都市 & "," & lng病人ID & "," & _
            Year(curDate) & "," & cur余额 & "," & cur帐户支出累计 & "," & cur进入统筹累计 & "," & _
            cur统筹报销累计 & "," & int住院次数累计 & "," & cur本次起付线 & ",NULL," & cur本次起付线 & "," & _
            cur金额 & ",NULL,NULL,NULL,NULL,NULL,NULL," & cur金额 & ",NULL)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "成都医保")
        '---------------------------------------------------------------------------------------------
    End With
    个人帐户转预交_成都 = True
    Exit Function
ErrH:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 入院登记_成都(lng病人ID As Long, lng主页ID As Long, ByRef str医保号 As String) As Boolean
'功能：将入院登记信息发送医保前置服务器确认；
'参数：lng病人ID-病人ID；lng主页ID-主页ID
'返回：交易成功返回true；否则，返回false
    Dim jysxh As String, INDate As String, strInNote As String
    Dim strSelfNo As String, strSelfPwd As String, strSerial As String, strKH As String
    Dim rsTmp As New ADODB.Recordset, curDate As Date

    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer
    Dim cur住院基数 As Currency, cur报销比例 As Currency, cur住院限额 As Currency

    jysxh = zlDatabase.GetNextID("部门表") & "2"
    'New
    gstrSQL = "Select A.入院日期,A.入院病床,B.名称,D.住院号,SysDate as 经办时间,C.卡号,C.医保号,C.密码 " & _
            " From 病案主页 A,部门表 B,保险帐户 C,病人信息 D " & _
            " Where A.病人ID=D.病人ID And A.病人ID=[1] And A.主页ID=[2]" & _
            " And A.入院科室ID=B.ID And A.病人ID=C.病人ID And C.险类=[3]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "成都医保", lng病人ID, lng主页ID, TYPE_成都市)
    
    strKH = TrimStr(IIf(IsNull(rsTmp!卡号), "", rsTmp!卡号))
    strSelfNo = TrimStr(IIf(IsNull(rsTmp!医保号), strKH, rsTmp!医保号))
    strSelfPwd = TrimStr(IIf(IsNull(rsTmp!密码), "", rsTmp!密码))
    
    If strSelfNo = "" Then
        MsgBox "没有此病人或此病人不是医保病人！", vbExclamation, gstrSysName
        Exit Function
    End If
    '个人状态的修改
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_成都市 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "成都医保")
            
    strInNote = 获取入出院诊断(lng病人ID, lng主页ID)   '入院诊断
    strSerial = getSerial(strSelfNo)
    
    Dim mSqlTemp As String
    mSqlTemp = ""
    '提交住院登记表
    mSqlTemp = "insert into zzydj(jysxh,pcno,yybm,grbm,ryzd,rysj,ryks,rycw,ryjbr,blh,zyh,sftzb,tzbbxbl,bpbz,jbsj)" & _
            " values('" & jysxh & "','" & UserInfo.站点 & "','" & Trim(gstr医院编码) & "','" & strSelfNo & "'," & _
            "'" & strInNote & "','" & Format(rsTmp!入院日期, "yyyy-MM-dd hh:mm:ss") & "','" & rsTmp("名称") & "','" & rsTmp("入院病床") & "','" & _
            UserInfo.编号 & "','','" & rsTmp("住院号") & "','',0,'','" & Format(rsTmp!经办时间, "yyyy-MM-dd hh:mm:ss") & "')"
    gcnSybase.Execute mSqlTemp
    rsTmp.Close
    
    '提交交易登记表
    'New:交易编号,客户机编号,交易顺序号,密码,操作员编号,就诊登记号,医保号,医院编码,交易时间,数据批号,支付类别,卡号
    gstrSQL = "z002('z002','" & UserInfo.站点 & "','" & jysxh & "','" & strSelfPwd & "','" & UserInfo.编号 & "'," & _
        "'" & strSerial & "','" & strSelfNo & "','" & Trim(gstr医院编码) & "','" & DateStr & "','" & jysxh & "','31','" & strKH & "')"
    gcnSybase.Execute gstrSQL, , adCmdStoredProc
    
    '检查是否正确(zjycl)
    gstrSQL = "Select code From zjycl Where jysxh='" & jysxh & "' And jybh='z002'"
    If rsTmp.State = 1 Then rsTmp.Close
    rsTmp.Open gstrSQL, gcnSybase, adOpenStatic, adLockReadOnly
    If Trim(rsTmp("code").Value) <> "0000" Then
        MsgBox "交易""z002""出现错误""" & rsTmp!CODE & """:" & vbCrLf & String(2, "　") & GetErrInfo(rsTmp!CODE, TYPE_成都市) & String(2, vbTab), vbInformation, gstrSysName
        入院登记_成都 = False
        Exit Function
    End If
    
    '200308z012:删除取顺序号,病人不再使用固定顺序号
    
    '填写帐户年度信息
    curDate = zlDatabase.Currentdate
    Call Get帐户信息(TYPE_成都市, lng病人ID, Year(curDate), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
    If int住院次数累计 = 0 Then int住院次数累计 = Get住院次数(lng病人ID)
        
    '200308z012:保存住院基数和报销比例
    cur住院基数 = Val(GetGrjbxx(strSelfNo, "zyjs")) '保存到"本次起付线"
    cur报销比例 = Val(GetGrjbxx(strSelfNo, "tcbxbl")) '保存到"起付线累计"
    cur住院限额 = Val(GetGrjbxx(strSelfNo, "zyxe")) '保存到"基本统筹限额"
    
    '200308z012:"本次起付线=住院基数","基本统筹限额=报销比例"
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & TYPE_成都市 & "," & Year(curDate) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 & "," & int住院次数累计 & "," & cur住院基数 & "," & cur报销比例 & "," & cur住院限额 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "成都医保")
    
    入院登记_成都 = True
End Function

Public Function 出院登记_成都(lng病人ID As Long, lng主页ID As Long, rs病人 As ADODB.Recordset) As Boolean
'功能：将出院信息发送医保前置服务器确认；
'参数：lng病人ID-病人ID；lng主页ID-主页ID
'返回：交易成功返回true；否则，返回false
    Dim rsTmp As New ADODB.Recordset
    Dim jysxh As String, OutDate As String, strOutNote As String
    Dim strSelfNo As String, strSelfPwd As String, strSerial As String, strKH As String
    
    'New
    strKH = TrimStr(IIf(IsNull(rs病人!卡号), "", rs病人!卡号))
    strSelfNo = TrimStr(IIf(IsNull(rs病人!医保号), strKH, rs病人!医保号))
    strSelfPwd = TrimStr(IIf(IsNull(rs病人!密码), "", rs病人!密码))
    
    strSerial = getSerial(strSelfNo)
    jysxh = zlDatabase.GetNextID("部门表") & "B"
    
    '个人状态的修改
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_成都市 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "成都医保")
    
    '提交出院登记表
    gstrSQL = "Select A.出院日期,A.出院病床,SysDate as 经办时间,B.住院号,A.出院方式,C.名称" & _
        " From 病案主页 A,病人信息 B,部门表 C" & _
        " Where A.出院科室ID=C.ID And A.病人ID=B.病人ID And A.病人ID=[1] And A.主页ID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "成都医保", lng病人ID, lng主页ID)
    
    '获取出院诊断
    strOutNote = 获取入出院诊断(lng病人ID, lng主页ID, False)
    
    gstrSQL = "insert into zcydj(jysxh,pcno,grbm,yybm,cysj,cyzd,cycw,cyjbr,blh,zyh,jbsj,cyyy,cyks,zyzt) " & _
            "values('" & jysxh & "','" & UserInfo.站点 & "','" & strSelfNo & "','" & Trim(gstr医院编码) & "','" & _
            Format(rsTmp!出院日期, "yyyy-MM-dd hh:mm:ss") & "','" & strOutNote & "','" & Nvl(rsTmp!出院病床) & "','" & UserInfo.编号 & "'," & _
            "'','" & Nvl(rsTmp!住院号) & "','" & Format(rsTmp!经办时间, "yyyy-MM-dd hh:mm:ss") & "'," & _
            "'" & Decode(Nvl(rsTmp!出院方式), "死亡", "1", "转院", "2", "0") & "','" & rsTmp!名称 & "','')"
    gcnSybase.Execute gstrSQL
    
    '提交交易登记表
    'New:交易编号,客户机编号,交易顺序号,密码,操作员编号,就诊登记号,医保号,医院编码,交易时间,数据批号,支付类别,卡号
    gstrSQL = "z011('z011','" & UserInfo.站点 & "','" & jysxh & "','" & strSelfPwd & "','" & UserInfo.编号 & "'," & _
        "'" & strSerial & "','" & strSelfNo & "','" & Trim(gstr医院编码) & "','" & DateStr & "','" & jysxh & "','31','" & strKH & "')"
    gcnSybase.Execute gstrSQL, , adCmdStoredProc
    
    '检查是否正确(zjycl)
    gstrSQL = "Select code From zjycl Where jysxh='" & jysxh & "' And jybh='z011'"
    If rsTmp.State = 1 Then rsTmp.Close
    rsTmp.Open gstrSQL, gcnSybase, adOpenStatic, adLockReadOnly
    If Trim(rsTmp("code").Value) <> "0000" Then
        MsgBox "交易""z011""出现错误""" & rsTmp!CODE & """:" & vbCrLf & String(2, "　") & GetErrInfo(rsTmp!CODE, TYPE_成都市) & String(2, vbTab), vbInformation, gstrSysName
        出院登记_成都 = False
        Exit Function
    End If
    出院登记_成都 = True
End Function

Public Function 住院虚拟结算_成都(rsList As ADODB.Recordset, str医保号 As String, str密码 As String) As String
'功能：获取该病人指定结帐内容的可报销金额；
'参数：rsList-需要结算的费用明细记录集合；str医保号-医保号；str密码-病人密码；
'返回：可报销金额串:"报销方式;金额;是否允许修改|...."
'注意：1)该函数主要使用模拟结算交易，查询结果返回获取基金报销额；
    Dim str顺序号 As String, str大类 As String, str数据批号 As String
    Dim lng序号 As Integer, lng病人ID As Long, lng结帐ID As Long
    Dim strSerial As String, str卡号 As String
    Dim strSQL As String, str备注 As String
    Dim blnTran As Boolean, i As Long
    Dim rsTmp As ADODB.Recordset
    Dim rs大类 As ADODB.Recordset
    
    Dim cur总额 As Currency, cur限额 As Currency, cur价格 As Currency
    Dim cur公费 As Currency, cur全自费 As Currency
    Dim cur比例自付 As Currency, cur比例报销 As Currency, sng比例 As Single
    Dim cur床位超限自付 As Currency, cur床位限额报销 As Currency
    Dim cur血费超限自付 As Currency, cur血费限额报销 As Currency
    
    On Error GoTo ErrH
    
    rsList.Filter = "婴儿费=0"
    If rsList.RecordCount = 0 Then Exit Function
    
    g结算数据.病人ID = rsList!病人ID
    g结算数据.主页ID = rsList!主页ID
    lng病人ID = rsList!病人ID
    
    '获取病人的一些帐户信息
    strSQL = "Select * From 保险帐户 Where 病人ID=" & lng病人ID & " And 险类=" & TYPE_成都市
    Set rsTmp = New ADODB.Recordset
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSQL, gcnOracle, adOpenKeyset, adLockReadOnly
    If rsTmp.EOF Then Exit Function
    
    str卡号 = TrimStr(IIf(IsNull(rsTmp!卡号), "", rsTmp!卡号))
    str医保号 = TrimStr(IIf(IsNull(rsTmp!医保号), str卡号, rsTmp!医保号))
    str密码 = TrimStr(IIf(IsNull(rsTmp!密码), "", rsTmp!密码))
    strSerial = getSerial(str医保号)
    
    '本次Z003交易的顺序号和开始序号
    lng序号 = 1
    str顺序号 = zlDatabase.GetNextID("病人结帐记录")
    str数据批号 = "D" & Format(DateStr, "YYYY-MM-DD")

    
    '从SybaseFace库获取收费细目大类清单
    strSQL = "select * from sfxmdl"
    Set rs大类 = New ADODB.Recordset
    rs大类.CursorLocation = adUseClient
    rs大类.Open strSQL, gcnSybase, adOpenKeyset, adLockReadOnly
    
    '插入费用明细zfymx
    gcnOracle.BeginTrans: blnTran = True
    
    For i = 1 To rsList.RecordCount
        If rsList!主页ID > g结算数据.主页ID Then g结算数据.主页ID = rsList!主页ID
        
        '床位费因为是单次限额,如果打折,单价也要打折
        If Left(Nvl(rsList!保险编码, rsList!医保项目编码), 3) = "002" And Mid(Nvl(rsList!保险编码, rsList!医保项目编码), 8, 1) = "2" Then
            cur价格 = rsList!金额 / IIf(rsList!数量 = 0, 1, rsList!数量)
        Else
            cur价格 = rsList!价格
        End If

        '只上传未上传部分
        '-----------------------------------------------------------------------------
        If rsList!是否上传 = 0 Then
            g成都结算信息 = "正在上传费用明细，请稍侯：" & vbCrLf & _
                "第" & i & "条明细，共" & rsList.RecordCount & "条明细。"
            frm成都结算提示.Show 1
            
            '获取收费大类名称
            str大类 = ""
            rs大类.Filter = "sfdlbm='" & Left(Nvl(rsList!保险编码, rsList!医保项目编码), 3) & "'"
            If Not rs大类.EOF Then str大类 = Nvl(rs大类!sfdlmc)

            '插入zfymx,该明细可用于交易(z003)
            'sfsj要用当前时间,不然作废再传时会违反唯一约束
            With rsList
                str备注 = "预结上传:" & !NO & ",序号:" & !序号
                strSQL = _
                    "insert into zfymx(" & _
                    "jysxh,sfsj,pcno,grbm,sfdlbm,sfxmbm,sl,sjjg,cd,gg,yfyl,fyze,zfbl,txbz,bpbz,qzfbf,ggzfbf,yxbxbf," & _
                    "fyshbz,sfy,jbr,bz,sfdlmc,sfxmmc,sjph,xh,yybm,ksbm,fylx,tjdm,ysxm,ksmc,blh,zyh) values (" & _
                    "'" & str顺序号 & "3',getdate()," & _
                    "'" & UserInfo.站点 & "','" & str医保号 & "','" & Left(Nvl(!保险编码, !医保项目编码), 3) & "','" & Nvl(!保险编码, !医保项目编码) & "'," & _
                    Format(!数量, "0.00") & "," & Format(cur价格, "0.00") & ",'" & IIf(IsNull(!产地), "", !产地) & "'," & _
                    "'" & IIf(IsNull(!规格), "", !规格) & "',''," & Format(!金额, "0.00") & ",0,'','',0,0,0,''," & _
                    "'" & UserInfo.姓名 & "','" & UserInfo.姓名 & "','" & str备注 & "','" & str大类 & "','" & !收费名称 & "'," & _
                    "'" & str顺序号 & "3','" & lng序号 & "','" & Trim(gstr医院编码) & "','','',''," & _
                    "'" & IIf(IsNull(!医生), "", !医生) & "','" & !开单部门 & "','" & lng病人ID & "','" & lng病人ID & "')"
            End With
            gcnSybase.Execute strSQL

            '标记该费用已上传(暂未提交)
            strSQL = "ZL_病人费用记录_上传('" & rsList!NO & "'," & rsList!序号 & "," & rsList!记录性质 & "," & rsList!记录状态 & ")"
            gcnOracle.Execute strSQL, , adCmdStoredProc

            lng序号 = lng序号 + 1
            
        Else
            '更新保险编码
            If IsNull(rsList!保险编码) Then
                strSQL = "ZL_病人费用记录_上传('" & rsList!NO & "'," & rsList!序号 & "," & rsList!记录性质 & "," & rsList!记录状态 & ",'" & rsList!医保项目编码 & "')"
                gcnOracle.Execute strSQL, , adCmdStoredProc
            End If
        End If
        
        cur总额 = cur总额 + Format(rsList!金额, "0.00")

        rsList.MoveNext
        
    Next
    
    '曾明春:2005-06-25 打开提示窗口
    g成都结算信息 = "正在进行预结算，请稍侯!"
    frm成都结算提示.Show 1
    
    '提交费用明细
    '-----------------------------------------------------------------------------
    If lng序号 > 1 Then
        'New:交易编号,客户机编号,交易顺序号,密码,操作员编号,就诊登记号,医保号,医院编码,交易时间,数据批号,支付类别,卡号
        strSQL = "z003('z003','" & UserInfo.站点 & "','" & str顺序号 & "3','" & str密码 & "','" & UserInfo.编号 & "'," & _
            "'" & strSerial & "','" & str医保号 & "','" & Trim(gstr医院编码) & "','" & DateStr & "','" & str数据批号 & "','31','" & str卡号 & "')"
        gcnSybase.Execute strSQL, , adCmdStoredProc
    
        '检查是否正确(zjycl)
        strSQL = "Select code From zjycl where grbm='" & str医保号 & "' and jysxh='" & str顺序号 & "3' and jybh='z003' and zflb='31' order by jyend desc"
        Set rsTmp = New ADODB.Recordset
        rsTmp.CursorLocation = adUseClient
        rsTmp.Open strSQL, gcnSybase, adOpenKeyset, adLockReadOnly
        If rsTmp.EOF Then
            gcnOracle.RollbackTrans
            MsgBox "未发现交易处理结果。", vbInformation, gstrSysName
            Exit Function
        ElseIf Trim(rsTmp!CODE) <> "0000" Then
            gcnOracle.RollbackTrans
            MsgBox "交易""z003""出现错误""" & rsTmp!CODE & """:" & vbCrLf & String(2, "　") & GetErrInfo(rsTmp!CODE, TYPE_成都市) & String(2, vbTab), vbInformation, gstrSysName
            Exit Function
        End If
    End If
    gcnOracle.CommitTrans: blnTran = False
    
    '辅助结算
    '---------------------------------------------------------------------------------------------------------------
    '删除对应顺序号的zjycl,zfzjs,以避免重复
    strSQL = "delete from zjycl where grbm='" & str医保号 & "' and jysxh='" & str顺序号 & "' and jybh='z008'"
    gcnSybase.Execute strSQL
    strSQL = "delete from zfzjs where grbm='" & str医保号 & "' and jysxh='" & str顺序号 & "'"
    gcnSybase.Execute strSQL
    
    'New:交易编号,客户机编号,交易顺序号,密码,操作员编号,就诊登记号,医保号,医院编码,交易时间,数据批号,支付类别,卡号
    'jysxh要使用当前的结帐ID,以便执行结帐作废z013之前获取相应信息
    strSQL = "z007('z007','" & UserInfo.站点 & "','" & str顺序号 & "','" & str密码 & "','" & UserInfo.编号 & "'," & _
        "'" & strSerial & "','" & str医保号 & "','" & Trim(gstr医院编码) & "','" & DateStr & "','" & str数据批号 & "','31','" & str卡号 & "')"
    gcnSybase.Execute strSQL, , adCmdStoredProc
         
     '检查是否正确(zjycl)
    strSQL = "Select code From zjycl Where grbm='" & str医保号 & "' and jysxh='" & str顺序号 & "' And jybh='z007' and zflb='31' order by jyend desc"
    Set rsTmp = New ADODB.Recordset
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSQL, gcnSybase, adOpenKeyset, adLockReadOnly
    If rsTmp.EOF Then
        MsgBox "未发现交易处理结果。", vbInformation, gstrSysName
        Exit Function
    ElseIf Trim(rsTmp!CODE) <> "0000" Then
        MsgBox "交易""z007""出现错误""" & rsTmp!CODE & """:" & vbCrLf & String(2, "　") & GetErrInfo(rsTmp!CODE, TYPE_成都市) & String(2, vbTab), vbInformation, gstrSysName
        Exit Function
    End If

    '返回:进入统筹部分,统筹支付部分,个人帐户支付
    strSQL = "Select fyze,jrjsbf,tczhifbf,grzhzf From zfzjs where grbm='" & str医保号 & "' and jysxh='" & str顺序号 & "' order by jbsj desc"
    Set rsTmp = New ADODB.Recordset
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSQL, gcnSybase, adOpenKeyset, adLockReadOnly
    If Not rsTmp.EOF Then
        If rsTmp!fyze <> cur总额 Then
            MsgBox "医院系统中的费用总金额与已经上传到医保的费用总金额不一致。" & vbCrLf & _
                "医院总金额：" & cur总额 & "元" & vbCrLf & _
                "医保总金额：" & rsTmp("fyze") & "元" & vbCrLf & _
                "请与管理员或中联工程师联系！" & String(2, " "), vbInformation, gstrSysName
            Exit Function
        Else
            住院虚拟结算_成都 = "医保基金;" & rsTmp!tczhifbf & ";0"
        End If
    End If
    Exit Function
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    If blnTran Then gcnOracle.RollbackTrans
End Function

Public Function 住院结算_成都(lng结帐ID As Long, rs帐户 As ADODB.Recordset) As Boolean
'功能：将需要本次结帐的费用明细和结帐数据发送医保前置服务器确认；
'参数: lng结帐ID -病人结帐记录ID, 从预交记录中可以检索医保号和密码
'返回：交易成功返回true；否则，返回false
'注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
'      2)理论上，由于我们通过模拟结算提取了基金报销额，保证了医保基金结算金额的正确性，
'        因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；
'        如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。
'        这样才能保证数据的完全统一。
'      3)由于结帐之后，可能使用结帐作废交易，这时需要结帐时执行结算交易的交易号，因此我们需要同时结帐交易号。
'        (由于门诊收费作废时，已经不再和医保有关系，所以不需要保存结帐的交易号)
    Dim str医保号 As String, str密码 As String, str卡号 As String
    Dim strSerial As String, lng病人ID As Long
    Dim str大类 As String, strSQL As String, i As Long
    
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer, cur本次起付线 As Currency, curDate As Date
    Dim cur起付线累计 As Currency, cur基本统筹限额 As Currency
    
    Dim cur住院基数 As Currency, cur发生费用 As Currency, cur支付比例 As Double
    Dim cur进入统筹 As Currency, cur统筹支付 As Currency
    Dim cur首先自付 As Currency, cur全自付 As Currency
    
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo ErrH
    
    '获取的一些帐户信息
    lng病人ID = rs帐户!病人ID
    str卡号 = TrimStr(IIf(IsNull(rs帐户!卡号), "", rs帐户!卡号))
    str医保号 = TrimStr(IIf(IsNull(rs帐户!医保号), str卡号, rs帐户!医保号))
    str密码 = TrimStr(IIf(IsNull(rs帐户!密码), "", rs帐户!密码))
    strSerial = getSerial(str医保号)
    
    '曾明春(2005-08-05):在结算的时候再次检查是否存在未上传的记录，避免操作员在预结算后使用工具清除银海数据，直接结帐。
    strSQL = "select Nvl(Sum(实收金额),0) as 未上传金额 from 住院费用记录 " & _
             " where 病人ID=" & lng病人ID & " and 门诊标志=2 and nvl(是否上传,0)=0 and 附加标志<>9 and  nvl(婴儿费,0)=0 and " & _
             " 主页ID=(Select distinct max(主页ID) from 病案主页 where 病人ID=" & lng病人ID & ")"
    Set rsTmp = New ADODB.Recordset
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSQL, gcnOracle, adOpenKeyset
    
    If rsTmp!未上传金额 <> 0 Then
       Err.Raise 9000, gstrSysName, "病人还存在未上传费用，请重新预结算后再结帐。"
       住院结算_成都 = False
       Exit Function
    End If
    
    '辅助结算
    '---------------------------------------------------------------------------------------------------------------
    '删除对应顺序号的zjycl,zfzjs,以避免重复
    strSQL = "delete from zjycl where grbm='" & str医保号 & "' and jysxh='" & lng结帐ID & "8' and jybh='z008'"
    gcnSybase.Execute strSQL
    strSQL = "delete from zfzjs where grbm='" & str医保号 & "' and jysxh='" & lng结帐ID & "8'"
    gcnSybase.Execute strSQL
    
    'New:交易编号,客户机编号,交易顺序号,密码,操作员编号,就诊登记号,医保号,医院编码,交易时间,数据批号,支付类别,卡号
    'jysxh要使用当前的结帐ID,以便执行结帐作废z013之前获取相应信息
    
    strSQL = "z008('z008','" & UserInfo.站点 & "','" & lng结帐ID & "8','" & str密码 & "','" & UserInfo.编号 & "'," & _
        "'" & strSerial & "','" & str医保号 & "','" & Trim(gstr医院编码) & "','" & DateStr & "','','31','" & str卡号 & "')"
    gcnSybase.Execute strSQL, , adCmdStoredProc
         
     '检查是否正确(zjycl)
    strSQL = "Select code From zjycl Where grbm='" & str医保号 & "' and jysxh='" & lng结帐ID & "8' And jybh='z008'"
    Set rsTmp = New ADODB.Recordset
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSQL, gcnSybase, adOpenKeyset, adLockReadOnly
    If rsTmp.EOF Then
        Err.Raise 9000, gstrSysName, "未发现交易处理结果。"
        Exit Function
    ElseIf Trim(rsTmp!CODE) <> "0000" Then
        Err.Raise 9000, gstrSysName, "交易""z008""出现错误""" & rsTmp!CODE & """:" & vbCrLf & String(2, "　") & GetErrInfo(rsTmp!CODE, TYPE_成都市) & String(2, vbTab) & vbCrLf & _
               "请在处理后重新提取病人信息。"
        Exit Function
    End If
    
    '填写结算表
    '---------------------------------------------------------------------------------------------------------------
    curDate = zlDatabase.Currentdate

    '住院基数,费用总额,进入统筹部分,统筹支付部份,全自付部份
    strSQL = "select zyjs,fyze,yxbxbf,tczhifbf,qzfbf,tczifbl from zfzjs where jysxh='" & lng结帐ID & "8' and grbm='" & str医保号 & "'"
    Set rsTmp = New ADODB.Recordset
    rsTmp.Open strSQL, gcnSybase, adOpenKeyset, adLockReadOnly
    If rsTmp.EOF Then
        Err.Raise 9000, gstrSysName, "未返回辅助结算记录！"
        Exit Function
    End If
    
    cur支付比例 = IIf(IsNull(rsTmp!tczifbl), 0, rsTmp!tczifbl) * 100 '为了保留有足够的小数位数，在原有比例上乘以100
    cur住院基数 = rsTmp!zyjs
    cur发生费用 = rsTmp!fyze
    cur进入统筹 = rsTmp!yxbxbf
    cur统筹支付 = rsTmp!tczhifbf
    cur全自付 = rsTmp!qzfbf
    cur首先自付 = cur发生费用 - cur全自付 - cur进入统筹
    
    '比较结算结果与预结结果是否一致
    strSQL = "Select 冲预交 From 病人预交记录 Where 记录性质=2 And 结算方式='医保基金' And 结帐ID=" & lng结帐ID
    Set rsTmp = New ADODB.Recordset
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSQL, gcnOracle, adOpenKeyset
    If rsTmp.EOF Then
        If cur统筹支付 <> 0 Then
            Err.Raise 9000, gstrSysName, "未发现预结记录！"
            Exit Function
        End If
    ElseIf cur统筹支付 <> Nvl(rsTmp!冲预交, 0) Then
        MsgBox "统筹支付金额为:" & Format(cur统筹支付, "0.00") & " ,与预结算的结果不一致！"
        Exit Function
    End If
    
    '将本次结帐记录标记为已上传
    strSQL = "zl_病人结帐记录_上传(" & lng结帐ID & ")"
    gcnOracle.Execute strSQL, , adCmdStoredProc
    
    '帐户年度信息
    Call Get帐户信息(TYPE_成都市, lng病人ID, Year(curDate), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计, cur本次起付线, cur起付线累计, cur基本统筹限额)
    If int住院次数累计 = 0 Then int住院次数累计 = Get住院次数(lng病人ID)
            
    '200308z012:"本次起付线=住院基数","基本统筹限额=报销比例"
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & TYPE_成都市 & "," & Year(curDate) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 & "," & cur进入统筹累计 + cur进入统筹 & "," & _
        cur统筹报销累计 + cur统筹支付 & "," & int住院次数累计 & "," & cur本次起付线 & "," & cur起付线累计 & "," & cur基本统筹限额 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "成都医保")
    
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & TYPE_成都市 & "," & lng病人ID & "," & _
        Year(curDate) & "," & cur帐户增加累计 & "," & cur帐户支出累计 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 & "," & int住院次数累计 & "," & cur住院基数 & "," & cur支付比例 & "," & cur住院基数 & "," & _
        cur发生费用 & "," & cur全自付 & "," & cur首先自付 & "," & cur进入统筹 & "," & cur统筹支付 & "," & _
        "NULL,NULL,NULL,NULL)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "成都医保")
    
    '保险结算计算
    gstrSQL = "zl_保险结算计算_insert(" & lng结帐ID & ",0," & cur进入统筹 & "," & cur统筹支付 & ",NULL)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "成都医保")
    
    住院结算_成都 = True
    Exit Function
ErrH:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 住院结算冲销_成都(lng结帐ID As Long, rs帐户 As ADODB.Recordset) As Boolean
'----------------------------------------------------------------
'功能：将指定结帐涉及的结帐交易和费用明细从医保数据中删除；
'参数：lng结帐ID-需要作废的结帐单ID号；
'返回：交易成功返回true；否则，返回false
'注意：1)主要使用结帐恢复交易和费用删除交易；
'      2)有关原结算交易号，在病人结帐记录中根据结帐单ID查找；原费用明细传输交易的交易号，
'        在病人费用记录中根据结帐ID查找；
'      3)作废的结帐记录(记录性质=2)其交易号，填写本次结帐恢复交易的交易号；
'        因结帐作废而产生的费用记录的交易号号，填写为本次费用删除交易的交易号。
'----------------------------------------------------------------
    Dim str医保号 As String, str密码 As String, str卡号 As String
    Dim cur费用总额 As Currency, strSerial As String, lng病人ID As Long
    Dim str结算编号 As String, str顺序号 As String
    
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer, cur本次起付线 As Currency
    Dim cur起付线累计 As Currency, cur基本统筹限额 As Currency
    Dim curDate As Date, lng新ID As Long
    
    Dim cur住院基数 As Currency, cur发生费用 As Currency, cur支付比例 As Double
    Dim cur进入统筹 As Currency, cur统筹支付 As Currency
    Dim cur首先自付 As Currency, cur全自付 As Currency
    
    Dim rsTmp As ADODB.Recordset, strSQL As String
        
    On Error GoTo ErrH
        
    '病人信息
    lng病人ID = rs帐户!病人ID
    str卡号 = TrimStr(IIf(IsNull(rs帐户!卡号), "", rs帐户!卡号))
    str医保号 = TrimStr(IIf(IsNull(rs帐户!医保号), str卡号, rs帐户!医保号))
    str密码 = TrimStr(IIf(IsNull(rs帐户!密码), "", rs帐户!密码))
    strSerial = getSerial(str医保号)
    
    '原"费用总额,结算编号"
    strSQL = "select fyze,jsbh from zfzjs where jysxh='" & lng结帐ID & "8' and grbm='" & str医保号 & "'"
    Set rsTmp = New ADODB.Recordset
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSQL, gcnSybase, adOpenKeyset, adLockReadOnly
    If rsTmp.EOF Then
        Err.Raise 9000, gstrSysName, "结帐记录未找到！", vbInformation, gstrSysName
        Exit Function
    End If
    str结算编号 = rsTmp!jsbh
    cur费用总额 = IIf(IsNull(rsTmp!fyze), 0, rsTmp!fyze)
    
    '插入费用结算表
    strSQL = _
        "insert into zfyjs(jysxh,pcno,grbm,yybm,zyjs," & _
        " nspgz,fyze,qzfbf,ggzfbf,yxbxbf,jrjsbf,tczifbl," & _
        " tczhifbf,grzhzf,zfsm,sbjkc,jbr,sfy,jbsj,bz,jsbh)" & _
        " values('" & lng结帐ID & "D','" & UserInfo.站点 & "','" & _
        str医保号 & "','" & Trim(gstr医院编码) & "',0,0," & _
        cur费用总额 & ",0,0,0,0,0,0,0,'',0,'" & UserInfo.编号 & "'," & _
        "'" & UserInfo.编号 & "',getdate() ,'','" & str结算编号 & "')"
    gcnSybase.Execute strSQL
    
    '提交交易登记表
    'New:交易编号,客户机编号,交易顺序号,密码,操作员编号,就诊登记号,医保号,医院编码,交易时间,数据批号,支付类别,卡号
    str顺序号 = zlDatabase.GetNextID("病人结帐记录") & "D"
    strSQL = "z013('z013','" & UserInfo.站点 & "','" & str顺序号 & "','" & str密码 & "','" & UserInfo.编号 & "'," & _
        "'" & strSerial & "','" & str医保号 & "','" & Trim(gstr医院编码) & "','" & DateStr & "','" & str顺序号 & "','31','" & str卡号 & "')"
    gcnSybase.Execute strSQL, , adCmdStoredProc
    
    '检查是否正确(zjycl)
    strSQL = "Select code From zjycl Where jysxh='" & str顺序号 & "' And jybh='z013'"
    Set rsTmp = New ADODB.Recordset
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSQL, gcnSybase, adOpenKeyset
    If rsTmp.EOF Then
        Err.Raise 9000, gstrSysName, "未发现交易处理结果。", vbInformation, gstrSysName
        Exit Function
    ElseIf Trim(rsTmp!CODE) <> "0000" Then
        Err.Raise 9000, gstrSysName, "交易""z013""出现错误""" & rsTmp!CODE & """:" & vbCrLf & String(2, "　") & GetErrInfo(rsTmp!CODE, TYPE_成都市) & String(2, vbTab), vbInformation, gstrSysName
        Exit Function
    End If
    
    '----------------------------------------------------------------------------------
    '填写结算表
    curDate = zlDatabase.Currentdate
    '获取作废后的结帐ID
    strSQL = "Select A.ID From 病人结帐记录 A,病人结帐记录 B" & _
        " Where A.NO=B.NO And A.记录状态=2 And B.记录状态=3" & _
        " And B.ID=" & lng结帐ID
    Set rsTmp = New ADODB.Recordset
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSQL, gcnOracle, adOpenKeyset
    If rsTmp.EOF Then
        Err.Raise 9000, gstrSysName, "未发现作废的结算数据！", vbInformation, gstrSysName
        Exit Function
    End If
    lng新ID = rsTmp!ID
    
    '帐户年度信息
    Call Get帐户信息(TYPE_成都市, lng病人ID, Year(curDate), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计, cur本次起付线, cur起付线累计, cur基本统筹限额)
    If int住院次数累计 = 0 Then int住院次数累计 = Get住院次数(lng病人ID)
    
    strSQL = "Select * From 保险结算计算 Where Nvl(档次,0)=0 And 结帐ID=" & lng结帐ID
    Set rsTmp = New ADODB.Recordset
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSQL, gcnOracle, adOpenKeyset
    If Not rsTmp.EOF Then
        cur进入统筹 = IIf(IsNull(rsTmp!进入统筹金额), 0, rsTmp!进入统筹金额)
        cur统筹支付 = IIf(IsNull(rsTmp!统筹报销金额), 0, rsTmp!统筹报销金额)
    End If
    
    strSQL = "Select * From 保险结算记录 Where 性质=2 And 记录ID=" & lng结帐ID
    Set rsTmp = New ADODB.Recordset
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSQL, gcnOracle, adOpenKeyset
    If Not rsTmp.EOF Then
        cur支付比例 = IIf(IsNull(rsTmp!封顶线), 0, rsTmp!封顶线)
        cur住院基数 = IIf(IsNull(rsTmp!实际起付线), 0, rsTmp!实际起付线)
        cur发生费用 = IIf(IsNull(rsTmp!发生费用金额), 0, rsTmp!发生费用金额)
        If cur进入统筹 = 0 Then cur进入统筹 = IIf(IsNull(rsTmp!进入统筹金额), 0, rsTmp!进入统筹金额)
        If cur统筹支付 = 0 Then cur统筹支付 = IIf(IsNull(rsTmp!统筹报销金额), 0, rsTmp!统筹报销金额)
        cur首先自付 = IIf(IsNull(rsTmp!首先自付金额), 0, rsTmp!首先自付金额)
        cur全自付 = IIf(IsNull(rsTmp!全自付金额), 0, rsTmp!全自付金额)
    End If
    
    '插入新的作废记录
    '200308z012:"本次起付线=住院基数","基本统筹限额=报销比例"
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & TYPE_成都市 & "," & Year(curDate) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 & "," & cur进入统筹累计 - cur进入统筹 & "," & _
        cur统筹报销累计 - cur统筹支付 & "," & int住院次数累计 & "," & cur本次起付线 & "," & cur起付线累计 & "," & cur基本统筹限额 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "成都医保")
    
    '保险结算计算
    gstrSQL = "zl_保险结算计算_insert(" & lng新ID & ",0," & -1 * cur进入统筹 & "," & -1 * cur统筹支付 & ",NULL)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "成都医保")
    
    '保险结算记录
    gstrSQL = "zl_保险结算记录_insert(2," & lng新ID & "," & TYPE_成都市 & "," & lng病人ID & "," & Year(curDate) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 & "," & cur进入统筹累计 & "," & cur统筹报销累计 & "," & _
        int住院次数累计 & "," & cur住院基数 & "," & cur支付比例 & "," & cur住院基数 & "," & -1 * cur发生费用 & "," & _
        -1 * cur全自付 & "," & -1 * cur首先自付 & "," & -1 * cur进入统筹 & "," & -1 * cur统筹支付 & "," & _
        "NULL,NULL,NULL,NULL)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "成都医保")

    住院结算冲销_成都 = True
    Exit Function
ErrH:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function getSerial(strSelfNo As String) As String
'----------------------------------------------------------
'功能：获取病人顺序号
'----------------------------------------------------------
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset

    strSQL = "select sxh from grjbxx where grbm='" & strSelfNo & "'"
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSQL, gcnSybase, adOpenKeyset
    If Not rsTmp.EOF Then getSerial = rsTmp.Fields(0).Value
End Function

Public Function GetGrjbxx(strSelfNo As String, strField As String) As Variant
'功能：获取grjbxx中指定字段的值
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset

    strSQL = "select " & strField & " from grjbxx where grbm='" & strSelfNo & "'"
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSQL, gcnSybase, adOpenKeyset
    If Not rsTmp.EOF Then
        GetGrjbxx = IIf(IsNull(rsTmp.Fields(strField).Value), "", rsTmp.Fields(strField).Value)
    End If
End Function

Public Sub ExecuteZ015(ByVal strCard As String, ByRef str医保号 As String, ByRef str卡号 As String)
'功能：执行Z015交易
'参数：
'   入：strCard=刷卡的内容
'   出：str医保号=根据卡内容解析的医保号
'   出：str卡号=根据卡内容解析的卡号
'说明：适用于成都新接口
    Dim cmdSybase As New ADODB.Command
    
    On Error GoTo ErrH
    
    With cmdSybase
        Set .ActiveConnection = gcnSybase
        .Parameters.Append .CreateParameter("vid", adVarChar, adParamInput, 30, strCard)
        .Parameters.Append .CreateParameter("vgrbm", adVarChar, adParamOutput, 20)
        .Parameters.Append .CreateParameter("vkh", adVarChar, adParamOutput, 20)
        .CommandType = adCmdStoredProc
        .CommandText = "z015"
        .Execute
        str医保号 = TrimStr(IIf(IsNull(.Parameters("vgrbm").Value), "", .Parameters("vgrbm").Value))
        str卡号 = TrimStr(IIf(IsNull(.Parameters("vkh").Value), "", .Parameters("vkh").Value))
    End With
    Exit Sub
ErrH:
    MsgBox Err.Number & vbCrLf & vbTab & Err.Description, vbInformation, gstrSysName
End Sub

Public Function 挂号结算_成都(lng结帐ID As Long, lng病人ID As Long, str医保号 As String, str密码 As String, str卡号 As String) As Boolean
'功能：将挂号收费的明细和结算数据转发送医保前置服务器确认；
'参数：lng结帐ID=挂号记录的结帐ID；
'权限：保险帐户,病人费用记录,收费细目,部门表,保险支付项目,病人预交记录,帐户年度信息,zl_帐户年度信息_insert,zl_保险结算记录_insert
    Dim strSerial As String, lngCount As Long, cur余额 As Currency
    Dim rsList As New ADODB.Recordset, rsTmp As New ADODB.Recordset
    Dim cur个帐支付 As Currency, cur发生费用 As Currency
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer, curDate As Date
    Dim cur本次起付线 As Currency, cur起付线累计 As Currency, cur基本统筹限额 As Currency
    
    Dim strFeeKind As String
    
    strSerial = getSerial(str医保号)
    
    '此时所有收费细目必然有对应的医保编码
    gstrSQL = "Select A.NO,A.登记时间,A.开单人 as 医生," & _
            "   A.数次*A.付数 as 数量,Round(A.结帐金额/(A.数次*A.付数),2) as 实际价格,A.结帐金额," & _
            "   D.项目编码 as 收费项目,B.名称 as 项目名称," & _
            "   decode(Instr(B.规格,'┆'),0,B.规格,substr(B.规格,1,Instr(B.规格,'┆')-1)) as 规格," & _
            "   decode(Instr(B.规格,'┆'),0,'',substr(B.规格,Instr(B.规格,'┆')+1)) as 产地," & _
            "   C.名称 as 科室名称" & _
            " From (Select * From 门诊费用记录 Where 结帐ID=[1]) A,收费细目 B,部门表 C,保险支付项目 D " & _
            " Where A.收费细目ID=B.ID And A.开单部门ID=C.ID And A.收费细目ID=D.收费细目ID And D.险类=[2]" & _
            " Order by A.ID"
    Set rsList = zlDatabase.OpenSQLRecord(gstrSQL, "成都医保", lng结帐ID, TYPE_成都市)
        
    With rsList
        If .EOF Then
            MsgBox "没有填写挂号记录。", vbExclamation, gstrSysName
            Exit Function
        End If
        
        '插入费用明细(Z003)
        lngCount = 0
        Do While Not .EOF
            lngCount = lngCount + 1
            gstrSQL = "Select sfdlmc From sfxmdl Where sfdlbm='" & Left(!收费项目, 3) & "'"
            With rsTmp
                If .State = 1 Then .Close
                .CursorLocation = adUseClient
                .Open gstrSQL, gcnSybase, adOpenKeyset
                strFeeKind = .Fields(0).Value
            End With
            gstrSQL = "insert into zfymx(jysxh,sfsj,pcno,grbm," & _
                    "   sfdlbm,sfxmbm,sl,sjjg," & _
                    "   cd,gg,yfyl,fyze,zfbl," & _
                    "   txbz,bpbz,qzfbf,ggzfbf,yxbxbf,fyshbz," & _
                    "   sfy,jbr,bz,sfdlmc,sfxmmc," & _
                    "   sjph,xh,yybm,ksbm,fylx," & _
                    "   tjdm,ysxm,ksmc,blh,zyh) " & _
                    " values ('" & lng结帐ID & "3',getdate(),'" & UserInfo.站点 & "','" & str医保号 & "'," & _
                    "   '" & Left(!收费项目, 3) & "','" & !收费项目 & "'," & !数量 & "," & !实际价格 & "," & _
                    "   '" & !产地 & "','" & !规格 & "',''," & !结帐金额 & ",0," & _
                    "   '','',0,0,0,''," & _
                    "   '" & UserInfo.姓名 & "','" & UserInfo.姓名 & "','','" & strFeeKind & "','" & !项目名称 & "'," & _
                    "   '" & lng结帐ID & "3','" & lngCount & "','" & Trim(gstr医院编码) & "','',''," & _
                    "   '','" & !医生 & "','" & !科室名称 & "','" & !NO & "','')"
            gcnSybase.Execute gstrSQL
            
            cur发生费用 = cur发生费用 + !结帐金额
            .MoveNext
        Loop
    End With
    
    'New:交易编号,客户机编号,交易顺序号,密码,操作员编号,就诊登记号,医保号,医院编码,交易时间,数据批号,支付类别,卡号
    gstrSQL = "z003('z003','" & UserInfo.站点 & "','" & lng结帐ID & "3','" & str密码 & "','" & UserInfo.编号 & "'," & _
        "'" & strSerial & "','" & str医保号 & "','" & Trim(gstr医院编码) & "','" & DateStr & "','" & lng结帐ID & "3','11','" & str卡号 & "')"
    gcnSybase.Execute gstrSQL, , adCmdStoredProc

    '检查是否正确(zjycl)
    With rsTmp
        gstrSQL = "select code from zjycl where jysxh='" & lng结帐ID & "3' And jybh='z003'"
        If .State = 1 Then .Close
        .Open gstrSQL, gcnSybase, adOpenKeyset
        If Trim(.Fields(0).Value) <> "0000" Then
            MsgBox "交易""z003""出现错误""" & !CODE & """:" & vbCrLf & String(2, "　") & GetErrInfo(!CODE, TYPE_成都市) & String(2, vbTab), vbInformation, gstrSysName
            挂号结算_成都 = False
            Exit Function
        End If
    End With
    
    'New:交易编号,客户机编号,交易顺序号,密码,操作员编号,就诊登记号,医保号,医院编码,交易时间,数据批号,支付类别,卡号
    gstrSQL = "z008('z008','" & UserInfo.站点 & "','" & lng结帐ID & "8','" & str密码 & "','" & UserInfo.编号 & "'," & _
        "'" & strSerial & "','" & str医保号 & "','" & Trim(gstr医院编码) & "','" & DateStr & "','','11','" & str卡号 & "')"
    gcnSybase.Execute gstrSQL, , adCmdStoredProc
    
    With rsTmp
        '检查是否正确(zjycl)
        gstrSQL = "select code from zjycl where jysxh='" & lng结帐ID & "8' And jybh='z008'"
        If .State = 1 Then .Close
        .Open gstrSQL, gcnSybase, adOpenKeyset
        If Trim(.Fields(0).Value) <> "0000" Then
            MsgBox "交易""z008""出现错误""" & !CODE & """:" & vbCrLf & String(2, "　") & GetErrInfo(!CODE, TYPE_成都市) & String(2, vbTab), vbInformation, gstrSysName
            挂号结算_成都 = False: Exit Function
        End If
        '---------------------------------------------------------------------------------------------
        '填写结算表
        curDate = zlDatabase.Currentdate
                
        cur余额 = 个人余额_成都(str医保号)
        
        '求个人帐户支付金额
        gstrSQL = "Select 冲预交 From 病人预交记录 Where 结算方式='个人帐户' And 记录性质 Not In (11,1) And 结帐ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "成都医保", lng结帐ID)
        
        If Not .EOF Then cur个帐支付 = IIf(IsNull(!冲预交), 0, !冲预交)
                
        '帐户年度信息
        Call Get帐户信息(TYPE_成都市, lng病人ID, Year(curDate), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计, cur本次起付线, cur起付线累计, cur基本统筹限额)
                        
        '200308z012:"本次起付线=住院基数","基本统筹限额=报销比例"
        gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & TYPE_成都市 & "," & Year(curDate) & "," & _
            cur余额 & "," & cur帐户支出累计 & "," & cur进入统筹累计 & "," & cur统筹报销累计 & "," & _
            int住院次数累计 & "," & cur本次起付线 & "," & cur起付线累计 & "," & cur基本统筹限额 & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "成都医保")
        
        '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
        gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & TYPE_成都市 & "," & lng病人ID & "," & _
            Year(curDate) & "," & cur余额 & "," & cur帐户支出累计 & "," & cur进入统筹累计 & "," & _
            cur统筹报销累计 & "," & int住院次数累计 & ",NULL,NULL,NULL," & cur发生费用 & "," & _
            0 & "," & 0 & ",NULL,NULL,NULL,NULL," & cur个帐支付 & ",NULL)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "成都医保")
        '---------------------------------------------------------------------------------------------
    End With
    挂号结算_成都 = True
End Function

Public Function 记帐传输_成都(strNO As String, int性质 As Integer, int状态 As Integer, Optional lng病人ID As Long) As Boolean
'功能：将住院病人的记帐单据上传到医保前置服务器
'参数：lng病人ID=是否只上传单据中指定病人的费用
    Dim rsBill As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, lngPatiID As Long
    Dim lng序号 As Long, str大类 As String
    Dim str备注 As String
    Dim i As Long
    
    On Error GoTo ErrH
    
    '读取单据明细(医保号,顺序号,登记时间,项目编码,项目名称,产地,规格,数量,单价,金额,医生,开单科室)
    '单据中非该医保的费用不传,未设置医保编码的不传,无顺序号的不传,婴儿费不上传。按病人排序
    strSQL = _
        "Select Nvl(A.价格父号,序号) as 序号," & _
        " A.病人ID,F.医保号,F.顺序号,A.登记时间,Nvl(A.保险编码,D.项目编码) as 项目编码,B.名称 as 项目名称, " & _
        " Decode(Instr(B.规格,'┆'),0,B.规格,Substr(B.规格,1,Instr(B.规格,'┆')-1)) as 规格," & _
        " Decode(Instr(B.规格,'┆'),0,'',Substr(B.规格,Instr(B.规格,'┆')+1)) as 产地," & _
        " Avg(Nvl(A.付数,1)*A.数次) as 数量,Sum(A.标准单价) as 单价,Sum(A.实收金额) as 金额," & _
        " A.开单人 as 医生,C.名称 as 开单科室" & _
        " From 住院费用记录 A,收费细目 B,部门表 C,保险支付项目 D,病案主页 E,保险帐户 F" & _
        " Where A.收费细目ID=B.ID And A.开单部门ID=C.ID And A.收费细目ID=D.收费细目ID" & _
        " And A.病人ID=E.病人ID And A.主页ID=E.主页ID And A.病人ID=F.病人ID" & _
        " And F.顺序号 is Not NULL And Nvl(A.婴儿费,0)=0 And A.记录状态<>0 And Nvl(A.是否上传,0)=0" & _
        " And D.险类=" & TYPE_成都市 & " And E.险类=" & TYPE_成都市 & " And F.险类=" & TYPE_成都市 & _
        " And A.NO='" & strNO & "' And A.记录性质=" & int性质 & " And A.记录状态=" & int状态 & _
        IIf(lng病人ID = 0, "", " And A.病人ID=" & lng病人ID) & _
        " Group by Nvl(A.价格父号,序号),A.病人ID,F.医保号,F.顺序号," & _
        " A.登记时间,Nvl(A.保险编码,D.项目编码),B.名称,B.规格,A.开单人,C.名称" & _
        " Order by 病人ID,序号"
    rsBill.CursorLocation = adUseClient
    rsBill.Open strSQL, gcnOracle, adOpenKeyset, adLockReadOnly
    
    For i = 1 To rsBill.RecordCount
        '记帐单中有多个病人,要分别处理
        If rsBill!病人ID <> lngPatiID Then
            lngPatiID = rsBill!病人ID
            
            '获取该病人已上传的最大序号
            strSQL = "select max(convert(integer,xh)) as xh from zfymx where jysxh='" & rsBill!顺序号 & "7' and grbm='" & rsBill!医保号 & "'"
            If rsTmp.State = 1 Then rsTmp.Close
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open strSQL, gcnSybase, adOpenKeyset, adLockReadOnly
            lng序号 = 1
            If Not rsTmp.EOF Then lng序号 = IIf(IsNull(rsTmp!xh), 0, rsTmp!xh) + 1
        End If
        
        '获取收费大类名称
        strSQL = "select sfdlmc from sfxmdl where sfdlbm='" & Left(rsBill!项目编码, 3) & "'"
        If rsTmp.State = 1 Then rsTmp.Close
        rsTmp.Open strSQL, gcnSybase, adOpenKeyset, adLockReadOnly
        str大类 = ""
        If Not rsTmp.EOF Then str大类 = rsTmp!sfdlmc
        
        '插入zfymx,该明细可用于虚拟结算(z007)
        With rsBill
            If int状态 = 1 Then
                str备注 = "记帐:" & strNO & ",序号:" & !序号
            Else
                str备注 = "销帐:" & strNO & ",序号:" & !序号
            End If
            strSQL = _
                "insert into zfymx(" & _
                "jysxh,sfsj,pcno,grbm,sfdlbm,sfxmbm,sl,sjjg,cd,gg,yfyl,fyze,zfbl,txbz,bpbz,qzfbf,ggzfbf,yxbxbf," & _
                "fyshbz,sfy,jbr,bz,sfdlmc,sfxmmc,sjph,xh,yybm,ksbm,fylx,tjdm,ysxm,ksmc,blh,zyh) values (" & _
                "'" & !顺序号 & "7','" & Format(!登记时间, "yyyy-MM-dd hh:mm:ss") & "'," & _
                "'" & UserInfo.站点 & "','" & !医保号 & "','" & Left(!项目编码, 3) & "','" & !项目编码 & "'," & _
                Format(!数量, "0.00") & "," & Format(!单价, "0.00") & ",'" & IIf(IsNull(!产地), "", !产地) & "'," & _
                "'" & IIf(IsNull(!规格), "", !规格) & "',''," & Format(!金额, "0.00") & ",0,'','',0,0,0,''," & _
                "'" & UserInfo.姓名 & "','" & UserInfo.姓名 & "','" & str备注 & "','" & str大类 & "','" & !项目名称 & "'," & _
                "'" & !顺序号 & "7','" & lng序号 & "','" & Trim(gstr医院编码) & "','','',''," & _
                "'" & IIf(IsNull(!医生), "", !医生) & "','" & !开单科室 & "','" & !病人ID & "','" & !病人ID & "')"
        End With
        gcnSybase.Execute strSQL
        
        '标记已上传
        strSQL = "ZL_病人费用记录_上传('" & strNO & "'," & rsBill!序号 & "," & int性质 & "," & int状态 & ")"
        gcnOracle.Execute strSQL, , adCmdStoredProc
        
        lng序号 = lng序号 + 1
        
        rsBill.MoveNext
    Next
    
    记帐传输_成都 = True
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
End Function
