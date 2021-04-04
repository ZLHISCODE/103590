Attribute VB_Name = "mdlBillRelate"
Option Explicit
Public Function zlGetBillChargeExistInsure(ByVal lng结帐ID As Long, _
    Optional lng病人ID As Long, Optional bln急诊 As Boolean) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取收费(或退费)记录中的指定医保险类
    '入参:lng结帐ID-结帐ID
    '出参:lng病人ID-返回病人ID
    '     bln急诊-是否急诊
    '返回:如果存在则返回单据当时的险类
    '编制:刘兴洪
    '日期:2014-06-18 16:22:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String

    On Error GoTo errHandle
    lng病人ID = 0:  bln急诊 = False
    strSQL = "" & _
        "    Select B.记录ID,B.险类,B.病人ID,A.是否急诊  " & _
        "    From 保险结算记录 B,门诊费用记录 A " & _
        "    Where B.性质=1  And B.记录ID=[1]    " & _
        "          And B.记录ID=A.结帐ID And A.序号=1 and Rownum <2"

    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", lng结帐ID)
    If Not rsTemp.EOF Then
        lng病人ID = NVL(rsTemp!病人ID, 0)
        lng结帐ID = NVL(rsTemp!记录ID, 0)
        bln急诊 = NVL(rsTemp!是否急诊, 0) = 1
        zlGetBillChargeExistInsure = NVL(rsTemp!险类, 0)
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlIsCheckExiseSingularity(ByVal lng结算序号 As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查一次收费是否存在异常的作废单据
    '入参:lng结算序号-指定的结算序号
    '出参:
    '返回:存在返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-06-18 17:08:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errHandle
    strSQL = "" & _
    "   Select  1  " & _
    "   From  门诊费用记录 A,病人预交记录 B, 门诊费用记录 C  " & _
    "   Where c.记录性质 = a.记录性质 And c.No = a.No And A.结帐ID=B.结帐ID And Mod(a.记录性质,10)=1 And c.记录状态=2 " & _
    "        And b.结算序号=[1] And Rownum <2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取指定单据一次收费是否存在已经作废的单据", lng结算序号)
    zlIsCheckExiseSingularity = rsTemp.RecordCount <> 0
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlIsCheckExistErrBill(ByVal lng结算序号 As Long, Optional ByVal bln补充结算 As Boolean, _
    Optional ByVal strNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查一次收费是否存在异常单据
    '入参:lng结算序号-指定的结算序号
    '出参:
    '返回:存在返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-06-18 17:08:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errHandle
    If bln补充结算 Then
        strSQL = "" & _
        "   Select  1  " & _
        "   From  费用补充记录 A" & _
        "   Where a.记录性质=1 And Nvl(a.费用状态,0)=1 And a.结算序号=[1] And Rownum <2"
    Else
        If strNos <> "" Then
            strSQL = "" & _
            "   Select /*+cardinality(j,10) */ 1" & vbNewLine & _
            "   From 门诊费用记录 A, Table(f_Str2list([2])) J" & vbNewLine & _
            "   Where Mod(A.记录性质, 10) = 1 And a.No = j.Column_Value And Nvl(a.费用状态,0)=1 And Rownum < 2"
        Else
            strSQL = "" & _
            "   Select  1  " & _
            "   From  门诊费用记录 A,病人预交记录 B" & _
            "   Where Mod(a.记录性质,10)=1 And A.结帐ID=B.结帐ID And Nvl(a.费用状态,0)=1 And b.结算序号=[1] And Rownum <2"
        End If
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取指定单据一次收费是否存在已经作废的单据", lng结算序号, strNos)
    zlIsCheckExistErrBill = rsTemp.RecordCount <> 0
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlFromIDGetChargeBalance(ByVal bytType As Byte, _
    ByVal strValue As String, Optional blnHistory As Boolean, _
    Optional ByRef blnDel As Boolean, Optional ByVal bln含异常 As Boolean, _
    Optional ByVal byt记录性质 As Byte = 1) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据结算ID获取收费结算信息
    '入参:bytType-查找类型:0-根据结帐ID查找;1-根据结算序号查找,2-根据单据号来获取结算方式
    '     strValue-要查找的值(为0时,结帐ID,为1时,结算序号,2时为一次收费所涉及的所有单据)
    '     blnDel-退费结算:true-查退费结算;false-非退费结算
    '     bln含异常-是否包含异常结算，根据单据号来获取结算数据时有效
    '     byt记录性质-2-根据单据号来获取结算方式时传入，区分挂号/收费
    '返回:收费结算的相关信息集
    '       字段:类型,记录性质,结算方式,摘要,卡类别ID,卡类别名称,自制卡,结算卡序号,结算号码,卡号,交易流水号, 交易说明,结算序号,校对标志,医保,消费卡id
    '            是否密文,是否全退,是否退现,冲预交
    '       类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
    '编制:刘兴洪
    '日期:2014-06-24 16:37:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strTable As String, strWhere As String
    Dim strTable1 As String
    
    On Error GoTo errHandle
    strTable = IIf(blnHistory, "H", "") & "病人预交记录"
    Select Case bytType
    Case 0  '0-根据结帐ID查找
        strWhere = " And  A.结帐ID= [1]"
    Case 1  ';1-根据结算序号查找
        strWhere = "  And A.结算序号= [1]"
    Case 2 '根据单据号来获取结算数据
        strTable1 = "Select distinct 结帐ID  " & _
            "    From 门诊费用记录 M " & _
            "    Where M.NO in (Select Column_value From Table(f_str2List([2])))  " & _
            "          And Mod(M.记录性质,10)=[3]" & IIf(bln含异常, "", " And Nvl(M.费用状态,0)<>1")
        strTable1 = ",(" & strTable1 & ") Q1"
        If blnHistory Then strTable1 = Replace(strTable1, "门诊费用记录", "H门诊费用记录")
        strWhere = " And A.结帐ID=Q1.结帐ID"
    End Select

    If blnDel Then
        '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
        strSQL = "" & _
            "   Select  A.ID,decode(A.记录状态,2,A.结帐ID,NULL) as 结帐ID," & _
            "        Case when Mod(A.记录性质,10)=1 then 1  " & _
            "             when B.名称 is not null then  2 " & _
            "             when nvl(A.卡类别ID,0)<>0  then  3 " & _
            "             when J.结算方式 is not null   then  4 " & _
            "             else 0 end as 类型, " & _
            "        Mod(A.记录性质,10) as 记录性质,A.结算方式,A.冲预交," & _
            "        decode(A.记录状态,2,A.摘要,NULL) as 摘要,decode(A.记录状态,2,1,0) as 退费," & _
            "        A.卡类别ID,A.结算卡序号, " & _
            "        decode(A.记录状态,2,A.结算号码,NULL) as 结算号码,decode(A.记录状态,2,A.卡号,NULL) as 卡号, " & _
            "        decode(A.记录状态,2,A.交易流水号,NULL) as 交易流水号,nvl(C.是否自制,0) as 自制卡, " & _
            "        nvl(C.是否退现,0) as 是否退现,nvl(C.是否全退,0) as 是否全退, " & _
            "        Decode(C.卡号密文,NULL,0,1) as  是否密文,Nvl(C.是否转帐及代扣,0) as 是否转帐及代扣," & _
            "        Nvl(C.是否退款验卡,0) as 是否退款验卡," & _
            "        C.名称 as 卡类别名称,decode(A.记录状态,2,A.交易说明,NULL) as 交易说明,A.结算序号,decode(A.记录状态,2,A.校对标志,0) as 校对标志, " & _
            "        decode(B.名称,Null,0,1) as 医保,0 as 消费卡id,nvl(q.性质,1) as 结算性质" & _
            "   From " & strTable & " A ,医疗卡类别 C,一卡通目录 J,结算方式 q," & _
            "        (Select 名称 From 结算方式 where 性质 in (3,4)) B " & strTable1 & _
            "   Where A.结算方式=J.结算方式(+) And A.卡类别ID=C.ID(+) " & _
            "         And A.结算方式=B.名称(+) and A.结算方式=q.名称(+) " & _
            "         And (a.记录性质 In (1, 11) Or Nvl(a.结算卡序号, 0) = 0) " & strWhere
            
        strSQL = strSQL & " Union ALL " & _
            "   Select A.ID,decode(A.记录状态,2,A.结帐ID,NULL) as 结帐ID,5 as  类型,Mod(A.记录性质,10) as 记录性质,A.结算方式,-1*nvl(b.应收金额,0) as 冲预交,A.摘要," & _
            "        decode(A.记录状态,2,1,0) as 退费,A.卡类别ID,A.结算卡序号," & _
            "        decode(A.记录状态,2,A.结算号码,NULL) as 结算号码,decode(A.记录状态,2,B.卡号,NULL) as 卡号, " & _
            "        decode(A.记录状态,2,B.交易流水号,NULL) as 交易流水号,nvl(M.自制卡,0) as 自制卡, " & _
            "        nvl( M.是否退现,0) as 是否退现,nvl(M.是否全退,0) as 是否全退, " & _
            "        nvl(M.是否密文,0) as  是否密文, 0 as 是否转帐及代扣,0 as 是否退款验卡," & _
            "        M.名称 as 卡类别名称,A.交易说明,A.结算序号,A.校对标志,0 as 医保,B.消费卡id,nvl(q.性质,1) as 结算性质" & _
            "   From  " & strTable & " A ,病人卡结算记录 B, " & _
            "        消费卡类别目录 M ,结算方式 q " & strTable1 & _
            "   Where  a.Id = b.结算id And a.结算卡序号 = m.编号  " & _
            "         and Mod(A.记录性质,10)<>1 and A.结算方式=q.名称(+) " & strWhere

        strSQL = "" & _
            "   Select /*+ Rule */ max(结帐id) as 结帐id,类型,max(退费) as 退费,记录性质,结算方式,Max(摘要) as 摘要,卡类别ID,卡类别名称,max(自制卡) as 自制卡,结算卡序号, " & _
            "         max(结算号码) as 结算号码,max(卡号) as 卡号,max(交易流水号) as 交易流水号, max(交易说明) as 交易说明, " & _
            "         结算序号,max(校对标志) as 校对标志,医保,消费卡id,结算性质,max(是否转帐及代扣) as 是否转帐及代扣," & _
            "         Max(是否退款验卡) as 是否退款验卡," & _
            "         max(是否密文) as 是否密文,max(是否全退) as 是否全退,max(是否退现) as 是否退现 , nvl(sum(冲预交),0) as 冲预交" & _
            "   From (" & strSQL & ") " & _
            "   Group by 类型, 记录性质,结算方式,卡类别ID,卡类别名称,结算卡序号,结算序号,医保,消费卡id,结算性质 having  sum(冲预交) <>0"
        Set zlFromIDGetChargeBalance = zlDatabase.OpenSQLRecord(strSQL, "获取收费结算方式", Val(strValue), strValue, byt记录性质)
        Exit Function
    End If
    
    '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
    strSQL = "" & _
        "   Select /*+ Rule */ A.ID,A.结帐ID," & _
        "        Case when Mod(A.记录性质,10)=1 then 1  " & _
        "             when B.名称 is not null then  2 " & _
        "             when nvl(A.卡类别ID,0)<>0  then  3 " & _
        "             when J.结算方式 is not null   then  4 " & _
        "             else 0 end as 类型, " & _
        "        Mod(A.记录性质,10) as 记录性质,A.结算方式,A.冲预交," & _
        "        A.摘要,decode(A.记录状态,2,1,0) as 退费," & _
        "        A.卡类别ID,A.结算卡序号, " & _
        "        A.结算号码,A.卡号,A.交易流水号,nvl(C.是否自制,0) as 自制卡, " & _
        "        nvl(C.是否退现,0) as 是否退现,nvl(C.是否全退,0) as 是否全退, " & _
        "        Decode(C.卡号密文,NULL,0,1) as  是否密文,Nvl(C.是否转帐及代扣,0) as 是否转帐及代扣," & _
        "        Nvl(C.是否退款验卡,0) as 是否退款验卡," & _
        "        C.名称 as 卡类别名称,A.交易说明,A.结算序号,A.校对标志, " & _
        "        decode(B.名称,Null,0,1) as 医保,0 as 消费卡id,nvl(q.性质,1) as 结算性质" & _
        "   From " & strTable & " A ,医疗卡类别 C,一卡通目录 J,结算方式 q," & _
        "        (Select 名称 From 结算方式 where 性质 in (3,4)) B " & strTable1 & _
        "   Where A.结算方式=J.结算方式(+) And A.卡类别ID=C.ID(+) " & _
        "         And A.结算方式=B.名称(+) and A.结算方式=q.名称(+) " & _
        "         And (a.记录性质 In (1, 11) Or Nvl(a.结算卡序号, 0) = 0) " & strWhere

    strSQL = strSQL & " Union ALL " & _
        "   Select /*+ Rule */ A.ID,A.结帐ID,5 as  类型,Mod(A.记录性质,10) as 记录性质,A.结算方式,-1*nvl(b.应收金额,0) as 冲预交,A.摘要," & _
        "        decode(A.记录状态,2,1,0) as 退费,A.卡类别ID,A.结算卡序号," & _
        "        A.结算号码,B.卡号,B.交易流水号,nvl( M.自制卡,0) as 自制卡, " & _
        "        nvl( M.是否退现,0) as 是否退现,nvl(M.是否全退,0) as 是否全退, " & _
        "        nvl(M.是否密文,0) as  是否密文,0 as 是否转帐及代扣,0 as 是否退款验卡," & _
        "        M.名称 as 卡类别名称,A.交易说明,A.结算序号,A.校对标志,0 as 医保,B.消费卡id,nvl(q.性质,1) as 结算性质" & _
        "   From  " & strTable & " A ,病人卡结算记录 B, " & _
        "        消费卡类别目录 M ,结算方式 q " & strTable1 & _
        "   Where  a.Id = b.结算id And a.结算卡序号 = m.编号  " & _
        "         and Mod(A.记录性质,10)<>1 and A.结算方式=q.名称(+) " & strWhere
    gstrSQL = "" & _
        "   Select  结帐ID,类型,退费,记录性质,结算方式,摘要,卡类别ID,卡类别名称,自制卡,结算卡序号," & _
        "           结算号码,卡号,交易流水号, 交易说明,结算序号,校对标志,医保,消费卡id,结算性质," & _
        "           max(是否转帐及代扣) as 是否转帐及代扣,max(是否密文) as 是否密文,max(是否全退) as 是否全退," & _
        "           max(是否退款验卡) as 是否退款验卡," & _
        "           max(是否退现) as 是否退现 , nvl(sum(冲预交),0) as 冲预交" & _
        "   From (" & gstrSQL & ") " & _
        "   Group by 结帐ID,类型,退费,记录性质,结算方式,摘要,卡类别ID,卡类别名称,自制卡,结算卡序号,结算号码,卡号,交易流水号, 交易说明,结算序号,校对标志,医保,消费卡id,结算性质"
    Set zlFromIDGetChargeBalance = zlDatabase.OpenSQLRecord(strSQL, "获取收费结算方式", Val(strValue), strValue, byt记录性质)
    
    Exit Function
errHandle:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGet结帐ID(ByVal lng结算序号 As Long, _
    ByRef strNos As String, Optional ByRef intInusre As Integer, _
    Optional ByVal blnNoMove As Boolean, _
    Optional ByRef lng冲销ID As Long, _
    Optional ByVal bln补结算 As Boolean) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取收费(或退费)记录中的指结帐ID
    '入参:lng结算序号-结算序号
    '     blnNoMove-是否转移到历史数据
    '     bln补结算-是否补充结算
    '出参:strNOs-返回涉及的单据号
    '     intInusre-医保序号
    '     lng冲销ID-冲销ID
    '返回:返回指定的结帐ID
    '编制:刘兴洪
    '日期:2014-06-18 16:22:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String, str结帐ID As String
    
    On Error GoTo errHandle
    strNos = ""
    If bln补结算 Then
        '79142,冉俊明,2014-11-3,补结算退费预结算失败产生的退费异常单据无法获取险类
        strSQL = "" & _
            "   Select Distinct a.结算id As 结帐id, a.No, b.险类," & _
            "          Decode(a.记录状态, 2, a.结算id, 0) As 冲销id" & _
            "   From 费用补充记录 A," & _
            "        (Select Distinct s.No, t.险类" & _
            "          From 费用补充记录 S, 保险结算记录 T" & _
            "          Where s.结算id = t.记录id And t.性质(+) = 1 And s.结算序号 = [1]) B" & _
            "   Where a.No = b.No(+) And a.结算序号 = [1]" & _
            "   Order By NO"
    Else
        strSQL = "Select Distinct a.结帐id, a.No, b.险类, Decode(a.记录状态, 2, a.结帐id, 0) As 冲销id" & vbNewLine & _
                " From 门诊费用记录 A," & vbNewLine & _
                "      (Select Distinct s.结帐id, t.险类" & vbNewLine & _
                "        From 病人预交记录 S, 保险结算记录 T" & vbNewLine & _
                "        Where s.结帐id = t.记录id(+) And s.结算序号 = [1] And t.性质(+) = 1) B" & vbNewLine & _
                " Where a.结帐id = b.结帐id And Mod(a.记录性质, 10) = 1" & vbNewLine & _
                " Order By NO"
    End If
    If blnNoMove Then
        strSQL = Replace(strSQL, "病人预交记录", "H病人预交记录")
        strSQL = Replace(strSQL, "门诊费用记录", "H门诊费用记录")
        strSQL = Replace(strSQL, "保险结算记录", "H保险结算记录")
        If bln补结算 Then
            strSQL = Replace(strSQL, "费用补充记录", "H费用补充记录")
        End If
    End If

    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "根据结算序号获取对应的结帐ID", lng结算序号)
    If rsTemp.EOF Then Exit Function
    
    lng冲销ID = 0
    With rsTemp
        strNos = "": str结帐ID = ""
        Do While Not .EOF
            If InStr(str结帐ID & ",", "," & !结帐ID & ",") = 0 Then
                str结帐ID = str结帐ID & "," & Val(NVL(!结帐ID))
            End If
            If InStr(strNos & ",", "," & !NO & ",") = 0 Then
                strNos = strNos & "," & NVL(!NO)
            End If
            If Val(NVL(rsTemp!冲销ID)) <> 0 And lng冲销ID = 0 Then
                lng冲销ID = Val(NVL(rsTemp!冲销ID))
            End If
            If intInusre = 0 Then intInusre = Val(NVL(!险类))
            .MoveNext
        Loop
    End With
    If strNos <> "" Then strNos = Mid(strNos, 2)
    If str结帐ID <> "" Then str结帐ID = Mid(str结帐ID, 2)
    zlGet结帐ID = str结帐ID
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlGetAdviceFromID(ByVal str医嘱IDs As String) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据医嘱ID,获取对应的医嘱内容
    '入参:str医嘱IDs-医嘱ID(多个用逗号分离)
    '出参:
    '返回:成功,返回医嘱数据集(医嘱ID,医嘱内容)
    '编制:刘兴洪
    '日期:2014-06-27 11:52:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    
    On Error GoTo errHandle
    strSQL = "" & _
    "   Select ID as 医嘱ID,医嘱内容 " & _
    "   From 病人医嘱记录 " & _
    "   Where ID in (Select Column_value From Table(f_num2List([1])))"
    Set zlGetAdviceFromID = zlDatabase.OpenSQLRecord(strSQL, "获取医嘱内容", str医嘱IDs)
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetFromNOToLastBalanceID(ByVal strNos As String, _
    Optional ByVal blnNOMoved As Boolean, _
    Optional ByVal bln历史表同步查 As Boolean = False, _
    Optional lng结算序号 As Long, Optional bln补结算 As Boolean = False) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据一张收费单据的NO，返回最后一次有效的结帐的ID
    '入参:blnNoMoved是否在后备表中，查询单据之前的判断需要用这个参数
    '     bln历史表同步查-是否连接历史表一起查询
    '     bln补结算-是否补充结算
    '出参:lng结算序号-返回最后一次有效的结帐序号
    '返回:结帐ID
    '编制:刘兴洪
    '日期:2014-06-17 17:06:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String, strSQL1 As String
    Dim i As Long
    
    On Error GoTo errHandle:
    '87975
    strSQL = "With c_单据 As (Select Column_Value As NO From Table(f_Str2list([1])))" & vbNewLine & _
            " Select Max(a.结帐id) As 结帐id" & vbNewLine & _
            " From 门诊费用记录 A, c_单据 M" & vbNewLine & _
            " Where a.No = m.No" & vbNewLine & _
            "       And a.登记时间 + 0 =" & vbNewLine & _
            "           (Select Max(m.登记时间)" & vbNewLine & _
            "            From 门诊费用记录 M, c_单据 J" & vbNewLine & _
            "            Where m.No = j.No And Mod(m.记录性质, 10) = 1 And m.记录状态 In (1, 3) And Nvl(m.费用状态, 0) <> 1)" & vbNewLine & _
            "            And Mod(a.记录性质, 10) = 1 And a.记录状态 In (1, 3) And Nvl(a.费用状态, 0) <> 1"

    If bln补结算 Then
        strSQL = Replace(strSQL, "门诊费用记录", "费用补充记录")
        strSQL = Replace(strSQL, "Max(a.结帐id)", "Max(a.结算id)")
    End If

    strSQL = "" & _
            "   Select /*+ Rule */ A.结帐ID,B.结算序号 " & _
            "   From (" & strSQL & ") A,病人预交记录 B " & _
            "   Where A.结帐ID=B.结帐ID(+) And Rownum<2"

    If Not blnNOMoved And bln历史表同步查 Then
        strSQL1 = Replace(strSQL, "门诊费用记录", "H门诊费用记录")
        strSQL1 = Replace(strSQL, "费用补充记录", "H费用补充记录")
        strSQL1 = Replace(strSQL, "病人预交记录", "H病人预交记录")
        strSQL = strSQL & " Union ALL " & strSQL1
    ElseIf blnNOMoved Then
        strSQL = Replace(strSQL, "门诊费用记录", "H门诊费用记录")
        strSQL1 = Replace(strSQL, "费用补充记录", "H费用补充记录")
        strSQL = Replace(strSQL, "病人预交记录", "H病人预交记录")
    End If

    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "根据单据获取最后一次正常结帐的结帐ID", strNos)

    If rsTemp.EOF Then Exit Function

    lng结算序号 = Val(NVL(rsTemp!结算序号))
    zlGetFromNOToLastBalanceID = Val(NVL(rsTemp!结帐ID))
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlInvoiceGetNOs(ByVal strInvioceNo As String, Optional cllInvoiceNoInfor As Collection, _
    Optional blnNOMoved As Boolean, Optional bln补结算 As Boolean = False) As String
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据发票号,获取对应的单据号
    '入参:strInvioceNo-发票号
    '     blnNOMoved-是否在历史表空间
    '     bln补结算-是否医保补充结算
    '出参:cllInvoiceNoInfor-array(No,序号)
    '返回:成功返回传入的发票所涉及的单据号
    '编制:刘兴洪
    '日期:2013-04-12 15:59:32
    '问题:25187
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strNos As String
    Dim strSQL1 As String, strSQL As String

    On Error GoTo errHandle
    Set cllInvoiceNoInfor = New Collection
    If gTy_Module_Para.byt票据分配规则 <> 0 And bln补结算 = False Then
        strSQL = "" & _
            "   Select  /*+ RULE */  A.NO,Max(A.序号) as 序号,Max(C.结算序号) as 结算序号" & _
            "   From 票据打印明细 A,门诊费用记录 B,病人预交记录 C" & _
            "   Where A.票号=[1] and 票种=1 and A.是否回收<>1" & _
            "         And A.No=B.NO And Mod(B.记录性质,10)=1  And nvl(B.记录状态,0)<>2 And B.结帐ID=C.结帐ID" & _
            "   Group by A.NO"
        If blnNOMoved Then
            strSQL = Replace(strSQL, "票据打印明细", "H票据打印明细")
            strSQL = Replace(strSQL, "门诊费用记录", "H门诊费用记录")
            strSQL = Replace(strSQL, "病人预交记录", "H病人预交记录")
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取对应单据的发票号", strInvioceNo)

        strNos = ""
        With rsTemp
            Do While Not .EOF
                strNos = strNos & "," & NVL(!NO)
                cllInvoiceNoInfor.Add Array(NVL(!NO), NVL(!序号))
                .MoveNext
            Loop
            If strNos <> "" Then
                zlInvoiceGetNOs = Mid(strNos, 2)
                Exit Function
            End If
        End With
    End If
    
    strSQL = "" & _
        "   Select  Distinct NO  " & _
        "   From 票据打印内容 A, " & _
        "           (   Select Max(M.打印ID) as 打印ID " & _
        "               From  票据使用明细 M   " & _
        "               Where M.票种=1 And M.性质=1 And M.号码=[1]  " & _
        "               Group by M.号码" & _
        "               )  Q" & _
        "   Where A.数据性质=1  And ID=Q.打印ID " & _
        "   Order by NO"
    If blnNOMoved Then
        strSQL = Replace(strSQL, "票据打印明细", "H票据打印明细")
        strSQL = Replace(strSQL, "票据打印内容", "H票据打印内容")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取对应单据的发票号", strInvioceNo)

    With rsTemp
        Do While Not .EOF
            strNos = strNos & "," & NVL(!NO)
            .MoveNext
        Loop
        If strNos <> "" Then
            zlInvoiceGetNOs = Mid(strNos, 2)
        End If
    End With
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetChargeInsure(ByVal lng结帐ID As Long, ByRef lng病人ID As Long, _
    Optional ByVal blnNOMoved As Boolean) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取收费的医保号
    '入参:lng结帐ID-结帐ID
    '     blnNOMoved-是否数据转移
    '出参:lng病人ID-病人ID
    '返回:险类
    '编制:刘兴洪
    '日期:2014-07-02 14:30:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String

    On Error GoTo errHandle
    lng病人ID = 0
    strSQL = "" & _
        "    Select B.记录ID,B.险类,B.病人ID,A.是否急诊  " & _
        "    From 门诊费用记录 A,保险结算记录 B " & _
        "    Where A.结帐ID=[1] And  mod(A.记录性质,10)=1 " & _
        "         And B.性质=1 And A.结帐ID=B.记录ID and Rownum<2 "
    If blnNOMoved Then
        strSQL = Replace(strSQL, "门诊费用记录", "H门诊费用记录")
        strSQL = Replace(strSQL, "保险结算记录", "H保险结算记录")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "根据结帐ID获取指定的医保险类", lng结帐ID)
    If rsTemp.EOF Then Exit Function
    lng病人ID = NVL(rsTemp!病人ID, 0)
    zlGetChargeInsure = NVL(rsTemp!险类, 0)
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlMakeClinicPreSwapData(ByVal strStartFact As String, _
    ByVal lng结帐ID As Long, ByRef strNos As String, Optional bln补结算 As Boolean) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据单据对象内容创建一个记录信息(以售价单位)
    '入参:strStartFact-开始发票号
    '     lng结帐ID-重新收费结帐IDs
    '出参:strNos-返回本次结算的Nos
    '返回:医保相关数据的数据集(单据序号(1--n),病人ID,收费细目ID,数量,单价,实收金额,统筹金额,保险支付大类ID,是否医保)
    '编制:刘兴洪
    '日期:2014-07-07 11:24:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNo As String
    Dim i As Integer, j As Integer, intStartPage As Integer, intPages As Integer
    Dim p As Integer, strSQL As String
    Dim dbl单价 As Double, cur实收 As Currency, cur统筹 As Currency
    Dim rsTmp As New ADODB.Recordset, rsNo As ADODB.Recordset
    Dim strTable  As String, strWhere As String
    
    Err = 0: On Error GoTo Errhand:
    With rsTmp.Fields
        .Append "单据序号", adBigInt, 50, adFldIsNullable
        .Append "费别", adVarChar, 50, adFldIsNullable
        .Append "NO", adVarChar, 8, adFldIsNullable
        .Append "序号", adBigInt, , adFldIsNullable '问题:42961
        .Append "实际票号", adVarChar, 20, adFldIsNullable
        .Append "结算时间", adDBTimeStamp, , adFldIsNullable
        .Append "病人ID", adBigInt, , adFldIsNullable
        .Append "收费类别", adVarChar, 2, adFldIsNullable
        .Append "收据费目", adVarChar, 20, adFldIsNullable
        .Append "计算单位", adVarChar, 50, adFldIsNullable
        '79420,李南春,2014/11/10:调整记录集字段大小
        .Append "开单人", adVarChar, 100, adFldIsNullable
        .Append "收费细目ID", adBigInt, , adFldIsNullable
        .Append "数量", adDouble, , adFldIsNullable
        .Append "单价", adDouble, , adFldIsNullable
        .Append "实收金额", adCurrency, , adFldIsNullable
        .Append "统筹金额", adCurrency, , adFldIsNullable
        .Append "保险支付大类ID", adBigInt, , adFldIsNullable
        .Append "是否医保", adBigInt, , adFldIsNullable
        .Append "保险编码", adVarChar, 50, adFldIsNullable
        .Append "摘要", adVarChar, 2000, adFldIsNullable
        .Append "是否急诊", adBigInt, , adFldIsNullable
        .Append "开单部门ID", adBigInt, , adFldIsNullable
        .Append "执行部门ID", adBigInt, , adFldIsNullable
    End With
    rsTmp.CursorLocation = adUseClient
    rsTmp.LockType = adLockOptimistic
    rsTmp.CursorType = adOpenStatic
    rsTmp.Open
    
    strTable = ""
    strWhere = " And A.结帐ID=[1]"
    If bln补结算 Then
       strTable = ",(Select distinct 收费结帐ID From 费用补充记录 Where 结算ID=[1]) B"
       strWhere = " And A.结帐ID=b.收费结帐ID"
    End If

    strSQL = "Select A.NO,Nvl( A.价格父号, A.序号) as 序号,To_char(max(A.登记时间),'YYYY-MM-DD HH24:MI:SS') as 结算时间," & _
            "       A.病人ID,A.费别,A.收费类别,A.收据费目,A.计算单位,A.开单人," & _
            "       A.收费细目ID,A.保险大类ID As 保险支付大类ID,Nvl(A.保险项目否,0) As 是否医保,A.保险编码," & _
            "       Avg(Nvl(A.付数,0)*A.数次) As 数量,Avg(A.标准单价) As 单价," & _
            "       Sum(A.实收金额) As 实收金额,Sum(A.统筹金额) As 统筹金额,max(A.摘要) as 摘要," & _
            "       nvl(A.加班标志,0) as 是否急诊,A.开单部门ID,A.执行部门ID,A.结帐ID" & _
            " From 门诊费用记录 A" & strTable & _
            " Where Mod(A.记录性质,10)=1 " & strWhere & _
            " Group By A.NO, Nvl(A.价格父号, A.序号),A.病人id, A.费别, A.收费类别, A.收据费目, A.计算单位, A.开单人, A.收费细目id, A.保险大类id, Nvl(A.保险项目否, 0), A.保险编码, A.摘要, Nvl(A.加班标志, 0)," & _
            "       A.开单部门id, A.执行部门id,A.结帐ID"
    
    strSQL = "Select '" & strStartFact & "' as 实际票号,A.NO,A.序号,max(A.结算时间) as 结算时间," & _
            "       A.病人ID,A.费别,A.收费类别,A.收据费目,A.计算单位,A.开单人," & _
            "       A.收费细目ID,A.保险支付大类ID,A.是否医保,A.保险编码," & _
            "       sum(A.数量) as 数量,max(A.单价) As 单价, Sum(A.实收金额) As 实收金额, " & _
            "       Sum(A.统筹金额) As 统筹金额,max(A.摘要) as 摘要," & _
            "       Max(A.是否急诊) as 是否急诊,max(A.开单部门ID) as 开单部门ID,max(A.执行部门ID ) as 执行部门ID " & _
            " From (" & strSQL & ") A" & _
            " Group By A.NO,A.序号,A.病人id, A.费别, A.收费类别, A.收据费目, A.计算单位, A.开单人, A.收费细目id, A.保险支付大类ID, " & _
            "       A.是否医保, A.保险编码" & _
            " Order by NO,序号 "

    Set rsNo = zlDatabase.OpenSQLRecord(strSQL, "获取重新收费数据-医保", lng结帐ID)
    If rsNo.RecordCount <> 0 Then rsNo.MoveFirst
    With rsNo
        p = 1: strNos = ""
        Do While Not rsNo.EOF
            rsTmp.AddNew
            rsTmp!单据序号 = p
            rsTmp!费别 = !费别
            rsTmp!NO = NVL(!NO)    '仅提取划价单时才有值
            rsTmp!序号 = Val(NVL(!序号))    '仅提取划价单时才有值
            rsTmp!实际票号 = NVL(!实际票号)
            rsTmp!结算时间 = !结算时间
            rsTmp!病人ID = Val(NVL(!病人ID))
            rsTmp!收费类别 = NVL(!收费类别)
            rsTmp!收据费目 = NVL(!收据费目)
            rsTmp!开单人 = NVL(!开单人)
            rsTmp!收费细目ID = Val(NVL(!收费细目ID))
            rsTmp!计算单位 = NVL(!计算单位)
            rsTmp!数量 = Val(NVL(!数量))
            rsTmp!单价 = Val(NVL(!单价))
            rsTmp!实收金额 = Val(NVL(!实收金额))
            rsTmp!统筹金额 = Val(NVL(!统筹金额))
            rsTmp!保险支付大类ID = IIf(Val(NVL(!保险支付大类ID)) = 0, Null, Val(NVL(!保险支付大类ID)))
            rsTmp!是否医保 = Val(NVL(!是否医保))
            rsTmp!保险编码 = NVL(!保险编码)
            rsTmp!摘要 = NVL(!摘要)
            rsTmp!是否急诊 = Val(NVL(!是否急诊))
            rsTmp!开单部门ID = Val(NVL(!开单部门ID))
            rsTmp!执行部门ID = Val(NVL(!执行部门ID))
            rsTmp.Update
            If InStr(strNos & ",", "," & !NO & ",") = 0 Then
                strNos = strNos & "," & !NO
                p = p + 1
            End If
            rsNo.MoveNext
        Loop
    End With
    If strNos <> "" Then strNos = Mid(strNos, 2)
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
    Set zlMakeClinicPreSwapData = rsTmp
    
    Exit Function
Errhand:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlInsureCheck(ByVal str保险结算 As String, ByVal strAdvance As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查当前的医保是否需要较对
    '入参:str保险结算-保险结算
    '       strAdvance-医保返回的结算
    '出参:
    '返回:需要较对,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-20 18:03:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnMedicareCheck As Boolean, strTmp As String, i As Long, j As Long
    Dim varData As Variant, varData1 As Variant
    Dim varTemp As Variant, varTemp1 As Variant

    On Error GoTo errHandle
    If Not (strAdvance <> "" And str保险结算 <> strAdvance) Then Exit Function
    '正式结算前后,结算方式和结算金额未发生变化时不校对
    blnMedicareCheck = True
    varData = Split(str保险结算, "||"): varData1 = Split(strAdvance, "||")

    If UBound(varData) = UBound(varData1) Then
        For i = 0 To UBound(varData)
            blnMedicareCheck = True
            strTmp = varData(i)
            varTemp = Split(strTmp, "|")
            For j = 0 To UBound(varData1)
                varTemp1 = Split(varData1(j), "|")
                If varTemp(0) = varTemp1(0) Then
                    If Val(varTemp(1)) = Val(varTemp1(1)) Then
                        blnMedicareCheck = False
                    End If
                End If
            Next
            If blnMedicareCheck Then Exit For
        Next
    End If
    zlInsureCheck = blnMedicareCheck
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetBalanceNos(ByVal bytType As Byte, _
    ByVal strFindValue As String, _
    Optional ByVal blnNOMoved As Boolean, _
    Optional bln补结算 As Boolean = False, _
    Optional int记录性质 As Integer = 1) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据一张收费单据的NO或结帐ID或结帐序号，返回同一次结算的NOs
    '入参:bytType-0-根据NO来查找;1-根据结帐ID来查找,2-根据结算序号来查找
    '    strFindValue-查找的值
    '    blnNOMoved-是否在后备表中，查询单据之前的判断需要用这个参数
    '    bln补结算-是否医保补结算
    '返回:格式如"AAA,BBB,CCC',..."
    '编制:刘兴洪
    '日期:2014-06-17 17:06:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String, strNos As String
    Dim i As Long, strHistory As String
    Dim strFeeTable As String, strSQL1 As String

    On Error GoTo errHandle:
    Select Case bytType
    Case 0 '0-根据NO来查找
        If bln补结算 Then
            strSQL = "" & _
            "   Select distinct A.NO " & _
            "   From 门诊费用记录 A,(Select distinct 收费结帐ID as 结帐ID From 费用补充记录 Where NO=[1] and 记录性质=1 ) B" & _
            "   Where A.结帐ID=B.结帐ID" & _
            "   Order by NO"
        Else
            strSQL = "" & _
            "   Select distinct B.NO " & _
            "   From 门诊费用记录 A,门诊费用记录 B" & _
            "   Where A.NO=[1] and Mod(A.记录性质,10)=1 And A.结帐ID=B.结帐ID And a.记录状态 In (1, 3)" & _
            "   Order by NO"
        End If
    Case 1  '1-根据结帐ID来查找
        If bln补结算 Then
            strSQL = "" & _
            "    Select Distinct A.No " & _
            "    From 门诊费用记录 A," & _
            "        (Select distinct C1.收费结帐ID as 结帐ID " & _
            "         From 费用补充记录 A1,费用补充记录 B1,费用补充记录 C1  " & _
            "         Where A1.结算ID=[2] and A1.记录性质=1  " & _
            "               And A1.NO=B1.NO and A1.记录性质=B1.记录性质 " & _
            "               And B1.结算序号=C1.结算序号 and C1.记录状态 in (1,3) ) B " & _
            "    Where A.结帐ID=B.结帐ID    " & _
            "    Order By NO"
        Else
            strSQL = "" & _
            "    Select Distinct c.No " & _
            "    From 门诊费用记录 A,门诊费用记录 B,门诊费用记录 C " & _
            "    Where A.结帐ID=[2] And Mod(a.记录性质, 10) = 1 And a.No = b.No And  b.记录性质 = 1 " & _
            "          and b.结帐ID=C.结帐ID    " & _
            "    Order By NO"
        End If
    Case 2  '2-根据结算序号来查找
        If bln补结算 Then
            strSQL = "" & _
            "    Select Distinct A.No " & _
            "    From 门诊费用记录 A," & _
            "        (Select distinct C1.收费结帐ID as 结帐ID " & _
            "         From 费用补充记录 A1,费用补充记录 B1,费用补充记录 C1  " & _
            "         Where A1.结算序号=[2] and A1.记录性质=1  " & _
            "               And A1.NO=B1.NO and A1.记录性质=B1.记录性质 " & _
            "               And B1.结算序号=C1.结算序号 and C1.记录状态 in (1,3) ) B " & _
            "    Where A.结帐ID=B.结帐ID    " & _
            "    Order By NO"
        Else
            strSQL = "" & _
            "   Select Distinct c.No " & _
            "   From 门诊费用记录 A,门诊费用记录 B,门诊费用记录 C," & _
            "        (Select 结帐ID From 病人预交记录 Where 结算序号=[2]) D" & _
            "   Where A.结帐ID=D.结帐ID And Mod(a.记录性质, 10) = 1 And a.No = b.No And  b.记录性质 = 1 " & _
            "         and b.结帐ID=C.结帐ID    " & _
            "   Order By NO"
        End If
    End Select
    If blnNOMoved Then
        strSQL = Replace(strSQL, "门诊费用记录", "H门诊费用记录")
        strSQL = Replace(strSQL, "病人预交记录", "H病人预交记录")
        If bln补结算 Then
            strSQL = Replace(strSQL, "费用补充记录", "H费用补充记录")
        End If
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "根据单据获取一次结帐的单据", strFindValue, Val(strFindValue))
    With rsTemp
        Do While Not .EOF
            strNos = strNos & "," & !NO
            .MoveNext
        Loop
    End With
    If strNos <> "" Then strNos = Mid(strNos, 2)
    zlGetBalanceNos = strNos
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlIsErrChargeCancel(ByVal strNo As String, Optional ByVal lng结帐ID As Long = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断指定的单据是否异常的收费作废操作异作
    '入参:strNO-按指定的单据判断
    '     lng结帐ID-按结帐ID进行判断
    '返回:是异常收费作废,则返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-07-29 14:11:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset

    On Error GoTo errHandle
    If strNo <> "" Then
        strSQL = "   Select 1 From 门诊费用记录 Where NO=[1] and 记录性质=1 And 记录状态 IN (1,3) and nvl(费用状态,0)=1"
    Else
        strSQL = "" & _
        "   Select 1 From 门诊费用记录 A,门诊费用记录 B  " & _
        "   Where A.NO=B.NO and A.记录性质=1 And A.记录状态 IN (1,3) and nvl(A.费用状态,0)=1 And B.结帐ID=[2]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查是否为异常收费作废", strNo, lng结帐ID)
    zlIsErrChargeCancel = Not rsTemp.EOF
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetFirstBalanceID(ByVal strNos As String, Optional ByVal blnNOMoved As Boolean, _
    Optional ByVal bln历史表同步查 As Boolean = False, _
    Optional lng结算序号 As Long) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取单据第一次结帐ID
    '入参:blnNoMoved是否在后备表中，查询单据之前的判断需要用这个参数
    '     bln历史表同步查-是否连接历史表一起查询
    '出参:lng结算序号-返回最后一次有效的结帐序号
    '返回:返回第一次结帐ID
    '编制:刘兴洪
    '日期:2014-07-30 13:57:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String, strSQL1 As String
    Dim i As Long
    
    On Error GoTo errHandle:
    strSQL = "" & _
        "   Select M.结帐ID,Q.结算序号" & _
        "   From 门诊费用记录 M,病人预交记录 Q" & _
        "   Where M.结帐ID=Q.结帐ID And  M.NO IN (select Column_value From Table(f_str2List([1])) ) " & _
        "         And  M.记录性质 =1 And M.记录状态 IN (1,3) And rownum <2 "

    If Not blnNOMoved And bln历史表同步查 Then
        strSQL1 = Replace(strSQL, "门诊费用记录", "H门诊费用记录")
        strSQL1 = Replace(strSQL1, "病人预交记录", "H病人预交记录")
        strSQL = strSQL & " Union ALL " & strSQL1
    ElseIf blnNOMoved Then
        strSQL = Replace(strSQL, "门诊费用记录", "H门诊费用记录")
        strSQL = Replace(strSQL, "病人预交记录", "H病人预交记录")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "根据单据获取最后一次正常结帐的结帐ID", strNos)
    If rsTemp.EOF Then Exit Function
    lng结算序号 = Val(NVL(rsTemp!结算序号))
    zlGetFirstBalanceID = Val(NVL(rsTemp!结帐ID))
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlExistDelFeeChargeBill(ByVal strNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据单据号,判断是否存在退费单据
    '入参:strNos-指定的收费单
    '返回:存在退费单据,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-06-17 17:06:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errHandle:
    strSQL = "" & _
        "   Select 1 " & _
        "   From 门诊费用记录 M" & _
        "   Where  M.NO in (Select Column_value From Table(f_str2List([1]))  )  " & _
        "       And Mod(M.记录性质,10)=1 And M.记录状态 =2 And Rownum <2 "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "判断是否存在退费单据", strNos)
    If rsTemp.EOF Then Exit Function
    zlExistDelFeeChargeBill = True
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlIsMulitOneBalance(ByVal strNos As String, Optional ByRef lng结帐ID As Long, _
    Optional ByRef lng结算序号 As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断指定单据是否多单据一次结算数据
    '入参:strNos-单据号(多个用逗号分隔)
    '出参:lng结帐ID-结帐ID
    '     lng结算序号-返回一次结算
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-08-01 14:42:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errHandle:
    strSQL = "" & _
        "   Select A.结帐ID,A.结算序号 " & _
        "   From  病人预交记录 A,门诊费用记录 M" & _
        "   Where  A.结帐ID=M.结帐ID And  M.NO in (Select Column_value From Table(f_str2List([1]))  )  " & _
        "       And  M.记录性质 =1 And M.记录状态 in (1,3) And Rownum <2 "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "判断是否存在退费单据", strNos)
    If rsTemp.EOF Then Exit Function
    lng结帐ID = Val(NVL(rsTemp!结帐ID))
    lng结算序号 = Val(NVL(rsTemp!结算序号))
    zlIsMulitOneBalance = lng结算序号 < 0
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlRePrintReplenishTheBalanceBill(frmParent As Object, ByVal lngModule As Long, _
    ByVal bytType As Byte, strNo As String, ByVal intInsure As Integer, _
    ByVal objInvoice As zlPublicExpense.clsInvoice, _
    ByVal objFact As zlPublicExpense.clsFactProperty, _
    Optional blnDelOpt As Boolean, Optional DateDel As Date, Optional blnVirtualPrint As Boolean, _
    Optional ByVal blnDelRecord As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新打印医保补充结算票据
    '入参:1-重打;2-补打
    '       strNO -指定要重打的单据号
    '       intInsure-医保号
    '       objInvoice-发票对象
    '       blnDelOpt-退费重打操作调用
    '       DateDel-退费时间
    '       blnVirtualPrint-医保接口打印票据，HIS不调打印只走票号
    '       blnDelRecord-重打时，是否是对退费记录进行重打(目前只有北京医保(医保接口打印票据)才允许)
    '出参:
    '返回:打印成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-09-30 10:22:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strInvoice As String, strInfo As String
    Dim j As Integer, blnValid As Boolean, blnInput As Boolean
    Dim blnDo As Boolean, blnHaveInvoice As Boolean
    Dim lngLastUseIDTemp As Long
    Dim strRptName As String, strInvoiceNO As String
    Dim bytPrintType As Byte
    Dim lngUseId As Long

    blnHaveInvoice = objFact.LastUseID <> 0     '主要是退费用,如果存在领用的,则必须回收发票,然后重打发票:30386
    If blnHaveInvoice = False And blnDelOpt Then
        blnHaveInvoice = objInvoice.zlCheckBillNOIsPrintInvoice(1, strNo)
    End If

    strRptName = "ZL" & glngSys \ 100 & "_BILL_1124"

    lngUseId = objFact.LastUseID
    '如果严格控制票据使用
    If objFact.严格控制 Then
        '此时只判断是否有,打印之前再根据张数判断是否够用
        If objInvoice.zlGetInvoiceGroupID(1124, UserInfo.姓名, EM_收费收据, objFact.使用类别, lngUseId, objFact.共享批次ID, lngUseId, 1, strInvoiceNO) = False Then Exit Function
        Select Case lngUseId
            Case -1
                MsgBox "你没有自用和共用的收费票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
            Case -2
                MsgBox "本地的共用票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
        End Select
        If lngUseId <= 0 Then Exit Function
    End If

    If Not gobjTax Is Nothing And gblnTax Then
        blnDo = True
    Else
        If blnDelOpt Then
            blnDo = True
        Else
            If objFact.打印格式 = 0 Then   '以缺省票据格式显示
                objFact.打印格式 = Val(GetReportPrintSet(gcnOracle, glngSys, strRptName, gstrDBUser, 1, , "Format"))
            End If
            SetReportPrintSet gcnOracle, glngSys, strRptName, "Format", objFact.打印格式
            '由于没有格式的传入,因此,需要强制缺省到指定格式
            blnDo = ReportPrintSet(gcnOracle, glngSys, strRptName, frmParent)
            '取出选择的格式
            objFact.打印格式 = Val(GetReportPrintSet(gcnOracle, glngSys, strRptName, gstrDBUser, 1, , "Format"))
        End If
    End If

    If blnDo Then
        '取下一个票据号码
        If Not objFact.严格控制 Then

            '有可能是第一次使用
            Do
                blnInput = False
                '非严格控制时直接从本地读取
                strInvoice = UCase(zlDatabase.GetPara("当前收费票据号", glngSys, 1124, ""))
                If strInvoice = "" Then
                    strInvoice = UCase(InputBox("没有找到已用的最大票据号码，无法确定将要使用的开始票据号。" & _
                                    vbCrLf & "请输入将要使用的开始票据号码：", gstrSysName, _
                                    "", frmParent.Left + 1500, frmParent.Top + 1500))
                    blnInput = True
                Else
                    strInvoice = zlstr.Increase(strInvoice)
                    strInvoice = UCase(InputBox("请确认" & IIf(bytType = 1, "重打", "补打") & "使用的开始票据号码：", gstrSysName, _
                                    strInvoice, frmParent.Left + 1500, frmParent.Top + 1500))
                    blnInput = True
                End If

                '用户取消输入,允许打印
                If strInvoice = "" Then
                    If MsgBox("你确定不输入票据号继续打印吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                    blnValid = True
                Else
                    '检查输入有效性
                    If blnInput Then
                        If zlCommFun.ActualLen(strInvoice) <> objFact.票号长度 Then
                            MsgBox "输入的票据号码长度应该为 " & objFact.票号长度 & " 位！", vbInformation, gstrSysName
                        Else
                            blnValid = True
                        End If
                    Else
                        blnValid = True
                    End If
                End If
            Loop While Not blnValid
        Else
            Do
                '根据票据领用读取
                blnInput = False
                If objInvoice.zlGetNextBill(1124, lngUseId, strInvoice) = False Then
                    strInvoice = ""
                End If

                If strInvoice = "" Then
                    '如果中途换用靠后的号码,可能造成未用完,但下一号码已超出范围
                    '30386:打印了发票的,必需重打再发出
                    If frmInputBox.InputBox(frmParent, "开始发票号", "无法根据票据领用情况获取将要使用的开始票据号，" & _
                                    vbCrLf & "请你输入将要使用的开始票据号码：", 30, 1, False, False, strInvoice, _
                                    blnHaveInvoice And blnDelOpt, frmParent.Left + 1500, frmParent.Top + 1500) = False Then
                                Exit Function
                    End If
                    blnInput = True
                Else
                    '30386
                    If frmInputBox.InputBox(frmParent, "开始发票号", "请确认" & IIf(bytType = 1, "重打", "补打") & "使用的开始票据号码：", 30, 1, False, False, strInvoice, _
                                    blnHaveInvoice And blnDelOpt, frmParent.Left + 1500, frmParent.Top + 1500) = False Then
                                Exit Function
                    End If
                    blnInput = True
                End If

                '用户取消输入,不打印
                If strInvoice = "" Then Exit Function

                '检查输入有效性
                If blnInput Then
                    If objInvoice.zlGetInvoiceGroupID(1124, UserInfo.姓名, EM_收费收据, objFact.使用类别, lngUseId, objFact.共享批次ID, lngLastUseIDTemp, 1, strInvoiceNO) = False Then Exit Function
                    If lngLastUseIDTemp = -3 Then
                        MsgBox "你输入的票据号码不在当前领用批次的有效领用范围内,请重新输入！", vbInformation, gstrSysName
                    Else
                        lngUseId = lngLastUseIDTemp
                        blnValid = True
                    End If
                Else
                    blnValid = True
                End If
            Loop While Not blnValid
        End If

        bytPrintType = IIf(blnDelOpt, 3, 2)
        If blnDelOpt Then
            Call frmReplenishTheBalancePrint.ReportPrint(bytPrintType, strNo, intInsure, objFact, "", lngUseId, strInvoice, DateDel, blnVirtualPrint)
        Else
            Call frmReplenishTheBalancePrint.ReportPrint(bytPrintType, strNo, intInsure, objFact, "", lngUseId, strInvoice, zlDatabase.Currentdate, blnVirtualPrint, blnDelRecord)
        End If
        zlRePrintReplenishTheBalanceBill = True
    End If
End Function

Public Function zlPrintReplenishTheDelBalanceBill(frmParent As Object, ByVal lngModule As Long, _
    ByVal lng结算序号 As Long, ByVal intInsure As Integer, _
    ByVal objInvoice As zlPublicExpense.clsInvoice, _
    ByVal objFact As zlPublicExpense.clsFactProperty, _
    Optional blnDelOpt As Boolean, Optional DateDel As Date, Optional blnVirtualPrint As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打印医保补充结算退费票据(红票)
    '入参:
    '       lng结算序号 -指定要打印单据的结算序号
    '       intInsure-医保号
    '       objInvoice-发票对象
    '       blnDelOpt-退费重打操作调用
    '       DateDel-退费时间
    '       blnVirtualPrint-医保接口打印票据，HIS不调打印只走票号
    '出参:
    '返回:打印成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-09-30 10:22:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strInvoice As String, strInfo As String
    Dim j As Integer, blnValid As Boolean, blnInput As Boolean
    Dim blnDo As Boolean, blnHaveInvoice As Boolean
    Dim lngLastUseIDTemp As Long
    Dim strRptName As String, strInvoiceNO As String
    Dim bytPrintType As Byte
    Dim lngUseId As Long
    
    strRptName = "ZL" & glngSys \ 100 & "_BILL_1124_3"
    
    lngUseId = objFact.LastUseID
    '如果严格控制票据使用
    If objFact.严格控制 Then
        '此时只判断是否有,打印之前再根据张数判断是否够用
        If objInvoice.zlGetInvoiceGroupID(1124, UserInfo.姓名, EM_收费收据, objFact.使用类别, lngUseId, objFact.共享批次ID, lngUseId, 1, strInvoiceNO) = False Then Exit Function
        Select Case lngUseId
            Case -1
                MsgBox "你没有自用和共用的收费票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
            Case -2
                MsgBox "本地的共用票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
        End Select
        If lngUseId <= 0 Then Exit Function
    End If


    If Not gobjTax Is Nothing And gblnTax Then
        blnDo = True
    Else
        If blnDelOpt Then
            blnDo = True
        Else
            If objFact.打印格式 = 0 Then   '以缺省票据格式显示
                objFact.打印格式 = Val(GetReportPrintSet(gcnOracle, glngSys, strRptName, gstrDBUser, 1, , "Format"))
            End If
            SetReportPrintSet gcnOracle, glngSys, strRptName, "Format", objFact.打印格式
            '由于没有格式的传入,因此,需要强制缺省到指定格式
            blnDo = ReportPrintSet(gcnOracle, glngSys, strRptName, frmParent)
            '取出选择的格式
            objFact.打印格式 = Val(GetReportPrintSet(gcnOracle, glngSys, strRptName, gstrDBUser, 1, , "Format"))
        End If
    End If
    
    If blnDo Then
        '取下一个票据号码
        If Not objFact.严格控制 Then
            
            '有可能是第一次使用
            Do
                blnInput = False
                '非严格控制时直接从本地读取
                strInvoice = UCase(zlDatabase.GetPara("当前收费票据号", glngSys, 1124, ""))
                If strInvoice = "" Then
                    strInvoice = UCase(InputBox("没有找到已用的最大票据号码，无法确定将要使用的开始票据号。" & _
                                    vbCrLf & "请输入将要使用的开始票据号码：", gstrSysName, _
                                    "", frmParent.Left + 1500, frmParent.Top + 1500))
                    blnInput = True
                Else
                    strInvoice = zlCommFun.IncStr(strInvoice)
                    strInvoice = UCase(InputBox("请确认使用的开始票据号码：", gstrSysName, _
                                    strInvoice, frmParent.Left + 1500, frmParent.Top + 1500))
                    blnInput = True
                End If
                    
                '用户取消输入,允许打印
                If strInvoice = "" Then
                    If MsgBox("你确定不输入票据号继续打印吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                    blnValid = True
                Else
                    '检查输入有效性
                    If blnInput Then
                        If zlCommFun.ActualLen(strInvoice) <> objFact.票号长度 Then
                            MsgBox "输入的票据号码长度应该为 " & objFact.票号长度 & " 位！", vbInformation, gstrSysName
                        Else
                            blnValid = True
                        End If
                    Else
                        blnValid = True
                    End If
                End If
            Loop While Not blnValid
        Else
            Do
                '根据票据领用读取
                blnInput = False
                If objInvoice.zlGetNextBill(1124, lngUseId, strInvoice) = False Then
                    strInvoice = ""
                End If
                
                If strInvoice = "" Then
                    '如果中途换用靠后的号码,可能造成未用完,但下一号码已超出范围
                    '30386:打印了发票的,必需重打再发出
                    If frmInputBox.InputBox(frmParent, "开始发票号", "无法根据票据领用情况获取将要使用的开始票据号，" & _
                                    vbCrLf & "请你输入将要使用的开始票据号码：", 30, 1, False, False, strInvoice, _
                                    False, frmParent.Left + 1500, frmParent.Top + 1500) = False Then
                                Exit Function
                    End If
                    blnInput = True
                Else
                    '30386
                    If frmInputBox.InputBox(frmParent, "开始发票号", "请确认使用的开始票据号码：", 30, 1, False, False, strInvoice, _
                                    blnHaveInvoice And blnDelOpt, frmParent.Left + 1500, frmParent.Top + 1500) = False Then
                                Exit Function
                    End If
                    blnInput = True
                End If
                
                '用户取消输入,不打印
                If strInvoice = "" Then Exit Function
                
                '检查输入有效性
                If blnInput Then
                    If objInvoice.zlGetInvoiceGroupID(1124, UserInfo.姓名, EM_收费收据, objFact.使用类别, lngUseId, objFact.共享批次ID, lngLastUseIDTemp, 1, strInvoiceNO) = False Then Exit Function
                    If lngLastUseIDTemp = -3 Then
                        MsgBox "你输入的票据号码不在当前领用批次的有效领用范围内,请重新输入！", vbInformation, gstrSysName
                    Else
                        lngUseId = lngLastUseIDTemp
                        blnValid = True
                    End If
                Else
                    blnValid = True
                End If
            Loop While Not blnValid
        End If
        
        '1-新单打印,2-重打,3-退费打印,4-补打票据(只有:2-按系统预定规则和3-用户自定规则时才转入),6-退费票据(红票)打印
        Call frmReplenishTheBalancePrint.ReportPrint(6, lng结算序号, intInsure, objFact, "", _
            lngUseId, strInvoice, DateDel, blnVirtualPrint)
        
        zlPrintReplenishTheDelBalanceBill = True
    End If
End Function

Public Function zlCheckRegBillIsExecuted(ByVal strNo As String, ByVal bln不含作废医嘱 As Boolean, _
    ByRef blnExecuted_Out As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断指定的挂号单据是否已经被执行,包括医生接诊下医嘱后作废后,取消接诊,也表示执行过了
    '入参:strNO-挂号单号
    '     bln不含作废医嘱-是否不包含已作废的医嘱
    '出参:
    '返回:True 表示已被执行
    '编制:刘兴洪
    '日期:2014-10-10 11:16:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String

    On Error GoTo errH
    blnExecuted_Out = False
    If bln不含作废医嘱 Then strSQL = " And 医嘱状态<>4"
    strSQL = _
        " Select count(ID) num From 病人挂号记录 Where NO=[1] And 执行状态>0 and 记录性质=1 and 记录状态 =1 " & _
        " Union All " & _
        " Select count(ID) num From 病人医嘱记录 Where 挂号单=[1] And (病人来源=1 or 病人来源=2)" & strSQL
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlRegEvent", strNo)
    Do While Not rsTmp.EOF
        If rsTmp!Num > 0 Then
            blnExecuted_Out = True
        End If
        rsTmp.MoveNext
    Loop
    zlCheckRegBillIsExecuted = True
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlGetInsureBalanceDetail(ByVal bytType As Byte, _
    ByVal strValue As String, Optional blnHistory As Boolean) As ADODB.Recordset
    '功能:获取医保结算明细数据
    '入参:bytType-查找类型:0-根据结帐ID查找;1-根据结算序号查找,2-根据单据号来获取结算方式
    '     strValue-要查找的值(为0时,结帐ID,为1时,结算序号,2时为一次收费所涉及的所有单据)
    '返回:返回医保结算明细记录
    '       字段:结帐id,NO,结算方式,金额
    '编制:冉俊明
    '日期:2015-07-13
    Dim strSQL As String, strWhere As String
    Dim strTable As String, strTable1 As String
    
    On Error GoTo errHandle
    strTable = IIf(blnHistory, "H", "") & "医保结算明细 A"
    Select Case bytType
    Case 0  '0-根据结帐ID查找
        strWhere = " And  A.结帐ID= [1]"
    Case 1  '1-根据结算序号查找
        strTable1 = "Select distinct 结帐ID  " & _
            "    From 病人预交记录 Where 结算序号= [1]"
        strTable1 = ",(" & strTable1 & ") B"
        If blnHistory Then strTable1 = Replace(strTable1, "病人预交记录", "H病人预交记录")
        strWhere = "  And A.结帐ID = B.结帐ID"
    Case 2 '2-根据单据号来获取结算数据
        strWhere = " And A.NO in (Select Column_value From Table(f_str2List([2])))"
    End Select
    
    strSQL = "Select a.结帐id, a.NO, a.结算方式, a.金额" & _
            " From " & strTable & strTable1 & _
            " Where 1=1 " & strWhere
    Set zlGetInsureBalanceDetail = zlDatabase.OpenSQLRecord(strSQL, "获取医保结算明细数据", Val(strValue), strValue)
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlGetYBBalanceNo(ByVal lng结帐ID As Long, Optional ByVal strNos As String, _
    Optional ByVal lng病人ID As Long, Optional ByVal intInsure As Integer, _
    Optional ByVal blnDelCheck As Boolean, Optional ByVal blnHistory As Boolean) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:按单据获取医保原结算方式和结算金额
    '参数：
    '   strNOs - 单据号,多个用逗号隔开：A0001,A0002,...
    '   blnDelCheck - 是否检查允许门诊结算作废
    '返回:返回结算信息,格式:结算方式|结算金额||...
    '编制:刘兴洪
    '日期:2014-07-07 09:57:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str结算方式 As String, varNos As Variant, strFilter As String
    Dim rsBalance As ADODB.Recordset, i As Integer, p As Integer
    Dim colBalance As Collection, strTemp As String
    
    On Error GoTo errHandle
    If blnDelCheck And intInsure = 0 Then Exit Function
    
    '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡)
    Set rsBalance = zlGetInsureBalanceDetail(0, lng结帐ID, blnHistory)
    If strNos <> "" Then
        varNos = Split(strNos, ",")
        For i = 0 To UBound(varNos)
            If UBound(varNos) < 1 Then '一张单据
                strFilter = " or No='" & varNos(i) & "'"
            Else '多张单据
                strFilter = strFilter & " or No='" & varNos(i) & "'"
            End If
        Next
        If strFilter <> "" Then strFilter = Mid(strFilter, 4)
        rsBalance.Filter = strFilter
    End If
    rsBalance.Sort = "No"
    If rsBalance.RecordCount = 0 Then Exit Function
    
    Set colBalance = New Collection
    p = 1: colBalance.Add Array()
    If rsBalance.RecordCount > 0 Then
        With rsBalance
            strTemp = NVL(!NO)
            Do While Not .EOF
                If strTemp <> NVL(!NO) Then
                    p = p + 1: colBalance.Add Array()
                    strTemp = NVL(!NO)
                End If
                If blnDelCheck Then
                    '如果这种结算方式不支持回退,要退为现金,则不用减去
                    If gclsInsure.GetCapability(support门诊结算作废, lng病人ID, intInsure, NVL(!结算方式)) Then
                        str结算方式 = NVL(!结算方式) & "|" & -1 * Val(NVL(!金额))
                    End If
                Else
                    str结算方式 = NVL(!结算方式) & "|" & Val(NVL(!金额))
                End If
                
                Call SetBalanceVal(colBalance, p, str结算方式)
                .MoveNext
            Loop
        End With
    End If
    zlGetYBBalanceNo = GetMedicareStr(colBalance)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlCheckOnlyUseTrans(ByVal str结算序号 As String) As Boolean
    '功能：检查医保报销金额是否大于总费用金额
    '入参：
    '   str结算序号 - 结算序号
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    Err = 0: On Error GoTo errHandler
    '剩余费用金额
    strSQL = "Select Nvl(Sum(冲预交), 0) As 金额" & vbNewLine & _
            " From (Select Nvl(Sum(冲预交), 0) As 冲预交" & vbNewLine & _
            "       From 病人预交记录" & vbNewLine & _
            "       Where 结帐id In (Select 收费结帐id From 费用补充记录 Where 记录性质 = 1 And 结算序号 = [1])" & vbNewLine & _
            "       Union All"
    '本次对剩余费用的医保报销金额
    strSQL = strSQL & vbNewLine & _
            "       Select -1 * Nvl(Sum(a.冲预交), 0) As 冲预交" & vbNewLine & _
            "       From 病人预交记录 A, 结算方式 B" & vbNewLine & _
            "       Where a.记录状态 = 1 And a.结算方式 = b.名称 And b.性质 In (3, 4) And a.结算序号= [1])"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查医保报销金额是否大于总费用金额", Val(str结算序号))
    zlCheckOnlyUseTrans = Val(NVL(rsTemp!金额)) < 0
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlGetCanDelBalanceRecords(ByVal lng结算序号 As Long, ByVal lng卡类别ID As Long) As ADODB.Recordset
    '功能：获取补充结算可退费的交易结算记录
    '入参：
    '   lng结算序号 - 结算序号
    '   lng卡类别ID - 医疗卡类别
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    
    Err = 0: On Error GoTo errHandler
    '原费用总金额，注意减去补结算前已退费的
    strSQL = _
        "Select Max(预交id) As 原交易ID, Max(原结帐id) As 结帐id, Max(卡号) As 卡号, Max(交易流水号) As 交易流水号," & vbNewLine & _
        "       Max(交易说明) As 交易说明,Max(冲预交) As 冲预交" & vbNewLine & _
        "From (Select Decode(a.记录状态, 2, f.Id, e.Id) As 预交id, Decode(a.记录状态, 2, f.结帐id, e.结帐id) As 原结帐id," & vbNewLine & _
        "             Decode(a.记录状态, 2, f.卡号, e.卡号) As 卡号, Decode(a.记录状态, 2, f.交易流水号, e.交易流水号) As 交易流水号," & vbNewLine & _
        "             Decode(a.记录状态, 2, f.交易说明, e.交易说明) As 交易说明, e.冲预交, a.结帐id" & vbNewLine & _
        "      From 门诊费用记录 A, 病人预交记录 E, 门诊费用记录 B, 病人预交记录 F, 费用补充记录 C" & vbNewLine & _
        "      Where a.记录性质 = b.记录性质 And a.No = b.No And a.序号 = b.序号 And a.结帐id = e.结帐id And b.结帐id = f.结帐id" & vbNewLine & _
        "            And b.结帐id = c.收费结帐id And b.记录状态 <> 2 And e.卡类别id = [2] And f.卡类别id = [2]" & vbNewLine & _
        "            And c.记录性质 = 1 And c.结算序号 = [1]" & vbNewLine & _
        "            And Not Exists (Select 1" & vbNewLine & _
        "                   From 病人预交记录" & vbNewLine & _
        "                   Where 结算序号 In (Select m.结算序号" & vbNewLine & _
        "                       From 费用补充记录 M, 费用补充记录 N" & vbNewLine & _
        "                       Where m.记录性质 = n.记录性质 And m.No = n.No And n.记录性质 = 1 And n.结算序号 = [1])" & vbNewLine & _
        "                             And 结帐id = a.结帐id))" & vbNewLine & _
        "Group By 结帐id"

    '补结算后已退费金额
    strSQL = strSQL & vbNewLine & _
        "Union All" & vbNewLine & _
        "Select 0, 记录id, '', '', '', -1 * 金额" & vbNewLine & _
        "From 三方退款信息" & vbNewLine & _
        "Where 记录id In (Select 收费结帐id From 费用补充记录 Where 记录性质 = 1 And 结算序号 = [1])"
    strSQL = _
        "Select Max(原交易ID) As 原交易ID, 结帐id, Max(卡号) As 卡号," & vbNewLine & _
        "       Max(交易流水号) As 交易流水号, Max(交易说明) As 交易说明, Nvl(Sum(冲预交), 0) As 金额" & vbNewLine & _
        "From (" & strSQL & ")" & vbNewLine & _
        "Group By 结帐id" & vbNewLine & _
        "Having Nvl(Sum(冲预交), 0) > 0" & vbNewLine & _
        "Order By 结帐id"
    Set zlGetCanDelBalanceRecords = zlDatabase.OpenSQLRecord(strSQL, "mdlBillRelate", lng结算序号, lng卡类别ID)
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlCheckOtherSessionDoing(ByVal lng结算序号 As Long, Optional ByVal strNos As String) As Boolean
    '功能:检查当前结算是否正在被其它会话处理
    '入参:lng结算序号-指定的结算序号
    '     strNos  - 单据号
    '出参:
    '返回:是返回true,否则返回False
    '说明："病人预交记录.会话号"格式：V$session.SID+'_'+V$session.SERIAL#
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errHandle
    If lng结算序号 = 0 And strNos = "" Then zlCheckOtherSessionDoing = False: Exit Function
    If strNos <> "" Then
        strSQL = "Select /*+cardinality(j,10) */ 1" & vbNewLine & _
                " From 病人预交记录 A, 门诊费用记录 B, Table(f_Str2list([2])) J, V$session C" & vbNewLine & _
                " Where a.结帐id = b.结帐id And Mod(b.记录性质, 10) = 1 And b.No = j.Column_Value" & vbNewLine & _
                "       And a.会话号 = c.Sid || '_' || c.Serial# And c.Username Is Not Null" & vbNewLine & _
                "       And c.Audsid <> Userenv('sessionid') And Upper(c.Status) In ('ACTIVE', 'INACTIVE') And Rownum < 2"
    Else
        strSQL = "Select 1" & vbNewLine & _
                " From 病人预交记录 A, V$session B" & vbNewLine & _
                " Where a.会话号 = b.Sid || '_' || b.Serial# And (a.结算序号 = [1] Or a.结帐ID = [1])" & vbNewLine & _
                "       And b.Username Is Not Null And b.Audsid <> Userenv('sessionid')" & vbNewLine & _
                "       And Upper(b.Status) In ('ACTIVE', 'INACTIVE') And Rownum < 2"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查当前结算是否正在被其它会话处理", lng结算序号, strNos)
    zlCheckOtherSessionDoing = Not rsTemp.EOF
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetForceDelToCashNote(ByRef cllForceDelToCash As Collection) As String
    '获取强制退现摘要，存入"交易说明"自段中，格式：XXXX强制退现:XXX卡;XXX卡
    '入参：
    '   cllForceDelToCash Array(操作员,卡类别名称)
    Dim str操作员 As String
    Dim strTemp As String, i As Integer
    
    On Error GoTo errHandler
    If cllForceDelToCash Is Nothing Then Exit Function
    If cllForceDelToCash.Count = 0 Then Exit Function
    
    str操作员 = cllForceDelToCash(1)(0)
    For i = 1 To cllForceDelToCash.Count
        strTemp = strTemp & ";" & cllForceDelToCash(i)(1)
    Next
    If strTemp <> "" Then strTemp = Mid(strTemp, 2)
    
    GetForceDelToCashNote = str操作员 & "强制退现：" & strTemp
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ThreeBalanceCheck(frmMain As Form, ByVal lngModule As Long, _
    ByVal objCard As Card, ByRef cllForceDelToCash As Collection, _
    Optional ByVal str卡类别名称 As String, Optional ByRef bln强制退现 As Boolean) As Boolean
    '三方卡强制退现检查
    '入参：
    '   objCard 医疗卡信息
    '   str卡类别名称 卡类别名称
    '出参：
    '   cllForceDelToCash 强制退现信息：Array(操作员,卡类别名称)
    '返回：允许强制退现，返回True；否则，返回False
    '105432
    Dim str操作员 As String
    
    On Error GoTo errHandler
    bln强制退现 = False
    If cllForceDelToCash Is Nothing Then Set cllForceDelToCash = New Collection
    
    If objCard Is Nothing Then
        If MsgBox("未找到指定的医疗卡，无法判断该医疗卡是否支持退现，你确定要强制退为其它结算方式吗？", _
                    vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    Else
        If Not (objCard.接口序号 > 0 And Not objCard.消费卡) Then ThreeBalanceCheck = True: Exit Function
        If objCard.是否退现 Then ThreeBalanceCheck = True: Exit Function
    End If
    
    If zlstr.IsHavePrivs(GetPrivFunc(glngSys, 1151), "三方退款强制退现") Then
        If MsgBox("『" & str卡类别名称 & "』不支持退现，你确定要将其强制退现吗？", _
            vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        'Array(操作员,卡类别名称)
        cllForceDelToCash.Add Array(UserInfo.姓名, str卡类别名称)
    Else
        str操作员 = zlDatabase.UserIdentifyByUser(frmMain, "『" & str卡类别名称 & "』强制退现，权限验证：", _
            glngSys, lngModule, "三方退款强制退现", , True)
        If str操作员 = "" Then Exit Function
        'Array(操作员,卡类别名称)
        cllForceDelToCash.Add Array(str操作员, str卡类别名称)
    End If
    bln强制退现 = True
    ThreeBalanceCheck = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlGetDelXMLExpend(ByVal lng结算序号 As String, ByVal bln异常作废 As Boolean) As String
    '获取传入三方卡退费接口zlRetuenCheck中strXMLExpend参数值
    '主要用于收费异常单据作废
    '入参：
    '   lng结算序号 - 结算序号
    '   bln异常作废 - 是否收费异常作废异常再作废
    Dim i As Integer, strPriorNO As String
    Dim strSQL As String, rsRecord As ADODB.Recordset
    Dim strXMLExpend As String, strXMLSub As String
    'strXMLExpend说明:
    '<TFDATA> //退费数据
    '  <YCTF>1</YCTF> //是否异常重退:1-异常重退;0-退费 此节点可能没有
    '  <TFLIST> //退费列表
    '    <NO></NO> // 退费单据
    '    <TFITEM> //退费项
    '      <SerialNum></SerialNum> //序号
    '      …
    '    </TFITEM>
    '  </TFLIST>
    '  ....
    '</TFDATA >
    
    On Error GoTo errHandler
    If lng结算序号 = 0 Then Exit Function
    strXMLExpend = "": strXMLSub = ""
    
    strSQL = "Select /*+cardinality(b,10)*/ Distinct a.NO, a.序号" & vbNewLine & _
            " From 门诊费用记录 A, 病人预交记录 B" & vbNewLine & _
            " Where a.记录性质 = 1 And a.结帐id = b.结帐id And b.结算序号 = [1]" & vbNewLine & _
            " Order By a.NO, a.序号"
    Set rsRecord = zlDatabase.OpenSQLRecord(strSQL, "获取费用单据", lng结算序号)
    If rsRecord.RecordCount = 0 Then Exit Function
    
    strXMLExpend = strXMLExpend & "<TFDATA>" & vbCrLf '退费数据
    strXMLExpend = strXMLExpend & "  <YCTF>" & IIf(bln异常作废, 1, 0) & "</YCTF>" & vbCrLf '是否异常重退:1-异常重退;0-退费
    Do While Not rsRecord.EOF
        If NVL(rsRecord!NO) <> strPriorNO Then
            If strPriorNO <> "" Then
                strXMLExpend = strXMLExpend & "    </TFITEM>" & vbCrLf
                strXMLExpend = strXMLExpend & "  </TFLIST>" & vbCrLf
            End If
            strXMLExpend = strXMLExpend & "  <TFLIST>" & vbNewLine '退费列表
            strXMLExpend = strXMLExpend & "    <NO>" & NVL(rsRecord!NO) & "</NO>" & vbCrLf '退费单据
            strXMLExpend = strXMLExpend & "    <TFITEM>" & vbCrLf '退费项
        End If
        strXMLExpend = strXMLExpend & "      <SerialNum>" & Val(NVL(rsRecord!序号)) & "</SerialNum>" & vbCrLf '序号
        
        strPriorNO = NVL(rsRecord!NO)
        rsRecord.MoveNext
    Loop
    strXMLExpend = strXMLExpend & "    </TFITEM>" & vbCrLf
    strXMLExpend = strXMLExpend & "  </TFLIST>" & vbCrLf
    strXMLExpend = strXMLExpend & "</TFDATA>"
    
    ZlGetDelXMLExpend = strXMLExpend
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlGetDelXMLExpendByGrid(ByVal vsfBill As VSFlexGrid) As String
    '从界面表格中获取传入三方卡退费接口zlRetuenCheck中strXMLExpend参数值
    Dim i As Integer
    Dim strXMLExpend As String, blnFindSelectItem As Boolean
    Dim strNo As String, strPriorNO As String
    'strXMLExpend说明:
    '<TFDATA> //退费数据
    '  <YCTF>1</YCTF> //是否异常重退:1-异常重退;0-退费 此节点可能没有
    '  <TFLIST> //退费列表
    '    <NO></NO> // 退费单据
    '    <TFITEM> //退费项
    '      <SerialNum></SerialNum> //序号
    '      …
    '    </TFITEM>
    '  </TFLIST>
    '  ....
    '</TFDATA >
    
    On Error GoTo errHandler
    strXMLExpend = "": blnFindSelectItem = False
    
    strXMLExpend = strXMLExpend & "<TFDATA>" & vbCrLf '退费数据
    strXMLExpend = strXMLExpend & "  <YCTF>" & 0 & "</YCTF>" & vbCrLf '是否异常重退:1-异常重退;0-退费
    With vsfBill
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("选择"))) <> 0 Then
                blnFindSelectItem = True
                strNo = .TextMatrix(i, .ColIndex("单据号"))
                If strNo <> strPriorNO Then
                    If strPriorNO <> "" Then
                        strXMLExpend = strXMLExpend & "    </TFITEM>" & vbCrLf
                        strXMLExpend = strXMLExpend & "  </TFLIST>" & vbCrLf
                    End If
                    strXMLExpend = strXMLExpend & "  <TFLIST>" & vbNewLine '退费列表
                    strXMLExpend = strXMLExpend & "    <NO>" & strNo & "</NO>" & vbCrLf '退费单据
                    strXMLExpend = strXMLExpend & "    <TFITEM>" & vbCrLf '退费项
                End If
                strXMLExpend = strXMLExpend & "      <SerialNum>" & .RowData(i) & "</SerialNum>" & vbCrLf '序号
                strPriorNO = strNo
            End If
        Next
    End With
    If blnFindSelectItem = False Then Exit Function
    strXMLExpend = strXMLExpend & "    </TFITEM>" & vbCrLf
    strXMLExpend = strXMLExpend & "  </TFLIST>" & vbCrLf
    strXMLExpend = strXMLExpend & "</TFDATA>"
    
    ZlGetDelXMLExpendByGrid = strXMLExpend
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
