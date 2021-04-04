VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisPackage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Get_RequestTime(ByVal lng病人ID As Long, ByVal lng主页ID As String, ByVal str挂号单号 As String) As ADODB.Recordset
'******************************************************************************************************************
'功能：获取检验申请的开始结束时间
'参数：lng主页ID=0:门诊，否者为住院
'返回：
'附注：
'******************************************************************************************************************

    Dim strSql As String

    On Error GoTo ErrH

    If Nvl(lng主页ID, 0) = 0 Then
        strSql = "select 急诊," & vbNewLine & _
               "       Null As 申请开始时间," & vbNewLine & _
               "       Trunc(sysdate) + 1 - 1 / 24 / 60 / 60 As 申请结束时间," & vbNewLine & _
               "       Null As 实际出院时间" & vbNewLine & _
               "  from 病人挂号记录 A" & vbNewLine & _
               " where A.NO = [2]" & vbNewLine & _
               "   And A.病人ID = [1]"
    Else
        strSql = "select 0 As 急诊," & vbNewLine & _
               "       A.入院日期 As 申请开始时间," & vbNewLine & _
               "       Nvl(A.出院日期,sysdate) As 申请结束时间," & vbNewLine & _
               "       A.出院日期 As 实际出院时间" & vbNewLine & _
               "  from 病案主页 A" & vbNewLine & _
               " where A.主页ID=[2]" & vbNewLine & _
               "   And A.病人ID = [1]"
    End If

    Set Get_RequestTime = gobjDatabase.OpenSQLRecord(strSql, gstrSysName, lng病人ID, IIf(lng主页ID = 0, str挂号单号, lng主页ID))

    Exit Function:
ErrH:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function Get_RequestOrder(ByVal lng病人ID As Long, ByVal lng主页ID As String, ByVal str申请开始时间 As String, ByVal str申请结束时间 As String) As ADODB.Recordset
'******************************************************************************************************************
'功能：获取检验申请记录
'参数：lng主页ID=0:门诊，否者为住院
'返回：
'附注：
'******************************************************************************************************************

    Dim strSql As String

    On Error GoTo ErrH

    If gblnNewLis Then
        strSql = "Select '检验' As 类别,C.分类,Nvl(C.微生物申请,0) As 微生物申请,F_List2Str(Cast(Collect(Nvl(C.名称,'')) As t_strlist)) As 检验项目,To_Char(A.核收时间, 'YYYYMMDD') || '-' || A.标本序号 As 标本序号," & vbNewLine & _
               "       Decode(A.急诊标本, 1, '急', '') As 急诊标本,D.名称 As 仪器,A.ID As 标本ID,A.姓名,Decode(A.病人来源, 1, '门诊', 2, '住院', 3, '体检') As 病人类型," & vbNewLine & _
               "       Decode(A.病人来源,1,'门 诊 号:' || A.病历号,2,'住 院 号:' || A.病历号,3,'健 康 号:' || A.病历号,'标 识 号:' || A.病历号) As 住院号," & vbNewLine & _
               "       A.挂号单,A.床号,Decode(A.性别, 0, '未区分', 1, '男', 2, '女', 9, '未说明') As 性别,A.年龄," & vbNewLine & _
               "       (case when A.病区 is not null then A.病区 else A.申请科室 end) As 申请科室,A.标本类型,A.申请人 As 申请医生," & vbNewLine & _
               "       To_Char(Min(A.申请时间), 'YYYY-MM-DD HH24:MI') As 申请时间,A.采样人,To_Char(A.采样时间, 'YYYY-MM-DD HH24:MI') As 采样时间," & vbNewLine & _
               "       A.送检人,To_Char(A.送检时间, 'YYYY-MM-DD HH24:MI') As 送检时间,A.签收人,To_Char(A.签收时间, 'YYYY-MM-DD HH24:MI') As 签收时间," & vbNewLine & _
               "       A.检验人 As 报告人,To_Char(A.报告时间, 'YYYY-MM-DD HH24:MI') As 报告时间,A.审核人," & vbNewLine & _
               "       To_Char(A.审核时间, 'YYYY-MM-DD HH24:MI') As 审核时间,A.备注 As 检验备注, A.临床特征" & vbNewLine & _
               "  From 检验报告记录 A, 检验申请组合 B, 检验组合项目 C, 检验仪器记录 D" & vbNewLine & _
               " Where A.ID = B.标本ID And B.组合ID = C.ID(+) And A.审核人 is not null And A.审核时间 is not null" & vbNewLine & _
               "   And A.仪器ID = D.ID And B.HIS病人ID = [1] And Nvl(B.主页ID,0) = [2]" & vbNewLine & _
               "   And B.申请时间 Between To_Date('" & Format(str申请开始时间, "yyyy-MM-dd HH:mm:ss") & "', 'YYYY-MM-DD HH24:MI:SS') And" & vbNewLine & _
               "       To_Date('" & Format(str申请结束时间, "yyyy-MM-dd HH:mm:ss") & "', 'YYYY-MM-DD HH24:MI:SS')" & IIf(gblnPrintAll, "", IIf(gblnDoctor, " and nvl(a.医生站打印,0)=0", " and nvl(a.打印次数,0)=0")) & vbNewLine & _
               " group by C.分类,Nvl(C.微生物申请,0),To_Char(A.核收时间, 'YYYYMMDD') || '-' || A.标本序号," & vbNewLine & _
               "       Decode(A.急诊标本, 1, '急', ''),D.名称,A.ID,A.姓名,Decode(A.病人来源, 1, '门诊', 2, '住院', 3, '体检')," & vbNewLine & _
               "       Decode(A.病人来源,1,'门 诊 号:' || A.病历号,2,'住 院 号:' || A.病历号,3,'健 康 号:' || A.病历号,'标 识 号:' || A.病历号)," & vbNewLine & _
               "       A.挂号单,A.床号,Decode(A.性别, 0, '未区分', 1, '男', 2, '女', 9, '未说明'),A.年龄," & vbNewLine & _
               "       (case when A.病区 is not null then A.病区 else A.申请科室 end),A.标本类型,A.申请人 ," & vbNewLine & _
               "       To_Char(A.申请时间, 'YYYY-MM-DD HH24:MI'),A.采样人,To_Char(A.采样时间, 'YYYY-MM-DD HH24:MI') ," & vbNewLine & _
               "       A.送检人,To_Char(A.送检时间, 'YYYY-MM-DD HH24:MI'),A.签收人,To_Char(A.签收时间, 'YYYY-MM-DD HH24:MI')," & vbNewLine & _
               "       A.检验人,To_Char(A.报告时间, 'YYYY-MM-DD HH24:MI'),A.审核人," & vbNewLine & _
               "       To_Char(A.审核时间, 'YYYY-MM-DD HH24:MI'),A.备注, A.临床特征" & vbNewLine & _
               " Order By C.分类, A.ID, Min(B.申请时间) Desc"
        ' And B.申请ID Is not Null
        Set Get_RequestOrder = gobjPublicHisCommLis.openSqlOtherDB(1, strSql, gstrSysName, lng病人ID, Val(lng主页ID))
    Else
        strSql = "Select distinct '检验' As 类别,D.操作类型 As 分类,Nvl(A.微生物标本, 0) As 微生物申请," & vbNewLine & _
               "  case when instr(a.标本类型,'血') > 0 then '1' || 标本类型 else '2' || 标本类型 end as 排序标本类型 , " & vbNewLine & _
               "    Decode(Instr(A.检验项目, '('),0,A.检验项目,substr(A.检验项目, 1, Instr(A.检验项目, '(') - 1)) As 检验项目," & vbNewLine & _
               "    Substr(Decode(Instr(A.检验项目, '('),0,A.检验项目,substr(A.检验项目, 1, Instr(A.检验项目, '(') - 1)),1,Decode(Instr(Decode(Instr(A.检验项目, '('),0,A.检验项目,substr(A.检验项目, 1, Instr(A.检验项目, '(') - 1)) ,',')-1,-1,4000,Instr(Decode(Instr(A.检验项目, '('),0,A.检验项目,substr(A.检验项目, 1, Instr(A.检验项目, '(') - 1)) ,',')-1 ) ) as 检验项目1 , " & vbNewLine & _
               "    To_Char(A.核收时间, 'YYYYMMDD') || '-' || A.标本序号 As 标本序号,Decode(A.紧急, 1, '急', '') As 急诊标本,F.名称 As 仪器," & vbNewLine & _
               "    A.ID As 标本ID,A.姓名,Decode(A.病人来源, 1, '门诊', 2, '住院', 3, '体检') As 病人类型," & vbNewLine & _
               "    Decode(A.病人来源,1,'门 诊 号:' || A.门诊号,2,'住 院 号:' || A.住院号,3,'健 康 号:' || A.标识号,'标 识 号:' || A.标识号) As 住院号," & vbNewLine & _
               "    A.床号,A.性别,A.年龄,E.名称 As 申请科室,A.标本类型,A.申请人 As 申请医生,To_Char(A.申请时间, 'YYYY-MM-DD HH24:MI') As 申请时间," & vbNewLine & _
               "    A.采样人,To_Char(A.采样时间, 'YYYY-MM-DD HH24:MI') As 采样时间,A.采样人 As 送检人,To_Char(A.核收时间, 'YYYY-MM-DD HH24:MI') As 送检时间," & vbNewLine & _
               "    A.核收人 As 签收人,To_Char(A.核收时间, 'YYYY-MM-DD HH24:MI') As 签收时间,A.检验人 As 报告人,To_Char(A.审核时间, 'YYYY-MM-DD HH24:MI') As 报告时间," & vbNewLine & _
               "    A.审核人,To_Char(A.审核时间, 'YYYY-MM-DD HH24:MI') As 审核时间,A.备注 As 检验备注,'' As 临床特征" & vbNewLine & _
               "  From 检验标本记录 A,病人医嘱记录 B,病人信息 C,诊疗项目目录 D,部门表 E,检验仪器 F,部门表 N,部门表 M,病案主页 S" & vbNewLine & _
               " Where A.申请科室ID = E.ID(+) And A.仪器ID = F.ID(+) And A.审核人 is not null And A.审核时间 is not null And A.病人ID = B.病人ID(+)" & vbNewLine & _
               "   And A.医嘱ID = B.相关ID And A.病人ID = C.病人ID And B.诊疗项目ID = D.ID And C.当前科室ID = M.ID(+) And C.当前病区ID = N.ID(+)" & vbNewLine & _
               "   And A.病人ID = S.病人ID(+) And A.主页ID = S.主页ID(+) And A.病人ID=[1]" & IIf(lng主页ID = 0, "", " And B.主页ID = [2]") & vbNewLine & _
               "   And A.申请时间 Between To_Date('" & Format(str申请开始时间, "yyyy-MM-dd HH:mm:ss") & "', 'YYYY-MM-DD HH24:MI:SS') And" & vbNewLine & _
               "   To_Date('" & Format(str申请结束时间, "yyyy-MM-dd HH:mm:ss") & "', 'YYYY-MM-DD HH24:MI:SS')" & IIf(gblnPrintAll, "", " and nvl(a.打印次数,0)=0") & vbNewLine & _
                 "Order By 排序标本类型, To_Char(A.审核时间, 'YYYY-MM-DD HH24:MI')  "
        Set Get_RequestOrder = gobjDatabase.OpenSQLRecord(strSql, gstrSysName, lng病人ID, lng主页ID)
    End If




    Exit Function
ErrH:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function Get_LisResult(ByVal lng标本id As Long, ByVal blnMultiline As Boolean) As ADODB.Recordset
'******************************************************************************************************************
'功能：获取检验结果
'参数：
'返回：
'附注：
'******************************************************************************************************************

    Dim strSql As String

    On Error GoTo ErrH

    If gblnNewLis Then
        If blnMultiline Then
            strSql = "select rownum As 序号,显示英文名,显示中文名,显示结果,显示标志,临床特征,结果参考,显示单位," & vbNewLine & _
                   "       中文名,英文名,检验结果,标志,单位,打印序号,排列序号,检验方法,显示OD,显示CUTOFF,显示SCO" & vbNewLine & _
                   "  from (select A.中文名,A.英文名,A.检验结果,A.标志,A.单位,A.结果参考,A.临床特征,A.打印序号,A.排列序号,A.检验方法," & vbNewLine & _
                   "               Decode(标记, 1, A.中文名, '') As 显示中文名,Decode(标记, 1, A.英文名, '') As 显示英文名," & vbNewLine & _
                   "               Decode(标记, 1, A.检验结果, '') As 显示结果,Decode(标记, 1, A.OD, '') As 显示标志," & vbNewLine & _
                   "               Decode(标记, 1, A.单位, '') As 显示单位,Decode(标记, 1, A.OD, '') As 显示OD," & vbNewLine & _
                   "               Decode(标记, 1, A.CUTOFF, '') As 显示CUTOFF,Decode(标记, 1, A.SCO, '') As 显示SCO" & vbNewLine & _
                   "          from (Select C.ID,B.中文名,B.英文名,A.项目ID,A.检验结果," & vbNewLine & _
                   "                       Decode(A.结果标志,'3','↑','2','↓','1','','4','↑','5','↓↓','6','↑↑','') As 标志," & vbNewLine & _
                   "                       B.单位,A.结果参考,ltrim(C.临床特征) As 临床特征,B.打印序号,B.检验方法,A.OD,A.CUTOFF,A.SCO," & vbNewLine & _
                   "                       B.排列序号,(case when Nvl(C.ID, 0) = Nvl(E.小号, 0) then '1' else '' End) As 标记" & vbNewLine & _
                   "                  From 检验报告明细 A,检验指标 B,检验指标参考 C," & vbNewLine & _
                   "                       (select min(ID) 小号, 仪器ID, 项目ID from 检验指标参考 group by 仪器ID, 项目ID) E" & vbNewLine & _
                   "                 Where A.项目ID = B.ID And A.项目ID = C.项目ID(+) And A.仪器ID = C.仪器ID(+) And A.项目ID = E.项目ID(+)" & vbNewLine & _
                   "                   And A.仪器ID = E.仪器ID(+) And A.标本ID = [1] And A.检验结果 Is Not Null" & vbNewLine & _
                   "                 Order By C.ID) A" & vbNewLine & _
                   "         where 检验结果 is not Null" & vbNewLine & _
                   "         order by A.排列序号, A.ID)"
        Else
            strSql = "select rownum As 序号,显示英文名,显示中文名,显示结果,显示标志,临床特征,结果参考,显示单位,中文名," & vbNewLine & _
                   "       英文名,检验结果,标志,单位,打印序号,排列序号,检验方法,显示OD,显示CUTOFF,显示SCO" & vbNewLine & _
                   "  from (select A.中文名,A.英文名,A.检验结果,A.标志,A.单位,A.结果参考,A.临床特征,A.打印序号,A.排列序号,A.检验方法," & vbNewLine & _
                   "               A.中文名 As 显示中文名,A.英文名 As 显示英文名,A.检验结果 As 显示结果,A.标志 As 显示标志,A.单位 As 显示单位," & vbNewLine & _
                   "               A.OD As 显示OD,A.CUTOFF As 显示CUTOFF,A.SCO As 显示SCO" & vbNewLine & _
                   "          from (Select B.中文名,B.英文名,A.项目ID,A.检验结果,B.单位,A.结果参考,'' As 临床特征,B.打印序号,B.检验方法," & vbNewLine & _
                   "                       Decode(A.结果标志,'3','↑','2','↓','1','','4','↑','5','↓↓','6','↑↑','') As 标志,B.排列序号," & vbNewLine & _
                   "                       A.OD,A.CUTOFF,A.SCO" & vbNewLine & _
                   "                  From 检验报告明细 A, 检验指标 B" & vbNewLine & _
                   "                 Where A.项目ID = B.ID And A.标本ID = [1] And A.检验结果 Is Not Null) A" & vbNewLine & _
                   "         where 检验结果 is not Null" & vbNewLine & _
                   "         order by A.排列序号)"
        End If
        Set Get_LisResult = gobjPublicHisCommLis.openSqlOtherDB(1, strSql, gstrSysName, lng标本id)

    Else
        If blnMultiline Then
            strSql = "select rownum As 序号,显示英文名,显示中文名,显示结果,显示标志,临床特征,结果参考,显示单位,中文名,英文名," & vbNewLine & _
                   "       检验结果,标志,单位,1 As 打印序号,排列序号,'' As 检验方法,显示OD,显示CUTOFF,显示SCO" & vbNewLine & _
                   "  from (Select decode(rownum, Y.内部序号, replace(英文名, '.', ''), '') As 显示英文名," & vbNewLine & _
                   "               decode(rownum, Y.内部序号, replace(中文名, '.', ''), '') As 显示中文名," & vbNewLine & _
                   "               decode(rownum, Y.内部序号, replace(OD, '.', ''), '') As 显示OD," & vbNewLine & _
                   "               decode(rownum, Y.内部序号, replace(CUTOFF, '.', ''), '') As 显示CUTOFF," & vbNewLine & _
                   "               decode(rownum, Y.内部序号, replace(SCO, '.', ''), '') As 显示SCO," & vbNewLine & _
                   "               decode(rownum,Y.内部序号,Decode(本次结果,'-','阴性(-)','+','阳性(+)','*','*.**',本次结果),'') As 显示结果," & vbNewLine & _
                   "               decode(rownum,Y.内部序号,decode(原始结果, '*', '', 标志),'') As 显示标志,Replace(临床特征, '111', '') As 临床特征," & vbNewLine & _
                   "               replace(trim(substr(参考,1,decode(instr(' ' || 参考,Chr(13) || Chr(10)),0,30,instr(' ' ||decode(substr(to_char(参考),1,1)," & vbNewLine & _
                   "               '.','0' || To_Char(参考),To_Char(参考)),Chr(13) || Chr(10))))),Chr(13) || Chr(10),'') As 结果参考," & vbNewLine & _
                   "               decode(Rownum, Y.内部序号, 单位, '') As 显示单位,中文名,英文名,本次结果 As 检验结果,标志,单位,排列序号"
            strSql = strSql & vbNewLine & "From (Select distinct A.检验项目ID,B.ID,A.检验项目 As 检验项目编码,A.排列序号,C.中文名,D.缩写 As 英文名,B.原始结果," & vbNewLine & _
                   "                      B.检验结果 As 本次结果,Decode(B.结果标志,3,'↑',2,'↓',1,'',4,'↑',5,'↓↓',6,'↑↑','') As 标志," & vbNewLine & _
                   "                      D.单位,Nvl(H.临床特征, '111') As 临床特征,B.OD,B.CUTOFF,B.SCO," & vbNewLine & _
                   "                      decode(substr(H.参考低值, 1, 1),'.','0' || To_Char(H.参考低值),To_Char(H.参考低值)) || '～' ||" & vbNewLine & _
                   "                      decode(substr(H.参考高值, 1, 1),'.','0' || To_Char(H.参考高值),To_Char(H.参考高值)) As 参考" & vbNewLine & _
                   "        From (Select A.检验项目ID," & vbNewLine & _
                   "                     Min(Decode(E.报告项目ID,A.检验项目ID,F.编码,99999)) As 检验项目," & vbNewLine & _
                   "                     Max(Decode(E.报告项目ID,A.检验项目ID,E.排列序号,F.编码)) As 排列序号" & vbNewLine & _
                   "                From 检验普通结果 A,检验报告项目 E,诊疗项目目录 F," & vbNewLine & _
                   "                     (Select Distinct C.诊疗项目ID From 检验项目分布 B, 病人医嘱记录 C Where B.标本ID=[1] And B.医嘱ID=C.相关ID) D" & vbNewLine & _
                   "               Where E.诊疗项目ID=D.诊疗项目ID And A.检验项目ID=E.报告项目ID(+) And E.诊疗项目ID=F.ID(+) And A.检验标本ID=[1]" & vbNewLine & _
                   "               Group By A.检验项目ID Order By Min(Decode(E.报告项目ID,A.检验项目ID,F.编码,99999))," & vbNewLine & _
                   "                        Max(Decode(E.报告项目ID,A.检验项目ID,E.排列序号,F.编码))) A," & vbNewLine & _
                "             检验普通结果 B,诊治所见项目 C,检验项目 D,检验标本记录 E,病人信息 F,检验仪器项目 G,检验项目参考 H"
            strSql = strSql & vbNewLine & "        Where B.检验项目ID=A.检验项目ID(+) And B.检验标本ID=[1] And B.检验项目ID=C.ID" & vbNewLine & _
                   "          And C.ID=D.诊治项目ID And B.检验标本ID=E.ID And E.病人ID=F.病人ID(+) And B.检验项目ID=G.项目ID(+)" & vbNewLine & _
                   "          And (G.仪器ID=E.仪器ID Or G.仪器ID Is Null Or E.仪器ID Is Null) And B.检验结果 is not Null And D.诊治项目ID=H.项目ID" & vbNewLine & _
                   "       Union All" & vbNewLine & _
                   "       Select distinct A.检验项目ID,B.ID,A.检验项目 As 检验项目编码,A.排列序号,C.中文名,D.缩写 As 英文名,B.原始结果," & vbNewLine & _
                   "                       B.检验结果 As 本次结果,Decode(B.结果标志,3,'↑',2,'↓',1,'',4,'↑',5,'↓↓',6,'↑↑','') As 标志," & vbNewLine & _
                   "                       D.单位,Nvl(临床特征, '111') As 临床特征,B.OD,B.CUTOFF,B.SCO,H.参考低值 || '～' || H.参考高值 As 参考" & vbNewLine & _
                   "         From (Select A.检验项目ID," & vbNewLine & _
                   "                      Min(Decode(E.报告项目ID,A.检验项目ID,F.编码,99999)) As 检验项目," & vbNewLine & _
                   "                      Max(Decode(E.报告项目ID,A.检验项目ID,E.排列序号,F.编码)) As 排列序号" & vbNewLine & _
                   "                 From 检验普通结果 A, 检验报告项目 E, 诊疗项目目录 F, 检验标本记录 G," & vbNewLine & _
                   "                      (Select Distinct C.诊疗项目ID From 检验项目分布 B,病人医嘱记录 C,检验标本记录 D Where D.合并ID=[1] And B.医嘱ID=C.相关ID And D.ID=B.标本ID) D" & vbNewLine & _
                   "                Where E.诊疗项目ID=D.诊疗项目ID And A.检验项目ID=E.报告项目ID(+) And E.诊疗项目ID=F.ID(+)" & vbNewLine & _
                   "                  And G.合并ID=[1] And G.ID=A.检验标本ID Group By A.检验项目ID" & vbNewLine & _
                   "                Order By Min(Decode(E.报告项目ID,A.检验项目ID,F.编码, 99999))," & vbNewLine & _
                   "                         Max(Decode(E.报告项目ID,A.检验项目ID,E.排列序号,F.编码))) A," & vbNewLine & _
                   "              检验普通结果 B,诊治所见项目 C,检验项目 D,检验标本记录 E,病人信息 F,检验仪器项目 G,检验项目参考 H" & vbNewLine & _
                   "        Where B.检验项目ID=A.检验项目ID(+) And E.合并ID=[1] And b.检验项目ID=c.ID" & vbNewLine & _
                   "          And C.ID=D.诊治项目ID And B.检验标本ID=E.ID And E.病人ID=F.病人ID(+) And B.检验项目ID=G.项目ID(+)" & vbNewLine & _
                   "          And (G.仪器ID=E.仪器ID Or G.仪器ID Is Null Or E.仪器ID Is Null) And B.检验结果 is not Null" & vbNewLine & _
                   "          And D.诊治项目ID=H.项目ID Order By ID, 检验项目编码, 排列序号, 临床特征) A," & vbNewLine & _
                   "      (Select min(Rownum) As 内部序号,C.ID From 检验标本记录 A,检验项目分布 B,检验普通结果 C,检验项目参考 D" & vbNewLine & _
                   "        Where A.ID=[1] And A.ID=B.标本ID And B.标本ID=C.检验标本ID And B.项目ID=D.项目ID" & vbNewLine & _
                   "          And C.检验项目ID=D.项目ID And 检验结果 is not null Group By C.ID Order By C.ID) Y" & vbNewLine & _
                     "Where A.ID=Y.ID order by 排列序号, nvl(临床特征, '111'))"
            '            strSql = "select Rownum As 序号, 显示英文名,显示中文名,显示结果,显示标志,临床特征,结果参考,显示单位,中文名,英文名,检验结果,标志,单位,打印序号,排列序号,检验方法,显示OD,显示CUTOFF,显示SCO,标本id from Lis_检验结果_多参考_Old where 标本ID=[1]"
        Else
            strSql = "select rownum As 序号,显示英文名,显示中文名,显示结果,显示标志,临床特征,结果参考,显示单位,中文名,英文名," & vbNewLine & _
                   "       检验结果,标志,单位,1 As 打印序号,排列序号,'' As 检验方法,显示OD,显示CUTOFF,显示SCO" & vbNewLine & _
                   "  from (Select replace(英文名, '.', '') As 显示英文名," & vbNewLine & _
                   "               replace(中文名, '.', '') As 显示中文名," & vbNewLine & _
                   "               replace(OD, '.', '') As 显示OD," & vbNewLine & _
                   "               replace(CUTOFF, '.', '') As 显示CUTOFF," & vbNewLine & _
                   "               replace(SCO, '.', '') As 显示SCO," & vbNewLine & _
                   "               Decode(本次结果,'-','阴性(-)','+','阳性(+)','*','*.**',本次结果) As 显示结果," & vbNewLine & _
                   "               decode(原始结果, '*', '', 标志) As 显示标志,Replace(临床特征, '111', '') As 临床特征," & vbNewLine & _
                   "               replace(trim(substr(参考,1,decode(instr(' ' || 参考,Chr(13) || Chr(10)),0,30,instr(' ' ||decode(substr(to_char(参考),1,1)," & vbNewLine & _
                   "               '.','0' || To_Char(参考),To_Char(参考)),Chr(13) || Chr(10))))),Chr(13) || Chr(10),'') As 结果参考," & vbNewLine & _
                "               单位 As 显示单位,中文名,英文名,本次结果 As 检验结果,标志,单位,排列序号"
            strSql = strSql & " From (Select distinct A.检验项目ID,B.ID,A.检验项目 As 检验项目编码,A.排列序号,C.中文名,D.缩写 As 英文名,B.原始结果," & vbNewLine & _
                   "                      B.检验结果 As 本次结果,Decode(B.结果标志,3,'↑',2,'↓',1,'',4,'↑',5,'↓↓',6,'↑↑','') As 标志," & vbNewLine & _
                   "                      D.单位,'' As 临床特征,B.OD,B.CUTOFF,B.SCO," & vbNewLine & _
                   "                      Trim(Replace(Replace(' ' ||" & vbNewLine & _
                   "                      Zlgetreference(C.ID, E.标本类型, Decode(E.性别, '男', 1, '女', 2, 0), F.出生日期," & vbNewLine & _
                   "                                                          E.仪器ID, E.年龄), ' .', '0.'), '～.', '～0.')) As 参考" & vbNewLine & _
                   "        From (Select A.检验项目ID," & vbNewLine & _
                   "                     Min(Decode(E.报告项目ID,A.检验项目ID,F.编码,99999)) As 检验项目," & vbNewLine & _
                   "                     Max(Decode(E.报告项目ID,A.检验项目ID,E.排列序号,F.编码)) As 排列序号" & vbNewLine & _
                   "                From 检验普通结果 A,检验报告项目 E,诊疗项目目录 F," & vbNewLine & _
                   "                     (Select Distinct C.诊疗项目ID From 检验项目分布 B, 病人医嘱记录 C Where B.标本ID=[1] And B.医嘱ID=C.相关ID) D" & vbNewLine & _
                   "               Where E.诊疗项目ID=D.诊疗项目ID And A.检验项目ID=E.报告项目ID(+) And E.诊疗项目ID=F.ID(+) And A.检验标本ID=[1]" & vbNewLine & _
                   "               Group By A.检验项目ID Order By Min(Decode(E.报告项目ID,A.检验项目ID,F.编码,99999))," & vbNewLine & _
                   "                        Max(Decode(E.报告项目ID,A.检验项目ID,E.排列序号,F.编码))) A," & vbNewLine & _
                   "             检验普通结果 B,诊治所见项目 C,检验项目 D,检验标本记录 E,病人信息 F,检验仪器项目 G" & vbNewLine & _
                   "        Where B.检验项目ID=A.检验项目ID(+) And B.检验标本ID=[1] And B.检验项目ID=C.ID" & vbNewLine & _
                   "          And C.ID=D.诊治项目ID And B.检验标本ID=E.ID And E.病人ID=F.病人ID(+) And B.检验项目ID=G.项目ID(+)" & vbNewLine & _
                   "          And (G.仪器ID=E.仪器ID Or G.仪器ID Is Null Or E.仪器ID Is Null) And B.检验结果 is not Null" & vbNewLine & _
                "       Union All"
            strSql = strSql & " Select distinct A.检验项目ID,B.ID,A.检验项目 As 检验项目编码,A.排列序号,C.中文名,D.缩写 As 英文名,B.原始结果," & vbNewLine & _
                   "                       B.检验结果 As 本次结果,Decode(B.结果标志,3,'↑',2,'↓',1,'',4,'↑',5,'↓↓',6,'↑↑','') As 标志," & vbNewLine & _
                   "                       D.单位,'' As 临床特征,B.OD,B.CUTOFF,B.SCO," & vbNewLine & _
                   "                       Trim(Replace(Replace(' ' ||" & vbNewLine & _
                   "                       Zlgetreference(C.ID, E.标本类型, Decode(E.性别, '男', 1, '女', 2, 0), F.出生日期," & vbNewLine & _
                   "                                                          E.仪器ID, E.年龄), ' .', '0.'), '～.', '～0.')) As 参考" & vbNewLine & _
                   "         From (Select A.检验项目ID," & vbNewLine & _
                   "                      Min(Decode(E.报告项目ID,A.检验项目ID,F.编码,99999)) As 检验项目," & vbNewLine & _
                   "                      Max(Decode(E.报告项目ID,A.检验项目ID,E.排列序号,F.编码)) As 排列序号" & vbNewLine & _
                   "                 From 检验普通结果 A, 检验报告项目 E, 诊疗项目目录 F, 检验标本记录 G," & vbNewLine & _
                   "                      (Select Distinct C.诊疗项目ID From 检验项目分布 B,病人医嘱记录 C,检验标本记录 D Where D.合并ID=[1] And B.医嘱ID=C.相关ID And D.ID=B.标本ID) D" & vbNewLine & _
                   "                Where E.诊疗项目ID=D.诊疗项目ID And A.检验项目ID=E.报告项目ID(+) And E.诊疗项目ID=F.ID(+)" & vbNewLine & _
                   "                  And G.合并ID=[1] And G.ID=A.检验标本ID Group By A.检验项目ID" & vbNewLine & _
                   "                Order By Min(Decode(E.报告项目ID,A.检验项目ID,F.编码, 99999))," & vbNewLine & _
                   "                         Max(Decode(E.报告项目ID,A.检验项目ID,E.排列序号,F.编码))) A," & vbNewLine & _
                   "              检验普通结果 B,诊治所见项目 C,检验项目 D,检验标本记录 E,病人信息 F,检验仪器项目 G" & vbNewLine & _
                   "        Where B.检验项目ID=A.检验项目ID(+) And E.合并ID=[1] And b.检验项目ID=c.ID" & vbNewLine & _
                   "          And C.ID=D.诊治项目ID And B.检验标本ID=E.ID And E.病人ID=F.病人ID(+) And B.检验项目ID=G.项目ID(+)" & vbNewLine & _
                   "          And (G.仪器ID=E.仪器ID Or G.仪器ID Is Null Or E.仪器ID Is Null) And B.检验结果 is not Null" & vbNewLine & _
                   "          Order By ID, 检验项目编码, 排列序号, 临床特征) A" & vbNewLine & _
                   " order by 排列序号, nvl(临床特征, '111'))"
            '            strSql = "select Rownum As 序号, 显示英文名, 显示中文名, 显示结果, 显示标志, 临床特征, 结果参考, 显示单位, 中文名, 英文名,检验结果, 标志, 单位, 打印序号, 排列序号, 检验方法, 显示od, 显示cutoff, 显示sco, 标本id from Lis_检验结果_非多参考_Old where 标本ID=[1]"
        End If
        Set Get_LisResult = gobjDatabase.OpenSQLRecord(strSql, gstrSysName, lng标本id)
    End If


    Exit Function:
ErrH:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function Get_LisGermFreeResult(ByVal lng标本id As Long) As ADODB.Recordset
'******************************************************************************************************************
'功能：获取微生物阴性结果
'参数：
'返回：
'附注：
'******************************************************************************************************************

    Dim strSql As String

    On Error GoTo ErrH

    If gblnNewLis Then
        strSql = "select 培养时间,正常菌,未检出,补充描述 正常菌群,镜检设备,镜检检出,镜检未检出,阴性评语 from 检验报告细菌 where 标本ID=[1]"
        Set Get_LisGermFreeResult = gobjPublicHisCommLis.openSqlOtherDB(1, strSql, gstrSysName, lng标本id)
    Else
        strSql = "select B.ID,D.报告结果,B.中文名 AS 检验项目,A.检验结果 AS 检验结果,A.培养描述 As 结果描述," & vbNewLine & _
               "   D.备注 As 阴性评语,D.检验备注 As 备注,A.耐药机制" & vbNewLine & _
               "  from 检验普通结果 A, 检验细菌 B, 检验标本记录 D" & vbNewLine & _
               " where A.细菌ID = B.ID And A.记录类型 = '0' And A.ID Not In(select 细菌结果ID from 检验药敏结果)" & vbNewLine & _
               "   And D.ID = A.检验标本ID And D.ID = [1]"
        Set Get_LisGermFreeResult = gobjDatabase.OpenSQLRecord(strSql, gstrSysName, lng标本id)

    End If


    Exit Function:
ErrH:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function Get_LisAntisResult(ByVal lng标本id As Long, ByVal lng细菌ID As Long) As ADODB.Recordset
'******************************************************************************************************************
'功能：获取抗生素结果
'参数：
'返回：
'附注：
'******************************************************************************************************************

    Dim strSql As String

    On Error GoTo ErrH

    If gblnNewLis Then
        strSql = "Select 行号, 抗生素1, 结果1, 耐药性1,阳性评语, 抗生素2, 结果2, 耐药性2" & vbNewLine & _
               "  From (With 抗生素结果 As (Select Decode(Mod(序号, 2), 1, 1, 2) As 栏数, Ceil(序号 / 2) As 行号,细菌ID,抗生素名,结果,结果类型,阳性评语" & vbNewLine & _
               "                         From (Select Rownum As 序号,C.ID,C.中文名 || Decode(length(Nvl(C.英文名,'')),null,'','(' || C.英文名 || ')') 抗生素名," & vbNewLine & _
               "                                      B.结果,B.结果类型,B.药敏方法,B.复查次数,B.参考描述,B.药敏组ID,A.细菌ID,A.阳性评语" & vbNewLine & _
               "                                 from 检验报告细菌 A,检验报告药敏 B,检验药敏     C" & vbNewLine & _
               "                                where A.ID = B.结果ID And B.药敏ID = C.ID And A.标本ID = [1] And A.细菌ID = [2]))" & vbNewLine & _
               "         Select M.行号,M.抗生素名 As 抗生素1,M.结果 As 结果1,M.结果类型 As 耐药性1,阳性评语," & vbNewLine & _
               "                N.抗生素名 As 抗生素2,N.结果 As 结果2,N.结果类型 As 耐药性2" & vbNewLine & _
               "           From (Select 行号, 细菌ID, 抗生素名, 结果, 结果类型,阳性评语 from 抗生素结果 T Where 栏数 = 1) M," & vbNewLine & _
               "                (Select 行号, 细菌ID, 抗生素名, 结果, 结果类型 from 抗生素结果 T Where 栏数 = 2) N" & vbNewLine & _
               "          Where M.行号 = N.行号(+)" & vbNewLine & _
               "            And M.细菌ID = N.细菌ID(+)" & vbNewLine & _
               "          Order By M.行号)"
        Set Get_LisAntisResult = gobjPublicHisCommLis.openSqlOtherDB(1, strSql, gstrSysName, lng标本id, lng细菌ID)
    Else
        strSql = "Select 行号, 抗生素1, 结果1, 耐药性1, 阳性评语, 抗生素2, 结果2, 耐药性2" & vbNewLine & _
               "  From (With 抗生素结果 As (Select Decode(Mod(序号, 2), 1, 1, 2) As 栏数,Ceil(序号/2) As 行号,细菌ID,抗生素名,结果,结果类型,阳性评语" & vbNewLine & _
               "                         From (select Rownum As 序号,B.ID,B.中文名 || Decode(length(Nvl(B.英文名, '')),null,'','(' || B.英文名 || ')') 抗生素名," & vbNewLine & _
               "                                      A.结果,Decode(A.结果类型,'R','R-耐药','I','I-中介','S','S-敏感',A.结果类型) AS 结果类型," & vbNewLine & _
               "                                      Decode(A.药敏方法,1,'1-MIC',2,'2-DISK',3,'3-K-B','') As 药敏方法,0 As 复查次数,'' As 参考描述," & vbNewLine & _
               "                                      C.药敏组ID,C.细菌ID,D.备注 As 阳性评语" & vbNewLine & _
               "                                 from 检验药敏结果 A,检验用抗生素 B,检验普通结果 C,检验标本记录 D" & vbNewLine & _
               "                                Where A.抗生素ID = B.ID And C.ID = A.细菌结果ID And C.记录类型 = A.记录类型" & vbNewLine & _
               "                                  And C.检验标本ID = D.ID And C.检验标本ID = [1] And C.细菌ID = [2] And C.记录类型 = '0'))" & vbNewLine & _
               "         Select M.行号, M.抗生素名 As 抗生素1, M.结果 As 结果1, M.结果类型 As 耐药性1, 阳性评语," & vbNewLine & _
               "                N.抗生素名 As 抗生素2, N.结果 As 结果2, N.结果类型 As 耐药性2" & vbNewLine & _
               "           From (Select 行号, 细菌ID, 抗生素名, 结果, 结果类型, 阳性评语 from 抗生素结果 T Where 栏数 = 1) M," & vbNewLine & _
               "                (Select 行号, 细菌ID, 抗生素名, 结果, 结果类型 From 抗生素结果 T Where 栏数 = 2) N" & vbNewLine & _
               "          Where M.行号 = N.行号(+)" & vbNewLine & _
               "            And M.细菌ID = N.细菌ID(+)" & vbNewLine & _
               "          Order By M.行号)"
        Set Get_LisAntisResult = gobjDatabase.OpenSQLRecord(strSql, gstrSysName, lng标本id, lng细菌ID)
    End If



    Exit Function:
ErrH:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function Get_LisGermsResult(ByVal lng标本id As Long) As ADODB.Recordset
'******************************************************************************************************************
'功能：获取微生物阳性结果
'参数：
'返回：
'附注：
'******************************************************************************************************************

    Dim strSql As String

    On Error GoTo ErrH

    If gblnNewLis Then
        strSql = "select A.细菌ID,A.标本ID," & vbNewLine & _
               "       B.中文名   As 细菌名称," & vbNewLine & _
               "       A.检验结果 As 细菌结果," & vbNewLine & _
               "       A.耐药机制," & vbNewLine & _
               "       A.培养描述" & vbNewLine & _
               "  from 检验报告细菌 A, 检验细菌记录 B" & vbNewLine & _
               " where A.细菌ID = B.ID(+)" & vbNewLine & _
               "   And A.标本ID = [1]"
        Set Get_LisGermsResult = gobjPublicHisCommLis.openSqlOtherDB(1, strSql, gstrSysName, lng标本id)
    Else
        strSql = "select distinct A.细菌ID,D.ID As 标本ID," & vbNewLine & _
               "       B.中文名   As 细菌名称," & vbNewLine & _
               "       A.检验结果 As 细菌结果," & vbNewLine & _
               "       A.耐药机制," & vbNewLine & _
               "       A.培养描述" & vbNewLine & _
               "  from 检验普通结果 A, 检验细菌 B, 检验标本记录 D, 检验药敏结果 E" & vbNewLine & _
               " where A.细菌id = B.ID And A.记录类型 = '0' And D.ID = A.检验标本ID" & vbNewLine & _
               "   And A.ID = E.细菌结果ID And D.ID = [1]"
        Set Get_LisGermsResult = gobjDatabase.OpenSQLRecord(strSql, gstrSysName, lng标本id)
    End If



    Exit Function:
ErrH:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function Get_LisMultiLineRefer(ByVal lng标本id As Long) As ADODB.Recordset
'******************************************************************************************************************
'功能：判断申请是否包含多参考项目
'参数：
'返回：
'附注：
'******************************************************************************************************************

    Dim strSql As String

    On Error GoTo ErrH

    If gblnNewLis Then
        strSql = "Select count(*) As 临床特诊" & vbNewLine & _
               "  From 检验报告明细 A, 检验指标参考 B" & vbNewLine & _
               " Where A.项目ID = B.项目ID" & vbNewLine & _
               "   And Nvl(A.仪器ID,0) = Nvl(B.仪器ID,0)" & vbNewLine & _
               "   And B.临床特征 is not null" & vbNewLine & _
               "   And A.标本ID = [1]"
        Set Get_LisMultiLineRefer = gobjPublicHisCommLis.openSqlOtherDB(1, strSql, gstrSysName, lng标本id)
    Else
        strSql = "Select count(*) As 临床特诊" & vbNewLine & _
               "  From 检验普通结果 A, 检验项目参考 B" & vbNewLine & _
               " Where A.检验项目ID = B.项目ID" & vbNewLine & _
               "   And Nvl(A.仪器ID,0) = Nvl(B.仪器ID,0)" & vbNewLine & _
               "   And B.临床特征 is not null" & vbNewLine & _
               "   And A.检验标本ID = [1]"
        Set Get_LisMultiLineRefer = gobjDatabase.OpenSQLRecord(strSql, gstrSysName, lng标本id)
    End If



    Exit Function:
ErrH:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function Get_LisEnzymeMarker(ByVal lng标本id As Long) As ADODB.Recordset
'******************************************************************************************************************
'功能：判断申请是否包含酶标结果项目
'参数：
'返回：
'附注：
'******************************************************************************************************************

    Dim strSql As String

    On Error GoTo ErrH

    If gblnNewLis Then
        strSql = "select count(*) As 数量" & vbNewLine & _
               "  from 检验报告明细" & vbNewLine & _
               " where 标本ID = [1]" & vbNewLine & _
               "   And (OD is not null or cutoff is not null or SCO is not null)"
        Set Get_LisEnzymeMarker = gobjPublicHisCommLis.openSqlOtherDB(1, strSql, gstrSysName, lng标本id)
    Else
        strSql = "select count(*) As 数量" & vbNewLine & _
               "  from 检验普通结果" & vbNewLine & _
               " where 检验标本ID = [1]" & vbNewLine & _
               "   And (OD is not null or cutoff is not null or SCO is not null)"
        Set Get_LisEnzymeMarker = gobjDatabase.OpenSQLRecord(strSql, gstrSysName, lng标本id)
    End If



    Exit Function:
ErrH:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function Get_LisGermsCount(ByVal lng标本id As Long) As ADODB.Recordset
'******************************************************************************************************************
'功能：判断是否为阳性报告
'参数：
'返回：
'附注：
'******************************************************************************************************************

    Dim strSql As String

    On Error GoTo ErrH

    If gblnNewLis Then
        strSql = "select Count(*) As 数量" & vbNewLine & _
               "  from 检验报告细菌 A, 检验报告药敏 B, 检验药敏 C" & vbNewLine & _
               " where A.ID = B.结果ID" & vbNewLine & _
               "   and B.药敏ID = C.ID" & vbNewLine & _
               "   and A.标本ID = [1]"
        Set Get_LisGermsCount = gobjPublicHisCommLis.openSqlOtherDB(1, strSql, gstrSysName, lng标本id)
    Else
        strSql = "select count(*) As 数量" & vbNewLine & _
               "  from 检验药敏结果 A, 检验用抗生素 B, 检验普通结果 C" & vbNewLine & _
               " Where A.抗生素ID = B.ID And C.ID = A.细菌结果ID" & vbNewLine & _
               "   And C.记录类型 = A.记录类型 And C.检验标本ID = [1]" & vbNewLine & _
               "   And C.记录类型 = '0'"
        Set Get_LisGermsCount = gobjDatabase.OpenSQLRecord(strSql, gstrSysName, lng标本id)

    End If


    Exit Function:
ErrH:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function Get_LisAdvice(ByVal lng标本id As Long) As ADODB.Recordset
'******************************************************************************************************************
'功能：获取检验标本自定义评语
'参数：
'返回：
'附注：
'******************************************************************************************************************

    Dim strSql As String
    Dim rsTmp As ADODB.Recordset

    On Error GoTo ErrH
    strSql = "select VIEW_NAME from all_views where upper(owner)='ZLHIS' And upper(view_name)='LIS_V_自定义结果评语'"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, gstrSysName)
    If Not ChkRsState(rsTmp) Then
        strSql = "select * from " & rsTmp!VIEW_NAME & " where 标本ID=[1]"
        Set Get_LisAdvice = gobjDatabase.OpenSQLRecord(strSql, gstrSysName, lng标本id)
    End If

    Exit Function:
ErrH:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function Get_LisVersion() As ADODB.Recordset
'******************************************************************************************************************
'功能：获取检验系统版本
'参数：
'返回：
'附注：
'******************************************************************************************************************

    Dim strSql As String
    Dim rsTmp As New ADODB.Recordset
    Dim rsLis As ADODB.Recordset
    Dim rsBh As ADODB.Recordset
    On Error GoTo ErrH
    strSql = "select  Nvl(A.核收时间1, To_Date('1900-01-01', 'YYYY-MM-DD HH24:MI:SS')) As 核收时间1 from (select Max(核收时间) As 核收时间1 from 检验报告记录) a"
    Set rsBh = gobjPublicHisCommLis.openSqlOtherDB(1, strSql, gstrSysName)
    strSql = "select  Nvl(B.核收时间2, To_Date('1900-01-01', 'YYYY-MM-DD HH24:MI:SS')) As 核收时间2 from (select Max(核收时间) As 核收时间2 from 检验标本记录) B"
    Set rsLis = gobjDatabase.OpenSQLRecord(strSql, gstrSysName)
    
    rsTmp.Fields.Append "核收时间1", adVarChar, 100
    rsTmp.Fields.Append "核收时间2", adVarChar, 100
    If rsTmp.State = adStateClosed Then rsTmp.Open
    rsTmp.AddNew
    If rsBh.RecordCount > 0 Then
        rsTmp("核收时间1") = rsBh("核收时间1")
    Else
        rsTmp("核收时间1") = rsBh("1900-01-01")
    End If
    If rsLis.RecordCount > 0 Then
        rsTmp("核收时间2") = rsLis("核收时间2")
    Else
        rsTmp("核收时间2") = rsBh("1900-01-01")
    End If

    rsTmp.Update

    Set Get_LisVersion = rsTmp
    
    
'    strSql = "select Nvl(A.核收时间1, To_Date('1900-01-01', 'YYYY-MM-DD HH24:MI:SS')) As 核收时间1," & vbNewLine & _
'           "       Nvl(B.核收时间2, To_Date('1900-01-01', 'YYYY-MM-DD HH24:MI:SS')) As 核收时间2" & vbNewLine & _
'           "  from (select Max(核收时间) As 核收时间1 from 检验报告记录) A," & vbNewLine & _
'           "       (select Max(核收时间) As 核收时间2 from 检验标本记录) B"
'
'    Set Get_LisVersion = gobjDatabase.OpenSQLRecord(strSql, gstrSysName)

    Exit Function:
ErrH:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function Get_LisSignPicture(ByVal lng人员ID As Long) As ADODB.Recordset
'******************************************************************************************************************
'功能：获取签名图片
'参数：
'返回：
'附注：
'******************************************************************************************************************

    Dim strSql As String
    Dim rsTmp As ADODB.Recordset

    On Error GoTo ErrH

    strSql = "select 签名图片 from 人员表 where ID=[1]"
    Set Get_LisSignPicture = gobjDatabase.OpenSQLRecord(strSql, gstrSysName, lng人员ID)

    Exit Function:
ErrH:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function Get_LisSignId(ByVal str审核人 As String) As ADODB.Recordset
'******************************************************************************************************************
'功能：获取签名图片
'参数：
'返回：
'附注：
'******************************************************************************************************************

    Dim strSql As String

    On Error GoTo ErrH

    strSql = "select ID from 人员表 where 姓名=[1]"

    Set Get_LisSignId = gobjDatabase.OpenSQLRecord(strSql, gstrSysName, Nvl(str审核人))

    Exit Function:
ErrH:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function Get_LisSampleSort() As String
'******************************************************************************************************************
'功能：获取标本顺序
'参数：
'返回：
'附注：
'******************************************************************************************************************

    Dim strSql As String
    Dim rsTmp As ADODB.Recordset

    On Error GoTo ErrH

    strSql = "select f_List2str(Cast(Collect(名称||'#'||顺序) As t_Strlist)) As 顺序 from (select * from 标本顺序 order by 顺序)"

    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, gstrSysName)
    If Not ChkRsState(rsTmp) Then Get_LisSampleSort = Nvl(rsTmp!顺序)
    Exit Function:
ErrH:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function Get_LisTypeSort() As String
'******************************************************************************************************************
'功能：获取类别顺序
'参数：
'返回：
'附注：
'******************************************************************************************************************

    Dim strSql As String
    Dim rsTmp As ADODB.Recordset

    On Error GoTo ErrH

    strSql = "select f_List2str(Cast(Collect(名称||'#'||顺序) As t_Strlist)) As 顺序 from (select * from 类别顺序 order by 顺序)"

    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, gstrSysName)

    If Not ChkRsState(rsTmp) Then Get_LisTypeSort = Nvl(rsTmp!顺序)
    Exit Function:
ErrH:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function Get_LisNameSort() As String
'******************************************************************************************************************
'功能：获取名称顺序
'参数：
'返回：
'附注：
'******************************************************************************************************************

    Dim strSql As String
    Dim rsTmp As ADODB.Recordset

    On Error GoTo ErrH

    strSql = "select f_List2str(Cast(Collect(名称||'#'||顺序) As t_Strlist)) As 顺序 from (select * from 项目顺序 order by 顺序)"

    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, gstrSysName)

    If Not ChkRsState(rsTmp) Then Get_LisNameSort = Nvl(rsTmp!顺序)
    Exit Function:
ErrH:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function IsNewVersion() As Boolean
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim strVersion As String
    strSql = "select 版本号 from zlsystems where 编号=100"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, gstrSysName)
    strVersion = rsTmp!版本号
    If Split(strVersion, ".")(1) >= 35 And Split(strVersion, ".")(2) > 0 Then
        IsNewVersion = True
    Else
        IsNewVersion = False
    End If
End Function

Public Function SetParameter(ByRef rs As ADODB.Recordset, ByVal strConditionName As String, ByVal strConditionValue As String, Optional ByVal strConditionType As String = "文本") As Boolean

    rs.Filter = ""
    rs.Filter = "参数名称='" & strConditionName & "'"
    If rs.RecordCount = 0 Then rs.AddNew
    rs("参数名称").Value = strConditionName
    rs("参数结果").Value = strConditionValue
    rs("参数类型").Value = strConditionType

    SetParameter = True

End Function

Public Function CreateParameter() As ADODB.Recordset

    Dim rs As New ADODB.Recordset

    With rs
        .Fields.Append "参数名称", adVarChar, 50
        .Fields.Append "参数结果", adVarChar, 4000
        .Fields.Append "参数类型", adVarChar, 30
        .Open
    End With

    Set CreateParameter = rs

End Function

Public Function GetParameter(ByRef rs As ADODB.Recordset, ByVal strParameterName As String) As String

    rs.Filter = ""
    rs.Filter = "参数名称='" & strParameterName & "'"
    If rs.RecordCount > 0 Then
        GetParameter = CStr(rs("参数结果").Value)
    End If

End Function
