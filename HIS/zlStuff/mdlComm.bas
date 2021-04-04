Attribute VB_Name = "mdlComm"
Option Explicit
Public Enum gEditType
     g新增 = 0
     g修改 = 1
     g审核 = 2
     g取消 = 3
     g查看 = 4
End Enum
Public Enum RecBillStatus  '记录状态信息
    正常记录 = 1
    冲销记录 = 2
    被冲销记录 = 3
End Enum
Public Enum ErrBillStatusInfor  '单据状态信息
    正常情况 = 1
    已经删除
    已经审核
    已经冲销
End Enum
Public Enum gRegType
    g注册信息 = 0
    g公共全局 = 1
    g公共模块 = 2
    g私有全局 = 3
    g私有模块 = 4
End Enum

Public gstrProductName As String

Public gblnCode As Boolean                   '是否具有卫材条码管理权限

Public Enum g小数类型
    g_数量 = 0
    g_成本价
    g_售价
    g_金额
End Enum

Private Type m_小数位
    数量小数 As Integer
    成本价小数 As Integer
    零售价小数 As Integer
    金额小数 As Integer
End Type

Public Type m_单位小数
     obj_散装小数 As m_小数位
     obj_包装小数 As m_小数位
     obj_最大小数 As m_小数位
End Type

Public g_小数位数 As m_单位小数

'小数格式化串
Public Type g_FmtString
    FM_数量 As String
    FM_成本价 As String
    FM_零售价 As String
    FM_金额 As String
    FM_散装零售价 As String
End Type
Private Type mSystem_para
    int简码方式 As Integer
    Para_输入方式   As String
    para_卫材填单下可用库存 As Boolean
    bln存在站点 As Boolean      '是否存在站点管理
    bln就诊卡密文显示 As Boolean  'true,密文显示,false 显示刷卡的卡号
    str就诊卡前缀符 As String    ' 存放就诊卡号中的字母前缀,不同前缀用|分隔,如:AA|BB|CC...
    P156_出库算法 As Integer    '0-批次优先;1-效期优先
End Type
Public gSystem_Para As mSystem_para


'小数位数设置
Public Const GFM_VBXS As String = "###0.000;-###0.000;0.000; "    '换算系数
Public Const GFM_VBCJL  As String = "#####0.00000;-#####0.00000;0.00000;"    '指导差价率
Public Const GFM_VBKL  As String = "#####0.0000;-#####0.0000;0.0000;"    '扣率
Public Const GFM_VBJCL  As String = "#####0.00;-#####0.00;0.00;"    '加成率


Public Const GFM_XS As String = "'999999999990.999'"    '换算系数
Public Const GFM_CJL  As String = "'999999999990.99999'"    '指导差价率
Public Const GFM_KL  As String = "'999999999990.99999'"    '扣率
Public Const GFM_JCL  As String = "'999999999990.99'"    '加成率

'卫材库存查询中，各批次报警的字体颜色常数
Public Const glng报警 As Long = &HC00000
Public Const glng正常 As Long = &H80000008
Public Const glng停用 As Long = &HC0

Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long


Public Function ExistsColObject(Col, index) As Boolean
    '判断集合中是否存在指定索引(关键字)的成员
    On Error GoTo ErrorHandler
    
    Dim v As Variant
    
    If TypeName(Col(index)) = "Collection" Then
        '索引对应的成员是集合时
        ExistsColObject = True
        Exit Function
    Else
        '索引对应的成员是非集合时
        v = Col(index)
        ExistsColObject = True
        Exit Function
    End If
ErrorHandler:
    '异常时表示无索引对应的成员
    ExistsColObject = False
End Function
Public Function GetCodePrivs() As Boolean
    '判读卫材系统是否具有条码管理的权限
    Dim rsTemp As ADODB.Recordset
    On Error GoTo ErrHandle
    gstrSQL = "Select 1 From Zltools.zlRegFunc Where 系统 = 1 And 序号 = 1711 And 功能 = '设置条码管理'"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, GetCodePrivs)
    GetCodePrivs = Not rsTemp.EOF
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

'Public Function getDept(strDept As String, Optional ByRef strID As String = "", Optional ByRef strName As String = "") As Boolean
'    Dim rsTemp As New ADODB.Recordset, strSQL As String
'    getDept = True
'    If strDept <> "" Then
'        If IsNumeric(strDept) Then
'            strSQL = "Select * From 供应商 Where 末级=1 And 编码=[1]"
'        Else
'            strSQL = "Select * From 供应商 Where 末级=1 And (简码=[1] Or 名称=[1])"
'        End If
'
'        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "供应商检查", UCase(strDept))
'
'        If rsTemp.EOF Then
'            getDept = False
'        Else
'            strID = rsTemp!Id
'            strName = rsTemp!名称
'        End If
'    End If
'End Function
Public Function GetMatchingSting(ByVal strString As String, Optional blnUpper As Boolean = True) As String
    '--------------------------------------------------------------------------------------------------------------------------------------
    '功能:加入匹配串%
    '参数:strString 需匹配的字串
    '     blnUpper-是否转换在大写
    '返回:返回加匹配串%dd%
    '--------------------------------------------------------------------------------------------------------------------------------------
    Dim strLeft As String
    Dim strRight As String
    
    If gstrMatchMethod = "0" Then
        strLeft = "%"
        strRight = "%"
    Else
        strLeft = ""
        strRight = "%"
    End If
    If blnUpper Then
        GetMatchingSting = strLeft & UCase(strString) & strRight
    Else
        GetMatchingSting = strLeft & strString & strRight
    End If
End Function

Public Sub SetCtlBackColor(objCtl As Object)
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:设置控件的背景色
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    If objCtl.Enabled Then
        objCtl.BackColor = &H80000005
    Else
        objCtl.BackColor = &H8000000A
    End If
End Sub

'取指定列头的列位置
Public Function GetCol(mshFlex As Object, ByVal ColName As String) As Integer
    Dim i As Integer
    
    On Error GoTo errH
    GetCol = -1
    With mshFlex
        For i = 0 To .Cols - 1
            If .TextMatrix(0, i) = ColName Then
                GetCol = i
                Exit Function
            End If
        Next
        
    End With
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function 检查单价(ByVal lng单据 As Long, ByVal strNo As String, Optional blnMsg As Boolean = True) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim lng材料_Last As Long, lng材料_Cur As Long
    
    '检查卫材的价格是否为最新的价格，允许继续操作
    '由于在保存前判断很麻烦，且各种单据的表格中保存的数据不一样，因此，待保存完成之后且提交前对已保存的数据进行检查
    '卫材相同的记录略过
    On Error GoTo ErrHandle

    gstrSQL = " Select '售价' As 类型, a.序号, a.药品id As 材料id, Nvl(a.批次, 0) As 批次, b.现价" & _
            " From 药品收发记录 A," & _
                 " (Select 收费细目id, Nvl(现价, 0) 现价, 执行日期" & _
                   " From 收费价目" & _
                   " Where (终止日期 Is Null Or Sysdate Between 执行日期 And Nvl(终止日期, To_Date('3000-01-01', 'yyyy-MM-dd')))" & _
                   GetPriceClassString("") & ") B, 收费项目目录 C" & _
            " Where a.单据 = [2] And a.No = [1] And a.药品id = b.收费细目id And c.Id = b.收费细目id And Round(a.零售价," & g_小数位数.obj_散装小数.零售价小数 & ") <> Round(b.现价, " & g_小数位数.obj_散装小数.零售价小数 & ") And" & _
              "    NVL(c.是否变价, 0) = 0" & _
            " Union All" & _
            " Select '售价' As 类型, a.序号, a.药品id As 材料id, Nvl(a.批次, 0) As 批次, decode(nvl(b.批次,0),0,b.实际金额 / b.实际数量,b.零售价) As 现价" & _
            " From 药品收发记录 A, 药品库存 B, 收费项目目录 C" & _
            " Where a.单据 = [2] And a.No = [1] And c.Id = a.药品id And Round(a.零售价," & g_小数位数.obj_散装小数.零售价小数 & ") <> Round(decode(nvl(b.批次,0),0,b.实际金额 / b.实际数量,b.零售价), " & g_小数位数.obj_散装小数.零售价小数 & ") And Nvl(c.是否变价, 0) = 1 And" & _
                  " b.性质 = 1 And b.库房id = a.库房id And b.药品id = a.药品id And NVL(b.批次, 0) = NVL(a.批次, 0) And NVL(b.实际数量, 0) <> 0 And a.入出系数 = -1" & _
            " Union All" & _
            " Select '成本价' As 类型, a.序号, a.药品id As 材料id, Nvl(a.批次, 0) As 批次, b.平均成本价 As 现价" & _
            " From 药品收发记录 A, 药品库存 B" & _
            " Where a.单据 = [2] And a.No = [1] And a.药品id = b.药品id And Nvl(a.批次, 0) = Nvl(b.批次, 0) and round(a.成本价," & g_小数位数.obj_散装小数.成本价小数 & ")<>round(b.平均成本价," & g_小数位数.obj_散装小数.成本价小数 & ") And a.库房id = b.库房id and a.入出系数=-1 and b.性质=1" & _
            " Order By 类型, 材料id, 序号"

    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "检查当前的价格", strNo, lng单据)
      
    If rsTemp.EOF Then
        检查单价 = True
        Exit Function
    End If
    
    lng材料_Last = 0
    With rsTemp
        Do While Not .EOF
            lng材料_Cur = !材料ID
            If lng材料_Cur <> lng材料_Last Then
                If blnMsg = True Then
                    If MsgBox("第" & !序号 & "行卫材的" & !类型 & "不是最新价格，是否继续保存单据？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                Else
                    Exit Function
                End If
            End If
            
            lng材料_Last = lng材料_Cur
            .MoveNext
        Loop
        检查单价 = True
    End With
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function 材料单据审核(ByVal str填制人 As String) As Boolean
    
    '材料单据审核时，是否判断审核人与填制人，其返回审核结果
    Dim blnBillVerify As Boolean
    
    材料单据审核 = True
    
    '暂无此功能,原因是卫生材料没有药品控制那么严.先预留此参数
    blnBillVerify = Val(zldatabase.GetPara(64, glngSys, 0)) = 1
    If Not blnBillVerify Then Exit Function
    
    材料单据审核 = (Trim(str填制人) <> Trim(UserInfo.用户名))
    If Not 材料单据审核 Then MsgBox "填制人与审核人不能是同一人，请检查！", vbInformation, gstrSysName
End Function

Public Function 检查库存数据(ByVal lng库房ID As Long, ByVal lng材料ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim bln库存是否分批 As Boolean, bln分批 As Boolean, bln库房 As Boolean
    
    '通过材料选择器输入卫材时，如果卫材库存中的数据与从部门性质、卫材目录中的分批属性判断出的不一致，则报错
    
    检查库存数据 = False
    On Error GoTo ErrHandle
    
    '如果没有库存记录，则直接退出
    gstrSQL = "" & _
        "   Select Count(*) 记录数 From 药品库存 " & _
        "   Where 库房ID=[1] And 性质=1 And 药品ID=[2]"
    
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "检查库存数据是否存在", lng库房ID, lng材料ID)
    If rsTemp!记录数 = 0 Then
        检查库存数据 = True
        Exit Function
    End If
    
    
    '存在分批记录则表明分批
    gstrSQL = " Select Count(*) 分批 From 药品库存 " & _
              " Where 库房ID=[1] And 性质=1 And Nvl(批次,0)<>0 And 药品ID=[2]"
              
    
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "检查库存数据是否存在", lng库房ID, lng材料ID)
    
    bln库存是否分批 = (rsTemp!分批 <> 0)
    
    '先判断是否是库房
    gstrSQL = "select 部门ID from 部门性质说明 where (工作性质 like '发料部门' Or 工作性质 like '%制剂室') And 部门id=[1]"
    
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "取部门性质", lng库房ID)
    bln库房 = (rsTemp.EOF)
        
    '判断对应的卫材目录中的分批属性
    gstrSQL = "" & _
        "   Select Nvl(库房分批,0) as 库房分批,nvl(在用分批,0) 在用分批 " & _
        "   From 材料特性 Where 材料ID=[1]"
    
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "取卫材目录中的分批属性", lng材料ID)
    
    If bln库房 Then
        bln分批 = (rsTemp!库房分批 = 1)
    Else
        bln分批 = (rsTemp!在用分批 = 1)
    End If
    检查库存数据 = (bln库存是否分批 = bln分批)
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub RefreshRowNO(ByRef mshBill As Object, ByVal lng序号列 As Long, Optional ByVal lngRow As Long = 1)
    Dim lngRows As Long
    '从指定行开始更新序号
    
    With mshBill
        lngRows = .Rows - 1
        For lngRow = lngRow To lngRows
            .TextMatrix(lngRow, lng序号列) = lngRow
        Next
    End With
End Sub

Public Sub CheckLapse(ByVal str效期 As String)
    '失效药品检查
    If Not IsDate(str效期) Then Exit Sub
    If Format(str效期, "yyyy-MM-dd") < Format(sys.Currentdate, "yyyy-MM-dd") Then
        MsgBox "该卫生材料已经失效了！", vbInformation, gstrSysName
    End If
End Sub

'转换数值为日期
Public Function TranNumToDate(ByVal strNum As String, Optional ByVal blnDec As Boolean = False) As String
    Dim strYear As String
    Dim strMonth As String
    Dim strDay As String
    Dim strDate As String
    
    TranNumToDate = ""
    strYear = Mid(strNum, 1, 4)
    strMonth = Mid(strNum, 5, 2)
    strDay = Mid(strNum, 7, 2)
        
    If strYear < 1000 Or strYear > 5000 Then Exit Function
    If strMonth = "" Then strMonth = "01"
    If strDay = "" Then strDay = "01"
    
    If strMonth > 12 Or strMonth < 1 Then Exit Function
    strDate = strYear & "-" & strMonth & "-" & strDay
        
    If Not IsDate(strDate) Then Exit Function
    
    strDate = Format(strDate, "yyyy-mm-dd")
    If blnDec Then strDate = DateAdd("d", -1, Format(strDate, "yyyy-mm-dd"))
    TranNumToDate = strDate
End Function

Public Function 相同符号(ByVal sinFirst As Single, ByVal sinSecond As Single) As Boolean
    Dim blnFirst_负数 As Boolean, blnSecond_负数 As Boolean
    相同符号 = False
    
    blnFirst_负数 = (sinFirst <= 0)
    blnSecond_负数 = (sinSecond <= 0)
    
    相同符号 = (blnFirst_负数 = blnSecond_负数)
End Function

'表示卫材出库时是否进行库存检查：0-不检查;1-检查，不足提醒；2-检查，不足禁止
Public Function Get出库检查(ByVal lng库房ID As Long) As Integer
    Dim rsSystemPara As New Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "select Nvl(检查方式,0) 出库检查 From 材料出库检查 Where 库房ID=[1]"
    
    Set rsSystemPara = zldatabase.OpenSQLRecord(gstrSQL, "出库检查", lng库房ID)
    
    If rsSystemPara.EOF Then
        Get出库检查 = 0
        Exit Function
    End If
    Get出库检查 = rsSystemPara!出库检查
    rsSystemPara.Close
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ExecuteSql(ByRef arrSQL As Variant, strTitle As String, _
Optional ByVal blnCommit As Boolean = True, Optional ByVal blnBeginTrans As Boolean = True) As Boolean
    Dim strTmp As Variant
    Dim i As Integer, j As Integer
    Dim intouter As Integer
    Dim intInner As Integer
    

    ExecuteSql = False
    If UBound(arrSQL) >= 0 Then
        '对SQL序列按药品ID升序排序
        intouter = UBound(arrSQL) - 1
        If Split(arrSQL(UBound(arrSQL)), ";" & vbCrLf)(0) = "出库" Then
            intouter = UBound(arrSQL) - 2
        Else
            intouter = UBound(arrSQL) - 1
        End If
        
        
        For i = 0 To intouter
            For j = i + 1 To intouter + 1
                If CLng(Split(arrSQL(j), ";" & vbCrLf)(0)) < CLng(Split(arrSQL(i), ";" & vbCrLf)(0)) Then
                    strTmp = CStr(arrSQL(j))
                    arrSQL(j) = arrSQL(i)
                    arrSQL(i) = strTmp
                End If
            Next
        Next
        
        '执行SQL语句
        On Error GoTo errH
        If blnBeginTrans Then gcnOracle.BeginTrans
        For i = 0 To UBound(arrSQL)
            zldatabase.ExecuteProcedure CStr(Split(arrSQL(i), ";" & vbCrLf)(1)), strTitle
        
'            Call SQLTest(App.ProductName, strTitle, CStr(Split(arrSql(i), ";" & vbCrLf)(1)))
'            gcnOracle.Execute CStr(Split(arrSql(i), ";" & vbCrLf)(1)), , adCmdStoredProc
'            Call SQLTest
        Next
        If blnCommit Then gcnOracle.CommitTrans
        ExecuteSql = True
    End If
    Exit Function
       
errH:
    If blnBeginTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ReturnSQL(ByVal lng库房ID As Long, ByVal strCaption As String, _
    Optional ByVal bln调拨 As Boolean = True, _
    Optional ByRef strOutSQL As String = "", _
    Optional ByVal lngModuleNO As Long = 0) As ADODB.Recordset
    '-----------------------------------------------------------------------------------------------------------
    '功能:返回相关的SQL数据
    '入参:
    '出参:strOutSQL-返回相关的SQL语句
    '返回:
    '编制:刘兴洪
    '日期:2008-08-22 17:26:47
    '-----------------------------------------------------------------------------------------------------------
        
    Dim str库房性质 As String, str材料流向 As String, str站点限制 As String
    '根据卫材流向控制表的数据，提取对方库房
    '-----------------调拨-----------------
    '所在库房是当前库房的，提取流向 In (1"可流向对方库房",3"可双向流通")
    '对方库房是当前库房的，提取流向 IN (2"可流向所在库房",3"可双向流通")
    '-----------------申领-----------------
    '所在库房是当前库房的，提取流向 In (2"可流向所在库房",3"可双向流通")
    '对方库房是当前库房的，提取流向 IN (1"可流向对方库房",3"可双向流通")
    Dim bln发料部门 As Boolean  '表示库房为发料部门
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    str站点限制 = GetDeptStationNode(lng库房ID)
    If lngModuleNO = 1716 Or lngModuleNO = 1717 Or lngModuleNO = 1722 Then
        str库房性质 = "('V','W','K')"
    Else
        str库房性质 = "('V','W','K','12')"
    End If
    
    str材料流向 = "" & _
    ",( Select 对方库房ID ID From 材料流向控制" & _
    "   Where 所在库房ID=[1] And 流向 In (" & IIf(bln调拨, 1, 2) & ",3)" & _
    "   Union" & _
    "   Select 所在库房ID ID From 材料流向控制" & _
    "   Where 对方库房ID=[1] And 流向 In (" & IIf(bln调拨, 2, 1) & ",3)) D "
    
    If bln调拨 Then
        '确定只是发料部门
        gstrSQL = " Select a.ID From 部门表 a,部门性质说明 B where A.id=B.部门ID and A.ID=[1] and b.工作性质  in ('卫材库','制剂室')"
        
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "获取不是发料部门的部门", lng库房ID)
        
        bln发料部门 = rsTemp.RecordCount = 0
        rsTemp.Close
        If bln发料部门 Then
            '如果只是发料部门特性的部门，则只能移到库房
            If lngModuleNO = 1716 Or lngModuleNO = 1722 Then    '移库管理、申领管理
                gstrSQL = "" & _
                    "   Select DISTINCT a.id,a.编码,a.名称,a.简码,a.位置 ,To_Char(a.建档时间, 'yyyy-mm-dd') As 建档时间, " & _
                    "          decode(To_Char(a.撤档时间, 'yyyy-mm-dd'),'3000-01-01','',To_Char(a.撤档时间, 'yyyy-mm-dd')) 撤档时间" & _
                    "   From 部门表 a,部门性质说明 B " & str材料流向 & _
                    "   where A.id=B.部门ID and a.id=d.id and b.工作性质  in ('卫材库','制剂室') " & _
                    "           AND (a.撤档时间 is null or a.撤档时间>= to_date('3000-01-01','yyyy-mm-dd'))"
            Else
                gstrSQL = "" & _
                    "   Select DISTINCT a.id,a.编码,a.名称,a.简码,a.位置 ,To_Char(a.建档时间, 'yyyy-mm-dd') As 建档时间, " & _
                    "          decode(To_Char(a.撤档时间, 'yyyy-mm-dd'),'3000-01-01','',To_Char(a.撤档时间, 'yyyy-mm-dd')) 撤档时间" & _
                    "   From 部门表 a,部门性质说明 B " & str材料流向 & _
                    "   where A.id=B.部门ID and a.id=d.id  and b.工作性质  in ('卫材库','制剂室','虚拟库房') " & _
                    IIf(str站点限制 <> "", " and (a.站点 = [2] or a.站点 is null) ", "") & "" & _
                    "           AND (a.撤档时间 is null or a.撤档时间>= to_date('3000-01-01','yyyy-mm-dd'))"
            End If

            strOutSQL = gstrSQL
            gstrSQL = gstrSQL & " Order by a.编码"
            Set ReturnSQL = zldatabase.OpenSQLRecord(gstrSQL, strCaption, lng库房ID, str站点限制)
            Exit Function
        End If
    End If

    '如果调拔对发料部门发限制.
    gstrSQL = "" & _
    " SELECT DISTINCT a.id,a.编码,a.名称,a.简码,a.位置 ,To_Char(a.建档时间, 'yyyy-mm-dd') As 建档时间, " & _
    "          decode(To_Char(a.撤档时间, 'yyyy-mm-dd'),'3000-01-01','',To_Char(a.撤档时间, 'yyyy-mm-dd')) 撤档时间" & _
    " FROM 部门性质说明 c, 部门性质分类 b, 部门表 a" & str材料流向 & _
    " Where c.工作性质 = b.名称 AND b.编码 in " & str库房性质 & _
    "       AND a.id = c.部门id And A.ID=D.ID" & _
    "       AND (a.撤档时间 is null or a.撤档时间>= to_date('3000-01-01','yyyy-mm-dd')) " & _
    IIf(str站点限制 <> "" And lngModuleNO <> 1716 And lngModuleNO <> 1722, " and (a.站点 = [2] or a.站点 is null) ", "")
    
    strOutSQL = Replace(gstrSQL, "[1]", lng库房ID)
    
    gstrSQL = gstrSQL & " Order by a.编码"
    Set ReturnSQL = zldatabase.OpenSQLRecord(gstrSQL, strCaption, lng库房ID, str站点限制)
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetBillInfo(ByVal lng单据 As Long, ByVal strNo As String, Optional ByVal bln填制日期 As Boolean = True, Optional ByVal bln配药日期 As Boolean = False) As String
    Dim rsTemp As New ADODB.Recordset
    '获取单据的最大修改时间
    
    On Error GoTo ErrHandle
    gstrSQL = "" & _
        "   Select to_char(Max(" & IIf(bln填制日期, "填制日期", IIf(bln配药日期, "配药日期", "审核日期")) & "),'yyyyMMddhh24miss') 日期 " & _
        "   From 药品收发记录 " & _
        "   Where 单据=[1] And NO=[2]"
    
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "获取单据的最大修改时间", lng单据, strNo)
    
    With rsTemp
        '返回空，表示已经删除
        If .EOF Then Exit Function
        If IsNull(!日期) Then Exit Function
        GetBillInfo = !日期
    End With
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function AutoAdd生产商(str编码 As String, str名称 As String, Optional strTittle As String = "增加生产商", Optional blnMsg As Boolean = False) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:自动增加生产商
    '--入参数:
    '--出参数:
    '--返  回:增加成功,返回true,否则返回false
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New Recordset
    Dim int编码 As Integer, strCode As String, strSpecify As String
    
    AutoAdd生产商 = False
    If blnMsg = True Then
        If MsgBox("没有找到你输入的材料生产商，你要把它加入材料生产商中吗？", vbYesNo + vbQuestion, strTittle) = vbNo Then
            Exit Function
        End If
    End If
    err = 0
    On Error GoTo ErrHand:
    
    gstrSQL = "SELECT Nvl(MAX(LENGTH(编码)),2) As Length FROM 材料生产商"
    zldatabase.OpenRecordset rsTemp, gstrSQL, strTittle
    
    int编码 = rsTemp!Length
    
    gstrSQL = "SELECT Nvl(MAX(LPAD(编码," & int编码 & ",'0')),'00') As Code FROM 材料生产商"
    zldatabase.OpenRecordset rsTemp, gstrSQL, strTittle
    strCode = rsTemp!Code
    
    int编码 = Len(strCode)
    strCode = strCode + 1
    
    If int编码 >= Len(strCode) Then
    strCode = String(int编码 - Len(strCode), "0") & strCode
    End If
    strSpecify = zlStr.GetCodeByVB(str名称)
    
    
    gstrSQL = "ZL_材料生产商_INSERT('" & strCode & "','" & str名称 & "','" & strSpecify & "')"
    Call zldatabase.ExecuteProcedure(gstrSQL, strTittle)
    str编码 = strCode
    AutoAdd生产商 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function Where撤档时间(Optional strAlias As String) As String
    If strAlias = "" Then
        Where撤档时间 = " (撤档时间=to_date('3000-01-01','yyyy-mm-dd') or 撤档时间 is null) "
    Else
        Where撤档时间 = " (" & strAlias & ".撤档时间=to_date('3000-01-01','yyyy-mm-dd') or " & strAlias & ".撤档时间 is null) "
    End If
End Function

'取时价卫生材料入库时，是否必须输入加价率
Public Function Get加价率() As Boolean
    Get加价率 = Val(zldatabase.GetPara(82, glngSys, 0)) = 1
End Function

Public Function Get定价单位() As Integer
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:取指导批发价的定价单位
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
   '20050106刘兴宏加入此参数
    Get定价单位 = Val(zldatabase.GetPara(88, glngSys, 0))
End Function

Public Function GetDigit() As Integer
    '获取金额小数位数
    Dim int小数 As Integer
    Dim rsTemp As ADODB.Recordset

    On Error GoTo ErrHandle
    gstrSQL = "Select nvl(精度,2) as 精度  From 药品卫材精度 Where 性质=0 and 类别 = 2 And 内容 = 4 And 单位 = 5"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "查询金额精度")
    If rsTemp.RecordCount = 0 Then
        GetDigit = 2
    Else
        GetDigit = rsTemp!精度
    End If
    Exit Function
ErrHandle:
    GetDigit = 2
End Function

Public Function GetDigit发料() As Integer
    '获取金额小数位数
    Dim int小数 As Integer
    
    Dim rsTemp As New ADODB.Recordset

    On Error GoTo ErrHandle
    int小数 = Trim(zldatabase.GetPara(9, glngSys, 0))
    If int小数 = 0 Then int小数 = 2
    GetDigit发料 = Val(int小数)
    Exit Function
ErrHandle:
    GetDigit发料 = 2
End Function

Public Function IS批次申领() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:是否按批次进行申领
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    IS批次申领 = Val(zldatabase.GetPara(83, glngSys, 0)) = 1
End Function

Public Function IS批次移库() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:是否按批次进行移库
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    IS批次移库 = Val(zldatabase.GetPara(280, glngSys, 0)) = 1
End Function

Public Function is时价卫材取上次售价() As Boolean
    '------------------------------------------------------------------------------------------------------
    '功能:卫材外购入库取上次售价
    '入参:
    '出参:
    '返回:返回true-取上次售价,否则返回false-默认方式
    '------------------------------------------------------------------------------------------------------
    is时价卫材取上次售价 = Val(zldatabase.GetPara(229, glngSys, 0)) = 1
End Function

Public Function Get分段加成率(ByVal dbl购价 As Double) As Double
    '功能:获取分段加成率
    '返回:返回加成率
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHandle
    gstrSQL = "Select 加成率 From 材料加成方案  " & _
             " Where  ([1] >最低价 And [1] <=最高价)  " & _
             "        Or ([1] <=最高价 And nvl(最低价,0)=0) " & _
             "        Or ([1] >最低价 And nvl(最高价,0)=0)"
                 
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "获取分段加成率", dbl购价)
             
    If rsTemp.EOF Then
        ShowMsgBox "未设置金额段为:" & dbl购价 & " 的加成率， " & vbCrLf & "请在卫材目录管理中设置，现以15%为加成率计算!"
        Get分段加成率 = 15
    Else
        Get分段加成率 = Val(zlStr.NVL(rsTemp!加成率))
    End If
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get分段加成售价(ByVal dbl采购价 As Double, ByVal dbl比例系数 As Double, ByVal strFormCaption As String, ByRef sng售价 As Double) As Boolean
    '------------------------------------------------------------------------------------------------------
    '功能:获取分断加成后的售价
    '入参:dbl采购价-采购价
    '     dbl比例系数-比例系数
    '     strFormCaption-窗体名称
    '出参:dbl售价-返回的分断加成后的售价
    '返回:计算的数据正确，返回true,否则返回false
    '修改人:刘兴宏
    '修改时间:2007/2/26
    '------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim byt计算方法 As Byte
    Dim dbl限价 As Double, dbl总限价 As Double
    Dim dblTemp As Double
    Dim dbl售价1 As Double
    Dim dbl成本价 As Double
    
    err = 0: On Error GoTo ErrHand:
    dbl成本价 = dbl采购价 / IIf(dbl比例系数 = 0, 1, dbl比例系数)
    
    gstrSQL = "Select 序号,最低价,最高价,1+加成率/100 as 加成率,计算方法,限价 from 材料加成方案 order by 序号"
    zldatabase.OpenRecordset rsTemp, gstrSQL, strFormCaption
    If rsTemp.EOF Then
        ShowMsgBox "未设置分断加成率,请在卫材目录管理中设置!"
        Exit Function
    End If
    byt计算方法 = Val(zlStr.NVL(rsTemp!计算方法))
    dbl总限价 = Val(zlStr.NVL(rsTemp!限价))
    dbl售价1 = 0
    dblTemp = 0
    
    '2010-8月底处理四川省要求，增加分段限价控制。问题号32282
    If rsTemp!序号 = 0 Then rsTemp.MoveNext
    If byt计算方法 = 0 Then
        '整体计算法
        With rsTemp
            Do While Not .EOF
                If (dbl成本价 > Val(zlStr.NVL(!最低价)) And dbl成本价 <= Val(zlStr.NVL(!最高价))) Or _
                   (dbl成本价 > Val(zlStr.NVL(!最低价)) And Val(zlStr.NVL(!最高价)) = 0) _
                Then
                    dbl限价 = Val(zlStr.NVL(!限价))
                    dblTemp = Val(zlStr.NVL(!加成率))
                    If dbl限价 > 0 Then
                        If Round((dblTemp - 1) * dbl成本价, 7) > Round(dbl限价, 7) Then
                            dbl售价1 = dbl成本价 + dbl限价
                        Else
                            dbl售价1 = dblTemp * dbl成本价
                        End If
                    Else
                        dbl售价1 = dblTemp * dbl成本价
                    End If
                    GoTo Check售价:
                End If
                .MoveNext
            Loop
        End With
        ShowMsgBox "未设置金额段为:" & dbl成本价 & " 的加成率， " & vbCrLf & "请在卫材目录管理中设置!"
        Exit Function
    End If
    '分断计算法
    '广东卫材加成算法:
    '    1000元以下，1000*10%；
    '    1000元以上，分两段，1000*10%+(采购价-1000)*8%
    With rsTemp
        Do While Not .EOF
            dbl限价 = Val(zlStr.NVL(rsTemp!限价))
            If dbl成本价 <= Val(zlStr.NVL(!最高价)) Or Val(zlStr.NVL(!最高价)) = 0 Then
                dblTemp = (dbl成本价 - Val(zlStr.NVL(!最低价))) * (Val(zlStr.NVL(!加成率)) - 1)
                If dbl限价 > 0 Then
                    If Round(dblTemp, 7) > Round(dbl限价, 7) Then
                        dbl售价1 = dbl成本价 + dbl售价1 + dbl限价
                    Else
                        dbl售价1 = dbl成本价 + dbl售价1 + dblTemp
                    End If
                Else
                    dbl售价1 = dbl成本价 + dbl售价1 + dblTemp
                End If
                GoTo Check售价:
            ElseIf dbl成本价 > Val(zlStr.NVL(!最高价)) Then
                dblTemp = (Val(zlStr.NVL(!最高价)) - Val(zlStr.NVL(!最低价))) * (Val(zlStr.NVL(!加成率)) - 1)
                If dbl限价 > 0 Then
                    If Round(dblTemp, 7) > Round(dbl限价, 7) Then
                        dbl售价1 = dbl售价1 + dbl限价
                    Else
                        dbl售价1 = dbl售价1 + dblTemp
                    End If
                Else
                    dbl售价1 = dbl售价1 + dblTemp
                End If
            End If
            .MoveNext
        Loop
    End With
    ShowMsgBox "未设置金额段为:" & dbl成本价 & " 的加成率， " & vbCrLf & "请在卫材目录管理中设置!"
    Exit Function
Check售价:
    If Round(dbl售价1 - dbl成本价, 7) > Round(dbl总限价, 7) Then
        ShowMsgBox "加成（￥" & Format(dbl售价1 - dbl成本价, "###0.0000000;-###0.0000000;0;0") _
                 & ")超过最高限价(￥" & Format(dbl总限价, "###0.0000000;-###0.0000000;0;0") _
                 & ")，将默认最高限价加成！"
        'Exit Function
        dbl售价1 = dbl成本价 + dbl总限价
    End If
    sng售价 = dbl售价1 * dbl比例系数
    
    Get分段加成售价 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function IS分段加成率() As Boolean
    '功能:检查是否加成率以分段金额来获取
    IS分段加成率 = Val(zldatabase.GetPara(121, glngSys, 0)) = 1
End Function

Public Function is时价卫材直接确定售价() As Boolean
    '------------------------------------------------------------------------------------------------------
    '功能:读取是否以is时价卫材直接确定售价的方式入库
    '入参:
    '出参:
    '返回:是以直接确定售价的方式,返回true,否则返回false
    '修改人:刘兴宏
    '修改时间:2007/1/25
    '------------------------------------------------------------------------------------------------------
    is时价卫材直接确定售价 = Val(zldatabase.GetPara(136, glngSys, 0)) = 1
End Function

Public Sub 初始小数位数()
    '------------------------------------------------------------------------------------------------------
    '功能:初始小数位数
    '入参:
    '出参:
    '返回:
    '修改人:刘兴宏
    '修改时间:2007/3/6
    '------------------------------------------------------------------------------------------------------
    Dim strSql As String
    Dim rsTemp As New ADODB.Recordset
    
    '    类别    Number(1)   1-药品,2-卫材
    '    内容    Number(1)   1-成本价，2-零售价,3-数量
    '    单位    Number(1)   1,2,3,4；药品分别为售价、门诊、住院、药库单位，卫材分别为散装、包装单位。
    '    精度    Number(1)   取值为2-4。
    On Error GoTo ErrHandle
    strSql = "Select * from 药品卫材精度 where 类别=[1] and 性质=0 order by 单位"
    Set rsTemp = zldatabase.OpenSQLRecord(strSql, "获取卫生材料的小数位数精度", 2)
    With g_小数位数
        With .obj_包装小数
            .成本价小数 = 7
            .零售价小数 = 7
            If glngModul = 1723 Then '卫材发料取费用金额精度，否则取药品卫材中设置的精度
                .金额小数 = GetDigit发料
            Else
                .金额小数 = GetDigit
            End If
            .数量小数 = 3
        End With
        With .obj_散装小数
            .成本价小数 = 7
            .零售价小数 = 7
            If glngModul = 1723 Then '卫材发料取费用金额精度，否则取药品卫材中设置的精度
                .金额小数 = GetDigit发料
            Else
                .金额小数 = GetDigit
            End If
            .数量小数 = 3
        End With
        With .obj_最大小数
            .成本价小数 = 7
            .金额小数 = 5
            .数量小数 = 5
            .零售价小数 = 7
        End With
    End With
    
    With gOraFmt_Max
        .FM_金额 = GetFmtString(-1, g小数类型.g_金额, True)
        .FM_成本价 = GetFmtString(-1, g小数类型.g_成本价, True)
        .FM_零售价 = GetFmtString(-1, g小数类型.g_售价, True)
        .FM_数量 = GetFmtString(-1, g小数类型.g_数量, True)
        .FM_散装零售价 = GetFmtString(-1, g小数类型.g_售价, True)
    End With
    
    If rsTemp.EOF Then Exit Sub
    Do While Not rsTemp.EOF
        If Val(zlStr.NVL(rsTemp!单位)) = 2 Then
            '包装单位
            If Val(zlStr.NVL(rsTemp!内容)) = 1 Then
                g_小数位数.obj_包装小数.成本价小数 = Val(zlStr.NVL(rsTemp!精度))
            ElseIf Val(zlStr.NVL(rsTemp!内容)) = 2 Then
                g_小数位数.obj_包装小数.零售价小数 = Val(zlStr.NVL(rsTemp!精度))
            ElseIf Val(zlStr.NVL(rsTemp!内容)) = 3 Then
                g_小数位数.obj_包装小数.数量小数 = Val(zlStr.NVL(rsTemp!精度))
            End If
        ElseIf Val(zlStr.NVL(rsTemp!单位)) = 1 Then
            '散装单位
            If Val(zlStr.NVL(rsTemp!内容)) = 1 Then
                g_小数位数.obj_散装小数.成本价小数 = Val(zlStr.NVL(rsTemp!精度))
            ElseIf Val(zlStr.NVL(rsTemp!内容)) = 3 Then
                g_小数位数.obj_散装小数.数量小数 = Val(zlStr.NVL(rsTemp!精度))
            Else
                g_小数位数.obj_散装小数.零售价小数 = Val(zlStr.NVL(rsTemp!精度))
            End If
        ElseIf Val(zlStr.NVL(rsTemp!单位)) = 5 Then
            '金额
            If glngModul <> 1723 Then '卫材发料的话 就不再用该精度，而是直接用费用业务设置的精度
                g_小数位数.obj_包装小数.金额小数 = Val(zlStr.NVL(rsTemp!精度))
                g_小数位数.obj_散装小数.金额小数 = Val(zlStr.NVL(rsTemp!精度))
            End If
        End If
        rsTemp.MoveNext
    Loop
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function GetFmtString(ByVal int单位 As Integer, ByVal 小数类型 As g小数类型, _
    Optional blnOracle As Boolean = False) As String
    '------------------------------------------------------------------------------------------------------
    '功能:返回指定的小数格式串
    '入参:int单位-0-散装单位,1-包装单位,<>0 or 1:数据库最大小数位
    '     lng小数位数-小数位数
    '     blnOracle-返回是oracle的格式串还是Vb的格式串
    '出参:
    '返回:返回指定的格式串
    '修改人:刘兴宏
    '修改时间:2007/3/6
    '------------------------------------------------------------------------------------------------------
    Dim strFmt As String
    Dim int位数 As Integer
    Select Case 小数类型
    Case g_数量
        Select Case int单位
        Case 0  ' 散装单位
            int位数 = g_小数位数.obj_散装小数.数量小数
        Case 1  '包装单位
            int位数 = g_小数位数.obj_包装小数.数量小数
        Case Else  '-1最大数据库单位
            int位数 = g_小数位数.obj_最大小数.数量小数
        End Select
    Case g_金额
        Select Case int单位
        Case 0  ' 散装单位
            int位数 = g_小数位数.obj_散装小数.金额小数
        Case 1  '包装单位
            int位数 = g_小数位数.obj_包装小数.金额小数
        Case Else  '-1最大数据库单位
            int位数 = g_小数位数.obj_最大小数.金额小数
        End Select
    Case g_成本价
        Select Case int单位
        Case 0  ' 散装单位
            int位数 = g_小数位数.obj_散装小数.成本价小数
        Case 1  '包装单位
            int位数 = g_小数位数.obj_包装小数.成本价小数
        Case Else  '-1最大数据库单位
            int位数 = g_小数位数.obj_最大小数.成本价小数
        End Select
    Case g_售价
        Select Case int单位
        Case 0  ' 散装单位
            int位数 = g_小数位数.obj_散装小数.零售价小数
        Case 1  '包装单位
            int位数 = g_小数位数.obj_包装小数.零售价小数
        Case Else  '-1最大数据库单位
            int位数 = g_小数位数.obj_最大小数.零售价小数
        End Select
    Case Else
        int位数 = 0
    End Select
    If blnOracle Then
       GetFmtString = "'9999999999990." & String(int位数, "9") & "'"
    Else
       GetFmtString = "#0." & String(int位数, "0")
    End If
End Function

Public Function InitSystemPara() As Boolean
    '------------------------------------------------------------------------------------------------------
    '功能:初始化相关的系统参数
    '入参:
    '出参:
    '返回:初始化成功,返回true,否则返回False
    '修改人:刘兴宏
    '修改时间:2007/6/28
    '------------------------------------------------------------------------------------------------------
    Dim strValue As String
    With gSystem_Para
        '0-拼音码,1-五笔码,2-两者
        .int简码方式 = Val(zldatabase.GetPara("简码方式"))
        '第1位1-全数字只查编码,第2位1-全字母只查简码,在HIS基础参数中设置
        .Para_输入方式 = zldatabase.GetPara(44, glngSys, 0): .Para_输入方式 = IIf(.Para_输入方式 = "", "11", .Para_输入方式)
        .para_卫材填单下可用库存 = Val(zldatabase.GetPara(95, glngSys, 0)) = 1
        '卡号显示方式
        .bln就诊卡密文显示 = Val(zldatabase.GetPara(12, glngSys)) = 1
        .str就诊卡前缀符 = zldatabase.GetPara(27, glngSys)
        .P156_出库算法 = zldatabase.GetPara(156, glngSys, 0, 0)
     End With
    '站点编号设置
    Call Init站点信息
    InitSystemPara = True
End Function
 
Public Function Check出院病人(ByVal strPrivs As String, ByVal lng单据 As Long, ByVal strNo As String, ByVal int记录性质 As Integer, ByVal int门诊标志 As Integer, Optional ByVal lng病人id As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:检查出院病人是否允许发料,需要根据权限控制(如果没有权限“发退出院病人处方”，则不允许发退料操作)
    '入参:
    '出参:
    '返回:允许,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-04-23 15:05:49
    '-----------------------------------------------------------------------------------------------------------

    '功能说明：如果当前病人是住院病人，
    Dim str姓名 As String
    Dim rsTemp As New ADODB.Recordset
    Dim lng主页ID As Long
    
    On Error GoTo ErrHandle
    If lng单据 = 24 Then
        Check出院病人 = True
        Exit Function
    End If
    
    '如果未传入病人ID，则自动提取
    gstrSQL = "Select A.病人ID,c.主页id From 门诊费用记录 A, 药品收发记录 B,病人医嘱记录 C Where A.ID = B.费用ID  And A.医嘱序号=C.id And b.单据 = [1] And b.No = [2] And Rownum = 1 "
    
    If int记录性质 = 1 Or (int记录性质 = 2 And (int门诊标志 = 1 Or int门诊标志 = 4)) Then
    Else
        gstrSQL = Replace(gstrSQL, "门诊费用记录", "住院费用记录")
    End If
    
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "取病人ID", lng单据, strNo)
    
    '特殊情况，找不到病人ID则不进行下一步检查
    If rsTemp.EOF Then
        Check出院病人 = True
        Exit Function
    End If
    
    lng病人id = rsTemp!病人ID
    lng主页ID = NVL(rsTemp!主页id, 0)

    '取病人姓名
    gstrSQL = "Select 姓名 From 病人信息 Where 病人ID=[1]"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "取病人姓名", lng病人id)

    str姓名 = rsTemp!姓名
    
    '如果当前病人是住院病人，如果没有权限“发退出院病人处方”，则不允许发退药操作
    If zlStr.IsHavePrivs(strPrivs, "发退出院病人处方") = False Then
        '检查病人已预出院或出院
        gstrSQL = " Select 1 From 病案主页" & _
                  " Where 病人ID=[1] and 主页id=[2] " & _
                  " And (出院日期 Is Not NULL)"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "判断是否已出院", lng病人id, lng主页ID)
        
        If rsTemp.RecordCount <> 0 Then
            MsgBox "在处方[" & strNo & "]中，病人“" & str姓名 & "”已出院，你没有对已出院病人的处方进行发料、退料的权限，操作中止！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    Check出院病人 = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Check结帐处方(ByVal strPrivs As String, ByVal lng单据 As Long, ByVal strNo As String, ByVal str序号 As String, ByVal int记录性质 As Integer, ByVal int门诊标志 As Integer) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:检查处方是否已经结帐了,结帐的处方不能发退料操作
    '入参:  lng单据    ：当前单据类型
    '       strNO      ：当前单据号
    '       lng病人ID  ：仅对多病人单有效
    '       str序号：相关单据序号,以,分离
    '出参:
    '返回:数据合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-04-23 14:58:47
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    If lng单据 = 24 Then
        Check结帐处方 = True
        Exit Function
    End If
    
    '如果没有权限“发退结帐处方”，检查该处方是否已结帐，已结帐处方不允许发退料操作
    If zlStr.IsHavePrivs(strPrivs, "发退结帐处方") = 0 Then
    
        gstrSQL = "Select Nvl(Sum(Nvl(结帐金额,0)),0) AS 结帐金额   " & _
                 "  From 门诊费用记录   " & _
                 "  Where Instr([1], ',' || 序号 || ',') > 0 " & _
                 "  And Mod(记录性质,10) = 2 and NO = [2]"
        If int记录性质 = 1 Or (int记录性质 = 2 And (int门诊标志 = 1 Or int门诊标志 = 4)) Then
        Else
            gstrSQL = Replace(gstrSQL, "门诊费用记录", "住院费用记录")
        End If
        gstrSQL = gstrSQL & " Order By 结帐金额 Desc"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "判断是否已结帐", "," & str序号 & ",", strNo)
        If zlStr.NVL(rsTemp!结帐金额, 0) <> 0 Then
            MsgBox "该处方[" & strNo & "]已结帐，你没有对已结帐处方进行发料、退料的权限，操作中止！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    Check结帐处方 = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function IsCtrlSetFocus(ByVal objCtl As Object) As Boolean
    '------------------------------------------------------------------------------
    '功能:判断控件是否可
    '返回:初如成功,返回true,否则返回False
    '编制:刘兴宏
    '日期:2008/01/24
    '------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    err = 0: On Error GoTo ErrHand:
    
    IsCtrlSetFocus = objCtl.Enabled And objCtl.Visible
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub Init站点信息()
    '-----------------------------------------------------------------------------------------------------------
    '功能:初始化站点的相关信息
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2008-09-01 11:32:00
    '-----------------------------------------------------------------------------------------------------------
    gSystem_Para.bln存在站点 = gstrNodeNo <> "-"
 End Sub
 
Public Function zl_获取站点限制(Optional ByVal blnAnd As Boolean = True, _
    Optional ByVal str别名 As String = "") As String
    '功能:获取站点条件限制:2008-09-02 14:30:17
    Dim strWhere As String
    Dim strAlia As String
    strAlia = IIf(str别名 = "", "", str别名 & ".") & "站点"
    strWhere = IIf(blnAnd, " And ", "") & " (" & strAlia & "='" & gstrNodeNo & "' Or " & strAlia & " is Null)"
     zl_获取站点限制 = strWhere
End Function

Public Function InputIsCard(txtInput As Object, KeyAscii As Integer) As Boolean
    '功能：判断指定文本框中当前输入是否在刷卡,根据处理密文显示
    Dim strText As String, blnCard As Boolean
    Dim arrMask As Variant, i As Long
    
    '当前键入后显示的内容(还未显示出来)
    strText = txtInput.Text
    If txtInput.SelLength = Len(txtInput.Text) Then strText = ""
    If KeyAscii = 8 Then
        If strText <> "" Then strText = Mid(strText, 1, Len(strText) - 1)
    Else
        strText = UCase(strText & Chr(KeyAscii))
    End If
    '判断是否在刷卡
    blnCard = False
    If IsNumeric(strText) And IsNumeric(Left(strText, 1)) Then
        blnCard = True
    ElseIf gSystem_Para.str就诊卡前缀符 <> "" Then
        arrMask = Split(gSystem_Para.str就诊卡前缀符, "|")
        For i = 0 To UBound(arrMask)
            If strText Like arrMask(i) & "*" Then
                If IsNumeric(Mid(strText, Len(arrMask(i)) + 1)) And IsNumeric(Mid(strText, Len(arrMask(i)) + 1, 1)) Then
                    blnCard = True
                End If
            End If
        Next
    End If
    
    '刷卡时卡号是否密文显示
    If blnCard Then
       txtInput.PasswordChar = IIf(gSystem_Para.bln就诊卡密文显示 = False, "", "*")
    Else
        txtInput.PasswordChar = ""
    End If
    InputIsCard = blnCard
End Function

Public Sub CheckKeyPress就诊卡号(ByVal txtPati As TextBox, KeyAscii As Integer)
    '检查就诊卡号的输入合法性
    
    If InStr(1, ":：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Public Function GetDeptStationNode(ByVal lngDeptId As Long) As String
'获取部门所属站点信息
    Dim rsSQL As ADODB.Recordset
    Dim strTmp As String
    
    On Error GoTo ErrHandle
    strTmp = "select 站点 from 部门表 where id=[1]"
    Set rsSQL = zldatabase.OpenSQLRecord(strTmp, "获取部门所属站点信息", lngDeptId)
    If Not rsSQL.EOF Then
        GetDeptStationNode = zlStr.NVL(rsSQL!站点)
    End If
    rsSQL.Close
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetArrayByStr(ByVal strInput As String, ByVal lngLength As Long, ByVal strSplitChar As String) As Variant
    '根据传入的字符串进行分解，大于指定字符长度就需要进行分解，结果保存到数组中
    '入参：strInput-输入的字符串；strSplitChar-字符串中内容的分隔符
    '返回：数组，其中数组成员的字符长度不超过指定长度
    Dim strArray As Variant
    Dim arrTmp As Variant
    Dim strTmp As String
    Dim lngCount As Long
    Dim i As Long
    
    strArray = Array()
   
    '大于指定字符时就需要分解
    If Len(strInput) > lngLength Then
        If strSplitChar = "" Then
            '无分隔符时
            strTmp = strInput
            Do While Len(strTmp) > lngLength
                ReDim Preserve strArray(UBound(strArray) + 1)
                strArray(UBound(strArray)) = Mid(strTmp, 1, lngLength)
                strTmp = Mid(strTmp, lngLength + 1)
            Loop
            
            If strTmp <> "" Then
                ReDim Preserve strArray(UBound(strArray) + 1)
                strArray(UBound(strArray)) = strTmp
            End If
        Else
            '有分隔符时
            arrTmp = Split(strInput & strSplitChar, strSplitChar)
            lngCount = UBound(arrTmp)
        
            For i = 0 To lngCount
                If arrTmp(i) <> "" Then
                    '有分隔符的需要保持分隔符之间字符的完整性，不能把分隔符之间的字符拆开
                    If Len(IIf(strTmp = "", "", strTmp & strSplitChar) & arrTmp(i)) > lngLength Then
                        ReDim Preserve strArray(UBound(strArray) + 1)
                        strArray(UBound(strArray)) = strTmp
                        strTmp = arrTmp(i)
                    Else
                        strTmp = IIf(strTmp = "", "", strTmp & strSplitChar) & arrTmp(i)
                    End If
                End If
                       
                If i = lngCount Then
                    ReDim Preserve strArray(UBound(strArray) + 1)
                    strArray(UBound(strArray)) = strTmp
                End If
            Next
        End If
    Else
        ReDim Preserve strArray(UBound(strArray) + 1)
        strArray(UBound(strArray)) = strInput
    End If
    
    GetArrayByStr = strArray
End Function
