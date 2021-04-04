Attribute VB_Name = "mdlSQL"
Option Explicit

Public Enum SQL
    
    人员基本资料
    分科项目结果
    分科项目结论
    分科项目诊断
    总检报告建议
    主检诊断结果
    
    药品执行科室
    诊疗执行科室
    收费执行科室
    体检项目价表
End Enum

Public Function GetPublicSQL(ByVal intMenu As SQL, Optional ByVal strParam As String) As String
    '------------------------------------------------------------------------------------------------------------------
    '功能:  集中产生SQL语句
    '参数:  strMenu             要产生的SQL名称
    '       strParam            参数串,格式:"参数值1'参数值2"
    '返回:  SQL语句
    '------------------------------------------------------------------------------------------------------------------
    
    Dim strSQL As String
    Dim varParam As Variant
    Dim strTmp As String
    Dim rs As New ADODB.Recordset
            
    On Error GoTo errHand
    
    If strParam = "" Then strParam = "'"
    
    varParam = Split(strParam, "'")
    
    Select Case intMenu
        Case SQL.人员基本资料
            
'            strSQL = "Select * From 体检人员档案_干保 A,体检登记记录_干保 B,病人信息_干保 C,病人信息 D,体检人员档案 E  " & _
'                        "WHERE D.病人id=A.病人id And C.病人id=A.病人id And A.任务包号=B.任务包号 And E.登记ID=B.登记ID And E.病人ID=A.病人ID " & _
'                                "AND B.任务包号='" & varParam(0) & "' And A.病人id=" & Val(varParam(1))
                                
            strSQL = "Select * From 体检人员档案_干保 A,体检登记记录_干保 B,病人信息 D,体检人员档案 E,体检组别_干保 c  " & _
                        "WHERE D.病人id=A.病人id And A.任务包号=B.任务包号 And E.登记ID=B.登记ID And E.病人ID=A.病人ID " & _
                                "AND c.登记id=b.登记id and c.组别名称=e.组别名称 and B.任务包号=[1] And A.病人id=[2]"
                                
        Case SQL.分科项目结果
            
            strSQL = _
            "SELECT 所见项id,体检项目id,执行部门id,结果,标志,单位,参考,组合编码,组合科室 " & _
            "FROM ( " & _
              "SELECT " & _
                     "R.执行部门ID, " & _
                     "R.结果, " & _
                     "R.单位,R.组合编码,R.组合科室, " & _
                     "DECODE(SIGN(INSTR(R.标志参考,'''')),1,SUBSTR(R.标志参考,1,INSTR(R.标志参考,'''')-1),'') AS 标志, " & _
                     "DECODE(SIGN(INSTR(R.标志参考,'''')),1,SUBSTR(R.标志参考,INSTR(R.标志参考,'''')+1,1000),'') AS 参考, " & _
                     "体检项目id, " & _
                     "所见项id " & _
              "FROM ( " & _
                    "Select " & _
                           "A.执行部门ID, " & _
                           "A.体检项目id,A.组合编码,A.组合科室," & _
                           "A.ID, " & _
                           "X.排列序号, " & _
                           "Y.内序号, " & _
                           "DECODE(SIGN(INSTR(Y.结果,'''')),1,SUBSTR(Y.结果,1,INSTR(Y.结果,'''')-1),Y.结果) AS 结果, " & _
                           "Y.单位, " & _
                           "Y.所见项id, " & _
                           "DECODE(SIGN(INSTR(Y.结果,'''')),1,SUBSTR(Y.结果,INSTR(Y.结果,'''')+1,1000),'') AS 标志参考 " & _
                    "From "
                    
            strSQL = strSQL & _
                         "( " & _
                         "Select DISTINCT A1.医嘱ID,A3.执行部门ID,A4.ID,A5.诊疗项目id AS 体检项目id,A6.干保编码 As 组合编码,A6.组合科室 " & _
                         "from 体检项目医嘱 A1, " & _
                               "病人医嘱记录 A2, " & _
                               "病人医嘱发送 A3, " & _
                               "病人病历记录 A4, " & _
                               "体检项目清单 A5,诊疗项目目录_干保 A6 " & _
                         "Where A1.病人id =[2] " & _
                                " AND A5.登记id=[1] " & _
                                " AND (A1.医嘱ID=A2.ID OR A1.医嘱ID=A2.相关id) " & _
                                "AND A3.医嘱ID=A2.ID " & _
                                "AND A4.ID=A3.报告ID AND A6.诊疗项目id=A2.诊疗项目id " & _
                                "AND A5.ID=A1.清单ID  AND A2.诊疗类别 In ('C','D') " & _
                         ") A, " & _
                         "病人病历内容 X, " & _
                         "( " & _
                         "select " & _
                                 "A.病历ID, " & _
                                 "A.控件号 AS 内序号, " & _
                                 "A.所见内容 AS 结果, " & _
                                 "B.单位, " & _
                                 "A.所见项id " & _
                          "From "
                          
            strSQL = strSQL & _
                            "病人病历所见单 A, " & _
                            "诊治所见项目 B " & _
                          "Where A.所见项id = B.ID And 所见项id > 0 " & _
                          ") Y " & _
                    "Where x.病历记录id = A.ID And X.ID = Y.病历ID " & _
                    ") R " & _
                ") A"
            
            strSQL = "Select W.*,T.干保编码 As 项目编码,T.项目分支,T.项目方法,T.干保名称 As 项目名称 From (" & strSQL & ") W,诊治所见项目_干保 T,体检人员档案_干保 K " & _
                        "WHERE T.诊治项目id=W.所见项id " & _
                                "AND K.病人id=[2] And K.任务包号=[3]"
        
        Case SQL.分科项目结论
                
            strSQL = _
                "Select " & _
                       "Distinct y.结论描述,0 As 疾病id,Y.诊断建议, A.执行部门ID, A.体检项目id,A.书写人,A.审阅日期 " & _
                "From " & _
                     "( " & _
                     "Select DISTINCT A1.医嘱ID,A3.执行部门ID,A4.ID,A5.诊疗项目id AS 体检项目id,A4.书写人,A4.审阅日期 " & _
                     "from 体检项目医嘱 A1, " & _
                           "病人医嘱记录 A2, " & _
                           "病人医嘱发送 A3, " & _
                           "病人病历记录 A4, " & _
                           "体检项目清单 A5 " & _
                     "Where A1.病人id =[2] " & _
                            " AND A5.登记id=[1] " & _
                            " AND (A1.医嘱ID=A2.ID OR A1.医嘱ID=A2.相关id) " & _
                            "AND A3.医嘱ID=A2.ID " & _
                            "AND A4.ID=A3.报告ID " & _
                            "AND A5.ID=A1.清单ID AND A2.诊疗类别 In ('C','D') " & _
                     ") A, " & _
                     "病人病历内容 X, " & _
                     "体检人员结论 Y " & _
                "Where x.病历记录id = A.ID And x.ID = y.病历ID And y.结论描述 Is Not Null "
            
            strSQL = "Select Distinct W.*,T.干保编码 As 组合编码,T.组合科室,T.干保名称 As 组合名称 From (" & strSQL & ") W,诊疗项目目录_干保 T,体检人员档案_干保 K " & _
                        "WHERE T.诊疗项目id=W.体检项目id " & _
                                "AND K.病人id=[2] And K.任务包号=[3] Order By T.组合科室,T.干保编码"
                                
        Case SQL.分科项目诊断
            
            strSQL = _
                "Select " & _
                       "Distinct 0 As 疾病id,Y.诊断建议,Y.结论id,A.执行部门ID,A.体检项目id,A.书写人,A.审阅日期 " & _
                "From " & _
                     "( " & _
                     "Select DISTINCT A1.医嘱ID,A3.执行部门ID,A4.ID,A5.诊疗项目id AS 体检项目id,A4.书写人,A4.审阅日期 " & _
                     "from 体检项目医嘱 A1, " & _
                           "病人医嘱记录 A2, " & _
                           "病人医嘱发送 A3, " & _
                           "病人病历记录 A4, " & _
                           "体检项目清单 A5 " & _
                     "Where A1.病人id =[2] " & _
                            " AND A5.登记id=[1] " & _
                            " AND (A1.医嘱ID=A2.ID OR A1.医嘱ID=A2.相关id) " & _
                            "AND A3.医嘱ID=A2.ID " & _
                            "AND A4.ID=A3.报告ID " & _
                            "AND A5.ID=A1.清单ID  AND A2.诊疗类别 In ('C','D') " & _
                     ") A, " & _
                     "病人病历内容 X, " & _
                     "体检人员结论 Y " & _
                "Where x.病历记录id = A.ID And x.ID = y.病历ID And Y.结论id Is Not Null"
            
            strSQL = "Select Distinct W.*,T.干保编码 As 组合编码,T.组合科室,T.干保名称 As 组合名称,X.干保编码 As 诊断编码,X.干保名称 As 诊断名称,X.疾病编码 From (" & strSQL & ") W,诊疗项目目录_干保 T,体检人员档案_干保 K,体检诊断建议_干保 X " & _
                        "WHERE T.诊疗项目id=W.体检项目id " & _
                                "AND X.结论id=W.结论id AND X.疾病编码 Is Not Null " & _
                                "AND K.病人id=[2] And K.任务包号=[3] Order By T.组合科室,T.干保编码"
        Case SQL.主检诊断结果
            
            strSQL = _
                "SELECT  Distinct 0 As 疾病id,X1.诊断建议,X1.结论id,Y.科室id,Y.书写人,Y.审阅日期 " & _
                "FROM 体检人员档案 A, " & _
                     "病人病历内容 X, " & _
                     "病人病历记录 Y, " & _
                      "体检人员结论 X1,病历元素目录 X2 " & _
                "Where X.病历记录id = A.体检病历ID " & _
                      "AND X.ID=X1.病历id " & _
                      "AND Y.ID=X.病历记录id " & _
                      "AND A.病人ID=[2] " & _
                      " AND A.登记ID=[1] " & _
                      " AND X.元素类型=4 and X.元素编码=X2.编码 AND Upper(X2.部件)='ZL9CISCORE.USRMEDICALSUM'"
            
            strSQL = "Select Distinct W.*,X.干保编码 As 诊断编码,X.干保名称 As 诊断名称,X.疾病编码 From (" & strSQL & ") W,体检人员档案_干保 K,体检诊断建议_干保 X " & _
                        "WHERE X.结论id=W.结论id AND X.疾病编码 Is Not Null " & _
                                "AND K.病人id=[2] And K.任务包号=[3] "
                                
        Case SQL.总检报告建议
        
            strSQL = _
                "SELECT  DECODE(SIGN(INSTR(结果,'二、建议：')),1,SUBSTR(结果,8,INSTR(结果,'二、建议：')-11),结果) AS 报告头," & _
                        "DECODE(SIGN(INSTR(结果,'二、建议：')),1,SUBSTR(结果,INSTR(结果,'二、建议：')+7,4000),结果) AS 健康指导," & _
                        "书写人, " & _
                        "TO_CHAR(书写日期,'yyyy-mm-dd') AS 书写日期 " & _
                "FROM ( " & _
                "select " & _
                       "X.排列序号, " & _
                       "X1.内序号, " & _
                       "X1.结果, " & _
                       "Y.书写人, " & _
                       "y.书写日期 " & _
                "From " & _
                     "体检人员档案 A, " & _
                     "病人病历内容 X, " & _
                     "病人病历记录 Y,病历元素目录 X2, " & _
                      "(select 病历id,0 AS 内序号,'' AS 项目,内容 AS 结果 from 病人病历文本段 ) X1 " & _
                "Where x.病历记录id = A.体检病历ID " & _
                      "AND X.ID=X1.病历id " & _
                      "AND Y.ID=X.病历记录id " & _
                      "AND A.病人ID=[2] " & _
                      " AND A.登记ID=[1] " & _
                      " AND X.元素类型=4 and X.元素编码=X2.编码 AND upper(X2.部件)='ZL9CISCORE.USRMEDICALSUM' " & _
                ") ORDER BY  排列序号,内序号"

        Case SQL.诊疗执行科室
                                                                                        
            '参数:诊疗项目id'病人科室id'开单科室id'查找内容
            
            strSQL = _
                "SELECT A.ID FROM 部门表 A,诊疗项目目录 X WHERE X.ID=[1] AND X.执行科室=1 AND A.ID=[2]"
            
            strSQL = strSQL & " UNION ALL " & _
                "SELECT A.ID FROM 部门表 A,床位状况记录 B,诊疗项目目录 X WHERE X.ID=[1] AND X.执行科室=2 AND A.ID=B.病区id AND B.科室ID=[2]"
                
            strSQL = strSQL & " UNION ALL " & _
                "SELECT A.ID FROM 部门表 A,诊疗项目目录 X WHERE X.ID=[1] AND X.执行科室=3 AND A.ID=[3]"
            
            strSQL = strSQL & " UNION ALL " & _
                "SELECT A.ID FROM 部门表 A,诊疗执行科室 B,诊疗项目目录 X WHERE X.ID=[1] AND X.执行科室=4 AND A.ID=B.执行科室id AND B.病人来源=1 AND B.诊疗项目id=X.ID"
                
            strSQL = strSQL & " UNION ALL " & _
                "SELECT A.ID FROM 部门表 A,诊疗执行科室 B,诊疗项目目录 X WHERE X.ID=[1] AND X.执行科室=4 AND " & _
                            "A.ID=B.执行科室id AND B.病人来源 IS NULL AND (B.开单科室id IS NULL OR B.开单科室id=[3]) AND B.诊疗项目id=X.ID "
            
            strSQL = _
                "SELECT 1 As 末级,A.编码,A.名称,A.简码,A.ID FROM 部门表 A WHERE A.ID IN (" & strSQL & ") AND (UPPER(A.编码) Like [4] OR UPPER(A.简码) Like [4] OR A.名称 Like [4])"
        
    Case SQL.药品执行科室
            
            strSQL = "SELECT Distinct 1 As 末级,A.编码,A.名称,A.ID " & _
                    "from 部门表 A,部门性质说明 B " & _
                    "where (A.撤档时间 IS NULL OR A.撤档时间 =TO_DATE('3000-01-01','YYYY-MM-DD'))" & _
                    "and A.ID=B.部门ID and B.服务对象 in (2,3) " & _
                    "and B.工作性质=Decode([1],'5','西药房','6','成药房','7','中药房','4','发料部门')"
                    
    Case SQL.收费执行科室
        
            '参数:诊疗项目id'病人科室id'开单科室id'查找内容
            
            strSQL = _
                "SELECT A.ID FROM 部门表 A,收费项目目录 X WHERE X.ID=[1] AND X.执行科室=1 AND A.ID=[2]"
            
            strSQL = strSQL & " UNION ALL " & _
                "SELECT A.ID FROM 部门表 A,床位状况记录 B,收费项目目录 X WHERE X.ID=[1] AND X.执行科室=2 AND A.ID=B.病区id AND B.科室ID=[2]"
                
            strSQL = strSQL & " UNION ALL " & _
                "SELECT A.ID FROM 部门表 A,收费项目目录 X WHERE X.ID=[1] AND X.执行科室=3 AND A.ID=[3]"
            
            strSQL = strSQL & " UNION ALL " & _
                "SELECT A.ID FROM 部门表 A,收费执行科室 B,收费项目目录 X WHERE X.ID=[1] AND X.执行科室=4 AND A.ID=B.执行科室id AND B.病人来源=1 AND B.收费细目id=X.ID"
                
            strSQL = strSQL & " UNION ALL " & _
                "SELECT A.ID FROM 部门表 A,收费执行科室 B,收费项目目录 X WHERE X.ID=[1] AND X.执行科室=4 AND " & _
                            "A.ID=B.执行科室id AND B.病人来源 IS NULL AND (B.开单科室id IS NULL OR B.开单科室id=[3]) AND B.收费细目id=X.ID "
            
            strSQL = _
                "SELECT 1 As 末级,A.编码,A.名称,A.ID FROM 部门表 A WHERE A.ID IN (" & strSQL & ") AND (UPPER(A.编码) Like [4] OR UPPER(A.简码) Like [4] OR A.名称 Like [4])"
                                    
                                    
    Case SQL.体检项目价表
            
            strTmp = Val(varParam(0)) & "," & varParam(2)
            If Right(strTmp, 1) = "," Then strTmp = strTmp & "0"
            
            strSQL = "Select y.名称,y.计算单位,z.收费数量,x.现价,y.id,1 As 计价性质,y.类别 " & _
                        "From ( " & _
                          "Select a.诊疗项目id,a.收费项目id,Sum(c.现价) As 现价 " & _
                          "From 收费价目 c, " & _
                               "诊疗收费关系 a, " & _
                               "诊疗项目目录 b " & _
                          "Where a.收费项目id = c.收费细目id " & _
                                "and c.执行日期<=SYSDATE and (c.终止日期 IS NULL OR c.终止日期>SYSDATE) " & _
                                "AND b.ID=a.诊疗项目id " & _
                                "AND NVL(b.计价性质,0)=0 " & _
                                "and a.诊疗项目id IN (" & strTmp & ") " & _
                          "Group by a.诊疗项目id,a.收费项目id " & _
                        ") x, " & _
                        "收费项目目录 y, " & _
                        "诊疗收费关系 z " & _
                        "Where x.收费项目id = y.ID " & _
                              "and z.收费项目id=x.收费项目id " & _
                              "and z.诊疗项目id=x.诊疗项目id"
                                          
            strSQL = strSQL & " Union All Select y.名称,y.计算单位,z.收费数量,x.现价,y.id,2 As 计价性质,y.类别 " & _
                        "From ( " & _
                          "Select a.诊疗项目id,a.收费项目id,Sum(c.现价) As 现价 " & _
                          "From 收费价目 c, " & _
                               "诊疗收费关系 a, " & _
                               "诊疗项目目录 b " & _
                          "Where a.收费项目id = c.收费细目id " & _
                                "and c.执行日期<=SYSDATE and (c.终止日期 IS NULL OR c.终止日期>SYSDATE) " & _
                                "AND b.ID=a.诊疗项目id " & _
                                "AND NVL(b.计价性质,0)=0 " & _
                                "and a.诊疗项目id=" & Val(varParam(1)) & " " & _
                          "Group by a.诊疗项目id,a.收费项目id " & _
                        ") x, " & _
                        "收费项目目录 y, " & _
                        "诊疗收费关系 z " & _
                        "Where x.收费项目id = y.ID " & _
                              "and z.收费项目id=x.收费项目id " & _
                              "and z.诊疗项目id=x.诊疗项目id"
    End Select
    
    GetPublicSQL = strSQL
    
    Exit Function
    
errHand:
    
End Function




