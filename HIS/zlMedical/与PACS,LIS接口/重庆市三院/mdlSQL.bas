Attribute VB_Name = "mdlSQL"
Option Explicit

Public Enum SQL
    
    人员基本资料
    分科项目结果
    分科项目结论
    分科项目诊断
    总检报告建议
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
            
            strSQL = "Select * From 体检人员档案_干保 A,体检登记记录_干保 B,病人信息_干保 C,病人信息 D,体检人员档案 E  " & _
                        "WHERE D.病人id=A.病人id And C.病人id=A.病人id And A.任务包号=B.任务包号 And E.登记ID=B.登记ID And E.病人ID=A.病人ID " & _
                                "AND B.任务包号='" & varParam(0) & "' And A.病人id=" & Val(varParam(1))
                                
        Case SQL.分科项目结果
            
            strSQL = _
            "SELECT 所见项id,体检项目id,执行部门id,结果,标志,单位,参考 " & _
            "FROM ( " & _
              "SELECT " & _
                     "R.执行部门ID, " & _
                     "R.结果, " & _
                     "R.单位, " & _
                     "DECODE(SIGN(INSTR(R.标志参考,'''')),1,SUBSTR(R.标志参考,1,INSTR(R.标志参考,'''')-1),'') AS 标志, " & _
                     "DECODE(SIGN(INSTR(R.标志参考,'''')),1,SUBSTR(R.标志参考,INSTR(R.标志参考,'''')+1,1000),'') AS 参考, " & _
                     "体检项目id, " & _
                     "所见项id " & _
              "FROM ( " & _
                    "Select " & _
                           "A.执行部门ID, " & _
                           "A.体检项目id, " & _
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
                         "Select DISTINCT A1.医嘱ID,A3.执行部门ID,A4.ID,A5.诊疗项目id AS 体检项目id " & _
                         "from 体检项目医嘱 A1, " & _
                               "病人医嘱记录 A2, " & _
                               "病人医嘱发送 A3, " & _
                               "病人病历记录 A4, " & _
                               "体检项目清单 A5 " & _
                         "Where A1.病人id = " & Val(varParam(1)) & _
                                " AND A5.登记id=" & Val(varParam(0)) & _
                                " AND (A1.医嘱ID=A2.ID OR A1.医嘱ID=A2.相关id) " & _
                                "AND A3.医嘱ID=A2.ID " & _
                                "AND A4.ID=A3.报告ID " & _
                                "AND A5.ID=A1.清单ID " & _
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
            
            strSQL = "Select W.*,P.组合编码,P.项目编码,P.项目分支,P.项目方法,P.组合科室,T.干保名称 As 项目名称 From (" & strSQL & ") W,诊治所见项目_干保 T,体检项目清单_干保 P,体检人员档案_干保 K " & _
                        "WHERE T.诊治项目id=W.所见项id " & _
                                "And T.干保编码=P.项目编码 " & _
                                "AND K.任务包号=P.任务包号 " & _
                                "AND K.套餐编码=P.套餐编码 " & _
                                "AND K.套餐序号=P.套餐序号 " & _
                                "AND K.病人id=" & Val(varParam(1)) & " And K.任务包号='" & varParam(2) & "'"
        
        Case SQL.分科项目结论
                
            strSQL = _
                "Select " & _
                       "Distinct y.结论描述,Y.疾病id,Y.诊断建议, A.执行部门ID, A.体检项目id,A.书写人,A.审阅日期 " & _
                "From " & _
                     "( " & _
                     "Select DISTINCT A1.医嘱ID,A3.执行部门ID,A4.ID,A5.诊疗项目id AS 体检项目id,A4.书写人,A4.审阅日期 " & _
                     "from 体检项目医嘱 A1, " & _
                           "病人医嘱记录 A2, " & _
                           "病人医嘱发送 A3, " & _
                           "病人病历记录 A4, " & _
                           "体检项目清单 A5 " & _
                     "Where A1.病人id = " & Val(varParam(1)) & _
                            " AND A5.登记id=" & Val(varParam(0)) & _
                            " AND (A1.医嘱ID=A2.ID OR A1.医嘱ID=A2.相关id) " & _
                            "AND A3.医嘱ID=A2.ID " & _
                            "AND A4.ID=A3.报告ID " & _
                            "AND A5.ID=A1.清单ID " & _
                     ") A, " & _
                     "病人病历内容 X, " & _
                     "体检人员结论 Y " & _
                "Where x.病历记录id = A.ID And x.ID = y.病历ID And y.结论描述 Is Not Null "
            
            strSQL = "Select Distinct W.*,P.组合编码,P.组合科室,T.干保名称 As 组合名称 From (" & strSQL & ") W,诊疗项目目录_干保 T,体检项目清单_干保 P,体检人员档案_干保 K " & _
                        "WHERE T.诊疗项目id=W.体检项目id " & _
                                "And T.干保编码=P.组合编码 " & _
                                "AND K.任务包号=P.任务包号 " & _
                                "AND K.套餐编码=P.套餐编码 " & _
                                "AND K.套餐序号=P.套餐序号 " & _
                                "AND K.病人id=" & Val(varParam(1)) & " And K.任务包号='" & varParam(2) & "' Order By P.组合科室,P.组合编码"
                                
        Case SQL.分科项目诊断
            
            strSQL = _
                "Select " & _
                       "Distinct Y.疾病id,Y.诊断建议,Y.结论id,A.执行部门ID,A.体检项目id,A.书写人,A.审阅日期 " & _
                "From " & _
                     "( " & _
                     "Select DISTINCT A1.医嘱ID,A3.执行部门ID,A4.ID,A5.诊疗项目id AS 体检项目id,A4.书写人,A4.审阅日期 " & _
                     "from 体检项目医嘱 A1, " & _
                           "病人医嘱记录 A2, " & _
                           "病人医嘱发送 A3, " & _
                           "病人病历记录 A4, " & _
                           "体检项目清单 A5 " & _
                     "Where A1.病人id = " & Val(varParam(1)) & _
                            " AND A5.登记id=" & Val(varParam(0)) & _
                            " AND (A1.医嘱ID=A2.ID OR A1.医嘱ID=A2.相关id) " & _
                            "AND A3.医嘱ID=A2.ID " & _
                            "AND A4.ID=A3.报告ID " & _
                            "AND A5.ID=A1.清单ID " & _
                     ") A, " & _
                     "病人病历内容 X, " & _
                     "体检人员结论 Y " & _
                "Where x.病历记录id = A.ID And x.ID = y.病历ID And Y.疾病id Is Not Null And Y.结论id Is Not Null"
            
            strSQL = "Select Distinct W.*,P.组合编码,P.项目分支,P.组合科室,T.干保名称 As 组合名称,X.干保编码 As 诊断编码,X.干保名称 As 诊断名称,L.编码 As 疾病编码 From (" & strSQL & ") W,诊疗项目目录_干保 T,体检项目清单_干保 P,体检人员档案_干保 K,体检诊断建议_干保 X,疾病编码目录 L " & _
                        "WHERE T.诊疗项目id=W.体检项目id " & _
                                "And T.干保编码=P.组合编码 " & _
                                "AND K.任务包号=P.任务包号 " & _
                                "AND K.套餐编码=P.套餐编码 " & _
                                "AND K.套餐序号=P.套餐序号 " & _
                                "AND X.结论id=W.结论id " & _
                                "AND L.ID=W.疾病id " & _
                                "AND K.病人id=" & Val(varParam(1)) & " And K.任务包号='" & varParam(2) & "' Order By P.组合科室,P.组合编码"
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
                 "病人病历记录 Y, " & _
                  "(select 病历id,0 AS 内序号,'' AS 项目,内容 AS 结果 from 病人病历文本段 ) X1 " & _
            "Where x.病历记录id = A.体检病历ID " & _
                  "AND X.ID=X1.病历id " & _
                  "AND Y.ID=X.病历记录id " & _
                  "AND A.病人ID=" & Val(varParam(1)) & _
                  " AND A.登记ID=" & Val(varParam(0)) & _
                  " AND X.元素类型=4 and X.元素编码='000055' " & _
            ") ORDER BY  排列序号,内序号"

    End Select
    
    GetPublicSQL = strSQL
    
    Exit Function
    
errHand:
    
End Function




