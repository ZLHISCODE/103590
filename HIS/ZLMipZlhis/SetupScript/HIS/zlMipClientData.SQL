
--zlmip_table
Insert Into zlmip_table(id,data_code,tab_type,tab_code,tab_title,tab_sqltext,tab_note)
          Select '93A3FF349B5A4C0FA15AFE6298F3B390','HIS',1,'ZLHIS_PUB_001','部门人员','Select b.id,b.姓名,b.性别,b.简码,b.身份证号,b.出生日期 from 部门人员 a,人员表 b where a.部门id=[0] And a.人员id=b.id and a.缺省=1','' From Dual
Union All Select '6DF2D77B66224BDBA99E67202C1BDB3F','HIS',1,'ZLHIS_PUB_002','部门信息','select * from 部门表 where id=[0]','' From Dual;

--zlmip_tab_field
Insert Into zlmip_tab_field(tab_id,fld_order,fld_title,fld_type)
          Select '93A3FF349B5A4C0FA15AFE6298F3B390',1,'ID',1 From Dual
Union All Select '93A3FF349B5A4C0FA15AFE6298F3B390',2,'姓名',2 From Dual
Union All Select '93A3FF349B5A4C0FA15AFE6298F3B390',3,'性别',2 From Dual
Union All Select '93A3FF349B5A4C0FA15AFE6298F3B390',4,'简码',2 From Dual
Union All Select '93A3FF349B5A4C0FA15AFE6298F3B390',5,'身份证号',2 From Dual
Union All Select '93A3FF349B5A4C0FA15AFE6298F3B390',6,'出生日期',3 From Dual
Union All Select '6DF2D77B66224BDBA99E67202C1BDB3F',1,'ID',1 From Dual
Union All Select '6DF2D77B66224BDBA99E67202C1BDB3F',2,'上级ID',1 From Dual
Union All Select '6DF2D77B66224BDBA99E67202C1BDB3F',3,'编码',2 From Dual
Union All Select '6DF2D77B66224BDBA99E67202C1BDB3F',4,'名称',2 From Dual
Union All Select '6DF2D77B66224BDBA99E67202C1BDB3F',5,'简码',2 From Dual
Union All Select '6DF2D77B66224BDBA99E67202C1BDB3F',6,'位置',2 From Dual
Union All Select '6DF2D77B66224BDBA99E67202C1BDB3F',7,'末级',1 From Dual
Union All Select '6DF2D77B66224BDBA99E67202C1BDB3F',8,'建档时间',3 From Dual
Union All Select '6DF2D77B66224BDBA99E67202C1BDB3F',9,'撤档时间',3 From Dual
Union All Select '6DF2D77B66224BDBA99E67202C1BDB3F',10,'站点',2 From Dual
Union All Select '6DF2D77B66224BDBA99E67202C1BDB3F',11,'环境类别',2 From Dual
Union All Select '6DF2D77B66224BDBA99E67202C1BDB3F',12,'部门负责人',1 From Dual;

--zlmip_tab_parameter
Insert Into zlmip_tab_parameter(tab_id,para_order,para_field,para_title,para_type,para_default,para_note)
          Select '93A3FF349B5A4C0FA15AFE6298F3B390',1,'','部门id',1,'','' From Dual
Union All Select '6DF2D77B66224BDBA99E67202C1BDB3F',1,'','部门id',1,'','' From Dual;

--zlmip_tab_extend
Insert Into zlmip_tab_extend(ID,SOURCE_TAB_ID,EXT_ORDER,EXT_TYPE,EXT_TITLE,TARGET_TAB_ID)
	  Select '1D680D93960C4F6185EF89D60884B0C8','6DF2D77B66224BDBA99E67202C1BDB3F',1,2,'部门人员','93A3FF349B5A4C0FA15AFE6298F3B390' From Dual;

--zlmip_tabext_condition
Insert Into zlmip_tabext_condition(EXT_ID,COND_ORDER,TARGET_FLD,SOURCE_FLD)
	  Select '1D680D93960C4F6185EF89D60884B0C8',1,'部门id','ID' From Dual;

--zlmip_item
Insert Into zlmip_item(id,data_code,item_code,item_title,item_request,item_type,trigger_type,again_policy,again_para,check_frequency,check_freq_internal,trigger_condition,trigger_frequency,tab_id,start_date,stop_date,item_note)
          Select '29C186B7661F4202BC60AB7319498FF9','HIS','ZLHIS_CHARGE_001','费用销帐申请通知',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select '2755C4E8A9324E33B29A08A0B2A0E480','HIS','ZLHIS_CHARGE_003','门诊费用单据',0,1,1,0,Null,0,0,'',0,Null,Sysdate,Null,'' From Dual
Union All Select '96D430F4877D406FA40672E641E323E7','HIS','ZLHIS_CIS_001','医生新开患者医嘱',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select 'D95BC90DB297434791792921393A5CD9','HIS','ZLHIS_CIS_002','医生新停患者医嘱',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select 'E0147BE1C6A640CEAC595C519490AD7B','HIS','ZLHIS_CIS_003','作废患者医嘱',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select 'C6E6C17F67C945C1903D23FD2F68BB87','HIS','ZLHIS_CIS_004','医嘱申请安排',0,1,1,0,Null,0,0,'',0,Null,Sysdate,Null,'' From Dual
Union All Select '73A421D7B7E44815A1BA40170BC2F509','HIS','ZLHIS_CIS_005','医技执行安排完成',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select 'F36F8875F1E540908E8DC62BA3C42168','HIS','ZLHIS_CIS_006','患者药品医嘱发送',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select 'D2F8C10D32AA48E19A485787F25C70A8','HIS','ZLHIS_CIS_007','门诊患者转诊',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select 'D862BD945CF74CA49C8110A1A2F1FDA5','HIS','ZLHIS_CIS_008','输液批次调整通知',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select 'B0C04DB87C284EB38B49141479B904A1','HIS','ZLHIS_CIS_009','门诊患者接诊通知',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select 'F02900C556BB4008A9DC046B1365FC26','HIS','ZLHIS_CIS_010','下达患者诊断',0,1,1,0,Null,0,0,'',0,Null,Sysdate,Null,'' From Dual
Union All Select '150A63E2931B4A309D65474C3E8ED00B','HIS','ZLHIS_CIS_011','撤消患者诊断',0,1,1,0,Null,0,0,'',0,Null,Sysdate,Null,'' From Dual
Union All Select '970C2D8BACB445618065E0B6851BECA9','HIS','ZLHIS_CIS_012','住院患者出院诊断',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select '78C19BDBF65D4784B3EC8A12B2489C61','HIS','ZLHIS_CIS_013','住院患者输液销帐申请',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select '8DF1D60F9642408FA984C1FB423C8E74','HIS','ZLHIS_CIS_014','检验危急值阅读通知',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select '10A0A36A7AE84876926575D5E67E2465','HIS','ZLHIS_CIS_015','医技拒绝执行通知',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select '409BD4665E1640669BB29EDDFBB09B10','HIS','ZLHIS_CIS_016','患者检验申请',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select 'D31F0420F60049B3837552290EB46703','HIS','ZLHIS_CIS_017','患者检查申请',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select 'C7CE07FDB5594C2DB11D50D1DD0BB795','HIS','ZLHIS_CIS_018','患者手术申请',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select '1BADC496188C4E2C8BCCD580D8070BD6','HIS','ZLHIS_CIS_019','患者输血申请',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select '323625B3583A4725A9B4C9A26D90C65D','HIS','ZLHIS_CIS_020','患者会诊申请',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select '11DF9A881F6D4EB4AC5141EA7785B6BA','HIS','ZLHIS_CIS_021','患者抢救医嘱',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select 'B3038E177E3448F0929B2938BE5CD443','HIS','ZLHIS_CIS_022','患者死亡医嘱',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select '877261D6A13849F8AC941E5E7E9FE562','HIS','ZLHIS_CIS_023','患者特殊治疗医嘱',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select '6082481F1E47484EB7D93ADCB2C94A13','HIS','ZLHIS_CIS_024','患者医嘱撤消',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select 'E52632D77EB843F1BF16F621855C3F66','HIS','ZLHIS_CIS_025','检查危急值阅读通知',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select 'D3D2A3A3EBD549CFB8761AA53BC5A9E9','HIS','ZLHIS_PACS_001','检查报告完成',0,1,1,0,Null,0,0,'',0,Null,Sysdate,Null,'' From Dual
Union All Select '98125826A65144D5A8330ECAC56F225B','HIS','ZLHIS_PACS_002','检查状态同步',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select '92F5E98A99A742299CF0FA52D98B0417','HIS','ZLHIS_PACS_003','检查状态回退',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select 'EB810C78495C46128FD31111C382407C','HIS','ZLHIS_PACS_004','检查报告撤销',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select '50146DFA544E4FAC9EBC4851DC1A4B5B','HIS','ZLHIS_PACS_005','检查危急值通知',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select 'F795028B5D8C4A9CAD49192A3A06FA5F','HIS','ZLHIS_PATIENT_001','住院患者入院登记',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select '1373E96D23D44E1388F27F268718C694','HIS','ZLHIS_PATIENT_002','住院患者入院入科',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select 'B67FC7C97266418D82E47209D7C41C9B','HIS','ZLHIS_PATIENT_003','住院患者转出科室',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select '61505E3D721544B8974558CC3420C19C','HIS','ZLHIS_PATIENT_004','住院患者床位变更',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select '447C3A204ED94FF3ACC8CDBBAA3C0C42','HIS','ZLHIS_PATIENT_005','住院患者病情变更',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select '3B793D544B914C72A958E1D7F8E8C247','HIS','ZLHIS_PATIENT_006','住院患者变动撤消',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select '20D31ACCAAC64CDCA242787B01C6D44E','HIS','ZLHIS_PATIENT_007','住院患者医护变更',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select '7BD50BB1610E4404B6B2B41A6B245E93','HIS','ZLHIS_PATIENT_009','住院患者预出院',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select '12FE4EEC12964216A2790B2AFFC55BF0','HIS','ZLHIS_PATIENT_010','住院患者出院',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select '1B38A2600FD749D68AF43FF8A6017A2A','HIS','ZLHIS_PATIENT_011','住院患者新生儿登记',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select '9034EA214816475B85CB9C83E377AE1D','HIS','ZLHIS_PATIENT_012','住院患者转入科室',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select '37F9BA5076D84A66B0DA03A8C293D42B','HIS','ZLHIS_PATIENT_013','新生儿登记作废',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select '39643F2E01CB4249B282CDB1CB278AB7','HIS','ZLHIS_PUB_001','生日祝福',0,1,2,0,Null,0,0,'Select a.Id, a.编号, a.姓名, a.性别, a.民族, Decode(a.性别, ''女'', ''女士'', ''男'', ''先生'', Null) As 称呼, 移动电话, a.电子邮件,
       To_Char(a.出生日期, ''yyyy-MM-dd'') As 出生日期
From 人员表 A
Where to_char(a.出生日期,''mm-dd'') = to_char(Sysdate,''mm-dd'')',0,'6DF2D77B66224BDBA99E67202C1BDB3F',Sysdate,Null,'通过短信、邮件、ZLHIS产品祝福当天过生的人员' From Dual
Union All Select '81B4204CC0B44877A3DF714BB13773BD','HIS','ZLHIS_QUEUE_001','队列排队消息',0,1,1,0,Null,0,0,'',0,Null,Sysdate,Null,'' From Dual
Union All Select '700CB92AC6D9428C9CEEAEE11B1EEA8C','HIS','ZLHIS_QUEUE_002','队列完成消息',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select '53E776624FA34A9D982C26AF7E705A98','HIS','ZLHIS_QUEUE_003','队列状态同步',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select '6BEB41BF1C054F41A839BD547A668426','HIS','ZLHIS_QUEUE_004','队列语音播放',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select 'D02FD471E85E4540B6D90E7AE1FCE922','HIS','ZLHIS_REGIST_001','门诊患者挂号',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select '36926C0DF2CA4C8390A469A338F1ABF3','HIS','ZLHIS_REGIST_002','门诊分诊通知',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select 'AC0482630353403DBA1D6A0B5B69A09A','HIS','ZLHIS_TRANSFUSION_001','输液审核未通过',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual;

--zlmip_item_field
Insert Into zlmip_item_field(item_id,fld_order,fld_title,fld_type)
          Select '39643F2E01CB4249B282CDB1CB278AB7',1,'ID',1 From Dual
Union All Select '39643F2E01CB4249B282CDB1CB278AB7',2,'编号',2 From Dual
Union All Select '39643F2E01CB4249B282CDB1CB278AB7',3,'姓名',2 From Dual
Union All Select '39643F2E01CB4249B282CDB1CB278AB7',4,'性别',2 From Dual
Union All Select '39643F2E01CB4249B282CDB1CB278AB7',5,'民族',2 From Dual
Union All Select '39643F2E01CB4249B282CDB1CB278AB7',6,'称呼',2 From Dual
Union All Select '39643F2E01CB4249B282CDB1CB278AB7',7,'移动电话',2 From Dual
Union All Select '39643F2E01CB4249B282CDB1CB278AB7',8,'电子邮件',2 From Dual
Union All Select '39643F2E01CB4249B282CDB1CB278AB7',9,'出生日期',2 From Dual;

--zlmip_item_deliver
Insert Into zlmip_item_deliver(id,item_id,deliver_order,deliver_code,deliver_title,deliver_object)
          Select '409B0EBEB96745CAA1F2023479F11F65','29C186B7661F4202BC60AB7319498FF9',1,'','能审核销帐申请的模块',xmltype('<deliverobject><depts><message>/cancel_reqeust/cancel_charge/audit_dept_id</message></depts><modules><module><title>1133-住院记帐管理</title><key>100_1133_基本</key></module><module><title>1134-科室分散记帐</title><key>100_1134_基本</key></module><module><title>1135-医技科室记帐</title><key>100_1135_基本</key></module><module><title>1139-病人费用查询</title><key>100_1139_基本</key></module><module><title>1342-药品部门发药</title><key>100_1342_基本</key></module><module><title>1262-住院护士工作站</title><key>100_1262_基本</key></module><module><title>1265-新版住院护士工作站</title><key>100_1265_基本</key></module><module><title>1263-医技工作站</title><key>100_1263_基本</key></module></modules></deliverobject>                                                                   ') From Dual
Union All Select 'FC4C67DB30244D54B3D69B418714B714','2755C4E8A9324E33B29A08A0B2A0E480',1,'01','检查系统的医技、采集、病理工作站(已收费)',xmltype('<deliverobject><depts><message>/charge_bill/bill_item/execute_dept_id</message></depts><modules><module><title>1290-影像医技工作站</title><key>100_1290_基本</key></module><module><title>1291-影像采集工作站</title><key>100_1291_基本</key></module><module><title>1294-影像病理工作站</title><key>100_1294_基本</key></module></modules></deliverobject>                           ') From Dual
Union All Select '52A94003E8AD4218A51A5ADA5435E036','2755C4E8A9324E33B29A08A0B2A0E480',2,'02','标准版的药品处方发药(药品)',xmltype('<deliverobject><depts><message>/charge_bill/bill_item/execute_dept_id</message></depts><modules><module><title>1341-药品处方发药</title><key>100_1341_基本</key></module></modules></deliverobject>        ') From Dual
Union All Select '91A7E35D3BDA43E093AA05A72C959DE4','96D430F4877D406FA40672E641E323E7',1,'','临床的护士工作站',xmltype('<deliverobject><depts><message>/patient_clinic/clinic_area_id</message></depts><modules><module><title>1262-住院护士工作站</title><key>100_1262_基本</key></module><module><title>1265-新版住院护士工作站</title><key>100_1265_基本</key></module></modules></deliverobject>                    ') From Dual
Union All Select '458D5AAAB45F470E9533F9D52B593A59','D95BC90DB297434791792921393A5CD9',1,'','临床的护士工作站',xmltype('<deliverobject><depts><message>/patient_clinic/clinic_area_id</message></depts><modules><module><title>1262-住院护士工作站</title><key>100_1262_基本</key></module><module><title>1265-新版住院护士工作站</title><key>100_1265_基本</key></module></modules></deliverobject>                    ') From Dual
Union All Select 'C8EABDA8199A437A98AEF743C38D9D97','E0147BE1C6A640CEAC595C519490AD7B',1,'01','临床的护士站(住院患者)',xmltype('<deliverobject><depts><message>/patient_clinic/clinic_area_id</message></depts><modules><module><title>1262-住院护士工作站</title><key>100_1262_基本</key></module><module><title>1265-新版住院护士工作站</title><key>100_1265_基本</key></module></modules></deliverobject>                    ') From Dual
Union All Select 'DBA5DFE62E5046E8A02EB8A510BA1C21','E0147BE1C6A640CEAC595C519490AD7B',2,'02','输液系统的配置中心管理(住病患者的输液)',xmltype('<deliverobject><depts><message>/cancel_order/execute_dept_id</message></depts><modules><module><title>1345-输液配置中心管理</title><key>100_1345_基本</key></module></modules></deliverobject>          ') From Dual
Union All Select 'D68E77243FD049FF88D6BB5CC50BD3AF','E0147BE1C6A640CEAC595C519490AD7B',3,'03','检查系统的影像医技工作站',xmltype('<deliverobject><depts><message>/cancel_order/execute_dept_id</message></depts><modules><module><title>1290-影像医技工作站</title><key>100_1290_基本</key></module></modules></deliverobject>         ') From Dual
Union All Select 'BBCC7CD75BB34278A02A84EC55F3E34E','C6E6C17F67C945C1903D23FD2F68BB87',1,'','临床的医技工作站',xmltype('<deliverobject><depts><message>/arrange_order/execute_dept_id</message></depts><modules><module><title>1263-医技工作站</title><key>100_1263_基本</key></module></modules></deliverobject>       ') From Dual
Union All Select '18C0EFCDDA8F482DA724800B651FB63D','73A421D7B7E44815A1BA40170BC2F509',1,'','临床的医生和护士站',xmltype('<deliverobject><depts><message>/patient_clinic/clinic_dept_id</message></depts><modules><module><title>1261-住院医生工作站</title><key>100_1261_基本</key></module><module><title>1262-住院护士工作站</title><key>100_1262_基本</key></module><module><title>1265-新版住院护士工作站</title><key>100_1265_基本</key></module></modules></deliverobject>                             ') From Dual
Union All Select '3030EE2A8B7648D3A9AE9E7087640EBB','F36F8875F1E540908E8DC62BA3C42168',1,'','标准版的药品处方发药',xmltype('<deliverobject><depts><message>/charge_bill/send_order/execute_dept_id</message></depts><modules><module><title>1341-药品处方发药</title><key>100_1341_基本</key></module></modules></deliverobject>        ') From Dual
Union All Select 'CF828B078E01474683B1A129C25CE23C','D2F8C10D32AA48E19A485787F25C70A8',1,'','临床的门诊医生站',xmltype('<deliverobject><depts><message>/change_clinic/change_dept_id</message></depts><modules><module><title>1260-门诊医生工作站</title><key>100_1260_基本</key></module></modules></deliverobject>         ') From Dual
Union All Select 'BFEA0B6404F54D15931812A5A1FD0238','D862BD945CF74CA49C8110A1A2F1FDA5',1,'','输液的配置中心管理',xmltype('<deliverobject><modules><module><title>1345-输液配置中心管理</title><key>100_1345_基本</key></module></modules></deliverobject>') From Dual
Union All Select '73922944F497408EBF4BC37F714DC11F','B0C04DB87C284EB38B49141479B904A1',1,'','病历系统的电子病历管理',xmltype('<deliverobject><modules><module><title>2250-电子病历管理</title><key>100_2250_基本</key></module></modules></deliverobject>') From Dual
Union All Select '6238E3C5DBD6413782CA33C11CD746AA','F02900C556BB4008A9DC046B1365FC26',1,'','病历系统的电子病历管理',xmltype('<deliverobject><depts><message>/patient_clinic/clinic_dept_id</message></depts><modules><module><title>2250-电子病历管理</title><key>100_2250_基本</key></module></modules></deliverobject>        ') From Dual
Union All Select 'F1322BBB3032487D9770289CD4059938','150A63E2931B4A309D65474C3E8ED00B',1,'','病历系统的电子病历管理',xmltype('<deliverobject><depts><message>/patient_clinic/clinic_dept_id</message></depts><modules><module><title>2250-电子病历管理</title><key>100_2250_基本</key></module></modules></deliverobject>        ') From Dual
Union All Select '195FB417251944A4A73BAD2F3799A6BE','78C19BDBF65D4784B3EC8A12B2489C61',1,'','输液系统的配置中心管理',xmltype('<deliverobject><depts><message>/cancel_reqeust/audit_dept_id</message></depts><modules><module><title>1345-输液配置中心管理</title><key>100_1345_基本</key></module></modules></deliverobject>          ') From Dual
Union All Select 'F6940EC364C9437E8D13BA2CC5EA77B1','8DF1D60F9642408FA984C1FB423C8E74',1,'','检验系统的技师工作站',xmltype('<deliverobject><modules><module><title>1208-检验技师工作站</title><key>100_1208_基本</key></module><module><title>2500-临床实验室管理</title><key>2500_2500_基本</key></module></modules></deliverobject>') From Dual
Union All Select 'DB24FBE2EAD842968000AFC88D98211B','10A0A36A7AE84876926575D5E67E2465',1,'','临床的医生和护士站',xmltype('<deliverobject><depts><message>/patient_clinic/clinic_area_id</message><message>/patient_clinic/clinic_dept_id</message></depts><modules><module><title>1261-住院医生工作站</title><key>100_1261_基本</key></module><module><title>1262-住院护士工作站</title><key>100_1262_基本</key></module><module><title>1265-新版住院护士工作站</title><key>100_1265_基本</key></module></modules></deliverobject>                             ') From Dual
Union All Select '70EDEC04842946D5A4B8F36056033C5D','D31F0420F60049B3837552290EB46703',1,'','检查系统的影像医技工作站',xmltype('<deliverobject><depts><message>/check_request/execute_dept_id</message></depts><modules><module><title>1290-影像医技工作站</title><key>100_1290_基本</key></module><module><title>1291-影像采集工作站</title><key>100_1291_基本</key></module><module><title>1294-影像病理工作站</title><key>100_1294_基本</key></module></modules></deliverobject>') From Dual
Union All Select '5D41FB8AB0944A3E846E8ABB5BDB13A3','C7CE07FDB5594C2DB11D50D1DD0BB795',1,'','手麻系统的手麻安排管理',xmltype('<deliverobject><depts><message>/oper_request/execute_dept_id</message></depts><modules><module><title>2409-手麻安排管理</title><key>2400_2409_基本</key></module></modules></deliverobject>        ') From Dual
Union All Select 'D80156CEE94349D49CC9D3E34F71D614','6082481F1E47484EB7D93ADCB2C94A13',1,'02','病历系统的电子病历管理',xmltype('<deliverobject><depts><message>/patient_clinic/clinic_dept_id</message></depts><modules><module><title>2250-电子病历管理</title><key>100_2250_基本</key></module></modules></deliverobject>') From Dual
Union All Select 'C9C571355FE245A896AE36D169707540','6082481F1E47484EB7D93ADCB2C94A13',2,'01','检查的医技、采集、病理工作站',xmltype('<deliverobject><modules><module><title>1290-影像医技工作站</title><key>100_1290_基本</key></module><module><title>1291-影像采集工作站</title><key>100_1291_基本</key></module><module><title>1294-影像病理工作站</title><key>100_1294_基本</key></module></modules></deliverobject>') From Dual
Union All Select 'AADF3E9D4DCA43C5BC08F754147C9777','E52632D77EB843F1BF16F621855C3F66',1,'','检查系统的影像技师工作站',xmltype('<deliverobject><depts><message>/patient_clinic/clinic_dept_id</message></depts><modules><module><title>1290-影像医技工作站</title><key>100_1290_基本</key></module></modules></deliverobject>         ') From Dual
Union All Select '6A78D3EEF18B4EA0A91ADF004F589821','D3D2A3A3EBD549CFB8761AA53BC5A9E9',1,'','检查系统的采集、医技、病理工作站',xmltype('<deliverobject><depts><message>/advice_info/study_execute_id</message></depts><modules><module><title>1291-影像采集工作站</title><key>100_1291_基本</key></module><module><title>1290-影像医技工作站</title><key>100_1290_基本</key></module><module><title>1294-影像病理工作站</title><key>100_1294_基本</key></module></modules></deliverobject>') From Dual
Union All Select '0BE44353BC5C448289AF37A9BC0373D0','98125826A65144D5A8330ECAC56F225B',1,'','检查系统的医技和采集工作站',xmltype('<deliverobject><depts><message>/study_state/study_execute_id</message></depts><modules><module><title>1290-影像医技工作站</title><key>100_1290_基本</key></module><module><title>1291-影像采集工作站</title><key>100_1291_基本</key></module><module><title>1294-影像病理工作站</title><key>100_1294_基本</key></module></modules></deliverobject>') From Dual
Union All Select '957D4132997D421887C157820F3EC151','92F5E98A99A742299CF0FA52D98B0417',1,'ZLHIS_PACS_003.01','检查系统的医技和采集工作站',xmltype('<deliverobject><depts><message>/study_state/study_execute_id</message></depts><modules><module><title>1290-影像医技工作站</title><key>100_1290_基本</key></module><module><title>1291-影像采集工作站</title><key>100_1291_基本</key></module><module><title>1294-影像病理工作站</title><key>100_1294_基本</key></module></modules></deliverobject>') From Dual
Union All Select '31DFF5AD2F334E978D1EE132DCB8C8D5','92F5E98A99A742299CF0FA52D98B0417',2,'ZLHIS_PACS_003.02','临床的住院医生工作站',xmltype('<deliverobject><modules><module><title>1261-住院医生工作站</title><key>100_1261_基本</key></module></modules></deliverobject>') From Dual
Union All Select '02F77EB636834599BAEFCA977D7EBFA5','EB810C78495C46128FD31111C382407C',1,'','检查系统的医技和采集工作站',xmltype('<deliverobject><depts><message>/advice_info/study_execute_id</message></depts><modules><module><title>1290-影像医技工作站</title><key>100_1290_基本</key></module><module><title>1291-影像采集工作站</title><key>100_1291_基本</key></module><module><title>1294-影像病理工作站</title><key>100_1294_基本</key></module></modules></deliverobject>') From Dual
Union All Select 'B224892D441A4BEBA4339406CC23D51D','50146DFA544E4FAC9EBC4851DC1A4B5B',1,'01','临床的医生工作站(住院病人)',xmltype('<deliverobject><depts><message>/patient_clinic/clinic_dept_id</message></depts><modules><module><title>1261-住院医生工作站</title><key>100_1261_基本</key></module><module><title>1262-住院护士工作站</title><key>100_1262_基本</key></module><module><title>1290-影像医技工作站</title><key>100_1290_基本</key></module><module><title>1291-影像采集工作站</title><key>100_1291_基本</key></module><module><title>1294-影像病理工作站</title><key>100_1294_基本</key></module></modules></deliverobject>') From Dual
Union All Select 'C285E72C49644CB2AB4C73EA297CB746','50146DFA544E4FAC9EBC4851DC1A4B5B',2,'02','住院患者的主任、主治、经治医师和责任护士',xmltype('<deliverobject><persons><message>/patient_clinic/in_doctor_id</message><message>/patient_clinic/director_doctor_id</message><message>/patient_clinic/treat_doctor_id</message><message>/patient_clinic/duty_nurse_id</message></persons></deliverobject>') From Dual
Union All Select 'CE1F550841A44220A2A152DAF6DECF70','F795028B5D8C4A9CAD49192A3A06FA5F',1,'01','病人入出管理及新版护士工作站',xmltype('<deliverobject><depts><message>/in_hospital/in_area_id</message><message>/in_hospital/in_dept_id</message></depts><modules><module><title>1132-病人入出管理</title><key>100_1132_基本</key></module><module><title>1265-新版住院护士工作站</title><key>100_1265_基本</key></module></modules></deliverobject>') From Dual
Union All Select '29CAF2DC70E4430C9FEE92555B928EDA','1373E96D23D44E1388F27F268718C694',1,'','住院医生站和住院护士站',xmltype('<deliverobject><depts><message>/in_hospital/in_area_id</message><message>/in_hospital/in_dept_id</message></depts><modules><module><title>1261-住院医生工作站</title><key>100_1261_基本</key></module><module><title>1262-住院护士工作站</title><key>100_1262_基本</key></module><module><title>1265-新版住院护士工作站</title><key>100_1265_基本</key></module></modules></deliverobject>                             ') From Dual
Union All Select 'C0FE290C8192463E871C1097DB67D0A1','B67FC7C97266418D82E47209D7C41C9B',1,'','标准版的病人入出管理和临床的护士站',xmltype('<deliverobject><depts><message>/change_state/change_area_id</message><message>/change_state/change_dept_id</message></depts><modules><module><title>1132-病人入出管理</title><key>100_1132_基本</key></module><module><title>1265-新版住院护士工作站</title><key>100_1265_基本</key></module><module><title>1262-住院护士工作站</title><key>100_1262_基本</key></module></modules></deliverobject>                            ') From Dual
Union All Select '9AAC03DADD624FEEA496A5E3E0F3D2AB','61505E3D721544B8974558CC3420C19C',1,'','病历系统的电子病历管理模块',xmltype('<deliverobject><modules><module><title>2250-电子病历管理</title><key>100_2250_基本</key></module></modules></deliverobject>') From Dual
Union All Select '0A07ADF6B23C461BB37BE848D4204B75','447C3A204ED94FF3ACC8CDBBAA3C0C42',1,'','病历系统的电子病历管理模块',xmltype('<deliverobject><modules><module><title>2250-电子病历管理</title><key>100_2250_基本</key></module></modules></deliverobject>') From Dual
Union All Select 'E8C535993E4C497F9340C4ABB707F393','3B793D544B914C72A958E1D7F8E8C247',1,'','电子病历/医生站/护士站',xmltype('<deliverobject><depts><message>/change_cancel/before_area_id</message><message>/change_cancel/before_dept_id</message><message>/change_cancel/after_area_id</message><message>/change_cancel/after_dept_id</message></depts><modules><module><title>2250-电子病历管理</title><key>100_2250_基本</key></module><module><title>1261-住院医生工作站</title><key>100_1261_基本</key></module><module><title>1262-住院护士工作站</title><key>100_1262_基本</key></module><module><title>1265-新版住院护士工作站</title><key>100_1265_基本</key></module></modules></deliverobject>                                                                          ') From Dual
Union All Select 'FF6C1C11A4FF44A391A9CBF452F5F587','20D31ACCAAC64CDCA242787B01C6D44E',1,'','病历系统的电子病历管理模块',xmltype('<deliverobject><modules><module><title>2250-电子病历管理</title><key>100_2250_基本</key></module></modules></deliverobject>') From Dual
Union All Select '74B66DACE0C7438987E5C06BD39692C5','7BD50BB1610E4404B6B2B41A6B245E93',1,'','临床的医生和护士站',xmltype('<deliverobject><depts><message>/out_prehospital/out_area_id</message><message>/out_prehospital/out_dept_id</message></depts><modules><module><title>1261-住院医生工作站</title><key>100_1261_基本</key></module><module><title>1262-住院护士工作站</title><key>100_1262_基本</key></module><module><title>1265-新版住院护士工作站</title><key>100_1265_基本</key></module></modules></deliverobject>                             ') From Dual
Union All Select 'C074E61723AD4B1F883EB6AF1BA09F96','12FE4EEC12964216A2790B2AFFC55BF0',1,'','临床的医生和护士站',xmltype('<deliverobject><depts><message>/out_hospital/out_area_id</message><message>/out_hospital/out_dept_id</message></depts><modules><module><title>1261-住院医生工作站</title><key>100_1261_基本</key></module><module><title>1262-住院护士工作站</title><key>100_1262_基本</key></module><module><title>1265-新版住院护士工作站</title><key>100_1265_基本</key></module></modules></deliverobject>                             ') From Dual
Union All Select 'AEDFDC629D9949BA81B6AFFA559D2EC6','9034EA214816475B85CB9C83E377AE1D',1,'','临床的医生和护士站',xmltype('<deliverobject><depts><message>/in_hospital/in_area_id</message><message>/in_hospital/in_dept_id</message></depts><modules><module><title>1261-住院医生工作站</title><key>100_1261_基本</key></module><module><title>1262-住院护士工作站</title><key>100_1262_基本</key></module><module><title>1265-新版住院护士工作站</title><key>100_1265_基本</key></module></modules></deliverobject>                             ') From Dual
Union All Select 'E2F5B3FF57BA4C3CA22A706375F98EF1','37F9BA5076D84A66B0DA03A8C293D42B',1,'','病历系统的电子病历管理',xmltype('<deliverobject><modules><module><title>2250-电子病历管理</title><key>100_2250_基本</key></module></modules></deliverobject>') From Dual
Union All Select '2F129D052B434B81AB4941F907D64043','39643F2E01CB4249B282CDB1CB278AB7',1,'','消息中的人员登录的工作站',xmltype('<deliverobject><persons><message>/person/person_id</message></persons></deliverobject>') From Dual
Union All Select '36932C7CD13347499680F0A675A42AA2','81B4204CC0B44877A3DF714BB13773BD',1,'01','检查的医技和采集工作站',xmltype('<deliverobject><modules><module><title>1290-影像医技工作站</title><key>100_1290_基本</key></module><module><title>1291-影像采集工作站</title><key>100_1291_基本</key></module><module><title>1160-排队叫号操作</title><key>100_1160_基本</key></module></modules></deliverobject>') From Dual
Union All Select 'ADD8A8B6EF33426C8FA3E470590A6773','700CB92AC6D9428C9CEEAEE11B1EEA8C',1,'01','检查的医技和采集工作站',xmltype('<deliverobject><modules><module><title>1290-影像医技工作站</title><key>100_1290_基本</key></module><module><title>1291-影像采集工作站</title><key>100_1291_基本</key></module><module><title>1160-排队叫号操作</title><key>100_1160_基本</key></module></modules></deliverobject>') From Dual
Union All Select '507CF75C6B6844F79DA37F4CF69F7EA2','53E776624FA34A9D982C26AF7E705A98',1,'01','检查的医技和采集工作站',xmltype('<deliverobject><modules><module><title>1290-影像医技工作站</title><key>100_1290_基本</key></module><module><title>1291-影像采集工作站</title><key>100_1291_基本</key></module><module><title>1160-排队叫号操作</title><key>100_1160_基本</key></module></modules></deliverobject>') From Dual
Union All Select '425C03224E9342BB88DA0D920EB433CB','6BEB41BF1C054F41A839BD547A668426',1,'','排队叫号操作模块',xmltype('<deliverobject><stations><message>/voice_station_ip</message></stations><modules><module><title>1160-排队叫号操作</title><key>100_1160_基本</key></module></modules></deliverobject>        ') From Dual
Union All Select '4F2205902DAB40A3BAB52574BEAC4CB6','D02FD471E85E4540B6D90E7AE1FCE922',1,'01','标准版的门诊分诊模块',xmltype('<deliverobject><modules><module><title>1113-门诊分诊管理</title><key>100_1113_基本</key></module></modules></deliverobject>') From Dual
Union All Select 'B615068542A84794AEE4F66C924497CE','D02FD471E85E4540B6D90E7AE1FCE922',2,'02','标准版的门诊医生站模块',xmltype('<deliverobject><depts><message>/register_info/register_dept_id</message></depts><modules><module><title>1260-门诊医生工作站</title><key>100_1260_基本</key></module></modules></deliverobject>         ') From Dual
Union All Select '8F92995E8D584518B094A08D1CB6BA64','36926C0DF2CA4C8390A469A338F1ABF3',1,'','标准版的门诊医生工作站模块',xmltype('<deliverobject><depts><message>/triage_info/register_dept_id</message></depts><modules><module><title>1260-门诊医生工作站</title><key>100_1260_基本</key></module></modules></deliverobject>         ') From Dual
Union All Select '5E0C8D480CCE4D9998BEE0C925145670','AC0482630353403DBA1D6A0B5B69A09A',1,'','病人病区对应的护士工作站',xmltype('<deliverobject><depts><message>/transfuse_order/patient_clinic/clinic_area_id</message></depts><modules><module><title>1262-住院护士工作站</title><key>100_1262_基本</key></module><module><title>1265-新版住院护士工作站</title><key>100_1265_基本</key></module></modules></deliverobject>') From Dual;

--zlmip_item_config
Insert Into zlmip_item_config(id,parent_id,item_id,node_order,node_type,node_title,data_type,min_occurs,max_occurs,config_occurs,config_occurs_key,config_express,config_express_key,config_note)
          Select '938880FABF0D4FE3B3FF339D997B3A00','','39643F2E01CB4249B282CDB1CB278AB7',1,2,'topic','xs:string','0','1','1','','生日祝福','生日祝福','' From Dual
Union All Select '1B822D4CD3394481B8DD6C5462FD84E0','','39643F2E01CB4249B282CDB1CB278AB7',2,2,'text','xs:string','1','1','1','','[T.姓名][T.称呼]：今天是您的生日，祝您身体健康，生日快乐！','[T.姓名][T.称呼]：今天是您的生日，祝您身体健康，生日快乐！','' From Dual
Union All Select '092913CAD52840D5BE57345B9EBB5754','','39643F2E01CB4249B282CDB1CB278AB7',3,1,'person','','1','1','','','','','' From Dual
Union All Select '4E95117BA6374359A117895D216CD7AC','092913CAD52840D5BE57345B9EBB5754','39643F2E01CB4249B282CDB1CB278AB7',4,2,'person_id','xs:integer','1','1','1','','[T.ID]','[T.ID]','' From Dual
Union All Select '4252A4341A8E45A28F9FCD737DB1B0AA','092913CAD52840D5BE57345B9EBB5754','39643F2E01CB4249B282CDB1CB278AB7',5,2,'person_name','xs:string','1','1','1','','[T.姓名]','[T.姓名]','' From Dual
Union All Select 'EC127AAEDA0440638189B60B8C0D188B','092913CAD52840D5BE57345B9EBB5754','39643F2E01CB4249B282CDB1CB278AB7',6,2,'person_sex','xs:string','1','1','1','','[T.性别]','[T.性别]','' From Dual
Union All Select 'F71F01B3EF5C435488BE7F6D89230868','092913CAD52840D5BE57345B9EBB5754','39643F2E01CB4249B282CDB1CB278AB7',7,2,'person_birth','xs:date','1','1','1','','[T.出生日期]','[T.出生日期]','' From Dual
Union All Select 'E83DD1E4E6EE446F8460A5DBD47EC775','092913CAD52840D5BE57345B9EBB5754','39643F2E01CB4249B282CDB1CB278AB7',8,2,'person_email','xs:string','0','1','1','','[T.电子邮件]','[T.电子邮件]','' From Dual
Union All Select '642C12A08B5741AEBF0A8B033A203E15','092913CAD52840D5BE57345B9EBB5754','39643F2E01CB4249B282CDB1CB278AB7',9,2,'person_mobiletelephone','xs:string','0','1','1','','[T.移动电话]','[T.移动电话]','' From Dual;
