
--zlmip_table

--zlmip_tab_field

--zlmip_tab_parameter

--zlmip_tab_extend

--zlmip_tabext_condition

--zlmip_item
Insert Into zlmip_item(id,data_code,item_code,item_title,item_request,item_type,trigger_type,again_policy,again_para,check_frequency,check_freq_internal,trigger_condition,trigger_frequency,tab_id,start_date,stop_date,item_note)
          Select '3C445D1CC7C947E38BB473E4A8925357','OPER','ZLHIS_OPER_001','手术安排完成',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select '59C13CDCA5C1472F9D563779AA48C2AE','OPER','ZLHIS_OPER_002','患者手术申请/撤消',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual;

--zlmip_item_field

--zlmip_item_deliver
Insert Into zlmip_item_deliver(id,item_id,deliver_order,deliver_code,deliver_title,deliver_object)
          Select '560564BBFF9C4436A1E985D74D78C92F','3C445D1CC7C947E38BB473E4A8925357',1,'','临床的住院医生工作站和护士工作站',xmltype('<deliverobject><depts><message>/patient_clinic/clinic_area_id</message><message>/patient_clinic/clinic_dept_id</message></depts><modules><module><title>1261-住院医生工作站</title><key>100_1261_基本</key></module><module><title>1262-住院护士工作站</title><key>100_1262_基本</key></module><module><title>1265-新版住院护士工作站</title><key>100_1265_基本</key></module></modules></deliverobject>                             ') From Dual
Union All Select 'A60BC3C8DDD849ABA701B9C632B8F9F7','3C445D1CC7C947E38BB473E4A8925357',2,'','移动消息客户端',xmltype('<deliverobject><mipusers><mipuser>ZLHIS</mipuser></mipusers></deliverobject>') From Dual;

--zlmip_item_config
