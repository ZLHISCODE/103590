
--zlmip_table

--zlmip_tab_field

--zlmip_tab_parameter

--zlmip_tab_extend

--zlmip_tabext_condition

--zlmip_item
Insert Into zlmip_item(id,data_code,item_code,item_title,item_request,item_type,trigger_type,again_policy,again_para,check_frequency,check_freq_internal,trigger_condition,trigger_frequency,tab_id,start_date,stop_date,item_note)
          Select '74035DABA8C3411A9D0D629B4934079F','LIS','ZLHIS_LIS_001','检验标本审核',0,1,1,0,Null,0,0,'',0,Null,Sysdate,Null,'' From Dual
Union All Select '34FD2A378F3E43A19FD3A4F86B580AC9','LIS','ZLHIS_LIS_002','检验标本审核撤消',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select '6CA39650E3984AD88D7B1DA8EE033CCB','LIS','ZLHIS_LIS_003','检验危急值通知',0,1,1,0,Null,0,0,'',0,Null,Sysdate,Null,'' From Dual;

--zlmip_item_field

--zlmip_item_deliver
Insert Into zlmip_item_deliver(id,item_id,deliver_order,deliver_code,deliver_title,deliver_object)
          Select '95C39771B55C45F283B0A91A2A815D71','34FD2A378F3E43A19FD3A4F86B580AC9',1,'','临床的医生工作站和护士工作站(住院病人)',xmltype('<deliverobject><modules><module><title>1261-住院医生工作站</title><key>100_1261_基本</key></module></modules></deliverobject>') From Dual
Union All Select '62F7FA19A60C42508E6C9E075A8236F8','6CA39650E3984AD88D7B1DA8EE033CCB',1,'01','临床的医生工作站和护士工作站(住院病人)',xmltype('<deliverobject><modules><module><title>1261-住院医生工作站</title><key>100_1261_基本</key></module><module><title>1262-住院护士工作站</title><key>100_1262_基本</key></module><module><title>1265-新版住院护士工作站</title><key>100_1265_基本</key></module></modules></deliverobject>') From Dual
Union All Select '20322AC0052A45CA95DFA60A331913F6','6CA39650E3984AD88D7B1DA8EE033CCB',2,'02','住院患者的主任、主治、经治医师和责任护士',xmltype('<deliverobject><persons><message>/patient_clinic/in_doctor_id</message><message>/patient_clinic/director_doctor_id</message><message>/patient_clinic/treat_doctor_id</message><message>/patient_clinic/duty_nurse_id</message></persons></deliverobject>') From Dual;

--zlmip_item_config
