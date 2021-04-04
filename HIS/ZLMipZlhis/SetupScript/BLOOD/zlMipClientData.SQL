
--zlmip_table

--zlmip_tab_field

--zlmip_tab_parameter

--zlmip_tab_extend

--zlmip_tabext_condition

--zlmip_item
Insert Into zlmip_item(id,data_code,item_code,item_title,item_request,item_type,trigger_type,again_policy,again_para,check_frequency,check_freq_internal,trigger_condition,trigger_frequency,tab_id,start_date,stop_date,item_note)
          Select '69F5990487C74B72AC23673E74F43BA7','BLOOD','ZLHIS_BLOOD_001','科室配血完成',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual;

--zlmip_item_field

--zlmip_item_deliver
Insert Into zlmip_item_deliver(id,item_id,deliver_order,deliver_code,deliver_title,deliver_object)
          Select '253548CB1D804EAC87B9B4D870E510D5','69F5990487C74B72AC23673E74F43BA7',1,'','血库系统的科室发血管理模块',xmltype('<deliverobject><modules><module><title>1936-科室发血管理</title><key>2200_1936_基本</key></module></modules></deliverobject>') From Dual;

--zlmip_item_config
