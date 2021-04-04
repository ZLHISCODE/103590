
--zlmip_table

--zlmip_tab_field

--zlmip_tab_parameter

--zlmip_tab_extend

--zlmip_tabext_condition

--zlmip_item
Insert Into zlmip_item(id,data_code,item_code,item_title,item_request,item_type,trigger_type,again_policy,again_para,check_frequency,check_freq_internal,trigger_condition,trigger_frequency,tab_id,start_date,stop_date,item_note)
          Select '1EDC4693D289498BBB4C6324FF6DFC49','','ZLHIS_PUB_002','邮件消息',0,1,1,0,Null,0,0,'',0,Null,Sysdate,Null,'' From Dual
Union All Select '20E0AB9BFCB841DBB59BC1D507F7EB4B','','ZLHIS_PUB_003','短信消息',0,1,1,0,Null,0,0,'',0,Null,Sysdate,Null,'' From Dual
Union All Select 'EE2BBEEB945F45A18DE30B88D050689B','','ZLHIS_PUB_004','通用消息',0,1,1,0,Null,0,0,'',0,Null,Sysdate,Null,'' From Dual
Union All Select '670446972A2B48049E735D079AA06FDE','','ZLHIS_PUB_005','产品预升级通知',0,1,2,0,Null,0,0,'Select Distinct 工作站, Ip, To_Char(预升时点, ''yyyy-mm-dd hh24:mi:ss'') As 预升时点
From zlClients
Where 升级标志 = 1 And 预升完成 = 0 And 预升时点 <= Sysdate',0,Null,Sysdate,Null,'' From Dual;

--zlmip_item_field
Insert Into zlmip_item_field(item_id,fld_order,fld_title,fld_type)
          Select '670446972A2B48049E735D079AA06FDE',1,'工作站',2 From Dual
Union All Select '670446972A2B48049E735D079AA06FDE',2,'IP',2 From Dual
Union All Select '670446972A2B48049E735D079AA06FDE',3,'预升时点',2 From Dual;

--zlmip_item_deliver
Insert Into zlmip_item_deliver(id,item_id,deliver_order,deliver_code,deliver_title,deliver_object)
          Select '6391D440E09A48F79C42C4ED7FB4EC9D','670446972A2B48049E735D079AA06FDE',1,'','需要升级的工作站',xmltype('<deliverobject><stations><message>/station_ip</message></stations></deliverobject>') From Dual;

--zlmip_item_config
Insert Into zlmip_item_config(id,parent_id,item_id,node_order,node_type,node_title,data_type,min_occurs,max_occurs,config_occurs,config_occurs_key,config_express,config_express_key,config_note)
          Select '07249D7784A74CCB82663231AC33503B','','670446972A2B48049E735D079AA06FDE',1,2,'station_title','xs:string','1','1','1','','[T.工作站]','[T.工作站]','' From Dual
Union All Select '60D46C8E9BBF438F99FBF96FE65F92BD','','670446972A2B48049E735D079AA06FDE',2,2,'station_ip','xs:string','1','1','1','','[T.IP]','[T.IP]','' From Dual
Union All Select '66E987573E6F43AB918655FE99DC9A1B','','670446972A2B48049E735D079AA06FDE',3,2,'upgrade_time','xs:string','1','1','1','','[T.预升时点]','[T.预升时点]','' From Dual;