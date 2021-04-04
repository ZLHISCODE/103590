
--zlmip_table

--zlmip_tab_field

--zlmip_tab_parameter

--zlmip_tab_extend

--zlmip_tabext_condition

--zlmip_item
Insert Into zlmip_item(id,data_code,item_code,item_title,item_request,item_type,trigger_type,again_policy,again_para,check_frequency,check_freq_internal,trigger_condition,trigger_frequency,tab_id,start_date,stop_date,item_note)
          Select '4AD940B8DDE045ACB0FD2C7AB0DC946A','EMR','ZLHIS_EMR_001','门诊患者接诊',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select '8453E9213E1348D98830A357FD63A2A5','EMR','ZLHIS_EMR_002','住院患者入院入科',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select 'ACA0FBE8D84340B78B4FA281A555B544','EMR','ZLHIS_EMR_003','住院患者转入科室',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select '39AB364D585C42DD9F884D5F1D1A5BF5','EMR','ZLHIS_EMR_004','住院患者床位变更',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select 'E85FDF7CD7824042933C32874AEB2135','EMR','ZLHIS_EMR_005','住院患者病情变更',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select '40F0D824400D4FDBB21F268F70A0B989','EMR','ZLHIS_EMR_006','住院患者医护变更',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select 'DE763B3131834FF6BBAA7E308ED7ACD9','EMR','ZLHIS_EMR_007','住院患者新生儿登记',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select '66FF97129D4245DA9B41799ED6A289F2','EMR','ZLHIS_EMR_008','住院患者会诊',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select 'F5D3A5D54BBA4DD5A242923C9DC2EEDC','EMR','ZLHIS_EMR_009','住院患者死亡',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select 'B7DD6D9698D648BB8675D12F9C305C7F','EMR','ZLHIS_EMR_010','住院患者抢救',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select '3091787D600D4B469BC5FF9C7FF1EF58','EMR','ZLHIS_EMR_011','住院患者输血',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select '151A5B5B6F324B5098C0D05D71D07DFB','EMR','ZLHIS_EMR_012','住院患者麻醉',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select '48F5EAAEB26B44A5ACDB976F3789F139','EMR','ZLHIS_EMR_013','住院患者手术',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select '7088BFC1D2FD485E96841B408EB1A450','EMR','ZLHIS_EMR_014','住院患者检查',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select '46C2CC5CE5AB444FB7A58C189F5993D7','EMR','ZLHIS_EMR_015','住院患者检验',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select '86D3A302743049499FBA0C91F1480BD0','EMR','ZLHIS_EMR_016','住院患者治疗',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select 'ED460FC9C20F4844B669C2255F2CA32C','EMR','ZLHIS_EMR_017','住院患者出院',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select '2734519190E34265B887E7C7A4D5F0A6','EMR','ZLHIS_EMR_018','住院患者门诊诊断',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select '28B6E931627A47C5B08B1ACE4E78EC99','EMR','ZLHIS_EMR_019','住院患者入院诊断',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select '4D7E9FD0FF0A4DAEAF07B18A0F04638E','EMR','ZLHIS_EMR_020','住院患者出院诊断',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual
Union All Select '310BDCA3A61148AFBFD1D3902B53EE20','EMR','ZLHIS_EMR_021','病历审阅提醒',0,1,1,0,Null,Null,Null,'',Null,Null,Sysdate,Null,'' From Dual;

--zlmip_item_field

--zlmip_item_deliver
Insert Into zlmip_item_deliver(id,item_id,deliver_order,deliver_code,deliver_title,deliver_object)
          Select '8FDDED756BED4F10A3488953A1464E94','4AD940B8DDE045ACB0FD2C7AB0DC946A',1,'','病历系统的电子病历管理',xmltype('<deliverobject><depts><message>/receive_dept_id</message></depts><modules><module><title>2250-电子病历管理</title><key>100_2250_基本</key></module></modules></deliverobject>                ') From Dual
Union All Select '9B2626EF43A84DCFAE62355A0C45C4C4','8453E9213E1348D98830A357FD63A2A5',1,'','电子病历管理',xmltype('<deliverobject><depts><message>/in_dept_id</message></depts><modules><module><title>2250-电子病历管理</title><key>100_2250_基本</key></module></modules></deliverobject>                ') From Dual
Union All Select '13C3AAD3F2B34217B5131B00DB22F6BE','ACA0FBE8D84340B78B4FA281A555B544',1,'','病历系统的电子病历管理',xmltype('<deliverobject><depts><message>/in_dept_id</message></depts><modules><module><title>2250-电子病历管理</title><key>100_2250_基本</key></module></modules></deliverobject>                ') From Dual
Union All Select 'C5EE5D0004654262952F0E5DC2006B19','39AB364D585C42DD9F884D5F1D1A5BF5',1,'','病历系统的电子病历管理',xmltype('<deliverobject><depts><message>/current_dept_id</message></depts><modules><module><title>2250-电子病历管理</title><key>100_2250_基本</key></module></modules></deliverobject>                ') From Dual
Union All Select '41594AC4C5B64BA3AE33A2E3B898B014','E85FDF7CD7824042933C32874AEB2135',1,'','病历系统的电子病历管理',xmltype('<deliverobject><depts><message>/clinic_dept_id</message></depts><modules><module><title>2250-电子病历管理</title><key>100_2250_基本</key></module></modules></deliverobject>                ') From Dual
Union All Select '60783E1D23914AC3A04E6D2DC22CCBD8','40F0D824400D4FDBB21F268F70A0B989',1,'','病历系统的电子病历管理',xmltype('<deliverobject><depts><message>/clinic_dept_id</message></depts><modules><module><title>2250-电子病历管理</title><key>100_2250_基本</key></module></modules></deliverobject>                ') From Dual
Union All Select 'F0274131D4714F21AF14F109A0EE2503','DE763B3131834FF6BBAA7E308ED7ACD9',1,'','病历系统的电子病历管理',xmltype('<deliverobject><depts><message>/clinic_dept_id</message></depts><modules><module><title>2250-电子病历管理</title><key>100_2250_基本</key></module></modules></deliverobject>                ') From Dual
Union All Select 'D7AAC75AAD9A4F16B5AE62656E6C14E6','66FF97129D4245DA9B41799ED6A289F2',1,'','病历系统的电子病历管理',xmltype('<deliverobject><depts><message>/clinic_dept_id</message></depts><modules><module><title>2250-电子病历管理</title><key>100_2250_基本</key></module></modules></deliverobject>                ') From Dual
Union All Select 'F109AE160C1046B19530DED692573FFE','F5D3A5D54BBA4DD5A242923C9DC2EEDC',1,'','病历系统的电子病历管理',xmltype('<deliverobject><depts><message>/clinic_dept_id</message></depts><modules><module><title>2250-电子病历管理</title><key>100_2250_基本</key></module></modules></deliverobject>                ') From Dual
Union All Select 'F6ADAC6C8ABF4E41B72EE53E18CE7853','B7DD6D9698D648BB8675D12F9C305C7F',1,'','病历系统的电子病历管理',xmltype('<deliverobject><depts><message>/clinic_dept_id</message></depts><modules><module><title>2250-电子病历管理</title><key>100_2250_基本</key></module></modules></deliverobject>                ') From Dual
Union All Select 'F9B2714E38E54DC994A2672717C4792F','3091787D600D4B469BC5FF9C7FF1EF58',1,'','病历系统的电子病历管理',xmltype('<deliverobject><depts><message>/clinic_dept_id</message></depts><modules><module><title>2250-电子病历管理</title><key>100_2250_基本</key></module></modules></deliverobject>                ') From Dual
Union All Select '9BA6000307994197A1BB820F55037DE3','151A5B5B6F324B5098C0D05D71D07DFB',1,'','病历系统的电子病历管理',xmltype('<deliverobject><depts><message>/clinic_dept_id</message></depts><modules><module><title>2250-电子病历管理</title><key>100_2250_基本</key></module></modules></deliverobject>                ') From Dual
Union All Select '08D5B19BA28F413DBD214F8EDA787478','48F5EAAEB26B44A5ACDB976F3789F139',1,'','病历系统的电子病历管理',xmltype('<deliverobject><depts><message>/clinic_dept_id</message></depts><modules><module><title>2250-电子病历管理</title><key>100_2250_基本</key></module></modules></deliverobject>                ') From Dual
Union All Select 'D5A882DA73404D90B91B7BDFEB740216','7088BFC1D2FD485E96841B408EB1A450',1,'','病历系统的电子病历管理',xmltype('<deliverobject><depts><message>/clinic_dept_id</message></depts><modules><module><title>2250-电子病历管理</title><key>100_2250_基本</key></module></modules></deliverobject>                ') From Dual
Union All Select 'AFF9FF7E13CF43E4AE4622F8B8819817','46C2CC5CE5AB444FB7A58C189F5993D7',1,'','病历系统的电子病历管理',xmltype('<deliverobject><depts><message>/clinic_dept_id</message></depts><modules><module><title>2250-电子病历管理</title><key>100_2250_基本</key></module></modules></deliverobject>                ') From Dual
Union All Select '5FAA13D1AF7F42439D0A41A58557AF39','86D3A302743049499FBA0C91F1480BD0',1,'','病历系统的电子病历管理',xmltype('<deliverobject><depts><message>/clinic_dept_id</message></depts><modules><module><title>2250-电子病历管理</title><key>100_2250_基本</key></module></modules></deliverobject>                ') From Dual
Union All Select '243CFBB3ED834FB69C5778B75F5EB237','ED460FC9C20F4844B669C2255F2CA32C',1,'','病历系统的电子病历管理',xmltype('<deliverobject><depts><message>/clinic_dept_id</message></depts><modules><module><title>2250-电子病历管理</title><key>100_2250_基本</key></module></modules></deliverobject>                ') From Dual
Union All Select '24FDE01FA7404C6F9EBB60FEC4B3DE8E','310BDCA3A61148AFBFD1D3902B53EE20',1,'','临床的医生工作站',xmltype('<deliverobject><persons><message>/patient_clinic/in_doctor_id</message><message>/patient_clinic/director_doctor_id</message><message>/patient_clinic/treat_doctor_id</message></persons><modules><module><title>1261-住院医生工作站</title><key>100_1261_基本</key></module></modules></deliverobject>                           ') From Dual;

--zlmip_item_config
