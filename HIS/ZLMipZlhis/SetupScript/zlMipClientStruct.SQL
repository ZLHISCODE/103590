-----------------------------------------------------------------------------------------------------------------------
--数据表部分
-----------------------------------------------------------------------------------------------------------------------
--zlmip_data_setup
Create Table zltools.zlmip_data_setup(
	data_code				varchar2(10),
	data_title				varchar2(30),	
	data_owner			varchar2(20),
	data_system			number(5),
	data_source			varchar2(200),
	data_db				varchar2(100),
	setup_time				date,
	upgrade_time			date);
	
--zlmip_table
Create Table zltools.zlmip_table(
	id					varchar2(32),
	data_code				varchar2(10),
	tab_type				number(3),
	tab_code				varchar2(30),
	tab_title				varchar2(30),
	tab_sqltext				varchar2(2000),
	tab_note				varchar2(200));

--zlmip_tab_field	
Create Table zltools.zlmip_tab_field(
	tab_id				varchar2(32),
	fld_order				number(5),
	fld_title				varchar2(30),
	fld_type				number(3));

--zlmip_tab_parameter	
Create Table zltools.zlmip_tab_parameter(
	tab_id				varchar2(32),
	para_order			number(5),
	para_field				varchar2(100),
	para_title				varchar2(1000),
	para_type				number(3),
	para_default			varchar2(100),
	para_note				varchar2(200));

--zlmip_tab_extend		
Create Table zltools.zlmip_tab_extend(
	id					varchar2(32),
	source_tab_id			varchar2(32),
	ext_order				number(5),
	ext_type				number(3),
	ext_title				varchar2(30),
	target_tab_id			varchar2(32));

--zlmip_tabext_condition		
Create Table zltools.zlmip_tabext_condition(
	ext_id				varchar2(32),
	cond_order			number(5),
	target_fld				varchar2(30),
	source_fld				varchar2(30));

--zlmip_item
Create Table zltools.zlmip_item(
	id					varchar2(32),
	data_code				varchar2(10),
	item_code				varchar2(30),
	item_title				varchar2(30),
	item_request			number(1),
	item_type				number(3),
	item_flow				varchar2(30),
	item_schema			varchar2(4000),
	trigger_type			number(3),
	check_frequency			number(3),
	check_freq_internal		number(3),
	trigger_condition		varchar2(4000),
	trigger_frequency		number(3),
	tab_id				varchar2(32),
	again_policy			number(3),
	again_para			varchar2(100),
	last_date				date,
	start_date				date,
	stop_date				date,	
	item_note				varchar2(200));

--zlmip_item_parameter
Create Table zltools.zlmip_item_parameter(
	item_id				varchar2(32),
	para_order			number(3),
	para_field				varchar2(100),
	para_title				varchar2(1000),
	para_type				number(3),
	para_default			varchar2(100),
	para_config			varchar2(100));

--zlmip_item_field	
Create Table zltools.zlmip_item_field(
	item_id				varchar2(32),
	fld_order				number(5),
	fld_title				varchar2(30),
	fld_type				number(3));

--zlmip_item_frequency
Create Table zltools.zlmip_item_frequency(
	item_id				varchar2(32),	
	freq_order				number(3),	
	freq_month			number(3),
	freq_day				number(3),
	freq_start				varchar2(10),
	freq_stop				varchar2(10));

--zlmip_item_config
Create Table zltools.zlmip_item_config(	
	id					varchar2(32),
	parent_id				varchar2(32),
	item_id				varchar2(32),
	node_order			number(5),
	node_type				number(3),
	node_title				varchar2(30),
	data_type				varchar2(10),
	min_occurs			varchar2(10),
	max_occurs			varchar2(10),
	config_occurs			varchar2(50),
	config_occurs_key		varchar2(50),
	config_express			varchar2(1000),
	config_express_key		varchar2(1000),
	config_note			varchar2(200));

--zlmip_item_deliver
Create Table zltools.zlmip_item_deliver(	
	id					varchar2(32),
	item_id				varchar2(32),
	deliver_order			number(3),
	deliver_code			varchar2(30),
	deliver_title			varchar2(50),
	deliver_object			xmltype);

--zlmip_send_log
Create Table zltools.zlmip_send_log(	
	id					varchar2(32),
	msg_id				varchar2(32),
	item_id				varchar2(32),
	item_identity			varchar2(32),
	trigger_date			date,
	send_content			xmltype,
	send_content_temp		clob,
	send_state				number(3),	
	send_date				date,	
	send_mipuser			varchar2(20),
	send_user				varchar2(20),
	send_username			varchar2(20),
	send_station			varchar2(100),
	send_module			varchar2(100),
	send_again			number(3),
	send_again_para		varchar2(20),
	send_again_number		number(18),
	send_again_date		date,
	send_note				varchar2(100));

--zlmip_sendlog_parameter
Create Table zltools.zlmip_sendlog_parameter(	
	send_log_id			varchar2(32),
	para_order			number(5),
	para_title				varchar2(100),
	para_value				varchar2(4000));

--zlmip_sendlog_again
Create Table zltools.zlmip_sendlog_again(	
	send_log_id			varchar2(32),
	send_order			number(5),
	send_state				number(3),
	send_date				date,
	send_mipuser			varchar2(20),
	send_user				varchar2(20),
	send_username			varchar2(20),
	send_station			varchar2(100),
	send_note				varchar2(100));

--zlmip_receive_log
Create Table zltools.zlmip_receive_log(	
	id					varchar2(32),
	msg_id				varchar2(32),
	item_identity			varchar2(32),
	receive_date			date,
	receive_mipuser		varchar2(20),
	receive_user			varchar2(20),
	receive_username		varchar2(20),
	receive_station			varchar2(30),
	receive_module			varchar2(1000),
	receive_content			xmltype,
	receive_content_temp		clob,
	receive_note			varchar2(100));

--zlmip_onlineusers
Create Table zltools.zlmip_onlineusers(
	session_id			number(18),
	station			varchar2(50),
	loginuser			varchar2(20));

--zlmip_onlinemodules
Create Table zltools.zlmip_onlinemodules(
	session_id				number(18),
	instance				varchar2(32),
	station				varchar2(50),
	system				number(5),
	module				number(5));

--zlmip_clientuser
Create Table zltools.zlmip_clientuser(
	station			varchar2(50),
	mipuser			Varchar2(20),
	mipuserpassword	varchar2(100));

-----------------------------------------------------------------------------------------------------------------------
--数据表约束部分
-----------------------------------------------------------------------------------------------------------------------
--zlmip_data_setup
Alter Table zltools.zlmip_data_setup Add Constraint zlmip_data_setup_PK Primary Key (data_code) Using Index;

--zlmip_table
Alter Table zltools.zlmip_table Add Constraint zlmip_table_PK Primary Key (ID) Using Index;
Alter Table zltools.zlmip_table Add Constraint zlmip_table_UQ_Code Unique (tab_code,tab_type) Using Index;
Alter Table zltools.zlmip_table Add Constraint zlmip_table_UQ_Title Unique (tab_code,tab_title) Using Index;

--zlmip_tab_field
Alter Table zltools.zlmip_tab_field Add Constraint zlmip_tab_field_PK Primary Key (tab_id,fld_order) Using Index;
Alter Table zltools.zlmip_tab_field Add Constraint zlmip_tab_field_FK_tabid Foreign Key (tab_id) References zltools.zlmip_table(ID) On Delete Cascade;
Alter Table zltools.zlmip_tab_field Add Constraint zlmip_tab_field_UQ_fldtitle Unique (tab_id,fld_title) Using Index;

--zlmip_tab_parameter
Alter Table zltools.zlmip_tab_parameter Add Constraint zlmip_tab_parameter_PK Primary Key (tab_id,para_order) Using Index;
Alter Table zltools.zlmip_tab_parameter Add Constraint zlmip_tab_parameter_FK_tabid Foreign Key (tab_id) References zltools.zlmip_table(ID) On Delete Cascade;
Alter Table zltools.zlmip_tab_parameter Add Constraint zlmip_tab_parameter_UQ_title Unique (tab_id,para_title) Using Index;

--zlmip_tab_extend
Alter Table zltools.zlmip_tab_extend Add Constraint zlmip_tab_extend_PK Primary Key (ID) Using Index;
Alter Table zltools.zlmip_tab_extend Add Constraint zlmip_tab_extend_FK_source Foreign Key (source_tab_id) References zltools.zlmip_table(ID) On Delete Cascade;
Alter Table zltools.zlmip_tab_extend Add Constraint zlmip_tab_extend_FK_target Foreign Key (target_tab_id) References zltools.zlmip_table(ID);
Alter Table zltools.zlmip_tab_extend Add Constraint zlmip_tab_extend_UQ_title Unique (source_tab_id,ext_title) Using Index;

--zlmip_tabext_conditon
Alter Table zltools.zlmip_tabext_condition Add Constraint zlmip_tabext_condition_PK Primary Key (ext_id,cond_order) Using Index;
Alter Table zltools.zlmip_tabext_condition Add Constraint zlmip_tabext_condition_FK_Exte Foreign Key (ext_id) References zltools.zlmip_tab_extend(ID) On Delete Cascade;
Alter Table zltools.zlmip_tabext_condition Add Constraint zlmip_tabext_condition_UQ_Targ Unique (ext_id,target_fld,source_fld) Using Index;

--zlmip_item
Alter Table zltools.zlmip_item Add Constraint zlmip_item_PK Primary Key (id) Using Index;
Alter Table zltools.zlmip_item Add Constraint zlmip_item_FK_tabid Foreign Key (tab_id) References zltools.zlmip_table(id);
Alter Table zltools.zlmip_item Add Constraint zlmip_item_UQ_Code Unique (item_code,item_type) Using Index;

--zlmip_item_parameter
Alter Table zltools.zlmip_item_parameter Add Constraint zlmip_item_para_FK_item_id Foreign Key (item_id) References zltools.zlmip_item(id);

--zlmip_item_field
Alter Table zltools.zlmip_item_field Add Constraint zlmip_item_field_FK_item_id Foreign Key (item_id) References zltools.zlmip_item(id);

--zlmip_item_frequency
Alter Table zltools.zlmip_item_frequency Add Constraint zlmip_item_freq_FK_item_id Foreign Key (item_id) References zltools.zlmip_item(id);

--zlmip_item_config
Alter Table zltools.zlmip_item_config Add Constraint zlmip_item_config_PK Primary Key (id) Using Index;
Alter Table zltools.zlmip_item_config Add Constraint zlmip_item_config_FK_parentid Foreign Key (parent_id) References zltools.zlmip_item_config(id);
Alter Table zltools.zlmip_item_config Add Constraint zlmip_item_config_FK_itemid Foreign Key (item_id) References zltools.zlmip_item(id) On Delete Cascade;
Alter Table zltools.zlmip_item_config Add Constraint zlmip_item_config_UQ_node Unique (item_id,node_order) Using Index;

--zlmip_item_deliver
Alter Table zltools.zlmip_item_deliver Add Constraint zlmip_item_deliver_PK Primary Key (item_id,deliver_order) Using Index;
Alter Table zltools.zlmip_item_deliver Add Constraint zlmip_item_deliver_FK_item_id Foreign Key (item_id) References zltools.zlmip_item(id);

--zlmip_send_log
Alter Table zltools.zlmip_send_log Add Constraint zlmip_send_log_PK Primary Key (id) Using Index;
Alter Table zltools.zlmip_send_log Add Constraint zlmip_send_log_UQ_msgid Unique (msg_id) Using Index;
Alter Table zltools.zlmip_send_log Add Constraint zlmip_send_log_FK_item_id Foreign Key (item_id) References zltools.zlmip_item(id);

--zlmip_sendlog_parameter
Alter Table zltools.zlmip_sendlog_parameter Add Constraint zlmip_sendlog_parameter_PK Primary Key (send_log_id,para_order) Using Index;
Alter Table zltools.zlmip_sendlog_parameter Add Constraint zlmip_sendlog_parameter_FK_rec Foreign Key (send_log_id) References zltools.zlmip_send_log(id) On Delete Cascade;

--zlmip_sendlog_again
Alter Table zltools.zlmip_sendlog_again Add Constraint zlmip_sendlog_again_PK Primary Key (send_log_id,send_order) Using Index;
Alter Table zltools.zlmip_sendlog_again Add Constraint zlmip_sendlog_again_FK_rec Foreign Key (send_log_id) References zltools.zlmip_send_log(id) On Delete Cascade;

--zlmip_receive_log
Alter Table zltools.zlmip_receive_log Add Constraint zlmip_receive_log_PK Primary Key (id) Using Index;

--zlmip_onlineusers
Alter Table zltools.zlmip_onlineusers Add Constraint zlmip_onlineusers_PK Primary Key (session_id,station,loginuser) Using Index;

--zlmip_onlinemodules
Alter Table zltools.zlmip_onlinemodules Add Constraint zlmip_onlinemodules_PK Primary Key (session_id,station,system,module,instance) Using Index;

--zlmip_clientuser
Alter Table zltools.zlmip_clientuser Add Constraint zlmip_clientuser_PK Primary Key (station) Using Index;
-----------------------------------------------------------------------------------------------------------------------
--数据表索引部分
-----------------------------------------------------------------------------------------------------------------------
--zlmip_tab_field
Create Index zltools.zlmip_tab_field_IX_tabid On zltools.zlmip_tab_field(tab_id);

--zlmip_tab_parameter
Create Index zltools.zlmip_tab_parameter_IX_tabid On zltools.zlmip_tab_parameter(tab_id);

--zlmip_item
Create Index zltools.zlmip_item_IX_tabid On zltools.zlmip_item(tab_id);

--zlmip_item_config
Create Index zltools.zlmip_item_config_IX_parentid On zltools.zlmip_item_config(parent_id);
-----------------------------------------------------------------------------------------------------------------------
--对象授权部分
-----------------------------------------------------------------------------------------------------------------------
Grant Select On zltools.zlmip_onlineusers To public;
Grant Select On zltools.zlmip_onlinemodules To public;
Grant Select On zltools.zlmip_receive_log To public;
Grant Select On zltools.zlmip_sendlog_again To public;
Grant Select On zltools.zlmip_item_deliver To public;
Grant Select On zltools.zlmip_item_config To public;
Grant Select On zltools.zlmip_item To public;
Grant Select On zltools.zlmip_tabext_condition To public;
Grant Select On zltools.zlmip_tab_extend To public;
Grant Select On zltools.zlmip_tab_parameter To public;
Grant Select On zltools.zlmip_tab_field To public;
Grant Select On zltools.zlmip_table To public;
Grant Select On zltools.zlmip_send_log To public;
Grant Select On zltools.zlmip_sendlog_parameter To public;
Grant Select On zltools.zlmip_item_field To public;
Grant Select On zltools.zlmip_item_frequency To public;
Grant Select On zltools.zlmip_item_parameter To public;
Grant Select On zltools.zlmip_data_setup To public;
Grant Select On zltools.zlmip_clientuser To public;
-----------------------------------------------------------------------------------------------------------------------
--对象同义词部分
-----------------------------------------------------------------------------------------------------------------------
Create Public Synonym zlmip_onlineusers for zltools.zlmip_onlineusers;
Create Public Synonym zlmip_onlinemodules for zltools.zlmip_onlinemodules;
Create Public Synonym zlmip_receive_log for zltools.zlmip_receive_log;
Create Public Synonym zlmip_sendlog_again for zltools.zlmip_sendlog_again;
Create Public Synonym zlmip_item_deliver for zltools.zlmip_item_deliver;
Create Public Synonym zlmip_item_config for zltools.zlmip_item_config;
Create Public Synonym zlmip_item for zltools.zlmip_item;
Create Public Synonym zlmip_tabext_condition for zltools.zlmip_tabext_condition;
Create Public Synonym zlmip_tab_extend for zltools.zlmip_tab_extend;
Create Public Synonym zlmip_tab_parameter for zltools.zlmip_tab_parameter;
Create Public Synonym zlmip_tab_field for zltools.zlmip_tab_field;
Create Public Synonym zlmip_table for zltools.zlmip_table;
Create Public Synonym zlmip_data_setup for zltools.zlmip_data_setup;
Create Public Synonym zlmip_send_log for zltools.zlmip_send_log;
Create Public Synonym zlmip_sendlog_parameter for zltools.zlmip_sendlog_parameter;
Create Public Synonym zlmip_item_field for zltools.zlmip_item_field;
Create Public Synonym zlmip_item_frequency for zltools.zlmip_item_frequency;
Create Public Synonym zlmip_item_parameter for zltools.zlmip_item_parameter;
Create Public Synonym zlmip_clientuser for zltools.zlmip_clientuser;