-----------------------------------------------------------------------------------------------------------------------
--过程
-----------------------------------------------------------------------------------------------------------------------
--ZLHIS客户端调用
CREATE OR REPLACE Procedure zltools.zlMip_OnLineUser_Insert
(
	station_In		In	zlmip_onlineusers.station%Type,
	loginuser_In	In	zlmip_onlineusers.loginuser%Type
)Is
	n_session		Number(18);
Begin
	Delete From zlmip_onlineusers A Where a.Session_Id Not In (Select Audsid From V$session Where Audsid > 0);
	n_session:=Sys_Context('USERENV', 'SESSIONID');
	Insert Into zlmip_onlineusers(session_id,station,loginuser)
	Select n_session,station_In,loginuser_In From Dual
	Where Not Exists (Select 1 From zlmip_onlineusers Where session_id=n_session And station=station_In And loginuser=loginuser_In);
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End zlMip_OnLineUser_Insert;
/

--ZLHIS客户端调用
CREATE OR REPLACE Procedure zltools.zlMip_OnLineUser_Delete
(
	station_In		In	zlmip_onlineusers.station%Type,
	loginuser_In	In	zlmip_onlineusers.loginuser%Type
)Is
	n_session		Number(18);
Begin
	n_session:=Sys_Context('USERENV', 'SESSIONID');
	Delete From zlmip_onlineusers Where session_id=n_session And station=station_In And loginuser=loginuser_In;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End zlMip_OnLineUser_Delete;
/

--ZLHIS客户端调用
CREATE OR REPLACE Procedure zltools.zlMip_OnLineModule_Insert
(	
	station_In		In	zlmip_onlinemodules.station%Type,
	system_In		In	zlmip_onlinemodules.system%Type,
	module_In		In	zlmip_onlinemodules.module%Type,
	instance_In		In	zlmip_onlinemodules.instance%Type:=Null
)Is
	n_session		Number(18);
Begin
	Delete From zlmip_onlinemodules A Where a.Session_Id Not In (Select Audsid From V$session Where Audsid > 0);
	n_session:=Sys_Context('USERENV', 'SESSIONID');
	Insert Into zlmip_onlinemodules(session_id,station,system,module,instance)
	Select n_session,station_In,system_In,module_In,instance_In From Dual
	Where Not Exists (Select 1 From zlmip_onlinemodules Where session_id=n_session And station=station_In And system=system_In And module=module_In And instance=instance_In);
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End zlMip_OnLineModule_Insert;
/
--ZLHIS客户端调用
CREATE OR REPLACE Procedure zltools.zlMip_OnLineModule_Delete
(
	station_In		In	zlmip_onlinemodules.station%Type,
	system_In		In	zlmip_onlinemodules.system%Type,
	module_In		In	zlmip_onlinemodules.module%Type,
	instance_In		In	zlmip_onlinemodules.instance%Type:=Null
)Is
	n_session		Number(18);
Begin
	n_session:=Sys_Context('USERENV', 'SESSIONID');
	Delete From zlmip_onlinemodules Where session_id=n_session And station=station_In And system=system_In And module=module_In And instance=instance_In;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End zlMip_OnLineModule_Delete;
/

--ZLHIS客户端调用
CREATE OR REPLACE Procedure zltools.zlMip_SendLog_Insert
(
	id_In			In	zlmip_send_log.id%Type,
	msg_id_In		In	zlmip_send_log.msg_id%Type,
	item_identity_In	In	zlmip_send_log.item_identity%Type,
	send_content_In	In	varchar2
)Is
Begin
	If send_content_In Is Null Then
		Insert Into zlmip_send_log(id,msg_id,item_identity,trigger_date,send_content,send_again)
		Select id_In,msg_id_In,item_identity_In,Sysdate,Null,0 From Dual;
	Else
		Insert Into zlmip_send_log(id,msg_id,item_identity,trigger_date,send_content,send_again)
		Select id_In,msg_id_In,item_identity_In,Sysdate,xmltype(send_content_In),0 From Dual;
	End If;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End zlMip_SendLog_Insert;
/

--ZLHIS客户端调用
CREATE OR REPLACE Procedure zltools.zlMip_SendLog_UpdateState
(
	id_In				In	zlmip_send_log.id%Type,
	send_state_In		In	zlmip_send_log.send_state%Type,
	send_mipuser_In	In	zlmip_send_log.send_mipuser%Type,
	send_user_In		In	zlmip_send_log.send_user%Type,
	send_username_In	In	zlmip_send_log.send_username%Type,
	send_station_In		In	zlmip_send_log.send_station%Type,
	send_module_In		In	zlmip_send_log.send_module%Type,
	send_content_In		In	varchar2,
	send_note_In		In	zlmip_send_log.send_note%Type:=Null
) Is
	n_again_policy		zlmip_item.again_policy%Type;
	v_again_para		zlmip_item.again_para%Type;
Begin
	n_again_policy:=0;
	v_again_para:='';
	If send_state_In=2 Then
		Select Nvl(Max(a.again_policy),0),Nvl(Max(a.again_para),'-') Into n_again_policy,v_again_para From zlmip_item a,zlmip_send_log b Where b.id=id_In And b.item_id=a.id;
		If v_again_para='-' Then
			v_again_para:='';
		End If;
	End If;
	
	Update zlmip_send_log Set send_state=send_state_In,
						send_mipuser=send_mipuser_In,
						send_user=send_user_In,
						send_username=send_username_In,
						send_station=send_station_In,
						send_module=send_module_In,
						send_again=Decode(send_state_In,1,0,n_again_policy),
						send_again_number=Decode(send_again_number,Null,0,send_again_number+1),
						send_again_date=Decode(send_date,Null,Null,Sysdate),
						send_date=Decode(send_date,Null,Sysdate,send_date),							
						send_again_para=Decode(send_state_In,1,Null,v_again_para),
						send_note=send_note_In
	Where id=id_In;	
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End zlMip_SendLog_UpdateState;
/

--ZLHIS客户端调用
Create Or Replace Procedure zltools.zlMip_Clob_Append
(
	Tab_In In Number,
	Key_In In Varchar2,
	Txt_In In Varchar2, 
	Cls_In In Number := 0, --是否清除原来的内容，第一片段传递时为1，以后为0
	Commit_In In Number := 0
	--参数说明：
	--Tab_In：包含CLOB的数据表
	--Key_In：数据记录的关键字
	--Txt_In：16进制的文件片段或文字片段
	--Cls_In：是否清除原来的内容，第一片段传递时为1，以后为0
) Is
	v_Clob Clob;
Begin	
	If Tab_In = 1 Then
		If Cls_In = 1 Then
			Update zlmip_send_log Set send_content_temp = Empty_Clob() Where id = Key_In;
		End If;
		Select send_content_temp Into v_Clob From zlmip_send_log Where id = Key_In For Update; 
		Dbms_Lob.Writeappend(v_Clob, Length(Txt_In), Txt_In);
		If Commit_In=1 Then
			Update zlmip_send_log Set send_content=xmltype(send_content_temp) Where  id = Key_In;
		End If;
	End If;

	If Tab_In = 2 Then
		If Cls_In = 1 Then
			Update zlmip_receive_log Set receive_content_temp = Empty_Clob() Where id = Key_In;
		End If;
		Select receive_content_temp Into v_Clob From zlmip_receive_log Where id = Key_In For Update; 
		Dbms_Lob.Writeappend(v_Clob, Length(Txt_In), Txt_In);

		If Commit_In=1 Then
			Update zlmip_receive_log Set receive_content=xmltype(receive_content_temp) Where  id = Key_In;
		End If;
	End If;
Exception
	When Others Then Zl_Errorcenter(Sqlcode, Sqlerrm);
End zlMip_Clob_Append;
/

--ZLHIS客户端调用
CREATE OR REPLACE Procedure zltools.zlMip_ReceiveLog_Insert
(
	id_In				In	zlmip_receive_log.id%Type,
	msg_id_In			In	zlmip_receive_log.msg_id%Type,
	item_identity_In		In	zlmip_receive_log.item_identity%Type,
	receive_mipuser_In	In	zlmip_receive_log.receive_mipuser%Type,
	receive_user_In		In	zlmip_receive_log.receive_user%Type,
	receive_username_In	In	zlmip_receive_log.receive_username%Type,
	receive_station_In	In	zlmip_receive_log.receive_station%Type,
	receive_module_In	In	zlmip_receive_log.receive_module%Type,
	receive_content_In	In	Varchar2,	
	receive_note_In		In	zlmip_receive_log.receive_note%Type:=Null
)Is
Begin
	Insert Into zlmip_receive_log(id,msg_id,item_identity,receive_date,receive_mipuser,receive_user,receive_username,receive_station,receive_module,receive_content,receive_note)
	Select id_In,msg_id_In,item_identity_In,Sysdate,receive_mipuser_In,receive_user_In,receive_username_In,receive_station_In,receive_module_In,Decode(receive_content_In,Null,Null,xmltype(receive_content_In)),receive_note_In From Dual ;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End zlMip_ReceiveLog_Insert;
/
-----------------------------------------------------------------------------------------------------------------------
--包
-----------------------------------------------------------------------------------------------------------------------
--zlMip_PkgManage 包头
Create Or Replace Package zltools.zlMip_PkgManage Is
	/*****************************************************************************************************************************************************/
	-- 功    能：增加、修改业务信息表信息
	/*****************************************************************************************************************************************************/
	Procedure Table_Update(
		id_In      In  zlmip_table.id%Type,
		data_code_In  In  zlmip_table.data_code%Type,
		tab_code_In  In  zlmip_table.tab_code%Type,
		tab_title_In  In  zlmip_table.tab_title%Type,
		tab_sqltext_In  In  zlmip_table.tab_sqltext%Type,
		tab_note_In  In  zlmip_table.tab_note%Type);

	/*****************************************************************************************************************************************************/
	-- 功    能：删除业务信息表的参数配置信息
	/*****************************************************************************************************************************************************/
	Procedure Tab_Parameter_Delete(tab_id_In In zlmip_tab_parameter.tab_id%Type);

	/*****************************************************************************************************************************************************/
	-- 功    能：增加业务信息表的参数配置信息
	/*****************************************************************************************************************************************************/
	Procedure Tab_Parameter_Insert(tab_id_In In zlmip_tab_parameter.tab_id%Type,para_In In varchar2);

	/*****************************************************************************************************************************************************/
	-- 功    能：删除业务信息表的字段信息
	/*****************************************************************************************************************************************************/
	Procedure Tab_Field_Delete(tab_id_In In zlmip_tab_field.tab_id%Type);

	/*****************************************************************************************************************************************************/
	-- 功    能：增加业务信息表的字段信息
	/*****************************************************************************************************************************************************/
	Procedure Tab_Field_Insert(tab_id_In In  zlmip_tab_field.tab_id%Type,field_In In  varchar2);

	/*****************************************************************************************************************************************************/
	-- 功    能：删除业务信息表信息
	/*****************************************************************************************************************************************************/
	Procedure Table_Delete(id_In In	zlmip_table.id%Type);

	/*****************************************************************************************************************************************************/
	-- 功    能：删除业务信息表的关系配置信息
	/*****************************************************************************************************************************************************/
	Procedure Tab_Extend_Delete(id_In In zlmip_tab_extend.id%Type);

	/*****************************************************************************************************************************************************/
	-- 功    能：增加、修改业务信息表的关系配置信息
	/*****************************************************************************************************************************************************/
	Procedure Tab_Extend_Update(
		id_In				In  zlmip_tab_extend.id%Type,
		source_tab_id_In	In  zlmip_tab_extend.source_tab_id%Type,
		ext_type_In		In  zlmip_tab_extend.ext_type%Type,
		ext_title_In		In  zlmip_tab_extend.ext_title%Type,
		target_tab_id_In	In  zlmip_tab_extend.target_tab_id%Type
	);

	/*****************************************************************************************************************************************************/
	-- 功    能：删除业务信息表的关系条件配置信息
	/*****************************************************************************************************************************************************/
	Procedure TabExt_Condition_Delete(ext_id_In In zlmip_tabext_condition.ext_id%Type);

	/*****************************************************************************************************************************************************/
	-- 功    能：增加业务信息表的关系条件配置信息
	/*****************************************************************************************************************************************************/
	Procedure TabExt_Condition_Insert(ext_id_In In zlmip_tabext_condition.ext_id%Type,Condition_In In varchar2);
	
	/*****************************************************************************************************************************************************/
	-- 功    能：增加、修改消息项目
	/*****************************************************************************************************************************************************/
	Procedure Item_Update(
		id_In					In	zlmip_item.id%Type,
		data_code_In			In	zlmip_item.data_code%Type,	
		item_code_In			In	zlmip_item.item_code%Type,
		item_title_In			In	zlmip_item.item_title%Type,
		item_flow_In			In	zlmip_item.item_flow%Type,
		check_frequency_In		In	zlmip_item.check_frequency%Type,
		check_freq_internal_In		In	zlmip_item.check_freq_internal%Type,
		again_policy_In			In	zlmip_item.again_policy%Type,
		again_para_In			In	zlmip_item.again_para%Type,
		start_date_In			In	zlmip_item.start_date%Type,
		stop_date_In			In	zlmip_item.stop_date%Type,
		freq_content_In			In	varchar2:=null,
		trigger_condition_In		In	zlmip_item.trigger_condition%Type:=null,
		trigger_frequency_In		In	zlmip_item.trigger_frequency%Type:=0,
		item_note_In			In	zlmip_item.item_note%Type:=null);

	/*****************************************************************************************************************************************************/
	-- 功    能：
	/*****************************************************************************************************************************************************/
	Procedure Item_Field_Delete(item_id_In In zlmip_item_field.item_id%Type);

	/*****************************************************************************************************************************************************/
	-- 功    能：
	/*****************************************************************************************************************************************************/
	Procedure Item_Field_Insert(item_id_In In  zlmip_item_field.item_id%Type,field_In In  varchar2);

	/*****************************************************************************************************************************************************/
	-- 功    能：
	/*****************************************************************************************************************************************************/
	Procedure Item_UpdateTab(id_In In zlmip_item.id%Type,tab_id_In In zlmip_item.tab_id%Type);
		
	/*****************************************************************************************************************************************************/
	-- 功    能：删除消息项目
	/*****************************************************************************************************************************************************/
	Procedure Item_Delete(id_In In	zlmip_item.id%Type);
	
	/*****************************************************************************************************************************************************/
	-- 功    能：删除业务事件关联的消息赋值信息
	/*****************************************************************************************************************************************************/
	Procedure Item_Config_Delete(item_id_In In	zlmip_item_config.item_id%Type);

	/*****************************************************************************************************************************************************/
	-- 功    能：增加业务事件关联的消息赋值信息
	/*****************************************************************************************************************************************************/
	Procedure Item_Config_Insert(item_id_In In zlmip_item_config.item_id%Type,Config_In In varchar2);

	/*****************************************************************************************************************************************************/
	-- 功    能：
	/*****************************************************************************************************************************************************/
	Procedure Item_Deliver_Update(
		id_In					In	zlmip_item_deliver.id%Type,
		item_id_In				In	zlmip_item_deliver.item_id%Type,
		deliver_title_In			In	zlmip_item_deliver.deliver_title%Type,
		deliver_code_In			In	zlmip_item_deliver.deliver_code%Type,
		deliver_object_In		In	varchar2);

	/*****************************************************************************************************************************************************/
	-- 功    能：
	/*****************************************************************************************************************************************************/
	Procedure Item_Deliver_Delete(id_In In zlmip_item_deliver.id%Type);

	/*****************************************************************************************************************************************************/
	-- 功    能：
	/*****************************************************************************************************************************************************/
	Procedure Options_Update(参数号_In In zlOptions.参数号%Type,参数名_In In zlOptions.参数名%Type,参数值_In In zlOptions.参数值%Type);

	/*****************************************************************************************************************************************************/
	-- 功    能：
	/*****************************************************************************************************************************************************/
	Procedure Clients_Update(工作站_In In zlmip_clientuser.station%Type,消息用户_In In zlmip_clientuser.mipuser%Type,消息密码_In In zlmip_clientuser.mipuserpassword%Type);

	/*****************************************************************************************************************************************************/
	-- 功    能：
	/*****************************************************************************************************************************************************/	
	Procedure Clients_Copy(来源工作站_In In zlmip_clientuser.station%Type,目标工作站_In In zlmip_clientuser.station%Type);

	/*****************************************************************************************************************************************************/
	-- 功    能：
	/*****************************************************************************************************************************************************/
	Procedure Send_Log_Delete(id_In In varchar2);	

	/*****************************************************************************************************************************************************/
	-- 功    能：
	/*****************************************************************************************************************************************************/
	Procedure Send_Log_Repeat(id_In In varchar2);	
	
	/*****************************************************************************************************************************************************/
	-- 功    能：
	/*****************************************************************************************************************************************************/
	Procedure Receive_Log_Delete(id_In In varchar2);	
	
End zlMip_PkgManage;
/

--zlMip_PkgManage 包体
Create Or Replace Package Body zltools.zlMip_PkgManage Is	
	/*****************************************************************************************************************************************************/
	-- 功    能：增加、修改业务信息表信息
	/*****************************************************************************************************************************************************/
	Procedure Table_Update(
		id_In      In  zlmip_table.id%Type,
		data_code_In  In  zlmip_table.data_code%Type,
		tab_code_In  In  zlmip_table.tab_code%Type,
		tab_title_In  In  zlmip_table.tab_title%Type,
		tab_sqltext_In  In  zlmip_table.tab_sqltext%Type,
		tab_note_In  In  zlmip_table.tab_note%Type) 
	Is
	Begin
		Update zlmip_table  Set  tab_code=tab_code_In,
						tab_title=tab_title_In,
						tab_sqltext=tab_sqltext_In,
						tab_note=tab_note_In
		Where id=id_In;

		If SQL%RowCount=0 Then
			Insert Into zlmip_table(id,data_code,tab_type,tab_code,tab_title,tab_sqltext,tab_note) 
			Values (id_In,data_code_In,2,tab_code_In,tab_title_In,tab_sqltext_In,tab_note_In);
		End If;

	Exception
		When Others Then Zl_ErrorCenter (SQLCODE, SQLERRM);
	End Table_Update;

	/*****************************************************************************************************************************************************/
	-- 功    能：删除业务信息表的参数配置信息
	/*****************************************************************************************************************************************************/
	Procedure Tab_Parameter_Delete(tab_id_In In zlmip_tab_parameter.tab_id%Type) Is
	Begin
		Delete From zlmip_tab_parameter Where tab_id=tab_id_In;
	Exception
		When Others Then Zl_ErrorCenter (SQLCODE, SQLERRM);
	End Tab_Parameter_Delete;
	
	/*****************************************************************************************************************************************************/
	-- 功    能：增加业务信息表的参数配置信息
	/*****************************************************************************************************************************************************/
	Procedure Tab_Parameter_Insert(
		tab_id_In		In  zlmip_tab_parameter.tab_id%Type,
		para_In		In  varchar2)
	Is
		n_para_order		zlmip_tab_parameter.para_order%Type;
		v_para_title		zlmip_tab_parameter.para_title%Type;
		n_para_type		zlmip_tab_parameter.para_type%Type;
		v_para_default		zlmip_tab_parameter.para_default%Type;
		v_para_note		zlmip_tab_parameter.para_note%Type;

		v_TmpRow			Varchar2(4000);
		v_Tmp			Varchar2(4000);
		n_Pos			Number(18);
	Begin
		v_Tmp := para_In||';';

		While v_Tmp Is Not Null Loop
			n_Pos := Instr(v_Tmp, ';');
			
			If n_Pos >0 Then
				
				v_TmpRow := Substr(v_Tmp, 1,n_Pos - 1);
				v_Tmp := Substr(v_Tmp, n_Pos + 1);
				
				n_Pos := Instr(v_TmpRow, ',');			
				n_para_order := To_Number(Substr(v_TmpRow, 1, n_Pos - 1));

				v_TmpRow := Substr(v_TmpRow, n_Pos + 1);
				n_Pos := Instr(v_TmpRow, ',');
				v_para_title := Substr(v_TmpRow, 1, n_Pos - 1);

				v_TmpRow := Substr(v_TmpRow, n_Pos + 1);
				n_Pos := Instr(v_TmpRow, ',');
				n_para_type := To_Number(Substr(v_TmpRow, 1, n_Pos - 1));

				v_TmpRow := Substr(v_TmpRow, n_Pos + 1);
				n_Pos := Instr(v_TmpRow, ',');
				v_para_default := Substr(v_TmpRow, 1, n_Pos - 1);

				v_para_note := Substr(v_TmpRow, n_Pos + 1);

				Insert Into zlmip_tab_parameter(tab_id,para_order,para_title,para_type,para_default,para_note) 
				Values (tab_id_In,n_para_order,v_para_title,n_para_type,v_para_default,v_para_note);

			End If;
		End Loop;
	Exception
	  When Others Then Zl_ErrorCenter (SQLCODE, SQLERRM);
	End Tab_Parameter_Insert;

	/*****************************************************************************************************************************************************/
	-- 功    能：删除业务信息表的字段信息
	/*****************************************************************************************************************************************************/
	Procedure Tab_Field_Delete(tab_id_In In zlmip_tab_field.tab_id%Type) Is
	Begin
		Delete From zlmip_tab_field Where tab_id=tab_id_In;
	Exception
		When Others Then Zl_ErrorCenter (SQLCODE, SQLERRM);
	End Tab_Field_Delete;

	/*****************************************************************************************************************************************************/
	-- 功    能：增加业务信息表的字段信息
	/*****************************************************************************************************************************************************/
	Procedure Tab_Field_Insert(
		tab_id_In		In  zlmip_tab_field.tab_id%Type,
		field_In		In  varchar2)
	Is
		n_fld_order		zlmip_tab_field.fld_order%Type;
		v_fld_title			zlmip_tab_field.fld_title%Type;
		n_fld_type			zlmip_tab_field.fld_type%Type;
		v_TmpRow			Varchar2(4000);
		v_Tmp			Varchar2(4000);
		n_Pos			Number(18);
	Begin
		v_Tmp := field_In||';';

		While v_Tmp Is Not Null Loop
			n_Pos := Instr(v_Tmp, ';');
			
			If n_Pos >0 Then
				
				v_TmpRow := Substr(v_Tmp, 1,n_Pos - 1);
				v_Tmp := Substr(v_Tmp, n_Pos + 1);
				
				n_Pos := Instr(v_TmpRow, ',');			
				n_fld_order := To_Number(Substr(v_TmpRow, 1, n_Pos - 1));

				v_TmpRow := Substr(v_TmpRow, n_Pos + 1);
				n_Pos := Instr(v_TmpRow, ',');
				v_fld_title := Substr(v_TmpRow, 1, n_Pos - 1);

				n_fld_type := To_Number(Substr(v_TmpRow, n_Pos + 1));
				Insert Into zlmip_tab_field(tab_id,fld_order,fld_title,fld_type) 
				Values (tab_id_In,n_fld_order,v_fld_title,n_fld_type);
			End If;
		End Loop;
	Exception
	  When Others Then Zl_ErrorCenter (SQLCODE, SQLERRM);
	End Tab_Field_Insert;

	/*****************************************************************************************************************************************************/
	-- 功    能：删除业务信息表信息
	/*****************************************************************************************************************************************************/
	Procedure Table_Delete(id_In In	zlmip_table.id%Type) Is
	Begin
	  Delete From zlmip_tab_parameter Where tab_id=id_In;
		Delete From zlmip_table  Where id=id_In;
	Exception
		When Others Then Zl_ErrorCenter (SQLCODE, SQLERRM);
	End Table_Delete;

	/*****************************************************************************************************************************************************/
	-- 功    能：删除业务信息表的关系配置信息
	/*****************************************************************************************************************************************************/
	Procedure Tab_Extend_Delete(id_In	In	zlmip_tab_extend.id%Type) Is
	Begin
		Delete From zlmip_tabext_condition Where ext_id=id_In;
		Delete From zlmip_tab_extend Where id=id_In;
	Exception
		When Others Then Zl_ErrorCenter (SQLCODE, SQLERRM);
	End Tab_Extend_Delete;

	/*****************************************************************************************************************************************************/
	-- 功    能：增加、修改业务信息表的关系配置信息
	/*****************************************************************************************************************************************************/
	Procedure Tab_Extend_Update(
		id_In				In  zlmip_tab_extend.id%Type,
		source_tab_id_In	In  zlmip_tab_extend.source_tab_id%Type,
		ext_type_In		In  zlmip_tab_extend.ext_type%Type,
		ext_title_In		In  zlmip_tab_extend.ext_title%Type,
		target_tab_id_In	In  zlmip_tab_extend.target_tab_id%Type)
	Is
		v_Error         Varchar2(255);
		Err_Custom    Exception;
	Begin
	  
	  If target_tab_id_In=source_tab_id_In Then
	     v_Error:='业务信息关系不能是自己关联自己，请重新指定关系！';
            Raise Err_Custom;
	  End If;

	  Update zlmip_tab_extend 
	  Set  ext_type=ext_type_In,
	       ext_title=ext_title_In,
	       target_tab_id=target_tab_id_In
	  Where id=id_In;

	  If SQL%RowCount=0 Then
	    Insert Into zlmip_tab_extend(id,source_tab_id,ext_order,ext_type,ext_title,target_tab_id) 
	    Select  id_In,source_tab_id_In,Nvl(Max(ext_order),0)+1,ext_type_In,ext_title_In,target_tab_id_In
	    From zlmip_tab_extend Where source_tab_id=source_tab_id_In;
	  End If;

	Exception
	  When Err_Custom Then Raise_application_error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
	  When Others Then Zl_ErrorCenter (SQLCODE, SQLERRM);
	End Tab_Extend_Update;

	/*****************************************************************************************************************************************************/
	-- 功    能：删除业务信息表的关系条件配置信息
	/*****************************************************************************************************************************************************/
	Procedure TabExt_Condition_Delete(ext_id_In	In zlmip_tabext_condition.ext_id%Type) Is
	Begin
		Delete From zlmip_tabext_condition Where ext_id=ext_id_In;
	Exception
		When Others Then Zl_ErrorCenter (SQLCODE, SQLERRM);
	End TabExt_Condition_Delete;

	/*****************************************************************************************************************************************************/
	-- 功    能：增加业务信息表的关系条件配置信息
	/*****************************************************************************************************************************************************/
	Procedure TabExt_Condition_Insert(
		ext_id_In		In	zlmip_tabext_condition.ext_id%Type,
		Condition_In		In	varchar2
	)
	Is
		n_cond_order			zlmip_tabext_condition.cond_order%Type;
		v_target_fld			zlmip_tabext_condition.target_fld%Type;
		v_source_fld			zlmip_tabext_condition.source_fld%Type;

		v_TmpRow			Varchar2(4000);
		v_Tmp			Varchar2(4000);
		n_Pos			Number(18);
	Begin
			
		v_Tmp := Condition_In||';';

		While v_Tmp Is Not Null Loop
			n_Pos := Instr(v_Tmp, ';');
			
			If n_Pos >0 Then
				
				v_TmpRow := Substr(v_Tmp, 1,n_Pos - 1);
				v_Tmp := Substr(v_Tmp, n_Pos + 1);
				
				n_Pos := Instr(v_TmpRow, ',');			
				n_cond_order := To_Number(Substr(v_TmpRow, 1, n_Pos - 1));

				v_TmpRow := Substr(v_TmpRow, n_Pos + 1);
				n_Pos := Instr(v_TmpRow, ',');
				v_target_fld := Substr(v_TmpRow, 1, n_Pos - 1);
				
				v_source_fld := Substr(v_TmpRow, n_Pos + 1);

				Insert Into zlmip_tabext_condition(ext_id,cond_order,target_fld,source_fld) 
				Values (ext_id_In,n_cond_order,v_target_fld,v_source_fld);	
			End If;
		End Loop;
	Exception
		When Others Then Zl_ErrorCenter (SQLCODE, SQLERRM);
	End TabExt_Condition_Insert;
	
	/*****************************************************************************************************************************************************/
	-- 功    能：增加、修改消息项目
	/*****************************************************************************************************************************************************/
	Procedure Item_Update(
		id_In					In	zlmip_item.id%Type,
		data_code_In			In	zlmip_item.data_code%Type,
		item_code_In			In	zlmip_item.item_code%Type,
		item_title_In			In	zlmip_item.item_title%Type,
		item_flow_In			In	zlmip_item.item_flow%Type,
		check_frequency_In		In	zlmip_item.check_frequency%Type,
		check_freq_internal_In		In	zlmip_item.check_freq_internal%Type,
		again_policy_In			In	zlmip_item.again_policy%Type,
		again_para_In			In	zlmip_item.again_para%Type,
		start_date_In			In	zlmip_item.start_date%Type,
		stop_date_In			In	zlmip_item.stop_date%Type,
		freq_content_In			In	varchar2:=null,
		trigger_condition_In		In	zlmip_item.trigger_condition%Type:=null,
		trigger_frequency_In		In	zlmip_item.trigger_frequency%Type:=0,
		item_note_In			In	zlmip_item.item_note%Type:=null)
	Is
		n_order				zlmip_item_frequency.freq_order%Type;
		n_freq_month			zlmip_item_frequency.freq_month%Type;
		n_freq_day				zlmip_item_frequency.freq_day%Type;
		v_freq_start			zlmip_item_frequency.freq_start%Type;
		v_freq_stop			zlmip_item_frequency.freq_stop%Type;

		v_TmpRow				Varchar2(4000);
		v_Tmp				Varchar2(4000);
		n_Pos				Number(18);
	Begin
		Update zlmip_item  Set	item_code=item_code_In,
							item_title=item_title_In,
							item_flow=item_flow_In,
							check_frequency=check_frequency_In,
							trigger_condition=trigger_condition_In,
							trigger_frequency=trigger_frequency_In,
							check_freq_internal=check_freq_internal_In,
							again_policy=again_policy_In,
							again_para=again_para_In,
							start_date=Decode(trunc(start_date_In),to_date('2000-01-01','yyyy-mm-dd'),sysdate,start_date_In),
							stop_date=stop_date_In,
							item_note=item_note_In
		Where id=id_In;

		If SQL%RowCount=0 Then
			Insert Into zlmip_item(id,data_code,item_code,item_title,item_type,item_flow,trigger_type,check_frequency,trigger_condition,trigger_frequency,check_freq_internal,again_policy,again_para,start_date,stop_date,item_note) 
			Values (id_In,data_code_In,item_code_In,item_title_In,2,item_flow_In,2,check_frequency_In,trigger_condition_In,trigger_frequency_In,check_freq_internal_In,again_policy_In,again_para_In,Decode(trunc(start_date_In),to_date('2000-01-01','yyyy-mm-dd'),sysdate,start_date_In),stop_date_In,item_note_In);
		End If;

		Delete From zlmip_item_frequency Where item_id=id_In;
		
		If freq_content_In Is Not Null Then
			v_Tmp := freq_content_In||';';
			n_order:=0;
			While v_Tmp Is Not Null Loop
				n_Pos := Instr(v_Tmp, ';');			
				If n_Pos >0 Then
					
					v_TmpRow := Substr(v_Tmp, 1,n_Pos - 1);
					v_Tmp := Substr(v_Tmp, n_Pos + 1);
					
					n_Pos := Instr(v_TmpRow, ',');			
					n_freq_month := To_Number(Substr(v_TmpRow, 1, n_Pos - 1));

					v_TmpRow := Substr(v_TmpRow, n_Pos + 1);
					n_Pos := Instr(v_TmpRow, ',');
					n_freq_day := To_Number(Substr(v_TmpRow, 1, n_Pos - 1));

					v_TmpRow := Substr(v_TmpRow, n_Pos + 1);
					n_Pos := Instr(v_TmpRow, ',');
					v_freq_start := Substr(v_TmpRow, 1, n_Pos - 1);
					
					v_freq_stop := Substr(v_TmpRow, n_Pos + 1);
					
					n_order:=n_order+1;
					Insert Into zlmip_item_frequency(item_id,freq_order,freq_month,freq_day,freq_start,freq_stop) 
					Values (id_In,n_order,Decode(n_freq_month,0,Null,n_freq_month),n_freq_day,v_freq_start,v_freq_stop);

				End If;
			End Loop;
		End If;
	Exception
		When Others Then Zl_ErrorCenter (SQLCODE, SQLERRM);
	End Item_Update;

	/*****************************************************************************************************************************************************/
	-- 功    能：
	/*****************************************************************************************************************************************************/
	Procedure Item_Field_Delete(item_id_In In zlmip_item_field.item_id%Type) Is
	Begin
		Delete From zlmip_item_field Where item_id=item_id_In;
	Exception
		When Others Then Zl_ErrorCenter (SQLCODE, SQLERRM);
	End Item_Field_Delete;

	/*****************************************************************************************************************************************************/
	-- 功    能：
	/*****************************************************************************************************************************************************/
	Procedure Item_Field_Insert(
		item_id_In		In  zlmip_item_field.item_id%Type,
		field_In		In  varchar2)
	Is
		n_fld_order		zlmip_item_field.fld_order%Type;
		v_fld_title			zlmip_item_field.fld_title%Type;
		n_fld_type			zlmip_item_field.fld_type%Type;
		v_TmpRow			Varchar2(4000);
		v_Tmp			Varchar2(4000);
		n_Pos			Number(18);
	Begin
		v_Tmp := field_In||';';

		While v_Tmp Is Not Null Loop
			n_Pos := Instr(v_Tmp, ';');
			
			If n_Pos >0 Then
				
				v_TmpRow := Substr(v_Tmp, 1,n_Pos - 1);
				v_Tmp := Substr(v_Tmp, n_Pos + 1);
				
				n_Pos := Instr(v_TmpRow, ',');			
				n_fld_order := To_Number(Substr(v_TmpRow, 1, n_Pos - 1));

				v_TmpRow := Substr(v_TmpRow, n_Pos + 1);
				n_Pos := Instr(v_TmpRow, ',');
				v_fld_title := Substr(v_TmpRow, 1, n_Pos - 1);

				n_fld_type := To_Number(Substr(v_TmpRow, n_Pos + 1));
				Insert Into zlmip_item_field(item_id,fld_order,fld_title,fld_type) 
				Values (item_id_In,n_fld_order,v_fld_title,n_fld_type);
			End If;
		End Loop;
	Exception
	  When Others Then Zl_ErrorCenter (SQLCODE, SQLERRM);
	End Item_Field_Insert;

	/*****************************************************************************************************************************************************/
	-- 功    能：
	/*****************************************************************************************************************************************************/
	Procedure Item_UpdateTab(id_In In zlmip_item.id%Type,tab_id_In In zlmip_item.tab_id%Type)
	Is
	Begin
		Update zlmip_item Set tab_id=tab_id_In Where id=id_In;
	Exception
		When Others Then Zl_ErrorCenter (SQLCODE, SQLERRM);
	End Item_UpdateTab;	

	/*****************************************************************************************************************************************************/
	-- 功    能：删除消息项目
	/*****************************************************************************************************************************************************/
	Procedure Item_Delete(id_In In	zlmip_item.id%Type) Is
	Begin
		Delete From zlmip_item_frequency Where item_id=id_In;
		Delete From zlmip_item Where id=id_In;
	Exception
		When Others Then Zl_ErrorCenter (SQLCODE, SQLERRM);
	End Item_Delete;
		
	/*****************************************************************************************************************************************************/
	-- 功    能：删除业务事件关联的消息赋值信息
	/*****************************************************************************************************************************************************/
	Procedure Item_Config_Delete(item_id_In In zlmip_item_config.item_id%Type) Is
	Begin
		Delete From zlmip_item_config Where item_id=item_id_In;
	Exception
		When Others Then Zl_ErrorCenter (SQLCODE, SQLERRM);
	End Item_Config_Delete;

	/*****************************************************************************************************************************************************/
	-- 功    能：增加业务事件关联的消息赋值信息
	/*****************************************************************************************************************************************************/
	Procedure Item_Config_Insert(item_id_In In	zlmip_item_config.item_id%Type,Config_In In varchar2)
	Is
		n_id					zlmip_item_config.id%Type;
		n_parent_id			zlmip_item_config.parent_id%Type;
		n_node_order			zlmip_item_config.node_order%Type;
		n_node_type			zlmip_item_config.node_type%Type;
		v_node_title			zlmip_item_config.node_title%Type;
		n_data_type			zlmip_item_config.data_type%Type;
		n_min_occurs			zlmip_item_config.min_occurs%Type;
		n_max_occurs			zlmip_item_config.max_occurs%Type;
		v_config_occurs			zlmip_item_config.config_occurs%Type;
		v_config_express		zlmip_item_config.config_express%Type;
		v_config_occurs_key		zlmip_item_config.config_occurs_key%Type;
		v_config_express_key		zlmip_item_config.config_express_key%Type;
		v_config_note			zlmip_item_config.config_note%Type;

		v_TmpRow				Varchar2(4000);
		v_Tmp				Varchar2(4000);
		n_Pos				Number(18);
	Begin
		--Item_Config_Insert('39643F2E01CB4249B282CDB1CB278AB7','AF259E08D5F840688F222D6994E5DE49,,1,2,text,,1,1,1,,,;9403BE2CF9D641479A5E9A4141B21549,,2,reports,,0,1,1,,,;78608BEB6D2F497D946315D269181E97,9403BE2CF9D641479A5E9A4141B21549,3,report,,0,u,1,,,;5DFFBD155D104A3C8390961C435C1623,78608BEB6D2F497D946315D269181E97,4,2,code,xs:string,1,1,,,,;EEDF2D6AEB7E4AAA90B7E7CA5D12F986,78608BEB6D2F497D946315D269181E97,5,2,title,xs:string,1,1,,,,;6745BAB0FC70483CBF713513212E33CF,,6,modules,,0,1,,,,;0E3A9E8B639F4F1D9847D26B8014019C,6745BAB0FC70483CBF713513212E33CF,7,module,,0,u,,,,;4BBC0A85D77F4134A9464DF28F754F35,0E3A9E8B639F4F1D9847D26B8014019C,8,2,sys,xs:string,1,1,,,,;F53AF488AF9745CA82573E7BEEEA4798,0E3A9E8B639F4F1D9847D26B8014019C,9,2,code,xs:string,1,1,,,,;36B564E5AD254F688542393C59C974E9,0E3A9E8B639F4F1D9847D26B8014019C,10,2,title,xs:string,1,1,,,,')

		v_Tmp := Config_In||';';

		While v_Tmp Is Not Null Loop
			n_Pos := Instr(v_Tmp, ';');
			
			If n_Pos >0 Then
				
				v_TmpRow := Substr(v_Tmp, 1,n_Pos - 1);
				v_Tmp := Substr(v_Tmp, n_Pos + 1);
				
				n_Pos := Instr(v_TmpRow, ',');			
				n_id := Substr(v_TmpRow, 1, n_Pos - 1);

				v_TmpRow := Substr(v_TmpRow, n_Pos + 1);
				n_Pos := Instr(v_TmpRow, ',');
				n_parent_id := Substr(v_TmpRow, 1, n_Pos - 1);

				v_TmpRow := Substr(v_TmpRow, n_Pos + 1);
				n_Pos := Instr(v_TmpRow, ',');
				n_node_order := To_Number(Substr(v_TmpRow, 1, n_Pos - 1));

				v_TmpRow := Substr(v_TmpRow, n_Pos + 1);
				n_Pos := Instr(v_TmpRow, ',');
				v_node_title := Substr(v_TmpRow, 1, n_Pos - 1);

				v_TmpRow := Substr(v_TmpRow, n_Pos + 1);
				n_Pos := Instr(v_TmpRow, ',');
				n_node_type := To_Number(Substr(v_TmpRow, 1, n_Pos - 1));
				
				v_TmpRow := Substr(v_TmpRow, n_Pos + 1);
				n_Pos := Instr(v_TmpRow, ',');
				n_data_type := Substr(v_TmpRow, 1, n_Pos - 1);

				v_TmpRow := Substr(v_TmpRow, n_Pos + 1);
				n_Pos := Instr(v_TmpRow, ',');
				n_min_occurs := Substr(v_TmpRow, 1, n_Pos - 1);

				v_TmpRow := Substr(v_TmpRow, n_Pos + 1);
				n_Pos := Instr(v_TmpRow, ',');
				n_max_occurs := Substr(v_TmpRow, 1, n_Pos - 1);

				v_TmpRow := Substr(v_TmpRow, n_Pos + 1);
				n_Pos := Instr(v_TmpRow, ',');
				v_config_occurs := Substr(v_TmpRow, 1, n_Pos - 1);

				v_TmpRow := Substr(v_TmpRow, n_Pos + 1);
				n_Pos := Instr(v_TmpRow, ',');
				v_config_express := Substr(v_TmpRow, 1, n_Pos - 1);

				v_TmpRow := Substr(v_TmpRow, n_Pos + 1);
				n_Pos := Instr(v_TmpRow, ',');
				v_config_occurs_key := Substr(v_TmpRow, 1, n_Pos - 1);
				
				v_config_express_key := Substr(v_TmpRow, n_Pos + 1);

				Insert Into zlmip_item_config(id,item_id,parent_id,node_order,node_type,node_title,data_type,min_occurs,max_occurs,config_occurs,config_occurs_key,config_express,config_express_key) 
				Values (n_id,item_id_In,n_parent_id,n_node_order,n_node_type,v_node_title,n_data_type,n_min_occurs,n_max_occurs,v_config_occurs,v_config_occurs_key,v_config_express,v_config_express_key);	
			End If;
		End Loop;
	Exception
		When Others Then Zl_ErrorCenter (SQLCODE, SQLERRM);
	End Item_Config_Insert;	

	/*****************************************************************************************************************************************************/
	-- 功    能：
	/*****************************************************************************************************************************************************/
	Procedure Item_Deliver_Update(
		id_In					In	zlmip_item_deliver.id%Type,
		item_id_In				In	zlmip_item_deliver.item_id%Type,
		deliver_title_In			In	zlmip_item_deliver.deliver_title%Type,
		deliver_code_In			In	zlmip_item_deliver.deliver_code%Type,
		deliver_object_In		In	varchar2)
	Is
	Begin
		If deliver_object_In Is Null Then
			Update zlmip_item_deliver  Set	deliver_title=deliver_title_In,	deliver_code=deliver_code_In,deliver_object=Null
			Where id=id_In;
			If SQL%RowCount=0 Then
				Insert Into zlmip_item_deliver(id,item_id,deliver_order,deliver_title,deliver_code,deliver_object)
				Select id_In,item_id_In,Nvl(Max(deliver_order),0)+1,deliver_title_In,deliver_code_In,Null
				From zlmip_item_deliver
				Where item_id=item_id_In;
			End If;
		Else
			Update zlmip_item_deliver  Set	deliver_title=deliver_title_In,	deliver_code=deliver_code_In,deliver_object=xmltype(deliver_object_In)
			Where id=id_In;

			If SQL%RowCount=0 Then
				Insert Into zlmip_item_deliver(id,item_id,deliver_order,deliver_title,deliver_code,deliver_object)
				Select id_In,item_id_In,Nvl(Max(deliver_order),0)+1,deliver_title_In,deliver_code_In,xmltype(deliver_object_In)
				From zlmip_item_deliver
				Where item_id=item_id_In;
			End If;
		End If;
	Exception
		When Others Then Zl_ErrorCenter (SQLCODE, SQLERRM);
	End Item_Deliver_Update;
	/*****************************************************************************************************************************************************/
	-- 功    能：
	/*****************************************************************************************************************************************************/
	Procedure Item_Deliver_Delete(id_In In zlmip_item_deliver.id%Type) Is
	Begin
		Delete From zlmip_item_deliver Where id=id_In;
	Exception
		When Others Then Zl_ErrorCenter (SQLCODE, SQLERRM);
	End Item_Deliver_Delete;

	/*****************************************************************************************************************************************************/
	-- 功    能：
	/*****************************************************************************************************************************************************/
	Procedure Options_Update(参数号_In In zlOptions.参数号%Type,参数名_In In zlOptions.参数名%Type,参数值_In In zlOptions.参数值%Type)
	Is
	Begin
		Update zlOptions  Set 参数值=参数值_In Where 参数号=参数号_In;

		If SQL%RowCount=0 Then
			Insert Into zlOptions(参数号,参数名,参数值) Values(参数号_In,参数名_In,参数值_In);
		End If;		
	Exception
		When Others Then Zl_ErrorCenter (SQLCODE, SQLERRM);
	End Options_Update;

	/*****************************************************************************************************************************************************/
	-- 功    能：
	/*****************************************************************************************************************************************************/
	Procedure Clients_Update(工作站_In In zlmip_clientuser.station%Type,消息用户_In In zlmip_clientuser.mipuser%Type,消息密码_In In zlmip_clientuser.mipuserpassword%Type)
	Is
	Begin
		If 消息用户_In Is Null Then
			Delete From zlmip_clientuser Where station=工作站_In;	
		Else
			Update zlmip_clientuser  Set mipuser=消息用户_In,mipuserpassword=消息密码_In Where station=工作站_In;	
			If SQL%RowCount=0 Then
				Insert Into zlmip_clientuser(station,mipuser,mipuserpassword)
				Values (工作站_In,消息用户_In,消息密码_In);
			End If;
		End If;
	Exception
		When Others Then Zl_ErrorCenter (SQLCODE, SQLERRM);
	End Clients_Update;

	/*****************************************************************************************************************************************************/
	-- 功    能：
	/*****************************************************************************************************************************************************/
	Procedure Clients_Copy(来源工作站_In In zlmip_clientuser.station%Type,目标工作站_In In zlmip_clientuser.station%Type)
	Is
	Begin
		For r_List In (Select mipuser,mipuserpassword From zlmip_clientuser Where station=来源工作站_In)
		Loop
			Update zlmip_clientuser  Set mipuser=r_List.mipuser,mipuserpassword=r_List.mipuserpassword Where station=目标工作站_In;	
			If SQL%RowCount=0 Then
				Insert Into zlmip_clientuser(station,mipuser,mipuserpassword)
				Values (目标工作站_In,r_List.mipuser,r_List.mipuserpassword);
			End If;
		End Loop;
	Exception
		When Others Then Zl_ErrorCenter (SQLCODE, SQLERRM);
	End Clients_Copy;

	/*****************************************************************************************************************************************************/
	-- 功    能：
	/*****************************************************************************************************************************************************/
	Procedure Send_Log_Delete(id_In In varchar2) 
	Is
		v_Tmp	varchar2(4000);
		n_Pos	number(18);
		n_id		varchar2(32);
	Begin
		v_Tmp := id_In||';';

		While v_Tmp Is Not Null Loop
			n_Pos := Instr(v_Tmp, ';');
			
			If n_Pos >0 Then
				
				n_id := Substr(v_Tmp, 1,n_Pos - 1);
				v_Tmp := Substr(v_Tmp, n_Pos + 1);
				
				If n_id Is Not Null Then
					Delete From zlmip_send_log Where id = n_id;
				End If;	
			End If;
		End Loop;
	Exception
		When Others Then Zl_ErrorCenter (SQLCODE, SQLERRM);
	End Send_Log_Delete;

	/*****************************************************************************************************************************************************/
	-- 功    能：
	/*****************************************************************************************************************************************************/
	Procedure Send_Log_Repeat(id_In In varchar2) 
	Is
		v_Tmp	varchar2(4000);
		n_Pos	number(18);
		n_id		varchar2(32);
	Begin
		v_Tmp := id_In||';';

		While v_Tmp Is Not Null Loop
			n_Pos := Instr(v_Tmp, ';');
			
			If n_Pos >0 Then
				
				n_id := Substr(v_Tmp, 1,n_Pos - 1);
				v_Tmp := Substr(v_Tmp, n_Pos + 1);
				
				If n_id Is Not Null Then
					Update zlmip_send_log Set send_again=1 Where id = n_id And Nvl(send_again,0)=0;
				End If;	
			End If;
		End Loop;
	Exception
		When Others Then Zl_ErrorCenter (SQLCODE, SQLERRM);
	End Send_Log_Repeat;
	
	/*****************************************************************************************************************************************************/
	-- 功    能：
	/*****************************************************************************************************************************************************/
	Procedure Receive_Log_Delete(id_In In varchar2) 
	Is
		v_Tmp	varchar2(4000);
		n_Pos	number(18);
		n_id		varchar2(32);
	Begin
		v_Tmp := id_In||';';

		While v_Tmp Is Not Null Loop
			n_Pos := Instr(v_Tmp, ';');
			
			If n_Pos >0 Then
				
				n_id := Substr(v_Tmp, 1,n_Pos - 1);
				v_Tmp := Substr(v_Tmp, n_Pos + 1);
				
				If n_id Is Not Null Then
					Delete From zlmip_Receive_log Where id = n_id;
				End If;	
			End If;
		End Loop;
	Exception
		When Others Then Zl_ErrorCenter (SQLCODE, SQLERRM);
	End Receive_Log_Delete;

End zlMip_PkgManage;
/
/*################################################################################################*/
--zlMip_PkgRollService 包头
Create Or Replace Package zltools.zlMip_PkgRollService Is
	
	Type t_StrArray Is Table of Varchar2(4000);
	Type t_Refcur Is Ref Cursor;
	
	/*****************************************************************************************************************************************************/
	-- 功    能：
	/*****************************************************************************************************************************************************/
	Procedure InsertMessage(	id_In		In	zlmip_send_log.id%Type,
						item_id_In	In	zlmip_send_log.item_id%Type);

	/*****************************************************************************************************************************************************/
	-- 功    能：
	/*****************************************************************************************************************************************************/
	Procedure InsertMessageParameter(  send_log_id_In	In	zlmip_sendlog_parameter.send_log_id%Type,
								para_order_In	In	zlmip_sendlog_parameter.para_order%Type,
								para_title_In	In	zlmip_sendlog_parameter.para_title%Type,
								para_value_In	In	zlmip_sendlog_parameter.para_value%Type);

	/*****************************************************************************************************************************************************/
	-- 功    能：
	/*****************************************************************************************************************************************************/
	Procedure BuildMessageContent(id_In In zlmip_send_log.id%Type,SendPara_In In Varchar2);	

	/*****************************************************************************************************************************************************/
	-- 功    能：
	/*****************************************************************************************************************************************************/
	Procedure RepeatMessageContent(id_In In zlmip_send_log.id%Type,SendPara_In In Varchar2,RepeatOrder_In In Number:=0);

	/*****************************************************************************************************************************************************/
	-- 功    能：
	/*****************************************************************************************************************************************************/
	Procedure UpdateSendLogState
	(
		id_In				In	zlmip_send_log.id%Type,
		send_state_In		In	zlmip_send_log.send_state%Type,
		send_mipuser_In	In	zlmip_send_log.send_mipuser%Type,
		send_user_In		In	zlmip_send_log.send_user%Type,
		send_username_In	In	zlmip_send_log.send_username%Type,
		send_station_In		In	zlmip_send_log.send_station%Type,
		send_note_In		In	zlmip_send_log.send_note%Type:=Null
	);

End zlMip_PkgRollService;
/
--zlMip_PkgRollService 包头
Create Or Replace Package Body zltools.zlMip_PkgRollService Is
	/*****************************************************************************************************************************************************/
	-- 功    能：
	/*****************************************************************************************************************************************************/
	Procedure InsertMessage(	id_In		In	zlmip_send_log.id%Type,
						item_id_In	In	zlmip_send_log.item_id%Type)
	Is
	Begin
		Insert Into zlmip_send_log(id,msg_id,item_id,item_identity,trigger_date,send_content,send_again)
		Select id_In,SYS_GUID(),id,item_code,Sysdate,Null,0 From zlmip_item Where id=item_id_In;
	Exception
		When Others Then Zl_ErrorCenter (SQLCODE, SQLERRM);
	End InsertMessage;

	/*****************************************************************************************************************************************************/
	-- 功    能：
	/*****************************************************************************************************************************************************/
	Procedure InsertMessageParameter(  send_log_id_In	In	zlmip_sendlog_parameter.send_log_id%Type,
								para_order_In	In	zlmip_sendlog_parameter.para_order%Type,
								para_title_In	In	zlmip_sendlog_parameter.para_title%Type,
								para_value_In	In	zlmip_sendlog_parameter.para_value%Type)
	Is
	Begin		
		Insert Into zlmip_sendlog_parameter(send_log_id,para_order,para_title,para_value)
		Values (send_log_id_In,para_order_In,para_title_In,para_value_In);
	Exception
		When Others Then Zl_ErrorCenter (SQLCODE, SQLERRM);
	End InsertMessageParameter;

	/*****************************************************************************************************************************************************/
	-- 功    能：
	/*****************************************************************************************************************************************************/
	Procedure UpdateSendLogState
	(
		id_In				In	zlmip_send_log.id%Type,
		send_state_In		In	zlmip_send_log.send_state%Type,
		send_mipuser_In	In	zlmip_send_log.send_mipuser%Type,
		send_user_In		In	zlmip_send_log.send_user%Type,
		send_username_In	In	zlmip_send_log.send_username%Type,
		send_station_In		In	zlmip_send_log.send_station%Type,
		send_note_In		In	zlmip_send_log.send_note%Type:=Null
	) Is
		n_again_policy		zlmip_item.again_policy%Type;
		v_again_para		zlmip_item.again_para%Type;
	Begin
		n_again_policy:=0;
		v_again_para:='';
		If send_state_In=2 Then
			Select Nvl(Max(a.again_policy),0),Nvl(Max(a.again_para),'-') Into n_again_policy,v_again_para From zlmip_item a,zlmip_send_log b Where b.id=id_In And b.item_id=a.id;
			If v_again_para='-' Then
				v_again_para:='';
			End If;
		End If;
		Update zlmip_send_log Set send_state=send_state_In,
							send_mipuser=send_mipuser_In,
							send_user=send_user_In,
							send_username=send_username_In,
							send_station=send_station_In,
							send_module='轮询服务程序',
							send_again=Decode(send_state_In,1,0,n_again_policy),
							send_again_number=Decode(send_again_number,Null,0,send_again_number+1),
							send_again_date=Decode(send_date,Null,Null,Sysdate),
							send_date=Decode(send_date,Null,Sysdate,send_date),							
							send_again_para=Decode(send_state_In,1,Null,v_again_para),
							send_note=send_note_In
		Where id=id_In;	
	Exception
	  When Others Then
	    Zl_Errorcenter(Sqlcode, Sqlerrm);
	End UpdateSendLogState;

	--****************************************************************************************************************************************************
	-- 功    能：
	--****************************************************************************************************************************************************
	Function GetSplitString(str_In In Varchar2) Return t_Strlist 
	As
		v_Str   Long Default str_In || ''''; 
		v_Index Number; 
		v_List  t_Strlist := t_Strlist(); 
	Begin 
		Loop 
			v_Index := Instr(v_Str, ''''); 
			Exit When(Nvl(v_Index, 0) = 0); 
			v_List.Extend; 
			v_List(v_List.Count) := Trim(Substr(v_Str, 1, v_Index - 1)); 
			v_Str := Substr(v_Str, v_Index + 1); 
		End Loop; 
		Return(v_List);
	End;
	--****************************************************************************************************************************************************
	-- 功    能：
	--****************************************************************************************************************************************************
	Function GetElement(strExpress_In In Varchar2) Return t_Strlist 
	As
		n_Count			Number(18);
		n_Loop			Number(18);
		v_Char			varchar2(2);
		v_Var				varchar2(4000);
		n_BeginVar		number(18):=0;
		n_EndVar			Number(18):=0;
		a_Return			t_Strlist := t_Strlist(); 
	Begin		
		n_Count:=length(strExpress_In);
		For n_Loop In 1..n_Count Loop
			v_Char:=substr(strExpress_In,n_Loop,1);
			If v_Char='[' Then
				n_BeginVar:=n_Loop;
			Elsif v_Char=']' Then
				IF n_BeginVar>0Then					
					n_EndVar:=n_Loop;
					v_Var:=v_Var||''''||substr(strExpress_In,n_BeginVar+1,n_EndVar-n_BeginVar-1);										
					n_BeginVar:=0;
					n_EndVar:=0;
				End If;
			End If;
		End Loop;			
		If v_Var Is Not Null Then
			v_Var:=rtrim(substr(v_Var,2));
			a_Return:=GetSplitString(v_Var);	
		End If;
		Return(a_Return);
	End;
	--****************************************************************************************************************************************************
	-- 功    能：
	--****************************************************************************************************************************************************
	Procedure CursorTableCreate
	Is
		v_SQL			Varchar2(4000);
	Begin		
		Begin
			Execute Immediate 'Drop Table m_Cursor';
		Exception
			When Others Then null;
		End;
		v_SQL:='Create Global Temporary Table  m_Cursor(Cursor_Key varchar2(32),Parent_Cursor_Hand Number(18),Cursor_Hand Number(18))  on commit delete rows';
		Execute Immediate v_SQL;
	Exception
		When Others Then Zl_ErrorCenter (SQLCODE, SQLERRM);
	End CursorTableCreate;
	--****************************************************************************************************************************************************
	-- 功    能：
	--****************************************************************************************************************************************************
	Function CursorTableSelect(Cursor_Key_In In Varchar2) Return Number 
	Is
		v_SQL			Varchar2(4000);
		n_Cursor			Number;
	Begin
		n_Cursor:=0;
		v_SQL:='Select Nvl(Max(Cursor_Hand),0) From m_Cursor Where Cursor_Key=:1';
		Execute Immediate v_SQL Into n_Cursor Using Cursor_Key_In;			
		Return(n_Cursor);
	Exception
		When Others Then Zl_ErrorCenter (SQLCODE, SQLERRM);
	End;
	--****************************************************************************************************************************************************
	-- 功    能：
	--****************************************************************************************************************************************************
	Procedure CursorTableAppend
	(
		Cursor_Key_In			In	Varchar2,
		Cursor_Hand_In			In	Number,
		Parent_Cursor_Hand_In	In	Number:=Null
	) Is
		v_SQL					Varchar2(4000);
	Begin
		v_SQL:='Insert Into m_Cursor(Cursor_Key,Cursor_Hand,Parent_Cursor_Hand) Values (:1,:2,:3)';
		Execute Immediate v_SQL Using Cursor_Key_In,Cursor_Hand_In,Parent_Cursor_Hand_In;
	Exception
		When Others Then Zl_ErrorCenter (SQLCODE, SQLERRM);
	End;
	--****************************************************************************************************************************************************
	-- 功    能：
	--****************************************************************************************************************************************************
	Procedure CursorTableDelete(Cursor_Key_In In Varchar2:=Null)
	Is
		v_SQL				Varchar2(4000);
		n_Cursor				number;
		TYPE cursor_type IS REF CURSOR; 
		Cur cursor_type;
	Begin
		If Cursor_Key_In Is Null Then
			n_Cursor:=0;
			v_SQL:='Select Cursor_Hand From m_Cursor ';
			OPEN Cur FOR v_SQL; 
			Loop				
				Fetch Cur Into n_Cursor;
				Exit When Cur%NOTFOUND;
				If n_Cursor>0 Then
					DBMS_SQL.close_cursor(n_Cursor);
					n_Cursor:=0;
				End If;
			End Loop;
			Execute Immediate 'Drop Table m_Cursor';
		Else
			n_Cursor:=CursorTableSelect(Cursor_Key_In);
			If n_Cursor>0 Then
				DBMS_SQL.close_cursor(n_Cursor);
				n_Cursor:=0;
				v_SQL:='Delete From m_Cursor Where Cursor_Key In (Select Cursor_Key From m_Cursor Start With Cursor_Key=:1 Connect By Prior Cursor_Hand=Parent_Cursor_Hand)';
				Execute Immediate v_SQL Using Cursor_Key_In;
			End If;	
		End If;
	Exception
		When Others Then Zl_ErrorCenter (SQLCODE, SQLERRM);
	End CursorTableDelete;
	--****************************************************************************************************************************************************
	-- 功    能：打开光标（SQL）
	--****************************************************************************************************************************************************
	Function OpenCursor(	Table_id_In In varchar2,u_Paravalue In t_Strlist) Return Number 
	Is
		n_Cursor			Number;
		n_State			number;
		v_SQL			Varchar2(4000);
		n_number			Number;
		v_varchar2			Varchar2(100);
		d_date			Date;
	Begin		
		---------------------------------------------------------------------------------------------------------------------
		n_Cursor := DBMS_SQL.open_cursor;					--打开游标
		v_SQL:='';
		For r_SQL In (Select a.tab_sqltext From zlmip_table a Where a.id=Table_id_In) Loop
			v_SQL := r_SQL.tab_sqltext;
		End Loop;
		--替换“[0]”形式为“参数名”形式
		---------------------------------------------------------------------------------------------------------------------
		For r_Para In (Select a.para_order,a.para_title,a.para_type From zlmip_tab_parameter a Where a.tab_id=Table_id_In Order By a.para_order) Loop	
			v_SQL:=Replace(v_SQL,'['||To_Char(r_Para.para_order-1)||']',':'||r_Para.para_title);		
		End Loop;

		--解析游标---------------------------------------------------------------------------------------------------------------------
		DBMS_SQL.parse(n_Cursor, v_SQL, DBMS_SQL.native);	
		
		--绑定参数---------------------------------------------------------------------------------------------------------------------
		For r_Para In (Select a.para_order,a.para_title,a.para_type From zlmip_tab_parameter a Where a.tab_id=Table_id_In Order By a.para_order) Loop	
			If r_Para.para_type=1 Then			--数值型
				DBMS_SQL.bind_variable(n_Cursor, ':'||r_Para.para_title,to_number(u_Paravalue(r_Para.para_order)));
			Elsif r_Para.para_type=2 Then		--字符型
				DBMS_SQL.bind_variable(n_Cursor, ':'||r_Para.para_title,u_Paravalue(r_Para.para_order));
			Elsif r_Para.para_type=3 Then		--日期型
				DBMS_SQL.bind_variable(n_Cursor, ':'||r_Para.para_title,to_date(u_Paravalue(r_Para.para_order),'yyyy-mm-dd hh24:mi:ss'));
			End If;
		End Loop;
		
		--定义列---------------------------------------------------------------------------------------------------------------------
		For r_Field In (Select a.fld_order,a.fld_type From zlmip_tab_field a Where a.tab_id=Table_id_In Order By a.fld_order) Loop	
			If r_Field.fld_type=1 Then			--数值型
				DBMS_SQL.define_column(n_Cursor, r_Field.fld_order, n_number);	
			Elsif r_Field.fld_type=2 Then		--字符型
				DBMS_SQL.define_column(n_Cursor, r_Field.fld_order, v_varchar2,4000);	
			Elsif r_Field.fld_type=3 Then		--日期型
				DBMS_SQL.define_column(n_Cursor, r_Field.fld_order, d_date);	
			End If;
		End Loop;	

		--执行SQL---------------------------------------------------------------------------------------------------------------------
		n_State := DBMS_SQL.Execute(n_Cursor);			
		Return(n_Cursor);
	End;	

	--****************************************************************************************************************************************************
	-- 功    能：
	--****************************************************************************************************************************************************
	Function GetTablePara
	(
		id_In				In	zlmip_send_log.id%Type,
		Table_id_In		In	varchar2,
		TabExt_id_In		In	varchar2:=null,
		Cursor_In			In	Number:=0
	) Return t_Strlist Is
		n_number				Number;
		v_varchar				Varchar2(100);
		d_date				Date;
		u_Para				t_Strlist:=t_Strlist();
		n_Order				Number;
	Begin		
		---------------------------------------------------------------------------------------------------------------------
		If cursor_In=0 Then
			--为入口信息表，参数从事件源处提供的中提取
			For r_Para In (Select a.para_order,a.para_title,a.para_type 
						From zlmip_tab_parameter a,zlmip_Table b 
						Where a.tab_id=b.id And b.id=Table_id_In Order By a.para_order) Loop	
			
				For r_Value In (Select b.para_value 
							From zlmip_send_log a,zlmip_sendlog_Parameter b  
							Where a.id=id_In And a.id=b.send_log_id And Upper(b.para_title)=Upper(r_Para.para_title)) Loop
					u_Para.Extend; 
					u_Para(u_Para.Count) := r_Value.para_value; 
				End Loop;
			End Loop;
		Else
			--非入口信息表，参数从上级信息表的对应字段值中提取
			For r_Para In (Select a.para_order,a.para_title,a.para_type 
						From zlmip_tab_parameter a,zlmip_Table b 
						Where a.tab_id=b.id And b.id=Table_id_In Order By a.para_order) Loop	
						
				--找出参数字段对应的原表字段的顺序
				n_Order:=0;
				Select Nvl(Max(c.fld_order),0) Into n_Order
				From zlmip_Tab_Extend a,zlmip_tabext_condition b,zlmip_tab_field c
				Where a.id=TabExt_id_In And b.ext_id=a.id And Upper(b.target_fld)=Upper(r_Para.para_title)
					And c.tab_id=a.source_tab_id And c.fld_title=b.source_fld;
				
				If n_Order>0 Then
					If r_Para.para_type=1 Then				--数值型
						DBMS_SQL.column_value(Cursor_In, n_Order, n_number);
						u_Para.Extend; 
						u_Para(u_Para.Count) := To_Char(n_number); 
					Elsif r_Para.para_type=2 Then			--字符型
						DBMS_SQL.column_value(Cursor_In, n_Order, v_varchar);
						u_Para.Extend; 
						u_Para(u_Para.Count) := v_varchar; 
					Elsif r_Para.para_type=3 Then			--日期型
						DBMS_SQL.column_value(Cursor_In, n_Order, d_date);
						u_Para.Extend; 
						u_Para(u_Para.Count) := To_Char(d_date,'yyyy-mm-dd hh24:mi:ss'); 
					End If;
				Else
					u_Para.Extend; 
					u_Para(u_Para.Count) := 'Null'; 
				End If;
			End Loop;
		End If;
		Return(u_Para);
	End;
	--****************************************************************************************************************************************************
	-- 功    能：
	--****************************************************************************************************************************************************
	Function OpenTable(id_In In zlmip_send_log.id%Type,Target_id_In In VARCHAR2) Return Number 
	Is
		v_SQL			Varchar2(2000);
		v_EnterTable_id		Varchar2(32);
		n_Cursor			Number;
		n_ParentCursor		Number;
		n_Ret				Number;
		u_Para			t_Strlist;
		n_Have			Number;
	Begin		
		---------------------------------------------------------------------------------------------------------------------
		--1.获取当前消息的入口表信息
		v_EnterTable_id:='';
		Begin
			Select c.tab_id Into v_EnterTable_id From zlmip_send_log a,zlmip_item c Where a.id=id_In And a.item_id=c.id;
		Exception
			When Others Then 
				v_EnterTable_id:='';
		End;
		---------------------------------------------------------------------------------------------------------------------	
		--2.判断当前消息的入口表游标是否打开，如果没有打开则立即打开并记录
		n_Cursor:=CursorTableSelect(v_EnterTable_id);					
		If n_Cursor=0 Then
			u_Para:=GetTablePara(id_In,v_EnterTable_id,Null,0);
			n_Cursor:=OpenCursor(v_EnterTable_id,u_Para);			
			If n_Cursor>0 Then
				CursorTableAppend(v_EnterTable_id,n_Cursor,Null);
				n_Ret:=DBMS_SQL.fetch_rows(n_Cursor);
			End If;		
		End If;
		---------------------------------------------------------------------------------------------------------------------
		--3.判断当前信息表是否就是入口表，如果不是，则通过信息表关系继续
		If Target_id_In<>v_EnterTable_id Then						
			n_ParentCursor:=n_Cursor;
			--循环信息表关系，找出上级业务信息表，并检查是否打开（感觉下面的SQL有问题，但验证好像又没有问题）
			For r_TabExt In (select a.id,a. source_tab_id,a.target_tab_id
						  from (select x.id,x.source_tab_id,x.target_tab_id
							  from zlmip_tab_extend x
							 start with x.id = Target_id_In
							Connect By prior x.source_tab_id = x.target_tab_id) a
						 start with a.source_tab_id = v_EnterTable_id
						Connect By a.source_tab_id = prior a.target_tab_id) Loop	
						
				n_Cursor:=CursorTableSelect(r_TabExt.id);
				If n_Cursor=0 Then					
					u_Para:=GetTablePara(id_In,r_TabExt.target_tab_id,r_TabExt.id,n_ParentCursor);
					n_Cursor:=OpenCursor(r_TabExt.target_tab_id,u_Para);					
					If n_Cursor>0 Then						
						CursorTableAppend( r_TabExt.id,n_Cursor,n_ParentCursor);
						n_Ret:=DBMS_SQL.fetch_rows(n_Cursor);
						If n_Ret=0 Then
							--无数据退出
							return(0);
						End If;
					Else
						--失败退出
						return(0);
					End If;
				End If;
				
				n_ParentCursor:=n_Cursor;
			End Loop;							
		End If;			
		Return(n_Cursor);
	End;

	--****************************************************************************************************************************************************
	-- 功    能：
	--****************************************************************************************************************************************************
	Function BuildSingleNodeData(id_In In zlmip_send_log.id%Type,Config_Express_In In VARCHAR2) Return Varchar2 
	Is
		---------------------------------------------------------------------------------------------------------------
		Cursor c_Item(n_id Varchar2) Is 
			Select b.item_code From zlmip_send_log a,zlmip_item b Where a.id=n_id And a.item_id=b.id;
		r_Item c_Item%Rowtype;

		---------------------------------------------------------------------------------------------------------------
		Cursor c_TriggerCond(n_id Varchar2,v_title Varchar2) Is 
			Select a.para_value From zlmip_sendlog_parameter a Where a.send_log_id=n_id And Upper(a.para_title)=Upper(v_title);
		r_TriggerCond c_TriggerCond%Rowtype;
		---------------------------------------------------------------------------------------------------------------
		v_Return			varchar2(4000);		
		n_Loop			Number(18);
		a_Return			t_Strlist := t_Strlist(); 
		u_Para			t_Strlist;
		n_Yes			Number:=0;
		v_SQL			Varchar2(1000);
		v_Target_id		Varchar2(32);
		v_Target_Field		Varchar2(30);
		v_EnterTable_id		Varchar2(32);
		n_Cursor			Number(18);
		n_ParentCursor		Number;
		n_Order			Number;
		n_Type			Number;		
		n_number			Number;
		v_varchar			varchar2(1000);
		d_date			date;
		v_Temp			Varchar(100);		
		---------------------------------------------------------------------------------------------------------------	
	Begin		
		v_Return:=Config_Express_In;

		a_Return:=GetElement(Config_Express_In);	

		For n_Loop In 1 .. a_Return.Count Loop
			---------------------------------------------------------------------------------------------------------------
			If a_Return(n_Loop)='当前时间' Then
				v_Return:=replace(v_Return,'['||a_Return(n_Loop)||']',To_Char(Sysdate,'yyyymmddhh24miss'));
			---------------------------------------------------------------------------------------------------------------
			ElsIf substr(a_Return(n_Loop),1,2)='项目标识' Then
				--从消息项目中提取信息
				Open c_Item(id_In);
				Fetch c_Item Into r_Item;
				If Not c_Item%NotFound Then
					v_Return:=replace(v_Return,'['||a_Return(n_Loop)||']',r_Item.item_code);
				End If;
				Close c_Item;
			---------------------------------------------------------------------------------------------------------------
			ElsIf substr(a_Return(n_Loop),1,2)='T.' Then
				--从触发条件结果中提取信息
				Open c_TriggerCond(id_In,Trim(substr(a_Return(n_Loop),3,100)));
				Fetch c_TriggerCond Into r_TriggerCond;
				If Not c_TriggerCond%NotFound Then
					v_Return:=replace(v_Return,'['||a_Return(n_Loop)||']',r_TriggerCond.para_value);
				End If;
				Close c_TriggerCond;
			Else
				--引用目标（可能是入口表id，可能是信息表关系id）及其字段（入口表的或关系对应的目标表的）
				v_Target_id:=substr(a_Return(n_Loop),1,instr(a_Return(n_Loop),'.')-1);
				v_Target_Field:=rtrim(substr(a_Return(n_Loop),instr(a_Return(n_Loop),'.')+1));
	
				--动态从临时表中读取此信息表对应的游标是否打开
				n_Cursor:=CursorTableSelect(v_Target_id);					
				----------------------------------------------------------------------------------------
				--没有打开时打开游标（包括消息树相关的游标）
				If n_Cursor=0 Then					
					n_Cursor:=OpenTable(id_In,v_Target_id);
				End If;
				----------------------------------------------------------------------------------------				
				If n_Cursor>0 Then

					--判断是否为入口表					
					Select Nvl(Max(1),0) Into n_Yes
					From zlmip_send_log a,zlmip_item c 
					Where a.id=id_In And a.item_id=c.id And c.tab_id=v_Target_id;

					n_Order:=0;
					n_Type:=0;
					Begin
						If n_Yes=1 Then
							Select b.fld_order,b.fld_type Into n_Order,n_Type 
							From zlmip_Table a,zlmip_Tab_Field b 
							Where a.id=v_Target_id And a.id=b.tab_id And Upper(b.fld_title)=Upper(v_Target_Field);
						Else
							Select b.fld_order,b.fld_type Into n_Order,n_Type 
							From zlmip_Tab_Extend c,zlmip_Table a,zlmip_Tab_Field b 
							Where c.id=v_Target_id And c.target_tab_id=a.id And a.id=b.tab_id And Upper(b.fld_title)=Upper(v_Target_Field);
						End If;
					Exception
						When Others Then n_Order:=0;
					End;

					If n_Order>0 Then
						If n_Type=1 Then			--数值型
							DBMS_SQL.column_value(n_Cursor, n_Order, n_number);
							v_Return:=replace(v_Return,'['||a_Return(n_Loop)||']',To_Char(n_number));
						Elsif n_Type=2 Then			--字符型
							DBMS_SQL.column_value(n_Cursor, n_Order, v_varchar);
							v_Return:=replace(v_Return,'['||a_Return(n_Loop)||']',v_varchar);
						Elsif n_Type=3 Then			--日期型
							DBMS_SQL.column_value(n_Cursor, n_Order, d_date);
							v_Return:=replace(v_Return,'['||a_Return(n_Loop)||']',To_Char(d_date,'yyyy-mm-dd hh24:mi:ss'));
						End If;	
					Else
						v_Return:=replace(v_Return,'['||a_Return(n_Loop)||']','');
					End If;
				Else
					v_Return:=replace(v_Return,'['||a_Return(n_Loop)||']','');
				End If;
			End If;
		End Loop;
		/*
		v_Return:=Replace(v_Return,'\n','\X000a');
		v_Return:=Replace(v_Return,'\r','\X000b');
		*/
		Return(v_Return);
	End;
	
	/*****************************************************************************************************************************************************/
	-- 功    能：
	/*****************************************************************************************************************************************************/
	Function BuildComplexNodeData(id_In In zlmip_send_log.id%Type,config_id_In In zlmip_item_config.id%Type) Return Varchar2 
	Is
		v_Return			Varchar2(4000);
		n_Position			Number(18);
		v_Target_id		Varchar2(32);
		n_Cursor			Number(18);
	Begin
		v_Return:='';

		--根据消息树的节点依次处理（结点有两种）
		For r_Row In (Select d.id,d.node_type,d.node_title,Nvl(d.config_occurs,'0') As config_occurs,d.config_occurs_key,config_express_key
						From zlmip_item_config d 
						Where d.parent_id=config_id_In order by d.node_order) 
		Loop	
			If r_Row.config_occurs='0' Then
				Null;
			ElsIf r_Row.config_occurs='1' Then
				--发生一次
				If r_Row.node_type=1 Then			--复合节点
					v_Return:=v_Return||'<'||r_Row.node_title||'>';
					v_Return:=v_Return||BuildComplexNodeData(id_In,r_Row.id);
					v_Return:=v_Return||'</'||r_Row.node_title||'>';
				ElsIf r_Row.node_type=2 Then			--单一节点
					v_Return:=v_Return||'<'||r_Row.node_title||'>';
					v_Return:=v_Return||BuildSingleNodeData(id_In,r_Row.config_express_key);	
					v_Return:=v_Return||'</'||r_Row.node_title||'>';
				End If;
			Else
				--发生多次
				n_Position:=Instr(Upper(r_Row.config_occurs_key),Upper('.Count'));				--如：[住院医嘱记录.Count]	
				If n_Position>0 Then
					v_Target_id:=substr(r_Row.config_occurs_key,2,n_Position-2);
					--先删除指定游标及其下级游标
					CursorTableDelete(v_Target_id);

					--重新打开指定游标
					n_Cursor:=OpenTable(id_In,v_Target_id);
					If n_Cursor>0 Then
						Loop								
							If r_Row.node_type=1 Then		--复合节点
								v_Return:=v_Return||'<'||r_Row.node_title||'>';
								v_Return:=v_Return||BuildComplexNodeData(id_In,r_Row.id);
								v_Return:=v_Return||'</'||r_Row.node_title||'>';

							Elsif r_Row.node_type=2 Then		--单一节点
								v_Return:=v_Return||'<'||r_Row.node_title||'>';
								v_Return:=v_Return||BuildSingleNodeData(id_In,r_Row.config_express_key);	
								v_Return:=v_Return||'</'||r_Row.node_title||'>';
							End If;
							
							Exit When DBMS_SQL.fetch_rows(n_Cursor) <= 0;
						End Loop;
					End If;
				End If;
			End If;
		End Loop;

		Return(v_Return);
	Exception
		When Others Then Zl_ErrorCenter (SQLCODE, SQLERRM);
	End BuildComplexNodeData;
	
	/*****************************************************************************************************************************************************/
	-- 功    能：
	/*****************************************************************************************************************************************************/
	Procedure BuildMessageContent(id_In In zlmip_send_log.id%Type,SendPara_In In Varchar2)
	Is
		v_MessageContent		Varchar2(4000);
		v_ItemIdentity			Varchar2(100);
		v_ItemTitle				Varchar2(100);
		v_MessageHead			Varchar2(4000);
		v_SendPara			Varchar2(2000);
	Begin
		v_MessageContent:='';

		CursorTableCreate();

		--根据消息树的节点依次处理（结点有两种）
		For r_Layer In (Select d.id,d.node_type,d.node_title,Nvl(d.config_occurs,'0') As config_occurs,d.config_occurs_key,d.config_express_key,a.item_code,a.item_title
						From zlmip_send_log e,zlmip_item_config d,zlmip_item a
						Where e.id=id_In And e.item_id=d.item_id And a.id=e.item_id And d.parent_id Is Null order by d.node_order) 
		Loop		
			v_ItemIdentity:=r_Layer.item_code;
			v_ItemTitle:=r_Layer.item_title;
			If r_Layer.node_type=1 Then			--复合节点				
				v_MessageContent:=v_MessageContent||'<'||r_Layer.node_title||'>';
				v_MessageContent:=v_MessageContent||BuildComplexNodeData(id_In,r_Layer.id);
				v_MessageContent:=v_MessageContent||'</'||r_Layer.node_title||'>';
			ElsIf r_Layer.node_type=2 Then			--单一节点				
				v_MessageContent:=v_MessageContent||'<'||r_Layer.node_title||'>';
				v_MessageContent:=v_MessageContent||BuildSingleNodeData(id_In,r_Layer.config_express_key);
				v_MessageContent:=v_MessageContent||'</'||r_Layer.node_title||'>';
			End If;
		End Loop;

		CursorTableDelete();
		If v_MessageContent Is Not Null Then
			
			v_MessageHead:='<head>';
			v_MessageHead:=v_MessageHead||'<'||lower(v_ItemIdentity)||'>'||v_ItemTitle||'</'||lower(v_ItemIdentity)||'>';
			v_MessageHead:=v_MessageHead||'<msg_id>'||id_In||'</msg_id>';
			v_MessageHead:=v_MessageHead||'<msg_item_code>'||v_ItemIdentity||'</msg_item_code>';
			
			v_SendPara:='';
			Select Extractvalue(xmltype(SendPara_In), '//send_station_name') Into v_SendPara From Dual;
			v_MessageHead:=v_MessageHead||'<send_station_name>'||v_SendPara||'</send_station_name>';
			
			v_SendPara:='';
			Select Extractvalue(xmltype(SendPara_In), '//send_station_ip') Into v_SendPara From Dual;
			v_MessageHead:=v_MessageHead||'<send_station_ip>'||v_SendPara||'</send_station_ip>';
			
			v_SendPara:='';
			Select Extractvalue(xmltype(SendPara_In), '//send_program') Into v_SendPara From Dual;
			v_MessageHead:=v_MessageHead||'<send_program>'||v_SendPara||'</send_program>';

			v_SendPara:='';
			Select Extractvalue(xmltype(SendPara_In), '//send_instance') Into v_SendPara From Dual;
			v_MessageHead:=v_MessageHead||'<send_instance>'||v_SendPara||'</send_instance>';
			v_MessageHead:=v_MessageHead||'<send_system_code></send_system_code>';
			v_MessageHead:=v_MessageHead||'<send_module_code></send_module_code>';
			
			v_SendPara:='';
			Select Extractvalue(xmltype(SendPara_In), '//send_mipuser') Into v_SendPara From Dual;
			v_MessageHead:=v_MessageHead||'<send_mipuser>'||v_SendPara||'</send_mipuser>';
			
			v_SendPara:='';
			Select Extractvalue(xmltype(SendPara_In), '//send_hisuser') Into v_SendPara From Dual;
			v_MessageHead:=v_MessageHead||'<send_hisuser>'||v_SendPara||'</send_hisuser>';

			v_MessageHead:=v_MessageHead||'<send_time>'||To_Char(Sysdate,'yyyy-mm-dd hh24:mi:ss')||'</send_time>';
			v_MessageHead:=v_MessageHead||'</head>';
			
			v_MessageContent:='<?xml version="1.0" encoding="gbk"?><message>'||v_MessageHead||v_MessageContent||'</message>';
		End If;
		Update zlmip_send_log Set send_content=xmltype(v_MessageContent) Where id=id_In;
	Exception
		When Others Then Zl_ErrorCenter (SQLCODE, SQLERRM);
	End BuildMessageContent;

	/*****************************************************************************************************************************************************/
	-- 功    能：
	/*****************************************************************************************************************************************************/
	Procedure RepeatMessageContent(id_In In zlmip_send_log.id%Type,SendPara_In In Varchar2,RepeatOrder_In In Number:=0)
	Is
		v_send_station_name		Varchar2(100);
		v_send_station_ip		Varchar2(20);
		v_send_program		Varchar2(100);
		v_send_instance			Varchar2(100);
		v_send_mipuser			Varchar2(100);
		v_send_hisuser			Varchar2(100);
	Begin		
		If RepeatOrder_In>0 Then
			v_send_station_name:='';
			v_send_station_ip:='';
			v_send_program:='';
			v_send_instance:='';
			v_send_mipuser:='';
			v_send_hisuser:='';
			Select Extractvalue(xmltype(SendPara_In), '//send_station_name') Into v_send_station_name From Dual;
			Select Extractvalue(xmltype(SendPara_In), '//send_station_ip') Into v_send_station_ip From Dual;
			Select Extractvalue(xmltype(SendPara_In), '//send_program') Into v_send_program From Dual;
			Select Extractvalue(xmltype(SendPara_In), '//send_instance') Into v_send_instance From Dual;
			Select Extractvalue(xmltype(SendPara_In), '//send_mipuser') Into v_send_mipuser From Dual;
			Select Extractvalue(xmltype(SendPara_In), '//send_hisuser') Into v_send_hisuser From Dual;
			Insert Into zlmip_sendlog_again(send_log_id,send_order,send_state,send_date,send_mipuser,send_user,send_station,send_note)
			Select id,RepeatOrder_In,send_state,Sysdate,v_send_mipuser,v_send_hisuser,v_send_station_name||'('||v_send_station_ip||')',Null From zlmip_send_log Where id=id_in;				
		End If;
	Exception
		When Others Then Zl_ErrorCenter (SQLCODE, SQLERRM);
	End RepeatMessageContent;

End zlMip_PkgRollService;
/

/*################################################################################################*/
--zlMip_PkgInterface 包头
Create Or Replace Package zltools.zlMip_PkgInterface Authid Current_User Is
	
	Type t_StrArray Is Table of Varchar2(4000);
	Type t_Refcur Is Ref Cursor;
	v_SQL		varchar2(4000);
	
	/*****************************************************************************************************************************************************/
	-- 功    能：解析投递目标
	/*****************************************************************************************************************************************************/
	Function GetDeliverTarget(MessageContent_In In Varchar2,DeliverLimit_In In Varchar2:=Null) Return Varchar2;
	
End zlMip_PkgInterface;
/

--zlMip_PkgInterface 包体
Create Or Replace Package Body zltools.zlMip_PkgInterface Is
	/**************************************************************************************************************************************************************************
	功能：内部函数
	**************************************************************************************************************************************************************************/	
	Function GetDataArray(Text In Varchar2,Split Varchar2) Return t_StrArray As
		v_Str   Long;
		v_Index Number; 
		v_List  t_StrArray := t_StrArray(); 
	Begin 
		v_Str:=Text||Split;
		Loop 
			v_Index := Instr(v_Str, Split); 
			Exit When(Nvl(v_Index, 0) = 0); 
			v_List.Extend; 
			v_List(v_List.Count) := Trim(Substr(v_Str, 1, v_Index - 1)); 
			v_Str := Substr(v_Str, v_Index + 1); 
		End Loop; 
		Return(v_List);
	End GetDataArray;

	/**************************************************************************************************************************************************************************
	功能：内部函数
	**************************************************************************************************************************************************************************/	
	Function GetNodeData(XmlText In xmltype,XmlNode Varchar2) Return Varchar2 As
		v_NodeData	varchar2(4000);
	Begin 
		Begin
			Select extractvalue(XmlText,XmlNode) Into v_NodeData From dual;
		Exception
			When Others Then v_NodeData:='';
		End;
		Return(v_NodeData);
	End GetNodeData;

	/**************************************************************************************************************************************************************************
	功能：内部函数
	**************************************************************************************************************************************************************************/	
	Function GetTargetIP(DeliverTarget_In varchar2,DeliverChild_In In varchar2,IP_In In varchar2) Return Varchar2 As
		v_Temp			varchar2(4000);
	Begin 
		v_Temp:=DeliverChild_In;
		If DeliverTarget_In Is Not Null Then
			If DeliverTarget_In='-' Then
				If Instr(','||DeliverChild_In||',',','||IP_In||',')=0 Then								
					v_Temp:=DeliverChild_In||IP_In||',';
				End If;
			Else
				If Instr(','||DeliverTarget_In||',',','||IP_In||',')>0 And Instr(','||DeliverChild_In||',',','||IP_In||',')=0 Then								
					v_Temp:=DeliverChild_In||IP_In||',';
				End If;				
			End If;
		Else
			v_Temp:='';
		End If;

		Return(v_Temp);
	End GetTargetIP;
	
	--****************************************************************************************************************************************************
	-- 功    能：内部函数
	--****************************************************************************************************************************************************
	Function GetCursor(ObjectType_In In Number,ParaValue_In In Varchar2,Owner_In In Varchar2) Return Number 
	Is
		n_Cursor				Number;
		n_State				Number;
		v_SQL				Varchar2(4000);
		v_varchar2				Varchar2(100);
		a_Para				t_StrArray := t_StrArray(); 
	Begin		
		---------------------------------------------------------------------------------------------------------------------
		If ObjectType_In=1 Then
			v_SQL:='Select  Distinct d.IP From '||Owner_In||'.上机人员表 b,zltools.zlmip_onlineusers c,zltools.zlClients d Where b.人员id=:personkey And b.用户名=c.loginuser And c.station=d.工作站';
		ElsIf ObjectType_In=2 Then
			v_SQL:='Select  Distinct d.IP From '||Owner_In||'.部门人员 a,'||Owner_In||'.上机人员表 b,zltools.zlmip_onlineusers c,zltools.zlClients d Where a.部门id=:deptkey And a.人员id=b.人员id And b.用户名=c.loginuser And c.station=d.工作站';
		ElsIf ObjectType_In=3 Then
			v_SQL:='Select  Distinct d.IP From '||Owner_In||'.人员性质说明 a,'||Owner_In||'.上机人员表 b,zltools.zlmip_onlineusers c,zltools.zlClients d Where a.人员性质=:property And a.人员id=b.人员id And b.用户名=c.loginuser And c.station=d.工作站';
		ElsIf ObjectType_In=4 Then
			v_SQL:='Select  Distinct d.IP From DBA_Role_Privs a,'||Owner_In||'.上机人员表 b,zltools.zlmip_onlineusers c,zltools.zlClients d Where a.Granted_Role=:role And a.Grantee=b.用户名 And b.用户名=c.loginuser And c.station=d.工作站';
		ElsIf ObjectType_In=6 Then
			v_SQL:='Select Distinct b.IP From zltools.zlmip_onlinemodules a,zltools.zlClients b Where a.system=:system And a.module=:module And a.station=b.工作站';
		End If;
		
		n_Cursor := DBMS_SQL.open_cursor;					--打开游标

		--解析游标---------------------------------------------------------------------------------------------------------------------
		DBMS_SQL.parse(n_Cursor, v_SQL, DBMS_SQL.native);	
		
		--绑定参数---------------------------------------------------------------------------------------------------------------------
		If ObjectType_In=1 Then			
			DBMS_SQL.bind_variable(n_Cursor, ':personkey',zl_to_number(ParaValue_In));
		ElsIf ObjectType_In=2 Then
			DBMS_SQL.bind_variable(n_Cursor, ':deptkey',zl_to_number(ParaValue_In));
		ElsIf ObjectType_In=3 Then
			DBMS_SQL.bind_variable(n_Cursor, ':property',Trim(ParaValue_In));
		ElsIf ObjectType_In=4 Then
			DBMS_SQL.bind_variable(n_Cursor, ':role',Trim(ParaValue_In));
		ElsIf ObjectType_In=6 Then
			a_Para:=GetDataArray(ParaValue_In,'_');
			DBMS_SQL.bind_variable(n_Cursor, ':system',zl_to_number(a_Para(1)));
			DBMS_SQL.bind_variable(n_Cursor, ':module',zl_to_number(a_Para(2)));
		End If;
		
		--定义列---------------------------------------------------------------------------------------------------------------------
		DBMS_SQL.define_column(n_Cursor, 1, v_varchar2,20);	
		
		--执行SQL---------------------------------------------------------------------------------------------------------------------
		n_State := DBMS_SQL.Execute(n_Cursor);			
		Return(n_Cursor);
	End;	
	
	/**************************************************************************************************************************************************************************
	功能：内部函数
	**************************************************************************************************************************************************************************/	
	Function GetChildDeliverTarget(ObjectType_In In Number,DeliverTarget_In varchar2,DeliverConfig_In In xmltype,MessageContent_In In xmltype,Owner_In In Varchar2) Return Varchar2 As
		n_Loop				Number(18); 
		n_Count				Number(18);
		n_Cursor				Number;
		n_Ret					Number;
		n_Pos1				Number;
		n_Pos2				Number;
		v_IP					Varchar2(20);
		v_productdata			Varchar2(50);
		v_messagedata			Varchar2(50);
		v_DeliverChild			varchar2(4000);
		v_Key				varchar2(4000);	
		v_Temp				Varchar2(4000);
		a_Keys				t_StrArray := t_StrArray(); 
		a_MsgKeys			t_StrArray := t_StrArray(); 

	Begin 
		--内容为 - 表示没有配置投递目标
		v_DeliverChild:='-';
		
		If ObjectType_In=1 Then
			v_productdata:='.//persons/person/key';
			v_messagedata:='.//persons/message';
		ElsIf ObjectType_In=2 Then
			v_productdata:='.//depts/dept/key';
			v_messagedata:='.//depts/message';
		ElsIf ObjectType_In=3 Then
			v_productdata:='.//personpropertys/personproperty';
			v_messagedata:='.//personpropertys/message';
		ElsIf ObjectType_In=4 Then
			v_productdata:='.//roles/role';
			v_messagedata:='.//roles/message';
		ElsIf ObjectType_In=5 Then
			v_productdata:='.//stations/station';
			v_messagedata:='.//stations/message';
		ElsIf ObjectType_In=6 Then
			v_productdata:='.//modules/module/key';
			v_messagedata:='.//modules/message';
		ElsIf ObjectType_In=7 Then
			v_productdata:='.//mipusers/mipuser';
			v_messagedata:='.//mipusers/message';
		End If;
		-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------				
		Select Reverse(Substr(Keys, 1, Instr(Keys, '/') - 1)) Into v_Key  From (Select Reverse(v_productdata) As Keys From Dual);
		Select extractValue(value(i),v_Key) Bulk Collect Into a_Keys From table(XMLSequence(extract(DeliverConfig_In,v_productdata))) i;
		If a_Keys.Count>0 Then
			For n_Loop In 1 .. a_Keys.Count Loop		
				If a_Keys(n_Loop) Is Not Null Then
					If v_DeliverChild='-' Then
						v_DeliverChild:='';
					End If;			
					If ObjectType_In=1 Or ObjectType_In=2 Or ObjectType_In=3 Or ObjectType_In=4 Or ObjectType_In=6 Then						
						n_Cursor:=GetCursor(ObjectType_In,a_Keys(n_Loop),Owner_In);
						n_Ret:=DBMS_SQL.fetch_rows(n_Cursor);
						If n_Ret>0 Then
							Loop
								DBMS_SQL.column_value(n_Cursor, 1, v_IP);
								v_IP:='IP|'''||v_IP||'''';
								v_DeliverChild:=GetTargetIP(DeliverTarget_In,v_DeliverChild,v_IP);
								Exit When DBMS_SQL.fetch_rows(n_Cursor) <= 0;
							End Loop;
						End If;
						DBMS_SQL.close_cursor(n_Cursor);
					ElsIf ObjectType_In=5 Then
						v_Temp:='IP|'''||a_Keys(n_Loop)||'''';
						n_Pos1:=instr(v_Temp,'(');
						n_Pos2:=instr(v_Temp,')');
						v_Temp:=substr(v_Temp,n_Pos1+1,n_Pos2-n_Pos1-1);
						v_DeliverChild:=GetTargetIP(DeliverTarget_In,v_DeliverChild,v_Temp);	
					ElsIf ObjectType_In=7 Then
						v_Temp:='USER|'''||a_Keys(n_Loop)||'''';
						v_DeliverChild:=GetTargetIP(DeliverTarget_In,v_DeliverChild,v_Temp);
					End If;
				End If;
			End Loop;
		End If;
		
		--来源于消息数据
		-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Select Reverse(Substr(Keys, 1, Instr(Keys, '/') - 1)) Into v_Key  From (Select Reverse(v_messagedata) As Keys From Dual);
		Select extractValue(value(i),v_Key) Bulk Collect Into a_Keys From table(XMLSequence(extract(DeliverConfig_In,v_messagedata))) i;
		If a_Keys.Count>0 Then
			For n_Loop In 1 .. a_Keys.Count Loop
				Select Reverse(Substr(Keys, 1, Instr(Keys, '/') - 1)) Into v_Key  From (Select Reverse(a_Keys(n_Loop)) As Keys From Dual);				
				Select extractValue(value(i),'/'||v_Key) Bulk Collect Into a_MsgKeys From table(XMLSequence(extract(MessageContent_In,'./'||a_Keys(n_Loop)))) i;
				If a_MsgKeys.Count>0 Then
					For n_Count In 1 .. a_MsgKeys.Count Loop		
						If a_MsgKeys(n_Count) Is Not Null Then
							If v_DeliverChild='-' Then
								v_DeliverChild:='';
							End If;
							If ObjectType_In=1 Or ObjectType_In=2 Or ObjectType_In=3 Or ObjectType_In=4 Or ObjectType_In=6 Then
								n_Cursor:=GetCursor(ObjectType_In,a_MsgKeys(n_Count),Owner_In);
								n_Ret:=DBMS_SQL.fetch_rows(n_Cursor);
								If n_Ret>0 Then
									Loop
										DBMS_SQL.column_value(n_Cursor, 1, v_IP);
										v_IP:='IP|'''||v_IP||'''';
										v_DeliverChild:=GetTargetIP(DeliverTarget_In,v_DeliverChild,v_IP);
										Exit When DBMS_SQL.fetch_rows(n_Cursor) <= 0;
									End Loop;
								End If;
								DBMS_SQL.close_cursor(n_Cursor);
							ElsIf ObjectType_In=5 Then
								v_Temp:='IP|'''||a_MsgKeys(n_Count)||'''';
								v_DeliverChild:=GetTargetIP(DeliverTarget_In,v_DeliverChild,v_Temp);
							ElsIf ObjectType_In=7 Then
								v_Temp:='USER|'''||a_MsgKeys(n_Count)||'''';
								v_DeliverChild:=GetTargetIP(DeliverTarget_In,v_DeliverChild,v_Temp);
							End If;
						End If;
					End Loop;
				Else
					--无节点时，认为不符合要求
					v_DeliverChild:='';
				End If;
			End Loop;
		End If;
		-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------		
		If v_DeliverChild Is Not Null Then
			If v_DeliverChild<>'-' Then
				v_DeliverChild:=substr(v_DeliverChild,1,length(v_DeliverChild)-1);
			End If;
		End If;			
		Return(v_DeliverChild);
	End GetChildDeliverTarget;
	
	/*****************************************************************************************************************************************************/
	-- 功    能：解析投递目标
	/*****************************************************************************************************************************************************/
	Function GetDeliverTarget(MessageContent_In In Varchar2,DeliverLimit_In In Varchar2:=Null) Return Varchar2 Is		
		v_DeliverTemp			Varchar2(4000);		
		v_DeliverGroup			Varchar2(4000);
		v_DeliverGroup_IP		Varchar2(4000);
		v_DeliverGroup_USER		Varchar2(4000);
		v_DeliverResult			Varchar2(4000);
		v_DeliverResult_IP		Varchar2(4000);
		v_DeliverResult_USER		Varchar2(4000);
		v_ItemKey				Varchar2(64);
		v_Owner				Varchar2(20);
		n_Loop				Number(18); 
		n_AllUser				Number(1);
		x_MessageContent		Xmltype;
		a_Array				t_StrArray := t_StrArray(); 
		a_Element				t_StrArray := t_StrArray(); 
		/***************************************************************************************************************************************************************************/	
	Begin
		--Return(MessageContent_In);
		If MessageContent_In Is Null Then
			Return('');
		End If;
		
		If Substr(MessageContent_In,1,1)<>'<' Then
			Return('');
		End If;
		
		n_AllUser:=0;

		Select xmltype(MessageContent_In) Into x_MessageContent From Dual;		
		v_ItemKey:=GetNodeData(x_MessageContent,'.//head/msg_item_code');
				
		Select Nvl(Max(b.data_owner),'ZLHIS') Into v_Owner From zltools.zlmip_item a,zltools.zlmip_data_setup b Where a.item_code=v_ItemKey and a.data_code=b.data_code;
		
		v_DeliverResult:='-';		--'-'表示没有配置
		For r_Group In (Select b.deliver_object From zltools.zlmip_item a,zltools.zlmip_item_deliver b Where a.item_code=v_ItemKey And a.id=b.item_id And (b.deliver_code=DeliverLimit_In Or DeliverLimit_In Is Null) And b.deliver_object Is Not Null Order By b.deliver_order)
		Loop
			--每次循环为一组条件，必须同时满足	
			v_DeliverGroup:='-';
			/*----------------------------------------------------------------------------------------------------------------------------------------------------------------------*/	
			For n_Loop In 1 .. 7 Loop	
				If v_DeliverGroup Is Not Null Then					
					v_DeliverTemp:=GetChildDeliverTarget(n_Loop,v_DeliverGroup,r_Group.deliver_object,x_MessageContent,v_Owner);
					If v_DeliverTemp Is Null Then
						v_DeliverGroup:='';
					ElsIf v_DeliverTemp<>'-' Then		
						v_DeliverGroup:=v_DeliverTemp;
					End If;
				End If;
			End Loop;
			
			--分析组中的各个投递对象(IP,USER)
			If v_DeliverGroup Is Not Null Or v_DeliverGroup<>'-' Then				
				v_DeliverGroup_IP:='';
				v_DeliverGroup_USER:='';

				a_Array:=GetDataArray(v_DeliverGroup,',');
				For n_Loop In 1 .. a_Array.Count Loop
					a_Element:=GetDataArray(a_Array(n_Loop),'|');
					If a_Element(1)='IP' Then
						v_DeliverGroup_IP:=v_DeliverGroup_IP||','||a_Element(2);
					ElsIf a_Element(1)='USER' Then
						If Upper(a_Element(2))='ALL' Then
							--发给所有用户，其他的不再分析
							n_AllUser:=1;
							Exit;
						End If;
						v_DeliverGroup_USER:=v_DeliverGroup_USER||','||a_Element(2);
					End If;
				End Loop;
				
				If n_AllUser=1 Then
					Exit;
				End If;

				If v_DeliverGroup_IP Is Not Null Then
					v_DeliverGroup_IP:=substr(v_DeliverGroup_IP,2,length(v_DeliverGroup_IP)-1);
				End If;
				If v_DeliverGroup_USER Is Not Null Then
					v_DeliverGroup_USER:=substr(v_DeliverGroup_USER,2,length(v_DeliverGroup_USER)-1);
				End If;
			End If;

			--合并组到最后的IP中
			If v_DeliverResult_IP='-' Or v_DeliverResult_IP Is Null Then
				v_DeliverResult_IP:=v_DeliverGroup_IP;
			Else
				v_DeliverTemp:='';
				For r_IP In (Select Column_Value As IP From Table(Cast(f_str2list(v_DeliverResult_IP) As zlTools.t_Strlist))
						Union 
						Select Column_Value As IP From Table(Cast(f_str2list(v_DeliverGroup_IP) As zlTools.t_Strlist)))
				Loop
					v_DeliverTemp:=v_DeliverTemp||','||r_IP.IP;
				End Loop;	

				If v_DeliverTemp Is Null Then
					v_DeliverResult_IP:='';
				Else
					v_DeliverResult_IP:=substr(v_DeliverTemp,2,length(v_DeliverTemp)-1);
				End If;
			End If;			

			--合并组到最后的USER中
			If v_DeliverResult_USER='-' Or v_DeliverResult_USER Is Null Then
				v_DeliverResult_USER:=v_DeliverGroup_USER;
			Else
				v_DeliverTemp:='';
				For r_USER In (Select Column_Value As MIPUSER From Table(Cast(f_str2list(v_DeliverResult_USER) As zlTools.t_Strlist))
						Union 
						Select Column_Value As MIPUSER From Table(Cast(f_str2list(v_DeliverGroup_USER) As zlTools.t_Strlist)))
				Loop
					v_DeliverTemp:=v_DeliverTemp||','||r_USER.MIPUSER;
				End Loop;	

				If v_DeliverTemp Is Null Then
					v_DeliverResult_USER:='';
				Else
					v_DeliverResult_USER:=substr(v_DeliverTemp,2,length(v_DeliverTemp)-1);
				End If;
			End If;			

		End Loop;
		/*----------------------------------------------------------------------------------------------------------------------------------------------------------------------*/

		If n_AllUser=1 Then
			v_DeliverResult:='all';
		Else 
			If v_DeliverResult_IP Is Not Null  Or v_DeliverResult_USER Is Not Null Then			
				v_DeliverResult:='condition|';
				If v_DeliverResult_IP Is Not Null And v_DeliverResult_IP<>'-' Then
					v_DeliverResult:=v_DeliverResult||'ip in ('||v_DeliverResult_IP||')';
				End If;

				If v_DeliverResult_USER Is Not Null And v_DeliverResult_USER<>'-' Then
					If v_DeliverResult<>'condition|' Then
						v_DeliverResult:=v_DeliverResult||' or ';
					End If;
					v_DeliverResult:=v_DeliverResult||'user in ('||v_DeliverResult_USER||')';
				End If;
			End If;
		End If;
		Return(v_DeliverResult);
	Exception
		WHEN OTHERS THEN Zl_ErrorCenter (SQLCODE, SQLERRM);
	End GetDeliverTarget;	
	
End zlMip_PkgInterface;
/

--重新编译部分过程,这些过程可能在安装完成之后无效。
Alter Package zltools.zlMip_PkgManage Compile
/
Alter Package zltools.zlMip_PkgRollService Compile
/
Alter Package zltools.zlMip_PkgInterface Compile
/

/*
-----------------------------------------------------------------------------------------------------------------------
--对象授权部分
-----------------------------------------------------------------------------------------------------------------------
Grant Execute On zltools.zlMip_OnLineUser_Insert To public;
Grant Execute On zltools.zlMip_OnLineUser_Delete To public;
Grant Execute On zltools.zlMip_OnLineModule_Insert To public;
Grant Execute On zltools.zlMip_OnLineModule_Delete To public;
Grant Execute On zltools.zlMip_PkgManage To public;
Grant Execute On zltools.zlMip_PkgInterface To public;
Grant Execute On zltools.zlMip_PkgRollService To public;
-----------------------------------------------------------------------------------------------------------------------
--对象同义词部分
-----------------------------------------------------------------------------------------------------------------------
Create Public Synonym zlMip_OnLineUser_Insert for zltools.zlMip_OnLineUser_Insert;
Create Public Synonym zlMip_OnLineUser_Delete for zltools.zlMip_OnLineUser_Delete;
Create Public Synonym zlMip_OnLineModule_Insert for zltools.zlMip_OnLineModule_Insert;
Create Public Synonym zlMip_OnLineModule_Delete for zltools.zlMip_OnLineModule_Delete;
Create Public Synonym zlMip_PkgManage for zltools.zlMip_PkgManage;
Create Public Synonym zlMip_PkgInterface for zltools.zlMip_PkgInterface;
Create Public Synonym zlMip_PkgRollService for zltools.zlMip_PkgRollService;
Create Public Synonym zlMip_Clob_Append for zltools.zlMip_Clob_Append;
Create Public Synonym zlMip_ReceiveLog_Insert for zltools.zlMip_ReceiveLog_Insert;
Create Public Synonym zlMip_SendLog_Insert for zltools.zlMip_SendLog_Insert;
Create Public Synonym zlMip_SendLog_UpdateState for zltools.zlMip_SendLog_UpdateState;
*/