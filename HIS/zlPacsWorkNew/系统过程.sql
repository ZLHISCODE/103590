--标本核收====================================================================================================


--添加标本信息
CREATE OR REPLACE function Zl_病理标本_新增
(
  医嘱ID_IN   病理标本信息.医嘱ID%Type,     
  标本名称_IN 病理标本信息.标本名称%Type,
  标本类型_IN 病理标本信息.标本类型%Type,
  采集部位_IN 病理标本信息.采集部位%Type,
  标本数量_IN 病理标本信息.数量%Type,
  材料类别_IN 病理标本信息.材料类别%Type,
  原有编号_IN 病理标本信息.原有编号%Type,
  存放位置_IN 病理标本信息.存放位置%Type,
  接收日期_IN 病理标本信息.接收日期%Type,
  备注信息_IN 病理标本信息.备注%Type
) return number Is
PRAGMA AUTONOMOUS_TRANSACTION;

v_标本ID 病理标本信息.标本ID%Type;

Begin
  select 病理标本信息_标本ID.NEXTVAL into  v_标本ID  from dual;
     
  insert into 病理标本信息(标本ID,医嘱ID,标本名称,标本类型,采集部位,数量,材料类别,原有编号,存放位置,接收日期,备注)
  values(v_标本ID, 医嘱ID_IN, 标本名称_IN, 标本类型_IN, 采集部位_IN, 标本数量_IN, 材料类别_IN, 原有编号_IN, 存放位置_IN, 接收日期_IN, 备注信息_IN);
  
  commit;
  
  return  v_标本ID;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病理标本_新增;
/


--更新标本信息
CREATE OR REPLACE Procedure Zl_病理标本_更新
(
  标本ID_IN   病理标本信息.标本ID%Type,     
  标本名称_IN 病理标本信息.标本名称%Type,
  标本类型_IN 病理标本信息.标本类型%Type,
  采集部位_IN 病理标本信息.采集部位%Type,
  标本数量_IN 病理标本信息.存放位置%Type,
  材料类别_IN 病理标本信息.材料类别%Type,
  原有编号_IN 病理标本信息.原有编号%Type,
  存放位置_IN 病理标本信息.存放位置%Type,
  备注信息_IN 病理标本信息.备注%Type
) Is
Begin
  update 病理标本信息 
  set 标本名称=标本名称_IN,标本类型=标本类型_IN,采集部位=采集部位_IN,
      数量=标本数量_IN,材料类别=材料类别_IN,原有编号=原有编号_IN,
      存放位置=存放位置_IN,备注=备注信息_IN
  where 标本ID=标本ID_IN;   
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病理标本_更新;
/

--核收送检标本
CREATE OR REPLACE Procedure Zl_病理标本_核收
(   
  医嘱ID_IN   病理检查信息.医嘱ID%Type,  
  检查类型_IN 病理检查信息.检查类型%Type,   
  送检单位_IN 病理送检信息.送检单位%Type,
  送检科室_IN 病理送检信息.送检科室%Type,
  送检人_IN   病理送检信息.送检人%Type,
  送检日期_IN 病理送检信息.送检日期%Type,
  联系方式_IN 病理送检信息.联系方式%Type,
  登记人_IN   病理送检信息.登记人%Type
) Is
  v_病理号 病理检查信息.病理号%Type := null;
  v_检查类型 病理检查信息.检查类型%Type;
Begin

  begin
    select 病理号 into v_病理号 from 病理检查信息 where 医嘱ID=医嘱ID_IN;      
  exception
    When Others Then v_病理号 := null;	
  end;              
     
  if v_病理号 is null then    
     --没有找到该医嘱对应的病理检查
     
     --生成病理号
     Select Lpad(病理检查信息_病理号.NEXTVAL, 8, 0) into v_病理号 from dual;  
      
     --取得当前病理检查类型
     v_检查类型 := 检查类型_IN;
  
     --添加病理送检信息
     insert into 病理送检信息(ID, 医嘱ID,送检单位,送检科室,送检人,送检日期,联系方式,登记人,核收状态)
     values(病理送检信息_ID.NEXTVAL, 医嘱ID_IN, 送检单位_IN, 送检科室_IN, 送检人_IN, 送检日期_IN, 联系方式_IN, 登记人_IN, 1);
     
     --添加病理检查信息,核收后，检查进入取材流程
     insert into 病理检查信息(病理号, 医嘱ID, 检查类型, 当前过程)
     values(v_病理号, 医嘱ID_IN, v_检查类型, decode(v_检查类型, 3, 3, 1));
  else
    --当该检查已被核收过时，则只添加送检信息   
     insert into 病理送检信息(ID, 医嘱ID,送检单位,送检科室,送检人,送检日期,联系方式,登记人,核收状态)
     values(病理送检信息_ID.NEXTVAL, 医嘱ID_IN, 送检单位_IN, 送检科室_IN, 送检人_IN, 送检日期_IN, 联系方式_IN, 登记人_IN, 1);    
  end if;  
  
  --更新对应医嘱的执行说明...             
  update 病人医嘱发送 
  set 执行说明=执行说明 || chr(13) || '标本已被核收 [ 时间:'|| 送检日期_IN  || '    登记人:' || 登记人_IN || '] '
  where 医嘱ID=医嘱ID_IN;  
  
  --更新执行过程
  update 病人医嘱发送 set 执行过程=2, 报到时间=送检日期_IN where 医嘱ID=医嘱ID_IN and nvl(执行过程,0) < 2;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病理标本_核收;
/

--拒收送检标本 
CREATE OR REPLACE Procedure Zl_病理标本_拒收
(
  医嘱ID_IN   病理检查信息.医嘱ID%Type,     
  送检单位_IN 病理送检信息.送检单位%Type,
  送检科室_IN 病理送检信息.送检科室%Type,
  送检人_IN   病理送检信息.送检人%Type,
  送检日期_IN 病理送检信息.送检日期%Type,
  联系方式_IN 病理送检信息.联系方式%Type,
  登记人_IN   病理送检信息.登记人%Type,
  拒收原因_IN 病理送检信息.拒收原因%Type,
  通知人_IN   病理送检信息.通知人%Type
) Is
Begin
     insert into 病理送检信息(ID, 医嘱ID,送检单位,送检科室,送检人,送检日期,联系方式,登记人,核收状态, 拒收原因, 通知人)
     values(病理送检信息_ID.NEXTVAL, 医嘱ID_IN, 送检单位_IN, 送检科室_IN, 送检人_IN, 送检日期_IN, 联系方式_IN, 登记人_IN, 2, 拒收原因_IN, 通知人_IN); 
     
     --更新对应医嘱的执行说明...   
     update 病人医嘱发送 
     set 执行说明=执行说明 || chr(13) || '标本已被拒收 [ 时间:'|| 送检日期_IN || '   拒收原因:' || 拒收原因_IN || '    登记人:' || 登记人_IN || '] '
     where 医嘱ID=医嘱ID_IN;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病理标本_拒收;
/






--抗体管理====================================================================================================





--新增抗体
CREATE OR REPLACE function Zl_病理抗体_新增
(
  抗体名称_IN 病理抗体信息.抗体名称%Type,     
  使用人份_IN 病理抗体信息.使用人份%Type,
  已用人份_IN 病理抗体信息.已用人份%Type,
  生产日期_IN 病理抗体信息.生产日期%Type,
  有效期_IN   病理抗体信息.有效期%Type,
  过期日期_IN 病理抗体信息.过期日期%Type,
  克隆性_IN   病理抗体信息.克隆性%Type,
  作用对象_IN 病理抗体信息.作用对象%Type,
  理化性质_IN 病理抗体信息.理化性质%Type,
  应用情况_IN 病理抗体信息.应用情况%Type,
  登记人_IN   病理抗体信息.登记人%Type,
  登记时间_IN 病理抗体信息.登记时间%Type,
  备注_IN     病理抗体信息.备注%Type
) return Number Is
PRAGMA AUTONOMOUS_TRANSACTION;
v_抗体ID 病理抗体信息.抗体ID%Type;
Begin
  select 病理抗体信息_抗体ID.NEXTVAL into v_抗体ID from dual;
       
  insert into 病理抗体信息(抗体ID,抗体名称,使用人份,已用人份,生产日期,有效期,过期日期,克隆性,作用对象,理化性质,应用情况,登记人,登记时间,使用状态,备注)
  values(v_抗体ID, 抗体名称_IN, 使用人份_IN, 已用人份_IN, 生产日期_IN, 
         有效期_IN, 过期日期_IN, 克隆性_IN, 作用对象_IN, 理化性质_IN, 应用情况_IN,登记人_IN,登记时间_IN,1,备注_IN);
         
  commit;                  
         
  return v_抗体ID;       
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病理抗体_新增;
/

--已使用过的抗体，不能更新抗体名称等内容
CREATE OR REPLACE Procedure Zl_病理抗体_更新
(
  抗体ID_IN   病理抗体信息.抗体ID%Type,           
  抗体名称_IN 病理抗体信息.抗体名称%Type,     
  使用人份_IN 病理抗体信息.使用人份%Type,
  已用人份_IN 病理抗体信息.已用人份%Type,
  生产日期_IN 病理抗体信息.生产日期%Type,
  有效期_IN   病理抗体信息.有效期%Type,
  过期日期_IN 病理抗体信息.过期日期%Type,
  克隆性_IN   病理抗体信息.克隆性%Type,
  作用对象_IN 病理抗体信息.作用对象%Type,
  理化性质_IN 病理抗体信息.理化性质%Type,
  应用情况_IN 病理抗体信息.应用情况%Type,
  登记人_IN   病理抗体信息.登记人%Type,
  登记时间_IN 病理抗体信息.登记时间%Type,
  备注_IN     病理抗体信息.备注%Type
) Is
Begin
  update 病理抗体信息
  set 抗体名称=抗体名称_IN, 使用人份=使用人份_IN,已用人份=已用人份_IN,生产日期=生产日期_IN,
      有效期=有效期_IN,过期日期=过期日期_IN,克隆性=克隆性_IN,作用对象=作用对象_IN,
      理化性质=理化性质_IN,应用情况=应用情况_IN,登记人=登记人_IN,登记时间=登记时间_IN,备注=备注_IN
  where 抗体ID=抗体ID_IN;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病理抗体_更新;
/

--更新抗体使用状态
CREATE OR REPLACE Procedure Zl_病理抗体_使用状态
(
  抗体ID_IN 病理抗体信息.抗体ID%Type,           
  使用状态_IN 病理抗体信息.使用状态%Type
) Is
Begin
  update 病理抗体信息 set 使用状态=使用状态_IN where 抗体ID=抗体ID_IN;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病理抗体_使用状态;
/

--删除抗体信息
CREATE OR REPLACE Procedure Zl_病理抗体_删除
(
  抗体ID_IN 病理抗体信息.抗体ID%Type
) Is
Begin
  delete 病理抗体信息 where 抗体ID=抗体ID_IN;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病理抗体_删除; 
/


--新增抗体反馈
CREATE OR REPLACE function Zl_病理抗体反馈_新增
(
  抗体ID_IN      病理抗体反馈.抗体ID%Type,   
  参考病理号_IN  病理抗体反馈.参考病理号%Type,
  实验类型_IN    病理抗体反馈.实验类型%Type,
  抗体评价_IN    病理抗体反馈.抗体评价%Type,
  反馈时间_IN    病理抗体反馈.反馈时间%Type,
  反馈医生_IN    病理抗体反馈.反馈医生%Type,
  反馈意见_IN    病理抗体反馈.反馈意见%Type
) return number Is
PRAGMA AUTONOMOUS_TRANSACTION;
v_id 病理抗体反馈.ID%Type;
Begin
  select 病理抗体反馈_Id.Nextval into v_id from dual;
  
  insert into 病理抗体反馈(ID, 抗体ID,参考病理号,实验类型,抗体评价,反馈时间,反馈医生,反馈意见)
  values(v_id, 抗体ID_IN, 参考病理号_IN, 实验类型_IN, 抗体评价_IN, 反馈时间_IN, 反馈医生_IN, 反馈意见_IN);
  
  commit;
  
  return v_id;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病理抗体反馈_新增;
/

--更新抗体反馈信息
CREATE OR REPLACE Procedure Zl_病理抗体反馈_更新
(
  ID_IN          病理抗体反馈.ID%Type,   
  参考病理号_IN  病理抗体反馈.参考病理号%Type,
  实验类型_IN    病理抗体反馈.实验类型%Type,
  抗体评价_IN    病理抗体反馈.抗体评价%Type,
  反馈时间_IN    病理抗体反馈.反馈时间%Type,
  反馈医生_IN    病理抗体反馈.反馈医生%Type,
  反馈意见_IN    病理抗体反馈.反馈意见%Type
) Is
Begin
  Update 病理抗体反馈
  set 参考病理号=参考病理号_IN,实验类型=实验类型_IN,抗体评价=抗体评价_IN,
       反馈时间=反馈时间_IN,反馈医生=反馈医生_IN,反馈意见=反馈意见_IN
  where ID=ID_IN;     
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病理抗体反馈_更新;
/


--删除反馈记录
CREATE OR REPLACE Procedure Zl_病理抗体反馈_删除
(
  ID_IN 病理抗体反馈.ID%Type
) Is
Begin
  Delete 病理抗体反馈 where ID=ID_IN;   
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病理抗体反馈_删除;
/


--新增病理套餐
CREATE OR REPLACE function Zl_病理套餐_新增
(
  套餐名称_IN   病理套餐信息.套餐名称%Type,   
  套餐说明_IN   病理套餐信息.套餐说明%Type,
  创建时间_IN   病理套餐信息.创建时间%Type,
  创建人_IN     病理套餐信息.创建人%Type
) return number Is
PRAGMA AUTONOMOUS_TRANSACTION;
v_id 病理套餐信息.套餐ID%Type;
Begin
  select 病理套餐信息_套餐ID.Nextval into v_id from dual;
  
  insert into 病理套餐信息(套餐ID,套餐名称,套餐说明,创建人,创建时间)
  values(v_id, 套餐名称_IN, 套餐说明_IN, 创建人_IN, 创建时间_IN);
  
  commit;
  
  return v_id;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病理套餐_新增;
/


--更新病理套餐
CREATE OR REPLACE procedure Zl_病理套餐_更新
(
  套餐ID_IN     病理套餐信息.套餐ID%Type,     
  套餐名称_IN   病理套餐信息.套餐名称%Type,   
  套餐说明_IN   病理套餐信息.套餐说明%Type
)Is
Begin
  --更新套餐信息
  update  病理套餐信息 set 套餐名称=套餐名称_IN, 套餐说明=套餐说明_IN where 套餐ID=套餐ID_IN;

Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病理套餐_更新;
/



--删除病理套餐
CREATE OR REPLACE procedure Zl_病理套餐_删除
(
  套餐ID_IN     病理套餐信息.套餐ID%Type
)Is
Begin
  --删除套餐信息
  delete 病理套餐信息 where 套餐ID=套餐ID_IN;

Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病理套餐_删除;
/
 

--创建套餐抗体关联
CREATE OR REPLACE function Zl_病理套餐关联_新增
(
  套餐ID_IN   病理套餐关联.套餐ID%Type,   
  抗体ID_IN   病理套餐关联.抗体ID%Type
) return number Is
PRAGMA AUTONOMOUS_TRANSACTION;
v_id 病理套餐关联.ID%Type;
Begin
  select 病理套餐关联_Id.Nextval into v_id from dual;
  
  insert into 病理套餐关联(ID, 套餐ID,抗体ID) values(v_id, 套餐ID_IN, 抗体ID_IN);
  
  commit;
  
  return v_id;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病理套餐关联_新增;
/


--根据套餐ID,删除套餐关联的抗体
CREATE OR REPLACE procedure Zl_病理套餐关联_删除
(
  套餐ID_IN   病理套餐关联.套餐ID%Type
)Is
Begin
  --删除关联的套餐信息          
  delete 病理套餐关联 where 套餐ID=套餐ID_IN;
  
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病理套餐关联_删除;
/


--根据关联ID删除套餐关联的抗体
CREATE OR REPLACE procedure Zl_病理套餐关联_删除1
(
  套餐关联ID_IN   病理套餐关联.ID%Type
)Is
Begin
  --删除关联的套餐信息          
  delete 病理套餐关联 where ID=套餐关联ID_IN;
  
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病理套餐关联_删除1;
/
 


--标本取材====================================================================================================


--病理脱钙
CREATE OR REPLACE function Zl_病理脱钙_开始
(
  标本ID_IN    病理脱钙信息.标本ID%Type,   
  开始时间_IN  病理脱钙信息.开始时间%Type,
  所需时长_IN  病理脱钙信息.所需时长%Type,
  操作员_IN    病理脱钙信息.操作员%Type
) return number Is
PRAGMA AUTONOMOUS_TRANSACTION;
v_id 病理脱钙信息.ID%Type;
Begin
  select 病理脱钙信息_Id.Nextval into v_id from dual;
  
  insert into 病理脱钙信息(ID, 标本ID,开始时间,所需时长,当前缸次,操作员,完成状态)
  values(v_id, 标本ID_IN, 开始时间_IN, 所需时长_IN, 1, 操作员_IN, 0);
  
  commit;
  
  return v_id;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病理脱钙_开始;
/



--病理脱钙换缸
CREATE OR REPLACE procedure Zl_病理脱钙_换缸
(
  ID_IN        病理脱钙信息.ID %Type,       
  开始时间_IN  病理脱钙信息.开始时间%Type,
  所需时长_IN  病理脱钙信息.所需时长%Type
)Is
Begin
  
  update 病理脱钙信息
  set 开始时间=开始时间_IN,所需时长=所需时长_IN,当前缸次=当前缸次+1
  where ID=ID_IN;
  
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病理脱钙_换缸;
/


--病理脱钙撤销
CREATE OR REPLACE procedure Zl_病理脱钙_撤销
(
  ID_IN        病理脱钙信息.ID %Type
) Is
Begin
  
  Delete 病理脱钙信息 where ID=ID_IN and 完成状态=0;
  
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病理脱钙_撤销;
/


--病理脱钙完成
CREATE OR REPLACE procedure Zl_病理脱钙_完成
(
  ID_IN        病理脱钙信息.ID %Type
)Is
Begin
  
  update 病理脱钙信息 set 完成状态=1 where ID=ID_IN;
  
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病理脱钙_完成;
/




--病理常规取材
CREATE OR REPLACE function Zl_病理取材_常规
(
  病理号_IN      病理取材信息.病理号%Type,   
  申请ID_IN      病理取材信息.申请ID%Type,
  标本ID_IN      病理取材信息.标本ID%Type,
  标本名称_IN    病理取材信息.标本名称%Type,
  取材位置_IN    病理取材信息.取材位置%Type,
  蜡块数_IN      病理取材信息.蜡块数%Type,
  主取医师_IN    病理取材信息.主取医师%Type,
  副取医师_IN    病理取材信息.副取医师%Type,
  记录医师_IN    病理取材信息.记录医师%Type,
  取材时间_IN    病理取材信息.取材时间%Type  
) return varchar2 Is
PRAGMA AUTONOMOUS_TRANSACTION;

v_id 病理取材信息.材块ID%Type;
v_seqNum 病理取材信息.序号%Type;

Begin                        
  --获取最大材块号序号  
  begin
    select  nvl(max(序号), 0) into v_seqNum from 病理取材信息 where 病理号=病理号_IN;
  exception
    When Others Then v_id := 0;	            
  end;    
  
  v_seqNum := v_seqNum + 1;
  select 病理取材信息_材块ID.Nextval into v_id from dual;
  
  --写入取材记录    
  insert into 病理取材信息(材块ID, 序号, 病理号, 申请ID, 标本ID, 标本名称,取材位置,蜡块数,主取医师,副取医师,记录医师,取材时间)
  values(v_id, v_seqNum, 病理号_IN, 申请ID_IN, 标本ID_IN, 标本名称_IN, 取材位置_IN, 蜡块数_IN,主取医师_IN,副取医师_IN,记录医师_IN,取材时间_IN);
  
  --写入制片记录
  insert into 病理制片信息(ID,病理号,材块ID,申请ID,制片类型,制片方式,制片数,当前状态)
  values(病理制片信息_ID.NEXTVAL,病理号_IN,v_id,申请ID_IN,0,0,蜡块数_IN,0);
  
  commit; 
  
  return v_id || '-' || v_seqNum;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病理取材_常规;
/



CREATE OR REPLACE Procedure Zl_病理取材_常规更新
(
  材块ID_IN      病理取材信息.病理号%Type,   
  取材位置_IN    病理取材信息.取材位置%Type,
  蜡块数_IN      病理取材信息.蜡块数%Type,
  主取医师_IN    病理取材信息.主取医师%Type,
  副取医师_IN    病理取材信息.副取医师%Type 
)Is
Begin
  --更新取材信息      
  update 病理取材信息
  set 取材位置=取材位置_IN,蜡块数=蜡块数_IN,主取医师=主取医师_IN,副取医师=副取医师_IN
  where 材块ID=材块ID_IN;

  --更新制片信息
  update 病理制片信息 set 制片数=蜡块数_IN where 材块ID=材块ID_IN and 当前状态=0;
  
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病理取材_常规更新;
/



--病理细胞取材
CREATE OR REPLACE function Zl_病理取材_细胞
(
  病理号_IN      病理取材信息.病理号%Type,   
  申请ID_IN      病理取材信息.申请ID%Type,
  标本ID_IN      病理取材信息.标本ID%Type,
  标本名称_IN    病理取材信息.标本名称%Type,
  形状_IN        病理取材信息.形状%Type,
  颜色_IN        病理取材信息.颜色%Type, 
  性质_IN        病理取材信息.性质%Type,   
  标本量_IN      病理取材信息.标本量%Type,
  细胞块数_IN    病理取材信息.蜡块数%Type,
  主取医师_IN    病理取材信息.主取医师%Type,
  副取医师_IN    病理取材信息.副取医师%Type,
  记录医师_IN    病理取材信息.记录医师%Type,
  取材时间_IN    病理取材信息.取材时间%Type  
) return varchar2 Is
PRAGMA AUTONOMOUS_TRANSACTION;

v_id 病理取材信息.材块ID%Type;
v_seqNum 病理取材信息.序号%Type;
v_slicesCount number;

Begin
  --获取最大材块号序号  
  begin
    select  nvl(max(序号), 0) into v_seqNum from 病理取材信息 where 病理号=病理号_IN;
  exception
    When Others Then v_id := 0;	            
  end;    
  
  v_seqNum := v_seqNum + 1;
  select 病理取材信息_材块ID.Nextval into v_id from dual;
  
  
  --写入取材记录    
  insert into 病理取材信息(材块ID,序号, 病理号, 申请ID, 标本ID, 标本名称,形状,颜色,性质,标本量,蜡块数,主取医师,副取医师,记录医师,取材时间)
  values(v_id, v_seqNum, 病理号_IN, 申请ID_IN, 标本ID_IN, 标本名称_IN, 性质_IN,颜色_IN,性质_IN, 标本量_IN,细胞块数_IN,主取医师_IN,副取医师_IN,记录医师_IN,取材时间_IN);

  if 细胞块数_IN is null then
     v_slicesCount := 1;
  elsif 细胞块数_IN <= 0 then
     v_slicesCount := 1;
  else
     v_slicesCount := 细胞块数_IN;
  end if;  

  --写入制片记录
  insert into 病理制片信息(ID,病理号,材块ID,申请ID,制片类型,制片方式,制片数,当前状态)
  values(病理制片信息_ID.NEXTVAL,病理号_IN,v_id,申请ID_IN,2, 0, v_slicesCount, 0);
  
  commit; 
  
  return v_id || '-' || v_seqNum;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病理取材_细胞;
/


CREATE OR REPLACE Procedure Zl_病理取材_细胞更新
(
  材块ID_IN      病理取材信息.病理号%Type,   
  形状_IN        病理取材信息.形状%Type,
  颜色_IN        病理取材信息.颜色%Type, 
  性质_IN        病理取材信息.性质%Type, 
  标本量_IN      病理取材信息.标本量%Type,   
  细胞块数_IN    病理取材信息.蜡块数%Type,
  主取医师_IN    病理取材信息.主取医师%Type,
  副取医师_IN    病理取材信息.副取医师%Type,
  取材时间_IN    病理取材信息.取材时间%Type  
)Is
v_slicesCount number;
Begin
  --更新取材信息      
  update 病理取材信息
  set 形状=形状_IN,颜色=颜色_IN,性质=性质_IN,标本量=标本量_IN,蜡块数=细胞块数_IN,主取医师=主取医师_IN,副取医师=副取医师_IN
  where 材块ID=材块ID_IN;

  if 细胞块数_IN is null then
     v_slicesCount := 1;
  elsif 细胞块数_IN <= 0 then
     v_slicesCount := 1;
  else
     v_slicesCount := 细胞块数_IN;     
  end if;  
  
  
  --更新制片信息
  update 病理制片信息  set 制片数=v_slicesCount where 材块ID=材块ID_IN and 当前状态=0;
  
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病理取材_细胞更新;
/



--病理冰冻取材
CREATE OR REPLACE function Zl_病理取材_冰冻
(
  病理号_IN      病理取材信息.病理号%Type,   
  申请ID_IN      病理取材信息.申请ID%Type,
  标本ID_IN      病理取材信息.标本ID%Type,
  标本名称_IN    病理取材信息.标本名称%Type,
  取材位置_IN    病理取材信息.取材位置%Type,
  是否冰余_IN    病理取材信息.是否冰余%Type,
  蜡块数_IN      病理取材信息.蜡块数%Type,
  主取医师_IN    病理取材信息.主取医师%Type,
  副取医师_IN    病理取材信息.副取医师%Type,
  记录医师_IN    病理取材信息.记录医师%Type,
  取材时间_IN    病理取材信息.取材时间%Type  
) return varchar2 Is
PRAGMA AUTONOMOUS_TRANSACTION;

v_id 病理取材信息.材块ID%Type;
v_seqNum 病理取材信息.序号%Type;

Begin
  --获取最大材块号序号  
  begin
    select  nvl(max(序号), 0) into v_seqNum from 病理取材信息 where 病理号=病理号_IN;
  exception
    When Others Then v_id := 0;	            
  end;    
  
  v_seqNum := v_seqNum + 1;
  select 病理取材信息_材块ID.Nextval into v_id from dual;
  
  --写入取材记录    
  insert into 病理取材信息(材块ID, 序号, 病理号, 申请ID, 标本ID, 标本名称,取材位置,是否冰余,蜡块数,主取医师,副取医师,记录医师,取材时间)
  values(v_id, v_seqNum, 病理号_IN, 申请ID_IN, 标本ID_IN, 标本名称_IN, 取材位置_IN, 是否冰余_IN,蜡块数_IN,主取医师_IN,副取医师_IN,记录医师_IN,取材时间_IN);


  --写入制片记录
  insert into 病理制片信息(ID,病理号,材块ID,申请ID,制片类型,制片方式,制片数,当前状态)
  values(病理制片信息_ID.NEXTVAL,病理号_IN,v_id,申请ID_IN,1,0,蜡块数_IN,0);
  
  
  commit; 
  
  return v_id || '-' || v_seqNum;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病理取材_冰冻;
/



CREATE OR REPLACE Procedure Zl_病理取材_冰冻更新
(
  材块ID_IN      病理取材信息.病理号%Type,   
  取材位置_IN    病理取材信息.取材位置%Type,
  是否冰余_IN    病理取材信息.是否冰余%Type,
  蜡块数_IN      病理取材信息.蜡块数%Type,
  主取医师_IN    病理取材信息.主取医师%Type,
  副取医师_IN    病理取材信息.副取医师%Type 
)Is
Begin
  
  --更新取材记录      
  update 病理取材信息
  set 取材位置=取材位置_IN,是否冰余=是否冰余_IN,蜡块数=蜡块数_IN,主取医师=主取医师_IN,副取医师=副取医师_IN
  where 材块ID=材块ID_IN;


  --更新制片信息
  update 病理制片信息  set 制片数=蜡块数_IN  where 材块ID=材块ID_IN and 当前状态=0;
  
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病理取材_冰冻更新;
/



--删除病理取材记录
CREATE OR REPLACE Procedure Zl_病理取材_删除
(
  材块ID_IN      病理取材信息.材块ID%Type
)Is
Begin
      
  delete 病理取材信息 where 材块ID=材块ID_IN;

Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病理取材_删除;
/

  


--保存取材时的相关信息
CREATE OR REPLACE Procedure Zl_病理取材_信息保存
(
  病理号_IN      病理检查信息.病理号%Type,   
  巨检描述_IN    病理检查信息.巨检描述%Type,
  剩余位置_IN    病理检查信息.剩余位置%Type
)Is
Begin
        
  update 病理检查信息  set 巨检描述=巨检描述_IN,剩余位置=剩余位置_IN  where 病理号=病理号_IN;

Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病理取材_信息保存;
/


--取材确认
CREATE OR REPLACE Procedure Zl_病理取材_确认
(
  病理号_IN      病理检查信息.病理号%Type
)Is
Begin
        
  --更新病理检查状态
  update 病理检查信息 set 当前过程=2 where 病理号=病理号_IN;
  
  --如果有补取申请，则更新申请状态  
  update 病理申请信息 set 申请状态=1 where 申请状态=0 and 申请类型=8 and 病理号=病理号_IN;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病理取材_确认;
/




--病理制片====================================================================================================

--接受制片处理（接受该病理检查的所有制片）
CREATE OR REPLACE Procedure Zl_病理制片_接受
(
  病理号_IN      病理取材信息.病理号%Type,
  制片人_IN      病理制片信息.制片人%Type
)Is
Begin
        
  --更新制片状态(为处理的制片才能接受)
  update 病理制片信息 set 当前状态=1,制片人=制片人_IN  where 病理号 = 病理号_IN and 当前状态=0;
  
  --更新申请状态（0-已申请，1-已接受，2-已完成）
  update 病理申请信息 set 申请状态 = 1 
  where 申请ID=(select distinct 申请ID from 病理制片信息  where 病理号=病理号_IN and 当前状态=0) 
        and 申请状态=0;
  
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病理制片_接受;
/



--接受制片处理（接受当前材块的制片）
CREATE OR REPLACE Procedure Zl_病理制片_清单打印
(
  ID_IN      病理制片信息.材块ID%Type
)Is
Begin        
  
  --更新清单状态
  update 病理制片信息  set 清单状态=1 where ID=ID_IN;
  
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病理制片_清单打印;
/



--确认制片
CREATE OR REPLACE Procedure Zl_病理制片_确认
(
  病理号_IN      病理取材信息.病理号%Type,
  制片时间_IN    病理制片信息.制片时间%Type
)Is
Begin
        
  --更新制片状态（将未完成的制片记录，修改为已完成状态，未接受的制片不能进行确认）
  update 病理制片信息 set 当前状态=2,制片时间=制片时间_IN where 病理号=病理号_IN and 当前状态=1;  
  --where 材块id in(select 材块id from 病理取材信息 where 病理号=病理号_IN) and 当前状态<>2;

  
  --修改检查的当前过程为下一阶段（制片完成后的下一阶段为诊断）
  update 病理检查信息  set 当前过程=3 where 病理号=病理号_IN;  
  
  
  --更新申请状态（如果有申请则更新，没有则不执行 0-已申请，1-已接受，2-已完成）
  update 病理申请信息 set 申请状态 = 2, 完成时间=制片时间_IN 
  where 申请ID=(select distinct 申请ID from 病理制片信息 where 病理号=病理号_IN and 当前状态=1)  
        and 申请状态=1; 
  
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病理制片_确认;
/


--确认制片
/*CREATE OR REPLACE Procedure Zl_病理制片_确认1
(
  材块ID_IN      病理制片信息.材块ID%Type,
  制片时间_IN    病理制片信息.制片时间%Type
)Is
  v_count number;
Begin
        
  --更新制片状态（将未完成的制片记录，修改为已完成状态）
  update 病理制片信息 
  set 当前状态=2,制片时间=制片时间_IN 
  where 材块ID = 材块ID_IN and 当前状态<>2;
  --where 材块id in(select 材块id from 病理取材信息 where 病理号=病理号_IN) and 当前状态=1;
  
  v_count := 0;
  begin
    select sum(制片数) into v_count from 病理制片信息 a, 病理取材信息 b
    where a.材块id = b.材块id and b.病理号 = (select 病理号 from 病理取材信息 where 材块id=材块ID_IN) and a.当前状态 <> 2;
  exception
    when others then v_count := 0;          
  end;    
  
  --修改检查的当前过程和制片状态(当所有材块均被确认后，修改检查的执行状态) 
  if v_count = 0 then
    update 病理检查信息  set 当前过程=3, 制片状态=2  where 病理号=(select 病理号 from 病理取材信息 where 材块id=材块ID_IN);  
  end if;  
  
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病理制片_确认1;
*/



--报告延迟===================================================================================================


--添加报告延迟记录
CREATE OR REPLACE function Zl_病理报告延迟_新增
(
  病理号_IN      病理报告延迟.病理号%Type,   
  延迟原因_IN    病理报告延迟.延迟原因%Type,
  延迟天数_IN    病理报告延迟.延迟天数%Type,
  临时诊断_IN    病理报告延迟.临时诊断%Type,
  转达人_IN      病理报告延迟.转达人%Type,  
  登记人_IN      病理报告延迟.登记人%Type,
  登记时间_IN    病理报告延迟.登记时间%Type
) return number Is
PRAGMA AUTONOMOUS_TRANSACTION;
v_id 病理报告延迟.ID%Type;
Begin
  --获取延迟报告ID
  select 病理报告延迟_ID.NEXTVAL into v_id from dual;

  
  --写入报告延迟记录    
  insert into 病理报告延迟(ID, 病理号, 延迟原因,延迟天数,临时诊断,转达人,登记人,登记时间,当前状态)
  values(v_id, 病理号_IN, 延迟原因_IN, 延迟天数_IN, 临时诊断_IN, 转达人_IN,登记人_IN,登记时间_IN, 0);
  
  commit; 
  
  return v_id;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病理报告延迟_新增;
/


--更新报告延迟记录
CREATE OR REPLACE procedure Zl_病理报告延迟_更新
(
  ID_IN          病理报告延迟.ID%Type,   
  延迟原因_IN    病理报告延迟.延迟原因%Type,
  延迟天数_IN    病理报告延迟.延迟天数%Type,
  临时诊断_IN    病理报告延迟.临时诊断%Type,
  转达人_IN      病理报告延迟.转达人%Type
) Is
Begin
  
  update 病理报告延迟
  set 延迟原因=延迟原因_IN,延迟天数=延迟天数_IN,临时诊断=临时诊断_IN,转达人=转达人_IN
  where ID=ID_IN;

Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病理报告延迟_更新;
/


--删除报告延迟
CREATE OR REPLACE procedure Zl_病理报告延迟_删除
(
  ID_IN          病理报告延迟.ID%Type
) Is
Begin

  --删除未打印的延迟报告，如果已打印则不能删除
  delete 病理报告延迟 where ID=ID_IN; -- and 当前状态=0;  

Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病理报告延迟_删除;
/

--打印报告延迟
CREATE OR REPLACE procedure Zl_病理报告延迟_打印
(
  ID_IN          病理报告延迟.ID%Type
) Is
Begin

  --当打印后，修改报告延迟记录的当前状态
  update 病理报告延迟 set 当前状态=1 where ID=ID_IN;  

Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病理报告延迟_打印;
/


--过程报告===================================================================================================


--添加过程报告记录
CREATE OR REPLACE function Zl_病理过程报告_新增
(
  病理号_IN      病理过程报告.病理号%Type,   
  标本名称_IN    病理过程报告.标本名称%Type,
  报告类型_IN    病理过程报告.报告类型%Type,
  检查意见_IN    病理过程报告.检查意见%Type,
  检查结果_IN    病理过程报告.检查结果%Type,  
  报告医师_IN    病理过程报告.报告医师%Type,
  报告日期_IN    病理过程报告.报告日期%Type,
  报告图像_IN    病理过程报告.报告图像%Type,
  备注_IN        病理过程报告.备注%Type
) return number Is
PRAGMA AUTONOMOUS_TRANSACTION;
v_id 病理过程报告.ID%Type;
Begin
  --获取过程报告ID
  select 病理过程报告_ID.NEXTVAL into v_id from dual;

  
  --写入过程报告记录    
  insert into 病理过程报告(ID, 病理号, 标本名称,报告类型,检查结果,检查意见,报告图像,报告医师,报告日期,当前状态,备注)
  values(v_id, 病理号_IN, 标本名称_IN, 报告类型_IN, 检查结果_IN, 检查意见_IN,报告图像_IN,报告医师_IN,报告日期_IN, 0,备注_IN);
  
  commit; 
  
  return v_id;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病理过程报告_新增;
/ 

--更新过程报告
CREATE OR REPLACE procedure Zl_病理过程报告_更新
(
  ID_IN          病理过程报告.ID%Type,   
  标本名称_IN    病理过程报告.标本名称%Type,
  报告类型_IN    病理过程报告.报告类型%Type,
  检查意见_IN    病理过程报告.检查意见%Type,
  检查结果_IN    病理过程报告.检查结果%Type,  
  报告图像_IN    病理过程报告.报告图像%Type,
  备注_IN        病理过程报告.备注%Type
)Is
Begin

  --更新过程报告记录    
  update 病理过程报告
  set 标本名称=标本名称_IN, 报告类型=报告类型_IN,检查意见=检查意见_IN,
      检查结果=检查结果_IN, 报告图像=报告图像_IN,备注=备注_IN
  where ID=ID_IN;
  
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病理过程报告_更新;
/


--删除过程报告
CREATE OR REPLACE procedure Zl_病理过程报告_删除
(
  ID_IN          病理过程报告.ID%Type
)Is
Begin

  --删除过程报告记录    
  delete 病理过程报告 where ID=ID_IN;
  
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病理过程报告_删除;
/


--过程报告状态设置
CREATE OR REPLACE procedure Zl_病理过程报告_状态
(
  ID_IN          病理过程报告.ID%Type,
  当前状态_IN    病理过程报告.当前状态%Type  
)Is
Begin

  --删除过程报告记录    
  update 病理过程报告 set 当前状态=当前状态_IN where ID=ID_IN;
  
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病理过程报告_状态;
/

 


--检查申请===================================================================================================



--添加检查申请
CREATE OR REPLACE function Zl_病理申请_新增
(
  病理号_IN      病理申请信息.病理号%Type,   
  申请人_IN      病理申请信息.申请人%Type,
  申请时间_IN    病理申请信息.申请时间%Type,
  申请类型_IN    病理申请信息.申请类型%Type,
  申请描述_IN    病理申请信息.申请描述%Type
) return number Is
PRAGMA AUTONOMOUS_TRANSACTION;
v_id 病理申请信息.申请ID%Type;
v_procedure 病理检查信息.当前过程%Type;
Begin
  --获取申请ID
  select 病理申请信息_申请ID.NEXTVAL into v_id from dual;

  
  --写入申请记录    
  insert into 病理申请信息(申请ID, 病理号, 申请人,申请时间,申请类型,申请描述,申请状态,是否打印)
  values(v_id, 病理号_IN, 申请人_IN, 申请时间_IN, 申请类型_IN, 申请描述_IN,0,0);
  
  --更新检查过程
  case 
    when 申请类型_IN = 0 then v_procedure := 4;  --免疫组化
    when 申请类型_IN = 1 then v_procedure := 5;  --特殊染色
    when 申请类型_IN = 2 then v_procedure := 6;  --分子病理
    when 申请类型_IN = 3 then v_procedure := 9;  --再制片
    when 申请类型_IN = 4 then v_procedure := 8;  --再取材
    else v_procedure := -1;
  end case;
  
  if v_procedure <= 0 then 
    Raise_Application_Error(-20101, '[ZLSOFT] 检查申请时，不能有效取得病理检查信息中的当前过程，终止执行。[ZLSOFT]');
  end if;
  
  
  update 病理检查信息 set 当前过程=v_procedure where 病理号=病理号_IN;
    
  
  commit; 
  
  return v_id;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病理申请_新增;
/ 



--删除检查申请
CREATE OR REPLACE procedure Zl_病理申请_删除
(
  申请ID_IN          病理申请信息.申请ID%Type
)Is
Begin

  --删除过程报告记录    
  delete 病理申请信息 where 申请ID=申请ID_IN;
  
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病理申请_删除;
/


--添加特检项目
CREATE OR REPLACE function Zl_病理申请_特检项目_新增
(
  病理号_IN      病理特检信息.病理号%Type,         
  材块ID_IN      病理特检信息.材块ID%Type,   
  申请ID_IN      病理特检信息.申请ID%Type,
  抗体ID_IN      病理特检信息.抗体ID%Type,
  特检类型_IN    病理特检信息.特检类型%Type,
  是否补做_IN    number := 0
) return number Is
PRAGMA AUTONOMOUS_TRANSACTION;
v_id 病理特检信息.ID%Type;
v_补做 病理特检信息.制作类型%Type;
Begin
  --获取特检信息ID
  select 病理特检信息_ID.NEXTVAL into v_id from dual;

  v_补做 := -1;
  if 是否补做_IN = 0 then
     v_补做 := 0;
  end if;
  
  --写入申请记录    
  insert into 病理特检信息(ID,病理号,材块ID,申请ID,抗体ID,特检类型,制作类型,当前状态,清单状态)
  values(v_id,病理号_IN, 材块ID_IN, 申请ID_IN, 抗体ID_IN, 特检类型_IN, v_补做, 0,0);
  
  --更新检查过程
  if v_补做 = -1 then
     update 病理检查信息 set 当前过程=decode(特检类型_IN, 0, 4, 1, 5, 6) where 病理号=病理号_IN;
  end if;
  
  commit; 
  
  return v_id;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病理申请_特检项目_新增;
/ 


--删除特检项目
CREATE OR REPLACE procedure Zl_病理申请_特检项目_删除
(
  ID_IN          病理特检信息.ID%Type
)Is
Begin

  --删除特检项目(只有已申请的项目才允许删除)
  delete 病理特检信息 where ID=ID_IN and 当前状态=0;
  
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病理申请_特检项目_删除;
/


--特检项目重做
CREATE OR REPLACE function Zl_病理申请_特检项目_重做
(
  ID_IN          病理特检信息.ID%Type
) return number Is
PRAGMA AUTONOMOUS_TRANSACTION;
  cursor c_SpeExamInf is
         select 病理号,材块ID,抗体ID,特检类型 from 病理特检信息 where Id=ID_IN;
 
  r_SpeExamInf  c_SpeExamInf%RowType;         
  v_newid      病理特检信息.ID%Type;
  v_count   病理特检信息.制作类型%Type;
  
  v_Error    varchar(255);
  Err_Custom Exception;
Begin
    
  
  Open c_SpeExamInf;
  Fetch c_SpeExamInf Into r_SpeExamInf;
    
  If c_SpeExamInf%Rowcount = 0 Then
    Close c_SpeExamInf;
    v_Error := '不能正确读取病理特检信息的相关数据，请检查项目ID是否为有效数据。';
    Raise Err_Custom;
  End If;  
  
  v_count := 0;
  begin
    select nvl(max(制作类型), 0) into v_count from 病理特检信息 
    where 材块Id=r_SpeExamInf.材块ID and 抗体Id=r_SpeExamInf.抗体ID and 特检类型=r_SpeExamInf.特检类型; 
  exception
    when others then v_count := 0;           
  end;
  
  select 病理特检信息_ID.NEXTVAL into v_newid from dual;

  --特检项目重做
  insert into 病理特检信息(ID, 病理号,材块ID,申请ID,抗体ID,特检类型,制作类型,当前状态,清单状态) 
  select v_newid as ID,病理号,材块ID,申请ID,抗体ID,特检类型, v_count+1,0,0 from 病理特检信息 where ID=ID_IN;
  
  
  --更新检查过程
  update 病理检查信息 set 当前过程=decode(r_SpeExamInf.特检类型, 0, 4, 1, 5, 6) where 病理号=r_SpeExamInf.病理号;

  commit;
  
  return v_newid;
  
  close c_SpeExamInf;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病理申请_特检项目_重做;
/




--新增制片项目
CREATE OR REPLACE function Zl_病理申请_制片项目_新增
(
  病理号_IN      病理制片信息.病理号%Type,    
  材块ID_IN      病理制片信息.材块ID%Type,   
  申请ID_IN      病理制片信息.申请ID%Type,
  制片类型_IN    病理制片信息.制片类型%Type,
  制片方式_IN    病理制片信息.制片方式%Type,
  制片数量_IN    病理制片信息.制片数%Type
) return number Is
PRAGMA AUTONOMOUS_TRANSACTION;
v_id       病理制片信息.ID%Type;
Begin
  --获取特检信息ID
  select 病理制片信息_ID.NEXTVAL into v_id from dual;

  
  --写入制片记录    
  insert into 病理制片信息(ID, 病理号, 材块ID,申请ID,制片类型,制片数,制片方式,当前状态,清单状态)
  values(v_id, 病理号_IN, 材块ID_IN, 申请ID_IN, 制片类型_IN, 制片数量_IN, 制片方式_IN, 0,0);
  
  commit; 
  
  return v_id;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病理申请_制片项目_新增;
/ 


--删除制片项目
CREATE OR REPLACE procedure Zl_病理申请_制片项目_删除
(
  ID_IN          病理制片信息.ID%Type
)Is
Begin

  --删除制片项目(只有未处理的项目才允许删除)
  delete 病理制片信息 where ID=ID_IN and 当前状态=0;
  
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病理申请_制片项目_删除;
/




--病理会诊===================================================================================================

--添加病理会诊
CREATE OR REPLACE function Zl_病理会诊_新增
(
  病理号_IN      病理会诊信息.病理号%Type,   
  申请医师_IN    病理会诊信息.申请医师%Type,
  会诊单位_IN    病理会诊信息.会诊单位%Type,
  会诊医师_IN    病理会诊信息.会诊医师%Type,
  会诊时间_IN    病理会诊信息.会诊时间%Type,
  截止时间_IN    病理会诊信息.截止时间%Type,
  会诊类型_IN    病理会诊信息.会诊类型%Type,
  检查描述_IN    病理会诊信息.检查描述%Type
) return number Is
PRAGMA AUTONOMOUS_TRANSACTION;
v_id 病理会诊信息.ID%Type;
Begin
  --获取申请ID
  select 病理会诊信息_ID.NEXTVAL into v_id from dual;

  
  --写入申请记录    
  insert into 病理会诊信息(id,病理号,申请医师,会诊单位,会诊医师,会诊时间,截止时间,会诊类型,检查描述,当前状态)
  values(v_id, 病理号_IN, 申请医师_IN, 会诊单位_IN, 会诊医师_IN, 会诊时间_IN,截止时间_IN,会诊类型_IN,检查描述_IN,0);
  
  commit; 
  
  return v_id;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病理会诊_新增;
/ 


--会诊反馈
CREATE OR REPLACE procedure Zl_病理会诊_反馈
(
  ID_IN          病理会诊信息.ID%Type,
  诊断结果_IN    病理会诊信息.诊断结果%Type,
  诊断意见_IN    病理会诊信息.诊断意见%Type,
  完成时间_IN    病理会诊信息.完成时间%Type,
  备注_IN        病理会诊信息.备注%Type
)Is
Begin

  --更新病理会诊记录 （当会诊记录反馈后，则将状态修改为完成状态）
  update 病理会诊信息 
  set 诊断结果 = 诊断结果_IN,诊断意见=诊断意见_IN,完成时间=完成时间_IN,备注=备注_IN,当前状态=2
  where ID=ID_IN;
  
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病理会诊_反馈;
/



--删除已申请的病理会诊信息
CREATE OR REPLACE procedure Zl_病理会诊_删除
(
  ID_IN          病理会诊信息.ID%Type
)Is
Begin

  --删除病理会诊记录 (已申请或者已撤销的会诊记录可被删除)
  delete 病理会诊信息 where ID=ID_IN and (当前状态=0 or 当前状态=1);
  
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病理会诊_删除;
/



--设置会诊当前状态
CREATE OR REPLACE procedure Zl_病理会诊_状态
(
  ID_IN          病理会诊信息.ID%Type,
  当前状态_IN    病理会诊信息.当前状态%Type  
)Is
Begin

  --设置会诊记录的当前状态
  update 病理会诊信息 set 当前状态=当前状态_IN where ID=ID_IN;
  
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病理会诊_状态;
/



--病理特检===================================================================================================


--接受特检处理
CREATE OR REPLACE Procedure Zl_病理特检_接受
(
  病理号_IN      病理取材信息.病理号%Type,
  特检类型_IN    病理特检信息.特检类型%Type, 
  特检医师_IN    病理特检信息.特检医师%Type
)Is
Begin
  --更新申请状态（0-已申请，1-已接受，2-已完成）
  update 病理申请信息 set 申请状态 = 1 
  where 申请ID=(select distinct 申请ID from 病理特检信息 where 病理号=病理号_IN and 特检类型=特检类型_IN and 当前状态=0)
        and 申请状态=0;
          
        
  --更新特检状态（只有当前状态为0的特检信息才进行更新）
  update 病理特检信息 
  set 当前状态=1,特检医师=特检医师_IN 
  where 材块id in(select 材块id from 病理取材信息 where 病理号=病理号_IN) and 特检类型=特检类型_IN and 当前状态=0;
  
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病理特检_接受;
/


--接受特检处理
CREATE OR REPLACE Procedure Zl_病理特检_清单打印
(
  ID_IN          病理特检信息.ID%Type
)Is 
Begin
        
  --更新特检状态（只有当前状态为0的特检信息才进行更新）
  --update 病理特检信息 set 当前状态=1,特检医师=特检医师_IN where id=ID_IN and 当前状态=0;
  
    --更新清单状态
  update 病理特检信息  set 清单状态=1 where ID=ID_IN;
  
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病理特检_清单打印;
/


--特检项目录入
CREATE OR REPLACE Procedure Zl_病理特检_项目录入
(
  ID_IN          病理特检信息.ID%Type,
  项目结果_IN    病理特检信息.项目结果%Type 
)Is
Begin
        
  --更新特检的项目结果（必须已接受的项目才能进行录入）
  update 病理特检信息 set 项目结果=项目结果_IN where ID=ID_IN; --and 当前状态=1;
  
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病理特检_项目录入;
/


--特检确认
CREATE OR REPLACE Procedure Zl_病理特检_确认
(
  病理号_IN      病理检查信息.病理号%Type,
  特检类型_IN    病理特检信息.特检类型%Type,
  完成时间_IN    病理特检信息.完成时间%Type  
)Is
  v_count number;
Begin  
  v_count := 0;
  begin
    select count(id) into v_count from 病理特检信息 where 病理号=病理号_IN and 特检类型=特检类型_IN and 当前状态<>2;
  exception
    when others then v_count := 0;     
  end;
  
  if v_count <= 0 then
    --更新检查过程
    update 病理检查信息 set 当前过程=3 where 病理号=病理号_IN;
  end if;
  
  --更新申请状态（0-已申请，1-已接受，2-已完成）
  update 病理申请信息 set 申请状态 = 2,完成时间=完成时间_IN 
  where 申请ID=(select distinct 申请ID from 病理特检信息 where 病理号=病理号_IN and 特检类型=特检类型_IN and 当前状态=1) 
        and 申请状态=1;  
        
  --更新特检状态（必须已接受的项目才能进行确认）
  update 病理特检信息 set 当前状态=2, 完成时间=完成时间_IN where 病理号=病理号_IN and 特检类型=特检类型_IN and 当前状态=1;        
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病理特检_确认;
/



--病理完成===================================================================================================

--检查完成
CREATE OR REPLACE Procedure Zl_病理检查_完成
(
  医嘱ID_IN    病理检查信息.医嘱ID%Type     
)Is
Begin
  update 病理检查信息 set 当前过程=10 where 医嘱id=医嘱ID_IN;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);  
End Zl_病理检查_完成;
/


--取消完成
CREATE OR REPLACE Procedure Zl_病理检查_取消完成
(
  医嘱ID_IN    病理检查信息.医嘱ID%Type     
)Is
Begin
  update 病理检查信息 set 当前过程=3 where 医嘱id=医嘱ID_IN;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);  
End Zl_病理检查_取消完成;
/




