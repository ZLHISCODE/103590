CREATE OR REPLACE Procedure Zl_历史数据_导入
(
  strDecode_IN  varchar2
)Is
  v_TempDecode varchar2(1000);
Begin

insert into 病理标本信息(标本ID, 医嘱ID, 标本名称,标本类型,数量,接收日期)
select 病理标本信息_标本ID.Nextval,a.医嘱ID,a.标本部位,0,a.块数,b.核收时间
from 影像病理标本 a, 影像标本核收取材 b where a.医嘱id=b.医嘱id 
and not exists(Select 1 From 病理标本信息 where 医嘱ID=a.医嘱ID and 标本名称=a.标本部位 and 数量=a.块数 and 接收日期=b.核收时间);

--连接传入Decode参数
v_TempDecode:= 'insert into 病理检查信息(病理医嘱ID,病理号,医嘱ID,检查类型,巨检描述,剩余位置)
               select 病理检查信息_病理医嘱ID.Nextval,病理号,医嘱ID,' || strDecode_IN || ',巨检所见,剩余标本位置
               from 影像标本核收取材 where 医嘱ID not in(select 医嘱ID from 病理检查信息)';

Execute Immediate v_TempDecode;

insert into 病理送检信息(ID,医嘱ID,送检单位,送检科室,送检人,送检日期,登记人,核收状态,拒收原因,备注)
select 病理送检信息_id.nextval,医嘱ID,'本院','', '未录入',核收时间,decode(核收技师,null,'未录入',核收技师),decode(核收情况,'1','1','0'),拒收原因,备注
from 影像标本核收取材 a where not exists(Select 1 From 病理送检信息 where 医嘱ID=a.医嘱id and 送检日期=a.核收时间 and 拒收原因=a.拒收原因);

update 病理标本信息 a set 送检ID=(select  ID from 病理送检信息 where 医嘱ID=a.医嘱ID and rownum=1);


Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_历史数据_导入;
