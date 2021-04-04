
insert into 病理标本信息(标本ID, 医嘱ID, 标本名称,标本类型,数量,接收日期)
select 病理标本信息_标本ID.Nextval,a.医嘱ID,a.标本部位,0,a.块数,b.核收时间 
from 影像病理标本 a, 影像标本核收取材 b where a.医嘱id=b.医嘱id;


insert into 病理检查信息(病理号,医嘱ID,检查类型,当前过程,巨检描述,剩余位置)
select 病理号,医嘱ID,decode(病理检查类别, '测试类别','0','1'),3,巨检所见,剩余标本位置
from 影像标本核收取材 where 医嘱ID not in(select 医嘱ID from 病理检查信息);


insert into 病理送检信息(ID,医嘱ID,送检单位,送检科室,送检人,送检日期,登记人,核收状态,拒收原因,备注)
select 病理送检信息_id.nextval,医嘱ID,'本院','', '未录入',核收时间,核收技师,decode(核收情况,'1','1','0'),拒收原因,备注
from 影像标本核收取材;












