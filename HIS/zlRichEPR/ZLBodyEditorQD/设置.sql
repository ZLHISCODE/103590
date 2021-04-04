
update 体温部件 set 启用=0 ;
insert into 体温部件(名称,适用地区,部件,启用) values ('青岛体温部件','青岛专用','zl9BODYEDITORQD',1);

update 护理记录项目 set 项目值域='20;180;' where 项目序号 IN (-1,2);
update 护理记录项目 set 项目值域='35,42;' where 项目序号=1;
update 体温记录项目 set 最大值=180,最小值=20,单位值=4,最高行=2 where 项目序号 IN (-1,2);
update 体温记录项目 set 最大值=42,最小值=35,单位值=0.2,最高行=2 where 项目序号 =1;


