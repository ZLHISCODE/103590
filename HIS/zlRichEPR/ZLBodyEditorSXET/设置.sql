Update 体温记录项目 Set 最大值=42,最小值=33,单位值=0.1,最高行=2 Where 项目序号=1;
Update 体温记录项目 Set 最大值=200,最小值=40,单位值=2,最高行=7 Where 项目序号=2;
Update 体温记录项目 Set 最大值=200,最小值=40,单位值=2,最高行=7 Where 项目序号=-1;
Update 护理记录项目 Set 项目值域='33;42;' Where 项目序号=1;
Update 护理记录项目 Set 项目值域='40;200;' Where 项目序号=2;
Update 护理记录项目 Set 项目值域='40;200;' Where 项目序号=-1;


UPDATE 体温部件 SET 启用=0;
INSERT INTO 体温部件 (名称,适用地区,部件,启用)
VALUES ('陕西西安儿童医院','目前使用儿童医院','ZL9BODYEDITORSXET',1);

commit;
