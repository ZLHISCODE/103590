Update 体温记录项目 Set 最大值=42,最小值=35,单位值=0.1,最高行=10 Where 项目序号=1;
Update 体温记录项目 Set 最大值=180,最小值=40,单位值=2,最高行=45 Where 项目序号=2;
Update 体温记录项目 Set 最大值=180,最小值=40,单位值=2,最高行=45 Where 项目序号=-1;
Update 护理记录项目 Set 项目值域='35;42;' Where 项目序号=1;
Update 护理记录项目 Set 项目值域='40;180;' Where 项目序号=2;
Update 护理记录项目 Set 项目值域='40;180;' Where 项目序号=-1;


UPDATE 体温部件 SET 启用=0;
INSERT INTO 体温部件 (名称,适用地区,部件,启用)
VALUES ('湖南省通用部件','适用湖南省所有地区','zl9BodyEditorHUN',1);

commit;
