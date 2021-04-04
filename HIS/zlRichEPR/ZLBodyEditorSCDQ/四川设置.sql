Update 体温记录项目 Set 最大值=42,最小值=35,单位值=0.1,最高行=2 Where 项目序号=1;
Update 体温记录项目 Set 最大值=180,最小值=40,单位值=2,最高行=2 Where 项目序号=2;
Update 体温记录项目 Set 最大值=180,最小值=40,单位值=2,最高行=2 Where 项目序号=-1;
Update 护理记录项目 Set 项目值域='35;42;' Where 项目序号=1;
Update 护理记录项目 Set 项目值域='40;180;' Where 项目序号=2;
Update 护理记录项目 Set 项目值域='40;180;' Where 项目序号=-1;
commit;