Update 体温记录项目 Set 最大值=42,最小值=35,单位值=0.1,最高行=2 Where 项目序号=1;
Update 体温记录项目 Set 最大值=200,最小值=40,单位值=2,最高行=2 Where 项目序号=2;
Update 体温记录项目 Set 最大值=200,最小值=40,单位值=2,最高行=2 Where 项目序号=-1;
Update 护理记录项目 Set 项目值域='35;42;' Where 项目序号=1;
Update 护理记录项目 Set 项目值域='40;200;' Where 项目序号=2;
Update 护理记录项目 Set 项目值域='40;200;' Where 项目序号=-1;

UPDATE 体温部件 SET 启用=0;

DECLARE 
  nCount number;
BEGIN
  BEGIN 
    SELECT 1 INTO nCount From 体温部件 WHERE upper(部件)='ZL9BODYEDITORSCDQ';
    EXCEPTION 
      WHEN OTHERS THEN nCount:=0;
  END;
  IF nCount=0 THEN 
    INSERT INTO 体温部件 (名称,适用地区,部件,启用)
      VALUES ('四川省通用体温部件','四川省地区所有用户','ZL9BODYEDITORSCDQ',1);
  ELSE 
    UPDATE 体温部件 SET 启用=1 WHERE upper(部件)='ZL9BODYEDITORSCDQ';
  END if;
end;
/
commit;
