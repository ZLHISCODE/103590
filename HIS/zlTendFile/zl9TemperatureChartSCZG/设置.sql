DECLARE 
BEGIN
  UPDATE 体温部件 SET 启用=0;
  BEGIN 
    UPDATE 体温部件 SET 启用=1 WHERE Upper(新部件)=Upper('zl9TemperatureChartSCZG');
    IF SQL%notfound then
      INSERT INTO 体温部件 (名称,适用地区,部件,启用,新部件)
      VALUES ('四川自贡市地区通用体温部件','适用于四川自贡市','zl9BodyEditor',1,'zl9TemperatureChartSCZG');
    END IF;
  END;
END; 
/

--修改对应参数值
UPDATE zlparameters SET 参数值='1' WHERE 系统=100 AND 模块=1255 AND 参数名='脉搏短绌以(心率/脉搏)方式录入';
UPDATE zlparameters SET 参数值='0' WHERE 系统=100 AND 模块=1255 AND 参数名='手术当天缺省格式';
COMMIT;


