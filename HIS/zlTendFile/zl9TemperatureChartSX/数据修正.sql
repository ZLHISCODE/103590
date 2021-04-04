--主要用于修正体温单编辑大便次数或入液量选择特殊符号保存后，病人护理明细项目序号=0的数据。
DECLARE
  --提取再院病人的体温单文件
  CURSOR Cur_File IS
    SELECT d.Rowid, d.项目序号, d.记录内容
    FROM 病人信息 e, 病案主页 f, 病人护理文件 b, 病历文件列表 a, 病人护理数据 c, 病人护理明细 d
    WHERE e.在院 = 1 AND e.病人id = f.病人id AND e.住院次数 = f.主页id AND e.病人id = b.病人id AND e.住院次数 = b.主页id AND f.出院日期 IS NULL AND
          a.种类 = 3 AND a.保留 = -1 AND b.格式id = a.Id AND b.Id = c.文件id AND c.Id = d.记录id AND d.项目类型 = 1 AND d.项目序号 = 0;
  TYPE t_File IS TABLE OF Cur_File%ROWTYPE;
  t_Count t_File;
BEGIN

  OPEN Cur_File;
  LOOP
    FETCH Cur_File BULK COLLECT
      INTO t_Count LIMIT 100;
    EXIT WHEN t_Count.Count = 0;
    FOR i IN 1 .. t_Count.Count LOOP
      IF t_Count(i).记录内容 LIKE '%C' THEN
        UPDATE 病人护理明细 SET 记录类型=1,项目序号 = 9, 项目名称 = '入液量' WHERE ROWID = t_Count(i).Rowid;
      END IF;
      IF t_Count(i).记录内容 = '※' OR t_Count(i).记录内容 = '☆' OR t_Count(i).记录内容 LIKE '%E' THEN
        UPDATE 病人护理明细 SET 记录类型=1,项目序号 = 10, 项目名称 = '大便次数' WHERE ROWID = t_Count(i).Rowid;
      END IF;
    END LOOP;
  END LOOP;
  CLOSE Cur_File;
END;
/ 
COMMIT;
