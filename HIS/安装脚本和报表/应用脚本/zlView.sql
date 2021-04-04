Create Or Replace View 在院病人自动记帐 as
Select p.类型,p.病人id, p.主页id, Nvl(a.姓名, i.姓名) As 姓名, Nvl(a.性别, i.性别) As 性别, Nvl(a.年龄, i.年龄) As 年龄, Nvl(a.住院号, i.住院号) As 住院号,
       a.费别, p.科室id, p.病区id, p.床号, p.附加床位, p.收费细目id, p.收入项目id, 1 As 标志, p.现价 As 标准单价, p.开始日期, p.终止日期,
       p.终止日期 - p.开始日期 As 天数, p.数量, p.经治医师, p.责任护士, p.操作员编号, p.操作员姓名,p.医疗小组id
From 病人信息 I, 病案主页 A,
     (Select 2 As 类型, b.病人id, b.主页id, b.科室id, b.病区id, b.床号, b.附加床位, p.收费细目id, p.收入项目id, p.现价, b.经治医师, b.责任护士, b.操作员编号, b.操作员姓名,
              Zl_Date_Half(Greatest(Least(Nvl(b.上次计算时间, b.开始时间), Nvl(b.终止时间, Greatest(Nvl(b.上次计算时间, b.开始时间))),
                                           Greatest(Nvl(b.上次计算时间, b.开始时间))), p.执行日期, Nvl(a.启用日期, Add_Months(Sysdate, -2)))) As 开始日期,
              Zl_Date_Half(Least(Nvl(b.终止时间, Greatest(b.开始时间, Sysdate)), Nvl(p.终止日期, Sysdate + 30) + 1)) As 终止日期, b.数量,
              b.医疗小组id
       From 自动计价项目 A,
            (Select a.病人id, a.主页id, a.开始时间, a.附加床位, a.病区id, a.科室id, a.床号, a.床位等级id, 1 As 数量, a.责任护士, a.经治医师, a.终止时间,
                     a.操作员编号, a.操作员姓名, a.上次计算时间, a.医疗小组id
              From 病人变动记录 A, 病人信息 B
              Where a.开始原因 <> 10 And a.病人id = b.病人id And a.主页id = b.主页id And b.在院 = 1
              Union All
              Select b.病人id, b.主页id, 开始时间, 附加床位, b.病区id, b.科室id, 床号, i.从项id As 床位等级id, i.从项数次 As 数量, 责任护士, 经治医师, 终止时间, 操作员编号,
                     操作员姓名, 上次计算时间, b.医疗小组id
              From 病人变动记录 B, 收费从属项目 I, 病人信息 C
              Where b.病人id = c.病人id And b.主页id = c.主页id And c.在院 = 1 And b.床位等级id = i.主项id And b.开始原因 <> 10 And i.固有从属 > 0) B,
            收费价目 P
       Where a.病区id = b.病区id And Zl_Date_Half(Nvl(b.上次计算时间, b.开始时间)) <> Zl_Date_Half(Nvl(b.终止时间, Sysdate)) And p.现价 <> 0 And
             a.计算标志 = 1 And b.床位等级id = p.收费细目id And Zl_Date_Half(Nvl(b.终止时间, Sysdate)) >= Zl_Date_Half(p.执行日期) And
             Zl_Date_Half(b.开始时间) <= Zl_Date_Half(Nvl(p.终止日期, Sysdate) + 1) And
             Zl_Date_Half(Least(Nvl(b.终止时间, Sysdate), Nvl(p.终止日期, Sysdate + 30) + 1)) >=
             Zl_Date_Half(Nvl(a.启用日期, Add_Months(Sysdate, -2))) And p.价格等级 Is Null
       Union All
       Select 1 As 类型, b.病人id, b.主页id, b.科室id, b.病区id, b.床号, b.附加床位, p.收费细目id, p.收入项目id, p.现价, b.经治医师, b.责任护士, b.操作员编号, b.操作员姓名,
              Zl_Date_Half(Greatest(Least(Nvl(b.上次计算时间, b.开始时间), Nvl(b.终止时间, Greatest(Nvl(b.上次计算时间, b.开始时间))),
                                           Greatest(Nvl(b.上次计算时间, b.开始时间))), p.执行日期, Nvl(a.启用日期, Add_Months(Sysdate, -2)))) As 开始日期,
              Zl_Date_Half(Least(Nvl(b.终止时间, Greatest(b.开始时间, Sysdate)), Nvl(p.终止日期, Sysdate + 30) + 1)) As 终止日期, b.数量,
              b.医疗小组id
       From 自动计价项目 A,
            (Select a.病人id, a.主页id, 开始时间, 附加床位, a.病区id, a.科室id, 床号, 护理等级id, 1 As 数量, 责任护士, 经治医师, 终止时间, 操作员编号, 操作员姓名, 上次计算时间,
                     a.医疗小组id
              From 病人变动记录 A, 病人信息 B
              Where 开始原因 <> 10 And a.病人id = b.病人id And a.主页id = b.主页id And b.在院 = 1
              Union All
              Select b.病人id, b.主页id, 开始时间, 附加床位, b.病区id, b.科室id, 床号, i.从项id As 护理等级id, i.从项数次 As 数量, 责任护士, 经治医师, 终止时间, 操作员编号,
                     操作员姓名, 上次计算时间, b.医疗小组id
              From 病人变动记录 B, 收费从属项目 I, 病人信息 C
              Where b.护理等级id = i.主项id And b.病人id = c.病人id And b.主页id = c.主页id And c.在院 = 1 And b.开始原因 <> 10 And i.固有从属 > 0) B,
            收费价目 P, 收费项目目录 C
       Where a.病区id = b.病区id And b.附加床位 <> 1 And
             Zl_Date_Half(Nvl(b.上次计算时间, b.开始时间)) <> Zl_Date_Half(Nvl(b.终止时间, Sysdate)) And p.现价 <> 0 And a.计算标志 = 2 And
             b.护理等级id = p.收费细目id And b.护理等级id = c.Id And Nvl(c.计算方式, 0) <> 1 And
             Zl_Date_Half(Nvl(b.终止时间, Sysdate)) >= Zl_Date_Half(p.执行日期) And
             Zl_Date_Half(b.开始时间) <= Zl_Date_Half(Nvl(p.终止日期, Sysdate) + 1) And
             Zl_Date_Half(Least(Nvl(b.终止时间, Sysdate), Nvl(p.终止日期, Sysdate + 30) + 1)) >=
             Zl_Date_Half(Nvl(a.启用日期, Add_Months(Sysdate, -2))) And p.价格等级 Is Null
       Union All
       Select 3 As 类型, b.病人id, b.主页id, b.科室id, b.病区id, b.床号, b.附加床位, p.收费细目id, p.收入项目id, p.现价, b.经治医师, b.责任护士, b.操作员编号, b.操作员姓名,
              Zl_Date_Half(Greatest(Least(Nvl(b.上次计算时间, b.开始时间), Nvl(b.终止时间, Greatest(Nvl(b.上次计算时间, b.开始时间))),
                                           Greatest(Nvl(b.上次计算时间, b.开始时间))), p.执行日期, Nvl(a.启用日期, Add_Months(Sysdate, -2)))) As 开始日期,
              Zl_Date_Half(Least(Nvl(b.终止时间, Greatest(b.开始时间, Sysdate)), Nvl(p.终止日期, Sysdate + 30) + 1)) As 终止日期, a.数量,
               b.医疗小组id
       From (Select 病区id, 计算标志, 收费细目id, 1 As 数量, 启用日期
              From 自动计价项目
              Union All
              Select 病区id, 计算标志, 从项id, i.从项数次 As 数量, 启用日期
              From 自动计价项目 A, 收费从属项目 I
              Where a.收费细目id = i.主项id And i.固有从属 > 0) A, 病人变动记录 B, 收费价目 P, 病人信息 C
       Where a.病区id = b.病区id And b.病人id = c.病人id And b.主页id = c.主页id And c.在院 = 1 And b.附加床位 <> 1 And b.开始原因 <> 10 And
             Zl_Date_Half(Nvl(b.上次计算时间, b.开始时间)) <> Zl_Date_Half(Nvl(b.终止时间, Sysdate)) And p.现价 <> 0 And
             a.收费细目id = p.收费细目id And (a.计算标志 = 6 And b.床位等级id Is Not Null Or a.计算标志=7) And
             Zl_Date_Half(Nvl(b.终止时间, Sysdate)) >= Zl_Date_Half(p.执行日期) And
             Zl_Date_Half(b.开始时间) <= Zl_Date_Half(Nvl(p.终止日期, Sysdate) + 1) And
             Zl_Date_Half(Least(Nvl(b.终止时间, Sysdate), Nvl(p.终止日期, Sysdate + 30) + 1)) >=
             Zl_Date_Half(Nvl(a.启用日期, Add_Months(Sysdate, -2))) And p.价格等级 Is Null) P
Where i.病人id = p.病人id And a.病人id = p.病人id And a.主页id = p.主页id;

Create Or Replace View 出院病人自动记帐 as
Select p.类型, p.病人id, p.主页id, Nvl(a.姓名, i.姓名) As 姓名, Nvl(a.性别, i.性别) As 性别, Nvl(a.年龄, i.年龄) As 年龄, Nvl(a.住院号, i.住院号) As 住院号,
       a.费别, p.科室id, p.病区id, p.床号, p.附加床位, p.收费细目id, p.收入项目id, 1 As 标志, p.现价 As 标准单价, p.开始日期, p.终止日期,
       p.终止日期 - p.开始日期 As 天数, p.数量, p.经治医师, p.责任护士, p.操作员编号, p.操作员姓名,p.医疗小组id
From 病人信息 I, 病案主页 A,
     (Select 2 As 类型, b.病人id, b.主页id, b.科室id, b.病区id, b.床号, b.附加床位, p.收费细目id, p.收入项目id, p.现价, b.经治医师, b.责任护士, b.操作员编号, b.操作员姓名,
              Zl_Date_Half(Greatest(Least(Nvl(b.上次计算时间, b.开始时间), Nvl(b.终止时间, Greatest(Nvl(b.上次计算时间, b.开始时间))),
                                           Greatest(Nvl(b.上次计算时间, b.开始时间))), p.执行日期, Nvl(a.启用日期, Add_Months(Sysdate, -2)))) As 开始日期,
              Zl_Date_Half(Least(Nvl(b.终止时间, Greatest(b.开始时间, Sysdate)), Nvl(p.终止日期, Sysdate + 30) + 1)) As 终止日期, b.数量,
               b.医疗小组id
       From 自动计价项目 A,
            (Select 病人id, 主页id, 开始时间, 附加床位, 病区id, 科室id, 床号, 床位等级id, 1 As 数量, 责任护士, 经治医师, 终止时间, 操作员编号, 操作员姓名, 上次计算时间,
                     a.医疗小组id
              From 病人变动记录 A
              Where 开始原因 <> 10
              Union All
              Select 病人id, 主页id, 开始时间, 附加床位, 病区id, 科室id, 床号, i.从项id As 床位等级id, i.从项数次 As 数量, 责任护士, 经治医师, 终止时间, 操作员编号, 操作员姓名,
                     上次计算时间, b.医疗小组id
              From 病人变动记录 B, 收费从属项目 I
              Where b.床位等级id = i.主项id And b.开始原因 <> 10 And i.固有从属 > 0) B, 收费价目 P
       Where a.病区id = b.病区id And Zl_Date_Half(Nvl(b.上次计算时间, b.开始时间)) <> Zl_Date_Half(Nvl(b.终止时间, Sysdate)) And p.现价 <> 0 And
             a.计算标志 = 1 And b.床位等级id = p.收费细目id And Zl_Date_Half(Nvl(b.终止时间, Sysdate)) >= Zl_Date_Half(p.执行日期) And
             Zl_Date_Half(b.开始时间) <= Zl_Date_Half(Nvl(p.终止日期, Sysdate) + 1) And
             Zl_Date_Half(Least(Nvl(b.终止时间, Sysdate), Nvl(p.终止日期, Sysdate + 30) + 1)) >=
             Zl_Date_Half(Nvl(a.启用日期, Add_Months(Sysdate, -2))) And p.价格等级 Is Null
       Union All
       Select 1 As 类型, b.病人id, b.主页id, b.科室id, b.病区id, b.床号, b.附加床位, p.收费细目id, p.收入项目id, p.现价, b.经治医师, b.责任护士, b.操作员编号, b.操作员姓名,
              Zl_Date_Half(Greatest(Least(Nvl(b.上次计算时间, b.开始时间), Nvl(b.终止时间, Greatest(Nvl(b.上次计算时间, b.开始时间))),
                                           Greatest(Nvl(b.上次计算时间, b.开始时间))), p.执行日期, Nvl(a.启用日期, Add_Months(Sysdate, -2)))) As 开始日期,
              Zl_Date_Half(Least(Nvl(b.终止时间, Greatest(b.开始时间, Sysdate)), Nvl(p.终止日期, Sysdate + 30) + 1)) As 终止日期, b.数量,
              b.医疗小组id
       From 自动计价项目 A,
            (Select 病人id, 主页id, 开始时间, 附加床位, 病区id, 科室id, 床号, 护理等级id, 1 As 数量, 责任护士, 经治医师, 终止时间, 操作员编号, 操作员姓名, 上次计算时间, 医疗小组id
              From 病人变动记录
              Where 开始原因 <> 10
              Union All
              Select 病人id, 主页id, 开始时间, 附加床位, 病区id, 科室id, 床号, i.从项id As 护理等级id, i.从项数次 As 数量, 责任护士, 经治医师, 终止时间, 操作员编号, 操作员姓名,
                     上次计算时间, b.医疗小组id
              From 病人变动记录 B, 收费从属项目 I
              Where b.护理等级id = i.主项id And b.开始原因 <> 10 And i.固有从属 > 0) B, 收费价目 P, 收费项目目录 C
       Where a.病区id = b.病区id And b.附加床位 <> 1 And
             Zl_Date_Half(Nvl(b.上次计算时间, b.开始时间)) <> Zl_Date_Half(Nvl(b.终止时间, Sysdate)) And p.现价 <> 0 And a.计算标志 = 2 And
             b.护理等级id = p.收费细目id And b.护理等级id = c.Id And Nvl(c.计算方式, 0) <> 1 And
             Zl_Date_Half(Nvl(b.终止时间, Sysdate)) >= Zl_Date_Half(p.执行日期) And
             Zl_Date_Half(b.开始时间) <= Zl_Date_Half(Nvl(p.终止日期, Sysdate) + 1) And
             Zl_Date_Half(Least(Nvl(b.终止时间, Sysdate), Nvl(p.终止日期, Sysdate + 30) + 1)) >=
             Zl_Date_Half(Nvl(a.启用日期, Add_Months(Sysdate, -2))) And p.价格等级 Is Null
       Union All
       Select 3 As 类型, b.病人id, b.主页id, b.科室id, b.病区id, b.床号, b.附加床位, p.收费细目id, p.收入项目id, p.现价, b.经治医师, b.责任护士, b.操作员编号, b.操作员姓名,
              Zl_Date_Half(Greatest(Least(Nvl(b.上次计算时间, b.开始时间), Nvl(b.终止时间, Greatest(Nvl(b.上次计算时间, b.开始时间))),
                                           Greatest(Nvl(b.上次计算时间, b.开始时间))), p.执行日期, Nvl(a.启用日期, Add_Months(Sysdate, -2)))) As 开始日期,
              Zl_Date_Half(Least(Nvl(b.终止时间, Greatest(b.开始时间, Sysdate)), Nvl(p.终止日期, Sysdate + 30) + 1)) As 终止日期, a.数量,
              b.医疗小组id
       From (Select 病区id, 计算标志, 收费细目id, 1 As 数量, 启用日期
              From 自动计价项目
              Union All
              Select 病区id, 计算标志, 从项id, i.从项数次 As 数量, 启用日期
              From 自动计价项目 A, 收费从属项目 I
              Where a.收费细目id = i.主项id And i.固有从属 > 0) A, 病人变动记录 B, 收费价目 P
       Where a.病区id = b.病区id And b.附加床位 <> 1 And b.开始原因 <> 10 And
             Zl_Date_Half(Nvl(b.上次计算时间, b.开始时间)) <> Zl_Date_Half(Nvl(b.终止时间, Sysdate)) And p.现价 <> 0 And
             a.收费细目id = p.收费细目id And (a.计算标志 = 6 And b.床位等级id Is Not Null Or a.计算标志 =7) And
             Zl_Date_Half(Nvl(b.终止时间, Sysdate)) >= Zl_Date_Half(p.执行日期) And
             Zl_Date_Half(b.开始时间) <= Zl_Date_Half(Nvl(p.终止日期, Sysdate) + 1) And
             Zl_Date_Half(Least(Nvl(b.终止时间, Sysdate), Nvl(p.终止日期, Sysdate + 30) + 1)) >=
             Zl_Date_Half(Nvl(a.启用日期, Add_Months(Sysdate, -2))) And p.价格等级 Is Null) P
Where i.病人id = p.病人id And a.病人id = p.病人id And a.主页id = p.主页id;

create or replace view xw_ris_studyinfo as
  Select b.医嘱ID As OrderID,a.病人ID As PatientID,a.门诊号 As OutPatientID,a.住院号 As InPatientID,a.健康号 As HealthID,
        a.姓名 As Name,a.性别 As Sex ,a.年龄 As Age,a.出生日期 As DateOfBirth,b.英文名 As PYName,b.影像类别 As Modality,
        b.检查号 As StudyID,c.病人来源 As Source,c.执行科室ID As DeptID,c.医嘱内容 As MedicalOrder,
        c.开嘱时间 As ApplyTime,d.首次时间 As CheckInTime
  From 病人信息 a,影像检查记录 b,病人医嘱记录 c,病人医嘱发送 d
  Where b.医嘱ID = c.Id And c.Id = d.医嘱ID And c.病人ID =a.病人ID;

create or replace view xw_pacs_imagepath as
  Select a.医嘱ID As OrderID,b.IP地址 As ServerIP,b.FTP目录 As RootPath,b.共享目录用户名 As ServerUserName,
         b.共享目录密码 As ServerPassWord,Decode(a.接收日期,Null,'',to_Char(a.接收日期,'YYYYMMDD')||'\')||a.检查UID||'\'||d.图像Uid As ImagePathName,
         a.检查UID As StudyUID,c.序列UID As SeriesUID,d.图像Uid As ImageUID,
         'FTP[;]'||b.IP地址||'[;]21[;]'||b.FTP用户名||'[;]'||b.FTP密码||'[;]'||'\'||b.FTP目录||'\[;]'||d.图像Uid As FTPString
  From 影像检查记录 a,影像设备目录 b,影像检查序列 c,影像检查图象 d
  Where a.位置一 =b.设备号 And C.检查UID = A.检查UID And D.序列UID = C.序列UID;

create or replace view xw_ris_wlm_info as
  Select '' As F_MACHINE_AET,'' As F_MACHINE_NAME , 影像类别 As F_MODALITY_DCMTYPE ,to_char(出生日期,'YYYYMMDD') As F_PAT_BIRTH ,英文名 As F_PAT_NAME ,英文名 As F_PAT_NAME_EN
         ,'' As F_PAT_ADDRESS,检查号 As F_PAT_NO,'' As F_PAT_OT_ID,'' As F_PAT_LOCATION,'' As F_ADD_HISTORY,'' As F_PAT_REGION,'' As F_MEDICAL_ALERTS,'' As F_CONTRAST,'' As F_PLACE_NO
         ,decode(性别,'男','M','女','F','O') As F_SEX,体重 As F_WEIGHT,身高 As F_HEIGHT,'' As F_PERFORM_DOC,'' As F_REQUEST_DOC,'' As F_DIAGNOSES,'' As F_STU_REASON,'' As F_STU_COMMENT
         ,'' As F_MEN_DATE,'' As F_LATERALITY,to_char(b.首次时间,'YYYYMMDD') As F_STU_DATE_DCM,a.医嘱ID As F_STU_ID,a.医嘱ID As F_STU_NO,to_char(b.首次时间,'hh24:mi:ss') As F_STU_TIME_DCM
         ,a.医嘱ID || '.' || a.发送号 As F_STU_UID,b.执行间 as F_ROOM_NAME, c.名称 as F_DEPT_NAME 
  From  影像检查记录 a ,病人医嘱发送 b,部门表 c
  Where a.医嘱ID=b.医嘱ID And a.发送号 = b.发送号 And b.执行部门id=c.Id And  b.执行状态=3 And b.执行过程=2 And a.检查UID IS Null AND b.首次时间>=SysDate-10;

create or replace view xw_ris_wlm_info_cn as
  Select '' As F_MACHINE_AET,'' As F_MACHINE_NAME , 影像类别 As F_MODALITY_DCMTYPE ,to_char(出生日期,'YYYYMMDD') As F_PAT_BIRTH ,姓名 As F_PAT_NAME ,英文名 As F_PAT_NAME_EN
         ,'' As F_PAT_ADDRESS,检查号 As F_PAT_NO,'' As F_PAT_OT_ID,'' As F_PAT_LOCATION,'' As F_ADD_HISTORY,'' As F_PAT_REGION,'' As F_MEDICAL_ALERTS,'' As F_CONTRAST,'' As F_PLACE_NO
         ,decode(性别,'男','M','女','F','O') As F_SEX,体重 As F_WEIGHT,身高 As F_HEIGHT,'' As F_PERFORM_DOC,'' As F_REQUEST_DOC,'' As F_DIAGNOSES,'' As F_STU_REASON,'' As F_STU_COMMENT
         ,'' As F_MEN_DATE,'' As F_LATERALITY,to_char(b.首次时间,'YYYYMMDD') As F_STU_DATE_DCM,a.医嘱ID As F_STU_ID,a.医嘱ID As F_STU_NO,to_char(b.首次时间,'hh24:mi:ss') As F_STU_TIME_DCM
         ,a.医嘱ID || '.' || a.发送号 As F_STU_UID,b.执行间 as F_ROOM_NAME, c.名称 as F_DEPT_NAME
  From  影像检查记录 a ,病人医嘱发送 b,部门表 c
  Where a.医嘱ID=b.医嘱ID And a.发送号 = b.发送号  And b.执行部门id=c.Id And  b.执行状态=3 And b.执行过程=2 And a.检查UID IS Null AND b.首次时间>=SysDate-10;


CREATE OR REPLACE VIEW 收费类别 AS 
    SELECT 编码,名称 AS 类别,简码 AS 说明,固定 AS 系统标志,0 AS 独立编辑
    FROM 收费项目类别;

CREATE OR REPLACE VIEW 收费细目 AS
SELECT 类别,ID, NULL AS 上级id,1 AS 末级,编码,名称,规格||'┆'||产地 AS 规格,计算单位,说明,
        费用类型,服务对象,0 AS 独立编辑,屏蔽费别,是否变价,加班加价,补充摘要,
        decode(执行科室,1,2,2,2,3,3,4,1,0) As 执行科室,标识主码,标识子码,建档时间,撤档时间
FROM 收费项目目录;

CREATE OR REPLACE VIEW 收费别名 AS 
    SELECT 收费细目id,名称,简码
    FROM 收费项目别名;

CREATE OR REPLACE VIEW 收费执行部门 AS 
    SELECT 收费细目id,执行科室id as 执行部门id
    FROM 收费执行科室;

CREATE OR REPLACE VIEW 挂号项目 AS 
    SELECT I.ID AS 序号, I.编码, I.名称, I.计算单位, N.简码, I.项目特性 AS 急诊标记, I.说明, I.建档时间, I.撤档时间
    FROM 收费项目目录 I, 收费项目别名 N
    WHERE I.ID=N.收费细目id And I.类别='1' And N.性质=1 And N.码类=1;

CREATE OR REPLACE VIEW 床位等级 AS 
    SELECT I.ID AS 序号, I.编码, I.名称, I.计算单位, N.简码, I.说明, I.建档时间, I.撤档时间
    FROM 收费项目目录 I, 收费项目别名 N
    WHERE I.ID=N.收费细目id And I.类别='J' And N.性质=1 And N.码类=1;

CREATE OR REPLACE VIEW 护理等级 AS 
    SELECT I.ID AS 序号, I.编码, I.名称, I.计算单位, N.简码, I.项目特性-1 AS 基本护理, I.说明, I.建档时间, I.撤档时间
    FROM 收费项目目录 I, 收费项目别名 N
    WHERE I.ID=N.收费细目id And I.类别='H' And I.项目特性>=1 And N.性质=1 And N.码类=1;
    
CREATE OR REPLACE VIEW 药品材质分类 AS 
    SELECT decode(编码,'5','1','6','2','3') AS 编码,名称, 简码
    FROM 诊疗项目类别
    WHERE 编码 in ('5','6','7');

CREATE OR REPLACE VIEW 药品用途分类 AS 
    SELECT decode(类型,1,'西成药',2,'中成药','中草药') AS 材质, ID, 编码, 名称, 简码, 上级id,1 AS 末级
    FROM 诊疗分类目录
    WHERE 类型 in (1,2,3) And (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'));

CREATE OR REPLACE VIEW 药品信息 AS 
    SELECT decode(I.类别,'5','西成药','6','中成药','中草药') AS 材质分类, S.药名ID, null as 药典id,I.分类ID AS 用途分类ID,
        K.编码 AS 剂型,I.编码, I.名称 AS 通用名称, I.计算单位 AS 剂量单位, S.毒理分类, S.货源情况, S.价值分类, S.用药梯次, 
        S.处方职务, S.急救药否, S.是否新药, S.是否皮试, S.是否原料, S.处方限量, S.药品类型, I.建档时间, I.撤档时间
    FROM 诊疗项目目录 I, 药品特性 S,药品剂型 K
    WHERE I.ID=S.药名ID AND S.药品剂型=K.名称(+) And I.类别 In ('5','6','7');

CREATE OR REPLACE VIEW 药品目录 AS 
    SELECT S.药品id, S.药名ID, I.编码, I.名称, I.规格, I.产地, I.计算单位 AS 售价单位, 
            S.剂量系数, S.门诊单位, S.门诊包装, S.住院单位, S.住院包装, S.药库单位, S.药库包装, 
            S.最大效期, S.药品来源, S.协定药品, S.自制药品, S.批准文号, S.标识码, 
            S.药价级别, I.是否变价, S.指导批发价, S.指导零售价, S.指导差价率, S.扣率, S.住院可否分零, S.门诊可否分零,
            S.药库分批 AS 分批核算, S.药房分批 AS 药房分批核算, I.费用类型, decode(I.服务对象,1,'100',2,'010',3,'110','000') AS 服务对象,
            S.招标药品,S.差价让利比,S.GMP认证,I.建档时间, I.撤档时间
    FROM 收费项目目录 I, 药品规格  S
    WHERE I.ID=S.药品id And I.类别 In ('5','6','7');

CREATE OR REPLACE VIEW 药品配伍禁忌 AS 
    SELECT R.组编号, R.项目ID AS 药名ID, R.类型
    FROM 诊疗互斥项目  R, 诊疗项目目录  I
    WHERE R.项目ID=I.ID And I.类别 In ('5','6','7');

Create Or Replace View 药品别名 As 
    Select T.药名id,N.名称,N.简码,decode(性质,3,N.收费细目id,Null) As 药品id,
           decode(N.码类,3,2,1) As 码类
    From 收费项目别名 N,药品规格 T
    Where N.收费细目id=T.药品id And N.码类<>2;

Create Or Replace View 挂号病人 As
  Select Distinct NO As 挂号单号, 病人id, 姓名, 性别, 年龄, 收费细目id As 挂号项目, 加班标志 As 急诊, 登记时间 As 日期, 执行部门id As 科室id, 发药窗口 As 诊室,
                  执行人 As 医生, 执行状态 As 状态
  From 门诊费用记录
  Where 记录性质 = 4 And 记录状态 = 1 And 收费类别 = '1' And 病人id Is Not Null And
        登记时间 > (Select Trunc(Sysdate) - To_Number(参数值)
                From zlParameters
                Where 系统 = (Select 编号
                            From zlSystems
                            Where Upper(所有者) = (Select Username From All_Users Where User_Id = Userenv('SchemaID')) And
                                  Trunc(编号 / 100) = 1) And 模块 Is Null And Nvl(私有, 0) = 0 And 参数号 = 21);

--提供给北航冠新的病人费用明细：
Create Or Replace View 病人费用明细 As 
Select L.病人id,
	   L.Id As 顺序号,
	   S.项目编码,
	   S.项目名称,
	   L.标准单价,
	   L.登记时间 As 收费日期
  From  (Select Id, 病人ID,收费细目ID,标准单价,登记时间 From  门诊费用记录 
	     Union ALL Select Id, 病人ID,收费细目ID,标准单价,登记时间 From  住院费用记录 
	     )  L, 收费项目目录 I, 标准医价规范 S
 Where L.收费细目id = I.Id And I.标识主码 = S.项目编码(+) And
	   I.类别 Not In ('4', '5', '6', '7');

--因医保接口要与ZLHIS9兼容而保留
Create Or Replace View 诊断情况 As 
Select 病人id, 主页Id,疾病id,诊断描述 As 描述信息,诊断类型, 
   出院情况,诊断次序,编码序号, 是否未治, 是否疑诊, 录入次序,编码类别
From 病人诊断记录 Where 记录来源=2;

--为保持与以前版本兼容
Create OR REPLACE View 医保核对表 AS Select 结帐ID,结算方式,金额 From 保险结算明细 Where 标志=1;

CREATE OR REPLACE VIEW 保险帐户 AS 
	SELECT A.病人ID,B.* FROM 医保病人关联表 A,医保病人档案 B 
	WHERE A.险类=B.险类 AND A.中心=B.中心 AND A.医保号=B.医保号 AND A.标志=1; 

--电子病案审查归档
create or replace view 病案评分标准视图 as
select decode(T.上级序号,null,序号,T.上级序号) as 上级序号, decode(T.序号,null,T.ID,T.序号) as 序号,T.ID,T.上级ID,T.方案ID,T.项目,T.标准分值,T.基本要求,T.缺陷内容,T.扣分标准,decode(T.子项个数,0,'否','是') as 隐藏,T.否决等级
from
(
  select B.上级序号,A.序号,A.方案ID,
  A.ID,
  A.上级ID,
  decode(A.子项个数,0,decode(A.上级ID,Null,A.名称,B.名称),A.名称) as 项目,
  decode(A.子项个数,0,decode(A.上级ID,Null,A.标准分值,B.标准分值),B.标准分值) as 标准分值,
  decode(A.子项个数,0,decode(A.上级ID,Null,A.描述,B.描述),A.描述) as 基本要求,
  decode(A.上级ID,Null,'',A.描述) as 缺陷内容,
  DECODE(A.缺陷等级,NULL,decode(sign(A.标准分值-1),-1,To_CHAR(A.标准分值,'0.9'),To_Char(A.标准分值))||decode(A.评分单位,NULL,'','/'||A.评分单位),A.缺陷等级) as 扣分标准,
  A.子项个数,
  A.否决等级
  from
      (
          select AA.序号,AA.ID,AA.方案ID,AA.上级ID,AA.名称,AA.描述,AA.标准分值,AA.缺陷等级,AA.评分单位,AA.否决等级,count(BB.ID) as 子项个数
          from 病案评分标准 AA,病案评分标准 BB
          where AA.ID=BB.上级ID(+)
          group by AA.序号,AA.ID,AA.方案ID,AA.上级ID,AA.名称,AA.描述,AA.标准分值,AA.缺陷等级,AA.评分单位,AA.否决等级
      ) A,
      (
          select 序号 as 上级序号,ID,名称,标准分值,描述 from 病案评分标准
      ) B
  where A.上级ID=B.ID(+)
) T
order by decode(T.上级序号,null,序号,T.上级序号),decode(T.序号,null,T.ID,T.序号);

create or replace view 病案质量报表视图 as
Select   Tb.姓名, Tb.性别, Ta."病人ID",Ta."主页ID",Ta."住院号",Ta."入院日期",Ta."出院日期",Ta."入院科室",Ta."出院科室",Ta."门诊医师",Ta."责任护士",Ta."住院医师",Ta."编目日期",Ta."结果ID",Ta."方案ID",Ta."总分",Ta."等级",Ta."评分人",Ta."评分时间",Ta."审核人",Ta."审核时间",Ta."返回修改",Ta."备注",Ta."病理类型" 
From (Select T1.病人id, T1.主页id,T1.住院号, T1.入院日期, T1.出院日期, T2.名称 As 入院科室, T3.名称 As 出院科室, T1.门诊医师, 
              T1.责任护士, T1.住院医师, T1.编目日期, T1.结果id, T1.方案id, T1.总分, T1.等级, T1.评分人, 
              To_Char(T1.评分时间, 'YYYY-MM-DD') As 评分时间, T1.审核人, To_Char(T1.审核时间, 'YYYY-MM-DD') As 审核时间, 
              T1.返回修改, T1.备注 ,T1.病理类型
       From (Select A.病人id, A.主页id, A.入院科室id, A.出院科室id, A.入院日期, A.出院日期, A.门诊医师, A.责任护士, 
                     A.住院医师, A.编目日期, B.ID As 结果id, B.方案id, B.总分, B.等级, B.评分人, B.评分时间, B.审核人, 
                     B.审核时间, B.返回修改, B.备注,B.病理类型,A.住院号 
              From 病案主页 A, 病案评分结果 B 
              Where A.病人id = B.病人id(+) And A.主页id = B.主页id(+)) T1, 部门表 T2, 部门表 T3 
       Where T1.入院科室id = T2.ID And T1.出院科室id = T3.ID) Ta, 病人信息 Tb 
Where Ta.病人id = Tb.病人id;