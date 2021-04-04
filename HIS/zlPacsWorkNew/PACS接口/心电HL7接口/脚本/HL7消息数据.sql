prompt PL/SQL Developer import file
prompt Created on 2012年3月28日 by HJ
set feedback off
set define off
prompt Loading HL7服务配置...
insert into HL7服务配置 (ID, IP地址, 端口号, 服务类型, 发送程序名称, 发送设备名称, 接收程序名称, 接收设备名称)
values (1, '127.0.0.1', '104', 1, 'MUSE ECG Result 1', 'MEI MUSE', 'ZLHIS', 'HIS001');
insert into HL7服务配置 (ID, IP地址, 端口号, 服务类型, 发送程序名称, 发送设备名称, 接收程序名称, 接收设备名称)
values (2, '127.0.0.1', '1024', 2, 'ZLHIS', 'HIS001', 'MUSE ECG Result 1', 'MEI MUSE');
commit;
prompt 2 records loaded
prompt Loading HL7消息定义...
insert into HL7消息定义 (ID, 服务ID, 动作类型, 消息名称, 消息类型, 消息段组合)
values (1, 1, '接收心电结果', 'ORU_R01', null, 'MSH|PID|[PV1]|{OBR|[{DG1}]|[{NTE}]}|ZEX|ZPH|[{OBX}]');
insert into HL7消息定义 (ID, 服务ID, 动作类型, 消息名称, 消息类型, 消息段组合)
values (2, 2, '发送新医嘱', 'ORM_O01', 'NW', 'MSH|PID|PV1|ORC|OBR');
insert into HL7消息定义 (ID, 服务ID, 动作类型, 消息名称, 消息类型, 消息段组合)
values (3, 2, '发送取消医嘱', 'ORM_O01', 'CA', 'MSH|PID|PV1|ORC|OBR');
insert into HL7消息定义 (ID, 服务ID, 动作类型, 消息名称, 消息类型, 消息段组合)
values (4, 2, '发送删除医嘱', 'ORM_O01', 'DD', 'MSH|PID|PV1|ORC|OBR');
commit;
prompt 4 records loaded
prompt Loading HL7消息段配置...
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (268, 4, 'MSH', 1, 'ST', null, '^~\&', 'Encoding Characters 代码字符');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (2, 2, 'MSH', 1, 'ST', null, '^~\&', 'Encoding Characters 代码字符');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (3, 2, 'MSH', 2, 'HD', 'MUSE ECG Result 1', 'ZLHIS', 'Sending Application 发送程序');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (4, 2, 'MSH', 3, 'HD', 'MEI MUSE', 'HIS001', 'Sending Facility 发送设备');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (5, 2, 'MSH', 4, 'HD', 'ZLHIS', 'MUSE ECG Result 1', 'Receiving Application 接收程序');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (6, 2, 'MSH', 5, 'HD', 'HIS001', 'MEI MUSE', 'Receiving Facility 接收设备');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (7, 2, 'MSH', 6, 'TS', null, '[当前时间]', 'Date/Time Of Message 消息的日期/时间');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (8, 2, 'MSH', 7, 'ST', null, null, 'Security 安全性');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (9, 2, 'MSH', 8, 'CM', 'ORM^O01', 'ORM^O01', 'Message Type 消息类型');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (10, 2, 'MSH', 9, 'ST', null, '[当前时间]', 'Message Control ID 消息控制ID');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (11, 2, 'MSH', 10, 'PT', 'P', 'P', 'Processing ID 处理ID');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (12, 2, 'MSH', 11, 'VID', '2.4', '2.4', 'Version ID 版本ID');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (135, 3, 'MSH', 1, 'ST', null, '^~\&', 'Encoding Characters 代码字符');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (136, 3, 'MSH', 2, 'HD', 'MUSE ECG Result 1', 'ZLHIS', 'Sending Application 发送程序');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (137, 3, 'MSH', 3, 'HD', 'MEI MUSE', 'HIS001', 'Sending Facility 发送设备');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (138, 3, 'MSH', 4, 'HD', 'ZLHIS', 'MUSE ECG Result 1', 'Receiving Application 接收程序');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (139, 3, 'MSH', 5, 'HD', 'HIS001', 'MEI MUSE', 'Receiving Facility 接收设备');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (140, 3, 'MSH', 6, 'TS', null, '[当前时间]', 'Date/Time Of Message 消息的日期/时间');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (141, 3, 'MSH', 7, 'ST', null, null, 'Security 安全性');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (142, 3, 'MSH', 8, 'CM', 'ORM^O01', 'ORM^O01', 'Message Type 消息类型');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (143, 3, 'MSH', 9, 'ST', null, '[当前时间]', 'Message Control ID 消息控制ID');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (144, 3, 'MSH', 10, 'PT', 'P', 'P', 'Processing ID 处理ID');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (145, 3, 'MSH', 11, 'VID', '2.4', '2.4', 'Version ID 版本ID');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (146, 3, 'PID', 1, 'SI', null, '1', 'Set ID - Patient ID 设置ID - 患者ID');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (147, 3, 'PID', 2, 'CX', null, null, 'Patient ID 患者ID');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (148, 3, 'PID', 3, 'CX', null, '[标识号] ', 'Patient Identifier List 患者标识符表');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (149, 3, 'PID', 4, 'CX', null, '[病人ID]', 'Alternate Patient ID 备用患者ID');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (150, 3, 'PID', 5, 'XPN', null, '[姓名]', 'Patient Name 患者姓名');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (151, 3, 'PID', 6, 'XPN', null, null, 'Mother''s Maiden Name 母亲的婚前姓');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (152, 3, 'PID', 7, 'TS', null, '[出生日期]', 'Date/Time of Birth 出生日期');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (153, 3, 'PID', 8, 'IS', null, '[性别]', 'Sex 性别');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (154, 3, 'PID', 9, 'XPN', null, '[姓名]', 'Patient Alias 患者别名');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (155, 3, 'PID', 10, 'CE', null, 'C', 'Race 种族');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (156, 3, 'PID', 11, 'XAD', null, '[联系人地址]', 'Patient Address 患者住址');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (157, 3, 'PID', 12, 'IS', null, null, 'County Code 县代码');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (158, 3, 'PID', 13, 'XTN', null, '[家庭电话]', 'Phone Number - Home 家庭电话');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (159, 3, 'PID', 14, 'XTN', null, '[联系人电话]', 'Phone Number - Business 单位电话');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (160, 3, 'PID', 15, 'CE', null, null, 'Primary Language 母语');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (161, 3, 'PID', 16, 'IS', null, '[婚姻状况]', 'Marital Status 婚姻状况');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (162, 3, 'PID', 17, 'CE', null, null, 'Religion 宗教信仰');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (163, 3, 'PID', 18, 'CX', null, '[身份证号]', 'Patient Account Number患者账号');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (164, 3, 'PID', 19, 'ST', null, null, 'SSN Number - Patient 患者社会保险号码');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (165, 3, 'PV1', 1, 'SI', null, '1', 'Set ID-PV1设置ID-PV1');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (166, 3, 'PV1', 2, 'IS', null, '[病人来源]', 'Patient Class 患者类别');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (167, 3, 'PV1', 3, 'PL', null, '[当前科室名称]^[病区名称]^[床号]^^^^^^[当前科室名称]', 'Assigned Patient Location 指定患者位置');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (168, 3, 'PV1', 4, 'IS', null, null, 'Admission Type 入院类型');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (169, 3, 'PV1', 5, 'CX', null, null, 'Preadmit Number 预收入院号码');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (170, 3, 'PV1', 6, 'PL', null, null, 'Prior Patient Location 患者原位置');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (171, 3, 'PV1', 7, 'XCN', null, '[开嘱医生]', 'Attending Doctor 接诊医生');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (172, 3, 'PV1', 8, 'XCN', null, null, 'Referring Doctor 转诊医生');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (173, 3, 'PV1', 9, 'XCN', null, null, 'Consulting Doctor 会诊医生');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (174, 3, 'PV1', 10, 'IS', null, null, 'Hospital Service 医院服务');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (175, 3, 'PV1', 11, 'PL', null, null, 'Temporary Location 临时位置');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (176, 3, 'PV1', 12, 'IS', null, null, 'Preadmit Test Indicator 预收入院检验标识符');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (177, 3, 'PV1', 13, 'IS', null, null, 'Readmission Indicator 再次入院标识符');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (178, 3, 'PV1', 14, 'IS', null, null, 'Admit Source 入院来源');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (179, 3, 'PV1', 15, 'IS', null, null, 'Ambulatory Status （手术后）走动状态');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (180, 3, 'PV1', 16, 'IS', null, null, 'VIP Indicator VIP标识符');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (181, 3, 'PV1', 17, 'XCN', null, null, 'Admitting Doctor 入院医生');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (182, 3, 'PV1', 18, 'IS', null, null, 'Patient Type 患者类型');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (183, 3, 'PV1', 19, 'CX', null, '1', 'Visit Number 就诊号码');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (184, 3, 'PV1', 20, 'FC', null, null, 'Financial Class 经济状况类别');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (185, 3, 'PV1', 21, 'IS', null, null, 'Charge Price Indicator 费用价格标识符');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (186, 3, 'PV1', 22, 'IS', null, null, 'Courtesy Code 礼貌代码');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (187, 3, 'PV1', 23, 'IS', null, null, 'Credit Rating 信用等级');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (188, 3, 'PV1', 24, 'IS', null, null, 'Contract Code 合同代码');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (189, 3, 'PV1', 25, 'DT', null, null, 'Contract Effective Date 合同生效日期');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (190, 3, 'PV1', 26, 'NM', null, null, 'Contract Amount 合同总量');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (191, 3, 'PV1', 27, 'NM', null, null, 'Contract Period 合同期限');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (192, 3, 'PV1', 28, 'IS', null, null, 'Interest Code 利率代码');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (193, 3, 'PV1', 29, 'IS', null, null, 'Transfer to Bad Debt Code 转为坏账代码');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (194, 3, 'PV1', 30, 'DT', null, null, 'Transfer to Bad Debt Date 转为坏账日期');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (195, 3, 'PV1', 31, 'IS', null, null, 'Bad Debt Agency Code 坏账代理代码');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (196, 3, 'PV1', 32, 'NM', null, null, 'Bad Debt Transfer Amount 转为坏账总额');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (197, 3, 'PV1', 33, 'NM', null, null, 'Bad Debt Recovery Amount 坏账恢复总额');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (198, 3, 'PV1', 34, 'IS', null, null, 'Delete Account Indicator 删除账户标识符');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (199, 3, 'PV1', 35, 'DT', null, null, 'Delete Account Date 删除账户日期');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (200, 3, 'PV1', 36, 'IS', null, 'UNK', 'Discharge Disposition 出院处置');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (201, 3, 'PV1', 37, 'CM', null, null, 'Discharge to Location 出院去往的位置');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (202, 3, 'PV1', 38, 'CE', null, null, 'Diet Type 饮食类型');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (203, 3, 'PV1', 39, 'IS', null, null, 'Servicing Facility 服务机构');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (204, 3, 'PV1', 40, 'IS', null, null, 'Bed Status 床位状况');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (205, 3, 'PV1', 41, 'IS', null, null, 'Account Status 账户状况');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (206, 3, 'PV1', 42, 'IS', null, null, 'Pending Location 待定位置');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (207, 3, 'PV1', 43, 'PL', null, null, 'Prior Temporary Location 前临时位置');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (208, 3, 'PV1', 44, 'TS', null, '[入院时间]', 'Admit Date/Time 入院日期/时间');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (209, 3, 'PV1', 45, 'TS', null, '[出院时间]', 'Discharge Date/Time 出院日期/时间');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (210, 3, 'PV1', 46, 'NM', null, null, 'Current Patient Balance 当前患者差额');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (211, 3, 'PV1', 47, 'NM', null, null, 'Total Charges 总费用');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (212, 3, 'PV1', 48, 'NM', null, null, 'Total Adjustments 总调度数');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (213, 3, 'PV1', 49, 'NM', null, null, 'Total Payments 总支付额');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (214, 3, 'PV1', 50, 'CX', null, null, 'Alternate Visit ID 备用就诊ID');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (215, 3, 'ORC', 1, 'ID', null, '[消息类型]', '医嘱控制码');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (216, 3, 'ORC', 2, 'EI', null, '[医嘱ID]', '开单者医嘱号码');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (217, 3, 'ORC', 3, 'EI', null, null, '执行者医嘱号码');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (218, 3, 'ORC', 4, 'EI', null, null, '开单者医嘱组号码');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (219, 3, 'ORC', 5, 'ID', null, null, '医嘱状态');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (220, 3, 'ORC', 6, 'ID', null, null, '应答标记');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (221, 3, 'ORC', 7, 'TQ', null, null, '数量/时间');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (222, 3, 'ORC', 8, 'CM', null, null, '父层标识');
commit;
prompt 100 records committed...
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (223, 3, 'ORC', 9, 'TS', null, '[开嘱时间]', '事务日期/时间');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (224, 3, 'ORC', 10, 'XCN', null, '[开嘱医生]', '录入者');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (225, 3, 'ORC', 11, 'XCN', null, '[校对护士]', '校正者');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (226, 3, 'ORC', 12, 'XCN', null, '[开嘱医生]', '医嘱提供者');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (231, 3, 'OBR', 1, 'SI', null, '1', '设置ID-OBR');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (232, 3, 'OBR', 2, 'EI', null, null, '开单者医嘱号码');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (233, 3, 'OBR', 3, 'EI', null, null, '执行者医嘱号码');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (234, 3, 'OBR', 4, 'CE', null, '[医嘱内容]', '通用服务标识');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (235, 3, 'OBR', 5, 'ID', null, null, '优先级-OBR');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (236, 3, 'OBR', 6, 'TS', null, '[开嘱时间]', '请求日期/时间');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (237, 3, 'OBR', 7, 'TS', null, '[开嘱时间]', '观察日期/时间#');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (238, 3, 'OBR', 8, 'TS', null, null, '观察结束日期/时间#');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (239, 3, 'OBR', 9, 'CQ', null, null, '样本收集量*');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (240, 3, 'OBR', 10, 'XCN', null, null, '收集者标识*');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (241, 3, 'OBR', 11, 'ID', null, null, '样本行为码*');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (242, 3, 'OBR', 12, 'CE', null, null, '危险因素代码');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (243, 3, 'OBR', 13, 'ST', null, null, '相关临床信息');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (244, 3, 'OBR', 14, 'TS', null, null, '收到样本日期/时间*');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (245, 3, 'OBR', 15, 'CM', null, null, '样本来源');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (246, 3, 'OBR', 16, 'XCN', null, '[开嘱医生]', '医嘱提供者');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (247, 3, 'OBR', 17, 'XTN', null, null, '医嘱回访电话号码');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (248, 3, 'OBR', 18, 'ST', null, null, '开单者字段1');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (249, 3, 'OBR', 19, 'ST', null, null, '开单者字段2');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (250, 3, 'OBR', 20, 'ST', null, null, '执行者字段1+');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (251, 3, 'OBR', 21, 'ST', null, null, '执行者字段2+');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (252, 3, 'OBR', 22, 'TS', null, null, '结果报告/状态改变日期/时间+');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (253, 3, 'OBR', 23, 'CM', null, null, '收费+');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (254, 3, 'OBR', 24, 'ID', null, null, '诊断服务标识ID');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (255, 3, 'OBR', 25, 'ID', null, null, '结果状态+');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (256, 3, 'OBR', 26, 'CM', null, null, '父层结果+');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (257, 3, 'OBR', 27, 'TQ', null, null, '数量/时间');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (258, 3, 'OBR', 28, 'XCN', null, null, '结果需要者');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (259, 3, 'OBR', 29, 'CM', null, null, '父层连结码');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (260, 3, 'OBR', 30, 'ID', null, null, '患者行动方式');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (261, 3, 'OBR', 31, 'CE', null, null, '研究原因');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (262, 3, 'OBX', 1, 'SI', null, null, 'Set ID -OBX 设置ID -OBX');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (263, 3, 'OBX', 2, 'ID', null, null, 'Value Type 值类型');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (264, 3, 'OBX', 3, 'CE', null, null, 'Observation Identifier 观察标识符');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (265, 3, 'OBX', 4, 'ST', null, null, 'Observation Sub -ID 观察子ID');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (266, 3, 'OBX', 5, 'TEXT', null, null, 'Observation Value 观察值');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (269, 4, 'MSH', 2, 'HD', 'MUSE ECG Result 1', 'ZLHIS', 'Sending Application 发送程序');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (13, 2, 'PID', 1, 'SI', null, '1', 'Set ID - Patient ID 设置ID - 患者ID');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (14, 2, 'PID', 2, 'CX', null, null, 'Patient ID 患者ID');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (15, 2, 'PID', 3, 'CX', null, '[标识号] ', 'Patient Identifier List 患者标识符表');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (16, 2, 'PID', 4, 'CX', null, '[病人ID]', 'Alternate Patient ID 备用患者ID');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (17, 2, 'PID', 5, 'XPN', null, '[姓名]', 'Patient Name 患者姓名');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (18, 2, 'PID', 6, 'XPN', null, null, 'Mother''s Maiden Name 母亲的婚前姓');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (19, 2, 'PID', 7, 'TS', null, '[出生日期]', 'Date/Time of Birth 出生日期');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (20, 2, 'PID', 8, 'IS', null, '[性别]', 'Sex 性别');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (21, 2, 'PID', 9, 'XPN', null, '[姓名]', 'Patient Alias 患者别名');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (22, 2, 'PID', 10, 'CE', null, 'C', 'Race 种族');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (23, 2, 'PID', 11, 'XAD', null, '[联系人地址]', 'Patient Address 患者住址');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (24, 2, 'PID', 12, 'IS', null, null, 'County Code 县代码');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (25, 2, 'PID', 13, 'XTN', null, '[家庭电话]', 'Phone Number - Home 家庭电话');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (26, 2, 'PID', 14, 'XTN', null, '[联系人电话]', 'Phone Number - Business 单位电话');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (27, 2, 'PID', 15, 'CE', null, null, 'Primary Language 母语');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (28, 2, 'PID', 16, 'IS', null, '[婚姻状况]', 'Marital Status 婚姻状况');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (29, 2, 'PID', 17, 'CE', null, null, 'Religion 宗教信仰');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (30, 2, 'PID', 18, 'CX', null, '[身份证号]', 'Patient Account Number患者账号');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (31, 2, 'PID', 19, 'ST', null, null, 'SSN Number - Patient 患者社会保险号码');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (32, 2, 'PV1', 1, 'SI', null, '1', 'Set ID-PV1设置ID-PV1');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (33, 2, 'PV1', 2, 'IS', null, '[病人来源]', 'Patient Class 患者类别');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (34, 2, 'PV1', 3, 'PL', null, '[当前科室名称]^[病区名称]^[床号]^^^^^^[当前科室名称]', 'Assigned Patient Location 指定患者位置');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (35, 2, 'PV1', 4, 'IS', null, null, 'Admission Type 入院类型');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (36, 2, 'PV1', 5, 'CX', null, null, 'Preadmit Number 预收入院号码');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (37, 2, 'PV1', 6, 'PL', null, null, 'Prior Patient Location 患者原位置');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (38, 2, 'PV1', 7, 'XCN', null, '[开嘱医生]', 'Attending Doctor 接诊医生');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (39, 2, 'PV1', 8, 'XCN', null, null, 'Referring Doctor 转诊医生');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (40, 2, 'PV1', 9, 'XCN', null, null, 'Consulting Doctor 会诊医生');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (41, 2, 'PV1', 10, 'IS', null, null, 'Hospital Service 医院服务');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (42, 2, 'PV1', 11, 'PL', null, null, 'Temporary Location 临时位置');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (43, 2, 'PV1', 12, 'IS', null, null, 'Preadmit Test Indicator 预收入院检验标识符');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (44, 2, 'PV1', 13, 'IS', null, null, 'Readmission Indicator 再次入院标识符');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (45, 2, 'PV1', 14, 'IS', null, null, 'Admit Source 入院来源');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (46, 2, 'PV1', 15, 'IS', null, null, 'Ambulatory Status （手术后）走动状态');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (47, 2, 'PV1', 16, 'IS', null, null, 'VIP Indicator VIP标识符');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (48, 2, 'PV1', 17, 'XCN', null, null, 'Admitting Doctor 入院医生');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (49, 2, 'PV1', 18, 'IS', null, null, 'Patient Type 患者类型');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (50, 2, 'PV1', 19, 'CX', null, '1', 'Visit Number 就诊号码');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (51, 2, 'PV1', 20, 'FC', null, null, 'Financial Class 经济状况类别');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (52, 2, 'PV1', 21, 'IS', null, null, 'Charge Price Indicator 费用价格标识符');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (53, 2, 'PV1', 22, 'IS', null, null, 'Courtesy Code 礼貌代码');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (54, 2, 'PV1', 23, 'IS', null, null, 'Credit Rating 信用等级');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (55, 2, 'PV1', 24, 'IS', null, null, 'Contract Code 合同代码');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (56, 2, 'PV1', 25, 'DT', null, null, 'Contract Effective Date 合同生效日期');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (57, 2, 'PV1', 26, 'NM', null, null, 'Contract Amount 合同总量');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (58, 2, 'PV1', 27, 'NM', null, null, 'Contract Period 合同期限');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (59, 2, 'PV1', 28, 'IS', null, null, 'Interest Code 利率代码');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (60, 2, 'PV1', 29, 'IS', null, null, 'Transfer to Bad Debt Code 转为坏账代码');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (61, 2, 'PV1', 30, 'DT', null, null, 'Transfer to Bad Debt Date 转为坏账日期');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (62, 2, 'PV1', 31, 'IS', null, null, 'Bad Debt Agency Code 坏账代理代码');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (63, 2, 'PV1', 32, 'NM', null, null, 'Bad Debt Transfer Amount 转为坏账总额');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (64, 2, 'PV1', 33, 'NM', null, null, 'Bad Debt Recovery Amount 坏账恢复总额');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (65, 2, 'PV1', 34, 'IS', null, null, 'Delete Account Indicator 删除账户标识符');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (66, 2, 'PV1', 35, 'DT', null, null, 'Delete Account Date 删除账户日期');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (67, 2, 'PV1', 36, 'IS', null, 'UNK', 'Discharge Disposition 出院处置');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (68, 2, 'PV1', 37, 'CM', null, null, 'Discharge to Location 出院去往的位置');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (69, 2, 'PV1', 38, 'CE', null, null, 'Diet Type 饮食类型');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (70, 2, 'PV1', 39, 'IS', null, null, 'Servicing Facility 服务机构');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (71, 2, 'PV1', 40, 'IS', null, null, 'Bed Status 床位状况');
commit;
prompt 200 records committed...
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (72, 2, 'PV1', 41, 'IS', null, null, 'Account Status 账户状况');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (73, 2, 'PV1', 42, 'IS', null, null, 'Pending Location 待定位置');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (74, 2, 'PV1', 43, 'PL', null, null, 'Prior Temporary Location 前临时位置');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (75, 2, 'PV1', 44, 'TS', null, '[入院时间]', 'Admit Date/Time 入院日期/时间');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (76, 2, 'PV1', 45, 'TS', null, '[出院时间]', 'Discharge Date/Time 出院日期/时间');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (77, 2, 'PV1', 46, 'NM', null, null, 'Current Patient Balance 当前患者差额');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (78, 2, 'PV1', 47, 'NM', null, null, 'Total Charges 总费用');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (79, 2, 'PV1', 48, 'NM', null, null, 'Total Adjustments 总调度数');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (80, 2, 'PV1', 49, 'NM', null, null, 'Total Payments 总支付额');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (81, 2, 'PV1', 50, 'CX', null, null, 'Alternate Visit ID 备用就诊ID');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (82, 2, 'ORC', 1, 'ID', null, '[消息类型]', '医嘱控制码');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (83, 2, 'ORC', 2, 'EI', null, '[医嘱ID]', '开单者医嘱号码');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (84, 2, 'ORC', 3, 'EI', null, null, '执行者医嘱号码');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (85, 2, 'ORC', 4, 'EI', null, null, '开单者医嘱组号码');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (86, 2, 'ORC', 5, 'ID', null, null, '医嘱状态');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (87, 2, 'ORC', 6, 'ID', null, null, '应答标记');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (88, 2, 'ORC', 7, 'TQ', null, null, '数量/时间');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (89, 2, 'ORC', 8, 'CM', null, null, '父层标识');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (90, 2, 'ORC', 9, 'TS', null, '[开嘱时间]', '事务日期/时间');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (91, 2, 'ORC', 10, 'XCN', null, '[开嘱医生]', '录入者');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (92, 2, 'ORC', 11, 'XCN', null, '[校对护士]', '校正者');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (93, 2, 'ORC', 12, 'XCN', null, '[开嘱医生]', '医嘱提供者');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (270, 4, 'MSH', 3, 'HD', 'MEI MUSE', 'HIS001', 'Sending Facility 发送设备');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (271, 4, 'MSH', 4, 'HD', 'ZLHIS', 'MUSE ECG Result 1', 'Receiving Application 接收程序');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (272, 4, 'MSH', 5, 'HD', 'HIS001', 'MEI MUSE', 'Receiving Facility 接收设备');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (273, 4, 'MSH', 6, 'TS', null, '[当前时间]', 'Date/Time Of Message 消息的日期/时间');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (98, 2, 'OBR', 1, 'SI', null, '1', '设置ID-OBR');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (99, 2, 'OBR', 2, 'EI', null, null, '开单者医嘱号码');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (100, 2, 'OBR', 3, 'EI', null, null, '执行者医嘱号码');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (101, 2, 'OBR', 4, 'CE', null, '[医嘱内容]', '通用服务标识');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (102, 2, 'OBR', 5, 'ID', null, null, '优先级-OBR');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (103, 2, 'OBR', 6, 'TS', null, '[开嘱时间]', '请求日期/时间');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (104, 2, 'OBR', 7, 'TS', null, '[开嘱时间]', '观察日期/时间#');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (105, 2, 'OBR', 8, 'TS', null, null, '观察结束日期/时间#');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (106, 2, 'OBR', 9, 'CQ', null, null, '样本收集量*');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (107, 2, 'OBR', 10, 'XCN', null, null, '收集者标识*');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (108, 2, 'OBR', 11, 'ID', null, null, '样本行为码*');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (109, 2, 'OBR', 12, 'CE', null, null, '危险因素代码');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (110, 2, 'OBR', 13, 'ST', null, null, '相关临床信息');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (111, 2, 'OBR', 14, 'TS', null, null, '收到样本日期/时间*');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (112, 2, 'OBR', 15, 'CM', null, null, '样本来源');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (113, 2, 'OBR', 16, 'XCN', null, '[开嘱医生]', '医嘱提供者');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (114, 2, 'OBR', 17, 'XTN', null, null, '医嘱回访电话号码');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (115, 2, 'OBR', 18, 'ST', null, null, '开单者字段1');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (116, 2, 'OBR', 19, 'ST', null, null, '开单者字段2');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (117, 2, 'OBR', 20, 'ST', null, null, '执行者字段1+');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (118, 2, 'OBR', 21, 'ST', null, null, '执行者字段2+');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (119, 2, 'OBR', 22, 'TS', null, null, '结果报告/状态改变日期/时间+');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (120, 2, 'OBR', 23, 'CM', null, null, '收费+');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (121, 2, 'OBR', 24, 'ID', null, null, '诊断服务标识ID');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (122, 2, 'OBR', 25, 'ID', null, null, '结果状态+');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (123, 2, 'OBR', 26, 'CM', null, null, '父层结果+');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (124, 2, 'OBR', 27, 'TQ', null, null, '数量/时间');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (125, 2, 'OBR', 28, 'XCN', null, null, '结果需要者');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (126, 2, 'OBR', 29, 'CM', null, null, '父层连结码');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (127, 2, 'OBR', 30, 'ID', null, null, '患者行动方式');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (128, 2, 'OBR', 31, 'CE', null, null, '研究原因');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (129, 2, 'OBX', 1, 'SI', null, null, 'Set ID -OBX 设置ID -OBX');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (130, 2, 'OBX', 2, 'ID', null, null, 'Value Type 值类型');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (131, 2, 'OBX', 3, 'CE', null, null, 'Observation Identifier 观察标识符');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (132, 2, 'OBX', 4, 'ST', null, null, 'Observation Sub -ID 观察子ID');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (133, 2, 'OBX', 5, 'TEXT', null, null, 'Observation Value 观察值');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (274, 4, 'MSH', 7, 'ST', null, null, 'Security 安全性');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (275, 4, 'MSH', 8, 'CM', 'ORM^O01', 'ORM^O01', 'Message Type 消息类型');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (276, 4, 'MSH', 9, 'ST', null, '[当前时间]', 'Message Control ID 消息控制ID');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (277, 4, 'MSH', 10, 'PT', 'P', 'P', 'Processing ID 处理ID');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (278, 4, 'MSH', 11, 'VID', '2.4', '2.4', 'Version ID 版本ID');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (279, 4, 'PID', 1, 'SI', null, '1', 'Set ID - Patient ID 设置ID - 患者ID');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (280, 4, 'PID', 2, 'CX', null, null, 'Patient ID 患者ID');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (281, 4, 'PID', 3, 'CX', null, '[标识号] ', 'Patient Identifier List 患者标识符表');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (282, 4, 'PID', 4, 'CX', null, '[病人ID]', 'Alternate Patient ID 备用患者ID');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (283, 4, 'PID', 5, 'XPN', null, '[姓名]', 'Patient Name 患者姓名');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (284, 4, 'PID', 6, 'XPN', null, null, 'Mother''s Maiden Name 母亲的婚前姓');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (285, 4, 'PID', 7, 'TS', null, '[出生日期]', 'Date/Time of Birth 出生日期');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (286, 4, 'PID', 8, 'IS', null, '[性别]', 'Sex 性别');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (287, 4, 'PID', 9, 'XPN', null, '[姓名]', 'Patient Alias 患者别名');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (288, 4, 'PID', 10, 'CE', null, 'C', 'Race 种族');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (289, 4, 'PID', 11, 'XAD', null, '[联系人地址]', 'Patient Address 患者住址');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (290, 4, 'PID', 12, 'IS', null, null, 'County Code 县代码');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (291, 4, 'PID', 13, 'XTN', null, '[家庭电话]', 'Phone Number - Home 家庭电话');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (292, 4, 'PID', 14, 'XTN', null, '[联系人电话]', 'Phone Number - Business 单位电话');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (293, 4, 'PID', 15, 'CE', null, null, 'Primary Language 母语');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (294, 4, 'PID', 16, 'IS', null, '[婚姻状况]', 'Marital Status 婚姻状况');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (295, 4, 'PID', 17, 'CE', null, null, 'Religion 宗教信仰');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (296, 4, 'PID', 18, 'CX', null, '[身份证号]', 'Patient Account Number患者账号');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (297, 4, 'PID', 19, 'ST', null, null, 'SSN Number - Patient 患者社会保险号码');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (298, 4, 'PV1', 1, 'SI', null, '1', 'Set ID-PV1设置ID-PV1');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (299, 4, 'PV1', 2, 'IS', null, '[病人来源]', 'Patient Class 患者类别');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (300, 4, 'PV1', 3, 'PL', null, '[当前科室名称]^[病区名称]^[床号]^^^^^^[当前科室名称]', 'Assigned Patient Location 指定患者位置');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (301, 4, 'PV1', 4, 'IS', null, null, 'Admission Type 入院类型');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (302, 4, 'PV1', 5, 'CX', null, null, 'Preadmit Number 预收入院号码');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (303, 4, 'PV1', 6, 'PL', null, null, 'Prior Patient Location 患者原位置');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (304, 4, 'PV1', 7, 'XCN', null, '[开嘱医生]', 'Attending Doctor 接诊医生');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (305, 4, 'PV1', 8, 'XCN', null, null, 'Referring Doctor 转诊医生');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (306, 4, 'PV1', 9, 'XCN', null, null, 'Consulting Doctor 会诊医生');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (307, 4, 'PV1', 10, 'IS', null, null, 'Hospital Service 医院服务');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (308, 4, 'PV1', 11, 'PL', null, null, 'Temporary Location 临时位置');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (309, 4, 'PV1', 12, 'IS', null, null, 'Preadmit Test Indicator 预收入院检验标识符');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (310, 4, 'PV1', 13, 'IS', null, null, 'Readmission Indicator 再次入院标识符');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (311, 4, 'PV1', 14, 'IS', null, null, 'Admit Source 入院来源');
commit;
prompt 300 records committed...
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (312, 4, 'PV1', 15, 'IS', null, null, 'Ambulatory Status （手术后）走动状态');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (313, 4, 'PV1', 16, 'IS', null, null, 'VIP Indicator VIP标识符');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (314, 4, 'PV1', 17, 'XCN', null, null, 'Admitting Doctor 入院医生');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (315, 4, 'PV1', 18, 'IS', null, null, 'Patient Type 患者类型');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (316, 4, 'PV1', 19, 'CX', null, '1', 'Visit Number 就诊号码');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (317, 4, 'PV1', 20, 'FC', null, null, 'Financial Class 经济状况类别');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (318, 4, 'PV1', 21, 'IS', null, null, 'Charge Price Indicator 费用价格标识符');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (319, 4, 'PV1', 22, 'IS', null, null, 'Courtesy Code 礼貌代码');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (320, 4, 'PV1', 23, 'IS', null, null, 'Credit Rating 信用等级');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (321, 4, 'PV1', 24, 'IS', null, null, 'Contract Code 合同代码');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (322, 4, 'PV1', 25, 'DT', null, null, 'Contract Effective Date 合同生效日期');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (323, 4, 'PV1', 26, 'NM', null, null, 'Contract Amount 合同总量');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (324, 4, 'PV1', 27, 'NM', null, null, 'Contract Period 合同期限');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (325, 4, 'PV1', 28, 'IS', null, null, 'Interest Code 利率代码');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (326, 4, 'PV1', 29, 'IS', null, null, 'Transfer to Bad Debt Code 转为坏账代码');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (327, 4, 'PV1', 30, 'DT', null, null, 'Transfer to Bad Debt Date 转为坏账日期');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (328, 4, 'PV1', 31, 'IS', null, null, 'Bad Debt Agency Code 坏账代理代码');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (329, 4, 'PV1', 32, 'NM', null, null, 'Bad Debt Transfer Amount 转为坏账总额');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (330, 4, 'PV1', 33, 'NM', null, null, 'Bad Debt Recovery Amount 坏账恢复总额');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (331, 4, 'PV1', 34, 'IS', null, null, 'Delete Account Indicator 删除账户标识符');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (332, 4, 'PV1', 35, 'DT', null, null, 'Delete Account Date 删除账户日期');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (333, 4, 'PV1', 36, 'IS', null, 'UNK', 'Discharge Disposition 出院处置');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (334, 4, 'PV1', 37, 'CM', null, null, 'Discharge to Location 出院去往的位置');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (335, 4, 'PV1', 38, 'CE', null, null, 'Diet Type 饮食类型');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (336, 4, 'PV1', 39, 'IS', null, null, 'Servicing Facility 服务机构');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (337, 4, 'PV1', 40, 'IS', null, null, 'Bed Status 床位状况');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (338, 4, 'PV1', 41, 'IS', null, null, 'Account Status 账户状况');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (339, 4, 'PV1', 42, 'IS', null, null, 'Pending Location 待定位置');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (340, 4, 'PV1', 43, 'PL', null, null, 'Prior Temporary Location 前临时位置');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (341, 4, 'PV1', 44, 'TS', null, '[入院时间]', 'Admit Date/Time 入院日期/时间');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (342, 4, 'PV1', 45, 'TS', null, '[出院时间]', 'Discharge Date/Time 出院日期/时间');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (343, 4, 'PV1', 46, 'NM', null, null, 'Current Patient Balance 当前患者差额');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (344, 4, 'PV1', 47, 'NM', null, null, 'Total Charges 总费用');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (345, 4, 'PV1', 48, 'NM', null, null, 'Total Adjustments 总调度数');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (346, 4, 'PV1', 49, 'NM', null, null, 'Total Payments 总支付额');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (347, 4, 'PV1', 50, 'CX', null, null, 'Alternate Visit ID 备用就诊ID');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (348, 4, 'ORC', 1, 'ID', null, '[消息类型]', '医嘱控制码');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (349, 4, 'ORC', 2, 'EI', null, '[医嘱ID]', '开单者医嘱号码');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (350, 4, 'ORC', 3, 'EI', null, null, '执行者医嘱号码');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (351, 4, 'ORC', 4, 'EI', null, null, '开单者医嘱组号码');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (352, 4, 'ORC', 5, 'ID', null, null, '医嘱状态');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (353, 4, 'ORC', 6, 'ID', null, null, '应答标记');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (354, 4, 'ORC', 7, 'TQ', null, null, '数量/时间');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (355, 4, 'ORC', 8, 'CM', null, null, '父层标识');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (356, 4, 'ORC', 9, 'TS', null, '[开嘱时间]', '事务日期/时间');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (357, 4, 'ORC', 10, 'XCN', null, '[开嘱医生]', '录入者');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (358, 4, 'ORC', 11, 'XCN', null, '[校对护士]', '校正者');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (359, 4, 'ORC', 12, 'XCN', null, '[开嘱医生]', '医嘱提供者');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (364, 4, 'OBR', 1, 'SI', null, '1', '设置ID-OBR');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (365, 4, 'OBR', 2, 'EI', null, null, '开单者医嘱号码');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (366, 4, 'OBR', 3, 'EI', null, null, '执行者医嘱号码');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (367, 4, 'OBR', 4, 'CE', null, '[医嘱内容]', '通用服务标识');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (368, 4, 'OBR', 5, 'ID', null, null, '优先级-OBR');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (369, 4, 'OBR', 6, 'TS', null, '[开嘱时间]', '请求日期/时间');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (370, 4, 'OBR', 7, 'TS', null, '[开嘱时间]', '观察日期/时间#');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (371, 4, 'OBR', 8, 'TS', null, null, '观察结束日期/时间#');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (372, 4, 'OBR', 9, 'CQ', null, null, '样本收集量*');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (373, 4, 'OBR', 10, 'XCN', null, null, '收集者标识*');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (374, 4, 'OBR', 11, 'ID', null, null, '样本行为码*');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (375, 4, 'OBR', 12, 'CE', null, null, '危险因素代码');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (376, 4, 'OBR', 13, 'ST', null, null, '相关临床信息');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (377, 4, 'OBR', 14, 'TS', null, null, '收到样本日期/时间*');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (378, 4, 'OBR', 15, 'CM', null, null, '样本来源');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (379, 4, 'OBR', 16, 'XCN', null, '[开嘱医生]', '医嘱提供者');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (380, 4, 'OBR', 17, 'XTN', null, null, '医嘱回访电话号码');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (381, 4, 'OBR', 18, 'ST', null, null, '开单者字段1');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (382, 4, 'OBR', 19, 'ST', null, null, '开单者字段2');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (383, 4, 'OBR', 20, 'ST', null, null, '执行者字段1+');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (384, 4, 'OBR', 21, 'ST', null, null, '执行者字段2+');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (385, 4, 'OBR', 22, 'TS', null, null, '结果报告/状态改变日期/时间+');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (386, 4, 'OBR', 23, 'CM', null, null, '收费+');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (387, 4, 'OBR', 24, 'ID', null, null, '诊断服务标识ID');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (388, 4, 'OBR', 25, 'ID', null, null, '结果状态+');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (389, 4, 'OBR', 26, 'CM', null, null, '父层结果+');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (390, 4, 'OBR', 27, 'TQ', null, null, '数量/时间');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (391, 4, 'OBR', 28, 'XCN', null, null, '结果需要者');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (392, 4, 'OBR', 29, 'CM', null, null, '父层连结码');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (393, 4, 'OBR', 30, 'ID', null, null, '患者行动方式');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (394, 4, 'OBR', 31, 'CE', null, null, '研究原因');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (395, 4, 'OBX', 1, 'SI', null, null, 'Set ID -OBX 设置ID -OBX');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (396, 4, 'OBX', 2, 'ID', null, null, 'Value Type 值类型');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (397, 4, 'OBX', 3, 'CE', null, null, 'Observation Identifier 观察标识符');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (398, 4, 'OBX', 4, 'ST', null, null, 'Observation Sub -ID 观察子ID');
insert into HL7消息段配置 (ID, 消息ID, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称)
values (399, 4, 'OBX', 5, 'TEXT', null, null, 'Observation Value 观察值');
commit;
prompt 384 records loaded
set feedback on
set define on
prompt Done.
