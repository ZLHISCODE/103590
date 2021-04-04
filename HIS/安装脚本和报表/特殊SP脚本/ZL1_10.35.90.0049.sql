----------------------------------------------------------------------------------------------------------------
--本脚本支持从ZLHIS+ v10.35.90升级到 v10.35.90
--请以数据空间的所有者登录PLSQL并执行下列脚本
Define n_System=100;
----------------------------------------------------------------------------------------------------------------
----------------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------
--结构修正部份
------------------------------------------------------------------------------



------------------------------------------------------------------------------
--数据修正部份
------------------------------------------------------------------------------


-------------------------------------------------------------------------------
--权限修正部份
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--报表修正部份
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--过程修正部份
-------------------------------------------------------------------------------
--137701:胡俊勇,2019-02-12,病人信息合并消息锚点修改
Create Or Replace Package b_Message Is
  Procedure p_Msg_Todo_Insert
  (
    Msg_Code_In  Zlmsg_Todo.Msg_Code%Type,
    Key_Value_In Zlmsg_Todo.Key_Value%Type
  );
  --设置平台调用类型
  Procedure Set_Platform_Call(Platform_Call Number);
  --新增部门
  Procedure Zlhis_Dict_001(Id_In 部门表.Id%Type);
  --修改部门
  Procedure Zlhis_Dict_002(部门id_In 部门表.Id%Type);
  --停用部门
  Procedure Zlhis_Dict_003(部门id_In 部门表.Id%Type);
  --启用部门
  Procedure Zlhis_Dict_004(部门id_In 部门表.Id%Type);
  --新增人员
  Procedure Zlhis_Dict_005(人员id_In 人员表.Id%Type);
  --修改人员
  Procedure Zlhis_Dict_006(人员id_In 人员表.Id%Type);
  --停用人员
  Procedure Zlhis_Dict_007(人员id_In 人员表.Id%Type);
  --启用人员
  Procedure Zlhis_Dict_008(人员id_In 人员表.Id%Type);
  --新增收费项目
  Procedure Zlhis_Dict_009(细目id_In 收费项目目录.Id%Type);
  --修改收费项目
  Procedure Zlhis_Dict_010(细目id_In 收费项目目录.Id%Type);
  --停用收费项目
  Procedure Zlhis_Dict_011(细目id_In 收费项目目录.Id%Type);
  --启用收费项目
  Procedure Zlhis_Dict_012(细目id_In 收费项目目录.Id%Type);
  --新增诊疗项目
  Procedure Zlhis_Dict_013(诊疗id_In 诊疗项目目录.Id%Type);
  --修改诊疗项目
  Procedure Zlhis_Dict_014(诊疗id_In 诊疗项目目录.Id%Type);
  --停用诊疗项目
  Procedure Zlhis_Dict_015(诊疗id_In 诊疗项目目录.Id%Type);
  --启用诊疗项目
  Procedure Zlhis_Dict_016(诊疗id_In 诊疗项目目录.Id%Type);
  --新增检验项目
  Procedure Zlhis_Dict_017(诊疗id_In 诊疗项目目录.Id%Type);
  --修改检验项目
  Procedure Zlhis_Dict_018(诊疗id_In 诊疗项目目录.Id%Type);
  --删除检验项目
  Procedure Zlhis_Dict_019
  (
    诊疗id_In 诊疗项目目录.Id%Type,
    编码_In   诊治所见项目.编码%Type,
    中文名_In 诊治所见项目.中文名%Type,
    英文名_In 诊治所见项目.英文名%Type
  );

  --新增疾病编码目录
  Procedure Zlhis_Dict_021(疾病id_In 疾病编码目录.Id%Type);
  --修改疾病编码目录
  Procedure Zlhis_Dict_022(疾病id_In 疾病编码目录.Id%Type);
  --停用疾病编码目录
  Procedure Zlhis_Dict_023(疾病id_In 疾病编码目录.Id%Type);
  --启用疾病编码目录
  Procedure Zlhis_Dict_024(疾病id_In 疾病编码目录.Id%Type);
  --新增药品分类
  Procedure Zlhis_Dict_025
  (
    类型_In 诊疗分类目录.类型%Type,
    Id_In   诊疗分类目录.Id%Type
  );
  --修改药品分类
  Procedure Zlhis_Dict_026
  (
    类型_In 诊疗分类目录.类型%Type,
    Id_In   诊疗分类目录.Id%Type
  );
  --删除药品分类
  Procedure Zlhis_Dict_027
  (
    类型_In 诊疗分类目录.类型%Type,
    Id_In   诊疗分类目录.Id%Type,
    编码_In 诊疗分类目录.编码%Type,
    名称_In 诊疗分类目录.名称%Type
  );
  --停用药品分类
  Procedure Zlhis_Dict_028
  (
    类型_In 诊疗分类目录.类型%Type,
    Id_In   诊疗分类目录.Id%Type
  );
  --启用药品分类
  Procedure Zlhis_Dict_029
  (
    类型_In 诊疗分类目录.类型%Type,
    Id_In   诊疗分类目录.Id%Type
  );
  --新增药品品种
  Procedure Zlhis_Dict_030
  (
    类别_In 诊疗项目目录.类别%Type,
    Id_In   诊疗项目目录.Id%Type
  );
  --修改药品品种
  Procedure Zlhis_Dict_031
  (
    类别_In 诊疗项目目录.类别%Type,
    Id_In   诊疗项目目录.Id%Type
  );
  --删除药品品种
  Procedure Zlhis_Dict_032
  (
    类别_In 诊疗项目目录.类别%Type,
    Id_In   诊疗项目目录.Id%Type,
    编码_In 诊疗项目目录.编码%Type,
    名称_In 诊疗项目目录.名称%Type
  );
  --停用药品品种
  Procedure Zlhis_Dict_033
  (
    类别_In 诊疗项目目录.类别%Type,
    Id_In   诊疗项目目录.Id%Type
  );
  --启用药品品种
  Procedure Zlhis_Dict_034
  (
    类别_In 诊疗项目目录.类别%Type,
    Id_In   诊疗项目目录.Id%Type
  );
  --新增药品规格
  Procedure Zlhis_Dict_035
  (
    类别_In   收费项目目录.类别%Type,
    药品id_In 收费项目目录.Id%Type
  );
  --修改药品规格
  Procedure Zlhis_Dict_036
  (
    类别_In   收费项目目录.类别%Type,
    药品id_In 收费项目目录.Id%Type
  );
  --删除药品规格
  Procedure Zlhis_Dict_037
  (
    类别_In   收费项目目录.类别%Type,
    药品id_In 收费项目目录.Id%Type,
    编码_In   收费项目目录.编码%Type,
    名称_In   收费项目目录.名称%Type,
    规格_In   收费项目目录.规格%Type,
    产地_In   收费项目目录.产地%Type
  );
  --停用药品规格
  Procedure Zlhis_Dict_038
  (
    类别_In   收费项目目录.类别%Type,
    药品id_In 收费项目目录.Id%Type
  );
  --启用药品规格
  Procedure Zlhis_Dict_039
  (
    类别_In   收费项目目录.类别%Type,
    药品id_In 收费项目目录.Id%Type
  );
  --设置药品存储库房
  Procedure Zlhis_Dict_040
  (
    类别_In   收费项目目录.类别%Type,
    药品id_In 收费项目目录.Id%Type
  );
  --设置药品储备限额
  Procedure Zlhis_Dict_041
  (
    类别_In   收费项目目录.类别%Type,
    药品id_In 收费项目目录.Id%Type
  );
  --新增卫材品种
  Procedure Zlhis_Dict_042(Id_In 诊疗项目目录.Id%Type);
  --新增卫材规格
  Procedure Zlhis_Dict_043(Id_In 收费项目目录.Id%Type);
  --修改卫材规格
  Procedure Zlhis_Dict_044(Id_In 收费项目目录.Id%Type);
  --删除卫材规格
  Procedure Zlhis_Dict_045
  (
    Id_In   收费项目目录.Id%Type,
    编码_In 收费项目目录.编码%Type,
    名称_In 收费项目目录.名称%Type,
    规格_In 收费项目目录.规格%Type,
    产地_In 收费项目目录.产地%Type
  );
  --停用卫材规格
  Procedure Zlhis_Dict_046(Id_In 收费项目目录.Id%Type);
  --启用卫材规格
  Procedure Zlhis_Dict_047(Id_In 收费项目目录.Id%Type);
  --医保对码
  Procedure Zlhis_Dict_048
  (
    险类_In       In 保险支付项目.险类%Type,
    收费细目id_In In 保险支付项目.收费细目id%Type
  );
  --删除医保对码
  Procedure Zlhis_Dict_049
  (
    险类_In       In 保险支付项目.险类%Type,
    收费细目id_In In 保险支付项目.收费细目id%Type,
    项目编码_In   In 收费项目目录.编码%Type,
    项目名称_In   In 收费项目目录.名称%Type,
    医保编码_In   In 保险支付项目.项目编码%Type,
    医保名称_In   In 保险支付项目.项目名称%Type
  );
  --新增卫材分类
  Procedure Zlhis_Dict_050
  (
    类型_In 诊疗分类目录.类型%Type,
    Id_In   诊疗分类目录.Id%Type
  );
  --修改卫材分类
  Procedure Zlhis_Dict_051
  (
    类型_In 诊疗分类目录.类型%Type,
    Id_In   诊疗分类目录.Id%Type
  );
  --删除卫材分类
  Procedure Zlhis_Dict_052
  (
    类型_In 诊疗分类目录.类型%Type,
    Id_In   诊疗分类目录.Id%Type,
    编码_In 诊疗分类目录.编码%Type,
    名称_In 诊疗分类目录.名称%Type
  );
  --收费价目变动
  Procedure Zlhis_Dict_053
  (
    收费项目Id_In       收费项目目录.Id%Type
  );
  --诊疗收费对照变动
  Procedure Zlhis_Dict_054
  (
    诊疗项目Id_In     诊疗分类目录.Id%Type
  );
  --新增诊疗检查类型
  Procedure Zlhis_Dictpacs_001
  (
    编码_In   诊疗检查类型.编码%Type,
    名称_In   诊疗检查类型.名称%Type,
    简码_In   诊疗检查类型.简码%Type,
    建病案_In 诊疗检查类型.建病案%Type
  );
  --修改诊疗检查类型
  Procedure Zlhis_Dictpacs_002
  (
    编码_In   诊疗检查类型.编码%Type,
    名称_In   诊疗检查类型.名称%Type,
    简码_In   诊疗检查类型.简码%Type,
    建病案_In 诊疗检查类型.建病案%Type
  );
  --删除诊疗检查类型
  Procedure Zlhis_Dictpacs_003
  (
    编码_In   诊疗检查类型.编码%Type,
    名称_In   诊疗检查类型.名称%Type,
    简码_In   诊疗检查类型.简码%Type,
    建病案_In 诊疗检查类型.建病案%Type
  );
  --新增诊疗检查部位
  Procedure Zlhis_Dictpacs_004
  (
    类型_In     诊疗检查部位.类型%Type,
    编码_In     诊疗检查部位.编码%Type,
    名称_In     诊疗检查部位.名称%Type,
    分组_In     诊疗检查部位.分组%Type,
    备注_In     诊疗检查部位.备注%Type,
    方法_In     诊疗检查部位.方法%Type,
    适用性别_In 诊疗检查部位.适用性别%Type
  );
  --修改诊疗检查部位
  Procedure Zlhis_Dictpacs_005
  (
    类型_In     诊疗检查部位.类型%Type,
    编码_In     诊疗检查部位.编码%Type,
    名称_In     诊疗检查部位.名称%Type,
    分组_In     诊疗检查部位.分组%Type,
    备注_In     诊疗检查部位.备注%Type,
    方法_In     诊疗检查部位.方法%Type,
    适用性别_In 诊疗检查部位.适用性别%Type
  );
  --删除诊疗检查部位
  Procedure Zlhis_Dictpacs_006
  (
    类型_In     诊疗检查部位.类型%Type,
    编码_In     诊疗检查部位.编码%Type,
    名称_In     诊疗检查部位.名称%Type,
    分组_In     诊疗检查部位.分组%Type,
    备注_In     诊疗检查部位.备注%Type,
    方法_In     诊疗检查部位.方法%Type,
    适用性别_In 诊疗检查部位.适用性别%Type
  );
  --新增诊疗项目部位
  Procedure Zlhis_Dictpacs_007
  (
    Id_In     诊疗项目部位.Id%Type,
    项目id_In 诊疗项目部位.项目id%Type,
    类型_In   诊疗项目部位.类型%Type,
    部位_In   诊疗项目部位.部位%Type,
    方法_In   诊疗项目部位.方法%Type,
    默认_In   诊疗项目部位.默认%Type
  );
  --修改诊疗项目部位
  Procedure Zlhis_Dictpacs_008
  (
    Id_In     诊疗项目部位.Id%Type,
    项目id_In 诊疗项目部位.项目id%Type,
    类型_In   诊疗项目部位.类型%Type,
    部位_In   诊疗项目部位.部位%Type,
    方法_In   诊疗项目部位.方法%Type,
    默认_In   诊疗项目部位.默认%Type
  );
  --删除诊疗项目部位
  Procedure Zlhis_Dictpacs_009
  (
    Id_In     诊疗项目部位.Id%Type,
    项目id_In 诊疗项目部位.项目id%Type,
    类型_In   诊疗项目部位.类型%Type,
    部位_In   诊疗项目部位.部位%Type,
    方法_In   诊疗项目部位.方法%Type,
    默认_In   诊疗项目部位.默认%Type
  );
  --新增诊疗检验标本
  Procedure Zlhis_Dictlis_004
  (
    编码_In     诊疗检验标本.编码%Type,
    名称_In     诊疗检验标本.名称%Type,
    简码_In     诊疗检验标本.简码%Type,
    适用性别_In 诊疗检验标本.适用性别%Type
  );
  --修改诊疗检验标本
  Procedure Zlhis_Dictlis_005
  (
    编码_In     诊疗检验标本.编码%Type,
    名称_In     诊疗检验标本.名称%Type,
    简码_In     诊疗检验标本.简码%Type,
    适用性别_In 诊疗检验标本.适用性别%Type
  );
  --删除诊疗项目部位
  Procedure Zlhis_Dictlis_006
  (
    编码_In     诊疗检验标本.编码%Type,
    名称_In     诊疗检验标本.名称%Type,
    简码_In     诊疗检验标本.简码%Type,
    适用性别_In 诊疗检验标本.适用性别%Type
  );
  --新增采血管类型
  Procedure Zlhis_Dictlis_007
  (
    编码_In   采血管类型.编码%Type,
    名称_In   采血管类型.名称%Type,
    简码_In   采血管类型.简码%Type,
    添加剂_In 采血管类型.添加剂%Type,
    采血量_In 采血管类型.采血量%Type,
    规格_In   采血管类型.规格%Type,
    颜色_In   采血管类型.颜色%Type,
    材料id_In 采血管类型.材料id%Type
  );
  --修改采血管类型
  Procedure Zlhis_Dictlis_008
  (
    编码_In   采血管类型.编码%Type,
    名称_In   采血管类型.名称%Type,
    简码_In   采血管类型.简码%Type,
    添加剂_In 采血管类型.添加剂%Type,
    采血量_In 采血管类型.采血量%Type,
    规格_In   采血管类型.规格%Type,
    颜色_In   采血管类型.颜色%Type,
    材料id_In 采血管类型.材料id%Type
  );
  --删除采血管类型
  Procedure Zlhis_Dictlis_009
  (
    编码_In   采血管类型.编码%Type,
    名称_In   采血管类型.名称%Type,
    简码_In   采血管类型.简码%Type,
    添加剂_In 采血管类型.添加剂%Type,
    采血量_In 采血管类型.采血量%Type,
    规格_In   采血管类型.规格%Type,
    颜色_In   采血管类型.颜色%Type,
    材料id_In 采血管类型.材料id%Type
  );

  --药品备药发送
  Procedure Zlhis_Drug_001(No_In 药品收发记录.No%Type);
  --取消药品备药发送
  Procedure Zlhis_Drug_002(No_In 药品收发记录.No%Type);
  --药品移库单接收
  Procedure Zlhis_Drug_003(No_In 药品收发记录.No%Type);
  --药品移库单冲销
  Procedure Zlhis_Drug_004(No_In 药品收发记录.No%Type);

  --部门发药
  Procedure Zlhis_Drug_005
  (
    库房id_In 药品收发记录.库房id%Type,
    收发id_In 药品收发记录.Id%Type
  );
  --部门退药
  Procedure Zlhis_Drug_006
  (
    冲销收发id_In 药品收发记录.Id%Type,
    待发收发id_In 药品收发记录.Id%Type,
    数量_In       药品收发记录.实际数量%Type,
    费用id_In     门诊费用记录.Id%Type
  );
  --药品调价
  Procedure Zlhis_Drug_007(价格id_In 药品价格记录.Id%Type);
  --静配发送
  Procedure Zlhis_Drug_008(记录ids_In Varchar2);
  --药品调售价
  Procedure Zlhis_Drug_009
  (
    价格id_In 药品价格记录.Id%Type,
    时价_In   Number
  );
  --卫材调成本价
  Procedure Zlhis_Drug_010(价格id_In 成本价调价信息.Id%Type);
  --卫材调售价
  Procedure Zlhis_Drug_011
  (
    价格id_In 收费价目.Id%Type,
    时价_In   Number
  );
  --2.停止患者医嘱，住院
  Procedure Zlhis_Cis_002
  (
    病人id_In  In 病人医嘱记录.病人id%Type,
    主页id_In  In 病人医嘱记录.主页id%Type,
    医嘱id_In  In 病人医嘱记录.Id%Type,
    医嘱ids_In In Varchar2
  );
  --3.作废患者医嘱，门诊/住院
  Procedure Zlhis_Cis_003
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    挂号单_In In 病人挂号记录.No%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  );

  --4.患者术后医嘱，住院
  Procedure Zlhis_Cis_004
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  );

  --5.撤消患者术后医嘱，住院
  Procedure Zlhis_Cis_005
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  );

  --6.患者护理常规医嘱，住院
  Procedure Zlhis_Cis_006
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  );

  --7.撤消患者护理常规医嘱，住院
  Procedure Zlhis_Cis_007
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type,
    No_In     In 病人医嘱发送.No%Type
  );

  --门诊患者接诊
  Procedure Zlhis_Cis_008
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    挂号单_In In 病人挂号记录.No%Type
  );

  --门诊患者取消接诊
  Procedure Zlhis_Cis_009
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    挂号单_In In 病人挂号记录.No%Type
  );

  --10.下达患者诊断，门诊/住院
  Procedure Zlhis_Cis_010
  (
    病人id_In In 病人挂号记录.病人id%Type,
    就诊id_In In 病人挂号记录.Id%Type, --门诊病人 挂号ID，住院病人 主页ID
    诊断id_In In 病人诊断记录.Id%Type
  );
  --11.撤消患者诊断
  Procedure Zlhis_Cis_011
  (
    病人id_In   In 病人挂号记录.病人id%Type,
    就诊id_In   In 病人挂号记录.Id%Type, --门诊病人 挂号ID，住院病人 主页ID
    Id_In       In 病人诊断记录.Id%Type,
    疾病id_In   In 病人诊断记录.疾病id%Type,
    诊断id_In   In 病人诊断记录.诊断id%Type,
    诊断描述_In In 病人诊断记录.诊断描述%Type
  );

  --病区执行医嘱校对
  Procedure Zlhis_Cis_012
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  );

  --13.检验危急值阅读通知
  Procedure Zlhis_Cis_014
  (
    病人id_In In 病人挂号记录.病人id%Type,
    就诊id_In In 病人挂号记录.Id%Type, --门诊病人 挂号ID，住院病人 主页ID
    医嘱id_In In 病人医嘱记录.Id%Type,
    消息id_In In 业务消息清单.Id%Type
  );

  --15.患者检验申请，门诊/住院
  Procedure Zlhis_Cis_016
  (
    病人id_In   In 病人医嘱记录.病人id%Type,
    主页id_In   In 病人医嘱记录.主页id%Type,
    挂号单_In   In 病人挂号记录.No%Type,
    发送号_In   In 病人医嘱发送.发送号%Type,
    医嘱id_In   In 病人医嘱记录.Id%Type,
    病人来源_In In 病人医嘱记录.病人来源%Type
  );
  --16.患者检查申请，门诊/住院
  Procedure Zlhis_Cis_017
  (
    病人id_In   In 病人医嘱记录.病人id%Type,
    主页id_In   In 病人医嘱记录.主页id%Type,
    挂号单_In   In 病人挂号记录.No%Type,
    发送号_In   In 病人医嘱发送.发送号%Type,
    医嘱id_In   In 病人医嘱记录.Id%Type,
    病人来源_In In 病人医嘱记录.病人来源%Type
  );
  --17.患者手术申请，门诊/住院
  Procedure Zlhis_Cis_018
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    挂号单_In In 病人挂号记录.No%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  );
  --18.患者输血申请，住院
  Procedure Zlhis_Cis_019
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    挂号单_In In 病人挂号记录.No%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  );
  --19.患者会诊申请，住院
  Procedure Zlhis_Cis_020
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  );
  --20.患者抢救医嘱，住院
  Procedure Zlhis_Cis_021
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  );
  --21.患者死亡医嘱，住院
  Procedure Zlhis_Cis_022
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  );
  --22.患者特殊治疗医嘱，住院
  Procedure Zlhis_Cis_023
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  );

  --24.检查危急值阅读通知
  Procedure Zlhis_Cis_025
  (
    病人id_In In 病人挂号记录.病人id%Type,
    就诊id_In In 病人挂号记录.Id%Type, --门诊病人 挂号ID，住院病人 主页ID
    医嘱id_In In 病人医嘱记录.Id%Type,
    消息id_In In 业务消息清单.Id%Type
  );

  --病区执行医嘱发送
  Procedure Zlhis_Cis_026
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  );

  --撤消患者检验申请
  Procedure Zlhis_Cis_036
  (
    病人id_In   In 病人医嘱记录.病人id%Type,
    主页id_In   In 病人医嘱记录.主页id%Type,
    挂号单_In   In 病人挂号记录.No%Type,
    发送号_In   In 病人医嘱发送.发送号%Type,
    医嘱id_In   In 病人医嘱记录.Id%Type,
    No_In       In 病人医嘱发送.No%Type,
    病人来源_In In 病人医嘱记录.病人来源%Type
  );

  --撤消患者检查申请
  Procedure Zlhis_Cis_037
  (
    病人id_In   In 病人医嘱记录.病人id%Type,
    主页id_In   In 病人医嘱记录.主页id%Type,
    挂号单_In   In 病人挂号记录.No%Type,
    发送号_In   In 病人医嘱发送.发送号%Type,
    医嘱id_In   In 病人医嘱记录.Id%Type,
    No_In       In 病人医嘱发送.No%Type,
    病人来源_In In 病人医嘱记录.病人来源%Type
  );

  --撤消患者手术申请
  Procedure Zlhis_Cis_038
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    挂号单_In In 病人挂号记录.No%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type,
    No_In     In 病人医嘱发送.No%Type
  );

  --撤消患者输血申请
  Procedure Zlhis_Cis_039
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    挂号单_In In 病人挂号记录.No%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type,
    No_In     In 病人医嘱发送.No%Type
  );

  --撤消患者会诊申请
  Procedure Zlhis_Cis_040
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type,
    No_In     In 病人医嘱发送.No%Type
  );

  --撤消患者抢救医嘱
  Procedure Zlhis_Cis_041
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type,
    No_In     In 病人医嘱发送.No%Type
  );

  --撤消患者死亡医嘱
  Procedure Zlhis_Cis_042
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type,
    No_In     In 病人医嘱发送.No%Type
  );

  --撤消特殊治疗医嘱
  Procedure Zlhis_Cis_043
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type,
    No_In     In 病人医嘱发送.No%Type
  );

  --撤消病区执行医嘱
  Procedure Zlhis_Cis_044
  (
    病人id_In   In 病人医嘱记录.病人id%Type,
    主页id_In   In 病人医嘱记录.主页id%Type,
    发送号_In   In 病人医嘱发送.发送号%Type,
    医嘱id_In   In 病人医嘱记录.Id%Type,
    No_In       In 病人医嘱发送.No%Type,
    发送数次_In In 病人医嘱发送.发送数次%Type,
    首次时间_In In 病人医嘱发送.首次时间%Type,
    末次时间_In In 病人医嘱发送.末次时间%Type,
    样本条码_In In 病人医嘱发送.样本条码%Type
  );
  --患者医嘱执行登记
  Procedure Zlhis_Cis_050
  (
    病人id_In   In 病人医嘱记录.病人id%Type,
    主页id_In   In 病人医嘱记录.主页id%Type,
    挂号单_In   In 病人挂号记录.No%Type,
    发送号_In   In 病人医嘱发送.发送号%Type,
    医嘱id_In   In 病人医嘱记录.Id%Type,
    要求时间_In In 病人医嘱执行.要求时间%Type,
    执行时间_In In 病人医嘱执行.执行时间%Type
  );

  --患者医嘱取消执行登记
  Procedure Zlhis_Cis_051
  (
    病人id_In   In 病人医嘱记录.病人id%Type,
    主页id_In   In 病人医嘱记录.主页id%Type,
    挂号单_In   In 病人挂号记录.No%Type,
    发送号_In   In 病人医嘱发送.发送号%Type,
    医嘱id_In   In 病人医嘱记录.Id%Type,
    要求时间_In In 病人医嘱执行.要求时间%Type,
    执行时间_In In 病人医嘱执行.执行时间%Type,
    本次数次_In In 病人医嘱执行.本次数次%Type,
    执行结果_In In 病人医嘱执行.执行结果%Type,
    执行摘要_In In 病人医嘱执行.执行摘要%Type,
    执行科室_In In 病人医嘱执行.执行科室id%Type,
    执行人_In   In 病人医嘱执行.执行人%Type,
    核对人_In   In 病人医嘱执行.核对人%Type,
    记录来源_In In 病人医嘱执行.记录来源%Type
  );
  --患者医嘱执行完成
  Procedure Zlhis_Cis_052
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    挂号单_In In 病人挂号记录.No%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  );
  --患者医嘱撤消执行完成
  Procedure Zlhis_Cis_053
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    挂号单_In In 病人挂号记录.No%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  );
  --病理申请发送后修改
  Procedure Zlhis_Cis_056
  (
    病人id_In   In 病人医嘱记录.病人id%Type,
    主页id_In   In 病人医嘱记录.主页id%Type,
    挂号单_In   In 病人挂号记录.No%Type,
    发送号_In   In 病人医嘱发送.发送号%Type,
    医嘱id_In   In 病人医嘱记录.Id%Type,
    病人来源_In In 病人医嘱记录.病人来源%Type
  );

  --门诊患者完成就诊
  Procedure Zlhis_Cis_057
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    挂号单_In In 病人挂号记录.No%Type
  );

  --门诊患者取消完成就诊
  Procedure Zlhis_Cis_058
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    挂号单_In In 病人挂号记录.No%Type
  );

  --确认停止患者医嘱 
  Procedure Zlhis_Cis_059
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  );

  --26.检查报告完成，检查完成时
  Procedure Zlhis_Pacs_001
  (
    医嘱id_In   In 影像检查记录.医嘱id%Type,
    报告id_Ins  In Varchar2,
    报告类型_In In Number --1-老版PACS报告，2-老版病历编辑器报告，3-新版编辑器报告
  );
  --27.检查状态同步，检查状态改变后
  Procedure Zlhis_Pacs_002
  (
    医嘱id_In In 影像检查记录.医嘱id%Type,
    原状态_In In 病人医嘱发送.执行过程%Type,
    新状态_In In 病人医嘱发送.执行过程%Type
  );
  --28.检查状态回退，检查状态回退后
  Procedure Zlhis_Pacs_003
  (
    医嘱id_In In 影像检查记录.医嘱id%Type,
    原状态_In In 病人医嘱发送.执行过程%Type,
    新状态_In In 病人医嘱发送.执行过程%Type
  );
  --29.检查报告撤销，撤销检查完成时
  Procedure Zlhis_Pacs_004
  (
    医嘱id_In   In 影像检查记录.医嘱id%Type,
    报告id_Ins  In Varchar2,
    报告类型_In In Number --1-老版PACS报告，2-老版病历编辑器报告，3-新版编辑器报告
  );
  --30.检查危急值通知，检查发生危急值时
  Procedure Zlhis_Pacs_005(医嘱id_In In 影像检查记录.医嘱id%Type);
  -- 检查预约通知，检查预约时
  Procedure Zlhis_Pacs_006
  (
    医嘱id_In In 影像检查记录.医嘱id%Type,
    预约id_In In Ris检查预约.预约id%Type
  );
  -- 取消检查预约，取消预约时
  Procedure Zlhis_Pacs_007
  (
    医嘱id_In       In 影像检查记录.医嘱id%Type,
    预约id_In       In Ris检查预约.预约id%Type,
    预约日期_In     In Ris检查预约.预约日期%Type,
    预约序号_In     In Ris检查预约.序号%Type,
    检查设备名称_In In Ris检查预约.检查设备名称%Type
  );

  --36.患者发卡
  Procedure Zlhis_Patient_018
  (
    变动id_In   In 病人变动记录.Id%Type,
    病人id_In   In 病人信息.病人id%Type,
    卡类别id_In In 医疗卡类别.Id%Type,
    卡号_In     In 病人医疗卡信息.卡号%Type
  );

  --37.患者退卡
  Procedure Zlhis_Patient_019
  (
    变动id_In   In 病人变动记录.Id%Type,
    病人id_In   In 病人信息.病人id%Type,
    卡类别id_In In 医疗卡类别.Id%Type,
    卡号_In     In 病人医疗卡信息.卡号%Type
  );

  --38.患者退卡
  Procedure Zlhis_Patient_020
  (
    变动id_In   In 病人变动记录.Id%Type,
    病人id_In   In 病人信息.病人id%Type,
    卡类别id_In In 医疗卡类别.Id%Type,
    原卡号_In   In 病人医疗卡信息.卡号%Type,
    新卡号_In   In 病人医疗卡信息.卡号%Type
  );

  --39.病人挂号登记（包含预约登记)
  Procedure Zlhis_Regist_001
  (
    挂号id_In In 病人挂号记录.Id%Type,
    No_In     In 病人挂号记录.No%Type
  );

  --40.病人分诊
  Procedure Zlhis_Regist_002
  (
    挂号id_In In 病人挂号记录.Id%Type,
    No_In     In 病人挂号记录.No%Type,
    诊室_In   In 病人挂号记录.诊室%Type
  );

  --41.病人退号
  Procedure Zlhis_Regist_003
  (
    挂号id_In In 病人挂号记录.Id%Type,
    No_In     In 病人挂号记录.No%Type
  );

  --42.临床出诊安排调整
  Procedure Zlhis_Regist_004
  (
    变动原因_In In Integer, --1-停诊;2-替诊;3-诊室变动
    记录id_In   In 临床出诊记录.Id%Type,
    变动id_In   In 临床出诊变动记录.Id%Type
  );

  --43.门诊患者挂号换号操作
  Procedure Zlhis_Regist_005
  (
    No_In         In 病人挂号记录.No%Type,
    变动原因_In   Integer, --1-替诊;2-换号;3-预约日期变动,
    就诊变动id_In 就诊变动记录.Id%Type
  );

  --费用门诊收费及补充结算
  --结算类型_In:1-收费结算，2-补充结算
  Procedure Zlhis_Charge_002
  (
    结算类型_In In Number,
    结帐id_In   In 门诊费用记录.结帐id%Type
  );

  --46.门诊退费单据
  Procedure Zlhis_Charge_004
  (
    退费类型_In In Number,
    结帐id_In   In 门诊费用记录.结帐id%Type
  );

  --47.收预交款
  Procedure Zlhis_Charge_005
  (
    预交id_In In 病人预交记录.Id%Type,
    单据号_In In 病人预交记录.No%Type
  );

  --48.退预交款(包含负数退预交款部分)
  Procedure Zlhis_Charge_006
  (
    退预交id_In In 病人预交记录.Id%Type,
    单据号_In   In 病人预交记录.No%Type
  );

  --住院记帐单据
  Procedure Zlhis_Charge_007
  (
    收费类别_In In 住院费用记录.收费类别%Type,
    费用id_In   In 住院费用记录.Id%Type
  );

  --住院记帐单据销账
  Procedure Zlhis_Charge_008
  (
    收费类别_In In 住院费用记录.收费类别%Type,
    费用id_In   In 住院费用记录.Id%Type,
    收发ids_In  In Varchar2 := Null --可能费用ID对应多个收发id，对应格式：收发id,数量|收发id,数量；非药品不传
  );

  --53.住院患者入院登记
  Procedure Zlhis_Patient_001
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  );
  --54.住院患者入院入科
  Procedure Zlhis_Patient_002
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  );
  --56.住院患者床位变更
  Procedure Zlhis_Patient_004
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  );
  --57.住院患者病情变更
  Procedure Zlhis_Patient_005
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  );
  --58.住院患者变更撤消
  Procedure Zlhis_Patient_006
  (
    病人id_In   In 病案主页.病人id%Type,
    主页id_In   In 病案主页.主页id%Type,
    撤销方式_In In Varchar2
  );
  --59.住院患者医护变更
  Procedure Zlhis_Patient_007
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  );
  --住院患者护理等级变更
  Procedure Zlhis_Patient_008
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  );
  --60.住院患者预出院
  Procedure Zlhis_Patient_009
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  );
  --61.住院患者出院
  Procedure Zlhis_Patient_010
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  );
  --62.住院患者新生儿登记
  Procedure Zlhis_Patient_011
  (
    病人id_In   In 病案主页.病人id%Type,
    主页id_In   In 病案主页.主页id%Type,
    婴儿序号_In 病人医嘱记录.婴儿%Type
  );
  --63.住院患者转入科室
  Procedure Zlhis_Patient_012
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  );
  --64.新生儿登记作废
  Procedure Zlhis_Patient_013
  (
    病人id_In   In 病案主页.病人id%Type,
    主页id_In   In 病案主页.主页id%Type,
    婴儿序号_In 病人医嘱记录.婴儿%Type
  );
  --65.门诊患者登记
  Procedure Zlhis_Patient_015(病人id_In In 病案主页.病人id%Type);
  --66.患者信息修改
  Procedure Zlhis_Patient_016(病人id_In In 病案主页.病人id%Type);

  --67.患者合并
  Procedure Zlhis_Patient_017 
  ( 
    病人id_In   In 病案主页.病人id%Type, 
    原病人id_In In 病案主页.病人id%Type,
    变化ids_In  In Varchar2 
  ); 

  --69.患者转病区转入
  Procedure Zlhis_Patient_026
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  );

  Procedure Zlhis_Patient_028(病人id_In In 病案主页.病人id%Type);

  --血库:科室配血完成
  Procedure Zlhis_Blood_001(医嘱id_In In 病人医嘱记录.Id%Type);
  --血库:科室配血拒绝
  Procedure Zlhis_Blood_002(医嘱id_In In 病人医嘱记录.Id%Type);

  --70.检验标本审核
  Procedure Zlhis_Lis_001(标本id_In In 检验标本记录.Id%Type);
  --71.检验标本审核撤消
  Procedure Zlhis_Lis_002(标本id_In In 检验标本记录.Id%Type);
  --73.检验标本条码打印
  Procedure Zlhis_Lis_004
  (
    样本条码_In In 病人医嘱发送.样本条码%Type,
    医嘱id_In   In 病人医嘱发送.医嘱id%Type,
    医嘱ids_In  In Varchar2
  );
  --74.检验标本条码打印撤销
  Procedure Zlhis_Lis_005
  (
    样本条码_In In 病人医嘱发送.样本条码%Type,
    医嘱id_In   In 病人医嘱发送.医嘱id%Type,
    医嘱ids_In  In Varchar2
  );
  --75.检验标本核收
  Procedure Zlhis_Lis_006(标本id_In In 检验标本记录.Id%Type);
  --76.检验标本核收撤销
  Procedure Zlhis_Lis_007(标本id_In In 检验标本记录.Id%Type);
  --77.检验标本拒收
  Procedure Zlhis_Lis_008(标本id_In In 检验标本记录.Id%Type);
End b_Message;
/
Create Or Replace Package Body b_Message Is
  --是否是平台调用
  Is_Platform_Call Number(1) := 0;
  --消息公共方法
  Message_Creator Zlmsg_Todo.Creator%Type := Null;
  --缓存消息查询结果
  Type Tmap_Msg_Using Is Table Of Number(1) Index By Varchar2(30);
  Zlmsg_Map Tmap_Msg_Using;
  --消息是否启用
  Function p_Msg_Using(Msg_Code_In Zlmsg_Lists.Code%Type) Return Number As
    n_Using Zlmsg_Lists.Using%Type;
    v_Code  Zlmsg_Lists.Code%Type;
  Begin
    If Is_Platform_Call = 1 Then
      Return 0;
    End If;
    v_Code := Upper(Msg_Code_In);
    Begin
      n_Using := Zlmsg_Map(v_Code);
      Return n_Using;
    Exception
      When No_Data_Found Then
        --不采取Max容错处理，错误相当于外键,用户可能没有采取同步修改或自己增加了消息类型但是未注册到Zlmsg_Lists，这两种情况会出现错误。



      
        Select Nvl(Using, 0) Into n_Using From Zlmsg_Lists A Where Code = v_Code;
        Zlmsg_Map(v_Code) := n_Using;
        --查询生成消息的人员，放在这里减少执行次数
        If Message_Creator Is Null Then
          Message_Creator := Zl_Username;
        End If;
        Return n_Using;
    End;
  Exception
    When Others Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || '未在Zlmsg_Lists中找到消息"' || v_Code || '"！请联系管理员进行处理。' || '[ZLSOFT]');
      Return 0;
  End;
  Procedure p_Msg_Todo_Insert
  (
    Msg_Code_In  Zlmsg_Todo.Msg_Code%Type,
    Key_Value_In Zlmsg_Todo.Key_Value%Type
  ) Is
  Begin
    If p_Msg_Using(Msg_Code_In) = 0 Then
      Return;
    End If;
    Insert Into Zlmsg_Todo
      (ID, Msg_Code, Key_Value, State, Create_Time, Creator)
    Values
      (Zlmsg_Todo_Id.Nextval, Upper(Msg_Code_In), Key_Value_In, 0, Sysdate, Message_Creator);
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Msg_Todo_Insert;
  --设置当前会话为平台调用
  Procedure Set_Platform_Call(Platform_Call Number) Is
  Begin
    Is_Platform_Call := Platform_Call;
  End Set_Platform_Call;
  --消息Zlhis_Dict_001
  Procedure Zlhis_Dict_001(Id_In 部门表.Id%Type) Is
    v_Define Xmltype;
    v_Value  Zlmsg_Todo.Key_Value%Type;
  Begin
    If p_Msg_Using('ZLHIS_DICT_001') = 0 Then
      Return;
    End If;
    Begin
      Select Xmltype(Key_Define) Into v_Define From Zlmsg_Lists Where Code = 'ZLHIS_DICT_001';
    Exception
      When Others Then
        v_Define := Xmltype('<root><ID>NULL</ID></root>');
    End;
    Select Updatexml(v_Define, '/root/ID/text()', Id_In).Getstringval() Into v_Value From Dual;
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_001', v_Value);
  End Zlhis_Dict_001;
  --修改部门
  Procedure Zlhis_Dict_002(部门id_In 部门表.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><部门ID>' || 部门id_In || '</部门ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_002', v_Value);
  End Zlhis_Dict_002;
  --停用部门
  Procedure Zlhis_Dict_003(部门id_In 部门表.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><部门ID>' || 部门id_In || '</部门ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_003', v_Value);
  End Zlhis_Dict_003;
  --启用部门
  Procedure Zlhis_Dict_004(部门id_In 部门表.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><部门ID>' || 部门id_In || '</部门ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_004', v_Value);
  End Zlhis_Dict_004;
  --新增人员
  Procedure Zlhis_Dict_005(人员id_In 人员表.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><人员ID>' || 人员id_In || '</人员ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_005', v_Value);
  End Zlhis_Dict_005;
  --修改人员
  Procedure Zlhis_Dict_006(人员id_In 人员表.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><人员ID>' || 人员id_In || '</人员ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_006', v_Value);
  End Zlhis_Dict_006;
  --停用人员
  Procedure Zlhis_Dict_007(人员id_In 人员表.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><人员ID>' || 人员id_In || '</人员ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_007', v_Value);
  End Zlhis_Dict_007;
  --启用人员
  Procedure Zlhis_Dict_008(人员id_In 人员表.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><人员ID>' || 人员id_In || '</人员ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_008', v_Value);
  End Zlhis_Dict_008;
  --新增收费项目
  Procedure Zlhis_Dict_009(细目id_In 收费项目目录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><细目ID>' || 细目id_In || '</细目ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_009', v_Value);
  End Zlhis_Dict_009;
  --修改收费项目
  Procedure Zlhis_Dict_010(细目id_In 收费项目目录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><细目ID>' || 细目id_In || '</细目ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_010', v_Value);
  End Zlhis_Dict_010;
  --停用收费项目
  Procedure Zlhis_Dict_011(细目id_In 收费项目目录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><细目ID>' || 细目id_In || '</细目ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_011', v_Value);
  End Zlhis_Dict_011;
  --启用收费项目
  Procedure Zlhis_Dict_012(细目id_In 收费项目目录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><细目ID>' || 细目id_In || '</细目ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_012', v_Value);
  End Zlhis_Dict_012;
  --新增诊疗项目
  Procedure Zlhis_Dict_013(诊疗id_In 诊疗项目目录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><诊疗ID>' || 诊疗id_In || '</诊疗ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_013', v_Value);
  End Zlhis_Dict_013;
  --修改诊疗项目
  Procedure Zlhis_Dict_014(诊疗id_In 诊疗项目目录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><诊疗ID>' || 诊疗id_In || '</诊疗ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_014', v_Value);
  End Zlhis_Dict_014;
  --停用诊疗项目
  Procedure Zlhis_Dict_015(诊疗id_In 诊疗项目目录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><诊疗ID>' || 诊疗id_In || '</诊疗ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_015', v_Value);
  End Zlhis_Dict_015;
  --启用诊疗项目
  Procedure Zlhis_Dict_016(诊疗id_In 诊疗项目目录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><诊疗ID>' || 诊疗id_In || '</诊疗ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_016', v_Value);
  End Zlhis_Dict_016;
  --新增检验项目
  Procedure Zlhis_Dict_017(诊疗id_In 诊疗项目目录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><诊疗ID>' || 诊疗id_In || '</诊疗ID><系统>1</系统></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_017', v_Value);
  End Zlhis_Dict_017;
  --修改检验项目
  Procedure Zlhis_Dict_018(诊疗id_In 诊疗项目目录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><诊疗ID>' || 诊疗id_In || '</诊疗ID><系统>1</系统></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_018', v_Value);
  End Zlhis_Dict_018;
  --删除检验项目
  Procedure Zlhis_Dict_019
  (
    诊疗id_In 诊疗项目目录.Id%Type,
    编码_In   诊治所见项目.编码%Type,
    中文名_In 诊治所见项目.中文名%Type,
    英文名_In 诊治所见项目.英文名%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><诊疗ID>' || 诊疗id_In || '</诊疗ID>' || '<编码>' || 编码_In || '</编码>' || '<中文名>' || 中文名_In || '</中文名>' ||
               '<英文名>' || 英文名_In || '</英文名>' || '<系统>1</系统></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_019', v_Value);
  End Zlhis_Dict_019;
  --新增疾病编码目录
  Procedure Zlhis_Dict_021(疾病id_In 疾病编码目录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><疾病ID>' || 疾病id_In || '</疾病ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_021', v_Value);
  End Zlhis_Dict_021;
  --修改疾病编码目录
  Procedure Zlhis_Dict_022(疾病id_In 疾病编码目录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><疾病ID>' || 疾病id_In || '</疾病ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_022', v_Value);
  End Zlhis_Dict_022;
  --停用疾病编码目录
  Procedure Zlhis_Dict_023(疾病id_In 疾病编码目录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><疾病ID>' || 疾病id_In || '</疾病ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_023', v_Value);
  End Zlhis_Dict_023;
  --启用疾病编码目录
  Procedure Zlhis_Dict_024(疾病id_In 疾病编码目录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><疾病ID>' || 疾病id_In || '</疾病ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_024', v_Value);
  End Zlhis_Dict_024;
  --新增药品分类
  Procedure Zlhis_Dict_025
  (
    类型_In 诊疗分类目录.类型%Type,
    Id_In   诊疗分类目录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类型>' || 类型_In || '</类型><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_025', v_Value);
  End Zlhis_Dict_025;
  --修改药品分类
  Procedure Zlhis_Dict_026
  (
    类型_In 诊疗分类目录.类型%Type,
    Id_In   诊疗分类目录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类型>' || 类型_In || '</类型><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_026', v_Value);
  End Zlhis_Dict_026;
  --删除药品分类
  Procedure Zlhis_Dict_027
  (
    类型_In 诊疗分类目录.类型%Type,
    Id_In   诊疗分类目录.Id%Type,
    编码_In 诊疗分类目录.编码%Type,
    名称_In 诊疗分类目录.名称%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类型>' || 类型_In || '</类型><ID>' || Id_In || '</ID><编码>' || 编码_In || '</编码><名称>' || 名称_In ||
               '</名称></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_027', v_Value);
  End Zlhis_Dict_027;
  --停用药品分类
  Procedure Zlhis_Dict_028
  (
    类型_In 诊疗分类目录.类型%Type,
    Id_In   诊疗分类目录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类型>' || 类型_In || '</类型><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_028', v_Value);
  End Zlhis_Dict_028;
  --启用药品分类
  Procedure Zlhis_Dict_029
  (
    类型_In 诊疗分类目录.类型%Type,
    Id_In   诊疗分类目录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类型>' || 类型_In || '</类型><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_029', v_Value);
  End Zlhis_Dict_029;
  --新增药品品种
  Procedure Zlhis_Dict_030
  (
    类别_In 诊疗项目目录.类别%Type,
    Id_In   诊疗项目目录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类别>' || 类别_In || '</类别><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_030', v_Value);
  End Zlhis_Dict_030;
  --修改药品品种
  Procedure Zlhis_Dict_031
  (
    类别_In 诊疗项目目录.类别%Type,
    Id_In   诊疗项目目录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类别>' || 类别_In || '</类别><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_031', v_Value);
  End Zlhis_Dict_031;
  --删除药品品种
  Procedure Zlhis_Dict_032
  (
    类别_In 诊疗项目目录.类别%Type,
    Id_In   诊疗项目目录.Id%Type,
    编码_In 诊疗项目目录.编码%Type,
    名称_In 诊疗项目目录.名称%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类别>' || 类别_In || '</类别><ID>' || Id_In || '</ID><编码>' || 编码_In || '</编码><名称>' || 名称_In ||
               '</名称></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_032', v_Value);
  End Zlhis_Dict_032;
  --停用药品品种
  Procedure Zlhis_Dict_033
  (
    类别_In 诊疗项目目录.类别%Type,
    Id_In   诊疗项目目录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类别>' || 类别_In || '</类别><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_033', v_Value);
  End Zlhis_Dict_033;
  --启用药品品种
  Procedure Zlhis_Dict_034
  (
    类别_In 诊疗项目目录.类别%Type,
    Id_In   诊疗项目目录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类别>' || 类别_In || '</类别><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_034', v_Value);
  End Zlhis_Dict_034;
  --新增药品规格
  Procedure Zlhis_Dict_035
  (
    类别_In   收费项目目录.类别%Type,
    药品id_In 收费项目目录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类别>' || 类别_In || '</类别><药品ID>' || 药品id_In || '</药品ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_035', v_Value);
  End Zlhis_Dict_035;
  --修改药品规格
  Procedure Zlhis_Dict_036
  (
    类别_In   收费项目目录.类别%Type,
    药品id_In 收费项目目录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类别>' || 类别_In || '</类别><药品ID>' || 药品id_In || '</药品ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_036', v_Value);
  End Zlhis_Dict_036;
  --删除药品规格
  Procedure Zlhis_Dict_037
  (
    类别_In   收费项目目录.类别%Type,
    药品id_In 收费项目目录.Id%Type,
    编码_In   收费项目目录.编码%Type,
    名称_In   收费项目目录.名称%Type,
    规格_In   收费项目目录.规格%Type,
    产地_In   收费项目目录.产地%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类别>' || 类别_In || '</类别><药品ID>' || 药品id_In || '</药品ID><编码>' || 编码_In || '</编码><名称>' || 名称_In ||
               '</名称><规格>' || 规格_In || '</规格><产地>' || 产地_In || '</产地></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_037', v_Value);
  End Zlhis_Dict_037;
  --停用药品规格
  Procedure Zlhis_Dict_038
  (
    类别_In   收费项目目录.类别%Type,
    药品id_In 收费项目目录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类别>' || 类别_In || '</类别><药品ID>' || 药品id_In || '</药品ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_038', v_Value);
  End Zlhis_Dict_038;
  --启用药品规格
  Procedure Zlhis_Dict_039
  (
    类别_In   收费项目目录.类别%Type,
    药品id_In 收费项目目录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类别>' || 类别_In || '</类别><药品ID>' || 药品id_In || '</药品ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_039', v_Value);
  End Zlhis_Dict_039;
  --设置药品存储库房
  Procedure Zlhis_Dict_040
  (
    类别_In   收费项目目录.类别%Type,
    药品id_In 收费项目目录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类别>' || 类别_In || '</类别><药品ID>' || 药品id_In || '</药品ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_040', v_Value);
  End Zlhis_Dict_040;
  --设置药品储备限额
  Procedure Zlhis_Dict_041
  (
    类别_In   收费项目目录.类别%Type,
    药品id_In 收费项目目录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类别>' || 类别_In || '</类别><药品ID>' || 药品id_In || '</药品ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_041', v_Value);
  End Zlhis_Dict_041;
  --新增卫材品种
  Procedure Zlhis_Dict_042(Id_In 诊疗项目目录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><材料ID>' || Id_In || '</材料ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_042', v_Value);
  End Zlhis_Dict_042;
  --新增卫材规格
  Procedure Zlhis_Dict_043(Id_In 收费项目目录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><材料ID>' || Id_In || '</材料ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_043', v_Value);
  End Zlhis_Dict_043;
  --修改卫材规格
  Procedure Zlhis_Dict_044(Id_In 收费项目目录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><材料ID>' || Id_In || '</材料ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_044', v_Value);
  End Zlhis_Dict_044;
  --删除卫材规格
  Procedure Zlhis_Dict_045
  (
    Id_In   收费项目目录.Id%Type,
    编码_In 收费项目目录.编码%Type,
    名称_In 收费项目目录.名称%Type,
    规格_In 收费项目目录.规格%Type,
    产地_In 收费项目目录.产地%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><材料ID>' || Id_In || '</材料ID><编码>' || 编码_In || '</编码><名称>' || 名称_In || '</名称><规格>' || 规格_In ||
               '</规格><产地>' || 产地_In || '</产地></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_045', v_Value);
  End Zlhis_Dict_045;
  --停用卫材规格
  Procedure Zlhis_Dict_046(Id_In 收费项目目录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><材料ID>' || Id_In || '</材料ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_046', v_Value);
  End Zlhis_Dict_046;
  --启用卫材规格
  Procedure Zlhis_Dict_047(Id_In 收费项目目录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><材料ID>' || Id_In || '</材料ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_047', v_Value);
  End Zlhis_Dict_047;
  --医保对码
  Procedure Zlhis_Dict_048
  (
    险类_In       In 保险支付项目.险类%Type,
    收费细目id_In In 保险支付项目.收费细目id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><险类>' || 险类_In || '</险类><收费细目ID>' || 收费细目id_In || '</收费细目ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_048', v_Value);
  End Zlhis_Dict_048;
  --删除医保对码
  Procedure Zlhis_Dict_049
  (
    险类_In       In 保险支付项目.险类%Type,
    收费细目id_In In 保险支付项目.收费细目id%Type,
    项目编码_In   In 收费项目目录.编码%Type,
    项目名称_In   In 收费项目目录.名称%Type,
    医保编码_In   In 保险支付项目.项目编码%Type,
    医保名称_In   In 保险支付项目.项目名称%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><险类>' || 险类_In || '</险类><收费细目ID>' || 收费细目id_In || '</收费细目ID><项目编码>' || 项目编码_In || '</项目编码><项目名称>' ||
               项目名称_In || '</项目名称><医保编码>' || 医保编码_In || '</医保编码><医保名称>' || 医保名称_In || '</医保名称></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_049', v_Value);
  End Zlhis_Dict_049;
  --新增卫材分类
  Procedure Zlhis_Dict_050
  (
    类型_In 诊疗分类目录.类型%Type,
    Id_In   诊疗分类目录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类型>' || 类型_In || '</类型><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_050', v_Value);
  End Zlhis_Dict_050;
  --修改卫材分类
  Procedure Zlhis_Dict_051
  (
    类型_In 诊疗分类目录.类型%Type,
    Id_In   诊疗分类目录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类型>' || 类型_In || '</类型><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_051', v_Value);
  End Zlhis_Dict_051;
  --删除卫材分类
  Procedure Zlhis_Dict_052
  (
    类型_In 诊疗分类目录.类型%Type,
    Id_In   诊疗分类目录.Id%Type,
    编码_In 诊疗分类目录.编码%Type,
    名称_In 诊疗分类目录.名称%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类型>' || 类型_In || '</类型><ID>' || Id_In || '</ID><编码>' || 编码_In || '</编码><名称>' || 名称_In ||
               '</名称></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_052', v_Value);
  End Zlhis_Dict_052;
  --收费价目变动
  Procedure Zlhis_Dict_053
  (
    收费项目Id_In       收费项目目录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><收费项目ID>' || 收费项目Id_In || '</收费项目ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_053', v_Value);
  End Zlhis_Dict_053;

  --诊疗收费对照变动
  Procedure Zlhis_Dict_054
  (
    诊疗项目Id_In     诊疗分类目录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><诊疗项目ID>' || 诊疗项目Id_In || '</诊疗项目ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_054', v_Value);
  End Zlhis_Dict_054;
  --新增诊疗检查类型
  Procedure Zlhis_Dictpacs_001
  (
    编码_In   诊疗检查类型.编码%Type,
    名称_In   诊疗检查类型.名称%Type,
    简码_In   诊疗检查类型.简码%Type,
    建病案_In 诊疗检查类型.建病案%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><编码>' || 编码_In || '</编码><名称>' || 名称_In || '</名称><简码>' || 简码_In || '</简码><建病案>' || 建病案_In ||
               '</建病案></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_001', v_Value);
  End Zlhis_Dictpacs_001;

  --修改诊疗检查类型
  Procedure Zlhis_Dictpacs_002
  (
    编码_In   诊疗检查类型.编码%Type,
    名称_In   诊疗检查类型.名称%Type,
    简码_In   诊疗检查类型.简码%Type,
    建病案_In 诊疗检查类型.建病案%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><编码>' || 编码_In || '</编码><名称>' || 名称_In || '</名称><简码>' || 简码_In || '</简码><建病案>' || 建病案_In ||
               '</建病案></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_002', v_Value);
  End Zlhis_Dictpacs_002;
  --删除诊疗检查类型
  Procedure Zlhis_Dictpacs_003
  (
    编码_In   诊疗检查类型.编码%Type,
    名称_In   诊疗检查类型.名称%Type,
    简码_In   诊疗检查类型.简码%Type,
    建病案_In 诊疗检查类型.建病案%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><编码>' || 编码_In || '</编码><名称>' || 名称_In || '</名称><简码>' || 简码_In || '</简码><建病案>' || 建病案_In ||
               '</建病案></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_003', v_Value);
  End Zlhis_Dictpacs_003;
  --新增诊疗检查部位
  Procedure Zlhis_Dictpacs_004
  (
    类型_In     诊疗检查部位.类型%Type,
    编码_In     诊疗检查部位.编码%Type,
    名称_In     诊疗检查部位.名称%Type,
    分组_In     诊疗检查部位.分组%Type,
    备注_In     诊疗检查部位.备注%Type,
    方法_In     诊疗检查部位.方法%Type,
    适用性别_In 诊疗检查部位.适用性别%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类型>' || 类型_In || '</类型><编码>' || 编码_In || '</编码><名称>' || 名称_In || '</名称><分组>' || 分组_In ||
               '</分组><备注>' || 备注_In || '</备注><方法>' || 方法_In || '</方法><适用性别>' || 适用性别_In || '</适用性别></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_004', v_Value);
  End Zlhis_Dictpacs_004;
  --修改诊疗检查部位
  Procedure Zlhis_Dictpacs_005
  (
    类型_In     诊疗检查部位.类型%Type,
    编码_In     诊疗检查部位.编码%Type,
    名称_In     诊疗检查部位.名称%Type,
    分组_In     诊疗检查部位.分组%Type,
    备注_In     诊疗检查部位.备注%Type,
    方法_In     诊疗检查部位.方法%Type,
    适用性别_In 诊疗检查部位.适用性别%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类型>' || 类型_In || '</类型><编码>' || 编码_In || '</编码><名称>' || 名称_In || '</名称><分组>' || 分组_In ||
               '</分组><备注>' || 备注_In || '</备注><方法>' || 方法_In || '</方法><适用性别>' || 适用性别_In || '</适用性别></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_005', v_Value);
  End Zlhis_Dictpacs_005;
  --删除诊疗检查部位
  Procedure Zlhis_Dictpacs_006
  (
    类型_In     诊疗检查部位.类型%Type,
    编码_In     诊疗检查部位.编码%Type,
    名称_In     诊疗检查部位.名称%Type,
    分组_In     诊疗检查部位.分组%Type,
    备注_In     诊疗检查部位.备注%Type,
    方法_In     诊疗检查部位.方法%Type,
    适用性别_In 诊疗检查部位.适用性别%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类型>' || 类型_In || '</类型><编码>' || 编码_In || '</编码><名称>' || 名称_In || '</名称><分组>' || 分组_In ||
               '</分组><备注>' || 备注_In || '</备注><方法>' || 方法_In || '</方法><适用性别>' || 适用性别_In || '</适用性别></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_006', v_Value);
  End Zlhis_Dictpacs_006;
  --新增诊疗项目部位
  Procedure Zlhis_Dictpacs_007
  (
    Id_In     诊疗项目部位.Id%Type,
    项目id_In 诊疗项目部位.项目id%Type,
    类型_In   诊疗项目部位.类型%Type,
    部位_In   诊疗项目部位.部位%Type,
    方法_In   诊疗项目部位.方法%Type,
    默认_In   诊疗项目部位.默认%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ID>' || Id_In || '</ID><项目ID>' || 项目id_In || '</项目ID><类型>' || 类型_In || '</类型><部位>' || 部位_In ||
               '</部位><方法>' || 方法_In || '</方法><默认>' || 默认_In || '</默认></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_007', v_Value);
  End Zlhis_Dictpacs_007;
  --修改诊疗项目部位
  Procedure Zlhis_Dictpacs_008
  (
    Id_In     诊疗项目部位.Id%Type,
    项目id_In 诊疗项目部位.项目id%Type,
    类型_In   诊疗项目部位.类型%Type,
    部位_In   诊疗项目部位.部位%Type,
    方法_In   诊疗项目部位.方法%Type,
    默认_In   诊疗项目部位.默认%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ID>' || Id_In || '</ID><项目ID>' || 项目id_In || '</项目ID><类型>' || 类型_In || '</类型><部位>' || 部位_In ||
               '</部位><方法>' || 方法_In || '</方法><默认>' || 默认_In || '</默认></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_008', v_Value);
  End Zlhis_Dictpacs_008;
  --删除诊疗项目部位
  Procedure Zlhis_Dictpacs_009
  (
    Id_In     诊疗项目部位.Id%Type,
    项目id_In 诊疗项目部位.项目id%Type,
    类型_In   诊疗项目部位.类型%Type,
    部位_In   诊疗项目部位.部位%Type,
    方法_In   诊疗项目部位.方法%Type,
    默认_In   诊疗项目部位.默认%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ID>' || Id_In || '</ID><项目ID>' || 项目id_In || '</项目ID><类型>' || 类型_In || '</类型><部位>' || 部位_In ||
               '</部位><方法>' || 方法_In || '</方法><默认>' || 默认_In || '</默认></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_009', v_Value);
  End Zlhis_Dictpacs_009;
  --新增诊疗项目部位
  Procedure Zlhis_Dictlis_004
  (
    编码_In     诊疗检验标本.编码%Type,
    名称_In     诊疗检验标本.名称%Type,
    简码_In     诊疗检验标本.简码%Type,
    适用性别_In 诊疗检验标本.适用性别%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><编码>' || 编码_In || '</编码><名称>' || 名称_In || '</名称><简码>' || 简码_In || '</简码><适用性别>' || 适用性别_In ||
               '</适用性别></root>';
    b_Message.p_Msg_Todo_Insert('Zlhis_DictLis_004', v_Value);
  End Zlhis_Dictlis_004;
  --修改诊疗项目部位
  Procedure Zlhis_Dictlis_005
  (
    编码_In     诊疗检验标本.编码%Type,
    名称_In     诊疗检验标本.名称%Type,
    简码_In     诊疗检验标本.简码%Type,
    适用性别_In 诊疗检验标本.适用性别%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><编码>' || 编码_In || '</编码><名称>' || 名称_In || '</名称><简码>' || 简码_In || '</简码><适用性别>' || 适用性别_In ||
               '</适用性别></root>';
    b_Message.p_Msg_Todo_Insert('Zlhis_DictLis_005', v_Value);
  End Zlhis_Dictlis_005;
  --删除诊疗项目部位
  Procedure Zlhis_Dictlis_006
  (
    编码_In     诊疗检验标本.编码%Type,
    名称_In     诊疗检验标本.名称%Type,
    简码_In     诊疗检验标本.简码%Type,
    适用性别_In 诊疗检验标本.适用性别%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><编码>' || 编码_In || '</编码><名称>' || 名称_In || '</名称><简码>' || 简码_In || '</简码><适用性别>' || 适用性别_In ||
               '</适用性别></root>';
    b_Message.p_Msg_Todo_Insert('Zlhis_DictLis_006', v_Value);
  End Zlhis_Dictlis_006;
  --新增采血管类型
  Procedure Zlhis_Dictlis_007
  (
    编码_In   采血管类型.编码%Type,
    名称_In   采血管类型.名称%Type,
    简码_In   采血管类型.简码%Type,
    添加剂_In 采血管类型.添加剂%Type,
    采血量_In 采血管类型.采血量%Type,
    规格_In   采血管类型.规格%Type,
    颜色_In   采血管类型.颜色%Type,
    材料id_In 采血管类型.材料id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><编码>' || 编码_In || '</编码><名称>' || 名称_In || '</名称><简码>' || 简码_In || '</简码><添加剂>' || 添加剂_In ||
               '</添加剂><采血量>' || 采血量_In || '</采血量><颜色>' || 颜色_In || '</颜色><规格>' || 规格_In || '</规格><材料ID_In>' || 材料id_In ||
               '</材料ID_In></root>';
    b_Message.p_Msg_Todo_Insert('Zlhis_DictLis_007', v_Value);
  End Zlhis_Dictlis_007;
  --新增采血管类型
  Procedure Zlhis_Dictlis_008
  (
    编码_In   采血管类型.编码%Type,
    名称_In   采血管类型.名称%Type,
    简码_In   采血管类型.简码%Type,
    添加剂_In 采血管类型.添加剂%Type,
    采血量_In 采血管类型.采血量%Type,
    规格_In   采血管类型.规格%Type,
    颜色_In   采血管类型.颜色%Type,
    材料id_In 采血管类型.材料id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><编码>' || 编码_In || '</编码><名称>' || 名称_In || '</名称><简码>' || 简码_In || '</简码><添加剂>' || 添加剂_In ||
               '</添加剂><采血量>' || 采血量_In || '</采血量><颜色>' || 颜色_In || '</颜色><规格>' || 规格_In || '</规格><材料ID_In>' || 材料id_In ||
               '</材料ID_In></root>';
    b_Message.p_Msg_Todo_Insert('Zlhis_DictLis_008', v_Value);
  End Zlhis_Dictlis_008;
  --新增采血管类型
  Procedure Zlhis_Dictlis_009
  (
    编码_In   采血管类型.编码%Type,
    名称_In   采血管类型.名称%Type,
    简码_In   采血管类型.简码%Type,
    添加剂_In 采血管类型.添加剂%Type,
    采血量_In 采血管类型.采血量%Type,
    规格_In   采血管类型.规格%Type,
    颜色_In   采血管类型.颜色%Type,
    材料id_In 采血管类型.材料id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><编码>' || 编码_In || '</编码><名称>' || 名称_In || '</名称><简码>' || 简码_In || '</简码><添加剂>' || 添加剂_In ||
               '</添加剂><采血量>' || 采血量_In || '</采血量><颜色>' || 颜色_In || '</颜色><规格>' || 规格_In || '</规格><材料ID_In>' || 材料id_In ||
               '</材料ID_In></root>';
    b_Message.p_Msg_Todo_Insert('Zlhis_DictLis_009', v_Value);
  End Zlhis_Dictlis_009;
  --药品备药发送
  Procedure Zlhis_Drug_001(No_In 药品收发记录.No%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><单据号>' || No_In || '</单据号></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_001', v_Value);
  End Zlhis_Drug_001;
  --取消药品备药发送
  Procedure Zlhis_Drug_002(No_In 药品收发记录.No%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><单据号>' || No_In || '</单据号></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_002', v_Value);
  End Zlhis_Drug_002;
  --药品移库单接收
  Procedure Zlhis_Drug_003(No_In 药品收发记录.No%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><单据号>' || No_In || '</单据号></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_003', v_Value);
  End Zlhis_Drug_003;
  --药品移库单冲销
  Procedure Zlhis_Drug_004(No_In 药品收发记录.No%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><单据号>' || No_In || '</单据号></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_004', v_Value);
  End Zlhis_Drug_004;
  --部门发药
  Procedure Zlhis_Drug_005
  (
    库房id_In 药品收发记录.库房id%Type,
    收发id_In 药品收发记录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><库房ID>' || 库房id_In || '</库房ID><收发ID>' || 收发id_In || '</收发ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_005', v_Value);
  End Zlhis_Drug_005;
  --部门退药
  Procedure Zlhis_Drug_006
  (
    冲销收发id_In 药品收发记录.Id%Type,
    待发收发id_In 药品收发记录.Id%Type,
    数量_In       药品收发记录.实际数量%Type,
    费用id_In     门诊费用记录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><冲销记录ID>' || 冲销收发id_In || '</冲销记录ID><待发记录ID>' || 待发收发id_In || '</待发记录ID><数量>' || 数量_In ||
               '</数量><费用ID>' || 费用id_In || '</费用ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_006', v_Value);
  End Zlhis_Drug_006;
  --药品调价
  Procedure Zlhis_Drug_007(价格id_In 药品价格记录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><价格ID>' || 价格id_In || '</价格ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_007', v_Value);
  End Zlhis_Drug_007;
  --静配发送
  Procedure Zlhis_Drug_008(记录ids_In Varchar2) Is
    v_Value  Zlmsg_Todo.Key_Value%Type;
    n_记录id 输液配药记录.Id%Type;
    v_Tmp    Varchar2(4000);
	n_Length Number(18);
  Begin
    If 记录ids_In Is Null Then
      v_Tmp := Null;
    Else
      v_Tmp := 记录ids_In || ',';
    End If;
  
    v_Value := '<root><记录IDS>';
  
    While v_Tmp Is Not Null Loop
      --分解单据ID串
      n_记录id := To_Number(Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1));
      v_Tmp    := Replace(',' || v_Tmp, ',' || n_记录id || ',');
      
      --判断当前长度是否即将超过缓存                                                                        
      Select Lengthb(v_Value || '<记录ID>' || n_记录id || '</记录ID>') Into n_Length From Dual;            
      If n_Length > 950 Then								                   
        v_Value := v_Value || '</记录IDs></root>';                                                         
        b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_008', v_Value);                                            
        v_Value := '<root><记录IDs>';                                                                      
      End If;

      v_Value := v_Value || '<记录ID>' || n_记录id || '</记录ID>';
    End Loop;
  
    v_Value := v_Value || '</记录IDS></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_008', v_Value);
  End Zlhis_Drug_008;
  --药品调售价
  Procedure Zlhis_Drug_009
  (
    价格id_In 药品价格记录.Id%Type,
    时价_In   Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><价格ID>' || 价格id_In || '</价格ID><时价>' || 时价_In || '</时价></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_009', v_Value);
  End Zlhis_Drug_009;
  --卫材调成本价
  Procedure Zlhis_Drug_010(价格id_In 成本价调价信息.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><价格ID>' || 价格id_In || '</价格ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_010', v_Value);
  End Zlhis_Drug_010;
  --卫材调售价
  Procedure Zlhis_Drug_011
  (
    价格id_In 收费价目.Id%Type,
    时价_In   Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><价格ID>' || 价格id_In || '</价格ID><时价>' || 时价_In || '</时价></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_011', v_Value);
  End Zlhis_Drug_011;

  --2.停止患者医嘱，住院
  Procedure Zlhis_Cis_002
  (
    病人id_In  In 病人医嘱记录.病人id%Type,
    主页id_In  In 病人医嘱记录.主页id%Type,
    医嘱id_In  In 病人医嘱记录.Id%Type,
    医嘱ids_In In Varchar2
  ) Is
  Begin
    If p_Msg_Using('ZLHIS_CIS_002') = 0 Then
      Return;
    End If;
    If 医嘱id_In Is Not Null Then
      b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_002',
                                  '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><ID>' || 医嘱id_In ||
                                   '</ID></root>');
    Else
      For R In (Select '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><ID>' || ID || '</ID></root>' As Xml_Value
                From 病人医嘱记录
                Where ID In (Select Column_Value From Table(f_Num2list(医嘱ids_In))) And 相关id Is Null) Loop
        b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_002', r.Xml_Value);
      End Loop;
    End If;
  End Zlhis_Cis_002;
  --3.作废患者医嘱，门诊/住院
  Procedure Zlhis_Cis_003
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    挂号单_In In 病人挂号记录.No%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><挂号单>' || 挂号单_In || '</挂号单><ID>' ||
               医嘱id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_003', v_Value);
  End Zlhis_Cis_003;

  --4.患者术后医嘱，住院
  Procedure Zlhis_Cis_004
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('Zlhis_Cis_004',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><ID>' || 医嘱id_In ||
                                 '</ID></root>');
  End Zlhis_Cis_004;

  --5.撤消患者术后医嘱，住院
  Procedure Zlhis_Cis_005
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('Zlhis_Cis_005',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><ID>' || 医嘱id_In ||
                                 '</ID></root>');
  End Zlhis_Cis_005;

  --6.患者护理常规医嘱，住院
  Procedure Zlhis_Cis_006
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('Zlhis_Cis_006',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><发送号>' || 发送号_In ||
                                 '</发送号><ID>' || 医嘱id_In || '</ID></root>');
  End Zlhis_Cis_006;

  --7.撤消患者护理常规医嘱，住院
  Procedure Zlhis_Cis_007
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type,
    No_In     In 病人医嘱发送.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><发送号>' || 发送号_In || '</发送号><ID>' ||
               医嘱id_In || '</ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('Zlhis_Cis_007', v_Value);
  End Zlhis_Cis_007;

  --门诊患者接诊
  Procedure Zlhis_Cis_008
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    挂号单_In In 病人挂号记录.No%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_008', '<root><病人ID>' || 病人id_In || '</病人ID><NO>' || 挂号单_In || '</NO></root>');
  End Zlhis_Cis_008;

  --门诊患者取消接诊
  Procedure Zlhis_Cis_009
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    挂号单_In In 病人挂号记录.No%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_009', '<root><病人ID>' || 病人id_In || '</病人ID><NO>' || 挂号单_In || '</NO></root>');
  End Zlhis_Cis_009;

  --10.下达患者诊断，门诊/住院
  Procedure Zlhis_Cis_010
  (
    病人id_In In 病人挂号记录.病人id%Type,
    就诊id_In In 病人挂号记录.Id%Type, --门诊病人 挂号ID，住院病人 主页ID
    诊断id_In In 病人诊断记录.Id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_010',
                                '<root><病人ID>' || 病人id_In || '</病人ID><就诊ID>' || 就诊id_In || '</就诊ID><ID>' || 诊断id_In ||
                                 '</ID></root>');
  End Zlhis_Cis_010;
  --11.撤消患者诊断
  Procedure Zlhis_Cis_011
  (
    病人id_In   In 病人挂号记录.病人id%Type,
    就诊id_In   In 病人挂号记录.Id%Type, --门诊病人 挂号ID，住院病人 主页ID
    Id_In       In 病人诊断记录.Id%Type,
    疾病id_In   In 病人诊断记录.疾病id%Type,
    诊断id_In   In 病人诊断记录.诊断id%Type,
    诊断描述_In In 病人诊断记录.诊断描述%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><就诊ID>' || 就诊id_In || '</就诊ID><ID>' || Id_In || '</ID><疾病ID>' ||
               疾病id_In || '</疾病ID><诊断ID>' || 诊断id_In || '</诊断ID><诊断描述>' || 诊断描述_In || '</诊断描述></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_011', v_Value);
  End Zlhis_Cis_011;

  --病区执行医嘱校对
  Procedure Zlhis_Cis_012
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_012',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><ID>' || 医嘱id_In ||
                                 '</ID></root>');
  End Zlhis_Cis_012;

  --13.检验危急值阅读通知
  Procedure Zlhis_Cis_014
  (
    病人id_In In 病人挂号记录.病人id%Type,
    就诊id_In In 病人挂号记录.Id%Type, --门诊病人 挂号ID，住院病人 主页ID
    医嘱id_In In 病人医嘱记录.Id%Type,
    消息id_In In 业务消息清单.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><就诊ID>' || 就诊id_In || '</就诊ID><医嘱ID>' || 医嘱id_In || '</医嘱ID><ID>' ||
               消息id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_014', v_Value);
  End Zlhis_Cis_014;
  --15.患者检验申请，门诊/住院
  Procedure Zlhis_Cis_016
  (
    病人id_In   In 病人医嘱记录.病人id%Type,
    主页id_In   In 病人医嘱记录.主页id%Type,
    挂号单_In   In 病人挂号记录.No%Type,
    发送号_In   In 病人医嘱发送.发送号%Type,
    医嘱id_In   In 病人医嘱记录.Id%Type,
    病人来源_In In 病人医嘱记录.病人来源%Type --1-门诊;2-住院;3-外来(今后用于辅诊部门接收外来病人);4-体检病人
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><挂号单>' || 挂号单_In || '</挂号单><发送号>' ||
               发送号_In || '</发送号><ID>' || 医嘱id_In || '</ID><病人来源>' || 病人来源_In || '</病人来源></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_016', v_Value);
  End Zlhis_Cis_016;
  --16.患者检查申请，门诊/住院
  Procedure Zlhis_Cis_017
  (
    病人id_In   In 病人医嘱记录.病人id%Type,
    主页id_In   In 病人医嘱记录.主页id%Type,
    挂号单_In   In 病人挂号记录.No%Type,
    发送号_In   In 病人医嘱发送.发送号%Type,
    医嘱id_In   In 病人医嘱记录.Id%Type,
    病人来源_In In 病人医嘱记录.病人来源%Type --1-门诊;2-住院;3-外来(今后用于辅诊部门接收外来病人);4-体检病人
  ) Is
    v_Value    Zlmsg_Todo.Key_Value%Type;
    v_操作类型 诊疗项目目录.操作类型%Type;
  Begin
    Select Max(a.操作类型)
    Into v_操作类型
    From 诊疗项目目录 A, 病人医嘱记录 B
    Where b.诊疗项目id = a.Id And b.Id = 医嘱id_In;
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><挂号单>' || 挂号单_In || '</挂号单><发送号>' ||
               发送号_In || '</发送号><ID>' || 医嘱id_In || '</ID><病人来源>' || 病人来源_In || '</病人来源></root>';
    If v_操作类型 = '病理' Then
      b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_054', v_Value);
    Else
      b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_017', v_Value);
    End If;
  End Zlhis_Cis_017;
  --17.患者手术申请，门诊/住院
  Procedure Zlhis_Cis_018
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    挂号单_In In 病人挂号记录.No%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><挂号单>' || 挂号单_In || '</挂号单><发送号>' ||
               发送号_In || '</发送号><ID>' || 医嘱id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_018', v_Value);
  End Zlhis_Cis_018;
  --18.患者输血申请，住院
  Procedure Zlhis_Cis_019
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    挂号单_In In 病人挂号记录.No%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><挂号单>' || 挂号单_In || '</挂号单><发送号>' ||
               发送号_In || '</发送号><ID>' || 医嘱id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_019', v_Value);
  End Zlhis_Cis_019;
  --19.患者会诊申请，住院
  Procedure Zlhis_Cis_020
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><发送号>' || 发送号_In || '</发送号><ID>' ||
               医嘱id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_020', v_Value);
  End Zlhis_Cis_020;
  --20.患者抢救医嘱，住院
  Procedure Zlhis_Cis_021
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><发送号>' || 发送号_In || '</发送号><ID>' ||
               医嘱id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_021', v_Value);
  End Zlhis_Cis_021;
  --21.患者死亡医嘱，住院
  Procedure Zlhis_Cis_022
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><发送号>' || 发送号_In || '</发送号><ID>' ||
               医嘱id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_022', v_Value);
  End Zlhis_Cis_022;
  --22.患者特殊治疗医嘱，住院
  Procedure Zlhis_Cis_023
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><发送号>' || 发送号_In || '</发送号><ID>' ||
               医嘱id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_023', v_Value);
  End Zlhis_Cis_023;

  --24.检查危急值阅读通知
  Procedure Zlhis_Cis_025
  (
    病人id_In In 病人挂号记录.病人id%Type,
    就诊id_In In 病人挂号记录.Id%Type, --门诊病人 挂号ID，住院病人 主页ID
    医嘱id_In In 病人医嘱记录.Id%Type,
    消息id_In In 业务消息清单.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><就诊ID>' || 就诊id_In || '</就诊ID><医嘱ID>' || 医嘱id_In || '</医嘱ID><ID>' ||
               消息id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_025', v_Value);
  End Zlhis_Cis_025;

  --病区执行医嘱发送
  Procedure Zlhis_Cis_026
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_026',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><发送号>' || 发送号_In ||
                                 '</发送号><ID>' || 医嘱id_In || '</ID></root>');
  End Zlhis_Cis_026;

  --撤消患者检验申请
  Procedure Zlhis_Cis_036
  (
    病人id_In   In 病人医嘱记录.病人id%Type,
    主页id_In   In 病人医嘱记录.主页id%Type,
    挂号单_In   In 病人挂号记录.No%Type,
    发送号_In   In 病人医嘱发送.发送号%Type,
    医嘱id_In   In 病人医嘱记录.Id%Type,
    No_In       In 病人医嘱发送.No%Type,
    病人来源_In In 病人医嘱记录.病人来源%Type --1-门诊;2-住院;3-外来(今后用于辅诊部门接收外来病人);4-体检病人
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><挂号单>' || 挂号单_In || '</挂号单><发送号>' ||
               发送号_In || '</发送号><ID>' || 医嘱id_In || '</ID><NO>' || No_In || '</NO><病人来源>' || 病人来源_In ||
               '</病人来源></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_036', v_Value);
  End Zlhis_Cis_036;

  --撤消患者检查申请
  Procedure Zlhis_Cis_037
  (
    病人id_In   In 病人医嘱记录.病人id%Type,
    主页id_In   In 病人医嘱记录.主页id%Type,
    挂号单_In   In 病人挂号记录.No%Type,
    发送号_In   In 病人医嘱发送.发送号%Type,
    医嘱id_In   In 病人医嘱记录.Id%Type,
    No_In       In 病人医嘱发送.No%Type,
    病人来源_In In 病人医嘱记录.病人来源%Type --1-门诊;2-住院;3-外来(今后用于辅诊部门接收外来病人);4-体检病人
  ) Is
    v_Value    Zlmsg_Todo.Key_Value%Type;
    v_操作类型 诊疗项目目录.操作类型%Type;
  Begin
    Select Max(a.操作类型)
    Into v_操作类型
    From 诊疗项目目录 A, 病人医嘱记录 B
    Where b.诊疗项目id = a.Id And b.Id = 医嘱id_In;
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><挂号单>' || 挂号单_In || '</挂号单><发送号>' ||
               发送号_In || '</发送号><ID>' || 医嘱id_In || '</ID><NO>' || No_In || '</NO><病人来源>' || 病人来源_In ||
               '</病人来源></root>';
    If v_操作类型 = '病理' Then
      b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_055', v_Value);
    Else
      b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_037', v_Value);
    End If;
  End Zlhis_Cis_037;

  --撤消患者手术申请
  Procedure Zlhis_Cis_038
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    挂号单_In In 病人挂号记录.No%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type,
    No_In     In 病人医嘱发送.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><挂号单>' || 挂号单_In || '</挂号单><发送号>' ||
               发送号_In || '</发送号><ID>' || 医嘱id_In || '</ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_038', v_Value);
  End Zlhis_Cis_038;

  --撤消患者输血申请
  Procedure Zlhis_Cis_039
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    挂号单_In In 病人挂号记录.No%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type,
    No_In     In 病人医嘱发送.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><挂号单>' || 挂号单_In || '</挂号单><发送号>' ||
               发送号_In || '</发送号><ID>' || 医嘱id_In || '</ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_039', v_Value);
  End Zlhis_Cis_039;

  --撤消患者会诊申请
  Procedure Zlhis_Cis_040
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type,
    No_In     In 病人医嘱发送.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><发送号>' || 发送号_In || '</发送号><ID>' ||
               医嘱id_In || '</ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_040', v_Value);
  End Zlhis_Cis_040;

  --撤消患者抢救医嘱
  Procedure Zlhis_Cis_041
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type,
    No_In     In 病人医嘱发送.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><发送号>' || 发送号_In || '</发送号><ID>' ||
               医嘱id_In || '</ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_041', v_Value);
  End Zlhis_Cis_041;

  --撤消患者死亡医嘱
  Procedure Zlhis_Cis_042
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type,
    No_In     In 病人医嘱发送.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><发送号>' || 发送号_In || '</发送号><ID>' ||
               医嘱id_In || '</ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_042', v_Value);
  End Zlhis_Cis_042;

  --撤消特殊治疗医嘱
  Procedure Zlhis_Cis_043
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type,
    No_In     In 病人医嘱发送.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><发送号>' || 发送号_In || '</发送号><ID>' ||
               医嘱id_In || '</ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_043', v_Value);
  End Zlhis_Cis_043;

  --撤消病区执行医嘱
  Procedure Zlhis_Cis_044
  (
    病人id_In   In 病人医嘱记录.病人id%Type,
    主页id_In   In 病人医嘱记录.主页id%Type,
    发送号_In   In 病人医嘱发送.发送号%Type,
    医嘱id_In   In 病人医嘱记录.Id%Type,
    No_In       In 病人医嘱发送.No%Type,
    发送数次_In In 病人医嘱发送.发送数次%Type,
    首次时间_In In 病人医嘱发送.首次时间%Type,
    末次时间_In In 病人医嘱发送.末次时间%Type,
    样本条码_In In 病人医嘱发送.样本条码%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><发送号>' || 发送号_In || '</发送号><ID>' ||
               医嘱id_In || '</ID><NO>' || No_In || '</NO><发送数次>' || 发送数次_In || '</发送数次><首次时间>' ||
               To_Char(首次时间_In, 'yyyy-mm-dd hh24:mi:ss') || '</首次时间><末次时间>' ||
               To_Char(末次时间_In, 'yyyy-mm-dd hh24:mi:ss') || '</末次时间><样本条码>' || 样本条码_In || '</样本条码></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_044', v_Value);
  End Zlhis_Cis_044;

  --患者医嘱执行登记
  Procedure Zlhis_Cis_050
  (
    病人id_In   In 病人医嘱记录.病人id%Type,
    主页id_In   In 病人医嘱记录.主页id%Type,
    挂号单_In   In 病人挂号记录.No%Type,
    发送号_In   In 病人医嘱发送.发送号%Type,
    医嘱id_In   In 病人医嘱记录.Id%Type,
    要求时间_In In 病人医嘱执行.要求时间%Type,
    执行时间_In In 病人医嘱执行.执行时间%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><挂号单>' || 挂号单_In || '</挂号单><发送号>' ||
               发送号_In || '</发送号><ID>' || 医嘱id_In || '</ID><要求时间>' || To_Char(要求时间_In, 'yyyy-mm-dd hh24:mi:ss') ||
               '</要求时间><执行时间>' || To_Char(执行时间_In, 'yyyy-mm-dd hh24:mi:ss') || '</执行时间></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_050', v_Value);
  End Zlhis_Cis_050;

  --患者医嘱取消执行登记
  Procedure Zlhis_Cis_051
  (
    病人id_In   In 病人医嘱记录.病人id%Type,
    主页id_In   In 病人医嘱记录.主页id%Type,
    挂号单_In   In 病人挂号记录.No%Type,
    发送号_In   In 病人医嘱发送.发送号%Type,
    医嘱id_In   In 病人医嘱记录.Id%Type,
    要求时间_In In 病人医嘱执行.要求时间%Type,
    执行时间_In In 病人医嘱执行.执行时间%Type,
    本次数次_In In 病人医嘱执行.本次数次%Type,
    执行结果_In In 病人医嘱执行.执行结果%Type,
    执行摘要_In In 病人医嘱执行.执行摘要%Type,
    执行科室_In In 病人医嘱执行.执行科室id%Type,
    执行人_In   In 病人医嘱执行.执行人%Type,
    核对人_In   In 病人医嘱执行.核对人%Type,
    记录来源_In In 病人医嘱执行.记录来源%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><挂号单>' || 挂号单_In || '</挂号单><发送号>' ||
               发送号_In || '</发送号><ID>' || 医嘱id_In || '</ID><要求时间>' || To_Char(要求时间_In, 'yyyy-mm-dd hh24:mi:ss') ||
               '</要求时间><执行时间>' || To_Char(执行时间_In, 'yyyy-mm-dd hh24:mi:ss') || '</执行时间><本次数次>' || 本次数次_In ||
               '</本次数次><执行结果>' || 执行结果_In || '</执行结果><执行摘要>' || 执行摘要_In || '</执行摘要><执行科室ID>' || 执行科室_In ||
               '</执行科室ID><执行人>' || 执行人_In || '</执行人><核对人>' || 核对人_In || '</核对人><记录来源>' || 记录来源_In || '</记录来源></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_051', v_Value);
  End Zlhis_Cis_051;
  --患者医嘱执行完成
  Procedure Zlhis_Cis_052
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    挂号单_In In 病人挂号记录.No%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_052',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><挂号单>' || 挂号单_In ||
                                 '</挂号单><发送号>' || 发送号_In || '</发送号><ID>' || 医嘱id_In || '</ID></root>');
  End Zlhis_Cis_052;
  --患者医嘱撤消执行完成
  Procedure Zlhis_Cis_053
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    挂号单_In In 病人挂号记录.No%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_053',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><挂号单>' || 挂号单_In ||
                                 '</挂号单><发送号>' || 发送号_In || '</发送号><ID>' || 医嘱id_In || '</ID></root>');
  End Zlhis_Cis_053;

  --病理申请发送后修改
  Procedure Zlhis_Cis_056
  (
    病人id_In   In 病人医嘱记录.病人id%Type,
    主页id_In   In 病人医嘱记录.主页id%Type,
    挂号单_In   In 病人挂号记录.No%Type,
    发送号_In   In 病人医嘱发送.发送号%Type,
    医嘱id_In   In 病人医嘱记录.Id%Type,
    病人来源_In In 病人医嘱记录.病人来源%Type --1-门诊;2-住院;3-外来(今后用于辅诊部门接收外来病人);4-体检病人
  ) Is
    v_Value    Zlmsg_Todo.Key_Value%Type;
    v_操作类型 诊疗项目目录.操作类型%Type;
  Begin
    Select Max(a.操作类型)
    Into v_操作类型
    From 诊疗项目目录 A, 病人医嘱记录 B
    Where b.诊疗项目id = a.Id And b.Id = 医嘱id_In;
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><挂号单>' || 挂号单_In || '</挂号单><发送号>' ||
               发送号_In || '</发送号><ID>' || 医嘱id_In || '</ID><病人来源>' || 病人来源_In || '</病人来源></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_056', v_Value);
  End Zlhis_Cis_056;

  --门诊患者完成就诊
  Procedure Zlhis_Cis_057
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    挂号单_In In 病人挂号记录.No%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_057', '<root><病人ID>' || 病人id_In || '</病人ID><NO>' || 挂号单_In || '</NO></root>');
  End Zlhis_Cis_057;

  --门诊患者取消完成就诊
  Procedure Zlhis_Cis_058
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    挂号单_In In 病人挂号记录.No%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_058', '<root><病人ID>' || 病人id_In || '</病人ID><NO>' || 挂号单_In || '</NO></root>');
  End Zlhis_Cis_058;

  --确认停止患者医嘱 
  Procedure Zlhis_Cis_059
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_059','<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><ID>' || 医嘱id_In || '</ID></root>');
  End Zlhis_Cis_059;

  --26.检查报告完成，检查完成时
  Procedure Zlhis_Pacs_001
  (
    医嘱id_In   In 影像检查记录.医嘱id%Type,
    报告id_Ins  In Varchar2,
    报告类型_In In Number --1-老版PACS报告，2-老版病历编辑器报告，3-新版编辑器报告
  ) Is
  Begin
    If p_Msg_Using('ZLHIS_PACS_001') = 0 Then
      Return;
    End If;
    For R In (Select '<root><医嘱ID>' || 医嘱id_In || '</医嘱ID><报告ID>' || Column_Value || '</报告ID><报告类型>' || 报告类型_In ||
                      '<报告类型></root>' As Xml_Value
              From Table(f_Str2list(报告id_Ins))) Loop
      b_Message.p_Msg_Todo_Insert('ZLHIS_PACS_001', r.Xml_Value);
    End Loop;
  End Zlhis_Pacs_001;
  --27.检查状态同步，检查状态改变后
  Procedure Zlhis_Pacs_002
  (
    医嘱id_In In 影像检查记录.医嘱id%Type,
    原状态_In In 病人医嘱发送.执行过程%Type,
    新状态_In In 病人医嘱发送.执行过程%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><医嘱ID>' || 医嘱id_In || '</医嘱ID><原状态>' || 原状态_In || '</原状态><新状态>' || 新状态_In || '</新状态></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PACS_002', v_Value);
  End Zlhis_Pacs_002;
  --28.检查状态回退，检查状态回退后
  Procedure Zlhis_Pacs_003
  (
    医嘱id_In In 影像检查记录.医嘱id%Type,
    原状态_In In 病人医嘱发送.执行过程%Type,
    新状态_In In 病人医嘱发送.执行过程%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><医嘱ID>' || 医嘱id_In || '</医嘱ID><原状态>' || 原状态_In || '</原状态><新状态>' || 新状态_In || '</新状态></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PACS_003', v_Value);
  End Zlhis_Pacs_003;
  --29.检查报告撤销，撤销检查完成时
  Procedure Zlhis_Pacs_004
  (
    医嘱id_In   In 影像检查记录.医嘱id%Type,
    报告id_Ins  In Varchar2,
    报告类型_In In Number --1-老版PACS报告，2-老版病历编辑器报告，3-新版编辑器报告
  ) Is
  Begin
    If p_Msg_Using('ZLHIS_PACS_004') = 0 Then
      Return;
    End If;
    For R In (Select '<root><医嘱ID>' || 医嘱id_In || '</医嘱ID><报告ID>' || Column_Value || '</报告ID><报告类型>' || 报告类型_In ||
                      '<报告类型></root>' As Xml_Value
              From Table(f_Str2list(报告id_Ins))) Loop
      b_Message.p_Msg_Todo_Insert('ZLHIS_PACS_004', r.Xml_Value);
    End Loop;
  End Zlhis_Pacs_004;
  --30.检查危急值通知，检查发生危急值时
  Procedure Zlhis_Pacs_005(医嘱id_In In 影像检查记录.医嘱id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><医嘱ID>' || 医嘱id_In || '</医嘱ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PACS_005', v_Value);
  End Zlhis_Pacs_005;
  -- 检查预约通知，检查预约时
  Procedure Zlhis_Pacs_006
  (
    医嘱id_In In 影像检查记录.医嘱id%Type,
    预约id_In In Ris检查预约.预约id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><医嘱ID>' || 医嘱id_In || '</医嘱ID><预约ID>' || 预约id_In || '</预约ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PACS_006', v_Value);
  End Zlhis_Pacs_006;
  -- 取消检查预约，取消预约时
  Procedure Zlhis_Pacs_007
  (
    医嘱id_In       In 影像检查记录.医嘱id%Type,
    预约id_In       In Ris检查预约.预约id%Type,
    预约日期_In     In Ris检查预约.预约日期%Type,
    预约序号_In     In Ris检查预约.序号%Type,
    检查设备名称_In In Ris检查预约.检查设备名称%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><医嘱ID>' || 医嘱id_In || '</医嘱ID><预约ID>' || 预约id_In || '</预约ID><预约日期>' || 预约日期_In || '</预约日期><预约序号>' ||
               预约序号_In || '</预约序号><检查设备名称>' || 检查设备名称_In || '</检查设备名称></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PACS_007', v_Value);
  End Zlhis_Pacs_007;

  --36.患者发卡或绑定卡
  Procedure Zlhis_Patient_018
  (
    变动id_In   In 病人变动记录.Id%Type,
    病人id_In   In 病人信息.病人id%Type,
    卡类别id_In In 医疗卡类别.Id%Type,
    卡号_In     In 病人医疗卡信息.卡号%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><变动ID>' || 变动id_In || '</变动ID><病人ID>' || 病人id_In || '</病人ID><卡类别ID>' || 卡类别id_In ||
               '</卡类别ID><卡号>' || 卡号_In || '</卡号></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_018', v_Value);
  End;

  --37.患者退卡
  Procedure Zlhis_Patient_019
  (
    变动id_In   In 病人变动记录.Id%Type,
    病人id_In   In 病人信息.病人id%Type,
    卡类别id_In In 医疗卡类别.Id%Type,
    卡号_In     In 病人医疗卡信息.卡号%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><变动ID>' || 变动id_In || '</变动ID><病人ID>' || 病人id_In || '</病人ID><卡类别ID>' || 卡类别id_In ||
               '</卡类别ID><卡号>' || 卡号_In || '</卡号></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_019', v_Value);
  End;

  --38.患者补卡/换卡
  Procedure Zlhis_Patient_020
  (
    变动id_In   In 病人变动记录.Id%Type,
    病人id_In   In 病人信息.病人id%Type,
    卡类别id_In In 医疗卡类别.Id%Type,
    原卡号_In   In 病人医疗卡信息.卡号%Type,
    新卡号_In   In 病人医疗卡信息.卡号%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><变动ID>' || 变动id_In || '</变动ID><病人ID>' || 病人id_In || '</病人ID><卡类别ID>' || 卡类别id_In ||
               '</卡类别ID><原卡号>' || 原卡号_In || '</原卡号><新卡号>' || 新卡号_In || '</新卡号></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_020', v_Value);
  End;

  --39.病人挂号登记（包含预约登记)
  Procedure Zlhis_Regist_001
  (
    挂号id_In In 病人挂号记录.Id%Type,
    No_In     In 病人挂号记录.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><挂号ID>' || 挂号id_In || '</挂号ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_REGIST_001', v_Value);
  End;

  --40.病人分诊
  Procedure Zlhis_Regist_002
  (
    挂号id_In In 病人挂号记录.Id%Type,
    No_In     In 病人挂号记录.No%Type,
    诊室_In   In 病人挂号记录.诊室%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><挂号ID>' || 挂号id_In || '</挂号ID><NO>' || No_In || '</NO><诊室>' || Nvl(诊室_In, '') || '</诊室></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_REGIST_002', v_Value);
  End;

  --41.病人退号（含取消预约)
  Procedure Zlhis_Regist_003
  (
    挂号id_In In 病人挂号记录.Id%Type,
    No_In     In 病人挂号记录.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><挂号ID>' || 挂号id_In || '</挂号ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_REGIST_003', v_Value);
  End;

  --42.临床出诊安排调整
  Procedure Zlhis_Regist_004
  (
    变动原因_In In Integer, --1-停诊;2-替诊;3-诊室变动
    记录id_In   In 临床出诊记录.Id%Type,
    变动id_In   In 临床出诊变动记录.Id%Type
    
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><变动原因>' || 变动原因_In || '</变动原因><记录ID>' || 记录id_In || '</记录ID><变动ID>' || 变动id_In ||
               '</变动ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_REGIST_004', v_Value);
  End;

  --43.门诊患者挂号换号操作
  Procedure Zlhis_Regist_005
  (
    No_In         In 病人挂号记录.No%Type,
    变动原因_In   Integer, --1-替诊;2-换号;3-预约日期变动,
    就诊变动id_In 就诊变动记录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><NO>' || No_In || '</NO><变动原因>' || 变动原因_In || '</变动原因><就诊变动ID>' || 就诊变动id_In ||
               '</就诊变动ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_REGIST_005', v_Value);
  End;

  --费用门诊收费及补充结算
  Procedure Zlhis_Charge_002
  (
    结算类型_In In Number,
    结帐id_In   In 门诊费用记录.结帐id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    --结算类型_In:1-收费结算，2-补充结算
    v_Value := '<root><结算类型>' || 结算类型_In || '</结算类型><结帐ID>' || 结帐id_In || '</结帐ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CHARGE_002', v_Value);
  End;

  --46.门诊退费单据
  Procedure Zlhis_Charge_004
  (
    退费类型_In In Number,
    结帐id_In   In 门诊费用记录.结帐id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    --退费类型_In:1-收费结算，2-补充结算
    v_Value := '<root><退费类型>' || 退费类型_In || '</退费类型><结帐ID>' || 结帐id_In || '</结帐ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CHARGE_004', v_Value);
  End;

  --47.收预交款
  Procedure Zlhis_Charge_005
  (
    预交id_In In 病人预交记录.Id%Type,
    单据号_In In 病人预交记录.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><预交ID>' || 预交id_In || '</预交ID><单据号>' || 单据号_In || '</单据号></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CHARGE_005', v_Value);
  End;

  --48.退预交款(包含负数退预交款部分)
  Procedure Zlhis_Charge_006
  (
    退预交id_In In 病人预交记录.Id%Type,
    单据号_In   In 病人预交记录.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><退预交ID>' || 退预交id_In || '</退预交ID><单据号>' || 单据号_In || '</单据号></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CHARGE_006', v_Value);
  End;

  --住院记帐单据
  Procedure Zlhis_Charge_007
  (
    收费类别_In In 住院费用记录.收费类别%Type,
    费用id_In   In 住院费用记录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><收费类别>' || 收费类别_In || '</收费类别><费用ID>' || 费用id_In || '</费用ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CHARGE_007', v_Value);
  End;

  --住院记帐单据销账
  Procedure Zlhis_Charge_008
  (
    收费类别_In In 住院费用记录.收费类别%Type,
    费用id_In   In 住院费用记录.Id%Type,
    收发ids_In  In Varchar2 := Null --可能费用ID对应多个收发id，对应格式：收发id,数量|收发id,数量；非药品不传
  ) Is
    v_Value   Zlmsg_Todo.Key_Value%Type;
    v_Tmp     Varchar2(4000);
    v_Infotmp Varchar2(4000);
    v_Fields  Varchar2(4000);
    v_收发id  Varchar2(50);
    v_数量    Varchar2(20);
  Begin
    If p_Msg_Using('ZLHIS_CHARGE_008') = 0 Then
      Return;
    End If;
    v_Value := '<root><收费类别>' || 收费类别_In || '</收费类别><费用ID>' || 费用id_In || '</费用ID>';
  
    If 收发ids_In Is Null Then
      v_Infotmp := Null;
      v_Tmp     := '<收发IDS>' || '<收发ID>' || '</收发ID>' || '<数量>' || '</数量>' || '</收发IDS>';
    Else
      v_Infotmp := 收发ids_In || '|';
      While v_Infotmp Is Not Null Loop
        --分解收发ID串
        v_Fields  := Substr(v_Infotmp, 1, Instr(v_Infotmp, '|') - 1);
        v_收发id  := Substr(v_Fields, 1, Instr(v_Fields, ',') - 1);
        v_数量    := Substr(v_Fields, Instr(v_Fields, ',') + 1);
        v_Infotmp := Replace('|' || v_Infotmp, '|' || v_Fields || '|');
      
        v_Tmp := v_Tmp || '<收发IDS>' || '<收发ID>' || v_收发id || '</收发ID>' || '<数量>' || v_数量 || '</数量>' || '</收发IDS>';
      End Loop;
    End If;
  
    v_Value := v_Value || v_Tmp || '</root>';
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_CHARGE_008', v_Value);
  End;

  --53.住院患者入院登记
  Procedure Zlhis_Patient_001
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  ) Is
    n_变动id 病人变动记录.Id%Type;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_001') = 0 Then
      Return;
    End If;
    Select Max(ID)
    Into n_变动id
    From 病人变动记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 开始原因 = 1 And Nvl(附加床位, 0) = 0;
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_001',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><变动ID>' || n_变动id ||
                                 '</变动ID></root>');
  End Zlhis_Patient_001;
  --54.住院患者入院入科
  Procedure Zlhis_Patient_002
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  ) Is
    n_变动id 病人变动记录.Id%Type;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_002') = 0 Then
      Return;
    End If;
    Select Max(ID)
    Into n_变动id
    From 病人变动记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 Is Null And Nvl(附加床位, 0) = 0;
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_002',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><变动ID>' || n_变动id ||
                                 '</变动ID></root>');
  End Zlhis_Patient_002;
  --56.住院患者床位变更
  Procedure Zlhis_Patient_004
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  ) Is
    v_原床号   Varchar2(255);
    v_新床号   Varchar2(255);
    n_变动id   Number(18);
    n_开始原因 Number(3);
    d_开始时间 Date;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_004') = 0 Then
      Return;
    End If;
    Select ID, 床号, 开始时间, 开始原因
    Into n_变动id, v_新床号, d_开始时间, n_开始原因
    From 病人变动记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 Is Null And Nvl(附加床位, 0) = 0;
  
    Select Max(床号)
    Into v_原床号
    From 病人变动记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 = d_开始时间 And 终止原因 = n_开始原因 And Nvl(附加床位, 0) = 0;
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_004',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID>' || '<原床号>' ||
                                 v_原床号 || '</原床号>' || '<新床号>' || v_新床号 || '</新床号>' || '<变动ID>' || n_变动id || '</变动ID>' ||
                                 '</root>');
  End Zlhis_Patient_004;
  --57.住院患者病情变更
  Procedure Zlhis_Patient_005
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  ) Is
    n_变动id 病人变动记录.Id%Type;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_005') = 0 Then
      Return;
    End If;
    Select Max(ID)
    Into n_变动id
    From 病人变动记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 Is Null And Nvl(附加床位, 0) = 0;
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_005',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><变动ID>' || n_变动id ||
                                 '</变动ID></root>');
  End Zlhis_Patient_005;
  --58.住院患者变更撤消
  Procedure Zlhis_Patient_006
  (
    病人id_In   In 病案主页.病人id%Type,
    主页id_In   In 病案主页.主页id%Type,
    撤销方式_In In Varchar2
  ) Is
    n_科室id     病人变动记录.科室id%Type;
    n_病区id     病人变动记录.病区id%Type;
    n_护理等级id 病人变动记录.护理等级id%Type;
    n_医疗小组id 病人变动记录.医疗小组id%Type;
    v_床号       病人变动记录.床号%Type;
    v_责任护士   病人变动记录.责任护士%Type;
    v_主任医师   病人变动记录.主任医师%Type;
    v_主治医师   病人变动记录.主治医师%Type;
    v_经治医师   病人变动记录.经治医师%Type;
    v_病情       病人变动记录.病情%Type;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_006') = 0 Then
      Return;
    End If;
    Select Max(科室id), Max(病区id), Max(护理等级id), Max(医疗小组id), Max(床号), Max(责任护士), Max(主任医师), Max(主治医师), Max(经治医师), Max(病情)
    Into n_科室id, n_病区id, n_护理等级id, n_医疗小组id, v_床号, v_责任护士, v_主任医师, v_主治医师, v_经治医师, v_病情
    From 病人变动记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And (终止时间 Is Null Or 终止原因 = 1) And Nvl(附加床位, 0) = 0;
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_006',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><撤销方式>' || 撤销方式_In ||
                                 '</撤销方式><科室ID>' || n_科室id || '</科室ID>' || '<病区ID>' || n_病区id || '</病区ID>' || '<护理等级ID>' ||
                                 n_护理等级id || '</护理等级ID>' || '<医疗小组ID>' || n_医疗小组id || '</医疗小组ID>' || '<床号>' || v_床号 ||
                                 '</床号>' || '<责任护士>' || v_责任护士 || '</责任护士>' || '<主任医师>' || v_主任医师 || '</主任医师>' ||
                                 '<主治医师>' || v_主治医师 || '</主治医师>' || '<经治医师>' || v_经治医师 || '</经治医师>' || '<病情>' || v_病情 ||
                                 '</病情>' || '</root>');
  End Zlhis_Patient_006;
  --59.住院患者医护变更
  Procedure Zlhis_Patient_007
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  ) Is
    v_原住院医生 Varchar2(100);
    v_新住院医生 Varchar2(100);
    v_原主治医生 Varchar2(100);
    v_新主治医生 Varchar2(100);
    v_原主任医生 Varchar2(100);
    v_新主任医生 Varchar2(100);
    v_原责任护士 Varchar2(100);
    v_新责任护士 Varchar2(100);
    n_变动id     Number(18);
    n_开始原因   Number(3);
    d_开始时间   Date;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_007') = 0 Then
      Return;
    End If;
    Select ID, 经治医师, 主治医师, 主任医师, 责任护士, 开始时间, 开始原因
    Into n_变动id, v_新住院医生, v_新主治医生, v_新主任医生, v_新责任护士, d_开始时间, n_开始原因
    From 病人变动记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 Is Null And Nvl(附加床位, 0) = 0;
  
    Select Max(经治医师), Max(主治医师), Max(主任医师), Max(责任护士)
    Into v_原住院医生, v_原主治医生, v_原主任医生, v_原责任护士
    From 病人变动记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 = d_开始时间 And 终止原因 = n_开始原因 And Nvl(附加床位, 0) = 0;
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_007',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID>' || '<原住院医生>' ||
                                 v_原住院医生 || '</原住院医生>' || '<新住院医生>' || v_新住院医生 || '</新住院医生>' || '<原主治医生>' || v_原主治医生 ||
                                 '</原主治医生>' || '<新主治医生>' || v_新主治医生 || '</新主治医生>' || '<原主任医生>' || v_原主任医生 || '</原主任医生>' ||
                                 '<新主任医生>' || v_新主任医生 || '</新主任医生>' || '<原责任护士>' || v_原责任护士 || '</原责任护士>' || '<新责任护士>' ||
                                 v_新责任护士 || '</新责任护士>' || '<变动ID>' || n_变动id || '</变动ID>' || '</root>');
  End Zlhis_Patient_007;
  --住院患者护理等级变更
  Procedure Zlhis_Patient_008
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  ) Is
    v_原护理等级id Number(18);
    v_新护理等级id Number(18);
    n_变动id       Number(18);
    n_开始原因     Number(3);
    d_开始时间     Date;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_008') = 0 Then
      Return;
    End If;
    Select ID, 护理等级id, 开始时间, 开始原因
    Into n_变动id, v_新护理等级id, d_开始时间, n_开始原因
    From 病人变动记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 Is Null And Nvl(附加床位, 0) = 0;
  
    Select Max(护理等级id)
    Into v_原护理等级id
    From 病人变动记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 = d_开始时间 And 终止原因 = n_开始原因 And Nvl(附加床位, 0) = 0;
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_008',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID>' || '<原护理等级ID>' ||
                                 v_原护理等级id || '</原护理等级ID>' || '<新护理等级ID>' || v_新护理等级id || '</新护理等级ID>' || '<变动ID>' ||
                                 n_变动id || '</变动ID>' || '</root>');
  End Zlhis_Patient_008;
  --60.住院患者预出院
  Procedure Zlhis_Patient_009
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  ) Is
    n_变动id 病人变动记录.Id%Type;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_009') = 0 Then
      Return;
    End If;
    Select Max(ID)
    Into n_变动id
    From 病人变动记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 Is Null And Nvl(附加床位, 0) = 0;
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_009',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><变动ID>' || n_变动id ||
                                 '</变动ID></root>');
  End Zlhis_Patient_009;
  --61.住院患者出院
  Procedure Zlhis_Patient_010
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_010',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID></root>');
  End Zlhis_Patient_010;
  --62.住院患者新生儿登记
  Procedure Zlhis_Patient_011
  (
    病人id_In   In 病案主页.病人id%Type,
    主页id_In   In 病案主页.主页id%Type,
    婴儿序号_In 病人医嘱记录.婴儿%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_011',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><婴儿序号>' || 婴儿序号_In ||
                                 '</婴儿序号></root>');
  End Zlhis_Patient_011;
  --63.住院患者转入科室
  Procedure Zlhis_Patient_012
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  ) Is
    v_转出科室id Number(18);
    v_转入科室id Number(18);
    n_变动id     Number(18);
    n_开始原因   Number(3);
    d_开始时间   Date;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_012') = 0 Then
      Return;
    End If;
    Select ID, 科室id, 开始时间, 开始原因
    Into n_变动id, v_转入科室id, d_开始时间, n_开始原因
    From 病人变动记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 Is Null And Nvl(附加床位, 0) = 0;
  
    Select Max(科室id)
    Into v_转出科室id
    From 病人变动记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 = d_开始时间 And 终止原因 = n_开始原因 And Nvl(附加床位, 0) = 0;
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_012',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID>' || '<转出科室ID>' ||
                                 v_转出科室id || '</转出科室ID>' || '<转入科室ID>' || v_转入科室id || '</转入科室ID>' || '<变动ID>' || n_变动id ||
                                 '</变动ID>' || '</root>');
  End Zlhis_Patient_012;
  --64.新生儿登记作废
  Procedure Zlhis_Patient_013
  (
    病人id_In   In 病案主页.病人id%Type,
    主页id_In   In 病案主页.主页id%Type,
    婴儿序号_In 病人医嘱记录.婴儿%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_013',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><婴儿序号>' || 婴儿序号_In ||
                                 '</婴儿序号></root>');
  End Zlhis_Patient_013;
  --65.门诊患者登记
  Procedure Zlhis_Patient_015(病人id_In In 病案主页.病人id%Type) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_015', '<root><病人ID>' || 病人id_In || '</病人ID></root>');
  End Zlhis_Patient_015;
  --66.患者信息修改
  Procedure Zlhis_Patient_016(病人id_In In 病案主页.病人id%Type) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_016', '<root><病人ID>' || 病人id_In || '</病人ID></root>');
  End Zlhis_Patient_016;

  --67.患者合并
  Procedure Zlhis_Patient_017 
  ( 
    病人id_In   In 病案主页.病人id%Type, 
    原病人id_In In 病案主页.病人id%Type,
    变化ids_In  In Varchar2
  ) Is 
  --参数： 1病人id,1主页id:1原病人id,1原主页id; 2病人id,2主页id:2原病人id,2原主页id;….
  Begin 
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_017', 
                                '<root><病人ID>' || 病人id_In || '</病人ID><原病人ID>' || 原病人id_In || '</原病人ID><CINFO>'||变化ids_In||'</CINFO></root>'); 
  End Zlhis_Patient_017;

  --69.住院患者转入病区
  Procedure Zlhis_Patient_026
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  ) Is
    v_转出病区id Number(18);
    v_转入病区id Number(18);
    n_变动id     Number(18);
    n_开始原因   Number(3);
    d_开始时间   Date;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_026') = 0 Then
      Return;
    End If;
    Select ID, 病区id, 开始时间, 开始原因
    Into n_变动id, v_转入病区id, d_开始时间, n_开始原因
    From 病人变动记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 Is Null And Nvl(附加床位, 0) = 0;
  
    Select Max(病区id)
    Into v_转出病区id
    From 病人变动记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 = d_开始时间 And 终止原因 = n_开始原因 And Nvl(附加床位, 0) = 0;
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_026',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID>' || '<转出病区ID>' ||
                                 v_转出病区id || '</转出病区ID>' || '<转入病区ID>' || v_转入病区id || '</转入病区ID>' || '<变动ID>' || n_变动id ||
                                 '</变动ID>' || '</root>');
  End Zlhis_Patient_026;

  Procedure Zlhis_Patient_028(病人id_In In 病案主页.病人id%Type) Is 
    v_姓名     病人信息.姓名%Type; 
    v_性别     病人信息.性别%Type; 
    v_年龄     病人信息.年龄%Type; 
    v_门诊号   病人信息.门诊号%Type; 
    v_身份证号 病人信息.身份证号%Type; 
    v_出生日期 varchar2(50); 
  Begin 
    If p_Msg_Using('ZLHIS_PATIENT_028') = 0 Then 
      Return; 
    End If; 
    Select 姓名, 性别, 年龄, To_Char(出生日期, 'yyyymmdd'), 门诊号, 身份证号 
    Into v_姓名, v_性别, v_年龄, v_出生日期, v_门诊号, v_身份证号 
    From 病人信息 
    Where 病人id = 病人id_In; 
 
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_028', 
                                '<root><病人ID>' || 病人id_In || '</病人ID><姓名>' || v_姓名 || '</姓名>' || '<性别>' || v_性别 || 
                                 '</性别>' || '<年龄>' || v_年龄 || '</年龄>' || '<出生日期>' || v_出生日期 || '</出生日期>' || '<门诊号>' || 
                                 v_门诊号 || '</门诊号>' || '<身份证号>' || v_身份证号 || '</身份证号>' || '</root>'); 
  End Zlhis_Patient_028; 

  --血库:科室配血完成
  Procedure Zlhis_Blood_001(医嘱id_In In 病人医嘱记录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If 医嘱id_In Is Not Null Then
      v_Value := '<root><医嘱ID>' || 医嘱id_In || '</医嘱ID></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_BLOOD_001', v_Value);
    End If;
  End Zlhis_Blood_001;

  --血库:科室拒绝配血
  Procedure Zlhis_Blood_002(医嘱id_In In 病人医嘱记录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If 医嘱id_In Is Not Null Then
      v_Value := '<root><医嘱ID>' || 医嘱id_In || '</医嘱ID></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_BLOOD_002', v_Value);
    End If;
  End Zlhis_Blood_002;

  --70.检验报告审核
  Procedure Zlhis_Lis_001(标本id_In In 检验标本记录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If 标本id_In Is Not Null Then
      v_Value := '<root><标本ID>' || 标本id_In || '</标本ID><系统>1</系统></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_001', v_Value);
    End If;
  End Zlhis_Lis_001;
  --71.检验报告审核撤消
  Procedure Zlhis_Lis_002(标本id_In In 检验标本记录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If 标本id_In Is Not Null Then
      v_Value := '<root><标本ID>' || 标本id_In || '</标本ID><系统>1</系统></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_002', v_Value);
    End If;
  End Zlhis_Lis_002;
  --73.检验标本条码打印
  Procedure Zlhis_Lis_004
  (
    样本条码_In In 病人医嘱发送.样本条码%Type,
    医嘱id_In   In 病人医嘱发送.医嘱id%Type,
    医嘱ids_In  In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If p_Msg_Using('ZLHIS_LIS_004') = 0 Then
      Return;
    End If;
    If 医嘱id_In Is Not Null Then
      v_Value := '<root><样本条码>' || 样本条码_In || '</样本条码><医嘱ID>' || 医嘱id_In || '</医嘱ID><系统>1</系统></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_004', v_Value);
    Else
      For R In (Select '<root><样本条码>' || 样本条码_In || '</样本条码><医嘱ID>' || 医嘱id_In || '</医嘱ID><系统>1</系统></root>' As Xml_Value
                From 病人医嘱发送
                Where 医嘱id In (Select Column_Value From Table(f_Num2list(医嘱ids_In)))) Loop
        b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_004', r.Xml_Value);
      End Loop;
    End If;
  End Zlhis_Lis_004;
  --74.检验标本条码打印撤销
  Procedure Zlhis_Lis_005
  (
    样本条码_In In 病人医嘱发送.样本条码%Type,
    医嘱id_In   In 病人医嘱发送.医嘱id%Type,
    医嘱ids_In  In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If p_Msg_Using('ZLHIS_LIS_005') = 0 Then
      Return;
    End If;
    If 医嘱id_In Is Not Null Then
      v_Value := '<root><样本条码>' || 样本条码_In || '</样本条码><医嘱ID>' || 医嘱id_In || '</医嘱ID><系统>1</系统></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_005', v_Value);
    Else
      For R In (Select '<root><样本条码>' || 样本条码_In || '</样本条码><医嘱ID>' || 医嘱id_In || '</医嘱ID><系统>1</系统></root>' As Xml_Value
                From 病人医嘱发送
                Where 医嘱id In (Select Column_Value From Table(f_Num2list(医嘱ids_In)))) Loop
        b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_005', r.Xml_Value);
      End Loop;
    End If;
  End Zlhis_Lis_005;
  --75.检验标本核收
  Procedure Zlhis_Lis_006(标本id_In In 检验标本记录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If 标本id_In Is Not Null Then
      v_Value := '<root><标本ID>' || 标本id_In || '</标本ID><系统>1</系统></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_006', v_Value);
    End If;
  End Zlhis_Lis_006;
  --76.检验标本核收撤销
  Procedure Zlhis_Lis_007(标本id_In In 检验标本记录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If 标本id_In Is Not Null Then
      v_Value := '<root><标本ID>' || 标本id_In || '</标本ID><系统>1</系统></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_007', v_Value);
    End If;
  End Zlhis_Lis_007;
  --77.检验标本拒收
  Procedure Zlhis_Lis_008(标本id_In In 检验标本记录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If 标本id_In Is Not Null Then
      v_Value := '<root><标本ID>' || 标本id_In || '</标本ID><系统>1</系统></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_008', v_Value);
    End If;
  End Zlhis_Lis_008;

End b_Message;
/

--137701:胡俊勇,2019-02-12,病人信息合并消息锚点修改
Create Or Replace Procedure Zl_病人信息_Merge
(
  A病人id_In    病人信息.病人id%Type, --要合并的病人信息
  B病人id_In    病人信息.病人id%Type, --要保留的病人信息
  合并原因_In   病人合并记录.合并原因%Type,
  操作员姓名_In 人员表.姓名%Type,
  强制保留_In   Number := 0
  --标准版
  ----------------------------------------------------------------------------
  --病人信息,病案主页,病案主页从表,病人变动记录,特殊病人
  --门诊病案记录,住院病案记录,床位状况记录
  --医保病人档案,保险模拟结算,保险结算记录,帐户年度信息
  --病人余额,病人未结费用,住院费用记录,门诊费用记录,病人预交记录,病人结帐记录,未发药品记录
  --病人挂号记录,病人过敏药物,病人过敏记录,病人诊断记录,诊断情况
  --病人医嘱记录,病人手麻记录
  --病人社区信息
  
  --后备表：
  --H病人结帐记录,H病人预交记录,H住院费用记录,H门诊费用记录
  --H病人医嘱记录,H病人诊断记录,H病人过敏记录
  --H病人病历记录,H病人手麻记录
  
  --病案系统
  ----------------------------------------------------------------------------
  --病人费用,随诊记录,借阅记录
  --新生儿诊断记录,病人分娩信息
  --诊断符合情况,病案评分结果
  
) As
  --病人相关表
  Cursor c_Patitable Is
    Select a.Table_Name, Max(Decode(b.Column_Name, '病人ID', 1, 0)) As 病人id,
           Max(Decode(b.Column_Name, '主页ID', 1, 0)) As 主页id
    From User_Tables A, User_Tab_Columns B
    Where a.Table_Name = b.Table_Name And b.Column_Name In ('病人ID', '主页ID') And
          a.Table_Name Not In
          ('病人信息', '病案主页', '病案主页从表', '病人变动记录','病人自动计算', '特殊病人', '门诊病案记录', '住院病案记录', '床位状况记录', '医保病人档案', '医保病人关联表', '保险模拟结算',
           '帐户年度信息', '病人余额', '病人未结费用', '住院费用记录', '门诊费用记录', '病人预交记录', '病人结帐记录', '未发药品记录', '病人挂号记录', '病人过敏药物', '病人过敏记录',
           '病人诊断记录', '诊断情况', '病人医嘱记录', '病人手麻记录', '病人费用', '随诊记录', '借阅记录', '病人分娩信息', '诊断符合情况', '病案评分结果', '病人担保记录', '病人社区信息',
           '病人免疫记录', '病人信息从表', '病人医疗卡属性') Having Max(Decode(b.Column_Name, '病人ID', 1, 0)) <> 0
    Group By a.Table_Name;

  --数组定义
  Type Array_Patitable Is Table Of Varchar2(100) Index By Binary_Integer;
  Arronbase Array_Patitable;
  Arronpage Array_Patitable;
  v_Loop    Number;
  n_Have    Number;

  -------------------------------------------------------
  --被合并的病人(住院号可能每次新产生,多次住院取最近一次)
  Cursor c_Infoa Is
    Select a.病人id, a.门诊号, a.住院号, a.就诊卡号, a.卡验证码, a.费别, a.医疗付款方式, a.姓名, a.性别, a.年龄, a.出生日期, a.出生地点, a.身份证号, a.其他证件, a.身份,
           a.职业, a.民族, a.国籍, a.籍贯, a.区域, a.学历, a.婚姻状况, a.家庭地址, a.家庭电话, a.家庭地址邮编, a.监护人, a.联系人姓名, a.联系人关系, a.联系人地址,
           a.联系人电话, a.户口地址, a.户口地址邮编, a.Email, a.Qq, a.合同单位id, a.工作单位, a.单位电话, a.单位邮编, a.单位开户行, a.单位帐号, a.担保人, a.担保额,
           a.担保性质, a.就诊时间, a.就诊状态, a.就诊诊室, a.住院次数, a.当前科室id, a.当前病区id, a.当前床号, a.入院时间, a.出院时间, a.在院, a.Ic卡号, a.健康号,
           a.医保号, a.险类, a.查询密码, a.登记时间, a.停用时间, a.锁定, a.联系人身份证号, b.主页id, b.入院日期, b.出院日期
    From 病人信息 A, 病案主页 B
    Where a.病人id = b.病人id(+) And a.病人id = A病人id_In
    Order By 出院日期 Desc, 主页id Desc;
  r_Infoa c_Infoa%RowType;

  --要保留的病人(住院号可能每次新产生,多次住院取最近一次)
  Cursor c_Infob Is
    Select a.病人id, a.门诊号, a.住院号, a.就诊卡号, a.卡验证码, a.费别, a.医疗付款方式, a.姓名, a.性别, a.年龄, a.出生日期, a.出生地点, a.身份证号, a.其他证件, a.身份,
           a.职业, a.民族, a.国籍, a.籍贯, a.区域, a.学历, a.婚姻状况, a.家庭地址, a.家庭电话, a.家庭地址邮编, a.监护人, a.联系人姓名, a.联系人关系, a.联系人地址,
           a.联系人电话, a.户口地址, a.户口地址邮编, a.Email, a.Qq, a.合同单位id, a.工作单位, a.单位电话, a.单位邮编, a.单位开户行, a.单位帐号, a.担保人, a.担保额,
           a.担保性质, a.就诊时间, a.就诊状态, a.就诊诊室, a.住院次数, a.当前科室id, a.当前病区id, a.当前床号, a.入院时间, a.出院时间, a.在院, a.Ic卡号, a.健康号,
           a.医保号, a.险类, a.查询密码, a.登记时间, a.停用时间, a.锁定, a.联系人身份证号, b.主页id, b.入院日期, b.出院日期
    From 病人信息 A, 病案主页 B
    Where a.病人id = b.病人id(+) And a.病人id = B病人id_In
    Order By 出院日期 Desc, 主页id Desc;
  r_Infob c_Infob%RowType;

  --合并后的信息
  Cursor c_Info(v_病人id 病人信息.病人id%Type) Is
    Select 病人id, 主页id, (Select Nvl(Max(主页id), 0) From 病案主页 Where 病人id = v_病人id) 最大主页id, 住院号, 病人性质, 医疗付款方式, 费别, 再入院,
           入院病区id, 入院科室id, 医疗小组id, 入院日期, 入院病况, 入院方式, 入院属性, 二级院转入, 住院目的, 入院病床, 是否陪伴, 当前病况, 当前病区id, 护理等级id, 出院科室id, 出院病床,
           出院日期, 住院天数, 出院方式, 是否确诊, 确诊日期, 新发肿瘤, 血型, 抢救次数, 成功次数, 随诊标志, 随诊期限, 尸检标志, 门诊医师, 责任护士, 住院医师, 病案号, 编目员编号, 编目员姓名,
           编目日期, 状态, 费用和, 年龄, 身高, 体重, 婚姻状况, 职业, 国籍, 学历, 单位电话, 单位邮编, 单位地址, 区域, 家庭地址, 家庭电话, 家庭地址邮编, 联系人姓名, 联系人关系, 联系人地址,
           联系人电话, 联系人身份证号, 户口地址, 户口地址邮编, 中医治疗类别, 险类, 社区, 审核标志, 审核人, 审核日期, 是否上传, 数据转出, 登记人, 登记时间, 备注, 病案状态, 病人类型
    From 病案主页
    Where 主页id = (Select Nvl(Max(主页id), 0)
                  From 病案主页
                  Where 病人id = v_病人id And Not Exists (Select 主页id From 病案主页 Where 病人id = v_病人id And 主页id = 0)) And
          病人id = v_病人id;
  r_Info c_Info%RowType;

  --合并两个住院病人
  Cursor c_Mergepati Is
    Select a.姓名, a.门诊号, a.住院号 当前住院号, b.病人id, b.主页id, b.住院号, b.留观号, b.病人性质, b.医疗付款方式, b.费别, b.再入院, b.入院病区id, b.入院科室id,
           b.医疗小组id, b.入院日期, b.入院病况, b.入院方式, b.入院属性, b.二级院转入, b.住院目的, b.入院病床, b.是否陪伴, b.当前病况, b.当前病区id, b.护理等级id,
           b.出院科室id, b.出院病床, b.出院日期, b.住院天数, b.出院方式, b.是否确诊, b.确诊日期, b.新发肿瘤, b.血型, b.抢救次数, b.成功次数, b.随诊标志, b.随诊期限,
           b.尸检标志, b.门诊医师, b.责任护士, b.住院医师, b.病案号, b.编目员编号, b.编目员姓名, b.编目日期, b.状态, b.费用和, b.性别, b.年龄, b.身高, b.体重, b.婚姻状况,
           b.职业, b.国籍, b.学历, b.单位电话, b.单位邮编, b.单位地址, b.区域, b.家庭地址, b.家庭电话, b.家庭地址邮编, b.联系人姓名, b.联系人关系, b.联系人地址, b.联系人电话,
           b.联系人身份证号, b.户口地址, b.户口地址邮编, b.中医治疗类别, b.险类, b.社区, b.审核标志, b.审核人, b.审核日期, b.是否上传, b.数据转出, b.登记人, b.登记时间, b.备注,
           b.病案状态, b.病人类型, b.封存时间, b.路径状态, b.单病种, b.婴儿科室id, b.婴儿病区id, b.母婴转科标志, b.医嘱重整时间
    From 病人信息 A, 病案主页 B
    Where a.病人id = b.病人id And a.病人id In (A病人id_In, B病人id_In)
    Order By b.入院日期 Desc, Nvl(b.出院日期, Sysdate) Desc;

  v_保留id 病人信息.病人id%Type;
  v_合并id 病人信息.病人id%Type;
  v_门诊号 病人信息.门诊号%Type;
  v_住院号 病人信息.住院号%Type;
  --病人未结费用(门诊部份)
  Cursor c_Owe(v_病人id 病人信息.病人id%Type) Is
    Select 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, Sum(金额) As 金额
    From 病人未结费用
    Where 主页id Is Null And 病人id = v_病人id
    Group By 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径;

  --病人余额
  Cursor c_Spare(v_病人id 病人信息.病人id%Type) Is
    Select 性质, 类型, 预交余额, 费用余额 From 病人余额 Where 病人id = v_病人id;

  --医保病人档案
  Cursor c_Insure(v_病人id 病人信息.病人id%Type) Is
    Select * From 保险帐户 Where 病人id = v_病人id Order By 险类;

  --要保留的医保病人档案
  Cursor c_Keepinsure
  (
    v_病人id 病人信息.病人id%Type,
    v_险类   医保病人档案.险类%Type
  ) Is
    Select * From 保险帐户 Where 病人id = v_病人id And 险类 = v_险类;
  r_Keepinsure c_Keepinsure%RowType;

  Cursor c_Year
  (
    v_病人id 病人信息.病人id%Type,
    v_险类   医保病人档案.险类%Type
  ) Is
    Select * From 帐户年度信息 Where 病人id = v_病人id And 险类 = v_险类;

  v_原信息   病人合并记录.原信息%Type;
  v_Count    Number;
  n_Readonly Number;
  v_Sql      Varchar2(1000);

  n_主页id       病人信息.主页id%Type;
  v_Error        Varchar2(255);
  n_担保额       病人担保记录.担保额%Type;
  v_担保人       病人信息.担保人%Type;
  n_担保性质     病人担保记录.担保性质%Type;
  n_Row          Number;
  n_独立病案     Number;
  n_每次新住院号 Number;
  n_Max主页id    Number;
  n_Cnt主页id    Number;
  n_Cur主页id    Number;
  n_Cnt住院次数  Number;
  n_Cur住院次数  Number;
  n_Max住院次数  Number;
  n_Loop主页id   病人信息.主页id%Type;
  v_Chgs         Varchar2(4000);

  n_Lengthb Number;
  Err_Custom Exception;
Begin
  Begin
    Select 只读 Into n_Readonly From zlBakSpaces Where 当前 = 1;
  Exception
    When Others Then
      Null;
  End;
  If n_Readonly = 1 Then
    n_Readonly := 0;
    For r_Bak In (Select a.表名 Table_Name
                  From Zltools.Zlbaktables A, User_Constraints B
                  Where a.表名 = b.Table_Name And b.r_Constraint_Name = '病人信息_PK' And b.Constraint_Type = 'R') Loop
      v_Sql := 'Select Count(病人Id) From H' || r_Bak.Table_Name || ' Where 病人Id In(:1,:2)';
      Execute Immediate v_Sql
        Into n_Readonly
        Using A病人id_In, B病人id_In;
      If n_Readonly > 0 Then
        v_Error := '病人在只读的当前转储空间存在数据,不能进行合并!';
        Raise Err_Custom;
      End If;
    End Loop;
  End If;

  --新门诊病人合并规则与ZLHIS不一致,禁止新门诊病人在ZLHIS合并操作 
  Select Count(1) Into v_Count From 病人挂号记录 A Where a.病人id In (A病人id_In, B病人id_In) And a.附加标志 = 3;
  If v_Count <> 0 Then
    v_Error := '本次合并的病人中含新门诊病人,不能进行合并！';
    Raise Err_Custom;
  End If;

  --程序中已检查：
  --1.选择了同一个病人
  --2.两个住院病人先入院的却在院(包括两个都在院)。
  --3.两个住院病人的住院期间存在交叉的情况
  --4.医保病人存在未结费用

  --先锁定病人不允许进行其他业务
  Zl_病人信息_锁定(A病人id_In, 1);
  Zl_病人信息_锁定(B病人id_In, 1);

  Open c_Infoa;
  Fetch c_Infoa
    Into r_Infoa;
  If c_Infoa%RowCount = 0 Then
    Close c_Infoa;
    v_Error := '没有发现被合并的病人信息！';
    Raise Err_Custom;
  End If;

  Open c_Infob;
  Fetch c_Infob
    Into r_Infob;
  If c_Infob%RowCount = 0 Then
    Close c_Infob;
    v_Error := '没有发现要保留的病人信息！';
    Raise Err_Custom;
  End If;

  --读取其它相关病人表到数组
  For r_Patitable In c_Patitable Loop
    If r_Patitable.主页id = 0 Then
      Arronbase(Arronbase.Count + 1) := r_Patitable.Table_Name;
    Else
      Arronpage(Arronpage.Count + 1) := r_Patitable.Table_Name;
    End If;
  End Loop;

  --以先住院或先登记的病人ID作为实际上要保留的病人ID
  If Nvl(强制保留_In, 0) = 1 Then
    v_保留id := B病人id_In;
  Else
    Select 病人id
    Into v_保留id
    From (Select /*+ CHOOSE */
            a.病人id
           From 病人信息 A, 病案主页 B
           Where a.病人id = b.病人id(+) And a.病人id In (A病人id_In, B病人id_In)
           Order By Nvl(b.入院日期, To_Date('3000-01-01', 'YYYY-MM-DD')), Nvl(b.出院日期, To_Date('3000-01-01', 'YYYY-MM-DD')),
                    a.登记时间, a.病人id --住院病人优先
           )
    Where Rownum = 1;
  End If;

  --先确定病案号的模式
  Select Zl_To_Number(Nvl(zl_GetSysParameter(39), '0')) Into n_独立病案 From Dual;
  --住院号模式
  Select Zl_To_Number(Nvl(zl_GetSysParameter(145), '0')) Into n_每次新住院号 From Dual;

  --另外一个就是实际最后要删除的病人ID
  If v_保留id = A病人id_In Then
    v_合并id := B病人id_In;
    --问题27445 保留指定病人的门诊号、住院号、医保号
    v_门诊号 := Nvl(r_Infob.门诊号, r_Infoa.门诊号);
    v_住院号 := Nvl(r_Infob.住院号, r_Infoa.住院号);
  Else
    v_合并id := A病人id_In;
    v_门诊号 := Nvl(r_Infob.门诊号, r_Infoa.门诊号);
    v_住院号 := Nvl(r_Infob.住院号, r_Infoa.住院号);
  End If;

  ---记录合并操作,在后面会根据r_PatiTable把合并病人的合并记录更新为保留病人的
  v_原信息 := v_合并id || ',' || r_Infoa.门诊号 || ',' || r_Infoa.住院号 || ',' || r_Infoa.就诊卡号 || ',' || r_Infoa.姓名 || ',' ||
           r_Infoa.性别 || ',' || r_Infoa.年龄 || ',' || To_Char(r_Infoa.出生日期, 'yyyy-mm-dd') || ',' || r_Infoa.身份证号 || ',' ||
           r_Infoa.婚姻状况 || ',' || r_Infoa.职业 || ',' || r_Infoa.家庭地址;
  Insert Into 病人合并记录
    (病人id, 原信息, 合并原因, 操作员姓名, 合并时间)
  Values
    (v_保留id, v_原信息, 合并原因_In, 操作员姓名_In, Sysdate);

  --开始合并
  --84398修改将住院次数计算放在外面，因需要考虑门诊和住院病人合并
  --10.34开始,住院次数不包含留关病人,合并后的住院次数=保留病人住院次数+合并病人正常入院的次数
  Select Nvl(住院次数, 0) Into n_Cur住院次数 From 病人信息 Where 病人id = v_保留id;
  Select Count(*) Into n_Cnt住院次数 From 病案主页 Where 病人id = v_合并id And 主页id <> 0 And 病人性质 = 0;
  n_Max住院次数 := n_Cur住院次数 + n_Cnt住院次数;
  --处理病案主页部份(涉及病人ID,主页ID字段的表)
  If (r_Infoa.主页id Is Not Null And r_Infob.主页id Is Not Null) Or (强制保留_In = 1 And r_Infoa.主页id Is Not Null) Then
    If r_Infoa.主页id = 0 And r_Infob.主页id = 0 Then
      Close c_Infoa;
      Close c_Infob;
      v_Error := '两个预约病人不能进行病人合并操作！';
      Raise Err_Custom;
    Elsif r_Infoa.主页id = 0 Then
      If r_Infob.入院日期 Is Not Null And r_Infob.出院日期 Is Null Then
        Close c_Infoa;
        Close c_Infob;
        v_Error := '预约病人和在院病人不能进行病人合并操作！';
        Raise Err_Custom;
      End If;
    Elsif r_Infob.主页id = 0 Then
      If r_Infoa.入院日期 Is Not Null And r_Infoa.出院日期 Is Null Then
        Close c_Infoa;
        Close c_Infob;
        v_Error := '预约病人和在院病人不能进行病人合并操作！';
        Raise Err_Custom;
      End If;
    End If;
    --求两个病人总共的住院就诊次数
    Select Count(*) Into v_Count From 病案主页 Where 病人id In (A病人id_In, B病人id_In) And 主页id <> 0;
    --因为10.19开始，入院时允许修改主页id，所以最大主页ID可能大于总的住院就诊次数
    Select Max(主页id) Into n_Max主页id From 病案主页 Where 病人id = v_保留id And 主页id <> 0;
    Select Count(*) Into n_Cnt主页id From 病案主页 Where 病人id = v_合并id And 主页id <> 0;
    If n_Max主页id + n_Cnt主页id > v_Count Then
      v_Count := n_Max主页id + n_Cnt主页id;
    End If;
    --求实际要更新的主页截至值,以前用v_Count >= n_Max主页id判断存在一个问题（对于两个病人多次交叉入院，可能导致A,B病人部分就诊次数没有更新）
    Select Nvl(Max(主页id), 0)
    Into n_Loop主页id
    From 病案主页 A, (Select Min(入院日期) 入院日期 From 病案主页 Where 病人id = v_合并id) B
    Where a.病人id = v_保留id And a.入院日期 < b.入院日期;
  
    For r_Merge In c_Mergepati Loop
      If Not (r_Merge.病人id = v_保留id And r_Merge.主页id = v_Count) And v_Count <> 0 Then
        --该病案主页要删除时,不能是已编目了的。
        If r_Merge.编目日期 Is Not Null Then
          Close c_Infoa;
          Close c_Infob;
          If r_Merge.当前住院号 Is Null Then
            v_Error := '病人' || r_Merge.姓名 || '(病人ID=' || r_Merge.病人id || ')存在已编目的病案,不允许合并该病人。';
          Else
            v_Error := '病人' || r_Merge.姓名 || '(病人ID=' || r_Merge.病人id || ',住院号=' || r_Merge.当前住院号 ||
                       ')存在已编目的病案,不允许合并该病人。';
          End If;
          Raise Err_Custom;
        End If;
        If v_Count >= Nvl(n_Loop主页id, 0) Then
          If r_Merge.主页id = 0 Then
            n_Cur主页id := 0;
            Update 病案主页
            Set 病人性质 = r_Merge.病人性质, 医疗付款方式 = r_Merge.医疗付款方式, 费别 = r_Merge.费别, 再入院 = r_Merge.再入院,
                入院病区id = r_Merge.入院病区id, 入院科室id = r_Merge.入院科室id, 入院日期 = r_Merge.入院日期, 入院病况 = r_Merge.入院病况,
                入院方式 = r_Merge.入院方式, 二级院转入 = r_Merge.二级院转入, 住院目的 = r_Merge.住院目的, 入院病床 = r_Merge.入院病床,
                是否陪伴 = r_Merge.是否陪伴, 当前病况 = r_Merge.当前病况, 当前病区id = r_Merge.当前病区id, 护理等级id = r_Merge.护理等级id,
                出院科室id = r_Merge.出院科室id, 出院病床 = r_Merge.出院病床, 出院日期 = r_Merge.出院日期, 住院天数 = r_Merge.住院天数,
                出院方式 = r_Merge.出院方式, 是否确诊 = r_Merge.是否确诊, 确诊日期 = r_Merge.确诊日期, 新发肿瘤 = r_Merge.新发肿瘤, 血型 = r_Merge.血型,
                抢救次数 = r_Merge.抢救次数, 成功次数 = r_Merge.成功次数, 随诊标志 = r_Merge.随诊标志, 随诊期限 = r_Merge.随诊期限, 尸检标志 = r_Merge.尸检标志,
                门诊医师 = r_Merge.门诊医师, 责任护士 = r_Merge.责任护士, 住院医师 = r_Merge.住院医师, 编目员编号 = r_Merge.编目员编号,
                编目员姓名 = r_Merge.编目员姓名, 编目日期 = r_Merge.编目日期, 状态 = r_Merge.状态, 费用和 = r_Merge.费用和, 姓名 = r_Merge.姓名,
                性别 = r_Merge.性别, 年龄 = r_Merge.年龄, 婚姻状况 = r_Merge.婚姻状况, 职业 = r_Merge.职业, 国籍 = r_Merge.国籍, 学历 = r_Merge.学历,
                单位电话 = r_Merge.单位电话, 单位邮编 = r_Merge.单位邮编, 单位地址 = r_Merge.单位地址, 区域 = r_Merge.区域, 家庭地址 = r_Merge.家庭地址,
                家庭电话 = r_Merge.家庭电话, 家庭地址邮编 = r_Merge.家庭地址邮编, 户口地址 = r_Merge.户口地址, 户口地址邮编 = r_Merge.户口地址邮编,
                联系人姓名 = r_Merge.联系人姓名, 联系人关系 = r_Merge.联系人关系, 联系人地址 = r_Merge.联系人地址, 联系人电话 = r_Merge.联系人电话,
                中医治疗类别 = r_Merge.中医治疗类别, 登记人 = r_Merge.登记人, 登记时间 = r_Merge.登记时间, 险类 = r_Merge.险类, 审核标志 = r_Merge.审核标志,
                是否上传 = r_Merge.是否上传, 备注 = r_Merge.备注, 数据转出 = r_Merge.数据转出, 病案号 = r_Merge.病案号,
                住院号 = Decode(n_每次新住院号, 1, r_Merge.住院号, v_住院号), 留观号 = r_Merge.留观号, 病人类型 = r_Merge.病人类型,
                封存时间 = r_Merge.封存时间, 路径状态 = r_Merge.路径状态, 单病种 = r_Merge.单病种, 婴儿科室id = r_Merge.婴儿科室id,
                婴儿病区id = r_Merge.婴儿病区id, 母婴转科标志 = r_Merge.母婴转科标志, 医嘱重整时间 = r_Merge.医嘱重整时间
            Where 病人id = v_保留id And 主页id = n_Cur主页id;
            If Sql%RowCount = 0 Then
              Insert Into 病案主页
                (病人id, 主页id, 病人性质, 医疗付款方式, 费别, 再入院, 入院病区id, 入院科室id, 入院日期, 入院病况, 入院方式, 二级院转入, 住院目的, 入院病床, 是否陪伴, 当前病况,
                 当前病区id, 护理等级id, 出院科室id, 出院病床, 出院日期, 住院天数, 出院方式, 是否确诊, 确诊日期, 新发肿瘤, 血型, 抢救次数, 成功次数, 随诊标志, 随诊期限, 尸检标志,
                 门诊医师, 责任护士, 住院医师, 编目员编号, 编目员姓名, 编目日期, 状态, 费用和, 姓名, 性别, 年龄, 婚姻状况, 职业, 国籍, 学历, 单位电话, 单位邮编, 单位地址, 区域, 家庭地址,
                 家庭电话, 家庭地址邮编, 户口地址, 户口地址邮编, 联系人姓名, 联系人关系, 联系人地址, 联系人电话, 中医治疗类别, 登记人, 登记时间, 险类, 审核标志, 是否上传, 备注, 数据转出,
                 病案号, 住院号, 留观号, 病人类型, 封存时间, 路径状态, 单病种, 婴儿科室id, 婴儿病区id, 母婴转科标志, 医嘱重整时间)
              Values
                (v_保留id, n_Cur主页id, r_Merge.病人性质, r_Merge.医疗付款方式, r_Merge.费别, r_Merge.再入院, r_Merge.入院病区id,
                 r_Merge.入院科室id, r_Merge.入院日期, r_Merge.入院病况, r_Merge.入院方式, r_Merge.二级院转入, r_Merge.住院目的, r_Merge.入院病床,
                 r_Merge.是否陪伴, r_Merge.当前病况, r_Merge.当前病区id, r_Merge.护理等级id, r_Merge.出院科室id, r_Merge.出院病床, r_Merge.出院日期,
                 r_Merge.住院天数, r_Merge.出院方式, r_Merge.是否确诊, r_Merge.确诊日期, r_Merge.新发肿瘤, r_Merge.血型, r_Merge.抢救次数,
                 r_Merge.成功次数, r_Merge.随诊标志, r_Merge.随诊期限, r_Merge.尸检标志, r_Merge.门诊医师, r_Merge.责任护士, r_Merge.住院医师,
                 r_Merge.编目员编号, r_Merge.编目员姓名, r_Merge.编目日期, r_Merge.状态, r_Merge.费用和, r_Merge.姓名, r_Merge.性别, r_Merge.年龄,
                 r_Merge.婚姻状况, r_Merge.职业, r_Merge.国籍, r_Merge.学历, r_Merge.单位电话, r_Merge.单位邮编, r_Merge.单位地址, r_Merge.区域,
                 r_Merge.家庭地址, r_Merge.家庭电话, r_Merge.家庭地址邮编, r_Merge.户口地址, r_Merge.户口地址邮编, r_Merge.联系人姓名, r_Merge.联系人关系,
                 r_Merge.联系人地址, r_Merge.联系人电话, r_Merge.中医治疗类别, r_Merge.登记人, r_Merge.登记时间, r_Merge.险类, r_Merge.审核标志,
                 r_Merge.是否上传, r_Merge.备注, r_Merge.数据转出, r_Merge.病案号, Decode(n_每次新住院号, 1, r_Merge.住院号, v_住院号),
                 r_Merge.留观号, r_Merge.病人类型, r_Merge.封存时间, r_Merge.路径状态, r_Merge.单病种, r_Merge.婴儿科室id, r_Merge.婴儿病区id,
                 r_Merge.母婴转科标志, r_Merge.医嘱重整时间);
            End If;
          Else
            n_Cur主页id := v_Count;
            Insert Into 病案主页
              (病人id, 主页id, 病人性质, 医疗付款方式, 费别, 再入院, 入院病区id, 入院科室id, 入院日期, 入院病况, 入院方式, 二级院转入, 住院目的, 入院病床, 是否陪伴, 当前病况,
               当前病区id, 护理等级id, 出院科室id, 出院病床, 出院日期, 住院天数, 出院方式, 是否确诊, 确诊日期, 新发肿瘤, 血型, 抢救次数, 成功次数, 随诊标志, 随诊期限, 尸检标志, 门诊医师,
               责任护士, 住院医师, 编目员编号, 编目员姓名, 编目日期, 状态, 费用和, 姓名, 性别, 年龄, 婚姻状况, 职业, 国籍, 学历, 单位电话, 单位邮编, 单位地址, 区域, 家庭地址, 家庭电话,
               家庭地址邮编, 户口地址, 户口地址邮编, 联系人姓名, 联系人关系, 联系人地址, 联系人电话, 中医治疗类别, 登记人, 登记时间, 险类, 审核标志, 是否上传, 备注, 数据转出, 病案号, 住院号,
               留观号, 病人类型, 封存时间, 路径状态, 单病种, 婴儿科室id, 婴儿病区id, 母婴转科标志, 医嘱重整时间)
            Values
              (v_保留id, n_Cur主页id, r_Merge.病人性质, r_Merge.医疗付款方式, r_Merge.费别, r_Merge.再入院, r_Merge.入院病区id, r_Merge.入院科室id,
               r_Merge.入院日期, r_Merge.入院病况, r_Merge.入院方式, r_Merge.二级院转入, r_Merge.住院目的, r_Merge.入院病床, r_Merge.是否陪伴,
               r_Merge.当前病况, r_Merge.当前病区id, r_Merge.护理等级id, r_Merge.出院科室id, r_Merge.出院病床, r_Merge.出院日期, r_Merge.住院天数,
               r_Merge.出院方式, r_Merge.是否确诊, r_Merge.确诊日期, r_Merge.新发肿瘤, r_Merge.血型, r_Merge.抢救次数, r_Merge.成功次数,
               r_Merge.随诊标志, r_Merge.随诊期限, r_Merge.尸检标志, r_Merge.门诊医师, r_Merge.责任护士, r_Merge.住院医师, r_Merge.编目员编号,
               r_Merge.编目员姓名, r_Merge.编目日期, r_Merge.状态, r_Merge.费用和, r_Merge.姓名, r_Merge.性别, r_Merge.年龄, r_Merge.婚姻状况,
               r_Merge.职业, r_Merge.国籍, r_Merge.学历, r_Merge.单位电话, r_Merge.单位邮编, r_Merge.单位地址, r_Merge.区域, r_Merge.家庭地址,
               r_Merge.家庭电话, r_Merge.家庭地址邮编, r_Merge.户口地址, r_Merge.户口地址邮编, r_Merge.联系人姓名, r_Merge.联系人关系, r_Merge.联系人地址,
               r_Merge.联系人电话, r_Merge.中医治疗类别, r_Merge.登记人, r_Merge.登记时间, r_Merge.险类, r_Merge.审核标志, r_Merge.是否上传,
               r_Merge.备注, r_Merge.数据转出, r_Merge.病案号, Decode(n_每次新住院号, 1, r_Merge.住院号, v_住院号), r_Merge.留观号, r_Merge.病人类型,
               r_Merge.封存时间, r_Merge.路径状态, r_Merge.单病种, r_Merge.婴儿科室id, r_Merge.婴儿病区id, r_Merge.母婴转科标志, r_Merge.医嘱重整时间);
          End If;
        Else
          Exit;
        End If;

        ---- v_保留id,n_Cur主页id:r_Merge.病人id, r_Merge.主页id
        v_Chgs := v_Chgs || ';' || v_保留id || ',' || n_Cur主页id || ':' || r_Merge.病人id || ',' || r_Merge.主页id;
      
        --更新病人相关表的病人指向
        ---------------------------------------------------------------
        --病人变动记录
        Update 病人变动记录
        Set 病人id = v_保留id, 主页id = n_Cur主页id
        Where 病人id = r_Merge.病人id And 主页id = r_Merge.主页id;
        
		--病人自动计算
        Update 病人自动计算
        Set 病人id = v_保留id, 主页id = n_Cur主页id
        Where 病人id = r_Merge.病人id And 主页id = r_Merge.主页id;

        --病案主页从表
        Update 病案主页从表
        Set 病人id = v_保留id, 主页id = n_Cur主页id
        Where 病人id = r_Merge.病人id And 主页id = r_Merge.主页id;
      
        --住院费用记录
        Update 住院费用记录
        Set 病人id = v_保留id, 主页id = n_Cur主页id,
            标识号 = Nvl(Decode(门诊标志, 1, v_门诊号, Decode(n_每次新住院号, 1, r_Merge.住院号, v_住院号)), 标识号)
        Where 病人id = r_Merge.病人id And 主页id = r_Merge.主页id;
        Update H住院费用记录
        Set 病人id = v_保留id, 主页id = n_Cur主页id,
            标识号 = Nvl(Decode(门诊标志, 1, v_门诊号, Decode(n_每次新住院号, 1, r_Merge.住院号, v_住院号)), 标识号)
        Where 病人id = r_Merge.病人id And 主页id = r_Merge.主页id;
        --门诊费用记录
        --Update 门诊费用记录
        --Set 病人id = v_保留id,
        --    标识号 = Nvl(Decode(门诊标志, 1, v_门诊号, Decode(n_每次新住院号, 1, r_Merge.住院号, v_住院号)), 标识号)
        --Where 病人id = r_Merge.病人id;
        --Update H门诊费用记录
        --Set 病人id = v_保留id,
        --    标识号 = Nvl(Decode(门诊标志, 1, v_门诊号, Decode(n_每次新住院号, 1, r_Merge.住院号, v_住院号)), 标识号)
        --Where 病人id = r_Merge.病人id;
      
        --病人预交记录
        Update 病人预交记录
        Set 病人id = v_保留id, 主页id = n_Cur主页id
        Where 病人id = r_Merge.病人id And 主页id = r_Merge.主页id;
        Update H病人预交记录
        Set 病人id = v_保留id, 主页id = n_Cur主页id
        Where 病人id = r_Merge.病人id And 主页id = r_Merge.主页id;
      
        --病人未结费用
        Update 病人未结费用
        Set 病人id = v_保留id, 主页id = n_Cur主页id
        Where 病人id = r_Merge.病人id And 主页id = r_Merge.主页id;
      
        --未发药品记录
        Update 未发药品记录
        Set 病人id = v_保留id, 主页id = n_Cur主页id
        Where 病人id = r_Merge.病人id And 主页id = r_Merge.主页id;
      
        --诊断情况
        Update 诊断情况
        Set 病人id = v_保留id, 主页id = n_Cur主页id
        Where 病人id = r_Merge.病人id And 主页id = r_Merge.主页id;
      
        --保险结算记录(病人ID和非住院病人一起在后面处理)
        Update 保险结算记录 Set 主页id = n_Cur主页id Where 病人id = r_Merge.病人id And 主页id = r_Merge.主页id;
      
        --保险模拟结算
        Update 保险模拟结算
        Set 病人id = v_保留id, 主页id = n_Cur主页id
        Where 病人id = r_Merge.病人id And 主页id = r_Merge.主页id;
      
        --病人医嘱记录(ZLHIS+)
        Update 病人医嘱记录
        Set 病人id = v_保留id, 主页id = n_Cur主页id
        Where 病人id = r_Merge.病人id And 主页id = r_Merge.主页id;
        Update H病人医嘱记录
        Set 病人id = v_保留id, 主页id = n_Cur主页id
        Where 病人id = r_Merge.病人id And 主页id = r_Merge.主页id;
      
        --病人过敏记录(ZLHIS+)
        Update 病人过敏记录
        Set 病人id = v_保留id, 主页id = n_Cur主页id
        Where 病人id = r_Merge.病人id And 主页id = r_Merge.主页id;
        Update H病人过敏记录
        Set 病人id = v_保留id, 主页id = n_Cur主页id
        Where 病人id = r_Merge.病人id And 主页id = r_Merge.主页id;
      
        --病人诊断记录(ZLHIS+)
        Update 病人诊断记录
        Set 病人id = v_保留id, 主页id = n_Cur主页id
        Where 病人id = r_Merge.病人id And 主页id = r_Merge.主页id;
        Update H病人诊断记录
        Set 病人id = v_保留id, 主页id = n_Cur主页id
        Where 病人id = r_Merge.病人id And 主页id = r_Merge.主页id;
      
        --病人手麻记录(ZLHIS+)
        Update 病人手麻记录
        Set 病人id = v_保留id, 主页id = n_Cur主页id
        Where 病人id = r_Merge.病人id And 主页id = r_Merge.主页id;
        Update H病人手麻记录
        Set 病人id = v_保留id, 主页id = n_Cur主页id
        Where 病人id = r_Merge.病人id And 主页id = r_Merge.主页id;
      
        --病人担保记录(zlhis+)
        Update 病人担保记录
        Set 病人id = v_保留id, 主页id = n_Cur主页id
        Where 病人id = r_Merge.病人id And 主页id = r_Merge.主页id;
      
        --病案系统的表
        Begin
          v_Sql := 'Update 病人费用 Set 病人ID=:1,主页ID=:2 Where 病人ID=:3 And 主页ID=:4';
          Execute Immediate v_Sql
            Using v_保留id, n_Cur主页id, r_Merge.病人id, r_Merge.主页id;
        Exception
          When Others Then
            Null;
        End;
      
        Begin
          v_Sql := 'Update 随诊记录 Set 病人ID=:1,主页ID=:2 Where 病人ID=:3 And 主页ID=:4';
          Execute Immediate v_Sql
            Using v_保留id, n_Cur主页id, r_Merge.病人id, r_Merge.主页id;
        Exception
          When Others Then
            Null;
        End;
      
        Begin
          v_Sql := 'Update 诊断符合情况 Set 病人ID=:1,主页ID=:2 Where 病人ID=:3 And 主页ID=:4';
          Execute Immediate v_Sql
            Using v_保留id, n_Cur主页id, r_Merge.病人id, r_Merge.主页id;
        Exception
          When Others Then
            Null;
        End;
      
        Begin
          v_Sql := 'Update 病案评分结果 Set 病人ID=:1,主页ID=:2 Where 病人ID=:3 And 主页ID=:4';
          Execute Immediate v_Sql
            Using v_保留id, n_Cur主页id, r_Merge.病人id, r_Merge.主页id;
        Exception
          When Others Then
            Null;
        End;
      
        Begin
          v_Sql := 'Insert Into 病人分娩信息(病人ID,主页ID,胎儿次序,分娩方式,出生胎位,分娩情况,出生缺陷,婴儿性别,婴儿体重,Apgar评分) ' ||
                   'Select :1,:2,胎儿次序,分娩方式,出生胎位,分娩情况,出生缺陷,婴儿性别,婴儿体重,Apgar评分 From 病人分娩信息 Where 病人ID=:3 And 主页ID=:4';
          Execute Immediate v_Sql
            Using v_保留id, n_Cur主页id, r_Merge.病人id, r_Merge.主页id;
        Exception
          When Others Then
            Null;
        End;
      
        Begin
          v_Sql := 'Delete From 病人分娩信息 Where 病人ID=:1 And 主页ID=:2';
          Execute Immediate v_Sql
            Using r_Merge.病人id, r_Merge.主页id;
        Exception
          When Others Then
            Null;
        End;
      
        Begin
          v_Sql := 'Update 借阅记录 Set 病人ID=:1,主页ID=:2 Where 病人ID=:3 And 主页ID=:4';
          Execute Immediate v_Sql
            Using v_保留id, n_Cur主页id, r_Merge.病人id, r_Merge.主页id;
        Exception
          When Others Then
            Null;
        End;
      
        --其它病案主页相关表
        For v_Loop In 1 .. Arronpage.Count Loop
          v_Sql := 'Update ' || Arronpage(v_Loop) || ' Set 病人ID=:1,主页ID=:2 Where 病人ID=:3 And 主页ID=:4';
          Execute Immediate v_Sql
            Using v_保留id, n_Cur主页id, r_Merge.病人id, r_Merge.主页id;
        End Loop;
      
        --删除已调整后的病案主页
        Delete From 病案主页 Where 病人id = r_Merge.病人id And 主页id = r_Merge.主页id;
      End If;
      If r_Merge.主页id <> 0 Then
        v_Count := v_Count - 1;
      End If;
    End Loop;
  End If;

  --不涉及主页ID部份的更改(无主页ID或主页ID可能为空)
  ---------------------------------------------------------------
  --住院费用记录
  Update 住院费用记录
  Set 病人id = v_保留id, 标识号 = Nvl(Decode(门诊标志, 2, v_住院号, v_门诊号), 标识号)
  Where 病人id = v_合并id;
  Update H住院费用记录
  Set 病人id = v_保留id, 标识号 = Nvl(Decode(门诊标志, 2, v_住院号, v_门诊号), 标识号)
  Where 病人id = v_合并id;
  --门诊费用记录
  Update 门诊费用记录
  Set 病人id = v_保留id, 标识号 = Nvl(Decode(门诊标志, 2, v_住院号, v_门诊号), 标识号)
  Where 病人id = v_合并id;
  Update H门诊费用记录
  Set 病人id = v_保留id, 标识号 = Nvl(Decode(门诊标志, 2, v_住院号, v_门诊号), 标识号)
  Where 病人id = v_合并id;

  --病人预交记录
  Update 病人预交记录 Set 病人id = v_保留id Where 病人id = v_合并id And 主页id Is Null;
  Update H病人预交记录 Set 病人id = v_保留id Where 病人id = v_合并id And 主页id Is Null;

  --未发药品记录
  Update 未发药品记录 Set 病人id = v_保留id Where 病人id = v_合并id And 主页id Is Null;

  --诊断情况
  Update 诊断情况 Set 病人id = v_保留id Where 病人id = v_合并id And 主页id Is Null;

  --病人医嘱记录(ZLHIS+)
  Update 病人医嘱记录 Set 病人id = v_保留id Where 病人id = v_合并id And 主页id Is Null;
  Update H病人医嘱记录 Set 病人id = v_保留id Where 病人id = v_合并id And 主页id Is Null;

  --病人过敏记录(ZLHIS+):主页ID可能是挂号ID
  Update 病人过敏记录 Set 病人id = v_保留id Where 病人id = v_合并id;
  Update H病人过敏记录 Set 病人id = v_保留id Where 病人id = v_合并id;

  --病人诊断记录(ZLHIS+):主页ID可能是挂号ID
  Update 病人诊断记录 Set 病人id = v_保留id Where 病人id = v_合并id;
  Update H病人诊断记录 Set 病人id = v_保留id Where 病人id = v_合并id;

  --病人手麻记录(ZLHIS+)
  Update 病人手麻记录 Set 病人id = v_保留id Where 病人id = v_合并id And 主页id Is Null;
  Update H病人手麻记录 Set 病人id = v_保留id Where 病人id = v_合并id And 主页id Is Null;

  --病人挂号记录(ZLHIS+)
  Update 病人挂号记录 Set 病人id = v_保留id, 门诊号 = Nvl(v_门诊号, 门诊号) Where 病人id = v_合并id;

  --病人结帐记录
  Update 病人结帐记录 Set 病人id = v_保留id Where 病人id = v_合并id;
  Update H病人结帐记录 Set 病人id = v_保留id Where 病人id = v_合并id;

  --床位状况记录
  Update 床位状况记录 Set 病人id = v_保留id Where 病人id = v_合并id;

  --病人担保记录
  Update 病人担保记录 Set 病人id = v_保留id Where 病人id = v_合并id;
  --特殊病人
  Select Count(*) Into v_Count From 特殊病人 Where 病人id = v_保留id;
  If v_Count = 0 Then
    Update 特殊病人 Set 病人id = v_保留id Where 病人id = v_合并id;
  Else
    Delete From 特殊病人 Where 病人id = v_合并id;
  End If;

  --病人未结费用
  For r_Owe In c_Owe(v_合并id) Loop
    Update 病人未结费用
    Set 金额 = Nvl(金额, 0) + Nvl(r_Owe.金额, 0)
    Where 主页id Is Null And 病人id = v_保留id And Nvl(病人病区id, 0) = Nvl(r_Owe.病人病区id, 0) And
          Nvl(病人科室id, 0) = Nvl(r_Owe.病人科室id, 0) And Nvl(开单部门id, 0) = Nvl(r_Owe.开单部门id, 0) And
          Nvl(执行部门id, 0) = Nvl(r_Owe.执行部门id, 0) And Nvl(收入项目id, 0) = Nvl(r_Owe.收入项目id, 0) And
          Nvl(来源途径, 0) = Nvl(r_Owe.来源途径, 0);
    If Sql%RowCount = 0 Then
      Insert Into 病人未结费用
        (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
      Values
        (v_保留id, Null, r_Owe.病人病区id, r_Owe.病人科室id, r_Owe.开单部门id, r_Owe.执行部门id, r_Owe.收入项目id, r_Owe.来源途径, r_Owe.金额);
    End If;
  End Loop;
  Delete From 病人未结费用 Where 病人id = v_合并id;
  Delete From 病人未结费用 Where 病人id = v_保留id And Nvl(金额, 0) = 0;

  --病人余额
  For r_Spare In c_Spare(v_合并id) Loop
    Update 病人余额
    Set 预交余额 = Nvl(预交余额, 0) + Nvl(r_Spare.预交余额, 0), 费用余额 = Nvl(费用余额, 0) + Nvl(r_Spare.费用余额, 0)
    Where Nvl(性质, 0) = Nvl(r_Spare.性质, 0) And 病人id = v_保留id And 类型 = Nvl(r_Spare.类型, 2);
    If Sql%RowCount = 0 Then
      Insert Into 病人余额
        (病人id, 性质, 类型, 预交余额, 费用余额)
      Values
        (v_保留id, r_Spare.性质, Nvl(r_Spare.类型, 2), r_Spare.预交余额, r_Spare.费用余额);
    End If;
  End Loop;
  Delete From 病人余额 Where 病人id = v_合并id;
  Delete From 病人余额 Where 病人id = v_保留id And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0 And 性质 = 1;

  --病人过敏药物
  Insert Into 病人过敏药物
    (病人id, 过敏药物id, 过敏药物)
    Select v_保留id, 过敏药物id, 过敏药物
    From 病人过敏药物
    Where 病人id = v_合并id And 过敏药物id Not In (Select 过敏药物id From 病人过敏药物 Where 病人id = v_保留id);
  Delete From 病人过敏药物 Where 病人id = v_合并id;

  --病人社区信息
  Insert Into 病人社区信息
    (病人id, 社区, 社区号, 标志, 就诊类型, 就诊时间)
    Select v_保留id, 社区, 社区号, 标志, 就诊类型, 就诊时间
    From 病人社区信息
    Where 病人id = v_合并id And 社区 Not In (Select 社区 From 病人社区信息 Where 病人id = v_保留id);
  Delete From 病人社区信息 Where 病人id = v_合并id;

  --病人免疫记录
  Insert Into 病人免疫记录
    (病人id, 接种时间, 接种名称)
    Select v_保留id, a.接种时间, a.接种名称
    From 病人免疫记录 A
    Where a.病人id = v_合并id And Not Exists (Select 1 From 病人免疫记录 Where 病人id = v_保留id And 接种时间 = a.接种时间);
  Delete From 病人免疫记录 Where 病人id = v_合并id;

  --病人信息从表
  Insert Into 病人信息从表
    (病人id, 信息名, 信息值, 就诊id)
    Select v_保留id, a.信息名, a.信息值, a.就诊id
    From 病人信息从表 A
    Where a.病人id = v_合并id And Not Exists (Select 1
           From 病人信息从表
           Where 病人id = v_保留id And 信息名 = a.信息名 And Nvl(就诊id, 0) = Nvl(a.就诊id, 0));
  Delete From 病人信息从表 Where 病人id = v_合并id;

  --病人医疗卡属性
  Insert Into 病人医疗卡属性
    (病人id, 卡类别id, 卡号, 信息名, 信息值)
    Select v_保留id, a.卡类别id, a.卡号, a.信息名, a.信息值
    From 病人医疗卡属性 A
    Where a.病人id = v_合并id And Not Exists (Select 1
           From 病人医疗卡属性
           Where 病人id = v_保留id And 卡类别id = a.卡类别id And 卡号 = a.卡号 And 信息名 = a.信息名);
  Delete From 病人医疗卡属性 Where 病人id = v_合并id;

  --门诊病案记录
  Select Count(*) Into v_Count From 门诊病案记录 Where 病人id = v_保留id;
  If v_Count = 0 Then
    Select Count(*) Into v_Count From 门诊病案记录 Where 病人id = v_合并id;
    If v_Count > 0 Then
      Update 门诊病案记录 Set 病人id = v_保留id Where 病人id = v_合并id;
    End If;
  Else
    Delete From 门诊病案记录 Where 病人id = v_合并id;
  End If;

  --住院病案记录
  Select Count(*) Into v_Count From 住院病案记录 Where 病人id = v_保留id;

  If v_Count = 0 Then
    Select Count(*) Into v_Count From 住院病案记录 Where 病人id = v_合并id;
    If v_Count > 0 Then
      Update 住院病案记录 Set 病人id = v_保留id Where 病人id = v_合并id;
    End If;
  Else
    Begin
      v_Sql := 'Delete From 借阅记录 Where 病人ID=:1';
      Execute Immediate v_Sql
        Using v_合并id;
    Exception
      When Others Then
        Null;
    End;
  
    Delete From 住院病案记录 Where 病人id = v_合并id;
  End If;

  --医保病人相关处理
  --即使合病或保留的病人当前不是医保帐户,只要曾是医保帐户,险类不同也不能合并
  Select Count(Distinct 险类) Into v_Count From 医保病人关联表 Where 病人id In (v_合并id, v_保留id);
  If v_Count = 2 Then
    Close c_Infoa;
    Close c_Infob;
    v_Error := '两个病人分别属于不同的保险类别，不允许合并。';
    Raise Err_Custom;
  End If;

  Select Count(*) Into v_Count From 医保病人关联表 Where 病人id = v_合并id And 标志 = 0;
  --a.合并的病人以前是医保帐户,现在不是
  If v_Count > 0 Then
    Select Count(*) Into v_Count From 医保病人关联表 Where 病人id = v_保留id;
    --a.1保留的病人现在是医保帐户
    --a.2.1保留的病人现在不是医保帐户,以前是,与a.1相同处理
    If v_Count > 0 Then
      Delete From 帐户年度信息 Where 病人id = v_合并id;
    
      Select Count(Distinct 医保号) Into v_Count From 医保病人关联表 Where 病人id In (v_合并id, v_保留id);
      If v_Count <> 2 Then
        --两个病人医保号相同时,不用处理医保病人档案
        For r_Insure In c_Insure(v_合并id) Loop
          --被合并的病人可能关联了多个医保病人,改为关联到保留的病人上
          --问题27445 保留指定病人的门诊号、住院号、医保号
          If v_合并id = B病人id_In Then
            Update 医保病人关联表
            Set 医保号 =
                 (Select 医保号 From 医保病人关联表 Where 病人id = v_合并id), 标志 = 0
            Where 险类 = r_Insure.险类 And 中心 = r_Insure.中心 And 医保号 = r_Insure.医保号 And 病人id <> v_合并id;
          Else
            Update 医保病人关联表
            Set 医保号 =
                 (Select 医保号 From 医保病人关联表 Where 病人id = v_保留id), 标志 = 0
            Where 险类 = r_Insure.险类 And 中心 = r_Insure.中心 And 医保号 = r_Insure.医保号 And 病人id <> v_合并id;
          End If;
          --合并的病人现在不是医保,即使是用户指定要保留该病人,也不保留它的帐户信息
          Delete From 医保病人档案 Where 险类 = r_Insure.险类 And 医保号 = r_Insure.医保号;
        End Loop;
      End If;
      Delete From 医保病人关联表 Where 病人id = v_合并id;
    Else
      --a.2.2保留的病人现在和以前都不是医保帐户
      Update 帐户年度信息 Set 病人id = v_保留id Where 病人id = v_合并id;
      Update 医保病人关联表 Set 病人id = v_保留id Where 病人id = v_合并id;
      --医保病人档案表不用处理,因为通过医保号关联<医保病人关联表>
    End If;
  Else
    Select Count(*) Into v_Count From 医保病人关联表 Where 病人id = v_合并id And 标志 = 1;
    --b.合并的病人现在是医保帐户
    If v_Count > 0 Then
      Select Count(*) Into v_Count From 医保病人关联表 Where 病人id = v_保留id;
      --b.1保留的病人现在也是医保帐户
      --b.2.1保留的病人现在不是医保帐户,以前是,与b.1相同处理
      If v_Count > 0 Then
        For r_Insure In c_Insure(v_合并id) Loop
          --转移帐户年度信息
          For r_Year In c_Year(v_合并id, r_Insure.险类) Loop
            Update 帐户年度信息
            Set 帐户增加累计 = Nvl(帐户增加累计, 0) + Nvl(r_Year.帐户增加累计, 0), 帐户支出累计 = Nvl(帐户支出累计, 0) + Nvl(r_Year.帐户支出累计, 0),
                进入统筹累计 = Nvl(进入统筹累计, 0) + Nvl(r_Year.进入统筹累计, 0), 统筹报销累计 = Nvl(统筹报销累计, 0) + Nvl(r_Year.统筹报销累计, 0),
                住院次数累计 = Nvl(住院次数累计, 0) + Nvl(r_Year.住院次数累计, 0), 大额统筹累计 = Nvl(大额统筹累计, 0) + Nvl(r_Year.大额统筹累计, 0),
                起付线累计 = Nvl(起付线累计, 0) + Nvl(r_Year.起付线累计, 0), 本次起付线 = Nvl(本次起付线, r_Year.本次起付线),
                基本统筹限额 = Nvl(基本统筹限额, r_Year.基本统筹限额), 大额统筹限额 = Nvl(大额统筹限额, r_Year.大额统筹限额), 封销信息 = Nvl(封销信息, r_Year.封销信息)
            Where 病人id = v_保留id And 险类 = r_Insure.险类 And 年度 = r_Year.年度;
            If Sql%RowCount = 0 Then
              Insert Into 帐户年度信息
                (病人id, 险类, 年度, 帐户增加累计, 帐户支出累计, 进入统筹累计, 统筹报销累计, 住院次数累计, 本次起付线, 基本统筹限额, 大额统筹限额, 起付线累计, 大额统筹累计, 封销信息)
              Values
                (v_保留id, r_Insure.险类, r_Year.年度, r_Year.帐户增加累计, r_Year.帐户支出累计, r_Year.进入统筹累计, r_Year.统筹报销累计,
                 r_Year.住院次数累计, r_Year.本次起付线, r_Year.基本统筹限额, r_Year.大额统筹限额, r_Year.起付线累计, r_Year.大额统筹累计, r_Year.封销信息);
            End If;
          End Loop;
          Delete From 帐户年度信息 Where 病人id = v_合并id;
        
          Select Count(Distinct 医保号) Into v_Count From 医保病人关联表 Where 病人id In (v_合并id, v_保留id);
          If v_Count <> 2 Then
            --两个病人医保号相同时,不用处理医保病人档案
            If v_合并id = B病人id_In Then
              Update 医保病人关联表
              Set 标志 = 0
              Where (险类, 中心, 医保号) In (Select 险类, 中心, 医保号 From 医保病人关联表 Where 病人id = v_保留id);
              Update 医保病人关联表 Set 标志 = 1 Where 病人id = v_保留id;
            End If;
            Delete From 医保病人关联表 Where 病人id = v_合并id;
          Else
            --被合并的病人可能关联了多个医保病人,改为关联到保留的病人上
            --问题27445 保留指定病人的门诊号、住院号、医保号
            If v_合并id = B病人id_In Then
              Update 医保病人关联表
              Set 医保号 =
                   (Select 医保号 From 医保病人关联表 Where 病人id = v_合并id), 标志 = 0
              Where 险类 = r_Insure.险类 And 中心 = r_Insure.中心 And 医保号 = r_Insure.医保号 And 病人id <> v_合并id;
            Else
              Update 医保病人关联表
              Set 医保号 =
                   (Select 医保号 From 医保病人关联表 Where 病人id = v_保留id), 标志 = 0
              Where 险类 = r_Insure.险类 And 中心 = r_Insure.中心 And 医保号 = r_Insure.医保号 And 病人id <> v_合并id;
            End If;
            --暂存用户指定要保留病人的帐户信息
            If v_合并id = B病人id_In Then
              Open c_Keepinsure(B病人id_In, r_Insure.险类);
              Fetch c_Keepinsure
                Into r_Keepinsure;
            End If;
          
            Delete From 医保病人关联表 Where 病人id = v_合并id;
            Delete From 医保病人档案 Where 险类 = r_Insure.险类 And 医保号 = r_Insure.医保号;
          
            --保留用户指定要保留病人的帐户信息
            If v_合并id = B病人id_In Then
              If c_Keepinsure%RowCount > 0 Then
                Update 医保病人档案
                Set 卡号 = r_Keepinsure.卡号, 医保号 = r_Keepinsure.医保号, 密码 = r_Keepinsure.密码, 人员身份 = r_Keepinsure.人员身份,
                    单位编码 = r_Keepinsure.单位编码, 顺序号 = r_Keepinsure.顺序号, 退休证号 = r_Keepinsure.退休证号, 帐户余额 = r_Keepinsure.帐户余额,
                    当前状态 = r_Keepinsure.当前状态, 病种id = r_Keepinsure.病种id, 在职 = r_Keepinsure.在职, 年龄段 = r_Keepinsure.年龄段,
                    灰度级 = r_Keepinsure.灰度级, 就诊时间 = r_Keepinsure.就诊时间
                Where (险类, 中心, 医保号) In (Select 险类, 中心, 医保号 From 医保病人关联表 Where 病人id = v_保留id);
                --保留病人可能关联了多个医保病人,都要更改医保号
                Update 医保病人关联表
                Set 医保号 = r_Keepinsure.医保号, 标志 = 0
                Where (险类, 中心, 医保号) In (Select 险类, 中心, 医保号 From 医保病人关联表 Where 病人id = v_保留id);
                Update 医保病人关联表 Set 标志 = 1 Where 病人id = v_保留id;
              End If;
              Close c_Keepinsure;
            End If;
          End If;
        End Loop;
      Else
        --b.2.2保留的病人现在和以前都不是医保帐户
        Update 帐户年度信息 Set 病人id = v_保留id Where 病人id = v_合并id;
        Update 医保病人关联表 Set 病人id = v_保留id Where 病人id = v_合并id;
        --医保病人档案表不用处理,因为通过医保号关联<医保病人关联表>
      End If;
    Else
      --c.合并的病人以前和现在都不是医保帐户,不作任何处理
      Null;
    End If;
  End If;

  --处理体检子系统的病人合并
  n_Have := 0;
  Begin
    Select 1 Into n_Have From zlSystems Where Floor(编号 / 100) = 21;
  Exception
    When Others Then
      Null;
  End;
  If n_Have = 1 Then
    v_Sql := 'Begin zl21_病人信息_Merge(:1,:2); End;';
    Execute Immediate v_Sql
      Using v_合并id, v_保留id;
  End If;

  --其它病人,病案主页相关表
  For v_Loop In 1 .. Arronpage.Count Loop
    --Executesql('Update ' || Arronpage(v_Loop) || ' Set 病人ID=' || v_保留id || ' Where 病人ID=' || v_合并id || ' And Nvl(主页ID,0) = 0');
    --"主页=0，主页ID is NULL，主页ID=挂号ID"都有可能，前面部分与主页ID关联都没处理到，因此不加条件
    v_Sql := 'Update ' || Arronpage(v_Loop) || ' Set 病人ID=:1 Where 病人ID=:2';
    Execute Immediate v_Sql
      Using v_保留id, v_合并id;
  End Loop;
  For v_Loop In 1 .. Arronbase.Count Loop
    If Arronbase(v_Loop) = '病人照片' Then
      Select Count(1) Into n_Have From 病人照片 Where 病人id = v_保留id;
      If n_Have = 1 Then
        Delete From 病人照片 Where 病人id = v_合并id;
      End If;
    End If;
    v_Sql := 'Update ' || Arronbase(v_Loop) || ' Set 病人ID=:1 Where 病人ID=:2';
    Execute Immediate v_Sql
      Using v_保留id, v_合并id;
  End Loop;

  --删除实际不保留的病人信息
  Delete From 病人信息 Where 病人id = v_合并id;

  --根据界面选择保留病人信息
  Update 病人信息
  Set 姓名 = Nvl(r_Infob.姓名, r_Infoa.姓名), 性别 = Nvl(r_Infob.性别, r_Infoa.性别), 年龄 = Nvl(r_Infob.年龄, r_Infoa.年龄), 门诊号 = v_门诊号,
      住院号 = v_住院号, 就诊卡号 = Nvl(r_Infob.就诊卡号, r_Infoa.就诊卡号), 卡验证码 = Decode(r_Infob.就诊卡号, Null, r_Infoa.卡验证码, r_Infob.卡验证码),
      费别 = Nvl(r_Infob.费别, r_Infoa.费别), 医疗付款方式 = Nvl(r_Infob.医疗付款方式, r_Infoa.医疗付款方式),
      出生日期 = Nvl(r_Infob.出生日期, r_Infoa.出生日期), 出生地点 = Nvl(r_Infob.出生地点, r_Infoa.出生地点),
      身份证号 = Nvl(r_Infob.身份证号, r_Infoa.身份证号), 身份 = Nvl(r_Infob.身份, r_Infoa.身份), 职业 = Nvl(r_Infob.职业, r_Infoa.职业),
      民族 = Nvl(r_Infob.民族, r_Infoa.民族), 国籍 = Nvl(r_Infob.国籍, r_Infoa.国籍), 学历 = Nvl(r_Infob.学历, r_Infoa.学历),
      籍贯 = Nvl(r_Infob.籍贯, r_Infoa.籍贯), 区域 = Nvl(r_Infob.区域, r_Infoa.区域), 婚姻状况 = Nvl(r_Infob.婚姻状况, r_Infoa.婚姻状况),
      家庭地址 = Nvl(r_Infob.家庭地址, r_Infoa.家庭地址), 家庭电话 = Nvl(r_Infob.家庭电话, r_Infoa.家庭电话),
      家庭地址邮编 = Nvl(r_Infob.家庭地址邮编, r_Infoa.家庭地址邮编), 户口地址 = Nvl(r_Infob.户口地址, r_Infoa.户口地址),
      户口地址邮编 = Nvl(r_Infob.户口地址邮编, r_Infoa.户口地址邮编), 联系人姓名 = Nvl(r_Infob.联系人姓名, r_Infoa.联系人姓名),
      联系人关系 = Nvl(r_Infob.联系人关系, r_Infoa.联系人关系), 联系人地址 = Nvl(r_Infob.联系人地址, r_Infoa.联系人地址),
      联系人电话 = Nvl(r_Infob.联系人电话, r_Infoa.联系人电话), 合同单位id = Nvl(r_Infob.合同单位id, r_Infoa.合同单位id),
      工作单位 = Nvl(r_Infob.工作单位, r_Infoa.工作单位), 单位电话 = Nvl(r_Infob.单位电话, r_Infoa.单位电话),
      单位邮编 = Nvl(r_Infob.单位邮编, r_Infoa.单位邮编), 单位开户行 = Nvl(r_Infob.单位开户行, r_Infoa.单位开户行),
      单位帐号 = Nvl(r_Infob.单位帐号, r_Infoa.单位帐号), 就诊时间 = Nvl(r_Infob.就诊时间, r_Infoa.就诊时间),
      就诊状态 = Nvl(r_Infob.就诊状态, r_Infoa.就诊状态), 就诊诊室 = Nvl(r_Infob.就诊诊室, r_Infoa.就诊诊室), 险类 = Nvl(r_Infob.险类, r_Infoa.险类),
      登记时间 = Nvl(r_Infob.登记时间, r_Infoa.登记时间), 住院次数 = Null, 主页id = Null, 当前床号 = Null, 当前科室id = Null, 当前病区id = Null,
      入院时间 = Null, 出院时间 = Null, 在院 = Decode(Nvl(r_Infob.在院, 0), 1, 1, Null), 健康号 = Nvl(r_Infob.健康号, r_Infoa.健康号)
  Where 病人id = v_保留id;

  Open c_Info(v_保留id);
  Fetch c_Info
    Into r_Info;
  If c_Info%RowCount > 0 Then
    --最后一次为预约病人,只需要更改住院次数和入院时间
    If r_Info.主页id = 0 Then
      Update 病人信息
      Set 主页id = Decode(r_Info.最大主页id, 0, Null, r_Info.最大主页id), 住院次数 = Decode(n_Max住院次数, 0, Null, n_Max住院次数)
      Where 病人id = v_保留id;
    Else
      Update 病人信息
      Set 主页id = Decode(r_Info.最大主页id, 0, Null, r_Info.最大主页id), 住院次数 = Decode(n_Max住院次数, 0, Null, n_Max住院次数),
          当前床号 = Decode(r_Info.出院日期, Null, r_Info.出院病床, Null), 当前病区id = Decode(r_Info.出院日期, Null, r_Info.当前病区id, Null),
          当前科室id = Decode(r_Info.出院日期, Null, r_Info.出院科室id, Null), 入院时间 = r_Info.入院日期, 出院时间 = r_Info.出院日期









      
      Where 病人id = v_保留id;
    End If;
    --处理担保信息
    Select Nvl(主页id, -1) Into n_主页id From 病人信息 Where 病人id = v_保留id;
    --提取当前有效的正常担保记录,确保正常担保与临时担保不并存
    Select Nvl(Sum(担保额), 0), Count(病人id)
    Into n_担保额, n_Row
    From 病人担保记录
    Where 病人id = v_保留id And Nvl(主页id, -1) = n_主页id And (到期时间 Is Null Or 到期时间 > Sysdate) And 担保性质 = 0 And 删除标志 = 1;
    If n_Row = 0 Then
      --保留最后一条临时担保记录,其余到期
      Update 病人担保记录
      Set 到期时间 = Sysdate - 1 / 24 / 60 / 60
      Where 病人id = v_保留id And Nvl(主页id, -1) = n_主页id And 担保性质 = 1 And (到期时间 Is Null Or 到期时间 > Sysdate) And 删除标志 = 1 And
            登记时间 <> (Select Max(登记时间)
                     From 病人担保记录
                     Where 病人id = v_保留id And Nvl(主页id, -1) = n_主页id And 担保性质 = 1 And (到期时间 Is Null Or 到期时间 > Sysdate) And
                           删除标志 = 1);
    Else
      --有正常担保就让临时担保失效
      Update 病人担保记录
      Set 到期时间 = Sysdate - 1 / 24 / 60 / 60
      Where 病人id = v_保留id And Nvl(主页id, -1) = n_主页id And 担保性质 = 1 And (到期时间 Is Null Or 到期时间 > Sysdate) And 删除标志 = 1;
    End If;
  
    --提取当前有效担保额及有效担保记录数
    n_Row    := 0;
    n_担保额 := 0;
    v_担保人 := '';
    For r_提保信息 In (Select 担保人, 担保额
                   From 病人担保记录
                   Where 病人id = v_保留id And Nvl(主页id, -1) = n_主页id And (到期时间 Is Null Or 到期时间 > Sysdate) And 删除标志 = 1) Loop
      n_Row     := n_Row + 1;
      n_担保额  := n_担保额 + r_提保信息.担保额;
      v_担保人  := v_担保人 || ',' || r_提保信息.担保人;
      n_Lengthb := Lengthb(v_担保人);
      If n_Lengthb >= 101 Then
        v_Error := '不能合并担保记录，在病人信息保存时超过担保人字段长度！';
        Raise Err_Custom;
      End If;
    End Loop;
    v_担保人 := Substr(v_担保人, 2, 100);
  
    If n_Row = 0 Then
      Update 病人信息 Set 担保人 = Null, 担保额 = Null, 担保性质 = Null Where 病人id = v_保留id;
    Else
      --提取最后一条有效担保人和担保性质
      Select 担保性质
      Into n_担保性质
      From 病人担保记录
      Where 病人id = v_保留id And Nvl(主页id, -1) = n_主页id And 删除标志 = 1 And
            登记时间 =
            (Select Max(登记时间)
             From 病人担保记录
             Where 病人id = v_保留id And Nvl(主页id, -1) = n_主页id And (到期时间 Is Null Or 到期时间 > Sysdate) And 删除标志 = 1);
    
      Update 病人信息 Set 担保人 = v_担保人, 担保额 = n_担保额, 担保性质 = n_担保性质 Where 病人id = v_保留id;
    End If;
  End If;

  Close c_Info;
  Close c_Infoa;
  Close c_Infob;

  --对病人进行解锁
  Update 病人信息 Set 锁定 = 0 Where 病人id In (A病人id_In, B病人id_In);  
  v_Chgs := Substr(v_Chgs, 2);
  b_Message.Zlhis_Patient_017(v_保留id, v_合并id, v_Chgs);
Exception
  When Err_Custom Then
    Begin
      Rollback; --不然会死锁
      Zl_病人信息_锁定(A病人id_In, 0);
      Zl_病人信息_锁定(B病人id_In, 0);
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
    End;
  When Others Then
    Begin
      Rollback; --不然会死锁
      Zl_病人信息_锁定(A病人id_In, 0);
      Zl_病人信息_锁定(B病人id_In, 0);
      zl_ErrorCenter(SQLCode, SQLErrM);
    End;
End Zl_病人信息_Merge;
/



------------------------------------------------------------------------------------
--系统版本号
Update zlSystems Set 版本号='10.35.90.0049' Where 编号=&n_System;
Commit;