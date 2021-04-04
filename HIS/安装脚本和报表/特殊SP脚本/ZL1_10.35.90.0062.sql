----------------------------------------------------------------------------------------------------------------
--本脚本支持从ZLHIS+ v10.35.90升级到 v10.35.90
--请以数据空间的所有者登录PLSQL并执行下列脚本
Define n_System=100;
----------------------------------------------------------------------------------------------------------------
----------------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------
--结构修正部份
------------------------------------------------------------------------------
--126017:蒋敏,2019-05-21,诊疗项目管理检查部位方法的选择处理
Alter Table 诊疗项目部位 Add 上级方法 Varchar2(30);
Alter Table 诊疗项目部位 Drop Constraint 诊疗项目部位_UQ_项目id Cascade Drop Index;
Alter Table 诊疗项目部位 Add Constraint 诊疗项目部位_UQ_项目id Unique(项目id,部位,方法,类型,上级方法)Using Index Tablespace Zl9indexhis;

--128157:蒋敏,2019-05-21,诊疗项目部位添加外键索引
Create Index 诊疗项目部位_IX_部位 on 诊疗项目部位(部位,类型) Tablespace zl9indexhis;
------------------------------------------------------------------------------
--数据修正部份
------------------------------------------------------------------------------
--126017:蒋敏,2019-05-21,诊疗项目管理检查部位增加上级方法锚点处理
Update Zlmsg_Lists
Set Key_Define = '<root><ID></ID><项目ID></项目ID><类型></类型><部位></部位><方法></方法><默认></默认><上级方法></上级方法></root>'
Where Code In ('ZLHIS_DICTPACS_007', 'ZLHIS_DICTPACS_008', 'ZLHIS_DICTPACS_009');
-------------------------------------------------------------------------------
--权限修正部份
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--报表修正部份
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--过程修正部份
-------------------------------------------------------------------------------
--126017:蒋敏,2019-05-21,诊疗项目管理检查部位增加上级方法锚点处理
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
    默认_In   诊疗项目部位.默认%Type,
    上级方法_In 诊疗项目部位.上级方法%Type
  );
  --修改诊疗项目部位
  Procedure Zlhis_Dictpacs_008
  (
    Id_In     诊疗项目部位.Id%Type,
    项目id_In 诊疗项目部位.项目id%Type,
    类型_In   诊疗项目部位.类型%Type,
    部位_In   诊疗项目部位.部位%Type,
    方法_In   诊疗项目部位.方法%Type,
    默认_In   诊疗项目部位.默认%Type,
    上级方法_In 诊疗项目部位.上级方法%Type
  );
  --删除诊疗项目部位
  Procedure Zlhis_Dictpacs_009
  (
    Id_In     诊疗项目部位.Id%Type,
    项目id_In 诊疗项目部位.项目id%Type,
    类型_In   诊疗项目部位.类型%Type,
    部位_In   诊疗项目部位.部位%Type,
    方法_In   诊疗项目部位.方法%Type,
    默认_In   诊疗项目部位.默认%Type,
    上级方法_In 诊疗项目部位.上级方法%Type
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


  --79.留观病人转住院病人
  Procedure Zlhis_Patient_029
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  );


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
  --管理工具上机人员变动消息
  Procedure Zltools_Users_001
  (
    用户名_In In 上机人员表.用户名%Type,
    人员id_In In 上机人员表.人员id%Type
  );
  Procedure Zltools_Users_002
  (
    用户名_In In 上机人员表.用户名%Type,
    人员id_In In 上机人员表.人员id%Type
  );
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
    默认_In   诊疗项目部位.默认%Type,
    上级方法_In 诊疗项目部位.上级方法%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ID>' || Id_In || '</ID><项目ID>' || 项目id_In || '</项目ID><类型>' || 类型_In || '</类型><部位>' || 部位_In ||
               '</部位><方法>' || 方法_In || '</方法><默认>' || 默认_In || '</默认><上级方法>' || 上级方法_In || '</上级方法></root>';
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
    默认_In   诊疗项目部位.默认%Type,
    上级方法_In 诊疗项目部位.上级方法%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ID>' || Id_In || '</ID><项目ID>' || 项目id_In || '</项目ID><类型>' || 类型_In || '</类型><部位>' || 部位_In ||
               '</部位><方法>' || 方法_In || '</方法><默认>' || 默认_In || '</默认><上级方法>' || 上级方法_In || '</上级方法></root>';
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
    默认_In   诊疗项目部位.默认%Type,
    上级方法_In 诊疗项目部位.上级方法%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ID>' || Id_In || '</ID><项目ID>' || 项目id_In || '</项目ID><类型>' || 类型_In || '</类型><部位>' || 部位_In ||
               '</部位><方法>' || 方法_In || '</方法><默认>' || 默认_In || '</默认><上级方法>' || 上级方法_In || '</上级方法></root>';
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

  --79.留观病人转住院病人
  Procedure Zlhis_Patient_029
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  ) Is
    n_变动id Number(18);
  Begin
    Select max(ID)
    Into n_变动id
    From 病人变动记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 Is Null And Nvl(附加床位, 0) = 0 And 开始原因 = 9;
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_029',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><变动ID>' || n_变动id ||
                                 '</变动ID></root>');
  
  End Zlhis_Patient_029;

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
  --管理工具上机人员变动消息
  Procedure Zltools_Users_001
  (
    用户名_In In 上机人员表.用户名%Type,
    人员id_In In 上机人员表.人员id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><用户名>' || 用户名_In || '</用户名><人员ID>' || 人员id_In || '</人员ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLTOOLS_USERS_001', v_Value);
  End Zltools_Users_001;
  Procedure Zltools_Users_002
  (
    用户名_In In 上机人员表.用户名%Type,
    人员id_In In 上机人员表.人员id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><用户名>' || 用户名_In || '</用户名><人员ID>' || 人员id_In || '</人员ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLTOOLS_USERS_002', v_Value);
  End Zltools_Users_002;
End b_Message;
/
--126017:蒋敏,2019-05-21,诊疗项目管理检查部位方法的选择处理
Create Or Replace Procedure Zl_诊疗项目部位_Delete(项目id_In In 诊疗项目部位.项目id%Type) As
Begin
  For R In (Select a.Id, a.项目id, a.类型, a.部位, a.方法, a.默认,a.上级方法 From 诊疗项目部位 A Where a.项目id = 项目id_In) Loop
    b_Message.Zlhis_Dictpacs_009(r.Id, r.项目id, r.类型, r.部位, r.方法, r.默认,r.上级方法);
  End Loop;
  Delete 诊疗项目部位 Where 项目id = 项目id_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_诊疗项目部位_Delete;
/
--126017:蒋敏,2019-05-21,诊疗项目管理检查部位方法的选择处理
CREATE OR REPLACE Procedure Zl_诊疗项目部位_Insert
(
  项目id_In In 诊疗项目部位.项目id%Type,
  类型_In   In 诊疗项目部位.类型%Type,
  部位_In   In 诊疗项目部位.部位%Type,
  方法_In   In 诊疗项目部位.方法%Type,
  默认_In   In 诊疗项目部位.默认%Type := Null,
  上级方法_In In 诊疗项目部位.上级方法%Type := Null
) As
  v_Code Varchar2(20); --编码
  Err_Notfind Exception;
  n_记录id 诊疗项目部位.Id%Type;
Begin
  Select RTrim(编码) Into v_Code From 诊疗项目目录 Where 类别 = 'D' And ID = 项目id_In;
  If v_Code Is Null Then
    Raise Err_Notfind;
  End If;
  Select 诊疗项目部位_Id.Nextval Into n_记录id From Dual;
  Insert Into 诊疗项目部位
    (ID, 项目id, 类型, 部位, 方法, 默认,上级方法)
  Values
    (n_记录id, 项目id_In, 类型_In, 部位_In, 方法_In, 默认_In,上级方法_In);
  b_Message.Zlhis_Dictpacs_007(n_记录id, 项目id_In, 类型_In, 部位_In, 方法_In, 默认_In,上级方法_In);
Exception
  When Err_Notfind Then
    Raise_Application_Error(-20101, '[ZLSOFT]该项目不存在，可能已被其他用户删除！[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_诊疗项目部位_Insert;
/
--137245:蒋敏,2019-05-20,过程中解析方法解析不对，导致方法前面多0或者多1
--126017:蒋敏,2019-05-21,诊疗项目管理检查部位方法的选择处理
CREATE OR REPLACE Procedure Zl_诊疗检查部位_Edit
(
  操作_In     In Number, --1:增加;2:修改;3:删除
  类型_In     In 诊疗检查部位.类型%Type,
  原编码_In   In 诊疗检查部位.编码%Type,
  新编码_In   In 诊疗检查部位.编码%Type := Null,
  名称_In     In 诊疗检查部位.名称%Type := Null,
  分组_In     In 诊疗检查部位.分组%Type := Null,
  备注_In     In 诊疗检查部位.备注%Type := Null,
  方法_In     In 诊疗检查部位.方法%Type := Null,
  适用性别_In In 诊疗检查部位.适用性别%Type := Null,
  上级方法_In In 诊疗检查部位.方法%Type := Null --格式：上级方法|方法;上级方法|方法...(若上级方法为空，则为|方法，若为同一个上级，则用逗号分隔)
) Is
  v_原名称 诊疗检查部位.名称%Type := Null;
  e_Notfind Exception;
  v_方法   Varchar2(1000);
  v_Fields Varchar2(1000);
  v_Tmp    Varchar2(1000);
  n_Count  Number;
  n_记录id 诊疗项目部位.Id%Type;
  v_上级 Varchar(100);
  v_格式方法 Varchar2(1000);
Begin
  If 操作_In = 1 Then
    Insert Into 诊疗检查部位
      (类型, 编码, 名称, 分组, 备注, 方法, 适用性别)
    Values
      (类型_In, 新编码_In, 名称_In, 分组_In, 备注_In, 方法_In, 适用性别_In);
    b_Message.Zlhis_Dictpacs_004(类型_In, 新编码_In, 名称_In, 分组_In, 备注_In, 方法_In, 适用性别_In);
  Elsif 操作_In = 2 Then
    Begin
      Select 名称 Into v_原名称 From 诊疗检查部位 Where 编码 = 原编码_In And 类型 = 类型_In;
    Exception
      When Others Then
        Null;
    End;
    If v_原名称 Is Null Then
      Raise e_Notfind;
    End If;
    Update 诊疗检查部位
    Set 编码 = 新编码_In, 名称 = 名称_In, 分组 = 分组_In, 备注 = 备注_In, 方法 = 方法_In, 适用性别 = 适用性别_In
    Where 编码 = 原编码_In And 类型 = 类型_In;
    b_Message.Zlhis_Dictpacs_005(类型_In, 新编码_In, 名称_In, 分组_In, 备注_In, 方法_In, 适用性别_In);
    --级联修改
    --原来没有的方法现在新增
    v_方法 := ';' || Replace(方法_In,Chr(32),';');
    v_方法 := Replace(v_方法, ';;', ';');
    v_方法 := Replace(v_方法, ';;', ';');  
    v_方法 := Replace(v_方法, ',', Chr(10));
    v_方法 := Replace(v_方法, Chr(9), ';');
    v_方法 := Replace(v_方法, ';0', Chr(10) || '(上级)0');
    v_方法 := Replace(v_方法, ';1', Chr(10) || '(上级)1');
    v_方法 := Replace(v_方法, Chr(10), ';');
    v_方法 := Replace(v_方法, ';;', ';');
    v_方法 := v_方法 || ';';
    v_方法 := Substr(v_方法, 2);
    While v_方法 Is Not Null Loop
      --依次取每个项目v_Tmp
      v_Fields := Substr(v_方法, 1, Instr(v_方法, ';') - 1);
      v_方法    := Substr(v_方法, Instr(v_方法, ';') + 1);
      If Substr(v_Fields, 1, 4) = '(上级)' Then
        v_Fields := Substr(v_Fields, 5);
        v_Fields:=Substr(v_Fields, 2);
        v_Tmp:=v_Fields;
        v_上级:=NULL;
      Else
        v_Fields:=Substr(v_Fields, 2);
        v_上级   := v_Tmp;
      End If;
      If v_Fields Is Not Null Then
      v_格式方法:=v_格式方法 ||';'|| v_Fields ||','|| Nvl(v_上级,' ');
        For r_Used In (Select Distinct 项目id From 诊疗项目部位 Where 部位 = 名称_In And 类型 = 类型_In) Loop
          Select Count(ID)
          Into n_Count
          From 诊疗项目部位
          Where 项目id = r_Used.项目id And 部位 = 名称_In And 类型 = 类型_In And 方法 = v_Fields And Nvl(上级方法,' ')=Nvl(v_上级,' ');
          If n_Count = 0 Then
            Select 诊疗项目部位_Id.Nextval Into n_记录id From Dual;
            Insert Into 诊疗项目部位
              (ID, 项目id, 类型, 部位, 方法,上级方法)
            Values
              (n_记录id, r_Used.项目id, 类型_In, 名称_In, v_Fields,v_上级);
            b_Message.Zlhis_Dictpacs_007(n_记录id, r_Used.项目id, 类型_In, 名称_In, v_Fields, Null,v_上级);
          End If;
        End Loop;
      End If;
    End Loop;
    --原有的方法，现在已经删除了或原有的部位的名称已经改变了
    v_格式方法 :=v_格式方法|| ';';
    For r_Used In (Select ID, 项目id, 部位, 方法, 类型, 默认,上级方法 From 诊疗项目部位 Where 部位 = v_原名称 And 类型 = 类型_In) Loop
      If Instr(v_格式方法, ';' || r_Used.方法 ||','|| Nvl(r_Used.上级方法,' ') || ';') = 0 Then
        b_Message.Zlhis_Dictpacs_009(r_Used.Id, r_Used.项目id, r_Used.类型, r_Used.部位, r_Used.方法, r_Used.默认,r_Used.上级方法);
        Delete 诊疗项目部位
        Where id=r_Used.id;
      Else
        Update 诊疗项目部位
        Set 部位 = 名称_In
        Where id=r_Used.id;
        b_Message.Zlhis_Dictpacs_008(r_Used.Id, r_Used.项目id, r_Used.类型, r_Used.部位, r_Used.方法, r_Used.默认,r_Used.上级方法);
      End If;
    End Loop;
  Elsif 操作_In = 3 Then
    For R In (Select a.类型, a.编码, a.名称, a.分组, a.备注, a.方法, a.适用性别
              From 诊疗检查部位 A
              Where a.编码 = 原编码_In And a.类型 = 类型_In) Loop
      b_Message.Zlhis_Dictpacs_006(r.类型, r.编码, r.名称, r.分组, r.备注, r.方法, r.适用性别);
    End Loop;
    Delete 诊疗检查部位 Where 编码 = 原编码_In And 类型 = 类型_In;
  End If;
Exception
  When e_Notfind Then
    Raise_Application_Error(-20101, '[ZLSOFT]该部位不存在，可能已被其他用户删除修改！[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_诊疗检查部位_Edit;
/
------------------------------------------------------------------------------------
--系统版本号
Update zlSystems Set 版本号='10.35.90.0062' Where 编号=&n_System;
Commit;
