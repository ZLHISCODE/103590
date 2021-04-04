------------------------------------------------------------
--数据组成分析：
------------------------------------------------------------
--  1.  诊疗分类目录
--  2.  诊疗项目目录
--  3.  诊疗项目别名
--  4.  检验报告项目：目前存放检验标本，今后随LIS取消，修改为诊疗用法用量
--  5.  诊疗收费关系
------------------------------------------------------------
--  1.  诊疗分类目录
Insert Into 诊疗分类目录 (id, 编码, 名称, 简码, 上级id, 类型)
Select 1,'11','一般诊疗项目','YBZLXM',-Null,5 From Dual Union All
Select 2,'1101','护理操作常规','HLCZCG',1,5 From Dual Union All
Select 3,'1102','过敏试验','GMSY',1,5 From Dual Union All
Select 4,'1103','给药与注射','GYYZS',1,5 From Dual Union All
Select 5,'1104','治疗处置','ZLCZ',1,5 From Dual Union All
Select 6,'1105','膳食','SS',1,5 From Dual Union All
Select 7,'1106','其他','QT',1,5 From Dual Union All
Select 8,'21','X线诊断','XXZD',-Null,5 From Dual Union All
Select 9,'2101','X线摄影','XXSY',8,5 From Dual Union All
Select 10,'2102','胸部X线检查','XBXXJC',8,5 From Dual Union All
Select 11,'2103','循环系统X线检查','XHXTXXJC',8,5 From Dual Union All
Select 12,'2104','消化系统X线检查','XHXTXXJC',8,5 From Dual Union All
Select 13,'2105','泌尿生殖系统','MNSZXT',8,5 From Dual Union All
Select 14,'2106','其它常规X线检查','QTCGXXJC',8,5 From Dual Union All
Select 15,'22','介入放射','JRFS',-Null,5 From Dual Union All
Select 16,'2201','介入放射常用术','JRFSCYS',15,5 From Dual Union All
Select 17,'2202','常见部位介入诊疗','CJBWJRZL',15,5 From Dual Union All
Select 18,'23','计算机体层扫描（CT）','JSJTCSM_CT',-Null,5 From Dual Union All
Select 19,'2301','各部位传统CT检查','GBWCTCTJC',18,5 From Dual Union All
Select 20,'2302','螺旋CT和高分辨率CT扫描','LXCTHGFBLC',18,5 From Dual Union All
Select 21,'2303','电子束CT(EBCT)','DZSCT(EBCT',18,5 From Dual Union All
Select 22,'24','磁共振成像（MRI）','CGZCX_MRI_',-Null,5 From Dual Union All
Select 23,'25','超声类诊疗','CSLZL',-Null,5 From Dual Union All
Select 24,'2501','各部位超声检查','GBWCSJC',23,5 From Dual Union All
Select 25,'25011','常规部位超声检查','CGBWCSJC',24,5 From Dual Union All
Select 26,'25012','其他部位超声检查','QTBWCSJC',24,5 From Dual Union All
Select 27,'25013','彩超特殊检查','CCTSJC',24,5 From Dual Union All
Select 28,'2502','腔内超声检查','QNCSJC',23,5 From Dual Union All
Select 29,'2503','介入超声','JRCS',23,5 From Dual Union All
Select 30,'2504','其他超声诊疗技术','QTCSZLJS',23,5 From Dual Union All
Select 31,'26','常用检验项目','CYJYXM',-Null,5 From Dual Union All
Select 32,'2601','临床检验与其他速检','LCJYYQTSJ',31,5 From Dual Union All
Select 33,'2602','临床化学检验','LCHXJY',31,5 From Dual Union All
Select 34,'2603','免疫学检','MYXJ',31,5 From Dual Union All
Select 35,'2604','放射免疫学检验','FSMYXJY',31,5 From Dual Union All
Select 36,'2605','病原微生物学检验','BYWSWXJY',31,5 From Dual Union All
Select 37,'27','血型配血与输血','XXPXYSX',-Null,5 From Dual Union All
Select 38,'28','常用病理项目','CYBLXM',-Null,5 From Dual Union All
Select 39,'2801','细胞病理学检查与诊断','',38,5 From Dual Union All
Select 40,'2802','组织病理学检查与诊断','',38,5 From Dual Union All
Select 41,'2803','冰冻切片与石蜡切片','',38,5 From Dual Union All
Select 42,'2804','常用特殊染色及组织化学技术','',38,5 From Dual Union All
Select 43,'2805','尸体解剖','',38,5 From Dual Union All
Select 44,'31','手术治疗','SSZL',-Null,5 From Dual Union All
Select 45,'3101','麻醉方式','MZFS',44,5 From Dual Union All
Select 46,'3102','神经系统手术','SJXTSS',44,5 From Dual Union All
Select 47,'31021','颅骨和脑手术','LGHNSS',46,5 From Dual Union All
Select 48,'31022','颅脑神经手术','LNSJSS',46,5 From Dual Union All
Select 49,'31023','脑、脑膜肿瘤切除术','N_NMZLQCS',46,5 From Dual Union All
Select 50,'31024','脊髓、椎管手术','JS_ZGSS',46,5 From Dual Union All
Select 51,'3103','内分泌系统手术','NFMXTSS',44,5 From Dual Union All
Select 52,'31031','甲状腺、甲状旁术','JZX_JZPS',51,5 From Dual Union All
Select 53,'31032','肾上腺、垂体腺、胸腺手术','SSX_CTX_XX',51,5 From Dual Union All
Select 54,'3104','眼部手术','YBSS',44,5 From Dual Union All
Select 55,'31041','眼睑手术','YJSS',54,5 From Dual Union All
Select 56,'31042','泪器手术','LQSS',54,5 From Dual Union All
Select 57,'31043','结膜手术','JMSS',54,5 From Dual Union All
Select 58,'31044','角膜手术','JMSS',54,5 From Dual Union All
Select 59,'31045','虹膜、睫状体、巩膜和前房手术','HM_JZT_GMH',54,5 From Dual Union All
Select 60,'31046','晶状体手术','JZTSS',54,5 From Dual Union All
Select 61,'31047','视网膜、脉络膜、玻璃体、后房手术','SWM_MLM_BL',54,5 From Dual Union All
Select 62,'31048','眼眶和眼球手术','YKHYQSS',54,5 From Dual Union All
Select 63,'3105','耳部手术','EBSS',44,5 From Dual Union All
Select 64,'31051','外耳手术','WESS',63,5 From Dual Union All
Select 65,'31052','中耳、内耳及其他内耳及其他耳部手术','ZE_NEJQTNE',63,5 From Dual Union All
Select 66,'3106','鼻、口、咽部手术','B_K_YBSS',44,5 From Dual Union All
Select 67,'31061','鼻部手术','BBSS',66,5 From Dual Union All
Select 68,'31062','咽部手术','YBSS',66,5 From Dual Union All
Select 69,'31063','鼻窦手术','BDSS',66,5 From Dual Union All
Select 70,'31064','口腔颌面手术','KQHMSS',66,5 From Dual Union All
Select 71,'31065','扁桃体和腺样体手术','BTTHXYTSS',66,5 From Dual Union All
Select 72,'3107','呼吸系统手术','HXXTSS',44,5 From Dual Union All
Select 73,'31071','喉及气管手术','HJQGSS',72,5 From Dual Union All
Select 74,'31072','肺和支气管手术','FHZQGSS',72,5 From Dual Union All
Select 75,'31073','胸壁、胸膜、纵隔、横隔膜手术','XB_XM_ZG_H',72,5 From Dual Union All
Select 76,'31074','喉切除术','HQCS',72,5 From Dual Union All
Select 77,'3108','心脏及血管系统手术','XZJXGXTSS',44,5 From Dual Union All
Select 78,'31081','心瓣膜和心间隔手术','XBMHXJGSS',77,5 From Dual Union All
Select 79,'31082','心脏血管手术','XZXGSS',77,5 From Dual Union All
Select 80,'31083','其他血管手术','QTXGSS',77,5 From Dual Union All
Select 81,'31084','心瓣膜、心隔手术','XBM_XGSS',77,5 From Dual Union All
Select 82,'31085','心血管手术','XXGSS',77,5 From Dual Union All
Select 83,'3109','血液及淋巴系统手术','XYJLBXTSS',44,5 From Dual Union All
Select 84,'31091','淋巴相关手术','LBXGSS',83,5 From Dual Union All
Select 85,'31092','骨髓移植与脾手术','GSYZYPSS',83,5 From Dual Union All
Select 86,'3110','消化系统手术','XHXTSS',44,5 From Dual Union All
Select 87,'31101','食管手术','SGSS',86,5 From Dual Union All
Select 88,'31102','胃手术','WSS',86,5 From Dual Union All
Select 89,'31103','肠手术及检查','CSSJJC',86,5 From Dual Union All
Select 90,'31104','肠外置造口修正闭合与阑尾切除','CWZZKXZBHY',86,5 From Dual Union All
Select 91,'31105','直肠肛门手术','ZCGMSS',86,5 From Dual Union All
Select 92,'31106','肝胆手术','GDSS',86,5 From Dual Union All
Select 93,'31107','胰腺手术','YXSS',86,5 From Dual Union All
Select 94,'31108','疝修补术','SXBS',86,5 From Dual Union All
Select 95,'31109','其他腹部手术','QTFBSS',86,5 From Dual Union All
Select 96,'3111','泌尿系统手术','MNXTSS',44,5 From Dual Union All
Select 97,'31111','肾脏与肾盂手术','SZYSYSS',96,5 From Dual Union All
Select 98,'31112','输尿管手术','SNGSS',96,5 From Dual Union All
Select 99,'31113','膀胱手术','BGSS',96,5 From Dual Union All
Select 100,'31114','尿道手术','NDSS',96,5 From Dual Union All
Select 101,'31115','其他泌尿系统手术','QTMNXTSS',96,5 From Dual Union All
Select 102,'3112','男性生殖系统手术','NXSZXTSS',44,5 From Dual Union All
Select 103,'31121','前列腺、精囊腺手术','QLX_JNXSS',102,5 From Dual Union All
Select 104,'31122','阴囊、睾丸手术','YN_GWSS',102,5 From Dual Union All
Select 105,'31123','附睾、输精管、精索手术','FG_SJG_JSS',102,5 From Dual Union All
Select 106,'31124','阴茎手术','YJSS',102,5 From Dual Union All
Select 107,'3113','女性生殖系统手术','NXSZXTSS',44,5 From Dual Union All
Select 108,'31131','卵巢手术','LCSS',107,5 From Dual Union All
Select 109,'31132','输卵管手术','SLGSS',107,5 From Dual Union All
Select 110,'31133','子宫手术','ZGSS',107,5 From Dual Union All
Select 111,'31134','阴道手术','YDSS',107,5 From Dual Union All
Select 112,'31135','外阴手术','WYSS',107,5 From Dual Union All
Select 113,'3114','产科手术与操作','CKSSYCZ',44,5 From Dual Union All
Select 114,'31141','臀位、器械分娩术','TW_QXFMS',113,5 From Dual Union All
Select 115,'31142','引产助产术','YCZCS',113,5 From Dual Union All
Select 116,'31143','刮腹取胎术','GFQTS',113,5 From Dual Union All
Select 117,'31144','其他产科手术','QTCKSS',113,5 From Dual Union All
Select 118,'3115','肌肉骨骼系统手术','JRGGXTSS',44,5 From Dual Union All
Select 119,'31151','面部骨手术','MBGSS',118,5 From Dual Union All
Select 120,'31152','脊柱四肢骨手术','JZSZGSS',118,5 From Dual Union All
Select 121,'31153','切开复位松懈术','QKFWSXS',118,5 From Dual Union All
Select 122,'31154','关节融合术','GJRHS',118,5 From Dual Union All
Select 123,'31155','手部肌肉手术','SBJRSS',118,5 From Dual Union All
Select 124,'31156','其他肌肉手术','QTJRSS',118,5 From Dual Union All
Select 125,'31157','截肢与再植手术','JZYZZSS',118,5 From Dual Union All
Select 126,'3116','体被系统手术','TBXTSS',44,5 From Dual Union All
Select 127,'31161','乳房手术','RFSS',126,5 From Dual Union All
Select 128,'31162','皮肤和皮下组织手术','PFHPXZZSS',126,5 From Dual Union All
Select 129,'32','临床各科系统','LCGKXT',-Null,5 From Dual Union All
Select 130,'3201','神经系统','SJXT',129,5 From Dual Union All
Select 131,'3202','呼吸系统','HXXT',129,5 From Dual Union All
Select 132,'3203','心脏及血管系统','XZJXGXT',129,5 From Dual Union All
Select 133,'3204','血液和淋巴系统','XYHLBXT',129,5 From Dual Union All
Select 134,'3205','消化系统','XHXT',129,5 From Dual Union All
Select 135,'3206','泌尿系统','MNXT',129,5 From Dual Union All
Select 136,'3207','妇产科与新生儿诊疗','FCKYXSEZL',129,5 From Dual Union All
Select 137,'3208','精神心理卫生','JSXLWS',129,5 From Dual Union All
Select 138,'32081','常用心理测验量表','CYXLCYLB',137,5 From Dual Union All
Select 139,'32082','精神科治疗项目','JSKZLXM',137,5 From Dual Union All
Select 140,'33','常用核医学技术','CYHYXJS',-Null,5 From Dual Union All
Select 141,'3301','影像和功能','YXHGN',140,5 From Dual Union All
Select 142,'3302','核素治疗','HSZL',140,5 From Dual Union All
Select 143,'34','常用放射治疗','CYFSZL',-Null,5 From Dual Union All
Select 144,'3401','外照射治疗','WZSZL',143,5 From Dual Union All
Select 145,'3402','后装治疗','HZZL',143,5 From Dual Union All
Select 146,'3403','其他','QT',143,5 From Dual Union All
Select 147,'35','常用理疗技术','CYLLJS',-Null,5 From Dual Union All
Select 148,'3501','电疗法','DLF',147,5 From Dual Union All
Select 149,'3502','光疗法','GLF',147,5 From Dual Union All
Select 150,'3503','超声波疗法','CSBLF',147,5 From Dual Union All
Select 151,'3504','其他理疗','QTLL',147,5 From Dual Union All
Select 152,'36','康复检查与训练','KFJCYXL',-Null,5 From Dual Union All
Select 153,'41','中医诊疗项目','ZYZLXM',-Null,5 From Dual Union All
Select 154,'4101','中药煎法服法','ZYJFFF',153,5 From Dual Union All
Select 155,'4102','中医针灸、外治与推拿','ZYZJ_WZYTN',153,5 From Dual Union All
Select 156,'4103','中医骨伤治疗','ZYGSZL',153,5 From Dual Union All
Select 157,'4104','其他中医疗法','QTZYLF',153,5 From Dual;

Update 诊疗分类目录 Set 建档时间=Sysdate Where 建档时间 IS Null;

--  2.  诊疗项目目录
Insert Into 诊疗项目目录(类别,分类id,id,编码,名称,计算单位,标本部位,计算方式,单独应用,执行频率,适用性别,组合项目,操作类型,执行安排,执行科室,服务对象,计价性质,建档时间,撤档时间)
Select 'H',2,1,'110100001','特级护理','','',0,1,2,0,0,'1',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,2,'110100002','一级护理','','',0,1,2,0,0,'1',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,3,'110100003','二级护理','','',0,1,2,0,0,'1',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,4,'110100004','三级护理','','',0,1,2,0,0,'1',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,5,'110100005','ICU护理常规','','',0,1,2,0,0,'1',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,6,'110100006','按内科护理常规','','',0,1,2,0,0,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,7,'110100007','按外科护理常规','','',0,1,2,0,0,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,8,'110100008','按普通外科护理常规','','',0,1,2,0,0,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,9,'110100009','按心外科护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,10,'110100010','按儿科护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,11,'110100011','按儿外科护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,12,'110100012','按儿外科术后护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,13,'110100013','按儿科肺炎护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,14,'110100014','按儿科血液病护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,15,'110100015','按新生儿护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,16,'110100016','按妇科护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,17,'110100017','按产科护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,18,'110100018','按眼科护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,19,'110100019','按耳鼻喉科护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,20,'110100020','按皮肤科护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,21,'110100021','按中医科护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,22,'110100022','按胸外科护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,23,'110100023','按神经外科护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,24,'110100024','按眼科术后护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,25,'110100025','按耳鼻喉科术后护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,26,'110100026','按脑系科护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,27,'110100027','按肝胆外科护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,28,'110100028','按康复医学科护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,29,'110100029','按泌尿外科护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,30,'110100030','按肾科护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,31,'110100031','按昏迷护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,32,'110100032','按急性心肌梗塞护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,33,'110100033','按急性胰腺炎护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,34,'110100034','按高热护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,35,'110100035','按高血压病护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,36,'110100036','按心脏病护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,37,'110100037','按肾脏病护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,38,'110100038','按糖尿病护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,39,'110100039','按血液病护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,40,'110100040','按早产儿护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,41,'110100041','按子痫护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,42,'110100042','按癫痫护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,43,'110100043','按尿崩症护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,44,'110100044','按偏瘫护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,45,'110100045','按褥疮护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,46,'110100046','按肝硬化腹水护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,47,'110100047','按截瘫护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,48,'110100048','按鼻腔后鼻孔出血填塞后护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,49,'110100049','按鼻饲管护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,50,'110100050','按层流室护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,51,'110100051','按上消化道出血护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,52,'110100052','按二囊三腔管护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,53,'110100053','按留置肛管护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,54,'110100054','按假肛护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,55,'110100055','按留置尿管护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,56,'110100056','按左右输尿管导管护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,57,'110100057','ERCP术后护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,58,'110100058','按PTCA术后护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,59,'110100059','按TAE术后护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,60,'110100060','按Tipss术后护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,61,'110100061','按X-刀术后护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,62,'110100062','按大静脉置管术后护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,63,'110100063','按大静脉穿刺置管术后护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,64,'110100064','按肺动脉栓塞术后护理常规护理','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,65,'110100065','按基底动脉瘤栓塞术后护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,66,'110100066','按局麻颈动脉造影,栓塞术后护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,67,'110100067','按脑部蛛网膜下腔持续引流护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,68,'110100068','按漂浮导管插管术后护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,69,'110100069','按经鼻气管插管护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,70,'110100070','按气管插管术后护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,71,'110100071','按气管切开术后护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,72,'110100072','按呼吸机使用护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,73,'110100073','按微波治疗术后护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,74,'110100074','按胃造瘘管护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,75,'110100075','按胃造瘘术后护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,76,'110100076','按心脏病术后护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,77,'110100077','按腹腔负压吸引管护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,78,'110100078','按脑室引流护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,79,'110100079','按头皮下引流管护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,80,'110100080','按胸腔闭式引流护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,81,'110100081','按血肿腔引流管护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,82,'110100082','按硬膜外引流管护理','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,83,'110100083','按硬脑膜下引流护理','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,84,'110100084','按肿瘤腔硬膜外引流护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,85,'110100085','按腰穿持续引流护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,86,'110100086','按硬脊膜外引流护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,87,'110100087','按全麻术后护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,88,'110100088','按全麻下腹腔镜检查术后护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,89,'110100089','按会阴切口感染护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,90,'110100090','按前置胎盘护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,91,'110100091','按妊高征护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,92,'110100092','按产褥感染护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,93,'110100093','按羊水早破护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,94,'110100094','按局麻下输卵管结扎术后护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,95,'110100095','按硬膜外麻醉术后常规护理','','',0,1,2,0,0,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,96,'110100096','按断指再植术后护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,97,'110100097','按右下肢长腿前后石膏托固定术后护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,98,'110100098','按右下肢前后石膏托术后护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,99,'110100099','按右小腿石膏托固定术后护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,100,'110100100','按右髋人字石膏护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,101,'110100101','按髋人字石膏术后护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,102,'110100102','按石膏绷带术后护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,103,'110100103','按头颈胸石膏绷带术后护理常规','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,104,'110100104','记出入量','','',0,1,2,0,0,'0',0,2,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,105,'110100105','观察瞳孔、生命体征','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,106,'110100106','观察呼吸','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,107,'110100107','观察患肢感觉及血循环','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,108,'110100108','观察右上颈区皮瓣颜色、温度','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,109,'110100109','记痰量','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,110,'110100110','置隔离婴儿室观察生命体征','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual;
Insert Into 诊疗项目目录(类别,分类id,id,编码,名称,计算单位,标本部位,计算方式,单独应用,执行频率,适用性别,组合项目,操作类型,执行安排,执行科室,服务对象,计价性质,建档时间,撤档时间)
Select 'H',2,111,'110100111','观察意识、生命体征、瞳孔变化','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,112,'110100112','观察足背皮肤血运','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,113,'110100113','观察生命体征','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,114,'110100114','观察肢体活动','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,115,'110100115','记尿量','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,116,'110100116','观察意识','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,117,'110100117','观察阴道出血情况','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,118,'110100118','分别记左右输尿管导管尿量','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,119,'110100119','观察宫缩及阴道分泌物','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,120,'110100120','观察患指血循环','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,121,'110100121','观察皮瓣颜色','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,122,'110100122','记三天热量','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,123,'110100123','观察血压','','',0,1,2,0,0,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,124,'110100124','记每小时尿量','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,125,'110100125','记阴道出血量','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,126,'110100126','测血压、脉搏、呼吸','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,127,'110100127','测体温、脉搏、血压','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,128,'110100128','测血压、呼吸、脉搏、瞳孔','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,129,'110100129','测基础体温','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,130,'110100130','测口温','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,131,'110100131','测肛温','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,132,'110100132','测体温','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,133,'110100133','测体重','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,134,'110100134','测血压','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,135,'110100135','呼吸监测','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,136,'110100136','测心率','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,137,'110100137','测心脉差','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,138,'110100138','血压监测','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,139,'110100139','持续血压监测','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,140,'110100140','胎心监测','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,141,'110100141','测腹围','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,142,'110100142','测残余尿','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,143,'110100143','洗头','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,144,'110100144','洗双手','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,145,'110100145','生理盐水洗眼','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,146,'110100146','清洗会阴','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,147,'110100147','外阴清洗','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,148,'110100148','清洗阴茎','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,149,'110100149','保护性隔离','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,150,'110100150','便盆隔离','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,151,'110100151','餐具隔离','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,152,'110100152','床旁隔离','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,153,'110100153','接触隔离','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,154,'110100154','消化道隔离','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'H',2,155,'110100155','新生儿隔离','','',0,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',3,156,'110200001','青霉素皮试','次','',3,1,1,0,0,'1',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',3,157,'110200002','氨苄青霉素皮试','次','',3,1,1,0,0,'1',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',3,158,'110200003','阿莫西林钠皮试','次','',3,1,1,0,0,'1',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',3,159,'110200004','哌拉西林钠皮试','次','',3,1,1,0,0,'1',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',3,160,'110200005','特治星皮试','次','',3,1,1,0,0,'1',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',3,161,'110200006','联邦他唑仙皮试','次','',3,1,1,0,0,'1',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',3,162,'110200007','氯唑西林钠皮试','次','',3,1,1,0,0,'1',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',3,163,'110200008','头孢唑啉钠皮试','次','',3,1,1,0,0,'1',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',3,164,'110200009','头孢拉定皮试','次','',3,1,1,0,0,'1',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',3,165,'110200010','西力欣(头孢呋辛)皮试','次','',3,1,1,0,0,'1',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',3,166,'110200011','力复乐(头孢呋辛)皮试','次','',3,1,1,0,0,'1',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',3,167,'110200012','头孢哌酮皮试','次','',3,1,1,0,0,'1',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',3,168,'110200013','头孢曲松钠皮试','次','',3,1,1,0,0,'1',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',3,169,'110200014','头孢三嗪皮试','次','',3,1,1,0,0,'1',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',3,170,'110200015','复达欣皮试','次','',3,1,1,0,0,'1',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',3,171,'110200016','头孢他啶皮试','次','',3,1,1,0,0,'1',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',3,172,'110200017','凯福定(头孢他啶)皮试','次','',3,1,1,0,0,'1',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',3,173,'110200018','头孢噻肟钠皮试','次','',3,1,1,0,0,'1',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',3,174,'110200019','头孢吡肟皮试','次','',3,1,1,0,0,'1',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',3,175,'110200020','凯福龙皮试','次','',3,1,1,0,0,'1',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',3,176,'110200021','头孢替唑钠皮试','次','',3,1,1,0,0,'1',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',3,177,'110200022','头孢米诺钠皮试','次','',3,1,1,0,0,'1',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',3,178,'110200023','链霉素皮试','次','',3,1,1,0,0,'1',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',3,179,'110200024','氨苄西林钠/舒巴坦皮试','次','',3,1,1,0,0,'1',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',3,180,'110200025','替卡西林/克拉维酸皮试','次','',3,1,1,0,0,'1',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',3,181,'110200026','阿莫西林/克拉维酸皮试','次','',3,1,1,0,0,'1',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',3,182,'110200027','亚胺培南西司他丁钠皮试','次','',3,1,1,0,0,'1',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',3,183,'110200028','美平(美罗培南)皮试','次','',3,1,1,0,0,'1',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',3,184,'110200029','铃兰欣皮试','次','',3,1,1,0,0,'1',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',3,185,'110200030','阿莫西林钠克拉维酸皮试','次','',3,1,1,0,0,'1',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',3,186,'110200031','美洛西林钠/舒巴坦皮试','次','',3,1,1,0,0,'1',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',3,187,'110200032','门冬酰胺酶皮试','次','',3,1,1,0,0,'1',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',3,188,'110200033','普鲁卡因皮试','次','',3,1,1,0,0,'1',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',3,189,'110200034','复方泛影葡胺皮试','次','',3,1,1,0,0,'1',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',3,190,'110200035','碘化油皮试','次','',3,1,1,0,0,'1',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',3,191,'110200036','荧光素钠皮试','次','',3,1,1,0,0,'1',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',3,192,'110200037','碘必乐(碘帕醇)皮试','次','',3,1,1,0,0,'1',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',3,193,'110200038','碘曲伦(伊索显)皮试','次','',3,1,1,0,0,'1',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',3,194,'110200039','优维显300(碘普罗胺)皮试','次','',3,1,1,0,0,'1',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',3,195,'110200040','优维显370(碘普罗胺)皮试','次','',3,1,1,0,0,'1',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',3,196,'110200041','碘海醇皮试','次','',3,1,1,0,0,'1',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',3,197,'110200042','欧乃影(扎双胺)皮试','次','',3,1,1,0,0,'1',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',3,198,'110200043','马根维显皮试','次','',3,1,1,0,0,'1',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',3,199,'110200044','破伤风抗毒素(TAT)皮试','次','',3,1,1,0,0,'1',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',3,200,'110200045','新瑞普欣皮试','次','',3,1,1,0,0,'1',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,201,'110300001','皮下注射','次','',3,0,0,0,0,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,202,'110300002','肌肉注射','次','',3,0,0,0,0,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,203,'110300003','静脉注射','次','',3,0,0,0,0,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,204,'110300004','静脉滴入','次','',3,0,0,0,0,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,205,'110300005','入液静滴','次','',3,0,0,0,1,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,206,'110300006','静脉泵内滴入','次','',3,0,0,0,1,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,207,'110300007','动静脉泵内注入','次','',3,0,0,0,1,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,208,'110300008','经泵动静脉注射','次','',3,0,0,0,1,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,209,'110300009','泵内注入','次','',3,0,0,0,1,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,210,'110300010','局部注射','次','',3,0,0,0,1,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,211,'110300011','腹腔注射','次','',3,0,0,0,1,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,212,'110300012','腹动脉插管注射','次','',3,0,0,0,1,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,213,'110300013','腹腔穿刺及注射','次','',3,0,0,0,1,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,214,'110300014','腹部皮下注射','次','',3,0,0,0,1,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,215,'110300015','肛门注入','次','',3,0,0,0,1,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,216,'110300016','动脉穿刺及注射术','次','',3,0,0,0,1,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,217,'110300017','股动脉注射','次','',3,0,0,0,1,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,218,'110300018','胸内注射','次','',3,0,0,0,0,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,219,'110300019','心内注射','次','',3,0,0,0,1,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,220,'110300020','鞘内注射','次','',3,0,0,0,1,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual;
Insert Into 诊疗项目目录(类别,分类id,id,编码,名称,计算单位,标本部位,计算方式,单独应用,执行频率,适用性别,组合项目,操作类型,执行安排,执行科室,服务对象,计价性质,建档时间,撤档时间)
Select 'E',4,221,'110300021','瘤周注射','次','',3,0,0,0,1,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,222,'110300022','骶管硬膜外注射','次','',3,0,0,0,1,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,223,'110300023','骶管注射','次','',3,0,0,0,1,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,224,'110300024','双颞浅注射','次','',3,0,0,0,1,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,225,'110300025','球后注射','次','',3,0,0,0,1,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,226,'110300026','眼球旁注射','次','',3,0,0,0,0,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,227,'110300027','眼结膜下注射','次','',3,0,0,0,0,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,228,'110300028','关节腔内注射','次','',3,0,0,0,0,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,229,'110300029','气管内注入','次','',3,0,0,0,1,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,230,'110300030','微量泵持续输入','次','',3,0,0,0,0,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,231,'110300031','动脉泵入','次','',3,0,0,0,0,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,232,'110300032','灌肠','次','',3,0,0,0,0,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,233,'110300033','喷入','次','',3,0,0,0,1,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,234,'110300034','喷雾吸入','次','',3,0,0,0,1,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,235,'110300035','雾化吸入','次','',3,0,0,0,1,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,236,'110300036','吸入','次','',3,0,0,0,1,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,237,'110300037','气管吸入','次','',3,0,0,0,0,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,238,'110300038','气管滴入','次','',3,0,0,0,0,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,239,'110300039','气管内滴入','次','',3,0,0,0,0,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,240,'110300040','U管滴入','次','',3,0,0,0,1,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,241,'110300041','滴斗入','次','',3,0,0,0,1,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,242,'110300042','腹腔滴入','次','',3,0,0,0,1,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,243,'110300043','鼻饲','次','',3,0,0,0,1,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,244,'110300044','置鼻饲管','次','',3,0,0,0,0,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,245,'110300045','应用鼻饲泵','次','',3,0,0,0,1,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,246,'110300046','滴眼','次','',3,0,0,0,1,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,247,'110300047','滴耳','次','',3,0,0,0,0,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,248,'110300048','滴鼻','次','',3,0,0,0,1,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,249,'110300049','滴内管','次','',3,0,0,0,1,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,250,'110300050','口服','次','',3,0,0,0,0,'2',0,0,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,251,'110300051','舌下含化','次','',3,0,0,0,1,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,252,'110300052','冲服','次','',3,0,0,0,1,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,253,'110300053','含化','次','',3,0,0,0,1,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,254,'110300054','含漱','次','',3,0,0,0,1,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,255,'110300055','漱口','次','',3,0,0,0,1,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,256,'110300056','纳肛','次','',3,0,0,0,1,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,257,'110300057','堵塞鼻腔','次','',3,0,0,0,1,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,258,'110300058','外用','次','',3,0,0,0,0,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,259,'110300059','局部贴敷','次','',3,0,0,0,0,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,260,'110300060','贴心前区','次','',3,0,0,0,1,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',4,261,'110300061','置阴道内','次','',3,0,0,0,1,'2',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,262,'110400001','备皮','次','',3,1,1,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,263,'110400002','全身皮肤备皮','次','',3,1,1,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,264,'110400003','剪短头发','次','',3,1,1,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,265,'110400004','剃头','次','',3,1,1,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,266,'110400005','剃刮胡须','次','',3,1,1,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,267,'110400006','刮剪耳毛','次','',3,1,1,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,268,'110400007','剪睫毛','次','',3,1,1,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,269,'110400008','剃剪眉毛','次','',3,1,1,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,270,'110400009','剪鼻毛','次','',3,1,1,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,271,'110400010','清洁腔道缝隙','次','',3,1,1,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,272,'110400011','局部清洗','次','',3,1,1,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,273,'110400012','组织缝合','次','',3,1,1,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,274,'110400013','血管缝合','次','',3,1,1,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,275,'110400014','筋腱缝合','次','',3,1,1,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,276,'110400015','创口缝合','次','',3,1,1,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,277,'110400016','创口包扎','次','',3,1,1,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,278,'110400017','拆线','次','',3,1,1,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,279,'110400018','腹腔灌注','次','',3,1,1,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,280,'110400019','经泵腹腔灌注','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,281,'110400020','胃管注入','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,282,'110400021','造瘘管注入','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,283,'110400022','胃造瘘管注入','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,284,'110400023','经空肠管注入','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,285,'110400024','伤口持续灌注','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,286,'110400025','关节腔持续灌注','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,287,'110400026','经营养管灌注','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,288,'110400027','保留灌肠','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,289,'110400028','冲洗口腔','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,290,'110400029','冲洗脓腔','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,291,'110400030','胆道引流管冲洗','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,292,'110400031','腹腔持续冲洗','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,293,'110400032','腹腔冲洗','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,294,'110400033','伤口持续滴入冲洗','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,295,'110400034','血液灌流','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,296,'110400035','胸腔冲洗','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,297,'110400036','造瘘口灌入','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,298,'110400037','直肠缓慢冲洗','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,299,'110400038','瘘管冲洗','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,300,'110400039','上颌窦穿刺冲洗','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,301,'110400040','U管冲洗','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,302,'110400041','鼻腔冲洗','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,303,'110400042','腹腔缓慢冲洗','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,304,'110400043','腹腔双套管冲洗','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,305,'110400044','回肠膀胱冲洗','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,306,'110400045','会阴部盆腔双套管持续冲洗','次','',3,1,1,2,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,307,'110400046','会阴部双套管冲洗接负压吸引器','次','',3,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,308,'110400047','会阴冲洗','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,309,'110400048','外阴冲洗','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,310,'110400049','会阴及阴道冲洗','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,311,'110400050','双上颌窦穿刺冲洗','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,312,'110400051','胃管冲洗','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,313,'110400052','冲洗胃管','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,314,'110400053','阴道灌洗','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,315,'110400054','阴道冲洗','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,316,'110400055','1：2：3液体灌肠','次','',3,1,1,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,317,'110400056','茶叶水保留灌肠','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,318,'110400057','肥皂水保留灌肠','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,319,'110400058','肥皂水灌肠','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,320,'110400059','假肛虹吸灌肠','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,321,'110400060','清洁灌肠','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,322,'110400061','温盐水灌肠','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,323,'110400062','胆囊空肠造口管灌入','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,324,'110400063','胃肠减压并冲洗','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,325,'110400064','双套管接吸引器持续负压吸引','次','',3,1,2,0,0,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,326,'110400065','引流管接吸引器持续负压吸引','次','',3,1,2,0,0,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,327,'110400066','双套管持续冲洗','次','',3,1,1,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,328,'110400067','双套管冲洗','次','',3,1,1,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,329,'110400068','引流管冲洗','次','',3,1,1,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,330,'110400069','经口鼻吸痰','次','',3,1,1,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual;
Insert Into 诊疗项目目录(类别,分类id,id,编码,名称,计算单位,标本部位,计算方式,单独应用,执行频率,适用性别,组合项目,操作类型,执行安排,执行科室,服务对象,计价性质,建档时间,撤档时间)
Select 'E',5,331,'110400070','经气管套管吸痰','次','',3,1,1,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,332,'110400071','吸痰','次','',3,1,2,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,333,'110400072','协助排痰','次','',3,1,2,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,334,'110400073','假肛接假肛袋','次','',3,1,1,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,335,'110400074','输尿管支架管接尿袋','次','',3,1,1,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,336,'110400075','输尿管导管接无菌尿袋','次','',3,1,1,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,337,'110400076','留置引流管','次','',3,1,1,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,338,'110400077','留置引流管接无菌袋','次','',3,1,1,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,339,'110400078','留置橡皮管引流','次','',3,1,1,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,340,'110400079','留置橡皮引流条','次','',3,1,1,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,341,'110400080','留置烟卷引流条','次','',3,1,1,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,342,'110400081','留置胃管','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,343,'110400082','留置肛管','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,344,'110400083','留置尿管','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,345,'110400084','留置尿管接无菌尿袋','次','',3,1,1,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,346,'110400085','留置气囊尿管','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,347,'110400086','留置气囊尿管接无菌尿袋','次','',3,1,1,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,348,'110400087','留置三腔尿管','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,349,'110400088','留置三腔气囊尿管','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,350,'110400089','留置三腔气囊尿管接床旁引流瓶','次','',3,1,1,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,351,'110400090','留置尿道支架管','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,352,'110400091','尿道留置猪尾巴支架管及输尿管导管','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,353,'110400092','留置T型引流管','次','',3,1,1,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,354,'110400093','引流管接无菌引流瓶','次','',3,1,1,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,355,'110400094','双侧输尿管引流管接无菌袋','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,356,'110400095','“T”管引流接无菌袋','次','',3,1,1,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,357,'110400096','“U”管引流接无菌袋','次','',3,1,1,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,358,'110400097','引流','次','',3,1,1,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,359,'110400098','肛管接床旁引流袋','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,360,'110400099','假肛接引流袋','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,361,'110400100','潘氏引流管接无菌袋','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,362,'110400101','乳胶引流管接无菌袋','次','',3,1,1,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,363,'110400102','闭式引流','次','',3,1,1,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,364,'110400103','导尿','次','',3,1,1,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,365,'110400104','漂浮导管插管术','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,366,'110400105','术前置胃管、尿管','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,367,'110400106','造瘘口换药','次','',3,1,1,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,368,'110400107','大疱抽液术','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,369,'110400108','间断胃肠减压','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,370,'110400109','胃肠减压','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,371,'110400110','胃肠减压管夹闭','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,372,'110400111','插胃管','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,373,'110400112','夹闭胃管','次','',3,1,1,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,374,'110400113','术前下胃管','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,375,'110400114','术前插胃管','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,376,'110400115','术前插胃管、尿管','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,377,'110400116','开放尿管','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,378,'110400117','下尿管','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,379,'110400118','冲管','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,380,'110400119','备无菌肛管','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,381,'110400120','插肛管','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,382,'110400121','肛管排气','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,383,'110400122','乳胶引流管长期开放','次','',3,1,1,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,384,'110400123','冰盐水洗胃','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,385,'110400124','冷盐水洗胃','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,386,'110400125','温盐水清洁洗肠(自造口远端及肛门)','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,387,'110400126','温盐水清洁洗肠(自造口远近端)','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,388,'110400127','温盐水洗胃','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,389,'110400128','洗胃','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,390,'110400129','气囊放气','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,391,'110400130','清洁胃肠','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,392,'110400131','吸氧','次','',2,1,2,0,0,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,393,'110400132','间断低流量吸氧','次','',2,1,2,0,0,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,394,'110400133','间断中流量吸氧','次','',2,1,2,0,0,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,395,'110400134','低流量吸氧','次','',2,1,2,0,0,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,396,'110400135','高流量吸氧','次','',2,1,2,0,0,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,397,'110400136','面罩吸氧','次','',2,1,2,0,0,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,398,'110400137','呼吸机用氧','次','',2,1,2,0,0,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,399,'110400138','呼吸机持续吸氧','次','',2,1,2,0,0,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,400,'110400139','呼吸机面罩给氧','次','',2,1,2,0,0,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,401,'110400140','呼吸机面罩加压给氧','次','',2,1,2,0,0,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,402,'110400141','物理降温','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,403,'110400142','冰袋降温','次','',3,1,0,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,404,'110400143','酒精擦浴','次','',3,1,0,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,405,'110400144','酒精擦浴,冰袋降温交替','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,406,'110400145','温水坐浴','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,407,'110400146','静脉留置针','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,408,'110400147','锁骨下静脉封管','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,409,'110400148','大静脉穿刺插管术','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,410,'110400149','锁骨下静脉穿刺插管术','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,411,'110400150','颈内静脉穿刺插管术','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,412,'110400151','右锁骨下静脉穿刺插管','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,413,'110400152','大静脉射管处换药','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,414,'110400153','大静脉插管处换药','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,415,'110400154','大静脉换药','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,416,'110400155','中心静脉压测定','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,417,'110400156','测中心静脉压','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,418,'110400157','周围静脉压测定','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,419,'110400158','肘静脉压测定','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,420,'110400159','局部砂袋压迫止血','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,421,'110400160','切口处沙袋压迫','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,422,'110400161','腰部砂袋压迫','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,423,'110400162','右腹股沟穿刺处沙袋压迫','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,424,'110400163','穿刺点加压包扎','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,425,'110400164','右下腹砂袋压迫','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,426,'110400165','股动脉穿刺处沙袋压迫','次','',3,1,1,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,427,'110400166','腹腔穿刺放腹水','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,428,'110400167','腹腔穿刺术','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,429,'110400168','肝脏穿刺活检','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,430,'110400169','肝脏穿刺术','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,431,'110400170','关节腔穿刺术','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,432,'110400171','滑囊穿刺术','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,433,'110400172','局麻下行B超引导下肝脏穿刺活检术','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,434,'110400173','局麻下行前列腺穿刺术','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,435,'110400174','上颌窦穿刺术','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,436,'110400175','头皮血肿穿刺术','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,437,'110400176','心包穿刺术','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,438,'110400177','心包穿刺插管术','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,439,'110400178','阴道后穹窿穿刺术','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,440,'110400179','诊断性腹腔穿刺术','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual;
Insert Into 诊疗项目目录(类别,分类id,id,编码,名称,计算单位,标本部位,计算方式,单独应用,执行频率,适用性别,组合项目,操作类型,执行安排,执行科室,服务对象,计价性质,建档时间,撤档时间)
Select 'E',5,441,'110400180','腕关节穿刺术','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,442,'110400181','关节穿刺术','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,443,'110400182','髌上囊穿刺','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,444,'110400183','腹水回输','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,445,'110400184','骨髓回输','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,446,'110400185','外周血造血干细胞回输','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,447,'110400186','关节镜下活检术','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,448,'110400187','右骶髂部活检术','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,449,'110400188','肾脏活检术','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,450,'110400189','小换药','次','',3,1,1,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,451,'110400190','中换药','次','',3,1,1,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,452,'110400191','大换药','次','',3,1,1,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,453,'110400192','口腔换药','次','',3,1,1,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,454,'110400193','鼻腔换药','次','',3,1,1,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,455,'110400194','颈部伤口大换药','次','',3,1,1,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,456,'110400195','会阴侧切伤口换药','次','',3,1,1,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,457,'110400196','会阴伤口换药','次','',3,1,1,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,458,'110400197','“T”管周围大换药','次','',3,1,1,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,459,'110400198','呼吸机辅助呼吸','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,460,'110400199','呼吸机湿化用水','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,461,'110400200','呼吸机湿化用无菌水','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,462,'110400201','无创血压监测','次','',3,1,0,0,0,'0',0,0,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,463,'110400202','血氧饱和度监测','次','',3,1,0,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,464,'110400203','临时起搏器监测','次','',3,1,0,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,465,'110400204','胎心监护','次','',2,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,466,'110400205','心电、血压监测','次','',3,1,0,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,467,'110400206','心电监测','次','',3,1,0,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,468,'110400207','心电监护','次','',2,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,469,'110400208','心排出量监测','次','',3,1,0,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,470,'110400209','氧饱和度监测','次','',3,1,0,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,471,'110400210','氧分压监测','次','',3,1,0,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,472,'110400211','血压监护','次','',2,1,2,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,473,'110400212','血液动力学监测','次','',2,1,0,0,0,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,474,'110400213','电除颤','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,475,'110400214','气管镜下经鼻气管插管','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,476,'110400215','局麻下经鼻气管插管术','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,477,'110400216','气管插管术','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,478,'110400217','清洗气管内套管','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,479,'110400218','堵塞气管套管','次','',3,1,1,0,1,'0',0,2,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',5,480,'110400219','煮沸内套管','次','',3,1,1,0,1,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,481,'110500001','半量半流食','','',0,1,2,0,0,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,482,'110500002','半流食','','',0,1,2,0,0,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,483,'110500003','半流食(免鱼虾蛋奶)','','',0,1,2,0,0,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,484,'110500004','鼻饲混合奶','','',0,1,2,0,0,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,485,'110500005','鼻饲无糖混合奶','','',0,1,2,0,0,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,486,'110500006','鼻饲饮食','','',0,1,2,0,0,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,487,'110500007','鼻饲饮食总热量','','',0,1,2,0,0,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,488,'110500008','扁桃体术后饮食','','',0,1,2,0,0,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,489,'110500009','产后饮食','','',0,1,2,0,0,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,490,'110500010','纯素口腔半流食','','',0,1,2,0,0,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,491,'110500011','低蛋白低嘌呤普食','','',0,1,2,0,0,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,492,'110500012','低嘌呤低蛋白普食','','',0,1,2,0,0,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,493,'110500013','低蛋白普食','','',0,1,2,0,0,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,494,'110500014','低蛋白软饭','','',0,1,2,0,0,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,495,'110500015','低蛋白糖尿病普食','','',0,1,2,0,0,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,496,'110500016','低钠普食','','',0,1,2,0,0,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,497,'110500017','低铜普食','','',0,1,2,0,0,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,498,'110500018','低盐半流食','','',0,1,2,0,0,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,499,'110500019','低盐低蛋白半流食','','',0,1,2,0,0,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,500,'110500020','低盐低蛋白低脂普食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,501,'110500021','低盐低蛋白低嘌呤普食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,502,'110500022','低盐低蛋白普食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,503,'110500023','低盐低蛋白软饭','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,504,'110500024','低盐低蛋白糖尿病普食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,505,'110500025','低盐低脂低蛋白糖尿病普食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,506,'110500026','低盐低脂普食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,507,'110500027','低盐低脂糖尿病普食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,508,'110500028','低盐高蛋白半流食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,509,'110500029','低盐高蛋白普食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,510,'110500030','低盐高蛋白软饭','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,511,'110500031','低盐普食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,512,'110500032','低盐软饭','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,513,'110500033','低盐糖尿病普食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,514,'110500034','低盐优质蛋白糖尿病普食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,515,'110500035','低盐优质高蛋白普食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,516,'110500036','低脂半流食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,517,'110500037','少油半流食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,518,'110500038','低脂高蛋白半流食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,519,'110500039','高蛋白低脂半流食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,520,'110500040','高蛋白低脂肪半流食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,521,'110500041','低脂高蛋白普食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,522,'110500042','高蛋白低脂肪普食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,523,'110500043','低脂口腔半流食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,524,'110500044','低脂流食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,525,'110500045','低脂普食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,526,'110500046','低脂软饭','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,527,'110500047','低脂少渣糖尿病软饭','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,528,'110500048','低脂糖尿病半流食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,529,'110500049','低脂糖尿病普食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,530,'110500050','低嘌呤半流食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,531,'110500051','低嘌呤低蛋白糖尿病普食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,532,'110500052','低嘌呤普食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,533,'110500053','低嘌呤糖尿病普食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,534,'110500054','高蛋白半流食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,535,'110500055','高蛋白流食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,536,'110500056','高蛋白普食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,537,'110500057','高蛋白软饭','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,538,'110500058','高蛋白糖尿病半流食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,539,'110500059','高蛋白糖尿病普食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,540,'110500060','高铁软食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,541,'110500061','高脂餐','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,542,'110500062','固定饮食、每天钾60mEq、钠120mEq','','',0,1,2,0,0,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,543,'110500063','固定饮食、每天钠120mEq','','',0,1,2,0,0,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,544,'110500064','固定饮食、每天钠160mEq、钾60mEq','','',0,1,2,0,0,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,545,'110500065','固定饮食、每天主食400g、钠160mEq','','',0,1,2,0,0,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,546,'110500066','管饲流食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,547,'110500067','管饲饮食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,548,'110500068','禁食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,549,'110500069','禁食、水','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,550,'110500070','口腔半流食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual;
Insert Into 诊疗项目目录(类别,分类id,id,编码,名称,计算单位,标本部位,计算方式,单独应用,执行频率,适用性别,组合项目,操作类型,执行安排,执行科室,服务对象,计价性质,建档时间,撤档时间)
Select 'I',6,551,'110500071','溃疡病高铁软饭','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,552,'110500072','流食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,553,'110500073','流质','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,554,'110500074','全流饮食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,555,'110500075','米汤','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,556,'110500076','免奶匀浆饮食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,557,'110500077','母乳喂养','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,558,'110500078','牛奶','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,559,'110500079','普食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,560,'110500080','普食(免鱼虾蛋奶)','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,561,'110500081','清半流食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,562,'110500082','清流食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,563,'110500083','全糖流食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,564,'110500084','软食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,565,'110500085','少碘普食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,566,'110500086','少碘糖尿病普食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,567,'110500087','少油糖尿病饮食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,568,'110500088','少渣半流食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,569,'110500089','少渣普食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,570,'110500090','少渣软食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,571,'110500091','术晨禁食、水','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,572,'110500092','术前禁食、水','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,573,'110500093','素半流食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,574,'110500094','素口腔半流食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,575,'110500095','素普食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,576,'110500096','糖尿病半流','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,577,'110500097','糖尿病半流食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,578,'110500098','糖尿病口腔半流食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,579,'110500099','糖尿病零号饭','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,580,'110500100','糖尿病流食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,581,'110500101','糖尿病普食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,582,'110500102','糖尿病隐血试验普食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,583,'110500103','胃切Ⅰ号饮食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,584,'110500104','无菌高蛋白流食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,585,'110500105','无糖混合奶','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,586,'110500106','无糖混合奶Ⅱ号','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,587,'110500107','无渣半流食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,588,'110500108','无渣普食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,589,'110500109','无脂餐','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,590,'110500110','小儿半流食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,591,'110500111','隐血试验饮食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,592,'110500112','婴儿辅助食品','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,593,'110500113','婴儿奶','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,594,'110500114','硬化口腔半流食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,595,'110500115','优质低蛋白普食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'I',6,596,'110500116','匀浆饮食','','',0,1,2,0,1,'',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'Z',7,597,'110600001','留院观察','','',3,1,1,0,0,'1',0,1,1,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'Z',7,598,'110600002','转院','','',3,1,1,0,0,'6',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'Z',7,599,'110600003','住院观察','','',3,1,1,0,0,'2',0,1,1,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'Z',7,600,'110600004','转科治疗','','',3,1,1,0,0,'3',0,1,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'Z',7,601,'110600005','病危通知','','',0,1,1,0,0,'0',0,0,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'Z',7,602,'110600006','术后','','',0,1,2,0,0,'4',0,1,2,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'Z',7,603,'110600007','他科会诊','','',3,1,1,0,0,'7',0,1,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'Z',7,604,'110600008','院外会诊','','',0,1,1,0,0,'0',0,1,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'Z',7,605,'110600009','陪伴','','',0,1,2,0,0,'0',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'Z',7,606,'110600010','出院','','',3,1,1,0,0,'5',0,1,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'Z',7,607,'110600011','房间紫外线消毒','','',0,1,2,0,0,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'Z',7,608,'110600012','风淋消毒','','',0,1,2,0,0,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'Z',7,609,'110600013','母婴分室','','',0,1,2,2,0,'0',0,0,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'Z',7,610,'110600014','母婴同室','','',0,1,2,2,0,'0',0,0,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'Z',7,611,'110600015','保持室温30℃','','',0,1,2,0,0,'0',0,0,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'Z',7,612,'110600016','住CCU单间','','',0,1,2,0,0,'0',0,0,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'Z',7,613,'110600017','住层流室','','',0,1,2,0,0,'0',0,0,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'Z',7,614,'110600018','住高危新生儿室','','',0,1,2,0,0,'0',0,0,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'Z',7,615,'110600019','住高危婴儿室','','',0,1,2,0,0,'0',0,0,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'Z',7,616,'110600020','住呼吸监护室','','',0,1,2,0,0,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'Z',7,617,'110600021','住恢复间','','',0,1,2,0,0,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'Z',7,618,'110600022','住急救室','','',0,1,2,0,0,'0',0,0,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'Z',7,619,'110600023','住洁净室','','',0,1,2,0,0,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'Z',7,620,'110600024','患肢保温','','',0,1,2,0,0,'0',0,1,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'Z',7,621,'110600025','住监护室','','',0,1,2,0,0,'0',0,2,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',9,622,'210100001','头颅定位测量平片摄影','次','',3,1,1,0,0,'X线',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',9,623,'210100002','头颅定位测量平片摄影','次','副鼻窦瓦氏位',3,1,1,0,0,'X线',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',9,624,'210100003','头颅定位测量平片摄影','次','乳突劳氏位',3,1,1,0,0,'X线',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',9,625,'210100004','头颅定位测量平片摄影','次','乳突伦氏位',3,1,1,0,0,'X线',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',9,626,'210100005','头颅定位测量平片摄影','次','乳突梅氏位',3,1,1,0,0,'X线',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',9,627,'210100006','头颅定位测量平片摄影','次','乳突斯氏位',3,1,1,0,0,'X线',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',9,628,'210100007','头颅定位测量平片摄影','次','视神经孔轴位(瑞氏位)',3,1,1,0,0,'X线',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',9,629,'210100008','头颅定位测量平片摄影','次','颅底颌顶位(下上轴位)',3,1,1,0,0,'X线',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',9,630,'210100009','头颅定位测量平片摄影','次','内听道正位',3,1,1,0,0,'X线',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',9,631,'210100010','胸部体层摄影','次','',3,1,1,0,0,'X线',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',9,632,'210100011','胸部体层摄影','次','胸部正位',3,1,1,0,0,'X线',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',9,633,'210100012','胸部体层摄影','次','胸部侧位',3,1,1,0,0,'X线',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',9,634,'210100013','气管支气管正位体层摄影','次','气管支气管',3,1,1,0,0,'X线',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',9,635,'210100014','支气管双倾斜位体层摄影','次','支气管',3,1,1,0,0,'X线',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',9,636,'210100015','喉部体层摄影','次','喉部',3,1,1,0,0,'X线',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',9,637,'210100016','蝶鞍体层摄影','次','蝶鞍',3,1,1,0,0,'X线',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',9,638,'210100017','上颌窦正位体层摄影','次','上颌窦',3,1,1,0,0,'X线',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',9,639,'210100018','颞颌关节体层摄影','次','颞颌关节',3,1,1,0,0,'X线',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',9,640,'210100019','数字化摄影','次','',3,1,1,0,0,'X线',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',9,641,'210100020','计算机化X线摄影','次','',3,1,1,0,0,'X线',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',10,642,'210200001','胸部透视检查','次','胸部',3,1,1,0,0,'X线',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',10,643,'210200002','支气管造影','次','支气管',3,1,1,0,0,'X线',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',11,644,'210300001','右心造影','次','右心',3,1,1,0,0,'X线',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',11,645,'210300002','左心房造影','次','左心房',3,1,1,0,0,'X线',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',11,646,'210300003','左心室造影','次','左心室',3,1,1,0,0,'X线',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',11,647,'210300004','主动脉造影','次','主动脉',3,1,1,0,0,'X线',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',11,648,'210300005','冠状动脉造影','次','冠状动脉',3,1,1,0,0,'X线',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',11,649,'210300006','选择性肺动脉造影','次','肺动脉',3,1,1,0,0,'X线',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',11,650,'210300007','右心导管检查','次','右心',3,1,1,0,0,'X线',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',12,651,'210400001','食道造影','次','',3,1,1,0,0,'X线',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',12,652,'210400002','胃、十二指肠造影','次','',3,1,1,0,0,'X线',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',12,653,'210400003','低张十二指肠造影','次','',3,1,1,0,0,'X线',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',12,654,'210400004','小肠造影','次','',3,1,1,0,0,'X线',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',12,655,'210400005','肠套叠空气灌肠的诊断和复位','次','',3,1,1,0,0,'X线',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',12,656,'210400006','结肠造影','次','',3,1,1,0,0,'X线',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',12,657,'210400007','内窥镜逆行胰胆管造影','次','',3,1,1,0,0,'X线',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',12,658,'210400008','口服法胆囊造影','次','',3,1,1,0,0,'X线',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',12,659,'210400009','静脉法胆囊造影','次','',3,1,1,0,0,'X线',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',12,660,'210400010','经皮肝穿刺胆管造影','次','',3,1,1,0,0,'X线',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual;
Insert Into 诊疗项目目录(类别,分类id,id,编码,名称,计算单位,标本部位,计算方式,单独应用,执行频率,适用性别,组合项目,操作类型,执行安排,执行科室,服务对象,计价性质,建档时间,撤档时间)
Select 'D',12,661,'210400011','术中胆道造影','次','',3,1,1,0,0,'X线',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',12,662,'210400012','经T形管胆道造影','次','',3,1,1,0,0,'X线',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',13,663,'210500001','静脉尿路造影','次','',3,1,1,0,0,'X线',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',13,664,'210500002','逆行肾盂造影','次','',3,1,1,0,0,'X线',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',13,665,'210500003','膀胱造影','次','',3,1,1,0,0,'X线',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',13,666,'210500004','男性尿道造影','次','',3,1,1,0,0,'X线',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',13,667,'210500005','子宫输卵管造影','次','',3,1,1,0,0,'X线',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',13,668,'210500006','盆腔充气造影','次','',3,1,1,0,0,'X线',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',14,669,'210600001','乳腺钼靶(或钼铑靶)Ｘ线摄影','次','',3,1,1,0,0,'X线',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',14,670,'210600002','乳腺导管造影','次','',3,1,1,0,0,'X线',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',14,671,'210600003','乳腺囊肿内充气造影','次','',3,1,1,0,0,'X线',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',14,672,'210600004','腮腺造影','次','',3,1,1,0,0,'X线',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',14,673,'210600005','泪道造影','次','',3,1,1,0,0,'X线',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',14,674,'210600006','脊髓造影','次','',3,1,1,0,0,'X线',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',14,675,'210600007','膝关节造影','次','',3,1,1,0,0,'X线',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',14,676,'210600008','窦道及瘘管造影','次','',3,1,1,0,0,'X线',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',14,677,'210600009','眼球异物定位(角膜缘环定位法)','次','',3,1,1,0,0,'X线',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',16,678,'220100001','选择性血管插管技术','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',16,679,'220100002','动脉内化疗术','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',16,680,'220100003','动脉内其他药物灌注术','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',16,681,'220100004','经导管动脉拴塞术','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',16,682,'220100005','经皮腔内血管成形术','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',16,683,'220100006','经皮左锁骨下动脉导管药盒系统植入术','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',16,684,'220100007','经皮肝门静脉导管药盒系统植入术','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',16,685,'220100008','经皮股动脉导管药盒系统植入术','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',16,686,'220100009','经颈静脉肝内门体支架分流术','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',16,687,'220100010','经皮肝胆道内外引流术','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',16,688,'220100011','经皮肾盂内外引流术','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',16,689,'220100012','脓肿、囊肿引流术','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',16,690,'220100013','胆管内支架(内涵管)置入术','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',16,691,'220100014','食管内支架置入术','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',16,692,'220100015','血管内支架置入术','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',16,693,'220100016','经皮活检术','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',16,694,'220100017','肿瘤内药物注射术','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',16,695,'220100018','经皮腹腔神经丛阻滞术','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',17,696,'220200001','椎动脉造影','次','',3,1,1,0,0,'其他',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',17,697,'220200002','颈动脉造影','次','',3,1,1,0,0,'其他',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',17,698,'220200003','脑膜瘤的介入治疗','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',17,699,'220200004','鼻咽部血管纤维瘤的介入治疗','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',17,700,'220200005','副神经节肿瘤的介入治疗','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',17,701,'220200006','脊髓血管病变的栓塞治疗','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',17,702,'220200007','经皮肺动脉瓣球囊成形术','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',17,703,'220200008','经皮二尖瓣球囊成形术','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',17,704,'220200009','经皮动脉导管未闭封堵术','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',17,705,'220200010','经导管房间隔缺损封闭术','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',17,706,'220200011','经皮腔内血管成形术','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',17,707,'220200012','右心导管检查','次','',3,1,1,0,0,'其他',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',17,708,'220200013','选择性冠状动脉造影','次','',3,1,1,0,0,'其他',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',17,709,'220200014','经皮穿刺冠状动脉腔内成形术','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',17,710,'220200015','主动脉成形术','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',17,711,'220200016','下腔静脉成形术','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',17,712,'220200017','心血管异物摘取','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',17,713,'220200018','支气管动脉造影','次','',3,1,1,0,0,'其他',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',17,714,'220200019','经皮肺脓肿穿刺引流术','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',17,715,'220200020','肺癌的化学栓塞治疗','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',17,716,'220200021','咯血的介入治疗','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',17,717,'220200022','肺动-静脉瘘的介入治疗','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',17,718,'220200023','肺栓塞的介入治疗','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',17,719,'220200024','肺部疾病经皮活检','次','',3,1,1,0,0,'其他',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',17,720,'220200025','纵隔活检','次','',3,1,1,0,0,'其他',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',17,721,'220200026','消化道出血的介入治疗','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',17,722,'220200027','胃冠状静脉栓塞治疗','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',17,723,'220200028','胃肠道狭窄扩张术','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',17,724,'220200029','食管支架置放术','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',17,725,'220200030','肝脏血管造影','次','',3,1,1,0,0,'其他',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',17,726,'220200031','肝硬化门静脉高压症的介入治疗','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',17,727,'220200032','肝癌的化疗栓塞','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',17,728,'220200033','经皮乙醇注射(PEI)治疗或TACE后残余癌灶','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',17,729,'220200034','肝海绵状血管瘤的介入治疗','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',17,730,'220200035','肝脏创伤的介入治疗','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',17,731,'220200036','肝脏脓肿引流','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',17,732,'220200037','肝脏经皮穿刺活检','次','',3,1,1,0,0,'其他',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',17,733,'220200038','经皮肝穿刺胆管造影','次','',3,1,1,0,0,'其他',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',17,734,'220200039','内窥镜逆行胆胰管造影','次','',3,1,1,0,0,'其他',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',17,735,'220200040','经颈静脉穿刺胆管造影','次','',3,1,1,0,0,'其他',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',17,736,'220200041','胆管扩张术','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',17,737,'220200042','经T管窦道网篮取石','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',17,738,'220200043','经T管窦道取石钳取石','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',17,739,'220200044','经T管窦道推石入肠','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',17,740,'220200045','经T管窦道处理胆囊管残端或总胆管憩室内结石','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',17,741,'220200046','经皮经肝排石入肠','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',17,742,'220200047','胰腺血管造影','次','',3,1,1,0,0,'其他',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',17,743,'220200048','胰腺肿瘤经皮活检','次','',3,1,1,0,0,'其他',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',17,744,'220200049','胰腺囊肿引流','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',17,745,'220200050','胰岛细胞功能性肿瘤经皮采集血标本','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',17,746,'220200051','脾动脉造影','次','',3,1,1,0,0,'其他',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',17,747,'220200052','脾疾病的介入治疗','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',17,748,'220200053','肾动脉造影','次','',3,1,1,0,0,'其他',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',17,749,'220200054','肾肿瘤栓塞','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',17,750,'220200055','肾动脉成形术','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',17,751,'220200056','前列腺肥大尿道狭窄扩张术','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',17,752,'220200057','输尿管狭窄扩张术','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',17,753,'220200058','妇科治疗肿瘤灌注化疗与栓塞','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',17,754,'220200059','妇产科大出血的经导管栓塞治疗','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',17,755,'220200060','经同轴导管输卵管选择性造影及再通术','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',17,756,'220200061','经皮穿刺椎间盘切割术','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',17,757,'220200062','骨肿瘤的动脉内化疗','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',19,758,'230100001','颅脑CT扫描','次','',3,1,1,0,0,'CT',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',19,759,'230100002','垂体和鞍区CT扫描','次','',3,1,1,0,0,'CT',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',19,760,'230100003','后颅窝及桥小脑角区CT扫描','次','',3,1,1,0,0,'CT',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',19,761,'230100004','颞颌关节CT扫描','次','',3,1,1,0,0,'CT',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',19,762,'230100005','眼和眼眶CT扫描','次','',3,1,1,0,0,'CT',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',19,763,'230100006','耳和颞骨CT扫描','次','',3,1,1,0,0,'CT',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',19,764,'230100007','鼻和副鼻窦CT扫描','次','',3,1,1,0,0,'CT',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',19,765,'230100008','鼻咽部和咽旁间隙CT扫描','次','',3,1,1,0,0,'CT',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',19,766,'230100009','喉CT扫描','次','',3,1,1,0,0,'CT',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',19,767,'230100010','甲状腺CT扫描','次','',3,1,1,0,0,'CT',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',19,768,'230100011','肺部CT扫描','次','',3,1,1,0,0,'CT',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',19,769,'230100012','纵隔CT扫描','次','',3,1,1,0,0,'CT',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',19,770,'230100013','肝、脾CT扫描','次','',3,1,1,0,0,'CT',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual;
Insert Into 诊疗项目目录(类别,分类id,id,编码,名称,计算单位,标本部位,计算方式,单独应用,执行频率,适用性别,组合项目,操作类型,执行安排,执行科室,服务对象,计价性质,建档时间,撤档时间)
Select 'D',19,771,'230100014','胆道系统CT扫描','次','',3,1,1,0,0,'CT',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',19,772,'230100015','胰腺CT扫描','次','',3,1,1,0,0,'CT',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',19,773,'230100016','肾上腺CT扫描','次','',3,1,1,0,0,'CT',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',19,774,'230100017','腹膜后间隙CT扫描','次','',3,1,1,0,0,'CT',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',19,775,'230100018','胃肠道CT扫描','次','',3,1,1,0,0,'CT',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',19,776,'230100019','盆腔CT扫描','次','',3,1,1,0,0,'CT',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',19,777,'230100020','脊髓和脊柱CT扫描','次','',3,1,1,0,0,'CT',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',19,778,'230100021','四肢及软组织CT扫描','次','',3,1,1,0,0,'CT',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',20,779,'230200001','垂体螺旋CT扫描','次','',3,1,1,0,0,'CT',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',20,780,'230200002','肺部螺旋CT扫描','次','',3,1,1,0,0,'CT',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',20,781,'230200003','肝脏螺旋CT扫描','次','',3,1,1,0,0,'CT',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',20,782,'230200004','胰腺螺旋CT扫描','次','',3,1,1,0,0,'CT',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',20,783,'230200005','肾脏和肾上腺螺旋CT扫描','次','',3,1,1,0,0,'CT',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',20,784,'230200006','高分辨率CT(HRCT)','次','',3,1,1,0,0,'CT',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',21,785,'230300001','单层容积扫描','次','',3,1,1,0,0,'CT',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',21,786,'230300002','多层容积扫描','次','',3,1,1,0,0,'CT',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',21,787,'230300003','连续容积扫描','次','',3,1,1,0,0,'CT',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',22,788,'240000001','颅脑MRI检查','次','',3,1,1,0,0,'MRI',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',22,789,'240000002','颅内MRA检查','次','',3,1,1,0,0,'MRI',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',22,790,'240000003','眼部MRI检查','次','',3,1,1,0,0,'MRI',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',22,791,'240000004','鼻及鼻窦MRI检查','次','',3,1,1,0,0,'MRI',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',22,792,'240000005','颞颌关节MRI检查','次','',3,1,1,0,0,'MRI',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',22,793,'240000006','耳部MRI检查','次','',3,1,1,0,0,'MRI',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',22,794,'240000007','鼻咽部MRI检查','次','',3,1,1,0,0,'MRI',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',22,795,'240000008','口咽部MRI检查','次','',3,1,1,0,0,'MRI',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',22,796,'240000009','喉及甲状腺MRI检查','次','',3,1,1,0,0,'MRI',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',22,797,'240000010','颅颈部MRA检查','次','',3,1,1,0,0,'MRI',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',22,798,'240000011','纵隔、肺、胸膜MRI检查','次','',3,1,1,0,0,'MRI',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',22,799,'240000012','心脏、大血管MRI检查','次','',3,1,1,0,0,'MRI',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',22,800,'240000013','乳房MRI检查','次','',3,1,1,0,0,'MRI',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',22,801,'240000014','肝胆脾MRI检查','次','',3,1,1,0,0,'MRI',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',22,802,'240000015','胰腺MRI检查','次','',3,1,1,0,0,'MRI',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',22,803,'240000016','肾及肾上腺MRI技术','次','',3,1,1,0,0,'MRI',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',22,804,'240000017','腹部血管MRA检查','次','',3,1,1,0,0,'MRI',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',22,805,'240000018','MRI胆道造影技术','次','',3,1,1,0,0,'MRI',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',22,806,'240000019','男性盆腔MRI检查','次','',3,1,1,0,0,'MRI',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',22,807,'240000020','女性盆腔和生殖器官MRI检查','次','',3,1,1,0,0,'MRI',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',22,808,'240000021','产科MRI检查','次','',3,1,1,0,0,'MRI',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',22,809,'240000022','脊柱、脊髓MRI检查','次','',3,1,1,0,0,'MRI',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',22,810,'240000023','骨髓MRI检查','次','',3,1,1,0,0,'MRI',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',22,811,'240000024','髋关节MRI检查','次','',3,1,1,0,0,'MRI',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',22,812,'240000025','膝关节MRI检查','次','',3,1,1,0,0,'MRI',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',22,813,'240000026','肩关节MRI检查','次','',3,1,1,0,0,'MRI',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',22,814,'240000027','上臂MRII检查','次','',3,1,1,0,0,'MRI',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',22,815,'240000028','小腿MRII检查','次','',3,1,1,0,0,'MRI',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',22,816,'240000029','MRI介入诊疗','次','',3,1,1,0,0,'MRI',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',25,817,'250110001','颅脑B超','次','',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',25,818,'250110002','颅脑彩超','次','',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',25,819,'250110003','眼部B超','次','双眼及附属器',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',25,820,'250110004','眼部彩超','次','双眼及附属器',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',25,821,'250110005','颌面部B超','次','',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',25,822,'250110006','颌面部彩超','次','',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',25,823,'250110007','颈部B超','次','',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',25,824,'250110008','颈部彩超','次','',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',25,825,'250110009','心脏B超检查','次','',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',25,826,'250110010','心脏彩超检查','次','',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',25,827,'250110011','超声心动图负荷试验','次','',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',25,828,'250110012','胸部常规B超检查','次','肺、胸腔、纵隔',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',25,829,'250110013','胸部常规彩超检查','次','肺、胸腔、纵隔',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',25,830,'250110014','腹部常规B超检查','次','肝、胆、胰、脾、双肾',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',25,831,'250110015','腹部常规彩超检查','次','肝、胆、胰、脾、双肾',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',25,832,'250110016','胃肠道B超检查','次','',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',25,833,'250110017','胃肠道彩超检查','次','',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',25,834,'250110018','泌尿系B超检查','次','双肾、输尿管、膀胱、前列腺',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',25,835,'250110019','泌尿系彩超检查','次','双肾、输尿管、膀胱、前列腺',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',25,836,'250110020','妇科B超检查','次','子宫、附件、膀胱及周围组织',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',25,837,'250110021','妇科彩超检查','次','子宫、附件、膀胱及周围组织',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',25,838,'250110022','产科B超检查','次','子宫、附件、胎儿及宫腔',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',25,839,'250110023','腹部大血管B超','次','',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',25,840,'250110024','腹部大血管彩超','次','',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',25,841,'250110025','阴囊与睾丸B超','次','阴囊、双侧睾丸、附睾',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',25,842,'250110026','阴囊与睾丸彩超','次','阴囊、双侧睾丸、附睾',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',26,843,'250120001','肾上腺超声','次','',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',26,844,'250120002','后腹膜超声','次','',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',26,845,'250120003','体腔积液超声','次','',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',26,846,'250120004','伤口感染超声','次','',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',26,847,'250120005','骨关节超声','次','',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',27,848,'250130001','彩超脐动脉血流监测','次','',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',27,849,'250130002','颅内段血管彩超','次','',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',27,850,'250130003','球后全部血管彩超','次','',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',27,851,'250130004','双颈动、静脉、椎动脉系彩超','次','',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',27,852,'250130005','门静脉系统彩超','次','',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',27,853,'250130006','四肢血管彩超','次','',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',27,854,'250130007','双肾及肾血管彩超','次','',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',27,855,'250130008','左肾静脉胡桃夹综合征检查','次','',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',27,856,'250130009','直立试验','次','',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',27,857,'250130010','药物血管功能试验','次','',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',27,858,'250130011','脏器或肿瘤声学照影','次','',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',28,859,'250200001','经食管超声心动图','次','',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',28,860,'250200002','经直肠B超检查','次','',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',28,861,'250200003','经直肠彩超检查','次','',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',28,862,'250200004','经阴道B超检查','次','',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',28,863,'250200005','经阴道彩超检查','次','',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',28,864,'250200006','血管内彩超检查','次','',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',29,865,'250300001','超声引导诊断','次','',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',29,866,'250300002','超声引导治疗','次','',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',30,867,'250400001','经颅多普勒血流图(TCD)','次','',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',30,868,'250400002','下(上)肢多普勒血流图','次','',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',30,869,'250400003','多普勒胎心监测','次','',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',30,870,'250400004','多普勒小儿血压检测','次','',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',30,871,'250400005','脏器灰阶三维超声立体成象','次','',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',30,872,'250400006','能量图血流三维超声立体成象','次','',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',30,873,'250400007','红外热象检查','次','',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',30,874,'250400008','心脏计算机三维重建技术(3DE)','次','',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',30,875,'250400009','心脏声学定量(AQ)','次','',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',30,876,'250400010','心脏彩色室壁动力(CK)','次','',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',30,877,'250400011','心脏组织多普勒显象(TDI)','次','',3,1,1,0,0,'超声',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',32,878,'260100001','血常规(血涂片检查)','次','手指血',3,1,1,0,0,'临检',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',32,879,'260100002','血细胞分析(23项)','次','静脉抗凝血',3,1,1,0,0,'临检',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',32,880,'260100003','血细胞分析(30项)','次','静脉抗凝血',3,1,1,0,0,'临检',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual;
Insert Into 诊疗项目目录(类别,分类id,id,编码,名称,计算单位,标本部位,计算方式,单独应用,执行频率,适用性别,组合项目,操作类型,执行安排,执行科室,服务对象,计价性质,建档时间,撤档时间)
Select 'C',32,881,'260100004','尿常规','次','尿液',3,1,1,0,0,'临检',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',32,882,'260100005','尿沉渣镜检','次','晨尿',3,1,1,0,0,'临检',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',32,883,'260100006','爱迪氏计数','次','24h尿',3,1,1,0,0,'临检',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',32,884,'260100007','尿HCG(人绒毛膜促性腺激素)','次','晨尿',3,1,1,0,0,'临检',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',32,885,'260100008','大便常规','次','新鲜粪便',3,1,1,0,0,'临检',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',32,886,'260100009','大便集虫卵','次','粪便',3,1,1,0,0,'临检',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',32,887,'260100010','穿刺液镜检','次','浆膜腔积液',3,1,1,0,0,'临检',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',32,888,'260100011','穿刺液常规','次','浆膜腔积液',3,1,1,0,0,'临检',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',32,889,'260100012','胸腹水常规','次','浆膜腔积液',3,1,1,0,0,'临检',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',32,890,'260100013','腹腔穿刺液常规','次','浆膜腔积液',3,1,1,0,0,'临检',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',32,891,'260100014','关节液常规','次','浆膜腔积液',3,1,1,0,0,'临检',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',32,892,'260100015','脑脊液常规','次','脑脊液',3,1,1,0,0,'临检',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',32,893,'260100016','PT消耗纠正试验','次','血浆',3,1,1,0,0,'临检',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',32,894,'260100017','凝血三项','次','血浆',3,1,1,0,0,'临检',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',32,895,'260100018','凝血四项','次','血浆',3,1,1,0,0,'临检',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',32,896,'260100019','PT四项(凝血酶原时间)','次','血浆',3,1,1,0,0,'临检',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',32,897,'260100020','血块收缩时间','次','血浆',3,1,1,0,0,'临检',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',32,898,'260100021','痰咽拭子涂片','次','痰,咽拭子',3,1,1,0,0,'临检',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',32,899,'260100022','精液查WBC','次','精液',3,1,1,0,0,'临检',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',32,900,'260100023','分泌物找细菌','次','分泌物',3,1,1,0,0,'临检',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',32,901,'260100024','分泌物常规','次','分泌物',3,1,1,0,0,'临检',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',32,902,'260100025','白带常规','次','阴道分泌物',3,1,1,0,0,'临检',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',32,903,'260100026','前列腺液检查','次','前列腺液',3,1,1,0,0,'临检',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',32,904,'260100027','肝肾生化+心功能','次','血浆',3,1,1,0,0,'其他',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',32,905,'260100028','肝功能(快速机检)','次','血浆',3,1,1,0,0,'临检',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',32,906,'260100029','肝肾生化(快速机检)','次','血浆',3,1,1,0,0,'其他',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',32,907,'260100030','梅毒抗体','次','血浆',3,1,1,0,0,'临检',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',32,908,'260100031','消化系统肿瘤','次','消化道活检物',3,1,1,0,0,'其他',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',32,909,'260100032','前列腺检测','次','前列腺液',3,1,1,0,0,'临检',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',32,910,'260100033','糖类抗原','次','血浆',3,1,1,0,0,'临检',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',32,911,'260100034','性激素+Ca125','次','血浆',3,1,1,0,0,'临检',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',32,912,'260100035','甲功+性激素','次','血浆',3,1,1,0,0,'临检',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',32,913,'260100036','肿瘤指标','次','血浆',3,1,1,0,0,'临检',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',32,914,'260100037','心肌标志物','次','血浆',3,1,1,0,0,'临检',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',32,915,'260100038','DPD/UCRE(尿肌酸)','次','尿液',3,1,1,0,0,'临检',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',32,916,'260100039','乙肝速检','次','血浆',3,1,1,0,0,'临检',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',33,917,'260200001','电解质测定','次','血浆',3,1,1,0,0,'生化',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',33,918,'260200002','血气分析','次','动脉血',3,1,1,0,0,'生化',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',33,919,'260200003','血氧分析','次','动脉血',3,1,1,0,0,'生化',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',33,920,'260200004','血沉(ESR)','次','动脉抗凝血',3,1,1,0,0,'生化',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',33,921,'260200005','体检肝功','次','空腹血清',3,1,1,0,0,'生化',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',33,922,'260200006','肝功能二项','次','空腹血清',3,1,1,0,0,'生化',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',33,923,'260200007','肝功能四项','次','空腹血清',3,1,1,0,0,'生化',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',33,924,'260200008','肝功能八项','次','空腹血清',3,1,1,0,0,'生化',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',33,925,'260200009','肝功能十二项','次','空腹血清',3,1,1,0,0,'生化',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',33,926,'260200010','肝功能','次','空腹血清',3,1,1,0,0,'生化',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',33,927,'260200011','全肝功能','次','空腹血清',3,1,1,0,0,'生化',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',33,928,'260200012','肝肾生化','次','空腹血清',3,1,1,0,0,'生化',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',33,929,'260200013','肾功能','次','空腹血清',3,1,1,0,0,'生化',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',33,930,'260200014','全肾功','次','空腹血清',3,1,1,0,0,'生化',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',33,931,'260200015','血糖','次','空腹血清',3,1,1,0,0,'生化',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',33,932,'260200016','尿糖生化(体检)','次','晨尿',3,1,1,0,0,'生化',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',33,933,'260200017','1小时血糖','次','空腹血清',3,1,1,0,0,'生化',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',33,934,'260200018','2小时血糖','次','空腹血清',3,1,1,0,0,'生化',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',33,935,'260200019','3小时血糖','次','空腹血清',3,1,1,0,0,'生化',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',33,936,'260200020','血脂二项','次','血浆',3,1,1,0,0,'生化',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',33,937,'260200021','血脂四项','次','血浆',3,1,1,0,0,'生化',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',33,938,'260200022','血脂全套','次','血浆',3,1,1,0,0,'生化',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',33,939,'260200023','血脂分析(一)','次','血浆',3,1,1,0,0,'生化',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',33,940,'260200024','血脂分析(二)','次','血浆',3,1,1,0,0,'生化',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',33,941,'260200025','体检血脂','次','血浆',3,1,1,0,0,'生化',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',33,942,'260200026','尿生化','次','晨尿',3,1,1,0,0,'生化',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',33,943,'260200027','尿液代谢产物','次','24h尿',3,1,1,0,0,'生化',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',33,944,'260200028','尿石分析','次','晨尿',3,1,1,0,0,'生化',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',33,945,'260200029','脑脊液生化(一)','次','脑脊液',3,1,1,0,0,'生化',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',33,946,'260200030','脑脊液生化(二)','次','脑脊液',3,1,1,0,0,'生化',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',33,947,'260200031','骨髓生化','次','骨髓',3,1,1,0,0,'生化',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',33,948,'260200032','胸腹水生化','次','浆膜腔积液',3,1,1,0,0,'生化',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',33,949,'260200033','关节液生化','次','浆膜腔积液',3,1,1,0,0,'生化',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',33,950,'260200034','心肌酶谱','次','血液',3,1,1,0,0,'生化',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',33,951,'260200035','胆碱脂酶(CHE)','次','血液',3,1,1,0,0,'生化',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',33,952,'260200036','血淀粉酶(AMY)','次','血液',3,1,1,0,0,'生化',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',33,953,'260200037','尿淀粉酶','次','血液',3,1,1,0,0,'生化',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',33,954,'260200038','尿蛋白定量(M-TP)','次','血液',3,1,1,0,0,'生化',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',33,955,'260200039','甲胎蛋白(AFP)','次','空腹血清',3,1,1,0,0,'生化',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',33,956,'260200040','铁蛋白(Fer)','次','血液',3,1,1,0,0,'生化',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',33,957,'260200041','血清蛋白电泳','次','血液',3,1,1,0,0,'生化',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',33,958,'260200042','糖化血红蛋白(HBAIC)','次','血液',3,1,1,0,0,'生化',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',33,959,'260200043','胰岛素(Isu)','次','血液',3,1,1,0,0,'生化',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',33,960,'260200044','皮质醇(Cortisol)','次','血液',3,1,1,0,0,'生化',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',33,961,'260200045','性激素六项','次','血液',3,1,1,0,0,'生化',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',33,962,'260200046','生化九项','次','空腹血清',3,1,1,0,0,'生化',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',33,963,'260200047','生化全套','次','空腹血清',3,1,1,0,0,'生化',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',34,964,'260300001','甲型肝炎病毒抗体','次','空腹血清',3,1,1,0,0,'免疫学',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',34,965,'260300002','乙型肝炎病毒抗原抗体检测(两对半)','次','空腹血清',3,1,1,0,0,'免疫学',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',34,966,'260300003','丙型肝炎病毒抗体','次','空腹血清',3,1,1,0,0,'免疫学',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',34,967,'260300004','丁型肝炎病毒抗体','次','空腹血清',3,1,1,0,0,'免疫学',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',34,968,'260300005','戊型肝炎病毒抗体','次','空腹血清',3,1,1,0,0,'免疫学',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',34,969,'260300006','庚型肝炎病毒抗体','次','空腹血清',3,1,1,0,0,'免疫学',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',34,970,'260300007','免疫功能1','次','血清',3,1,1,0,0,'免疫学',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',34,971,'260300008','炎症指标','次','血清',3,1,1,0,0,'免疫学',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',34,972,'260300009','脑脊液功能','次','脑脊液',3,1,1,0,0,'免疫学',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',34,973,'260300010','肥达氏反应','次','血清',3,1,1,0,0,'免疫学',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',34,974,'260300011','外斐氏反应','次','血清',3,1,1,0,0,'免疫学',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',34,975,'260300012','HIV-HCV(爱滋病抗体)','次','血清',3,1,1,0,0,'免疫学',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',34,976,'260300013','VCA-EA(EB病毒抗体)','次','血清',3,1,1,0,0,'免疫学',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',34,977,'260300014','前列腺特异抗原(PSA)','次','前列腺液',3,1,1,0,0,'免疫学',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',34,978,'260300015','结核抗体(TB抗体)','次','血清',3,1,1,0,0,'免疫学',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',34,979,'260300016','RPR、梅毒血清确证试验(TPPA)','次','血清',3,1,1,0,0,'免疫学',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',34,980,'260300017','免疫功能2','次','血清',3,1,1,0,0,'免疫学',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',34,981,'260300018','补体C3、C4','次','血清',3,1,1,0,0,'免疫学',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',34,982,'260300019','过敏原皮试(每种物质)','次','血清',3,1,1,0,0,'免疫学',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',34,983,'260300020','血清β2微球蛋白','次','血清',3,1,1,0,0,'免疫学',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',34,984,'260300021','尿肾功','次','尿液',3,1,1,0,0,'免疫学',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',34,985,'260300022','营养贫血指标','次','血清',3,1,1,0,0,'免疫学',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',34,986,'260300023','风湿指标','次','血清',3,1,1,0,0,'免疫学',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',34,987,'260300024','风湿三项','次','血清',3,1,1,0,0,'免疫学',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',34,988,'260300025','风湿十项','次','血清',3,1,1,0,0,'免疫学',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',34,989,'260300026','抗O(ASO)','次','血清',3,1,1,0,0,'免疫学',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',34,990,'260300027','血小板抗体','次','血浆',3,1,1,0,0,'免疫学',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual;
Insert Into 诊疗项目目录(类别,分类id,id,编码,名称,计算单位,标本部位,计算方式,单独应用,执行频率,适用性别,组合项目,操作类型,执行安排,执行科室,服务对象,计价性质,建档时间,撤档时间)
Select 'C',34,991,'260300028','游离血小板抗体','次','血清',3,1,1,0,0,'免疫学',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',34,992,'260300029','ENA(可提取性核抗原)','次','血清',3,1,1,0,0,'免疫学',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',34,993,'260300030','ANCA(抗中性粒细胞胞浆抗体)','次','血清',3,1,1,0,0,'免疫学',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',34,994,'260300031','抗双链DNA抗体','次','血清',3,1,1,0,0,'免疫学',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',34,995,'260300032','抗核抗体(ANA)','次','血清',3,1,1,0,0,'免疫学',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',34,996,'260300033','癌胚抗原(CEA)','次','血清',3,1,1,0,0,'免疫学',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',34,997,'260300034','甲胎蛋白','次','空腹血清',3,1,1,0,0,'免疫学',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',34,998,'260300035','产检','次','血清',3,1,1,0,0,'免疫学',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',34,999,'260300036','新生儿免疫','次','血清',3,1,1,0,0,'免疫学',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',35,1000,'260400001','甲功Ⅱ','次','血清',3,1,1,0,0,'免疫学',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',35,1001,'260400002','甲功Ⅰ','次','血清',3,1,1,0,0,'免疫学',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',35,1002,'260400003','TGTM(抗甲状腺抗体)','次','血清',3,1,1,0,0,'免疫学',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',35,1003,'260400004','胃放免','次','血清',3,1,1,0,0,'免疫学',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',35,1004,'260400005','HA组合','次','血清',3,1,1,0,0,'免疫学',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',35,1005,'260400006','新生儿筛查','次','血清',3,1,1,0,0,'免疫学',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',35,1006,'260400007','肝纤维化','次','血清',3,1,1,0,0,'免疫学',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',35,1007,'260400008','乙肝放免','次','血清',3,1,1,0,0,'免疫学',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',35,1008,'260400009','骨代谢','次','血清',3,1,1,0,0,'免疫学',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',35,1009,'260400010','肿瘤放免','次','血清',3,1,1,0,0,'免疫学',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',35,1010,'260400011','胶原试验','次','血清',3,1,1,0,0,'免疫学',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',35,1011,'260400012','高血压肽类','次','血清',3,1,1,0,0,'免疫学',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',35,1012,'260400013','冠心病肽类','次','血清',3,1,1,0,0,'免疫学',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',35,1013,'260400014','肾素','次','血清',3,1,1,0,0,'免疫学',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',35,1014,'260400015','细胞因子','次','血清',3,1,1,0,0,'免疫学',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',35,1015,'260400016','产检放免','次','血清',3,1,1,0,0,'免疫学',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',35,1016,'260400017','免疫体检','次','血清',3,1,1,0,0,'免疫学',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1017,'260500001','涂片查淋球菌','次','分泌物',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1018,'260500002','直接涂片查细菌','次','分泌物',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1019,'260500003','腹水涂片','次','浆膜腔积液',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1020,'260500004','胸水涂片','次','胸水',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1021,'260500005','集菌法查结核杆菌','次','痰液',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1022,'260500006','痰涂片找细菌','次','痰液',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1023,'260500007','普通细菌涂片检查','次','分泌物',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1024,'260500008','直接涂片查真菌','次','分泌物',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1025,'260500009','结核杆菌培养','次','痰液',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1026,'260500010','血培养','次','血液',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1027,'260500011','骨髓细菌培养','次','骨髓',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1028,'260500012','脑脊液细菌培养','次','脑脊液',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1029,'260500013','尿培养','次','尿液',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1030,'260500014','大便培养','次','粪便',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1031,'260500015','胆汁细菌培养','次','胆汁',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1032,'260500016','脓及各种分泌物细菌培养','次','脓分泌物',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1033,'260500017','痰培养','次','痰液',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1034,'260500018','咽拭子培养','次','咽拭子',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1035,'260500019','各种穿刺液培养','次','浆膜腔积液',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1036,'260500020','胸、腹水细菌培养','次','浆膜腔积液',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1037,'260500021','血布氏杆菌培养','次','血液',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1038,'260500022','血液真菌培养','次','血液',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1039,'260500023','骨髓查黑热病小体','次','骨髓',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1040,'260500024','药敏联合试验','次','病理材料',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1041,'260500025','药物敏感试验(定性)','次','病理材料',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1042,'260500026','药物敏感试验(定量)','次','病理材料',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1043,'260500027','大便培养加药敏','次','粪便',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1044,'260500028','胆汁培养加药敏','次','胆汁',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1045,'260500029','脑脊液培养加药敏','次','脑脊液',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1046,'260500030','各种分泌物培养加药敏','次','分泌物',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1047,'260500031','血培养加药敏','次','血液',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1048,'260500032','痰培养加药敏','次','痰液',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1049,'260500033','胸腹水培养加药敏','次','浆膜腔积液',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1050,'260500034','咽试子培养加药敏','次','咽拭子',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1051,'260500035','尿培养加细菌计数加药敏','次','尿液',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1052,'260500036','结核菌药敏试验','次','痰液',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1053,'260500037','新型隐球菌培养','次','病理材料',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1054,'260500038','霉菌培养加药敏','次','病理材料',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1055,'260500039','真菌药敏试验','次','病理材料',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1056,'260500040','真菌培养加药敏','次','病理材料',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1057,'260500041','血液真菌培养加药敏','次','血液',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1058,'260500042','微生物鉴定','次','病理材料',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1059,'260500043','细菌计数','次','病理材料',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1060,'260500044','菌种鉴定','次','病理材料',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1061,'260500045','细菌内毒素','次','血浆',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1062,'260500046','各种无菌试验','次','病理材料',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1063,'260500047','普通细菌培养','次','病理材料',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1064,'260500048','少见病原菌培养','次','病理材料',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1065,'260500049','耐甲氧苯青霉素葡球菌培养','次','病理材料',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1066,'260500050','L型菌培养','次','病理材料',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1067,'260500051','幽门螺杆菌培养','次','胃液',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1068,'260500052','致病性大肠杆菌培养','次','病理材料',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1069,'260500053','幽门螺杆菌感染快速试验','次','胃液',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1070,'260500054','O2增菌及培养(含鉴定)','次','病理材料',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1071,'260500055','O2悬滴及培养(含增菌,培养)','次','病理材料',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1072,'260500056','支原体培养','次','病理材料',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1073,'260500057','衣原体培养','次','病理材料',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1074,'260500058','真菌培养','次','病理材料',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1075,'260500059','霉菌培养','次','病理材料',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1076,'260500060','疟原虫检查','次','血浆',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1077,'260500061','大便找虫卵','次','粪便',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1078,'260500062','血液查寄生虫','次','血液',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1079,'260500063','尿液查寄生虫','次','尿液',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1080,'260500064','痰液寄生虫检验','次','痰液',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1081,'260500065','肝吸虫检查','次','血液',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1082,'260500066','寄生虫或幼虫鉴定','次','粪便',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1083,'260500067','虫卵计数或成虫计数','次','粪便',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1084,'260500068','血吸虫皮肤试验','次','皮肤',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1085,'260500069','肺吸虫皮肤试验','次','皮肤',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1086,'260500070','肝吸虫皮肤试验','次','皮肤',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',36,1087,'260500071','猪囊虫皮肤试验','次','皮肤',3,1,1,0,0,'微生物',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',37,1088,'270000001','ABO红细胞定型','次','',3,1,1,0,0,'其他',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',37,1089,'270000002','ABO血型鉴定','次','',3,1,1,0,0,'其他',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',37,1090,'270000003','ABO亚型鉴定','次','',3,1,1,0,0,'其他',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',37,1091,'270000004','Rh血型鉴定','次','',3,1,1,0,0,'其他',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',37,1092,'270000005','Rh血型其他抗原鉴定','次','',3,1,1,0,0,'其他',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',37,1093,'270000006','血型抗体特异性鉴定','次','',3,1,1,0,0,'其他',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',37,1094,'270000007','交叉配血','次','',3,1,1,0,0,'其他',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',37,1095,'270000008','唾液ABH血型物质测定','次','',3,1,1,0,0,'其他',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'K',37,1096,'270000009','全血','单位','',1,1,1,0,0,'',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'K',37,1097,'270000010','红细胞悬液','单位','',1,1,1,0,0,'',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'K',37,1098,'270000011','洗涤红细胞','单位','',1,1,1,0,0,'',0,1,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'K',37,1099,'270000012','血小板悬液','单位','',1,1,1,0,0,'',0,1,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'K',37,1100,'270000013','单采血小板','单位','',1,1,1,0,0,'',0,1,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual;
Insert Into 诊疗项目目录(类别,分类id,id,编码,名称,计算单位,标本部位,计算方式,单独应用,执行频率,适用性别,组合项目,操作类型,执行安排,执行科室,服务对象,计价性质,建档时间,撤档时间)
Select 'K',37,1101,'270000014','新鲜冷冻血浆','单位','',1,1,1,0,0,'',0,1,2,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',39,1102,'280100001','体液细胞学检查与诊断','次','',3,1,1,0,0,'其他',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',39,1103,'280100002','拉网细胞学检查与诊断','次','',3,1,1,0,0,'其他',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',39,1104,'280100003','细针穿刺细胞学检查与诊断','次','',3,1,1,0,0,'其他',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',39,1105,'280100004','脱落细胞学检查与诊断','次','',3,1,1,0,0,'其他',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',39,1106,'280100005','细胞学计数','次','',3,1,1,0,0,'其他',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',40,1107,'280200001','穿刺组织活检检查与诊断','次','',3,1,1,0,0,'其他',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',40,1108,'280200002','内镜组织活检检查与诊断','次','',3,1,1,0,0,'其他',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',40,1109,'280200003','局部切除活检检查与诊断','次','',3,1,1,0,0,'其他',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',40,1110,'280200004','骨髓组织活检检查与诊断','次','',3,1,1,0,0,'其他',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',40,1111,'280200005','大手术标本检查与诊断','次','',3,1,1,0,0,'其他',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',40,1112,'280200006','中手术标本检查与诊断','次','',3,1,1,0,0,'其他',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',40,1113,'280200007','小手术标本检查与诊断','次','',3,1,1,0,0,'其他',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',40,1114,'280200008','大截肢标本病理检查与诊断','次','',3,1,1,0,0,'其他',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',40,1115,'280200009','小截肢标本病理检查与诊断','次','',3,1,1,0,0,'其他',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',40,1116,'280200010','牙齿及骨骼磨片诊断(不脱钙)','次','',3,1,1,0,0,'其他',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',40,1117,'280200011','牙齿及骨骼磨片诊断(脱钙)','次','',3,1,1,0,0,'其他',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',40,1118,'280200012','部分颌骨样本及牙体牙周样本','次','',3,1,1,0,0,'其他',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',40,1119,'280200013','单侧颌骨切除样本诊断','次','',3,1,1,0,0,'其他',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',40,1120,'280200014','全器官大切片','次','',3,1,1,0,0,'其他',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',41,1121,'280300001','冰冻切片检查与诊断','次','',3,1,1,0,0,'其他',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',41,1122,'280300002','石蜡切片检查与诊断','次','',3,1,1,0,0,'其他',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',42,1123,'280400001','结缔组织及肌肉组织的鉴别染色','次','',3,1,1,0,0,'其他',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',42,1124,'280400002','脂质染色','次','',3,1,1,0,0,'其他',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',42,1125,'280400003','核酸染色','次','',3,1,1,0,0,'其他',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',42,1126,'280400004','色素和无机物染色','次','',3,1,1,0,0,'其他',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',42,1127,'280400005','内分泌腺细胞、产肽细胞(APUD)染色','次','',3,1,1,0,0,'其他',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',42,1128,'280400006','神经组织染色','次','',3,1,1,0,0,'其他',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',42,1129,'280400007','微生物染色','次','',3,1,1,0,0,'其他',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',42,1130,'280400008','显示酶的组织化学法','次','',3,1,1,0,0,'其他',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',42,1131,'280400009','免疫酶技术','次','',3,1,1,0,0,'其他',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',42,1132,'280400010','免疫荧光技术','次','',3,1,1,0,0,'其他',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',43,1133,'280500001','成人尸体解剖','次','',3,1,1,0,0,'其他',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'C',43,1134,'280500002','婴幼胎儿尸体解剖','次','',3,1,1,0,0,'其他',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'G',45,1135,'310100001','基础麻醉','次','',3,0,1,0,0,'局麻',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'G',45,1136,'310100002','静脉全身麻醉','次','',3,0,1,0,0,'静脉',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'G',45,1137,'310100003','静脉普鲁卡因复合麻醉','次','',3,0,1,0,0,'其他',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'G',45,1138,'310100004','吸入全身麻醉','次','',3,0,1,0,0,'全麻',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'G',45,1139,'310100005','气管、支气管内插管术','次','',3,0,1,0,0,'其他',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'G',45,1140,'310100006','连续硬膜外腔阻滞麻醉','次','',3,0,1,0,0,'局麻',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'G',45,1141,'310100007','骶管阻滞麻醉','次','',3,0,1,0,0,'局麻',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'G',45,1142,'310100008','脊椎麻醉(中位腰麻)','次','',3,0,1,0,0,'局麻',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'G',45,1143,'310100009','脊椎麻醉(低位腰麻)','次','',3,0,1,0,0,'局麻',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'G',45,1144,'310100010','脊椎麻醉(鞍麻)','次','',3,0,1,0,0,'局麻',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'G',45,1145,'310100011','颈丛神经阻滞','次','',3,0,1,0,0,'颈丛',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'G',45,1146,'310100012','臂丛神经阻滞','次','',3,0,1,0,0,'臂丛',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'G',45,1147,'310100013','中心静脉穿剌置管术','次','',3,0,1,0,0,'其他',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'G',45,1148,'310100014','动脉穿刺置管术','次','',3,0,1,0,0,'其他',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'G',45,1149,'310100015','控制性降压','次','',3,0,1,0,0,'其他',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'G',45,1150,'310100016','低温麻醉','次','',3,0,1,0,0,'其他',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1151,'310210001','脑池穿刺术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1152,'310210002','小脑延髓池穿刺','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1153,'310210003','脑室穿刺,经植入导管','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1154,'310210004','硬脑下抽吸术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1155,'310210005','颅穿剌术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1156,'310210006','蛛网膜下腔抽吸术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1157,'310210007','前囱门穿刺','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1158,'310210008','小脑穿刺术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1159,'310210009','脑膜活组织检查,经皮','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1160,'310210010','脑膜活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1161,'310210011','脑活组织检查,颅骨穿孔','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1162,'310210012','脑活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1163,'310210013','颅静脉窦切开引流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1164,'310210014','颅内神经刺激器的去除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1165,'310210015','开颅探查术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1166,'310210016','硬脑膜外脓肿清除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1167,'310210017','硬脑膜外血肿清除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1168,'310210018','颅内减压术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1169,'310210019','颅内脓肿引流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1170,'310210020','颅内异物取出术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1171,'310210021','颅死骨切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1172,'310210022','颅外伤清创术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1173,'310210023','脑蛛网膜下脓肿清除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1174,'310210024','蛛网膜下腔血肿清除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1175,'310210025','硬脑膜下脓肿清除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1176,'310210026','蛛网膜下血肿清除术,脑部','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1177,'310210027','硬脑膜下血肿清除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1178,'310210028','蛛网膜下脓肿清除术,脑部','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1179,'310210029','脑脓肿切开引流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1180,'310210030','脑脓肿清除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1181,'310210031','脑室引流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1182,'310210032','脑血肿清除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1183,'310210033','丘脑切开术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1184,'310210034','苍白球切开术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1185,'310210035','脑叶切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1186,'310210036','颅骨凹陷骨折切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1187,'310210037','颅骨病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1188,'310210038','颅骨肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1189,'310210039','颅内肉芽切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1190,'310210040','线形颅骨切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1191,'310210041','颅骨骨折复位术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1192,'310210042','颅骨骨折清创术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1193,'310210043','颅骨骨折减压术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1194,'310210044','颅骨修补伴有骨瓣','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1195,'310210045','颅骨修补伴有骨移植','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1196,'310210046','颅骨金属板置换术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1197,'310210047','颅骨修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1198,'310210048','颅骨金属板去除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1199,'310210049','硬脑膜单纯缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1200,'310210050','脑脊液瘘修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1201,'310210051','脑膨出修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1202,'310210052','脑膨出修补术伴颅内成形术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1203,'310210053','中脑膜动脉结扎术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1204,'310210054','脑室脑池造瘘术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1205,'310210055','脑室鼻咽分流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1206,'310210056','脑室乳突吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1207,'310210057','脑室腔静脉分流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1208,'310210058','脑室心房分流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1209,'310210059','脑室胸腔分流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1210,'310210060','脑室胆囊分流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual;
Insert Into 诊疗项目目录(类别,分类id,id,编码,名称,计算单位,标本部位,计算方式,单独应用,执行频率,适用性别,组合项目,操作类型,执行安排,执行科室,服务对象,计价性质,建档时间,撤档时间)
Select 'F',47,1211,'310210061','脑室腹腔分流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1212,'310210062','脑室输尿管分流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1213,'310210063','脑室骨髓分流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1214,'310210064','脑室分流管冲洗术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1215,'310210065','脑室导管置换术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1216,'310210066','脑室分流管却除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1217,'310210067','脑修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1218,'310210068','大脑皮层粘连松解术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',47,1219,'310210069','中脑导水管粘连松解术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',48,1220,'310220001','听神经探查术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',48,1221,'310220002','三叉神经切断术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',48,1222,'310220003','闭孔神经切断术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',48,1223,'310220004','神经松解术,周围神经','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',48,1224,'310220005','鼓室神经丛切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',48,1225,'310220006','颅神经瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',48,1226,'310220007','面神经解剖术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',48,1227,'310220008','神经病损切除术,颅或周围神经','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',48,1228,'310220009','神经瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',48,1229,'310220010','神经活组织检查,颅或周围神经','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',48,1230,'310220011','三叉神经半月节热射频治疗','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',48,1231,'310220012','正中神经缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',48,1232,'310220013','尺神经缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',48,1233,'310220014','神经缝合术,颅或周围神经','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',48,1234,'310220015','三叉神经减压术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',48,1235,'310220016','三叉神经松解术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',48,1236,'310220017','颅神经减压术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',48,1237,'310220018','面神经减压术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',48,1238,'310220019','神经松解术,颅神经','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',48,1239,'310220020','视神经减压术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',48,1240,'310220021','听神经减压术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',48,1241,'310220022','跗管内神经松解术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',48,1242,'310220023','腕管内神经松解术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',48,1243,'310220024','正中神经松解术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',48,1244,'310220025','尺神经松解术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',48,1245,'310220026','神经移植术,颅及周围神经','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',48,1246,'310220027','指神经移位术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',48,1247,'310220028','尺神经移位术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',48,1248,'310220029','神经移位术,颅或周围神经','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',48,1249,'310220030','舌下--面神经吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',48,1250,'310220031','副-面神经吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',48,1251,'310220032','副-舌下神经吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',48,1252,'310220033','闭孔神经吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',48,1253,'310220034','尺神经吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',48,1254,'310220035','桡神经吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',48,1255,'310220036','神经吻合术,颅或周围神经','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',48,1256,'310220037','交感神经切断术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',48,1257,'310220038','交感神经活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',48,1258,'310220039','交感神经病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',48,1259,'310220040','交感神经神经瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',48,1260,'310220041','交感神经缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',49,1261,'310230001','蛛网膜囊肿切除术,脑部','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',49,1262,'310230002','脑膜肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',49,1263,'310230003','脑脑膜瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',49,1264,'310230004','脑蛛网膜囊肿切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',49,1265,'310230005','异位小脑扁桃体切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',49,1266,'310230006','脑病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',49,1267,'310230007','脑囊肿切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',49,1268,'310230008','脑囊肿造袋术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',49,1269,'310230009','脑清创术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',49,1270,'310230010','脑肿瘤部分切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',49,1271,'310230011','脑肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',49,1272,'310230012','脉络丛血管瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',49,1273,'310230013','脉络丛灼烙术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',50,1274,'310240001','椎管内异物去除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',50,1275,'310240002','椎管减压术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',50,1276,'310240003','椎管探查术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',50,1277,'310240004','脊髓神经根探查术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',50,1278,'310240005','脊髓探查术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',50,1279,'310240006','椎板切除术(减压)','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',50,1280,'310240007','椎管内神经根切断术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',50,1281,'310240008','腰椎穿刺术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',50,1282,'310240009','脊髓活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',50,1283,'310240010','脊髓膜活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',50,1284,'310240011','椎管内病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',50,1285,'310240012','椎管内脑膜瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',50,1286,'310240013','椎管内脓肿切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',50,1287,'310240014','椎管内肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',50,1288,'310240015','脊髓膜病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',50,1289,'310240016','脊髓肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',50,1290,'310240017','马尾神经肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',50,1291,'310240018','脊髓脑膜疝出修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',50,1292,'310240019','脊髓膜疝出修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',50,1293,'310240020','脊椎骨折复位术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',50,1294,'310240021','脊柱裂修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',50,1295,'310240022','蛛网膜(脊髓)粘连松解术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',50,1296,'310240023','脊髓粘连松解术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',50,1297,'310240024','脊髓神经根粘连松解术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',50,1298,'310240025','脊髓蛛网膜下腹腔分流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',50,1299,'310240026','脊髓蛛网膜下输尿管分流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',52,1300,'310310001','甲状腺全部切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',52,1301,'310310002','甲状旁腺探查术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',52,1302,'310310003','甲状腺探查术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',52,1303,'310310004','甲状腺活组织检查,穿剌术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',52,1304,'310310005','甲状腺活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',52,1305,'310310006','甲状旁腺活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',52,1306,'310310007','甲状腺叶切除术, 单侧','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',52,1307,'310310008','甲状腺病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',52,1308,'310310009','甲状腺结节切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',52,1309,'310310010','甲状腺囊肿切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',52,1310,'310310011','甲状腺腺瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',52,1311,'310310012','甲状腺肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',52,1312,'310310013','甲状腺部分切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',52,1313,'310310014','甲状腺次全切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',52,1314,'310310015','胸骨下甲状腺切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',52,1315,'310310016','胸骨下甲状腺部分切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',52,1316,'310310017','甲状舌管瘘切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',52,1317,'310310018','甲状舌管囊肿切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',52,1318,'310310019','甲状旁腺全部切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',52,1319,'310310020','甲状旁腺部分切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',52,1320,'310310021','甲状旁腺腺瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual;
Insert Into 诊疗项目目录(类别,分类id,id,编码,名称,计算单位,标本部位,计算方式,单独应用,执行频率,适用性别,组合项目,操作类型,执行安排,执行科室,服务对象,计价性质,建档时间,撤档时间)
Select 'F',52,1321,'310310022','甲状腺血管结扎术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',52,1322,'310310023','甲状腺缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',52,1323,'310310024','甲状腺自体移植术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',52,1324,'310310025','甲状旁腺同种异体移植术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',52,1325,'310310026','甲状旁腺自体移植术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',53,1326,'310320001','肾上腺活组织检查, 经皮','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',53,1327,'310320002','肾上腺活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',53,1328,'310320003','垂体腺活组织检查, 经额','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',53,1329,'310320004','垂体腺活组织检查, 经蝶部','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',53,1330,'310320005','胸腺活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',53,1331,'310320006','松果体活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',53,1332,'310320007','肾上腺病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',53,1333,'310320008','肾上腺囊肿切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',53,1334,'310320009','肾上腺肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',53,1335,'310320010','肾上腺切除术,单侧','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',53,1336,'310320011','肾上腺部分切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',53,1337,'310320012','肾上腺切除术,双侧','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',53,1338,'310320013','肾上腺探查术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',53,1339,'310320014','肾上腺神经切断术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',53,1340,'310320015','肾上腺血管结扎术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',53,1341,'310320016','肾上腺修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',53,1342,'310320017','肾上腺自体移植术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',53,1343,'310320018','松果腺探查术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',53,1344,'310320019','松果腺切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',53,1345,'310320020','垂体病损切除术, 经额','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',53,1346,'310320021','垂体瘤切除术, 经额','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',53,1347,'310320022','垂体瘤切除术, 经蝶窦','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',53,1348,'310320023','垂体腺切除术, 经额','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',53,1349,'310320024','垂体腺切除术, 经蝶窦','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',53,1350,'310320025','垂体窝探查术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',53,1351,'310320026','垂体腺探查术, 经蝶窦','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',53,1352,'310320027','胸腺部分切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',53,1353,'310320028','胸腺瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',53,1354,'310320029','胸腺囊肿切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',53,1355,'310320030','胸腺切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',53,1356,'310320031','胸腺探查术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',53,1357,'310320032','胸腺修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',53,1358,'310320033','胸腺移植术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',53,1359,'310320034','胸腺固定术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',55,1360,'310410001','眼睑粘连分离术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',55,1361,'310410002','眼睑切开术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',55,1362,'310410003','眼睑探查术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',55,1363,'310410004','眼睑病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',55,1364,'310410005','睑板腺囊肿切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',55,1365,'310410006','霰粒肿切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',55,1366,'310410007','眼睑大病损切除术(切除眼脸缘板层1/4或以上','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',55,1367,'310410008','眼睑大病损切除术, 全层(1/4或更多,全层)','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',55,1368,'310410009','眼睑病损破坏术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',55,1369,'310410010','额肌缝线术(眼睑下垂矫正术)','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',55,1370,'310410011','眼睑下垂矫正术, 额肌缝线法','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',55,1371,'310410012','额肌-筋膜吊带术(眼睑下垂矫正术)','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',55,1372,'310410013','眼睑下垂矫正术, 额肌-筋膜吊带术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',55,1373,'310410014','上睑提肌缩短术(睑下垂矫正)','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',55,1374,'310410015','眼睑下垂矫正术, 睑板法','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',55,1375,'310410016','眼睑内翻热灼修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',55,1376,'310410017','眼睑外翻热灼修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',55,1377,'310410018','眼睑内翻缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',55,1378,'310410019','眼睑外翻缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',55,1379,'310410020','眼睑内翻楔形切除修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',55,1380,'310410021','眼睑外翻楔形切除修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',55,1381,'310410022','眼睑内翻矫正术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',55,1382,'310410023','眼睑外翻矫正术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',55,1383,'310410024','外眦缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',55,1384,'310410025','外眦成形术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',55,1385,'310410026','眼睑皮瓣重建术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',55,1386,'310410027','眼睑皮肤粘膜移植术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',55,1387,'310410028','眼睑毛囊移植片重建术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',55,1388,'310410029','眼睑结膜睑板移植重建术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',55,1389,'310410030','眼睑成形术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',55,1390,'310410031','眼睑重建术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',55,1391,'310410032','眼睑重建术涉及眼睑缘和板层','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',55,1392,'310410033','眼睑重建术涉及眼睑缘及全层','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',56,1393,'310420001','泪囊切七引流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',56,1394,'310420002','泪腺活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',56,1395,'310420003','泪囊活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',56,1396,'310420004','泪腺病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',56,1397,'310420005','泪腺部分切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',56,1398,'310420006','泪腺切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',56,1399,'310420007','泪点扩张术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',56,1400,'310420008','泪小管扩张术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',56,1401,'310420009','鼻泪管狭窄探通术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',56,1402,'310420010','鼻泪管扩张模置入术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',56,1403,'310420011','泪点切开术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',56,1404,'310420012','泪小管切开术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',56,1405,'310420013','泪囊切开术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',56,1406,'310420014','鼻泪管切开引流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',56,1407,'310420015','泪囊病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',56,1408,'310420016','泪囊囊肿切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',56,1409,'310420017','泪囊切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',56,1410,'310420018','泪小管切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',56,1411,'310420019','泪点外翻纠正术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',56,1412,'310420020','泪点修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',56,1413,'310420021','泪管成形术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',56,1414,'310420022','泪囊鼻腔造口术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',56,1415,'310420023','结膜泪囊鼻腔造口术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',57,1416,'310430001','结膜异物切开去除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',57,1417,'310430002','球结膜环状切开术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',57,1418,'310430003','结膜病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',57,1419,'310430004','阴道隔膜切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',57,1420,'310430005','睑球粘连游离移植物修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',57,1421,'310430006','结膜穹窿游离移植物重建术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',57,1422,'310430007','结膜穹窿重建术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',57,1423,'310430008','结膜移植术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',57,1424,'310430009','结膜成形术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',57,1425,'310430010','结膜穹窿成形术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',57,1426,'310430011','睑球(眼睑结膜)粘连分离术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',57,1427,'310430012','眼睑结膜(睑球)粘连分离术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',57,1428,'310430013','结膜缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',57,1429,'310430014','结膜撕裂修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',58,1430,'310440001','角膜异物磁吸术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual;
Insert Into 诊疗项目目录(类别,分类id,id,编码,名称,计算单位,标本部位,计算方式,单独应用,执行频率,适用性别,组合项目,操作类型,执行安排,执行科室,服务对象,计价性质,建档时间,撤档时间)
Select 'F',58,1431,'310440002','角膜切开术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',58,1432,'310440003','角膜异物切开去除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',58,1433,'310440004','角膜活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',58,1434,'310440005','翼状胬肉切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',58,1435,'310440006','翼状胬肉切除术伴骨膜移植术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',58,1436,'310440007','角膜病损热烙术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',58,1437,'310440008','角膜病损冷冻术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',58,1438,'310440009','角巩膜缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',58,1439,'310440010','角膜缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',58,1440,'310440011','白内障手术伤口修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',58,1441,'310440012','角膜结膜成形术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',58,1442,'310440013','角膜修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',58,1443,'310440014','角膜移植术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',58,1444,'310440015','角巩板层自体移植术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',58,1445,'310440016','角膜板层移植术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',58,1446,'310440017','角膜自体全层移植术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',58,1447,'310440018','角膜全层移植术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',58,1448,'310440019','植入角膜去除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',58,1449,'310440020','眼肌活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',58,1450,'310440021','眼肌延长术, 一条眼肌','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',58,1451,'310440022','眼肌缩短术, 一条眼肌','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',58,1452,'310440023','降结肠部分切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',58,1453,'310440024','斜视手术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',59,1454,'310450001','虹膜贯穿术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',59,1455,'310450002','虹膜剪除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',59,1456,'310450003','虹膜括约肌切断术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',59,1457,'310450004','虹膜切开术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',59,1458,'310450005','虹膜脱出切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',59,1459,'310450006','虹膜分离切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',59,1460,'310450007','虹膜激光切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',59,1461,'310450008','虹膜切除嵌顿术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',59,1462,'310450009','虹膜周边切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',59,1463,'310450010','虹膜前粘连分离术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',59,1464,'310450011','虹膜粘连剥离术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',59,1465,'310450012','瞳孔成形术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',59,1466,'310450013','虹膜病损破坏术, 非切除性','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',59,1467,'310450014','虹膜,病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',59,1468,'310450015','虹膜囊肿切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',59,1469,'310450016','虹膜睫状体切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',59,1470,'310450017','睫状体切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',59,1471,'310450018','前房角穿刺术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',59,1472,'310450019','前房角切开术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',59,1473,'310450020','前房角切开伴穿刺术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',59,1474,'310450021','小梁切开术,外入路','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',59,1475,'310450022','睫状体分离术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',59,1476,'310450023','睫状体切开术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',59,1477,'310450024','巩膜环钻术伴虹膜切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',59,1478,'310450025','虹膜切除伴巩膜环钻术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',59,1479,'310450026','谢氏SCHEIE巩膜造口术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',59,1480,'310450027','谢氏SCHEIE巩膜灼烙术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',59,1481,'310450028','虹膜嵌顿术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',59,1482,'310450029','滤帘切除术(小梁切除术),外路','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',59,1483,'310450030','虹膜巩膜切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',59,1484,'310450031','虹膜切除伴滤过术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',59,1485,'310450032','睫状体透热术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',59,1486,'310450033','睫状体冷冻术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',59,1487,'310450034','睫状体光凝固术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',59,1488,'310450035','睫状体贫血术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',59,1489,'310450036','前房异管术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',59,1490,'310450037','巩膜缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',59,1491,'310450038','巩膜漏修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',59,1492,'310450039','巩膜病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',59,1493,'310450040','巩膜灼烙术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',59,1494,'310450041','巩膜外加压术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',59,1495,'310450042','巩膜移植物加固术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',59,1496,'310450043','巩膜成形术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',59,1497,'310450044','巩膜修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',59,1498,'310450045','角巩膜环钻术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',59,1499,'310450046','角巩膜咬切术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',59,1500,'310450047','角膜穿剌术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',59,1501,'310450048','睫状体放液术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',59,1502,'310450049','前房抽吸术,治疗性','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',59,1503,'310450050','前房穿刺术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',59,1504,'310450051','前房注气术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',59,1505,'310450052','睫状体缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',59,1506,'310450053','前房导管取出术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',60,1507,'310460001','晶状体异物去除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',60,1508,'310460002','晶状体异物磁吸术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',60,1509,'310460003','晶状体异物去除术,非磁吸性','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',60,1510,'310460004','白内障囊内摘除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',60,1511,'310460005','白内障吸取术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',60,1512,'310460006','白内障摘除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',60,1513,'310460007','晶状体冷冻摘除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',60,1514,'310460008','晶状体囊切开伴晶状体切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',60,1515,'310460009','晶状体囊外线形摘除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',60,1516,'310460010','外伤性白内障冲洗术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',60,1517,'310460011','白内障晶状体乳化抽吸术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',60,1518,'310460012','白内障后路切割吸出术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',60,1519,'310460013','晶状体切割吸出术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',60,1520,'310460014','晶状体囊膜剪除术,伴晶状体摘除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',60,1521,'310460015','白内障囊外摘除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',60,1522,'310460016','晶状体囊膜剪除术,原发性白内障','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',60,1523,'310460017','原发膜性白内障截囊术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',60,1524,'310460018','原发性白内障晶体囊膜剪除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',60,1525,'310460019','晶状体囊切开术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',60,1526,'310460020','原发膜性白内障切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',60,1527,'310460021','复发性白内障晶体囊膜剪除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',60,1528,'310460022','后发性膜(白内障后)截囊术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',60,1529,'310460023','晶状体囊膜剪除术,复发性白内障','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',60,1530,'310460024','白内障剪除术, 复发性','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',60,1531,'310460025','复发性白内障切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',60,1532,'310460026','后发性膜(白内障后)切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',60,1533,'310460027','晶状体囊切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',60,1534,'310460028','人工晶体植入术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',60,1535,'310460029','白内障摘除伴人工晶体植入术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',60,1536,'310460030','人工晶体伴白内障摘除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',60,1537,'310460031','人工晶体植入术,白内障摘除术后','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',60,1538,'310460032','植入晶体去除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',60,1539,'310460033','人工晶体复位术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',61,1540,'310470001','眼后节异物去除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual;
Insert Into 诊疗项目目录(类别,分类id,id,编码,名称,计算单位,标本部位,计算方式,单独应用,执行频率,适用性别,组合项目,操作类型,执行安排,执行科室,服务对象,计价性质,建档时间,撤档时间)
Select 'F',61,1541,'310470002','咽后节异物磁吸术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',61,1542,'310470003','玻璃体豚囊虫取出术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',61,1543,'310470004','眼后节异物去除术, 未用磁吸','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',61,1544,'310470005','脉络膜病损透热术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',61,1545,'310470006','视网膜病损透热破坏术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',61,1546,'310470007','脉络膜血管瘤冷冻破坏术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',61,1547,'310470008','视网膜病损冷冻破坏术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',61,1548,'310470009','脉络膜病损氙弧光凝固术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',61,1549,'310470010','视网膜病损氙弧光凝固术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',61,1550,'310470011','脉络膜病损激光光凝固术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',61,1551,'310470012','视网膜病损激光光凝固术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',61,1552,'310470013','视网膜撕裂透热修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',61,1553,'310470014','视网膜撕裂冷冻修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',61,1554,'310470015','视网膜撕裂氙弧光凝固术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',61,1555,'310470016','视网膜撕裂激光光凝固术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',61,1556,'310470017','巩膜环扎伴填充术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',61,1557,'310470018','巩膜环扎术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',61,1558,'310470019','视网膜脱离透热再接术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',61,1559,'310470020','视网膜脱离再接合冷冻术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',61,1560,'310470021','视网膜脱离氙弧光凝固术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',61,1561,'310470022','视网膜脱离激光治疗','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',61,1562,'310470023','巩膜缩短术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',61,1563,'310470024','玻璃体抽吸术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',61,1564,'310470025','玻璃体切割术, 经瞳孔','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',61,1565,'310470026','玻璃体切割术, 前入路','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',61,1566,'310470027','玻璃体切割术, 后入路','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',61,1567,'310470028','玻璃体内注射代替物','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',61,1568,'310470029','玻璃体置入术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',61,1569,'310470030','视网膜下放液术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',62,1570,'310480001','眶外侧壁切开术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',62,1571,'310480002','开眶探查术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',62,1572,'310480003','眶内容物剜出术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',62,1573,'310480004','眶内血管瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',62,1574,'310480005','眶内肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',64,1575,'310510001','外耳道囊肿切开引流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',64,1576,'310510002','耳前脓肿切开引流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',64,1577,'310510003','耳前窦道切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',64,1578,'310510004','耳前囊肿切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',64,1579,'310510005','耳前肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',64,1580,'310510006','耳廓部分切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',64,1581,'310510007','耳廓血管瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',64,1582,'310510008','耳廓肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',64,1583,'310510009','副耳切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',64,1584,'310510010','外耳道瘢痕切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',64,1585,'310510011','外耳道病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',64,1586,'310510012','外耳道胆脂瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',64,1587,'310510013','外耳道囊肿切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',64,1588,'310510014','外耳道血管瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',64,1589,'310510015','外耳道痣切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',64,1590,'310510016','外耳肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',64,1591,'310510017','外耳道缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',64,1592,'310510018','外耳缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',64,1593,'310510019','耳后贴矫正术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',64,1594,'310510020','外耳道成形术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',64,1595,'310510021','耳廓建造术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',64,1596,'310510022','断耳再接术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',64,1597,'310510023','外耳整形术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',64,1598,'310510024','耳部切口扩创术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',65,1599,'310520001','镫骨松动术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',65,1600,'310520002','眶前路切开术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',65,1601,'310520003','镫骨切除伴砧骨置换术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',65,1602,'310520004','镫骨部分切除伴脂肪移植术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',65,1603,'310520005','镫骨切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',65,1604,'310520006','鼓膜成形术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',65,1605,'310520007','鼓膜修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',65,1606,'310520008','鼓膜灼烙术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',65,1607,'310520009','鼓室成形术, Ⅰ型','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',65,1608,'310520010','鼓室成形术, Ⅱ型','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',65,1609,'310520011','鼓室成形术, Ⅲ型','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',65,1610,'310520012','鼓室成形术, Ⅳ型','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',65,1611,'310520013','鼓室成形术, Ⅴ型','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',65,1612,'310520014','鼓室成形术后修正术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',65,1613,'310520015','中耳修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',65,1614,'310520016','乳突瘘关闭术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',65,1615,'310520017','鼓膜切开植管术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',65,1616,'310520018','鼓膜造口术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',65,1617,'310520019','鼓膜切开术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',65,1618,'310520020','鼓室切开术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',65,1619,'310520021','鼓室探查术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',65,1620,'310520022','中耳抽吸术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',65,1621,'310520023','鼓室造瘘管取出术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',65,1622,'310520024','乳突探查术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',65,1623,'310520025','中耳粘连松解术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',65,1624,'310520026','中耳探查术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',65,1625,'310520027','耳蜗电图','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',65,1626,'310520028','中耳活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',65,1627,'310520029','内耳活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',65,1628,'310520030','乳突单纯切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',65,1629,'310520031','乳突根治术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',65,1630,'310520032','乳头改良根治术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',65,1631,'310520033','中耳病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',65,1632,'310520034','鼓膜切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',65,1633,'310520035','半规管开窗术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',65,1634,'310520036','迷路开窗术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',65,1635,'310520037','前庭开窗术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',65,1636,'310520038','内淋巴分流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',65,1637,'310520039','迷路减压术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',65,1638,'310520040','内耳引流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',65,1639,'310520041','耳蜗单道接收器伴电极植入术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',65,1640,'310520042','耳蜗多道接收器伴电极植入术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',65,1641,'310520043','半规管瘘修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',65,1642,'310520044','耳蜗电极植入术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',67,1643,'310610001','鼻出血止血术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',67,1644,'310610002','鼻填塞止血术, 前鼻腔','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',67,1645,'310610003','鼻后部填塞止血术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',67,1646,'310610004','鼻甲电凝止血术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',67,1647,'310610005','鼻烙止血术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',67,1648,'310610006','颈外动脉结扎术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',67,1649,'310610007','鼻部切开引流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',67,1650,'310610008','鼻活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual;
Insert Into 诊疗项目目录(类别,分类id,id,编码,名称,计算单位,标本部位,计算方式,单独应用,执行频率,适用性别,组合项目,操作类型,执行安排,执行科室,服务对象,计价性质,建档时间,撤档时间)
Select 'F',67,1651,'310610009','鼻腔活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',67,1652,'310610010','鼻病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',67,1653,'310610011','鼻腔内囊肿切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',67,1654,'310610012','鼻腔内肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',67,1655,'310610013','鼻腔血管瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',67,1656,'310610014','鼻息肉切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',67,1657,'310610015','鼻前庭囊肿切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',67,1658,'310610016','鼻死骨切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',67,1659,'310610017','鼻切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',67,1660,'310610018','鼻中膈粘膜下切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',67,1661,'310610019','鼻甲部分切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',67,1662,'310610020','鼻甲切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',67,1663,'310610021','鼻骨骨折闭合复位术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',67,1664,'310610022','鼻骨骨折开放性复位术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',67,1665,'310610023','鼻缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',67,1666,'310610024','鼻瘘修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',67,1667,'310610025','鼻重建术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',67,1668,'310610026','鼻畸形矫正术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',67,1669,'310610027','弯鼻鼻成形术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',67,1670,'310610028','鼻中膈矫正术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',67,1671,'310610029','鼻尖整形术(增高)','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',67,1672,'310610030','鼻翼矫正术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',67,1673,'310610031','鼻成形术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',67,1674,'310610032','鼻矫正术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',67,1675,'310610033','鼻中膈穿孔修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',67,1676,'310610034','断鼻再接术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',67,1677,'310610035','鼻腔粘连分离术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',67,1678,'310610036','鼻外伤清创术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',68,1679,'310620001','声门上肿块活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',68,1680,'310620002','咽部活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',68,1681,'310620003','鳃裂囊肿切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',68,1682,'310620004','鳃裂瘘切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',68,1683,'310620005','咽瘘修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',68,1684,'310620006','咽后壁修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',69,1685,'310630001','鼻窦穿剌抽吸术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',69,1686,'310630002','鼻窦穿剌灌洗术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',69,1687,'310630003','鼻窦活组织检查术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',69,1688,'310630004','上颌窦开窗术, 经鼻腔','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',69,1689,'310630005','上颌开窗术(单纯上颌窦切开术),经鼻腔','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',69,1690,'310630006','上颌窦引流术,经鼻腔','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',69,1691,'310630007','上颌窦根治术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',69,1692,'310630008','上颌窦开窗术, 外入路','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',69,1693,'310630009','上颌窦开窗术,外入路','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',69,1694,'310630010','上颌窦探查术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',69,1695,'310630011','额窦切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',69,1696,'310630012','额窦囊肿切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',69,1697,'310630013','额窦肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',69,1698,'310630014','鼻窦探查术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',69,1699,'310630015','鼻窦探查术, 经鼻','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',69,1700,'310630016','鼻窦探查术, 经外部','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',69,1701,'310630017','筛窦探查术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',69,1702,'310630018','蝶窦切开术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',69,1703,'310630019','蝶窦探查术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',69,1704,'310630020','鼻窦病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',69,1705,'310630021','鼻窦肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',69,1706,'310630022','上颌窦病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',69,1707,'310630023','上颌窦囊肿切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',69,1708,'310630024','上颌窦肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',69,1709,'310630025','筛窦囊肿切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',69,1710,'310630026','筛窦切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',69,1711,'310630027','筛窦肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',69,1712,'310630028','蝶窦囊肿切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',69,1713,'310630029','蝶窦肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',69,1714,'310630030','鼻窦瘘修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1715,'310640001','拔牙, 齿钳','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1716,'310640002','残留牙根拔除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1717,'310640003','阻生智齿拔出术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1718,'310640004','拔牙, 手术性','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1719,'310640005','牙填充修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1720,'310640006','牙嵌体修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1721,'310640007','人工齿冠','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1722,'310640008','牙植入术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1723,'310640009','义齿骨内植入','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1724,'310640010','牙槽切开术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1725,'310640011','牙龈活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1726,'310640012','牙槽活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1727,'310640013','牙龈成形术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1728,'310640014','牙龈肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1729,'310640015','牙周病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1730,'310640016','牙龈缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1731,'310640017','牙槽肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1732,'310640018','牙囊肿摘除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1733,'310640019','齿槽骨修整术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1734,'310640020','牙槽切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1735,'310640021','牙槽修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1736,'310640022','牙槽整形术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1737,'310640023','舌活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1738,'310640024','舌病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1739,'310640025','舌血管瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1740,'310640026','舌肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1741,'310640027','舌部分切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1742,'310640028','舌根治性切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1743,'310640029','舌缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1744,'310640030','舌系带切开术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1745,'310640031','舌系带切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1746,'310640032','唾液腺活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1747,'310640033','颌下腺结石切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1748,'310640034','颌下腺肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1749,'310640035','腮腺病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1750,'310640036','腮腺肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1751,'310640037','舌下腺囊肿切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1752,'310640038','涎腺病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1753,'310640039','面部的其他手术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1754,'310640040','颌下腺部分切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1755,'310640041','腮腺部分切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1756,'310640042','颌下腺切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1757,'310640043','腮腺全部切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1758,'310640044','舌下腺切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1759,'310640045','面部脓肿引流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1760,'310640046','悬雍垂活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual;
Insert Into 诊疗项目目录(类别,分类id,id,编码,名称,计算单位,标本部位,计算方式,单独应用,执行频率,适用性别,组合项目,操作类型,执行安排,执行科室,服务对象,计价性质,建档时间,撤档时间)
Select 'F',70,1761,'310640047','唇活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1762,'310640048','腭囊肿切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1763,'310640049','腭肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1764,'310640050','腭广泛切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1765,'310640051','唇病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1766,'310640052','唇血管瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1767,'310640053','唇肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1768,'310640054','颊内部病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1769,'310640055','口腔粘膜病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1770,'310640056','软腭肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1771,'310640057','舌下肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1772,'310640058','唇撕裂缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1773,'310640059','唇瘘管修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1774,'310640060','松果腺部分切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1775,'310640061','唇裂修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1776,'310640062','唇成形术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1777,'310640063','唇矫形术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1778,'310640064','巨口矩正术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1779,'310640065','口变形矫正术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1780,'310640066','口底重建术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1781,'310640067','腭裂修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1782,'310640068','软腭成形术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1783,'310640069','悬雍垂切开术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1784,'310640070','悬雍切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1785,'310640071','悬雍垂修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',70,1786,'310640072','悬雍垂肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',71,1787,'310650001','扁桃体脓肿切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',71,1788,'310650002','扁桃体周围脓肿引流','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',71,1789,'310650003','咽后脓肿引流','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',71,1790,'310650004','咽旁脓肿引流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',71,1791,'310650005','扁桃体切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',71,1792,'310650006','扁桃体伴腺样体切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',71,1793,'310650007','扁桃残体摘除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',71,1794,'310650008','舌扁桃体切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',71,1795,'310650009','增殖腺切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',71,1796,'310650010','腺样体切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',71,1797,'310650011','扁桃体病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',71,1798,'310650012','扁桃体肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',73,1799,'310710001','声带注射术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',73,1800,'310710002','气管造口术,暂时性','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',73,1801,'310710003','气管造口术,永久性','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',73,1802,'310710004','喉脓肿切开引流术, 经喉镜','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',73,1803,'310710005','喉切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',73,1804,'310710006','会厌脓肿切开引流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',73,1805,'310710007','喉活组织检查, 纤维喉镜','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',73,1806,'310710008','喉活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',73,1807,'310710009','气管活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',73,1808,'310710010','气管囊肿切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',73,1809,'310710011','气管肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',73,1810,'310710012','气管肿瘤切除术,经支气管','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',73,1811,'310710013','喉缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',73,1812,'310710014','喉瘘修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',73,1813,'310710015','喉成形术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',73,1814,'310710016','声带移位术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',73,1815,'310710017','气管造口闭合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',73,1816,'310710018','气管瘘闭合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',73,1817,'310710019','气管成形术伴人工喉','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',73,1818,'310710020','发音再造术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',73,1819,'310710021','喉模置换术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',73,1820,'310710022','喉模取出术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',74,1821,'310720001','支气管囊肿切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',74,1822,'310720002','支气管肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',74,1823,'310720003','肺大泡结扎术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',74,1824,'310720004','肺病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',74,1825,'310720005','肺楔形切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',74,1826,'310720006','肺血管瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',74,1827,'310720007','肺肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',74,1828,'310720008','肺叶切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',74,1829,'310720009','肺切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',74,1830,'310720010','支气管异物切开取出术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',74,1831,'310720011','肺脓肿引流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',74,1832,'310720012','支气管光学纤维镜检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',74,1833,'310720013','支气管镜检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',74,1834,'310720014','肺活组织检查, 经支气管镜','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',74,1835,'310720015','肺活组织检查, 经皮肤穿剌','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',74,1836,'310720016','肺活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',74,1837,'310720017','胸廓成形术,用于肺萎陷','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',74,1838,'310720018','食管支气管瘘修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',74,1839,'310720019','肺修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',74,1840,'310720020','肺移植术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',74,1841,'310720021','支气管扩张术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',74,1842,'310720022','支气管结扎术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',74,1843,'310720023','肺穿剌术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',75,1844,'310730001','胸膜外引流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',75,1845,'310730002','开胸探查术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',75,1846,'310730003','胸膜闭式引流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',75,1847,'310730004','开胸止血术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',75,1848,'310730005','胸膜开窗引流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',75,1849,'310730006','胸腔开放式引流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',75,1850,'310730007','胸腔内异物取出术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',75,1851,'310730008','纵隔探查术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',75,1852,'310730009','纵隔引流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',75,1853,'310730010','胸壁活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',75,1854,'310730011','胸膜活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',75,1855,'310730012','纵隔活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',75,1856,'310730013','横膈活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',75,1857,'310730014','纵隔血管瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',75,1858,'310730015','纵隔肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',75,1859,'310730016','胸壁病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',75,1860,'310730017','胸膜病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',75,1861,'310730018','胸壁撕裂缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',75,1862,'310730019','胸廓造口闭合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',75,1863,'310730020','支气管胸膜瘘管闭合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',75,1864,'310730021','支气管胸膜皮肤瘘管闭合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',75,1865,'310730022','支气管胸膜纵隔瘘管闭合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',75,1866,'310730023','胸廓畸形矫正术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',75,1867,'310730024','胸壁修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',75,1868,'310730025','横膈病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',75,1869,'310730026','横膈部分切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',75,1870,'310730027','横膈缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual;
Insert Into 诊疗项目目录(类别,分类id,id,编码,名称,计算单位,标本部位,计算方式,单独应用,执行频率,适用性别,组合项目,操作类型,执行安排,执行科室,服务对象,计价性质,建档时间,撤档时间)
Select 'F',75,1871,'310730028','胸肠瘘管切断术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',75,1872,'310730029','胸腹瘘管切断术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',75,1873,'310730030','胸胃瘘管切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',75,1874,'310730031','横膈修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',75,1875,'310730032','横膈脓肿引流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',75,1876,'310730033','胸腔穿刺术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',75,1877,'310730034','胸膜修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',76,1878,'310740001','喉囊肿造袋术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',76,1879,'310740002','喉蹼切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',76,1880,'310740003','喉肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',76,1881,'310740004','会厌病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',76,1882,'310740005','声带息肉切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',76,1883,'310740006','声带息肉切除术,经喉镜','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',76,1884,'310740007','声门病损切除术,经喉镜','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',76,1885,'310740008','半喉切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',76,1886,'310740009','会厌切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',76,1887,'310740010','声带切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',76,1888,'310740011','喉软骨切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',76,1889,'310740012','喉全切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',76,1890,'310740013','喉根治性切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',76,1891,'310740014','肺叶部分切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',78,1892,'310810001','心包穿剌术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',78,1893,'310810002','心肌切开术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',78,1894,'310810003','心内膜切开术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',78,1895,'310810004','心室切开术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',78,1896,'310810005','心包切开术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',78,1897,'310810006','心包松解术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',78,1898,'310810007','右心导管术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',78,1899,'310810008','左心导管术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',78,1900,'310810009','左右心联合导管术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',78,1901,'310810010','心包活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',78,1902,'310810011','心脏活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',78,1903,'310810012','心包部分切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',78,1904,'310810013','心包囊肿切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',78,1905,'310810014','心脏动脉瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',78,1906,'310810015','心脏肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',78,1907,'310810016','心脏移植术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',78,1908,'310810017','搏动性气囊植入术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',78,1909,'310810018','心脏泵植入术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',78,1910,'310810019','心室暂时性起搏器植入','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',78,1911,'310810020','心房暂时性起搏器植入术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',78,1912,'310810021','心脏暂时性起搏器植入','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',78,1913,'310810022','心房永久性起搏器植入,经静脉','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',78,1914,'310810023','心室永久性起搏器植入,经静脉','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',78,1915,'310810024','心脏永久性起搏器植入,经静脉','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',78,1916,'310810025','心脏永锥性起搏器植入','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',78,1917,'310810026','心脏起搏器电池置换术,经静脉','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',78,1918,'310810027','心外膜电极置换术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',78,1919,'310810028','心外膜电极去除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',78,1920,'310810029','心脏起搏器电池置换术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',78,1921,'310810030','暂时性心脏起搏器去除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',78,1922,'310810031','心脏开胸按摩','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',78,1923,'310810032','心内局部注射酒精','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1924,'310820001','动脉血栓切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1925,'310820002','静脉探查术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1926,'310820003','颈内动脉血栓切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1927,'310820004','主动脉血栓切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1928,'310820005','腹主动脉血栓切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1929,'310820006','肠系膜上动脉取栓术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1930,'310820007','静脉瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1931,'310820008','下肢动脉取栓术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1932,'310820009','下肢静脉切开术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1933,'310820010','下肢静脉取栓术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1934,'310820011','动脉内膜伴血栓切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1935,'310820012','动脉内膜剥脱术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1936,'310820013','血管活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1937,'310820014','动脉瘤切除伴吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1938,'310820015','血管切除伴吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1939,'310820016','腹主动脉瘤切除术伴人工血管移植术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1940,'310820017','动脉瘤切除伴有静脉移植','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1941,'310820018','动脉探查术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1942,'310820019','阴茎背静脉曲张结扎术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1943,'310820020','下腔静脉曲张切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1944,'310820021','大隐静脉剥脱术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1945,'310820022','大隐静脉高位结扎剥脱术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1946,'310820023','下肢曲张静脉剥脱术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1947,'310820024','下肢曲张静脉切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1948,'310820025','小隐静脉高位结扎剥术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1949,'310820026','移植血管切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1950,'310820027','动静脉瘘切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1951,'310820028','动脉瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1952,'310820029','假动脉瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1953,'310820030','心畸形血管切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1954,'310820031','血管病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1955,'310820032','颅内畸形血管切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1956,'310820033','颈动脉体瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1957,'310820034','颈静脉球瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1958,'310820035','下腔静脉病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1959,'310820036','下腔静脉隔膜切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1960,'310820037','腔静脉结扎术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1961,'310820038','椎动脉结扎术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1962,'310820039','颈内动脉结扎术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1963,'310820040','颈前静脉结扎术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1964,'310820041','颈总动脉结扎术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1965,'310820042','动脉导结扎术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1966,'310820043','胸壁血管结扎术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1967,'310820044','肠系膜下动脉结扎术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1968,'310820045','肝动脉结扎术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1969,'310820046','肝动脉栓塞术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1970,'310820047','脾动脉结扎术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1971,'310820048','肾动脉栓塞术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1972,'310820049','子宫动静脉高位结扎术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1973,'310820050','腹部静脉结扎术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1974,'310820051','胃底静脉结扎','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1975,'310820052','颈前动脉结扎术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1976,'310820053','髂内动脉结扎术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1977,'310820054','大隐静脉结扎术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1978,'310820055','下肢静脉结所术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1979,'310820056','动脉导管插入术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1980,'310820057','肝动脉插管术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual;
Insert Into 诊疗项目目录(类别,分类id,id,编码,名称,计算单位,标本部位,计算方式,单独应用,执行频率,适用性别,组合项目,操作类型,执行安排,执行科室,服务对象,计价性质,建档时间,撤档时间)
Select 'F',79,1981,'310820058','髂内动脉插管术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1982,'310820059','脐静脉导管插入术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1983,'310820060','肝静脉导管插入术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1984,'310820061','上腔静脉导管插入术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1985,'310820062','肾静脉导管插入术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',79,1986,'310820063','股动脉穿剌术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',80,1987,'310830001','降主动脉-肺动脉吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',80,1988,'310830002','锁骨下--肺动脉吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',80,1989,'310830003','肠系膜静脉-下腔静脉吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',80,1990,'310830004','门--腔静脉分流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',80,1991,'310830005','脾--肾静脉吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',80,1992,'310830006','腔静脉--肺动脉吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',80,1993,'310830007','主动脉-锁骨下动脉-肱动脉搭桥术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',80,1994,'310830008','主动脉-锁骨下动脉-颈动脉搭桥术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',80,1995,'310830009','颈外-颈内动脉人工血管架桥术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',80,1996,'310830010','颈总动脉-锁骨下动脉搭桥术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',80,1997,'310830011','颈总动脉-腋动脉人工血管架桥术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',80,1998,'310830012','锁骨下动脉--肱动脉吻合,大隐静脉架桥术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',80,1999,'310830013','胸腔内动脉搭桥术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',80,2000,'310830014','腹主动脉-肾动脉搭桥术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',80,2001,'310830015','腹主动脉-股动脉人工血管架桥术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',80,2002,'310830016','腹主动脉-髂动脉吻合, 人工血管Y移植术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',80,2003,'310830017','髂总动脉--股动脉人工血管架桥术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',80,2004,'310830018','腹主动脉-肠系膜上动脉人工血管架桥术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',80,2005,'310830019','动静脉造瘘术, 为肾透析','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',80,2006,'310830020','隐静脉分流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',80,2007,'310830021','左右大隐静脉吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',80,2008,'310830022','大隐静脉-股静脉吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',80,2009,'310830023','股腓动脉搭桥术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',80,2010,'310830024','股国动脉搭桥术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',80,2011,'310830025','股颈动脉搭桥术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',80,2012,'310830026','颈外颈内静脉吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',80,2013,'310830027','下肢动脉搭桥术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',80,2014,'310830028','下肢动脉人造血管架桥术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',80,2015,'310830029','腋-肱动脉搭桥术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',80,2016,'310830030','动脉缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',80,2017,'310830031','静脉缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',80,2018,'310830032','静脉扩张缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',80,2019,'310830033','血管手术后出血的止血','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',80,2020,'310830034','动脉瘤钳夹术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',80,2021,'310830035','动脉瘤破裂修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',80,2022,'310830036','动静脉修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',80,2023,'310830037','股动脉瘘结扎术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',80,2024,'310830038','动脉修补术, 组织补片移植','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',80,2025,'310830039','静脉修补术,组织补片移植','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',80,2026,'310830040','动脉修补术, 合成补片','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',80,2027,'310830041','静脉修补术,合成补片','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',80,2028,'310830042','动脉修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',80,2029,'310830043','颈内动脉成形术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',80,2030,'310830044','人工心肺,体外循环','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',80,2031,'310830045','体外循环','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',80,2032,'310830046','化学感受组织切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',80,2033,'310830047','颈动脉球切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',80,2034,'310830048','锁骨下静脉松解术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',80,2035,'310830049','人工肾','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',80,2036,'310830050','肾透析','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',80,2037,'310830051','血液透析','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',80,2038,'310830052','颈动脉扩张术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2039,'310840001','主动脉瓣闭式扩张术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2040,'310840002','二尖瓣闭式扩张术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2041,'310840003','肺动闭瓣闭式扩张术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2042,'310840004','三尖瓣闭式切开术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2043,'310840005','主动脉瓣修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2044,'310840006','二尖瓣闭锁不全修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2045,'310840007','二尖瓣成形术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2046,'310840008','二尖瓣缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2047,'310840009','二尖瓣切开扩张术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2048,'310840010','肺动脉瓣切开术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2049,'310840011','三尖瓣修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2050,'310840012','主动脉生物瓣膜置换术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2051,'310840013','主动脉瓣机械瓣膜置换术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2052,'310840014','二尖瓣生物瓣膜置换术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2053,'310840015','二尖瓣机械瓣膜置换术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2054,'310840016','肺动脉瓣生物瓣膜置换术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2055,'310840017','肺动脉瓣机械瓣膜置换术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2056,'310840018','三尖瓣生物瓣膜置换术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2057,'310840019','三尖瓣机械瓣膜置换术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2058,'310840020','心脏乳头肌切断术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2059,'310840021','心脏乳头肌修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2060,'310840022','心脏乳头肌再接术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2061,'310840023','心脏修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2062,'310840024','腱索切断术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2063,'310840025','腱索修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2064,'310840026','瓣环折迭术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2065,'310840027','右心室动脉圆锥切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2066,'310840028','主动脉瓣膜下环切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2067,'310840029','瓦耳萨瓦耳窦VALASALVA(动脉瘤)修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2068,'310840030','房间隔缺损假体修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2069,'310840031','室间隔缺损闭式假体修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2070,'310840032','室间隔缺损假体修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2071,'310840033','瓣膜缺损合并房室间膈缺损假体修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2072,'310840034','房室间隔缺损伴瓣膜缺损假体修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2073,'310840035','房室通道假体修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2074,'310840036','心内膜缺损假体修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2075,'310840037','房间隔成形伴组织移植','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2076,'310840038','房间隔缺损补片修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2077,'310840039','卵圆孔未闭组织移植修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2078,'310840040','室间隔缺损补片修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2079,'310840041','房室间隔缺损伴瓣膜缺损补片修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2080,'310840042','房室通道补片修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2081,'310840043','房间隔修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2082,'310840044','室间隔修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2083,'310840045','房室通道修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2084,'310840046','法乐氏四联症全部矫正伴肺动脉瓣联合切开术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2085,'310840047','法乐氏四联症全部修补术伴动脉贺锥切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2086,'310840048','法乐氏四联症全部修补术伴流出道补片修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2087,'310840049','法乐氏四联症全部修补术伴流出道修复术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2088,'310840050','法乐氏四症全部修补术伴室间隔缺损修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2089,'310840051','法乐氏四联症一期全部矫正术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2090,'310840052','肺静脉异常全部矫正术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual;
Insert Into 诊疗项目目录(类别,分类id,id,编码,名称,计算单位,标本部位,计算方式,单独应用,执行频率,适用性别,组合项目,操作类型,执行安排,执行科室,服务对象,计价性质,建档时间,撤档时间)
Select 'F',81,2091,'310840053','肺静脉异常全部修补术伴房间隔缺损修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2092,'310840054','肺静脉异常全部修补术伴肺总干和左房壁吻合','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2093,'310840055','肺静脉异常全部修补伴卵圆孔扩大','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2094,'310840056','动脉干全部矫正术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2095,'310840057','动脉干全部矫正术伴室间缺损修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2096,'310840058','动脉干全部矫正术伴右室代替肺动脉供血建造','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2097,'310840059','动脉干全部矫正术伴主动脉和肺动脉连接处结','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2098,'310840060','右心室肺动脉分流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2099,'310840061','左心室尖主动脉分流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2100,'310840062','方坦FONTAN氏手术(右心肺动脉带瓣管道转流','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2101,'310840063','心脏瓣假体重新缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2102,'310840064','心脏间隔假体重新缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',81,2103,'310840065','心瓣膜气囊成形术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',82,2104,'310850001','冠状动脉血管成形术, 经皮经管腔[PTCA]','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',82,2105,'310850002','冠状动脉血管成形术, [PTCA]伴血栓溶解剂注','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',82,2106,'310850003','冠状动脉梗阻直视去除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',82,2107,'310850004','冠状动脉内膜切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',82,2108,'310850005','冠状动脉内膜切除术伴补片移植术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',82,2109,'310850006','冠状动脉内血栓溶解剂注射','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',82,2110,'310850007','主动脉-冠状动脉架桥术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',82,2111,'310850008','主动脉-冠状动脉架桥术,一根冠状动脉','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',82,2112,'310850009','主动脉-冠状动脉架桥术,二根冠状动脉','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',82,2113,'310850010','主动脉-冠状动脉架桥术,三根冠状动脉','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',82,2114,'310850011','主动脉-冠状动脉架桥术,四根或更多根冠状动','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',82,2115,'310850012','乳房内动脉冠状动脉吻合术,单个','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',82,2116,'310850013','胸动脉--冠状动脉吻合术,单个','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',82,2117,'310850014','乳房内动脉冠状动脉吻合术,双个','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',82,2118,'310850015','胸动脉--冠状动脉吻合术,双个','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',82,2119,'310850016','冠状血管动脉瘤修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',82,2120,'310850017','冠状动脉结扎术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',82,2121,'310850018','冠状动脉探查术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',84,2122,'310910001','淋巴管探查术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',84,2123,'310910002','淋巴活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',84,2124,'310910003','颈深部淋巴结切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',84,2125,'310910004','乳腺淋巴结切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',84,2126,'310910005','腋下淋巴结切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',84,2127,'310910006','腹股沟淋巴结切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',84,2128,'310910007','淋巴管瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',84,2129,'310910008','淋巴管囊肿切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',84,2130,'310910009','淋巴结切除术,扩大性区域性','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',84,2131,'310910010','区域性淋巴结切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',84,2132,'310910011','颈淋巴结清除术,单调','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',84,2133,'310910012','颈淋巴结清除术,双调','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',84,2134,'310910013','腋下淋巴结清扫术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',84,2135,'310910014','主动脉旁根治性淋巴结切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',84,2136,'310910015','髂淋巴结根治性切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',84,2137,'310910016','腹股沟淋巴结清扫术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',84,2138,'310910017','腹腔淋巴清扫术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',84,2139,'310910018','胸导管插入术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',84,2140,'310910019','胸导管造兼术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',84,2141,'310910020','胸导管瘘关闭术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',84,2142,'310910021','胸导管结扎术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',84,2143,'310910022','胸导管--颈内静脉吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',84,2144,'310910023','周围淋巴管闭合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',84,2145,'310910024','周围淋巴管重建术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',84,2146,'310910025','周围淋管结扎术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',84,2147,'310910026','周围淋巴管扩张术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',84,2148,'310910027','周围淋巴管吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',84,2149,'310910028','周围淋巴管修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',84,2150,'310910029','周围淋巴管移植术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',84,2151,'310910030','下肢淋巴管-静脉分流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',85,2152,'310920001','骨髓移植术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',85,2153,'310920002','脾穿刺','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',85,2154,'310920003','脾切开术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',85,2155,'310920004','骨髓活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',85,2156,'310920005','脾穿刺活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',85,2157,'310920006','脾活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',85,2158,'310920007','脾囊肿造袋术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',85,2159,'310920008','脾病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',85,2160,'310920009','脾部分切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',85,2161,'310920010','副脾切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',85,2162,'310920011','脾移植术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',85,2163,'310920012','脾组织移植术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',85,2164,'310920013','脾修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',85,2165,'310920014','脾切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',87,2166,'311010001','食管蹼膜切开术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',87,2167,'311010002','食管内异物切开取出术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',87,2168,'311010003','颈部食管造口术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',87,2169,'311010004','胸部食管造口术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',87,2170,'311010005','食管镜检查,食管切开','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',87,2171,'311010006','食管镜检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',87,2172,'311010007','食管活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',87,2173,'311010008','食管憩室切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',87,2174,'311010009','食管病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',87,2175,'311010010','食管囊肿切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',87,2176,'311010011','食管肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',87,2177,'311010012','食管部分切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',87,2178,'311010013','食管切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',87,2179,'311010014','胃大部切除伴食管-胃吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',87,2180,'311010015','胃近端切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',87,2181,'311010016','食管--食管吻合术,胸腔内','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',87,2182,'311010017','食管--胃吻合术,胸腔内','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',87,2183,'311010018','食管--胃转流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',87,2184,'311010019','空肠食道间置术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',87,2185,'311010020','小肠食管间置术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',87,2186,'311010021','结肠代食道吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',87,2187,'311010022','食管--食管吻合术,胸骨前','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',87,2188,'311010023','食管--胃吻合术,胸骨前','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',87,2189,'311010024','小肠食管间置术,胸骨前','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',87,2190,'311010025','结肠食管间置术, 胸骨前','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',87,2191,'311010026','改良HELLER手术(食管肌切开术)','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',87,2192,'311010027','海伦HELLER手术(食管肌切开术)','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',87,2193,'311010028','食管肌切开术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',87,2194,'311010029','食管永久性插管植入术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',87,2195,'311010030','食管撕裂缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',87,2196,'311010031','食管造口关闭术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',87,2197,'311010032','食管狭窄修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',87,2198,'311010033','食管修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',87,2199,'311010034','食管静脉结扎术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',87,2200,'311010035','食管扩张术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual;
Insert Into 诊疗项目目录(类别,分类id,id,编码,名称,计算单位,标本部位,计算方式,单独应用,执行频率,适用性别,组合项目,操作类型,执行安排,执行科室,服务对象,计价性质,建档时间,撤档时间)
Select 'F',88,2201,'311020001','胃切开异物取出术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',88,2202,'311020002','胃造瘘术,暂时性','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',88,2203,'311020003','胃造瘘术,永久性','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',88,2204,'311020004','幽门肌切开术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',88,2205,'311020005','胃息肉切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',88,2206,'311020006','胃病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',88,2207,'311020007','胃十二指肠切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',88,2208,'311020008','胃癌根治术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',88,2209,'311020009','胃切除伴食管空肠吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',88,2210,'311020010','胃切除伴十二指肠吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',88,2211,'311020011','胃切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',88,2212,'311020012','迷走神经切断术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',88,2213,'311020013','胃活组织检查,经胃镜','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',88,2214,'311020014','胃镜检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',88,2215,'311020015','胃刷洗活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',88,2216,'311020016','幽门切开括张术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',88,2217,'311020017','气囊内窥镜幽门扩张术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',88,2218,'311020018','胃空肠吻合口扩张,内窥镜下','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',88,2219,'311020019','幽门成形术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',88,2220,'311020020','幽门梗阻松解术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',88,2221,'311020021','胃空肠RONX-Y吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',88,2222,'311020022','胃溃疡部位缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',88,2223,'311020023','十二指肠溃疡部分缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',88,2224,'311020024','胃空肠吻合口闭合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',88,2225,'311020025','胃十二指肠吻合口闭合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',88,2226,'311020026','胃吻合口闭合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',88,2227,'311020027','胃大部切除伴胃十二指肠吻合术,I式','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',88,2228,'311020028','胃大部切除伴胃十二指肠吻合术(毕罗特手术I','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',88,2229,'311020029','胃末端切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',88,2230,'311020030','胃幽门切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',88,2231,'311020031','胃裂伤缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',88,2232,'311020032','胃造瘘关闭术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',88,2233,'311020033','胃固定术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',88,2234,'311020034','贲门成形术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',88,2235,'311020035','食管和胃贲门成形术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',88,2236,'311020036','胃贲门成形术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',88,2237,'311020037','胃大部切除伴胃十二指肠吻合术,II式','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',88,2238,'311020038','残胃部分切除胃空肠吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',88,2239,'311020039','胃大部切除伴胃空肠吻合(毕罗特手术II式)','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',88,2240,'311020040','胃部分切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',88,2241,'311020041','胃静脉曲张结扎术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',89,2242,'311030001','肠切开异物取出术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',89,2243,'311030002','十二指肠探查术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',89,2244,'311030003','小肠内窥镜检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',89,2245,'311030004','小肠刷洗活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',89,2246,'311030005','纤维结肠镜检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',89,2247,'311030006','乙状结肠镜检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',89,2248,'311030007','结肠刷洗活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',89,2249,'311030008','大肠活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',89,2250,'311030009','十二指肠憩室切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',89,2251,'311030010','十二指肠肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',89,2252,'311030011','小肠病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',89,2253,'311030012','小肠憩室切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',89,2254,'311030013','小肠肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',89,2255,'311030014','乙状结肠肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',89,2256,'311030015','大肠憩室切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',89,2257,'311030016','结肠息肉切除术, 经纤维结肠镜','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',89,2258,'311030017','结肠肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',89,2259,'311030018','盲肠肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',89,2260,'311030019','小肠部分切除术为间置术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',89,2261,'311030020','大肠部分切除术为间置术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',89,2262,'311030021','回肠部分切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',89,2263,'311030022','回肠切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',89,2264,'311030023','空肠部分切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',89,2265,'311030024','十二指肠部分切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',89,2266,'311030025','小肠部分切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',89,2267,'311030026','小肠全部切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',89,2268,'311030027','盲肠切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',89,2269,'311030028','右半结肠切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',89,2270,'311030029','回肠结肠切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',89,2271,'311030030','横结肠切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',89,2272,'311030031','横结肠部分切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',89,2273,'311030032','左半结肠切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',89,2274,'311030033','乙状结肠部分切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',89,2275,'311030034','乙状结肠切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',89,2276,'311030035','降结肠切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',89,2277,'311030036','大肠部分切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',89,2278,'311030037','结肠部分切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',89,2279,'311030038','结肠大部切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',89,2280,'311030039','升结肠部分切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',89,2281,'311030040','升结肠切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',89,2282,'311030041','大肠全切术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',89,2283,'311030042','结肠全部切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',89,2284,'311030043','空肠--空肠吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',89,2285,'311030044','十二指肠--空肠吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',89,2286,'311030045','小肠--小肠吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',89,2287,'311030046','回肠-横结肠吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',89,2288,'311030047','回肠-乙状结肠吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',89,2289,'311030048','小肠--大肠吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',89,2290,'311030049','乙状结肠-乙状结肠吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',89,2291,'311030050','乙状结肠-直肠吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',89,2292,'311030051','直肠-乙状结肠吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',89,2293,'311030052','横结肠-乙状结肠吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',89,2294,'311030053','降结肠-乙状结肠吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',89,2295,'311030054','降结肠-直肠吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',89,2296,'311030055','结肠-直肠吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',89,2297,'311030056','脾曲结肠--乙状结肠侧侧吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',89,2298,'311030057','升结肠--降结肠吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',90,2299,'311040001','十二指肠旷置术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',90,2300,'311040002','小肠外置术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',90,2301,'311040003','大肠外置术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',90,2302,'311040004','结肠置管引流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',90,2303,'311040005','乙状结肠破口造口术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',90,2304,'311040006','乙状结肠造口术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',90,2305,'311040007','横结肠造口术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',90,2306,'311040008','降结肠造口术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',90,2307,'311040009','结肠造口术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',90,2308,'311040010','盲肠造口术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',90,2309,'311040011','结肠暂时造口术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',90,2310,'311040012','结肠永久性造口术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual;
Insert Into 诊疗项目目录(类别,分类id,id,编码,名称,计算单位,标本部位,计算方式,单独应用,执行频率,适用性别,组合项目,操作类型,执行安排,执行科室,服务对象,计价性质,建档时间,撤档时间)
Select 'F',90,2311,'311040013','回肠造口术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',90,2312,'311040014','回肠暂时性造口术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',90,2313,'311040015','回肠永久性造口术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',90,2314,'311040016','空肠造口术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',90,2315,'311040017','十二指肠造口术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',90,2316,'311040018','肠造瘘口修正术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',90,2317,'311040019','小肠造瘘口修正术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',90,2318,'311040020','大肠造瘘口修正术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',90,2319,'311040021','肠造口关闭术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',90,2320,'311040022','小肠造瘘口闭合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',90,2321,'311040023','乙状结肠造瘘口闭合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',90,2322,'311040024','结肠造口闭合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',90,2323,'311040025','盲肠造瘘口闭合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',90,2324,'311040026','空肠折迭术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',90,2325,'311040027','诺布尔NOBLE氏小肠折迭术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',90,2326,'311040028','乙状结肠固定术(莫斯科茨氏)','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',90,2327,'311040029','盲肠升结肠固定术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',90,2328,'311040030','莫斯科茨氏术(乙状结肠固定术)','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',90,2329,'311040031','结肠固定术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',90,2330,'311040032','盲肠固定术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',90,2331,'311040033','十二指肠修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',90,2332,'311040034','乙状结肠瘘修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',90,2333,'311040035','肠穿孔修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',90,2334,'311040036','肠瘘关闭术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',90,2335,'311040037','肠憩室缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',90,2336,'311040038','肠修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',90,2337,'311040039','结肠修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',90,2338,'311040040','十二指肠憩室修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',90,2339,'311040041','小肠修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',90,2340,'311040042','肠扭转复位术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',90,2341,'311040043','肠套迭复位术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',90,2342,'311040044','乙状结肠肌切开术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',90,2343,'311040045','阑尾脓肿引流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',90,2344,'311040046','阑尾切除伴引流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',90,2345,'311040047','阑尾切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',90,2346,'311040048','阑尾切除术,附带的','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',90,2347,'311040049','阑尾内翻包埋术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',91,2348,'311050001','直肠切开术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',91,2349,'311050002','直肠造口术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',91,2350,'311050003','直肠乙状结肠镜检查,经腹','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',91,2351,'311050004','直肠乙状结肠镜检查,经人工瘘口','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',91,2352,'311050005','直肠乙状结肠镜检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',91,2353,'311050006','直肠刷洗活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',91,2354,'311050007','直肠活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',91,2355,'311050008','直肠肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',91,2356,'311050009','直肠病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',91,2357,'311050010','直肠囊肿切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',91,2358,'311050011','直肠内膜拖出切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',91,2359,'311050012','阿尔特迈氏ALTEMEIER手术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',91,2360,'311050013','斯温森氏SWENSON直肠切断术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',91,2361,'311050014','直肠腹会阴拖出切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',91,2362,'311050015','腹, 会阴, 直肠联合切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',91,2363,'311050016','迈乐斯MILES氏术(腹,会阴,直肠联合切除)','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',91,2364,'311050017','MILES迈尔斯氏术(腹.会阴.直肠联合切除)','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',91,2365,'311050018','直肠全部切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',91,2366,'311050019','直肠乙状结肠切除术,经骶','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',91,2367,'311050020','直肠乙状结肠部分切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',91,2368,'311050021','直肠部分切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',91,2369,'311050022','直肠裂伤缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',91,2370,'311050023','直肠造口闭合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',91,2371,'311050024','直肠瘘修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',91,2372,'311050025','直肠会阴瘘修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',91,2373,'311050026','弗里克曼氏FRICKMAN手术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',91,2374,'311050027','直肠脱垂里普斯坦氏RIPSTEIN修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',91,2375,'311050028','直肠脱垂德洛姆氏DELORME修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',91,2376,'311050029','直肠修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',91,2377,'311050030','直肠阴道隔膜切开术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',91,2378,'311050031','直肠阴道隔病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',91,2379,'311050032','直肠阴道隔肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',91,2380,'311050033','直肠狭窄切开术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',91,2381,'311050034','肛门直肠肌部分切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',91,2382,'311050035','直肠周围瘘管修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',91,2383,'311050036','肛周脓肿切开引流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',91,2384,'311050037','肛旁病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',91,2385,'311050038','肛瘘切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',91,2386,'311050039','肛门镜检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',91,2387,'311050040','肛周围组织活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',91,2388,'311050041','肛门活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',91,2389,'311050042','肛门裂切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',91,2390,'311050043','肛门息肉切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',91,2391,'311050044','肛门肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',91,2392,'311050045','肛乳头切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',91,2393,'311050046','痔结扎术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',91,2394,'311050047','痔切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',91,2395,'311050048','肛门切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',91,2396,'311050049','肛门裂伤缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',91,2397,'311050050','肛门环扎术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',91,2398,'311050051','肛门瘘挂线结扎术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',91,2399,'311050052','肛门瘘管闭合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',91,2400,'311050053','肛门成形术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',91,2401,'311050054','肛门隔膜切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',91,2402,'311050055','肛门脱垂复位','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2403,'311060001','肝脓肿引流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2404,'311060002','肝穿剌活组织检查, 经皮','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2405,'311060003','肝活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2406,'311060004','肝部分切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2407,'311060005','肝包囊虫切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2408,'311060006','肝病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2409,'311060007','肝囊肿切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2410,'311060008','肝血管瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2411,'311060009','肝肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2412,'311060010','肝叶切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2413,'311060011','肝缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2414,'311060012','肝抽吸术, 经皮','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2415,'311060013','肝止血术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2416,'311060014','胆囊造口术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2417,'311060015','胆囊切开取石术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2418,'311060016','胆囊引流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2419,'311060017','胆道内窥镜检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2420,'311060018','胆囊部分切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual;
Insert Into 诊疗项目目录(类别,分类id,id,编码,名称,计算单位,标本部位,计算方式,单独应用,执行频率,适用性别,组合项目,操作类型,执行安排,执行科室,服务对象,计价性质,建档时间,撤档时间)
Select 'F',92,2421,'311060019','残余胆囊切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2422,'311060020','胆囊切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2423,'311060021','胆囊肝管吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2424,'311060022','胆囊-空肠吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2425,'311060023','胆囊-十二指肠吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2426,'311060024','胆囊胰腺吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2427,'311060025','胆囊胃吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2428,'311060026','胆总管-空肠吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2429,'311060027','胆总管-十二指肠吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2430,'311060028','肝管-空肠吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2431,'311060029','胆管-十二指肠吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2432,'311060030','胆管-胃吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2433,'311060031','胆总管切开取石术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2434,'311060032','胆总管切开取蛔虫','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2435,'311060033','胆管切开取石术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2436,'311060034','胆总管切开T管引流','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2437,'311060035','胆总管切开引流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2438,'311060036','胆总管探查术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2439,'311060037','胆管引流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2440,'311060038','残余胆囊管切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2441,'311060039','法特氏壶腹切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2442,'311060040','总胆管切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2443,'311060041','胆管病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2444,'311060042','胆管囊肿切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2445,'311060043','胆管肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2446,'311060044','肝胆管切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2447,'311060045','肝管病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2448,'311060046','肝管切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2449,'311060047','肝总管切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2450,'311060048','总胆管缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2451,'311060049','胆总管瘘修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2452,'311060050','奥狄氏括约肌扩张术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2453,'311060051','法特氏壶腹扩张术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2454,'311060052','奥狄氏括约肌切开术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2455,'311060053','壶腹括约肌切开术, 经十二指','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2456,'311060054','胰括约肌切开术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2457,'311060055','奥狄氏括约肌成形术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2458,'311060056','胆囊修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2459,'311060057','胆囊造口闭合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2460,'311060058','胆囊空肠瘘切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2461,'311060059','胆囊-十二指肠瘘修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2462,'311060060','胆囊胃瘘修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',92,2463,'311060061','胆道结石去除术, 经内窥镜','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',93,2464,'311070001','胰腺囊肿引流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',93,2465,'311070002','胰腺脓肿引流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',93,2466,'311070003','胰腺切开取石术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',93,2467,'311070004','胰腺切开术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',93,2468,'311070005','胰腺探查术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',93,2469,'311070006','胰腺穿剌活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',93,2470,'311070007','胰腺活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',93,2471,'311070008','胰腺病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',93,2472,'311070009','胰腺胂瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',93,2473,'311070010','胰腺囊肿造袋术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',93,2474,'311070011','胰腺囊肿-空肠吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',93,2475,'311070012','胰腺囊肿胃吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',93,2476,'311070013','胰腺囊肿十二指肠吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',93,2477,'311070014','胰腺囊肿造口术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',93,2478,'311070015','胰腺囊肿-空肠R-Y内引流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',93,2479,'311070016','胰近端切除伴十二指肠切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',93,2480,'311070017','胰头伴部分胰体切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',93,2481,'311070018','胰头切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',93,2482,'311070019','胰尾切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',93,2483,'311070020','胰尾伴部分胰体切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',93,2484,'311070021','胰腺次根治术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',93,2485,'311070022','胰腺根治性次全切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',93,2486,'311070023','胰腺部分切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',93,2487,'311070024','胰十二指肠切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',93,2488,'311070025','胰腺全部切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',93,2489,'311070026','惠普尔WHIPPLE氏术(根治性胰十二指肠切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',93,2490,'311070027','胰十二指肠根治性切除术(惠普尔氏WHIPPLE术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',93,2491,'311070028','胰腺同种移植术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',93,2492,'311070029','胰腺异种移植术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',93,2493,'311070030','逆行胰管内窥镜检查(ERCP)','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',93,2494,'311070031','胰管套管置换术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',93,2495,'311070032','维尔松氏WIRSUNG管扩张术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',93,2496,'311070033','维尔松氏WIRSUNG管修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',93,2497,'311070034','胰瘘管切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',93,2498,'311070035','胰腺缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',93,2499,'311070036','胰腺-空肠吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',93,2500,'311070037','胰腺胃吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',94,2501,'311080001','巴西尼氏术(腹股沟疝修补术)','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',94,2502,'311080002','腹股沟滑动疝修补术, 单侧','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',94,2503,'311080003','腹股沟疝修补术, 单侧','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',94,2504,'311080004','腹股沟直疝修补术, 单侧','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',94,2505,'311080005','腹股沟斜疝修补术, 单侧','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',94,2506,'311080006','腹股沟直疝补片修补术,单侧','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',94,2507,'311080007','腹股沟斜疝补片修补术,单侧','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',94,2508,'311080008','腹股沟疝补片修补术, 单侧','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',94,2509,'311080009','腹股沟疝修补术, 双侧','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',94,2510,'311080010','腹股沟直疝修补术, 双侧','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',94,2511,'311080011','腹股沟斜疝修补术, 双侧','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',94,2512,'311080012','腹股沟疝修补术, 一侧直疝一侧斜疝','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',94,2513,'311080013','腹股沟直疝补片修补术,双侧','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',94,2514,'311080014','腹股沟斜疝补片修补术,双侧','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',94,2515,'311080015','腹股沟疝补片修补术, 一侧直疝一侧斜疝','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',94,2516,'311080016','腹股沟疝补征修补术, 双侧','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',94,2517,'311080017','股疝补片修补术, 单侧','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',94,2518,'311080018','股疝修补术, 单侧','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',94,2519,'311080019','股疝补片修补术, 双侧','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',94,2520,'311080020','股疝修补术, 双侧','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',94,2521,'311080021','脐疝假体修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',94,2522,'311080022','脐疝修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',94,2523,'311080023','腹壁切口疝修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',94,2524,'311080024','腹壁疝修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',94,2525,'311080025','腹壁切口疝补片或假体修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',94,2526,'311080026','腹壁疝假体修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',94,2527,'311080027','膈疝修补术, 经腹','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',94,2528,'311080028','食管裂孔疝修补术,经腹','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',94,2529,'311080029','膈疝修补术, 经胸','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',94,2530,'311080030','坐骨孔疝修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual;
Insert Into 诊疗项目目录(类别,分类id,id,编码,名称,计算单位,标本部位,计算方式,单独应用,执行频率,适用性别,组合项目,操作类型,执行安排,执行科室,服务对象,计价性质,建档时间,撤档时间)
Select 'F',94,2531,'311080031','坐骨直肠窝疝修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',94,2532,'311080032','闭孔疝修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',94,2533,'311080033','肠疝还纳术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',94,2534,'311080034','腹膜后疝修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',94,2535,'311080035','网膜疝修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',94,2536,'311080036','腰疝修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',95,2537,'311090001','腹壁脓肿切开引流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',95,2538,'311090002','腹股沟探查术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',95,2539,'311090003','腹膜后脓肿引流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',95,2540,'311090004','腹膜外脓肿引流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',95,2541,'311090005','剖腹探查术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',95,2542,'311090006','腹壁血肿清除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',95,2543,'311090007','腹腔引流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',95,2544,'311090008','膈下脓肿切开引流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',95,2545,'311090009','盆腔脓肿切开引流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',95,2546,'311090010','腹腔镜检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',95,2547,'311090011','腹膜活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',95,2548,'311090012','腹壁病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',95,2549,'311090013','腹壁肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',95,2550,'311090014','腹股沟病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',95,2551,'311090015','腹股沟肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',95,2552,'311090016','盆腔壁病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',95,2553,'311090017','脐病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',95,2554,'311090018','脐切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',95,2555,'311090019','肠系膜病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',95,2556,'311090020','肠系膜囊肿切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',95,2557,'311090021','肠系膜肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',95,2558,'311090022','大网膜病扣切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',95,2559,'311090023','大网膜肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',95,2560,'311090024','腹膜后肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',95,2561,'311090025','子宫粘连松解术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',95,2562,'311090026','肠粘连分离术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',95,2563,'311090027','腹膜粘连松解术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',95,2564,'311090028','盆腔粘连分离术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',95,2565,'311090029','胃十二指肠肝粘连松解术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',95,2566,'311090030','腹壁切口裂开缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',95,2567,'311090031','腹膜缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',95,2568,'311090032','腹壁加强修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',95,2569,'311090033','大网膜包肝术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',95,2570,'311090034','大网膜包肾术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',95,2571,'311090035','网膜缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',95,2572,'311090036','网膜固定术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',95,2573,'311090037','网膜扭转复位术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',95,2574,'311090038','网膜修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',95,2575,'311090039','网膜移植术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',95,2576,'311090040','肠系膜固定术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',95,2577,'311090041','肠系膜修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',95,2578,'311090042','腹腔穿剌术, 经皮','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',95,2579,'311090043','腹腔静脉分流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',95,2580,'311090044','腹膜切开术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',95,2581,'311090045','拉德氏LADD手术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',95,2582,'311090046','腹膜透析','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',95,2583,'311090047','腹腔病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',95,2584,'311090048','腹腔肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',95,2585,'311090049','盆腔病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',95,2586,'311090050','盆腔肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',97,2587,'311110001','肾切开取石术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',97,2588,'311110002','肾探查术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',97,2589,'311110003','肾盂切开取石术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',97,2590,'311110004','肾造瘘术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',97,2591,'311110005','肾盂探查术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',97,2592,'311110006','肾盂内T管引流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',97,2593,'311110007','肾盂造口术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',97,2594,'311110008','肾活组织检查,经皮肤针吸','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',97,2595,'311110009','肾活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',97,2596,'311110010','肾病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',97,2597,'311110011','肾囊肿切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',97,2598,'311110012','肾肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',97,2599,'311110013','肾部分切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',97,2600,'311110014','肾盏切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',97,2601,'311110015','肾切除术,单侧','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',97,2602,'311110016','肾输尿管切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',97,2603,'311110017','移植肾切除术(切除移植的肾)','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',97,2604,'311110018','肾自体移植术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',97,2605,'311110019','肾异体移植术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',97,2606,'311110020','肾固定术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',97,2607,'311110021','肾盂造口闭合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',97,2608,'311110022','肾造口关闭术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',97,2609,'311110023','肾蒂扭转复位术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',97,2610,'311110024','肾盂-输尿管-膀胱吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',97,2611,'311110025','肾盂--输尿管吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',97,2612,'311110026','肾盏--输尿管吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',97,2613,'311110027','肾盂成形术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',97,2614,'311110028','肾盂输尿管成形术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',97,2615,'311110029','肾修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',97,2616,'311110030','肾包膜剥除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',97,2617,'311110031','肾穿刺术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',97,2618,'311110032','肾囊肿抽吸术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',97,2619,'311110033','肾造瘘管置换术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',97,2620,'311110034','肾盂造瘘管置换术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',98,2621,'311120001','输尿管取石术,经输尿管镜','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',98,2622,'311120002','输尿管切开取石术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',98,2623,'311120003','输尿管探查术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',98,2624,'311120004','输尿管镜检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',98,2625,'311120005','输尿管切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',98,2626,'311120006','输尿管病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',98,2627,'311120007','输尿管部分切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',98,2628,'311120008','输尿管囊肿切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',98,2629,'311120009','输尿管全部切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',98,2630,'311120010','回肠代尿管建造术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',98,2631,'311120011','腹壁-输尿管吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',98,2632,'311120012','皮肤--输尿管造口术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',98,2633,'311120013','输尿管皮肤吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',98,2634,'311120014','输尿管--回肠吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',98,2635,'311120015','输尿管--乙状结肠吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',98,2636,'311120016','输尿管--直肠吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',98,2637,'311120017','输尿管肠管吻合口修正术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',98,2638,'311120018','输尿管--膀胱吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',98,2639,'311120019','输尿管--输尿管吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',98,2640,'311120020','输尿管管腔内粘连松解术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual;
Insert Into 诊疗项目目录(类别,分类id,id,编码,名称,计算单位,标本部位,计算方式,单独应用,执行频率,适用性别,组合项目,操作类型,执行安排,执行科室,服务对象,计价性质,建档时间,撤档时间)
Select 'F',98,2641,'311120021','输尿管裂伤缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',98,2642,'311120022','输尿管造口闭合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',98,2643,'311120023','输尿管瘘修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',98,2644,'311120024','输尿管阴道瘘修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',98,2645,'311120025','输尿管固定术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',98,2646,'311120026','输尿管成形术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',98,2647,'311120027','输尿管移植术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',98,2648,'311120028','输尿管口扩张术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',99,2649,'311130001','膀胱镜碎石术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',99,2650,'311130002','膀胱穿刺抽吸术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',99,2651,'311130003','膀胱切开异物取出术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',99,2652,'311130004','膀胱探查术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',99,2653,'311130005','耻骨上膀胱造口术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',99,2654,'311130006','膀胱造口术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',99,2655,'311130007','膀胱镜检查,经人工造口','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',99,2656,'311130008','膀胱镜检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',99,2657,'311130009','膀胱活组织检查,经尿道','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',99,2658,'311130010','膀胱活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',99,2659,'311130011','膀胱病损切除术,经尿道','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',99,2660,'311130012','膀胱肿瘤切除术,经尿道','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',99,2661,'311130013','膀胱取石术，经尿道','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',99,2662,'311130014','膀胱病损激光治疗','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',99,2663,'311130015','膀胱病损切开电灼术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',99,2664,'311130016','膀胱颈V形切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',99,2665,'311130017','膀胱部分切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',99,2666,'311130018','膀胱穹隆切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',99,2667,'311130019','膀胱三角区切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',99,2668,'311130020','膀胱根治性切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',99,2669,'311130021','膀胱全部切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',99,2670,'311130022','膀胱缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',99,2671,'311130023','膀胱造瘘口闭合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',99,2672,'311130024','膀胱乙状结肠瘘闭合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',99,2673,'311130025','膀胱瘘管闭合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',99,2674,'311130026','膀胱阴道瘘修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',99,2675,'311130027','乙状结肠代膀胱','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',99,2676,'311130028','直肠代膀胱术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',99,2677,'311130029','回肠代膀胱术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',99,2678,'311130030','膀胱修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',99,2679,'311130031','膀胱颈切开术,经尿道','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',99,2680,'311130032','耻骨上膀胱切开, 膀胱颈扩开术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',99,2681,'311130033','膀胱颈扩开术,耻骨上膀胱切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',99,2682,'311130034','导尿管插入','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',99,2683,'311130035','导尿管置换术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',100,2684,'311140001','尿道会阴造瘘术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',100,2685,'311140002','尿道切开术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',100,2686,'311140003','尿道造口术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',100,2687,'311140004','尿道外口切开术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',100,2688,'311140005','尿道周围组织活检','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',100,2689,'311140006','尿道瓣膜切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',100,2690,'311140007','尿道病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',100,2691,'311140008','尿道部分切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',100,2692,'311140009','尿道息肉电切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',100,2693,'311140010','尿道狭窄电切术,经尿道','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',100,2694,'311140011','尿道狭窄电切术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',100,2695,'311140012','尿道肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',100,2696,'311140013','尿道裂伤缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',100,2697,'311140014','尿道造口闭合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',100,2698,'311140015','尿道阴道瘘修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',100,2699,'311140016','尿道直肠瘘修补术,经尿道','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',100,2700,'311140017','尿道吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',100,2701,'311140018','尿道下裂成形术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',100,2702,'311140019','尿道口成形术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',100,2703,'311140020','尿道口紧缩术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',100,2704,'311140021','耻骨弓下尿道修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',100,2705,'311140022','尿道成形术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',100,2706,'311140023','尿道会师术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',100,2707,'311140024','尿道修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',100,2708,'311140025','尿道后电切开术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',100,2709,'311140026','尿道内口切开术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',100,2710,'311140027','尿道切开术,经尿道','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',100,2711,'311140028','尿道扩张术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',100,2712,'311140029','尿道旁腺囊肿切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',101,2713,'311150001','肾周脓肿引流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',101,2714,'311150002','肾周区域探查术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',101,2715,'311150003','膀胱周围粘连松解术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',101,2716,'311150004','耻骨后探查术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',101,2717,'311150005','膀胱切开取石术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',101,2718,'311150006','膀胱周围组织探查术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',101,2719,'311150007','膀胱周围活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',101,2720,'311150008','肾周围活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',101,2721,'311150009','凯利-肯尼迪氏KELLEY-KENNED手术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',101,2722,'311150010','凯利-斯托克尔氏KELLEY-STOEKEL尿道折迭术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',101,2723,'311150011','奥克斯福德氏OXFORD尿失禁手术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',101,2724,'311150012','耻骨上悬吊尿道膀胱固定术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',101,2725,'311150013','尿道膀胱悬吊术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',101,2726,'311150014','米林--里德氏MILLIN-READNI尿道膀胱悬吊术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',101,2727,'311150015','尿道膀胱戈-弗-斯G-F-S氏悬吊术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',101,2728,'311150016','耻骨后尿道悬吊术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',101,2729,'311150017','尿道旁悬吊术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',101,2730,'311150018','膀胱尿道提肌悬带固定术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',101,2731,'311150019','尿道前固定术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',101,2732,'311150020','尿失禁修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',101,2733,'311150021','输尿管扩张术,膀胱镜下','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',101,2734,'311150022','输尿管膀胱口扩张术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',101,2735,'311150023','输尿管造瘘管置换术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',101,2736,'311150024','泌尿糸超声碎石术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',101,2737,'311150025','肾超声碎石术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',101,2738,'311150026','体外休克波碎石术(ESWL)','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',103,2739,'311210001','前列腺脓肿引流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',103,2740,'311210002','前列腺针刺活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',103,2741,'311210003','前列腺活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',103,2742,'311210004','精囊针吸活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',103,2743,'311210005','前列腺周围活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',103,2744,'311210006','前列腺切除术,经尿道','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',103,2745,'311210007','前列腺切除术,耻骨上经膀胱','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',103,2746,'311210008','前列腺切除术,耻骨后膀胱前','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',103,2747,'311210009','前列腺根治切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',103,2748,'311210010','前列腺病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',103,2749,'311210011','前列腺切除术,经会阴','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',103,2750,'311210012','前列腺切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual;
Insert Into 诊疗项目目录(类别,分类id,id,编码,名称,计算单位,标本部位,计算方式,单独应用,执行频率,适用性别,组合项目,操作类型,执行安排,执行科室,服务对象,计价性质,建档时间,撤档时间)
Select 'F',103,2751,'311210013','精囊切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',103,2752,'311210014','前列腺周围脓肿引流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',103,2753,'311210015','前列腺周围组织病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',103,2754,'311210016','前列腺术后止血术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',104,2755,'311220001','睾丸鞘膜切开引流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',104,2756,'311220002','阴囊切开引流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',104,2757,'311220003','鞘膜部分切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',104,2758,'311220004','鞘膜切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',104,2759,'311220005','阴囊病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',104,2760,'311220006','阴囊象皮病复原术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',104,2761,'311220007','阴囊输精管瘘切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',104,2762,'311220008','鞘膜翻转修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',104,2763,'311220009','睾丸鞘膜积液抽吸术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',104,2764,'311220010','鞘膜囊肿切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',104,2765,'311220011','睾丸探查术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',104,2766,'311220012','睾丸活组织检查, 经皮肤','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',104,2767,'311220013','睾丸活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',104,2768,'311220014','睾丸囊肿切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',104,2769,'311220015','睾丸肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',104,2770,'311220016','莫尔加尼氏MORGAGNI囊肿切除术,男性','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',104,2771,'311220017','男性MORGAGNI莫尔加尼氏囊肿切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',104,2772,'311220018','输卵管内窥镜下结扎伴挤压术,双侧','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',104,2773,'311220019','睾丸附睾切除术, 单侧','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',104,2774,'311220020','睾丸切除术, 单侧','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',104,2775,'311220021','睾丸根治性切除术, 双侧','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',104,2776,'311220022','睾丸切除术, 双侧','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',104,2777,'311220023','睾丸复位术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',104,2778,'311220024','睾丸固定术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',105,2779,'311230001','精索静脉高位结扎术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',105,2780,'311230002','附睾囊肿切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',105,2781,'311230003','附睾病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',105,2782,'311230004','精索病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',105,2783,'311230005','精索囊肿切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',105,2784,'311230006','附睾切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',105,2785,'311230007','输精管造口术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',105,2786,'311230008','输精管结扎术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',105,2787,'311230009','输精管切断术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',105,2788,'311230010','精索结扎术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',105,2789,'311230011','输精管切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',105,2790,'311230012','输精管吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',105,2791,'311230013','输精管附睾吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',105,2792,'311230014','输精管结扎去除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',105,2793,'311230015','精索粘连松解术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',105,2794,'311230016','输精管病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',106,2795,'311240001','包皮环切术','次','',3,1,1,0,0,'小',0,4,3,2,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',106,2796,'311240002','阴茎活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',106,2797,'311240003','包皮瘢痕切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',106,2798,'311240004','阴茎病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',106,2799,'311240005','阴茎全部切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',106,2800,'311240006','阴茎缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',106,2801,'311240007','痛形阴茎勃起松解术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',106,2802,'311240008','阴茎重建术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',106,2803,'311240009','阴茎截断再接术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',106,2804,'311240010','阴茎矫直术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',106,2805,'311240011','阴茎海绵体-阴茎头分流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',108,2806,'311310001','卵巢切开术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',108,2807,'311310002','卵巢切开探查术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',108,2808,'311310003','卵巢活组织检查,经皮肤针吸','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',108,2809,'311310004','卵巢活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',108,2810,'311310005','卵巢囊肿袋形缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',108,2811,'311310006','卵巢楔形切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',108,2812,'311310007','卵巢病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',108,2813,'311310008','卵巢部分切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',108,2814,'311310009','卵巢囊肿切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',108,2815,'311310010','卵巢肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',108,2816,'311310011','卵巢切除术,单侧','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',108,2817,'311310012','卵巢输卵管切除术,单侧','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',108,2818,'311310013','卵巢切除术,双侧','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',108,2819,'311310014','残留卵巢切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',108,2820,'311310015','卵巢输卵管切除术,双侧','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',108,2821,'311310016','卵巢缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',108,2822,'311310017','卵巢输卵管成形术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',108,2823,'311310018','卵巢固定术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',108,2824,'311310019','卵巢修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',108,2825,'311310020','卵巢输卵管粘连松解术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',108,2826,'311310021','输卵管粘连分离术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',108,2827,'311310022','卵巢扭转复位术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',108,2828,'311310023','腹腔镜下取卵术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',109,2829,'311320001','输卵管切开术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',109,2830,'311320002','输卵管妊娠清除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',109,2831,'311320003','输卵管探查术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',109,2832,'311320004','输卵管引流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',109,2833,'311320005','输卵管活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',109,2834,'311320006','输卵管内窥镜下结扎伴切断术,双侧','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',109,2835,'311320007','腹腔镜输卵管套环结扎术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',109,2836,'311320008','输卵管腹腔镜结扎术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',109,2837,'311320009','输卵管结扎术,双侧','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',109,2838,'311320010','输卵管粘堵术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',109,2839,'311320011','输卵管结扎术，单侧','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',109,2840,'311320012','输卵管切除术,单侧','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',109,2841,'311320013','输卵管切除术,双侧','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',109,2842,'311320014','莫尔加尼氏MORGAGNI囊肿切除术,女性','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',109,2843,'311320015','女性MORGAGNI莫尔加尼氏囊肿切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',109,2844,'311320016','输精管病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',109,2845,'311320017','输卵管囊肿切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',109,2846,'311320018','输卵管妊娠取出术伴输卵管切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',109,2847,'311320019','输卵管部分切除术,单侧','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',109,2848,'311320020','输卵管缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',109,2849,'311320021','输卵管--卵巢吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',109,2850,'311320022','输卵管吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',109,2851,'311320023','输卵管造口术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',109,2852,'311320024','输卵管成形术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',109,2853,'311320025','输卵管移植术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',109,2854,'311320026','输卵管通液术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',109,2855,'311320027','输卵管注气术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',109,2856,'311320028','输卵管抽吸术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',109,2857,'311320029','输卵管扩张术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2858,'311330001','子宫颈管扩张术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2859,'311330002','子宫颈成形术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2860,'311330003','子宫颈内膜活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual;
Insert Into 诊疗项目目录(类别,分类id,id,编码,名称,计算单位,标本部位,计算方式,单独应用,执行频率,适用性别,组合项目,操作类型,执行安排,执行科室,服务对象,计价性质,建档时间,撤档时间)
Select 'F',110,2861,'311330004','子宫颈楔形切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2862,'311330005','子宫颈囊肿袋形缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2863,'311330006','子宫颈息肉切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2864,'311330007','子宫颈肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2865,'311330008','子宫颈切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2866,'311330009','子宫颈环扎术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2867,'311330010','子宫颈陈旧性撕裂修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2868,'311330011','子宫切开探查术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2869,'311330012','子宫镜检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2870,'311330013','子宫活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2871,'311330014','子宫韧带活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2872,'311330015','子宫诊断性探查术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2873,'311330016','子宫腔粘连松解术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2874,'311330017','子宫隔切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2875,'311330018','子宫病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2876,'311330019','子宫肌瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2877,'311330020','子宫次全切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2878,'311330021','子宫部分切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2879,'311330022','子宫扩大性切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2880,'311330023','子宫全部切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2881,'311330024','全子宫切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2882,'311330025','全子宫伴双侧附件切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2883,'311330026','全子宫伴单侧附件切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2884,'311330027','子宫次全伴单侧附件切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2885,'311330028','子宫阴道式切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2886,'311330029','子宫腹式根治性切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2887,'311330030','子宫改良根治性切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2888,'311330031','子宫根治性切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2889,'311330032','子宫阴道式根治性切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2890,'311330033','子宫扩张刮宫术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2891,'311330034','人工流产,刮宫(扩宫)','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2892,'311330035','刮宫(扩宫)术, 产后或人工流产后','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2893,'311330036','诊断性刮宫','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2894,'311330037','阔韧带内异位妊娠切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2895,'311330038','子宫韧带病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2896,'311330039','子宫韧带肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2897,'311330040','子宫固定术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2898,'311330041','沃特金斯氏WATKINS手术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2899,'311330042','圆韧带缩短术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2900,'311330043','主韧带缩短术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2901,'311330044','子宫悬吊术(曼彻斯特MANCHERSTER手术)','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2902,'311330045','曼彻斯特MANCHERSTER手术(子宫悬吊)','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2903,'311330046','子宫缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2904,'311330047','子宫瘘管闭合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2905,'311330048','子宫陈旧破裂修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2906,'311330049','子宫成形术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2907,'311330050','电吸人工流产术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2908,'311330051','刮吸宫术, 分娩后','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2909,'311330052','刮吸宫术, 人工流产后','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2910,'311330053','钳刮术，流产后','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2911,'311330054','葡萄胎吸宫术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2912,'311330055','人工月经周期','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2913,'311330056','子宫内避孕器植入术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2914,'311330057','精液输卵管内移植术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2915,'311330058','人工授精(AID)','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2916,'311330059','人工授精(AIH)','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2917,'311330060','输卵管配子移植术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',110,2918,'311330061','内翻子宫手法复位术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',111,2919,'311340001','后穹窿穿剌术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',111,2920,'311340002','处女膜切开术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',111,2921,'311340003','后穹窿切开引流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',111,2922,'311340004','阴道粘连分离术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',111,2923,'311340005','阴道隔膜切断术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',111,2924,'311340006','阴道切开探查术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',111,2925,'311340007','阴道血肿切开引流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',111,2926,'311340008','阴道异物切开取出术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',111,2927,'311340009','阴道镜检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',111,2928,'311340010','后穹窿镜检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',111,2929,'311340011','阴道活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',111,2930,'311340012','处女膜切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',111,2931,'311340013','阴道病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',111,2932,'311340014','阴道肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',111,2933,'311340015','尿道阴道膜切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',111,2934,'311340016','阴道闭合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',111,2935,'311340017','阴道切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',111,2936,'311340018','阴道前后壁修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',111,2937,'311340019','阴道前壁修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',111,2938,'311340020','阴道后壁修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',111,2939,'311340021','直肠阴道隔疝修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',111,2940,'311340022','阴道建造术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',111,2941,'311340023','阴道再建术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',111,2942,'311340024','阴道缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',111,2943,'311340025','结肠阴道瘘修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',111,2944,'311340026','直肠阴道瘘修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',111,2945,'311340027','耻骨疏韧带悬吊术(阴道固定术)','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',111,2946,'311340028','阴道固定术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',111,2947,'311340029','阴道壁修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',111,2948,'311340030','阴道陈归性撕裂修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',111,2949,'311340031','阴道成形术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',111,2950,'311340032','阴道修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',111,2951,'311340033','阴道小肠疝修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',111,2952,'311340034','直肠子宫凹陷缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',112,2953,'311350001','外阴粘连松解术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',112,2954,'311350002','会阴切开术, 非产科','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',112,2955,'311350003','外阴活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',112,2956,'311350004','巴氏腺(前庭大腺)脓肿切开引流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',112,2957,'311350005','前庭大腺(巴氏腺)脓肿切开引流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',112,2958,'311350006','巴氏腺(前庭大腺)切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',112,2959,'311350007','前庭大腺(巴氏腺)切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',112,2960,'311350008','外阴病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',112,2961,'311350009','外阴肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',112,2962,'311350010','阴蒂部分切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',112,2963,'311350011','阴蒂成形术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',112,2964,'311350012','阴蒂切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',112,2965,'311350013','外阴根治切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',112,2966,'311350014','阴唇部分切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',112,2967,'311350015','阴唇切除术,单侧','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',112,2968,'311350016','外阴单纯切除术,双侧','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',112,2969,'311350017','阴唇切除术,双侧','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',112,2970,'311350018','会阴裂伤缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual;
Insert Into 诊疗项目目录(类别,分类id,id,编码,名称,计算单位,标本部位,计算方式,单独应用,执行频率,适用性别,组合项目,操作类型,执行安排,执行科室,服务对象,计价性质,建档时间,撤档时间)
Select 'F',112,2971,'311350019','外阴裂伤缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',112,2972,'311350020','会阴成形术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',112,2973,'311350021','会阴修补术, 女性','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',112,2974,'311350022','小阴唇切开成形术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',114,2975,'311410001','出口产钳助产术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',114,2976,'311410002','低位产钳助产术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',114,2977,'311410003','出口产钳助产伴会阴切开术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',114,2978,'311410004','低位产钳助产伴会阴切开术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',114,2979,'311410005','中位产钳助产伴会阴切开术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',114,2980,'311410006','中位产钳助产术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',114,2981,'311410007','高位产钳助产伴会阴切开术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',114,2982,'311410008','高位产钳助产术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',114,2983,'311410009','产钳部分臀位牵引术,头后出','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',114,2984,'311410010','臀位牵引术助产(臀抽术)','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',114,2985,'311410011','产钳完全臀信位牵引术, 头后出','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',114,2986,'311410012','臀位完全牵引术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',114,2987,'311410013','真空吸引伴会阴切开术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',114,2988,'311410014','胎头吸引助产','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',114,2989,'311410015','真空吸引助产','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',115,2990,'311420001','人工破膜引产','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',115,2991,'311420002','人工破膜,分娩时','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',115,2992,'311420003','剥膜引产','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',115,2993,'311420004','宫颈封闭术, 助产','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',115,2994,'311420005','水囊引产','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',115,2995,'311420006','胎儿转位术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',115,2996,'311420007','催产素点滴引产','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',115,2997,'311420008','改良药物引产','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',115,2998,'311420009','宫缩素点滴引产','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',115,2999,'311420010','手转胎头助产','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',115,3000,'311420011','手法助产术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',115,3001,'311420012','会阴切开伴缝合术, 产科','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',115,3002,'311420013','毁胎术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',115,3003,'311420014','胎儿穿颅术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',115,3004,'311420015','耻骨联合切开助产术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',116,3005,'311430001','古典式剖宫产术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',116,3006,'311430002','剖宫产术,古典式','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',116,3007,'311430003','子宫低位剖腹产术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',116,3008,'311430004','剖宫产术,子宫下段横切口','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',116,3009,'311430005','剖宫产样,子宫下段直切口','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',116,3010,'311430006','腹膜外剖宫产术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',116,3011,'311430007','膀胱上剖腹产术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',116,3012,'311430008','剖宫产术,腹膜外','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',116,3013,'311430009','卵巢妊娠清除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',116,3014,'311430010','异位妊娠去除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',116,3015,'311430011','子宫切开终止妊娠','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',116,3016,'311430012','剖腹产术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',117,3017,'311440001','雷夫诺尔羊膜腔内注射,人工流产','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',117,3018,'311440002','前列腺素羊膜腔内注射,人工流产','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',117,3019,'311440003','人工流产,雷诺尔羊膜腔内注射','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',117,3020,'311440004','人工流产,前列腺素羊腔内注射','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',117,3021,'311440005','羊膜腔内注射终止妊娠','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',117,3022,'311440006','羊膜穿剌术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',117,3023,'311440007','胎儿镜检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',117,3024,'311440008','胎儿监护','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',117,3025,'311440009','手取胎盘','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',117,3026,'311440010','胎盘钳夹术,产后','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',117,3027,'311440011','胎盘人工取出术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',117,3028,'311440012','子宫撕裂修补术,产科近期','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',117,3029,'311440013','子宫颈延期产科裂伤修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',117,3030,'311440014','子宫颈撕裂修补术,产科近期','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',117,3031,'311440015','子宫体近期产科裂伤修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',117,3032,'311440016','会阴撕裂修补术, 产科近期','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',117,3033,'311440017','小阴唇撕裂修补术,产科近期','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',117,3034,'311440018','阴道撕裂修补术,产科近期','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',117,3035,'311440019','产后手法子宫探查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',117,3036,'311440020','外阴血肿切开术,产科','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',119,3037,'311510001','面骨死骨取骨术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',119,3038,'311510002','面骨活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',119,3039,'311510003','鼻窦骨肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',119,3040,'311510004','颌骨囊肿切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',119,3041,'311510005','面骨病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',119,3042,'311510006','半下颌骨切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',119,3043,'311510007','下颌骨半切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',119,3044,'311510008','颧骨部分切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',119,3045,'311510009','上颌骨半切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',119,3046,'311510010','下颌骨全部切除伴骨移植术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',119,3047,'311510011','下颌骨全部切骨术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',119,3048,'311510012','上颌骨全部切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',119,3049,'311510013','颞下颌关节成形术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',119,3050,'311510014','下颌骨角闭合性成形术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',119,3051,'311510015','下颌骨角切骨术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',119,3052,'311510016','下颌骨体切骨术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',119,3053,'311510017','下颌骨成形术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',119,3054,'311510018','下颌骨切骨术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',119,3055,'311510019','上颌骨成形术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',119,3056,'311510020','上颌骨切骨术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',119,3057,'311510021','面骨切骨术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',119,3058,'311510022','颧骨骨折闭合性复位术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',119,3059,'311510023','颧骨骨折开放性复位术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',119,3060,'311510024','上颌骨骨折闭合复位术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',119,3061,'311510025','上颌骨骨折切开复位术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',119,3062,'311510026','下颌骨骨折闭合性复位术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',119,3063,'311510027','下颌骨骨折切开复位术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',119,3064,'311510028','眶骨骨折闭合性复位术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',119,3065,'311510029','眶骨骨折开放性复位术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',119,3066,'311510030','面骨移植术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',119,3067,'311510031','颞下颌关节脱位闭合复位术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',119,3068,'311510032','颞下颌关节脱位切开复位术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3069,'311520001','股骨死骨切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3070,'311520002','指死骨切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3071,'311520003','趾死骨切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3072,'311520004','椎骨死骨取出术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3073,'311520005','骨切开引流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3074,'311520006','股骨头切开术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3075,'311520007','肩甲骨楔形切骨术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3076,'311520008','锁骨楔形切骨术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3077,'311520009','胸骨楔形切骨术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3078,'311520010','肱骨楔形切骨术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3079,'311520011','尺骨楔形切骨术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3080,'311520012','桡骨楔形切骨术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual;
Insert Into 诊疗项目目录(类别,分类id,id,编码,名称,计算单位,标本部位,计算方式,单独应用,执行频率,适用性别,组合项目,操作类型,执行安排,执行科室,服务对象,计价性质,建档时间,撤档时间)
Select 'F',120,3081,'311520013','腕骨楔形切骨术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3082,'311520014','股骨楔形切骨术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3083,'311520015','髌骨楔形切骨术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3084,'311520016','腓骨楔形切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3085,'311520017','胫骨楔形切骨术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3086,'311520018','跗骨楔形切骨术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3087,'311520019','指(趾)骨楔形切骨术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3088,'311520020','椎骨楔形切骨术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3089,'311520021','骨盆楔形切骨术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3090,'311520022','脊柱后路楔形截骨矫形术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3091,'311520023','肩甲骨切骨术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3092,'311520024','锁骨切骨术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3093,'311520025','胸骨切骨术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3094,'311520026','肱骨切骨术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3095,'311520027','尺骨切骨术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3096,'311520028','桡骨切骨术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3097,'311520029','腕骨切骨术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3098,'311520030','股骨切骨术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3099,'311520031','髌骨切骨术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3100,'311520032','腓骨切骨术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3101,'311520033','胫骨切骨术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3102,'311520034','跗骨切骨术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3103,'311520035','跖骨切骨术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3104,'311520036','骨盆切骨术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3105,'311520037','指(趾)骨切骨术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3106,'311520038','椎骨切骨术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3107,'311520039','骨活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3108,'311520040','跖骨楔形切骨母外翻矫正术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3109,'311520041','母囊肿切除伴关节固定术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3110,'311520042','母囊肿切除伴软组织修整术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3111,'311520043','小趾囊肿切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3112,'311520044','凯勒氏KELLER手术(母囊肿切除术)','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3113,'311520045','母囊肿切除KELLER凯勒氏手术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3114,'311520046','骨病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3115,'311520047','骨棘切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3116,'311520048','骨囊肿切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3117,'311520049','骨髓炎刮除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3118,'311520050','骨肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3119,'311520051','半椎切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3120,'311520052','手指骨肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3121,'311520053','肋骨切除为骨移植','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3122,'311520054','髂骨部分切除,用于移植术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3123,'311520055','尺骨小骨切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3124,'311520056','桡骨小头切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3125,'311520057','掌骨部分切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3126,'311520058','股骨颈切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3127,'311520059','股骨头切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3128,'311520060','腓骨小头切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3129,'311520061','耻骨部分切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3130,'311520062','肋骨切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3131,'311520063','籽骨切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3132,'311520064','骨移植术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3133,'311520065','肱骨植骨术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3134,'311520066','桡或尺骨植骨术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3135,'311520067','股骨植骨术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3136,'311520068','髌骨植骨术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3137,'311520069','腓骨植骨术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3138,'311520070','胫或腓骨植骨术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3139,'311520071','指骨植骨术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3140,'311520072','趾骨植骨术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3141,'311520073','骨缩短术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3142,'311520074','骨延长术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3143,'311520075','骨融合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3144,'311520076','肱骨内固定术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3145,'311520077','脊柱哈林顿氏棍植入术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3146,'311520078','脊柱卢奎LUQUE棍内固定','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3147,'311520079','骨折内固定物取出术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3148,'311520080','钢板内固定取出术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3149,'311520081','脊柱哈林顿氏棍切取出术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',120,3150,'311520082','脊柱卢奎LUQUE棍取出术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',121,3151,'311530001','小腿骨折闭合复位术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',121,3152,'311530002','髌骨骨抓髌器外固定术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',121,3153,'311530003','骨折闭合复位伴内固定术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',121,3154,'311530004','骨折开放复位','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',121,3155,'311530005','骨折开放复位伴内固定术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',121,3156,'311530006','股骨骨折切开复位伴内固定','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',121,3157,'311530007','开放性骨折清创术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',121,3158,'311530008','肩关节脱位闭合复位术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',121,3159,'311530009','肘节脱位闭合复位术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',121,3160,'311530010','腕关节脱位闭合性复位术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',121,3161,'311530011','趾关节脱位闭合复位术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',121,3162,'311530012','指关节脱位闭合复位术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',121,3163,'311530013','髋关节脱位闭合性复位术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',121,3164,'311530014','膝关节脱位闭合复位术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',121,3165,'311530015','踝关节脱位闭合复位术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',121,3166,'311530016','肩关节脱位切开复位术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',121,3167,'311530017','肘关节脱位切开复位术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',121,3168,'311530018','腕关节脱位切开复位术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',121,3169,'311530019','指关节脱位切开复位术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',121,3170,'311530020','髋关节脱位切开复位术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',121,3171,'311530021','膝关节脱位切开复位术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',121,3172,'311530022','踝关节脱位切开复位术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',121,3173,'311530023','趾关节脱位切开复位术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',121,3174,'311530024','关节内部假体装置取出术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',121,3175,'311530025','膝关节切开探查术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',121,3176,'311530026','椎间盘探查术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',121,3177,'311530027','关节镜检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',121,3178,'311530028','膝关节镜检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',121,3179,'311530029','关节活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',121,3180,'311530030','肩关节活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',121,3181,'311530031','肘关节活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',121,3182,'311530032','髋关节活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',121,3183,'311530033','膝关节活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',121,3184,'311530034','踝关节活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',121,3185,'311530035','关节松解术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',121,3186,'311530036','腕横韧带松解术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',121,3187,'311530037','足韧带松解术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',121,3188,'311530038','膝韧带松解术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',121,3189,'311530039','椎板切除伴椎间盘疝切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',121,3190,'311530040','膝半月板切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual;
Insert Into 诊疗项目目录(类别,分类id,id,编码,名称,计算单位,标本部位,计算方式,单独应用,执行频率,适用性别,组合项目,操作类型,执行安排,执行科室,服务对象,计价性质,建档时间,撤档时间)
Select 'F',121,3191,'311530041','滑膜病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',121,3192,'311530042','滑膜切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',121,3193,'311530043','膝滑膜切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',121,3194,'311530044','关节病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',121,3195,'311530045','韧带病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',121,3196,'311530046','韧带切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',122,3197,'311540001','脊柱融合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',122,3198,'311540002','颅颈融合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',122,3199,'311540003','颈椎融合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',122,3200,'311540004','胸椎融合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',122,3201,'311540005','胸腰椎哈林顿棍融合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',122,3202,'311540006','胸腰椎融合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',122,3203,'311540007','腰疝融合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',122,3204,'311540008','腰骶部脊柱融合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',122,3205,'311540009','脊柱假关节矫形术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',122,3206,'311540010','颈椎假关节融合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',122,3207,'311540011','腰疝假关节融合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',122,3208,'311540012','颅颈假关节融合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',122,3209,'311540013','胸部假关节融合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',122,3210,'311540014','踝关节融合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',122,3211,'311540015','胫距骨融合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',122,3212,'311540016','三关节融合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',122,3213,'311540017','趾关节融合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',122,3214,'311540018','髋关节融合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',122,3215,'311540019','膝关节融合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',122,3216,'311540020','肩关节融合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',122,3217,'311540021','肘关节融合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',122,3218,'311540022','腕关节融合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',122,3219,'311540023','(全)膝关节置换术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',122,3220,'311540024','膝关节五合一手术(内半月板,内韧带,股内肌','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',122,3221,'311540025','膝关节三联修补术(内半月板,前交叉韧带,内','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',122,3222,'311540026','髌骨稳定术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',122,3223,'311540027','膝关节修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',122,3224,'311540028','(全)踝关节置换术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',122,3225,'311540029','髋关节重用术,用甲基丙烯酸酯','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',122,3226,'311540030','(全)髋关节甲基丙烯酸酯置换术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',122,3227,'311540031','股骨头和髋臼假体置换术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',122,3228,'311540032','髋关节重建术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',122,3229,'311540033','(全)髋关节置换术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',122,3230,'311540034','股骨头置换, 甲基丙烯酸脂','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',122,3231,'311540035','股骨头假体置换术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',122,3232,'311540036','髋臼假体置换,甲基丙烯酸酯','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',122,3233,'311540037','髋关节修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',122,3234,'311540038','指关节修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',122,3235,'311540039','肩关节合成假体成形术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',122,3236,'311540040','习惯性肩关节脱位修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',122,3237,'311540041','肩关节修正术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',122,3238,'311540042','肘关节修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',122,3239,'311540043','腕关节修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',122,3240,'311540044','关节穿剌术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',123,3241,'311550001','手部腱鞘切开术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',123,3242,'311550002','手部腱鞘松解术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',123,3243,'311550003','手部肌肉切开术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',123,3244,'311550004','手部肌肉异物切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',123,3245,'311550005','手部粘液囊切开术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',123,3246,'311550006','手部肌腱切断术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',123,3247,'311550007','手部筋膜切断术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',123,3248,'311550008','手部肌肉松解术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',123,3249,'311550009','手部腱鞘病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',123,3250,'311550010','手部腱鞘囊肿切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',123,3251,'311550011','手部肌肉病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',123,3252,'311550012','手粘液囊切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',123,3253,'311550013','手腱鞘切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',123,3254,'311550014','手部筋膜切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',123,3255,'311550015','手腱鞘缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',123,3256,'311550016','桡侧屈腕肌腱缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',123,3257,'311550017','手屈肌腱缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',123,3258,'311550018','手部肌腱缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',123,3259,'311550019','伸拇长肌缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',123,3260,'311550020','手筋膜缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',123,3261,'311550021','手部肌腱前徙术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',123,3262,'311550022','手部肌腱退缩术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',123,3263,'311550023','手部肌腱再接术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',123,3264,'311550024','手部肌肉再接术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',123,3265,'311550025','手部肌腱延长术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',123,3266,'311550026','手部肌腱移位术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',123,3267,'311550027','手部肌肉移位术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',123,3268,'311550028','脚趾代拇指术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',123,3269,'311550029','拇指再建术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',123,3270,'311550030','手部肌腱成形术,用其它部位移植的肌腱','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',123,3271,'311550031','手指肌腱成形术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',123,3272,'311550032','手部肌腱松解术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',124,3273,'311560001','腱鞘探查术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',124,3274,'311560002','肌肉筋膜减压术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',124,3275,'311560003','肌肉内异物取出术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',124,3276,'311560004','前臂肌腱松解术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',124,3277,'311560005','髂胫束松解术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',124,3278,'311560006','跖筋膜松解术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',124,3279,'311560007','肌肉松解术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',124,3280,'311560008','前斜角肌松解术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',124,3281,'311560009','臀部肌肉松解术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',124,3282,'311560010','胸锁乳突肌切断术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',124,3283,'311560011','肌肉活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',124,3284,'311560012','肌腱囊肿切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',124,3285,'311560013','腱鞘病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',124,3286,'311560014','腱鞘囊肿切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',124,3287,'311560015','肌肉病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',124,3288,'311560016','肌肉骨化性损害切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',124,3289,'311560017','肌肉血管瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',124,3290,'311560018','肌肉肿瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',124,3291,'311560019','贝克氏BAKERS国窝囊肿切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',124,3292,'311560020','国窝贝克BAKERS氏囊肿切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',124,3293,'311560021','滑膜囊肿切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',124,3294,'311560022','颈部软组织病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',124,3295,'311560023','肘窝囊肿切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',124,3296,'311560024','肌腱切除为移植','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',124,3297,'311560025','筋膜切除为移植','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',124,3298,'311560026','颈伸肌部分切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',124,3299,'311560027','粘液囊切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',124,3300,'311560028','腱鞘缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual;
Insert Into 诊疗项目目录(类别,分类id,id,编码,名称,计算单位,标本部位,计算方式,单独应用,执行频率,适用性别,组合项目,操作类型,执行安排,执行科室,服务对象,计价性质,建档时间,撤档时间)
Select 'F',124,3301,'311560029','前臂肌腱缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',124,3302,'311560030','趾肌腱缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',124,3303,'311560031','肱二头肌缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',124,3304,'311560032','股二头肌腱徙前术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',124,3305,'311560033','肌腱退缩术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',124,3306,'311560034','肌腱再接术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',124,3307,'311560035','肌肉再接术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',124,3308,'311560036','肌腱移位术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',124,3309,'311560037','肌腱瓣转移术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',124,3310,'311560038','肌肉移位术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',124,3311,'311560039','股外斜肌代臂中肌','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',124,3312,'311560040','股外斜肌代股四头肌','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',124,3313,'311560041','肌肉移植术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',124,3314,'311560042','胫后肌移植术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',124,3315,'311560043','跟腱缩短术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',124,3316,'311560044','跟腱延长术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',124,3317,'311560045','肌腱延长术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',124,3318,'311560046','股四头肌成形术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',124,3319,'311560047','肌肉修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',124,3320,'311560048','肌疝修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',124,3321,'311560049','胸大肌修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',124,3322,'311560050','足母伸肌腱固定术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',124,3323,'311560051','跟腱修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',124,3324,'311560052','肌腱成形术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',124,3325,'311560053','肌腱固定术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',124,3326,'311560054','胫后肌健吻合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',124,3327,'311560055','足畸形矫形术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',124,3328,'311560056','筋膜成形术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',124,3329,'311560057','筋膜延长术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',124,3330,'311560058','滑囊穿剌术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',125,3331,'311570001','手指截指术,拇指除外','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',125,3332,'311570002','腕关节离断术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',125,3333,'311570003','前臂截肢术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',125,3334,'311570004','上臂切断术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',125,3335,'311570005','肩关节离断术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',125,3336,'311570006','趾离断术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',125,3337,'311570007','足截断术,前部','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',125,3338,'311570008','踝关节离断术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',125,3339,'311570009','小腿截肢术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',125,3340,'311570010','膝关节离断术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',125,3341,'311570011','大腿截肢术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',125,3342,'311570012','髋关节离断术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',125,3343,'311570013','拇指再植术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',125,3344,'311570014','断指再植术(手指再植术)','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',125,3345,'311570015','手指再植术,(断指再植)','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',125,3346,'311570016','前臂再植术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',125,3347,'311570017','母趾再植术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',125,3348,'311570018','断肢再植术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',125,3349,'311570019','截肢残端修整术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',125,3350,'311570020','假肢安装','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',125,3351,'311570021','前臂假肢安装','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',125,3352,'311570022','臂假肢安装','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',125,3353,'311570023','膝关节上假肢安装','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',125,3354,'311570024','膝关节下假肢安装','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',125,3355,'311570025','腿假肢安装','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',127,3356,'311610001','乳腺脓肿引流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',127,3357,'311610002','乳腺活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',127,3358,'311610003','乳腺病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',127,3359,'311610004','乳腺纤维瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',127,3360,'311610005','乳腺次全切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',127,3361,'311610006','副乳头切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',127,3362,'311610007','乳头切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',127,3363,'311610008','乳房单纯切除术,单侧','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',127,3364,'311610009','乳房单纯切除术.双侧','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',127,3365,'311610010','乳房改良根治术,单侧','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',127,3366,'311610011','乳房根治术,单侧','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',127,3367,'311610012','乳房根治术,双侧','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',127,3368,'311610013','乳房扩大性根治术,单侧','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',127,3369,'311610014','乳房扩大性根治术,双侧','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',127,3370,'311610015','乳房增大成形术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',127,3371,'311610016','乳房注入增大术,双侧','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',127,3372,'311610017','乳房植入术,双侧','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',127,3373,'311610018','乳房重建术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',127,3374,'311610019','乳房固定术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',127,3375,'311610020','乳房修补术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',128,3376,'311620001','皮肤和皮下组织的切开引流术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',128,3377,'311620002','皮下组织异物切开取出术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',128,3378,'311620003','皮肤活组织检查','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',128,3379,'311620004','背部伤口清创术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',128,3380,'311620005','臂伤口清创术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',128,3381,'311620006','腹部伤口清创术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',128,3382,'311620007','颈部清创术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',128,3383,'311620008','皮肤清创术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',128,3384,'311620009','手清创术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',128,3385,'311620010','下肢清创术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',128,3386,'311620011','趾甲取除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',128,3387,'311620012','指甲去除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',128,3388,'311620013','腹壁皮肤瘢痕切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',128,3389,'311620014','皮肤瘢痕切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',128,3390,'311620015','皮肤病损激光治疗','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',128,3391,'311620016','皮肤病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',128,3392,'311620017','皮肤及皮下血管瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',128,3393,'311620018','皮肤囊肿切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',128,3394,'311620019','皮下囊瘤切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',128,3395,'311620020','皮下组织病损切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',128,3396,'311620021','皮脂囊肿切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',128,3397,'311620022','汗腺肿瘤根治切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',128,3398,'311620023','头皮缝合术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',128,3399,'311620024','头皮再植术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',128,3400,'311620025','手部游离植皮术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',128,3401,'311620026','上肢植皮术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',128,3402,'311620027','下肢植皮术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',128,3403,'311620028','皮瓣自体植皮术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',128,3404,'311620029','瓣状或蒂状移植皮片向手固定','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',128,3405,'311620030','面部皱纹切除术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',128,3406,'311620031','皮肤瘢痕松解术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',128,3407,'311620032','手部瘢痕松解术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'F',128,3408,'311620033','并指分指术','次','',3,1,1,0,0,'小',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',130,3409,'320100001','脑电图','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',130,3410,'320100002','脑电地形图','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual;
Insert Into 诊疗项目目录(类别,分类id,id,编码,名称,计算单位,标本部位,计算方式,单独应用,执行频率,适用性别,组合项目,操作类型,执行安排,执行科室,服务对象,计价性质,建档时间,撤档时间)
Select 'D',130,3411,'320100003','24小时动态脑电图','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',130,3412,'320100004','神经传导速度测定','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',130,3413,'320100005','体感诱发电位','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',130,3414,'320100006','运动诱发电位','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',130,3415,'320100007','事件相关电位','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',130,3416,'320100008','脑干诱发电位','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',130,3417,'320100009','感觉阈值测量','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',130,3418,'320100010','雷诺氏现象试验','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',130,3419,'320100011','肌电图','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',130,3420,'320100012','单纤维肌电图','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',130,3421,'320100013','动态肌电图','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',130,3422,'320100014','经皮穿刺三叉神经半月节注射治疗术','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',130,3423,'320100015','经皮穿刺三叉神经半月节及感觉根射频温控热凝术','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',131,3424,'320200001','肺功能检查','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',131,3425,'320200002','低压氧疗','次','',3,1,1,0,0,'0',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',131,3426,'320200003','机械通气(呼吸机辅助呼吸)','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',131,3427,'320200004','纤维支气管镜检查','次','',3,1,1,0,0,'其他',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',131,3428,'320200005','支气管肺泡灌洗术','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',131,3429,'320200006','湿化和气溶胶吸入','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',131,3430,'320200007','气体代谢分析','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',131,3431,'320200008','胸腔镜检查','次','',3,1,1,0,0,'其他',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',131,3432,'320200009','纵隔镜检查','次','',3,1,1,0,0,'其他',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',131,3433,'320200010','高压氧治疗','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',132,3434,'320300001','心电图检查','次','',3,1,1,0,0,'心电',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',132,3435,'320300002','食管内心电图','次','',3,1,1,0,0,'心电',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',132,3436,'320300003','动态心电图','次','',3,1,1,0,0,'心电',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',132,3437,'320300004','心电图平板运动试验','次','',3,1,1,0,0,'心电',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',132,3438,'320300005','无创心功能检查','次','',3,1,1,0,0,'心电',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',132,3439,'320300006','射频消融术','次','',3,1,1,0,0,'0',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',132,3440,'320300007','永久起博器治疗','次','',3,1,1,0,0,'0',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',132,3441,'320300008','临时起博器治疗','次','',3,1,1,0,0,'0',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',132,3442,'320300009','心脏电复律','次','',3,1,1,0,0,'0',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',133,3443,'320400001','骨髓穿刺术','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',133,3444,'320400002','骨髓活检术','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',133,3445,'320400003','血浆置换','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',133,3446,'320400004','血液稀释疗法','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',134,3447,'320500001','纤维食管镜检查','次','',3,1,1,0,0,'其他',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',134,3448,'320500002','胃电图、肠电图','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',134,3449,'320500003','纤维胃、十二指肠镜检查','次','',3,1,1,0,0,'其他',0,4,3,0,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',134,3450,'320500004','纤维小肠镜检查','次','',3,1,1,0,0,'其他',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',134,3451,'320500005','纤维全结肠镜检查','次','',3,1,1,0,0,'其他',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',134,3452,'320500006','乙状结肠镜检查','次','',3,1,1,0,0,'其他',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',134,3453,'320500007','直肠镜检查','次','',3,1,1,0,0,'其他',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',134,3454,'320500008','肛门镜检查','次','',3,1,1,0,0,'其他',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',134,3455,'320500009','肛门指检','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',134,3456,'320500010','肛、直肠肌电测量','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',134,3457,'320500011','直肠、肛管癌冷冻治疗','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',134,3458,'320500012','内痔冷冻治疗','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',134,3459,'320500013','内痔微波治疗','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',134,3460,'320500014','直肠粘膜激光烧灼','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',134,3461,'320500015','胆道镜检查','次','',3,1,1,0,0,'其他',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',134,3462,'320500016','腹腔镜检查','次','',3,1,1,0,0,'其他',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',135,3463,'320600001','血液透析','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',135,3464,'320600002','血液滤过','次','',3,1,1,0,0,'0',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',135,3465,'320600003','血液透析滤过','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',135,3466,'320600004','血液灌流(加透析)','次','',3,1,1,0,0,'0',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',135,3467,'320600005','连续性动静脉血液滤过(透析)','次','',3,1,1,0,0,'0',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',135,3468,'320600006','腹膜透析疗法','次','',3,1,1,0,0,'0',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',135,3469,'320600007','经皮肾盂镜检查','次','',3,1,1,0,0,'其他',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',135,3470,'320600008','经尿道输尿管镜检查','次','',3,1,1,0,0,'其他',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',135,3471,'320600009','膀胱冲洗','次','',3,1,1,0,0,'0',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',135,3472,'320600010','膀胱注射','次','',3,1,1,0,0,'0',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',135,3473,'320600011','膀胱灌注','次','',3,1,1,0,0,'0',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',135,3474,'320600012','尿道冲洗','次','',3,1,1,0,0,'0',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',135,3475,'320600013','尿道肉阜电灼术','次','',3,1,1,0,0,'0',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',135,3476,'320600014','经尿道治疗尿失禁','次','',3,1,1,0,0,'0',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',135,3477,'320600015','体外冲击波碎石','次','',3,1,1,0,0,'0',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',135,3478,'320600016','前列腺注射','次','',3,1,1,0,0,'0',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',135,3479,'320600017','前列腺灌注','次','',3,1,1,0,0,'0',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',135,3480,'320600018','前列腺封闭','次','',3,1,1,0,0,'0',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',135,3481,'320600019','微波或射频前列腺治疗','次','',3,1,1,0,0,'0',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',136,3482,'320700001','荧光检查','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',136,3483,'320700002','子宫托治疗','次','',3,1,1,0,0,'0',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',136,3484,'320700003','子宫输卵管通液术','次','',3,1,1,0,0,'0',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',136,3485,'320700004','子宫内翻复位术','次','',3,1,1,0,0,'0',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',136,3486,'320700005','妇科激光治疗','次','',3,1,1,0,0,'0',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',136,3487,'320700006','妇科微波治疗','次','',3,1,1,0,0,'0',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',136,3488,'320700007','妇科冷冻治疗','次','',3,1,1,0,0,'0',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',136,3489,'320700008','妇科电熨治疗','次','',3,1,1,0,0,'0',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',136,3490,'320700009','外阴病光照射治疗','次','',3,1,1,0,0,'0',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',136,3491,'320700010','产前检查','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',136,3492,'320700011','双合诊检查','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',136,3493,'320700012','胎儿心电图','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',136,3494,'320700013','羊膜镜检查','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',136,3495,'320700014','新生儿兰光治疗','次','',3,1,1,0,0,'0',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',136,3496,'320700015','新生儿暖箱','次','',3,1,1,0,0,'0',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',136,3497,'320700016','新生儿油浴','次','',3,1,1,0,0,'0',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',138,3498,'320810001','韦氏幼儿智力量表(C－WYCSI)','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',138,3499,'320810002','修订韦氏儿童智力量表(C－WISC)','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',138,3500,'320810003','修订韦氏成人智力量表(WAIS－RC)','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',138,3501,'320810004','修订韦氏记忆量表(WMS)成人本','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',138,3502,'320810005','修订韦氏记忆量表(WMS)儿童本','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',138,3503,'320810006','计算机多相个性测查表(CMCIS)或明尼苏达多相个性测查表(MMPI)','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',138,3504,'320810007','艾森克个性问卷','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',138,3505,'320810008','卡特尔16种人格因素测验(16PF)','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',138,3506,'320810009','数字划销测验','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',138,3507,'320810010','症状自评量表(SCL－90)','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',138,3508,'320810011','瑞文测验联合型','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',138,3509,'320810012','中国比内测验','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',138,3510,'320810013','修订HR神经心理成套测验(成人式)','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',138,3511,'320810014','修订HR神经心理成套测验(幼儿式)','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',138,3512,'320810015','汉密顿抑郁量表(HAND)','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',138,3513,'320810016','汉密顿焦虑量表(HAMA)','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',138,3514,'320810017','抑郁自评量表(SDS)','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',138,3515,'320810018','焦虑自评量表(SAS)','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',138,3516,'320810019','Marks 恐怖强迫量表(MSCPOR)','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',138,3517,'320810020','Bech－Rafaelsen躁狂量表(BRMS)','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',138,3518,'320810021','Achenbach儿童行为量表(CBCL)','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',138,3519,'320810022','副反应量表(TESS)','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',138,3520,'320810023','临床疗效总评量表(CGI)','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual;
Insert Into 诊疗项目目录(类别,分类id,id,编码,名称,计算单位,标本部位,计算方式,单独应用,执行频率,适用性别,组合项目,操作类型,执行安排,执行科室,服务对象,计价性质,建档时间,撤档时间)
Select 'D',138,3521,'320810024','简明精神病量表(BPRS)','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',138,3522,'320810025','神经精神病学临床评定量表(SCAN)','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',138,3523,'320810026','精神障碍诊断量表(DSMD)','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',138,3524,'320810027','精神障碍定式检查量表(SCID)','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',138,3525,'320810028','阳性阴性症状评定量表(PANSS)','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',138,3526,'320810029','Yale－Brown强迫量表(Y－BOCS)','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',138,3527,'320810030','社会功能缺陷筛选量表(SDSS)','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',138,3528,'320810031','套瓦(TOVA)注意力竞量测试','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',139,3529,'320820001','抽搐电休克治疗','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',139,3530,'320820002','无抽搐电休克治疗','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',139,3531,'320820003','暴露疗法和半暴露疗法','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',139,3532,'320820004','冷光光量子血疗','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',139,3533,'320820005','胰岛素低血糖和休克治疗','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',139,3534,'320820006','行为观察和治疗','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',139,3535,'320820007','电磁场治疗','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',139,3536,'320820008','新电治疗','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',139,3537,'320820009','脑电生物反馈治疗','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',139,3538,'320820010','脑反射治疗','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',139,3539,'320820011','智能电针','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',139,3540,'320820012','经络氧疗法','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',139,3541,'320820013','感觉统合治疗','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',139,3542,'320820014','工娱治疗','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',139,3543,'320820015','音乐治疗','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',139,3544,'320820016','暗示疗法','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',139,3545,'320820017','松驰治疗','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',139,3546,'320820018','漂浮治疗','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',139,3547,'320820019','听力整合及语言训练','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',139,3548,'320820020','心理咨询','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',139,3549,'320820021','心理治疗','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',139,3550,'320820022','麻醉分析','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',139,3551,'320820023','催眠治疗','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',139,3552,'320820024','森田疗法','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',139,3553,'320820025','行为矫正','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',139,3554,'320820026','厌恶治疗','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',141,3555,'330100001','脑灌注显像','次','',3,1,1,0,0,'其他',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',141,3556,'330100002','静息心肌灌注显像','次','',3,1,1,0,0,'其他',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',141,3557,'330100003','运动心肌灌注显像','次','',3,1,1,0,0,'其他',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',141,3558,'330100004','心室显像','次','',3,1,1,0,0,'其他',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',141,3559,'330100005','肝脏显像','次','',3,1,1,0,0,'其他',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',141,3560,'330100006','肝血池显像','次','',3,1,1,0,0,'其他',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',141,3561,'330100007','肝胆道系统显像','次','',3,1,1,0,0,'其他',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',141,3562,'330100008','胃食道返流显像','次','',3,1,1,0,0,'其他',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',141,3563,'330100009','肺灌注显像','次','',3,1,1,0,0,'其他',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',141,3564,'330100010','肺通气显像','次','',3,1,1,0,0,'其他',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',141,3565,'330100011','肾功能显像(GFR)','次','',3,1,1,0,0,'其他',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',141,3566,'330100012','甲状腺显像','次','',3,1,1,0,0,'其他',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',141,3567,'330100013','乳腺肿瘤显像','次','',3,1,1,0,0,'其他',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',141,3568,'330100014','淋巴显像','次','',3,1,1,0,0,'其他',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',141,3569,'330100015','甲状腺肿瘤显像','次','',3,1,1,0,0,'其他',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',141,3570,'330100016','全身骨显像','次','',3,1,1,0,0,'其他',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',142,3571,'330200001','711I治疗甲状腺功能亢进症','次','',3,1,1,0,0,'5',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',142,3572,'330200002','711I治疗功能自主性甲状腺腺瘤','次','',3,1,1,0,0,'5',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',142,3573,'330200003','β敷贴治疗','次','',3,1,1,0,0,'5',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',144,3574,'340100001','深部X线照射','次','',3,1,1,0,0,'5',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',144,3575,'340100002','60钴外照射','次','',3,1,1,0,0,'5',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',144,3576,'340100003','直线加速器放疗(固定照射)','次','',3,1,1,0,0,'5',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',144,3577,'340100004','X刀治疗','次','',3,1,1,0,0,'5',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',144,3578,'340100005','伽玛刀治疗','次','',3,1,1,0,0,'5',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',144,3579,'340100006','全身60钴照射','次','',3,1,1,0,0,'5',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',144,3580,'340100007','全身X线照射','次','',3,1,1,0,0,'5',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',144,3581,'340100008','全身电子线照射','次','',3,1,1,0,0,'5',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',145,3582,'340200001','浅表部位后装治疗','次','',3,1,1,0,0,'5',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',145,3583,'340200002','腔内放疗','次','',3,1,1,0,0,'5',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',145,3584,'340200003','组织间插置放疗','次','',3,1,1,0,0,'5',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',145,3585,'340200004','皮肤贴敷放疗','次','',3,1,1,0,0,'5',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',145,3586,'340200005','血管内放疗','次','',3,1,1,0,0,'5',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',145,3587,'340200006','冠状动脉内放疗','次','',3,1,1,0,0,'5',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',146,3588,'340300001','浅表部位热疗','次','',3,1,1,0,0,'5',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',146,3589,'340300002','深部热疗','次','',3,1,1,0,0,'5',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',148,3590,'350100001','直流电及药物离子导入疗法','次','',3,1,1,0,0,'0',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',148,3591,'350100002','直流电水浴疗法','次','',3,1,1,0,0,'',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',148,3592,'350100003','神经肌肉低频电刺激疗法','次','',3,1,1,0,0,'',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',148,3593,'350100004','失神经支配肌肉低频电刺激法','次','',3,1,1,0,0,'',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',148,3594,'350100005','痉挛肌低频电刺激法','次','',3,1,1,0,0,'',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',148,3595,'350100006','低频感应电疗法','次','',3,1,1,0,0,'',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',148,3596,'350100007','低频电兴奋疗法','次','',3,1,1,0,0,'',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',148,3597,'350100008','间动低频电疗法','次','',3,1,1,0,0,'',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',148,3598,'350100009','音频中频电疗法','次','',3,1,1,0,0,'',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',148,3599,'350100010','正弦调制中频电疗法','次','',3,1,1,0,0,'',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',148,3600,'350100011','静态干扰中频电疗法','次','',3,1,1,0,0,'',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',148,3601,'350100012','动态干扰中频电疗法','次','',3,1,1,0,0,'',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',148,3602,'350100013','长波高频电疗法','次','',3,1,1,0,0,'',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',148,3603,'350100014','中波高频电疗法','次','',3,1,1,0,0,'',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',148,3604,'350100015','短波高频电疗法','次','',3,1,1,0,0,'',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',148,3605,'350100016','超短波高频电疗法','次','',3,1,1,0,0,'',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',148,3606,'350100017','微波高频电疗法','次','',3,1,1,0,0,'',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',148,3607,'350100018','共鸣火花电疗法','次','',3,1,1,0,0,'',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',149,3608,'350200001','红外线疗法','次','',3,1,1,0,0,'',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',149,3609,'350200002','紫外线疗法','次','',3,1,1,0,0,'',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',149,3610,'350200003','激光疗法','次','',3,1,1,0,0,'',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',149,3611,'350200004','特定电磁波疗法','次','',3,1,1,0,0,'',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',150,3612,'350300001','超声波疗法','次','',3,1,1,0,0,'',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',150,3613,'350300002','超声药物透入疗法','次','',3,1,1,0,0,'',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',150,3614,'350300003','超声雾化吸入疗法','次','',3,1,1,0,0,'',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',150,3615,'350300004','超声—间动电疗法','次','',3,1,1,0,0,'',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',150,3616,'350300005','超声—调制中频电叠加疗法','次','',3,1,1,0,0,'',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',151,3617,'350400001','石蜡疗法','次','',3,1,1,0,0,'',0,4,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',152,3618,'360000001','平衡功能检查','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',152,3619,'360000002','平衡功能评定','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',152,3620,'360000003','康复评定','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',152,3621,'360000004','日常生活活动能力评定','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',152,3622,'360000005','等速肌力测定','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',152,3623,'360000006','手功能评定','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',152,3624,'360000007','疲劳度测定','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',152,3625,'360000008','步态分析检查','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',152,3626,'360000009','语言能力评定','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',152,3627,'360000010','失语症检查','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',152,3628,'360000011','口吃检查','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',152,3629,'360000012','吞咽功能障碍评定','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',152,3630,'360000013','纯音听力检查','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual;
Insert Into 诊疗项目目录(类别,分类id,id,编码,名称,计算单位,标本部位,计算方式,单独应用,执行频率,适用性别,组合项目,操作类型,执行安排,执行科室,服务对象,计价性质,建档时间,撤档时间)
Select 'D',152,3631,'360000014','认知知觉功能检查','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',152,3632,'360000015','记忆力评定','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',152,3633,'360000016','失认、失用评定','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',152,3634,'360000017','职业能力评定','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'D',152,3635,'360000018','记忆广度检查','次','',3,1,1,0,0,'其他',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',152,3636,'360000019','运动疗法','次','',3,1,1,0,0,'0',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',152,3637,'360000020','轮椅功能训练','次','',3,1,1,0,0,'0',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',152,3638,'360000021','电动起立床训练','次','',3,1,1,0,0,'0',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',152,3639,'360000022','平衡功能训练','次','',3,1,1,0,0,'0',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',152,3640,'360000023','手功能训练','次','',3,1,1,0,0,'0',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',152,3641,'360000024','关节松动术','次','',3,1,1,0,0,'0',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',152,3642,'360000025','有氧训练','次','',3,1,1,0,0,'0',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',152,3643,'360000026','文体训练','次','',3,1,1,0,0,'0',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',152,3644,'360000027','引导式教育训练','次','',3,1,1,0,0,'0',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',152,3645,'360000028','等速肌力训练','次','',3,1,1,0,0,'0',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',152,3646,'360000029','作业疗法','次','',3,1,1,0,0,'0',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',152,3647,'360000030','职业功能训练','次','',3,1,1,0,0,'0',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',152,3648,'360000031','口吃训练','次','',3,1,1,0,0,'0',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',152,3649,'360000032','语言训练','次','',3,1,1,0,0,'0',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',152,3650,'360000033','儿童听力障碍语言训练','次','',3,1,1,0,0,'0',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',152,3651,'360000034','构音障碍训练','次','',3,1,1,0,0,'0',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',152,3652,'360000035','吞咽功能障碍训练','次','',3,1,1,0,0,'0',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',152,3653,'360000036','认知知觉功能障碍训练','次','',3,1,1,0,0,'0',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',152,3654,'360000037','社区康复测查','次','',3,1,1,0,0,'0',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',152,3655,'360000038','偏瘫肢体综合训练','次','',3,1,1,0,0,'0',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',152,3656,'360000039','脑瘫肢体综合训练','次','',3,1,1,0,0,'0',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',154,3657,'410100001','水煎','次','',3,0,0,0,0,'3',0,0,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',154,3658,'410100002','水煎(煎药机)','次','',3,0,0,0,0,'3',0,0,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',154,3659,'410100003','研粉','次','',3,0,0,0,0,'3',0,0,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',154,3660,'410100004','制丸','次','',3,0,0,0,0,'3',0,0,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',154,3661,'410100005','蜜炼','次','',3,0,0,0,0,'3',0,0,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',154,3662,'410100006','制膏','次','',3,0,0,0,0,'3',0,0,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',154,3663,'410100007','口服','次','',3,0,0,0,0,'4',0,0,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',154,3664,'410100008','开水冲服','次','',3,0,0,0,0,'4',0,0,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',154,3665,'410100009','嚼服','次','',3,0,0,0,0,'4',0,0,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',154,3666,'410100010','伴食服','次','',3,0,0,0,0,'4',0,0,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',154,3667,'410100011','酒服','次','',3,0,0,0,0,'4',0,0,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',154,3668,'410100012','外用','次','',3,0,0,0,0,'4',0,0,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',154,3669,'410100013','外搽','次','',3,0,0,0,0,'4',0,0,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',154,3670,'410100014','外敷','次','',3,0,0,0,0,'4',0,0,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',154,3671,'410100015','熏洗','次','',3,0,0,0,0,'4',0,0,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',154,3672,'410100016','熏蒸','次','',3,0,0,0,0,'4',0,0,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',154,3673,'410100017','药浴','次','',3,0,0,0,0,'4',0,0,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',154,3674,'410100018','滴鼻','次','',3,0,0,0,0,'4',0,0,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',154,3675,'410100019','保留灌肠','次','',3,0,0,0,0,'4',0,0,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',154,3676,'410100020','直肠灌注','次','',3,0,0,0,0,'4',0,0,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',154,3677,'410100021','耳咽吹粉','次','',3,0,0,0,0,'4',0,0,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',155,3678,'410200001','中药化腐清创术','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',155,3679,'410200002','赘生物中药腐蚀','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',155,3680,'410200003','挑治','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',155,3681,'410200004','割治','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',155,3682,'410200005','针刺治疗','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',155,3683,'410200006','埋针治疗','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',155,3684,'410200007','电针治疗','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',155,3685,'410200008','微波针治疗','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',155,3686,'410200009','激光针治疗','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',155,3687,'410200010','磁热疗法','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',155,3688,'410200011','放血疗法','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',155,3689,'410200012','穴位注射','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',155,3690,'410200013','子午流注开穴法','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',155,3691,'410200014','灸治','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',155,3692,'410200015','拔罐','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',155,3693,'410200016','全身推拿','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',155,3694,'410200017','落枕推拿','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',155,3695,'410200018','其他推拿','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',155,3696,'410200019','小儿捏脊','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',155,3697,'410200020','人工按摩','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',155,3698,'410200021','按摩器按摩','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',155,3699,'410200022','药棒穴位按摩','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',155,3700,'410200023','刮痧疗法','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',155,3701,'410200024','烫熨疗法','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',156,3702,'410300001','骨折手法修复','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',156,3703,'410300002','骨折撬拨复位术','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',156,3704,'410300003','骨折经皮钳夹复位术','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',156,3705,'410300004','关节脱位手法修复','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',156,3706,'410300005','骨折外固定架固定术','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',156,3707,'410300006','骨折夹板外固定术','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',156,3708,'410300007','关节错缝术','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',156,3709,'410300008','麻醉下腰突症大手法治疗','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',156,3710,'410300009','外固定架使用','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',157,3711,'410400001','直肠脱出复位','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',157,3712,'410400002','直肠周围硬化剂治疗','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',157,3713,'410400003','内痔硬化剂注射治疗(枯痔治疗)','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',157,3714,'410400004','高位、复杂肛瘘挂线治疗','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',157,3715,'410400005','肛瘘封堵术','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',157,3716,'410400006','白内障针拔术','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',157,3717,'410400007','白内障针拔吸出术','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',157,3718,'410400008','白内障针拔套出术','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',157,3719,'410400009','眼结膜囊穴位注射','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',157,3720,'410400010','小针刀治疗','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',157,3721,'410400011','红皮病清消术','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',157,3722,'410400012','扁桃体烙法','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',157,3723,'410400013','药线引流疗法','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',157,3724,'410400014','医疗气功疗法','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual Union All
Select 'E',157,3725,'410400015','辨证施膳指导','次','',3,1,1,0,0,'',0,1,3,1,trunc(sysdate)-30,to_date('3000-1-1','yyyy-mm-dd') From Dual;

--  3.  诊疗项目别名
Insert Into 诊疗项目别名(诊疗项目id,名称,性质,简码,码类)
Select id, 名称, 1, zlSpellCode(名称), 1
  From 诊疗项目目录
 Where 类别 >'A';
Insert Into 诊疗项目别名(诊疗项目id,名称,性质,简码,码类)
Select id,名称,1,zlWBCode(名称),2 
  From 诊疗项目目录 I
 Where 类别 >'A';

Insert Into 诊疗项目别名(诊疗项目id,名称,性质,简码,码类)
Select 158,'益萨林皮试',9,'YSLPS',1 From Dual Union All
Select 158,'益萨林皮试',9,'UASHY',2 From Dual Union All
Select 160,'哌拉西林钠+他唑巴坦皮试',9,'PLXLN+TZBT',1 From Dual Union All
Select 160,'哌拉西林钠+他唑巴坦皮试',9,'KRSSQ+WKCF',2 From Dual Union All
Select 161,'哌拉西林+他巴唑坦皮试',9,'PLXL+TBZTP',1 From Dual Union All
Select 161,'哌拉西林+他巴唑坦皮试',9,'KRSS+WCKFH',2 From Dual Union All
Select 164,'泛捷复皮试',9,'FJFPS',1 From Dual Union All
Select 164,'泛捷复皮试',9,'IRTHY',2 From Dual Union All
Select 167,'先锋必皮试',9,'XFBPS',1 From Dual Union All
Select 167,'先锋必皮试',9,'TQNHY',2 From Dual Union All
Select 168,'罗氏芬皮试',9,'LSFPS',1 From Dual Union All
Select 168,'罗氏芬皮试',9,'LQAHY',2 From Dual Union All
Select 174,'马斯平皮试',9,'MSPPS',1 From Dual Union All
Select 174,'马斯平皮试',9,'CAGHY',2 From Dual Union All
Select 176,'特子社复皮试',9,'TZSFPS',1 From Dual Union All
Select 176,'特子社复皮试',9,'TBPTHY',2 From Dual Union All
Select 179,'丽安林皮试',9,'LALPS',1 From Dual Union All
Select 179,'丽安林皮试',9,'GPSHY',2 From Dual Union All
Select 180,'特美汀皮试',9,'TMTPS',1 From Dual Union All
Select 180,'特美汀皮试',9,'TUIHY',2 From Dual Union All
Select 181,'海夫佳皮试',9,'HFJPS',1 From Dual Union All
Select 181,'海夫佳皮试',9,'IFWHY',2 From Dual Union All
Select 182,'泰能皮试',9,'TNPS',1 From Dual Union All
Select 182,'泰能皮试',9,'DCHY',2 From Dual Union All
Select 185,'强力阿莫仙皮试',9,'QLAMXPS',1 From Dual Union All
Select 185,'强力阿莫仙皮试',9,'XLBAWHY',2 From Dual Union All
Select 186,'开林皮试',9,'KLPS',1 From Dual Union All
Select 186,'开林皮试',9,'GSHY',2 From Dual Union All
Select 202,'肌注',9,'JZ',1 From Dual Union All
Select 202,'肌注',9,'EI',2 From Dual Union All
Select 203,'静注',9,'JZ',1 From Dual Union All
Select 203,'静注',9,'GI',2 From Dual Union All
Select 204,'静滴',9,'JD',1 From Dual Union All
Select 204,'静滴',9,'GI',2 From Dual Union All
Select 597,'留观',9,'LG',1 From Dual Union All
Select 597,'留观',9,'QC',2 From Dual Union All
Select 598,'转诊',9,'ZZ',1 From Dual Union All
Select 598,'转诊',9,'LY',2 From Dual Union All
Select 599,'住院治疗',9,'ZYZL',1 From Dual Union All
Select 599,'住院治疗',9,'WBIU',2 From Dual Union All
Select 601,'病危',9,'BW',1 From Dual Union All
Select 601,'病危',9,'UQ',2 From Dual Union All
Select 602,'术后医嘱',9,'SHYZ',1 From Dual Union All
Select 602,'术后医嘱',9,'SRAK',2 From Dual Union All
Select 640,'DR',9,'DR',1 From Dual Union All
Select 640,'DR',9,'DR',2 From Dual Union All
Select 641,'CR',9,'CR',1 From Dual Union All
Select 641,'CR',9,'CR',2 From Dual;

--  4.  检验报告项目：目前存放检验标本，今后随LIS取消，修改为诊疗用法用量
Insert Into 检验报告项目(id,诊疗项目id,检验标本)
Select 1,878,'手指血' From Dual Union All
Select 2,879,'静脉抗凝血' From Dual Union All
Select 3,880,'静脉抗凝血' From Dual Union All
Select 4,881,'尿液' From Dual Union All
Select 5,882,'晨尿' From Dual Union All
Select 6,883,'24h尿' From Dual Union All
Select 7,884,'晨尿' From Dual Union All
Select 8,885,'新鲜粪便' From Dual Union All
Select 9,886,'粪便' From Dual Union All
Select 10,887,'浆膜腔积液' From Dual Union All
Select 11,888,'浆膜腔积液' From Dual Union All
Select 12,889,'浆膜腔积液' From Dual Union All
Select 13,890,'浆膜腔积液' From Dual Union All
Select 14,891,'浆膜腔积液' From Dual Union All
Select 15,892,'脑脊液' From Dual Union All
Select 16,893,'血浆' From Dual Union All
Select 17,894,'血浆' From Dual Union All
Select 18,895,'血浆' From Dual Union All
Select 19,896,'血浆' From Dual Union All
Select 20,897,'血浆' From Dual Union All
Select 21,898,'痰液' From Dual Union All
Select 22,899,'精液' From Dual Union All
Select 23,900,'分泌物' From Dual Union All
Select 24,901,'分泌物' From Dual Union All
Select 25,902,'白带' From Dual Union All
Select 26,903,'前列腺液' From Dual Union All
Select 27,904,'血浆' From Dual Union All
Select 28,905,'空腹血清' From Dual Union All
Select 29,906,'血浆' From Dual Union All
Select 30,907,'血浆' From Dual Union All
Select 31,908,'胃肠粘膜' From Dual Union All
Select 32,909,'前列腺液' From Dual Union All
Select 33,910,'血浆' From Dual Union All
Select 34,911,'血浆' From Dual Union All
Select 35,912,'血浆' From Dual Union All
Select 36,913,'血浆' From Dual Union All
Select 37,914,'血浆' From Dual Union All
Select 38,915,'尿液' From Dual Union All
Select 39,916,'血浆' From Dual Union All
Select 40,917,'血清抗凝血' From Dual Union All
Select 41,918,'动脉抗凝血' From Dual Union All
Select 42,919,'动脉抗凝血' From Dual Union All
Select 43,920,'动脉抗凝血' From Dual Union All
Select 44,921,'空腹血清' From Dual Union All
Select 45,922,'空腹血清' From Dual Union All
Select 46,923,'空腹血清' From Dual Union All
Select 47,924,'空腹血清' From Dual Union All
Select 48,925,'空腹血清' From Dual Union All
Select 49,926,'空腹血清' From Dual Union All
Select 50,927,'空腹血清' From Dual Union All
Select 51,928,'空腹血清' From Dual Union All
Select 52,929,'空腹血浆' From Dual Union All
Select 53,929,'空腹血清' From Dual Union All
Select 54,930,'空腹血清' From Dual Union All
Select 55,931,'空腹血清' From Dual Union All
Select 56,932,'晨尿' From Dual Union All
Select 57,933,'空腹血清' From Dual Union All
Select 58,934,'空腹血清' From Dual Union All
Select 59,935,'空腹血清' From Dual Union All
Select 60,936,'空腹血浆' From Dual Union All
Select 61,937,'空腹血浆' From Dual Union All
Select 62,938,'空腹血浆' From Dual Union All
Select 63,939,'空腹血浆' From Dual Union All
Select 64,940,'空腹血浆' From Dual Union All
Select 65,941,'空腹血浆' From Dual Union All
Select 66,942,'24h尿' From Dual Union All
Select 67,943,'24h尿' From Dual Union All
Select 68,944,'晨尿' From Dual Union All
Select 69,945,'脑脊液' From Dual Union All
Select 70,946,'脑脊液' From Dual Union All
Select 71,947,'骨髓' From Dual Union All
Select 72,948,'浆膜腔积液' From Dual Union All
Select 73,949,'浆膜腔积液' From Dual Union All
Select 74,950,'血液' From Dual Union All
Select 75,951,'血液' From Dual Union All
Select 76,952,'血液' From Dual Union All
Select 77,953,'尿液' From Dual Union All
Select 78,954,'尿液' From Dual Union All
Select 79,955,'空腹血清' From Dual Union All
Select 80,956,'血液' From Dual Union All
Select 81,957,'空腹血清' From Dual Union All
Select 82,958,'血液' From Dual Union All
Select 83,959,'空腹血清' From Dual Union All
Select 84,960,'血液' From Dual Union All
Select 85,961,'血液' From Dual Union All
Select 86,962,'空腹血清' From Dual Union All
Select 87,963,'血清' From Dual Union All
Select 88,964,'空腹血清' From Dual Union All
Select 89,965,'空腹血清' From Dual Union All
Select 90,966,'空腹血清' From Dual Union All
Select 91,967,'空腹血清' From Dual Union All
Select 92,968,'空腹血清' From Dual Union All
Select 93,969,'空腹血清' From Dual Union All
Select 94,970,'血清' From Dual Union All
Select 95,971,'血清' From Dual Union All
Select 96,972,'脑脊液' From Dual Union All
Select 97,973,'血清' From Dual Union All
Select 98,974,'血清' From Dual Union All
Select 99,975,'血清' From Dual Union All
Select 100,976,'血清' From Dual Union All
Select 101,978,'血清' From Dual Union All
Select 102,979,'血清' From Dual Union All
Select 103,980,'血清' From Dual Union All
Select 104,981,'血清' From Dual Union All
Select 105,983,'血清' From Dual Union All
Select 106,984,'尿液' From Dual Union All
Select 107,985,'血清' From Dual Union All
Select 108,986,'血清' From Dual Union All
Select 109,987,'血清' From Dual Union All
Select 110,988,'血清' From Dual;
Insert Into 检验报告项目(id,诊疗项目id,检验标本)
Select 111,989,'血清' From Dual Union All
Select 112,990,'血浆' From Dual Union All
Select 113,991,'血清' From Dual Union All
Select 114,992,'血清' From Dual Union All
Select 115,993,'血清' From Dual Union All
Select 116,994,'血清' From Dual Union All
Select 117,995,'血清' From Dual Union All
Select 118,996,'血清' From Dual Union All
Select 119,997,'血清' From Dual Union All
Select 120,998,'血清' From Dual Union All
Select 121,999,'血清' From Dual Union All
Select 122,1000,'血清' From Dual Union All
Select 123,1001,'血清' From Dual Union All
Select 124,1002,'血清' From Dual Union All
Select 125,1003,'血清' From Dual Union All
Select 126,1004,'血清' From Dual Union All
Select 127,1005,'血清' From Dual Union All
Select 128,1006,'血清' From Dual Union All
Select 129,1007,'血清' From Dual Union All
Select 130,1008,'血清' From Dual Union All
Select 131,1009,'血清' From Dual Union All
Select 132,1010,'血清' From Dual Union All
Select 133,1011,'血清' From Dual Union All
Select 134,1012,'血清' From Dual Union All
Select 135,1013,'血清' From Dual Union All
Select 136,1014,'血清' From Dual Union All
Select 137,1015,'血清' From Dual Union All
Select 138,1016,'血清' From Dual Union All
Select 139,1018,'分泌物' From Dual Union All
Select 140,1019,'浆膜腔积液' From Dual Union All
Select 141,1020,'胸水' From Dual Union All
Select 142,1021,'痰液' From Dual Union All
Select 143,1022,'痰液' From Dual Union All
Select 144,1023,'分泌物' From Dual Union All
Select 145,1024,'分泌物' From Dual Union All
Select 146,1025,'痰液' From Dual Union All
Select 147,1026,'血液' From Dual Union All
Select 148,1028,'脑脊液' From Dual Union All
Select 149,1029,'尿液' From Dual Union All
Select 150,1030,'粪便' From Dual Union All
Select 151,1031,'胆汁' From Dual Union All
Select 152,1032,'分泌物' From Dual Union All
Select 153,1033,'痰液' From Dual Union All
Select 154,1034,'咽拭子' From Dual Union All
Select 155,1038,'血液' From Dual Union All
Select 156,1043,'粪便' From Dual Union All
Select 157,1044,'胆汁' From Dual Union All
Select 158,1045,'脑脊液' From Dual Union All
Select 159,1046,'分泌物' From Dual Union All
Select 160,1046,'宫腔分泌物' From Dual Union All
Select 161,1046,'前列腺液分泌物' From Dual Union All
Select 162,1046,'阴道分泌物' From Dual Union All
Select 163,1047,'血液' From Dual Union All
Select 164,1048,'痰液' From Dual Union All
Select 165,1050,'咽拭子' From Dual Union All
Select 166,1051,'尿液' From Dual Union All
Select 167,1052,'痰液' From Dual Union All
Select 168,1053,'病理材料' From Dual Union All
Select 169,1054,'病理材料' From Dual Union All
Select 170,1056,'病理材料' From Dual Union All
Select 171,1057,'血液' From Dual Union All
Select 172,1059,'病理材料' From Dual Union All
Select 173,1066,'血液' From Dual Union All
Select 174,1067,'胃肠粘膜' From Dual Union All
Select 175,1069,'胃液' From Dual Union All
Select 176,1072,'病理材料' From Dual Union All
Select 177,1075,'病理材料' From Dual Union All
Select 178,1076,'血浆' From Dual Union All
Select 179,1077,'新鲜粪便' From Dual Union All
Select 180,1078,'血液' From Dual Union All
Select 181,1079,'尿液' From Dual Union All
Select 182,1080,'痰液' From Dual Union All
Select 183,1081,'血液' From Dual Union All
Select 184,1082,'新鲜粪便' From Dual Union All
Select 185,1083,'新鲜粪便' From Dual;

--  5.  诊疗收费关系
Insert Into 诊疗收费关系(诊疗项目id, 收费项目id, 收费数量, 固有对照)
Select 157,ID,1,0 From 收费项目目录 Where ID=124 Union All
Select 2,ID,1,0 From 收费项目目录 Where ID=74 Union All
Select 3,ID,1,0 From 收费项目目录 Where ID=75 Union All
Select 4,ID,1,0 From 收费项目目录 Where ID=76 Union All
Select 156,ID,1,0 From 收费项目目录 Where ID=124 Union All
Select 437,ID,1,0 From 收费项目目录 Where ID=3084 Union All
Select 3443,ID,1,0 From 收费项目目录 Where ID=3086 Union All
Select 3444,ID,1,0 From 收费项目目录 Where ID=3087 Union All
Select 3446,ID,1,0 From 收费项目目录 Where ID=3108 Union All
Select 2152,ID,1,0 From 收费项目目录 Where ID=3119 Union All
Select 467,ID,1,0 From 收费项目目录 Where ID=3044 Union All
Select 418,ID,1,0 From 收费项目目录 Where ID=3048 Union All
Select 463,ID,1,0 From 收费项目目录 Where ID=3050 Union All
Select 3439,ID,1,0 From 收费项目目录 Where ID=3056 Union All
Select 3431,ID,1,0 From 收费项目目录 Where ID=2979 Union All
Select 3432,ID,1,0 From 收费项目目录 Where ID=2980 Union All
Select 3435,ID,1,0 From 收费项目目录 Where ID=3010 Union All
Select 3436,ID,1,0 From 收费项目目录 Where ID=3011 Union All
Select 3427,ID,1,0 From 收费项目目录 Where ID=2959 Union All
Select 918,ID,1,0 From 收费项目目录 Where ID=2920 Union All
Select 459,ID,1,0 From 收费项目目录 Where ID=2945 Union All
Select 1876,ID,1,0 From 收费项目目录 Where ID=2953 Union All
Select 302,ID,1,0 From 收费项目目录 Where ID=2579 Union All
Select 435,ID,1,0 From 收费项目目录 Where ID=2582 Union All
Select 1677,ID,1,0 From 收费项目目录 Where ID=2592 Union All
Select 1503,ID,1,0 From 收费项目目录 Where ID=2453 Union All
Select 1504,ID,1,0 From 收费项目目录 Where ID=2457 Union All
Select 1625,ID,1,0 From 收费项目目录 Where ID=2494 Union All
Select 225,ID,1,0 From 收费项目目录 Where ID=2444 Union All
Select 3412,ID,1,0 From 收费项目目录 Where ID=2213 Union All
Select 3413,ID,1,0 From 收费项目目录 Where ID=2216 Union All
Select 3414,ID,1,0 From 收费项目目录 Where ID=2223 Union All
Select 3415,ID,1,0 From 收费项目目录 Where ID=2225 Union All
Select 3417,ID,1,0 From 收费项目目录 Where ID=2239 Union All
Select 1281,ID,1,0 From 收费项目目录 Where ID=2241 Union All
Select 3419,ID,1,0 From 收费项目目录 Where ID=2250 Union All
Select 3420,ID,1,0 From 收费项目目录 Where ID=2253 Union All
Select 3422,ID,1,0 From 收费项目目录 Where ID=2257 Union All
Select 3409,ID,1,0 From 收费项目目录 Where ID=2261 Union All
Select 1104,ID,1,0 From 收费项目目录 Where ID=2119 Union All
Select 1105,ID,1,0 From 收费项目目录 Where ID=2120 Union All
Select 1106,ID,1,0 From 收费项目目录 Where ID=2128 Union All
Select 1107,ID,1,0 From 收费项目目录 Where ID=2133 Union All
Select 1108,ID,1,0 From 收费项目目录 Where ID=2137 Union All
Select 1110,ID,1,0 From 收费项目目录 Where ID=2139 Union All
Select 1116,ID,1,0 From 收费项目目录 Where ID=2144 Union All
Select 1117,ID,1,0 From 收费项目目录 Where ID=2145 Union All
Select 1121,ID,1,0 From 收费项目目录 Where ID=2151 Union All
Select 1102,ID,1,0 From 收费项目目录 Where ID=2107 Union All
Select 1103,ID,1,0 From 收费项目目录 Where ID=2116 Union All
Select 1088,ID,1,0 From 收费项目目录 Where ID=2042 Union All
Select 1089,ID,1,0 From 收费项目目录 Where ID=2043 Union All
Select 1090,ID,1,0 From 收费项目目录 Where ID=2045 Union All
Select 1091,ID,1,0 From 收费项目目录 Where ID=2047 Union All
Select 1092,ID,1,0 From 收费项目目录 Where ID=2048 Union All
Select 1095,ID,1,0 From 收费项目目录 Where ID=2072 Union All
Select 1074,ID,1,0 From 收费项目目录 Where ID=1978 Union All
Select 1073,ID,1,0 From 收费项目目录 Where ID=1986 Union All
Select 1072,ID,1,0 From 收费项目目录 Where ID=1989 Union All
Select 1096,ID,1,0 From 收费项目目录 Where ID=2029 Union All
Select 1098,ID,1,0 From 收费项目目录 Where ID=2031 Union All
Select 1026,ID,1,0 From 收费项目目录 Where ID=1908 Union All
Select 1066,ID,1,0 From 收费项目目录 Where ID=1932 Union All
Select 1067,ID,1,0 From 收费项目目录 Where ID=1936 Union All
Select 1055,ID,1,0 From 收费项目目录 Where ID=1947 Union All
Select 1052,ID,1,0 From 收费项目目录 Where ID=1949 Union All
Select 1052,ID,1,0 From 收费项目目录 Where ID=1950 Union All
Select 1094,ID,1,0 From 收费项目目录 Where ID=7290 Union All
Select 948,ID,1,0 From 收费项目目录 Where ID=7292 Union All
Select 665,ID,1,0 From 收费项目目录 Where ID=7312 Union All
Select 667,ID,1,0 From 收费项目目录 Where ID=7315 Union All
Select 674,ID,1,0 From 收费项目目录 Where ID=7317 Union All
Select 973,ID,1,0 From 收费项目目录 Where ID=1803 Union All
Select 974,ID,1,0 From 收费项目目录 Where ID=1804 Union All
Select 957,ID,1,0 From 收费项目目录 Where ID=1089 Union All
Select 882,ID,1,0 From 收费项目目录 Where ID=833 Union All
Select 3589,ID,1,0 From 收费项目目录 Where ID=769 Union All
Select 3576,ID,1,0 From 收费项目目录 Where ID=729 Union All
Select 3577,ID,1,0 From 收费项目目录 Where ID=734 Union All
Select 3579,ID,1,0 From 收费项目目录 Where ID=742 Union All
Select 3580,ID,1,0 From 收费项目目录 Where ID=743 Union All
Select 3581,ID,1,0 From 收费项目目录 Where ID=744 Union All
Select 3582,ID,1,0 From 收费项目目录 Where ID=748 Union All
Select 3584,ID,1,0 From 收费项目目录 Where ID=750 Union All
Select 857,ID,1,0 From 收费项目目录 Where ID=483 Union All
Select 870,ID,1,0 From 收费项目目录 Where ID=496 Union All
Select 862,ID,1,0 From 收费项目目录 Where ID=442 Union All
Select 860,ID,1,0 From 收费项目目录 Where ID=443 Union All
Select 873,ID,1,0 From 收费项目目录 Where ID=410 Union All
Select 658,ID,1,0 From 收费项目目录 Where ID=334 Union All
Select 665,ID,1,0 From 收费项目目录 Where ID=344 Union All
Select 676,ID,1,0 From 收费项目目录 Where ID=351 Union All
Select 643,ID,1,0 From 收费项目目录 Where ID=318 Union All
Select 673,ID,1,0 From 收费项目目录 Where ID=315 Union All
Select 670,ID,1,0 From 收费项目目录 Where ID=319 Union All
Select 452,ID,1,0 From 收费项目目录 Where ID=174 Union All
Select 451,ID,1,0 From 收费项目目录 Where ID=176 Union All
Select 450,ID,1,0 From 收费项目目录 Where ID=178 Union All
Select 370,ID,1,0 From 收费项目目录 Where ID=195 Union All
Select 389,ID,1,0 From 收费项目目录 Where ID=199 Union All
Select 404,ID,1,0 From 收费项目目录 Where ID=204 Union All
Select 329,ID,1,0 From 收费项目目录 Where ID=217 Union All
Select 232,ID,1,0 From 收费项目目录 Where ID=221 Union All
Select 288,ID,1,0 From 收费项目目录 Where ID=223 Union All
Select 3675,ID,1,0 From 收费项目目录 Where ID=223 Union All
Select 321,ID,1,0 From 收费项目目录 Where ID=227 Union All
Select 364,ID,1,0 From 收费项目目录 Where ID=229 Union All
Select 382,ID,1,0 From 收费项目目录 Where ID=233 Union All
Select 3497,ID,1,0 From 收费项目目录 Where ID=102 Union All
Select 202,ID,1,0 From 收费项目目录 Where ID=120;
Insert Into 诊疗收费关系(诊疗项目id, 收费项目id, 收费数量, 固有对照)
Select 201,ID,1,0 From 收费项目目录 Where ID=124 Union All
Select 203,ID,1,0 From 收费项目目录 Where ID=128 Union All
Select 219,ID,1,0 From 收费项目目录 Where ID=133 Union All
Select 1148,ID,1,0 From 收费项目目录 Where ID=160 Union All
Select 1,ID,1,0 From 收费项目目录 Where ID=73 Union All
Select 853,ID,1,0 From 收费项目目录 Where ID=7203 Union All
Select 2214,ID,1,0 From 收费项目目录 Where ID=7207 Union All
Select 881,ID,1,0 From 收费项目目录 Where ID=7211 Union All
Select 3713,ID,1,0 From 收费项目目录 Where ID=6933 Union All
Select 3715,ID,1,0 From 收费项目目录 Where ID=6945 Union All
Select 3719,ID,1,0 From 收费项目目录 Where ID=6950 Union All
Select 3720,ID,1,0 From 收费项目目录 Where ID=6951 Union All
Select 3710,ID,1,0 From 收费项目目录 Where ID=6821 Union All
Select 3707,ID,1,0 From 收费项目目录 Where ID=6817 Union All
Select 3708,ID,1,0 From 收费项目目录 Where ID=6819 Union All
Select 3687,ID,1,0 From 收费项目目录 Where ID=6866 Union All
Select 3688,ID,1,0 From 收费项目目录 Where ID=6867 Union All
Select 3689,ID,1,0 From 收费项目目录 Where ID=6870 Union All
Select 3690,ID,1,0 From 收费项目目录 Where ID=6874 Union All
Select 3706,ID,1,0 From 收费项目目录 Where ID=6815 Union All
Select 3637,ID,1,0 From 收费项目目录 Where ID=6745 Union All
Select 3638,ID,1,0 From 收费项目目录 Where ID=6746 Union All
Select 3639,ID,1,0 From 收费项目目录 Where ID=6747 Union All
Select 3640,ID,1,0 From 收费项目目录 Where ID=6748 Union All
Select 3642,ID,1,0 From 收费项目目录 Where ID=6752 Union All
Select 3643,ID,1,0 From 收费项目目录 Where ID=6753 Union All
Select 3644,ID,1,0 From 收费项目目录 Where ID=6754 Union All
Select 3645,ID,1,0 From 收费项目目录 Where ID=6755 Union All
Select 3646,ID,1,0 From 收费项目目录 Where ID=6756 Union All
Select 3647,ID,1,0 From 收费项目目录 Where ID=6757 Union All
Select 3648,ID,1,0 From 收费项目目录 Where ID=6758 Union All
Select 3650,ID,1,0 From 收费项目目录 Where ID=6760 Union All
Select 3651,ID,1,0 From 收费项目目录 Where ID=6761 Union All
Select 3652,ID,1,0 From 收费项目目录 Where ID=6762 Union All
Select 3653,ID,1,0 From 收费项目目录 Where ID=6763 Union All
Select 3620,ID,1,0 From 收费项目目录 Where ID=6764 Union All
Select 3655,ID,1,0 From 收费项目目录 Where ID=6765 Union All
Select 3656,ID,1,0 From 收费项目目录 Where ID=6766 Union All
Select 3680,ID,1,0 From 收费项目目录 Where ID=6792 Union All
Select 3681,ID,1,0 From 收费项目目录 Where ID=6793 Union All
Select 3704,ID,1,0 From 收费项目目录 Where ID=6804 Union All
Select 3622,ID,1,0 From 收费项目目录 Where ID=6714 Union All
Select 3623,ID,1,0 From 收费项目目录 Where ID=6715 Union All
Select 3624,ID,1,0 From 收费项目目录 Where ID=6718 Union All
Select 3625,ID,1,0 From 收费项目目录 Where ID=6719 Union All
Select 3627,ID,1,0 From 收费项目目录 Where ID=6724 Union All
Select 3628,ID,1,0 From 收费项目目录 Where ID=6725 Union All
Select 3629,ID,1,0 From 收费项目目录 Where ID=6726 Union All
Select 3631,ID,1,0 From 收费项目目录 Where ID=6727 Union All
Select 3632,ID,1,0 From 收费项目目录 Where ID=6729 Union All
Select 3634,ID,1,0 From 收费项目目录 Where ID=6732 Union All
Select 3635,ID,1,0 From 收费项目目录 Where ID=6733 Union All
Select 3636,ID,1,0 From 收费项目目录 Where ID=6737 Union All
Select 3610,ID,1,0 From 收费项目目录 Where ID=6616 Union All
Select 3277,ID,1,0 From 收费项目目录 Where ID=6393 Union All
Select 3218,ID,1,0 From 收费项目目录 Where ID=6266 Union All
Select 3286,ID,1,0 From 收费项目目录 Where ID=6338 Union All
Select 3217,ID,1,0 From 收费项目目录 Where ID=6199 Union All
Select 3212,ID,1,0 From 收费项目目录 Where ID=6203 Union All
Select 3318,ID,1,0 From 收费项目目录 Where ID=6218 Union All
Select 3132,ID,1,0 From 收费项目目录 Where ID=6232 Union All
Select 3335,ID,1,0 From 收费项目目录 Where ID=6233 Union All
Select 3342,ID,1,0 From 收费项目目录 Where ID=6240 Union All
Select 3341,ID,1,0 From 收费项目目录 Where ID=6241 Union All
Select 3339,ID,1,0 From 收费项目目录 Where ID=6242 Union All
Select 3348,ID,1,0 From 收费项目目录 Where ID=6247 Union All
Select 3166,ID,1,0 From 收费项目目录 Where ID=6126 Union All
Select 3170,ID,1,0 From 收费项目目录 Where ID=6130 Union All
Select 1228,ID,1,0 From 收费项目目录 Where ID=6031 Union All
Select 1222,ID,1,0 From 收费项目目录 Where ID=6034 Union All
Select 3200,ID,1,0 From 收费项目目录 Where ID=5953 Union All
Select 2877,ID,1,0 From 收费项目目录 Where ID=5817 Union All
Select 3485,ID,1,0 From 收费项目目录 Where ID=5836 Union All
Select 2970,ID,1,0 From 收费项目目录 Where ID=5843 Union All
Select 2949,ID,1,0 From 收费项目目录 Where ID=5851 Union All
Select 2936,ID,1,0 From 收费项目目录 Where ID=5854 Union All
Select 2920,ID,1,0 From 收费项目目录 Where ID=5873 Union All
Select 2778,ID,1,0 From 收费项目目录 Where ID=5730 Union All
Select 2784,ID,1,0 From 收费项目目录 Where ID=5736 Union All
Select 2791,ID,1,0 From 收费项目目录 Where ID=5738 Union All
Select 2786,ID,1,0 From 收费项目目录 Where ID=5745 Union All
Select 2790,ID,1,0 From 收费项目目录 Where ID=5748 Union All
Select 2795,ID,1,0 From 收费项目目录 Where ID=5752 Union All
Select 2824,ID,1,0 From 收费项目目录 Where ID=5782 Union All
Select 2811,ID,1,0 From 收费项目目录 Where ID=5783 Union All
Select 2807,ID,1,0 From 收费项目目录 Where ID=5784 Union All
Select 2717,ID,1,0 From 收费项目目录 Where ID=5631 Union All
Select 2665,ID,1,0 From 收费项目目录 Where ID=5633 Union All
Select 2674,ID,1,0 From 收费项目目录 Where ID=5653 Union All
Select 2707,ID,1,0 From 收费项目目录 Where ID=5672 Union All
Select 2706,ID,1,0 From 收费项目目录 Where ID=5674 Union All
Select 2698,ID,1,0 From 收费项目目录 Where ID=5691 Union All
Select 2587,ID,1,0 From 收费项目目录 Where ID=5596 Union All
Select 2589,ID,1,0 From 收费项目目录 Where ID=5597 Union All
Select 2613,ID,1,0 From 收费项目目录 Where ID=5615 Union All
Select 2646,ID,1,0 From 收费项目目录 Where ID=5616 Union All
Select 2622,ID,1,0 From 收费项目目录 Where ID=5618 Union All
Select 2523,ID,1,0 From 收费项目目录 Where ID=5544 Union All
Select 2541,ID,1,0 From 收费项目目录 Where ID=5547 Union All
Select 2560,ID,1,0 From 收费项目目录 Where ID=5559 Union All
Select 2549,ID,1,0 From 收费项目目录 Where ID=5561 Union All
Select 2606,ID,1,0 From 收费项目目录 Where ID=5581 Union All
Select 2599,ID,1,0 From 收费项目目录 Where ID=5588 Union All
Select 2597,ID,1,0 From 收费项目目录 Where ID=5593 Union All
Select 2468,ID,1,0 From 收费项目目录 Where ID=5495 Union All
Select 2522,ID,1,0 From 收费项目目录 Where ID=5543 Union All
Select 2422,ID,1,0 From 收费项目目录 Where ID=5475 Union All
Select 1968,ID,1,0 From 收费项目目录 Where ID=5483 Union All
Select 2400,ID,1,0 From 收费项目目录 Where ID=5428 Union All
Select 2406,ID,1,0 From 收费项目目录 Where ID=5458;
Insert Into 诊疗收费关系(诊疗项目id, 收费项目id, 收费数量, 固有对照)
Select 2345,ID,1,0 From 收费项目目录 Where ID=5382 Union All
Select 2208,ID,1,0 From 收费项目目录 Where ID=5329 Union All
Select 2219,ID,1,0 From 收费项目目录 Where ID=5346 Union All
Select 2250,ID,1,0 From 收费项目目录 Where ID=5349 Union All
Select 2137,ID,1,0 From 收费项目目录 Where ID=5263 Union All
Select 2142,ID,1,0 From 收费项目目录 Where ID=5268 Union All
Select 2160,ID,1,0 From 收费项目目录 Where ID=5278 Union All
Select 2164,ID,1,0 From 收费项目目录 Where ID=5279 Union All
Select 2165,ID,1,0 From 收费项目目录 Where ID=5280 Union All
Select 2161,ID,1,0 From 收费项目目录 Where ID=5281 Union All
Select 2482,ID,1,0 From 收费项目目录 Where ID=5282 Union All
Select 2173,ID,1,0 From 收费项目目录 Where ID=5292 Union All
Select 1907,ID,1,0 From 收费项目目录 Where ID=5144 Union All
Select 1903,ID,1,0 From 收费项目目录 Where ID=5118 Union All
Select 2107,ID,1,0 From 收费项目目录 Where ID=5048 Union All
Select 2040,ID,1,0 From 收费项目目录 Where ID=4999 Union All
Select 1866,ID,1,0 From 收费项目目录 Where ID=4950 Union All
Select 1825,ID,1,0 From 收费项目目录 Where ID=4908 Union All
Select 1828,ID,1,0 From 收费项目目录 Where ID=4910 Union All
Select 1823,ID,1,0 From 收费项目目录 Where ID=4918 Union All
Select 1839,ID,1,0 From 收费项目目录 Where ID=4921 Union All
Select 1840,ID,1,0 From 收费项目目录 Where ID=4923 Union All
Select 1845,ID,1,0 From 收费项目目录 Where ID=4932 Union All
Select 1847,ID,1,0 From 收费项目目录 Where ID=4933 Union All
Select 3130,ID,1,0 From 收费项目目录 Where ID=4935 Union All
Select 1889,ID,1,0 From 收费项目目录 Where ID=4845 Union All
Select 1804,ID,1,0 From 收费项目目录 Where ID=4884 Union All
Select 1791,ID,1,0 From 收费项目目录 Where ID=4820 Union All
Select 1794,ID,1,0 From 收费项目目录 Where ID=4822 Union All
Select 1226,ID,1,0 From 收费项目目录 Where ID=4619 Union All
Select 1226,ID,1,0 From 收费项目目录 Where ID=4624 Union All
Select 1681,ID,1,0 From 收费项目目录 Where ID=4630 Union All
Select 1758,ID,1,0 From 收费项目目录 Where ID=4637 Union All
Select 1756,ID,1,0 From 收费项目目录 Where ID=4639 Union All
Select 3040,ID,1,0 From 收费项目目录 Where ID=4590 Union All
Select 1727,ID,1,0 From 收费项目目录 Where ID=4541 Union All
Select 1657,ID,1,0 From 收费项目目录 Where ID=4433 Union All
Select 1239,ID,1,0 From 收费项目目录 Where ID=4333 Union All
Select 1586,ID,1,0 From 收费项目目录 Where ID=4345 Union All
Select 1594,ID,1,0 From 收费项目目录 Where ID=4360 Union All
Select 1617,ID,1,0 From 收费项目目录 Where ID=4364 Union All
Select 1423,ID,1,0 From 收费项目目录 Where ID=4184 Union All
Select 1434,ID,1,0 From 收费项目目录 Where ID=4203 Union All
Select 1443,ID,1,0 From 收费项目目录 Where ID=4210 Union All
Select 1462,ID,1,0 From 收费项目目录 Where ID=4216 Union All
Select 1454,ID,1,0 From 收费项目目录 Where ID=4218 Union All
Select 1468,ID,1,0 From 收费项目目录 Where ID=4219 Union All
Select 1472,ID,1,0 From 收费项目目录 Where ID=4226 Union All
Select 1562,ID,1,0 From 收费项目目录 Where ID=4241 Union All
Select 1510,ID,1,0 From 收费项目目录 Where ID=4244 Union All
Select 1521,ID,1,0 From 收费项目目录 Where ID=4245 Union All
Select 1539,ID,1,0 From 收费项目目录 Where ID=4248 Union All
Select 1320,ID,1,0 From 收费项目目录 Where ID=4095 Union All
Select 1312,ID,1,0 From 收费项目目录 Where ID=4103 Union All
Select 1310,ID,1,0 From 收费项目目录 Where ID=4104 Union All
Select 1309,ID,1,0 From 收费项目目录 Where ID=4105 Union All
Select 1313,ID,1,0 From 收费项目目录 Where ID=4106 Union All
Select 1316,ID,1,0 From 收费项目目录 Where ID=4113 Union All
Select 1317,ID,1,0 From 收费项目目录 Where ID=4114 Union All
Select 1355,ID,1,0 From 收费项目目录 Where ID=4118 Union All
Select 1336,ID,1,0 From 收费项目目录 Where ID=4131 Union All
Select 1962,ID,1,0 From 收费项目目录 Where ID=4056 Union All
Select 1648,ID,1,0 From 收费项目目录 Where ID=4057 Union All
Select 1964,ID,1,0 From 收费项目目录 Where ID=4058 Union All
Select 1197,ID,1,0 From 收费项目目录 Where ID=3858 Union All
Select 1149,ID,1,0 From 收费项目目录 Where ID=3838 Union All
Select 2031,ID,1,0 From 收费项目目录 Where ID=3839 Union All
Select 1135,ID,1,0 From 收费项目目录 Where ID=3804 Union All
Select 477,ID,1,0 From 收费项目目录 Where ID=3822 Union All
Select 703,ID,1,0 From 收费项目目录 Where ID=3741 Union All
Select 702,ID,1,0 From 收费项目目录 Where ID=3744 Union All
Select 3545,ID,1,0 From 收费项目目录 Where ID=3693 Union All
Select 3546,ID,1,0 From 收费项目目录 Where ID=3694 Union All
Select 3547,ID,1,0 From 收费项目目录 Where ID=3695 Union All
Select 3548,ID,1,0 From 收费项目目录 Where ID=3696 Union All
Select 3549,ID,1,0 From 收费项目目录 Where ID=3697 Union All
Select 3550,ID,1,0 From 收费项目目录 Where ID=3700 Union All
Select 3551,ID,1,0 From 收费项目目录 Where ID=3701 Union All
Select 3552,ID,1,0 From 收费项目目录 Where ID=3702 Union All
Select 3554,ID,1,0 From 收费项目目录 Where ID=3704 Union All
Select 3531,ID,1,0 From 收费项目目录 Where ID=3677 Union All
Select 3534,ID,1,0 From 收费项目目录 Where ID=3680 Union All
Select 3537,ID,1,0 From 收费项目目录 Where ID=3682 Union All
Select 3538,ID,1,0 From 收费项目目录 Where ID=3683 Union All
Select 3541,ID,1,0 From 收费项目目录 Where ID=3685 Union All
Select 3542,ID,1,0 From 收费项目目录 Where ID=3689 Union All
Select 3543,ID,1,0 From 收费项目目录 Where ID=3691 Union All
Select 3496,ID,1,0 From 收费项目目录 Where ID=3456 Union All
Select 3495,ID,1,0 From 收费项目目录 Where ID=3470 Union All
Select 3177,ID,1,0 From 收费项目目录 Where ID=3481 Union All
Select 442,ID,1,0 From 收费项目目录 Where ID=3482 Union All
Select 220,ID,1,0 From 收费项目目录 Where ID=3495 Union All
Select 3491,ID,1,0 From 收费项目目录 Where ID=3405 Union All
Select 3493,ID,1,0 From 收费项目目录 Where ID=3407 Union All
Select 140,ID,1,0 From 收费项目目录 Where ID=3408 Union All
Select 3494,ID,1,0 From 收费项目目录 Where ID=3410 Union All
Select 3483,ID,1,0 From 收费项目目录 Where ID=3376 Union All
Select 3484,ID,1,0 From 收费项目目录 Where ID=3379 Union All
Select 3470,ID,1,0 From 收费项目目录 Where ID=3285 Union All
Select 3472,ID,1,0 From 收费项目目录 Where ID=3304 Union All
Select 3473,ID,1,0 From 收费项目目录 Where ID=3305 Union All
Select 3476,ID,1,0 From 收费项目目录 Where ID=3317 Union All
Select 3477,ID,1,0 From 收费项目目录 Where ID=3320 Union All
Select 3463,ID,1,0 From 收费项目目录 Where ID=3322 Union All
Select 2037,ID,1,0 From 收费项目目录 Where ID=3322 Union All
Select 3464,ID,1,0 From 收费项目目录 Where ID=3325 Union All
Select 3465,ID,1,0 From 收费项目目录 Where ID=3326 Union All
Select 295,ID,1,0 From 收费项目目录 Where ID=3328 Union All
Select 3478,ID,1,0 From 收费项目目录 Where ID=3347 Union All
Select 3482,ID,1,0 From 收费项目目录 Where ID=3354;
Insert Into 诊疗收费关系(诊疗项目id, 收费项目id, 收费数量, 固有对照)
Select 3490,ID,1,0 From 收费项目目录 Where ID=3359 Union All
Select 2927,ID,1,0 From 收费项目目录 Where ID=3362 Union All
Select 2713,ID,1,0 From 收费项目目录 Where ID=3280 Union All
Select 3453,ID,1,0 From 收费项目目录 Where ID=3207 Union All
Select 3454,ID,1,0 From 收费项目目录 Where ID=3209 Union All
Select 2386,ID,1,0 From 收费项目目录 Where ID=3209 Union All
Select 3455,ID,1,0 From 收费项目目录 Where ID=3210 Union All
Select 428,ID,1,0 From 收费项目目录 Where ID=3218 Union All
Select 3461,ID,1,0 From 收费项目目录 Where ID=3228 Union All
Select 3462,ID,1,0 From 收费项目目录 Where ID=3230 Union All
Select 2546,ID,1,0 From 收费项目目录 Where ID=3230 Union All
Select 3463,ID,1,0 From 收费项目目录 Where ID=3260 Union All
Select 2037,ID,1,0 From 收费项目目录 Where ID=3260 Union All
Select 3464,ID,1,0 From 收费项目目录 Where ID=3263 Union All
Select 3465,ID,1,0 From 收费项目目录 Where ID=3264 Union All
Select 295,ID,1,0 From 收费项目目录 Where ID=3265 Union All
Select 138,ID,1,0 From 收费项目目录 Where ID=3270 Union All
Select 2617,ID,1,0 From 收费项目目录 Where ID=3276 Union All
Select 2590,ID,1,0 From 收费项目目录 Where ID=3277 Union All
Select 2246,ID,1,0 From 收费项目目录 Where ID=3184 Union All
Select 3452,ID,1,0 From 收费项目目录 Where ID=3185 Union All
Select 2247,ID,1,0 From 收费项目目录 Where ID=3185 Union All
Select 3447,ID,1,0 From 收费项目目录 Where ID=3136 Union All
Select 158,ID,1,0 From 收费项目目录 Where ID=124 Union All
Select 159,ID,1,0 From 收费项目目录 Where ID=124 Union All
Select 160,ID,1,0 From 收费项目目录 Where ID=124 Union All
Select 161,ID,1,0 From 收费项目目录 Where ID=124 Union All
Select 162,ID,1,0 From 收费项目目录 Where ID=124 Union All
Select 163,ID,1,0 From 收费项目目录 Where ID=124 Union All
Select 164,ID,1,0 From 收费项目目录 Where ID=124 Union All
Select 165,ID,1,0 From 收费项目目录 Where ID=124 Union All
Select 166,ID,1,0 From 收费项目目录 Where ID=124 Union All
Select 167,ID,1,0 From 收费项目目录 Where ID=124 Union All
Select 168,ID,1,0 From 收费项目目录 Where ID=124 Union All
Select 169,ID,1,0 From 收费项目目录 Where ID=124 Union All
Select 170,ID,1,0 From 收费项目目录 Where ID=124 Union All
Select 171,ID,1,0 From 收费项目目录 Where ID=124 Union All
Select 172,ID,1,0 From 收费项目目录 Where ID=124 Union All
Select 173,ID,1,0 From 收费项目目录 Where ID=124 Union All
Select 174,ID,1,0 From 收费项目目录 Where ID=124 Union All
Select 175,ID,1,0 From 收费项目目录 Where ID=124 Union All
Select 176,ID,1,0 From 收费项目目录 Where ID=124 Union All
Select 177,ID,1,0 From 收费项目目录 Where ID=124 Union All
Select 178,ID,1,0 From 收费项目目录 Where ID=124 Union All
Select 179,ID,1,0 From 收费项目目录 Where ID=124 Union All
Select 180,ID,1,0 From 收费项目目录 Where ID=124 Union All
Select 181,ID,1,0 From 收费项目目录 Where ID=124 Union All
Select 182,ID,1,0 From 收费项目目录 Where ID=124 Union All
Select 183,ID,1,0 From 收费项目目录 Where ID=124 Union All
Select 184,ID,1,0 From 收费项目目录 Where ID=124 Union All
Select 185,ID,1,0 From 收费项目目录 Where ID=124 Union All
Select 186,ID,1,0 From 收费项目目录 Where ID=124 Union All
Select 187,ID,1,0 From 收费项目目录 Where ID=124 Union All
Select 188,ID,1,0 From 收费项目目录 Where ID=124 Union All
Select 189,ID,1,0 From 收费项目目录 Where ID=124 Union All
Select 190,ID,1,0 From 收费项目目录 Where ID=124 Union All
Select 191,ID,1,0 From 收费项目目录 Where ID=124 Union All
Select 192,ID,1,0 From 收费项目目录 Where ID=124 Union All
Select 193,ID,1,0 From 收费项目目录 Where ID=124 Union All
Select 194,ID,1,0 From 收费项目目录 Where ID=124 Union All
Select 195,ID,1,0 From 收费项目目录 Where ID=124 Union All
Select 196,ID,1,0 From 收费项目目录 Where ID=124 Union All
Select 197,ID,1,0 From 收费项目目录 Where ID=124 Union All
Select 198,ID,1,0 From 收费项目目录 Where ID=124 Union All
Select 199,ID,1,0 From 收费项目目录 Where ID=124 Union All
Select 200,ID,1,0 From 收费项目目录 Where ID=124;