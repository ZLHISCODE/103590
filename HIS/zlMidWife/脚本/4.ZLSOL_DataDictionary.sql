数据字典：
create table SOL_STD_FetalPosition--胎方位
(
code varchar2(10),
name varchar2(50),
Description varchar2(100)
);
create table SOL_STD_Delivery--分娩方式
(
code varchar2(10),
name varchar2(50),
Description varchar2(100)
);
create table SOL_STD_PerinealLaceration--会阴裂伤情况
(
code varchar2(10),
name varchar2(50),
Description varchar2(100)
);
create table SOL_STD_Anesthesia--麻醉方式
(
code varchar2(10),
name varchar2(50),
Description varchar2(500)
);
create table SOL_STD_FetalPresentation--胎先露
(
code varchar2(10),
name varchar2(50),
Description varchar2(100)
);
create table SOL_STD_NeonatalAbnormality--新生儿异常情况
(
code varchar2(10),
name varchar2(50),
Description varchar2(100)
);

--胎方位
Insert Into SOL_STD_FetalPosition(code,name,Description) Values('01','左枕前(LOA)','');
Insert Into SOL_STD_FetalPosition(code,name,Description) Values('02','右枕前(ROA)','');
Insert Into SOL_STD_FetalPosition(code,name,Description) Values('03','左枕后(LOP)','');
Insert Into SOL_STD_FetalPosition(code,name,Description) Values('04','右枕后(ROP)','');
Insert Into SOL_STD_FetalPosition(code,name,Description) Values('05','左枕横(LOT)','');
Insert Into SOL_STD_FetalPosition(code,name,Description) Values('06','右枕横(ROT)','');
Insert Into SOL_STD_FetalPosition(code,name,Description) Values('07','左颏前(LMA)','');
Insert Into SOL_STD_FetalPosition(code,name,Description) Values('08','右颏前(RMA)','');
Insert Into SOL_STD_FetalPosition(code,name,Description) Values('09','左颏后(LMP)','');
Insert Into SOL_STD_FetalPosition(code,name,Description) Values('10','右颏后(RMP)','');
Insert Into SOL_STD_FetalPosition(code,name,Description) Values('11','左颏横(LMT)','');
Insert Into SOL_STD_FetalPosition(code,name,Description) Values('12','右颏横(RMT)','');
Insert Into SOL_STD_FetalPosition(code,name,Description) Values('13','左骶前(LSA)','');
Insert Into SOL_STD_FetalPosition(code,name,Description) Values('14','右骶前(RSA)','');
Insert Into SOL_STD_FetalPosition(code,name,Description) Values('15','左骶后(LSP)','');
Insert Into SOL_STD_FetalPosition(code,name,Description) Values('16','右骶后(RSP)','');
Insert Into SOL_STD_FetalPosition(code,name,Description) Values('17','左骶横(LST)','');
Insert Into SOL_STD_FetalPosition(code,name,Description) Values('18','右骶横(RST)','');
Insert Into SOL_STD_FetalPosition(code,name,Description) Values('19','左肩前(LScA)','');
Insert Into SOL_STD_FetalPosition(code,name,Description) Values('20','右肩前(RscA)','');
Insert Into SOL_STD_FetalPosition(code,name,Description) Values('21','左肩后(LScP)','');
Insert Into SOL_STD_FetalPosition(code,name,Description) Values('22','右肩后(RScP)','');
Insert Into SOL_STD_FetalPosition(code,name,Description) Values('99','不祥','');
--分娩方式
Insert Into SOL_STD_Delivery(code,name,Description) Values('1','阴道自然分娩','');
Insert Into SOL_STD_Delivery(code,name,Description) Values('11','会阴切开','');
Insert Into SOL_STD_Delivery(code,name,Description) Values('12','会阴未切','');
Insert Into SOL_STD_Delivery(code,name,Description) Values('2','阴道手术助产','');
Insert Into SOL_STD_Delivery(code,name,Description) Values('21','产钳助产','');
Insert Into SOL_STD_Delivery(code,name,Description) Values('22','臀位助产','');
Insert Into SOL_STD_Delivery(code,name,Description) Values('23','胎头吸引','');
Insert Into SOL_STD_Delivery(code,name,Description) Values('3','剖宫产','');
Insert Into SOL_STD_Delivery(code,name,Description) Values('31','子宫下段横切口剖宫产','');
Insert Into SOL_STD_Delivery(code,name,Description) Values('32','子宫体剖宫产','');
Insert Into SOL_STD_Delivery(code,name,Description) Values('33','腹膜外剖宫产','');
Insert Into SOL_STD_Delivery(code,name,Description) Values('9','其他','');
--会阴裂伤情况
Insert Into SOL_STD_PerinealLaceration(code,name,Description) Values('1','无裂伤','');
Insert Into SOL_STD_PerinealLaceration(code,name,Description) Values('2','Ⅰ°裂伤','');
Insert Into SOL_STD_PerinealLaceration(code,name,Description) Values('3','Ⅱ°裂伤','');
Insert Into SOL_STD_PerinealLaceration(code,name,Description) Values('4','Ⅲ°裂伤','');
Insert Into SOL_STD_PerinealLaceration(code,name,Description) Values('5','会阴切开','');
--麻醉方式
Insert Into SOL_STD_Anesthesia(code,name,Description) Values('1','全身麻醉','用麻醉剂使全身处于麻醉状态');
Insert Into SOL_STD_Anesthesia(code,name,Description) Values('11','吸入麻醉','用吸入麻醉剂的方法使全身处于麻醉状态');
Insert Into SOL_STD_Anesthesia(code,name,Description) Values('12','静脉麻醉','经静脉注入麻醉剂使全身处于麻醉状态');
Insert Into SOL_STD_Anesthesia(code,name,Description) Values('13','基础麻醉','麻醉前先使患者神志消失的方法');
Insert Into SOL_STD_Anesthesia(code,name,Description) Values('2','椎管内麻醉','将麻醉药注入椎管内达到局部麻醉效果的方法');
Insert Into SOL_STD_Anesthesia(code,name,Description) Values('21','蛛网膜下腔阻滞麻醉','将麻醉药注入蛛网膜下腔达到局部麻醉效果的方法');
Insert Into SOL_STD_Anesthesia(code,name,Description) Values('22','硬脊膜外腔阻滞麻醉','将麻醉药注入硬脊膜外腔产生局部麻醉效果的方法');
Insert Into SOL_STD_Anesthesia(code,name,Description) Values('3','局部麻醉','将麻醉药直接注入施行手术的组织内或手术部位周围的麻醉方法');
Insert Into SOL_STD_Anesthesia(code,name,Description) Values('31','神经丛阻滞麻醉','将局部麻醉药注射于神经丛附近，使通过神经丛的神经及其所分布的区域产生局部麻醉的方法');
Insert Into SOL_STD_Anesthesia(code,name,Description) Values('32','神经节阻滞麻醉','将局部麻醉药注射于神经节附近，使通过神经节的神经及其所分布的区域产生局部麻醉的方法');
Insert Into SOL_STD_Anesthesia(code,name,Description) Values('33','神经阻滞麻醉','将局麻药物注射于神经干的周围，使该神经分布的区域产生麻醉作用的方法');
Insert Into SOL_STD_Anesthesia(code,name,Description) Values('34','区域阻滞麻醉','将局麻药注射于手术野外周，使通往手术野以及由手术野传出的神经末梢皆受到阻滞的局部麻醉方法');
Insert Into SOL_STD_Anesthesia(code,name,Description) Values('35','局部浸润麻醉','将局麻药沿手术切口线分层注入组织内，以阻滞组织中的神经末梢的麻醉方法');
Insert Into SOL_STD_Anesthesia(code,name,Description) Values('36','表面麻醉','将麻醉药直接与粘膜或皮肤接触，使支配该部分粘膜或皮肤内的神经末梢被阻滞的麻醉方法');
Insert Into SOL_STD_Anesthesia(code,name,Description) Values('4','复合麻醉','用一种以上药物或采用多种麻醉方法以增强麻醉效果');
Insert Into SOL_STD_Anesthesia(code,name,Description) Values('41','静吸复合全麻','静脉麻醉和吸入麻醉共同作用产生麻醉效果');
Insert Into SOL_STD_Anesthesia(code,name,Description) Values('42','针药复合麻醉','针刺麻醉和药物麻醉共同作用产生麻醉效果');
Insert Into SOL_STD_Anesthesia(code,name,Description) Values('43','神经丛与硬膜外阻滞复合麻醉','神经丛阻滞麻醉和硬脊膜外腔阻滞麻醉共同作用产生麻醉效果');
Insert Into SOL_STD_Anesthesia(code,name,Description) Values('44','全麻复合全身降温','在全身麻醉的同时主动降低患者血压');
Insert Into SOL_STD_Anesthesia(code,name,Description) Values('45','全麻复合控制性降压','在全身麻醉的同时降低患者的体温');
Insert Into SOL_STD_Anesthesia(code,name,Description) Values('9','其他麻醉方法','以上未提及的其他麻醉方法');
--胎先露
Insert Into SOL_STD_FetalPresentation(code,name,Description) Values('1','头先露','');
Insert Into SOL_STD_FetalPresentation(code,name,Description) Values('2','臀先露','');
Insert Into SOL_STD_FetalPresentation(code,name,Description) Values('3','肩先露','');
Insert Into SOL_STD_FetalPresentation(code,name,Description) Values('4','足先露','');
Insert Into SOL_STD_FetalPresentation(code,name,Description) Values('9','不详','');
--新生儿异常情况
Insert Into SOL_STD_NeonatalAbnormality(code,name,Description) Values('1','无','');
Insert Into SOL_STD_NeonatalAbnormality(code,name,Description) Values('2','早期新生儿死亡','');
Insert Into SOL_STD_NeonatalAbnormality(code,name,Description) Values('3','畸形','');
Insert Into SOL_STD_NeonatalAbnormality(code,name,Description) Values('4','早产','');
Insert Into SOL_STD_NeonatalAbnormality(code,name,Description) Values('5','窒息','');
Insert Into SOL_STD_NeonatalAbnormality(code,name,Description) Values('6','低出生体重','');
Insert Into SOL_STD_NeonatalAbnormality(code,name,Description) Values('9','其他','');
