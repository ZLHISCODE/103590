--[连续升级]1
--[管理工具版本号]10.35.40
--本脚本支持从ZLHIS+ v10.35.30 升级到 v10.35.40
--请以系统所有者登录PLSQL并执行下列脚本
--脚本执行后，请手工升级导出报表
Define n_System=100;
-------------------------------------------------------------------------------
--结构修正部份
-------------------------------------------------------------------------------




-------------------------------------------------------------------------------
--数据修正部份
-------------------------------------------------------------------------------
--103000:李南春,2017-03-06,将家庭电话中满足手机号形式的数据填写到手机号中
--适用范围：病人信息增加了手机号的版本
--修正内容:家庭电话满足手机号格式，将家庭电话填写到手机号中
--修正范围：修正所有病人信息记录
--耗时说明: 数据共451W条，待修正数据记录345652条，该数据修正脚本在24分钟内执行完成，测试环境如下:
--1.硬件环境
--     IBM 3650 M4,cpu E5-2620 2.1GHz,2*6核，32G内存
--     V3700存储,SAS硬盘,10K RPM,Raid 10
--2.软件环境
--     Windows 2008,Oracle 10.2.0.4 64bit
--     日志文件500M/个，Log Buffer设置为200M,PGA为9G,SGA为自动管理，最大25G
--3.数据环境
--     病人信息中需要修正的数据记录：345652
Create Or Replace Procedure Zl1_Optional_病人信息修正 As
  Cursor c_Pati Is
    Select 病人ID From 病人信息 Where  Length(家庭电话) = 11 and substr(家庭电话,1,3) in 
            ('139','138','137','136','135','134','159','158','157',
            '150','151','152','147','188','187','182','183','184','178',
            '130','131','132','156','155','186','185','145','176','133','153','189','180','181','177','173','170') And 手机号 is Null;

  t_PatiId      t_Strlist := t_Strlist();
  n_Array_Size Number := 100000; --每批十万，多了可能PGA不够
  I            Number(8) := 0; --每修正100万条记录提交一次,多了可能Undo不够,少了提交过于频繁
  v_内容 varchar2(500);
Begin
  Select Max(内容) Into v_内容 From Zlupgradeconfig Where 项目 = '病人信息手机号修正_20170228';
  Update Zlupgradeconfig Set 内容 = Null Where 项目 = '病人信息手机号修正_20170228';
  If Sql%Notfound  Then
    Insert Into Zlupgradeconfig (项目, 内容) Values ('病人信息手机号修正_20170228', Null);
  End If;
  If Nvl(v_内容, 0) = '成功' Then
    --数据已修正成功
    Return;
  End If;
  Open c_Pati;
  Loop
    Fetch c_Pati Bulk Collect
      Into t_PatiId Limit n_Array_Size;
    Exit When t_PatiId.Count = 0;
  
    Forall I In 1 .. t_PatiId.Count
      Update 病人信息 set 手机号 = 家庭电话 Where 病人ID = t_PatiId(I);
  
    If I = 9 Then
      Update Zlupgradeconfig Set 内容 = To_Number(Nvl(内容, 0)) + I * n_Array_Size Where 项目 = '病人信息手机号修正_20170228';
      Commit;
      I := 0;
    Else
      I := I + 1;
    End If;
  End Loop;
  Update Zlupgradeconfig Set 内容 = '成功' Where 项目 = '病人信息手机号修正_20170228';
  Commit;
  Close c_Pati;
End Zl1_Optional_病人信息修正;
/

--101262:王振涛,2017-02-13,有“病人医嘱报告”的错误数据，同时修正升级数据
--问题描述：在使用第三方LIS系统(例如：智方LIS)中，调用第三方LIS接口程序（zlLISInterface）中的Zl_检验报告单_Insert过程
--会在电子病历记录中产生错误数据，对于门诊病人来讲，是没有主页id的，此时会造成在电子病历记录中，填写的主页为0，正确应该填写挂号单id。
--适用范围：使用第三方LIS系统，此问题影响所有版本，修正脚本对所有版本通用
--修正内容:
--1.修正电子病历记录中， 病人为门诊的病人，进行填写挂号单id。
--2.病人来源是其他的，填写对应的主页id。
--修正范围：
--1.电子病历记录中已经产生的历史错误数据,主页id是空或者0 的数据，进行修正。
--2.直接登记的检验、检查病人没有挂号单号,主页ID也是为0，这类数据不修正。
--耗时说明: 错误数据记录2225269条，该数据修正脚本在45分钟内执行完成，测试环境如下:
--1.硬件环境
--     IBM 3650 M4,cpu E5-2620 2.1GHz,2*6核，32G内存
--     V3700存储,SAS硬盘,10K RPM,Raid 10
--2.软件环境
--     Windows 2008,Oracle 10.2.0.4 64bit
--     日志文件500M/个，Log Buffer设置为200M,PGA为9G,SGA为自动管理，最大25G
--3.数据环境
--     电子病历记录中需要修正的数据记录：2225269
Create Or Replace Procedure Zl1_Optional_门诊病历修正_1 As
  n_主页id 电子病历记录.主页id%Type;
  I        Number(8) := 0;
   v_内容 varchar2(500);
Begin
  Update Zlupgradeconfig Set 内容 = Null Where 项目 = '门诊检验申请病历修正_20160217';
  If Sql%NotFound Then
    Insert Into Zlupgradeconfig (项目, 内容) Values ('门诊检验申请病历修正_20160217', Null);
  End If;
  Commit;
  If Nvl(v_内容, '空') = '成功' Then
    --数据已修正成功
    Return;
  End If;
  For c_Rec In (Select ID, 病人id, 科室id, 创建时间, 病人来源 From 电子病历记录 A Where Nvl(主页id, 0) = 0) Loop
  
    If c_Rec.病人来源 = 1 Then
	  --有挂号单id的病人，进行数据修正
      Select Nvl(Max(d.Id), 0)
      Into n_主页id
      From 病人医嘱报告 B, 病人医嘱记录 C, 病人挂号记录 D
      Where b.病历id = c_Rec.Id And b.医嘱id = c.Id And c.挂号单 = d.No;
    
      --直接登记的检验、检查病人没有挂号单号,主页ID也是为0，这类数据不修正
    Else
      Select Nvl(Max(主页id), 0)
      Into n_主页id
      From 病案主页
      Where 病人id = c_Rec.病人id And c_Rec.创建时间 Between 入院日期 AND  Nvl(出院日期,c_Rec.创建时间);        
    End If;
  
    If n_主页id > 0 Then
      Update 电子病历记录 Set 主页id = n_主页id Where Nvl(主页id, 0) = 0 And ID = c_Rec.Id;
    
      --每一万条提交一次
      I := I + 1;
      If I = 10000 Then
        Update Zlupgradeconfig Set 内容 = To_Number(Nvl(内容, 0)) + I Where 项目 = '门诊检验申请病历修正_20160217';
        Commit;
        I := 0;
      End If;
    End If;
  End Loop;
  Update Zlupgradeconfig Set 内容 = To_Number(Nvl(内容, 0)) + I Where 项目 = '门诊检验申请病历修正_20160217';
  Update Zlupgradeconfig Set 内容 = '成功' Where 项目 = User || '门诊检验申请病历修正_20160217';
  Commit;
End Zl1_Optional_门诊病历修正_1;
/
--101262:王振涛,2017-02-13,缺少“病人医嘱报告”的错误数据，同时修正升级数据
--问题描述：在使用第三方LIS系统(例如：智方LIS)中，调用第三方LIS接口程序（zlLISInterface）
--会造成产生缺少“病人医嘱报告”的错误数据
--适用范围：使用第三方LIS系统，此问题影响所有版本，修正脚本对所有版本通用
--修正内容:
--1.修正电子病历记录中， 病人为门诊的病人，进行填写挂号单id。
--2.病人来源是其他的，填写对应的主页id。
--修正范围：
--1.电子病历记录中已经产生的历史错误数据,主页id是空或者0 的数据，进行修正。
--2.直接登记的检验、检查病人没有挂号单号,主页ID也是为0，这类数据不修正。
--耗时说明: 修正有“病人医嘱报告”的错误数据，数据记录2225269条，该数据修正脚本在2分钟内执行完成，测试环境如下:
--1.硬件环境
--     IBM 3650 M4,cpu E5-2620 2.1GHz,2*6核，32G内存
--     V3700存储,SAS硬盘,10K RPM,Raid 10
--2.软件环境
--     Windows 2008,Oracle 10.2.0.4 64bit
--     日志文件500M/个，Log Buffer设置为200M,PGA为9G,SGA为自动管理，最大25G
--3.数据环境
--     XX医院运行10年的数据
--     电子病历记录中需要修正的数据记录：2225269
--以下修正处理结果：1891
--剩余127行病人来源为2的数据未能修正（缺病案主页数据）
Create Or Replace Procedure Zl1_Optional_门诊病历修正_2 As
  n_主页id 电子病历记录.主页id%Type;
  I        Number(8) := 0;
   v_内容 varchar2(500);
Begin
  Update Zlupgradeconfig Set 内容 = Null Where 项目 = '门诊检验申请病历修正_20160218';
  If Sql%NotFound Then
    Insert Into Zlupgradeconfig (项目, 内容) Values ('门诊检验申请病历修正_20160218', Null);
  End If;
  Commit;
  If Nvl(v_内容, '空') = '成功' Then
    --数据已修正成功
    Return;
  End If;
  For c_Rec In (Select ID, 病人id, 科室id, 创建时间, 病人来源
                From 电子病历记录 A
                Where Nvl(主页id, 0) = 0 And Not Exists (Select 1 From 病人医嘱报告 B Where a.Id = b.病历id)) Loop
  
    If c_Rec.病人来源 = 1 Then
      Select Nvl(Max(ID), 0)
      Into n_主页id
      From 病人挂号记录
      Where 病人id = c_Rec.病人id And 执行部门id = c_Rec.科室id And 登记时间 < c_Rec.创建时间;
    
      If n_主页id = 0 Then
        Select Nvl(Max(ID), 0)
        Into n_主页id
        From 病人挂号记录
        Where 病人id = c_Rec.病人id And 登记时间 < c_Rec.创建时间;
      End If;
    Else
      Select Nvl(Max(主页id), 0)
      Into n_主页id
      From 病案主页
      Where 病人id = c_Rec.病人id And c_Rec.创建时间 Between 入院日期 And Nvl(出院日期, c_Rec.创建时间);
    
      --如果没有病案主页数据，但3天内有门诊挂号数据（不管执行科室），说明：病人来源=2，数据是错误的，应该是1
      If n_主页id = 0 Then
        Select Nvl(Max(ID), 0)
        Into n_主页id
        From 病人挂号记录
        Where 病人id = c_Rec.病人id And 登记时间 Between Trunc(c_Rec.创建时间 - 3) And c_Rec.创建时间;
        
        If n_主页id <> 0 Then
          Update 电子病历记录 Set 主页id = n_主页id, 病人来源 = 1 Where Nvl(主页id, 0) = 0 And ID = c_Rec.Id;
        End If;      
      End If;
    End If;
  
    If n_主页id > 0 Then
      Update 电子病历记录 Set 主页id = n_主页id Where Nvl(主页id, 0) = 0 And ID = c_Rec.Id;
    
      --每一万条提交一次
      I := I + 1;
      If I = 10000 Then
        Update Zlupgradeconfig Set 内容 = To_Number(Nvl(内容, 0)) + I Where 项目 = '门诊检验申请病历修正_20160218';
        Commit;
        I := 0;
      End If;
    End If;
  End Loop;
  Update Zlupgradeconfig Set 内容 = To_Number(Nvl(内容, 0)) + I Where 项目 = '门诊检验申请病历修正_20160218';
  Update Zlupgradeconfig Set 内容 = '成功' Where 项目 = User || '门诊检验申请病历修正_20160218';
  Commit;
End Zl1_Optional_门诊病历修正_2;
/




-------------------------------------------------------------------------------
--权限修正部份
-------------------------------------------------------------------------------





-------------------------------------------------------------------------------
--报表修正部份
-------------------------------------------------------------------------------






-------------------------------------------------------------------------------
--过程修正部份
-------------------------------------------------------------------------------





---------------------------------------------------------------------------------------------------
--更改系统及部件的版本号
-------------------------------------------------------------------------------------------------------
--系统版本号
--部件版本号
Commit;