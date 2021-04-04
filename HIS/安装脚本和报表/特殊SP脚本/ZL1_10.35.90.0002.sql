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
--122063:殷瑞,2018-03-12,输液配置中心可以根据选择的病区来显示对应的输液单
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 部门, 性质, 参数号, 参数名, 参数值, 缺省值, 影响控制说明, 参数值含义, 关联说明, 适用说明, 警告说明)
  Select Zlparameters_Id.Nextval, &n_System, 1345, 0, 1, 0, 0, 0, 0, 44, '显示来源病区', '', '', '在查询输液单时作为查询条件，如果输液单中的病区在来源病区列表中时则满足查询条件',
         '来源病区名称列表，按病区ID1，病区ID2…格式保存', '', '输液配置中心如果有多台机器可以分别对不同病区进行操作，用设置不同的来源病区的方式来提取对应的病区的输液单据', Null
  From Dual;


--122724:刘鹏飞,2018-03-12,输血科直接发血提示医生站
insert into 业务消息类型(编码,名称,说明,保留天数) values ('ZLHIS_BLOOD_004','输血科直接发血提醒','输血科直接发血完成，提醒医生站进行医嘱审核',7);


-------------------------------------------------------------------------------
--权限修正部份
-------------------------------------------------------------------------------





-------------------------------------------------------------------------------
--报表修正部份
-------------------------------------------------------------------------------




-------------------------------------------------------------------------------
--过程修正部份
-------------------------------------------------------------------------------
--122937:胡俊勇,2018-03-16,护理接口出参添加的结点属性
Create Or Replace Procedure Zl_Third_Getadviceinfo
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --功能：获取医嘱基本信息/查询
  --参数：
  --入参数 Xml_In
  --<IN>
  --     <YZID>1156789</YZID>--主医嘱ID
  --</IN>

  --出参 Xml_Out
  --<OUTPUT>
  --    <YZ>
  --       <PATIID></PATIID>     --病人医嘱记录.病人ID
  --       <PAGEID></PAGEID>     --病人医嘱记录.主页ID
  --       <BABY></BABY>   --病人医嘱记录.婴儿
  --       <YZID>1145878</YZID>   --病人医嘱记录.医嘱ID， 主医嘱ID
  --       <RELATEDID></RELATEDID>   --病人医嘱记录.相关ID
  --       <ZXKSID></ZXKSID>   --病人医嘱记录.执行科室id
  --       <YZQX>0</YZQX>      --病人医嘱记录.医嘱期效
  --       <STATE>8</STATE>    --病人医嘱记录.医嘱状态
  --       <JJBZ>0</JJBZ>      --病人医嘱记录.紧急标志
  --       <KZYS>代翔</KZYS>   --病人医嘱记录.开嘱医生
  --       <KZSJ>2015-03-25 16:37:00</KZSJ>   --病人医嘱记录.开嘱时间
  --       <ZLXMID></ZLXMID>   --诊疗项目目录.ID
  --       <ZLLB>E</ZLLB>      --诊疗项目目录.类别
  --       <ZLXMMC></ZLXMMC>   --诊疗项目目录.名称 ，检查，检验(检验行 C)，手术(主手术行 F)，输血(K)，中药配方(服法行 E)，其它(本身)
  --       <ZLXMCZLX></<ZLXMCZLX>   --诊疗项目目录.操作类型
  --       <ZLXMZXFL></ZLXMZXFL>   --诊疗项目目录.执行分类
  --       <BZ>21</BZ> 诊疗项目目录.操作类型||诊疗项目目录.执行分类
  --       <YF>静脉滴注</YF>   --病人医嘱记录.医嘱内容 ，主医嘱行中的  医嘱内容
  --       <PC>BID</PC>   --诊疗频率项目.英文名称
  --       <ZXSJFY>18-20</ZXSJFY>   --病人医嘱记录.执行时间方案
  --       <PLCS>2</PLCS>   --病人医嘱记录.频率次数
  --       <PLJG>1</PLJG>   --病人医嘱记录.频率间隔
  --       <PSJG></PSJG>   --病人医嘱记录.皮试结果
  --       <YSZT></YSZT>   --病人医嘱记录.医生嘱托
  --       <KSZXSJ>2015-03-25 16:35:00</KSZXSJ>  --病人医嘱记录.开始执行时间
  --       <ZXZZSJ></ZXZZSJ>   --病人医嘱记录.执行终止时间
  --       <TZYS></TZYS>   --病人医嘱记录.停嘱医生
  --       <TZSJ></TZSJ>   --病人医嘱记录.停嘱时间
  --       <DW>次</DW>   --诊疗项目目录.计算单位
  --       <DL></DL>   --病人医嘱记录.单次用量
  --       <ZL></ZL>   --病人医嘱记录.总给予量

  --       <ITEMLIST> 仅输血项目和西/成药医嘱项目明细相关信息；输血的血袋信息，药品行明细信息
  --        <ITEM>
  --         <YSZT></YSZT>   --病人医嘱记录.医生嘱托
  --         <YZID>1145878</YZID>   --病人医嘱记录.医嘱ID
  --         <RELATEDID></RELATEDID>   --病人医嘱记录.相关ID
  --         <ZLXMID></ZLXMID>   --诊疗项目目录.ID
  --         <SFXMID></SFXMID>   --收费项目目录.id
  --         <SFXMMC></SFXMMC>   --收费项目目录.名称
  --         <SFXMGG></SFXMGG>   --收费项目目录.规格
  --         <BM></BM>           --收费项目别名.名称（商品名）
  --         <ZL></ZL>           --病人医嘱记录.总给予量
  --         <DL>10</DL>         --病人医嘱记录.单次用量
  --         <DW>ml</DW>         --收费项目目录.计算单位
  --         <ZLDW>ml</ZLDW>   --诊疗项目目录.计算单位
  --         <ZXXZ></ZXXZ>   --病人医嘱记录.执行性质
  --         <ZXKS></ZXKS>   --诊疗项目目录.执行科室
  --         <XDBH></XDBH>   --血液收发记录.血袋编号
  --         <SXXH></SXXH>   --血液收发记录.序号
  --        </ITEM>
  --        <ITEM/>...
  --       </ITEMLIST>
  --      </YZ>
  --</OUTPUT>

  n_医嘱id  病人医嘱记录.Id%Type;
  x_医嘱    Xmltype;
  x_Item    Xmltype;
  v_Xtmp    Clob; --临时XML
  n_Cnt     Number;
  x_Templet Xmltype;

  v_英文名     诊疗频率项目.英文名称%Type;
  v_试管名称   采血管类型.名称%Type;
  v_添加剂     采血管类型.添加剂%Type;
  v_试管规格   采血管类型.规格%Type;
  n_试管颜色   采血管类型.颜色%Type;
  v_收费商品名 收费项目别名.名称%Type;
  n_启用血库   Number; 
  v_Sql血库    Varchar2(4000);
  n_血库申请id Number(18);
  v_Tmp输血    Varchar2(4000);

  Type Bloodlist_Type Is Ref Cursor;
  Cbloodlist Bloodlist_Type;

  Type t_Code Is Record(
    ID       收费项目目录.Id%Type,
    名称     收费项目目录.名称%Type,
    规格     收费项目目录.规格%Type,
    单位     收费项目目录.计算单位%Type,
    血袋编号 Varchar2(50),
    序号     Number(5));
  r_b t_Code;

Begin

  Select Extractvalue(Value(A), 'IN/YZID') Into n_医嘱id From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;
  n_Cnt := 0;
  For R In (Select a.病人id, a.主页id, a.婴儿, a.Id As 医嘱id, a.相关id, a.执行科室id, a.医嘱期效, a.医嘱状态, a.紧急标志, a.开嘱医生, a.开嘱时间, a.诊疗项目id,
                   a.诊疗类别, a.医嘱内容, a.执行时间方案, a.执行频次, a.频率次数, a.频率间隔, a.皮试结果, a.医生嘱托, a.开始执行时间, a.执行终止时间, a.停嘱医生, a.停嘱时间,
                   b.名称 As 项目名称, b.操作类型, b.执行分类, b.计算单位 As 诊疗单位, a.单次用量, a.总给予量, a.标本部位, a.检查方法, a.收费细目id, c.名称 As 收费名称,
                   c.规格, Null As 收费商品名, c.计算单位 As 收费单位, a.执行性质, b.执行科室, b.试管编码
            From 病人医嘱记录 A, 诊疗项目目录 B, 收费项目目录 C
            Where a.诊疗项目id = b.Id And a.收费细目id = c.Id(+) And (a.Id = n_医嘱id Or a.相关id = n_医嘱id)
            Order By a.序号) Loop
    n_Cnt := n_Cnt + 1;
    If n_Cnt = 1 Then
      Select Max(a.英文名称) Into v_英文名 From 诊疗频率项目 A Where a.名称 = r.执行频次;
    End If;
    v_试管名称 := Null;
    v_添加剂   := Null;
    v_试管规格 := Null;
    n_试管颜色 := Null;
    If r.试管编码 Is Not Null Then
      Select Max(a.名称), Max(a.添加剂), Max(a.规格), Max(a.颜色)
      Into v_试管名称, v_添加剂, v_试管规格, n_试管颜色
      From 采血管类型 A
      Where a.编码 = r.试管编码;
    End If;
    --主医行
    If r.相关id Is Null Then
      v_Xtmp := '<YZ>';
      v_Xtmp := v_Xtmp || '<PATIID>' || r.病人id || '</PATIID>'; --病人医嘱记录.病人ID
      v_Xtmp := v_Xtmp || '<PAGEID>' || r.主页id || '</PAGEID>'; --病人医嘱记录.主页ID
      v_Xtmp := v_Xtmp || '<BABY>' || r.婴儿 || '</BABY>'; --病人医嘱记录.婴儿
      v_Xtmp := v_Xtmp || '<YZID>' || r.医嘱id || '</YZID>'; --病人医嘱记录.医嘱ID
      v_Xtmp := v_Xtmp || '<RELATEDID>' || r.相关id || '</RELATEDID>'; --病人医嘱记录.相关ID
      v_Xtmp := v_Xtmp || '<ZXKSID>' || r.执行科室id || '</ZXKSID>'; --病人医嘱记录.执行科室id
      v_Xtmp := v_Xtmp || '<YZQX>' || r.医嘱期效 || '</YZQX>'; --病人医嘱记录.医嘱期效
      v_Xtmp := v_Xtmp || '<STATE>' || r.医嘱状态 || '</STATE>'; --病人医嘱记录.医嘱状态
      v_Xtmp := v_Xtmp || '<JJBZ>' || r.紧急标志 || '</JJBZ>'; --病人医嘱记录.紧急标志
      v_Xtmp := v_Xtmp || '<KZYS>' || r.开嘱医生 || '</KZYS>'; --病人医嘱记录.开嘱医生
      v_Xtmp := v_Xtmp || '<KZSJ>' || To_Char(r.开嘱时间, 'yyyy-mm-dd hh24:mi:ss') || '</KZSJ>'; --病人医嘱记录.开嘱时间
      v_Xtmp := v_Xtmp || '<BZ>' || r.操作类型 || r.执行分类 || '</BZ>'; -- 诊疗项目目录.操作类型||诊疗项目目录.执行分类
      v_Xtmp := v_Xtmp || '<ZLXMID>' || r.诊疗项目id || '</ZLXMID>'; --诊疗项目目录.ID
      v_Xtmp := v_Xtmp || '<ZLLB>' || r.诊疗类别 || '</ZLLB>'; --诊疗项目目录.类别
      v_Xtmp := v_Xtmp || '<YZNR>' || r.医嘱内容 || '</YZNR>'; --医嘱内容
      v_Xtmp := v_Xtmp || '<YF>' || r.项目名称 || '</YF>'; --病人医嘱记录.医嘱内容
      v_Xtmp := v_Xtmp || '<PC>' || v_英文名 || '</PC>'; --诊疗频率项目.英文名称
      v_Xtmp := v_Xtmp || '<ZXSJFY>' || r.执行时间方案 || '</ZXSJFY>'; --病人医嘱记录.执行时间方案
      v_Xtmp := v_Xtmp || '<PLCS>' || r.频率次数 || '</PLCS>'; --病人医嘱记录.频率次数
      v_Xtmp := v_Xtmp || '<PLJG>' || r.频率间隔 || '</PLJG>'; --病人医嘱记录.频率间隔
      v_Xtmp := v_Xtmp || '<PSJG>' || r.皮试结果 || '</PSJG>'; --病人医嘱记录.皮试结果
      v_Xtmp := v_Xtmp || '<YSZT>' || r.医生嘱托 || '</YSZT>'; --病人医嘱记录.医生嘱托
      v_Xtmp := v_Xtmp || '<KSZXSJ>' || To_Char(r.开始执行时间, 'yyyy-mm-dd hh24:mi:ss') || '</KSZXSJ>'; --病人医嘱记录.开始执行时间
      v_Xtmp := v_Xtmp || '<ZXZZSJ>' || To_Char(r.执行终止时间, 'yyyy-mm-dd hh24:mi:ss') || '</ZXZZSJ>'; --病人医嘱记录.执行终止时间
      v_Xtmp := v_Xtmp || '<TZYS>' || r.停嘱医生 || '</TZYS>'; --病人医嘱记录.停嘱医生
      v_Xtmp := v_Xtmp || '<TZSJ>' || To_Char(r.停嘱时间, 'yyyy-mm-dd hh24:mi:ss') || '</TZSJ>'; --病人医嘱记录.停嘱时间
      v_Xtmp := v_Xtmp || '<ZLXMMC>' || r.项目名称 || '</ZLXMMC>'; --诊疗项目目录.名称
      v_Xtmp := v_Xtmp || '<ZLXMCZLX>' || r.操作类型 || '</ZLXMCZLX>'; --诊疗项目目录.操作类型
      v_Xtmp := v_Xtmp || '<ZLXMZXFL>' || r.执行分类 || '</ZLXMZXFL>'; --诊疗项目目录.执行分类
      --       (仅采血管返回)
      v_Xtmp := v_Xtmp || '<CXGMC>' || v_试管名称 || '</CXGMC>'; --采血管名称
      v_Xtmp := v_Xtmp || '<CXGTJJ>' || v_添加剂 || '</CXGTJJ>'; --采血管添加剂
      v_Xtmp := v_Xtmp || '<CXGGG>' || v_试管规格 || '</CXGGG>'; --采血管规格
      v_Xtmp := v_Xtmp || '<CXGYS>' || n_试管颜色 || '</CXGYS>'; --采血管颜色
      v_Xtmp := v_Xtmp || '<DW>' || r.诊疗单位 || '</DW>'; --诊疗项目目录.计算单位
      v_Xtmp := v_Xtmp || '<DL>' || r.单次用量 || '</DL>'; --病人医嘱记录.单次用量
      v_Xtmp := v_Xtmp || '<ZL>' || r.总给予量 || '</ZL>'; --病人医嘱记录.总给予量
      v_Xtmp := v_Xtmp || '</YZ>';
      x_医嘱 := Xmltype(v_Xtmp);
    End If;
  
    --输血
    If r.诊疗类别 = 'K' Then
      --判断是否安装血库
      Select Zl_Checkobject(1, '血液收发记录') Into n_启用血库 From Dual;
      If n_启用血库 > 0 Then
        n_血库申请id := r.医嘱id;
        --医嘱部分
        v_Xtmp    := '<YSZT>' || r.医生嘱托 || '</YSZT>'; --病人医嘱记录.医生嘱托
        v_Xtmp    := v_Xtmp || '<YZID>' || r.医嘱id || '</YZID>'; --病人医嘱记录.医嘱ID
        v_Xtmp    := v_Xtmp || '<RELATEDID>' || r.相关id || '</RELATEDID>'; --病人医嘱记录.相关ID
        v_Xtmp    := v_Xtmp || '<ZLXMID>' || r.诊疗项目id || '</ZLXMID>'; --诊疗项目目录.ID
        v_Xtmp    := v_Xtmp || '<ZL>' || r.总给予量 || '</ZL>'; --病人医嘱记录.总给予量
        v_Xtmp    := v_Xtmp || '<DL>' || r.单次用量 || '</DL>'; --病人医嘱记录.单次用量
        v_Xtmp    := v_Xtmp || '<ZLDW>' || r.诊疗单位 || '</ZLDW>'; --诊疗项目目录.计算单位
        v_Xtmp    := v_Xtmp || '<ZXXZ>' || r.执行性质 || '</ZXXZ>'; --病人医嘱记录.执行性质
        v_Xtmp    := v_Xtmp || '<ZXKS>' || r.执行科室 || '</ZXKS>'; --诊疗项目目录.执行科室
        v_Tmp输血 := v_Xtmp;
        If r.检查方法 = '1' Then
          v_Sql血库 := 'Select d.Id,d.名称,d.规格,d.计算单位 as 单位, a.血袋编号,a.序号
                       From 血液收发记录 a,血液发送记录 b,血液配血记录 c,收费项目目录 d
                       Where a.Id = b.收发id And b.配发id = c.Id and a.血液id =d.id  And c.申请id =:1';
        End If;
      End If;
    Elsif r.相关id Is Not Null And r.诊疗类别 = 'E' And r.操作类型 = '8' And Nvl(r.执行分类, 0) = 0 And n_启用血库 = 1 And
          v_Sql血库 Is Null Then
      v_Sql血库 := 'Select b.Id,b.名称,  b.规格,b.计算单位 as 单位, a.血袋编号,a.序号
                  From 血液收发记录 a,收费项目目录 b
                  Where a.血液id =b.id and a.配发id = (Select Id From 血液配血记录 Where 申请id=:1)';
    Else
      v_Sql血库 := Null;
    End If;
  
    If v_Sql血库 Is Not Null And n_血库申请id Is Not Null Then
      --输血医嘱，只有发医嘱后才可能有血袋信息
      x_Item := Xmltype('<ITEMLIST></ITEMLIST>');
      Open Cbloodlist For v_Sql血库
        Using n_血库申请id;
      Loop
        Fetch Cbloodlist
          Into r_b.Id, r_b.名称, r_b.规格, r_b.单位, r_b.血袋编号, r_b.序号;
        Exit When Cbloodlist%NotFound;
        v_收费商品名 := Null;
        If r_b.Id Is Not Null Then
          For Z In (Select a.名称, a.性质
                    From 收费项目别名 A
                    Where a.收费细目id = r_b.Id
                    Group By a.名称, a.性质
                    Order By a.性质) Loop
            v_收费商品名 := z.名称;
            If z.性质 = 3 Then
              v_收费商品名 := z.名称;
              Exit;
            End If;
          End Loop;
        End If;
      
        v_Xtmp := '<ITEM jsonArray="True" >';
      
        v_Xtmp := v_Xtmp || v_Tmp输血;
      
        --血库部分
        v_Xtmp := v_Xtmp || '<SFXMID>' || r_b.Id || '</SFXMID>'; --收费项目目录.id
        v_Xtmp := v_Xtmp || '<SFXMMC>' || r_b.名称 || '</SFXMMC>'; --收费项目目录.名称
        v_Xtmp := v_Xtmp || '<SFXMGG>' || r_b.规格 || '</SFXMGG>'; --收费项目目录.规格
        v_Xtmp := v_Xtmp || '<BM>' || v_收费商品名 || '</BM>'; --收费项目别名.名称（商品名）
        v_Xtmp := v_Xtmp || '<DW>' || r_b.单位 || '</DW>'; --收费项目目录.计算单位
        v_Xtmp := v_Xtmp || '<XDBH>' || r_b.血袋编号 || '</XDBH>'; --血液收发记录.血袋编号
        v_Xtmp := v_Xtmp || '<SXXH>' || r_b.序号 || '</SXXH>'; --血液收发记录.序号
      
        v_Xtmp := v_Xtmp || '</ITEM>';
        Select Appendchildxml(x_Item, '/ITEMLIST', Xmltype(v_Xtmp)) Into x_Item From Dual;
      End Loop;
      Close Cbloodlist;
    End If;
  
    --西药成药医嘱
    If r.诊疗类别 = '5' Or r.诊疗类别 = '6' Then
      --西/成 药
      If x_Item Is Null Then
        --只初始化一次
        x_Item := Xmltype('<ITEMLIST></ITEMLIST>');
      End If;
      v_收费商品名 := Null;
      If r.收费细目id Is Not Null Then
        For Z In (Select a.名称, a.性质
                  From 收费项目别名 A
                  Where a.收费细目id = r.收费细目id
                  Group By a.名称, a.性质
                  Order By a.性质) Loop
          v_收费商品名 := z.名称;
          If z.性质 = 3 Then
            v_收费商品名 := z.名称;
            Exit;
          End If;
        End Loop;
      End If;
    
      v_Xtmp := '<ITEM jsonArray="True" >';
      v_Xtmp := v_Xtmp || '<YSZT>' || r.医生嘱托 || '</YSZT>'; --病人医嘱记录.医生嘱托
      v_Xtmp := v_Xtmp || '<YZID>' || r.医嘱id || '</YZID>'; --病人医嘱记录.医嘱ID
      v_Xtmp := v_Xtmp || '<RELATEDID>' || r.相关id || '</RELATEDID>'; --病人医嘱记录.相关ID
      v_Xtmp := v_Xtmp || '<ZLXMID>' || r.诊疗项目id || '</ZLXMID>'; --诊疗项目目录.ID
      v_Xtmp := v_Xtmp || '<SFXMID>' || r.收费细目id || '</SFXMID>'; --收费项目目录.id
      v_Xtmp := v_Xtmp || '<SFXMMC>' || r.收费名称 || '</SFXMMC>'; --收费项目目录.名称
      v_Xtmp := v_Xtmp || '<SFXMGG>' || r.规格 || '</SFXMGG>'; --收费项目目录.规格
      v_Xtmp := v_Xtmp || '<BM>' || v_收费商品名 || '</BM>'; --收费项目别名.名称（商品名）
      v_Xtmp := v_Xtmp || '<ZL>' || r.总给予量 || '</ZL>'; --病人医嘱记录.总给予量
      v_Xtmp := v_Xtmp || '<DL>' || r.单次用量 || '</DL>'; --病人医嘱记录.单次用量
      v_Xtmp := v_Xtmp || '<DW>' || r.收费单位 || '</DW>'; --收费项目目录.计算单位
      v_Xtmp := v_Xtmp || '<ZLDW>' || r.诊疗单位 || '</ZLDW>'; --诊疗项目目录.计算单位
      v_Xtmp := v_Xtmp || '<ZXXZ>' || r.执行性质 || '</ZXXZ>'; --病人医嘱记录.执行性质
      v_Xtmp := v_Xtmp || '<ZXKS>' || r.执行科室 || '</ZXKS>'; --诊疗项目目录.执行科室
      v_Xtmp := v_Xtmp || '<XDBH></XDBH>'; --血液收发记录.血袋编号
      v_Xtmp := v_Xtmp || '<SXXH></SXXH>'; --血液收发记录.序号
      v_Xtmp := v_Xtmp || '</ITEM>';
      Select Appendchildxml(x_Item, '/ITEMLIST', Xmltype(v_Xtmp)) Into x_Item From Dual;
    End If;
  End Loop;
  If x_Item Is Not Null Then
    Select Appendchildxml(x_医嘱, '/YZ', x_Item) Into x_医嘱 From Dual;
  End If;
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');
  Select Appendchildxml(x_Templet, '/OUTPUT', x_医嘱) Into x_Templet From Dual;
  Xml_Out := x_Templet;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getadviceinfo;
/

--122937:胡俊勇,2018-03-15,护理接口出参添加的结点属性
CREATE OR REPLACE Procedure Zl_Third_Getpathway
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --功能：病人临床路径表单,阶段信息/查询
  --入参：Xml_In
  --<IN>
  --     <PATIID>29</PATIID>     --病人ID
  --     <PAGEID>1</PAGEID>     --主页ID
  --</IN>

  --出参：Xml_Out
  --<OUTPUT>
  --  <LJID>43</LJID>     --病人临床路径.路径ID
  --  <YSLJID></YSLJID>   --病人路径评估.原路径id
  --  <BBH></BBH>         --病人临床路径.版本号
  --  <YSBBH></YSBBH>     --病人路径评估.原路径版本
  --  <LJMC>急性单纯性阑尾炎临床路径</LJMC>   --临床路径目录.名称
  --  <ZXZT>执行中</ZXZT>   --病人临床路径.状态
  --  <BZZYR>6</BZZYR>   --临床路径版本.标准住院日
  --  <DQTS>5</DQTS>   --病人临床路径.当前天数

  --  <PHASELIST>
  --   <PHASES>
  --    <PHASE>
  --      <JDID>1</JDID>   --病人路径执行.阶段ID
  --      <JD>住院第1天 (住院日,手术日)</JD>   --临床路径阶段.名称
  --      <DQJD>0</DQJD>   --病人临床路径.当前阶段ID
  --      <DAYS>
  --        <DAY>
  --          <TS>1</TS>                     --病人路径执行.天数
  --          <RQ>2011-09-16 00:00:00</RQ>   --病人路径执行.日期
  --          <PGJG>正常</PGJG>              --病人路径评估.评估结果
  --          <PGSM>嘀咕</PGSM>              --病人路径评估.评估说明
  --          <PGR>代翔</PGR>                --病人路径评估.评估人
  --          <PGSJ>2011-09-16 10:51:40</PGSJ>   --病人路径评估.评估时间
  --          <BYYY></BYYY>                      --病人路径评估.变异原因（变异常见原因.名称）
  --          <ITEMLIST>
  --             <ITEM>
  --                <FL>主要诊疗工作</FL>   --临床路径分类.名称
  --                <TBID />   --病人路径执行.图标ID
  --                <ZXID>3366</ZXID>   --病人路径执行.ID
  --                <XMID>1</XMID>   --病人路径执行.项目ID    
  --                <XMXH>1</XMXH>   --临床路径项目.项目序号（XMID为空时，取病人路径执行 .项目序号）   
  --                <XMNR>询问病史体格检查</XMNR>   --临床路径项目.项目内容（XMID为空时，取病人路径执行 .项目内容）
  --                <ZXFS>1</ZXFS>   --临床路径项目.执行方式
  --                <ZXJG>已经执行</ZXJG>   --病人路径执行.执行结果
  --                <TJYY />   --病人路径执行.添加原因
  --                <ZXBYYY />  变异原因
  --              </ITEM>
  --           </ITEMLIST>
  --        </DAY>
  --        <DAY/>...
  --     </DAYS>
  --    </PHASE>
  --    <PHASE/>...
  --  </PHASES>
  -- </PHASELIST>
  --</OUTPUT>

  n_病人id     病人医嘱记录.病人id%Type;
  n_主页id     病人医嘱记录.主页id%Type;
  n_路径id     临床路径阶段.路径id%Type;
  n_版本号     临床路径阶段.版本号%Type;
  n_当前阶段id 病人临床路径.当前阶段id%Type;
  n_路径记录id 病人临床路径.Id%Type;
  v_Xtmp       Clob; --临时XML
  x_Templet    Xmltype;
  x_Phase      Xmltype;
  x_Day        Xmltype;
  x_Item       Xmltype;
  v_评估信息   Varchar2(4000);
  v_Err_Msg    Varchar2(255);
  Err_Item Exception;

  Cursor c_Main Is
    Select a.Id, a.路径id, c.原路径id, a.版本号, c.原路径版本, e.名称 As 路径名称,
           Decode(a.状态, 0, '不符合导入条件', 1, '执行中', 2, '正常结束', 3, '变异结束', Null) As 状态, f.标准住院日, a.当前天数, a.当前阶段id
    From 病人临床路径 A, 病人路径评估 C, 临床路径目录 E, 临床路径版本 F
    Where a.病人id = n_病人id And a.主页id = n_主页id And a.路径id = e.Id And a.当前阶段id = c.阶段id(+) And a.Id = c.路径记录id(+) And
          a.当前天数 = c.天数(+) And a.路径id = f.路径id And a.版本号 = f.版本号;

  Cursor c_定义阶段 Is
    Select a.Id As 阶段id, a.名称 As 阶段名称
    From 临床路径阶段 A
    Where a.路径id = n_路径id And a.版本号 = n_版本号 And a.父id Is Null
    Order By a.序号;

  Type t_定义阶段 Is Table Of c_定义阶段%RowType;
  r_定义阶段 t_定义阶段;

  --已生成的阶段
  Cursor c_Phase Is
    Select a.阶段id, a.天数, To_Char(a.日期, 'yyyy-mm-dd') 日期, To_Char(a.日期, 'day') 星期, b.名称 As 阶段名, b.序号, b.说明, b.父id,
           Decode(g.路径id, b.路径id, 1, 0) As 排序
    From (Select a.阶段id, a.天数, a.日期, a.路径记录id
           From 病人路径执行 A
           Where a.路径记录id = n_路径记录id
           Group By a.阶段id, a.天数, a.日期, a.路径记录id) A, 临床路径阶段 B, 临床路径阶段 C, 病人临床路径 G
    Where a.阶段id = b.Id And b.父id = c.Id(+) And g.Id = a.路径记录id
    Order By 日期, 排序, Nvl(c.序号, b.序号);

  Type t_Phase Is Table Of c_Phase%RowType;
  r_Phase t_Phase;

  --明细项目
  Cursor c_Item Is
    Select a.Id, Nvl(b.图标id, a.图标id) As 图标id, a.分类, To_Char(a.日期, 'yyyy-mm-dd') As 日期, a.天数, a.阶段id,
           Nvl(a.项目序号, b.项目序号) As 项目序号, Nvl(b.项目内容, a.项目内容) 项目内容, a.项目id, Decode(a.执行人, Null, 0, 1) 执行状态,
           Nvl(b.执行方式, 1) 执行方式, a.添加原因, Nvl(a.生成时间性质, 0) As 生成时间性质, c.名称 As 变异原因, Nvl(b.项目结果, a.项目结果) As 项目结果, a.执行结果,
           d.路径id, d.分支id, Nvl(Nvl(a.生成者, b.生成者), 1) As 生成者, d.名称 As 阶段名
    From 病人路径执行 A, 临床路径项目 B, 变异常见原因 C, 临床路径阶段 D
    Where a.路径记录id = n_路径记录id And a.项目id = b.Id(+) And a.变异原因 = c.编码(+) And a.阶段id + 0 = d.Id
    Order By a.日期, 分类, 项目序号;

  Type t_Item Is Table Of c_Item%RowType;
  r_Item t_Item;

  --阶段评估
  Cursor c_Eval Is
    Select a.阶段id, a.天数, Decode(a.评估结果, 1, '正常', -1, '变异', Null) As 评估结果, a.评估说明, a.评估人, a.评估时间, c.名称 As 变异原因, a.变异审核人,
           Nvl(a.时间进度, 0) 时间进度, a.跳转审核人, a.原路径id
    From 病人路径评估 A, 病人路径变异 B, 变异常见原因 C
    Where a.路径记录id = b.路径记录id(+) And a.阶段id = b.阶段id(+) And a.日期 = b.日期(+) And a.路径记录id = n_路径记录id And b.变异原因 = c.编码(+);
  Type t_Eval Is Table Of c_Eval%RowType;
  r_Eval t_Eval;
Begin
  Select Extractvalue(Value(A), 'IN/PATIID') As 病人id, Extractvalue(Value(A), 'IN/PAGEID') As 主页id
  Into n_病人id, n_主页id
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  For R In c_Main Loop
    v_Xtmp := '<LJID>' || r.路径id || '</LJID>'; --病人临床路径.路径ID
    v_Xtmp := v_Xtmp || '<YSLJID>' || r.原路径id || '</YSLJID>'; --病人路径评估.原路径id
    v_Xtmp := v_Xtmp || '<BBH>' || r.版本号 || '</BBH>'; --病人临床路径.版本号
    v_Xtmp := v_Xtmp || '<YSBBH>' || r.原路径版本 || '</YSBBH>'; --病人路径评估.原路径版本
    v_Xtmp := v_Xtmp || '<LJMC>' || r.路径名称 || '</LJMC>'; --临床路径目录.名称
    v_Xtmp := v_Xtmp || '<ZXZT>' || r.状态 || '</ZXZT>'; --病人临床路径.状态
    v_Xtmp := v_Xtmp || '<BZZYR>' || r.标准住院日 || '</BZZYR>'; --临床路径版本.标准住院日
    v_Xtmp := v_Xtmp || '<DQTS>' || r.当前天数 || '</DQTS>'; --病人临床路径.当前天数;
  
    n_路径id     := r.路径id;
    n_版本号     := r.版本号;
    n_当前阶段id := r.当前阶段id;
    n_路径记录id := r.Id;
  End Loop;
  x_Templet := Xmltype('<OUTPUT>' || v_Xtmp || '<PHASELIST></PHASELIST></OUTPUT>');

  If n_路径记录id Is Null Then
    v_Err_Msg := '未找到路径信息！';
    Raise Err_Item;
  End If;

  Open c_定义阶段;
  Fetch c_定义阶段 Bulk Collect
    Into r_定义阶段;
  Close c_定义阶段;

  Open c_Phase;
  Fetch c_Phase Bulk Collect
    Into r_Phase;
  Close c_Phase;

  Open c_Item;
  Fetch c_Item Bulk Collect
    Into r_Item;
  Close c_Item;

  Open c_Eval;
  Fetch c_Eval Bulk Collect
    Into r_Eval;
  Close c_Eval;

  For I In 1 .. r_定义阶段.Count Loop
    v_Xtmp  := '<PHASE jsonArray="True" ><JDID>' || r_定义阶段(I).阶段id || '</JDID><JD>' || r_定义阶段(I).阶段名称 || '</JD><DQJD>' || n_当前阶段id ||
               '</DQJD><DAYS></DAYS></PHASE>';
    x_Phase := Xmltype(v_Xtmp);
  
    For J In 1 .. r_Phase.Count Loop
      If r_定义阶段(I).阶段id = r_Phase(J).阶段id Then
        --day 
        v_Xtmp     := '<DAY jsonArray="True" >';
        v_Xtmp     := v_Xtmp || '<TS>' || r_Phase(J).天数 || '</TS>'; --病人路径执行.天数
        v_Xtmp     := v_Xtmp || '<RQ>' || r_Phase(J).日期 || ' 00:00:00</RQ>'; --病人路径执行.日期      
        v_评估信息 := '<PGJG></PGJG><PGSM></PGSM><PGR></PGR><PGSJ></PGSJ><BYYY></BYYY>';
        For K In 1 .. r_Eval.Count Loop
          If r_Phase(J).阶段id = r_Eval(K).阶段id And r_Phase(J).天数 = r_Eval(K).天数 Then
            v_评估信息 := '<PGJG>' || r_Eval(K).评估结果 || '</PGJG>'; --病人路径评估.评估结果
            v_评估信息 := v_评估信息 || '<PGSM>' || r_Eval(K).评估说明 || '</PGSM>'; --病人路径评估.评估说明
            v_评估信息 := v_评估信息 || '<PGR>' || r_Eval(K).评估人 || '</PGR>'; --病人路径评估.评估人
            v_评估信息 := v_评估信息 || '<PGSJ>' || To_Char(r_Eval(K).评估时间, 'yyyy-mm-dd hh24:mi:ss') || '</PGSJ>'; --病人路径评估.评估时间
            v_评估信息 := v_评估信息 || '<BYYY>' || r_Eval(K).变异原因 || '</BYYY>'; --病人路径评估.变异原因（变异常见原因.名称）         
          End If;
        End Loop;
        v_Xtmp := v_Xtmp || v_评估信息;
        v_Xtmp := v_Xtmp || '</DAY>';
        x_Day  := Xmltype(v_Xtmp);
      
        --item
        x_Item := Xmltype('<ITEMLIST></ITEMLIST>');
        For K In 1 .. r_Item.Count Loop
          If r_Phase(J).阶段id = r_Item(K).阶段id And r_Phase(J).天数 = r_Item(K).天数 Then
            v_Xtmp := '<ITEM jsonArray="True" >';
            v_Xtmp := v_Xtmp || '<FL>' || r_Item(K).分类 || '</FL>'; --临床路径分类.名称
            v_Xtmp := v_Xtmp || '<TBID>' || r_Item(K).图标id || '</TBID>'; --病人路径执行.图标ID
            v_Xtmp := v_Xtmp || '<ZXID>' || r_Item(K).Id || '</ZXID>'; --病人路径执行.ID
            v_Xtmp := v_Xtmp || '<XMID>' || r_Item(K).项目id || '</XMID>'; --病人路径执行.项目ID    
            v_Xtmp := v_Xtmp || '<XMXH>' || r_Item(K).项目序号 || '</XMXH>'; --临床路径项目.项目序号（XMID为空时，取病人路径执行 .项目序号）   
            v_Xtmp := v_Xtmp || '<XMNR>' || r_Item(K).项目内容 || '</XMNR>'; --临床路径项目.项目内容（XMID为空时，取病人路径执行 .项目内容）
            v_Xtmp := v_Xtmp || '<ZXFS>' || r_Item(K).执行方式 || '</ZXFS>'; --临床路径项目.执行方式
            v_Xtmp := v_Xtmp || '<ZXJG>' || r_Item(K).执行结果 || '</ZXJG>'; --病人路径执行.执行结果
            v_Xtmp := v_Xtmp || '<TJYY>' || r_Item(K).添加原因 || '</TJYY>'; --病人路径执行.添加原因
            v_Xtmp := v_Xtmp || '<ZXBYYY>' || r_Item(K).变异原因 || '</ZXBYYY>'; --变异原因
            v_Xtmp := v_Xtmp || '</ITEM>';
            Select Appendchildxml(x_Item, '/ITEMLIST', Xmltype(v_Xtmp)) Into x_Item From Dual;
          End If;
        End Loop;
        Select Appendchildxml(x_Day, '/DAY', x_Item) Into x_Day From Dual;
        Select Appendchildxml(x_Phase, '/PHASE/DAYS', x_Day) Into x_Phase From Dual;
      End If;
    End Loop;
    Select Appendchildxml(x_Templet, '/OUTPUT/PHASELIST', x_Phase) Into x_Templet From Dual;
  End Loop;
  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getpathway;
/

--122937:胡俊勇,2018-03-15,护理接口出参添加的结点属性
CREATE OR REPLACE Procedure Zl_Third_Getpathwaydetail
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --功能：临床路径中某个项目具体的执行明细及医嘱信息/查询
  --入参：Xml_In
  --<IN>
  --     <ZXID>29</ZXID>     --病人路径执行.ID
  --</IN>
  --出参：Xml_Out
  --<OUTPUT>
  -- <LJEXEC>
  --  <ZXQK>
  --   <ZXZ />   --病人路径执行.执行者
  --   <ZXR>代翔</ZXR>   --病人路径执行.执行人
  --   <ZXSJ>2011-10-24 17:28:53</ZXSJ>   --病人路径执行.执行时间
  --   <ZXJG>已经执行</ZXJG>   --病人路径执行.执行结果
  --   <ZXSM />   --病人路径执行..执行说明
  --  </ZXQK>
  --  <YZLIST>
  --   <YZXX>
  --    <YZQX>0</YZQX>   --病人医嘱记录.医嘱期效
  --    <YZNR>注射用克林霉素 0.6g/支 苏州第壹制药有限公司</YZNR>   --病人医嘱记录.医嘱内容
  --    <DL>每次1.2g</DL>   -病人医嘱记录.单次用量
  --    <ZL />   --病人医嘱记录.总给予量
  --    <GYTJ>静脉滴注（门诊）</GYTJ>   --诊疗项目目录.名称
  --    <ZXPL>每天二次</ZXPL>   --病人医嘱记录.执行频次
  --    <ZXSJ>10-16</ZXSJ>   -病人医嘱记录.执行时间方案
  --    <YSZT />   --病人医嘱记录..医生嘱托
  --   </YZXX>
  --  </YZLIST>
  -- </LJEXEC>
  --</OUTPUT>

  n_执行id  病人路径执行.Id%Type;
  v_Xtmp    Clob; --临时XML
  x_Templet Xmltype;
Begin
  Select Extractvalue(Value(A), 'IN/ZXID') Into n_执行id From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  For R In (Select a.执行者, a.执行人, a.执行时间, a.执行结果, a.执行说明 From 病人路径执行 A Where a.Id = n_执行id) Loop
    v_Xtmp := '<ZXQK>';
    v_Xtmp := v_Xtmp || '<ZXZ>' || r.执行者 || '</ZXZ>';
    v_Xtmp := v_Xtmp || '<ZXR>' || r.执行人 || '</ZXR>';
    v_Xtmp := v_Xtmp || '<ZXSJ>' || To_Char(r.执行时间, 'yyyy-mm-dd hh24:mi:ss') || '</ZXSJ>';
    v_Xtmp := v_Xtmp || '<ZXJG>' || r.执行结果 || '</ZXJG>';
    v_Xtmp := v_Xtmp || '<ZXSM>' || r.执行说明 || '</ZXSM>';
    v_Xtmp := v_Xtmp || '</ZXQK>';
  End Loop;

  x_Templet := Xmltype('<OUTPUT><LJEXEC>' || v_Xtmp || '</LJEXEC><YZLIST></YZLIST></OUTPUT>');

  --(西/成药，返回药品行，其它医嘱返回主医嘱行)
  For R In (Select a.医嘱期效, a.医嘱内容, a.单量, a.总量, a.结药途径, a.执行频次, a.时间方案, a.医生嘱托
            From (Select a.Id, a.相关id, a.诊疗类别, d.诊疗类别 As 主类别, a.医嘱期效, a.医嘱内容, a.单次用量 As 单量, a.总给予量 As 总量, e.名称 As 结药途径,
                          a.执行频次, a.执行时间方案 As 时间方案, a.医生嘱托, c.操作类型
                   From 病人医嘱记录 A, 病人路径医嘱 B, 诊疗项目目录 C, 病人医嘱记录 D, 诊疗项目目录 E
                   Where b.路径执行id = n_执行id And a.Id = b.病人医嘱id And a.诊疗项目id = c.Id(+) And a.相关id = d.Id(+) And
                         d.诊疗项目id = e.Id(+)) A
            Where a.相关id Is Null And Not (a.诊疗类别 = 'E' And a.操作类型 = '2')) Loop
    v_Xtmp := '<YZXX jsonArray="True" >';
    v_Xtmp := v_Xtmp || '<YZQX>' || r.医嘱期效 || '</YZQX>';
    v_Xtmp := v_Xtmp || '<YZNR>' || r.医嘱内容 || '</YZNR>';
    v_Xtmp := v_Xtmp || '<DL>' || r.单量 || '</DL>';
    v_Xtmp := v_Xtmp || '<ZL>' || r.总量 || '</ZL>';
    v_Xtmp := v_Xtmp || '<GYTJ>' || r.结药途径 || '</GYTJ>';
    v_Xtmp := v_Xtmp || '<ZXPL>' || r.执行频次 || '</ZXPL>';
    v_Xtmp := v_Xtmp || '<ZXSJ>' || r.时间方案 || '</ZXSJ>';
    v_Xtmp := v_Xtmp || '<YSZT>' || r.医生嘱托 || '</YSZT>';
    v_Xtmp := v_Xtmp || '</YZXX>';
    Select Appendchildxml(x_Templet, '/OUTPUT/YZLIST', Xmltype(v_Xtmp)) Into x_Templet From Dual;
  End Loop;
  Xml_Out := x_Templet;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getpathwaydetail;
/

--122937:胡俊勇,2018-03-15,护理接口出参添加的结点属性
CREATE OR REPLACE Procedure Zl_Third_Getdiagnosis
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --功能：获取病人诊断信息/查询
  --入参：Xml_In
  --<IN>
  --     <PATIID></PATIID>         --病人ID
  --     <PAGEID></PAGEID>     --主页ID
  --</IN>
  --出参：Xml_Out
  --<OUTPUT>
  --  <ZDLIST>
  --    <ZD>
  --      <ZDLX></ZDLX> --诊断类型。类型的名称，门诊诊断、入院诊断、出院诊断等
  --      <ZDCX></ZDCX> --诊断次序
  --      <ZDBM></ZDBM> --诊断编码
  --      <ZDMC></ZDMC> --诊断名称
  --    </ZD>
  --  </ZDLIST>
  --</OUTPUT>

  n_病人id   病人医嘱记录.病人id%Type;
  n_主页id   病人医嘱记录.主页id%Type;
  v_Xtmp     Varchar(5000); --临时XML
  v_Tmp      Varchar2(800);
  v_诊断编码 Varchar2(1000);
  v_诊断名称 Varchar2(1000);
  x_Templet  Xmltype;
Begin
  Select Extractvalue(Value(A), 'IN/PATIID'), Extractvalue(Value(A), 'IN/PAGEID')
  Into n_病人id, n_主页id
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  x_Templet := Xmltype('<OUTPUT><ZDLIST></ZDLIST></OUTPUT>');

  For R In (Select Decode(a.诊断类型, 1, '西医门诊诊断', 2, '西医入院诊断', 3, '西医出院诊断', 5, '院内感染', 6, '病理诊断', 7, '损伤中毒码', 8, '术前诊断', 9,
                           '术后诊断', 10, '并发症', 11, '中医门诊诊断', 12, '中医入院诊断', 13, '中医出院诊断', 21, '病原学诊断', 22, '影像学诊断') As 诊断类型,
                   a.诊断次序, a.诊断描述
            From 病人诊断记录 A
            Where a.病人id = n_病人id And a.主页id = n_主页id
            Order By a.诊断类型, a.诊断次序) Loop
  
    v_诊断编码 := Null;
    v_诊断名称 := r.诊断描述;
    v_Tmp      := r.诊断描述;
    If Substr(v_Tmp, 1, 1) = '(' Then
      v_诊断编码 := Substr(v_Tmp, 2, Instr(v_Tmp, ')') - 2);
      v_诊断名称 := Substr(v_Tmp, Instr(v_Tmp, ')') + 1);
    End If;
  
    v_Xtmp := '<ZD jsonArray="True" >';
    v_Xtmp := v_Xtmp || '<ZDLX>' || r.诊断类型 || '</ZDLX>';
    v_Xtmp := v_Xtmp || '<ZDCX>' || r.诊断次序 || '</ZDCX>';
    v_Xtmp := v_Xtmp || '<ZDBM>' || v_诊断编码 || '</ZDBM>';
    v_Xtmp := v_Xtmp || '<ZDMC>' || v_诊断名称 || '</ZDMC>';
    v_Xtmp := v_Xtmp || '</ZD>';
    
    Select Appendchildxml(x_Templet, '/OUTPUT/ZDLIST', Xmltype(v_Xtmp)) Into x_Templet From Dual;
  End Loop;
  Xml_Out := x_Templet;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getdiagnosis;
/

--122937:胡俊勇,2018-03-15,护理接口出参添加的结点属性
CREATE OR REPLACE Procedure Zl_Third_Getallergy
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --功能：获取病人过敏记录/查询
  --入参：Xml_In
  --<IN>
  --     <PATIID></PATIID>         --病人ID
  --     <PAGEID></PAGEID>     --主页ID
  --</IN>
  --出参：Xml_Out
  --<OUTPUT>
  --  <GMLIST>
  --    <GM>
  --      <GMYW></GMYW> --过敏药物
  --      <GMSJ></GMSJ> --过敏时间
  --      <JLSJ></JLSJ> --记录时间
  --      <JLR></JLR> --记录人
  --    </GM>
  --  </GMLIST>
  --</OUTPUT>
  n_病人id  病人医嘱记录.病人id%Type;
  n_主页id  病人医嘱记录.主页id%Type;
  v_Xtmp    Varchar(5000); --临时XML
  x_Templet Xmltype;
Begin
  Select Extractvalue(Value(A), 'IN/PATIID'), Extractvalue(Value(A), 'IN/PAGEID')
  Into n_病人id, n_主页id
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  x_Templet := Xmltype('<OUTPUT><GMLIST></GMLIST></OUTPUT>');
  For R In (Select a.药物名, a.过敏时间, a.记录时间, a.记录人
            From 病人过敏记录 A
            Where a.病人id = n_病人id And a.主页id = n_主页id
            Order By a.记录时间) Loop
    v_Xtmp := '<GM jsonArray="True" >';
    v_Xtmp := v_Xtmp || '<GMYW>' || r.药物名 || '</GMYW>'; --过敏药物
    v_Xtmp := v_Xtmp || '<GMSJ>' || To_Char(r.过敏时间, 'yyyy-mm-dd hh24:mi:ss') || '</GMSJ>'; --过敏时间
    v_Xtmp := v_Xtmp || '<JLSJ>' || To_Char(r.记录时间, 'yyyy-mm-dd hh24:mi:ss') || '</JLSJ>'; --记录时间
    v_Xtmp := v_Xtmp || '<JLR>' || r.记录人 || '</JLR>'; --记录人
    v_Xtmp := v_Xtmp || '</GM>';

    Select Appendchildxml(x_Templet, '/OUTPUT/GMLIST', Xmltype(v_Xtmp)) Into x_Templet From Dual;
  End Loop;
  Xml_Out := x_Templet;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getallergy;
/

--122937:胡俊勇,2018-03-15,护理接口出参添加的结点属性
Create Or Replace Procedure Zl_Third_Getpatichange
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --功能：获取病人变动记录/查询
  --入参：Xml_In
  --<IN>
  --     <PATIID></PATIID>         --病人ID
  --     <PAGEID></PAGEID>     --主页ID
  --</IN>
  --出参：Xml_Out
  --<OUTPUT>
  --  <BDLIST>--变动数据
  --    <ITEM>
  --      <SJ></SJ>--变动时间
  --      <LSXLIST>--变动类型列表
  --         <ITEM>
  --           <MC></MC>--变动类型名称（调整护理等级/调整住院医师/调整护士/留观转住院/调整主任医师/调整病况/调整医疗小组/调整病区/入院/入住/换床/调整床位等级/预出院/调整主治医师）
  --           <XXLIST>--变动信息列表
  --              <ITEM>
  --                <XXM></XXM>  ----信息名称（医疗小组/病区/科室/床号/床位等级/护理等级/护士/住院医师/主治医生/主任医生/当前病况/包床情况）         
  --                <YXX></YXX>  ----原信息值        
  --                <XXX></XXX>  ----现信息值             
  --              </ITEM> 
  --              ...
  --           </XXLIST>
  --         </ITEM>
  --         ...
  --      </LSXLIST>  
  --    </ITEM>
  --    ... 
  --  </BDLIST>
  --</OUTPUT>

  n_病人id 病人医嘱记录.病人id%Type;
  n_主页id 病人医嘱记录.主页id%Type;

  v_Tmp  Varchar(5000);
  v_Tmp1 Varchar(5000);

  v_Pre时间  Varchar(500);
  v_Cur时间  Varchar(500);
  v_变动名称 Varchar(500);

  n_Preidx Number;
  n_Curidx Number;

  v_Value    Varchar(500);
  x_Templet  Xmltype;
  x_变动时间 Xmltype;
  x_变动类型 Xmltype;

  Cursor c_Pati Is
    Select a.Id, f.名称 As 医疗小组名, b.名称 As 病区, c.名称 As 科室, a.附加床位, Decode(a.附加床位, 0, '主床', '包床') As 床位性质, a.床号,
           d.名称 As 床位等级, e.名称 As 护理等级, a.责任护士 As 护士, a.经治医师 As 住院医师, a.主治医师 As 主治医生, a.主任医师 As 主任医生, a.病情 As 当前病况,
           a.操作员姓名 As 开始操作员,
           Decode(a.开始原因, 1, '入院', 2, '入住', 3, '转科', 4, '换床', 5, '调整床位等级', 6, '调整护理等级', 7, '调整住院医师', 8, '调整护士', 9,
                   '留观转住院', 10, '预出院', 11, '调整主治医师', 12, '调整主任医师', 13, '调整病况', 14, '调整医疗小组', 15, '调整病区') As 开始原因,
           To_Char(a.开始时间, 'YYYY-MM-DD HH24:MI:SS') As 开始时间, a.终止人员 As 终止操作员,
           Decode(a.终止原因, 1, '出院', 2, '入住', 3, '转科', 4, '换床', 5, '调整床位等级', 6, '调整护理等级', 7, '调整住院医师', 8, '调整护士', 9,
                   '留观转住院', 10, '预出院', 11, '调整主治医师', 12, '调整主任医师', 13, '调整病况', 14, '调整医疗小组', 15, '调整病区') As 终止原因,
           To_Char(a.终止时间, 'YYYY-MM-DD HH24:MI:SS') As 终止时间
    From 病人变动记录 A, 部门表 B, 部门表 C, 收费项目目录 D, 收费项目目录 E, 临床医疗小组 F
    Where a.病区id = b.Id And a.科室id = c.Id And a.床位等级id = d.Id(+) And a.护理等级id = e.Id(+) And a.病人id = n_病人id And
          a.主页id = n_主页id And a.开始时间 Is Not Null And a.医疗小组id = f.Id(+)
    Order By a.终止时间, a.开始时间, a.附加床位, a.床号;

  Type t_Pati Is Table Of c_Pati%RowType;
  r_Pati t_Pati;
  r_Seek t_Pati;

Begin
  Select Extractvalue(Value(A), 'IN/PATIID') As 病人id, Extractvalue(Value(A), 'IN/PAGEID') As 主页id
  Into n_病人id, n_主页id
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;
  Open c_Pati;
  Fetch c_Pati Bulk Collect
    Into r_Pati;
  Close c_Pati;
  r_Seek := r_Pati;

  v_Pre时间 := '0';
  n_Preidx  := 0;
  n_Curidx  := 0;

  x_Templet := Xmltype('<OUTPUT><BDLIST></BDLIST></OUTPUT>');

  For I In 1 .. r_Pati.Count Loop
    --以时间点为一次变动
    If v_Pre时间 <> r_Pati(I).开始时间 And r_Pati(I).附加床位 = 0 Then
      v_Cur时间 := r_Pati(I).开始时间;
      --这中间查出变动情况
      x_变动时间 := Xmltype('<ITEM jsonArray="True" ><SJ>' || v_Cur时间 || '</SJ><LSXLIST></LSXLIST></ITEM>');
    
      v_变动名称 := '0';
      For J In 1 .. r_Seek.Count Loop
        If r_Seek(J).开始时间 = v_Cur时间 And r_Seek(J).附加床位 = 0 Then
          If v_变动名称 <> r_Seek(J).开始原因 Then
            v_变动名称 := r_Seek(J).开始原因;
            n_Curidx   := J;
            x_变动类型 := Xmltype('<ITEM jsonArray="True" ><MC>' || v_变动名称 || '</MC><XXLIST></XXLIST></ITEM>');
          
            --医疗小组名
            v_Value := Null;
            If n_Preidx <> 0 Then
              If Nvl(r_Seek(n_Preidx).医疗小组名, 'XXX') <> Nvl(r_Seek(n_Curidx).医疗小组名, 'XXX') Then
                v_Value := '<ITEM jsonArray="True" ><XXM>医疗小组名</XXM><YXX>' || r_Seek(n_Preidx).医疗小组名 || '</YXX><XXX>' || r_Seek(n_Curidx)
                          .医疗小组名 || '</XXX></ITEM>';
              End If;
            Else
              If r_Seek(n_Curidx).医疗小组名 Is Not Null Then
                v_Value := '<ITEM jsonArray="True" ><XXM>医疗小组名</XXM><YXX></YXX><XXX>' || r_Seek(n_Curidx).医疗小组名 || '</XXX></ITEM>';
              End If;
            End If;
            If v_Value Is Not Null Then
              Select Appendchildxml(x_变动类型, '/ITEM/XXLIST', Xmltype(v_Value)) Into x_变动类型 From Dual;
            End If;
          
            --病区
            v_Value := Null;
            If n_Preidx <> 0 Then
              If Nvl(r_Seek(n_Preidx).病区, 'XXX') <> Nvl(r_Seek(n_Curidx).病区, 'XXX') Then
                v_Value := '<ITEM jsonArray="True" ><XXM>病区</XXM><YXX>' || r_Seek(n_Preidx).病区 || '</YXX><XXX>' || r_Seek(n_Curidx).病区 ||
                           '</XXX></ITEM>';
              End If;
            Else
              If r_Seek(n_Curidx).病区 Is Not Null Then
                v_Value := '<ITEM jsonArray="True" ><XXM>病区</XXM><YXX></YXX><XXX>' || r_Seek(n_Curidx).病区 || '</XXX></ITEM>';
              End If;
            End If;
            If v_Value Is Not Null Then
              Select Appendchildxml(x_变动类型, '/ITEM/XXLIST', Xmltype(v_Value)) Into x_变动类型 From Dual;
            End If;
          
            --科室
            v_Value := Null;
            If n_Preidx <> 0 Then
              If Nvl(r_Seek(n_Preidx).科室, 'XXX') <> Nvl(r_Seek(n_Curidx).科室, 'XXX') Then
                v_Value := '<ITEM jsonArray="True" ><XXM>科室</XXM><YXX>' || r_Seek(n_Preidx).科室 || '</YXX><XXX>' || r_Seek(n_Curidx).科室 ||
                           '</XXX></ITEM>';
              End If;
            Else
              If r_Seek(n_Curidx).科室 Is Not Null Then
                v_Value := '<ITEM jsonArray="True" ><XXM>科室</XXM><YXX></YXX><XXX>' || r_Seek(n_Curidx).科室 || '</XXX></ITEM>';
              End If;
            End If;
            If v_Value Is Not Null Then
              Select Appendchildxml(x_变动类型, '/ITEM/XXLIST', Xmltype(v_Value)) Into x_变动类型 From Dual;
            End If;
          
            --床号
            v_Value := Null;
            If n_Preidx <> 0 Then
              If Nvl(r_Seek(n_Preidx).床号, 'XXX') <> Nvl(r_Seek(n_Curidx).床号, 'XXX') Then
                v_Value := '<ITEM jsonArray="True" ><XXM>床号</XXM><YXX>' || r_Seek(n_Preidx).床号 || '</YXX><XXX>' || r_Seek(n_Curidx).床号 ||
                           '</XXX></ITEM>';
              End If;
            Else
              If r_Seek(n_Curidx).床号 Is Not Null Then
                v_Value := '<ITEM jsonArray="True" ><XXM>床号</XXM><YXX></YXX><XXX>' || r_Seek(n_Curidx).床号 || '</XXX></ITEM>';
              End If;
            End If;
            If v_Value Is Not Null Then
              Select Appendchildxml(x_变动类型, '/ITEM/XXLIST', Xmltype(v_Value)) Into x_变动类型 From Dual;
            End If;
          
            --床位等级
            v_Value := Null;
            If n_Preidx <> 0 Then
              If Nvl(r_Seek(n_Preidx).床位等级, 'XXX') <> Nvl(r_Seek(n_Curidx).床位等级, 'XXX') Then
                v_Value := '<ITEM jsonArray="True" ><XXM>床位等级</XXM><YXX>' || r_Seek(n_Preidx).床位等级 || '</YXX><XXX>' || r_Seek(n_Curidx).床位等级 ||
                           '</XXX></ITEM>';
              End If;
            Else
              If r_Seek(n_Curidx).床位等级 Is Not Null Then
                v_Value := '<ITEM jsonArray="True" ><XXM>床位等级</XXM><YXX></YXX><XXX>' || r_Seek(n_Curidx).床位等级 || '</XXX></ITEM>';
              End If;
            End If;
            If v_Value Is Not Null Then
              Select Appendchildxml(x_变动类型, '/ITEM/XXLIST', Xmltype(v_Value)) Into x_变动类型 From Dual;
            End If;
          
            --护理等级
            v_Value := Null;
            If n_Preidx <> 0 Then
              If Nvl(r_Seek(n_Preidx).护理等级, 'XXX') <> Nvl(r_Seek(n_Curidx).护理等级, 'XXX') Then
                v_Value := '<ITEM jsonArray="True" ><XXM>护理等级</XXM><YXX>' || r_Seek(n_Preidx).护理等级 || '</YXX><XXX>' || r_Seek(n_Curidx).护理等级 ||
                           '</XXX></ITEM>';
              End If;
            Else
              If r_Seek(n_Curidx).护理等级 Is Not Null Then
                v_Value := '<ITEM jsonArray="True" ><XXM>护理等级</XXM><YXX></YXX><XXX>' || r_Seek(n_Curidx).护理等级 || '</XXX></ITEM>';
              End If;
            End If;
            If v_Value Is Not Null Then
              Select Appendchildxml(x_变动类型, '/ITEM/XXLIST', Xmltype(v_Value)) Into x_变动类型 From Dual;
            End If;
          
            --护士
            v_Value := Null;
            If n_Preidx <> 0 Then
              If Nvl(r_Seek(n_Preidx).护士, 'XXX') <> Nvl(r_Seek(n_Curidx).护士, 'XXX') Then
                v_Value := '<ITEM jsonArray="True" ><XXM>护士</XXM><YXX>' || r_Seek(n_Preidx).护士 || '</YXX><XXX>' || r_Seek(n_Curidx).护士 ||
                           '</XXX></ITEM>';
              End If;
            Else
              If r_Seek(n_Curidx).护士 Is Not Null Then
                v_Value := '<ITEM jsonArray="True" ><XXM>护士</XXM><YXX></YXX><XXX>' || r_Seek(n_Curidx).护士 || '</XXX></ITEM>';
              End If;
            End If;
            If v_Value Is Not Null Then
              Select Appendchildxml(x_变动类型, '/ITEM/XXLIST', Xmltype(v_Value)) Into x_变动类型 From Dual;
            End If;
          
            --住院医师
            v_Value := Null;
            If n_Preidx <> 0 Then
              If Nvl(r_Seek(n_Preidx).住院医师, 'XXX') <> Nvl(r_Seek(n_Curidx).住院医师, 'XXX') Then
                v_Value := '<ITEM jsonArray="True" ><XXM>住院医师</XXM><YXX>' || r_Seek(n_Preidx).住院医师 || '</YXX><XXX>' || r_Seek(n_Curidx).住院医师 ||
                           '</XXX></ITEM>';
              End If;
            Else
              If r_Seek(n_Curidx).住院医师 Is Not Null Then
                v_Value := '<ITEM jsonArray="True" ><XXM>住院医师</XXM><YXX></YXX><XXX>' || r_Seek(n_Curidx).住院医师 || '</XXX></ITEM>';
              End If;
            End If;
            If v_Value Is Not Null Then
              Select Appendchildxml(x_变动类型, '/ITEM/XXLIST', Xmltype(v_Value)) Into x_变动类型 From Dual;
            End If;
          
            --主治医生
            v_Value := Null;
            If n_Preidx <> 0 Then
              If Nvl(r_Seek(n_Preidx).主治医生, 'XXX') <> Nvl(r_Seek(n_Curidx).主治医生, 'XXX') Then
                v_Value := '<ITEM jsonArray="True" ><XXM>主治医生</XXM><YXX>' || r_Seek(n_Preidx).主治医生 || '</YXX><XXX>' || r_Seek(n_Curidx).主治医生 ||
                           '</XXX></ITEM>';
              End If;
            Else
              If r_Seek(n_Curidx).主治医生 Is Not Null Then
                v_Value := '<ITEM jsonArray="True" ><XXM>主治医生</XXM><YXX></YXX><XXX>' || r_Seek(n_Curidx).主治医生 || '</XXX></ITEM>';
              End If;
            End If;
            If v_Value Is Not Null Then
              Select Appendchildxml(x_变动类型, '/ITEM/XXLIST', Xmltype(v_Value)) Into x_变动类型 From Dual;
            End If;
          
            --主任医生
            v_Value := Null;
            If n_Preidx <> 0 Then
              If Nvl(r_Seek(n_Preidx).主任医生, 'XXX') <> Nvl(r_Seek(n_Curidx).主任医生, 'XXX') Then
                v_Value := '<ITEM jsonArray="True" ><XXM>主任医生</XXM><YXX>' || r_Seek(n_Preidx).主任医生 || '</YXX><XXX>' || r_Seek(n_Curidx).主任医生 ||
                           '</XXX></ITEM>';
              End If;
            Else
              If r_Seek(n_Curidx).主任医生 Is Not Null Then
                v_Value := '<ITEM jsonArray="True" ><XXM>主任医生</XXM><YXX></YXX><XXX>' || r_Seek(n_Curidx).主任医生 || '</XXX></ITEM>';
              End If;
            End If;
            If v_Value Is Not Null Then
              Select Appendchildxml(x_变动类型, '/ITEM/XXLIST', Xmltype(v_Value)) Into x_变动类型 From Dual;
            End If;
          
            --当前病况
            v_Value := Null;
            If n_Preidx <> 0 Then
              If Nvl(r_Seek(n_Preidx).当前病况, 'XXX') <> Nvl(r_Seek(n_Curidx).当前病况, 'XXX') Then
                v_Value := '<ITEM jsonArray="True" ><XXM>当前病况</XXM><YXX>' || r_Seek(n_Preidx).当前病况 || '</YXX><XXX>' || r_Seek(n_Curidx).当前病况 ||
                           '</XXX></ITEM>';
              End If;
            Else
              If r_Seek(n_Curidx).当前病况 Is Not Null Then
                v_Value := '<ITEM jsonArray="True" ><XXM>当前病况</XXM><YXX></YXX><XXX>' || r_Seek(n_Curidx).当前病况 || '</XXX></ITEM>';
              End If;
            End If;
            If v_Value Is Not Null Then
              Select Appendchildxml(x_变动类型, '/ITEM/XXLIST', Xmltype(v_Value)) Into x_变动类型 From Dual;
            End If;
          
            ----包床情况  用床号拼串即可（简单处理）
            v_Value := Null;
            v_Tmp   := Null;
            v_Tmp1  := Null;
            If n_Preidx <> 0 Then
              --检查之前的变动是否是包了床,用床号拼串即可（简单处理）
              For K In 1 .. r_Seek.Count Loop
                If v_变动名称 = r_Seek(K).终止原因 And v_Cur时间 = r_Seek(K).终止时间 And r_Seek(K).附加床位 = 1 Then
                  If v_Tmp Is Null Then
                    v_Tmp := r_Seek(K).床号;
                  Else
                    v_Tmp := v_Tmp || ',' || r_Seek(K).床号;
                  End If;
                End If;
              End Loop;
            End If;
            --检查当前的变动是否是包了床,用床号拼串即可（简单处理）
            For K In 1 .. r_Seek.Count Loop
              If v_变动名称 = r_Seek(K).开始原因 And v_Cur时间 = r_Seek(K).开始时间 And r_Seek(K).附加床位 = 1 Then
                If v_Tmp1 Is Null Then
                  v_Tmp1 := r_Seek(K).床号;
                Else
                  v_Tmp1 := v_Tmp1 || ',' || r_Seek(K).床号;
                End If;
              End If;
            End Loop;
            If Nvl(v_Tmp, 'XXX') <> Nvl(v_Tmp1, 'XXX') Then
              v_Value := '<ITEM jsonArray="True" ><XXM>包床情况</XXM><YXX>' || v_Tmp || '</YXX><XXX>' || v_Tmp1 || '</XXX></ITEM>';
            End If;
            If v_Value Is Not Null Then
              Select Appendchildxml(x_变动类型, '/ITEM/XXLIST', Xmltype(v_Value)) Into x_变动类型 From Dual;
            End If;
          
            Select Appendchildxml(x_变动时间, '/ITEM/LSXLIST', x_变动类型) Into x_变动时间 From Dual;
          End If;
        End If;
      End Loop;
    
      v_Pre时间 := v_Cur时间;
      n_Preidx  := I;
    
      Select Appendchildxml(x_Templet, '/OUTPUT/BDLIST', x_变动时间) Into x_Templet From Dual;
    End If;
  End Loop;
  Xml_Out := x_Templet;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getpatichange;
/

--122937:胡俊勇,2018-03-15,护理接口出参添加的结点属性
CREATE OR REPLACE Procedure Zl_Third_Getoperation
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --功能：获取病人手术信息/查询
  --用于获取指定病人的手术信息
  --当监听到病人发送手术医嘱时调用
  --当监听到病人完成手术时调用
  --需要手动同步时调用
  --获取所有病人信息后调用
  --入参：Xml_In
  --<IN>
  --  <PATIID></PATIID>     --病人ID
  --  <PAGEID></PAGEID>     --主页ID
  --</IN>

  --出参：Xml_Out 
  --<OUTPUT>
  --  <SSLIST>
  --    <SS>
  --      <SSMC></SSMC>  //手术名称
  --      <SSSJ></SSSJ>  //手术时间,yyyy-mm-dd hh24:mi
  --      <MZFS></MZFS>  //麻醉方式
  --      <SSQK></SSQK>  //手术情况  择期、急诊、限期
  --      <ZXKSID></ZXKSID>   //执行科室ID
  --      <ZXKSMC></ZXKSMC>    //执行科室名称
  --      <FJSS></FJSS>     //附加手术   1-是，0-否
  --    <SS>
  --  <SSLIST>
  --</OUTPUT>

  n_病人id   病人医嘱记录.病人id%Type;
  n_主页id   病人医嘱记录.主页id%Type;
  n_主医嘱id 病人医嘱记录.Id%Type;
  v_麻醉     诊疗项目目录.名称%Type;
  v_Xtmp     Clob; --临时XML 

  Cursor c_医嘱 Is
    Select b.名称, a.标本部位 As 手术时间, Nvl(a.相关id, a.Id) As 主医嘱id, a.执行科室id, c.名称 As 执行科室名称,
           Decode(a.相关id, Null, 0, 1) As 附加手术, a.诊疗类别, Decode(a.手术情况, Null, '择期', 1, '急诊', 2, '限期', Null) As 手术情况
    From 病人医嘱记录 A, 诊疗项目目录 B, 部门表 C
    Where a.诊疗项目id = b.Id And a.执行科室id = c.Id And a.诊疗类别 In ('F', 'G') And a.医嘱状态 <> 4 And Nvl(a.执行标记, 0) <> -1 And
          a.病人id = n_病人id And a.主页id = n_主页id
    Order By a.诊疗类别 Desc, a.序号;
Begin

  Select Extractvalue(Value(A), 'IN/PATIID') As 病人id, Extractvalue(Value(A), 'IN/PAGEID') As 主页id
  Into n_病人id, n_主页id
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  n_主医嘱id := 0;

  For R In c_医嘱 Loop
    If n_主医嘱id <> r.主医嘱id Then
      n_主医嘱id := r.主医嘱id;
      If r.诊疗类别 = 'G' Then
        v_麻醉 := r.名称;
      End If;
    Else
      v_麻醉 := Null;
    End If;
  
    If r.诊疗类别 = 'F' Then
      v_Xtmp := v_Xtmp || '<SS jsonArray="True" >';
      v_Xtmp := v_Xtmp || '<SSMC>' || r.名称 || '</SSMC>'; --  //手术名称
      v_Xtmp := v_Xtmp || '<SSSJ>' || r.手术时间 || '</SSSJ>'; --  //手术时间,yyyy-mm-dd hh24:mi
      v_Xtmp := v_Xtmp || '<MZFS>' || v_麻醉 || '</MZFS>'; --  //麻醉方式
      v_Xtmp := v_Xtmp || '<SSQK>' || r.手术情况 || '</SSQK>'; --  //手术情况  择期、急诊、限期
      v_Xtmp := v_Xtmp || '<ZXKSID>' || r.执行科室id || '</ZXKSID>'; --   //执行科室ID
      v_Xtmp := v_Xtmp || '<ZXKSMC>' || r.执行科室名称 || '</ZXKSMC>'; --    //执行科室名称
      v_Xtmp := v_Xtmp || '<FJSS>' || r.附加手术 || '</FJSS>'; --    //附加手术   1-是，0-否    
      v_Xtmp := v_Xtmp || '</SS>';
    End If;
  End Loop;

  If v_Xtmp Is Not Null Then
    Xml_Out := Xmltype('<OUTPUT><SSLIST>' || v_Xtmp || '</SSLIST></OUTPUT>');
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getoperation;
/

--122937:胡俊勇,2018-03-15,护理接口出参添加的结点属性
CREATE OR REPLACE Procedure Zl_Third_Getallpatiinfo
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --功能：获取病区所有病人基本信息/查询
  --用于获取某病区所有病人的基本信息。
  --在第一次获取病人信息时调用
  --每天晚上自动同步时调用
  --需要手动同步时调用

  --入参：Xml_In
  --<INPUT>
  --  <BQID></BQID>      --病区ID
  --  <CYTS></CYTS>      --出院天数，获取多少天之内出院的病人
  --</INPUT>

  --出参：Xml_Out 
  --<OUTPUT>
  --  <PATILIST>
  --    <PATI>
  --      <JBXX>    --基本信息
  --        <PATIID></PATIID>         --病人ID
  --        <PAGEID></PAGEID>  --主页ID
  --        <BABY></BABY>  --婴儿序号
  --        <XM></XM>   --姓名
  --        <XB></XB>   --性别
  --        <NL></NL>   --年龄
  --        <CSRQ></CSRQ>  --出生日期
  --        <ZYH></ZYH>  --住院号
  --        <HY></HY>   --婚姻
  --        <GJ></GJ>   --国籍
  --        <MZ></MZ>   --民族
  --        <XL></XL>   --学历
  --        <SF></SF>   --身份
  --        <ZY></ZY>   --职业
  --        <SFZH></SFZH>  --身份证号
  --        <FKFS></FKFS>  --付款方式
  --        <LXFS></LXFS>  --联系方式
  --        <LXRXM></LXRXM>  --联系人姓名
  --        <LXRDH></LXRDH>  --联系人电话
  --        <LXRDZ></LXRDZ>  --联系人地址
  --        <JTDH></JTDH>  --家庭电话
  --        <JTDZ></JTDZ>  --家庭地址
  --        <CSDD></CSDD>  --出生地点
  --        <GMS></GMS>  --过敏史
  --      </JBXX>
  --      <ZYXX>    --住院信息
  --        <RYRQ></RYRQ>  --入院日期
  --        <RKRQ></RKRQ>  --入科日期
  --        <CYRQ></CYRQ>  --出院日期
  --        <ZYTS></ZYTS>  --住院天数
  --        <RYFS></RYFS>  --入院方式
  --        <KSID></KSID>  --科室ID
  --        <KSMC></KSMC>  --科室名称
  --        <BQID></BQID>  --病区ID
  --        <BQMC></BQMC>  --病区名称
  --        <CH></CH>   --床号
  --        <BQ></BQ>   --病情
  --        <ZZYS></ZZYS>  --主治医师
  --        <ZRYS></ZRYS>  --主任医师
  --        <ZYYS></ZYYS>  --住院医师
  --        <ZRHS></ZRHS>  --责任护士
  --        <HLDJ></HLDJ>  --护理等级
  --        <YLZ></YLZ>    --医疗小组id
  --        <YBH></YBH>    --医保号
  --        <YBMC></YBMC>  --医保名称
  --      </ZYXX>
  --    </PATI>
  --  </PATILIST>
  --</OUTPUT>

  n_病区id 部门表.Id%Type;
  v_病区   部门表.名称%Type;
  n_天数   Number;
  v_Xtmp   Clob; --临时XML 
  x_Item   Xmltype;
  d_开始   Date;
  d_结束   Date;

  v_过敏信息 Varchar2(5000);
  v_主治医师 Varchar2(500);
  v_主任医师 Varchar2(500);
  x_Templet  Xmltype;

  Cursor c_在院 Is
    Select a.病人id, b.主页id, 0 As 婴儿序号, Nvl(b.姓名, a.姓名) As 姓名, Nvl(b.性别, a.性别) As 性别, Nvl(b.年龄, a.年龄) As 年龄, a.出生日期, b.住院号,
           a.婚姻状况 As 婚姻, a.国籍, a.民族, a.学历, a.身份, a.职业, a.身份证号, a.医疗付款方式 As 付款方式, a.手机号 As 联系方式, a.联系人姓名, a.联系人电话,
           a.联系人地址, a.家庭电话, a.家庭地址, a.出生地点, '待定单独查询' As 过敏史, Decode(b.入科时间, Null, b.入院日期, b.入科时间) As 入院日期,
           b.入科时间 As 入科日期, Null As 出院日期, (Trunc(Sysdate) - Trunc(Decode(b.入科时间, Null, b.入院日期, b.入科时间))) As 住院天数, b.入院方式,
           b.出院科室id As 科室id, c.名称 As 科室名称, r.病区id, v_病区 As 病区名称, b.出院病床 As 床号, b.当前病况 As 病情, '待定病案主页从表' As 主治医师,
           '待定病案主页从表' As 主任医师, b.住院医师, b.责任护士, e.名称 As 护理等级, b.医疗小组id, a.医保号, d.名称 As 医保名称

    From 病人信息 A, 病案主页 B, 部门表 C, 保险类别 D, 收费项目目录 E, 在院病人 R
    Where a.病人id = b.病人id And a.主页id = b.主页id And b.出院科室id = c.Id And b.险类 = d.序号(+) And Nvl(b.状态, 0) <> 1 And
          b.护理等级id = e.Id(+) And (r.病区id = n_病区id Or b.婴儿病区id = n_病区id) And a.病人id = r.病人id And a.当前病区id + 0 = r.病区id And
          Nvl(b.病案状态, 0) <> 5 And b.封存时间 Is Null
    Order By b.出院病床;

  Cursor c_出院 Is
    Select a.病人id, b.主页id, 0 As 婴儿序号, Nvl(b.姓名, a.姓名) As 姓名, Nvl(b.性别, a.性别) As 性别, Nvl(b.年龄, a.年龄) As 年龄, a.出生日期, b.住院号,
           a.婚姻状况 As 婚姻, a.国籍, a.民族, a.学历, a.身份, a.职业, a.身份证号, a.医疗付款方式 As 付款方式, a.手机号 As 联系方式, a.联系人姓名, a.联系人电话,
           a.联系人地址, a.家庭电话, a.家庭地址, a.出生地点, '待定单独查询' As 过敏史, Decode(b.入科时间, Null, b.入院日期, b.入科时间) As 入院日期,
           b.入科时间 As 入科日期, b.出院日期, (Trunc(b.出院日期) - Trunc(Decode(b.入科时间, Null, b.入院日期, b.入科时间))) As 住院天数, b.入院方式,
           b.出院科室id As 科室id, c.名称 As 科室名称, b.当前病区id As 病区id, v_病区 As 病区名称, b.出院病床 As 床号, b.当前病况 As 病情,
           '待定病案主页从表' As 主治医师, '待定病案主页从表' As 主任医师, b.住院医师, b.责任护士, e.名称 As 护理等级, b.医疗小组id, a.医保号, d.名称 As 医保名称
    From 病人信息 A, 病案主页 B, 部门表 C, 保险类别 D, 收费项目目录 E
    Where a.病人id = b.病人id And Nvl(b.主页id, 0) <> 0 And b.状态 = 0 And b.出院科室id = c.Id And b.险类 = d.序号(+) And
          b.护理等级id = e.Id(+) And b.当前病区id + 0 = n_病区id And Nvl(b.病案状态, 0) <> 5 And b.封存时间 Is Null And
          b.出院日期 Between d_开始 And d_结束
    Order By b.出院病床;

  --放到循环中执行的，可能有性能问题
  Procedure p_Getother
  (
    病人id_In    In 病人信息.病人id%Type,
    主页id_In    In 病案主页.主页id%Type,
    过敏信息_Out Out Varchar2,
    主治医师_Out Out Varchar2,
    主任医师_Out Out Varchar2
  ) Is
  Begin
  
    过敏信息_Out := Null;
    主治医师_Out := Null;
    主任医师_Out := Null;
 
    For R In (Select a.信息名, a.信息值
              From 病案主页从表 A
              Where a.病人id = 病人id_In And a.主页id = 主页id_In And a.信息名 In ('主治医师', '主任医师')) Loop
      If r.信息名 = '主治医师' Then
        主治医师_Out := r.信息值;
      Elsif r.信息名 = '主任医师' Then
        主任医师_Out := r.信息值;
      End If;
    End Loop;
  
    For R In (Select a.药物名
              From 病人过敏记录 A, 病人挂号记录 B, 病案主页 C
              Where a.病人id = b.病人id(+) And a.主页id = b.Id(+) And b.记录性质(+) = 1 And b.记录状态(+) = 1 And a.病人id = c.病人id(+) And
                    a.主页id = c.主页id(+) And a.结果 = 1 And 药物名 Is Not Null And a.病人id = 202 And Not Exists
               (Select 药物id
                     From 病人过敏记录
                     Where (Nvl(药物id, 0) = Nvl(a.药物id, 0) Or Nvl(药物名, 'Null') = Nvl(a.药物名, 'Null')) And Nvl(结果, 0) = 0 And
                           记录时间 > a.记录时间 And 病人id = 202)
              Group By a.药物名
              Order By a.药物名) Loop
      过敏信息_Out := 过敏信息_Out || ',' || r.药物名;
    End Loop;
    过敏信息_Out := Substr(过敏信息_Out, 2);
  End;

Begin
  Select Extractvalue(Value(A), 'IN/BQID') As 病区id, Extractvalue(Value(A), 'IN/CYTS') As 天数
  Into n_病区id, n_天数
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  Select Sysdate Into d_结束 From Dual;

  d_开始 := Trunc(d_结束) - n_天数; --当天的 00:00:00  
  d_结束 := Trunc(d_结束) + 1 - 1 / 24 / 60; --当天的 23:59:59

  Select Max(a.名称) Into v_病区 From 部门表 A Where a.Id = n_病区id;

  x_Templet := Xmltype('<OUTPUT></OUTPUT>');
  x_Item    := Xmltype('<PATILIST></PATILIST>');

  For R In c_在院 Loop
    p_Getother(r.病人id, r.主页id, v_过敏信息, v_主治医师, v_主任医师);
    v_Xtmp := '<PATI jsonArray="True" >';
    v_Xtmp := v_Xtmp || '<JBXX>';
    v_Xtmp := v_Xtmp || '<PATIID>' || r.病人id || '</PATIID>';
    v_Xtmp := v_Xtmp || '<PAGEID>' || r.主页id || '</PAGEID>';
    v_Xtmp := v_Xtmp || '<BABY>' || r.婴儿序号 || '</BABY>';
    v_Xtmp := v_Xtmp || '<XM>' || r.姓名 || '</XM>';
    v_Xtmp := v_Xtmp || '<XB>' || r.性别 || '</XB>';
    v_Xtmp := v_Xtmp || '<NL>' || r.年龄 || '</NL>';
    v_Xtmp := v_Xtmp || '<CSRQ>' || To_Char(r.出生日期, 'yyyy-mm-dd hh24:mi:ss') || '</CSRQ>';
    v_Xtmp := v_Xtmp || '<ZYH>' || r.住院号 || '</ZYH>';
    v_Xtmp := v_Xtmp || '<HY>' || r.婚姻 || '</HY>';
    v_Xtmp := v_Xtmp || '<GJ>' || r.国籍 || '</GJ>';
    v_Xtmp := v_Xtmp || '<MZ>' || r.民族 || '</MZ>';
    v_Xtmp := v_Xtmp || '<XL>' || r.学历 || '</XL>';
    v_Xtmp := v_Xtmp || '<SF>' || r.身份 || '</SF>';
    v_Xtmp := v_Xtmp || '<ZY>' || r.职业 || '</ZY>';
    v_Xtmp := v_Xtmp || '<SFZH>' || r.身份证号 || '</SFZH>';
    v_Xtmp := v_Xtmp || '<FKFS>' || r.付款方式 || '</FKFS>';
    v_Xtmp := v_Xtmp || '<LXFS>' || r.联系方式 || '</LXFS>';
    v_Xtmp := v_Xtmp || '<LXRXM>' || r.联系人姓名 || '</LXRXM>';
    v_Xtmp := v_Xtmp || '<LXRDH>' || r.联系人电话 || '</LXRDH>';
    v_Xtmp := v_Xtmp || '<LXRDZ>' || r.联系人地址 || '</LXRDZ>';
    v_Xtmp := v_Xtmp || '<JTDH>' || r.家庭电话 || '</JTDH>';
    v_Xtmp := v_Xtmp || '<JTDZ>' || r.家庭地址 || '</JTDZ>';
    v_Xtmp := v_Xtmp || '<CSDD>' || r.出生地点 || '</CSDD>';
    v_Xtmp := v_Xtmp || '<GMS>' || v_过敏信息 || '</GMS>'; -- r.过敏史 
    v_Xtmp := v_Xtmp || '</JBXX>';
    v_Xtmp := v_Xtmp || '<ZYXX>';
    v_Xtmp := v_Xtmp || '<RYRQ>' || To_Char(r.入院日期, 'yyyy-mm-dd hh24:mi:ss') || '</RYRQ>';
    v_Xtmp := v_Xtmp || '<RKRQ>' || To_Char(r.入科日期, 'yyyy-mm-dd hh24:mi:ss') || '</RKRQ>';
    v_Xtmp := v_Xtmp || '<CYRQ>' || To_Char(r.出院日期, 'yyyy-mm-dd hh24:mi:ss') || '</CYRQ>';
    v_Xtmp := v_Xtmp || '<ZYTS>' || r.住院天数 || '</ZYTS>';
    v_Xtmp := v_Xtmp || '<RYFS>' || r.入院方式 || '</RYFS>';
    v_Xtmp := v_Xtmp || '<KSID>' || r.科室id || '</KSID>';
    v_Xtmp := v_Xtmp || '<KSMC>' || r.科室名称 || '</KSMC>';
    v_Xtmp := v_Xtmp || '<BQID>' || r.病区id || '</BQID>';
    v_Xtmp := v_Xtmp || '<BQMC>' || r.病区名称 || '</BQMC>';
    v_Xtmp := v_Xtmp || '<CH>' || r.床号 || '</CH>';
    v_Xtmp := v_Xtmp || '<BQ>' || r.病情 || '</BQ>';
    v_Xtmp := v_Xtmp || '<ZZYS>' || v_主治医师 || '</ZZYS>';
    v_Xtmp := v_Xtmp || '<ZRYS>' || v_主任医师 || '</ZRYS>';
    v_Xtmp := v_Xtmp || '<ZYYS>' || r.住院医师 || '</ZYYS>';
    v_Xtmp := v_Xtmp || '<ZRHS>' || r.责任护士 || '</ZRHS>';
    v_Xtmp := v_Xtmp || '<HLDJ>' || r.护理等级 || '</HLDJ>';
    v_Xtmp := v_Xtmp || '<YLZ>' || r.医疗小组id || '</YLZ>';
    v_Xtmp := v_Xtmp || '<YBH>' || r.医保号 || '</YBH>';
    v_Xtmp := v_Xtmp || '<YBMC>' || r.医保名称 || '</YBMC>';
    v_Xtmp := v_Xtmp || '</ZYXX>';
    v_Xtmp := v_Xtmp || '</PATI>';
    Select Appendchildxml(x_Item, '/PATILIST', Xmltype(v_Xtmp)) Into x_Item From Dual;
  End Loop;

  For R In c_出院 Loop
    p_Getother(r.病人id, r.主页id, v_过敏信息, v_主治医师, v_主任医师);
    v_Xtmp := '<PATI jsonArray="True" >';
    v_Xtmp := v_Xtmp || '<JBXX>';
    v_Xtmp := v_Xtmp || '<PATIID>' || r.病人id || '</PATIID>';
    v_Xtmp := v_Xtmp || '<PAGEID>' || r.主页id || '</PAGEID>';
    v_Xtmp := v_Xtmp || '<BABY>' || r.婴儿序号 || '</BABY>';
    v_Xtmp := v_Xtmp || '<XM>' || r.姓名 || '</XM>';
    v_Xtmp := v_Xtmp || '<XB>' || r.性别 || '</XB>';
    v_Xtmp := v_Xtmp || '<NL>' || r.年龄 || '</NL>';
    v_Xtmp := v_Xtmp || '<CSRQ>' || To_Char(r.出生日期, 'yyyy-mm-dd hh24:mi:ss') || '</CSRQ>';
    v_Xtmp := v_Xtmp || '<ZYH>' || r.住院号 || '</ZYH>';
    v_Xtmp := v_Xtmp || '<HY>' || r.婚姻 || '</HY>';
    v_Xtmp := v_Xtmp || '<GJ>' || r.国籍 || '</GJ>';
    v_Xtmp := v_Xtmp || '<MZ>' || r.民族 || '</MZ>';
    v_Xtmp := v_Xtmp || '<XL>' || r.学历 || '</XL>';
    v_Xtmp := v_Xtmp || '<SF>' || r.身份 || '</SF>';
    v_Xtmp := v_Xtmp || '<ZY>' || r.职业 || '</ZY>';
    v_Xtmp := v_Xtmp || '<SFZH>' || r.身份证号 || '</SFZH>';
    v_Xtmp := v_Xtmp || '<FKFS>' || r.付款方式 || '</FKFS>';
    v_Xtmp := v_Xtmp || '<LXFS>' || r.联系方式 || '</LXFS>';
    v_Xtmp := v_Xtmp || '<LXRXM>' || r.联系人姓名 || '</LXRXM>';
    v_Xtmp := v_Xtmp || '<LXRDH>' || r.联系人电话 || '</LXRDH>';
    v_Xtmp := v_Xtmp || '<LXRDZ>' || r.联系人地址 || '</LXRDZ>';
    v_Xtmp := v_Xtmp || '<JTDH>' || r.家庭电话 || '</JTDH>';
    v_Xtmp := v_Xtmp || '<JTDZ>' || r.家庭地址 || '</JTDZ>';
    v_Xtmp := v_Xtmp || '<CSDD>' || r.出生地点 || '</CSDD>';
    v_Xtmp := v_Xtmp || '<GMS>' || v_过敏信息 || '</GMS>'; -- r.过敏史 
    v_Xtmp := v_Xtmp || '</JBXX>';
    v_Xtmp := v_Xtmp || '<ZYXX>';
    v_Xtmp := v_Xtmp || '<RYRQ>' || To_Char(r.入院日期, 'yyyy-mm-dd hh24:mi:ss') || '</RYRQ>';
    v_Xtmp := v_Xtmp || '<RKRQ>' || To_Char(r.入科日期, 'yyyy-mm-dd hh24:mi:ss') || '</RKRQ>';
    v_Xtmp := v_Xtmp || '<CYRQ>' || To_Char(r.出院日期, 'yyyy-mm-dd hh24:mi:ss') || '</CYRQ>';
    v_Xtmp := v_Xtmp || '<ZYTS>' || r.住院天数 || '</ZYTS>';
    v_Xtmp := v_Xtmp || '<RYFS>' || r.入院方式 || '</RYFS>';
    v_Xtmp := v_Xtmp || '<KSID>' || r.科室id || '</KSID>';
    v_Xtmp := v_Xtmp || '<KSMC>' || r.科室名称 || '</KSMC>';
    v_Xtmp := v_Xtmp || '<BQID>' || r.病区id || '</BQID>';
    v_Xtmp := v_Xtmp || '<BQMC>' || r.病区名称 || '</BQMC>';
    v_Xtmp := v_Xtmp || '<CH>' || r.床号 || '</CH>';
    v_Xtmp := v_Xtmp || '<BQ>' || r.病情 || '</BQ>';
    v_Xtmp := v_Xtmp || '<ZZYS>' || v_主治医师 || '</ZZYS>';
    v_Xtmp := v_Xtmp || '<ZRYS>' || v_主任医师 || '</ZRYS>';
    v_Xtmp := v_Xtmp || '<ZYYS>' || r.住院医师 || '</ZYYS>';
    v_Xtmp := v_Xtmp || '<ZRHS>' || r.责任护士 || '</ZRHS>';
    v_Xtmp := v_Xtmp || '<HLDJ>' || r.护理等级 || '</HLDJ>';
    v_Xtmp := v_Xtmp || '<YLZ>' || r.医疗小组id || '</YLZ>';
    v_Xtmp := v_Xtmp || '<YBH>' || r.医保号 || '</YBH>';
    v_Xtmp := v_Xtmp || '<YBMC>' || r.医保名称 || '</YBMC>';
    v_Xtmp := v_Xtmp || '</ZYXX>';
    v_Xtmp := v_Xtmp || '</PATI>';
    Select Appendchildxml(x_Item, '/PATILIST', Xmltype(v_Xtmp)) Into x_Item From Dual;
  End Loop;

  Select Appendchildxml(x_Templet, '/OUTPUT', x_Item) Into x_Templet From Dual;
  Xml_Out := x_Templet;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getallpatiinfo;
/

--122937:胡俊勇,2018-03-15,护理接口出参添加的结点属性
Create Or Replace Procedure Zl_Third_Getkfcws
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is

  --功能：开放床位数/查询
  --开放床位数：每日夜晚12点开放病床数总和，不论该床是否被病人占用,都应计算在内。包括消毒和小修理等暂停使用的病床，
  --             以及超过半年的加床。不包括因病房扩建或大修而停用的病床及临时增设病床

  --平均开放床位数：每月平均开放床位数=当月每日开放床位数之和/当月天数
  --基于ZLHIS系统数据结构的理解：床位增减记录，中增加的床位记录
  --入参：xml_in
  --<IN>
  --    <BQID></BQID>    //病区ID，传空取所有病区
  --    <KSRQ></KSRQ>  //开始日期   yyyy-mm-dd
  --    <JSRQ></JSRQ>   //结束日期  yyyy-mm-dd
  --</IN>

  --出参：xml_out
  --<OUTPUT>
  --  <BQLIST>
  --    <ITEM>
  --      <BQID></BQID>  //病区ID
  --      <BQMC></BQMC>  //病区名称
  --      <DATALIST>
  --        <ITEM>
  --           <YF></YF>  //月份
  --           <KFCR></KFCR>  //开放床位数
  --        </ITEM>
  --      </DATALIST>
  --    </ITEM>
  --  </BQLIST>
  --</OUTPUT>

  n_病区id   部门表.Id%Type;
  d_开始     Date;
  d_结束     Date;
  v_Xtmp     Varchar(5000);
  x_Tmp      Xmltype;
  x_Templet  Xmltype;
  n_初床位数 病人医嘱记录.Id%Type;
  n_总天数   Number;
  v_病区名称 部门表.名称%Type;

  v_月份   Varchar(30);
  d_Tmp    Date;
  n_开放数 Number;

  Cursor c_Item(病区id_In 部门表.Id%Type) Is
    Select a.病区id, a.天, Sum(a.变动) As 开放床位数
    From (Select a.病区id, a.变动, To_Char(a.日期, 'yyyy-mm-dd') As 天
           From 床位增减记录 A
           Where a.日期 Between d_开始 And d_结束 And a.病区id = 病区id_In) A
    Group By a.病区id, a.天
    Order By a.天;

  Type t_Item Is Table Of c_Item%RowType;
  r_Item t_Item;
Begin
  Select Extractvalue(Value(A), 'IN/BQID') As 病区id,
         To_Date(Extractvalue(Value(A), 'IN/KSRQ'), 'yyyy-mm-dd hh24:mi:ss') As 开始日期,
         To_Date(Extractvalue(Value(A), 'IN/JSRQ'), 'yyyy-mm-dd hh24:mi:ss') As 结束日期
  Into n_病区id, d_开始, d_结束
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  x_Templet := Xmltype('<OUTPUT><BQLIST></BQLIST></OUTPUT>');

  --单个病区
  If Nvl(n_病区id, 0) <> 0 Then
    Select 名称 Into v_病区名称 From 部门表 Where ID = n_病区id;
    Select Nvl(Sum(a.变动), 0) Into n_初床位数 From 床位增减记录 A Where a.病区id = n_病区id And a.日期 < d_开始;
    Open c_Item(n_病区id);
    Fetch c_Item Bulk Collect
      Into r_Item;
    Close c_Item;
    v_Xtmp   := '<ITEM jsonArray="True" ><BQID>' || n_病区id || '</BQID><BQMC>' || v_病区名称 || '</BQMC>';
    v_Xtmp   := v_Xtmp || '<DATALIST></DATALIST></ITEM>';
    x_Tmp    := Xmltype(v_Xtmp);
    n_开放数 := 0;
    d_Tmp    := d_开始;
    v_月份   := '-';
    --循环天数
    While d_Tmp <= d_结束 Loop
      For J In 1 .. r_Item.Count Loop
        If r_Item(J).天 = To_Char(d_Tmp, 'yyyy-mm-dd') Then
          n_初床位数 := n_初床位数 + r_Item(J).开放床位数;
        End If;
      End Loop;
    
      If Substr(To_Char(d_Tmp, 'yyyy-mm-dd'), 1, 7) <> v_月份 Then
        --1.拼接之前的
        If v_月份 <> '-' Then
          n_开放数 := Round(n_开放数 / n_总天数);
          v_Xtmp   := '<ITEM jsonArray="True" ><YF>' || Substr(v_月份, 6, 2) || '</YF><KFCR>' || n_开放数 || '</KFCR></ITEM>';
          Select Appendchildxml(x_Tmp, '/ITEM/DATALIST', Xmltype(v_Xtmp)) Into x_Tmp From Dual;
        End If;
        v_月份 := Substr(To_Char(d_Tmp, 'yyyy-mm-dd'), 1, 7);
        --2.重置
        n_总天数 := 1;
        n_开放数 := n_初床位数;
      Else
        n_总天数 := n_总天数 + 1;
        n_开放数 := n_开放数 + n_初床位数;
      End If;
      d_Tmp := d_Tmp + 1;
    End Loop;
    n_开放数 := Round(n_开放数 / n_总天数);
    v_Xtmp   := '<ITEM jsonArray="True" ><YF>' || Substr(v_月份, 6, 2) || '</YF><KFCR>' || n_开放数 || '</KFCR></ITEM>';
    Select Appendchildxml(x_Tmp, '/ITEM/DATALIST', Xmltype(v_Xtmp)) Into x_Tmp From Dual;
    Select Appendchildxml(x_Templet, '/OUTPUT/BQLIST', x_Tmp) Into x_Templet From Dual;
  Else
    --所有病区  
    For R In (Select a.Id, a.名称, a.编码
              From 部门表 A, 部门性质说明 B
              Where a.Id = b.部门id And b.工作性质 = '护理' And 服务对象 = 2
              Group By a.Id, a.名称, a.编码
              Order By a.编码) Loop
      v_病区名称 := r.名称;
      n_病区id   := r.Id;
    
      Select Nvl(Sum(a.变动), 0) Into n_初床位数 From 床位增减记录 A Where a.病区id = n_病区id And a.日期 < d_开始;
      Open c_Item(n_病区id);
      Fetch c_Item Bulk Collect
        Into r_Item;
      Close c_Item;
    
      v_Xtmp   := '<ITEM jsonArray="True" ><BQID>' || n_病区id || '</BQID><BQMC>' || v_病区名称 || '</BQMC>';
      v_Xtmp   := v_Xtmp || '<DATALIST></DATALIST></ITEM>';
      x_Tmp    := Xmltype(v_Xtmp);
      n_开放数 := 0;
      d_Tmp    := d_开始;
      v_月份   := '-';
      --循环天数
      While d_Tmp <= d_结束 Loop
        For J In 1 .. r_Item.Count Loop
          If r_Item(J).天 = To_Char(d_Tmp, 'yyyy-mm-dd') Then
            n_初床位数 := n_初床位数 + r_Item(J).开放床位数;
          End If;
        End Loop;
      
        If Substr(To_Char(d_Tmp, 'yyyy-mm-dd'), 1, 7) <> v_月份 Then
          --1.拼接之前的
          If v_月份 <> '-' Then
            n_开放数 := Round(n_开放数 / n_总天数);
            v_Xtmp   := '<ITEM jsonArray="True" ><YF>' || Substr(v_月份, 6, 2) || '</YF><KFCR>' || n_开放数 || '</KFCR></ITEM>';
            Select Appendchildxml(x_Tmp, '/ITEM/DATALIST', Xmltype(v_Xtmp)) Into x_Tmp From Dual;
          End If;
          v_月份 := Substr(To_Char(d_Tmp, 'yyyy-mm-dd'), 1, 7);
          --2.重置
          n_总天数 := 1;
          n_开放数 := n_初床位数;
        Else
          n_总天数 := n_总天数 + 1;
          n_开放数 := n_开放数 + n_初床位数;
        End If;
        d_Tmp := d_Tmp + 1;
      End Loop;
      n_开放数 := Round(n_开放数 / n_总天数);
      v_Xtmp   := '<ITEM jsonArray="True" ><YF>' || Substr(v_月份, 6, 2) || '</YF><KFCR>' || n_开放数 || '</KFCR></ITEM>';
      Select Appendchildxml(x_Tmp, '/ITEM/DATALIST', Xmltype(v_Xtmp)) Into x_Tmp From Dual;
      Select Appendchildxml(x_Templet, '/OUTPUT/BQLIST', x_Tmp) Into x_Templet From Dual;
    End Loop;
  End If;
  Xml_Out := x_Templet;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getkfcws;
/

--122937:胡俊勇,2018-03-15,护理接口出参添加的结点属性
Create Or Replace Procedure Zl_Third_Getzyrs
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --功能：住院患者人数/查询
  --入参：Xml_In
  --<IN>
  --  <BQID></BQID>    //病区ID，传空取所有病区
  --  <KSRQ></KSRQ>  //开始日期   yyyy-mm-dd
  --  <JSRQ></JSRQ>   //结束日期  yyyy-mm-dd
  --</IN>
  --出参：xml_out
  --<OUTPUT>
  --  <BQLIST>
  --    <ITEM>
  --      <BQID></BQID>  //病区ID
  --      <BQMC></BQMC>  //病区名称
  --      <DATALIST>
  --        <ITEM>
  --           <YF></YF>  //月份
  --           <QCRS></QCRS>  //期初人数，即开始时间的在院病人数
  --           <XRRS></XRRS>  //新入人数，即时间段内新入病区的人数，包括入院、转入
  --           <XCRS></XCRS>  //新出人数，即时间段内新出病区的人数，包括出院、转出、死亡
  --           <QMRS></QMRS>  //期末人数，即结束时间的在院病人数
  --        </ITEM>
  --      </DATALIST>
  --    </ITEM>
  --  </BQLIST>
  --</OUTPUT>

  n_病区id   病人医嘱记录.执行科室id%Type;
  v_病区名称 部门表.名称%Type;
  v_Xtmp     Varchar(5000); --临时XML
  x_Tmp      Xmltype;
  d_开始     Date;
  d_结束     Date;
  x_Templet  Xmltype;
  d_Tmp      Date;
  n_期初人数 Number;
  n_新入人数 Number;
  n_新出人数 Number;
  n_期末人数 Number;
  d_s        Date;
  d_e        Date;

  v_月份 Varchar(50);

  --病区指定时间点的人数
  Cursor c_当前人数
  (
    时间_In   Date,
    病区id_In 病人医嘱记录.执行科室id%Type
  ) Is
    Select Count(1) As 人数
    From 病人变动记录 A
    Where 开始时间 < 时间_In And (终止时间 Is Null Or 终止时间 > 时间_In) And Nvl(a.附加床位, 0) = 0 And 病区id = 病区id_In;

  r_当前人数 c_当前人数%RowType;

  Cursor c_入人数
  (
    时间起_In Date,
    时间止_In Date,
    病区id_In 病人医嘱记录.执行科室id%Type
  ) Is
    Select Count(1) As 人数
    From 病人变动记录 A
    Where (a.开始原因 In (2, 3, 15) Or a.开始原因 = 1 And Not Exists
           (Select 1 From 病人变动记录 B Where a.病人id = b.病人id And a.主页id = b.主页id And b.开始原因 = 2)) And a.病区id = 病区id_In And
          a.开始时间 Between 时间起_In And 时间止_In And Nvl(a.附加床位, 0) = 0;
  r_入人数 c_入人数%RowType;

Begin
  Select Extractvalue(Value(A), 'IN/BQID') As 病区id,
         To_Date(Extractvalue(Value(A), 'IN/KSRQ'), 'yyyy-mm-dd hh24:mi:ss') As 开始日期,
         To_Date(Extractvalue(Value(A), 'IN/JSRQ'), 'yyyy-mm-dd hh24:mi:ss') As 结束日期
  Into n_病区id, d_开始, d_结束
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  x_Templet := Xmltype('<OUTPUT><BQLIST></BQLIST></OUTPUT>');

  If Nvl(n_病区id, 0) <> 0 Then
    Select 名称 Into v_病区名称 From 部门表 Where ID = n_病区id;
  
    v_Xtmp := '<ITEM jsonArray="True" ><BQID>' || n_病区id || '</BQID><BQMC>' || v_病区名称 || '</BQMC>';
    v_Xtmp := v_Xtmp || '<DATALIST></DATALIST></ITEM>';
    x_Tmp  := Xmltype(v_Xtmp);
  
    d_Tmp  := d_开始;
    v_月份 := '-';
    d_s    := d_开始;
  
    --循环天数取出每个月份
    While d_Tmp <= d_结束 Loop
      If Substr(To_Char(d_Tmp, 'yyyy-mm-dd'), 1, 7) <> v_月份 Then
        If v_月份 <> '-' Then
        
          d_e := Trunc(d_Tmp) - 1 / 24 / 60;
        
          Open c_当前人数(d_s, n_病区id);
          Fetch c_当前人数
            Into r_当前人数;
          If c_当前人数%RowCount = 0 Then
            n_期初人数 := 0;
          Else
            n_期初人数 := r_当前人数.人数;
          End If;
          Close c_当前人数;
        
          Open c_当前人数(d_e, n_病区id);
          Fetch c_当前人数
            Into r_当前人数;
          If c_当前人数%RowCount = 0 Then
            n_期末人数 := 0;
          Else
            n_期末人数 := r_当前人数.人数;
          End If;
          Close c_当前人数;
        
          Open c_入人数(d_s, d_e, n_病区id);
          Fetch c_入人数
            Into r_入人数;
          If c_入人数%RowCount = 0 Then
            n_新入人数 := 0;
          Else
            n_新入人数 := r_入人数.人数;
          End If;
          Close c_入人数;
        
          n_新出人数 := n_新入人数 + n_期初人数 - n_期末人数;
        
          v_Xtmp := '<ITEM jsonArray="True" ><YF>' || Substr(v_月份, 6, 2) || '</YF><QCRS>' || n_期初人数 || '</QCRS><XRRS>' || n_新入人数 ||
                    '</XRRS><XCRS>' || n_新出人数 || '</XCRS><QMRS>' || n_期末人数 || '</QMRS></ITEM>';
          Select Appendchildxml(x_Tmp, '/ITEM/DATALIST', Xmltype(v_Xtmp)) Into x_Tmp From Dual;
        
        End If;
        v_月份 := Substr(To_Char(d_Tmp, 'yyyy-mm-dd'), 1, 7);
        d_s    := d_Tmp;
      
      End If;
      d_Tmp := d_Tmp + 1;
    End Loop;
  
    d_e := d_结束;
  
    Open c_当前人数(d_s, n_病区id);
    Fetch c_当前人数
      Into r_当前人数;
    If c_当前人数%RowCount = 0 Then
      n_期初人数 := 0;
    Else
      n_期初人数 := r_当前人数.人数;
    End If;
    Close c_当前人数;
  
    Open c_当前人数(d_e, n_病区id);
    Fetch c_当前人数
      Into r_当前人数;
    If c_当前人数%RowCount = 0 Then
      n_期末人数 := 0;
    Else
      n_期末人数 := r_当前人数.人数;
    End If;
    Close c_当前人数;
  
    Open c_入人数(d_s, d_e, n_病区id);
    Fetch c_入人数
      Into r_入人数;
    If c_入人数%RowCount = 0 Then
      n_新入人数 := 0;
    Else
      n_新入人数 := r_入人数.人数;
    End If;
    Close c_入人数;
  
    n_新出人数 := n_新入人数 + n_期初人数 - n_期末人数;
  
    v_Xtmp := '<ITEM jsonArray="True" ><YF>' || Substr(v_月份, 6, 2) || '</YF><QCRS>' || n_期初人数 || '</QCRS><XRRS>' || n_新入人数 ||
              '</XRRS><XCRS>' || n_新出人数 || '</XCRS><QMRS>' || n_期末人数 || '</QMRS></ITEM>';
    Select Appendchildxml(x_Tmp, '/ITEM/DATALIST', Xmltype(v_Xtmp)) Into x_Tmp From Dual;
    If x_Tmp Is Not Null Then
      Select Appendchildxml(x_Templet, '/OUTPUT/BQLIST', x_Tmp) Into x_Templet From Dual;
    End If;
  Else
    --所有病区
    For R In (Select a.Id, a.名称, a.编码
              From 部门表 A, 部门性质说明 B
              Where a.Id = b.部门id And b.工作性质 = '护理' And 服务对象 = 2
              Group By a.Id, a.名称, a.编码
              Order By a.编码) Loop
      v_病区名称 := r.名称;
      n_病区id   := r.Id;
    
      v_Xtmp := '<ITEM jsonArray="True" ><BQID>' || n_病区id || '</BQID><BQMC>' || v_病区名称 || '</BQMC>';
      v_Xtmp := v_Xtmp || '<DATALIST></DATALIST></ITEM>';
      x_Tmp  := Xmltype(v_Xtmp);
    
      d_Tmp  := d_开始;
      v_月份 := '-';
      d_s    := d_开始;
    
      --循环天数取出每个月份
      While d_Tmp <= d_结束 Loop
        If Substr(To_Char(d_Tmp, 'yyyy-mm-dd'), 1, 7) <> v_月份 Then
          If v_月份 <> '-' Then
          
            d_e := Trunc(d_Tmp) - 1 / 24 / 60;
          
            Open c_当前人数(d_s, n_病区id);
            Fetch c_当前人数
              Into r_当前人数;
            If c_当前人数%RowCount = 0 Then
              n_期初人数 := 0;
            Else
              n_期初人数 := r_当前人数.人数;
            End If;
            Close c_当前人数;
          
            Open c_当前人数(d_e, n_病区id);
            Fetch c_当前人数
              Into r_当前人数;
            If c_当前人数%RowCount = 0 Then
              n_期末人数 := 0;
            Else
              n_期末人数 := r_当前人数.人数;
            End If;
            Close c_当前人数;
          
            Open c_入人数(d_s, d_e, n_病区id);
            Fetch c_入人数
              Into r_入人数;
            If c_入人数%RowCount = 0 Then
              n_新入人数 := 0;
            Else
              n_新入人数 := r_入人数.人数;
            End If;
            Close c_入人数;
          
            n_新出人数 := n_新入人数 + n_期初人数 - n_期末人数;
          
            v_Xtmp := '<ITEM jsonArray="True" ><YF>' || Substr(v_月份, 6, 2) || '</YF><QCRS>' || n_期初人数 || '</QCRS><XRRS>' || n_新入人数 ||
                      '</XRRS><XCRS>' || n_新出人数 || '</XCRS><QMRS>' || n_期末人数 || '</QMRS></ITEM>';
            Select Appendchildxml(x_Tmp, '/ITEM/DATALIST', Xmltype(v_Xtmp)) Into x_Tmp From Dual;
          
          End If;
          v_月份 := Substr(To_Char(d_Tmp, 'yyyy-mm-dd'), 1, 7);
          d_s    := d_Tmp;
        
        End If;
        d_Tmp := d_Tmp + 1;
      End Loop;
    
      d_e := d_结束;
    
      Open c_当前人数(d_s, n_病区id);
      Fetch c_当前人数
        Into r_当前人数;
      If c_当前人数%RowCount = 0 Then
        n_期初人数 := 0;
      Else
        n_期初人数 := r_当前人数.人数;
      End If;
      Close c_当前人数;
    
      Open c_当前人数(d_e, n_病区id);
      Fetch c_当前人数
        Into r_当前人数;
      If c_当前人数%RowCount = 0 Then
        n_期末人数 := 0;
      Else
        n_期末人数 := r_当前人数.人数;
      End If;
      Close c_当前人数;
    
      Open c_入人数(d_s, d_e, n_病区id);
      Fetch c_入人数
        Into r_入人数;
      If c_入人数%RowCount = 0 Then
        n_新入人数 := 0;
      Else
        n_新入人数 := r_入人数.人数;
      End If;
      Close c_入人数;
    
      n_新出人数 := n_新入人数 + n_期初人数 - n_期末人数;
    
      v_Xtmp := '<ITEM jsonArray="True" ><YF>' || Substr(v_月份, 6, 2) || '</YF><QCRS>' || n_期初人数 || '</QCRS><XRRS>' || n_新入人数 ||
                '</XRRS><XCRS>' || n_新出人数 || '</XCRS><QMRS>' || n_期末人数 || '</QMRS></ITEM>';
      Select Appendchildxml(x_Tmp, '/ITEM/DATALIST', Xmltype(v_Xtmp)) Into x_Tmp From Dual;
      If x_Tmp Is Not Null Then
        Select Appendchildxml(x_Templet, '/OUTPUT/BQLIST', x_Tmp) Into x_Templet From Dual;
      End If;
    End Loop;
  End If;
  Xml_Out := x_Templet;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getzyrs;
/

--122937:胡俊勇,2018-03-15,护理接口出参添加的结点属性
CREATE OR REPLACE Procedure Zl_Third_Getzycws
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --功能：实际占用床日数/查询
  --实际占用总床日数：指每日夜晚12点实际占用病床数(即每日夜晚12点住院人数)总和。
  --                   包括实际占用的临时加床在内。病人入院后于当晚12点前死亡或因故出院的病人, 作为实际占用床位1天进行统计
  --入参：Xml_In
  --<IN>
  --    <BQID></BQID>    //病区ID，传空取所有病区
  --    <KSRQ></KSRQ>    //开始日期   yyyy-mm-dd
  --    <JSRQ></JSRQ>    //结束日期  yyyy-mm-dd
  --</IN>

  --出参：xml_out
  --<OUTPUT>
  --  <BQLIST>
  --    <ITEM>
  --      <BQID></BQID>  //病区ID
  --      <BQMC></BQMC>  //病区名称
  --      <DATALIST>
  --        <ITEM>
  --           <YF></YF>  //月份
  --           <ZYCR></ZYCR>  //实际占用床日数
  --        </ITEM>
  --      </DATALIST>
  --    </ITEM>
  --  </BQLIST>
  --</OUTPUT>
  n_病区id     部门表.Id%Type;
  v_病区名称   部门表.名称%Type;
  d_开始       Date;
  d_结束       Date;
  v_Xtmp       Varchar(5000); --临时XML
  x_Tmp        Xmltype;
  x_Templet    Xmltype;
  n_床日数     Number; --病人住院天数之和
  n_病人床日数 Number;
  n_病区床日数 Number(18);
  v_月份       Varchar2(50);
  d_Tmp        Date;
  v_Pre病人    Varchar2(100);
  n_Index      Number;
  d_s          Date;
  d_e          Date;

  d_天数起    Date;
  d_天数止    Date;
  d_Pre天数止 Date;

  Cursor c_Item
  (
    时间起_In Date,
    时间止_In Date,
    病区id_In 病人医嘱记录.执行科室id%Type
  ) Is
    Select Case
             When 终止原因 = 1 And 开始原因 In (1, 2, 3, 15) And 病区id = 病区id_In Then
              '转入加转出'
             When 终止原因 = 1 Or 开始原因 In (3, 15) And 病区id <> 病区id_In Then
              '转出'
             Else
              '转入'
           End As 类型, 病人id, 主页id, Trunc(开始时间) AS 开始时间, Trunc(终止时间) AS 终止时间, 病区id, 开始原因, 终止原因
    From (Select a.病人id, a.主页id,
                  Case
                    When Trunc(a.开始时间) < 时间起_In Then
                     时间起_In
                    Else
                     a.开始时间
                  End As 开始时间,
                  Case
                    When Trunc(a.终止时间) > 时间止_In Then
                     时间止_In
                    Else
                     a.终止时间
                  End As 终止时间, a.病区id, a.开始原因, a.终止原因
           From 病人变动记录 A
           Where a.开始时间 Between 时间起_In And 时间止_In And Exists
            (Select 1 From 病人变动记录 B Where a.病人id = b.病人id And a.主页id = b.主页id And b.病区id = 病区id_In) And
                 ((((a.开始原因 = 2 Or a.开始原因 = 1 And Not Exists
                  (Select 1 From 病人变动记录 C Where a.病人id = c.病人id And a.主页id = c.主页id And c.开始原因 = 2)) And a.病区id = 病区id_In Or
                 a.开始原因 In (3, 15))) Or a.终止原因 = 1) And Nvl(a.附加床位, 0) = 0
           Union All
           Select a.病人id, a.主页id,
                  Case
                    When Trunc(a.开始时间) < 时间起_In Then
                     时间起_In
                    Else
                     a.开始时间
                  End As 开始时间,
                  Case
                    When Trunc(a.终止时间) > 时间止_In Then
                     时间止_In
                    Else
                     a.终止时间
                  End As 终止时间, a.病区id, a.开始原因, a.终止原因
           From 病人变动记录 A
           Where a.开始时间 < 时间起_In And a.病区id = 病区id_In And
                 ((a.开始原因 = 2 Or a.开始原因 = 1 And Not Exists
                  (Select 1 From 病人变动记录 C Where a.病人id = c.病人id And a.主页id = c.主页id And c.开始原因 = 2)) And a.病区id = 病区id_In) And
                 Nvl(a.附加床位, 0) = 0 And Not Exists
            (Select 1
                  From 病人变动记录 B
                  Where a.病人id = b.病人id And a.主页id = b.主页id And b.开始时间 < 时间止_In And
                        (b.开始原因 In (3, 15) And b.病区id <> 病区id_In Or b.终止原因 = 1 And 病区id = 病区id_In)))
    Order By 病人id, 主页id, 开始时间, 终止时间;

  Type t_Item Is Table Of c_Item%RowType;
  r_Item t_Item;
Begin

  Select Extractvalue(Value(A), 'IN/BQID') As 病区id,
         To_Date(Extractvalue(Value(A), 'IN/KSRQ'), 'yyyy-mm-dd hh24:mi:ss') As 开始日期,
         To_Date(Extractvalue(Value(A), 'IN/JSRQ'), 'yyyy-mm-dd hh24:mi:ss') As 结束日期
  Into n_病区id, d_开始, d_结束
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  x_Templet := Xmltype('<OUTPUT><BQLIST></BQLIST></OUTPUT>');

  If Nvl(n_病区id, 0) <> 0 Then
    Select 名称 Into v_病区名称 From 部门表 Where ID = n_病区id;
    v_Xtmp   := '<ITEM jsonArray="True" ><BQID>' || n_病区id || '</BQID><BQMC>' || v_病区名称 || '</BQMC>';
    v_Xtmp   := v_Xtmp || '<DATALIST></DATALIST></ITEM>';
    x_Tmp    := Xmltype(v_Xtmp);
    d_Tmp    := d_开始;
    v_月份   := '-';
    d_s      := d_开始;
    n_床日数 := 0;

    --循环天数取出每个月份
    While d_Tmp <= d_结束 Loop
      If Substr(To_Char(d_Tmp, 'yyyy-mm-dd'), 1, 7) <> v_月份 Then
        If v_月份 <> '-' Then
          d_e       := Trunc(d_Tmp) - 1 / 24 / 60;
          n_床日数  := 0;
          v_Pre病人 := '-';
          Open c_Item(d_s, d_e, n_病区id);
          Fetch c_Item Bulk Collect
            Into r_Item;
          Close c_Item;
          n_病区床日数 := 0;
          For I In 1 .. r_Item.Count Loop
            If r_Item(I).病人id || '_' || r_Item(I).主页id <> v_Pre病人 Or v_Pre病人 = '-' Then
              --新病人开始
              If r_Item(I).类型 = '转入加转出' Then
                d_天数起 := r_Item(I).开始时间;
                d_天数止 := r_Item(I).终止时间;
                If (d_天数止 - d_天数起)=0 Then
                  n_病人床日数 :=1;
                Else
                  n_病人床日数 :=(d_天数止 - d_天数起);
                End If;
                n_病区床日数 := n_病区床日数 + n_病人床日数 ;
              Elsif r_Item(I).类型 = '转入' Then
                If I >=r_Item.Count Then
                  d_天数起 := r_Item(I).开始时间;
                  d_天数止 := Trunc(d_e);
                  If (d_天数止 - d_天数起)=0 Then
                    n_病人床日数 :=1;
                  Else
                    n_病人床日数 :=(d_天数止 - d_天数起);
                  End If;
                  n_病区床日数 := n_病区床日数 + n_病人床日数 ;
                Else
                  If r_Item(I+1).类型 = '转出' And r_Item(I+1).病人id || '_' || r_Item(I+1).主页id = r_Item(I).病人id || '_' || r_Item(I).主页id Then
                     d_天数起 := r_Item(I).开始时间;
                     d_天数止 := r_Item(I+1).开始时间;
                     If (d_天数止 - d_天数起)=0 Then
                       n_病人床日数 :=1;
                     Else
                       n_病人床日数 :=(d_天数止 - d_天数起);
                     End If;
                     n_病区床日数 := n_病区床日数 + n_病人床日数 ;
                  Else
                     d_天数起 := r_Item(I).开始时间;
                     d_天数止 := Trunc(d_e);
                     If (d_天数止 - d_天数起)=0 Then
                       n_病人床日数 :=1;
                     Else
                       n_病人床日数 :=(d_天数止 - d_天数起);
                     End If;
                     n_病区床日数 := n_病区床日数 + n_病人床日数 ;
                  End If;
                End If;
              Elsif r_Item(I).类型 = '转出' Then
                d_天数起 := Trunc(d_s);
                If NVL(r_Item(I).终止原因,0) = 1 then
                   d_天数止 := r_Item(I).终止时间;
                Else
                  d_天数止 := r_Item(I).开始时间;
                End If;
                If (d_天数止 - d_天数起)=0 Then
                  n_病人床日数 :=1;
                Else
                  n_病人床日数 :=(d_天数止 - d_天数起);
                End If;
                n_病区床日数 := n_病区床日数 + n_病人床日数 ;
              End If;
            Else
              If r_Item(I).类型 = '转入加转出' Then
                --如果上一个起始时间和这次的起始时间相同，则减一，例如：一天内在同一个科室转了多次；
                if d_天数起 = r_Item(I).开始时间 Then
                  n_病区床日数 := n_病区床日数 - 1;
                End If;
                d_天数起 := r_Item(I).开始时间;
                d_天数止 := r_Item(I).终止时间;
                If (d_天数止 - d_天数起)=0 Then
                  n_病人床日数 :=1;
                Else
                  n_病人床日数 :=(d_天数止 - d_天数起);
                End If;
                n_病区床日数 := n_病区床日数 + n_病人床日数 ;
              Elsif r_Item(I).类型 = '转入' Then
                --如果上一个起始时间和这次的起始时间相同，则减一，例如：一天内在同一个科室转了多次；
                if d_天数起 = r_Item(I).开始时间 Then
                  n_病区床日数 := n_病区床日数 - 1;
                End If;
                If I >=r_Item.Count Then
                  d_天数起 := r_Item(I).开始时间;
                  d_天数止 := Trunc(d_e);
                  If (d_天数止 - d_天数起)=0 Then
                    n_病人床日数 :=1;
                  Else
                    n_病人床日数 :=(d_天数止 - d_天数起);
                  End If;
                  n_病区床日数 := n_病区床日数 + n_病人床日数 ;
                Else
                  If r_Item(I+1).类型 = '转出' And r_Item(I+1).病人id || '_' || r_Item(I+1).主页id = r_Item(I).病人id || '_' || r_Item(I).主页id Then
                     d_天数起 := r_Item(I).开始时间;
                     d_天数止 := r_Item(I+1).开始时间;
                     If (d_天数止 - d_天数起)=0 Then
                       n_病人床日数 :=1;
                     Else
                       n_病人床日数 :=(d_天数止 - d_天数起);
                     End If;
                     n_病区床日数 := n_病区床日数 + n_病人床日数 ;
                  Else
                     d_天数起 := r_Item(I).开始时间;
                     d_天数止 := Trunc(d_e);
                     If (d_天数止 - d_天数起)=0 Then
                       n_病人床日数 :=1;
                     Else
                       n_病人床日数 :=(d_天数止 - d_天数起);
                     End If;
                     n_病区床日数 := n_病区床日数 + n_病人床日数 ;
                  End If;
                End If;
              End If;
            End If;
            v_Pre病人 := r_Item(I).病人id || '_' || r_Item(I).主页id;
          End Loop;
          v_Xtmp := '<ITEM jsonArray="True" ><YF>' || Substr(v_月份, 6, 2) || '</YF><ZYCR>' || n_病区床日数 || '</ZYCR></ITEM>';
          Select Appendchildxml(x_Tmp, '/ITEM/DATALIST', Xmltype(v_Xtmp)) Into x_Tmp From Dual;
        End If;
        v_月份 := Substr(To_Char(d_Tmp, 'yyyy-mm-dd'), 1, 7);
        d_s    := d_Tmp;
      End If;
      d_Tmp := d_Tmp + 1;
    End Loop;

    d_e       := d_结束;
    n_床日数  := 0;
    v_Pre病人 := '-';
    Open c_Item(d_s, d_e, n_病区id);
    Fetch c_Item Bulk Collect
      Into r_Item;
    Close c_Item;
    n_病区床日数 := 0;
    For I In 1 .. r_Item.Count Loop
      If r_Item(I).病人id || '_' || r_Item(I).主页id <> v_Pre病人 Or v_Pre病人 = '-' Then
        --新病人开始
        If r_Item(I).类型 = '转入加转出' Then
          d_天数起 := r_Item(I).开始时间;
          d_天数止 := r_Item(I).终止时间;
          If (d_天数止 - d_天数起)=0 Then
            n_病人床日数 :=1;
          Else
            n_病人床日数 :=(d_天数止 - d_天数起);
          End If;
          n_病区床日数 := n_病区床日数 + n_病人床日数 ;
        Elsif r_Item(I).类型 = '转入' Then
          If I >=r_Item.Count Then
            d_天数起 := r_Item(I).开始时间;
            d_天数止 := Trunc(d_e);
            If (d_天数止 - d_天数起)=0 Then
              n_病人床日数 :=1;
            Else
              n_病人床日数 :=(d_天数止 - d_天数起);
            End If;
            n_病区床日数 := n_病区床日数 + n_病人床日数 ;
          Else
            If r_Item(I+1).类型 = '转出' And r_Item(I+1).病人id || '_' || r_Item(I+1).主页id = r_Item(I).病人id || '_' || r_Item(I).主页id Then
               d_天数起 := r_Item(I).开始时间;
               d_天数止 := r_Item(I+1).开始时间;
               If (d_天数止 - d_天数起)=0 Then
                 n_病人床日数 :=1;
               Else
                 n_病人床日数 :=(d_天数止 - d_天数起);
               End If;
               n_病区床日数 := n_病区床日数 + n_病人床日数 ;
            Else
               d_天数起 := r_Item(I).开始时间;
               d_天数止 := Trunc(d_e);
               If (d_天数止 - d_天数起)=0 Then
                 n_病人床日数 :=1;
               Else
                 n_病人床日数 :=(d_天数止 - d_天数起);
               End If;
               n_病区床日数 := n_病区床日数 + n_病人床日数 ;
            End If;
          End If;
        Elsif r_Item(I).类型 = '转出' Then
          d_天数起 := Trunc(d_s);
          If NVL(r_Item(I).终止原因,0) = 1 then
             d_天数止 := r_Item(I).终止时间;
          Else
            d_天数止 := r_Item(I).开始时间;
          End If;
          If (d_天数止 - d_天数起)=0 Then
            n_病人床日数 :=1;
          Else
            n_病人床日数 :=(d_天数止 - d_天数起);
          End If;
          n_病区床日数 := n_病区床日数 + n_病人床日数 ;
        End If;
      Else
        If r_Item(I).类型 = '转入加转出' Then
          --如果上一个起始时间和这次的起始时间相同，则减一，例如：一天内在同一个科室转了多次；
          if d_天数起 = r_Item(I).开始时间 Then
            n_病区床日数 := n_病区床日数 - 1;
          End If;
          d_天数起 := r_Item(I).开始时间;
          d_天数止 := r_Item(I).终止时间;
          If (d_天数止 - d_天数起)=0 Then
            n_病人床日数 :=1;
          Else
            n_病人床日数 :=(d_天数止 - d_天数起);
          End If;
          n_病区床日数 := n_病区床日数 + n_病人床日数 ;
        Elsif r_Item(I).类型 = '转入' Then
          --如果上一个起始时间和这次的起始时间相同，则减一，例如：一天内在同一个科室转了多次；
          if d_天数起 = r_Item(I).开始时间 Then
            n_病区床日数 := n_病区床日数 - 1;
          End If;
          If I >=r_Item.Count Then
            d_天数起 := r_Item(I).开始时间;
            d_天数止 := Trunc(d_e);
            If (d_天数止 - d_天数起)=0 Then
              n_病人床日数 :=1;
            Else
              n_病人床日数 :=(d_天数止 - d_天数起);
            End If;
            n_病区床日数 := n_病区床日数 + n_病人床日数 ;
          Else
            If r_Item(I+1).类型 = '转出' And r_Item(I+1).病人id || '_' || r_Item(I+1).主页id = r_Item(I).病人id || '_' || r_Item(I).主页id Then
               d_天数起 := r_Item(I).开始时间;
               d_天数止 := r_Item(I+1).开始时间;
               If (d_天数止 - d_天数起)=0 Then
                 n_病人床日数 :=1;
               Else
                 n_病人床日数 :=(d_天数止 - d_天数起);
               End If;
               n_病区床日数 := n_病区床日数 + n_病人床日数 ;
            Else
               d_天数起 := r_Item(I).开始时间;
               d_天数止 := Trunc(d_e);
               If (d_天数止 - d_天数起)=0 Then
                 n_病人床日数 :=1;
               Else
                 n_病人床日数 :=(d_天数止 - d_天数起);
               End If;
               n_病区床日数 := n_病区床日数 + n_病人床日数 ;
            End If;
          End If;
        End If;
      End If;
      v_Pre病人 := r_Item(I).病人id || '_' || r_Item(I).主页id;
    End Loop;
    v_Xtmp := '<ITEM jsonArray="True" ><YF>' || Substr(v_月份, 6, 2) || '</YF><ZYCR>' || n_病区床日数 || '</ZYCR></ITEM>';
    Select Appendchildxml(x_Tmp, '/ITEM/DATALIST', Xmltype(v_Xtmp)) Into x_Tmp From Dual;
    If x_Tmp Is Not Null Then
      Select Appendchildxml(x_Templet, '/OUTPUT/BQLIST', x_Tmp) Into x_Templet From Dual;
    End If;
  Else
    --所有病区
    For R In (Select a.Id, a.名称, a.编码
              From 部门表 A, 部门性质说明 B
              Where a.Id = b.部门id And b.工作性质 = '护理' And 服务对象 = 2
              Group By a.Id, a.名称, a.编码
              Order By a.编码) Loop
      v_病区名称 := r.名称;
      n_病区id   := r.Id;

      v_Xtmp   := '<ITEM jsonArray="True" ><BQID>' || n_病区id || '</BQID><BQMC>' || v_病区名称 || '</BQMC>';
      v_Xtmp   := v_Xtmp || '<DATALIST></DATALIST></ITEM>';
      x_Tmp    := Xmltype(v_Xtmp);
      d_Tmp    := d_开始;
      v_月份   := '-';
      d_s      := d_开始;
      n_床日数 := 0;

      --循环天数取出每个月份
      While d_Tmp <= d_结束 Loop
        If Substr(To_Char(d_Tmp, 'yyyy-mm-dd'), 1, 7) <> v_月份 Then
          If v_月份 <> '-' Then
            d_e       := Trunc(d_Tmp) - 1 / 24 / 60;
            n_床日数  := 0;
            v_Pre病人 := '-';
            Open c_Item(d_s, d_e, n_病区id);
            Fetch c_Item Bulk Collect
              Into r_Item;
            Close c_Item;
            n_病区床日数 := 0;
            For I In 1 .. r_Item.Count Loop
              If r_Item(I).病人id || '_' || r_Item(I).主页id <> v_Pre病人 Or v_Pre病人 = '-' Then
                --新病人开始
                If r_Item(I).类型 = '转入加转出' Then
                  d_天数起 := r_Item(I).开始时间;
                  d_天数止 := r_Item(I).终止时间;
                  If (d_天数止 - d_天数起)=0 Then
                    n_病人床日数 :=1;
                  Else
                    n_病人床日数 :=(d_天数止 - d_天数起);
                  End If;
                  n_病区床日数 := n_病区床日数 + n_病人床日数 ;
                Elsif r_Item(I).类型 = '转入' Then
                  If I >=r_Item.Count Then
                    d_天数起 := r_Item(I).开始时间;
                    d_天数止 := Trunc(d_e);
                    If (d_天数止 - d_天数起)=0 Then
                      n_病人床日数 :=1;
                    Else
                      n_病人床日数 :=(d_天数止 - d_天数起);
                    End If;
                    n_病区床日数 := n_病区床日数 + n_病人床日数 ;
                  Else
                    If r_Item(I+1).类型 = '转出' And r_Item(I+1).病人id || '_' || r_Item(I+1).主页id = r_Item(I).病人id || '_' || r_Item(I).主页id Then
                       d_天数起 := r_Item(I).开始时间;
                       d_天数止 := r_Item(I+1).开始时间;
                       If (d_天数止 - d_天数起)=0 Then
                         n_病人床日数 :=1;
                       Else
                         n_病人床日数 :=(d_天数止 - d_天数起);
                       End If;
                       n_病区床日数 := n_病区床日数 + n_病人床日数 ;
                    Else
                       d_天数起 := r_Item(I).开始时间;
                       d_天数止 := Trunc(d_e);
                       If (d_天数止 - d_天数起)=0 Then
                         n_病人床日数 :=1;
                       Else
                         n_病人床日数 :=(d_天数止 - d_天数起);
                       End If;
                       n_病区床日数 := n_病区床日数 + n_病人床日数 ;
                    End If;
                  End If;
                Elsif r_Item(I).类型 = '转出' Then
                  d_天数起 := Trunc(d_s);
                  If NVL(r_Item(I).终止原因,0) = 1 then
                     d_天数止 := r_Item(I).终止时间;
                  Else
                    d_天数止 := r_Item(I).开始时间;
                  End If;
                  If (d_天数止 - d_天数起)=0 Then
                    n_病人床日数 :=1;
                  Else
                    n_病人床日数 :=(d_天数止 - d_天数起);
                  End If;
                  n_病区床日数 := n_病区床日数 + n_病人床日数 ;
                End If;
              Else
                If r_Item(I).类型 = '转入加转出' Then
                  --如果上一个起始时间和这次的起始时间相同，则减一，例如：一天内在同一个科室转了多次；
                  if d_天数起 = r_Item(I).开始时间 Then
                    n_病区床日数 := n_病区床日数 - 1;
                  End If;
                  d_天数起 := r_Item(I).开始时间;
                  d_天数止 := r_Item(I).终止时间;
                  If (d_天数止 - d_天数起)=0 Then
                    n_病人床日数 :=1;
                  Else
                    n_病人床日数 :=(d_天数止 - d_天数起);
                  End If;
                  n_病区床日数 := n_病区床日数 + n_病人床日数 ;
                Elsif r_Item(I).类型 = '转入' Then
                  --如果上一个起始时间和这次的起始时间相同，则减一，例如：一天内在同一个科室转了多次；
                  if d_天数起 = r_Item(I).开始时间 Then
                    n_病区床日数 := n_病区床日数 - 1;
                  End If;
                  If I >=r_Item.Count Then
                    d_天数起 := r_Item(I).开始时间;
                    d_天数止 := Trunc(d_e);
                    If (d_天数止 - d_天数起)=0 Then
                      n_病人床日数 :=1;
                    Else
                      n_病人床日数 :=(d_天数止 - d_天数起);
                    End If;
                    n_病区床日数 := n_病区床日数 + n_病人床日数 ;
                  Else
                    If r_Item(I+1).类型 = '转出' And r_Item(I+1).病人id || '_' || r_Item(I+1).主页id = r_Item(I).病人id || '_' || r_Item(I).主页id Then
                       d_天数起 := r_Item(I).开始时间;
                       d_天数止 := r_Item(I+1).开始时间;
                       If (d_天数止 - d_天数起)=0 Then
                         n_病人床日数 :=1;
                       Else
                         n_病人床日数 :=(d_天数止 - d_天数起);
                       End If;
                       n_病区床日数 := n_病区床日数 + n_病人床日数 ;
                    Else
                       d_天数起 := r_Item(I).开始时间;
                       d_天数止 := Trunc(d_e);
                       If (d_天数止 - d_天数起)=0 Then
                         n_病人床日数 :=1;
                       Else
                         n_病人床日数 :=(d_天数止 - d_天数起);
                       End If;
                       n_病区床日数 := n_病区床日数 + n_病人床日数 ;
                    End If;
                  End If;
                End If;
              End If;
              v_Pre病人 := r_Item(I).病人id || '_' || r_Item(I).主页id;
            End Loop;
            v_Xtmp := '<ITEM jsonArray="True" ><YF>' || Substr(v_月份, 6, 2) || '</YF><ZYCR>' || n_病区床日数 || '</ZYCR></ITEM>';
            Select Appendchildxml(x_Tmp, '/ITEM/DATALIST', Xmltype(v_Xtmp)) Into x_Tmp From Dual;
          End If;
          v_月份 := Substr(To_Char(d_Tmp, 'yyyy-mm-dd'), 1, 7);
          d_s    := d_Tmp;
        End If;
        d_Tmp := d_Tmp + 1;
      End Loop;

      d_e       := d_结束;
      n_床日数  := 0;
      v_Pre病人 := '-';
      Open c_Item(d_s, d_e, n_病区id);
      Fetch c_Item Bulk Collect
        Into r_Item;
      Close c_Item;
      n_病区床日数 := 0;
      For I In 1 .. r_Item.Count Loop
        If r_Item(I).病人id || '_' || r_Item(I).主页id <> v_Pre病人 Or v_Pre病人 = '-' Then
          --新病人开始
          If r_Item(I).类型 = '转入加转出' Then
            d_天数起 := r_Item(I).开始时间;
            d_天数止 := r_Item(I).终止时间;
            If (d_天数止 - d_天数起)=0 Then
              n_病人床日数 :=1;
            Else
              n_病人床日数 :=(d_天数止 - d_天数起);
            End If;
            n_病区床日数 := n_病区床日数 + n_病人床日数 ;
          Elsif r_Item(I).类型 = '转入' Then
            If I >=r_Item.Count Then
              d_天数起 := r_Item(I).开始时间;
              d_天数止 := Trunc(d_e);
              If (d_天数止 - d_天数起)=0 Then
                n_病人床日数 :=1;
              Else
                n_病人床日数 :=(d_天数止 - d_天数起);
              End If;
              n_病区床日数 := n_病区床日数 + n_病人床日数 ;
            Else
              If r_Item(I+1).类型 = '转出' And r_Item(I+1).病人id || '_' || r_Item(I+1).主页id = r_Item(I).病人id || '_' || r_Item(I).主页id Then
                 d_天数起 := r_Item(I).开始时间;
                 d_天数止 := r_Item(I+1).开始时间;
                 If (d_天数止 - d_天数起)=0 Then
                   n_病人床日数 :=1;
                 Else
                   n_病人床日数 :=(d_天数止 - d_天数起);
                 End If;
                 n_病区床日数 := n_病区床日数 + n_病人床日数 ;
              Else
                 d_天数起 := r_Item(I).开始时间;
                 d_天数止 := Trunc(d_e);
                 If (d_天数止 - d_天数起)=0 Then
                   n_病人床日数 :=1;
                 Else
                   n_病人床日数 :=(d_天数止 - d_天数起);
                 End If;
                 n_病区床日数 := n_病区床日数 + n_病人床日数 ;
              End If;
            End If;
          Elsif r_Item(I).类型 = '转出' Then
            d_天数起 := Trunc(d_s);
            If NVL(r_Item(I).终止原因,0) = 1 then
               d_天数止 := r_Item(I).终止时间;
            Else
              d_天数止 := r_Item(I).开始时间;
            End If;
            If (d_天数止 - d_天数起)=0 Then
              n_病人床日数 :=1;
            Else
              n_病人床日数 :=(d_天数止 - d_天数起);
            End If;
            n_病区床日数 := n_病区床日数 + n_病人床日数 ;
          End If;
        Else
          If r_Item(I).类型 = '转入加转出' Then
            --如果上一个起始时间和这次的起始时间相同，则减一，例如：一天内在同一个科室转了多次；
            if d_天数起 = r_Item(I).开始时间 Then
              n_病区床日数 := n_病区床日数 - 1;
            End If;
            d_天数起 := r_Item(I).开始时间;
            d_天数止 := r_Item(I).终止时间;
            If (d_天数止 - d_天数起)=0 Then
              n_病人床日数 :=1;
            Else
              n_病人床日数 :=(d_天数止 - d_天数起);
            End If;
            n_病区床日数 := n_病区床日数 + n_病人床日数 ;
          Elsif r_Item(I).类型 = '转入' Then
            --如果上一个起始时间和这次的起始时间相同，则减一，例如：一天内在同一个科室转了多次；
            if d_天数起 = r_Item(I).开始时间 Then
              n_病区床日数 := n_病区床日数 - 1;
            End If;
            If I >=r_Item.Count Then
              d_天数起 := r_Item(I).开始时间;
              d_天数止 := Trunc(d_e);
              If (d_天数止 - d_天数起)=0 Then
                n_病人床日数 :=1;
              Else
                n_病人床日数 :=(d_天数止 - d_天数起);
              End If;
              n_病区床日数 := n_病区床日数 + n_病人床日数 ;
            Else
              If r_Item(I+1).类型 = '转出' And r_Item(I+1).病人id || '_' || r_Item(I+1).主页id = r_Item(I).病人id || '_' || r_Item(I).主页id Then
                 d_天数起 := r_Item(I).开始时间;
                 d_天数止 := r_Item(I+1).开始时间;
                 If (d_天数止 - d_天数起)=0 Then
                   n_病人床日数 :=1;
                 Else
                   n_病人床日数 :=(d_天数止 - d_天数起);
                 End If;
                 n_病区床日数 := n_病区床日数 + n_病人床日数 ;
              Else
                 d_天数起 := r_Item(I).开始时间;
                 d_天数止 := Trunc(d_e);
                 If (d_天数止 - d_天数起)=0 Then
                   n_病人床日数 :=1;
                 Else
                   n_病人床日数 :=(d_天数止 - d_天数起);
                 End If;
                 n_病区床日数 := n_病区床日数 + n_病人床日数 ;
              End If;
            End If;
          End If;
        End If;
        v_Pre病人 := r_Item(I).病人id || '_' || r_Item(I).主页id;
      End Loop;
      v_Xtmp := '<ITEM jsonArray="True" ><YF>' || Substr(v_月份, 6, 2) || '</YF><ZYCR>' || n_病区床日数 || '</ZYCR></ITEM>';
      Select Appendchildxml(x_Tmp, '/ITEM/DATALIST', Xmltype(v_Xtmp)) Into x_Tmp From Dual;

      If x_Tmp Is Not Null Then
        Select Appendchildxml(x_Templet, '/OUTPUT/BQLIST', x_Tmp) Into x_Templet From Dual;
      End If;
    End Loop;
  End If;
  Xml_Out := x_Templet;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getzycws;
/

--122937:胡俊勇,2018-03-15,护理接口出参添加的结点属性
CREATE OR REPLACE Procedure Zl_Third_Getzyhzrrs
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --功能：出院者占用总床日数/查询
  --出院者占用总床日数：指当天出院患者的住院床日之总和，即住院总天数。病人入院后于当晚12点前死亡或因故出院的病人, 作为占用床日数1天进行统计
  --基于ZLHIS系统的理解：在指定时间范围出院的病人的住院天数的总和
  --入参：Xml_In
  --<IN>
  --    <BQID></BQID>    //病区ID，传空取所有病区
  --    <KSRQ></KSRQ>  //开始日期   yyyy-mm-dd
  --    <JSRQ></JSRQ>   //结束日期  yyyy-mm-dd
  --</IN>

  --出参：xml_out
  --<OUTPUT>
  --  <BQLIST>
  --    <ITEM>
  --      <BQID></BQID>  //病区ID
  --      <BQMC></BQMC>  //病区名称
  --      <DATALIST>
  --        <ITEM>
  --           <YF></YF>  //月份
  --           <CRS></CRS>  //出院患者占用总床日数
  --        </ITEM>
  --      </DATALIST>
  --    </ITEM>
  --  </BQLIST>
  --</OUTPUT> 

  n_病区id    部门表.Id%Type;
  n_Pre病区id 部门表.Id%Type;
  d_开始      Date;
  d_结束      Date;
  v_Xtmp      Varchar(5000); --临时XML 
  x_Tmp       Xmltype;
  x_Templet   Xmltype;

  Cursor c_Item Is
    Select m.病区id, m.病区名称, m.月, Sum(m.住院天数) As 床日数
    From (Select b.当前病区id As 病区id, b.病人id, b.主页id, a.名称 As 病区名称, To_Char(b.出院日期, 'mm') As 月,
                  (Trunc(b.出院日期) - Trunc(Decode(b.入科时间, Null, b.入院日期, b.入科时间))) As 住院天数
           From 病案主页 B, 部门表 A
           Where b.当前病区id = a.Id And b.当前病区id = n_病区id And b.出院日期 Between d_开始 And d_结束) M
    Group By m.病区id, m.病区名称, m.月
    Having Sum(m.住院天数) > 0;

  Type t_Item Is Table Of c_Item%RowType;
  r_Item t_Item;

  Cursor c_Itemall Is
    Select m.病区id, m.病区名称, m.月, Sum(m.住院天数) As 床日数
    From (Select b.当前病区id As 病区id, b.病人id, b.主页id, a.名称 As 病区名称, To_Char(b.出院日期, 'mm') As 月,
                  (Trunc(b.出院日期) - Trunc(Decode(b.入科时间, Null, b.入院日期, b.入科时间))) As 住院天数
           From 病案主页 B, 部门表 A
           Where b.当前病区id = a.Id And b.出院日期 Between d_开始 And d_结束) M
    Group By m.病区id, m.病区名称, m.月
    Having Sum(m.住院天数) > 0
    Order By m.病区id;
Begin
  Select Extractvalue(Value(A), 'IN/BQID') As 病区id,
         To_Date(Extractvalue(Value(A), 'IN/KSRQ'), 'yyyy-mm-dd hh24:mi:ss') As 开始日期,
         To_Date(Extractvalue(Value(A), 'IN/JSRQ'), 'yyyy-mm-dd hh24:mi:ss') As 结束日期
  Into n_病区id, d_开始, d_结束
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;
  n_Pre病区id := -1;

  If n_病区id Is Null Then
    Open c_Itemall;
    Fetch c_Itemall Bulk Collect
      Into r_Item;
    Close c_Itemall;
  Else
    Open c_Item;
    Fetch c_Item Bulk Collect
      Into r_Item;
    Close c_Item;
  End If;
  x_Templet := Xmltype('<OUTPUT><BQLIST></BQLIST></OUTPUT>');
  For I In 1 .. r_Item.Count Loop
    If n_Pre病区id <> r_Item(I).病区id Then
      If x_Tmp Is Not Null Then
        Select Appendchildxml(x_Templet, '/OUTPUT/BQLIST', x_Tmp) Into x_Templet From Dual;
      End If;
      n_Pre病区id := r_Item(I).病区id;
      v_Xtmp      := '<ITEM jsonArray="True" ><BQID>' || r_Item(I).病区id || '</BQID><BQMC>' || r_Item(I).病区名称 || '</BQMC>';
      v_Xtmp      := v_Xtmp || '<DATALIST></DATALIST></ITEM>';
      x_Tmp       := Xmltype(v_Xtmp);
    End If;
    v_Xtmp := '<ITEM jsonArray="True" ><YF>' || r_Item(I).月 || '</YF><CRS>' || r_Item(I).床日数 || '</CRS></ITEM>';
    Select Appendchildxml(x_Tmp, '/ITEM/DATALIST', Xmltype(v_Xtmp)) Into x_Tmp From Dual;
  End Loop;
  If x_Tmp Is Not Null Then
    Select Appendchildxml(x_Templet, '/OUTPUT/BQLIST', x_Tmp) Into x_Templet From Dual;
  End If;
  Xml_Out := x_Templet;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getzyhzrrs;
/

--122832:陈刘,2018-03-13,移动HIS接口增加节点标记
Create Or Replace Procedure Zl_Third_Tendfile_Gettemphdata
(
  Xmlfilelist_In  Xmltype,
  Xmlfilelist_Out Out Xmltype
) Is
  ---------------------------------------------------------------------------------------------------------------------------
  --功能:获取某病人指定范围内的体温单历史数据
  --入参:Xml_In
  --<IN>
  --<BQ></BQ>       --病区ID
  --<PATIID></PATIID>     --病人ID
  --<PAGEID></PAGEID>     --主页ID
  --<BABY></BABY>      --婴儿
  -- <FW></FW>   --范围：当天、三天、 一周
  --</IN>
  -- 出参:Xml_Out
  --<OUTPUT>
  -- <GROUPS>
  --  <GROUP>
  --   <SJ></SJ>   --发生时间
  --   <CZY></CZY>  --操作员
  --   <ITEMS>
  --    <ITEM>
  --     <XH></XH>   --序号
  --     <MC></MC>   --名称
  --     <NR></NR>   --内容
  --     <WJ />     --未记说明
  --     <BW />     --部位
  --    </ITEM>
  --   </ITEMS>
  --  </GROUP>
  -- </GROUPS>
  --</OUTPUT>
  ----------------------------------------------------------------------------------------------------------------------------
  n_Patiid   Number(18);
  n_Pageid   Number(18);
  n_Baby     Number(18);
  n_Areaid   Number(18);
  n_Fw       Number(18);
  d_开始时间 Date;
  d_结束时间 Date;
  v_Temp     Varchar2(32767);
  x_Templet  Xmltype; --模板XML
Begin
  Select To_Number(Extractvalue(Value(A), 'IN/BQ'))
  Into n_Areaid
  From Table(Xmlsequence(Extract(Xmlfilelist_In, 'IN'))) A;

  Select To_Number(Extractvalue(Value(A), 'IN/PATIID'))
  Into n_Patiid
  From Table(Xmlsequence(Extract(Xmlfilelist_In, 'IN'))) A;

  Select To_Number(Extractvalue(Value(A), 'IN/PAGEID'))
  Into n_Pageid
  From Table(Xmlsequence(Extract(Xmlfilelist_In, 'IN'))) A;

  Select To_Number(Extractvalue(Value(A), 'IN/BABY'))
  Into n_Baby
  From Table(Xmlsequence(Extract(Xmlfilelist_In, 'IN'))) A;

  Select To_Number(Extractvalue(Value(A), 'IN/FW')) Into n_Fw From Table(Xmlsequence(Extract(Xmlfilelist_In, 'IN'))) A;
  d_结束时间 := Sysdate;
  d_开始时间 := d_结束时间 - n_Fw;

  x_Templet := Xmltype('<OUTPUT><GROUPS></GROUPS></OUTPUT>');
  For r_File In (Select a.Id
                 From 病人护理文件 A, 病历文件列表 B
                 Where 病人id = n_Patiid And 主页id = n_Pageid And 婴儿 = n_Baby And a.格式id = b.Id And 保留 = -1 And
                       a.开始时间 > d_开始时间 And (a.结束时间 < d_结束时间 Or a.结束时间 Is Null)
                 Order By a.Id) Loop
    For r_Twd In (Select ID, To_Char(发生时间, 'yyyy-mm-dd hh24:mi:ss') 发生时间, 保存人
                  From 病人护理数据
                  Where 文件id = r_File.Id And 发生时间 Between d_开始时间 And d_结束时间
                  Order By 发生时间 Desc) Loop
      v_Temp := '<GROUP jsonArray="True"><SJ>' || r_Twd.发生时间 || '</SJ><CZY>' || r_Twd.保存人 ||
                '</CZY><ITEMS jsonArray="True"></ITEMS></GROUP>';
      Select Appendchildxml(x_Templet, '/OUTPUT/GROUPS', Xmltype(v_Temp)) Into x_Templet From Dual;
      For r_Nr In (Select 项目序号, 项目名称, 记录内容, 未记说明, 体温部位 From 病人护理明细 Where 记录id = r_Twd.Id) Loop
        v_Temp := '<ITEM jsonArray="True"><XH>' || r_Nr.项目序号 || '</XH><MC>' || r_Nr.项目名称 || '</MC><NR>' || r_Nr.记录内容 ||
                  '</NR><WJ>' || r_Nr.未记说明 || '</WJ><BW>' || r_Nr.体温部位 || '</BW></ITEM>';
        Select Appendchildxml(x_Templet, '/OUTPUT/GROUPS/GROUP/ITEMS', Xmltype(v_Temp)) Into x_Templet From Dual;
      End Loop;
    End Loop;
  End Loop;

  Xmlfilelist_Out := x_Templet;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Tendfile_Gettemphdata;
/

--122832:陈刘,2018-03-13,移动HIS接口增加节点标记
Create Or Replace Procedure Zl_Third_Tendfile_Getmitems
(
  Xmlfilelist_In  Xmltype,
  Xmlfilelist_Out Out Xmltype
) Is
  ---------------------------------------------------------------------------------------------------------------------------
  --功能:获取体温单/护理记录单中的可用活动项目
  --入参:Xml_In
  --无
  -- 出参:Xml_Out
  --<OUTPUT>
  -- <ITEMLIST>
  --  <ITEM>
  --    <XH/>      --序号
  --    <MC/>      --名称
  --    <LX/>      --项目类型
  --    <BS/>      --项目表示
  --    <CD/>      --项目长度
  --    <XS/>      --项目小数
  --    <DW/>     --项目单位
  --    <ZY/>      --项目值域
  --    <SYBR/>      --适用病人
  --    <YYFS/>      --应用方式
  --    <BW/>       --活动项目部位
  --  <ITEM/>
  -- </ITEMLIST>
  --</OUTPUT>
  ----------------------------------------------------------------------------------------------------------------------------
  v_Temp    Varchar2(32767);
  v_Bw      Varchar2(100);
  x_Templet Xmltype; --模板XML
Begin

  x_Templet := Xmltype('<OUTPUT><ITEMLIST></ITEMLIST></OUTPUT>');

  For r_Hdxm In (Select a.项目序号, a.项目名称, a.项目类型, a.项目表示, a.项目长度, a.项目小数, a.项目单位, a.项目值域, a.适用病人, a.应用方式
                 From 护理记录项目 A
                 Where a.项目性质 = 2 And Nvl(a.应用场合, 0) <> 1) Loop
    Select f_List2str(Cast(Collect(部位) As t_Strlist)) Into v_Bw From 体温部位 Where 项目序号 = r_Hdxm.项目序号;
  
    v_Temp := '<ITEM jsonArray="True"><XH>' || r_Hdxm.项目序号 || '</XH><MC>' || r_Hdxm.项目名称 || '</MC><LX>' || r_Hdxm.项目类型 ||
              '</LX><BS>' || r_Hdxm.项目表示 || '</BS><CD>' || r_Hdxm.项目长度 || '</CD><XS>' || r_Hdxm.项目小数 || '</XS><DW>' ||
              r_Hdxm.项目单位 || '</DW><ZY>' || r_Hdxm.项目值域 || '</ZY><SYBR>' || r_Hdxm.适用病人 || '</SYBR><YYFS>' ||
              r_Hdxm.项目表示 || '</YYFS><BW>' || v_Bw || '</BW></ITEM>';
    Select Appendchildxml(x_Templet, '/OUTPUT/ITEMLIST', Xmltype(v_Temp)) Into x_Templet From Dual;
  End Loop;

  Xmlfilelist_Out := x_Templet;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Tendfile_Getmitems;
/

--122832:陈刘,2018-03-13,移动HIS接口增加节点标记
Create Or Replace Procedure Zl_Third_Tendfile_Getdetail
(
  Xmlfilelist_In  Xmltype,
  Xmlfilelist_Out Out Xmltype
) Is
  ---------------------------------------------------------------------------------------------------------------------------
  --功能:获取指定的护理记录单中，某一个护理项目最近若干次的记录内容，按时间由近到远排序
  --入参:Xml_In
  --<IN>
  -- <FILE></FILE>       --文件id
  -- <XH></XH>   --项目序号
  -- <FW></FW>   --范围数值，传3，表示最近3次
  --</IN>
  --出参:Xml_Out
  --<OUTPUT>
  -- <LISTS>
  --  <ITEM>
  --   <TIME></TIME>   --发生时间
  --   <DATA></DATA>   --内容
  --  </ITEM>
  -- </LISTS>
  --</OUTPUT>
  ----------------------------------------------------------------------------------------------------------------------------
  n_Fileid  Number(18);
  n_Xh      Number(18);
  n_Fw      Number(18);
  v_Temp    Varchar2(32767);
  x_Templet Xmltype; --模板XML
Begin
  Select To_Number(Extractvalue(Value(A), 'IN/FILE'))
  Into n_Fileid
  From Table(Xmlsequence(Extract(Xmlfilelist_In, 'IN'))) A;
  Select To_Number(Extractvalue(Value(A), 'IN/XH')) Into n_Xh From Table(Xmlsequence(Extract(Xmlfilelist_In, 'IN'))) A;
  Select To_Number(Extractvalue(Value(A), 'IN/FW')) Into n_Fw From Table(Xmlsequence(Extract(Xmlfilelist_In, 'IN'))) A;

  x_Templet := Xmltype('<OUTPUT><LISTS></LISTS></OUTPUT>');
  For r_Hljl In (Select To_Char(时间, 'YYYY-MM-DD hh24:mi:ss') 时间, 内容
                 From (Select b.发生时间 时间, Decode(c.记录内容, Null, c.未记说明, c.记录内容) 内容,
                               Row_Number() Over(Partition By b.文件id Order By b.发生时间 Desc) As Top
                        From 病人护理数据 B, 病人护理明细 C
                        Where b.Id = c.记录id And 项目序号 = n_Xh And 文件id = n_Fileid And 记录类型 = 1)
                 Where Top <= n_Fw) Loop
    v_Temp := '<ITEM jsonArray="True"><TIME>' || r_Hljl.时间 || '</TIME><DATA>' || r_Hljl.内容 || '</DATA></ITEM>';
    Select Appendchildxml(x_Templet, '/OUTPUT/LISTS', Xmltype(v_Temp)) Into x_Templet From Dual;
  End Loop;
  Xmlfilelist_Out := x_Templet;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Tendfile_Getdetail;
/

--122832:陈刘,2018-03-13,移动HIS接口增加节点标记
Create Or Replace Procedure Zl_Third_Tendfile_Getitems
(
  Xmlfilelist_In  Xmltype,
  Xmlfilelist_Out Out Xmltype
) Is
  ---------------------------------------------------------------------------------------------------------------------------
  --功能:获取可填写的护理项目（含活动项目及空项目）
  --入参:Xml_In
  --<IN>
  -- <BQ></BQ>       --病区ID
  -- <PATIID></PATIID>     --病人ID
  -- <PAGEID></PAGEID>     --主页ID
  -- <BABY></BABY>      --婴儿
  -- <FILE></FILE>   --传空表示获取在用体温单项目，否则获取id对应的护理记录单项目
  --</IN>
  --出参:Xml_Out
  --<OUTPUT />
  -- <YH></YH>      --页号，用于绑定活动项目
  -- <FILE></FILE>   --文件id
  -- <ITEMLIST>
  --  <ITEM>
  --   <LH></LH>     --列号，用于绑定活动项目
  --   <XH></XH>     --项目序号
  --   <MC></MC>     --项目名称
  --   <LX></LX>     --项目类型0数值1-文本
  --   <BS></BS>     --项目表示
  --   <CD></CD>     --项目长度
  --   <XS></XS>     --项目小数
  --   <DW</DW>      --单位
  --   <ZY></ZY>     --值域
  --   <SYBR></SYBR>    --适用病人0所有1病人本人2婴儿
  --   <YYFS></YYFS>    --应用方式0禁止使用1单独使用2与脉搏共用
  --   <BW></BW>   --部位
  --   <XMXZ></XMXZ>  --项目性质1-普通2-活动项目,表示该项目为预留的活动项目位置
  --  </ITEM>
  -- </ITEMLIST>
  --<OUTPUT/>
  ----------------------------------------------------------------------------------------------------------------------------
  n_Patiid  Number(18);
  n_Pageid  Number(18);
  n_Baby    Number(18);
  n_Areaid  Number(18);
  n_Fileid  Number(18);
  n_Format  Number(18);
  n_Yh      Number(18);
  v_Nnit    Varchar2(100);
  v_Hdlh    Varchar2(40);
  v_Temp    Varchar2(32767);
  v_Temp2   Varchar2(32767);
  x_Templet Xmltype; --模板XML
Begin
  Select To_Number(Extractvalue(Value(A), 'IN/BQ'))
  Into n_Areaid
  From Table(Xmlsequence(Extract(Xmlfilelist_In, 'IN'))) A;

  Select To_Number(Extractvalue(Value(A), 'IN/PATIID'))
  Into n_Patiid
  From Table(Xmlsequence(Extract(Xmlfilelist_In, 'IN'))) A;

  Select To_Number(Extractvalue(Value(A), 'IN/PAGEID'))
  Into n_Pageid
  From Table(Xmlsequence(Extract(Xmlfilelist_In, 'IN'))) A;

  Select To_Number(Extractvalue(Value(A), 'IN/BABY'))
  Into n_Baby
  From Table(Xmlsequence(Extract(Xmlfilelist_In, 'IN'))) A;

  Select To_Number(Extractvalue(Value(A), 'IN/FILE'))
  Into n_Fileid
  From Table(Xmlsequence(Extract(Xmlfilelist_In, 'IN'))) A;

  x_Templet := Xmltype('<OUTPUT><ITEMLIST></ITEMLIST></OUTPUT>');
  If Nvl(n_Fileid, 0) = 0 Then
    Select a.格式id
    Into n_Format
    From 病人护理文件 A, 病历文件列表 B
    Where a.格式id = b.Id And b.种类 = 3 And b.保留 = -1 And a.病人id = n_Patiid And a.主页id = n_Pageid And a.婴儿 = n_Baby And
          a.结束时间 Is Null;
    If n_Format = 30 Then
      For r_Twd In (Select b.项目序号, b.项目名称, b.项目类型, b.项目表示, b.项目长度, b.项目小数, b.项目单位, b.项目值域, b.适用病人, b.应用方式, b.项目性质
                    From 体温记录项目 F, 护理记录项目 B
                    Where f.项目序号 = b.项目序号) Loop
        Select f_List2str(Cast(Collect(部位) As t_Strlist)) Into v_Nnit From 体温部位 Where 项目序号 = r_Twd.项目序号;
        v_Temp2 := '<ITEM jsonArray="True"><XH>' || r_Twd.项目序号 || '</XH><MC>' || r_Twd.项目名称 || '</MC><LX>' || r_Twd.项目类型 || '</LX><BS>' ||
                   r_Twd.项目表示 || '</BS><CD>' || r_Twd.项目长度 || '</CD><XS>' || r_Twd.项目小数 || '</XS><DW>' || r_Twd.项目单位 ||
                   '</DW><ZY>' || r_Twd.项目值域 || '</ZY><SYBR>' || r_Twd.适用病人 || '</SYBR><YYFS>' || r_Twd.应用方式 ||
                   '</YYFS><BW>' || v_Nnit || '</BW><XMXZ>' || r_Twd.项目性质 || '</XMXZ></ITEM>';
        Select Appendchildxml(x_Templet, '/OUTPUT/ITEMLIST', Xmltype(v_Temp2)) Into x_Templet From Dual;
      End Loop;
    Else
      For r_Twd In (Select d.对象序号 列号, b.项目序号, b.项目名称, b.项目类型, b.项目表示, b.项目长度, b.项目小数, b.项目单位, b.项目值域, b.适用病人, b.应用方式,
                           b.项目性质
                    From 病历文件结构 C, 病历文件结构 D, 护理记录项目 B
                    Where c.文件id = n_Format And c.父id Is Null And c.对象序号 In (2, 3) And d.父id = c.Id And b.项目名称 = d.要素名称) Loop
        Select f_List2str(Cast(Collect(部位) As t_Strlist)) Into v_Nnit From 体温部位 Where 项目序号 = r_Twd.项目序号;
        v_Temp2 := '<ITEM jsonArray="True"><XH>' || r_Twd.项目序号 || '</XH><MC>' || r_Twd.项目名称 || '</MC><LX>' || r_Twd.项目类型 || '</LX><BS>' ||
                   r_Twd.项目表示 || '</BS><CD>' || r_Twd.项目长度 || '</CD><XS>' || r_Twd.项目小数 || '</XS><DW>' || r_Twd.项目单位 ||
                   '</DW><ZY>' || r_Twd.项目值域 || '</ZY><SYBR>' || r_Twd.适用病人 || '</SYBR><YYFS>' || r_Twd.应用方式 ||
                   '</YYFS><BW>' || v_Nnit || '</BW><XMXZ>' || r_Twd.项目性质 || '</XMXZ></ITEM>';
        Select Appendchildxml(x_Templet, '/OUTPUT/ITEMLIST', Xmltype(v_Temp2)) Into x_Templet From Dual;
      End Loop;
    
    End If;
  Else
    Select Max(结束页号) Into n_Yh From 病人护理打印 Where 文件id = n_Fileid;
    v_Temp := '<YH>' || n_Yh || '</YH>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := ' <FILE>' || n_Fileid || '</FILE>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    For r_Jld In (Select a.页号, a.文件id, a.列号, b.项目序号, b.项目名称, b.项目类型, b.项目表示, b.项目长度, b.项目小数, b.项目单位, b.项目值域, b.适用病人,
                         b.应用方式, b.项目性质
                  From 病人护理活动项目 A, 护理记录项目 B
                  Where b.项目序号 = a.项目序号 And b.项目序号 = a.项目序号 And a.文件id = n_Fileid And a.页号 = n_Yh) Loop
    
      v_Temp2 := '<ITEM jsonArray="True"><LH>' || r_Jld.列号 || '</LH><XH>' || r_Jld.项目序号 || '</XH><MC>' || r_Jld.项目名称 || '</MC><LX>' ||
                 r_Jld.项目类型 || '</LX><BS>' || r_Jld.项目表示 || '</BS><CD>' || r_Jld.项目长度 || '</CD><XS>' || r_Jld.项目小数 ||
                 '</XS><DW>' || r_Jld.项目单位 || '</DW><ZY>' || r_Jld.项目值域 || '</ZY><SYBR>' || r_Jld.适用病人 ||
                 '</SYBR><YYFS>' || r_Jld.应用方式 || '</YYFS><BW>' || v_Nnit || '</BW><XMXZ>' || r_Jld.项目性质 ||
                 '</XMXZ></ITEM>';
      Select Appendchildxml(x_Templet, '/OUTPUT/ITEMLIST', Xmltype(v_Temp2)) Into x_Templet From Dual;
      v_Hdlh := v_Hdlh || ',' || r_Jld.列号;
    End Loop;
  
    --记录单已绑定的项目
    For r_Jldb In (Select d.对象序号 列号, b.项目序号, b.项目名称, b.项目类型, b.项目表示, b.项目长度, b.项目小数, b.项目单位, b.项目值域, b.适用病人, b.应用方式,
                          b.项目性质
                   From 护理记录项目 B, 病历文件结构 C, 病历文件结构 D, 病人护理文件 E
                   Where c.文件id = e.格式id And e.Id = n_Fileid And c.内容文本 = '表列集合' And d.父id = c.Id And b.项目名称(+) = d.要素名称) Loop
      Select f_List2str(Cast(Collect(部位) As t_Strlist)) Into v_Nnit From 体温部位 Where 项目序号 = r_Jldb.项目序号;
      If Not Instr(v_Hdlh || ',', ',' || r_Jldb.列号 || ',') > 0 Then
        v_Temp2 := '<ITEM jsonArray="True"><LH>' || r_Jldb.列号 || '</LH><XH>' || r_Jldb.项目序号 || '</XH><MC>' || r_Jldb.项目名称 || '</MC><LX>' ||
                   r_Jldb.项目类型 || '</LX><BS>' || r_Jldb.项目表示 || '</BS><CD>' || r_Jldb.项目长度 || '</CD><XS>' ||
                   r_Jldb.项目小数 || '</XS><DW>' || r_Jldb.项目单位 || '</DW><ZY>' || r_Jldb.项目值域 || '</ZY><SYBR>' ||
                   r_Jldb.适用病人 || '</SYBR><YYFS>' || r_Jldb.应用方式 || '</YYFS><BW>' || v_Nnit || '</BW><XMXZ>' ||
                   r_Jldb.项目性质 || '</XMXZ></ITEM>';
        Select Appendchildxml(x_Templet, '/OUTPUT/ITEMLIST', Xmltype(v_Temp2)) Into x_Templet From Dual;
      End If;
    End Loop;
  
  End If;
  Xmlfilelist_Out := x_Templet;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Tendfile_Getitems;
/

--122832:陈刘,2018-03-13,移动HIS接口增加节点标记
Create Or Replace Procedure Zl_Third_Tendfile_Getall
(
  Xmlfilelist_In  Xmltype,
  Xmlfilelist_Out Out Xmltype
) Is
  ---------------------------------------------------------------------------------------------------------------------------
  --功能:获取所选病人当前已创建的护理记录单和可创建的护理记录单列表
  --入参:Xml_In
  --<IN>
  -- <BQ></BQ>        --病区ID
  -- <PATIID></PATIID>      --病人ID
  -- <PAGEID></PAGEID>      --主页ID
  -- <BABY></BABY>       --婴儿
  --</IN>
  --出参:Xml_Out
  --<OUTPUT>
  -- <ITEMLIST>
  --  <ITEM>
  --   <ID></ID>   --已创建的为文件ID，未创建的为格式ID
  --   <MC></MC>
  --   <TYPE></TYPE>   --0表示未创建，1表示已创建
  --  </ITEM>
  -- </ITEMLIST>
  --<OUTPUT/>
  ----------------------------------------------------------------------------------------------------------------------------
  n_Patiid  Number(18);
  n_Pageid  Number(18);
  n_Baby    Number(18);
  n_Areaid  Number(18);
  v_Temp    Varchar2(32767);
  x_Templet Xmltype; --模板XML
Begin
  Select To_Number(Extractvalue(Value(A), 'IN/BQ'))
  Into n_Areaid
  From Table(Xmlsequence(Extract(Xmlfilelist_In, 'IN'))) A;

  Select To_Number(Extractvalue(Value(A), 'IN/PATIID'))
  Into n_Patiid
  From Table(Xmlsequence(Extract(Xmlfilelist_In, 'IN'))) A;

  Select To_Number(Extractvalue(Value(A), 'IN/PAGEID'))
  Into n_Pageid
  From Table(Xmlsequence(Extract(Xmlfilelist_In, 'IN'))) A;

  Select To_Number(Extractvalue(Value(A), 'IN/BABY'))
  Into n_Baby
  From Table(Xmlsequence(Extract(Xmlfilelist_In, 'IN'))) A;

  x_Templet := Xmltype('<OUT><ITEMLIST></ITEMLIST></OUT>');

  For r_Ycj In (Select a.Id, a.格式id, a.科室id, c.名称 As 科室, a.文件名称, a.开始时间, a.创建时间, b.保留, b.编号
                From 病人护理文件 A, 病历文件列表 B, 部门表 C
                Where a.格式id = b.Id And a.科室id = c.Id And a.病人id = n_Patiid And a.主页id = n_Pageid And a.婴儿 = n_Baby
                Order By b.保留, a.开始时间) Loop
    v_Temp := '<ITEM jsonArray="True"><ID>' || r_Ycj.Id || '</ID><MC>' || r_Ycj.文件名称 || '</MC><TYPE>1</TYPE></ITEM>';
    Select Appendchildxml(x_Templet, '/OUT/ITEMLIST', Xmltype(v_Temp)) Into x_Templet From Dual;
  End Loop;

  For r_Kcj In (Select ID, 保留, 编号, 格式
                From (Select ID, 保留, 编号, 名称 As 格式
                       From 病历文件列表
                       Where 种类 = 3 And 保留 <> 1 And
                             (通用 = 1 Or (通用 = 2 And ID In (Select 文件id From 病历应用科室 Where 科室id = n_Areaid))))
                Order By 保留, 编号) Loop
    v_Temp := '<ITEM jsonArray="True"><ID>' || r_Kcj.Id || '</ID><MC>' || r_Kcj.格式 || '</MC><TYPE>0</TYPE></ITEM>';
    Select Appendchildxml(x_Templet, '/OUT/ITEMLIST', Xmltype(v_Temp)) Into x_Templet From Dual;
  End Loop;

  Xmlfilelist_Out := x_Templet;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Tendfile_Getall;
/

--122763:冉俊明,2018-03-12,XML循环节点增加 jsonArray 属性
Create Or Replace Procedure Zl_Third_Getexessort
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  -------------------------------------------------------------------------------------------------- 
  --功能:获取住院病人费用分类汇总 
  --入参:Xml_In 
  -- <IN> 
  --  <PATIID></PATIID>         --病人ID 
  --  <PAGEID></PAGEID>     --主页ID 
  --</IN> 
  --出参:Xml_Out 
  --<OUTPUT> 
  --  <ZFY></ZFY>    --总费用 
  --  <FYLIST> 
  --    <ITEM> 
  --      <XM></XM> --收据费目 
  --      <JE></JE> --金额 
  --    </ITEM> 
  --  <FYLIST> 
  --</OUTPUT> 
  -------------------------------------------------------------------------------------------------- 
  n_病人id 病人信息.病人id%Type;
  n_主页id 病案主页.主页id%Type;
  n_总费用 住院费用记录.实收金额%Type;

  x_Templet Xmltype; --模板XML 
Begin
  --获取入参 
  Select Extractvalue(Value(A), 'IN/PATIID'),
         Decode(Extractvalue(Value(A), 'IN/PAGEID'), 0, Null, Extractvalue(Value(A), 'IN/PAGEID'))
  Into n_病人id, n_主页id
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  Select Nvl(Sum(a.实收金额), 0)
  Into n_总费用
  From 住院费用记录 A
  Where a.病人id = n_病人id And a.主页id = n_主页id And Nvl(a.门诊标志, 0) = 2;

  Select Xmlelement("OUTPUT",
                     Xmlforest(n_总费用 As "ZFY",
                                Xmlagg(Xmlelement("ITEM", Xmlattributes('true' As "jsonArray"),
                                                   Xmlforest(a.收据费目 As "XM", Nvl(Sum(a.实收金额), 0) As "JE"))) As "FYLIST"))
  Into x_Templet
  From 住院费用记录 A
  Where a.病人id = n_病人id And a.主页id = n_主页id And Nvl(a.门诊标志, 0) = 2
  Group By a.收据费目;
  Xml_Out := x_Templet;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getexessort;
/

--122714:廖思奇,2018-03-12,Zl_病人医嘱执行_拒绝执行  修正逻辑错误
Create Or Replace Procedure Zl_病人医嘱执行_拒绝执行
(
  医嘱id_In     In 病人医嘱执行.医嘱id%Type,
  发送号_In     In 病人医嘱执行.发送号%Type,
  操作员编号_In In 人员表.编号%Type := Null,
  操作员姓名_In In 人员表.姓名%Type := Null,
  执行部门id_In In 门诊费用记录.执行部门id%Type := 0,
  拒绝原因_In   In 病人医嘱发送.执行说明%Type := Null
  --参数：医嘱ID_IN=单独执行的医嘱ID，检验组合为显示的检验项目的ID。
) Is
  Cursor c_Advice Is
    Select a.Id, a.相关id, a.诊疗类别, a.病人id, a.主页id, a.挂号单, b.No, a.病人来源
    From 病人医嘱记录 A, 病人医嘱发送 B
    Where ID = 医嘱id_In And a.Id = b.医嘱id;
  r_Advice c_Advice%RowType;

  n_Temp     Number;
  v_Temp     Varchar2(255);
  v_人员姓名 人员表.姓名%Type;

Begin
  --当前操作人员
  If 操作员编号_In Is Not Null And 操作员姓名_In Is Not Null Then
    v_人员姓名 := 操作员姓名_In;
  Else
    v_Temp     := Zl_Identity;
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    v_人员姓名 := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  End If;

  Open c_Advice;
  Fetch c_Advice
    Into r_Advice;

  If r_Advice.诊疗类别 = 'C' And r_Advice.相关id Is Not Null Then
    --包含一并采集的所有检验项目
    Update 病人医嘱发送
    Set 执行状态 = 2, 完成人 = v_人员姓名, 完成时间 = Sysdate, 执行说明 = 拒绝原因_In
    Where 发送号 + 0 = 发送号_In And 医嘱id In (Select ID From 病人医嘱记录 Where 相关id = r_Advice.相关id);
  Else
    --包含附加手术,检验部位,以及其它独立医嘱;麻醉和中药煎法是单独安排
    Update 病人医嘱发送
    Set 执行状态 = 2, 完成人 = v_人员姓名, 完成时间 = Sysdate, 执行说明 = 拒绝原因_In
    Where 发送号 + 0 = 发送号_In And 医嘱id In (Select ID
                                        From 病人医嘱记录
                                        Where ID = 医嘱id_In
                                        Union All
                                        Select ID
                                        From 病人医嘱记录
                                        Where 相关id = 医嘱id_In And 诊疗类别 In ('F', 'D'));
  End If;
  If r_Advice.诊疗类别 = 'D' Then
    Select Count(1) Into n_Temp From 部门性质说明 Where 部门id = 执行部门id_In And 工作性质 = '检查';
    If n_Temp > 0 Then
      b_Message.Zlhis_Cis_037(r_Advice.病人id, r_Advice.主页id, r_Advice.挂号单, 发送号_In, 医嘱id_In, r_Advice.No, r_Advice.病人来源);
    End If;
  End If;
  Close c_Advice;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人医嘱执行_拒绝执行;
/

--122731:刘硕,2018-03-09,移动护理接口循环节点属性添加
Create Or Replace Procedure Zl_Third_Getfeeitem
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --------------------------------------------------------------------------------------------------
  --功能:获取收费项目目录,用于在设置护理项目对应收费项目时，输入简码、名称查询HIS收费项目
  --入参:Xml_In:
  --<IN>
  --  <KEY></KEY>   //查询关键字，简码、名称、编码，简码为拼音简码，编码左匹配，名称与简码全匹配。为NULL则不进行匹配查询
  --  <PAGENOW></PAGENOW>  //当前页数，当PAGESIZE与PAGENOW为空或<1,则返回所有数据。否则返回指定页数的数据
  --  <PAGESIZE></PAGESIZE>  //记录条数，当PAGESIZE与PAGENOW为空或<1,则返回所有数据。否则返回指定页数的数据
  --</IN>
  --出参:Xml_Out--以类别与编码排序返回分页
  --<OUTPUT>
  --  <XMLIST>
  --    <XM jsonArray="true">
  --      <LB></LB>   //类别，治疗、护理等
  --      <ID></ID>   //收费项目Id
  --      <BM></BM>   //收费项目编码
  --      <MC></MC>   //收费项目名称
  --      <GG></GG>   //规格
  --      <DW></DW>   //单位
  --      <DJ></DJ>   //单价
  --      <SM></SM>   //说明
  --    </XM>
  --  </XMLIST>
  --</OUTPUT>
  --------------------------------------------------------------------------------------------------
  v_Key      收费项目目录.名称%Type;
  n_Cur_Page Number(5);
  n_Pagesize Number(5);
  x_Templet  Xmltype; --模板XML
Begin
  --获取查询参数
  Select Max(b.Key), Max(b.Pagenow), Max(Pagesize)
  Into v_Key, n_Cur_Page, n_Pagesize
  From Xmltable('$a/IN' Passing Xml_In As "a" Columns Key Varchar2(100) Path 'KEY', Pagenow Number(5) Path 'PAGENOW',
                 Pagesize Number(5) Path 'PAGESIZE') B;

  --获取所有的数据，不匹配
  --查询SQL来源于诊疗项目管理中设置收费项目时输入匹配，变动点为（将Sum修改为Max）
  If v_Key Is Null Then
    --不进行分页，直接返回所有数据
    If Nvl(n_Cur_Page, 0) < 1 Or Nvl(n_Pagesize, 0) < 1 Then
      Select Xmlelement("OUTPUT",
                         Xmlelement("XMLIST",
                                     Xmlagg(Xmlelement("XM", Xmlattributes('true' As "jsonArray"),
                                                        Xmlforest(e.类别名称 As "LB", e.Id As "ID", e.编码 As "BM", e.名称 As "MC",
                                                                   e.规格 As "GG", e.计算单位 As "DW", e.售价 As "DJ", e.说明 As "SM"))))) 部门性质
      Into x_Templet
      From (Select b.名称 类别名称, a.Id, a.编码, a.名称, a.规格, a.产地, a.计算单位, a.说明,
                    Decode(Nvl(a.是否变价, 0), 0, LTrim(RTrim(To_Char(Sum(Nvl(d.现价, 0)), '9999999990.0000'))),
                            Decode(Instr('4567', a.类别), 0, LTrim(RTrim(To_Char(Sum(Nvl(d.缺省价格, 0)), '9999999990.0000'))),
                                    '时价')) As 售价
             From 收费项目目录 A, 收费项目类别 B, 收费价目 D
             Where a.类别 = b.编码 And a.Id = d.收费细目id(+) And a.类别 Not In ('1', 'J') And
                   (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And
                   (a.服务对象 = 1 Or a.服务对象 = 2 Or a.服务对象 = 3) And d.执行日期 <= Sysdate And
                   (d.终止日期 > Sysdate Or d.终止日期 Is Null) And d.价格等级 Is Null
             Group By b.名称, a.Id, a.编码, a.名称, a.规格, a.产地, a.计算单位, a.说明, a.是否变价, a.类别) E;
      --分页查询
    Else
      Select Xmlelement("OUTPUT",
                         Xmlelement("XMLIST",
                                     Xmlagg(Xmlelement("XM", Xmlattributes('true' As "jsonArray"),
                                                        Xmlforest(e.类别名称 As "LB", e.Id As "ID", e.编码 As "BM", e.名称 As "MC",
                                                                   e.规格 As "GG", e.计算单位 As "DW", e.售价 As "DJ", e.说明 As "SM"))))) 部门性质
      Into x_Templet
      From (Select b.名称 类别名称, a.Id, a.编码, a.名称, a.规格, a.产地, a.计算单位, a.说明,
                    Decode(Nvl(a.是否变价, 0), 0, LTrim(RTrim(To_Char(Sum(Nvl(d.现价, 0)), '9999999990.0000'))),
                            Decode(Instr('4567', a.类别), 0, LTrim(RTrim(To_Char(Sum(Nvl(d.缺省价格, 0)), '9999999990.0000'))),
                                    '时价')) As 售价, Row_Number() Over(Order By b.名称, a.编码) As Rn
             From 收费项目目录 A, 收费项目类别 B, 收费价目 D
             Where a.类别 = b.编码 And a.Id = d.收费细目id(+) And a.类别 Not In ('1', 'J') And
                   (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And
                   (a.服务对象 = 1 Or a.服务对象 = 2 Or a.服务对象 = 3) And d.执行日期 <= Sysdate And
                   (d.终止日期 > Sysdate Or d.终止日期 Is Null) And d.价格等级 Is Null
             Group By b.名称, a.Id, a.编码, a.名称, a.规格, a.产地, a.计算单位, a.说明, a.是否变价, a.类别) E
      Where Rn Between n_Pagesize * (n_Cur_Page - 1) + 1 And n_Pagesize * n_Cur_Page;
    End If;
    --获取指定的匹配数据
  Else
    --不进行分页，直接返回所有数据
    If Nvl(n_Cur_Page, 0) < 1 Or Nvl(n_Pagesize, 0) < 1 Then
      Select Xmlelement("OUTPUT",
                         Xmlelement("XMLIST",
                                     Xmlagg(Xmlelement("XM", Xmlattributes('true' As "jsonArray"),
                                                        Xmlforest(e.类别名称 As "LB", e.Id As "ID", e.编码 As "BM", e.名称 As "MC",
                                                                   e.规格 As "GG", e.计算单位 As "DW", e.售价 As "DJ", e.说明 As "SM"))))) 部门性质
      Into x_Templet
      From (Select f.类别名称, f.Id, f.编码, f.名称, f.规格, f.产地, f.计算单位, f.说明,
                    Decode(Nvl(f.是否变价, 0), 0, LTrim(RTrim(To_Char(Sum(Nvl(d.现价, 0)), '9999999990.0000'))),
                            Decode(Instr('4567', f.类别), 0, LTrim(RTrim(To_Char(Sum(Nvl(d.缺省价格, 0)), '9999999990.0000'))),
                                    '时价')) As 售价
             From (Select Distinct b.名称 类别名称, a.Id, a.编码, a.名称, a.规格, a.产地, a.计算单位, a.说明, a.是否变价, a.类别
                    
                    From 收费项目目录 A, 收费项目类别 B, 收费项目别名 C
                    Where a.类别 = b.编码 And a.Id = c.收费细目id And a.类别 Not In ('1', 'J') And
                          (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And
                          (a.服务对象 = 1 Or a.服务对象 = 2 Or a.服务对象 = 3) And
                          (a.编码 Like v_Key || '%' Or c.名称 Like '%' || v_Key || '%' Or c.简码 Like '%' || v_Key || '%') And
                          c.码类 = '1') F,
                  
                  收费价目 D
             Where f.Id = d.收费细目id(+) And d.价格等级 Is Null And d.执行日期 <= Sysdate And (d.终止日期 > Sysdate Or d.终止日期 Is Null)
             Group By f.名称, f.Id, f.编码, f.名称, f.规格, f.产地, f.计算单位, f.说明, f.是否变价, f.类别名称, f.类别) E;
      --分页查询
    Else
      Select Xmlelement("OUTPUT",
                         Xmlelement("XMLIST",
                                     Xmlagg(Xmlelement("XM", Xmlattributes('true' As "jsonArray"),
                                                        Xmlforest(e.类别名称 As "LB", e.Id As "ID", e.编码 As "BM", e.名称 As "MC",
                                                                   e.规格 As "GG", e.计算单位 As "DW", e.售价 As "DJ", e.说明 As "SM"))))) 部门性质
      Into x_Templet
      From (Select f.类别名称, f.Id, f.编码, f.名称, f.规格, f.产地, f.计算单位, f.说明,
                    Decode(Nvl(f.是否变价, 0), 0, LTrim(RTrim(To_Char(Sum(Nvl(d.现价, 0)), '9999999990.0000'))),
                            Decode(Instr('4567', f.类别), 0, LTrim(RTrim(To_Char(Sum(Nvl(d.缺省价格, 0)), '9999999990.0000'))),
                                    '时价')) As 售价, Row_Number() Over(Order By f.名称, f.编码) As Rn
             From (Select Distinct b.名称 类别名称, a.Id, a.编码, a.名称, a.规格, a.产地, a.计算单位, a.说明, a.是否变价, a.类别
                    From 收费项目目录 A, 收费项目类别 B, 收费项目别名 C
                    Where a.类别 = b.编码 And a.Id = c.收费细目id And a.类别 Not In ('1', 'J') And
                          (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And
                          (a.服务对象 = 1 Or a.服务对象 = 2 Or a.服务对象 = 3) And
                          (a.编码 Like v_Key || '%' Or c.名称 Like '%' || v_Key || '%' Or c.简码 Like '%' || v_Key || '%') And
                          c.码类 = '1') F,
                  
                  收费价目 D
             Where f.Id = d.收费细目id(+) And d.价格等级 Is Null And d.执行日期 <= Sysdate And (d.终止日期 > Sysdate Or d.终止日期 Is Null)
             Group By f.名称, f.Id, f.编码, f.名称, f.规格, f.产地, f.计算单位, f.说明, f.是否变价, f.类别名称, f.类别) E
      Where Rn Between n_Pagesize * (n_Cur_Page - 1) + 1 And n_Pagesize * n_Cur_Page;
    End If;
  End If;
  Xml_Out := x_Templet;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getfeeitem;
/

--122731:刘硕,2018-03-09,移动护理接口循环节点属性添加
Create Or Replace Procedure Zl_Third_Getdeptmatch
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --------------------------------------------------------------------------------------------------
  --功能:用于获取科室病区对照关系
  --入参:Xml_In:
  --<IN>
  --  <BMID></BMID>         --科室ID,传空时为获取所有对照关系
  --</IN>
  --出参:Xml_Out
  --<OUTPUT>
  --  <BMKSLIST>
  --    <ITEM jsonArray="true">
  --      <BQID></BQID>  --病区ID
  --      <KSID></KSID>  --科室ID
  --    </ITEM>
  --  <BMKSLIST>
  --</OUTPUT>
  --------------------------------------------------------------------------------------------------
  n_部门id  部门表.Id%Type;
  x_Templet Xmltype; --模板XML
Begin
  --获取部门ID
  Select Max(b.Bmid) Into n_部门id From Xmltable('$a/IN' Passing Xml_In As "a" Columns Bmid Number(18) Path 'BMID') B;

  --获取所有对应关系
  If n_部门id Is Null Then
  
    Select Xmlelement("OUTPUT",
                       Xmlelement("BMKSLIS",
                                   Xmlagg(Xmlelement("ITEM", Xmlattributes('true' As "jsonArray"),
                                                      Xmlforest(a.病区id As "BQID", a.科室id As "KSID"))))) 部门性质
    
    Into x_Templet
    From 病区科室对应 A;
    --获取指定科室对应的病区
  Else
    Select Xmlelement("OUTPUT",
                       Xmlelement("BMKSLIS",
                                   Xmlagg(Xmlelement("ITEM", Xmlattributes('true' As "jsonArray"),
                                                      Xmlforest(a.病区id As "BQID", a.科室id As "KSID"))))) 部门性质
    
    Into x_Templet
    From 病区科室对应 A
    Where a.病区id = n_部门id;
  End If;
  Xml_Out := x_Templet;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getdeptmatch;
/

--122731:刘硕,2018-03-09,移动护理接口循环节点属性添加
Create Or Replace Procedure Zl_Third_Getdept(Xml_Out Out Xmltype) Is
  --------------------------------------------------------------------------------------------------
  --功能:获取部门信息
  --入参:无
  --出参:Xml_Out
  --<OUTPUT>
  --  <BMLIST>
  --    <ITEM jsonArray=”true”>
  --      <ID></ID>  --ID
  --      <MC></MC>  --名称
  --      <BH></BH>  --编号
  --      <JM></JM>  --简码
  --      <ZD></ZD>  --站点
  --      <XZ></XZ>  --性质，多个性质用“,”号分隔
  --    </ITEM>
  --  <BMLIST>
  --</OUTPUT>
  --------------------------------------------------------------------------------------------------
  x_Templet Xmltype; --模板XML
Begin
  --获取所有部门
  Select Xmlelement("OUTPUT",
                     Xmlelement("BMLIST",
                                 Xmlagg(Xmlelement("ITEM", Xmlattributes('true' As "jsonArray"),
                                                    Xmlforest(a.Id As "ID", Max(a.名称) As "MC", Max(a.编码) As "BH",
                                                               Max(a.简码) As "JM", Max(a.站点) As "ZD",
                                                               f_List2str(Cast(Collect(b.工作性质) As t_Strlist)) As "XZ"))))) 部门性质

  
  Into x_Templet
  From 部门表 A, 部门性质说明 B
  Where a.Id = b.部门id
  Group By a.Id;
  Xml_Out := x_Templet;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getdept;
/

--122731:刘硕,2018-03-09,移动护理接口循环节点增加属性
Create Or Replace Procedure Zl_Third_Getperson
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --------------------------------------------------------------------------------------------------
  --功能:获取人员信息
  --入参:Xml_In:
  --<IN>
  --  <BMID></BMID>         --部门ID,传空时为获取所有人员
  --</IN>
  --出参:Xml_Out
  --<OUTPUT>
  --  <RYLIST>
  --    <ITEM jsonArray=”true”>
  --      <ID></ID>  --ID
  --      <XM></XM>  --姓名
  --      <BH></BH>  --编号
  --      <JM></JM>  --简码
  --      <XB></XB>  --性别
  --      <CSRQ></CSRQ> --出生日期
  --      <SFZH></SFZH> --身份证号
  --      <MZ></MZ>  --民族
  --      <XL></XL>  --学历
  --      <ZYJSZW></ZYJSZW> --专业技术职务
  --      <XZ></XZ>  --人员性质，字符串“医生,护士,其他”
  --      <BMLIST>  --所属部门列表
  --        <ITEM jsonArray=”true”>
  --          <ID></ID> --部门ID
  --          <MC></MC> --部门名称
  --        </ITEM>
  --      </BMLIST>
  --    </ITEM>
  --  <RYLIST>
  --</OUTPUT>
  --------------------------------------------------------------------------------------------------
  n_部门id  部门表.Id%Type;
  x_Templet Xmltype; --模板XML
Begin
  --获取部门ID
  Select Max(b.Bmid) Into n_部门id From Xmltable('$a/IN' Passing Xml_In As "a" Columns Bmid Number(18) Path 'BMID') B;
  --获取所有部门人员
  If n_部门id Is Null Then
    Select Xmlelement("OUTPUT",
                       Xmlelement("RYLIST",
                                   Xmlagg(Xmlelement("ITEM", Xmlattributes('true' As "jsonArray"),
                                                      Xmlforest(a.Id As "ID", Max(a.姓名) As "XM", Max(a.编号) As "BH",
                                                                 Max(a.简码) As "JM", Max(a.性别) As "XB",
                                                                 To_Char(Max(a.出生日期), 'YYYY-MM-DD HH24:MI:SS') As "CSRQ",
                                                                 Max(a.身份证号) As "SFZH", Max(a.民族) As "MZ", Max(a.学历) As "XL",
                                                                 Max(a.专业技术职务) As "ZYJSZW", Max(b.人员性质) As "XZ",
                                                                 Xmlagg(Xmlelement("ITEM", Xmlattributes('true' As "jsonArray"),
                                                                                    Xmlforest(g.Id As "ID", g.名称 As "MC"))) As
                                                                  "BMLIST")))))
    Into x_Templet
    From 人员表 A,
         (Select e.人员id, f_List2str(Cast(Collect(e.人员性质) As t_Strlist)) 人员性质
           From (Select c.Id 人员id, d.人员性质
                  From 人员表 C, 人员性质说明 D
                  Where d.人员id = c.Id And d.人员性质 In ('医生', '护士')
                  Union All
                  Select c.Id 人员id, '其他' 人员性质
                  From 人员表 C, 人员性质说明 D
                  Where d.人员id = c.Id And d.人员性质 Not In ('医生', '护士')
                  Group By c.Id) E
           Group By e.人员id) B, 部门人员 F, 部门表 G
    Where a.Id = b.人员id(+) And a.Id = f.人员id(+) And f.部门id = g.Id(+)
    Group By a.Id;
    --获取指定部门人员
  Else
    With People As
     (Select r.Id, r.姓名, r.编号, r.简码, r.性别, r.出生日期, r.身份证号, r.民族, r.学历, r.专业技术职务
      From 人员表 R
      Where r.Id In (Select 人员id From 部门人员 H Where h.部门id = n_部门id))
    Select Xmlelement("OUTPUT",
                       Xmlelement("RYLIST",
                                   Xmlagg(Xmlelement("ITEM", Xmlattributes('true' As "jsonArray"),
                                                      Xmlforest(a.Id As "ID", Max(a.姓名) As "XM", Max(a.编号) As "BH",
                                                                 Max(a.简码) As "JM", Max(a.性别) As "XB",
                                                                 To_Char(Max(a.出生日期), 'YYYY-MM-DD HH24:MI:SS') As "CSRQ",
                                                                 Max(a.身份证号) As "SFZH", Max(a.民族) As "MZ", Max(a.学历) As "XL",
                                                                 Max(a.专业技术职务) As "ZYJSZW", Max(b.人员性质) As "XZ",
                                                                 Xmlagg(Xmlelement("ITEM", Xmlattributes('true' As "jsonArray"),
                                                                                    Xmlforest(g.Id As "ID", g.名称 As "MC"))) As
                                                                  "BMLIST")))))
    Into x_Templet
    From People A,
         (Select e.人员id, f_List2str(Cast(Collect(e.人员性质) As t_Strlist)) 人员性质
           From (Select c.Id 人员id, d.人员性质
                  From People C, 人员性质说明 D
                  Where d.人员id = c.Id And d.人员性质 In ('医生', '护士')
                  Union All
                  Select c.Id 人员id, '其他' 人员性质
                  From People C, 人员性质说明 D
                  Where d.人员id = c.Id And d.人员性质 Not In ('医生', '护士')
                  Group By c.Id) E
           Group By e.人员id) B, 部门人员 F, 部门表 G
    Where a.Id = b.人员id(+) And a.Id = f.人员id(+) And f.部门id = g.Id(+)
    Group By a.Id;
  End If;
  Xml_Out := x_Templet;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getperson;
/







------------------------------------------------------------------------------------
--系统版本号
Update zlSystems Set 版本号='10.35.90.0002' Where 编号=&n_System;
Commit;
