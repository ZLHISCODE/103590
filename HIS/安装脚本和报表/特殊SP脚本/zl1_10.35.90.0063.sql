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
--139596:焦博,2019-05-30,调整Oracle过程Zl_票据领用记录_Update,在内部检查入库ID
Create Or Replace Procedure Zl_票据领用记录_Update
(
  Id_In       In 票据领用记录.Id%Type,
  使用类别_In In 票据领用记录.使用类别%Type,
  领用人_In   In 票据领用记录.领用人%Type,
  开始号码_In In 票据领用记录.开始号码%Type,
  终止号码_In In 票据领用记录.终止号码%Type,
  前缀文本_In In 票据领用记录.前缀文本%Type := Null,
  使用方式_In In 票据领用记录.使用方式%Type := 1,
  登记时间_In In 票据领用记录.登记时间%Type := Null,
  登记人_In   In 票据领用记录.登记人%Type := Null,
  批次_In     In 票据领用记录.批次%Type := Null,
  签字人_In   In 票据领用记录.签字人%Type := Null,
  入库id_In   In 票据领用记录.入库id%Type := Null
  
) Is
  Cursor c_领用记录 Is
    Select * From 票据领用记录 Where ID = Id_In For Update;

  c_记录     票据领用记录%RowType;
  n_使用数量 票据领用记录.剩余数量%Type;
  n_剩余数量 票据领用记录.剩余数量%Type;
  n_原领用数 票据领用记录.剩余数量%Type;
  n_现领用数 票据领用记录.剩余数量%Type;

  v_开始号码 票据领用记录.开始号码%Type;
  v_终止号码 票据领用记录.终止号码%Type;
  v_Err_Msg  Varchar2(500);
  Err_Item Exception;
  n_Count  Number(18);
  n_剩余数 票据入库记录.剩余数量%Type;
  n_入库id 票据领用记录.入库id%Type;
Begin
  Open c_领用记录;
  Fetch c_领用记录
    Into c_记录;

  If c_领用记录%NotFound Then
    --记录未找到 
    v_Err_Msg := '[ZLSOFT]该条记录已经被删除，不能修改。[ZLSOFT]';
    Raise Err_Item;
  End If;

  n_入库id := 入库id_In;
  If Nvl(n_入库id, 0) = 0 Then
    n_入库id := c_记录.入库id;
  End If;
  Select Min(号码), Max(号码) Into v_开始号码, v_终止号码 From 票据使用明细 Where 领用id = Id_In;

  If 前缀文本_In Is Null Then
    n_剩余数量 := To_Number(终止号码_In) - To_Number(开始号码_In) + 1;
  Else
    n_剩余数量 := To_Number(Substr(终止号码_In, Length(前缀文本_In) + 1)) - To_Number(Substr(开始号码_In, Length(前缀文本_In) + 1)) + 1;
  End If;

  n_现领用数 := n_剩余数量;
  If c_记录.前缀文本 Is Null Then
    n_原领用数 := To_Number(c_记录.终止号码) - To_Number(c_记录.开始号码) + 1;
  Else
    n_原领用数 := To_Number(Substr(c_记录.终止号码, Length(c_记录.前缀文本) + 1)) - To_Number(Substr(c_记录.开始号码, Length(c_记录.前缀文本) + 1)) + 1;
  End If;

  If v_开始号码 Is Not Null Then
    --已经使用，对一些项目进行验证 
    If Nvl(前缀文本_In, ' ') <> Nvl(c_记录.前缀文本, ' ') Then
      v_Err_Msg := '[ZLSOFT]该条记录领用的票据已经使用，不能修改号码的前缀。[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    If Length(开始号码_In) <> Length(c_记录.开始号码) Then
      v_Err_Msg := '[ZLSOFT]该条记录领用的票据已经使用，不能修改号码的长度。[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    If 开始号码_In > v_开始号码 Then
      v_Err_Msg := '[ZLSOFT]该条记录领用的票据已经使用，开始号码最大只能是' || v_开始号码 || '。[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    If 终止号码_In < v_终止号码 Then
      v_Err_Msg := '[ZLSOFT]该条记录领用的票据已经使用，终止号码最小只能是' || v_终止号码 || '。[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    --下面计算数量 
    If 前缀文本_In Is Null Then
      n_使用数量 := To_Number(c_记录.终止号码) - To_Number(c_记录.开始号码) + 1 - c_记录.剩余数量;
    Else
      n_使用数量 := To_Number(Substr(c_记录.终止号码, Length(前缀文本_In) + 1)) - To_Number(Substr(c_记录.开始号码, Length(前缀文本_In) + 1)) + 1 -
                c_记录.剩余数量;
    End If;
  
    n_剩余数量 := n_剩余数量 - n_使用数量;
  End If;

  For v_入库 In (Select ID, 前缀文本, 使用类别, 开始号码, Nvl(终止号码, 开始号码) As 终止号码
               From 票据入库记录
               Where ID = Nvl(n_入库id, 0) And 票种 = c_记录.票种) Loop
  
    If Nvl(使用类别_In, 'LXH') <> Nvl(v_入库.使用类别, 'LXH') Then
      v_Err_Msg := '[ZLSOFT]当前领用的使用类别『' || Nvl(使用类别_In, '') || '』与入库的使用类别不一致『' || Nvl(v_入库.使用类别, '') || '』![ZLSOFT]';
      Raise Err_Item;
    End If;
  
    --1. 入库检查 
    If 开始号码_In < v_入库.开始号码 Or 开始号码_In > v_入库.终止号码 Then
      --不在入库范围,不能领用 
      v_Err_Msg := '[ZLSOFT]当前领用的开始号码『' || 开始号码_In || '』不在入库范围' || Chr(10) || Chr(13) || '『' || v_入库.开始号码 || '-' ||
                   v_入库.终止号码 || '』不能领用该票据![ZLSOFT]';
      Raise Err_Item;
    End If;
    If 终止号码_In < v_入库.开始号码 Or 终止号码_In > v_入库.终止号码 Then
      --不在入库范围,不能领用 
      v_Err_Msg := '[ZLSOFT]当前领用的终止号码『' || 终止号码_In || '』不在入库范围' || Chr(10) || Chr(13) || '『' || v_入库.开始号码 || '-' ||
                   v_入库.终止号码 || '』不能领用该票据![ZLSOFT]';
      Raise Err_Item;
    End If;
    --2.检查票据是否已经被报损,不能重复报损 
    Select Count(*)
    Into n_Count
    From 票据报损记录
    Where 入库id = Nvl(n_入库id, 0) And ((开始号码_In Between 开始号码 And 终止号码) Or (终止号码_In Between 开始号码 And 终止号码) Or
          (开始号码 Between 开始号码_In And 终止号码_In) Or (终止号码 Between 开始号码_In And 终止号码_In));
  
    If n_Count <> 0 Then
      If 开始号码_In = 终止号码_In Then
        v_Err_Msg := '[ZLSOFT]号码:' || 开始号码_In || '在报损记录中已经存在,不能再进行领用![ZLSOFT]';
      Else
        v_Err_Msg := '[ZLSOFT]号码:' || 开始号码_In || '-' || 终止号码_In || '在报损记录中已经存在,不能再进行领用![ZLSOFT]';
      End If;
      Raise Err_Item;
    End If;
  
    --3.检查票据是否已经被领用,领用的不能再进行报损 
    Select Count(*)
    Into n_Count
    From 票据领用记录
    Where 批次 = Nvl(批次_In, 0) And ID <> Id_In And
          ((开始号码_In Between 开始号码 And 终止号码) Or (终止号码_In Between 开始号码 And 终止号码) Or (开始号码 Between 开始号码_In And 终止号码_In) Or
          (终止号码 Between 开始号码_In And 终止号码_In));
    If n_Count <> 0 Then
      If 开始号码_In = 终止号码_In Then
        v_Err_Msg := '[ZLSOFT]号码:' || 开始号码_In || '在票据领用记录中已经存在,不能再进行领用![ZLSOFT]';
      Else
        v_Err_Msg := '[ZLSOFT]号码:' || 开始号码_In || '-' || 终止号码_In || '在票据领用记录中已经存在,不能再进行领用![ZLSOFT]';
      End If;
      Raise Err_Item;
    End If;
  
    --减少库存 
    Update 票据入库记录
    Set 剩余数量 = Nvl(剩余数量, 0) + (Nvl(n_原领用数, 0) - Nvl(n_现领用数, 0))
    Where ID = Nvl(n_入库id, 0) And 票种 = c_记录.票种
    Returning Nvl(剩余数量, 0) Into n_剩余数;
    If n_剩余数 < 0 Then
      v_Err_Msg := '[ZLSOFT]入库票据的剩余票据数不足,请检查![ZLSOFT]';
      Raise Err_Item;
    End If;
    Update 票据入库记录
    Set 有无票据 = Decode(Sign(Nvl(n_剩余数, 0)), 1, 1, Null)
    Where ID = (Select 入库id From 票据领用记录 Where ID = Id_In) And 票种 = c_记录.票种;
  End Loop;

  Update 票据领用记录
  Set 领用人 = 领用人_In, 前缀文本 = 前缀文本_In, 开始号码 = 开始号码_In, 终止号码 = 终止号码_In, 使用方式 = 使用方式_In, 登记时间 = 登记时间_In, 登记人 = 登记人_In,
      剩余数量 = n_剩余数量, 批次 = 批次_In, 使用类别 = 使用类别_In, 签字人 = 签字人_In, 签字时间 = Decode(签字人_In, Null, Null + Sysdate, Sysdate)
  Where ID = Id_In;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_票据领用记录_Update;
/





------------------------------------------------------------------------------------
--系统版本号
Update zlSystems Set 版本号='10.35.90.0063' Where 编号=&n_System;
Commit;
