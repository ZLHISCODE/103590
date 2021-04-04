--新增字段“显示状态”
Alter Table 未发药品记录 Add 显示状态 Number(1);

--新增过程
Create Or Replace Procedure Zl_未发药品记录_显示
(
  No_In     药品收发记录.No%Type,
  单据_In   药品收发记录.单据%Type,
  库房id_In 药品收发记录.库房id%Type
) Is
Begin
  --更改单据的显示状态
  Update 未发药品记录 Set 显示状态 = 1 Where NO = No_In And 单据 = 单据_In And 库房id = 库房id_In;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_未发药品记录_显示;
/

--过程修正
Create Or Replace Procedure Zl_未发药品记录_呼叫
(
  No_In       药品收发记录.No%Type,
  单据_In     药品收发记录.单据%Type,
  药房id_In   药品收发记录.库房id%Type,
  发药窗口_In 药品收发记录.发药窗口%Type,
  呼叫内容_In 未发药品记录.呼叫内容%Type := Null,
  呼叫终端_In 未发药品记录.呼叫终端%Type := Null
) Is
Begin
  If 呼叫内容_In Is Null Then
    --呼叫内容为空时，将当前的呼叫状态的单据的呼叫内容清空 
    Update 未发药品记录
    Set 呼叫内容 = Null
    Where 库房id = 药房id_In And 单据 = 单据_In And
          (发药窗口 = 发药窗口_In Or 发药窗口 In (Select 名称 From 发药窗口 Where 叫号窗口 = 发药窗口_In)) And NO = No_In And 排队状态 = 3 And
          填制日期 Between Sysdate - 3 And Sysdate;
  Else
    --呼叫内容不为空时，先将以前的呼叫状态中的单据设置为已呼叫，再将当前单据设置为呼叫状态，并填写呼叫内容和呼叫时间 
    --可以满足同一单据反复呼叫的情况 
    Update 未发药品记录
    Set 排队状态 = 4, 呼叫内容 = Null
    Where 库房id = 药房id_In And (发药窗口 = 发药窗口_In Or 发药窗口 In (Select 名称 From 发药窗口 Where 叫号窗口 = 发药窗口_In)) And 排队状态 = 3;
  
    Update 未发药品记录
    Set 排队状态 = 3, 呼叫内容 = 呼叫内容_In, 呼叫时间 = Sysdate, 呼叫终端 = 呼叫终端_In, 显示状态 = 0
    Where 库房id = 药房id_In And 单据 = 单据_In And NO = No_In And 填制日期 Between Sysdate - 3 And Sysdate;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_未发药品记录_呼叫;
/
