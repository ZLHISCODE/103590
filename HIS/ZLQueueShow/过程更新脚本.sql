CREATE OR REPLACE Procedure Zl_排队叫号队列_清除业务
(
       业务类型_IN 排队叫号队列.业务类型%Type,
       有效天数_IN Number := 1
)
Is
Begin
  case 业务类型_IN
    when -1 then Null;
    else
      --清除当前业务类型，而且时间在有效时间之前的排队信息
      delete from 排队语音呼叫 where 站点=userenv('TERMINAL') and nvl(业务类型,0) = 业务类型_IN And 生成时间 <=  sysdate - (1 / 48);
     
      Delete From 排队叫号队列 
      Where 业务类型 = 业务类型_IN And To_Number(Trunc(Sysdate - 排队叫号队列.排队时间)) >= 有效天数_In;
  end case;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_排队叫号队列_清除业务;
/


CREATE OR REPLACE Function Zl_排队叫号队列_获取排序方式
( 
    业务类型_In 排队叫号队列.业务类型%Type 
    --获取不同业务分别所需的排序方式 
) Return Varchar2 Is 
Begin 
  Case 业务类型_In 
    When -1 Then NULL; 
    when 0 then  --临床业务
      return '优先 desc , 排队序号, 排队时间';
    Else 
      Return 'to_number(排队序号) asc'; 
  End Case; 
 
Exception 
  When Others Then 
    Return ''; 
End Zl_排队叫号队列_获取排序方式;
/
