CREATE OR REPLACE Function Zl_排队叫号队列_获取排序方式
(
    业务类型_In 排队叫号队列.业务类型%Type
    --获取不同业务分别所需的排序方式
) Return Varchar2 Is
Begin
  Case 业务类型_In
    When -1 Then NULL;
    Else
      Return 'to_number(排队序号) asc';
  End Case;

Exception
  When Others Then
    Return '';
End Zl_排队叫号队列_获取排序方式;

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values
(100,1160,'基本','ZLHIS','Zl_排队叫号队列_获取排序方式','EXECUTE');