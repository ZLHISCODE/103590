CREATE OR REPLACE Procedure Zl_�ŶӽкŶ���_���ҵ��
(
       ҵ������_IN �ŶӽкŶ���.ҵ������%Type,
       ��Ч����_IN Number := 1
)
Is
Begin
  case ҵ������_IN
    when -1 then Null;
    else
      --�����ǰҵ�����ͣ�����ʱ������Чʱ��֮ǰ���Ŷ���Ϣ
      delete from �Ŷ��������� where վ��=userenv('TERMINAL') and nvl(ҵ������,0) = ҵ������_IN And ����ʱ�� <=  sysdate - (1 / 48);
     
      Delete From �ŶӽкŶ��� 
      Where ҵ������ = ҵ������_IN And To_Number(Trunc(Sysdate - �ŶӽкŶ���.�Ŷ�ʱ��)) >= ��Ч����_In;
  end case;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ŶӽкŶ���_���ҵ��;
/


CREATE OR REPLACE Function Zl_�ŶӽкŶ���_��ȡ����ʽ
( 
    ҵ������_In �ŶӽкŶ���.ҵ������%Type 
    --��ȡ��ͬҵ��ֱ����������ʽ 
) Return Varchar2 Is 
Begin 
  Case ҵ������_In 
    When -1 Then NULL; 
    when 0 then  --�ٴ�ҵ��
      return '���� desc , �Ŷ����, �Ŷ�ʱ��';
    Else 
      Return 'to_number(�Ŷ����) asc'; 
  End Case; 
 
Exception 
  When Others Then 
    Return ''; 
End Zl_�ŶӽкŶ���_��ȡ����ʽ;
/
