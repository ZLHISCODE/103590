CREATE OR REPLACE Function Zl_�ŶӽкŶ���_��ȡ����ʽ
(
    ҵ������_In �ŶӽкŶ���.ҵ������%Type
    --��ȡ��ͬҵ��ֱ����������ʽ
) Return Varchar2 Is
Begin
  Case ҵ������_In
    When -1 Then NULL;
    Else
      Return 'to_number(�Ŷ����) asc';
  End Case;

Exception
  When Others Then
    Return '';
End Zl_�ŶӽкŶ���_��ȡ����ʽ;

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values
(100,1160,'����','ZLHIS','Zl_�ŶӽкŶ���_��ȡ����ʽ','EXECUTE');