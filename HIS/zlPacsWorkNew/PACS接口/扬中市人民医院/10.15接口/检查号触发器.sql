---��������Ҫ��ģ�����д�������š���ʱ��ͬʱ�޸ġ����š�
create or replace trigger zl_����ҽ������_����ID
  after update on ����ҽ������  
  for each row
Declare
  v_���� Number(18);
Begin
  If :New.����id Is Null Then
    Return;
  End If;

  Begin
    Select To_Number(����)
    Into v_����
    From ���˲������� A, ���˲����ı��� B
    Where A.ID = B.����id And A.������¼id = :New.����id And A.�����ı� = '������' And Rownum = 1;
  Exception
    When Others Then
      v_���� := Null;
  End;
  If v_���� Is Null Then
    Return;
  End If;

  Update Ӱ�����¼ Set ���� = v_���� Where ҽ��id = :Old.ҽ��id;
End Zl_����ҽ������_����id;
/
