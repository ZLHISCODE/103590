--�����ֶΡ���ʾ״̬��
Alter Table δ��ҩƷ��¼ Add ��ʾ״̬ Number(1);

--��������
Create Or Replace Procedure Zl_δ��ҩƷ��¼_��ʾ
(
  No_In     ҩƷ�շ���¼.No%Type,
  ����_In   ҩƷ�շ���¼.����%Type,
  �ⷿid_In ҩƷ�շ���¼.�ⷿid%Type
) Is
Begin
  --���ĵ��ݵ���ʾ״̬
  Update δ��ҩƷ��¼ Set ��ʾ״̬ = 1 Where NO = No_In And ���� = ����_In And �ⷿid = �ⷿid_In;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_δ��ҩƷ��¼_��ʾ;
/

--��������
Create Or Replace Procedure Zl_δ��ҩƷ��¼_����
(
  No_In       ҩƷ�շ���¼.No%Type,
  ����_In     ҩƷ�շ���¼.����%Type,
  ҩ��id_In   ҩƷ�շ���¼.�ⷿid%Type,
  ��ҩ����_In ҩƷ�շ���¼.��ҩ����%Type,
  ��������_In δ��ҩƷ��¼.��������%Type := Null,
  �����ն�_In δ��ҩƷ��¼.�����ն�%Type := Null
) Is
Begin
  If ��������_In Is Null Then
    --��������Ϊ��ʱ������ǰ�ĺ���״̬�ĵ��ݵĺ���������� 
    Update δ��ҩƷ��¼
    Set �������� = Null
    Where �ⷿid = ҩ��id_In And ���� = ����_In And
          (��ҩ���� = ��ҩ����_In Or ��ҩ���� In (Select ���� From ��ҩ���� Where �кŴ��� = ��ҩ����_In)) And NO = No_In And �Ŷ�״̬ = 3 And
          �������� Between Sysdate - 3 And Sysdate;
  Else
    --�������ݲ�Ϊ��ʱ���Ƚ���ǰ�ĺ���״̬�еĵ�������Ϊ�Ѻ��У��ٽ���ǰ��������Ϊ����״̬������д�������ݺͺ���ʱ�� 
    --��������ͬһ���ݷ������е���� 
    Update δ��ҩƷ��¼
    Set �Ŷ�״̬ = 4, �������� = Null
    Where �ⷿid = ҩ��id_In And (��ҩ���� = ��ҩ����_In Or ��ҩ���� In (Select ���� From ��ҩ���� Where �кŴ��� = ��ҩ����_In)) And �Ŷ�״̬ = 3;
  
    Update δ��ҩƷ��¼
    Set �Ŷ�״̬ = 3, �������� = ��������_In, ����ʱ�� = Sysdate, �����ն� = �����ն�_In, ��ʾ״̬ = 0
    Where �ⷿid = ҩ��id_In And ���� = ����_In And NO = No_In And �������� Between Sysdate - 3 And Sysdate;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_δ��ҩƷ��¼_����;
/
